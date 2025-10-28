VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmTxn_ReservationInvoice 
   Caption         =   "Reservation Invoice"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame5 
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
      Left            =   4200
      TabIndex        =   102
      Top             =   720
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
         Height          =   2400
         Left            =   1080
         MaxLength       =   2000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   105
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
         TabIndex        =   104
         Top             =   2880
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
         TabIndex        =   103
         Top             =   2880
         Width           =   1455
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
         Left            =   135
         TabIndex        =   106
         Top             =   1410
         Width           =   1260
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H008080FF&
      Caption         =   "Frame Work Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   5280
      TabIndex        =   97
      Top             =   4560
      Visible         =   0   'False
      Width           =   5055
      Begin VB.ComboBox CmbFramework 
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
         ItemData        =   "FrmTxn_ReservationInvoice.frx":0000
         Left            =   2280
         List            =   "FrmTxn_ReservationInvoice.frx":000A
         Sorted          =   -1  'True
         TabIndex        =   101
         Top             =   270
         Width           =   2535
      End
      Begin VB.CommandButton CmdCancelFrameWork 
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
         Left            =   3615
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   99
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton CmdPrintFrameWork 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Print"
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
         Left            =   2415
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   98
         Top             =   840
         Width           =   1200
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         Caption         =   "Management Type"
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
         TabIndex        =   100
         Top             =   330
         Width           =   1950
      End
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      TabIndex        =   16
      Top             =   10080
      Width           =   15015
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   9975
      Left            =   15000
      TabIndex        =   6
      Top             =   -120
      Width           =   255
   End
   Begin VB.Frame Frame4 
      Height          =   9495
      Left            =   120
      TabIndex        =   17
      Top             =   0
      Width           =   14895
      Begin VB.Frame Frame0 
         Height          =   2415
         Left            =   120
         TabIndex        =   43
         Top             =   4560
         Visible         =   0   'False
         Width           =   14850
         Begin VB.TextBox TxtSInvoiceNo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   3765
            TabIndex        =   107
            TabStop         =   0   'False
            Top             =   720
            Width           =   1575
         End
         Begin VB.CommandButton CmdGo 
            Caption         =   "Go"
            Height          =   350
            Left            =   12450
            TabIndex        =   50
            Top             =   750
            Width           =   1000
         End
         Begin VB.CommandButton CmdExcel 
            Caption         =   "Excel"
            Height          =   350
            Left            =   13515
            TabIndex        =   49
            Top             =   750
            Width           =   1000
         End
         Begin VB.CheckBox ChkBatchNo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Check1"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   7440
            TabIndex        =   48
            Top             =   780
            Width           =   225
         End
         Begin VB.ComboBox CmbSStatus 
            Height          =   315
            ItemData        =   "FrmTxn_ReservationInvoice.frx":0036
            Left            =   750
            List            =   "FrmTxn_ReservationInvoice.frx":0043
            Sorted          =   -1  'True
            TabIndex        =   47
            Top             =   720
            Width           =   1875
         End
         Begin VB.ComboBox CmbSContract 
            Height          =   315
            Left            =   6645
            Sorted          =   -1  'True
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   360
            Width           =   4470
         End
         Begin VB.ComboBox CmbSLocation 
            Height          =   315
            Left            =   3585
            Sorted          =   -1  'True
            TabIndex        =   45
            Top             =   360
            Width           =   3045
         End
         Begin VB.ComboBox CmbNonNcdexClient 
            Height          =   315
            Left            =   11130
            Sorted          =   -1  'True
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   360
            Width           =   3390
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   1095
            Left            =   120
            TabIndex        =   51
            Top             =   1200
            Width           =   14595
            _ExtentX        =   25744
            _ExtentY        =   1931
            _Version        =   393216
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSComCtl2.DTPicker TxtFromDate 
            Height          =   300
            Left            =   120
            TabIndex        =   109
            Top             =   375
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
            Format          =   121896961
            CurrentDate     =   39884
         End
         Begin MSComCtl2.DTPicker TxtToDate 
            Height          =   300
            Left            =   1815
            TabIndex        =   110
            Top             =   375
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
            Format          =   121896961
            CurrentDate     =   39884
         End
         Begin VB.Label Label35 
            Caption         =   "To Date"
            Height          =   225
            Left            =   2227
            TabIndex        =   112
            Top             =   120
            Width           =   630
         End
         Begin VB.Label Label34 
            Caption         =   "From Date"
            Height          =   225
            Left            =   472
            TabIndex        =   111
            Top             =   120
            Width           =   750
         End
         Begin VB.Label Label32 
            Caption         =   "Invoice No"
            Height          =   255
            Left            =   2880
            TabIndex        =   108
            Top             =   780
            Width           =   975
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
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
            Left            =   10395
            TabIndex        =   58
            Top             =   795
            Width           =   75
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
            Left            =   7920
            TabIndex        =   57
            Top             =   780
            Width           =   2415
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
            Left            =   5640
            TabIndex        =   56
            Top             =   780
            Width           =   1680
         End
         Begin VB.Label Label10 
            Caption         =   "Status"
            Height          =   225
            Left            =   135
            TabIndex        =   55
            Top             =   765
            Width           =   465
         End
         Begin VB.Label Label11 
            Caption         =   "Contract"
            Height          =   225
            Left            =   8543
            TabIndex        =   54
            Top             =   120
            Width           =   675
         End
         Begin VB.Label Label36 
            Caption         =   "Location"
            Height          =   300
            Left            =   4755
            TabIndex        =   53
            Top             =   120
            Width           =   705
         End
         Begin VB.Label Label37 
            Caption         =   "Client"
            Height          =   225
            Left            =   12615
            TabIndex        =   52
            Top             =   150
            Width           =   435
         End
      End
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
         Left            =   4680
         TabIndex        =   65
         Top             =   3840
         Visible         =   0   'False
         Width           =   6375
         Begin VB.TextBox TxtReverseRemark 
            Appearance      =   0  'Flat
            Height          =   480
            Left            =   1560
            MaxLength       =   250
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   68
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
            TabIndex        =   67
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
            TabIndex        =   66
            Top             =   1440
            Width           =   1455
         End
         Begin MSComCtl2.DTPicker TxtReverseDate 
            Height          =   375
            Left            =   1560
            TabIndex        =   69
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
            Format          =   121896961
            CurrentDate     =   39884
         End
         Begin VB.Label Label5 
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
            TabIndex        =   71
            Top             =   930
            Width           =   1260
         End
         Begin VB.Label Label6 
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
            TabIndex        =   70
            Top             =   390
            Width           =   1800
         End
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   3855
         Left            =   120
         TabIndex        =   74
         Top             =   4320
         Width           =   14595
         _ExtentX        =   25744
         _ExtentY        =   6800
         _Version        =   393216
         TabHeight       =   520
         TabCaption(0)   =   "Billing Details"
         TabPicture(0)   =   "FrmTxn_ReservationInvoice.frx":0062
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "MSHFlexGrid2"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "TI details / Log History"
         TabPicture(1)   =   "FrmTxn_ReservationInvoice.frx":007E
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label28"
         Tab(1).Control(1)=   "Label26"
         Tab(1).Control(2)=   "Label25"
         Tab(1).Control(3)=   "Label20"
         Tab(1).Control(4)=   "Label15"
         Tab(1).Control(5)=   "Label17"
         Tab(1).Control(6)=   "Label16"
         Tab(1).Control(7)=   "Label38"
         Tab(1).Control(8)=   "Frame2"
         Tab(1).Control(9)=   "TxtBalAmt"
         Tab(1).Control(10)=   "TxtEntryNo"
         Tab(1).Control(11)=   "TxtRecAmt"
         Tab(1).Control(12)=   "TxtTISTAmt"
         Tab(1).Control(13)=   "TxtTINo"
         Tab(1).Control(14)=   "TxtTIDate"
         Tab(1).Control(15)=   "TxtTIAmount"
         Tab(1).Control(16)=   "TxtPostInvoiceNo"
         Tab(1).ControlCount=   17
         TabCaption(2)   =   "Stock Statement"
         TabPicture(2)   =   "FrmTxn_ReservationInvoice.frx":009A
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "CmdStockExcel"
         Tab(2).Control(1)=   "MSHFlexGrid4"
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
            Left            =   -73080
            Locked          =   -1  'True
            TabIndex        =   113
            Top             =   1065
            Width           =   5295
         End
         Begin VB.CommandButton CmdStockExcel 
            Caption         =   "Excel"
            Height          =   495
            Left            =   -61800
            TabIndex        =   93
            Top             =   480
            Width           =   1215
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
            Left            =   -71250
            Locked          =   -1  'True
            TabIndex        =   83
            Top             =   645
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
            Left            =   -73065
            Locked          =   -1  'True
            TabIndex        =   82
            Top             =   645
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
            Left            =   -74880
            Locked          =   -1  'True
            TabIndex        =   81
            Top             =   645
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
            Left            =   -69435
            Locked          =   -1  'True
            TabIndex        =   80
            Top             =   645
            Width           =   1695
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
            Left            =   -67740
            Locked          =   -1  'True
            TabIndex        =   79
            Top             =   645
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
            Left            =   -64110
            Locked          =   -1  'True
            TabIndex        =   78
            Top             =   645
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
            Left            =   -65925
            Locked          =   -1  'True
            TabIndex        =   77
            Top             =   645
            Width           =   1815
         End
         Begin VB.Frame Frame2 
            Height          =   2055
            Left            =   -74880
            TabIndex        =   75
            Top             =   1560
            Width           =   14415
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
               Height          =   1695
               Left            =   120
               TabIndex        =   76
               Top             =   240
               Width           =   14175
               _ExtentX        =   25003
               _ExtentY        =   2990
               _Version        =   393216
               AllowUserResizing=   3
               _NumberOfBands  =   1
               _Band(0).Cols   =   2
            End
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
            Height          =   3255
            Left            =   120
            TabIndex        =   91
            Top             =   480
            Width           =   14295
            _ExtentX        =   25215
            _ExtentY        =   5741
            _Version        =   393216
            AllowUserResizing=   3
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid4 
            Height          =   3255
            Left            =   -74880
            TabIndex        =   92
            Top             =   480
            Width           =   12975
            _ExtentX        =   22886
            _ExtentY        =   5741
            _Version        =   393216
            AllowUserResizing=   3
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
            Height          =   255
            Left            =   -74880
            TabIndex        =   114
            Top             =   1125
            Width           =   1455
         End
         Begin VB.Label Label16 
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
            Left            =   -72495
            TabIndex        =   90
            Top             =   375
            Width           =   735
         End
         Begin VB.Label Label17 
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
            Left            =   -70935
            TabIndex        =   89
            Top             =   375
            Width           =   975
         End
         Begin VB.Label Label15 
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
            Left            =   -74205
            TabIndex        =   88
            Top             =   375
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
            Left            =   -69135
            TabIndex        =   87
            Top             =   375
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
            Left            =   -67635
            TabIndex        =   86
            Top             =   375
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
            Left            =   -63615
            TabIndex        =   85
            Top             =   375
            Width           =   855
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
            Left            =   -65745
            TabIndex        =   84
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.Frame Frame1 
         Height          =   4125
         Left            =   120
         TabIndex        =   20
         Top             =   120
         Width           =   14655
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
            Left            =   12465
            Locked          =   -1  'True
            TabIndex        =   115
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
            Left            =   9120
            TabIndex        =   2
            Top             =   1740
            Width           =   1905
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
            Left            =   5160
            TabIndex        =   1
            Top             =   1740
            Width           =   2025
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
            TabIndex        =   64
            TabStop         =   0   'False
            Top             =   3555
            Width           =   5250
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
            Height          =   765
            Left            =   1920
            MaxLength       =   2000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   5
            Top             =   2685
            Width           =   12570
         End
         Begin VB.TextBox TxtActualAmount 
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
            Left            =   12855
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   7
            Top             =   1725
            Width           =   1695
         End
         Begin VB.TextBox TxtContractName 
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
            TabIndex        =   62
            Top             =   730
            Width           =   5295
         End
         Begin VB.TextBox TxtClient 
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
            Left            =   1920
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   61
            Top             =   1220
            Width           =   5295
         End
         Begin VB.TextBox TxtInvoiceNo 
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
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   60
            TabStop         =   0   'False
            Top             =   240
            Width           =   1695
         End
         Begin VB.ComboBox CmbServiceTax 
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
            ItemData        =   "FrmTxn_ReservationInvoice.frx":00B6
            Left            =   1920
            List            =   "FrmTxn_ReservationInvoice.frx":00B8
            Sorted          =   -1  'True
            TabIndex        =   3
            Top             =   2205
            Width           =   5295
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
            Height          =   375
            Left            =   5400
            Locked          =   -1  'True
            TabIndex        =   27
            Top             =   240
            Width           =   1815
         End
         Begin VB.TextBox TxtBillingAmount 
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
            Left            =   1920
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   26
            Top             =   1710
            Width           =   1815
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
            IMEMode         =   3  'DISABLE
            Left            =   9120
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   25
            Top             =   1216
            Width           =   5400
         End
         Begin VB.TextBox TxtLocation 
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
            Left            =   9120
            Locked          =   -1  'True
            TabIndex        =   24
            Top             =   728
            Width           =   5400
         End
         Begin VB.TextBox TxtInvoiceMonth 
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
            Left            =   9120
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   23
            Top             =   240
            Width           =   1815
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
            Height          =   375
            IMEMode         =   3  'DISABLE
            Left            =   12630
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   22
            Top             =   3540
            Width           =   1815
         End
         Begin VB.TextBox TxtNAVCode 
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
            Left            =   9240
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   21
            Top             =   3525
            Width           =   1815
         End
         Begin VB.TextBox TxtServiceTax 
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
            Locked          =   -1  'True
            TabIndex        =   63
            TabStop         =   0   'False
            Top             =   2205
            Width           =   5280
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
            ItemData        =   "FrmTxn_ReservationInvoice.frx":00BA
            Left            =   9120
            List            =   "FrmTxn_ReservationInvoice.frx":00C7
            Sorted          =   -1  'True
            TabIndex        =   4
            Top             =   2190
            Width           =   5415
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
            Left            =   9120
            Locked          =   -1  'True
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   2190
            Width           =   5415
         End
         Begin VB.Label Label39 
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
            Left            =   11160
            TabIndex        =   116
            Top             =   188
            Width           =   1035
         End
         Begin VB.Label Label29 
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
            Left            =   7755
            TabIndex        =   95
            Top             =   1770
            Width           =   855
         End
         Begin VB.Label Label27 
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
            Left            =   3960
            TabIndex        =   94
            Top             =   1800
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
            Height          =   255
            Left            =   120
            TabIndex        =   59
            Top             =   300
            Width           =   975
         End
         Begin VB.Label Label9 
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
            Left            =   7680
            TabIndex        =   42
            Top             =   795
            Width           =   855
         End
         Begin VB.Label Label7 
            Caption         =   "Client"
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
            TabIndex        =   41
            Top             =   1280
            Width           =   615
         End
         Begin VB.Label Label3 
            Caption         =   "Contract Name"
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
            TabIndex        =   40
            Top             =   790
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "Invoice Month and Year"
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
            Left            =   7680
            TabIndex        =   39
            Top             =   180
            Width           =   1335
         End
         Begin VB.Label Label13 
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
            Left            =   7680
            TabIndex        =   38
            Top             =   1305
            Width           =   1095
         End
         Begin VB.Label Label8 
            Caption         =   "Billing Amount"
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
            TabIndex        =   37
            Top             =   1770
            Width           =   1335
         End
         Begin VB.Label Label12 
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
            Height          =   240
            Left            =   7680
            TabIndex        =   36
            Top             =   2250
            Width           =   720
         End
         Begin VB.Label Label4 
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
            Height          =   225
            Left            =   3960
            TabIndex        =   35
            Top             =   315
            Width           =   1215
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Actual Amount"
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
            TabIndex        =   34
            Top             =   1800
            Width           =   1695
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
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
            TabIndex        =   33
            Top             =   2940
            Width           =   975
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label18 
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
            Height          =   255
            Left            =   11190
            TabIndex        =   32
            Top             =   3600
            Width           =   855
         End
         Begin VB.Label Label21 
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
            Left            =   7800
            TabIndex        =   31
            Top             =   3585
            Width           =   975
         End
         Begin VB.Label Label24 
            Caption         =   "Service Tax"
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
            TabIndex        =   30
            Top             =   2265
            Width           =   1440
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
            Left            =   120
            TabIndex        =   29
            Top             =   3615
            Width           =   975
         End
      End
      Begin VB.Frame ButtonFrm 
         Height          =   1185
         Left            =   120
         TabIndex        =   18
         Top             =   8190
         Width           =   14625
         Begin VB.CommandButton CmdPRint 
            Caption         =   "Frame work Print"
            Height          =   805
            Left            =   13260
            TabIndex        =   96
            Top             =   195
            Width           =   1230
         End
         Begin VB.CommandButton CmdViewInvoice 
            Caption         =   "View Invoice"
            Height          =   400
            Left            =   11790
            TabIndex        =   73
            Top             =   600
            Width           =   1470
         End
         Begin VB.CommandButton CmdReverse 
            Caption         =   "Reverse"
            Height          =   400
            Left            =   11790
            TabIndex        =   72
            Top             =   195
            Width           =   1470
         End
         Begin VB.CommandButton ExitCmd 
            Caption         =   "E&xit"
            Height          =   400
            Left            =   9465
            TabIndex        =   15
            Top             =   600
            Width           =   2310
         End
         Begin VB.CommandButton LastCmd 
            Caption         =   "&Last"
            Height          =   400
            Left            =   7155
            TabIndex        =   14
            Top             =   600
            Width           =   2310
         End
         Begin VB.CommandButton FirstCmd 
            Caption         =   "&First"
            Height          =   400
            Left            =   4845
            TabIndex        =   13
            Top             =   600
            Width           =   2310
         End
         Begin VB.CommandButton PreviousCmd 
            Caption         =   "&Previous"
            Height          =   400
            Left            =   2535
            TabIndex        =   12
            Top             =   600
            Width           =   2310
         End
         Begin VB.CommandButton NextCmd 
            Caption         =   "&Next"
            Height          =   400
            Left            =   225
            TabIndex        =   11
            Top             =   600
            Width           =   2310
         End
         Begin VB.CommandButton AddCmd 
            Caption         =   "&Add New"
            Height          =   400
            Left            =   225
            TabIndex        =   19
            Top             =   195
            Width           =   2310
         End
         Begin VB.CommandButton SaveCmd 
            Caption         =   "Save"
            Height          =   400
            Left            =   2535
            TabIndex        =   8
            Top             =   195
            Width           =   2310
         End
         Begin VB.CommandButton EditCmd 
            Caption         =   "&Edit"
            Height          =   400
            Left            =   4845
            TabIndex        =   0
            Top             =   195
            Width           =   2310
         End
         Begin VB.CommandButton DeleteCmd 
            Caption         =   "&Delete"
            Height          =   400
            Left            =   7155
            TabIndex        =   9
            Top             =   195
            Width           =   2310
         End
         Begin VB.CommandButton SearchCmd 
            Caption         =   "Search"
            Height          =   400
            Left            =   9465
            TabIndex        =   10
            Top             =   195
            Width           =   2310
         End
      End
   End
End
Attribute VB_Name = "FrmTxn_ReservationInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mStatus, mForwardToID, mFlag, mServiceTaxID As Variant
Dim mFromDate, mToDate, vMonth, mBillingMonth, mMonthListindex, mAmount, mSPRateFlag As Variant
Dim mMonth, mTotalDays, mBillingCycleID, mTxtTotDays As Variant
Dim mContractIDRow, mClientIDRow, mSEmployee As Variant
Dim mForwardToEmailID, Edit_Flg As Variant
Dim mEntryNo, mSM_Prefix As Variant
Dim mSMonth, mSYear, mSTotalDays As Variant
Dim mClientID, mLocationID As Variant

Private Sub CmbFramework_LostFocus()
If CmbFramework.Text = "" Then Exit Sub
If LCase(CmbFramework.Text) <> "stock management" And LCase(CmbFramework.Text) <> "non stock management" Then
   MsgBox "Invalid selection!!!"
   CmbFramework.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmbSContract_GotFocus()
tmp = CmbSContract.Text
Call TableMst_StorageContract
CmbSContract.Clear
If RsMst_StorageContract.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_StorageContract.RecordCount
    CmbSContract.AddItem RsMst_StorageContract!ContractName
    RsMst_StorageContract.MoveNext
Next
CmbSContract.Text = tmp
End Sub
Private Sub CmbSContract_LostFocus()
If CmbSContract.Text = "" Then
   Exit Sub
End If
RsMst_StorageContract.MoveFirst
RsMst_StorageContract.Find "ContractName = '" & CmbSContract.Text & "'"
If RsMst_StorageContract.EOF Then
   MsgBox "Invalid selection "
   CmbSContract.SetFocus
   Exit Sub
End If
mContractIDRow = RsMst_StorageContract!ContractID

End Sub

Private Sub CmbNonNcdexClient_GotFocus()
tmp = CmbNonNcdexClient.Text
Call TableMst_Client("where ExchangeTypeId = 1 ") '" & mExchangeTypeID)
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
mClientIDRow = RsMst_Client!ClientID

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
       mStatus = True
       Exit For
    End If
Next

If mStatus = False Then
   MsgBox "Invalid Selection"
   CmbMonth.SetFocus
   Exit Sub
End If
'If CmbMonth.Text = "" Then Exit Sub
'mStatus = False
'For I = 0 To 11
'    If Trim(CmbMonth.Text) = CmbMonth.List(I) Then
'       CmbMonth.Text = CmbMonth.List(I)
'       mMonthListindex = I '+ 1
'       mStatus = True
'       Exit For
'    End If
'Next
'
'If mStatus = False Then
'   MsgBox "Invalid Selection"
'   CmbMonth.SetFocus
'   Exit Sub
'End If
End Sub

Private Sub CmbServiceTax_GotFocus()
tmp = CmbServiceTax.Text
Call TableMst_ServiceTax(" Where Flag = 'A' ")
CmbServiceTax.Clear
If RsMst_ServiceTax.RecordCount = 0 Then
   MsgBox "No Employee Found!!!"
   CmbServiceTax.SetFocus
   'RsMst_Employee.MoveFirst
   Exit Sub
End If
For I = 1 To RsMst_ServiceTax.RecordCount
    CmbServiceTax.AddItem RsMst_ServiceTax!Description
    RsMst_ServiceTax.MoveNext
Next

CmbServiceTax.Text = tmp
End Sub

Private Sub CmbServiceTax_LostFocus()
If CmbServiceTax.Text = "" Then
   Exit Sub
End If
RsMst_ServiceTax.MoveFirst
RsMst_ServiceTax.Find "Description = '" & CmbServiceTax.Text & "'"
If RsMst_ServiceTax.EOF Then
   MsgBox "Invalid selection "
   CmbServiceTax.SetFocus
   Exit Sub
End If
TxtNAVCode.Text = RsMst_ServiceTax!NavCode
mServiceTaxID = RsMst_ServiceTax!ServiceTaxID
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
'Call ClearFrame
'Call Grid_Head
'Frame4.Enabled = False
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

End Sub
Private Sub CmbSLocation_GotFocus()
tmp = CmbSLocation.Text

If Gen_WcLocation <> "" Then
   Call TableMst_Location(" Where " & Gen_WcLocation)   '(" Where LocationId > 0 ")
Else
   Call TableMst_Location
End If
CmbSLocation.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbSLocation.AddItem RsMst_Location!Locationname
    RsMst_Location.MoveNext
Next
CmbSLocation.Text = tmp
End Sub

Private Sub CmbSLocation_LostFocus()
If CmbSLocation.Text = "" Then
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbSLocation.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection "
   CmbSLocation.SetFocus
   Exit Sub
End If
mSLocationID = RsMst_Location!LocationID
End Sub
Private Sub CmbSStatus_LostFocus()
If CmbSStatus.Text = "" Then Exit Sub

If LCase(CmbSStatus.Text) = "approved" Then
   mSStatus = "A"
ElseIf LCase(CmbSStatus.Text) = "pending" Then
   mSStatus = "P"
ElseIf LCase(CmbSStatus.Text) = "freeze" Then
   mSStatus = "F"
Else
   MsgBox "Invalid selection!!!"
   CmbSStatus.SetFocus
   Exit Sub
End If
End Sub


Private Sub Grid_Head2()
MSHFlexGrid3.Clear
MSHFlexGrid3.Rows = 3
MSHFlexGrid3.FixedRows = 1

MSHFlexGrid3.Cols = 4
MSHFlexGrid3.Font.Size = 10

MSHFlexGrid3.TextMatrix(0, 0) = "Invoice No"
MSHFlexGrid3.TextMatrix(0, 1) = "Updated By "
MSHFlexGrid3.TextMatrix(0, 2) = "Updated On"
MSHFlexGrid3.TextMatrix(0, 3) = "Remarks"

MSHFlexGrid3.ColWidth(0) = 800 '500
MSHFlexGrid3.ColWidth(1) = 1280
MSHFlexGrid3.ColWidth(2) = 2050
MSHFlexGrid3.ColWidth(3) = 5000
MSHFlexGrid3.RowHeight(0) = 600
MSHFlexGrid3.WordWrap = True
End Sub


Private Sub CmbStatus_LostFocus()
If CmbStatus.Text = "" Then Exit Sub
'FREEZE
If LCase(CmbStatus.Text) = "approved" Then
   mStatus = "A"
ElseIf LCase(CmbStatus.Text) = "pending" Then
   mStatus = "P"
ElseIf LCase(CmbStatus.Text) = LCase("FREEZE") Then
   mStatus = "F"
Else
   MsgBox "Invalid selection!!!"
   CmbStatus.SetFocus
   Exit Sub
End If

If LCase(Gen_UserRole) <> "fag" And LCase(Gen_UserRole) <> "fagadmin" And LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "admin" And LCase(Gen_UserRole) <> "superadmin" And LCase(Gen_UserRole) = "headcmg" Then
   If mStatus = "A" Then
      MsgBox "Please forward this Invoice to CSG department / FAG department, Only CSG/FAG department have access to approve invoice '!!!"
      CmbStatus.SetFocus
      Exit Sub
   End If
End If

End Sub


Private Sub CmdCancelFrameWork_Click()
Frame3.Visible = False
End Sub

Private Sub CmdCancelLog_Click()
Frame5.Visible = False
TxtLogRemark.Text = ""
End Sub

Private Sub CmdCancelReverse_Click()
TxtReverseDate.Value = Date
TxtReverseRemark = ""
Frame6.Visible = False
End Sub

Private Sub CmdDeleteLog_Click()

If Trim(TxtLogRemark.Text) = "" Then
   MsgBox "Log Remark not allowed!!!"
   TxtLogRemark.SetFocus
   Exit Sub
End If
Call TableLog_InvoiceDeletetion(" Where DocumentType = ''")
Call TableMst_Location
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationID = " & RsTxn_ReservationInvoice!LocationID
RsLog_InvoiceDeletetion.AddNew
RsLog_InvoiceDeletetion!DocumentType = "Reservation"
RsLog_InvoiceDeletetion!DcoumentID = RsTxn_ReservationInvoice!InvoiceNo
RsLog_InvoiceDeletetion!SM_Prefix = RsMst_Location!SM_Prefix
RsLog_InvoiceDeletetion!ClientID = RsTxn_ReservationInvoice!ContractID
RsLog_InvoiceDeletetion!Amount = RsTxn_ReservationInvoice!ActualAmount
RsLog_InvoiceDeletetion!DeletetionRemark = Trim(TxtLogRemark.Text)
RsLog_InvoiceDeletetion!DeletedBy = Gen_UserLoginID
RsLog_InvoiceDeletetion!DeletedOn = Now
RsLog_InvoiceDeletetion.Update

tmp = "Delete From Log_ReservationInvoice Where InvoiceNo= " & TxtInvoiceNo.Text & ""
Call TmpTable

tmp = "Delete From Txn_ReservationInvoiceDetail Where InvoiceNo= " & TxtInvoiceNo.Text & ""
Call TmpTable

tmp = "Delete From Txn_ReservationDailyStock Where ContractID=" & mContractIDRow & " And MonthYear='" & RsTxn_ReservationInvoice!YearMonth & "'"
Call TmpTable

tmp = "Update Txn_InvMonGSLDetail set invoiceAmount=null,PINo=Null Where YearMonth='" & Format(RsTxn_ReservationInvoice!YearMonth, Gen_DatFormat) & "' And ContractID=" & mContractIDRow & " And PINo=" & Val(TxtInvoiceNo) & ""
Call TmpTable

RsTxn_ReservationInvoice.Delete
RsTxn_ReservationInvoice.Update
RsTxn_ReservationInvoice.MoveNext
If RsTxn_ReservationInvoice.EOF() Then
   RsTxn_ReservationInvoice.MoveLast
End If

Frame5.Visible = False
TxtLogRemark.Text = ""

Call Indata



End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Txn_ReservationInvoice.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Txn_ReservationInvoice.xls")
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

Call Grid_Head3
'MSHFlexGrid1.Clear
mSEmployee = ""

'If CmbMonth.Text = "" Then
'   MsgBox " Please Select Month"
'   CmbMonth.SetFocus
'   Exit Sub
'End If
'
'If CmbYear.Text = "" Then
'   MsgBox " Please Select Year"
'   CmbYear.SetFocus
'   Exit Sub
'End If

'If mMonthListindex < 10 Then
'   mMonthListindex = "0" & mMonthListindex
'End If
'
'mMonthListindex = CmbMonth.ListIndex + 1
'vMonth = Month(Now)
'
'
'mYear = Val(CmbYear.Text)
'mTotalDays = Day(DateSerial(mYear, mMonthListindex + 1, 0))

'mFromDate = Format(CmbYear.Text & "-" & mMonthListindex & "-" & "01", "yyyy-mm-dd")
'mToDate = Format(CmbYear.Text & "-" & mMonthListindex & "-" & mTotalDays, "yyyy-mm-dd")


tmp = " Select ML.LocationName,TRI.Flag,TRI.InvoiceNo,TRI.TINO,TRI.TIAmount,TRI.TIDate,TRI.ActualAmount,MSC.ContractName,"
tmp = tmp & " TRI.BillingAmount ,TRI.BillingDate,TRI.BatchID , TRI.YearMonth, TRI.Remarks,"
tmp = tmp & " MC.Commodity , MCL.ClientNAme, " '",TRI.ApprovalRemarks ,TRI.ApprovalDate"
tmp = tmp & " TTD.TIID,TTD.TINo,TTD.TIDate,TTD.TIAmount,TTD.TISTAmount,TTD.EntryNo, "
tmp = tmp & " (Select isnull(Sum(TCRT.RecAmount),0) as RecAmount From Txn_NewCashReceiptTIDetail TCRT"
tmp = tmp & " Inner Join Txn_NewCashReceipt TCR on TCR.NewCRID= TCRT.NewCRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'And TIID=TTD.TIID)+isnull(TTD.AdjAmountFAG,0) + "
tmp = tmp & " (Select isnull(Sum(TCRT.RecTDS),0) as RecTDS From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.NewCRID= TCRT.NewCRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'And TIID=TTD.TIID ) + (Select isnull(Sum(TCRT.RecTDS),0) as RecTDS From Txn_CRTDSAdjustmentDetail TCRT Inner Join Txn_CRTDSAdjustment TCR on TCR.CRTDSAdjustmentID= TCRT.CRTDSAdjustmentID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.Flag<>'R'And TIID=TTD.TIID ) +"
tmp = tmp & " (Select isnull(Sum(TCRT.RecOtherDed),0) as RecOtherDed From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.NewCRID= TCRT.NewCRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'And TIID=TTD.TIID )+ (Select isnull(Sum(TCRT.RecOtherDed),0) as RecOtherDed From Txn_CRTDSAdjustmentDetail TCRT Inner Join Txn_CRTDSAdjustment TCR on TCR.CRTDSAdjustmentID= TCRT.CRTDSAdjustmentID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.Flag<>'R'And TIID=TTD.TIID) 'ReceivedAmount','' 'BalanceAmount' ,"
tmp = tmp & " TRI.ReverseFlag,ReverseRemark,ReverseDate,RBatchID,TRI.ServiceTaxNAVCode "
''-- Ullhas Patil 27 Sep 2017
tmp = tmp & " ,MSI.StateIndiaName ,MSI.GSTIN 'NCMLGSTNo',MCSG.GSTNo 'ClientGSTNo',MCSG.RGSTNAVCode,'' 'TaxApplied' "
''-------------------------
tmp = tmp & " from Txn_ReservationInvoice TRI"
tmp = tmp & " inner join Mst_StorageContract MSC on MSC.ContractID=TRI.ContractID"
tmp = tmp & " Left join MSt_Commodity MC on MC.CommodityID=MSC.CommodityID"
tmp = tmp & " inner join Mst_Client MCL on MCL.ClientID=MSC.ClientID"
tmp = tmp & " inner join mst_Location ML on ML.LocationID=TRI.LocationID"
'tmp = tmp & " inner join Mst_Employee ME on ME.Employeeid=TRI.ForwardToID"
'tmp = tmp & " Where InvoiceNo=" & RsTxn_ReservationInvoice!InvoiceNo & " "
tmp = tmp & " Left Join Txn_TIDetails TTD On TRI.InvoiceNo =  TTD.PIID And PIFlag = 'Reservation' "
''-- Ullhas Patil 27 Sep 2017
tmp = tmp & " Inner Join Mst_State MS on ML.StateID = MS.StateID "
tmp = tmp & " Inner Join Mst_StateIndia MSI On MS.StateIndiaID = MSI.StateIndiaID "
tmp = tmp & " Left  Join Mst_ClientStateWiseGST MCSG On MSC.ClientID = MCSG.ClientID ANd MSI.StateIndiaID = MCSG.StateINdiaID "
''-------------------

If Gen_DBType = "MDB" Then
   tmp = tmp & " Where BillingDate Between #" & TxtFromDate.Value & "# And #" & TxtToDate.Value & "# "     'And ExchangeTypeID = 1 And (TRI.ContractID = 0 Or TRI.ContractID is Null)"
Else
   tmp = tmp & " Where BillingDate Between '" & TxtFromDate.Value & "' And '" & TxtToDate.Value & "'" 'And ExchangeTypeID = 1 And (TRI.ContractID = 0 Or TRI.ContractID is Null)"
End If

If CmbSContract.Text <> "" Then
   tmp = tmp & " And ContractName = '" & CmbSContract.Text & "'"
End If

If CmbSLocation.Text <> "" Then
   tmp = tmp & " And LocationName = '" & CmbSLocation.Text & "'"
Else
   If Gen_WcLocation <> "" Then
      tmp = tmp & " And ML." & Gen_WcLocation
   End If
End If

If CmbNonNcdexClient.Text <> "" Then
   tmp = tmp & " And MCL.ClientNAme = '" & CmbNonNcdexClient.Text & "'"
End If



'If Trim(CmbSForwardTo.Text) <> "" Then
'
'   strSplitString = Split(CmbSForwardTo, "~")
'   EmpName_ = Trim(strSplitString(0))
'   EmpNo_ = Trim(strSplitString(1))
'   tmp = tmp & " And ME.EmployeeNo = '" & EmpNo_ & "'"
'End If

mFlag = Left(CmbSStatus.Text, 1)
If CmbSStatus.Text <> "" Then
   tmp = tmp & " And TRI.Flag = '" & mFlag & "'"
End If

If ChkBatchNo.Value = 0 Then
   'tmp = tmp & " AND CS.BatchID = 0 "
   tmp = tmp & " And BatchID = 0 "
End If

If TxtSInvoiceNo <> "" Then
   tmp = tmp & " And TRI.InvoiceNo = " & TxtSInvoiceNo & " "
End If

tmp = tmp & " Order By InvoiceNo"
Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   MSHFlexGrid1.Rows = TmpRs1.RecordCount + 2
   For I = 1 To TmpRs1.RecordCount
       MSHFlexGrid1.TextMatrix(I, 0) = TmpRs1!InvoiceNo
       MSHFlexGrid1.TextMatrix(I, 1) = TmpRs1!ContractName
       'mLocationId = TmpRs1!LocationID
       
       MSHFlexGrid1.TextMatrix(I, 2) = TmpRs1!Locationname
       MSHFlexGrid1.TextMatrix(I, 3) = IIf(IsNull(TmpRs1!Commodity), "", TmpRs1!Commodity)
'       MSHFlexGrid1.TextMatrix(I, 4) = TmpRs1!YearMonth
       MSHFlexGrid1.TextMatrix(I, 4) = Format(TmpRs1!YearMonth, "dd-mm-yyyy")
'       MSHFlexGrid1.TextMatrix(I, 5) = TmpRs1!BillingDate
       MSHFlexGrid1.TextMatrix(I, 5) = Format(TmpRs1!BillingDate, "dd-mm-yyyy")
       MSHFlexGrid1.TextMatrix(I, 6) = IIf(IsNull(TmpRs1!BillingAmount), 0, TmpRs1!BillingAmount)
       MSHFlexGrid1.TextMatrix(I, 7) = TmpRs1!ClientName
       MSHFlexGrid1.TextMatrix(I, 8) = IIf(IsNull(TmpRs1!ActualAmount), "", TmpRs1!ActualAmount) '
       MSHFlexGrid1.TextMatrix(I, 9) = GetServiceTaxDescFromNAVCode(TmpRs1!ServiceTaxNAVCode) 'IIf(IsNull(TmpRs1!TINo), "", TmpRs1!TINo) 'TmpRs1!TINO
       MSHFlexGrid1.TextMatrix(I, 10) = "" 'IIf(IsNull(TmpRs1!TIDate), "", Format(TmpRs1!TIDate, "dd-mm-yyyy")) '
'       MSHFlexGrid1.TextMatrix(I, 10) = Format(TmpRs1!TIDate, "dd-mm-yyyy")
       MSHFlexGrid1.TextMatrix(I, 11) = "" 'IIf(IsNull(TmpRs1!TIAmount), "", TmpRs1!TIAmount)
       MSHFlexGrid1.TextMatrix(I, 12) = IIf(IsNull(TmpRs1!BatchID), 0, TmpRs1!BatchID) '
       MSHFlexGrid1.TextMatrix(I, 13) = TmpRs1!Flag
       If LCase(MSHFlexGrid1.TextMatrix(I, 13)) = "p" Then
          MSHFlexGrid1.TextMatrix(I, 13) = "Pending"
       ElseIf LCase(MSHFlexGrid1.TextMatrix(I, 13)) = "a" Then
          MSHFlexGrid1.TextMatrix(I, 13) = "Approved"
       End If
       
       MSHFlexGrid1.TextMatrix(I, 14) = IIf(IsNull(TmpRs1!Remarks), "", TmpRs1!Remarks)
       
       MSHFlexGrid1.TextMatrix(I, 15) = IIf(IsNull(TmpRs1!TIID), "", TmpRs1!TIID)
       MSHFlexGrid1.TextMatrix(I, 16) = IIf(IsNull(TmpRs1!TINo), "", TmpRs1!TINo)
       MSHFlexGrid1.TextMatrix(I, 17) = IIf(IsNull(TmpRs1!TIDate), "", Format(TmpRs1!TIDate, "dd-mmm-yyyy"))
       MSHFlexGrid1.TextMatrix(I, 18) = IIf(IsNull(TmpRs1!TIAmount), "", TmpRs1!TIAmount)
       MSHFlexGrid1.TextMatrix(I, 19) = IIf(IsNull(TmpRs1!TISTAmount), "", TmpRs1!TISTAmount)
       MSHFlexGrid1.TextMatrix(I, 20) = IIf(IsNull(TmpRs1!EntryNo), "", TmpRs1!EntryNo)
       MSHFlexGrid1.TextMatrix(I, 21) = IIf(IsNull(TmpRs1!ReceivedAmount), "", TmpRs1!ReceivedAmount)
       MSHFlexGrid1.TextMatrix(I, 22) = Val(MSHFlexGrid1.TextMatrix(I, 18)) - Val(MSHFlexGrid1.TextMatrix(I, 21))
       
       MSHFlexGrid1.TextMatrix(I, 23) = IIf(IsNull(TmpRs1!ReverseFlag), "", TmpRs1!ReverseFlag)
       If LCase(MSHFlexGrid1.TextMatrix(I, 23)) = "y" Then
          MSHFlexGrid1.TextMatrix(I, 23) = "Yes"
       Else
          MSHFlexGrid1.TextMatrix(I, 23) = "No"
       End If
       MSHFlexGrid1.TextMatrix(I, 24) = IIf(IsNull(TmpRs1!ReverseRemark), "", TmpRs1!ReverseRemark)
       MSHFlexGrid1.TextMatrix(I, 25) = IIf(IsNull(TmpRs1!ReverseDate), "", Format(TmpRs1!ReverseDate, "dd-mmm-yyyy"))
       MSHFlexGrid1.TextMatrix(I, 26) = IIf(IsNull(TmpRs1!RBatchID), "", TmpRs1!RBatchID)

       ''-- Ullhas Patil 27 Sep 2017
       MSHFlexGrid1.TextMatrix(I, 27) = IIf(IsNull(TmpRs1!StateIndiaName), "", TmpRs1!StateIndiaName)   '"State India Name"
       MSHFlexGrid1.TextMatrix(I, 28) = IIf(IsNull(TmpRs1!NCMLGSTNo), "", TmpRs1!NCMLGSTNo)             '"NCML GST No"
       MSHFlexGrid1.TextMatrix(I, 29) = IIf(IsNull(TmpRs1!ClientGSTNo), "", TmpRs1!ClientGSTNo)         '"Client GST No"
       MSHFlexGrid1.TextMatrix(I, 30) = IIf(IsNull(TmpRs1!RGSTNavCode), "", TmpRs1!RGSTNavCode)           '"Client GST NAVCode"
       MSHFlexGrid1.TextMatrix(I, 31) = "" '"Tax Applied (Inter/Intra)"
       
       If Trim(MSHFlexGrid1.TextMatrix(I, 28)) <> "" And Trim(MSHFlexGrid1.TextMatrix(I, 29)) <> "" Then
          If Left(MSHFlexGrid1.TextMatrix(I, 28), 2) = Left(MSHFlexGrid1.TextMatrix(I, 29), 2) Then
             MSHFlexGrid1.TextMatrix(I, 31) = "Intra"
          Else
             MSHFlexGrid1.TextMatrix(I, 31) = "Inter"
          End If
       End If
       
       ''-----------------
       
       TmpRs1.MoveNext
   Next
    
End If
'If TmpRs1.RecordCount > 0 Then
'   MSHFlexGrid1.Rows = TmpRs1.RecordCount + 2
'   For I = 1 To TmpRs1.RecordCount
'       For c = 0 To MSHFlexGrid1.Cols - 1
'           MSHFlexGrid1.TextMatrix(I, c) = IIf(IsNull(TmpRs1.Fields(c)), "", TmpRs1.Fields(c))
'       Next
'       TmpRs1.MoveNext
'   Next
'End If
Label22.Caption = ""
Label22.Caption = TmpRs1.RecordCount
Label22.Refresh

End Sub

Private Sub CmdPrint_Click()
Frame3.Visible = True
CmbFramework.Text = ""
End Sub

Public Sub Xls_Detail_Rpt2()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Integer
Dim mCol As Integer
Dim mAmount As Double
MsgBox ("Wait till next message")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "BillFrameWork.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows

If LCase(CmbFramework.Text) = "non stock management" Then
   oWS.Name = "Non Stock Management"
   oWS.Range("a1:g1").Merge
   oWS.Range("a:a").ColumnWidth = 7.86
   oWS.Range("b:b").ColumnWidth = 15.14
   oWS.Range("c:c").ColumnWidth = 17.43
   oWS.Range("d:d").ColumnWidth = 12.14
   oWS.Range("e:e").ColumnWidth = 12.71
   oWS.Range("f:f").ColumnWidth = 9.43
   oWS.Range("g:g").ColumnWidth = 10.29
   
   
   mRow = 2
   tmp = "a" & mRow & ":" & "g" & mRow
   oWS.Range(tmp).Merge
   oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
   oWS.Cells(mRow, 1) = TxtClient.Text & " - Storage charges as per the agreement"
   oWS.Cells(mRow, 1).HorizontalAlignment = -4108
   oWS.Cells(mRow, 1).VerticalAlignment = -4108
   oWS.Cells(mRow, 1).Font.Bold = True
   oWS.Cells(mRow, 1).Font.Size = 14
   oWS.Cells(mRow, C + 1).RowHeight = 30
   
   mRow = mRow + 2
   tmp = "a" & mRow & ":" & "b" & mRow
   oWS.Range(tmp).Merge
   oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
   oWS.Cells(mRow, 1) = "Location"
   oWS.Cells(mRow, 1).Font.Bold = True
   oWS.Cells(mRow, 1).Font.Size = 12
   tmp = "c" & mRow & ":" & "g" & mRow
   oWS.Range(tmp).Merge
   oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
   oWS.Cells(mRow, 3) = TxtLocation.Text
   oWS.Cells(mRow, 3).Font.Bold = True
   oWS.Cells(mRow, 3).Font.Size = 12
   
   'mRow = mRow + 1
   'tmp = "a" & mRow & ":" & "b" & mRow
   'oWS.Range(tmp).Merge
   'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
   'oWS.Cells(mRow, 1) = "Warehouse Name"
   'oWS.Cells(mRow, 1).Font.Bold = True
   'oWS.Cells(mRow, 1).Font.Size = 12
   'tmp = "c" & mRow & ":" & "g" & mRow
   'oWS.Range(tmp).Merge
   'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
   'oWS.Cells(mRow, 3) = "Warehouse Name"
   'oWS.Cells(mRow, 3).Font.Bold = True
   'oWS.Cells(mRow, 3).Font.Size = 12
   
   mRow = mRow + 1
   tmp = "a" & mRow & ":" & "b" & mRow
   oWS.Range(tmp).Merge
   oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
   oWS.Cells(mRow, 1) = "Commodity"
   oWS.Cells(mRow, 1).Font.Bold = True
   oWS.Cells(mRow, 1).Font.Size = 12
   tmp = "c" & mRow & ":" & "g" & mRow
   oWS.Range(tmp).Merge
   oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
   oWS.Cells(mRow, 3) = TxtCommodity.Text
   oWS.Cells(mRow, 3).Font.Bold = True
   oWS.Cells(mRow, 3).Font.Size = 12
   
   mRow = mRow + 2
   oWS.Cells(mRow, 1).RowHeight = 30
   oWS.Cells(mRow, 1).WrapText = True
   oWS.Cells(mRow, 1).Borders.Color = RGB(0, 0, 0)
   oWS.Cells(mRow, 1) = "Sr. No"
   oWS.Cells(mRow, 1).Font.Bold = True
   oWS.Cells(mRow, 1).Font.Size = 12
   oWS.Cells(mRow, 2).WrapText = True
   oWS.Cells(mRow, 2).Borders.Color = RGB(0, 0, 0)
   oWS.Cells(mRow, 2) = "Period"
   oWS.Cells(mRow, 2).Font.Bold = True
   oWS.Cells(mRow, 2).Font.Size = 12
   oWS.Cells(mRow, 3).WrapText = True
   oWS.Cells(mRow, 3).Borders.Color = RGB(0, 0, 0)
   oWS.Cells(mRow, 3) = "Reserve Qty/ Peak Qty (MT)"
   oWS.Cells(mRow, 3).Font.Bold = True
   oWS.Cells(mRow, 3).Font.Size = 12
   oWS.Cells(mRow, 4).WrapText = True
   oWS.Cells(mRow, 4).Borders.Color = RGB(0, 0, 0)
   oWS.Cells(mRow, 4) = "Rate ( per MT month)"
   oWS.Cells(mRow, 4).Font.Bold = True
   oWS.Cells(mRow, 4).Font.Size = 12
   oWS.Cells(mRow, 5).WrapText = True
   oWS.Cells(mRow, 5).Borders.Color = RGB(0, 0, 0)
   oWS.Cells(mRow, 5) = "No of Days"
   oWS.Cells(mRow, 5).Font.Bold = True
   oWS.Cells(mRow, 5).Font.Size = 12
   oWS.Cells(mRow, 6).WrapText = True
   oWS.Cells(mRow, 6).Borders.Color = RGB(0, 0, 0)
   oWS.Cells(mRow, 6) = "Amount"
   oWS.Cells(mRow, 6).Font.Bold = True
   oWS.Cells(mRow, 6).Font.Size = 12
   oWS.Cells(mRow, 7).WrapText = True
   oWS.Cells(mRow, 7).Borders.Color = RGB(0, 0, 0)
   oWS.Cells(mRow, 7) = "Remarks"
   oWS.Cells(mRow, 7).Font.Bold = True
   oWS.Cells(mRow, 7).Font.Size = 12
   
   With MSHFlexGrid2
      mAmount = 0
      For I = 1 To .Rows - 1
         mRow = mRow + 1
         oWS.Cells(mRow, 1).RowHeight = 30
         oWS.Cells(mRow, 1).WrapText = True
         oWS.Cells(mRow, 1).Borders.Color = RGB(0, 0, 0)
         oWS.Cells(mRow, 1) = I
         oWS.Cells(mRow, 1).Font.Size = 12
         oWS.Cells(mRow, 2).WrapText = True
         oWS.Cells(mRow, 2).Borders.Color = RGB(0, 0, 0)
         oWS.Cells(mRow, 2) = .TextMatrix(I, 0) & " To " & .TextMatrix(I, 1)
         oWS.Cells(mRow, 2).Font.Size = 12
         oWS.Cells(mRow, 3).WrapText = True
         oWS.Cells(mRow, 3).Borders.Color = RGB(0, 0, 0)
         oWS.Cells(mRow, 3) = .TextMatrix(I, 4)
         oWS.Cells(mRow, 3).Font.Size = 12
         oWS.Cells(mRow, 4).WrapText = True
         oWS.Cells(mRow, 4).Borders.Color = RGB(0, 0, 0)
         oWS.Cells(mRow, 4) = .TextMatrix(I, 5)
         oWS.Cells(mRow, 4).Font.Size = 12
         oWS.Cells(mRow, 5).WrapText = True
         oWS.Cells(mRow, 5).Borders.Color = RGB(0, 0, 0)
         If .TextMatrix(I, 0) <> "" Then
            oWS.Cells(mRow, 5) = DateDiff("d", CDate(.TextMatrix(I, 0)), CDate(.TextMatrix(I, 1))) + 1
         Else
            oWS.Cells(mRow, 5) = ""
         End If
         oWS.Cells(mRow, 5).Font.Size = 12
         oWS.Cells(mRow, 6).WrapText = True
         oWS.Cells(mRow, 6).Borders.Color = RGB(0, 0, 0)
         oWS.Cells(mRow, 6) = Format(Val(.TextMatrix(I, 6)), "#0")
         mAmount = mAmount + Val(.TextMatrix(I, 6))
         oWS.Cells(mRow, 6).Font.Size = 12
         oWS.Cells(mRow, 7).WrapText = True
         oWS.Cells(mRow, 7).Borders.Color = RGB(0, 0, 0)
         oWS.Cells(mRow, 7) = ""
         oWS.Cells(mRow, 7).Font.Size = 12
      Next
   End With
   
   mRow = mRow + 2
   tmp = "a" & mRow & ":" & "e" & mRow
   oWS.Range(tmp).Merge
   oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
   oWS.Cells(mRow, 1) = "Grand Total"
   oWS.Cells(mRow, 1).Font.Bold = True
   oWS.Cells(mRow, 1).Font.Size = 12
   oWS.Cells(mRow, 1).HorizontalAlignment = -4108
   tmp = "f" & mRow & ":" & "f" & mRow
   oWS.Range(tmp).Merge
   oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
   oWS.Cells(mRow, 6) = Format(Val(mAmount), "#0")
   oWS.Cells(mRow, 6).Font.Bold = True
   oWS.Cells(mRow, 6).Font.Size = 12
      
   mRow = mRow + 3
   tmp = "a" & mRow & ":" & "g" & mRow
   oWS.Range(tmp).Merge
   oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
   oWS.Cells(mRow, 1) = "Rupees:" & NumStrConv(Val(oWS.Cells(mRow - 3, 6)))
   oWS.Cells(mRow, 1).Font.Bold = True
   oWS.Cells(mRow, 1).Font.Size = 12
   
   tmp = "Select MSC.Remarks from Mst_StorageContract MSC Where MSC.ContractID = " & Val(mContractIDRow)
   Call TmpTable
   mRow = mRow + 3
   tmp = "a" & mRow & ":" & "g" & mRow
   oWS.Range(tmp).Merge
   oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
   If TmpRs.RecordCount > 0 Then
      oWS.Cells(mRow, 1) = IIf(IsNull(TmpRs!Remarks), "", TmpRs!Remarks)
   End If
   oWS.Cells(mRow, 1).Font.Bold = True
   oWS.Cells(mRow, 1).Font.Size = 10
   oWS.Cells(mRow, 1).WrapText = True
   oWS.Cells(mRow, 1).RowHeight = 30
   
   
ElseIf LCase(CmbFramework.Text) = "stock management" Then
   oWS.Name = "Stock Management"
   oWS.Range("a1:g1").Merge
   oWS.Range("a:a").ColumnWidth = 7.86
   oWS.Range("b:b").ColumnWidth = 15.14
   oWS.Range("c:c").ColumnWidth = 17.43
   oWS.Range("d:d").ColumnWidth = 12.14
   oWS.Range("e:e").ColumnWidth = 12.71
   oWS.Range("f:f").ColumnWidth = 9.43
   oWS.Range("g:g").ColumnWidth = 10.29
   oWS.Range("h:h").ColumnWidth = 10.29
      
   mRow = 2
   tmp = "a" & mRow & ":" & "h" & mRow
   oWS.Range(tmp).Merge
   oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
   oWS.Cells(mRow, 1) = TxtClient.Text & " - Storage charges as per the agreement"
   oWS.Cells(mRow, 1).HorizontalAlignment = -4108
   oWS.Cells(mRow, 1).VerticalAlignment = -4108
   oWS.Cells(mRow, 1).Font.Bold = True
   oWS.Cells(mRow, 1).Font.Size = 14
   oWS.Cells(mRow, C + 1).RowHeight = 30
   
   mRow = mRow + 2
   tmp = "a" & mRow & ":" & "b" & mRow
   oWS.Range(tmp).Merge
   oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
   oWS.Cells(mRow, 1) = "Location"
   oWS.Cells(mRow, 1).Font.Bold = True
   oWS.Cells(mRow, 1).Font.Size = 12
   tmp = "c" & mRow & ":" & "h" & mRow
   oWS.Range(tmp).Merge
   oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
   oWS.Cells(mRow, 3) = TxtLocation.Text
   oWS.Cells(mRow, 3).Font.Bold = True
   oWS.Cells(mRow, 3).Font.Size = 12
   
   'mRow = mRow + 1
   'tmp = "a" & mRow & ":" & "b" & mRow
   'oWS.Range(tmp).Merge
   'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
   'oWS.Cells(mRow, 1) = "Warehouse Name"
   'oWS.Cells(mRow, 1).Font.Bold = True
   'oWS.Cells(mRow, 1).Font.Size = 12
   'tmp = "c" & mRow & ":" & "g" & mRow
   'oWS.Range(tmp).Merge
   'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
   'oWS.Cells(mRow, 3) = "Warehouse Name"
   'oWS.Cells(mRow, 3).Font.Bold = True
   'oWS.Cells(mRow, 3).Font.Size = 12
   
   mRow = mRow + 1
   tmp = "a" & mRow & ":" & "b" & mRow
   oWS.Range(tmp).Merge
   oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
   oWS.Cells(mRow, 1) = "Commodity"
   oWS.Cells(mRow, 1).Font.Bold = True
   oWS.Cells(mRow, 1).Font.Size = 12
   tmp = "c" & mRow & ":" & "h" & mRow
   oWS.Range(tmp).Merge
   oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
   oWS.Cells(mRow, 3) = TxtCommodity.Text
   oWS.Cells(mRow, 3).Font.Bold = True
   oWS.Cells(mRow, 3).Font.Size = 12
   
   mRow = mRow + 2
   oWS.Cells(mRow, 1).RowHeight = 30
   oWS.Cells(mRow, 1).WrapText = True
   oWS.Cells(mRow, 1).Borders.Color = RGB(0, 0, 0)
   oWS.Cells(mRow, 1) = "Sr. No"
   oWS.Cells(mRow, 1).Font.Bold = True
   oWS.Cells(mRow, 1).Font.Size = 12
   oWS.Cells(mRow, 2).WrapText = True
   oWS.Cells(mRow, 2).Borders.Color = RGB(0, 0, 0)
   oWS.Cells(mRow, 2) = "Particulars"
   oWS.Cells(mRow, 2).Font.Bold = True
   oWS.Cells(mRow, 2).Font.Size = 12
   oWS.Cells(mRow, 3).WrapText = True
   oWS.Cells(mRow, 3).Borders.Color = RGB(0, 0, 0)
   oWS.Cells(mRow, 3) = "Godown Name"
   oWS.Cells(mRow, 3).Font.Bold = True
   oWS.Cells(mRow, 3).Font.Size = 12
   oWS.Cells(mRow, 4).WrapText = True
   oWS.Cells(mRow, 4).Borders.Color = RGB(0, 0, 0)
   oWS.Cells(mRow, 4) = "Period"
   oWS.Cells(mRow, 4).Font.Bold = True
   oWS.Cells(mRow, 4).Font.Size = 12
   oWS.Cells(mRow, 5).WrapText = True
   oWS.Cells(mRow, 5).Borders.Color = RGB(0, 0, 0)
   oWS.Cells(mRow, 5) = "Rate per month"
   oWS.Cells(mRow, 5).Font.Bold = True
   oWS.Cells(mRow, 5).Font.Size = 12
   oWS.Cells(mRow, 6).WrapText = True
   oWS.Cells(mRow, 6).Borders.Color = RGB(0, 0, 0)
   oWS.Cells(mRow, 6) = "No of Days"
   oWS.Cells(mRow, 6).Font.Bold = True
   oWS.Cells(mRow, 6).Font.Size = 12
   oWS.Cells(mRow, 7).WrapText = True
   oWS.Cells(mRow, 7).Borders.Color = RGB(0, 0, 0)
   oWS.Cells(mRow, 7) = "Amount"
   oWS.Cells(mRow, 7).Font.Bold = True
   oWS.Cells(mRow, 7).Font.Size = 12
   oWS.Cells(mRow, 8).WrapText = True
   oWS.Cells(mRow, 8).Borders.Color = RGB(0, 0, 0)
   oWS.Cells(mRow, 8) = "Remarks"
   oWS.Cells(mRow, 8).Font.Bold = True
   oWS.Cells(mRow, 8).Font.Size = 12
   
   With MSHFlexGrid2
      mAmount = 0
      For I = 1 To .Rows - 1
         mRow = mRow + 1
         oWS.Cells(mRow, 1).RowHeight = 45
         oWS.Cells(mRow, 1).WrapText = True
         oWS.Cells(mRow, 1).Borders.Color = RGB(0, 0, 0)
         oWS.Cells(mRow, 1) = I
         oWS.Cells(mRow, 1).Font.Size = 12
         oWS.Cells(mRow, 2).WrapText = True
         oWS.Cells(mRow, 2).Borders.Color = RGB(0, 0, 0)
         oWS.Cells(mRow, 2) = "Service Charges for stock management" '
         oWS.Cells(mRow, 2).Font.Size = 12
         oWS.Cells(mRow, 3).WrapText = True
         oWS.Cells(mRow, 3).Borders.Color = RGB(0, 0, 0)
         oWS.Cells(mRow, 3) = ""
         oWS.Cells(mRow, 3).Font.Size = 12
         oWS.Cells(mRow, 4).WrapText = True
         oWS.Cells(mRow, 4).Borders.Color = RGB(0, 0, 0)
         oWS.Cells(mRow, 4) = .TextMatrix(I, 0) & " To " & .TextMatrix(I, 1)
         oWS.Cells(mRow, 4).Font.Size = 12
         oWS.Cells(mRow, 5).WrapText = True
         oWS.Cells(mRow, 5).Borders.Color = RGB(0, 0, 0)
         oWS.Cells(mRow, 5) = .TextMatrix(I, 5)
         oWS.Cells(mRow, 5).Font.Size = 12
         oWS.Cells(mRow, 6).WrapText = True
         oWS.Cells(mRow, 6).Borders.Color = RGB(0, 0, 0)
         oWS.Cells(mRow, 6) = DateDiff("d", CDate(.TextMatrix(I, 0)), CDate(.TextMatrix(I, 1))) + 1
         oWS.Cells(mRow, 6).Font.Size = 12
         oWS.Cells(mRow, 7).WrapText = True
         oWS.Cells(mRow, 7).Borders.Color = RGB(0, 0, 0)
         oWS.Cells(mRow, 7) = Format(Val(.TextMatrix(I, 6)), "#0")
         mAmount = mAmount + Val(.TextMatrix(I, 6))
         oWS.Cells(mRow, 7).Font.Size = 12
         oWS.Cells(mRow, 8).WrapText = True
         oWS.Cells(mRow, 8).Borders.Color = RGB(0, 0, 0)
         oWS.Cells(mRow, 8) = ""
         oWS.Cells(mRow, 8).Font.Size = 12
      Next
   End With
   
   mRow = mRow + 2
   tmp = "a" & mRow & ":" & "f" & mRow
   oWS.Range(tmp).Merge
   oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
   oWS.Cells(mRow, 1) = "Service Tax @ 18%" '"Service Tax @ 12.36%"
   oWS.Cells(mRow, 1).Font.Bold = True
   oWS.Cells(mRow, 1).Font.Size = 12
   oWS.Cells(mRow, 1).HorizontalAlignment = -4108
   tmp = "g" & mRow & ":" & "g" & mRow
   oWS.Range(tmp).Merge
   oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
   oWS.Cells(mRow, 7) = Format(Val(mAmount) * 18 / 100, "#0") 'Format(Val(mAmount) * 12.36 / 100, "#0")
   oWS.Cells(mRow, 7).Font.Bold = True
   oWS.Cells(mRow, 7).Font.Size = 12
      
   mRow = mRow + 2
   tmp = "a" & mRow & ":" & "f" & mRow
   oWS.Range(tmp).Merge
   oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
   oWS.Cells(mRow, 1) = "Grand Total"
   oWS.Cells(mRow, 1).Font.Bold = True
   oWS.Cells(mRow, 1).Font.Size = 12
   oWS.Cells(mRow, 1).HorizontalAlignment = -4108
   tmp = "g" & mRow & ":" & "g" & mRow
   oWS.Range(tmp).Merge
   oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
   oWS.Cells(mRow, 7) = Format(Val(mAmount) + (Val(mAmount) * 18 / 100), "#0") 'Format(Val(mAmount) + (Val(mAmount) * 12.36 / 100), "#0")
   oWS.Cells(mRow, 7).Font.Bold = True
   oWS.Cells(mRow, 7).Font.Size = 12
      
   mRow = mRow + 3
   tmp = "a" & mRow & ":" & "h" & mRow
   oWS.Range(tmp).Merge
   oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
   oWS.Cells(mRow, 1) = "Rupees:" & NumStrConv(Val(oWS.Cells(mRow - 3, 7)))
   oWS.Cells(mRow, 1).Font.Bold = True
   oWS.Cells(mRow, 1).Font.Size = 12
   
   tmp = "Select MSC.Remarks from Mst_StorageContract MSC Where MSC.ContractID = " & Val(mContractIDRow)
   Call TmpTable
   mRow = mRow + 3
   tmp = "a" & mRow & ":" & "h" & mRow
   oWS.Range(tmp).Merge
   oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
   If TmpRs.RecordCount > 0 Then
      oWS.Cells(mRow, 1) = IIf(IsNull(TmpRs!Remarks), "", TmpRs!Remarks)
   End If
   oWS.Cells(mRow, 1).Font.Bold = True
   oWS.Cells(mRow, 1).Font.Size = 10
   oWS.Cells(mRow, 1).WrapText = True
   oWS.Cells(mRow, 1).RowHeight = 45
End If

oWB.Save
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub

Private Sub CmdPrintFrameWork_Click()
If CmbFramework.Text = "" Then
   MsgBox "Select Management Type!!!"
   CmbFramework.SetFocus
   Exit Sub
End If
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "BillFrameWork.xls")
Call Xls_Detail_Rpt2
Frame3.Visible = False
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
If TxtReverseDate.Value < RsTxn_ReservationInvoice!BillingDate Then
   MsgBox "'Reverse Date' should not less than 'Billing Date' !!"
   TxtReverseDate.SetFocus
   Exit Sub
End If
If TxtReverseRemark.Text = "" Then
   MsgBox "Blank Remark not allowed!!"
   TxtReverseRemark.SetFocus
   Exit Sub
End If

TxtReverseRemark = Replace(TxtReverseRemark, ",", "-")
RsTxn_ReservationInvoice!ReverseDate = TxtReverseDate.Value
RsTxn_ReservationInvoice!ReverseFlag = "Y"
RsTxn_ReservationInvoice!ReverseRemark = TxtReverseRemark.Text
RsTxn_ReservationInvoice!UpdatedBy = Gen_UserLoginID
RsTxn_ReservationInvoice!UpdatedDate = Now
RsTxn_ReservationInvoice.Update

tmp = " Update Txn_TIDetails Set TIAmount=0,ReverseFlag='Y',UpdatedBy = '" & Gen_UserLoginID & "',UpdatedDate = getdate() Where PIID=" & RsTxn_ReservationInvoice!InvoiceNo & " and PIFLag='Reservation'"

Call TmpTable

Frame6.Visible = False
CmdReverse.Enabled = False

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
'   End If
   
   If fso.FileExists(mFileName) = True Then
      If wsPathAdobe <> "" Then
         iRet = Shell(wsPathAdobe & " " & mFileName, vbMaximizedFocus)
         Call Log_Print_History(mSM_Prefix, "Reservation", TxtInvoiceNo, mEntryNo)
      End If
   Else
      MsgBox "Invoice not found. Please contact CMG Department for same!!!"
   End If
   
End If
Exit Sub

Error:
      MsgBox "Please contact Technology department!!!"
End Sub

Private Sub Command1_Click()

End Sub

Private Sub DeleteCmd_Click()

'If LCase(TxtStatus) = "approved" Then
'   MsgBox "This Invoice is " & TxtStatus & ". You can't delete this Invoice!!!"
'   Exit Sub
'End If

If Val(TxtBatchID) <> 0 Then
   MsgBox "Batch generated for this invoice. You can't delete this Invoice!!!"
   Exit Sub
End If

Dim ans As Variant
On Error GoTo msgError
Frame3.Visible = False
ans = MsgBox("Are you Sure you want to Delete this Record ", vbYesNo)

If ans = vbYes Then
   Frame5.Visible = True
   TxtLogRemark.Text = ""
'   tmp = "Delete From Log_ReservationInvoice Where InvoiceNo= " & TxtInvoiceNo.Text & ""
'   Call TmpTable
'
'   tmp = "Delete From Txn_ReservationInvoiceDetail Where InvoiceNo= " & TxtInvoiceNo.Text & ""
'   Call TmpTable
'
'   tmp = "Delete From Txn_ReservationDailyStock Where ContractID=" & mContractIDRow & " And MonthYear='" & RsTxn_ReservationInvoice!yearmonth & "'"
'   Call TmpTable
'
'   RsTxn_ReservationInvoice.Delete
'   RsTxn_ReservationInvoice.Update
'   RsTxn_ReservationInvoice.MoveNext
'   If RsTxn_ReservationInvoice.EOF() Then
'      RsTxn_ReservationInvoice.MoveLast
'   End If
'   Call InData
End If

Exit Sub

msgError:
MsgBox "Child record Present "
RsTxn_ReservationInvoice.CancelUpdate

End Sub

Private Sub EditCmd_Click()

If LCase(TxtStatus) = "approved" Then
   MsgBox "This Invoice is " & TxtStatus & ". You can't edit this Invoice any more!!!"
   Exit Sub
End If

If LCase(Gen_UserRole) <> "power" Then
   If LCase(TxtStatus) = LCase("FREEZE") Then
      MsgBox "This Invoice is " & TxtStatus & ". You can't edit this Invoice any more!!!"
      Exit Sub
   End If
End If

CmdPrint.Enabled = False

Edit_Flg = "E"
Call GButtonLock(Me, False)
CmdReverse.Enabled = False
Frame6.Visible = False
Frame3.Visible = False
If TxtActualAmount.Text = "" Then
   TxtActualAmount.Text = TxtBillingAmount.Text
End If
'TxtClaimDate.Enabled = True
''TxtEmployeeID.Visible = False
''CmbEmployeeID.Visible = True
''CmbEmployeeID.Text = TxtEmployeeID
''CmdEmployeeID.Visible = True

'If TxtNAVCode.Text = "" Then
'   CmbServiceTax.Visible = True
'End If

'If TxtCommodity.Text = "" Then
'   CmbServiceTax.Visible = True
'   TxtServiceTax.Visible = False
'   CmbServiceTax.Text = TxtServiceTax.Text
'Else
'   CmbServiceTax.Visible = False
'   TxtServiceTax.Visible = True
'End If

If LCase(Gen_UserRole) = "admin" Or LCase(Gen_UserRole) = "superadmin" Or LCase(Gen_UserRole) = "fag" Or LCase(Gen_UserRole) = "power" Or LCase(Gen_UserRole) = "headcmg" Or LCase(Gen_UserRole) = "fagadmin" Then
   CmbServiceTax.Visible = True
   TxtServiceTax.Visible = False
   CmbServiceTax.Text = TxtServiceTax.Text
Else
   CmbServiceTax.Visible = False
   TxtServiceTax.Visible = True
End If


'TxtForwardTo.Visible = False
'CmbForwardedTo.Visible = True
'CmbForwardedTo.Text = TxtForwardTo
TxtStatus.Visible = False
CmbStatus.Visible = True
CmbStatus.Text = TxtStatus
'TxtFAGRemarks.Locked = False
TxtRemarks.Locked = False
TxtActualAmount.Locked = False
TxtDiscount.Locked = False
TxtPremium.Locked = False
TxtDiscount.SetFocus

CmdViewInvoice.Enabled = False
End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub FirstCmd_Click()
RsTxn_ReservationInvoice.MoveFirst
Call Indata
'TmpRs.MoveLast
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
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
Call Grid_Head3

'SFrom.Value = CDate(From_GenDate) - 900
'STo.Value = Date 'To_GenDate
CmbSLocation.Text = ""
'CmbSEmployee.Text = ""
'CmbSStatus.Text = "Pending"
'CmbSForwardTo.Text = ""
ChkBatchNo.Value = 0
TxtSClaimID = ""
Label22.Caption = ""
mSStatus = "P"


TxtFromDate.Value = Date
TxtToDate.Value = Date
TxtFromDate.Value = GetLastDateOfMonth(TxtFromDate.Value)
TxtToDate.Value = GetLastDateOfMonth(TxtToDate.Value)
Call TxtFromDate_LostFocus
Call TxtToDate_LostFocus


If Gen_WcLocation <> "" Then
   Call TableTxn_ReservationInvoice(" Where " & Gen_WcLocation & "  Order by InvoiceNo")
Else
   Call TableTxn_ReservationInvoice(" Order by InvoiceNo")
End If
If RsTxn_ReservationInvoice.RecordCount = 0 Then
   If UserType = "Auditor" Then
      Call GButtonLock(Me, False)
      SaveCmd.Enabled = False
      SearchCmd.Enabled = False
      Exit Sub
   End If
   Call GButtonLockAD(Me)
   Exit Sub
End If

'For I = 1 To 12
'    CmbMonth.AddItem MonthName(I), I - 1
'Next

Call GButtonLock(Me, True)

Call Indata
 
End Sub

Private Sub Form_Resize()
If Me.WindowState = vbMaximized Then
   SetScrollBars
End If
End Sub


Private Sub LastCmd_Click()
RsTxn_ReservationInvoice.MoveLast
Call Indata
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False
End Sub

Private Sub MSHFlexGrid1_Click()
Dim x As Variant
Dim mCaId As Variant

x = MSHFlexGrid1.RowSel
If MSHFlexGrid1.TextMatrix(x, 0) = "" Then Exit Sub
mCaId = MSHFlexGrid1.TextMatrix(x, 0)
RsTxn_ReservationInvoice.MoveFirst

'If RsTxn_ReservationInvoice.RecordCount = 0 Then
'   Call AddCmd_Click
'   Exit Sub
'End If
RsTxn_ReservationInvoice.MoveFirst
RsTxn_ReservationInvoice.Find "InvoiceNo = " & Val(mCaId) & ""
If RsTxn_ReservationInvoice.EOF Then
   RsTxn_ReservationInvoice.MoveLast
End If
Call Indata
Frame0.Visible = False

End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsTxn_ReservationInvoice.MoveNext
If RsTxn_ReservationInvoice.EOF Then
   RsTxn_ReservationInvoice.MoveLast
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
RsTxn_ReservationInvoice.MovePrevious
If RsTxn_ReservationInvoice.BOF Then
   RsTxn_ReservationInvoice.MoveFirst
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

Public Sub Indata()


Call GButtonLock(Me, True)

CmdReverse.Enabled = False
Frame6.Visible = False
Frame3.Visible = False
TxtServiceTax.Visible = True
'CmbForwardedTo.Visible = False
CmbStatus.Visible = False
TxtStatus.Visible = True
'TxtForwardTo.Visible = True
mStatus = ""
'mForwardToID = ""
CmbServiceTax.Visible = False
CmdViewInvoice.Enabled = True

tmp = " Select MST.Description,MST.NAVCode,ML.LocationName,TRI.Flag,TRI.InvoiceNo,TRI.Discount,TRI.Premium,TRI.ActualAmount,MSC.ContractName,"
tmp = tmp & " TRI.BillingAmount ,TRI.ServiceTaxNAVCode,TRI.BillingDate,TRI.BatchID , TRI.YearMonth, TRI.Remarks, "
tmp = tmp & " MC.Commodity , MCL.ClientNAme,TRI.CreatedDate,TRI.CreatedBy,TRI.ContractID,ML.SM_Prefix,ML.LocationID,MCL.ClientID"
tmp = tmp & " from Txn_ReservationInvoice TRI"
tmp = tmp & " inner join Mst_StorageContract MSC on MSC.ContractID=TRI.ContractID"
tmp = tmp & " Left join MSt_Commodity MC on MC.CommodityID=MSC.CommodityID"
tmp = tmp & " inner join Mst_Client MCL on MCL.ClientID=MSC.ClientID"
tmp = tmp & " inner join mst_Location ML on ML.LocationID=TRI.LocationID"
'tmp = tmp & " inner join Mst_Employee ME on ME.Employeeid=TRI.ForwardToID"
tmp = tmp & " left join MSt_ServiceTax MST on MST.SErviceTAxID=MC.SErviceTAxID"
tmp = tmp & " Where InvoiceNo=" & RsTxn_ReservationInvoice!InvoiceNo & "  "


Call TmpTable


TxtActualAmount.Locked = True
TxtPremium.Locked = True
TxtDiscount.Locked = True
mSM_Prefix = TmpRs!SM_Prefix
TxtRemarks.Locked = True
TxtInvoiceNo = TmpRs!InvoiceNo
mContractIDRow = TmpRs!ContractID
TxtInvoiceDate.Text = Format(TmpRs!BillingDate, "dd-mm-yyyy")
TxtInvoiceMonth = Format(TmpRs!YearMonth, "mmm - yyyy")
TxtLocation = TmpRs!Locationname
TxtContractName = TmpRs!ContractName
TxtClient = TmpRs!ClientName
TxtCommodity = IIf(IsNull(TmpRs!Commodity), "", TmpRs!Commodity)

'If TxtCommodity.Text <> "" Then
'  TxtNAVCode.Text = IIf(IsNull(TmpRs!NAVCode), "", TmpRs!NAVCode)
'   TxtServiceTax.Text = IIf(IsNull(TmpRs!Description), "", TmpRs!Description)
'Else
   TxtNAVCode.Text = IIf(IsNull(TmpRs!ServiceTaxNAVCode), "", TmpRs!ServiceTaxNAVCode)
   TxtServiceTax.Text = GetServiceTaxDescFromNAVCode(TxtNAVCode.Text)
'End If

mStatus = TmpRs!Flag

If LCase(mStatus) = "p" Then
   TxtStatus.Text = "Pending"
ElseIf LCase(mStatus) = "a" Then
   TxtStatus.Text = "Approved"
ElseIf LCase(mStatus) = "f" Then
   TxtStatus.Text = "Freeze"
End If

TxtActualAmount = IIf(IsNull(TmpRs!ActualAmount), "", TmpRs!ActualAmount)
TxtPremium = IIf(IsNull(TmpRs!Premium), 0, TmpRs!Premium)
TxtDiscount = IIf(IsNull(TmpRs!Discount), 0, TmpRs!Discount)

TxtBillingAmount = IIf(IsNull(TmpRs!BillingAmount), 0, TmpRs!BillingAmount)

TxtBatchID = IIf(IsNull(TmpRs!BatchID), "", TmpRs!BatchID)
TxtRemarks.Text = IIf(IsNull(TmpRs!Remarks), "", TmpRs!Remarks)
TxtCreated = IIf(IsNull(TmpRs!CreatedBy), "", TmpRs!CreatedBy) & " :- " & IIf(IsNull(TmpRs!CreatedDate), "", TmpRs!CreatedDate)

''tmp = " Select TDI.TINO,TDI.TIAmount,TDI.TISTAmount,TDI.TIDate,TDI.TIID,TDI.EntryNo From Txn_TIDetails TDI"
''tmp = tmp & " Where PIID=" & RsTxn_ReservationInvoice!InvoiceNo & " and TDI.PIFLag='Reservation'  "
''
''Call TmpTable
''
''If TmpRs.RecordCount > 0 Then
''   TxtTINo = IIf(IsNull(TmpRs!TINO), "", TmpRs!TINO)
''   TxtTIDate = IIf(IsNull(TmpRs!TIDate), "", Format(TmpRs!TIDate, "dd-mm-yyyy"))
''   TxtTIAmount = IIf(IsNull(TmpRs!TIAmount), "", TmpRs!TIAmount)
''   mEntryNo = IIf(IsNull(TmpRs!EntryNo), "", TmpRs!EntryNo)
''   If LCase(Gen_UserRole) = "fag" Or LCase(Gen_UserRole) = "power" Then
''      If IIf(IsNull(RsTxn_ReservationInvoice!ReverseFlag), "N", RsTxn_ReservationInvoice!ReverseFlag) <> "Y" Then
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

ans = GetTIDetails(RsTxn_ReservationInvoice!InvoiceNo, "Reservation")
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
   If IIf(IsNull(RsTxn_ReservationInvoice!ReverseFlag), "N", RsTxn_ReservationInvoice!ReverseFlag) <> "Y" Then
      If CheckTIInCR(Val(ans(0))) = False Then
         CmdReverse.Enabled = True
      End If
   End If
End If

mLocationID = IIf(IsNull(TmpRs!LocationID), 0, TmpRs!LocationID)
mClientID = IIf(IsNull(TmpRs!ClientID), 0, TmpRs!ClientID)

Call InDataReservationInvoiceLogHistory
Call InDataReservationInvoiceDetail
Call InDataStockStatement

TxtAdHocAmt = GetADhocInvoiceAmount(mLocationID, mClientID)


CmdPrint.Enabled = True
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

Private Sub TxtSInvoiceNo_LostFocus()

If TxtSInvoiceNo = "" Then Exit Sub
If IsNumeric(TxtSInvoiceNo) = False Or Val(TxtSInvoiceNo) < 0 Then
   MsgBox "Invalid numeric format !!!!"
   TxtSInvoiceNo.SetFocus
   Exit Sub
End If

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

Private Sub SaveCmd_Click()

TxtActualAmount.Text = Format(Val(TxtBillingAmount.Text) - Val(TxtDiscount.Text) + Val(TxtPremium.Text), "#####0.00")
If Val(TxtActualAmount.Text) < 0 Then
   MsgBox "Actual amount should not negative!!"
   TxtPremium.SetFocus
   Exit Sub
End If

''If TxtActualAmount.Text = "" Then
''   MsgBox "Blank Actual amount not allowed!!!"
''   TxtActualAmount.SetFocus
''   Exit Sub
''End If

If TxtRemarks.Text = "" Then
   MsgBox "Blank Remarks not allowed!!!"
   TxtRemarks.SetFocus
   Exit Sub
End If

If CmbStatus.Text = "" Then
   MsgBox "Blank Status not allowed!!!"
   CmbStatus.SetFocus
   Exit Sub
End If

If CmbServiceTax.Visible = True Then
   If CmbServiceTax.Text = "" Then
      MsgBox "Blank Service Tax not allowed!!!"
      CmbServiceTax.SetFocus
      Exit Sub
   End If
End If

If Val(TxtActualAmount.Text) > 0 Then
   If LCase(mStatus) = "a" Then
      If GetStateWiseNAVCode(mClientID, mLocationID, "R") = "" Then
         MsgBox "Nav code mapping missing in State Wise GST mapping screen!!!"
         CmbStatus.SetFocus
         Exit Sub
      End If
   End If
End If

RsTxn_ReservationInvoice.MoveFirst
RsTxn_ReservationInvoice.Find " InvoiceNo= " & TxtInvoiceNo.Text & ""

RsTxn_ReservationInvoice!Flag = mStatus
'RsTxn_ReservationInvoice!ForwardToID = mForwardToID

TxtRemarks = Replace(TxtRemarks, ",", "-")

RsTxn_ReservationInvoice!Remarks = Trim(TxtRemarks.Text)
RsTxn_ReservationInvoice!ActualAmount = Val(TxtActualAmount.Text)
RsTxn_ReservationInvoice!Discount = Val(TxtDiscount.Text)
RsTxn_ReservationInvoice!Premium = Val(TxtPremium.Text)
RsTxn_ReservationInvoice!ServiceTaxNAVCode = TxtNAVCode.Text

If IsNull(RsTxn_ReservationInvoice!CreatedBy) = True Or RsTxn_ReservationInvoice!CreatedBy = "" Then
   RsTxn_ReservationInvoice!CreatedBy = Gen_UserLoginID
   RsTxn_ReservationInvoice!CreatedDate = Now
Else
   RsTxn_ReservationInvoice!UpdatedBy = Gen_UserLoginID
   RsTxn_ReservationInvoice!UpdatedDate = Now
End If

RsTxn_ReservationInvoice.Update

'Call Gen_Email
If Val(TxtBillingAmount.Text) <> Val(TxtActualAmount.Text) Then
   Call UpdateInvMonGSL
End If

RsTxn_ReservationInvoice.Requery
RsTxn_ReservationInvoice.MoveFirst
RsTxn_ReservationInvoice.Find "InvoiceNo = " & TxtInvoiceNo.Text & ""

Call GButtonLock(Me, True)

Call Indata

End Sub

Private Sub SearchCmd_Click()
Frame3.Visible = False
If SearchCmd.Caption = "Cancel" Then
   If RsTxn_ReservationInvoice.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   RsTxn_ReservationInvoice.MoveFirst
 
   RsTxn_ReservationInvoice.Find "InvoiceNo = '" & TxtInvoiceNo.Text & "'"
   If RsTxn_ReservationInvoice.EOF Then
      RsTxn_ReservationInvoice.MoveLast
   End If
   Call Indata
   Exit Sub
End If

Frame0.Height = Frame4.Height - 200 'Me.Height - 500
Frame0.Left = Frame1.Left - 100
Frame0.Top = Frame1.Top
Frame0.Width = Frame4.Width - 200
Frame0.Visible = True

MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 1400

RsTxn_ReservationInvoice.MoveFirst

''===========comment on 7 july 2012 for not clearing search grid=============
'Call Grid_Head3
'
''SFrom.Value = CDate(From_GenDate) - 900
''STo.Value = Date 'To_GenDate
'CmbSLocation.Text = ""
''CmbSEmployee.Text = ""
''CmbSStatus.Text = "Pending"
''CmbSForwardTo.Text = ""
'ChkBatchNo.Value = 0
'TxtSClaimID = ""
'Label22.Caption = ""
'mSStatus = "P"
''===========comment on 7 july 2012 for not clearing search grid=============

'tmp = "Select EmployeeID,EmployeeName +'~'+ EmployeeNo as Employee,DesignationID,EmployeeEmailID From Mst_Employee Where EmployeeNo = '" & Gen_UserLoginID & "'"
'Call TmpTable
'
'If TmpRs.RecordCount = 1 Then
'   CmbSForwardTo.Text = TmpRs!Employee
'   mSForwardToID = TmpRs!EmployeeID
'Else
'   CmbSForwardTo.Text = ""
'   mSForwardToID = Null
'End If
'
End Sub
Private Sub Grid_Head3()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.FixedRows = 1

MSHFlexGrid1.Cols = 32 '30 '25
MSHFlexGrid1.Font.Size = 10

MSHFlexGrid1.TextMatrix(0, 0) = "Invoice No"
MSHFlexGrid1.TextMatrix(0, 1) = "Contract Name"
MSHFlexGrid1.TextMatrix(0, 2) = "Location Name"
MSHFlexGrid1.TextMatrix(0, 3) = "Commodity"
MSHFlexGrid1.TextMatrix(0, 4) = "Year Month"
MSHFlexGrid1.TextMatrix(0, 5) = "Billing Date"
MSHFlexGrid1.TextMatrix(0, 6) = "Billing Amount"
MSHFlexGrid1.TextMatrix(0, 7) = "Client Name"
MSHFlexGrid1.TextMatrix(0, 8) = "Actual Amount"
MSHFlexGrid1.TextMatrix(0, 9) = "GST Group Code" '"Service Tax" '"TI No"
MSHFlexGrid1.TextMatrix(0, 10) = "TI Date"
MSHFlexGrid1.TextMatrix(0, 11) = "TI Amount"
MSHFlexGrid1.TextMatrix(0, 12) = "BatchID"
MSHFlexGrid1.TextMatrix(0, 13) = "Status"
MSHFlexGrid1.TextMatrix(0, 14) = "Remarks"

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

''--- Ullhas Patil 27 Sep 2017 GST Columns
MSHFlexGrid1.TextMatrix(0, 27) = "State India Name":                MSHFlexGrid1.ColWidth(27) = 2000
MSHFlexGrid1.TextMatrix(0, 28) = "NCML GST No":                     MSHFlexGrid1.ColWidth(28) = 2000
MSHFlexGrid1.TextMatrix(0, 29) = "Client GST No":                   MSHFlexGrid1.ColWidth(29) = 2000
MSHFlexGrid1.TextMatrix(0, 30) = "Client GST NAVCode":              MSHFlexGrid1.ColWidth(30) = 2000
MSHFlexGrid1.TextMatrix(0, 31) = "Tax Applied (Inter/Intra)":       MSHFlexGrid1.ColWidth(31) = 2000
''--------------

'MSHFlexGrid1.TextMatrix(0, 15) = "Approval Remarks"
'MSHFlexGrid1.TextMatrix(0, 16) = "Approval Date"
'MSHFlexGrid1.TextMatrix(0, 17) = "Forwarded To"
 
MSHFlexGrid1.ColWidth(0) = 720 '500
MSHFlexGrid1.ColWidth(1) = 2900
MSHFlexGrid1.ColWidth(2) = 1800
MSHFlexGrid1.ColWidth(3) = 1150
MSHFlexGrid1.ColWidth(4) = 1100
MSHFlexGrid1.ColWidth(5) = 1100

MSHFlexGrid1.ColWidth(6) = 900 '500
MSHFlexGrid1.ColWidth(7) = 3200
MSHFlexGrid1.ColWidth(8) = 1200
MSHFlexGrid1.ColWidth(9) = 1200
MSHFlexGrid1.ColWidth(10) = 0
MSHFlexGrid1.ColWidth(11) = 0
MSHFlexGrid1.ColWidth(12) = 1200
MSHFlexGrid1.ColWidth(13) = 1000
MSHFlexGrid1.ColWidth(14) = 2500
MSHFlexGrid1.ColWidth(15) = 1600
MSHFlexGrid1.ColWidth(16) = 1600
MSHFlexGrid1.ColWidth(17) = 1600
MSHFlexGrid1.ColWidth(18) = 1600
MSHFlexGrid1.ColWidth(19) = 1600
MSHFlexGrid1.ColWidth(20) = 1600
'MSHFlexGrid1.ColWidth(15) = 2500
'MSHFlexGrid1.ColWidth(16) = 1100
'MSHFlexGrid1.ColWidth(17) = 2200

MSHFlexGrid1.RowHeight(0) = 600
MSHFlexGrid1.WordWrap = True
End Sub

Private Sub TxtActualAmount_LostFocus()
If TxtActualAmount = "" Then Exit Sub
If IsNumeric(TxtActualAmount) = False Or Val(TxtActualAmount) < 0 Then
   MsgBox "Invalid numeric format!!!"
   TxtActualAmount.SetFocus
   Exit Sub
End If
End Sub
Private Sub InDataReservationInvoiceLogHistory()

Call Grid_Head2
 
tmp = " select INvoiceNo,LRI.UpdatedDate,LRI.UpdatedBy,LRI.Remarks " & _
      " From Log_ReservationInvoice LRI " & _
      " Where InvoiceNo= " & TxtInvoiceNo.Text & " "
Call TmpTable
If TmpRs.RecordCount > 0 Then
   MSHFlexGrid3.Rows = TmpRs.RecordCount + 1
   For iu = 1 To TmpRs.RecordCount

       MSHFlexGrid3.TextMatrix(iu, 0) = IIf(IsNull(TmpRs!InvoiceNo), "", TmpRs!InvoiceNo)
       MSHFlexGrid3.TextMatrix(iu, 1) = IIf(IsNull(TmpRs!UpdatedBy), "", TmpRs!UpdatedBy)
       MSHFlexGrid3.TextMatrix(iu, 2) = IIf(IsNull(TmpRs!UpdatedDate), TmpRs!UpdatedDate & ":- " & Format(TmpRs!UpdatedDate, "dd-mmm-yyyy"), TmpRs!UpdatedDate)
       MSHFlexGrid3.TextMatrix(iu, 3) = IIf(IsNull(TmpRs!Remarks), "", TmpRs!Remarks)

       TmpRs.MoveNext
   Next
End If
End Sub


Private Sub Grid_Head4()
MSHFlexGrid2.Clear
MSHFlexGrid2.Rows = 3
MSHFlexGrid2.FixedRows = 1

MSHFlexGrid2.Cols = 8
MSHFlexGrid2.Font.Size = 10

MSHFlexGrid2.TextMatrix(0, 0) = "From Date"
MSHFlexGrid2.TextMatrix(0, 1) = "To Date"
MSHFlexGrid2.TextMatrix(0, 2) = "Billing Cycle"
MSHFlexGrid2.TextMatrix(0, 3) = "Billing Unit"
MSHFlexGrid2.TextMatrix(0, 4) = "Peak Value"
MSHFlexGrid2.TextMatrix(0, 5) = "Rate"
MSHFlexGrid2.TextMatrix(0, 6) = "Amount"
MSHFlexGrid2.TextMatrix(0, 7) = "Contract ID"

MSHFlexGrid2.ColWidth(0) = 1400 '500
MSHFlexGrid2.ColWidth(1) = 1400
MSHFlexGrid2.ColWidth(2) = 2200
MSHFlexGrid2.ColWidth(3) = 1400
MSHFlexGrid2.ColWidth(4) = 1200
MSHFlexGrid2.ColWidth(5) = 1200
MSHFlexGrid2.ColWidth(6) = 1200
MSHFlexGrid2.ColWidth(7) = 0

MSHFlexGrid2.RowHeight(0) = 600
MSHFlexGrid2.WordWrap = True
End Sub


Private Sub InDataReservationInvoiceDetail()

Call Grid_Head4
 
Call TableTxn_ReservationInvoiceDetail(" Where InvoiceNo=" & Val(TxtInvoiceNo.Text))
If RsTxn_ReservationInvoiceDetail.RecordCount > 0 Then
   MSHFlexGrid2.Rows = RsTxn_ReservationInvoiceDetail.RecordCount + 1
   For iu = 1 To RsTxn_ReservationInvoiceDetail.RecordCount
       MSHFlexGrid2.TextMatrix(iu, 0) = IIf(IsNull(RsTxn_ReservationInvoiceDetail!FromDate), "", Format(RsTxn_ReservationInvoiceDetail!FromDate, "dd-MMM-yyyy"))
       MSHFlexGrid2.TextMatrix(iu, 1) = IIf(IsNull(RsTxn_ReservationInvoiceDetail!ToDate), "", Format(RsTxn_ReservationInvoiceDetail!ToDate, "dd-MMM-yyyy"))
       MSHFlexGrid2.TextMatrix(iu, 2) = IIf(IsNull(RsTxn_ReservationInvoiceDetail!BillingCycleID), "", GetBillingCycle(RsTxn_ReservationInvoiceDetail!BillingCycleID, "N"))
       MSHFlexGrid2.TextMatrix(iu, 3) = IIf(IsNull(RsTxn_ReservationInvoiceDetail!Flag), "", RsTxn_ReservationInvoiceDetail!Flag)
       If MSHFlexGrid2.TextMatrix(iu, 3) = "Q" Then
          MSHFlexGrid2.TextMatrix(iu, 3) = "Wt in MT"
       ElseIf MSHFlexGrid2.TextMatrix(iu, 3) = "B" Then
          MSHFlexGrid2.TextMatrix(iu, 3) = "No of Bags"
       ElseIf MSHFlexGrid2.TextMatrix(iu, 3) = "F" Then
          MSHFlexGrid2.TextMatrix(iu, 3) = "Fixed"
       End If
       MSHFlexGrid2.TextMatrix(iu, 4) = IIf(IsNull(RsTxn_ReservationInvoiceDetail!PeakValue), "", Format(RsTxn_ReservationInvoiceDetail!PeakValue, "#####0.000"))
       MSHFlexGrid2.TextMatrix(iu, 5) = IIf(IsNull(RsTxn_ReservationInvoiceDetail!Rate), "", Format(RsTxn_ReservationInvoiceDetail!Rate, "#####0.00"))
       MSHFlexGrid2.TextMatrix(iu, 6) = IIf(IsNull(RsTxn_ReservationInvoiceDetail!Amount), "", Format(RsTxn_ReservationInvoiceDetail!Amount, "#####0.00"))
       MSHFlexGrid2.TextMatrix(iu, 7) = IIf(IsNull(RsTxn_ReservationInvoiceDetail!ContractID), "", RsTxn_ReservationInvoiceDetail!ContractID)
       
       RsTxn_ReservationInvoiceDetail.MoveNext
   Next
End If
End Sub

Private Sub Grid_Head5()
MSHFlexGrid4.Clear
MSHFlexGrid4.Rows = 2
MSHFlexGrid4.FixedRows = 1

MSHFlexGrid4.Cols = 10
MSHFlexGrid4.Font.Size = 10

MSHFlexGrid4.TextMatrix(0, 0) = "CxTF No"
MSHFlexGrid4.TextMatrix(0, 1) = "Date"
MSHFlexGrid4.TextMatrix(0, 2) = "Opening Balance (Bags)"
MSHFlexGrid4.TextMatrix(0, 3) = "Opening Balance (Qty)"
MSHFlexGrid4.TextMatrix(0, 4) = "Deposit Bags"
MSHFlexGrid4.TextMatrix(0, 5) = "Deposit Qty"
MSHFlexGrid4.TextMatrix(0, 6) = "Withdrawal Bags"
MSHFlexGrid4.TextMatrix(0, 7) = "Withdrawal Qty"
MSHFlexGrid4.TextMatrix(0, 8) = "Closing Balance (Bags)"
MSHFlexGrid4.TextMatrix(0, 9) = "Closing Balance (Qty)"

MSHFlexGrid4.RowHeight(0) = 900
MSHFlexGrid4.ColWidth(1) = 1400
MSHFlexGrid4.WordWrap = True
End Sub

Private Sub InDataStockStatement()

Call Grid_Head5
 
Call TableTxn_ReservationDailyStock(" Where ContractID=" & mContractIDRow & " And MonthYear='" & Format(RsTxn_ReservationInvoice!YearMonth, Gen_DatFormat) & "' Order By TxnDate")
If RsTxn_ReservationDailyStock.RecordCount > 0 Then
   MSHFlexGrid4.Rows = RsTxn_ReservationDailyStock.RecordCount + 1
   For iu = 1 To RsTxn_ReservationDailyStock.RecordCount
       MSHFlexGrid4.TextMatrix(iu, 0) = IIf(IsNull(RsTxn_ReservationDailyStock!CxTFNO), "", RsTxn_ReservationDailyStock!CxTFNO)
       MSHFlexGrid4.TextMatrix(iu, 1) = IIf(IsNull(RsTxn_ReservationDailyStock!TxnDate), "", Format(RsTxn_ReservationDailyStock!TxnDate, "dd-MMM-yyyy"))
       MSHFlexGrid4.TextMatrix(iu, 2) = RsTxn_ReservationDailyStock!OpnBalBags
       MSHFlexGrid4.TextMatrix(iu, 3) = RsTxn_ReservationDailyStock!OpnBalQty
       MSHFlexGrid4.TextMatrix(iu, 4) = RsTxn_ReservationDailyStock!DepoBags
       MSHFlexGrid4.TextMatrix(iu, 5) = RsTxn_ReservationDailyStock!DepoQty
       MSHFlexGrid4.TextMatrix(iu, 6) = RsTxn_ReservationDailyStock!WidraBags
       MSHFlexGrid4.TextMatrix(iu, 7) = RsTxn_ReservationDailyStock!WidraQty
       MSHFlexGrid4.TextMatrix(iu, 8) = RsTxn_ReservationDailyStock!BalanceBags
       MSHFlexGrid4.TextMatrix(iu, 9) = RsTxn_ReservationDailyStock!BalanceWeight
       
       RsTxn_ReservationDailyStock.MoveNext
   Next
End If
End Sub

Private Sub CmdStockExcel_Click()
If MSHFlexGrid4.Rows < 3 Then Exit Sub
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Stock Statement_" & TxtInvoiceNo & "_" & TxtInvoiceMonth.Text & ".xls")
Call Xls_Detail_Rpt1
End Sub

Public Sub Xls_Detail_Rpt1()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Integer
Dim mCol As Integer
MsgBox ("Wait till next message")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Stock Statement_" & TxtInvoiceNo & "_" & TxtInvoiceMonth.Text & ".xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid4.Rows - 1
    mRow = I + 1 ' + 6
    For C = 0 To MSHFlexGrid4.Cols - 1
        If C = 0 Then
           oWS.Cells(mRow, C + 1) = "'" & MSHFlexGrid4.TextMatrix(I, C)
        Else
        
           If InStr(MSHFlexGrid4.TextMatrix(I, C), "/") > 0 Then
              If IsDate(MSHFlexGrid4.TextMatrix(I, C)) Then
                 oWS.Cells(mRow, C + 1) = Format(MSHFlexGrid4.TextMatrix(I, C), "MM/dd/yyyy")
              Else
                 oWS.Cells(mRow, C + 1) = MSHFlexGrid4.TextMatrix(I, C)
              End If
           Else
              oWS.Cells(mRow, C + 1) = MSHFlexGrid4.TextMatrix(I, C)
           End If
        End If
    Next
Next
oWB.Save
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub


Private Sub TxtDiscount_LostFocus()
If TxtDiscount = "" Then Exit Sub
If IsNumeric(TxtDiscount) = False Or Val(TxtDiscount) < 0 Then
   MsgBox "Invalid numeric format !!!!"
   TxtDiscount.SetFocus
   Exit Sub
End If
TxtActualAmount.Text = Format(Val(TxtBillingAmount.Text) - Val(TxtDiscount.Text) + Val(TxtPremium.Text), "#####0.00")
If Val(TxtActualAmount.Text) < 0 Then
   MsgBox "Actual amount should not negative!!"
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
TxtActualAmount.Text = Format(Val(TxtBillingAmount.Text) - Val(TxtDiscount.Text) + Val(TxtPremium.Text), "#####0.00")
If Val(TxtActualAmount.Text) < 0 Then
   MsgBox "Actual amount should not negative!!"
   TxtPremium.SetFocus
   Exit Sub
End If

End Sub

Private Sub UpdateInvMonGSL()
tmp = "Select Mst_GSL_ID,PINO,InvoiceDate,ContractID,isnull(BalanceWeight,0)BalanceWeight,InvoiceAmount,MonG_UID from Txn_InvMonGSLDetail  Where PINO=" & Val(TxtInvoiceNo) & " And ContractID=" & mContractIDRow & " And YearMonth='" & Format(RsTxn_ReservationInvoice!YearMonth, Gen_DatFormat) & "'"    '' 'Where YearMonth ='" & mFromDate & "' and isnull(ContractID,0)>0 And PINO is not null"
Call Tmp2Table

mTotQty = 0

If TmpRs2.RecordCount > 0 Then
   For Ln = 1 To TmpRs2.RecordCount
       If TmpRs2!BalanceWeight > 0 Then
          mTotQty = mTotQty + TmpRs2!BalanceWeight
       End If
       TmpRs2.MoveNext
   Next Ln
   If mTotQty > 0 Then
      tmp = " Update Txn_InvMonGSLDetail Set InvoiceAmount=" & Val(TxtActualAmount.Text) & " * BalanceWeight/" & mTotQty & "  Where PINO=" & Val(TxtInvoiceNo) & " And ContractID=" & Val(mContractIDRow) & " And YearMonth ='" & Format(RsTxn_ReservationInvoice!YearMonth, Gen_DatFormat) & "' And BalanceWeight>0"
      Call TmpTable
   End If
End If

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

Private Function GetADhocInvoiceAmount(mLocID_ As Variant, mClientID_ As Variant)
Dim Retval As Variant
Retval = 0

tmp = "Select LocationID,ClientID,Sum(Amount) As 'Amount' From Txn_TempInvoice "
tmp = tmp & " Where IETypeID = 4 And LocationID = " & mLocID_ & " And ClientID = " & mClientID_ & " And Month(InvoiceDate) = " & Month(CDate(TxtInvoiceMonth)) & "  And Year(InvoiceDate) = " & Year(CDate(TxtInvoiceMonth)) & ""
tmp = tmp & " Group By LocationID,ClientID "

Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   Retval = TmpRs1!Amount
End If

GetADhocInvoiceAmount = Retval

End Function
