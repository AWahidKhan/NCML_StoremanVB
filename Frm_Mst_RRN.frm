VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmMst_RRN 
   Caption         =   "RRN Master"
   ClientHeight    =   4020
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4995
   MDIChild        =   -1  'True
   ScaleHeight     =   4020
   ScaleWidth      =   4995
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   4005
      Left            =   120
      TabIndex        =   0
      Top             =   9480
      Visible         =   0   'False
      Width           =   15000
      Begin VB.TextBox TxtSearchISIN 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4440
         TabIndex        =   39
         Top             =   420
         Width           =   2205
      End
      Begin VB.ComboBox CmbSCommodity 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   42
         Top             =   1155
         Width           =   5745
      End
      Begin VB.ComboBox CmbSCmp 
         Height          =   315
         Left            =   9630
         Sorted          =   -1  'True
         TabIndex        =   41
         Top             =   420
         Width           =   3810
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "Search"
         Height          =   350
         Left            =   12750
         TabIndex        =   43
         Top             =   1155
         Width           =   1000
      End
      Begin VB.ComboBox CmbSLocation 
         Height          =   315
         Left            =   6675
         Sorted          =   -1  'True
         TabIndex        =   40
         Top             =   420
         Width           =   2910
      End
      Begin VB.TextBox TxtSearchRRN 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2880
         TabIndex        =   38
         Top             =   420
         Width           =   1530
      End
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   350
         Left            =   13755
         TabIndex        =   45
         Top             =   1155
         Width           =   1000
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1455
         Left            =   75
         TabIndex        =   10
         Top             =   1560
         Width           =   4665
         _ExtentX        =   8229
         _ExtentY        =   2566
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSComCtl2.DTPicker STo 
         Height          =   300
         Left            =   1515
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   420
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
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
         Format          =   48824321
         CurrentDate     =   36494
      End
      Begin MSComCtl2.DTPicker SFrom 
         Height          =   300
         Left            =   120
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   420
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
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
         Format          =   48824321
         CurrentDate     =   36494
      End
      Begin VB.Label Label45 
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
         Height          =   255
         Left            =   2032
         TabIndex        =   110
         Top             =   180
         Width           =   300
      End
      Begin VB.Label Label21 
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
         Height          =   255
         Left            =   525
         TabIndex        =   109
         Top             =   180
         Width           =   525
      End
      Begin VB.Label LblISINNo 
         Caption         =   "CMSE Lot No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4890
         TabIndex        =   68
         Top             =   180
         Width           =   1305
      End
      Begin VB.Label Label5 
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
         Left            =   2407
         TabIndex        =   67
         Top             =   900
         Width           =   1170
      End
      Begin VB.Label Label3 
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
         Left            =   11385
         TabIndex        =   66
         Top             =   180
         Width           =   435
      End
      Begin VB.Label Label4 
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
         Height          =   300
         Left            =   7770
         TabIndex        =   65
         Top             =   180
         Width           =   675
      End
      Begin VB.Label LblSearchRRN 
         Caption         =   "RRN No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3225
         TabIndex        =   44
         Top             =   180
         Width           =   840
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4575
      Left            =   135
      TabIndex        =   76
      Top             =   4080
      Width           =   14805
      _ExtentX        =   26114
      _ExtentY        =   8070
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "CommTrack Lot Detail"
      TabPicture(0)   =   "Frm_Mst_RRN.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "MSHFlexGrid4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Pic_TTD"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "CA - CITF Detail"
      TabPicture(1)   =   "Frm_Mst_RRN.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(1)=   "Frame3"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Log Detail"
      TabPicture(2)   =   "Frm_Mst_RRN.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label9"
      Tab(2).Control(1)=   "Label6"
      Tab(2).Control(2)=   "Label13"
      Tab(2).Control(3)=   "TxtUpdated"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "TxtCreated"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "TxtOutOfSystemUpdated"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).ControlCount=   6
      Begin VB.TextBox TxtOutOfSystemUpdated 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   325
         Left            =   -73635
         Locked          =   -1  'True
         TabIndex        =   99
         TabStop         =   0   'False
         Top             =   1740
         Width           =   6120
      End
      Begin VB.TextBox TxtCreated 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   325
         Left            =   -73635
         Locked          =   -1  'True
         TabIndex        =   96
         TabStop         =   0   'False
         Top             =   600
         Width           =   6120
      End
      Begin VB.TextBox TxtUpdated 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   325
         Left            =   -73635
         Locked          =   -1  'True
         TabIndex        =   95
         TabStop         =   0   'False
         Top             =   1185
         Width           =   6120
      End
      Begin VB.PictureBox Pic_TTD 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1575
         Left            =   120
         ScaleHeight     =   1545
         ScaleWidth      =   14505
         TabIndex        =   83
         Top             =   360
         Width           =   14535
         Begin VB.CommandButton CmdDeleteList 
            Caption         =   "Remove from List"
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
            Height          =   375
            Left            =   11760
            TabIndex        =   108
            Top             =   960
            Width           =   1950
         End
         Begin VB.TextBox TxtLotCommodity 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   375
            Left            =   2760
            Locked          =   -1  'True
            TabIndex        =   107
            TabStop         =   0   'False
            Top             =   420
            Width           =   3120
         End
         Begin VB.TextBox TxtLotLocation 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   325
            Left            =   795
            Locked          =   -1  'True
            TabIndex        =   106
            TabStop         =   0   'False
            Top             =   960
            Width           =   5085
         End
         Begin VB.TextBox TxtLotExchangeType 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   325
            Left            =   7800
            Locked          =   -1  'True
            TabIndex        =   104
            TabStop         =   0   'False
            Top             =   960
            Width           =   2640
         End
         Begin VB.TextBox TxtLotConfirmedQtyDP 
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
            TabIndex        =   89
            Top             =   420
            Width           =   1695
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
            Left            =   10560
            TabIndex        =   88
            Top             =   960
            Width           =   1215
         End
         Begin VB.TextBox TxtLotRematQty 
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
            Left            =   6000
            TabIndex        =   87
            Top             =   420
            Width           =   1575
         End
         Begin VB.TextBox TxtLotNo 
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
            TabIndex        =   86
            Top             =   420
            Width           =   2535
         End
         Begin VB.TextBox TxtLotConfirmedQtyCT 
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
            TabIndex        =   85
            Top             =   420
            Width           =   1935
         End
         Begin VB.TextBox TxtLotActualPending 
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
            Left            =   11895
            TabIndex        =   84
            Top             =   420
            Width           =   1815
         End
         Begin VB.Label Label20 
            Caption         =   "Exchange Type"
            Height          =   300
            Left            =   6000
            TabIndex        =   105
            Top             =   975
            Width           =   1440
         End
         Begin VB.Label Label19 
            Caption         =   "Location"
            Height          =   300
            Left            =   120
            TabIndex        =   103
            Top             =   972
            Width           =   1200
         End
         Begin VB.Label Label15 
            Caption         =   "Commodity"
            Height          =   300
            Left            =   3600
            TabIndex        =   102
            Top             =   135
            Width           =   1140
         End
         Begin VB.Label Label17 
            Caption         =   "Confirmed Qty DP"
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
            Left            =   7800
            TabIndex        =   94
            Top             =   120
            Width           =   1575
         End
         Begin VB.Label Label16 
            Caption         =   "Remat Qty"
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
            Left            =   6120
            TabIndex        =   93
            Top             =   120
            Width           =   1335
         End
         Begin VB.Label Label30 
            Caption         =   "CMSE- CommTrack Lot No"
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
            Left            =   240
            TabIndex        =   92
            Top             =   120
            Width           =   2295
         End
         Begin VB.Label Label18 
            Caption         =   "Confirmed Qty ComTrack"
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
            Left            =   9600
            TabIndex        =   91
            Top             =   120
            Width           =   2175
         End
         Begin VB.Label Label54 
            Caption         =   "Actual Pending Qty"
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
            Left            =   12075
            TabIndex        =   90
            Top             =   120
            Width           =   1695
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "CITF And Converted Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   -74880
         TabIndex        =   79
         Top             =   2400
         Width           =   14685
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
            Height          =   1680
            Left            =   75
            TabIndex        =   80
            Top             =   270
            Width           =   14520
            _ExtentX        =   25612
            _ExtentY        =   2963
            _Version        =   393216
            Rows            =   3
            FixedRows       =   2
            FixedCols       =   0
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "CA Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   -74880
         TabIndex        =   77
         Top             =   360
         Width           =   14565
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
            Height          =   1680
            Left            =   75
            TabIndex        =   78
            Top             =   270
            Width           =   14400
            _ExtentX        =   25400
            _ExtentY        =   2963
            _Version        =   393216
            Rows            =   3
            FixedRows       =   2
            FixedCols       =   0
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid4 
         Height          =   2400
         Left            =   120
         TabIndex        =   81
         Top             =   2040
         Width           =   14520
         _ExtentX        =   25612
         _ExtentY        =   4233
         _Version        =   393216
         Rows            =   3
         FixedRows       =   2
         FixedCols       =   0
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label13 
         Caption         =   "Qty Updated :"
         Height          =   195
         Left            =   -74880
         TabIndex        =   100
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Created  :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   98
         Top             =   635
         Width           =   1695
      End
      Begin VB.Label Label9 
         Caption         =   "Updated :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   97
         Top             =   1220
         Width           =   1695
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1200
      Left            =   135
      TabIndex        =   30
      Top             =   8640
      Width           =   14805
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   400
         Left            =   11775
         TabIndex        =   31
         Top             =   645
         Width           =   2835
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "Search"
         Height          =   400
         Left            =   11775
         TabIndex        =   32
         Top             =   240
         Width           =   2835
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   400
         Left            =   8940
         TabIndex        =   11
         Top             =   645
         Width           =   2835
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   400
         Left            =   8940
         TabIndex        =   12
         Top             =   240
         Width           =   2835
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   400
         Left            =   5985
         TabIndex        =   13
         Top             =   645
         Width           =   2955
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   400
         Left            =   5985
         TabIndex        =   14
         Top             =   240
         Width           =   2955
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   400
         Left            =   3030
         TabIndex        =   15
         Top             =   645
         Width           =   2955
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   400
         Left            =   3030
         TabIndex        =   9
         Top             =   240
         Width           =   2955
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   400
         Left            =   75
         TabIndex        =   16
         Top             =   645
         Width           =   2955
      End
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New"
         Height          =   400
         Left            =   75
         TabIndex        =   17
         Top             =   240
         Width           =   2955
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3915
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   14820
      Begin VB.ComboBox CmbCommodity 
         Height          =   315
         Left            =   11520
         Sorted          =   -1  'True
         TabIndex        =   101
         Top             =   1425
         Visible         =   0   'False
         Width           =   3120
      End
      Begin VB.CheckBox Check1 
         Caption         =   "CommTrack RRN"
         Height          =   375
         Left            =   4920
         TabIndex        =   82
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox TxtAdjQty 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   360
         Left            =   13320
         Locked          =   -1  'True
         TabIndex        =   75
         Top             =   240
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.TextBox TxtDeliveredQty 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   360
         Left            =   8880
         Locked          =   -1  'True
         TabIndex        =   73
         Top             =   2940
         Width           =   1365
      End
      Begin VB.TextBox TxtRemarks 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   1800
         MaxLength       =   500
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   3405
         Width           =   5580
      End
      Begin VB.TextBox TxtOutOfSystemQty 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   5580
         TabIndex        =   7
         Top             =   2940
         Width           =   1800
      End
      Begin VB.TextBox TxtFlag 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Left            =   1800
         MaxLength       =   1
         TabIndex        =   70
         Top             =   3990
         Visible         =   0   'False
         Width           =   3600
      End
      Begin VB.TextBox TxtRRNQty 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1800
         TabIndex        =   6
         Top             =   2940
         Width           =   1200
      End
      Begin VB.TextBox TxtCommodity 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   325
         Left            =   11520
         Locked          =   -1  'True
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   1905
         Width           =   3120
      End
      Begin VB.ComboBox CmbClientDP 
         Height          =   315
         Left            =   1800
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   1410
         Visible         =   0   'False
         Width           =   8445
      End
      Begin VB.TextBox TxtOldDPID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   325
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   1440
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.TextBox TxtISINDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   325
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   63
         Top             =   1800
         Width           =   5565
      End
      Begin VB.TextBox TxtExchangeTypeID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   325
         Left            =   11520
         Locked          =   -1  'True
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   2265
         Width           =   3120
      End
      Begin VB.TextBox TxtLocationID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   325
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   2505
         Width           =   8445
      End
      Begin VB.TextBox TxtCMPID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   325
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   2160
         Width           =   8445
      End
      Begin VB.TextBox TxtGrade 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   325
         Left            =   11520
         Locked          =   -1  'True
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   2625
         Width           =   3120
      End
      Begin VB.TextBox TxtDPID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   325
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   1425
         Width           =   2850
      End
      Begin VB.TextBox TxtDPName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   325
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   1425
         Width           =   4125
      End
      Begin VB.TextBox TxtISINID 
         Appearance      =   0  'Flat
         Height          =   325
         Left            =   1800
         TabIndex        =   5
         Top             =   1800
         Width           =   2850
      End
      Begin VB.TextBox TxtClientName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   360
         Left            =   6120
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   1035
         Width           =   4125
      End
      Begin VB.TextBox TxtClientID 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1800
         TabIndex        =   3
         Top             =   1005
         Width           =   2880
      End
      Begin VB.TextBox TxtBalanceQty 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         CausesValidation=   0   'False
         Height          =   325
         Left            =   11520
         Locked          =   -1  'True
         TabIndex        =   50
         Top             =   3060
         Width           =   3120
      End
      Begin VB.CommandButton CmdClientID 
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
         Height          =   300
         Left            =   4710
         TabIndex        =   35
         Top             =   1035
         Width           =   375
      End
      Begin VB.CommandButton CmdISINID 
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
         Height          =   300
         Left            =   10305
         TabIndex        =   34
         Top             =   1395
         Width           =   375
      End
      Begin VB.TextBox TxtRRNID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   360
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   240
         Width           =   2880
      End
      Begin VB.TextBox TxtRRNNo 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1800
         MaxLength       =   20
         TabIndex        =   1
         Top             =   615
         Width           =   2880
      End
      Begin MSComCtl2.DTPicker RRNDate 
         Height          =   360
         Left            =   6120
         TabIndex        =   2
         Top             =   630
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   635
         _Version        =   393216
         Format          =   48824321
         CurrentDate     =   39864
      End
      Begin MSComCtl2.DTPicker EntryDate 
         Height          =   360
         Left            =   9375
         TabIndex        =   28
         Top             =   270
         Visible         =   0   'False
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   48824321
         CurrentDate     =   39864
      End
      Begin VB.Label Label14 
         Caption         =   "Converted Adjst Qty"
         Height          =   405
         Left            =   12120
         TabIndex        =   74
         Top             =   285
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.Label Label12 
         Caption         =   "Remarks"
         Height          =   210
         Left            =   240
         TabIndex        =   72
         Top             =   3450
         Width           =   1800
      End
      Begin VB.Label Label7 
         Caption         =   "Qty Delived before StoreMan Starting Date"
         Height          =   400
         Left            =   3150
         TabIndex        =   71
         Top             =   2920
         Width           =   1680
      End
      Begin VB.Label lblNameDid 
         Caption         =   "ClientDPID"
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
         Left            =   255
         TabIndex        =   46
         Top             =   1425
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label LblExchangeTypeID 
         Caption         =   "Exchange Type"
         Height          =   300
         Left            =   10305
         TabIndex        =   62
         Top             =   2325
         Width           =   1440
      End
      Begin VB.Label LblLocationID 
         Caption         =   "Location"
         Height          =   300
         Left            =   240
         TabIndex        =   61
         Top             =   2520
         Width           =   1200
      End
      Begin VB.Label LblCMPID 
         Caption         =   "CMP"
         Height          =   300
         Left            =   240
         TabIndex        =   60
         Top             =   2175
         Width           =   1080
      End
      Begin VB.Label Label2 
         Caption         =   "Grade"
         Height          =   180
         Left            =   10320
         TabIndex        =   59
         Top             =   2685
         Width           =   900
      End
      Begin VB.Label LblCommodity 
         Caption         =   "Commodity"
         Height          =   300
         Left            =   10320
         TabIndex        =   58
         Top             =   1995
         Width           =   1140
      End
      Begin VB.Label Label8 
         Caption         =   "Client Name "
         Height          =   255
         Left            =   5160
         TabIndex        =   49
         Top             =   1035
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "DP ID"
         Height          =   255
         Left            =   240
         TabIndex        =   48
         Top             =   1455
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Balance Qty in MT"
         Height          =   405
         Left            =   10320
         TabIndex        =   33
         Top             =   3045
         Width           =   840
      End
      Begin VB.Label LblRRNID 
         Caption         =   "RRN ID"
         Height          =   240
         Left            =   240
         TabIndex        =   19
         Top             =   300
         Width           =   1800
      End
      Begin VB.Label LblRRNNo 
         Caption         =   "RRN No"
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
         Left            =   240
         TabIndex        =   21
         Top             =   690
         Width           =   1800
      End
      Begin VB.Label LblRRNDate 
         Caption         =   "RRN Date"
         Height          =   180
         Left            =   4920
         TabIndex        =   22
         Top             =   720
         Width           =   1440
      End
      Begin VB.Label LblISINID 
         Caption         =   "CMSE Lot  No"
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
         TabIndex        =   23
         Top             =   1875
         Width           =   1800
      End
      Begin VB.Label LblClientID 
         Caption         =   "Client ID"
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
         TabIndex        =   24
         Top             =   1035
         Width           =   1800
      End
      Begin VB.Label LblRRNQty 
         Caption         =   "RRN Qty in MT"
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
         Left            =   240
         TabIndex        =   25
         Top             =   3008
         Width           =   1560
      End
      Begin VB.Label LblDeliveredQty 
         Caption         =   "Qty Delivered in StoreMan"
         Height          =   405
         Left            =   7620
         TabIndex        =   26
         Top             =   2925
         Width           =   1200
      End
      Begin VB.Label LblEntryDate 
         Caption         =   "Entry Date"
         Enabled         =   0   'False
         Height          =   300
         Left            =   8415
         TabIndex        =   27
         Top             =   300
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label LblFlag 
         Caption         =   "Flag"
         Enabled         =   0   'False
         Height          =   300
         Left            =   255
         TabIndex        =   29
         Top             =   4020
         Visible         =   0   'False
         Width           =   1800
      End
      Begin VB.Label Label11 
         Caption         =   "DP Name "
         Height          =   255
         Left            =   5160
         TabIndex        =   47
         Top             =   1440
         Width           =   735
      End
   End
End
Attribute VB_Name = "FrmMst_RRN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg As Variant
Dim mISINID, mClientID As String
Dim mISIN_LotID As String
Dim mClientIDRow As Double
Dim mCMPName, mCMPID, mLocationID, mCommodityID, mExchangeTypeID As Double
Dim mSLocationID As Double
Dim x As Integer


Private Sub Check1_Click()
Call Grid_Head3
TxtRRNQty = ""
If Check1.Value = 1 Then
    TxtRRNQty.BackColor = -2147483648#
    TxtRRNQty.Locked = True
    Pic_TTD.Enabled = True
    
'    TxtCommodity.Visible = False
'    CmbCommodity.Visible = True
Else
    TxtRRNQty.BackColor = -2147483643
    TxtRRNQty.Locked = False
    Pic_TTD.Enabled = False

'    TxtCommodity.Visible = True
'    CmbCommodity.Visible = False

End If

End Sub

'Private Sub CmbCommodity_GotFocus()
'tmp = CmbCommodity.Text
'Call TableMst_Commodity(" Where CommTrackFlag = 1")
'CmbCommodity.Clear
'If RsMst_Commodity.RecordCount = 0 Then
'   MsgBox "No Record found !!!! "
'   Exit Sub
'End If
'For i = 1 To RsMst_Commodity.RecordCount
' CmbCommodity.AddItem RsMst_Commodity!Commodity
' RsMst_Commodity.MoveNext
'Next
'CmbCommodity.Text = tmp
'End Sub
'
'Private Sub CmbCommodity_LostFocus()
'If CmbCommodity.Text = "" Then
'   Exit Sub
'End If
'RsMst_Commodity.MoveFirst
'RsMst_Commodity.Find "Commodity = '" & CmbCommodity.Text & "'"
'If RsMst_Commodity.EOF Then
'   MsgBox "Invalid selection "
'   CmbCommodity.SetFocus
'   Exit Sub
'End If
'
'mCommodityID = RsMst_Commodity!CommodityID
'
'End Sub


Private Sub CmbSCmp_GotFocus()
If CmbSLocation.Text = "" Then
   MsgBox "Please select location!!!"
   CmbSLocation.SetFocus
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

Private Sub CmbSLocation_GotFocus()
tmp = CmbSLocation.Text
If Gen_WcLocation <> "" Then
   Call TableMst_Location(" Where " & Gen_WcLocation & " ")
Else
   Call TableMst_Location
End If
CmbSLocation.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbSLocation.AddItem RsMst_Location!LocationName
    RsMst_Location.MoveNext
Next
CmbSLocation.Text = tmp
End Sub
Private Sub CmbSLocation_LostFocus()
If CmbSLocation.Text = "" Then
   mSLocationID = 0
   CmbSCmp.Text = ""
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

Private Sub CmdDeleteList_Click()

I = MSHFlexGrid4.RowSel
If I > 0 Then
   If MSHFlexGrid4.TextMatrix(I, 0) <> "" Then
      MSHFlexGrid4.TextMatrix(0, 3) = Val(MSHFlexGrid4.TextMatrix(0, 3)) - Val(MSHFlexGrid4.TextMatrix(I, 3))
      MSHFlexGrid4.TextMatrix(0, 4) = Val(MSHFlexGrid4.TextMatrix(0, 4)) - Val(MSHFlexGrid4.TextMatrix(I, 4))
      MSHFlexGrid4.TextMatrix(0, 5) = Val(MSHFlexGrid4.TextMatrix(0, 5)) - Val(MSHFlexGrid4.TextMatrix(I, 5))
      MSHFlexGrid4.TextMatrix(0, 2) = Val(MSHFlexGrid4.TextMatrix(0, 2)) - Val(MSHFlexGrid4.TextMatrix(I, 2))
      
      For RO = I To MSHFlexGrid4.Rows - 2
          For C = 0 To MSHFlexGrid4.Cols - 1
              MSHFlexGrid4.TextMatrix(RO, C) = MSHFlexGrid4.TextMatrix(RO + 1, C)
          Next
      Next
      MSHFlexGrid4.Rows = MSHFlexGrid4.Rows - 1
   End If
End If
TxtLotNo = ""
TxtLotCommodity = ""
TxtLotRematQty = ""
TxtLotConfirmedQtyDP = ""
TxtLotConfirmedQtyCT = ""
TxtLotActualPending = ""
TxtLotLocation = ""
TxtLotExchangeType = ""
TxtRRNQty = MSHFlexGrid4.TextMatrix(0, 2)
CmdDeleteList.Enabled = False


End Sub

Private Sub CmdSearch_Click()
Dim x As Integer
Dim wc As String
wc = ""

RsMst_ISIN.MoveFirst
Call Grid_Head

'If LCase(Gen_UserRole) <> "power" Then
'   If TxtSearchRRN = "" Or TxtSearchISIN = "" Then
'      MsgBox "Please provide some details for both RRN and ISIN!!!"
'      TxtSearchRRN.SetFocus
'      Exit Sub
'   End If
'End If

wc = " Where RRN.RRNDate between '" & SFrom.Value & "' And '" & STo.Value & "' "

If TxtSearchRRN <> "" Then
   'tmp = "Select * from Mst_ISIN where ISINID like '%" & Trim(TxtSearchISIN) & "%'"
   'Wc = " where rrn.RRNNo like '%" & Trim(TxtSearchRRN) & "%' "
'   If LCase(Gen_UserRole) = "power" Then
      wc = GenWc(wc, TxtSearchRRN, "RRN.RRNNo", "S", "L")
'   Else
'      Wc = GenWc(Wc, TxtSearchRRN, "RRN.RRNNo", "S", "N")
'   End If
End If
If CmbSCmp.Text <> "" Then
   wc = GenWc(wc, mCMPName, "cmp.CMPID", "N", "N")
End If
If CmbSLocation.Text <> "" Then
   wc = GenWc(wc, mSLocationID, "l.LocationID", "N", "N")
Else
   If Gen_WcLocation <> "" Then
      If wc = "" Then
         wc = " Where l." & Gen_WcLocation
      Else
         wc = " And l." & Gen_WcLocation
      End If
   End If
End If
If CmbSCommodity.Text <> "" Then
   wc = GenWc(wc, mCommodityID, "com.CommodityID", "N", "N")
End If
If TxtSearchISIN.Text <> "" Then
'   If LCase(Gen_UserRole) = "power" Then
      wc = GenWc(wc, TxtSearchISIN, "rrn.ISINID", "S", "L")
'   Else
'      Wc = GenWc(Wc, TxtSearchISIN, "rrn.ISINID", "S", "N")
'   End If
End If

tmp = "select rrn.RRNID, rrn.RRNNo, rrn.RRNDate, rrn.ISINID, l.LocationName, cmp.CMPName, com.Commodity, "
tmp = tmp & " c.ClientID + ':'+ c.ClientName As Client , dp.DepositoryID + ':'+ dp.DepositoryName As [DP Name], "
tmp = tmp & " rrn.RRNQty,rrn.DeliveredQty,rrn.OutOfSystemQty,rrn.remarks,rrn.EntryDate,rrn.Flag,rrn.G_UID,rrn.CommTrackFlag "
tmp = tmp & " from mst_rrn rrn "
tmp = tmp & " inner join Mst_ISIN isin on rrn.ISINID = isin.ISINID "
tmp = tmp & " inner join Mst_CMP cmp on isin.CMPID = cmp.CMPID "
tmp = tmp & " inner join mst_location l on cmp.LocationID = l.LocationID "
tmp = tmp & " inner join Mst_Commodity com on isin.CommodityID = com.CommodityID "
tmp = tmp & " inner join Mst_Client c on  rrn.ClientIDRow = c.ClientIDRow  "
tmp = tmp & " inner join Mst_Depository dp on c.DepositoryID = dp.DepositoryID and dp.ExchangeTypeID = 2 "
tmp = tmp & wc
tmp = tmp & " order by rrn.RRNNo,cmp.CMPName,com.Commodity "
Call Tmp4Table

If TmpRs4.RecordCount <= 0 Then
   MsgBox "No Record found!!!"
   TxtSearchRRN.SetFocus
   Exit Sub
End If
Call ShowGrid
End Sub
Private Sub ShowGrid()
MSHFlexGrid1.Rows = TmpRs4.RecordCount + 2
For x = 1 To TmpRs4.RecordCount
    For C = 0 To MSHFlexGrid1.Cols - 1
        MSHFlexGrid1.TextMatrix(x, C) = IIf(IsNull(TmpRs4.Fields(C)), "", TmpRs4.Fields(C))
    Next
    If MSHFlexGrid1.TextMatrix(x, 16) = "" Then
       MSHFlexGrid1.TextMatrix(x, 16) = 0
    End If
    If MSHFlexGrid1.TextMatrix(x, 16) = 1 Then
       MSHFlexGrid1.TextMatrix(x, 16) = "Yes"
    Else
       MSHFlexGrid1.TextMatrix(x, 16) = "No"
    End If
    TmpRs4.MoveNext
Next
TmpRs4.Close
End Sub

Private Sub CmdTTDAdd2List_Click()
Dim r As Variant

Call TxtLotNo_LostFocus

If TxtLotNo = "" Then
   MsgBox "Blank Lot No!!! "
   TxtLotNo.SetFocus
   Exit Sub
End If
If Val(TxtLotRematQty) = 0 Then
   MsgBox "Blank  Remat Qty !!! "
   TxtLotRematQty.SetFocus
   Exit Sub
End If

If TxtLotExchangeType <> TxtExchangeTypeID Then
   MsgBox "Invalid Exchange Type Selection !!!"
   TxtLotNo.SetFocus
   Exit Sub
End If

If TxtLotLocation <> TxtLocationID Then
   MsgBox "Invalid Location Selection!!!"
   TxtLotNo.SetFocus
   Exit Sub
End If


If TxtLotCommodity <> TxtCommodity Then
   MsgBox "Invalid  Commodity Selection!!!"
   TxtLotNo.SetFocus
   Exit Sub
End If


If CheckDuplicateLotNo() = True Then
   MsgBox "Duplicate Lot No not Allowed!!! "
   TxtLotNo.SetFocus
   Exit Sub
End If


Call TableMst_RRN_LotDetail(" Where ISINID = '" & mISIN_LotID & "'")

If Edit_Flg = "A" Then
    If RsMst_RRNLotDetail.RecordCount > 0 Then
       MsgBox "This Lot No is already used in RRN ID " & RsMst_RRNLotDetail!RRNID & " !!! "
       TxtLotNo.SetFocus
       Exit Sub
    End If
End If


If CmdTTDAdd2List.Caption = "Update List" Then
   r = MSHFlexGrid4.RowSel
   MSHFlexGrid4.TextMatrix(0, 3) = Val(MSHFlexGrid4.TextMatrix(0, 3)) - Val(MSHFlexGrid4.TextMatrix(r, 3))
   MSHFlexGrid4.TextMatrix(0, 4) = Val(MSHFlexGrid4.TextMatrix(0, 4)) - Val(MSHFlexGrid4.TextMatrix(r, 4))
   MSHFlexGrid4.TextMatrix(0, 5) = Val(MSHFlexGrid4.TextMatrix(0, 5)) - Val(MSHFlexGrid4.TextMatrix(r, 5))
   MSHFlexGrid4.TextMatrix(0, 2) = Val(MSHFlexGrid4.TextMatrix(0, 2)) - Val(MSHFlexGrid4.TextMatrix(r, 2))
Else
   r = MSHFlexGrid4.Rows - 1
   MSHFlexGrid4.Rows = MSHFlexGrid4.Rows + 1
   MSHFlexGrid4.TextMatrix(r, 6) = ""
End If


MSHFlexGrid4.TextMatrix(r, 0) = TxtLotNo
MSHFlexGrid4.TextMatrix(r, 1) = TxtLotCommodity
MSHFlexGrid4.TextMatrix(r, 2) = TxtLotRematQty
MSHFlexGrid4.TextMatrix(r, 3) = TxtLotConfirmedQtyDP
MSHFlexGrid4.TextMatrix(r, 4) = TxtLotConfirmedQtyCT
MSHFlexGrid4.TextMatrix(r, 5) = TxtLotActualPending


MSHFlexGrid4.TextMatrix(0, 2) = Val(MSHFlexGrid4.TextMatrix(0, 2)) + Val(MSHFlexGrid4.TextMatrix(r, 2))
MSHFlexGrid4.TextMatrix(0, 3) = Val(MSHFlexGrid4.TextMatrix(0, 3)) + Val(MSHFlexGrid4.TextMatrix(r, 3))
MSHFlexGrid4.TextMatrix(0, 4) = Val(MSHFlexGrid4.TextMatrix(0, 4)) + Val(MSHFlexGrid4.TextMatrix(r, 4))
MSHFlexGrid4.TextMatrix(0, 5) = Val(MSHFlexGrid4.TextMatrix(0, 5)) + Val(MSHFlexGrid4.TextMatrix(r, 5))

TxtLotNo = ""
TxtLotCommodity = ""
TxtLotRematQty = ""
TxtLotConfirmedQtyDP = ""
TxtLotConfirmedQtyCT = ""
TxtLotActualPending = ""
TxtLotLocation = ""
TxtLotExchangeType = ""
TxtRRNQty = MSHFlexGrid4.TextMatrix(0, 2)
CmdDeleteList.Enabled = False
                           
End Sub
Function CheckDuplicateLotNo() As Boolean
Dim Retval As Boolean
Retval = False
If MSHFlexGrid4.Rows > 3 Then
   For u = 1 To MSHFlexGrid4.Rows - 1
       If LCase(MSHFlexGrid4.TextMatrix(u, 0)) = LCase(TxtLotNo.Text) Then
          If LCase(CmdTTDAdd2List.Caption) = "add to list" Then
             Retval = True
             Exit For
          Else
             If u <> MSHFlexGrid4.RowSel Then
                Retval = True
                Exit For
             End If
          End If
       End If
   Next
End If

CheckDuplicateLotNo = Retval
End Function
Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Call GButtonLockAD(Me)
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If
Call Grid_Head

SFrom.Value = Date - 30
STo.Value = Date
TxtSearchRRN = ""
TxtSearchISIN.Text = ""
'If LCase(Gen_UserRole) = "power" Then
   Label4.Visible = True
   Label5.Visible = True
   Label3.Visible = True
   CmbSCmp.Visible = True
   CmbSLocation.Visible = True
   CmbSCommodity.Visible = True
'Else
'   Label4.Visible = False
'   Label5.Visible = False
'   Label3.Visible = False
'   CmbSCmp.Visible = False
'   CmbSLocation.Visible = False
'   CmbSCommodity.Visible = False
'End If


CmbSCmp.Text = ""
CmbSLocation.Text = ""
CmbSCommodity.Text = ""

Call TableMst_RRN
If RsMst_RRN.RecordCount = 0 Then
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
Call Grid_Head1
Call Grid_Head2
Call Grid_Head3


Call Indata

End Sub


Private Sub MSHFlexGrid3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyC And Shift = vbCtrlMask Then
   With MSHFlexGrid3
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
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Mst_RRN.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Mst_RRN.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = I + 1 ' + 6
    For C = 0 To MSHFlexGrid1.Cols - 2
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



Private Sub MSHFlexGrid4_Click()

If SaveCmd.Enabled = False Then
   Exit Sub
End If
CmdTTDAdd2List.Caption = "Update List"
r = MSHFlexGrid4.RowSel
If r <= 1 Then Exit Sub
If MSHFlexGrid4.TextMatrix(r, 1) = "" Then
   CmdTTDAdd2List.Caption = "Add to List"
   Exit Sub
End If

TxtLotNo = MSHFlexGrid4.TextMatrix(r, 0)

TxtLotCommodity = MSHFlexGrid4.TextMatrix(r, 1)
TxtLotRematQty = MSHFlexGrid4.TextMatrix(r, 2)
TxtLotConfirmedQtyDP = MSHFlexGrid4.TextMatrix(r, 3)
TxtLotConfirmedQtyCT = MSHFlexGrid4.TextMatrix(r, 4)
TxtLotActualPending = MSHFlexGrid4.TextMatrix(r, 5)
CmdDeleteList.Enabled = True

End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_RRN.MoveNext
If RsMst_RRN.EOF Then
   RsMst_RRN.MoveLast
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
RsMst_RRN.MovePrevious
If RsMst_RRN.BOF Then
   RsMst_RRN.MoveFirst
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
RsMst_RRN.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsMst_RRN.MoveLast
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
   If RsMst_RRN.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   Call FirstCmd_Click
   Exit Sub
End If
Frame0.Height = Me.Height - 600
Frame0.Left = Frame1.Left - 100
Frame0.Top = Frame1.Top - 100
Frame0.Width = Me.Width - 150
Frame0.Visible = True
MSHFlexGrid1.Width = Frame0.Width - 300
MSHFlexGrid1.Height = Frame0.Height - 1600
RsMst_RRN.MoveFirst


''===========comment on 7 july 2012 for not clearing search grid=============
'MSHFlexGrid1.Clear
'MSHFlexGrid1.Cols = 3
'Call Grid_Head
'TxtSearchRRN = ""
'TxtSearchISIN.Text = ""
''If LCase(Gen_UserRole) = "power" Then
'   Label4.Visible = True
'   Label5.Visible = True
'   Label3.Visible = True
'   CmbSCmp.Visible = True
'   CmbSLocation.Visible = True
'   CmbSCommodity.Visible = True
''Else
''   Label4.Visible = False
''   Label5.Visible = False
''   Label3.Visible = False
''   CmbSCmp.Visible = False
''   CmbSLocation.Visible = False
''   CmbSCommodity.Visible = False
''End If
'
'CmbSCmp.Text = ""
'CmbSLocation.Text = ""
'CmbSCommodity.Text = ""
''===========comment on 7 july 2012 for not clearing search grid=============
End Sub
Public Sub Indata()
Check1.Value = 0
Call Grid_Head3
Pic_TTD.Enabled = False
TxtRRNID.Locked = True
TxtRRNNo.Locked = True
RRNDate.Enabled = False
TxtISINID.Locked = True
'CmbISINID.Visible = False
CmdISINID.Visible = False
TxtISINID.Visible = True
TxtClientID.Locked = True
'CmbClientId.Visible = False
CmdClientID.Visible = False
CmbClientDP.Visible = False
lblNameDid.Visible = False
TxtClientID.Visible = True
TxtRRNQty.Locked = True
TxtDeliveredQty.Locked = True
TxtBalanceQty.Locked = True
'TxtAdjQty.Locked = True
EntryDate.Enabled = False
TxtFlag.Locked = True
TxtOutOfSystemQty.Locked = True
TxtRemarks.Locked = True

Check1.Value = IIf(IsNull(RsMst_RRN!CommTrackFlag), 0, RsMst_RRN!CommTrackFlag)

TxtRRNID = IIf(IsNull(RsMst_RRN!RRNID), "", RsMst_RRN!RRNID)
TxtRRNNo = IIf(IsNull(RsMst_RRN!RRNNo), "", RsMst_RRN!RRNNo)
RRNDate = IIf(IsNull(RsMst_RRN!RRNDate), "", RsMst_RRN!RRNDate)
mISINID = IIf(IsNull(RsMst_RRN!ISINID), "", RsMst_RRN!ISINID)
'TxtISINID = GetISINName(RsMst_RRN!ISINID)
Call TableMst_ISIN(" Where ISINID = '" & RsMst_RRN!ISINID & "'")
TxtISINID = RsMst_RRN!ISINID
TxtISINDesc = RsMst_ISIN!ISINDescription
mCMPID = IIf(IsNull(RsMst_ISIN!CMPID), "", RsMst_ISIN!CMPID)
TxtCMPID = GetCMPName(RsMst_ISIN!CMPID) '(IsNull(RsMst_ISIN!CMPName), "", RsMst_ISIN!CMPName)
mLocationID = IIf(IsNull(RsMst_ISIN!LocationID), "", RsMst_ISIN!LocationID)
TxtLocationID = GetLocationName(RsMst_ISIN!LocationID) '(IsNull(RsMst_ISIN!LocationName), "", RsMst_ISIN!LocationName)
mCommodityID = IIf(IsNull(RsMst_ISIN!CommodityID), "", RsMst_ISIN!CommodityID)
TxtCommodity = GetCommodityName(RsMst_ISIN!CommodityID) 'IsNull(RsMst_ISIN!Commodity), "", RsMst_ISIN!Commodity)
mExchangeTypeID = IIf(IsNull(RsMst_ISIN!ExchangeTypeID), "", RsMst_ISIN!ExchangeTypeID)
TxtExchangeTypeID = GetExchangeName(RsMst_ISIN!ExchangeTypeID) '(IsNull(RsMst_ISIN!ExchangeType), "", RsMst_ISIN!ExchangeType)
TxtGrade = IIf(IsNull(RsMst_ISIN!Grade), "", RsMst_ISIN!Grade)
'TxtISINID = IIf(IsNull(RsMst_RRN!ISINDescription), "", RsMst_RRN!ISINDescription)

mClientIDRow = IIf(IsNull(RsMst_RRN!ClientIDRow), "", RsMst_RRN!ClientIDRow)


Dim RetGetVal As Variant

RetGetVal = GetClientNameByRow((RsMst_RRN!ClientIDRow))   '(IsNull(RsMst_RRN!ClientIDRow), "", RsMst_RRN!ClientIDRow)

TxtClientID = RetGetVal(0) 'IIf(IsNull(RsMst_RRN!ClientID), "", RsMst_RRN!ClientID)
TxtClientName = RetGetVal(1)
TxtDPID = RetGetVal(2) 'GetDPFromClient(TxtClientID)
TxtOldDPID = RetGetVal(2) 'GetDPFromClient(TxtClientID)
TxtDPName = RetGetVal(3) 'GetDepositoryName(TxtDPID)

TxtRRNQty = IIf(IsNull(RsMst_RRN!RRNQty), "", RsMst_RRN!RRNQty)
TxtDeliveredQty = IIf(IsNull(RsMst_RRN!DeliveredQty), "", RsMst_RRN!DeliveredQty)
TxtOutOfSystemQty = IIf(IsNull(RsMst_RRN!OutOfSystemQty), "", RsMst_RRN!OutOfSystemQty)
'TxtAdjQty = IIf(IsNull(RsMst_RRN!AdjstQty), "", RsMst_RRN!AdjstQty)
TxtBalanceQty = Format(Val(TxtRRNQty) - (Val(TxtDeliveredQty) + Val(TxtOutOfSystemQty) + Val(TxtAdjQty)), "########0.000")
TxtRemarks = IIf(IsNull(RsMst_RRN!Remarks), "", RsMst_RRN!Remarks)

If IsNull(RsMst_RRN!EntryDate) = False Then
   'EntryDate = IIf(IsNull(RsMst_RRN!EntryDate), "", RsMst_RRN!EntryDate)
   EntryDate = RsMst_RRN!EntryDate
End If
TxtFlag = IIf(IsNull(RsMst_RRN!Flag), "", RsMst_RRN!Flag)


TxtCreated = IIf(IsNull(RsMst_RRN!CreatedBy), "", RsMst_RRN!CreatedBy) & " :- " & IIf(IsNull(RsMst_RRN!CreatedDate), "", RsMst_RRN!CreatedDate)
TxtUpdated = IIf(IsNull(RsMst_RRN!UpdatedBy), "", RsMst_RRN!UpdatedBy) & " :- " & IIf(IsNull(RsMst_RRN!UpdatedDate), "", RsMst_RRN!UpdatedDate)

TxtOutOfSystemUpdated = IIf(IsNull(RsMst_RRN!OutSystemUpdatedBy), "", RsMst_RRN!OutSystemUpdatedBy) & " :- " & IIf(IsNull(RsMst_RRN!OutSystemUpdatedDate), "", RsMst_RRN!OutSystemUpdatedDate)
'TxtOutOfSystemUpdated = ""

Call Grid_Head3
If Check1.Value = 1 Then
   Call TableMst_RRN_LotDetail(" where RRNID = " & RsMst_RRN!RRNID)
   For I = 1 To RsMst_RRNLotDetail.RecordCount
          MSHFlexGrid4.Rows = MSHFlexGrid4.Rows + 1
          MSHFlexGrid4.TextMatrix(I + 1, 0) = RsMst_RRNLotDetail!ISINID
          MSHFlexGrid4.TextMatrix(I + 1, 1) = TxtCommodity
          MSHFlexGrid4.TextMatrix(I + 1, 2) = RsMst_RRNLotDetail!RematQty
          MSHFlexGrid4.TextMatrix(I + 1, 6) = RsMst_RRNLotDetail!LotDetailID
          RsMst_RRNLotDetail.MoveNext
   Next
   MSHFlexGrid4.TextMatrix(0, 2) = TxtRRNQty
End If
Pic_TTD.Enabled = False
Call Indata_CA
Call Indata_CITF

If RsMst_RRN.RecordCount = 1 Then
   Call GButtonLock_1(Me, True)
Else
   Call GButtonLock(Me, True)
End If

If LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "superadmin" Then
   NextCmd.Enabled = False
   PreviousCmd.Enabled = False
   FirstCmd.Enabled = False
   LastCmd.Enabled = False
End If


End Sub
Private Sub Indata_CA()
Call Grid_Head1

tmp = "Select CANo,CADate,Quantity,PartialAcceptance from Mst_CA Where RRNID = " & TxtRRNID & ""

Call TmpTable

If TmpRs.RecordCount > 0 Then
   MSHFlexGrid2.Rows = TmpRs.RecordCount + 2
   mRow = 2
   For x = 1 To TmpRs.RecordCount
       For C = 0 To MSHFlexGrid2.Cols - 1
           MSHFlexGrid2.TextMatrix(mRow, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
       Next
       MSHFlexGrid2.TextMatrix(mRow, 1) = Format(MSHFlexGrid2.TextMatrix(mRow, 1), "dd-mmm-yyyy")
       MSHFlexGrid2.TextMatrix(0, 3) = Val(MSHFlexGrid2.TextMatrix(0, 3)) + Val(MSHFlexGrid2.TextMatrix(mRow, 3))
       mRow = mRow + 1
       TmpRs.MoveNext
   Next
End If

End Sub

Private Sub Indata_CITF()
Call Grid_Head2

'tmp = "Select CANo,CADate,Quantity,PartialAcceptance from Mst_CA Where RRNID = " & TxtRRNID & ""

tmp = "Select a.SM_Prefix,a.CxTFNo,b.CxTFDate,a.AdjRRNQty,a.AdjRRNNoofBags,b.CreatedDate From Txn_CxTF_ClientDetail a " & _
"inner join Txn_CxTF_Details b on a.SM_Prefix = b.SM_Prefix And a.cxtfno = b.cxtfno  And a.txntype = b.txntype " & _
"where b.BookedSpaceFlag <> 1 And b.exchangetypeid= 2 and b.txntype = 'W' and a.rrnid = " & Val(TxtRRNID)

Call TmpTable

If TmpRs.RecordCount > 0 Then
   MSHFlexGrid3.Rows = MSHFlexGrid3.Rows + TmpRs.RecordCount '+ 2
   mRow = 2
   For x = 1 To TmpRs.RecordCount
       For C = 0 To MSHFlexGrid3.Cols - 1
           MSHFlexGrid3.TextMatrix(mRow, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
       Next
       MSHFlexGrid3.TextMatrix(mRow, 2) = Format(MSHFlexGrid3.TextMatrix(mRow, 2), "dd-mmm-yyyy")
       MSHFlexGrid3.TextMatrix(0, 3) = Val(MSHFlexGrid3.TextMatrix(0, 3)) + Val(MSHFlexGrid3.TextMatrix(mRow, 3))
       MSHFlexGrid3.TextMatrix(0, 4) = Val(MSHFlexGrid3.TextMatrix(0, 4)) + Val(MSHFlexGrid3.TextMatrix(mRow, 4))
       MSHFlexGrid3.TextMatrix(mRow, 5) = Format(MSHFlexGrid3.TextMatrix(mRow, 5), "dd-mmm-yyyy hh:mm AM/PM")
       mRow = mRow + 1
       TmpRs.MoveNext
   Next
End If

''''''Add Converted RRN_PWR Detail
tmp = " Select ML.SM_Prefix, -TRP.ConvertedID ConvertedID, TRP.PWRDate, sum(TRPG.AdjstBags) AdjstBags, sum(TRPG.AdjstQty)+TRP.AdjstQty AdjstQty, TRP.CreatedDate  FROM Txn_RRNPWR TRP  "
tmp = tmp & " INNER JOIN Txn_RRNPWRGSL TRPG on TRPG.ConvertedID = TRP.ConvertedID "
tmp = tmp & " INNER JOIN Mst_Location ML on ML.LocationID = TRP.LocationID"
tmp = tmp & " WHERE TRP.RRNID = " & Val(TxtRRNID)
tmp = tmp & " GROUP BY ML.SM_Prefix, TRP.ConvertedID, TRP.PWRDate, TRP.CreatedDate, TRP.AdjstQty "
Call TmpTable

If TmpRs.RecordCount > 0 Then
   For I = 1 To TmpRs.RecordCount
      mRow = MSHFlexGrid3.Rows - 1
      
      MSHFlexGrid3.TextMatrix(mRow, 0) = IIf(IsNull(TmpRs!SM_Prefix), "", TmpRs!SM_Prefix)
      MSHFlexGrid3.TextMatrix(mRow, 1) = IIf(IsNull(TmpRs!ConvertedID), "", TmpRs!ConvertedID)
      MSHFlexGrid3.TextMatrix(mRow, 2) = IIf(IsNull(TmpRs!PWRDate), "", Format(TmpRs!PWRDate, "dd-MMM-yyyy"))
      MSHFlexGrid3.TextMatrix(mRow, 3) = IIf(IsNull(TmpRs!AdjstQty), "", TmpRs!AdjstQty)
      MSHFlexGrid3.TextMatrix(0, 3) = Val(MSHFlexGrid3.TextMatrix(0, 3)) + Val(MSHFlexGrid3.TextMatrix(mRow, 3))
      MSHFlexGrid3.TextMatrix(mRow, 4) = IIf(IsNull(TmpRs!AdjstBags), "", TmpRs!AdjstBags)
      MSHFlexGrid3.TextMatrix(0, 4) = Val(MSHFlexGrid3.TextMatrix(0, 4)) + Val(MSHFlexGrid3.TextMatrix(mRow, 4))
      MSHFlexGrid3.TextMatrix(mRow, 5) = IIf(IsNull(TmpRs!CreatedDate), "", Format(TmpRs!CreatedDate, "dd-mmm-yyyy hh:mm AM/PM"))
      MSHFlexGrid3.Rows = MSHFlexGrid3.Rows + 1
      TmpRs.MoveNext
   Next
End If
End Sub

Private Sub TxtISINID_GotFocus()
tmp = TxtISINID.Text
End Sub

Private Sub TxtISINID_LostFocus()
If SaveCmd.Enabled = False Then Exit Sub
If TxtISINID.Text = "" Then
   mCommodityID = 0
   mCMPID = 0
    'mExchangeTypeID = 0
   TxtCMPID = ""
   TxtLocationID = ""
   TxtCommodity = ""
   TxtExchangeTypeID = ""
   TxtGrade = ""
   Call Grid_Head
   Exit Sub
End If

If tmp <> TxtISINID Then
   mCommodityID = 0
   mCMPID = 0
    'mExchangeTypeID = 0
   TxtCMPID = ""
   TxtLocationID = ""
   TxtCommodity = ""
   TxtExchangeTypeID = ""
   TxtGrade = ""
   Call Grid_Head
End If


Call TableMst_ISIN("where ISINID = '" & TxtISINID.Text & "'")



If RsMst_ISIN.RecordCount = 0 Then
   MsgBox "No Such CMSE Lot No Exist !!!"
   TxtISINID.SetFocus
   Exit Sub
End If
mISINID = RsMst_ISIN!ISINID
mCMPID = RsMst_ISIN!CMPID
TxtCMPID = GetCMPName(RsMst_ISIN!CMPID) '(IsNull(RsMst_ISIN!CMPName), "", RsMst_ISIN!CMPName)
TxtISINDesc = RsMst_ISIN!ISINDescription
mLocationID = IIf(IsNull(RsMst_ISIN!LocationID), "", RsMst_ISIN!LocationID)
TxtLocationID = GetLocationName(RsMst_ISIN!LocationID) '(IsNull(RsMst_ISIN!LocationName), "", RsMst_ISIN!LocationName)
mCommodityID = IIf(IsNull(RsMst_ISIN!CommodityID), "", RsMst_ISIN!CommodityID)
TxtCommodity = GetCommodityName(RsMst_ISIN!CommodityID) 'IsNull(RsMst_ISIN!Commodity), "", RsMst_ISIN!Commodity)
mExchangeTypeID = IIf(IsNull(RsMst_ISIN!ExchangeTypeID), "", RsMst_ISIN!ExchangeTypeID)
TxtExchangeTypeID = GetExchangeName(RsMst_ISIN!ExchangeTypeID) '(IsNull(RsMst_ISIN!ExchangeType), "", RsMst_ISIN!ExchangeType)
TxtGrade = IIf(IsNull(RsMst_ISIN!Grade), "", RsMst_ISIN!Grade)


If Check1.Value = 1 And MSHFlexGrid4.Rows = 3 Then
   TxtLotNo = TxtISINID
   Call TxtLotNo_LostFocus
End If

End Sub

Private Sub TxtLotNo_LostFocus()

If SaveCmd.Enabled = False Then Exit Sub

TxtLotNo = Val(TxtLotNo)
If TxtLotNo = 0 Then Exit Sub

Call TableMst_ISIN("where ISINID = '" & TxtLotNo.Text & "'")
If RsMst_ISIN.RecordCount = 0 Then
   MsgBox "No Such CMSE Lot No /Lot  Exist !!!"
   TxtLotNo.SetFocus
   Exit Sub
End If

 
mISIN_LotID = RsMst_ISIN!ISINID

TxtLotLocation = GetLocationName(RsMst_ISIN!LocationID)  '(IsNull(RsMst_ISIN!LocationName), "", RsMst_ISIN!LocationName)
TxtLotCommodity = GetCommodityName(RsMst_ISIN!CommodityID) 'IsNull(RsMst_ISIN!Commodity), "", RsMst_ISIN!Commodity)
TxtLotExchangeType = GetExchangeName(RsMst_ISIN!ExchangeTypeID) '(IsNull(RsMst_ISIN!ExchangeType), "", RsMst_ISIN!ExchangeType)


End Sub

Private Sub TxtLotRematQty_LostFocus()

If TxtLotRematQty = "" Then Exit Sub
If IsNumeric(TxtLotRematQty) = False Or Val(TxtLotRematQty) < 0 Then
   MsgBox "Invalid numeric format !!!!"
   TxtLotRematQty.SetFocus
   Exit Sub
End If

End Sub

Private Sub TxtOutOfSystemQty_LostFocus()

If SaveCmd.Enabled = False Then Exit Sub
If TxtOutOfSystemQty = "" Then Exit Sub
If IsNumeric(TxtOutOfSystemQty) = False Or Val(TxtOutOfSystemQty) < 0 Then
   MsgBox "Invalid numeric format !!!!"
   TxtOutOfSystemQty.SetFocus
   Exit Sub
End If

If Val(TxtOutOfSystemQty) > Val(TxtRRNQty) Then
   MsgBox "Out of System Qty must be less than RRN Qty!!!"
   TxtOutOfSystemQty.SetFocus
   Exit Sub
End If


TxtOutOfSystemQty = Val(TxtOutOfSystemQty)
'TxtBalanceQty = Val(TxtRRNQty) - (Val(TxtDeliveredQty) + Val(TxtOutOfSystemQty) + Val(TxtAdjQty))
TxtBalanceQty = Val(TxtRRNQty) - (Val(TxtDeliveredQty) + Val(TxtOutOfSystemQty))
TxtBalanceQty = Format(TxtBalanceQty, "########0.000")

End Sub

Private Sub TxtRRNID_LostFocus()
If TxtRRNID = "" Then Exit Sub
If IsNumeric(TxtRRNID) = False Then
   MsgBox "Invalid numeric format !!!!"
   TxtRRNID.SetFocus
   Exit Sub
End If
TxtRRNID = Val(TxtRRNID)
End Sub
Private Sub TxtRRNNo_LostFocus()
If TxtRRNNo = "" Then Exit Sub
If SaveCmd.Enabled = False Then Exit Sub
If RsMst_RRN.RecordCount = 0 Then Exit Sub
TxtRRNNo = UCase(Trim(TxtRRNNo))

If Check1.Value = 1 Then
    RsMst_RRN.MoveFirst
    RsMst_RRN.Filter = "CommTrackFlag = 1"
    RsMst_RRN.Find "RRNNo = '" & TxtRRNNo & "'"
    If Not RsMst_RRN.EOF Then
       MsgBox "RRN already exist !!! "
       Call Indata
       Exit Sub
    End If
End If
End Sub

Private Sub TxtRRNQty_LostFocus()
If SaveCmd.Enabled = False Then Exit Sub
If TxtRRNQty = "" Then Exit Sub
If IsNumeric(TxtRRNQty) = False Then
   MsgBox "Invalid numeric format !!!!"
   TxtRRNQty.SetFocus
   Exit Sub
End If
TxtRRNQty = Val(TxtRRNQty)
'TxtBalanceQty = Val(TxtRRNQty) - (Val(TxtDeliveredQty) + Val(TxtOutOfSystemQty) + Val(TxtAdjQty))
TxtBalanceQty = Val(TxtRRNQty) - (Val(TxtDeliveredQty) + Val(TxtOutOfSystemQty))
TxtBalanceQty = Format(TxtBalanceQty, "########0.000")

End Sub
Private Sub TxtDeliveredQty_LostFocus()
If TxtDeliveredQty = "" Then Exit Sub
If IsNumeric(TxtDeliveredQty) = False Then
   MsgBox "Invalid numeric format !!!!"
   TxtDeliveredQty.SetFocus
   Exit Sub
End If
TxtDeliveredQty = Val(TxtDeliveredQty)
TxtBalanceQty = Val(TxtRRNQty) - Val(TxtDeliveredQty)
End Sub

Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)
Call Grid_Head3
TxtRRNID.Locked = True
TxtRRNNo.Locked = False
RRNDate.Enabled = True
TxtISINID.Locked = False
'CmbISINID.Visible = True
CmdISINID.Visible = True
'TxtISINID.Visible = False
TxtClientID.Locked = False
'CmbClientId.Visible = True
CmdClientID.Visible = True
'TxtClientID.Visible = False
TxtRRNQty.Locked = False
TxtOutOfSystemQty.Locked = False
TxtDeliveredQty.Locked = True
CmbClientDP.Visible = False
lblNameDid.Visible = False
EntryDate.Enabled = False
TxtFlag.Locked = False
TxtRRNID = ""
TxtRRNNo = ""
RRNDate = Date
TxtISINID = ""
TxtISINDesc = ""
TxtCMPID = ""
TxtLocationID = ""
TxtCommodity = ""
TxtGrade = ""
TxtExchangeTypeID = ""
CmbClientDP.Text = ""
TxtClientID = ""
TxtClientName = ""
TxtRRNQty = ""
TxtDPID.Text = ""
TxtDPName.Text = ""
TxtDeliveredQty = ""
TxtBalanceQty = ""
'TxtAdjQty = ""
EntryDate = Date
TxtFlag = ""
TxtOutOfSystemQty = ""
TxtRemarks = ""
TxtCreated = ""
TxtUpdated = ""
TxtRemarks.Locked = False
TxtOutOfSystemUpdated = ""
Check1.Value = 0
Pic_TTD.Enabled = False

If RsMst_RRN.RecordCount > 0 Then
   TxtRRNNo.SetFocus
End If


End Sub
Private Sub EditCmd_Click()

'tmp = " Select RRNID from Txn_RRNPWR where RRNID = " & Val(TxtRRNID.Text)
'Call TmpTable
'If TmpRs.RecordCount > 0 Then
'   MsgBox "RRN is converted into PWR, you can not edit!!!"
'   Exit Sub
'End If

If LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "headcmg" Then
   If Val(TxtDeliveredQty) <> 0 Then
      MsgBox "you Can't Edit This RRN since, CITF is generated !!! "
      Exit Sub
   End If
End If

Edit_Flg = "E"
Call GButtonLock(Me, False)

TxtRRNID.Locked = True
TxtRRNNo.Locked = False
RRNDate.Enabled = True
TxtISINID.Locked = False
TxtISINDesc.Locked = True
'CmbISINID.Visible = True
CmdISINID.Visible = True
'TxtISINID.Visible = False
TxtClientID.Locked = False
'CmbClientId.Visible = True
CmbClientDP.Visible = False
lblNameDid.Visible = False
CmdClientID.Visible = True
'TxtClientID.Visible = False
TxtRRNQty.Locked = False
TxtOutOfSystemQty.Locked = False
'TxtDeliveredQty.Locked = False
EntryDate.Enabled = False
TxtFlag.Locked = False
TxtRemarks.Locked = False

If Check1.Value = 1 Then
   Pic_TTD.Enabled = True
Else
   Pic_TTD.Enabled = False
End If


End Sub
Private Sub DeleteCmd_Click()

If Val(TxtDeliveredQty) <> 0 Then
   MsgBox "you Can't Delete This RRN since, CITF is generated !!! "
   Exit Sub
End If

tmp = "Select * from Mst_CA Where RRNID = " & TxtRRNID & ""
Call TmpTable

If TmpRs.RecordCount > 0 Then
   MsgBox "You Can't Delete This RRN !!! "
   Exit Sub
End If


Dim ans As Variant
On Error GoTo msgError
ans = MsgBox("Do you really want to DELETE this record???", vbYesNo)
If ans = vbYes Then
   If Check1.Value = 1 Then
      tmp = "delete from Mst_RRN_LotDetail where RRNid  = " & TxtRRNID & ""
      Call Tmp1Table
   End If
   RsMst_RRN.Delete
   RsMst_RRN.Update
   If RsMst_RRN.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_RRN.MoveNext
   If RsMst_RRN.EOF() Then
      RsMst_RRN.MoveLast
   End If
   Call Indata
End If
Exit Sub
msgError:
MsgBox "Child record Present "
RsMst_RRN.CancelUpdate
End Sub
Private Sub SaveCmd_Click()

If TxtRRNNo.Text = "" Then
   MsgBox "Blank RRNNo Not Allowed !!! "
   TxtRRNNo.SetFocus
   Exit Sub
End If


'If Check1.Value = 1 Then
'   If Edit_Flg = "A" Then
'
'
'   End If
'End If

If TxtISINID.Text = "" Then
    MsgBox "Blank CMSE Lot No not Allowed!!!"
    CmbISINID.SetFocus
    Exit Sub
End If

If TxtClientID.Text = "" Then
   MsgBox "Blank Client Id Not Allowed!!!"
   CmbClientID.SetFocus
   Exit Sub
End If

If TxtRRNQty.Text = "" Then
   MsgBox "Blank RRNQty Not Allowed !!! "
   TxtRRNQty.SetFocus
   Exit Sub
End If

If Val(TxtOutOfSystemQty.Text) <> 0 Then
   
   If Val(TxtBalanceQty.Text) < 0 Then
      MsgBox "Balance Qty should not be negative!!!"
      TxtOutOfSystemQty.SetFocus
      Exit Sub
   End If
   If Trim(TxtRemarks) = "" Then
      MsgBox "Blank Remarks Not Allowed !!! "
      TxtRemarks.SetFocus
      Exit Sub
   End If
End If

If Check1.Value = 1 Then
   If ChkISINinGrid = False Then
      MsgBox "ISIN " & mISINID & " should present in Commtrack lot details also!!!"
      Exit Sub
   End If
End If

If Edit_Flg = "A" Then
   If RsMst_RRN.RecordCount > 0 Then
      RsMst_RRN.MoveFirst
      tmp = "DepositoryID = '" & TxtDPID & "'"
      RsMst_RRN.Filter = tmp
      
      RsMst_RRN.Find "RRNNO= '" & TxtRRNNo.Text & "'"
      If Not RsMst_RRN.EOF Then
        MsgBox "Duplicate RRN No Not Allowed !!! "
         TxtRRNNo.SetFocus
         Exit Sub
      End If
   End If
      RsMst_RRN.AddNew
      RsMst_RRN!RRNID = GetRRNID
      RsMst_RRN!G_UID = GetGUID
      TxtRRNID.Text = RsMst_RRN!RRNID
Else
    RsMst_RRN.MoveFirst
    tmp = "DepositoryID = '" & TxtDPID & "'"
    RsMst_RRN.Filter = tmp
    RsMst_RRN.Find "RRNNo= '" & TxtRRNNo.Text & "'"
    If Not RsMst_RRN.EOF Then
       If RsMst_RRN!RRNID <> TxtRRNID.Text Then
          MsgBox "Duplicate RRN No Not Allowed !!! "
          TxtRRNNo.SetFocus
          Exit Sub
       End If
    End If
    If RsMst_RRN.RecordCount > 0 Then
       RsMst_RRN.MoveFirst
    End If
    tmp = "DepositoryID = '" & TxtOldDPID & "'"
    RsMst_RRN.Filter = tmp
    RsMst_RRN.Find " RRNID= " & TxtRRNID.Text
End If

RsMst_RRN!RRNID = TxtRRNID.Text
RsMst_RRN!RRNNo = TxtRRNNo.Text
RsMst_RRN!RRNDate = RRNDate.Value
RsMst_RRN!ISINID = mISINID
RsMst_RRN!ClientIDRow = mClientIDRow
RsMst_RRN!RRNQty = Val(TxtRRNQty.Text)
RsMst_RRN!DeliveredQty = Val(TxtDeliveredQty.Text)
RsMst_RRN!EntryDate = EntryDate.Value
RsMst_RRN!Flag = TxtFlag.Text
RsMst_RRN!DepositoryID = TxtDPID
RsMst_RRN!OutOfSystemQty = Val(TxtOutOfSystemQty)
RsMst_RRN!Remarks = TxtRemarks

If IsNull(RsMst_RRN!CreatedBy) = True Or RsMst_RRN!CreatedBy = "" Then
   RsMst_RRN!CreatedBy = Gen_UserLoginID
   RsMst_RRN!CreatedDate = Now
Else
   RsMst_RRN!UpdatedBy = Gen_UserLoginID
   RsMst_RRN!UpdatedDate = Now
End If

If TxtOutOfSystemQty <> "" Then
   RsMst_RRN!OutSystemUpdatedBy = Gen_UserLoginID
   RsMst_RRN!OutSystemUpdatedDate = Now
End If

RsMst_RRN!CommTrackFlag = Check1.Value
RsMst_RRN.Update

If Check1.Value = 1 Then
   Call TableMst_RRN_LotDetail
   For I = 2 To MSHFlexGrid4.Rows - 1
       If MSHFlexGrid4.TextMatrix(I, 0) <> "" Then
          If Val(MSHFlexGrid4.TextMatrix(I, 6)) <> 0 Then
             RsMst_RRNLotDetail.MoveFirst
             RsMst_RRNLotDetail.Find "LotDetailID = " & MSHFlexGrid4.TextMatrix(I, 6) & " "
          Else
             RsMst_RRNLotDetail.AddNew
             RsMst_RRNLotDetail!LotDetailID = GetRRNLotID
          End If
          RsMst_RRNLotDetail!RRNID = RsMst_RRN!RRNID
          RsMst_RRNLotDetail!ISINID = MSHFlexGrid4.TextMatrix(I, 0)
          RsMst_RRNLotDetail!RematQty = MSHFlexGrid4.TextMatrix(I, 2)
          RsMst_RRNLotDetail.Update
          RsMst_RRNLotDetail.Requery
       End If
   Next
End If
RsMst_RRN.Filter = ""
RsMst_RRN.Find "RRNID = " & TxtRRNID

Call Indata

End Sub
Private Function ChkISINinGrid() As Boolean

ChkISINinGrid = False
For x = 2 To MSHFlexGrid4.Rows - 1
    If MSHFlexGrid4.TextMatrix(x, 0) = mISINID Then
       ChkISINinGrid = True
       Exit For
    End If
Next

End Function


Private Function GetRRNID() As Double
Dim Retval As Double
tmp = "Select max(RRNID) from Mst_RRN"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetRRNID = Retval
End Function

Private Function GetRRNLotID() As Double
Dim Retval As Double
tmp = "Select max(LotDetailID) from Mst_RRN_LotDetail"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetRRNLotID = Retval
End Function

Private Sub Grid_Head()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 17
MSHFlexGrid1.TextMatrix(0, 0) = "RRN ID"
MSHFlexGrid1.TextMatrix(0, 1) = "RRN No"
MSHFlexGrid1.TextMatrix(0, 2) = "RRN Date"
MSHFlexGrid1.TextMatrix(0, 3) = "CMSE Lot No" '"ISIN No"
MSHFlexGrid1.TextMatrix(0, 4) = "Location"
MSHFlexGrid1.TextMatrix(0, 5) = "CMP"
MSHFlexGrid1.TextMatrix(0, 6) = "Commodity"
MSHFlexGrid1.TextMatrix(0, 7) = "Client"
MSHFlexGrid1.TextMatrix(0, 8) = "DP"
MSHFlexGrid1.TextMatrix(0, 9) = "RRN Qty in MT"
MSHFlexGrid1.TextMatrix(0, 10) = "Delivered Qty in MT"
MSHFlexGrid1.TextMatrix(0, 11) = "Out of System Qty in MT"
MSHFlexGrid1.TextMatrix(0, 12) = "Remarks"
MSHFlexGrid1.TextMatrix(0, 13) = "Entry Date"
MSHFlexGrid1.TextMatrix(0, 14) = "Flag"
MSHFlexGrid1.TextMatrix(0, 15) = "G-UID"
MSHFlexGrid1.TextMatrix(0, 16) = "Commtrack Flag"

MSHFlexGrid1.ColWidth(0) = 1200
MSHFlexGrid1.ColWidth(1) = 2000
MSHFlexGrid1.ColWidth(2) = 1000
MSHFlexGrid1.ColWidth(3) = 2000
MSHFlexGrid1.ColWidth(4) = 2000
MSHFlexGrid1.ColWidth(5) = 1600
MSHFlexGrid1.ColWidth(6) = 1600
MSHFlexGrid1.ColWidth(7) = 1400
MSHFlexGrid1.ColWidth(8) = 1200
MSHFlexGrid1.ColWidth(11) = 1200
MSHFlexGrid1.ColWidth(12) = 1600
MSHFlexGrid1.ColWidth(15) = 0

End Sub
Private Sub Grid_Head2()
MSHFlexGrid3.Clear
MSHFlexGrid3.Rows = 3
MSHFlexGrid3.FixedCols = 0
MSHFlexGrid3.Cols = 6

MSHFlexGrid3.TextMatrix(1, 0) = "Location ID"
MSHFlexGrid3.TextMatrix(1, 1) = "CITF No"
MSHFlexGrid3.TextMatrix(1, 2) = "CITF Date"
MSHFlexGrid3.TextMatrix(1, 3) = "CITF Qty"
MSHFlexGrid3.TextMatrix(1, 4) = "CITF No of Bags"
MSHFlexGrid3.TextMatrix(1, 5) = "Created Date"

MSHFlexGrid3.ColWidth(0) = 2400
MSHFlexGrid3.ColWidth(1) = 2400
MSHFlexGrid3.ColWidth(2) = 2400
MSHFlexGrid3.ColWidth(3) = 2400
MSHFlexGrid3.ColWidth(4) = 2400
MSHFlexGrid3.ColWidth(5) = 2400


End Sub



Private Sub Grid_Head1()
MSHFlexGrid2.Clear
MSHFlexGrid2.Rows = 3
MSHFlexGrid2.Cols = 4

MSHFlexGrid2.TextMatrix(1, 0) = "CA N0"
MSHFlexGrid2.TextMatrix(1, 1) = "CA Date"
MSHFlexGrid2.TextMatrix(1, 2) = "RRN Qty "
MSHFlexGrid2.TextMatrix(1, 3) = "CA Qty"


MSHFlexGrid2.ColWidth(0) = 2000
MSHFlexGrid2.ColWidth(1) = 2000
MSHFlexGrid2.ColWidth(2) = 2000
MSHFlexGrid2.ColWidth(3) = 2000


End Sub
Private Sub Grid_Head3()
MSHFlexGrid4.Clear
MSHFlexGrid4.Rows = 3
MSHFlexGrid4.FixedCols = 0
MSHFlexGrid4.Cols = 7

MSHFlexGrid4.TextMatrix(1, 0) = "CMSE- CommTrack Lot No" '"ISIN- CommTrack Lot No"
MSHFlexGrid4.TextMatrix(1, 1) = "Commodity"
MSHFlexGrid4.TextMatrix(1, 2) = "Remat Qty"
MSHFlexGrid4.TextMatrix(1, 3) = "Confirmed Qty DP"
MSHFlexGrid4.TextMatrix(1, 4) = "Confirmed Qty ComTrack"
MSHFlexGrid4.TextMatrix(1, 5) = "Actual Pending Qty"
MSHFlexGrid4.TextMatrix(1, 6) = "ISINID"

MSHFlexGrid4.ColWidth(0) = 2400
MSHFlexGrid4.ColWidth(1) = 2400
MSHFlexGrid4.ColWidth(2) = 2400
MSHFlexGrid4.ColWidth(3) = 2400
MSHFlexGrid4.ColWidth(4) = 2400
MSHFlexGrid4.ColWidth(5) = 2400
MSHFlexGrid4.ColWidth(6) = 0


End Sub

Private Sub MSHFlexGrid1_Click()
Dim I As Variant
I = MSHFlexGrid1.RowSel
If I <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(I, 0) = "" Then Exit Sub
RsMst_RRN.Filter = ""
RsMst_RRN.MoveFirst
'If LCase(Gen_UserRole) = "power" Then
   RsMst_RRN.Find "G_UID = '" & MSHFlexGrid1.TextMatrix(I, 15) & "'"
'Else
'   RsMst_RRN.Filter = "RRNNo = '" & MSHFlexGrid1.TextMatrix(I, 1) & "' And ISINID = '" & MSHFlexGrid1.TextMatrix(I, 3) & "'"
'End If

If Not RsMst_RRN.EOF Then
   Call Indata
   Frame0.Visible = False
End If

End Sub
Private Sub CmdISINID_Click()
    FrmMst_ISIN.Show
End Sub

'----
Private Sub CmdClientID_Click()
    FrmMst_Client.Show
End Sub
Private Sub CmbClientDP_LostFocus()

Dim Str_() As String
Str_ = Split(CmbClientDP.Text, "~")
If UBound(Str_) <> 1 Then
   MsgBox "Invalid selection !!!!! "
   CmbClientDP.SetFocus
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

End Sub


Private Sub ShowGridData(RSType As Variant, k As Integer)
For C = 0 To MSHFlexGrid1.Cols - 1
   MSHFlexGrid1.TextMatrix(k, C) = IIf(IsNull(RSType.Fields(C)), "", RSType.Fields(C))
Next
End Sub
Private Sub TxtClientID_LostFocus()
If SaveCmd.Enabled = False Then Exit Sub
CmbClientDP.Visible = False
If TxtClientID = "" Then Exit Sub
'If CmbExchangeTypeID.Text = "" Then Exit Sub

Dim ans As Variant
mExchangeTypeID = GetExchangeID("NCDEX")
ans = GetClientNameDtl(TxtClientID, mExchangeTypeID)

If ans(0, 0) = False Then
   MsgBox "Not Found !!!"
   TxtClientID.SetFocus
   Exit Sub
End If
If ans(3, 0) > 1 Then
   CmbClientDP.Visible = True
End If
CmbClientDP.Text = ""
CmbClientDP.Clear
tmp = "SELECT Mst_Depository.DepositoryID, Mst_Depository.DepositoryName, Mst_Client.ClientID, " & _
      "Mst_Client.ClientName,  Mst_Client.ClientIDRow " & _
      "FROM Mst_Depository INNER JOIN Mst_Client ON Mst_Depository.DepositoryID = Mst_Client.DepositoryID " & _
      "WHERE Mst_Depository.ExchangeTypeID = " & mExchangeTypeID & _
      "  And  (((Mst_Client.ExchangeTypeID)=" & mExchangeTypeID & ") AND ((Mst_Client.ClientID)='" & TxtClientID & "'));"
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
Else
   If CmbClientDP.Visible = True Then
      CmbClientDP.SetFocus
   End If
End If
'Call GetGridData(mClientIDRow, mExchangeTypeID, mCMPID, mCommodityID)

 
End Sub

Private Sub MSHFlexGrid2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyC And Shift = vbCtrlMask Then
   With MSHFlexGrid2
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
