VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmTxn_CashReceipt 
   Caption         =   "Cash Bank Receipt"
   ClientHeight    =   9675
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13125
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9675
   ScaleWidth      =   13125
   WindowState     =   2  'Maximized
   Begin VB.TextBox TxtG_UID 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
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
      Height          =   360
      Left            =   10920
      Locked          =   -1  'True
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   240
      Visible         =   0   'False
      Width           =   2400
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
      Height          =   1935
      Left            =   4200
      TabIndex        =   47
      Top             =   5280
      Visible         =   0   'False
      Width           =   4695
      Begin VB.CheckBox Check2 
         BackColor       =   &H008080FF&
         Caption         =   "Generate Excel"
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   1440
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton CmdCancel 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cancel"
         Height          =   400
         Left            =   2880
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton CmdSaveHologram 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Save"
         Height          =   400
         Left            =   1440
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   840
         Width           =   1440
      End
      Begin VB.TextBox TxtHologramNo 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1440
         MaxLength       =   25
         TabIndex        =   48
         Top             =   240
         Width           =   2880
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
         TabIndex        =   51
         Top             =   390
         Width           =   1260
      End
   End
   Begin VB.Frame Frame0 
      Height          =   1845
      Left            =   600
      TabIndex        =   0
      Top             =   6480
      Visible         =   0   'False
      Width           =   13680
      Begin VB.TextBox TxtSCRID 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1440
         MaxLength       =   15
         TabIndex        =   100
         TabStop         =   0   'False
         Top             =   720
         Width           =   1245
      End
      Begin VB.TextBox TxtAmountTo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   13920
         MaxLength       =   15
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   330
         Width           =   990
      End
      Begin VB.TextBox TxtAmountFrom 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   12600
         MaxLength       =   15
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   330
         Width           =   1005
      End
      Begin VB.ComboBox CmbSLocation 
         Height          =   315
         Left            =   2805
         Sorted          =   -1  'True
         TabIndex        =   63
         Top             =   330
         Width           =   3525
      End
      Begin VB.ComboBox CmbSInstrumentType 
         Height          =   315
         ItemData        =   "FrmTxn_CashReceipt.frx":0000
         Left            =   10290
         List            =   "FrmTxn_CashReceipt.frx":000D
         Sorted          =   -1  'True
         TabIndex        =   62
         Top             =   330
         Width           =   1845
      End
      Begin VB.ComboBox CmbSAgent 
         Height          =   315
         Left            =   6345
         Sorted          =   -1  'True
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   330
         Width           =   3885
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "Search"
         Height          =   350
         Left            =   12600
         TabIndex        =   60
         Top             =   750
         Width           =   1000
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1455
         Left            =   120
         TabIndex        =   29
         Top             =   1200
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   2566
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   350
         Left            =   13620
         TabIndex        =   28
         Top             =   750
         Width           =   1000
      End
      Begin MSComCtl2.DTPicker STo 
         Height          =   300
         Left            =   1440
         TabIndex        =   59
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
         Format          =   59834369
         CurrentDate     =   36494
      End
      Begin MSComCtl2.DTPicker SFrom 
         Height          =   300
         Left            =   120
         TabIndex        =   70
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
         Format          =   59834369
         CurrentDate     =   36494
      End
      Begin VB.Label Label19 
         Caption         =   "ID"
         Height          =   255
         Left            =   240
         TabIndex        =   101
         Top             =   750
         Width           =   405
      End
      Begin VB.Label Label11 
         Caption         =   "To"
         Height          =   255
         Left            =   13680
         TabIndex        =   73
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label10 
         Caption         =   "From"
         Height          =   255
         Left            =   12240
         TabIndex        =   72
         Top             =   360
         Width           =   405
      End
      Begin VB.Label Label44 
         Caption         =   "From Date"
         Height          =   255
         Left            =   360
         TabIndex        =   71
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Amount Range"
         Height          =   225
         Left            =   13200
         TabIndex        =   68
         Top             =   120
         Width           =   1080
      End
      Begin VB.Label Label36 
         Caption         =   "Location"
         Height          =   300
         Left            =   4207
         TabIndex        =   67
         Top             =   120
         Width           =   720
      End
      Begin VB.Label Label37 
         Caption         =   "Agent Name"
         Height          =   300
         Left            =   7822
         TabIndex        =   66
         Top             =   120
         Width           =   1080
      End
      Begin VB.Label Label38 
         Caption         =   "Instrument Type"
         Height          =   225
         Left            =   10612
         TabIndex        =   65
         Top             =   120
         Width           =   1200
      End
      Begin VB.Label Label45 
         Caption         =   "To Date"
         Height          =   255
         Left            =   1680
         TabIndex        =   64
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1095
      Left            =   120
      TabIndex        =   34
      Top             =   8685
      Width           =   14205
      Begin VB.CommandButton CmdPrintPreview 
         Caption         =   "Print Preview"
         Height          =   400
         Left            =   10800
         TabIndex        =   55
         Top             =   165
         Width           =   2535
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Print"
         Height          =   400
         Left            =   10800
         TabIndex        =   46
         Top             =   585
         Width           =   2535
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "Search"
         Height          =   400
         Left            =   8925
         TabIndex        =   19
         Top             =   165
         Width           =   1875
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   400
         Left            =   6810
         TabIndex        =   23
         Top             =   165
         Width           =   2115
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   400
         Left            =   4575
         TabIndex        =   22
         Top             =   165
         Width           =   2235
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   400
         Left            =   2340
         TabIndex        =   17
         Top             =   165
         Width           =   2235
      End
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New"
         Height          =   400
         Left            =   100
         TabIndex        =   21
         Top             =   165
         Width           =   2235
      End
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   400
         Left            =   8925
         TabIndex        =   18
         Top             =   585
         Width           =   1875
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   400
         Left            =   6810
         TabIndex        =   27
         Top             =   585
         Width           =   2115
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   400
         Left            =   4575
         TabIndex        =   26
         Top             =   585
         Width           =   2235
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   400
         Left            =   2340
         TabIndex        =   25
         Top             =   585
         Width           =   2235
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   400
         Left            =   100
         TabIndex        =   24
         Top             =   585
         Width           =   2235
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4200
      Left            =   120
      TabIndex        =   30
      Top             =   90
      Width           =   14205
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
         ItemData        =   "FrmTxn_CashReceipt.frx":0023
         Left            =   1995
         List            =   "FrmTxn_CashReceipt.frx":0025
         Sorted          =   -1  'True
         TabIndex        =   10
         Top             =   3765
         Width           =   3255
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
         Left            =   10320
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   2850
         Width           =   375
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
         Left            =   1995
         Sorted          =   -1  'True
         TabIndex        =   8
         Top             =   2850
         Width           =   8205
      End
      Begin VB.TextBox TxtCRID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
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
         Height          =   360
         Left            =   2010
         Locked          =   -1  'True
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   195
         Width           =   3120
      End
      Begin VB.TextBox TxtPreFix 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   360
         Left            =   11295
         Locked          =   -1  'True
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   1500
         Width           =   1635
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
         Left            =   1995
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   600
         Width           =   10980
      End
      Begin VB.CommandButton CmdLocationID 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   13050
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   660
         Width           =   375
      End
      Begin MSComCtl2.DTPicker TxtInstrumentDate 
         Height          =   375
         Left            =   8055
         TabIndex        =   5
         Top             =   1500
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   661
         _Version        =   393216
         Format          =   59834369
         CurrentDate     =   39884
      End
      Begin VB.TextBox TxtInstrumentAmount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   360
         Left            =   8040
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1065
         Width           =   2040
      End
      Begin VB.ComboBox CmbInstrumentType 
         Height          =   315
         ItemData        =   "FrmTxn_CashReceipt.frx":0027
         Left            =   1995
         List            =   "FrmTxn_CashReceipt.frx":0034
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   1065
         Width           =   3180
      End
      Begin VB.TextBox TxtInstrumentNo 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   2010
         MaxLength       =   50
         TabIndex        =   4
         Top             =   1500
         Width           =   3165
      End
      Begin VB.TextBox TxtRemark 
         Appearance      =   0  'Flat
         Height          =   435
         Left            =   1995
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   3255
         Width           =   10920
      End
      Begin VB.TextBox TxtDrawnOn 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1995
         MaxLength       =   150
         TabIndex        =   6
         Top             =   1980
         Width           =   10920
      End
      Begin VB.TextBox TxtBranchName 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1995
         MaxLength       =   150
         TabIndex        =   7
         Top             =   2415
         Width           =   10920
      End
      Begin VB.TextBox TxtLocationID 
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
         Left            =   1995
         Locked          =   -1  'True
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   637
         Width           =   10950
      End
      Begin MSComCtl2.DTPicker TxtCRDate 
         Height          =   375
         Left            =   8055
         TabIndex        =   1
         Top             =   195
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         Format          =   59834369
         CurrentDate     =   39884
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
         Left            =   2010
         ScrollBars      =   2  'Vertical
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   2850
         Width           =   8175
      End
      Begin VB.TextBox TxtExchangeTypeID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
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
         Left            =   1995
         Locked          =   -1  'True
         TabIndex        =   75
         Top             =   3765
         Width           =   3255
      End
      Begin VB.Label Label14 
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
         Left            =   240
         TabIndex        =   76
         Top             =   3818
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
         Left            =   240
         TabIndex        =   58
         Top             =   2903
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "Prefix"
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
         Left            =   10320
         TabIndex        =   45
         Top             =   1530
         Width           =   720
      End
      Begin VB.Label Label8 
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
         Height          =   300
         Left            =   240
         TabIndex        =   43
         Top             =   660
         Width           =   1440
      End
      Begin VB.Label Label7 
         Caption         =   "Instrument Date"
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
         Left            =   5955
         TabIndex        =   40
         Top             =   1537
         Width           =   1800
      End
      Begin VB.Label Label6 
         Caption         =   "Amount"
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
         Left            =   5940
         TabIndex        =   39
         Top             =   1095
         Width           =   1800
      End
      Begin VB.Label Label5 
         Caption         =   "Instrument No"
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
         TabIndex        =   38
         Top             =   1530
         Width           =   1800
      End
      Begin VB.Label Label4 
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
         TabIndex        =   37
         Top             =   3322
         Width           =   1800
      End
      Begin VB.Label Label2 
         Caption         =   "Instrument Type"
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
         TabIndex        =   36
         Top             =   1065
         Width           =   1800
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   2468
         Width           =   1800
      End
      Begin VB.Label LblFlag 
         Caption         =   "Receipt Date"
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
         Left            =   5955
         TabIndex        =   33
         Top             =   232
         Width           =   1800
      End
      Begin VB.Label LblStateID 
         Caption         =   "Drawn On"
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
         TabIndex        =   32
         Top             =   1965
         Width           =   1800
      End
      Begin VB.Label LblLocationID 
         Caption         =   "Receipt ID"
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
         Left            =   240
         TabIndex        =   31
         Top             =   225
         Width           =   1200
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4350
      Left            =   120
      TabIndex        =   77
      Top             =   4320
      Width           =   13965
      _ExtentX        =   24633
      _ExtentY        =   7673
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "PI Detials"
      TabPicture(0)   =   "FrmTxn_CashReceipt.frx":004A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "MSHFlexGrid3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Pic_TTD"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Log History"
      TabPicture(1)   =   "FrmTxn_CashReceipt.frx":0066
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "TxtCreated"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "TxtUpdated"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label13"
      Tab(1).Control(3)=   "Label12"
      Tab(1).ControlCount=   4
      Begin VB.PictureBox Pic_TTD 
         Height          =   1455
         Left            =   120
         ScaleHeight     =   1395
         ScaleMode       =   0  'User
         ScaleWidth      =   9868.087
         TabIndex        =   80
         Top             =   360
         Width           =   13740
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
            TabIndex        =   81
            Top             =   600
            Visible         =   0   'False
            Width           =   495
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
            TabIndex        =   87
            TabStop         =   0   'False
            Top             =   960
            Width           =   1215
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
            TabIndex        =   86
            TabStop         =   0   'False
            Top             =   960
            Width           =   3855
         End
         Begin VB.TextBox TxtPINo 
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
            Left            =   90
            MaxLength       =   50
            TabIndex        =   11
            Top             =   255
            Width           =   1815
         End
         Begin VB.CommandButton CmdDeleteList 
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
            Left            =   11940
            TabIndex        =   85
            Top             =   915
            Width           =   1710
         End
         Begin VB.CommandButton CmdAdd2List 
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
            Left            =   10380
            TabIndex        =   16
            Top             =   915
            Width           =   1575
         End
         Begin VB.TextBox TxtInvoiceDetails 
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
            Left            =   5745
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   14
            Top             =   255
            Width           =   7800
         End
         Begin VB.TextBox TxtAmount 
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
            Left            =   3660
            TabIndex        =   13
            Top             =   255
            Width           =   2040
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
            Left            =   1245
            TabIndex        =   84
            TabStop         =   0   'False
            Top             =   675
            Width           =   375
         End
         Begin VB.ComboBox CmbClientDP 
            Height          =   315
            Left            =   5160
            Sorted          =   -1  'True
            TabIndex        =   83
            Top             =   960
            Visible         =   0   'False
            Width           =   5055
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
            TabIndex        =   15
            Top             =   960
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
            TabIndex        =   82
            TabStop         =   0   'False
            Top             =   960
            Width           =   3495
         End
         Begin MSComCtl2.DTPicker TxtPIDate 
            Height          =   345
            Left            =   1950
            TabIndex        =   12
            Top             =   255
            Width           =   1680
            _ExtentX        =   2963
            _ExtentY        =   609
            _Version        =   393216
            Format          =   59834369
            CurrentDate     =   39884
         End
         Begin VB.Label LblPINo 
            Caption         =   "PI No."
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
            Left            =   652
            TabIndex        =   96
            Top             =   30
            Width           =   690
         End
         Begin VB.Label LblInvoiceDetails 
            Caption         =   "Invoice Details"
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
            Left            =   8858
            TabIndex        =   95
            Top             =   30
            Width           =   1575
         End
         Begin VB.Label LblInvoiceAmount 
            Caption         =   "Invoice Amount"
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
            Left            =   3833
            TabIndex        =   94
            Top             =   30
            Width           =   1695
         End
         Begin VB.Label LblPIDate 
            Caption         =   "PI Date"
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
            Left            =   2363
            TabIndex        =   93
            Top             =   30
            Width           =   855
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
            TabIndex        =   92
            Top             =   660
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label15 
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
            TabIndex        =   91
            Top             =   660
            Width           =   1095
         End
         Begin VB.Label Label16 
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
            TabIndex        =   90
            Top             =   660
            Width           =   735
         End
         Begin VB.Label Label17 
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
            TabIndex        =   89
            Top             =   660
            Width           =   1095
         End
         Begin VB.Label Label18 
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
            TabIndex        =   88
            Top             =   660
            Width           =   975
         End
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
         Left            =   -73575
         Locked          =   -1  'True
         TabIndex        =   79
         TabStop         =   0   'False
         Top             =   840
         Width           =   8970
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
         Left            =   -73575
         Locked          =   -1  'True
         TabIndex        =   78
         TabStop         =   0   'False
         Top             =   1425
         Width           =   8970
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
         Height          =   2370
         Left            =   120
         TabIndex        =   99
         Top             =   1845
         Width           =   13515
         _ExtentX        =   23839
         _ExtentY        =   4180
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label13 
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
         Left            =   -74520
         TabIndex        =   98
         Top             =   900
         Width           =   1695
      End
      Begin VB.Label Label12 
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
         Left            =   -74520
         TabIndex        =   97
         Top             =   1485
         Width           =   975
      End
   End
End
Attribute VB_Name = "FrmTxn_CashReceipt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg As Variant
Dim mClientIDRow, mLocationID As Double
Dim mAgentID As Double
Dim mExchangeTypeID As Double

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
   mExchangeTypeID = 0
   CmbNonNcdexClient.Clear
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
   CmbNonNcdexClient.Clear
   Call Grid_Head1
End If
mExchangeTypeID = RsMst_ExchangeType!ExchangeTypeID

If mExchangeTypeID = 1 Then
   CmbNonNcdexClient.Width = 7300
   CmbNonNcdexClient.Top = TxtClientID.Top
   CmbNonNcdexClient.Left = TxtClientID.Left
   Label15.Visible = False
   Label16.Visible = False
   Label18.Visible = False
   CmbNonNcdexClient.Visible = True
Else
   Label15.Visible = True
   Label16.Visible = True
   Label18.Visible = True
   CmbNonNcdexClient.Visible = False
End If

Call ClearInvoiceDetails

End Sub

Private Sub CmbInstrumentType_LostFocus()
If CmbInstrumentType.Text = "" Then Exit Sub
CmbInstrumentType.Text = Trim(CmbInstrumentType.Text)

If LCase(CmbInstrumentType.Text) = "cash" Then
   TxtDrawnOn.Enabled = False
   TxtBranchName.Enabled = False
   TxtInstrumentNo.Enabled = False
   TxtInstrumentDate.Enabled = False
   TxtDrawnOn = ""
   TxtBranchName = ""
   TxtInstrumentNo = ""
'   TxtInstrumentDate = ""
   Exit Sub
End If
If LCase(CmbInstrumentType.Text) = "cheque" Or LCase(CmbInstrumentType.Text) = "dd" Then
   TxtDrawnOn.Enabled = True
   TxtBranchName.Enabled = True
   TxtInstrumentNo.Enabled = True
   TxtInstrumentDate.Enabled = True
   Exit Sub
End If
MsgBox "Invalid selection !!!!"
CmbInstrumentType.SetFocus

End Sub

Private Sub CmbLocationID_GotFocus()
tmp = CmbLocationID.Text
Call TableMst_Location(" Where LocationId > 0 ")
CmbLocationID.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_Location.RecordCount
    CmbLocationID.AddItem RsMst_Location!LocationName
    RsMst_Location.MoveNext
Next
CmbLocationID.Text = tmp
End Sub
'----
Private Sub CmbLocationID_LostFocus()
If CmbLocationID.Text = "" Then
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocationID.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection "
   CmbLocationID.SetFocus
   Exit Sub
End If
mLocationID = RsMst_Location!LocationID
TxtPreFix = GetPreFixByLocation(RsMst_Location!LocationID, "P")
End Sub

Private Sub CmbNonNcdexClient_GotFocus()
If CmbExchangeTypeID.Text = "" Then
   MsgBox "Please select Exchange Type!!!"
   CmbNonNcdexClient.Clear
   CmbExchangeTypeID.SetFocus
   Exit Sub
End If
tmp = CmbNonNcdexClient.Text
Call TableMst_Client("where ExchangeTypeId = " & mExchangeTypeID & " ")
CmbNonNcdexClient.Clear
If RsMst_Client.RecordCount = 0 Then
   MsgBox "No record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_Client.RecordCount
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
mClientIDRow = RsMst_Client!ClientIDRow

Dim ans As Variant

ans = GetClientNameByRow(RsMst_Client!ClientIDRow)
TxtClientID = ans(0)
TxtClientName = ans(1)
TxtDPID = ans(2)
TxtDPName = ans(3)


End Sub

Private Sub CmbSAgent_GotFocus()
tmp = CmbSAgent.Text
Call TableMst_Agent
CmbSAgent.Clear
If RsMst_Agent.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_Agent.RecordCount
    CmbSAgent.AddItem RsMst_Agent!AgentName
    RsMst_Agent.MoveNext
Next
CmbSAgent.Text = tmp
End Sub

Private Sub CmbSAgent_LostFocus()
If CmbSAgent.Text = "" Then
   Exit Sub
End If
RsMst_Agent.MoveFirst
RsMst_Agent.Find "agentName = '" & CmbSAgent.Text & "'"
If RsMst_Agent.EOF Then
   MsgBox "Invalid selection "
   CmbSAgent.SetFocus
   Exit Sub
End If
'mAgentID = RsMst_Agent!AgentID
End Sub

Private Sub CmbSInstrumentType_LostFocus()
If CmbSInstrumentType.Text = "" Then Exit Sub
If LCase(CmbSInstrumentType.Text) = "cash" Then
ElseIf LCase(CmbSInstrumentType.Text) = "cheque" Then
ElseIf LCase(CmbSInstrumentType.Text) = "dd" Then
Else
   MsgBox "Invalid Selection"
   CmbSInstrumentType.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmbSLocation_GotFocus()
tmp = CmbSLocation.Text
Call TableMst_Location(" Where LocationId > 0 ")
CmbSLocation.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_Location.RecordCount
    CmbSLocation.AddItem RsMst_Location!LocationName
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
'mLocationId = RsMst_Location!LocationID
End Sub

Private Sub CmdCancel_Click()
TxtHologramNo = ""
Frame3.Visible = False
End Sub

Private Sub CmdClient_Click()
FrmMst_Client.Show
End Sub

Private Sub CmdLocationID_Click()
FrmMst_Location.Show
End Sub

Private Sub CmdPrint_Click()
If RsTxn_CashReceipt!HologramNo = "" Or IsNull(RsTxn_CashReceipt!HologramNo) Then
   If LCase(UserType) = "power" Then
      Check2.Visible = True
   End If

   Frame3.Visible = True
   TxtHologramNo.SetFocus
   Call TableMst_Hologram
   Exit Sub
Else
   TxtHologramNo = ""
   Frame3.Visible = False
End If
RsTxn_CashReceipt!Flag = "P"
RsTxn_CashReceipt.Update
Gen_mTimeStamp = GetTimeStamp
Call Create_Xls(Gen_mTimeStamp & Gen_UserName & "_CashReceipt-" & Trim(TxtPreFix) & "" & Trim(TxtCRID) & ".xls")
Call Xls_Print_Rpt
End Sub

Private Sub CmdPrintPreview_Click()
Gen_mTimeStamp = GetTimeStamp
Call Create_Xls(Gen_mTimeStamp & Gen_UserName & "_CashReceipt-" & Trim(TxtPreFix) & "" & Trim(TxtCRID) & ".xls")
Call Xls_Print_Rpt
End Sub
Private Sub CmdSearch_Click()

If SFrom.Value > STo.Value Then
   MsgBox "'From Date' must be less than 'To date'!!!"
   SFrom.SetFocus
   Exit Sub
End If

If Val(TxtAmountFrom) > Val(TxtAmountTo) Then
   MsgBox "In Amount range 'From amount' must be less than 'To amount'!!!"
   TxtAmountFrom.SetFocus
   Exit Sub
End If

tmp = "Select CR.CRID,LM.SM_Prefix,CR.CRDate,CR.ClientIDRow,CR.InstrumentType,CR.DrawnOn,CR.BranchName,CR.InstrumentNo,CR.InstrumentDate,"
tmp = tmp & " CR.InstrumentAmount , AM.AgentName, CR.Remark,ME.ExchangeType, CR.Flag, CR.HologramNo"
tmp = tmp & " From Txn_CashReceipt CR"
tmp = tmp & " Inner Join Mst_Agent AM On CR.AgentID = AM.AgentID"
tmp = tmp & " Inner Join Mst_Location LM On CR.LocationID = LM.LocationID"
tmp = tmp & " Inner Join Mst_ExchangeType ME On CR.ExchangeTypeID = ME.ExchangeTypeID"


'tmp = tmp & " Where CR.CRDate between '2009-08-20 00:00:00' and '2009-08-20 00:00:00'"
If Gen_DBType = "MDB" Then
   tmp = tmp & " Where CR.CRDate Between #" & Format(SFrom, Gen_DatFormat) & "# And #" & Format(STo, Gen_DatFormat) & "#) "
Else
    tmp = tmp & " where CR.CRDate Between '" & Format(SFrom, Gen_DatFormat) & "' And '" & Format(STo, Gen_DatFormat) & "' "
End If

If TxtAmountFrom <> "" And TxtAmountTo <> "" Then
   tmp = tmp & " AND CR.InstrumentAmount between " & Val(TxtAmountFrom) & " And " & Val(TxtAmountTo) & " "
End If

If CmbSLocation.Text <> "" Then
   tmp = tmp & " AND LM.LocationName = '" & CmbSLocation.Text & "' "
End If
If CmbSAgent <> "" Then
   tmp = tmp & " AND AM.AgentName = '" & CmbSAgent.Text & "' "
End If
If CmbSInstrumentType.Text <> "" Then
   tmp = tmp & " AND CR.InstrumentType = '" & CmbSInstrumentType.Text & "' "
End If
If TxtSCRID <> "" Then
   tmp = tmp & " And CR.CRID = " & Val(TxtSCRID)
End If

Call TmpTable
Call Grid_Head
For iu = 1 To TmpRs.RecordCount
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
    For c = 0 To MSHFlexGrid1.Cols - 1
       MSHFlexGrid1.TextMatrix(iu, c) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
    Next
    If MSHFlexGrid1.TextMatrix(iu, 13) = "0" Then
       MSHFlexGrid1.TextMatrix(iu, 13) = ""
    Else
       MSHFlexGrid1.TextMatrix(iu, 13) = "P"
    End If
    TmpRs.MoveNext
Next

End Sub

Private Sub TxtAmount_LostFocus()
If TxtAmount = "" Then Exit Sub
If IsNumeric(TxtAmount) = False Or Val(TxtAmount) < 0 Then
   MsgBox "Invalid numeric format!!!"
   TxtAmount.SetFocus
   Exit Sub
End If
End Sub

Private Sub TxtAmountFrom_LostFocus()
If TxtAmountFrom = "" Then Exit Sub
If IsNumeric(TxtAmountFrom) = False Or Val(TxtAmountFrom) < 0 Then
   MsgBox "Invalid numeric format!!!"
   TxtAmountFrom.SetFocus
   Exit Sub
End If
End Sub

Private Sub TxtAmountTo_LostFocus()
If TxtAmountTo = "" Then Exit Sub
If IsNumeric(TxtAmountTo) = False Or Val(TxtAmountTo) < 0 Then
   MsgBox "Invalid numeric format!!!"
   TxtAmountTo.SetFocus
   Exit Sub
End If
End Sub

Private Sub TxtCRDate_LostFocus()
If TxtCRDate.Value > Date Then
   MsgBox "Future Cash Receipt Date is not allowed !!!! "
   TxtCRDate.SetFocus
End If

End Sub

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
ans = MsgBox("Are you sure you want to save this Hologram No??", vbYesNo)
If ans = vbYes Then
   Call TableMst_Hologram
   RsMst_Hologram.AddNew
   RsMst_Hologram!HologramNo = TxtHologramNo
   RsMst_Hologram!DocType = "Cash Receipt"
   RsMst_Hologram!DocNo = TxtCRID
   RsMst_Hologram!SM_Prefix = TxtPreFix
   RsMst_Hologram.Update
   
   'RsTxn_CashReceipt.Filter = " CRID = " & TxtCRID & " And SM_Prefix = '" & TxtPreFix & "' "
   RsTxn_CashReceipt!HologramNo = TxtHologramNo
   RsTxn_CashReceipt!Flag = "P"
   RsTxn_CashReceipt.Update
   'RsTxn_CashReceipt.Filter = ""
   TxtHologramNo = ""
   Frame3.Visible = False
   Gen_mTimeStamp = GetTimeStamp
   'Call Create_Xls(Gen_mTimeStamp & Gen_UserName & "_CDTF-" & Trim(TxtPreFix) & "" & Trim(TxtCxTFNo) & ".xls")
    If Check2.Value = 1 Then
       Call Create_Xls(Gen_mTimeStamp & Gen_UserName & "_CashReceipt-" & Trim(TxtPreFix) & "" & Trim(TxtCRID) & ".xls")
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
oBook.SaveAs (Pat & "\Excel\" & FileName)

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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "_CashReceipt-" & Trim(TxtPreFix) & "" & Trim(TxtCRID) & ".xls")

i = 2
Set oWS = oWB.Worksheets("Sheet" & 1)
''oWS.PageSetup.RightMargin = 0.05
'oWS.PageSetup.TopMargin = 0.05
'oWS.PageSetup.RightMargin = 0.05
'oWS.PageSetup.TopMargin = 4.6
oWS.Name = TxtPreFix & "-" & TxtCRID  'MSHFlexGrid2.TextMatrix(i, 23)
oWS.ClearArrows
mRow = 8

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
oWS.Cells(mRow, c + 1) = "NATIONAL COLLATERAL MANAGEMENT SERVICES LIMITED"
oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).Font.Size = 10

'oWS.Cells(1, 1).Font.ColorIndex = 4
mRow = mRow + 1

tmp = "a" & mRow & ":" & "g" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = "Location -" & TxtLocationID
oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).Font.Size = 10

mRow = mRow + 1

oWS.Cells(mRow, c + 1) = "Receipt No"
oWS.Cells(mRow, c + 1).Font.Size = 7
oWS.Cells(mRow, c + 1).Font.Bold = True

tmp = "b" & mRow & ":" & "c" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 2) = TxtPreFix & "-" & TxtCRID
oWS.Cells(mRow, c + 2).Font.Size = 7
oWS.Cells(mRow, c + 2).HorizontalAlignment = -4108

oWS.Cells(mRow, c + 4) = "Receipt Date"
oWS.Cells(mRow, c + 4).Font.Size = 7
oWS.Cells(mRow, c + 4).Font.Bold = True

tmp = "e" & mRow & ":" & "g" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 5) = " " & Format(TxtCRDate, "dd-mmm-yyyy")
oWS.Cells(mRow, c + 5).Font.Size = 7
oWS.Cells(mRow, c + 5).HorizontalAlignment = -4108


mRow = mRow + 1

Dim Rsp As String
Rsp = NumStrConv(Val(TxtInstrumentAmount))
tmp = "a" & mRow & ":" & "g" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = "Received with thanks from Shri " & TxtAgent & " a sum of Rs " & TxtInstrumentAmount & " ( Rupees " & Rsp & " )" & _
                         " towards the following, by " & CmbInstrumentType & " "
'No:................. (instrument no) dated ................. ( instrument date) drawn on ............................ (drawn on Bank) ........................ branch (drawn on branch)."
If CmbInstrumentType.Text <> "Cash" Then
   oWS.Cells(mRow, c + 1) = oWS.Cells(mRow, c + 1) & " No: " & TxtInstrumentNo & " dated " & TxtInstrumentDate & " drawn on " & TxtDrawnOn & " " & TxtBranchName & " branch (subject to realisation). "
End If
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).WrapText = True
oWS.Cells(mRow, c + 1).HorizontalAlignment = 2
oWS.Cells(mRow, c + 1).Font.Size = 10
oWS.Cells(mRow, c + 1).RowHeight = 60

mRow = mRow + 1
'mCol = 5

tmp = "a" & mRow & ":" & "g" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = "Cash Receipt Details"
'oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).Font.Size = 10

mRow = mRow + 1
'mCol = 5


''New
oWS.Cells(mRow, c + 1) = "Proforma Invoice No"
oWS.Cells(mRow, c + 1).Font.Size = 7
oWS.Cells(mRow, c + 1).Font.Bold = True

oWS.Cells(mRow, c + 2) = "Proforma Invoice Date"
oWS.Cells(mRow, c + 2).Font.Size = 7
oWS.Cells(mRow, c + 2).Font.Bold = True

oWS.Cells(mRow, c + 3) = "Receipt Amount"
oWS.Cells(mRow, c + 3).Font.Size = 7
oWS.Cells(mRow, c + 3).Font.Bold = True

oWS.Cells(mRow, c + 4) = "Client Name"
oWS.Cells(mRow, c + 4).Font.Size = 7
oWS.Cells(mRow, c + 4).Font.Bold = True
tmp = "d" & mRow & ":" & "e" & mRow
oWS.Range(tmp).Merge

oWS.Cells(mRow, c + 6) = "Invoice Details"
oWS.Cells(mRow, c + 6).Font.Size = 7
oWS.Cells(mRow, c + 6).Font.Bold = True
tmp = "f" & mRow & ":" & "g" & mRow
oWS.Range(tmp).Merge

mRow = mRow + 1

For k = 1 To MSHFlexGrid3.Rows - 2
'k = 1
    oWS.Cells(mRow, c + 1) = MSHFlexGrid3.TextMatrix(k, 0)
    oWS.Cells(mRow, c + 1).Font.Size = 7
    
    oWS.Cells(mRow, c + 2) = " " & Format(MSHFlexGrid3.TextMatrix(k, 1), "dd-mmm-yyyy")
    oWS.Cells(mRow, c + 2).Font.Size = 7
    
    oWS.Cells(mRow, c + 3) = MSHFlexGrid3.TextMatrix(k, 2)
    oWS.Cells(mRow, c + 3).Font.Size = 7
    
    oWS.Cells(mRow, c + 4) = MSHFlexGrid3.TextMatrix(k, 4)
    oWS.Cells(mRow, c + 4).Font.Size = 7
    tmp = "d" & mRow & ":" & "e" & mRow
    oWS.Range(tmp).Merge
    If Len(MSHFlexGrid3.TextMatrix(k, 3)) > 10 Then
'        oWS.Cells(mRow, c + 3).RowHeight = 20
    End If
    
    oWS.Cells(mRow, c + 6) = MSHFlexGrid3.TextMatrix(k, 7)
    oWS.Cells(mRow, c + 6).Font.Size = 7
    tmp = "f" & mRow & ":" & "g" & mRow
    oWS.Range(tmp).Merge
    If Len(MSHFlexGrid3.TextMatrix(k, 7)) > 10 Then
        oWS.Cells(mRow, c + 7).RowHeight = Len(MSHFlexGrid3.TextMatrix(k, 7)) / 2
    End If
    
    mRow = mRow + 1
Next k

'tmp = "a" & mRow & ":" & "a" & mRow + 1
'oWS.Range(tmp).Merge

tmp = "a" & mRow & ":" & "d" & mRow
oWS.Range(tmp).Merge

tmp = "e" & mRow & ":" & "g" & mRow
oWS.Range(tmp).Merge

oWS.Cells(mRow, c + 5) = "Affix Hologram here"
oWS.Cells(mRow, c + 5).Font.Size = 7
oWS.Cells(mRow, c + 5).Font.Bold = True
oWS.Cells(mRow, c + 5).RowHeight = 45

mRow = mRow + 1

tmp = "a" & mRow & ":" & "d" & mRow
oWS.Range(tmp).Merge

oWS.Cells(mRow, c + 1) = "Name and Signature of authorised representative of NCMSL"
oWS.Cells(mRow, c + 1).Font.Size = 7
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).WrapText = True

tmp = "e" & mRow & ":" & "g" & mRow
oWS.Range(tmp).Merge

If IsNull(RsTxn_CashReceipt!HologramNo) Then
   oWS.Cells(mRow, c + 5) = "Hologram No:- Preview"
Else
  oWS.Cells(mRow, c + 5) = "Hologram No:- " & RsTxn_CashReceipt!HologramNo
End If
oWS.Cells(mRow, c + 5).Font.Size = 7
oWS.Cells(mRow, c + 5).Font.Bold = True
oWS.Cells(mRow, c + 5).RowHeight = 40

mRow = mRow + 1
tmp = "a" & mRow & ":" & "g" & mRow
oWS.Range(tmp).Merge

k = mRow

For Cell = 8 To k
    tmp = "a" & Cell & ":" & "g" & Cell
    oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
    oWS.Range(tmp).Font.Name = "Zurich BT"
    'oWS.Range(tmp).Font.Size = 7
    oWS.Range(tmp).WrapText = True
Next Cell
If Gen_ExcelPassword <> "" Then
   oWS.Protect Gen_ExcelPassword, , , , , , True, True
End If
oWB.Save

Call Log_Print_History(TxtPreFix, "CR", TxtCRID, Gen_FLName)

MsgBox ("Excel file created !!! ")
oXL.Visible = True

End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Call GButtonLockAD(Me)
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If
Call Grid_Head1

Call TableTxn_CashReceipt(" Order by SM_Prefix,CRID")
If RsTxn_CashReceipt.RecordCount = 0 Then
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
Private Sub CmdExcel_Click()
Call Gen_Create_Xls("Txn_CashReceipt.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\Txn_CashReceipt.xls")
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

Private Sub TxtInstrumentAmount_LostFocus()
If TxtInstrumentAmount = "" Then Exit Sub
If Not IsNumeric(TxtInstrumentAmount) Then
   MsgBox "Invalid numeric format!!! "
   TxtInstrumentAmount.SetFocus
   Exit Sub
End If
End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsTxn_CashReceipt.MoveNext
If RsTxn_CashReceipt.EOF Then
   RsTxn_CashReceipt.MoveLast
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
RsTxn_CashReceipt.MovePrevious
If RsTxn_CashReceipt.BOF Then
   RsTxn_CashReceipt.MoveFirst
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
RsTxn_CashReceipt.MoveFirst
Call InData
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsTxn_CashReceipt.MoveLast
Call InData
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
   If RsTxn_CashReceipt.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   RsTxn_CashReceipt.MoveFirst
 
   RsTxn_CashReceipt.Find "G_UID = '" & TxtG_UID & "'"
   If RsTxn_CashReceipt.EOF Then
      RsTxn_CashReceipt.MoveLast
   End If
   Call InData
   Exit Sub
End If
Frame0.Height = Me.Height - 600
Frame0.Left = Frame1.Left - 100
Frame0.Top = Frame1.Top - 100
Frame0.Width = Me.Width - 150
Frame0.Visible = True
MSHFlexGrid1.Clear
MSHFlexGrid1.Cols = 11
MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 1400
RsTxn_CashReceipt.Filter = ""
RsTxn_CashReceipt.MoveFirst
Call Grid_Head

SFrom.Value = Date - 730
STo.Value = Date
CmbSLocation.Text = ""
CmbSAgent.Text = ""
CmbSInstrumentType.Text = ""
TxtAmountFrom = ""
TxtAmountTo = ""
End Sub
Public Sub InData()
Frame1.Enabled = False
Pic_TTD.Enabled = False
CmdClient.Visible = False
CmbLocationID.Visible = False
CmdLocationID.Visible = False
CmbAgent.Visible = False
CmdAgent.Visible = False
TxtAgent.Visible = True
TxtLocationID.Visible = True
CmbLocationID.Text = ""

TxtCRID = RsTxn_CashReceipt!CRID
TxtCRDate = RsTxn_CashReceipt!CRDate
mClientIDRow = RsTxn_CashReceipt!ClientIDRow
TxtClient = GetClientName(RsTxn_CashReceipt!ClientIDRow)
CmbInstrumentType = RsTxn_CashReceipt!InstrumentType
TxtDrawnOn = RsTxn_CashReceipt!DrawnOn
TxtBranchName = RsTxn_CashReceipt!BranchName
TxtInstrumentNo = RsTxn_CashReceipt!InstrumentNo
TxtInstrumentDate = RsTxn_CashReceipt!InstrumentDate
TxtInstrumentAmount = RsTxn_CashReceipt!InstrumentAmount
TxtRemark = RsTxn_CashReceipt!Remark
TxtInvoiceDetail = RsTxn_CashReceipt!InvoiceDetail
TxtFlag = IIf(IsNull(RsTxn_CashReceipt!Flag), "", RsTxn_CashReceipt!Flag)
mLocationID = IIf(IsNull(RsTxn_CashReceipt!LocationID), 0, RsTxn_CashReceipt!LocationID)
TxtLocationID = GetLocationName(mLocationID)
TxtPreFix = RsTxn_CashReceipt!SM_Prefix
TxtG_UID = RsTxn_CashReceipt!G_UID
mAgentID = IIf(IsNull(RsTxn_CashReceipt!AgentID), 0, RsTxn_CashReceipt!AgentID)
If mAgentID <> 0 Then
   TxtAgent = GetAgentName(mAgentID)
Else
   TxtAgent = ""
End If

mExchangeTypeID = IIf(IsNull(RsTxn_CashReceipt!ExchangeTypeID), 0, RsTxn_CashReceipt!ExchangeTypeID)


CmbExchangeTypeID.Visible = False
TxtExchangeTypeID.Visible = True


If mExchangeTypeID <> 0 Then
   TxtExchangeTypeID = GetExchangeName(mExchangeTypeID)
Else
   TxtExchangeTypeID = ""
End If

TxtCreated = IIf(IsNull(RsTxn_CashReceipt!CreatedBy), "", RsTxn_CashReceipt!CreatedBy) & " :- " & IIf(IsNull(RsTxn_CashReceipt!CreatedDate), "", RsTxn_CashReceipt!CreatedDate)
TxtUpdated = IIf(IsNull(RsTxn_CashReceipt!UpdatedBy), "", RsTxn_CashReceipt!UpdatedBy) & " :- " & IIf(IsNull(RsTxn_CashReceipt!UpdatedDate), "", RsTxn_CashReceipt!UpdatedDate)
SSTab1.Tab = 0

Call InDataCRInvoiceDetails
If RsTxn_CashReceipt.RecordCount = 1 Then
   Call GButtonLock_1(Me, True)
Else
   Call GButtonLock(Me, True)
End If
CmdPrint.Enabled = True
End Sub
Private Sub InDataCRInvoiceDetails()
Call Grid_Head1
Call TableTxn_CashReceiptInvoiceDetails(" Where CRID= " & TxtCRID & " AND SM_Prefix = '" & TxtPreFix & "'")
If RsTxn_CashReceiptInvoiceDetails.RecordCount > 0 Then
   MSHFlexGrid3.Rows = MSHFlexGrid3.Rows + RsTxn_CashReceiptInvoiceDetails.RecordCount
   For i = 1 To MSHFlexGrid3.Rows - 2
       MSHFlexGrid3.TextMatrix(i, 0) = RsTxn_CashReceiptInvoiceDetails!PI_No
       MSHFlexGrid3.TextMatrix(i, 1) = RsTxn_CashReceiptInvoiceDetails!PI_Date
       MSHFlexGrid3.TextMatrix(i, 2) = RsTxn_CashReceiptInvoiceDetails!Amount
       ans = GetClientNameByRow(RsTxn_CashReceiptInvoiceDetails!ClientIDRow)
       MSHFlexGrid3.TextMatrix(i, 3) = ans(0)
       MSHFlexGrid3.TextMatrix(i, 4) = ans(1) 'TmpRs!ClientName
       MSHFlexGrid3.TextMatrix(i, 5) = ans(2) 'TmpRs!DepositoryID
       MSHFlexGrid3.TextMatrix(i, 6) = ans(3) 'TmpRs!DepositoryName
'        MSHFlexGrid3.TextMatrix(i, 3) = GetClientName(RsTxn_CashReceiptInvoiceDetails!ClientIDRow)
       MSHFlexGrid3.TextMatrix(i, 7) = RsTxn_CashReceiptInvoiceDetails!InvoiceDetail
       MSHFlexGrid3.TextMatrix(i, 8) = RsTxn_CashReceiptInvoiceDetails!ClientIDRow
       RsTxn_CashReceiptInvoiceDetails.MoveNext
   Next
End If
End Sub

Private Sub TxtLocationID_LostFocus()
'If TxtLocationID = "" Then Exit Sub
'If IsNumeric(TxtLocationID) = False Then
'   MsgBox "Invalid numeric format !!!!"
'   TxtLocationID.SetFocus
'   Exit Sub
'End If
'TxtLocationID = Val(TxtLocationID)
End Sub
Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)
Frame1.Enabled = True
Pic_TTD.Enabled = True
Call Grid_Head1
TxtCRDate.Enabled = True
TxtCRDate.Enabled = True
TxtInstrumentDate.Enabled = True
CmbInstrumentType.Visible = True
'CmbNonNcdexClient.Visible = True
CmdClient.Visible = True
CmbLocationID.Visible = True
CmdLocationID.Visible = True
TxtLocationID.Visible = False
CmbLocationID.Text = ""
CmbAgent.Visible = True
CmbAgent.Text = ""
CmdAgent.Visible = True
TxtAgent.Visible = False
TxtCRID.Locked = True
TxtInstrumentNo.Locked = False
'TxtInstrumentAmount.Locked = False
TxtDrawnOn.Locked = False
TxtBranchName.Locked = False
'TxtInvoiceDetail.Locked = False
TxtRemark.Locked = False
CmbExchangeTypeID.Visible = True
CmbExchangeTypeID.Text = ""
TxtExchangeTypeID.Visible = False

SSTab1.Tab = 0

TxtCRID = ""
TxtInstrumentNo = ""
TxtCRDate.Value = Date
TxtInstrumentDate = Date
CmbInstrumentType.Text = ""
CmbNonNcdexClient.Text = ""
TxtInstrumentAmount = ""
Call ClearInvoiceDetails
TxtDrawnOn = ""
TxtBranchName = ""
TxtInvoiceDetail = ""
TxtRemark = ""
TxtCreated = ""
TxtUpdated = ""

TxtPIDate.Value = Date
CmdPrint.Enabled = False
mExchangeTypeID = 0

If RsTxn_CashReceipt.RecordCount > 0 Then
    TxtCRDate.SetFocus
End If
End Sub
Private Sub EditCmd_Click()
If RsTxn_CashReceipt!Flag = "P" Then
   MsgBox "Print out is generated already. You can't edit this receipt!!!! "
   Exit Sub
End If

Edit_Flg = "E"

Call GButtonLock(Me, False)
Frame1.Enabled = False
Pic_TTD.Enabled = True
CmdClient.Visible = True
'CmbNonNcdexClient.Visible = False
'TxtClient.Visible = True
'CmbNonNcdexClient.Text = TxtClient
CmbLocationID.Text = TxtLocationID
CmbAgent.Text = TxtAgent


CmbExchangeTypeID.Text = TxtExchangeTypeID


CmdPrint.Enabled = False
End Sub
Private Sub DeleteCmd_Click()
If RsTxn_CashReceipt!Flag = "P" Then
   MsgBox "Print out is generated already. You can't delete this receipt!!!! "
   Exit Sub
End If

Dim ans As Variant
On Error GoTo msgError
ans = MsgBox("Do you really want to DELETE this Record??", vbYesNo)
If ans = vbYes Then
   Call DeleteCRInvoiceDetails
   RsTxn_CashReceipt.Delete
   RsTxn_CashReceipt.Update
   If RsTxn_CashReceipt.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsTxn_CashReceipt.MoveNext
   If RsTxn_CashReceipt.EOF() Then
      RsTxn_CashReceipt.MoveLast
   End If
   Call InData
End If
Exit Sub
msgError:
MsgBox "Child record Present "
RsTxn_CashReceipt.CancelUpdate
End Sub
Private Sub SaveCmd_Click()

If CmbLocationID.Text = "" Then
   MsgBox "Blank Location Name not allowed !!! "
   CmbLocationID.SetFocus
   Exit Sub
End If


If CmbInstrumentType.Text = "" Then
   MsgBox "Blank Instrument Type not allowed !!! "
   CmbInstrumentType.SetFocus
   Exit Sub
End If

'If Val(TxtInstrumentAmount) <= 0 Then
'   MsgBox "Blank Instrument Amount Not Allowed !!! "
'   TxtInstrumentAmount.SetFocus
'   Exit Sub
'End If

If LCase(CmbInstrumentType.Text) <> "cash" Then
   If TxtDrawnOn.Text = "" Then
      MsgBox "Blank 'Drawn On' not allowed !!! "
      TxtDrawnOn.SetFocus
      Exit Sub
   End If
   If TxtBranchName.Text = "" Then
      MsgBox "Blank Branch Name not allowed !!! "
      TxtBranchName.SetFocus
      Exit Sub
   End If
   If TxtInstrumentNo.Text = "" Then
      MsgBox "Blank Instrument No not allowed !!! "
      TxtInstrumentNo.SetFocus
      Exit Sub
   End If

End If

If CmbAgent.Text = "" Then
   MsgBox "Blank Agent Name not allowed !!! "
   CmbAgent.SetFocus
   Exit Sub
End If


If TxtRemark.Text = "" Then
   MsgBox "Blank Remark not allowed !!! "
   TxtRemark.SetFocus
   Exit Sub
End If

If CmbExchangeTypeID.Text = "" Then
   MsgBox "Blank Exchange Type not allowed !!! "
   CmbExchangeTypeID.SetFocus
   Exit Sub
End If


If MSHFlexGrid3.Rows < 3 Then
   MsgBox "Atleast one invoice detail is required!!!"
   TxtPINo.SetFocus
   Exit Sub
End If


If Edit_Flg = "A" Then
   RsTxn_CashReceipt.AddNew
   RsTxn_CashReceipt!CRID = GetCRID
   RsTxn_CashReceipt!G_UID = GetGUID
   TxtCRID.Text = RsTxn_CashReceipt!CRID
   TxtG_UID = RsTxn_CashReceipt!G_UID
Else
   RsTxn_CashReceipt.MoveFirst
   RsTxn_CashReceipt.Filter = "G_UID= '" & TxtG_UID & "'"
End If

RsTxn_CashReceipt!CRID = TxtCRID
RsTxn_CashReceipt!CRDate = TxtCRDate.Value
'RsTxn_CashReceipt!ClientIDRow = mClientIDRow
RsTxn_CashReceipt!InstrumentType = CmbInstrumentType.Text
RsTxn_CashReceipt!DrawnOn = TxtDrawnOn
RsTxn_CashReceipt!BranchName = TxtBranchName
RsTxn_CashReceipt!InstrumentNo = TxtInstrumentNo
RsTxn_CashReceipt!InstrumentDate = TxtInstrumentDate.Value
RsTxn_CashReceipt!InstrumentAmount = Val(TxtInstrumentAmount)
RsTxn_CashReceipt!Remark = TxtRemark
RsTxn_CashReceipt!InvoiceDetail = TxtInvoiceDetail
RsTxn_CashReceipt!Flag = "0" 'TxtFlag.Text
RsTxn_CashReceipt!SM_Prefix = TxtPreFix  'TxtFlag.Text
RsTxn_CashReceipt!LocationID = mLocationID
RsTxn_CashReceipt!AgentID = mAgentID
If mExchangeTypeID <> 0 Then
   RsTxn_CashReceipt!ExchangeTypeID = mExchangeTypeID
End If
If IsNull(RsTxn_CashReceipt!CreatedBy) = True Or RsTxn_CashReceipt!CreatedBy = "" Then
   RsTxn_CashReceipt!CreatedBy = Gen_UserLoginID
   RsTxn_CashReceipt!CreatedDate = Now
Else
   RsTxn_CashReceipt!UpdatedBy = Gen_UserLoginID
   RsTxn_CashReceipt!UpdatedDate = Now
End If

RsTxn_CashReceipt.Update

Call DeleteCRInvoiceDetails
Call SaveTxn_CashReceiptInvoiceDetails

RsTxn_CashReceipt.Filter = ""
RsTxn_CashReceipt.Requery
RsTxn_CashReceipt.MoveFirst
RsTxn_CashReceipt.Filter = "G_UID= '" & TxtG_UID & "'"
If RsTxn_CashReceipt.EOF Then
   RsTxn_CashReceipt.MoveLast
End If
Call InData
End Sub
Private Sub SaveTxn_CashReceiptInvoiceDetails()
Call TableTxn_CashReceiptInvoiceDetails(" Where CRID=" & TxtCRID & " AND SM_Prefix = '" & TxtPreFix & "'")
For i = 1 To MSHFlexGrid3.Rows - 1
    If MSHFlexGrid3.TextMatrix(i, 1) <> "" Then
       RsTxn_CashReceiptInvoiceDetails.AddNew
       RsTxn_CashReceiptInvoiceDetails!CRIDID = getCRInvoiceDetailsMax
       RsTxn_CashReceiptInvoiceDetails!CRID = TxtCRID
       RsTxn_CashReceiptInvoiceDetails!SM_Prefix = TxtPreFix
       RsTxn_CashReceiptInvoiceDetails!PI_No = MSHFlexGrid3.TextMatrix(i, 0)
       RsTxn_CashReceiptInvoiceDetails!PI_Date = MSHFlexGrid3.TextMatrix(i, 1)
       RsTxn_CashReceiptInvoiceDetails!Amount = Val(MSHFlexGrid3.TextMatrix(i, 2))
       RsTxn_CashReceiptInvoiceDetails!ClientIDRow = MSHFlexGrid3.TextMatrix(i, 8)
       RsTxn_CashReceiptInvoiceDetails!InvoiceDetail = MSHFlexGrid3.TextMatrix(i, 7)
       RsTxn_CashReceiptInvoiceDetails.Update
       RsTxn_CashReceiptInvoiceDetails.Requery
    End If
Next
End Sub

Private Sub DeleteCRInvoiceDetails()
tmp = "Delete From Txn_CashReceiptInvoiceDetails Where CRID=" & TxtCRID & " AND SM_Prefix = '" & TxtPreFix & "'"
'Call TableTxn_CashReceiptInvoiceDetails(" Where CRID=" & TxtCRID & " AND SM_Prefix = '" & TxtPreFix & "'")
Call Tmp4Table
End Sub
Private Function GetCRID() As Double
Dim RetVal As Double
tmp = "Select max(CRID) from Txn_CashReceipt where sm_prefix = '" & TxtPreFix & "'"
Call TmpTable
RetVal = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
RetVal = RetVal + 1
GetCRID = RetVal
End Function

Function getCRInvoiceDetailsMax() As Double
Dim RetVal As Double
RetVal = 0
tmp = "select max(CRIDID) from Txn_CashReceiptInvoiceDetails where SM_Prefix = '" & TxtPreFix & "'"
Call TmpTable
If TmpRs.RecordCount > 0 Then
   RetVal = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
End If
RetVal = RetVal + 1
TmpRs.Close
getCRInvoiceDetailsMax = RetVal

End Function
Private Sub Grid_Head()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 15
MSHFlexGrid1.TextMatrix(0, 0) = "Receipt ID"
MSHFlexGrid1.TextMatrix(0, 1) = "Location ID"
MSHFlexGrid1.TextMatrix(0, 2) = "Receipt Date"
MSHFlexGrid1.TextMatrix(0, 3) = "Client Name"
MSHFlexGrid1.TextMatrix(0, 4) = "Type"
MSHFlexGrid1.TextMatrix(0, 5) = "Drawn On"
MSHFlexGrid1.TextMatrix(0, 6) = "Branch Name"
MSHFlexGrid1.TextMatrix(0, 7) = "Instrument No"
MSHFlexGrid1.TextMatrix(0, 8) = "Instrument Date"
MSHFlexGrid1.TextMatrix(0, 9) = "Instrument Amount"
MSHFlexGrid1.TextMatrix(0, 10) = "Agent Name"
MSHFlexGrid1.TextMatrix(0, 11) = "Remark"
MSHFlexGrid1.TextMatrix(0, 12) = "Exchange Type"
MSHFlexGrid1.TextMatrix(0, 13) = "Print Flag"
MSHFlexGrid1.TextMatrix(0, 14) = "Hologram No"


MSHFlexGrid1.ColWidth(0) = 1200
MSHFlexGrid1.ColWidth(1) = 1000
MSHFlexGrid1.ColWidth(2) = 1200
MSHFlexGrid1.ColWidth(3) = 0
MSHFlexGrid1.ColWidth(4) = 1400
MSHFlexGrid1.ColWidth(5) = 1800
MSHFlexGrid1.ColWidth(6) = 1800
MSHFlexGrid1.ColWidth(7) = 1600
MSHFlexGrid1.ColWidth(8) = 1400
MSHFlexGrid1.ColWidth(9) = 1800
MSHFlexGrid1.ColWidth(10) = 1800
MSHFlexGrid1.ColWidth(11) = 3400
MSHFlexGrid1.ColWidth(12) = 1000
MSHFlexGrid1.ColWidth(13) = 1400
MSHFlexGrid1.ColWidth(14) = 1400
End Sub

Private Sub MSHFlexGrid1_Click()
Dim i As Variant
i = MSHFlexGrid1.RowSel
If i <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(i, 0) = "" Then Exit Sub
RsTxn_CashReceipt.MoveFirst
RsTxn_CashReceipt.Find "SM_Prefix = '" & MSHFlexGrid1.TextMatrix(i, 1) & "'"
If Not RsTxn_CashReceipt.EOF Then
   RsTxn_CashReceipt.Find "CRID = " & MSHFlexGrid1.TextMatrix(i, 0)
   Call InData
   Frame0.Visible = False
End If
End Sub

Private Sub TxtLocationName_LostFocus()
TxtLocationName = Trim(TxtLocationName)
End Sub

Private Sub Grid_Head1()
MSHFlexGrid3.Clear
MSHFlexGrid3.Rows = 2
MSHFlexGrid3.Cols = 9

MSHFlexGrid3.TextMatrix(0, 0) = "PI No"
MSHFlexGrid3.TextMatrix(0, 1) = "PI Date"
MSHFlexGrid3.TextMatrix(0, 2) = "Amount"
MSHFlexGrid3.TextMatrix(0, 3) = "Client ID"
MSHFlexGrid3.TextMatrix(0, 4) = "Client Name"
MSHFlexGrid3.TextMatrix(0, 5) = "DP ID"
MSHFlexGrid3.TextMatrix(0, 6) = "DP Name"
MSHFlexGrid3.TextMatrix(0, 7) = "Invoice Details"
MSHFlexGrid3.TextMatrix(0, 8) = "ClientIDRow"

MSHFlexGrid3.ColWidth(0) = 1400
MSHFlexGrid3.ColWidth(1) = 1200
MSHFlexGrid3.ColWidth(2) = 1400
MSHFlexGrid3.ColWidth(3) = 1200
MSHFlexGrid3.ColWidth(4) = 1600
MSHFlexGrid3.ColWidth(5) = 1200
MSHFlexGrid3.ColWidth(6) = 2000
MSHFlexGrid3.ColWidth(7) = 2800
MSHFlexGrid3.ColWidth(8) = 0

End Sub

Private Sub CmdAdd2List_Click()
Dim r As Variant
If TxtPINo = "" Then
   MsgBox "Blank PI No !!! "
   TxtPINo.SetFocus
   Exit Sub
End If

If TxtPIDate = "" Then
   MsgBox "Blank PI Date!!! "
   TxtPIDate.SetFocus
   Exit Sub
End If
If TxtPIDate > TxtCRDate Then
   MsgBox "Receipt Date cannot be greater than PI Date!!! "
   TxtPIDate.SetFocus
   Exit Sub
End If

If TxtAmount = "" Then
   MsgBox "Blank Invoice Amount!!! "
   TxtAmount.SetFocus
   Exit Sub
End If

If TxtInvoiceDetails = "" Then
   MsgBox "Blank Invoice Details!!! "
   TxtInvoiceDetails.SetFocus
   Exit Sub
End If


If CmbClientDP.Visible = True Then
   MsgBox "Invalid selection !!! "
   Exit Sub
End If
If CmbNonNcdexClient.Visible = True Then
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


If ChkDuplicatePINo = False Then
     MsgBox "Duplicate PI No not allowed !!!"
     TxtPINo.SetFocus
     Exit Sub
End If

If CmdAdd2List.Caption = "Update List" Then
   r = MSHFlexGrid3.RowSel
   TxtInstrumentAmount = Val(TxtInstrumentAmount) - Val(MSHFlexGrid3.TextMatrix(r, 2))
Else
  r = MSHFlexGrid3.Rows - 1
  MSHFlexGrid3.Rows = MSHFlexGrid3.Rows + 1
End If

MSHFlexGrid3.TextMatrix(r, 0) = TxtPINo
MSHFlexGrid3.TextMatrix(r, 1) = TxtPIDate
MSHFlexGrid3.TextMatrix(r, 2) = TxtAmount
MSHFlexGrid3.TextMatrix(r, 3) = TxtClientID
MSHFlexGrid3.TextMatrix(r, 4) = TxtClientName 'CmbNonNcdexClient.Text
MSHFlexGrid3.TextMatrix(r, 5) = TxtDPID 'CmbNonNcdexClient.Text
MSHFlexGrid3.TextMatrix(r, 6) = TxtDPName ' .Text
MSHFlexGrid3.TextMatrix(r, 7) = TxtInvoiceDetails
MSHFlexGrid3.TextMatrix(r, 8) = mClientIDRow
TxtInstrumentAmount = Val(TxtInstrumentAmount) + Val(MSHFlexGrid3.TextMatrix(r, 2))
TxtCRDate.Enabled = False
Call ClearInvoiceDetails
TxtPINo.SetFocus
End Sub
Function ChkDuplicatePINo() As Boolean
Dim RetVal As Boolean
RetVal = True
If MSHFlexGrid3.Rows > 2 Then
   For u = 1 To MSHFlexGrid3.Rows - 1
       If MSHFlexGrid3.TextMatrix(u, 0) = TxtPINo Then
          If LCase(CmdAdd2List.Caption) = "add to list" Then
             RetVal = False
             Exit For
          Else
             If u <> MSHFlexGrid3.RowSel Then
                RetVal = False
             End If
          End If
       End If
   Next
End If

ChkDuplicatePINo = RetVal
End Function

Sub ClearInvoiceDetails()
TxtPINo = ""
TxtPIDate.Value = Date
TxtAmount = ""
TxtClientID = ""
TxtClientName = ""
TxtDPID = ""
TxtDPName = ""
CmbClientDP.Text = ""
CmbClientDP.Visible = False
CmbNonNcdexClient.Text = ""
TxtInvoiceDetails = ""
CmdAdd2List.Caption = "Add to List"
CmdDeleteList.Enabled = False
End Sub

Private Sub MSHFlexGrid3_Click()
If SaveCmd.Enabled = False Then Exit Sub
r = MSHFlexGrid3.RowSel
If MSHFlexGrid3.TextMatrix(r, 0) = "" Then Exit Sub
CmdAdd2List.Caption = "Update List"
'TxtClientID.Locked = False
If r <= 0 Then Exit Sub
TxtPINo = MSHFlexGrid3.TextMatrix(r, 0)
TxtPIDate = MSHFlexGrid3.TextMatrix(r, 1)
TxtAmount = MSHFlexGrid3.TextMatrix(r, 2)
'CmbNonNcdexClient.Text = MSHFlexGrid3.TextMatrix(r, 3)
TxtClientID = MSHFlexGrid3.TextMatrix(r, 3)
TxtClientName = MSHFlexGrid3.TextMatrix(r, 4)
TxtDPID = MSHFlexGrid3.TextMatrix(r, 5)
TxtDPName = MSHFlexGrid3.TextMatrix(r, 6)
TxtInvoiceDetails = MSHFlexGrid3.TextMatrix(r, 7)

If CmbNonNcdexClient.Visible = True Then
   CmbNonNcdexClient.Text = MSHFlexGrid3.TextMatrix(r, 4)
End If

CmdDeleteList.Enabled = True
End Sub

Private Sub CmdDeleteList_Click()
i = MSHFlexGrid3.RowSel
If i > 0 Then
   If MSHFlexGrid3.TextMatrix(i, 0) <> "" Then
      TxtInstrumentAmount = Val(TxtInstrumentAmount) - Val(MSHFlexGrid3.TextMatrix(i, 2))
      For RO = i To MSHFlexGrid3.Rows - 2
          For c = 0 To MSHFlexGrid3.Cols - 1
               MSHFlexGrid3.TextMatrix(RO, c) = MSHFlexGrid3.TextMatrix(RO + 1, c)
          Next
      Next
      MSHFlexGrid3.Rows = MSHFlexGrid3.Rows - 1
   End If
End If

Call ClearInvoiceDetails
End Sub


Private Sub TxtPIDate_LostFocus()
If TxtPIDate.Value > Date Then
   MsgBox "Future Proforma Invoice Date is not allowed!!!"
   TxtPIDate.SetFocus
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


Private Sub TxtClientID_KeyPress(KeyAscii As Integer)
tmp = KeyAscii
If KeyAscii = 14 Then
   SSTab1.Tab = 1
   TxtClientID.SetFocus
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
CmdAdd2List.SetFocus
End Sub
