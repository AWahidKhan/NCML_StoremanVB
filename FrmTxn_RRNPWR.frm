VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmTxn_RRNPWR 
   Caption         =   "RRN PWR Mapping"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   12915
   ScaleWidth      =   21360
   WindowState     =   2  'Maximized
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      TabIndex        =   83
      Top             =   10560
      Width           =   10935
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   9015
      Left            =   15000
      TabIndex        =   82
      Top             =   0
      Width           =   255
   End
   Begin VB.Frame Frame4 
      Height          =   9735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15135
      Begin VB.Frame Frame0 
         Height          =   1245
         Left            =   720
         TabIndex        =   75
         Top             =   6600
         Visible         =   0   'False
         Width           =   13920
         Begin VB.ComboBox CmbSStatus 
            Height          =   315
            ItemData        =   "FrmTxn_RRNPWR.frx":0000
            Left            =   4200
            List            =   "FrmTxn_RRNPWR.frx":000A
            Sorted          =   -1  'True
            TabIndex        =   88
            Top             =   855
            Width           =   1755
         End
         Begin VB.TextBox TxtSISINNo 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4080
            TabIndex        =   24
            Top             =   510
            Width           =   2085
         End
         Begin VB.ComboBox CmbSCommodity 
            Height          =   315
            Left            =   915
            Sorted          =   -1  'True
            TabIndex        =   27
            Top             =   855
            Width           =   2325
         End
         Begin VB.ComboBox CmbSCmp 
            Height          =   315
            Left            =   8670
            Sorted          =   -1  'True
            TabIndex        =   26
            Top             =   510
            Width           =   3690
         End
         Begin VB.CommandButton CmdSearch 
            Caption         =   "Search"
            Height          =   350
            Left            =   12390
            TabIndex        =   28
            Top             =   495
            Width           =   1000
         End
         Begin VB.ComboBox CmbSLocation 
            Height          =   315
            Left            =   6195
            Sorted          =   -1  'True
            TabIndex        =   25
            Top             =   510
            Width           =   2430
         End
         Begin VB.TextBox TxtSRRNNo 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2850
            TabIndex        =   23
            Top             =   510
            Width           =   1185
         End
         Begin VB.CommandButton CmdExcel 
            Caption         =   "Excel"
            Height          =   350
            Left            =   13395
            TabIndex        =   29
            Top             =   495
            Width           =   1000
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   1455
            Left            =   75
            TabIndex        =   76
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
         Begin MSComCtl2.DTPicker STo 
            Height          =   300
            Left            =   1470
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   495
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
            Format          =   48955393
            CurrentDate     =   36494
         End
         Begin MSComCtl2.DTPicker SFrom 
            Height          =   300
            Left            =   135
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   495
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
            Format          =   48955393
            CurrentDate     =   36494
         End
         Begin VB.Label Label19 
            Caption         =   "From Date"
            Height          =   255
            Left            =   420
            TabIndex        =   94
            Top             =   240
            Width           =   825
         End
         Begin VB.Label Label45 
            Caption         =   "To Date"
            Height          =   255
            Left            =   1755
            TabIndex        =   93
            Top             =   240
            Width           =   675
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
            Height          =   225
            Left            =   3555
            TabIndex        =   89
            Top             =   900
            Width           =   810
         End
         Begin VB.Label LblISINNo 
            AutoSize        =   -1  'True
            Caption         =   "ISIN No"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   4837
            TabIndex        =   81
            Top             =   240
            Width           =   570
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Commodity"
            Height          =   195
            Left            =   90
            TabIndex        =   80
            Top             =   915
            Width           =   765
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "CMP"
            Height          =   195
            Left            =   10350
            TabIndex        =   79
            Top             =   240
            Width           =   345
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Location"
            Height          =   195
            Left            =   7103
            TabIndex        =   78
            Top             =   240
            Width           =   615
         End
         Begin VB.Label LblSearchRRN 
            AutoSize        =   -1  'True
            Caption         =   "RRN No"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   3135
            TabIndex        =   77
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame ButtonFrm 
         Height          =   1200
         Left            =   135
         TabIndex        =   74
         Top             =   8400
         Width           =   14805
         Begin VB.CommandButton ExitCmd 
            Caption         =   "E&xit"
            Height          =   400
            Left            =   11775
            TabIndex        =   20
            Top             =   645
            Width           =   2835
         End
         Begin VB.CommandButton SearchCmd 
            Caption         =   "Search"
            Height          =   400
            Left            =   11775
            TabIndex        =   15
            Top             =   240
            Width           =   2835
         End
         Begin VB.CommandButton LastCmd 
            Caption         =   "&Last"
            Height          =   400
            Left            =   8940
            TabIndex        =   19
            Top             =   645
            Width           =   2835
         End
         Begin VB.CommandButton DeleteCmd 
            Caption         =   "&Delete"
            Height          =   400
            Left            =   8940
            TabIndex        =   14
            Top             =   240
            Width           =   2835
         End
         Begin VB.CommandButton FirstCmd 
            Caption         =   "&First"
            Height          =   400
            Left            =   5985
            TabIndex        =   18
            Top             =   645
            Width           =   2955
         End
         Begin VB.CommandButton EditCmd 
            Caption         =   "&Edit"
            Height          =   400
            Left            =   5985
            TabIndex        =   13
            Top             =   240
            Width           =   2955
         End
         Begin VB.CommandButton PreviousCmd 
            Caption         =   "&Previous"
            Height          =   400
            Left            =   3030
            TabIndex        =   17
            Top             =   645
            Width           =   2955
         End
         Begin VB.CommandButton SaveCmd 
            Caption         =   "Save"
            Height          =   400
            Left            =   3030
            TabIndex        =   12
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
            TabIndex        =   11
            Top             =   240
            Width           =   2955
         End
      End
      Begin VB.Frame Frame1 
         Height          =   4395
         Left            =   120
         TabIndex        =   34
         Top             =   120
         Width           =   14820
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
            ItemData        =   "FrmTxn_RRNPWR.frx":0021
            Left            =   10005
            List            =   "FrmTxn_RRNPWR.frx":002B
            Sorted          =   -1  'True
            TabIndex        =   4
            Top             =   1005
            Width           =   4650
         End
         Begin VB.TextBox TxtSM_Prefix 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   325
            Left            =   14640
            Locked          =   -1  'True
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   225
            Visible         =   0   'False
            Width           =   360
         End
         Begin VB.TextBox TxtGrade 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   360
            Left            =   10080
            Locked          =   -1  'True
            TabIndex        =   45
            TabStop         =   0   'False
            Top             =   2280
            Width           =   1875
         End
         Begin VB.TextBox TxtBalanceQty 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            CausesValidation=   0   'False
            Height          =   360
            Left            =   13200
            Locked          =   -1  'True
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   2700
            Width           =   1440
         End
         Begin VB.TextBox TxtRRNNo 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   12615
            MaxLength       =   20
            TabIndex        =   3
            Top             =   600
            Width           =   2010
         End
         Begin VB.TextBox TxtDeliveredQty 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   360
            Left            =   10080
            Locked          =   -1  'True
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   2700
            Width           =   1890
         End
         Begin VB.TextBox TxtFEDDate 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   360
            Left            =   13200
            Locked          =   -1  'True
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   2280
            Width           =   1440
         End
         Begin VB.ComboBox CmbCommodityID 
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
            ItemData        =   "FrmTxn_RRNPWR.frx":0042
            Left            =   1440
            List            =   "FrmTxn_RRNPWR.frx":0044
            Sorted          =   -1  'True
            TabIndex        =   2
            Top             =   600
            Width           =   10410
         End
         Begin VB.TextBox TxtConvertedID 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   360
            Left            =   1440
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   180
            Width           =   2880
         End
         Begin VB.TextBox TxtAdjQty 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   1440
            TabIndex        =   5
            Top             =   3120
            Width           =   2880
         End
         Begin VB.TextBox TxtClientID 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   360
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   1860
            Width           =   2880
         End
         Begin VB.TextBox TxtISINID 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   360
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   2280
            Width           =   2880
         End
         Begin VB.TextBox TxtDPID 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   360
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   1425
            Width           =   2880
         End
         Begin VB.TextBox TxtRRNQty 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   360
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   2700
            Width           =   2880
         End
         Begin VB.TextBox TxtRemarks 
            Appearance      =   0  'Flat
            Height          =   795
            Left            =   1440
            MaxLength       =   500
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   7
            Top             =   3525
            Width           =   13215
         End
         Begin VB.TextBox TxtRRNID 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   360
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   84
            TabStop         =   0   'False
            Top             =   1005
            Width           =   2880
         End
         Begin VB.TextBox TxtCommodityID 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   51
            TabStop         =   0   'False
            Top             =   600
            Width           =   10395
         End
         Begin VB.TextBox TxtStatus 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   10005
            Locked          =   -1  'True
            TabIndex        =   87
            TabStop         =   0   'False
            Top             =   1005
            Width           =   4650
         End
         Begin VB.ComboBox CmbLocation 
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
            ItemData        =   "FrmTxn_RRNPWR.frx":0046
            Left            =   10005
            List            =   "FrmTxn_RRNPWR.frx":0048
            Sorted          =   -1  'True
            TabIndex        =   1
            Top             =   180
            Width           =   4650
         End
         Begin VB.TextBox TxtLocation 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   10005
            Locked          =   -1  'True
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   180
            Width           =   4650
         End
         Begin MSComCtl2.DTPicker TxtPWRDate 
            Height          =   375
            Left            =   5880
            TabIndex        =   90
            Top             =   165
            Width           =   1335
            _ExtentX        =   2355
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
            Format          =   48955393
            CurrentDate     =   36494
         End
         Begin VB.TextBox TxtOutOfSystemQty 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   360
            Left            =   5880
            Locked          =   -1  'True
            TabIndex        =   49
            TabStop         =   0   'False
            Top             =   2700
            Width           =   2775
         End
         Begin VB.TextBox TxtCMPID 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   360
            Left            =   5880
            Locked          =   -1  'True
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   2280
            Width           =   2760
         End
         Begin VB.TextBox TxtDPName 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   325
            Left            =   5880
            Locked          =   -1  'True
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   1425
            Width           =   8760
         End
         Begin VB.TextBox TxtClientName 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   360
            Left            =   5880
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   1860
            Width           =   8760
         End
         Begin VB.TextBox TxtRRNDate 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   360
            Left            =   5880
            Locked          =   -1  'True
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   1005
            Width           =   2745
         End
         Begin VB.TextBox TxtConvertedClient 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   360
            Left            =   5880
            Locked          =   -1  'True
            TabIndex        =   92
            TabStop         =   0   'False
            Top             =   3120
            Width           =   8760
         End
         Begin VB.ComboBox CmbConvertedClient 
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
            ItemData        =   "FrmTxn_RRNPWR.frx":004A
            Left            =   5880
            List            =   "FrmTxn_RRNPWR.frx":004C
            Sorted          =   -1  'True
            TabIndex        =   6
            Top             =   3120
            Visible         =   0   'False
            Width           =   8760
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            Caption         =   "PWR Date"
            Height          =   195
            Left            =   4440
            TabIndex        =   91
            Top             =   240
            Width           =   780
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Status"
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
            Left            =   8880
            TabIndex        =   86
            Top             =   1080
            Width           =   555
         End
         Begin VB.Label Label12 
            Caption         =   "Remarks"
            Height          =   210
            Left            =   120
            TabIndex        =   73
            Top             =   3810
            Width           =   1800
         End
         Begin VB.Label Label7 
            Caption         =   "Adjusted Qty by CMG"
            Height          =   390
            Left            =   4440
            TabIndex        =   72
            Top             =   2685
            Width           =   1320
         End
         Begin VB.Label LblExchangeTypeID 
            Caption         =   "SM_Prefix"
            Height          =   210
            Left            =   14280
            TabIndex        =   71
            Top             =   240
            Visible         =   0   'False
            Width           =   1440
         End
         Begin VB.Label LblLocationID 
            AutoSize        =   -1  'True
            Caption         =   "Location"
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
            Left            =   8880
            TabIndex        =   70
            Top             =   240
            Width           =   750
         End
         Begin VB.Label LblCMPID 
            Caption         =   "CMP"
            Height          =   210
            Left            =   4440
            TabIndex        =   69
            Top             =   2355
            Width           =   1080
         End
         Begin VB.Label Label2 
            Caption         =   "Grade"
            Height          =   180
            Left            =   8880
            TabIndex        =   68
            Top             =   2370
            Width           =   900
         End
         Begin VB.Label LblCommodity 
            Caption         =   "Commodity"
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
            Left            =   120
            TabIndex        =   67
            Top             =   668
            Width           =   1140
         End
         Begin VB.Label Label8 
            Caption         =   "Client Name "
            Height          =   255
            Left            =   4440
            TabIndex        =   66
            Top             =   1913
            Width           =   975
         End
         Begin VB.Label Label10 
            Caption         =   "DP ID"
            Height          =   195
            Left            =   120
            TabIndex        =   65
            Top             =   1508
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "RRN Balance Qty"
            Height          =   420
            Left            =   12120
            TabIndex        =   64
            Top             =   2670
            Width           =   1200
         End
         Begin VB.Label LblRRNID 
            Caption         =   "RRN ID"
            Height          =   240
            Left            =   120
            TabIndex        =   63
            Top             =   1065
            Width           =   1245
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
            Left            =   11880
            TabIndex        =   62
            Top             =   675
            Width           =   795
         End
         Begin VB.Label LblRRNDate 
            Caption         =   "RRN Date"
            Height          =   240
            Left            =   4440
            TabIndex        =   61
            Top             =   1065
            Width           =   945
         End
         Begin VB.Label LblISINID 
            Caption         =   "ISIN No"
            Height          =   225
            Left            =   120
            TabIndex        =   60
            Top             =   2348
            Width           =   825
         End
         Begin VB.Label LblClientID 
            Caption         =   "Client ID"
            Height          =   225
            Left            =   120
            TabIndex        =   59
            Top             =   1928
            Width           =   1800
         End
         Begin VB.Label LblRRNQty 
            Caption         =   "RRN Qty in MT"
            Height          =   225
            Left            =   120
            TabIndex        =   58
            Top             =   2768
            Width           =   1560
         End
         Begin VB.Label LblDeliveredQty 
            Caption         =   "Qty Delivered in StoreMan"
            Height          =   360
            Left            =   8880
            TabIndex        =   57
            Top             =   2700
            Width           =   1080
         End
         Begin VB.Label Label11 
            Caption         =   "DP Name "
            Height          =   225
            Left            =   4440
            TabIndex        =   56
            Top             =   1475
            Width           =   840
         End
         Begin VB.Label Label13 
            Caption         =   "Conversion margin Qty"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   120
            TabIndex        =   55
            Top             =   3083
            Width           =   1245
         End
         Begin VB.Label Label14 
            Caption         =   "Converted Client Name "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   4440
            TabIndex        =   54
            Top             =   3075
            Width           =   1140
         End
         Begin VB.Label Label15 
            Caption         =   "FED Date"
            Height          =   240
            Left            =   12120
            TabIndex        =   53
            Top             =   2340
            Width           =   945
         End
         Begin VB.Label Label16 
            Caption         =   "Converted ID"
            Height          =   210
            Left            =   120
            TabIndex        =   52
            Top             =   255
            Width           =   1155
         End
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   3855
         Left            =   120
         TabIndex        =   8
         Top             =   4560
         Width           =   14820
         _ExtentX        =   26141
         _ExtentY        =   6800
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "GSL Details"
         TabPicture(0)   =   "FrmTxn_RRNPWR.frx":004E
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "MSHFlexGrid11"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "MSHFlexGrid2"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "TxtEnterData"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Check1"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).ControlCount=   4
         TabCaption(1)   =   "Log History"
         TabPicture(1)   =   "FrmTxn_RRNPWR.frx":006A
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "TxtUpdated"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "TxtCreated"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "Label9"
         Tab(1).Control(3)=   "Label6"
         Tab(1).ControlCount=   4
         Begin VB.CheckBox Check1 
            Caption         =   "Select All"
            Height          =   255
            Left            =   120
            TabIndex        =   85
            Top             =   480
            Width           =   1095
         End
         Begin VB.TextBox TxtEnterData 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   1320
            TabIndex        =   9
            Top             =   960
            Visible         =   0   'False
            Width           =   1800
         End
         Begin VB.TextBox TxtUpdated 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   325
            Left            =   -73875
            Locked          =   -1  'True
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   945
            Width           =   6120
         End
         Begin VB.TextBox TxtCreated 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   325
            Left            =   -73875
            Locked          =   -1  'True
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   480
            Width           =   6120
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
            Height          =   2880
            Left            =   120
            TabIndex        =   10
            Top             =   840
            Width           =   14520
            _ExtentX        =   25612
            _ExtentY        =   5080
            _Version        =   393216
            Rows            =   3
            FixedRows       =   2
            FixedCols       =   0
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid11 
            Height          =   2880
            Left            =   120
            TabIndex        =   95
            Top             =   720
            Visible         =   0   'False
            Width           =   14520
            _ExtentX        =   25612
            _ExtentY        =   5080
            _Version        =   393216
            Rows            =   3
            FixedRows       =   2
            FixedCols       =   0
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin VB.Label Label9 
            Caption         =   "Updated :"
            Height          =   255
            Left            =   -74760
            TabIndex        =   33
            Top             =   975
            Width           =   1695
         End
         Begin VB.Label Label6 
            Caption         =   "Created  :"
            Height          =   255
            Left            =   -74760
            TabIndex        =   32
            Top             =   510
            Width           =   1695
         End
      End
   End
End
Attribute VB_Name = "FrmTxn_RRNPWR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mLocationID, mCommodityID, mConvertedClientIDRow, mNewGSLID, mCMPID, mSCommodityID As Variant
Dim Edit_Flg As Variant
Dim mColsel, mRowsel As Single
Const strChecked = "þ"
Const strUnChecked = "q"
Dim mBODDate As Variant
Dim mCommtrackFlag As Variant
Dim mCommodityGroupId As Variant
Dim mRejInvIDs As Variant

Public Sub Grid_Head2()
Check1.Value = vbUnchecked
With MSHFlexGrid2
   .Clear
   .Cols = 10
   .Rows = 3
   .FixedRows = 2
   .WordWrap = True
   
   .TextMatrix(1, 0) = "Selection"
   .TextMatrix(1, 1) = "Godown"
   .TextMatrix(1, 2) = "Stack No"
   .TextMatrix(1, 3) = "Lot No"
   .TextMatrix(1, 4) = "Balance Bags"
   .TextMatrix(1, 5) = "Balance Qty"
   .TextMatrix(1, 6) = "GSLID"
   .TextMatrix(1, 7) = "Converted Qty"
   .TextMatrix(1, 8) = "Converted Bags"
   .TextMatrix(1, 9) = "ConvertedGSLID"
   
   .ColWidth(0) = 800
   .ColWidth(1) = 1500
   .ColWidth(2) = 1500
   .ColWidth(3) = 1500
   .ColWidth(4) = 1300
   .ColWidth(5) = 1300
   .ColWidth(6) = 0
   .ColWidth(7) = 1300
   .ColWidth(8) = 1300
   .ColWidth(9) = 0
   
   .RowHeight(1) = 300
End With
End Sub

Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)
Call Grid_HeadRRNNCDEXInvoice
Call ClearControls
Check1.Enabled = True
CmbCommodityID.Visible = True
'CmbConvertedClient.Visible = True
CmbLocation.Visible = True
TxtCommodityID.Visible = False
'TxtConvertedClient.Visible = False
TxtLocation.Visible = False
TxtAdjQty.Locked = False
TxtRRNNo.Locked = False
TxtPWRDate.Enabled = True

TxtStatus.Visible = True
CmbStatus.Visible = False
TxtStatus.Text = "Pending"
CmbStatus.Text = TxtStatus.Text

Frame1.Enabled = True
If RsTxn_RRNPWR.RecordCount > 0 Then
   CmbLocation.SetFocus
End If
End Sub

Private Sub ClearControls()
TxtConvertedID.Text = ""
TxtPWRDate.Value = Date
TxtLocation.Text = ""
CmbLocation.Text = ""
TxtCommodityID.Text = ""
CmbCommodityID.Text = ""
TxtRRNNo.Text = ""
TxtRRNID.Text = ""
TxtRRNDate.Text = ""
TxtDPID.Text = ""
TxtDPName.Text = ""
TxtClientID.Text = ""
TxtClientName.Text = ""
TxtISINID.Text = ""
TxtCMPID.Text = ""
TxtGrade.Text = ""
TxtFEDDate.Text = ""
TxtRRNQty.Text = ""
TxtOutOfSystemQty.Text = ""
TxtDeliveredQty.Text = ""
TxtBalanceQty.Text = ""
TxtAdjQty.Text = ""
'TxtConvertedClient.Text = ""
'CmbConvertedClient.Text = ""
TxtRemarks.Text = ""
CmbStatus.Text = ""
TxtStatus.Text = ""

TxtCreated.Text = ""
TxtUpdated.Text = ""

Call Grid_Head2
End Sub

Private Sub Check1_Click()
If MSHFlexGrid2.Rows = 3 Then
   Check1.Value = vbUnchecked
   Exit Sub
End If

If Check1.Value = vbUnchecked Then
   MSHFlexGrid2.TextMatrix(0, 7) = ""
   MSHFlexGrid2.TextMatrix(0, 8) = ""
   For I = 2 To MSHFlexGrid2.Rows - 2
      MSHFlexGrid2.TextMatrix(I, 0) = strUnChecked
      MSHFlexGrid2.TextMatrix(I, 7) = ""
      MSHFlexGrid2.TextMatrix(I, 8) = ""
   Next
Else
   MSHFlexGrid2.TextMatrix(0, 7) = MSHFlexGrid2.TextMatrix(0, 5)
   MSHFlexGrid2.TextMatrix(0, 8) = MSHFlexGrid2.TextMatrix(0, 4)
   For I = 2 To MSHFlexGrid2.Rows - 2
      MSHFlexGrid2.TextMatrix(I, 0) = strChecked
      MSHFlexGrid2.TextMatrix(I, 7) = MSHFlexGrid2.TextMatrix(I, 5)
      MSHFlexGrid2.TextMatrix(I, 8) = MSHFlexGrid2.TextMatrix(I, 4)
   Next
End If
End Sub

Private Sub CmbConvertedClient_GotFocus()
tmp = CmbConvertedClient.Text
Call TableMst_Client("where ExchangeTypeId = 1 And Status = 'A' And Converted = 1")
CmbConvertedClient.Clear
If RsMst_Client.RecordCount = 0 Then
   MsgBox "No record found!!! "
   Exit Sub
End If
For I = 1 To RsMst_Client.RecordCount
    CmbConvertedClient.AddItem RsMst_Client!ClientName
    RsMst_Client.MoveNext
Next
CmbConvertedClient.Text = tmp
End Sub

Private Sub CmbConvertedClient_LostFocus()
If CmbConvertedClient.Text = "" Then
   Exit Sub
End If
RsMst_Client.MoveFirst
RsMst_Client.Find "ClientName = '" & CmbConvertedClient.Text & "'"
If RsMst_Client.EOF Then
   MsgBox "Invalid selection!!! "
   CmbConvertedClient.SetFocus
   Exit Sub
End If
mConvertedClientIDRow = RsMst_Client!ClientIDRow
End Sub

Private Sub CmbLocation_GotFocus()
tmp = CmbLocation.Text
If Gen_WcLocation = "" Then
    Call TableMst_Location("")
Else
    Call TableMst_Location(" WHERE " & Gen_WcLocation)
End If
CmbLocation.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Record found!!!"
   'CmbLocation.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbLocation.AddItem RsMst_Location!LocationName
    RsMst_Location.MoveNext
Next
CmbLocation.Text = tmp
End Sub

Private Sub CmbLocation_LostFocus()
If Trim(CmbLocation.Text) = "" Then
   TxtSM_Prefix.Text = ""
   CmbCommodityID.Text = ""
   TxtRRNNo.Text = ""
   Call ClearRRNDetail
   Exit Sub
End If

If tmp <> CmbLocation.Text Then
   CmbCommodityID.Text = ""
   TxtRRNNo.Text = ""
   Call ClearRRNDetail
End If

RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocation.Text & "'"

If RsMst_Location.EOF Then
    MsgBox "Invalid Selection!!!"
    CmbLocation.SetFocus
    Exit Sub
End If
mLocationID = RsMst_Location!LocationID
TxtSM_Prefix.Text = RsMst_Location!SM_Prefix
End Sub

Private Sub CmbCommodityID_GotFocus()
If Trim(CmbLocation.Text) = "" Then
   MsgBox "First Select Location!!!"
   CmbLocation.SetFocus
   Exit Sub
End If

tmp1 = CmbCommodityID.Text
If Edit_Flg <> "A" Then
   tmp = "Select MCom.CommodityID, MCom.Commodity From Mst_Commodity MCom Where MCom.CommodityID = " & mCommodityID
Else
   tmp = "Select MCom.CommodityID, MCom.Commodity From Mst_Commodity MCom "
   tmp = tmp & " Inner Join Mst_LocationCOmmodityMapping MLCM on MLCM.CommodityID = MCom.CommodityID "
   tmp = tmp & " Where MLCM.LocationID = " & mLocationID
End If
Call Tmp7Table

CmbCommodityID.Clear
If TmpRs7.RecordCount = 0 Then
   MsgBox "No Record found!!!"
   Exit Sub
End If

For I = 1 To TmpRs7.RecordCount
   CmbCommodityID.AddItem TmpRs7!Commodity
   TmpRs7.MoveNext
Next
CmbCommodityID.Text = tmp1
End Sub

Private Sub CmbCommodityID_LostFocus()
If Trim(CmbCommodityID.Text) = "" Then
   TxtRRNNo.Text = ""
   Call ClearRRNDetail
   Exit Sub
End If
If tmp1 <> CmbCommodityID.Text Then
   TxtRRNNo.Text = ""
   Call ClearRRNDetail
End If
TmpRs7.MoveFirst
TmpRs7.Find "Commodity = '" & CmbCommodityID.Text & "'"
If TmpRs7.EOF Then
   MsgBox "Invalid selection!!!"
   CmbCommodityID.SetFocus
   Exit Sub
End If

mCommodityID = TmpRs7!CommodityID
End Sub

Private Sub CmbSCmp_GotFocus()
If CmbSLocation.Text = "" Then
   MsgBox "Select Location!!!"
   CmbSLocation.SetFocus
   Exit Sub
End If

tmp = CmbSCmp.Text
Call TableMst_CMP(" Where CloseDate is null And LocationID = " & mLocationID)
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
If CmbSCmp.Text = "" Then Exit Sub

RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPName = '" & CmbSCmp.Text & "'"
If RsMst_CMP.EOF Then
   MsgBox "Invalid selection "
   CmbSCmp.SetFocus
   Exit Sub
End If

mCMPID = RsMst_CMP!CMPID

End Sub

Private Sub CmbSCommodity_GotFocus()
If Trim(CmbSLocation.Text) = "" Then
   MsgBox "First Select Location!!!"
   CmbSLocation.SetFocus
   Exit Sub
End If

tmp1 = CmbSCommodity.Text

tmp = "Select MCom.CommodityID, MCom.Commodity From Mst_Commodity MCom "
tmp = tmp & " Inner Join Mst_LocationCOmmodityMapping MLCM on MLCM.CommodityID = MCom.CommodityID "
tmp = tmp & " Where MLCM.LocationID = " & mLocationID
Call Tmp7Table

CmbSCommodity.Clear
If TmpRs7.RecordCount = 0 Then
   MsgBox "No Record found!!!"
   Exit Sub
End If

For I = 1 To TmpRs7.RecordCount
   CmbSCommodity.AddItem TmpRs7!Commodity
   TmpRs7.MoveNext
Next
CmbSCommodity.Text = tmp1

End Sub

Private Sub CmbSCommodity_LostFocus()
If Trim(CmbSCommodity.Text) = "" Then Exit Sub

TmpRs7.MoveFirst
TmpRs7.Find "Commodity = '" & CmbSCommodity.Text & "'"
If TmpRs7.EOF Then
   MsgBox "Invalid selection!!!"
   CmbSCommodity.SetFocus
   Exit Sub
End If

mSCommodityID = TmpRs7!CommodityID

End Sub

Private Sub CmbSLocation_GotFocus()
tmp = CmbSLocation.Text
Call TableMst_Location
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
   CmbSCmp.Text = ""
   CmbSCommodity.Text = ""
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbSLocation.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection "
   CmbSLocation.SetFocus
   Exit Sub
End If

If tmp <> CmbSLocation.Text Then
   CmbSCmp.Text = ""
   CmbSCommodity.Text = ""
End If

mLocationID = RsMst_Location!LocationID
End Sub

Private Sub CmbSStatus_LostFocus()
If Trim(CmbSStatus.Text) = "" Then Exit Sub

If LCase(CmbSStatus.Text) <> "pending" And LCase(CmbSStatus.Text) <> "approved" Then
   MsgBox "Invalid Selection!!!"
   CmbSStatus.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmbStatus_LostFocus()
If Trim(CmbStatus.Text) = "" Then Exit Sub

If LCase(CmbStatus.Text) <> "pending" And LCase(CmbStatus.Text) <> "approved" Then
   MsgBox "Invalid Selection!!!"
   CmbStatus.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmdSearch_Click()
Dim wc As Variant

Call Grid_Head
'MsgBox "Wait till next message!!!"
wc = "Where TRP.PWRDate Between '" & SFrom.Value & "' And '" & STo.Value & "'"

If Trim(TxtSRRNNo.Text) <> "" Then
   If wc = "" Then
      wc = " Where RRN.RRNNo = '" & TxtSRRNNo.Text & "'"
   Else
      wc = wc & " And RRN.RRNNo = '" & TxtSRRNNo.Text & "'"
   End If
End If
If Trim(TxtSISINNo.Text) <> "" Then
   If wc = "" Then
      wc = " Where ISIN.ISINID = '" & TxtSISINNo.Text & "'"
   Else
      wc = wc & " And ISIN.ISINID = '" & TxtSISINNo.Text & "'"
   End If
End If
If Trim(CmbSLocation.Text) <> "" Then
   If wc = "" Then
      wc = " Where ML.LocationID = " & mLocationID
   Else
      wc = wc & " And ML.LocationID = " & mLocationID
   End If
End If
If Trim(CmbSCmp.Text) <> "" Then
   If wc = "" Then
      wc = " Where CMP.CMPID = " & mCMPID
   Else
      wc = wc & " And CMP.CMPID = " & mCMPID
   End If
End If
If Trim(CmbSCommodity.Text) <> "" Then
   If wc = "" Then
      wc = " Where MCom.CommodityID = " & mSCommodityID
   Else
      wc = wc & " And MCom.CommodityID = " & mSCommodityID
   End If
End If
If Trim(CmbSStatus.Text) <> "" Then
   If wc = "" Then
      wc = " Where TRP.Flag= '" & Left$(CmbSStatus.Text, 1) & "'"
   Else
      wc = wc & " And TRP.Flag= '" & Left$(CmbSStatus.Text, 1) & "'"
   End If
End If

tmp = " Select TRP.ConvertedID, TRP.PWRDate, ML.LocationName, MCom.Commodity, RRN.RRNNo, TRP.RRNID, RRN.RRNDate, DP.DepositoryID, DP.DepositoryName,"
tmp = tmp & " C.ClientID, C.ClientName, ISIN.ISINID, CMP.CMPName, ISIN.Grade, ISIN.FED, RRN.RRNQty, RRN.OutOfSystemQty, RRN.DeliveredQty,"
tmp = tmp & " RRN.RRNQty - RRN.OutOfSystemQty - RRN.DeliveredQty BalanceQty, TRP.AdjstQty , MC.ClientName 'ConvertedClient',"
tmp = tmp & " TRP.Remarks, TRP.Flag,RRN.CommtrackFlag"
tmp = tmp & " from Txn_RRNPWR TRP"
tmp = tmp & " INNER JOIN Mst_Location ML ON ML.LocationID = TRP.LocationID"
tmp = tmp & " INNER JOIN Mst_Commodity MCom ON MCom.CommodityID = TRP.CommodityID"
tmp = tmp & " INNER JOIN Mst_RRN RRN ON RRN.RRNID = TRP.RRNID"
tmp = tmp & " INNER JOIN Mst_Client MC ON MC.ClientIDRow = TRP.ConvertedClientIDRow"
tmp = tmp & " INNER JOIN Mst_ISIN ISIN ON RRN.ISINID = ISIN.ISINID"
tmp = tmp & " INNER JOIN Mst_CMP CMP ON ISIN.CMPID = CMP.CMPID"
tmp = tmp & " INNER JOIN Mst_Client C ON  RRN.ClientIDRow = C.ClientIDRow"
tmp = tmp & " INNER JOIN Mst_Depository DP ON C.DepositoryID = DP.DepositoryID AND DP.ExchangeTypeID = 2"
tmp = tmp & wc
tmp = tmp & " Order by TRP.ConvertedID"
Call TmpTable
If TmpRs.RecordCount > 0 Then
   CmdExcel.Enabled = True
   With MSHFlexGrid1
      .Rows = .Rows + TmpRs.RecordCount
      For iu = 1 To TmpRs.RecordCount
         .TextMatrix(iu, 0) = IIf(IsNull(TmpRs!ConvertedID), "", TmpRs!ConvertedID)
         .TextMatrix(iu, 1) = IIf(IsNull(TmpRs!PWRDate), Date, TmpRs!PWRDate)
         .TextMatrix(iu, 2) = IIf(IsNull(TmpRs!LocationName), "", TmpRs!LocationName)
         .TextMatrix(iu, 3) = IIf(IsNull(TmpRs!Commodity), "", TmpRs!Commodity)
         .TextMatrix(iu, 4) = IIf(IsNull(TmpRs!RRNNo), "", TmpRs!RRNNo)
         .TextMatrix(iu, 5) = IIf(IsNull(TmpRs!RRNID), "", TmpRs!RRNID)
         .TextMatrix(iu, 6) = IIf(IsNull(TmpRs!RRNDate), "", Format(TmpRs!RRNDate, "dd-MMM-yyyy"))
         .TextMatrix(iu, 7) = IIf(IsNull(TmpRs!DepositoryID), "", TmpRs!DepositoryID)
         .TextMatrix(iu, 8) = IIf(IsNull(TmpRs!DepositoryName), "", TmpRs!DepositoryName)
         .TextMatrix(iu, 9) = IIf(IsNull(TmpRs!ClientID), "", TmpRs!ClientID)
         .TextMatrix(iu, 10) = IIf(IsNull(TmpRs!ClientName), "", TmpRs!ClientName)
         .TextMatrix(iu, 11) = IIf(IsNull(TmpRs!ISINID), "", TmpRs!ISINID)
         .TextMatrix(iu, 12) = IIf(IsNull(TmpRs!CMPName), "", TmpRs!CMPName)
         .TextMatrix(iu, 13) = IIf(IsNull(TmpRs!Grade), "", TmpRs!Grade)
         .TextMatrix(iu, 14) = IIf(IsNull(TmpRs!FED), "", Format(TmpRs!FED, "dd-MMM-yyyy"))
         .TextMatrix(iu, 15) = IIf(IsNull(TmpRs!RRNQty), "", TmpRs!RRNQty)
         .TextMatrix(iu, 16) = IIf(IsNull(TmpRs!OutOfSystemQty), "", TmpRs!OutOfSystemQty)
         .TextMatrix(iu, 17) = IIf(IsNull(TmpRs!DeliveredQty), "", TmpRs!DeliveredQty)
         .TextMatrix(iu, 18) = IIf(IsNull(TmpRs!BalanceQty), "", TmpRs!BalanceQty)
         .TextMatrix(iu, 19) = IIf(IsNull(TmpRs!AdjstQty), "", TmpRs!AdjstQty)
         .TextMatrix(iu, 20) = IIf(IsNull(TmpRs!ConvertedClient), "", TmpRs!ConvertedClient)
         .TextMatrix(iu, 21) = IIf(IsNull(TmpRs!Remarks), "", TmpRs!Remarks)
         .TextMatrix(iu, 22) = IIf(IsNull(TmpRs!Flag), "P", TmpRs!Flag)
         .TextMatrix(iu, 23) = IIf(IsNull(TmpRs!CommTrackFlag), 0, TmpRs!CommTrackFlag)
         If LCase(.TextMatrix(iu, 22)) = "p" Then
            .TextMatrix(iu, 22) = "Pending"
         ElseIf LCase(.TextMatrix(iu, 22)) = "a" Then
            .TextMatrix(iu, 22) = "Approved"
         End If
         TmpRs.MoveNext
      Next
   End With
'   MsgBox "Done!!!"
'Else
'   MsgBox "No record found!!!"
End If

End Sub

Private Sub Grid_Head()
CmdExcel.Enabled = False
With MSHFlexGrid1
   .Clear
   .Rows = 2
   .Cols = 24
   .TextMatrix(0, 0) = "ConvertedID"
   .TextMatrix(0, 1) = "PWR Date"
   .TextMatrix(0, 2) = "LocationName"
   .TextMatrix(0, 3) = "Commodity"
   .TextMatrix(0, 4) = "RRNNo"
   .TextMatrix(0, 5) = "RRNID"
   .TextMatrix(0, 6) = "RRNDate"
   .TextMatrix(0, 7) = "DPID"
   .TextMatrix(0, 8) = "DPName"
   .TextMatrix(0, 9) = "ClientID"
   .TextMatrix(0, 10) = "ClientName"
   .TextMatrix(0, 11) = "ISINNo"
   .TextMatrix(0, 12) = "CMP"
   .TextMatrix(0, 13) = "Grade"
   .TextMatrix(0, 14) = "FEDDate"
   .TextMatrix(0, 15) = "RRNQty"
   .TextMatrix(0, 16) = "Adjusted Qty by CMG"
   .TextMatrix(0, 17) = "Qty Delivered in StoreMan"
   .TextMatrix(0, 18) = "BalanceQty"
   .TextMatrix(0, 19) = "Adjusted Qty "
   .TextMatrix(0, 20) = "Converted Client"
   .TextMatrix(0, 21) = "Remarks"
   .TextMatrix(0, 22) = "Status"
   .TextMatrix(0, 23) = "Commtrack Flag"
   
   .RowHeight(0) = 400
End With
End Sub

Private Sub DeleteCmd_Click()
If LCase(TxtStatus.Text) = "approved" Then
   MsgBox "ConvertedID is approved, you can not delete!!!"
   Exit Sub
End If

Dim ans As Variant
On Error GoTo msgError
ans = MsgBox("Are you Sure you want to Delete this Record?", vbYesNo)
If ans = vbYes Then
   tmp = " Update Mst_RRN set DeliveredQty = DeliveredQty - " & Val(TxtAdjQty.Text) + Val(MSHFlexGrid2.TextMatrix(0, 7)) & " Where RRNID = " & TxtRRNID.Text
   Call TmpTable
   
   For I = 2 To MSHFlexGrid2.Rows - 2
      tmp = " Update Mst_GSL Set "
      If Val(MSHFlexGrid2.TextMatrix(I, 5)) = 0 Then
         tmp = tmp & " Flag = Flag1,"
      End If
      tmp = tmp & " BalanceBags = BalanceBags + " & Val(MSHFlexGrid2.TextMatrix(I, 8)) & " , "
      tmp = tmp & " BalanceWeight = BalanceWeight + " & Val(MSHFlexGrid2.TextMatrix(I, 7)) & " ,"
      tmp = tmp & " CDTFBags = CDTFBags - " & Val(MSHFlexGrid2.TextMatrix(I, 8)) & ", "
      tmp = tmp & " CDTFWeight = CDTFWeight - " & Val(MSHFlexGrid2.TextMatrix(I, 7)) & ""
      tmp = tmp & " Where GSLID = " & MSHFlexGrid2.TextMatrix(I, 6) & " And SM_Prefix = '" & TxtSM_Prefix.Text & "'"
      Call TmpTable
   Next

   tmp = " Delete from Mst_GSL Where SM_Prefix = '" & TxtSM_Prefix.Text & "' And  GSLID in (Select NewGSLID from Txn_RRNPWRGSL Where ConvertedID = " & Val(TxtConvertedID.Text) & " )"
   Call TmpTable
   
   tmp = " Delete from Txn_RRNPWRGSL Where ConvertedID = " & Val(TxtConvertedID.Text)
   Call TmpTable

   RsTxn_RRNPWR.Delete
   RsTxn_RRNPWR.Update

   If RsTxn_RRNPWR.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsTxn_RRNPWR.MoveNext
   If RsTxn_RRNPWR.EOF() Then
      RsTxn_RRNPWR.MoveLast
   End If
   Call Indata
End If
Exit Sub
msgError:
   MsgBox "Child Record Present "
   RsTxn_RRNPWR.CancelUpdate
End Sub

Private Sub EditCmd_Click()

If LCase(TxtStatus.Text) = "approved" Then
   MsgBox "ConvertedID is approved, you can not edit!!!"
   Exit Sub
End If
Edit_Flg = "E"
Frame1.Enabled = True
Call GButtonLock(Me, False)

TxtStatus.Visible = False
CmbStatus.Visible = True
TxtAdjQty.Locked = True
TxtRRNNo.Locked = True
TxtPWRDate.Enabled = False
Call Grid_HeadRRNNCDEXInvoice

End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub Form_Activate()
If Me.WindowState = vbMaximized Then
   SetScrollBars
End If
End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Call GButtonLockAD(Me)
   MsgBox "Access Denied !!!!!!!!!"
   Exit Sub
End If

Call Grid_Head
Call Grid_HeadRRNNCDEXInvoice


SFrom.Value = Date - 30
STo.Value = Date
TxtSRRNNo.Text = ""
TxtSISINNo.Text = ""
CmbSLocation.Text = ""
CmbSCmp.Text = ""
CmbSCommodity.Text = ""

Call TableTxn_RRNPWR(" Order by ConvertedID ")

If RsTxn_RRNPWR.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If

Call GButtonLock(Me, True)
Call Indata

End Sub

Private Sub Form_Resize()
If Me.WindowState = vbMaximized Then
   SetScrollBars
End If
End Sub

Private Function GetGSLID() As Double
Dim Retval As Double
tmp = "Select max(GSLID) from Mst_GSL where SM_Prefix = '" & TxtSM_Prefix.Text & "'"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetGSLID = Retval
End Function

Private Sub MSHFlexGrid1_Click()
Dim x As Variant
Dim mCaId As Variant

x = MSHFlexGrid1.RowSel
If MSHFlexGrid1.TextMatrix(x, 0) = "" Then Exit Sub
mCaId = MSHFlexGrid1.TextMatrix(x, 0)
RsTxn_RRNPWR.MoveFirst

RsTxn_RRNPWR.MoveFirst
RsTxn_RRNPWR.Find "ConvertedID = " & Val(mCaId) & ""
If RsTxn_RRNPWR.EOF Then
   RsTxn_RRNPWR.MoveLast
End If
Call Indata
Frame0.Visible = False
End Sub

Private Sub MSHFlexGrid2_Click()
If SaveCmd.Enabled = False Or Edit_Flg <> "A" Then Exit Sub
mColsel = MSHFlexGrid2.ColSel
mRowsel = MSHFlexGrid2.RowSel

If mColsel = 0 And mRowsel <> MSHFlexGrid2.Rows - 1 Then
   If MSHFlexGrid2.TextMatrix(mRowsel, mColsel) = strUnChecked Then
      MSHFlexGrid2.TextMatrix(mRowsel, mColsel) = strChecked
      MSHFlexGrid2.TextMatrix(0, 7) = Val(MSHFlexGrid2.TextMatrix(0, 7)) - Val(MSHFlexGrid2.TextMatrix(mRowsel, 7))
      MSHFlexGrid2.TextMatrix(0, 8) = Val(MSHFlexGrid2.TextMatrix(0, 8)) - Val(MSHFlexGrid2.TextMatrix(mRowsel, 8))
      MSHFlexGrid2.TextMatrix(mRowsel, 7) = MSHFlexGrid2.TextMatrix(mRowsel, 5)
      MSHFlexGrid2.TextMatrix(mRowsel, 8) = MSHFlexGrid2.TextMatrix(mRowsel, 4)
      MSHFlexGrid2.TextMatrix(0, 7) = Val(MSHFlexGrid2.TextMatrix(0, 7)) + Val(MSHFlexGrid2.TextMatrix(mRowsel, 7))
      MSHFlexGrid2.TextMatrix(0, 8) = Val(MSHFlexGrid2.TextMatrix(0, 8)) + Val(MSHFlexGrid2.TextMatrix(mRowsel, 8))
      
      
   Else
      MSHFlexGrid2.TextMatrix(mRowsel, mColsel) = strUnChecked
      MSHFlexGrid2.TextMatrix(0, 7) = Val(MSHFlexGrid2.TextMatrix(0, 7)) - Val(MSHFlexGrid2.TextMatrix(mRowsel, 7))
      MSHFlexGrid2.TextMatrix(0, 8) = Val(MSHFlexGrid2.TextMatrix(0, 8)) - Val(MSHFlexGrid2.TextMatrix(mRowsel, 8))
      MSHFlexGrid2.TextMatrix(mRowsel, 7) = ""
      MSHFlexGrid2.TextMatrix(mRowsel, 8) = ""
   End If
ElseIf MSHFlexGrid2.TextMatrix(mRowsel, 0) = strChecked Then
   If mColsel = 7 And mRowsel <> MSHFlexGrid2.Rows - 1 Then
      Call VisibleEnterTxt
   End If
End If
End Sub

Private Sub VisibleEnterTxt()
With TxtEnterData
   .Text = MSHFlexGrid2.TextMatrix(mRowsel, mColsel)
   .Left = MSHFlexGrid2.Left + MSHFlexGrid2.CellLeft
   .Top = MSHFlexGrid2.Top + MSHFlexGrid2.CellTop
   .Width = MSHFlexGrid2.CellWidth - 10
   .Height = MSHFlexGrid2.CellHeight - 10
   .Visible = True
   .ZOrder
   .SetFocus
End With
End Sub

Private Sub MSHFlexGrid2_Scroll()
TxtEnterData.Text = ""
TxtEnterData.Visible = False
End Sub

Private Function CheckGSLFill() As Boolean
CheckGSLFill = True
If MSHFlexGrid2.Rows = 3 Then
   MsgBox "GSL Detail not found!!!"
   CheckGSLFill = False
   Exit Function
End If
If Val(MSHFlexGrid2.TextMatrix(0, 7)) = 0 Then
   MsgBox "Enter atleast one GSL Detail in grid!!!"
   CheckGSLFill = False
   Exit Function
End If

End Function

Private Sub SaveCmd_Click()
If TxtPWRDate.Value > Date Then
   MsgBox "Future PWRDate not allowed!!!"
   TxtPWRDate.SetFocus
   Exit Sub
End If

If CmbLocation.Text = "" Then
   MsgBox "Blank Location not allowed!!!"
   CmbLocation.SetFocus
   Exit Sub
End If

If CmbCommodityID.Text = "" Then
   MsgBox "Blank Commodity not allowed!!!"
   CmbCommodityID.SetFocus
   Exit Sub
End If

If TxtRRNNo.Text = "" Then
   MsgBox "Blank RRNNo not allowed!!!"
   TxtRRNNo.SetFocus
   Exit Sub
End If

If Trim(CmbStatus.Text) = "" Then
   MsgBox "Blank Status not allowed!!!"
   CmbStatus.SetFocus
   Exit Sub
End If

If mCommtrackFlag = 1 Then
   If ChkGSLforPWR = False Then
      MsgBox "For Commtrack RRN all lots in grid should be transfered!!!"
      Exit Sub
   End If
End If

If Edit_Flg = "A" Then
   If Format(TxtPWRDate.Value, "yyyy-mm-dd") < Format(CDate(TxtRRNDate), "yyyy-mm-dd") Then
      MsgBox "PWR Date must be greater than RRN Date!!!"
      TxtPWRDate.SetFocus
      Exit Sub
   End If
   
   If Format(TxtPWRDate.Value, "yyyy-mm-dd") < Format(mBODDate, "yyyy-mm-dd") Then
      MsgBox "PWR Date must be greater than BOD Date(" & Format(mBODDate, "dd-mmm-yyyy") & ")!!!"
      TxtPWRDate.SetFocus
      Exit Sub
   End If
End If

If TxtAdjQty.Text = "" Then
   MsgBox "Blank Adjusted Qty not allowed!!!"
   TxtAdjQty.SetFocus
   Exit Sub
End If

'If CmbConvertedClient.Text = "" Then
'   MsgBox "Blank Converted Client Name not allowed!!!"
'   CmbConvertedClient.SetFocus
'   Exit Sub
'End If

If Edit_Flg = "A" Then
   If CheckGSLFill = False Then
      SSTab1.SetFocus
      Exit Sub
   End If
   
   If Format(Val(TxtAdjQty.Text) + Val(MSHFlexGrid2.TextMatrix(0, 7)), "##0.000") <> Format(Val(TxtBalanceQty.Text), "##0.000") Then
      MsgBox "Sum of 'Conversion margin Qty' and 'Converted Qty' is not equal to 'RRN Balance Qty'!!! Difference is " & Format(Val(TxtBalanceQty.Text) - (Val(TxtAdjQty.Text) + Val(MSHFlexGrid2.TextMatrix(0, 7))), "##0.000")
      TxtAdjQty.SetFocus
      Exit Sub
   End If
End If

If LCase(CmbStatus.Text) = "approved" Then
   If CheckNCDEXRate = False Then
      MsgBox "No Rate updated in NCDEX rate card master. Please contact CMG Dept for updating NCDEX rate card master!!!"
      Exit Sub
   End If
   If CheckRateCards = False Then
      MsgBox "No Rate updated in Deposit Contract. Please update Deposit Contract !!!"
      Exit Sub
   End If

End If

If MsgBox("Are you sure you want to save?", vbYesNo) = vbNo Then Exit Sub

If Edit_Flg = "A" Then
   RsTxn_RRNPWR.AddNew
   RsTxn_RRNPWR!ConvertedID = GetMaxConvertedID()
   TxtConvertedID = RsTxn_RRNPWR!ConvertedID
   RsTxn_RRNPWR!G_UID = GetGUID
Else
   RsTxn_RRNPWR.MoveFirst
   RsTxn_RRNPWR.Find "ConvertedID  = " & TxtConvertedID & ""
   If RsTxn_RRNPWR.EOF Then
      Exit Sub
   End If
End If
RsTxn_RRNPWR!PWRDate = TxtPWRDate.Value
RsTxn_RRNPWR!LocationID = mLocationID
RsTxn_RRNPWR!CommodityID = mCommodityID
RsTxn_RRNPWR!RRNID = TxtRRNID
RsTxn_RRNPWR!ConvertedClientIDRow = mConvertedClientIDRow
RsTxn_RRNPWR!AdjstQty = Val(TxtAdjQty.Text)
RsTxn_RRNPWR!Remarks = Trim(TxtRemarks.Text)
RsTxn_RRNPWR!Flag = Left$(CmbStatus.Text, 1)

If IsNull(RsTxn_RRNPWR!CreatedBy) = True Or RsTxn_RRNPWR!CreatedBy = "" Then
   RsTxn_RRNPWR!CreatedBy = Gen_UserLoginID
   RsTxn_RRNPWR!CreatedDate = Now
Else
   RsTxn_RRNPWR!UpdatedBy = Gen_UserLoginID
   RsTxn_RRNPWR!UpdatedDate = Now
End If
RsTxn_RRNPWR.Update

Call TableTxn_RRNPWRGSL(" Where ConvertedID = " & Val(TxtConvertedID))

If Edit_Flg = "A" Then
'If LCase(CmbStatus.Text) = "approved" Then
   ''Update RRN Master with BalanceQty(sum of adjst qty and GSl convered qty
   tmp = " Update Mst_RRN set DeliveredQty = DeliveredQty + " & Val(TxtBalanceQty.Text) & " Where RRNID = " & TxtRRNID.Text
   Call TmpTable
   Call TableMst_GSL(" Where SM_Prefix = '" & TxtSM_Prefix.Text & "' And ExchangeTypeID = 2")
'End If

   For I = 2 To MSHFlexGrid2.Rows - 2
      If MSHFlexGrid2.TextMatrix(I, 0) = strChecked And Val(MSHFlexGrid2.TextMatrix(I, 7)) > 0 Then
         'If LCase(CmbStatus.Text) = "approved" Then
            Call UpdateGSLDetail(I)
         'End If
         If Edit_Flg = "A" Then
            RsTxn_RRNPWRGSL.AddNew
            RsTxn_RRNPWRGSL!ConvertedGSLID = GetMaxConvertedGSLID()
            RsTxn_RRNPWRGSL!G_UID = GetGUID
         Else
            RsTxn_RRNPWRGSL.MoveFirst
            RsTxn_RRNPWRGSL.Find "ConvertedGSLID = " & Val(MSHFlexGrid2.TextMatrix(I, 9))
         End If
         RsTxn_RRNPWRGSL!ConvertedID = TxtConvertedID
         RsTxn_RRNPWRGSL!GSLID = MSHFlexGrid2.TextMatrix(I, 6)
         RsTxn_RRNPWRGSL!SM_Prefix = TxtSM_Prefix.Text
         RsTxn_RRNPWRGSL!AdjstQty = MSHFlexGrid2.TextMatrix(I, 7)
         RsTxn_RRNPWRGSL!AdjstBags = MSHFlexGrid2.TextMatrix(I, 8)
         'If LCase(CmbStatus.Text) = "approved" Then
            RsTxn_RRNPWRGSL!NewGSLID = mNewGSLID '''' ''remaining
         'End If
         If IsNull(RsTxn_RRNPWRGSL!CreatedBy) = True Or RsTxn_RRNPWRGSL!CreatedBy = "" Then
            RsTxn_RRNPWRGSL!CreatedBy = Gen_UserLoginID
            RsTxn_RRNPWRGSL!CreatedDate = Now
         Else
            RsTxn_RRNPWRGSL!UpdatedBy = Gen_UserLoginID
            RsTxn_RRNPWRGSL!UpdatedDate = Now
         End If
         RsTxn_RRNPWRGSL.Update
         ''Update BalanceBags, BalanceWeight, CDTFBags, CDTFWeight in Mst_GSL, Flag = 'E' When BalanceBags = 0
      End If
   Next
End If

RsTxn_RRNPWR.Requery
RsTxn_RRNPWR.MoveFirst
RsTxn_RRNPWR.Find "ConvertedID = " & TxtConvertedID
If RsTxn_RRNPWR.EOF Then
   RsTxn_RRNPWR.MoveLast
End If

If LCase(CmbStatus.Text) = "approved" Then

   Call TableTxn_CDTFBOE(" Where SM_Prefix='" & TxtSM_Prefix & "' And CDTFNo=" & Val(TxtConvertedID) & " And TxnType='P' And ((Amount + Premium) - Discount) > 0 ")
          
   If RsTxn_CDTFBOE.RecordCount = 0 Then
      Call CreateRRNNCDEXInvoice
      If mRejInvIDs <> "" Then
         MsgBox "PWR Invoice " & mRejInvIDs & " generated, please confirm payment!!!"
         Exit Sub
      End If
   End If
End If

Call Indata


End Sub

Private Sub UpdateGSLDetail(i_ As Variant)
tmp = " Update Mst_GSL Set "
'If Format(Val(MSHFlexGrid2.TextMatrix(i_, 5)), "###0.000") - Format(Val(MSHFlexGrid2.TextMatrix(i_, 7)), "###0.000") = 0 Then
'If Format(Val(MSHFlexGrid2.TextMatrix(i_, 5)), "###0.000") = 0 Then
If Format(Val(MSHFlexGrid2.TextMatrix(i_, 4)), "###0") - Format(Val(MSHFlexGrid2.TextMatrix(i_, 8)), "###0") = 0 Then
   tmp = tmp & " Flag1 = Flag, Flag = 'E',"
End If
tmp = tmp & " BalanceBags = BalanceBags - " & Val(MSHFlexGrid2.TextMatrix(i_, 8)) & " , "
tmp = tmp & " BalanceWeight = BalanceWeight - " & Val(MSHFlexGrid2.TextMatrix(i_, 7)) & " ,"
tmp = tmp & " CDTFBags = CDTFBags + " & Val(MSHFlexGrid2.TextMatrix(i_, 8)) & ", "
tmp = tmp & " CDTFWeight = CDTFWeight + " & Val(MSHFlexGrid2.TextMatrix(i_, 7)) & ""
tmp = tmp & " Where GSLID = " & MSHFlexGrid2.TextMatrix(i_, 6) & " And SM_Prefix = '" & TxtSM_Prefix.Text & "'"
Call Tmp2Table

tmp = " Select * from Mst_GSL Where GSLID = " & MSHFlexGrid2.TextMatrix(i_, 6) & " And SM_Prefix = '" & TxtSM_Prefix.Text & "'"
Call Tmp2Table

If TmpRs2.RecordCount > 0 Then
   RsMst_GSL.AddNew
   RsMst_GSL!GSLID = GetGSLID
   mNewGSLID = RsMst_GSL!GSLID
   RsMst_GSL!G_UID = GetGUID
   RsMst_GSL!SM_Prefix = TxtSM_Prefix
   RsMst_GSL!CMPID = TmpRs2!CMPID
   If mCommodityID <> 31 And mCommodityID <> 114 Then
      RsMst_GSL!CommodityID = mCommodityID
   Else
      RsMst_GSL!CommodityID = TmpRs2!CommodityID
   End If
   RsMst_GSL!GodownID = TmpRs2!GodownID
   RsMst_GSL!StackNo = TmpRs2!StackNo
   RsMst_GSL!LotNo = TmpRs2!LotNo & "-Con" & TxtConvertedID.Text
   RsMst_GSL!DepositeBags = Val(MSHFlexGrid2.TextMatrix(i_, 8))
   RsMst_GSL!DepositeWeight = Val(MSHFlexGrid2.TextMatrix(i_, 7))
   RsMst_GSL!BalanceBags = Val(MSHFlexGrid2.TextMatrix(i_, 8))
   RsMst_GSL!BalanceWeight = Val(MSHFlexGrid2.TextMatrix(i_, 7))
   RsMst_GSL!CDTFBags = 0
   RsMst_GSL!CDTFWeight = 0
   RsMst_GSL!WHRBags = 0
   RsMst_GSL!WHRWeight = 0
   RsMst_GSL!ExchangeTypeID = 1
   RsMst_GSL!ClientIDRow = mConvertedClientIDRow
   RsMst_GSL!PledgeBags = 0
   RsMst_GSL!PledgeWeight = 0
   RsMst_GSL!SpillageFlag = 0
   RsMst_GSL!DepositAdjWt = 0
   RsMst_GSL!WithdrawalAdjWt = 0
   RsMst_GSL!Flag = "O"
   RsMst_GSL!CdtfNo = -Val(TxtConvertedID.Text)
   RsMst_GSL!CDTFDate = TxtPWRDate.Value 'Date
   RsMst_GSL!SCMFlag = 0
   RsMst_GSL!Madeupflag = 0
   RsMst_GSL!CreatedBy = Gen_UserLoginID
   RsMst_GSL!CreatedDate = Now
   RsMst_GSL.Update
End If

End Sub

Private Sub Indata()
Edit_Flg = ""
Call Grid_HeadRRNNCDEXInvoice
Check1.Value = vbUnchecked
Check1.Enabled = False
Call GButtonLock(Me, True)
Frame1.Enabled = False

CmbCommodityID.Visible = False
'CmbConvertedClient.Visible = False
CmbLocation.Visible = False
CmbStatus.Visible = False

TxtStatus.Visible = True
TxtCommodityID.Visible = True
'TxtConvertedClient.Visible = True
TxtLocation.Visible = True

TxtConvertedID = RsTxn_RRNPWR!ConvertedID
TxtPWRDate.Value = IIf(IsNull(RsTxn_RRNPWR!PWRDate), Date, RsTxn_RRNPWR!PWRDate)
TxtAdjQty.Text = IIf(IsNull(RsTxn_RRNPWR!AdjstQty), 0, RsTxn_RRNPWR!AdjstQty)
TxtRemarks.Text = IIf(IsNull(RsTxn_RRNPWR!Remarks), "", RsTxn_RRNPWR!Remarks)
TxtStatus.Text = IIf(IsNull(RsTxn_RRNPWR!Flag), "P", RsTxn_RRNPWR!Flag)
If LCase(TxtStatus.Text) = "p" Then
   TxtStatus.Text = "Pending"
Else
   TxtStatus.Text = "Approved"
End If
CmbStatus.Text = TxtStatus.Text

TxtCreated.Text = IIf(IsNull(RsTxn_RRNPWR!CreatedBy), "", RsTxn_RRNPWR!CreatedBy) & " :- " & IIf(IsNull(RsTxn_RRNPWR!CreatedDate), "", RsTxn_RRNPWR!CreatedDate)
TxtUpdated.Text = IIf(IsNull(RsTxn_RRNPWR!UpdatedBy), "", RsTxn_RRNPWR!UpdatedBy) & " :- " & IIf(IsNull(RsTxn_RRNPWR!UpdatedDate), "", RsTxn_RRNPWR!UpdatedDate)

tmp = " Select TRP.LocationID, TRP.CommodityID, TRP.RRNID, ML.LocationName, ML.SM_Prefix, MCom.Commodity, RRN.RRNNo, TRP.ConvertedClientIDRow, MC.ClientName from Txn_RRNPWR TRP "
tmp = tmp & " Inner Join Mst_Location ML on ML.LocationID = TRP.LocationID"
tmp = tmp & " Inner Join Mst_Commodity MCom on MCom.CommodityID = TRP.CommodityID "
tmp = tmp & " Inner Join Mst_RRN RRN on RRN.RRNID = TRP.RRNID"
tmp = tmp & " Inner Join Mst_Client MC on MC.ClientIDRow = TRP.ConvertedClientIDRow"
tmp = tmp & " Where TRP.ConvertedID = " & RsTxn_RRNPWR!ConvertedID
Call Tmp3Table

If TmpRs3.RecordCount > 0 Then
   mLocationID = IIf(IsNull(TmpRs3!LocationID), 0, TmpRs3!LocationID)
   TxtLocation.Text = IIf(IsNull(TmpRs3!LocationName), "", TmpRs3!LocationName)
   CmbLocation.Text = TxtLocation.Text
   TxtSM_Prefix.Text = IIf(IsNull(TmpRs3!SM_Prefix), "", TmpRs3!SM_Prefix)
   
   mCommodityID = IIf(IsNull(TmpRs3!CommodityID), 0, TmpRs3!CommodityID)
   TxtCommodityID.Text = IIf(IsNull(TmpRs3!Commodity), "", TmpRs3!Commodity)
   CmbCommodityID.Text = TxtCommodityID.Text
   
   mConvertedClientIDRow = IIf(IsNull(TmpRs3!ConvertedClientIDRow), 0, TmpRs3!ConvertedClientIDRow)
   TxtConvertedClient.Text = IIf(IsNull(TmpRs3!ClientName), "", TmpRs3!ClientName)
   CmbConvertedClient.Text = TxtConvertedClient.Text
   
   TxtRRNNo.Text = IIf(IsNull(TmpRs3!RRNNo), "", TmpRs3!RRNNo)
   Call GetRRNDetail
   If TxtISINID.Text <> "" Then
      Call FillGSLDetail
   End If
End If
End Sub

Private Function GetMaxConvertedID() As Double
Dim Retval As Variant
Retval = 0
tmp = "select max(ConvertedID) from Txn_RRNPWR "
Call Tmp4Table

If TmpRs4.RecordCount > 0 Then
   Retval = IIf(IsNull(TmpRs4.Fields(0)), 0, TmpRs4.Fields(0))
End If
TmpRs4.Close
Retval = Retval + 1
GetMaxConvertedID = Retval
End Function

Private Function GetMaxConvertedGSLID() As Double
Dim Retval As Variant
Retval = 0
tmp = "select max(ConvertedGSLID) from Txn_RRNPWRGSL "
Call Tmp4Table
If TmpRs4.RecordCount > 0 Then
   Retval = IIf(IsNull(TmpRs4.Fields(0)), 0, TmpRs4.Fields(0))
End If
TmpRs4.Close
Retval = Retval + 1
GetMaxConvertedGSLID = Retval
End Function

Private Sub SearchCmd_Click()
Dim x As Integer

If SearchCmd.Caption = "Cancel" Then
   If RsTxn_RRNPWR.RecordCount = 0 Then Exit Sub

   Call GButtonLock(Me, True)
   If RsTxn_RRNPWR.EOF Then
      RsTxn_RRNPWR.MoveFirst
   End If
   Call Indata
   Exit Sub
End If

Frame0.Height = Frame4.Height - 200
Frame0.Left = Frame1.Left
Frame0.Top = Frame1.Top
Frame0.Width = Frame1.Width
Frame0.Visible = True
'MSHFlexGrid1.Clear
MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 1600

End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsTxn_RRNPWR.MoveNext
If RsTxn_RRNPWR.EOF Then
   RsTxn_RRNPWR.MoveLast
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
RsTxn_RRNPWR.MovePrevious
If RsTxn_RRNPWR.BOF Then
   RsTxn_RRNPWR.MoveFirst
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
RsTxn_RRNPWR.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsTxn_RRNPWR.MoveLast
Call Indata
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = Fasle
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Txn_RRNPWR.xls")
Call Xls_Detail_Rpt
End Sub

Public Sub Xls_Detail_Rpt()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Integer
Dim mCol As Integer
MsgBox ("Wait till Next Message!!!")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Txn_RRNPWR.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid1.Rows - 1
   mRow = I + 1
   For C = 0 To MSHFlexGrid1.Cols - 1
      oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
   Next
Next
oWB.Save
MsgBox ("Excel File Created!!!")
oXL.Visible = True
End Sub


Private Sub TxtAdjQty_LostFocus()
If TxtAdjQty = "" Then Exit Sub

If IsNumeric(TxtAdjQty) = False Or Val(TxtAdjQty) < 0 Then
   MsgBox "Invalid Numeric format!!!"
   TxtAdjQty.SetFocus
   Exit Sub
End If
TxtAdjQty.Text = Format(Val(TxtAdjQty.Text), "##0.000")
End Sub

Private Sub TxtEnterData_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
   If CheckTxtEnterData = False Then Exit Sub
   MSHFlexGrid2.SetFocus
   DoEvents
   If MSHFlexGrid2.row < MSHFlexGrid2.Rows - 2 Then
      MSHFlexGrid2.row = MSHFlexGrid2.row + 1
      mRowsel = MSHFlexGrid2.row
      mColsel = MSHFlexGrid2.Col
      If mRowsel <> MSHFlexGrid2.Rows - 1 And MSHFlexGrid2.TextMatrix(mRowsel, 0) = strChecked Then
         'Call MSHFlexGrid2_Click
         Call VisibleEnterTxt
      End If
   Else
      Call MSHFlexGrid2_Click
   End If
ElseIf KeyCode = vbKeyUp Then
   If CheckTxtEnterData = False Then Exit Sub
   MSHFlexGrid2.SetFocus
   DoEvents
   If MSHFlexGrid2.row > 2 Then
      MSHFlexGrid2.row = MSHFlexGrid2.row - 1
      mRowsel = MSHFlexGrid2.row
      mColsel = MSHFlexGrid2.Col
      If mRowsel <> 1 And MSHFlexGrid2.TextMatrix(mRowsel, 0) = strChecked Then
         'Call MSHFlexGrid2_Click
         Call VisibleEnterTxt
      End If
   Else
      Call MSHFlexGrid2_Click
   End If
End If
End Sub

Private Sub TxtEnterData_LostFocus()
If Trim(TxtEnterData.Text) = "" Then
   MSHFlexGrid2.TextMatrix(0, 7) = Val(MSHFlexGrid2.TextMatrix(0, 7)) - Val(MSHFlexGrid2.TextMatrix(mRowsel, 7))
   MSHFlexGrid2.TextMatrix(0, 8) = Val(MSHFlexGrid2.TextMatrix(0, 8)) - Val(MSHFlexGrid2.TextMatrix(mRowsel, 8))
   MSHFlexGrid2.TextMatrix(mRowsel, 7) = ""
   MSHFlexGrid2.TextMatrix(mRowsel, 8) = ""
   TxtEnterData.Visible = False
   Exit Sub
End If
If CheckTxtEnterData = False Then
   TxtEnterData.SetFocus
   Exit Sub
Else
   MSHFlexGrid2.TextMatrix(0, mColsel) = Val(MSHFlexGrid2.TextMatrix(0, mColsel)) - Val(MSHFlexGrid2.TextMatrix(mRowsel, mColsel))
   MSHFlexGrid2.TextMatrix(0, mColsel + 1) = Val(MSHFlexGrid2.TextMatrix(0, mColsel + 1)) - Val(MSHFlexGrid2.TextMatrix(mRowsel, mColsel + 1))
   TxtEnterData.Text = Format(Val(TxtEnterData.Text), "##0.000")
   
   MSHFlexGrid2.TextMatrix(mRowsel, mColsel) = TxtEnterData.Text
   If MSHFlexGrid2.TextMatrix(mRowsel, 5) > 0 Then
      MSHFlexGrid2.TextMatrix(mRowsel, mColsel + 1) = CInt((Val(MSHFlexGrid2.TextMatrix(mRowsel, mColsel)) * Val(MSHFlexGrid2.TextMatrix(mRowsel, 4))) / Val(MSHFlexGrid2.TextMatrix(mRowsel, 5)))
   End If
   MSHFlexGrid2.TextMatrix(0, mColsel) = Val(MSHFlexGrid2.TextMatrix(0, mColsel)) + Val(MSHFlexGrid2.TextMatrix(mRowsel, mColsel))
   MSHFlexGrid2.TextMatrix(0, mColsel + 1) = Val(MSHFlexGrid2.TextMatrix(0, mColsel + 1)) + Val(MSHFlexGrid2.TextMatrix(mRowsel, mColsel + 1))
   TxtEnterData.Visible = False
End If
End Sub

Private Function CheckTxtEnterData() As Boolean
CheckTxtEnterData = True
If Trim(TxtEnterData.Text) = "" Then Exit Function
If IsNumeric(TxtEnterData.Text) = False Or Val(TxtEnterData.Text) <= 0 Then
   MsgBox "Enter Positive Numeric  Value!!!"
   TxtEnterData.SetFocus
   CheckTxtEnterData = False
   Exit Function
ElseIf Val(TxtEnterData.Text) > Val(MSHFlexGrid2.TextMatrix(mRowsel, 5)) Then
   MsgBox "You can not enter more than Balance Qty!!!"
   TxtEnterData.SetFocus
   CheckTxtEnterData = False
   Exit Function
End If
End Function

Private Sub TxtRRNNo_GotFocus()
If Edit_Flg <> "A" Then Exit Sub
tmp1 = TxtRRNNo.Text
If Trim(CmbLocation.Text) = "" Then
   MsgBox "First Select Location!!!"
   CmbLocation.SetFocus
   Exit Sub
End If
If Trim(CmbCommodityID.Text) = "" Then
   MsgBox "First Select Commodity!!!"
   CmbCommodityID.SetFocus
   Exit Sub
End If
End Sub

Private Sub ClearRRNDetail()
Call Grid_Head2
TxtRRNID.Text = ""
TxtRRNDate.Text = ""
TxtDPID.Text = ""
TxtDPName.Text = ""
TxtClientID.Text = ""
TxtClientName.Text = ""
TxtISINID.Text = ""
TxtCMPID.Text = ""
TxtGrade.Text = ""
TxtFEDDate.Text = ""
TxtRRNQty.Text = ""
TxtOutOfSystemQty.Text = ""
TxtDeliveredQty.Text = ""
TxtBalanceQty.Text = ""
End Sub

Private Sub TxtRRNNo_LostFocus()
If Edit_Flg <> "A" Then Exit Sub
If Trim(TxtRRNNo.Text) = "" Then
   Call ClearRRNDetail
   Exit Sub
End If

'If tmp1 <> TxtRRNNo.Text Then
   Call ClearRRNDetail
   tmp = " SELECT TC.CxTFNo, TCD.RRNID from Txn_CxTF_ClientDetail TCD"
   tmp = tmp & " INNER JOIN Mst_RRN RRN on RRN.RRNID = TCD.RRNID"
   tmp = tmp & " INNER JOIN Txn_CxTF_Details TC on TC.CxTFNo = TCD.CxTFNo AND TC.SM_Prefix = TCD.SM_Prefix And TC.TxnType = TCD.TxnType"
   tmp = tmp & " Where RRN.RRNNo = '" & TxtRRNNo.Text & "' and isnull(TC.Flag, '') <> 'P'"
   Call Tmp6Table
   If TmpRs6.RecordCount > 0 Then
      MsgBox "CITF No: " & IIf(IsNull(TmpRs6!CxTFNO), "", TmpRs6!CxTFNO) & " is not printed, Print pending CITF and then enter this data!!!"
      TxtRRNNo.SetFocus
      Exit Sub
   End If

   Call GetRRNDetail
   If TxtISINID.Text <> "" Then
      Call FillGSLDetail
   End If
   If Edit_Flg = "A" Then
      If TmpRs6.RecordCount > 0 Then
         mConvertedClientIDRow = IIf(IsNull(TmpRs6!ConvertedClientIDRow), "", TmpRs6!ConvertedClientIDRow)
         TxtConvertedClient.Text = IIf(IsNull(TmpRs6!ConvertedClientName), "", TmpRs6!ConvertedClientName)
         If TxtConvertedClient.Text = "" Then
            MsgBox "Please map Converted Client to above RRN Client!!!"
            TxtRRNNo.SetFocus
            Exit Sub
         End If
       End If
   End If

'End If
End Sub

Private Sub GetRRNDetail()
tmp = "select RRN.RRNID, RRN.RRNNo, RRN.RRNDate, RRN.ISINID, ML.LocationName, CMP.CMPName, CMP.StartDate, CMP.CurrentDate, COM.Commodity, C.ClientID, C.ClientName,"
tmp = tmp & " DP.DepositoryID, DP.DepositoryName, RRN.RRNQty, RRN.DeliveredQty, RRN.OutOfSystemQty, RRN.remarks, RRN.EntryDate, "
tmp = tmp & " ISIN.FED , ISIN.grade, C.ConvertedClientIDRow, CC.CLientName ConvertedClientName,RRN.CommtrackFlag"
tmp = tmp & " from Mst_RRN RRN"
tmp = tmp & " INNER JOIN Mst_ISIN ISIN ON RRN.ISINID = ISIN.ISINID"
tmp = tmp & " INNER JOIN Mst_CMP CMP ON ISIN.CMPID = CMP.CMPID"
tmp = tmp & " INNER JOIN Mst_Location ML ON CMP.LocationID = ML.LocationID"
tmp = tmp & " INNER JOIN Mst_Commodity COM ON ISIN.CommodityID = COM.CommodityID"
tmp = tmp & " INNER JOIN Mst_Client C ON  RRN.ClientIDRow = C.ClientIDRow"
tmp = tmp & " INNER JOIN Mst_Depository dp ON C.DepositoryID = DP.DepositoryID AND DP.ExchangeTypeID = 2"
tmp = tmp & " LEFT JOIN Mst_Client CC ON  C.ConvertedClientIDRow = CC.ClientIDRow"
tmp = tmp & " WHERE ML.LocationID = " & mLocationID & " AND COM.CommodityID = " & mCommodityID & " AND RRN.RRNNo = '" & TxtRRNNo.Text & "'"
If Edit_Flg = "A" Then
   tmp = tmp & " AND isnull(rrn.RRNQty,0) - isnull(rrn.DeliveredQty,0) - isnull(rrn.OutOfSystemQty,0) > 0 " '- isnull(rrn.AdjstQty,0)
'Else
   'tmp = tmp & " AND isnull(rrn.RRNQty,0) - isnull(rrn.DeliveredQty,0) - isnull(rrn.OutOfSystemQty,0) > 0"
End If
Call Tmp6Table

If TmpRs6.RecordCount > 0 Then
   TxtRRNID.Text = IIf(IsNull(TmpRs6!RRNID), "", TmpRs6!RRNID)
'   If Edit_Flg = "A" Then
'      tmp = " Select ConvertedID, RRNID from Txn_RRNPWR Where RRNID = " & TxtRRNID.Text
'      Call Tmp1Table
'      If TmpRs1.RecordCount > 0 Then
'         TxtRRNID.Text = ""
'         MsgBox "RRNNo Already Converted, Check ConvertedID " & IIf(IsNull(TmpRs1!ConvertedID), "", TmpRs1!ConvertedID) & "!!!"
'         TxtRRNNo.SetFocus
'         Exit Sub
'      End If
'   End If
   TxtRRNDate.Text = IIf(IsNull(TmpRs6!RRNDate), "", Format(TmpRs6!RRNDate, "dd-MMM-yyyy"))
   TxtDPID.Text = IIf(IsNull(TmpRs6!DepositoryID), "", TmpRs6!DepositoryID)
   TxtDPName.Text = IIf(IsNull(TmpRs6!DepositoryName), "", TmpRs6!DepositoryName)
   TxtClientID.Text = IIf(IsNull(TmpRs6!ClientID), "", TmpRs6!ClientID)
   TxtClientName.Text = IIf(IsNull(TmpRs6!ClientName), "", TmpRs6!ClientName)
   TxtISINID.Text = IIf(IsNull(TmpRs6!ISINID), "", TmpRs6!ISINID)
   TxtCMPID.Text = IIf(IsNull(TmpRs6!CMPName), "", TmpRs6!CMPName)
   TxtGrade.Text = IIf(IsNull(TmpRs6!Grade), "", TmpRs6!Grade)
   TxtFEDDate.Text = IIf(IsNull(TmpRs6!FED), "", Format(TmpRs6!FED, "dd-MMM-yyyy"))
   TxtRRNQty.Text = IIf(IsNull(TmpRs6!RRNQty), "", TmpRs6!RRNQty)
   TxtOutOfSystemQty.Text = IIf(IsNull(TmpRs6!OutOfSystemQty), "", TmpRs6!OutOfSystemQty)
   TxtDeliveredQty.Text = IIf(IsNull(TmpRs6!DeliveredQty), "", TmpRs6!DeliveredQty)
   TxtBalanceQty.Text = Format(Val(TxtRRNQty.Text) - (Val(TxtDeliveredQty.Text) + Val(TxtOutOfSystemQty.Text)), "########0.000")
   mBODDate = IIf(IsNull(TmpRs6!CurrentDate), Format(TmpRs6!StartDate, "yyyy-mm-dd"), Format(TmpRs6!CurrentDate, "yyyy-mm-dd"))
   mCommtrackFlag = IIf(IsNull(TmpRs6!CommTrackFlag), 0, TmpRs6!CommTrackFlag)
Else
   Call ClearRRNDetail
   MsgBox "No RRNNo found!!!"
   TxtRRNNo.SetFocus
   Exit Sub
End If
End Sub

Private Sub FillGSLDetail()
Call Grid_Head2
tmp = " SELECT GSL.GSLID, MG.GodownNo Godown, GSL.StackNo, GSL.LotNo, GSL.BalanceBags, GSL.Balanceweight "
tmp = tmp & " FROM Mst_GSL GSL "
tmp = tmp & " INNER JOIN Mst_Godown MG ON MG.GodownID = GSL.GodownID "
tmp = tmp & " INNER JOIN Mst_CMP MCMP ON MCMP.CMPID = MG.CMPID"
If Edit_Flg = "A" Then
   tmp = tmp & " WHERE GSL.Flag in ('O', 'R') and GSL.Balanceweight > 0 and GSL.BalanceBags > 0 and ExchangeTypeID = 2 "
Else
   tmp = tmp & " WHERE GSLID in (SELECT GSLID FROM Txn_RRNPWRGSL WHERE ConvertedID = " & Val(TxtConvertedID.Text) & " )"
End If

'Commented on 04 Feb 2014
'tmp = tmp & " and GSL.ISINID = '" & TxtISINID.Text & "' AND GSL.SM_Prefix = '" & TxtSM_Prefix.Text & "' and GSL.CommodityID = " & mCommodityID
If mCommtrackFlag = 0 Then
   If mCommodityID <> 31 And mCommodityID <> 114 Then
      tmp = tmp & " and GSL.ISINID = '" & TxtISINID.Text & "' AND GSL.SM_Prefix = '" & TxtSM_Prefix.Text & "' and GSL.CommodityID = " & mCommodityID
   Else
      tmp = tmp & " and GSL.ISINID = '" & TxtISINID.Text & "' AND GSL.SM_Prefix = '" & TxtSM_Prefix.Text & "' and GSL.CommodityID in (31,114)"
   End If
Else
   If mCommodityID <> 31 And mCommodityID <> 114 Then
      tmp = tmp & " and GSL.ISINID in  (Select Distinct ISINID From Mst_RRN_LotDetail Where RRNID = " & Val(TxtRRNID.Text) & " ) AND GSL.SM_Prefix = '" & TxtSM_Prefix.Text & "' and GSL.CommodityID = " & mCommodityID
   Else
      tmp = tmp & " and GSL.ISINID in  (Select Distinct ISINID From Mst_RRN_LotDetail Where RRNID = " & Val(TxtRRNID.Text) & " ) AND GSL.SM_Prefix = '" & TxtSM_Prefix.Text & "' and GSL.CommodityID in (31,114)"
   End If
End If
Call Tmp5Table

If TmpRs5.RecordCount > 0 Then
   With MSHFlexGrid2
      .Rows = .Rows + TmpRs5.RecordCount
      For j = 2 To TmpRs5.RecordCount + 1
         .TextMatrix(j, 1) = IIf(IsNull(TmpRs5!Godown), "", TmpRs5!Godown)
         .TextMatrix(j, 2) = IIf(IsNull(TmpRs5!StackNo), "", TmpRs5!StackNo)
         .TextMatrix(j, 3) = IIf(IsNull(TmpRs5!LotNo), "", TmpRs5!LotNo)
         .TextMatrix(j, 4) = IIf(IsNull(TmpRs5!BalanceBags), 0, TmpRs5!BalanceBags)
         .TextMatrix(0, 4) = Val(.TextMatrix(j, 4)) + Val(.TextMatrix(0, 4))
         .TextMatrix(j, 5) = Format(IIf(IsNull(TmpRs5!BalanceWeight), 0, TmpRs5!BalanceWeight), "##0.000")
         .TextMatrix(0, 5) = Format(Val(.TextMatrix(j, 5)) + Val(.TextMatrix(0, 5)), "##0.000")
         .TextMatrix(j, 6) = IIf(IsNull(TmpRs5!GSLID), "", TmpRs5!GSLID)
                  
         .row = j
         .Col = 0
         .CellFontName = "Wingdings"
         .CellFontSize = 12
         .CellAlignment = flexAlignCenterCenter
         If Edit_Flg <> "A" Then
            .Text = strChecked
         Else
            .Text = strUnChecked
         End If
         TmpRs5.MoveNext
      Next
   End With
End If
If Edit_Flg <> "A" Then
   tmp = " select ConvertedGSLID, GSLID, AdjstBags, AdjstQty from Txn_RRNPWRGSL Where ConvertedID = " & Val(TxtConvertedID.Text)
   Call Tmp5Table
   If TmpRs5.RecordCount > 0 Then
      For j = 1 To MSHFlexGrid2.Rows - 2
         TmpRs5.MoveFirst
         TmpRs5.Find "GSLID = " & Val(MSHFlexGrid2.TextMatrix(j, 6))
         If Not TmpRs5.EOF Then
            MSHFlexGrid2.TextMatrix(j, 7) = Format(IIf(IsNull(TmpRs5!AdjstQty), 0, TmpRs5!AdjstQty), "##0.000")
            MSHFlexGrid2.TextMatrix(0, 7) = Format(Val(MSHFlexGrid2.TextMatrix(j, 7)) + Val(MSHFlexGrid2.TextMatrix(0, 7)), "##0.000")
            MSHFlexGrid2.TextMatrix(j, 8) = IIf(IsNull(TmpRs5!AdjstBags), "", TmpRs5!AdjstBags)
            MSHFlexGrid2.TextMatrix(0, 8) = Val(MSHFlexGrid2.TextMatrix(j, 8)) + Val(MSHFlexGrid2.TextMatrix(0, 8))
            MSHFlexGrid2.TextMatrix(j, 9) = IIf(IsNull(TmpRs5!ConvertedGSLID), "", TmpRs5!ConvertedGSLID)
         End If
      Next
   End If
End If
End Sub

Private Sub SetScrollBars()
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
HScroll1.Max = Me.ScaleWidth - Frame1.Width
HScroll1.LargeChange = Me.ScaleWidth
HScroll1.SmallChange = Me.ScaleWidth / 5
If Frame4.Height >= Me.Height Then VScroll1.Enabled = True
If Frame4.Height < Me.Height Then VScroll1.Enabled = False
If Frame1.Width >= Me.Width Then HScroll1.Enabled = True
If Frame1.Width < Me.Width Then HScroll1.Enabled = False
End Sub
Private Sub VScroll1_Change()

   Dim l As Double
   Dim a As Double
   Dim x As Double
   
   x = VScroll1.Value
   a = Me.Height - Frame4.Height
   l = (a * x) / 1000
   Frame4.Top = -l

End Sub

Private Sub VScroll1_Scroll()

   Dim l As Double
   Dim a As Double
   Dim x As Double
   
   x = VScroll1.Value
   a = Me.Height - Frame4.Height
   l = (a * x) / 1000
   Frame4.Top = -l

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

Private Function ChkGSLforPWR() As Boolean
Dim Retval As Boolean
Retval = True

For I = 2 To MSHFlexGrid2.Rows - 1
    If MSHFlexGrid2.TextMatrix(I, 1) = "" Then Exit For
    If MSHFlexGrid2.TextMatrix(I, 0) = strUnChecked Then
       Retval = False
       Exit For
    End If
Next
ChkGSLforPWR = Retval

End Function

Private Sub Grid_HeadRRNNCDEXInvoice()
MSHFlexGrid11.Clear
MSHFlexGrid11.Rows = 3
MSHFlexGrid11.FixedCols = 0
MSHFlexGrid11.Cols = 17

MSHFlexGrid11.TextMatrix(1, 0) = "RRN No"
MSHFlexGrid11.TextMatrix(1, 1) = "ISIN No"
MSHFlexGrid11.TextMatrix(1, 2) = "RRN Date"
MSHFlexGrid11.TextMatrix(1, 3) = "FED Date"
MSHFlexGrid11.TextMatrix(1, 4) = "NE Date"
MSHFlexGrid11.TextMatrix(1, 5) = "CITF Date"
MSHFlexGrid11.TextMatrix(1, 6) = "CITF Qty"
MSHFlexGrid11.TextMatrix(1, 7) = "Commodity"
MSHFlexGrid11.TextMatrix(1, 8) = "Storage Charge to be recovered"
MSHFlexGrid11.TextMatrix(1, 9) = "LocationID"
MSHFlexGrid11.TextMatrix(1, 10) = "ClientID"
MSHFlexGrid11.TextMatrix(1, 11) = "CommodityID"
MSHFlexGrid11.TextMatrix(1, 12) = "ServiceTaxID"
MSHFlexGrid11.TextMatrix(1, 13) = "MaxRate"
MSHFlexGrid11.TextMatrix(1, 14) = "GSLID"
MSHFlexGrid11.TextMatrix(1, 15) = "NoOfBags"
MSHFlexGrid11.TextMatrix(1, 16) = "No Of Days"

End Sub

Private Function CheckNCDEXRate() As Boolean
Dim Retval As Boolean
Retval = True
tmp = "Select GodownID,Rate from Mst_NCDEXStorageRateCardDetail MNRC"
tmp = tmp & " inner Join Mst_NCDEXStorageRateCard MNR on MNRC.NSRCID=MNR.NSRCID"
tmp = tmp & " Inner Join Mst_CommodityGroup MCG On MNR.CommodityGroupID = MCG.CommodityGroupID"
tmp = tmp & " Inner Join Mst_Commodity MComm On MCG.CommodityGroupID = MComm.CommodityGroupID"
tmp = tmp & " Where MComm.CommodityID=" & mCommodityID & "  And MNR.LocationID=" & mLocationID
 
Call Tmp2Table

'SELECT GSLID FROM Txn_RRNPWRGSL WHERE ConvertedID = " & Val(TxtConvertedID.Text) & "

tmp = "Select distinct gsl.GodownID From"
tmp = tmp & " Txn_RRNPWRGSL tgsl"
tmp = tmp & " inner join mst_gsl gsl on gsl.GSLID = tgsl.GSLID and gsl.SM_Prefix = tgsl.SM_Prefix"
tmp = tmp & " Where tgsl.ConvertedID = " & Val(TxtConvertedID.Text) & " "

Call Tmp3Table

For I = 1 To TmpRs3.RecordCount
    If TmpRs2.RecordCount > 0 Then TmpRs2.MoveFirst
    TmpRs2.Find "GodownID=" & TmpRs3!GodownID
    If TmpRs2.EOF = False Then
       mRate = TmpRs2!Rate
    Else
       mRate = 0
    End If
    If Val(mRate) = 0 Then
       Retval = False
       Exit For
    End If
    
    TmpRs3.MoveNext
Next
CheckNCDEXRate = Retval
End Function


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


Private Sub CreateRRNNCDEXInvoice()

If TxtPWRDate.Value < CDate("10/01/2013") Then Exit Sub
       
Call TableTxn_CDTFBOE(" Where SM_Prefix='" & TxtSM_Prefix & "' And CDTFNo=" & Val(TxtConvertedID) & " And TxnType='P'")

If RsTxn_CDTFBOE.RecordCount > 0 Then Exit Sub

tmp = "Select CommodityGroupID From Mst_Commodity Where CommodityID=" & mCommodityID
Call TmpTable
If TmpRs.RecordCount > 0 Then
   mCommodityGroupId = TmpRs!CommodityGroupID
End If


tmp = "Select  Max(Rate) as MaxRate From MSt_NcDEXStorageRateCardDetail Where NSRCID in("
tmp = tmp & " Select NSRCID From MSt_NcDEXStorageRateCard Where LocationID = " & mLocationID & " And CommodityGroupID =" & mCommodityGroupId & ")" ' And GodownID in (Select GodownID from MSt_Godown Where CMPID=" & mCMPID & ")"
Call Tmp3Table
If Not TmpRs3.EOF Then
   mMaxRate = TmpRs3!MaxRate
Else
   mRateExist = False
End If

Call TableTxn_CDTFBOE_GSL(" Where SM_Prefix='" & TxtSM_Prefix & "'")


tmp = "Select MR.RRNID,MR.RRNNo,MR.RRNDate,isnull(ISIN.ISINID,ISIN1.ISINID)ISINID,isnull(ISIN.FED,ISIN1.FED)FED,isnull(ISIN.NE_Date,ISIN1.NE_Date)NE_Date, isnull(ISIN.Grade,ISIN1.Grade)Grade,MR.Flag, MR.CommTrackFlag,sum(TRG.AdjstQty)DematWeight,sum(TRG.AdjstBags)NoOfBags,MR.ClientIDRow,TRG.GSLID"
tmp = tmp & " from Txn_RRNPWRGSL TRG"
tmp = tmp & " inner join Txn_RRNPWR TR on TR.ConvertedID=TRG.ConvertedID"
tmp = tmp & " inner join Mst_RRN MR on TR.RRNID=MR.RRNID"
tmp = tmp & " inner join mst_GSL GSL on GSL.GSLID=TRG.GSLID and GSL.SM_prefix=TRG.SM_prefix"
tmp = tmp & " left Join Mst_ISIN as ISIN on ISIN.ISINID = GSl.ISINID"
tmp = tmp & " left Join Mst_ISIN as ISIN1 on ISIN1.ISINID = MR.ISINID"
tmp = tmp & " Where TRG.ConvertedID = " & Val(TxtConvertedID) & " And TRG.SM_prefix='" & TxtSM_Prefix & "'"
tmp = tmp & " Group by MR.RRNID,MR.RRNNo,MR.RRNDate,ISIN.ISINID,ISIN.FED, ISIN.NE_Date, ISIN.Grade,MR.Flag, MR.CommTrackFlag,MR.ClientIDRow,TRG.GSLID,ISIN1.ISINID,ISIN1.FED, ISIN1.NE_Date, ISIN1.Grade"

Call Tmp2Table

Call Grid_HeadRRNNCDEXInvoice

If TmpRs2.RecordCount > 0 Then
   MSHFlexGrid11.Rows = MSHFlexGrid11.Rows + TmpRs2.RecordCount
    For I = 1 To TmpRs2.RecordCount
        mNewRRNDate = TmpRs2!RRNDate + 1
        If TmpRs2!RRNDate < CDate("10/01/2013") Then
           If TmpRs2!CommTrackFlag = "1" Then
              mNewRRNDate = CDate("10/01/2013")
           End If
        End If
        If TmpRs2!CommTrackFlag = "1" Then
           mLastFEDDate = GetLastDateOfMonth(CDate(TmpRs2!FED))
           If TxtPWRDate.Value >= CDate(mLastFEDDate) Then 'CDate(TmpRs2!FED) Then
              mLastFEDDate = GetLastDateOfMonth(CDate(TmpRs2!FED))
              If CDate(mNewRRNDate) > CDate(mLastFEDDate) Then
                 mTotDays1 = DateDiff("d", CDate(mLastFEDDate) + 1, CDate(mNewRRNDate) - 1) + 1
                 mRRNFED = mTotDays1 * mMaxRate * Val(TmpRs2!DematWeight)
                 mNextToFED = CDate(mLastFEDDate) + 1
                 mTotDays = DateDiff("d", CDate(mNewRRNDate), TxtPWRDate.Value) + 1
              
                'Introduced this line for month comparison on 13 Dec 2013 after mail from Mayur Sir and Dharmendra Sir
                 If DateDiff("m", CDate(TmpRs2!FED), TxtPWRDate.Value) > 0 Then
                    mFEDCITF = (mTotDays * mMaxRate * Val(TmpRs2!DematWeight)) * 2
                 Else
                    mFEDCITF = (mTotDays * mMaxRate * Val(TmpRs2!DematWeight))
                 End If
                'mFEDCITF = (mTotDays * mMaxRate * Val(TmpRs2!DematWeight)) * 2
                 mSCToBerecovered = mRRNFED + mFEDCITF
              Else
                 mTotDays1 = DateDiff("d", CDate(mNewRRNDate), CDate(mLastFEDDate)) + 1
                 mRRNFED = mTotDays1 * mMaxRate * Val(TmpRs2!DematWeight)
                 mNextToFED = CDate(mLastFEDDate) + 1
                 mTotDays = DateDiff("d", CDate(mNextToFED), TxtPWRDate.Value) + 1
              
                'Introduced this line for month comparison on 13 Dec 2013 after mail from Mayur Sir and Dharmendra Sir
                 If DateDiff("m", CDate(TmpRs2!FED), TxtPWRDate.Value) > 0 Then
                    mFEDCITF = (mTotDays * mMaxRate * Val(TmpRs2!DematWeight)) * 2
                 Else
                    mFEDCITF = (mTotDays * mMaxRate * Val(TmpRs2!DematWeight))
                 End If
                'mFEDCITF = (mTotDays * mMaxRate * Val(TmpRs2!DematWeight)) * 2
                 mSCToBerecovered = mRRNFED + mFEDCITF
              End If
              
              mTotDays = mTotDays + mTotDays1
           
           
           ElseIf TxtPWRDate.Value <= CDate(mLastFEDDate) Then 'CDate(TmpRs2!FED) Then
           'Single Charges From RRNDate to CITFDate then
              mTotDays = DateDiff("d", CDate(mNewRRNDate), TxtPWRDate.Value) + 1
              mSCToBerecovered = (mTotDays * mMaxRate * Val(TmpRs2!DematWeight))
           End If
        Else
            If TxtPWRDate.Value <= CDate(TmpRs2!FED) Then
               mSCToBerecovered = 0
            ElseIf TxtPWRDate.Value >= CDate(TmpRs2!FED) Then
               tmp = "Select * From Txn_RRNPWR Where RRNID =" & TmpRs2!RRNID
               Call Tmp3Table
               If Not TmpRs3.EOF Then
                  mTotDays1 = DateDiff("d", TmpRs3!PWRDate, CDate(TmpRs2!FED)) + 1
                  mPWRFED = mTotDays1 * mMaxRate * Val(TmpRs2!DematWeight)
                  
                  mTotDays = DateDiff("d", TmpRs3!PWRDate, TxtPWRDate.Value) + 1
                  
                  If DateDiff("m", TmpRs3!PWRDate, TxtPWRDate.Value) > 0 Then
                     mFEDCITF = (mTotDays * mMaxRate * Val(TmpRs2!DematWeight)) * 2
                  Else
                     mFEDCITF = (mTotDays * mMaxRate * Val(TmpRs2!DematWeight))
                  End If
                  'mFEDCITF = (mTotDays * mMaxRate * Val(TmpRs2!DematWeight)) * 2
                  
                  mSCToBerecovered = mPWRFED + mFEDCITF
                  mTotDays = mTotDays + mTotDays1
               Else
                  'mTotDays = DateDiff("d", Format(mFedDate, Gen_DatFormat), CITFDate.Value) + 1
                  mTotDays = DateDiff("d", CDate(TmpRs2!FED), TxtPWRDate.Value) + 1
                  mSCToBerecovered = mTotDays * mMaxRate * Val(TmpRs2!DematWeight)
                  
               End If
            End If
        End If
        r = I + 1
        MSHFlexGrid11.TextMatrix(r, 0) = TmpRs2!RRNNo
        MSHFlexGrid11.TextMatrix(r, 1) = TmpRs2!ISINID
        MSHFlexGrid11.TextMatrix(r, 2) = Format(TmpRs2!RRNDate, "dd-MMM-yyyy")
        MSHFlexGrid11.TextMatrix(r, 3) = Format(TmpRs2!FED, "dd-MMM-yyyy")
        MSHFlexGrid11.TextMatrix(r, 4) = Format(TmpRs2!NE_Date, "dd-MMM-yyyy")
        MSHFlexGrid11.TextMatrix(r, 5) = Format(TxtPWRDate.Value, "dd-MMM-yyyy")
        MSHFlexGrid11.TextMatrix(r, 6) = TmpRs2!DematWeight
        MSHFlexGrid11.TextMatrix(r, 7) = TxtCommodityID
        MSHFlexGrid11.TextMatrix(r, 8) = mSCToBerecovered
        MSHFlexGrid11.TextMatrix(r, 9) = mLocationID
        MSHFlexGrid11.TextMatrix(r, 10) = TmpRs2!ClientIDRow
        MSHFlexGrid11.TextMatrix(r, 11) = mCommodityID
        MSHFlexGrid11.TextMatrix(r, 12) = mServiceTaxID
        MSHFlexGrid11.TextMatrix(r, 13) = mMaxRate
        MSHFlexGrid11.TextMatrix(r, 14) = TmpRs2!GSLID
        MSHFlexGrid11.TextMatrix(r, 15) = TmpRs2!NoofBags
        MSHFlexGrid11.TextMatrix(r, 16) = mTotDays
        MSHFlexGrid11.TextMatrix(0, 8) = Val(MSHFlexGrid11.TextMatrix(0, 8)) + Val(MSHFlexGrid11.TextMatrix(r, 8))
        MSHFlexGrid11.TextMatrix(0, 6) = Val(MSHFlexGrid11.TextMatrix(0, 6)) + Val(MSHFlexGrid11.TextMatrix(r, 6))
        
        If RsTxn_CDTFBOE.RecordCount > 0 Then RsTxn_CDTFBOE.MoveFirst
        RsTxn_CDTFBOE.Find "ClientIDRow=" & TmpRs2!ClientIDRow
        If RsTxn_CDTFBOE.EOF = True Then
           Call AddCDTFBOE(Val(MSHFlexGrid11.TextMatrix(r, 10)), "P")
        End If
        
        Call SaveCDTFBOEGSL(Val(MSHFlexGrid11.TextMatrix(r, 14)), Val(MSHFlexGrid11.TextMatrix(r, 15)), Val(MSHFlexGrid11.TextMatrix(r, 6)), Val(MSHFlexGrid11.TextMatrix(r, 16)), Val(MSHFlexGrid11.TextMatrix(r, 13)), Val(MSHFlexGrid11.TextMatrix(r, 8)))
            
        TmpRs2.MoveNext
    Next
End If
End Sub


Private Sub AddCDTFBOE(mClientID_ As Variant, TxnType_ As Variant)

RsTxn_CDTFBOE.AddNew
RsTxn_CDTFBOE!BOEID = getCDTFBOE_Max
RsTxn_CDTFBOE!LocationID = mLocationID
RsTxn_CDTFBOE!SM_Prefix = TxtSM_Prefix
RsTxn_CDTFBOE!CdtfNo = Val(TxtConvertedID)
RsTxn_CDTFBOE!CDTFDate = TxtPWRDate.Value
RsTxn_CDTFBOE!ClientIDRow = mClientID_
RsTxn_CDTFBOE!InvoiceDate = TxtPWRDate.Value
RsTxn_CDTFBOE!BillingMonth = CDate(Year(TxtPWRDate.Value) & "-" & Month(TxtPWRDate.Value) & "-01")
RsTxn_CDTFBOE!Discount = 0
RsTxn_CDTFBOE!Premium = 0
RsTxn_CDTFBOE!BatchID = 0
RsTxn_CDTFBOE!RBatchID = 0
RsTxn_CDTFBOE!Flag = "P"
RsTxn_CDTFBOE!TxnType = TxnType_
RsTxn_CDTFBOE!G_UID = GetGUID
RsTxn_CDTFBOE!CreatedDate = Now
RsTxn_CDTFBOE!CreatedBy = Gen_UserLoginID
RsTxn_CDTFBOE.Update

If mRejInvIDs = "" Then
   mRejInvIDs = RsTxn_CDTFBOE!BOEID
Else
   mRejInvIDs = mRejInvIDs & "," & RsTxn_CDTFBOE!BOEID
End If

End Sub


'Private Sub SaveCDTFBOEGSL(mGSLID_ As Variant, mNoofBags_ As Variant, mDematWeight_ As Variant, mNoOfDays_ As Variant, mRate_ As Variant)
'
'RsTxn_CDTFBOE_GSL.AddNew
'RsTxn_CDTFBOE_GSL!BOEGSLID = getCDTFBOE_GSL_Max
'RsTxn_CDTFBOE_GSL!BOEID = RsTxn_CDTFBOE!BOEID
'RsTxn_CDTFBOE_GSL!SM_Prefix = TxtSM_Prefix
'RsTxn_CDTFBOE_GSL!gslid = mGSLID_
'RsTxn_CDTFBOE_GSL!CDTFBags = mNoofBags_
'RsTxn_CDTFBOE_GSL!CDTFQty = mDematWeight_
'RsTxn_CDTFBOE_GSL!NoOfDays = mNoOfDays_
'RsTxn_CDTFBOE_GSL!Rate = mRate_ 'TmpRs2!Rate
'RsTxn_CDTFBOE_GSL!BillingUnit = "Q" 'TmpRs2!Unit
'RsTxn_CDTFBOE_GSL!BillingCycleID = 1 ''TmpRs2!BillingCycleID
'RsTxn_CDTFBOE_GSL!Amount = Format(mNoOfDays_ * mDematWeight_ * mRate_, "#####0.00")
'RsTxn_CDTFBOE_GSL.Update
'
'RsTxn_CDTFBOE!Amount = IIf(IsNull(RsTxn_CDTFBOE!Amount), 0, RsTxn_CDTFBOE!Amount) + RsTxn_CDTFBOE_GSL!Amount
'RsTxn_CDTFBOE.Update
'
'tmp = "Update Txn_CxTF_GSL Set BillFlag='I' Where SM_Prefix='" & TxtSM_Prefix & "' And CxTFNo=" & TxtCxTFNo & " And TxnType='W' And GSLID=" & mGSLID_ & " And BillFlag is null"
'Call Tmp5Table
'
'End Sub

Private Sub SaveCDTFBOEGSL(mGSLID_ As Variant, mNoofBags_ As Variant, mDematWeight_ As Variant, mNoOfDays_ As Variant, mRate_ As Variant, amt_ As Variant)

RsTxn_CDTFBOE_GSL.AddNew
RsTxn_CDTFBOE_GSL!BOEGSLID = getCDTFBOE_GSL_Max
RsTxn_CDTFBOE_GSL!BOEID = RsTxn_CDTFBOE!BOEID
RsTxn_CDTFBOE_GSL!SM_Prefix = TxtSM_Prefix
RsTxn_CDTFBOE_GSL!GSLID = mGSLID_
RsTxn_CDTFBOE_GSL!CDTFBags = mNoofBags_
RsTxn_CDTFBOE_GSL!CDTFQty = mDematWeight_
RsTxn_CDTFBOE_GSL!NoOfDays = mNoOfDays_
RsTxn_CDTFBOE_GSL!Rate = mRate_ 'TmpRs2!Rate
RsTxn_CDTFBOE_GSL!BillingUnit = "Q" 'TmpRs2!Unit
RsTxn_CDTFBOE_GSL!BillingCycleID = 1 ''TmpRs2!BillingCycleID
RsTxn_CDTFBOE_GSL!Amount = amt_ 'Format(mNoOfDays_ * mDematWeight_ * mRate_, "#####0.00")
RsTxn_CDTFBOE_GSL.Update

RsTxn_CDTFBOE!Amount = IIf(IsNull(RsTxn_CDTFBOE!Amount), 0, RsTxn_CDTFBOE!Amount) + RsTxn_CDTFBOE_GSL!Amount
RsTxn_CDTFBOE.Update

End Sub
Private Function GetLastDateOfMonth(mDate_ As Date)
Dim mSMonth, mSYear, mSTotalDays As Variant
mSMonth = Month(mDate_)
mSYear = Year(mDate_)
mSTotalDays = Day(DateSerial(mSYear, mSMonth + 1, 0))
GetLastDateOfMonth = CDate(mSYear & "-" & mSMonth & "-" & mSTotalDays)
End Function

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
   
      tmp = "Select TC.PWRDate,GSL.CDTFDate,day(TC.PWRDate) NoOfDay,GSL.ClientIDRow,MC.ClientName,TCG.GSLID,MG.GodownID,GSL.CommodityID,MG.GodownNo,MG.GodownNo+'~'+"
      tmp = tmp & " GSL.StackNo +'~'+ GSL.LotNo As GSL,GSL.SpillageFlag,TCG.AdjstBags,TCG.AdjstQty,(TCG.AdjstQty/TCG.AdjstBags)*1000 as 'BagSize',GSL.NContractID"
      tmp = tmp & " From  Txn_RRNPWRGSL TCG"
      tmp = tmp & " Inner Join Txn_RRNPWR TC on TCG.ConvertedID=TC.ConvertedID "
      tmp = tmp & " left Join Mst_GSL GSL on TCG.NewGSLID=GSL.GSLID And TCG.SM_Prefix=GSL.SM_Prefix"
      tmp = tmp & " Inner Join Mst_Godown MG on GSL.GodownID=MG.GodownID"
      tmp = tmp & " Inner Join Mst_Client MC on GSL.ClientIDRow=MC.ClientIDRow"
      tmp = tmp & " Where GSL.ExchangeTypeID=1 And TC.ConvertedID=" & Val(TxtConvertedID) & " And TC.LOcationID= " & mLocationID & " And GSL.SpillageFlag = 0 And GSL.MadeUpFlag = 0 and isnull(GSL.NContractID,0)=0"
    
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
End If

If CheckfranchiseeRateCard(mCommodityGroupId) = False Then
   Retval = False
End If

CheckRateCards = Retval

End Function
Private Function CheckBillingDetails(mCommodityGroupId_ As Variant) As Boolean

Dim Retval As Boolean
Retval = True

tmp = " Select Distinct TDCD.DepConID,TDC.ClientID,BagSizeFlag  from Txn_DepositContractGodownDetail TDCD"
tmp = tmp & " Inner Join Txn_DepositContract TDC On TDCD.DepConID = TDC.DepConID"
tmp = tmp & " Where TDC.SM_Prefix = '" & TxtPreFix & "' And TDC.ClientID = " & TmpRs!ClientIDRow & " And TDCD.GodownId = " & TmpRs!GodownID & " And TDC.CommodityID = " & TmpRs!CommodityID & " And ApprovalFlag in ('S','A')  And '" & CDate(TxtPWRDate) & "' Between StartDate and EndDate"

Call Tmp2Table

If TmpRs2.RecordCount > 0 Then
   If LCase(TmpRs2!BagSizeFlag) = LCase("N") Then
   Else
      tmp = " Select DepositFinalPrice From Txn_DepositContractGodownDetail Where DepConID  = " & TmpRs2!DepConID & " And " & Round(TmpRs!BagSize, 0) & " <= MaxBagSize And " & Round(TmpRs!BagSize, 0) & " >= MinBagSize"
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
      MsgBox "Rate Card not found for Client " & TmpRs!ClientName & " , Godown - " & TmpRs!GodownNo & " , Commodity - " & TxtCommodityID & " !!"
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
tmp = tmp & " From Txn_RRNPWR TCD"
tmp = tmp & " Inner Join Txn_RRNPWRGSL TCG On TCD.ConvertedID = TCG.ConvertedID "
tmp = tmp & " Inner Join Mst_GSL MG On TCG.NewGSLID = MG.GSLID And TCG.SM_Prefix = MG.SM_Prefix"
tmp = tmp & " Inner Join Mst_Godown MGod On MG.GodownID = MGod.GodownID"
tmp = tmp & " Inner Join Mst_CMP MC On MG.CMPID = MC.CMPID"
tmp = tmp & " Inner Join Mst_Location ML On MC.LocationID = ML.LocationID"
tmp = tmp & " Inner Join Mst_Commodity MComm On TCD.CommodityID = MComm.CommodityID"
tmp = tmp & " Where MGod.UnitTypeID=2  and TCD.ConvertedID=" & Val(TxtConvertedID) & " And TCD.LocationID = " & mLocationID & " And MG.SpillageFlag = 0 And MG.MadeUpFlag= 0"

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
             tmp = tmp & " And '" & Format(TxtPWRDate.Value, Gen_DatFormat) & "'>=c.StartDate And '" & Format(TxtPWRDate.Value, Gen_DatFormat) & "'<=c.EndDate"
             Call Tmp1Table
             If TmpRs1.RecordCount = 0 Then
             
               
                tmp = "Select TC.PWRDate,GSL.CDTFDate,GSL.ClientIDRow,MC.ClientName,TCG.GSLID,MG.GodownNo+'~'+"
                tmp = tmp & " GSL.StackNo +'~'+ GSL.LotNo As GSL,GSL.SpillageFlag,TCG.AdjstBags,TCG.AdjstQty,(TCG.AdjstQty/TCG.AdjstBags)*1000 as 'BagSize',GSL.NContractID,GSL.GodownID"
                tmp = tmp & " From  Txn_RRNPWRGSL TCG"
                tmp = tmp & " Inner Join Txn_RRNPWR TC on TCG.ConvertedID=TC.ConvertedID "
                tmp = tmp & " left Join Mst_GSL GSL on TCG.NewGSLID=GSL.GSLID And TCG.SM_Prefix=GSL.SM_Prefix"
                tmp = tmp & " Inner Join Mst_Godown MG on GSL.GodownID=MG.GodownID"
                tmp = tmp & " Inner Join Mst_Client MC on GSL.ClientIDRow=MC.ClientIDRow"
                tmp = tmp & " Where TC.ConvertedID=" & Val(TxtConvertedID) & " And TC.LocationID= " & mLocationID & " And GSL.SpillageFlag = 0 And GSL.MadeUpFlag = 0 and isnull(GSL.NContractID,0)=0 And GSL.GodownID=" & TmpRs!GodownID
                               
                Call Tmp2Table
                
                For Ln = 1 To TmpRs2.RecordCount
                
                    tmp = "Select  MNR.FF_RateCardID"
                    tmp = tmp & " from  Mst_FF_RateCard_Godown  b"
                    tmp = tmp & " inner Join Mst_FF_RateCard MNR On b.FF_RateCardID = MNR.FF_RateCardID"
                    tmp = tmp & " inner Join Mst_FF_RateCard_Godown_Bag c on c.FF_RateCard_GodownID=b.FF_RateCard_GodownID"
                    tmp = tmp & " Where " & TmpRs2!BagSize & " <= BagSizeTo And " & TmpRs2!BagSize & " >= BagSizeFrom  And MNR.CommodityGroupID=" & mCommodityGroupId_ & " And b.GodownID=" & TmpRs2!GodownID & " And b.BillingUnit in ('B')  And MNR.Flag='A'"

                    
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



Private Sub UpdateBillingDetails()

tmp = "Select ConvertedID,LocationID,CxTFDate,CDTFDate,NoOfDay,ClientIDRow,GSLID,GodownID,CommodityID,SpillageFlag,NoOfBags,DematWeight,BagSize,NContractID,BillingCycleID,DepBillingCycleID,BagSizeFlag ,FinalPrice,DepositFinalPrice,G_UID,DepConDID"
tmp = tmp & " From (Select TC.ConvertedID,TC.LocationID ,TC.PWRDate 'CxTFDate',GSL.CDTFDate,day(TC.PWRDate) NoOfDay,GSL.ClientIDRow,TCG.GSLID,GSL.GodownID,GSL.CommodityID ,GSL.SpillageFlag,TCG.NoOfBags,TCG.DematWeight,(TCG.DematWeight/TCG.NoOfBags)*1000 as 'BagSize',GSL.NContractID,TDC.BillingCycleID,TDC.DepBillingCycleID,TDC.BagSizeFlag ,TDCD.FinalPrice,TDCD.DepositFinalPrice,TCG.G_UID,TDCD.DepConDID"
tmp = tmp & " From  Txn_RRNPWRGSL TCG"
tmp = tmp & " Inner Join Txn_RRNPWR TC on TCG.ConvertedID=TC.ConvertedID "
tmp = tmp & " left Join Mst_GSL GSL on TCG.NewGSLID=GSL.GSLID And TCG.SM_Prefix=GSL.SM_Prefix"
tmp = tmp & " Inner Join Txn_DepositContractGodownDetail TDCD On  GSL.GodownID = TDCD.GodownId And ((TCG.AdjstQty/TCG.AdjstBags)*1000) Between  MinBagSize And MaxBagSize"
tmp = tmp & " Inner Join Txn_DepositContract TDC On TDCD.DepConID = TDC.DepConID And TDC.SM_Prefix = GSL.SM_Prefix And GSL.ClientIDRow = TDC.ClientID And TC.CommodityID = TDC.CommodityID And TDC.ApprovalFlag in ('S','A')  And TC.CxTFDate Between StartDate and EndDate"
tmp = tmp & " Where TC.LocationID = " & mLocationID & " And TC.ConvertedID = " & Val(TxtConvertedID) & " And TC.PWRDate = '" & TxtPWRDate.Value & "'  And GSL.ExchangeTypeID=1 And TCG.DepConDID is null And GSL.SpillageFlag = 0 and isnull(GSL.NContractID,0)=0 and isnull(GSL.ContractID,0)=0 And TDC.BagSizeFlag = 'Y'"
tmp = tmp & " Union All"
tmp = tmp & " Select TC.ConvertedID,TC.LocationID ,TC.PWRDate 'CxTFDate',GSL.CDTFDate,day(TC.PWRDate) NoOfDay,GSL.ClientIDRow,TCG.GSLID,GSL.GodownID,GSL.CommodityID ,GSL.SpillageFlag,TCG.NoOfBags,TCG.DematWeight,(TCG.DematWeight/TCG.NoOfBags)*1000 as 'BagSize',GSL.NContractID,TDC.BillingCycleID,TDC.DepBillingCycleID,TDC.BagSizeFlag,TDCD.FinalPrice,TDCD.DepositFinalPrice,TCG.G_UID,TDCD.DepConDID"
tmp = tmp & " From  Txn_RRNPWRGSL TCG"
tmp = tmp & " Inner Join Txn_RRNPWR TC on TCG.ConvertedID=TC.ConvertedID "
tmp = tmp & " left Join Mst_GSL GSL on TCG.NewGSLID=GSL.GSLID And TCG.SM_Prefix=GSL.SM_Prefix"
tmp = tmp & " Inner Join Txn_DepositContractGodownDetail TDCD On GSL.GodownID = TDCD.GodownId"
tmp = tmp & " Inner Join Txn_DepositContract TDC On TDCD.DepConID = TDC.DepConID And  TDC.SM_Prefix = GSL.SM_Prefix And GSL.ClientIDRow = TDC.ClientID And TC.CommodityID = TDC.CommodityID And TDC.ApprovalFlag in ('S','A')  And TC.PWRDate Between StartDate and EndDate"
tmp = tmp & " Where TC.LocationID = " & mLocationID & " And TC.ConvertedID = " & Val(TxtConvertedID) & " And TC.PWRDate = '" & TxtPWRDate.Value & "'  And GSL.ExchangeTypeID=1 And TCG.DepConDID is null And GSL.SpillageFlag = 0 and isnull(GSL.NContractID,0)=0 and isnull(GSL.ContractID,0)=0 And TDC.BagSizeFlag = 'N') A"
tmp = tmp & " Order By ConvertedID,GSLID"

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
       tmp = "Update Txn_RRNPWRGSL Set RateCardType = 'D',BillingCycleID = " & mDepoBillingCycleID & ",BillingUnit = '" & mDepoBillingUnit & "',Rate = " & mDepoBillingRate & ",RateForDailyPnL = " & mDepoBillingRate & ",DepConDID =  " & mDepConDID & ",NoofDays = " & mNoofDays & " Where ConvertedID = '" & TmpRs!ConvertedID & "' And DepConDID is null And G_UID = '" & TmpRs!G_UID & "' "
       Call Tmp1Table
    End If
    
    TmpRs.MoveNext
Next

End Sub

