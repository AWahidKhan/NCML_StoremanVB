VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmMst_GSL 
   Caption         =   "GSL Master"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7380
   MDIChild        =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   7380
   WindowState     =   2  'Maximized
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   -120
      TabIndex        =   77
      Top             =   10080
      Width           =   10935
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   9015
      Left            =   15000
      TabIndex        =   0
      Top             =   0
      Width           =   255
   End
   Begin VB.Frame Frame4 
      Height          =   10095
      Left            =   0
      TabIndex        =   11
      Top             =   -120
      Width           =   15135
      Begin VB.Frame Frame0 
         Height          =   2565
         Left            =   120
         TabIndex        =   47
         Top             =   4080
         Visible         =   0   'False
         Width           =   14745
         Begin VB.CheckBox Check3 
            Appearance      =   0  'Flat
            Caption         =   "Data with First CITF Date"
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   3240
            TabIndex        =   96
            Top             =   1635
            Width           =   2100
         End
         Begin VB.ComboBox CmbSMadeupFlag 
            Height          =   315
            ItemData        =   "Frm_Mst_GSL.frx":0000
            Left            =   1800
            List            =   "Frm_Mst_GSL.frx":000A
            TabIndex        =   92
            Top             =   1680
            Width           =   1365
         End
         Begin VB.CheckBox Check2 
            Appearance      =   0  'Flat
            Caption         =   "CDTF Date"
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   8280
            TabIndex        =   83
            Top             =   1350
            Width           =   735
         End
         Begin VB.TextBox TxtSGSLID 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   6600
            MaxLength       =   10
            TabIndex        =   78
            Top             =   1320
            Width           =   1125
         End
         Begin VB.CommandButton CmdSearch 
            Caption         =   "Search"
            Height          =   350
            Left            =   12720
            TabIndex        =   66
            Top             =   1320
            Width           =   1000
         End
         Begin VB.ComboBox CmbSCmp 
            Height          =   315
            Left            =   1845
            Sorted          =   -1  'True
            TabIndex        =   50
            Top             =   600
            Width           =   3165
         End
         Begin VB.ComboBox CmbSGodown 
            Height          =   315
            Left            =   5085
            Sorted          =   -1  'True
            TabIndex        =   52
            Top             =   600
            Width           =   1725
         End
         Begin VB.ComboBox CmbSExchangeType 
            Height          =   315
            Left            =   9720
            TabIndex        =   61
            Top             =   960
            Width           =   2085
         End
         Begin VB.ComboBox CmbSCommodity 
            Height          =   315
            Left            =   8565
            Sorted          =   -1  'True
            TabIndex        =   58
            Top             =   600
            Width           =   3285
         End
         Begin VB.CommandButton CmdExcel 
            Caption         =   "Excel"
            Height          =   350
            Left            =   13680
            TabIndex        =   67
            Top             =   1320
            Width           =   1000
         End
         Begin VB.ComboBox CmbPreFix 
            Height          =   315
            Left            =   165
            Sorted          =   -1  'True
            TabIndex        =   49
            Top             =   600
            Width           =   1605
         End
         Begin VB.ComboBox CmbStack 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "Frm_Mst_GSL.frx":002A
            Left            =   6840
            List            =   "Frm_Mst_GSL.frx":002C
            Sorted          =   -1  'True
            TabIndex        =   56
            Top             =   600
            Width           =   1695
         End
         Begin VB.TextBox TxtSClientName 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1200
            TabIndex        =   59
            Top             =   960
            Width           =   4365
         End
         Begin VB.TextBox TxtSISINNo 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   6450
            TabIndex        =   60
            Top             =   960
            Width           =   2085
         End
         Begin VB.ComboBox CmbSFlag 
            Height          =   315
            ItemData        =   "Frm_Mst_GSL.frx":002E
            Left            =   11880
            List            =   "Frm_Mst_GSL.frx":0044
            TabIndex        =   62
            Top             =   600
            Width           =   1365
         End
         Begin VB.TextBox TxtSBalBags 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1800
            MaxLength       =   10
            TabIndex        =   63
            Top             =   1320
            Width           =   1365
         End
         Begin VB.TextBox TxtSBalWt 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4560
            MaxLength       =   10
            TabIndex        =   64
            Top             =   1320
            Width           =   1365
         End
         Begin VB.ComboBox CmbSSplilageFlag 
            Height          =   315
            ItemData        =   "Frm_Mst_GSL.frx":00AE
            Left            =   13320
            List            =   "Frm_Mst_GSL.frx":00B8
            TabIndex        =   65
            Top             =   960
            Width           =   1365
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   1215
            Left            =   120
            TabIndex        =   54
            Top             =   2160
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   2143
            _Version        =   393216
            Rows            =   3
            FixedRows       =   2
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSComCtl2.DTPicker TxtFromDate 
            Height          =   315
            Left            =   9720
            TabIndex        =   84
            Top             =   1320
            Visible         =   0   'False
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   556
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
            Format          =   110690305
            CurrentDate     =   39864
         End
         Begin MSComCtl2.DTPicker TxtToDate 
            Height          =   315
            Left            =   11280
            TabIndex        =   85
            Top             =   1320
            Visible         =   0   'False
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   556
            _Version        =   393216
            Format          =   110690305
            CurrentDate     =   39864
         End
         Begin VB.ComboBox CmbSSCMFlag 
            Height          =   315
            ItemData        =   "Frm_Mst_GSL.frx":00DD
            Left            =   13320
            List            =   "Frm_Mst_GSL.frx":00E7
            TabIndex        =   82
            Top             =   600
            Width           =   1365
         End
         Begin VB.Label Label19 
            Caption         =   "Made Up Flag"
            Height          =   225
            Left            =   240
            TabIndex        =   93
            Top             =   1725
            Width           =   1080
         End
         Begin VB.Label Label12 
            Caption         =   "Flag"
            Height          =   225
            Left            =   12382
            TabIndex        =   91
            Top             =   240
            Width           =   360
         End
         Begin VB.Label Label21 
            Caption         =   "Stack No"
            Height          =   180
            Left            =   7320
            TabIndex        =   90
            Top             =   240
            Width           =   720
         End
         Begin VB.Label Label3 
            Caption         =   "Commodity"
            Height          =   180
            Left            =   9960
            TabIndex        =   89
            Top             =   262
            Width           =   840
         End
         Begin VB.Label Label18 
            Caption         =   "SCM Flag"
            Height          =   195
            Left            =   13635
            TabIndex        =   88
            Top             =   255
            Width           =   720
         End
         Begin VB.Label LblTo 
            Caption         =   "To"
            Height          =   225
            Left            =   11040
            TabIndex        =   87
            Top             =   1365
            Visible         =   0   'False
            Width           =   360
         End
         Begin VB.Label LblFrom 
            Caption         =   "From"
            Height          =   225
            Left            =   9240
            TabIndex        =   86
            Top             =   1365
            Visible         =   0   'False
            Width           =   360
         End
         Begin VB.Label Label16 
            Caption         =   "GSL ID"
            Height          =   180
            Left            =   6000
            TabIndex        =   79
            Top             =   1380
            Width           =   1560
         End
         Begin VB.Label Label2 
            Caption         =   "CMP"
            Height          =   300
            Left            =   3240
            TabIndex        =   76
            Top             =   240
            Width           =   360
         End
         Begin VB.Label Label4 
            Caption         =   "Godown"
            Height          =   300
            Left            =   5647
            TabIndex        =   75
            Top             =   240
            Width           =   600
         End
         Begin VB.Label Label8 
            Caption         =   "Location ID"
            Height          =   300
            Left            =   240
            TabIndex        =   74
            Top             =   240
            Width           =   1320
         End
         Begin VB.Label Label6 
            Caption         =   "Client Name"
            Height          =   225
            Left            =   240
            TabIndex        =   73
            Top             =   1005
            Width           =   1320
         End
         Begin VB.Label Label5 
            Caption         =   "Exchange Type"
            Height          =   225
            Left            =   8565
            TabIndex        =   72
            Top             =   1005
            Width           =   1320
         End
         Begin VB.Label Label9 
            Caption         =   "CMSE Lot"
            Height          =   225
            Left            =   5640
            TabIndex        =   71
            Top             =   1005
            Width           =   1320
         End
         Begin VB.Label Label13 
            Caption         =   "Balance No of Bags "
            Height          =   180
            Left            =   240
            TabIndex        =   70
            Top             =   1380
            Width           =   1560
         End
         Begin VB.Label Label14 
            Caption         =   "Balance Weight"
            Height          =   180
            Left            =   3240
            TabIndex        =   69
            Top             =   1380
            Width           =   1560
         End
         Begin VB.Label Label15 
            Caption         =   "Spillage Flag"
            Height          =   225
            Left            =   11880
            TabIndex        =   68
            Top             =   1005
            Width           =   1080
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Transaction Details"
         Height          =   9735
         Left            =   7320
         TabIndex        =   44
         Top             =   120
         Width           =   7575
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
            Height          =   1215
            Left            =   120
            TabIndex        =   45
            Top             =   255
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   2143
            _Version        =   393216
            FixedRows       =   0
            FixedCols       =   0
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin VB.Frame ButtonFrm 
         Height          =   1200
         Left            =   120
         TabIndex        =   42
         Top             =   8760
         Width           =   7125
         Begin VB.CommandButton AddCmd 
            Caption         =   "&Add New"
            Height          =   400
            Left            =   100
            TabIndex        =   43
            Top             =   240
            Width           =   1400
         End
         Begin VB.CommandButton NextCmd 
            Caption         =   "&Next"
            Height          =   400
            Left            =   100
            TabIndex        =   51
            Top             =   645
            Width           =   1400
         End
         Begin VB.CommandButton SaveCmd 
            Caption         =   "Save"
            Height          =   400
            Left            =   1500
            TabIndex        =   6
            Top             =   240
            Width           =   1400
         End
         Begin VB.CommandButton PreviousCmd 
            Caption         =   "&Previous"
            Height          =   400
            Left            =   1500
            TabIndex        =   53
            Top             =   645
            Width           =   1400
         End
         Begin VB.CommandButton EditCmd 
            Caption         =   "&Edit"
            Height          =   400
            Left            =   2895
            TabIndex        =   46
            Top             =   240
            Width           =   1400
         End
         Begin VB.CommandButton FirstCmd 
            Caption         =   "&First"
            Height          =   400
            Left            =   2895
            TabIndex        =   55
            Top             =   645
            Width           =   1400
         End
         Begin VB.CommandButton DeleteCmd 
            Caption         =   "&Delete"
            Height          =   400
            Left            =   4290
            TabIndex        =   48
            Top             =   240
            Width           =   1400
         End
         Begin VB.CommandButton LastCmd 
            Caption         =   "&Last"
            Height          =   400
            Left            =   4290
            TabIndex        =   57
            Top             =   645
            Width           =   1400
         End
         Begin VB.CommandButton SearchCmd 
            Caption         =   "Search"
            Height          =   400
            Left            =   5685
            TabIndex        =   8
            Top             =   240
            Width           =   1400
         End
         Begin VB.CommandButton ExitCmd 
            Caption         =   "E&xit"
            Height          =   400
            Left            =   5685
            TabIndex        =   7
            Top             =   645
            Width           =   1400
         End
      End
      Begin VB.Frame Frame1 
         Height          =   8655
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   7125
         Begin VB.CommandButton CmdStackID 
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
            Height          =   315
            Left            =   6240
            TabIndex        =   102
            TabStop         =   0   'False
            Top             =   1395
            Width           =   375
         End
         Begin VB.ComboBox CmbStackID 
            Height          =   315
            Left            =   2160
            Sorted          =   -1  'True
            TabIndex        =   3
            Top             =   1395
            Width           =   3885
         End
         Begin VB.TextBox TxtStackNo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   360
            Left            =   2160
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   101
            Top             =   1395
            Width           =   3840
         End
         Begin VB.TextBox TxtCITFMoisturePer 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   360
            Left            =   4680
            TabIndex        =   99
            TabStop         =   0   'False
            Top             =   4170
            Width           =   1320
         End
         Begin VB.TextBox TxtMoisturePer 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   360
            Left            =   2160
            TabIndex        =   98
            TabStop         =   0   'False
            Top             =   4170
            Width           =   1185
         End
         Begin VB.Frame Frame5 
            Caption         =   "Made Up Details"
            Height          =   1290
            Left            =   120
            TabIndex        =   94
            Top             =   7320
            Visible         =   0   'False
            Width           =   6900
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid4 
               Height          =   975
               Left            =   60
               TabIndex        =   95
               Top             =   225
               Width           =   6735
               _ExtentX        =   11880
               _ExtentY        =   1720
               _Version        =   393216
               BackColor       =   16777215
               WordWrap        =   -1  'True
               AllowUserResizing=   1
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
         End
         Begin VB.CheckBox ChkSCMFlag 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3930
            TabIndex        =   80
            Top             =   4635
            Width           =   240
         End
         Begin VB.TextBox TxtGSLID 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Enabled         =   0   'False
            Height          =   360
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   240
            Width           =   1200
         End
         Begin VB.TextBox TxtExchangeTypeID 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   360
            Left            =   2160
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   2565
            Width           =   3840
         End
         Begin VB.TextBox TxtClientID 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   360
            Left            =   2160
            MaxLength       =   50
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   2955
            Width           =   3840
         End
         Begin VB.TextBox TxtFlag 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Enabled         =   0   'False
            Height          =   285
            Left            =   5520
            MaxLength       =   1
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   4590
            Width           =   480
         End
         Begin VB.TextBox TxtISINID 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   360
            Left            =   2160
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   3765
            Width           =   3840
         End
         Begin VB.TextBox TxtPledgeNo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   360
            Left            =   2160
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   3360
            Width           =   3840
         End
         Begin VB.TextBox TxtLotNo 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   2160
            MaxLength       =   10
            TabIndex        =   4
            Top             =   1785
            Width           =   3840
         End
         Begin VB.ComboBox CmbGodownID 
            Height          =   315
            Left            =   2160
            Sorted          =   -1  'True
            TabIndex        =   2
            Top             =   1005
            Width           =   3885
         End
         Begin VB.ComboBox CmbCMPName 
            Height          =   315
            Left            =   2160
            Sorted          =   -1  'True
            TabIndex        =   1
            Top             =   615
            Width           =   3885
         End
         Begin VB.CommandButton CmdGodownID 
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
            Height          =   315
            Left            =   6240
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   1005
            Width           =   375
         End
         Begin VB.CommandButton CmdCMPName 
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
            Height          =   315
            Left            =   6240
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   615
            Width           =   375
         End
         Begin VB.Frame Frame3 
            Caption         =   "Store Details"
            Height          =   1560
            Left            =   120
            TabIndex        =   16
            Top             =   5760
            Width           =   6900
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
               Height          =   1290
               Left            =   60
               TabIndex        =   17
               Top             =   225
               Width           =   6735
               _ExtentX        =   11880
               _ExtentY        =   2275
               _Version        =   393216
               BackColor       =   16777215
               WordWrap        =   -1  'True
               AllowUserResizing=   1
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
         End
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2160
            TabIndex        =   5
            Top             =   4650
            Width           =   255
         End
         Begin VB.TextBox TxtPreFix 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Enabled         =   0   'False
            Height          =   360
            Left            =   4560
            Locked          =   -1  'True
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   240
            Width           =   1440
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
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   4950
            Width           =   3855
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
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   5370
            Width           =   3855
         End
         Begin VB.TextBox TxtCMPName 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   2160
            MaxLength       =   100
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   615
            Width           =   3840
         End
         Begin VB.TextBox TxtGodownID 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   2160
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   1005
            Width           =   3840
         End
         Begin VB.TextBox TxtCommodityID 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   360
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   2175
            Width           =   3840
         End
         Begin VB.Label Label22 
            Caption         =   "CITF Moisture %"
            Height          =   300
            Left            =   3420
            TabIndex        =   100
            Top             =   4200
            Width           =   1440
         End
         Begin VB.Label Label20 
            Caption         =   "CDTF Moisture %"
            Height          =   300
            Left            =   240
            TabIndex        =   97
            Top             =   4200
            Width           =   1440
         End
         Begin VB.Label Label17 
            Caption         =   "SCM Lot"
            Height          =   240
            Left            =   3000
            TabIndex        =   81
            Top             =   4650
            Width           =   720
         End
         Begin VB.Label LblISINID 
            Caption         =   "CMSE Lot No"
            Height          =   240
            Left            =   225
            TabIndex        =   41
            Top             =   3825
            Width           =   1800
         End
         Begin VB.Label LblFlag 
            Caption         =   "Flag"
            Enabled         =   0   'False
            Height          =   225
            Left            =   4800
            TabIndex        =   40
            Top             =   4620
            Width           =   1080
         End
         Begin VB.Label LblPledgeNo 
            Caption         =   "Pledge No"
            Height          =   300
            Left            =   240
            TabIndex        =   39
            Top             =   3390
            Width           =   1800
         End
         Begin VB.Label LblClientID 
            Caption         =   "Client Name"
            Height          =   300
            Left            =   240
            TabIndex        =   38
            Top             =   2955
            Width           =   1800
         End
         Begin VB.Label LblExchangeTypeID 
            Caption         =   "Exchange Type"
            Height          =   300
            Left            =   255
            TabIndex        =   37
            Top             =   2565
            Width           =   1800
         End
         Begin VB.Label LblLotNo 
            Caption         =   "Lot No"
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
            TabIndex        =   36
            Top             =   1815
            Width           =   1800
         End
         Begin VB.Label LblStackNo 
            Caption         =   "Stack No"
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
            Left            =   255
            TabIndex        =   35
            Top             =   1470
            Width           =   1800
         End
         Begin VB.Label LblGodownID 
            Caption         =   "Godown"
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
            TabIndex        =   34
            Top             =   1005
            Width           =   1800
         End
         Begin VB.Label LblCommodityID 
            Caption         =   "Commodity"
            Height          =   300
            Left            =   240
            TabIndex        =   33
            Top             =   2175
            Width           =   1800
         End
         Begin VB.Label LblCMPName 
            Caption         =   "CMP"
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
            TabIndex        =   32
            Top             =   615
            Width           =   1800
         End
         Begin VB.Label LblGSLID 
            Caption         =   "GSL ID"
            Height          =   300
            Left            =   240
            TabIndex        =   31
            Top             =   240
            Width           =   1800
         End
         Begin VB.Label Label1 
            Caption         =   "Spillage"
            Height          =   240
            Left            =   240
            TabIndex        =   30
            Top             =   4650
            Width           =   1080
         End
         Begin VB.Label Label7 
            Caption         =   "Location ID"
            Height          =   180
            Left            =   3480
            TabIndex        =   29
            Top             =   330
            Width           =   1080
         End
         Begin VB.Label Label10 
            Caption         =   "Created  :"
            Height          =   255
            Left            =   240
            TabIndex        =   28
            Top             =   5010
            Width           =   1695
         End
         Begin VB.Label Label11 
            Caption         =   "Updated :"
            Height          =   255
            Left            =   240
            TabIndex        =   27
            Top             =   5415
            Width           =   1695
         End
      End
   End
End
Attribute VB_Name = "FrmMst_GSL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg As Variant
Dim mCMPName, mCommodityID, mGodownID, mExchangeTypeID, mClientID, mISINID, mClientIDRow As Double
Dim msSM_prefix, mSLocationID, mSFlag, mSSplilageFlag, mSCMFlag, mMadeupFlag, mSMadeUpFlag As Variant
Dim mStackNo As Variant

Private Sub Check2_Click()
If Check2.Value = 1 Then
    LblFrom.Visible = True
    LblTo.Visible = True
    TxtFromDate.Visible = True
    TxtToDate.Visible = True
    TxtFromDate.Value = Date - 365
Else
    LblFrom.Visible = False
    LblTo.Visible = False
    TxtFromDate.Visible = False
    TxtToDate.Visible = False
    TxtToDate.Value = Date
End If
End Sub

Private Sub CmbPreFix_GotFocus()
tmp = CmbPreFix.Text

'Call TableMst_StoreManLocations("where " & Gen_WcSMLocation)

If Gen_WcLocation = "" Then
   Call TableMst_Location("")
Else
   Call TableMst_Location("where " & Gen_WcLocation)
End If


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
   MsgBox "Invalid PreFix Selection !!!!"
   CmbPreFix.SetFocus
   Exit Sub
End If
msSM_prefix = RsMst_Location!SM_Prefix
mSLocationID = RsMst_Location!LocationID
End Sub

Private Sub CmbSClient_Change()

End Sub

Private Sub CmbSCmp_GotFocus()
If CmbPreFix.Text = "" Then
   MsgBox "Plase Select Location !!!"
   CmbPreFix.SetFocus
   Exit Sub
End If

tmp1 = ""
If Gen_WcCMP <> "" Then
   tmp1 = "And " & Gen_WcCMP
End If

tmp = CmbSCmp.Text
Call TableMst_CMP(" Where LocationID = " & mSLocationID & "" & tmp1)
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

Private Sub CmbSFlag_LostFocus()
mSFlag = ""
If CmbSFlag.Text = "" Then Exit Sub
If CmbSFlag.Text = "B ~ Blank" Or CmbSFlag.Text = "O ~ Original" _
   Or CmbSFlag.Text = "Z ~ Transfer From" Or CmbSFlag.Text = "J ~ Rejected" _
   Or CmbSFlag.Text = "E ~ Expired (Withdrawan)" Or CmbSFlag.Text = "R ~ Revalidation" Then
   mSFlag = mID(CmbSFlag.Text, 1, 1)
Else
   MsgBox "Invalid Selection !!!"
   CmbSFlag.SetFocus
End If
End Sub

Private Sub CmbSGodown_GotFocus()
If CmbSCmp.Text = "" Then
   MsgBox "Select CMP first !!! "
   CmbSCmp.SetFocus
   Exit Sub
End If
tmp = CmbSGodown.Text
Call TableMst_Godown("Where CMPID = " & mCMPName)
CmbSGodown.Clear
If RsMst_Godown.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Godown.RecordCount
    CmbSGodown.AddItem RsMst_Godown!GodownNo
    RsMst_Godown.MoveNext
Next
CmbGodownID.Text = tmp
End Sub

Private Sub CmbSGodown_LostFocus()
If CmbSGodown.Text = "" Then
   Exit Sub
End If
RsMst_Godown.MoveFirst
RsMst_Godown.Find "GodownNo = '" & CmbSGodown.Text & "'"
If RsMst_Godown.EOF Then
   MsgBox "Invalid selection "
   CmbSGodown.SetFocus
   Exit Sub
End If
mGodownID = RsMst_Godown!GodownID
End Sub


Private Sub CmbSMadeupFlag_LostFocus()

mSMadeUpFlag = ""
If CmbSMadeupFlag.Text = "" Then Exit Sub

If CmbSMadeupFlag.Text = "0 ~ Non-Madeup" Or CmbSMadeupFlag.Text = "1 ~ Madeup" Then
   mSMadeUpFlag = mID(CmbSMadeupFlag.Text, 1, 1)
Else
   MsgBox "Invalid Selection !!!"
   CmbSMadeupFlag.SetFocus
End If

End Sub

Private Sub CmbSSCMFlag_LostFocus()
mSCMFlag = ""
If CmbSSCMFlag.Text = "" Then Exit Sub

If CmbSSCMFlag.Text = "0 ~ Not SCM" Or CmbSSCMFlag.Text = "1 ~ SCM" Then
   mSCMFlag = mID(CmbSSCMFlag.Text, 1, 1)
Else
   MsgBox "Invalid Selection !!!"
   CmbSSCMFlag.SetFocus
End If
End Sub

Private Sub CmbSSplilageFlag_LostFocus()
mSSplilageFlag = ""
If CmbSSplilageFlag.Text = "" Then Exit Sub

If CmbSSplilageFlag.Text = "0 ~ Non-Spillage" Or CmbSSplilageFlag.Text = "1 ~ Splillage" Then
   mSSplilageFlag = mID(CmbSSplilageFlag.Text, 1, 1)
Else
   MsgBox "Invalid Selection !!!"
   CmbSSplilageFlag.SetFocus
End If
End Sub

Private Sub CmbStack_GotFocus()

If CmbSGodown.Text = "" Or mGodownID = "" Then
  MsgBox "Please Select Godown No!!!"
  Exit Sub
End If
tmp = CmbStack.Text
Call TableMst_GSL(" Where CMPID = " & mCMPName & " And GodownID = " & mGodownID, "D")
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

mCmbStack = CmbStack.Text
End Sub

Private Sub CmbStackID_GotFocus()
If Edit_Flg = "A" Then
If CmbGodownID.Text = "" Then
   MsgBox "Select Godown first !!! "
   CmbGodownID.SetFocus
   Exit Sub
End If
    tmp = CmbStackID.Text
Else
    tmp = TxtStackNo
End If

Call TableMst_Stack("Where Flag = 'A' AND GodownID = " & mGodownID)
CmbStackID.Clear
If RsMst_Stack.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Stack.RecordCount
    CmbStackID.AddItem RsMst_Stack!StackNo
    RsMst_Stack.MoveNext
Next
CmbStackID.Text = tmp
End Sub
Private Sub CmbStackID_LostFocus()
If CmbStackID.Text = "" Then
   TxtLotNo = ""
   Exit Sub
End If

'If tmp <> CmbStackID.Text Then
'   TxtLotNo = ""
'End If

RsMst_Stack.MoveFirst
RsMst_Stack.Find "StackNo = '" & CmbStackID.Text & "'"
If RsMst_Stack.EOF Then
   MsgBox "Invalid selection "
   CmbStackID.SetFocus
   Exit Sub
End If
mStackNo = RsMst_Stack!StackNo
TxtLotNo.Text = GetNextLotNo(mGodownID, mStackNo)
End Sub

Private Sub CmdSearch_Click()
Dim wc As String
Dim x As Double
Call Grid_Head
wc = ""
If CmbSCmp.Text <> "" Then
   wc = GenWc(wc, mCMPName, "Mst_CMP.CMPID", "N", "N")
Else
   If Gen_WcCMP <> "" Then
      tmp = Replace(Gen_WcCMP, "CMPID", "mst_CMP.CMPID")
      wc = GenWc(wc, tmp, "", "N", "I")
   End If
End If
If CmbSGodown.Text <> "" Then
   wc = GenWc(wc, mGodownID, "Mst_Godown.GodownID", "N", "N")
End If

If CmbSCommodity.Text <> "" Then
   wc = GenWc(wc, mCommodityID, "Mst_Commodity.CommodityID", "N", "N")
End If

If CmbSExchangeType.Text <> "" Then
   wc = GenWc(wc, mExchangeTypeID, "Mst_ExchangeType.ExchangeTypeID", "N", "N")
End If
If CmbPreFix.Text <> "" Then
   wc = GenWc(wc, msSM_prefix, "Mst_GSL.sm_prefix", "S", "N")
End If

If CmbStack.Text <> "" Then
   wc = GenWc(wc, CmbStack.Text, "Mst_GSL.StackNo", "S", "N")
End If

If TxtSClientName.Text <> "" Then
   wc = GenWc(wc, TxtSClientName, "Mst_Client.ClientName", "S", "L")
End If
If TxtSISINNo.Text <> "" Then
   wc = GenWc(wc, TxtSISINNo, "Mst_GSL.ISINID", "S", "L")
End If

If CmbSFlag.Text <> "" Then
   wc = GenWc(wc, mSFlag, "Mst_GSL.Flag", "S", "N")
End If

If CmbSSplilageFlag.Text <> "" Then
   wc = GenWc(wc, mSSplilageFlag, "Mst_GSL.SpillageFlag", "S", "N")
End If


If TxtSBalBags <> "" Then
    If Not IsOprater(TxtSBalBags) Then
       wc = GenWc(wc, TxtSBalBags, "Mst_GSL.BalanceBags", "N", "N")
    Else
       wc = GenWc(wc, TxtSBalBags, "Mst_GSL.BalanceBags", "N", "I")
    End If
End If

If TxtSBalWt <> "" Then
    If Not IsOprater(TxtSBalWt) Then
       wc = GenWc(wc, TxtSBalWt, "Mst_GSL.BalanceWeight", "N", "N")
    Else
       wc = GenWc(wc, TxtSBalWt, "Mst_GSL.BalanceWeight", "N", "I")
    End If
End If

If Check2.Value = 1 Then
    If TxtToDate < TxtFromDate Then
        MsgBox "'From Date' Must be less than 'To Date' !!!"
        TxtFromDate.SetFocus
        Exit Sub
    End If
    If Gen_DBType = "MDB" Then
       If wc <> "" Then
          wc = wc & " And (Mst_GSL.CDTFDate Between #" & Format(TxtFromDate, Gen_DatFormat) & "# And #" & Format(TxtToDate, Gen_DatFormat) & "#) "
       Else
          wc = " Where (Mst_GSL.CDTFDate Between #" & Format(TxtFromDate, Gen_DatFormat) & "# And #" & Format(TxtToDate, Gen_DatFormat) & "#) "
       End If
    Else
       If wc <> "" Then
          wc = wc & " And (Mst_GSL.CDTFDate Between '" & Format(TxtFromDate, Gen_DatFormat) & "' And '" & Format(TxtToDate, Gen_DatFormat) & "') "
       Else
          wc = " Where (Mst_GSL.CDTFDate Between '" & Format(TxtFromDate, Gen_DatFormat) & "' And '" & Format(TxtToDate, Gen_DatFormat) & "') "
       End If
    End If
End If

If CmbSSCMFlag.Text <> "" Then
   wc = GenWc(wc, mSCMFlag, "Mst_GSL.SCMFlag", "S", "N")
   'Wc = Wc & " And Mst_GSL.SCMFlag = " & mSCMFlag
End If

If CmbSMadeupFlag.Text <> "" Then
   wc = GenWc(wc, mSMadeUpFlag, "Mst_GSL.MAdeupFlag", "S", "N")
   'Wc = Wc & " And Mst_GSL.SCMFlag = " & mSCMFlag
End If


If TxtSGSLID <> "" Then
   wc = GenWc(wc, TxtSGSLID, "Mst_GSL.Mst_GSL_ID", "N", "N")
End If


'
'tmp = "SELECT Mst_GSL.GSLID, Mst_GSL.sm_prefix , Mst_CMP.CMPName, Mst_Commodity.Commodity, Mst_Godown.GodownNo, Mst_GSL.StackNo,"
'tmp = tmp & " Mst_GSL.LotNo , Mst_GSL.DepositeBags, Mst_GSL.CDTFBags, Mst_GSL.BalanceBags, Mst_GSL.DepositeWeight,"
'tmp = tmp & " Mst_GSL.CDTFWeight , Mst_GSL.BalanceWeight, Mst_ExchangeType.ExchangeType, Mst_Client.ClientName,"
'tmp = tmp & " Mst_GSL.Flag , Mst_GSL.ISINID, Mst_GSL.CDTFNo, Mst_GSL.CDTFDate, Mst_GSL.PledgeNo, Mst_GSL.PledgeBags,"
'tmp = tmp & " Mst_GSL.PledgeWeight , Mst_GSL.SpillageFlag , Mst_GSL.DepositAdjWt,Mst_GSL.WithdrawalAdjWt,mst_gsl.g_uid, "
'tmp = tmp & " Mst_GSL.KRNo, Mst_GSL.GrossWt, "
'tmp = tmp & " Mst_GSL.BGSLID , Mst_GSL.FGSLID, mst_gsl.WHRBags,mst_gsl.WHRWeight"
'tmp = tmp & " FROM (Mst_Client RIGHT JOIN ((Mst_Godown RIGHT JOIN (Mst_Commodity RIGHT JOIN (Mst_CMP RIGHT JOIN Mst_GSL ON"
'tmp = tmp & " Mst_CMP.CMPID = Mst_GSL.CMPID) ON Mst_Commodity.CommodityID = Mst_GSL.CommodityID) ON Mst_Godown.GodownID ="
'tmp = tmp & " Mst_GSL.GodownID) LEFT JOIN Mst_ExchangeType ON Mst_GSL.ExchangeTypeID = Mst_ExchangeType.ExchangeTypeID)"
'tmp = tmp & " ON Mst_Client.ClientIDRow = Mst_GSL.ClientIDRow) LEFT JOIN Mst_ISIN ON Mst_GSL.ISINID = Mst_ISIN.ISINID"


'tmp = "SELECT Mst_GSL.GSLID, Mst_GSL.sm_prefix , Mst_CMP.CMPName, Mst_Commodity.Commodity, Mst_Godown.GodownNo,"
'tmp = tmp & " Mst_GSL.StackNo, Mst_GSL.LotNo , Mst_GSL.DepositeBags, Mst_GSL.CDTFBags, Mst_GSL.BalanceBags, Mst_GSL.DepositeWeight,"
'tmp = tmp & " Mst_GSL.CDTFWeight , Mst_GSL.BalanceWeight, Mst_ExchangeType.ExchangeType, Mst_Client.ClientName, Mst_GSL.Flag ,"
'tmp = tmp & " Mst_GSL.ISINID, Mst_GSL.CDTFNo, Mst_GSL.CDTFDate, Mst_GSL.PledgeNo, Mst_GSL.PledgeBags, Mst_GSL.PledgeWeight ,"
'tmp = tmp & " Mst_GSL.SpillageFlag , Mst_GSL.DepositAdjWt,Mst_GSL.WithdrawalAdjWt,mst_gsl.g_uid,  Mst_GSL.KRNo, Mst_GSL.GrossWt,"
'tmp = tmp & " Mst_GSL.BGSLID , Mst_GSL.FGSLID, mst_gsl.WHRBags,mst_gsl.WHRWeight FROM"
'tmp = tmp & " Mst_GSL"
'tmp = tmp & " left JOIN Mst_CMP on Mst_CMP.CMPID = Mst_GSL.CMPID"
'tmp = tmp & " left JOIN Mst_Godown on Mst_Godown.GodownID = Mst_GSL.GodownID"
'tmp = tmp & " left JOIN Mst_Commodity on Mst_Commodity.CommodityID = Mst_GSL.CommodityID"
'tmp = tmp & " LEFT JOIN Mst_ExchangeType ON Mst_GSL.ExchangeTypeID = Mst_ExchangeType.ExchangeTypeID"
'tmp = tmp & " left JOIN Mst_Client on Mst_Client.ClientIDRow = Mst_GSL.ClientIDRow"
'tmp = tmp & " LEFT JOIN Mst_ISIN ON Mst_GSL.ISINID = Mst_ISIN.ISINID"



tmp = "SELECT Mst_GSL.Mst_GSL_ID, Mst_Location.LocationName , Mst_CMP.CMPName, Mst_Commodity.Commodity, Mst_Godown.GodownNo,"
tmp = tmp & " Mst_GSL.StackNo, Mst_GSL.LotNo , Mst_GSL.DepositeBags, Mst_GSL.CDTFBags, Mst_GSL.BalanceBags, Mst_GSL.DepositeWeight,"
tmp = tmp & " Mst_GSL.CDTFWeight , Mst_GSL.BalanceWeight, Mst_ExchangeType.ExchangeType, Replace(Mst_Client.ClientName,Char(10),'') 'ClientName', Replace(MA.AgentName,Char(10),'') 'AgentName' ,Mst_GSL.Flag ,"
tmp = tmp & " Mst_GSL.ISINID, Mst_GSL.CDTFNo, Mst_GSL.CDTFDate, Mst_GSL.PledgeNo, Mst_GSL.PledgeBags, Mst_GSL.PledgeWeight ,"
tmp = tmp & " Mst_GSL.SpillageFlag , Mst_GSL.DepositAdjWt,Mst_GSL.WithdrawalAdjWt,mst_gsl.g_uid,  Mst_GSL.KRNo, Mst_GSL.GrossWt,"
tmp = tmp & " Mst_GSL.BGSLID_Mst_GSL_ID , Mst_GSL.FGSLID_Mst_GSL_ID, mst_gsl.WHRBags,mst_gsl.WHRWeight,Mst_StorageContract.ContractName, mst_gsl.NContractID , "
tmp = tmp & " Mst_GSL.BagSize,Mst_GSL.BagEquivalent,Mst_GSL.WtQty,Mst_GSL.CRAmount,Mst_GSL.CRDate,Mst_GSL.SCMFLag,Mst_GSL.MadeUpFlag,Mst_GSL.Connected_Mst_GSL_ID,Mst_GSL.BSM_Prefix , Mst_GSL.FSM_Prefix,Mst_GSL.MoisturePer ,Mst_GSL.CITF_MoisturePer,Mst_GSL.WHRNo,Mst_GSL.CMPID FROM "
tmp = tmp & " Mst_GSL"
tmp = tmp & " left JOIN Mst_CMP on Mst_CMP.CMPID = Mst_GSL.CMPID"
tmp = tmp & " left JOIN Mst_Location on Mst_CMP.LocationID = Mst_Location.LocationID "
tmp = tmp & " left JOIN Mst_Godown on Mst_Godown.GodownID = Mst_GSL.GodownID"
tmp = tmp & " left JOIN Mst_Commodity on Mst_Commodity.CommodityID = Mst_GSL.CommodityID"
tmp = tmp & " LEFT JOIN Mst_ExchangeType ON Mst_GSL.ExchangeTypeID = Mst_ExchangeType.ExchangeTypeID"
tmp = tmp & " left JOIN Mst_Client on Mst_Client.ClientID = Mst_GSL.ClientID"
tmp = tmp & " LEFT JOIN Mst_ISIN ON Mst_GSL.ISINID = Mst_ISIN.ISINID"
tmp = tmp & " LEFT JOIN Mst_StorageContract ON Mst_GSL.ContractID = Mst_StorageContract.ContractID"
'tmp = tmp & " Left Join Txn_CxTF_GSL TCG On Mst_GSL.GSLID = TCG.GSLID And Mst_GSL.SM_Prefix = TCG.SM_Prefix And TCG.TxnType = 'D' "
tmp = tmp & " Left Join Txn_CxTF_Details TCD On Mst_GSL.CDTFNo = TCD.CxTFNo And Mst_GSL.SM_Prefix = TCD.SM_Prefix And TCD.TxnType = 'D' "
tmp = tmp & " Left Join Mst_Agent MA On TCD.AgentID = MA.AgentID "


tmp = tmp & " " & wc
tmp = tmp & " ORDER BY Mst_Location.LocationName,Mst_CMP.CMPName,Mst_GSL.Mst_GSL_ID;"

Call TmpTable

If Check3.Value = 1 Then
   tmp = "Select MG.Mst_GSL_ID,MG.SM_PRefix,MG.G_UID,Min(TCD.CxTFDate) as  'FrstCITFDate'"
   tmp = tmp & " From Txn_CxTF_GSL TCG"
   tmp = tmp & " Inner Join Txn_CxTF_Details TCD On TCG.CxTFNo = TCD.CxTFNo And TCG.SM_PRefix = TCD.SM_Prefix And TCG.TxnType = TCD.TxnType"
   tmp = tmp & " Inner Join Mst_GSL MG On TCG.Mst_GSL_ID  = MG.Mst_GSL_ID And TCG.SM_Prefix = MG.SM_Prefix"
   tmp = tmp & " Where TCG.TxnType = 'W' "

   If CmbPreFix.Text <> "" Then
      tmp = tmp & " And TCG.SM_PRefix = '" & msSM_prefix & "'"
   End If

   If TxtSGSLID <> "" Then
      tmp = tmp & " And TCG.Mst_GSL_ID = " & Val(TxtSGSLID) & ""
   End If

   tmp = tmp & " Group By MG.Mst_GSL_ID,MG.SM_PRefix,MG.G_UID"

   Call Tmp1Table

End If

MSHFlexGrid1.Rows = TmpRs.RecordCount + 2
'x = 1
For I = 2 To TmpRs.RecordCount + 1
    'MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1 'TmpRs.RecordCount + 2
    For C = 0 To TmpRs.Fields.Count - 3
        MSHFlexGrid1.TextMatrix(I, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
    Next
     
    MSHFlexGrid1.TextMatrix(I, 2) = Replace(Replace(MSHFlexGrid1.TextMatrix(I, 2), vbCrLf, ""), ",", "~")
    MSHFlexGrid1.TextMatrix(I, 5) = Replace(Replace(MSHFlexGrid1.TextMatrix(I, 5), vbCrLf, ""), ",", "~")
    MSHFlexGrid1.TextMatrix(I, 6) = Replace(Replace(MSHFlexGrid1.TextMatrix(I, 6), vbCrLf, ""), ",", "~")
    MSHFlexGrid1.TextMatrix(I, 17) = Replace(Replace(MSHFlexGrid1.TextMatrix(I, 17), vbCrLf, ""), ",", "~")
     
    MSHFlexGrid1.TextMatrix(I, 14) = Replace(Replace(MSHFlexGrid1.TextMatrix(I, 14), vbCrLf, ""), ",", "~")
    MSHFlexGrid1.TextMatrix(I, 15) = Replace(Replace(MSHFlexGrid1.TextMatrix(I, 15), vbCrLf, ""), ",", "~")
    MSHFlexGrid1.TextMatrix(I, 33) = Replace(Replace(MSHFlexGrid1.TextMatrix(I, 33), vbCrLf, ""), ",", "~")
    MSHFlexGrid1.TextMatrix(I, 34) = Replace(Replace(MSHFlexGrid1.TextMatrix(I, 34), vbCrLf, ""), ",", "~")
    
    MSHFlexGrid1.TextMatrix(I, 14) = Replace(Replace(MSHFlexGrid1.TextMatrix(I, 14), vbNewLine, ""), ",", "~")
    MSHFlexGrid1.TextMatrix(I, 15) = Replace(Replace(MSHFlexGrid1.TextMatrix(I, 15), vbNewLine, ""), ",", "~")
    MSHFlexGrid1.TextMatrix(I, 33) = Replace(Replace(MSHFlexGrid1.TextMatrix(I, 33), vbNewLine, ""), ",", "~")
    MSHFlexGrid1.TextMatrix(I, 34) = Replace(Replace(MSHFlexGrid1.TextMatrix(I, 34), vbNewLine, ""), ",", "~")
     
    MSHFlexGrid1.TextMatrix(I, 7) = Format(MSHFlexGrid1.TextMatrix(I, 7), "########0")
    MSHFlexGrid1.TextMatrix(I, 8) = Format(MSHFlexGrid1.TextMatrix(I, 8), "########0")
    MSHFlexGrid1.TextMatrix(I, 9) = Format(MSHFlexGrid1.TextMatrix(I, 9), "########0")
    MSHFlexGrid1.TextMatrix(I, 10) = Format(MSHFlexGrid1.TextMatrix(I, 10), "########0.000")
    MSHFlexGrid1.TextMatrix(I, 11) = Format(MSHFlexGrid1.TextMatrix(I, 11), "########0.000")
    MSHFlexGrid1.TextMatrix(I, 12) = Format(MSHFlexGrid1.TextMatrix(I, 12), "########0.000")
    MSHFlexGrid1.TextMatrix(I, 21) = Format(MSHFlexGrid1.TextMatrix(I, 21), "########0")
    MSHFlexGrid1.TextMatrix(I, 22) = Format(MSHFlexGrid1.TextMatrix(I, 22), "########0.000")
    
    
    MSHFlexGrid1.TextMatrix(0, 7) = Val(MSHFlexGrid1.TextMatrix(0, 7)) + Val(MSHFlexGrid1.TextMatrix(I, 7))
    MSHFlexGrid1.TextMatrix(0, 8) = Val(MSHFlexGrid1.TextMatrix(0, 8)) + Val(MSHFlexGrid1.TextMatrix(I, 8))
    MSHFlexGrid1.TextMatrix(0, 9) = Val(MSHFlexGrid1.TextMatrix(0, 9)) + Val(MSHFlexGrid1.TextMatrix(I, 9))
    MSHFlexGrid1.TextMatrix(0, 10) = Val(MSHFlexGrid1.TextMatrix(0, 10)) + Val(MSHFlexGrid1.TextMatrix(I, 10))
    MSHFlexGrid1.TextMatrix(0, 11) = Val(MSHFlexGrid1.TextMatrix(0, 11)) + Val(MSHFlexGrid1.TextMatrix(I, 11))
    MSHFlexGrid1.TextMatrix(0, 12) = Val(MSHFlexGrid1.TextMatrix(0, 12)) + Val(MSHFlexGrid1.TextMatrix(I, 12))
    MSHFlexGrid1.TextMatrix(0, 21) = Val(MSHFlexGrid1.TextMatrix(0, 21)) + Val(MSHFlexGrid1.TextMatrix(I, 21))
    MSHFlexGrid1.TextMatrix(0, 22) = Val(MSHFlexGrid1.TextMatrix(0, 22)) + Val(MSHFlexGrid1.TextMatrix(I, 22))
      
    If Check3.Value = 1 Then
       If TmpRs1.RecordCount > 0 Then
          TmpRs1.MoveFirst
          TmpRs1.Find "G_UID = '" & MSHFlexGrid1.TextMatrix(I, 26) & "'"
          If TmpRs1.EOF = False Then
             MSHFlexGrid1.TextMatrix(I, 43) = IIf(IsNull(TmpRs1!FrstCITFDate), "", Format(TmpRs1!FrstCITFDate, "dd-mmm-yyyy"))
          End If
       End If
    End If
    MSHFlexGrid1.TextMatrix(I, 44) = IIf(IsNull(TmpRs!BSM_Prefix), "", TmpRs!BSM_Prefix)
    MSHFlexGrid1.TextMatrix(I, 45) = IIf(IsNull(TmpRs!FSM_Prefix), "", TmpRs!FSM_Prefix)
    MSHFlexGrid1.TextMatrix(I, 19) = Format(MSHFlexGrid1.TextMatrix(I, 19), "dd-mmm-yyyy")
    'MSHFlexGrid1.TextMatrix(I, 40) = IIf(MSHFlexGrid1.TextMatrix(I, 40) = "", 0, MSHFlexGrid1.TextMatrix(I, 40))
    If MSHFlexGrid1.TextMatrix(I, 40) = "1" Then
       MSHFlexGrid1.TextMatrix(I, 40) = "Yes"
    Else
       MSHFlexGrid1.TextMatrix(I, 40) = "No"
    End If
    
    If MSHFlexGrid1.TextMatrix(I, 41) = "1" Then
       MSHFlexGrid1.TextMatrix(I, 41) = "Yes"
    Else
       MSHFlexGrid1.TextMatrix(I, 41) = "No"
    End If
    MSHFlexGrid1.TextMatrix(I, 46) = IIf(IsNull(TmpRs!MoisturePer), "", TmpRs!MoisturePer)
    MSHFlexGrid1.TextMatrix(I, 47) = IIf(IsNull(TmpRs!CITF_MoisturePer), "", TmpRs!CITF_MoisturePer)
    
    If IsNull(TmpRs!WHRNo) = False Then
       MSHFlexGrid1.TextMatrix(I, 48) = IIf(IsNull(TmpRs!CMPID), 0, TmpRs!CMPID) & "-" & TmpRs!WHRNo - (IIf(IsNull(TmpRs!CMPID), 0, TmpRs!CMPID) * 1000000)
    End If
    MSHFlexGrid1.TextMatrix(I, 49) = MSHFlexGrid1.TextMatrix(I, 4) & MSHFlexGrid1.TextMatrix(I, 5) & MSHFlexGrid1.TextMatrix(I, 6)
    
    'x = x + 1
    TmpRs.MoveNext
Next

RsMst_GSL.Close
Call TableMst_GSL(" Order By sm_prefix,CMPID,GodownID,StackNo,LotNo ")
MsgBox "Done !!! "
End Sub



Private Sub CmdStackID_Click()
 FrmMst_Stack.Show
End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean

'If LCase(Gen_UserRole) = "power" Then
   Frame2.Visible = True
'Else
'   Frame2.Visible = False
'End If

FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Call GButtonLockAD(Me)
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If
Dim wc As Variant
wc = ""
If Gen_WcCMP <> "" Then
   wc = "Where " & Gen_WcCMP
End If

Call TableMst_GSL(wc & " Order By sm_prefix,CMPID,GodownID,StackNo,LotNo ")

Call Grid_Head
CmbPreFix.Text = ""
CmbSCmp.Text = ""
CmbSGodown.Text = ""
CmbStack.Text = ""
CmbSCommodity.Text = ""
CmbSFlag.Text = ""
CmbSSCMFlag.Text = ""
CmbSSplilageFlag.Text = ""
CmbSExchangeType.Text = ""
TxtSClientName.Text = ""
TxtSISINNo.Text = ""
TxtSBalBags.Text = ""
TxtSBalWt.Text = ""
TxtSGSLID.Text = ""
Check2.Value = 0
Check2_Click

Call DetailGrid_Head
If RsMst_GSL.RecordCount = 0 Then
   If UserType = "Auditor" Then
      Call GButtonLock(Me, False)
      SaveCmd.Enabled = False
      SearchCmd.Enabled = False
      Exit Sub
   End If
   Call AddCmd_Click
   Exit Sub
End If

' This code will display Record from Transaction

If Gen_GSLID(0) = False Then
   tmp = ""
   If Gen_GSLID(1) <> "" Then
      tmp = tmp & "CMPID = " & Gen_GSLID(1)
   End If
   If Gen_GSLID(2) <> "" Then
      tmp = tmp & " And GodownID = " & Gen_GSLID(2)
   End If
   If Gen_GSLID(3) <> "" Then
      tmp = tmp & " And StackNo = '" & Gen_GSLID(3) & "'"
   End If
   If Gen_GSLID(4) <> "" Then
      tmp = tmp & " And LotNo = '" & Gen_GSLID(4) & "'"
   End If
   RsMst_GSL.Filter = tmp
   If RsMst_GSL.RecordCount > 0 Then
      tmp = RsMst_GSL!Mst_GSL_ID
      RsMst_GSL.Filter = ""
      RsMst_GSL.MoveFirst
      RsMst_GSL.Find "Mst_GSL_ID = " & tmp
   Else
      RsMst_GSL.Filter = ""
      RsMst_GSL.MoveFirst
   End If
End If
Call GButtonLock(Me, True)
Call Indata

End Sub
Private Sub CmdExcel_Click()

Dim ans As Variant


Gen_mTimeStamp = GetTimeStamp


ans = MsgBox("Do you want to generate CSV for searched result? Click 'Yes' for CSV and 'No' for Excel generation!!!", vbYesNo)
If ans = vbYes Then
   Call GenerateCSV
Else
   Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Mst_GSL.xls")
   Call Xls_Detail_Rpt
End If


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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Mst_GSL.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = I + 1 ' + 6
    For C = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(I, C), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(I, C)) Then
              oWS.Cells(mRow, C + 1) = Format(MSHFlexGrid1.TextMatrix(I, C), "MMM/dd/yyyy")
           Else
              If C >= 3 And C <= 6 Then
                  oWS.Cells(mRow, C + 1) = "'" & MSHFlexGrid1.TextMatrix(I, C)
              Else
                 oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
              End If
           End If
        Else
           If C >= 3 And C <= 6 Then
              oWS.Cells(mRow, C + 1) = "'" & MSHFlexGrid1.TextMatrix(I, C)
           ElseIf C = 49 Then
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



Private Sub Form_Resize()
If Me.WindowState = vbMaximized Then
   SetScrollBars
End If
End Sub

Private Sub MSHFlexGrid3_Click()
Dim I, mG_UID As Variant
I = MSHFlexGrid3.RowSel
If I <= 0 Then Exit Sub
If MSHFlexGrid3.TextMatrix(I, 0) = "" Then Exit Sub
If LCase(MSHFlexGrid3.TextMatrix(I, 0)) <> "transfer-from" And LCase(MSHFlexGrid3.TextMatrix(I, 0)) <> "transfer-to" And LCase(MSHFlexGrid3.TextMatrix(I, 0)) <> "gslmadeupbags" And LCase(MSHFlexGrid3.TextMatrix(I, 0)) <> "originalgsl" And LCase(MSHFlexGrid3.TextMatrix(I, 0)) <> "pwr" Then Exit Sub
RsMst_GSL.MoveFirst
RsMst_GSL.Filter = "Mst_GSL_ID = " & (MSHFlexGrid3.TextMatrix(I, 6)) & " And SM_prefix = '" & TxtPreFix & "'"
If RsMst_GSL.EOF = False Then
   mG_UID = RsMst_GSL!G_UID
End If
RsMst_GSL.Filter = ""
RsMst_GSL.MoveFirst
RsMst_GSL.Find "g_uid = '" & mG_UID & "'"
If Not RsMst_GSL.EOF Then
   Call Indata
End If


End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_GSL.MoveNext
If RsMst_GSL.EOF Then
   RsMst_GSL.MoveLast
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
RsMst_GSL.MovePrevious
If RsMst_GSL.BOF Then
   RsMst_GSL.MoveFirst
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
RsMst_GSL.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsMst_GSL.MoveLast
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
   If RsMst_GSL.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   'Call FirstCmd_Click
   Call Indata
   Exit Sub
End If
Frame0.Height = Frame4.Height - 200
Frame0.Left = Frame1.Left - 100
Frame0.Top = Frame1.Top - 100
Frame0.Width = Frame4.Width - 300
Frame0.Visible = True
MSHFlexGrid1.Width = Frame0.Width - 400
MSHFlexGrid1.Height = Frame0.Height - 2400
RsMst_GSL.MoveFirst
CmdExcel.Visible = False
Check2.Visible = False
Check3.Visible = False
'If LCase(Gen_UserRole) = "power" Then
CmdExcel.Visible = True
Check2.Visible = True
Check3.Visible = True
'End If
''===========comment on 7 july 2012 for not clearing search grid=============
'MSHFlexGrid1.Clear
'MSHFlexGrid1.Cols = 3
'Call Grid_Head
'CmbPreFix.Text = ""
'CmbSCmp.Text = ""
'CmbSGodown.Text = ""
'CmbStack.Text = ""
'CmbSCommodity.Text = ""
'CmbSFlag.Text = ""
'CmbSSCMFlag.Text = ""
'CmbSSplilageFlag.Text = ""
'CmbSExchangeType.Text = ""
'TxtSClientName.Text = ""
'TxtSISINNo.Text = ""
'TxtSBalBags.Text = ""
'TxtSBalWt.Text = ""
'TxtSGSLID.Text = ""
'Check2.Value = 0
'Check2_Click
''===========comment on 7 july 2012 for not clearing search grid=============


MSHFlexGrid1.ColWidth(0) = 1000
Dim ans As Variant

End Sub
Public Sub Indata()

Frame5.Visible = False
TxtGSLID.Locked = True
TxtCMPName.Locked = True
CmbCMPName.Visible = False
CmdCMPName.Visible = False
TxtCMPName.Visible = True
TxtCommodityID.Locked = True
TxtCommodityID.Visible = True
TxtGodownID.Locked = True
CmbGodownID.Visible = False
CmdGodownID.Visible = False
TxtGodownID.Visible = True

CmbStackID.Visible = False
CmdStackID.Visible = False
TxtStackNo.Visible = True
TxtStackNo.Locked = True

TxtLotNo.Locked = True

TxtExchangeTypeID.Locked = True
TxtExchangeTypeID.Visible = True
TxtClientID.Locked = True
TxtClientID.Visible = True
'TxtPledgeNo.Locked = True
TxtFlag.Locked = True
TxtISINID.Locked = True
TxtISINID.Visible = True

TxtGSLID = IIf(IsNull(RsMst_GSL!Mst_GSL_ID), "", RsMst_GSL!Mst_GSL_ID)
TxtPreFix = IIf(IsNull(RsMst_GSL!SM_Prefix), "", RsMst_GSL!SM_Prefix)

TxtCommodityID = ""
TxtExchangeTypeID = ""
TxtISINID = ""
TxtClientID = ""
mCMPID = IIf(IsNull(RsMst_GSL!CMPID), "", RsMst_GSL!CMPID)
mCMPName = IIf(IsNull(RsMst_GSL!CMPID), "", RsMst_GSL!CMPID)
TxtCMPName = GetCMPName(RsMst_GSL!CMPID) 'IIf(IsNull(RsMst_GSL!CMPID), "", RsMst_GSL!CMPID)
If RsMst_GSL!CommodityID <> "" Then
   mCommodityID = IIf(IsNull(RsMst_GSL!CommodityID), "", RsMst_GSL!CommodityID)
   TxtCommodityID = GetCommodityName(IIf(IsNull(RsMst_GSL!CommodityID), "0", RsMst_GSL!CommodityID))
End If
mGodownID = IIf(IsNull(RsMst_GSL!GodownID), "", RsMst_GSL!GodownID)
TxtGodownID = GetGodownName(RsMst_GSL!GodownID, "N") 'IIf(IsNull(RsMst_GSL!GodownNo), "", RsMst_GSL!GodownNo)

TxtStackNo = IIf(IsNull(RsMst_GSL!StackNo), "", RsMst_GSL!StackNo)
CmbStackID = TxtStackNo

TxtLotNo = IIf(IsNull(RsMst_GSL!LotNo), "", RsMst_GSL!LotNo)
TxtDepositeWeight = IIf(IsNull(RsMst_GSL!DepositeWeight), "", RsMst_GSL!DepositeWeight)
TxtDepositeBags = IIf(IsNull(RsMst_GSL!DepositeBags), "", RsMst_GSL!DepositeBags)
TxtWHRBags = IIf(IsNull(RsMst_GSL!WHRBags), "", RsMst_GSL!WHRBags)
TxtCDTFBags = IIf(IsNull(RsMst_GSL!CDTFBags), "", RsMst_GSL!CDTFBags)
TxtBalanceBags = IIf(IsNull(RsMst_GSL!BalanceBags), "", RsMst_GSL!BalanceBags)
TxtWHRWeight = IIf(IsNull(RsMst_GSL!WHRWeight), "", RsMst_GSL!WHRWeight)
TxtCDTFWeight = IIf(IsNull(RsMst_GSL!CDTFWeight), "", RsMst_GSL!CDTFWeight)
TxtBalanceWeight = IIf(IsNull(RsMst_GSL!BalanceWeight), "", RsMst_GSL!BalanceWeight)
TxtMoisturePer = IIf(IsNull(RsMst_GSL!MoisturePer), "", RsMst_GSL!MoisturePer)
TxtCITFMoisturePer = IIf(IsNull(RsMst_GSL!CITF_MoisturePer), "", RsMst_GSL!CITF_MoisturePer)

If IsNull(RsMst_GSL!SpillageFlag) Or RsMst_GSL!SpillageFlag = "0" Then
   Check1.Value = 0
Else
   Check1.Value = 1
End If

'If IsNull(RsMst_GSL!SCMFlag) Or RsMst_GSL!SCMFlag = "0" Then
'   ChkSCMFlag.Value = 0
'Else
'   ChkSCMFlag.Value = 1
'End If

ChkSCMFlag.Value = IIf(IsNull(RsMst_GSL!SCMFlag), 0, RsMst_GSL!SCMFlag)
mMadeupFlag = IIf(IsNull(RsMst_GSL!MadeupFlag), 0, RsMst_GSL!MadeupFlag)

ChkSCMFlag.Enabled = False
Check1.Enabled = False
mExchangeTypeID = IIf(IsNull(RsMst_GSL!ExchangeTypeID), "", RsMst_GSL!ExchangeTypeID)
If mExchangeTypeID <> "" Then
   TxtExchangeTypeID = GetExchangeName(RsMst_GSL!ExchangeTypeID) 'IIf(IsNull(RsMst_GSL!ExchangeType), "", RsMst_GSL!ExchangeType)
End If

mClientIDRow = 0
If Not IsNull(RsMst_GSL!ClientID) Then
   mClientIDRow = IIf(IsNull(RsMst_GSL!ClientID), 0, RsMst_GSL!ClientID)
End If
If mClientIDRow <> 0 Then
   TxtClientID = GetClientName(RsMst_GSL!ClientID) '(IsNull(RsMst_GSL!ClientName), "", RsMst_GSL!ClientName)
End If
TxtPledgeNo = IIf(IsNull(RsMst_GSL!PledgeNo), "", RsMst_GSL!PledgeNo)
TxtFlag = IIf(IsNull(RsMst_GSL!Status), "", RsMst_GSL!Status)

mISINID = IIf(IsNull(RsMst_GSL!ISINID), "0", RsMst_GSL!ISINID)
TxtISINID = IIf(IsNull(RsMst_GSL!ISINID), "0", RsMst_GSL!ISINID)

Call DetailGrid_Head
Call MadeUpDetailGrid_Head

MSHFlexGrid2.TextMatrix(1, 1) = IIf(IsNull(RsMst_GSL!DepositeBags), "", RsMst_GSL!DepositeBags)
MSHFlexGrid2.TextMatrix(1, 2) = IIf(IsNull(RsMst_GSL!DepositeWeight), "", RsMst_GSL!DepositeWeight)
MSHFlexGrid2.TextMatrix(2, 1) = IIf(IsNull(RsMst_GSL!CDTFBags), "0", RsMst_GSL!CDTFBags) '"WHR Withdrawn Against"
MSHFlexGrid2.TextMatrix(2, 2) = IIf(IsNull(RsMst_GSL!CDTFWeight), "", RsMst_GSL!CDTFWeight) '"WHR Withdrawn Against"
MSHFlexGrid2.TextMatrix(3, 1) = IIf(IsNull(RsMst_GSL!BalanceBags), "0", RsMst_GSL!BalanceBags) '"Balance Bags"
MSHFlexGrid2.TextMatrix(3, 2) = IIf(IsNull(RsMst_GSL!BalanceWeight), "0", RsMst_GSL!BalanceWeight) '"Balance Weight

mConnectedGSLID = IIf(IsNull(RsMst_GSL!Connected_Mst_GSL_ID), "0", RsMst_GSL!Connected_Mst_GSL_ID)

If mMadeupFlag = 0 And mConnectedGSLID <> 0 Then
   tmp = "Select DepositeBags,DepositeWeight,CDTFBags,CDTFWeight,BalanceBags,BalanceWeight from Mst_GSL Where Mst_GSL_ID = " & RsMst_GSL!Connected_Mst_GSL_ID & " And SM_Prefix = '" & TxtPreFix & "' "
   Call TmpTable
   
   If TmpRs.RecordCount > 0 Then
      Frame5.Visible = True
      MSHFlexGrid4.TextMatrix(1, 1) = IIf(IsNull(TmpRs!DepositeBags), "", TmpRs!DepositeBags)
      MSHFlexGrid4.TextMatrix(1, 2) = IIf(IsNull(TmpRs!DepositeWeight), "", TmpRs!DepositeWeight)
      MSHFlexGrid4.TextMatrix(2, 1) = IIf(IsNull(TmpRs!CDTFBags), "0", TmpRs!CDTFBags) '"WHR Withdrawn Against"
      MSHFlexGrid4.TextMatrix(2, 2) = IIf(IsNull(TmpRs!CDTFWeight), "", TmpRs!CDTFWeight) '"WHR Withdrawn Against"
      MSHFlexGrid4.TextMatrix(3, 1) = IIf(IsNull(TmpRs!BalanceBags), "0", TmpRs!BalanceBags) '"Balance Bags"
      MSHFlexGrid4.TextMatrix(3, 2) = IIf(IsNull(TmpRs!BalanceWeight), "0", TmpRs!BalanceWeight) '"Balance Weight
   End If
End If

TxtCreated = IIf(IsNull(RsMst_GSL!CreatedBy), "", RsMst_GSL!CreatedBy) & " :- " & IIf(IsNull(RsMst_GSL!CreatedDate), "", RsMst_GSL!CreatedDate)
TxtUpdated = IIf(IsNull(RsMst_GSL!UpdatedBy), "", RsMst_GSL!UpdatedBy) & " :- " & IIf(IsNull(RsMst_GSL!UpdatedDate), "", RsMst_GSL!UpdatedDate)


For I = 1 To 3
    MSHFlexGrid2.TextMatrix(I, 2) = Format(MSHFlexGrid2.TextMatrix(I, 2), "########0.000") '"Balance WHR Against"
Next


'If LCase(Gen_UserRole) = "power" Then
   Call Grid_Head1
   Call ShowGSLTransactionDetails
'Else
'   Call Grid_Head1
'End If

If RsMst_GSL.RecordCount = 1 Then
   Call GButtonLock_1(Me, True)
Else
   Call GButtonLock(Me, True)
End If

End Sub
Private Sub ShowGSLTransactionDetails()
'
'tmp = "select 'CxTF' tt,CxTFNo,NoOfBags,DematWeight,txntype,0 GSLID " & _
'      " from Txn_CxTF_GSL where gslid = " & TxtGSLID & "   and sm_prefix = '" & TxtPreFix & "' " & _
'      " Union All " & _
'      " select 'Adjustment' tt ,a.txnid,a.NoOfBags,a.Quantity,b.TxnType,0 GSLID " & _
'      " from Txn_GSLAdjTransaction a " & _
'      " inner join Txn_AdjTransaction b on a.txnid = b.txnid " & _
'      " Where a.gslid = " & TxtGSLID & " and b.sm_prefix = '" & TxtPreFix & "' " & _
'      " Union All " & _
'      " Select 'Transfer-From' tt,GSLTransferID ,0 NoOfBags,0 DematWeight,'' txntype,OldGSLID  from Txn_GSL_Transfer where newgslid = " & TxtGSLID & " and sm_prefix = '" & TxtPreFix & "'  " & _
'      " Union All " & _
'      " Select 'Transfer-To' tt,GSLTransferID ,0 NoOfBags,0 DematWeight,'' txntype,NewGSLID  from Txn_GSL_Transfer where oldgslid = " & TxtGSLID & " and sm_prefix = '" & TxtPreFix & "' " & _
'      " Union All " & _
'      " Select 'Spillage' , Spillageid , SpillageNoOfBags,SpillageQuantity , '' Txntype , 0 GSLID " & _
'      " From Txn_Spillage where gslid = " & TxtGSLID & " and sm_prefix = '" & TxtPreFix & "' "
'

tmp = "Select 'CxTF' tt,tcg.CxTFNo,tcg.NoOfBags,tcg.DematWeight,tcg.txntype,tcd.CxTFdate 'TxnDate',0 Mst_GSL_ID"
tmp = tmp & " From Txn_CxTF_GSL TCG "
tmp = tmp & " Inner Join  Txn_CxTF_Details TCD On TCG.CxTFNo = TCD.CxTFNo And TCG.SM_Prefix = TCD.SM_prefix And TCG.TxnType = TCD.TxnType"
tmp = tmp & " Where TCG.Mst_GSL_ID = " & TxtGSLID & "   and TCG.sm_prefix = '" & TxtPreFix & "' "
tmp = tmp & " Union All"
tmp = tmp & " Select 'Adjustment' tt ,a.txnid,a.NoOfBags,a.Quantity,b.TxnType,b.TxnDate 'TxnDate',0 Mst_GSL_ID"
tmp = tmp & " From Txn_GSLAdjTransaction a  inner join Txn_AdjTransaction b on a.txnid = b.txnid"
tmp = tmp & " Where a.Mst_GSL_ID = " & TxtGSLID & " and b.sm_prefix = '" & TxtPreFix & "'"
tmp = tmp & " Union All"
tmp = tmp & " Select 'Transfer-From' tt,GSLTransferID ,0 NoOfBags,0 DematWeight,'' txntype,GSLTransferDate 'TxnDate',Old_Mst_GSL_ID"
tmp = tmp & " From Txn_GSL_Transfer"
tmp = tmp & " Where New_Mst_GSL_ID = " & TxtGSLID & " and sm_prefix = '" & TxtPreFix & "'"
tmp = tmp & " Union All"
tmp = tmp & " Select 'Transfer-To' tt,GSLTransferID ,0 NoOfBags,0 DematWeight,'' txntype,GSLTransferDate 'TxnDate',New_Mst_GSL_ID"
tmp = tmp & " From Txn_GSL_Transfer"
tmp = tmp & " Where Old_Mst_GSL_ID = " & TxtGSLID & " and sm_prefix = '" & TxtPreFix & "'"
tmp = tmp & " Union All  Select 'Spillage' , Spillageid , SpillageNoOfBags,SpillageQuantity , '' Txntype ,SpillageDate 'TxnDate', 0 GSLID"
tmp = tmp & " From Txn_Spillage where Mst_GSL_ID = " & TxtGSLID & "   and sm_prefix = '" & TxtPreFix & "' And Flag = 'S' "
tmp = tmp & " Union All  Select 'MadeUpBags' , Spillageid , SpillageNoOfBags,SpillageQuantity , '' Txntype ,SpillageDate 'TxnDate', 0 GSLID"
tmp = tmp & " From Txn_Spillage where NewMst_GSL_ID = " & TxtGSLID & "   and sm_prefix = '" & TxtPreFix & "' And Flag = 'M' "
tmp = tmp & " Union All"
tmp = tmp & " Select 'GSLMadeUpBags' , Mst_GSL_ID , DepositeBags,DepositeWeight , '' Txntype ,'' 'TxnDate', Mst_GSL_ID"
tmp = tmp & " From Mst_GSL"
tmp = tmp & " where Connected_Mst_GSL_ID = " & TxtGSLID & "   and sm_prefix = '" & TxtPreFix & "' And MadeUpFlag = '1'"
tmp = tmp & " Union All"
tmp = tmp & " Select 'OriginalGSL' , Mst_GSL_ID , DepositeBags,DepositeWeight , '' Txntype ,'' 'TxnDate', Mst_GSL_ID"
tmp = tmp & " From Mst_GSL"
tmp = tmp & " where Connected_Mst_GSL_ID = " & TxtGSLID & "   and sm_prefix = '" & TxtPreFix & "' And MadeUpFlag = '0'"
tmp = tmp & " Union All"
tmp = tmp & " Select 'Transfer-To' tt,Mst_GSL_ID ,0 NoOfBags,0 DematWeight,'' txntype,'' 'TxnDate',Mst_GSL_ID"
tmp = tmp & " From Mst_GSL"
tmp = tmp & " Where BGSLID_Mst_GSL_ID = " & TxtGSLID & " and sm_prefix = '" & TxtPreFix & "' And MadeUpFlag = '1'"
tmp = tmp & " Union All"
tmp = tmp & " Select 'Transfer-From' tt,Mst_GSL_ID ,0 NoOfBags,0 DematWeight,'' txntype,'' 'TxnDate',Mst_GSL_ID"
tmp = tmp & " From Mst_GSL"
tmp = tmp & " Where FGSLID_Mst_GSL_ID = " & TxtGSLID & " and sm_prefix = '" & TxtPreFix & "' And MadeUpFlag = '1'"
tmp = tmp & " Union All"
tmp = tmp & " Select 'PWR' tt,Mst_GSL_ID ,AdjstBags NoOfBags,AdjstQty DematWeight,'W' txntype,Concurrency"
tmp = tmp & "  'TxnDate',Mst_GSL_ID"
tmp = tmp & " From Txn_RRNPWRGSL Where Mst_GSL_ID = " & TxtGSLID & " and sm_prefix = '" & TxtPreFix & "'"
tmp = tmp & " Union All"
tmp = tmp & " Select 'PWR' tt,NewMst_GSL_ID ,AdjstBags NoOfBags,AdjstQty DematWeight,'D' txntype,Concurrency"
tmp = tmp & " 'TxnDate',Mst_GSL_ID"
tmp = tmp & " From Txn_RRNPWRGSL Where NewMst_GSL_ID = " & TxtGSLID & " and sm_prefix = '" & TxtPreFix & "'"

Call TmpTable

For I = 1 To TmpRs.RecordCount
    MSHFlexGrid3.Rows = TmpRs.RecordCount + 1
    MSHFlexGrid3.TextMatrix(I, 0) = IIf(IsNull(TmpRs!TT), "", TmpRs!TT)
    MSHFlexGrid3.TextMatrix(I, 1) = IIf(IsNull(TmpRs!CxTFNO), "", TmpRs!CxTFNO)
    MSHFlexGrid3.TextMatrix(I, 2) = IIf(IsNull(TmpRs!NoofBags), 0, TmpRs!NoofBags)
    MSHFlexGrid3.TextMatrix(I, 3) = IIf(IsNull(TmpRs!DematWeight), 0, TmpRs!DematWeight)
    MSHFlexGrid3.TextMatrix(I, 4) = IIf(IsNull(TmpRs!TxnType), "", TmpRs!TxnType)
    If LCase(TmpRs!TT) = "gslmadeupbags" Or LCase(TmpRs!TT) = "originalgsl" Then
       MSHFlexGrid3.TextMatrix(I, 5) = ""
    Else
       MSHFlexGrid3.TextMatrix(I, 5) = IIf(IsNull(TmpRs!TxnDate), "", TmpRs!TxnDate)
    End If
    If MSHFlexGrid3.TextMatrix(I, 5) <> "" Then
       MSHFlexGrid3.TextMatrix(I, 5) = Format(MSHFlexGrid3.TextMatrix(I, 5), "dd-mmm-yyyy")
    End If
    MSHFlexGrid3.TextMatrix(I, 6) = IIf(IsNull(TmpRs!Mst_GSL_ID), "", TmpRs!Mst_GSL_ID)
    TmpRs.MoveNext
Next

End Sub

Private Sub TxtGSLID_LostFocus()
If TxtGSLID = "" Then Exit Sub
If IsNumeric(TxtGSLID) = False Then
   MsgBox "Invalid numeric format !!!!"
   TxtGSLID.SetFocus
   Exit Sub
End If
TxtGSLID = Val(TxtGSLID)
End Sub
Private Sub TxtDepositeWeight_LostFocus()
If TxtDepositeWeight = "" Then Exit Sub
If IsNumeric(TxtDepositeWeight) = False Then
   MsgBox "Invalid numeric format !!!!"
   TxtDepositeWeight.SetFocus
   Exit Sub
End If
TxtDepositeWeight = Val(TxtDepositeWeight)
End Sub
Private Sub TxtWHRWeight_LostFocus()
If TxtWHRWeight = "" Then Exit Sub
If IsNumeric(TxtWHRWeight) = False Then
   MsgBox "Invalid numeric format !!!!"
   TxtWHRWeight.SetFocus
   Exit Sub
End If
TxtWHRWeight = Val(TxtWHRWeight)
End Sub
Private Sub TxtCDTFWeight_LostFocus()
If TxtCDTFWeight = "" Then Exit Sub
If IsNumeric(TxtCDTFWeight) = False Then
   MsgBox "Invalid numeric format !!!!"
   TxtCDTFWeight.SetFocus
   Exit Sub
End If
TxtCDTFWeight = Val(TxtCDTFWeight)
End Sub
Private Sub TxtBalanceWeight_LostFocus()
If TxtBalanceWeight = "" Then Exit Sub
If IsNumeric(TxtBalanceWeight) = False Then
   MsgBox "Invalid numeric format !!!!"
   TxtBalanceWeight.SetFocus
   Exit Sub
End If
TxtBalanceWeight = Val(TxtBalanceWeight)
End Sub
Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)

TxtGSLID.Locked = True
TxtCMPName.Locked = False
CmbCMPName.Visible = True
CmdCMPName.Visible = True
TxtCMPName.Visible = False

TxtGodownID.Locked = False
CmbGodownID.Visible = True
CmdGodownID.Visible = True
TxtGodownID.Visible = False

CmbStackID.Text = ""
CmbStackID.Visible = True
CmdStackID.Visible = True
TxtStackNo.Locked = True
TxtStackNo.Visible = False

TxtLotNo.Locked = False
TxtExchangeTypeID.Locked = True
TxtExchangeTypeID.Visible = True
TxtClientID.Locked = True
TxtFlag.Locked = False
TxtISINID.Locked = True

TxtCreated = ""
TxtUpdated = ""
TxtPreFix = ""
TxtGSLID = ""
TxtCMPName = ""
TxtCommodityID = ""
TxtGodownID = ""
TxtStackNo = ""
TxtLotNo = ""
TxtDepositeWeight = ""
TxtDepositeBags = ""
TxtWHRBags = ""
TxtCDTFBags = ""
TxtBalanceBags = ""
TxtPledgeNo = ""
TxtWHRWeight = ""
TxtCDTFWeight = ""
TxtBalanceWeight = ""
TxtExchangeTypeID = ""
TxtClientID = ""
TxtFlag = "B"
TxtISINID = ""
CmbCMPName.Text = ""
CmbGodownID.Text = ""
Check1.Enabled = True
ChkSCMFlag.Enabled = False
ChkSCMFlag.Value = 0
TxtMoisturePer.Text = ""
TxtCITFMoisturePer.Text = ""

If Gen_GSLID(0) = False Then
   mCMPName = Gen_GSLID(1)
   mGodownID = Gen_GSLID(2)
   CmbCMPName.Text = Gen_GSLID(5)
   CmbGodownID.Text = Gen_GSLID(6)
   TxtStackNo = Gen_GSLID(3)
   TxtLotNo.Text = GetNextLotNo(mGodownID, TxtStackNo)
End If

MSHFlexGrid2.TextMatrix(1, 1) = 0
MSHFlexGrid2.TextMatrix(1, 2) = 0
MSHFlexGrid2.TextMatrix(2, 1) = 0
MSHFlexGrid2.TextMatrix(2, 2) = 0
MSHFlexGrid2.TextMatrix(3, 1) = 0
MSHFlexGrid2.TextMatrix(3, 2) = 0

Call Grid_Head1

Frame5.Visible = False

If RsMst_GSL.RecordCount > 0 Then
   CmbCMPName.SetFocus
End If
End Sub
Private Sub EditCmd_Click()

If TxtCommodityID <> "" Then
   MsgBox "You can't edit this GSL Detail "
   Exit Sub
End If

Call TableTxn_CxTF_GSL("where sm_prefix = '" & TxtPreFix & "' And Mst_GSL_ID = " & TxtGSLID)
If RsTxn_CxTF_GSL.RecordCount > 0 Then
   MsgBox "GSL used in Deposit/Withdrawal Transaction so you can't edit this record "
   Exit Sub
End If

Call TableTxn_Spillage("Where sm_prefix = '" & TxtPreFix & "' And Mst_GSL_ID = " & TxtGSLID)
If RsTxn_Spillage.RecordCount > 0 Then
   MsgBox "GSL used in Spillage Transaction so you can't edit this record "
   Exit Sub
End If
Edit_Flg = "E"

Call GButtonLock(Me, False)

TxtGSLID.Locked = True
TxtCMPName.Locked = True
TxtCMPName.Visible = True
TxtGodownID.Locked = True
TxtGodownID.Visible = True

CmbStackID = TxtStackNo
CmbStackID.Visible = True
TxtStackNo.Visible = False
TxtStackNo.Locked = True

TxtLotNo.Locked = False

TxtFlag.Locked = False
Check1.Enabled = True
ChkSCMFlag.Enabled = False

End Sub
Private Sub DeleteCmd_Click()
If TxtFlag = "E" Then
   MsgBox "You Can't Delete this GSL !!!!! "
   Exit Sub
End If
If TxtCommodityID <> "" Then
   MsgBox "You Can't Delete this GSL !!!!! "
   Exit Sub
End If
Dim ans As Variant
On Error GoTo msgError
ans = MsgBox("Do you really want to DELETE this record???", vbYesNo)
If ans = vbYes Then
   RsMst_GSL.Delete
   RsMst_GSL.Update
   If RsMst_GSL.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_GSL.MoveNext
   If RsMst_GSL.EOF() Then
      RsMst_GSL.MoveLast
   End If
   Call Indata
End If
Exit Sub
msgError:
MsgBox "Child record Present "
RsMst_GSL.CancelUpdate
End Sub
Private Sub SaveCmd_Click()

If CmbStackID.Text = "" Then
   MsgBox "Blank StackNo Not Allowed !!! "
   CmbStackID.SetFocus
   Exit Sub
End If

'If TxtStackNo.Text = "" Then
'   MsgBox "Blank StackNo Not Allowed !!! "
'   TxtStackNo.SetFocus
'   Exit Sub
'End If

If TxtPreFix = "" Then
   MsgBox "Blank Pre-fix is not Allowed !!! Contact Technology Department !! "
   Exit Sub
End If

If TxtLotNo.Text = "" Then
   MsgBox "Blank LotNo Not Allowed !!! "
   TxtLotNo.SetFocus
   Exit Sub
Else
   If InStr(1, LCase(TxtLotNo), "mub") > 0 Then
      MsgBox "'MUB' word not allowed in Lot No !!! "
      TxtLotNo.SetFocus
      Exit Sub
   End If
   If InStr(1, LCase(TxtLotNo), "con") > 0 Then
      MsgBox "'CON' word not allowed in Lot No !!! "
      TxtLotNo.SetFocus
      Exit Sub
   End If
End If


If Edit_Flg = "A" Then
   If RsMst_GSL.RecordCount > 0 Then
      RsMst_GSL.MoveFirst
      tmp = "select * from Mst_GSL where sm_prefix = '" & TxtPreFix & "' And CMPID = " & mCMPName & " And GodownID = " & mGodownID & " And StackNo = '" & CmbStackID.Text & "' And LotNo = '" & TxtLotNo & "'"
      Call Tmp3Table
      
      If TmpRs3.RecordCount > 0 Then
        MsgBox "Duplicate GSLID  Not Allowed !!! "
         TxtLotNo.SetFocus
         RsMst_GSL.Filter = ""
         Exit Sub
      End If
   End If
   RsMst_GSL.AddNew
   'RsMst_GSL!Mst_GSL_ID = GetGSLID
   RsMst_GSL!G_UID = GetGUID
   mGUIDForID = RsMst_GSL!G_UID
   'TxtGSLID.Text = RsMst_GSL!GSLID
Else
    RsMst_GSL.MoveFirst
    tmp = "sm_prefix = '" & TxtPreFix & "' And CMPID = " & mCMPName & " And GodownID = " & mGodownID & " And StackNo = '" & CmbStackID.Text & "' And LotNo = '" & TxtLotNo & "'"
    RsMst_GSL.Filter = tmp
    If Not RsMst_GSL.EOF Then
       If RsMst_GSL!Mst_GSL_ID <> TxtGSLID.Text Then
          MsgBox "Duplicate GSLID  Not Allowed !!! "
          TxtLotNo.SetFocus
          Exit Sub
       End If
    End If
    tmp = "sm_prefix = '" & TxtPreFix & "'"
    RsMst_GSL.Filter = tmp
    RsMst_GSL.MoveFirst
    RsMst_GSL.Find " Mst_GSL_ID= " & Val(TxtGSLID.Text)
End If

'RsMst_GSL!GSLID = TxtGSLID.Text
RsMst_GSL!SM_Prefix = TxtPreFix
RsMst_GSL!CMPID = mCMPName
RsMst_GSL!GodownID = mGodownID
RsMst_GSL!StackNo = CmbStackID.Text  ''TxtStackNo.Text
RsMst_GSL!LotNo = TxtLotNo.Text
RsMst_GSL!SpillageFlag = Check1.Value
RsMst_GSL!PledgeNo = TxtPledgeNo.Text
RsMst_GSL!Status = TxtFlag.Text
RsMst_GSL!SCMFlag = 0
RsMst_GSL!MadeupFlag = 0
RsMst_GSL!GSLID = 0


If IsNull(RsMst_GSL!CreatedBy) = True Or RsMst_GSL!CreatedBy = "" Then
   RsMst_GSL!CreatedBy = Gen_UserLoginID
   RsMst_GSL!CreatedDate = Now
Else
   RsMst_GSL!UpdatedBy = Gen_UserLoginID
   RsMst_GSL!UpdatedDate = Now
End If

RsMst_GSL.Update

tmp = "Select Mst_GSL_ID from Mst_GSL Where G_UID = '" & mGUIDForID & "'"
Call TmpTable

If TmpRs.RecordCount > 0 Then
   TxtGSLID.Text = TmpRs!Mst_GSL_ID
End If

RsMst_GSL.Filter = ""
RsMst_GSL.MoveFirst
RsMst_GSL.Find " sm_prefix = '" & TxtPreFix.Text & "'"
RsMst_GSL.Find " Mst_GSL_ID= " & TxtGSLID.Text

Call Indata

End Sub
Private Function GetGSLID() As Double
Dim Retval As Double
tmp = "Select max(GSLID) from Mst_GSL where SM_Prefix = '" & TxtPreFix & "'"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetGSLID = Retval
End Function
Private Sub Grid_Head()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 3
MSHFlexGrid1.Cols = 50 '49 ' 48

MSHFlexGrid1.TextMatrix(1, 0) = "GSLID"
MSHFlexGrid1.TextMatrix(1, 1) = "Location ID"
MSHFlexGrid1.TextMatrix(1, 2) = "CMP Name"
MSHFlexGrid1.TextMatrix(1, 3) = "Commodity"
MSHFlexGrid1.TextMatrix(1, 4) = "Godown No"
MSHFlexGrid1.TextMatrix(1, 5) = "Stack No"
MSHFlexGrid1.TextMatrix(1, 6) = "Lot No"
MSHFlexGrid1.TextMatrix(1, 7) = "Deposited Bags"
MSHFlexGrid1.TextMatrix(1, 8) = "Withdrawn Bags"
MSHFlexGrid1.TextMatrix(1, 9) = "Balance Bags"
MSHFlexGrid1.TextMatrix(1, 10) = "Deposited Weight"
MSHFlexGrid1.TextMatrix(1, 11) = "Withdrawn Weight"
MSHFlexGrid1.TextMatrix(1, 12) = "Balance Weight"
MSHFlexGrid1.TextMatrix(1, 13) = "Exchange Type"
MSHFlexGrid1.TextMatrix(1, 14) = "Client Name"

MSHFlexGrid1.TextMatrix(1, 15) = "Agent Name"

MSHFlexGrid1.TextMatrix(1, 16) = "Flag"
MSHFlexGrid1.TextMatrix(1, 17) = "CMSE Lot No / CDC No" '"ISIN ID / CDC No"
MSHFlexGrid1.TextMatrix(1, 18) = "CDTF No"
MSHFlexGrid1.TextMatrix(1, 19) = "CDTF Date"
MSHFlexGrid1.TextMatrix(1, 20) = "Pledge No"
MSHFlexGrid1.TextMatrix(1, 21) = "Pledge Bags"
MSHFlexGrid1.TextMatrix(1, 22) = "Pledge Weight"
MSHFlexGrid1.TextMatrix(1, 23) = "Spillage Flag"
MSHFlexGrid1.TextMatrix(1, 24) = "Deposit Adj Wt"
MSHFlexGrid1.TextMatrix(1, 25) = "Withdrawal Adj Wt"
MSHFlexGrid1.TextMatrix(1, 26) = "G_uid"

MSHFlexGrid1.TextMatrix(1, 27) = "KR No"
MSHFlexGrid1.TextMatrix(1, 28) = "Gross Qty"
MSHFlexGrid1.TextMatrix(1, 29) = "From GSL"
MSHFlexGrid1.TextMatrix(1, 30) = "To GSL"
MSHFlexGrid1.TextMatrix(1, 31) = "WHR Bags"
MSHFlexGrid1.TextMatrix(1, 32) = "WHR Qty"

MSHFlexGrid1.TextMatrix(1, 33) = "Current Contract"
MSHFlexGrid1.TextMatrix(1, 34) = "Billing Contract"

MSHFlexGrid1.TextMatrix(1, 35) = "Bag Size"
MSHFlexGrid1.TextMatrix(1, 36) = "Bag Equivalent"
MSHFlexGrid1.TextMatrix(1, 37) = "Weighted Qty"
MSHFlexGrid1.TextMatrix(1, 38) = "Amount Received"
MSHFlexGrid1.TextMatrix(1, 39) = "SC upto Date"
MSHFlexGrid1.TextMatrix(1, 40) = "SCM Lot"
MSHFlexGrid1.TextMatrix(1, 41) = "Made Up Flag"
MSHFlexGrid1.TextMatrix(1, 42) = "Connected GSL ID"
MSHFlexGrid1.TextMatrix(1, 43) = "First CITF Date"
MSHFlexGrid1.TextMatrix(1, 44) = "From Location"
MSHFlexGrid1.TextMatrix(1, 45) = "To Location"
MSHFlexGrid1.TextMatrix(1, 46) = "Moisture %"
MSHFlexGrid1.TextMatrix(1, 47) = "CITF Moisture %"
MSHFlexGrid1.TextMatrix(1, 48) = "WHR No"
MSHFlexGrid1.TextMatrix(1, 49) = "Concatinated GSL data"

MSHFlexGrid1.RowHeight(1) = 500

MSHFlexGrid1.ColWidth(26) = 0

End Sub

Private Sub MSHFlexGrid1_Click()
Dim I As Variant
I = MSHFlexGrid1.RowSel
If I <= 1 Then Exit Sub
If MSHFlexGrid1.TextMatrix(I, 0) = "" Then Exit Sub
RsMst_GSL.MoveFirst

RsMst_GSL.Find "g_uid = '" & (MSHFlexGrid1.TextMatrix(I, 26)) & "'"

If Not RsMst_GSL.EOF Then
   Call Indata
   Frame0.Visible = False
End If
End Sub
Private Sub DetailGrid_Head()
MSHFlexGrid2.Clear
MSHFlexGrid2.Rows = 4
MSHFlexGrid2.Cols = 3
MSHFlexGrid2.TextMatrix(0, 1) = "No Of Bags"
MSHFlexGrid2.TextMatrix(0, 2) = "Weight in MT"
MSHFlexGrid2.TextMatrix(1, 0) = "Deposits"
MSHFlexGrid2.TextMatrix(2, 0) = "Withdrawals "
MSHFlexGrid2.TextMatrix(3, 0) = "Electronic Balance "

Call GridCellColour(1, 1, "&H80000000")
Call GridCellColour(1, 2, "&H80000000")

Call GridCellColour(2, 1, "&H80000000")
Call GridCellColour(2, 2, "&H80000000")

Call GridCellColour(3, 1, "&H80000000")
Call GridCellColour(3, 2, "&H80000000")


MSHFlexGrid2.ColWidth(0) = 2800
MSHFlexGrid2.ColWidth(1) = 1800
MSHFlexGrid2.ColWidth(2) = 1800

End Sub
Public Sub GridCellColour(mRow As Integer, mCol As Integer, colour As Double)
    MSHFlexGrid2.Col = mCol
    MSHFlexGrid2.row = mRow
    MSHFlexGrid2.CellBackColor = colour
End Sub
Private Sub MadeUpDetailGrid_Head()

MSHFlexGrid4.Clear
MSHFlexGrid4.Rows = 4
MSHFlexGrid4.Cols = 3
MSHFlexGrid4.TextMatrix(0, 1) = "No Of Bags"
MSHFlexGrid4.TextMatrix(0, 2) = "Weight in MT"
MSHFlexGrid4.TextMatrix(1, 0) = "Deposits"
MSHFlexGrid4.TextMatrix(2, 0) = "Withdrawals "
MSHFlexGrid4.TextMatrix(3, 0) = "Electronic Balance "

Call GridMCellColour(1, 1, "&H80000000")
Call GridMCellColour(1, 2, "&H80000000")

Call GridMCellColour(2, 1, "&H80000000")
Call GridMCellColour(2, 2, "&H80000000")

Call GridMCellColour(3, 1, "&H80000000")
Call GridMCellColour(3, 2, "&H80000000")


MSHFlexGrid4.ColWidth(0) = 2800
MSHFlexGrid4.ColWidth(1) = 1800
MSHFlexGrid4.ColWidth(2) = 1800

End Sub
Public Sub GridMCellColour(mRow As Integer, mCol As Integer, colour As Double)
    MSHFlexGrid4.Col = mCol
    MSHFlexGrid4.row = mRow
    MSHFlexGrid4.CellBackColor = colour
End Sub

Private Sub CmdCMPName_Click()
    FrmMst_CMP.Show
End Sub
'----
Private Sub CmbCMPName_GotFocus()
tmp = CmbCMPName.Text
tmp1 = ""
If Gen_WcCMP <> "" Then
   tmp1 = "And " & Gen_WcCMP
End If
Call TableMst_CMP(" where CloseDate is null " & tmp1)


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
'----
Private Sub CmbCMPName_LostFocus()
If CmbCMPName.Text = "" Then
   Exit Sub
End If
RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPName = '" & CmbCMPName.Text & "'"
If RsMst_CMP.EOF Then
   MsgBox "Invalid selection "
   CmbCMPName.SetFocus
   Exit Sub
End If
mCMPName = RsMst_CMP!CMPID
TxtPreFix = GetPreFix(RsMst_CMP!CMPID, "P")
End Sub
Private Sub CmdGodownID_Click()
    FrmMst_CMP.Show
End Sub
'----
Private Sub CmbGodownID_GotFocus()
If CmbCMPName.Text = "" Then
   MsgBox "Select CMP first !!! "
   CmbCMPName.SetFocus
   Exit Sub
End If
tmp = CmbGodownID.Text
Call TableMst_Godown("Where CloseDate is null And CMPID = " & mCMPName)
CmbGodownID.Clear
If RsMst_Godown.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Godown.RecordCount
    CmbGodownID.AddItem RsMst_Godown!GodownNo
    RsMst_Godown.MoveNext
Next
CmbGodownID.Text = tmp
End Sub
Private Sub CmbGodownID_LostFocus()
If CmbGodownID.Text = "" Then
   CmbStackID = ""
   TxtLotNo = ""
   Exit Sub
End If

If tmp <> CmbGodownID.Text Then
   CmbStackID = ""
End If

RsMst_Godown.MoveFirst
RsMst_Godown.Find "GodownNo = '" & CmbGodownID.Text & "'"
If RsMst_Godown.EOF Then
   MsgBox "Invalid selection "
   CmbGodownID.SetFocus
   Exit Sub
End If
mGodownID = RsMst_Godown!GodownID
TxtLotNo.Text = GetNextLotNo(mGodownID, TxtStackNo)
End Sub

Private Sub TxtLotNo_LostFocus()
TxtLotNo = Trim(TxtLotNo)
End Sub


Private Sub TxtStackNo_LostFocus()
'If TxtStackNo = "" Then
'   TxtLotNo = ""
'   Exit Sub
'End If
'TxtStackNo = Trim(TxtStackNo)
'TxtLotNo.Text = GetNextLotNo(mGodownID, TxtStackNo)
End Sub


Private Sub Grid_Head1()

MSHFlexGrid3.Clear
MSHFlexGrid3.Rows = 2
MSHFlexGrid3.Cols = 7
MSHFlexGrid3.TextMatrix(0, 0) = "Module Type"
MSHFlexGrid3.TextMatrix(0, 1) = " Module ID"
MSHFlexGrid3.TextMatrix(0, 2) = "No of Bags"
MSHFlexGrid3.TextMatrix(0, 3) = "Demat Weight"
MSHFlexGrid3.TextMatrix(0, 4) = "Txn Type"
MSHFlexGrid3.TextMatrix(0, 5) = "Txn Date"
MSHFlexGrid3.TextMatrix(0, 6) = "GSL ID"

MSHFlexGrid3.RowHeight(0) = 500

MSHFlexGrid3.ColWidth(0) = 1200
MSHFlexGrid3.ColWidth(1) = 1000
MSHFlexGrid3.ColWidth(2) = 1000
MSHFlexGrid3.ColWidth(3) = 1000
MSHFlexGrid3.ColWidth(4) = 1000
MSHFlexGrid3.ColWidth(5) = 1000
MSHFlexGrid3.ColWidth(6) = 800

MSHFlexGrid3.Height = Frame2.Height - 400
MSHFlexGrid3.Width = Frame2.Width - 250

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

Private Function GetNextLotNo(mGodownID_ As Variant, mStackNo_ As Variant) As Variant
Dim Retval As Variant
Retval = ""
If CmbStackID = "" Then Exit Function
If CmbGodownID.Text = "" Then Exit Function
If mGodownID_ > 0 Then
   tmp = "Select LotNo from mst_GSL where Mst_GSL_ID= (Select max(Mst_GSL_ID) from mst_GSL where GodownID=" & mGodownID_ & " And  StackNo='" & mStackNo_ & "') And GodownID=" & mGodownID_ & ""
   Call Tmp3Table
   If TmpRs3.RecordCount > 0 Then
      If IsNumeric(TmpRs3!LotNo) = True Then
         Retval = Val(TmpRs3!LotNo) + 1
      End If
   Else
      Retval = 1
   End If
End If

GetNextLotNo = Retval
End Function

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


Private Sub GenerateCSV()
Dim Path, strFilName, strFilLocation As Variant
Path = App.Path
strFilName = "\excel\" & Gen_mTimeStamp & Gen_UserName & "Mst_GSL.csv"
strFilLocation = Path + strFilName
If Len(Dir(strFilLocation)) > 1 Then Kill strFilLocation
Open strFilLocation For Output As #1

For I = 1 To MSHFlexGrid1.Rows - 1
    tmp = ""
    For C = 0 To MSHFlexGrid1.Cols - 1
        If tmp = "" Then
           tmp = tmp & MSHFlexGrid1.TextMatrix(I, C)
        Else
           tmp = tmp & "," & MSHFlexGrid1.TextMatrix(I, C)
        End If
    Next
    Print #1, tmp
Next
MsgBox "Done !!!"
Close #1
End Sub
