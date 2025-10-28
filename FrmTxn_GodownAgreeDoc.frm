VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmTxn_GodownAgreeDoc 
   Caption         =   "Godown Agreement Document Cycle"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2805
      Left            =   120
      TabIndex        =   44
      Top             =   9840
      Visible         =   0   'False
      Width           =   14895
      Begin VB.CommandButton CmdSearch 
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
         Height          =   350
         Left            =   12480
         TabIndex        =   113
         Top             =   480
         Width           =   1125
      End
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   13680
         TabIndex        =   112
         Top             =   480
         Width           =   1125
      End
      Begin VB.TextBox TxtSAgreeID 
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
         Left            =   7920
         MaxLength       =   12
         TabIndex        =   111
         Top             =   1290
         Width           =   2490
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
         ItemData        =   "FrmTxn_GodownAgreeDoc.frx":0000
         Left            =   3795
         List            =   "FrmTxn_GodownAgreeDoc.frx":0002
         Sorted          =   -1  'True
         TabIndex        =   98
         Top             =   1290
         Width           =   4065
      End
      Begin VB.ComboBox CmbSStatus 
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
         ItemData        =   "FrmTxn_GodownAgreeDoc.frx":0004
         Left            =   7755
         List            =   "FrmTxn_GodownAgreeDoc.frx":000B
         Sorted          =   -1  'True
         TabIndex        =   56
         Top             =   2010
         Visible         =   0   'False
         Width           =   2790
      End
      Begin MSComCtl2.DTPicker STo 
         Height          =   375
         Left            =   1935
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   2715
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
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
         Format          =   111083521
         CurrentDate     =   36494
      End
      Begin MSComCtl2.DTPicker SFrom 
         Height          =   375
         Left            =   375
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   2715
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
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
         Format          =   111083521
         CurrentDate     =   36494
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1455
         Left            =   135
         TabIndex        =   48
         Top             =   1920
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   2566
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
      Begin VB.ComboBox CmbSForwardedTo 
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
         ItemData        =   "FrmTxn_GodownAgreeDoc.frx":0018
         Left            =   3960
         List            =   "FrmTxn_GodownAgreeDoc.frx":001A
         Sorted          =   -1  'True
         TabIndex        =   107
         Top             =   2160
         Visible         =   0   'False
         Width           =   4200
      End
      Begin VB.ComboBox CmbSState 
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
         ItemData        =   "FrmTxn_GodownAgreeDoc.frx":001C
         Left            =   120
         List            =   "FrmTxn_GodownAgreeDoc.frx":001E
         Sorted          =   -1  'True
         TabIndex        =   95
         Top             =   1290
         Width           =   3600
      End
      Begin VB.Frame Frame4 
         Height          =   735
         Left            =   120
         TabIndex        =   74
         Top             =   240
         Width           =   12255
         Begin VB.OptionButton OptA 
            Caption         =   "Agreement Effect Date"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   2040
            TabIndex        =   78
            Top             =   252
            Width           =   2415
         End
         Begin VB.OptionButton OptE 
            Caption         =   "Expiry Date"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   4560
            TabIndex        =   77
            Top             =   252
            Width           =   1455
         End
         Begin VB.OptionButton OptN 
            Caption         =   "None"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   6120
            TabIndex        =   76
            Top             =   252
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton OptS 
            Caption         =   "Agreement Date"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   120
            TabIndex        =   75
            Top             =   252
            Width           =   1815
         End
         Begin MSComCtl2.DTPicker SSTo 
            Height          =   375
            Left            =   10680
            TabIndex        =   79
            TabStop         =   0   'False
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
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
            Format          =   111083521
            CurrentDate     =   36494
         End
         Begin MSComCtl2.DTPicker SSFrom 
            Height          =   375
            Left            =   8280
            TabIndex        =   81
            TabStop         =   0   'False
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
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
            Format          =   111083521
            CurrentDate     =   36494
         End
         Begin VB.Label Label5 
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
            Left            =   7200
            TabIndex        =   82
            Top             =   300
            Width           =   975
         End
         Begin VB.Label Label7 
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
            Left            =   9840
            TabIndex        =   80
            Top             =   300
            Width           =   855
         End
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
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
         Left            =   5445
         TabIndex        =   99
         Top             =   1035
         Width           =   765
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Doc Lifycycle ID"
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
         Left            =   8400
         TabIndex        =   97
         Top             =   1035
         Width           =   1440
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "State"
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
         Left            =   1695
         TabIndex        =   96
         Top             =   1035
         Width           =   465
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
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
         Left            =   8880
         TabIndex        =   57
         Top             =   1725
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "-"
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
         Left            =   13095
         TabIndex        =   51
         Top             =   1335
         Width           =   90
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
         Left            =   10620
         TabIndex        =   50
         Top             =   1320
         Width           =   2415
      End
      Begin VB.Label Label44 
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
         Left            =   720
         TabIndex        =   49
         Top             =   2415
         Visible         =   0   'False
         Width           =   975
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
         Left            =   2235
         TabIndex        =   47
         Top             =   2415
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Forwarded To"
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
         TabIndex        =   108
         Top             =   2040
         Visible         =   0   'False
         Width           =   1275
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1150
      Left            =   120
      TabIndex        =   38
      Top             =   8640
      Width           =   15015
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Print"
         Height          =   815
         Left            =   13005
         TabIndex        =   114
         Top             =   195
         Width           =   1860
      End
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New"
         Height          =   400
         Left            =   105
         TabIndex        =   0
         Top             =   195
         Width           =   2580
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   400
         Left            =   105
         TabIndex        =   18
         Top             =   600
         Width           =   2580
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   400
         Left            =   2685
         TabIndex        =   14
         Top             =   195
         Width           =   2580
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   400
         Left            =   2685
         TabIndex        =   19
         Top             =   600
         Width           =   2580
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   400
         Left            =   5265
         TabIndex        =   15
         Top             =   195
         Width           =   2580
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   400
         Left            =   5265
         TabIndex        =   20
         Top             =   600
         Width           =   2580
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   400
         Left            =   7845
         TabIndex        =   16
         Top             =   195
         Width           =   2580
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   400
         Left            =   7845
         TabIndex        =   21
         Top             =   600
         Width           =   2580
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "Search"
         Height          =   400
         Left            =   10425
         TabIndex        =   17
         Top             =   195
         Width           =   2580
      End
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   400
         Left            =   10425
         TabIndex        =   22
         Top             =   600
         Width           =   2580
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2205
      Left            =   120
      TabIndex        =   23
      Top             =   0
      Width           =   15015
      Begin VB.TextBox TxtUserRole 
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
         Left            =   10245
         Locked          =   -1  'True
         TabIndex        =   91
         TabStop         =   0   'False
         Top             =   2760
         Visible         =   0   'False
         Width           =   4650
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Agreement not Received"
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
         Left            =   6000
         TabIndex        =   61
         Top             =   3173
         Visible         =   0   'False
         Width           =   2640
      End
      Begin VB.TextBox TxtBalance 
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
         Left            =   13680
         Locked          =   -1  'True
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   3300
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox TxtAdvanceRemain 
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
         Left            =   11280
         Locked          =   -1  'True
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   3300
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox TxtAdvAdjusted 
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
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   3300
         Visible         =   0   'False
         Width           =   1455
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
         Height          =   720
         Left            =   1320
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   1320
         Width           =   13575
      End
      Begin VB.TextBox TxtAgreementID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   375
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   270
         Width           =   1785
      End
      Begin MSComCtl2.DTPicker TxtEntryDate 
         Height          =   375
         Left            =   4560
         TabIndex        =   29
         Top             =   300
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   111083521
         CurrentDate     =   39884
      End
      Begin MSComCtl2.DTPicker TxtStartDate 
         Height          =   375
         Left            =   7560
         TabIndex        =   1
         Top             =   300
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
         Format          =   111083521
         CurrentDate     =   39884
      End
      Begin MSComCtl2.DTPicker TxtExpiryDate 
         Height          =   375
         Left            =   13560
         TabIndex        =   3
         Top             =   300
         Width           =   1365
         _ExtentX        =   2408
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
         Format          =   111083521
         CurrentDate     =   39884
      End
      Begin VB.ComboBox CmbAgreRecAs 
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
         ItemData        =   "FrmTxn_GodownAgreeDoc.frx":0020
         Left            =   9000
         List            =   "FrmTxn_GodownAgreeDoc.frx":002A
         Sorted          =   -1  'True
         TabIndex        =   28
         Top             =   3300
         Visible         =   0   'False
         Width           =   3360
      End
      Begin VB.TextBox TxtAgreeRecAs 
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
         Left            =   8640
         Locked          =   -1  'True
         MaxLength       =   200
         ScrollBars      =   2  'Vertical
         TabIndex        =   43
         Top             =   3240
         Visible         =   0   'False
         Width           =   3360
      End
      Begin MSComCtl2.DTPicker TxtAAggrStartDate 
         Height          =   375
         Left            =   11040
         TabIndex        =   2
         Top             =   300
         Width           =   1365
         _ExtentX        =   2408
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
         Format          =   111083521
         CurrentDate     =   39884
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid6 
         Height          =   1200
         Left            =   12360
         TabIndex        =   59
         Top             =   3480
         Visible         =   0   'False
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   2117
         _Version        =   393216
         FixedCols       =   0
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
         ItemData        =   "FrmTxn_GodownAgreeDoc.frx":0044
         Left            =   11640
         List            =   "FrmTxn_GodownAgreeDoc.frx":004B
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   840
         Visible         =   0   'False
         Width           =   3285
      End
      Begin VB.TextBox TxtForwardedTo 
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
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   2760
         Visible         =   0   'False
         Width           =   7800
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
         Left            =   11640
         Locked          =   -1  'True
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   840
         Width           =   3240
      End
      Begin VB.ComboBox CmbForwardedTo 
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
         ItemData        =   "FrmTxn_GodownAgreeDoc.frx":0058
         Left            =   1320
         List            =   "FrmTxn_GodownAgreeDoc.frx":005A
         Sorted          =   -1  'True
         TabIndex        =   8
         Top             =   2760
         Visible         =   0   'False
         Width           =   7800
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
         ItemData        =   "FrmTxn_GodownAgreeDoc.frx":005C
         Left            =   1320
         List            =   "FrmTxn_GodownAgreeDoc.frx":005E
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   840
         Width           =   4680
      End
      Begin VB.ComboBox CmbAgreeType 
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
         ItemData        =   "FrmTxn_GodownAgreeDoc.frx":0060
         Left            =   7920
         List            =   "FrmTxn_GodownAgreeDoc.frx":0062
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   840
         Width           =   3000
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
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   89
         TabStop         =   0   'False
         Top             =   840
         Width           =   4680
      End
      Begin VB.TextBox TxtAgreeType 
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
         Left            =   7920
         Locked          =   -1  'True
         TabIndex        =   93
         TabStop         =   0   'False
         Top             =   840
         Width           =   3000
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "User Role"
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
         Left            =   9240
         TabIndex        =   92
         Top             =   2820
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Agreement Type"
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
         Left            =   6120
         TabIndex        =   90
         Top             =   900
         Width           =   1740
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   120
         TabIndex        =   88
         Top             =   900
         Width           =   900
      End
      Begin VB.Label Label27 
         Caption         =   "Forwarded To"
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
         Left            =   90
         TabIndex        =   60
         Top             =   2700
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Label Label26 
         Caption         =   "Agreement Effect Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9000
         TabIndex        =   58
         Top             =   360
         Width           =   2025
      End
      Begin VB.Label Label15 
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
         Height          =   255
         Left            =   11040
         TabIndex        =   55
         Top             =   900
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label LblBalance 
         AutoSize        =   -1  'True
         Caption         =   "Balance"
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
         Left            =   12840
         TabIndex        =   53
         Top             =   3360
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.Label LblAdvRem 
         AutoSize        =   -1  'True
         Caption         =   "SD Returned"
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
         Left            =   9960
         TabIndex        =   41
         Top             =   3360
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Label LblAgrePecAs 
         AutoSize        =   -1  'True
         Caption         =   "Agreement Received As"
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
         Left            =   6240
         TabIndex        =   39
         Top             =   3360
         Width           =   2535
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Expiry Date"
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
         Left            =   12480
         TabIndex        =   37
         Top             =   360
         Width           =   1035
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Agreement Date"
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
         Left            =   6000
         TabIndex        =   36
         Top             =   360
         Width           =   1470
      End
      Begin VB.Label LblAdvAdjust 
         AutoSize        =   -1  'True
         Caption         =   "SD Adjusted"
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
         Left            =   7080
         TabIndex        =   35
         Top             =   3360
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   120
         TabIndex        =   34
         Top             =   1560
         Width           =   825
      End
      Begin VB.Label LblEntryDate 
         Caption         =   "Doc     Lifecycle Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   3240
         TabIndex        =   33
         Top             =   150
         Width           =   1275
      End
      Begin VB.Label LblAgreemID 
         Caption         =   "Doc Lifecycle ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   135
         TabIndex        =   32
         Top             =   225
         Width           =   1140
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6300
      Left            =   120
      TabIndex        =   42
      Top             =   2280
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   11113
      _Version        =   393216
      TabHeight       =   794
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Document Attachment"
      TabPicture(0)   =   "FrmTxn_GodownAgreeDoc.frx":0064
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "MSHFlexGrid2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Godown Details"
      TabPicture(1)   =   "FrmTxn_GodownAgreeDoc.frx":0080
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(1)=   "MSHFlexGrid3"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Log Details"
      TabPicture(2)   =   "FrmTxn_GodownAgreeDoc.frx":009C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "TxtCreated"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "TxtUpdated"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "MSHFlexGrid5"
      Tab(2).Control(3)=   "Label9"
      Tab(2).Control(4)=   "Label20"
      Tab(2).ControlCount=   5
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
         Left            =   -73800
         Locked          =   -1  'True
         TabIndex        =   84
         TabStop         =   0   'False
         Top             =   720
         Width           =   6210
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
         Left            =   -66360
         Locked          =   -1  'True
         TabIndex        =   83
         TabStop         =   0   'False
         Top             =   720
         Width           =   6210
      End
      Begin VB.Frame Frame2 
         Height          =   3150
         Left            =   120
         TabIndex        =   69
         Top             =   480
         Width           =   14775
         Begin VB.CheckBox Check2 
            Appearance      =   0  'Flat
            Caption         =   "All"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   8925
            TabIndex        =   115
            Top             =   2760
            Width           =   495
         End
         Begin VB.TextBox TxtDocRemarks 
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
            Height          =   725
            Left            =   1440
            MaxLength       =   4000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   11
            Top             =   1875
            Width           =   13215
         End
         Begin VB.ComboBox CmbDocForwardedTo 
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
            ItemData        =   "FrmTxn_GodownAgreeDoc.frx":00B8
            Left            =   4710
            List            =   "FrmTxn_GodownAgreeDoc.frx":00BA
            Sorted          =   -1  'True
            TabIndex        =   13
            Top             =   2700
            Width           =   4080
         End
         Begin VB.ComboBox CmbDocument 
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
            Left            =   1440
            Sorted          =   -1  'True
            TabIndex        =   9
            Top             =   240
            Width           =   13215
         End
         Begin VB.ComboBox CmbDocStatus 
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
            ItemData        =   "FrmTxn_GodownAgreeDoc.frx":00BC
            Left            =   1440
            List            =   "FrmTxn_GodownAgreeDoc.frx":00D2
            Sorted          =   -1  'True
            TabIndex        =   12
            Top             =   2700
            Width           =   2055
         End
         Begin VB.TextBox TxtFileName 
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
            Height          =   420
            Left            =   1440
            Locked          =   -1  'True
            MaxLength       =   4000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   105
            Top             =   720
            Width           =   13215
         End
         Begin VB.CommandButton CmdGetDocAttach 
            Caption         =   "Get Doc Attachement"
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
            Left            =   9600
            Style           =   1  'Graphical
            TabIndex        =   104
            Top             =   2700
            Width           =   2160
         End
         Begin VB.TextBox TxtRemDes 
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
            Height          =   600
            Left            =   1440
            MaxLength       =   4000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   10
            Top             =   1200
            Width           =   13250
         End
         Begin VB.CommandButton CmdDeleteDocList 
            Caption         =   "Delete from List"
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
            Left            =   13080
            Style           =   1  'Graphical
            TabIndex        =   71
            Top             =   2700
            Width           =   1560
         End
         Begin VB.CommandButton CmdAddDocList 
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
            Height          =   375
            Left            =   11815
            Style           =   1  'Graphical
            TabIndex        =   70
            Top             =   2700
            Width           =   1200
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "To view the uploaded document, click on the 'Document  Type' in the table below"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   3360
            TabIndex        =   110
            Top             =   2355
            Width           =   8355
         End
         Begin VB.Label Label30 
            Caption         =   "Forwarded To"
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
            Left            =   3600
            TabIndex        =   109
            Top             =   2640
            Width           =   1035
         End
         Begin VB.Label Label29 
            Caption         =   "File Name"
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
            Left            =   120
            TabIndex        =   106
            Top             =   818
            Width           =   1020
         End
         Begin VB.Label Label25 
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
            Height          =   255
            Left            =   120
            TabIndex        =   103
            Top             =   2753
            Width           =   720
         End
         Begin VB.Label Label16 
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
            Height          =   225
            Left            =   120
            TabIndex        =   102
            Top             =   2125
            Width           =   1020
         End
         Begin VB.Label Label21 
            Caption         =   "Document Description"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   120
            TabIndex        =   101
            Top             =   1268
            Width           =   1260
         End
         Begin VB.Label Label14 
            Caption         =   "Document Type"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Left            =   120
            TabIndex        =   100
            Top             =   128
            Width           =   1620
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1740
         Left            =   -74880
         TabIndex        =   62
         Top             =   570
         Width           =   14715
         Begin VB.ComboBox CmbGodown 
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
            Sorted          =   -1  'True
            TabIndex        =   25
            Top             =   525
            Width           =   2565
         End
         Begin VB.TextBox TxtGodownNo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   375
            Left            =   7155
            Locked          =   -1  'True
            TabIndex        =   67
            Top             =   525
            Width           =   5670
         End
         Begin VB.ComboBox CmbCMP 
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
            Left            =   165
            Sorted          =   -1  'True
            TabIndex        =   24
            Top             =   525
            Width           =   4365
         End
         Begin VB.CommandButton CmdDeleteFrmListGodown 
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
            Left            =   12915
            TabIndex        =   27
            Top             =   1080
            Width           =   1710
         End
         Begin VB.CommandButton CmdAdd2ListGodown 
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
            Left            =   12915
            TabIndex        =   26
            Top             =   600
            Width           =   1710
         End
         Begin VB.TextBox TxtGodownAddress 
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
            Height          =   600
            Left            =   960
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   63
            TabStop         =   0   'False
            Top             =   1000
            Width           =   11895
         End
         Begin VB.Label Label4 
            Caption         =   "Viability ID"
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
            Left            =   5167
            TabIndex        =   94
            Top             =   195
            Width           =   1320
         End
         Begin VB.Label Label3 
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
            Height          =   300
            Left            =   9390
            TabIndex        =   68
            Top             =   195
            Width           =   1200
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "CMP"
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
            Left            =   2130
            TabIndex        =   66
            Top             =   195
            Width           =   495
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Address"
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
            TabIndex        =   64
            Top             =   1180
            Width           =   765
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
         Height          =   3840
         Left            =   -74880
         TabIndex        =   65
         Top             =   2370
         Width           =   14715
         _ExtentX        =   25956
         _ExtentY        =   6773
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
         Height          =   2535
         Left            =   120
         TabIndex        =   72
         Top             =   3645
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   4471
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid5 
         Height          =   4935
         Left            =   -74880
         TabIndex        =   85
         Top             =   1200
         Width           =   14700
         _ExtentX        =   25929
         _ExtentY        =   8705
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
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Created :"
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
         TabIndex        =   87
         Top             =   787
         Width           =   810
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   -67440
         TabIndex        =   86
         Top             =   720
         Width           =   885
      End
   End
   Begin MSComDlg.CommonDialog cdialog 
      Left            =   120
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmTxn_GodownAgreeDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mLocationID, mCMPID, mSCMPID, mLessorID, mGdLessorID, mSMPrefix, mForwardID, Edit_Flg, mDocForwardID As Variant
Dim mGodownID, mStatus, mServiceTaxID, mSStatus, mInspectionID, mType As Variant
Dim mExpiryDate, mSGodownID, mSForwardID, mPrevExpiryDate, mAgreeTypeID, mSStateID, mSLocationID As Variant
Dim File_Path As String
Dim mMsg As Variant
Dim mAmount, mGodownRentAmount As Double
Dim mGodownExists As Boolean
Dim mGodownMsg As Variant
Dim mSelRow, mSelCol As Variant
Dim mDocumentTypeID, strFilename As Variant
Dim Gen_WCGADC As Variant
Const strChecked = "þ"
Const strUnChecked = "q"
Dim mDocStatus As Variant
Dim mCreateId As Variant
Dim FileName As Variant
Dim mStrMode As Variant
Dim mCreatedEmailID As Variant

'Declaration for Auto Resize of Mshflexgrid Row Height
Private Declare Function SendMessage Lib "user32" Alias _
         "SendMessageA" _
() '() '         (ByVal hWnd As Long, _
          ByVal wMsg As Long, _
          ByVal wParam As Long, _
          lParam As Any) As Long

      Const EM_GETLINECOUNT = &HBA
Dim oFs As New FileSystemObject
Dim oFs1 As New FileSystemObject
Dim oFolder As Folder
Dim oFile As File

Private Sub Gen_EmailForError(TOMail As Variant, CCMail As Variant, ErrorMsg As Variant, Mode As Variant)
If ErrorMsg <> "" Then
   mMsg = ""
   mMsg = "<tr>Godown Agreement Document Cycle (from Store-Man) " & Mode & " ID - " & Val(TxtAgreementID.Text) & "</tr>"
   mMsg = mMsg & "<tr><td> Error Description </td><td> " & ErrorMsg & " </td></tr>"
   mMsg = mMsg & "<tr><td> Login By </td><td> " & Gen_UserLoginID & ":" & Gen_UserName & "  </td></tr>"
   mMsg = mMsg & "<tr><td> Entry Date/Time </td><td> " & Now & "  </td></tr>"

   mMsgSubject = "Error Found in Godown Agreement Document Cycle ID - " & Val(TxtAgreementID.Text) & ""
   TOADD = TOMail
   semail = CCMail
   Call SendMail1(TOADD, semail, mMsgSubject, mMsg)
End If
End Sub
Public Function SendMail1(ToADD_ As Variant, semail_ As Variant, MsgSubject_ As Variant, Message1_ As Variant)

Dim htmlStringhead As String
Dim htmlStringtail As String
Set oSMTP = New OSSMTP.SMTPSession

'servertxt = "10.0.0.10"
servertxt = "mail.ncml.com"


htmlStringhead = "<html><body><font size= 3 face= Zurich BT><table border = 1>"
htmlStringtail = "</table></Font></body></html>"

'With oSMTP
'  'connection
'  .Server = servertxt
'  'message
'  .MailFrom = Gen_UserEmailID  '"mayur.d@ncml.com"
'  If ToADD_ = "" Or IsNull(ToADD_) = True Then
'     ToADD_ = "parag.r@ncml.com"
'  End If
'  .SendTo = ToADD_
'  TOCC_ = semail_
'  If ToADD_ <> TOCC_ Then
'     .CC = TOCC_
'  End If
''  .BCC = "mayur.d@ncml.com"
'  .MessageSubject = MsgSubject_
'  .MessageHTML = htmlStringhead & Message1_ & htmlStringtail
'  .TimeStamp = Now()
'  If LCase(Gen_EmailConfigFlag) = "y" Then
'     .SendEmail
'  End If
'End With

Dim objSMTP
Set objSMTP = CreateObject("MailBee.SMTP")
objSMTP.LicenseKey = "MBC900-5F470FAD67-D8E202A2E22927335F24F5A667BFD196"
objSMTP.ServerName = Gen_ReadMailServerIPINI("MailIP") ' = "mail.ncml.com" '"mail.ncml.com"
objSMTP.FromAddr = Gen_UserEmailID
If IsNull(Trim(ToADD_)) = True Or Trim(ToADD_) = "" Then
   ToADD_ = "manu.ag@ncml.com"
End If
objSMTP.ToAddr = ToADD_
TOCC_ = semail_
If ToADD_ <> TOCC_ Then
   objSMTP.CCAddr = TOCC_
End If

objSMTP.Subject = MsgSubject_ '& " (from Store-Man)" 'MsgSubject_ '"Message in HTML format"
objSMTP.BodyFormat = 1
objSMTP.BodyText = htmlStringhead & Message1_ & htmlStringtail 'htmlStringhead & MsgSubject_ & " :- " & Message1_ & htmlStringtail '"<html>This word is <b>bold</b></html>"
If LCase(Gen_EmailConfigFlag) = "y" Then
   objSMTP.Send
End If

End Function
Private Sub AddCmd_Click()
TxtDocRemarks.Height = 725
Check2.Value = 0

Edit_Flg = "A"
Call GButtonLock(Me, False)
Frame3.Enabled = True
Frame1.Enabled = True
Frame2.Enabled = True
SSTab1.Tab = 0

TxtAgreementID = ""
TxtEntryDate.Value = Date
TxtStartDate.Value = Date
TxtAAggrStartDate.Value = Date
TxtExpiryDate.Value = Date
TxtForwardedTo.Text = ""
TxtLocationID.Text = ""
TxtAgreeType.Text = ""
TxtStatus.Text = ""
TxtUserRole.Text = ""
TxtRemarks.Text = ""
CmbForwardedTo.Text = ""

CmbForwardedTo.Visible = True
TxtForwardedTo.Visible = False

CmbLocationID.Visible = True
TxtLocationID.Visible = False

CmbAgreeType.Visible = True
TxtAgreeType.Visible = False

CmbStatus.Visible = False
TxtStatus.Visible = False
TxtStatus.Text = "Pending"
CmbStatus.Text = TxtStatus.Text

TxtCreated = ""
TxtUpdated = ""
Call ClearGodownFrame
Call ClearDocumentFrame

Call Grid_Head1
Call Grid_Head2
Call Grid_Head3
Call Grid_Head5
End Sub

Private Sub Grid_Head1()
With MSHFlexGrid1
 .Clear
 .Rows = 2
 .FixedRows = 1
 .Cols = 9 '10
 .WordWrap = True
 .TextMatrix(0, 0) = "Doc Lifecycle ID"
 .TextMatrix(0, 1) = "Doc Lifecycle Date"
 .TextMatrix(0, 2) = "Agreement Date"
 .TextMatrix(0, 3) = "Agreement Effect Date"
 .TextMatrix(0, 4) = "Expiry Date"
 .TextMatrix(0, 5) = "Forwarded To"
 .TextMatrix(0, 6) = "State"
 .TextMatrix(0, 7) = "Location"
 '.TextMatrix(0, 8) = "Status"
 .TextMatrix(0, 8) = "Remarks"

 .ColWidth(0) = 1200
 .ColWidth(1) = 1200
 .ColWidth(2) = 1200
 .ColWidth(3) = 1200
 .ColWidth(4) = 1200
 .ColWidth(5) = 0
 .ColWidth(6) = 2000
 .ColWidth(7) = 1500
 '.ColWidth(8) = 2000
 .ColWidth(8) = 4000

 .RowHeight(0) = 800
End With
End Sub

Private Sub Grid_Head3()
With MSHFlexGrid3
 .Clear
 .Rows = 3
 .FixedRows = 2
 .Cols = 9

 .TextMatrix(1, 0) = "RowID"
 .TextMatrix(1, 1) = "Location"
 .TextMatrix(1, 2) = "CMP Name"
 .TextMatrix(1, 3) = "Godown Viability"
 .TextMatrix(1, 4) = "Godown No"
 .TextMatrix(1, 5) = "SM_Prefix"
 .TextMatrix(1, 6) = "Address"

 .TextMatrix(1, 7) = "LocationID"
 .TextMatrix(1, 8) = "CMP ID"

 .ColWidth(0) = 0
 .ColWidth(1) = 0
 .ColWidth(2) = 2500
 .ColWidth(3) = 2000
 .ColWidth(4) = 2000
 .ColWidth(5) = 0
 .ColWidth(6) = 6000
 .ColWidth(7) = 0
 .ColWidth(8) = 0

End With
End Sub

Private Sub Check2_Click()
'If Check2.Value = vbChecked Then
'   For i = 1 To MSHFlexGrid2.Rows - 2
'       MSHFlexGrid2.Row = i
'       MSHFlexGrid2.Col = 6
'       MSHFlexGrid2.CellFontName = "Wingdings"
'       MSHFlexGrid2.CellFontSize = 12
'       MSHFlexGrid2.CellAlignment = flexAlignCenterCenter
'       MSHFlexGrid2.Text = strChecked
'   Next
'Else
'   For i = 1 To MSHFlexGrid2.Rows - 2
'       MSHFlexGrid2.Row = i
'       MSHFlexGrid2.Col = 6
'       MSHFlexGrid2.CellFontName = "Wingdings"
'       MSHFlexGrid2.CellFontSize = 12
'       MSHFlexGrid2.CellAlignment = flexAlignCenterCenter
'       MSHFlexGrid2.Text = strUnChecked
'   Next
'End If

If LCase(CmbDocStatus.Text) <> "approved" Or CmbDocForwardedTo.Text = "" Then
   Check2.Value = 0
   Exit Sub
End If
If Check2.Value = 1 Then
   For I = 1 To MSHFlexGrid2.Rows - 2
       If LCase(MSHFlexGrid2.TextMatrix(I, 7)) = "g" Or LCase(MSHFlexGrid2.TextMatrix(I, 7)) = "l" Then
          MSHFlexGrid2.TextMatrix(I, 6) = CmbDocStatus.Text
          MSHFlexGrid2.TextMatrix(I, 8) = CmbDocForwardedTo.Text
          MSHFlexGrid2.TextMatrix(I, 9) = mDocForwardID
       End If
   Next
Else
      tmp = " Select TGAD.DocumentTypeID,TGAD.DocDID,TGAD.Flag,ME.EmployeeName+'~'+ME.EmployeeNo as EmpName,ME.EmployeeID,ME.EmployeeEmailID"
      tmp = tmp & " from Txn_GodownAgreeDocCycle TGAC"
      tmp = tmp & " Inner Join Txn_GADCDocumentAttachDetail TGAD on TGAC.AgreementID=TGAD.AgreementID"
      tmp = tmp & " Inner Join Mst_Employee ME on TGAD.ForwardID=ME.EmployeeID"
      tmp = tmp & " Where TGAC.AgreementID = " & Val(TxtAgreementID.Text) & ""
      Call TmpTable
      If TmpRs.RecordCount > 0 Then
         For I = 1 To MSHFlexGrid2.Rows - 2
             TmpRs.MoveFirst
             TmpRs.Find "DocDID=" & Val(MSHFlexGrid2.TextMatrix(I, 5)) & ""
             If TmpRs.EOF = False Then
                If LCase(TmpRs!Flag) = "g" Then
                   MSHFlexGrid2.TextMatrix(I, 6) = "Forwarded to CMG"
                ElseIf LCase(TmpRs!Flag) = "p" Then
                   MSHFlexGrid2.TextMatrix(I, 6) = "Forwarded to CMPM"
                ElseIf LCase(TmpRs!Flag) = "f" Then
                   MSHFlexGrid2.TextMatrix(I, 6) = "Forwarded to FAG"
                ElseIf LCase(TmpRs!Flag) = "l" Then
                   MSHFlexGrid2.TextMatrix(I, 6) = "Forwarded to Legal"
                ElseIf LCase(TmpRs!Flag) = "a" Then
                   MSHFlexGrid2.TextMatrix(I, 6) = "Approved"
                ElseIf LCase(TmpRs!Flag) = "m" Then
                   MSHFlexGrid2.TextMatrix(I, 6) = "Forwarded to AM"
                End If
                MSHFlexGrid2.TextMatrix(I, 8) = TmpRs!EmpName
                MSHFlexGrid2.TextMatrix(I, 9) = TmpRs!EmployeeID
             End If
         Next
   End If
End If

End Sub

Private Sub CmbAgreeType_GotFocus()
tmp = CmbAgreeType.Text
CmbCMP.Clear
CmbGodown.Clear
Call TableMst_UnitType("Where Status = 'A'")
CmbAgreeType.Clear
If RsMst_UnitType.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   CmbAgreeType.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_UnitType.RecordCount
    CmbAgreeType.AddItem RsMst_UnitType!UnitType
    RsMst_UnitType.MoveNext
Next
CmbAgreeType.Text = tmp
End Sub

Private Sub CmbAgreeType_LostFocus()
If CmbAgreeType.Text = "" Then
   Exit Sub
End If
RsMst_UnitType.MoveFirst
RsMst_UnitType.Find "UnitType = '" & CmbAgreeType.Text & "'"
If RsMst_UnitType.EOF Then
   MsgBox "Invalid selection "
   CmbAgreeType.SetFocus
   Exit Sub
End If
mAgreeTypeID = RsMst_UnitType!UnitTypeID
End Sub

Private Sub CmbAgreRecAs_LostFocus()
If CmbAgreRecAs.Text = "" Then Exit Sub
   If CmbAgreRecAs.Text <> "Hard Copy" And CmbAgreRecAs.Text <> "Soft Copy" Then
      MsgBox "Invalid selection "
      CmbAgreRecAs.SetFocus
      Exit Sub
   End If
End Sub
Private Sub CmbDocForwardedTo_GotFocus()
Dim mUserRole

If CmbDocStatus.Text = "" Then Exit Sub

tmp1 = CmbDocForwardedTo.Text
mUserRole = ""

CmbDocForwardedTo.Clear

If LCase(CmbDocStatus.Text) = "forwarded to cmg" Then
   mUserRole = "'superadmin','admin','headcmg'"
   tmp = " Select ME.EmployeeID,ME.EmployeeName,ME.EmployeeNo,MUR.UserRoleName"
   tmp = tmp & " from Mst_Employee ME"
   'Parag
   'tmp = tmp & " Inner Join Mst_USers MU on ME.EmployeeNo=MU.Employee_ID"
   tmp = tmp & " Inner Join Mst_USers MU on ME.EmployeeNo=MU.EmployeeNo"
   'tmp = tmp & " Inner Join Mst_UserRole MUR on MU.UserGroupID=MUR.UserGroupID"
   tmp = tmp & " Inner Join Mst_UserRole MUR on MU.UserRoleID=MUR.UserRoleID"
   tmp = tmp & " Where ME.Flag = 'Y' And MUR.UserRoleName in (" & mUserRole & ")"
ElseIf LCase(CmbDocStatus.Text) = "forwarded to cmpm" Then
'   mUserRole = "'superadmin','admin','head cmg'"
'   tmp = " Select ME.EmployeeID,ME.EmployeeName,ME.EmployeeNo,MUR.GroupName"
'   tmp = tmp & " from Mst_Employee ME"
'   tmp = tmp & " Inner Join Mst_USers MU on ME.EmployeeNo=MU.Employee_ID"
'   tmp = tmp & " Inner Join Mst_UserRole MUR on MU.UserGroupID=MUR.UserGroupID"
'   tmp = tmp & " Where ME.Flag = 'Y' And MUR.GroupName in (" & mUserRole & ")"
'   tmp = tmp & " Union All"

   'Call UpperHierarchy(mCreateId)
   tmp = " Select ME.EmployeeID,ME.EmployeeName,ME.EmployeeNo"
   tmp = tmp & " from Mst_Employee ME Where ME.EmployeeID=" & mCreateId & ""

ElseIf LCase(CmbDocStatus.Text) = "forwarded to legal" Then
   mUserRole = "'legal'"
   tmp = " Select ME.EmployeeID,ME.EmployeeName,ME.EmployeeNo,MUR.UserRoleName"
   tmp = tmp & " from Mst_Employee ME"
   'Parag
   'tmp = tmp & " Inner Join Mst_USers MU on ME.EmployeeNo=MU.Employee_ID"
   'tmp = tmp & " Inner Join Mst_UserRole MUR on MU.UserGroupID=MUR.UserGroupID"
   
   tmp = tmp & " Inner Join Mst_USers MU on ME.EmployeeNo=MU.EmployeeNo"
   tmp = tmp & " Inner Join Mst_UserRole MUR on MU.UserRoleID=MUR.UserRoleID"
   
   
   tmp = tmp & " Where ME.Flag = 'Y' And MUR.UserRoleName in (" & mUserRole & ")"
ElseIf LCase(CmbDocStatus.Text) = "approved" Then
   tmp = " Select ME.EmployeeID,ME.EmployeeName,ME.EmployeeNo"
   tmp = tmp & " from Mst_Employee ME Where ME.EmployeeID=" & Gen_EmployeeID & ""
ElseIf LCase(CmbDocStatus.Text) = "forwarded to am" Then
'   mUserRole = "'amsandrms'"
'   tmp = " Select ME.EmployeeID,ME.EmployeeName,ME.EmployeeNo,MUR.GroupName"
'   tmp = tmp & " from Mst_Employee ME"
'   tmp = tmp & " Inner Join Mst_USers MU on ME.EmployeeNo=MU.Employee_ID"
'   tmp = tmp & " Inner Join Mst_UserRole MUR on MU.UserGroupID=MUR.UserGroupID"
'   tmp = tmp & " Where ME.Flag = 'Y' And MUR.GroupName in (" & mUserRole & ")"
   If Edit_Flg = "A" Then
      tmp = " Select ME1.EmployeeID,ME1.EmployeeName,ME1.EmployeeNo"
      tmp = tmp & " from Mst_Employee ME"
      tmp = tmp & " Inner Join Mst_Employee ME1 on ME.DirectReportingTo=ME1.EmployeeID"
      tmp = tmp & " Where ME.EmployeeNo='" & Gen_UserLoginID & "'"
   Else
      tmp = " Select ME1.EmployeeID,ME1.EmployeeName,ME1.EmployeeNo"
      tmp = tmp & " from Mst_Employee ME"
      tmp = tmp & " Inner Join Mst_Employee ME1 on ME.DirectReportingTo=ME1.EmployeeID"
      tmp = tmp & " Where ME.EmployeeID=" & mCreateId & ""
   End If

'ElseIf LCase(CmbDocStatus.Text) = "rejected" Then
'   tmp = " Select ME.EmployeeID,ME.EmployeeName,ME.EmployeeNo"
'   tmp = tmp & " from Mst_Employee ME Where ME.EmployeeID=" & Gen_EmployeeID & ""
End If

'If LCase(Gen_UserRole) <> "superadmin" And LCase(Gen_UserRole) <> "admin" And LCase(Gen_UserRole) <> "head cmg" And LCase(Gen_UserRole) <> "fag" And LCase(Gen_UserRole) <> "fagadmin" And LCase(Gen_UserRole) <> "legal" Then
'   mUserRole = "'superadmin','admin','head cmg'"
'
'   tmp = " Select ME.EmployeeID,ME.EmployeeName,ME.EmployeeNo,MUR.GroupName"
'   tmp = tmp & " from Mst_Employee ME"
'   tmp = tmp & " Inner Join Mst_USers MU on ME.EmployeeNo=MU.Employee_ID"
'   tmp = tmp & " Inner Join Mst_UserRole MUR on MU.UserGroupID=MUR.UserGroupID"
'   tmp = tmp & " Where ME.Flag = 'Y' And MUR.GroupName in (" & mUserRole & ")"
'Else
'   If LCase(Gen_UserRole) = "superadmin" Or LCase(Gen_UserRole) = "admin" Then
'      mUserRole = "'superadmin','admin','head cmg','fag','fagadmin','power'"
'
'      tmp = " Select ME.EmployeeID,ME.EmployeeName,ME.EmployeeNo,MUR.GroupName"
'      tmp = tmp & " from Mst_Employee ME"
'      tmp = tmp & " Inner Join Mst_USers MU on ME.EmployeeNo=MU.Employee_ID"
'      tmp = tmp & " Inner Join Mst_UserRole MUR on MU.UserGroupID=MUR.UserGroupID"
'      tmp = tmp & " Where ME.Flag = 'Y' And MUR.GroupName Not in (" & mUserRole & ")"
'   ElseIf LCase(Gen_UserRole) = "head cmg" Then
'      mUserRole = "'legal','fag','fag admin'"
'      tmp = " Select ME.EmployeeID,ME.EmployeeName,ME.EmployeeNo,MUR.GroupName"
'      tmp = tmp & " from Mst_Employee ME"
'      tmp = tmp & " Inner Join Mst_USers MU on ME.EmployeeNo=MU.Employee_ID"
'      tmp = tmp & " Inner Join Mst_UserRole MUR on MU.UserGroupID=MUR.UserGroupID"
'      tmp = tmp & " Where ME.Flag = 'Y' And MUR.GroupName in (" & mUserRole & ")"
'   ElseIf LCase(Gen_UserRole) = "legal" Or LCase(Gen_UserRole) = "fag" Or LCase(Gen_UserRole) = "fag admin" Then
'
'      tmp = " Select ME.EmployeeID,ME.EmployeeName,ME.EmployeeNo,MUR.GroupName"
'      tmp = tmp & " from Mst_Employee ME"
'      tmp = tmp & " Inner Join Mst_USers MU on ME.EmployeeNo=MU.Employee_ID"
'      tmp = tmp & " Inner Join Mst_UserRole MUR on MU.UserGroupID=MUR.UserGroupID"
'      tmp = tmp & " Where ME.Flag = 'Y' "
'   Else
'      tmp = " Select ME.EmployeeID,ME.EmployeeName,ME.EmployeeNo,MUR.GroupName"
'      tmp = tmp & " from Mst_Employee ME"
'      tmp = tmp & " Inner Join Mst_USers MU on ME.EmployeeNo=MU.Employee_ID"
'      tmp = tmp & " Inner Join Mst_UserRole MUR on MU.UserGroupID=MUR.UserGroupID"
'      tmp = tmp & " Where ME.Flag = 'Y' "
'   End If
'End If
Call Tmp2Table
If TmpRs2.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   CmbDocForwardedTo.SetFocus
   Exit Sub
End If
For I = 1 To TmpRs2.RecordCount
    CmbDocForwardedTo.AddItem TmpRs2!EmployeeName & "~" & TmpRs2!EmployeeNo
    TmpRs2.MoveNext
Next
CmbDocForwardedTo.Text = tmp1
End Sub

Private Sub UpperHierarchy(mEmployeeID_ As Variant)
tmp = "Select EmployeeID,EmployeeName + '~' + EmployeeNo As EmployeeName,DirectReportingTo From Mst_Employee Where EmployeeID=" & mEmployeeID_ & " And Flag='Y' And DesignationID<>1"
Call TmpTable
If TmpRs.RecordCount > 0 Then
   If Not IsNull(TmpRs!DirectReportingTo) Then
      CmbDocForwardedTo.AddItem TmpRs!EmployeeName
      Call UpperHierarchy(TmpRs!DirectReportingTo)
   End If
End If
End Sub
Private Sub CmbDocForwardedTo_LostFocus()
If CmbDocForwardedTo.Text = "" Then Exit Sub

Dim strSplitString() As String
Dim EmpNo_, EmpName_ As String

If InStr(1, CmbDocForwardedTo, "~", vbTextCompare) = 0 Then
   MsgBox "Invalid selection "
   CmbDocForwardedTo.SetFocus
   Exit Sub
End If

strSplitString = Split(CmbDocForwardedTo, "~")
EmpName_ = Trim(strSplitString(0))
EmpNo_ = Trim(strSplitString(1))

If CmbDocForwardedTo.Text = "" Then
   mForwardID = ""
   Exit Sub
End If

tmp = "Select EmployeeID From Mst_Employee Where EmployeeNo='" & EmpNo_ & "'"
Call Tmp1Table
If TmpRs1.EOF Then
   MsgBox "Invalid selection "
   CmbDocForwardedTo.SetFocus
   Exit Sub
End If

mDocForwardID = TmpRs1!EmployeeID
End Sub

Private Sub CmbDocStatus_GotFocus()
CmbDocStatus.Clear
CmbDocForwardedTo.Clear
If LCase(Gen_UserRole) <> "superadmin" And LCase(Gen_UserRole) <> "admin" And LCase(Gen_UserRole) <> "headcmg" And LCase(Gen_UserRole) <> "fag" And LCase(Gen_UserRole) <> "fagadmin" And LCase(Gen_UserRole) <> "legal" And LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "amsandrms" Then
   'CmbDocStatus.AddItem "Forwarded to CMG"
    CmbDocStatus.AddItem "Forwarded to AM"
Else
   If LCase(Gen_UserRole) = "superadmin" Or LCase(Gen_UserRole) = "admin" Then
      If Edit_Flg = "A" Then
         CmbDocStatus.AddItem "Forwarded to CMG"
      Else
         CmbDocStatus.AddItem "Forwarded to CMG"
'         CmbDocStatus.AddItem "Forwarded To CMPM"
         CmbDocStatus.AddItem "Forwarded to AM"
         CmbDocStatus.AddItem "Forwarded To Legal"
         If LCase(Gen_UserRole) = "superadmin" Then
            CmbDocStatus.AddItem "Approved"
         End If
      End If
   ElseIf LCase(Gen_UserRole) = "headcmg" Then
      If Edit_Flg = "A" Then
         CmbDocStatus.AddItem "Forwarded to CMG"
      Else
         CmbDocStatus.AddItem "Forwarded to CMG"
'         CmbDocStatus.AddItem "Forwarded To CMPM"
         CmbDocStatus.AddItem "Forwarded to AM"
         CmbDocStatus.AddItem "Forwarded To Legal"
         CmbDocStatus.AddItem "Approved"
      End If
   ElseIf LCase(Gen_UserRole) = "legal" Or LCase(Gen_UserRole) = "power" Then
      If Edit_Flg = "A" Then
         CmbDocStatus.AddItem "Forwarded to CMG"
      Else
         CmbDocStatus.AddItem "Forwarded to CMG"
'         CmbDocStatus.AddItem "Forwarded To CMPM"
         CmbDocStatus.AddItem "Forwarded to AM"
         CmbDocStatus.AddItem "Forwarded To Legal"
         CmbDocStatus.AddItem "Approved"
      End If
   ElseIf LCase(Gen_UserRole) = "amsandrms" Then
      If Edit_Flg = "A" Then
         CmbDocStatus.AddItem "Forwarded to CMG"
      Else
         CmbDocStatus.AddItem "Forwarded to CMG"
         CmbDocStatus.AddItem "Forwarded To CMPM"
      End If
   End If
End If
End Sub

Private Sub CmbDocStatus_LostFocus()
Dim iCtr As Variant
Dim Flag As Boolean
Flag = False

If CmbDocStatus.Text = "" Then Exit Sub


For iCtr = 0 To CmbDocStatus.ListCount - 1
    If LCase(CmbDocStatus.Text) = LCase(CmbDocStatus.List(iCtr)) Then
       Flag = True
    End If
Next
If Flag = False Then
   MsgBox "Invalid selection "
   CmbDocStatus.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmbDocument_GotFocus()
tmp = CmbDocument.Text
Call TableMst_DocumentType(" Where Flag='A'")
CmbDocument.Clear
If RsMst_DocumentType.RecordCount = 0 Then
   MsgBox "No Document Type Found!!!"
   CmbDocument.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_DocumentType.RecordCount
    CmbDocument.AddItem RsMst_DocumentType!DocumentType
    RsMst_DocumentType.MoveNext
Next
CmbDocument.Text = tmp
End Sub

Private Sub CmbDocument_LostFocus()
If CmbDocument.Text = "" Then
   Exit Sub
End If

RsMst_DocumentType.MoveFirst
RsMst_DocumentType.Find "DocumentType = '" & CmbDocument.Text & "'"
If RsMst_DocumentType.EOF Then
   MsgBox "Invalid selection "
   CmbDocument.SetFocus
   Exit Sub
End If
mDocumentTypeID = RsMst_DocumentType!DocumentTypeID
End Sub
Private Sub ReSizeCellHeight(MyRow As Long, MyCol As Long)
Dim LinesOfText As Long
Dim HeightOfLine As Long

'Set MSFlexGrid to appropriate Cell
MSHFlexGrid2.row = MyRow
MSHFlexGrid2.Col = MyCol

'Set textbox width to match current width of selected cell
TxtForRowHeight.Width = MSHFlexGrid2.ColWidth(MyCol)

'Set font info of textbox to match FlexGrid control
TxtForRowHeight.Font.Name = MSHFlexGrid2.Font.Name
TxtForRowHeight.Font.Size = MSHFlexGrid2.Font.Size
TxtForRowHeight.Font.Bold = MSHFlexGrid2.Font.Bold
TxtForRowHeight.Font.Italic = MSHFlexGrid2.Font.Italic
TxtForRowHeight.Font.Strikethrough = MSHFlexGrid2.Font.Strikethrough
TxtForRowHeight.Font.Underline = MSHFlexGrid2.Font.Underline

'Set font info of form to match FlexGrid control
Me.Font.Name = MSHFlexGrid2.Font.Name
Me.Font.Size = MSHFlexGrid2.Font.Size
Me.Font.Bold = MSHFlexGrid2.Font.Bold
Me.Font.Italic = MSHFlexGrid2.Font.Italic
Me.Font.Strikethrough = MSHFlexGrid2.Font.Strikethrough
Me.Font.Underline = MSHFlexGrid2.Font.Underline

'Put the text from the selected cell into the textbox
TxtForRowHeight.Text = MSHFlexGrid2.Text
TxtForRowHeight.Text = Replace(TxtForRowHeight.Text, vbNewLine, "")

'Get the height of the text in the textbox
HeightOfLine = Me.TextHeight(TxtForRowHeight.Text)

'Call API to determine how many lines of text are in text box
LinesOfText = SendMessage(TxtForRowHeight.hWnd, EM_GETLINECOUNT, 0&, 0&)

'Check to see if row is not tall enough
If Val(MyCol) = 1 Then
If MSHFlexGrid2.RowHeight(MyRow) < (LinesOfText * HeightOfLine) Then
  'Adjust the RowHeight based on the number of lines in textbox
   MSHFlexGrid2.RowHeight(MyRow) = LinesOfText * HeightOfLine
End If
End If
End Sub

Private Sub CmbForwardedTo_GotFocus()
Dim mUserRole
tmp1 = CmbForwardedTo.Text
CmbForwardedTo.Clear
mUserRole = ""

If LCase(Gen_UserRole) <> "superadmin" And LCase(Gen_UserRole) <> "admin" And LCase(Gen_UserRole) <> "headcmg" And LCase(Gen_UserRole) <> "fag" And LCase(Gen_UserRole) <> "fagadmin" And LCase(Gen_UserRole) <> "legal" Then
   mUserRole = "'superadmin','admin','headcmg'"

   tmp = " Select ME.EmployeeID,ME.EmployeeName,ME.EmployeeNo,MUR.GroupName"
   tmp = tmp & " from Mst_Employee ME"
   tmp = tmp & " Inner Join Mst_USers MU on ME.EmployeeNo=MU.EmployeeNo"
   tmp = tmp & " Inner Join Mst_UserRole MUR on MU.UserGroupID=MUR.UserGroupID"
   tmp = tmp & " Where ME.Flag = 'Y' And MUR.GroupName in (" & mUserRole & ")"
Else
   If LCase(Gen_UserRole) = "superadmin" Or LCase(Gen_UserRole) = "admin" Then
      mUserRole = "'superadmin','admin','headcmg','fag','fagadmin','power'"

      tmp = " Select ME.EmployeeID,ME.EmployeeName,ME.EmployeeNo,MUR.GroupName"
      tmp = tmp & " from Mst_Employee ME"
      tmp = tmp & " Inner Join Mst_USers MU on ME.EmployeeNo=MU.EmployeeNo"
      tmp = tmp & " Inner Join Mst_UserRole MUR on MU.UserGroupID=MUR.UserGroupID"
      tmp = tmp & " Where ME.Flag = 'Y' And MUR.GroupName Not in (" & mUserRole & ")"

   ElseIf LCase(Gen_UserRole) = "headcmg" Then
      mUserRole = "'legal','fag','fagadmin'"
      tmp = " Select ME.EmployeeID,ME.EmployeeName,ME.EmployeeNo,MUR.GroupName"
      tmp = tmp & " from Mst_Employee ME"
      tmp = tmp & " Inner Join Mst_USers MU on ME.EmployeeNo=MU.EmployeeNo"
      tmp = tmp & " Inner Join Mst_UserRole MUR on MU.UserGroupID=MUR.UserGroupID"
      tmp = tmp & " Where ME.Flag = 'Y' And MUR.GroupName in (" & mUserRole & ")"

   ElseIf LCase(Gen_UserRole) = "legal" Or LCase(Gen_UserRole) = "fag" Or LCase(Gen_UserRole) = "fagadmin" Then

      tmp = " Select ME.EmployeeID,ME.EmployeeName,ME.EmployeeNo,MUR.GroupName"
      tmp = tmp & " from Mst_Employee ME"
      tmp = tmp & " Inner Join Mst_USers MU on ME.EmployeeNo=MU.EmployeeNo"
      tmp = tmp & " Inner Join Mst_UserRole MUR on MU.UserGroupID=MUR.UserGroupID"
      tmp = tmp & " Where ME.Flag = 'Y' "
   Else
      tmp = " Select ME.EmployeeID,ME.EmployeeName,ME.EmployeeNo,MUR.GroupName"
      tmp = tmp & " from Mst_Employee ME"
      tmp = tmp & " Inner Join Mst_USers MU on ME.EmployeeNo=MU.EmployeeNo"
      tmp = tmp & " Inner Join Mst_UserRole MUR on MU.UserGroupID=MUR.UserGroupID"
      tmp = tmp & " Where ME.Flag = 'Y' "
   End If
End If

Call Tmp2Table
CmbForwardedTo.Clear
If TmpRs2.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   CmbForwardedTo.SetFocus
   Exit Sub
End If
For I = 1 To TmpRs2.RecordCount
    CmbForwardedTo.AddItem TmpRs2!EmployeeName & "~" & TmpRs2!EmployeeNo
    TmpRs2.MoveNext
Next
CmbForwardedTo.Text = tmp1
End Sub

Private Sub CmbForwardedTo_LostFocus()
If CmbForwardedTo.Text = "" Then Exit Sub

Dim strSplitString() As String
Dim EmpNo_, EmpName_ As String

strSplitString = Split(CmbForwardedTo, "~")
EmpName_ = Trim(strSplitString(0))
EmpNo_ = Trim(strSplitString(1))

If CmbForwardedTo.Text = "" Then
   mForwardID = ""
   Exit Sub
End If

TmpRs2.MoveFirst
TmpRs2.Find "EmployeeName = '" & EmpName_ & "'"
If TmpRs2.EOF Then
   MsgBox "Invalid selection "
   CmbForwardedTo.SetFocus
   Exit Sub
End If
mForwardID = TmpRs2!EmployeeID
TxtUserRole.Text = TmpRs2!GroupName
End Sub

Private Sub CmbLocationID_GotFocus()
tmp = CmbLocationID.Text
CmbCMP.Clear
CmbGodown.Clear
If Gen_WcLocation = "" Then
   Call TableMst_Location("Where LocationID <> 0 ")
Else
   Call TableMst_Location("where " & Gen_WcLocation & " And LocationID <> 0 ")
End If
CmbLocationID.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   CmbLocationID.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbLocationID.AddItem RsMst_Location!Locationname
    RsMst_Location.MoveNext
Next
CmbLocationID.Text = tmp
End Sub

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
mSMPrefix = RsMst_Location!SM_Prefix
If tmp <> CmbLocationID.Text Then
   Call Grid_Head3
End If
End Sub

Private Sub CmbSAgreRecAs_LostFocus()
If CmbSAgreRecAs.Text = "" Then Exit Sub

If CmbSAgreRecAs.Text <> "Hard Copy" And CmbSAgreRecAs.Text <> "Soft Copy" Then
   MsgBox "Invalid selection "
   CmbSAgreRecAs.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmbCMP_GotFocus()

If mLocationID = "" Then
   MsgBox "Please Select Location !!!"
   If CmbLocationID.Enabled = True And CmbLocationID.Visible = True Then CmbLocationID.SetFocus
   Exit Sub
End If

CmbGodown.Clear
tmp = CmbCMP.Text
Call TableMst_CMP(" Where LocationID = " & mLocationID & "")
CmbCMP.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   CmbCMP.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_CMP.RecordCount
    CmbCMP.AddItem RsMst_CMP!CMPName
    RsMst_CMP.MoveNext
Next
CmbCMP.Text = tmp
End Sub

Private Sub CmbCMP_LostFocus()
If CmbCMP.Text = "" Then
   Exit Sub
End If
RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPName = '" & CmbCMP.Text & "'"
If RsMst_CMP.EOF Then
   MsgBox "Invalid selection "
   CmbCMP.SetFocus
   Exit Sub
End If
mCMPID = RsMst_CMP!CMPID
End Sub

Private Sub CmbGodown_GotFocus()

If CmbCMP.Text = "" Then
   MsgBox "Select CMP first !!! "
   CmbCMP.SetFocus
   Exit Sub
End If

If mAgreeTypeID = "" Then
   MsgBox "Select Agreement Type first !!! "
   CmbAgreeType.SetFocus
   Exit Sub
End If

tmp = CmbGodown.Text
'Parag
'Call TableMst_GodownInspection("Where CMPID = " & mCMPID & " And Status ='A' And GodownTypeID=" & mAgreeTypeID & " And InspectionID Not In (Select isnull(InspectionID,0) from Txn_GADCGodownDetail where AgreementID <> " & Val(TxtAgreementID.Text) & ")")
Call TableMst_GodownInspection("Where CMPID = " & mCMPID & " And Status ='A' And UnitTypeID=" & mAgreeTypeID & " And InspectionID Not In (Select isnull(InspectionID,0) from Txn_GADCGodownDetail where AgreementID <> " & Val(TxtAgreementID.Text) & ")")
'UnitTypeID
CmbGodown.Clear
If RsMst_GodownInspection.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
'   CmbGodown.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_GodownInspection.RecordCount
    CmbGodown.AddItem RsMst_GodownInspection!InspectionID
    RsMst_GodownInspection.MoveNext
Next
CmbGodown.Text = tmp
End Sub

Private Sub CmbGodown_LostFocus()
If CmbGodown.Text = "" Then
   Exit Sub
End If

If RsMst_GodownInspection.RecordCount = 0 Then
   CmbGodown.SetFocus
   Exit Sub
End If

RsMst_GodownInspection.MoveFirst
RsMst_GodownInspection.Find "InspectionID =" & CmbGodown.Text & ""
If RsMst_GodownInspection.EOF Then
   MsgBox "Invalid selection "
   CmbGodown.SetFocus
   Exit Sub
End If
mInspectionID = RsMst_GodownInspection!InspectionID
mType = RsMst_GodownInspection!Type
TxtGodownAddress.Text = IIf(IsNull(RsMst_GodownInspection!Address), "", RsMst_GodownInspection!Address)
If LCase(mType) = "r" Then
   mGodownID = RsMst_GodownInspection!GodownID
   tmp = " Select GodownNo from Mst_Godown where GodownID=" & mGodownID & ""
   Call Tmp3Table
   If TmpRs3.RecordCount > 0 Then
      TxtGodownNo.Text = TmpRs3!GodownNo
   End If
ElseIf LCase(mType) = "n" Then
   TxtGodownNo.Text = RsMst_GodownInspection!GodownNo
End If
End Sub

Private Sub ClearGodownFrame()
CmbCMP.Text = ""
CmbGodown.Text = ""
TxtGodownAddress.Text = ""
TxtGodownNo.Text = ""
CmdAdd2ListGodown.Caption = "Add to List"
End Sub
Private Sub ClearDocumentFrame()
CmbDocument.Text = ""
CmbDocStatus.Text = ""
CmbDocForwardedTo.Text = ""
TxtFileName.Text = ""
TxtRemDes.Text = ""
TxtDocRemarks.Text = ""
Check2.Value = 0

'If LCase(Gen_UserRole) <> "superadmin" And LCase(Gen_UserRole) <> "admin" And LCase(Gen_UserRole) <> "head cmg" And LCase(Gen_UserRole) <> "fag" And LCase(Gen_UserRole) <> "fagadmin" And LCase(Gen_UserRole) <> "legal" And LCase(Gen_UserRole) <> "power" Then
'   CmbDocStatus.Clear
'   CmbDocStatus.AddItem "Forwarded to CMG"
'Else
'   If LCase(Gen_UserRole) = "superadmin" Or LCase(Gen_UserRole) = "admin" Then
'      CmbDocStatus.Clear
'      CmbDocStatus.AddItem "Forwarded to CMG"
'      CmbDocStatus.AddItem "Forwarded To CMPM"
'      CmbDocStatus.AddItem "Forwarded To Legal"
'      CmbDocStatus.AddItem "Rejected"
'   ElseIf LCase(Gen_UserRole) = "head cmg" Then
'      CmbDocStatus.Clear
'      CmbDocStatus.AddItem "Forwarded to CMG"
'      CmbDocStatus.AddItem "Forwarded To CMPM"
'      CmbDocStatus.AddItem "Forwarded To Legal"
'      CmbDocStatus.AddItem "Forwarded To FAG"
'      CmbDocStatus.AddItem "Rejected"
'   ElseIf LCase(Gen_UserRole) = "legal" Then
'      CmbDocStatus.Clear
'      CmbDocStatus.AddItem "Forwarded to CMG"
'      CmbDocStatus.AddItem "Forwarded To CMPM"
'      CmbDocStatus.AddItem "Forwarded To Legal"
'      CmbDocStatus.AddItem "Forwarded To FAG"
'      CmbDocStatus.AddItem "Approved"
'      CmbDocStatus.AddItem "Rejected"
'   ElseIf LCase(Gen_UserRole) = "fag" Or LCase(Gen_UserRole) = "fagadmin" Or LCase(Gen_UserRole) = "power" Then
'      CmbDocStatus.Clear
'      CmbDocStatus.AddItem "Forwarded to CMG"
'      CmbDocStatus.AddItem "Forwarded To CMPM"
'      CmbDocStatus.AddItem "Forwarded To Legal"
'      CmbDocStatus.AddItem "Forwarded To FAG"
'      CmbDocStatus.AddItem "Approved"
'   End If
'End If

CmdAddDocList.Caption = "Add to List"
End Sub
Private Sub CmbSForwardedTo_GotFocus()
tmp1 = CmbSForwardedTo.Text
CmbSForwardedTo.Clear
tmp = " Select ME.EmployeeID,ME.EmployeeName,ME.EmployeeNo,MUR.GroupName"
tmp = tmp & " from Mst_Employee ME"
tmp = tmp & " Inner Join Mst_USers MU on ME.EmployeeNo=MU.EmployeeNo"
tmp = tmp & " Inner Join Mst_UserRole MUR on MU.UserGroupID=MUR.UserGroupID"
tmp = tmp & " Where ME.Flag = 'Y'"
Call Tmp2Table
CmbSForwardedTo.Clear
If TmpRs2.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   CmbSForwardedTo.SetFocus
   Exit Sub
End If
For I = 1 To TmpRs2.RecordCount
    CmbSForwardedTo.AddItem TmpRs2!EmployeeName & "~" & TmpRs2!EmployeeNo
    TmpRs2.MoveNext
Next
CmbSForwardedTo.Text = tmp1
End Sub


Private Sub CmbSForwardedTo_LostFocus()
If CmbSForwardedTo.Text = "" Then Exit Sub

Dim strSplitString() As String
Dim EmpNo_, EmpName_ As String

strSplitString = Split(CmbSForwardedTo, "~")
EmpName_ = Trim(strSplitString(0))
EmpNo_ = Trim(strSplitString(1))

If CmbSForwardedTo.Text = "" Then
   mForwardID = ""
   Exit Sub
End If

TmpRs2.MoveFirst
TmpRs2.Find "EmployeeName = '" & EmpName_ & "'"
If TmpRs2.EOF Then
   MsgBox "Invalid selection "
   CmbSForwardedTo.SetFocus
   Exit Sub
End If
mSForwardID = TmpRs2!EmployeeID
End Sub

Private Sub CmbSLocationID_GotFocus()
If CmbSState.Text = "" Then Exit Sub
tmp = CmbSLocationID.Text

If Gen_WcLocation = "" Then
   Call TableMst_Location("Where LocationID <> 0 And StateID=" & mSStateID & "")
Else
   Call TableMst_Location("where " & Gen_WcLocation & " And LocationID <> 0 And StateID=" & mSStateID & "")
End If

CmbSLocationID.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   CmbSLocationID.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbSLocationID.AddItem RsMst_Location!Locationname
    RsMst_Location.MoveNext
Next
CmbSLocationID.Text = tmp
End Sub

Private Sub CmbSLocationID_LostFocus()
If CmbSState.Text = "" Then Exit Sub

If CmbSLocationID.Text = "" Then
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbSLocationID.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection "
   CmbSLocationID.SetFocus
   Exit Sub
End If
mSLocationID = RsMst_Location!LocationID
End Sub

Private Sub CmbSState_GotFocus()
tmp = CmbSState.Text
If Gen_WcState <> "" Then
   Call TableMst_State("Where" & Gen_WcState)
Else
   Call TableMst_State
End If
CmbSState.Clear
If RsMst_State.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   CmbSState.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_State.RecordCount
    CmbSState.AddItem RsMst_State!StateName
    RsMst_State.MoveNext
Next
CmbSState.Text = tmp
End Sub

Private Sub CmbSState_LostFocus()
If CmbSState.Text = "" Then
   Exit Sub
End If
RsMst_State.MoveFirst
RsMst_State.Find "StateName = '" & CmbSState.Text & "'"
If RsMst_State.EOF Then
   MsgBox "Invalid selection "
   CmbSState.SetFocus
   Exit Sub
End If
mSStateID = RsMst_State!StateID
End Sub

Private Sub CmbSStatus_LostFocus()
If CmbSStatus.Text = "" Then
   Exit Sub
End If

'If LCase(CmbSStatus.Text) <> "forwarded to cmg" And LCase(CmbSStatus.Text) <> "forwarded to cmpm" And LCase(CmbSStatus.Text) <> "forwarded to fag" And LCase(CmbSStatus.Text) <> "forwarded to legal" And LCase(CmbSStatus.Text) <> "approved" And LCase(CmbSStatus.Text) <> "rejected" Then
'   MsgBox "Invalid selection "
'   CmbSStatus.SetFocus
'   Exit Sub
'End If
If LCase(CmbSStatus.Text) <> "pending" Then
   MsgBox "Invalid selection "
   CmbSStatus.SetFocus
   Exit Sub
End If
If LCase(CmbSStatus.Text) = "forwarded to cmg" Then
   mSStatus = "G"
ElseIf LCase(CmbSStatus.Text) = "pending" Then
   mSStatus = "P"
ElseIf LCase(CmbSStatus.Text) = "forwarded to fag" Then
   mSStatus = "F"
ElseIf LCase(CmbSStatus.Text) = "forwarded to legal" Then
   mSStatus = "L"
ElseIf LCase(CmbSStatus.Text) = "approved" Then
   mSStatus = "A"
'ElseIf LCase(CmbSStatus.Text) = "rejected" Then
'   mSStatus = "R"
End If

End Sub

Private Sub CmbStatus_LostFocus()
If CmbStatus.Text = "" Then
   Exit Sub
End If
If LCase(CmbStatus.Text) <> "pending" Then
   MsgBox "Invalid selection "
   CmbStatus.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmdAdd2ListGodown_Click()

If CmbCMP.Text = "" Then
   MsgBox "Select CMP Name"
   CmbCMP.SetFocus
   Exit Sub
End If

If CmbGodown.Text = "" Then
   MsgBox "Select Viability ID"
   CmbGodown.SetFocus
   Exit Sub
End If

If CheckDuplicateGodown = True Then
   MsgBox "Duplicate Godown Viability Not Allowed."
   CmbGodown.SetFocus
   Exit Sub
End If

If CmdAdd2ListGodown.Caption = "Update List" Then
   r = MSHFlexGrid3.RowSel
Else
   r = MSHFlexGrid3.Rows - 1
   MSHFlexGrid3.Rows = MSHFlexGrid3.Rows + 1
   MSHFlexGrid3.TextMatrix(r, 0) = ""
End If

MSHFlexGrid3.TextMatrix(r, 1) = IIf(Edit_Flg = "A", CmbLocationID.Text, TxtLocationID.Text)
MSHFlexGrid3.TextMatrix(r, 2) = CmbCMP.Text
MSHFlexGrid3.TextMatrix(r, 3) = CmbGodown.Text
MSHFlexGrid3.TextMatrix(r, 4) = TxtGodownNo.Text
MSHFlexGrid3.TextMatrix(r, 5) = mSMPrefix
MSHFlexGrid3.TextMatrix(r, 6) = TxtGodownAddress.Text
MSHFlexGrid3.TextMatrix(r, 7) = mLocationID
MSHFlexGrid3.TextMatrix(r, 8) = mCMPID

Call ClearGodownFrame
End Sub
Private Function CheckGodownExpiry(mGodownID_ As Variant, Flag_ As Variant) As Boolean

Dim Retval As Boolean
Retval = True

tmp = "Select TGL.LAID,TGL.GodownID,Max(TL.ExpiryDate) as ExpiryDate  from Txn_GodownLessorAgreeDetail TGL"
tmp = tmp & " inner join Txn_LessorAgreement TL on TGL.LAID=TL.LAID"
tmp = tmp & " Where TGL.GodownID = " & mGodownID_ & " ANd TL.FLag <> 'R' "
tmp = tmp & " Group by TGL.LAID,TGL.GodownID"
tmp = tmp & " order by TGL.LAID"

Call TmpTable

If TmpRs.RecordCount > 0 Then
   For I = 1 To TmpRs.RecordCount
       If Flag_ = "E" Then
          If CDate(TmpRs!ExpiryDate) >= Date Then
             If TmpRs!LAID <> Val(TxtAgreementID) Then
                mExpiryDate = TmpRs!ExpiryDate
                Retval = False
                Exit For

             End If
          End If
       ElseIf Flag_ = "S" Then
          If CDate(TmpRs!ExpiryDate) >= TxtStartDate.Value Then
             If TmpRs!LAID <> Val(TxtAgreementID) Then
                mExpiryDate = TmpRs!ExpiryDate
                Retval = False
                Exit For
             End If
          End If
       Else
          If TmpRs!LAID <> Val(TxtAgreementID) Then
             If CDate(TmpRs!ExpiryDate) > TxtExpiryDate.Value Then   ' for Check Expiry Date of New Agreement
                mExpiryDate = TmpRs!ExpiryDate
                Retval = False
                Exit For
             End If
          End If
       End If
       TmpRs.MoveNext
   Next
End If
CheckGodownExpiry = Retval

End Function

Function CheckDuplicateLessorName() As Boolean
Dim Retval As Boolean
Retval = False
If MSHFlexGrid2.Rows > 1 Then
   For u = 1 To MSHFlexGrid2.Rows - 1
       If MSHFlexGrid2.TextMatrix(u, 1) = CmbLessorName.Text Then
          If LCase(CmdAdd2ListLessor.Caption) = "add to list" Then
             Retval = True
             Exit For
          Else
             If u <> MSHFlexGrid2.RowSel Then
                Retval = True
             End If
          End If
       End If
   Next
End If

CheckDuplicateLessorName = Retval
End Function

Function CheckDuplicateGodown() As Boolean
Dim Retval As Boolean
If mCMPID = "" Then Exit Function
Retval = False
If MSHFlexGrid3.Rows > 1 Then
   For u = 1 To MSHFlexGrid3.Rows - 1
       If MSHFlexGrid3.TextMatrix(u, 3) = mInspectionID And MSHFlexGrid3.TextMatrix(u, 8) = mCMPID Then
          If LCase(CmdAdd2ListGodown.Caption) = "add to list" Then
             Retval = True
             Exit For
          Else
             If u <> MSHFlexGrid3.RowSel Then
                Retval = True
             End If
          End If
       End If
   Next
End If
CheckDuplicateGodown = Retval
End Function

Function CheckDuplicateDocument() As Boolean
Dim Retval As Boolean
Retval = False
If MSHFlexGrid2.Rows > 1 Then
   For u = 1 To MSHFlexGrid2.Rows - 1
       If MSHFlexGrid2.TextMatrix(u, 4) = mDocumentTypeID Then
          If LCase(CmdAddDocList.Caption) = "add to list" Then
             Retval = True
             Exit For
          Else
             If u <> MSHFlexGrid2.RowSel Then
                Retval = True
             End If
          End If
       End If
   Next
End If
CheckDuplicateDocument = Retval
End Function

Function CheckDuplicateDocumentAttach() As Boolean
Dim Retval As Boolean
Retval = False
If MSHFlexGrid2.Rows > 1 Then
   For u = 1 To MSHFlexGrid2.Rows - 1
       For v = 1 To MSHFlexGrid2.Rows - 1
           If v <> u Then
              If MSHFlexGrid2.TextMatrix(u, 1) = MSHFlexGrid2.TextMatrix(v, 1) Then
                 Retval = True
                 CheckDuplicateDocumentAttach = Retval
                 Exit Function
              End If
           End If
       Next
   Next
End If
CheckDuplicateDocumentAttach = Retval
End Function
Private Sub CmdAddDocList_Click()

If CmbDocument.Text = "" Then
   MsgBox "Select Document Type"
   CmbDocument.SetFocus
   Exit Sub
End If

If TxtFileName.Text = "" Then
   MsgBox "Blank File Name Not Allowed"
   Exit Sub
End If

If TxtRemDes.Text = "" Then
   MsgBox "Blank Document Description Not Allowed"
   TxtRemDes.SetFocus
   Exit Sub
End If

If CmbDocStatus.Text = "" Then
   MsgBox "Select Status"
   CmbDocStatus.SetFocus
   Exit Sub
End If

If CmbDocForwardedTo.Text = "" Then
   MsgBox "Blank Forward To Not Allowed"
   CmbDocForwardedTo.SetFocus
   Exit Sub
End If

If CheckDuplicateDocument = True Then
   MsgBox "Duplicate Document Type Not Allowed."
   CmbDocument.SetFocus
   Exit Sub
End If

If CmdAddDocList.Caption = "Update List" Then
   r = MSHFlexGrid2.RowSel
Else
   r = MSHFlexGrid2.Rows - 1
   MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
   MSHFlexGrid2.TextMatrix(r, 5) = ""
End If

MSHFlexGrid2.TextMatrix(r, 0) = CmbDocument.Text
MSHFlexGrid2.TextMatrix(r, 1) = TxtFileName.Text
MSHFlexGrid2.TextMatrix(r, 2) = TxtRemDes.Text
MSHFlexGrid2.TextMatrix(r, 3) = TxtDocRemarks.Text
MSHFlexGrid2.TextMatrix(r, 4) = mDocumentTypeID
MSHFlexGrid2.TextMatrix(r, 6) = CmbDocStatus.Text
MSHFlexGrid2.TextMatrix(r, 8) = CmbDocForwardedTo.Text
MSHFlexGrid2.TextMatrix(r, 9) = mDocForwardID

Call ClearDocumentFrame

End Sub

Private Sub CmdDeleteDocList_Click()
Dim iCtr
iCtr = 1
I = MSHFlexGrid2.RowSel
If I > 0 Then
   If MSHFlexGrid2.TextMatrix(I, 0) = "" Then Exit Sub
      If MSHFlexGrid2.TextMatrix(I, 0) <> "" Then
         For RO = I To MSHFlexGrid2.Rows - 2
             For C = 0 To MSHFlexGrid2.Cols - 1
                 MSHFlexGrid2.TextMatrix(RO, C) = MSHFlexGrid2.TextMatrix(RO + 1, C)
                 MSHFlexGrid2.row = RO
             Next
         Next
         MSHFlexGrid2.Rows = MSHFlexGrid2.Rows - 1
      End If
End If
Call ClearDocumentFrame
End Sub

Private Sub CmdDeleteFrmListGodown_Click()

If CmbCMP.Text = "" Or CmbGodown.Text = "" Then Exit Sub
I = MSHFlexGrid3.RowSel
If I > 0 Then
   If MSHFlexGrid3.TextMatrix(I, 1) <> "" Then
      For RO = I To MSHFlexGrid3.Rows - 2
          For C = 0 To MSHFlexGrid3.Cols - 1
              MSHFlexGrid3.TextMatrix(RO, C) = MSHFlexGrid3.TextMatrix(RO + 1, C)
          Next
      Next
      MSHFlexGrid3.Rows = MSHFlexGrid3.Rows - 1
   End If
End If
Call ClearGodownFrame
End Sub


Private Function CheckLessorInGodown(LessorID_ As Variant) As Boolean
Dim Retval As Boolean
Retval = False
For r = 2 To MSHFlexGrid1.Rows - 2
    If Val(MSHFlexGrid1.TextMatrix(r, 6)) = Val(LessorID_) Then
       Retval = True
       Exit For
    End If
Next
CheckLessorInGodown = Retval
End Function
Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "GodownAgreementDocumentCycle.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "GodownAgreementDocumentCycle.xls")
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
Public Sub OnEditCreateAttachFile()
On Error GoTo Edit_Err

If Right(App.Path, 1) = "\" Then
   File_Path = App.Path & "Agreement Attachment\Godown Agreements\" & Gen_UserLoginID
Else
   File_Path = App.Path & "\Agreement Attachment\Godown Agreements\" & Gen_UserLoginID
End If
If oFs.FolderExists(File_Path) = True Then
   Set oFolder = oFs.GetFolder(File_Path)
       For Each oFile In oFolder.Files
           Call oFs.DeleteFile(oFile.Path, True)
       Next
Else
   If Not oFs.FolderExists(File_Path) Then
      Call oFs.CreateFolder(File_Path)
   End If
End If


tmp = " Select MDT.DocumentType,TGAD.FileName,TGAD.[File],TGAD.[Desc],TGAD.Remarks,TGAD.DocumentTypeID,TGAD.DocDID,TGAD.Flag,ME.EmployeeName+'~'+ME.EmployeeNo as EmpName,ME.EmployeeID,ME.EmployeeEmailID"
tmp = tmp & " from Txn_GodownAgreeDocCycle TGAC"
tmp = tmp & " Inner Join Txn_GADCDocumentAttachDetail TGAD on TGAC.AgreementID=TGAD.AgreementID"
tmp = tmp & " Inner Join Mst_DocumentType MDT on TGAD.DocumentTypeID=MDT.DocumentTypeID"
tmp = tmp & " Inner Join Mst_Employee ME on TGAD.ForwardID=ME.EmployeeID"
tmp = tmp & " Where TGAC.AgreementID = " & Val(TxtAgreementID.Text) & " "
Call Tmp2Table
If TmpRs2.RecordCount > 0 Then
   For I = 1 To TmpRs2.RecordCount
      
       If Right(App.Path, 1) = "\" Then
          File_Path = App.Path & "Agreement Attachment\Godown Agreements\" & Gen_UserLoginID & "\" & TmpRs2!FileName
       Else
          File_Path = App.Path & "\Agreement Attachment\Godown Agreements\" & Gen_UserLoginID & "\" & TmpRs2!FileName
       End If

       Call LoadFileFromDB(File_Path, TmpRs2, "File")
       
       File_Path = App.Path & "\Agreement Attachment\Godown Agreements\" & Gen_UserLoginID

   TmpRs2.MoveNext
   Next
End If

Edit_Err:
If Err.Number > 0 Then MsgBox Err.Description
End Sub

Private Sub CmdGetDocAttach_Click()
If CmbDocument.Text = "" Then Exit Sub

With cdialog
     .InitDir = App.Path
     .DialogTitle = "Select File"
     .FilterIndex = 1
'     .Filter = "Pdf files(*.pdf)|*.pdf"
     .ShowOpen
End With
If cdialog.CancelError Then
   TxtFileName.Text = ""
Else
   strFilename = cdialog.FileName
   If strFilename = "" Then Exit Sub
   TxtFileName.Text = strFilename
End If
End Sub

Private Sub CmdPrint_Click()

If TxtAgreementID.Text = "" Then
   Exit Sub
End If

Gen_mTimeStamp = GetTimeStamp

Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Godown Agreement Document Cycle " & Trim(TxtAgreementID) & ".xls")
FileName = Gen_mTimeStamp & Gen_UserName & "Godown Agreement Document Cycle " & Trim(TxtAgreementID) & ".xls"
Call Xls_File
End Sub

Private Sub Gen_Create_PAXls(FileName_ As String)

Dim oExcel As Object
Dim oBook As Object
Dim oSheet As Object
Dim Pat As String
Pat = App.Path
Set oExcel = CreateObject("Excel.Application")
Set oBook = oExcel.Workbooks.Add
Set oSheet = oBook.Worksheets(1)

tmp = "Select Password from mst_users where EmployeeNo  = '" & Gen_UserLoginID & "'"
Call TmpTable
If TmpRs.RecordCount > 0 Then
   mPassword = IIf(IsNull(TmpRs!Password), Gen_ExcelPassword, TmpRs!Password)
Else
   mPassword = Gen_ExcelPassword
End If

mPassword = mEmployeeNo & "_" & mPassword
oBook.SetPasswordEncryptionOptions "Microsoft Enhanced Cryptographic Provider v1.0", "RC4", 128, True
Pat = App.Path
oBook.SaveAs Pat & "\excel\" & FileName_, , mPassword
oExcel.Quit
End Sub
Private Sub CmdSearch_Click()

Dim Wc1, mSerTax, mStatus As Variant
Call Grid_Head1

If SSFrom.Value > SSTo.Value Then
   MsgBox "From Date Must be less than To Date!!!"
   SFrom.SetFocus
   Exit Sub
End If

tmp = " Select TGAC.AgreementID,TGAC.EntryDate,TGAC.AgreementActualStartDate,TGAC.StartDate,TGAC.ExpiryDate, "
tmp = tmp & " MS.StateName,ML.LocationName,TGAC.Remarks " 'TGAC.Status,
tmp = tmp & " from Txn_GodownAgreeDocCycle TGAC "
tmp = tmp & " Inner Join Mst_Location ML on TGAC.LocationID=ML.LocationID "
tmp = tmp & " Inner Join Mst_State MS on ML.StateID=MS.StateID "
'tmp = tmp & " Inner Join Mst_Employee ME on TGAC.ForwardEmpolyeeID=ME.EmployeeID"

'If Gen_DBType = "MDB" Then
'   wc1 = " Where TGAC.EntryDate Between #" & Format(SFrom.Value, Gen_DatFormat) & "# And #" & Format(STo.Value, Gen_DatFormat) & "# "
'Else
'   wc1 = " Where TGAC.EntryDate Between '" & Format(SFrom.Value, Gen_DatFormat) & "' And '" & Format(STo.Value, Gen_DatFormat) & "' "
'End If

If OptS.Value = True Then
   If Gen_DBType = "MDB" Then
      Wc1 = " Where TGAC.StartDate Between #" & Format(SSFrom.Value, Gen_DatFormat) & "# And #" & Format(SSTo.Value, Gen_DatFormat) & "# "
   Else
      Wc1 = " Where TGAC.StartDate Between '" & Format(SSFrom.Value, Gen_DatFormat) & "' And '" & Format(SSTo.Value, Gen_DatFormat) & "' "
   End If
ElseIf OptA.Value = True Then
   If Gen_DBType = "MDB" Then
      Wc1 = " Where TGAC.AgreementActualStartDate Between #" & Format(SSFrom.Value, Gen_DatFormat) & "# And #" & Format(SSTo.Value, Gen_DatFormat) & "# "
   Else
      Wc1 = " Where TGAC.AgreementActualStartDate Between '" & Format(SSFrom.Value, Gen_DatFormat) & "' And '" & Format(SSTo.Value, Gen_DatFormat) & "' "
   End If
ElseIf OptE.Value = True Then
   If Gen_DBType = "MDB" Then
      Wc1 = " Where TGAC.ExpiryDate Between #" & Format(SSFrom.Value, Gen_DatFormat) & "# And #" & Format(SSTo.Value, Gen_DatFormat) & "# "
   Else
      Wc1 = " Where TGAC.ExpiryDate Between '" & Format(SSFrom.Value, Gen_DatFormat) & "' And '" & Format(SSTo.Value, Gen_DatFormat) & "' "
   End If
Else

End If

If CmbSForwardedTo.Text <> "" Then
   If Wc1 <> "" Then
      Wc1 = Wc1 & " And TGAC.ForwardEmpolyeeID=" & mSForwardID & ""
   Else
      Wc1 = " Where TGAC.ForwardEmpolyeeID=" & mSForwardID & ""
   End If
End If

If TxtSAgreeID.Text <> "" Then
   If Wc1 <> "" Then
      Wc1 = Wc1 & " And TGAC.AgreementID=" & TxtSAgreeID.Text & ""
   Else
      Wc1 = " Where TGAC.AgreementID=" & TxtSAgreeID.Text & ""
   End If
End If

'If CmbSStatus.Text <> "" Then
'   If Wc1 <> "" Then
'      Wc1 = Wc1 & " And TGAC.Status='" & mSStatus & "'"
'   Else
'      Wc1 = " Where TGAC.Status='" & mSStatus & "'"
'   End If
'End If

If CmbSState.Text <> "" Then
   If Wc1 <> "" Then
      Wc1 = Wc1 & " And MS.StateName='" & CmbSState & "'"
   Else
      Wc1 = Wc1 & " Where MS.StateName='" & CmbSState & "'"
   End If
End If

If CmbSLocationID.Text <> "" Then
   If Wc1 <> "" Then
      Wc1 = Wc1 & " And ML.LocationName='" & CmbSLocationID & "'"
   Else
      Wc1 = Wc1 & " Where ML.LocationName='" & CmbSLocationID & "'"
   End If
End If

If Gen_WcLocation <> "" Then
   If Wc1 <> "" Then
      Wc1 = Wc1 & " And ML." & Gen_WcLocation & ""
   Else
      Wc1 = Wc1 & " Where ML." & Gen_WcLocation & ""
   End If
End If

tmp = tmp & Wc1 & " Order By TGAC.AgreementID"

Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   MSHFlexGrid1.Rows = TmpRs1.RecordCount + 2
   For I = 1 To TmpRs1.RecordCount
       MSHFlexGrid1.TextMatrix(I, 0) = IIf(IsNull(TmpRs1!AgreementID), "", TmpRs1!AgreementID)
       MSHFlexGrid1.TextMatrix(I, 1) = IIf(IsNull(TmpRs1!EntryDate), Date, Format(TmpRs1!EntryDate, "dd-MMM-yyyy"))
       MSHFlexGrid1.TextMatrix(I, 2) = IIf(IsNull(TmpRs1!AgreementActualStartDate), Date, Format(TmpRs1!AgreementActualStartDate, "dd-MMM-yyyy"))
       MSHFlexGrid1.TextMatrix(I, 3) = IIf(IsNull(TmpRs1!StartDate), Date, Format(TmpRs1!StartDate, "dd-MMM-yyyy"))
       MSHFlexGrid1.TextMatrix(I, 4) = IIf(IsNull(TmpRs1!ExpiryDate), Date, Format(TmpRs1!ExpiryDate, "dd-MMM-yyyy"))
       'MSHFlexGrid1.TextMatrix(i, 5) = IIf(IsNull(TmpRs1!ForwardedName), "", TmpRs1!ForwardedName)
       MSHFlexGrid1.TextMatrix(I, 6) = IIf(IsNull(TmpRs1!StateName), "", TmpRs1!StateName)
       MSHFlexGrid1.TextMatrix(I, 7) = IIf(IsNull(TmpRs1!Locationname), "", TmpRs1!Locationname)
       'MSHFlexGrid1.TextMatrix(i, 8) = IIf(IsNull(TmpRs1!Status), "", TmpRs1!Status)

       'If LCase(TmpRs1!Status) = "g" Then
       '   MSHFlexGrid1.TextMatrix(i, 8) = "Forwarded to CMG"
       'ElseIf LCase(TmpRs1!Status) = "p" Then
       '   MSHFlexGrid1.TextMatrix(i, 8) = "Forwarded to CMPM"
       'ElseIf LCase(TmpRs1!Status) = "f" Then
       '   MSHFlexGrid1.TextMatrix(i, 8) = "Forwarded to CMG"
       'ElseIf LCase(TmpRs1!Status) = "l" Then
       '   MSHFlexGrid1.TextMatrix(i, 8) = "Forwarded to Legal"
       'ElseIf LCase(TmpRs1!Status) = "a" Then
       '   MSHFlexGrid1.TextMatrix(i, 8) = "Approved"
'      ' ElseIf LCase(TmpRs1!Status) = "r" Then
'      '    MSHFlexGrid1.TextMatrix(i, 8) = "Rejected"
       'End If

       MSHFlexGrid1.TextMatrix(I, 8) = IIf(IsNull(TmpRs1!Remarks), "", TmpRs1!Remarks)
       TmpRs1.MoveNext
   Next
End If
Label22.Caption = TmpRs1.RecordCount
Label22.Refresh
End Sub

Private Sub DeleteCmd_Click()
'Dim Ans As Variant
'On Error GoTo msgError
'If LCase(TxtStatus) = "approved" Or LCase(TxtStatus) = "rejected" Then
'   MsgBox "This agreement is " & TxtStatus & ". You can't delete this agreement!!!"
'   Exit Sub
'End If
'Ans = MsgBox("Do you really want to DELETE this record???", vbYesNo)
'If Ans = vbYes Then
'
'   Call DeleteChildTableDetail
'
'   RsTxn_GodownAgreeDocCycle.Delete
'   RsTxn_GodownAgreeDocCycle.Update
'   If RsTxn_GodownAgreeDocCycle.RecordCount = 0 Then
'      Call AddCmd_Click
'      Exit Sub
'   End If
'   RsTxn_GodownAgreeDocCycle.MoveNext
'   If RsTxn_GodownAgreeDocCycle.EOF() Then
'      RsTxn_GodownAgreeDocCycle.MoveLast
'   End If
'   Call InData
'End If
'Exit Sub
'msgError:
'MsgBox "Child record Present "
'RsTxn_GodownAgreeDocCycle.CancelUpdate
End Sub
Sub DeleteChildTableDetail()
tmp = "Delete from Txn_GADCDocumentAttachDetail where AgreementID = " & Val(TxtAgreementID.Text) & ""
Call TmpTable
tmp = "Delete from Txn_GADCGodownDetail where AgreementID = " & Val(TxtAgreementID.Text) & ""
Call TmpTable
tmp = "Delete from Txn_GADCAgreementAttachDetail where AgreementID = " & Val(TxtAgreementID.Text) & ""
Call TmpTable
End Sub

Private Sub EditCmd_Click()
Edit_Flg = "E"
TxtDocRemarks.Height = 725

'If LCase(mStatus) = "a" Or LCase(mStatus) = "r" Then
'   MsgBox "Approved or Rejected Godown Aggreement Can't Edit"
'   Exit Sub
'End If

If MSHFlexGrid3.Rows > 3 And MSHFlexGrid2.Rows > 2 Then
   Call TableTxn_GADCDocumentAttachDetail("Where AgreementID=" & TxtAgreementID.Text & " And Not Flag in ('A','R')")
   If RsTxn_GADCDocumentAttachDetail.RecordCount = 0 Then
      ''--- 06 Sep 2017 ---
      If LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "superadmin" And LCase(Gen_UserRole) <> "admin" And LCase(Gen_UserRole) <> "headcmg" Then
        MsgBox "Can't Edit, All Documents Approved."
        Exit Sub
      End If
      ''-------------------
   End If
End If

CmbDocStatus.Clear
If LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "superadmin" And LCase(Gen_UserRole) <> "admin" And LCase(Gen_UserRole) <> "headcmg" And LCase(Gen_UserRole) <> "fag" And LCase(Gen_UserRole) <> "fagadmin" And LCase(Gen_UserRole) <> "legal" And LCase(Gen_UserRole) <> "amsandrms" Then
'   CmbDocStatus.AddItem "Forwarded To CMG"
   CmbDocStatus.AddItem "Forwarded To AM"
ElseIf LCase(Gen_UserRole) = "superadmin" Or LCase(Gen_UserRole) = "admin" Then
   CmbDocStatus.AddItem "Forwarded To CMG"
'   CmbDocStatus.AddItem "Forwarded To CMPM"
   CmbDocStatus.AddItem "Forwarded To AM"
   CmbDocStatus.AddItem "Forwarded To Legal"
   'CmbDocStatus.AddItem "Rejected"
ElseIf LCase(Gen_UserRole) = "headcmg" Then
   CmbDocStatus.AddItem "Forwarded To CMG"
'   CmbDocStatus.AddItem "Forwarded To CMPM"
   CmbDocStatus.AddItem "Forwarded To AM"
   CmbDocStatus.AddItem "Forwarded To Legal"
   'CmbDocStatus.AddItem "Rejected"
ElseIf LCase(Gen_UserRole) = "legal" Or LCase(Gen_UserRole) = "power" Then
   CmbDocStatus.AddItem "Forwarded To CMG"
'   CmbDocStatus.AddItem "Forwarded To CMPM"
   CmbDocStatus.AddItem "Forwarded To AM"
   CmbDocStatus.AddItem "Approved"
   'CmbDocStatus.AddItem "Rejected"
ElseIf LCase(Gen_UserRole) = "amsandrms" Then
   CmbDocStatus.AddItem "Forwarded to CMG"
   CmbDocStatus.AddItem "Forwarded To CMPM"
End If

If CmbDocStatus.ListCount = 0 Then
   MsgBox "Can't Edit"
   Exit Sub
End If

Frame1.Enabled = True
Frame2.Enabled = True
Frame3.Enabled = True

CmbAgreRecAs.Text = TxtAgreeRecAs
TxtAgreeRecAs.Visible = False
CmbAgreRecAs.Visible = True

CmbStatus.Text = TxtStatus
TxtStatus.Visible = False
CmbStatus.Visible = False

CmbForwardedTo.Text = TxtForwardedTo
TxtForwardedTo.Visible = False
CmbForwardedTo.Visible = True

'CmbLocationID.Text = TxtLocationID
'TxtLocationID.Visible = False
'CmbLocationID.Visible = True
'
'CmbAgreeType.Text = TxtAgreeType
'TxtAgreeType.Visible = False
'CmbAgreeType.Visible = True
Call OnEditCreateAttachFile

SSTab1.Tab = 0
Call GButtonLock(Me, False)

End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub FirstCmd_Click()
RsTxn_GodownAgreeDocCycle.MoveFirst
Call Indata
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
   Frame2.Enabled = False
   Frame3.Enabled = False
   Frame1.Enabled = False
   Call GButtonLockAD(Me)
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If

Call Grid_Head1
Call Grid_Head2
Call Grid_Head3
Call Grid_Head5
TxtDocRemarks.Height = 600

If Gen_WcLocation <> "" Then
   Call TableTxn_GodownAgreeDocCycle("Where " & Gen_WcLocation & " Order by AgreementID")
Else
   Call TableTxn_GodownAgreeDocCycle("Order by AgreementID")
End If
If RsTxn_GodownAgreeDocCycle.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
Call GButtonLock(Me, True)
Call Indata
End Sub

Private Sub Grid_Head5()
MSHFlexGrid5.Clear
MSHFlexGrid5.Rows = 3
MSHFlexGrid5.FixedRows = 1
MSHFlexGrid5.Cols = 6

MSHFlexGrid5.TextMatrix(0, 0) = "Row ID"
MSHFlexGrid5.TextMatrix(0, 1) = "Document Type"
MSHFlexGrid5.TextMatrix(0, 2) = "Old Employee to Whom was Forwarded"
MSHFlexGrid5.TextMatrix(0, 3) = "New Employee to Whom is Forwarded"
MSHFlexGrid5.TextMatrix(0, 4) = "Updated By Employee"
MSHFlexGrid5.TextMatrix(0, 5) = "Concurrency"

MSHFlexGrid5.ColWidth(0) = 0
MSHFlexGrid5.ColWidth(1) = 2000
MSHFlexGrid5.ColWidth(2) = 3800
MSHFlexGrid5.ColWidth(3) = 3800
MSHFlexGrid5.ColWidth(4) = 2000
MSHFlexGrid5.ColWidth(5) = 2500

End Sub
Public Sub Indata()

TxtDocRemarks.Height = 450

Call ClearGodownFrame
Call ClearDocumentFrame

Frame1.Enabled = False
Frame2.Enabled = False
Frame3.Enabled = False

SSTab1.Tab = 0

Call Grid_Head2
Call Grid_Head3

CmbStatus.Visible = False
TxtStatus.Visible = False

CmbForwardedTo.Visible = False
TxtForwardedTo.Visible = True

CmbLocationID.Visible = False
TxtLocationID.Visible = True

CmbAgreeType.Visible = False
TxtAgreeType.Visible = True

tmp = " Select TGAC.AgreementID,TGAC.StartDate,TGAC.AgreementActualStartDate,TGAC.EntryDate,TGAC.ExpiryDate, "
tmp = tmp & " TGAC.Remarks,TGAC.CreatedBy,TGAC.CreatedDate,MEE.EmployeeID as CreateID,MEE.EmployeeEmailID as CreatedEmailID,"
tmp = tmp & " TGAC.UpdatedBy,TGAC.UpdatedDate,TGAC.Status,ML.LocationName,MUT.UnitType,ML.LocationID,MUT.UnitTypeID"
tmp = tmp & " from Txn_GodownAgreeDocCycle TGAC "
tmp = tmp & " Inner Join Mst_Employee MEE on TGAC.CreatedBy=MEE.EmployeeNo"
tmp = tmp & " Inner Join Mst_Location ML on TGAC.LocationID=ML.LocationID"
tmp = tmp & " Inner Join Mst_UnitType MUT on TGAC.UnitTypeID=MUT.UnitTypeID"
tmp = tmp & " Where TGAC.AgreementID=" & RsTxn_GodownAgreeDocCycle!AgreementID & ""
Call Tmp5Table
If TmpRs5.RecordCount > 0 Then
   TxtAgreementID = IIf(IsNull(TmpRs5!AgreementID), "", TmpRs5!AgreementID)
   TxtStartDate.Value = IIf(IsNull(TmpRs5!StartDate), Date, TmpRs5!StartDate)
   TxtAAggrStartDate.Value = IIf(IsNull(TmpRs5!AgreementActualStartDate), Date, TmpRs5!AgreementActualStartDate)
   TxtEntryDate.Value = IIf(IsNull(TmpRs5!EntryDate), Date, TmpRs5!EntryDate)
   TxtExpiryDate.Value = IIf(IsNull(TmpRs5!ExpiryDate), Date, TmpRs5!ExpiryDate)
   'TxtForwardedTo.Text = TmpRs5!EmployeeName & "~" & TmpRs5!EmployeeNo
   TxtLocationID.Text = IIf(IsNull(TmpRs5!Locationname), "", TmpRs5!Locationname)
   TxtAgreeType.Text = IIf(IsNull(TmpRs5!UnitType), "", TmpRs5!UnitType)
   TxtRemarks = IIf(IsNull(TmpRs5!Remarks), "", TmpRs5!Remarks)
   TxtCreated = IIf(IsNull(TmpRs5!CreatedBy), "", TmpRs5!CreatedBy) & " :- " & IIf(IsNull(TmpRs5!CreatedDate), "", TmpRs5!CreatedDate)
   TxtUpdated = IIf(IsNull(TmpRs5!UpdatedBy), "", TmpRs5!UpdatedBy) & " :- " & IIf(IsNull(TmpRs5!UpdatedDate), "", TmpRs5!UpdatedDate)
   mStatus = IIf(IsNull(TmpRs5!Status), "", TmpRs5!Status)
'   TxtUserRole.Text = IIf(IsNull(TmpRs5!groupname), "", TmpRs5!groupname)

   mLocationID = IIf(IsNull(TmpRs5!LocationID), "", TmpRs5!LocationID)
   mAgreeTypeID = IIf(IsNull(TmpRs5!UnitTypeID), "", TmpRs5!UnitTypeID)
   mCreateId = IIf(IsNull(TmpRs5!CreateID), "", TmpRs5!CreateID)
   mCreatedEmailID = IIf(IsNull(TmpRs5!CreatedEmailID), "", TmpRs5!CreatedEmailID)

   If LCase(mStatus) = "p" Then
      TxtStatus.Text = "Pending"
   End If

'   mForwardID = IIf(IsNull(TmpRs5!ForwardEmpolyeeID), 0, TmpRs5!ForwardEmpolyeeID)

   Call InDataDocument
   Call InDataGodown
   Call InDataLogHistory


End If

If RsTxn_GodownAgreeDocCycle.RecordCount = 1 Then
   Call GButtonLock_1(Me, True)
Else
   Call GButtonLock(Me, True)
End If
End Sub

Private Sub InDataDocument()
Call Grid_Head2

Gen_DbPathFromAzure = "N"

tmp = " Select MDT.DocumentType,TGAD.FileName,TGAD.[File],TGAD.[Desc],TGAD.Remarks,TGAD.DocumentTypeID,TGAD.DocDID,TGAD.Flag,ME.EmployeeName+'~'+ME.EmployeeNo as EmpName,ME.EmployeeID,ME.EmployeeEmailID"
tmp = tmp & " from Txn_GodownAgreeDocCycle TGAC"
tmp = tmp & " Inner Join Txn_GADCDocumentAttachDetail TGAD on TGAC.AgreementID=TGAD.AgreementID"
tmp = tmp & " Inner Join Mst_DocumentType MDT on TGAD.DocumentTypeID=MDT.DocumentTypeID"
tmp = tmp & " Inner Join Mst_Employee ME on TGAD.ForwardID=ME.EmployeeID"
tmp = tmp & " Where TGAC.AgreementID = " & TxtAgreementID.Text & ""

Call Tmp2Table

Gen_DbPathFromAzure = "Y"

MSHFlexGrid2.Rows = TmpRs2.RecordCount + 2
For I = 1 To TmpRs2.RecordCount
    MSHFlexGrid2.TextMatrix(I, 0) = IIf(IsNull(TmpRs2!DocumentType), "", TmpRs2!DocumentType)
    MSHFlexGrid2.TextMatrix(I, 1) = IIf(IsNull(TmpRs2!FileName), "", TmpRs2!FileName)
    MSHFlexGrid2.TextMatrix(I, 2) = IIf(IsNull(TmpRs2!Desc), "", TmpRs2!Desc)
    MSHFlexGrid2.TextMatrix(I, 3) = IIf(IsNull(TmpRs2!Remarks), "", TmpRs2!Remarks)

    If Right(App.Path, 1) = "\" Then
       File_Path = App.Path & "Agreement Attachment\Godown Agreements\" & Gen_UserLoginID & "\" & MSHFlexGrid2.TextMatrix(I, 1)
    Else
       File_Path = App.Path & "\Agreement Attachment\Godown Agreements\" & Gen_UserLoginID & "\" & MSHFlexGrid2.TextMatrix(I, 1)
    End If '    Comment by prajakta 22 Dec 2014
'    File_Path = App.Path & "\Agreement Attachment\Godown Agreements\" & MSHFlexGrid2.TextMatrix(i, 1)
'    Call LoadFileFromDB(File_Path, TmpRs2, "File")

    MSHFlexGrid2.TextMatrix(I, 4) = IIf(IsNull(TmpRs2!DocumentTypeID), "", TmpRs2!DocumentTypeID)
    MSHFlexGrid2.TextMatrix(I, 5) = IIf(IsNull(TmpRs2!DocDID), "", TmpRs2!DocDID)
    MSHFlexGrid2.TextMatrix(I, 1) = File_Path

    If LCase(TmpRs2!Flag) = "g" Then
       MSHFlexGrid2.TextMatrix(I, 6) = "Forwarded to CMG"
    ElseIf LCase(TmpRs2!Flag) = "p" Then
       MSHFlexGrid2.TextMatrix(I, 6) = "Forwarded to CMPM"
    ElseIf LCase(TmpRs2!Flag) = "f" Then
       MSHFlexGrid2.TextMatrix(I, 6) = "Forwarded to FAG"
    ElseIf LCase(TmpRs2!Flag) = "l" Then
       MSHFlexGrid2.TextMatrix(I, 6) = "Forwarded to Legal"
    ElseIf LCase(TmpRs2!Flag) = "a" Then
       MSHFlexGrid2.TextMatrix(I, 6) = "Approved"
    ElseIf LCase(TmpRs2!Flag) = "m" Then
       MSHFlexGrid2.TextMatrix(I, 6) = "Forwarded to AM"
'    ElseIf LCase(TmpRs2!Flag) = "r" Then
'       MSHFlexGrid2.TextMatrix(i, 6) = "Rejected"
    End If
    MSHFlexGrid2.TextMatrix(I, 7) = TmpRs2!Flag
    MSHFlexGrid2.TextMatrix(I, 8) = TmpRs2!EmpName
    MSHFlexGrid2.TextMatrix(I, 9) = TmpRs2!EmployeeID
    MSHFlexGrid2.TextMatrix(I, 10) = TmpRs2!EmployeeEmailID
    TmpRs2.MoveNext
Next
End Sub
Private Sub InDataLogHistory()
Call Grid_Head5
'Parag
'tmp = "Select TCS.AgreementID,MDT.DocumentType,MU.Employee_Name+'~'+ MU.Employee_ID,MU1.Employee_Name+'~'+ MU1.Employee_ID, " & _
'      " TCS.ForwardID,EMP.EmployeeName+'~'+ EMP.EmployeeNo as 'OldFwdTo', " & _
'      " EMP1.EmployeeName+'~'+ EMP1.EmployeeNo as 'NewFwdTo',LCFTH.UpdatedBy,LCFTH.Concurrency,TCS.CreatedDate,TCS.UpdatedDate  " & _
'      " From Log_GADCForwardToHistory LCFTH  " & _
'      " Left Join Mst_Employee EMP On LCFTH.OldForwardToID = EMP.EmployeeID  " & _
'      " Left Join Mst_Employee EMP1 On LCFTH.NewForwardToID = EMP1.EmployeeID " & _
'      " Right Join Txn_GADCDocumentAttachDetail TCS On LCFTH.AgreementRowID = TCS.AgreementID  " & _
'      " Inner Join Mst_DocumentType MDT On TCS.DocumentTypeID = MDT.DocumentTypeID And TCS.DocumentTypeID=LCFTH.DocumentTypeID " & _
'      " Inner Join Mst_Employee EMP2 On TCS.ForwardID = EMP2.EmployeeID  " & _
'      " Inner Join Mst_Users MU On TCS.CreatedBy = MU.Employee_ID  " & _
'      " Left Join Mst_Users MU1 On TCS.UpdatedBy = MU1.Employee_ID  " & _
'      " Where TCS.AgreementID = " & TxtAgreementID.Text & " " & _
'      " Order by LCFTH.Concurrency desc,MDT.DocumentType asc "


tmp = "Select TCS.AgreementID,MDT.DocumentType,MU.EmployeeName+'~'+ MU.EmployeeNo,MU1.EmployeeName+'~'+ MU1.EmployeeNo, " & _
      " TCS.ForwardID,EMP.EmployeeName+'~'+ EMP.EmployeeNo as 'OldFwdTo', " & _
      " EMP1.EmployeeName+'~'+ EMP1.EmployeeNo as 'NewFwdTo',LCFTH.UpdatedBy,LCFTH.Concurrency,TCS.CreatedDate,TCS.UpdatedDate  " & _
      " From Log_GADCForwardToHistory LCFTH  " & _
      " Left Join Mst_Employee EMP On LCFTH.OldForwardToID = EMP.EmployeeID  " & _
      " Left Join Mst_Employee EMP1 On LCFTH.NewForwardToID = EMP1.EmployeeID " & _
      " Right Join Txn_GADCDocumentAttachDetail TCS On LCFTH.AgreementRowID = TCS.AgreementID  " & _
      " Inner Join Mst_DocumentType MDT On TCS.DocumentTypeID = MDT.DocumentTypeID And TCS.DocumentTypeID=LCFTH.DocumentTypeID " & _
      " Inner Join Mst_Employee EMP2 On TCS.ForwardID = EMP2.EmployeeID  " & _
      " Inner Join Mst_Users MU On TCS.CreatedBy = MU.EmployeeNo  " & _
      " Left Join Mst_Users MU1 On TCS.UpdatedBy = MU1.EmployeeNo  " & _
      " Where TCS.AgreementID = " & TxtAgreementID.Text & " " & _
      " Order by LCFTH.Concurrency desc,MDT.DocumentType asc "

Call TmpTable
For iu = 1 To TmpRs.RecordCount
    MSHFlexGrid5.Rows = TmpRs.RecordCount + 2
    For C = 0 To MSHFlexGrid5.Cols - 1
         MSHFlexGrid5.TextMatrix(iu, 0) = IIf(IsNull(TmpRs!AgreementID), "", TmpRs!AgreementID)
         MSHFlexGrid5.TextMatrix(iu, 1) = IIf(IsNull(TmpRs!DocumentType), "", TmpRs!DocumentType)
         MSHFlexGrid5.TextMatrix(iu, 2) = IIf(IsNull(TmpRs.Fields(5)), TmpRs.Fields(2) & ":- " & Format(TmpRs!CreatedDate, "dd-mmm-yyyy"), TmpRs.Fields(5))
         MSHFlexGrid5.TextMatrix(iu, 3) = IIf(IsNull(TmpRs.Fields(6)), TmpRs.Fields(3) & ":- " & Format(TmpRs!UpdatedDate, "dd-mmm-yyyy"), TmpRs.Fields(6))
         MSHFlexGrid5.TextMatrix(iu, 4) = IIf(IsNull(TmpRs!UpdatedBy), "", TmpRs!UpdatedBy)
         MSHFlexGrid5.TextMatrix(iu, 5) = IIf(IsNull(TmpRs!Concurrency), "", TmpRs!Concurrency)
         MSHFlexGrid5.TextMatrix(iu, 5) = Format(MSHFlexGrid5.TextMatrix(iu, 5), "dd-mmm-yyyy hh:mm")
    Next
    TmpRs.MoveNext
Next
End Sub
Private Sub InDataGodown()
Call Grid_Head3

tmp = " Select TGD.GodownDID,ML.LocationName,MC.CMPName,MGI.InspectionID,TGD.GodOwnNo,ML.SM_Prefix,TGD.Remarks,MGI.CMPID,ML.LocationID "
tmp = tmp & " From Txn_GodownAgreeDocCycle TGAC"
tmp = tmp & " Inner Join Txn_GADCGodownDetail TGD on TGAC.AgreementID=TGD.AgreementID"
tmp = tmp & " Inner Join Mst_GodownInspection MGI on TGD.InspectionID=MGI.InspectionID"
tmp = tmp & " Inner Join Mst_CMP MC on MGI.CMPID=MC.CMPID"
tmp = tmp & " Inner Join Mst_Location ML on MC.LocationID=ML.LocationID"
tmp = tmp & " Where TGAC.AgreementID=" & Val(TxtAgreementID.Text) & ""

Call Tmp5Table
MSHFlexGrid3.Rows = TmpRs5.RecordCount + 3
For I = 2 To MSHFlexGrid3.Rows - 2
   MSHFlexGrid3.TextMatrix(I, 0) = TmpRs5!GodownDID
   MSHFlexGrid3.TextMatrix(I, 1) = TmpRs5!Locationname
   MSHFlexGrid3.TextMatrix(I, 2) = TmpRs5!CMPName
   MSHFlexGrid3.TextMatrix(I, 3) = TmpRs5!InspectionID
   MSHFlexGrid3.TextMatrix(I, 4) = TmpRs5!GodownNo
   MSHFlexGrid3.TextMatrix(I, 5) = TmpRs5!SM_Prefix
   MSHFlexGrid3.TextMatrix(I, 6) = TmpRs5!Remarks
   MSHFlexGrid3.TextMatrix(I, 7) = TmpRs5!LocationID
   MSHFlexGrid3.TextMatrix(I, 8) = TmpRs5!CMPID
   TmpRs5.MoveNext
Next
End Sub
Private Function getMaxLAID() As Double
Dim Retval As Double
tmp = "Select max(LAID) from Txn_LessorAgreement"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
getMaxLAID = Retval
End Function
Private Function getMaxLADetailID() As Double
Dim Retval As Double
tmp = "Select max(LADetailID) from Txn_LessorAgreementDetail"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
getMaxLADetailID = Retval
End Function

Private Function getMaxGLADID() As Double
Dim Retval As Double
tmp = "Select max(GLADID) from Txn_GodownLessorAgreeDetail"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
getMaxGLADID = Retval
End Function




Private Sub LastCmd_Click()
RsTxn_GodownAgreeDocCycle.MoveLast
Call Indata
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False
End Sub

Private Sub MSHFlexGrid1_Click()
Dim I As Variant
I = MSHFlexGrid1.RowSel
If I <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(I, 0) = "" Then Exit Sub
RsTxn_GodownAgreeDocCycle.MoveFirst
RsTxn_GodownAgreeDocCycle.Find "AgreementID = " & MSHFlexGrid1.TextMatrix(I, 0)
If Not RsTxn_GodownAgreeDocCycle.EOF Then
   Call Indata
   Frame0.Visible = False
End If
End Sub

Private Sub MSHFlexGrid2_Click()
'On Error GoTo MsgErr

Dim ans, r As Variant
Dim ID As Integer

r = MSHFlexGrid2.RowSel
File_Path = ""

If SaveCmd.Enabled = False Then
   If r < 1 Then Exit Sub
      If MSHFlexGrid2.TextMatrix(r, 1) = "" Then Exit Sub
      ans = MsgBox("Do You want to view this document ", vbYesNo)
      If ans = vbYes Then
         
         tmp = " Select MDT.DocumentType,TGAD.FileName,TGAD.[File],TGAD.[Desc],TGAD.Remarks,TGAD.DocumentTypeID,TGAD.DocDID,TGAD.Flag,ME.EmployeeName+'~'+ME.EmployeeNo as EmpName,ME.EmployeeID,ME.EmployeeEmailID"
         tmp = tmp & " from Txn_GodownAgreeDocCycle TGAC"
         tmp = tmp & " Inner Join Txn_GADCDocumentAttachDetail TGAD on TGAC.AgreementID=TGAD.AgreementID"
         tmp = tmp & " Inner Join Mst_DocumentType MDT on TGAD.DocumentTypeID=MDT.DocumentTypeID"
         tmp = tmp & " Inner Join Mst_Employee ME on TGAD.ForwardID=ME.EmployeeID"
         tmp = tmp & " Where TGAC.AgreementID = " & Val(TxtAgreementID.Text) & " And DocDID=" & Val(MSHFlexGrid2.TextMatrix(r, 5)) & ""
         Call Tmp2Table
         
         
         
         
         If TmpRs2.RecordCount > 0 Then
            
            Gen_DbPathFromAzure = "N"
            Call TableTxn_GADCDocumentAttachDetail(" Where AgreementID = " & Val(TxtAgreementID.Text) & " And DocDID=" & TmpRs2!DocDID & " ")
            Gen_DbPathFromAzure = "Y"
            
            If Right(App.Path, 1) = "\" Then
               File_Path = App.Path & "Agreement Attachment\Godown Agreements\" & Gen_UserLoginID
            Else
               File_Path = App.Path & "\Agreement Attachment\Godown Agreements\" & Gen_UserLoginID
            End If
            
            If oFs.FolderExists(File_Path) = True Then
               Set oFolder = oFs.GetFolder(File_Path)
               
'               For Each oFile In oFolder.Files
'                   Call oFs.DeleteFile(oFile.Path, True)
'               Next
            Else
               If Not oFs.FolderExists(File_Path) Then
                  Call oFs.CreateFolder(File_Path)
               End If
            End If
            
            If Right(App.Path, 1) = "\" Then
               File_Path = App.Path & "Agreement Attachment\Godown Agreements\" & Gen_UserLoginID & "\" & RsTxn_GADCDocumentAttachDetail!FileName
            Else
               File_Path = App.Path & "\Agreement Attachment\Godown Agreements\" & Gen_UserLoginID & "\" & RsTxn_GADCDocumentAttachDetail!FileName
            End If
            Call LoadFileFromDB(File_Path, RsTxn_GADCDocumentAttachDetail, "File")
         End If
         If Right(App.Path, 1) = "\" Then
            File_Path = App.Path & "Agreement Attachment\Godown Agreements\" & Gen_UserLoginID
         Else
            File_Path = App.Path & "\Agreement Attachment\Godown Agreements\" & Gen_UserLoginID
         End If
'         If Right(App.Path, 1) = "\" Then
'            Shell "Explorer.exe " & File_Path, vbMaximizedFocus
'         Else
            Shell "Explorer.exe " & File_Path, vbMaximizedFocus
'         End If
      End If
   Exit Sub
End If

If Frame2.Enabled = False Then Exit Sub

If MSHFlexGrid2.TextMatrix(r, 5) <> "0" And MSHFlexGrid2.TextMatrix(r, 5) <> "" Then
   If LCase(MSHFlexGrid2.TextMatrix(r, 7)) = "a" Then
      MsgBox "Can't Edit, Document is " & MSHFlexGrid2.TextMatrix(r, 6) & ""
      Exit Sub
   End If
   If CheckStatus(MSHFlexGrid2.TextMatrix(r, 7)) = True Then
      MsgBox "Can't Edit, Document is " & MSHFlexGrid2.TextMatrix(r, 6) & ""
      Call ClearDocumentFrame
      Exit Sub
   End If
End If

If r < 1 Or r = MSHFlexGrid2.Rows - 1 Then
   Call ClearDocumentFrame
   Exit Sub
End If

If MSHFlexGrid2.TextMatrix(r, 1) = "" Then Exit Sub

CmbDocument.Text = MSHFlexGrid2.TextMatrix(r, 0)
TxtFileName.Text = MSHFlexGrid2.TextMatrix(r, 1)
TxtRemDes.Text = MSHFlexGrid2.TextMatrix(r, 2)
TxtDocRemarks.Text = MSHFlexGrid2.TextMatrix(r, 3)
mDocumentTypeID = MSHFlexGrid2.TextMatrix(r, 4)
CmbDocStatus.Text = MSHFlexGrid2.TextMatrix(r, 6)
CmbDocForwardedTo.Text = MSHFlexGrid2.TextMatrix(r, 8)
mDocForwardID = MSHFlexGrid2.TextMatrix(r, 9)
CmdAddDocList.Caption = "Update List"
CmdAddDocList.SetFocus

Exit Sub

MsgErr:
MsgBox Err.Description
End Sub

Private Sub MSHFlexGrid2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF5 Then
   MsgBox MSHFlexGrid2.TextMatrix(I, 1)
End If
End Sub


Private Sub MSHFlexGrid3_Click()
If SaveCmd.Enabled = False Then Exit Sub
Dim r As Variant

If Frame3.Enabled = False Then Exit Sub

r = MSHFlexGrid3.RowSel
If r < 2 Or r = MSHFlexGrid3.Rows - 1 Then
   Call ClearGodownFrame
   Exit Sub
End If
If MSHFlexGrid3.TextMatrix(r, 1) = "" Then Exit Sub

mInspectionID = MSHFlexGrid3.TextMatrix(r, 3)
CmbCMP.Text = MSHFlexGrid3.TextMatrix(r, 2)
mCMPID = MSHFlexGrid3.TextMatrix(r, 8)
CmbGodown.Text = MSHFlexGrid3.TextMatrix(r, 3)
TxtGodownNo.Text = MSHFlexGrid3.TextMatrix(r, 4)
TxtGodownAddress.Text = MSHFlexGrid3.TextMatrix(r, 6)
CmdAdd2ListGodown.Caption = "Update List"
CmdAdd2ListGodown.SetFocus
End Sub

Private Sub MSHFlexGrid4_Click()
Dim ans, r As Variant
Dim ID As Integer
r = MSHFlexGrid4.RowSel
If r < 1 Then Exit Sub
If MSHFlexGrid4.TextMatrix(r, 1) = "" Then Exit Sub

If SaveCmd.Enabled = False Then

   ans = MsgBox("Do You want to view this document ", vbYesNo)
   If ans = vbYes Then
      If Right(App.Path, 1) = "\" Then
         Shell "Explorer.exe " & App.Path & "Agreement Attachment", vbMaximizedFocus
      Else
         Shell "Explorer.exe " & App.Path & "\Agreement Attachment", vbMaximizedFocus
      End If
   End If
   Exit Sub
End If

lblAttachment.Caption = MSHFlexGrid4.TextMatrix(r, 3)
TxtAttachmentRemarks = MSHFlexGrid4.TextMatrix(r, 2)

CmdRemoveAttachment.Enabled = True
If MSHFlexGrid4.TextMatrix(r, 4) <> "" Then
   CmdRemoveAttachment.Enabled = False
End If
CmdAttachment.Caption = "Update Attachment"
End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsTxn_GodownAgreeDocCycle.MoveNext
If RsTxn_GodownAgreeDocCycle.EOF Then
   RsTxn_GodownAgreeDocCycle.MoveLast
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
RsTxn_GodownAgreeDocCycle.MovePrevious
If RsTxn_GodownAgreeDocCycle.BOF Then
   RsTxn_GodownAgreeDocCycle.MoveFirst
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

Private Sub SaveCmd_Click()
On Error GoTo MsgErr

mStrMode = "Save"

If TxtExpiryDate <= TxtStartDate Then
   MsgBox "Expiry Date should be Greater than Start Date."
   TxtExpiryDate.SetFocus
   Exit Sub
End If

If TxtExpiryDate <= TxtAAggrStartDate Then
   MsgBox "Expiry Date should be Greater than Actual Agreement Start Date."
   TxtExpiryDate.SetFocus
   Exit Sub
End If

'If CmbForwardedTo.Text = "" Then
'   MsgBox "Blank Forwarded To Not Allowed"
'   Exit Sub
'End If

If CheckDuplicateDocumentAttach = True Then
   MsgBox "Duplicate Document Attachment Not Allowed"
   Exit Sub
End If

If MSHFlexGrid2.Rows < 2 Then
   MsgBox "Document Attachment Missing."
   Exit Sub
ElseIf MSHFlexGrid2.Rows = 2 Then
   If MSHFlexGrid2.TextMatrix(1, 0) = "" Or MSHFlexGrid2.TextMatrix(1, 1) = "" Then
      MsgBox "Document Attachment Missing."
      Exit Sub
   End If
End If

If MSHFlexGrid3.Rows < 4 Then
   MsgBox "Please provide Godown details!!!"
   SSTab1.Tab = 1
   Exit Sub
End If

If CheckDescription = True Then
   MsgBox "Blank Document Attachment Description Not Allowed"
   Exit Sub
End If

mStatus = "P"

If Edit_Flg = "A" Then
   RsTxn_GodownAgreeDocCycle.AddNew
   RsTxn_GodownAgreeDocCycle!AgreementID = GetMaxID("AgreementID", "Txn_GodownAgreeDocCycle")
   RsTxn_GodownAgreeDocCycle!G_UID = GetGUID
   TxtAgreementID.Text = RsTxn_GodownAgreeDocCycle!AgreementID
Else
   RsTxn_GodownAgreeDocCycle.MoveFirst
   RsTxn_GodownAgreeDocCycle.Find " AgreementID= '" & Val(TxtAgreementID.Text) & "'"
End If

RsTxn_GodownAgreeDocCycle!AgreementID = Val(TxtAgreementID.Text)
RsTxn_GodownAgreeDocCycle!EntryDate = Format(TxtEntryDate.Value, Gen_DatFormat)
RsTxn_GodownAgreeDocCycle!StartDate = Format(TxtStartDate.Value, Gen_DatFormat)
RsTxn_GodownAgreeDocCycle!AgreementActualStartDate = Format(TxtAAggrStartDate.Value, Gen_DatFormat)
RsTxn_GodownAgreeDocCycle!ExpiryDate = Format(TxtExpiryDate.Value, Gen_DatFormat)
RsTxn_GodownAgreeDocCycle!LocationID = mLocationID
RsTxn_GodownAgreeDocCycle!UnitTypeID = mAgreeTypeID
RsTxn_GodownAgreeDocCycle!Remarks = TxtRemarks.Text
RsTxn_GodownAgreeDocCycle!Status = mStatus

If IsNull(RsTxn_GodownAgreeDocCycle!CreatedBy) = True Or RsTxn_GodownAgreeDocCycle!CreatedBy = "" Then
   RsTxn_GodownAgreeDocCycle!CreatedBy = Gen_UserLoginID
   RsTxn_GodownAgreeDocCycle!CreatedDate = Now
Else
   RsTxn_GodownAgreeDocCycle!UpdatedBy = Gen_UserLoginID
   RsTxn_GodownAgreeDocCycle!UpdatedDate = Now
End If

RsTxn_GodownAgreeDocCycle.Update

Call SaveDocumentAttachDetails
Call SaveGodownDetails
SaveCmd.Enabled = False
Call Indata
Call Gen_Mail

Exit Sub

MsgErr:
MsgBox Err.Description
Call Gen_EmailForError("parag.r@ncml.com", "", Err.Description, mStrMode)
End Sub
Private Function GetMaxID(mColumn_ As Variant, mTable_ As Variant)
Dim Retval As Double
tmp = "Select Max(" & mColumn_ & ") from " & mTable_ & ""
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 1, TmpRs.Fields(0) + 1)
GetMaxID = Retval
End Function
Private Sub GetGodwn()
mFlag = True
mGodownID = ""
mRow = 0
j = 1
If MSHFlexGrid1.Rows < 3 Then
   MsgBox "No Data Exists"
   Exit Sub
End If

Call Grid_Head5
For I = 2 To MSHFlexGrid1.Rows - 2
    mGodownID = ""
    mFlag = True
    If MSHFlexGrid1.TextMatrix(I, 0) <> "" Then
       mGodownID = MSHFlexGrid1.TextMatrix(I, 0)
       For j = 1 To MSHFlexGrid6.Rows - 1
           If MSHFlexGrid6.TextMatrix(j, 0) = mGodownID Then
              mFlag = False
              Exit For
           End If
       Next
       If mFlag = True Then
          mRow = MSHFlexGrid6.Rows - 1
          MSHFlexGrid6.TextMatrix(mRow, 0) = mGodownID
          MSHFlexGrid6.TextMatrix(mRow, 1) = MSHFlexGrid1.TextMatrix(I, 3)
          MSHFlexGrid6.Rows = MSHFlexGrid6.Rows + 1
        End If
    End If
Next
'Call CheckGodownRent
End Sub

Private Function CheckRaintPaidUpto() As Boolean
Dim Retval As Boolean
Retval = True
For I = 2 To MSHFlexGrid1.Rows - 1
    If MSHFlexGrid1.TextMatrix(I, 10) <> "" Then
       If TxtExpiryDate < CDate(MSHFlexGrid1.TextMatrix(I, 10)) Then
          Retval = False
          Exit For
       End If
    End If
Next
CheckRaintPaidUpto = Retval
End Function
Private Function CheckGodownsExpiryDate() As Boolean
Dim Retval As Boolean
Retval = True
For I = 2 To MSHFlexGrid1.Rows - 1
    If MSHFlexGrid1.TextMatrix(I, 1) <> "" Then
       If CheckGodownExpiry(Val(MSHFlexGrid1.TextMatrix(I, 0)), "S") = False Then
          Retval = False
          Exit For
       End If
    End If
Next
CheckGodownsExpiryDate = Retval
End Function
Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   If RsTxn_GodownAgreeDocCycle.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   RsTxn_GodownAgreeDocCycle.MoveFirst
   Call ClearGodownFrame
   RsTxn_GodownAgreeDocCycle.Find "AgreementID = " & Val(TxtAgreementID) & ""
   If RsTxn_GodownAgreeDocCycle.EOF Then
      RsTxn_GodownAgreeDocCycle.MoveLast
   End If
   Call Indata
   Exit Sub
End If


Frame0.Left = Frame1.Left
Frame0.Top = Frame1.Top
Frame0.Height = Me.Height - 600
Frame0.Width = Frame1.Width
Frame0.Visible = True

MSHFlexGrid1.Width = Frame0.Width - 300
MSHFlexGrid1.Height = Frame0.Height - 2200
STo = Date
SFrom = Date - 30
SSFrom = Date - 365
SSTo = Date
RsTxn_GodownAgreeDocCycle.MoveFirst
RsTxn_GodownAgreeDocCycle.Filter = ""

End Sub

Private Sub TxtAdvanceGiven_LostFocus()

If TxtAdvanceGiven = "" Then Exit Sub
If IsNumeric(TxtAdvanceGiven) = False Or Val(TxtAdvanceGiven) < 0 Then
   MsgBox "Invalid numeric format!!!"
   TxtAdvanceGiven.SetFocus
   Exit Sub
End If
TxtAdvanceGiven = Format(Val(TxtAdvanceGiven), "#####0.00")
TxtBalance = Format(Val(TxtAdvanceGiven) - (Val(TxtAdvAdjusted) + Val(TxtAdvanceRemain)), "#####0.00")

End Sub


Private Sub CmdGetFile_Click()

With cdialog
     .InitDir = App.Path
     .DialogTitle = "Select File"
     .FilterIndex = 1
     .ShowOpen
End With

If cdialog.CancelError Then
   lblAttachment.Caption = ""
Else
   strFilename = cdialog.FileName
   If strFilename = "" Then Exit Sub
   lblAttachment.Refresh
   lblAttachment.Caption = strFilename 'GetFileFromPath(strFilename)
End If

End Sub

Private Sub Grid_Head2()
With MSHFlexGrid2
 .Clear
 .Rows = 2
 .Cols = 11
 .FixedRows = 1
 .FixedCols = 0
 .WordWrap = True

 .TextMatrix(0, 0) = "Document Type"
 .TextMatrix(0, 1) = "Attachment"
 .TextMatrix(0, 2) = "Attachment Description"
 .TextMatrix(0, 3) = "Remarks"
 .TextMatrix(0, 4) = "Document ID"
 .TextMatrix(0, 5) = "Row Id"
 .TextMatrix(0, 6) = "Status"
 .TextMatrix(0, 7) = "Database Status"
 .TextMatrix(0, 8) = "Forwarded To"
 .TextMatrix(0, 9) = "EmpID"
 .TextMatrix(0, 10) = "EmailID"

 .ColWidth(0) = 1600
 .ColWidth(1) = 3500
 .ColWidth(2) = 3500
 .ColWidth(3) = 4000
 .ColWidth(4) = 0
 .ColWidth(5) = 0
 .ColWidth(6) = 2000
 .ColWidth(7) = 0
 .ColWidth(8) = 3000
 .ColWidth(9) = 0
 .ColWidth(10) = 0
End With
End Sub

Public Sub SaveDocumentAttachDetails()
Dim DAD As Variant

mStrMode = "SaveDocumentAttachDetails"

If MSHFlexGrid2.Rows <= 1 Then Exit Sub

For I = 1 To MSHFlexGrid2.Rows - 1
    If MSHFlexGrid2.TextMatrix(I, 1) <> "" Then
       tmp = " Insert into Log_GADCForwardToHistory" & _
             " (AgreementRowID," & _
             " OldForwardToID," & _
             " NewForwardToID," & _
             " DocumentTypeID," & _
             " UpdatedBy)" & _
             " ( Select AgreementID,ForwardID," & MSHFlexGrid2.TextMatrix(I, 9) & ",DocumentTypeID,'" & Gen_UserLoginID & "'" & _
             " from Txn_GADCDocumentAttachDetail Where AgreementId=" & TxtAgreementID.Text & " And DocumentTypeId=" & MSHFlexGrid2.TextMatrix(I, 4) & ") "
             Call Tmp1Table
    End If
Next

Gen_DbPathFromAzure = "N"

tmp = " Delete From Txn_GADCDocumentAttachDetail where AgreementID=" & TxtAgreementID.Text & " "
Call TmpTable

Call TableTxn_GADCDocumentAttachDetail("Where AgreementID=" & TxtAgreementID.Text & "")
For I = 1 To MSHFlexGrid2.Rows - 1
    If MSHFlexGrid2.TextMatrix(I, 1) = "" Then Exit For
       RsTxn_GADCDocumentAttachDetail.AddNew
       RsTxn_GADCDocumentAttachDetail!DocDID = GetMaxID("DocDID", "Txn_GADCDocumentAttachDetail")
       RsTxn_GADCDocumentAttachDetail!AgreementID = TxtAgreementID.Text
       RsTxn_GADCDocumentAttachDetail!DocumentTypeID = MSHFlexGrid2.TextMatrix(I, 4)

       mFlag = SaveFileToDB(MSHFlexGrid2.TextMatrix(I, 1), RsTxn_GADCDocumentAttachDetail, "File")
       mFlag = GetFileFromPath(MSHFlexGrid2.TextMatrix(I, 1))
       If InStr(1, mFlag, "Agreement Attachment-" & TxtAgreementID & "-", vbTextCompare) > 0 Then
          RsTxn_GADCDocumentAttachDetail!FileName = mFlag
       Else
          'RsTxn_GADCDocumentAttachDetail!FileName = GetTimeStamp & "Agreement Attachment" & TxtAgreementID & "-" & mFlag
          RsTxn_GADCDocumentAttachDetail!FileName = "Agreement Attachment-" & TxtAgreementID & "-" & GetTimeStamp & "-" & mFlag
       End If
       RsTxn_GADCDocumentAttachDetail!Desc = MSHFlexGrid2.TextMatrix(I, 2)
       RsTxn_GADCDocumentAttachDetail!Remarks = MSHFlexGrid2.TextMatrix(I, 3)
       RsTxn_GADCDocumentAttachDetail!G_UID = GetGUID

    If LCase(MSHFlexGrid2.TextMatrix(I, 6)) = "forwarded to cmg" Then
       mDocStatus = "G" 'cmg
    ElseIf LCase(MSHFlexGrid2.TextMatrix(I, 6)) = "forwarded to cmpm" Then
       mDocStatus = "P" 'cmpm
    ElseIf LCase(MSHFlexGrid2.TextMatrix(I, 6)) = "forwarded to fag" Then
       mDocStatus = "F" 'fag
    ElseIf LCase(MSHFlexGrid2.TextMatrix(I, 6)) = "forwarded to legal" Then
       mDocStatus = "L" 'legal
    ElseIf LCase(MSHFlexGrid2.TextMatrix(I, 6)) = "approved" Then
       mDocStatus = "A"
    ElseIf LCase(MSHFlexGrid2.TextMatrix(I, 6)) = "forwarded to am" Then
       mDocStatus = "M"
'    ElseIf LCase(MSHFlexGrid2.TextMatrix(i, 6)) = "rejected" Then
'       mDocStatus = "R"
    End If

    RsTxn_GADCDocumentAttachDetail!Flag = mDocStatus
    RsTxn_GADCDocumentAttachDetail!ForwardID = MSHFlexGrid2.TextMatrix(I, 9)

    If IsNull(RsTxn_GADCDocumentAttachDetail!CreatedBy) = True Or RsTxn_GADCDocumentAttachDetail!CreatedBy = "" Then
       RsTxn_GADCDocumentAttachDetail!CreatedBy = Gen_UserLoginID
       RsTxn_GADCDocumentAttachDetail!CreatedDate = Now
    Else
       RsTxn_GADCDocumentAttachDetail!UpdatedBy = Gen_UserLoginID
       RsTxn_GADCDocumentAttachDetail!UpdatedDate = Now
    End If
    RsTxn_GADCDocumentAttachDetail.Update
Next

Gen_DbPathFromAzure = "Y"

End Sub

Public Sub SaveGodownDetails()
Dim GD As Variant

mStrMode = "SaveGodownDetails"

If MSHFlexGrid3.Rows <= 1 Then Exit Sub

tmp = " Delete From Txn_GADCGodownDetail where AgreementID=" & TxtAgreementID.Text & ""
Call TmpTable

Call TableTxn_GADCGodownDetail("Where AgreementID=" & TxtAgreementID.Text & "")
For I = 2 To MSHFlexGrid3.Rows - 1
    If MSHFlexGrid3.TextMatrix(I, 1) = "" Then Exit For
       RsTxn_GADCGodownDetail.AddNew
       RsTxn_GADCGodownDetail!GodownDID = GetMaxID("GodownDID", "Txn_GADCGodownDetail")
       RsTxn_GADCGodownDetail!AgreementID = TxtAgreementID.Text
       RsTxn_GADCGodownDetail!InspectionID = MSHFlexGrid3.TextMatrix(I, 3)
       RsTxn_GADCGodownDetail!GodownNo = MSHFlexGrid3.TextMatrix(I, 4)
       RsTxn_GADCGodownDetail!Remarks = MSHFlexGrid3.TextMatrix(I, 6)
       RsTxn_GADCGodownDetail!G_UID = GetGUID

       If IsNull(RsTxn_GADCGodownDetail!CreatedBy) = True Or RsTxn_GADCGodownDetail!CreatedBy = "" Then
          RsTxn_GADCGodownDetail!CreatedBy = Gen_UserLoginID
          RsTxn_GADCGodownDetail!CreatedDate = Now
       Else
          RsTxn_GADCGodownDetail!UpdatedBy = Gen_UserLoginID
          RsTxn_GADCGodownDetail!UpdatedDate = Now
       End If
       RsTxn_GADCGodownDetail.Update
Next
End Sub


Public Function CheckStatus(Status As Variant) As Boolean
'G-F to CMG , P-F to CMPM , L-F to Legal , F-F to FAG , R-Rejected , A-Approved


If LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "superadmin" And LCase(Gen_UserRole) <> "admin" And LCase(Gen_UserRole) <> "headcmg" And LCase(Gen_UserRole) <> "fag" And LCase(Gen_UserRole) <> "fagadmin" And LCase(Gen_UserRole) <> "legal" And LCase(Gen_UserRole) <> "amsandrms" Then
   If Status <> "P" Then
      CheckStatus = True
      Exit Function
   End If
''ElseIf LCase(Gen_UserRole) = "superadmin" Then
''   If Status = "P" Or Status = "M" Then  'Status = "L" Or
''      CheckStatus = True
''      Exit Function
''   End If
''-- 07 Sep 2017 Ullhas Patil to Change status and Forward to Allow
ElseIf LCase(Gen_UserRole) = "superadmin" Then
   If Status = "A" Then
      CheckStatus = True
      Exit Function
   End If
ElseIf LCase(Gen_UserRole) = "admin" Then
   If Status = "P" Or Status = "L" Or Status = "M" Then
      CheckStatus = True
      Exit Function
   End If
''ElseIf LCase(Gen_UserRole) = "head cmg" Then
''   If Status = "P" Or Status = "M" Then 'Status = "L" Or Comment by prajakta on 20Feb 2015
''      CheckStatus = True
''      Exit Function
''   End If
''-- 07 Sep 2017 Ullhas Patil to Change Status and Forward to Allow
ElseIf LCase(Gen_UserRole) = "headcmg" Then
   If Status = "A" Then
      CheckStatus = True
      Exit Function
   End If
ElseIf LCase(Gen_UserRole) = "legal" Then
   If Status = "P" Or Status = "M" Then
      CheckStatus = True
      Exit Function
   End If
ElseIf LCase(Gen_UserRole) = "amsandrms" Then
   If Status <> "M" Then
      CheckStatus = True
      Exit Function
   End If
End If

CheckStatus = False
End Function

Public Function Get_Gen_WCGADC() As Variant
If Gen_WcLocation <> "" Then
   tmp = " Select TGAC.AgreementID"
   tmp = tmp & " From Txn_GodownAgreeDocCycle TGAC"
   tmp = tmp & " Inner Join Txn_GADCGodownDetail TGD on TGAC.AgreementID=TGD.AgreementID"
   tmp = tmp & " Inner Join Mst_GodownInspection MGI on TGD.InspectionID=MGI.InspectionID"
   tmp = tmp & " Inner Join Mst_CMP MC on MGI.CMPID=MC.CMPID"
   tmp = tmp & " Inner Join Mst_Location ML on MC.LocationID=ML.LocationID"
   tmp = tmp & " Where ML." & Gen_WcLocation & ""
   Call Tmp1Table
   If TmpRs1.RecordCount > 0 Then
      For I = 1 To TmpRs1.RecordCount
         If Get_Gen_WCGADC = "" Then
            Get_Gen_WCGADC = TmpRs1!AgreementID
         Else
            Get_Gen_WCGADC = Get_Gen_WCGADC & "," & TmpRs1!AgreementID
         End If
         TmpRs1.MoveNext
      Next
   End If
End If
End Function

Public Function CheckDescription() As Boolean
Dim I
For I = 1 To MSHFlexGrid2.Rows - 1
    If MSHFlexGrid2.TextMatrix(I, 0) <> "" Then
       If MSHFlexGrid2.TextMatrix(I, 2) = "" Then
          CheckDescription = True
          Exit Function
       End If
    End If
Next
CheckDescription = False
End Function

Private Sub TxtRemarks_LostFocus()
If CmbDocument.Enabled = True Then CmbDocument.SetFocus
End Sub



Public Sub Gen_Mail()
Dim TOADD As Variant
Dim TOCC, mMsg, mGridMsg, mMsgSubject As Variant
Dim strSplitString() As String
Dim EmpNo_, EmpName_ As String
Dim mEmailId As Variant
TOADD = ""
TOCC = ""

mEmailId = ""
mMsg = ""
mMsg = "Dear Sir/Madam,"
mMsg = mMsg & "<br><br>"
mMsg = mMsg & "The following electronic documents are entered/updated in StoreMan. Please look into the documents pertaining to you/your department and if they"
mMsg = mMsg & "<br>"
mMsg = mMsg & "require any modifications/changes, please forward the documents to the field with your comments. "
mMsg = mMsg & "<br><br>"
mMsg = mMsg & "If the documents are OK in all respects, you may either forward them to the next level or APPROVE them in StoreMan (only Legal department can approve the documents) so that we can arrange for the finalised paper documents from the field for the record of Legal and CMG."
mMsg = mMsg & "<br><br>"
mMsg = mMsg & "<table border = 1>"
mMsg = mMsg & "<tr><td style=""color: blue""> Document Life Cycle ID  </td><td> " & TxtAgreementID.Text & " </td></tr>"
mMsg = mMsg & "<tr><td> Document Life Cycle Date  </td><td> " & Format(TxtEntryDate.Value, "dd-mmm-yyyy") & "  </td></tr>"
mMsg = mMsg & "<tr><td> Location </td><td> " & TxtLocationID.Text & "</td></tr>"
mMsg = mMsg & "<tr><td> Remarks </td><td> " & Trim(TxtRemarks.Text) & "</td></tr>"

'TxtDocRemarks

mGridMsg = "" 'Document details
For I = 0 To MSHFlexGrid2.Rows - 1
    If MSHFlexGrid2.TextMatrix(I, 1) <> "" Then
       If I = 0 Then
          mGridMsg = mGridMsg & "<tr><td style=""color: blue""> " & MSHFlexGrid2.TextMatrix(I, 0) & " </td><td style=""color: blue""> " & MSHFlexGrid2.TextMatrix(I, 8) & " </td><td style=""color: blue""> " & MSHFlexGrid2.TextMatrix(I, 6) & " </td><td style=""color: blue""> " & MSHFlexGrid2.TextMatrix(I, 3) & " </td></tr>"
       Else
          mGridMsg = mGridMsg & "<tr><td> " & MSHFlexGrid2.TextMatrix(I, 0) & " </td><td> " & MSHFlexGrid2.TextMatrix(I, 8) & " </td><td> " & MSHFlexGrid2.TextMatrix(I, 6) & " </td><td> " & MSHFlexGrid2.TextMatrix(I, 3) & " </td></tr>"
          If MSHFlexGrid2.TextMatrix(I, 10) <> "" Then
             If InStr(1, mEmailId, MSHFlexGrid2.TextMatrix(I, 10), vbTextCompare) = 0 Then
                mEmailId = IIf(mEmailId = "", MSHFlexGrid2.TextMatrix(I, 10), mEmailId & "," & MSHFlexGrid2.TextMatrix(I, 10))
             End If
          End If
       End If
    End If
Next

mMsg = mMsg & "<tr><td colspan=2><table border = 1><tr><td> Document Details </td><td><table border = 1>" & mGridMsg & "</table></td></tr></table>"

TOADD = mEmailId
TOCC = Gen_UserEmailID & ",cmg@ncml.com"  ',mayur.d@ncml.com,parag.r@ncml.com,sankar.k@ncml.com"

If InStr(1, TOCC, mCreatedEmailID, vbTextCompare) = 0 Then
   If InStr(1, TOADD, mCreatedEmailID, vbTextCompare) = 0 Then
      TOCC = TOCC & "," & mCreatedEmailID
   End If
End If

mMsgSubject = "Electronic Document sent for your approval in StoreMan Document Life Cycle ID- " & Val(TxtAgreementID) & "  "

'MsgBox "To : " & TOADD
'MsgBox "CC : " & TOCC

Call SendMail(TOADD, TOCC, mMsgSubject, mMsg)

End Sub
Private Function SendMail(ToADD_ As Variant, TOCC_ As Variant, MsgSubject_ As Variant, Message1_ As Variant)
Dim htmlStringhead As String
Dim htmlStringtail As String
Set oSMTP = New OSSMTP.SMTPSession

servertxt = "mail.ncml.com"

'htmlStringhead = "<html><body><font size= 3 face= Zurich BT color = Blue><table border = 1>"
'htmlStringtail = "</table></Font></body></html>"

htmlStringhead = "<html><body><font size= 3 face= Zurich BT color = black>"
htmlStringtail = "</Font></body></html>"


'With oSMTP
'  'connection
'  .Server = servertxt
'  'message
'  .MailFrom = Gen_UserEmailID
'  If IsNull(Trim(ToADD_)) = True Or Trim(ToADD_) = "" Then
'     ToADD_ = "parag.r@ncml.com" '"bhagwat.k@ncml.com"
'  End If
'  .SendTo = ToADD_
'  If ToADD_ <> TOCC_ Then
'     .CC = TOCC_
'  End If
'  .BCC = "parag.r@ncml.com"
'
''  .SendTo = "prajakta.p@ncml.com"
''  .CC = "prajakta.p@ncml.com"
'
'  .MessageSubject = MsgSubject_ '"Confidential Mail"
'  .MessageHTML = htmlStringhead & Message1_ & htmlStringtail
'  .TimeStamp = Now()
'
'  If LCase(Gen_EmailConfigFlag) = "y" Then
'     .SendEmail
'  End If
'End With

Dim objSMTP
Set objSMTP = CreateObject("MailBee.SMTP")
objSMTP.LicenseKey = "MBC900-5F470FAD67-D8E202A2E22927335F24F5A667BFD196"
objSMTP.ServerName = Gen_ReadMailServerIPINI("MailIP") ' = "mail.ncml.com" '"mail.ncml.com"
objSMTP.FromAddr = Gen_UserEmailID
If IsNull(Trim(ToADD_)) = True Or Trim(ToADD_) = "" Then
   ToADD_ = "manu.ag@ncml.com"
End If
objSMTP.ToAddr = ToADD_
If ToADD_ <> TOCC_ Then
   objSMTP.CCAddr = TOCC_
End If

objSMTP.Subject = MsgSubject_ & " (from Store-Man)" 'MsgSubject_ '"Message in HTML format"
objSMTP.BodyFormat = 1
objSMTP.BodyText = htmlStringhead & Message1_ & htmlStringtail 'htmlStringhead & MsgSubject_ & " :- " & Message1_ & htmlStringtail '"<html>This word is <b>bold</b></html>"
If LCase(Gen_EmailConfigFlag) = "y" Then
   objSMTP.Send
End If

End Function





Public Sub Xls_File()
On Error GoTo PassErr

Dim ans As Variant
MsgBox ("Wait till Next Message")
Set oXL = New Excel.Application
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & FileName)
Set oWS = oWB.Worksheets("Sheet1")
oWS.Name = "Godown Agreement Document Cycle"

oWS.ClearArrows
mRow = 1

oWS.Range("a:a").ColumnWidth = 12
oWS.Range("b:b").ColumnWidth = 9
oWS.Range("c:c").ColumnWidth = 14
oWS.Range("d:d").ColumnWidth = 8
oWS.Range("e:e").ColumnWidth = 9
oWS.Range("f:f").ColumnWidth = 8
oWS.Range("g:g").ColumnWidth = 10
oWS.Range("h:h").ColumnWidth = 11
oWS.Range("i:i").ColumnWidth = 10
oWS.Range("j:j").ColumnWidth = 20
oWS.Range("k:k").ColumnWidth = 20

tmp = "a" & mRow & ":" & "k" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "NATIONAL COLLATERAL MANAGEMENT SERVICES LIMITED"
oWS.Cells(mRow, C + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 10

mRow = mRow + 1
tmp = "a" & mRow & ":" & "k" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "Godown Agreement Document Cycle"
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 10
oWS.Cells(mRow, C + 1).HorizontalAlignment = -4108

mRow = mRow + 1
tmp = "a" & mRow & ":" & "b" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "Document Life Cycle ID :"
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 10

tmp = "c" & mRow & ":" & "d" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 3) = TxtAgreementID.Text
oWS.Cells(mRow, C + 3).Font.Size = 10
oWS.Cells(mRow, C + 3).HorizontalAlignment = -4108

tmp = "e" & mRow & ":" & "g" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 5) = "Document Life Cycle Date :"
oWS.Cells(mRow, C + 5).Font.Bold = True
oWS.Cells(mRow, C + 5).Font.Size = 10

tmp = "h" & mRow & ":" & "k" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 8) = "'" & Format(TxtEntryDate.Value, "dd-mmm-yyyy")
oWS.Cells(mRow, C + 8).Font.Size = 10
oWS.Cells(mRow, C + 8).HorizontalAlignment = -4108

mRow = mRow + 1
tmp = "a" & mRow & ":" & "b" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "Agreement Date :"
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 10

tmp = "c" & mRow & ":" & "d" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 3) = Format(TxtStartDate.Value, "dd-mmm-yyyy")
oWS.Cells(mRow, C + 3).Font.Size = 10
oWS.Cells(mRow, C + 3).HorizontalAlignment = -4108

tmp = "e" & mRow & ":" & "g" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 5) = "Agreement Effect Date :"
oWS.Cells(mRow, C + 5).Font.Bold = True
oWS.Cells(mRow, C + 5).Font.Size = 10

tmp = "h" & mRow & ":" & "h" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 8) = Format(TxtAAggrStartDate.Value, "dd-mmm-yyyy")
oWS.Cells(mRow, C + 8).Font.Bold = False
oWS.Cells(mRow, C + 8).Font.Size = 10
oWS.Cells(mRow, C + 8).HorizontalAlignment = -4108

tmp = "i" & mRow & ":" & "j" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 9) = "Expiry Date :"
oWS.Cells(mRow, C + 9).Font.Bold = True
oWS.Cells(mRow, C + 9).Font.Size = 10

tmp = "k" & mRow & ":" & "k" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 11) = Format(TxtExpiryDate.Value, "dd-mmm-yyyy")
oWS.Cells(mRow, C + 11).Font.Bold = False
oWS.Cells(mRow, C + 11).Font.Size = 10
oWS.Cells(mRow, C + 11).HorizontalAlignment = -4108

mRow = mRow + 1
tmp = "a" & mRow & ":" & "b" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "Location :"
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 10

tmp = "c" & mRow & ":" & "h" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 3) = TxtLocationID.Text
oWS.Cells(mRow, C + 3).Font.Size = 10
oWS.Cells(mRow, C + 3).HorizontalAlignment = -4108

tmp = "i" & mRow & ":" & "j" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 9) = "Agreement Type :"
oWS.Cells(mRow, C + 9).Font.Bold = True
oWS.Cells(mRow, C + 9).Font.Size = 10

tmp = "k" & mRow & ":" & "k" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 11) = TxtAgreeType.Text
oWS.Cells(mRow, C + 11).Font.Size = 10
oWS.Cells(mRow, C + 11).HorizontalAlignment = -4108

mRow = mRow + 1
tmp = "a" & mRow & ":" & "b" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "Remarks :"
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 10

tmp = "c" & mRow & ":" & "k" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 11) = TxtRemarks.Text
oWS.Cells(mRow, C + 11).Font.Bold = True
oWS.Cells(mRow, C + 11).Font.Size = 10

mRow = mRow + 1
tmp = "a" & mRow & ":" & "k" & mRow
oWS.Range(tmp).Merge
mRow = mRow + 1

tmp = "a" & mRow & ":" & "k" & mRow 'Goals Tab
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "Document Details  "
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 10

mRow = mRow + 1
tmp = "a" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "Document Type"
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 10
oWS.Cells(mRow, C + 1).RowHeight = 27.75
oWS.Cells(mRow, C + 1).WrapText = True

tmp = "b" & mRow & ":" & "c" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 2) = "FileName"
oWS.Cells(mRow, C + 2).Font.Bold = True
oWS.Cells(mRow, C + 2).Font.Size = 10
oWS.Cells(mRow, C + 2).RowHeight = 27.75

tmp = "d" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 4) = "Description"
oWS.Cells(mRow, C + 4).Font.Bold = True
oWS.Cells(mRow, C + 4).Font.Size = 10
oWS.Cells(mRow, C + 4).RowHeight = 27.75

tmp = "g" & mRow & ":" & "i" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 7) = "Remarks"
oWS.Cells(mRow, C + 7).Font.Bold = True
oWS.Cells(mRow, C + 7).Font.Size = 10
oWS.Cells(mRow, C + 7).RowHeight = 27.75

tmp = "j" & mRow & ":" & "j" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 10) = "Status"
oWS.Cells(mRow, C + 10).Font.Bold = True
oWS.Cells(mRow, C + 10).Font.Size = 10
oWS.Cells(mRow, C + 10).RowHeight = 27.75
oWS.Cells(mRow, C + 10).WrapText = True

tmp = "k" & mRow & ":" & "k" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 11) = "ForwardTo"
oWS.Cells(mRow, C + 11).Font.Bold = True
oWS.Cells(mRow, C + 11).Font.Size = 10
oWS.Cells(mRow, C + 11).RowHeight = 27.75


For I = 1 To MSHFlexGrid2.Rows - 1
    If MSHFlexGrid2.TextMatrix(I, 0) = "" Then Exit For

    mRow = mRow + 1
    C = 0
    tmp = "a" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 1) = MSHFlexGrid2.TextMatrix(I, 0)

    tmp = "b" & mRow & ":c" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 2) = MSHFlexGrid2.TextMatrix(I, 1)
    oWS.Cells(mRow, C + 2).WrapText = True

    tmp = "d" & mRow & ":f" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 4) = Replace(MSHFlexGrid2.TextMatrix(I, 2), vbNewLine, "")
    oWS.Cells(mRow, C + 4).WrapText = True

    tmp = "g" & mRow & ":i" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 7) = MSHFlexGrid2.TextMatrix(I, 3)
    oWS.Cells(mRow, C + 7).WrapText = True

    tmp = "j" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 10) = MSHFlexGrid2.TextMatrix(I, 6)

    tmp = "k" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 11) = MSHFlexGrid2.TextMatrix(I, 8)
Next I

For j = 1 To mRow
    tmp = "a" & j & ":k" & j
    oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
Next j
oWS.PageSetup.Orientation = xlPortrait

mRow = mRow + 1
tmp = "a" & mRow & ":" & "k" & mRow
oWS.Range(tmp).Merge
mRow = mRow + 1

tmp = "a" & mRow & ":" & "k" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "Godown Details  "
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 10

mRow = mRow + 1
tmp = "a" & mRow & ":" & "a" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "Sno"
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 10
oWS.Cells(mRow, C + 1).RowHeight = 27.75
oWS.Cells(mRow, C + 1).WrapText = True

tmp = "b" & mRow & ":" & "d" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 2) = "CMP Name"
oWS.Cells(mRow, C + 2).Font.Bold = True
oWS.Cells(mRow, C + 2).Font.Size = 10
oWS.Cells(mRow, C + 2).RowHeight = 27.75
oWS.Cells(mRow, C + 2).WrapText = True

tmp = "e" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 5) = "Viability ID"
oWS.Cells(mRow, C + 5).Font.Bold = True
oWS.Cells(mRow, C + 5).Font.Size = 10
oWS.Cells(mRow, C + 5).RowHeight = 27.75

tmp = "g" & mRow & ":" & "g" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 7) = "Godown No"
oWS.Cells(mRow, C + 7).Font.Bold = True
oWS.Cells(mRow, C + 7).Font.Size = 10
oWS.Cells(mRow, C + 7).RowHeight = 27.75

tmp = "h" & mRow & ":" & "k" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 8) = "Address"
oWS.Cells(mRow, C + 8).Font.Bold = True
oWS.Cells(mRow, C + 8).Font.Size = 10
oWS.Cells(mRow, C + 8).RowHeight = 27.75
oWS.Cells(mRow, C + 8).WrapText = True


For I = 2 To MSHFlexGrid3.Rows - 2
    If MSHFlexGrid3.TextMatrix(I, 0) = "" Then Exit For
    mRow = mRow + 1
    C = 0

    tmp = "a" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 1) = I

    tmp = "b" & mRow & ":d" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 2) = MSHFlexGrid3.TextMatrix(I, 2)

    tmp = "e" & mRow & ":f" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 5) = MSHFlexGrid3.TextMatrix(I, 3)

    tmp = "g" & mRow & ":g" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 7) = MSHFlexGrid3.TextMatrix(I, 4)

    tmp = "h" & mRow & ":k" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 8) = MSHFlexGrid3.TextMatrix(I, 6)
    oWS.Cells(mRow, C + 8).WrapText = True
Next I

mRow = mRow + 1
tmp = "a" & mRow & ":" & "k" & mRow
oWS.Range(tmp).Merge

'mRow = mRow + 1
For j = 1 To mRow
    tmp = "a" & j & ":k" & j
    oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
Next j
oWS.PageSetup.Orientation = xlPortrait

oWB.Save
oXL.Visible = True
Exit Sub

PassErr:
oXL.Quit
MsgBox Err.Description, vbExclamation, "Error"

End Sub


