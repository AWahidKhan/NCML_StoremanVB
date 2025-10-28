VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmTxn_LessorAgreement 
   Caption         =   "Lessor Agreement"
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
      Height          =   1965
      Left            =   360
      TabIndex        =   63
      Top             =   6720
      Visible         =   0   'False
      Width           =   13995
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
         Left            =   7170
         MaxLength       =   12
         TabIndex        =   90
         Top             =   1425
         Width           =   1890
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
         ItemData        =   "FrmTxn_LessorAgreement.frx":0000
         Left            =   4665
         List            =   "FrmTxn_LessorAgreement.frx":000D
         Sorted          =   -1  'True
         TabIndex        =   88
         Top             =   1440
         Width           =   2415
      End
      Begin VB.ComboBox CmbSLessor 
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
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   69
         Top             =   1440
         Width           =   4485
      End
      Begin VB.ComboBox CmbSCMP 
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
         Left            =   6240
         Sorted          =   -1  'True
         TabIndex        =   67
         Top             =   570
         Width           =   3645
      End
      Begin VB.ComboBox CmbSGodown 
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
         Left            =   10080
         Sorted          =   -1  'True
         TabIndex        =   68
         Top             =   600
         Width           =   2925
      End
      Begin VB.ComboBox CmbSAgreRecAs 
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
         ItemData        =   "FrmTxn_LessorAgreement.frx":002E
         Left            =   3210
         List            =   "FrmTxn_LessorAgreement.frx":0038
         Sorted          =   -1  'True
         TabIndex        =   65
         Top             =   570
         Width           =   3000
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
         Left            =   10185
         TabIndex        =   71
         Top             =   1425
         Width           =   1000
      End
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
         Left            =   9165
         TabIndex        =   70
         Top             =   1425
         Width           =   1000
      End
      Begin MSComCtl2.DTPicker STo 
         Height          =   375
         Left            =   1695
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   555
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
         Format          =   40697857
         CurrentDate     =   36494
      End
      Begin MSComCtl2.DTPicker SFrom 
         Height          =   375
         Left            =   135
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   555
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
         Format          =   40697857
         CurrentDate     =   36494
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
         Height          =   1455
         Left            =   120
         TabIndex        =   73
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
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Agreement ID"
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
         Left            =   7470
         TabIndex        =   91
         Top             =   1125
         Width           =   1230
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
         Left            =   5520
         TabIndex        =   89
         Top             =   1080
         Width           =   555
      End
      Begin VB.Label Label14 
         Caption         =   "Lessor Name"
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
         Left            =   1710
         TabIndex        =   82
         Top             =   1110
         Width           =   1320
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   7755
         TabIndex        =   81
         Top             =   210
         Width           =   435
      End
      Begin VB.Label Label12 
         Caption         =   "Godown"
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
         Left            =   11040
         TabIndex        =   80
         Top             =   240
         Width           =   1320
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Agreement Received As"
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
         Left            =   3630
         TabIndex        =   77
         Top             =   255
         Width           =   2205
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
         Left            =   13815
         TabIndex        =   76
         Top             =   1455
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
         Left            =   11340
         TabIndex        =   75
         Top             =   1440
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
         Left            =   480
         TabIndex        =   74
         Top             =   255
         Width           =   855
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
         Left            =   1995
         TabIndex        =   72
         Top             =   255
         Width           =   855
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1080
      Left            =   120
      TabIndex        =   31
      Top             =   8760
      Width           =   15015
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New"
         Height          =   400
         Left            =   105
         TabIndex        =   41
         Top             =   195
         Width           =   3060
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   400
         Left            =   105
         TabIndex        =   40
         Top             =   600
         Width           =   3060
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   400
         Left            =   3165
         TabIndex        =   39
         Top             =   195
         Width           =   3060
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   400
         Left            =   3165
         TabIndex        =   38
         Top             =   600
         Width           =   3060
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   400
         Left            =   6225
         TabIndex        =   37
         Top             =   195
         Width           =   2940
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   400
         Left            =   6225
         TabIndex        =   36
         Top             =   600
         Width           =   2940
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   400
         Left            =   9165
         TabIndex        =   35
         Top             =   195
         Width           =   2940
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   400
         Left            =   9165
         TabIndex        =   34
         Top             =   600
         Width           =   2940
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "Search"
         Height          =   400
         Left            =   12105
         TabIndex        =   33
         Top             =   195
         Width           =   2820
      End
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   400
         Left            =   12105
         TabIndex        =   32
         Top             =   600
         Width           =   2820
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2835
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   15015
      Begin VB.CheckBox ChkBRAgreement 
         Caption         =   "BR Agreement"
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
         Left            =   5640
         TabIndex        =   108
         Top             =   2520
         Width           =   2640
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
         Left            =   3615
         TabIndex        =   92
         Top             =   2280
         Width           =   1920
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
         Height          =   315
         Left            =   13620
         TabIndex        =   87
         Top             =   255
         Width           =   345
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
         Left            =   12105
         Locked          =   -1  'True
         TabIndex        =   78
         TabStop         =   0   'False
         Top             =   1290
         Width           =   1455
      End
      Begin VB.TextBox TxtAdvanceGiven 
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
         Left            =   1770
         MaxLength       =   12
         TabIndex        =   5
         Top             =   1290
         Width           =   1440
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
         Left            =   9480
         Locked          =   -1  'True
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   1290
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
         Left            =   5640
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   1290
         Width           =   1455
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
         Height          =   600
         Left            =   5640
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   1785
         Width           =   7890
      End
      Begin VB.TextBox TxtAgreementID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   375
         Left            =   1770
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   270
         Width           =   1425
      End
      Begin MSComCtl2.DTPicker TxtEntryDate 
         Height          =   375
         Left            =   5640
         TabIndex        =   9
         Top             =   270
         Width           =   1455
         _ExtentX        =   2566
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
         Format          =   40697857
         CurrentDate     =   39884
      End
      Begin MSComCtl2.DTPicker TxtStartDate 
         Height          =   375
         Left            =   1770
         TabIndex        =   2
         Top             =   780
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
         Format          =   40697857
         CurrentDate     =   39884
      End
      Begin MSComCtl2.DTPicker TxtExpiryDate 
         Height          =   375
         Left            =   5640
         TabIndex        =   3
         Top             =   780
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
         Format          =   40697857
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
         ItemData        =   "FrmTxn_LessorAgreement.frx":0052
         Left            =   10200
         List            =   "FrmTxn_LessorAgreement.frx":005C
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   787
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
         Left            =   10200
         Locked          =   -1  'True
         MaxLength       =   200
         ScrollBars      =   2  'Vertical
         TabIndex        =   61
         Top             =   780
         Width           =   3360
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
         ItemData        =   "FrmTxn_LessorAgreement.frx":0076
         Left            =   1740
         List            =   "FrmTxn_LessorAgreement.frx":0083
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   1785
         Width           =   1500
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
         Left            =   1785
         Locked          =   -1  'True
         TabIndex        =   83
         TabStop         =   0   'False
         Top             =   1785
         Width           =   1425
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
         Left            =   10200
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   3360
      End
      Begin VB.TextBox TxtServiceTaxID 
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
         Left            =   10200
         Locked          =   -1  'True
         TabIndex        =   86
         Top             =   240
         Width           =   3360
      End
      Begin MSComCtl2.DTPicker TxtAAggrStartDate 
         Height          =   375
         Left            =   1755
         TabIndex        =   105
         Top             =   2235
         Width           =   1485
         _ExtentX        =   2619
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
         Format          =   40697857
         CurrentDate     =   39884
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid6 
         Height          =   1200
         Left            =   3360
         TabIndex        =   107
         Top             =   360
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
      Begin VB.Label Label26 
         Caption         =   "Actual Agreement Start Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   120
         TabIndex        =   106
         Top             =   2160
         Width           =   1665
      End
      Begin VB.Label Label16 
         Caption         =   "Service Tax Description"
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
         Left            =   7440
         TabIndex        =   85
         Top             =   300
         Width           =   2700
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
         Left            =   135
         TabIndex        =   84
         Top             =   1845
         Width           =   960
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
         Left            =   11100
         TabIndex        =   79
         Top             =   1350
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
         Left            =   7440
         TabIndex        =   44
         Top             =   1350
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
         Left            =   7440
         TabIndex        =   42
         Top             =   847
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
         Left            =   3720
         TabIndex        =   30
         Top             =   840
         Width           =   1035
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Start Date"
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
         TabIndex        =   29
         Top             =   840
         Width           =   885
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
         Left            =   3720
         TabIndex        =   28
         Top             =   1350
         Width           =   1125
      End
      Begin VB.Label LblAdvGiven 
         Caption         =   "Security Deposit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   135
         TabIndex        =   27
         Top             =   1245
         Width           =   1680
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
         Left            =   3720
         TabIndex        =   26
         Top             =   1920
         Width           =   825
      End
      Begin VB.Label LblEntryDate 
         AutoSize        =   -1  'True
         Caption         =   "Entry Date"
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
         Left            =   3720
         TabIndex        =   25
         Top             =   330
         Width           =   930
      End
      Begin VB.Label LblAgreemID 
         AutoSize        =   -1  'True
         Caption         =   "Agreement ID"
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
         TabIndex        =   24
         Top             =   330
         Width           =   1230
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5820
      Left            =   120
      TabIndex        =   45
      Top             =   2880
      Width           =   14985
      _ExtentX        =   26432
      _ExtentY        =   10266
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Lessor Details"
      TabPicture(0)   =   "FrmTxn_LessorAgreement.frx":00A4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "MSHFlexGrid2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Godown Details"
      TabPicture(1)   =   "FrmTxn_LessorAgreement.frx":00C0
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "MSHFlexGrid1"
      Tab(1).Control(1)=   "Frame3"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Agreement Attachment Details"
      TabPicture(2)   =   "FrmTxn_LessorAgreement.frx":00DC
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label25"
      Tab(2).Control(1)=   "lblAttachment"
      Tab(2).Control(2)=   "MSHFlexGrid5"
      Tab(2).Control(3)=   "TxtAttachmentRemarks"
      Tab(2).Control(4)=   "CmdRemoveAttachment"
      Tab(2).Control(5)=   "CmdAttachment"
      Tab(2).Control(6)=   "CmdGetFile"
      Tab(2).ControlCount=   7
      TabCaption(3)   =   "Log Details"
      TabPicture(3)   =   "FrmTxn_LessorAgreement.frx":00F8
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label21"
      Tab(3).Control(1)=   "Label20"
      Tab(3).Control(2)=   "MSHFlexGrid4"
      Tab(3).Control(3)=   "TxtCreated"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "TxtUpdated"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).ControlCount=   5
      Begin VB.CommandButton CmdGetFile 
         Caption         =   "Get File"
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
         Left            =   -67560
         Style           =   1  'Graphical
         TabIndex        =   101
         Top             =   1410
         Width           =   2070
      End
      Begin VB.CommandButton CmdAttachment 
         Caption         =   "Add Attachment"
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
         Left            =   -65490
         Style           =   1  'Graphical
         TabIndex        =   100
         Top             =   1410
         Width           =   2070
      End
      Begin VB.CommandButton CmdRemoveAttachment 
         Caption         =   "Remove Attachment"
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
         Left            =   -63420
         Style           =   1  'Graphical
         TabIndex        =   99
         Top             =   1410
         Width           =   2070
      End
      Begin VB.TextBox TxtAttachmentRemarks 
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
         Height          =   510
         Left            =   -68880
         MaxLength       =   5000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   98
         Top             =   765
         Width           =   7515
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
         Left            =   -72990
         Locked          =   -1  'True
         TabIndex        =   94
         TabStop         =   0   'False
         Top             =   885
         Width           =   7410
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
         Left            =   -72990
         Locked          =   -1  'True
         TabIndex        =   93
         TabStop         =   0   'False
         Top             =   480
         Width           =   7410
      End
      Begin VB.Frame Frame2 
         Height          =   1020
         Left            =   120
         TabIndex        =   56
         Top             =   360
         Width           =   14730
         Begin VB.CommandButton CmdLessorName 
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
            Left            =   4440
            TabIndex        =   57
            TabStop         =   0   'False
            Top             =   510
            Width           =   375
         End
         Begin VB.CommandButton CmdAdd2ListLessor 
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
            Left            =   11295
            TabIndex        =   11
            Top             =   480
            Width           =   1335
         End
         Begin VB.CommandButton CmdDeleteListLessor 
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
            Left            =   12630
            TabIndex        =   12
            Top             =   480
            Width           =   1710
         End
         Begin VB.TextBox TxtLessorAmount 
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
            Left            =   4920
            Locked          =   -1  'True
            MaxLength       =   12
            TabIndex        =   62
            Top             =   480
            Width           =   1890
         End
         Begin VB.TextBox TxtLessorRemark 
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
            Left            =   6885
            MaxLength       =   250
            ScrollBars      =   2  'Vertical
            TabIndex        =   10
            Top             =   480
            Width           =   4305
         End
         Begin VB.ComboBox CmbLessorName 
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
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   8
            Top             =   480
            Width           =   4245
         End
         Begin VB.Label LblAmt 
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
            Height          =   255
            Left            =   5520
            TabIndex        =   60
            Top             =   165
            Width           =   855
         End
         Begin VB.Label Label5 
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
            Left            =   8760
            TabIndex        =   59
            Top             =   165
            Width           =   1200
         End
         Begin VB.Label Label3 
            Caption         =   "Lessor Name"
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
            Left            =   1335
            TabIndex        =   58
            Top             =   150
            Width           =   1530
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1740
         Left            =   -74880
         TabIndex        =   46
         Top             =   360
         Width           =   14715
         Begin VB.TextBox TxtRentAmt 
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
            Left            =   7545
            MaxLength       =   12
            TabIndex        =   18
            Top             =   1267
            Width           =   1890
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
            Left            =   12435
            TabIndex        =   20
            Top             =   1200
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
            Left            =   11100
            TabIndex        =   19
            Top             =   1200
            Width           =   1335
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
            Left            =   105
            Sorted          =   -1  'True
            TabIndex        =   14
            Top             =   525
            Width           =   3285
         End
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
            Left            =   7185
            Sorted          =   -1  'True
            TabIndex        =   16
            Top             =   525
            Width           =   2925
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
            Left            =   3465
            Sorted          =   -1  'True
            TabIndex        =   15
            Top             =   525
            Width           =   3645
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
            Left            =   10200
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   525
            Width           =   4215
         End
         Begin VB.TextBox TxtAreaSqFit 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   375
            Left            =   90
            Locked          =   -1  'True
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   1260
            Width           =   1695
         End
         Begin VB.ComboBox CmbGdLessor 
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
            Left            =   1950
            Sorted          =   -1  'True
            TabIndex        =   17
            Top             =   1267
            Width           =   5265
         End
         Begin VB.Label Label8 
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
            Left            =   705
            TabIndex        =   55
            Top             =   165
            Width           =   1320
         End
         Begin VB.Label Label4 
            Caption         =   "Godown"
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
            Left            =   8145
            TabIndex        =   54
            Top             =   165
            Width           =   1320
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
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
            Height          =   240
            Left            =   5025
            TabIndex        =   53
            Top             =   195
            Width           =   435
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
            Left            =   11505
            TabIndex        =   52
            Top             =   165
            Width           =   765
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Area Sq.Ft"
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
            Left            =   330
            TabIndex        =   51
            Top             =   960
            Width           =   945
         End
         Begin VB.Label Label9 
            Caption         =   "Lessor Name"
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
            TabIndex        =   50
            Top             =   930
            Width           =   1320
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Rent Amount"
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
            Left            =   7845
            TabIndex        =   49
            Top             =   960
            Width           =   1140
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
         Height          =   4215
         Left            =   120
         TabIndex        =   13
         Top             =   1440
         Width           =   14700
         _ExtentX        =   25929
         _ExtentY        =   7435
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   3480
         Left            =   -74880
         TabIndex        =   21
         Top             =   2160
         Width           =   14715
         _ExtentX        =   25956
         _ExtentY        =   6138
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid4 
         Height          =   4215
         Left            =   -74880
         TabIndex        =   95
         Top             =   1440
         Width           =   14700
         _ExtentX        =   25929
         _ExtentY        =   7435
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
         Height          =   3750
         Left            =   -74880
         TabIndex        =   102
         Top             =   1830
         Width           =   13545
         _ExtentX        =   23892
         _ExtentY        =   6615
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
      Begin VB.Label lblAttachment 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Path of  file selected for attachment"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1140
         Left            =   -74880
         TabIndex        =   104
         Top             =   510
         Width           =   5880
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Attachment Remarks"
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
         Left            =   -67680
         TabIndex        =   103
         Top             =   480
         Width           =   2145
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
         Left            =   -74760
         TabIndex        =   97
         Top             =   945
         Width           =   885
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
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
         TabIndex        =   96
         Top             =   540
         Width           =   855
      End
   End
   Begin MSComDlg.CommonDialog cdialog 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmTxn_LessorAgreement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mLocationID, mCMPID, mSCMPID, mLessorID, mGdLessorID, mSMPrefix, Edit_Flg As Variant
Dim mGodownID, mStatus, mServiceTaxID, mSStatus As Variant
Dim mExpiryDate, mSGodownID, mSLessorID, mPrevExpiryDate As Variant
Dim File_Path As String
Dim mMsg As Variant
Dim mAmount, mGodownRentAmount As Double
Dim mGodownExists As Boolean
Dim mGodownMsg As Variant


Private Sub AddCmd_Click()

Edit_Flg = "A"
Call GButtonLock(Me, False)
Frame2.Enabled = True
Frame3.Enabled = True
Frame1.Enabled = True
SSTab1.Tab = 0
TxtAgreementID = ""
TxtEntryDate.Value = Date
TxtStartDate.Value = Date
TxtAAggrStartDate.Value = Date
TxtExpiryDate.Value = Date
TxtAdvanceGiven = ""
TxtAdvanceRemain = ""
TxtAdvAdjusted = ""
TxtRemark = ""
CmbAgreRecAs.Text = ""
TxtAgreeRecAs.Visible = False
CmbAgreRecAs.Visible = True
CmbServiceTaxID.Visible = True
TxtServiceTaxID.Visible = False
CmbServiceTaxID.Text = ""
Check1.Value = vbUnchecked
ChkBRAgreement.Value = vbUnchecked
If LCase(Gen_UserRole) = "fag" Or LCase(Gen_UserRole) = "fag admin" Then
   CmdServiceTaxID.Visible = True
Else
   CmdServiceTaxID.Visible = False
End If

TxtStatus = "Pending"
CmbStatus.Text = TxtStatus

TxtCreated = ""
TxtUpdated = ""
Call ClearLessorFrame
Call ClearGodownFrame

Call Grid_Head1
Call Grid_Head2

lblAttachment.Enabled = True
TxtAttachmentRemarks.Locked = False
CmdAttachment.Enabled = True
CmdGetFile.Enabled = True
CmdRemoveAttachment.Enabled = True

Call Grid_Head4

End Sub
Private Sub Grid_Head1()

MSHFlexGrid2.Clear
MSHFlexGrid2.Rows = 3
MSHFlexGrid2.FixedRows = 2
MSHFlexGrid2.Cols = 4

MSHFlexGrid2.TextMatrix(1, 0) = "Lessor ID"
MSHFlexGrid2.TextMatrix(1, 1) = "Lessor Name"
MSHFlexGrid2.TextMatrix(1, 2) = "Amount"
MSHFlexGrid2.TextMatrix(1, 3) = "Remark"

MSHFlexGrid2.ColWidth(0) = 0
MSHFlexGrid2.ColWidth(1) = 3400
MSHFlexGrid2.ColWidth(2) = 1500
MSHFlexGrid2.ColWidth(3) = 2800

End Sub
Private Sub Grid_Head()

MSHFlexGrid3.Clear
MSHFlexGrid3.Rows = 2
MSHFlexGrid3.FixedRows = 1
MSHFlexGrid3.Cols = 20 '19 '18
MSHFlexGrid3.WordWrap = True
MSHFlexGrid3.TextMatrix(0, 0) = "Lessor Agreement ID"

MSHFlexGrid3.TextMatrix(0, 1) = "Status"

MSHFlexGrid3.TextMatrix(0, 2) = "CMP Name"
MSHFlexGrid3.TextMatrix(0, 3) = "Godown No"
MSHFlexGrid3.TextMatrix(0, 4) = "Lessor Name"

MSHFlexGrid3.TextMatrix(0, 5) = "Entry Date"
MSHFlexGrid3.TextMatrix(0, 6) = "Start Date"
MSHFlexGrid3.TextMatrix(0, 7) = "Expiry Date"

MSHFlexGrid3.TextMatrix(0, 8) = "Rent Amount"


MSHFlexGrid3.TextMatrix(0, 9) = "Agreement Rec.As"
MSHFlexGrid3.TextMatrix(0, 10) = "SD Given"
MSHFlexGrid3.TextMatrix(0, 11) = "SD Adjusted"
MSHFlexGrid3.TextMatrix(0, 12) = "SD Returned"
MSHFlexGrid3.TextMatrix(0, 13) = "Balance"

MSHFlexGrid3.TextMatrix(0, 14) = "Area in Sq. ft."
MSHFlexGrid3.TextMatrix(0, 15) = "Service Tax Description"

MSHFlexGrid3.TextMatrix(0, 16) = "Remark"
MSHFlexGrid3.TextMatrix(0, 17) = "Godown Close Date"
MSHFlexGrid3.TextMatrix(0, 18) = "Agreement Receieve Flag"
MSHFlexGrid3.TextMatrix(0, 19) = "BR Agreement"

MSHFlexGrid3.ColWidth(0) = 800
MSHFlexGrid3.ColWidth(1) = 1200
MSHFlexGrid3.ColWidth(2) = 1200
MSHFlexGrid3.ColWidth(3) = 1000
MSHFlexGrid3.ColWidth(4) = 2000
MSHFlexGrid3.ColWidth(5) = 1400
MSHFlexGrid3.ColWidth(6) = 1400
MSHFlexGrid3.ColWidth(7) = 1400
MSHFlexGrid3.ColWidth(8) = 1400
MSHFlexGrid3.ColWidth(9) = 1400
MSHFlexGrid3.ColWidth(10) = 1200
MSHFlexGrid3.ColWidth(11) = 1200
MSHFlexGrid3.ColWidth(12) = 1200

MSHFlexGrid3.ColWidth(13) = 1200
MSHFlexGrid3.ColWidth(14) = 1400
MSHFlexGrid3.ColWidth(15) = 1800
MSHFlexGrid3.ColWidth(16) = 2000

MSHFlexGrid3.RowHeight(0) = 800

End Sub
Private Sub Grid_Head2()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 3
MSHFlexGrid1.FixedRows = 2
MSHFlexGrid1.Cols = 11

MSHFlexGrid1.TextMatrix(1, 0) = "Godown ID"
MSHFlexGrid1.TextMatrix(1, 1) = "Location"
MSHFlexGrid1.TextMatrix(1, 2) = "CMP Name"
MSHFlexGrid1.TextMatrix(1, 3) = "Godown No"
MSHFlexGrid1.TextMatrix(1, 4) = "Address"
MSHFlexGrid1.TextMatrix(1, 5) = "Area in Sq.Ft"
MSHFlexGrid1.TextMatrix(1, 6) = "Lessor ID "
MSHFlexGrid1.TextMatrix(1, 7) = "Lessor Name"
MSHFlexGrid1.TextMatrix(1, 8) = "Rent Amount"
MSHFlexGrid1.TextMatrix(1, 9) = "SM_Prefix"
MSHFlexGrid1.TextMatrix(1, 10) = "RentPaidUpTo"

MSHFlexGrid1.ColWidth(0) = 0
MSHFlexGrid1.ColWidth(1) = 2200
MSHFlexGrid1.ColWidth(2) = 2200
MSHFlexGrid1.ColWidth(3) = 1300
MSHFlexGrid1.ColWidth(4) = 3000
MSHFlexGrid1.ColWidth(5) = 1300
MSHFlexGrid1.ColWidth(6) = 0
MSHFlexGrid1.ColWidth(7) = 2600
MSHFlexGrid1.ColWidth(8) = 1600
MSHFlexGrid1.ColWidth(9) = 0
MSHFlexGrid1.ColWidth(10) = 1200

End Sub



Private Sub CmbAgreRecAs_LostFocus()
If CmbAgreRecAs.Text = "" Then Exit Sub
If CmbAgreRecAs.Text <> "Hard Copy" And CmbAgreRecAs.Text <> "Soft Copy" Then
   MsgBox "Invalid selection "
   CmbAgreRecAs.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmbGdLessor_GotFocus()
tmp = CmbGdLessor.Text
CmbGdLessor.Clear
If MSHFlexGrid2.Rows < 2 Then Exit Sub

For I = 2 To MSHFlexGrid2.Rows - 2
   CmbGdLessor.AddItem MSHFlexGrid2.TextMatrix(I, 1)
Next I
CmbGdLessor.Text = tmp
End Sub

Private Sub CmbGdLessor_LostFocus()
If CmbGdLessor.Text = "" Then Exit Sub

For I = 1 To MSHFlexGrid2.Rows - 2
   If CmbGdLessor.Text = MSHFlexGrid2.TextMatrix(I, 1) Then
      mGdLessorID = Val(MSHFlexGrid2.TextMatrix(I, 0))
      Exit Sub
   End If
Next I
MsgBox "Invalid selection "
CmbGdLessor.SetFocus
End Sub

Private Sub CmbLessorName_GotFocus()
tmp = CmbLessorName.Text
mLessorID = ""
Call TableMst_Lessor("")
CmbLessorName.Clear
If RsMst_Lessor.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Lessor.RecordCount
    CmbLessorName.AddItem RsMst_Lessor!LessorName
    RsMst_Lessor.MoveNext
Next
CmbLessorName.Text = tmp
End Sub

Private Sub CmbLessorName_LostFocus()
If CmbLessorName.Text = "" Then
   Exit Sub
End If
RsMst_Lessor.MoveFirst
RsMst_Lessor.Find "LessorName = '" & CmbLessorName.Text & "'"
If RsMst_Lessor.EOF Then
   MsgBox "Invalid selection "
   CmbLessorName.SetFocus
   Exit Sub
End If
mLessorID = RsMst_Lessor!LessorID
End Sub

Private Sub CmbLocationID_GotFocus()
tmp = CmbLocationID.Text
CmbCMP.Clear
CmbGodown.Clear
Call TableMst_Location(" Where LocationID <> 0 ")
CmbLocationID.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbLocationID.AddItem RsMst_Location!LocationName
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
If CmbLocationID.Text = "" Then
   MsgBox "Plase Select Location !!!"
   CmbLocationID.SetFocus
   Exit Sub
End If
CmbGodown.Clear
tmp = CmbCMP.Text
Call TableMst_CMP(" Where LocationID = " & mLocationID & "")
CmbCMP.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
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
   CmbGodown.Text = ""
   Exit Sub
End If
RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPName = '" & CmbCMP.Text & "'"
If RsMst_CMP.EOF Then
   MsgBox "Invalid selection "
   CmbCMP.SetFocus
   Exit Sub
End If

If tmp <> RsMst_CMP!CMPName Then
   CmbGodown.Text = ""
End If

mCMPID = RsMst_CMP!CMPID
End Sub

Private Sub CmbGodown_GotFocus()
If CmbCMP.Text = "" Then
   MsgBox "Select CMP first !!! "
   CmbCMP.SetFocus
   Exit Sub
End If
tmp = CmbGodown.Text
Call TableMst_Godown("Where CMPID = " & mCMPID & " And UnitTypeID = 1")
CmbGodown.Clear
If RsMst_Godown.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Godown.RecordCount
    CmbGodown.AddItem RsMst_Godown!GodownNo
    RsMst_Godown.MoveNext
Next
CmbGodown.Text = tmp
End Sub

Private Sub CmbGodown_LostFocus()

If CmbGodown.Text = "" Then
   Exit Sub
End If
RsMst_Godown.MoveFirst
RsMst_Godown.Find "GodownNo = '" & CmbGodown.Text & "'"
If RsMst_Godown.EOF Then
   MsgBox "Invalid selection "
   CmbGodown.SetFocus
   Exit Sub
End If
mGodownID = RsMst_Godown!GodownID
TxtGodownAddress.Text = IIf(IsNull(RsMst_Godown!Address), "", RsMst_Godown!Address)
TxtAreaSqFit.Text = IIf(IsNull(RsMst_Godown!Area), "", RsMst_Godown!Area)
End Sub
Private Sub ClearLessorFrame()
CmbLessorName.Text = ""
TxtLessorAmount = ""
TxtLessorRemark = ""
 CmdAdd2ListLessor.Caption = "Add to List"

End Sub
Private Sub ClearGodownFrame()
CmbLocationID.Text = ""
CmbCMP.Text = ""
CmbGodown.Text = ""
TxtGodownAddress = ""
TxtAreaSqFit = ""
CmbGdLessor.Text = ""
TxtRentAmt = ""
CmbServiceTaxID.Text = ""
CmdAdd2ListGodown.Caption = "Add to List"
End Sub

Private Sub CmbSCmp_GotFocus()
tmp = CmbSCmp.Text
Call TableMst_CMP("")
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
If tmp <> CmbSCmp.Text Then
   CmbSGodown.Clear
End If
mSCMPID = RsMst_CMP!CMPID
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
If CmbServiceTaxID.Text = "" Then Exit Sub

RsMst_ServiceTax.MoveFirst
RsMst_ServiceTax.Find "Description = '" & CmbServiceTaxID.Text & "'"
If RsMst_ServiceTax.EOF Then
   MsgBox "Invalid selection "
   CmbServiceTaxID.SetFocus
   Exit Sub
End If
mServiceTaxID = RsMst_ServiceTax!ServiceTaxID

End Sub

Private Sub CmbSGodown_GotFocus()

If CmbSCmp.Text = "" Then
   MsgBox "Select CMP first !!! "
   CmbSCmp.SetFocus
   Exit Sub
End If
tmp = CmbSGodown.Text
Call TableMst_Godown("Where CMPID = " & mSCMPID & " And UnitTypeID = 1")
CmbSGodown.Clear
If RsMst_Godown.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Godown.RecordCount
    CmbSGodown.AddItem RsMst_Godown!GodownNo
    RsMst_Godown.MoveNext
Next
CmbSGodown.Text = tmp
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
mSGodownID = RsMst_Godown!GodownID
End Sub

Private Sub CmbSLessor_GotFocus()
tmp = CmbSLessor.Text
mLessorID = ""
Call TableMst_Lessor("")
CmbSLessor.Clear
If RsMst_Lessor.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Lessor.RecordCount
    CmbSLessor.AddItem RsMst_Lessor!LessorName
    RsMst_Lessor.MoveNext
Next
CmbSLessor.Text = tmp
End Sub

Private Sub CmbSLessor_LostFocus()
If CmbSLessor.Text = "" Then
   Exit Sub
End If
RsMst_Lessor.MoveFirst
RsMst_Lessor.Find "LessorName = '" & CmbSLessor.Text & "'"
If RsMst_Lessor.EOF Then
   MsgBox "Invalid selection "
   CmbSLessor.SetFocus
   Exit Sub
End If
mSLessorID = RsMst_Lessor!LessorID
End Sub


Private Sub CmbSStatus_LostFocus()
If CmbSStatus.Text = "" Then
   Exit Sub
End If
If LCase(CmbSStatus.Text) <> "approved" And LCase(CmbSStatus.Text) <> "pending" And LCase(CmbSStatus.Text) <> "rejected" Then
   MsgBox "Invalid selection "
   CmbSStatus.SetFocus
   Exit Sub
End If
If LCase(CmbSStatus.Text) = "approved" Then
   mSStatus = "A"
ElseIf LCase(CmbSStatus.Text) = "rejected" Then
   mSStatus = "R"
ElseIf LCase(CmbSStatus.Text) = "pending" Then
   mSStatus = "P"
End If
End Sub



Private Sub CmbStatus_LostFocus()
If CmbStatus.Text = "" Then
   Exit Sub
End If
If LCase(CmbStatus.Text) <> "approved" And LCase(CmbStatus.Text) <> "pending" And LCase(CmbStatus.Text) <> "rejected" Then
   MsgBox "Invalid selection "
   CmbStatus.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmdAdd2ListGodown_Click()
If CmbLocationID.Text = "" Then
   MsgBox "Select Location Name"
   CmbLocationID.SetFocus
   Exit Sub
End If
If CmbCMP.Text = "" Then
   MsgBox "Select CMP Name"
   CmbCMP.SetFocus
   Exit Sub
End If

If CmbGodown.Text = "" Then
   MsgBox "Select Godown"
   CmbGodown.SetFocus
   Exit Sub
End If

If CmbGdLessor.Text = "" Then
   MsgBox "Select Lessor Name"
   CmbGdLessor.SetFocus
   Exit Sub
End If

If TxtRentAmt.Text = "" Then
   MsgBox "Blank Rent Amount not Allowed."
   TxtRentAmt.SetFocus
   Exit Sub
End If



If CheckDuplicateGodown = True Then
   MsgBox "Duplicate Godown Not Allowed."
   CmbGodown.SetFocus
   Exit Sub
End If

If CheckGodownExpiry(mGodownID, "E") = False Then
   MsgBox "Agreemet of this Godown already Exist And Expire on " & mExpiryDate & " ."
   CmbGodown.SetFocus
   Exit Sub
End If
'If CheckGodownExpiry(mGodownID, "S") = False Then
'   If TxtStartDate <= mExpiryDate Then
'      MsgBox "Agreemet of this Godown already Exist And Expire on " & mExpiryDate & " ."
'      CmbGodown.SetFocus
'      Exit Sub
'   End If
'End If
If CmdAdd2ListGodown.Caption = "Update List" Then
   r = MSHFlexGrid1.RowSel
   Call SubtractLessorAmount(MSHFlexGrid1.TextMatrix(r, 8), MSHFlexGrid1.TextMatrix(r, 6))
Else
   r = MSHFlexGrid1.Rows - 1
   MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
End If

MSHFlexGrid1.TextMatrix(r, 0) = mGodownID
MSHFlexGrid1.TextMatrix(r, 1) = CmbLocationID.Text
MSHFlexGrid1.TextMatrix(r, 2) = CmbCMP.Text
MSHFlexGrid1.TextMatrix(r, 3) = CmbGodown.Text
MSHFlexGrid1.TextMatrix(r, 4) = TxtGodownAddress.Text
MSHFlexGrid1.TextMatrix(r, 5) = TxtAreaSqFit.Text
MSHFlexGrid1.TextMatrix(r, 6) = mGdLessorID
MSHFlexGrid1.TextMatrix(r, 7) = CmbGdLessor.Text
MSHFlexGrid1.TextMatrix(r, 8) = Format(Val(TxtRentAmt.Text), "#####0.00")
MSHFlexGrid1.TextMatrix(r, 9) = mSMPrefix

Call AddLessorAmount(MSHFlexGrid1.TextMatrix(r, 8), MSHFlexGrid1.TextMatrix(r, 6))
Call ClearGodownFrame
End Sub
Private Sub AddLessorAmount(mAmt As Variant, mLessorID_ As Variant)
For I = 2 To MSHFlexGrid2.Rows - 2
    If Val(MSHFlexGrid2.TextMatrix(I, 0)) = Val(mLessorID_) Then
       MSHFlexGrid2.TextMatrix(I, 2) = Format(Val(MSHFlexGrid2.TextMatrix(I, 2)) + Val(mAmt), "#####0.00")
       MSHFlexGrid2.TextMatrix(0, 2) = Format(Val(MSHFlexGrid2.TextMatrix(0, 2)) + Val(mAmt), "#####0.00")
       MSHFlexGrid1.TextMatrix(0, 8) = Format(Val(MSHFlexGrid1.TextMatrix(0, 8)) + Val(mAmt), "#####0.00")
       Exit For
    End If
Next
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
                If LCase(Gen_UserRole) <> "head cmg" And LCase(Gen_UserRole) <> "power" Then
                   mExpiryDate = TmpRs!ExpiryDate
                   Retval = False
                   Exit For
                End If
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
Private Sub SubtractLessorAmount(mAmt As Variant, mLessorID_ As Variant)
For I = 2 To MSHFlexGrid2.Rows - 2
    If Val(MSHFlexGrid2.TextMatrix(I, 0)) = Val(mLessorID_) Then
       MSHFlexGrid2.TextMatrix(I, 2) = Format(Val(MSHFlexGrid2.TextMatrix(I, 2)) - Val(mAmt), "#####0.00")
       MSHFlexGrid2.TextMatrix(0, 2) = Format(Val(MSHFlexGrid2.TextMatrix(0, 2)) - Val(mAmt), "#####0.00")
       MSHFlexGrid1.TextMatrix(0, 8) = Format(Val(MSHFlexGrid1.TextMatrix(0, 8)) - Val(mAmt), "#####0.00")
       Exit For
    End If
Next
End Sub

Private Sub CmdAdd2ListLessor_Click()

If CmbLessorName.Text = "" Then
   MsgBox "Select Lessor Name"
   CmbLessorName.SetFocus
   Exit Sub
End If

If CheckDuplicateLessorName = True Then
   MsgBox "Duplicate Lessor Name Not Allowed."
   CmbLessorName.SetFocus
   Exit Sub
End If
If CmdAdd2ListLessor.Caption = "Update List" Then
   r = MSHFlexGrid2.RowSel
Else
   r = MSHFlexGrid2.Rows - 1
   MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
End If

MSHFlexGrid2.TextMatrix(r, 0) = mLessorID
MSHFlexGrid2.TextMatrix(r, 1) = CmbLessorName.Text
MSHFlexGrid2.TextMatrix(r, 2) = Format(Val(TxtLessorAmount.Text), "#####0.00")
MSHFlexGrid2.TextMatrix(r, 3) = TxtLessorRemark.Text
Call ClearLessorFrame
End Sub
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
Retval = False
If MSHFlexGrid1.Rows > 1 Then
   For u = 1 To MSHFlexGrid1.Rows - 1
       If MSHFlexGrid1.TextMatrix(u, 0) = mGodownID And MSHFlexGrid1.TextMatrix(u, 7) = CmbGdLessor.Text Then
          If LCase(CmdAdd2ListGodown.Caption) = "add to list" Then
             Retval = True
             Exit For
          Else
             If u <> MSHFlexGrid1.RowSel Then
                Retval = True
             End If
          End If
       End If
   Next
End If
CheckDuplicateGodown = Retval
End Function
Private Sub CmdDeleteFrmListGodown_Click()
I = MSHFlexGrid1.RowSel
If I > 0 Then
   If MSHFlexGrid1.TextMatrix(I, 1) <> "" Then
      Call SubtractLessorAmount(MSHFlexGrid1.TextMatrix(I, 8), MSHFlexGrid1.TextMatrix(I, 6))
      For RO = I To MSHFlexGrid1.Rows - 2
          For C = 0 To MSHFlexGrid1.Cols - 1
              MSHFlexGrid1.TextMatrix(RO, C) = MSHFlexGrid1.TextMatrix(RO + 1, C)
          Next
      Next
      MSHFlexGrid1.Rows = MSHFlexGrid1.Rows - 1
   End If
End If

Call ClearGodownFrame
End Sub

Private Sub CmdDeleteListLessor_Click()
I = MSHFlexGrid2.RowSel
If I > 0 Then
   If MSHFlexGrid2.TextMatrix(I, 1) <> "" Then
      If CheckLessorInGodown(MSHFlexGrid2.TextMatrix(I, 0)) = True Then
         MsgBox " Lessor Name Used in Godown Details!!"
         Exit Sub
      End If
      For RO = I To MSHFlexGrid2.Rows - 2
          For C = 0 To MSHFlexGrid2.Cols - 1
              MSHFlexGrid2.TextMatrix(RO, C) = MSHFlexGrid2.TextMatrix(RO + 1, C)
          Next
      Next
      MSHFlexGrid2.Rows = MSHFlexGrid2.Rows - 1
   End If
End If
Call ClearLessorFrame
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
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Txn_LessorAgreement.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Txn_LessorAgreement.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid3.Rows - 1
    mRow = I + 1 ' + 6
    For C = 0 To MSHFlexGrid3.Cols - 1
        If C = 3 Then
           oWS.Cells(mRow, C + 1) = "'" & MSHFlexGrid3.TextMatrix(I, C)
        Else
            If InStr(MSHFlexGrid3.TextMatrix(I, C), "/") > 0 Then
               If IsDate(MSHFlexGrid3.TextMatrix(I, C)) Then
                  oWS.Cells(mRow, C + 1) = Format(MSHFlexGrid3.TextMatrix(I, C), "MM/dd/yyyy")
               Else
                  oWS.Cells(mRow, C + 1) = MSHFlexGrid3.TextMatrix(I, C)
               End If
            Else
               oWS.Cells(mRow, C + 1) = MSHFlexGrid3.TextMatrix(I, C)
            End If
        End If
    Next
Next
oWB.Save
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub




Private Sub CmdLessorName_Click()
FrmMst_Lessor.Show
End Sub

Private Sub CmdSearch_Click()

Dim Wc1, mSerTax, mStatus As Variant
Call Grid_Head

If SFrom.Value > STo.Value Then
   MsgBox "From Date Must be less than To Date!!!"
   SFrom.SetFocus
   Exit Sub
End If

tmp = "Select TL.LAID,MC.CMPName, MG.GodownNO, ML.LessorName,TL.EntryDate,TL.StartDate,TL.ExpiryDate,"
tmp = tmp & " TL.AgreementRecvAs,TGD.RentAmount,MG.Area,TL.Flag,TL.AdvanceGiven ,TL.AdvanceAdjusted,"
tmp = tmp & " TL.AdvanceReturn, TL.AdvanceBalance, TL.Remark, TGD.GodownID, TGD.LessorID,"
tmp = tmp & " TL.ServiceTaxID,MG.CloseDate,TL.AgreenotReceieveFlag,TL.BRAgreement "
tmp = tmp & " From Txn_LessorAgreement TL "
tmp = tmp & " Inner join Txn_GodownLessorAgreeDetail TGD on TL.LAID=TGD.LAID"
tmp = tmp & " Inner join Mst_Lessor ML on TGD.LessorID=ML.LessorID"
tmp = tmp & " Inner join Mst_Godown MG on TGD.GodownID=MG.GodownID"
tmp = tmp & " Inner Join Mst_CMP MC on MG.CMPID=MC.CMPID"


If Gen_DBType = "MDB" Then
   Wc1 = " Where TL.EntryDate Between #" & Format(SFrom.Value, Gen_DatFormat) & "# And #" & Format(STo.Value, Gen_DatFormat) & "# "
Else
   Wc1 = " Where TL.EntryDate Between '" & Format(SFrom.Value, Gen_DatFormat) & "' And '" & Format(STo.Value, Gen_DatFormat) & "' "
End If
If CmbSAgreRecAs.Text <> "" Then
   Wc1 = Wc1 & " And TL.AgreementRecvAs='" & CmbSAgreRecAs.Text & "'"
End If

If CmbSCmp.Text <> "" Then
   Wc1 = Wc1 & " And MC.CMPName='" & CmbSCmp.Text & "'"
Else
   If Gen_WcCMP <> "" Then
      Wc1 = Wc1 & " And MC." & Gen_WcCMP
   End If
End If

If CmbSGodown.Text <> "" Then
   Wc1 = Wc1 & " And MG.GodownNo='" & CmbSGodown.Text & "'"
End If

If CmbSLessor.Text <> "" Then
   Wc1 = Wc1 & " And ML.LessorName='" & CmbSLessor.Text & "'"
End If

If CmbSStatus.Text <> "" Then
   Wc1 = Wc1 & " And TL.Flag='" & mSStatus & "'"
End If

If TxtSAgreeID.Text <> "" Then
   Wc1 = Wc1 & " And TL.LAID=" & Val(TxtSAgreeID) & ""
End If


tmp = tmp & Wc1 & " Order By TL.LAID"
'
'tmp = "Select LAID,EntryDate,StartDate,ExpiryDate,AgreementRecvAs,"
'tmp = tmp & " AdvanceGiven,AdvanceAdjusted,AdvanceReturn,AdvanceBalance,Remark From Txn_LessorAgreement"



Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   MSHFlexGrid3.Rows = TmpRs1.RecordCount + 2
   For I = 1 To TmpRs1.RecordCount
   
       MSHFlexGrid3.TextMatrix(I, 0) = TmpRs1!LAID
       If LCase(TmpRs1!Flag) = "a" Then
          mStatus = "Approved"
       ElseIf LCase(TmpRs1!Flag) = "p" Then
          mStatus = "Pending"
       ElseIf LCase(TmpRs1!Flag) = "r" Then
          mStatus = "Rejected"
       End If
       
       MSHFlexGrid3.TextMatrix(I, 1) = mStatus
       
       MSHFlexGrid3.TextMatrix(I, 2) = TmpRs1!CMPName
       MSHFlexGrid3.TextMatrix(I, 3) = TmpRs1!GodownNo
       MSHFlexGrid3.TextMatrix(I, 4) = TmpRs1!LessorName
       
       MSHFlexGrid3.TextMatrix(I, 5) = Format(TmpRs1!EntryDate, "dd-MMM-yyyy")
       MSHFlexGrid3.TextMatrix(I, 6) = Format(TmpRs1!StartDate, "dd-MMM-yyyy")
       MSHFlexGrid3.TextMatrix(I, 7) = Format(TmpRs1!ExpiryDate, "dd-MMM-yyyy")
       
       MSHFlexGrid3.TextMatrix(I, 8) = Format(TmpRs1!RentAmount, "#######0.00")
       
       MSHFlexGrid3.TextMatrix(I, 9) = Format(TmpRs1!AgreementRecvAs, "#######0.00")
       MSHFlexGrid3.TextMatrix(I, 10) = Format(TmpRs1!AdvanceGiven, "#######0.00")
       MSHFlexGrid3.TextMatrix(I, 11) = Format(TmpRs1!AdvanceAdjusted, "#######0.00")
       MSHFlexGrid3.TextMatrix(I, 12) = Format(TmpRs1!AdvanceReturn, "#######0.00")
       MSHFlexGrid3.TextMatrix(I, 13) = Format(TmpRs1!AdvanceBalance, "#######0.00")
       
       MSHFlexGrid3.TextMatrix(I, 14) = TmpRs1!Area
       mSerTax = IIf(IsNull(TmpRs1!ServiceTaxID), 0, TmpRs1!ServiceTaxID)
       If mSerTax <> 0 Then
          MSHFlexGrid3.TextMatrix(I, 15) = GetServiceTaxDescription(mSerTax)
       End If
       MSHFlexGrid3.TextMatrix(I, 16) = TmpRs1!Remark
       
       MSHFlexGrid3.TextMatrix(I, 17) = IIf(IsNull(TmpRs1!CloseDate), "", Format(TmpRs1!CloseDate, "dd-mmm-yyyy"))
       MSHFlexGrid3.TextMatrix(I, 18) = IIf(IsNull(TmpRs1!AgreenotReceieveFlag), "", TmpRs1!AgreenotReceieveFlag)
       If Val(MSHFlexGrid3.TextMatrix(I, 18)) = 0 Then
          MSHFlexGrid3.TextMatrix(I, 18) = "Yes"
       Else
          MSHFlexGrid3.TextMatrix(I, 18) = "No"
       End If
       MSHFlexGrid3.TextMatrix(I, 19) = IIf(IsNull(TmpRs1!BRAgreement), False, TmpRs1!BRAgreement)
       If MSHFlexGrid3.TextMatrix(I, 19) = "True" Then
          MSHFlexGrid3.TextMatrix(I, 19) = "Yes"
       Else
          MSHFlexGrid3.TextMatrix(I, 19) = "No"
       End If
        

'''       For c = 0 To MSHFlexGrid3.Cols - 1

'''           MSHFlexGrid3.TextMatrix(i, c) = IIf(IsNull(TmpRs1.Fields(c)), "", TmpRs1.Fields(c))
'''       Next
       TmpRs1.MoveNext
   Next
   
End If
Wc1 = "EntryDate >='" & Format(SFrom.Value, Gen_DatFormat) & "' And EntryDate <='" & Format(STo.Value, Gen_DatFormat) & "'"

If CmbSAgreRecAs.Text <> "" Then
   Wc1 = Wc1 & " And AgreementRecvAs='" & CmbSAgreRecAs.Text & "'"
End If
RsTxn_LessorAgreement.Filter = Wc1
Label22.Caption = TmpRs1.RecordCount
Label22.Refresh
End Sub

Private Sub CmdServiceTaxID_Click()
FrmMst_ServiceTax.Show
End Sub

Private Sub DeleteCmd_Click()
Dim ans As Variant
On Error GoTo msgError
If LCase(TxtStatus) = "approved" Or LCase(TxtStatus) = "rejected" Then
   MsgBox "This agreement is " & TxtStatus & ". You can't delete this agreement!!!"
   Exit Sub
End If
ans = MsgBox("Do you really want to DELETE this record???", vbYesNo)
If ans = vbYes Then

   Call DeleteChildTableDetail
   
   RsTxn_LessorAgreement.Delete
   RsTxn_LessorAgreement.Update
   If RsTxn_LessorAgreement.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsTxn_LessorAgreement.MoveNext
   If RsTxn_LessorAgreement.EOF() Then
      RsTxn_LessorAgreement.MoveLast
   End If
   Call Indata
End If
Exit Sub
msgError:
MsgBox "Child record Present "
RsTxn_LessorAgreement.CancelUpdate

End Sub

Private Sub EditCmd_Click()

'And LCase(Gen_UserRole) <> "superadmin"

If LCase(Gen_UserRole) <> "fag" And LCase(Gen_UserRole) <> "fag admin" And LCase(Gen_UserRole) <> "power" Then
   If LCase(TxtStatus) = "approved" Or LCase(TxtStatus) = "rejected" Then
      MsgBox "This agreement is " & TxtStatus & ". You can't edit this agreement any more!!!"
      Exit Sub
   End If
End If

If CheckNextAgreement = True Then
   MsgBox "You Can't Edit this Record. New Agreement is created for one or more godowns"
   Exit Sub
End If

Edit_Flg = "E"

Call GButtonLock(Me, False)


Frame1.Enabled = True
CmbStatus.Text = TxtStatus.Text

'Commented on 09 July 2013 after mail for Mayur Sir and Sankar Sir
tmp = "Select a.LAID,Max(b.ExpenseFrom) 'RentMonth' from Txn_GodownRentDetail a"
tmp = tmp & " Inner Join Txn_GodownRent b On a.ClaimID  = b.ClaimID"
tmp = tmp & " Where b.Flag in ('P','A') And a.LAID = " & Val(TxtAgreementID) & ""
tmp = tmp & " Group By a.LAID"
tmp = tmp & " Order By LAID"

Call TmpTable

If TmpRs.RecordCount > 0 Then
   Frame2.Enabled = False
   Frame3.Enabled = False
   TxtAdvanceGiven.Locked = True
'   TxtStartDate.Enabled = False
'   TxtAAggrStartDate.Enabled = False
   TxtRemark.Locked = False
   CmbStatus.Text = TxtStatus.Text
   CmbServiceTaxID.Text = TxtServiceTaxID.Text
   If TxtServiceTaxID = "" Then
      CmbServiceTaxID.Visible = True
      TxtServiceTaxID.Visible = False
      CmbServiceTaxID.Text = TxtServiceTaxID.Text
   End If
   
Else
   Frame2.Enabled = True
   Frame3.Enabled = True
   TxtAdvanceGiven.Locked = False
'   TxtStartDate.Enabled = False
'   TxtAAggrStartDate.Enabled = False
   TxtRemark.Locked = False
   If LCase(Gen_UserRole) = "fag" Or LCase(Gen_UserRole) = "power" Or LCase(Gen_UserRole) = "fag admin" Then
      CmbStatus.Visible = True
      TxtStatus.Visible = False
      CmbStatus.Text = TxtStatus.Text
      CmdServiceTaxID.Visible = True
   End If
   
   CmbServiceTaxID.Visible = True
   TxtServiceTaxID.Visible = False
   CmbServiceTaxID.Text = TxtServiceTaxID.Text
    
   CmbAgreRecAs.Text = TxtAgreeRecAs
   TxtAgreeRecAs.Visible = False
   CmbAgreRecAs.Visible = True
   
   lblAttachment.Enabled = True
   TxtAttachmentRemarks.Locked = False
   CmdAttachment.Enabled = True
   CmdGetFile.Enabled = True
   CmdRemoveAttachment.Enabled = True
   SSTab1.Tab = 0
End If



'Or LCase(Gen_UserRole) = "superadmin"


'mPrevExpiryDate = TxtExpiryDate.Value



End Sub
Private Function CheckNextAgreement() As Boolean
Dim Retval As Boolean
Retval = False
For I = 2 To MSHFlexGrid1.Rows - 1
    If MSHFlexGrid1.TextMatrix(I, 1) <> "" Then
       If CheckGodownExpiry(Val(MSHFlexGrid1.TextMatrix(I, 0)), "") = False Then
          Retval = True
          Exit For
       End If
    End If
Next
CheckNextAgreement = Retval
End Function
Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub FirstCmd_Click()
RsTxn_LessorAgreement.MoveFirst
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
Call Grid_Head
SFrom.Value = Date - 365
STo.Value = Date
Label22.Caption = ""
CmbSAgreRecAs.Text = ""
CmbSCmp.Text = ""
CmbSGodown.Text = ""
CmbSLessor.Text = ""
CmbSStatus.Text = ""
TxtSAgreeID.Text = ""

Call TableTxn_LessorAgreement(" Order by LAID")
If RsTxn_LessorAgreement.RecordCount = 0 Then
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

Public Sub Indata()
Dim mSerTax As Variant
Call ClearFrame
Call ClearGodownFrame
Call ClearLessorFrame
Frame2.Enabled = False
Frame3.Enabled = False
Frame1.Enabled = False
SSTab1.Tab = 0
Call Grid_Head1
Call Grid_Head2

CmbAgreRecAs.Visible = False
TxtAgreeRecAs.Visible = True
CmbStatus.Visible = False
TxtStatus.Visible = True

CmbServiceTaxID.Visible = False
TxtServiceTaxID.Visible = True
CmdServiceTaxID.Visible = False

'TxtStartDate.Enabled = True
'TxtAAggrStartDate.Enabled = True

TxtAgreementID = IIf(IsNull(RsTxn_LessorAgreement!LAID), "", RsTxn_LessorAgreement!LAID)
TxtStartDate.Value = RsTxn_LessorAgreement!StartDate
TxtAAggrStartDate.Value = RsTxn_LessorAgreement!AgreemenActualtStartDate
TxtEntryDate.Value = RsTxn_LessorAgreement!EntryDate
TxtExpiryDate.Value = RsTxn_LessorAgreement!ExpiryDate
mServiceTaxID = IIf(IsNull(RsTxn_LessorAgreement!ServiceTaxID), 0, RsTxn_LessorAgreement!ServiceTaxID)
mSerTax = GetServiceTaxDescription(IIf(IsNull(RsTxn_LessorAgreement!ServiceTaxID), 0, RsTxn_LessorAgreement!ServiceTaxID))
TxtServiceTaxID.Text = IIf(mSerTax = 0, "", mSerTax)

TxtAgreeRecAs = IIf(IsNull(RsTxn_LessorAgreement!AgreementRecvAs), "", RsTxn_LessorAgreement!AgreementRecvAs)
CmbAgreRecAs.Text = TxtAgreeRecAs
TxtAdvanceGiven = Format(IIf(IsNull(RsTxn_LessorAgreement!AdvanceGiven), "", RsTxn_LessorAgreement!AdvanceGiven), "#####0.00")
TxtAdvAdjusted = Format(IIf(IsNull(RsTxn_LessorAgreement!AdvanceAdjusted), "", RsTxn_LessorAgreement!AdvanceAdjusted), "#####0.00")
TxtAdvanceRemain = Format(IIf(IsNull(RsTxn_LessorAgreement!AdvanceReturn), "", RsTxn_LessorAgreement!AdvanceReturn), "#####0.00")
TxtBalance.Text = Format(IIf(IsNull(RsTxn_LessorAgreement!AdvanceBalance), "", RsTxn_LessorAgreement!AdvanceBalance), "#####0.00")
TxtRemark = IIf(IsNull(RsTxn_LessorAgreement!Remark), "", RsTxn_LessorAgreement!Remark)
TxtCreated = IIf(IsNull(RsTxn_LessorAgreement!CreatedBy), "", RsTxn_LessorAgreement!CreatedBy) & " :- " & IIf(IsNull(RsTxn_LessorAgreement!CreatedDate), "", RsTxn_LessorAgreement!CreatedDate)
TxtUpdated = IIf(IsNull(RsTxn_LessorAgreement!UpdatedBy), "", RsTxn_LessorAgreement!UpdatedBy) & " :- " & IIf(IsNull(RsTxn_LessorAgreement!UpdatedDate), "", RsTxn_LessorAgreement!UpdatedDate)

mStatus = IIf(IsNull(RsTxn_LessorAgreement!Flag), "", RsTxn_LessorAgreement!Flag)

Check1.Value = IIf(IsNull(RsTxn_LessorAgreement!AgreenotReceieveFlag), 0, RsTxn_LessorAgreement!AgreenotReceieveFlag)
mBRAgreement = IIf(IsNull(RsTxn_LessorAgreement!BRAgreement), False, RsTxn_LessorAgreement!BRAgreement)
If mBRAgreement = True Then
   ChkBRAgreement.Value = vbChecked
Else
   ChkBRAgreement.Value = vbUnchecked
End If
If LCase(mStatus) = "a" Then
   TxtStatus = "Approved"
ElseIf LCase(mStatus) = "p" Then
   TxtStatus = "Pending"
ElseIf LCase(mStatus) = "r" Then
   TxtStatus = "Rejected"
Else
   TxtStatus = ""
End If

'TxtG_UID = IIf(IsNull(RsTxn_LessorAgreement!G_UID), "", RsTxn_LessorAgreement!G_UID)
Call InDataLessor
Call InDataGodown
Call InDataLogHistory
Call IndataAttachmentDetails
If RsTxn_LessorAgreement.RecordCount = 1 Then
   Call GButtonLock_1(Me, True)
Else
   Call GButtonLock(Me, True)
End If

End Sub

Private Sub InDataLessor()
'Call TableTxn_LessorAgreementDetail(" Where LAID = " & Val(TxtAgreementID) & " ")

tmp = "Select TLD.LAID,TLD.LessorID,TLD.LessorAmt,TLD.Remark,ML.LessorName"
tmp = tmp & " From Txn_LessorAgreementDetail TLD"
tmp = tmp & " inner Join Mst_Lessor ML on TLD.LessorID=ML.LessorID"
tmp = tmp & " Where TLD.LAID=" & Val(TxtAgreementID) & ""
Call Tmp1Table
MSHFlexGrid2.Rows = TmpRs1.RecordCount + 3
For I = 2 To MSHFlexGrid2.Rows - 2
   MSHFlexGrid2.TextMatrix(I, 0) = TmpRs1!LessorID
   MSHFlexGrid2.TextMatrix(I, 1) = TmpRs1!LessorName
   MSHFlexGrid2.TextMatrix(I, 2) = Format(TmpRs1!LessorAmt, "#####0.00")
   MSHFlexGrid2.TextMatrix(I, 3) = IIf(IsNull(TmpRs1!Remark), "", TmpRs1!Remark)
   MSHFlexGrid2.TextMatrix(0, 2) = Format(Val(MSHFlexGrid2.TextMatrix(0, 2)) + TmpRs1!LessorAmt, "#####0.00")
   TmpRs1.MoveNext
Next
End Sub
Private Sub InDataLogHistory()

Call Grid_Head3
tmp = "Select LAID, UpdatedBy, UpdatedOn from Log_LessorAgreementHistory Where LAID=" & Val(TxtAgreementID) & ""
Call TmpTable
MSHFlexGrid4.Rows = MSHFlexGrid4.Rows + TmpRs.RecordCount
For I = 1 To MSHFlexGrid4.Rows - 2
   MSHFlexGrid4.TextMatrix(I, 0) = TmpRs!LAID
   MSHFlexGrid4.TextMatrix(I, 1) = IIf(IsNull(TmpRs!UpdatedBy), "", TmpRs!UpdatedBy)
   MSHFlexGrid4.TextMatrix(I, 2) = TmpRs!UpdatedOn
   TmpRs.MoveNext
Next
End Sub
Private Sub InDataGodown()
',TGD.ServiceTaxID,MS.Description,
'Call TableTxn_GodownLessorAgreeDetail(" Where LAID = " & Val(TxtAgreementID) & " ")
tmp = "select TGD.LAID,TGD.GodownID,TGD.SM_Prefix,TGD.LessorID,TGD.RentAmount,"
tmp = tmp & " MG.GodownNo , MG.Address, MG.Area, MC.CMPName, ML.LocationName, MLes.LessorName,TGD.RentPaidUpTo"
tmp = tmp & " From Txn_GodownLessorAgreeDetail TGD"
tmp = tmp & " Inner Join Mst_Godown MG on TGD.GodownID=MG.GodownID"
tmp = tmp & " Inner Join Mst_CMP MC on MG.CMPID=MC.CMPID"
tmp = tmp & " Inner Join Mst_Location ML on MC.LocationID=ML.LocationID"
tmp = tmp & " inner Join Mst_Lessor MLes on TGD.LessorID=MLes.LessorID"
'tmp = tmp & " Inner Join Mst_ServiceTax MS on TGD.ServiceTaxID=MS.ServiceTaxID"
tmp = tmp & " Where TGD.LAID=" & Val(TxtAgreementID) & ""
tmp = tmp & " Order By TGD.GLADID"

Call TmpTable

MSHFlexGrid1.Rows = TmpRs.RecordCount + 3
For I = 2 To MSHFlexGrid1.Rows - 2
   MSHFlexGrid1.TextMatrix(I, 0) = TmpRs!GodownID
   MSHFlexGrid1.TextMatrix(I, 3) = TmpRs!GodownNo
   MSHFlexGrid1.TextMatrix(I, 5) = TmpRs!Area
   MSHFlexGrid1.TextMatrix(I, 4) = TmpRs!Address
   MSHFlexGrid1.TextMatrix(I, 2) = TmpRs!CMPName
   MSHFlexGrid1.TextMatrix(I, 1) = TmpRs!LocationName
   MSHFlexGrid1.TextMatrix(I, 6) = TmpRs!LessorID
   MSHFlexGrid1.TextMatrix(I, 7) = TmpRs!LessorName
   MSHFlexGrid1.TextMatrix(I, 8) = Format(TmpRs!RentAmount, "#####0.00")
   MSHFlexGrid1.TextMatrix(I, 9) = TmpRs!SM_Prefix
   MSHFlexGrid1.TextMatrix(I, 10) = IIf(IsNull(TmpRs!RentPaidUpTo), "", TmpRs!RentPaidUpTo)
''   MSHFlexGrid1.TextMatrix(i, 11) = TmpRs!ServiceTaxID
   
   MSHFlexGrid1.TextMatrix(0, 8) = Format(Val(MSHFlexGrid1.TextMatrix(0, 8)) + TmpRs!RentAmount, "#####0.00")
   TmpRs.MoveNext
Next
End Sub
Private Sub SaveLessorDetails()
Call TableTxn_LessorAgreementDetail(" Where LAID=" & Val(TxtAgreementID) & "")

For I = 2 To MSHFlexGrid2.Rows - 1
    If MSHFlexGrid2.TextMatrix(I, 1) <> "" Then
       RsTxn_LessorAgreementDetail.AddNew
       RsTxn_LessorAgreementDetail!G_UID = GetGUID
       RsTxn_LessorAgreementDetail!LADetailID = getMaxLADetailID
       RsTxn_LessorAgreementDetail!LessorID = Val(MSHFlexGrid2.TextMatrix(I, 0))
       RsTxn_LessorAgreementDetail!LAID = Val(TxtAgreementID)
       RsTxn_LessorAgreementDetail!LessorAmt = Val(MSHFlexGrid2.TextMatrix(I, 2))
       RsTxn_LessorAgreementDetail!Remark = MSHFlexGrid2.TextMatrix(I, 3)
       RsTxn_LessorAgreementDetail.Update
       RsTxn_LessorAgreementDetail.Requery
    End If
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

Private Sub SaveGodownLessorDetails()

Call TableTxn_GodownLessorAgreeDetail(" Where LAID=" & Val(TxtAgreementID) & "")
tmp = " Select * from Mst_Godown WHere UnitTypeID = 1 "
Call Tmp1Table

For I = 2 To MSHFlexGrid1.Rows - 1
    If MSHFlexGrid1.TextMatrix(I, 1) <> "" Then
       RsTxn_GodownLessorAgreeDetail.AddNew
       RsTxn_GodownLessorAgreeDetail!G_UID = GetGUID
       RsTxn_GodownLessorAgreeDetail!GLADID = getMaxGLADID
       RsTxn_GodownLessorAgreeDetail!LAID = Val(TxtAgreementID)
       RsTxn_GodownLessorAgreeDetail!GodownID = Val(MSHFlexGrid1.TextMatrix(I, 0))
       RsTxn_GodownLessorAgreeDetail!SM_Prefix = MSHFlexGrid1.TextMatrix(I, 9)
       RsTxn_GodownLessorAgreeDetail!LessorID = Val(MSHFlexGrid1.TextMatrix(I, 6))
       RsTxn_GodownLessorAgreeDetail!RentAmount = Val(MSHFlexGrid1.TextMatrix(I, 8))
       If MSHFlexGrid1.TextMatrix(I, 10) <> "" Then
          RsTxn_GodownLessorAgreeDetail!RentPaidUpTo = MSHFlexGrid1.TextMatrix(I, 10)
       End If
       RsTxn_GodownLessorAgreeDetail.Update
       RsTxn_GodownLessorAgreeDetail.Requery
       If TmpRs1.RecordCount > 0 Then
          TmpRs1.MoveFirst
          TmpRs1.Find "GodownID = " & Val(MSHFlexGrid1.TextMatrix(I, 0)) & ""
          If IsNull(TmpRs1!AgreementEndDate) = False Then
             If TmpRs1!AgreementEndDate < TxtExpiryDate.Value Then
                TmpRs1!AgreementEndDate = TxtExpiryDate.Value
                TmpRs1.Update
                TmpRs1.Requery
             End If
          Else
             TmpRs1!AgreementEndDate = TxtExpiryDate.Value
             TmpRs1.Update
             TmpRs1.Requery
          End If
       End If
    End If
Next

End Sub




Private Sub LastCmd_Click()
RsTxn_LessorAgreement.MoveLast
Call Indata
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False
End Sub

Private Sub MSHFlexGrid1_Click()
If SaveCmd.Enabled = False Then Exit Sub
Dim r As Variant

If Frame3.Enabled = False Then Exit Sub

r = MSHFlexGrid1.RowSel
If r < 2 Or r = MSHFlexGrid1.Rows - 1 Then
   Call ClearGodownFrame
   Exit Sub
End If
If MSHFlexGrid1.TextMatrix(r, 1) = "" Then Exit Sub

mGodownID = MSHFlexGrid1.TextMatrix(r, 0)
CmbLocationID.Text = MSHFlexGrid1.TextMatrix(r, 1)
mLocationID = GetLocationID(CmbLocationID.Text)
CmbCMP.Text = MSHFlexGrid1.TextMatrix(r, 2)
mCMPID = GetCMPID(CmbCMP.Text)
CmbGodown.Text = MSHFlexGrid1.TextMatrix(r, 3)
TxtGodownAddress.Text = MSHFlexGrid1.TextMatrix(r, 4)
TxtAreaSqFit.Text = MSHFlexGrid1.TextMatrix(r, 5)
mGdLessorID = MSHFlexGrid1.TextMatrix(r, 6)
CmbGdLessor.Text = MSHFlexGrid1.TextMatrix(r, 7)
TxtRentAmt.Text = Val(MSHFlexGrid1.TextMatrix(r, 8))


CmdAdd2ListGodown.Caption = "Update List"
CmdAdd2ListGodown.SetFocus
End Sub

Private Sub MSHFlexGrid2_Click()
If SaveCmd.Enabled = False Then Exit Sub
Dim r As Variant
r = MSHFlexGrid2.RowSel

If Frame2.Enabled = False Then Exit Sub

If r < 2 Or r = MSHFlexGrid2.Rows - 1 Then
   Call ClearLessorFrame
   Exit Sub
End If
If MSHFlexGrid2.TextMatrix(r, 1) = "" Then
   Exit Sub
End If

CmbLessorName.Text = MSHFlexGrid2.TextMatrix(r, 1)
mLessorID = MSHFlexGrid2.TextMatrix(r, 0)
TxtLessorAmount = Val(MSHFlexGrid2.TextMatrix(r, 2))
TxtLessorRemark = MSHFlexGrid2.TextMatrix(r, 3)
CmdAdd2ListLessor.Caption = "Update List"
CmbLessorName.SetFocus
End Sub

Private Sub MSHFlexGrid3_Click()
Dim I As Variant
I = MSHFlexGrid3.RowSel
If I <= 0 Then Exit Sub
If MSHFlexGrid3.TextMatrix(I, 0) = "" Then Exit Sub
RsTxn_LessorAgreement.MoveFirst
RsTxn_LessorAgreement.Find "LAID= " & Val(MSHFlexGrid3.TextMatrix(I, 0)) & ""
If Not RsTxn_LessorAgreement.EOF Then
   Call Indata
   Frame0.Visible = False
End If
End Sub

Private Sub MSHFlexGrid5_Click()
Dim ans, r As Variant
Dim ID As Integer
r = MSHFlexGrid5.RowSel
If r < 1 Then Exit Sub
If MSHFlexGrid5.TextMatrix(r, 1) = "" Then Exit Sub

If SaveCmd.Enabled = False Then
'Agreement Attachment\Lessor
   ans = MsgBox("Do You want to view this document ", vbYesNo)
   If ans = vbYes Then
      If Right(App.Path, 1) = "\" Then
         Shell "Explorer.exe " & App.Path & "Agreement Attachment\Lessor", vbMaximizedFocus
      Else
         Shell "Explorer.exe " & App.Path & "Agreement Attachment\Lessor\", vbMaximizedFocus
      End If
   End If
   Exit Sub
End If

lblAttachment.Caption = MSHFlexGrid5.TextMatrix(r, 3)
TxtAttachmentRemarks = MSHFlexGrid5.TextMatrix(r, 2)

CmdRemoveAttachment.Enabled = True
If MSHFlexGrid5.TextMatrix(r, 4) <> "" Then
   CmdRemoveAttachment.Enabled = False
End If

CmdAttachment.Caption = "Update Attachment"
End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsTxn_LessorAgreement.MoveNext
If RsTxn_LessorAgreement.EOF Then
   RsTxn_LessorAgreement.MoveLast
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
RsTxn_LessorAgreement.MovePrevious
If RsTxn_LessorAgreement.BOF Then
   RsTxn_LessorAgreement.MoveFirst
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

'If TxtEntryDate < Date Then
'   MsgBox ""
'   CmbAgreRecAs.SetFocus
'   Exit Sub
'End If

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
If CheckGodownsExpiryDate = False Then
   MsgBox "Previous Expiry Date of one or more selected Godown is Greater than Start Date."
   TxtStartDate.SetFocus
   Exit Sub
End If

If CmbAgreRecAs.Text = "" Then
   MsgBox "Select Agreement Received As."
   CmbAgreRecAs.SetFocus
   Exit Sub
End If

If TxtAdvanceGiven.Text = "" Then
   MsgBox "Blank Advance given Not Allowed."
   TxtAdvanceGiven.SetFocus
   Exit Sub
End If

If CmbServiceTaxID.Text = "" Then
   MsgBox "Blank Service Tax description not allowed!!!"
   CmbServiceTaxID.SetFocus
   Exit Sub
End If

If MSHFlexGrid1.Rows < 4 Then
   MsgBox "Godown Details Missing."
   Exit Sub
End If

If MSHFlexGrid2.Rows < 4 Then
   MsgBox "Lessor Details Missing."
   Exit Sub
End If


If CmbStatus.Visible = True Then
   If LCase(CmbStatus.Text) = "approved" Then
      If Format(Gen_BlockingDate, "yyyy-mm") >= Format(TxtStartDate.Value, "yyyy-mm") Then
         If LCase(Gen_UserRole) <> "fag admin" And LCase(Gen_UserRole) <> "power" Then
            MsgBox "Agreement Start Date allowed for next month of " & Format(Gen_BlockingDate, "MMM-yyyy") & "!!! "
            TxtStartDate.SetFocus
            Exit Sub
         End If
      End If
      mStatus = "A"
   ElseIf LCase(CmbStatus.Text) = "rejected" Then
      mStatus = "R"
   ElseIf LCase(CmbStatus.Text) = "pending" Then
      mStatus = "P"
   End If
Else
   mStatus = Left(CmbStatus.Text, 1) '"P"
End If

mMsg = ""
mGodownMsg = ""

Call GetGodwn

If mGodownMsg <> "" Then
   MsgBox mGodownMsg
   Exit Sub
End If


If mMsg <> "" Then
   MsgBox mMsg
   Exit Sub
End If


If Edit_Flg = "A" Then
   RsTxn_LessorAgreement.AddNew
   RsTxn_LessorAgreement!LAID = getMaxLAID
   RsTxn_LessorAgreement!G_UID = GetGUID
   TxtAgreementID.Text = RsTxn_LessorAgreement!LAID
Else
'   If TxtExpiryDate.Value < mPrevExpiryDate Then
'      MsgBox "Expiry Date should be greater than previous Expiry Date"
'      TxtExpiryDate.SetFocus
'      Exit Sub
'   End If
   If CheckRaintPaidUpto = False Then
      MsgBox "Rent Paid Upto Date of one or more selected Godown is Greater than Expiry Date."
      TxtExpiryDate.SetFocus
      Exit Sub
   End If

   RsTxn_LessorAgreement.MoveFirst
   RsTxn_LessorAgreement.Find " LAID= '" & Val(TxtAgreementID.Text) & "'"
   Call DeleteChildTableDetail

End If

RsTxn_LessorAgreement!LAID = Val(TxtAgreementID.Text)
RsTxn_LessorAgreement!EntryDate = Format(TxtEntryDate.Value, Gen_DatFormat)
RsTxn_LessorAgreement!StartDate = Format(TxtStartDate.Value, Gen_DatFormat)
RsTxn_LessorAgreement!AgreemenActualtStartDate = Format(TxtAAggrStartDate.Value, Gen_DatFormat)
RsTxn_LessorAgreement!ExpiryDate = Format(TxtExpiryDate.Value, Gen_DatFormat)
RsTxn_LessorAgreement!ServiceTaxID = mServiceTaxID
RsTxn_LessorAgreement!AgreementRecvAs = CmbAgreRecAs.Text
RsTxn_LessorAgreement!AdvanceGiven = Val(TxtAdvanceGiven.Text)
RsTxn_LessorAgreement!AdvanceAdjusted = Val(TxtAdvAdjusted.Text)
RsTxn_LessorAgreement!AdvanceReturn = Val(TxtAdvanceRemain.Text)
RsTxn_LessorAgreement!AdvanceBalance = Val(TxtBalance.Text)
RsTxn_LessorAgreement!Remark = TxtRemark.Text
RsTxn_LessorAgreement!Flag = mStatus
RsTxn_LessorAgreement!AgreenotReceieveFlag = Check1.Value
RsTxn_LessorAgreement!BRAgreement = ChkBRAgreement.Value
If ChkBRAgreement.Value = 1 Then
   RsTxn_LessorAgreement!ServiceTaxID = GetServiceTaxID("GST without Tax")
End If
If IsNull(RsTxn_LessorAgreement!CreatedBy) = True Or RsTxn_LessorAgreement!CreatedBy = "" Then
   RsTxn_LessorAgreement!CreatedBy = Gen_UserLoginID
   RsTxn_LessorAgreement!CreatedDate = Now
Else
   RsTxn_LessorAgreement!UpdatedBy = Gen_UserLoginID
   RsTxn_LessorAgreement!UpdatedDate = Now
End If

RsTxn_LessorAgreement.Update

Call SaveLessorDetails
Call SaveGodownLessorDetails
Call SaveAttachmentDetails
SaveCmd.Enabled = False
Call Indata


End Sub
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

Call CheckGodownRent

End Sub
Private Sub CheckGodownRent()

'mMsg = ""
'mAmount = 0
'
'
'tmp = "Select * From Mst_Godown"
'Call Tmp4Table
'
'
'tmp = "Select * From Mst_InspectionExpenseDetail Where InsExpTypeID = 1"
'
'Call Tmp1Table
'
'tmp = "Select * From Mst_GodownInspection Where Status in ('G') order by StartDate Desc"
'
'Call Tmp2Table
'
'tmp = "Select InspectionID,convert(varchar,CMPID) +'-' +convert(varchar,GodownNo) as keyvalue  From Mst_GodownInspection"
'tmp = tmp & " Where Status in ('G') And convert(varchar,CMPID) +'-' +convert(varchar,GodownNo)<>'NULL'"
'tmp = tmp & " order by StartDate Desc"
'Call Tmp3Table
'mAmount = 0
'mGodownID = ""
'mGodownRentAmount = 0
'For i = 1 To MSHFlexGrid6.Rows - 2
'    mAmount = 0
'    mGodownID = ""
'    mGodownRentAmount = 0
'    mGodownID = Val(MSHFlexGrid6.TextMatrix(i, 0))
'
'    TmpRs2.MoveFirst
'    TmpRs2.Find "GodownID = " & Val(MSHFlexGrid6.TextMatrix(i, 0))
'    If Not TmpRs2.EOF Then
'       TmpRs1.MoveFirst
'       TmpRs1.Find "InspectionID = " & TmpRs2!InspectionID
'       If Not TmpRs1.EOF Then
'          mAmount = TmpRs1!Amount
'       End If
'    Else
'       TmpRs4.MoveFirst
'       TmpRs4.Find "GodownID = " & Val(MSHFlexGrid6.TextMatrix(i, 0))
'       If Not TmpRs4.EOF Then
'          TmpRs3.MoveFirst
'          TmpRs3.Find "keyvalue='" & TmpRs4!CMPID & "-" & TmpRs4!GodownNo & "'"
'          If Not TmpRs3.EOF Then
'             TmpRs1.MoveFirst
'             TmpRs1.Find "InspectionID = " & TmpRs3!InspectionID
'             If Not TmpRs1.EOF Then
'                mAmount = TmpRs1!Amount
'              End If
'          End If
'       End If
'    End If
'
'    For j = 2 To MSHFlexGrid1.Rows - 2
'        If MSHFlexGrid1.TextMatrix(j, 0) <> "" Then
'           If Val(MSHFlexGrid1.TextMatrix(j, 0)) = mGodownID Then
'              mGodownRentAmount = mGodownRentAmount + Val(MSHFlexGrid1.TextMatrix(j, 8))
'           End If
'        End If
'    Next
'    If mGodownRentAmount > mAmount Then
'       mMsg = mMsg & " Godown Rate is greater then Viability Rate for given Godown No " & MSHFlexGrid6.TextMatrix(i, 1) & vbCrLf
'    End If
'
'    If mGodownRentAmount < mAmount Then
'       mMsg = mMsg & "Godown Rate is less than Viability Rate for given Godown No  " & MSHFlexGrid6.TextMatrix(i, 1) & vbCrLf
'    End If
'Next

mGodownExists = False
mMsg = ""
mAmount = 0


tmp = "Select * From Mst_Godown"
Call Tmp4Table


tmp = "Select * From Mst_InspectionExpenseDetail Where InsExpTypeID = 1"

Call Tmp1Table

tmp = "Select * From Mst_GodownInspection Where Status in ('G') order by StartDate Desc"

Call Tmp2Table

tmp = "Select InspectionID,convert(varchar,CMPID) +'-' +convert(varchar,GodownNo) as keyvalue  From Mst_GodownInspection"
tmp = tmp & " Where Status in ('G') And convert(varchar,CMPID) +'-' +convert(varchar,GodownNo)<>'NULL'"
tmp = tmp & " order by StartDate Desc"
Call Tmp3Table
mAmount = 0
mGodownID = ""
mGodownRentAmount = 0
For I = 1 To MSHFlexGrid6.Rows - 2
    mAmount = 0
    mGodownID = ""
    mGodownRentAmount = 0
    mGodownID = Val(MSHFlexGrid6.TextMatrix(I, 0))
    
    TmpRs2.MoveFirst
    TmpRs2.Find "GodownID = " & Val(MSHFlexGrid6.TextMatrix(I, 0))
    If Not TmpRs2.EOF Then
       mGodownExists = True
       TmpRs1.MoveFirst
       TmpRs1.Find "InspectionID = " & TmpRs2!InspectionID
       If Not TmpRs1.EOF Then
          mAmount = TmpRs1!Amount
       End If
    Else
       TmpRs4.MoveFirst
       TmpRs4.Find "GodownID = " & Val(MSHFlexGrid6.TextMatrix(I, 0))
       If Not TmpRs4.EOF Then
          TmpRs3.MoveFirst
          TmpRs3.Find "keyvalue='" & TmpRs4!CMPID & "-" & TmpRs4!GodownNo & "'"
          If Not TmpRs3.EOF Then
             mGodownExists = True
             TmpRs1.MoveFirst
             TmpRs1.Find "InspectionID = " & TmpRs3!InspectionID
             If Not TmpRs1.EOF Then
                mAmount = TmpRs1!Amount
              End If
          End If
       End If
    End If
    
    If mGodownExists = False And Edit_Flg = "A" Then
       mGodownMsg = "Godown No " & MSHFlexGrid6.TextMatrix(I, 1) & " not exists in Viability !!!!"
       Exit Sub
    End If
      
    For j = 2 To MSHFlexGrid1.Rows - 2
        If MSHFlexGrid1.TextMatrix(j, 0) <> "" Then
           If Val(MSHFlexGrid1.TextMatrix(j, 0)) = mGodownID Then
              mGodownRentAmount = mGodownRentAmount + Val(MSHFlexGrid1.TextMatrix(j, 8))
           End If
        End If
    Next
    
    If mGodownExists = True Then
       If mGodownRentAmount > mAmount Then
          mMsg = mMsg & " Godown Rate is greater then Viability Rate for given Godown No " & MSHFlexGrid6.TextMatrix(I, 1) & vbCrLf
       End If
    
       If mGodownRentAmount < mAmount Then
          mMsg = mMsg & "Godown Rate is less than Viability Rate for given Godown No  " & MSHFlexGrid6.TextMatrix(I, 1) & vbCrLf
       End If
    End If
Next



End Sub
Sub DeleteChildTableDetail()
tmp = "Delete from Txn_LessorAgreementDetail where LAID = " & Val(TxtAgreementID.Text) & ""
Call TmpTable
tmp = "Delete from Txn_GodownLessorAgreeDetail where LAID = " & Val(TxtAgreementID.Text) & ""
Call TmpTable
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
   If RsTxn_LessorAgreement.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   RsTxn_LessorAgreement.MoveFirst
   Call ClearGodownFrame
   Call ClearLessorFrame
   RsTxn_LessorAgreement.Find "LAID = " & Val(TxtAgreementID) & ""
   If RsTxn_LessorAgreement.EOF Then
      RsTxn_LessorAgreement.MoveLast
   End If
   Call Indata
   Exit Sub
End If

Frame0.Height = Me.Height - 500
Frame0.Left = Frame1.Left - 100
Frame0.Top = Frame1.Top
Frame0.Width = Me.Width - 200
Frame0.Visible = True

MSHFlexGrid3.Width = Frame0.Width - 200
MSHFlexGrid3.Height = Frame0.Height - 2500

RsTxn_LessorAgreement.MoveFirst
RsTxn_LessorAgreement.Filter = ""

''===========comment on 7 july 2012 for not clearing search grid=============
'Call Grid_Head
'SFrom.Value = Date - 365
'STo.Value = Date
'Label22.Caption = ""
'CmbSAgreRecAs.Text = ""
'CmbSCmp.Text = ""
'CmbSGodown.Text = ""
'CmbSLessor.Text = ""
'CmbSStatus.Text = ""
'TxtSAgreeID.Text = ""
''===========comment on 7 july 2012 for not clearing search grid=============

End Sub
Private Sub Grid_Head3()

MSHFlexGrid4.Clear
MSHFlexGrid4.Rows = 2
MSHFlexGrid4.FixedRows = 1
MSHFlexGrid4.Cols = 3

MSHFlexGrid4.TextMatrix(0, 0) = "Lessor Agreement ID"
MSHFlexGrid4.TextMatrix(0, 1) = "Updated By"
MSHFlexGrid4.TextMatrix(0, 2) = "Updated On"

MSHFlexGrid4.ColWidth(0) = 2200
MSHFlexGrid4.ColWidth(1) = 2000
MSHFlexGrid4.ColWidth(2) = 3000

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

Private Sub TxtRentAmt_LostFocus()
If TxtRentAmt = "" Then Exit Sub
If IsNumeric(TxtRentAmt) = False Or Val(TxtRentAmt) < 0 Then
   MsgBox "Invalid numeric format!!!"
   TxtRentAmt.SetFocus
   Exit Sub
End If
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

Private Sub CmdRemoveAttachment_Click()
I = MSHFlexGrid5.RowSel
If I > 0 Then
    If MSHFlexGrid5.TextMatrix(I, 4) = "" And MSHFlexGrid5.TextMatrix(I, 1) <> "" Then
        For RO = I To MSHFlexGrid5.Rows - 2
            For C = 0 To MSHFlexGrid5.Cols - 1
                MSHFlexGrid5.TextMatrix(RO, C) = MSHFlexGrid5.TextMatrix(RO + 1, C)
            Next
        Next
        MSHFlexGrid5.Rows = MSHFlexGrid5.Rows - 1
    End If
End If

Call ClearFrame
End Sub
Private Sub ClearFrame()

lblAttachment.Caption = "File Attachment"
CmdAttachment.Caption = "Add Attachment"
TxtAttachmentRemarks = ""

End Sub

Private Sub SaveAttachmentDetails()

If MSHFlexGrid5.Rows <= 1 Then Exit Sub

'tmp = "Select * from Txn_StorageContractAttachment Where SContractID = " & TxtContractID & " "

Call TableTxn_LessorAgreeAttachment

For I = 1 To MSHFlexGrid5.Rows - 1
    If MSHFlexGrid5.TextMatrix(I, 1) = "" Then Exit For
    If MSHFlexGrid5.TextMatrix(I, 4) = "" Then
       RsTxn_LessorAgreeAttachment.AddNew
 
       RsTxn_LessorAgreeAttachment!TxnID = GetMaxAttachID
       RsTxn_LessorAgreeAttachment!LAID = TxtAgreementID
       mFlag = SaveFileToDB(MSHFlexGrid5.TextMatrix(I, 3), RsTxn_LessorAgreeAttachment, "File")
       mFlag = GetFileFromPath(MSHFlexGrid5.TextMatrix(I, 3))
       RsTxn_LessorAgreeAttachment!FileName = "LessorAgreement" & TxtAgreementID & "-" & mFlag
    Else
       RsTxn_LessorAgreeAttachment.MoveFirst
       RsTxn_LessorAgreeAttachment.Find "TxnID = " & MSHFlexGrid5.TextMatrix(I, 4) & ""
    End If
    RsTxn_LessorAgreeAttachment!Remarks = MSHFlexGrid5.TextMatrix(I, 2)
    RsTxn_LessorAgreeAttachment!UpdatedDate = Now
    RsTxn_LessorAgreeAttachment!UpdatedBy = Gen_UserLoginID
    RsTxn_LessorAgreeAttachment!G_UID = GetGUID
    If IsNull(RsTxn_LessorAgreeAttachment!CreatedBy) = True Or RsTxn_LessorAgreeAttachment!CreatedBy = "" Then
       RsTxn_LessorAgreeAttachment!CreatedBy = Gen_UserLoginID
       RsTxn_LessorAgreeAttachment!CreatedDate = Now
    Else
       RsTxn_LessorAgreeAttachment!UpdatedBy = Gen_UserLoginID
       RsTxn_LessorAgreeAttachment!UpdatedDate = Now
    End If
    RsTxn_LessorAgreeAttachment.Update
Next

End Sub

Private Sub IndataAttachmentDetails()
Call Grid_Head4
Call ClearFrame

Call TableTxn_LessorAgreeAttachment(" Where LAID = " & TxtAgreementID.Text & " ")

MSHFlexGrid5.Rows = RsTxn_LessorAgreeAttachment.RecordCount + 2
For I = 1 To RsTxn_LessorAgreeAttachment.RecordCount

    MSHFlexGrid5.TextMatrix(I, 0) = I
    MSHFlexGrid5.TextMatrix(I, 1) = RsTxn_LessorAgreeAttachment!FileName
    MSHFlexGrid5.TextMatrix(I, 2) = RsTxn_LessorAgreeAttachment!Remarks
    'MSHFlexGrid5.TextMatrix(i, 3) = App.Path & "\Agreement Attachment\" & MSHFlexGrid5.TextMatrix(i, 1)
    MSHFlexGrid5.TextMatrix(I, 3) = App.Path & "\Agreement Attachment\Lessor\" & MSHFlexGrid5.TextMatrix(I, 1)
    File_Path = MSHFlexGrid5.TextMatrix(I, 3) 'App.Path & "\Project Attachment\"
    Call LoadFileFromDB(File_Path, RsTxn_LessorAgreeAttachment, "File")
    MSHFlexGrid5.TextMatrix(I, 4) = RsTxn_LessorAgreeAttachment!TxnID
    RsTxn_LessorAgreeAttachment.MoveNext
Next
lblAttachment.Enabled = False
TxtAttachmentRemarks.Locked = True
CmdAttachment.Enabled = False
CmdGetFile.Enabled = False
CmdRemoveAttachment.Enabled = False
End Sub

Private Function GetMaxAttachID() As Double
Dim Retval As Double
tmp = "Select max(TxnID) from Txn_LessorAgreeAttachment"
Call TmpTableSMAttachment
Retval = IIf(IsNull(TmpRsSM_Attachment.Fields(0)), 0, TmpRsSM_Attachment.Fields(0))
Retval = Retval + 1
GetMaxAttachID = Retval
End Function

Private Sub CmdAttachment_Click()
Dim r As Variant

If lblAttachment.Caption = "File Attachment" Then
   MsgBox "Blank Attachment not allowed!!!"
   Exit Sub
End If

If TxtAttachmentRemarks = "" Then
   MsgBox "Blank Remarks not allowed!!!"
   TxtAttachmentRemarks.SetFocus
   Exit Sub
End If

If CmdAttachment.Caption = "Update Attachment" Then
   r = MSHFlexGrid3.RowSel
Else
   r = MSHFlexGrid5.Rows - 1
   MSHFlexGrid5.Rows = MSHFlexGrid5.Rows + 1
   MSHFlexGrid5.TextMatrix(r, 0) = ""
End If

MSHFlexGrid5.TextMatrix(r, 1) = GetFileFromPath(lblAttachment.Caption)
MSHFlexGrid5.TextMatrix(r, 2) = TxtAttachmentRemarks
MSHFlexGrid5.TextMatrix(r, 3) = lblAttachment.Caption
Call ClearFrame

End Sub

Private Sub Grid_Head4()
With MSHFlexGrid5
    .Clear
    .Rows = 2
    .Cols = 5
    .Font.Size = 10
    .WordWrap = True
    .TextMatrix(0, 0) = "SR No"
    .TextMatrix(0, 1) = "Attachment File Name"
    .TextMatrix(0, 2) = "Remarks"
    .TextMatrix(0, 3) = "File Path"
    .TextMatrix(0, 4) = "ID"
    .ColWidth(0) = 1000
    .ColWidth(1) = 5000
    .ColWidth(2) = 8200
    .ColWidth(3) = 0
    .ColWidth(4) = 0
End With
End Sub
Private Sub Grid_Head5()
MSHFlexGrid6.Clear
MSHFlexGrid6.Rows = 2
MSHFlexGrid6.Cols = 2
MSHFlexGrid6.FixedRows = 1
MSHFlexGrid6.WordWrap = True
MSHFlexGrid6.TextMatrix(0, 0) = "GodownID"
MSHFlexGrid6.TextMatrix(0, 1) = "GodownNo"
MSHFlexGrid6.ColWidth(0) = 1000
MSHFlexGrid6.ColWidth(1) = 1000
End Sub

