VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmMst_GodownInspection 
   Caption         =   "Godown Viability"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      TabIndex        =   127
      Top             =   10080
      Width           =   10935
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   9015
      Left            =   15120
      TabIndex        =   126
      Top             =   0
      Width           =   255
   End
   Begin VB.Frame Frame0 
      Height          =   2175
      Left            =   15480
      TabIndex        =   0
      Top             =   4440
      Visible         =   0   'False
      Width           =   5775
      Begin VB.TextBox TxtSViabilityID 
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
         Left            =   9075
         TabIndex        =   76
         Top             =   960
         Width           =   2160
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
         ItemData        =   "FrmMst_GodownInspection.frx":0000
         Left            =   5040
         List            =   "FrmMst_GodownInspection.frx":0010
         Sorted          =   -1  'True
         TabIndex        =   75
         Top             =   960
         Width           =   2685
      End
      Begin VB.ComboBox CmbSGodownType 
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
         Left            =   11880
         Sorted          =   -1  'True
         TabIndex        =   72
         Top             =   480
         Width           =   2325
      End
      Begin VB.ComboBox CmbSType 
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
         ItemData        =   "FrmMst_GodownInspection.frx":0042
         Left            =   120
         List            =   "FrmMst_GodownInspection.frx":004C
         Sorted          =   -1  'True
         TabIndex        =   68
         Top             =   480
         Width           =   2325
      End
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   350
         Left            =   13125
         TabIndex        =   78
         Top             =   960
         Width           =   1000
      End
      Begin VB.ComboBox CmbSCmp 
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
         Left            =   5475
         Sorted          =   -1  'True
         TabIndex        =   70
         Top             =   480
         Width           =   4005
      End
      Begin VB.ComboBox CmbSLocation 
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
         ItemData        =   "FrmMst_GodownInspection.frx":005E
         Left            =   2640
         List            =   "FrmMst_GodownInspection.frx":0060
         Sorted          =   -1  'True
         TabIndex        =   69
         Top             =   480
         Width           =   2685
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "Search"
         Height          =   350
         Left            =   12120
         TabIndex        =   77
         Top             =   960
         Width           =   1000
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
         Left            =   9600
         Sorted          =   -1  'True
         TabIndex        =   71
         Top             =   480
         Width           =   2085
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1335
         Left            =   120
         TabIndex        =   79
         Top             =   1440
         Width           =   3150
         _ExtentX        =   5556
         _ExtentY        =   2355
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
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   120
         TabIndex        =   146
         Top             =   1920
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.TextBox TxtSGodown 
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
         Left            =   9600
         TabIndex        =   152
         Top             =   480
         Visible         =   0   'False
         Width           =   2085
      End
      Begin MSComCtl2.DTPicker STo 
         Height          =   300
         Left            =   3000
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   960
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
         Format          =   115343361
         CurrentDate     =   36494
      End
      Begin MSComCtl2.DTPicker SFrom 
         Height          =   300
         Left            =   1200
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   960
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
         Format          =   115343361
         CurrentDate     =   36494
      End
      Begin VB.Label Label84 
         Caption         =   "Start From"
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
         Left            =   180
         TabIndex        =   221
         Top             =   990
         Width           =   990
      End
      Begin VB.Label Label83 
         Caption         =   "To"
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
         Left            =   2685
         TabIndex        =   220
         Top             =   990
         Width           =   300
      End
      Begin VB.Label Label82 
         AutoSize        =   -1  'True
         Caption         =   "Viability ID"
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
         Left            =   8040
         TabIndex        =   214
         Top             =   1020
         Width           =   945
      End
      Begin VB.Label Label71 
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
         Left            =   4440
         TabIndex        =   193
         Top             =   1020
         Width           =   555
      End
      Begin VB.Label Label51 
         AutoSize        =   -1  'True
         Caption         =   "Godown Type"
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
         Left            =   12405
         TabIndex        =   151
         Top             =   210
         Width           =   1275
      End
      Begin VB.Label Label54 
         AutoSize        =   -1  'True
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
         Left            =   1042
         TabIndex        =   150
         Top             =   225
         Width           =   480
      End
      Begin VB.Label Label50 
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
         Left            =   3600
         TabIndex        =   149
         Top             =   225
         Width           =   765
      End
      Begin VB.Label Label49 
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
         Left            =   7260
         TabIndex        =   81
         Top             =   225
         Width           =   435
      End
      Begin VB.Label Label53 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   10260
         TabIndex        =   80
         Top             =   225
         Width           =   750
      End
   End
   Begin VB.PictureBox Frame4 
      Height          =   9555
      Left            =   45
      ScaleHeight     =   9495
      ScaleWidth      =   14835
      TabIndex        =   82
      Top             =   45
      Width           =   14895
      Begin VB.PictureBox Frame5 
         Height          =   4935
         Left            =   3960
         ScaleHeight     =   4875
         ScaleWidth      =   8475
         TabIndex        =   215
         Top             =   3480
         Width           =   8535
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            Caption         =   "Select All"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   219
            Top             =   60
            Width           =   2055
         End
         Begin VB.CommandButton CmdCancel 
            Caption         =   "&Cancel"
            Height          =   375
            Left            =   1440
            TabIndex        =   218
            Top             =   4440
            Width           =   1335
         End
         Begin VB.CommandButton CmdSave 
            Caption         =   "&Save"
            Height          =   375
            Left            =   120
            TabIndex        =   217
            Top             =   4440
            Width           =   1335
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid8 
            Height          =   3975
            Left            =   120
            TabIndex        =   216
            Top             =   360
            Width           =   8250
            _ExtentX        =   14552
            _ExtentY        =   7011
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
      End
      Begin VB.Frame FrmPreRent 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Previous Rent Details"
         Height          =   1455
         Left            =   6000
         TabIndex        =   206
         Top             =   3600
         Visible         =   0   'False
         Width           =   6015
         Begin VB.CommandButton Command1 
            Caption         =   "Hide"
            Height          =   375
            Left            =   4920
            TabIndex        =   211
            Top             =   960
            Width           =   975
         End
         Begin VB.Label LblLeaseRent 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "--"
            Height          =   195
            Left            =   3360
            TabIndex        =   210
            Top             =   840
            Width           =   90
         End
         Begin VB.Label LblPRenewalRent 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "--"
            Height          =   195
            Left            =   3360
            TabIndex        =   209
            Top             =   480
            Width           =   90
         End
         Begin VB.Label Label81 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Lease Rent (from Lessor Agreement) :"
            Height          =   195
            Left            =   360
            TabIndex        =   208
            Top             =   840
            Width           =   2670
         End
         Begin VB.Label Label80 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Previous Renewal Rent (if any):"
            Height          =   195
            Left            =   360
            TabIndex        =   207
            Top             =   480
            Width           =   2235
         End
      End
      Begin VB.PictureBox Picture3 
         Height          =   975
         Left            =   120
         ScaleHeight     =   915
         ScaleWidth      =   14595
         TabIndex        =   125
         Top             =   8520
         Width           =   14655
         Begin VB.CommandButton AddCmd 
            Caption         =   "&Add New "
            Height          =   375
            Left            =   45
            TabIndex        =   58
            Top             =   105
            Width           =   2895
         End
         Begin VB.CommandButton SaveCmd 
            Caption         =   "&Save"
            Height          =   375
            Left            =   2940
            TabIndex        =   59
            Top             =   105
            Width           =   2895
         End
         Begin VB.CommandButton EditCmd 
            Caption         =   "&Edit"
            Height          =   375
            Left            =   5835
            TabIndex        =   60
            Top             =   105
            Width           =   2895
         End
         Begin VB.CommandButton DeleteCmd 
            Caption         =   "&Delete"
            Height          =   375
            Left            =   8730
            TabIndex        =   61
            Top             =   105
            Width           =   2895
         End
         Begin VB.CommandButton SearchCmd 
            Caption         =   "&Search"
            Height          =   375
            Left            =   11625
            TabIndex        =   62
            Top             =   105
            Width           =   2895
         End
         Begin VB.CommandButton NextCmd 
            Caption         =   "&Next"
            Height          =   375
            Left            =   45
            TabIndex        =   63
            Top             =   480
            Width           =   2895
         End
         Begin VB.CommandButton PreviousCmd 
            Caption         =   "&Previous"
            Height          =   375
            Left            =   2940
            TabIndex        =   64
            Top             =   480
            Width           =   2895
         End
         Begin VB.CommandButton FirstCmd 
            Caption         =   "&First"
            Height          =   375
            Left            =   5835
            TabIndex        =   65
            Top             =   480
            Width           =   2895
         End
         Begin VB.CommandButton LastCmd 
            Caption         =   "&Last"
            Height          =   375
            Left            =   8730
            TabIndex        =   66
            Top             =   480
            Width           =   2895
         End
         Begin VB.CommandButton ExitCmd 
            Caption         =   "E&xit"
            Height          =   375
            Left            =   11625
            TabIndex        =   67
            Top             =   480
            Width           =   2895
         End
      End
      Begin VB.PictureBox Picture2 
         Height          =   2835
         Left            =   120
         ScaleHeight     =   2775
         ScaleWidth      =   14595
         TabIndex        =   85
         Top             =   120
         Width           =   14655
         Begin VB.TextBox TxtWeightPerStack 
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
            Left            =   5670
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   1005
            Width           =   1845
         End
         Begin VB.TextBox TxtNoOfStack 
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
            Locked          =   -1  'True
            TabIndex        =   8
            Top             =   1005
            Width           =   1965
         End
         Begin VB.TextBox TxtInspectionID 
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
            Left            =   13110
            Locked          =   -1  'True
            TabIndex        =   123
            TabStop         =   0   'False
            Top             =   2340
            Width           =   1395
         End
         Begin VB.CommandButton CmdCMP 
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
            Height          =   360
            Left            =   14100
            TabIndex        =   117
            TabStop         =   0   'False
            Top             =   75
            Width           =   375
         End
         Begin MSComCtl2.DTPicker TxtStartDate 
            Height          =   360
            Left            =   10320
            TabIndex        =   14
            Top             =   2340
            Width           =   1320
            _ExtentX        =   2328
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
            Format          =   115343361
            CurrentDate     =   39867
         End
         Begin MSComCtl2.DTPicker TxtCloseDate 
            Height          =   360
            Left            =   11700
            TabIndex        =   15
            Top             =   2340
            Width           =   1320
            _ExtentX        =   2328
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
            Format          =   115343361
            CurrentDate     =   39867
         End
         Begin VB.ComboBox CmbType 
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
            ItemData        =   "FrmMst_GodownInspection.frx":0062
            Left            =   1680
            List            =   "FrmMst_GodownInspection.frx":006C
            Sorted          =   -1  'True
            TabIndex        =   1
            Top             =   75
            Width           =   2085
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
            Left            =   1680
            Sorted          =   -1  'True
            TabIndex        =   4
            Top             =   525
            Width           =   2085
         End
         Begin VB.TextBox TxtGodown 
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
            MaxLength       =   25
            TabIndex        =   5
            Top             =   525
            Width           =   2085
         End
         Begin VB.TextBox TxtType 
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
            TabIndex        =   86
            TabStop         =   0   'False
            Top             =   75
            Width           =   2085
         End
         Begin VB.TextBox TxtCapacity 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
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
            Left            =   12540
            Locked          =   -1  'True
            TabIndex        =   20
            Top             =   975
            Width           =   1965
         End
         Begin VB.TextBox TxtArea 
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
            Left            =   8940
            Locked          =   -1  'True
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   975
            Width           =   1725
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
            Left            =   4950
            Sorted          =   -1  'True
            TabIndex        =   2
            Top             =   75
            Width           =   2595
         End
         Begin VB.TextBox TxtAddress 
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
            Height          =   915
            Left            =   1680
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   13
            Top             =   1845
            Width           =   8565
         End
         Begin VB.TextBox TxtLocation 
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
            Left            =   4950
            Locked          =   -1  'True
            TabIndex        =   88
            TabStop         =   0   'False
            Top             =   75
            Width           =   2595
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
            Left            =   8910
            Sorted          =   -1  'True
            TabIndex        =   3
            Top             =   75
            Width           =   5115
         End
         Begin VB.TextBox TxtCMP 
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
            Left            =   8910
            Locked          =   -1  'True
            TabIndex        =   90
            TabStop         =   0   'False
            Top             =   75
            Width           =   5115
         End
         Begin VB.ComboBox CmbLicenseNo 
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
            TabIndex        =   11
            Top             =   1440
            Width           =   9165
         End
         Begin VB.TextBox TxtLicenseNo 
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
            Left            =   1710
            Locked          =   -1  'True
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   1440
            Width           =   9090
         End
         Begin VB.ComboBox CmbGodownType 
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
            TabIndex        =   6
            Top             =   525
            Width           =   2100
         End
         Begin VB.TextBox TxtGodownType 
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
            Left            =   5400
            Locked          =   -1  'True
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   525
            Width           =   2085
         End
         Begin VB.ComboBox CmbGodownCategory 
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
            Left            =   8910
            Sorted          =   -1  'True
            TabIndex        =   7
            Top             =   525
            Width           =   2205
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
            ItemData        =   "FrmMst_GodownInspection.frx":007E
            Left            =   12060
            List            =   "FrmMst_GodownInspection.frx":008E
            Sorted          =   -1  'True
            TabIndex        =   12
            Top             =   1440
            Width           =   2445
         End
         Begin VB.TextBox TxtGodownCategory 
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
            Left            =   8910
            Locked          =   -1  'True
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   525
            Width           =   2205
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
            Left            =   12060
            Locked          =   -1  'True
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   1440
            Width           =   1965
         End
         Begin VB.Label Label90 
            AutoSize        =   -1  'True
            Caption         =   "Weight Per Stack"
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
            Left            =   3840
            TabIndex        =   235
            Top             =   1065
            Width           =   1800
         End
         Begin VB.Label Label89 
            AutoSize        =   -1  'True
            Caption         =   "No of Stack"
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
            TabIndex        =   234
            Top             =   1065
            Width           =   1335
         End
         Begin VB.Label Label77 
            AutoSize        =   -1  'True
            Caption         =   "License No"
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
            TabIndex        =   202
            Top             =   1500
            Width           =   1185
         End
         Begin VB.Label Label72 
            AutoSize        =   -1  'True
            Caption         =   "Category"
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
            TabIndex        =   194
            Top             =   585
            Width           =   960
         End
         Begin VB.Label Label48 
            AutoSize        =   -1  'True
            Caption         =   "Close Date"
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
            Left            =   11760
            TabIndex        =   147
            Top             =   2040
            Width           =   1170
         End
         Begin VB.Label LblVendorID 
            AutoSize        =   -1  'True
            Caption         =   "Viability ID"
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
            Left            =   13335
            TabIndex        =   124
            Top             =   2040
            Width           =   945
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            Caption         =   "Start Date"
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
            Left            =   10440
            TabIndex        =   120
            Top             =   2040
            Width           =   1050
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Address"
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
            TabIndex        =   93
            Top             =   2182
            Width           =   885
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "CMP Name"
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
            TabIndex        =   91
            Top             =   135
            Width           =   1170
         End
         Begin VB.Label Label2 
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
            Left            =   3840
            TabIndex        =   89
            Top             =   135
            Width           =   900
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "Capacity in MT"
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
            TabIndex        =   116
            Top             =   1035
            Width           =   1560
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Godown No"
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
            TabIndex        =   115
            Top             =   585
            Width           =   1215
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Area in Sq.Ft."
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
            Left            =   7560
            TabIndex        =   114
            Top             =   1035
            Width           =   1410
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Type"
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
            TabIndex        =   87
            Top             =   135
            Width           =   555
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
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
            Left            =   10920
            TabIndex        =   122
            Top             =   1500
            Width           =   660
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Godown Type"
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
            Left            =   3840
            TabIndex        =   92
            Top             =   585
            Width           =   1455
         End
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   5415
         Left            =   120
         TabIndex        =   21
         Top             =   3000
         Width           =   14655
         _ExtentX        =   25850
         _ExtentY        =   9551
         _Version        =   393216
         Tabs            =   9
         TabsPerRow      =   9
         TabHeight       =   970
         TabCaption(0)   =   "Staff && Utilisation Details"
         TabPicture(0)   =   "FrmMst_GodownInspection.frx":00C0
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "LblStorageDays"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label11"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label37"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label9"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Label45"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Label42"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "Label10"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "Label47"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "Label55"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "Label59"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "Label60"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "Label61"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "Label62"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "Label63"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "Label43"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "Label68"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "Label69"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "TxtStaffSalarySys"
         Tab(0).Control(17).Enabled=   0   'False
         Tab(0).Control(18)=   "TxtStaffsys"
         Tab(0).Control(18).Enabled=   0   'False
         Tab(0).Control(19)=   "TxtSecStaffSys"
         Tab(0).Control(19).Enabled=   0   'False
         Tab(0).Control(20)=   "TxtSecStaffSalarySys"
         Tab(0).Control(20).Enabled=   0   'False
         Tab(0).Control(21)=   "TxtStaffSalaryUser"
         Tab(0).Control(21).Enabled=   0   'False
         Tab(0).Control(22)=   "TxtSecStaffSalaryUser"
         Tab(0).Control(22).Enabled=   0   'False
         Tab(0).Control(23)=   "TxtSecStaffUser"
         Tab(0).Control(23).Enabled=   0   'False
         Tab(0).Control(24)=   "TxtStaffUser"
         Tab(0).Control(24).Enabled=   0   'False
         Tab(0).Control(25)=   "MSHFlexGrid6"
         Tab(0).Control(25).Enabled=   0   'False
         Tab(0).Control(26)=   "MSHFlexGrid2"
         Tab(0).Control(26).Enabled=   0   'False
         Tab(0).Control(27)=   "TxtExpUtilisationMT"
         Tab(0).Control(27).Enabled=   0   'False
         Tab(0).Control(28)=   "TxtRemarks"
         Tab(0).Control(28).Enabled=   0   'False
         Tab(0).Control(29)=   "TxtEnterData"
         Tab(0).Control(29).Enabled=   0   'False
         Tab(0).Control(30)=   "Frame3"
         Tab(0).Control(30).Enabled=   0   'False
         Tab(0).Control(31)=   "TxtExpUtilisation"
         Tab(0).Control(31).Enabled=   0   'False
         Tab(0).Control(32)=   "CmdGodownExcel"
         Tab(0).Control(32).Enabled=   0   'False
         Tab(0).ControlCount=   33
         TabCaption(1)   =   "Commodity Details"
         TabPicture(1)   =   "FrmMst_GodownInspection.frx":00DC
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "MSHFlexGrid3"
         Tab(1).Control(1)=   "Frame1"
         Tab(1).ControlCount=   2
         TabCaption(2)   =   "Expense Details"
         TabPicture(2)   =   "FrmMst_GodownInspection.frx":00F8
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Label20"
         Tab(2).Control(1)=   "Label21"
         Tab(2).Control(2)=   "Label22"
         Tab(2).Control(3)=   "Label23"
         Tab(2).Control(4)=   "Label24"
         Tab(2).Control(5)=   "Label25"
         Tab(2).Control(6)=   "Label19"
         Tab(2).Control(7)=   "Label18"
         Tab(2).Control(8)=   "Label17"
         Tab(2).Control(9)=   "Label16"
         Tab(2).Control(10)=   "Text10"
         Tab(2).Control(10).Enabled=   0   'False
         Tab(2).Control(11)=   "Text11"
         Tab(2).Control(11).Enabled=   0   'False
         Tab(2).Control(12)=   "Text17"
         Tab(2).Control(12).Enabled=   0   'False
         Tab(2).Control(13)=   "Text16"
         Tab(2).Control(13).Enabled=   0   'False
         Tab(2).Control(14)=   "Text15"
         Tab(2).Control(14).Enabled=   0   'False
         Tab(2).Control(15)=   "Text14"
         Tab(2).Control(15).Enabled=   0   'False
         Tab(2).Control(16)=   "Text13"
         Tab(2).Control(16).Enabled=   0   'False
         Tab(2).Control(17)=   "Text12"
         Tab(2).Control(17).Enabled=   0   'False
         Tab(2).Control(18)=   "MSHFlexGrid4"
         Tab(2).Control(19)=   "Frame2"
         Tab(2).ControlCount=   20
         TabCaption(3)   =   "Close Godown Details"
         TabPicture(3)   =   "FrmMst_GodownInspection.frx":0114
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "MSHFlexGrid5"
         Tab(3).Control(1)=   "CmdCloseGodown"
         Tab(3).ControlCount=   2
         TabCaption(4)   =   "Profit Summary"
         TabPicture(4)   =   "FrmMst_GodownInspection.frx":0130
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "Label30"
         Tab(4).Control(1)=   "Label65"
         Tab(4).Control(2)=   "lblaverageleasedrent"
         Tab(4).Control(3)=   "Label66"
         Tab(4).Control(4)=   "Label26"
         Tab(4).Control(5)=   "Label27"
         Tab(4).Control(6)=   "Label70"
         Tab(4).Control(7)=   "Label76"
         Tab(4).Control(8)=   "Label79"
         Tab(4).Control(9)=   "MSHFlexGrid7"
         Tab(4).Control(10)=   "TxtNetProfit"
         Tab(4).Control(10).Enabled=   0   'False
         Tab(4).Control(11)=   "TxtGrossProfit"
         Tab(4).Control(11).Enabled=   0   'False
         Tab(4).Control(12)=   "TxtIndirectExpense"
         Tab(4).Control(12).Enabled=   0   'False
         Tab(4).Control(13)=   "TxtTotalRevenue"
         Tab(4).Control(13).Enabled=   0   'False
         Tab(4).Control(14)=   "TxtDirectExpense"
         Tab(4).Control(14).Enabled=   0   'False
         Tab(4).Control(15)=   "TxtCurrentUtilisation"
         Tab(4).Control(15).Enabled=   0   'False
         Tab(4).Control(16)=   "TxtCMShare"
         Tab(4).Control(16).Enabled=   0   'False
         Tab(4).Control(17)=   "TxtSRentChange"
         Tab(4).Control(17).Enabled=   0   'False
         Tab(4).ControlCount=   18
         TabCaption(5)   =   "Area Details"
         TabPicture(5)   =   "FrmMst_GodownInspection.frx":014C
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "Frame6"
         Tab(5).Control(1)=   "MSHFlexGrid9"
         Tab(5).ControlCount=   2
         TabCaption(6)   =   "Utilization for previous month(s)"
         TabPicture(6)   =   "FrmMst_GodownInspection.frx":0168
         Tab(6).ControlEnabled=   0   'False
         Tab(6).Control(0)=   "MSHFlexGrid10"
         Tab(6).ControlCount=   1
         TabCaption(7)   =   "PnL Details for previous month(s)"
         TabPicture(7)   =   "FrmMst_GodownInspection.frx":0184
         Tab(7).ControlEnabled=   0   'False
         Tab(7).Control(0)=   "MSHFlexGrid11"
         Tab(7).ControlCount=   1
         TabCaption(8)   =   "Log History Detail"
         TabPicture(8)   =   "FrmMst_GodownInspection.frx":01A0
         Tab(8).ControlEnabled=   0   'False
         Tab(8).Control(0)=   "TxtUpdated"
         Tab(8).Control(0).Enabled=   0   'False
         Tab(8).Control(1)=   "TxtCreated"
         Tab(8).Control(1).Enabled=   0   'False
         Tab(8).Control(2)=   "Label44"
         Tab(8).Control(3)=   "Label46"
         Tab(8).ControlCount=   4
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
            Left            =   -73635
            Locked          =   -1  'True
            TabIndex        =   237
            TabStop         =   0   'False
            Top             =   1965
            Width           =   9315
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
            Left            =   -73635
            Locked          =   -1  'True
            TabIndex        =   236
            TabStop         =   0   'False
            Top             =   1440
            Width           =   9330
         End
         Begin VB.Frame Frame6 
            Height          =   1095
            Left            =   -74880
            TabIndex        =   223
            Top             =   600
            Width           =   14385
            Begin VB.CommandButton CmdAreaAddToList 
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
               Height          =   360
               Left            =   10440
               TabIndex        =   233
               Top             =   690
               Width           =   1935
            End
            Begin VB.CommandButton CmdAreaDeleteFromList 
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
               Height          =   360
               Left            =   12375
               TabIndex        =   232
               Top             =   690
               Width           =   1935
            End
            Begin VB.TextBox TxtSubArea 
               Alignment       =   1  'Right Justify
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
               Left            =   11460
               Locked          =   -1  'True
               TabIndex        =   230
               TabStop         =   0   'False
               Top             =   240
               Width           =   1965
            End
            Begin VB.TextBox TxtHeight 
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
               TabIndex        =   226
               Top             =   240
               Width           =   1605
            End
            Begin VB.TextBox TxtWidth 
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
               Left            =   4680
               TabIndex        =   225
               Top             =   240
               Width           =   1725
            End
            Begin VB.TextBox TxtLength 
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
               Left            =   1395
               TabIndex        =   224
               Top             =   240
               Width           =   1725
            End
            Begin VB.Label Label88 
               AutoSize        =   -1  'True
               Caption         =   "Area in Sq.Ft."
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
               Left            =   9720
               TabIndex        =   231
               Top             =   270
               Width           =   1410
            End
            Begin VB.Label Label85 
               AutoSize        =   -1  'True
               Caption         =   "Length(ft.in)"
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
               TabIndex        =   229
               Top             =   300
               Width           =   1215
            End
            Begin VB.Label Label86 
               AutoSize        =   -1  'True
               Caption         =   "Width(ft.in)"
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
               Left            =   3480
               TabIndex        =   228
               Top             =   300
               Width           =   1110
            End
            Begin VB.Label Label87 
               AutoSize        =   -1  'True
               Caption         =   "Height(ft.in)"
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
               Left            =   6720
               TabIndex        =   227
               Top             =   300
               Width           =   1200
            End
         End
         Begin VB.TextBox TxtSRentChange 
            Alignment       =   1  'Right Justify
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
            Left            =   -72960
            Locked          =   -1  'True
            TabIndex        =   212
            TabStop         =   0   'False
            Top             =   4080
            Width           =   1725
         End
         Begin VB.TextBox TxtCMShare 
            Alignment       =   1  'Right Justify
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
            Left            =   -72960
            Locked          =   -1  'True
            TabIndex        =   200
            TabStop         =   0   'False
            Top             =   3600
            Width           =   1725
         End
         Begin VB.CommandButton CmdCloseGodown 
            Caption         =   "Excel"
            Height          =   315
            Left            =   -61510
            TabIndex        =   191
            Top             =   4995
            Width           =   1000
         End
         Begin VB.CommandButton CmdGodownExcel 
            Caption         =   "Excel"
            Height          =   350
            Left            =   13490
            TabIndex        =   190
            Top             =   4995
            Width           =   1000
         End
         Begin VB.TextBox TxtCurrentUtilisation 
            Alignment       =   1  'Right Justify
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
            Left            =   -72960
            Locked          =   -1  'True
            TabIndex        =   188
            TabStop         =   0   'False
            Top             =   840
            Width           =   1725
         End
         Begin VB.TextBox TxtDirectExpense 
            Alignment       =   1  'Right Justify
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
            Left            =   -72960
            Locked          =   -1  'True
            TabIndex        =   181
            TabStop         =   0   'False
            Top             =   1785
            Width           =   1725
         End
         Begin VB.TextBox TxtTotalRevenue 
            Alignment       =   1  'Right Justify
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
            Left            =   -72960
            Locked          =   -1  'True
            TabIndex        =   179
            TabStop         =   0   'False
            Top             =   1320
            Width           =   1725
         End
         Begin VB.TextBox TxtIndirectExpense 
            Alignment       =   1  'Right Justify
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
            Left            =   -72960
            Locked          =   -1  'True
            TabIndex        =   177
            TabStop         =   0   'False
            Top             =   2691
            Width           =   1725
         End
         Begin VB.TextBox TxtExpUtilisation 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   9735
            Locked          =   -1  'True
            TabIndex        =   22
            Top             =   930
            Width           =   645
         End
         Begin VB.Frame Frame3 
            Caption         =   "Stock Detail"
            Height          =   1215
            Left            =   6360
            TabIndex        =   24
            Top             =   1260
            Width           =   5415
            Begin VB.TextBox TxtNCDEXStock 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   3375
               Locked          =   -1  'True
               TabIndex        =   25
               Top             =   165
               Width           =   645
            End
            Begin VB.TextBox TxtNNNRStock 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   3375
               Locked          =   -1  'True
               TabIndex        =   27
               Top             =   510
               Width           =   645
            End
            Begin VB.TextBox TxtReservationStock 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   3375
               Locked          =   -1  'True
               TabIndex        =   29
               Top             =   855
               Width           =   645
            End
            Begin VB.TextBox TxtReservationMT 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000004&
               Height          =   315
               Left            =   4080
               Locked          =   -1  'True
               TabIndex        =   30
               Top             =   855
               Width           =   1245
            End
            Begin VB.TextBox TxtNNNRMT 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000004&
               Height          =   315
               Left            =   4080
               Locked          =   -1  'True
               TabIndex        =   28
               Top             =   510
               Width           =   1245
            End
            Begin VB.TextBox TxtNCDEXMT 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000004&
               Height          =   315
               Left            =   4080
               Locked          =   -1  'True
               TabIndex        =   26
               Top             =   165
               Width           =   1245
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               Caption         =   "NCDEX Stock"
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
               Left            =   240
               TabIndex        =   174
               Top             =   225
               Width           =   1200
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "NNNR (including SCM) Stock"
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
               Left            =   240
               TabIndex        =   173
               Top             =   570
               Width           =   2505
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Reservation Stock"
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
               Left            =   240
               TabIndex        =   172
               Top             =   915
               Width           =   1590
            End
         End
         Begin VB.TextBox TxtGrossProfit 
            Alignment       =   1  'Right Justify
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
            Left            =   -72960
            Locked          =   -1  'True
            TabIndex        =   170
            TabStop         =   0   'False
            Top             =   2238
            Width           =   1725
         End
         Begin VB.TextBox TxtNetProfit 
            Alignment       =   1  'Right Justify
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
            Left            =   -72960
            Locked          =   -1  'True
            TabIndex        =   169
            TabStop         =   0   'False
            Top             =   3144
            Width           =   1725
         End
         Begin VB.TextBox TxtEnterData 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1560
            TabIndex        =   31
            Top             =   1440
            Visible         =   0   'False
            Width           =   645
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
            Height          =   915
            Left            =   1095
            MaxLength       =   4000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   32
            Top             =   2520
            Width           =   13365
         End
         Begin VB.Frame Frame2 
            Height          =   2250
            Left            =   -74880
            TabIndex        =   131
            Top             =   600
            Width           =   14415
            Begin VB.TextBox TxtChangeInPer 
               Alignment       =   1  'Right Justify
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
               Left            =   12750
               Locked          =   -1  'True
               TabIndex        =   203
               Top             =   1065
               Width           =   1545
            End
            Begin VB.TextBox TxtRemarksExpense 
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
               Height          =   480
               Left            =   1440
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   195
               Top             =   1680
               Width           =   9705
            End
            Begin VB.CommandButton CmdClearData 
               Caption         =   "Clear Data"
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
               Left            =   12870
               TabIndex        =   184
               Top             =   1800
               Width           =   1425
            End
            Begin VB.TextBox TxtFormula 
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
               Height          =   675
               Left            =   1440
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   50
               Top             =   900
               Width           =   3465
            End
            Begin VB.TextBox TxtAmountPer 
               Alignment       =   1  'Right Justify
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
               Left            =   11100
               TabIndex        =   48
               TabStop         =   0   'False
               Top             =   420
               Width           =   1485
            End
            Begin VB.CommandButton CmdEUpdateList 
               Caption         =   "Update List"
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
               Left            =   11280
               TabIndex        =   52
               Top             =   1800
               Width           =   1575
            End
            Begin VB.TextBox TxtCalculationOn 
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
               Left            =   5160
               Locked          =   -1  'True
               TabIndex        =   46
               TabStop         =   0   'False
               Top             =   420
               Width           =   3525
            End
            Begin VB.TextBox TxtQtySq 
               Alignment       =   1  'Right Justify
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
               Left            =   8760
               Locked          =   -1  'True
               TabIndex        =   47
               Top             =   420
               Width           =   2205
            End
            Begin VB.TextBox TxtExpenseType 
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
               Left            =   90
               Locked          =   -1  'True
               TabIndex        =   45
               TabStop         =   0   'False
               Top             =   420
               Width           =   4935
            End
            Begin VB.TextBox TxtExpenseRemarks 
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
               Height          =   675
               Left            =   6000
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   51
               Top             =   900
               Width           =   6330
            End
            Begin VB.TextBox TxtAmount 
               Alignment       =   1  'Right Justify
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
               Left            =   12750
               Locked          =   -1  'True
               TabIndex        =   49
               Top             =   420
               Width           =   1545
            End
            Begin VB.Label LblViewPrevRent 
               AutoSize        =   -1  'True
               Caption         =   "View Previous rent"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   12585
               MousePointer    =   10  'Up Arrow
               TabIndex        =   205
               Top             =   1470
               Width           =   1650
            End
            Begin VB.Label Label78 
               AutoSize        =   -1  'True
               Caption         =   "Change in %"
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
               Left            =   12915
               TabIndex        =   204
               Top             =   780
               Width           =   1125
            End
            Begin VB.Label Label73 
               Caption         =   "Expense Remarks"
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
               TabIndex        =   196
               Top             =   1620
               Width           =   900
            End
            Begin VB.Label Label64 
               AutoSize        =   -1  'True
               Caption         =   "Formula"
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
               TabIndex        =   176
               Top             =   1117
               Width           =   735
            End
            Begin VB.Label Label52 
               AutoSize        =   -1  'True
               Caption         =   "Amount per Unit"
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
               Left            =   11137
               TabIndex        =   148
               Top             =   120
               Width           =   1410
            End
            Begin VB.Label Label33 
               AutoSize        =   -1  'True
               Caption         =   "Expense Type"
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
               Left            =   1597
               TabIndex        =   145
               Top             =   120
               Width           =   1320
            End
            Begin VB.Label Label40 
               AutoSize        =   -1  'True
               Caption         =   "Calculation On"
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
               Left            =   6277
               TabIndex        =   144
               Top             =   120
               Width           =   1290
            End
            Begin VB.Label Label41 
               AutoSize        =   -1  'True
               Caption         =   "Units (Qty/Sq.Ft/No of)"
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
               Left            =   8872
               TabIndex        =   143
               Top             =   120
               Width           =   1980
            End
            Begin VB.Label Label58 
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
               Left            =   5040
               TabIndex        =   142
               Top             =   1110
               Width           =   825
            End
            Begin VB.Label Label31 
               AutoSize        =   -1  'True
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
               Left            =   13155
               TabIndex        =   141
               Top             =   120
               Width           =   675
            End
         End
         Begin VB.Frame Frame1 
            Height          =   1935
            Left            =   -74880
            TabIndex        =   130
            Top             =   570
            Width           =   14415
            Begin VB.TextBox TxtCMShareR 
               Alignment       =   1  'Right Justify
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
               Left            =   1770
               Locked          =   -1  'True
               TabIndex        =   198
               TabStop         =   0   'False
               Top             =   1515
               Width           =   1515
            End
            Begin VB.TextBox TxtCMShareP 
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
               Left            =   1770
               TabIndex        =   41
               Top             =   1125
               Width           =   1515
            End
            Begin VB.CommandButton CmdStorageRate 
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
               Height          =   360
               Left            =   2910
               TabIndex        =   192
               TabStop         =   0   'False
               Top             =   690
               Width           =   375
            End
            Begin VB.TextBox TxtFumigableQty 
               Alignment       =   1  'Right Justify
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
               Left            =   12720
               Locked          =   -1  'True
               TabIndex        =   40
               TabStop         =   0   'False
               Top             =   690
               Width           =   1485
            End
            Begin VB.ComboBox CmbCommodity 
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
               Left            =   4800
               Sorted          =   -1  'True
               TabIndex        =   34
               Top             =   240
               Width           =   2250
            End
            Begin VB.ComboBox CmbCategory 
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
               ItemData        =   "FrmMst_GodownInspection.frx":01BC
               Left            =   1770
               List            =   "FrmMst_GodownInspection.frx":01C9
               Sorted          =   -1  'True
               TabIndex        =   33
               Top             =   240
               Width           =   1515
            End
            Begin VB.TextBox TxtExpQtyP 
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
               TabIndex        =   35
               Top             =   240
               Width           =   765
            End
            Begin VB.CommandButton CmdCDeleteFromList 
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
               Height          =   360
               Left            =   12390
               TabIndex        =   44
               Top             =   1530
               Width           =   1935
            End
            Begin VB.CommandButton CmdCAddToList 
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
               Height          =   360
               Left            =   12390
               TabIndex        =   43
               Top             =   1170
               Width           =   1935
            End
            Begin VB.TextBox TxtStorageRate 
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
               Left            =   1770
               TabIndex        =   36
               Top             =   690
               Width           =   1050
            End
            Begin VB.TextBox TxtPrice 
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
               Left            =   4800
               TabIndex        =   37
               Top             =   690
               Width           =   2250
            End
            Begin VB.TextBox TxtRevenue 
               Alignment       =   1  'Right Justify
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
               Left            =   8160
               Locked          =   -1  'True
               TabIndex        =   38
               TabStop         =   0   'False
               Top             =   690
               Width           =   1740
            End
            Begin VB.TextBox TxtAUM 
               Alignment       =   1  'Right Justify
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
               Left            =   11040
               Locked          =   -1  'True
               TabIndex        =   39
               TabStop         =   0   'False
               Top             =   690
               Width           =   1605
            End
            Begin VB.TextBox TxtStorageRemarks 
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
               Height          =   735
               Left            =   4800
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   42
               Top             =   1125
               Width           =   7410
            End
            Begin VB.TextBox TxtExpQtyMT 
               Alignment       =   1  'Right Justify
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
               Left            =   11040
               Locked          =   -1  'True
               TabIndex        =   57
               TabStop         =   0   'False
               Top             =   240
               Width           =   1605
            End
            Begin VB.Label Label75 
               AutoSize        =   -1  'True
               Caption         =   "CM Share(in Rs.)"
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
               Left            =   105
               TabIndex        =   199
               Top             =   1575
               Width           =   1500
            End
            Begin VB.Label Label74 
               AutoSize        =   -1  'True
               Caption         =   "CM Share(in %)"
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
               TabIndex        =   197
               Top             =   1185
               Width           =   1380
            End
            Begin VB.Label Label67 
               Caption         =   "Fumigable Qty ( in MT)"
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
               Left            =   12795
               TabIndex        =   183
               Top             =   180
               Width           =   1335
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               Caption         =   "Category"
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
               TabIndex        =   140
               Top             =   300
               Width           =   825
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
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
               Left            =   3480
               TabIndex        =   139
               Top             =   300
               Width           =   1020
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               Caption         =   "Expected Qty (in %)"
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
               Left            =   7200
               TabIndex        =   138
               Top             =   300
               Width           =   1740
            End
            Begin VB.Label Label32 
               Caption         =   "Storage Rate per month per MT"
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
               TabIndex        =   137
               Top             =   645
               Width           =   1620
            End
            Begin VB.Label Label34 
               AutoSize        =   -1  'True
               Caption         =   "Price (per MT)"
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
               Left            =   3480
               TabIndex        =   136
               Top             =   750
               Width           =   1275
            End
            Begin VB.Label Label36 
               AutoSize        =   -1  'True
               Caption         =   "Revenue"
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
               Left            =   7200
               TabIndex        =   135
               Top             =   750
               Width           =   825
            End
            Begin VB.Label Label39 
               AutoSize        =   -1  'True
               Caption         =   "AUM"
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
               TabIndex        =   134
               Top             =   750
               Width           =   450
            End
            Begin VB.Label Label56 
               Caption         =   "Storage Rate Remarks"
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
               Left            =   3480
               TabIndex        =   133
               Top             =   1215
               Width           =   1215
            End
            Begin VB.Label Label57 
               AutoSize        =   -1  'True
               Caption         =   "Qty (in MT)"
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
               TabIndex        =   132
               Top             =   300
               Width           =   960
            End
         End
         Begin VB.TextBox TxtExpUtilisationMT 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Height          =   315
            Left            =   10440
            Locked          =   -1  'True
            TabIndex        =   23
            Top             =   930
            Width           =   1245
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
            Height          =   1455
            Left            =   120
            TabIndex        =   83
            Top             =   3480
            Width           =   14370
            _ExtentX        =   25347
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
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
            Height          =   2775
            Left            =   -74880
            TabIndex        =   95
            Top             =   2520
            Width           =   14370
            _ExtentX        =   25347
            _ExtentY        =   4895
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
            Height          =   2415
            Left            =   -74880
            TabIndex        =   118
            Top             =   2880
            Width           =   14370
            _ExtentX        =   25347
            _ExtentY        =   4260
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
         Begin VB.TextBox Text12 
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
            Left            =   -62400
            Locked          =   -1  'True
            TabIndex        =   96
            TabStop         =   0   'False
            Top             =   2760
            Visible         =   0   'False
            Width           =   1605
         End
         Begin VB.TextBox Text13 
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
            Left            =   -62400
            Locked          =   -1  'True
            TabIndex        =   98
            TabStop         =   0   'False
            Top             =   3240
            Visible         =   0   'False
            Width           =   1605
         End
         Begin VB.TextBox Text14 
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
            Left            =   -62400
            Locked          =   -1  'True
            TabIndex        =   100
            TabStop         =   0   'False
            Top             =   3720
            Visible         =   0   'False
            Width           =   1605
         End
         Begin VB.TextBox Text15 
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
            Left            =   -62400
            Locked          =   -1  'True
            TabIndex        =   102
            TabStop         =   0   'False
            Top             =   4200
            Visible         =   0   'False
            Width           =   1605
         End
         Begin VB.TextBox Text16 
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
            Left            =   -62400
            Locked          =   -1  'True
            TabIndex        =   104
            TabStop         =   0   'False
            Top             =   4680
            Visible         =   0   'False
            Width           =   1605
         End
         Begin VB.TextBox Text17 
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
            Left            =   -62400
            Locked          =   -1  'True
            TabIndex        =   106
            TabStop         =   0   'False
            Top             =   5160
            Visible         =   0   'False
            Width           =   1605
         End
         Begin VB.TextBox Text11 
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
            Left            =   -62400
            Locked          =   -1  'True
            TabIndex        =   108
            TabStop         =   0   'False
            Top             =   6120
            Visible         =   0   'False
            Width           =   1605
         End
         Begin VB.TextBox Text10 
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
            Left            =   -62400
            Locked          =   -1  'True
            TabIndex        =   109
            TabStop         =   0   'False
            Top             =   5280
            Visible         =   0   'False
            Width           =   1605
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid5 
            Height          =   4335
            Left            =   -74880
            TabIndex        =   119
            Top             =   600
            Width           =   14370
            _ExtentX        =   25347
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
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid6 
            Height          =   1710
            Left            =   120
            TabIndex        =   166
            Top             =   645
            Width           =   5940
            _ExtentX        =   10478
            _ExtentY        =   3016
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
         Begin VB.TextBox TxtStaffUser 
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
            Left            =   2895
            TabIndex        =   53
            Top             =   1470
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.TextBox TxtSecStaffUser 
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
            Left            =   2895
            TabIndex        =   55
            Top             =   1860
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.TextBox TxtSecStaffSalaryUser 
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
            Left            =   4680
            TabIndex        =   56
            Top             =   1860
            Visible         =   0   'False
            Width           =   1365
         End
         Begin VB.TextBox TxtStaffSalaryUser 
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
            Left            =   4680
            TabIndex        =   54
            Top             =   1470
            Visible         =   0   'False
            Width           =   1365
         End
         Begin VB.TextBox TxtSecStaffSalarySys 
            Alignment       =   1  'Right Justify
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
            Left            =   4680
            Locked          =   -1  'True
            TabIndex        =   156
            Top             =   1050
            Visible         =   0   'False
            Width           =   1365
         End
         Begin VB.TextBox TxtSecStaffSys 
            Alignment       =   1  'Right Justify
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
            Left            =   2895
            Locked          =   -1  'True
            TabIndex        =   157
            Top             =   1050
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.TextBox TxtStaffsys 
            Alignment       =   1  'Right Justify
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
            Left            =   2895
            Locked          =   -1  'True
            TabIndex        =   158
            Top             =   660
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.TextBox TxtStaffSalarySys 
            Alignment       =   1  'Right Justify
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
            Left            =   4680
            Locked          =   -1  'True
            TabIndex        =   162
            Top             =   660
            Visible         =   0   'False
            Width           =   1365
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid7 
            Height          =   4290
            Left            =   -71040
            TabIndex        =   187
            Top             =   840
            Width           =   10530
            _ExtentX        =   18574
            _ExtentY        =   7567
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
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid9 
            Height          =   3495
            Left            =   -74880
            TabIndex        =   222
            Top             =   1800
            Width           =   14370
            _ExtentX        =   25347
            _ExtentY        =   6165
            _Version        =   393216
            Rows            =   3
            FixedRows       =   2
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
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid10 
            Height          =   4575
            Left            =   -74880
            TabIndex        =   240
            Top             =   720
            Width           =   14370
            _ExtentX        =   25347
            _ExtentY        =   8070
            _Version        =   393216
            Rows            =   3
            FixedRows       =   2
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
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid11 
            Height          =   4695
            Left            =   -74880
            TabIndex        =   241
            Top             =   600
            Width           =   14370
            _ExtentX        =   25347
            _ExtentY        =   8281
            _Version        =   393216
            Rows            =   3
            FixedRows       =   2
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
         Begin VB.Label Label44 
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
            Left            =   -74880
            TabIndex        =   239
            Top             =   2025
            Width           =   885
         End
         Begin VB.Label Label46 
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
            Left            =   -74880
            TabIndex        =   238
            Top             =   1500
            Width           =   855
         End
         Begin VB.Label Label79 
            AutoSize        =   -1  'True
            Caption         =   "Rent Change in %"
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
            TabIndex        =   213
            Top             =   4140
            Width           =   1590
         End
         Begin VB.Label Label76 
            AutoSize        =   -1  'True
            Caption         =   "CM Share"
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
            TabIndex        =   201
            Top             =   3660
            Width           =   885
         End
         Begin VB.Label Label70 
            Caption         =   "Current Capacity Utilisation (in %)"
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
            Left            =   -74760
            TabIndex        =   189
            Top             =   788
            Width           =   1560
         End
         Begin VB.Label Label69 
            Caption         =   "color indicates either zero stock or zero rent"
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
            Left            =   12360
            TabIndex        =   186
            Top             =   1200
            Width           =   1935
         End
         Begin VB.Label Label68 
            BackColor       =   &H000000FF&
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Left            =   11880
            TabIndex        =   185
            Top             =   1290
            Width           =   375
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "Direct Expense"
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
            TabIndex        =   182
            Top             =   1860
            Width           =   1365
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Total Revenue"
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
            TabIndex        =   180
            Top             =   1380
            Width           =   1335
         End
         Begin VB.Label Label66 
            AutoSize        =   -1  'True
            Caption         =   "Indirect Expense"
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
            TabIndex        =   178
            Top             =   2790
            Width           =   1485
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            Caption         =   "(in %)"
            Height          =   195
            Left            =   9870
            TabIndex        =   175
            Top             =   675
            Width           =   375
         End
         Begin VB.Label lblaverageleasedrent 
            AutoSize        =   -1  'True
            Caption         =   "Average Leased Rent : 0"
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
            TabIndex        =   171
            Top             =   480
            Visible         =   0   'False
            Width           =   2220
         End
         Begin VB.Label Label65 
            AutoSize        =   -1  'True
            Caption         =   "Gross Profit"
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
            TabIndex        =   168
            Top             =   2310
            Width           =   1035
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Net Profit"
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
            TabIndex        =   167
            Top             =   3225
            Width           =   810
         End
         Begin VB.Label Label63 
            Caption         =   "As per User"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   120
            TabIndex        =   165
            Top             =   1560
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.Label Label62 
            Caption         =   "From System"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   120
            TabIndex        =   164
            Top             =   720
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.Label Label61 
            AutoSize        =   -1  'True
            Caption         =   "Salary"
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
            TabIndex        =   163
            Top             =   720
            Visible         =   0   'False
            Width           =   585
         End
         Begin VB.Label Label60 
            AutoSize        =   -1  'True
            Caption         =   "No. of security staff"
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
            Left            =   1080
            TabIndex        =   161
            Top             =   1110
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.Label Label59 
            AutoSize        =   -1  'True
            Caption         =   "No. of staff"
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
            Left            =   1080
            TabIndex        =   160
            Top             =   720
            Visible         =   0   'False
            Width           =   930
         End
         Begin VB.Label Label55 
            AutoSize        =   -1  'True
            Caption         =   "Salary"
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
            TabIndex        =   159
            Top             =   1110
            Visible         =   0   'False
            Width           =   585
         End
         Begin VB.Label Label47 
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
            TabIndex        =   155
            Top             =   2857
            Width           =   945
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Salary"
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
            TabIndex        =   154
            Top             =   1530
            Visible         =   0   'False
            Width           =   585
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            Caption         =   "Qty (in MT)"
            Height          =   195
            Left            =   10665
            TabIndex        =   153
            Top             =   675
            Width           =   780
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            Caption         =   "Salary"
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
            TabIndex        =   129
            Top             =   1920
            Visible         =   0   'False
            Width           =   585
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "No. of staff"
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
            Left            =   1080
            TabIndex        =   128
            Top             =   1530
            Visible         =   0   'False
            Width           =   930
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            Caption         =   "Expected Capacity Utilisation"
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
            Left            =   6600
            TabIndex        =   121
            Top             =   990
            Width           =   2505
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Additional staff"
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
            Left            =   -64800
            TabIndex        =   113
            Top             =   5340
            Visible         =   0   'False
            Width           =   1305
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Expected monthly CTC : x * 100000"
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
            Left            =   -68280
            TabIndex        =   112
            Top             =   5340
            Visible         =   0   'False
            Width           =   3060
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Expected monthly salary : x * 5500"
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
            Left            =   -68280
            TabIndex        =   111
            Top             =   6180
            Visible         =   0   'False
            Width           =   3000
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Additional Security "
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
            Left            =   -64800
            TabIndex        =   110
            Top             =   6180
            Visible         =   0   'False
            Width           =   1710
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "Dunnage rate"
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
            Left            =   -64800
            TabIndex        =   107
            Top             =   5220
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Spraying rate"
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
            Left            =   -64800
            TabIndex        =   105
            Top             =   4740
            Visible         =   0   'False
            Width           =   1200
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Fumigation rate"
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
            Left            =   -64800
            TabIndex        =   103
            Top             =   4260
            Visible         =   0   'False
            Width           =   1380
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Monthly Insurance (%)"
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
            Left            =   -64800
            TabIndex        =   101
            Top             =   3780
            Visible         =   0   'False
            Width           =   1950
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Franchisee share (%)"
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
            Left            =   -64800
            TabIndex        =   99
            Top             =   3300
            Visible         =   0   'False
            Width           =   1905
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Lease Rent"
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
            Left            =   -64800
            TabIndex        =   97
            Top             =   2820
            Visible         =   0   'False
            Width           =   1035
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "No. of security staff"
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
            Left            =   1080
            TabIndex        =   94
            Top             =   1920
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.Label LblStorageDays 
            AutoSize        =   -1  'True
            BackColor       =   &H80000000&
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
            Left            =   9480
            TabIndex        =   84
            Top             =   720
            Width           =   75
         End
      End
   End
   Begin MSScriptControlCtl.ScriptControl SC 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
End
Attribute VB_Name = "FrmMst_GodownInspection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mLocationID, mCMPID, mGodownID, mGodownTypeID, mCommodityID, mExpenseTypeID, mGodownCategoryID, mLicenseID As Variant
Dim mNCDEXPer As Double, mNNNRPer As Double, mReservationPer As Double
Dim Edit_Flg As String
Dim mGodownIDs As String, mEmployeeNos As String
Dim mStartDate, mEndDate As Variant
Dim mColsel As Single, mRowsel As Single
Dim mFormula As String, mFumigationFlag As String, mUserInputAmount As Integer, mType As String, mUserInputCalc As Integer, mClearFlag As Integer
Dim mAgreementExpiryDate As Variant
Dim mApprovalmailIDs As String
Dim mSaveFail As Boolean
Dim mLExpenseID, mCMPStatrDate As Variant

Dim mStateID, mmGodownID As Variant
Const strChecked = "þ"
Const strUnChecked = "q"
Dim mLength, mWidth, mHeight As Variant
Dim mSelRow, mSelCol, mSGodownID As Variant
Dim mBackMonths As Variant

Private Sub Grid_Head()
With MSHFlexGrid1
   .Clear
   .Cols = 27 '30 '27
   .Rows = 2
  
   .TextMatrix(0, 0) = "Viability ID"
   .TextMatrix(0, 1) = "Type"
   .TextMatrix(0, 2) = "Location Name"
   .TextMatrix(0, 3) = "CMP Name"
   .TextMatrix(0, 4) = "Godown No"
   .TextMatrix(0, 5) = "Capacity (in MT)"
   .TextMatrix(0, 6) = "Godown Type"
   .TextMatrix(0, 7) = "Area (in Sq.Ft)"
   .TextMatrix(0, 8) = "Address"
   .TextMatrix(0, 9) = "Status"
   .TextMatrix(0, 10) = "Start Date"
   .TextMatrix(0, 11) = "Close Date"
   .TextMatrix(0, 12) = "No. of Staff"
   .TextMatrix(0, 13) = "Staff Salary"
   .TextMatrix(0, 14) = "No. of Sec Staff"
   .TextMatrix(0, 15) = "Sec Staff Salary"
   .TextMatrix(0, 16) = "Exp. Utilisation(%)"
   .TextMatrix(0, 17) = "NCDEX Stock(%)"
   .TextMatrix(0, 18) = "NNNR Stock(%)"
   .TextMatrix(0, 19) = "Reservation Stock(%)"
   .TextMatrix(0, 20) = "Total Revenue"
   .TextMatrix(0, 21) = "Direct Expense"
   .TextMatrix(0, 22) = "Gross Profit"
   .TextMatrix(0, 23) = "Indirect Expense"
   .TextMatrix(0, 24) = "Net Profit"
   .TextMatrix(0, 25) = "Created Date"
   .TextMatrix(0, 26) = "Updated Date"
   
   .ColWidth(0) = 1200
   .ColWidth(1) = 1500
   .ColWidth(2) = 1800
   .ColWidth(3) = 1800
   .ColWidth(4) = 1500
   .ColWidth(5) = 3000
   .ColWidth(6) = 1500
   .ColWidth(7) = 1500
   
End With
End Sub
Private Sub Grid_Head8()
With MSHFlexGrid8
   .Clear
   .Cols = 4
   .Rows = 2
   .FixedCols = 0
   .TextMatrix(0, 0) = ""
   .TextMatrix(0, 1) = "License Type"
   .TextMatrix(0, 2) = "License Type ID"
   .TextMatrix(0, 3) = "License ID"
   
   .ColWidth(0) = 1200
   .ColWidth(1) = 2000
   .ColWidth(2) = 0
   .ColWidth(3) = 2000
End With
End Sub

Private Sub Grid_Head9()
With MSHFlexGrid9
   .Clear
   .Cols = 5
   .Rows = 3
   .FixedCols = 0
   .FixedRows = 2
   .TextMatrix(1, 0) = "ID"
   .TextMatrix(1, 1) = "Length(ft.in)"
   .TextMatrix(1, 2) = "Width(ft.in)"
   .TextMatrix(1, 3) = "Height(ft.in)"
   .TextMatrix(1, 4) = "Area"
   
   
   .ColWidth(0) = 1200
   .ColWidth(1) = 2000
   .ColWidth(2) = 2000
   .ColWidth(3) = 2000
   .ColWidth(4) = 2000
   
End With
End Sub

Private Sub Grid_Head10()
With MSHFlexGrid10
   .Clear
   .Cols = 5
   .Rows = 3
   .FixedCols = 0
   .FixedRows = 2
   
   
   .TextMatrix(1, 0) = "Sr No"
   .TextMatrix(1, 1) = "Month Year"
   .TextMatrix(1, 2) = "Capacity in MT"
   .TextMatrix(1, 3) = "Balance Stock at end of the month"
   .TextMatrix(1, 4) = "Utilization in MT ((Stock/Capacity)*100)"
   
   .WordWrap = True
   .ColWidth(0) = 1200
   .ColWidth(1) = 2000
   .ColWidth(2) = 2000
   .ColWidth(3) = 2000
   .ColWidth(4) = 4500
   
End With
End Sub
Private Sub Grid_Head11()
With MSHFlexGrid11
   .Clear
   .Cols = 6
   .Rows = 3
   .FixedCols = 0
   .FixedRows = 2

   .TextMatrix(1, 0) = "SR No"
   .TextMatrix(1, 1) = "Month Year"
   .TextMatrix(1, 2) = "Total Income"
   .TextMatrix(1, 3) = "Total Expense"
   .TextMatrix(1, 4) = "Net Profit / Loss"
   .TextMatrix(1, 5) = "Rent"
   
   .WordWrap = True
   .ColWidth(0) = 1200
   .ColWidth(1) = 2000
   .ColWidth(2) = 2000
   .ColWidth(3) = 2000
   .ColWidth(4) = 2000
   
End With
End Sub

Private Sub Grid_Head2()
If Edit_Flg = "A" Then
   Call Grid_Head6
   'TxtStaffUser.Text = ""
   'TxtStaffSalaryUser.Text = ""
   'TxtSecStaffUser.Text = ""
   'TxtSecStaffSalaryUser.Text = ""
   'TxtStaffsys.Text = ""
   'TxtStaffSalarySys.Text = ""
   'TxtSecStaffSys.Text = ""
   'TxtSecStaffSalarySys.Text = ""
End If
With MSHFlexGrid2
   .Clear
   .Rows = 3
   .FixedRows = 2
   .Cols = 9
   .FixedCols = 1
   .TextMatrix(0, 1) = "Total No of Godown : xxxx"
   .TextMatrix(0, 2) = "Average Rent (per Sq.Ft) : xxxx"
     
   .TextMatrix(1, 0) = "CMP Name"
   .TextMatrix(1, 1) = "Godown No"
   .TextMatrix(1, 2) = "Godown Type"
   .TextMatrix(1, 3) = "Capacity (in MT)"
   .TextMatrix(1, 4) = "Stock (in MT)"
   .TextMatrix(1, 5) = "Area (in Sq.Ft)"
   .TextMatrix(1, 6) = "Rent"
   .TextMatrix(1, 7) = "Address"
   .TextMatrix(1, 8) = "GodownID"
   
   .ColWidth(0) = 3000
   .ColWidth(1) = 1500
   .ColWidth(2) = 1700
   .ColWidth(3) = 1100
   .ColWidth(4) = 1000
   .ColWidth(5) = 1000
   .ColWidth(6) = 1000
   .ColWidth(7) = 3500
   .ColWidth(8) = 0
   
   .RowHeight(0) = 600
   .RowHeight(1) = 600
End With
End Sub

Private Sub Grid_Head3()
mNCDEXPer = 0
mNNNRPer = 0
mReservationPer = 0
With MSHFlexGrid3
   .Clear
   .Rows = 3
   .Cols = 15
   .FixedRows = 2
  
   .TextMatrix(1, 0) = "Category"
   .TextMatrix(1, 1) = "Commodity"
   .TextMatrix(1, 2) = "Expected Qty (in %)"
   .TextMatrix(1, 3) = "Expected Qty (in MT)"
   .TextMatrix(1, 4) = "Storage Rate per month per MT"
   .TextMatrix(1, 5) = "Price"
   .TextMatrix(1, 6) = "Revenue"
   .TextMatrix(1, 7) = "AUM"
   .TextMatrix(1, 8) = "Storage Rate Remarks"
   .TextMatrix(1, 9) = "InspectionCommodityID"
   .TextMatrix(1, 10) = "CommodityID"
   .TextMatrix(1, 11) = "Fumigable Qty (in MT)"
   .TextMatrix(1, 12) = "FumigationFlag"
   .TextMatrix(1, 13) = "CM Share (in %)"
   .TextMatrix(1, 14) = "CM Share (in Rs.)"
         
   .ColWidth(0) = 1200
   .ColWidth(1) = 1200
   .ColWidth(2) = 1000
   .ColWidth(3) = 1000
   .ColWidth(4) = 1500
   .ColWidth(5) = 800
   .ColWidth(6) = 900
   .ColWidth(7) = 1400
   .ColWidth(8) = 1800
   .ColWidth(9) = 0
   .ColWidth(10) = 0
   .ColWidth(11) = 1000
   .ColWidth(12) = 0
   .ColWidth(13) = 1200
   .ColWidth(14) = 1200
   
   .RowHeight(1) = 600
End With
Call Grid_Head7
End Sub

Private Sub Grid_Head4()
With MSHFlexGrid4
   .Clear
   .Rows = 3
   .Cols = 17
   .FixedRows = 2
  
   .TextMatrix(1, 0) = "Expense Type"
   .TextMatrix(1, 1) = "Calculation On"
   .TextMatrix(1, 2) = "Units (Qty/ Sq.Ft/No of)" '"Qty/Sq.Ft/No of"
   .TextMatrix(1, 3) = "Amount"
   .TextMatrix(1, 4) = "Remarks"
   .TextMatrix(1, 5) = "InspectionExpenseID"
   .TextMatrix(1, 6) = "InsExpTypeID"
   .TextMatrix(1, 7) = "Amount Per Unit"
   .TextMatrix(1, 8) = "Formula"
   .TextMatrix(1, 9) = "User Input Amount"
   .TextMatrix(1, 10) = "Type"
   .TextMatrix(1, 11) = "User Input Calculation"
   .TextMatrix(1, 12) = "Clear Flag"
   .TextMatrix(1, 13) = "Expense Remarks"
   .TextMatrix(1, 14) = "Previous Renewal rent"
   .TextMatrix(1, 15) = "Previous Lease rent"
   .TextMatrix(1, 16) = "Change in %"
      
   .ColWidth(0) = 3500
   .ColWidth(1) = 1800
   .ColWidth(2) = 1200
   .ColWidth(3) = 1100
   .ColWidth(4) = 4500
   .ColWidth(5) = 0
   .ColWidth(6) = 0
   .ColWidth(7) = 1200
   .ColWidth(8) = 0
   .ColWidth(9) = 0
   .ColWidth(10) = 0
   .ColWidth(11) = 0
   .ColWidth(12) = 0
   .ColWidth(14) = 1200
   .ColWidth(15) = 1200
   .ColWidth(16) = 1200
   
   .RowHeight(1) = 600
End With
End Sub

Private Sub Grid_Head5()
With MSHFlexGrid5
   .Clear
   .Rows = 3
   .FixedRows = 2
   .Cols = 9
  
   .TextMatrix(0, 1) = "Total No of Godown : xxxx"
   .TextMatrix(0, 2) = "Average Rent (per Sq.Ft) : xxxx"
  
   .TextMatrix(1, 0) = "CMP Name"
   .TextMatrix(1, 1) = "Godown No"
   .TextMatrix(1, 2) = "Godown Type"
   .TextMatrix(1, 3) = "Capacity"
   .TextMatrix(1, 4) = "Stock"
   .TextMatrix(1, 5) = "Area (in Sq.Ft)"
   .TextMatrix(1, 6) = "Rent"
   .TextMatrix(1, 7) = "Address"
   .TextMatrix(1, 8) = "GodownID"
   
   .ColWidth(0) = 3000
   .ColWidth(1) = 1500
   .ColWidth(2) = 1700
   .ColWidth(3) = 1100
   .ColWidth(4) = 1000
   .ColWidth(5) = 1000
   .ColWidth(6) = 1000
   .ColWidth(7) = 3500
   .ColWidth(8) = 0
      
   .RowHeight(0) = 600
   .RowHeight(1) = 600
End With
End Sub

Private Sub Grid_Head6()
With MSHFlexGrid6
   .Clear
   .Rows = 4
   .FixedRows = 2
   .Cols = 5
   .FixedCols = 3
   .WordWrap = True
   .MergeCells = flexMergeRestrictRows
  
   .TextMatrix(2, 0) = "Employee"
   .TextMatrix(3, 0) = "Security"
  
   .TextMatrix(0, 1) = "Existing for the Location"
   .TextMatrix(0, 2) = "Existing for the Location"
   .TextMatrix(0, 3) = "Existing + Proposed for the Location"
   .TextMatrix(0, 4) = "Existing + Proposed for the Location"
   
   .TextMatrix(1, 1) = "No of Staff"
   .TextMatrix(1, 2) = "Salary Cost"
   .TextMatrix(1, 3) = "No of Staff"
   .TextMatrix(1, 4) = "Salary Cost"
   
   .MergeRow(0) = True
      
   .RowHeight(0) = 600
   .RowHeight(1) = 600
End With
End Sub

Private Sub Grid_Head7()
With MSHFlexGrid7
   .Clear
   .Rows = 11
   .FixedRows = 1
   .Cols = 6
   .FixedCols = 1
   .WordWrap = True
   
   .TextMatrix(0, 0) = "Expected Capacity Utilisation(in %)"
   .TextMatrix(0, 1) = "Total Revenue"
   .TextMatrix(0, 2) = "Direct Expense"
   .TextMatrix(0, 3) = "Gross Profit"
   .TextMatrix(0, 4) = "Indirect Expense"
   .TextMatrix(0, 5) = "Net Profit"

   .TextMatrix(1, 0) = "10"
   .TextMatrix(2, 0) = "20"
   .TextMatrix(3, 0) = "30"
   .TextMatrix(4, 0) = "40"
   .TextMatrix(5, 0) = "50"
   .TextMatrix(6, 0) = "60"
   .TextMatrix(7, 0) = "70"
   .TextMatrix(8, 0) = "80"
   .TextMatrix(9, 0) = "90"
   .TextMatrix(10, 0) = "100"
   
   .ColWidth(0) = 2200
   .ColWidth(1) = 1600
   .ColWidth(2) = 1600
   .ColWidth(3) = 1600
   .ColWidth(4) = 1600
   .ColWidth(5) = 1600
   
   .RowHeight(0) = 500
End With
End Sub


Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)

Call ClearControl
Call ClearCommodityDetail
Call ClearExpenseDetail
Call ClearAreaDetail
Call ComboVisible(True)
Call FillExpenseDetails
Frame6.Enabled = False

TxtStatus.Text = "Pending"
CmbStatus.Text = TxtStatus.Text
TxtStatus.Visible = True
CmbStatus.Visible = False

Picture2.Enabled = True
Frame1.Enabled = True
Frame2.Enabled = True
Frame5.Visible = False

'TxtStaffSalaryUser.Locked = False
'TxtStaffUser.Locked = False
'TxtSecStaffUser.Locked = False
'TxtSecStaffSalaryUser.Locked = False



TxtRemarks.Locked = False
TxtCapacity.Locked = False
TxtAddress.Locked = False
TxtLength.Locked = False
TxtWidth.Locked = False
TxtHeight.Locked = False
'TxtArea.Locked = False
TxtGodown.Locked = False
TxtExpUtilisation.Locked = False
TxtNCDEXStock.Locked = False
TxtNNNRStock.Locked = False
TxtReservationStock.Locked = False
TxtStartDate.Enabled = True
TxtCloseDate.Enabled = True
TxtNoofStack.Locked = False
TxtWeightPerStack.Locked = False

If RsMst_GodownInspection.RecordCount > 0 Then
   CmbType.SetFocus
End If

End Sub

Private Sub FillExpenseDetails()

Call ClearExpenseDetail
tmp = " Select  MIE.InsExpTypeID , MIE.ExpenseType, MIE.CalculationOn, MIE.Amount, MIE.Formula, MIE.UserInputAmount, MIE.Type, MIE.UserInputCalc, MIE.ClearFlag, MIE.Remarks 'ExpenseRemarks' "
tmp = tmp & " From Mst_InspectionExpenseType MIE"
tmp = tmp & " Where MIE.Status = 'A'"
If CmbGodownType.Text <> "" Then
   tmp = tmp & " And (MIE.UnitTypeID is null OR MIE.UnitTypeID = " & mGodownTypeID & ") "
End If
If Edit_Flg <> "A" Then
   MSHFlexGrid4.TextMatrix(0, 3) = Format(Val(MSHFlexGrid4.TextMatrix(0, 3)) - Val(MSHFlexGrid4.TextMatrix(MSHFlexGrid4.Rows - 2, 3)), "#0")
   MSHFlexGrid4.TextMatrix(0, 3) = Format(Val(MSHFlexGrid4.TextMatrix(0, 3)) - Val(MSHFlexGrid4.TextMatrix(MSHFlexGrid4.Rows - 3, 3)), "#0")
   MSHFlexGrid4.Rows = MSHFlexGrid4.Rows - 2
   tmp = tmp & " And MIE.InsExpTypeID not in ( Select InsExpTypeID from Mst_InspectionExpenseDetail Where InspectionID = " & Val(TxtInspectionID.Text) & " )   "
End If
tmp = tmp & " Order by MIE.OrderNo"
Call Tmp7Table
With MSHFlexGrid4
   If TmpRs7.RecordCount > 0 Then
      For j = 1 To TmpRs7.RecordCount
         I = .Rows - 1
         .TextMatrix(I, 0) = IIf(IsNull(TmpRs7!ExpenseType), "", TmpRs7!ExpenseType)
         .TextMatrix(I, 1) = IIf(IsNull(TmpRs7!CalculationOn), "", TmpRs7!CalculationOn)
         .TextMatrix(I, 2) = "" 'IIf(IsNull(TmpRs7!QtySQ), "", TmpRs7!QtySQ)
         .TextMatrix(I, 3) = "" 'IIf(IsNull(TmpRs7!Amount), "", TmpRs7!Amount)
         .TextMatrix(I, 4) = "" 'IIf(IsNull(TmpRs7!Remarks), "", TmpRs7!Remarks)
         .TextMatrix(I, 5) = "" 'IIf(IsNull(TmpRs7!InspectionExpenseID), "", TmpRs7!InspectionExpenseID)
         .TextMatrix(I, 6) = IIf(IsNull(TmpRs7!InsExpTypeID), "", TmpRs7!InsExpTypeID)
         .TextMatrix(I, 7) = IIf(IsNull(TmpRs7!Amount), "", TmpRs7!Amount)
         .TextMatrix(I, 8) = IIf(IsNull(TmpRs7!Formula), "", TmpRs7!Formula)
         .TextMatrix(I, 9) = IIf(IsNull(TmpRs7!UserInputAmount), "", TmpRs7!UserInputAmount)
         .TextMatrix(I, 10) = IIf(IsNull(TmpRs7!Type), "", TmpRs7!Type)
         .TextMatrix(I, 11) = IIf(IsNull(TmpRs7!UserInputCalc), "", TmpRs7!UserInputCalc)
         .TextMatrix(I, 12) = IIf(IsNull(TmpRs7!ClearFlag), "", TmpRs7!ClearFlag)
         .TextMatrix(I, 13) = IIf(IsNull(TmpRs7!ExpenseRemarks), "", TmpRs7!ExpenseRemarks)
         .Rows = .Rows + 1
         TmpRs7.MoveNext
      Next
   End If
   
'============================for security and staff salary expense detail============================
Call SecandUserSalaryExpense
'============================for security and staff salary expense detail============================
End With
End Sub

Private Sub ComboVisible(var1 As Boolean)
CmbType.Visible = var1
CmbLocation.Visible = var1
CmbCMP.Visible = var1
CmbGodown.Visible = var1
CmbGodownType.Visible = var1
CmbGodownCategory.Visible = var1
CmbLicenseNo.Visible = var1
CmbStatus.Visible = var1

If var1 = False Then
   TxtType.Visible = True
   TxtLocation.Visible = True
   TxtCmp.Visible = True
   TxtGodown.Visible = True
   TxtGodownType.Visible = True
   TxtGodownCategory.Visible = True
   TxtLicenseNo.Visible = True
   TxtStatus.Visible = True
Else
   TxtType.Visible = False
   TxtLocation.Visible = False
   TxtCmp.Visible = False
   TxtGodown.Visible = False
   TxtGodownType.Visible = False
   TxtGodownCategory.Visible = False
   TxtLicenseNo.Visible = False
   TxtStatus.Visible = False
End If
End Sub

Private Sub ClearCommodityDetail()
CmbCategory.Text = ""
CmbCommodity.Text = ""
TxtExpQtyP.Text = ""
TxtExpQtyMT.Text = ""
TxtStorageRate.Text = ""
TxtPrice.Text = ""
TxtRevenue.Text = ""
TxtAUM.Text = ""
TxtStorageRemarks.Text = ""
TxtFumigableQty.Text = ""
TxtCMShareP.Text = ""
TxtCMShareR.Text = ""
CmdCAddToList.Caption = "Add to List"
CmdCDeleteFromList.Enabled = False
CmbCategory.Enabled = True
CmbCommodity.Enabled = True
CmdStorageRate.Enabled = False
TxtCMShareP.Enabled = False
End Sub

Private Sub ClearExpenseDetail()
TxtExpenseType.Text = ""
TxtCalculationOn.Text = ""
TxtQtySq.Text = ""
TxtAmount.Text = ""
TxtAmountPer.Text = ""
TxtFormula.Text = ""
TxtExpenseRemarks.Text = ""
CmdEUpdateList.Enabled = False
CmdClearData.Enabled = False
TxtRemarksExpense.Text = ""
mLExpenseID = 0
LblLeaseRent.Caption = "--"
LblPRenewalRent.Caption = "--"
TxtChangeInPer.Text = ""
LblViewPrevRent.Enabled = False
'TxtQtySq.Locked = False
End Sub
Private Sub ClearAreaDetail()
TxtLength = ""
TxtWidth = ""
TxtHeight = ""
TxtSubArea = ""
CmdAreaAddToList.Caption = "Add to List"
CmdAreaDeleteFromList.Enabled = False
End Sub

Private Sub ClearControl()
Call Grid_Head2
Call Grid_Head3
Call Grid_Head4
Call Grid_Head5
Call Grid_Head6
Call Grid_Head7
Call Grid_Head9
Call Grid_Head10
Call Grid_Head11

TxtInspectionID.Text = ""
CmbType.Text = ""
TxtType.Text = ""
CmbLocation.Text = ""
TxtLocation.Text = ""
CmbCMP.Text = ""
TxtCmp.Text = ""
CmbGodown.Text = ""
TxtGodown.Text = ""
CmbGodownType.Text = ""
TxtGodownType.Text = ""
CmbGodownCategory.Text = ""
TxtGodownCategory.Text = ""
CmbLicenseNo.Text = ""
TxtLicenseNo.Text = ""
CmbStatus.Text = ""
TxtStatus.Text = ""

TxtLength = ""
TxtWidth = ""
TxtHeight = ""

TxtCapacity.Text = ""
TxtAddress.Text = ""
TxtArea.Text = ""

TxtStartDate.Value = Date
TxtCloseDate.Value = Date + 365
'TxtStaffUser.Text = ""
'TxtStaffSalaryUser.Text = ""
'TxtSecStaffUser.Text = ""
'TxtSecStaffSalaryUser.Text = ""
'TxtStaffsys.Text = ""
'TxtStaffSalarySys.Text = ""
'TxtSecStaffSys.Text = ""
'TxtSecStaffSalarySys.Text = ""
TxtRemarks.Text = ""
TxtTotalRevenue.Text = ""
TxtDirectExpense.Text = ""
TxtNetProfit.Text = ""
TxtGrossProfit.Text = ""
TxtIndirectExpense.Text = ""
TxtExpUtilisation.Text = ""
TxtNCDEXStock.Text = ""
TxtNNNRStock.Text = ""
TxtReservationStock.Text = ""
TxtExpUtilisationMT.Text = ""
TxtNCDEXMT.Text = ""
TxtNNNRMT.Text = ""
TxtReservationMT.Text = ""
TxtCreated.Text = ""
TxtUpdated.Text = ""
TxtCMShare.Text = ""
TxtSRentChange.Text = ""

End Sub

Private Sub ChkCloseDate_Click()
If ChkCloseDate.Value = vbChecked Then
   TxtCloseDate.Visible = True
   TxtCloseDate.Value = Date
Else
   TxtCloseDate.Visible = False
End If
End Sub

Private Sub Check1_Click()

If Check1.Value = 0 Then
   For I = 1 To MSHFlexGrid8.Rows - 1
       If MSHFlexGrid8.TextMatrix(I, 1) <> "" Then
          MSHFlexGrid8.row = I
          MSHFlexGrid8.Col = 0
          MSHFlexGrid8.CellFontName = "Wingdings"
          MSHFlexGrid8.CellFontSize = 12
          MSHFlexGrid8.CellAlignment = flexAlignCenterCenter
          MSHFlexGrid8.Text = strUnChecked
       End If
   Next
Else
   For I = 1 To MSHFlexGrid8.Rows - 1
       If MSHFlexGrid8.TextMatrix(I, 1) <> "" Then
          MSHFlexGrid8.row = I
          MSHFlexGrid8.Col = 0
          MSHFlexGrid8.CellFontName = "Wingdings"
          MSHFlexGrid8.CellFontSize = 12
          MSHFlexGrid8.CellAlignment = flexAlignCenterCenter
          MSHFlexGrid8.Text = strChecked
       End If
   Next
End If
End Sub

Private Sub CmbCategory_GotFocus()
tmp = CmbCategory.Text
End Sub

Private Sub CmbCategory_LostFocus()
If CmbCategory.Text = "" Then Exit Sub

'If LCase(CmbCategory.Text) <> "ncdex" And LCase(CmbCategory.Text) <> "nnnr" And LCase(CmbCategory.Text) <> "reservation" Then
'End If
If LCase(CmbCategory.Text) = "ncdex" Then
   If Val(TxtNCDEXMT) <= 0 Then
      MsgBox "NCDEX Qty is not available, you can not select!!! "
      CmbCategory.SetFocus
      Exit Sub
   End If
   CmdStorageRate.Enabled = True
   TxtCMShareP.Text = ""
   TxtCMShareP.Enabled = False
ElseIf LCase(CmbCategory.Text) = "nnnr" Then
   If Val(TxtNNNRMT) <= 0 Then
      MsgBox "NNNR Qty is not available, you can not select!!! "
      CmbCategory.SetFocus
      Exit Sub
   End If
   CmdStorageRate.Enabled = True
   TxtCMShareP.Enabled = True
ElseIf LCase(CmbCategory.Text) = "reservation" Then
   If Val(TxtReservationMT) <= 0 Then
      MsgBox "Reservation Qty is not available, you can not select!!! "
      CmbCategory.SetFocus
      Exit Sub
   End If
   CmdStorageRate.Enabled = False
   TxtCMShareP.Text = ""
   TxtCMShareP.Enabled = False
Else
   MsgBox "Invalid selection!!! "
   CmbCategory.SetFocus
   Exit Sub
End If
If tmp <> CmbCategory.Text Then
   TxtExpQtyP.Text = ""
   TxtExpQtyMT.Text = ""
   TxtStorageRate.Text = ""
   TxtPrice.Text = ""
   TxtRevenue.Text = ""
   TxtAUM.Text = ""
   TxtStorageRemarks.Text = ""
   TxtFumigableQty.Text = ""
   TxtCMShareP.Text = ""
   TxtCMShareR.Text = ""
End If
End Sub

Private Sub CmbCMP_GotFocus()
If Trim(CmbLocation.Text) = "" Then
   MsgBox "Select Location!!!"
   CmbLocation.SetFocus
   Exit Sub
End If
tmp = CmbCMP.Text
If Gen_WcCMP = "" Then
   Call TableMst_CMP(" where CloseDate is null and LocationID = " & mLocationID)
Else
   Call TableMst_CMP(" where " & Gen_WcCMP & " And CloseDate is null and LocationID = " & mLocationID)
End If
CmbCMP.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "No Record found!!!"
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
   mCMPStatrDate = ""
   Exit Sub
End If

If tmp <> CmbCMP.Text Then
   CmbGodown.Text = ""
End If

RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPName = '" & CmbCMP.Text & "'"
If RsMst_CMP.EOF Then
   MsgBox "Invalid selection!!!"
   CmbCMP.SetFocus
   Exit Sub
End If

mCMPID = RsMst_CMP!CMPID
mCMPStatrDate = RsMst_CMP!StartDate

End Sub

Private Sub CmbCommodity_GotFocus()
If Trim(CmbLocation.Text) = "" Then
   MsgBox "Select Location!!!"
   CmbLocation.SetFocus
   Exit Sub
End If
tmp = CmbCommodity.Text
If CmdCAddToList.Caption = "Update List" Then
   Call TableMst_Commodity(" Where CommodityID = " & mCommodityID & " OR CommodityID in (Select CommodityID from Mst_LocationCommodityMapping Where LocationID = " & mLocationID & ")")
Else
   Call TableMst_Commodity(" Where CommodityID in (Select CommodityID from Mst_LocationCommodityMapping Where LocationID = " & mLocationID & ")")
End If

CmbCommodity.Clear
If RsMst_Commodity.RecordCount = 0 Then
   MsgBox "No Commodity found"
   CmbCommodity.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Commodity.RecordCount
   CmbCommodity.AddItem RsMst_Commodity!Commodity
   RsMst_Commodity.MoveNext
Next
CmbCommodity.Text = tmp
End Sub

Private Sub CmbCommodity_LostFocus()
If CmbCommodity.Text = "" Then
   Exit Sub
End If
RsMst_Commodity.MoveFirst
RsMst_Commodity.Find "Commodity = '" & CmbCommodity.Text & "'"

If RsMst_Commodity.EOF Then
   MsgBox "Invalid Commodity selection "
   CmbCommodity.SetFocus
   Exit Sub
End If
mCommodityID = RsMst_Commodity!CommodityID
mFumigationFlag = IIf(IsNull(RsMst_Commodity!FumigationFlag), "", RsMst_Commodity!FumigationFlag)
'If LCase(mFumigationFlag) = "y" Then
'   TxtFumigableQty.Locked = False
'Else
'   TxtFumigableQty.Locked = True
'   TxtFumigableQty.Text = ""
'End If
If LCase(mFumigationFlag) <> "y" Then
   TxtFumigableQty.Text = ""
End If

End Sub

Private Sub CmbGodown_GotFocus()
If Trim(CmbCMP.Text) = "" Then
   MsgBox "Select CMP!!!"
   CmbCMP.SetFocus
   Exit Sub
End If
tmp = CmbGodown.Text
Call TableMst_Godown(" where CloseDate is null and CMPID = " & mCMPID)
CmbGodown.Clear
If RsMst_Godown.RecordCount = 0 Then
   MsgBox "No Record found!!!"
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
   TxtCapacity.Text = ""
   TxtArea.Text = ""
   TxtAddress.Text = ""
   CmbGodownType.Text = ""
   CmbGodownCategory.Text = ""
   CmbLicenseNo.Text = ""
   Call CapacityCalculation
   Exit Sub
End If

RsMst_Godown.MoveFirst
RsMst_Godown.Find "GodownNo= '" & CmbGodown.Text & "'"
If RsMst_Godown.EOF Then
   MsgBox "Invalid selection!!!"
   CmbGodown.SetFocus
   Exit Sub
End If

mGodownID = RsMst_Godown!GodownID
If tmp <> CmbGodown.Text Then
   TxtCapacity.Text = IIf(IsNull(RsMst_Godown!GodownCapacity), "0", RsMst_Godown!GodownCapacity)
   TxtArea.Text = IIf(IsNull(RsMst_Godown!Area), "0", RsMst_Godown!Area)
   TxtAddress.Text = IIf(IsNull(RsMst_Godown!Address), "", RsMst_Godown!Address)
   TxtNoofStack = IIf(IsNull(RsMst_Godown!NoOfStack), "", RsMst_Godown!NoOfStack)
   TxtWeightPerStack = IIf(IsNull(RsMst_Godown!WeightPerStack), "", RsMst_Godown!WeightPerStack)
   CmbGodownType.Text = GetUnitName(IIf(IsNull(RsMst_Godown!UnitTypeID), 0, RsMst_Godown!UnitTypeID))
   If LCase(CmbGodownType.Text) = "leased franchisee" Then
      MsgBox "Renewal of Leased Franchisee godowns is not allowed!!!"
      CmbGodown.SetFocus
      Exit Sub
   End If
   
   Call CmbGodownType_GotFocus
   Call CmbGodownType_LostFocus
   CmbGodownCategory.Text = GetCategoryName(IIf(IsNull(RsMst_Godown!CategoryID), 0, RsMst_Godown!CategoryID))
   Call CmbGodownCategory_GotFocus
   Call CmbGodownCategory_LostFocus
   CmbLicenseNo.Text = GetLicenseNo(IIf(IsNull(RsMst_Godown!LicenseID), 0, RsMst_Godown!LicenseID))
   Call CmbLicenseNo_GotFocus
   Call CmbLicenseNo_LostFocus
   tmp = " select TLA.ExpiryDate from Txn_GodownLessorAgreeDetail TGLA "
   tmp = tmp & " Inner Join Txn_LessorAgreement TLA on TLA.LAID = TGLA.LAID"
   tmp = tmp & " Where TLA.Flag in ('A', 'P') And TGLA.GodownID = " & mGodownID
   tmp = tmp & " Order by TLA.ExpiryDate Desc"
   Call TmpTable
   If TmpRs.RecordCount > 0 Then
      TxtStartDate.Value = IIf(IsNull(TmpRs!ExpiryDate), Date, TmpRs!ExpiryDate)
      mAgreementExpiryDate = TxtStartDate.Value
      TxtStartDate.Value = TxtStartDate.Value + 1
      TxtCloseDate.Value = TxtStartDate.Value + 365
   End If
   Call CapacityCalculation
   ''====================='for security and staff salary expense'=====================
   MSHFlexGrid4.TextMatrix(0, 3) = Format(Val(MSHFlexGrid4.TextMatrix(0, 3)) - Val(MSHFlexGrid4.TextMatrix(MSHFlexGrid4.Rows - 2, 3)), "#0")
   MSHFlexGrid4.TextMatrix(0, 3) = Format(Val(MSHFlexGrid4.TextMatrix(0, 3)) - Val(MSHFlexGrid4.TextMatrix(MSHFlexGrid4.Rows - 3, 3)), "#0")
   MSHFlexGrid4.Rows = MSHFlexGrid4.Rows - 2
   Call SecandUserSalaryExpense
   '====================='for security and staff salary expense'=====================
End If

End Sub

Private Sub CmbGodownType_GotFocus()
tmp = CmbGodownType.Text
Call TableMst_UnitType(" Where Status  = 'A'")
CmbGodownType.Clear
If RsMst_UnitType.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_UnitType.RecordCount
    CmbGodownType.AddItem RsMst_UnitType!UnitType
    RsMst_UnitType.MoveNext
Next
CmbGodownType.Text = tmp
End Sub

Private Sub CmbGodownType_LostFocus()
If CmbGodownType.Text = "" Then
   'Call Grid_Head4
   Exit Sub
End If

If tmp <> CmbGodownType.Text Then
   'Call Grid_Head4
End If

RsMst_UnitType.MoveFirst
RsMst_UnitType.Find "UnitType = '" & CmbGodownType.Text & "'"
If RsMst_UnitType.EOF Then
   MsgBox "Invalid selection!!!"
   CmbGodownType.SetFocus
   Exit Sub
End If

mGodownTypeID = RsMst_UnitType!UnitTypeID
Call Grid_Head4
Call FillExpenseDetails

Call TxtNoofStack_LostFocus


End Sub

Private Sub CmbGodownCategory_GotFocus()
tmp = CmbGodownCategory.Text
If Edit_Flg = "A" Then
   Call TableMst_Category("Where CategoryID <> 4")
Else
   Call TableMst_Category
End If
CmbGodownCategory.Clear
If RsMst_Category.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Category.RecordCount
    CmbGodownCategory.AddItem RsMst_Category!Category
    RsMst_Category.MoveNext
Next
CmbGodownCategory.Text = tmp
End Sub

Private Sub CmbGodownCategory_LostFocus()
If CmbGodownCategory.Text = "" Then
   Exit Sub
End If

If tmp <> CmbGodownCategory.Text Then
   ''write if clear date
End If

RsMst_Category.MoveFirst
RsMst_Category.Find "Category= '" & CmbGodownCategory.Text & "'"
If RsMst_Category.EOF Then
   MsgBox "Invalid selection!!!"
   CmbGodownCategory.SetFocus
   Exit Sub
End If

mGodownCategoryID = RsMst_Category!CategoryID

If mGodownCategoryID = 2 Then
'1 And 7
   If Val(mGodownTypeID) = 1 Or Val(mGodownTypeID) = 7 Then
      If LCase(CmbType.Text) = "renewal" Then
         
         TxtArea.Locked = True
         TxtArea.BackColor = &H80000004
         Call Grid_Head9
         SSTab1.Tab = 5
         If LCase(Gen_UserRole) = "power" Or LCase(Gen_UserRole) = "superadmin" Or LCase(Gen_UserRole) = "headcmg" Then
            Frame6.Enabled = True
         Else
            Frame6.Enabled = False
         End If
         Call GetPreviousInspectionAreaDetail
      Else
         TxtArea.Locked = True
         TxtArea.BackColor = &H80000004
         Frame6.Enabled = True
         Call Grid_Head9
         SSTab1.Tab = 5
       End If
   Else
      TxtArea.Locked = False
      TxtArea.BackColor = &H80000005
      Frame6.Enabled = False
      Call Grid_Head9
   End If
Else
   TxtArea.Locked = False
   TxtArea.BackColor = &H80000005
   Frame6.Enabled = False
   Call Grid_Head9
   
   
End If

Call TxtNoofStack_LostFocus

End Sub


Private Sub CmbLicenseNo_GotFocus()
If Trim(CmbLocation.Text) = "" Then
   MsgBox "Select Location!!!"
   CmbLocation.SetFocus
   Exit Sub
End If

tmp = CmbLicenseNo.Text
If LCase(CmbType.Text) = "renewal" Then
   Call TableMst_License(" Where LicenseID in (Select Distinct LicenseID from Txn_LicenseGodownDetail Where GodownID = " & mGodownID & ") ") '(" Where LocationID = 0 or LocationID = " & mLocationID)
ElseIf LCase(CmbType.Text) = "new" Then
   Call TableMst_License(" Where LocationID = 0 or LocationID = " & mLocationID)
End If
CmbLicenseNo.Clear
If RsMst_License.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_License.RecordCount
   CmbLicenseNo.AddItem RsMst_License!LicenseNo
   CmbLicenseNo.Text = RsMst_License!LicenseNo
   RsMst_License.MoveNext
Next


End Sub

Private Sub CmbLicenseNo_LostFocus()
If CmbLicenseNo.Text = "" Then
   Exit Sub
End If
RsMst_License.MoveFirst
RsMst_License.Find "LicenseNo = '" & CmbLicenseNo.Text & "'"
If RsMst_License.EOF Then
   MsgBox "Invalid selection!!!"
   CmbLicenseNo.SetFocus
   Exit Sub
End If
mLicenseID = RsMst_License!LicenseID
End Sub

Private Sub CmbLocation_GotFocus()
If Trim(CmbType.Text) = "" Then
   MsgBox "Select Type!!!"
   CmbType.SetFocus
   Exit Sub
End If
tmp = CmbLocation.Text
If Gen_WcLocation = "" Then
   Call TableMst_Location(" Where LocationID in (Select LocationID from Mst_CMP Where CloseDate is null)")
Else
   Call TableMst_Location(" Where " & Gen_WcLocation)
End If
CmbLocation.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Record found!!!"
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
   CmbLocation.AddItem RsMst_Location!Locationname
   RsMst_Location.MoveNext
Next
CmbLocation.Text = tmp
End Sub

Private Sub CmbLocation_LostFocus()
If CmbLocation.Text = "" Then
   Call Grid_Head2
   CmbCMP.Text = ""
   CmbGodown.Text = ""
   Exit Sub
End If

RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocation.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid  selection!!!"
   CmbLocation.SetFocus
   Exit Sub
End If

mLocationID = RsMst_Location!LocationID
mStateID = RsMst_Location!StateID

If tmp <> CmbLocation.Text Then
   CmbCMP.Text = ""
   CmbGodown.Text = ""
   Call FillLiveGodowns
   Call FillCloseGodowns
   Call FillSecurityandStaffDetail
End If
End Sub

Private Sub FillLiveGodowns()
Call Grid_Head2
tmp = " Select MCMP.CMPName, MG.GodownNo, MUT.UnitType, MG.GodownCapacity, MG.Address, MG.GodownID, MG.Area,  "
tmp = tmp & " (Select sum(MGSL.BalanceWeight) from Mst_GSL MGSL Where MGSL.GodownID = MG.GodownID And MGSL.Flag not in ('E', 'Z', 'B')) BalanceWeight, "
tmp = tmp & " (select sum(TGLA.RentAmount) from Txn_GodownLessorAgreeDetail TGLA Inner Join Txn_LessorAgreement TLA on TLA.LAID = TGLA.LAID Where TLA.Flag in ('A', 'P') And TGLA.GodownID = MG.GodownID And '" & Format(Date, "yyyy-mm-dd") & "' between TLA.StartDate And TLA.ExpiryDate) Rent"
tmp = tmp & " From Mst_Godown MG"
tmp = tmp & " Inner Join Mst_CMP MCMP on MCMP.CMPID =MG.CMPID"
tmp = tmp & " Inner Join Mst_UnitType MUT on MUT.UnitTypeID = MG.UnitTypeID"
tmp = tmp & " Where MG.CloseDate is null And MCMP.LocationID = " & mLocationID
tmp = tmp & " order by MCMP.CMPName, MG.GodownNo"
Call TmpTable

mGodownIDs = "(0 ,"
If TmpRs.RecordCount > 0 Then
   With MSHFlexGrid2
      .Rows = .Rows + TmpRs.RecordCount
      .TextMatrix(0, 1) = "Total No of Godown : " & TmpRs.RecordCount
      For I = 2 To TmpRs.RecordCount + 1
         .TextMatrix(I, 0) = IIf(IsNull(TmpRs!CMPName), "", TmpRs!CMPName)
         .TextMatrix(I, 1) = IIf(IsNull(TmpRs!GodownNo), "", TmpRs!GodownNo)
         .TextMatrix(I, 2) = IIf(IsNull(TmpRs!UnitType), "", TmpRs!UnitType)
         .TextMatrix(I, 3) = Format(IIf(IsNull(TmpRs!GodownCapacity), "0", TmpRs!GodownCapacity), "#0")
         .TextMatrix(I, 4) = Format(IIf(IsNull(TmpRs!BalanceWeight), 0, TmpRs!BalanceWeight), "#0")
         .TextMatrix(I, 5) = Format(IIf(IsNull(TmpRs!Area), "", TmpRs!Area), "#0")
         .TextMatrix(I, 6) = Format(IIf(IsNull(TmpRs!Rent), "0", TmpRs!Rent), "#0")
         .TextMatrix(I, 7) = IIf(IsNull(TmpRs!Address), "", TmpRs!Address)
         .TextMatrix(I, 8) = IIf(IsNull(TmpRs!GodownID), "", TmpRs!GodownID)
         
         mGodownIDs = mGodownIDs & IIf(IsNull(TmpRs!GodownID), 0, TmpRs!GodownID) & " ,"
         
         .TextMatrix(0, 3) = Format(Val(.TextMatrix(0, 3)) + Val(.TextMatrix(I, 3)), "#0")
         .TextMatrix(0, 4) = Format(Val(.TextMatrix(0, 4)) + Val(.TextMatrix(I, 4)), "#0")
         If Val(.TextMatrix(I, 4)) > 0 And Val(.TextMatrix(I, 6)) > 0 Then ' for zero stock
            .TextMatrix(0, 5) = Format(Val(.TextMatrix(0, 5)) + Val(.TextMatrix(I, 5)), "#0")
         Else
            .row = I
            If Val(.TextMatrix(I, 4)) = 0 And Val(.TextMatrix(I, 6)) = 0 Then
               For j = 0 To .Cols - 1
                  .Col = j
                  .CellBackColor = vbRed
               Next
            ElseIf Val(.TextMatrix(I, 4)) = 0 Then
               .Col = 4
               .CellBackColor = vbRed
            ElseIf Val(.TextMatrix(I, 6)) = 0 Then
               .Col = 6
               .CellBackColor = vbRed
            End If
            .TextMatrix(0, 7) = Format(Val(.TextMatrix(0, 7)) + Val(.TextMatrix(I, 5)), "#0")
         End If
         .TextMatrix(0, 6) = Format(Val(.TextMatrix(0, 6)) + Val(.TextMatrix(I, 6)), "#0")
         TmpRs.MoveNext
      Next
      If Val(.TextMatrix(0, 5)) > 0 Then
         .TextMatrix(0, 2) = "Average Rent (per Sq.Ft) : " & Format(Val(.TextMatrix(0, 6)) / Val(.TextMatrix(0, 5)), "#0.00")
      End If
      If Val(.TextMatrix(0, 7)) > 0 Then
         .TextMatrix(0, 7) = "Area with zero stock and rent : " & .TextMatrix(0, 7)
      End If
   End With
End If
mGodownIDs = mID(mGodownIDs, 1, Len(mGodownIDs) - 1) & " )"
End Sub

Private Sub FillCloseGodowns()
Call Grid_Head5
tmp = " Select MCMP.CMPName, MG.GodownNo, MUT.UnitType, MG.GodownCapacity, MG.Address, MG.GodownID, MG.Area,  "
tmp = tmp & " (Select sum(MGSL.BalanceWeight) from Mst_GSL MGSL Where MGSL.GodownID = MG.GodownID And MGSL.Flag not in ('E', 'Z', 'B')) BalanceWeight, "
tmp = tmp & " (select sum(TGLA.RentAmount) from Txn_GodownLessorAgreeDetail TGLA Inner Join Txn_LessorAgreement TLA on TLA.LAID = TGLA.LAID Where TLA.Flag in ('A', 'P') And TGLA.GodownID = MG.GodownID And '" & Format(Date, "yyyy-mm-dd") & "' between TLA.StartDate And TLA.ExpiryDate) Rent"
tmp = tmp & " From Mst_Godown MG"
tmp = tmp & " Inner Join Mst_CMP MCMP on MCMP.CMPID =MG.CMPID"
tmp = tmp & " Inner Join Mst_UnitType MUT on MUT.UnitTypeID = MG.UnitTypeID"
tmp = tmp & " Where MG.CloseDate is not null And MCMP.LocationID = " & mLocationID
tmp = tmp & " order by MCMP.CMPName, MG.GodownNo"
Call TmpTable

If TmpRs.RecordCount > 0 Then
   With MSHFlexGrid5
      .Rows = .Rows + TmpRs.RecordCount
      .TextMatrix(0, 1) = "Total No of Godown : " & TmpRs.RecordCount
      For I = 2 To TmpRs.RecordCount + 1
         .TextMatrix(I, 0) = IIf(IsNull(TmpRs!CMPName), "", TmpRs!CMPName)
         .TextMatrix(I, 1) = IIf(IsNull(TmpRs!GodownNo), "", TmpRs!GodownNo)
         .TextMatrix(I, 2) = IIf(IsNull(TmpRs!UnitType), "", TmpRs!UnitType)
         .TextMatrix(I, 3) = Format(IIf(IsNull(TmpRs!GodownCapacity), "0", TmpRs!GodownCapacity), "#0")
         .TextMatrix(I, 4) = Format(IIf(IsNull(TmpRs!BalanceWeight), 0, TmpRs!BalanceWeight), "#0")
         .TextMatrix(I, 5) = Format(IIf(IsNull(TmpRs!Area), "", TmpRs!Area), "#0")
         .TextMatrix(I, 6) = Format(IIf(IsNull(TmpRs!Rent), "0", TmpRs!Rent), "#0")
         .TextMatrix(I, 7) = IIf(IsNull(TmpRs!Address), "", TmpRs!Address)
         .TextMatrix(I, 8) = IIf(IsNull(TmpRs!GodownID), "", TmpRs!GodownID)
         .TextMatrix(0, 3) = Format(Val(.TextMatrix(0, 3)) + Val(.TextMatrix(I, 3)), "#0")
         .TextMatrix(0, 4) = Format(Val(.TextMatrix(0, 4)) + Val(.TextMatrix(I, 4)), "#0")
         .TextMatrix(0, 5) = Format(Val(.TextMatrix(0, 5)) + Val(.TextMatrix(I, 5)), "#0")
         .TextMatrix(0, 6) = Format(Val(.TextMatrix(0, 6)) + Val(.TextMatrix(I, 6)), "#0")
         TmpRs.MoveNext
      Next
      If Val(.TextMatrix(0, 5)) > 0 Then
         .TextMatrix(0, 2) = "Average Rent (per Sq.Ft) : " & Format(Val(.TextMatrix(0, 6)) / Val(.TextMatrix(0, 5)), "#0.00")
      End If
   End With
End If
End Sub


Private Sub FillSecurityandStaffDetail()
tmp = "select count(MSLD.TxnDetailID) SecStaff, sum(MSLD.RateperMonth) SecSalStaff from Mst_SecLocMappingDetail MSLD"
tmp = tmp & " Where getdate() between MSLD.StartDate and MSLD.EndDate"
tmp = tmp & " And MSLD.TxnDetailID in (Select MSLGD.TxnDetailID from Mst_SecLocGodMappingDetail MSLGD Where MSLGD.GodownID in " & mGodownIDs & " )"
Call TmpTable
If TmpRs.RecordCount > 0 Then
   MSHFlexGrid6.TextMatrix(3, 1) = Format(IIf(IsNull(TmpRs!SecStaff), "", TmpRs!SecStaff), "#0")
   MSHFlexGrid6.TextMatrix(3, 3) = Format(IIf(IsNull(TmpRs!SecStaff), "", TmpRs!SecStaff), "#0")
   MSHFlexGrid6.TextMatrix(3, 2) = Format(IIf(IsNull(TmpRs!SecSalStaff), "", TmpRs!SecSalStaff), "#0")
   MSHFlexGrid6.TextMatrix(3, 4) = Format(IIf(IsNull(TmpRs!SecSalStaff), "", TmpRs!SecSalStaff), "#0")
'============================for security and staff salary expense detail============================
MSHFlexGrid4.TextMatrix(0, 3) = Format(Val(MSHFlexGrid4.TextMatrix(0, 3)) - Val(MSHFlexGrid4.TextMatrix(MSHFlexGrid4.Rows - 2, 3)), "#0")
MSHFlexGrid4.TextMatrix(0, 3) = Format(Val(MSHFlexGrid4.TextMatrix(0, 3)) - Val(MSHFlexGrid4.TextMatrix(MSHFlexGrid4.Rows - 3, 3)), "#0")
MSHFlexGrid4.Rows = MSHFlexGrid4.Rows - 2
Call SecandUserSalaryExpense
'============================for security and staff salary expense detail============================
End If

tmp = " select ME.EmployeeNo from Mst_EmpCMPMapping MECM"
tmp = tmp & " Inner Join Mst_Employee ME on ME.EmployeeID = MECM.EmployeeID and ME.Flag = 'Y'"
tmp = tmp & " Inner Join Mst_Godown MG on MG.CMPID = MECM.CMPID"
tmp = tmp & " Where MG.GodownID in " & mGodownIDs
tmp = tmp & " union"
tmp = tmp & " select ME1.EmployeeNo"
tmp = tmp & " From Mst_EMPCMPMapping MECM"
tmp = tmp & " Inner Join Mst_Godown MG on MG.CMPID = MECM.CMPID"
tmp = tmp & " Inner Join Mst_Employee ME on ME.EmployeeID = MECM.EmployeeID and ME.Flag = 'Y'"
tmp = tmp & " Inner Join Mst_Employee ME1 On ME.EmployeeID  = ME1.DirectReportingTo and ME1.Flag = 'Y'"
tmp = tmp & " where ME1.BaseLocationID = " & mLocationID
tmp = tmp & " And MG.GodownID in " & mGodownIDs
Call TmpTable

mEmployeeNos = ""
If TmpRs.RecordCount > 0 Then
   MSHFlexGrid6.TextMatrix(2, 1) = TmpRs.RecordCount
   MSHFlexGrid6.TextMatrix(2, 3) = TmpRs.RecordCount
   mEmployeeNos = "('0' ,"
   For I = 1 To TmpRs.RecordCount
      mEmployeeNos = mEmployeeNos & "'" & IIf(IsNull(TmpRs!EmployeeNo), 0, TmpRs!EmployeeNo) & "' ,"
      TmpRs.MoveNext
   Next
   mEmployeeNos = mID(mEmployeeNos, 1, Len(mEmployeeNos) - 1) & " )"
End If

If mEmployeeNos <> "" Then
   mStartDate = DateSerial(Year(Date), Month(Date) - 1, 1)
   mEndDate = DateSerial(Year(Date), Month(Date), 0)
   
   tmp = " Select sum(isnull(BasicMonthly,0)) + sum(isnull(HRAMonthly,0)) + sum(isnull(ConveyanceMonthly,0)) + "
   tmp = tmp & " sum(isnull(OtherAllMonthly,0)) + sum(isnull(MEDICAL,0)) + sum(isnull(SplAllMonthly,0)) + sum(isnull(LTA,0)) +"
   tmp = tmp & " Sum (IsNull(LaptopAllowMonthly, 0)) + Sum(IsNull(LeaveEncashment, 0)) + Sum(IsNull(GainShare, 0)) + Sum(IsNull(Incentive, 0))"
   tmp = tmp & " + sum(isnull(Perquisites,0)) + sum(isnull(ARREARSBASIC,0)) + sum(isnull(ARREARSHRA,0)) +"
   tmp = tmp & " sum(isnull(ARREARSCONVEYANCE,0)) + sum(isnull(ARREARSMEDICAL,0)) + sum(isnull(PERFORMANCEBONUS,0)) +"
   tmp = tmp & " Sum (IsNull(PROFESSIONALFEES, 0)) + Sum(IsNull(STIPEND, 0)) + Sum(IsNull(SERVICECHARGES, 0))"
   tmp = tmp & " + sum(isnull(ACCRUEDVARIABLEPAY,0)) + sum(isnull(ACCRUEDRETIRALS,0))+sum(isnull(employersesic,0))"
   tmp = tmp & " + sum(isnull(carmaintandpetrol,0))+ sum(isnull(booksandperiodical,0))+ sum(isnull(mealallowance,0))"
   tmp = tmp & " + sum(isnull(LWFEMPLOYER,0)) + sum(isnull(EMPLOYERSPF,0)) + sum(isnull(STaxandEducationcess,0)) StaffSalary "
   tmp = tmp & " From Txn_SalaryDistribution "
   tmp = tmp & " Where MonthYear in (Select max(monthyear) from Txn_SalaryDistribution)"
   tmp = tmp & " And BusinessSegmentID = 1 And EmployeeNo in " & mEmployeeNos
   tmp = tmp & " And LocationID = " & mLocationID
   Call TmpTableHRPro
   If TmpRsHRPro.RecordCount > 0 Then
      MSHFlexGrid6.TextMatrix(2, 2) = Format(IIf(IsNull(TmpRsHRPro!staffsalary), "", TmpRsHRPro!staffsalary), "#0")
      MSHFlexGrid6.TextMatrix(2, 4) = Format(IIf(IsNull(TmpRsHRPro!staffsalary), "", TmpRsHRPro!staffsalary), "#0")
      
      '============================for security and staff salary expense detail============================
      MSHFlexGrid4.TextMatrix(0, 3) = Format(Val(MSHFlexGrid4.TextMatrix(0, 3)) - Val(MSHFlexGrid4.TextMatrix(MSHFlexGrid4.Rows - 2, 3)), "#0")
      MSHFlexGrid4.TextMatrix(0, 3) = Format(Val(MSHFlexGrid4.TextMatrix(0, 3)) - Val(MSHFlexGrid4.TextMatrix(MSHFlexGrid4.Rows - 3, 3)), "#0")
      MSHFlexGrid4.Rows = MSHFlexGrid4.Rows - 2
      Call SecandUserSalaryExpense
      '============================for security and staff salary expense detail============================
   End If
End If
End Sub

Private Sub CmbSCmp_GotFocus()
If Trim(CmbSLocation.Text) = "" Then
   MsgBox "Select Location!!!"
   CmbSLocation.SetFocus
   Exit Sub
End If
tmp = CmbSCmp.Text
If Gen_WcCMP = "" Then
   Call TableMst_CMP(" where CloseDate is null and LocationID = " & mLocationID)
Else
   Call TableMst_CMP(" where " & Gen_WcCMP & " And CloseDate is null and LocationID = " & mLocationID)
End If
CmbSCmp.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "No Record found!!!"
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
   CmbSGodown.Text = ""
   Exit Sub
End If

If tmp <> CmbSCmp.Text Then
   CmbSGodown.Text = ""
End If

RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPName = '" & CmbSCmp.Text & "'"
If RsMst_CMP.EOF Then
   MsgBox "Invalid selection!!!"
   CmbSCmp.SetFocus
   Exit Sub
End If

mCMPID = RsMst_CMP!CMPID
End Sub

Private Sub CmbSGodown_GotFocus()
If Trim(CmbSType.Text) = "" Then
   MsgBox "Select Type!!!"
   CmbSType.SetFocus
   Exit Sub
End If
If Trim(CmbSCmp.Text) = "" Then
   MsgBox "Select CMP!!!"
   CmbSCmp.SetFocus
   Exit Sub
End If
tmp = CmbSGodown.Text
Call TableMst_Godown(" where CloseDate is null and CMPID = " & mCMPID)
CmbSGodown.Clear
If RsMst_Godown.RecordCount = 0 Then
   MsgBox "No Record found!!!"
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

If tmp <> CmbSGodown.Text Then
   ''write if clear date
End If

RsMst_Godown.MoveFirst
RsMst_Godown.Find "GodownNo= '" & CmbSGodown.Text & "'"
If RsMst_Godown.EOF Then
   MsgBox "Invalid selection!!!"
   CmbSGodown.SetFocus
   Exit Sub
End If

mSGodownID = RsMst_Godown!GodownID
End Sub

Private Sub CmbSGodownType_GotFocus()
tmp = CmbSGodownType.Text
Call TableMst_UnitType
CmbSGodownType.Clear
If RsMst_UnitType.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_UnitType.RecordCount
    CmbSGodownType.AddItem RsMst_UnitType!UnitType
    RsMst_UnitType.MoveNext
Next
CmbSGodownType.Text = tmp
End Sub

Private Sub CmbSGodownType_LostFocus()
If CmbSGodownType.Text = "" Then
   Exit Sub
End If

If tmp <> CmbSGodownType.Text Then
   ''write if clear date
End If

RsMst_UnitType.MoveFirst
RsMst_UnitType.Find "UnitType = '" & CmbSGodownType.Text & "'"
If RsMst_UnitType.EOF Then
   MsgBox "Invalid selection!!!"
   CmbSGodownType.SetFocus
   Exit Sub
End If

mGodownTypeID = RsMst_UnitType!UnitTypeID
End Sub

Private Sub CmbSLocation_GotFocus()
'If Trim(CmbSType.Text) = "" Then
'   MsgBox "Select Type!!!"
'   CmbSType.SetFocus
'   Exit Sub
'End If
tmp = CmbSLocation.Text
If Gen_WcLocation = "" Then
   Call TableMst_Location
Else
   Call TableMst_Location(" Where " & Gen_WcLocation)
End If
CmbSLocation.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Record found!!!"
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
   CmbSCmp.Text = ""
   CmbSGodown.Text = ""
   Exit Sub
End If

If tmp <> CmbSLocation.Text Then
   CmbSCmp.Text = ""
   CmbSGodown.Text = ""
End If

RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbSLocation.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid  selection!!!"
   CmbLocation.SetFocus
   Exit Sub
End If

mLocationID = RsMst_Location!LocationID
End Sub

Private Sub CmbSStatus_LostFocus()
If CmbSStatus.Text = "" Then Exit Sub

If LCase(CmbSStatus.Text) <> "pending" And LCase(CmbSStatus.Text) <> "approved" And LCase(CmbSStatus.Text) <> "rejected" And LCase(CmbSStatus.Text) <> "godown creation" Then
   MsgBox "Invalid selection!!! "
   CmbSStatus.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmbStatus_LostFocus()

If CmbStatus.Text = "" Then Exit Sub

If LCase(CmbStatus.Text) <> "pending" And LCase(CmbStatus.Text) <> "approved" And LCase(CmbStatus.Text) <> "rejected" And LCase(CmbStatus.Text) <> "godown creation" Then
   MsgBox "Invalid selection!!! "
   CmbStatus.SetFocus
   Exit Sub
End If

If LCase(CmbStatus.Text) = "approved" Or LCase(CmbStatus.Text) = "godown creation" Then
   If LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "superadmin" And LCase(Gen_UserRole) <> "headcmg" Then
      MsgBox "Rights for Approval/Godown Creation are with CMG Department!!! "
      CmbStatus.SetFocus
      Exit Sub
   End If
End If

If TxtInspectionID.Text <> "" Then
   If RsMst_GodownInspection!Status = "P" And LCase(CmbStatus.Text) = "godown creation" Then
      MsgBox "You can not select 'Godown Creation', when current status is 'Pending'!!!"
      CmbStatus.SetFocus
      Exit Sub
   End If
End If


End Sub

Private Sub CmbSType_GotFocus()
tmp = CmbSType.Text
End Sub

Private Sub CmbSType_LostFocus()
If CmbSType.Text = "" Then Exit Sub

If LCase(CmbSType.Text) <> "new" And LCase(CmbSType.Text) <> "renewal" Then
   MsgBox "Invalid selection!!! "
   CmbSType.SetFocus
   Exit Sub
End If
If CmbSType.Text <> tmp Then
   If LCase(CmbType.Text) = "new" Then
      CmbSGodown.Text = ""
      TxtSGodown.Text = ""
      CmbSGodown.Visible = False
      TxtSGodown.Visible = True
   ElseIf LCase(CmbSType.Text) = "renewal" Then
      CmbSGodown.Text = ""
      TxtSGodown.Text = ""
      CmbSGodown.Visible = True
      TxtSGodown.Visible = False
   End If
End If
End Sub

Private Sub CmbType_GotFocus()
tmp = CmbType.Text
End Sub

Private Sub CmbType_LostFocus()
If CmbType.Text = "" Then Exit Sub

If LCase(CmbType.Text) <> "new" And LCase(CmbType.Text) <> "renewal" Then
   MsgBox "Invalid selection!!! "
   CmbType.SetFocus
   Exit Sub
End If
If CmbType.Text <> tmp Then
   If LCase(CmbType.Text) = "new" Then
      CmbGodown.Text = ""
      TxtGodown.Text = ""
      CmbGodown.Visible = False
      TxtGodown.Visible = True
      'TxtCapacity.Enabled = True
      TxtArea.Enabled = True
      TxtAddress.Enabled = True
      CmbGodownType.Enabled = True
      CmbGodownCategory.Enabled = True
      CmbLicenseNo.Enabled = True
      TxtNoofStack.Locked = False
      TxtWeightPerStack.Locked = False
   ElseIf LCase(CmbType.Text) = "renewal" Then
      mAgreementExpiryDate = ""
      CmbGodown.Text = ""
      TxtGodown.Text = ""
      CmbGodown.Visible = True
      TxtGodown.Visible = False
      'TxtCapacity.Enabled = False
      TxtArea.Enabled = False
      TxtAddress.Enabled = False
      CmbGodownType.Enabled = False
      CmbGodownCategory.Enabled = False
      CmbLicenseNo.Enabled = False
      TxtNoofStack.Locked = True
      TxtWeightPerStack.Locked = True
   End If
End If

End Sub

Private Sub CmdAreaAddToList_Click()


If Val(TxtLength) = 0 And Val(TxtWidth) = 0 And Val(TxtHeight) = 0 Then
   MsgBox "All values can't be zero!!!"
   TxtLength.SetFocus
End If

If CmdAreaAddToList.Caption = "Update List" Then
   r = MSHFlexGrid9.RowSel
   MSHFlexGrid9.TextMatrix(0, 4) = Val(MSHFlexGrid9.TextMatrix(0, 4)) - Val(MSHFlexGrid9.TextMatrix(r, 4))
   TxtArea = Val(TxtArea) - Val(MSHFlexGrid9.TextMatrix(r, 4))
Else
   r = MSHFlexGrid9.Rows - 1
   MSHFlexGrid9.Rows = MSHFlexGrid9.Rows + 1
   MSHFlexGrid9.TextMatrix(r, 0) = ""
End If

MSHFlexGrid9.TextMatrix(r, 1) = Val(TxtLength)
MSHFlexGrid9.TextMatrix(r, 2) = Val(TxtWidth)
MSHFlexGrid9.TextMatrix(r, 3) = Val(TxtHeight)
MSHFlexGrid9.TextMatrix(r, 4) = Val(TxtSubArea)

MSHFlexGrid9.TextMatrix(0, 4) = Val(MSHFlexGrid9.TextMatrix(0, 4)) + Val(MSHFlexGrid9.TextMatrix(r, 4))
If r = 2 Then
   If LCase(CmdAreaAddToList.Caption) <> "update list" Then
      TxtArea = Val(MSHFlexGrid9.TextMatrix(r, 4))
   Else
      TxtArea = Val(MSHFlexGrid9.TextMatrix(0, 4)) 'Val(TxtArea) + Val(MSHFlexGrid9.TextMatrix(r, 4))
   End If
Else
   TxtArea = Val(MSHFlexGrid9.TextMatrix(0, 4)) 'Val(TxtArea) + Val(MSHFlexGrid9.TextMatrix(r, 4))
End If

Call TxtArea_LostFocus

Call ClearAreaDetail



End Sub

Private Sub CmdAreaDeleteFromList_Click()

If MSHFlexGrid9.Rows < 3 Then Exit Sub
Dim ans As Variant

If TxtLength.Text = "" Then
   MsgBox "Please Select atleast one row from grid!!!"
   Exit Sub
End If

With MSHFlexGrid9
   I = .RowSel
      
   If I > 1 Then
      .TextMatrix(0, 4) = Val(.TextMatrix(0, 4)) - Val(.TextMatrix(I, 4))
      TxtArea = Val(.TextMatrix(0, 4))
      For RO = I To .Rows - 2
         For C = 0 To .Cols - 1
            .TextMatrix(RO, C) = .TextMatrix(RO + 1, C)
         Next
      Next
      .Rows = .Rows - 1
      Call TxtArea_LostFocus
   End If
   
End With

Call ClearAreaDetail

End Sub

Private Sub CmdCAddToList_Click()
If Val(TxtNCDEXStock.Text) + Val(TxtNNNRStock.Text) + Val(TxtReservationStock.Text) <> 100 Then
   MsgBox "Sum of Stock will be 100%!!!"
   SSTab1.Tab = 0
   TxtNCDEXStock.SetFocus
   Exit Sub
End If
If Trim(CmbCategory.Text) = "" Then
   MsgBox "Blank Category not allowed!!!"
   CmbCategory.SetFocus
   Exit Sub
End If
If Trim(CmbCommodity.Text) = "" Then
   MsgBox "Blank Commodity not allowed!!!"
   CmbCommodity.SetFocus
   Exit Sub
End If
If Trim(TxtExpQtyP.Text) = "" Then
   MsgBox "Blank Expected Qty (in %)not allowed!!!"
   TxtExpQtyP.SetFocus
   Exit Sub
End If
If Trim(TxtStorageRate.Text) = "" Then
   MsgBox "Blank Storage Rate not allowed!!!"
   TxtStorageRate.SetFocus
   Exit Sub
End If
If Trim(TxtPrice.Text) = "" Then
   MsgBox "Blank Price not allowed!!!"
   TxtPrice.SetFocus
   Exit Sub
End If
If Val(TxtFumigableQty.Text) > Val(TxtExpQtyMT.Text) Then
   MsgBox "Fumigation Qty should not greater than Expected Qty!!!"
   TxtFumigableQty.SetFocus
   Exit Sub
End If

With MSHFlexGrid3
   For I = 2 To .Rows - 2
      If CmdCAddToList.Caption = "Update List" Then
         If I <> .RowSel Then
            If .TextMatrix(I, 0) = CmbCategory.Text And .TextMatrix(I, 1) = CmbCommodity.Text Then
               MsgBox "Duplicate Category and Commodity not Allowed!!!"
               Exit Sub
            End If
         End If
      Else
         If .TextMatrix(I, 0) = CmbCategory.Text And .TextMatrix(I, 1) = CmbCommodity.Text Then
            MsgBox "Duplicate Category and Commodity not Allowed!!!"
            Exit Sub
         End If
      End If
   Next

   If CmdCAddToList.Caption = "Update List" Then
      
      I = .RowSel
      If LCase(CmbCategory.Text) = "ncdex" Then
         If mNCDEXPer - Val(.TextMatrix(I, 2)) + Val(TxtExpQtyP.Text) > 100 Then
            MsgBox "Expected Qty (in %) under 'NCDEX' greater than 100% not allowed!!!"
            SSTab1.Tab = 1
            TxtExpQtyP.SetFocus
            Exit Sub
         End If
         mNCDEXPer = mNCDEXPer - Val(.TextMatrix(I, 2))
      ElseIf LCase(CmbCategory.Text) = "nnnr" Then
         If mNNNRPer - Val(.TextMatrix(I, 2)) + Val(TxtExpQtyP.Text) > 100 Then
            MsgBox "Expected Qty (in %) under 'NNNR' greater than 100% not allowed!!!"
            SSTab1.Tab = 1
            TxtExpQtyP.SetFocus
            Exit Sub
         End If
         mNNNRPer = mNNNRPer - Val(.TextMatrix(I, 2))
      ElseIf LCase(CmbCategory.Text) = "reservation" Then
         If mReservationPer - Val(.TextMatrix(I, 2)) + Val(TxtExpQtyP.Text) > 100 Then
            MsgBox "Expected Qty (in %) under 'Reservation' greater than 100% not allowed!!!"
            SSTab1.Tab = 1
            TxtExpQtyP.SetFocus
            Exit Sub
         End If
         mReservationPer = mReservationPer - Val(.TextMatrix(I, 2))
      End If
      .TextMatrix(0, 3) = Val(.TextMatrix(0, 3)) - Val(.TextMatrix(I, 3))
      .TextMatrix(0, 6) = Val(.TextMatrix(0, 6)) - Val(.TextMatrix(I, 6))
      .TextMatrix(0, 7) = Val(.TextMatrix(0, 7)) - Val(.TextMatrix(I, 7))
      .TextMatrix(0, 11) = Val(.TextMatrix(0, 11)) - Val(.TextMatrix(I, 11))
      .TextMatrix(0, 14) = Val(.TextMatrix(0, 14)) - Val(.TextMatrix(I, 14))
      
   Else
      If LCase(CmbCategory.Text) = "ncdex" Then
         If mNCDEXPer + Val(TxtExpQtyP.Text) > 100 Then
            MsgBox "Expected Qty (in %) under 'NCDEX' greater than 100% not allowed!!!"
            SSTab1.Tab = 1
            TxtExpQtyP.SetFocus
            Exit Sub
         End If
      ElseIf LCase(CmbCategory.Text) = "nnnr" Then
         If mNNNRPer + Val(TxtExpQtyP.Text) > 100 Then
            MsgBox "Expected Qty (in %) under 'NNNR' greater than 100% not allowed!!!"
            SSTab1.Tab = 1
            TxtExpQtyP.SetFocus
            Exit Sub
         End If
      ElseIf LCase(CmbCategory.Text) = "reservation" Then
         If mReservationPer + Val(TxtExpQtyP.Text) > 100 Then
            MsgBox "Expected Qty (in %) under 'Reservation' greater than 100% not allowed!!!"
            SSTab1.Tab = 1
            TxtExpQtyP.SetFocus
            Exit Sub
         End If
      End If
      I = .Rows - 1
      .Rows = .Rows + 1
   End If
   .TextMatrix(I, 0) = CmbCategory.Text
   .TextMatrix(I, 1) = CmbCommodity.Text
   .TextMatrix(I, 2) = TxtExpQtyP.Text
   .TextMatrix(I, 3) = TxtExpQtyMT.Text
   .TextMatrix(I, 4) = TxtStorageRate.Text
   .TextMatrix(I, 5) = TxtPrice.Text
   .TextMatrix(I, 6) = TxtRevenue.Text
   .TextMatrix(I, 7) = TxtAUM.Text
   .TextMatrix(I, 8) = TxtStorageRemarks.Text
   .TextMatrix(I, 10) = mCommodityID
   .TextMatrix(I, 11) = TxtFumigableQty.Text
   .TextMatrix(I, 12) = mFumigationFlag
   .TextMatrix(I, 13) = TxtCMShareP.Text
   .TextMatrix(I, 14) = TxtCMShareR.Text
   
   .TextMatrix(0, 3) = Format(Val(.TextMatrix(0, 3)) + Val(.TextMatrix(I, 3)), "#0")
   .TextMatrix(0, 6) = Format(Val(.TextMatrix(0, 6)) + Val(.TextMatrix(I, 6)), "#0")
   .TextMatrix(0, 7) = Format(Val(.TextMatrix(0, 7)) + Val(.TextMatrix(I, 7)), "#0")
   .TextMatrix(0, 11) = Format(Val(.TextMatrix(0, 11)) + Val(.TextMatrix(I, 11)), "#0")
   .TextMatrix(0, 14) = Format(Val(.TextMatrix(0, 14)) + Val(.TextMatrix(I, 14)), "#0")
   
   If LCase(CmbCategory.Text) = "ncdex" Then
      mNCDEXPer = mNCDEXPer + Val(.TextMatrix(I, 2))
   ElseIf LCase(CmbCategory.Text) = "nnnr" Then
      mNNNRPer = mNNNRPer + Val(.TextMatrix(I, 2))
   ElseIf LCase(CmbCategory.Text) = "reservation" Then
      mReservationPer = mReservationPer + Val(.TextMatrix(I, 2))
   End If
End With
Call ClearCommodityDetail
If Val(MSHFlexGrid4.TextMatrix(0, 3)) > Val(MSHFlexGrid4.TextMatrix(MSHFlexGrid4.Rows - 2, 3)) + Val(MSHFlexGrid4.TextMatrix(MSHFlexGrid4.Rows - 3, 3)) Then
   MsgBox "Expense will be clear!!!"
   Call ClearAllExpenseDetails
End If
End Sub

Private Sub CmdCancel_Click()
Frame5.Visible = False
End Sub

Private Sub CmdCDeleteFromList_Click()
If MSHFlexGrid3.Rows < 3 Then Exit Sub
Dim ans As Variant
If CmbCategory.Text = "" Then
   MsgBox "Please Select atleast one row from grid!!!"
   Exit Sub
End If

With MSHFlexGrid3
   I = .RowSel
      
   If I > 1 Then
      If LCase(.TextMatrix(I, 0)) = "ncdex" Then
         mNCDEXPer = mNCDEXPer - Val(.TextMatrix(I, 2))
      ElseIf LCase(CmbCategory.Text) = "nnnr" Then
         mNNNRPer = mNNNRPer - Val(.TextMatrix(I, 2))
      ElseIf LCase(CmbCategory.Text) = "reservation" Then
         mReservationPer = mReservationPer - Val(.TextMatrix(I, 2))
      End If
      .TextMatrix(0, 3) = Val(.TextMatrix(0, 3)) - Val(.TextMatrix(I, 3))
      .TextMatrix(0, 6) = Val(.TextMatrix(0, 6)) - Val(.TextMatrix(I, 6))
      .TextMatrix(0, 7) = Val(.TextMatrix(0, 7)) - Val(.TextMatrix(I, 7))
      .TextMatrix(0, 11) = Val(.TextMatrix(0, 11)) - Val(.TextMatrix(I, 11))
      .TextMatrix(0, 14) = Val(.TextMatrix(0, 14)) - Val(.TextMatrix(I, 14))
         
      For RO = I To .Rows - 2
         For C = 0 To .Cols - 1
            .TextMatrix(RO, C) = .TextMatrix(RO + 1, C)
         Next
      Next
      .Rows = .Rows - 1
   End If
End With
Call ClearCommodityDetail
If Val(MSHFlexGrid4.TextMatrix(0, 3)) > Val(MSHFlexGrid4.TextMatrix(MSHFlexGrid4.Rows - 2, 3)) + Val(MSHFlexGrid4.TextMatrix(MSHFlexGrid4.Rows - 3, 3)) Then
   MsgBox "Expense will be clear!!!"
   Call ClearAllExpenseDetails
End If
End Sub

Private Sub CmdClearData_Click()

With MSHFlexGrid4
   I = .RowSel
   .TextMatrix(0, 3) = Val(.TextMatrix(0, 3)) - Val(.TextMatrix(I, 3))
   .TextMatrix(I, 2) = 0
   .TextMatrix(I, 3) = 0
End With

Call ClearExpenseDetail

End Sub

Private Sub CmdCloseGodown_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & "Mst_CloseGodownViabilityDetails.xls")
Call Xls_Detail_Rpt2
End Sub

Public Sub Xls_Detail_Rpt2()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Integer
Dim mCol As Integer
MsgBox ("Wait till next message!!!")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & "Mst_CloseGodownViabilityDetails.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid5.Rows - 1
   mRow = I + 1 ' + 6
   For C = 0 To MSHFlexGrid5.Cols - 1
      oWS.Cells(mRow, C + 1) = MSHFlexGrid5.TextMatrix(I, C)
   Next
Next
oWB.Save
MsgBox ("Excel file created!!!")
oXL.Visible = True
End Sub

Private Sub CmdCMP_Click()
FrmMst_CMP.Show
End Sub

Private Sub CmdEUpdateList_Click()
'If Trim(TxtQtySq.Text) = "" Then
'   MsgBox "Blank Qty/Sq.Ft/No of not allowed!!!"
'   TxtQtySq.SetFocus
'   Exit Sub
'End If
If Trim(TxtAmount.Text) = "" Then
   MsgBox "Blank Amount not allowed!!!"
   TxtAmount.SetFocus
   Exit Sub
End If
With MSHFlexGrid4
   I = .RowSel
   .TextMatrix(0, 3) = Val(.TextMatrix(0, 3)) - Val(.TextMatrix(I, 3))
   .TextMatrix(I, 2) = Format(TxtQtySq.Text, "#0")
   .TextMatrix(I, 3) = Format(TxtAmount.Text, "#0")
   .TextMatrix(I, 4) = TxtExpenseRemarks.Text
   .TextMatrix(0, 3) = Format(Val(.TextMatrix(0, 3)) + Val(.TextMatrix(I, 3)), "#0")
   .TextMatrix(I, 7) = TxtAmountPer.Text
'   TxtDirectExpense.Text = Format(Val(.TextMatrix(0, 3)), "#0")
'   'TxtNetProfit = Format(Val(TxtTotalRevenue.Text) - Val(TxtDirectExpense.Text) - Val(TxtSecStaffSalaryUser.Text) - Val(TxtStaffSalaryUser.Text), "#0")
'   TxtGrossProfit.Text = Format(Val(TxtTotalRevenue.Text) - Val(TxtDirectExpense.Text), "#0")
'   If LCase(CmbGodownType.Text) = "renewal" Then
'      TxtIndirectExpense.Text = (Val(MSHFlexGrid6.TextMatrix(2, 4)) + Val(MSHFlexGrid6.TextMatrix(3, 4))) * Val(TxtCapacity.Text) / Val(MSHFlexGrid2.TextMatrix(0, 3))
'   ElseIf LCase(CmbGodownType.Text) = "new" Then
'      TxtIndirectExpense.Text = (Val(MSHFlexGrid6.TextMatrix(2, 4)) + Val(MSHFlexGrid6.TextMatrix(3, 4))) * Val(TxtCapacity.Text) / (Val(MSHFlexGrid2.TextMatrix(0, 3)) + Val(TxtCapacity.Text))
'   End If
'   TxtNetProfit = Format(Val(TxtTotalRevenue.Text) - Val(TxtDirectExpense.Text) - Val(TxtIndirectExpense.Text), "#0")

End With
If LCase(TxtExpenseType.Text) = "leased rent" Then
   lblaverageleasedrent.Caption = "Average Leased Rent : 0"
End If
   
Call ClearExpenseDetail
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & "Mst_GodownViability.xls")
Call Xls_Detail_Rpt
End Sub

Public Sub Xls_Detail_Rpt()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Integer
Dim mCol As Integer
MsgBox ("Wait till next message!!!")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & "Mst_GodownViability.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid1.Rows - 1
   mRow = I + 1 ' + 6
   For C = 0 To MSHFlexGrid1.Cols - 1
      If C = 4 Then
         oWS.Cells(mRow, C + 1) = "'" & MSHFlexGrid1.TextMatrix(I, C)
      Else
         oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
      End If
   Next
Next
oWB.Save
MsgBox ("Excel file created!!!")
oXL.Visible = True
End Sub

Private Sub CmdSave_Click()
With MSHFlexGrid8
     For I = 1 To MSHFlexGrid8.Rows - 1
         If MSHFlexGrid8.TextMatrix(I, 1) <> "" Then
            tmp = " Select GodownID from Txn_LicenseGodownDetail where GodownID=" & mmGodownID & ""
            Call Tmp4Table
            If TmpRs4.RecordCount = 0 Then
               Call TableTxn_LicenseGodownDetail("Where LicenseID=" & Val(MSHFlexGrid8.TextMatrix(I, 3)) & "")
               
               RsTxn_LicenseGodownDetail.AddNew
               RsTxn_LicenseGodownDetail!ID = GetMaxID("ID", "Txn_LicenseGodownDetail")
               RsTxn_LicenseGodownDetail!LicenseID = Val(MSHFlexGrid8.TextMatrix(I, 3))
               RsTxn_LicenseGodownDetail!GodownID = mmGodownID
               RsTxn_LicenseGodownDetail!G_UID = GetGUID
               If IsNull(RsTxn_LicenseGodownDetail!CreatedBy) = True Or RsTxn_LicenseGodownDetail!CreatedBy = "" Then
                  RsTxn_LicenseGodownDetail!CreatedBy = Gen_UserLoginID
                  RsTxn_LicenseGodownDetail!CreatedDate = Now
               Else
                  RsTxn_LicenseGodownDetail!UpdatedBy = Gen_UserLoginID
                  RsTxn_LicenseGodownDetail!UpdatedDate = Now
               End If
               RsTxn_LicenseGodownDetail.Update
            End If
         End If
     Next
Frame5.Visible = False
End With
End Sub

Private Sub CmdSearch_Click()
Dim wc As Variant

wc = ""

Call Grid_Head

wc = " Where MGI.StartDate between '" & SFrom.Value & "' And '" & STo.Value & "'"

If CmbSType.Text <> "" Then
   wc = wc & " And MGI.Type = '" & Left$(CmbSType.Text, 1) & "'"
End If

If CmbSLocation.Text <> "" Then
   If wc = "" Then
      wc = " Where MGI.LocationID = " & mLocationID
   Else
      wc = wc & " And MGI.LocationID = " & mLocationID
   End If
ElseIf Gen_WcLocation <> "" Then
   If wc = "" Then
      wc = " Where MGI." & Gen_WcLocation
   Else
      wc = wc & " And MGI." & Gen_WcLocation
   End If
End If
If CmbSCmp.Text <> "" Then
   If wc = "" Then
      wc = " Where MGI.CMPID = " & mCMPID
   Else
      wc = wc & " And MGI.CMPID = " & mCMPID
   End If
ElseIf Gen_WcCMP <> "" Then
   If wc = "" Then
      wc = " Where MGI." & Gen_WcCMP
   Else
      wc = wc & " And MGI." & Gen_WcCMP
   End If
End If
If CmbSGodown.Text <> "" Then
   If wc = "" Then
      wc = " Where MGI.GodownID = " & mGodownID
   Else
      wc = wc & " And MGI.GodownID = " & mGodownID
   End If
End If

If TxtSGodown.Text <> "" Then
   If wc = "" Then
      wc = " Where MGI.GodownNo like '%" & TxtSGodown.Text & "%'"
   Else
      wc = wc & " And MGI.GodownNo like '%" & TxtSGodown.Text & "%'"
   End If
End If

If CmbSGodownType.Text <> "" Then
   If wc = "" Then
      wc = " Where MGI.UnitTypeID = " & mGodownTypeID
   Else
      wc = wc & " And MGI.UnitTypeID = " & mGodownTypeID
   End If
End If
'If CmbSGodownCategory.Text <> "" Then
'   If wc = "" Then
'      wc = " Where MGI.GodownCategoryID = " & mGodownCategoryID
'   Else
'      wc = wc & " And MGI.GodownCategoryID = " & mGodownCategoryID
'   End If
'End If

If CmbSStatus.Text <> "" Then
   If wc = "" Then
      wc = " Where MGI.Status = '" & Left$(CmbSStatus.Text, 1) & "'"
   Else
      wc = wc & " And MGI.Status = '" & Left$(CmbSStatus.Text, 1) & "'"
   End If
End If

If TxtSViabilityID <> "" Then
   If wc = "" Then
      wc = " Where MGI.InspectionID = " & Val(TxtSViabilityID) & ""
   Else
      wc = wc & " And MGI.InspectionID = " & Val(TxtSViabilityID) & ""
   End If
End If

tmp = " Select MGI.InspectionID, MGI.Type"
tmp = tmp & " ,ML.LocationName, MGI.LocationID ,MCMP.CMPName, MGI.CMPID , isnull(MGI.GodownNo, MG.GodownNo) GodownNo, MGI.GodownID, MUT.UnitType, MGI.UnitTypeID , MGI.GodownCategoryID , MGI.LicenseID,  MC.Category, MLC.LicenseNo, MGI.Capacity , MGI.Address , MGI.Area ,MGI.Status ,MGI.StartDate ,MGI.CloseDate ,MGI.CapacityUtilisation"
tmp = tmp & " ,MGI.NCDEXStock ,MGI.NNNRStock ,MGI.ReservationStock ,MGI.G_UID ,MGI.Concurrency ,MGI.CreatedDate ,MGI.CreatedBy ,MGI.UpdatedDate ,MGI.UpdatedBy"
tmp = tmp & " ,MGI.noofstaff ,MGI.staffsalary ,MGI.noofsec ,MGI.secsalary ,MGI.totalrevenue ,MGI.directexpense ,MGI.grossprofit, MGI.indirectexpense ,MGI.netprofit  ,MGI.CreatedDate,MGI.UpdatedDate "
tmp = tmp & " from Mst_GodownInspection MGI"
tmp = tmp & " Left Join Mst_Location ML on ML.LocationID = MGI.LocationID"
tmp = tmp & " Left Join Mst_CMP MCMP on MCMP.CMPID = MGI.CMPID"
tmp = tmp & " Left Join Mst_Godown MG on MG.GodownID = MGI.GodownID"
tmp = tmp & " Left Join Mst_UnitType MUT on MUT.UnitTypeID = MGI.UnitTypeID"
tmp = tmp & " Left Join Mst_Category MC on MC.CategoryID = MGI.GodownCategoryID"
tmp = tmp & " Left Join Mst_License MLC on MLC.LicenseID = MGI.LicenseID"
tmp = tmp & wc & " Order By MGI.InspectionID"

Call TmpTable
With MSHFlexGrid1
   .Rows = .Rows + TmpRs.RecordCount
   For iu = 1 To TmpRs.RecordCount
      .TextMatrix(iu, 0) = IIf(IsNull(TmpRs!InspectionID), "", TmpRs!InspectionID)
      .TextMatrix(iu, 1) = IIf(IsNull(TmpRs!Type), "", TmpRs!Type)
      .TextMatrix(iu, 2) = IIf(IsNull(TmpRs!Locationname), "", TmpRs!Locationname) '
      .TextMatrix(iu, 3) = IIf(IsNull(TmpRs!CMPName), "", TmpRs!CMPName) '"CMP Name"
      .TextMatrix(iu, 4) = IIf(IsNull(TmpRs!GodownNo), "", TmpRs!GodownNo) '"Godown No"
      .TextMatrix(iu, 5) = IIf(IsNull(TmpRs!Capacity), "", TmpRs!Capacity) '"Capacity (in MT)"
      .TextMatrix(iu, 6) = IIf(IsNull(TmpRs!UnitType), "", TmpRs!UnitType) '"Godown Type"
      .TextMatrix(iu, 7) = IIf(IsNull(TmpRs!Area), "", TmpRs!Area) '"Area (in Sq.Ft)"
      .TextMatrix(iu, 8) = IIf(IsNull(TmpRs!Address), "", TmpRs!Address) '"Address"
      .TextMatrix(iu, 9) = IIf(IsNull(TmpRs!Status), "", TmpRs!Status) '"Status"
      .TextMatrix(iu, 10) = IIf(IsNull(TmpRs!StartDate), "", Format(TmpRs!StartDate, "dd-MMM-yyyy")) '"Start Date"
      .TextMatrix(iu, 11) = IIf(IsNull(TmpRs!CloseDate), "", Format(TmpRs!CloseDate, "dd-MMM-yyyy")) '"Close Date"
      .TextMatrix(iu, 12) = IIf(IsNull(TmpRs!noofstaff), "", TmpRs!noofstaff) '"No. of Staff"
      .TextMatrix(iu, 13) = IIf(IsNull(TmpRs!staffsalary), "", TmpRs!staffsalary) '"Staff Salary"
      .TextMatrix(iu, 14) = IIf(IsNull(TmpRs!noofsec), "", TmpRs!noofsec) '"No. of Sec Staff"
      .TextMatrix(iu, 15) = IIf(IsNull(TmpRs!secsalary), "", TmpRs!secsalary) '"Sec Staff Salary"
      .TextMatrix(iu, 16) = IIf(IsNull(TmpRs!CapacityUtilisation), "", TmpRs!CapacityUtilisation) '"Exp. Utilisation(%)"
      .TextMatrix(iu, 17) = IIf(IsNull(TmpRs!NcdexStock), "", TmpRs!NcdexStock) '"NCDEX Stock(%)"
      .TextMatrix(iu, 18) = IIf(IsNull(TmpRs!NNNRStock), "", TmpRs!NNNRStock) '"NNNR Stock(%)"
      .TextMatrix(iu, 19) = IIf(IsNull(TmpRs!ReservationStock), "", TmpRs!ReservationStock) '"Reservation Stock(%)"
      .TextMatrix(iu, 20) = IIf(IsNull(TmpRs!totalrevenue), "", TmpRs!totalrevenue) '"Total Revenue"
      .TextMatrix(iu, 21) = IIf(IsNull(TmpRs!DirectExpense), "", TmpRs!DirectExpense) '"Total Expense"
      .TextMatrix(iu, 22) = IIf(IsNull(TmpRs!GrossProfit), "", TmpRs!GrossProfit) '"Profit/Loss"
      .TextMatrix(iu, 23) = IIf(IsNull(TmpRs!IndirectExpense), "", TmpRs!IndirectExpense) '"Profit/Loss"
      .TextMatrix(iu, 24) = IIf(IsNull(TmpRs!NetProfit), "", TmpRs!NetProfit) '"Profit/Loss"
      .TextMatrix(iu, 25) = IIf(IsNull(TmpRs!CreatedDate), "", Format(TmpRs!CreatedDate, "dd-MMM-yyyy HH:MM:SSS")) '"Profit/Loss"
      .TextMatrix(iu, 26) = IIf(IsNull(TmpRs!UpdatedDate), "", Format(TmpRs!UpdatedDate, "dd-MMM-yyyy HH:MM:SSS")) '"Profit/Loss"
      
      If LCase(.TextMatrix(iu, 1)) = "n" Then
         .TextMatrix(iu, 1) = "New"
      ElseIf LCase(.TextMatrix(iu, 1)) = "r" Then
         .TextMatrix(iu, 1) = "Renewal"
      End If
      If LCase(.TextMatrix(iu, 9)) = "p" Then
         .TextMatrix(iu, 9) = "Pending"
      ElseIf LCase(.TextMatrix(iu, 9)) = "a" Then
         .TextMatrix(iu, 9) = "Approved"
      ElseIf LCase(.TextMatrix(iu, 9)) = "r" Then
         .TextMatrix(iu, 9) = "Rejected"
      ElseIf LCase(.TextMatrix(iu, 9)) = "g" Then
         .TextMatrix(iu, 9) = "Godown Creation"
      End If
      TmpRs.MoveNext
   Next
End With
End Sub

Private Sub CmdGodownExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & "Mst_LiveGodownViabilityDetails.xls")
Call Xls_Detail_Rpt1
End Sub

Public Sub Xls_Detail_Rpt1()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Integer
Dim mCol As Integer
MsgBox ("Wait till next message!!!")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & "Mst_LiveGodownViabilityDetails.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid2.Rows - 1
   mRow = I + 1 ' + 6
   For C = 0 To MSHFlexGrid2.Cols - 1
      oWS.Cells(mRow, C + 1) = MSHFlexGrid2.TextMatrix(I, C)
   Next
Next
oWB.Save
MsgBox ("Excel file created!!!")
oXL.Visible = True

End Sub

Private Sub CmdStorageRate_Click()
If Trim(CmbLocation.Text) = "" Then
   MsgBox "Select Location first!!!"
   Exit Sub
End If
If Trim(CmbCategory.Text) = "" Then
   MsgBox "Select Category first!!!"
   Exit Sub
End If
If Trim(CmbCommodity.Text) = "" Then
   MsgBox "Select Commodity first!!!"
   Exit Sub
End If
Gen_GodownInspection(0) = mLocationID
tmp = " Select CommodityGroupID from Mst_Commodity Where CommodityID = " & mCommodityID
Call TmpTable
Gen_GodownInspection(1) = IIf(IsNull(TmpRs!CommodityGroupID), 0, TmpRs!CommodityGroupID)

If LCase(CmbCategory.Text) = "ncdex" Then
   FrmMst_NCDEXStorageRateCard.Show
ElseIf LCase(CmbCategory.Text) = "nnnr" Then
   FrmMst_StorageRateCard.Show
End If
End Sub

Private Sub Command1_Click()
FrmPreRent.Visible = False
End Sub

Private Sub EditCmd_Click()

If LCase(TxtStatus.Text) = "rejected" Or LCase(TxtStatus.Text) = "godown creation" Then
   MsgBox "Viability is " & TxtStatus.Text & ", You can not edit!!!"
   Exit Sub
End If

If LCase(TxtStatus.Text) = "approved" Or LCase(TxtStatus.Text) = "godown creation" Then
   If LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "superadmin" And LCase(Gen_UserRole) <> "headcmg" Then
      MsgBox "Rights for Approval/Godown Creation are with CMG Department. You can't edit!!! "
      Exit Sub
   End If
End If


Edit_Flg = "E"
Call GButtonLock(Me, False)

Picture2.Enabled = True
Frame1.Enabled = True
Frame2.Enabled = True

'TxtStaffSalaryUser.Locked = False
'TxtStaffUser.Locked = False
'TxtSecStaffUser.Locked = False
'TxtSecStaffSalaryUser.Locked = False
'TxtCapacity.Locked = False
'TxtAddress.Locked = False
'TxtArea.Locked = False
'TxtGodown.Locked = False

TxtRemarks.Locked = False
TxtExpUtilisation.Locked = False
TxtNCDEXStock.Locked = False
TxtNNNRStock.Locked = False
TxtReservationStock.Locked = False
TxtStartDate.Enabled = True
TxtCloseDate.Enabled = True

TxtStatus.Visible = False
CmbStatus.Visible = True

Call FillExpenseDetails
Call Grid_Head7

TxtCurrentUtilisation.Text = ""
TxtTotalRevenue.Text = ""
TxtDirectExpense.Text = ""
TxtDirectExpense.Text = ""
TxtGrossProfit.Text = ""
TxtIndirectExpense.Text = ""
TxtIndirectExpense.Text = ""
TxtNetProfit.Text = ""
TxtCMShare.Text = ""

If mGodownCategoryID = 2 Then
   If Val(mGodownTypeID) = 1 Or Val(mGodownTypeID) = 7 Then
      TxtArea.Locked = True
      TxtArea.BackColor = &H80000004
      If LCase(Gen_UserRole) = "power" Or LCase(Gen_UserRole) = "superadmin" Or LCase(Gen_UserRole) = "headcmg" Then
         Frame6.Enabled = True
      Else
         Frame6.Enabled = False
      End If
   Else
      TxtArea.Locked = False
      TxtArea.BackColor = &H80000005
      Frame6.Enabled = False
   End If
Else
   TxtArea.Locked = False
   TxtArea.BackColor = &H80000005
   Frame6.Enabled = False
End If

End Sub

Private Sub DeleteCmd_Click()
Dim ans As Variant
On Error GoTo msgError
If LCase(TxtStatus.Text) = "rejected" Or LCase(TxtStatus.Text) = "godown creation" Then
   MsgBox "Viability is " & TxtStatus.Text & ", You can not edit!!!"
   Exit Sub
End If
ans = MsgBox("Do you really want to DELETE this record???", vbYesNo)
If ans = vbYes Then
   tmp = "Delete from Mst_InspectionGodownDetail where InspectionID = " & Val(TxtInspectionID.Text)
   Call TmpTable
   
   tmp = "Delete from Mst_InspectionCommodityDetail where InspectionID = " & Val(TxtInspectionID.Text)
   Call TmpTable
   
   tmp = "Delete from Mst_InspectionExpenseDetail where InspectionID = " & Val(TxtInspectionID.Text)
   Call TmpTable
   
   
   RsMst_GodownInspection.Find "InspectionID = " & Val(TxtInspectionID.Text)
   If RsMst_GodownInspection.EOF Then Exit Sub
   RsMst_GodownInspection.Delete
   RsMst_GodownInspection.Update
   If RsMst_GodownInspection.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_GodownInspection.MoveNext
   If RsMst_GodownInspection.EOF() Then
      RsMst_GodownInspection.MoveLast
   End If
   Call Indata
End If
Exit Sub
msgError:
MsgBox "Child record Present "
RsMst_GodownInspection.CancelUpdate
End Sub

Private Sub ExitCmd_Click()
Unload Me
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


Private Sub LblViewPrevRent_Click()
If mLExpenseID = 1 And TxtExpenseType.Text <> "" Then
   FrmPreRent.Visible = True
End If
End Sub

Private Sub MSHFlexGrid1_Click()
Dim I As Variant
I = MSHFlexGrid1.RowSel
If I < 1 Then Exit Sub
If MSHFlexGrid1.TextMatrix(I, 0) = "" Then Exit Sub

RsMst_GodownInspection.MoveFirst
RsMst_GodownInspection.Find " InspectionID= " & Val(MSHFlexGrid1.TextMatrix(I, 0))

If Not RsMst_GodownInspection.EOF Then
   Call Indata
   Frame0.Visible = False
End If
End Sub

Private Sub MSHFlexGrid3_Click()
If SaveCmd.Enabled = False Then Exit Sub
With MSHFlexGrid3
   I = .RowSel
   If I < 2 Then Exit Sub
   If .TextMatrix(I, 1) = "" Then Exit Sub

   CmbCategory.Text = .TextMatrix(I, 0)
   CmbCommodity.Text = .TextMatrix(I, 1)
   TxtExpQtyP.Text = .TextMatrix(I, 2)
   TxtExpQtyMT.Text = .TextMatrix(I, 3)
   TxtStorageRate.Text = .TextMatrix(I, 4)
   TxtPrice.Text = .TextMatrix(I, 5)
   TxtRevenue.Text = .TextMatrix(I, 6)
   TxtAUM.Text = .TextMatrix(I, 7)
   TxtStorageRemarks.Text = .TextMatrix(I, 8)
   mCommodityID = .TextMatrix(I, 10)
   TxtFumigableQty.Text = .TextMatrix(I, 11)
   mFumigationFlag = .TextMatrix(I, 12)
   TxtCMShareP.Text = .TextMatrix(I, 13)
   TxtCMShareR.Text = .TextMatrix(I, 14)
   
   CmdCAddToList.Caption = "Update List"
   
   CmbCategory.Enabled = False
   CmbCommodity.Enabled = False
   'If LCase(mFumigationFlag) = "y" Then
   '   TxtFumigableQty.Locked = False
   'Else
   '   TxtFumigableQty.Locked = True
   '   TxtFumigableQty.Text = ""
   'End If
End With
If LCase(CmbCategory.Text) = "ncdex" Or LCase(CmbCategory.Text) = "nnnr" Then
   CmdStorageRate.Enabled = True
Else
   CmdStorageRate.Enabled = False
End If
If LCase(CmbCategory.Text) = "nnnr" Then
   TxtCMShareP.Enabled = True
Else
   TxtCMShareP.Enabled = False
End If
CmdCDeleteFromList.Enabled = True
End Sub

Private Sub MSHFlexGrid4_Click()
'On Error GoTo ErrorHandler
Call ClearExpenseDetail
If SaveCmd.Enabled = False Then Exit Sub
With MSHFlexGrid4
   I = .RowSel
   If I < 2 Then Exit Sub
   If .TextMatrix(I, 1) = "" Then Exit Sub

   TxtExpenseType.Text = .TextMatrix(I, 0)
   TxtCalculationOn.Text = .TextMatrix(I, 1)
   TxtQtySq.Text = .TextMatrix(I, 2)
   TxtAmount.Text = .TextMatrix(I, 3)
   TxtExpenseRemarks.Text = .TextMatrix(I, 4)
   TxtAmountPer.Text = .TextMatrix(I, 7)
   TxtFormula.Text = .TextMatrix(I, 8)
   mUserInputAmount = Val(.TextMatrix(I, 9))
   mType = .TextMatrix(I, 10)
   mUserInputCalc = Val(.TextMatrix(I, 11))
   mClearFlag = Val(.TextMatrix(I, 12))
   TxtRemarksExpense.Text = .TextMatrix(I, 13)
   mLExpenseID = Val(.TextMatrix(I, 6))
   If mUserInputAmount = 1 Then
      TxtAmountPer.Locked = False
      TxtAmountPer.BackColor = &H80000005
   Else
      TxtAmountPer.Locked = True
      TxtAmountPer.BackColor = &H80000004
   End If
   If mUserInputCalc = 1 Then
      TxtQtySq.Locked = False
      TxtQtySq.BackColor = &H80000005
   Else
      TxtQtySq.Locked = True
      TxtQtySq.BackColor = &H80000004
   End If
''   If TxtQtySq.Text = "" Then
''      If LCase(TxtCalculationOn.Text) = "qty(per mt)" Then
''         TxtQtySq.Text = Txtcapacity.Text
''         TxtAmount.Text = Format(Val(TxtAmountPer.Text) * Val(TxtQtySq.Text), "#0")
''      ElseIf LCase(TxtCalculationOn.Text) = "area(per sq.ft)" Then
''         TxtQtySq.Text = TxtArea.Text
''         TxtAmount.Text = Format(Val(TxtAmountPer.Text) * Val(TxtQtySq.Text), "#0")
''      ElseIf LCase(TxtCalculationOn.Text) = "fumigation qty(per mt)" Then
''         TxtQtySq.Text = MSHFlexGrid3.TextMatrix(0, 11)
''         TxtAmount.Text = Format(Val(TxtAmountPer.Text) * Val(TxtQtySq.Text), "#0")
''      ElseIf LCase(TxtCalculationOn.Text) = "fixed" Then
''         TxtQtySq.Text = 1
''         TxtAmount.Text = TxtAmountPer.Text
''         TxtQtySq.Locked = True
''      End If
''   End If

   If Val(TxtQtySq.Text) = 0 Then
      If LCase(TxtCalculationOn.Text) = "qty(per mt)" Then
         If Val(TxtCapacity.Text) > Val(TxtExpUtilisationMT.Text) Then
            TxtQtySq.Text = TxtCapacity.Text
         Else
            TxtQtySq.Text = TxtExpUtilisationMT.Text
         End If
      ElseIf LCase(TxtCalculationOn.Text) = "area(per sq.ft)" Then
         TxtQtySq.Text = TxtArea.Text
      ElseIf LCase(TxtCalculationOn.Text) = "fixed" Then
         TxtQtySq.Text = 0
      ElseIf LCase(TxtCalculationOn.Text) = "fumigation qty(per mt)" Then
         TxtQtySq.Text = MSHFlexGrid3.TextMatrix(0, 11)
      ElseIf LCase(TxtCalculationOn.Text) = "per aum" Then
         TxtQtySq.Text = MSHFlexGrid3.TextMatrix(0, 7)
      End If
   End If
   
   If Trim(TxtExpenseType.Text) <> "" Then
      If mLExpenseID = 1 Then
         If LCase(CmbType.Text) <> "new" Then
            Call CalculateRentChange
         End If
      End If
   End If
   
   If Trim(TxtFormula.Text) <> "" Then
      TxtFormula.Text = Replace(TxtFormula.Text, "godownarea", "Area in Sq.Ft.", 1, , vbTextCompare)
      TxtFormula.Text = Replace(TxtFormula.Text, "amountperunit", "Amount per Unit", 1, , vbTextCompare)
      TxtFormula.Text = Replace(TxtFormula.Text, "godowncapacity", "Capacity in MT", 1, , vbTextCompare)
      TxtFormula.Text = Replace(TxtFormula.Text, "totalcapacity", "Location Capacity", 1, , vbTextCompare)
      TxtFormula.Text = Replace(TxtFormula.Text, "aum", "Total AUM", 1, , vbTextCompare)
      TxtFormula.Text = Replace(TxtFormula.Text, "fumigableqty", "Total fumigable Qty", 1, , vbTextCompare)
      TxtFormula.Text = Replace(TxtFormula.Text, "godownincome", "Total Revenue", 1, , vbTextCompare)
      TxtFormula.Text = Replace(TxtFormula.Text, "coexpense", "Co Expense", 1, , vbTextCompare)
      TxtFormula.Text = Replace(TxtFormula.Text, "cmpbudgetincome", "Budget Income", 1, , vbTextCompare)
      TxtFormula.Text = Replace(TxtFormula.Text, "cmpbankinterest", "Bank Interest", 1, , vbTextCompare)
      TxtFormula.Text = Replace(TxtFormula.Text, "ncdexincome", "NCDEX Income", 1, , vbTextCompare)
   End If
   Call AmountCalculation
End With
If I > MSHFlexGrid4.Rows - 4 Then Exit Sub

CmdEUpdateList.Enabled = True
CmdClearData.Enabled = True

Exit Sub
ErrorHandler:
   MsgBox "Calculation is not done correctly!!!"
End Sub

Private Sub AmountCalculation()
On Error GoTo ErrorHandler
With MSHFlexGrid4
   I = .RowSel
   mFormula = .TextMatrix(I, 8)
   If mFormula = "" Then
      TxtAmount.Text = TxtAmountPer.Text
   Else
      If LCase(TxtCalculationOn.Text) = "qty(per mt)" Then
         mFormula = Replace(mFormula, "godowncapacity", Val(TxtQtySq.Text), 1, , vbTextCompare)
      ElseIf LCase(TxtCalculationOn.Text) = "area(per sq.ft)" Then
         mFormula = Replace(mFormula, "godownarea", Val(TxtQtySq.Text), 1, , vbTextCompare)
      'ElseIf LCase(TxtCalculationOn.Text) = "fixed" Then
      '   TxtQtySq.Text = 0
      ElseIf LCase(TxtCalculationOn.Text) = "fumigation qty(per mt)" Then
         mFormula = Replace(mFormula, "fumigableqty", Val(TxtQtySq.Text), 1, , vbTextCompare)
      ElseIf LCase(TxtCalculationOn.Text) = "per aum" Then
         mFormula = Replace(mFormula, "aum", Val(TxtQtySq.Text), 1, , vbTextCompare)
      End If
      mFormula = Replace(mFormula, "godownarea", Val(TxtArea.Text), 1, , vbTextCompare)
      mFormula = Replace(mFormula, "amountperunit", Val(TxtAmountPer.Text), 1, , vbTextCompare)
      If Val(TxtCapacity.Text) > Val(TxtExpUtilisationMT.Text) Then
         mFormula = Replace(mFormula, "godowncapacity", Val(TxtCapacity.Text), 1, , vbTextCompare)
      Else
         mFormula = Replace(mFormula, "godowncapacity", Val(TxtExpUtilisationMT.Text), 1, , vbTextCompare)
      End If
      
      If LCase(CmbType.Text) = "renewal" Then
         mFormula = Replace(mFormula, "totalcapacity", Val(MSHFlexGrid2.TextMatrix(0, 3)), 1, , vbTextCompare)
      ElseIf LCase(CmbType.Text) = "new" Then
         If Val(TxtCapacity.Text) > Val(TxtExpUtilisationMT.Text) Then
            mFormula = Replace(mFormula, "totalcapacity", Val(MSHFlexGrid2.TextMatrix(0, 3)) + Val(TxtCapacity.Text), 1, , vbTextCompare)
         Else
            mFormula = Replace(mFormula, "totalcapacity", Val(MSHFlexGrid2.TextMatrix(0, 3)) + Val(TxtExpUtilisationMT.Text), 1, , vbTextCompare)
         End If
         
      End If
      mFormula = Replace(mFormula, "aum", Val(MSHFlexGrid3.TextMatrix(0, 7)), 1, , vbTextCompare)
      mFormula = Replace(mFormula, "fumigableqty", Val(MSHFlexGrid3.TextMatrix(0, 11)), 1, , vbTextCompare)
      mFormula = Replace(mFormula, "godownincome", Val(MSHFlexGrid3.TextMatrix(0, 6)), 1, , vbTextCompare)
      mFormula = Replace(mFormula, "month", DateDiff("m", TxtStartDate.Value, TxtCloseDate.Value) + 1, 1, , vbTextCompare)
      
      If InStr(1, mFormula, "coexpense", vbTextCompare) > 0 Then
         tmp = " Select ExpenseAmt From Mst_CoExpense Where '" & Format(Date, "yyyy-mm-dd") & "' between FromDate and ToDate"
         Call Tmp7Table
         If TmpRs7.RecordCount > 0 Then
            mFormula = Replace(mFormula, "coexpense", IIf(IsNull(TmpRs7!ExpenseAmt), 0, TmpRs7!ExpenseAmt), 1, , vbTextCompare)
         Else
            mFormula = Replace(mFormula, "coexpense", 1, 1, , vbTextCompare)
         End If
      End If
      If InStr(1, mFormula, "cmpbudgetincome", vbTextCompare) > 0 Then
         tmp = " Select BudgetIncome From Mst_CoExpense Where '" & Format(Date, "yyyy-mm-dd") & "' between FromDate and ToDate"
         Call Tmp7Table
         If TmpRs7.RecordCount > 0 Then
            mFormula = Replace(mFormula, "cmpbudgetincome", IIf(IsNull(TmpRs7!BudgetIncome), 0, TmpRs7!BudgetIncome), 1, , vbTextCompare)
         Else
            mFormula = Replace(mFormula, "cmpbudgetincome", 1, 1, , vbTextCompare)
         End If
      End If
      If InStr(1, mFormula, "cmpbankinterest", vbTextCompare) > 0 Then
         tmp = " Select BankInterest From Mst_CoExpense Where '" & Format(Date, "yyyy-mm-dd") & "' between FromDate and ToDate"
         Call Tmp7Table
         If TmpRs7.RecordCount > 0 Then
            mFormula = Replace(mFormula, "cmpbankinterest", IIf(IsNull(TmpRs7!BankInterest), 0, TmpRs7!BankInterest), 1, , vbTextCompare)
         Else
            mFormula = Replace(mFormula, "cmpbankinterest", 1, 1, , vbTextCompare)
         End If
      End If
      If InStr(1, mFormula, "aorocmgexpense", vbTextCompare) > 0 Then
         tmp = " Select aorocmgexpense From Mst_CoExpense Where '" & Format(Date, "yyyy-mm-dd") & "' between FromDate and ToDate"
         Call Tmp7Table
         If TmpRs7.RecordCount > 0 Then
            mFormula = Replace(mFormula, "aorocmgexpense", IIf(IsNull(TmpRs7!aorocmgexpense), 0, TmpRs7!aorocmgexpense), 1, , vbTextCompare)
         Else
            mFormula = Replace(mFormula, "aorocmgexpense", 1, 1, , vbTextCompare)
         End If
      End If

      mNCDEXIncome = 0
      For k = 2 To MSHFlexGrid3.Rows - 2
         If LCase(MSHFlexGrid3.TextMatrix(k, 0)) = "ncdex" Then
            mNCDEXIncome = mNCDEXIncome + Val(MSHFlexGrid3.TextMatrix(k, 6))
         End If
      Next
      mFormula = Replace(mFormula, "ncdexincome", mNCDEXIncome, 1, , vbTextCompare)
      TxtAmount.Text = Format(SC.Eval(mFormula), "#0")
   End If
End With
Exit Sub
ErrorHandler:
   MsgBox "Calculation is not done correctly!!!"
End Sub

Private Sub MSHFlexGrid6_Click()

If SaveCmd.Enabled = False Then Exit Sub


'If LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "superadmin" And LCase(Gen_UserRole) <> "head cmg" And LCase(Gen_UserRole) <> "admin" Then
'   Exit Sub
'End If

With MSHFlexGrid6
   mColsel = .ColSel
   mRowsel = .RowSel
   If mColsel = 3 Or mColsel = 4 Then
      If mRowsel = 2 Or mRowsel = 3 Then
         Call VisibleEnterTxt
      End If
   End If
End With

End Sub

Private Sub MSHFlexGrid8_Click()
With MSHFlexGrid8
     If .Col = 0 Then
        If .TextMatrix(.row, .Col) = "" Then Exit Sub
        .Col = 0
        If .TextMatrix(.row, .Col) = strUnChecked Then
           .TextMatrix(.row, .Col) = strChecked
        Else
           .TextMatrix(.row, .Col) = strUnChecked
        End If
     End If
End With
End Sub

Private Sub MSHFlexGrid9_Click()

If SaveCmd.Enabled = False Then Exit Sub
Dim r As Variant
r = MSHFlexGrid9.RowSel
If r < 2 Then Exit Sub
If MSHFlexGrid9.TextMatrix(r, 1) = "" Then Exit Sub
If Frame6.Enabled = False Then Exit Sub
With MSHFlexGrid9
   TxtLength = Val(.TextMatrix(r, 1))
   TxtWidth = Val(.TextMatrix(r, 2))
   TxtHeight = Val(.TextMatrix(r, 3))
   TxtSubArea = Val(.TextMatrix(r, 4))
   CmdAreaAddToList.Caption = "Update List"
   CmdAreaDeleteFromList.Enabled = True
End With


End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_GodownInspection.MoveNext
If RsMst_GodownInspection.EOF Then
   RsMst_GodownInspection.MoveLast
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
RsMst_GodownInspection.MovePrevious
If RsMst_GodownInspection.BOF Then
   RsMst_GodownInspection.MoveFirst
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
RsMst_GodownInspection.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub

Private Sub LastCmd_Click()
RsMst_GodownInspection.MoveLast
Call Indata
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False
End Sub


Private Sub SearchCmd_Click()

If SearchCmd.Caption = "Cancel" Then
   If RsMst_GodownInspection.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   Call Indata
   Exit Sub
End If

Frame0.Height = Me.Height - 600
Frame0.Left = Picture2.Left
Frame0.Top = Picture2.Top
Frame0.Width = Picture2.Width + 100
Frame0.Visible = True

MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 1800

ProgressBar1.Width = MSHFlexGrid1.Width
ProgressBar1.Left = MSHFlexGrid1.Left
ProgressBar1.Top = MSHFlexGrid1.Top + MSHFlexGrid1.Height



End Sub

Private Sub TxtAmountPer_LostFocus()
If Trim(TxtAmountPer.Text) = "" Then Exit Sub
If IsNumeric(TxtAmountPer.Text) = False Or Val(TxtAmountPer.Text) < 0 Then
   MsgBox "Invalid numeric format"
   TxtAmountPer.SetFocus
   Exit Sub
End If

TxtAmountPer.Text = Format(TxtAmountPer.Text, "#0.0000")
If Trim(TxtExpenseType.Text) <> "" Then
   Call AmountCalculation
End If

If Trim(TxtExpenseType.Text) <> "" Then
'   Call AmountCalculation
   If mLExpenseID = 1 Then
      If LCase(CmbType.Text) <> "new" Then
         Call CalculateRentChange
      End If
   End If
End If

End Sub

Private Sub TxtArea_GotFocus()
tmp = TxtArea.Text
End Sub

Private Sub TxtCapacity_GotFocus()
tmp = TxtCapacity.Text
End Sub

Private Sub TxtCloseDate_GotFocus()
tmp = TxtCloseDate.Value
End Sub

Private Sub TxtCloseDate_LostFocus()
If TxtStartDate.Value > TxtCloseDate.Value Then
   MsgBox "Start Date should not greater than close date!!!"
   TxtCloseDate.SetFocus
   Exit Sub
End If
If TxtCloseDate.Value <> tmp Then
   If Val(MSHFlexGrid4.TextMatrix(0, 3)) > Val(MSHFlexGrid4.TextMatrix(MSHFlexGrid4.Rows - 2, 3)) + Val(MSHFlexGrid4.TextMatrix(MSHFlexGrid4.Rows - 3, 3)) Then
      MsgBox "Expense will be clear!!!"
      Call ClearAllExpenseDetails
   End If
End If
End Sub

Private Sub TxtCMShareP_LostFocus()
If Trim(TxtCMShareP.Text) = "" Then Exit Sub
If IsNumeric(TxtCMShareP.Text) = False Or Val(TxtCMShareP.Text) < 0 Then
   MsgBox "Invalid numeric format"
   TxtCMShareP.SetFocus
   Exit Sub
End If
If Val(TxtCMShareP.Text) > 100 Then
   MsgBox "CM Share (in %) never greater than 100!!!"
   TxtCMShareP.SetFocus
   Exit Sub
End If
TxtCMShareP.Text = Format(TxtCMShareP.Text, "#0")
TxtCMShareR.Text = Format(Val(TxtAUM.Text) * Val(TxtCMShareP.Text) * 0.00007 * 60 / 1200, "#0")
End Sub

Private Sub TxtEnterData_LostFocus()
If Trim(TxtEnterData.Text) = "" Then
   TxtEnterData.Visible = False
   Exit Sub
End If
If IsNumeric(TxtEnterData.Text) = False Or Val(TxtEnterData.Text) < 0 Then
   MsgBox "Invalid numeric format!!!"
   TxtEnterData.SetFocus
   Exit Sub
End If

TxtEnterData.Text = Format(TxtEnterData.Text, "#0")
If LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "superadmin" And LCase(Gen_UserRole) <> "headcmg" And LCase(Gen_UserRole) <> "admin" Then
    If mRowsel = 2 Then
       If mColsel = 3 Then
          If Val(TxtEnterData.Text) < Val(MSHFlexGrid6.TextMatrix(2, 1)) Then
             MsgBox "No of Staff should not less than existing"
             TxtEnterData.SetFocus
             Exit Sub
          End If
       End If
       If mColsel = 4 Then
          If Val(TxtEnterData.Text) < Val(MSHFlexGrid6.TextMatrix(2, 2)) Then
             MsgBox "Salary cost should not less than existing"
             TxtEnterData.SetFocus
             Exit Sub
          End If
       End If
    End If
    
    If mRowsel = 3 Then
       If mColsel = 3 Then
          If Val(TxtEnterData.Text) < Val(MSHFlexGrid6.TextMatrix(3, 1)) Then
             MsgBox "No of Security should not less than existing"
             TxtEnterData.SetFocus
             Exit Sub
          End If
       End If
       If mColsel = 4 Then
          If Val(TxtEnterData.Text) < Val(MSHFlexGrid6.TextMatrix(3, 2)) Then
             MsgBox "Security cost should not less than existing"
             TxtEnterData.SetFocus
             Exit Sub
          End If
       End If
    End If
End If
MSHFlexGrid6.TextMatrix(mRowsel, mColsel) = Val(TxtEnterData.Text)
TxtEnterData.Visible = False

'====================='for security and staff salary expense'=====================
MSHFlexGrid4.TextMatrix(0, 3) = Format(Val(MSHFlexGrid4.TextMatrix(0, 3)) - Val(MSHFlexGrid4.TextMatrix(MSHFlexGrid4.Rows - 2, 3)), "#0")
MSHFlexGrid4.TextMatrix(0, 3) = Format(Val(MSHFlexGrid4.TextMatrix(0, 3)) - Val(MSHFlexGrid4.TextMatrix(MSHFlexGrid4.Rows - 3, 3)), "#0")
MSHFlexGrid4.Rows = MSHFlexGrid4.Rows - 2
Call SecandUserSalaryExpense
'====================='for security and staff salary expense'=====================
End Sub

Private Sub TxtExpUtilisation_GotFocus()
tmp = TxtExpUtilisation.Text
End Sub

Private Sub TxtFumigableQty_LostFocus()
If Trim(TxtFumigableQty.Text) = "" Then Exit Sub
If IsNumeric(TxtFumigableQty.Text) = False Or Val(TxtFumigableQty.Text) < 0 Then
   MsgBox "Invalid numeric format"
   TxtFumigableQty.SetFocus
   Exit Sub
End If
TxtFumigableQty.Text = Format(TxtFumigableQty.Text, "#0")
End Sub

Private Sub TxtGodown_LostFocus()

If TxtGodown = "" Then Exit Sub

If IsClean(TxtGodown) = False Then
   MsgBox "Special Character not allowed!!!"
   TxtGodown.SetFocus
   Exit Sub
End If

TxtGodown = Trim(TxtGodown)

End Sub

Private Sub TxtHeight_LostFocus()

mHeight = 0
If Trim(TxtHeight.Text) = "" Then
   Call CalArea
   Exit Sub
End If
If IsNumeric(TxtHeight.Text) = False Or Val(TxtHeight.Text) < 0 Then
   MsgBox "Invalid numeric format"
   TxtHeight.SetFocus
   Exit Sub
End If

mHeight = Val(TxtHeight)

Call CalArea

End Sub

Private Sub TxtLength_LostFocus()

mLength = 0
If Trim(TxtLength.Text) = "" Then
   Call CalArea
   Exit Sub
End If

If IsNumeric(TxtLength.Text) = False Or Val(TxtLength.Text) < 0 Then
   MsgBox "Invalid numeric format"
   TxtLength.SetFocus
   Exit Sub
End If


TxtLength = Format(TxtLength, "########0.00")
Dim strSplitString() As String
Dim mLPart1, mLPart2 As Variant

strSplitString = Split(TxtLength, ".")
If UBound(strSplitString) <> 1 Then
   MsgBox "Invalid Selection "
   TxtLength.SetFocus
   Exit Sub
End If

If Val(strSplitString(1)) >= 12 Then
   MsgBox "Value after decimal should be less than 12"
   TxtLength.SetFocus
   Exit Sub
End If

mLPart1 = Val(strSplitString(0))
mLPart2 = Format(Val(strSplitString(1)) / 120, "#######0.00")

mLength = Val(mLPart1) + Val(mLPart2) 'Format(TxtLength.Text, "#0")


'MSHFlexGrid9.TextMatrix(mSelRow, mSelCol) = Val(mLength)
'MSHFlexGrid2.TextMatrix(0, 7) = Val(MSHFlexGrid2.TextMatrix(0, 7)) - Val(MSHFlexGrid2.TextMatrix(mSelRow, 7))
'MSHFlexGrid2.TextMatrix(mSelRow, 8) = Val(MSHFlexGrid2.TextMatrix(mSelRow, 6)) + Val(TxtArrears)

Call CalArea

End Sub
Private Sub CalArea()

If SaveCmd.Enabled = False Then Exit Sub

If Val(TxtLength) = 0 Then
   TxtSubArea = 0
   Exit Sub
End If

If Val(TxtWidth) = 0 Then
   TxtSubArea = 0
   Exit Sub
End If

If Val(TxtHeight) = 0 Then
   TxtSubArea = 0
   Exit Sub
End If

If Val(mLength) = 0 Then
   mLength = Val(TxtLength)
End If
If Val(mWidth) = 0 Then
   mWidth = Val(TxtWidth)
End If
mAdhocArea = Val(mLength) * Val(mWidth)

If Val(mHeight) = 0 Then
   mHeight = Val(TxtHeight)
End If

If Val(mHeight) > 15 Then
   TxtSubArea = Format(Val(mAdhocArea), "#######0.00")
Else
   TxtSubArea = Format(Val(mAdhocArea) * Val(mHeight) / 15, "#######0.00")
End If

'Call TxtArea_LostFocus

End Sub

Private Sub TxtNCDEXStock_GotFocus()
tmp = TxtNCDEXStock.Text
End Sub

Private Sub TxtNNNRStock_GotFocus()
tmp = TxtNNNRStock.Text
End Sub

Private Sub TxtNoofStack_LostFocus()
If TxtNoofStack.Locked = True Then Exit Sub
If Trim(TxtNoofStack.Text) = "" Then Exit Sub
If IsNumeric(TxtNoofStack.Text) = False Or Val(TxtNoofStack.Text) < 0 Then
   MsgBox "Invalid numeric format"
   TxtNoofStack.SetFocus
   Exit Sub
End If
'
TxtNoofStack.Text = Format(TxtNoofStack.Text, "#0")

If mGodownCategoryID = 2 And Val(mGodownTypeID) = 7 Then
   TxtCapacity.Text = Format(Val(TxtNoofStack) * Val(TxtWeightPerStack), "#0.00")
Else
   If TxtArea <> "" Then
      Call TxtArea_LostFocus
   Else
      TxtCapacity.Text = ""
   End If
End If

End Sub

Private Sub TxtReservationStock_GotFocus()
tmp = TxtReservationStock.Text
End Sub

Private Sub TxtSecStaffUser_LostFocus()
If Trim(TxtSecStaffUser.Text) = "" Then Exit Sub
If IsNumeric(TxtSecStaffUser.Text) = False Or Val(TxtSecStaffUser.Text) < 0 Then
   MsgBox "Invalid numeric format"
   TxtSecStaffUser.SetFocus
   Exit Sub
End If

TxtSecStaffUser.Text = Format(TxtSecStaffUser.Text, "#0")
End Sub

Private Sub TxtSecStaffSalaryUser_LostFocus()
If Trim(TxtSecStaffSalaryUser.Text) = "" Then Exit Sub
If IsNumeric(TxtSecStaffSalaryUser.Text) = False Or Val(TxtSecStaffSalaryUser.Text) < 0 Then
   MsgBox "Invalid numeric format"
   TxtSecStaffSalaryUser.SetFocus
   Exit Sub
End If

TxtSecStaffSalaryUser.Text = Format(TxtSecStaffSalaryUser.Text, "#0")
'====================='for security and staff salary expense'=====================
MSHFlexGrid4.TextMatrix(0, 3) = Format(Val(MSHFlexGrid4.TextMatrix(0, 3)) - Val(MSHFlexGrid4.TextMatrix(MSHFlexGrid4.Rows - 2, 3)), "#0")
MSHFlexGrid4.TextMatrix(0, 3) = Format(Val(MSHFlexGrid4.TextMatrix(0, 3)) - Val(MSHFlexGrid4.TextMatrix(MSHFlexGrid4.Rows - 3, 3)), "#0")
MSHFlexGrid4.Rows = MSHFlexGrid4.Rows - 2
Call SecandUserSalaryExpense
'====================='for security and staff salary expense'=====================
End Sub

Private Sub TxtStaffUser_LostFocus()
If Trim(TxtStaffUser.Text) = "" Then Exit Sub
If IsNumeric(TxtStaffUser.Text) = False Or Val(TxtStaffUser.Text) < 0 Then
   MsgBox "Invalid numeric format"
   TxtStaffUser.SetFocus
   Exit Sub
End If

TxtStaffUser.Text = Format(TxtStaffUser.Text, "#0")
'====================='for security and staff salary expense'=====================
MSHFlexGrid4.TextMatrix(0, 3) = Format(Val(MSHFlexGrid4.TextMatrix(0, 3)) - Val(MSHFlexGrid4.TextMatrix(MSHFlexGrid4.Rows - 2, 3)), "#0")
MSHFlexGrid4.TextMatrix(0, 3) = Format(Val(MSHFlexGrid4.TextMatrix(0, 3)) - Val(MSHFlexGrid4.TextMatrix(MSHFlexGrid4.Rows - 3, 3)), "#0")
MSHFlexGrid4.Rows = MSHFlexGrid4.Rows - 2
Call SecandUserSalaryExpense
'====================='for security and staff salary expense'=====================

End Sub

Private Sub TxtStaffSalaryUser_LostFocus()
If Trim(TxtStaffSalaryUser.Text) = "" Then Exit Sub
If IsNumeric(TxtStaffSalaryUser.Text) = False Or Val(TxtStaffSalaryUser.Text) < 0 Then
   MsgBox "Invalid numeric format"
   TxtStaffSalaryUser.SetFocus
   Exit Sub
End If

TxtStaffSalaryUser.Text = Format(TxtStaffSalaryUser.Text, "#0")
'====================='for security and staff salary expense'=====================
MSHFlexGrid4.TextMatrix(0, 3) = Format(Val(MSHFlexGrid4.TextMatrix(0, 3)) - Val(MSHFlexGrid4.TextMatrix(MSHFlexGrid4.Rows - 2, 3)), "#0")
MSHFlexGrid4.TextMatrix(0, 3) = Format(Val(MSHFlexGrid4.TextMatrix(0, 3)) - Val(MSHFlexGrid4.TextMatrix(MSHFlexGrid4.Rows - 3, 3)), "#0")
MSHFlexGrid4.Rows = MSHFlexGrid4.Rows - 2
Call SecandUserSalaryExpense
'====================='for security and staff salary expense'=====================
End Sub

Private Sub TxtStartDate_GotFocus()
tmp = TxtStartDate.Value
End Sub

Private Sub TxtStartDate_LostFocus()
If mAgreementExpiryDate <> "" And LCase(CmbType.Text) = "renewal" Then
   If Format(CDate(mAgreementExpiryDate), "yyyy-mm-dd") >= Format(TxtStartDate.Value, "yyyy-mm-dd") Then
      'MsgBox " StartDate never less than or equal to Agreement Expiry Date!!!"
      MsgBox "Renewal Start date should be later than the godown's current Expiry Date!!!"
      TxtStartDate.SetFocus
      Exit Sub
   End If
End If
If TxtStartDate.Value > TxtCloseDate.Value Then
   MsgBox "Start Date should not greater than close date!!!"
   TxtStartDate.SetFocus
   Exit Sub
End If

If mCMPStatrDate <> "" Then
   If Format(CDate(mCMPStatrDate), "yyyy-mm-dd") > Format(TxtStartDate.Value, "yyyy-mm-dd") Then
      MsgBox "Godown Start date should be later than the CMP Start Date!!!"
      TxtStartDate.SetFocus
      Exit Sub
   End If
End If

If TxtStartDate.Value <> tmp Then
   If Val(MSHFlexGrid4.TextMatrix(0, 3)) > Val(MSHFlexGrid4.TextMatrix(MSHFlexGrid4.Rows - 2, 3)) + Val(MSHFlexGrid4.TextMatrix(MSHFlexGrid4.Rows - 3, 3)) Then
      MsgBox "Expense will be clear!!!"
      Call ClearAllExpenseDetails
   End If
End If
End Sub

Private Sub TxtWeightPerStack_LostFocus()
If TxtWeightPerStack.Locked = True Then Exit Sub
If Trim(TxtWeightPerStack.Text) = "" Then Exit Sub
If IsNumeric(TxtWeightPerStack.Text) = False Or Val(TxtWeightPerStack.Text) < 0 Then
   MsgBox "Invalid numeric format"
   TxtWeightPerStack.SetFocus
   Exit Sub
End If
'
TxtWeightPerStack.Text = Format(TxtWeightPerStack.Text, "#0.00")

If mGodownCategoryID = 2 And Val(mGodownTypeID) = 7 Then
      TxtCapacity.Text = Format(Val(TxtNoofStack) * Val(TxtWeightPerStack), "#0.00")
End If

End Sub

Private Sub TxtWidth_LostFocus()

mWidth = 0
If Trim(TxtWidth.Text) = "" Then
   Call CalArea
   Exit Sub
End If
If IsNumeric(TxtWidth.Text) = False Or Val(TxtWidth.Text) < 0 Then
   MsgBox "Invalid numeric format"
   TxtWidth.SetFocus
   Exit Sub
End If

TxtWidth = Format(TxtWidth, "########0.00")

Dim strSplitString() As String
Dim mWPart1, mWPart2 As Variant


strSplitString = Split(TxtWidth, ".")
If UBound(strSplitString) <> 1 Then
   MsgBox "Invalid Selection "
   TxtWidth.SetFocus
   Exit Sub
End If

If Val(strSplitString(1)) >= 12 Then
   MsgBox "Value after decimal should be less than 12"
   TxtWidth.SetFocus
   Exit Sub
End If

mWPart1 = Val(strSplitString(0))
mWPart2 = Format(Val(strSplitString(1)) / 120, "#######0.00")


mWidth = Val(mWPart1) + Val(mWPart2) 'Format(TxtLength.Text, "#0")


Call CalArea

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

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Call GButtonLockAD(Me)
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If

Call Grid_Head
Call Grid_Head2
Call Grid_Head3
Call Grid_Head4
Call Grid_Head5
Call Grid_Head6
Call Grid_Head7
Call Grid_Head9
Call Grid_Head10
Call Grid_Head11

Frame6.Enabled = False

SFrom.Value = Date - 30
STo.Value = Date
mBackMonths = 12
If Gen_WcLocation <> "" Then
   Call TableMst_GodownInspection(" Where " & Gen_WcLocation)
Else
   Call TableMst_GodownInspection
End If

If RsMst_GodownInspection.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
Call GButtonLock(Me, True)
RsMst_GodownInspection.MoveLast
Call Indata
End Sub

Private Sub SaveCmd_Click()

If Trim(CmbType.Text) = "" Then
   MsgBox "Blank Type not allowed!!!"
   CmbType.SetFocus
   Exit Sub
End If
If Trim(CmbLocation.Text) = "" Then
   MsgBox "Blank Location not allowed!!!"
   CmbLocation.SetFocus
   Exit Sub
End If
If Trim(CmbCMP.Text) = "" Then
   MsgBox "Blank CMP not allowed!!!"
   CmbCMP.SetFocus
   Exit Sub
End If
If LCase(Trim(CmbType.Text)) = "new" Then
   If Trim(TxtGodown.Text) = "" Then
      MsgBox "Blank Godown not allowed!!!"
      TxtGodown.SetFocus
      Exit Sub
   End If
   
   tmp = "Select GodownNo From Mst_GodownInspection Where InspectionID <> " & Val(TxtInspectionID.Text)
   tmp = tmp & " And CMPID = " & mCMPID & " And GodownNo = '" & Trim(TxtGodown.Text) & "' And Status not in ('R') " 'Status = 'P'"
   Call TmpTable
   If TmpRs.RecordCount > 0 Then
      MsgBox "Duplicate Godown No not allowed!!!"
      TxtGodown.SetFocus
      Exit Sub
   End If
   tmp = "Select GodownID From Mst_Godown Where CMPID = " & mCMPID & " And GodownNo = '" & Trim(TxtGodown.Text) & "'"
   Call TmpTable
   If TmpRs.RecordCount > 0 Then
      MsgBox "Duplicate Godown No not allowed!!!"
      TxtGodown.SetFocus
      Exit Sub
   End If
End If
If LCase(Trim(CmbType.Text)) = "renewal" Then
   If Trim(CmbGodown.Text) = "" Then
      MsgBox "Blank Godown not allowed!!!"
      CmbGodown.SetFocus
      Exit Sub
   End If
   tmp = " Select GodownID from Mst_GodownInspection Where InspectionID <> " & Val(TxtInspectionID.Text)
   tmp = tmp & " And GodownID = " & mGodownID & " And Status not in ('R') And (('" & TxtStartDate & "' >= StartDate And '" & TxtStartDate & "' <= CloseDate) Or '" & TxtCloseDate & "' >= StartDate And '" & TxtCloseDate & "' <= CloseDate Or ('" & TxtStartDate & "' <= StartDate And '" & TxtCloseDate & "' >= CloseDate))"
   tmp = tmp & " Union "
   tmp = tmp & " Select GodownID from Mst_GodownInspection Where InspectionID <> " & Val(TxtInspectionID.Text)
   tmp = tmp & " And CMPID = " & mCMPID & " And  GodownNo = '" & Trim(CmbGodown.Text) & "' And Status not in ('R') And (('" & TxtStartDate & "' >= StartDate And '" & TxtStartDate & "' <= CloseDate) Or '" & TxtCloseDate & "' >= StartDate And '" & TxtCloseDate & "' <= CloseDate Or ('" & TxtStartDate & "' <= StartDate And '" & TxtCloseDate & "' >= CloseDate))"
   
   Call TmpTable
   If TmpRs.RecordCount > 0 Then
      'MsgBox "Godown is already pending, Duplicate Godown not allowed!!!"
      MsgBox "Vaibility for this Godown for above selected period already exist, Duplicate Godown not allowed!!!"
      CmbGodown.SetFocus
      Exit Sub
   End If
End If

If Edit_Flg = "A" Then
   If mGodownCategoryID = 2 Then
      If Val(mGodownTypeID) = 1 Or Val(mGodownTypeID) = 7 Then
         If MSHFlexGrid9.Rows < 3 Then
            MsgBox "Area Details not updated!!!"
            TxtLength.SetFocus
            SSTab1.Tab = 5
            Exit Sub
         End If
      End If
      If LCase(CmbType.Text) = "new" Then
         If mGodownCategoryID = 2 And Val(mGodownTypeID) = 7 Then
            If Val(TxtNoofStack) = 0 Then
               MsgBox "Blank No Of Stack. not allowed!!!"
               TxtNoofStack.SetFocus
               Exit Sub
            End If
            If Val(TxtWeightPerStack) = 0 Then
               MsgBox "Blank Weight per Stack not allowed!!!"
               TxtWeightPerStack.SetFocus
               Exit Sub
            End If
         End If
      End If
   End If
End If


If Trim(TxtArea.Text) = "" Then
   MsgBox "Blank Area not allowed!!!"
   TxtArea.SetFocus
   Exit Sub
End If
'If Trim(TxtCapacity.Text) = "" Then
'   MsgBox "Blank Capacity not allowed!!!"
'   TxtCapacity.SetFocus
'   Exit Sub
'End If
If Trim(CmbLicenseNo.Text) = "" And Edit_Flg = "A" Then
   MsgBox "Blank License No not allowed!!!"
   CmbLicenseNo.SetFocus
   Exit Sub
End If
If Trim(CmbGodownType.Text) = "" Then
   MsgBox "Blank Godown Type not allowed!!!"
   CmbGodownType.SetFocus
   Exit Sub
End If
If Trim(TxtAddress.Text) = "" Then
   MsgBox "Blank Address not allowed!!!"
   TxtAddress.SetFocus
   Exit Sub
End If
If Trim(CmbGodownCategory.Text) = "" Then
   MsgBox "Blank Godown Category not allowed!!!"
   CmbGodownCategory.SetFocus
   Exit Sub
End If
If Trim(CmbStatus.Text) = "" Then
   MsgBox "Blank Status not allowed!!!"
   CmbStatus.SetFocus
   Exit Sub
End If
If TxtStartDate.Value > TxtCloseDate.Value Then
   MsgBox "Start Date should not greater than close date!!!"
   TxtCloseDate.SetFocus
   Exit Sub
End If
'If Trim(TxtStaffUser.Text) = "" Then
'   MsgBox "Blank 'No. of Staff' not allowed!!!"
'   TxtStaffUser.SetFocus
'   Exit Sub
'End If
'If Trim(TxtStaffSalaryUser.Text) = "" Then
'   MsgBox "Blank Staff Salary not allowed!!!"
'   TxtStaffSalaryUser.SetFocus
'   Exit Sub
'End If
'If Trim(TxtSecStaffUser.Text) = "" Then
'   MsgBox "Blank 'No. of Security Staff' not allowed!!!"
'   TxtSecStaffUser.SetFocus
'   Exit Sub
'End If
'If Trim(TxtSecStaffSalaryUser.Text) = "" Then
'   MsgBox "Blank Security Salary not allowed!!!"
'   TxtSecStaffSalaryUser.SetFocus
'   Exit Sub
'End If
If Trim(MSHFlexGrid6.TextMatrix(2, 3)) = "" Then
   MsgBox "Blank 'No. of Staff' not allowed!!!"
   SSTab1.Tab = 0
   Exit Sub
End If
If Trim(MSHFlexGrid6.TextMatrix(2, 4)) = "" Then
   MsgBox "Blank Staff Salary not allowed!!!"
   SSTab1.Tab = 0
   Exit Sub
End If
If Trim(MSHFlexGrid6.TextMatrix(3, 3)) = "" Then
   MsgBox "Blank 'No. of Security Staff' not allowed!!!"
   TxtSecStaffUser.SetFocus
   Exit Sub
End If
If Trim(MSHFlexGrid6.TextMatrix(3, 4)) = "" Then
   MsgBox "Blank Security Salary not allowed!!!"
   SSTab1.Tab = 0
   Exit Sub
End If
If Trim(TxtExpUtilisation.Text) = "" Then
   MsgBox "Blank Expected capacity utilisation (%) not allowed!!!"
   SSTab1.Tab = 0
   TxtExpUtilisation.SetFocus
   Exit Sub
End If
'If Trim(TxtNCDEXStock.Text) = "" Then
'   MsgBox "Blank NCDEX Stock (%) not allowed!!!"
'   TxtNCDEXStock.SetFocus
'   Exit Sub
'End If
'If Trim(TxtNNNRStock.Text) = "" Then
'   MsgBox "Blank NNNR Stock (%) not allowed!!!"
'   TxtNNNRStock.SetFocus
'   Exit Sub
'End If
'If Trim(TxtReservationStock.Text) = "" Then
'   MsgBox "Blank Reservation Stock (%) not allowed!!!"
'   TxtReservationStock.SetFocus
'   Exit Sub
'End If
If Val(TxtNCDEXStock.Text) + Val(TxtNNNRStock.Text) + Val(TxtReservationStock.Text) <> 100 Then
   MsgBox "Sum of Stock will be 100%!!!"
   SSTab1.Tab = 0
   TxtNCDEXStock.SetFocus
   Exit Sub
End If
If Trim(TxtRemarks.Text) = "" Then
   MsgBox "Blank Remarks not allowed!!!"
   SSTab1.Tab = 0
   TxtRemarks.SetFocus
   Exit Sub
End If
If MSHFlexGrid3.Rows = 3 Then
   MsgBox "Blank Commodity detail not allowed!!!"
   SSTab1.Tab = 1
   CmbCategory.SetFocus
   Exit Sub
End If

If Val(TxtNCDEXMT.Text) > 0 And mNCDEXPer <> 100 Then
   MsgBox "Expected Qty (in %) under 'NCDEX' is not 100%!!!"
   SSTab1.Tab = 1
   CmbCategory.SetFocus
   Exit Sub
End If
If Val(TxtNNNRMT.Text) > 0 And mNNNRPer <> 100 Then
   MsgBox "Expected Qty (in %) under 'NNNR' is not 100%!!!"
   SSTab1.Tab = 1
   CmbCategory.SetFocus
   Exit Sub
End If
If Val(TxtReservationMT.Text) > 0 And mReservationPer <> 100 Then
   MsgBox "Expected Qty (in %) under 'Reservation' is not 100%!!!"
   SSTab1.Tab = 1
   CmbCategory.SetFocus
   Exit Sub
End If
'If ChkExpenseDetail = False Then
'   SSTab1.Tab = 2
'   Exit Sub
'End If
If Val(MSHFlexGrid4.TextMatrix(0, 3)) = 0 Then
   MsgBox "Blank Expense detail not allowed!!!"
   SSTab1.Tab = 2
   Exit Sub
End If

'===================================for profit calculation===================================
TxtTotalRevenue.Text = Format(Val(MSHFlexGrid3.TextMatrix(0, 6)), "#0")
TxtDirectExpense.Text = 0
For I = 2 To MSHFlexGrid4.Rows - 2
   If LCase(MSHFlexGrid4.TextMatrix(I, 10)) = "d" Then
      TxtDirectExpense.Text = Format(Val(TxtDirectExpense.Text) + Val(MSHFlexGrid4.TextMatrix(I, 3)), "#0")
   End If
Next
TxtGrossProfit.Text = Format(Val(TxtTotalRevenue.Text) - Val(TxtDirectExpense.Text), "#0")
TxtIndirectExpense.Text = 0
For I = 2 To MSHFlexGrid4.Rows - 2
   If LCase(MSHFlexGrid4.TextMatrix(I, 10)) = "i" Then
      TxtIndirectExpense.Text = Format(Val(TxtIndirectExpense.Text) + Val(MSHFlexGrid4.TextMatrix(I, 3)), "#0")
   End If
Next
TxtNetProfit.Text = Format(Val(TxtGrossProfit.Text) - Val(TxtIndirectExpense.Text), "#0")
TxtCMShare.Text = Format(Val(MSHFlexGrid3.TextMatrix(0, 14)), "#0")
'===================================for profit calculation===================================

If Edit_Flg = "A" Then
   'If RsMst_GodownInspection.RecordCount > 0 Then
   '   RsMst_GodownInspection.MoveFirst
   '   RsMst_GodownInspection.Find "ExpenseType= '" & Trim(TxtExpenseType.Text) & "'"
   '   If Not RsMst_GodownInspection.EOF Then
   '      MsgBox "Duplicate Expense Type Not Allowed!!!"
   '      TxtExpenseType.SetFocus
   '      Exit Sub
   '   End If
   'End If
   RsMst_GodownInspection.AddNew
   RsMst_GodownInspection!InspectionID = GetInspectionID
   RsMst_GodownInspection!G_UID = GetGUID
   TxtInspectionID.Text = RsMst_GodownInspection!InspectionID
Else
   tmp = "Select UpdatedBy, UpdatedDate From Mst_GodownInspection Where InspectionID= " & Val(TxtInspectionID.Text)
   Call TmpTable
   If TmpRs.RecordCount > 0 Then
      If TxtUpdated.Text <> IIf(IsNull(TmpRs!UpdatedBy), "", TmpRs!UpdatedBy & " :- " & TmpRs!UpdatedDate) Then
         MsgBox "Data has been changed by Other User, You can't save!!!"
         mSaveFail = True
         Call Gen_Email
         Exit Sub
      End If
   End If
   
   RsMst_GodownInspection.MoveFirst
   RsMst_GodownInspection.Find " InspectionID= " & TxtInspectionID.Text
End If
RsMst_GodownInspection!Type = Left$(CmbType.Text, 1)
RsMst_GodownInspection!LocationID = mLocationID
RsMst_GodownInspection!CMPID = mCMPID
If LCase(CmbType.Text) = "new" Then
   RsMst_GodownInspection!GodownNo = TxtGodown.Text
ElseIf LCase(CmbType.Text) = "renewal" Then
   RsMst_GodownInspection!GodownID = mGodownID
End If
RsMst_GodownInspection!UnitTypeID = mGodownTypeID
RsMst_GodownInspection!GodownCategoryID = mGodownCategoryID
If CmbLicenseNo.Text <> "" Then
   RsMst_GodownInspection!LicenseID = mLicenseID
End If
RsMst_GodownInspection!Capacity = Val(TxtCapacity.Text)
RsMst_GodownInspection!Address = TxtAddress.Text

RsMst_GodownInspection!Length = Val(TxtLength.Text)
RsMst_GodownInspection!Width = Val(TxtWidth.Text)
RsMst_GodownInspection!Height = Val(TxtHeight.Text)

RsMst_GodownInspection!NoOfStack = Val(TxtNoofStack.Text)
RsMst_GodownInspection!WeightPerStack = Val(TxtWeightPerStack.Text)

RsMst_GodownInspection!Area = Val(TxtArea.Text)
RsMst_GodownInspection!Status = Left$(CmbStatus.Text, 1)
RsMst_GodownInspection!StartDate = TxtStartDate.Value
RsMst_GodownInspection!CloseDate = TxtCloseDate.Value
'RsMst_GodownInspection!noofstaffuser = Val(TxtStaffUser.Text)
'RsMst_GodownInspection!staffsalaryuser = Val(TxtStaffSalaryUser.Text)
'RsMst_GodownInspection!noofsecuser = Val(TxtSecStaffUser.Text)
'RsMst_GodownInspection!secsalaryuser = Val(TxtSecStaffSalaryUser.Text)
'RsMst_GodownInspection!noofstaff = Val(TxtStaffsys.Text)
'RsMst_GodownInspection!staffsalary = Val(TxtStaffSalarySys.Text)
'RsMst_GodownInspection!noofsec = Val(TxtSecStaffSys.Text)
'RsMst_GodownInspection!secsalary = Val(TxtSecStaffSalarySys.Text)

RsMst_GodownInspection!noofstaffuser = Val(MSHFlexGrid6.TextMatrix(2, 3))
RsMst_GodownInspection!staffsalaryuser = Val(MSHFlexGrid6.TextMatrix(2, 4))
RsMst_GodownInspection!noofsecuser = Val(MSHFlexGrid6.TextMatrix(3, 3))
RsMst_GodownInspection!secsalaryuser = Val(MSHFlexGrid6.TextMatrix(3, 4))
RsMst_GodownInspection!noofstaff = Val(MSHFlexGrid6.TextMatrix(2, 1))
RsMst_GodownInspection!staffsalary = Val(MSHFlexGrid6.TextMatrix(2, 2))
RsMst_GodownInspection!noofsec = Val(MSHFlexGrid6.TextMatrix(3, 1))
RsMst_GodownInspection!secsalary = Val(MSHFlexGrid6.TextMatrix(3, 2))
RsMst_GodownInspection!Remarks = TxtRemarks.Text
RsMst_GodownInspection!totalrevenue = Val(TxtTotalRevenue.Text)
RsMst_GodownInspection!DirectExpense = Val(TxtDirectExpense.Text)
RsMst_GodownInspection!GrossProfit = Val(TxtGrossProfit.Text)
RsMst_GodownInspection!IndirectExpense = Val(TxtIndirectExpense.Text)
RsMst_GodownInspection!NetProfit = Val(TxtNetProfit.Text)
RsMst_GodownInspection!CapacityUtilisation = Val(TxtExpUtilisation.Text)
RsMst_GodownInspection!NcdexStock = Val(TxtNCDEXStock.Text)
RsMst_GodownInspection!NNNRStock = Val(TxtNNNRStock.Text)
RsMst_GodownInspection!ReservationStock = Val(TxtReservationStock.Text)
If IsNull(RsMst_GodownInspection!CreatedBy) = True Or RsMst_GodownInspection!CreatedBy = "" Then
   RsMst_GodownInspection!CreatedBy = Gen_UserLoginID
   RsMst_GodownInspection!CreatedDate = Now
Else
   RsMst_GodownInspection!UpdatedBy = Gen_UserLoginID
   RsMst_GodownInspection!UpdatedDate = Now
End If
RsMst_GodownInspection.Update

If Edit_Flg = "A" Then
   Call SaveInspectionGodownDetail
End If
If mGodownCategoryID = 2 Then
   If Val(mGodownTypeID) = 1 Or Val(mGodownTypeID) = 7 Then
      Call SaveInspectionAreaDetail
   End If
End If
Call SaveInspectionCommodityDetail
Call SaveInspectionExpenseDetail
'If LCase(CmbStatus.Text) = "approved" Then
'   Call UpdateGodownDetail
'End If
If LCase(CmbStatus.Text) = "godown creation" Then
   Call UpdateGodownDetail
End If

RsMst_GodownInspection.Requery
RsMst_GodownInspection.MoveFirst
RsMst_GodownInspection.Find " InspectionID= " & Val(TxtInspectionID.Text)

Call Indata
mSaveFail = False
Call Gen_Email

If LCase(CmbType.Text) = "new" And LCase(CmbStatus.Text) = "godown creation" Then
   Call GetGodownList
End If
End Sub

Private Sub UpdateGodownDetail()
Call TableMst_Godown("Where CMPID = " & mCMPID)
If LCase(CmbType.Text) = "new" Then
   RsMst_Godown.AddNew
   RsMst_Godown!GodownID = GetMaxGodownID
   RsMst_Godown!Concurrency = Now
   RsMst_Godown!CreatedBy = Gen_UserLoginID
   RsMst_Godown!CreatedDate = Now
   RsMst_Godown!G_UID = GetGUID
   RsMst_Godown!CMPID = mCMPID
   RsMst_Godown!UnitTypeID = mGodownTypeID
   RsMst_Godown!CategoryID = mGodownCategoryID
   RsMst_Godown!GodownNo = TxtGodown.Text
   RsMst_Godown!Address = TxtAddress.Text
   RsMst_Godown!NoOfStack = Val(TxtNoofStack.Text)
   RsMst_Godown!WeightPerStack = Val(TxtWeightPerStack.Text)
   RsMst_Godown!Area = TxtArea.Text
   RsMst_Godown!BagsCapacity = Val(TxtCapacity.Text) * 10 'ask mayur sir
   RsMst_Godown!GodownCapacity = Val(TxtCapacity.Text)
   RsMst_Godown!StartDate = TxtStartDate.Value

   If CmbLicenseNo.Text <> "" Then
      RsMst_Godown!LicenseID = mLicenseID
   End If
   RsMst_Godown!TelephoneNo = 0 'ask mayur sir
   'RsMst_Godown!Amount = Null
   'RsMst_Godown!AdvAmount = Null
   'RsMst_Godown!NCDEXAccreditedQty = TxtCapacity.Text ' ask mayur sir
   'If MSHFlexGrid2.TextMatrix(I, 22) <> "" Then
   '   RsMst_Godown!AccreditedID = Val(MSHFlexGrid2.TextMatrix(I, 22))
   'Else
   '   RsMst_Godown!AccreditedID = Null
   'End If
   'RsMst_Godown!CPEN = MSHFlexGrid2.TextMatrix(I, 24)
   'RsMst_Godown!CPED = MSHFlexGrid2.TextMatrix(I, 25)
   'RsMst_Godown!COPEN = MSHFlexGrid2.TextMatrix(I, 26)
   'RsMst_Godown!COPED = MSHFlexGrid2.TextMatrix(I, 27)
Else
'   tmp = " select TLA.LAID, TLA.ExpiryDate from Txn_GodownLessorAgreeDetail TGLA "
'   tmp = tmp & " Inner Join Txn_LessorAgreement TLA on TLA.LAID = TGLA.LAID"
'   tmp = tmp & " Where TLA.Flag in ('A', 'P') And TGLA.GodownID = " & mGodownID
'   tmp = tmp & " Order by TLA.ExpiryDate Desc"
'   Call Tmp7Table
'   If TmpRs7.RecordCount > 0 Then
'      tmp = "Update Txn_LessorAgreement set ExpiryDate = '" & Format(TxtCloseDate.Value, "yyyy-mm-dd") & "' Where LAID = " & IIf(IsNull(TmpRs7!LAID), 0, TmpRs7!LAID)
'      Call Tmp6Table
'   End If

   RsMst_Godown.MoveFirst
   RsMst_Godown.Find "GodownID = " & mGodownID
   RsMst_Godown!UpdatedBy = Gen_UserLoginID
   RsMst_Godown!UpdatedDate = Now
End If

If LCase(CmbGodownType.Text) = "leased franchisee" Then
   RsMst_Godown!LeasedUptoDate = TxtCloseDate.Value
Else
   RsMst_Godown!LeasedUptoDate = Null
End If

'"IF" COndition introduced on 25 July 2013 after mail from Mayur Sir and Sankar SIr to show Agreement date from Lessor Agreement Master

If mGodownTypeID <> 1 Then
   RsMst_Godown!AgreementEndDate = TxtCloseDate.Value
End If

RsMst_Godown.Update
mmGodownID = RsMst_Godown!GodownID
End Sub

Private Sub SaveInspectionGodownDetail()
If MSHFlexGrid2.Rows > 3 Then
   Call TableMst_InspectionGodownDetail("Where InspectionID = " & Val(TxtInspectionID.Text))
   For I = 2 To MSHFlexGrid2.Rows - 2
      RsMst_InspectionGodownDetail.AddNew
      RsMst_InspectionGodownDetail!InspectionGodownID = GetInspectionGodownID
      RsMst_InspectionGodownDetail!CreatedBy = Gen_UserLoginID
      RsMst_InspectionGodownDetail!CreatedDate = Now
      RsMst_InspectionGodownDetail!G_UID = GetGUID
      RsMst_InspectionGodownDetail!InspectionID = Val(TxtInspectionID.Text)
      RsMst_InspectionGodownDetail!GodownID = Val(MSHFlexGrid2.TextMatrix(I, 8))
      RsMst_InspectionGodownDetail!Capacity = Val(MSHFlexGrid2.TextMatrix(I, 3))
      RsMst_InspectionGodownDetail!Rent = Val(MSHFlexGrid2.TextMatrix(I, 6))
      RsMst_InspectionGodownDetail!Stock = Val(MSHFlexGrid2.TextMatrix(I, 4))
      RsMst_InspectionGodownDetail!Area = Val(MSHFlexGrid2.TextMatrix(I, 5))
      RsMst_InspectionGodownDetail.Update
   Next
End If
End Sub

Private Sub SaveInspectionCommodityDetail()
Dim mInspectionCommodityIDs As String
mInspectionCommodityIDs = "(0 "
For I = 2 To MSHFlexGrid3.Rows - 2
   If Val(MSHFlexGrid3.TextMatrix(I, 9)) > 0 Then
      mInspectionCommodityIDs = mInspectionCommodityIDs & " ," & Val(MSHFlexGrid3.TextMatrix(I, 9))
   End If
Next
mInspectionCommodityIDs = mInspectionCommodityIDs & ")"
tmp = " Delete from Mst_InspectionCommodityDetail Where InspectionID = " & Val(TxtInspectionID.Text) & " and InspectionCommodityID not in " & mInspectionCommodityIDs
Call Tmp6Table

If MSHFlexGrid3.Rows > 3 Then
   Call TableMst_InspectionCommodityDetail("Where InspectionID = " & Val(TxtInspectionID.Text))
   For I = 2 To MSHFlexGrid3.Rows - 2
      If MSHFlexGrid3.TextMatrix(I, 9) = "" Then
         RsMst_InspectionCommodityDetail.AddNew
         RsMst_InspectionCommodityDetail!InspectionCommodityID = GetInspectionCommodityID
         RsMst_InspectionCommodityDetail!CreatedBy = Gen_UserLoginID
         RsMst_InspectionCommodityDetail!CreatedDate = Now
         RsMst_InspectionCommodityDetail!G_UID = GetGUID
      Else
         RsMst_InspectionCommodityDetail.MoveFirst
         RsMst_InspectionCommodityDetail.Find "InspectionCommodityID = " & Val(MSHFlexGrid3.TextMatrix(I, 9))
         RsMst_InspectionCommodityDetail!UpdatedBy = Gen_UserLoginID
         RsMst_InspectionCommodityDetail!UpdatedDate = Now
      End If
      RsMst_InspectionCommodityDetail!InspectionID = Val(TxtInspectionID.Text)
      RsMst_InspectionCommodityDetail!Category = Left$(MSHFlexGrid3.TextMatrix(I, 0), 2)
      RsMst_InspectionCommodityDetail!CommodityID = Val(MSHFlexGrid3.TextMatrix(I, 10))
      RsMst_InspectionCommodityDetail!ExpQtyP = Val(MSHFlexGrid3.TextMatrix(I, 2))
      RsMst_InspectionCommodityDetail!ExpQtyMT = Val(MSHFlexGrid3.TextMatrix(I, 3))
      RsMst_InspectionCommodityDetail!StorageRate = Val(MSHFlexGrid3.TextMatrix(I, 4))
      RsMst_InspectionCommodityDetail!Price = Val(MSHFlexGrid3.TextMatrix(I, 5))
      RsMst_InspectionCommodityDetail!Revenue = Val(MSHFlexGrid3.TextMatrix(I, 6))
      RsMst_InspectionCommodityDetail!AUM = Val(MSHFlexGrid3.TextMatrix(I, 7))
      RsMst_InspectionCommodityDetail!StorageRateRemarks = MSHFlexGrid3.TextMatrix(I, 8)
      RsMst_InspectionCommodityDetail!FumigableQty = Val(MSHFlexGrid3.TextMatrix(I, 11))
      RsMst_InspectionCommodityDetail!CMShareP = Val(MSHFlexGrid3.TextMatrix(I, 13))
      RsMst_InspectionCommodityDetail!CMShareR = Val(MSHFlexGrid3.TextMatrix(I, 14))
      
      RsMst_InspectionCommodityDetail.Update
   Next
End If
End Sub
Private Sub SaveInspectionAreaDetail()

Dim mInspectionAreaIDs As String
mInspectionAreaIDs = "(0 "
For I = 2 To MSHFlexGrid9.Rows - 2
   If Val(MSHFlexGrid9.TextMatrix(I, 0)) > 0 Then
      mInspectionAreaIDs = mInspectionAreaIDs & " ," & Val(MSHFlexGrid9.TextMatrix(I, 0))
   End If
Next
mInspectionAreaIDs = mInspectionAreaIDs & ")"
tmp = " Delete from Mst_InspectionAreaDetail Where InspectionID = " & Val(TxtInspectionID.Text) & " and InspectionAreaID not in " & mInspectionAreaIDs
Call Tmp6Table

If MSHFlexGrid9.Rows > 3 Then
   'Call TableMst_InspectionCommodityDetail("Where InspectionID = " & Val(TxtInspectionID.Text))
   Call TableMst_InspectionAreaDetail("Where InspectionID = " & Val(TxtInspectionID.Text))
   For I = 2 To MSHFlexGrid9.Rows - 2
      If MSHFlexGrid9.TextMatrix(I, 0) = "" Then
         RsMst_InspectionAreaDetail.AddNew
         RsMst_InspectionAreaDetail!InspectionAreaID = GetInspectionAreaID 'GetInspectionCommodityID
         RsMst_InspectionAreaDetail!CreatedBy = Gen_UserLoginID
         RsMst_InspectionAreaDetail!CreatedDate = Now
         RsMst_InspectionAreaDetail!G_UID = GetGUID
      Else
         RsMst_InspectionAreaDetail.MoveFirst
         RsMst_InspectionAreaDetail.Find "InspectionAreaID = " & Val(MSHFlexGrid9.TextMatrix(I, 0))
         RsMst_InspectionAreaDetail!UpdatedBy = Gen_UserLoginID
         RsMst_InspectionAreaDetail!UpdatedDate = Now
      End If
      RsMst_InspectionAreaDetail!InspectionID = Val(TxtInspectionID.Text)
      RsMst_InspectionAreaDetail!Length = Val(MSHFlexGrid9.TextMatrix(I, 1))
      RsMst_InspectionAreaDetail!Width = Val(MSHFlexGrid9.TextMatrix(I, 2))
      RsMst_InspectionAreaDetail!Height = Val(MSHFlexGrid9.TextMatrix(I, 3))
      RsMst_InspectionAreaDetail!Area = Val(MSHFlexGrid9.TextMatrix(I, 4))
      RsMst_InspectionAreaDetail.Update
   Next
End If


End Sub

Private Sub SaveInspectionExpenseDetail()
'Dim mInspectionExpenseIDs As String
'mInspectionExpenseIDs = "(0 "
'For I = 2 To MSHFlexGrid4.Rows - 2
'   If Val(MSHFlexGrid4.TextMatrix(I, 5)) > 0 Then
'      mInspectionExpenseIDs = mInspectionExpenseIDs & " ," & Val(MSHFlexGrid4.TextMatrix(I, 5))
'   End If
'Next
'mInspectionExpenseIDs = mInspectionExpenseIDs & ")"
'tmp = " Delete from Mst_InspectionExpenseDetail Where InspectionID = " & Val(TxtInspectionID.Text) & " and InspectionExpenseID not in " & mInspectionExpenseIDs
'Call Tmp6Table

If MSHFlexGrid3.Rows > 3 Then
   Call TableMst_InspectionExpenseDetail("Where InspectionID = " & Val(TxtInspectionID.Text))
   For I = 2 To MSHFlexGrid4.Rows - 4
      If MSHFlexGrid4.TextMatrix(I, 5) = "" Then
         RsMst_InspectionExpenseDetail.AddNew
         RsMst_InspectionExpenseDetail!InspectionExpenseID = GetInspectionExpenseID
         RsMst_InspectionExpenseDetail!CreatedBy = Gen_UserLoginID
         RsMst_InspectionExpenseDetail!CreatedDate = Now
         RsMst_InspectionExpenseDetail!G_UID = GetGUID
      Else
         RsMst_InspectionExpenseDetail.MoveFirst
         RsMst_InspectionExpenseDetail.Find "InspectionExpenseID = " & Val(MSHFlexGrid4.TextMatrix(I, 5))
         RsMst_InspectionExpenseDetail!UpdatedBy = Gen_UserLoginID
         RsMst_InspectionExpenseDetail!UpdatedDate = Now
      End If
      RsMst_InspectionExpenseDetail!InspectionID = Val(TxtInspectionID.Text)
      RsMst_InspectionExpenseDetail!InsExpTypeID = Val(MSHFlexGrid4.TextMatrix(I, 6))
      RsMst_InspectionExpenseDetail!QtySQ = Val(MSHFlexGrid4.TextMatrix(I, 2))
      RsMst_InspectionExpenseDetail!Amount = Val(MSHFlexGrid4.TextMatrix(I, 3))
      RsMst_InspectionExpenseDetail!AmountPer = Val(MSHFlexGrid4.TextMatrix(I, 7))
      RsMst_InspectionExpenseDetail!Remarks = MSHFlexGrid4.TextMatrix(I, 4)
      RsMst_InspectionExpenseDetail!Formula = MSHFlexGrid4.TextMatrix(I, 8)
      RsMst_InspectionExpenseDetail.Update
   Next
End If
End Sub


Private Function GetInspectionID() As Double
Dim Retval As Double
tmp = "Select max(InspectionID) from Mst_GodownInspection"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetInspectionID = Retval
End Function

Private Function GetInspectionGodownID() As Double
Dim Retval As Double
tmp = "Select max(InspectionGodownID) from Mst_InspectionGodownDetail"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetInspectionGodownID = Retval
End Function

Private Function GetInspectionCommodityID() As Double
Dim Retval As Double
tmp = "Select max(InspectionCommodityID) from Mst_InspectionCommodityDetail"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetInspectionCommodityID = Retval
End Function
Private Function GetInspectionAreaID() As Double
Dim Retval As Double
tmp = "Select max(InspectionAreaID) from Mst_InspectionAreaDetail"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetInspectionAreaID = Retval
End Function

Private Function GetInspectionExpenseID() As Double
Dim Retval As Double
tmp = "Select max(InspectionExpenseID) from Mst_InspectionExpenseDetail"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetInspectionExpenseID = Retval
End Function

Private Sub TxtAmount_LostFocus()
If Trim(TxtAmount.Text) = "" Then Exit Sub
If IsNumeric(TxtAmount.Text) = False Or Val(TxtAmount.Text) < 0 Then
   MsgBox "Invalid numeric format"
   TxtAmount.SetFocus
   Exit Sub
End If

TxtAmount.Text = Format(TxtAmount.Text, "#0")
End Sub

Private Sub TxtArea_LostFocus()

If Trim(TxtArea.Text) = "" Then Exit Sub
If IsNumeric(TxtArea.Text) = False Or Val(TxtArea.Text) < 0 Then
   MsgBox "Invalid numeric format"
   TxtArea.SetFocus
   Exit Sub
End If
'
TxtArea.Text = Format(TxtArea.Text, "#0.00")
'if tmp <> TxtArea.Text Then
   Call TxtCapacity_GotFocus
   If mGodownCategoryID = 2 And Val(mGodownTypeID) = 7 Then
      TxtCapacity.Text = Format(Val(TxtNoofStack) * Val(TxtWeightPerStack), "#0.00")
   Else
      TxtCapacity.Text = Format(Val(TxtArea.Text) / 6, "#0.00")
   End If
   Call TxtCapacity_LostFocus
'End If

End Sub

Private Sub TxtCapacity_LostFocus()
If Trim(TxtCapacity.Text) = "" Then
   If Val(MSHFlexGrid3.TextMatrix(0, 7)) > 0 Then
      If mGodownCategoryID = 2 Then
         If Val(mGodownTypeID) <> 1 And Val(mGodownTypeID) <> 7 Then
            If MsgBox("Commodity and expense detail will clear, do you want to continue?", vbYesNo) = vbNo Then
               TxtCapacity.Text = tmp
'         TxtCapacity.SetFocus
               Exit Sub
            End If
            Call Grid_Head3
            Call ClearAllExpenseDetails
         Else
            MsgBox "Commodity and expense detail will get clear!!!"
            Call Grid_Head3
            Call ClearAllExpenseDetails
         End If
      Else
         Call Grid_Head3
         Call ClearAllExpenseDetails
      End If
   End If
   
   Call CapacityCalculation
   ''====================='for security and staff salary expense'=====================
   MSHFlexGrid4.TextMatrix(0, 3) = Format(Val(MSHFlexGrid4.TextMatrix(0, 3)) - Val(MSHFlexGrid4.TextMatrix(MSHFlexGrid4.Rows - 2, 3)), "#0")
   MSHFlexGrid4.TextMatrix(0, 3) = Format(Val(MSHFlexGrid4.TextMatrix(0, 3)) - Val(MSHFlexGrid4.TextMatrix(MSHFlexGrid4.Rows - 3, 3)), "#0")
   MSHFlexGrid4.Rows = MSHFlexGrid4.Rows - 2
   Call SecandUserSalaryExpense
   '====================='for security and staff salary expense'=====================
   Exit Sub
End If
If IsNumeric(TxtCapacity.Text) = False Or Val(TxtCapacity.Text) < 0 Then
   MsgBox "Invalid numeric format"
   TxtCapacity.SetFocus
   Exit Sub
End If
TxtCapacity.Text = Format(TxtCapacity.Text, "#0")
If tmp <> TxtCapacity.Text Then
   If Val(MSHFlexGrid3.TextMatrix(0, 7)) > 0 Then
      If mGodownCategoryID = 2 Then
         If Val(mGodownTypeID) = 1 Or Val(mGodownTypeID) = 7 Then
            MsgBox "Commodity and expense detail will get clear!!!"
            Call Grid_Head3
            Call ClearAllExpenseDetails
         Else
            If MsgBox("Commodity and expense detail will clear, do you want to continue?", vbYesNo) = vbNo Then
               Exit Sub
            Else
               TxtCapacity.Text = tmp
               TxtCapacity.SetFocus
               Exit Sub
            End If
         End If
      Else
         TxtCapacity.Text = tmp
         TxtCapacity.SetFocus
         Exit Sub
      End If
      Call Grid_Head3
      Call ClearAllExpenseDetails
   End If
   Call CapacityCalculation
   ''====================='for security and staff salary expense'=====================
   MSHFlexGrid4.TextMatrix(0, 3) = Format(Val(MSHFlexGrid4.TextMatrix(0, 3)) - Val(MSHFlexGrid4.TextMatrix(MSHFlexGrid4.Rows - 2, 3)), "#0")
   MSHFlexGrid4.TextMatrix(0, 3) = Format(Val(MSHFlexGrid4.TextMatrix(0, 3)) - Val(MSHFlexGrid4.TextMatrix(MSHFlexGrid4.Rows - 3, 3)), "#0")
   MSHFlexGrid4.Rows = MSHFlexGrid4.Rows - 2
   Call SecandUserSalaryExpense
   '====================='for security and staff salary expense'=====================
End If
End Sub

Private Sub TxtExpUtilisation_LostFocus()
If Trim(TxtExpUtilisation.Text) = "" Then
   If Val(MSHFlexGrid3.TextMatrix(0, 7)) > 0 Then
      If MsgBox("Commodity and expense detail will clear, do you want to continue?", vbYesNo) = vbNo Then
         TxtExpUtilisation.Text = tmp
         TxtExpUtilisation.SetFocus
         Exit Sub
      End If
      Call Grid_Head3
      Call ClearAllExpenseDetails
   End If
   Call CapacityCalculation
   Exit Sub
End If
If IsNumeric(TxtExpUtilisation.Text) = False Or Val(TxtExpUtilisation.Text) < 0 Then
   MsgBox "Invalid numeric format!!!"
   TxtExpUtilisation.SetFocus
   Exit Sub
End If
'If Val(TxtExpUtilisation.Text) > 100 Then
'   MsgBox "Expected Capacity Utilisation (%) never greater than 100!!!"
'   TxtExpUtilisation.SetFocus
'   Exit Sub
'End If

If Val(TxtArea) <= 0 Then
   MsgBox "Area details not updated!!!"
   TxtExpUtilisation.Text = ""
   SSTab1.Tab = 5
   If Frame6.Enabled = False Then TxtLength.SetFocus
   Exit Sub
End If

TxtExpUtilisation.Text = Format(TxtExpUtilisation.Text, "#0")
If tmp <> TxtExpUtilisation.Text Then
   If Val(MSHFlexGrid3.TextMatrix(0, 7)) > 0 Then
      If MsgBox("Commodity and expense detail will clear, do you want to continue?", vbYesNo) = vbNo Then
         TxtExpUtilisation.Text = tmp
         TxtExpUtilisation.SetFocus
         Exit Sub
      End If
      Call Grid_Head3
      Call ClearAllExpenseDetails
   End If
   Call CapacityCalculation
End If

End Sub

Private Sub CapacityCalculation()
TxtExpUtilisationMT.Text = Format(Val(TxtExpUtilisation.Text) * Val(TxtCapacity.Text) / 100, "#0")
TxtNCDEXMT.Text = Format(Val(TxtNCDEXStock.Text) * Val(TxtExpUtilisationMT.Text) / 100, "#0")
TxtNNNRMT.Text = Format(Val(TxtNNNRStock.Text) * Val(TxtExpUtilisationMT.Text) / 100, "#0")
TxtReservationMT.Text = Format(Val(TxtReservationStock.Text) * Val(TxtExpUtilisationMT.Text) / 100, "#0")
''====================='for security and staff salary expense'=====================
'MSHFlexGrid4.TextMatrix(0, 3) = Format(Val(MSHFlexGrid4.TextMatrix(0, 3)) - Val(MSHFlexGrid4.TextMatrix(MSHFlexGrid4.Rows - 2, 3)), "#0")
'MSHFlexGrid4.TextMatrix(0, 3) = Format(Val(MSHFlexGrid4.TextMatrix(0, 3)) - Val(MSHFlexGrid4.TextMatrix(MSHFlexGrid4.Rows - 3, 3)), "#0")
'MSHFlexGrid4.Rows = MSHFlexGrid4.Rows - 2
'Call SecandUserSalaryExpense
''====================='for security and staff salary expense'=====================
End Sub

Private Sub TxtNCDEXStock_LostFocus()
If TxtNCDEXStock.Text <> tmp Then
   If Val(MSHFlexGrid3.TextMatrix(0, 7)) > 0 Then
      If MsgBox("Commodity and expense detail will clear, do you want to continue?", vbYesNo) = vbNo Then
         TxtNCDEXStock.Text = tmp
         TxtNCDEXStock.SetFocus
         Exit Sub
      End If
      Call Grid_Head3
      Call ClearAllExpenseDetails
   End If
End If
If Trim(TxtNCDEXStock.Text) = "" Then
   TxtNCDEXMT.Text = ""
   Exit Sub
End If

If IsNumeric(TxtNCDEXStock.Text) = False Or Val(TxtNCDEXStock.Text) < 0 Then
   MsgBox "Invalid numeric format"
   TxtNCDEXStock.SetFocus
   Exit Sub
End If
If Val(TxtNCDEXStock.Text) > 100 Then
   MsgBox "NCDEX Stock (%) never greater than 100!!!"
   TxtNCDEXStock.SetFocus
   Exit Sub
End If

TxtNCDEXStock.Text = Format(TxtNCDEXStock.Text, "#0")
TxtNCDEXMT.Text = Format(Val(TxtNCDEXStock.Text) * Val(TxtExpUtilisationMT.Text) / 100, "#0")
End Sub

Private Sub TxtNNNRStock_LostFocus()
If Trim(TxtNNNRStock.Text) = "" Then
   If Val(MSHFlexGrid3.TextMatrix(0, 7)) > 0 Then
      If MsgBox("Commodity and expense detail will clear, do you want to continue?", vbYesNo) = vbNo Then
         TxtNNNRStock.Text = tmp
         TxtNNNRStock.SetFocus
         Exit Sub
      End If
      Call Grid_Head3
      Call ClearAllExpenseDetails
   End If
   TxtNNNRMT.Text = ""
   Exit Sub
End If
If IsNumeric(TxtNNNRStock.Text) = False Or Val(TxtNNNRStock.Text) < 0 Then
   MsgBox "Invalid numeric format"
   TxtNNNRStock.SetFocus
   Exit Sub
End If
If Val(TxtNNNRStock.Text) > 100 Then
   MsgBox "NNNR Stock (%) never greater than 100!!!"
   TxtNNNRStock.SetFocus
   Exit Sub
End If
TxtNNNRStock.Text = Format(TxtNNNRStock.Text, "#0")
If tmp <> TxtNNNRStock.Text Then
   If Val(MSHFlexGrid3.TextMatrix(0, 7)) > 0 Then
      If MsgBox("Commodity and expense detail will clear, do you want to continue?", vbYesNo) = vbNo Then
         TxtNNNRStock.Text = tmp
         TxtNNNRStock.SetFocus
         Exit Sub
      End If
      Call Grid_Head3
      Call ClearAllExpenseDetails
   End If
End If
TxtNNNRMT.Text = Format(Val(TxtNNNRStock.Text) * Val(TxtExpUtilisationMT.Text) / 100, "#0")
End Sub

Private Sub TxtExpQtyP_LostFocus()
If Trim(TxtExpQtyP.Text) = "" Then Exit Sub
If IsNumeric(TxtExpQtyP.Text) = False Or Val(TxtExpQtyP.Text) < 0 Then
   MsgBox "Invalid numeric format"
   TxtExpQtyP.SetFocus
   Exit Sub
End If
If Val(TxtExpQtyP.Text) > 100 Then
   MsgBox "Expected Qty (in %) never greater than 100!!!"
   TxtExpQtyP.SetFocus
   Exit Sub
End If
TxtExpQtyP.Text = Format(TxtExpQtyP.Text, "#0")

If LCase(CmbCategory.Text) = "ncdex" Then
   TxtExpQtyMT.Text = Format(Val(TxtNCDEXMT.Text) * Val(TxtExpQtyP.Text) / 100, "#0")
ElseIf LCase(CmbCategory.Text) = "nnnr" Then
   TxtExpQtyMT.Text = Format(Val(TxtNNNRMT.Text) * Val(TxtExpQtyP.Text) / 100, "#0")
ElseIf LCase(CmbCategory.Text) = "reservation" Then
   TxtExpQtyMT.Text = Format(Val(TxtReservationMT.Text) * Val(TxtExpQtyP.Text) / 100, "#0")
End If
TxtAUM.Text = Format(Val(TxtExpQtyMT.Text) * Val(TxtPrice.Text), "#0")
TxtCMShareR.Text = Format(Val(TxtAUM.Text) * Val(TxtCMShareP.Text) * 0.00007 * 60 / 1200, "#0")

TxtRevenue.Text = Format(Val(TxtExpQtyMT.Text) * Val(TxtStorageRate.Text), "#0")
If LCase(mFumigationFlag) = "y" Then
   TxtFumigableQty.Text = TxtExpQtyMT.Text
End If

End Sub

Private Sub TxtPrice_LostFocus()
If Trim(TxtPrice.Text) = "" Then Exit Sub
If IsNumeric(TxtPrice.Text) = False Or Val(TxtPrice.Text) < 0 Then
   MsgBox "Invalid numeric format"
   TxtPrice.SetFocus
   Exit Sub
End If

TxtPrice.Text = Format(TxtPrice.Text, "#0")
TxtAUM.Text = Format(Val(TxtExpQtyMT.Text) * Val(TxtPrice.Text), "#0")
TxtCMShareR.Text = Format(Val(TxtAUM.Text) * Val(TxtCMShareP.Text) * 0.00007 * 60 / 1200, "#0")
End Sub

Private Sub TxtQtySq_LostFocus()
If Trim(TxtQtySq.Text) = "" Then Exit Sub
If IsNumeric(TxtQtySq.Text) = False Or Val(TxtQtySq.Text) < 0 Then
   MsgBox "TxtQtySq numeric format"
   TxtQtySq.SetFocus
   Exit Sub
End If
TxtQtySq.Text = Format(TxtQtySq.Text, "#0.0000")
If Trim(TxtExpenseType.Text) <> "" Then
   Call AmountCalculation
   If mLExpenseID = 1 Then
      If LCase(CmbType.Text) <> "new" Then
         Call CalculateRentChange
      End If
   End If
End If
'If LCase(TxtCalculationOn.Text) = "qty(per mt)" Then
'   TxtQtySq.Text = Format(TxtQtySq.Text, "#0")
'   TxtAmount.Text = Format(Val(TxtAmountPer.Text) * Val(TxtQtySq.Text), "#0")
'ElseIf LCase(TxtCalculationOn.Text) = "area(per sq.ft)" Then
'   TxtQtySq.Text = Format(TxtQtySq.Text, "#0")
'   TxtAmount.Text = Format(Val(TxtAmountPer.Text) * Val(TxtQtySq.Text), "#0")
'ElseIf LCase(TxtCalculationOn.Text) = "number(per unit)" Then
'   TxtQtySq.Text = Format(TxtQtySq.Text, "#0")
'   TxtAmount.Text = Format(Val(TxtAmountPer.Text) * Val(TxtQtySq.Text), "#0")
'ElseIf LCase(TxtCalculationOn.Text) = "fumigation qty(per mt)" Then
'   TxtQtySq.Text = MSHFlexGrid3.TextMatrix(0, 11)
'   TxtAmount.Text = Format(Val(TxtAmountPer.Text) * Val(TxtQtySq.Text), "#0")
'ElseIf LCase(TxtCalculationOn.Text) = "fixed" Then
'   TxtQtySq.Text = 1
'   TxtAmount.Text = TxtAmountPer.Text
'End If
End Sub

Private Sub TxtReservationStock_LostFocus()
If tmp <> TxtReservationStock.Text Then
   If Val(MSHFlexGrid3.TextMatrix(0, 7)) > 0 Then
      If MsgBox("Commodity and expense detail will clear, do you want to continue?", vbYesNo) = vbNo Then
         TxtReservationStock.Text = tmp
         TxtReservationStock.SetFocus
         Exit Sub
      End If
      Call Grid_Head3
      Call ClearAllExpenseDetails
   End If
End If

If Trim(TxtReservationStock.Text) = "" Then
   TxtReservationMT.Text = ""
   Exit Sub
End If

If IsNumeric(TxtReservationStock.Text) = False Or Val(TxtReservationStock.Text) < 0 Then
   MsgBox "Invalid numeric format"
   TxtReservationStock.SetFocus
   Exit Sub
End If
If Val(TxtReservationStock.Text) > 100 Then
   MsgBox "Reservation Stock (%) never greater than 100!!!"
   TxtReservationStock.SetFocus
   Exit Sub
End If
TxtReservationStock.Text = Format(TxtReservationStock.Text, "#0")
TxtReservationMT.Text = Format(Val(TxtReservationStock.Text) * Val(TxtExpUtilisationMT.Text) / 100, "#0")
End Sub

Private Sub ClearAllExpenseDetails()
Call Grid_Head7
With MSHFlexGrid4
   For I = 2 To .Rows - 4
      If Val(.TextMatrix(I, 12)) = 1 Then
         .TextMatrix(0, 3) = Val(.TextMatrix(0, 3)) - Val(.TextMatrix(I, 3))
         .TextMatrix(I, 2) = 0
         .TextMatrix(I, 3) = 0
      End If
   Next
End With
End Sub

Private Sub TxtStorageRate_LostFocus()
If Trim(TxtStorageRate.Text) = "" Then Exit Sub
If IsNumeric(TxtStorageRate.Text) = False Or Val(TxtStorageRate.Text) < 0 Then
   MsgBox "Invalid numeric format"
   TxtStorageRate.SetFocus
   Exit Sub
End If

TxtStorageRate.Text = Format(TxtStorageRate.Text, "#0.00")
TxtRevenue.Text = Format(Val(TxtExpQtyMT.Text) * Val(TxtStorageRate.Text), "#0")
End Sub

Private Sub VisibleEnterTxt()
With TxtEnterData
   .Text = MSHFlexGrid6.TextMatrix(mRowsel, mColsel)
   .Left = MSHFlexGrid6.Left + MSHFlexGrid6.CellLeft
   .Top = MSHFlexGrid6.Top + MSHFlexGrid6.CellTop
   .Width = MSHFlexGrid6.CellWidth - 10
   .Height = MSHFlexGrid6.CellHeight - 10
   .Visible = True
   .ZOrder
   .SetFocus
End With
End Sub


Private Sub Indata()
Edit_Flg = ""

Call GButtonLock(Me, True)
Call ComboVisible(False)
Call ClearControl
Call Grid_Head6
Call Grid_Head9
Call Grid_Head10
Call Grid_Head11

Picture2.Enabled = False
Frame1.Enabled = False
Frame2.Enabled = False
Frame5.Visible = False
Frame6.Enabled = False

'TxtStaffSalaryUser.Locked = True
'TxtStaffUser.Locked = True
'TxtSecStaffUser.Locked = True
'TxtSecStaffSalaryUser.Locked = True

'TxtLength.Locked = True
'TxtWidth.Locked = True
'TxtHeight.Locked = True
TxtRemarks.Locked = True
TxtCapacity.Locked = True
TxtAddress.Locked = True
TxtArea.Locked = True
TxtGodown.Locked = True
TxtExpUtilisation.Locked = True
TxtNCDEXStock.Locked = True
TxtNNNRStock.Locked = True
TxtReservationStock.Locked = True
TxtStartDate.Enabled = False
TxtCloseDate.Enabled = False

tmp = " Select MGI.InspectionID, MGI.Type"
tmp = tmp & " ,ML.LocationName, MGI.LocationID ,MCMP.CMPName, MGI.CMPID , isnull(MGI.GodownNo, MG.GodownNo) GodownNo, MGI.GodownID, MUT.UnitType, MGI.UnitTypeID , MC.Category, MGI.GodownCategoryID ,  MGI.Capacity , MGI.Address , MGI.Area ,MGI.Status ,MGI.StartDate ,MGI.CloseDate ,MGI.CapacityUtilisation"
tmp = tmp & " ,MGI.NCDEXStock ,MGI.NNNRStock ,MGI.ReservationStock ,MGI.G_UID ,MGI.Concurrency ,MGI.CreatedDate ,MGI.CreatedBy ,MGI.UpdatedDate ,MGI.UpdatedBy"
tmp = tmp & " ,MGI.noofstaff ,MGI.staffsalary ,MGI.noofsec ,MGI.secsalary ,MGI.noofstaffuser ,MGI.staffsalaryuser ,MGI.noofsecuser ,MGI.secsalaryuser ,MGI.Remarks ,MGI.totalrevenue ,MGI.directexpense, MGI.grossprofit ,MGI.indirectexpense, MGI.netprofit, MLC.LicenseNo, MGI.LicenseID ,MS.StateID,MGI.Length,MGI.Width,MGI.Height,MGI.NoOfStack,MGI.WeightPerStack"
tmp = tmp & " from Mst_GodownInspection MGI"
tmp = tmp & " Left Join Mst_Location ML on ML.LocationID = MGI.LocationID"
tmp = tmp & " Inner Join Mst_State MS on ML.StateID = MS.StateID"
tmp = tmp & " Left Join Mst_CMP MCMP on MCMP.CMPID = MGI.CMPID"
tmp = tmp & " Left Join Mst_Godown MG on MG.GodownID = MGI.GodownID"
tmp = tmp & " Left Join Mst_UnitType MUT on MUT.UnitTypeID = MGI.UnitTypeID"
tmp = tmp & " Left Join Mst_Category MC on MC.CategoryID = MGI.GodownCategoryID"
tmp = tmp & " Left Join Mst_License MLC on MLC.LicenseID = MGI.LicenseID"
tmp = tmp & " Where MGI.InspectionID = " & RsMst_GodownInspection!InspectionID
Call Tmp7Table

If TmpRs7.RecordCount > 0 Then
   TxtType.Text = IIf(IsNull(TmpRs7!Type), "", TmpRs7!Type)
   If LCase(TxtType.Text) = "n" Then
      TxtType.Text = "New"
   ElseIf LCase(TxtType.Text) = "r" Then
      TxtType.Text = "Renewal"
   End If
   TxtInspectionID.Text = IIf(IsNull(TmpRs7!InspectionID), "", TmpRs7!InspectionID)
   TxtLocation.Text = IIf(IsNull(TmpRs7!Locationname), "", TmpRs7!Locationname)
   TxtCmp.Text = IIf(IsNull(TmpRs7!CMPName), "", TmpRs7!CMPName)
   TxtGodown.Text = IIf(IsNull(TmpRs7!GodownNo), "", TmpRs7!GodownNo)
   TxtCapacity.Text = Format(IIf(IsNull(TmpRs7!Capacity), "", TmpRs7!Capacity), "#0")
   TxtAddress.Text = IIf(IsNull(TmpRs7!Address), "", TmpRs7!Address)
   TxtGodownType.Text = IIf(IsNull(TmpRs7!UnitType), "", TmpRs7!UnitType)
   TxtGodownCategory.Text = IIf(IsNull(TmpRs7!Category), "", TmpRs7!Category)
   TxtLicenseNo.Text = IIf(IsNull(TmpRs7!LicenseNo), "", TmpRs7!LicenseNo)
   TxtNoofStack = Format(IIf(IsNull(TmpRs7!NoOfStack), "", TmpRs7!NoOfStack), "#0")
   TxtWeightPerStack = Format(IIf(IsNull(TmpRs7!WeightPerStack), "", TmpRs7!WeightPerStack), "#0.00")
   TxtArea.Text = Format(IIf(IsNull(TmpRs7!Area), "", TmpRs7!Area), "#0.000")
   TxtStatus.Text = IIf(IsNull(TmpRs7!Status), "", TmpRs7!Status)
   If LCase(TxtStatus.Text) = "p" Then
      TxtStatus.Text = "Pending"
   ElseIf LCase(TxtStatus.Text) = "a" Then
      TxtStatus.Text = "Approved"
   ElseIf LCase(TxtStatus.Text) = "r" Then
      TxtStatus.Text = "Rejected"
   ElseIf LCase(TxtStatus.Text) = "g" Then
      TxtStatus.Text = "Godown Creation"
   End If
   TxtStartDate.Value = IIf(IsNull(TmpRs7!StartDate), Date, TmpRs7!StartDate)
   TxtCloseDate.Value = IIf(IsNull(TmpRs7!CloseDate), Date, TmpRs7!CloseDate)
   
   'TxtStaffsys.Text = Format(IIf(IsNull(TmpRs7!noofstaff), "", TmpRs7!noofstaff), "#0")
   'TxtStaffSalarySys.Text = Format(IIf(IsNull(TmpRs7!staffsalary), "", TmpRs7!staffsalary), "#0")
   'TxtSecStaffSys.Text = Format(IIf(IsNull(TmpRs7!noofsec), "", TmpRs7!noofsec), "#0")
   'TxtSecStaffSalarySys.Text = Format(IIf(IsNull(TmpRs7!secsalary), "", TmpRs7!secsalary), "#0")
   'TxtStaffUser.Text = Format(IIf(IsNull(TmpRs7!noofstaffuser), "", TmpRs7!noofstaffuser), "#0")
   'TxtStaffSalaryUser.Text = Format(IIf(IsNull(TmpRs7!staffsalaryuser), "", TmpRs7!staffsalaryuser), "#0")
   'TxtSecStaffUser.Text = Format(IIf(IsNull(TmpRs7!noofsecuser), "", TmpRs7!noofsecuser), "#0")
   'TxtSecStaffSalaryUser.Text = Format(IIf(IsNull(TmpRs7!secsalaryuser), "", TmpRs7!secsalaryuser), "#0")
   
   MSHFlexGrid6.TextMatrix(2, 1) = Format(IIf(IsNull(TmpRs7!noofstaff), "", TmpRs7!noofstaff), "#0")
   MSHFlexGrid6.TextMatrix(2, 2) = Format(IIf(IsNull(TmpRs7!staffsalary), "", TmpRs7!staffsalary), "#0")
   MSHFlexGrid6.TextMatrix(3, 1) = Format(IIf(IsNull(TmpRs7!noofsec), "", TmpRs7!noofsec), "#0")
   MSHFlexGrid6.TextMatrix(3, 2) = Format(IIf(IsNull(TmpRs7!secsalary), "", TmpRs7!secsalary), "#0")
   MSHFlexGrid6.TextMatrix(2, 3) = Format(IIf(IsNull(TmpRs7!noofstaffuser), "", TmpRs7!noofstaffuser), "#0")
   MSHFlexGrid6.TextMatrix(2, 4) = Format(IIf(IsNull(TmpRs7!staffsalaryuser), "", TmpRs7!staffsalaryuser), "#0")
   MSHFlexGrid6.TextMatrix(3, 3) = Format(IIf(IsNull(TmpRs7!noofsecuser), "", TmpRs7!noofsecuser), "#0")
   MSHFlexGrid6.TextMatrix(3, 4) = Format(IIf(IsNull(TmpRs7!secsalaryuser), "", TmpRs7!secsalaryuser), "#0")
   TxtRemarks.Text = IIf(IsNull(TmpRs7!Remarks), "", TmpRs7!Remarks)
   TxtTotalRevenue.Text = Format(IIf(IsNull(TmpRs7!totalrevenue), "", TmpRs7!totalrevenue), "#0")
   TxtDirectExpense.Text = Format(IIf(IsNull(TmpRs7!DirectExpense), "", TmpRs7!DirectExpense), "#0")
   TxtGrossProfit.Text = Format(IIf(IsNull(TmpRs7!GrossProfit), "", TmpRs7!GrossProfit), "#0")
   TxtIndirectExpense.Text = Format(IIf(IsNull(TmpRs7!IndirectExpense), "", TmpRs7!IndirectExpense), "#0")
   TxtNetProfit.Text = Format(IIf(IsNull(TmpRs7!NetProfit), "", TmpRs7!NetProfit), "#0")
   TxtExpUtilisation.Text = Format(IIf(IsNull(TmpRs7!CapacityUtilisation), "", TmpRs7!CapacityUtilisation), "#0")
   TxtCurrentUtilisation.Text = TxtExpUtilisation.Text
   TxtNCDEXStock.Text = Format(IIf(IsNull(TmpRs7!NcdexStock), "", TmpRs7!NcdexStock), "#0")
   TxtNNNRStock.Text = Format(IIf(IsNull(TmpRs7!NNNRStock), "", TmpRs7!NNNRStock), "#0")
   TxtReservationStock.Text = Format(IIf(IsNull(TmpRs7!ReservationStock), "", TmpRs7!ReservationStock), "#0")
   TxtCreated.Text = IIf(IsNull(TmpRs7!CreatedBy), "", TmpRs7!CreatedBy & " :- " & TmpRs7!CreatedDate)
   TxtUpdated.Text = IIf(IsNull(TmpRs7!UpdatedBy), "", TmpRs7!UpdatedBy & " :- " & TmpRs7!UpdatedDate)
   Call CapacityCalculation
   
   CmbType.Text = TxtType.Text
   CmbLocation.Text = TxtLocation.Text
   CmbCMP.Text = TxtCmp.Text
   CmbGodown.Text = TxtGodown.Text
   CmbGodownType.Text = TxtGodownType.Text
   CmbGodownCategory.Text = TxtGodownCategory.Text
   CmbLicenseNo.Text = TxtLicenseNo.Text
   CmbStatus.Text = TxtStatus.Text
   
   mLocationID = IIf(IsNull(TmpRs7!LocationID), "", TmpRs7!LocationID)
   mCMPID = IIf(IsNull(TmpRs7!CMPID), "", TmpRs7!CMPID)
   mGodownID = IIf(IsNull(TmpRs7!GodownID), "", TmpRs7!GodownID)
   mGodownTypeID = IIf(IsNull(TmpRs7!UnitTypeID), "", TmpRs7!UnitTypeID)
   mGodownCategoryID = IIf(IsNull(TmpRs7!GodownCategoryID), "", TmpRs7!GodownCategoryID)
   mLicenseID = IIf(IsNull(TmpRs7!LicenseID), "", TmpRs7!LicenseID)
   mStateID = IIf(IsNull(TmpRs7!StateID), "", TmpRs7!StateID)
   
   If mGodownCategoryID = 2 Then
      If Val(mGodownTypeID) = 1 Or Val(mGodownTypeID) = 7 Then
         TxtArea.BackColor = &H80000004
      Else
         TxtArea.BackColor = &H80000005
      End If
   Else
      TxtArea.BackColor = &H80000005
   End If
   
End If
Call IndataInspectionGodownDetail
Call IndataInspectionCommodityDetail
Call IndataInspectionExpenseDetail
If mGodownCategoryID = 2 Then
   If Val(mGodownTypeID) = 1 Or Val(mGodownTypeID) = 7 Then
      Call IndataInspectionAreaDetail
   End If
End If

'====================='for security and staff salary expense'=====================
Call SecandUserSalaryExpense
'====================='for security and staff salary expense'=====================

Call FillCloseGodowns
Call IndataProfileSummaryDetail


Call IndataGodownUtilizationDetail
Call IndataGodownPnLDetail

End Sub

Private Sub IndataProfileSummaryDetail()
Dim mProportionalFormula As String
Dim mDirectExpenseProportional As Double, mInDirectExpenseProportional As Double
Dim mDirectExpenseInProportional As Double, mInDirectExpenseInProportional As Double
''contain godownincome, aum, fumigableqty, ncdexincome then it is proportional to capacity
Call Grid_Head7
For I = 2 To MSHFlexGrid4.Rows - 2
   mProportionalFormula = MSHFlexGrid4.TextMatrix(I, 8)
   If LCase(MSHFlexGrid4.TextMatrix(I, 10)) = "d" Then
      If InStr(1, mProportionalFormula, "fumigableqty", vbTextCompare) > 0 Or InStr(1, mProportionalFormula, "aum", vbTextCompare) > 0 Or InStr(1, mProportionalFormula, "ncdexincome", vbTextCompare) > 0 Or InStr(1, mProportionalFormula, "godownincome", vbTextCompare) > 0 Then
         mDirectExpenseProportional = mDirectExpenseProportional + Val(MSHFlexGrid4.TextMatrix(I, 3))
      Else
         mDirectExpenseInProportional = mDirectExpenseInProportional + Val(MSHFlexGrid4.TextMatrix(I, 3))
      End If
   ElseIf LCase(MSHFlexGrid4.TextMatrix(I, 10)) = "i" Then
      If InStr(1, mProportionalFormula, "fumigableqty", vbTextCompare) > 0 Or InStr(1, mProportionalFormula, "aum", vbTextCompare) > 0 Or InStr(1, mProportionalFormula, "ncdexincome", vbTextCompare) > 0 Or InStr(1, mProportionalFormula, "godownincome", vbTextCompare) > 0 Then
         mInDirectExpenseProportional = mInDirectExpenseProportional + Val(MSHFlexGrid4.TextMatrix(I, 3))
      Else
         mInDirectExpenseInProportional = mInDirectExpenseInProportional + Val(MSHFlexGrid4.TextMatrix(I, 3))
      End If
   End If
Next

With MSHFlexGrid7
   If Val(TxtExpUtilisation.Text) > 0 Then
      For j = 1 To .Rows - 1
         .TextMatrix(j, 1) = Format(Val(.TextMatrix(j, 0)) * Val(TxtTotalRevenue.Text) / Val(TxtExpUtilisation.Text), "#0")
         .TextMatrix(j, 2) = Format((Val(.TextMatrix(j, 0)) * mDirectExpenseProportional / Val(TxtExpUtilisation.Text)) + mDirectExpenseInProportional, "#0")
         .TextMatrix(j, 3) = Format(Val(.TextMatrix(j, 1)) - Val(.TextMatrix(j, 2)), "#0")
         .TextMatrix(j, 4) = Format((Val(.TextMatrix(j, 0)) * mInDirectExpenseProportional / Val(TxtExpUtilisation.Text)) + mInDirectExpenseInProportional, "#0")
         .TextMatrix(j, 5) = Format(Val(.TextMatrix(j, 3)) - Val(.TextMatrix(j, 4)), "#0")
      Next
   End If
End With
End Sub

Private Sub IndataInspectionGodownDetail()
Call Grid_Head2
tmp = " Select MCMP.CMPName, MG.GodownNo, MUT.UnitType, MIGD.Capacity, MIGD.Rent, MG.Address, MIGD.Stock, MIGD.GodownID, MIGD.Area"
tmp = tmp & " From Mst_InspectionGodownDetail MIGD"
tmp = tmp & " Inner Join Mst_Godown MG on MG.GodownID = MIGD.GodownID"
tmp = tmp & " Inner Join Mst_CMP MCMP on MCMP.CMPID = MG.CMPID"
tmp = tmp & " Inner Join Mst_Location ML on ML.LocationID = MCMP.LocationID"
tmp = tmp & " Inner Join Mst_UnitType MUT on MUT.UnitTypeID = MG.UnitTypeID"
tmp = tmp & " Where MIGD.InspectionID = " & Val(TxtInspectionID.Text)
tmp = tmp & " order by MCMP.CMPName, MG.GodownNo"
Call Tmp7Table
If TmpRs7.RecordCount > 0 Then
   With MSHFlexGrid2
      .Rows = .Rows + TmpRs7.RecordCount
      .TextMatrix(0, 1) = "Total No of Godown : " & TmpRs7.RecordCount
      For I = 2 To TmpRs7.RecordCount + 1
         .TextMatrix(I, 0) = IIf(IsNull(TmpRs7!CMPName), "", TmpRs7!CMPName)
         .TextMatrix(I, 1) = IIf(IsNull(TmpRs7!GodownNo), "", TmpRs7!GodownNo)
         .TextMatrix(I, 2) = IIf(IsNull(TmpRs7!UnitType), "", TmpRs7!UnitType)
         .TextMatrix(I, 3) = Format(IIf(IsNull(TmpRs7!Capacity), "0", TmpRs7!Capacity), "#0")
         .TextMatrix(I, 4) = Format(IIf(IsNull(TmpRs7!Stock), 0, TmpRs7!Stock), "#0")
         .TextMatrix(I, 5) = Format(IIf(IsNull(TmpRs7!Area), "", TmpRs7!Area), "#0")
         .TextMatrix(I, 6) = Format(IIf(IsNull(TmpRs7!Rent), "0", TmpRs7!Rent), "#0")
         .TextMatrix(I, 7) = IIf(IsNull(TmpRs7!Address), "", TmpRs7!Address)
         .TextMatrix(I, 8) = IIf(IsNull(TmpRs7!GodownID), "", TmpRs7!GodownID)
         
         .TextMatrix(0, 3) = Format(Val(.TextMatrix(0, 3)) + Val(.TextMatrix(I, 3)), "#0")
         .TextMatrix(0, 4) = Format(Val(.TextMatrix(0, 4)) + Val(.TextMatrix(I, 4)), "#0")
         If Val(.TextMatrix(I, 4)) > 0 And Val(.TextMatrix(I, 6)) > 0 Then ' for zero stock
            .TextMatrix(0, 5) = Format(Val(.TextMatrix(0, 5)) + Val(.TextMatrix(I, 5)), "#0")
         Else
            .row = I
            If Val(.TextMatrix(I, 4)) = 0 And Val(.TextMatrix(I, 6)) = 0 Then
               For j = 0 To .Cols - 1
                  .Col = j
                  .CellBackColor = vbRed
               Next
            ElseIf Val(.TextMatrix(I, 4)) = 0 Then
               .Col = 4
               .CellBackColor = vbRed
            ElseIf Val(.TextMatrix(I, 6)) = 0 Then
               .Col = 6
               .CellBackColor = vbRed
            End If
            
            .TextMatrix(0, 7) = Format(Val(.TextMatrix(0, 7)) + Val(.TextMatrix(I, 5)), "#0")
         End If
         .TextMatrix(0, 6) = Format(Val(.TextMatrix(0, 6)) + Val(.TextMatrix(I, 6)), "#0")
         
         TmpRs7.MoveNext
      Next
      If Val(.TextMatrix(0, 5)) > 0 Then
         .TextMatrix(0, 2) = "Average Rent (per Sq.Ft) : " & Format(Val(.TextMatrix(0, 6)) / Val(.TextMatrix(0, 5)), "#0.00")
      End If
      If Val(.TextMatrix(0, 7)) > 0 Then
         .TextMatrix(0, 7) = "Area with zero stock and rent : " & .TextMatrix(0, 7)
      End If
   End With
End If
End Sub

Private Sub IndataInspectionCommodityDetail()

Call ClearCommodityDetail
Call Grid_Head3
tmp = " Select MICD.InspectionCommodityID ,MICD.InspectionID ,MCom.Commodity, MICD.Category ,MICD.CommodityID ,MICD.ExpQtyP ,MICD.ExpQtyMT"
tmp = tmp & "  ,MICD.StorageRate ,MICD.Price ,MICD.Revenue ,MICD.AUM ,MICD.StorageRateRemarks, MICD.FumigableQty, MCom.FumigationFlag, MICD.CMShareP, MICD.CMShareR"
tmp = tmp & " From Mst_InspectionCommodityDetail MICD"
tmp = tmp & " Inner Join Mst_commodity MCom on MCom.CommodityID = MICD.CommodityID"
tmp = tmp & " Where MICD.InspectionID = " & Val(TxtInspectionID.Text)
Call Tmp7Table
If TmpRs7.RecordCount > 0 Then
   With MSHFlexGrid3
      .Rows = .Rows + TmpRs7.RecordCount
      For I = 2 To TmpRs7.RecordCount + 1
         .TextMatrix(I, 0) = IIf(IsNull(TmpRs7!Category), "", TmpRs7!Category)
         If LCase(.TextMatrix(I, 0)) = "nc" Then
            .TextMatrix(I, 0) = "NCDEX"
            mNCDEXPer = mNCDEXPer + IIf(IsNull(TmpRs7!ExpQtyP), 0, TmpRs7!ExpQtyP)
         ElseIf LCase(.TextMatrix(I, 0)) = "nn" Then
            .TextMatrix(I, 0) = "NNNR"
            mNNNRPer = mNNNRPer + IIf(IsNull(TmpRs7!ExpQtyP), 0, TmpRs7!ExpQtyP)
         ElseIf LCase(.TextMatrix(I, 0)) = "re" Then
            .TextMatrix(I, 0) = "Reservation"
            mReservationPer = mReservationPer + IIf(IsNull(TmpRs7!ExpQtyP), 0, TmpRs7!ExpQtyP)
         End If
         .TextMatrix(I, 1) = IIf(IsNull(TmpRs7!Commodity), "", TmpRs7!Commodity)
         .TextMatrix(I, 2) = IIf(IsNull(TmpRs7!ExpQtyP), "", TmpRs7!ExpQtyP)
         .TextMatrix(I, 3) = IIf(IsNull(TmpRs7!ExpQtyMT), "", TmpRs7!ExpQtyMT)
         .TextMatrix(I, 4) = IIf(IsNull(TmpRs7!StorageRate), "", TmpRs7!StorageRate)
         .TextMatrix(I, 5) = IIf(IsNull(TmpRs7!Price), "", TmpRs7!Price)
         .TextMatrix(I, 6) = IIf(IsNull(TmpRs7!Revenue), "", TmpRs7!Revenue)
         .TextMatrix(I, 7) = IIf(IsNull(TmpRs7!AUM), "", TmpRs7!AUM)
         .TextMatrix(I, 8) = IIf(IsNull(TmpRs7!StorageRateRemarks), "", TmpRs7!StorageRateRemarks)
         .TextMatrix(I, 9) = IIf(IsNull(TmpRs7!InspectionCommodityID), "", TmpRs7!InspectionCommodityID)
         .TextMatrix(I, 10) = IIf(IsNull(TmpRs7!CommodityID), "", TmpRs7!CommodityID)
         .TextMatrix(I, 11) = IIf(IsNull(TmpRs7!FumigableQty), "", TmpRs7!FumigableQty)
         .TextMatrix(I, 12) = IIf(IsNull(TmpRs7!FumigationFlag), "", TmpRs7!FumigationFlag)
         .TextMatrix(I, 13) = IIf(IsNull(TmpRs7!CMShareP), "", TmpRs7!CMShareP)
         .TextMatrix(I, 14) = IIf(IsNull(TmpRs7!CMShareR), "", TmpRs7!CMShareR)
         
         .TextMatrix(0, 3) = Val(.TextMatrix(0, 3)) + Val(.TextMatrix(I, 3))
         .TextMatrix(0, 6) = Val(.TextMatrix(0, 6)) + Val(.TextMatrix(I, 6))
         .TextMatrix(0, 7) = Val(.TextMatrix(0, 7)) + Val(.TextMatrix(I, 7))
         .TextMatrix(0, 11) = Val(.TextMatrix(0, 11)) + Val(.TextMatrix(I, 11))
         .TextMatrix(0, 14) = Val(.TextMatrix(0, 14)) + Val(.TextMatrix(I, 14))
         TmpRs7.MoveNext
      Next
      TxtCMShare.Text = Format(Val(.TextMatrix(0, 14)), "#0")
   End With
End If
End Sub

Private Sub IndataInspectionExpenseDetail()
Call ClearExpenseDetail
Call Grid_Head4
'tmp = " Select  MICE.InspectionExpenseID ,MICE.InspectionID ,MICE.InsExpTypeID ,MICE.QtySQ ,MICE.Amount ,MICE.Remarks, MIE.ExpenseType, MIE.CalculationOn, MIE.Amount Amountper, MICE.Formula, MIE.UserInputAmount"
tmp = " Select  MICE.InspectionExpenseID ,MICE.InspectionID ,MICE.InsExpTypeID ,MICE.QtySQ ,MICE.Amount ,MICE.Remarks, MIE.ExpenseType, MIE.CalculationOn, MICE.AmountPer, MIE.Formula, MIE.UserInputAmount, MIE.Type, MIE.UserInputCalc, MIE.ClearFlag, MIE.Remarks 'ExpenseRemarks'"
tmp = tmp & " From Mst_InspectionExpenseDetail MICE"
tmp = tmp & " Inner Join Mst_InspectionExpenseType MIE on MIE.InsExpTypeID = MICE.InsExpTypeID"
tmp = tmp & " Where MICE.InspectionID = " & Val(TxtInspectionID.Text)
tmp = tmp & " Order by MIE.OrderNo"
Call Tmp7Table

With MSHFlexGrid4
   If TmpRs7.RecordCount > 0 Then
      .Rows = .Rows + TmpRs7.RecordCount
      For I = 2 To TmpRs7.RecordCount + 1
         .TextMatrix(I, 0) = IIf(IsNull(TmpRs7!ExpenseType), "", TmpRs7!ExpenseType)
         .TextMatrix(I, 1) = IIf(IsNull(TmpRs7!CalculationOn), "", TmpRs7!CalculationOn)
         .TextMatrix(I, 2) = Format(IIf(IsNull(TmpRs7!QtySQ), "", TmpRs7!QtySQ), "#0")
         .TextMatrix(I, 3) = Format(IIf(IsNull(TmpRs7!Amount), "", TmpRs7!Amount), "#0")
         .TextMatrix(0, 3) = Format(Val(.TextMatrix(0, 3)) + Val(.TextMatrix(I, 3)), "#0")
         .TextMatrix(I, 4) = IIf(IsNull(TmpRs7!Remarks), "", TmpRs7!Remarks)
         .TextMatrix(I, 5) = IIf(IsNull(TmpRs7!InspectionExpenseID), "", TmpRs7!InspectionExpenseID)
         .TextMatrix(I, 6) = IIf(IsNull(TmpRs7!InsExpTypeID), "", TmpRs7!InsExpTypeID)
         .TextMatrix(I, 7) = Format(IIf(IsNull(TmpRs7!AmountPer), "", TmpRs7!AmountPer), "#0.00")
         .TextMatrix(I, 8) = IIf(IsNull(TmpRs7!Formula), "", TmpRs7!Formula)
         .TextMatrix(I, 9) = IIf(IsNull(TmpRs7!UserInputAmount), "", TmpRs7!UserInputAmount)
         .TextMatrix(I, 10) = IIf(IsNull(TmpRs7!Type), "", TmpRs7!Type)
         .TextMatrix(I, 11) = IIf(IsNull(TmpRs7!UserInputCalc), "", TmpRs7!UserInputCalc)
         .TextMatrix(I, 12) = IIf(IsNull(TmpRs7!ClearFlag), "", TmpRs7!ClearFlag)
         .TextMatrix(I, 13) = IIf(IsNull(TmpRs7!ExpenseRemarks), "", TmpRs7!ExpenseRemarks)
         If TmpRs7!InsExpTypeID = 1 And mGodownTypeID = 1 Then
            If LCase(TxtType.Text) = "renewal" Then
               Call GetRentChange(I)
            End If
         End If
         TmpRs7.MoveNext
      Next
   End If
End With
End Sub

Private Sub SecandUserSalaryExpense()
With MSHFlexGrid4
   I = .Rows - 1
   .Rows = .Rows + 1
   .TextMatrix(I, 0) = "Security Salary Cost"
   .TextMatrix(I, 1) = "Fixed"
   .TextMatrix(I, 2) = Val(MSHFlexGrid6.TextMatrix(3, 3))
   If LCase(CmbType.Text) = "renewal" And Val(MSHFlexGrid2.TextMatrix(0, 3)) > 0 Then
      .TextMatrix(I, 3) = Format(Val(MSHFlexGrid6.TextMatrix(3, 4)) * Val(TxtCapacity.Text) / Val(MSHFlexGrid2.TextMatrix(0, 3)), "#0")
   ElseIf LCase(CmbType.Text) = "new" And (Val(MSHFlexGrid2.TextMatrix(0, 3)) + Val(TxtCapacity.Text)) > 0 Then
      .TextMatrix(I, 3) = Format(Val(MSHFlexGrid6.TextMatrix(3, 4)) * Val(TxtCapacity.Text) / (Val(MSHFlexGrid2.TextMatrix(0, 3)) + Val(TxtCapacity.Text)), "#0")
   End If
   .TextMatrix(0, 3) = Format(Val(.TextMatrix(0, 3)) + Val(.TextMatrix(I, 3)), "#0")
   .TextMatrix(I, 4) = "Security Salary Cost * Godown Capacity / Location Capacity "
   .TextMatrix(I, 5) = ""
   .TextMatrix(I, 6) = ""
   If LCase(CmbType.Text) = "renewal" And Val(MSHFlexGrid2.TextMatrix(0, 3)) > 0 Then
      .TextMatrix(I, 7) = Format(Val(MSHFlexGrid6.TextMatrix(3, 4)) * Val(TxtCapacity.Text) / Val(MSHFlexGrid2.TextMatrix(0, 3)), "#0")
   ElseIf LCase(CmbType.Text) = "new" And (Val(MSHFlexGrid2.TextMatrix(0, 3)) + Val(TxtCapacity.Text)) > 0 Then
      .TextMatrix(I, 7) = Format(Val(MSHFlexGrid6.TextMatrix(3, 4)) * Val(TxtCapacity.Text) / (Val(MSHFlexGrid2.TextMatrix(0, 3)) + Val(TxtCapacity.Text)), "#0")
   End If
   .TextMatrix(I, 8) = ""
   .TextMatrix(I, 9) = ""
   .TextMatrix(I, 10) = "d"
   
   I = .Rows - 1
   .Rows = .Rows + 1
   .TextMatrix(I, 0) = "Staff Salary Cost"
   .TextMatrix(I, 1) = "Fixed"
   .TextMatrix(I, 2) = Val(MSHFlexGrid6.TextMatrix(2, 3))
   If LCase(CmbType.Text) = "renewal" And Val(MSHFlexGrid2.TextMatrix(0, 3)) > 0 Then
      .TextMatrix(I, 3) = Format(Val(MSHFlexGrid6.TextMatrix(2, 4)) * Val(TxtCapacity.Text) / Val(MSHFlexGrid2.TextMatrix(0, 3)), "#0")
   ElseIf LCase(CmbType.Text) = "new" And (Val(MSHFlexGrid2.TextMatrix(0, 3)) + Val(TxtCapacity.Text)) > 0 Then
      .TextMatrix(I, 3) = Format(Val(MSHFlexGrid6.TextMatrix(2, 4)) * Val(TxtCapacity.Text) / (Val(MSHFlexGrid2.TextMatrix(0, 3)) + Val(TxtCapacity.Text)), "#0")
   End If
   .TextMatrix(0, 3) = Format(Val(.TextMatrix(0, 3)) + Val(.TextMatrix(I, 3)), "#0")
   .TextMatrix(I, 4) = "Staff Salary Cost * Godown Capacity / Location Capacity "
   .TextMatrix(I, 5) = ""
   .TextMatrix(I, 6) = ""
   If LCase(CmbType.Text) = "renewal" And Val(MSHFlexGrid2.TextMatrix(0, 3)) > 0 Then
      .TextMatrix(I, 7) = Format(Val(MSHFlexGrid6.TextMatrix(2, 4)) * Val(TxtCapacity.Text) / Val(MSHFlexGrid2.TextMatrix(0, 3)), "#0")
   ElseIf LCase(CmbType.Text) = "new" And (Val(MSHFlexGrid2.TextMatrix(0, 3)) + Val(TxtCapacity.Text)) > 0 Then
      .TextMatrix(I, 7) = Format(Val(MSHFlexGrid6.TextMatrix(2, 4)) * Val(TxtCapacity.Text) / (Val(MSHFlexGrid2.TextMatrix(0, 3)) + Val(TxtCapacity.Text)), "#0")
   End If
   .TextMatrix(I, 8) = ""
   .TextMatrix(I, 9) = ""
   .TextMatrix(I, 10) = "d"
End With

End Sub

Private Function GetMaxGodownID() As Double
Dim Retval As Double
tmp = "Select max(GodownID) from Mst_Godown"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
TmpRs.Close
GetMaxGodownID = Retval
End Function

Private Sub Gen_Email()
Dim TOADD As Variant
Dim TOCC, mMsg As Variant

mMsg = ""
TOADD = ""
TOCC = ""
mMsg = "<table border=1>"
mMsg = mMsg & "<tr><td>"
mMsg = mMsg & "<table>"

mMsg = mMsg & "<tr>"
mMsg = mMsg & "<td><table border=1>"
mMsg = mMsg & "<tr><td colspan=3 style=""background-color: #ffff99"">Godown Viability (From Store-Man)</td></tr>"
mMsg = mMsg & "<tr><td>Viability ID</td><td  colspan=2>" & TxtInspectionID.Text & "</td> </tr>"
mMsg = mMsg & "<tr><td>Type</td><td  colspan=2>" & TxtType.Text & "</td> </tr>"
mMsg = mMsg & "<tr><td>Location</td><td  colspan=2>" & TxtLocation.Text & "</td> </tr>"
mMsg = mMsg & "<tr><td>CMP Name</td><td  colspan=2>" & TxtCmp.Text & "</td> </tr>"
mMsg = mMsg & "<tr><td>Godown No</td><td  colspan=2>" & TxtGodown.Text & "</td> </tr>"
mMsg = mMsg & "<tr><td>Godown Type</td><td  colspan=2>" & TxtGodownType.Text & "</td> </tr>"
mMsg = mMsg & "<tr><td>License No</td><td  colspan=2>" & TxtLicenseNo.Text & "</td> </tr>"
mMsg = mMsg & "<tr><td>Address</td><td  colspan=2>" & TxtAddress.Text & "</td> </tr>"
mMsg = mMsg & "<tr><td>Area in Sq.Ft.</td><td  colspan=2 align = right>" & TxtArea.Text & "</td> </tr>"
mMsg = mMsg & "<tr><td>Capacity in MT</td><td  colspan=2 align = right>" & TxtCapacity.Text & "</td> </tr>"
mMsg = mMsg & "<tr><td>Category</td><td  colspan=2>" & TxtGodownCategory.Text & "</td> </tr>"
mMsg = mMsg & "<tr><td>Status</td><td  colspan=2>" & TxtStatus.Text & "</td> </tr>"
mMsg = mMsg & "<tr><td>Start Date</td><td  colspan=2>" & Format(TxtStartDate.Value, "dd-MMM-yyyy") & "</td> </tr>"
mMsg = mMsg & "<tr><td>Close Date</td><td  colspan=2>" & Format(TxtCloseDate.Value, "dd-MMM-yyyy") & "</td> </tr>"
tmp = "Select EmployeeName from Mst_Employee Where EmployeeNo = '" & IIf(IsNull(RsMst_GodownInspection!CreatedBy), 0, RsMst_GodownInspection!CreatedBy) & "'"
Call Tmp4Table
If TmpRs4.RecordCount > 0 Then
   mCreated = IIf(IsNull(TmpRs4!EmployeeName), "", TmpRs4!EmployeeName) & " :- " & IIf(IsNull(RsMst_GodownInspection!CreatedDate), "", RsMst_GodownInspection!CreatedDate)
   mMsg = mMsg & "<tr><td>Created By</td><td colspan=2>" & mCreated & "</td></tr>"
Else
   mMsg = mMsg & "<tr><td>Created By</td><td colspan=2>" & TxtCreated.Text & "</td></tr>"
End If

tmp = "Select EmployeeName from Mst_Employee Where EmployeeNo = '" & IIf(IsNull(RsMst_GodownInspection!UpdatedBy), 0, RsMst_GodownInspection!UpdatedBy) & "'"
Call Tmp4Table
If TmpRs4.RecordCount > 0 Then
   mUpdated = IIf(IsNull(TmpRs4!EmployeeName), "", TmpRs4!EmployeeName) & " :- " & IIf(IsNull(RsMst_GodownInspection!UpdatedDate), "", RsMst_GodownInspection!UpdatedDate)
   mMsg = mMsg & "<tr><td>Updated By</td><td colspan=2>" & mUpdated & "</td></tr>"
Else
   mMsg = mMsg & "<tr><td>Updated By</td><td colspan=2>" & TxtUpdated.Text & "</td></tr>"
End If

mMsg = mMsg & "</table></td>"

mMsg = mMsg & "<td valign = top><table border=1>"

mMsg = mMsg & "<tr><td colspan=6 style=""background-color: #ffff99""> Profit Summary at different Utilisation</td></tr>"
For I = 0 To MSHFlexGrid7.Rows - 1
   mMsg = mMsg & "<tr>"
   For j = 0 To MSHFlexGrid7.Cols - 1
      If MSHFlexGrid7.ColWidth(j) > 0 Then
         If I = 0 Then
            mMsg = mMsg & "<td> " & MSHFlexGrid7.TextMatrix(I, j) & " </td>"
         Else
            mMsg = mMsg & "<td align = right> " & MSHFlexGrid7.TextMatrix(I, j) & " </td>"
         End If
      End If
   Next
   mMsg = mMsg & "</tr>"
Next
mMsg = mMsg & "<tr style=""background-color: #ffff99"">"
mMsg = mMsg & "<td align = right> Projected Utilisation : - " & TxtCurrentUtilisation & " </td>"
mMsg = mMsg & "<td align = right> " & TxtTotalRevenue & " </td>"
mMsg = mMsg & "<td align = right> " & TxtDirectExpense & " </td>"
mMsg = mMsg & "<td align = right> " & TxtGrossProfit & " </td>"
mMsg = mMsg & "<td align = right> " & TxtIndirectExpense & " </td>"
mMsg = mMsg & "<td align = right> " & TxtNetProfit & " </td>"
mMsg = mMsg & "</tr>"
mMsg = mMsg & "</table></td></tr>"
mMsg = mMsg & "</table></td></tr>"
        
mMsg = mMsg & "<tr><td><table><tr><td><table border=1>"
mMsg = mMsg & "<tr><td style=""background-color: #ffff99""> Utilisation Details </td><td style=""background-color: #ffff99""> (in %) </td><td style=""background-color: #ffff99""> Qty (in MT) </td></tr>"
mMsg = mMsg & "<tr><td> Expected Capacity Utilisation </td> <td align = right>" & TxtExpUtilisation.Text & "</td> <td align = right>" & TxtExpUtilisationMT.Text & "</td></tr>"
mMsg = mMsg & "<tr><td> NCDEX Stock </td> <td align = right>" & TxtNCDEXStock.Text & "</td> <td align = right>" & TxtNCDEXMT.Text & "</td></tr>"
mMsg = mMsg & "<tr><td> NNNR (including SCM) Stock </td> <td align = right>" & TxtNNNRStock.Text & "</td> <td align = right>" & TxtNNNRMT.Text & "</td></tr>"
mMsg = mMsg & "<tr><td> Reservation Stock </td> <td align = right>" & TxtReservationStock.Text & "</td> <td align = right>" & TxtReservationMT.Text & "</td></tr>"
'mMsg = mMsg & "</table></td></tr>"
mMsg = mMsg & "</table></td>"
mMsg = mMsg & "<td valign = top><table border=1>"
'mMsg = mMsg & "<tr><td><table border=1>"
mMsg = mMsg & "<tr><td colspan=5 style=""background-color: #ffff99""> Staff Details </td></tr>"
For I = 0 To MSHFlexGrid6.Rows - 1
   mMsg = mMsg & "<tr>"
   For j = 0 To MSHFlexGrid6.Cols - 1
      If I > 1 And j > 0 Then
         mMsg = mMsg & "<td align = right> " & MSHFlexGrid6.TextMatrix(I, j) & " </td>"
      Else
         mMsg = mMsg & "<td> " & MSHFlexGrid6.TextMatrix(I, j) & " </td>"
      End If
   Next
   mMsg = mMsg & "</tr>"
Next
mMsg = mMsg & "</table></td></tr></table></td></tr>"

mMsg = mMsg & "<tr><td><table border=1>"
mMsg = mMsg & "<tr><td colspan=12 style=""background-color: #ffff99""> Commodity Details </td></tr>"
For I = 0 To MSHFlexGrid3.Rows - 2
   mMsg = mMsg & "<tr>"
   For j = 0 To MSHFlexGrid3.Cols - 1
      If MSHFlexGrid3.ColWidth(j) > 0 Then
         If I > 1 And j > 1 Then
            mMsg = mMsg & "<td align = right> " & MSHFlexGrid3.TextMatrix(I, j) & " </td>"
         Else
            mMsg = mMsg & "<td> " & MSHFlexGrid3.TextMatrix(I, j) & " </td>"
         End If
      End If
   Next
   mMsg = mMsg & "</tr>"
Next
mMsg = mMsg & "</table></td></td></tr>"

mMsg = mMsg & "<tr><td><table border=1>"
mMsg = mMsg & "<tr><td colspan=6 style=""background-color: #ffff99""> Expense Details </td></tr>"
For I = 0 To MSHFlexGrid4.Rows - 2
   mMsg = mMsg & "<tr>"
   For j = 0 To MSHFlexGrid4.Cols - 1
      If MSHFlexGrid4.ColWidth(j) > 0 Then
         If I > 1 And j > 1 Then
            mMsg = mMsg & "<td align = right> " & MSHFlexGrid4.TextMatrix(I, j) & " </td>"
         Else
            mMsg = mMsg & "<td> " & MSHFlexGrid4.TextMatrix(I, j) & " </td>"
         End If
      End If
   Next
   mMsg = mMsg & "</tr>"
Next
mMsg = mMsg & "</table></td></tr>"
mMsg = mMsg & "</table>"

Call GetEmailIDsforGodownViability
TOADD = mApprovalmailIDs '"priti.g@ncml.COM, kuldip.s@ncml.com"
TOCC = Gen_UserEmailID & ",manu.ag@ncml.com,anilkumar.par@NCML.COM,taufique.k@ncml.com"

tmp = " Select Distinct  ME1.EmployeeEmailID 'AMID', ME2.EmployeeEmailID 'RMID'"
tmp = tmp & " From Mst_EmpCMPMapping MECM"
tmp = tmp & " Inner Join Mst_CMP MC On MECM.CMPID = MC.CMPID"
tmp = tmp & " Inner Join Mst_Employee ME On MECM.EmployeeID = ME.EmployeeID And ME.Flag = 'Y'"
tmp = tmp & " Inner Join Mst_Employee ME1 On ME.DirectReportingTo = ME1.EmployeeID And ME1.Flag = 'Y'"
tmp = tmp & " Inner Join Mst_Employee ME2 On ME1.DirectReportingTo = ME2.EmployeeID And ME2.Flag = 'Y'"
tmp = tmp & " Where MC.LocationID = " & mLocationID
Call Tmp1Table
If TmpRs1.RecordCount > 0 Then
   TOCC = TOCC & IIf(IsNull(TmpRs1!AMID), "", "," & TmpRs1!AMID)
   TOCC = TOCC & IIf(IsNull(TmpRs1!RMID), "", "," & TmpRs1!RMID)
End If

mMsgSubject = "Godown Viability"

Call SendMail(TOADD, TOCC, mMsgSubject, mMsg)

End Sub

Private Function SendMail(ToADD_ As Variant, TOCC_ As Variant, MsgSubject_ As Variant, Message1_ As Variant)
Dim htmlStringhead As String
Dim htmlStringtail As String
Set oSMTP = New OSSMTP.SMTPSession

servertxt = "mail.ncml.com"

htmlStringhead = "<html><body><font size= 3 face= Zurich BT color = Blue><table border = 1>"
htmlStringtail = "</table></Font></body></html>"
   
'With oSMTP
'  .Server = servertxt
'  .MailFrom = Gen_UserEmailID 'TOCC_ '"mayur.d@ncml.com"
'  .MessageHTML = htmlStringhead & Message1_ & htmlStringtail
'  If mSaveFail = True Then
'     .SendTo = "mayur.d@ncml.com"
'     .CC = "parag.r@ncml.com"
'     .BCC = "rakesh.go@ncml.com"
'     .MessageSubject = MsgSubject_ & " Save Failure (from Store-Man)" '"Confidential Mail"
'  Else
'     .SendTo = ToADD_
'     .BCC = "parag.r@ncml.com"
'     If ToADD_ <> TOCC_ Then
'         .CC = TOCC_
'     End If
'     .MessageSubject = MsgSubject_ & " (from Store-Man)" '"Confidential Mail"
'  End If
'
'  'comment when live it
'  '.SendTo = "rakesh.go@ncml.com" '
'  '.CC = Gen_UserEmailID
'  'comment when live it
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

If mSaveFail = True Then
   objSMTP.ToAddr = "mayur.d@ncml.com"
   objSMTP.CCAddr = "parag.r@ncml.com"
   objSMTP.Subject = MsgSubject_ & " Save Failure (from Store-Man)" '"Confidential Mail"
Else
   objSMTP.ToAddr = ToADD_
   If ToADD_ <> TOCC_ Then
      objSMTP.CCAddr = TOCC_
   End If
   objSMTP.Subject = MsgSubject_ & " (from Store-Man)" '"Confidential Mail"
End If
'objSMTP.ToAddr = ToADD_
'If ToADD_ <> TOCC_ Then
'   objSMTP.CCAddr = TOCC_
'End If

'objSMTP.Subject = MsgSubject_ & " (from Store-Man)" 'MsgSubject_ '"Message in HTML format"
objSMTP.BodyFormat = 1
objSMTP.BodyText = htmlStringhead & Message1_ & htmlStringtail 'htmlStringhead & MsgSubject_ & " :- " & Message1_ & htmlStringtail '"<html>This word is <b>bold</b></html>"
If LCase(Gen_EmailConfigFlag) = "y" Then
   objSMTP.Send
End If

End Function

Private Function ChkExpenseDetail() As Boolean
ChkExpenseDetail = True
If Val(MSHFlexGrid4.TextMatrix(0, 3)) = 0 Then
   MsgBox "Blank Expense detail not allowed!!!"
   ChkExpenseDetail = False
   Exit Function
End If
For l = 2 To MSHFlexGrid4.Rows - 2
   If Val(MSHFlexGrid4.TextMatrix(l, 3)) = 0 And Trim(MSHFlexGrid4.TextMatrix(l, 4)) = "" Then
      MsgBox "Blank remarks not allowed for zero expense!!!"
      ChkExpenseDetail = False
      Exit Function
   End If
Next
End Function

Public Sub GetEmailIDsforGodownViability()

Dim objFS As Scripting.FileSystemObject
Dim objFile As File
Dim objTS As TextStream
Dim lngFreefile As Long
Dim sFileLine As Variant
Set objFS = New Scripting.FileSystemObject
Set objFile = objFS.GetFile(App.Path & "\" & "EmailIDsforGodownViability.ini")
Set objTS = objFile.OpenAsTextStream(ForReading, TristateUseDefault)

sFileLine = objTS.ReadLine
'If LCase(CmbStatus.Text) = "approved" Then
'   sFileLine = objTS.ReadAll
'End If
If LCase(CmbStatus.Text) = "godown creation" Then
   sFileLine = objTS.ReadAll
End If
mApprovalmailIDs = mID(Trim(sFileLine), 3, Len(sFileLine))
objTS.Close
End Sub
Private Sub CalculateRentChange()
If mGodownTypeID = 1 Then
   mPrevAmt = 0
   tmp = "Select  MICE.InspectionExpenseID ,MICE.InspectionID ,MICE.InsExpTypeID ,MICE.QtySQ ,MICE.Amount ,MICE.Remarks, MIE.ExpenseType, MIE.CalculationOn, MICE.AmountPer, MIE.Formula, MIE.UserInputAmount, MIE.Type, MIE.UserInputCalc, MIE.ClearFlag, MIE.Remarks 'ExpenseRemarks'"
   tmp = tmp & " From Mst_InspectionExpenseDetail MICE"
   tmp = tmp & " Inner Join Mst_InspectionExpenseType MIE on MIE.InsExpTypeID = MICE.InsExpTypeID"
   tmp = tmp & " Where MICE.InspectionID = (Select isnull(max(InspectionID),0) from Mst_GodownInspection Where InspectionID < " & Val(TxtInspectionID) & ""
   tmp = tmp & " And Type='R' And GodownID =" & mGodownID & " ) And MIE.InsExpTypeID=1"  ' 1 for lease rent
   Call Tmp2Table
   If TmpRs2.RecordCount > 0 Then
      LblPRenewalRent.Caption = IIf(IsNull(TmpRs2!Amount), 0, TmpRs2!Amount)
      LblLeaseRent.Caption = "--"
      mPrevAmt = IIf(IsNull(TmpRs2!Amount), 0, TmpRs2!Amount)
   Else
      tmp = "select Top 1 TG.LAID, TGD.GodownID,TGD.SM_Prefix,TGD.RentAmount RentAmount"
      tmp = tmp & " From Txn_GodownLessorAgreeDetail TGD"
      tmp = tmp & " inner join Txn_LessorAgreement TG on TG.LAID=TGD.LAID"
      'tmp = tmp & " Where TG.ExpiryDate>='" & Format(Date, Gen_DatFormat) & "' And TG.Flag in ('A','P') and TGD.GodownID=" & mGodownID
      tmp = tmp & " Where TG.ExpiryDate<='" & Format(Date, Gen_DatFormat) & "' And TG.Flag in ('A') and TGD.GodownID=" & mGodownID
      'tmp = tmp & " Group by TGD.GodownID,TGD.SM_Prefix"
      tmp = tmp & " Order by TG.LAID Desc "
      Call Tmp2Table
      
      If TmpRs2.RecordCount > 0 Then
         LblLeaseRent.Caption = IIf(IsNull(TmpRs2!RentAmount), 0, TmpRs2!RentAmount)
         LblPRenewalRent.Caption = "--"
         mPrevAmt = IIf(IsNull(TmpRs2!RentAmount), 0, TmpRs2!RentAmount)
      End If
   End If
   If mPrevAmt > 0 Then
      TxtChangeInPer.Text = Format(((Val(TxtAmount) - Val(mPrevAmt)) / Val(mPrevAmt)) * 100, "#0.00")
      LblViewPrevRent.Enabled = True
   End If
End If
End Sub

Private Sub GetRentChange(Row_ As Variant)
mPrevAmt = 0
tmp = "Select  MICE.InspectionExpenseID ,MICE.InspectionID ,MICE.InsExpTypeID ,MICE.QtySQ ,MICE.Amount ,MICE.Remarks, MIE.ExpenseType, MIE.CalculationOn, MICE.AmountPer, MIE.Formula, MIE.UserInputAmount, MIE.Type, MIE.UserInputCalc, MIE.ClearFlag, MIE.Remarks 'ExpenseRemarks'"
tmp = tmp & " From Mst_InspectionExpenseDetail MICE"
tmp = tmp & " Inner Join Mst_InspectionExpenseType MIE on MIE.InsExpTypeID = MICE.InsExpTypeID"
tmp = tmp & " Where MICE.InspectionID = (Select isnull(max(InspectionID),0) from Mst_GodownInspection Where InspectionID < " & Val(TxtInspectionID) & ""
tmp = tmp & " And Type='R' And GodownID =" & mGodownID & " ) And MIE.InsExpTypeID=1"  ' 1 for lease rent
Call Tmp2Table
If TmpRs2.RecordCount > 0 Then
   mPrevAmt = IIf(IsNull(TmpRs2!Amount), 0, TmpRs2!Amount)
   MSHFlexGrid4.TextMatrix(Row_, 14) = mPrevAmt
Else
   tmp = "select Top 1 TG.LAID, TGD.GodownID,TGD.SM_Prefix,sum(TGD.RentAmount) RentAmount"
   tmp = tmp & " From Txn_GodownLessorAgreeDetail TGD"
   tmp = tmp & " inner join Txn_LessorAgreement TG on TG.LAID=TGD.LAID"
   'tmp = tmp & " Where TG.ExpiryDate>='" & Format(Date, Gen_DatFormat) & "' And TG.Flag in ('A','P') and TGD.GodownID=" & mGodownID
   tmp = tmp & " Where TG.ExpiryDate<='" & Format(Date, Gen_DatFormat) & "' And TG.Flag in ('A') and TGD.GodownID=" & mGodownID
   tmp = tmp & " Group by TG.LAID,TGD.GodownID,TGD.SM_Prefix"
   tmp = tmp & " Order by TG.LAID Desc "
   Call Tmp2Table
   
   If TmpRs2.RecordCount > 0 Then
      mPrevAmt = IIf(IsNull(TmpRs2!RentAmount), 0, TmpRs2!RentAmount)
      MSHFlexGrid4.TextMatrix(Row_, 15) = mPrevAmt
   End If
End If
If mPrevAmt > 0 Then
   MSHFlexGrid4.TextMatrix(Row_, 16) = Format(((Val(MSHFlexGrid4.TextMatrix(Row_, 3)) - Val(mPrevAmt)) / Val(mPrevAmt)) * 100, "#0.00")
   TxtSRentChange.Text = MSHFlexGrid4.TextMatrix(Row_, 16)
End If
End Sub


Public Sub GetGodownList()
If CmbLocation.Text <> "" And CmbCMP.Text <> "" Then
   tmp = " Select 'State' As State, MLT.LicenseTypeName,ML.LicenseTypeID,ML.LicenseID,ML.LicenseNo"
   tmp = tmp & " from Mst_License  ML"
   tmp = tmp & " Inner Join Txn_LicenseAllocationDetail TLAD on TLAD.LicenseID=ML.LicenseID"
   tmp = tmp & " Inner Join Mst_LicenseType MLT on ML.LicenseTypeID=MLT.LicenseTypeID"
   tmp = tmp & " where TLAD.SLGID=" & mStateID & " And  TLAD.Flag='S'"
   tmp = tmp & " Union"
   tmp = tmp & " Select 'Location' As Location, MLT.LicenseTypeName,ML.LicenseTypeID,ML.LicenseID,ML.LicenseNo"
   tmp = tmp & " from Mst_License  ML"
   tmp = tmp & " Inner Join Txn_LicenseAllocationDetail TLAD on TLAD.LicenseID=ML.LicenseID"
   tmp = tmp & " Inner Join Mst_LicenseType MLT on ML.LicenseTypeID=MLT.LicenseTypeID"
   tmp = tmp & " where TLAD.SLGID=" & mLocationID & " And TLAD.Flag='L'"
   tmp = tmp & " Union"
   tmp = tmp & " Select 'CMP' As CMP, MLT.LicenseTypeName,ML.LicenseTypeID,ML.LicenseID,ML.LicenseNo"
   tmp = tmp & " from Mst_License  ML"
   tmp = tmp & " Inner Join Txn_LicenseAllocationDetail TLAD on TLAD.LicenseID=ML.LicenseID"
   tmp = tmp & " Inner Join Mst_LicenseType MLT on ML.LicenseTypeID=MLT.LicenseTypeID"
   tmp = tmp & " where TLAD.SLGID=" & mCMPID & " And TLAD.Flag='C'"
   Call Tmp1Table
   If TmpRs1.RecordCount > 0 Then
      If MsgBox("Do you want to mapped this godown in license master", vbYesNo) = vbYes Then
         Frame5.Visible = True
         Call Grid_Head8
         With MSHFlexGrid8
              For I = 1 To TmpRs1.RecordCount
                  MSHFlexGrid8.Rows = MSHFlexGrid8.Rows + 1
                  
                  MSHFlexGrid8.Col = 0
                  MSHFlexGrid8.row = I
                  MSHFlexGrid8.CellFontName = "Wingdings"
                  MSHFlexGrid8.CellFontSize = 12
                  MSHFlexGrid8.CellAlignment = flexAlignCenterCenter
          
                  MSHFlexGrid8.TextMatrix(I, 0) = strUnChecked
                  MSHFlexGrid8.TextMatrix(I, 1) = TmpRs1!LicenseTypeName
                  MSHFlexGrid8.TextMatrix(I, 2) = TmpRs1!LicenseTypeID
                  MSHFlexGrid8.TextMatrix(I, 3) = TmpRs1!LicenseID
                  TmpRs1.MoveNext
              Next
         End With
      End If
   End If
End If
End Sub


Private Sub IndataInspectionAreaDetail()

Call ClearAreaDetail
Call Grid_Head9
Call TableMst_InspectionAreaDetail(" Where InspectionID = " & Val(TxtInspectionID) & " ")
If RsMst_InspectionAreaDetail.RecordCount > 0 Then
   With MSHFlexGrid9
      .Rows = .Rows + RsMst_InspectionAreaDetail.RecordCount
      For I = 2 To RsMst_InspectionAreaDetail.RecordCount + 1
         .TextMatrix(I, 0) = IIf(IsNull(RsMst_InspectionAreaDetail!InspectionAreaID), "", RsMst_InspectionAreaDetail!InspectionAreaID)
         .TextMatrix(I, 1) = IIf(IsNull(RsMst_InspectionAreaDetail!Length), "", RsMst_InspectionAreaDetail!Length)
         .TextMatrix(I, 2) = IIf(IsNull(RsMst_InspectionAreaDetail!Width), "", RsMst_InspectionAreaDetail!Width)
         .TextMatrix(I, 3) = IIf(IsNull(RsMst_InspectionAreaDetail!Height), "", RsMst_InspectionAreaDetail!Height)
         .TextMatrix(I, 4) = IIf(IsNull(RsMst_InspectionAreaDetail!Area), "", RsMst_InspectionAreaDetail!Area)
         .TextMatrix(0, 4) = Val(.TextMatrix(0, 4)) + Val(.TextMatrix(I, 4))
         RsMst_InspectionAreaDetail.MoveNext
      Next
   End With
End If
End Sub


Private Sub IndataGodownUtilizationDetail()

Call Grid_Head10

'TxtStartDate

'mMonthStartDate = Year(TxtStartDate.Value) & "-" & Month(TxtStartDate.Value) & "-01"
mMonthUtilizationDate = DateAdd("M", CDbl("-" & mBackMonths), CDate(TxtStartDate.Value))
mMonthUtilizationDate = Year(CDate(mMonthUtilizationDate)) & "-" & Month(CDate(mMonthUtilizationDate)) & "-01"


tmp = " Select YearMonth,GodownCapacity,Sum(BalWt) 'BalWt' From ("
tmp = tmp & " Select a.YearMonth,g.GodownCapacity,Sum(a.BalanceWeight) 'BalWt' from Txn_InvMOnGSLDetail a"
tmp = tmp & " Inner Join Mst_Godown g on a.GodownID = g.GodownID"
tmp = tmp & " Where a.TFlag in ('O','D') And a.YearMonth >= '" & mMonthUtilizationDate & "' And a.CMPID = '" & mCMPID & "' And (a.GodownID = " & Val(mGodownID) & " Or g.GodownNo = '" & CmbGodown.Text & "')"
tmp = tmp & " Group By a.YearMonth, g.GodownCapacity"
tmp = tmp & " Union All"
tmp = tmp & " Select a.YearMonth,g.GodownCapacity,-Sum(a.BalanceWeight) 'BalWt' from Txn_InvMOnGSLDetail a"
tmp = tmp & " Inner Join Mst_Godown g on a.GodownID = g.GodownID"
tmp = tmp & " Where a.TFlag in ('W') And a.YearMonth >= '" & mMonthUtilizationDate & "' And a.CMPID = '" & mCMPID & "' And (a.GodownID = " & Val(mGodownID) & " Or g.GodownNo = '" & CmbGodown.Text & "')"
tmp = tmp & " Group By a.YearMonth,g.GodownCapacity) A"
tmp = tmp & " Group By YearMonth, GodownCapacity"
tmp = tmp & " Order By YearMonth Desc"

Call Tmp7Table

If TmpRs7.RecordCount > 0 Then
   With MSHFlexGrid10
      .Rows = .Rows + TmpRs7.RecordCount
      For I = 2 To TmpRs7.RecordCount + 1
         .TextMatrix(I, 0) = I - 1
         .TextMatrix(I, 1) = IIf(IsNull(TmpRs7!YearMonth), "", Format(TmpRs7!YearMonth, "MMM-YYYY"))
         .TextMatrix(I, 2) = IIf(IsNull(TmpRs7!GodownCapacity), "", TmpRs7!GodownCapacity)
         .TextMatrix(I, 3) = IIf(IsNull(TmpRs7!BalWt), "", TmpRs7!BalWt)
         .TextMatrix(I, 4) = Format((Val(.TextMatrix(I, 3)) / Val(.TextMatrix(I, 2))) * 100, "#######.00")
         '.TextMatrix(0, 4) = Val(.TextMatrix(0, 4)) + Val(.TextMatrix(I, 4))
         TmpRs7.MoveNext
      Next
   End With
End If
End Sub

Private Sub IndataGodownPnLDetail()


Call Grid_Head11

mMonthUtilizationDate = DateAdd("M", CDbl("-" & mBackMonths), CDate(TxtStartDate.Value))
mMonthUtilizationDate = Year(CDate(mMonthUtilizationDate)) & "-" & Month(CDate(mMonthUtilizationDate)) & "-01"

tmp = " Select MonthYear,(NNNRIncomeAdhoc+NNNRIncome+ReservationIncomeAdhoc+ReservationIncome+SevenDaysIncome+DPIncomeAdHoc+FinalDPIncome+RejectedIncome+SpotIncomeAdHoc+OtherIncome+SCMIncome+CMUnbilledIncome+CMPendingIncome+CMApprovedIncome+LeftOverStock+PledgeIncome-PartiallyReversedIncome) 'TotIncome' ,"
tmp = tmp & " (LeaseRentPaid+LeaseRentFAG+LeaseRentDRC+EmpClaimFAG+EmpClaimAccrual+VendorExpenses+VendorExpensesAccrued+SecurityFeesPaid+SecurityFeesAccrual+FumigationFeesVendor+FumigationFeesPaid+FumigationFeesAccrual+FranchiseeFeesPaid+FranchiseeFeesCMG+FranchiseeFeesFAG+Salary+NCDEXCharges+LeaseRentAccrued+CoExpense+InsurancePremium+InterestCharges) 'TotExp',"
tmp = tmp & " 0 'ProfitOrLoss', (LeaseRentPaid+LeaseRentFAG+LeaseRentDRC+LeaseRentAccrued) 'Rent'"
tmp = tmp & " From Txn_MonGodownWiseProfitability a "
tmp = tmp & " Inner Join Mst_Godown g on a.GodownID = g.GodownID "
tmp = tmp & " Where MonthYear >= '" & mMonthUtilizationDate & "' And g.CMPID = " & mCMPID & " And (a.GodownID = " & Val(mGodownID) & " Or g.GodownNo = '" & CmbGodown.Text & "')"
tmp = tmp & " Order By MonthYear Desc"

Call Tmp7Table

If TmpRs7.RecordCount > 0 Then
   With MSHFlexGrid11
      .Rows = .Rows + TmpRs7.RecordCount
      For I = 2 To TmpRs7.RecordCount + 1
         .TextMatrix(I, 0) = I - 1
         .TextMatrix(I, 1) = IIf(IsNull(TmpRs7!MonthYear), "", Format(TmpRs7!MonthYear, "MMM-YYYY"))
         .TextMatrix(I, 2) = IIf(IsNull(TmpRs7!TotIncome), "", TmpRs7!TotIncome)
         .TextMatrix(I, 3) = IIf(IsNull(TmpRs7!TotExp), "", TmpRs7!TotExp)
         .TextMatrix(I, 4) = Format((Val(.TextMatrix(I, 2)) - Val(.TextMatrix(I, 3))), "#######.00")
         .TextMatrix(I, 5) = IIf(IsNull(TmpRs7!Rent), "", TmpRs7!Rent)
         .TextMatrix(0, 4) = Val(.TextMatrix(0, 4)) + Val(.TextMatrix(I, 4))
         TmpRs7.MoveNext
      Next
   End With
End If
End Sub

Private Sub GetPreviousInspectionAreaDetail()

Call ClearAreaDetail
Call Grid_Head9
Call TableMst_InspectionAreaDetail(" Where InspectionID in (Select Max(InspectionID) from Mst_GodownInspection Where  CMPID = " & mCMPID & " And  (GodownNo = '" & CmbGodown.Text & "' or GodownID = " & mGodownID & ") And Status not in ('R') ) ")
If RsMst_InspectionAreaDetail.RecordCount > 0 Then
   With MSHFlexGrid9
      .Rows = .Rows + RsMst_InspectionAreaDetail.RecordCount
      For I = 2 To RsMst_InspectionAreaDetail.RecordCount + 1
         .TextMatrix(I, 0) = ""
         .TextMatrix(I, 1) = IIf(IsNull(RsMst_InspectionAreaDetail!Length), "", RsMst_InspectionAreaDetail!Length)
         .TextMatrix(I, 2) = IIf(IsNull(RsMst_InspectionAreaDetail!Width), "", RsMst_InspectionAreaDetail!Width)
         .TextMatrix(I, 3) = IIf(IsNull(RsMst_InspectionAreaDetail!Height), "", RsMst_InspectionAreaDetail!Height)
         .TextMatrix(I, 4) = IIf(IsNull(RsMst_InspectionAreaDetail!Area), "", RsMst_InspectionAreaDetail!Area)
         .TextMatrix(0, 4) = Val(.TextMatrix(0, 4)) + Val(.TextMatrix(I, 4))
         RsMst_InspectionAreaDetail.MoveNext
      Next
   End With
End If
End Sub
