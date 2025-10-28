VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCDCTransaction 
   Caption         =   "Commodity Deposite Certificate"
   ClientHeight    =   9285
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14685
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9285
   ScaleWidth      =   14685
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   2175
      Left            =   720
      TabIndex        =   110
      Top             =   5040
      Width           =   14175
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
         Left            =   9960
         Sorted          =   -1  'True
         TabIndex        =   136
         Top             =   480
         Width           =   2565
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "Search"
         Height          =   350
         Left            =   10650
         TabIndex        =   135
         Top             =   1185
         Width           =   1000
      End
      Begin VB.TextBox TxtSCDCNo 
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
         Left            =   8040
         TabIndex        =   133
         Top             =   1200
         Width           =   2505
      End
      Begin VB.ComboBox CmbSCommodity 
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
         TabIndex        =   131
         Top             =   1200
         Width           =   3405
      End
      Begin VB.ComboBox CmbPreFix 
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
         Left            =   3480
         TabIndex        =   128
         Top             =   480
         Width           =   2685
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
         Left            =   6315
         Sorted          =   -1  'True
         TabIndex        =   127
         Top             =   480
         Width           =   3525
      End
      Begin VB.TextBox TxtSClientName 
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
         Left            =   3600
         TabIndex        =   125
         Top             =   1200
         Width           =   4305
      End
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   350
         Left            =   11640
         TabIndex        =   112
         Top             =   1185
         Width           =   1000
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1335
         Left            =   120
         TabIndex        =   111
         Top             =   1800
         Width           =   3150
         _ExtentX        =   5556
         _ExtentY        =   2355
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSComCtl2.DTPicker DtpSFromCDCDate 
         Height          =   375
         Left            =   120
         TabIndex        =   123
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   72155137
         CurrentDate     =   39884
      End
      Begin MSComCtl2.DTPicker DtpSToCDCDate 
         Height          =   375
         Left            =   1800
         TabIndex        =   138
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   72155137
         CurrentDate     =   39884
      End
      Begin VB.Label Label54 
         Caption         =   "To CDC Date"
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
         Left            =   1920
         TabIndex        =   139
         Top             =   150
         Width           =   1215
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
         Left            =   10920
         TabIndex        =   137
         Top             =   120
         Width           =   750
      End
      Begin VB.Label Label52 
         Caption         =   "CDC No"
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
         Left            =   8640
         TabIndex        =   134
         Top             =   900
         Width           =   1335
      End
      Begin VB.Label Label51 
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
         Height          =   300
         Left            =   1440
         TabIndex        =   132
         Top             =   885
         Width           =   1200
      End
      Begin VB.Label Label50 
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
         Left            =   4320
         TabIndex        =   130
         Top             =   150
         Width           =   1320
      End
      Begin VB.Label Label49 
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
         Left            =   7875
         TabIndex        =   129
         Top             =   150
         Width           =   1320
      End
      Begin VB.Label Label48 
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
         Height          =   255
         Left            =   5070
         TabIndex        =   126
         Top             =   915
         Width           =   1335
      End
      Begin VB.Label Label45 
         Caption         =   "From CDC Date"
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
         TabIndex        =   124
         Top             =   150
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture3 
      Height          =   975
      Left            =   120
      ScaleHeight     =   915
      ScaleWidth      =   14475
      TabIndex        =   47
      Top             =   7920
      Width           =   14535
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New "
         Height          =   375
         Left            =   360
         TabIndex        =   87
         Top             =   120
         Width           =   2295
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   375
         Left            =   2655
         TabIndex        =   32
         Top             =   120
         Width           =   2295
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   4920
         TabIndex        =   86
         Top             =   120
         Width           =   2325
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   7260
         TabIndex        =   85
         Top             =   120
         Width           =   2280
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "&Search"
         Height          =   375
         Left            =   9540
         TabIndex        =   84
         Top             =   120
         Width           =   2295
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   375
         Left            =   360
         TabIndex        =   33
         Top             =   495
         Width           =   2295
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   375
         Left            =   2655
         TabIndex        =   83
         Top             =   495
         Width           =   2295
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   375
         Left            =   4950
         TabIndex        =   82
         Top             =   495
         Width           =   2295
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   375
         Left            =   7245
         TabIndex        =   81
         Top             =   495
         Width           =   2295
      End
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   375
         Left            =   9540
         TabIndex        =   80
         Top             =   495
         Width           =   2295
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Print"
         Height          =   750
         Left            =   11835
         TabIndex        =   79
         Top             =   120
         Width           =   2535
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   7755
      Left            =   120
      ScaleHeight     =   7695
      ScaleWidth      =   14475
      TabIndex        =   35
      Top             =   120
      Width           =   14535
      Begin TabDlg.SSTab SSTab1 
         Height          =   4575
         Left            =   120
         TabIndex        =   8
         Top             =   3000
         Width           =   14175
         _ExtentX        =   25003
         _ExtentY        =   8070
         _Version        =   393216
         Tabs            =   4
         TabsPerRow      =   4
         TabHeight       =   520
         TabCaption(0)   =   "Godown Stack Detail"
         TabPicture(0)   =   "FrmCDCTransaction.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label11"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label23"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "LblStorageDays"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label27"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "StorageTo"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "StorageFrom"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "MSHFlexGrid2"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "CmdUpdate"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "CmdCancel"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).ControlCount=   9
         TabCaption(1)   =   "Assayer Details"
         TabPicture(1)   =   "FrmCDCTransaction.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame2"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Insurance Detail"
         TabPicture(2)   =   "FrmCDCTransaction.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Frame3"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "History Detail"
         TabPicture(3)   =   "FrmCDCTransaction.frx":0054
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "TxtCreated"
         Tab(3).Control(0).Enabled=   0   'False
         Tab(3).Control(1)=   "TxtUpdated"
         Tab(3).Control(1).Enabled=   0   'False
         Tab(3).Control(2)=   "Label46"
         Tab(3).Control(3)=   "Label44"
         Tab(3).ControlCount=   4
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
            TabIndex        =   120
            TabStop         =   0   'False
            Top             =   600
            Width           =   9330
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
            Left            =   -73635
            Locked          =   -1  'True
            TabIndex        =   119
            TabStop         =   0   'False
            Top             =   1365
            Width           =   9315
         End
         Begin VB.CommandButton CmdCancel 
            Caption         =   "Cancel"
            Height          =   420
            Left            =   12480
            TabIndex        =   91
            Top             =   435
            Width           =   1455
         End
         Begin VB.CommandButton CmdUpdate 
            Caption         =   "Update"
            Height          =   420
            Left            =   11040
            TabIndex        =   106
            Top             =   435
            Width           =   1455
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
            Height          =   3375
            Left            =   120
            TabIndex        =   34
            Top             =   960
            Width           =   13950
            _ExtentX        =   24606
            _ExtentY        =   5953
            _Version        =   393216
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin VB.Frame Frame2 
            Height          =   3900
            Left            =   -74880
            TabIndex        =   59
            Top             =   360
            Width           =   13935
            Begin VB.TextBox TxtGrainSize 
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
               MaxLength       =   30
               TabIndex        =   16
               Top             =   2700
               Width           =   1575
            End
            Begin VB.TextBox TxtPolarization 
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
               MaxLength       =   30
               TabIndex        =   15
               Top             =   2280
               Width           =   1575
            End
            Begin VB.CheckBox ChkAcceptFlag 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   240
               Left            =   4800
               TabIndex        =   28
               Top             =   3120
               Width           =   255
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
               Height          =   2940
               Left            =   10440
               MaxLength       =   255
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   29
               Top             =   600
               Width           =   3375
            End
            Begin VB.TextBox TxtOther 
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
               MaxLength       =   30
               TabIndex        =   17
               Top             =   3120
               Width           =   1575
            End
            Begin VB.TextBox TxtOrganic 
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
               Left            =   8280
               MaxLength       =   30
               TabIndex        =   23
               Top             =   1440
               Width           =   2055
            End
            Begin VB.TextBox TxtColour 
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
               MaxLength       =   30
               TabIndex        =   20
               Top             =   1860
               Width           =   1575
            End
            Begin VB.TextBox TxtStapleLength 
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
               Left            =   8280
               MaxLength       =   30
               TabIndex        =   22
               Top             =   1020
               Width           =   2055
            End
            Begin VB.TextBox TxtWax 
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
               Left            =   8280
               MaxLength       =   30
               TabIndex        =   24
               Top             =   1860
               Width           =   2055
            End
            Begin VB.TextBox TxtAgeing 
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
               MaxLength       =   30
               TabIndex        =   21
               Top             =   2280
               Width           =   1575
            End
            Begin VB.TextBox TxtDamaged 
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
               MaxLength       =   30
               TabIndex        =   13
               Top             =   1440
               Width           =   1575
            End
            Begin VB.TextBox TxtOil 
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
               MaxLength       =   30
               TabIndex        =   19
               Top             =   1440
               Width           =   1575
            End
            Begin VB.TextBox TxtForeign 
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
               MaxLength       =   30
               TabIndex        =   14
               Top             =   1845
               Width           =   1575
            End
            Begin VB.TextBox TxtMoisture 
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
               MaxLength       =   30
               TabIndex        =   12
               Top             =   1020
               Width           =   1575
            End
            Begin VB.TextBox TxtBrokan 
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
               MaxLength       =   30
               TabIndex        =   18
               Top             =   1020
               Width           =   1575
            End
            Begin VB.TextBox TxtAssayerNo 
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
               MaxLength       =   100
               TabIndex        =   9
               Top             =   195
               Width           =   4455
            End
            Begin VB.TextBox TxtGradeByAssayer 
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
               Left            =   8280
               MaxLength       =   30
               TabIndex        =   25
               Top             =   2280
               Width           =   2055
            End
            Begin VB.TextBox TxtAssayerName 
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
               MaxLength       =   150
               TabIndex        =   11
               Top             =   600
               Width           =   8415
            End
            Begin MSComCtl2.DTPicker ARDate 
               Height          =   375
               Left            =   8265
               TabIndex        =   10
               Top             =   180
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   661
               _Version        =   393216
               Format          =   72155137
               CurrentDate     =   39884
            End
            Begin MSComCtl2.DTPicker NED 
               Height          =   375
               Left            =   4800
               TabIndex        =   26
               Top             =   2685
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   661
               _Version        =   393216
               Format          =   72155137
               CurrentDate     =   39884
            End
            Begin MSComCtl2.DTPicker FED 
               Height          =   375
               Left            =   8280
               TabIndex        =   27
               Top             =   2685
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   661
               _Version        =   393216
               Format          =   72155137
               CurrentDate     =   39884
            End
            Begin VB.Label Label28 
               Caption         =   "Grain Size(mm)"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               TabIndex        =   95
               Top             =   2700
               Width           =   1695
            End
            Begin VB.Label Label25 
               Caption         =   "Polarization"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               TabIndex        =   94
               Top             =   2340
               Width           =   1335
            End
            Begin VB.Label Label5 
               Caption         =   "Accepted"
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
               Left            =   3720
               TabIndex        =   88
               Top             =   3120
               Width           =   1095
            End
            Begin VB.Label Label47 
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
               Height          =   15
               Left            =   3480
               TabIndex        =   78
               Top             =   2220
               Width           =   15
            End
            Begin VB.Label Label37 
               Caption         =   "Over All Quality"
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
               TabIndex        =   77
               Top             =   3120
               Width           =   1695
            End
            Begin VB.Label Label36 
               Caption         =   "Colour"
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
               Left            =   3720
               TabIndex        =   76
               Top             =   1860
               Width           =   735
            End
            Begin VB.Label Label35 
               Caption         =   "Organic - InOrganic"
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
               Left            =   6480
               TabIndex        =   75
               Top             =   1440
               Width           =   1935
            End
            Begin VB.Label Label34 
               Caption         =   "Staple Length"
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
               Left            =   6480
               TabIndex        =   74
               Top             =   1020
               Width           =   1455
            End
            Begin VB.Label Label33 
               Caption         =   "Wax-Honey"
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
               Left            =   6480
               TabIndex        =   73
               Top             =   1860
               Width           =   1695
            End
            Begin VB.Label Label32 
               Caption         =   "Ageing"
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
               Left            =   3720
               TabIndex        =   72
               Top             =   2280
               Width           =   735
            End
            Begin VB.Label Label24 
               Caption         =   "Damaged"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               TabIndex        =   71
               Top             =   1440
               Width           =   1095
            End
            Begin VB.Label Label22 
               Caption         =   "Oil Contain"
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
               Left            =   3720
               TabIndex        =   70
               Top             =   1440
               Width           =   975
            End
            Begin VB.Label Label21 
               Caption         =   "Foreign"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               TabIndex        =   69
               Top             =   1845
               Width           =   1455
            End
            Begin VB.Label Label20 
               Caption         =   "Brokan"
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
               Left            =   3720
               TabIndex        =   68
               Top             =   1020
               Width           =   855
            End
            Begin VB.Label Label17 
               Caption         =   "Moisture"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               TabIndex        =   67
               Top             =   1020
               Width           =   975
            End
            Begin VB.Label Label15 
               Caption         =   "AR No"
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
               TabIndex        =   66
               Top             =   195
               Width           =   1335
            End
            Begin VB.Label Label16 
               Caption         =   "Assayer Name"
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
               TabIndex        =   65
               Top             =   600
               Width           =   1575
            End
            Begin VB.Label Label29 
               Caption         =   "AR Date"
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
               Left            =   6720
               TabIndex        =   64
               Top             =   195
               Width           =   975
            End
            Begin VB.Label Label1 
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
               Height          =   255
               Left            =   11520
               TabIndex        =   63
               Top             =   360
               Width           =   1095
            End
            Begin VB.Label Label14 
               Caption         =   "Grade"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   6480
               TabIndex        =   62
               Top             =   2280
               Width           =   855
            End
            Begin VB.Label Label18 
               Caption         =   "NED"
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
               Left            =   3720
               TabIndex        =   61
               Top             =   2700
               Width           =   615
            End
            Begin VB.Label Label19 
               Caption         =   "FED"
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
               Left            =   6495
               TabIndex        =   60
               Top             =   2700
               Width           =   615
            End
         End
         Begin VB.Frame Frame3 
            Height          =   3975
            Left            =   -74880
            TabIndex        =   48
            Top             =   360
            Width           =   13815
            Begin VB.CommandButton CmdInsurance 
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
               Height          =   330
               Left            =   9120
               TabIndex        =   114
               Top             =   720
               Width           =   375
            End
            Begin VB.CommandButton CmdInsuranceCo 
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
               Height          =   330
               Left            =   9120
               TabIndex        =   113
               Top             =   255
               Width           =   375
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
               Height          =   375
               Left            =   7080
               Locked          =   -1  'True
               TabIndex        =   50
               TabStop         =   0   'False
               Top             =   1320
               Width           =   1815
            End
            Begin VB.TextBox TxtRisks 
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
               Height          =   1815
               Left            =   2160
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   49
               TabStop         =   0   'False
               Top             =   1800
               Width           =   11415
            End
            Begin VB.ComboBox CmbPolicyNo 
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
               Left            =   2160
               Sorted          =   -1  'True
               TabIndex        =   31
               Top             =   720
               Width           =   6645
            End
            Begin VB.ComboBox CmbInsuranceCo 
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
               Left            =   2160
               Sorted          =   -1  'True
               TabIndex        =   30
               Top             =   240
               Width           =   6645
            End
            Begin VB.TextBox TxtInsuranceCo 
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
               Left            =   2160
               TabIndex        =   51
               Top             =   240
               Width           =   6645
            End
            Begin VB.TextBox TxtPolicyNo 
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
               Left            =   2160
               TabIndex        =   52
               Top             =   720
               Width           =   6645
            End
            Begin MSComCtl2.DTPicker ValidFrom 
               Height          =   375
               Left            =   2160
               TabIndex        =   107
               TabStop         =   0   'False
               Top             =   1320
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   661
               _Version        =   393216
               Enabled         =   0   'False
               Format          =   72155137
               CurrentDate     =   39884
            End
            Begin MSComCtl2.DTPicker ValidTo 
               Height          =   375
               Left            =   4680
               TabIndex        =   108
               TabStop         =   0   'False
               Top             =   1320
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   661
               _Version        =   393216
               Enabled         =   0   'False
               Format          =   72155137
               CurrentDate     =   39884
            End
            Begin VB.Label Label43 
               Caption         =   "Risks"
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
               TabIndex        =   58
               Top             =   1800
               Width           =   1095
            End
            Begin VB.Label Label42 
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
               Height          =   375
               Left            =   6000
               TabIndex        =   57
               Top             =   1320
               Width           =   1695
            End
            Begin VB.Label Label41 
               Caption         =   "Valid From"
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
               TabIndex        =   56
               Top             =   1320
               Width           =   1695
            End
            Begin VB.Label Label40 
               Caption         =   "Valid To"
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
               Left            =   3720
               TabIndex        =   55
               Top             =   1320
               Width           =   975
            End
            Begin VB.Label Label39 
               Caption         =   "Policy No /  Cover note."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   120
               TabIndex        =   54
               Top             =   645
               Width           =   1815
            End
            Begin VB.Label Label38 
               Caption         =   "Insurance Co."
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
               TabIndex        =   53
               Top             =   285
               Width           =   1455
            End
         End
         Begin MSComCtl2.DTPicker StorageFrom 
            Height          =   375
            Left            =   1800
            TabIndex        =   103
            TabStop         =   0   'False
            Top             =   450
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   72155137
            CurrentDate     =   39884
         End
         Begin MSComCtl2.DTPicker StorageTo 
            Height          =   375
            Left            =   4920
            TabIndex        =   104
            Top             =   450
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   72155137
            CurrentDate     =   39884
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
            Left            =   -74880
            TabIndex        =   122
            Top             =   660
            Width           =   1695
         End
         Begin VB.Label Label44 
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
            Left            =   -74880
            TabIndex        =   121
            Top             =   1425
            Width           =   975
         End
         Begin VB.Label Label27 
            Caption         =   "Additional Storage Days :"
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
            Left            =   6975
            TabIndex        =   93
            Top             =   480
            Width           =   2310
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
            TabIndex        =   92
            Top             =   480
            Width           =   75
         End
         Begin VB.Label Label23 
            Caption         =   "Storage To"
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
            Left            =   3600
            TabIndex        =   90
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label Label11 
            Caption         =   "Storage From"
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
            TabIndex        =   89
            Top             =   480
            Width           =   1455
         End
      End
      Begin VB.PictureBox Picture2 
         Height          =   2790
         Left            =   120
         ScaleHeight     =   2730
         ScaleWidth      =   14115
         TabIndex        =   36
         Top             =   120
         Width           =   14175
         Begin VB.TextBox TxtPreFix 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
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
            Left            =   8280
            Locked          =   -1  'True
            TabIndex        =   117
            TabStop         =   0   'False
            Top             =   120
            Width           =   1095
         End
         Begin VB.TextBox TxtCommodityCondition 
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
            Left            =   11640
            TabIndex        =   7
            Top             =   2280
            Width           =   2235
         End
         Begin VB.TextBox TxtNoofDays 
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
            Left            =   13560
            Locked          =   -1  'True
            TabIndex        =   115
            TabStop         =   0   'False
            Top             =   2280
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox TxtStorageRate 
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
            Left            =   8280
            TabIndex        =   6
            Top             =   2280
            Width           =   1815
         End
         Begin VB.ComboBox CmbClientDP 
            Height          =   315
            Left            =   1920
            Sorted          =   -1  'True
            TabIndex        =   4
            Top             =   1800
            Visible         =   0   'False
            Width           =   12090
         End
         Begin VB.ComboBox CmbCMPID 
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
            Sorted          =   -1  'True
            TabIndex        =   2
            Top             =   525
            Width           =   4680
         End
         Begin VB.ComboBox CmbCommodityID 
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
            Sorted          =   -1  'True
            TabIndex        =   5
            Top             =   2235
            Width           =   4680
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
            Height          =   375
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   101
            TabStop         =   0   'False
            Top             =   945
            Width           =   4650
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
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   98
            Top             =   1800
            Width           =   4665
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
            Left            =   8280
            Locked          =   -1  'True
            TabIndex        =   96
            Top             =   1365
            Width           =   5730
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
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   45
            Top             =   2250
            Width           =   4680
         End
         Begin VB.TextBox TxtCMPAddress 
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
            Height          =   795
            Left            =   8280
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   525
            Width           =   5775
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
            Left            =   1920
            TabIndex        =   3
            Top             =   1365
            Width           =   4665
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
            Left            =   8280
            Locked          =   -1  'True
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   1785
            Width           =   5730
         End
         Begin VB.TextBox TxtCMPID 
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
            TabIndex        =   38
            Top             =   525
            Width           =   4650
         End
         Begin VB.TextBox TxtCDCID 
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
            TabIndex        =   0
            TabStop         =   0   'False
            Top             =   120
            Width           =   1575
         End
         Begin MSComCtl2.DTPicker CDCDate 
            Height          =   375
            Left            =   5040
            TabIndex        =   1
            Top             =   120
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            Format          =   72155137
            CurrentDate     =   39884
         End
         Begin VB.Label Label31 
            Caption         =   "Location ID"
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
            Left            =   6720
            TabIndex        =   118
            Top             =   180
            Width           =   1695
         End
         Begin VB.Label Label13 
            Caption         =   "Condition Of Commodity"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   10170
            TabIndex        =   116
            Top             =   2265
            Width           =   1455
         End
         Begin VB.Label Label4 
            Caption         =   "Storage Rate"
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
            Left            =   6720
            TabIndex        =   109
            Top             =   2340
            Width           =   1455
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
            Left            =   120
            TabIndex        =   105
            Top             =   1815
            Visible         =   0   'False
            Width           =   1740
         End
         Begin VB.Label Label2 
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
            Height          =   255
            Left            =   120
            TabIndex        =   102
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Label30 
            Caption         =   "Trading Member Name "
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
            Left            =   6690
            TabIndex        =   100
            Top             =   1770
            Width           =   1575
         End
         Begin VB.Label Label10 
            Caption         =   "Trading Member ID "
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
            TabIndex        =   99
            Top             =   1815
            Width           =   1800
         End
         Begin VB.Label Label8 
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
            Height          =   255
            Left            =   6720
            TabIndex        =   97
            Top             =   1425
            Width           =   1335
         End
         Begin VB.Label Label12 
            Caption         =   "WH Address (CMP)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   6720
            TabIndex        =   46
            Top             =   675
            Width           =   1575
         End
         Begin VB.Label Label7 
            Caption         =   "Commodity"
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
            Left            =   135
            TabIndex        =   44
            Top             =   2295
            Width           =   1215
         End
         Begin VB.Label Label26 
            Caption         =   "CDC Date"
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
            Left            =   3840
            TabIndex        =   42
            Top             =   120
            Width           =   1215
         End
         Begin VB.Label Label6 
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
            Left            =   120
            TabIndex        =   41
            Top             =   1425
            Width           =   975
         End
         Begin VB.Label Label3 
            Caption         =   "WH Name (CMP)"
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
            Left            =   120
            TabIndex        =   39
            Top             =   570
            Width           =   1695
         End
         Begin VB.Label Label9 
            Caption         =   "CDC ID"
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
            TabIndex        =   37
            Top             =   120
            Width           =   1095
         End
      End
   End
End
Attribute VB_Name = "frmCDCTransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Flag_EditDelete, Edit_Flg, mInsuranceCoID As String
Dim PolicyNo_, CNo As String
Dim mStorageDays, mStorageAmount As Integer
Dim mExchangeTypeID As Double
Dim mClientIDRow As Double
Dim mCMPID As Double
Dim mCommodityID As Double
Dim mSCMPName, mSLocationId, mSCommodityID, mSGodownID, mSClientID, mSClientIDRow As Double
Dim msSM_prefix As Variant
Dim mBackGSL As Variant
Private Sub AddCmd_Click()
Call TableMst_Config
TxtNoofDays = IIf(IsNull(RsMst_Config!SportStorageChargesDays), 0, RsMst_Config!SportStorageChargesDays) - 1
Picture2.Enabled = True
Edit_Flg = "A"
LockTxtBox (False)

mExchangeTypeID = GetExchangeID("SPOT")
Call GButtonLock(Me, False)
Call ClearData
SSTab1.Tab = 0
CmdUpdate.Enabled = False
CmdCancel.Enabled = False
CmdPrint.Enabled = False
CmdInsuranceCo.Visible = True
CmdInsurance.Visible = True
StorageFrom.Value = Date
TxtPreFix = ""
CDCDate.Enabled = True
ChkAcceptFlag.Value = 1
If RsTxn_CDC.RecordCount > 0 Then
   CDCDate.Value = Date
   If CDCDate.Enabled = True Then
      CDCDate.SetFocus
   End If
End If


TxtCreated = ""
TxtUpdated = ""

End Sub

Private Sub CDCDate_LostFocus()

If CDCDate.Value > Date Then
   MsgBox "Future Date is not Allowed !!!! "
   CDCDate.SetFocus
End If

Call GetGridData(mClientIDRow, mExchangeTypeID, mCMPID, mCommodityID)

End Sub

Private Sub CmbCMPID_GotFocus()
tmp = CmbCMPID.Text
tmp1 = ""
If Gen_WcCMP <> "" Then
   Call TableMst_CMP(" Where " & Gen_WcCMP & "")
Else
   Call TableMst_CMP
End If

CmbCMPID.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_CMP.RecordCount
    CmbCMPID.AddItem RsMst_CMP!CMPName
    RsMst_CMP.MoveNext
Next
CmbCMPID.Text = tmp
End Sub
'----
Private Sub CmbCMPID_LostFocus()
If CmbCMPID.Text = "" Then
   Exit Sub
End If
Call TableMst_CMP
RsMst_CMP.Find "CMPName = '" & CmbCMPID.Text & "'"
If RsMst_CMP.EOF Then
   MsgBox "Invalid selection "
   CmbCMPID.SetFocus
   Exit Sub
End If
mCMPID = RsMst_CMP!CMPID
TxtCMPAddress = RsMst_CMP!Address
TxtLocationID = GetLocationFromCMP(RsMst_CMP!CMPID)
TxtPreFix = GetPreFix(RsMst_CMP!CMPID, "P")
'mLocationID = RsMst_CMP!LocationID
Call GetGridData(mClientIDRow, mExchangeTypeID, mCMPID, mCommodityID)

End Sub

Private Sub CmbCommodityID_GotFocus()
tmp = CmbCommodityID.Text
Call TableMst_Commodity
CmbCommodityID.Clear
If RsMst_Commodity.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Commodity.RecordCount
    CmbCommodityID.AddItem RsMst_Commodity!Commodity
    RsMst_Commodity.MoveNext
Next
CmbCommodityID.Text = tmp
End Sub
'----
Private Sub CmbCommodityID_LostFocus()
If CmbCommodityID.Text = "" Then
   Exit Sub
End If
RsMst_Commodity.MoveFirst
RsMst_Commodity.Find "Commodity = '" & CmbCommodityID.Text & "'"
If RsMst_Commodity.EOF Then
   MsgBox "Invalid selection "
   CmbCommodityID.SetFocus
   Exit Sub
End If
mCommodityID = RsMst_Commodity!CommodityID
Call TableMst_CommodityDetail("where CommodityID = " & mCommodityID & " and ExchangeTypeID = " & mExchangeTypeID & "")
If RsMst_CommodityDetail.RecordCount > 0 Then
  mGunnyWt = RsMst_CommodityDetail!GunnyWt
  mStdDed = IIf(IsNull(RsMst_CommodityDetail!StdDeduction), "", RsMst_CommodityDetail!StdDeduction)
  mLotSize = IIf(IsNull(RsMst_CommodityDetail!LotSize), "", RsMst_CommodityDetail!LotSize)
  mVariation = IIf(IsNull(RsMst_CommodityDetail!Variation), "", RsMst_CommodityDetail!Variation)
  TxtStorageRate = IIf(IsNull(RsMst_CommodityDetail!ScRate), 0, RsMst_CommodityDetail!ScRate)
Else
  MsgBox "No Record Found For " & CmbCommodityID.Text
  CmbCommodityID.SetFocus
  Exit Sub
End If
Call GetGridData(mClientIDRow, mExchangeTypeID, mCMPID, mCommodityID)

End Sub
Private Sub GetGridData(mClientRowId_, mExchangeID_, mCMPID_, mCommodityID_)
Dim mStdDedLess As Double
Call Grid_Head
If IsNull(mClientRowId_) Or mClientRowId_ = 0 Then
   Exit Sub
End If
If IsNull(mExchangeID_) Or mExchangeID_ = 0 Then
   Exit Sub
End If
If IsNull(mCMPID_) Or mCMPID_ = 0 Then
   Exit Sub
End If
If IsNull(mCommodityID_) Or mCommodityID_ = 0 Then
   Exit Sub
End If
Dim Sdat As Variant
Sdat = Format(CDCDate.Value, Gen_DatFormat)
If TxtStorageRate.Text = "" Then Exit Sub


If Gen_DBType = "MDB" Then
   tmp = "SELECT Mst_Godown.GodownNo,Mst_Godown.Address, Mst_GSL.StackNo, Mst_GSL.LotNo, Mst_GSL.DepositeBags, "
   tmp = tmp & " Mst_GSL.DepositeWeight,Mst_GSL.ISINID,Mst_GSL.GSLID , Mst_GSL.CDTFNo, Mst_GSL.CDTFDate,Mst_GSL.BGSLID  "
   tmp = tmp & " FROM Mst_Godown INNER JOIN Mst_GSL ON Mst_Godown.GodownID = Mst_GSL.GodownID "
   tmp = tmp & " WHERE Mst_GSL.sm_prefix = '" & TxtPreFix & "' And  (Mst_GSL.ClientIDRow= " & mClientRowId_ & " AND Mst_GSL.CommodityID=" & mCommodityID_ & "  AND "
   tmp = tmp & " Mst_GSL.ExchangeTypeID=" & mExchangeID_ & " AND Mst_Godown.CMPID= " & mCMPID_ & " AND  "
   tmp = tmp & " Mst_GSL.CDTFDate <= #" & Sdat & "# And Mst_GSL.MadeupFlag = '0' )"
Else
   tmp = "SELECT Mst_Godown.GodownNo,Mst_Godown.Address, Mst_GSL.StackNo, Mst_GSL.LotNo, Mst_GSL.DepositeBags, "
   tmp = tmp & " Mst_GSL.DepositeWeight,Mst_GSL.ISINID,Mst_GSL.GSLID , Mst_GSL.CDTFNo, Mst_GSL.CDTFDate,Mst_GSL.BGSLID  "
   tmp = tmp & " FROM Mst_Godown INNER JOIN Mst_GSL ON Mst_Godown.GodownID = Mst_GSL.GodownID "
   tmp = tmp & " WHERE Mst_GSL.sm_prefix = '" & TxtPreFix & "' And  (Mst_GSL.ClientIDRow= " & mClientRowId_ & " AND Mst_GSL.CommodityID=" & mCommodityID_ & "  AND "
   tmp = tmp & " Mst_GSL.ExchangeTypeID=" & mExchangeID_ & " AND Mst_Godown.CMPID= " & mCMPID_ & " AND  "
   tmp = tmp & " Mst_GSL.CDTFDate <= '" & Sdat & "' And Mst_GSL.MadeupFlag = '0' )"
End If
'If SaveCmd.Enabled = False Then
'   tmp = tmp & " And (Mst_GSL.ISINID= '" & TxtISINID & "')"
'Else
'   tmp = tmp & " And (Mst_GSL.ISINID Is Null  OR Mst_GSL.ISINID= '' OR Mst_GSL.ISINID= '" & TxtISINID & "')"
'End If
  
Call TmpTable

If TmpRs.RecordCount <= 0 Then
   MsgBox "No Record Found For Grid!!!"
   Exit Sub
Else
   
End If
k = 0
For I = 0 To TmpRs.RecordCount - 1
  
  If TmpRs!ISINID = "" Or TmpRs!ISINID = "0" Or IsNull(TmpRs!ISINID) Then
     k = k + 1
     mBackGSL = TmpRs!GSLID
     If IsNull(TmpRs!BGSLID) = False Then
        Call GetBGSLID(I, TmpRs!GSLID)
     End If
     
     tmp = " Where  sm_prefix = '" & TxtPreFix & "' And  TxnType = 'D' And  CxTFNo = " & TmpRs!CdtfNo
     Call TableTxn_CxTF_Details(tmp)
     tmp = " Where   sm_prefix = '" & TxtPreFix & "' And  GSLID = " & mBackGSL '" & TmpRs!GSLID
     Call TableTxn_CxTF_GSL(tmp)
     
     'WHERE (((Txn_CxTF_GSL.GSLID)=6)) OR (((Txn_CxTF_GSL.CxTFNo)=8));
     
     MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
     MSHFlexGrid2.TextMatrix(k, 1) = TmpRs!Address
     MSHFlexGrid2.TextMatrix(k, 2) = TmpRs!GodownNo
     MSHFlexGrid2.TextMatrix(k, 3) = TmpRs!StackNo
     MSHFlexGrid2.TextMatrix(k, 4) = TmpRs!LotNo
     For j = 1 To RsTxn_CxTF_GSL.RecordCount
         
         MSHFlexGrid2.TextMatrix(k, 5) = Val(MSHFlexGrid2.TextMatrix(k, 5)) + RsTxn_CxTF_GSL!NoofBags
         MSHFlexGrid2.TextMatrix(k, 6) = Val(MSHFlexGrid2.TextMatrix(k, 6)) + RsTxn_CxTF_GSL!AdjNCMSLWeight  'in MT
         MSHFlexGrid2.TextMatrix(k, 7) = Val(MSHFlexGrid2.TextMatrix(k, 7)) + RsTxn_CxTF_GSL!AdjNCMSLWeight * 1000 'in Kg
         MSHFlexGrid2.TextMatrix(k, 8) = Val(MSHFlexGrid2.TextMatrix(k, 8)) + RsTxn_CxTF_GSL!GunnyWeight * 1000 'TmpRs!GunnyWeight
         MSHFlexGrid2.TextMatrix(k, 9) = Val(MSHFlexGrid2.TextMatrix(k, 9)) + IIf(IsNull(RsTxn_CxTF_GSL!StdDeduction), 0, RsTxn_CxTF_GSL!StdDeduction) * 1000 'TmpRs!StdDeduction
         MSHFlexGrid2.TextMatrix(k, 10) = Val(MSHFlexGrid2.TextMatrix(k, 10)) + RsTxn_CxTF_GSL!DematWeight * 1000  '  net weight in KGS
         
         If IsNull(RsTxn_CxTF_Details!Flag) Or RsTxn_CxTF_Details!Flag = "" Then
            MSHFlexGrid2.TextMatrix(k, 20) = "* - " & TmpRs!CdtfNo
         Else
            MSHFlexGrid2.TextMatrix(k, 20) = TmpRs!CdtfNo
         End If
         RsTxn_CxTF_GSL.MoveNext
     Next
     If RsTxn_CxTF_GSL.RecordCount > 0 Then
        RsTxn_CxTF_GSL.MoveLast
     End If
     
     MSHFlexGrid2.TextMatrix(k, 11) = TmpRs!CDTFDate 'TmpRs!StorageFrom
     MSHFlexGrid2.TextMatrix(k, 12) = TmpRs!CDTFDate + Val(TxtNoofDays) '7 'TmpRs!StorageTo
     MSHFlexGrid2.TextMatrix(k, 13) = "0" '(MSHFlexGrid1.TextMatrix(k, 12) - MSHFlexGrid1.TextMatrix(k, 11)) 'TmpRs!TotStorageDays
     MSHFlexGrid2.TextMatrix(k, 14) = Val(TxtStorageRate.Text)
     tmp = Val(MSHFlexGrid2.TextMatrix(k, 6)) * Val(TxtStorageRate.Text) * Val(MSHFlexGrid2.TextMatrix(k, 13))    'TmpRs!StorageAmount
     MSHFlexGrid2.TextMatrix(k, 15) = tmp
     MSHFlexGrid2.TextMatrix(k, 16) = IIf(IsNull(TmpRs!ISINID), "", TmpRs!ISINID)
       'MSHFlexGrid2.TextMatrix(k, 17) = TmpRs!CDCDATE
     MSHFlexGrid2.TextMatrix(k, 18) = "0" 'TmpRs!PrintFlag
     MSHFlexGrid2.TextMatrix(k, 19) = TmpRs!GSLID 'Val(mBackGSL) 'TmpRs!GSLID
     
     RsTxn_CxTF_GSL.Close
  
  End If
  TmpRs.MoveNext
Next
If MSHFlexGrid2.Rows < 3 Then
   MsgBox "No Record Found For Grid!!!!"
   Exit Sub
End If

End Sub
Private Sub CmdCancel_Click()
If MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 0) = "*" Then
  MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 0) = ""
'Else
  'MSHFlexGrid2.TextMatrix(MSHFlexGrid1.RowSel, 0) = ""
End If
End Sub

Private Sub CmdExcel_Click()
Call Gen_Create_Xls("Txn_CDC.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\Txn_CDC.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = I + 1 ' + 6
    For c = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(I, c), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(I, c)) Then
              oWS.Cells(mRow, c + 1) = Format(MSHFlexGrid1.TextMatrix(I, c), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(I, c)
           End If
        Else
           oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(I, c)
        End If
    Next
Next
oWB.Save
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub

Private Sub CmdInsurance_Click()
  FrmMstInsurance.Show
End Sub

Private Sub CmdInsuranceCo_Click()
  FrmMstInsuranceCo.Show
End Sub

Private Sub CmdPrintold_Click()
On Error Resume Next
RsTxn_CDC!PrintFlag = "P"
RsTxn_CDC.Update
'RsTxn_CxTF_Details!Flag = "P"
'RsTxn_CxTF_Details.Update
tmp = "Delete from PrintDetail "
Call TmpTable

tmp = "Delete from PrintMain "
Call TmpTable
tmp = "select * from PrintMain"
Call TmpTable
k = 0
For I = 1 To MSHFlexGrid2.Rows - 1
    k = k + 1
    If MSHFlexGrid2.TextMatrix(k, 16) = "" Then Exit For
    TmpRs.AddNew
    TmpRs!ID = MSHFlexGrid2.TextMatrix(k, 16)
    TmpRs!F1 = TxtLocationID & "," & TxtCMPAddress
    TmpRs!F2 = TxtCMPID
    TmpRs!F3 = MSHFlexGrid2.TextMatrix(k, 4)
    TmpRs!F4 = TxtClientID & "-" & TxtClientName
    'TmpRs!F5 = TxtClientName
    TmpRs!F6 = TxtCommodity
    TmpRs!F7 = MSHFlexGrid2.TextMatrix(k, 5)
    TmpRs!F8 = MSHFlexGrid2.TextMatrix(k, 10)
    TmpRs!F9 = TxtAssayerNo & "-" & TxtAssayerName
    TmpRs!F10 = TxtGradeByAssayer
    TmpRs!F11 = FED.Value
    TmpRs!F12 = TxtRisks
    TmpRs!F13 = TxtInsuranceCo
    TmpRs!F14 = TxtAmount
    TmpRs!F15 = MSHFlexGrid2.TextMatrix(k, 11)
    TmpRs!F16 = MSHFlexGrid2.TextMatrix(k, 12)
    TmpRs!F17 = TxtCommodityCondition
    TmpRs.Update
Next
TmpRs.Close
cn.Close
If DataEnvironment1.rsCommand1.State <> 0 Then
  DataEnvironment1.rsCommand1.Close
End If

Rpt_CDC.Refresh
Rpt_CDC.Width = 15200
Rpt_CDC.Show

End Sub

Private Sub CmdSearch_Click()
Dim wc1 As Variant

Call Grid_Head1

If DtpSFromCDCDate.Value > DtpSToCDCDate.Value Then
   MsgBox "'From Date' must be less than 'To Date'!!! "
   DtpSFromCDCDate.SetFocus
   Exit Sub
End If


tmp = "Select  TC.CDCID,TCG.CDCNo,TC.CDCDate,ML.LocationName,MC.CMPName,MG.GodownNo,GSL.StackNo,"
tmp = tmp & " GSL.LotNo,TCG.StorageFrom,TCG.StorageTo,TCG.StorageRate,TCG.StorageAmount, MCL.ClientName, MCm.Commodity, TC.Flag, TC.PrintFlag,TC.G_UID"
tmp = tmp & " from Txn_CDC TC"
tmp = tmp & " inner join Mst_CMP MC on TC.CMPID=MC.CMPID"
tmp = tmp & " inner join Mst_Location ML on MC.LocationID=ML.LocationID"
tmp = tmp & " inner join Mst_Commodity MCm on TC.CommodityID=MCm.CommodityID"
tmp = tmp & " inner join Mst_Client MCl on TC.ClientIDRow=MCl.ClientIDRow"
tmp = tmp & " inner join Txn_CDC_GSL_Detail TCG on TC.CDCID=TCG.CDCID And TC.SM_Prefix=TCG.SM_Prefix"
tmp = tmp & " inner join Mst_GSL GSL on TCG.GSLID=GSL.GSLID And TCG.SM_Prefix=GSL.SM_Prefix"
tmp = tmp & " inner join Mst_Godown MG on GSL.GodownID=MG.GodownID"

If Gen_DBType = "MDB" Then
   wc1 = " Where TC.CDCDate Between #" & Format(DtpSFromCDCDate.Value, Gen_DatFormat) & "# And #" & Format(DtpSToCDCDate.Value, Gen_DatFormat) & "#"
Else
   wc1 = " Where TC.CDCDate Between '" & Format(DtpSFromCDCDate.Value, Gen_DatFormat) & "' And '" & Format(DtpSToCDCDate.Value, Gen_DatFormat) & "'"
End If
If CmbPreFix.Text <> "" Then
   wc1 = wc1 & " And ML.LocationName='" & CmbPreFix.Text & "'"
Else
   If Gen_WcLocation <> "" Then
      wc1 = wc1 & " And  ML." & Gen_WcLocation
   End If
End If

If CmbSCmp.Text <> "" Then
   wc1 = wc1 & " And MC.CMPName='" & CmbSCmp.Text & "'"
Else
   If Gen_WcCMP <> "" Then
      wc1 = wc1 & " And  MC." & Gen_WcCMP
   End If
End If

If CmbSGodown.Text <> "" Then
   wc1 = wc1 & " And MG.GodownID='" & mSGodownID & "'"
End If

If CmbSCommodity.Text <> "" Then
   wc1 = wc1 & " And MCm.Commodity='" & CmbSCommodity.Text & "'"
End If

If TxtSClientName.Text <> "" Then
   wc1 = wc1 & " And MCL.ClientName Like '" & TxtSClientName.Text & "%'"
End If

If TxtSCDCNo.Text <> "" Then
   wc1 = wc1 & " And TCG.CDCNo = " & Val(TxtSCDCNo.Text) & ""
End If

tmp = tmp & wc1

tmp = tmp & " order By TC.CDCID,TC.SM_Prefix"

Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   MSHFlexGrid1.Rows = TmpRs1.RecordCount + 2
   For I = 1 To TmpRs1.RecordCount
       For c = 0 To MSHFlexGrid1.Cols - 1
           MSHFlexGrid1.TextMatrix(I, c) = IIf(IsNull(TmpRs1.Fields(c)), "", TmpRs1.Fields(c))
       Next
       TmpRs1.MoveNext
   Next
End If

End Sub

Private Sub CmdUpdate_Click()
If StorageFrom.Value > StorageTo.Value Then
   MsgBox "'Storage To Date' Must Be Greater Than 'Storage From Date' Always!!!"
   StorageTo.SetFocus
   Exit Sub
End If

If MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 1) = "" Then
   CmdUpdate.Enabled = False
   CmdCancel.Enabled = False
   Exit Sub
End If
Call Fix_StartoGrid(True)
StorageTo.Enabled = False
CmdUpdate.Enabled = False
CmdCancel.Enabled = False
End Sub

Sub Fix_StartoGrid(ans As Boolean)
If ans Then
   MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 0) = "*"
   MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 11) = StorageFrom.Value
   MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 12) = StorageTo.Value
   MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 13) = (StorageTo.Value - StorageFrom.Value) - Val(TxtNoofDays) '7  'mStorageDays 'CDate(TxtStorageTo) - CDate(TxtStoragefrom)
   MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 15) = MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 6) * MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 14) * (MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 13))    'TmpRs!StorageAmount
Else
   MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 0) = ""
   MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 11) = StorageFrom.Value
   MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 12) = StorageFrom.Value + Val(TxtNoofDays) '7
   MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 13) = (StorageTo.Value - StorageFrom.Value) - Val(TxtNoofDays) '7  'mStorageDays 'CDate(TxtStorageTo) - CDate(TxtStoragefrom)
   MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 15) = MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 6) * MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 14) * (MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 13))    'TmpRs!StorageAmount
End If

End Sub

Private Sub DeleteCmd_Click()

If Flag_EditDelete = "1" Then
   MsgBox "Payout completed; cannot edit/delete !!!! "
   Exit Sub
End If

If RsTxn_CDC!PrintFlag = "P" Then
  MsgBox "You can't Delete this CDC,Since printout is taken !!!! "
  Exit Sub
End If

Dim ans As Variant
On Error GoTo msgError
ans = MsgBox("Do you really want to DELETE this record???", vbYesNo)
If ans = vbYes Then
Call DeleteChildTableDetail
   tmp = "select * from Txn_CDC where sm_prefix = '" & TxtPreFix & "' And CDCID = " & TxtCDCID
   Call TmpTable
   TmpRs.Close
   
   RsTxn_CDC.Delete
   RsTxn_CDC.Update
   If RsTxn_CDC.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsTxn_CDC.MoveNext
   If RsTxn_CDC.EOF() Then
      RsTxn_CDC.MoveLast
   End If
   Call Indata
End If
Exit Sub
msgError:
MsgBox "Child record Present "
RsTxn_CDC.CancelUpdate
End Sub

Sub DeleteChildTableDetail()
Call TableMst_GSL("where sm_prefix = '" & TxtPreFix & "'")
tmp = "select * from Txn_CDC_GSL_Detail  where sm_prefix = '" & TxtPreFix & "' And CDCID = " & TxtCDCID
Call TmpTable
For I = 1 To TmpRs.RecordCount
    RsMst_GSL.MoveFirst
    RsMst_GSL.Find "GSLID = " & TmpRs!GSLID
    If Not RsMst_GSL.EOF Then
       'RsMst_GSL!DepositeBags = IIf(IsNull(RsMst_GSL!DepositeBags), 0, RsMst_GSL!DepositeBags) - TmpRs!NoOfBags
       'RsMst_GSL!DepositeWeight = IIf(IsNull(RsMst_GSL!DepositeWeight), 0, RsMst_GSL!DepositeWeight) - TmpRs!DematWeight
       'RsMst_GSL!BalanceBags = IIf(IsNull(RsMst_GSL!BalanceBags), 0, RsMst_GSL!BalanceBags) - TmpRs!NoOfBags
       'RsMst_GSL!BalanceWeight = IIf(IsNull(RsMst_GSL!BalanceWeight), 0, RsMst_GSL!BalanceWeight) - TmpRs!DematWeight
       If RsMst_GSL!DepositeWeight = 0 Then
          RsMst_GSL!Flag = "B"
          RsMst_GSL!ClientIDRow = 0
          RsMst_GSL!CommodityID = Null
          RsMst_GSL!ExchangeTypeID = Null
          RsMst_GSL!ISINID = "0"
       Else
          RsMst_GSL!Flag = "O"
          RsMst_GSL!ISINID = "0"
       End If
       RsMst_GSL.Update
    End If
    TmpRs.MoveNext
Next
tmp = "Delete from Txn_CDC_Assayer_Detail where sm_prefix = '" & TxtPreFix & "' And CDCID = " & TxtCDCID
Call TmpTable
tmp = "Delete from Txn_CDC_GSL_Detail where sm_prefix = '" & TxtPreFix & "' And CDCID = " & TxtCDCID
Call TmpTable
tmp = "Delete from Txn_CDC_Insurance_Detail where sm_prefix = '" & TxtPreFix & "' And CDCID = " & TxtCDCID
Call TmpTable
tmp = "select * from Txn_CDC_GSL_Detail where sm_prefix = '" & TxtPreFix & "' And CDCID = " & TxtCDCID
Call TmpTable
TmpRs.Close
End Sub
Private Sub EditCmd_Click()
  Edit_Flg = "E"
  If Flag_EditDelete = "1" Then
    MsgBox "Payout completed; cannot edit/delete"
    Exit Sub
  End If
  If RsTxn_CDC!PrintFlag = "P" Then
    MsgBox "You can't Edit this CDC,Since printout is taken !!!! "
    Exit Sub
  End If
  
  Call TableMst_Config
  TxtNoofDays = IIf(IsNull(RsMst_Config!SportStorageChargesDays), 0, RsMst_Config!SportStorageChargesDays) - 1
  
  Call LockTxtBox(False)
  CmbCMPID.Visible = False
  CmbCommodityID.Visible = False
  TxtCMPID.Visible = True
  TxtCommodity.Visible = True
  Picture2.Enabled = False
  TxtAssayerName.Visible = True
  CmbInsuranceCo.Text = TxtInsuranceCo.Text
  CmbPolicyNo = TxtPolicyNo.Text
  CmdInsuranceCo.Visible = True
  CmdInsurance.Visible = True
  Call GButtonLock(Me, False)
End Sub

Private Sub ExitCmd_Click()
   Unload Me
End Sub

Private Sub FirstCmd_Click()
RsTxn_CDC.MoveFirst
Call Indata
'TmpRs.MoveFirst
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Call GButtonLockAD(Me)
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If

LockTxtBox (True)
Frame0.Visible = False
Call Grid_Head
If Gen_WcCMP <> "" Then
   Call TableTxn_CDC(" Where Flag = 'O' And " & Gen_WcCMP)
Else
   Call TableTxn_CDC(" Where Flag = 'O'")
End If
Call Grid_Head1
DtpSFromCDCDate.Value = Date - 365
DtpSToCDCDate.Value = Date
CmbPreFix.Text = ""
CmbSCmp.Text = ""
CmbSCommodity.Text = ""
TxtSClientName.Text = ""
TxtSCDCNo.Text = ""

'SSTab1.Tab = 0
If RsTxn_CDC.RecordCount = 0 Then
  Call AddCmd_Click
  Exit Sub
End If

Call GButtonLock(Me, True)
Call Indata

End Sub

Private Sub Grid_Head()
MSHFlexGrid2.Clear
MSHFlexGrid2.Rows = 2
MSHFlexGrid2.Clear
MSHFlexGrid2.Cols = 21
MSHFlexGrid2.Font.Size = 11
MSHFlexGrid2.FixedRows = 1

MSHFlexGrid2.TextMatrix(0, 0) = ""
MSHFlexGrid2.TextMatrix(0, 1) = "Godown Address"
MSHFlexGrid2.TextMatrix(0, 2) = "Godown No"
MSHFlexGrid2.TextMatrix(0, 3) = "Stack No"
MSHFlexGrid2.TextMatrix(0, 4) = "Lot No"
MSHFlexGrid2.TextMatrix(0, 5) = "No of Bags"
MSHFlexGrid2.TextMatrix(0, 6) = "Lot Wise Quantity in MT (Gross Weight)"
MSHFlexGrid2.TextMatrix(0, 7) = "Lot Wise Quantity in KGs (Gross Weight)"
MSHFlexGrid2.TextMatrix(0, 8) = "Gunny Weight"
MSHFlexGrid2.TextMatrix(0, 9) = "Std Deduction"
MSHFlexGrid2.TextMatrix(0, 10) = "Net Weight"
MSHFlexGrid2.TextMatrix(0, 11) = "Storage From"
MSHFlexGrid2.TextMatrix(0, 12) = "Storage To"
MSHFlexGrid2.TextMatrix(0, 13) = "Tot. Storage Days"
MSHFlexGrid2.TextMatrix(0, 14) = "Storage Rate"
MSHFlexGrid2.TextMatrix(0, 15) = "Storage Amount"
MSHFlexGrid2.TextMatrix(0, 16) = "CDC-No"
MSHFlexGrid2.TextMatrix(0, 17) = "CDC-DATE"
MSHFlexGrid2.TextMatrix(0, 18) = "Print Flag"
MSHFlexGrid2.TextMatrix(0, 19) = "GSLID"
MSHFlexGrid2.TextMatrix(0, 20) = "CDTF No"


MSHFlexGrid2.ColWidth(0) = 500
MSHFlexGrid2.ColWidth(1) = 2600
MSHFlexGrid2.ColWidth(2) = 900
MSHFlexGrid2.ColWidth(3) = 800
MSHFlexGrid2.ColWidth(4) = 800
MSHFlexGrid2.ColWidth(5) = 900
MSHFlexGrid2.ColWidth(6) = 1200
MSHFlexGrid2.ColWidth(7) = 1200
MSHFlexGrid2.ColWidth(8) = 800
MSHFlexGrid2.ColWidth(9) = 1200
MSHFlexGrid2.ColWidth(10) = 1100
MSHFlexGrid2.ColWidth(11) = 1100
MSHFlexGrid2.ColWidth(12) = 1100
MSHFlexGrid2.ColWidth(13) = 1200
MSHFlexGrid2.ColWidth(14) = 800
MSHFlexGrid2.ColWidth(15) = 800
MSHFlexGrid2.ColWidth(16) = 1000
MSHFlexGrid2.ColWidth(17) = 1100
MSHFlexGrid2.ColWidth(18) = 0
MSHFlexGrid2.ColWidth(19) = 0
MSHFlexGrid2.ColWidth(20) = 0
MSHFlexGrid2.RowHeight(0) = 1000
CmdUpdate.Enabled = False
CmdCancel.Enabled = False
    
End Sub
Private Sub Grid_Head1()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Clear
MSHFlexGrid1.Cols = 17
MSHFlexGrid1.Font.Size = 11
MSHFlexGrid1.FixedRows = 1

MSHFlexGrid1.TextMatrix(0, 0) = "CDCID"
MSHFlexGrid1.TextMatrix(0, 1) = "CDC No"
MSHFlexGrid1.TextMatrix(0, 2) = "CDCDate"
MSHFlexGrid1.TextMatrix(0, 3) = "Location"
MSHFlexGrid1.TextMatrix(0, 4) = "CMP Name"
MSHFlexGrid1.TextMatrix(0, 5) = "Godown No"
MSHFlexGrid1.TextMatrix(0, 6) = "Stack No"
MSHFlexGrid1.TextMatrix(0, 7) = "Lot No"
MSHFlexGrid1.TextMatrix(0, 8) = "Storage From"
MSHFlexGrid1.TextMatrix(0, 9) = "Storage To"
MSHFlexGrid1.TextMatrix(0, 10) = "Storage Rate"
MSHFlexGrid1.TextMatrix(0, 11) = "Storage Amt"

MSHFlexGrid1.TextMatrix(0, 12) = "Client Name"
MSHFlexGrid1.TextMatrix(0, 13) = "Commodity Name"
MSHFlexGrid1.TextMatrix(0, 14) = "Flag"
MSHFlexGrid1.TextMatrix(0, 15) = "Print Flag"
MSHFlexGrid1.TextMatrix(0, 16) = "GUID"

MSHFlexGrid1.ColWidth(0) = 800
MSHFlexGrid1.ColWidth(1) = 1000
MSHFlexGrid1.ColWidth(2) = 1000
MSHFlexGrid1.ColWidth(3) = 2200
MSHFlexGrid1.ColWidth(4) = 2400
MSHFlexGrid1.ColWidth(5) = 1200
MSHFlexGrid1.ColWidth(6) = 1200
MSHFlexGrid1.ColWidth(7) = 1200
MSHFlexGrid1.ColWidth(8) = 1400
MSHFlexGrid1.ColWidth(9) = 1200
MSHFlexGrid1.ColWidth(10) = 1400
MSHFlexGrid1.ColWidth(11) = 1200
MSHFlexGrid1.ColWidth(12) = 2600
MSHFlexGrid1.ColWidth(13) = 2600
MSHFlexGrid1.ColWidth(14) = 800
MSHFlexGrid1.ColWidth(15) = 800
MSHFlexGrid1.ColWidth(16) = 0
    'MSHFlexGrid1.RowHeight(0) = 1000
    
    
End Sub

Private Sub LastCmd_Click()
RsTxn_CDC.MoveLast
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
RsTxn_CDC.MoveFirst
RsTxn_CDC.Find "G_UID = '" & MSHFlexGrid1.TextMatrix(I, 16) & "'"
If Not RsTxn_CDC.EOF Then
   Call Indata
   Frame0.Visible = False
End If
End Sub

Private Sub MSHFlexGrid2_Click()
If SaveCmd.Enabled = False Then Exit Sub
If MSHFlexGrid2.RowSel <> "0" Then
   If MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 2) = "" Then Exit Sub
   If MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 19) <> "" Then
      tmp = Mid(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 20), 1, 1)
      If tmp = "*" Then
         MsgBox "CDTF Printout is not Generated You can't select this Record !!!!! "
         Exit Sub
      End If
     
      LblStorageDays.Refresh
      StorageFrom.Value = MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 11)
      StorageTo.Value = MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 12)
      LblStorageDays.Caption = (MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 13))
      If MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 0) = "" Then
         Call Fix_StartoGrid(True)
      Else
         Call Fix_StartoGrid(False)
      End If
      StorageTo.Enabled = True
      StorageTo.SetFocus
   Else
      CmdUpdate.Enabled = False
      CmdCancel.Enabled = False
   End If
End If
End Sub
Private Function LockTxtBox(Yn As Boolean)
Dim Ny As Boolean
    
Ny = IIf(Yn, False, True)
TxtCDCID.Enabled = Ny

CmbCMPID.Visible = Ny
CmbCommodityID.Visible = Ny
TxtCMPID.Visible = Yn
TxtCommodity.Visible = Yn
'Tab Controls
'CmbAssayerName.Visible = Ny
CmbInsuranceCo.Visible = Ny
CmbPolicyNo.Visible = Ny
TxtClientID.Locked = Yn
TxtStorageRate.Locked = True
TxtCommodityCondition.Locked = Yn
TxtAssayerName.Visible = Ny
TxtAssayerName.Locked = Yn
TxtInsuranceCo.Visible = Yn
TxtPolicyNo.Visible = Yn
TxtCDCID.Enabled = Ny
CDCDate.Enabled = Ny
'StorageFrom.Enabled = Ny
'StorageTo.Enabled = Ny
'MSHFlexGrid1.Enabled = Yn
TxtAssayerNo.Locked = Yn
TxtMoisture.Locked = Yn
TxtForeign.Locked = Yn
TxtOil.Locked = Yn
TxtOrganic.Locked = Yn
TxtPolarization.Locked = Yn
TxtGrainSize.Locked = Yn
TxtGradeByAssayer.Locked = Yn
TxtOther.Locked = Yn
TxtColour.Locked = Yn
TxtBrokan.Locked = Yn
TxtAgeing.Locked = Yn
TxtDamaged.Locked = Yn
TxtStapleLength.Locked = Yn
TxtWax.Locked = Yn
ARDate.Enabled = Ny
NED.Enabled = Ny
FED.Enabled = Ny
TxtRemark.Locked = Yn
ChkAcceptFlag.Enabled = Ny
TxtInsuranceCo.Locked = Yn
TxtPolicyNo.Locked = Yn
    
End Function

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsTxn_CDC.MoveNext
If RsTxn_CDC.EOF Then
   RsTxn_CDC.MoveLast
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
RsTxn_CDC.MovePrevious

If RsTxn_CDC.BOF Then
   RsTxn_CDC.MoveFirst
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
If Edit_Flg = "A" Then
    If CmbCMPID.Text = "" Then
       MsgBox "Blank CMP Not Allowed!!!"
       CmbCMPID.SetFocus
       Exit Sub
    End If
    
    If TxtClientID.Text = "" Then
       MsgBox "Blank ClientID Not Allowed!!!"
       TxtClientID.SetFocus
       Exit Sub
    End If
    
    If CmbCommodityID.Text = "" Then
       MsgBox "Blank Commodity Not Allowed!!!"
       CmbCommodityID.SetFocus
       Exit Sub
    End If
End If

If TxtStorageRate.Text = "" Then
   MsgBox "Blank Storage Rate No Allowed!!!"
   TxtStorageRate.SetFocus
   Exit Sub
End If
If TxtCommodityCondition = "" Then
   MsgBox "Blank Condition Of Commodity Not Allowed!!!"
   TxtCommodityCondition.SetFocus
   Exit Sub
End If
If MSHFlexGrid2.Rows < 3 Then
   MsgBox "Atleast One GSL Detail Should Be In Grid"
   CmbCommodityID.SetFocus
   Exit Sub
End If

If CheckSelection() = True Then
   MsgBox "Please Select Atleast 1 GSL For CDC !!! "
   MSHFlexGrid2.SetFocus
   Exit Sub
End If

'If TxtAssayerNo.Text = "" Then
'   MsgBox "AR No Cant Left Balnk!!!!"
'   TxtAssayerNo.SetFocus
'   Exit Sub
'End If
'If ARDate.Value > CDCDate.Value Then
'   MsgBox "AR Date Must be Less than CDC Date"
'   Exit Sub
'End If
'
'If TxtAssayerName.Text = "" Then
'   MsgBox "Assayer Name Cant Left Blank!!!"
'   TxtAssayerName.SetFocus
'   Exit Sub
'End If
'
'If TxtMoisture = "" Then
'    MsgBox "Blank Moisture Not Allowed !!!! "
'    TxtMoisture.SetFocus
'    Exit Sub
' End If
'
'If TxtDamaged = "" Then
'   MsgBox "Blank Damaged Not Allowed !!!! "
'   TxtDamaged.SetFocus
'   Exit Sub
'End If
'
'If TxtForeign = "" Then
'   MsgBox "Blank Foreign Not Allowed !!!! "
'   TxtForeign.SetFocus
'   Exit Sub
'End If
'
'If TxtPolarization.Text = "" Then
'   MsgBox "Blank Polarization Not Allowed!!!"
'   TxtPolarization.SetFocus
'   Exit Sub
'End If
'
'If TxtGrainSize.Text = "" Then
'   MsgBox "Blank Gran Size Not Allowed!!!"
'   TxtGrainSize.SetFocus
'   Exit Sub
'End If
'
'If TxtGradeByAssayer.Text = "" Then
'   MsgBox "Blank Grade Not Allowed!!!"
'   TxtGradeByAssayer.SetFocus
'   Exit Sub
'End If
'
'If TxtOther.Text = "" Then
'   MsgBox "Blank OverAll Quality Not Allowed!!!"
'   TxtOther.SetFocus
'   Exit Sub
'End If
'
'If NED.Value > FED.Value Then
'   MsgBox "NED Must Be Less Than FED Always!!!"
'   FED.SetFocus
'   Exit Sub
'End If
'
'
'If NED.Value < ARDate.Value Then
'   MsgBox "NED Must Be Greater than AR Date Always!!!"
'   NED.SetFocus
'   Exit Sub
'End If

'If ChkAcceptFlag.Value = 0 Then
'   MsgBox "Please Mark The Acceptance of Report"
'   ChkAcceptFlag.SetFocus
'   Exit Sub
'End If
 
'If TxtRemark.Text = "" Then
'   MsgBox "Please Give Some Remarks!!!"
'   TxtRemark.SetFocus
'   Exit Sub
'End If
 
If CmbInsuranceCo.Text = "" Then
   MsgBox "Blank InsuranceCompany Not Allowed!!!"
   CmbInsuranceCo.SetFocus
   Exit Sub
End If
      
If CmbPolicyNo.Text = "" Then
   MsgBox "Blank Policy Not Allowed!!!"
   CmbPolicyNo.SetFocus
   Exit Sub
End If
       
If Edit_Flg = "A" Then
   RsTxn_CDC.AddNew
   RsTxn_CDC!CDCID = GetCDCID
   RsTxn_CDC!G_UID = GetGUID
   TxtCDCID.Text = RsTxn_CDC!CDCID
Else
   RsTxn_CDC.MoveFirst
   RsTxn_CDC.Find "CDCID= " & TxtCDCID
   If Not RsTxn_CDC.EOF Then
     If RsTxn_CDC!CDCID <> TxtCDCID.Text Then
       MsgBox "Duplicate CDC Name Not Allowed !!! "
       Exit Sub
     End If
   End If
   RsTxn_CDC.MoveFirst
   RsTxn_CDC.Find " CDCID= " & TxtCDCID.Text
End If

RsTxn_CDC!CDCID = TxtCDCID.Text
RsTxn_CDC!CDCDate = CDCDate.Value
RsTxn_CDC!CMPID = mCMPID
RsTxn_CDC!ClientIDRow = mClientIDRow
RsTxn_CDC!CommodityID = mCommodityID
RsTxn_CDC!SM_Prefix = TxtPreFix
RsTxn_CDC!Flag = "O"
RsTxn_CDC!PrintFlag = "0"
RsTxn_CDC!CommodityConditon = IIf(IsNull(TxtCommodityCondition), "", TxtCommodityCondition)


If IsNull(RsTxn_CDC!CreatedBy) = True Or RsTxn_CDC!CreatedBy = "" Then
   RsTxn_CDC!CreatedBy = Gen_UserLoginID
   RsTxn_CDC!CreatedDate = Now
Else
   RsTxn_CDC!UpdatedBy = Gen_UserLoginID
   RsTxn_CDC!UpdatedDate = Now
End If


RsTxn_CDC.Update
   
Call SaveCDC_GSData
Call SaveCDC_AssayerData
Call SaveCDC_InsuranceData
End Sub
Private Sub SaveCDC_GSData()
Call TableTxn_CDC_GSLDetail(" where sm_prefix = '" & TxtPreFix & "' And CDCID = " & TxtCDCID)
'RsTxn_CDC_GSLDetail.Find ("CDCID = " & TxtCDCID)
For I = 1 To MSHFlexGrid2.Rows - 1
   If MSHFlexGrid2.TextMatrix(I, 1) <> "" Then
     If MSHFlexGrid2.TextMatrix(I, 0) <> "" Then
       RsTxn_CDC_GSLDetail.Find "CDCNO = " & Val(MSHFlexGrid2.TextMatrix(I, 16))
       If Not RsTxn_CDC_GSLDetail.EOF Then
           'RsTxn_CDC_GSLDetail.MoveFirst
           'RsTxn_CDC_GSLDetail.Find "CDCNo = " & MSHFlexGrid1.TextMatrix(i, 16)
           'If RsTxn_CDC_GSLDetail.EOF Then
              'MsgBox "Not Found GSL !!!!"
              'Exit Sub
           'End If
       Else
         RsTxn_CDC_GSLDetail.AddNew
         RsTxn_CDC_GSLDetail!CDCNo = GetCDCNo
         RsTxn_CDC_GSLDetail!G_UID = GetGUID
       End If
       
       RsTxn_CDC_GSLDetail!CDCID = TxtCDCID
       RsTxn_CDC_GSLDetail!GSLID = MSHFlexGrid2.TextMatrix(I, 19)
       RsTxn_CDC_GSLDetail!StorageFrom = MSHFlexGrid2.TextMatrix(I, 11)
       RsTxn_CDC_GSLDetail!StorageTo = MSHFlexGrid2.TextMatrix(I, 12)
       RsTxn_CDC_GSLDetail!StorageRate = MSHFlexGrid2.TextMatrix(I, 14)
       RsTxn_CDC_GSLDetail!StorageAmount = MSHFlexGrid2.TextMatrix(I, 15)
       RsTxn_CDC_GSLDetail!CDCDate = CDCDate.Value 'MSHFlexGrid1.TextMatrix(i, 16)
       RsTxn_CDC_GSLDetail!PrintFlag = "0" 'MSHFlexGrid1.TextMatrix(i, 18)
       'RsTxn_CDC_GSLDetail!Flag = "F" 'MSHFlexGrid1.TextMatrix(i, 18)
       ' By defult it will marked as Payout changes effected from 1.04.2010
       RsTxn_CDC_GSLDetail!Flag = "P" 'MSHFlexGrid1.TextMatrix(i, 18)
       RsTxn_CDC_GSLDetail!OldCDCNo = "0"
       RsTxn_CDC_GSLDetail!CdtfNo = MSHFlexGrid2.TextMatrix(I, 20)
       RsTxn_CDC_GSLDetail!CDTFDate = MSHFlexGrid2.TextMatrix(I, 11)
       RsTxn_CDC_GSLDetail!ClientIDRow = mClientIDRow
       RsTxn_CDC_GSLDetail!SM_Prefix = TxtPreFix
       RsTxn_CDC_GSLDetail.Update
       Call TableMst_GSL(" Where sm_prefix = '" & TxtPreFix & "' And GSLID = " & MSHFlexGrid2.TextMatrix(I, 19))
       'RsMst_GSL.Find (")
       If MSHFlexGrid2.TextMatrix(I, 0) = "*" Then
         RsMst_GSL!ISINID = RsTxn_CDC_GSLDetail!CDCNo
         'RsMst_GSL!Flag = Check1.Value
       Else
         RsMst_GSL!ISINID = ""
         RsMst_GSL!Flag = 0
       End If
       RsMst_GSL.Update
       RsMst_GSL.Close
       RsTxn_CDC_GSLDetail.Requery
     End If
   End If
Next
'RsMst_GSL.Close
End Sub
Private Sub SaveCDC_AssayerData()

Call TableTxn_CDC_AssayerDetail(" where sm_prefix = '" & TxtPreFix & "' And CDCID = " & TxtCDCID)
RsTxn_CDC_AssayerDetail.Find "CDCID = " & TxtCDCID

If Not RsTxn_CDC_AssayerDetail.EOF Then
   'RsTxn_CDC_AssayerDetail.MoveFirst
Else
   RsTxn_CDC_AssayerDetail.AddNew
   RsTxn_CDC_AssayerDetail!G_UID = GetGUID
End If

RsTxn_CDC_AssayerDetail!CDCID = TxtCDCID
RsTxn_CDC_AssayerDetail!AssayerReportNo = TxtAssayerNo
RsTxn_CDC_AssayerDetail!AssayerName = TxtAssayerName
RsTxn_CDC_AssayerDetail!AssayerReportDate = ARDate.Value
RsTxn_CDC_AssayerDetail!Moisture = TxtMoisture
RsTxn_CDC_AssayerDetail!Damaged = TxtDamaged
RsTxn_CDC_AssayerDetail!Foreign = TxtForeign
RsTxn_CDC_AssayerDetail!Polarization = TxtPolarization
RsTxn_CDC_AssayerDetail!GrainSize = TxtGrainSize
RsTxn_CDC_AssayerDetail!Brokan = TxtBrokan
RsTxn_CDC_AssayerDetail!OilContain = TxtOil
RsTxn_CDC_AssayerDetail!colour = TxtColour
RsTxn_CDC_AssayerDetail!Ageing = TxtAgeing
RsTxn_CDC_AssayerDetail!StapleLength = TxtStapleLength
RsTxn_CDC_AssayerDetail!OrganicInorganic = TxtOrganic
RsTxn_CDC_AssayerDetail!WaxHoney = TxtWax
RsTxn_CDC_AssayerDetail!Grade = TxtGradeByAssayer
RsTxn_CDC_AssayerDetail!NED = NED.Value
RsTxn_CDC_AssayerDetail!FED = FED.Value
RsTxn_CDC_AssayerDetail!OverAllQuality = TxtOther  'OverAll Quality
RsTxn_CDC_AssayerDetail!AcceptenceFlag = ChkAcceptFlag
RsTxn_CDC_AssayerDetail!Remarks = TxtRemark
RsTxn_CDC_AssayerDetail!Flag = ""
RsTxn_CDC_AssayerDetail!SM_Prefix = TxtPreFix
RsTxn_CDC_AssayerDetail.Update
RsTxn_CDC_AssayerDetail.Close

End Sub
Private Sub SaveCDC_InsuranceData()
Dim strSplitString() As String
Call TableTxn_CDC_InsuranceDetail(" where sm_prefix = '" & TxtPreFix & "' And CDCID = " & TxtCDCID)
RsTxn_CDC_InsuranceDetail.Find "CDCID = " & TxtCDCID
If Not RsTxn_CDC_InsuranceDetail.EOF Then
  strSplitString = Split(CmbPolicyNo.Text, "~")
  PolicyNo_ = Trim(strSplitString(0))
  CNo = Trim(strSplitString(1))
Else
  RsTxn_CDC_InsuranceDetail.AddNew
  RsTxn_CDC_InsuranceDetail!G_UID = GetGUID
End If
RsTxn_CDC_InsuranceDetail!CDCID = TxtCDCID
RsTxn_CDC_InsuranceDetail!PolicyNo = PolicyNo_
RsTxn_CDC_InsuranceDetail!CoverNoteNo = CNo
RsTxn_CDC_InsuranceDetail!Flag = ""
RsTxn_CDC_InsuranceDetail!SM_Prefix = TxtPreFix
RsTxn_CDC_InsuranceDetail.Update
RsTxn_CDC_InsuranceDetail.Close
Call Indata
End Sub
Private Sub Indata()
Call TableMst_Config
TxtNoofDays = IIf(IsNull(RsMst_Config!SportStorageChargesDays), 0, RsMst_Config!SportStorageChargesDays) - 1

Call ClearData
Call LockTxtBox(True)
Call GButtonLock(Me, True)
mExchangeTypeID = GetExchangeID("SPOT")
Call InDataCDC
Call InDataCDC_GSL
Call InDataCDC_Assayer
Call IndataCDC_Insurance
CmdPrint.Enabled = True
End Sub
Private Sub InDataCDC()
Dim ans As Variant
TxtCDCID = RsTxn_CDC!CDCID
TxtPreFix = RsTxn_CDC!SM_Prefix
CDCDate.Value = RsTxn_CDC!CDCDate
mClientIDRow = RsTxn_CDC!ClientIDRow
ans = GetClientNameByRow(mClientIDRow)
TxtClientID = ans(0)
TxtClientName = ans(1)
TxtDPID = ans(2)
TxtDPName = ans(3)

mCMPID = RsTxn_CDC!CMPID
TxtPreFix = GetPreFix(RsTxn_CDC!CMPID, "P")
TxtCMPID = GetCMPName(mCMPID)
TxtCMPAddress = GetCMPAddress(mCMPID, "A")
TxtLocationID = GetLocationFromCMP(mCMPID)

mCommodityID = RsTxn_CDC!CommodityID
TxtCommodity = GetCommodityName(mCommodityID)
TxtCommodityCondition = IIf(IsNull(RsTxn_CDC!CommodityConditon), "", RsTxn_CDC!CommodityConditon)


TxtCreated = IIf(IsNull(RsTxn_CDC!CreatedBy), "", RsTxn_CDC!CreatedBy) & " :- " & IIf(IsNull(RsTxn_CDC!CreatedDate), "", RsTxn_CDC!CreatedDate)
TxtUpdated = IIf(IsNull(RsTxn_CDC!UpdatedBy), "", RsTxn_CDC!UpdatedBy) & " :- " & IIf(IsNull(RsTxn_CDC!UpdatedDate), "", RsTxn_CDC!UpdatedDate)


End Sub
Private Sub InDataCDC_GSL()
Dim mStdDedLess As Double
Call Grid_Head

'Call TableTxn_CDC_GSLDetail(" where sm_prefix = '" & TxtPreFix & "' And CDCID = " & TxtCDCID)
'MSHFlexGrid2.Rows = RsTxn_CDC_GSLDetail.RecordCount + 1

tmp = " Select TCGD.GSLID,MG.Address,MG.GodownNo,MGSL.StackNo,MGSL.LotNo,MGSL.BGSLID,TCGD.StorageFrom,TCGD.StorageTo,TCGD.StorageRate,TCGD.StorageAmount,TCGD.CDCNo,TCGD.CDCDate,TCGD.PrintFlag,TCGD.Flag,"
tmp = tmp & " TCGD.CdtfNo"
tmp = tmp & " From Txn_CDC_GSL_Detail TCGD"
tmp = tmp & " Inner Join Mst_GSL MGSL On TCGD.SM_PRefix = MGSL.SM_Prefix And TCGD.GSLID = MGSL.GSLID And MGSL.ExchangeTypeID = 3"
tmp = tmp & " Inner Join Mst_Godown MG On MGSL.GodownID = MG.GodownID"
tmp = tmp & " Where TCGD.CDCID = " & TxtCDCID & " And TCGD.SM_Prefix = '" & TxtPreFix & "'"
Call TmpTable

'MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + RsTxn_CDC_GSLDetail.RecordCount
MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + TmpRs.RecordCount

'TxtStorageRate = RsTxn_CDC_GSLDetail!StorageRate

TxtStorageRate = TmpRs!StorageRate

k = 0
Dim ans As Variant
Flag_EditDelete = 0
For I = 0 To MSHFlexGrid2.Rows - 1
k = k + 1
    'MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
'    If RsTxn_CDC_GSLDetail.EOF Then Exit For
'
'    If RsTxn_CDC_GSLDetail!Flag = "P" Or RsTxn_CDC_GSLDetail!Flag = "T" Or RsTxn_CDC_GSLDetail!Flag = "W" Then
'       Flag_EditDelete = "1"
'    End If
'
'    tmp = " Where  sm_prefix = '" & TxtPreFix & "' And Txn_CxTF_GSL.GSLID = " & RsTxn_CDC_GSLDetail!GSLID
'    Call TableTxn_CxTF_GSL(tmp)
'
'    MSHFlexGrid2.TextMatrix(k, 19) = RsTxn_CDC_GSLDetail!GSLID
'
'    ans = GetGSLByGLSID(RsTxn_CDC_GSLDetail!GSLID, TxtPreFix)
'
'
'    MSHFlexGrid2.TextMatrix(k, 0) = "*" 'RsTxn_CDC_GSLDetail!TruckDetailID
'    MSHFlexGrid2.TextMatrix(k, 1) = ans(11) ' Godown Address
'    MSHFlexGrid2.TextMatrix(k, 2) = ans(3)  ' GodownNo
'    MSHFlexGrid2.TextMatrix(k, 3) = ans(4)  ' StackNo
'    MSHFlexGrid2.TextMatrix(k, 4) = ans(5)  ' TmpRs!LotNo
'
''    MSHFlexGrid2.TextMatrix(k, 5) = RsTxn_CxTF_GSL!NoofBags
''    MSHFlexGrid2.TextMatrix(k, 6) = RsTxn_CxTF_GSL!AdjNCMSLWeight  'in MT
''    MSHFlexGrid2.TextMatrix(k, 7) = RsTxn_CxTF_GSL!AdjNCMSLWeight * 1000 'in Kg
''    MSHFlexGrid2.TextMatrix(k, 8) = RsTxn_CxTF_GSL!GunnyWeight * 1000 'TmpRs!GunnyWeight
''    MSHFlexGrid2.TextMatrix(k, 9) = RsTxn_CxTF_GSL!StdDeduction * 1000 'TmpRs!StdDeduction
''    MSHFlexGrid2.TextMatrix(k, 10) = RsTxn_CxTF_GSL!DematWeight * 1000 '  net weight in KGS
'
'    For y = 1 To RsTxn_CxTF_GSL.RecordCount
'        MSHFlexGrid2.TextMatrix(k, 5) = Val(MSHFlexGrid2.TextMatrix(k, 5)) + RsTxn_CxTF_GSL!NoofBags
'        MSHFlexGrid2.TextMatrix(k, 6) = Val(MSHFlexGrid2.TextMatrix(k, 6)) + RsTxn_CxTF_GSL!AdjNCMSLWeight  'in MT
'        MSHFlexGrid2.TextMatrix(k, 7) = Val(MSHFlexGrid2.TextMatrix(k, 7)) + RsTxn_CxTF_GSL!AdjNCMSLWeight * 1000 'in Kg
'        MSHFlexGrid2.TextMatrix(k, 8) = Val(MSHFlexGrid2.TextMatrix(k, 8)) + RsTxn_CxTF_GSL!GunnyWeight * 1000 'TmpRs!GunnyWeight
'        MSHFlexGrid2.TextMatrix(k, 9) = Val(MSHFlexGrid2.TextMatrix(k, 9)) + IIf(IsNull(RsTxn_CxTF_GSL!StdDeduction), 0, RsTxn_CxTF_GSL!StdDeduction) * 1000 'TmpRs!StdDeduction
'        MSHFlexGrid2.TextMatrix(k, 10) = Val(MSHFlexGrid2.TextMatrix(k, 10)) + RsTxn_CxTF_GSL!DematWeight * 1000 '  net weight in KGS
'        RsTxn_CxTF_GSL.MoveNext
'    Next
'
'
'    MSHFlexGrid2.TextMatrix(k, 11) = RsTxn_CDC_GSLDetail!StorageFrom
'    MSHFlexGrid2.TextMatrix(k, 12) = RsTxn_CDC_GSLDetail!StorageTo
'    MSHFlexGrid2.TextMatrix(k, 13) = (RsTxn_CDC_GSLDetail!StorageTo - RsTxn_CDC_GSLDetail!StorageFrom) - Val(TxtNoofDays) '7
'    MSHFlexGrid2.TextMatrix(k, 14) = RsTxn_CDC_GSLDetail!StorageRate
'    MSHFlexGrid2.TextMatrix(k, 15) = RsTxn_CDC_GSLDetail!StorageAmount
'    MSHFlexGrid2.TextMatrix(k, 16) = RsTxn_CDC_GSLDetail!CDCNo
'    MSHFlexGrid2.TextMatrix(k, 17) = RsTxn_CDC_GSLDetail!CDCDate
'    MSHFlexGrid2.TextMatrix(k, 18) = RsTxn_CDC_GSLDetail!PrintFlag
'    MSHFlexGrid2.TextMatrix(k, 19) = RsTxn_CDC_GSLDetail!GSLID
'    MSHFlexGrid2.TextMatrix(k, 20) = RsTxn_CDC_GSLDetail!CdtfNo
'
'    RsTxn_CDC_GSLDetail.MoveNext
'    RsTxn_CxTF_GSL.Close



If TmpRs.EOF Then Exit For
    
    If TmpRs!Flag = "P" Or TmpRs!Flag = "T" Or TmpRs!Flag = "W" Then
       Flag_EditDelete = "1"
    End If
    
    mBackGSL = TmpRs!GSLID
    
    If IsNull(TmpRs!BGSLID) = False Then
       Call GetBGSLID(k, TmpRs!GSLID)
    End If
    
    tmp = " Where  sm_prefix = '" & TxtPreFix & "' And GSLID = " & mBackGSL 'TmpRs!GSLID
    Call TableTxn_CxTF_GSL(tmp)
    
    MSHFlexGrid2.TextMatrix(k, 19) = TmpRs!GSLID
    
    'ans = GetGSLByGLSID(RsTxn_CDC_GSLDetail!GSLID, TxtPreFix)
    
    
    MSHFlexGrid2.TextMatrix(k, 0) = "*" 'RsTxn_CDC_GSLDetail!TruckDetailID
    MSHFlexGrid2.TextMatrix(k, 1) = TmpRs!Address 'ans(11) ' Godown Address
    MSHFlexGrid2.TextMatrix(k, 2) = TmpRs!GodownNo 'ans(3)  ' GodownNo
    MSHFlexGrid2.TextMatrix(k, 3) = TmpRs!StackNo 'ans(4)  ' StackNo
    MSHFlexGrid2.TextMatrix(k, 4) = TmpRs!LotNo 'ans(5)  ' TmpRs!LotNo
       
'    MSHFlexGrid2.TextMatrix(k, 5) = RsTxn_CxTF_GSL!NoofBags
'    MSHFlexGrid2.TextMatrix(k, 6) = RsTxn_CxTF_GSL!AdjNCMSLWeight  'in MT
'    MSHFlexGrid2.TextMatrix(k, 7) = RsTxn_CxTF_GSL!AdjNCMSLWeight * 1000 'in Kg
'    MSHFlexGrid2.TextMatrix(k, 8) = RsTxn_CxTF_GSL!GunnyWeight * 1000 'TmpRs!GunnyWeight
'    MSHFlexGrid2.TextMatrix(k, 9) = RsTxn_CxTF_GSL!StdDeduction * 1000 'TmpRs!StdDeduction
'    MSHFlexGrid2.TextMatrix(k, 10) = RsTxn_CxTF_GSL!DematWeight * 1000 '  net weight in KGS

    For y = 1 To RsTxn_CxTF_GSL.RecordCount
        MSHFlexGrid2.TextMatrix(k, 5) = Val(MSHFlexGrid2.TextMatrix(k, 5)) + RsTxn_CxTF_GSL!NoofBags
        MSHFlexGrid2.TextMatrix(k, 6) = Val(MSHFlexGrid2.TextMatrix(k, 6)) + RsTxn_CxTF_GSL!AdjNCMSLWeight  'in MT
        MSHFlexGrid2.TextMatrix(k, 7) = Val(MSHFlexGrid2.TextMatrix(k, 7)) + RsTxn_CxTF_GSL!AdjNCMSLWeight * 1000 'in Kg
        MSHFlexGrid2.TextMatrix(k, 8) = Val(MSHFlexGrid2.TextMatrix(k, 8)) + RsTxn_CxTF_GSL!GunnyWeight * 1000 'TmpRs!GunnyWeight
        MSHFlexGrid2.TextMatrix(k, 9) = Val(MSHFlexGrid2.TextMatrix(k, 9)) + IIf(IsNull(RsTxn_CxTF_GSL!StdDeduction), 0, RsTxn_CxTF_GSL!StdDeduction) * 1000 'TmpRs!StdDeduction
        MSHFlexGrid2.TextMatrix(k, 10) = Val(MSHFlexGrid2.TextMatrix(k, 10)) + RsTxn_CxTF_GSL!DematWeight * 1000 '  net weight in KGS
        RsTxn_CxTF_GSL.MoveNext
    Next

    
    MSHFlexGrid2.TextMatrix(k, 11) = TmpRs!StorageFrom 'RsTxn_CDC_GSLDetail!StorageFrom
    MSHFlexGrid2.TextMatrix(k, 12) = TmpRs!StorageTo 'RsTxn_CDC_GSLDetail!StorageTo
    MSHFlexGrid2.TextMatrix(k, 13) = (TmpRs!StorageTo - TmpRs!StorageFrom) - Val(TxtNoofDays) '7 '(RsTxn_CDC_GSLDetail!StorageTo - RsTxn_CDC_GSLDetail!StorageFrom) - Val(TxtNoofDays) '7
    MSHFlexGrid2.TextMatrix(k, 14) = TmpRs!StorageRate 'RsTxn_CDC_GSLDetail!StorageRate
    MSHFlexGrid2.TextMatrix(k, 15) = TmpRs!StorageAmount 'RsTxn_CDC_GSLDetail!StorageAmount
    MSHFlexGrid2.TextMatrix(k, 16) = TmpRs!CDCNo 'RsTxn_CDC_GSLDetail!CDCNo
    MSHFlexGrid2.TextMatrix(k, 17) = TmpRs!CDCDate 'RsTxn_CDC_GSLDetail!CDCDate
    MSHFlexGrid2.TextMatrix(k, 18) = TmpRs!PrintFlag 'RsTxn_CDC_GSLDetail!PrintFlag
    MSHFlexGrid2.TextMatrix(k, 19) = TmpRs!GSLID 'RsTxn_CDC_GSLDetail!GSLID
    MSHFlexGrid2.TextMatrix(k, 20) = TmpRs!CdtfNo 'RsTxn_CDC_GSLDetail!CdtfNo
    
    'RsTxn_CDC_GSLDetail.MoveNext
    TmpRs.MoveNext
    RsTxn_CxTF_GSL.Close
Next
'If TmpRs.State = 1 Then
'   TmpRs.Close
'End If

End Sub
Private Sub InDataCDC_Assayer()

Call TableTxn_CDC_AssayerDetail(" where sm_prefix = '" & TxtPreFix & "' And CDCID = " & TxtCDCID)
If RsTxn_CDC_AssayerDetail.RecordCount = 0 Then Exit Sub

TxtAssayerNo = RsTxn_CDC_AssayerDetail!AssayerReportNo
TxtAssayerName = RsTxn_CDC_AssayerDetail!AssayerName
TxtAssayerName.Visible = True
ARDate.Value = RsTxn_CDC_AssayerDetail!AssayerReportDate
TxtMoisture = RsTxn_CDC_AssayerDetail!Moisture
TxtDamaged = RsTxn_CDC_AssayerDetail!Damaged
TxtForeign = RsTxn_CDC_AssayerDetail!Foreign
TxtPolarization = RsTxn_CDC_AssayerDetail!Polarization
TxtGrainSize = RsTxn_CDC_AssayerDetail!GrainSize
TxtBrokan = RsTxn_CDC_AssayerDetail!Brokan
TxtOil = RsTxn_CDC_AssayerDetail!OilContain
TxtColour = RsTxn_CDC_AssayerDetail!colour
TxtAgeing = RsTxn_CDC_AssayerDetail!Ageing
TxtStapleLength = RsTxn_CDC_AssayerDetail!StapleLength
TxtOrganic = RsTxn_CDC_AssayerDetail!OrganicInorganic
TxtWax = RsTxn_CDC_AssayerDetail!WaxHoney
TxtGradeByAssayer = RsTxn_CDC_AssayerDetail!Grade
NED.Value = RsTxn_CDC_AssayerDetail!NED
FED.Value = RsTxn_CDC_AssayerDetail!FED
TxtOther = RsTxn_CDC_AssayerDetail!OverAllQuality  'OverAll Quality
ChkAcceptFlag.Value = RsTxn_CDC_AssayerDetail!AcceptenceFlag
TxtRemark = RsTxn_CDC_AssayerDetail!Remarks

End Sub
Private Sub IndataCDC_Insurance()

Call TableTxn_CDC_InsuranceDetail(" where sm_prefix = '" & TxtPreFix & "' And CDCID = " & TxtCDCID)
If RsTxn_CDC_InsuranceDetail.RecordCount = 0 Then Exit Sub

TxtPolicyNo = RsTxn_CDC_InsuranceDetail!PolicyNo & " ~ " & RsTxn_CDC_InsuranceDetail!CoverNoteNo
Call TableMst_PolicyDetail("where PolicyNo = '" & RsTxn_CDC_InsuranceDetail!PolicyNo & "' And CoverNoteNo = '" & RsTxn_CDC_InsuranceDetail!CoverNoteNo & "'")
TxtInsuranceCo = GetInsuranceCoName(RsMst_PolicyDetail!InsuranceCoID)
ValidFrom.Value = RsMst_PolicyDetail!ValidFrom
ValidTo.Value = RsMst_PolicyDetail!ValidTo
TxtAmount = RsMst_PolicyDetail!Amount
TxtRisks = RsMst_PolicyDetail!Risks
CmdInsuranceCo.Visible = False
CmdInsurance.Visible = False

RsMst_PolicyDetail.Close


End Sub
Private Function GetCDCID() As Double
Dim Retval As Double
tmp = "Select max(CDCID) from Txn_CDC where sm_prefix = '" & TxtPreFix & "'"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetCDCID = Retval
End Function
Private Function GetCDCNo() As Double
Dim Retval As Double
tmp = "Select max(CDCNo) from Txn_CDC_GSL_Detail where sm_prefix = '" & TxtPreFix & "'"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
TmpRs.Close
GetCDCNo = Retval

End Function
Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   If RsTxn_CDC.RecordCount = 0 Then Exit Sub
   Call ClearData
   Call GButtonLock(Me, True)
   Call LockTxtBox(True)
   Call FirstCmd_Click
   Exit Sub
End If

Frame0.Height = Me.Height - 1000
Frame0.Left = Picture1.Left - 50
Frame0.Top = Picture1.Top - 40
Frame0.Width = Me.Width - 300
Frame0.Visible = True
MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 2000
RsTxn_CDC.MoveFirst

''===========comment on 7 july 2012 for not clearing search grid=============
'MSHFlexGrid1.Clear
'MSHFlexGrid1.Cols = 3
'Call Grid_Head1
'DtpSFromCDCDate.Value = Date - 365
'DtpSToCDCDate.Value = Date
'CmbPreFix.Text = ""
'CmbSCmp.Text = ""
'CmbSCommodity.Text = ""
'TxtSClientName.Text = ""
'TxtSCDCNo.Text = ""
''===========comment on 7 july 2012 for not clearing search grid=============

End Sub
Private Sub SSTab1_Click(PreviousTab As Integer)
'If SaveCmd.Enabled = True Then
'  If SSTab1.Tab = 0 Then
'    If Edit_Flg = "A" Then
'
'    ElseIf Edit_Flg = "E" Then
'
'    Else
'
'    End If
'  ElseIf SSTab1.Tab = 1 Then
'    If Edit_Flg = "A" Then
'
'     'CmbAssayerName.Visible = True
'      TxtAssayerName.Visible = True
'    ElseIf Edit_Flg = "E" Then
'     'CmbAssayerName.Visible = True
'      TxtAssayerName.Visible = True
'    Else
'      TxtAssayerName.Visible = True
'    End If
'  ElseIf SSTab1.Tab = 2 Then
'    If Edit_Flg = "A" Then
'      CmbInsuranceCo.Visible = True
'      CmbPolicyNo.Visible = True
'      TxtInsuranceCo.Visible = False
'      TxtPolicyNo.Visible = False
'    ElseIf Edit_Flg = "E" Then
'      CmbInsuranceCo.Visible = True
'      CmbPolicyNo.Visible = True
'      TxtInsuranceCo.Visible = False
'      TxtPolicyNo.Visible = False
'    Else
'      CmbInsuranceCo.Visible = False
'      CmbPolicyNo.Visible = False
'      TxtInsuranceCo.Visible = True
'      TxtPolicyNo.Visible = True
'    End If
'
'  End If
'End If
End Sub

Private Sub CmbInsuranceCo_GotFocus()
tmp = CmbInsuranceCo.Text
Call TableMst_InsuranceCo("")

CmbInsuranceCo.Clear
If RsMst_InsuranceCo.RecordCount = 0 Then
   MsgBox "No InsuranceCo found"
   'CmdInsuranceCo.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_InsuranceCo.RecordCount
    CmbInsuranceCo.AddItem RsMst_InsuranceCo!InsuranceCoName
    RsMst_InsuranceCo.MoveNext
Next
CmbInsuranceCo.Text = tmp
   
End Sub
Private Sub CmbInsuranceCo_LostFocus()

If CmbInsuranceCo.Text = "" Then
   Exit Sub
End If
RsMst_InsuranceCo.MoveFirst
RsMst_InsuranceCo.Find "InsuranceCoName = '" & CmbInsuranceCo.Text & "'"

If RsMst_InsuranceCo.EOF Then
   MsgBox "Invalid InsuranceCo selection "
   CmbInsuranceCo.SetFocus
   Exit Sub
End If
mInsuranceCoID = RsMst_InsuranceCo.Fields(0)
If CmbInsuranceCo.Text <> tmp Then
   CmbPolicyNo.Text = ""
   TxtRisks = ""
   TxtAmount = ""
   ValidFrom.Value = Date
   ValidTo.Value = Date
End If

End Sub

Private Sub CmbPolicyNo_GotFocus()
If CmbInsuranceCo.Text = "" Or mInsuranceCoID = "" Then Exit Sub

tmp = CmbPolicyNo.Text
Call TableMst_PolicyDetail(" Where InsuranceCoID =  " & mInsuranceCoID & " Order by  PolicyNo,CoverNoteNo")

CmbPolicyNo.Clear
If RsMst_PolicyDetail.RecordCount = 0 Then
   MsgBox "No Insurance Detail found"
   CmbPolicyNo.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_PolicyDetail.RecordCount
    CmbPolicyNo.AddItem RsMst_PolicyDetail!PolicyNo & " ~ " & RsMst_PolicyDetail!CoverNoteNo
    RsMst_PolicyDetail.MoveNext
Next
CmbPolicyNo.Text = tmp
End Sub

Private Sub CmbPolicyNo_LostFocus()
Dim strSplitString() As String
If CmbPolicyNo.Text = "" Then
   Exit Sub
End If
strSplitString = Split(CmbPolicyNo.Text, "~")
'On Error GoTo M_lbl:
tmp = UBound(strSplitString)
If tmp <> 1 Then
   MsgBox "Invalid selection"
   CmbPolicyNo.SetFocus
   Exit Sub
End If
PolicyNo_ = Trim(strSplitString(0))
CNo = Trim(strSplitString(1))
RsMst_PolicyDetail.MoveFirst
For I = 1 To RsMst_PolicyDetail.RecordCount
  If PolicyNo_ = Trim(RsMst_PolicyDetail!PolicyNo) And CNo = Trim(RsMst_PolicyDetail!CoverNoteNo) Then
     TxtRisks = RsMst_PolicyDetail!Risks
     TxtAmount = RsMst_PolicyDetail!Amount
     ValidFrom.Value = RsMst_PolicyDetail!ValidFrom
     ValidTo.Value = RsMst_PolicyDetail!ValidTo
     Exit Sub
  End If
  RsMst_PolicyDetail.MoveNext
Next
If RsMst_PolicyDetail.EOF Then
  MsgBox "Invalid InsuranceCo selection "
  CmbPolicyNo.SetFocus
  Exit Sub
End If
mInsuranceCoID = RsMst_PolicyDetail.Fields(0)

End Sub

Private Sub StorageTo_LostFocus()
If StorageFrom.Value > StorageTo.Value Then
    MsgBox "To Date must be Greater Than From Date !!!!!! "
    StorageTo.SetFocus
    Exit Sub
Else
    mStorageDays = (StorageTo.Value - StorageFrom) - Val(TxtNoofDays) '7
    If mStorageDays < 0 Then
       MsgBox "To Date must be Greater Than From Date plus " & (TxtNoofDays) & " Days !!!!!! "
       StorageTo.SetFocus
       Exit Sub
    End If
    LblStorageDays = mStorageDays 'mStorageDays - 6
    CmdUpdate.Enabled = True
    CmdCancel.Enabled = True
    CmdUpdate.SetFocus
End If
End Sub

Private Sub TxtClientID_LostFocus()
If SaveCmd.Enabled = False Then Exit Sub
CmbClientDP.Visible = False
If TxtClientID = "" Then Exit Sub
 'If CmbExchangeTypeID.Text = "" Then Exit Sub
 
Dim ans As Variant
mExchangeTypeID = GetExchangeID("SPOT")
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
Else
   If CmbClientDP.Visible = True Then
      CmbClientDP.SetFocus
   End If
End If
'If SaveCmd.Enabled = True Then
  'Call GetGridData(mClientIDRow, mExchangeTypeID, mCMPID, mCommodityID)
'End If
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
Call GetGridData(mClientIDRow, mExchangeTypeID, mCMPID, mCommodityID)
End Sub
Public Sub ClearData()

TxtCDCID.Text = ""
CDCDate.Value = Date
TxtCMPID = ""
TxtCMPAddress = ""
TxtWHRLocation = ""
'TxtDepositor = GetDepositorName()
TxtClientID = ""
TxtCommodity = ""
TxtLocationID = ""
TxtClientName = ""
TxtDPID = ""
TxtDPName = ""
TxtStorageRate = ""
TxtCommodityCondition = ""
TxtAssayerNo = ""
TxtAssayerName = ""
ARDate.Value = Date
TxtMoisture = ""
TxtDamaged = ""
TxtForeign = ""
TxtPolarization = ""
TxtGrainSize = ""
TxtBrokan = ""
TxtOil = ""
TxtColour = ""
TxtAgeing = ""
TxtStapleLength = ""
TxtOrganic = ""
TxtWax = ""
TxtGradeByAssayer = ""
NED.Value = Date
FED.Value = Date
TxtOther = ""  'OverAll Quality
ChkAcceptFlag.Value = 0
TxtRemark = ""
ValidFrom.Value = Date
ValidTo.Value = Date
TxtAmount = ""
TxtRisks = ""
CmbInsuranceCo.Text = ""
CmbPolicyNo.Text = ""
CmdUpdate.Enabled = False
CmdCancel.Enabled = False
LblStorageDays.Caption = ""
Call Grid_Head
mClientIDRow = 0
mCMPID = 0
mCommodityID = 0
End Sub
Private Sub TxtStorageRate_LostFocus()
If TxtStorageRate.Text = "" Then Exit Sub
If IsNumeric(TxtStorageRate) = False Then
   MsgBox "Invalid numeric formate"
   TxtStorageRate.SetFocus
   Exit Sub
End If
If SaveCmd.Enabled = True Then
   Call GetGridData(mClientIDRow, mExchangeTypeID, mCMPID, mCommodityID)
End If
End Sub
Function CheckSelection() As Boolean
Dim Retval As Boolean
Retval = True
If MSHFlexGrid2.Rows > 2 Then
   For u = 1 To MSHFlexGrid2.Rows - 1
     If MSHFlexGrid2.TextMatrix(u, 0) = "*" Then
        Retval = False
        Exit For
     End If
   Next
End If
CheckSelection = Retval
End Function


Private Sub CmdPrint_Click()

RsTxn_CDC!PrintFlag = "P"
RsTxn_CDC.Update


Gen_mTimeStamp = GetTimeStamp
Call Create_Xls(Gen_mTimeStamp & Gen_UserName & "_CDC-" & Trim(TxtPreFix) & "" & Trim(TxtCDCID) & ".xls")
Call Xls_Print_Rpt

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
oBook.SaveAs (Pat & "\excel\" & FileName)

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
'Dim oXL As Excel.Application
'Dim oWB As Excel.Workbook
'Dim oWS As Excel.Worksheet
'Dim mRow As Integer
'Dim mCol As Integer
Dim ans As Variant
MsgBox ("Wait till next message")
Set oXL = New Excel.Application
'Dim Pat As String
Pat = App.Path
'Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_UserName & "_CDC-" & Trim(TxtPreFix) & "" & Trim(TxtCDCID) & ".xls")
 
 
 Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "_CDC-" & Trim(TxtPreFix) & "" & Trim(TxtCDCID) & ".xls")
 
 
 
 
 
For I = 1 To MSHFlexGrid2.Rows - 2

    Set oWS = oWB.Worksheets("Sheet" & I)
    'oWS.PageSetup.RightMargin = 0.05
    oWS.Name = TxtPreFix & "-" & MSHFlexGrid2.TextMatrix(I, 16)
    oWS.ClearArrows
    mRow = 1

    oWS.Range("a1:d1").Merge
    oWS.Range("a:a").ColumnWidth = 21.86
    oWS.Range("b:b").ColumnWidth = 21.86
    oWS.Range("c:c").ColumnWidth = 21.86
    oWS.Range("d:d").ColumnWidth = 21.86
    
    'oWS.Range("A1:E1").Borders.Color = RGB(0, 0, 0)
    oWS.Cells(mRow, c + 1) = "NATIONAL COLLATERAL MANAGEMENT SERVICES LIMITED"
    oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
    oWS.Cells(mRow, c + 1).Font.Bold = True
    oWS.Cells(mRow, c + 1).Font.Size = 12
    
    'oWS.Cells(1, 1).Font.ColorIndex = 4
    mRow = mRow + 1
    'oWS.Range("a3:e3").Merge
    tmp = "a" & mRow & ":" & "d" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, c + 1) = "Commodity Deposit Certificate (CDC)"
    oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
    oWS.Cells(mRow, c + 1).Font.Bold = True
    oWS.Cells(mRow, c + 1).Font.Size = 9
    oWS.Cells(mRow, c + 1).RowHeight = 24
    
    mRow = mRow + 1
    
    oWS.Cells(mRow, c + 1) = "Warehouse License Number"
    oWS.Cells(mRow, c + 1).Font.Bold = True
    oWS.Cells(mRow, c + 1).WrapText = True
    oWS.Cells(mRow, c + 1).Font.Size = 9
    
    oWS.Cells(mRow, c + 2) = GetLicenseFromGSL(MSHFlexGrid2.TextMatrix(I, 19))
    oWS.Cells(mRow, c + 2).HorizontalAlignment = 2
    oWS.Cells(mRow, c + 2).Font.Size = 9
    
    oWS.Cells(mRow, c + 3) = "CDC No"
    oWS.Cells(mRow, c + 3).Font.Size = 9
    oWS.Cells(mRow, c + 3).Font.Bold = True
    oWS.Cells(mRow, c + 4) = TxtPreFix & "-" & MSHFlexGrid2.TextMatrix(I, 16)
    oWS.Cells(mRow, c + 4).HorizontalAlignment = 2
    oWS.Cells(mRow, c + 4).Font.Size = 9
    
    mRow = mRow + 1
    
    oWS.Cells(mRow, c + 1) = "Location & Address of Warehouse"
    oWS.Cells(mRow, c + 1).Font.Size = 9
    oWS.Cells(mRow, c + 1).Font.Bold = True
    tmp = "b" & mRow & ":" & "d" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, c + 2) = TxtLocationID & "--" & TxtCMPAddress
    oWS.Cells(mRow, c + 2).HorizontalAlignment = 2
    oWS.Cells(mRow, c + 2).Font.Size = 9
    oWS.Cells(mRow, c + 2).WrapText = True
    
    mRow = mRow + 1
    
    oWS.Cells(mRow, c + 1) = "CMP ID & Address"
    oWS.Cells(mRow, c + 1).Font.Size = 9
    oWS.Cells(mRow, c + 1).WrapText = True
    oWS.Cells(mRow, c + 1).Font.Bold = True
    
    tmp = "b" & mRow & ":" & "d" & mRow
    oWS.Range(tmp).Merge
    
    oWS.Cells(mRow, c + 2) = TxtCMPID
    oWS.Cells(mRow, c + 2).HorizontalAlignment = 2
    oWS.Cells(mRow, c + 2).Font.Size = 9
    oWS.Cells(mRow, c + 2).WrapText = True
    
    mRow = mRow + 1
    
    oWS.Cells(mRow, c + 1) = "Lot No"
    oWS.Cells(mRow, c + 1).Font.Bold = True
    oWS.Cells(mRow, c + 1).Font.Size = 9
    oWS.Cells(mRow, c + 1).WrapText = True
    
    tmp = "b" & mRow & ":" & "d" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, c + 2) = MSHFlexGrid2.TextMatrix(I, 4)
    oWS.Cells(mRow, c + 2).Font.Size = 9
    
    mRow = mRow + 1
    
    oWS.Cells(mRow, c + 1) = "Name of Depositor"
    oWS.Cells(mRow, c + 1).Font.Bold = True
    oWS.Cells(mRow, c + 1).Font.Size = 9
    
    tmp = "b" & mRow & ":" & "d" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, c + 2) = TxtClientID & "-" & TxtClientName
    oWS.Cells(mRow, c + 2).Font.Size = 9
    oWS.Cells(mRow, c + 2).WrapText = True
    
    mRow = mRow + 1
    
    oWS.Cells(mRow, c + 1) = "Commodity"
    oWS.Cells(mRow, c + 1).Font.Bold = True
    oWS.Cells(mRow, c + 1).Font.Size = 9
    oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
    oWS.Cells(mRow, c + 2) = "No of Units"
    oWS.Cells(mRow, c + 2).Font.Size = 9
    oWS.Cells(mRow, c + 2).Font.Bold = True
    oWS.Cells(mRow, c + 2).HorizontalAlignment = -4108
    oWS.Cells(mRow, c + 3) = "Net Quantity(Kgs)"
    oWS.Cells(mRow, c + 3).Font.Size = 9
    oWS.Cells(mRow, c + 3).Font.Bold = True
    oWS.Cells(mRow, c + 3).HorizontalAlignment = -4108
    oWS.Cells(mRow, c + 4) = "Condition of Commodity"
    oWS.Cells(mRow, c + 4).Font.Size = 9
    oWS.Cells(mRow, c + 4).Font.Bold = True
    oWS.Cells(mRow, c + 4).HorizontalAlignment = -4108
    
    mRow = mRow + 1
    oWS.Cells(mRow, c + 1) = TxtCommodity
    oWS.Cells(mRow, c + 1).Font.Size = 9
    oWS.Cells(mRow, c + 2) = Format(MSHFlexGrid2.TextMatrix(I, 5), "#######0.000") 'MSHFlexGrid2.TextMatrix(i, 5)
    oWS.Cells(mRow, c + 2).Font.Size = 9
    oWS.Cells(mRow, c + 3) = Format(MSHFlexGrid2.TextMatrix(I, 10), "#######0.000") 'MSHFlexGrid2.TextMatrix(i, 10)
    oWS.Cells(mRow, c + 3).Font.Size = 9
    oWS.Cells(mRow, c + 4) = TxtCommodityCondition
    oWS.Cells(mRow, c + 4).Font.Size = 9
    
    mRow = mRow + 1
    
    oWS.Cells(mRow, c + 1) = "Name & License No of Assayer"
    oWS.Cells(mRow, c + 1).Font.Size = 9
    oWS.Cells(mRow, c + 1).Font.Bold = True
    oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
    oWS.Cells(mRow, c + 1).WrapText = True
    oWS.Cells(mRow, c + 2) = "Grade Parameter Values as per Assayer's Certificate"
    oWS.Cells(mRow, c + 2).Font.Size = 9
    oWS.Cells(mRow, c + 2).Font.Bold = True
    oWS.Cells(mRow, c + 2).HorizontalAlignment = -4108
    oWS.Cells(mRow, c + 2).WrapText = True
    oWS.Cells(mRow, c + 3) = "Grade as per Assayer's Certificate"
    oWS.Cells(mRow, c + 3).Font.Size = 9
    oWS.Cells(mRow, c + 3).Font.Bold = True
    oWS.Cells(mRow, c + 3).HorizontalAlignment = -4108
    oWS.Cells(mRow, c + 3).WrapText = True
    oWS.Cells(mRow, c + 4) = "Grade Valid till"
    oWS.Cells(mRow, c + 4).Font.Size = 9
    oWS.Cells(mRow, c + 4).Font.Bold = True
    oWS.Cells(mRow, c + 4).HorizontalAlignment = -4108
    oWS.Cells(mRow, c + 4).WrapText = True
    
    mRow = mRow + 1
    
    oWS.Cells(mRow, c + 1) = TxtAssayerNo & "-" & TxtAssayerName
    oWS.Cells(mRow, c + 1).Font.Size = 9
    oWS.Cells(mRow, c + 1).WrapText = True
    oWS.Cells(mRow, c + 2) = ""
    oWS.Cells(mRow, c + 2).Font.Size = 9
    oWS.Cells(mRow, c + 2).WrapText = True
    oWS.Cells(mRow, c + 3) = TxtGradeByAssayer
    oWS.Cells(mRow, c + 3).Font.Size = 9
    oWS.Cells(mRow, c + 3).WrapText = True
    oWS.Cells(mRow, c + 4) = Format(FED.Value, "dd-mmm-yyyy")
    oWS.Cells(mRow, c + 4).Font.Size = 9
    oWS.Cells(mRow, c + 4).WrapText = True
    mRow = mRow + 1
    
    oWS.Cells(mRow, c + 1) = "Insurance Details"
    oWS.Cells(mRow, c + 1).Font.Size = 9
    oWS.Cells(mRow, c + 1).Font.Bold = True
    oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
    oWS.Cells(mRow, c + 1).WrapText = True
    
    tmp = "b" & mRow & ":" & "c" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, c + 2) = "Insurance Company"
    oWS.Cells(mRow, c + 2).Font.Size = 9
    oWS.Cells(mRow, c + 2).Font.Bold = True
    oWS.Cells(mRow, c + 2).HorizontalAlignment = -4108
    oWS.Cells(mRow, c + 2).WrapText = True
    
    oWS.Cells(mRow, c + 4) = "Amount Insured"
    oWS.Cells(mRow, c + 4).Font.Size = 9
    oWS.Cells(mRow, c + 4).Font.Bold = True
    oWS.Cells(mRow, c + 4).HorizontalAlignment = -4108
    oWS.Cells(mRow, c + 4).WrapText = True
    
    mRow = mRow + 1
    
    oWS.Cells(mRow, c + 1) = TxtRisks
    oWS.Cells(mRow, c + 1).Font.Size = 9
    oWS.Cells(mRow, c + 1).WrapText = True
    
    tmp = "b" & mRow & ":" & "c" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, c + 2) = TxtInsuranceCo
    oWS.Cells(mRow, c + 2).Font.Size = 9
    oWS.Cells(mRow, c + 2).WrapText = True
           
    oWS.Cells(mRow, c + 4) = TxtAmount
    oWS.Cells(mRow, c + 4).Font.Size = 9
    oWS.Cells(mRow, c + 4).WrapText = True
    
    mRow = mRow + 1
    
    oWS.Cells(mRow, c + 1) = "Goods Are Accepted for Storage From"
    oWS.Cells(mRow, c + 1).Font.Size = 9
    oWS.Cells(mRow, c + 1).Font.Bold = True
    oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
    oWS.Cells(mRow, c + 1).WrapText = True
    
    oWS.Cells(mRow, c + 2) = Format(MSHFlexGrid2.TextMatrix(I, 11), "dd-mmm-yyyy")
    oWS.Cells(mRow, c + 2).Font.Size = 9
    
    oWS.Cells(mRow, c + 3) = "Storage To"
    oWS.Cells(mRow, c + 3).Font.Size = 9
    oWS.Cells(mRow, c + 3).Font.Bold = True
    oWS.Cells(mRow, c + 3).HorizontalAlignment = -4108
    oWS.Cells(mRow, c + 3).WrapText = True
    
    oWS.Cells(mRow, c + 4) = Format(MSHFlexGrid2.TextMatrix(I, 12), "dd-mmm-yyyy")
    oWS.Cells(mRow, c + 4).Font.Size = 9
    
    
    mRow = mRow + 1
    
    tmp = "a" & mRow & ":" & "c" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, c + 4) = "Rs." & " " & MSHFlexGrid2.TextMatrix(I, 15)
    oWS.Cells(mRow, c + 4).Font.Size = 9
    oWS.Cells(mRow, c + 4).Font.Bold = True
    
    mRow = mRow + 1
    tmp = "a" & mRow & ":" & "b" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, c + 3) = "Hologram No:-"
    oWS.Cells(mRow, c + 3).Font.Size = 9
    oWS.Cells(mRow, c + 3).Font.Bold = True
    
    oWS.Cells(mRow, c + 4) = ""
    oWS.Cells(mRow, c + 4).Font.Size = 9
    oWS.Cells(mRow, c + 4).RowHeight = 35
    
    mRow = mRow + 1
    
    oWS.Cells(mRow, c + 1) = "Date"
    oWS.Cells(mRow, c + 1).Font.Size = 9
    oWS.Cells(mRow, c + 1).Font.Bold = True
    
    
    oWS.Cells(mRow, c + 2) = Format(CDCDate.Value, "dd-mmm-yyyy")
    oWS.Cells(mRow, c + 2).Font.Size = 9
    
    oWS.Cells(mRow, c + 3) = "Signature & Stamp of Warehouse Manager"
    oWS.Cells(mRow, c + 3).Font.Size = 9
    oWS.Cells(mRow, c + 3).Font.Bold = True
    
    oWS.Cells(mRow, c + 4) = ""
    oWS.Cells(mRow, c + 4).Font.Size = 9
    oWS.Cells(mRow, c + 4).RowHeight = 24
    mRow = mRow + 1
    
    k = mRow
    
    For Cell = 1 To k
        tmp = "a" & Cell & ":" & "d" & Cell
        oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
        oWS.Range(tmp).Font.Name = "Zurich BT"
        'oWS.Range(tmp).Font.Size = 9
        oWS.Range(tmp).WrapText = True
    Next Cell
    Call Log_Print_History(TxtPreFix, "CDC", MSHFlexGrid2.TextMatrix(I, 16), Gen_FLName)
Next

If Gen_ExcelPassword <> "" Then
   oWS.Protect Gen_ExcelPassword, , , , , , True, True
End If
oWB.Save


MsgBox ("Excel file created !!! ")
oXL.Visible = True

End Sub

Private Sub CmbPreFix_GotFocus()
tmp = CmbPreFix.Text

'Call TableMst_StoreManLocations("where " & Gen_WcSMLocation)

If Gen_WcLocation = "" Then
   Call TableMst_StoreManLocations("")
Else
   Call TableMst_StoreManLocations("where " & Gen_WcSMLocation)
End If


CmbPreFix.Clear
If RsMst_StoreManLocations.RecordCount = 0 Then
   MsgBox "No Bank found"
   CmbPreFix.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_StoreManLocations.RecordCount
    CmbPreFix.AddItem RsMst_StoreManLocations!StoreManLocations

    RsMst_StoreManLocations.MoveNext
Next
CmbPreFix.Text = tmp

End Sub

Private Sub CmbPreFix_LostFocus()
If CmbPreFix.Text = "" Then
   Exit Sub
End If
RsMst_StoreManLocations.MoveFirst
RsMst_StoreManLocations.Find "StoreManLocations = '" & CmbPreFix.Text & "'"
If RsMst_StoreManLocations.EOF Then
   MsgBox "Invalid PreFix Selection !!!!"
   CmbPreFix.SetFocus
   Exit Sub
End If
msSM_prefix = RsMst_StoreManLocations!SM_Prefix
mSLocationId = RsMst_StoreManLocations!StoreManLocationsID
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
Call TableMst_CMP(" Where LocationID = " & mSLocationId & "" & tmp1)
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
mSCMPName = RsMst_CMP!CMPID
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

Private Sub CmbSGodown_GotFocus()
If CmbPreFix.Text = "" Then
   MsgBox "Plase Select Location !!!"
   CmbPreFix.SetFocus
   Exit Sub
End If
If CmbSCmp.Text = "" Then
   MsgBox "Select CMP first !!! "
   CmbSCmp.SetFocus
   Exit Sub
End If
tmp = CmbSGodown.Text
Call TableMst_Godown("Where CMPID = " & mSCMPName)
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
mSGodownID = RsMst_Godown!godownid
End Sub

Private Sub GetBGSLID(x As Variant, mBGSLID As Variant)

tmp = "Select * from Mst_GSL Where SM_Prefix = '" & TxtPreFix.Text & "'"
Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
L:   TmpRs1.MoveFirst
     TmpRs1.Find "GSLID = " & mBGSLID
     If TmpRs1.EOF = False Then
        If IsNull(TmpRs1!BGSLID) = True Then
           mBackGSL = TmpRs1!GSLID
        Else
           mBGSLID = TmpRs1!BGSLID
           GoTo L
        End If
     End If
End If

End Sub
