VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmWithdrawal 
   Caption         =   "Withdrawal Transaction Detail"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame FrmInvoice 
      BackColor       =   &H008080FF&
      Caption         =   "Enter BOE Invoice Number."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   6000
      TabIndex        =   193
      Top             =   4560
      Visible         =   0   'False
      Width           =   5895
      Begin VB.OptionButton OptPostNo 
         BackColor       =   &H008080FF&
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
         Left            =   120
         TabIndex        =   198
         Top             =   720
         Width           =   2175
      End
      Begin VB.OptionButton OptTINo 
         BackColor       =   &H008080FF&
         Caption         =   "TI No. (Voucher No)"
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
         TabIndex        =   197
         Top             =   360
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.TextBox TxtInvNo 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   2760
         MaxLength       =   25
         TabIndex        =   196
         Top             =   480
         Width           =   2880
      End
      Begin VB.CommandButton CmdInvOK 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Approve"
         Height          =   400
         Left            =   2760
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   195
         Top             =   960
         Width           =   1440
      End
      Begin VB.CommandButton CmdInvCancel 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cancel"
         Height          =   400
         Left            =   4200
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   194
         Top             =   960
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "GSL Detail for Balance Check due to excess bags delivery from GSL Stock"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3000
      Left            =   360
      TabIndex        =   169
      Top             =   5160
      Visible         =   0   'False
      Width           =   14535
      Begin VB.CommandButton CmdGSLChkOK 
         Caption         =   "Ok"
         Height          =   375
         Left            =   13425
         TabIndex        =   171
         Top             =   8235
         Width           =   1215
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid7 
         Height          =   2475
         Left            =   120
         TabIndex        =   170
         Top             =   360
         Width           =   14175
         _ExtentX        =   25003
         _ExtentY        =   4366
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
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
      Height          =   3015
      Left            =   6480
      TabIndex        =   110
      Top             =   4680
      Visible         =   0   'False
      Width           =   4695
      Begin VB.CheckBox Check3 
         BackColor       =   &H008080FF&
         Caption         =   "Generate PDF"
         Height          =   255
         Left            =   240
         TabIndex        =   137
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
         TabIndex        =   113
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
         TabIndex        =   112
         Top             =   840
         Width           =   1440
      End
      Begin VB.TextBox TxtHologramNo 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1440
         MaxLength       =   25
         TabIndex        =   111
         Top             =   360
         Width           =   2880
      End
      Begin VB.Label Label48 
         Alignment       =   2  'Center
         BackColor       =   &H008080FF&
         Caption         =   $"FrmWithdrawal.frx":0000
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1260
         Left            =   120
         TabIndex        =   143
         Top             =   1800
         Width           =   4380
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
         TabIndex        =   114
         Top             =   390
         Width           =   1260
      End
   End
   Begin VB.Frame Frame0 
      Height          =   1965
      Left            =   480
      TabIndex        =   95
      Top             =   4800
      Visible         =   0   'False
      Width           =   14655
      Begin VB.ComboBox CmbSSCMFlag 
         Height          =   315
         ItemData        =   "FrmWithdrawal.frx":0092
         Left            =   7890
         List            =   "FrmWithdrawal.frx":009C
         Sorted          =   -1  'True
         TabIndex        =   126
         Top             =   840
         Width           =   1005
      End
      Begin VB.TextBox TxtSTruckNo 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   9810
         MaxLength       =   25
         TabIndex        =   127
         Top             =   840
         Width           =   2550
      End
      Begin VB.ComboBox CmbSRejection 
         Height          =   315
         ItemData        =   "FrmWithdrawal.frx":00A9
         Left            =   6120
         List            =   "FrmWithdrawal.frx":00B3
         Sorted          =   -1  'True
         TabIndex        =   125
         Top             =   840
         Width           =   915
      End
      Begin VB.TextBox TxtSCxTFNo 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1650
         MaxLength       =   25
         TabIndex        =   123
         Top             =   840
         Width           =   1080
      End
      Begin VB.TextBox TxtSManualCxTFNo 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   4170
         MaxLength       =   25
         TabIndex        =   124
         Top             =   840
         Width           =   1080
      End
      Begin VB.ComboBox CmbPreFix 
         Height          =   315
         Left            =   3165
         Sorted          =   -1  'True
         TabIndex        =   119
         Top             =   450
         Width           =   1605
      End
      Begin VB.ComboBox CmbSCommodity 
         Height          =   315
         Left            =   9000
         Sorted          =   -1  'True
         TabIndex        =   121
         Top             =   450
         Width           =   2685
      End
      Begin VB.ComboBox CmbSExchangeType 
         Height          =   315
         Left            =   11730
         Sorted          =   -1  'True
         TabIndex        =   122
         Top             =   450
         Width           =   2805
      End
      Begin VB.ComboBox CmbSCmp 
         Height          =   315
         Left            =   4800
         Sorted          =   -1  'True
         TabIndex        =   120
         Top             =   450
         Width           =   4125
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "Search"
         Height          =   350
         Left            =   12600
         TabIndex        =   129
         Top             =   870
         Width           =   1000
      End
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   350
         Left            =   13560
         TabIndex        =   96
         Top             =   870
         Width           =   1005
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1455
         Left            =   120
         TabIndex        =   97
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
         Left            =   1680
         TabIndex        =   118
         TabStop         =   0   'False
         Top             =   450
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
         Format          =   120061953
         CurrentDate     =   36494
      End
      Begin MSComCtl2.DTPicker SFrom 
         Height          =   300
         Left            =   120
         TabIndex        =   117
         TabStop         =   0   'False
         Top             =   450
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
         Format          =   120061953
         CurrentDate     =   36494
      End
      Begin VB.ComboBox CmbSAgent 
         Height          =   315
         Left            =   2280
         Sorted          =   -1  'True
         TabIndex        =   128
         TabStop         =   0   'False
         Top             =   1305
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label Label65 
         Caption         =   "SCM Deposit"
         Height          =   435
         Left            =   7080
         TabIndex        =   215
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label52 
         Caption         =   "Truck No"
         Height          =   225
         Left            =   9105
         TabIndex        =   166
         Top             =   930
         Width           =   735
      End
      Begin VB.Label Label51 
         AutoSize        =   -1  'True
         Caption         =   "Rejected "
         Height          =   195
         Left            =   5400
         TabIndex        =   149
         Top             =   900
         Width           =   720
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label49 
         Caption         =   "Store Man CITF No"
         Height          =   225
         Left            =   120
         TabIndex        =   147
         Top             =   900
         Width           =   1530
      End
      Begin VB.Label Label50 
         Caption         =   "Manual CITF No"
         Height          =   225
         Left            =   3000
         TabIndex        =   146
         Top             =   930
         Width           =   1530
      End
      Begin VB.Label Label40 
         Caption         =   "Location ID"
         Height          =   300
         Left            =   3600
         TabIndex        =   136
         Top             =   240
         Width           =   960
      End
      Begin VB.Label Label37 
         Caption         =   "Agent Name"
         Height          =   300
         Left            =   1320
         TabIndex        =   135
         Top             =   1335
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.Label Label36 
         Caption         =   "Exchange Type"
         Height          =   225
         Left            =   12532
         TabIndex        =   134
         Top             =   240
         Width           =   1200
      End
      Begin VB.Label Label39 
         Caption         =   "Commodity"
         Height          =   300
         Left            =   9975
         TabIndex        =   133
         Top             =   240
         Width           =   960
      End
      Begin VB.Label Label34 
         Caption         =   "CMP"
         Height          =   300
         Left            =   6735
         TabIndex        =   132
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label44 
         Caption         =   "From Date"
         Height          =   255
         Left            =   420
         TabIndex        =   131
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label45 
         Caption         =   "To Date"
         Height          =   255
         Left            =   2100
         TabIndex        =   130
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      TabIndex        =   165
      Top             =   9960
      Width           =   10935
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   8895
      Left            =   15000
      TabIndex        =   164
      Top             =   240
      Width           =   255
   End
   Begin VB.PictureBox Frame4 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   9555
      Left            =   120
      ScaleHeight     =   9525
      ScaleWidth      =   14880
      TabIndex        =   0
      Top             =   0
      Width           =   14910
      Begin VB.Frame Frame5 
         Caption         =   "RRN NCDEX Invoice"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3000
         Left            =   960
         TabIndex        =   212
         Top             =   4920
         Visible         =   0   'False
         Width           =   14655
         Begin VB.CommandButton Command3 
            Caption         =   "Ok"
            Height          =   375
            Left            =   13425
            TabIndex        =   213
            Top             =   8235
            Width           =   1215
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid11 
            Height          =   2475
            Left            =   120
            TabIndex        =   214
            Top             =   360
            Width           =   14415
            _ExtentX        =   25426
            _ExtentY        =   4366
            _Version        =   393216
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H008080FF&
         Caption         =   "Approve Withdrawal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   5400
         TabIndex        =   173
         Top             =   4800
         Visible         =   0   'False
         Width           =   6495
         Begin VB.TextBox TxtARemark 
            Appearance      =   0  'Flat
            Height          =   600
            Left            =   1440
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   176
            Top             =   360
            Width           =   4920
         End
         Begin VB.CommandButton CmdApprovePrint 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Approve"
            Height          =   400
            Left            =   1920
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   175
            Top             =   1080
            Width           =   1440
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Cancel"
            Height          =   400
            Left            =   3360
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   174
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label Label67 
            BackColor       =   &H008080FF&
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
            Height          =   300
            Left            =   120
            TabIndex        =   177
            Top             =   390
            Width           =   900
         End
      End
      Begin VB.PictureBox Picture3 
         Height          =   945
         Left            =   120
         ScaleHeight     =   885
         ScaleWidth      =   14880
         TabIndex        =   150
         Top             =   8520
         Width           =   14940
         Begin VB.CommandButton CmdPrint 
            Caption         =   "&Print"
            Height          =   375
            Left            =   10380
            TabIndex        =   163
            Top             =   450
            Width           =   2310
         End
         Begin VB.CommandButton ExitCmd 
            Caption         =   "E&xit"
            Height          =   375
            Left            =   8325
            TabIndex        =   162
            Top             =   450
            Width           =   2055
         End
         Begin VB.CommandButton LastCmd 
            Caption         =   "&Last"
            Height          =   375
            Left            =   6270
            TabIndex        =   161
            Top             =   450
            Width           =   2055
         End
         Begin VB.CommandButton FirstCmd 
            Caption         =   "&First"
            Height          =   375
            Left            =   4215
            TabIndex        =   160
            Top             =   450
            Width           =   2055
         End
         Begin VB.CommandButton PreviousCmd 
            Caption         =   "&Previous"
            Height          =   375
            Left            =   2160
            TabIndex        =   159
            Top             =   450
            Width           =   2055
         End
         Begin VB.CommandButton NextCmd 
            Caption         =   "&Next"
            Height          =   375
            Left            =   105
            TabIndex        =   158
            Top             =   450
            Width           =   2055
         End
         Begin VB.CommandButton SearchCmd 
            Caption         =   "&Search"
            Height          =   375
            Left            =   8325
            TabIndex        =   157
            Top             =   75
            Width           =   2070
         End
         Begin VB.CommandButton DeleteCmd 
            Caption         =   "&Delete"
            Height          =   375
            Left            =   6270
            TabIndex        =   156
            Top             =   75
            Width           =   2055
         End
         Begin VB.CommandButton EditCmd 
            Caption         =   "&Edit"
            Height          =   375
            Left            =   4215
            TabIndex        =   155
            Top             =   75
            Width           =   2055
         End
         Begin VB.CommandButton SaveCmd 
            Caption         =   "Save"
            Height          =   375
            Left            =   2160
            TabIndex        =   154
            Top             =   75
            Width           =   2055
         End
         Begin VB.CommandButton AddCmd 
            Caption         =   "&Add New "
            Height          =   375
            Left            =   105
            TabIndex        =   153
            Top             =   75
            Width           =   2055
         End
         Begin VB.CommandButton cmdPrintPreview 
            Caption         =   "Print Preview"
            Height          =   375
            Left            =   10395
            TabIndex        =   152
            Top             =   75
            Width           =   2295
         End
         Begin VB.CommandButton CmdBOE 
            Caption         =   "Bill of Estimation"
            Height          =   750
            Left            =   12690
            TabIndex        =   151
            Top             =   75
            Width           =   1935
         End
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   6420
         Left            =   120
         TabIndex        =   39
         Top             =   2040
         Width           =   14655
         _ExtentX        =   25850
         _ExtentY        =   11324
         _Version        =   393216
         Tabs            =   4
         Tab             =   2
         TabsPerRow      =   4
         TabHeight       =   520
         ForeColor       =   128
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "RWC Detail"
         TabPicture(0)   =   "FrmWithdrawal.frx":00CE
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "PicRRN"
         Tab(0).Control(1)=   "MSHFlexGrid2"
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Truck Details"
         TabPicture(1)   =   "FrmWithdrawal.frx":00EA
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "MSHFlexGrid3"
         Tab(1).Control(1)=   "Pic_TTD"
         Tab(1).ControlCount=   2
         TabCaption(2)   =   "GSL Details"
         TabPicture(2)   =   "FrmWithdrawal.frx":0106
         Tab(2).ControlEnabled=   -1  'True
         Tab(2).Control(0)=   "MSHFlexGrid4"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "Pic_GCTD"
         Tab(2).Control(1).Enabled=   0   'False
         Tab(2).ControlCount=   2
         TabCaption(3)   =   "Stock / Invoice Details"
         TabPicture(3)   =   "FrmWithdrawal.frx":0122
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "FrmRRNDetails"
         Tab(3).Control(1)=   "FrameCRDetail"
         Tab(3).Control(2)=   "MSHFlexGrid9"
         Tab(3).Control(3)=   "TxtUpdated"
         Tab(3).Control(3).Enabled=   0   'False
         Tab(3).Control(4)=   "TxtCreated"
         Tab(3).Control(4).Enabled=   0   'False
         Tab(3).Control(5)=   "TxtBOEAmount"
         Tab(3).Control(5).Enabled=   0   'False
         Tab(3).Control(6)=   "CmdAPrint"
         Tab(3).Control(7)=   "MSHFlexGrid5"
         Tab(3).Control(8)=   "MSHFlexGrid8"
         Tab(3).Control(9)=   "MSHFlexGrid6"
         Tab(3).Control(10)=   "Label55"
         Tab(3).Control(11)=   "Label47"
         Tab(3).Control(12)=   "Label46"
         Tab(3).Control(13)=   "Label56"
         Tab(3).Control(14)=   "Label57"
         Tab(3).Control(15)=   "Label59"
         Tab(3).ControlCount=   16
         Begin VB.Frame FrmRRNDetails 
            Caption         =   "RRN Details"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5280
            Left            =   -74880
            TabIndex        =   209
            Top             =   480
            Visible         =   0   'False
            Width           =   12975
            Begin VB.CommandButton Command2 
               Caption         =   "Ok"
               Height          =   375
               Left            =   13425
               TabIndex        =   210
               Top             =   8235
               Width           =   1215
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid10 
               Height          =   7275
               Left            =   120
               TabIndex        =   211
               Top             =   360
               Width           =   14535
               _ExtentX        =   25638
               _ExtentY        =   12832
               _Version        =   393216
               WordWrap        =   -1  'True
               AllowUserResizing=   3
               _NumberOfBands  =   1
               _Band(0).Cols   =   2
            End
         End
         Begin VB.Frame FrameCRDetail 
            Height          =   525
            Left            =   -72480
            TabIndex        =   199
            Top             =   3360
            Width           =   5655
            Begin VB.Label LblVBatchID 
               AutoSize        =   -1  'True
               Caption         =   "-"
               Height          =   195
               Left            =   4560
               TabIndex        =   205
               Top             =   240
               Width           =   45
            End
            Begin VB.Label Label64 
               AutoSize        =   -1  'True
               Caption         =   "Batch ID:"
               Height          =   195
               Left            =   3720
               TabIndex        =   204
               Top             =   240
               Width           =   675
            End
            Begin VB.Label LblVPrefix 
               Caption         =   "-"
               Height          =   255
               Left            =   2880
               TabIndex        =   203
               Top             =   240
               Width           =   1335
            End
            Begin VB.Label Label62 
               AutoSize        =   -1  'True
               Caption         =   "Prefix :"
               Height          =   195
               Left            =   2280
               TabIndex        =   202
               Top             =   240
               Width           =   480
            End
            Begin VB.Label lblVCRID 
               AutoSize        =   -1  'True
               Caption         =   "-"
               Height          =   195
               Left            =   1440
               TabIndex        =   201
               Top             =   240
               Width           =   45
            End
            Begin VB.Label Label60 
               Caption         =   "Cash Receipt ID: "
               Height          =   255
               Left            =   120
               TabIndex        =   200
               Top             =   240
               Width           =   1335
            End
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid9 
            Height          =   2490
            Left            =   -74880
            TabIndex        =   185
            Top             =   3840
            Width           =   14460
            _ExtentX        =   25506
            _ExtentY        =   4392
            _Version        =   393216
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
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
            Left            =   -66060
            Locked          =   -1  'True
            TabIndex        =   183
            TabStop         =   0   'False
            Top             =   480
            Width           =   5595
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
            Left            =   -73260
            Locked          =   -1  'True
            TabIndex        =   182
            TabStop         =   0   'False
            Top             =   480
            Width           =   5730
         End
         Begin VB.TextBox TxtBOEAmount 
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
            Left            =   -73260
            Locked          =   -1  'True
            TabIndex        =   181
            TabStop         =   0   'False
            Top             =   900
            Width           =   1890
         End
         Begin VB.CommandButton CmdAPrint 
            BackColor       =   &H008080FF&
            Caption         =   "Approve Withdrawal"
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
            Left            =   -62625
            Style           =   1  'Graphical
            TabIndex        =   180
            Top             =   915
            Width           =   2190
         End
         Begin VB.PictureBox Pic_TTD 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   1815
            Left            =   -74940
            ScaleHeight     =   1785
            ScaleWidth      =   14385
            TabIndex        =   73
            Top             =   420
            Width           =   14415
            Begin VB.ComboBox CmbVehicleTypeID 
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
               Left            =   1875
               Sorted          =   -1  'True
               TabIndex        =   14
               Top             =   60
               Width           =   5295
            End
            Begin VB.TextBox TxtMaximum 
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
               Left            =   13035
               Locked          =   -1  'True
               MaxLength       =   75
               TabIndex        =   220
               Top             =   75
               Width           =   1245
            End
            Begin VB.TextBox TxtMinimum 
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
               Left            =   9465
               Locked          =   -1  'True
               MaxLength       =   75
               TabIndex        =   219
               Top             =   75
               Width           =   1200
            End
            Begin VB.TextBox TxtMoisturePer 
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
               Left            =   11805
               TabIndex        =   21
               Top             =   780
               Width           =   1575
            End
            Begin VB.TextBox TxtTTDNoOfBags 
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
               Left            =   10215
               TabIndex        =   20
               Top             =   780
               Width           =   1575
            End
            Begin VB.TextBox TxtTTDGrossWt 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
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
               TabIndex        =   19
               Top             =   780
               Width           =   1815
            End
            Begin VB.TextBox TxtTTDTruckNo 
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
               Top             =   780
               Width           =   2415
            End
            Begin VB.TextBox TxtWeighBridgNo 
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
               Left            =   2520
               TabIndex        =   16
               Top             =   780
               Width           =   2535
            End
            Begin VB.TextBox TxtTTDTotalWt 
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
               Left            =   5040
               TabIndex        =   17
               Top             =   780
               Width           =   1575
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
               Left            =   11880
               TabIndex        =   23
               Top             =   1320
               Width           =   1575
            End
            Begin VB.TextBox txtTTDRemarks 
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
               Height          =   495
               Left            =   2520
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   22
               Top             =   1200
               Width           =   9240
            End
            Begin VB.TextBox TxtTTDTareWt 
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
               Left            =   6600
               TabIndex        =   18
               Top             =   780
               Width           =   1815
            End
            Begin VB.Label Label71 
               Caption         =   "Vehicle Model"
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
               Left            =   0
               TabIndex        =   223
               Top             =   105
               Width           =   1875
            End
            Begin VB.Label Label70 
               Caption         =   "Max Tare Weight in MT"
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
               Left            =   10770
               TabIndex        =   222
               Top             =   135
               Width           =   2145
            End
            Begin VB.Label Label69 
               Caption         =   "Min Tare Weight in MT"
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
               Left            =   7275
               TabIndex        =   221
               Top             =   113
               Width           =   2100
            End
            Begin VB.Label Label66 
               Caption         =   "Moisture %"
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
               Left            =   11985
               TabIndex        =   216
               Top             =   465
               Width           =   1215
            End
            Begin VB.Label Label54 
               Caption         =   "No of Bags"
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
               Left            =   10395
               TabIndex        =   168
               Top             =   480
               Width           =   1215
            End
            Begin VB.Label Label18 
               Caption         =   "Gross Weight"
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
               Left            =   8520
               TabIndex        =   79
               Top             =   480
               Width           =   1455
            End
            Begin VB.Label Label30 
               Caption         =   "Truck No."
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
               Left            =   600
               TabIndex        =   78
               Top             =   480
               Width           =   1095
            End
            Begin VB.Label Label15 
               Caption         =   "Weighbridge Slip No."
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
               Left            =   2880
               TabIndex        =   77
               Top             =   480
               Width           =   1935
            End
            Begin VB.Label Label16 
               Caption         =   "Total Weight"
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
               Left            =   5160
               TabIndex        =   76
               Top             =   480
               Width           =   1335
            End
            Begin VB.Label Label17 
               Caption         =   "Tare Weight"
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
               Left            =   6840
               TabIndex        =   75
               Top             =   480
               Width           =   1335
            End
            Begin VB.Label Label6 
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
               Height          =   255
               Left            =   1320
               TabIndex        =   74
               Top             =   1320
               Width           =   855
            End
         End
         Begin VB.PictureBox Pic_GCTD 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   1935
            Left            =   120
            ScaleHeight     =   1905
            ScaleWidth      =   14430
            TabIndex        =   54
            Top             =   360
            Width           =   14460
            Begin VB.TextBox TxtAverageBagSize 
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
               Left            =   5040
               Locked          =   -1  'True
               TabIndex        =   226
               TabStop         =   0   'False
               Top             =   1500
               Width           =   1065
            End
            Begin VB.TextBox TxtMoisturePer1 
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
               Left            =   7095
               Locked          =   -1  'True
               TabIndex        =   217
               TabStop         =   0   'False
               Top             =   990
               Width           =   675
            End
            Begin VB.CommandButton CmdGCTDAdd2List 
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
               Left            =   12720
               TabIndex        =   32
               Top             =   1440
               Width           =   1605
            End
            Begin VB.CheckBox Check4 
               Caption         =   "Made Up"
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
               Left            =   7350
               TabIndex        =   206
               Top             =   1545
               Width           =   1170
            End
            Begin VB.ComboBox CmbGSLTruckNo 
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
               ItemData        =   "FrmWithdrawal.frx":013E
               Left            =   5235
               List            =   "FrmWithdrawal.frx":0140
               Sorted          =   -1  'True
               TabIndex        =   27
               Top             =   540
               Width           =   2550
            End
            Begin VB.TextBox TxtWHRGenQty 
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
               Left            =   11400
               Locked          =   -1  'True
               TabIndex        =   141
               TabStop         =   0   'False
               Top             =   1920
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.TextBox TxtWHRGenBags 
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
               Left            =   11760
               Locked          =   -1  'True
               TabIndex        =   140
               TabStop         =   0   'False
               Top             =   1920
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.TextBox txtmGSLClientIDRow 
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
               Left            =   11280
               Locked          =   -1  'True
               TabIndex        =   138
               TabStop         =   0   'False
               Top             =   2040
               Visible         =   0   'False
               Width           =   1185
            End
            Begin VB.TextBox TxtGSLAvailableNofBags 
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
               Left            =   3900
               Locked          =   -1  'True
               TabIndex        =   63
               TabStop         =   0   'False
               Top             =   540
               Width           =   1320
            End
            Begin VB.TextBox TxtGCTDNoofBags 
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
               Left            =   7800
               TabIndex        =   28
               Top             =   540
               Width           =   1335
            End
            Begin VB.TextBox TxtGSLPledgedWeight 
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
               Left            =   3195
               Locked          =   -1  'True
               TabIndex        =   106
               TabStop         =   0   'False
               Top             =   1500
               Width           =   1065
            End
            Begin VB.TextBox TxtGSLPledgedNofBags 
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
               Left            =   1455
               Locked          =   -1  'True
               TabIndex        =   104
               TabStop         =   0   'False
               Top             =   1470
               Width           =   825
            End
            Begin VB.ComboBox CmbGCTDClient 
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
               ItemData        =   "FrmWithdrawal.frx":0142
               Left            =   1395
               List            =   "FrmWithdrawal.frx":0144
               TabIndex        =   30
               Top             =   990
               Width           =   4575
            End
            Begin VB.TextBox TxtAdjWt 
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
               Left            =   11400
               TabIndex        =   33
               Top             =   1920
               Visible         =   0   'False
               Width           =   1560
            End
            Begin VB.CheckBox Check2 
               Caption         =   "Spillage"
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
               Left            =   6255
               TabIndex        =   98
               Top             =   1545
               Width           =   1170
            End
            Begin VB.ComboBox CmbLot 
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
               ItemData        =   "FrmWithdrawal.frx":0146
               Left            =   2537
               List            =   "FrmWithdrawal.frx":0148
               Sorted          =   -1  'True
               TabIndex        =   26
               Top             =   540
               Width           =   1350
            End
            Begin VB.ComboBox CmbStack 
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
               ItemData        =   "FrmWithdrawal.frx":014A
               Left            =   1336
               List            =   "FrmWithdrawal.frx":014C
               Sorted          =   -1  'True
               TabIndex        =   25
               Top             =   540
               Width           =   1185
            End
            Begin VB.ComboBox CmbGodownNo 
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
               ItemData        =   "FrmWithdrawal.frx":014E
               Left            =   30
               List            =   "FrmWithdrawal.frx":0150
               Sorted          =   -1  'True
               TabIndex        =   24
               Top             =   540
               Width           =   1290
            End
            Begin VB.TextBox TxtGSLISINno 
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
               Left            =   1395
               Locked          =   -1  'True
               TabIndex        =   71
               TabStop         =   0   'False
               Top             =   990
               Width           =   4575
            End
            Begin VB.TextBox TxtGSLWeightGainLoss 
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
               Left            =   13110
               Locked          =   -1  'True
               TabIndex        =   69
               TabStop         =   0   'False
               Top             =   540
               Visible         =   0   'False
               Width           =   1245
            End
            Begin VB.TextBox TxtGSLNetWeight 
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
               Left            =   11783
               Locked          =   -1  'True
               TabIndex        =   67
               TabStop         =   0   'False
               Top             =   540
               Width           =   1305
            End
            Begin VB.TextBox TxtGSLAvailableGrossWgt 
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
               Left            =   11640
               Locked          =   -1  'True
               TabIndex        =   65
               TabStop         =   0   'False
               Top             =   1920
               Visible         =   0   'False
               Width           =   1590
            End
            Begin VB.TextBox TxtGCTDremarks 
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
               Left            =   8880
               MaxLength       =   255
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   31
               Top             =   975
               Width           =   5490
            End
            Begin VB.TextBox TxtGCTDGrossWt 
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
               Left            =   9156
               TabIndex        =   29
               Top             =   540
               Width           =   1500
            End
            Begin VB.TextBox TxtGGunnyWeight 
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
               Left            =   10672
               Locked          =   -1  'True
               TabIndex        =   55
               TabStop         =   0   'False
               Top             =   540
               Width           =   1095
            End
            Begin VB.Label Label73 
               Caption         =   "Bag Size"
               Height          =   375
               Left            =   4560
               TabIndex        =   227
               Top             =   1440
               Width           =   555
            End
            Begin VB.Label Label68 
               Caption         =   "Moisture %"
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
               Left            =   6015
               TabIndex        =   218
               Top             =   990
               Width           =   1095
            End
            Begin VB.Label Label63 
               Caption         =   "Lots in yellow color are reservation lots"
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
               Left            =   9165
               TabIndex        =   208
               Top             =   1545
               Width           =   3495
            End
            Begin VB.Label Label61 
               BackColor       =   &H0000FFFF&
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
               Left            =   8670
               TabIndex        =   207
               Top             =   1545
               Width           =   435
            End
            Begin VB.Label Label58 
               Caption         =   "Truck No."
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
               Left            =   5850
               TabIndex        =   172
               Top             =   120
               Width           =   1095
            End
            Begin VB.Label Label11 
               Caption         =   "Pledged Wt."
               Height          =   375
               Left            =   2475
               TabIndex        =   107
               Top             =   1500
               Width           =   795
            End
            Begin VB.Label Label9 
               Caption         =   "Pledged Bags"
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
               Left            =   75
               TabIndex        =   105
               Top             =   1515
               Width           =   1380
            End
            Begin VB.Label LblClient_WHR_RRN 
               Caption         =   "Client Name ~ DP Name"
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
               Left            =   45
               TabIndex        =   101
               Top             =   960
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.Label Label1 
               Caption         =   "Adjusted Weight"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   11640
               TabIndex        =   99
               Top             =   1920
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.Label Label32 
               Caption         =   "ISIN No."
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
               Left            =   75
               TabIndex        =   72
               Top             =   1035
               Width           =   1080
            End
            Begin VB.Label Label31 
               Alignment       =   2  'Center
               Caption         =   "Weight  Gain / Loss"
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
               Left            =   13200
               TabIndex        =   70
               Top             =   15
               Visible         =   0   'False
               Width           =   1185
            End
            Begin VB.Label Label26 
               Alignment       =   2  'Center
               Caption         =   "Electronic Weight"
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
               Left            =   11790
               TabIndex        =   68
               Top             =   30
               Width           =   1410
            End
            Begin VB.Label Label13 
               Caption         =   "Available Electronic Weight"
               Height          =   270
               Left            =   10560
               TabIndex        =   66
               Top             =   1920
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.Label Label12 
               Caption         =   "Available Bags"
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
               Left            =   3900
               TabIndex        =   64
               Top             =   120
               Width           =   1395
            End
            Begin VB.Label Label25 
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
               Height          =   255
               Left            =   7965
               TabIndex        =   62
               Top             =   1035
               Width           =   975
            End
            Begin VB.Label Label24 
               Alignment       =   2  'Center
               Caption         =   "Delivered Gross Weight"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   525
               Left            =   9180
               TabIndex        =   61
               Top             =   30
               Width           =   1455
            End
            Begin VB.Label Label23 
               Alignment       =   2  'Center
               Caption         =   "Delivered Bags"
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
               Left            =   7680
               TabIndex        =   60
               Top             =   30
               Width           =   1560
            End
            Begin VB.Label Label22 
               Caption         =   "Lot No"
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
               Left            =   2845
               TabIndex        =   59
               Top             =   135
               Width           =   735
            End
            Begin VB.Label Label21 
               Caption         =   "Stack No"
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
               Left            =   1410
               TabIndex        =   58
               Top             =   120
               Width           =   975
            End
            Begin VB.Label Label20 
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
               Height          =   255
               Left            =   60
               TabIndex        =   57
               Top             =   120
               Width           =   1215
            End
            Begin VB.Label Label27 
               Alignment       =   2  'Center
               Caption         =   "Gunny Weight"
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
               Left            =   10755
               TabIndex        =   56
               Top             =   15
               Width           =   855
            End
         End
         Begin VB.PictureBox PicRRN 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   1635
            Left            =   -74940
            ScaleHeight     =   1605
            ScaleWidth      =   14490
            TabIndex        =   40
            Top             =   465
            Width           =   14520
            Begin VB.ComboBox CmbRRNISIN 
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
               ItemData        =   "FrmWithdrawal.frx":0152
               Left            =   1920
               List            =   "FrmWithdrawal.frx":0154
               Sorted          =   -1  'True
               TabIndex        =   148
               Top             =   360
               Visible         =   0   'False
               Width           =   1935
            End
            Begin VB.TextBox TxtRRNBags 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               Height          =   375
               Left            =   12480
               Locked          =   -1  'True
               TabIndex        =   103
               TabStop         =   0   'False
               Top             =   720
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.TextBox TxtBalanceBags 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               Height          =   375
               Left            =   13965
               Locked          =   -1  'True
               TabIndex        =   102
               TabStop         =   0   'False
               Top             =   720
               Visible         =   0   'False
               Width           =   1440
            End
            Begin VB.ComboBox CmbClientDP 
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
               Left            =   6300
               Sorted          =   -1  'True
               TabIndex        =   12
               Top             =   1080
               Visible         =   0   'False
               Width           =   6105
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
               Left            =   8025
               Locked          =   -1  'True
               TabIndex        =   91
               TabStop         =   0   'False
               Top             =   1080
               Width           =   4350
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
               Left            =   1800
               Locked          =   -1  'True
               TabIndex        =   52
               TabStop         =   0   'False
               Top             =   1080
               Width           =   4470
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
               Left            =   6285
               Locked          =   -1  'True
               TabIndex        =   50
               TabStop         =   0   'False
               Top             =   1080
               Width           =   1725
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
               ForeColor       =   &H80000012&
               Height          =   375
               Left            =   30
               Locked          =   -1  'True
               TabIndex        =   11
               TabStop         =   0   'False
               Top             =   1080
               Width           =   1755
            End
            Begin VB.TextBox TxtRRNRemarks 
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
               Height          =   390
               Left            =   7590
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   10
               Top             =   360
               Width           =   6840
            End
            Begin VB.CommandButton CmdAdd2ListRRN 
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
               Height          =   405
               Left            =   12480
               TabIndex        =   13
               Top             =   1065
               Width           =   1935
            End
            Begin VB.TextBox TxtBalanceQty 
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
               Left            =   5760
               Locked          =   -1  'True
               TabIndex        =   43
               TabStop         =   0   'False
               Top             =   360
               Width           =   1800
            End
            Begin VB.TextBox TxtRRNNo 
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
               Left            =   0
               Locked          =   -1  'True
               TabIndex        =   9
               Top             =   360
               Width           =   1830
            End
            Begin VB.TextBox TxtISINno 
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
               Left            =   1890
               Locked          =   -1  'True
               TabIndex        =   42
               TabStop         =   0   'False
               Top             =   360
               Width           =   2010
            End
            Begin VB.TextBox TxtRRNQty 
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
               Left            =   3915
               Locked          =   -1  'True
               TabIndex        =   41
               TabStop         =   0   'False
               Top             =   360
               Width           =   1815
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
               Left            =   7380
               TabIndex        =   100
               Top             =   795
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.Label LblDPName 
               Caption         =   "DP Name"
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
               Left            =   8685
               TabIndex        =   92
               Top             =   810
               Width           =   1695
            End
            Begin VB.Label Label10 
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
               Left            =   3450
               TabIndex        =   53
               Top             =   795
               Width           =   1095
            End
            Begin VB.Label LblDPID 
               Caption         =   "DP-ID"
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
               Left            =   6525
               TabIndex        =   51
               Top             =   795
               Width           =   1335
            End
            Begin VB.Label Label8 
               Caption         =   "Client ID"
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
               Left            =   600
               TabIndex        =   49
               Top             =   810
               Width           =   855
            End
            Begin VB.Label Label42 
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
               Height          =   255
               Left            =   10485
               TabIndex        =   48
               Top             =   90
               Width           =   975
            End
            Begin VB.Label Label38 
               Caption         =   "Balance Qty"
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
               Left            =   6158
               TabIndex        =   47
               Top             =   75
               Width           =   975
            End
            Begin VB.Label LblRRNQty 
               Caption         =   "RRN Qty"
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
               Left            =   4350
               TabIndex        =   46
               Top             =   75
               Width           =   1545
            End
            Begin VB.Label LblISINNo 
               Caption         =   "CMSE Lot  No."
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
               Left            =   2318
               TabIndex        =   45
               Top             =   105
               Width           =   1155
            End
            Begin VB.Label LblRRNNo 
               Alignment       =   2  'Center
               Caption         =   "RRN No"
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
               Left            =   30
               TabIndex        =   44
               Top             =   105
               Width           =   1815
            End
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid4 
            Height          =   3885
            Left            =   90
            TabIndex        =   142
            Top             =   2430
            Width           =   14535
            _ExtentX        =   25638
            _ExtentY        =   6853
            _Version        =   393216
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
            Height          =   3930
            Left            =   -74880
            TabIndex        =   144
            Top             =   2280
            Width           =   14460
            _ExtentX        =   25506
            _ExtentY        =   6932
            _Version        =   393216
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
            Height          =   3915
            Left            =   -74925
            TabIndex        =   145
            Top             =   2340
            Width           =   14415
            _ExtentX        =   25426
            _ExtentY        =   6906
            _Version        =   393216
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid5 
            Height          =   1725
            Left            =   -74160
            TabIndex        =   179
            Top             =   4440
            Visible         =   0   'False
            Width           =   6615
            _ExtentX        =   11668
            _ExtentY        =   3043
            _Version        =   393216
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid8 
            Height          =   1890
            Left            =   -74865
            TabIndex        =   184
            Top             =   1560
            Width           =   14460
            _ExtentX        =   25506
            _ExtentY        =   3334
            _Version        =   393216
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid6 
            Height          =   1365
            Left            =   -74760
            TabIndex        =   192
            Top             =   4800
            Visible         =   0   'False
            Width           =   7215
            _ExtentX        =   12726
            _ExtentY        =   2408
            _Version        =   393216
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin VB.Label Label55 
            Caption         =   "Stock Details:"
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
            Left            =   -74865
            TabIndex        =   191
            Top             =   1275
            Width           =   1575
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
            Left            =   -67305
            TabIndex        =   190
            Top             =   540
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
            Left            =   -74745
            TabIndex        =   189
            Top             =   540
            Width           =   1695
         End
         Begin VB.Label Label56 
            Caption         =   "Outstanding Details:"
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
            Left            =   -74880
            TabIndex        =   188
            Top             =   3510
            Width           =   2175
         End
         Begin VB.Label Label57 
            Caption         =   "BOE Amount:"
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
            Left            =   -74745
            TabIndex        =   187
            Top             =   960
            Width           =   1695
         End
         Begin VB.Label Label59 
            Caption         =   $"FrmWithdrawal.frx":0156
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   540
            Left            =   -71145
            TabIndex        =   186
            Top             =   960
            Width           =   8415
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1860
         Left            =   120
         ScaleHeight     =   1830
         ScaleWidth      =   14625
         TabIndex        =   36
         Top             =   120
         Width           =   14655
         Begin VB.TextBox TxtSpotPrice 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2057
               SubFormatType   =   0
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
            Left            =   7320
            Locked          =   -1  'True
            MaxLength       =   18
            TabIndex        =   224
            Top             =   1350
            Width           =   1815
         End
         Begin VB.TextBox TxtGUID 
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
            Left            =   14160
            Locked          =   -1  'True
            TabIndex        =   178
            TabStop         =   0   'False
            Top             =   1440
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.CheckBox ChkSCMFlag 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   5280
            TabIndex        =   7
            Top             =   1410
            Width           =   225
         End
         Begin VB.TextBox TxtManualCDTFNo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2057
               SubFormatType   =   0
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
            Left            =   1200
            Locked          =   -1  'True
            MaxLength       =   18
            TabIndex        =   2
            Top             =   1350
            Width           =   1815
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
            Left            =   11400
            TabIndex        =   115
            TabStop         =   0   'False
            Top             =   960
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
            Left            =   5280
            Sorted          =   -1  'True
            TabIndex        =   6
            Top             =   930
            Width           =   6015
         End
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
            Left            =   11445
            Locked          =   -1  'True
            TabIndex        =   108
            TabStop         =   0   'False
            Top             =   930
            Visible         =   0   'False
            Width           =   975
         End
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
            ItemData        =   "FrmWithdrawal.frx":01E3
            Left            =   1215
            List            =   "FrmWithdrawal.frx":01E5
            Sorted          =   -1  'True
            TabIndex        =   1
            Top             =   930
            Width           =   1815
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
            Left            =   5280
            Sorted          =   -1  'True
            TabIndex        =   5
            Top             =   510
            Width           =   6015
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
            Left            =   5280
            Sorted          =   -1  'True
            TabIndex        =   4
            Top             =   90
            Width           =   6015
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
            Left            =   5280
            ScrollBars      =   2  'Vertical
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   930
            Width           =   6000
         End
         Begin VB.TextBox TxtCGunnyWt 
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
            Left            =   13530
            Locked          =   -1  'True
            TabIndex        =   82
            TabStop         =   0   'False
            Top             =   90
            Width           =   975
         End
         Begin VB.TextBox TxtCLotSize 
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
            Left            =   13530
            Locked          =   -1  'True
            TabIndex        =   81
            TabStop         =   0   'False
            Top             =   510
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.TextBox TxtCVariation 
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
            Left            =   13530
            Locked          =   -1  'True
            TabIndex        =   80
            TabStop         =   0   'False
            Top             =   930
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Check1"
            Height          =   255
            Left            =   4020
            TabIndex        =   3
            Top             =   945
            Visible         =   0   'False
            Width           =   225
         End
         Begin VB.TextBox TxtCxTFNo 
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
            Left            =   1215
            Locked          =   -1  'True
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   90
            Width           =   1815
         End
         Begin MSComCtl2.DTPicker TxtCxTFDate 
            Height          =   360
            Left            =   1215
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   510
            Width           =   1815
            _ExtentX        =   3201
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
            CalendarBackColor=   -2147483648
            Format          =   120061953
            CurrentDate     =   39083
            MinDate         =   39083
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
            Height          =   360
            Left            =   5280
            Locked          =   -1  'True
            TabIndex        =   93
            TabStop         =   0   'False
            Top             =   120
            Width           =   6015
         End
         Begin VB.TextBox TxtCommodityID 
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
            Left            =   5280
            Locked          =   -1  'True
            TabIndex        =   94
            TabStop         =   0   'False
            Top             =   510
            Width           =   6015
         End
         Begin VB.TextBox TxtCxTFDispalyDate 
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
            Left            =   1215
            Locked          =   -1  'True
            TabIndex        =   139
            TabStop         =   0   'False
            Top             =   510
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.Label Label72 
            Caption         =   "Spot Price"
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
            TabIndex        =   225
            Top             =   1403
            Width           =   1095
         End
         Begin VB.Label Label53 
            Caption         =   "SCM Withdrawal"
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
            TabIndex        =   167
            Top             =   1455
            Width           =   1515
         End
         Begin VB.Label Label35 
            Caption         =   "Manual CITF #"
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
            Left            =   120
            TabIndex        =   116
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            Caption         =   "Location ID"
            Height          =   255
            Left            =   11520
            TabIndex        =   109
            Top             =   660
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label Label14 
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
            Left            =   4080
            TabIndex        =   90
            Top             =   563
            Width           =   1215
         End
         Begin VB.Label Label41 
            Alignment       =   1  'Right Justify
            Caption         =   "Lot Size"
            Height          =   255
            Left            =   11925
            TabIndex        =   89
            Top             =   570
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label Label5 
            Caption         =   "Txn Date"
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
            TabIndex        =   88
            Top             =   563
            Width           =   1215
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "Rejected "
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
            Left            =   3120
            TabIndex        =   87
            Top             =   945
            Visible         =   0   'False
            Width           =   840
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label28 
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
            Left            =   4590
            TabIndex        =   86
            Top             =   990
            Width           =   855
         End
         Begin VB.Label Label4 
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
            Height          =   255
            Left            =   4680
            TabIndex        =   85
            Top             =   150
            Width           =   615
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Gunny Wt. per Bag"
            Height          =   255
            Left            =   11925
            TabIndex        =   84
            Top             =   150
            Width           =   1455
         End
         Begin VB.Label Label43 
            Alignment       =   1  'Right Justify
            Caption         =   "% Variation "
            Height          =   255
            Left            =   11925
            TabIndex        =   83
            Top             =   990
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label Label7 
            Caption         =   "Mode"
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
            TabIndex        =   38
            Top             =   983
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "CITF No"
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
            Top             =   143
            Width           =   1215
         End
      End
   End
End
Attribute VB_Name = "FrmWithdrawal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg, Diff_wt As Variant
Dim mAgentID, mCommodityID, mContractID, mGodownID, mGSLID, mRRNinGCTD As Double
Dim mExchangeTypeID As Double
Dim mClientIDRow As Double
Dim mRWCID As Double
Dim mGSLClientIDRow As Double
Dim mRWCBalanceQty As Double
Dim mCMPName As Double
Dim mPledgeID, mTimeStamp, msSM_prefix As Variant
Dim PntPrv, mSLocationID As Variant
Dim BalBagNeg, mRejection As Variant
Dim mCommodityGroupId, mLocationID As Variant
Dim wc As Variant
Dim mCMPID As Double

'Declarations for PDF Report

Dim pgval As Double
Dim objPDF As New mjwPDF
Dim mLeft As Double
Dim mTop As Double
Dim mHeight As Double
Dim mRowHeight As Double
Dim mRow As Double
Dim mGlobalLeft As Double
Dim mTxtTop As Double
Dim mTxtLeft As Double
Dim mFontSize As Integer
Dim mTxtFromTop As Double
Dim mPDFFileName As Variant


Dim mLastDateofCDTF As Date
Dim mPostDepoFromDate As Date
Dim xRow As Variant
Dim mTruckWtRatio As Variant
Dim mASM_prefix As String
Dim mALocationID As Double
Dim mACommodityID As Double
Dim mNAVCode, mRNAVCode As Variant

Dim mMultOfAll, mMultOfAM As Variant
Dim mBODDate As Date
Dim mChkSameMonth As Boolean
Dim mTIID As Variant

Dim mTotOutstAmt As Variant
Dim mTotCommPrice, mTotBalQty, mtimes As Variant
Dim mVoucherNo As String
Dim mGridOutAmt As Variant
Dim mRejInvIDs As String

Dim mBGSLIDs As String
Dim Totpages As Double
Dim mGSLContractID As Variant
Dim mChilliGunnyWeight, mRRNCommTrackFlag As Variant
Dim mCalMoisturePer As Double
Dim I As Long
Dim mNewSM_Prefix As Variant
Dim mCommodityYearlyFlag, mGodownYearlyFlag As Variant
Dim mVehicleTypeID As Variant

Dim mBlockDays, mBlockMessage As Variant
Dim mBlockDate As Date

Dim mBlockingDays As Double

Dim mSpaceChargesForInvoice As Variant
Dim mWHRNo As Variant

Dim mCxTF_DetailsID As Variant

Dim mInvAmountOnPrint, mBOEIDNew As Variant

Private Sub Grid_Head_Error()

MSHFlexGrid5.Clear
MSHFlexGrid5.Rows = 3
MSHFlexGrid5.FixedRows = 2
MSHFlexGrid5.Cols = 7
MSHFlexGrid5.TextMatrix(1, 0) = ""
If mExchangeTypeID = 1 Then
   MSHFlexGrid5.TextMatrix(1, 1) = "Client Name"
   MSHFlexGrid5.TextMatrix(1, 2) = "WHR Balance Qty"
Else
   MSHFlexGrid5.TextMatrix(1, 1) = "CMSE Lot No" '"ISIN No"
   MSHFlexGrid5.TextMatrix(1, 2) = "RRN Balance Qty"
End If
MSHFlexGrid5.TextMatrix(1, 3) = "Delivered Qty (Electronic Weight)"
MSHFlexGrid5.TextMatrix(1, 4) = "Adjusted Weight"
MSHFlexGrid5.TextMatrix(1, 5) = "Excess Qty Delivered"
MSHFlexGrid5.TextMatrix(1, 6) = "Partial Qty Delivered"

MSHFlexGrid5.ColWidth(0) = 500
MSHFlexGrid5.ColWidth(1) = 3000
MSHFlexGrid5.ColWidth(2) = 2000
MSHFlexGrid5.ColWidth(3) = 2500
MSHFlexGrid5.ColWidth(4) = 1800
MSHFlexGrid5.ColWidth(5) = 1800
MSHFlexGrid5.ColWidth(6) = 1800


End Sub
Private Sub Grid_Head_GSLBalChk()

MSHFlexGrid7.Clear
MSHFlexGrid7.Rows = 2
MSHFlexGrid7.FixedRows = 1
MSHFlexGrid7.Cols = 5
MSHFlexGrid7.TextMatrix(0, 0) = "GSL ID"
MSHFlexGrid7.TextMatrix(0, 1) = "G~S~L"
MSHFlexGrid7.TextMatrix(0, 2) = "Balance Bags"
MSHFlexGrid7.TextMatrix(0, 3) = "Delivered Bags"
MSHFlexGrid7.TextMatrix(0, 4) = "Balance Bags"

MSHFlexGrid7.ColWidth(0) = 2200
MSHFlexGrid7.ColWidth(1) = 3000
MSHFlexGrid7.ColWidth(2) = 2000
MSHFlexGrid7.ColWidth(3) = 2500
MSHFlexGrid7.ColWidth(4) = 1800


End Sub

Private Sub Grid_Head_RRN()

MSHFlexGrid2.Clear
MSHFlexGrid2.Rows = 3
MSHFlexGrid2.FixedRows = 2
MSHFlexGrid2.Cols = 20

MSHFlexGrid2.TextMatrix(1, 0) = ""
MSHFlexGrid2.TextMatrix(1, 1) = "RRN No"
MSHFlexGrid2.TextMatrix(1, 2) = "CMSE Lot No" '"ISIN No"
MSHFlexGrid2.TextMatrix(1, 3) = "Client ID"
MSHFlexGrid2.TextMatrix(1, 4) = "Client Name"
MSHFlexGrid2.TextMatrix(1, 5) = "DP ID"
MSHFlexGrid2.TextMatrix(1, 6) = "DP Name"
MSHFlexGrid2.TextMatrix(1, 7) = "RRN Qty"
MSHFlexGrid2.TextMatrix(1, 8) = "Balance Qty (excluding this CITF)"
MSHFlexGrid2.TextMatrix(1, 9) = "Delivered Qty (Electronic Weight)"
MSHFlexGrid2.TextMatrix(1, 10) = "Remark"
MSHFlexGrid2.TextMatrix(1, 11) = "ClientRowID"
MSHFlexGrid2.TextMatrix(1, 12) = "R/W/C ID"
MSHFlexGrid2.TextMatrix(1, 13) = "WHR Bags"
MSHFlexGrid2.TextMatrix(1, 14) = "Balance Bags"
MSHFlexGrid2.TextMatrix(1, 15) = "Delivered Bags"
MSHFlexGrid2.TextMatrix(1, 16) = "Spillage/Madeup Delivered Bags"
MSHFlexGrid2.TextMatrix(1, 17) = "Spillage/Madeup Delivered Qty"
MSHFlexGrid2.TextMatrix(1, 18) = "NAVCode"
MSHFlexGrid2.TextMatrix(1, 19) = "RNAVCode"

CmdAdd2ListRRN.Enabled = True
MSHFlexGrid2.ColWidth(0) = 500
MSHFlexGrid2.ColWidth(1) = 1300
MSHFlexGrid2.ColWidth(2) = 1200
MSHFlexGrid2.ColWidth(3) = 1200
MSHFlexGrid2.ColWidth(4) = 1500
MSHFlexGrid2.ColWidth(5) = 1500
MSHFlexGrid2.ColWidth(6) = 1500
MSHFlexGrid2.ColWidth(10) = 1500
MSHFlexGrid2.ColWidth(11) = 0
MSHFlexGrid2.ColWidth(12) = 0
MSHFlexGrid2.ColWidth(13) = 0
MSHFlexGrid2.ColWidth(14) = 1000
MSHFlexGrid2.ColWidth(15) = 1000

MSHFlexGrid2.ColWidth(16) = 1000
MSHFlexGrid2.ColWidth(17) = 1000
MSHFlexGrid2.ColWidth(18) = 1000
MSHFlexGrid2.ColWidth(19) = 1000

MSHFlexGrid2.RowHeight(1) = 800

End Sub
Private Sub Grid_Head_TTD()
MSHFlexGrid3.Clear
MSHFlexGrid3.Rows = 3
MSHFlexGrid3.FixedRows = 2

MSHFlexGrid3.Cols = 13 '9
MSHFlexGrid3.TextMatrix(1, 0) = ""
MSHFlexGrid3.TextMatrix(1, 1) = "Truck No"
MSHFlexGrid3.TextMatrix(1, 2) = "Weighbridge Slip No"
MSHFlexGrid3.TextMatrix(1, 3) = "Total Weight"
MSHFlexGrid3.TextMatrix(1, 4) = "Tare Weight"
MSHFlexGrid3.TextMatrix(1, 5) = "Gross Weight"
MSHFlexGrid3.TextMatrix(1, 6) = "No Of Bags"
MSHFlexGrid3.TextMatrix(1, 7) = "Remark"
MSHFlexGrid3.TextMatrix(1, 8) = "Moisture %"

MSHFlexGrid3.TextMatrix(1, 9) = "Vehicle Model"
MSHFlexGrid3.TextMatrix(1, 10) = "Vehicle Type ID"
MSHFlexGrid3.TextMatrix(1, 11) = "Minimum tare weight"
MSHFlexGrid3.TextMatrix(1, 12) = "Maximum tare weight"


MSHFlexGrid3.ColWidth(0) = 0
MSHFlexGrid3.ColWidth(1) = 2000
MSHFlexGrid3.ColWidth(2) = 2500
MSHFlexGrid3.ColWidth(3) = 1500
MSHFlexGrid3.ColWidth(4) = 1500
MSHFlexGrid3.ColWidth(5) = 1500
MSHFlexGrid3.ColWidth(6) = 1500
MSHFlexGrid3.ColWidth(7) = 3800
MSHFlexGrid3.ColWidth(8) = 1500
End Sub
Private Sub Grid_Head_GCTD()

MSHFlexGrid4.Clear
MSHFlexGrid4.Rows = 3
MSHFlexGrid4.FixedRows = 2
MSHFlexGrid4.Cols = 24 '23 '22

MSHFlexGrid4.TextMatrix(1, 0) = ""
MSHFlexGrid4.TextMatrix(1, 1) = "Godown No"
MSHFlexGrid4.TextMatrix(1, 2) = "Stack No"
MSHFlexGrid4.TextMatrix(1, 3) = "Lot No"
MSHFlexGrid4.TextMatrix(1, 4) = "Available Bags"
MSHFlexGrid4.TextMatrix(1, 5) = "Available Ele. Wt."
MSHFlexGrid4.TextMatrix(1, 6) = "No of Bags"
MSHFlexGrid4.TextMatrix(1, 7) = "Gross Wt.    (A) "
'MSHFlexGrid4.TextMatrix(1, 8) = "Adjusted Wt. (B)"
MSHFlexGrid4.TextMatrix(1, 9) = "Gunny Weight (B)"
MSHFlexGrid4.TextMatrix(1, 10) = "Electronic Weight (A-B)" '"Electronic Weight (A+B-C)"
MSHFlexGrid4.TextMatrix(1, 11) = "Weight Gain/Loss"
MSHFlexGrid4.TextMatrix(1, 12) = "CMSE Lot No" '"ISIN No"
MSHFlexGrid4.TextMatrix(1, 13) = "Remark"
MSHFlexGrid4.TextMatrix(1, 14) = "GSLID"
MSHFlexGrid4.TextMatrix(1, 15) = "Godown Close"
MSHFlexGrid4.TextMatrix(1, 16) = "ClientIDRow"
MSHFlexGrid4.TextMatrix(1, 17) = "SpillageLot"
MSHFlexGrid4.TextMatrix(1, 18) = "Balance Bags"
MSHFlexGrid4.TextMatrix(1, 19) = "Truck No"
MSHFlexGrid4.TextMatrix(1, 20) = "MadeUp Lot"
MSHFlexGrid4.TextMatrix(1, 21) = "Moisture %"
MSHFlexGrid4.TextMatrix(1, 22) = "Yearly Billing Flag"
MSHFlexGrid4.TextMatrix(1, 23) = "Average Bag Size"

MSHFlexGrid4.ColWidth(0) = 0
MSHFlexGrid4.ColWidth(1) = 1200
MSHFlexGrid4.ColWidth(2) = 1200
MSHFlexGrid4.ColWidth(3) = 1200

MSHFlexGrid4.ColWidth(4) = 0
MSHFlexGrid4.ColWidth(5) = 0

MSHFlexGrid4.ColWidth(12) = 1500
MSHFlexGrid4.ColWidth(13) = 2000
MSHFlexGrid4.ColWidth(14) = 0
MSHFlexGrid4.ColWidth(15) = 0
MSHFlexGrid4.ColWidth(16) = 0
MSHFlexGrid4.ColWidth(11) = 0
MSHFlexGrid4.ColWidth(8) = 0
MSHFlexGrid4.ColWidth(19) = 1500

MSHFlexGrid4.RowHeight(1) = 800

End Sub
Private Sub AddCmd_Click()

Edit_Flg = "A"
Frame5.Visible = False
Call Grid_Head_RRN
Call Grid_Head_Error
Call Grid_Head_TTD
Call Grid_Head_GCTD
Call TableMst_RRN
Call Grid_Head_RRNDetail
mCommodityYearlyFlag = ""
CmdAPrint.Enabled = False
CmdPrintPreview.Enabled = False
TxtCxTFDate = CDate("01/01/2009")
TxtManualCDTFNo.Locked = False
TxtCMPID.Visible = False
TxtCommodityID.Visible = False
CmbCMPID.Visible = True
CmbCommodityID.Visible = True
CmbExchangeTypeID.Enabled = True
PicRRN.Enabled = True
Pic_GCTD.Enabled = True
Pic_TTD.Enabled = True
TxtRRNNo.Locked = False
TxtAgent.Locked = False
TxtCxTFNo = ""
TxtManualCDTFNo = ""
TxtAgent = ""
Check1.Value = 0
CmbExchangeTypeID.Text = ""
CmbCMPID.Text = ""
CmbCommodityID.Text = ""
TxtCxTFDate.Enabled = True
Check1.Enabled = True
Check1.Value = 0
ChkSCMFlag.Value = 0
TxtSpotPrice = ""
TxtSpotPrice.Locked = False
If RsTxn_CxTF_Details.RecordCount > 0 Then
   If CmbExchangeTypeID.Enabled = True Then
      CmbExchangeTypeID.SetFocus
   End If
End If
Call CheckRejection(Check1.Value)
Call GButtonLock(Me, False)
CmdPrint.Enabled = False
CmdBOE.Enabled = False
TxtAgent.Locked = False
CmbAgent.Visible = True
CmdAgent.Visible = True
TxtAgent.Visible = False
CmbAgent.Text = ""
TxtTTDGrossWt.Locked = True
ChkSCMFlag.Enabled = True
TxtCreated = ""
TxtUpdated = ""
TxtBOEAmount = ""
mTruckWtRatio = 0
Call Grid_Head_OutStand
Call Grid_Head_Stock
FrameCRDetail.Visible = False
Frame1.Visible = False

End Sub
Private Sub Check1_Click()
'Check1.Enabled = False
'SSTab1.Tab = 0
'MsgBox SSTab1.Caption

If mExchangeTypeID = 3 Then 'SPOT
   Call CheckRejection(Check1.Value)
   Call SetCaptions("CDC No", "CDC-No", "CDC-No", "CDC Qty", "TM ID", "Trading Member Name", True)
   Call CDCDetails(mCMPID, mCommodityID, TxtRRNNo.Text)
Else
   Call CheckRejection(Check1.Value)
   If Check1.Value = 1 Then
      Call SetCaptions("client details", "CDTF No", "Client Name", "Rejected Qty", "DP ID", "DP Name", True)
   Else
      Call SetCaptions("Remat Request No. (RRN)", "RRN No", "CMSE Lot No", "RRN Qty", "DP ID", "DP Name", True)
   End If
End If
End Sub
Private Sub CheckRejection(Value_)
'If mExchangeTypeID = 1 Then Exit Sub

Call Clear_RWCTabControl
Call Grid_Head_RRN

Label8.Top = 810
TxtClientID.Top = 1080
Label42.Top = 90
Label42.Left = 10485
TxtRRNRemarks.Top = 360
TxtRRNRemarks.Left = 7590

Label10.Top = Label8.Top
TxtClientName.Top = TxtClientID.Top
LblDPID.Top = Label8.Top
LblDPName.Top = Label8.Top
lblNameDid.Top = Label8.Top
CmbClientDP.Top = TxtClientID.Top
TxtDPID.Top = TxtClientID.Top
TxtDPName.Top = TxtClientID.Top

LblRRNQty.Top = 75
LblRRNQty.Left = 4470
TxtRRNQty.Top = 360
TxtRRNQty.Left = 3915

TxtBalanceQty.Top = 360
TxtBalanceQty.Left = 5760
Label38.Top = 75
Label38.Left = 6158


LblRRNNo.Visible = True
TxtRRNNo.Visible = True
LblISINNo.Visible = True
TxtISINno.Visible = True
TxtRRNQty.Visible = True
LblRRNQty.Visible = True
TxtBalanceQty.Visible = True
CmbGCTDClient.Visible = True
LblClient_WHR_RRN.Visible = False
TxtGSLISINno.Visible = True
Label38.Visible = True
TxtClientID.Locked = True
TxtClientID.BackColor = &H80000000
MSHFlexGrid2.ColWidth(1) = 1300
MSHFlexGrid2.ColWidth(2) = 1200
MSHFlexGrid2.ColWidth(7) = 1000
MSHFlexGrid2.ColWidth(8) = 1000
MSHFlexGrid2.ColWidth(9) = 1000
MSHFlexGrid2.ColWidth(10) = 1500


If mExchangeTypeID = 2 And Value_ = 1 Then
   SSTab1.TabCaption(0) = "Client Detail "
   LblRRNNo.Visible = False
   TxtRRNNo.Visible = False
   LblISINNo.Visible = False
   TxtISINno.Visible = False
   
   TxtRRNQty.Visible = True
   LblRRNQty.Visible = True
   
   TxtBalanceQty.Visible = True
   Label38.Visible = True
   
   CmbGCTDClient.Visible = True
   LblClient_WHR_RRN.Visible = True
   TxtGSLISINno.Visible = False
   
   TxtClientID.Locked = False
   TxtClientID.BackColor = &H80000005
   MSHFlexGrid2.ColWidth(1) = 0
   MSHFlexGrid2.ColWidth(2) = 0
'   MSHFlexGrid2.ColWidth(7) = 0
'   MSHFlexGrid2.ColWidth(8) = 0
'   MSHFlexGrid2.ColWidth(9) = 0
   MSHFlexGrid2.ColWidth(10) = 1500
   Label8.Top = 105
   TxtClientID.Top = 360
   Label42.Top = 810
   Label42.Left = 3450
   TxtRRNRemarks.Top = 1080
   TxtRRNRemarks.Left = 50
   TxtRRNRemarks.Width = 8500
   
   LblRRNQty.Top = 810
   LblRRNQty.Left = 9000
   TxtRRNQty.Top = 1080
   TxtRRNQty.Left = 8650
   
   Label38.Top = 810
   Label38.Left = 10500
   TxtBalanceQty.Top = 1080
   TxtBalanceQty.Left = 10500
   
   Label10.Top = Label8.Top
   
   TxtClientName.Top = TxtClientID.Top
   LblDPID.Top = Label8.Top
   LblDPName.Top = Label8.Top
   lblNameDid.Top = Label8.Top
   CmbClientDP.Top = TxtClientID.Top
   TxtDPID.Top = TxtClientID.Top
   TxtDPName.Top = TxtClientID.Top
    

ElseIf (mExchangeTypeID = 2 Or mExchangeTypeID = 1) And Value_ = 0 Then
       SSTab1.TabCaption(0) = "Remat Request No. (RRN)"
'       LblRRNNo.Visible = True
'       TxtRRNNo.Visible = True
'       LblISINNo.Visible = True
'       TxtISINno.Visible = True
'       TxtRRNQty.Visible = True
'       LblRRNQty.Visible = True
'       TxtBalanceQty.Visible = True
'       CmbGCTDClient.Visible = False
'       LblClient_WHR_RRN.Visible = False
'       TxtGSLISINno.Visible = True
'       Label38.Visible = True
'       TxtClientID.Locked = True
'       TxtClientID.BackColor = &H80000000
'       MSHFlexGrid2.ColWidth(1) = 1300
'       MSHFlexGrid2.ColWidth(2) = 1200
'       MSHFlexGrid2.ColWidth(7) = 1000
'       MSHFlexGrid2.ColWidth(8) = 1000
'       MSHFlexGrid2.ColWidth(9) = 1000
'       MSHFlexGrid2.ColWidth(10) = 1500
'       Label8.Top = 810
'       TxtClientID.Top = 1080
'       Label42.Top = 90
'       Label42.Left = 10485
'       TxtRRNRemarks.Top = 360
'       TxtRRNRemarks.Left = 7590
       
'
'       LblRRNQty.Top = 75
'       LblRRNQty.Left = 4470
'       TxtRRNQty.Top = 360
'       TxtRRNQty.Left = 3915
'
'       TxtBalanceQty.Top = 360
'       TxtBalanceQty.Left = 5760
'       Label38.Top = 75
'       Label38.Left = 6158
End If


End Sub

Private Sub Check2_Click()
'CmbGCTDClient.Visible = False
If Check4.Value = 1 Then
   MsgBox "Please un check Madeup!!!"
   Check2.Value = 0
   Exit Sub
End If

TxtGSLISINno.Visible = False

If Check2.Value = 1 Then
   'CmbGCTDClient.Visible = True
Else
   If mExchangeTypeID = 2 And Check1.Value = 1 Then
      'CmbGCTDClient.Visible = True
   Else
      TxtGSLISINno.Visible = True
   End If
End If

CmbLot.Text = ""

If mExchangeTypeID = 3 And Check2.Value = 0 Then
   CmbGodownNo.Enabled = False
   CmbStack.Enabled = False
   CmbLot.Enabled = False
   CmdGCTDAdd2List.Enabled = False
Else
   CmbGodownNo.Enabled = True
   CmbStack.Enabled = True
   CmbLot.Enabled = True
   CmdGCTDAdd2List.Caption = "Add to List"
   CmdGCTDAdd2List.Enabled = True
End If




End Sub

Private Sub Check4_Click()

If Check2.Value = 1 Then
   MsgBox "Please deselect Spillage Check!!!"
   Check4.Value = 0
   Exit Sub
End If

TxtGSLISINno.Visible = False

If Check4.Value = 1 Then
   'CmbGCTDClient.Visible = True
Else
   If mExchangeTypeID = 2 And Check1.Value = 1 Then
      'CmbGCTDClient.Visible = True
   Else
      TxtGSLISINno.Visible = True
   End If
End If

CmbLot.Text = ""

If mExchangeTypeID = 3 And Check4.Value = 0 Then
   CmbGodownNo.Enabled = False
   CmbStack.Enabled = False
   CmbLot.Enabled = False
   CmdGCTDAdd2List.Enabled = False
Else
   CmbGodownNo.Enabled = True
   CmbStack.Enabled = True
   CmbLot.Enabled = True
   CmdGCTDAdd2List.Caption = "Add to List"
   CmdGCTDAdd2List.Enabled = True
End If

End Sub

Private Sub ChkSCMFlag_Click()
Call Clear_RWCTabControl
Call Grid_Head_RRN
Call DisplayRWCModification
Call Clear_GCTDTabControl
Call Grid_Head_GCTD
End Sub

Private Sub CmbAgent_GotFocus()
tmp = CmbAgent.Text
Call TableMst_Agent
CmbAgent.Clear
If RsMst_Agent.RecordCount = 0 Then
   MsgBox "No record found!!!"
   Exit Sub
End If
For I = 1 To RsMst_Agent.RecordCount
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
   MsgBox "Invalid selection!!!"
   CmbAgent.SetFocus
   Exit Sub
End If
mAgentID = RsMst_Agent!AgentID
SSTab1.Tab = 0
If TxtRRNNo.Visible = True And TxtRRNNo.Enabled = True Then
   TxtRRNNo.SetFocus
End If
End Sub

Private Sub CmbCMPID_GotFocus()
tmp = CmbCMPID.Text

wc = ""
If Gen_WcCMP <> "" Then
   wc = "And " & Gen_WcCMP
End If

Call TableMst_CMP(" Where CloseDate is Null " & wc)
CmbCMPID.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "No record found!!!"
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
RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPName = '" & CmbCMPID.Text & "'"
If RsMst_CMP.EOF Then
   MsgBox "Invalid selection!!!"
   CmbCMPID.SetFocus
   Exit Sub
End If

If mCMPID <> RsMst_CMP!CMPID Then
   CmbCommodityID.Text = ""
End If


mCMPID = RsMst_CMP!CMPID



'If RsMst_CMP!CurrentDate + 1 >= RsMst_CMP!CMPLockedDate Then
'   MsgBox "Month end process is pending for this CMP, Please Call CMG Department to Execute Month End Process !!! "
'   CmbCMPID.SetFocus
'   Exit Sub
'End If


TxtPreFix = GetPreFix(RsMst_CMP!CMPID, "P")
If mExchangeTypeID = 1 Then 'NON NCDEX
   Call WHRDetails(mCMPID, mCommodityID, TxtRRNNo.Text)
End If
If mExchangeTypeID = 2 Then 'NCDEX
   Call RRNDetails(mCMPID, mCommodityID, TxtRRNNo.Text)
End If
If mExchangeTypeID = 3 Then 'SPOT
   Call CDCDetails(mCMPID, mCommodityID, TxtRRNNo.Text)
End If
Call Clear_GCTDTabControl
Call Grid_Head_GCTD
 


TxtCxTFDate.Value = IIf(IsNull(RsMst_CMP!CurrentDate), RsMst_CMP!StartDate, RsMst_CMP!CurrentDate) + 1

mBODDate = TxtCxTFDate.Value

TxtCxTFDispalyDate = TxtCxTFDate.Value

mBlockDays = RsMst_CMP!BlockDays
mBlockDate = Date - Val(mBlockDays)
mBlockMessage = "CITF Date less than BOD Date is not allowed!!!"

If mBODDate < mBlockDate Then
   TxtCxTFDate.Value = mBlockDate
   mBODDate = TxtCxTFDate.Value
   TxtCxTFDispalyDate = TxtCxTFDate.Value
   mBlockMessage = "CITF Date less than Block Date (which is back " & mBlockDays & " no. of days from todays date) is not allowed!!! "
End If

'==Commented on 22 May 2012
''TxtCxTFDate.Enabled = False


'If LCase(CmbCMPID.Text) = "unjha" Or LCase(CmbCMPID.Text) = "rajkot" Then
'   TxtCxTFDate.Enabled = True
'Else
'   TxtCxTFDate.Enabled = False
'End If

'If SaveCmd.Enabled = True Then
'   MsgBox "CITF Date is " & TxtCxTFDate
'End If

End Sub


'Private Sub CmbCommodityID_GotFocus()
'If CmbExchangeTypeID.Text <> "" Then
'   If CmbCMPID.Text <> "" Then
'      tmp = CmbCommodityID.Text
'      Call TableMst_Commodity
'      CmbCommodityID.Clear
'      If RsMst_Commodity.RecordCount = 0 Then
'         MsgBox "No record found!!!"
'         Exit Sub
'      End If
'      For I = 1 To RsMst_Commodity.RecordCount
'          CmbCommodityID.AddItem RsMst_Commodity!Commodity
'          RsMst_Commodity.MoveNext
'      Next
'      CmbCommodityID.Text = tmp
'   Else
'      MsgBox "Please select a CMP!!!"
'      CmbCMPID.SetFocus
'      Exit Sub
'   End If
'Else
'   MsgBox "Please select a Mode!!!"
'   CmbExchangeTypeID.SetFocus
'   Exit Sub
'End If
'End Sub
''----
'Private Sub CmbCommodityID_LostFocus()
'If CmbCommodityID.Text = "" Then
'   Exit Sub
'End If
'
'RsMst_Commodity.MoveFirst
'RsMst_Commodity.Find "Commodity = '" & CmbCommodityID.Text & "'"
'If RsMst_Commodity.EOF Then
'   MsgBox "Invalid selection!!!"
'   CmbCommodityID.SetFocus
'   Exit Sub
'End If
'mCommodityID = RsMst_Commodity!CommodityID
'Call TableMst_CommodityDetail("Where CommodityId = " & mCommodityID & " and ExchangeTypeId = " & mExchangeTypeID)
'If RsMst_CommodityDetail.RecordCount > 0 Then
'   TxtCGunnyWt = RsMst_CommodityDetail!GunnyWt
'   TxtCLotSize = IIf(IsNull(RsMst_CommodityDetail!LotSize), "", RsMst_CommodityDetail!LotSize)
'   TxtCVariation = IIf(IsNull(RsMst_CommodityDetail!Variation), "", RsMst_CommodityDetail!Variation)
'Else
'   MsgBox "No record found for Exchange" & CmbExchangeTypeID.Text & " - " & CmbCommodityID.Text
'   CmbCommodityID.SetFocus
'   Exit Sub
'End If
'Call Clear_RWCTabControl
'Call Grid_Head_RRN
'Call Clear_TTDTabControl
'Call Grid_Head_TTD
'Call Clear_GCTDTabControl
'Call Grid_Head_GCTD
'
''Call CheckRejection(Check1.Value)
''Call Check1_Click
'Call DisplayRWCModification
'
''If mExchangeTypeID = 1 Then 'NON NCDEX
''   Call WHRDetails(mCMPID, mCommodityID, TxtRRNNo.Text)
''End If
''If mExchangeTypeID = 2 Then 'NCDEX
''   Call RRNDetails(mCMPID, mCommodityID, TxtRRNNo.Text)
''End If
''If mExchangeTypeID = 3 Then 'SPOT
''   Call CDCDetails(mCMPID, mCommodityID, TxtRRNNo.Text)
''End If
'
''SSTab1.Tab = 0
'End Sub

'---------------------------

Private Sub CmbCommodityID_GotFocus()

If CmbExchangeTypeID.Text = "" Then
    MsgBox "Please Select Exchange Mode !!!!"
    CmbExchangeTypeID.SetFocus
    Exit Sub
End If

If CmbCMPID.Text = "" Then
    MsgBox "Please Select CMP!!!!!"
    CmbCMPID.SetFocus
    Exit Sub
End If

tmp1 = CmbCommodityID.Text
'Call TableMst_Commodity
If Edit_Flg = "" Then
   tmp = "Select MCom.CommodityID, MCom.Commodity,Mcom.YearlyBillingFlg From Mst_Commodity MCom Where MCom.CommodityID = " & mCommodityID
Else
   tmp = "Select MCom.CommodityID, MCom.Commodity,Mcom.YearlyBillingFlg From Mst_Commodity MCom Inner Join Mst_LocationCOmmodityMapping MLCM on MLCM.CommodityID = MCom.CommodityID "
   tmp = tmp & " Inner Join Mst_Location ML on ML.LocationID = MLCM.LocationID "
   tmp = tmp & " Where ML.SM_Prefix = '" & TxtPreFix & "'"
End If
Call Tmp7Table

CmbCommodityID.Clear
'If RsMst_Commodity.RecordCount = 0 Then
If TmpRs7.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
'For I = 1 To RsMst_Commodity.RecordCount
'    CmbCommodityID.AddItem RsMst_Commodity!Commodity
'    RsMst_Commodity.MoveNext
'Next
For I = 1 To TmpRs7.RecordCount
    CmbCommodityID.AddItem TmpRs7!Commodity
    TmpRs7.MoveNext
Next
CmbCommodityID.Text = tmp1

'If CmbExchangeTypeID.Text <> "" Then
'   If CmbCMPID.Text <> "" Then
'      tmp = CmbCommodityID.Text
'      Call TableMst_Commodity
'      CmbCommodityID.Clear
'      If RsMst_Commodity.RecordCount = 0 Then
'         MsgBox "No record found!!!"
'         Exit Sub
'      End If
'      For I = 1 To RsMst_Commodity.RecordCount
'          CmbCommodityID.AddItem RsMst_Commodity!Commodity
'          RsMst_Commodity.MoveNext
'      Next
'      CmbCommodityID.Text = tmp
'   Else
'      MsgBox "Please select a CMP!!!"
'      CmbCMPID.SetFocus
'      Exit Sub
'   End If
'Else
'   MsgBox "Please select a Mode!!!"
'   CmbExchangeTypeID.SetFocus
'   Exit Sub
'End If
End Sub
'----
Private Sub CmbCommodityID_LostFocus()
If CmbCommodityID.Text = "" Then Exit Sub

'RsMst_Commodity.MoveFirst
'RsMst_Commodity.Find "Commodity = '" & CmbCommodityID.Text & "'"
'If RsMst_Commodity.EOF Then
'   MsgBox "Invalid selection!!!"
'   CmbCommodityID.SetFocus
'   Exit Sub
'End If

TmpRs7.MoveFirst
TmpRs7.Find "Commodity = '" & CmbCommodityID.Text & "'"
If TmpRs7.EOF Then
   MsgBox "Invalid selection "
   CmbCommodityID.SetFocus
   Exit Sub
End If

mCommodityID = TmpRs7!CommodityID
'mCommodityID = RsMst_Commodity!CommodityID
Call TableMst_CommodityDetail("Where CommodityId = " & mCommodityID & " and ExchangeTypeId = " & mExchangeTypeID)
If RsMst_CommodityDetail.RecordCount > 0 Then
   TxtCGunnyWt = RsMst_CommodityDetail!GunnyWt
   TxtCLotSize = IIf(IsNull(RsMst_CommodityDetail!LotSize), "", RsMst_CommodityDetail!LotSize)
   TxtCVariation = IIf(IsNull(RsMst_CommodityDetail!Variation), "", RsMst_CommodityDetail!Variation)
Else
   MsgBox "No record found for Exchange" & CmbExchangeTypeID.Text & " - " & CmbCommodityID.Text
   CmbCommodityID.SetFocus
   Exit Sub
End If
mCommodityYearlyFlag = IIf(IsNull(TmpRs7!YearlyBillingFlg), 0, TmpRs7!YearlyBillingFlg)

Call Clear_RWCTabControl
Call Grid_Head_RRN
Call Clear_TTDTabControl
Call Grid_Head_TTD
Call Clear_GCTDTabControl
Call Grid_Head_GCTD

'Call CheckRejection(Check1.Value)
'Call Check1_Click
Call DisplayRWCModification

'If mExchangeTypeID = 1 Then 'NON NCDEX
'   Call WHRDetails(mCMPID, mCommodityID, TxtRRNNo.Text)
'End If
'If mExchangeTypeID = 2 Then 'NCDEX
'   Call RRNDetails(mCMPID, mCommodityID, TxtRRNNo.Text)
'End If
'If mExchangeTypeID = 3 Then 'SPOT
'   Call CDCDetails(mCMPID, mCommodityID, TxtRRNNo.Text)
'End If

'SSTab1.Tab = 0
End Sub



Private Sub CmbExchangeTypeID_GotFocus()
tmp = CmbExchangeTypeID.Text
Call TableMst_ExchangeType
CmbExchangeTypeID.Clear
If RsMst_ExchangeType.RecordCount = 0 Then
   MsgBox "No record found!!!"
   Exit Sub
End If
For I = 1 To RsMst_ExchangeType.RecordCount
    CmbExchangeTypeID.AddItem RsMst_ExchangeType!ExchangeType
    RsMst_ExchangeType.MoveNext
Next
CmbExchangeTypeID.Text = tmp
End Sub
'----
Private Sub CmbExchangeTypeID_LostFocus()
If CmbExchangeTypeID.Text = "" Then
   Exit Sub
End If
RsMst_ExchangeType.MoveFirst
RsMst_ExchangeType.Find "ExchangeType = '" & CmbExchangeTypeID.Text & "'"
If RsMst_ExchangeType.EOF Then
   MsgBox "Invalid selection!!!"
   CmbExchangeTypeID.SetFocus
   Exit Sub
End If
mExchangeTypeID = RsMst_ExchangeType!ExchangeTypeID
CmbCommodityID.Text = ""
mCommodityID = 0
TxtCGunnyWt = ""
TxtCLotSize = ""
TxtCVariation = ""

If mExchangeTypeID = 1 Then
   ChkSCMFlag.Enabled = True
Else
   ChkSCMFlag.Value = 0
   ChkSCMFlag.Enabled = False
End If

Call Clear_RWCTabControl
Call Clear_TTDTabControl
Call Clear_GCTDTabControl
Call Grid_Head_RRN
Call Grid_Head_TTD
Call Grid_Head_GCTD


Check1.Value = 0
Call DisplayRWCModification


'Call Check1_Click

End Sub
Private Sub DisplayRWCModification()

If mExchangeTypeID = 1 Then 'NON NCDEX
   Call SetCaptions("WHR No", "WHR No", "Client Name", "WHR Qty", "Depository ID", "Depository Name", False)
   Call WHRDetails(mCMPID, mCommodityID, TxtRRNNo.Text)
End If
If mExchangeTypeID = 2 Then 'NCDEX
   If Check1.Value = 0 Then
      Call SetCaptions("Remat Request No. (RRN)", "RRN No", "CMSE Lot No", "RRN Qty", "DP ID", "DP Name", True)
      Call RRNDetails(mCMPID, mCommodityID, TxtRRNNo.Text)
   Else
      Call Check1_Click
   End If
End If
If mExchangeTypeID = 3 Then 'SPOT
   Call SetCaptions("CDC No", "CDC-No", "CDC-No", "CDC Qty", "TM ID", "Trading Member Name", True)
   Call CDCDetails(mCMPID, mCommodityID, TxtRRNNo.Text)
End If

End Sub
Private Sub SetCaptions(TabCaption_, RRNNo_, ISINNo_, RRNQty_, DPID_, DPName_, Yn)

SSTab1.TabCaption(0) = TabCaption_
LblRRNNo.Caption = RRNNo_
LblISINNo.Caption = ISINNo_
LblRRNQty.Caption = RRNQty_
LblDPID.Caption = DPID_
LblDPName.Caption = DPName_
Label32.Caption = ISINNo_
Label29.Visible = Yn
Check1.Visible = Yn
If Yn = False Then
   MSHFlexGrid2.ColWidth(2) = 0
   LblISINNo.Visible = Yn
   TxtISINno.Visible = Yn
   If mExchangeTypeID = 3 Then
      Label32.Visible = Yn
      TxtGSLISINno.Visible = Yn
   Else
      Label32.Visible = True
      TxtGSLISINno.Visible = True
   End If
Else
   If Check1.Value = 0 Then
      LblISINNo.Visible = Yn
      TxtISINno.Visible = Yn
      Label32.Visible = Yn
      TxtGSLISINno.Visible = Yn
   Else
      If mExchangeTypeID = 2 Then
         SSTab1.TabCaption(0) = "Client Detail "
      End If
      If mExchangeTypeID = 3 Then
         SSTab1.TabCaption(0) = "CDC Detail "
      End If
      LblISINNo.Visible = False
      TxtISINno.Visible = False
      If mExchangeTypeID = 3 Then
         Label32.Visible = True
         TxtGSLISINno.Visible = True
      End If
   End If
End If
MSHFlexGrid2.TextMatrix(1, 1) = RRNNo_
MSHFlexGrid2.TextMatrix(1, 5) = DPID_
MSHFlexGrid2.TextMatrix(1, 6) = DPName_
MSHFlexGrid2.TextMatrix(1, 7) = RRNQty_
MSHFlexGrid4.TextMatrix(1, 12) = ISINNo_

End Sub
Private Sub CmbRRNISIN_LostFocus()
If CmbRRNISIN.Text = "" Then Exit Sub
RsMst_RRN.MoveFirst
RsMst_RRN.Find "ISINID = '" & Trim(CmbRRNISIN.Text) & "'"

If Not RsMst_RRN.EOF Then
   TxtISINno = IIf(IsNull(RsMst_RRN!ISINID), "", RsMst_RRN!ISINID)
   mRRNCommTrackFlag = IIf(IsNull(RsMst_RRN!CommTrackFlag), 0, RsMst_RRN!CommTrackFlag)
   If mRRNCommTrackFlag = 1 Then
      If MSHFlexGrid2.Rows > 3 And MSHFlexGrid10.Rows < 4 Then
         MsgBox "Please enter non Comtrack RRN only !!"
         TxtRRNNo.SetFocus
         Exit Sub
      End If
   Else
      If MSHFlexGrid10.Rows > 3 Then
         MsgBox "Please enter Comtrack RRN only!!"
         TxtRRNNo.SetFocus
         Exit Sub
      End If
   End If
   
   Call TableMst_ISIN("Where ISINID = '" & TxtISINno & "'")
   Call TableMst_Location("Where LocationID = '" & RsMst_ISIN!LocationID & "'")
   If Not RsMst_ISIN.EOF Then
'      If TxtPreFix <> RsMst_Location!SM_Prefix Then
'         MsgBox "ISIN Location and above CMP location do not match!!!"
'         TxtRRNNo.SetFocus
'         Exit Sub
'      End If
      
      If TxtPreFix <> RsMst_Location!SM_Prefix Then
'         If LCase(TxtPreFix) <> "bknrp" Then
         If LCase(TxtPreFix) <> "bknrp" And LCase(TxtPreFix) <> "ujjap" And LCase(TxtPreFix) <> "dewsp" And LCase(TxtPreFix) <> "vdshp" And LCase(TxtPreFix) <> "adlbp" And LCase(TxtPreFix) <> "aklap" And LCase(TxtPreFix) <> "sngrp" And LCase(TxtPreFix) <> "nzmdp" And LCase(TxtPreFix) <> "unjap" And LCase(TxtPreFix) <> "sglip" And LCase(TxtPreFix) <> "gntrp" And LCase(TxtPreFix) <> "dvgrp" And LCase(TxtPreFix) <> "jdhrp" And LCase(TxtPreFix) <> "hngtp" And LCase(TxtPreFix) <> "kadip" And LCase(TxtPreFix) <> "guntr" And LCase(TxtPreFix) <> "glbrg" And LCase(TxtPreFix) <> "gbrgp" And LCase(TxtPreFix) <> "kngym" And LCase(TxtPreFix) <> "lturp" And LCase(TxtPreFix) <> "ktap" And LCase(TxtPreFix) <> "rsgrp" And LCase(TxtPreFix) <> "rsngr" And LCase(TxtPreFix) <> "ttcrp" Then
            MsgBox "CMSE Lot No Location and above CMP location do not match!!!"
            TxtRRNNo.SetFocus
            Exit Sub
         Else
            If LCase(TxtPreFix) = "bknrp" Then
               If LCase(RsMst_Location!SM_Prefix) <> "bkner" Then
                  MsgBox "CMSE Lot No Location should be Bikaner only!!!"
                  TxtRRNNo.SetFocus
                  Exit Sub
               End If
            ElseIf LCase(TxtPreFix) = "ujjap" Then
               If LCase(RsMst_Location!SM_Prefix) <> "ujjan" Then
                  MsgBox "CMSE Lot No Location should be Ujjain only!!!"
                  TxtRRNNo.SetFocus
                  Exit Sub
               End If
            ElseIf LCase(TxtPreFix) = "dewsp" Then
               If LCase(RsMst_Location!SM_Prefix) <> "dewas" Then
                  MsgBox "CMSE Lot No Location should be Dewas only!!!"
                  TxtRRNNo.SetFocus
                  Exit Sub
               End If
            ElseIf LCase(TxtPreFix) = "vdshp" Then
               If LCase(RsMst_Location!SM_Prefix) <> "vdisa" Then
                  MsgBox "CMSE Lot No Location should be Vidisha only!!!"
                  TxtRRNNo.SetFocus
                  Exit Sub
               End If
            ElseIf LCase(TxtPreFix) = "adlbp" Then 'Add by prajakta 29 Jan 2015
               If LCase(RsMst_Location!SM_Prefix) <> "adlbd" Then
                  MsgBox "CMSE Lot No Location should be Adilabad only!!!"
                  TxtRRNNo.SetFocus
                  Exit Sub
               End If
            ElseIf LCase(TxtPreFix) = "aklap" Then 'Add by Parag 02 May 2015
               If LCase(RsMst_Location!SM_Prefix) <> "akola" Then
                  MsgBox "CMSE Lot No Location should be Akola only!!!"
                  TxtRRNNo.SetFocus
                  Exit Sub
               End If
            ElseIf LCase(TxtPreFix) = "sngrp" Then 'Add by Parag 02 May 2015
               If LCase(RsMst_Location!SM_Prefix) <> "sgngr" Then
                  MsgBox "CMSE Lot No Location should be Sriganganagar only!!!"
                  TxtRRNNo.SetFocus
                  Exit Sub
               End If
            ElseIf LCase(TxtPreFix) = "nzmdp" Then 'Add by Parag 07 May 2016
               If LCase(RsMst_Location!SM_Prefix) <> "nzmbd" Then
                  MsgBox "CMSE Lot No Location should be Nizamabad only!!!"
                  TxtRRNNo.SetFocus
                  Exit Sub
               End If
            ElseIf LCase(TxtPreFix) = "unjap" Then 'Add by Parag 23 May 2016
               If LCase(RsMst_Location!SM_Prefix) <> "unjha" Then
                  MsgBox "CMSE Lot No Location should be Unjha only!!!"
                  TxtRRNNo.SetFocus
                  Exit Sub
               End If
            ElseIf LCase(TxtPreFix) = "sglip" Then 'Add by Parag 30 Sept 2016
               If LCase(RsMst_Location!SM_Prefix) <> "sngli" Then
                  MsgBox "CMSE Lot No Location should be Sangli only!!!"
                  TxtRRNNo.SetFocus
                  Exit Sub
               End If
            ElseIf LCase(TxtPreFix) = "gntrp" Then 'Add by Parag 04 Nov 2016
               If LCase(RsMst_Location!SM_Prefix) <> "guntr" Then
                  MsgBox "CMSE Lot No Location should be Guntur only!!!"
                  TxtRRNNo.SetFocus
                  Exit Sub
               End If
            ElseIf LCase(TxtPreFix) = "guntr" Then 'Add by Parag 07 March 2018
               If LCase(RsMst_Location!SM_Prefix) <> "gntrp" Then
                  MsgBox "CMSE Lot No Location should be Guntur P only!!!"
                  TxtRRNNo.SetFocus
                  Exit Sub
               End If
            
            ElseIf LCase(TxtPreFix) = "dvgrp" Then 'Add by Parag 04 Nov 2016
               If LCase(RsMst_Location!SM_Prefix) <> "dvngr" Then
                  MsgBox "CMSE Lot No Location should be Davangeere only!!!"
                  TxtRRNNo.SetFocus
                  Exit Sub
               End If
            ElseIf LCase(TxtPreFix) = "jdhrp" Then 'Add by Parag 26 June 2017
               If LCase(RsMst_Location!SM_Prefix) <> "jdhpr" Then
                  MsgBox "CMSE Lot No Location should be Jodhpur only!!!"
                  TxtRRNNo.SetFocus
                  Exit Sub
               End If
            ElseIf LCase(TxtPreFix) = "hngtp" Then 'Add by Parag 13 Sept 2017
               If LCase(RsMst_Location!SM_Prefix) <> "hghat" Then
                  MsgBox "CMSE Lot No Location should be Hinghanghat only!!!"
                  TxtRRNNo.SetFocus
                  Exit Sub
               End If
            ElseIf LCase(TxtPreFix) = "kadip" Then 'Add by Parag 03 Jan 2018
               If LCase(RsMst_Location!SM_Prefix) <> "kadi" Then
                  MsgBox "CMSE Lot No Location should be Kadi only!!!"
                  TxtRRNNo.SetFocus
                  Exit Sub
               End If
            ElseIf LCase(TxtPreFix) = "gbrgp" Then 'Add by Parag 09 March 2018
               If LCase(RsMst_Location!SM_Prefix) <> "glbrg" Then
                  MsgBox "CMSE Lot No Location should be Gulburga only!!!"
                  TxtRRNNo.SetFocus
                  Exit Sub
               End If
            ElseIf LCase(TxtPreFix) = "glbrg" Then 'Add by Parag 09 March 2018
               If LCase(RsMst_Location!SM_Prefix) <> "gbrgp" Then
                  MsgBox "CMSE Lot No Location should be Gulburga only!!!"
                  TxtRRNNo.SetFocus
                  Exit Sub
               End If
            ElseIf LCase(TxtPreFix) = "kngmp" Then 'Add by Parag 30 April 2018
               If LCase(RsMst_Location!SM_Prefix) <> "kngym" Then
                  MsgBox "CMSE Lot No Location should be Kangayam P only!!!"
                  TxtRRNNo.SetFocus
                  Exit Sub
               End If
            ElseIf LCase(TxtPreFix) = "lturp" Then 'Add by Parag 19 Feb 2019
               If LCase(RsMst_Location!SM_Prefix) <> "latur" Then
                  MsgBox "CMSE Lot No Location should be Latur only!!!"
                  TxtRRNNo.SetFocus
                  Exit Sub
               End If
            ElseIf LCase(TxtPreFix) = "ktap" Then 'Add by Parag 19 Feb 2019
               If LCase(RsMst_Location!SM_Prefix) <> "kota" Then
                  MsgBox "CMSE Lot No Location should be Kota only!!!"
                  TxtRRNNo.SetFocus
                  Exit Sub
               End If
            ElseIf LCase(TxtPreFix) = "rsgrp" Then 'Add by Parag 19 Mar 2019
               If LCase(RsMst_Location!SM_Prefix) <> "rsngr" Then
                  MsgBox "CMSE Lot No Location should be Raisingnagar only!!!"
                  TxtRRNNo.SetFocus
                  Exit Sub
               End If
            ElseIf LCase(TxtPreFix) = "rsngr" Then 'Add by Parag 02 Aug 2019
               If LCase(RsMst_Location!SM_Prefix) <> "rsgrp" Then
                  MsgBox "CMSE Lot No Location should be Raisingnagar P only!!!"
                  TxtRRNNo.SetFocus
                  Exit Sub
               End If
            ElseIf LCase(TxtPreFix) = "ttcrp" Then 'Add by Parag 02 Aug 2019
               If LCase(RsMst_Location!SM_Prefix) <> "tcrin" Then
                  MsgBox "CMSE Lot No Location should be Tuticorin only!!!"
                  TxtRRNNo.SetFocus
                  Exit Sub
               End If
            End If
         End If
      End If
      
      ' Removed CMP Level Check on 31.03.2010
'      If mCMPID_ <> RsMst_ISIN!CMPID Then
'         MsgBox "ISIN CMP and above CMP do not match!!!"
'         TxtRRNNo.SetFocus
'         Exit Sub
'      End If
'Commented on 09 July 2013
'      If mCommodityID_ <> RsMst_ISIN!CommodityID Then
'         MsgBox "ISIN commodity and selected Commodity do not match!!!"
'         TxtRRNNo.SetFocus
'         Exit Sub
'      End If
      
      'Condition introduced for withdrawal of Chilli and Chilli 25 mail from Mayur Sir on 09 July 2013
      If mCommodityID <> 31 And mCommodityID <> 114 Then
         If mCommodityID <> RsMst_ISIN!CommodityID Then
            MsgBox "CMSE Lot No commodity and selected Commodity do not match!!!"
            TxtRRNNo.SetFocus
            Exit Sub
         End If
      ElseIf mCommodityID = 31 Then
         If RsMst_ISIN!CommodityID <> 31 And RsMst_ISIN!CommodityID <> 114 Then
            MsgBox "CMSE Lot No commodity should be either Chilli or Chilli 25!!!"
            TxtRRNNo.SetFocus
            Exit Sub
         End If
      ElseIf mCommodityID = 114 Then
         If RsMst_ISIN!CommodityID <> 31 And RsMst_ISIN!CommodityID <> 114 Then
            MsgBox "CMSE Lot No commodity should be either Chilli or Chilli 25!!!"
            TxtRRNNo.SetFocus
            Exit Sub
         End If
      End If
  
  End If
  mRWCID = RsMst_RRN!RRNID
  TxtRRNQty = IIf(IsNull(RsMst_RRN!RRNQty), "", RsMst_RRN!RRNQty)
  'TxtBalanceQty = Format(Val(TxtRRNQty) - Val(RsMst_RRN!DeliveredQty), "#######0.000")
  TxtBalanceQty = Format(Val(TxtRRNQty) - (Val(RsMst_RRN!DeliveredQty) + Val(RsMst_RRN!OutOfSystemQty)), "#######0.000")
  mClientIDRow = RsMst_RRN!ClientID
  Call GetClientDetails(mClientIDRow)
  CmbRRNISIN.Visible = False
  Exit Sub
End If

MsgBox "Invalid CMSE Lot No selection !!! "
CmbRRNISIN.SetFocus

End Sub

Private Sub CmbSRejection_LostFocus()
If CmbSRejection.Text = "" Then Exit Sub
'If LCase(CmbSRejection.Text) <> "rejected" And LCase(CmbSRejection.Text) <> "un rejected" Then
'   MsgBox "Invalid selection!!!"
'   CmbSRejection.SetFocus
'   Exit Sub
'End If
'
If LCase(CmbSRejection.Text) = "rejected" Then
   mRejection = 1
ElseIf LCase(CmbSRejection.Text) = "un rejected" Then
   mRejection = 0
Else
   MsgBox "Invalid selection!!!"
   CmbSRejection.SetFocus
   Exit Sub
End If

End Sub

Private Sub CmbSSCMFlag_LostFocus()
If CmbSSCMFlag.Text = "" Then Exit Sub

If LCase(CmbSSCMFlag.Text) <> "yes" And LCase(CmbSSCMFlag.Text) <> "no" Then
   MsgBox "Invalid Selection!!!"
   CmbSSCMFlag.SetFocus
   Exit Sub
End If

End Sub

Private Sub CmdAgent_Click()
FrmMst_Agent.Show
End Sub

Private Sub CmdApprovePrint_Click()

Retval = True
mTotOutstAmt = 0

If LCase(CmbExchangeTypeID.Text) = "non ncdex" Then
   If LCase(Gen_UserRole) = "power" Or LCase(Gen_UserRole) = "admin" Or LCase(Gen_UserRole) = "superadmin" Or LCase(Gen_UserRole) = "headcmg" Then
      mtimes = 1
   ElseIf LCase(Gen_UserRole) = "amsandrms" Then
      mtimes = mMultOfAM
   Else
      mtimes = mMultOfAll
   End If
   
   If MsgBox("Is invoice generated for this BOE amount?", vbYesNo) = vbYes Then
      mTIID = ""
      ''mVoucherNo = InputBox("Enter TI No. (Voucher No)", "Enter TI No.")
      FrmInvoice.Visible = True
      TxtInvNo.SetFocus
   ''   If TxtInvNo = "" Then Exit Sub
   ''   mVoucherNo = Trim(TxtInvNo.Text)
   ''   If CheckVoucheNo(mVoucherNo) = False Then
   ''      'MsgBox "Voucher No not found"
   ''      Exit Sub
   ''   End If
   ''   Call CmdInvCancel_Click
   Else
      mTotOutstAmt = Val(TxtBOEAmount)
      Call ApproveCITFPrint
   End If
Else
   Call ApproveCITFPrint
End If

End Sub

Private Sub CmdASearch_Click()
Dim wc As String
Call Grid_Head_ApprovaPrint
wc = ""
If TxtAToDate.Value < TxtAFromDate.Value Then
   MsgBox "'To Date' must be greater than 'From Date'!!!"
   TxtAToDate.SetFocus
   Exit Sub
End If

If CmbACMPID.Text <> "" Then
   wc = GenWc(wc, CmbACMPID.Text, "Mst_CMP.CMPID", "N")
Else
   If Gen_WcCMP <> "" Then
      tmp = Replace(Gen_WcCMP, "CMPID", "mst_CMP.CMPID")
      wc = GenWc(wc, tmp, "", "I")
   End If
End If


If CmbACommodityID.Text <> "" Then
   wc = GenWc(wc, mACommodityID, "Mst_Commodity.CommodityID", "N")
End If
wc = GenWc(wc, 1, "Mst_ExchangeType.ExchangeTypeID", "N")
If CmbALocationID.Text <> "" Then
   wc = GenWc(wc, mASM_prefix, "Txn_CxTf_Details.sm_prefix", "S")
End If

If TxtACITFNo <> "" Then
   wc = GenWc(wc, TxtACITFNo, "Txn_CxTf_Details.CxTF_DetailsID", "N")
End If

tmp = "Select Txn_CxTf_Details.CxTF_DetailsID,Txn_CxTf_Details.CxTFDate,Mst_ExchangeType.ExchangeType,Mst_CMP.CMPName,Txn_CxTf_Details.SM_Prefix,"
tmp = tmp & " Mst_Commodity.Commodity,Txn_CxTf_Details.Flag,"
tmp = tmp & " Txn_CxTf_Details.Agent,Txn_CxTf_Details.Remark,Txn_CxTf_Details.HologramNo,"
tmp = tmp & " Mst_Agent.AgentName , Txn_CxTf_Details.ManualCDTFNo, "
tmp = tmp & " Sum (Txn_CxTf_Gsl.NoofBags) 'NoOfBags',sum(Txn_CxTf_Gsl.GrossWeight) 'GrossWeight',Txn_CxTf_Details.G_UID,Txn_CxTf_Details.BookedSpaceFlag "
tmp = tmp & " From Txn_CxTf_Details"
tmp = tmp & " Inner Join Txn_CxTf_Gsl  on Txn_CxTf_Details.SM_Prefix = Txn_CxTf_Gsl.SM_Prefix And"
tmp = tmp & " Txn_CxTf_Details.CxTF_DetailsID = Txn_CxTf_Gsl.CxTF_DetailsID And Txn_CxTf_Details.TxnType = Txn_CxTf_Gsl.TxnType"
tmp = tmp & " Inner Join Mst_CMP On Txn_CxTf_Details.CMPID = Mst_CMP.CMPID"
tmp = tmp & " Inner Join Mst_Commodity On Txn_CxTf_Details.CommodityID = Mst_Commodity.CommodityID"
tmp = tmp & " Inner Join Mst_ExchangeType ON Txn_CxTf_Details.ExchangeTypeID = Mst_ExchangeType.ExchangeTypeID"
tmp = tmp & " Left Join Mst_Agent ON Txn_CxTf_Details.AgentID = Mst_Agent.AgentID"
If TxtSTruckNo <> "" Then
   tmp = tmp & " Inner Join Txn_CxTF_TruckDetail  on Txn_CxTf_Details.SM_Prefix = Txn_CxTF_TruckDetail.SM_Prefix And"
   tmp = tmp & " Txn_CxTf_Details.CxTF_DetailsID = Txn_CxTF_TruckDetail.CxTF_DetailsID And Txn_CxTf_Details.TxnType = Txn_CxTF_TruckDetail.TxnType"
End If
tmp = tmp & " Where Txn_CxTf_Details.TxnType = 'W' "

If Gen_DBType = "MDB" Then
   tmp = tmp & " and Txn_CxTf_Details.CxTFDate Between #" & Format(TxtAFromDate.Value, Gen_DatFormat) & "# And #" & Format(TxtAToDate.Value, Gen_DatFormat) & "#) "
Else
    tmp = tmp & " and Txn_CxTf_Details.CxTFDate Between '" & Format(TxtAFromDate.Value, Gen_DatFormat) & "' And '" & Format(TxtAToDate.Value, Gen_DatFormat) & "' "
End If

If wc <> "" Then
   tmp = tmp & wc
End If

tmp = tmp & " group by Txn_CxTf_Details.CxTF_DetailsID,Txn_CxTf_Details.CxTFDate,Mst_ExchangeType.ExchangeType,Mst_CMP.CMPName,"
tmp = tmp & " Txn_CxTf_Details.SM_Prefix, Mst_Commodity.Commodity,Txn_CxTf_Details.Flag, Txn_CxTf_Details.Agent,"
tmp = tmp & " Txn_CxTf_Details.Remark,Txn_CxTf_Details.HologramNo, Mst_Agent.AgentName , Txn_CxTf_Details.ManualCDTFNo,"
tmp = tmp & " Txn_CxTf_Details.G_UID,Txn_CxTf_Details.BookedSpaceFlag"
tmp = tmp & " ORDER BY Txn_CxTf_Details.sm_prefix,Mst_CMP.CMPName,Txn_CxTf_Details.CxTF_DetailsID "

Call TmpTable

MSHFlexGrid1.Rows = TmpRs.RecordCount + 2
For I = 1 To TmpRs.RecordCount
    For C = 0 To TmpRs.Fields.Count - 1
        MSHFlexGrid1.TextMatrix(I, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
    Next
    If MSHFlexGrid1.TextMatrix(I, 15) = 1 Then
       MSHFlexGrid1.TextMatrix(I, 15) = "Yes"
    Else
       If LCase(MSHFlexGrid1.TextMatrix(I, 2)) = "ncdex" Then
          MSHFlexGrid1.TextMatrix(I, 15) = "No"
       Else
          MSHFlexGrid1.TextMatrix(I, 15) = "Not Applicable"
       End If
    End If
    TmpRs.MoveNext
Next

End Sub

Private Sub CmdAPrint_Click()


If CheckForBOEAmount() = True And Val(TxtBOEAmount.Text) = 0 Then
   MsgBox "BOE amount should not zero for this CITF!!"
   Exit Sub
End If

TxtARemark.Text = ""
Frame2.Visible = True
End Sub

Private Sub CmdBOE_Click()

'If LCase(CmbExchangeTypeID) = "ncdex" Then
'   MsgBox "BOE is for Non NCDEX Clients only"
'   Exit Sub
'End If

If LCase(CmbExchangeTypeID.Text) = "non ncdex" Then
   Call UpdateBillingDetails
End If

Call FillBOE
Call SaveCxTFBOE
TxtBOEAmount = GetBOEAmount()
Call CreateExcel

End Sub

Private Sub CmdCancel_Click()
TxtHologramNo = ""
Frame3.Visible = False
End Sub

Private Sub CmdGSLChkOK_Click()
Frame1.Visible = False
Picture3.Enabled = True
End Sub

Private Sub CmdInvCancel_Click()
OptTINo.Value = True
TxtInvNo.Text = ""
FrmInvoice.Visible = False
End Sub

Private Sub CmdInvOK_Click()
If TxtInvNo.Text = "" Then
   MsgBox "Blank Invoice no not allowed!!"
   TxtInvNo.SetFocus
   Exit Sub
End If
'If TxtInvNo = "" Then Exit Sub
mVoucherNo = Trim(TxtInvNo.Text)
If CheckVoucheNo(mVoucherNo) = False Then
   'MsgBox "Voucher No not found"
   Exit Sub
End If
Call ApproveCITFPrint
Call CmdInvCancel_Click

End Sub

Private Sub CmdPrintPreview_Click()

If BalBagNeg = "Y" Then
   MsgBox "Duplicate RRN/WHR Or One or more GSL having Negative  'No of Bags' please contact CMG/IT for assistance!!!"
   Exit Sub
End If

If Val(MSHFlexGrid3.TextMatrix(0, 5)) <> Val(MSHFlexGrid4.TextMatrix(0, 7)) Then
   MsgBox "Gross Weight in 'Truck Details' tab is different from Gross Weight in 'GSL Details' tab!!! "
   Exit Sub
End If

If Val(MSHFlexGrid2.TextMatrix(0, 9)) <> Val(MSHFlexGrid4.TextMatrix(0, 10)) Then
   If mExchangeTypeID = 1 Then
      MsgBox "Delivered Qty in 'WHR No' Tab and Electronic Weight in 'GSL Details' tab are not equal !!!!"
   ElseIf mExchangeTypeID = 2 Then
      MsgBox "Delivered Qty in 'RRN Detail' Tab and Electronic Weight in 'GSL Details' tab are not equal !!!!"
   Else
      MsgBox "Delivered Qty in 'CDC Detail' Tab and Electronic Weight in 'GSL Detail' tab are not equal !!!!"
   End If
   Exit Sub
End If

If Val(MSHFlexGrid2.TextMatrix(0, 15)) <> Val(MSHFlexGrid4.TextMatrix(0, 6)) Then
   If mExchangeTypeID = 1 Then
      MsgBox "Delivered Bags in 'WHR No' Tab and 'No. of Bags' in 'GSL Details' tab are not equal !!!!"
'      MsgBox "Delivered Bags in 'RRN Detail' Tab and 'No of Bags' in 'GSL Details' tab are not equal !!!!"
      Exit Sub
   ElseIf mExchangeTypeID = 3 Then
      MsgBox "Delivered Bags in 'CDC Detail' Tab and 'No of Bags' in 'GSL Details' tab are not equal !!!!"
      Exit Sub
   End If
End If


Gen_mTimeStamp = GetTimeStamp
''Call Create_Xls(Gen_mTimeStamp & Gen_UserName & "_CITF-" & Trim(TxtPreFix) & "" & Trim(TxtCxTFNo) & ".xls")
''Call Xls_Print_Rpt

'''If LCase(CmbExchangeTypeID.Text) = "non ncdex" Then
'''   If CheckClientOutStanding = False Then
'''      Exit Sub
'''   End If
'''End If
PntPrv = "Y"
Call PDF_Print_Rpt

End Sub

Private Sub CmdSearch_Click()
Dim wc As String
Call Grid_Head
wc = ""
If STo.Value < SFrom.Value Then
   MsgBox "'To Date' must be greater than 'From Date'!!!"
   STo.SetFocus
   Exit Sub
End If

If CmbSCmp.Text <> "" Then
   wc = GenWc(wc, mCMPName, "Mst_CMP.CMPID", "N")
Else
   If Gen_WcCMP <> "" Then
      tmp = Replace(Gen_WcCMP, "CMPID", "mst_CMP.CMPID")
      wc = GenWc(wc, tmp, "", "I")
   End If
End If


If CmbSCommodity.Text <> "" Then
   wc = GenWc(wc, mCommodityID, "Mst_Commodity.CommodityID", "N")
End If
If CmbSExchangeType.Text <> "" Then
   wc = GenWc(wc, mExchangeTypeID, "Mst_ExchangeType.ExchangeTypeID", "N")
End If
If CmbPreFix.Text <> "" Then
   wc = GenWc(wc, msSM_prefix, "Txn_CxTf_Details.sm_prefix", "S")
End If

If TxtSCxTFNo <> "" Then
   wc = GenWc(wc, TxtSCxTFNo, "Txn_CxTf_Details.CxTF_DetailsID", "N")
End If

If TxtSManualCxTFNo <> "" Then
   wc = GenWc(wc, TxtSManualCxTFNo, "Txn_CxTf_Details.ManualCDTFNo", "S")
End If

If CmbSRejection.Text <> "" Then
   wc = GenWc(wc, mRejection, "Txn_CxTf_Details.BookedSpaceFlag", "S")
End If

If TxtSTruckNo <> "" Then
   wc = wc & " And Txn_CxTF_TruckDetail.TruckNo like '%" & TxtSTruckNo & "%'"
End If


If CmbSSCMFlag.Text <> "" Then
   If LCase(CmbSSCMFlag.Text) = "yes" Then
      wc = wc & " And isnull(Txn_CxTf_Details.SCMFlag,0) = '1' "
   ElseIf LCase(CmbSSCMFlag.Text) = "no" Then
      wc = wc & " And isnull(Txn_CxTf_Details.SCMFlag,0) = '0' "
   End If
End If

'tmp = "Select Txn_CxTf_Details.CxTFNo,Txn_CxTf_Details.CxTFDate,Mst_ExchangeType.ExchangeType,Mst_CMP.CMPName,Txn_CxTf_Details.SM_Prefix,"
'tmp = tmp & " Mst_Commodity.Commodity,Txn_CxTf_Details.Flag,"
'tmp = tmp & " Txn_CxTf_Details.Agent,Txn_CxTf_Details.Remark,Txn_CxTf_Details.HologramNo,"
'tmp = tmp & " Mst_Agent.AgentName , Txn_CxTf_Details.ManualCDTFNo, Txn_CxTf_Details.G_UID"
'tmp = tmp & " From Txn_CxTf_Details"
'tmp = tmp & " Inner Join Mst_CMP On Txn_CxTf_Details.CMPID = Mst_CMP.CMPID"
'tmp = tmp & " Inner Join Mst_Commodity On Txn_CxTf_Details.CommodityID = Mst_Commodity.CommodityID"
'tmp = tmp & " Inner Join Mst_ExchangeType ON Txn_CxTf_Details.ExchangeTypeID = Mst_ExchangeType.ExchangeTypeID"
'tmp = tmp & " Left Join Mst_Agent ON Txn_CxTf_Details.AgentID = Mst_Agent.AgentID"
'tmp = tmp & " Where Txn_CxTf_Details.TxnType = 'W'"
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " and Txn_CxTf_Details.CxTFDate Between #" & Format(SFrom, Gen_DatFormat) & "# And #" & Format(STo, Gen_DatFormat) & "#) "
'Else
'   tmp = tmp & " and Txn_CxTf_Details.CxTFDate Between '" & Format(SFrom, Gen_DatFormat) & "' And '" & Format(STo, Gen_DatFormat) & "' "
'End If
'
'If wc <> "" Then
'   tmp = tmp & wc
'End If
'
'tmp = tmp & " ORDER BY Txn_CxTf_Details.sm_prefix,Mst_CMP.CMPName,Txn_CxTf_Details.CxTFNo "




tmp = "Select Txn_CxTf_Details.CxTF_DetailsID,Txn_CxTf_Details.CxTFDate,Mst_ExchangeType.ExchangeType,Mst_CMP.CMPName,Txn_CxTf_Details.SM_Prefix,"
tmp = tmp & " Mst_Commodity.Commodity,Txn_CxTf_Details.Flag,"
tmp = tmp & " Txn_CxTf_Details.Agent,Txn_CxTf_Details.Remark,Txn_CxTf_Details.HologramNo,"
'tmp = tmp & " Mst_Agent.AgentName,MC.ClientName+'~'+MC.ClientID 'ClientName' , Txn_CxTf_Details.ManualCDTFNo, "
tmp = tmp & " Mst_Agent.AgentName,'' 'ClientName' , Txn_CxTf_Details.ManualCDTFNo, "
tmp = tmp & " Sum (Txn_CxTf_Gsl.NoofBags) 'NoOfBags',sum(Txn_CxTf_Gsl.GrossWeight) 'GrossWeight',Txn_CxTf_Details.G_UID,Txn_CxTf_Details.BookedSpaceFlag,Txn_CxTf_Details.ApprovePrintFlag,isnull(ME.EmployeeName + '~' + ME.EmployeeNo,ApprovePrintBy)  as ApprovePrintBy,Txn_CxTf_Details.ApprovePrintRemark,Txn_CxTf_Details.ApprovePrintDate,TI.TINo,TI.PostInvoiceNo "
tmp = tmp & " ,Txn_CxTf_Details.OutstandingAmt,Txn_CxTf_Details.BalQty,Txn_CxTf_Details.CommodityValue,Txn_CxTf_Details.SCMFlag "
tmp = tmp & " From Txn_CxTf_Details"
tmp = tmp & " Left Join Txn_CxTf_Gsl  on Txn_CxTf_Details.SM_Prefix = Txn_CxTf_Gsl.SM_Prefix And"
tmp = tmp & " Txn_CxTf_Details.CxTF_DetailsID = Txn_CxTf_Gsl.CxTF_DetailsID And Txn_CxTf_Details.TxnType = Txn_CxTf_Gsl.TxnType"
tmp = tmp & " Inner Join Mst_CMP On Txn_CxTf_Details.CMPID = Mst_CMP.CMPID"
tmp = tmp & " Inner Join Mst_Commodity On Txn_CxTf_Details.CommodityID = Mst_Commodity.CommodityID"
tmp = tmp & " Inner Join Mst_ExchangeType ON Txn_CxTf_Details.ExchangeTypeID = Mst_ExchangeType.ExchangeTypeID"
tmp = tmp & " Left Join Mst_Agent ON Txn_CxTf_Details.AgentID = Mst_Agent.AgentID"
tmp = tmp & " Left Join Mst_Employee ME ON Txn_CxTf_Details.ApprovePrintBy = ME.EmployeeNo"
tmp = tmp & " Left Join Txn_TIDetails TI On TI.TIID=Txn_CxTf_Details.TIID"
'tmp = tmp & " Left Join Mst_GSL MGSL ON Txn_CxTf_Gsl.GSLID = MGSL.GSLID And Txn_CxTf_Gsl.SM_Prefix = MGSL.SM_Prefix "
'tmp = tmp & " Left Join Mst_Client MC ON MGSL.ClientIDRow = MC.ClientIDRow "
If TxtSTruckNo <> "" Then
   tmp = tmp & " Left Join Txn_CxTF_TruckDetail  on Txn_CxTf_Details.SM_Prefix = Txn_CxTF_TruckDetail.SM_Prefix And"
   tmp = tmp & " Txn_CxTf_Details.CxTF_DetailsID = Txn_CxTF_TruckDetail.CxTF_DetailsID And Txn_CxTf_Details.TxnType = Txn_CxTF_TruckDetail.TxnType"
End If
tmp = tmp & " Where Txn_CxTf_Details.TxnType = 'W' "

If Gen_DBType = "MDB" Then
   tmp = tmp & " and Txn_CxTf_Details.CxTFDate Between #" & Format(SFrom, Gen_DatFormat) & "# And #" & Format(STo, Gen_DatFormat) & "#) "
Else
    tmp = tmp & " and Txn_CxTf_Details.CxTFDate Between '" & Format(SFrom, Gen_DatFormat) & "' And '" & Format(STo, Gen_DatFormat) & "' "
End If

If wc <> "" Then
   tmp = tmp & wc
End If
tmp = tmp & " group by Txn_CxTf_Details.CxTF_DetailsID,Txn_CxTf_Details.CxTFDate,Mst_ExchangeType.ExchangeType,Mst_CMP.CMPName,"
tmp = tmp & " Txn_CxTf_Details.SM_Prefix, Mst_Commodity.Commodity,Txn_CxTf_Details.Flag, Txn_CxTf_Details.Agent,"
tmp = tmp & " Txn_CxTf_Details.Remark,Txn_CxTf_Details.HologramNo, Mst_Agent.AgentName , Txn_CxTf_Details.ManualCDTFNo,"
tmp = tmp & " Txn_CxTf_Details.G_UID,Txn_CxTf_Details.BookedSpaceFlag,Txn_CxTf_Details.ApprovePrintFlag,ME.EmployeeName,ME.EmployeeNo,Txn_CxTf_Details.ApprovePrintRemark,Txn_CxTf_Details.ApprovePrintDate,TI.TINo,TI.PostInvoiceNo,Txn_CxTf_Details.OutstandingAmt,Txn_CxTf_Details.CommodityValue,Txn_CxTf_Details.BalQty,Txn_CxTf_Details.ApprovePrintBy,Txn_CxTf_Details.SCMFlag" ',MC.ClientName,MC.ClientID"
tmp = tmp & " ORDER BY Txn_CxTf_Details.sm_prefix,Mst_CMP.CMPName,Txn_CxTf_Details.CxTF_DetailsID "

Call TmpTable

MSHFlexGrid1.Rows = TmpRs.RecordCount + 2
For I = 1 To TmpRs.RecordCount
    For C = 0 To TmpRs.Fields.Count - 1
        MSHFlexGrid1.TextMatrix(I, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
    Next
    If MSHFlexGrid1.TextMatrix(I, 16) = 1 Then
       MSHFlexGrid1.TextMatrix(I, 16) = "Yes"
    Else
       If LCase(MSHFlexGrid1.TextMatrix(I, 2)) = "ncdex" Then
          MSHFlexGrid1.TextMatrix(I, 16) = "No"
       Else
          MSHFlexGrid1.TextMatrix(I, 16) = "Not Applicable"
       End If
    End If
    If MSHFlexGrid1.TextMatrix(I, 26) = 1 Then
       MSHFlexGrid1.TextMatrix(I, 26) = "Yes"
    Else
       MSHFlexGrid1.TextMatrix(I, 26) = "No"
    End If
    TmpRs.MoveNext
Next


wc = ""
If Gen_WcCMP <> "" Then
   wc = "And " & Gen_WcCMP
End If

RsTxn_CxTF_Details.Close
Call TableTxn_CxTF_Details("Where TxnType = 'W' " & wc & " Order By SM_Prefix,CMPID,CxTF_DetailsID  ")

End Sub
Function GenWc(wc_ As String, ID As Variant, Field_ As String, Type_ As String) As String
Dim Retval
Retval = wc_
If Retval = "" Then
   If Type_ = "S" Then
      Retval = "And " & Field_ & "= '" & ID & "'"
   ElseIf Type_ = "I" Then
      Retval = "And " & Field_ & " " & ID
   Else
      Retval = "And " & Field_ & "= " & ID
   End If
Else
   If Type_ = "S" Then
      Retval = Retval & " And " & Field_ & "= '" & ID & "'"
   ElseIf Type_ = "I" Then
      Retval = Retval & " And " & Field_ & " " & ID
   Else
      Retval = Retval & " And " & Field_ & "= " & ID
   End If
End If
GenWc = Retval
End Function

Private Sub Command1_Click()
TxtARemark.Text = ""
Frame2.Visible = False
End Sub

Private Sub DeleteCmd_Click()

If LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "headcmg" And LCase(Gen_UserRole) <> "superadmin" And LCase(Gen_UserRole) <> "admin" Then
   If TxtCxTFDate.Value < Date - 1 Then
      MsgBox "This CITF is more than 1 days old. You Can't Delete This CITF!!! "
      Exit Sub
   End If
End If

If mExchangeTypeID = 2 Then
   If Check1.Value = 1 Then
      Call TableTxn_CDTFBOE(" Where SM_Prefix='" & TxtPreFix & "' And CxTF_DetailsID=" & Val(TxtCxTFNo) & " And TxnType='W'")
      If RsTxn_CDTFBOE.RecordCount > 0 Then
         MsgBox "Rejected invoice is generated for this CITF, you can't delete!!"
         Exit Sub
      End If
   Else
      Call TableTxn_CDTFBOE(" Where SM_Prefix='" & TxtPreFix & "' And CxTF_DetailsID=" & Val(TxtCxTFNo) & " And TxnType='R'")
      If RsTxn_CDTFBOE.RecordCount > 0 Then
         MsgBox "RRN invoice is generated for this CITF, you can't delete!!"
         Exit Sub
      End If
   End If
End If

If RsTxn_CxTF_Details!Flag = "P" Then
   MsgBox "Print out is already generated. You can't edit this CITF!!!"
   Exit Sub
End If

Dim ans As Variant
'On Error GoTo msgError
ans = MsgBox("Do you really want to DELETE this record?", vbYesNo)
If ans = vbYes Then
   
   Call TableLog_CxTF_Details
   
   tmp = "Select * from Log_CxTF_Details Where TxnType = 'W' And SM_prefix = '" & RsTxn_CxTF_Details!SM_Prefix & "' And CxTF_DetailsID = " & RsTxn_CxTF_Details!CxTF_DetailsID & ""
   Call TmpTable
   
   If TmpRs.RecordCount = 0 Then
    
      RsLog_CxTF_Details.AddNew
      RsLog_CxTF_Details!CxTFNO = RsTxn_CxTF_Details!CxTFNO
      RsLog_CxTF_Details!CxTF_DetailsID = RsTxn_CxTF_Details!CxTF_DetailsID
      RsLog_CxTF_Details!TxnType = RsTxn_CxTF_Details!TxnType
      RsLog_CxTF_Details!SM_Prefix = RsTxn_CxTF_Details!SM_Prefix
      RsLog_CxTF_Details!CxTFDate = RsTxn_CxTF_Details!CxTFDate
      RsLog_CxTF_Details!CMPID = RsTxn_CxTF_Details!CMPID
      RsLog_CxTF_Details!CommodityID = RsTxn_CxTF_Details!CommodityID
      RsLog_CxTF_Details!ExchangeTypeID = RsTxn_CxTF_Details!ExchangeTypeID
      RsLog_CxTF_Details!BookedSpaceFlag = RsTxn_CxTF_Details!BookedSpaceFlag
      RsLog_CxTF_Details!Flag = RsTxn_CxTF_Details!Flag
      RsLog_CxTF_Details!Agent = RsTxn_CxTF_Details!Agent
      RsLog_CxTF_Details!Remark = RsTxn_CxTF_Details!Remark
      RsLog_CxTF_Details!ContractID = RsTxn_CxTF_Details!ContractID
      RsLog_CxTF_Details!G_UID = RsTxn_CxTF_Details!G_UID
      RsLog_CxTF_Details!HologramNo = RsTxn_CxTF_Details!HologramNo
      RsLog_CxTF_Details!AgentID = RsTxn_CxTF_Details!AgentID
      RsLog_CxTF_Details!ManualCDTFNo = RsTxn_CxTF_Details!ManualCDTFNo
      RsLog_CxTF_Details!CreatedBy = RsTxn_CxTF_Details!CreatedBy
      RsLog_CxTF_Details!CreatedDate = RsTxn_CxTF_Details!CreatedDate
      RsLog_CxTF_Details!UpdatedDate = RsTxn_CxTF_Details!UpdatedDate
      RsLog_CxTF_Details!UpdatedBy = RsTxn_CxTF_Details!UpdatedBy
      RsLog_CxTF_Details!DeletedBy = Gen_UserLoginID
      RsLog_CxTF_Details.Update
      RsLog_CxTF_Details.Close
   End If
   
   Call DeleteChildTableDetail_RRN
   
   
   tmp = "Delete from Txn_GSL where CxTF_DetailsID = " & TxtCxTFNo
   Call TmpTable
   
   tmp = "select * from Txn_CxTF_TruckDetail where sm_prefix = '" & TxtPreFix & "' And TxnType = 'W'  And  CxTF_DetailsID = " & TxtCxTFNo
   Call TmpTable
   
   
   TmpRs.Close
   RsTxn_CxTF_Details.Delete
   RsTxn_CxTF_Details.Update
   If RsTxn_CxTF_Details.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsTxn_CxTF_Details.MoveNext
   If RsTxn_CxTF_Details.EOF() Then
      RsTxn_CxTF_Details.MoveLast
   End If
   Call Indata
End If
Exit Sub
msgError:
MsgBox "Child record present!!!"
RsTxn_CxTF_Details.CancelUpdate
End Sub

Sub DeleteChildTableDetail_RRN()
Call TableMst_GSL("where  sm_prefix = '" & TxtPreFix & "'")
tmp = "select * from Txn_CxTF_GSL where  sm_prefix = '" & TxtPreFix & "' And TxnType = 'W' And CxTF_DetailsID = " & TxtCxTFNo
Call TmpTable
For I = 1 To TmpRs.RecordCount
    RsMst_GSL.MoveFirst
    RsMst_GSL.Find "Mst_GSL_ID = " & TmpRs!Mst_GSL_ID
    If Not RsMst_GSL.EOF Then
       RsMst_GSL!CDTFBags = IIf(IsNull(RsMst_GSL!CDTFBags), 0, RsMst_GSL!CDTFBags) - Val(MSHFlexGrid4.TextMatrix(I + 1, 6))
       RsMst_GSL!BalanceBags = IIf(IsNull(RsMst_GSL!BalanceBags), 0, RsMst_GSL!BalanceBags) + Val(MSHFlexGrid4.TextMatrix(I + 1, 6))
'       If RsMst_GSL!MadeupFlag = "0" Then
          RsMst_GSL!CDTFWeight = IIf(IsNull(RsMst_GSL!CDTFWeight), 0, RsMst_GSL!CDTFWeight) - Val(MSHFlexGrid4.TextMatrix(I + 1, 10))
          RsMst_GSL!BalanceWeight = IIf(IsNull(RsMst_GSL!BalanceWeight), 0, RsMst_GSL!BalanceWeight) + Val(MSHFlexGrid4.TextMatrix(I + 1, 10))
          RsMst_GSL!WithdrawalAdjWt = IIf(IsNull(RsMst_GSL!WithdrawalAdjWt), 0, RsMst_GSL!WithdrawalAdjWt) - Val(MSHFlexGrid4.TextMatrix(I + 1, 8))
'       End If
       RsMst_GSL.Update
       If Not IsNull(RsMst_GSL!Flag1) Then
          RsMst_GSL!Status = RsMst_GSL!Flag1
       End If
       RsMst_GSL.Update
    End If
    TmpRs.MoveNext
Next

Call ChkDeleteGSL

If mExchangeTypeID = 2 Then
   Call TableMst_RRN
End If

If LCase(TxtPreFix) = "bknrp" Then
   mNewSM_Prefix = "BKNER"
ElseIf LCase(TxtPreFix) = "ujjap" Then
   mNewSM_Prefix = "Ujjan"
ElseIf LCase(TxtPreFix) = "dewsp" Then
   mNewSM_Prefix = "DEWAS"
ElseIf LCase(TxtPreFix) = "vdshp" Then
   mNewSM_Prefix = "VDISA"
ElseIf LCase(TxtPreFix) = "adlbp" Then ' 'adlbp
   mNewSM_Prefix = "adlbd"
ElseIf LCase(TxtPreFix) = "aklap" Then ' 'adlbp
   mNewSM_Prefix = "akola"
ElseIf LCase(TxtPreFix) = "sngrp" Then ' 'adlbp
   mNewSM_Prefix = "sgngr"
ElseIf LCase(TxtPreFix) = "adlbd" Then ' 'adlbp
   mNewSM_Prefix = "adlbp"
ElseIf LCase(TxtPreFix) = "nzmdp" Then ' 'adlbp
   mNewSM_Prefix = "nzmbd"
ElseIf LCase(TxtPreFix) = "unjap" Then ' 'adlbp
   mNewSM_Prefix = "unjha"
ElseIf LCase(TxtPreFix) = "sglip" Then ' 'adlbp
   mNewSM_Prefix = "sngli"
ElseIf LCase(TxtPreFix) = "gntrp" Then ' 'adlbp
   mNewSM_Prefix = "guntr"
ElseIf LCase(TxtPreFix) = "guntr" Then ' 'adlbp
   mNewSM_Prefix = "gntrp"
ElseIf LCase(TxtPreFix) = "dvgrp" Then ' 'adlbp
   mNewSM_Prefix = "dvngr"
ElseIf LCase(TxtPreFix) = "jdhrp" Then ' 'adlbp
   mNewSM_Prefix = "jdhpr"
ElseIf LCase(TxtPreFix) = "hngtp" Then ' 'adlbp
   mNewSM_Prefix = "hghat"
ElseIf LCase(TxtPreFix) = "kadip" Then ' 'kadi
   mNewSM_Prefix = "kadi"
ElseIf LCase(TxtPreFix) = "gbrgp" Then ' 'kadi
   mNewSM_Prefix = "glbrg"
ElseIf LCase(TxtPreFix) = "glbrg" Then ' 'kadi
   mNewSM_Prefix = "gbrgp"
ElseIf LCase(TxtPreFix) = "kngmp" Then ' 'kadi
   mNewSM_Prefix = "kngym"
ElseIf LCase(TxtPreFix) = "lturp" Then ' 'kadi
   mNewSM_Prefix = "latur"
ElseIf LCase(TxtPreFix) = "ktap" Then ' 'KOTA
   mNewSM_Prefix = "kota"
ElseIf LCase(TxtPreFix) = "rsgrp" Then ' 'KOTA
   mNewSM_Prefix = "rsngr"
ElseIf LCase(TxtPreFix) = "rsngr" Then ' 'KOTA
   mNewSM_Prefix = "rsgrp"
ElseIf LCase(TxtPreFix) = "ttcrp" Then ' 'KOTA
   mNewSM_Prefix = "tcrin"
Else
   mNewSM_Prefix = ""
End If

If mExchangeTypeID = 1 Then
   Call TableTxn_WHRDetail(" where  (sm_prefix = '" & TxtPreFix & "' or sm_prefix = '" & mNewSM_Prefix & "')")
End If

'If mExchangeTypeID = 1 Then
'   Call TableTxn_WHRDetail(" where  sm_prefix = '" & TxtPreFix & "'")
'End If
    
If mExchangeTypeID = 3 Then
   Call TableTxn_CDC_GSLDetail(" Where  sm_prefix = '" & TxtPreFix & "'")
End If
For I = 2 To MSHFlexGrid2.Rows - 1
    
    If mExchangeTypeID = 2 Then
       Call TableMst_RRN
       If Check1.Value = 0 Then
          RsMst_RRN.MoveFirst
          If MSHFlexGrid2.TextMatrix(I, 12) <> "" Then
             RsMst_RRN.Find "RRNID = " & Val(MSHFlexGrid2.TextMatrix(I, 12))
             If Not RsMst_RRN.EOF Then
                RsMst_RRN!DeliveredQty = RsMst_RRN!DeliveredQty - Val(MSHFlexGrid2.TextMatrix(I, 9))
                RsMst_RRN.Update
                RsMst_RRN.Requery
             End If
          End If
       End If
    End If
    
    
    If mExchangeTypeID = 1 Then
       If Check1.Value = 0 Then
          RsTxn_WHRDetail.MoveFirst
          RsTxn_WHRDetail.Find "WHRNo = " & Val(MSHFlexGrid2.TextMatrix(I, 12))
          If Not RsTxn_WHRDetail.EOF Then
             RsTxn_WHRDetail!DeliveredQuantity = RsTxn_WHRDetail!DeliveredQuantity - (Val(MSHFlexGrid2.TextMatrix(I, 9)))
             RsTxn_WHRDetail!DeliveredNoOfBags = RsTxn_WHRDetail!DeliveredNoOfBags - (Val(MSHFlexGrid2.TextMatrix(I, 15)) - Val(MSHFlexGrid2.TextMatrix(I, 16)))
             RsTxn_WHRDetail.Update
             RsTxn_WHRDetail.Requery
          End If
       End If
    End If
    
    If mExchangeTypeID = 3 Then
       If Check1.Value = 0 Then
          RsTxn_CDC_GSLDetail.MoveFirst
          RsTxn_CDC_GSLDetail.Find "CDCNo = " & Val(MSHFlexGrid2.TextMatrix(I, 12))
          If Not RsTxn_CDC_GSLDetail.EOF Then
             RsTxn_CDC_GSLDetail!DeliveredQty = RsTxn_CDC_GSLDetail!DeliveredQty - Val(MSHFlexGrid2.TextMatrix(I, 9))
             RsTxn_CDC_GSLDetail!Flag = "P"
             RsTxn_CDC_GSLDetail.Update
             RsTxn_CDC_GSLDetail.Requery
          End If
       End If
    End If


Next
tmp = "Delete from Txn_CxTF_ClientDetail where  sm_prefix = '" & TxtPreFix & "' And TxnType = 'W' And  CxTF_DetailsID = " & TxtCxTFNo
Call TmpTable
tmp = "Delete from Txn_CxTF_GSL where  sm_prefix = '" & TxtPreFix & "' And TxnType = 'W' And  CxTF_DetailsID = " & TxtCxTFNo
Call TmpTable
tmp = "Delete from Txn_CxTF_TruckDetail where  sm_prefix = '" & TxtPreFix & "' And TxnType = 'W'  And  CxTF_DetailsID = " & TxtCxTFNo
Call TmpTable
tmp = "Delete from Txn_CxTF_BOE Where CxTF_DetailsID=" & Val(TxtCxTFNo.Text) & " And LocationID=" & mLocationID
Call TmpTable

tmp = "select * from Txn_CxTF_GSL where  sm_prefix = '" & TxtPreFix & "' And TxnType = 'W'  And  CxTF_DetailsID = " & TxtCxTFNo
Call TmpTable
TmpRs.Close

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

Private Sub MSHFlexGrid4_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyC And Shift = vbCtrlMask Then
   With MSHFlexGrid4
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

Private Sub MSHFlexGrid9_Click()
FrameCRDetail.Visible = False
If SaveCmd.Enabled = True Then Exit Sub

If Val(MSHFlexGrid9.TextMatrix(MSHFlexGrid9.RowSel, 0)) = 0 Then Exit Sub

tmp = "Select TCR.NewCRID,TCR.SM_Prefix,TCR.BatchID from Txn_NewCashReceiptTIDetail TCRD"
tmp = tmp & " Inner join Txn_NewCashReceipt TCR on TCR.NewCRID=TCRD.NewCRID And TCRD.SM_Prefix=TCR.SM_Prefix"
tmp = tmp & " Where TCR.ReturnFlag<>'Y' And TIID = " & MSHFlexGrid9.TextMatrix(MSHFlexGrid9.RowSel, 0)
tmp = tmp & " Order by  TCRD.Concurrency Desc"

Call Tmp7Table

If TmpRs7.RecordCount > 0 Then
   FrameCRDetail.Visible = True
   lblVCRID.Caption = TmpRs7!NewCRID
   LblVPrefix.Caption = TmpRs7!SM_Prefix
   LblVBatchID.Caption = TmpRs7!BatchID
End If

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)

'If SSTab1.Tab = 0 Then
'   If SaveCmd.Enabled = True Then
'      Call UpdateRRNGridStock
'   End If
'End If
If SSTab1.Tab = 3 Then
   If SaveCmd.Enabled = True Then
      Call Grid_Head_Error
      If mExchangeTypeID = 1 Then
         Call UpdateErrorGrid
      End If
   End If
End If

End Sub
Private Function UpdateGSLBalChkGrid() As Boolean
Dim Retval As Boolean
Retval = True

Call Grid_Head_GSLBalChk
xRow = 1
For I = 2 To MSHFlexGrid4.Rows - 1
'    Frame1.Visible = True
    If ChkGrid(MSHFlexGrid4.TextMatrix(I, 14)) = False Then
       MSHFlexGrid7.TextMatrix(xRow, 3) = Val(MSHFlexGrid7.TextMatrix(xRow, 3)) + Val(MSHFlexGrid4.TextMatrix(I, 6))
       MSHFlexGrid7.TextMatrix(xRow, 4) = Val(MSHFlexGrid7.TextMatrix(xRow, 2)) - Val(MSHFlexGrid7.TextMatrix(xRow, 3))
    Else
       MSHFlexGrid7.Rows = MSHFlexGrid7.Rows + 1
       MSHFlexGrid7.TextMatrix(xRow, 0) = MSHFlexGrid4.TextMatrix(I, 14)
       MSHFlexGrid7.TextMatrix(xRow, 1) = MSHFlexGrid4.TextMatrix(I, 1) & "~" & MSHFlexGrid4.TextMatrix(I, 2) & "~" & MSHFlexGrid4.TextMatrix(I, 3)
       MSHFlexGrid7.TextMatrix(xRow, 2) = Val(MSHFlexGrid4.TextMatrix(I, 4))
       MSHFlexGrid7.TextMatrix(xRow, 3) = Val(MSHFlexGrid4.TextMatrix(I, 6))
       MSHFlexGrid7.TextMatrix(xRow, 4) = Val(MSHFlexGrid7.TextMatrix(xRow, 2)) - Val(MSHFlexGrid7.TextMatrix(xRow, 3))
       xRow = xRow + 1
    End If
Next

For y = 1 To MSHFlexGrid7.Rows - 1
    If Val(MSHFlexGrid7.TextMatrix(y, 4)) < 0 Then
       Retval = False
       For a = 0 To MSHFlexGrid7.Cols - 1
           MSHFlexGrid7.Col = a
           MSHFlexGrid7.row = y
           MSHFlexGrid7.CellBackColor = vbRed
       Next
       'Exit For
    End If
Next

UpdateGSLBalChkGrid = Retval

End Function
Private Function ChkGrid(mGSLCHKID_ As Variant) As Boolean
Dim Retval As Boolean
Retval = True

For y = 1 To MSHFlexGrid7.Rows - 1
    If mGSLCHKID_ = MSHFlexGrid7.TextMatrix(y, 0) Then
       Retval = False
       xRow = y
       Exit For
    End If
Next

ChkGrid = Retval


End Function
Private Sub UpdateErrorGrid()
Call Grid_Head_Error
For I = 2 To MSHFlexGrid2.Rows - 1
    If mExchangeTypeID = 1 Then
       If MSHFlexGrid2.TextMatrix(I, 11) <> "" Then
          Call UpdateDataISIN_WHR(MSHFlexGrid2.TextMatrix(I, 4), Val(MSHFlexGrid2.TextMatrix(I, 8)), "R", 0)
       End If
   Else
       If MSHFlexGrid2.TextMatrix(I, 2) <> "" Then
          Call UpdateDataISIN_WHR(MSHFlexGrid2.TextMatrix(I, 2), Val(MSHFlexGrid2.TextMatrix(I, 8)), "R", 0)
       End If
    End If
Next

For I = 2 To MSHFlexGrid4.Rows - 1
    If MSHFlexGrid4.TextMatrix(I, 12) <> "" Then
       If mExchangeTypeID = 1 Then
          
          Dim strSplitString() As String
          Dim mGSLClient As String
          strSplitString = Split(MSHFlexGrid4.TextMatrix(I, 12), "~")
          
          If UBound(strSplitString) = 1 Then
             mGSLClient = strSplitString(1)
          End If
          'Call UpdateDataISIN_WHR(MSHFlexGrid4.TextMatrix(I, 12), Val(MSHFlexGrid4.TextMatrix(I, 10)), "G", Val(MSHFlexGrid4.TextMatrix(I, 8)))
          Call UpdateDataISIN_WHR(mGSLClient, Val(MSHFlexGrid4.TextMatrix(I, 10)), "G", Val(MSHFlexGrid4.TextMatrix(I, 8)))
       Else
          If Val(MSHFlexGrid4.TextMatrix(I, 17)) = 1 Then
             Call UpdateDataISIN_WHR(MSHFlexGrid4.TextMatrix(I, 12), Val(MSHFlexGrid4.TextMatrix(I, 10)), "G", Val(MSHFlexGrid4.TextMatrix(I, 8)))
          Else
             Call UpdateDataISIN_WHR(MSHFlexGrid4.TextMatrix(I, 12), Val(MSHFlexGrid4.TextMatrix(I, 10)), "G", Val(MSHFlexGrid4.TextMatrix(I, 8)))
          End If
       End If
    End If
Next


For u = 2 To MSHFlexGrid5.Rows - 1
    MSHFlexGrid5.TextMatrix(0, 2) = Val(MSHFlexGrid5.TextMatrix(0, 2)) + Val(MSHFlexGrid5.TextMatrix(u, 2))
    MSHFlexGrid5.TextMatrix(0, 3) = Val(MSHFlexGrid5.TextMatrix(0, 3)) + Val(MSHFlexGrid5.TextMatrix(u, 3))
    MSHFlexGrid5.TextMatrix(0, 4) = Val(MSHFlexGrid5.TextMatrix(0, 4)) + Val(MSHFlexGrid5.TextMatrix(u, 4))
    MSHFlexGrid5.TextMatrix(0, 5) = Val(MSHFlexGrid5.TextMatrix(0, 5)) + Val(MSHFlexGrid5.TextMatrix(u, 5))
    MSHFlexGrid5.TextMatrix(0, 6) = Val(MSHFlexGrid5.TextMatrix(0, 6)) + Val(MSHFlexGrid5.TextMatrix(u, 6))
    
    MSHFlexGrid5.TextMatrix(u, 2) = Format(MSHFlexGrid5.TextMatrix(u, 2), "######0.000")
    MSHFlexGrid5.TextMatrix(u, 3) = Format(MSHFlexGrid5.TextMatrix(u, 3), "######0.000")
    MSHFlexGrid5.TextMatrix(u, 4) = Format(MSHFlexGrid5.TextMatrix(u, 4), "######0.000")
    MSHFlexGrid5.TextMatrix(u, 5) = Format(MSHFlexGrid5.TextMatrix(u, 5), "######0.000")
    MSHFlexGrid5.TextMatrix(u, 6) = Format(MSHFlexGrid5.TextMatrix(u, 6), "######0.000")
    
    MSHFlexGrid5.TextMatrix(0, 2) = Format(MSHFlexGrid5.TextMatrix(0, 2), "######0.000")
    MSHFlexGrid5.TextMatrix(0, 3) = Format(MSHFlexGrid5.TextMatrix(0, 3), "######0.000")
    MSHFlexGrid5.TextMatrix(0, 4) = Format(MSHFlexGrid5.TextMatrix(0, 4), "######0.000")
    MSHFlexGrid5.TextMatrix(0, 5) = Format(MSHFlexGrid5.TextMatrix(0, 5), "######0.000")
    MSHFlexGrid5.TextMatrix(0, 6) = Format(MSHFlexGrid5.TextMatrix(0, 6), "######0.000")
    MSHFlexGrid5.row = u
    If Val(MSHFlexGrid5.TextMatrix(u, 5)) <> 0 Then
       MSHFlexGrid5.Col = 5
       MSHFlexGrid5.CellBackColor = vbYellow
       MSHFlexGrid5.Col = 1
       MSHFlexGrid5.CellBackColor = vbYellow
   End If
Next

End Sub

Private Sub UpdateDataISIN_WHR(ISIN_ As String, Wght_ As Double, Mod_ As String, AdjWght_ As Double)

Dim u As Variant
For u = 2 To MSHFlexGrid5.Rows - 1
    If LCase(MSHFlexGrid5.TextMatrix(u, 1)) = LCase(Trim(ISIN_)) Then
       If Mod_ = "R" Then
          MSHFlexGrid5.TextMatrix(u, 2) = Val(MSHFlexGrid5.TextMatrix(u, 2)) + Wght_
       Else
          MSHFlexGrid5.TextMatrix(u, 3) = Val(MSHFlexGrid5.TextMatrix(u, 3)) + Wght_
          MSHFlexGrid5.TextMatrix(u, 4) = Val(MSHFlexGrid5.TextMatrix(u, 4)) + AdjWght_
          If Val(MSHFlexGrid5.TextMatrix(u, 3)) > Val(MSHFlexGrid5.TextMatrix(u, 2)) Then
             MSHFlexGrid5.TextMatrix(u, 5) = Val(MSHFlexGrid5.TextMatrix(u, 3)) - (Val(MSHFlexGrid5.TextMatrix(u, 2)))
             MSHFlexGrid5.TextMatrix(u, 6) = 0
          Else
             MSHFlexGrid5.TextMatrix(u, 6) = Val(MSHFlexGrid5.TextMatrix(u, 2)) - (Val(MSHFlexGrid5.TextMatrix(u, 3)))
             MSHFlexGrid5.TextMatrix(u, 6) = Format(MSHFlexGrid5.TextMatrix(u, 6), "######0.000")
          End If
       End If
       Exit Sub
    End If
Next
If Mod_ = "R" Then
   u = MSHFlexGrid5.Rows - 1
   MSHFlexGrid5.Rows = MSHFlexGrid5.Rows + 1
   MSHFlexGrid5.TextMatrix(u, 1) = ISIN_
   MSHFlexGrid5.TextMatrix(u, 2) = Wght_
End If

End Sub

Private Sub UpdateRRNGridStock()
Diff_wt = 0
For I = 2 To MSHFlexGrid2.Rows - 1
    MSHFlexGrid2.TextMatrix(I, 9) = 0 '"Delived Qty"
    MSHFlexGrid2.TextMatrix(I, 15) = 0 '"Delived Bags"
Next
MSHFlexGrid2.TextMatrix(0, 9) = 0 '"Delived Qty"
MSHFlexGrid2.TextMatrix(0, 15) = 0 '"Delived Bags"
For I = 2 To MSHFlexGrid4.Rows - 1
    If MSHFlexGrid4.TextMatrix(I, 12) <> "" Then
       If mExchangeTypeID = 1 Then
'          If MSHFlexGrid4.TextMatrix(i, 17) = 1 Then
'             Call UpdateDataClient(MSHFlexGrid4.TextMatrix(i, 16), Val(MSHFlexGrid4.TextMatrix(i, 10)), 0)
'          Else
             Call UpdateDataClient(MSHFlexGrid4.TextMatrix(I, 16), Val(MSHFlexGrid4.TextMatrix(I, 10)), Val(MSHFlexGrid4.TextMatrix(I, 6)))
'          End If
       Else
'          If MSHFlexGrid4.TextMatrix(i, 17) = 1 Then
'             Call UpdateDataClient(MSHFlexGrid4.TextMatrix(i, 12), Val(MSHFlexGrid4.TextMatrix(i, 10)), 0)
'          Else
             Call UpdateDataClient(MSHFlexGrid4.TextMatrix(I, 12), Val(MSHFlexGrid4.TextMatrix(I, 10)), Val(MSHFlexGrid4.TextMatrix(I, 6)))
'          End If
       End If
    End If
Next
End Sub
Sub UpdateDataClient(ISIN_ As Variant, EWeight As Double, DBags As Double)

Dim LstRow, Ci, Bal, PQty, Cl_IS As Variant
PQty = EWeight


For I = 2 To MSHFlexGrid2.Rows - 1
   
    If mExchangeTypeID = 1 Then
       Cl_IS = MSHFlexGrid2.TextMatrix(I, 11)
    ElseIf mExchangeTypeID = 2 Then
       If Check1.Value = 0 Then
          Cl_IS = MSHFlexGrid2.TextMatrix(I, 2)
       Else
          Cl_IS = MSHFlexGrid2.TextMatrix(I, 4)
       End If
    Else
       Cl_IS = MSHFlexGrid2.TextMatrix(I, 1)
    End If
    If LCase(Cl_IS) = LCase(ISIN_) Then
       Bal = Val(MSHFlexGrid2.TextMatrix(I, 8)) - Val(MSHFlexGrid2.TextMatrix(I, 9))
       If Bal > 0 Then
          If Bal >= (EWeight) Then
             MSHFlexGrid2.TextMatrix(I, 9) = Val(MSHFlexGrid2.TextMatrix(I, 9)) + EWeight '"Weight"
             MSHFlexGrid2.TextMatrix(0, 9) = Val(MSHFlexGrid2.TextMatrix(0, 9)) + EWeight '"Weight"
             EWeight = 0
          Else
             EWeight = EWeight - Bal
             MSHFlexGrid2.TextMatrix(I, 9) = Val(MSHFlexGrid2.TextMatrix(I, 9)) + Bal '"Weight"
             MSHFlexGrid2.TextMatrix(0, 9) = Val(MSHFlexGrid2.TextMatrix(0, 9)) + Bal '"Weight"
          End If
       End If
       If EWeight = 0 Then Exit For
       LstRow = I
    End If
Next
If EWeight > 0 And mExchangeTypeID = 1 Then
   MSHFlexGrid2.TextMatrix(LstRow, 9) = Val(MSHFlexGrid2.TextMatrix(LstRow, 9)) + EWeight '"Weight"
   MSHFlexGrid2.TextMatrix(0, 9) = Val(MSHFlexGrid2.TextMatrix(0, 9)) + EWeight '"Weight"
End If
' Adj No of Bags
If mExchangeTypeID = 1 Or (mExchangeTypeID = 2 And Check1.Value = 1) Then
   PQty = DBags
   
   For I = 2 To MSHFlexGrid2.Rows - 1
       
       If mExchangeTypeID = 1 Then
          Cl_IS = MSHFlexGrid2.TextMatrix(I, 11)
       ElseIf mExchangeTypeID = 2 Then
          If Check1.Value = 0 Then
             Cl_IS = MSHFlexGrid2.TextMatrix(I, 2)
          Else
             Cl_IS = MSHFlexGrid2.TextMatrix(I, 4)
          End If
       
       Else
          Cl_IS = MSHFlexGrid2.TextMatrix(I, 1)
       End If
       If Cl_IS = ISIN_ Then
           Bal = Val(MSHFlexGrid2.TextMatrix(I, 14)) - Val(MSHFlexGrid2.TextMatrix(I, 15))
           If Bal > 0 Then
              If Bal >= (DBags) Then
                 MSHFlexGrid2.TextMatrix(I, 15) = Val(MSHFlexGrid2.TextMatrix(I, 15)) + DBags '"Weight"
                 MSHFlexGrid2.TextMatrix(0, 15) = Val(MSHFlexGrid2.TextMatrix(0, 15)) + DBags '"Weight"
                 DBags = 0
              Else
                 DBags = DBags - Bal
                 MSHFlexGrid2.TextMatrix(I, 15) = Val(MSHFlexGrid2.TextMatrix(I, 15)) + Bal '"Weight"
                 MSHFlexGrid2.TextMatrix(0, 15) = Val(MSHFlexGrid2.TextMatrix(0, 15)) + Bal '"Weight"
              End If
           End If
           If DBags = 0 Then Exit For
           LstRow = I
        End If
    Next
    If DBags <> 0 Then
       Diff_wt = 1
    End If
    If DBags > 0 And mExchangeTypeID = 1 Then
       MSHFlexGrid2.TextMatrix(LstRow, 15) = Val(MSHFlexGrid2.TextMatrix(LstRow, 15)) + DBags '"Weight"
       MSHFlexGrid2.TextMatrix(0, 15) = Val(MSHFlexGrid2.TextMatrix(0, 15)) + DBags '"Weight"
   End If
End If
If EWeight <> 0 Then
   Diff_wt = 1
End If
End Sub



Private Sub TxtClientID_LostFocus()
CmbClientDP.Visible = False
If TxtClientID = "" Then Exit Sub
If CmbExchangeTypeID.Text = "" Then Exit Sub

If CmbCommodityID.Text = "" Then Exit Sub
If CmbCMPID.Text = "" Then Exit Sub
If CmbExchangeTypeID.Text = "" Then Exit Sub
 
Dim ans As Variant
ans = GetClientNameDtl(TxtClientID, mExchangeTypeID)

If ans(0, 0) = False Then
   MsgBox "Not found!!!"
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
      " And  (((Mst_Client.ExchangeTypeID)=" & mExchangeTypeID & ") AND ((Mst_Client.ClientIDRow)='" & TxtClientID & "'));"
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

Private Sub CmbClientDP_LostFocus()

Dim Str_() As String
Str_ = Split(CmbClientDP.Text, "~")
If UBound(Str_) <> 1 Then
   MsgBox "Invalid selection!!!"
   CmbClientDP.SetFocus
   Exit Sub
End If
TmpRs1.MoveFirst
TmpRs1.Find "DepositoryID  = '" & Str_(0) & "'"
If TmpRs1.EOF Then
   MsgBox "Invalid selection!!!"
   CmbClientDP.SetFocus
   Exit Sub
End If


TxtClientName = TmpRs1!ClientName
TxtDPID = TmpRs1!DepositoryID
TxtDPName = TmpRs1!DepositoryName
CmbClientDP.Visible = False
mClientIDRow = TmpRs1!ClientID


tmp = "SELECT Sum(Mst_GSL.BalanceWeight) AS SBalWeight, Sum(Mst_GSL.BalanceBags) AS SBalBags, "
tmp = tmp & "  Sum(Mst_GSL.DepositeWeight) AS SDWeight,Sum(Mst_GSL.DepositeBags) AS SDBags FROM Mst_GSL "
tmp = tmp & " WHERE (((Mst_GSL.SM_Prefix)='" & TxtPreFix & "') AND ((Mst_GSL.CMPID)=" & mCMPID & ") AND "
tmp = tmp & " ((Mst_GSL.CommodityID)=" & mCommodityID & " ) AND ((Mst_GSL.ClientID)=" & mClientIDRow & ") AND "
tmp = tmp & " ((Mst_GSL.ExchangeTypeID)= " & mExchangeTypeID & ") AND ((Mst_GSL.State)='J') AND ((Mst_GSL.Status)='O'));"
Call TmpTable
'If TmpRs.RecordCount = 0 Then Exit Sub
TxtRRNQty = IIf(IsNull(TmpRs.Fields(2)), 0, TmpRs.Fields(2))
TxtBalanceQty = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0)) 'TmpRs.Fields(0)
TxtBalanceBags = IIf(IsNull(TmpRs.Fields(1)), 0, TmpRs.Fields(1)) 'TmpRs.Fields(1)
TxtRRNBags = IIf(IsNull(TmpRs.Fields(3)), 0, TmpRs.Fields(3)) 'TmpRs.Fields(3)
TxtRRNNo = mClientIDRow
CmdAdd2ListRRN.SetFocus
End Sub

Private Sub CmbGodownNo_GotFocus()
If CmbCMPID.Text = "" Then
   MsgBox "Select a CMP!!!"
   CmbCMPID.SetFocus
   Exit Sub
End If
If CmbCommodityID.Text <> "" Then
   tmp = CmbGodownNo.Text
   Call TableMst_Godown(" Where CloseDate is null And CMPID = " & mCMPID)
   CmbGodownNo.Clear
   If RsMst_Godown.RecordCount = 0 Then
      MsgBox "No record found!!!"
      Exit Sub
   End If
   For I = 1 To RsMst_Godown.RecordCount
       CmbGodownNo.AddItem RsMst_Godown!GodownNo
       RsMst_Godown.MoveNext
   Next
   CmbGodownNo.Text = tmp
Else
   MsgBox "Please select a Commodity!!!"
   CmbCommodityID.SetFocus
   Exit Sub
End If
End Sub
'----
Private Sub CmbGodownNo_LostFocus()
If CmbGodownNo.Text = "" Then
   Exit Sub
End If
RsMst_Godown.MoveFirst
RsMst_Godown.Find "GodownNo = '" & CmbGodownNo.Text & "'"
If RsMst_Godown.EOF Then
   MsgBox "Invalid selection!!!"
   CmbGodownNo.SetFocus
   Exit Sub
End If
mGodownID = RsMst_Godown!GodownID
mGodownYearlyFlag = IIf(IsNull(RsMst_Godown!YearlyBillingFlg), 0, RsMst_Godown!YearlyBillingFlg)

If LCase(CmbExchangeTypeID.Text) = "non ncdex" Then
   For I = 2 To MSHFlexGrid4.Rows - 1
       If MSHFlexGrid4.TextMatrix(I, 22) = "" Then Exit For
       If Val(MSHFlexGrid4.TextMatrix(I, 22)) <> Val(mGodownYearlyFlag) Then
          MsgBox "All godown through which wihtdrawal is going should have same yearly billing flag!!!"
          CmbGodownNo.SetFocus
          Exit Sub
       End If
   Next
End If

End Sub

Private Sub CmbLot_GotFocus()
If CmbStack.Text = "" Then
   MsgBox "Please select Stack No!!!"
   CmbStack.SetFocus
   CmbLot.Text = ""
   Exit Sub
End If
tmp = CmbLot.Text
If Check2.Value = 1 Then
   Call TableMst_GSL(" Where  (SpillageFlag = '1' And CMPID = " & mCMPID & " And CommodityID = " & mCommodityID & "  And GodownID = " & mGodownID & " And StackNo = '" & CmbStack.Text & "' And  " & Gen_ExpireGSLFlag & "  )") 'And SCMFLag = '" & ChkSCMFlag.Value & "'
ElseIf Check4.Value = 1 Then
   Call TableMst_GSL(" Where  (MadeupFlag = '1' And ExchangeTypeId = " & mExchangeTypeID & " And CMPID = " & mCMPID & " And CommodityID = " & mCommodityID & "  And GodownID = " & mGodownID & " And StackNo = '" & CmbStack.Text & "' And  " & Gen_ExpireGSLFlag & " And SCMFLag = '" & ChkSCMFlag.Value & "' And isnull(Connected_Mst_GSL_ID,0) not in (Select Mst_GSL_ID From Txn_PledgeCreation_GSL Where SM_Prefix = '" & TxtPreFix & "' And PledgeNoOfBags > 0) )")
ElseIf mExchangeTypeID <> 1 Then
   If Check1.Value = 0 Then
      If mCommodityID <> 31 And mCommodityID <> 114 Then
         Call TableMst_GSL(" Where ((MadeupFlag = '0' And ExchangeTypeId = " & mExchangeTypeID & " And CMPID = " & mCMPID & " And CommodityID = " & mCommodityID & " And GodownID = " & mGodownID & " And StackNo = '" & CmbStack.Text & "' And ISINID <> '0' And  " & Gen_ExpireGSLFlag & " And SCMFLag = '" & ChkSCMFlag.Value & "' )) ")
      Else
         Call TableMst_GSL(" Where ((MadeupFlag = '0' And ExchangeTypeId = " & mExchangeTypeID & " And CMPID = " & mCMPID & " And CommodityID in (31,114) And GodownID = " & mGodownID & " And StackNo = '" & CmbStack.Text & "' And ISINID <> '0' And  " & Gen_ExpireGSLFlag & " And SCMFLag = '" & ChkSCMFlag.Value & "' )) ")
      End If
   ElseIf Check1.Value = 1 Then
      If mExchangeTypeID = 3 Then
         Call TableMst_GSL(" Where ((MadeupFlag = '0' And ExchangeTypeId = " & mExchangeTypeID & " And CMPID = " & mCMPID & " And CommodityID = " & mCommodityID & " And GodownID = " & mGodownID & " And StackNo = '" & CmbStack.Text & "' And SCMFLag = '" & ChkSCMFlag.Value & "' And ISINID <> '0' And ( " & Gen_ExpireGSLFlag & "  ))) ")
      Else
         ' change done by mayur on 12.03.2010 ' Unjha - narendra rejected lot f3 combo not showing rejection
         'Call TableMst_GSL(" Where ((ExchangeTypeId = " & mExchangeTypeID & " And CMPID = " & mCMPID & " And CommodityID = " & mCommodityID & " And GodownID = " & mGodownID & " And StackNo = '" & CmbStack.Text & "' And ISINID <> '0' And ( " & Gen_ExpireGSLFlag & "  And Flag = 'J'))) ")
         
         'Call TableMst_GSL(" Where ((ExchangeTypeId = " & mExchangeTypeID & " And CMPID = " & mCMPID & " And CommodityID = " & mCommodityID & " And GodownID = " & mGodownID & " And StackNo = '" & CmbStack.Text & "' And ISINID <> '0' And ( " & Gen_ExpireGSLFlag & "  And Flag = 'J'))) ")
         If mCommodityID <> 31 And mCommodityID <> 114 Then
            Call TableMst_GSL(" Where ((MadeupFlag = '0' And ExchangeTypeId = " & mExchangeTypeID & " And CMPID = " & mCMPID & " And CommodityID = " & mCommodityID & " And GodownID = " & mGodownID & " And StackNo = '" & CmbStack.Text & "'  And SCMFLag = '" & ChkSCMFlag.Value & "' And  State = 'J' And Status = 'O')) ")
         Else
            Call TableMst_GSL(" Where ((MadeupFlag = '0' And ExchangeTypeId = " & mExchangeTypeID & " And CMPID = " & mCMPID & " And CommodityID in (31,114) And GodownID = " & mGodownID & " And StackNo = '" & CmbStack.Text & "'  And SCMFLag = '" & ChkSCMFlag.Value & "' And  State = 'J' And Status = 'O')) ")
         End If
      End If
   End If
Else
   Call TableMst_GSL(" Where  (MadeupFlag = '0' And ExchangeTypeId = " & mExchangeTypeID & " And CMPID = " & mCMPID & " And CommodityID = " & mCommodityID & "  And GodownID = " & mGodownID & " And StackNo = '" & CmbStack.Text & "' And  " & Gen_ExpireGSLFlag & " And SCMFLag = '" & ChkSCMFlag.Value & "'  )")
End If
CmbLot.Clear
If RsMst_GSL.RecordCount = 0 Then
   MsgBox "No record found!!!"
   Exit Sub
End If
For I = 1 To RsMst_GSL.RecordCount
    CmbLot.AddItem RsMst_GSL!LotNo
    RsMst_GSL.MoveNext
Next I
CmbLot.Text = tmp
End Sub
'----
Private Sub CmbLot_LostFocus()
Dim ans As Variant
If CmbLot.Text = "" Then
   mWHRNo = ""
   Exit Sub
End If
RsMst_GSL.MoveFirst
RsMst_GSL.Find "LotNo = '" & CmbLot.Text & "'"
If RsMst_GSL.EOF Then
   MsgBox "Invalid selection!!!"
   CmbLot.SetFocus
   Exit Sub
End If

mGSLID = RsMst_GSL!Mst_GSL_ID
mWHRNo = IIf(IsNull(RsMst_GSL!WHRNo), "", RsMst_GSL!WHRNo)

If Val(RsMst_GSL!DepositeBags) < 0 Or Val(RsMst_GSL!CDTFBags) < 0 Or Val(RsMst_GSL!BalanceBags) < 0 Then
   MsgBox "Some -ve figure issue with this GSL. Please Contact CMG Department for rectification in this GSL!!!"
   CmbLot.SetFocus
   Exit Sub
End If

TxtGSLAvailableNofBags = IIf(IsNull(RsMst_GSL!DepositeBags), 0, RsMst_GSL!DepositeBags) - IIf(IsNull(RsMst_GSL!CDTFBags), 0, RsMst_GSL!CDTFBags)
TxtGSLAvailableNofBags = Val(TxtGSLAvailableNofBags) - getRemainingGSLBag(mGSLID)
TxtGSLAvailableGrossWgt = IIf(IsNull(RsMst_GSL!DepositeWeight), 0, RsMst_GSL!DepositeWeight) - IIf(IsNull(RsMst_GSL!CDTFWeight), 0, RsMst_GSL!CDTFWeight)
TxtGSLAvailableGrossWgt = Val(TxtGSLAvailableGrossWgt) - getRemainingGSLWt(mGSLID)

mPledgeID = IIf(IsNull(RsMst_GSL!PledgeNo), "", RsMst_GSL!PledgeNo)
TxtGSLPledgedNofBags = IIf(IsNull(RsMst_GSL!PledgeBags), "", RsMst_GSL!PledgeBags)
TxtGSLPledgedWeight = IIf(IsNull(RsMst_GSL!PledgeWeight), "", RsMst_GSL!PledgeWeight)
TxtGSLISINno = ""
txtmGSLClientIDRow = 0

mGSLContractID = IIf(IsNull(RsMst_GSL!NContractID), 0, RsMst_GSL!NContractID)

If mExchangeTypeID <> 1 Then
   If Check1.Value = 0 Then
      TxtGSLISINno = IIf(IsNull(RsMst_GSL!ISINID), "", RsMst_GSL!ISINID)
      txtmGSLClientIDRow = IIf(IsNull(RsMst_GSL!ISINID), 0, RsMst_GSL!ISINID)
   Else
      txtmGSLClientIDRow = IIf(IsNull(RsMst_GSL!ClientID), 0, RsMst_GSL!ClientID)
      If mExchangeTypeID = 3 Then
         TxtGSLISINno = IIf(IsNull(RsMst_GSL!ISINID), "", RsMst_GSL!ISINID)
      End If
   End If
   '--Added by sagar to check RRn having Comtrack flag
   If mExchangeTypeID = 2 And MSHFlexGrid10.Rows > 3 Then
      TxtGSLISINno = IIf(IsNull(RsMst_GSL!ISINID), "", RsMst_GSL!ISINID)
      'TxtGSLISINno.Visible = True
      'CmbGCTDClient.Visible = True
      If Check2.Value = 0 Then
         If Check4.Value = 1 Then
            Dim Stra() As String
            Stra = GetGSLByGLSID(IIf(IsNull(RsMst_GSL!Connected_Mst_GSL_ID), 0, RsMst_GSL!Connected_Mst_GSL_ID), TxtPreFix)
            If UBound(Stra) > 0 Then
               TxtGSLISINno = Stra(10)
            End If
         End If
         
         CmbGCTDClient.Text = GetHeaderISINRRN(TxtGSLISINno)
         CmbGCTDClient.Enabled = False
         
      Else
         CmbGCTDClient.Enabled = True
      End If
   Else
      CmbGCTDClient.Enabled = True
   End If
Else
   TxtGSLISINno = GetClientName(txtmGSLClientIDRow) 'IIf(IsNull(RsMst_GSL!ClientIDRow), "", RsMst_GSL!ClientIDRow)
   txtmGSLClientIDRow = IIf(IsNull(RsMst_GSL!ClientID), 0, RsMst_GSL!ClientID)
   TxtWHRGenBags = IIf(IsNull(RsMst_GSL!WHRBags), 0, RsMst_GSL!WHRBags) - IIf(IsNull(RsMst_GSL!CDTFBags), 0, RsMst_GSL!CDTFBags)
   TxtWHRGenQty = IIf(IsNull(RsMst_GSL!WHRWeight), 0, RsMst_GSL!WHRWeight) - IIf(IsNull(RsMst_GSL!CDTFWeight), 0, RsMst_GSL!CDTFWeight)
End If

If mCommodityID = 31 Or mCommodityID = 114 Then
   Call TableMst_CommodityDetail("Where CommodityId = " & RsMst_GSL!CommodityID & " and ExchangeTypeId = " & mExchangeTypeID)
   If RsMst_CommodityDetail.RecordCount > 0 Then
      mChilliGunnyWeight = RsMst_CommodityDetail!GunnyWt
   End If
End If

End Sub
Private Sub CmbStack_GotFocus()
If CmbGodownNo.Text = "" Or mGodownID = "" Then    'IsNull(mGodownID)
   MsgBox "Please select Godown No!!!"
   Exit Sub
End If
tmp = CmbStack.Text
Call TableMst_GSL(" Where CMPID = " & mCMPID & " And GodownID = " & mGodownID, "D")
CmbStack.Clear
If RsMst_GSL.RecordCount = 0 Then
   MsgBox "No record found!!!"
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
   MsgBox "Invalid selection!!!"
   CmbStack.SetFocus
   Exit Sub
End If

End Sub
Private Sub CmdAdd2ListRRN_Click()
If CmbExchangeTypeID.Text = "" Then
   MsgBox "Select a Mode!!!"
   CmbExchangeTypeID.SetFocus
   Exit Sub
End If

If CmbCMPID.Text = "" Then
   MsgBox "Select a CMP!!!"
   CmbCMPID.SetFocus
   Exit Sub
End If
If CmbCommodityID.Text = "" Then
   MsgBox "Select a Commodity!!!"
   CmbCommodityID.SetFocus
   Exit Sub
End If
Dim r As Variant
If Check1.Value = 0 Then
   If TxtRRNNo = "" Then
      If mExchangeTypeID = 1 Then 'NON NCDEX
         MsgBox "Blank WHR Number not allowed!!! "
      ElseIf mExchangeTypeID = 2 Then 'NCDEX
         MsgBox "Blank RRN Number not allowed!!! "
      ElseIf mExchangeTypeID = 3 Then 'SPOT
         MsgBox "Blank CDC Number not allowed!!! "
      End If
      TxtRRNNo.SetFocus
      Exit Sub
    End If
Else
    If TxtClientID = "" Then
       MsgBox "Blank Client ID not allowed!!!"
       TxtClientID.SetFocus
       Exit Sub
    End If
End If
If Check1.Value = 0 Then
   If CheckDuplicateRWC = True Then
     'MsgBox "Duplicate Client ID  not allowd !!!"
      If mExchangeTypeID = 1 Then 'NON NCDEX
         MsgBox "Duplicate WHR Number not allowed!!! "
      ElseIf mExchangeTypeID = 2 Then 'NCDEX
         MsgBox "Duplicate RRN Number not allowed!!! "
      ElseIf mExchangeTypeID = 3 Then 'SPOT
         MsgBox "Duplicate CDC Number not allowed!!! "
      End If
      TxtRRNNo.SetFocus
      Exit Sub
   End If
Else
   If CheckDuplicateClient = True Then
      MsgBox "Duplicate Client ID not allowed!!!"
      TxtClientID.SetFocus
      Exit Sub
   End If
End If

If Check1.Value = 0 Then
   If Val(TxtBalanceQty) <= 0 Then
      If mExchangeTypeID = 1 Then
         If Val(TxtBalanceBags) <= 0 Then
            MsgBox "Warning : No Quantity is available for delivery !!!"
            'TxtRRNNo.SetFocus
            'Exit Sub
         End If
      ElseIf mExchangeTypeID = 2 Then
         MsgBox "All Quantity delivered!!!"
         TxtRRNNo.SetFocus
         Exit Sub
      Else
         MsgBox "Warning :  All Quantity delivered!!!"
         'TxtRRNNo.SetFocus
         'Exit Sub
      End If
   Else
      If mExchangeTypeID = 1 Then
         If Val(TxtBalanceBags) <= 0 Then
            MsgBox "Warning : No Bag is available for delivery !!!"
            'TxtRRNNo.SetFocus
            'Exit Sub
         End If
      End If
   End If
End If


If mExchangeTypeID = 3 Then
   If Val(TxtBalanceQty) = 0 Then
      MsgBox "Warning : All Quantity delivered!!!"
      'TxtRRNNo.SetFocus
      'Exit Sub
   End If
   If Val(TxtBalanceBags) <= 0 Then
      MsgBox "Warning : All Bags delivered!!!"
      'TxtRRNNo.SetFocus
      'Exit Sub
   End If
End If

If mExchangeTypeID = 1 Then
   Dim strSplitString() As String
   strSplitString = Split(TxtRRNNo.Text, "-")
   tmp = UBound(strSplitString)
   If tmp <> 1 Then
      MsgBox "Not a valid WHR Number!!!"
      TxtRRNNo.SetFocus
      Exit Sub
   End If
   tmp = (Val(strSplitString(0)) * 1000000) + Val(strSplitString(1))
   RsTxn_WHRDetail.MoveFirst
   RsTxn_WHRDetail.Find "WHRNo = " & tmp
   ' 30;sept 2011 please check
'   If RsTxn_WHRDetail!PledgeID <> 0 Then
'      If MSHFlexGrid2.Rows > 3 Then
'         MsgBox "Error:This warehouse receipt is already pledged!!!"
'         TxtRRNNo.SetFocus
'         Exit Sub
'      End If
'      CmdAdd2ListRRN.Enabled = False
'   End If
End If
If CmdAdd2ListRRN.Caption = "Update List" Then
   r = MSHFlexGrid2.RowSel
   MSHFlexGrid2.TextMatrix(0, 7) = Val(MSHFlexGrid2.TextMatrix(0, 7)) - Val(MSHFlexGrid2.TextMatrix(r, 7))
   MSHFlexGrid2.TextMatrix(0, 8) = Val(MSHFlexGrid2.TextMatrix(0, 8)) - Val(MSHFlexGrid2.TextMatrix(r, 8))
   MSHFlexGrid2.TextMatrix(0, 9) = Val(MSHFlexGrid2.TextMatrix(0, 9)) - Val(MSHFlexGrid2.TextMatrix(r, 9))
   
   MSHFlexGrid2.TextMatrix(0, 13) = Val(MSHFlexGrid2.TextMatrix(0, 13)) - Val(MSHFlexGrid2.TextMatrix(r, 13))
   MSHFlexGrid2.TextMatrix(0, 14) = Val(MSHFlexGrid2.TextMatrix(0, 14)) - Val(MSHFlexGrid2.TextMatrix(r, 14))
   MSHFlexGrid2.TextMatrix(0, 15) = Val(MSHFlexGrid2.TextMatrix(0, 15)) - Val(MSHFlexGrid2.TextMatrix(r, 15))
   If mExchangeTypeID = 2 Then
      If mRRNCommTrackFlag = 1 Then
         Call DeleteFromRRNDetail(MSHFlexGrid2.TextMatrix(r, 12))
      End If
   End If
Else
   If mExchangeTypeID = 2 Then
      If mRRNCommTrackFlag = 1 Then
         Call FillRRNDetails(mRWCID)
      End If
   End If

   If mExchangeTypeID = 3 Then 'SPOT
      Call DisplayGSLData(mGSLID)
   End If
   r = MSHFlexGrid2.Rows - 1
   MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
End If

MSHFlexGrid2.TextMatrix(r, 1) = Trim(TxtRRNNo)
MSHFlexGrid2.TextMatrix(r, 2) = TxtISINno
MSHFlexGrid2.TextMatrix(r, 3) = TxtClientID
MSHFlexGrid2.TextMatrix(r, 4) = TxtClientName
MSHFlexGrid2.TextMatrix(r, 5) = TxtDPID
MSHFlexGrid2.TextMatrix(r, 6) = TxtDPName
MSHFlexGrid2.TextMatrix(r, 7) = TxtRRNQty
MSHFlexGrid2.TextMatrix(r, 8) = TxtBalanceQty
If mExchangeTypeID = 3 Then 'SPOT
   MSHFlexGrid2.TextMatrix(r, 9) = Val(TxtBalanceQty)
   MSHFlexGrid2.TextMatrix(r, 15) = Val(TxtBalanceBags)
End If
'MSHFlexGrid2.TextMatrix(r, 9) = (TxtBalanceQty)
MSHFlexGrid2.TextMatrix(r, 10) = TxtRRNRemarks
MSHFlexGrid2.TextMatrix(r, 11) = mClientIDRow
MSHFlexGrid2.TextMatrix(r, 12) = mRWCID
MSHFlexGrid2.TextMatrix(r, 13) = TxtRRNBags
MSHFlexGrid2.TextMatrix(r, 14) = TxtBalanceBags

 
MSHFlexGrid2.TextMatrix(r, 18) = mNAVCode
MSHFlexGrid2.TextMatrix(r, 19) = mRNAVCode

'If mExchangeTypeID = 1 Then
'   MSHFlexGrid2.TextMatrix(r, 12) = mRWCID
'ElseIf mExchangeTypeID = 2 Then
'   MSHFlexGrid2.TextMatrix(r, 12) = mRWCID
'ElseIf mExchangeTypeID = 3 Then
'   MSHFlexGrid2.TextMatrix(r, 12) = mRWCID
'End If
MSHFlexGrid2.TextMatrix(0, 7) = Val(MSHFlexGrid2.TextMatrix(0, 7)) + Val(MSHFlexGrid2.TextMatrix(r, 7))
MSHFlexGrid2.TextMatrix(0, 8) = Val(MSHFlexGrid2.TextMatrix(0, 8)) + Val(MSHFlexGrid2.TextMatrix(r, 8))
MSHFlexGrid2.TextMatrix(0, 9) = Val(MSHFlexGrid2.TextMatrix(0, 9)) + Val(MSHFlexGrid2.TextMatrix(r, 9))

MSHFlexGrid2.TextMatrix(0, 13) = Val(MSHFlexGrid2.TextMatrix(0, 13)) + Val(MSHFlexGrid2.TextMatrix(r, 13))
MSHFlexGrid2.TextMatrix(0, 14) = Val(MSHFlexGrid2.TextMatrix(0, 14)) + Val(MSHFlexGrid2.TextMatrix(r, 14))
MSHFlexGrid2.TextMatrix(0, 15) = Val(MSHFlexGrid2.TextMatrix(0, 15)) + Val(MSHFlexGrid2.TextMatrix(r, 15))


Call Clear_RWCTabControl
TxtRRNNo.Locked = False
If TxtRRNNo.Visible = True And TxtRRNNo.Enabled = True Then
   TxtRRNNo.SetFocus
End If
End Sub
Function CheckDuplicateRWC() As Boolean
Dim Retval As Boolean
Retval = False
If MSHFlexGrid2.Rows > 3 Then
   For u = 2 To MSHFlexGrid2.Rows - 1
       If MSHFlexGrid2.TextMatrix(u, 1) = Trim(TxtRRNNo) Then
          If LCase(CmdAdd2ListRRN.Caption) = "add to list" Then
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

CheckDuplicateRWC = Retval
End Function

Private Sub CmdGCTDAdd2List_Click()

Dim r As Variant

If (CmbGodownNo.Text) = "" Then
   MsgBox "Blank Godown Number not allowed!!! "
   If CmbGodownNo.Enabled = True Then
      CmbGodownNo.SetFocus
   End If
   Exit Sub
End If

If (CmbStack.Text) = "" Then
   MsgBox "Blank Stack not allowed!!! "
   CmbStack.SetFocus
   Exit Sub
End If

If (CmbLot.Text) = "" Then
   MsgBox "Blank Lot Number not allowed!!! "
   CmbLot.SetFocus
   Exit Sub
End If


If mExchangeTypeID = 2 Then 'NCDEX
   If Check1.Value = 0 Then
      If Check2.Value = 0 And Check4.Value = 0 Then
         If MSHFlexGrid10.Rows > 3 Then
            If CheckISINinRRNDetailTab = True Then
               If mCommodityID_ <> 31 And mCommodityID_ <> 114 Then
                  MsgBox "This lot belongs to one CMSE Lot No but the selected RRN belongs to a different CMSE Lot No!!!"
                  CmbLot.SetFocus
                  Exit Sub
               End If
            End If
         End If
      End If
   End If
End If
If CmbGCTDClient.Visible And (CmbGCTDClient.Text) = "" Then
   MsgBox "Blank " & LblClient_WHR_RRN.Caption & " not allowed!!! "
   CmbGCTDClient.SetFocus
   Exit Sub
End If

If CmbGSLTruckNo.Text = "" Then
   MsgBox "Blank Truck No. not allowed!!! "
   CmbGSLTruckNo.SetFocus
   Exit Sub
End If


If CheckDuplicateGSL() = True Then
   'MsgBox "This lot is already selected!!! "
   MsgBox "This lot, client & Truck No. is already selected!!! "
   CmbGodownNo.SetFocus
   Exit Sub
End If

If Check2.Value = 1 Then
   If Val(TxtGCTDNoofBags.Text) < 0 Then
      MsgBox "Number of Bags can't be Negative !!! "
      TxtGCTDNoofBags.SetFocus
      Exit Sub
   End If
Else
   If Val(TxtGCTDNoofBags.Text) <= 0 Then
      MsgBox "Number of Bags should be greater than zero!!! "
      TxtGCTDNoofBags.SetFocus
      Exit Sub
   End If
End If

If Check4.Value = 0 Then
   If Val(TxtGCTDGrossWt) <= 0 Then
      MsgBox "Weight should be greater than zero!!! "
      TxtGCTDGrossWt.SetFocus
      Exit Sub
   End If
End If

If Check4.Value = 1 Then
   If Val(TxtGCTDGrossWt) < 0 Then
      MsgBox "Negetive Weight not allowed!!! "
      TxtGCTDGrossWt.SetFocus
      Exit Sub
   End If
End If


'If mExchangeTypeID = 1 Then
'   If Val(TxtWHRGenBags) < Val(TxtGCTDNoofBags) Then
'      MsgBox "Number of Bags is greater than avalibale for withdrawal !!! "
'      TxtGCTDNoofBags.SetFocus
'      Exit Sub
'   End If
'   If Val(TxtWHRGenQty) < Val(TxtGCTDGrossWt) Then
'      MsgBox "Weight is greater than avalibale for withdrawal !!! "
'      TxtGCTDGrossWt.SetFocus
'      Exit Sub
'   End If
'End If

'Condition introduced for withdrawal of Chilli and Chilli 25 mail from Mayur Sir on 09 July 2013
If mCommodityID_ <> 31 And mCommodityID_ <> 114 Then
   If Check2.Value = 0 And Check4.Value = 0 Then
      tmp = Validate_ComBags(mCommodityID, Val(TxtGCTDNoofBags), Val(TxtGCTDGrossWt))
      If Len(tmp) > 1 Then
         MsgBox tmp
         TxtGCTDGrossWt.SetFocus
         Exit Sub
      End If
   End If
End If

If MSHFlexGrid2.Rows <= 3 Then
   If mExchangeTypeID = 1 Then 'NON NCDEX
      MsgBox "Missing WHR Detail!!! "
   ElseIf mExchangeTypeID = 2 Then 'NCDEX
      MsgBox "Missing RRN Detail!!! "
   ElseIf mExchangeTypeID = 3 Then 'SPOT
      MsgBox "Missing CDC Detail!!! "
   End If
   Exit Sub
End If

'Commented on 06 Aug 2013 after mail from Sandeep Mohite and Mayur Sir for check of Client Check for selected lot and WHR Client
'If mExchangeTypeID = 1 And Check2.Value = 0 Then And Check4.Value = 0 Then
If mExchangeTypeID = 1 And Check2.Value = 0 Then 'And Check4.Value = 0 Then
   If CheckClientinRRNTab() = True Then
      MsgBox "This lot belongs to one client but the selected WHR belongs to a different client!!!"
      CmbLot.SetFocus
      Exit Sub
   End If

   If Val(TxtGSLPledgedNofBags) > 0 Or Val(TxtGSLPledgedWeight) > 0 Then 'IIf(IsNull(RsMst_GSL!PledgeBags), "", RsMst_GSL!PledgeBags)
      MsgBox "This lot is pledged. " 'Can't withdraw!!!"
    '      CmbLot.SetFocus
    '      Exit Sub
   End If

   If Val(TxtGCTDNoofBags) > (Val(TxtGSLAvailableNofBags) - Val(TxtGSLPledgedNofBags)) Then
      MsgBox "Delivering more than pledge bags is not allowed!!! " 'Can't withdraw!!!"
      TxtGCTDNoofBags.SetFocus
      Exit Sub
   End If
End If

If mExchangeTypeID = 3 And Check2.Value = 0 And Check4.Value = 0 Then
   If CheckISINinRRNTab() = True Then
      MsgBox "This lot belongs to one client but the selected CDC belongs to a different client!!!"
      If CmbLot.Enabled Then
         CmbLot.SetFocus
      End If
      Exit Sub
   End If
End If


If Val(TxtMoisturePer.Text) <= 0 Then
   MsgBox "No moisture value is present. Please provide same in TTD details!!!"
   If CmbGSLTruckNo.Enabled Then
      CmbGSLTruckNo.SetFocus
   End If
   Exit Sub
End If


If mExchangeTypeID = 1 Then
      Dim strSplitString() As String
      Dim mWHR() As String
   If mWHRNo <> "" Then

      strSplitString = Split(CmbGCTDClient, "~")
      mWHR = Split(strSplitString(0), "-")
      tmp = (Val(mWHR(0)) * 1000000) + Val(mWHR(1))
      If mWHRNo <> tmp Then
         MsgBox "WHR Connected to this Lot is different than selected WHR!!! " 'Can't withdraw!!!"
         CmbLot.SetFocus
         Exit Sub
      End If
  Else
      If Check2.Value = 0 And Check4.Value = 0 Then
         strSplitString = Split(CmbGCTDClient, "~")
         mWHR = Split(strSplitString(0), "-")
         tmp = (Val(mWHR(0)) * 1000000) + Val(mWHR(1))
         Call TableTxn_WHRDetail("Where WHRNo = " & tmp & " ")
         If RsTxn_WHRDetail.RecordCount > 0 Then
            If RsTxn_WHRDetail!LotWiseWHR = "1" Then
               MsgBox "No WHR is connected to this Lot and WHR selected is lot wise!!!"
               CmbLot.SetFocus
               Exit Sub
            End If
         End If
      End If
  End If
End If

If mExchangeTypeID = 2 Then 'NCDEX
   If Check1.Value = 0 Then
      If Check2.Value = 0 And Check4.Value = 0 Then
         If MSHFlexGrid10.Rows > 3 Then
            If CheckISINinRRNDetailTab = True Then
               If mCommodityID_ <> 31 And mCommodityID_ <> 114 Then
                  MsgBox "This lot belongs to one CMSE Lot No but the selected RRN belongs to a different CMSE Lot No!!!"
                  CmbLot.SetFocus
                  Exit Sub
               End If
            End If
         Else
            If CheckISINinRRNTab() = True Then
               'Condition introduced for withdrawal of Chilli and Chilli 25 mail from Mayur Sir on 09 July 2013
               If mCommodityID_ <> 31 And mCommodityID_ <> 114 Then
                  MsgBox "This lot belongs to one CMSE Lot No but the selected RRN belongs to a different CMSE Lot No!!!"
                  CmbLot.SetFocus
                  Exit Sub
               End If
            End If
         End If
     End If
   Else
     If CheckClientinRRNTab() = True And Check2.Value = 0 And Check4.Value = 0 Then
        MsgBox "This is not a rejected lot of the selected client!!!"
        CmbLot.SetFocus
        Exit Sub
     End If
     If CmbGCTDClient.Text = "" Then
        MsgBox "Blank Client-DP not allowed!!!"
        CmbGCTDClient.SetFocus
        Exit Sub
     End If
  End If
End If
   
If Val(TxtGCTDNoofBags) > (Val(TxtGSLAvailableNofBags) - Val(TxtGSLPledgedNofBags)) Then
   MsgBox "Delivered number of bags are more than available unpledged number of bags!!!"
End If

mBag = getRemainingTruckBag(CmbGSLTruckNo.Text)

'If getRemainingTruckBag(CmbGSLTruckNo.Text) < Val(TxtGCTDNoofBags.Text) Then
If Val(mBag) < Val(TxtGCTDNoofBags.Text) Then
   MsgBox "Delivered number of bags are more than available number of bags of selected truck!!!"
   TxtGCTDNoofBags.SetFocus
   Exit Sub
End If

mWt = Format(getRemainingTruckWt(CmbGSLTruckNo.Text), "#######0.000")

'If getRemainingTruckWt(CmbGSLTruckNo.Text) < Val(TxtGCTDGrossWt.Text) Then
If Val(mWt) < Val(TxtGCTDGrossWt.Text) Then
   MsgBox "Delivered gross weight is more than available gross weight of selected truck!!!"
   TxtGCTDGrossWt.SetFocus
   Exit Sub
End If

Dim OAdjBags, OAdjWt As Variant
OAdjBags = 0
OAdjWt = 0
If CmdGCTDAdd2List.Caption = "Update List" Then
   r = MSHFlexGrid4.RowSel
   OAdjBags = Val(MSHFlexGrid4.TextMatrix(r, 6))
   OAdjWt = Val(MSHFlexGrid4.TextMatrix(r, 10))
   
  If mExchangeTypeID = 2 And MSHFlexGrid10.Rows > 3 Then
     mRRNISIN = UpdateRRNDetailTab(MSHFlexGrid4.TextMatrix(r, 12), OAdjWt, OAdjBags, "U")
     For ir = 2 To MSHFlexGrid2.Rows - 2
         If LCase(MSHFlexGrid4.TextMatrix(r, 12)) = LCase(MSHFlexGrid2.TextMatrix(ir, 1)) & " ~ " & LCase(MSHFlexGrid2.TextMatrix(ir, 2)) Then
   '         If ((OAdjBags + Val(MSHFlexGrid2.TextMatrix(ir, 14))) - Val(MSHFlexGrid2.TextMatrix(ir, 15))) < Val(TxtGCTDNoofBags) Then
   '         MsgBox "Delivered number of bags are more than avalibale number of bags!!!"
   '         Exit Sub
   '         End If
   '         If ((OAdjWt + Val(MSHFlexGrid2.TextMatrix(ir, 8))) - Val(MSHFlexGrid2.TextMatrix(ir, 9))) < Val(TxtGSLNetWeight) Then
   '            MsgBox "Delivered Quantity is more than avalibale Quantity!!!"
   '            Exit Sub
   '         End If
            MSHFlexGrid2.TextMatrix(0, 15) = (Val(MSHFlexGrid2.TextMatrix(0, 15)) - OAdjBags)
            MSHFlexGrid2.TextMatrix(0, 9) = (Val(MSHFlexGrid2.TextMatrix(0, 9)) - OAdjWt)

            MSHFlexGrid2.TextMatrix(ir, 15) = (Val(MSHFlexGrid2.TextMatrix(ir, 15)) - OAdjBags)
            MSHFlexGrid2.TextMatrix(ir, 9) = (Val(MSHFlexGrid2.TextMatrix(ir, 9)) - OAdjWt)

            If Check2.Value = 1 Or Check4.Value = 1 Then
               MSHFlexGrid2.TextMatrix(0, 16) = (Val(MSHFlexGrid2.TextMatrix(0, 16)) - OAdjBags)
               MSHFlexGrid2.TextMatrix(0, 17) = (Val(MSHFlexGrid2.TextMatrix(0, 17)) - OAdjWt)

               MSHFlexGrid2.TextMatrix(ir, 16) = (Val(MSHFlexGrid2.TextMatrix(ir, 16)) - OAdjBags)
               MSHFlexGrid2.TextMatrix(ir, 17) = (Val(MSHFlexGrid2.TextMatrix(ir, 17)) - OAdjWt)
            End If

            Exit For
          End If
       Next
   Else
      For ir = 2 To MSHFlexGrid2.Rows - 2
          If MSHFlexGrid4.TextMatrix(r, 12) = MSHFlexGrid2.TextMatrix(ir, 1) & " ~ " & MSHFlexGrid2.TextMatrix(ir, 4) Then
   '         If ((OAdjBags + Val(MSHFlexGrid2.TextMatrix(ir, 14))) - Val(MSHFlexGrid2.TextMatrix(ir, 15))) < Val(TxtGCTDNoofBags) Then
   '         MsgBox "Delivered number of bags are more than avalibale number of bags!!!"
   '         Exit Sub
   '         End If
   '         If ((OAdjWt + Val(MSHFlexGrid2.TextMatrix(ir, 8))) - Val(MSHFlexGrid2.TextMatrix(ir, 9))) < Val(TxtGSLNetWeight) Then
   '            MsgBox "Delivered Quantity is more than avalibale Quantity!!!"
   '            Exit Sub
   '         End If
            MSHFlexGrid2.TextMatrix(0, 15) = (Val(MSHFlexGrid2.TextMatrix(0, 15)) - OAdjBags)
            MSHFlexGrid2.TextMatrix(0, 9) = (Val(MSHFlexGrid2.TextMatrix(0, 9)) - OAdjWt)
            
            MSHFlexGrid2.TextMatrix(ir, 15) = (Val(MSHFlexGrid2.TextMatrix(ir, 15)) - OAdjBags)
            MSHFlexGrid2.TextMatrix(ir, 9) = (Val(MSHFlexGrid2.TextMatrix(ir, 9)) - OAdjWt)
            
            If Check2.Value = 1 Or Check4.Value = 1 Then
               MSHFlexGrid2.TextMatrix(0, 16) = (Val(MSHFlexGrid2.TextMatrix(0, 16)) - OAdjBags)
               MSHFlexGrid2.TextMatrix(0, 17) = (Val(MSHFlexGrid2.TextMatrix(0, 17)) - OAdjWt)
            
               MSHFlexGrid2.TextMatrix(ir, 16) = (Val(MSHFlexGrid2.TextMatrix(ir, 16)) - OAdjBags)
               MSHFlexGrid2.TextMatrix(ir, 17) = (Val(MSHFlexGrid2.TextMatrix(ir, 17)) - OAdjWt)
            End If
            
            Exit For
          End If
       Next
   End If
End If
If mExchangeTypeID = 2 And MSHFlexGrid10.Rows > 3 Then
   mRRNISIN = UpdateRRNDetailTab(CmbGCTDClient.Text, TxtGSLNetWeight, TxtGCTDNoofBags, "A")
   For ir = 2 To MSHFlexGrid2.Rows - 2
       If LCase(CmbGCTDClient.Text) = LCase(MSHFlexGrid2.TextMatrix(ir, 1)) & " ~ " & LCase(MSHFlexGrid2.TextMatrix(ir, 2)) Then
          ' If mExchangeTypeID = 2 And Check2.Value = 0 Then changed on 29.March.2010 Access allowed in case of NCDEX Rejection

          'Commented on 22 Nov 2012 Meeting with Sankar Sir Mayur Sir
          'If mExchangeTypeID = 2 And Check2.Value = 0 And Check1.Value = 0 And Check4.Value = 0 Then

          If mExchangeTypeID = 2 And Check1.Value = 0 Then
             If (Round((Val(MSHFlexGrid2.TextMatrix(ir, 8))) - Val(MSHFlexGrid2.TextMatrix(ir, 9)), 3)) < Round(Val(TxtGSLNetWeight), 3) Then
                MsgBox "Delivered Quantity is more than avalibale Quantity!!!"
                Exit Sub
             End If
          End If
          MSHFlexGrid2.TextMatrix(0, 15) = (Val(MSHFlexGrid2.TextMatrix(0, 15)) + Val(TxtGCTDNoofBags))
          MSHFlexGrid2.TextMatrix(0, 9) = (Val(MSHFlexGrid2.TextMatrix(0, 9)) + Val(TxtGSLNetWeight))

          MSHFlexGrid2.TextMatrix(ir, 15) = (Val(MSHFlexGrid2.TextMatrix(ir, 15)) + Val(TxtGCTDNoofBags))
          MSHFlexGrid2.TextMatrix(ir, 9) = (Val(MSHFlexGrid2.TextMatrix(ir, 9)) + Val(TxtGSLNetWeight))

          If Check2.Value = 1 Or Check4.Value = 1 Then
             MSHFlexGrid2.TextMatrix(0, 16) = (Val(MSHFlexGrid2.TextMatrix(0, 16)) + Val(TxtGCTDNoofBags))
             MSHFlexGrid2.TextMatrix(0, 17) = (Val(MSHFlexGrid2.TextMatrix(0, 17)) + Val(TxtGSLNetWeight))
          
             MSHFlexGrid2.TextMatrix(ir, 16) = (Val(MSHFlexGrid2.TextMatrix(ir, 16)) + Val(TxtGCTDNoofBags))
             MSHFlexGrid2.TextMatrix(ir, 17) = (Val(MSHFlexGrid2.TextMatrix(ir, 17)) + Val(TxtGSLNetWeight))
          End If
          Exit For
       End If
   Next
Else

   For ir = 2 To MSHFlexGrid2.Rows - 2
       If CmbGCTDClient.Text = MSHFlexGrid2.TextMatrix(ir, 1) & " ~ " & MSHFlexGrid2.TextMatrix(ir, 4) Then
          If mExchangeTypeID <> 2 And Check2.Value = 0 And Check4.Value = 0 Then ' Check1.Value =1 '30-01-10
          
             If ((Val(MSHFlexGrid2.TextMatrix(ir, 14))) - (Val(MSHFlexGrid2.TextMatrix(ir, 15)) - Val(MSHFlexGrid2.TextMatrix(ir, 16)))) < Val(TxtGCTDNoofBags) Then
                
    'Commented on 28 Sept 2011 for stopingof excess delivery from WHR
    '            MsgBox "Warring : Delivered number of bags are more than avalibale number of bags!!!"
    '            'Exit Sub
                
                MsgBox "Delivered number of bags are more than avalibale number of bags!!!"
                Exit Sub
             End If
          End If
          ' If mExchangeTypeID = 2 And Check2.Value = 0 Then changed on 29.March.2010 Access allowed in case of NCDEX Rejection
          
          'Commented on 22 Nov 2012 Meeting with Sankar Sir Mayur Sir
          'If mExchangeTypeID = 2 And Check2.Value = 0 And Check1.Value = 0 And Check4.Value = 0 Then
          
          If mExchangeTypeID = 2 And Check1.Value = 0 Then
             If (Round((Val(MSHFlexGrid2.TextMatrix(ir, 8))) - Val(MSHFlexGrid2.TextMatrix(ir, 9)), 3)) < Round(Val(TxtGSLNetWeight), 3) Then
                MsgBox "Delivered Quantity is more than avalibale Quantity!!!"
                Exit Sub
             End If
          ElseIf mExchangeTypeID = 1 Then
             tmp = tmp
          ElseIf mExchangeTypeID <> 2 And Check2.Value = 0 And Check4.Value = 0 Then  'Check1.Value = 0
             If (Round((Val(MSHFlexGrid2.TextMatrix(ir, 8))) - Val(MSHFlexGrid2.TextMatrix(ir, 9)), 3)) < Round(Val(TxtGSLNetWeight), 3) Then
                MsgBox "Delivered Quantity is more than avalibale Quantity!!!"
                Exit Sub
             End If
          End If
          
          MSHFlexGrid2.TextMatrix(0, 15) = (Val(MSHFlexGrid2.TextMatrix(0, 15)) + Val(TxtGCTDNoofBags))
          MSHFlexGrid2.TextMatrix(0, 9) = (Val(MSHFlexGrid2.TextMatrix(0, 9)) + Val(TxtGSLNetWeight))
          
          MSHFlexGrid2.TextMatrix(ir, 15) = (Val(MSHFlexGrid2.TextMatrix(ir, 15)) + Val(TxtGCTDNoofBags))
          MSHFlexGrid2.TextMatrix(ir, 9) = (Val(MSHFlexGrid2.TextMatrix(ir, 9)) + Val(TxtGSLNetWeight))
          
          If Check2.Value = 1 Or Check4.Value = 1 Then
             MSHFlexGrid2.TextMatrix(0, 16) = (Val(MSHFlexGrid2.TextMatrix(0, 16)) + Val(TxtGCTDNoofBags))
             MSHFlexGrid2.TextMatrix(0, 17) = (Val(MSHFlexGrid2.TextMatrix(0, 17)) + Val(TxtGSLNetWeight))
          
             MSHFlexGrid2.TextMatrix(ir, 16) = (Val(MSHFlexGrid2.TextMatrix(ir, 16)) + Val(TxtGCTDNoofBags))
             MSHFlexGrid2.TextMatrix(ir, 17) = (Val(MSHFlexGrid2.TextMatrix(ir, 17)) + Val(TxtGSLNetWeight))
          End If
          
          Exit For
       End If
   Next
End If

If CmdGCTDAdd2List.Caption = "Update List" Then
   r = MSHFlexGrid4.RowSel
   MSHFlexGrid4.TextMatrix(0, 4) = Val(MSHFlexGrid4.TextMatrix(0, 4)) - Val(MSHFlexGrid4.TextMatrix(r, 4))
   MSHFlexGrid4.TextMatrix(0, 5) = Val(MSHFlexGrid4.TextMatrix(0, 5)) - Val(MSHFlexGrid4.TextMatrix(r, 5))
   MSHFlexGrid4.TextMatrix(0, 6) = Val(MSHFlexGrid4.TextMatrix(0, 6)) - Val(MSHFlexGrid4.TextMatrix(r, 6))
   MSHFlexGrid4.TextMatrix(0, 7) = Val(MSHFlexGrid4.TextMatrix(0, 7)) - Val(MSHFlexGrid4.TextMatrix(r, 7))
   MSHFlexGrid4.TextMatrix(0, 8) = Val(MSHFlexGrid4.TextMatrix(0, 8)) - Val(MSHFlexGrid4.TextMatrix(r, 8))
   MSHFlexGrid4.TextMatrix(0, 9) = Val(MSHFlexGrid4.TextMatrix(0, 9)) - Val(MSHFlexGrid4.TextMatrix(r, 9))
   MSHFlexGrid4.TextMatrix(0, 10) = Val(MSHFlexGrid4.TextMatrix(0, 10)) - Val(MSHFlexGrid4.TextMatrix(r, 10))
   MSHFlexGrid4.TextMatrix(0, 11) = Val(MSHFlexGrid4.TextMatrix(0, 11)) - Val(MSHFlexGrid4.TextMatrix(r, 11))
   
'   If MSHFlexGrid4.TextMatrix(r, 17) = 1 Then
'      MSHFlexGrid4.TextMatrix(0, 17) = Val(MSHFlexGrid4.TextMatrix(0, 17)) - Val(MSHFlexGrid4.TextMatrix(r, 6))
'   End If
Else
   r = MSHFlexGrid4.Rows - 1
   MSHFlexGrid4.Rows = MSHFlexGrid4.Rows + 1
End If

MSHFlexGrid4.TextMatrix(r, 1) = CmbGodownNo.Text
MSHFlexGrid4.TextMatrix(r, 2) = CmbStack.Text
MSHFlexGrid4.TextMatrix(r, 3) = CmbLot.Text
MSHFlexGrid4.TextMatrix(r, 4) = Val(TxtGSLAvailableNofBags)
MSHFlexGrid4.TextMatrix(r, 5) = Val(TxtGSLAvailableGrossWgt)
MSHFlexGrid4.TextMatrix(r, 6) = Val(TxtGCTDNoofBags)
MSHFlexGrid4.TextMatrix(r, 7) = Val(TxtGCTDGrossWt)
MSHFlexGrid4.TextMatrix(r, 8) = Val(TxtAdjWt)
MSHFlexGrid4.TextMatrix(r, 9) = Val(TxtGGunnyWeight)
MSHFlexGrid4.TextMatrix(r, 10) = Val(TxtGSLNetWeight)
MSHFlexGrid4.TextMatrix(r, 11) = Val(TxtGSLWeightGainLoss)
'If Check2.Value = 1 Or Check1.Value = 1 Then
'   MSHFlexGrid4.TextMatrix(r, 12) = CmbGCTDClient.Text
'Else
'   MSHFlexGrid4.TextMatrix(r, 12) = TxtGSLISINno
'End If
MSHFlexGrid4.TextMatrix(r, 19) = CmbGSLTruckNo.Text


MSHFlexGrid4.Col = 7
MSHFlexGrid4.row = r
MSHFlexGrid4.CellBackColor = &HFFC0C0

If CmbGCTDClient.Visible Then
   MSHFlexGrid4.TextMatrix(r, 12) = CmbGCTDClient.Text
Else
   MSHFlexGrid4.TextMatrix(r, 12) = TxtGSLISINno
End If

MSHFlexGrid4.TextMatrix(r, 13) = TxtGCTDremarks
MSHFlexGrid4.TextMatrix(r, 14) = mGSLID
MSHFlexGrid4.TextMatrix(r, 16) = IIf(IsNull(mGSLClientIDRow), "", mGSLClientIDRow)
MSHFlexGrid4.TextMatrix(r, 17) = Check2.Value
MSHFlexGrid4.TextMatrix(r, 20) = Check4.Value
MSHFlexGrid4.TextMatrix(r, 21) = Val(TxtMoisturePer.Text)
MSHFlexGrid4.TextMatrix(r, 22) = Val(mGodownYearlyFlag)
MSHFlexGrid4.TextMatrix(r, 23) = Val(TxtAverageBagSize)

MSHFlexGrid4.TextMatrix(0, 4) = Val(MSHFlexGrid4.TextMatrix(0, 4)) + Val(MSHFlexGrid4.TextMatrix(r, 4))
MSHFlexGrid4.TextMatrix(0, 5) = Val(MSHFlexGrid4.TextMatrix(0, 5)) + Val(MSHFlexGrid4.TextMatrix(r, 5))
MSHFlexGrid4.TextMatrix(0, 6) = Val(MSHFlexGrid4.TextMatrix(0, 6)) + Val(MSHFlexGrid4.TextMatrix(r, 6))
MSHFlexGrid4.TextMatrix(0, 7) = Val(MSHFlexGrid4.TextMatrix(0, 7)) + Val(MSHFlexGrid4.TextMatrix(r, 7))
MSHFlexGrid4.TextMatrix(0, 8) = Val(MSHFlexGrid4.TextMatrix(0, 8)) + Val(MSHFlexGrid4.TextMatrix(r, 8))
MSHFlexGrid4.TextMatrix(0, 9) = Val(MSHFlexGrid4.TextMatrix(0, 9)) + Val(MSHFlexGrid4.TextMatrix(r, 9))
MSHFlexGrid4.TextMatrix(0, 10) = Val(MSHFlexGrid4.TextMatrix(0, 10)) + Val(MSHFlexGrid4.TextMatrix(r, 10))
MSHFlexGrid4.TextMatrix(0, 11) = Val(MSHFlexGrid4.TextMatrix(0, 11)) + Val(MSHFlexGrid4.TextMatrix(r, 11))

'If IIf(IsNull(RsMst_GSL!ContractID), 0, RsMst_GSL!ContractID) > 0 Then
If IIf(IsNull(mGSLContractID), 0, mGSLContractID) > 0 Then
   For m = 0 To MSHFlexGrid4.Cols - 1
      MSHFlexGrid4.row = r
      MSHFlexGrid4.Col = m
      MSHFlexGrid4.CellBackColor = vbYellow
   Next
End If

Call Clear_GCTDTabControl
If CmbGodownNo.Enabled = True Then
   CmbGodownNo.SetFocus
End If

End Sub
Function CheckISINinRRNTab() As Boolean
Dim Retval As Boolean
Dim TxtGSLISINno_ As Variant
Retval = True
If Check2.Value = 0 Then
   TxtGSLISINno_ = CmbGCTDClient.Text ' TxtGSLISINno
Else
   TxtGSLISINno_ = CmbGCTDClient.Text
End If

If MSHFlexGrid2.Rows > 3 Then
   For u = 2 To MSHFlexGrid2.Rows - 1
       If MSHFlexGrid2.TextMatrix(u, 1) <> "" Then
          If LCase(TxtGSLISINno_) = LCase(MSHFlexGrid2.TextMatrix(u, 1)) & " ~ " & LCase(MSHFlexGrid2.TextMatrix(u, 4)) Then
             If mExchangeTypeID = 2 Then
                If LCase(TxtGSLISINno) = LCase(MSHFlexGrid2.TextMatrix(u, 2)) Then
                   'mRRNinGCTD = MSHFlexGrid2.TextMatrix(u, 1)
                   Retval = False
                End If
                Exit For
             Else
                If LCase(TxtGSLISINno) = LCase(MSHFlexGrid2.TextMatrix(u, 1)) Then
                   'mRRNinGCTD = MSHFlexGrid2.TextMatrix(u, 1)
                   Retval = False
                End If
                Exit For
             End If
           Else
             Retval = True
           End If
       End If
   Next
End If

CheckISINinRRNTab = Retval

End Function
Function CheckClientinRRNTab() As Boolean
Dim Retval As Boolean
Dim TxtGSLISINno_ As Variant
Retval = True
If MSHFlexGrid2.Rows > 3 Then
   For u = 2 To MSHFlexGrid2.Rows - 1
       If CmbGCTDClient.Text = MSHFlexGrid2.TextMatrix(u, 1) & " ~ " & MSHFlexGrid2.TextMatrix(u, 4) Then
          If txtmGSLClientIDRow = MSHFlexGrid2.TextMatrix(u, 11) Then
             Retval = False
             Exit For
          End If
       End If
   Next
End If
CheckClientinRRNTab = Retval
End Function
Function CheckDuplicateGSL() As Boolean
Dim Retval As Boolean
Retval = False
If MSHFlexGrid4.Rows > 3 Then
   For u = 1 To MSHFlexGrid4.Rows - 1
       If MSHFlexGrid4.TextMatrix(u, 1) = CmbGodownNo.Text And _
          MSHFlexGrid4.TextMatrix(u, 2) = CmbStack.Text And _
          MSHFlexGrid4.TextMatrix(u, 3) = CmbLot.Text And _
          MSHFlexGrid4.TextMatrix(u, 12) = CmbGCTDClient.Text And MSHFlexGrid4.TextMatrix(u, 19) = CmbGSLTruckNo.Text Then
          If LCase(CmdGCTDAdd2List.Caption) = "add to list" Then
             Retval = True
             Exit For
          Else
             If u <> MSHFlexGrid4.RowSel Then
                Retval = True
             End If
          End If
       End If
   Next
End If

CheckDuplicateGSL = Retval

End Function

Private Sub ExitCmd_Click()
If SaveCmd.Enabled = True Then
   ans = MsgBox("Do you really want to exit?", vbYesNo)
   If ans = vbNo Then
      Exit Sub
   End If
End If
Unload Me
End Sub
Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Call GButtonLockAD(Me)
   MsgBox "Access denied!!!"
   Exit Sub
End If
SSTab1.Tab = 0
  
wc = ""
If Gen_WcCMP <> "" Then
   wc = "And " & Gen_WcCMP
End If
Call Grid_Head
CmbSCmp.Text = ""
CmbSCommodity.Text = ""
CmbSExchangeType.Text = ""
CmbSAgent.Text = ""
SFrom.Value = Date - 730
STo.Value = Date
TxtSCxTFNo = ""
TxtSManualCxTFNo = ""
CmbSRejection.Text = ""
Frame5.Visible = False
Call TableTxn_CxTF_Details("Where TxnType = 'W' " & wc & " Order By SM_Prefix,CMPID,CxTF_DetailsID  ")
If RsTxn_CxTF_Details.RecordCount = 0 Then
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


'Call AddCmd_Click
End Sub
Private Sub CmdTTDAdd2List_Click()
Dim r As Variant


If CmbExchangeTypeID.Text = "" Then
   MsgBox "Select a Mode!!!"
   CmbExchangeTypeID.SetFocus
   Exit Sub
End If

If CmbCMPID.Text = "" Then
   MsgBox "Select a CMP!!!"
   CmbCMPID.SetFocus
   Exit Sub
End If
If CmbCommodityID.Text = "" Then
   MsgBox "Select a Commodity!!!"
   CmbCommodityID.SetFocus
   Exit Sub
End If

If CmbVehicleTypeID.Text = "" Then
   MsgBox "Blank Vehicle Model not allowed!!!"
   CmbVehicleTypeID.SetFocus
   Exit Sub
End If

If TxtTTDTruckNo = "" Then
   MsgBox "Blank Truck No!!! "
   TxtTTDTruckNo.SetFocus
   Exit Sub
End If

If TxtWeighBridgNo = "" Then
   MsgBox "Blank Weigh Bridge No!!! "
   TxtWeighBridgNo.SetFocus
   Exit Sub
End If

If Val(TxtTTDTotalWt) <= 0 Then
   MsgBox "Invalid Weight!!! "
   TxtTTDTotalWt.SetFocus
   Exit Sub
End If

If Val(TxtTTDTareWt) < 0 Then
   MsgBox "Invalid Weight!!! "
   TxtTTDTareWt.SetFocus
   Exit Sub
End If

If Val(TxtTTDGrossWt) <= 0 Then
   MsgBox "Invalid Weight!!! "
   TxtTTDGrossWt.SetFocus
   Exit Sub
End If

If Val(TxtTTDNoOfBags) <= 0 Then
   MsgBox "Invalid No of Bags !!! "
   TxtTTDNoOfBags.SetFocus
   Exit Sub
End If


' Removal of Gross weight check
'tmp = Format(Val(TxtTTDTotalWt) - Val(TxtTTDTareWt), "#######0.000")
'r = Format(Val(TxtTTDGrossWt), "#######0.000")
'
'tmp = (Val(TxtTTDTotalWt) - Val(TxtTTDTareWt))
'r = (Val(TxtTTDGrossWt))
'
'If (tmp) <> r Then
'   MsgBox "Invalid Weight!!! "
'   Exit Sub
'End If

If Val(TxtTTDTotalWt) <= 0 Then
   MsgBox "Invalid Weight !!! "
   TxtTTDTotalWt.SetFocus
   Exit Sub
End If

If Val(TxtTTDGrossWt) <= 0 Then
   MsgBox "Invalid Weight !!! "
   TxtTTDGrossWt.SetFocus
   Exit Sub
End If
If CheckDuplicateTruck() = True Then
   MsgBox "Duplicate Truck No. and Weighbridge Slip No. not Allowed!!! "
   TxtTTDTruckNo.SetFocus
   Exit Sub
End If

If Val(TxtMoisturePer) = 0 Then
   MsgBox "Blank  Moisture % Not Allowed !!! "
   TxtMoisturePer.SetFocus
   Exit Sub
End If

If Val(TxtMoisturePer) < 0 Then
   MsgBox "Invalid Moisture % !!! "
   TxtMoisturePer.SetFocus
   Exit Sub
End If

If Val(TxtTTDTareWt) < Val(TxtMinimum) Or Val(TxtTTDTareWt) > Val(TxtMaximum) Then
   MsgBox "The Tare weight of the Truck appears to be beyond the limits set for this Vehicle Model. Please make an enquiry and ensure that there is no fraudulent activity taking place before going ahead with this data entry!!" '" Tare weight must be between " & Val(TxtMinimum) & " and " & Val(TxtMaximum) & "  range of Vehicle Model Master !!!", vbExclamation, "Message"
   'TxtTTDTareWt.SetFocus
   'Exit Sub
End If

tmp = Format((Val(TxtTTDTotalWt) - Val(TxtTTDTareWt)) - Val(TxtTTDGrossWt), "#####0.000")

If Val(tmp) <> 0 Then
   MsgBox "Invalid Weight !!! "
   Exit Sub
End If

If CmdTTDAdd2List.Caption = "Update List" Then

   If MsgBox("This will clear GSL details of selected Truck, You wan't to continue?", vbYesNo) = vbNo Then
      Call Clear_TTDTabControl
      Exit Sub
   End If
   r = MSHFlexGrid3.RowSel
   MSHFlexGrid3.TextMatrix(0, 3) = Val(MSHFlexGrid3.TextMatrix(0, 3)) - Val(MSHFlexGrid3.TextMatrix(r, 3))
   MSHFlexGrid3.TextMatrix(0, 4) = Val(MSHFlexGrid3.TextMatrix(0, 4)) - Val(MSHFlexGrid3.TextMatrix(r, 4))
   MSHFlexGrid3.TextMatrix(0, 5) = Val(MSHFlexGrid3.TextMatrix(0, 5)) - Val(MSHFlexGrid3.TextMatrix(r, 5))
   MSHFlexGrid3.TextMatrix(0, 6) = Val(MSHFlexGrid3.TextMatrix(0, 6)) - Val(MSHFlexGrid3.TextMatrix(r, 6))
   
   Call DeleteFromGSLDetail
   Call Clear_GCTDTabControl
Else
   r = MSHFlexGrid3.Rows - 1
   MSHFlexGrid3.Rows = MSHFlexGrid3.Rows + 1
End If

MSHFlexGrid3.TextMatrix(r, 1) = TxtTTDTruckNo
MSHFlexGrid3.TextMatrix(r, 2) = TxtWeighBridgNo
MSHFlexGrid3.TextMatrix(r, 3) = TxtTTDTotalWt
MSHFlexGrid3.TextMatrix(r, 4) = TxtTTDTareWt
MSHFlexGrid3.TextMatrix(r, 5) = TxtTTDGrossWt
MSHFlexGrid3.TextMatrix(r, 6) = Val(TxtTTDNoOfBags)
MSHFlexGrid3.TextMatrix(r, 7) = txtTTDRemarks
MSHFlexGrid3.TextMatrix(r, 8) = TxtMoisturePer.Text

MSHFlexGrid3.TextMatrix(r, 9) = CmbVehicleTypeID.Text
MSHFlexGrid3.TextMatrix(r, 10) = mVehicleTypeID
MSHFlexGrid3.TextMatrix(r, 11) = Val(TxtMinimum)
MSHFlexGrid3.TextMatrix(r, 12) = Val(TxtMaximum)

MSHFlexGrid3.Col = 5
MSHFlexGrid3.row = r
MSHFlexGrid3.CellBackColor = &HFFC0C0


MSHFlexGrid3.TextMatrix(0, 3) = Val(MSHFlexGrid3.TextMatrix(0, 3)) + Val(MSHFlexGrid3.TextMatrix(r, 3))
MSHFlexGrid3.TextMatrix(0, 4) = Val(MSHFlexGrid3.TextMatrix(0, 4)) + Val(MSHFlexGrid3.TextMatrix(r, 4))
MSHFlexGrid3.TextMatrix(0, 5) = Val(MSHFlexGrid3.TextMatrix(0, 5)) + Val(MSHFlexGrid3.TextMatrix(r, 5))
MSHFlexGrid3.TextMatrix(0, 6) = Val(MSHFlexGrid3.TextMatrix(0, 6)) + Val(MSHFlexGrid3.TextMatrix(r, 6))

Call Clear_TTDTabControl

'TxtTTDTruckNo.SetFocus
CmbVehicleTypeID.SetFocus
End Sub
Sub Clear_TTDTabControl()
CmbVehicleTypeID.Text = ""
TxtMinimum = ""
TxtMaximum = ""
mVehicleTypeID = ""
TxtTTDTruckNo = ""
TxtWeighBridgNo = ""
TxtTTDTotalWt = ""
TxtTTDTareWt = ""
TxtTTDGrossWt = ""
txtTTDRemarks = ""
TxtTTDNoOfBags = ""
CmdTTDAdd2List.Caption = "Add to List"
TxtTTDGrossWt.Locked = True
End Sub

Private Sub MSHFlexGrid2_Click()
If SaveCmd.Enabled = False Then
   Exit Sub
End If

CmdAdd2ListRRN.Caption = "Add to List"
TxtRRNNo.Locked = False

r = MSHFlexGrid2.RowSel
If r <= 0 Then Exit Sub

'Changed by Parag , dt 30;09-2011 whr edit mode was not working
'If MSHFlexGrid2.TextMatrix(r, 2) = "" Then
'   Exit Sub
'End If

If MSHFlexGrid2.TextMatrix(r, 1) = "" Then
   Exit Sub
End If


CmdAdd2ListRRN.Caption = "Update List"
TxtRRNNo.Locked = False

If Val(MSHFlexGrid2.TextMatrix(r, 9)) <> 0 Then
   MsgBox "You can't modify this row since corresponding GSL details are already selected!!!"
   Call Clear_RWCTabControl
   Exit Sub
End If


TxtRRNNo = MSHFlexGrid2.TextMatrix(r, 1)
TxtISINno = MSHFlexGrid2.TextMatrix(r, 2)
TxtClientID = MSHFlexGrid2.TextMatrix(r, 3)
TxtClientName = MSHFlexGrid2.TextMatrix(r, 4)
TxtDPID = MSHFlexGrid2.TextMatrix(r, 5)
TxtDPName = MSHFlexGrid2.TextMatrix(r, 6)
TxtRRNQty = MSHFlexGrid2.TextMatrix(r, 7)
TxtBalanceQty = MSHFlexGrid2.TextMatrix(r, 8)
TxtRRNRemarks = MSHFlexGrid2.TextMatrix(r, 10)
mClientIDRow = Val(MSHFlexGrid2.TextMatrix(r, 11))
mNAVCode = MSHFlexGrid2.TextMatrix(r, 18)
mRNAVCode = MSHFlexGrid2.TextMatrix(r, 19)
mRWCID = MSHFlexGrid2.TextMatrix(r, 12)
TxtBalanceBags = MSHFlexGrid2.TextMatrix(r, 14)

End Sub

Private Sub MSHFlexGrid3_Click()
If SaveCmd.Enabled = False Then
   Exit Sub
End If
CmdTTDAdd2List.Caption = "Update List"
r = MSHFlexGrid3.RowSel
If r <= 1 Then Exit Sub
If MSHFlexGrid3.TextMatrix(r, 1) = "" Then
   CmdTTDAdd2List.Caption = "Add to List"
   Exit Sub
End If
TxtTTDTruckNo = MSHFlexGrid3.TextMatrix(r, 1)
TxtWeighBridgNo = MSHFlexGrid3.TextMatrix(r, 2)
TxtTTDTotalWt = MSHFlexGrid3.TextMatrix(r, 3)
TxtTTDTareWt = MSHFlexGrid3.TextMatrix(r, 4)
TxtTTDGrossWt = MSHFlexGrid3.TextMatrix(r, 5)
TxtTTDNoOfBags = MSHFlexGrid3.TextMatrix(r, 6)
txtTTDRemarks = MSHFlexGrid3.TextMatrix(r, 7)
TxtMoisturePer = MSHFlexGrid3.TextMatrix(r, 8)
CmbVehicleTypeID.Text = MSHFlexGrid3.TextMatrix(r, 9)
mVehicleTypeID = Val(MSHFlexGrid3.TextMatrix(r, 10))
TxtMinimum = Val(MSHFlexGrid3.TextMatrix(r, 11))
TxtMaximum = Val(MSHFlexGrid3.TextMatrix(r, 12))

End Sub
Private Sub MSHFlexGrid4_Click()

If SaveCmd.Enabled = False Then
   Exit Sub
End If

Call Clear_GCTDTabControl
r = MSHFlexGrid4.RowSel
If r <= 1 Then Exit Sub
If MSHFlexGrid4.TextMatrix(r, 15) <> "" Then
   MsgBox "Since this godown is closed, edit is not allowed!!!"
   Exit Sub
End If

If MSHFlexGrid4.TextMatrix(r, 1) = "" Then
   Call Clear_GCTDTabControl
   Exit Sub
End If

'CmdGCTDAdd2List.Caption = "Update List"  Error Found while  edit
'CmdGCTDAdd2List.Enabled = True

CmbGodownNo.Text = MSHFlexGrid4.TextMatrix(r, 1)
CmbGSLTruckNo.Text = MSHFlexGrid4.TextMatrix(r, 19)
Call CmbGSLTruckNo_LostFocus
TxtGSLAvailableNofBags = MSHFlexGrid4.TextMatrix(r, 4)
TxtGSLAvailableGrossWgt = MSHFlexGrid4.TextMatrix(r, 5)
TxtGCTDNoofBags = MSHFlexGrid4.TextMatrix(r, 6)
TxtGCTDGrossWt = MSHFlexGrid4.TextMatrix(r, 7)
TxtAdjWt = MSHFlexGrid4.TextMatrix(r, 8)
TxtGGunnyWeight = MSHFlexGrid4.TextMatrix(r, 9)
TxtGSLNetWeight = MSHFlexGrid4.TextMatrix(r, 10)
TxtGSLWeightGainLoss = MSHFlexGrid4.TextMatrix(r, 11)
TxtGSLISINno = MSHFlexGrid4.TextMatrix(r, 12)
TxtGCTDremarks = MSHFlexGrid4.TextMatrix(r, 13)
mGSLID = MSHFlexGrid4.TextMatrix(r, 14)
'mGodownID = Val(MSHFlexGrid4.TextMatrix(r, 15))
'mRWCID = MSHFlexGrid4.TextMatrix(r, 16)
Check2.Value = Val(MSHFlexGrid4.TextMatrix(r, 17))
Check4.Value = Val(MSHFlexGrid4.TextMatrix(r, 20))
mGodownYearlyFlag = Val(MSHFlexGrid4.TextMatrix(r, 22))
CmdGCTDAdd2List.Caption = "Update List"
CmdGCTDAdd2List.Enabled = True

If CmbGCTDClient.Visible = True Then
   CmbGCTDClient.Text = MSHFlexGrid4.TextMatrix(r, 12)
End If
Call CmbGodownNo_GotFocus
Call CmbGodownNo_LostFocus
CmbStack.Text = MSHFlexGrid4.TextMatrix(r, 2)
Call CmbStack_GotFocus
Call CmbStack_LostFocus
CmbLot.Text = MSHFlexGrid4.TextMatrix(r, 3)
Call CmbLot_GotFocus
Call CmbLot_LostFocus

If MSHFlexGrid4.TextMatrix(r, 0) <> "" Or mExchangeTypeID = 3 Then
   CmbGodownNo.Enabled = False
   CmbStack.Enabled = False
   CmbLot.Enabled = False
Else
   CmbGodownNo.Enabled = True
   CmbStack.Enabled = True
   CmbLot.Enabled = True
End If
If mExchangeTypeID = 3 Then
   Check2.Enabled = False
   'Call Check2_Click
End If
CmbLot.Text = MSHFlexGrid4.TextMatrix(r, 3)
CmbGSLTruckNo.Text = MSHFlexGrid4.TextMatrix(r, 19)
TxtAverageBagSize = MSHFlexGrid4.TextMatrix(r, 23)
CmdGCTDAdd2List.Enabled = True



End Sub

Private Sub SaveCmd_Click()
Dim ans As Variant

'/Call UpdateRRNGridStock

If mExchangeTypeID = 1 Then
   Call UpdateErrorGrid
End If

If UpdateGSLBalChkGrid = False Then
   Frame1.Visible = True
   Frame1.Height = 9450
   Frame1.Width = 14775
   Frame1.Top = 120
   Frame1.Left = 120
   MSHFlexGrid7.Height = 7275
   MSHFlexGrid7.Width = 14535
   MSHFlexGrid7.Top = 360
   MSHFlexGrid7.Left = 135
   Picture3.Enabled = False
   Exit Sub
Else
   Frame1.Visible = False
   Picture3.Enabled = True
End If

'If TxtCxTFDate.Value > Date Then
'   MsgBox "Future Date is not Allowed !!!! "
'   TxtCxTFDate.SetFocus
'   Exit Sub
'End If

If CmbCMPID.Text <> "" Then
   If TxtCxTFDate.Value < CDate(TxtCxTFDispalyDate) Then
      MsgBox "Date less than BOD Date is not Allowed !!!! "
      TxtCxTFDate.SetFocus
      Exit Sub
   End If
End If

If CmbExchangeTypeID.Text = "" Then
   MsgBox "Select a Mode!!!"
   CmbExchangeTypeID.SetFocus
   Exit Sub
End If

If CmbCMPID.Text = "" Then
   MsgBox "Select a CMP!!!"
   CmbCMPID.SetFocus
   Exit Sub
End If

If CmbAgent.Text = "" Then
   MsgBox "Select an Agent!!!"
   CmbAgent.SetFocus
   Exit Sub
End If

If CmbCommodityID.Text = "" Then
   MsgBox "Select a Commodity!!!"
   CmbCommodityID.SetFocus
   Exit Sub
End If

If MSHFlexGrid2.Rows <= 3 Then
   If mExchangeTypeID = 1 Then 'NON NCDEX
      MsgBox "Missing WHR detail!!! "
   ElseIf mExchangeTypeID = 2 Then 'NCDEX
      MsgBox "Missing RRN detail!!! "
   ElseIf mExchangeTypeID = 3 Then 'SPOT
      MsgBox "Missing CDC detail!!! "
   End If
   Exit Sub
End If
If MSHFlexGrid3.Rows <= 3 Then
   MsgBox "Missing Truck wise detail!!! "
   Exit Sub
End If
If MSHFlexGrid4.Rows <= 3 Then
   MsgBox "Missing GSL Detail!!! "
   Exit Sub
End If

If Val(MSHFlexGrid3.TextMatrix(0, 5)) <> Val(MSHFlexGrid4.TextMatrix(0, 7)) Then
   MsgBox "Gross Weight in 'Truck Detail' tab is different from Gross Weight in 'GSL Detail' tab!!! "
   Exit Sub
End If

If Val(MSHFlexGrid3.TextMatrix(0, 6)) <> Val(MSHFlexGrid4.TextMatrix(0, 6)) Then
''   If MsgBox("No of Bags in 'Truck Detail' tab is different from No of Bags in 'GSL Detail' tab, do you wan't to continue?", vbYesNo) = vbNo Then
''      Exit Sub
''   End If
   MsgBox "No of Bags in 'Truck Detail' tab is different from No of Bags in 'GSL Detail' tab!!!"
   Exit Sub
End If


'If mExchangeTypeID = 1 Then
'   If Val(MSHFlexGrid2.TextMatrix(0, 15)) <> Val(MSHFlexGrid4.TextMatrix(0, 6)) Then
'      MsgBox "'Delivered Bags' in 'WHR No' tab is different from delivered bags in 'GSL Detail' tab!!!! "
'      Exit Sub
'   End If
'End If

If mExchangeTypeID = 2 Then
   If Check1.Value = 0 Then
      If TxtSpotPrice = "" Then
         MsgBox "Blank Spot Price not allowed!!!"
         TxtSpotPrice.SetFocus
         Exit Sub
      End If
   End If
   SSTab1.Tab = 3
   If Val(MSHFlexGrid5.TextMatrix(0, 5)) > 0 Then
      tmp = "Excess deliveries are made against certain CMSE Lot Nos which are highlighted in yellow." & _
             "Please make appropriate adjustment entries against those CMSE Lot Nos in the 'GSL Details' tab." & _
             "Please note that these adjustment entries should be Negative  entries."
      MsgBox tmp
      SSTab1.Tab = 3
      Exit Sub
   End If
   If Val(MSHFlexGrid5.TextMatrix(0, 4)) < 0 And Val(MSHFlexGrid5.TextMatrix(0, 6)) > 0 Then
      If Abs(Val(MSHFlexGrid5.TextMatrix(0, 4))) < Val(MSHFlexGrid5.TextMatrix(0, 6)) Then
         
         For I = 2 To MSHFlexGrid5.Rows - 1
             If Val(MSHFlexGrid5.TextMatrix(I, 6)) > 0 Then
                MSHFlexGrid5.row = I
                MSHFlexGrid5.Col = 1
                MSHFlexGrid5.CellBackColor = vbRed
             End If
         Next
         
         
'         tmp = "You have made a net adjustment of " & Abs(Val(MSHFlexGrid5.TextMatrix(0, 4))) & " MT towards Excess " & _
'               "Delivery. Additionally you have also short delivered to the tune of " & Val(MSHFlexGrid5.TextMatrix(0, 6)) & " MT. " & _
'               "You need to adjust this Excess Delivery Adjustment by a corresponding Short Delivery Adjustment " & _
'               "Hence please make appropriate Adjustment Entry of " & Abs(Val(MSHFlexGrid5.TextMatrix(0, 4))) & "  MT " & _
'               "in those ISINs where short delivery are made."
'
         tmp = "Please make Adjustment Entry of " & Abs(Val(MSHFlexGrid5.TextMatrix(0, 4))) & "  MT " & _
               "in 'GSL Details' tab for some of those CMSE Lot Nos which are highlighted in red."
         MsgBox tmp
'         For i = 2 To MSHFlexGrid5.Rows - 1
'             If Val(MSHFlexGrid5.TextMatrix(i, 6)) > 0 Then
'                MSHFlexGrid5.Row = i
'                MSHFlexGrid5.Col = 6
'                MSHFlexGrid5.CellBackColor = vbRed
'                MSHFlexGrid5.Col = 1
'                MSHFlexGrid5.CellBackColor = vbRed
'             End If
'         Next
         'ans = MsgBox(tmp, vbOK)
                 'MsgBox "Create entries  for adjusted weight for  " & Abs(Val(MSHFlexGrid5.TextMatrix(0, 4))) & " MT  'GSL Detail' tab !!!!"
         
      Else
'         tmp = "You have made a net adjustment of " & Abs(Val(MSHFlexGrid5.TextMatrix(0, 4))) & " MT towards Excess " & _
'               "Delivery. Additionally you have also short delivered to the tune of " & Val(MSHFlexGrid5.TextMatrix(0, 6)) & " MT. " & _
'               "You need to adjust this Excess Delivery Adjustment by a corresponding Short Delivery Adjustment " & _
'               "Hence please make appropriate Adjustment Entry of " & Abs(Val(MSHFlexGrid5.TextMatrix(0, 6))) & "  MT " & _
'               "in those ISINs where short delivery are made."
         
         tmp = "Please make Adjustment Entry of " & Abs(Val(MSHFlexGrid5.TextMatrix(0, 6))) & "  MT " & _
               "in 'GSL Details' tab for some of those CMSE Lot Nos which are highlighted in red."
         MsgBox tmp
         
         'MsgBox "Create entries  for adjusted weight for  " & Abs(Val(MSHFlexGrid5.TextMatrix(0, 6))) & " MT  'GSL Detail' tab !!!!"
      End If
      Exit Sub
   End If
''   If CheckWtRRNDetailTab = False Then
''      MsgBox " Withdrawal Quantity should not greater than Remat quantity in RRN!!"
''      Exit Sub
''   End If
End If



If Val(MSHFlexGrid2.TextMatrix(0, 9)) <> Val(MSHFlexGrid4.TextMatrix(0, 10)) Then
   If mExchangeTypeID = 1 Then
      ''Delivered Qty' in 'WHR No' tab is different from 'Electronic Weight' in 'GSL Detail' tab!!!! "
      MsgBox "Warning: You are delivering more than WHR Balance Quantity!!! "
   ElseIf mExchangeTypeID = 2 Then
     If Check1.Value = 0 Then
        'MsgBox "'Delivered Qty' in 'RRN' tab is different from 'Electronic Weight' in 'GSL Detail' tab!!!! "
        MsgBox "You are delivering more than RRN Balance Quantity!!!"
     ElseIf Check1.Value = 1 Then
        'MsgBox "Delivered Quantity in 'Client Detail tab' is not equal to Electronic Weight in 'Delivery Lot(G~S~L) Details' "
        MsgBox "You are delivering more than the Balance Quantity in 'Client Detail tab'!!!"
     End If
     Exit Sub
   ElseIf mExchangeTypeID = 3 Then
        'MsgBox "'Delivered Qty' in 'CDC No' tab is different from 'Electronic Weight' in 'GSL Detail' tab!!!! "
        MsgBox "You are delivering more than CDC Balance Quantity!!!"
        Exit Sub
   End If
End If


For x = 2 To MSHFlexGrid2.Rows - 1
    If mExchangeTypeID <> 1 Then
       If Val(MSHFlexGrid2.TextMatrix(x, 8)) < Val(MSHFlexGrid2.TextMatrix(x, 9)) Then
          MsgBox "Warning: You are delivering more than Rejected Balance Quantity!!! "
          Exit For
       End If
    End If
    If mExchangeTypeID <> 2 Then
       If Val(MSHFlexGrid2.TextMatrix(x, 14)) < Val(MSHFlexGrid2.TextMatrix(x, 15)) Then
          MsgBox "Warning: You are delivering more than WHR Balance Bags!!! "
          Exit For
       End If
    End If
Next



'MsgBox ("Since no Edit or Delete option available for CITF, Data once saved can't be modified therefor go through all the tabs carfully and ensure correctness of data entered by you !!! ",vbYesNo
tmp = MsgBox("Data once saved can't be modified. Are you sure you want to save?", vbYesNo)
If tmp = vbNo Then
   Exit Sub
End If


If Edit_Flg = "A" Then
   
   Call TableTxn_CxTF_Details(" Where TxnType = 'W' And CxTF_DetailsID = 0 ")
   
   RsTxn_CxTF_Details.AddNew
   RsTxn_CxTF_Details!CxTFNO = GetCxTFNo
   RsTxn_CxTF_Details!G_UID = GetGUID
   TxtGUID = RsTxn_CxTF_Details!G_UID
   TxtCxTFNo.Text = RsTxn_CxTF_Details!CxTFNO
Else
   RsTxn_CxTF_Details.MoveFirst
   RsTxn_CxTF_Details.Find " CxTFNo= " & TxtCxTFNo.Text
    'Call DeleteChildTableDetail
End If

RsTxn_CxTF_Details!CxTFNO = TxtCxTFNo.Text
RsTxn_CxTF_Details!ManualCDTFNo = TxtManualCDTFNo
RsTxn_CxTF_Details!CxTFDate = TxtCxTFDate.Value
RsTxn_CxTF_Details!CMPID = mCMPID
RsTxn_CxTF_Details!CommodityID = mCommodityID
RsTxn_CxTF_Details!ExchangeTypeID = mExchangeTypeID
RsTxn_CxTF_Details!BookedSpaceFlag = Check1.Value
'RsTxn_CxTF_Details!Flag = TxtFlag.Text
RsTxn_CxTF_Details!Agent = TxtAgent.Text
RsTxn_CxTF_Details!AgentID = mAgentID

'RsTxn_CxTF_Details!Remark = TxtRemark.Text
RsTxn_CxTF_Details!SM_Prefix = TxtPreFix
RsTxn_CxTF_Details!SCMFlag = ChkSCMFlag.Value
RsTxn_CxTF_Details!TxnType = "W"
If mContractID <> "" Then
   RsTxn_CxTF_Details!ContractID = mContractID
End If

RsTxn_CxTF_Details!SpotPrice = Val(TxtSpotPrice)

If IsNull(RsTxn_CxTF_Details!CreatedBy) = True Or RsTxn_CxTF_Details!CreatedBy = "" Then
   RsTxn_CxTF_Details!CreatedBy = Gen_UserLoginID
   RsTxn_CxTF_Details!CreatedDate = Now
Else
   RsTxn_CxTF_Details!UpdatedBy = Gen_UserLoginID
   RsTxn_CxTF_Details!UpdatedDate = Now
End If

RsTxn_CxTF_Details.Update

tmp = "Select CxTF_DetailsID From Txn_CxTF_Details Where G_UID='" & TxtGUID & "'"
Call TmpTable

If TmpRs.RecordCount > 0 Then
'   mCxTF_DetailsID = TmpRs!CxTF_DetailsID
   TxtCxTFNo.Text = TmpRs!CxTF_DetailsID
End If

Call SaveClient
Call SaveTruck
Call SaveGSL


SaveCmd.Enabled = False
Check1.Value = 0

'RsTxn_CxTF_Details.Requery

Call TableTxn_CxTF_Details(" Where TxnType = 'W'")

If RsTxn_CxTF_Details.RecordCount > 0 Then
   RsTxn_CxTF_Details.MoveFirst
End If

RsTxn_CxTF_Details.Find "G_UID='" & TxtGUID & "'"

If ChkTareWeightDiscripency = True Then
   Call Gen_MailTareWeightDiscripency
End If

Call Indata

If LCase(CmbExchangeTypeID.Text) = "non ncdex" Or (LCase(CmbExchangeTypeID.Text) = "ncdex" And Check1.Value = 1) Then
   Call CmdBOE_Click
End If

End Sub
Private Sub UpdateBillingDetails()

tmp = "Select CxTF_DetailsID,SM_Prefix,CxTFDate,CDTFDate,NoOfDay,ClientID,Mst_GSL_ID,GodownID,CommodityID,SpillageFlag,NoOfBags,DematWeight,BagSize,NContractID,BillingCycleID,WBillingCycleID,BagSizeFlag ,FinalPrice,WithdrawalFinalPrice,G_UID,DepConDID"
tmp = tmp & " From (Select TC.CxTF_DetailsID,TC.SM_Prefix ,TC.CxTFDate,GSL.CDTFDate,day(TC.CxTFDate) NoOfDay,GSL.ClientID,TCG.Mst_GSL_ID,GSL.GodownID,GSL.CommodityID ,GSL.SpillageFlag,TCG.NoOfBags,TCG.DematWeight,(TCG.DematWeight/TCG.NoOfBags)*1000 as 'BagSize',GSL.NContractID,TDC.BillingCycleID,TDC.WBillingCycleID,TDC.BagSizeFlag ,TDCD.FinalPrice,TDCD.WithdrawalFinalPrice,TCG.G_UID,TDCD.DepConDID"
tmp = tmp & " From  Txn_CxTF_GSL TCG"
tmp = tmp & " Inner Join Txn_CxTF_Details TC on TCG.SM_Prefix=TC.SM_Prefix And TCG.CxTF_DetailsID=TC.CxTF_DetailsID And TCG.TxnType=TC.TxnType"
tmp = tmp & " left Join Mst_GSL GSL on TCG.Mst_GSL_ID=GSL.Mst_GSL_ID And TCG.SM_Prefix=GSL.SM_Prefix"
tmp = tmp & " Inner Join Txn_DepositContractGodownDetail TDCD On GSL.DepConDID = TDCD.DepConDID "
tmp = tmp & " Inner Join Txn_DepositContract TDC On TDCD.DepConID = TDC.DepConID  And TDC.ApprovalFlag in ('S','A','C')  "
tmp = tmp & " Where TC.SM_Prefix = '" & TxtPreFix & "' And  TC.CxTF_DetailsID = " & Val(TxtCxTFNo) & " And  TC.CxTFDate = '" & TxtCxTFDate.Value & "'   And TC.ExchangeTypeID=1 And  TC.TxnType='W'  And TCG.DepConDID is null And GSL.SpillageFlag = 0  and isnull(GSL.NContractID,0)=0 And TDC.BagSizeFlag = 'Y'"
tmp = tmp & " Union All"
tmp = tmp & " Select TC.CxTF_DetailsID,TC.SM_Prefix ,TC.CxTFDate,GSL.CDTFDate,day(TC.CxTFDate) NoOfDay,GSL.ClientID,TCG.Mst_GSL_ID,GSL.GodownID,GSL.CommodityID ,GSL.SpillageFlag,TCG.NoOfBags,TCG.DematWeight,(TCG.DematWeight/TCG.NoOfBags)*1000 as 'BagSize',GSL.NContractID,TDC.BillingCycleID,TDC.WBillingCycleID,TDC.BagSizeFlag,TDCD.FinalPrice,TDCD.WithdrawalFinalPrice,TCG.G_UID,TDCD.DepConDID"
tmp = tmp & " From  Txn_CxTF_GSL TCG"
tmp = tmp & " Inner Join Txn_CxTF_Details TC on TCG.SM_Prefix=TC.SM_Prefix And TCG.CxTF_DetailsID=TC.CxTF_DetailsID And TCG.TxnType=TC.TxnType"
tmp = tmp & " left Join Mst_GSL GSL on TCG.Mst_GSL_ID=GSL.Mst_GSL_ID And TCG.SM_Prefix=GSL.SM_Prefix"
tmp = tmp & " Inner Join Txn_DepositContractGodownDetail TDCD On GSL.DepConDID = TDCD.DepConDID "
tmp = tmp & " Inner Join Txn_DepositContract TDC On TDCD.DepConID = TDC.DepConID  And TDC.ApprovalFlag in ('S','A','C')  "
tmp = tmp & " Where TC.SM_Prefix = '" & TxtPreFix & "' And TC.CxTF_DetailsID = " & Val(TxtCxTFNo) & " And TC.CxTFDate = '" & TxtCxTFDate.Value & "'   And TC.ExchangeTypeID=1 And  TC.TxnType='W'  And TCG.DepConDID is null And GSL.SpillageFlag = 0  and isnull(GSL.NContractID,0)=0 And TDC.BagSizeFlag = 'N') A"
tmp = tmp & " Order By SM_Prefix,CxTF_DetailsID,Mst_GSL_ID"


Call TmpTable

For I = 1 To TmpRs.RecordCount

    mDepoBillingCycleID = ""
    mDepoBillingUnit = ""
    mNoofDays = ""
    mDepConDID = ""
    
    
    
    mDepoBillingCycleID = IIf(IsNull(TmpRs!WBillingCycleID), "", TmpRs!WBillingCycleID)
    If TmpRs!BagSizeFlag = "Y" Then
       mDepoBillingUnit = "B"
    Else
       mDepoBillingUnit = "Q"
    End If
    mDepoBillingRate = IIf(IsNull(TmpRs!WithdrawalFinalPrice), 0, TmpRs!WithdrawalFinalPrice)
    mDepConDID = IIf(IsNull(TmpRs!DepConDID), "", TmpRs!DepConDID)
    
    If mDepoBillingCycleID = 1 Then
       mNoofDays = Val(Day(TmpRs!CxTFDate))
    ElseIf mDepoBillingCycleID = 2 Then
       If Day(TmpRs!CxTFDate) < 8 Then
          mNoofDays = "1"
       ElseIf Day(TmpRs!CxTFDate) < 15 Then
          mNoofDays = "2"
       ElseIf Day(TmpRs!CxTFDate) < 22 Then
          mNoofDays = "3"
       Else
          mNoofDays = "4"
       End If
    ElseIf mDepoBillingCycleID = 3 Then
       If Day(TmpRs!CxTFDate) < 16 Then
          mNoofDays = "1"
       Else
          mNoofDays = "2"
       End If
    ElseIf mDepoBillingCycleID = 4 Then
       mNoofDays = "1"
    End If
    
    If mDepoBillingCycleID <> "" And mDepoBillingUnit <> "" And Val(mDepoBillingRate) <> 0 And mNoofDays <> "" Then
       tmp = "Update Txn_CxTF_GSL Set RateCardType = 'D',BillingCycleID = " & mDepoBillingCycleID & ",BillingUnit = '" & mDepoBillingUnit & "',Rate = " & mDepoBillingRate & ",RateForDailyPnL = " & mDepoBillingRate & ",DepConDID =  " & mDepConDID & ",NoofDays = " & mNoofDays & " Where SM_Prefix = '" & TmpRs!SM_Prefix & "' And CxTF_DetailsID = '" & TmpRs!CxTF_DetailsID & "' And TxnType='W' And DepConDID is null And G_UID = '" & TmpRs!G_UID & "' "
       Call Tmp1Table
    End If
    
    TmpRs.MoveNext
Next


End Sub


'Private Function GetCxTFNo() As Double
'Dim Retval As Double
'tmp = "Select max(CxTFNo) from Txn_CxTF_Details Where  sm_prefix = '" & TxtPreFix & "' And TxnType = 'W'  "
'Call TmpTable
'Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
'Retval = Retval + 1
'GetCxTFNo = Retval
'End Function

Private Function GetCxTFNo() As Double

Dim Retval As Double
Dim LogRetVal As Double

tmp = "Select max(CxTFNo) from Txn_CxTF_Details Where  sm_prefix = '" & TxtPreFix & "' And TxnType = 'W'  "
Call TmpTable

'Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
'Retval = Retval + 1

tmp = "Select max(CxTFNo) from Log_CxTF_Details Where TxnType = 'W' And SM_Prefix = '" & TxtPreFix & "'"
Call Tmp1Table

Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
LogRetVal = IIf(IsNull(TmpRs1.Fields(0)), 0, TmpRs1.Fields(0))

If Retval > LogRetVal Then
   Retval = Retval + 1
Else
   Retval = LogRetVal + 1
End If

GetCxTFNo = Retval

End Function


Private Sub SaveClient()

Call TableTxn_CxTF_ClientDetail(" Where  sm_prefix = '" & TxtPreFix & "' And CxTF_DetailsID = " & TxtCxTFNo & " And TxnType = 'W' ")
If mExchangeTypeID = 2 Then
   Call TableMst_RRN
End If

'old code
'If mExchangeTypeID = 1 Then
'   Call TableTxn_WHRDetail(" where  sm_prefix = '" & TxtPreFix & "' ")
'End If
'If mExchangeTypeID = 3 Then
'   Call TableTxn_CDC_GSLDetail(" where  sm_prefix = '" & TxtPreFix & "' ")
'End If

'New Code by prajakta 30 Jan 2015
If LCase(TxtPreFix) = "bknrp" Then
   mNewSM_Prefix = "BKNER"
ElseIf LCase(TxtPreFix) = "ujjap" Then
   mNewSM_Prefix = "Ujjan"
ElseIf LCase(TxtPreFix) = "dewsp" Then
   mNewSM_Prefix = "DEWAS"
ElseIf LCase(TxtPreFix) = "vdshp" Then
   mNewSM_Prefix = "VDISA"
ElseIf LCase(TxtPreFix) = "adlbp" Then ' 'adlbp
   mNewSM_Prefix = "adlbd"
ElseIf LCase(TxtPreFix) = "aklap" Then ' 'adlbp
   mNewSM_Prefix = "AKOLA"
ElseIf LCase(TxtPreFix) = "sngrp" Then ' 'adlbp
   mNewSM_Prefix = "sgngr"
ElseIf LCase(TxtPreFix) = "adlbd" Then ' 'adlbp
   mNewSM_Prefix = "adlbp"
ElseIf LCase(TxtPreFix) = "nzmdp" Then ' 'adlbp
   mNewSM_Prefix = "nzmbd"
ElseIf LCase(TxtPreFix) = "unjap" Then ' 'adlbp
   mNewSM_Prefix = "unjha"
ElseIf LCase(TxtPreFix) = "sglip" Then ' 'adlbp
   mNewSM_Prefix = "sngli"
ElseIf LCase(TxtPreFix) = "gntrp" Then ' 'adlbp
   mNewSM_Prefix = "guntr"
ElseIf LCase(TxtPreFix) = "guntr" Then ' 'adlbp
   mNewSM_Prefix = "gntrp"
ElseIf LCase(TxtPreFix) = "dvgrp" Then ' 'adlbp
   mNewSM_Prefix = "dvngr"
ElseIf LCase(TxtPreFix) = "jdhrp" Then ' 'adlbp
   mNewSM_Prefix = "jdhpr"
ElseIf LCase(TxtPreFix) = "hngtp" Then ' 'adlbp
   mNewSM_Prefix = "hghat"
ElseIf LCase(TxtPreFix) = "kadip" Then ' 'adlbp
   mNewSM_Prefix = "kadi"
ElseIf LCase(TxtPreFix) = "gbrgp" Then ' 'adlbp
   mNewSM_Prefix = "glbrg"
ElseIf LCase(TxtPreFix) = "glbrg" Then ' 'adlbp
   mNewSM_Prefix = "gbrgp"
ElseIf LCase(TxtPreFix) = "kngmp" Then ' 'adlbp
   mNewSM_Prefix = "kngym"
ElseIf LCase(TxtPreFix) = "lturp" Then ' 'adlbp
   mNewSM_Prefix = "latur"
ElseIf LCase(TxtPreFix) = "ktap" Then ' 'Kota
   mNewSM_Prefix = "kota"
ElseIf LCase(TxtPreFix) = "rsgrp" Then ' 'Raisingnagar
   mNewSM_Prefix = "rsngr"
ElseIf LCase(TxtPreFix) = "rsngr" Then ' 'Raisingnagar
   mNewSM_Prefix = "rsgrp"
ElseIf LCase(TxtPreFix) = "ttcrp" Then ' 'Kota
   mNewSM_Prefix = "tcrin"
Else
   mNewSM_Prefix = ""
End If

If mExchangeTypeID = 1 Then
   Call TableTxn_WHRDetail(" where  (sm_prefix = '" & TxtPreFix & "' or sm_prefix = '" & mNewSM_Prefix & "')")
End If

If mExchangeTypeID = 3 Then
   Call TableTxn_CDC_GSLDetail(" Where  (sm_prefix = '" & TxtPreFix & "' or sm_prefix = '" & mNewSM_Prefix & "')")
End If

For I = 2 To MSHFlexGrid2.Rows - 1
    If MSHFlexGrid2.TextMatrix(I, 1) <> "" Then
       RsTxn_CxTF_ClientDetail.AddNew
       RsTxn_CxTF_ClientDetail!ClientDetailID = getCxtf_ClientMax
       RsTxn_CxTF_ClientDetail!G_UID = GetGUID
       RsTxn_CxTF_ClientDetail!CxTF_DetailsID = TxtCxTFNo
       RsTxn_CxTF_ClientDetail!ClientID = MSHFlexGrid2.TextMatrix(I, 11)
       RsTxn_CxTF_ClientDetail!TxnType = "W"
       RsTxn_CxTF_ClientDetail!RRNID = MSHFlexGrid2.TextMatrix(I, 12)
       RsTxn_CxTF_ClientDetail!AdjRRNQty = Val(MSHFlexGrid2.TextMatrix(I, 9))
       RsTxn_CxTF_ClientDetail!AdjRRNNoofBags = Val(MSHFlexGrid2.TextMatrix(I, 15))
       
       RsTxn_CxTF_ClientDetail!SplRRNQty = Val(MSHFlexGrid2.TextMatrix(I, 17))
       RsTxn_CxTF_ClientDetail!SplRRNNoofBags = Val(MSHFlexGrid2.TextMatrix(I, 16))
       
       RsTxn_CxTF_ClientDetail!SM_Prefix = TxtPreFix
       RsTxn_CxTF_ClientDetail.Update
       RsTxn_CxTF_ClientDetail.Requery
       
       If mExchangeTypeID = 1 Then
          If Check1.Value = 0 Then
             RsTxn_WHRDetail.MoveFirst
             RsTxn_WHRDetail.Find "WHRNo = " & Val(MSHFlexGrid2.TextMatrix(I, 12))
             If Not RsTxn_WHRDetail.EOF Then
                RsTxn_WHRDetail!DeliveredQuantity = RsTxn_WHRDetail!DeliveredQuantity + Val(MSHFlexGrid2.TextMatrix(I, 9))
                
                'Commented on 01 Oct 2011 after discussion with Mayur Sir and Sankar Sir
                'RsTxn_WHRDetail!DeliveredNoOfBags = RsTxn_WHRDetail!DeliveredNoOfBags + Val(MSHFlexGrid2.TextMatrix(I, 15))
               
                RsTxn_WHRDetail!DeliveredNoOfBags = RsTxn_WHRDetail!DeliveredNoOfBags + (Val(MSHFlexGrid2.TextMatrix(I, 15)) - Val(MSHFlexGrid2.TextMatrix(I, 16)))
                
                RsTxn_WHRDetail.Update
                RsTxn_WHRDetail.Requery
             End If
          End If
       End If
       If mExchangeTypeID = 2 Then
          If Check1.Value = 0 Then
             RsMst_RRN.MoveFirst
             RsMst_RRN.Find "RRNID = " & Val(MSHFlexGrid2.TextMatrix(I, 12))
             If Not RsMst_RRN.EOF Then
                RsMst_RRN!DeliveredQty = RsMst_RRN!DeliveredQty + Val(MSHFlexGrid2.TextMatrix(I, 9))
                RsMst_RRN.Update
                RsMst_RRN.Requery
             End If
          End If
       End If
    
       If mExchangeTypeID = 3 Then
          If Check1.Value = 0 Then
             RsTxn_CDC_GSLDetail.MoveFirst
             RsTxn_CDC_GSLDetail.Find "CDCNo = " & Val(MSHFlexGrid2.TextMatrix(I, 12))
             If Not RsTxn_CDC_GSLDetail.EOF Then
                RsTxn_CDC_GSLDetail!DeliveredQty = RsTxn_CDC_GSLDetail!DeliveredQty + Val(MSHFlexGrid2.TextMatrix(I, 9))
                
                tmp = "Select DepositeWeight From Mst_GSL Where SM_Prefix = '" & TxtPreFix & "' And Mst_GSL_ID = " & RsTxn_CDC_GSLDetail!Mst_GSL_ID & " "
                Call TmpTable
                
                If TmpRs.RecordCount > 0 Then
                   If TmpRs!DepositeWeight = RsTxn_CDC_GSLDetail!DeliveredQty Then
                      RsTxn_CDC_GSLDetail!Flag = "W"
                   End If
                End If
                RsTxn_CDC_GSLDetail.Update
                RsTxn_CDC_GSLDetail.Requery
             End If
          End If
       End If
    
    End If
Next


End Sub
Function getCxtf_ClientMax() As Double
Dim Retval As Double
Retval = 0
tmp = "select max(ClientDetailID) from Txn_CxTF_ClientDetail where sm_prefix = '" & TxtPreFix & "'"
Call TmpTable
If TmpRs.RecordCount > 0 Then
   Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
End If
Retval = Retval + 1
TmpRs.Close
getCxtf_ClientMax = Retval

End Function
Private Sub SaveTruck()
Call TableTxn_CxTF_TruckDetail(" Where sm_prefix = '" & TxtPreFix & "' And CxTF_DetailsID = " & TxtCxTFNo & " And TxnType = 'W'")
For I = 2 To MSHFlexGrid3.Rows - 1
    If MSHFlexGrid3.TextMatrix(I, 1) <> "" Then
       RsTxn_CxTF_TruckDetail.AddNew
       RsTxn_CxTF_TruckDetail!TruckDetailID = getCxtf_TruckMax
       RsTxn_CxTF_TruckDetail!G_UID = GetGUID
       RsTxn_CxTF_TruckDetail!CxTF_DetailsID = TxtCxTFNo
       RsTxn_CxTF_TruckDetail!TruckNo = MSHFlexGrid3.TextMatrix(I, 1)
       RsTxn_CxTF_TruckDetail!WeighBridgeSlipno = MSHFlexGrid3.TextMatrix(I, 2)
       RsTxn_CxTF_TruckDetail!TotalWeight = MSHFlexGrid3.TextMatrix(I, 3)
       RsTxn_CxTF_TruckDetail!TareWeight = Val(MSHFlexGrid3.TextMatrix(I, 4))
       RsTxn_CxTF_TruckDetail!GrossWeight = MSHFlexGrid3.TextMatrix(I, 5)
       RsTxn_CxTF_TruckDetail!NoofBags = Val(MSHFlexGrid3.TextMatrix(I, 6))
       RsTxn_CxTF_TruckDetail!Remarks = MSHFlexGrid3.TextMatrix(I, 7)
       RsTxn_CxTF_TruckDetail!MoisturePer = MSHFlexGrid3.TextMatrix(I, 8)
       RsTxn_CxTF_TruckDetail!TxnType = "W"
       RsTxn_CxTF_TruckDetail!SM_Prefix = TxtPreFix
       RsTxn_CxTF_TruckDetail!VehicleTypeID = Val(MSHFlexGrid3.TextMatrix(I, 10))
       RsTxn_CxTF_TruckDetail.Update
       RsTxn_CxTF_TruckDetail.Requery
    End If
Next
End Sub
Function getCxtf_TruckMax() As Double
Dim Retval As Double
Retval = 0
tmp = "select max(TruckDetailID) from Txn_CxTF_TruckDetail where sm_prefix = '" & TxtPreFix & "'"
Call TmpTable
If TmpRs.RecordCount > 0 Then
   Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
End If
Retval = Retval + 1
TmpRs.Close
getCxtf_TruckMax = Retval

End Function
Private Sub SaveGSL()
Dim PQty, TQty As Variant

Call TableTxn_CxTF_GSL(" Where sm_prefix = '" & TxtPreFix & "' And CxTF_DetailsID = " & TxtCxTFNo & " And TxnType = 'W' ")
Call TableMst_GSL(" where sm_prefix = '" & TxtPreFix & "'")
For I = 2 To MSHFlexGrid4.Rows - 1
    If MSHFlexGrid4.TextMatrix(I, 1) <> "" Then
       RsTxn_CxTF_GSL.AddNew
       RsTxn_CxTF_GSL!GSLDetailID = getCxtf_GSLMax
       RsTxn_CxTF_GSL!G_UID = GetGUID
       RsMst_GSL.MoveFirst
       RsMst_GSL.Find "Mst_GSL_ID = " & Val(MSHFlexGrid4.TextMatrix(I, 14))
       
       RsTxn_CxTF_GSL!CxTF_DetailsID = TxtCxTFNo
       RsTxn_CxTF_GSL!NoofBags = Val(MSHFlexGrid4.TextMatrix(I, 6))
       RsTxn_CxTF_GSL!GrossWeight = Val(MSHFlexGrid4.TextMatrix(I, 7))
       RsTxn_CxTF_GSL!AdjNCMSLWeight = Val(MSHFlexGrid4.TextMatrix(I, 8))
       'RsTxn_CxTF_GSL!AdjNClientWeight = Val(MSHFlexGrid4.TextMatrix(i, 8))
       RsTxn_CxTF_GSL!GunnyWeight = Val(MSHFlexGrid4.TextMatrix(I, 9))
       RsTxn_CxTF_GSL!DematWeight = Val(MSHFlexGrid4.TextMatrix(I, 10))
       RsTxn_CxTF_GSL!WeightGainLoss = Val(MSHFlexGrid4.TextMatrix(I, 11)) 'Weight Gain Or Loss
      
       
       If mExchangeTypeID = 1 Then
          RsTxn_CxTF_GSL!ClientID = MSHFlexGrid4.TextMatrix(I, 16)
          'mWHRSplit(0) 'MSHFlexGrid4.TextMatrix(I, 16)
       End If
       
       'Introduced on 01 Oct 2011 after discussion with Mayur Sir and Sankar Sir
       Dim strSplitString() As String
       strSplitString = Split(MSHFlexGrid4.TextMatrix(I, 12), "~")
       If mExchangeTypeID = 1 Then
          Dim mWHR() As String
          mWHR = Split(strSplitString(0), "-")
          tmp = (Val(mWHR(0)) * 1000000) + Val(mWHR(1))
       ElseIf mExchangeTypeID = 2 Then
          tmp = GetRRNID(strSplitString(0))
       ElseIf mExchangeTypeID = 3 Then
          tmp = Trim(strSplitString(0))
       End If
       If Check1.Value = 0 Then
          RsTxn_CxTF_GSL!WHRNo = tmp
       End If
       RsTxn_CxTF_GSL!Remarks = MSHFlexGrid4.TextMatrix(I, 13)
       RsTxn_CxTF_GSL!Mst_GSL_ID = MSHFlexGrid4.TextMatrix(I, 14)
       RsTxn_CxTF_GSL!TxnType = "W"
       RsTxn_CxTF_GSL!SM_Prefix = TxtPreFix
       strSplitString = Split(MSHFlexGrid4.TextMatrix(I, 19), "~")
       RsTxn_CxTF_GSL!TruckNo = strSplitString(0)
       RsTxn_CxTF_GSL!WeighBridgeSlipno = strSplitString(1)
       RsTxn_CxTF_GSL!MoisturePer = Val(MSHFlexGrid4.TextMatrix(I, 21))
       RsTxn_CxTF_GSL.Update
       RsTxn_CxTF_GSL.Requery
       If Not RsMst_GSL.EOF Then
       
'          If mExchangeTypeID = 2 Then
             mCalMoisturePer = CheckMoisturePer(I)
             RsMst_GSL!CITF_MoisturePer = Format(Val(mCalMoisturePer), "#######0.000")
'          End If
       
          RsMst_GSL!CDTFBags = IIf(IsNull(RsMst_GSL!CDTFBags), 0, RsMst_GSL!CDTFBags) + Val(MSHFlexGrid4.TextMatrix(I, 6))
          RsMst_GSL!BalanceBags = IIf(IsNull(RsMst_GSL!BalanceBags), 0, RsMst_GSL!BalanceBags) - Val(MSHFlexGrid4.TextMatrix(I, 6))
'          If MSHFlexGrid4.TextMatrix(I, 20) = 0 Then
             RsMst_GSL!CDTFWeight = IIf(IsNull(RsMst_GSL!CDTFWeight), 0, RsMst_GSL!CDTFWeight) + Val(MSHFlexGrid4.TextMatrix(I, 10))
             RsMst_GSL!BalanceWeight = IIf(IsNull(RsMst_GSL!BalanceWeight), 0, RsMst_GSL!BalanceWeight) - Val(MSHFlexGrid4.TextMatrix(I, 10))
             RsMst_GSL!WithdrawalAdjWt = IIf(IsNull(RsMst_GSL!WithdrawalAdjWt), 0, RsMst_GSL!WithdrawalAdjWt) + Val(MSHFlexGrid4.TextMatrix(I, 8))
'          End If
          RsMst_GSL!Flag1 = RsMst_GSL!Status
          If Val(RsMst_GSL!BalanceBags) = 0 Then
             RsMst_GSL!Flag1 = RsMst_GSL!Status
             RsMst_GSL!Status = "E"
          End If
          RsMst_GSL.Update
       End If

    End If
Next

End Sub
Private Sub ChkDeleteGSL()

Call TableMst_GSL(" Where sm_prefix = '" & TxtPreFix & "' And BalanceBags > 0 and Status = 'E' And SpillageFlag = 0 And MadeUpFlag = 0 ")

'tmp = "Select * from mst_gsl where BalanceBags > 0 and flag = 'E'"
'Call Tmp1Table

For I = 2 To MSHFlexGrid4.Rows - 1
    If RsMst_GSL.RecordCount > 0 Then
       RsMst_GSL.MoveFirst
    End If
    If MSHFlexGrid4.TextMatrix(I, 14) = "" Then Exit For
    RsMst_GSL.Find "Mst_GSL_ID = " & MSHFlexGrid4.TextMatrix(I, 14) & ""
    If RsMst_GSL.EOF = False Then
       If mExchangeTypeID = 2 Then
          If Check1.Value = 1 Then
             RsMst_GSL!State = "J"
             RsMst_GSL!Status = "O"
          ElseIf RsMst_GSL!ISINID <> "" And RsMst_GSL!FirstISINID <> "" Then
             RsMst_GSL!State = "R"
             RsMst_GSL!Status = "O"
          Else
             RsMst_GSL!State = "A"
             RsMst_GSL!Status = "O"
          End If
       ElseIf mExchangeTypeID = 1 Then
          RsMst_GSL!Status = "O"
       End If
       RsMst_GSL.Update
    End If
'    RsMst_GSL.Requery
Next


End Sub

Private Sub InDataTruck()
Call TableTxn_CxTF_TruckDetail("  Where sm_prefix = '" & TxtPreFix & "' And CxTF_DetailsID = " & TxtCxTFNo & " And TxnType = 'W'")
MSHFlexGrid3.Rows = MSHFlexGrid3.Rows + RsTxn_CxTF_TruckDetail.RecordCount
For I = 2 To MSHFlexGrid3.Rows - 2
    MSHFlexGrid3.TextMatrix(I, 0) = RsTxn_CxTF_TruckDetail!TruckDetailID
    MSHFlexGrid3.TextMatrix(I, 1) = RsTxn_CxTF_TruckDetail!TruckNo
    MSHFlexGrid3.TextMatrix(I, 2) = RsTxn_CxTF_TruckDetail!WeighBridgeSlipno
    MSHFlexGrid3.TextMatrix(I, 3) = RsTxn_CxTF_TruckDetail!TotalWeight
    MSHFlexGrid3.TextMatrix(I, 4) = RsTxn_CxTF_TruckDetail!TareWeight
    MSHFlexGrid3.TextMatrix(I, 5) = RsTxn_CxTF_TruckDetail!GrossWeight
    MSHFlexGrid3.TextMatrix(I, 6) = IIf(IsNull(RsTxn_CxTF_TruckDetail!NoofBags), 0, RsTxn_CxTF_TruckDetail!NoofBags)
    MSHFlexGrid3.TextMatrix(I, 7) = IIf(IsNull(RsTxn_CxTF_TruckDetail!Remarks), "", RsTxn_CxTF_TruckDetail!Remarks)
    MSHFlexGrid3.TextMatrix(I, 8) = IIf(IsNull(RsTxn_CxTF_TruckDetail!MoisturePer), "", RsTxn_CxTF_TruckDetail!MoisturePer)
    MSHFlexGrid3.TextMatrix(I, 9) = IIf(IsNull(RsTxn_CxTF_TruckDetail!VehicleTypeID), "", GetVehicleType(RsTxn_CxTF_TruckDetail!VehicleTypeID, "N"))
    MSHFlexGrid3.TextMatrix(I, 10) = IIf(IsNull(RsTxn_CxTF_TruckDetail!VehicleTypeID), "", RsTxn_CxTF_TruckDetail!VehicleTypeID)
    MSHFlexGrid3.TextMatrix(I, 11) = IIf(IsNull(RsTxn_CxTF_TruckDetail!VehicleTypeID), "", GetVehicleType(RsTxn_CxTF_TruckDetail!VehicleTypeID, "Min"))
    MSHFlexGrid3.TextMatrix(I, 12) = IIf(IsNull(RsTxn_CxTF_TruckDetail!VehicleTypeID), "", GetVehicleType(RsTxn_CxTF_TruckDetail!VehicleTypeID, "Max"))
    
    MSHFlexGrid3.Col = 5
    MSHFlexGrid3.row = I
    MSHFlexGrid3.CellBackColor = &HFFC0C0
    
    MSHFlexGrid3.TextMatrix(0, 3) = Val(MSHFlexGrid3.TextMatrix(0, 3)) + RsTxn_CxTF_TruckDetail!TotalWeight
    MSHFlexGrid3.TextMatrix(0, 4) = Val(MSHFlexGrid3.TextMatrix(0, 4)) + RsTxn_CxTF_TruckDetail!TareWeight
    MSHFlexGrid3.TextMatrix(0, 5) = Val(MSHFlexGrid3.TextMatrix(0, 5)) + RsTxn_CxTF_TruckDetail!GrossWeight
    MSHFlexGrid3.TextMatrix(0, 6) = Val(MSHFlexGrid3.TextMatrix(0, 6)) + IIf(IsNull(RsTxn_CxTF_TruckDetail!NoofBags), 0, RsTxn_CxTF_TruckDetail!NoofBags)
    
    RsTxn_CxTF_TruckDetail.MoveNext
Next
End Sub
Function getCxtf_GSLMax() As Double
Dim Retval As Double
Retval = 0
tmp = "select max(GSLDetailID) from Txn_CxTF_GSL where sm_prefix = '" & TxtPreFix & "'"
Call TmpTable
If TmpRs.RecordCount > 0 Then
   Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
End If
Retval = Retval + 1
TmpRs.Close
getCxtf_GSLMax = Retval

End Function
Private Sub InDataClient()
Dim ans As Variant

If Check1.Value <> 1 Then
   tmp = "Select rrnid,count(rrnid) from Txn_CxTF_ClientDetail " & _
         " Where TxnType = 'W' And CxTF_DetailsID = " & TxtCxTFNo & " And sm_prefix = '" & TxtPreFix & "' " & _
         " Group by rrnid order by count(rrnid)"
   Call Tmp4Table

   If TmpRs4.RecordCount > 0 Then
      TmpRs4.MoveLast
      If TmpRs4.Fields(1) > 1 Then
         BalBagNeg = "Y"
      End If
   End If
   TmpRs4.Close
End If
Call TableTxn_CxTF_ClientDetail(" Where sm_prefix = '" & TxtPreFix & "' And CxTF_DetailsID = " & TxtCxTFNo & " And TxnType = 'W' ")
If mExchangeTypeID = 1 Then
   Call TableTxn_WHRDetail(" where sm_prefix = '" & TxtPreFix & "'")
End If
If mExchangeTypeID = 2 Then
   Call TableMst_RRN
End If
If mExchangeTypeID = 3 Then
   Call TableTxn_CDC_GSLDetail(" where sm_prefix = '" & TxtPreFix & "'")
   Call TableMst_GSL(" where sm_prefix = '" & TxtPreFix & "'")
End If

MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + RsTxn_CxTF_ClientDetail.RecordCount
For I = 2 To MSHFlexGrid2.Rows - 2
    MSHFlexGrid2.TextMatrix(I, 0) = RsTxn_CxTF_ClientDetail!ClientDetailID
    Call GetRWCDetails(I, mExchangeTypeID, RsTxn_CxTF_ClientDetail!RRNID)
    ans = GetClientNameByRow(RsTxn_CxTF_ClientDetail!ClientID) 'ans = GetClientNameByRow(RsTxn_CxTF_ClientDetail!ClientIDRow)
    MSHFlexGrid2.TextMatrix(I, 3) = ans(0)
    MSHFlexGrid2.TextMatrix(I, 4) = ans(1) 'TmpRs!ClientName
    MSHFlexGrid2.TextMatrix(I, 5) = ans(2) 'TmpRs!DepositoryID
    MSHFlexGrid2.TextMatrix(I, 6) = ans(3) 'TmpRs!DepositoryName
    If mExchangeTypeID = 1 Then
       RsTxn_WHRDetail.Find "WHRNo = " & RsTxn_CxTF_ClientDetail!RRNID
       MSHFlexGrid2.TextMatrix(I, 7) = Val(RsTxn_WHRDetail!WHRQuantity)
       MSHFlexGrid2.TextMatrix(I, 8) = Format(Val(RsTxn_WHRDetail!WHRQuantity) - Val(RsTxn_WHRDetail!DeliveredQuantity), "#######0.000")
       MSHFlexGrid2.TextMatrix(I, 9) = RsTxn_CxTF_ClientDetail!AdjRRNQty
       
       MSHFlexGrid2.TextMatrix(I, 13) = Val(RsTxn_WHRDetail!WHRNoOfBags)
       MSHFlexGrid2.TextMatrix(I, 14) = Format(Val(RsTxn_WHRDetail!WHRNoOfBags) - Val(RsTxn_WHRDetail!DeliveredNoOfBags), "#######0.000")
       MSHFlexGrid2.TextMatrix(I, 15) = RsTxn_CxTF_ClientDetail!AdjRRNNoofBags
    
       MSHFlexGrid2.TextMatrix(I, 17) = IIf(IsNull(RsTxn_CxTF_ClientDetail!SplRRNQty), 0, RsTxn_CxTF_ClientDetail!SplRRNQty)
       MSHFlexGrid2.TextMatrix(I, 16) = IIf(IsNull(RsTxn_CxTF_ClientDetail!SplRRNNoofBags), 0, RsTxn_CxTF_ClientDetail!SplRRNNoofBags)
    ElseIf mExchangeTypeID = 2 Then
       If Check1.Value = 0 Then
          RsMst_RRN.Find "RRNID = " & RsTxn_CxTF_ClientDetail!RRNID
          MSHFlexGrid2.TextMatrix(I, 7) = RsMst_RRN!RRNQty
          MSHFlexGrid2.TextMatrix(I, 8) = Format(RsMst_RRN!RRNQty - RsMst_RRN!DeliveredQty, "#######0.000")
          MSHFlexGrid2.TextMatrix(I, 9) = RsTxn_CxTF_ClientDetail!AdjRRNQty
       ElseIf Check1.Value = 1 Then
          MSHFlexGrid2.TextMatrix(I, 9) = RsTxn_CxTF_ClientDetail!AdjRRNQty
          MSHFlexGrid2.TextMatrix(I, 15) = RsTxn_CxTF_ClientDetail!AdjRRNNoofBags
       End If
    ElseIf mExchangeTypeID = 3 Then
       RsTxn_CDC_GSLDetail.Find "CDCNo = " & RsTxn_CxTF_ClientDetail!RRNID
       RsMst_GSL.MoveFirst
       RsMst_GSL.Find "Mst_GSL_ID = " & RsTxn_CDC_GSLDetail!Mst_GSL_ID
       MSHFlexGrid2.TextMatrix(I, 7) = Val(RsMst_GSL!DepositeWeight)
       MSHFlexGrid2.TextMatrix(I, 8) = Format(Val(RsMst_GSL!DepositeWeight) - Val(RsMst_GSL!CDTFWeight), "#######0.000")
       MSHFlexGrid2.TextMatrix(I, 9) = Val(RsTxn_CxTF_ClientDetail!AdjRRNQty)
       MSHFlexGrid2.TextMatrix(I, 15) = RsTxn_CxTF_ClientDetail!AdjRRNNoofBags
       
    End If
    MSHFlexGrid2.TextMatrix(I, 11) = RsTxn_CxTF_ClientDetail!ClientID 'RsTxn_CxTF_ClientDetail!ClientIDRow
    MSHFlexGrid2.TextMatrix(I, 12) = IIf(IsNull(RsTxn_CxTF_ClientDetail!RRNID), "", RsTxn_CxTF_ClientDetail!RRNID)
    MSHFlexGrid2.TextMatrix(0, 7) = Val(MSHFlexGrid2.TextMatrix(0, 7)) + Val(MSHFlexGrid2.TextMatrix(I, 7))
    MSHFlexGrid2.TextMatrix(0, 8) = Format(Val(MSHFlexGrid2.TextMatrix(0, 8)) + Val(MSHFlexGrid2.TextMatrix(I, 8)), "#######0.000")
    MSHFlexGrid2.TextMatrix(0, 9) = Val(MSHFlexGrid2.TextMatrix(0, 9)) + Val(MSHFlexGrid2.TextMatrix(I, 9))
    MSHFlexGrid2.TextMatrix(0, 15) = Val(MSHFlexGrid2.TextMatrix(0, 15)) + Val(MSHFlexGrid2.TextMatrix(I, 15))
    
    RsTxn_CxTF_ClientDetail.MoveNext
Next
'MsgBox TmpRs.State

End Sub
Private Sub GetRWCDetails(i_, mExchangeTypeID_, RWCID_)
'Dim mRWCID_ As Double
If mExchangeTypeID_ = 1 Then
   If LCase(TxtPreFix) = "bknrp" Then
      mNewSM_Prefix = "BKNER"
      Call TableTxn_WHRDetail("Where (sm_prefix = '" & TxtPreFix & "' Or sm_prefix = '" & mNewSM_Prefix & "') And WHRNo = " & RWCID_ & " ")
   ElseIf LCase(TxtPreFix) = "ujjap" Then
      mNewSM_Prefix = "Ujjan"
      Call TableTxn_WHRDetail("Where (sm_prefix = '" & TxtPreFix & "' Or sm_prefix = '" & mNewSM_Prefix & "') And WHRNo = " & RWCID_ & " ")
   ElseIf LCase(TxtPreFix) = "dewsp" Then
      mNewSM_Prefix = "DEWAS"
      Call TableTxn_WHRDetail("Where (sm_prefix = '" & TxtPreFix & "' Or sm_prefix = '" & mNewSM_Prefix & "') And WHRNo = " & RWCID_ & " ")
   ElseIf LCase(TxtPreFix) = "vdshp" Then
      mNewSM_Prefix = "VDISA"
      Call TableTxn_WHRDetail("Where (sm_prefix = '" & TxtPreFix & "' Or sm_prefix = '" & mNewSM_Prefix & "') And WHRNo = " & RWCID_ & " ")
   ElseIf LCase(TxtPreFix) = "adlbp" Then ' 'adlbp
      mNewSM_Prefix = "adlbd"
      Call TableTxn_WHRDetail("Where (sm_prefix = '" & TxtPreFix & "' Or sm_prefix = '" & mNewSM_Prefix & "') And WHRNo = " & RWCID_ & " ")
   ElseIf LCase(TxtPreFix) = "aklap" Then ' 'adlbp
      mNewSM_Prefix = "akola"
      Call TableTxn_WHRDetail("Where (sm_prefix = '" & TxtPreFix & "' Or sm_prefix = '" & mNewSM_Prefix & "') And WHRNo = " & RWCID_ & " ")
   ElseIf LCase(TxtPreFix) = "sngrp" Then ' 'adlbp
      mNewSM_Prefix = "sgngr"
      Call TableTxn_WHRDetail("Where (sm_prefix = '" & TxtPreFix & "' Or sm_prefix = '" & mNewSM_Prefix & "') And WHRNo = " & RWCID_ & " ")
   ElseIf LCase(TxtPreFix) = "adlbd" Then ' 'adlbp
      mNewSM_Prefix = "adlbp"
      Call TableTxn_WHRDetail("Where (sm_prefix = '" & TxtPreFix & "' Or sm_prefix = '" & mNewSM_Prefix & "') And WHRNo = " & RWCID_ & " ")
   ElseIf LCase(TxtPreFix) = "nzmdp" Then ' 'adlbp
      mNewSM_Prefix = "nzmbd"
      Call TableTxn_WHRDetail("Where (sm_prefix = '" & TxtPreFix & "' Or sm_prefix = '" & mNewSM_Prefix & "') And WHRNo = " & RWCID_ & " ")
   ElseIf LCase(TxtPreFix) = "unjap" Then ' 'adlbp
      mNewSM_Prefix = "unjha"
      Call TableTxn_WHRDetail("Where (sm_prefix = '" & TxtPreFix & "' Or sm_prefix = '" & mNewSM_Prefix & "') And WHRNo = " & RWCID_ & " ")
   ElseIf LCase(TxtPreFix) = "sglip" Then ' 'adlbp
      mNewSM_Prefix = "sngli"
      Call TableTxn_WHRDetail("Where (sm_prefix = '" & TxtPreFix & "' Or sm_prefix = '" & mNewSM_Prefix & "') And WHRNo = " & RWCID_ & " ")
   ElseIf LCase(TxtPreFix) = "gntrp" Then ' 'adlbp
      mNewSM_Prefix = "guntr"
      Call TableTxn_WHRDetail("Where (sm_prefix = '" & TxtPreFix & "' Or sm_prefix = '" & mNewSM_Prefix & "') And WHRNo = " & RWCID_ & " ")
   ElseIf LCase(TxtPreFix) = "dvgrp" Then ' 'adlbp
      mNewSM_Prefix = "dvngr"
      Call TableTxn_WHRDetail("Where (sm_prefix = '" & TxtPreFix & "' Or sm_prefix = '" & mNewSM_Prefix & "') And WHRNo = " & RWCID_ & " ")
   ElseIf LCase(TxtPreFix) = "jdhrp" Then ' 'adlbp
      mNewSM_Prefix = "jdhpr"
      Call TableTxn_WHRDetail("Where (sm_prefix = '" & TxtPreFix & "' Or sm_prefix = '" & mNewSM_Prefix & "') And WHRNo = " & RWCID_ & " ")
   ElseIf LCase(TxtPreFix) = "hngtp" Then ' 'adlbp
      mNewSM_Prefix = "hghat"
      Call TableTxn_WHRDetail("Where (sm_prefix = '" & TxtPreFix & "' Or sm_prefix = '" & mNewSM_Prefix & "') And WHRNo = " & RWCID_ & " ")
   ElseIf LCase(TxtPreFix) = "kadip" Then ' 'adlbp
      mNewSM_Prefix = "kadi"
      Call TableTxn_WHRDetail("Where (sm_prefix = '" & TxtPreFix & "' Or sm_prefix = '" & mNewSM_Prefix & "') And WHRNo = " & RWCID_ & " ")
   ElseIf LCase(TxtPreFix) = "guntr" Then ' 'adlbp
      mNewSM_Prefix = "gntrp"
      Call TableTxn_WHRDetail("Where (sm_prefix = '" & TxtPreFix & "' Or sm_prefix = '" & mNewSM_Prefix & "') And WHRNo = " & RWCID_ & " ")
   ElseIf LCase(TxtPreFix) = "glbrg" Then ' 'adlbp
      mNewSM_Prefix = "gbrgp"
      Call TableTxn_WHRDetail("Where (sm_prefix = '" & TxtPreFix & "' Or sm_prefix = '" & mNewSM_Prefix & "') And WHRNo = " & RWCID_ & " ")
   ElseIf LCase(TxtPreFix) = "gbrgp" Then ' 'adlbp
      mNewSM_Prefix = "glbrg"
      Call TableTxn_WHRDetail("Where (sm_prefix = '" & TxtPreFix & "' Or sm_prefix = '" & mNewSM_Prefix & "') And WHRNo = " & RWCID_ & " ")
   ElseIf LCase(TxtPreFix) = "kngym" Then ' 'adlbp
      mNewSM_Prefix = "kngmp"
      Call TableTxn_WHRDetail("Where (sm_prefix = '" & TxtPreFix & "' Or sm_prefix = '" & mNewSM_Prefix & "') And WHRNo = " & RWCID_ & " ")
   ElseIf LCase(TxtPreFix) = "lturp" Then ' 'Latur
      mNewSM_Prefix = "latur"
      Call TableTxn_WHRDetail("Where (sm_prefix = '" & TxtPreFix & "' Or sm_prefix = '" & mNewSM_Prefix & "') And WHRNo = " & RWCID_ & " ")
   ElseIf LCase(TxtPreFix) = "ktap" Then ' 'KOTA
      mNewSM_Prefix = "kota"
      Call TableTxn_WHRDetail("Where (sm_prefix = '" & TxtPreFix & "' Or sm_prefix = '" & mNewSM_Prefix & "') And WHRNo = " & RWCID_ & " ")
   ElseIf LCase(TxtPreFix) = "rsgrp" Then ' 'KOTA
      mNewSM_Prefix = "rsngr"
      Call TableTxn_WHRDetail("Where (sm_prefix = '" & TxtPreFix & "' Or sm_prefix = '" & mNewSM_Prefix & "') And WHRNo = " & RWCID_ & " ")
   ElseIf LCase(TxtPreFix) = "rsngr" Then ' 'KOTA
      mNewSM_Prefix = "rsgrp"
      Call TableTxn_WHRDetail("Where (sm_prefix = '" & TxtPreFix & "' Or sm_prefix = '" & mNewSM_Prefix & "') And WHRNo = " & RWCID_ & " ")
   ElseIf LCase(TxtPreFix) = "ttcrp" Then ' 'KOTA
      mNewSM_Prefix = "tcrin"
      Call TableTxn_WHRDetail("Where (sm_prefix = '" & TxtPreFix & "' Or sm_prefix = '" & mNewSM_Prefix & "') And WHRNo = " & RWCID_ & " ")
   Else
      Call TableTxn_WHRDetail("Where sm_prefix = '" & TxtPreFix & "' And WHRNo = " & RWCID_ & " ")
   End If
   
   'mRWCID_ = mCMPID & "-" & (RsTxn_WHRDetail!WHRNo - mCMPID * 1000000)
   mRWCID_ = Format((RsTxn_WHRDetail!WHRNo / 1000000), "#####")
   'mRWCID_  =
   mRWCID_ = mRWCID_ & "-" & (RsTxn_WHRDetail!WHRNo - (mRWCID_ * 1000000))
   
   mRWCQty_ = RsTxn_WHRDetail!WHRQuantity
ElseIf mExchangeTypeID_ = 2 And Check1.Value = 0 Then
   Call TableMst_RRN("Where RRNID = " & RWCID_ & " ")
   mRWCID_ = RsMst_RRN!RRNNo
   mISINNo_ = RsMst_RRN!ISINID
   mRWCQty_ = RsMst_RRN!RRNQty
ElseIf mExchangeTypeID_ = 3 Then
   Call TableTxn_CDC_GSLDetail("Where sm_prefix = '" & TxtPreFix & "' And CDCNo = " & RWCID_ & "")
   mRWCID_ = RsTxn_CDC_GSLDetail!CDCNo
   mRWCQty_ = RsTxn_CDC_GSLDetail!StorageAmount
End If
MSHFlexGrid2.TextMatrix(i_, 1) = mRWCID_
MSHFlexGrid2.TextMatrix(i_, 2) = mISINNo_
MSHFlexGrid2.TextMatrix(i_, 7) = mRWCQty_
'MSHFlexGrid2.TextMatrix(x, 8) = mRWCID_
'MSHFlexGrid2.TextMatrix(x, 9) = mRWCID_
End Sub

Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   If RsTxn_CxTF_Details.RecordCount = 0 Then Exit Sub
   ans = MsgBox("Do you really want to Cancel entered details?", vbYesNo)
   If ans = vbNo Then
      Exit Sub
   End If
   Call GButtonLock(Me, True)
   Call FirstCmd_Click
   Exit Sub
End If
'
'Frame0.Height = Me.Height - 350
'Frame0.Left = Frame4.Left - 100
'Frame0.Top = Frame4.Top - 100
'Frame0.Width = Me.Width - 150
'

Frame0.Height = Frame4.Height - 100
Frame0.Left = Picture2.Left '- 100
Frame0.Top = Picture2.Top '- 100
Frame0.Width = Frame4.Width '- 150

Frame0.Visible = True
MSHFlexGrid1.Width = Frame0.Width - 400
MSHFlexGrid1.Height = Frame0.Height - 1800
RsTxn_CxTF_Details.MoveFirst

''===========comment on 7 july 2012 for not clearing search grid=============
'MSHFlexGrid1.Clear
'MSHFlexGrid1.Cols = 3
'Call Grid_Head
'CmbSCmp.Text = ""
'CmbSCommodity.Text = ""
'CmbSExchangeType.Text = ""
'CmbSAgent.Text = ""
'SFrom.Value = Date - 730
'STo.Value = Date
'TxtSCxTFNo = ""
'TxtSManualCxTFNo = ""
'CmbSRejection.Text = ""
''===========comment on 7 july 2012 for not clearing search grid=============

'MSHFlexGrid1.Rows = RsTxn_CxTF_Details.RecordCount + 1
'MSHFlexGrid1.ColWidth(0) = 1000
'For i = 1 To RsTxn_CxTF_Details.RecordCount
'    MSHFlexGrid1.TextMatrix(i, 0) = RsTxn_CxTF_Details!CxTFNo
'    MSHFlexGrid1.TextMatrix(i, 1) = RsTxn_CxTF_Details!CxTFDate
'    MSHFlexGrid1.TextMatrix(i, 2) = GetCMPName(RsTxn_CxTF_Details!CMPID)
'    MSHFlexGrid1.TextMatrix(i, 3) = GetCommodityName(RsTxn_CxTF_Details!CommodityID)
'    MSHFlexGrid1.TextMatrix(i, 4) = GetExchangeName(RsTxn_CxTF_Details!ExchangeTypeID)
'    MSHFlexGrid1.TextMatrix(i, 5) = RsTxn_CxTF_Details!SM_Prefix
'    MSHFlexGrid1.TextMatrix(i, 6) = IIf(IsNull(RsTxn_CxTF_Details!Flag), "", RsTxn_CxTF_Details!Flag)
'    MSHFlexGrid1.TextMatrix(i, 7) = IIf(IsNull(RsTxn_CxTF_Details!Agent), "", RsTxn_CxTF_Details!Agent)
'    MSHFlexGrid1.TextMatrix(i, 8) = IIf(IsNull(RsTxn_CxTF_Details!Remark), "", RsTxn_CxTF_Details!Remark)
'    MSHFlexGrid1.TextMatrix(i, 9) = "Withdrawal"
'    RsTxn_CxTF_Details.MoveNext
'Next
End Sub
Private Sub Grid_Head()
'MSHFlexGrid1.Clear
'MSHFlexGrid1.Rows = 2
'MSHFlexGrid1.Cols = 11
'MSHFlexGrid1.TextMatrix(0, 0) = "CITF No"
'MSHFlexGrid1.TextMatrix(0, 1) = "CITF Date"
'MSHFlexGrid1.TextMatrix(0, 2) = "CMP"
'MSHFlexGrid1.TextMatrix(0, 3) = "Commodity"
'MSHFlexGrid1.TextMatrix(0, 4) = "Exchange Type"
'MSHFlexGrid1.TextMatrix(0, 5) = "Prefix"
'MSHFlexGrid1.TextMatrix(0, 6) = "Flag"
'MSHFlexGrid1.TextMatrix(0, 7) = "Agent"
'MSHFlexGrid1.TextMatrix(0, 8) = "Remark"
'MSHFlexGrid1.TextMatrix(0, 9) = "Txn Type"
'MSHFlexGrid1.TextMatrix(0, 10) = "Booked Space ID"


MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 27 '26 '25
MSHFlexGrid1.TextMatrix(0, 0) = "CITF No"
MSHFlexGrid1.TextMatrix(0, 1) = "CITF Date"
MSHFlexGrid1.TextMatrix(0, 2) = "Exchange Type"
MSHFlexGrid1.TextMatrix(0, 3) = "CMP"
MSHFlexGrid1.TextMatrix(0, 4) = "Location ID"
MSHFlexGrid1.TextMatrix(0, 5) = "Commodity"
MSHFlexGrid1.TextMatrix(0, 6) = "Flag"
MSHFlexGrid1.TextMatrix(0, 7) = "Agent"
MSHFlexGrid1.TextMatrix(0, 8) = "Remark"
MSHFlexGrid1.TextMatrix(0, 9) = "Hologram No"
MSHFlexGrid1.TextMatrix(0, 10) = "Agent Name"
MSHFlexGrid1.TextMatrix(0, 11) = "Client Name"
MSHFlexGrid1.TextMatrix(0, 12) = "Manual CITF No"
MSHFlexGrid1.TextMatrix(0, 13) = "No of Bags"
MSHFlexGrid1.TextMatrix(0, 14) = "Gross Wt."
MSHFlexGrid1.TextMatrix(0, 15) = "Guid"
MSHFlexGrid1.TextMatrix(0, 16) = "Rejected"
MSHFlexGrid1.TextMatrix(0, 17) = "Print Approve Flag"
MSHFlexGrid1.TextMatrix(0, 18) = "Print Approved By"
MSHFlexGrid1.TextMatrix(0, 19) = "Print Approve Remark"
MSHFlexGrid1.TextMatrix(0, 20) = "Print Approve Date"
MSHFlexGrid1.TextMatrix(0, 21) = "TI No"
MSHFlexGrid1.TextMatrix(0, 22) = "Post Invoice No"
MSHFlexGrid1.TextMatrix(0, 23) = "Outstanding Amount"
MSHFlexGrid1.TextMatrix(0, 24) = "Balance Stock"
MSHFlexGrid1.TextMatrix(0, 25) = "Commodity Value"
MSHFlexGrid1.TextMatrix(0, 26) = "SCM Flag"

MSHFlexGrid1.ColWidth(0) = 1000
MSHFlexGrid1.ColWidth(1) = 1000
MSHFlexGrid1.ColWidth(2) = 1400
MSHFlexGrid1.ColWidth(3) = 1600
MSHFlexGrid1.ColWidth(4) = 1200
MSHFlexGrid1.ColWidth(5) = 1200
MSHFlexGrid1.ColWidth(6) = 1200
MSHFlexGrid1.ColWidth(7) = 1200
MSHFlexGrid1.ColWidth(8) = 1600
MSHFlexGrid1.ColWidth(9) = 1000
MSHFlexGrid1.ColWidth(10) = 1200
MSHFlexGrid1.ColWidth(11) = 1200
MSHFlexGrid1.ColWidth(15) = 0
MSHFlexGrid1.ColWidth(16) = 1000

MSHFlexGrid1.RowHeight(0) = 700

End Sub
Private Sub MSHFlexGrid1_Click()
'Dim i As Variant
'i = MSHFlexGrid1.RowSel
'If i <= 0 Then Exit Sub
'If MSHFlexGrid1.TextMatrix(i, 0) = "" Then Exit Sub
'RsTxn_CxTF_Details.MoveFirst
'RsTxn_CxTF_Details.Find "sm_prefix  = '" & MSHFlexGrid1.TextMatrix(i, 5) & "'"
'If Not RsTxn_CxTF_Details.EOF Then
'   RsTxn_CxTF_Details.Find "CxTFNo = " & MSHFlexGrid1.TextMatrix(i, 0)
'   Call InData
'   Frame0.Visible = False
'End If

Dim I As Variant
I = MSHFlexGrid1.RowSel
If I <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(I, 0) = "" Then Exit Sub
RsTxn_CxTF_Details.MoveFirst
RsTxn_CxTF_Details.Find "g_uid = '" & MSHFlexGrid1.TextMatrix(I, 15) & "'"
If Not RsTxn_CxTF_Details.EOF Then
   Call Indata
   Frame0.Visible = False
End If


End Sub
Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Txn_CxTF_Details.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Txn_CxTF_Details.xls")
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
MsgBox ("Excel file created!!! ")
oXL.Visible = True
End Sub

Private Sub TxtAdjWt_LostFocus()
If TxtAdjWt = "" Then
   TxtAdjWt = Val(TxtAdjWt)
Else
    If IsNumeric(TxtAdjWt) = False Then
       MsgBox "Invalid numeric format!!!"
       TxtAdjWt.SetFocus
       Exit Sub
    End If
End If
Call CalDmatWt
End Sub
Private Sub TxtCxTFDate_LostFocus()
If TxtCxTFDate.Value > Date Then
   MsgBox "Future date is not allowed!!! "
   TxtCxTFDate.SetFocus
End If
'If TxtCxTFDate.Value < mBODDate Then
'   MsgBox "CITF Date less than BOD Date is not allowed!!! "
'   TxtCxTFDate.SetFocus
'End If

If TxtCxTFDate.Value < mBODDate Then
   MsgBox mBlockMessage '"CDTF Date less than BOD Date is not allowed!!! "
   TxtCxTFDate.SetFocus
End If

TxtCxTFDispalyDate = TxtCxTFDate.Value
End Sub
Private Sub TxtGCTDGrossWt_LostFocus()
If TxtGCTDGrossWt = "" Then Exit Sub

If Not IsNumeric(TxtGCTDGrossWt) Or Val(TxtGCTDGrossWt) < 0 Then
   MsgBox "Invalid numeric format!!!"
   TxtGCTDGrossWt.SetFocus
   Exit Sub
End If
Call CalDmatWt

End Sub
Private Sub TxtGCTDNoofBags_LostFocus()
If TxtGCTDNoofBags = "" Then Exit Sub
If Not IsNumeric(TxtGCTDNoofBags) Or Val(TxtGCTDNoofBags) < 0 Then
   MsgBox "Invalid numeric format!!!"
   TxtGCTDNoofBags.SetFocus
   Exit Sub
End If

TxtGCTDNoofBags = Format(TxtGCTDNoofBags, "#######0")


If Val(TxtGSLAvailableNofBags) < Val(TxtGCTDNoofBags) Then
   MsgBox "Delivered number of bags must not be more than available number of bags!!!"
   TxtGCTDNoofBags.SetFocus
   Exit Sub
End If
TxtGCTDGrossWt = Format(Val(TxtGCTDNoofBags) * mTruckWtRatio, "#####0.000")
Call CalDmatWt
End Sub
Private Sub CalDmatWt()

If mCommodityID = 31 Or mCommodityID = 114 Then
   TxtGGunnyWeight = Format((Val(TxtGCTDNoofBags) * Val(mChilliGunnyWeight)) / 1000, "#######0.000")
Else
   TxtGGunnyWeight = Format((Val(TxtGCTDNoofBags) * Val(TxtCGunnyWt)) / 1000, "#######0.000")
End If

TxtGSLNetWeight = Format(Val(TxtGCTDGrossWt) + Val(TxtAdjWt) - (Val(TxtGGunnyWeight)), "#######0.000")
If Val(TxtGSLAvailableNofBags) = Val(TxtGCTDNoofBags) Then
   TxtGSLWeightGainLoss = Format((Val(TxtGCTDGrossWt) - Val(TxtGGunnyWeight)) - (Val(TxtGSLAvailableGrossWgt)), "#######0.000")
Else
   TxtGSLWeightGainLoss = 0
End If

TxtAverageBagSize = ""
If Val(TxtGCTDNoofBags) > 0 Then
   TxtAverageBagSize = Format(Val(TxtGSLNetWeight) * 1000 / Val(TxtGCTDNoofBags), "#######0.00")
End If

End Sub

Private Sub TxtRecoverClient_LostFocus()
If TxtRecoverClient = "" Then Exit Sub
If IsNumeric(TxtRecoverClient) = False Or Val(TxtRecoverClient) < 0 Then
   MsgBox "Invalid numeric format!!!"
   TxtRecoverClient.SetFocus
   Exit Sub
End If
End Sub

Private Sub TxtRecoverNCMSL_LostFocus()
If TxtRecoverNCMSL = "" Then Exit Sub
If IsNumeric(TxtRecoverNCMSL) = False Or Val(TxtRecoverNCMSL) < 0 Then
   MsgBox "Invalid numeric format!!!"
   TxtRecoverNCMSL.SetFocus
   Exit Sub
End If
End Sub

Private Sub TxtMoisturePer_LostFocus()
If TxtMoisturePer.Text = "" Then Exit Sub

If Not IsNumeric(TxtMoisturePer.Text) Then
   MsgBox "Invalid Format"
   TxtMoisturePer.SetFocus
   Exit Sub
End If
End Sub


Private Sub TxtRRNNo_GotFocus()
CmbRRNISIN.Visible = False
End Sub

Private Sub TxtRRNNo_KeyPress(KeyAscii As Integer)
tmp = KeyAscii
If KeyAscii = 14 Then
   SSTab1.Tab = 1
   TxtTTDTruckNo.SetFocus
End If
End Sub

Private Sub TxtRRNNo_LostFocus()
If CmbExchangeTypeID.Text = "" Then
   MsgBox "Please select a Mode!!!"
   CmbExchangeTypeID.SetFocus
   Exit Sub
End If
If TxtRRNNo = "" Then
   Exit Sub
End If
If mExchangeTypeID = 1 Then 'NON NCDEX
   Dim strSplitString() As String
   strSplitString = Split(TxtRRNNo.Text, "-")
   tmp = UBound(strSplitString)
   If tmp <> 1 Then
      MsgBox "Invalid WHR No selection!!!"
      TxtRRNNo.SetFocus
      Exit Sub
   End If
   tmp = (Val(strSplitString(0)) * 1000000) + Val(strSplitString(1))
   Call WHRDetails(mCMPID, mCommodityID, tmp)
End If
If mExchangeTypeID = 2 Then 'NCDEX
   Call RRNDetails(mCMPID, mCommodityID, TxtRRNNo.Text)
End If
If mExchangeTypeID = 3 Then 'SPOT
   Call CDCDetails(mCMPID, mCommodityID, TxtRRNNo.Text)
End If
End Sub
Private Sub RRNDetails(mCMPID_, mCommodityID_, mTxtRRNNo_)
If mExchangeTypeID = 0 Or IsNull(mExchangeTypeID) Then Exit Sub
If mCMPID_ = 0 Or IsNull(mCMPID_) Then Exit Sub
If mCommodityID_ = 0 Or IsNull(mCommodityID_) Then Exit Sub
If mTxtRRNNo_ = "" Then Exit Sub

Call TableMst_RRN("Where RRNNo = '" & mTxtRRNNo_ & "'")
mRRNIds_ = ""
If RsMst_RRN.RecordCount = 0 Then
   MsgBox "No Record found!!!"
   TxtRRNNo.SetFocus
   Exit Sub
End If
RsMst_RRN.MoveFirst
RsMst_RRN.Filter = "RRNNo = '" & TxtRRNNo & "'"

'RsMst_RRN.Find "RRNNo = '" & TxtRRNNo & "'"
If RsMst_RRN.RecordCount > 1 Then
   CmbRRNISIN.Visible = True
   CmbRRNISIN.Clear
   For iu = 1 To RsMst_RRN.RecordCount
       CmbRRNISIN.AddItem RsMst_RRN!ISINID
       RsMst_RRN.MoveNext
   Next
   Exit Sub
End If

If Not RsMst_RRN.EOF Then
   TxtISINno = IIf(IsNull(RsMst_RRN!ISINID), "", RsMst_RRN!ISINID)
   mRRNCommTrackFlag = IIf(IsNull(RsMst_RRN!CommTrackFlag), 0, RsMst_RRN!CommTrackFlag)
   If mRRNCommTrackFlag = 1 Then
      If MSHFlexGrid2.Rows > 3 And MSHFlexGrid10.Rows < 4 Then
         MsgBox "Please enter non Comtrack RRN only!!"
         TxtRRNNo.SetFocus
         Exit Sub
      End If
   Else
      If MSHFlexGrid10.Rows > 3 Then
         MsgBox "Please enter comtrack RRN only!!"
         TxtRRNNo.SetFocus
         Exit Sub
      End If
   End If
   Call TableMst_ISIN("Where ISINID = '" & TxtISINno & "'")
   Call TableMst_Location("Where LocationID = '" & RsMst_ISIN!LocationID & "'")
   If Not RsMst_ISIN.EOF Then
      If TxtPreFix <> RsMst_Location!SM_Prefix Then
'         If LCase(TxtPreFix) <> "bknrp" Then
'            MsgBox "ISIN Location and above CMP location do not match!!!"
'            TxtRRNNo.SetFocus
'            Exit Sub
'         Else
'            If LCase(RsMst_Location!SM_Prefix) <> "bkner" Then
'               MsgBox "ISIN Location should be Bikaner only!!!"
'               TxtRRNNo.SetFocus
'               Exit Sub
'            End If
'         End If
      
         If LCase(TxtPreFix) <> "bknrp" And LCase(TxtPreFix) <> "ujjap" And LCase(TxtPreFix) <> "dewsp" And LCase(TxtPreFix) <> "vdshp" And LCase(TxtPreFix) <> "adlbp" And LCase(TxtPreFix) <> "aklap" And LCase(TxtPreFix) <> "sngrp" And LCase(TxtPreFix) <> "adlbd" And LCase(TxtPreFix) <> "nzmdp" And LCase(TxtPreFix) <> "unjap" And LCase(TxtPreFix) <> "sglip" And LCase(TxtPreFix) <> "gntrp" And LCase(TxtPreFix) <> "dvgrp" And LCase(TxtPreFix) <> "jdhrp" And LCase(TxtPreFix) <> "hngtp" And LCase(TxtPreFix) <> "kadip" And LCase(TxtPreFix) <> "guntr" And LCase(TxtPreFix) <> "glbrg" And LCase(TxtPreFix) <> "gbrgp" And LCase(TxtPreFix) <> "kngmp" And LCase(TxtPreFix) <> "lturp" And LCase(TxtPreFix) <> "ktap" And LCase(TxtPreFix) <> "rsgrp" And LCase(TxtPreFix) <> "rsngr" And LCase(TxtPreFix) <> "ttcrp" Then
            MsgBox "CMSE Lot No Location and above CMP location do not match!!!"
            TxtRRNNo.SetFocus
            Exit Sub
         Else
            If LCase(TxtPreFix) = "bknrp" Then
               If LCase(RsMst_Location!SM_Prefix) <> "bkner" Then
                  MsgBox "CMSE Lot No Location should be Bikaner only!!!"
                  TxtRRNNo.SetFocus
                  Exit Sub
               End If
            ElseIf LCase(TxtPreFix) = "ujjap" Then
               If LCase(RsMst_Location!SM_Prefix) <> "ujjan" Then
                  MsgBox "CMSE Lot No Location should be Ujjain only!!!"
                  TxtRRNNo.SetFocus
                  Exit Sub
               End If
            ElseIf LCase(TxtPreFix) = "dewsp" Then
               If LCase(RsMst_Location!SM_Prefix) <> "dewas" Then
                  MsgBox "CMSE Lot No Location should be Dewas only!!!"
                  TxtRRNNo.SetFocus
                  Exit Sub
               End If
            ElseIf LCase(TxtPreFix) = "vdshp" Then
               If LCase(RsMst_Location!SM_Prefix) <> "vdisa" Then
                  MsgBox "CMSE Lot No Location should be Vidisha only!!!"
                  TxtRRNNo.SetFocus
                  Exit Sub
               End If
            ElseIf LCase(TxtPreFix) = "adlbp" Then
               If LCase(RsMst_Location!SM_Prefix) <> "adlbd" Then
                  MsgBox "CMSE Lot No Location should be Adilabad only!!!"
                  TxtRRNNo.SetFocus
                  Exit Sub
               End If
            ElseIf LCase(TxtPreFix) = "aklap" Then
               If LCase(RsMst_Location!SM_Prefix) <> "akola" Then
                  MsgBox "CMSE Lot No Location should be Akola only!!!"
                  TxtRRNNo.SetFocus
                  Exit Sub
               End If
            ElseIf LCase(TxtPreFix) = "sngrp" Then
               If LCase(RsMst_Location!SM_Prefix) <> "sgngr" Then
                  MsgBox "CMSE Lot No Location should be Sriganganagar only!!!"
                  TxtRRNNo.SetFocus
                  Exit Sub
               End If
            ElseIf LCase(TxtPreFix) = "adlbd" Then
               If LCase(RsMst_Location!SM_Prefix) <> "adlbp" Then
                  MsgBox "CMSE Lot No Location should be Adilabad P only!!!"
                  TxtRRNNo.SetFocus
                  Exit Sub
               End If
            ElseIf LCase(TxtPreFix) = "nzmdp" Then
               If LCase(RsMst_Location!SM_Prefix) <> "nzmbd" Then
                  MsgBox "CMSE Lot No Location should be Nizamabad only!!!"
                  TxtRRNNo.SetFocus
                  Exit Sub
               End If
            ElseIf LCase(TxtPreFix) = "unjap" Then
               If LCase(RsMst_Location!SM_Prefix) <> "unjha" Then
                  MsgBox "CMSE Lot No Location should be Unjha only!!!"
                  TxtRRNNo.SetFocus
                  Exit Sub
               End If
            ElseIf LCase(TxtPreFix) = "sglip" Then
               If LCase(RsMst_Location!SM_Prefix) <> "sngli" Then
                  MsgBox "CMSE Lot No Location should be Sangli only!!!"
                  TxtRRNNo.SetFocus
                  Exit Sub
               End If
            ElseIf LCase(TxtPreFix) = "gntrp" Then
               If LCase(RsMst_Location!SM_Prefix) <> "guntr" Then
                  MsgBox "CMSE Lot No Location should be Guntur only!!!"
                  TxtRRNNo.SetFocus
                  Exit Sub
               End If
            ElseIf LCase(TxtPreFix) = "dvgrp" Then
               If LCase(RsMst_Location!SM_Prefix) <> "dvngr" Then
                  MsgBox "CMSE Lot No Location should be Davangeere Only!!!"
                  TxtRRNNo.SetFocus
                  Exit Sub
               End If
            ElseIf LCase(TxtPreFix) = "jdhrp" Then
               If LCase(RsMst_Location!SM_Prefix) <> "jdhpr" Then
                  MsgBox "CMSE Lot No Location should be Jodhpur Only!!!"
                  TxtRRNNo.SetFocus
                  Exit Sub
               End If
            ElseIf LCase(TxtPreFix) = "hngtp" Then
               If LCase(RsMst_Location!SM_Prefix) <> "hghat" Then
                  MsgBox "CMSE Lot No Location should be Hinghanghat Only!!!"
                  TxtRRNNo.SetFocus
                  Exit Sub
               End If
            ElseIf LCase(TxtPreFix) = "kadip" Then
               If LCase(RsMst_Location!SM_Prefix) <> "kadi" Then
                  MsgBox "CMSE Lot No Location should be Kadi Only!!!"
                  TxtRRNNo.SetFocus
                  Exit Sub
               End If
            ElseIf LCase(TxtPreFix) = "gbrgp" Then
               If LCase(RsMst_Location!SM_Prefix) <> "glbrg" Then
                  MsgBox "CMSE Lot No Location should be Kadi Only!!!"
                  TxtRRNNo.SetFocus
                  Exit Sub
               End If
            ElseIf LCase(TxtPreFix) = "glbrg" Then
               If LCase(RsMst_Location!SM_Prefix) <> "gbrgp" Then
                  MsgBox "CMSE Lot No Location should be Gulbarga Only!!!"
                  TxtRRNNo.SetFocus
                  Exit Sub
               End If
            ElseIf LCase(TxtPreFix) = "guntr" Then
               If LCase(RsMst_Location!SM_Prefix) <> "gntrp" Then
                  MsgBox "CMSE Lot No Location should be Guntur only!!!"
                  TxtRRNNo.SetFocus
                  Exit Sub
               End If
            ElseIf LCase(TxtPreFix) = "kngym" Then
               If LCase(RsMst_Location!SM_Prefix) <> "kngmp" Then
                  MsgBox "CMSE Lot No Location should be Kangayam only!!!"
                  TxtRRNNo.SetFocus
                  Exit Sub
               End If
            ElseIf LCase(TxtPreFix) = "lturp" Then
               If LCase(RsMst_Location!SM_Prefix) <> "latur" Then
                  MsgBox "CMSE Lot No Location should be Latur only!!!"
                  TxtRRNNo.SetFocus
                  Exit Sub
               End If
            ElseIf LCase(TxtPreFix) = "ktap" Then
               If LCase(RsMst_Location!SM_Prefix) <> "kota" Then
                  MsgBox "CMSE Lot No Location should be Kota only!!!"
                  TxtRRNNo.SetFocus
                  Exit Sub
               End If
            ElseIf LCase(TxtPreFix) = "rsgrp" Then
               If LCase(RsMst_Location!SM_Prefix) <> "rsngr" Then
                  MsgBox "CMSE Lot No Location should be Raisingnagar only!!!"
                  TxtRRNNo.SetFocus
                  Exit Sub
               End If
            ElseIf LCase(TxtPreFix) = "rsngr" Then
               If LCase(RsMst_Location!SM_Prefix) <> "rsgrp" Then
                  MsgBox "CMSE Lot No Location should be Raisingnagar P only!!!"
                  TxtRRNNo.SetFocus
                  Exit Sub
               End If
            ElseIf LCase(TxtPreFix) = "ttcrp" Then
               If LCase(RsMst_Location!SM_Prefix) <> "tcrin" Then
                  MsgBox "CMSE Lot No Location should be Kota only!!!"
                  TxtRRNNo.SetFocus
                  Exit Sub
               End If
            End If
         End If
      End If
      ' Removed CMP Level Check on 31.03.2010
'      If mCMPID_ <> RsMst_ISIN!CMPID Then
'         MsgBox "ISIN CMP and above CMP do not match!!!"
'         TxtRRNNo.SetFocus
'         Exit Sub
'      End If
      
'Commented on 09 July 2013
'      If mCommodityID_ <> RsMst_ISIN!CommodityID Then
'         MsgBox "ISIN commodity and selected Commodity do not match!!!"
'         TxtRRNNo.SetFocus
'         Exit Sub
'      End If
      
      'Condition introduced for withdrawal of Chilli and Chilli 25 mail from Mayur Sir on 09 July 2013
      If mCommodityID_ <> 31 And mCommodityID_ <> 114 Then
         If mCommodityID_ <> RsMst_ISIN!CommodityID Then
            MsgBox "CMSE Lot No commodity and selected Commodity do not match!!!"
            TxtRRNNo.SetFocus
            Exit Sub
         End If
      ElseIf mCommodityID_ = 31 Then
         If RsMst_ISIN!CommodityID <> 31 And RsMst_ISIN!CommodityID <> 114 Then
            MsgBox "CMSE Lot No commodity should be either Chilli or Chilli 25!!!"
            TxtRRNNo.SetFocus
            Exit Sub
         End If
      ElseIf mCommodityID_ = 114 Then
         If RsMst_ISIN!CommodityID <> 31 And RsMst_ISIN!CommodityID <> 114 Then
            MsgBox "CMSE Lot No commodity should be either Chilli or Chilli 25!!!"
            TxtRRNNo.SetFocus
            Exit Sub
         End If
      End If
  End If
  mRWCID = RsMst_RRN!RRNID
  TxtRRNQty = IIf(IsNull(RsMst_RRN!RRNQty), "", RsMst_RRN!RRNQty)
  'Commneted On
  'TxtBalanceQty = Format(Val(TxtRRNQty) - Val(RsMst_RRN!DeliveredQty), "#######0.000")
  TxtBalanceQty = Format(Val(TxtRRNQty) - (Val(RsMst_RRN!DeliveredQty) + Val(RsMst_RRN!OutOfSystemQty)), "#######0.000")
  mClientIDRow = RsMst_RRN!ClientID
  Call GetClientDetails(mClientIDRow)
End If

End Sub
Private Sub WHRDetails(mCMPID_, mCommodityID_, mTxtRRNNo_)
If mExchangeTypeID = 0 Or IsNull(mExchangeTypeID) Then Exit Sub
If mCMPID_ = 0 Or IsNull(mCMPID_) Then Exit Sub
If mCommodityID_ = 0 Or IsNull(mCommodityID_) Then Exit Sub
If mTxtRRNNo_ = "" Then Exit Sub

'Call TableTxn_WHRDetail("Where CMPID = " & mCMPID_ & " And CommodityId = " & mCommodityID_ & "")

If LCase(TxtPreFix) <> "bknrp" And LCase(TxtPreFix) <> "ujjap" And LCase(TxtPreFix) <> "dewsp" And LCase(TxtPreFix) <> "vdshp" And LCase(TxtPreFix) <> LCase("ADLBP") And LCase(TxtPreFix) <> LCase("aklap") And LCase(TxtPreFix) <> LCase("sngrp") And LCase(TxtPreFix) <> LCase("adlbd") And LCase(TxtPreFix) <> LCase("nzmdp") And LCase(TxtPreFix) <> LCase("unjap") And LCase(TxtPreFix) <> LCase("sglip") And LCase(TxtPreFix) <> "gntrp" And LCase(TxtPreFix) <> LCase("guntr") And LCase(TxtPreFix) <> LCase("dvgrp") And LCase(TxtPreFix) <> LCase("jdhrp") And LCase(TxtPreFix) <> LCase("hngtp") And LCase(TxtPreFix) <> LCase("kadip") And LCase(TxtPreFix) <> LCase("gbrgp") And LCase(TxtPreFix) <> LCase("glbrg") And LCase(TxtPreFix) <> LCase("kngym") And LCase(TxtPreFix) <> "lturp" And LCase(TxtPreFix) <> "ktap" And LCase(TxtPreFix) <> "rsgrp" And LCase(TxtPreFix) <> "rsngr" And LCase(TxtPreFix) <> "ttcrp" Then
   If ChkSCMFlag.Value = 0 Then
      Call TableTxn_WHRDetail("Where  sm_prefix = '" & TxtPreFix & "' And WHRNo = " & mTxtRRNNo_ & " And SCMFlag = 0 ")
   Else
      Call TableTxn_WHRDetail("Where  sm_prefix = '" & TxtPreFix & "' And WHRNo = " & mTxtRRNNo_ & " And SCMFlag = 1 ")
   End If
Else
   If LCase(TxtPreFix) = "bknrp" Then
      mNewSM_Prefix = "BKNER"
   ElseIf LCase(TxtPreFix) = "ujjap" Then
      mNewSM_Prefix = "Ujjan"
   ElseIf LCase(TxtPreFix) = "dewsp" Then
      mNewSM_Prefix = "DEWAS"
   ElseIf LCase(TxtPreFix) = "vdshp" Then
      mNewSM_Prefix = "VDISA"
   ElseIf LCase(TxtPreFix) = "adlbp" Then 'ADLBD
      mNewSM_Prefix = "adlbd"
   ElseIf LCase(TxtPreFix) = "aklap" Then 'ADLBD
      mNewSM_Prefix = "AKOLA"
   ElseIf LCase(TxtPreFix) = "sngrp" Then 'ADLBD
      mNewSM_Prefix = "sgngr"
   ElseIf LCase(TxtPreFix) = "adlbd" Then 'ADLBD
      mNewSM_Prefix = "adlbp"
   ElseIf LCase(TxtPreFix) = "nzmdp" Then 'ADLBD
      mNewSM_Prefix = "nzmbd"
   ElseIf LCase(TxtPreFix) = "unjap" Then 'ADLBD
      mNewSM_Prefix = "unjha"
   ElseIf LCase(TxtPreFix) = "sglip" Then 'ADLBD
      mNewSM_Prefix = "sngli"
   ElseIf LCase(TxtPreFix) = "gntrp" Then 'ADLBD
      mNewSM_Prefix = "guntr"
   ElseIf LCase(TxtPreFix) = "dvgrp" Then 'ADLBD
      mNewSM_Prefix = "dvngr"
   ElseIf LCase(TxtPreFix) = "jdhrp" Then 'ADLBD
      mNewSM_Prefix = "jdhpr"
   ElseIf LCase(TxtPreFix) = "hngtp" Then 'ADLBD
      mNewSM_Prefix = "hghat"
   ElseIf LCase(TxtPreFix) = "kadip" Then 'ADLBD
      mNewSM_Prefix = "kadi"
   ElseIf LCase(TxtPreFix) = "guntr" Then 'ADLBD
      mNewSM_Prefix = "gntrp"
   ElseIf LCase(TxtPreFix) = "glbrg" Then 'ADLBD
      mNewSM_Prefix = "gbrgp"
   ElseIf LCase(TxtPreFix) = "gbrgp" Then 'ADLBD
      mNewSM_Prefix = "glbrg"
   ElseIf LCase(TxtPreFix) = "kngym" Then 'ADLBD
      mNewSM_Prefix = "kngmp"
   ElseIf LCase(TxtPreFix) = "lturp" Then 'ADLBD
      mNewSM_Prefix = "latur"
   ElseIf LCase(TxtPreFix) = "ktap" Then 'Kota P
      mNewSM_Prefix = "kota"
   ElseIf LCase(TxtPreFix) = "rsgrp" Then 'Kota P
      mNewSM_Prefix = "rsngr"
   ElseIf LCase(TxtPreFix) = "rsngr" Then 'Kota P
      mNewSM_Prefix = "rsgrp"
   ElseIf LCase(TxtPreFix) = "ttcrp" Then 'Kota P
      mNewSM_Prefix = "tcrin"
   End If
   
   If ChkSCMFlag.Value = 0 Then
      Call TableTxn_WHRDetail("Where  (sm_prefix = '" & TxtPreFix & "' Or sm_prefix = '" & mNewSM_Prefix & "') And WHRNo = " & mTxtRRNNo_ & " And SCMFlag = 0 ")
   Else
      Call TableTxn_WHRDetail("Where  (sm_prefix = '" & TxtPreFix & "' Or sm_prefix = '" & mNewSM_Prefix & "') And WHRNo = " & mTxtRRNNo_ & " And SCMFlag = 1 ")
   End If
End If

If Check1.Value = 0 Then
   If RsTxn_WHRDetail.RecordCount = 0 Then
      MsgBox "No record found!!!"
      TxtRRNNo.SetFocus
      Exit Sub
   End If
   RsTxn_WHRDetail.MoveFirst
   RsTxn_WHRDetail.Find "WHRNo = " & mTxtRRNNo_
   'If RsTxn_WHRDetail!Flag <> "P" And IsNull(RsTxn_WHRDetail!Flag) = False And RsTxn_WHRDetail!Flag <> "N" Then
   If IsNull(RsTxn_WHRDetail!Flag) = True Then
      MsgBox "WHR is not yet printed!!! "
      TxtRRNNo.SetFocus
      Exit Sub
   End If
   
   
   'If RsTxn_WHRDetail!PledgeID <> 0 Then
   '   MsgBox "This WHR is already pledged!!!"
   '   TxtRRNNo.SetFocus
   '   Exit Sub
   'End If
   mPledgeBags = 0
   mPledgeQty = 0
   
   If Not IsNull(RsTxn_WHRDetail!Txn_PledgeCreation_ID) Then
      tmp = "Select Sum(PledgeNoOfBags) PledgeBags,Sum(PledgeQuantity) PledgeQty from Txn_PledgeCreation_GSL Where  SM_Prefix = '" & TxtPreFix & "' And Txn_PledgeCreation_ID = " & RsTxn_WHRDetail!Txn_PledgeCreation_ID & ""
      Call Tmp6Table
        
      If TmpRs6.RecordCount > 0 Then
         mPledgeBags = IIf(IsNull(TmpRs6!PledgeBags), 0, TmpRs6!PledgeBags)
         mPledgeQty = IIf(IsNull(TmpRs6!PledgeQty), 0, TmpRs6!PledgeQty)
      End If
   End If
   If Not RsTxn_WHRDetail.EOF Then
      'If LCase(UserType) <> "power" And LCase(UserType) <> "superadmin" Then
         If mCMPID_ <> RsTxn_WHRDetail!CMPID Then
            MsgBox "Warning : WHR CMP is different from the CMP selected above!!!"
            'CmbCMPID.SetFocus
            'Exit Sub
         End If
     ' End If
      
      If mCommodityID_ <> RsTxn_WHRDetail!CommodityID Then
         MsgBox "WHR Commodity is different from the commodity selected above!!!"
         CmbCommodityID.SetFocus
         Exit Sub
      End If
      
      mRWCID = RsTxn_WHRDetail!WHRNo
      TxtRRNQty = IIf(IsNull(RsTxn_WHRDetail!WHRQuantity), "0", RsTxn_WHRDetail!WHRQuantity)
      TxtBalanceQty = IIf(IsNull(RsTxn_WHRDetail!WHRQuantity), "0", RsTxn_WHRDetail!WHRQuantity) - IIf(IsNull(RsTxn_WHRDetail!DeliveredQuantity), "0", RsTxn_WHRDetail!DeliveredQuantity) - mPledgeQty ' IIf(IsNull(RsTxn_WHRDetail!DeliveredQuantity), "0", RsTxn_WHRDetail!DeliveredQuantity)
      
      TxtRRNBags = IIf(IsNull(RsTxn_WHRDetail!WHRNoOfBags), "0", RsTxn_WHRDetail!WHRNoOfBags)
      TxtBalanceBags = Val(TxtRRNBags) - IIf(IsNull(RsTxn_WHRDetail!DeliveredNoOfBags), 0, RsTxn_WHRDetail!DeliveredNoOfBags) - mPledgeBags 'IIf(IsNull(RsTxn_WHRDetail!DeliveredNoOfBags), 0, RsTxn_WHRDetail!DeliveredNoOfBags)
      
      mClientIDRow = RsTxn_WHRDetail!ClientID
'      If RsTxn_WHRDetail!PledgeID <> 0 Then
'         MsgBox "ERROR : This Warehouse Receipt is already pledged!!!"
'      End If
     
      Call GetClientDetailsWHRAdd(mClientIDRow)
   End If
End If
End Sub
Private Sub CDCDetails(mCMPID_, mCommodityID_, mTxtRRNNo_)
If mExchangeTypeID = 0 Or IsNull(mExchangeTypeID) Then Exit Sub
If mCMPID_ = 0 Or IsNull(mCMPID_) Then Exit Sub
If mCommodityID_ = 0 Or IsNull(mCommodityID_) Then Exit Sub
If mTxtRRNNo_ = "" Then Exit Sub

Call TableTxn_CDC_GSLDetail("Where  sm_prefix = '" & TxtPreFix & "' And CDCNo = " & Val(mTxtRRNNo_))
If RsTxn_CDC_GSLDetail.RecordCount = 0 Then
   MsgBox "No such CDC record found!!!"
   If TxtRRNNo.Enabled = True Then
      TxtRRNNo.SetFocus
   End If
   Exit Sub
End If

If Check1.Value = 0 Then
   
   
   If RsTxn_CDC_GSLDetail!Flag <> "P" Then
      
      If RsTxn_CDC_GSLDetail!Flag = "F" Then
         MsgBox "Payout is not executed!!! "
         TxtRRNNo.SetFocus
         Exit Sub
      End If
      If RsTxn_CDC_GSLDetail!Flag = "T" Then
         MsgBox "This CDC is reissued.It can't be withdrawn!!! "
         TxtRRNNo.SetFocus
         Exit Sub
      End If
      If RsTxn_CDC_GSLDetail!Flag = "W" Then
         MsgBox "This CDC is already withdrawn!!! "
         TxtRRNNo.SetFocus
         Exit Sub
      End If
   End If
Else
   tmp = "SELECT Txn_CDC_Assayer_Detail.AcceptenceFlag " & _
   "FROM Txn_CDC_Assayer_Detail INNER JOIN Txn_CDC_GSL_Detail ON (Txn_CDC_Assayer_Detail.SM_Prefix = " & _
   "Txn_CDC_GSL_Detail.SM_Prefix) AND (Txn_CDC_Assayer_Detail.CDCID = Txn_CDC_GSL_Detail.CDCID) " & _
   "WHERE (((Txn_CDC_GSL_Detail.CDCNo)=" & mTxtRRNNo_ & ") AND ((Txn_CDC_GSL_Detail.SM_Prefix)= '" & TxtPreFix & "'));"
   Call Tmp1Table
   If TmpRs1.RecordCount > 0 Then
      If TmpRs1.Fields(0) <> 0 Then
         MsgBox "Selected CDC is not rejected. You can't withdraw this CDC!!! "
         TxtRRNNo.SetFocus
         Exit Sub
      End If
   End If
End If

mCDCID = RsTxn_CDC_GSLDetail!CDCID
Call TableTxn_CDC("Where sm_prefix = '" & TxtPreFix & "' And  CDCID =  " & mCDCID & " ")
If Not RsTxn_CDC.EOF Then
   If mCMPID_ <> RsTxn_CDC!CMPID Then
      MsgBox "CDC CMP is different from the CMP selected above!!!"
      CmbCMPID.SetFocus
      Exit Sub
   End If
   If RsTxn_CDC!PrintFlag <> "P" Then
      MsgBox "CDC is not printed!!!"
      TxtRRNNo.SetFocus
      Exit Sub
   End If
   
   If mCommodityID_ <> RsTxn_CDC!CommodityID Then
      MsgBox "CDC Commodity is different from the commodity selected above!!!"
      CmbCommodityID.SetFocus
      Exit Sub
   End If
End If

mRWCID = RsTxn_CDC_GSLDetail!CDCNo
mGSLID = RsTxn_CDC_GSLDetail!Mst_GSL_ID
mClientIDRow = RsTxn_CDC_GSLDetail!ClientID



'Code by Parag for Checking Flag of CDC GSL 31 May 2010
' Copy this code in CDCDetails procedure below line "mClientIDRow = RsTxn_CDC_GSLDetail!ClientIDRow" and above line "Call TableMst_GSL(" Where sm_prefix = '" & TxtPreFix & "' And  GSLID = " & mGSLID)"
Call GetGSLTransferDetail(TxtPreFix, mGSLID)

Call TableMst_GSL(" Where sm_prefix = '" & TxtPreFix & "' And  Mst_GSL_ID = " & mGSLID)
TxtRRNQty = RsMst_GSL!DepositeWeight
TxtBalanceQty = RsMst_GSL!DepositeWeight - RsMst_GSL!CDTFWeight


TxtRRNBags = RsMst_GSL!DepositeBags
TxtBalanceBags = RsMst_GSL!DepositeBags - RsMst_GSL!CDTFBags


Call GetClientDetails(mClientIDRow)
'Call DisplayGSLData(mGSLID)

CmbGCTDClient.Text = RsTxn_CDC_GSLDetail!CDCNo & " ~ " & TxtClientName


'Condition removed on 29-01-2010 mayur

'If mExchangeTypeID = 3 And Check2.Value = 0 Then
'   If CheckISINinRRNTab() = True Then
'      MsgBox "This lot belongs to one client but the selected CDC belongs to a different client!!!"
'      If CmbLot.Enabled = True Then
'         CmbLot.SetFocus
'      End If
'      Exit Sub
'   End If
'End If

End Sub

Private Sub GetGSLTransferDetail(SM_Prefix_ As Variant, gslID_ As Variant)
tmp = "Select Mst_GSL_ID,Status,FGSLID_Mst_GSL_ID From Mst_GSL Where SM_Prefix = '" & SM_Prefix_ & "' And Mst_GSL_ID = " & gslID_ & " "
Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   If TmpRs1!Status = "Z" Then
      mGSLID = TmpRs1!FGSLID_Mst_GSL_ID
      Call GetGSLTransferDetail(SM_Prefix_, mGSLID)
   ElseIf TmpRs1!Status = "O" Then
      mGSLID = TmpRs1!Mst_GSL_ID
      Exit Sub
   End If
End If

End Sub

Private Sub GetClientDetails(mClientIDRow_ As Double)
If mClientIDRow_ = 0 Or IsNull(mClientIDRow_) Then Exit Sub
ans = GetClientNameByRow(mClientIDRow_)
If ans(0) = "" Then
   MsgBox "Not found!!!"
   Exit Sub
End If
TxtClientID = ans(0)
TxtClientName = ans(1)
TxtDPID = ans(2)
TxtDPName = ans(3)
'mNAVCode = ans(6)
'mRNAVCode = ans(7)

End Sub
Private Sub DisplayGSLData(mGSLID_)
Dim r As Variant
r = MSHFlexGrid4.Rows - 1
MSHFlexGrid4.Rows = MSHFlexGrid4.Rows + 1
Call TableMst_GSL("Where sm_prefix = '" & TxtPreFix & "' And Mst_GSL_ID = " & mGSLID_ & " ")
If RsMst_GSL.RecordCount = 0 Then
   'MsgBox "No GSL Attach To This CDC!!!"
   Exit Sub
End If


MSHFlexGrid4.TextMatrix(r, 1) = GetGodownName(RsMst_GSL!GodownID, "N")
MSHFlexGrid4.TextMatrix(r, 2) = RsMst_GSL!StackNo
MSHFlexGrid4.TextMatrix(r, 3) = RsMst_GSL!LotNo
MSHFlexGrid4.TextMatrix(r, 4) = IIf(IsNull(RsMst_GSL!DepositeBags), 0, RsMst_GSL!DepositeBags) - IIf(IsNull(RsMst_GSL!CDTFBags), 0, RsMst_GSL!CDTFBags)
MSHFlexGrid4.TextMatrix(r, 5) = IIf(IsNull(RsMst_GSL!DepositeWeight), 0, RsMst_GSL!DepositeWeight) - IIf(IsNull(RsMst_GSL!CDTFWeight), 0, RsMst_GSL!CDTFWeight)
MSHFlexGrid4.TextMatrix(r, 8) = Val(TxtAdjWt)
MSHFlexGrid4.TextMatrix(r, 9) = Format((Val(MSHFlexGrid4.TextMatrix(r, 4)) * Val(TxtCGunnyWt)) / 1000, "#######0.000")
MSHFlexGrid4.TextMatrix(r, 10) = IIf(IsNull(RsMst_GSL!DepositeWeight), 0, RsMst_GSL!DepositeWeight) - IIf(IsNull(RsMst_GSL!CDTFWeight), 0, RsMst_GSL!CDTFWeight) 'Val(TxtGSLNetWeight)
MSHFlexGrid4.TextMatrix(r, 6) = IIf(IsNull(RsMst_GSL!DepositeBags), 0, RsMst_GSL!DepositeBags) - IIf(IsNull(RsMst_GSL!CDTFBags), 0, RsMst_GSL!CDTFBags)
MSHFlexGrid4.TextMatrix(r, 7) = Val(MSHFlexGrid4.TextMatrix(r, 10)) - Val(MSHFlexGrid4.TextMatrix(r, 8)) + Val(MSHFlexGrid4.TextMatrix(r, 9))   'IIf(IsNull(RsMst_GSL!DepositeWeight), 0, RsMst_GSL!DepositeWeight) - IIf(IsNull(RsMst_GSL!CDTFWeight), 0, RsMst_GSL!CDTFWeight)
MSHFlexGrid4.TextMatrix(r, 11) = Val(TxtGSLWeightGainLoss)
MSHFlexGrid4.TextMatrix(r, 12) = TxtRRNNo & " ~ " & TxtClientName
MSHFlexGrid4.TextMatrix(r, 13) = TxtGCTDremarks
MSHFlexGrid4.TextMatrix(r, 14) = mGSLID_
'MSHFlexGrid4.TextMatrix(r, 15) = mGodownID
'MSHFlexGrid4.TextMatrix(r, 16) = mRWCID
 
MSHFlexGrid4.TextMatrix(0, 4) = Val(MSHFlexGrid4.TextMatrix(0, 4)) + Val(MSHFlexGrid4.TextMatrix(r, 4))
MSHFlexGrid4.TextMatrix(0, 5) = Val(MSHFlexGrid4.TextMatrix(0, 5)) + Val(MSHFlexGrid4.TextMatrix(r, 5))
MSHFlexGrid4.TextMatrix(0, 6) = Val(MSHFlexGrid4.TextMatrix(0, 6)) + Val(MSHFlexGrid4.TextMatrix(r, 6))
MSHFlexGrid4.TextMatrix(0, 7) = Val(MSHFlexGrid4.TextMatrix(0, 7)) + Val(MSHFlexGrid4.TextMatrix(r, 7))
MSHFlexGrid4.TextMatrix(0, 8) = Val(MSHFlexGrid4.TextMatrix(0, 8)) + Val(MSHFlexGrid4.TextMatrix(r, 8))
MSHFlexGrid4.TextMatrix(0, 9) = Val(MSHFlexGrid4.TextMatrix(0, 9)) + Val(MSHFlexGrid4.TextMatrix(r, 9))
MSHFlexGrid4.TextMatrix(0, 10) = Val(MSHFlexGrid4.TextMatrix(0, 10)) + Val(MSHFlexGrid4.TextMatrix(r, 10))
MSHFlexGrid4.TextMatrix(0, 11) = Val(MSHFlexGrid4.TextMatrix(0, 11)) + Val(MSHFlexGrid4.TextMatrix(r, 11))

'MSHFlexGrid4.ColWidth(12) = 0
Call Clear_GCTDTabControl
CmdGCTDAdd2List.Enabled = False
End Sub




Private Sub TxtSpotPrice_LostFocus()
'TxtSpotPrice.Locked = False
If TxtSpotPrice = "" Then Exit Sub
If Not IsNumeric(TxtSpotPrice) Or Val(TxtSpotPrice) < 0 Then
   MsgBox "Invalid numeric format!!!"
   TxtSpotPrice.SetFocus
   Exit Sub
End If

TxtSpotPrice = Format(TxtSpotPrice, "#######0.00")


End Sub

Private Sub TxtTTDTareWt_LostFocus()
If TxtTTDTareWt = "" Then Exit Sub
If Not IsNumeric(TxtTTDTareWt) Then
   MsgBox "Invalid numeric format!!!"
   TxtTTDTareWt.SetFocus
End If
Call CalTruckWt
End Sub

Private Sub TxtTTDGrossWt_LostFocus()
If TxtTTDGrossWt = "" Then Exit Sub
If Not IsNumeric(TxtTTDGrossWt) Then
   MsgBox "Invalid numeric format!!!"
   TxtTTDGrossWt.SetFocus
End If

End Sub

Private Sub TxtTTDTotalWt_LostFocus()
If TxtTTDTotalWt = "" Then Exit Sub
If Not IsNumeric(TxtTTDTotalWt) Then
   MsgBox "Invalid numeric format!!!"
   TxtTTDTotalWt.SetFocus
End If
Call CalTruckWt
End Sub

Private Sub CalTruckWt()

TxtTTDGrossWt = Val(TxtTTDTotalWt) - Val(TxtTTDTareWt)
TxtTTDGrossWt = Format(TxtTTDGrossWt, "#######0.000")

End Sub
Private Sub Clear_RWCTabControl()
TxtRRNNo = ""
TxtISINno = ""
TxtRRNQty = ""
TxtBalanceQty = ""
TxtAdjustedQty = ""
TxtRRNRemarks = ""
TxtClientID = ""
TxtClientName = ""
TxtDPID = ""
TxtDPName = ""
Check2.Value = 0
mNAVCode = ""
mRNAVCode = ""
mRWCID = 0
TxtBalanceBags = ""
CmdAdd2ListRRN.Caption = "Add to List"
'Call Grid_Head_RRN
End Sub
Private Sub Clear_GCTDTabControl()
CmbGodownNo.Text = ""
mGodownYearlyFlag = ""
CmbStack.Text = ""
CmbLot.Text = ""
'CmbGCTDClient.Text = ""
TxtGCTDNoofBags = ""
TxtGCTDGrossWt = ""
TxtGGunnyWeight = ""
TxtGSLISINno = ""
TxtGSLNetWeight = ""
TxtGCTDremarks = ""
TxtAdjWt = ""
'TxtAdjWt.Locked = True
TxtGSLAvailableNofBags = ""
TxtGSLAvailableGrossWgt = ""
CmbGCTDClient.Text = ""
TxtGSLWeightGainLoss = ""
'CmbNonNcdexClient.Text = ""
CmbGSLTruckNo.Text = ""
mGSLID = 0
TxtMoisturePer1.Text = ""
CmdGCTDAdd2List.Caption = "Add to List"
Check2.Enabled = False
Check4.Enabled = False
Call EnDnGCTDTab(True)
If mExchangeTypeID = 3 Then
   CmbGodownNo.Enabled = False
   CmbStack.Enabled = False
   CmbLot.Enabled = False
End If

If mExchangeTypeID = 1 Then
   TxtAdjWt.Locked = True
Else
   If Check1.Value = 1 Then
      TxtAdjWt.Locked = True
   Else
      TxtAdjWt.Locked = False
   End If
End If
Check4.Value = 0
Check2.Value = 0
CmbGCTDClient.Enabled = True
mPledgeID = ""
TxtAverageBagSize = ""

End Sub
Private Sub EnDnGCTDTab(Yn As Boolean)
Dim Ny As Boolean
Ny = IIf(Yn, False, True)

CmdGCTDAdd2List.Enabled = Yn
CmbGodownNo.Enabled = Yn
CmbStack.Enabled = Yn
CmbLot.Enabled = Yn
Check2.Enabled = Yn
Check4.Enabled = Yn
TxtGCTDNoofBags.Locked = Ny
TxtGCTDGrossWt.Locked = Ny

TxtAdjWt.Locked = Ny  '16/03/2010
TxtGCTDremarks.Locked = Ny

End Sub
Public Sub Indata()
Frame5.Visible = False
BalBagNeg = "N"
PicRRN.Enabled = False
Pic_GCTD.Enabled = False
Pic_TTD.Enabled = False
FrmInvoice.Visible = False
PntPrv = "N"
Edit_Flg = ""
Call Grid_Head_TTD
Call Grid_Head_RRN
Call Grid_Head_GCTD
Call Grid_Head_GSLBalChk
Call Grid_Head_Stock
Call Grid_Head_OutStand

CmdPrintPreview.Enabled = True
Frame1.Visible = False
FrameCRDetail.Visible = False
TxtCxTFNo.Locked = True
TxtCxTFDate.Enabled = False
TxtCMPID.Locked = True
CmbCMPID.Visible = False
TxtCMPID.Visible = True
TxtCommodityID.Locked = True
CmbCommodityID.Visible = False
TxtCommodityID.Visible = True
CmbExchangeTypeID.Enabled = False
TxtAgent.Locked = True
Check1.Enabled = False

TxtAgent.Locked = True
CmbAgent.Visible = False
CmdAgent.Visible = False
TxtAgent.Visible = True
TxtManualCDTFNo.Locked = True
TxtSpotPrice.Locked = True

TxtCxTFNo = IIf(IsNull(RsTxn_CxTF_Details!CxTF_DetailsID), "", RsTxn_CxTF_Details!CxTF_DetailsID)
TxtManualCDTFNo = IIf(IsNull(RsTxn_CxTF_Details!ManualCDTFNo), "", RsTxn_CxTF_Details!ManualCDTFNo)

TxtPreFix = IIf(IsNull(RsTxn_CxTF_Details!SM_Prefix), "", RsTxn_CxTF_Details!SM_Prefix)
TxtCxTFDate.Value = IIf(IsNull(RsTxn_CxTF_Details!CxTFDate), "", RsTxn_CxTF_Details!CxTFDate)
TxtCxTFDispalyDate = TxtCxTFDate.Value
TxtCMPID = GetCMPName(IIf(IsNull(RsTxn_CxTF_Details!CMPID), "", RsTxn_CxTF_Details!CMPID))
mCMPID = IIf(IsNull(RsTxn_CxTF_Details!CMPID), 0, RsTxn_CxTF_Details!CMPID)

mLocationID = GetLocationIDFromCMP(mCMPID)

TxtCommodityID = GetCommodityName(IIf(IsNull(RsTxn_CxTF_Details!CommodityID), "", RsTxn_CxTF_Details!CommodityID))
mCommodityID = IIf(IsNull(RsTxn_CxTF_Details!CommodityID), "", RsTxn_CxTF_Details!CommodityID)
CmbExchangeTypeID.Text = GetExchangeName(IIf(IsNull(RsTxn_CxTF_Details!ExchangeTypeID), "", RsTxn_CxTF_Details!ExchangeTypeID))
mExchangeTypeID = IIf(IsNull(RsTxn_CxTF_Details!ExchangeTypeID), "", RsTxn_CxTF_Details!ExchangeTypeID)
Check1.Value = IIf(IsNull(RsTxn_CxTF_Details!BookedSpaceFlag), "", RsTxn_CxTF_Details!BookedSpaceFlag)
TxtFlag = IIf(IsNull(RsTxn_CxTF_Details!Flag), "", RsTxn_CxTF_Details!Flag)
TxtAgent = IIf(IsNull(RsTxn_CxTF_Details!Agent), "", RsTxn_CxTF_Details!Agent)
ChkSCMFlag.Enabled = False
ChkSCMFlag.Value = IIf(IsNull(RsTxn_CxTF_Details!SCMFlag), 0, RsTxn_CxTF_Details!SCMFlag)

If Not IsNull(RsTxn_CxTF_Details!AgentID) Then
   mAgentID = IIf(IsNull(RsTxn_CxTF_Details!AgentID), "", RsTxn_CxTF_Details!AgentID)
   TxtAgent = GetAgentName(IIf(IsNull(RsTxn_CxTF_Details!AgentID), "", RsTxn_CxTF_Details!AgentID))
End If

TxtRemark = IIf(IsNull(RsTxn_CxTF_Details!Remark), "", RsTxn_CxTF_Details!Remark)
TxtTxnType = IIf(IsNull(RsTxn_CxTF_Details!TxnType), "", RsTxn_CxTF_Details!TxnType)
'TxtContractID = GetContract(IIf(IsNull(RsTxn_CxTF_Details!ContractID), "", RsTxn_CxTF_Details!ContractID))
mContractID = IIf(IsNull(RsTxn_CxTF_Details!ContractID), "", RsTxn_CxTF_Details!ContractID)
CmbCMPID.Text = TxtCMPID

'Call CmbExchangeTypeID_GotFocus
'Call CmbExchangeTypeID_LostFocus
'Call CmbCMPID_GotFocus
'Call CmbCMPID_LostFocus
CmbCommodityID.Text = TxtCommodityID
TxtSpotPrice = IIf(IsNull(RsTxn_CxTF_Details!SpotPrice), "", RsTxn_CxTF_Details!SpotPrice)
TxtCreated = IIf(IsNull(RsTxn_CxTF_Details!CreatedBy), "", RsTxn_CxTF_Details!CreatedBy) & " :- " & IIf(IsNull(RsTxn_CxTF_Details!CreatedDate), "", RsTxn_CxTF_Details!CreatedDate)
TxtUpdated = IIf(IsNull(RsTxn_CxTF_Details!UpdatedBy), "", RsTxn_CxTF_Details!UpdatedBy) & " :- " & IIf(IsNull(RsTxn_CxTF_Details!UpdatedDate), "", RsTxn_CxTF_Details!UpdatedDate)

TxtGUID = IIf(IsNull(RsTxn_CxTF_Details!G_UID), "", RsTxn_CxTF_Details!G_UID)

Call CmbCommodityID_GotFocus
Call CmbCommodityID_LostFocus
If mExchangeTypeID = 1 Then 'NON NCDEX
   'Call SetCaptions("WHR No", "WHR No", "ClientIDRow", "WHR Qty", "Depository ID", "Depository Name", False)
'Client Name
   Call SetCaptions("WHR No", "WHR No", "Client Name", "WHR Qty", "Depository ID", "Depository Name", False)
End If
If mExchangeTypeID = 2 Then 'NCDEX
   If Check1.Value = 1 Then
      Call SetCaptions("client details", "CDTF No", "Client Name", "Rejected Qty", "DP ID", "DP Name", True)
   Else
      Call SetCaptions("Remat Request No. (RRN)", "RRN No", "CMSE Lot No", "RRN Qty", "DP ID", "DP Name", True)
   End If
End If
If mExchangeTypeID = 3 Then 'SPOT
   Call SetCaptions("CDC No", "CDC-No", "CDC-No", "CDC Qty", "TM ID", "Trading Member Name", True)
End If
Call InDataClient
Call InDataTruck
Call InDataGSL
If RsTxn_CxTF_Details.RecordCount = 1 Then
   Call GButtonLock_1(Me, True)
Else
   Call GButtonLock(Me, True)
End If
Frame3.Visible = False
EditCmd.Enabled = False
'DeleteCmd.Enabled = False

'
'If LCase(CmbExchangeTypeID.Text) = "non ncdex" Then
'   CmdBOE.Enabled = True
'Else
'   CmdBOE.Enabled = False
'End If
'

If LCase(CmbExchangeTypeID.Text) = "non ncdex" Or (LCase(CmbExchangeTypeID.Text) = "ncdex" And Check1.Value = 1) Then
   CmdBOE.Enabled = True
Else
   CmdBOE.Enabled = False
End If

CmdPrintPreview.Enabled = True
''If RsTxn_CxTF_Details!HologramNo <> "" Or IsNull(RsTxn_CxTF_Details!HologramNo) = False Then
''   'MsgBox "Print out already generated for this CDTF . Can't preview!!!"
''   'Exit Sub
''   CmdPrintPreview.Enabled = False
''End If

If IIf(IsNull(RsTxn_CxTF_Details!HologramNo), "", RsTxn_CxTF_Details!HologramNo) <> "" Then
   'MsgBox "Print out already generated for this CDTF . Can't preview!!!"
   'Exit Sub
   CmdPrintPreview.Enabled = False
End If

CmdAPrint.Enabled = False
TxtBOEAmount = GetBOEAmount()

If LCase(CmbExchangeTypeID.Text) = "non ncdex" Then
   Call AssignMultOf
   Call GetClientOutstanding
   Call FillGSLBalanceStockWithPrice
   If CheckForBOEAmount() = True And Val(TxtBOEAmount.Text) = 0 Then
   
   Else
      If RsTxn_CxTF_Details!HologramNo = "" Or IsNull(RsTxn_CxTF_Details!HologramNo) Then
         mTotBalQty = Val(MSHFlexGrid8.TextMatrix(0, 4))
         mTotCommPrice = Val(MSHFlexGrid8.TextMatrix(0, 6))
         mTotOutstAmt = Val(MSHFlexGrid9.TextMatrix(0, 12)) + Val(TxtBOEAmount.Text)
         If mTotBalQty > 0 Then
            If mTotCommPrice >= (mTotOutstAmt * mMultOfAll) Then
               If IsNull(RsTxn_CxTF_Details!ApprovePrintFlag) = True Or RsTxn_CxTF_Details!ApprovePrintFlag = "N" Then
                  RsTxn_CxTF_Details!ApprovePrintFlag = "Y"
                  RsTxn_CxTF_Details!ApprovePrintBy = "System"
                  RsTxn_CxTF_Details!ApprovePrintRemark = "System Approved(Stock value (" & mTotCommPrice & ") is more than " & mMultOfAll & " time of Outstanding (" & mTotOutstAmt & "))"
                  RsTxn_CxTF_Details!ApprovePrintDate = Now
                  RsTxn_CxTF_Details!OutstandingAmt = mTotOutstAmt
                  RsTxn_CxTF_Details!CommodityValue = mTotCommPrice
                  RsTxn_CxTF_Details!BalQty = mTotBalQty
                  RsTxn_CxTF_Details.Update
                  RsTxn_CxTF_Details.Requery
               End If
            End If
         Else
            If mTotOutstAmt = 0 Then
               If IsNull(RsTxn_CxTF_Details!ApprovePrintFlag) = True Or RsTxn_CxTF_Details!ApprovePrintFlag = "N" Then
                  RsTxn_CxTF_Details!ApprovePrintFlag = "Y"
                  RsTxn_CxTF_Details!ApprovePrintBy = "System"
                  RsTxn_CxTF_Details!ApprovePrintRemark = "System Approved(Stock & Outstanding is zero)"
                  RsTxn_CxTF_Details!ApprovePrintDate = Now
                  RsTxn_CxTF_Details!OutstandingAmt = mTotOutstAmt
                  RsTxn_CxTF_Details!CommodityValue = mTotCommPrice
                  RsTxn_CxTF_Details!BalQty = mTotBalQty
                  RsTxn_CxTF_Details.Update
                  RsTxn_CxTF_Details.Requery
               End If
            End If
         End If
      End If
      RsTxn_CxTF_Details.MoveFirst
      RsTxn_CxTF_Details.Find "G_UID='" & TxtGUID & "'"
   End If
   If IsNull(RsTxn_CxTF_Details!ApprovePrintFlag) = True Or RsTxn_CxTF_Details!ApprovePrintFlag = "N" Then
      CmdAPrint.Enabled = True
   End If
End If

CmdPrint.Enabled = True

End Sub
Private Sub InDataGSL()

Call TableTxn_CxTF_GSL("  Where  sm_prefix = '" & TxtPreFix & "' And CxTF_DetailsID = " & TxtCxTFNo & " And TxnType = 'W'")
Call TableMst_Godown(" Where CMPID = " & mCMPID & " ")

MSHFlexGrid4.Rows = MSHFlexGrid4.Rows + RsTxn_CxTF_GSL.RecordCount
If RsTxn_CxTF_GSL.RecordCount = 0 Then Exit Sub

Dim ans As Variant

For I = 2 To MSHFlexGrid4.Rows - 2
    
    
    MSHFlexGrid4.Col = 7
    MSHFlexGrid4.row = I
    MSHFlexGrid4.CellBackColor = &HFFC0C0


    ans = GetGSLByGLSID(RsTxn_CxTF_GSL!Mst_GSL_ID, TxtPreFix)
    MSHFlexGrid4.TextMatrix(I, 0) = RsTxn_CxTF_GSL!GSLDetailID
    
    MSHFlexGrid4.TextMatrix(I, 1) = ans(3) ' TmpRs!GodownNo
    MSHFlexGrid4.TextMatrix(I, 2) = ans(4) 'TmpRs!StackNo
    MSHFlexGrid4.TextMatrix(I, 3) = ans(5) 'TmpRs!LotNo
    If RsMst_Godown.RecordCount > 0 Then
       RsMst_Godown.MoveFirst
       RsMst_Godown.Find "GodownID =  " & Val(ans(2)) & ""
       If RsMst_Godown.EOF = False Then
          MSHFlexGrid4.TextMatrix(I, 22) = IIf(IsNull(RsMst_Godown!YearlyBillingFlg), "", RsMst_Godown!YearlyBillingFlg)
       End If
    End If
    MSHFlexGrid4.TextMatrix(I, 18) = ans(15) 'TmpRs!BalanceBags
    If ans(15) < 0 Then
       BalBagNeg = "Y"
    End If
    If MSHFlexGrid4.TextMatrix(1, 12) = "Client Name" And Not (IsNull(RsTxn_CxTF_GSL!ClientID)) Then 'ans(9) <> "" Then
       If IsNull(RsTxn_CxTF_GSL!WHRNo) = False Then
          mWHR_ = Format((RsTxn_CxTF_GSL!WHRNo / 1000000), "#####")
          mWHR_ = mWHR_ & "-" & (RsTxn_CxTF_GSL!WHRNo - (mWHR_ * 1000000))
    
          MSHFlexGrid4.TextMatrix(I, 12) = mWHR_ & "~" & GetClientName(RsTxn_CxTF_GSL!ClientID) 'GetClientName(IIf(IsNull(ans(9)), 0, ans(9)))
       Else
          MSHFlexGrid4.TextMatrix(I, 12) = GetClientName(RsTxn_CxTF_GSL!ClientID) 'GetClientName(IIf(IsNull(ans(9)), 0, ans(9)))
       End If
    Else
       If IsNull(RsTxn_CxTF_GSL!WHRNo) = False Then
          mWHR_ = GetRRNNo(RsTxn_CxTF_GSL!WHRNo)
          MSHFlexGrid4.TextMatrix(I, 12) = IIf(IsNull(ans(10)) Or ans(10) = 0, "", mWHR_ & "~" & ans(10))
       Else
          MSHFlexGrid4.TextMatrix(I, 12) = IIf(IsNull(ans(10)) Or ans(10) = 0, "", ans(10))
       End If
       
    End If
    MSHFlexGrid4.TextMatrix(I, 15) = IIf(IsNull(ans(6)), "", ans(6))
    MSHFlexGrid4.TextMatrix(I, 6) = RsTxn_CxTF_GSL!NoofBags
    MSHFlexGrid4.TextMatrix(I, 7) = RsTxn_CxTF_GSL!GrossWeight
    MSHFlexGrid4.TextMatrix(I, 8) = RsTxn_CxTF_GSL!AdjNCMSLWeight
    MSHFlexGrid4.TextMatrix(I, 9) = RsTxn_CxTF_GSL!GunnyWeight
    MSHFlexGrid4.TextMatrix(I, 10) = RsTxn_CxTF_GSL!DematWeight
    'MSHFlexGrid4.TextMatrix(i, 11) = Weight Gain/LossIif(Isnull(RsTxn_CxTF_GSL!DematWeight),0,RsTxn_CxTF_GSL!De
    MSHFlexGrid4.TextMatrix(I, 13) = IIf(IsNull(RsTxn_CxTF_GSL!Remarks), "", RsTxn_CxTF_GSL!Remarks)
    MSHFlexGrid4.TextMatrix(I, 21) = IIf(IsNull(RsTxn_CxTF_GSL!MoisturePer), "", RsTxn_CxTF_GSL!MoisturePer)
    MSHFlexGrid4.TextMatrix(I, 23) = Format(Val(MSHFlexGrid4.TextMatrix(I, 10)) * 1000 / Val(MSHFlexGrid4.TextMatrix(I, 6)), "#######0.00")
    
    ''''yellow color indication for reservation  lot
    If ans(17) > 0 Then
       For m = 0 To MSHFlexGrid4.Cols - 1
          MSHFlexGrid4.row = I
          MSHFlexGrid4.Col = m
          MSHFlexGrid4.CellBackColor = vbYellow
       Next
    End If
    
    If Not (IsNull(RsTxn_CxTF_GSL!ClientID)) Then
       ans = GetClientNameByRow(RsTxn_CxTF_GSL!ClientID)
       MSHFlexGrid4.TextMatrix(I, 14) = RsTxn_CxTF_GSL!Mst_GSL_ID
       MSHFlexGrid4.TextMatrix(I, 16) = RsTxn_CxTF_GSL!ClientID
       'MSHFlexGrid4.TextMatrix(i, 17) = RRNID
    Else
       MSHFlexGrid4.TextMatrix(I, 14) = RsTxn_CxTF_GSL!Mst_GSL_ID
    End If
    
    MSHFlexGrid4.Col = 7
    MSHFlexGrid4.row = I
    MSHFlexGrid4.CellBackColor = &HFFC0C0



    MSHFlexGrid4.TextMatrix(0, 4) = Val(MSHFlexGrid4.TextMatrix(0, 4)) + Val(MSHFlexGrid4.TextMatrix(I, 4))
    MSHFlexGrid4.TextMatrix(0, 5) = Val(MSHFlexGrid4.TextMatrix(0, 5)) + Val(MSHFlexGrid4.TextMatrix(I, 5))
    MSHFlexGrid4.TextMatrix(0, 6) = Val(MSHFlexGrid4.TextMatrix(0, 6)) + Val(MSHFlexGrid4.TextMatrix(I, 6))
    MSHFlexGrid4.TextMatrix(0, 7) = Val(MSHFlexGrid4.TextMatrix(0, 7)) + Val(MSHFlexGrid4.TextMatrix(I, 7))
    MSHFlexGrid4.TextMatrix(0, 8) = Val(MSHFlexGrid4.TextMatrix(0, 8)) + Val(MSHFlexGrid4.TextMatrix(I, 8))
    MSHFlexGrid4.TextMatrix(0, 9) = Val(MSHFlexGrid4.TextMatrix(0, 9)) + Val(MSHFlexGrid4.TextMatrix(I, 9))
    MSHFlexGrid4.TextMatrix(0, 10) = Val(MSHFlexGrid4.TextMatrix(0, 10)) + Val(MSHFlexGrid4.TextMatrix(I, 10))
    MSHFlexGrid4.TextMatrix(0, 11) = Val(MSHFlexGrid4.TextMatrix(0, 11)) + Val(MSHFlexGrid4.TextMatrix(I, 11))
    
    If mExchangeTypeID = 3 Then
       MSHFlexGrid4.TextMatrix(I, 12) = "" 'IIf(IsNull(Ans(10)) Or Ans(10) = 0, "", Ans(10))
       MSHFlexGrid4.ColWidth(12) = 0
    End If
    MSHFlexGrid4.TextMatrix(I, 19) = IIf(IsNull(RsTxn_CxTF_GSL!TruckNo), "", RsTxn_CxTF_GSL!TruckNo) & "~" & IIf(IsNull(RsTxn_CxTF_GSL!WeighBridgeSlipno), "", RsTxn_CxTF_GSL!WeighBridgeSlipno)
    
       
    RsTxn_CxTF_GSL.MoveNext
Next

End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsTxn_CxTF_Details.MoveNext
If RsTxn_CxTF_Details.EOF Then
   RsTxn_CxTF_Details.MoveLast
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
RsTxn_CxTF_Details.MovePrevious
If RsTxn_CxTF_Details.BOF Then
   RsTxn_CxTF_Details.MoveFirst
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
RsTxn_CxTF_Details.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsTxn_CxTF_Details.MoveLast
Call Indata
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False
End Sub
Private Sub CmbGCTDClient_GotFocus()
tmp = CmbGCTDClient.Text
CmbGCTDClient.Clear
If mExchangeTypeID = 2 And MSHFlexGrid10.Rows > 3 Then
   If Check2.Value = 1 Then
      If MSHFlexGrid2.Rows < 4 Then Exit Sub
      For I = 2 To MSHFlexGrid2.Rows - 2
          CmbGCTDClient.AddItem MSHFlexGrid2.TextMatrix(I, 1) & " ~ " & MSHFlexGrid2.TextMatrix(I, 2)
      Next I
   Else
      CmbGCTDClient.AddItem GetHeaderISINRRN(TxtGSLISINno)
   End If
Else
   If MSHFlexGrid2.Rows < 4 Then Exit Sub

   For I = 2 To MSHFlexGrid2.Rows - 2
       CmbGCTDClient.AddItem MSHFlexGrid2.TextMatrix(I, 1) & " ~ " & MSHFlexGrid2.TextMatrix(I, 4)
   Next I
End If
CmbGCTDClient.Text = tmp
End Sub
Private Sub CmbGCTDClient_LostFocus()
If CmbGCTDClient.Text = "" Then Exit Sub
If mExchangeTypeID = 2 And MSHFlexGrid10.Rows > 3 Then
   If Check2.Value = 1 Then
      For I = 2 To MSHFlexGrid2.Rows - 2
          If CmbGCTDClient.Text = MSHFlexGrid2.TextMatrix(I, 1) & " ~ " & MSHFlexGrid2.TextMatrix(I, 2) Then
             mGSLClientIDRow = MSHFlexGrid2.TextMatrix(I, 11)
             Exit Sub
          End If
      Next I
   Else
      
   End If
   
Else
   For I = 2 To MSHFlexGrid2.Rows - 2
       If CmbGCTDClient.Text = MSHFlexGrid2.TextMatrix(I, 1) & " ~ " & MSHFlexGrid2.TextMatrix(I, 4) Then
          mGSLClientIDRow = MSHFlexGrid2.TextMatrix(I, 11)
          Exit Sub
       End If
   Next I
End If
MsgBox "Invalid selection!!!"
CmbGCTDClient.SetFocus

End Sub
Function CheckDuplicateClient() As Boolean
Dim Retval As Boolean
Retval = False
If MSHFlexGrid2.Rows > 3 Then
   For u = 1 To MSHFlexGrid2.Rows - 1
       If MSHFlexGrid2.TextMatrix(u, 3) = TxtClientID And MSHFlexGrid2.TextMatrix(u, 5) = TxtDPID Then
          If LCase(CmdAdd2ListRRN.Caption) = "add to list" Then
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

CheckDuplicateClient = Retval
End Function

Private Sub TxtTTDTruckNo_KeyPress(KeyAscii As Integer)
tmp = KeyAscii
If KeyAscii = 14 Then
   SSTab1.Tab = 2
   TxtTTDTruckNo.SetFocus
End If
End Sub
Private Sub CmdPrint_Click()

If BalBagNeg = "Y" Then
   MsgBox "Duplicate RRN/WHR or One or more GSL having Negative 'No of Bags' please contact CMG/IT for assistance!!!"
   Exit Sub
End If
If Val(MSHFlexGrid3.TextMatrix(0, 5)) <> Val(MSHFlexGrid4.TextMatrix(0, 7)) Then
   MsgBox "Gross Weight in 'Truck Details' tab is different from Gross Weight in 'GSL Details' tab!!! "
   Exit Sub
End If

If Val(MSHFlexGrid2.TextMatrix(0, 9)) <> Val(MSHFlexGrid4.TextMatrix(0, 10)) Then
   If mExchangeTypeID = 1 Then
      MsgBox "Delivered Qty in 'WHR No' Tab and Electronic Weight in 'GSL Details' tab are not equal !!!!"
   ElseIf mExchangeTypeID = 2 Then
      MsgBox "Delivered Qty in 'RRN Detail' Tab and Electronic Weight in 'GSL Details' tab are not equal !!!!"
   Else
      MsgBox "Delivered Qty in 'CDC Detail' Tab and Electronic Weight in 'GSL Detail' tab are not equal !!!!"
   End If
   Exit Sub
End If

If Val(MSHFlexGrid2.TextMatrix(0, 15)) <> Val(MSHFlexGrid4.TextMatrix(0, 6)) Then
   If mExchangeTypeID = 1 Then
      MsgBox "Delivered Bags in 'WHR No' Tab and 'No. of Bags' in 'GSL Details' tab are not equal !!!!"
'      MsgBox "Delivered Bags in 'RRN Detail' Tab and 'No of Bags' in 'GSL Details' tab are not equal !!!!"
      Exit Sub
   ElseIf mExchangeTypeID = 3 Then
      MsgBox "Delivered Bags in 'CDC Detail' Tab and 'No of Bags' in 'GSL Details' tab are not equal !!!!"
      Exit Sub
   End If
End If


'Commented on 13 th Dec 2011 after Mail from Mayur Sir and approval of Sankar Sir
'Dim ans As Variant
'
'If mExchangeTypeID = 1 Then
'   If ChkCorporateClient = False Then
'      If ChkCRMapping = False Then
'         If LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "headcmg" Then
'            MsgBox "No Cash receipt mapped for this CITF. Please map Cash Receipt using Cash Receipt CITF Mapping!!!"
'            Exit Sub
'         Else
'            ans = MsgBox("Warning : No Cash receipt mapped for this CITF. Still you want to map hologram to CITF!!!", vbYesNo)
'            If ans = vbNo Then
'               Exit Sub
'            End If
'         End If
'        End If
'   End If
'ElseIf mExchangeTypeID = 2 And Check1.Value = 1 Then
'   If ChkCRMapping = False Then
'      If LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "head cmg" And LCase(Gen_UserRole) <> "superadmin" Then
'         MsgBox "No Cash receipt mapped for this CITF. Please map Cash Receipt using Cash Receipt CITF Mapping!!!"
'         Exit Sub
'      Else
'         ans = MsgBox("Warning : No Cash receipt mapped for this CITF. Still you want to map hologram to CITF!!!", vbYesNo)
'         If ans = vbNo Then
'            Exit Sub
'         End If
'      End If
'   End If
'End If



Check3.Visible = False
mSpaceChargesForInvoice = 0
If LCase(UserType) = "power" Then
   Check3.Visible = True
End If

If RsTxn_CxTF_Details!HologramNo = "" Or IsNull(RsTxn_CxTF_Details!HologramNo) Then
   

   If CheckForBOEAmount() = True And Val(TxtBOEAmount.Text) = 0 Then
      MsgBox "BOE amount should not zero for this CITF!!"
      Exit Sub
   End If
   
   
   Call GetClientFeedbackBlockingDays

     Dim mR, mC As Double
     Dim mWHRandNNR As Variant
    
     With MSHFlexGrid2
        If .Rows > 3 Then
            tmp = " SELECT TCFD.ClientID,MAX(TCFD.TrancationDate)[TrancationDate] FROM Txn_ClientFeedbackDoc AS TCFD Inner Join Mst_Client AS MC On  TCFD.ClientID = MC.ClientID WHERE "
            tmp = tmp & " MC.NAVCode in (SELECT NAVCode FROM Mst_Client Where NAVCode IS NOT NULL AND ClientID IN (SELECT ClientID FROM Txn_CxTF_ClientDetail Where TxnType = 'W' AND CxTF_DetailsID = " & Val(TxtCxTFNo) & " AND SM_Prefix = '" & TxtPreFix & "')) OR "
            tmp = tmp & " MC.RNAVCode in (SELECT RNAVCode FROM Mst_Client Where RNAVCode IS NOT NULL AND ClientID IN (SELECT ClientID FROM Txn_CxTF_ClientDetail Where TxnType = 'W' AND CxTF_DetailsID = " & Val(TxtCxTFNo) & " AND SM_Prefix = '" & TxtPreFix & "')) "
            tmp = tmp & " GROUP BY TCFD.ClientID"
    
            Call TmpTable
    
            If TmpRs.RecordCount > 0 Then
               For mC = 1 To TmpRs.RecordCount
                   If DateDiff("d", Date, CDate(TmpRs!TrancationDate)) > mBlockingDays Then
                      MsgBox "Client's feedback document is not submitted from last " & mBlockingDays & " Days.", vbOKOnly + vbExclamation, App.Title
                      Exit Sub
                   End If
                Next mC
            Else
                MsgBox "Client's feedback document is not submitted from last " & mBlockingDays & " Days.", vbOKOnly + vbExclamation, App.Title
                Exit Sub
            End If
        End If
     End With


   
     
   If LCase(CmbExchangeTypeID.Text) = "non ncdex" Then
   
      
      If Val(mCommodityYearlyFlag) = 1 And ChkGodownYearlyFlag = True Then
         Call TableTxn_CDTFBOE(" Where SM_Prefix='" & TxtPreFix & "' And CxTF_DetailsID=" & Val(TxtCxTFNo) & " And TxnType='Y' And ((Amount + Premium) - Discount) > 0")
               
         If RsTxn_CDTFBOE.RecordCount > 0 Then
            
            For Ln = 1 To RsTxn_CDTFBOE.RecordCount
                mBal = GetBalAmt(RsTxn_CDTFBOE!BOEID, "Y")
                If mBal = "NA" Then
                   MsgBox "TI Details not uploaded for Yearly Invoice " & RsTxn_CDTFBOE!BOEID & "!!"
                   Exit Sub
                ElseIf Val(mBal) > 0 Then
                   If IsNull(RsTxn_CxTF_Details!ApprovePrintFlag) = True Or RsTxn_CxTF_Details!ApprovePrintFlag = "N" Then
                      
                      If LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "headcmg" Then
                         MsgBox "Full Invoice amount not received against Yearly Invoice : " & RsTxn_CDTFBOE!BOEID & "!!"
                         Exit Sub
                      Else
                         If MsgBox("Full Invoice amount not received against Yearly Invoice : " & RsTxn_CDTFBOE!BOEID & ". Still you want to take print", vbYesNo) = vbNo Then
                            Exit Sub
                         Else
                            TxtARemark.Text = ""
                            Frame2.Visible = True
                            Exit Sub
                         End If
                      End If
                   End If
                End If
                RsTxn_CDTFBOE.MoveNext
            Next
         Else
            mRejInvIDs = ""
            Call SaveYearlyBilling
            If mRejInvIDs <> "" Then
               MsgBox "Yearly Invoice " & mRejInvIDs & " generated, please confirm payment & then take print out!!"
               Exit Sub
            End If
         End If
      Else
         MsgBox "Please ensure that an invoice is generated for the BoE amount shown on the left side before approving and taking a print out of this CITF."
         
         If IsNull(RsTxn_CxTF_Details!ApprovePrintFlag) = True Or RsTxn_CxTF_Details!ApprovePrintFlag = "N" Then
            MsgBox "As withdrawal is not yet approved, you can't take print out!!"
            Exit Sub
         End If
      End If
   Else
      If Check1.Value = 1 Then
         Call TableTxn_CDTFBOE(" Where SM_Prefix='" & TxtPreFix & "' And CxTF_DetailsID=" & Val(TxtCxTFNo) & " And TxnType='W' And ((Amount + Premium) - Discount) > 0 ")
               
         If RsTxn_CDTFBOE.RecordCount > 0 Then
            
            For Ln = 1 To RsTxn_CDTFBOE.RecordCount
                mBal = GetBalAmt(RsTxn_CDTFBOE!BOEID, "W")
                If mBal = "NA" Then
                   MsgBox "TI Details not uploaded for Rejected Invoice " & RsTxn_CDTFBOE!BOEID & "!!"
                   Exit Sub
                ElseIf Val(mBal) > 0 Then
                   If IsNull(RsTxn_CxTF_Details!ApprovePrintFlag) = True Or RsTxn_CxTF_Details!ApprovePrintFlag = "N" Then
                      
                      If LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "headcmg" Then
                         MsgBox "Full Invoice amount not received against Rejected Invoice : " & RsTxn_CDTFBOE!BOEID & "!!"
                         Exit Sub
                      Else
                         If MsgBox("Full Invoice amount not received against Rejected Invoice : " & RsTxn_CDTFBOE!BOEID & ". Still you want to take print", vbYesNo) = vbNo Then
                            Exit Sub
                         Else
                            TxtARemark.Text = ""
                            Frame2.Visible = True
                            Exit Sub
                         End If
                      End If
                   End If
                End If
                RsTxn_CDTFBOE.MoveNext
            Next
         Else
            If CheckNCDEXRate = False Then
               MsgBox "No Rate updated in NCDEX rate card master or Franchisee Rate Card master. Please contact CMG Dept for updating NCDEX rate card master!!!"
               Exit Sub
            End If
            mRejInvIDs = ""
            Call SaveCITFBOE
            If mRejInvIDs <> "" Then
               MsgBox "Rejected Invoice " & mRejInvIDs & " generated, please confirm payment & then take print out!!"
               Exit Sub
            End If
         End If
      Else
'         Frame5.Visible = True
         Call TableTxn_CDTFBOE(" Where SM_Prefix='" & TxtPreFix & "' And CxTF_DetailsID=" & Val(TxtCxTFNo) & " And TxnType='R' And ((Amount + Premium) - Discount) > 0 ")
               
         If RsTxn_CDTFBOE.RecordCount > 0 Then
            
            For Ln = 1 To RsTxn_CDTFBOE.RecordCount
                mBal = GetBalAmt(RsTxn_CDTFBOE!BOEID, "R")
                If mBal = "NA" Then
                   MsgBox "TI Details not uploaded for RRN Withdrawal Invoice " & RsTxn_CDTFBOE!BOEID & "!!"
                   Exit Sub
                ElseIf Val(mBal) > 0 Then
                   If IsNull(RsTxn_CxTF_Details!ApprovePrintFlag) = True Or RsTxn_CxTF_Details!ApprovePrintFlag = "N" Then
                      
                      If LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "headcmg" Then
                         MsgBox "Full Invoice amount not received against RRN Withdrawal Invoice : " & RsTxn_CDTFBOE!BOEID & "!!"
                         Exit Sub
                      Else
                         If MsgBox("Full Invoice amount not received against RRN Withdrawal Invoice : " & RsTxn_CDTFBOE!BOEID & ". Still you want to take print", vbYesNo) = vbNo Then
                            Exit Sub
                         Else
                            TxtARemark.Text = ""
                            Frame2.Visible = True
                            Exit Sub
                         End If
                      End If
                   End If
                End If
                RsTxn_CDTFBOE.MoveNext
            Next
         Else
            If CheckNCDEXRate = False Then
               MsgBox "No Rate updated in NCDEX rate card master. Please contact CMG Dept for updating NCDEX rate card master!!!"
               Exit Sub
            End If
            mRejInvIDs = ""
            Call CreateRRNNCDEXInvoice
            If mRejInvIDs <> "" Then
               MsgBox "RRN Invoice " & mRejInvIDs & " generated, please confirm payment & then take print out!!"
               Exit Sub
            End If
         End If
      End If
   End If

   If PntPrv = "N" Then
      ans = MsgBox("We strongly recommend that you take a look at Print-Preview first before going ahead with this Print option. Do you still want to continue?", vbYesNo)
      If ans = vbNo Then
         Exit Sub
      End If
   End If
   
   Frame3.Visible = True
   TxtHologramNo.SetFocus
   Call TableMst_Hologram
   Exit Sub

Else
   TxtHologramNo = ""
   Frame3.Visible = False
End If
RsTxn_CxTF_Details!Flag = "P"
RsTxn_CxTF_Details.Update
''
Gen_mTimeStamp = GetTimeStamp
''Call Create_Xls(Gen_mTimeStamp & Gen_UserName & "_CITF-" & Trim(TxtPreFix) & "" & Trim(TxtCxTFNo) & ".xls")
''Call Xls_Print_Rpt

PntPrv = "N"
Call PDF_Print_Rpt
Call Gen_Email
End Sub
Private Function ChkCRMapping() As Boolean

Dim Retval As Boolean

Retval = False

tmp = " Select TCC.Flag,TCCD.* from Txn_CITFCRDetail TCCD "
tmp = tmp & " Inner Join Txn_CITFCR TCC  On TCCD.CITFCRID = TCC.CITFCRID"
tmp = tmp & " Where  TCCD.SM_Prefix = '" & TxtPreFix & "' And TCCD.CITFNo = " & TxtCxTFNo & " "

Call Tmp4Table

If TmpRs4.RecordCount > 0 Then
   Retval = True
End If

ChkCRMapping = Retval

End Function
Private Function ChkCorporateClient() As Boolean

Dim Retval As Boolean

Retval = True
mCorporateGSLID = ""
For I = 2 To MSHFlexGrid4.Rows - 1
    If mCorporateGSLID = "" Then
       mCorporateGSLID = "(" & MSHFlexGrid4.TextMatrix(I, 14)
    Else
       If MSHFlexGrid4.TextMatrix(I, 14) <> "" Then
          mCorporateGSLID = mCorporateGSLID & "," & MSHFlexGrid4.TextMatrix(I, 14)
       End If
    End If
    If I = MSHFlexGrid4.Rows - 1 Then
       mCorporateGSLID = mCorporateGSLID & ")"
    End If
Next

tmp = " Select * from Mst_GSL Where Mst_GSL_ID in " & mCorporateGSLID & " And SM_Prefix = '" & TxtPreFix & "' And (ContractID = 0 Or ContractID is null) And SpillageFlag = 0 "
Call Tmp4Table

If TmpRs4.RecordCount > 0 Then
   Retval = False
End If

ChkCorporateClient = Retval

End Function

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
'Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & mTimeStamp & Gen_UserName & "_CITF-" & Trim(TxtPreFix) & "" & Trim(TxtCxTFNo) & ".xls")
 
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "_CITF-" & Trim(TxtPreFix) & "" & Trim(TxtCxTFNo) & ".xls")
 
I = 2
Set oWS = oWB.Worksheets("Sheet" & 1)
'oWS.PageSetup.RightMargin = 0.05
oWS.Name = TxtPreFix & "-" & TxtCxTFNo  'MSHFlexGrid2.TextMatrix(i, 23)
oWS.ClearArrows
mRow = 4

oWS.Range("a1:f1").Merge
oWS.Range("a:a").ColumnWidth = 10.86
oWS.Range("b:b").ColumnWidth = 12.86
oWS.Range("c:c").ColumnWidth = 12.86
oWS.Range("d:d").ColumnWidth = 12.86
oWS.Range("e:e").ColumnWidth = 17.86 '32.14
oWS.Range("f:f").ColumnWidth = 17.86


tmp = "a" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = GetCMPAddress(RsTxn_CxTF_Details!CMPID, "D") '& "-" & GetCMPName(RsTxn_CxTF_Details!CMPID) '& "-" & GetCMPAddress(RsTxn_CxTF_Details!CMPID)
oWS.Cells(mRow, C + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 10
'oWS.Cells(mRow, c + 1).RowHeight = 24
mRow = mRow + 1
'oWS.Range("A1:E1").Borders.Color = RGB(0, 0, 0)

tmp = "a" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "Commodity Issue Transaction Form (CITF)"
oWS.Cells(mRow, C + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 10
oWS.Cells(mRow, C + 1).RowHeight = 24


mRow = mRow + 1

tmp = "a" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "(This is NOT an instrument of title to the goods)"
oWS.Cells(mRow, C + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 7


mRow = mRow + 1

oWS.Cells(mRow, C + 1) = "CITF No"
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).WrapText = True
'oWS.Range("b5:c5").Merge
tmp = "b" & mRow & ":" & "c" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1).Font.Size = 7
oWS.Cells(mRow, C + 2) = TxtPreFix & "-" & TxtCxTFNo
oWS.Cells(mRow, C + 2).HorizontalAlignment = 2
oWS.Cells(mRow, C + 2).Font.Size = 7


oWS.Cells(mRow, C + 4) = "Date"
oWS.Cells(mRow, C + 4).Font.Size = 7
oWS.Cells(mRow, C + 4).Font.Bold = True

'oWS.Range(tmp).Merge
tmp = "e" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 5) = Format(TxtCxTFDate, "dd-mmm-yyyy")
oWS.Cells(mRow, C + 5).HorizontalAlignment = 2
oWS.Cells(mRow, C + 5).Font.Size = 7

mRow = mRow + 1

oWS.Cells(mRow, C + 1) = "Location"
oWS.Cells(mRow, C + 1).Font.Size = 7
oWS.Cells(mRow, C + 1).Font.Bold = True
tmp = "b" & mRow & ":" & "c" & mRow

oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 2) = GetLocationFromCMP(RsTxn_CxTF_Details!CMPID)
oWS.Cells(mRow, C + 2).HorizontalAlignment = 2
oWS.Cells(mRow, C + 2).Font.Size = 7

oWS.Cells(mRow, C + 4) = "State"
oWS.Cells(mRow, C + 4).Font.Size = 7
oWS.Cells(mRow, C + 4).Font.Bold = True

tmp = "e" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 5) = GetStateFromLocation(GetLocationIDFromCMP(RsTxn_CxTF_Details!CMPID))
oWS.Cells(mRow, C + 5).HorizontalAlignment = 2
oWS.Cells(mRow, C + 5).Font.Size = 7

mRow = mRow + 1

oWS.Cells(mRow, C + 1) = "CMP ID & Name"
oWS.Cells(mRow, C + 1).Font.Size = 7
oWS.Cells(mRow, C + 1).WrapText = True
oWS.Cells(mRow, C + 1).Font.Bold = True

tmp = "b" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 2) = RsTxn_CxTF_Details!CMPID & "-" & GetCMPName(RsTxn_CxTF_Details!CMPID)
oWS.Cells(mRow, C + 2).HorizontalAlignment = 2
oWS.Cells(mRow, C + 2).Font.Size = 7
oWS.Cells(mRow, C + 2).WrapText = True

mRow = mRow + 1

oWS.Cells(mRow, C + 1) = "CMP Address"
oWS.Cells(mRow, C + 1).Font.Size = 7
oWS.Cells(mRow, C + 1).WrapText = True
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).RowHeight = 50

tmp = "b" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 2) = GetCMPAddress(RsTxn_CxTF_Details!CMPID, "A")
oWS.Cells(mRow, C + 2).HorizontalAlignment = 2
oWS.Cells(mRow, C + 2).Font.Size = 7
oWS.Cells(mRow, C + 2).WrapText = True

mRow = mRow + 1

oWS.Cells(mRow, C + 1) = "Transaction Type"
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 7
oWS.Cells(mRow, C + 1).WrapText = True

tmp = "b" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 2) = CmbExchangeTypeID.Text
oWS.Cells(mRow, C + 2).Font.Size = 7

mRow = mRow + 1

oWS.Cells(mRow, C + 1) = "Name Of Agent"
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 7

tmp = "b" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 2) = TxtAgent
oWS.Cells(mRow, C + 2).Font.Size = 7

mRow = mRow + 1

oWS.Cells(mRow, C + 1) = "Commodity"
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 7

tmp = "b" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 2) = TxtCommodityID
oWS.Cells(mRow, C + 2).Font.Size = 7

mRow = mRow + 1

tmp = "a" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "Withdrawer Details"
oWS.Cells(mRow, C + 1).Font.Size = 9
oWS.Cells(mRow, C + 1).Font.Bold = True

mRow = mRow + 1
If CmbExchangeTypeID.Text = "NCDEX" Then
   If Check1.Value = 0 Then
      oWS.Cells(mRow, C + 1) = "RRN No"
   Else
      oWS.Cells(mRow, C + 1) = "Rejected Withdrawal"
   End If
ElseIf CmbExchangeTypeID.Text = "Non NCDEX" Then
   oWS.Cells(mRow, C + 1) = "WHR No"
ElseIf CmbExchangeTypeID.Text = "SPOT" Then
   oWS.Cells(mRow, C + 1) = "CDC No"
End If
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 7
oWS.Cells(mRow, C + 2) = "Client-ID"
oWS.Cells(mRow, C + 2).Font.Bold = True
oWS.Cells(mRow, C + 2).Font.Size = 7

If mExchangeTypeID <> 2 Then
   tmp = "c" & mRow & ":" & "d" & mRow
   oWS.Range(tmp).Merge
Else
   tmp = "c" & mRow & ":" & "e" & mRow
   oWS.Range(tmp).Merge
End If

oWS.Cells(mRow, C + 3) = "Client Name"
oWS.Cells(mRow, C + 3).Font.Bold = True
oWS.Cells(mRow, C + 3).Font.Size = 7
If mExchangeTypeID <> 2 Then
   oWS.Cells(mRow, C + 5) = "No. of Bags"
   oWS.Cells(mRow, C + 5).Font.Bold = True
   oWS.Cells(mRow, C + 5).Font.Size = 7
End If

oWS.Cells(mRow, C + 6) = "Electronic Weight in MT"
oWS.Cells(mRow, C + 6).Font.Bold = True
oWS.Cells(mRow, C + 6).Font.Size = 7

mRow = mRow + 1

For k = 2 To MSHFlexGrid2.Rows - 2
'k = 1
    oWS.Cells(mRow, C + 1) = "'" & MSHFlexGrid2.TextMatrix(k, 1)
    oWS.Cells(mRow, C + 1).Font.Size = 7
    oWS.Cells(mRow, C + 2) = MSHFlexGrid2.TextMatrix(k, 3)
    oWS.Cells(mRow, C + 2).Font.Size = 7
    If mExchangeTypeID <> 2 Then
       tmp = "c" & mRow & ":" & "d" & mRow
       oWS.Range(tmp).Merge
    Else
       tmp = "c" & mRow & ":" & "e" & mRow
       oWS.Range(tmp).Merge
    End If
    oWS.Cells(mRow, C + 3) = MSHFlexGrid2.TextMatrix(k, 4)
    oWS.Cells(mRow, C + 3).Font.Size = 7
    If mExchangeTypeID <> 2 Then
       oWS.Cells(mRow, C + 5) = Format(MSHFlexGrid2.TextMatrix(k, 15), "#######0.000")
       oWS.Cells(mRow, C + 5).Font.Size = 7
    End If
    oWS.Cells(mRow, C + 6) = Format(MSHFlexGrid2.TextMatrix(k, 9), "#######0.000")
    oWS.Cells(mRow, C + 6).Font.Size = 7
    mRow = mRow + 1
Next k

If mExchangeTypeID <> 2 Then
   tmp = "c" & mRow & ":" & "d" & mRow
   oWS.Range(tmp).Merge
Else
   tmp = "c" & mRow & ":" & "e" & mRow
   oWS.Range(tmp).Merge
End If
mRow = mRow + 1

If mExchangeTypeID <> 2 Then
   tmp = "a" & mRow & ":" & "d" & mRow
   oWS.Range(tmp).Merge
Else
   tmp = "a" & mRow & ":" & "e" & mRow
   oWS.Range(tmp).Merge
End If
oWS.Cells(mRow, C + 1) = "Total"
oWS.Cells(mRow, C + 1).Font.Size = 7
oWS.Cells(mRow, C + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, C + 1).Font.Bold = True

If mExchangeTypeID <> 2 Then
   oWS.Cells(mRow, C + 5) = Format(MSHFlexGrid2.TextMatrix(0, 15), "#######0.000")
   oWS.Cells(mRow, C + 5).Font.Size = 7
End If
oWS.Cells(mRow, C + 6) = Format(MSHFlexGrid2.TextMatrix(0, 9), "#######0.000")
oWS.Cells(mRow, C + 6).Font.Size = 7

mRow = mRow + 1

tmp = "a" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "Truck-wise Withdrawal Details"
oWS.Cells(mRow, C + 1).Font.Size = 9
oWS.Cells(mRow, C + 1).Font.Bold = True

mRow = mRow + 1

oWS.Cells(mRow, C + 1) = "Truck No"
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 7
oWS.Cells(mRow, C + 2) = "Total Weight(MT)"
oWS.Cells(mRow, C + 2).Font.Bold = True
oWS.Cells(mRow, C + 2).Font.Size = 7
oWS.Cells(mRow, C + 3) = "Tare Weight(MT)"
oWS.Cells(mRow, C + 3).Font.Bold = True
oWS.Cells(mRow, C + 3).Font.Size = 7
oWS.Cells(mRow, C + 4) = "Gross Weight"
oWS.Cells(mRow, C + 4).Font.Bold = True
oWS.Cells(mRow, C + 4).Font.Size = 7
oWS.Cells(mRow, C + 5) = "Weighbridge Slip No"
oWS.Cells(mRow, C + 5).Font.Bold = True
oWS.Cells(mRow, C + 5).Font.Size = 7
oWS.Cells(mRow, C + 6) = "Remarks"
oWS.Cells(mRow, C + 6).Font.Bold = True
oWS.Cells(mRow, C + 6).Font.Size = 7

mRow = mRow + 1

For k = 2 To MSHFlexGrid3.Rows - 2
'k = 1
    oWS.Cells(mRow, C + 1) = MSHFlexGrid3.TextMatrix(k, 1)
    oWS.Cells(mRow, C + 1).Font.Size = 7
    oWS.Cells(mRow, C + 2) = Format(MSHFlexGrid3.TextMatrix(k, 3), "#######0.000")
    oWS.Cells(mRow, C + 2).Font.Size = 7
    oWS.Cells(mRow, C + 3) = Format(MSHFlexGrid3.TextMatrix(k, 4), "#######0.000")
    oWS.Cells(mRow, C + 3).Font.Size = 7
    oWS.Cells(mRow, C + 4) = Format(MSHFlexGrid3.TextMatrix(k, 5), "#######0.000")
    oWS.Cells(mRow, C + 4).Font.Size = 7
    oWS.Cells(mRow, C + 5) = MSHFlexGrid3.TextMatrix(k, 2)
    oWS.Cells(mRow, C + 5).Font.Size = 7
    oWS.Cells(mRow, C + 6) = MSHFlexGrid3.TextMatrix(k, 6)
    oWS.Cells(mRow, C + 6).Font.Size = 7
    mRow = mRow + 1
Next k
mRow = mRow + 1

tmp = "a" & mRow & ":" & "c" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "Total"
oWS.Cells(mRow, C + 1).Font.Size = 7
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).HorizontalAlignment = -4108

oWS.Cells(mRow, C + 4) = Format(MSHFlexGrid3.TextMatrix(0, 5), "#######0.000")
oWS.Cells(mRow, C + 4).Font.Size = 7

tmp = "e" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge

tmp = "e" & mRow & ":" & "e" & (mRow + 1)
oWS.Range(tmp).Merge

oWS.Cells(mRow, C + 5) = "Affix Hologram here."
oWS.Cells(mRow, C + 5).Font.Size = 7
oWS.Cells(mRow, C + 5).Font.Bold = True


mRow = mRow + 1

tmp = "a" & mRow & ":" & "c" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "Less Gunny Weight(MT)"
oWS.Cells(mRow, C + 1).Font.Size = 7
oWS.Cells(mRow, C + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, C + 1).Font.Bold = True

oWS.Cells(mRow, C + 4) = Format(MSHFlexGrid4.TextMatrix(0, 9), "#######0.000")
oWS.Cells(mRow, C + 4).Font.Size = 7

tmp = "e" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge

tmp = "e" & mRow & ":" & "e" & (mRow + 1)
oWS.Range(tmp).Merge

mRow = mRow + 1

tmp = "a" & mRow & ":" & "c" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "Less Standard Deductions(MT)"
oWS.Cells(mRow, C + 1).Font.Size = 7
oWS.Cells(mRow, C + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, C + 1).Font.Bold = True

oWS.Cells(mRow, C + 4) = "" '
oWS.Cells(mRow, C + 4).Font.Size = 7

tmp = "e" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge

tmp = "e" & mRow & ":" & "e" & (mRow + 1)
oWS.Range(tmp).Merge

mRow = mRow + 1

tmp = "a" & mRow & ":" & "c" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "Add Adjustments(MT)"
oWS.Cells(mRow, C + 1).Font.Size = 7
oWS.Cells(mRow, C + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, C + 1).Font.Bold = True

oWS.Cells(mRow, C + 4) = Format(MSHFlexGrid4.TextMatrix(0, 8), "#######0.000")
oWS.Cells(mRow, C + 4).Font.Size = 7

tmp = "e" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge

If mExchangeTypeID = 2 Then
   tmp = "e" & mRow & ":" & "e" & (mRow + 1)
   oWS.Range(tmp).Merge
End If
mRow = mRow + 1


tmp = "a" & mRow & ":" & "c" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "Electronic Weight(MT)"
oWS.Cells(mRow, C + 1).Font.Size = 7
oWS.Cells(mRow, C + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, C + 1).Font.Bold = True

oWS.Cells(mRow, C + 4) = Format(MSHFlexGrid4.TextMatrix(0, 10), "#######0.000")
oWS.Cells(mRow, C + 4).Font.Size = 7

mRow = mRow + 1

If mExchangeTypeID = 2 Then
   tmp = "a" & mRow & ":" & "c" & mRow
   oWS.Range(tmp).Merge
   oWS.Cells(mRow, C + 1) = "Total No. of Bags"
   oWS.Cells(mRow, C + 1).Font.Size = 7
   oWS.Cells(mRow, C + 1).HorizontalAlignment = -4108
   oWS.Cells(mRow, C + 1).Font.Bold = True

   oWS.Cells(mRow, C + 4) = Format(MSHFlexGrid4.TextMatrix(0, 6)) 'Format(MSHFlexGrid4.TextMatrix(0, 8), "#######0.000")
   oWS.Cells(mRow, C + 4).Font.Size = 7
'   tmp = "e" & mRow & ":" & "f" & mRow
'   oWS.Range(tmp).Merge
   
'   tmp = "e" & mRow & ":" & "e" & (mRow + 1)
'   oWS.Range(tmp).Merge
   oWS.Cells(mRow, C + 5) = "Hologram No"
   oWS.Cells(mRow, C + 5).Font.Size = 7
   oWS.Cells(mRow, C + 5).Font.Bold = True
   If RsTxn_CxTF_Details!HologramNo = "" Or IsNull(RsTxn_CxTF_Details!HologramNo) Then
      oWS.Cells(mRow, C + 6) = "Preview"
   Else
      oWS.Cells(mRow, C + 6) = RsTxn_CxTF_Details!HologramNo
   End If
   oWS.Cells(mRow, C + 6).Font.Size = 7
   oWS.Cells(mRow, C + 6).HorizontalAlignment = -4108
   mRow = mRow + 1
Else
   oWS.Cells(mRow - 1, C + 5) = "Hologram No"
   oWS.Cells(mRow - 1, C + 5).Font.Size = 7
   oWS.Cells(mRow - 1, C + 5).Font.Bold = True
   If RsTxn_CxTF_Details!HologramNo = "" Or IsNull(RsTxn_CxTF_Details!HologramNo) Then
      oWS.Cells(mRow - 1, C + 6) = "Preview"
   Else
      oWS.Cells(mRow - 1, C + 6) = RsTxn_CxTF_Details!HologramNo
   End If
   oWS.Cells(mRow - 1, C + 6).Font.Size = 7
   oWS.Cells(mRow - 1, C + 6).HorizontalAlignment = -4108
End If

mRow = mRow + 1
mRow = mRow + 1
mRow = mRow + 1

tmp = "a" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "Name & Signature of CMP Manager/CMP DRC:"
oWS.Cells(mRow, C + 1).Font.Size = 9
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).RowHeight = 40
mRow = mRow + 1

tmp = "a" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "Received the above mentioned goods in good condition."
oWS.Cells(mRow, C + 1).Font.Size = 10
oWS.Cells(mRow, C + 1).WrapText = True
oWS.Cells(mRow, C + 1).Font.Name = "Zurich BT"

mRow = mRow + 1

tmp = "a" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
mRow = mRow + 1

tmp = "a" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
mRow = mRow + 1


tmp = "a" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge

mRow = mRow + 1

tmp = "a" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge

mRow = mRow + 1

tmp = "a" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "Name & Signature of Withdrawer/Agent of Withdrawer:"
oWS.Cells(mRow, C + 1).Font.Size = 9
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Name = "Zurich BT"
k = mRow - 6

For Cell = 4 To k
    tmp = "a" & Cell & ":" & "f" & Cell
    oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
    oWS.Range(tmp).Font.Name = "Zurich BT"
    'oWS.Range(tmp).Font.Size = 7
    oWS.Range(tmp).WrapText = True
Next Cell
    
If Gen_ExcelPassword <> "" Then
    oWS.Protect Gen_ExcelPassword, , , , , , True, True
End If
oWB.Save
Call Log_Print_History(TxtPreFix, "CITF", TxtCxTFNo, Gen_FLName)
    
    
'oWB.Close
'oXL.Quit
MsgBox ("Excel file created!!!")
oXL.Visible = True





End Sub


Private Sub TxtHologramNo_LostFocus()
'If TxtHologramNo = "" Then Exit Sub
'If Check_HologramNo(TxtHologramNo) = False Then
'   TxtHologramNo.SetFocus
'   CmdSaveHologram.Enabled = False
'   Exit Sub
'End If
'CmdSaveHologram.Enabled = True
'CmdSaveHologram.SetFocus

If TxtHologramNo = "" Then Exit Sub
'If Check_HologramNo(TxtHologramNo) = False Then
'   TxtHologramNo.SetFocus
'   CmdSaveHologram.Enabled = False
'   Exit Sub
'End If


If TxtHologramNo <> "0" Then

   If Check_HologramNo(TxtHologramNo) = False Then
      TxtHologramNo.SetFocus
      CmdSaveHologram.Enabled = False
      Exit Sub
   End If
    
   If CheckHologramInRange(TxtHologramNo, mLocationID) = False Then
      MsgBox "Hologram No not found in any range."
      TxtHologramNo.SetFocus
      CmdSaveHologram.Enabled = False
      Exit Sub
   End If
End If

CmdSaveHologram.Enabled = True
CmdSaveHologram.SetFocus

End Sub

Private Sub CmdSaveHologram_Click()
On Error GoTo msgError

Dim ans As Variant

If BalBagNeg = "Y" Then
   MsgBox "One or more GSL having Negative  'No of Bags' please contact CMG/IT for assistance!!!"
   Exit Sub
End If
If TxtHologramNo = "" Then
   MsgBox "Blank Hologram Number not allowed!!!"
   TxtHologramNo.SetFocus
   Exit Sub
End If
ans = MsgBox("Are you sure you want to save this Hologram number?", vbYesNo)
Call TableMst_Hologram
If ans = vbYes Then
   If TxtHologramNo <> "0" Then
      RsMst_Hologram.AddNew
      RsMst_Hologram!NewHologramNo = TxtHologramNo
      RsMst_Hologram!DocType = "CITF"
      RsMst_Hologram!DocNo = TxtCxTFNo
      RsMst_Hologram!SM_Prefix = TxtPreFix
      RsMst_Hologram!CreatedBy = Gen_UserLoginID
      RsMst_Hologram.Update
   End If
   RsTxn_CxTF_Details!HologramNo = TxtHologramNo
   RsTxn_CxTF_Details.Update
   TxtHologramNo = ""
   Frame3.Visible = False
   RsTxn_CxTF_Details!Flag = "P"
   RsTxn_CxTF_Details.Update
   RsTxn_CxTF_Details.Requery
   RsTxn_CxTF_Details.MoveFirst
   RsTxn_CxTF_Details.Find "G_UID = '" & TxtGUID & "'"
   'TxtGUID
''   If LCase(CmbExchangeTypeID.Text) = "non ncdex" Then
''      Call FillBOE
''      Call SaveCxTFBOE
''   End If
   Gen_mTimeStamp = GetTimeStamp
   If Check3.Value = 1 Then
''      Call Create_Xls(Gen_mTimeStamp & Gen_UserName & "_CITF-" & Trim(TxtPreFix) & "" & Trim(TxtCxTFNo) & ".xls")
''      Call Xls_Print_Rpt
      PntPrv = "N"
      Call PDF_Print_Rpt
      Call Gen_Email
   End If
Else
   TxtHologramNo.SetFocus
End If
'''''
Exit Sub

msgError:

MsgBox Err.Description, vbExclamation, "Error" '"Some error occurred during process!!!"


End Sub


'Private Sub CmbPreFix_GotFocus()
'tmp = CmbPreFix.Text
'Call TableMst_StoreManLocations("")
'
'CmbPreFix.Clear
'If RsMst_StoreManLocations.RecordCount = 0 Then
'   MsgBox "No Bank found!!!"
'   CmbPreFix.SetFocus
'   Exit Sub
'End If
'For i = 1 To RsMst_StoreManLocations.RecordCount
'    CmbPreFix.AddItem RsMst_StoreManLocations!SM_Prefix
'    RsMst_StoreManLocations.MoveNext
'Next
'CmbPreFix.Text = tmp
'End Sub
'
'Private Sub CmbPreFix_LostFocus()
'If CmbPreFix.Text = "" Then
'   Exit Sub
'End If
'RsMst_StoreManLocations.MoveFirst
'RsMst_StoreManLocations.Find "SM_Prefix = '" & CmbPreFix.Text & "'"
'
'If RsMst_StoreManLocations.EOF Then
'   MsgBox "Invalid Location-ID selection!!!"
'   CmbPreFix.SetFocus
'   Exit Sub
'End If
'msSM_prefix = RsMst_StoreManLocations!SM_Prefix
'
'End Sub



Private Sub CmbPreFix_GotFocus()
tmp = CmbPreFix.Text
'Call TableMst_Location

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
   mLocationID = Null
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

If tmp <> CmbPreFix.Text Then
   CmbSCmp.Text = ""
   mCMPID = 0
End If
mSLocationID = RsMst_Location!LocationID
mLocationName = RsMst_Location!Locationname
msSM_prefix = RsMst_Location!SM_Prefix
End Sub

Private Sub CmbSCmp_GotFocus()
tmp = CmbSCmp.Text
If msSM_prefix = "" Then
   MsgBox "Select Location !!!"
   CmbPreFix.SetFocus
   Exit Sub
End If
Call TableMst_CMP(" Where LocationID = " & mSLocationID & "")
CmbSCmp.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "No record found!!! "
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
   MsgBox "Invalid selection!!!"
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
   MsgBox "No record found!!! "
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
   MsgBox "Invalid selection!!!"
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
   MsgBox "No record found!!!"
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
   MsgBox "Invalid selection!!!"
   CmbSExchangeType.SetFocus
   Exit Sub
End If
mExchangeTypeID = RsMst_ExchangeType!ExchangeTypeID

End Sub

Private Sub Grid_Head_Bill()

MSHFlexGrid6.Clear
MSHFlexGrid6.Rows = 2
MSHFlexGrid6.FixedRows = 1
MSHFlexGrid6.Cols = 22

MSHFlexGrid6.TextMatrix(0, 0) = "CIDF Date"
MSHFlexGrid6.TextMatrix(0, 1) = "CDTF Date"
MSHFlexGrid6.TextMatrix(0, 2) = "No. of days"
MSHFlexGrid6.TextMatrix(0, 3) = "ClientRowID"
MSHFlexGrid6.TextMatrix(0, 4) = "Client Name"
MSHFlexGrid6.TextMatrix(0, 5) = "GSLID"
MSHFlexGrid6.TextMatrix(0, 6) = "G-S-L"
MSHFlexGrid6.TextMatrix(0, 7) = "Lot Type"
MSHFlexGrid6.TextMatrix(0, 8) = "No of Bags withdrawn"
MSHFlexGrid6.TextMatrix(0, 9) = "Withdrawn Qty"
MSHFlexGrid6.TextMatrix(0, 10) = "Bag Size"
MSHFlexGrid6.TextMatrix(0, 11) = "NContractID"
MSHFlexGrid6.TextMatrix(0, 12) = "Billing Cycle"
MSHFlexGrid6.TextMatrix(0, 13) = "Billing Unit"
MSHFlexGrid6.TextMatrix(0, 14) = "No. of Cycle"
MSHFlexGrid6.TextMatrix(0, 15) = "Rate per Cycle"
MSHFlexGrid6.TextMatrix(0, 16) = "Storage Charges Amount"
MSHFlexGrid6.TextMatrix(0, 17) = "Deposit Billing Cycle"
MSHFlexGrid6.TextMatrix(0, 18) = "Depo No. of Cycle"
MSHFlexGrid6.TextMatrix(0, 19) = "Depo Rate per Cycle"
MSHFlexGrid6.TextMatrix(0, 20) = "Depo Storage Charges Amount"
MSHFlexGrid6.TextMatrix(0, 21) = "Total Storage Charges Amount"


MSHFlexGrid6.ColWidth(0) = 1200
MSHFlexGrid6.ColWidth(1) = 1200
MSHFlexGrid6.ColWidth(2) = 1000
MSHFlexGrid6.ColWidth(3) = 0
MSHFlexGrid6.ColWidth(4) = 2400
MSHFlexGrid6.ColWidth(5) = 0
MSHFlexGrid6.ColWidth(6) = 1400
MSHFlexGrid6.ColWidth(7) = 1200
MSHFlexGrid6.ColWidth(8) = 1200
MSHFlexGrid6.ColWidth(9) = 1200
MSHFlexGrid6.ColWidth(10) = 1000
MSHFlexGrid6.ColWidth(11) = 0
MSHFlexGrid6.ColWidth(12) = 1200
MSHFlexGrid6.ColWidth(13) = 1200
MSHFlexGrid6.ColWidth(14) = 1200
MSHFlexGrid6.ColWidth(15) = 1200
MSHFlexGrid6.ColWidth(16) = 1600
MSHFlexGrid6.ColWidth(17) = 1600
MSHFlexGrid6.ColWidth(18) = 1600
MSHFlexGrid6.ColWidth(19) = 1600
MSHFlexGrid6.ColWidth(20) = 1600

MSHFlexGrid6.RowHeight(0) = 600

End Sub

Private Sub DeleteCxTFBOE()
tmp = "Delete from Txn_CxTF_BOE Where CxTF_DetailsID=" & Val(TxtCxTFNo.Text) & " And LocationID=" & mLocationID
Call TmpTable
End Sub
Private Sub SaveCxTFBOE()

Call DeleteCxTFBOE

For I = 1 To MSHFlexGrid6.Rows - 2
    If MSHFlexGrid6.TextMatrix(I, 0) = "" Then Exit For
    Call TableTxn_CxTF_BOE(" where CxTF_DetailsID=" & Val(TxtCxTFNo.Text) & " And LocationID=" & mLocationID)
    If RsTxn_CxTF_BOE.RecordCount > 0 Then
       RsTxn_CxTF_BOE.MoveFirst
       RsTxn_CxTF_BOE.Find "ClientID= " & Val(MSHFlexGrid6.TextMatrix(I, 3))
    End If
    If RsTxn_CxTF_BOE.EOF = False Then
       'RsTxn_CxTF_BOE!CxTFDate = MSHFlexGrid6.TextMatrix(i, 3)
       RsTxn_CxTF_BOE!WithdrawalBags = IIf(IsNull(RsTxn_CxTF_BOE!WithdrawalBags), 0, RsTxn_CxTF_BOE!WithdrawalBags) + Val(MSHFlexGrid6.TextMatrix(I, 8))
       RsTxn_CxTF_BOE!WithdrawalQty = IIf(IsNull(RsTxn_CxTF_BOE!WithdrawalQty), 0, RsTxn_CxTF_BOE!WithdrawalQty) + Val(MSHFlexGrid6.TextMatrix(I, 9))
       RsTxn_CxTF_BOE!BOEAmount = IIf(IsNull(RsTxn_CxTF_BOE!BOEAmount), 0, RsTxn_CxTF_BOE!BOEAmount) + Val(MSHFlexGrid6.TextMatrix(I, 21))
       RsTxn_CxTF_BOE.Update
    Else
       RsTxn_CxTF_BOE.AddNew
'       RsTxn_CxTF_BOE!BOEID = getMaxBOEID
       RsTxn_CxTF_BOE!CxTF_DetailsID = TxtCxTFNo.Text
       RsTxn_CxTF_BOE!LocationID = mLocationID
       RsTxn_CxTF_BOE!CommodityID = mCommodityID
       RsTxn_CxTF_BOE!SM_Prefix = TxtPreFix.Text
       RsTxn_CxTF_BOE!ClientID = Val(MSHFlexGrid6.TextMatrix(I, 3))
       RsTxn_CxTF_BOE!CxTFDate = MSHFlexGrid6.TextMatrix(I, 0)
       RsTxn_CxTF_BOE!WithdrawalBags = Val(MSHFlexGrid6.TextMatrix(I, 8))
       RsTxn_CxTF_BOE!WithdrawalQty = Val(MSHFlexGrid6.TextMatrix(I, 9))
       RsTxn_CxTF_BOE!BOEAmount = Val(MSHFlexGrid6.TextMatrix(I, 21))
       RsTxn_CxTF_BOE!ExchangeTypeID = mExchangeTypeID
       RsTxn_CxTF_BOE!G_UID = GetGUID
       RsTxn_CxTF_BOE.Update
    End If
Next

End Sub


Private Function getMaxBOEID() As Double
Dim Retval As Double
tmp = "Select max(BOEID) from Txn_CxTF_BOE"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
getMaxBOEID = Retval
End Function


Private Sub CreateExcel()
Gen_mTimeStamp = GetTimeStamp
Call Create_Xls(Gen_mTimeStamp & Gen_UserName & "_CITF-BOE" & Trim(TxtPreFix) & "" & Trim(TxtCxTFNo) & ".xls")
Call Xls_Print_RptBOE
End Sub


Private Sub Xls_Print_RptBOE()

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
 
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "_CITF-BOE" & Trim(TxtPreFix) & "" & Trim(TxtCxTFNo) & ".xls")
 
I = 2
Set oWS = oWB.Worksheets("Sheet" & 1)
'oWS.PageSetup.RightMargin = 0.05
oWS.Name = TxtPreFix & "-" & TxtCxTFNo  'MSHFlexGrid2.TextMatrix(i, 23)
oWS.ClearArrows
mRow = 2

oWS.Range("a1:f1").Merge
oWS.Range("a:a").ColumnWidth = 10.86
oWS.Range("b:b").ColumnWidth = 12.86
oWS.Range("c:c").ColumnWidth = 12.86
oWS.Range("d:d").ColumnWidth = 12.86
oWS.Range("e:e").ColumnWidth = 17.86 '32.14
oWS.Range("f:f").ColumnWidth = 17.86

tmp = "a" & mRow & ":" & "b" & mRow
oWS.Range(tmp).Merge

oWS.Cells(mRow, 1) = "CMP Name"
oWS.Cells(mRow, 1).Font.Bold = True
tmp = "c" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, 3) = TxtCMPID.Text


tmp = "g" & mRow & ":" & "h" & mRow
oWS.Range(tmp).Merge

oWS.Cells(mRow, 7) = "Commodity"
oWS.Cells(mRow, 7).Font.Bold = True
tmp = "i" & mRow & ":" & "l" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, 9) = TxtCommodityID.Text

mRow = mRow + 1

tmp = "a" & mRow & ":" & "b" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, 1) = "Agent Name"
oWS.Cells(mRow, 1).Font.Bold = True
tmp = "c" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, 3) = TxtAgent.Text

tmp = "g" & mRow & ":" & "h" & mRow
oWS.Range(tmp).Merge

oWS.Cells(mRow, 7) = "CITF No"
oWS.Cells(mRow, 7).Font.Bold = True
tmp = "i" & mRow & ":" & "l" & mRow

oWS.Range(tmp).Merge
oWS.Cells(mRow, 9) = TxtCxTFNo.Text
oWS.Cells(mRow, 9).HorizontalAlignment = 2

mRow = mRow + 1

oWS.Cells(mRow, C + 1) = "CITF Date"
oWS.Cells(mRow, C + 2) = "CDTF Date"
oWS.Cells(mRow, C + 3) = "No of Days"
oWS.Cells(mRow, C + 4) = "Client Name"
oWS.Cells(mRow, C + 5) = "G-S-L"
oWS.Cells(mRow, C + 6) = "Lot Type"
oWS.Cells(mRow, C + 7) = "No of Bags Withdrawn"
oWS.Cells(mRow, C + 8) = "Withdrawn Qty"
oWS.Cells(mRow, C + 9) = "Billing Cycle"
oWS.Cells(mRow, C + 10) = "No of Cycle"
oWS.Cells(mRow, C + 11) = "Rate Per Cycle"
oWS.Cells(mRow, C + 12) = "Storage Charges Amount"

''    oWS.Cells(mRow, c + 13) = "Depo Billing Cycle"
''    oWS.Cells(mRow, c + 14) = "Depo No of Cycle"
''    oWS.Cells(mRow, c + 15) = "Depo Rate Per Cycle"
''    oWS.Cells(mRow, c + 16) = "Depo Storage Charges Amount"
''    oWS.Cells(mRow, c + 17) = "Total Storage Charges Amount"


tmp = "a" & mRow & ":q" & mRow
oWS.Range(tmp).Font.Bold = True

mRow = mRow + 1

For k = 1 To MSHFlexGrid6.Rows - 2

    oWS.Cells(mRow, C + 1) = MSHFlexGrid6.TextMatrix(k, 0)
    oWS.Cells(mRow, C + 2) = MSHFlexGrid6.TextMatrix(k, 1)
    oWS.Cells(mRow, C + 3) = MSHFlexGrid6.TextMatrix(k, 2)
    oWS.Cells(mRow, C + 4) = MSHFlexGrid6.TextMatrix(k, 4)
    oWS.Cells(mRow, C + 5) = MSHFlexGrid6.TextMatrix(k, 6)
    oWS.Cells(mRow, C + 6) = MSHFlexGrid6.TextMatrix(k, 7)
    oWS.Cells(mRow, C + 7) = MSHFlexGrid6.TextMatrix(k, 8)
    oWS.Cells(mRow, C + 8) = MSHFlexGrid6.TextMatrix(k, 9)
    oWS.Cells(mRow, C + 9) = MSHFlexGrid6.TextMatrix(k, 12)
    oWS.Cells(mRow, C + 10) = MSHFlexGrid6.TextMatrix(k, 14)
    oWS.Cells(mRow, C + 11) = MSHFlexGrid6.TextMatrix(k, 15)
    oWS.Cells(mRow, C + 12) = MSHFlexGrid6.TextMatrix(k, 16)

''    oWS.Cells(mRow, c + 13) = MSHFlexGrid6.TextMatrix(k, 17)
''    oWS.Cells(mRow, c + 14) = MSHFlexGrid6.TextMatrix(k, 18)
''    oWS.Cells(mRow, c + 15) = MSHFlexGrid6.TextMatrix(k, 19)
''    oWS.Cells(mRow, c + 16) = MSHFlexGrid6.TextMatrix(k, 20)
''    oWS.Cells(mRow, c + 17) = MSHFlexGrid6.TextMatrix(k, 21) 'Val(MSHFlexGrid6.TextMatrix(k, 16)) + Val(MSHFlexGrid6.TextMatrix(k, 20))

    mRow = mRow + 1
Next k
tmp = "a" & 2 & ":l" & mRow - 1

oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
    
If Gen_ExcelPassword <> "" Then
    oWS.Protect Gen_ExcelPassword, , , , , , True, True
End If
oWB.Save
'Call Log_Print_History(TxtPreFix, "CITF", TxtCxTFNo, Gen_FLName)
    
'oWB.Close
'oXL.Quit
MsgBox ("Excel file created!!!")
oXL.Visible = True

End Sub

'''Private Sub FillBillingDetails(Row_ As Variant, mFromDate As Date, mToDate As Date)
'''Dim mMonth, mYear As Variant
'''
'''If Val(MSHFlexGrid6.TextMatrix(Row_, 11)) <> 0 Then Exit Sub
'''
'''tmp = "Select SPRCID,MS.BillingCycleID,MB.BillingCycle,MB.NoofDays,BillingUnit,Amount,MS.DepoBillingCycleID "
'''tmp = tmp & " From Mst_SpecialRateCard MS"
'''tmp = tmp & " Inner Join Mst_BillingCycle MB on MS.BillingCycleID=MB.BillingCycleID"
'''tmp = tmp & " Where ClientIDRow = " & Val(MSHFlexGrid6.TextMatrix(Row_, 3)) & " And CommodityGroupID = " & mCommodityGroupId & " And LocationID =" & mLocationID & "  And ExchangeTypeID = 1 And MS.Flag = 'A'"
'''
'''Call Tmp2Table
'''If TmpRs2.RecordCount > 0 Then
'''   MSHFlexGrid6.TextMatrix(Row_, 12) = TmpRs2!BillingCycle
'''   MSHFlexGrid6.TextMatrix(Row_, 13) = IIf(TmpRs2!BillingUnit = "B", "Bag", "Qty")
'''
'''    'MSHFlexGrid6.TextMatrix(Row_, 14) = Val(MSHFlexGrid6.TextMatrix(Row_, 2)) / TmpRs2!NoofDays
'''    'MSHFlexGrid6.TextMatrix(Row_, 14) = CDbl(CLng(Val(MSHFlexGrid6.TextMatrix(Row_, 14)) + 0.5))
'''
'''   mLastDateofCDTF = GetLastDateOfMonth(mFromDate)
'''   mMonth = Month(mFromDate)
'''   mYear = Year(mFromDate)
'''
'''
'''   If mMonth = "12" Then
'''      mMonth = "01"
'''      mYear = mYear + 1
'''   Else
'''      mMonth = mMonth + 1
'''   End If
'''
'''
'''   mPostDepoFromDate = CDate(mMonth & "-" & "01" & "-" & mYear)
'''   MSHFlexGrid6.TextMatrix(Row_, 14) = CalculateNoOfCycle(mPostDepoFromDate, mToDate, TmpRs2!BillingCycleID, TmpRs2!NoofDays)
'''   MSHFlexGrid6.TextMatrix(Row_, 18) = CalculateDepoNoOfCycle(mFromDate, mLastDateofCDTF, TmpRs2!DepoBillingCycleID, TmpRs2!DepoNoofDays)
'''
''''   MSHFlexGrid6.TextMatrix(Row_, 14) = CalculateNoOfCycle(mFromDate, mToDate, TmpRs2!BillingCycleID, TmpRs2!NoofDays)
'''
'''   If TmpRs2!BillingUnit = "Q" Then
'''      MSHFlexGrid6.TextMatrix(Row_, 15) = TmpRs2!Amount
'''      MSHFlexGrid6.TextMatrix(Row_, 19) = TmpRs2!Amount
'''   Else
''''      tmp = "Select RatePerBag From Mst_SpecialRateCardDetail Where " & Val(MSHFlexGrid6.TextMatrix(Row_, 10)) & " <= ToBagSize And " & Val(MSHFlexGrid6.TextMatrix(Row_, 10)) & ">=FromBagSize And SPRCID=" & TmpRs2!SPRCID
''''      Call Tmp1Table
''''      If TmpRs1.RecordCount > 0 Then
''''         MSHFlexGrid6.TextMatrix(Row_, 15) = TmpRs1!RatePerBag
''''      End If
'''
'''      If DateDiff("M", mFromDate, mToDate) = 0 Then
'''         MSHFlexGrid6.TextMatrix(Row_, 15) = 0
'''      Else
'''         tmp = "Select RatePerBag From Mst_SpecialRateCardDetail Where " & Val(MSHFlexGrid6.TextMatrix(Row_, 10)) & " <= ToBagSize And " & Val(MSHFlexGrid6.TextMatrix(Row_, 10)) & ">=FromBagSize And SPRCID=" & TmpRs2!SPRCID
'''         Call Tmp1Table
'''         If TmpRs1.RecordCount > 0 Then
'''            MSHFlexGrid6.TextMatrix(Row_, 15) = TmpRs1!RatePerBag
'''         End If
'''
'''         tmp = " Select RatePerBag From Mst_SpecialRateCardDepoDetail Where " & Val(MSHFlexGrid6.TextMatrix(Row_, 10)) & " <= ToBagSize And " & Val(MSHFlexGrid6.TextMatrix(Row_, 10)) & " >= FromBagSize And  SPRCID = " & TmpRs2!SRCID & " "
'''         Call Tmp1Table
'''         If TmpRs1.RecordCount > 0 Then
'''            MSHFlexGrid6.TextMatrix(Row_, 19) = TmpRs1!RatePerBag
'''         End If
'''
'''      End If
'''
'''   End If
'''Else
'''
'''   tmp = "Select SRCID,MS.BillingCycleID,MB.BillingCycle,MB.NoofDays,MS.Flag,RateOnQty,MB1.BillingCycle as DepoBillingCycle,MS.DepoBillingCycleID,MB1.NoofDays as DepoNoOfDays"
'''   tmp = tmp & " From Mst_StorageRateCard MS "
'''   tmp = tmp & " Inner Join Mst_BillingCycle MB on MS.BillingCycleID=MB.BillingCycleID "
'''   tmp = tmp & " Inner Join Mst_BillingCycle MB1 on MS.DepoBillingCycleID=MB1.BillingCycleID "
'''   tmp = tmp & " Where CommodityGroupID = " & mCommodityGroupId & " And LocationID =" & mLocationID & "  And ExchangeTypeID = 1 And DefaultRate = 1 "
'''
'''   Call Tmp2Table
'''
'''   If TmpRs2.RecordCount > 0 Then
'''      MSHFlexGrid6.TextMatrix(Row_, 12) = TmpRs2!BillingCycle
'''      MSHFlexGrid6.TextMatrix(Row_, 17) = TmpRs2!DepoBillingCycle
'''      MSHFlexGrid6.TextMatrix(Row_, 13) = IIf(TmpRs2!Flag = "B", "Bag", "Qty")
'''       'MSHFlexGrid6.TextMatrix(Row_, 14) = Val(MSHFlexGrid6.TextMatrix(Row_, 2)) / TmpRs2!NoOfDays
'''       'MSHFlexGrid6.TextMatrix(Row_, 14) = CDbl(CLng(Val(MSHFlexGrid6.TextMatrix(Row_, 14)) + 0.5))
'''
'''      mLastDateofCDTF = GetLastDateOfMonth(mFromDate)
'''
'''
'''      mMonth = Month(mFromDate)
'''      mYear = Year(mFromDate)
'''
'''
'''      If mMonth = "12" Then
'''         mMonth = "01"
'''         mYear = mYear + 1
'''      Else
'''         mMonth = mMonth + 1
'''      End If
'''
'''
'''      mPostDepoFromDate = CDate(mMonth & "-" & "01" & "-" & mYear)
'''      MSHFlexGrid6.TextMatrix(Row_, 14) = CalculateNoOfCycle(mPostDepoFromDate, mToDate, TmpRs2!BillingCycleID, TmpRs2!NoofDays)
'''      MSHFlexGrid6.TextMatrix(Row_, 18) = CalculateDepoNoOfCycle(mFromDate, mLastDateofCDTF, TmpRs2!DepoBillingCycleID, TmpRs2!DepoNoofDays)
'''
'''      If TmpRs2!Flag = "Q" Then
'''         MSHFlexGrid6.TextMatrix(Row_, 15) = TmpRs2!RateOnQty
'''         MSHFlexGrid6.TextMatrix(Row_, 19) = TmpRs2!RateOnQty
'''      Else
''''         tmp = "Select RatePerBag From Mst_StorageRateCardDetail Where " & Val(MSHFlexGrid6.TextMatrix(Row_, 10)) & " <= ToBagSize And " & Val(MSHFlexGrid6.TextMatrix(Row_, 10)) & ">=FromBagSize And SRCID=" & TmpRs2!SRCID
''''         Call Tmp1Table
''''         If TmpRs1.RecordCount > 0 Then
''''            MSHFlexGrid6.TextMatrix(Row_, 15) = TmpRs1!RatePerBag
''''         End If
'''
'''         If DateDiff("M", mFromDate, mToDate) = 0 Then
'''            MSHFlexGrid6.TextMatrix(Row_, 15) = 0
'''         Else
'''            tmp = "Select RatePerBag From Mst_StorageRateCardDetail Where " & Val(MSHFlexGrid6.TextMatrix(Row_, 10)) & " <= ToBagSize And " & Val(MSHFlexGrid6.TextMatrix(Row_, 10)) & ">=FromBagSize And SRCID=" & TmpRs2!SRCID
'''            Call Tmp1Table
'''            If TmpRs1.RecordCount > 0 Then
'''               MSHFlexGrid6.TextMatrix(Row_, 15) = TmpRs1!RatePerBag
'''            End If
'''
'''            tmp = "Select RatePerBag From Mst_StorageRateCardDepoDetail Where " & Val(MSHFlexGrid6.TextMatrix(Row_, 10)) & " <= ToBagSize And " & Val(MSHFlexGrid6.TextMatrix(Row_, 10)) & ">=FromBagSize And SRCID=" & TmpRs2!SRCID
'''            Call Tmp1Table
'''            If TmpRs1.RecordCount > 0 Then
'''               MSHFlexGrid6.TextMatrix(Row_, 19) = TmpRs1!RatePerBag
'''            End If
'''         End If
'''
'''       End If
'''    End If
'''End If
'''End Sub

Private Function CalculateNoOfCycle(mFromDate_ As Date, mToDate_ As Date, mBillingCycleID_ As Variant, mNoOfDays_ As Variant)
Dim mCycleCount As Variant
Dim MonthCount, DepositMonthCycles, WithdrawalMonthCycles As Variant
'D=1 Daily
If mBillingCycleID_ = 1 Then
   mCycleCount = DateDiff("d", mFromDate_, mToDate_) + 1
'w=2 'Weekly
ElseIf mBillingCycleID_ = 2 Then
   MonthCount = DateDiff("M", mFromDate_, mToDate_) ' - 1
  
   WithdrawalMonthCycles = RoundUp(DatePart("d", mToDate_) / mNoOfDays_)
   mCycleCount = (MonthCount * 4) + WithdrawalMonthCycles '+ DepositMonthCycles
'F=3 'Fortnightly
   If mCycleCount > 4 Then
      mCycleCount = 4
   End If
ElseIf mBillingCycleID_ = 3 Then
   MonthCount = DateDiff("M", mFromDate_, mToDate_) '- 1
'   DepositMonthCycles = IIf(DatePart("d", mFromDate_) - mNoOfDays_ >= 0, 1, 2)
   WithdrawalMonthCycles = IIf(DatePart("d", mToDate_) - mNoOfDays_ >= 0, 2, 1)
   mCycleCount = (MonthCount * 2) + WithdrawalMonthCycles '+ DepositMonthCycles
   
   If Format(mToDate_, "dd") > 15 Then
      mCycleCount = 2
   Else
      mCycleCount = 1
   End If
   
'M=4 'Monthly
   If mCycleCount > 2 Then
      mCycleCount = 2
   End If
ElseIf mBillingCycleID_ = 4 Then
   mCycleCount = DateDiff("M", mFromDate_, mToDate_) + 1
ElseIf mBillingCycleID_ = 6 Then
   mCycleCount = "1" 'DateDiff("M", mFromDate_, mToDate_) + 1
End If

CalculateNoOfCycle = mCycleCount

End Function
Private Function CalculateDepoNoOfCycle(mFromDate_ As Date, mToDate_ As Date, mBillingCycleID_ As Variant, mNoOfDays_ As Variant)
Dim mCycleCount As Variant
Dim MonthCount, DepositMonthCycles, WithdrawalMonthCycles As Variant
'D=1 Daily
If mBillingCycleID_ = 1 Then
   mCycleCount = DateDiff("d", mFromDate_, mToDate_) + 1
'w=2 'Weekly
ElseIf mBillingCycleID_ = 2 Then
   MonthCount = DateDiff("M", mFromDate_, mToDate_) ' - 1
'   DepositMonthCycles = 4 - Left((DatePart("d", mFromDate_) / mNoOfDays_), 2)

   If Format(DatePart("d", mFromDate_) / mNoOfDays_) = 1 Then
      DepositMonthCycles = 4 '- 0 'Left(Format(DatePart("d", mFromDate_) / mNoOfDays_), 1)
   ElseIf Format(DatePart("d", mFromDate_) / mNoOfDays_) > 4 Then
      DepositMonthCycles = 4 - (Left(Format(DatePart("d", mFromDate_) / mNoOfDays_), 1) - 1)
   Else
      If DatePart("d", mFromDate_) Mod mNoOfDays_ = 0 Then
         DepositMonthCycles = 4 - Left(Format(DatePart("d", mFromDate_) / mNoOfDays_), 1) + 1
      Else
         DepositMonthCycles = 4 - Left(Format(DatePart("d", mFromDate_) / mNoOfDays_), 1)
      End If
'      DepositMonthCycles = 4 - Left(Format(DatePart("d", mFromDate_) / mNoOfDays_), 1)
   End If
   'WithdrawalMonthCycles = RoundUp(DatePart("d", mToDate_) / mNoOfDays_)
   mCycleCount = (MonthCount * 4) + DepositMonthCycles
'F=3 'Fortnightly
   If mCycleCount > 4 Then
      mCycleCount = 4
   End If
ElseIf mBillingCycleID_ = 3 Then
   MonthCount = DateDiff("M", mFromDate_, mToDate_) '- 1
   DepositMonthCycles = IIf(DatePart("d", mFromDate_) - mNoOfDays_ >= 0, 1, 2)
'   WithdrawalMonthCycles = IIf(DatePart("d", mToDate_) - mNoOfDays_ >= 0, 2, 1)
   mCycleCount = (MonthCount * 2) + DepositMonthCycles
'M=4 'Monthly
   If mCycleCount > 2 Then
      mCycleCount = 2
   End If
ElseIf mBillingCycleID_ = 4 Then
   mCycleCount = DateDiff("M", mFromDate_, mToDate_) + 1
End If

CalculateDepoNoOfCycle = mCycleCount

End Function
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

Private Sub TxtTTDTruckNo_LostFocus()
TxtTTDTruckNo.Text = Trim$(TxtTTDTruckNo.Text)
End Sub

Private Sub TxtWeighBridgNo_LostFocus()
TxtWeighBridgNo.Text = Trim$(TxtWeighBridgNo.Text)
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

'Private Sub PDF_Print_Rpt()
'Set objPDF = New mjwPDF
'
'' Create a simple PDF file using the mjwPDF class
'
'Gen_mTimeStamp = GetTimeStamp
''Call Create_Xls(Gen_mTimeStamp & Gen_UserName & "_CDTF-" & Trim(TxtPreFix) & "" & Trim(TxtCxTFNo) & ".xls")
'
'
'' Set the PDF title and filename
'objPDF.PDFTitle = "CDTF PDF Document"
'objPDF.PDFFileName = App.Path & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "_CITF-" & Trim(TxtPreFix) & "" & Trim(TxtCxTFNo) & ".pdf"
'mPDFFileName = App.Path & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "_CITF-" & Trim(TxtPreFix) & "" & Trim(TxtCxTFNo) & ".pdf"
'' We must tell the class where the PDF fonts are located
'objPDF.PDFLoadAfm = App.Path & "\Fonts"
'
'' Set the file properties
'objPDF.PDFSetLayoutMode = LAYOUT_DEFAULT
'objPDF.PDFFormatPage = FORMAT_A4
'objPDF.PDFOrientation = ORIENT_PORTRAIT
'objPDF.PDFSetUnit = UNIT_PT
'objPDF.PDFSetViewerPreferences = VIEW_HIDEMENUBAR & VIEW_HIDETOOLBAR & VIEW_HIDEWINDOWUI & VIEW_HIDERIGHTCLICK '& VIEW_DISABLEFILEMENU
'
'
'' Lets us set see the bookmark pane when we view the PDF
''objPDF.PDFUseOutlines = True
'
'' View the PDF file after we create it
'objPDF.PDFView = True
'
'' Begin our PDF document
'
'objPDF.PDFBeginDoc
'    ' Lets add a heading
'objPDF.PDFSetFont FONT_ARIAL, 12, FONT_BOLD
'objPDF.PDFSetDrawColor = vbWhite
'objPDF.PDFSetTextColor = vbBlack
'objPDF.PDFSetAlignement = ALIGN_CENTER
'objPDF.PDFSetBorder = BORDER_ALL
'objPDF.PDFSetFill = True
'
'objPDF.PDFCell " ", 15, 15, _
'objPDF.PDFGetPageWidth - 30, objPDF.PDFGetPageHeight - 20
'
'If PntPrv = "Y" Then
'   objPDF.PDFWatermark = "Preview"
'Else
'   objPDF.PDFWatermark = "NCMSL"
'End If
'
'objPDF.PDFSetFont FONT_ARIAL, 12, FONT_BOLD
'objPDF.PDFSetDrawColor = vbWhite
'objPDF.PDFSetTextColor = vbBlack
'objPDF.PDFSetAlignement = ALIGN_CENTER
'objPDF.PDFSetBorder = BORDER_ALL
'objPDF.PDFSetFill = True
'
'
'pgval = ((objPDF.PDFGetPageWidth - 18))
'
''objPDF.PDFDrawPolygon Array(12, 15, objPDF.PDFGetPageWidth - 18, 15, objPDF.PDFGetPageWidth - 18, 30, 12, 30)
'
'mRowHeight = 15
'mGlobalLeft = 12
''objPDF.PDFDrawPolygon Array(12, 15, objPDF.PDFGetPageWidth - 18, 15, objPDF.PDFGetPageWidth - 18, 30, 12, 30)
'mLeft = mGlobalLeft
'mTop = 15
'mHeight = mTop + mRowHeight
'mTxtFromTop = mRowHeight + 2
'
' 'Call DrawTable(10, 2, 15, 30, 12)
'
'Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
'
'mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
'mStrLength = Len("NCMSL " & TxtCMPID & "")
'mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
'mTxtLeft = mPrintPoint - (mStrLength * 2) '* 5) '270
'mFontSize = 12
'
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_CENTER, mTxtLeft, mTxtTop, "NCMSL " & TxtCMPID.Text & "")
'
'mHeight = mTop + mRowHeight
'mLeft = mGlobalLeft
'
'
'Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
'
'mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
'mStrLength = Len("Commodity Issue Transaction Form (CITF)")
'mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
'mTxtLeft = mPrintPoint - (mStrLength * 2) '* 5) '270
'mFontSize = 10
'
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_CENTER, mTxtLeft, mTxtTop, "Commodity Issue Transaction Form (CITF)")
'
'mHeight = mTop + mRowHeight
'mLeft = mGlobalLeft
'
'Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
'
'
'mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
'mStrLength = Len("(This is NOT an instrument of title of the goods)")
'mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
'mTxtLeft = mPrintPoint '- (mStrLength) '* 5) '270
'mFontSize = 6
'
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_CENTER, mTxtLeft - 30, mTxtTop - 3, "(This is NOT an instrument of title of the goods)")
'
''--Code added By Kishor
''Row 1
'
'
'mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
'mStrLength = Len("Depositor Details")
'mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
'mTxtLeft = mPrintPoint '- (mStrLength) '* 5) '270
'mFontSize = 7
'Dim mDispTextSize  As Integer
'mDispTextSize = 6
'
'mTxtTop = mTxtTop + mRowHeight
'
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "CITF No")
'Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop, TxtCxTFNo)
''--
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 280, mTxtTop, "Date")
'Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 430, mTxtTop, Format(TxtCxTFDate.Value, "dd-MMM-yyyy"))
''--
'mTxtTop = mTxtTop + mRowHeight
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "Location")
'Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop, GetLocationFromCMP(Val(mCmpID)))
''--
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 280, mTxtTop, "State")
'Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 430, mTxtTop, GetStateFromLocation(GetLocationIDFromCMP(Val(mCmpID))))
''--
'mTxtTop = mTxtTop + mRowHeight
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "CMPID & Address")
'Dim mCMPIDnAddr As String
'mCMPIDnAddr = mCmpID & "-" & GetCMPAddress(Val(mCmpID), "A")
'Dim mLen As Integer, mMax As Integer, mPara As Integer
'Dim t1 As Double
'mMax = 35
'mPara = 0
't1 = 1
'mLen = Len(mCMPIDnAddr)
'mLen = Round(mLen / mMax) '+ 1
'
'For t = 1 To mLen
'    If t > 1 Then
'       t1 = mMax + t1
'       'mMax '= mMax + 35
'       mPara = mPara + 7
'    End If
'    Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop + mPara, Mid$(mCMPIDnAddr, t1, mMax))
'Next
'
''--
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 280, mTxtTop, "Commodity")
'Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 430, mTxtTop, TxtCommodityID)
'
'Call DrawTable(2, 4, mTop, mHeight, mLeft, mRowHeight)
'mHeight = mHeight + 15
'Call DrawTable(, 4, mTop, mHeight, mLeft, mRowHeight)
'
'
''Row 2
'
'mTxtTop = mHeight - (mTxtFromTop) + 15 'mTop ' + 15
'mStrLength = Len("Transaction Tyepe")
'mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
'mTxtLeft = mPrintPoint '- (mStrLength) '* 5) '270
'mFontSize = 7
'
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "Transaction Type")
'Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop, CmbExchangeTypeID)
'
'mTxtTop = mTxtTop + mRowHeight
'
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "Name Of Agent")
'Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop, TxtAgent)
'
'objPDF.PDFDrawLine mGlobalLeft + 132, mTop - 4 * mRowHeight, mGlobalLeft + 132, mHeight + mRowHeight
'Call DrawTable(2, , mTop, mHeight, mLeft, mRowHeight)
'
'
''Row 3
'
'mTxtTop = mHeight '- (mTxtFromTop) 'mTop ' + 15
'mStrLength = Len("Withdrawer Details")
'mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
'mTxtLeft = mPrintPoint '- (mStrLength) '* 5) '270
'mFontSize = 8
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mLeft, mTxtTop - 2, "Withdrawer Details")
'
'Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
'
'objPDF.PDFDrawLine mGlobalLeft + 72, mTop, mGlobalLeft + 72, mTop + ((MSHFlexGrid2.Rows - 2) * mRowHeight)
'objPDF.PDFDrawLine mGlobalLeft + 142, mTop, mGlobalLeft + 142, mTop + ((MSHFlexGrid2.Rows - 2) * mRowHeight)
'If mExchangeTypeID <> 2 Then
'   objPDF.PDFDrawLine 2.8 * (mGlobalLeft + 142), mTop, 2.8 * (mGlobalLeft + 142), mTop + ((MSHFlexGrid2.Rows - 1) * mRowHeight)
'End If
'objPDF.PDFDrawLine 3.3 * (mGlobalLeft + 142), mTop, 3.3 * (mGlobalLeft + 142), mTop + ((MSHFlexGrid2.Rows - 1) * mRowHeight)
'
'Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
'
'
'mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
'mStrLength = Len("Depositor Details")
'mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
'mTxtLeft = mPrintPoint '- (mStrLength) '* 5) '270
'mFontSize = 7
'
'If CmbExchangeTypeID.Text = "NCDEX" Then
'   If Check1.Value = 0 Then
'      Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mLeft, mTxtTop, "RRN No")
'   Else
'
'      Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mLeft, mTxtTop, "Rejected Withdrawal")
'   End If
'ElseIf CmbExchangeTypeID.Text = "Non NCDEX" Then
'
'   Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mLeft, mTxtTop, "WHR No")
'ElseIf CmbExchangeTypeID.Text = "SPOT" Then
'
'   Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mLeft, mTxtTop, "CDC No")
'End If
'
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mLeft + 80, mTxtTop, "Client ID")
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 150, mTxtTop, "Client Name")
'If mExchangeTypeID <> 2 Then
'   Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 430, mTxtTop, "No of Bags")
'   Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 500, mTxtTop - 6, "Electronic Weight")
'   Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 500, mTxtTop, "in MT")
'
'Else
'   Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 500, mTxtTop - 6, "Electronic Weight")
'   Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 500, mTxtTop, "in MT")
'
'End If
'
'objPDF.PDFDrawLine mGlobalLeft + 72, mTop, mGlobalLeft + 72, mTop + ((MSHFlexGrid2.Rows - 2) * mRowHeight)
'objPDF.PDFDrawLine mGlobalLeft + 142, mTop, mGlobalLeft + 142, mTop + ((MSHFlexGrid2.Rows - 2) * mRowHeight)
'If mExchangeTypeID <> 2 Then
'   objPDF.PDFDrawLine 2.8 * (mGlobalLeft + 142), mTop, 2.8 * (mGlobalLeft + 142), mTop + ((MSHFlexGrid2.Rows - 1) * mRowHeight)
'End If
'objPDF.PDFDrawLine 3.3 * (mGlobalLeft + 142), mTop, 3.3 * (mGlobalLeft + 142), mTop + ((MSHFlexGrid2.Rows - 1) * mRowHeight)
'
'Call DrawTable(MSHFlexGrid2.Rows - 1, , mTop, mHeight, mLeft, mRowHeight)
'
''FILL GRID WITH DETAILS
'For t = 2 To MSHFlexGrid2.Rows - 1
'    mTxtTop = mTxtTop + mRowHeight
'    Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mLeft, mTxtTop, MSHFlexGrid2.TextMatrix(t, 1))
'    Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 80, mTxtTop, MSHFlexGrid2.TextMatrix(t, 3))
'    Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 150, mTxtTop, MSHFlexGrid2.TextMatrix(t, 4))
'    If mExchangeTypeID <> 2 Then
'       Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 430, mTxtTop, Format(MSHFlexGrid2.TextMatrix(t, 15), "#######0.000"))
'    End If
'    Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 500, mTxtTop, Format(MSHFlexGrid2.TextMatrix(t, 9), "#######0.000"))
'
'Next
'mTxtTop = mTxtTop + mRowHeight
'
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 220, mTxtTop, "Total")
'If mExchangeTypeID <> 2 Then
'   Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 430, mTxtTop, Format(MSHFlexGrid2.TextMatrix(0, 15), "#######0.000"))
'End If
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 500, mTxtTop, Format(MSHFlexGrid2.TextMatrix(0, 9), "#######0.000"))
'Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
'
'mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
'mStrLength = Len("Depositor Details")
'mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
'mTxtLeft = mPrintPoint '- (mStrLength) '* 5) '270
'mFontSize = 7
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mLeft, mTxtTop - 2, "Truck-wise Withdrawal Details")
'''
'''objPDF.PDFDrawLine mGlobalLeft + 50, mTop, mGlobalLeft + 50, mTop + ((MSHFlexGrid3.Rows - 1) * mRowHeight)
'''objPDF.PDFDrawLine (mGlobalLeft + 120), mTop, (mGlobalLeft + 120), mTop + ((MSHFlexGrid3.Rows - 1) * mRowHeight)
'''objPDF.PDFDrawLine (mGlobalLeft + 190), mTop, (mGlobalLeft + 190), mTop + ((MSHFlexGrid3.Rows + 6) * mRowHeight)
'''objPDF.PDFDrawLine (mGlobalLeft + 256), mTop, (mGlobalLeft + 256), mTop + ((MSHFlexGrid3.Rows + 6) * mRowHeight)
'''objPDF.PDFDrawLine (mGlobalLeft + 356), mTop, (mGlobalLeft + 356), mTop + ((MSHFlexGrid3.Rows - 1) * mRowHeight)
'''
'''Dim mTempTop As Integer
'''mTempTop = mTxtTop - 3
'''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 90, mTempTop + (MSHFlexGrid3.Rows * mRowHeight), "Total")
'''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 60, mTempTop + ((MSHFlexGrid3.Rows + 1) * mRowHeight), "Less Gunny Weight(MT)")
'''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 50, mTempTop + ((MSHFlexGrid3.Rows + 2) * mRowHeight), "Less Standard Deductions(MT)")
'''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 60, mTempTop + ((MSHFlexGrid3.Rows + 3) * mRowHeight), "Add Adjustments(MT)")
'''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 60, mTempTop + ((MSHFlexGrid3.Rows + 4) * mRowHeight), "Electronic Weight(MT)")
''''pp 'Hologram No
'''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 270, mTempTop + ((MSHFlexGrid3.Rows + 4) * mRowHeight), "Hologram No")
'''
''''Actual Values
'''Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 200, mTempTop + (MSHFlexGrid3.Rows * mRowHeight), MSHFlexGrid3.TextMatrix(0, 5))
'''Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 200, mTempTop + ((MSHFlexGrid3.Rows + 1) * mRowHeight), MSHFlexGrid4.TextMatrix(0, 8))
'''Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 200, mTempTop + ((MSHFlexGrid3.Rows + 2) * mRowHeight), MSHFlexGrid4.TextMatrix(0, 9))
'''Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 200, mTempTop + ((MSHFlexGrid3.Rows + 3) * mRowHeight), MSHFlexGrid4.TextMatrix(0, 6))
'''Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 200, mTempTop + ((MSHFlexGrid3.Rows + 4) * mRowHeight), MSHFlexGrid4.TextMatrix(0, 10))
''''Hologram No -- Use Record Set
'''
'''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 370, mTempTop + ((MSHFlexGrid3.Rows + 4) * mRowHeight), " " & RsTxn_CxTF_Details!HologramNo & " ")
'''
'''Dim mTemp As Integer
'''For t = 1 To MSHFlexGrid3.Rows - 1
'''    If t = 1 Then
'''       mTemp = mFontSize
'''    Else
'''       mTemp = mDispTextSize
'''    End If
'''    mTxtTop = mTxtTop + mRowHeight
'''    Call PrintText(mTemp, FONT_NORMAL, ALIGN_LEFT, mLeft, mTxtTop, MSHFlexGrid3.TextMatrix(t, 1))
'''    Call PrintText(mTemp, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 60, mTxtTop, MSHFlexGrid3.TextMatrix(t, 3))
'''    Call PrintText(mTemp, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 130, mTxtTop, MSHFlexGrid3.TextMatrix(t, 4))
'''    Call PrintText(mTemp, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 200, mTxtTop, MSHFlexGrid3.TextMatrix(t, 5))
'''    Call PrintText(mTemp, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 256, mTxtTop, MSHFlexGrid3.TextMatrix(t, 2))
'''    Call PrintText(mTemp, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 356, mTxtTop, MSHFlexGrid3.TextMatrix(t, 6))
'''Next
'''
'''Call DrawTable(MSHFlexGrid3.Rows + 6, , mTop, mHeight, mLeft, mRowHeight)
''''Clear Line
'''mTxtTop = mHeight - (mTxtFromTop)
'''objPDF.PDFSetTextColor = vbWhite
'''
'''Call PrintText(15, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 255.2, mTxtTop - (6 * mRowHeight), "_____________________________________")
'''Call PrintText(15, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 255.2, mTxtTop - (5 * mRowHeight), "_____________________________________")
'''Call PrintText(15, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 255.2, mTxtTop - (4 * mRowHeight), "_____________________________________")
'''
'''objPDF.PDFSetTextColor = vbBlack
'''Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 256, mTxtTop - (3 * mRowHeight), "Affix Hologram Here")
''''--
'''
'''objPDF.PDFDrawLine (mGlobalLeft + 50), mTop - (2 * mRowHeight), (mGlobalLeft + 50), mTop
'''objPDF.PDFDrawLine (mGlobalLeft + 120), mTop - (2 * mRowHeight), (mGlobalLeft + 120), mTop
'''
'''objPDF.PDFDrawLine (mGlobalLeft + 360), mTop - (3 * mRowHeight), (mGlobalLeft + 360), mTop
'''objPDF.PDFDrawLine (mGlobalLeft + 420), mTop - (2 * mRowHeight), (mGlobalLeft + 420), mTop
'''
'''If mHeight > objPDF.PDFGetPageHeight - 50 Then
'''   objPDF.PDFEndPage
'''   objPDF.PDFNewPage
'''   objPDF.PDFSetFont FONT_ARIAL, 12, FONT_BOLD
'''   objPDF.PDFSetDrawColor = vbWhite
'''   objPDF.PDFSetTextColor = vbBlack
'''   objPDF.PDFSetAlignement = ALIGN_CENTER
'''   objPDF.PDFSetBorder = BORDER_ALL
'''   objPDF.PDFSetFill = True
'''
'''   objPDF.PDFCell " ", 15, 15, _
'''   objPDF.PDFGetPageWidth - 30, objPDF.PDFGetPageHeight - 20
'''   mLeft = mGlobalLeft
'''   mTop = 15
'''   mHeight = mTop + mRowHeight
'''   mTxtFromTop = mRowHeight + 2
'''End If
'''
'''Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
'''mTxtTop = mHeight - (mTxtFromTop)
'''
'''Call PrintText(5.4, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "NOTE: THIS DOCUMENT IS NOT VALID WITHOUT A HOLOGRAM OF NCMSL AFFIXED ABOVE & THE SERIAL NUMBER IN THE HOLOGRAM SHOULD MATCH WITH THE HOLOGRAM NUMBER PRINTED ABOVE.")
'''
'''mHeight = mHeight + 50
'''
'''Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
'''mTxtTop = mHeight - (mTxtFromTop) '+ 15
'''
'''Call PrintText(9, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "Name & Signature of CMP Manager/CMP DRC:")
'''
'''mHeight = mHeight + 30
'''Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
'''mTxtTop = mHeight - (mTxtFromTop)
'''
'''Call PrintText(6.5, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop - 17, "I/We hereby declare that the commodities being deposited herein belong to the clients as stated above and are free from any encumbrances. I/We request  you to take custody of the said")
'''Call PrintText(6.5, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop - 10, "commodities and I/We undertake to indemnify you against any losses incurred by you due to any false information provided by me/us. I/We also undertake to pay all required charges as per")
'''Call PrintText(6.5, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop - 3, "your charges schedule from time to time. I have gone throught the attached 'Terms and Conditions'** of storage and agree to abide by the same.")
'''
'
'objPDF.PDFDrawLine mGlobalLeft + 50, mTop, mGlobalLeft + 50, mTop + ((MSHFlexGrid3.Rows - 1) * mRowHeight)
'objPDF.PDFDrawLine (mGlobalLeft + 120), mTop, (mGlobalLeft + 120), mTop + ((MSHFlexGrid3.Rows - 1) * mRowHeight)
'objPDF.PDFDrawLine (mGlobalLeft + 190), mTop, (mGlobalLeft + 190), mTop + ((MSHFlexGrid3.Rows) * mRowHeight)
'objPDF.PDFDrawLine (mGlobalLeft + 256), mTop, (mGlobalLeft + 256), mTop + ((MSHFlexGrid3.Rows + 6) * mRowHeight)
'objPDF.PDFDrawLine (mGlobalLeft + 356), mTop, (mGlobalLeft + 356), mTop + ((MSHFlexGrid3.Rows - 1) * mRowHeight)
''''
''''
'''''Dim mTempTop As Integer
'''''mTempTop = mTxtTop - 3
'''''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 90, mTempTop + (MSHFlexGrid3.Rows * mRowHeight), "Total")
''''
'''''----Commented on 18-01-2011 by Sagar
''''
'''''''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 60, mTempTop + ((MSHFlexGrid3.Rows + 1) * mRowHeight), "Less Gunny Weight(MT)")
'''''''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 50, mTempTop + ((MSHFlexGrid3.Rows + 2) * mRowHeight), "Less Standard Deductions(MT)")
'''''''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 60, mTempTop + ((MSHFlexGrid3.Rows + 3) * mRowHeight), "Add Adjustments(MT)")
'''''''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 60, mTempTop + ((MSHFlexGrid3.Rows + 4) * mRowHeight), "Electronic Weight(MT)")
''''
''''objPDF.PDFDrawLine (mGlobalLeft), mTempTop + (MSHFlexGrid3.Rows + 4 * mRowHeight), (mGlobalLeft + 256), mTempTop + (MSHFlexGrid3.Rows + 4 * mRowHeight)
''''
'''''pp 'Hologram No
''''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 270, mTempTop + ((MSHFlexGrid3.Rows + 4) * mRowHeight), "Hologram No")
''''
'''''Actual Values
''''Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 200, mTempTop + (MSHFlexGrid3.Rows * mRowHeight), MSHFlexGrid3.TextMatrix(0, 5))
''''
'''''----Commented on 18-01-2011 by Sagar
''''
'''''''Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 200, mTempTop + ((MSHFlexGrid3.Rows + 1) * mRowHeight), MSHFlexGrid4.TextMatrix(0, 8))
'''''''Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 200, mTempTop + ((MSHFlexGrid3.Rows + 2) * mRowHeight), MSHFlexGrid4.TextMatrix(0, 9))
'''''''Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 200, mTempTop + ((MSHFlexGrid3.Rows + 3) * mRowHeight), MSHFlexGrid4.TextMatrix(0, 6))
'''''''Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 200, mTempTop + ((MSHFlexGrid3.Rows + 4) * mRowHeight), MSHFlexGrid4.TextMatrix(0, 10))
'''''Hologram No -- Use Record Set
''''
''''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 370, mTempTop + ((MSHFlexGrid3.Rows + 4) * mRowHeight), " " & RsTxn_CxTF_Details!HologramNo & " ")
''''
'Dim mTemp As Integer
'For t = 1 To MSHFlexGrid3.Rows - 1
'    If t = 1 Then
'       mTemp = mFontSize
'    Else
'       mTemp = mDispTextSize
'    End If
'    mTxtTop = mTxtTop + mRowHeight
'    Call PrintText(mTemp, FONT_NORMAL, ALIGN_LEFT, mLeft, mTxtTop, MSHFlexGrid3.TextMatrix(t, 1))
'    Call PrintText(mTemp, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 60, mTxtTop, MSHFlexGrid3.TextMatrix(t, 3))
'    Call PrintText(mTemp, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 130, mTxtTop, MSHFlexGrid3.TextMatrix(t, 4))
'    Call PrintText(mTemp, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 200, mTxtTop, MSHFlexGrid3.TextMatrix(t, 5))
'    Call PrintText(mTemp, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 256, mTxtTop, MSHFlexGrid3.TextMatrix(t, 2))
'    Call PrintText(mTemp, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 356, mTxtTop, MSHFlexGrid3.TextMatrix(t, 6))
'Next
'
'Call DrawTable(MSHFlexGrid3.Rows - 1, , mTop, mHeight, mLeft, mRowHeight)
'
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 90, mTop + mRowHeight - 3, "Total")
'Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 200, mTop + mRowHeight - 3, MSHFlexGrid3.TextMatrix(0, 5))
'objPDF.PDFDrawLine (mGlobalLeft), mTop + mRowHeight, (mGlobalLeft + 256), mTop + mRowHeight
'mTop = mTop + (mRowHeight * 4)
'
'objPDF.PDFDrawLine mGlobalLeft + 50, mTop + mRowHeight, mGlobalLeft + 50, mTop + (mRowHeight * 3)
'objPDF.PDFDrawLine (mGlobalLeft + 120), mTop + mRowHeight, (mGlobalLeft + 120), mTop + (mRowHeight * 3)
'objPDF.PDFDrawLine (mGlobalLeft + 190), mTop + mRowHeight, (mGlobalLeft + 190), mTop + (mRowHeight * 3)
'objPDF.PDFDrawLine (mGlobalLeft + 256), mTop, (mGlobalLeft + 256), mTop + (mRowHeight * 3)
'objPDF.PDFDrawLine (mGlobalLeft + 356), mTop, (mGlobalLeft + 356), mTop + (mRowHeight * 3)
'objPDF.PDFDrawLine (mGlobalLeft + 420), mTop + mRowHeight, (mGlobalLeft + 420), mTop + (mRowHeight * 3)
'
'Call PrintText(6, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 265, mTop - 3, "Affix Hologram Here")
'
'If mExchangeTypeID = 2 Then
'   Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 70, mTop + mRowHeight - 3, "Total No. of Bags")
'   Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 200, mTop + mRowHeight - 3, Format(MSHFlexGrid4.TextMatrix(0, 6)))
'   objPDF.PDFDrawLine (mGlobalLeft + 190), mTop, (mGlobalLeft + 190), mTop + mRowHeight
'End If
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 270, mTop + mRowHeight - 3, "Hologram No")
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 370, mTop + mRowHeight - 3, " " & RsTxn_CxTF_Details!HologramNo & " ")
'
'Call DrawTable(4, , mTop, mTop, mLeft, mRowHeight)
'
'If mHeight > objPDF.PDFGetPageHeight - 50 Then
'   objPDF.PDFEndPage
'   objPDF.PDFNewPage
'   objPDF.PDFSetFont FONT_ARIAL, 12, FONT_BOLD
'   objPDF.PDFSetDrawColor = vbWhite
'   objPDF.PDFSetTextColor = vbBlack
'   objPDF.PDFSetAlignement = ALIGN_CENTER
'   objPDF.PDFSetBorder = BORDER_ALL
'   objPDF.PDFSetFill = True
'
'   objPDF.PDFCell " ", 15, 15, _
'   objPDF.PDFGetPageWidth - 30, objPDF.PDFGetPageHeight - 20
'   mLeft = mGlobalLeft
'   mTop = 15
'   mHeight = mTop + mRowHeight
'   mTxtFromTop = mRowHeight + 2
'End If
'mHeight = mTop
'Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
'mTxtTop = mHeight - (mTxtFromTop)
''''
'Call PrintText(5.4, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "NOTE: THIS DOCUMENT IS NOT VALID WITHOUT A HOLOGRAM OF NCMSL AFFIXED ABOVE & THE SERIAL NUMBER IN THE HOLOGRAM SHOULD MATCH WITH THE HOLOGRAM NUMBER PRINTED ABOVE.")
''''
'mHeight = mHeight + 50
''''
'Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
'mTxtTop = mHeight - (mTxtFromTop) '+ 15
''''
'Call PrintText(9, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "Name & Signature of CMP Manager/CMP DRC:")
'
'Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
'
'mTxtTop = mTxtTop + mRowHeight
'Call PrintText(7, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "Received the above mentioned goods in good condition.")
''mHeight = mHeight + 30
''Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
'
''''
'''Call PrintText(6.5, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop - 17, "I/We hereby declare that the commodities being deposited herein belong to the clients as stated above and are free from any encumbrances. I/We request  you to take custody of the said")
'''Call PrintText(6.5, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop - 10, "commodities and I/We undertake to indemnify you against any losses incurred by you due to any false information provided by me/us. I/We also undertake to pay all required charges as per")
'''Call PrintText(6.5, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop - 3, "your charges schedule from time to time. I have gone throught the attached 'Terms and Conditions'** of storage and agree to abide by the same.")
''''''
'
'mHeight = mHeight + 70
'mTxtTop = mHeight - (mTxtFromTop)
'Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
'
'Call PrintText(9, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "Name & Signature of Withdrawer/Agent of Withdrawer:")
'
'mHeight = mHeight + 50
''Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
'mTxtTop = mHeight - (mTxtFromTop)
'Call PrintText(8, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "**'Terms and Conditions' can be viewed in http://www.ncmsl.com/whrtNc.html")
'
'If mHeight > objPDF.PDFGetPageHeight Then
'   objPDF.PDFEndPage
'   objPDF.PDFNewPage
'   objPDF.PDFSetFont FONT_ARIAL, 12, FONT_BOLD
'   objPDF.PDFSetDrawColor = vbWhite
'   objPDF.PDFSetTextColor = vbBlack
'   objPDF.PDFSetAlignement = ALIGN_CENTER
'   objPDF.PDFSetBorder = BORDER_ALL
'   objPDF.PDFSetFill = True
'
'   objPDF.PDFCell " ", 15, 15, _
'   objPDF.PDFGetPageWidth - 30, objPDF.PDFGetPageHeight - 20
'   mLeft = mGlobalLeft
'   mTop = 15
'   mHeight = mTop + mRowHeight
'   mTxtFromTop = mRowHeight + 2
'End If
'
'objPDF.PDFEndDoc
'
'End Sub
Private Sub DrawTable(Optional mROw_ As Double, Optional mCol_ As Double, Optional mTop_ As Double, Optional mHeight_ As Double, Optional mLeft_ As Double, Optional mRowHeight_ As Double)

If mROw_ = 0 Then mROw_ = 1
If mCol_ = 0 Then mCol_ = 1
If mLeft_ = 0 Then mLeft_ = mmasterLeft
If mTop_ = 0 Then mTop_ = 0
If mHeight_ = 0 Then mHeight_ = mTop_ + mRowHeight_

Dim pgval_ As Double

For ir = 1 To mROw_
    pgval_ = pgval
    pgval_ = pgval_ / mCol_
    mWidth_ = pgval_
    For ic = 1 To mCol_
        objPDF.PDFDrawPolygon Array(mLeft_, mTop_, mWidth_, mTop_, mWidth_, mHeight_, mLeft_, mHeight_)
        mLeft_ = pgval_ * ic
        mWidth_ = pgval_ * (ic + 1)
    Next
    mTop_ = mHeight_
    mHeight_ = mHeight_ + mRowHeight_
    'mLeft_ = mmasterLeft

    mLeft_ = mGlobalLeft

Next

mTop = mTop_
mRow = mROw_ + 1

End Sub
Private Sub PrintText(mFontSize_ As Integer, mFontFormat_ As PDFFontStl, mAlign_ As Variant, mTxtLeft_ As Double, mTxtTop_ As Double, mTexttoPrint As String)

'        objPDF.PDFSetFont FONT_ARIAL, 12, FONT_BOLD
'        objPDF.PDFSetAlignement = ALIGN_CENTER
        
objPDF.PDFSetFont FONT_ARIAL, mFontSize_, mFontFormat_
objPDF.PDFSetAlignement = mAlign_
objPDF.PDFTextOut mTexttoPrint, mTxtLeft_, mTxtTop_


End Sub

Private Sub Gen_Email()
Dim TOADD As Variant
Dim TOCC, mMsg, mGridMsg As Variant
Dim strSplitString() As String
Dim EmpNo_, EmpName_ As String

mMsg = ""
TOADD = ""
TOCC = ""


mMsg = "<table border = 1><tr><td colspan=2>CITF Details (From Store-Man)</td></tr>"
mMsg = mMsg & "<tr><td>CITF No</td><td>" & TxtCxTFNo.Text & "</td> </tr>"
mMsg = mMsg & "<tr><td>CITF Date</td><td>" & Format(TxtCxTFDate.Value, "dd-MMM-yyyy") & "</td></tr>"
mMsg = mMsg & "<tr><td>Location</td><td>" & GetLocationFromCMP(RsTxn_CxTF_Details!CMPID) & "</td> </tr>"
mMsg = mMsg & "<tr><td>State</td><td>" & GetStateFromLocation(GetLocationIDFromCMP(Val(mCMPID))) & "</td></tr>"
mMsg = mMsg & "<tr><td>CMPID & Address</td><td>" & mCMPID & "-" & GetCMPAddress(Val(mCMPID), "A") & "</td></tr>"
mMsg = mMsg & "<tr><td>Commodity</td><td>" & TxtCommodityID.Text & "</td></tr>"
mMsg = mMsg & "<tr><td>Transaction Type</td><td>" & CmbExchangeTypeID.Text & "</td></tr>"
mMsg = mMsg & "<tr><td>Name Of Agent </td><td>" & TxtAgent.Text & "</td></tr>"
If mExchangeTypeID <> 2 Then
   mMsg = mMsg & "<tr><td>Total No of Bags</td><td>" & MSHFlexGrid2.TextMatrix(0, 15) & "</td></tr>"
End If

mMsg = mMsg & "<tr><td>Total Electronic Weight in MT </td><td>" & MSHFlexGrid2.TextMatrix(0, 9) & "</td></tr>"

mMsg = mMsg & "<tr><td>Hologram No</td><td>" & IIf(IsNull(RsTxn_CxTF_Details!HologramNo), "", RsTxn_CxTF_Details!HologramNo) & "</td></tr></table>"

tmp = " Select EMP.EmployeeID,EMP.EmployeeNo,EMP.EmployeeName,EMP.EmployeeEmailID,EMP1.EmployeeID as 'AMEmpID',EMP1.EmployeeNo as 'AMEmpNo',EMP1.EmployeeName as 'AMName',EMP1.EmployeeEmailID as 'AMEmailID'" & _
      " From Mst_EmpCMPMapping EGM" & _
      " Inner Join Mst_Employee EMP on EGM.EmployeeID = EMP.EmployeeID" & _
      " Inner Join Mst_Employee EMP1 on EMP.DirectReportingTo = EMP1.EmployeeID" & _
      " Where EMP.Flag = 'Y' And EGM.CMPID = " & mCMPID
'      " Group By EMP.EmployeeID,EMP.EmployeeNo,EMP.EmployeeName,EMP.EmployeeEmailID"

Call Tmp3Table
If TmpRs3.RecordCount > 0 Then
   For ab = 1 To TmpRs3.RecordCount
       
        If TmpRs3!EmployeeEmailID <> "" Then
           If InStr(1, TmpRs3!EmployeeEmailID, "@") Then
              If TOADD <> "" Then
                 If InStr(1, LCase(TOADD), LCase(TmpRs3!EmployeeEmailID)) = 0 Then
                    TOADD = TOADD & "," & TmpRs3!EmployeeEmailID
                 End If
              Else
                 TOADD = TmpRs3!EmployeeEmailID
              End If
           End If
        End If
       
        If TmpRs3!AMEmailID <> "" Then
           If InStr(1, TmpRs3!AMEmailID, "@") Then
              If TOCC <> "" Then
                 If InStr(1, LCase(TOCC), LCase(TmpRs3!AMEmailID)) = 0 Then
                    TOCC = TOCC & "," & TmpRs3!AMEmailID
                 End If
              Else
                 TOCC = TmpRs3!AMEmailID
              End If
           End If
        End If
       
'       TOADD = TmpRs3!EmployeeEmailID
'       TOCC = IIf(IsNull(TmpRs3!AMEmailID), "", TmpRs3!AMEmailID) '""
       TmpRs3.MoveNext
   Next
End If

If TOADD = "" Then
   TOADD = "manu.ag@ncml.com"
End If
 
 '"cmg@ncml.com" & "," & Gen_UserEmailID  'Gen_UserEmailID
If TOCC <> "" Then
   TOCC = TOCC & "," & Gen_UserEmailID
Else
   TOCC = Gen_UserEmailID
End If
mMsgSubject = "NCMSL " & TxtCMPID.Text & " CITF Details"

Call SendMail(TOADD, TOCC, mMsgSubject, mMsg, "N")

End Sub

Private Function SendMail(ToADD_ As Variant, TOCC_ As Variant, MsgSubject_ As Variant, Message1_ As Variant, mMailTypeFlg As Variant)
Dim htmlStringhead As String
Dim htmlStringtail As String
Set oSMTP = New OSSMTP.SMTPSession

servertxt = "mail.ncml.com"

htmlStringhead = "<html><body><font size= 3 face= Zurich BT color = Blue>" '<table border = 1>"
htmlStringtail = "</Font></body></html>" '</table>
   
'With oSMTP
'  'connection
'  .Server = servertxt
'  'message
'  '.MailFrom = "info@ncml.com" 'Gen_UserEmailID 'TOCC_ '"mayur.d@ncml.com"
'  If LCase(mMailTypeFlg) = "n" Then
'     .MailFrom = "info@ncml.com" 'Gen_UserEmailID 'TOCC_ '"mayur.d@ncml.com"
'  Else
'     .MailFrom = Gen_UserEmailID
'  End If
'  .SendTo = ToADD_
''  TOCC_ = TOCC_ '& "," & mVendorEmailID
'  If ToADD_ <> TOCC_ Then
'     .CC = TOCC_
'  End If
''  .BCC = "mayur.d@ncml.com"
'  If LCase(mMailTypeFlg) = "n" Then
'     .Attachments.Add (mPDFFileName)
'  End If
'  .MessageSubject = MsgSubject_ & " (from Store-Man)" '"Confidential Mail"
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
If LCase(mMailTypeFlg) = "n" Then
   objSMTP.FromAddr = "info@ncml.com"
Else
   objSMTP.FromAddr = Gen_UserEmailID
End If
If IsNull(Trim(ToADD_)) = True Or Trim(ToADD_) = "" Then
   ToADD_ = "manu.ag@ncml.com"
End If
objSMTP.ToAddr = ToADD_
If ToADD_ <> TOCC_ Then
   objSMTP.CCAddr = TOCC_
End If
If LCase(mMailTypeFlg) = "n" Then
   objSMTP.AddAttachment mPDFFileName
End If
objSMTP.Subject = MsgSubject_ & " (from Store-Man)" 'MsgSubject_ '"Message in HTML format"
objSMTP.BodyFormat = 1
objSMTP.BodyText = htmlStringhead & Message1_ & htmlStringtail 'htmlStringhead & MsgSubject_ & " :- " & Message1_ & htmlStringtail '"<html>This word is <b>bold</b></html>"
If LCase(Gen_EmailConfigFlag) = "y" Then
   objSMTP.Send
End If

End Function

'Private Sub FillBOE()
'
'Dim I As Variant
'Dim mdate As Date
'
'
'tmp = "Select CommodityGroupID From Mst_Commodity Where CommodityID=" & mCommodityID
'Call TmpTable
'If TmpRs.RecordCount > 0 Then
'   mCommodityGroupId = TmpRs!CommodityGroupID
'End If
'mLocationID = GetLocationIDFromCMP(Val(mCMPID))
'
'Call Grid_Head_Bill
'mChkSameMonth = False
'If LCase(CmbExchangeTypeID) = "non ncdex" Then
'
''   tmp = "Select TC.CxTFDate,GSL.CDTFDate,datediff(day,GSL.CDTFDate,TC.CxTFDate)+ 1 NoOfDay,GSL.ClientIDRow,MC.ClientName,TCG.GSLID,MG.GodownNo+'~'+"
''   tmp = tmp & " GSL.StackNo +'~'+ GSL.LotNo As GSL,GSL.SpillageFlag,TCG.NoOfBags,TCG.DematWeight,(TCG.DematWeight/TCG.NoOfBags)*1000 as 'BagSize',GSL.NContractID"
''   tmp = tmp & " From  Txn_CxTF_GSL TCG"
''   tmp = tmp & " Inner Join Txn_CxTF_Details TC on TCG.SM_Prefix=TC.SM_Prefix And TCG.CxTFNo=TC.CxTFNo And TCG.TxnType=TC.TxnType"
''   tmp = tmp & " left Join Mst_GSL GSL on TCG.GSLID=GSL.GSLID And TCG.SM_Prefix=GSL.SM_Prefix"
''   tmp = tmp & " Inner Join Mst_Godown MG on GSL.GodownID=MG.GodownID"
''   tmp = tmp & " Inner Join Mst_Client MC on GSL.ClientIDRow=MC.ClientIDRow"
''   tmp = tmp & " Where TC.ExchangeTypeID=1 And TC.CxTFNo=" & Val(TxtCxTFNo) & " And TC.TxnType='W' And TC.SM_Prefix='" & TxtPreFix.Text & "'"
''
''   Call TmpTable
''
''   If TmpRs.RecordCount > 0 Then
''      For I = 1 To TmpRs.RecordCount
''          For c = 0 To TmpRs.Fields.Count - 1
''              MSHFlexGrid6.TextMatrix(I, c) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
''          Next
''          MSHFlexGrid6.TextMatrix(I, 0) = Format(MSHFlexGrid6.TextMatrix(I, 0), "dd-MMM-yyyy")
''          MSHFlexGrid6.TextMatrix(I, 1) = Format(MSHFlexGrid6.TextMatrix(I, 1), "dd-MMM-yyyy")
''          If MSHFlexGrid6.TextMatrix(I, 7) = "1" Then
''             MSHFlexGrid6.TextMatrix(I, 7) = "Spillage"
''          Else
''             MSHFlexGrid6.TextMatrix(I, 7) = "Normal"
''             'Call FillBillingDetails(i)
''
''             Call FillBillingDetails(I, TmpRs!CDTFDate, TmpRs!CxtfDate)
''
''          End If
''          If MSHFlexGrid6.TextMatrix(I, 13) = "Bag" Then
''             MSHFlexGrid6.TextMatrix(I, 16) = Format(Val(MSHFlexGrid6.TextMatrix(I, 14)) * Val(MSHFlexGrid6.TextMatrix(I, 15)) * Val(MSHFlexGrid6.TextMatrix(I, 8)), "#####0.00")
''          Else
''             MSHFlexGrid6.TextMatrix(I, 16) = Format(Val(MSHFlexGrid6.TextMatrix(I, 14)) * Val(MSHFlexGrid6.TextMatrix(I, 15)) * Val(MSHFlexGrid6.TextMatrix(I, 9)), "#####0.00")
''          End If
''          TmpRs.MoveNext
''          MSHFlexGrid6.Rows = MSHFlexGrid6.Rows + 1
''      Next
''   End If
'
'
''  Commented on 12 Dec 2011 549 PM To Stop Spillage in BOE Calculation
'
''   tmp = "Select TC.CxTFDate,GSL.CDTFDate,datediff(day,GSL.CDTFDate,TC.CxTFDate)+ 1 NoOfDay,GSL.ClientIDRow,MC.ClientName,TCG.GSLID,MG.GodownNo+'~'+"
''   tmp = tmp & " GSL.StackNo +'~'+ GSL.LotNo As GSL,GSL.SpillageFlag,TCG.NoOfBags,TCG.DematWeight,(TCG.DematWeight/TCG.NoOfBags)*1000 as 'BagSize',GSL.NContractID"
''   tmp = tmp & " From  Txn_CxTF_GSL TCG"
''   tmp = tmp & " Inner Join Txn_CxTF_Details TC on TCG.SM_Prefix=TC.SM_Prefix And TCG.CxTFNo=TC.CxTFNo And TCG.TxnType=TC.TxnType"
''   tmp = tmp & " left Join Mst_GSL GSL on TCG.GSLID=GSL.GSLID And TCG.SM_Prefix=GSL.SM_Prefix"
''   tmp = tmp & " Inner Join Mst_Godown MG on GSL.GodownID=MG.GodownID"
''   tmp = tmp & " Inner Join Mst_Client MC on GSL.ClientIDRow=MC.ClientIDRow"
''   tmp = tmp & " Where TC.ExchangeTypeID=1 And TC.CxTFNo=" & Val(TxtCxTFNo) & " And TC.TxnType='W' And TC.SM_Prefix='" & TxtPreFix.Text & "'"
'
''' Commented on 5 May 2012 to Calculate NoofDays from 1 st of CITF Dates Month
'
'''   tmp = "Select TC.CxTFDate,GSL.CDTFDate,datediff(day,GSL.CDTFDate,TC.CxTFDate)+ 1 NoOfDay,GSL.ClientIDRow,MC.ClientName,TCG.GSLID,MG.GodownNo+'~'+"
'''   tmp = tmp & " GSL.StackNo +'~'+ GSL.LotNo As GSL,GSL.SpillageFlag,TCG.NoOfBags,TCG.DematWeight,(TCG.DematWeight/TCG.NoOfBags)*1000 as 'BagSize',GSL.NContractID"
'''   tmp = tmp & " From  Txn_CxTF_GSL TCG"
'''   tmp = tmp & " Inner Join Txn_CxTF_Details TC on TCG.SM_Prefix=TC.SM_Prefix And TCG.CxTFNo=TC.CxTFNo And TCG.TxnType=TC.TxnType"
'''   tmp = tmp & " left Join Mst_GSL GSL on TCG.GSLID=GSL.GSLID And TCG.SM_Prefix=GSL.SM_Prefix"
'''   tmp = tmp & " Inner Join Mst_Godown MG on GSL.GodownID=MG.GodownID"
'''   tmp = tmp & " Inner Join Mst_Client MC on GSL.ClientIDRow=MC.ClientIDRow"
'''   tmp = tmp & " Where TC.ExchangeTypeID=1 And TC.CxTFNo=" & Val(TxtCxTFNo) & " And TC.TxnType='W' And TC.SM_Prefix='" & TxtPreFix.Text & "' And GSL.SpillageFlag = 0 "
'
'   tmp = "Select TC.CxTFDate,GSL.CDTFDate,day(TC.CxTFDate) NoOfDay,GSL.ClientIDRow,MC.ClientName,TCG.GSLID,MG.GodownNo+'~'+"
'   tmp = tmp & " GSL.StackNo +'~'+ GSL.LotNo As GSL,GSL.SpillageFlag,TCG.NoOfBags,TCG.DematWeight,(TCG.DematWeight/TCG.NoOfBags)*1000 as 'BagSize',GSL.NContractID,GSL.MadeUpFlag"
'   tmp = tmp & " From  Txn_CxTF_GSL TCG"
'   tmp = tmp & " Inner Join Txn_CxTF_Details TC on TCG.SM_Prefix=TC.SM_Prefix And TCG.CxTFNo=TC.CxTFNo And TCG.TxnType=TC.TxnType"
'   tmp = tmp & " left Join Mst_GSL GSL on TCG.GSLID=GSL.GSLID And TCG.SM_Prefix=GSL.SM_Prefix"
'   tmp = tmp & " Inner Join Mst_Godown MG on GSL.GodownID=MG.GodownID"
'   tmp = tmp & " Inner Join Mst_Client MC on GSL.ClientIDRow=MC.ClientIDRow"
'   tmp = tmp & " Where TC.ExchangeTypeID=1 And TC.CxTFNo=" & Val(TxtCxTFNo) & " And TC.TxnType='W' And TC.SM_Prefix='" & TxtPreFix.Text & "' And GSL.SpillageFlag = 0 And GSL.MadeUpFlag = 0 "
'
'   Call TmpTable
'
'   If TmpRs.RecordCount > 0 Then
'      For I = 1 To TmpRs.RecordCount
'          For c = 0 To TmpRs.Fields.Count - 2
'              MSHFlexGrid6.TextMatrix(I, c) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
'          Next
'          MSHFlexGrid6.TextMatrix(I, 0) = Format(MSHFlexGrid6.TextMatrix(I, 0), "dd-MMM-yyyy")
'          MSHFlexGrid6.TextMatrix(I, 1) = Format(MSHFlexGrid6.TextMatrix(I, 1), "dd-MMM-yyyy")
'          MSHFlexGrid6.TextMatrix(I, 10) = Round(Val(MSHFlexGrid6.TextMatrix(I, 10)), 0)
'          If MSHFlexGrid6.TextMatrix(I, 7) = "1" Then
'             MSHFlexGrid6.TextMatrix(I, 7) = "Spillage"
'          ElseIf TmpRs!MadeUpFlag = "1" Then
'             MSHFlexGrid6.TextMatrix(I, 7) = "Madeup"
'          Else
'             MSHFlexGrid6.TextMatrix(I, 7) = "Normal"
'             'Call FillBillingDetails(i)
'
'             ''Call FillBillingDetails(I, TmpRs!CDTFDate, TmpRs!CxTFDate)
'             mMonth = Month(TmpRs!CxTFDate)
'             mYear = Year(TmpRs!CxTFDate)
'             mdate = CDate(mMonth & "-" & "01" & "-" & mYear)
'             'Call FillBillingDetails(I, mdate, TmpRs!CxTFDate)
'
'             If Month(TmpRs!CDTFDate) & "~" & Year(TmpRs!CDTFDate) = Month(TmpRs!CxTFDate) & "~" & Year(TmpRs!CxTFDate) Then
'             '   MSHFlexGrid6.TextMatrix(I, 16) = 0
'             '   MSHFlexGrid6.TextMatrix(I, 20) = 0
'             '   MSHFlexGrid6.TextMatrix(I, 21) = 0
'                mChkSameMonth = True
'                Call FillBillingDetails(I, TmpRs!CDTFDate, TmpRs!CxTFDate)
'             Else
'                Call FillBillingDetails(I, mdate, TmpRs!CxTFDate)
'             End If
'
'          End If
'          If MSHFlexGrid6.TextMatrix(I, 13) = "Bag" Then
'             MSHFlexGrid6.TextMatrix(I, 16) = Format(Val(MSHFlexGrid6.TextMatrix(I, 14)) * Val(MSHFlexGrid6.TextMatrix(I, 15)) * Val(MSHFlexGrid6.TextMatrix(I, 8)), "#####0.00")
'             MSHFlexGrid6.TextMatrix(I, 20) = Format(Val(MSHFlexGrid6.TextMatrix(I, 18)) * Val(MSHFlexGrid6.TextMatrix(I, 19)) * Val(MSHFlexGrid6.TextMatrix(I, 8)), "#####0.00")
'             MSHFlexGrid6.TextMatrix(I, 21) = Val(MSHFlexGrid6.TextMatrix(I, 16)) + Val(MSHFlexGrid6.TextMatrix(I, 20))
'          Else
'             MSHFlexGrid6.TextMatrix(I, 16) = Format(Val(MSHFlexGrid6.TextMatrix(I, 14)) * Val(MSHFlexGrid6.TextMatrix(I, 15)) * Val(MSHFlexGrid6.TextMatrix(I, 9)), "#####0.00")
'             MSHFlexGrid6.TextMatrix(I, 20) = Format(Val(MSHFlexGrid6.TextMatrix(I, 18)) * Val(MSHFlexGrid6.TextMatrix(I, 19)) * Val(MSHFlexGrid6.TextMatrix(I, 9)), "#####0.00")
'             MSHFlexGrid6.TextMatrix(I, 21) = Val(MSHFlexGrid6.TextMatrix(I, 16)) + Val(MSHFlexGrid6.TextMatrix(I, 20))
'          End If
'
''          If Month(TmpRs!CDTFDate) & "~" & Year(TmpRs!CDTFDate) = Month(TmpRs!CxTFDate) & "~" & Year(TmpRs!CxTFDate) Then
''             MSHFlexGrid6.TextMatrix(I, 16) = 0
''             MSHFlexGrid6.TextMatrix(I, 20) = 0
''             MSHFlexGrid6.TextMatrix(I, 21) = 0
''          End If
'
'          TmpRs.MoveNext
'          MSHFlexGrid6.Rows = MSHFlexGrid6.Rows + 1
'          mChkSameMonth = False
'      Next
'   End If
'
'ElseIf LCase(CmbExchangeTypeID) = "ncdex" Then
'
''   tmp = " Select TCD.CxTFDate,MG.CDTFDate,(datediff(day,MG.CDTFDate,TCD.CxTFDate) + 1) NoOfDay,MG.ClientIDRow,MClient.ClientName,"
''   tmp = tmp & " TCG.GSLID, MGod.GodownNo +'~'+ MG.StackNo  +'~'+ MG.LotNo as GSL,MG.SpillageFlag,TCG.NoOfBags,TCG.DematWeight,0 as 'BagSize',MG.NContractID"
''   tmp = tmp & " ,'Daily' Cycle,'Qty' Unit,(datediff(day,MG.CDTFDate,TCD.CxTFDate) + 1) NoOfCycle,MNRC.Rate,(datediff(day,MG.CDTFDate,TCD.CxTFDate) + 1) * TCG.DematWeight * Rate as 'Amount'"
''   tmp = tmp & " From Txn_CxTF_Details TCD"
''   tmp = tmp & " Inner Join Txn_CxTF_GSL TCG On TCD.CxTFNo = TCG.CxTFNo And TCD.SM_Prefix = TCG.SM_Prefix And TCD.TxnType = TCG.TxnType"
''   tmp = tmp & " Inner Join Mst_GSL MG On TCG.GSLID = MG.GSLID And TCG.SM_Prefix = MG.SM_Prefix"
''   tmp = tmp & " Inner Join Mst_Godown MGod On MG.GodownID = MGod.GodownID"
''   tmp = tmp & " Inner Join Mst_CMP MC On TCD.CMPID = MC.CMPID"
''   tmp = tmp & " Inner Join Mst_Location ML On MC.LocationID = ML.LocationID"
''   tmp = tmp & " Inner Join Mst_Commodity MComm On TCD.CommodityID = MComm.CommodityID"
''   tmp = tmp & " Left Join Mst_Client MClient On MG.ClientIDRow = MClient.ClientIDRow"
''   tmp = tmp & " Left Join Mst_NCDEXStorageRateCard MNR On ML.LocationID = MNR.LocationID And MComm.CommodityGroupID = MNR.CommodityGroupID"
''   tmp = tmp & " Left Join Mst_NCDEXStorageRateCardDetail MNRC On MNR.NSRCID = MNRC.NSRCID And MG.GodownID = MNRC.GodownID"
''   tmp = tmp & " Where TCD.TxnType = 'W' and TCD.ExchangeTypeID = 2 and TCD.BookedSpaceFlag = 1 And TCD.CxTFNo=" & Val(TxtCxTFNo) & " And TCD.SM_Prefix='" & TxtPreFix.Text & "'"
'
'
''  Commented on 12 Dec 2011 549 PM To Stop Spillage in BOE Calculation
'
''   tmp = " Select TCD.CxTFDate,MG.CDTFDate,((datediff(day,MG.CDTFDate,TCD.CxTFDate) + 1) - 7) NoOfDay,MG.ClientIDRow,MClient.ClientName,"
''   tmp = tmp & " TCG.GSLID, MGod.GodownNo +'~'+ MG.StackNo  +'~'+ MG.LotNo as GSL,MG.SpillageFlag,TCG.NoOfBags,TCG.DematWeight,0 as 'BagSize',MG.NContractID"
''   tmp = tmp & " ,'Daily' Cycle,'Qty' Unit,((datediff(day,MG.CDTFDate,TCD.CxTFDate) + 1) - 7) NoOfCycle,MNRC.Rate,((datediff(day,MG.CDTFDate,TCD.CxTFDate) + 1) - 7) * TCG.DematWeight * Rate as 'Amount'"
''   tmp = tmp & " From Txn_CxTF_Details TCD"
''   tmp = tmp & " Inner Join Txn_CxTF_GSL TCG On TCD.CxTFNo = TCG.CxTFNo And TCD.SM_Prefix = TCG.SM_Prefix And TCD.TxnType = TCG.TxnType"
''   tmp = tmp & " Inner Join Mst_GSL MG On TCG.GSLID = MG.GSLID And TCG.SM_Prefix = MG.SM_Prefix"
''   tmp = tmp & " Inner Join Mst_Godown MGod On MG.GodownID = MGod.GodownID"
''   tmp = tmp & " Inner Join Mst_CMP MC On TCD.CMPID = MC.CMPID"
''   tmp = tmp & " Inner Join Mst_Location ML On MC.LocationID = ML.LocationID"
''   tmp = tmp & " Inner Join Mst_Commodity MComm On TCD.CommodityID = MComm.CommodityID"
''   tmp = tmp & " Left Join Mst_Client MClient On MG.ClientIDRow = MClient.ClientIDRow"
''   tmp = tmp & " Left Join Mst_NCDEXStorageRateCard MNR On ML.LocationID = MNR.LocationID And MComm.CommodityGroupID = MNR.CommodityGroupID"
''   tmp = tmp & " Left Join Mst_NCDEXStorageRateCardDetail MNRC On MNR.NSRCID = MNRC.NSRCID And MG.GodownID = MNRC.GodownID"
''   tmp = tmp & " Where TCD.TxnType = 'W' and TCD.ExchangeTypeID = 2 and TCD.BookedSpaceFlag = 1 And TCD.CxTFNo=" & Val(TxtCxTFNo) & " And TCD.SM_Prefix='" & TxtPreFix.Text & "'"
'
''' Commented on 5 May 2012 to Calculate NoofDays from 1 st of CITF Dates Month
'
'''   tmp = " Select TCD.CxTFDate,MG.CDTFDate,((datediff(day,MG.CDTFDate,TCD.CxTFDate) + 1) - 7) NoOfDay,MG.ClientIDRow,MClient.ClientName,"
'''   tmp = tmp & " TCG.GSLID, MGod.GodownNo +'~'+ MG.StackNo  +'~'+ MG.LotNo as GSL,MG.SpillageFlag,TCG.NoOfBags,TCG.DematWeight,0 as 'BagSize',MG.NContractID"
'''   tmp = tmp & " ,'Daily' Cycle,'Qty' Unit,((datediff(day,MG.CDTFDate,TCD.CxTFDate) + 1) - 7) NoOfCycle,MNRC.Rate,((datediff(day,MG.CDTFDate,TCD.CxTFDate) + 1) - 7) * TCG.DematWeight * Rate as 'Amount'"
'''   tmp = tmp & " From Txn_CxTF_Details TCD"
'''   tmp = tmp & " Inner Join Txn_CxTF_GSL TCG On TCD.CxTFNo = TCG.CxTFNo And TCD.SM_Prefix = TCG.SM_Prefix And TCD.TxnType = TCG.TxnType"
'''   tmp = tmp & " Inner Join Mst_GSL MG On TCG.GSLID = MG.GSLID And TCG.SM_Prefix = MG.SM_Prefix"
'''   tmp = tmp & " Inner Join Mst_Godown MGod On MG.GodownID = MGod.GodownID"
'''   tmp = tmp & " Inner Join Mst_CMP MC On TCD.CMPID = MC.CMPID"
'''   tmp = tmp & " Inner Join Mst_Location ML On MC.LocationID = ML.LocationID"
'''   tmp = tmp & " Inner Join Mst_Commodity MComm On TCD.CommodityID = MComm.CommodityID"
'''   tmp = tmp & " Left Join Mst_Client MClient On MG.ClientIDRow = MClient.ClientIDRow"
'''   tmp = tmp & " Left Join Mst_NCDEXStorageRateCard MNR On ML.LocationID = MNR.LocationID And MComm.CommodityGroupID = MNR.CommodityGroupID"
'''   tmp = tmp & " Left Join Mst_NCDEXStorageRateCardDetail MNRC On MNR.NSRCID = MNRC.NSRCID And MG.GodownID = MNRC.GodownID"
'''   tmp = tmp & " Where TCD.TxnType = 'W' and TCD.ExchangeTypeID = 2 and TCD.BookedSpaceFlag = 1 And TCD.CxTFNo=" & Val(TxtCxTFNo) & " And TCD.SM_Prefix='" & TxtPreFix.Text & "' And MG.SpillageFlag = 0 "
'
'
'   tmp = " Select TCD.CxTFDate,MG.CDTFDate,((datediff(day,MG.CDTFDate,TCD.CxTFDate) + 1) - 7) NoOfDay,MG.ClientIDRow,MClient.ClientName,"
'   tmp = tmp & " TCG.GSLID, MGod.GodownNo +'~'+ MG.StackNo  +'~'+ MG.LotNo as GSL,MG.SpillageFlag,TCG.NoOfBags,TCG.DematWeight,0 as 'BagSize',MG.NContractID"
'   tmp = tmp & " ,'Daily' Cycle,'Qty' Unit,((datediff(day,MG.CDTFDate,TCD.CxTFDate) + 1) - 7) NoOfCycle,MNRC.Rate,((datediff(day,MG.CDTFDate,TCD.CxTFDate) + 1) - 7) * TCG.DematWeight * Rate as 'Amount',MG.MadeupFlag"
'   tmp = tmp & " From Txn_CxTF_Details TCD"
'   tmp = tmp & " Inner Join Txn_CxTF_GSL TCG On TCD.CxTFNo = TCG.CxTFNo And TCD.SM_Prefix = TCG.SM_Prefix And TCD.TxnType = TCG.TxnType"
'   tmp = tmp & " Inner Join Mst_GSL MG On TCG.GSLID = MG.GSLID And TCG.SM_Prefix = MG.SM_Prefix"
'   tmp = tmp & " Inner Join Mst_Godown MGod On MG.GodownID = MGod.GodownID"
'   tmp = tmp & " Inner Join Mst_CMP MC On TCD.CMPID = MC.CMPID"
'   tmp = tmp & " Inner Join Mst_Location ML On MC.LocationID = ML.LocationID"
'   tmp = tmp & " Inner Join Mst_Commodity MComm On TCD.CommodityID = MComm.CommodityID"
'   tmp = tmp & " Left Join Mst_Client MClient On MG.ClientIDRow = MClient.ClientIDRow"
'   tmp = tmp & " Left Join Mst_NCDEXStorageRateCard MNR On ML.LocationID = MNR.LocationID And MComm.CommodityGroupID = MNR.CommodityGroupID"
'   tmp = tmp & " Left Join Mst_NCDEXStorageRateCardDetail MNRC On MNR.NSRCID = MNRC.NSRCID And MG.GodownID = MNRC.GodownID"
'   tmp = tmp & " Where TCD.TxnType = 'W' and TCD.ExchangeTypeID = 2 and TCD.BookedSpaceFlag = 1 And TCD.CxTFNo=" & Val(TxtCxTFNo) & " And TCD.SM_Prefix='" & TxtPreFix.Text & "' And MG.SpillageFlag = 0 "
'
'   Call TmpTable
'
'   If TmpRs.RecordCount > 0 Then
'      For I = 1 To TmpRs.RecordCount
'          For c = 0 To TmpRs.Fields.Count - 2
'              MSHFlexGrid6.TextMatrix(I, c) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
'          Next
'          MSHFlexGrid6.TextMatrix(I, 0) = Format(MSHFlexGrid6.TextMatrix(I, 0), "dd-MMM-yyyy")
'          MSHFlexGrid6.TextMatrix(I, 1) = Format(MSHFlexGrid6.TextMatrix(I, 1), "dd-MMM-yyyy")
'          If MSHFlexGrid6.TextMatrix(I, 7) = "1" Then
'             MSHFlexGrid6.TextMatrix(I, 7) = "Spillage"
'          ElseIf TmpRs!MadeUpFlag = "1" Then
'             MSHFlexGrid6.TextMatrix(I, 7) = "Madeup"
'          Else
'             MSHFlexGrid6.TextMatrix(I, 7) = "Normal"
'          End If
'
'          If Val(MSHFlexGrid6.TextMatrix(I, 2)) < 0 Then
'             MSHFlexGrid6.TextMatrix(I, 2) = 0
'             MSHFlexGrid6.TextMatrix(I, 14) = 0
'             MSHFlexGrid6.TextMatrix(I, 16) = 0
'          End If
'
'          TmpRs.MoveNext
'          MSHFlexGrid6.Rows = MSHFlexGrid6.Rows + 1
'      Next
'   End If
'End If
'
'End Sub

Private Sub FillBOE()

Dim I As Variant
Dim mDate As Date

tmp = "Select CommodityGroupID From Mst_Commodity Where CommodityID=" & mCommodityID
Call TmpTable
If TmpRs.RecordCount > 0 Then
   mCommodityGroupId = TmpRs!CommodityGroupID
End If
mLocationID = GetLocationIDFromCMP(Val(mCMPID))

Call Grid_Head_Bill
mChkSameMonth = False

If LCase(CmbExchangeTypeID) = "non ncdex" Then

 
   tmp = "Select TC.CxTFDate,GSL.CDTFDate,day(TC.CxTFDate) NoOfDay,GSL.ClientID,MC.ClientName,TCG.Mst_GSL_ID,MG.GodownNo+'~'+"
   tmp = tmp & " GSL.StackNo +'~'+ GSL.LotNo As GSL,GSL.SpillageFlag,TCG.NoOfBags,TCG.DematWeight,(TCG.DematWeight/TCG.NoOfBags)*1000 as 'BagSize',GSL.NContractID, MBC.BillingCycle 'WBillingCycle', TCG.BillingUnit,MBC.NoofDays, TCG.Rate, '' 'StorageAmount' , MBC1.BillingCycle 'DepBillingCycle' ,MBC1.NoOfDays 'DepoNoOfDays' ,TDCD.DepositFinalPrice,'' DepoStorageChargesAmount, '' TotalStorageChargesAmount,MBC.BillingCycleID,MBC1.BillingCycleID 'DepoBillingCycleId' "
   tmp = tmp & " From  Txn_CxTF_GSL TCG"
   tmp = tmp & " Inner Join Txn_CxTF_Details TC on TCG.SM_Prefix=TC.SM_Prefix And TCG.CxTF_DetailsID=TC.CxTF_DetailsID And TCG.TxnType=TC.TxnType"
   tmp = tmp & " left Join Mst_GSL GSL on TCG.Mst_GSL_ID=GSL.Mst_GSL_ID And TCG.SM_Prefix=GSL.SM_Prefix"
   tmp = tmp & " Inner Join Mst_Godown MG on GSL.GodownID=MG.GodownID"
   tmp = tmp & " Inner Join Mst_Client MC on GSL.ClientID=MC.ClientID"
   tmp = tmp & " Left Join Txn_DepositContractGodownDetail TDCD On TCG.DepConDID = TDCD.DepConDID "
   tmp = tmp & " Left Join Txn_DepositContract TDC On TDCD.DepConID = TDC.DepConID  And TDC.ApprovalFlag in ('S','A','C')  "
   tmp = tmp & " Left Join Mst_BillingCycle MBC On TCG.BillingCycleID = MBC.BillingCycleID "
   tmp = tmp & " Left Join Mst_BillingCycle MBC1 On TDC.DepBillingCycleID = MBC1.BillingCycleID "
   tmp = tmp & " Where TCG.NoOfBags > 0 And TC.ExchangeTypeID=1 And TC.CxTF_DetailsID=" & Val(TxtCxTFNo) & " And TC.TxnType='W' And TC.SM_Prefix='" & TxtPreFix.Text & "' And GSL.SpillageFlag = 0 And GSL.MadeUpFlag = 0 "
    
   Call TmpTable
    
   If TmpRs.RecordCount > 0 Then
      For I = 1 To TmpRs.RecordCount
          For C = 0 To TmpRs.Fields.Count - 3
              MSHFlexGrid6.TextMatrix(I, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
          Next
          MSHFlexGrid6.TextMatrix(I, 0) = Format(MSHFlexGrid6.TextMatrix(I, 0), "dd-MMM-yyyy")
          MSHFlexGrid6.TextMatrix(I, 1) = Format(MSHFlexGrid6.TextMatrix(I, 1), "dd-MMM-yyyy")
          MSHFlexGrid6.TextMatrix(I, 10) = Round(Val(MSHFlexGrid6.TextMatrix(I, 10)), 0)
          If MSHFlexGrid6.TextMatrix(I, 7) = "1" Then
             MSHFlexGrid6.TextMatrix(I, 7) = "Spillage"
          Else
             MSHFlexGrid6.TextMatrix(I, 7) = "Normal"
             mMonth = Month(TmpRs!CxTFDate)
             mYear = Year(TmpRs!CxTFDate)
             mDate = CDate(mMonth & "-" & "01" & "-" & mYear)
          
             If Month(TmpRs!CDTFDate) & "~" & Year(TmpRs!CDTFDate) = Month(TmpRs!CxTFDate) & "~" & Year(TmpRs!CxTFDate) Then
                mChkSameMonth = True
                Call FillBillingDetails(I, TmpRs!CDTFDate, TmpRs!CxTFDate)
             Else
                Call FillBillingDetails(I, mDate, TmpRs!CxTFDate)
             End If
          
          End If
          If MSHFlexGrid6.TextMatrix(I, 13) = "Bag" Then
             MSHFlexGrid6.TextMatrix(I, 16) = Format(Val(MSHFlexGrid6.TextMatrix(I, 14)) * Val(MSHFlexGrid6.TextMatrix(I, 15)) * Val(MSHFlexGrid6.TextMatrix(I, 8)), "#####0.00")
             MSHFlexGrid6.TextMatrix(I, 20) = Format(Val(MSHFlexGrid6.TextMatrix(I, 18)) * Val(MSHFlexGrid6.TextMatrix(I, 19)) * Val(MSHFlexGrid6.TextMatrix(I, 8)), "#####0.00")
             MSHFlexGrid6.TextMatrix(I, 21) = Val(MSHFlexGrid6.TextMatrix(I, 16)) '+ Val(MSHFlexGrid6.TextMatrix(i, 20))
          Else
             MSHFlexGrid6.TextMatrix(I, 16) = Format(Val(MSHFlexGrid6.TextMatrix(I, 14)) * Val(MSHFlexGrid6.TextMatrix(I, 15)) * Val(MSHFlexGrid6.TextMatrix(I, 9)), "#####0.00")
             MSHFlexGrid6.TextMatrix(I, 20) = Format(Val(MSHFlexGrid6.TextMatrix(I, 18)) * Val(MSHFlexGrid6.TextMatrix(I, 19)) * Val(MSHFlexGrid6.TextMatrix(I, 9)), "#####0.00")
             MSHFlexGrid6.TextMatrix(I, 21) = Val(MSHFlexGrid6.TextMatrix(I, 16)) '+ Val(MSHFlexGrid6.TextMatrix(i, 20))
          End If
          
          TmpRs.MoveNext
          MSHFlexGrid6.Rows = MSHFlexGrid6.Rows + 1
          mChkSameMonth = False
      Next
   End If

ElseIf LCase(CmbExchangeTypeID) = "ncdex" Then
   
   tmp = " Select TCD.CxTFDate,MG.CDTFDate,((datediff(day,MG.CDTFDate,TCD.CxTFDate) + 1) - 7) NoOfDay,MG.ClientID,MClient.ClientName,"
   tmp = tmp & " TCG.Mst_GSL_ID, MGod.GodownNo +'~'+ MG.StackNo  +'~'+ MG.LotNo as GSL,MG.SpillageFlag,TCG.NoOfBags,TCG.DematWeight,0 as 'BagSize',MG.NContractID"
   tmp = tmp & " ,'Daily' Cycle,'Qty' Unit,((datediff(day,MG.CDTFDate,TCD.CxTFDate) + 1) - 7) NoOfCycle,MNRC.Rate,((datediff(day,MG.CDTFDate,TCD.CxTFDate) + 1) - 7) * TCG.DematWeight * MNRC.Rate as 'Amount',ML.Collect7Days,MG.LCDTF_Date"
   tmp = tmp & " From Txn_CxTF_Details TCD"
   tmp = tmp & " Inner Join Txn_CxTF_GSL TCG On TCD.CxTF_DetailsID = TCG.CxTF_DetailsID And TCD.SM_Prefix = TCG.SM_Prefix And TCD.TxnType = TCG.TxnType"
   tmp = tmp & " Inner Join Mst_GSL MG On TCG.Mst_GSL_ID = MG.Mst_GSL_ID And TCG.SM_Prefix = MG.SM_Prefix"
   tmp = tmp & " Inner Join Mst_Godown MGod On MG.GodownID = MGod.GodownID"
   tmp = tmp & " Inner Join Mst_CMP MC On TCD.CMPID = MC.CMPID"
   tmp = tmp & " Inner Join Mst_Location ML On MC.LocationID = ML.LocationID"
   tmp = tmp & " Inner Join Mst_Commodity MComm On TCD.CommodityID = MComm.CommodityID"
   tmp = tmp & " Left Join Mst_Client MClient On MG.ClientID = MClient.ClientID"
   tmp = tmp & " Left Join Mst_NCDEXStorageRateCard MNR On ML.LocationID = MNR.LocationID And MComm.CommodityGroupID = MNR.CommodityGroupID"
   tmp = tmp & " Left Join Mst_NCDEXStorageRateCardDetail MNRC On MNR.NSRCID = MNRC.NSRCID And MG.GodownID = MNRC.GodownID"
   tmp = tmp & " Where TCD.TxnType = 'W' and TCD.ExchangeTypeID = 2 and TCD.BookedSpaceFlag = 1 And TCD.CxTF_DetailsID=" & Val(TxtCxTFNo) & " And TCD.SM_Prefix='" & TxtPreFix.Text & "' And MG.SpillageFlag = 0 "
      
   Call TmpTable
    
   If TmpRs.RecordCount > 0 Then
      For I = 1 To TmpRs.RecordCount
          For C = 0 To TmpRs.Fields.Count - 3
              MSHFlexGrid6.TextMatrix(I, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
          Next
          
          If TmpRs!Collect7Days = 1 Then
    
             If Check7DaysInvoice(TmpRs!Mst_GSL_ID) = True Then
                mTotDays = DateDiff("d", DateAdd("d", 7, TmpRs!LCDTF_Date), TmpRs!CxTFDate) + 1
             Else
                mBGSLIDs = ""
                Call GetBGSLIDs(TxtPreFix, TmpRs!Mst_GSL_ID)
                If mBGSLIDs <> "" Then
                   If Check7DaysInvoice(mBGSLIDs) = True Then
                      mTotDays = DateDiff("d", DateAdd("d", 7, TmpRs!LCDTF_Date), TmpRs!CxTFDate) + 1
                   Else
                      mTotDays = DateDiff("d", TmpRs!CDTFDate, TmpRs!CxTFDate) + 1
                   End If
                Else
                   mTotDays = DateDiff("d", TmpRs!CDTFDate, TmpRs!CxTFDate) + 1
                End If
             End If
          Else
             mTotDays = DateDiff("d", TmpRs!CDTFDate, TmpRs!CxTFDate) + 1
          End If
          
          MSHFlexGrid6.TextMatrix(I, 2) = mTotDays
          MSHFlexGrid6.TextMatrix(I, 14) = mTotDays
          
          MSHFlexGrid6.TextMatrix(I, 0) = Format(MSHFlexGrid6.TextMatrix(I, 0), "dd-MMM-yyyy")
          MSHFlexGrid6.TextMatrix(I, 1) = Format(MSHFlexGrid6.TextMatrix(I, 1), "dd-MMM-yyyy")
          If MSHFlexGrid6.TextMatrix(I, 7) = "1" Then
             MSHFlexGrid6.TextMatrix(I, 7) = "Spillage"
          Else
             MSHFlexGrid6.TextMatrix(I, 7) = "Normal"
          End If
          
          If Val(MSHFlexGrid6.TextMatrix(I, 2)) < 0 Then
             MSHFlexGrid6.TextMatrix(I, 2) = 0
             MSHFlexGrid6.TextMatrix(I, 14) = 0
             MSHFlexGrid6.TextMatrix(I, 16) = 0
          End If
          
          MSHFlexGrid6.TextMatrix(I, 16) = Format(Val(MSHFlexGrid6.TextMatrix(I, 14)) * Val(MSHFlexGrid6.TextMatrix(I, 15)) * Val(MSHFlexGrid6.TextMatrix(I, 9)), "#####0.00")
          
          TmpRs.MoveNext
          MSHFlexGrid6.Rows = MSHFlexGrid6.Rows + 1
      Next
   End If
End If
End Sub


'Private Sub PDF_Print_Rpt()
'Set objPDF = New mjwPDF
'
'' Create a simple PDF file using the mjwPDF class
'
'Gen_mTimeStamp = GetTimeStamp
''Call Create_Xls(Gen_mTimeStamp & Gen_UserName & "_CDTF-" & Trim(TxtPreFix) & "" & Trim(TxtCxTFNo) & ".xls")
'
'
'' Set the PDF title and filename
'objPDF.PDFTitle = "CDTF PDF Document"
'objPDF.PDFFileName = App.Path & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "_CITF-" & Trim(TxtPreFix) & "" & Trim(TxtCxTFNo) & ".pdf"
'mPDFFileName = App.Path & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "_CITF-" & Trim(TxtPreFix) & "" & Trim(TxtCxTFNo) & ".pdf"
'' We must tell the class where the PDF fonts are located
'objPDF.PDFLoadAfm = App.Path & "\Fonts"
'
'' Set the file properties
'objPDF.PDFSetLayoutMode = LAYOUT_DEFAULT
'objPDF.PDFFormatPage = FORMAT_A4
'objPDF.PDFOrientation = ORIENT_PORTRAIT
'objPDF.PDFSetUnit = UNIT_PT
'objPDF.PDFSetViewerPreferences = VIEW_HIDEMENUBAR & VIEW_HIDETOOLBAR & VIEW_HIDEWINDOWUI & VIEW_HIDERIGHTCLICK '& VIEW_DISABLEFILEMENU
'
'
'' Lets us set see the bookmark pane when we view the PDF
''objPDF.PDFUseOutlines = True
'
'' View the PDF file after we create it
'objPDF.PDFView = True
'
'' Begin our PDF document
'
'objPDF.PDFBeginDoc
'    ' Lets add a heading
'objPDF.PDFSetFont FONT_ARIAL, 12, FONT_BOLD
'objPDF.PDFSetDrawColor = vbWhite
'objPDF.PDFSetTextColor = vbBlack
'objPDF.PDFSetAlignement = ALIGN_CENTER
'objPDF.PDFSetBorder = BORDER_ALL
'objPDF.PDFSetFill = True
'
'objPDF.PDFCell " ", 15, 15, _
'objPDF.PDFGetPageWidth - 30, objPDF.PDFGetPageHeight - 20
'
'If PntPrv = "Y" Then
'   objPDF.PDFWatermark = "Preview"
'Else
'   objPDF.PDFWatermark = "NCMSL"
'End If
'
'objPDF.PDFSetFont FONT_ARIAL, 12, FONT_BOLD
'objPDF.PDFSetDrawColor = vbWhite
'objPDF.PDFSetTextColor = vbBlack
'objPDF.PDFSetAlignement = ALIGN_CENTER
'objPDF.PDFSetBorder = BORDER_ALL
'objPDF.PDFSetFill = True
'
'
'pgval = ((objPDF.PDFGetPageWidth - 18))
'
''objPDF.PDFDrawPolygon Array(12, 15, objPDF.PDFGetPageWidth - 18, 15, objPDF.PDFGetPageWidth - 18, 30, 12, 30)
'
'mRowHeight = 15
'mGlobalLeft = 12
''objPDF.PDFDrawPolygon Array(12, 15, objPDF.PDFGetPageWidth - 18, 15, objPDF.PDFGetPageWidth - 18, 30, 12, 30)
'mLeft = mGlobalLeft
'mTop = 15
'mHeight = mTop + mRowHeight
'mTxtFromTop = mRowHeight + 2
'
' 'Call DrawTable(10, 2, 15, 30, 12)
'
'Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
'
'mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
'mStrLength = Len("NCMSL " & TxtCMPID & "")
'mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
'mTxtLeft = mPrintPoint - (mStrLength * 2) '* 5) '270
'mFontSize = 12
'
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_CENTER, mTxtLeft, mTxtTop, "NCMSL " & TxtCMPID.Text & "")
'
'mHeight = mTop + mRowHeight
'mLeft = mGlobalLeft
'
'
'Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
'
'mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
'mStrLength = Len("Commodity Issue Transaction Form (CITF)")
'mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
'mTxtLeft = mPrintPoint - (mStrLength * 2) '* 5) '270
'mFontSize = 10
'
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_CENTER, mTxtLeft, mTxtTop, "Commodity Issue Transaction Form (CITF)")
'
'mHeight = mTop + mRowHeight
'mLeft = mGlobalLeft
'
'Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
'
'
'mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
'mStrLength = Len("(This is NOT an instrument of title of the goods)")
'mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
'mTxtLeft = mPrintPoint '- (mStrLength) '* 5) '270
'mFontSize = 6
'
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_CENTER, mTxtLeft - 30, mTxtTop - 3, "(This is NOT an instrument of title of the goods)")
'
''--Code added By Kishor
''Row 1
'
'
'mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
'mStrLength = Len("Depositor Details")
'mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
'mTxtLeft = mPrintPoint '- (mStrLength) '* 5) '270
'mFontSize = 7
'Dim mDispTextSize  As Integer
'mDispTextSize = 6
'
'mTxtTop = mTxtTop + mRowHeight
'
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "CITF No")
'Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop, TxtCxTFNo)
''--
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 280, mTxtTop, "Date")
'Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 430, mTxtTop, Format(TxtCxTFDate.Value, "dd-MMM-yyyy"))
''--
'mTxtTop = mTxtTop + mRowHeight
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "Location")
'Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop, GetLocationFromCMP(Val(mCMPID)))
''--
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 280, mTxtTop, "State")
'Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 430, mTxtTop, GetStateFromLocation(GetLocationIDFromCMP(Val(mCMPID))))
''--
'mTxtTop = mTxtTop + mRowHeight
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "CMPID & Address")
'Dim mCMPIDnAddr As String
'mCMPIDnAddr = mCMPID & "-" & GetCMPAddress(Val(mCMPID), "A")
'Dim mLen As Integer, mMax As Integer, mPara As Integer
'Dim t1 As Double
'mMax = 35
'mPara = 0
't1 = 1
'mLen = Len(mCMPIDnAddr)
'mLen = Round(mLen / mMax) '+ 1
'
'For t = 1 To mLen
'    If t > 1 Then
'       t1 = mMax + t1
'       'mMax '= mMax + 35
'       mPara = mPara + 7
'    End If
'    Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop + mPara, Mid$(mCMPIDnAddr, t1, mMax))
'Next
'
''--
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 280, mTxtTop, "Commodity")
'Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 430, mTxtTop, TxtCommodityID)
'
'Call DrawTable(2, 4, mTop, mHeight, mLeft, mRowHeight)
'mHeight = mHeight + 15
'Call DrawTable(, 4, mTop, mHeight, mLeft, mRowHeight)
'
'
''Row 2
'
'mTxtTop = mHeight - (mTxtFromTop) + 15 'mTop ' + 15
'mStrLength = Len("Transaction Tyepe")
'mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
'mTxtLeft = mPrintPoint '- (mStrLength) '* 5) '270
'mFontSize = 7
'
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "Transaction Type")
'Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop, CmbExchangeTypeID)
'
'mTxtTop = mTxtTop + mRowHeight
'
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "Name Of Agent")
'Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop, TxtAgent)
'
'objPDF.PDFDrawLine mGlobalLeft + 132, mTop - 4 * mRowHeight, mGlobalLeft + 132, mHeight + mRowHeight
'Call DrawTable(2, , mTop, mHeight, mLeft, mRowHeight)
'
'
''Row 3
'
'mTxtTop = mHeight '- (mTxtFromTop) 'mTop ' + 15
'mStrLength = Len("Withdrawer Details")
'mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
'mTxtLeft = mPrintPoint '- (mStrLength) '* 5) '270
'mFontSize = 8
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mLeft, mTxtTop - 2, "Withdrawer Details")
'
'Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
'
'objPDF.PDFDrawLine mGlobalLeft + 72, mTop, mGlobalLeft + 72, mTop + ((MSHFlexGrid2.Rows - 2) * mRowHeight)
'objPDF.PDFDrawLine mGlobalLeft + 142, mTop, mGlobalLeft + 142, mTop + ((MSHFlexGrid2.Rows - 2) * mRowHeight)
'If mExchangeTypeID <> 2 Then
'   objPDF.PDFDrawLine 2.8 * (mGlobalLeft + 142), mTop, 2.8 * (mGlobalLeft + 142), mTop + ((MSHFlexGrid2.Rows - 1) * mRowHeight)
'End If
'objPDF.PDFDrawLine 3.3 * (mGlobalLeft + 142), mTop, 3.3 * (mGlobalLeft + 142), mTop + ((MSHFlexGrid2.Rows - 1) * mRowHeight)
'
'Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
'
'
'mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
'mStrLength = Len("Depositor Details")
'mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
'mTxtLeft = mPrintPoint '- (mStrLength) '* 5) '270
'mFontSize = 7
'
'If CmbExchangeTypeID.Text = "NCDEX" Then
'   If Check1.Value = 0 Then
'      Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mLeft, mTxtTop, "RRN No")
'   Else
'
'      Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mLeft, mTxtTop, "Rejected Withdrawal")
'   End If
'ElseIf CmbExchangeTypeID.Text = "Non NCDEX" Then
'
'   Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mLeft, mTxtTop, "WHR No")
'ElseIf CmbExchangeTypeID.Text = "SPOT" Then
'
'   Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mLeft, mTxtTop, "CDC No")
'End If
'
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mLeft + 80, mTxtTop, "Client ID")
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 150, mTxtTop, "Client Name")
'If mExchangeTypeID <> 2 Then
'   Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 430, mTxtTop, "No of Bags")
'   Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 500, mTxtTop - 6, "Electronic Weight")
'   Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 500, mTxtTop, "in MT")
'
'Else
'   Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 500, mTxtTop - 6, "Electronic Weight")
'   Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 500, mTxtTop, "in MT")
'
'End If
'
'objPDF.PDFDrawLine mGlobalLeft + 72, mTop, mGlobalLeft + 72, mTop + ((MSHFlexGrid2.Rows - 2) * mRowHeight)
'objPDF.PDFDrawLine mGlobalLeft + 142, mTop, mGlobalLeft + 142, mTop + ((MSHFlexGrid2.Rows - 2) * mRowHeight)
'If mExchangeTypeID <> 2 Then
'   objPDF.PDFDrawLine 2.8 * (mGlobalLeft + 142), mTop, 2.8 * (mGlobalLeft + 142), mTop + ((MSHFlexGrid2.Rows - 1) * mRowHeight)
'End If
'objPDF.PDFDrawLine 3.3 * (mGlobalLeft + 142), mTop, 3.3 * (mGlobalLeft + 142), mTop + ((MSHFlexGrid2.Rows - 1) * mRowHeight)
'objPDF.PDFDrawLine mGlobalLeft, mTop + mRowHeight, mGlobalLeft + 560, mTop + mRowHeight
'Call DrawTable(1, , mTop, mHeight, mLeft, mRowHeight)
'
''FILL GRID WITH DETAILS
'Dim mTemp As Integer
'For t = 2 To MSHFlexGrid2.Rows - 1
'    If mHeight > objPDF.PDFGetPageHeight Then
'       Call CheckForNewPage
'       mTop = 15
'       mTxtTop = 10
'    End If
'    If t = 1 Then
'       mTemp = mFontSize
'    Else
'       mTemp = mDispTextSize
'    End If
'    mTxtTop = mTxtTop + mRowHeight
'
'    Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mLeft, mTxtTop, MSHFlexGrid2.TextMatrix(t, 1))
'    Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 80, mTxtTop, MSHFlexGrid2.TextMatrix(t, 3))
'    Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 150, mTxtTop, MSHFlexGrid2.TextMatrix(t, 4))
'    If mExchangeTypeID <> 2 Then
'       Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 430, mTxtTop, Format(MSHFlexGrid2.TextMatrix(t, 15), "#######0.000"))
'    End If
'    Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 500, mTxtTop, Format(MSHFlexGrid2.TextMatrix(t, 9), "#######0.000"))
'
'    objPDF.PDFDrawLine mGlobalLeft + 72, mTop, mGlobalLeft + 72, mTop + mRowHeight
'    objPDF.PDFDrawLine mGlobalLeft + 142, mTop, mGlobalLeft + 142, mTop + mRowHeight
'    If mExchangeTypeID <> 2 Then
'       objPDF.PDFDrawLine 2.8 * (mGlobalLeft + 142), mTop, 2.8 * (mGlobalLeft + 142), mTop + mRowHeight
'    End If
'    objPDF.PDFDrawLine 3.3 * (mGlobalLeft + 142), mTop, 3.3 * (mGlobalLeft + 142), mTop + mRowHeight
'
'    'Call DrawTable(MSHFlexGrid2.Rows - 1, , mTop, mHeight, mLeft, mRowHeight)
'    objPDF.PDFDrawLine mGlobalLeft, mTop + mRowHeight, mGlobalLeft + 565, mTop + mRowHeight
'
'    mTop = mTop + mRowHeight
'    mHeight = mHeight + mRowHeight
'Next
'
'Call CheckForNewPage
'mTxtTop = mTxtTop + mRowHeight
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 220, mTxtTop, "Total")
'If mExchangeTypeID <> 2 Then
'   Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 430, mTxtTop, Format(MSHFlexGrid2.TextMatrix(0, 15), "#######0.000"))
'End If
'Call CheckForNewPage
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 500, mTxtTop, Format(MSHFlexGrid2.TextMatrix(0, 9), "#######0.000"))
'Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
'
''mTxtTop = mTxtTop + mRowHeight
'mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
'mStrLength = Len("Truck-wise Withdrawal Details")
'mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
'mTxtLeft = mPrintPoint '- (mStrLength) '* 5) '270
'mFontSize = 7
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mLeft, mTxtTop - 2, "Truck-wise Withdrawal Details")
'''objPDF.PDFDrawLine mGlobalLeft, mTxtTop + mRowHeight, mGlobalLeft + 565, mTxtTop + mRowHeight
'''objPDF.PDFDrawLine mGlobalLeft + 50, mTop, mGlobalLeft + 50, mTop + ((MSHFlexGrid3.Rows - 1) * mRowHeight)
'''objPDF.PDFDrawLine (mGlobalLeft + 120), mTop, (mGlobalLeft + 120), mTop + ((MSHFlexGrid3.Rows - 1) * mRowHeight)
'''objPDF.PDFDrawLine (mGlobalLeft + 190), mTop, (mGlobalLeft + 190), mTop + ((MSHFlexGrid3.Rows + 6) * mRowHeight)
'''objPDF.PDFDrawLine (mGlobalLeft + 256), mTop, (mGlobalLeft + 256), mTop + ((MSHFlexGrid3.Rows + 6) * mRowHeight)
'''objPDF.PDFDrawLine (mGlobalLeft + 356), mTop, (mGlobalLeft + 356), mTop + ((MSHFlexGrid3.Rows - 1) * mRowHeight)
'''
'''Dim mTempTop As Integer
'''mTempTop = mTxtTop - 3
'''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 90, mTempTop + (MSHFlexGrid3.Rows * mRowHeight), "Total")
'''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 60, mTempTop + ((MSHFlexGrid3.Rows + 1) * mRowHeight), "Less Gunny Weight(MT)")
'''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 50, mTempTop + ((MSHFlexGrid3.Rows + 2) * mRowHeight), "Less Standard Deductions(MT)")
'''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 60, mTempTop + ((MSHFlexGrid3.Rows + 3) * mRowHeight), "Add Adjustments(MT)")
'''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 60, mTempTop + ((MSHFlexGrid3.Rows + 4) * mRowHeight), "Electronic Weight(MT)")
''''pp 'Hologram No
'''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 270, mTempTop + ((MSHFlexGrid3.Rows + 4) * mRowHeight), "Hologram No")
'''
''''Actual Values
'''Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 200, mTempTop + (MSHFlexGrid3.Rows * mRowHeight), MSHFlexGrid3.TextMatrix(0, 5))
'''Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 200, mTempTop + ((MSHFlexGrid3.Rows + 1) * mRowHeight), MSHFlexGrid4.TextMatrix(0, 8))
'''Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 200, mTempTop + ((MSHFlexGrid3.Rows + 2) * mRowHeight), MSHFlexGrid4.TextMatrix(0, 9))
'''Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 200, mTempTop + ((MSHFlexGrid3.Rows + 3) * mRowHeight), MSHFlexGrid4.TextMatrix(0, 6))
'''Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 200, mTempTop + ((MSHFlexGrid3.Rows + 4) * mRowHeight), MSHFlexGrid4.TextMatrix(0, 10))
''''Hologram No -- Use Record Set
'''
'''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 370, mTempTop + ((MSHFlexGrid3.Rows + 4) * mRowHeight), " " & RsTxn_CxTF_Details!HologramNo & " ")
'''
'''Dim mTemp As Integer
'''For t = 1 To MSHFlexGrid3.Rows - 1
'''    If t = 1 Then
'''       mTemp = mFontSize
'''    Else
'''       mTemp = mDispTextSize
'''    End If
'''    mTxtTop = mTxtTop + mRowHeight
'''    Call PrintText(mTemp, FONT_NORMAL, ALIGN_LEFT, mLeft, mTxtTop, MSHFlexGrid3.TextMatrix(t, 1))
'''    Call PrintText(mTemp, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 60, mTxtTop, MSHFlexGrid3.TextMatrix(t, 3))
'''    Call PrintText(mTemp, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 130, mTxtTop, MSHFlexGrid3.TextMatrix(t, 4))
'''    Call PrintText(mTemp, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 200, mTxtTop, MSHFlexGrid3.TextMatrix(t, 5))
'''    Call PrintText(mTemp, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 256, mTxtTop, MSHFlexGrid3.TextMatrix(t, 2))
'''    Call PrintText(mTemp, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 356, mTxtTop, MSHFlexGrid3.TextMatrix(t, 6))
'''Next
'''
'''Call DrawTable(MSHFlexGrid3.Rows + 6, , mTop, mHeight, mLeft, mRowHeight)
''''Clear Line
'''mTxtTop = mHeight - (mTxtFromTop)
'''objPDF.PDFSetTextColor = vbWhite
'''
'''Call PrintText(15, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 255.2, mTxtTop - (6 * mRowHeight), "_____________________________________")
'''Call PrintText(15, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 255.2, mTxtTop - (5 * mRowHeight), "_____________________________________")
'''Call PrintText(15, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 255.2, mTxtTop - (4 * mRowHeight), "_____________________________________")
'''
'''objPDF.PDFSetTextColor = vbBlack
'''Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 256, mTxtTop - (3 * mRowHeight), "Affix Hologram Here")
''''--
'''
'''objPDF.PDFDrawLine (mGlobalLeft + 50), mTop - (2 * mRowHeight), (mGlobalLeft + 50), mTop
'''objPDF.PDFDrawLine (mGlobalLeft + 120), mTop - (2 * mRowHeight), (mGlobalLeft + 120), mTop
'''
'''objPDF.PDFDrawLine (mGlobalLeft + 360), mTop - (3 * mRowHeight), (mGlobalLeft + 360), mTop
'''objPDF.PDFDrawLine (mGlobalLeft + 420), mTop - (2 * mRowHeight), (mGlobalLeft + 420), mTop
'''
'''If mHeight > objPDF.PDFGetPageHeight - 50 Then
'''   objPDF.PDFEndPage
'''   objPDF.PDFNewPage
'''   objPDF.PDFSetFont FONT_ARIAL, 12, FONT_BOLD
'''   objPDF.PDFSetDrawColor = vbWhite
'''   objPDF.PDFSetTextColor = vbBlack
'''   objPDF.PDFSetAlignement = ALIGN_CENTER
'''   objPDF.PDFSetBorder = BORDER_ALL
'''   objPDF.PDFSetFill = True
'''
'''   objPDF.PDFCell " ", 15, 15, _
'''   objPDF.PDFGetPageWidth - 30, objPDF.PDFGetPageHeight - 20
'''   mLeft = mGlobalLeft
'''   mTop = 15
'''   mHeight = mTop + mRowHeight
'''   mTxtFromTop = mRowHeight + 2
'''End If
'''
'''Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
'''mTxtTop = mHeight - (mTxtFromTop)
'''
'''Call PrintText(5.4, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "NOTE: THIS DOCUMENT IS NOT VALID WITHOUT A HOLOGRAM OF NCMSL AFFIXED ABOVE & THE SERIAL NUMBER IN THE HOLOGRAM SHOULD MATCH WITH THE HOLOGRAM NUMBER PRINTED ABOVE.")
'''
'''mHeight = mHeight + 50
'''
'''Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
'''mTxtTop = mHeight - (mTxtFromTop) '+ 15
'''
'''Call PrintText(9, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "Name & Signature of CMP Manager/CMP DRC:")
'''
'''mHeight = mHeight + 30
'''Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
'''mTxtTop = mHeight - (mTxtFromTop)
'''
'''Call PrintText(6.5, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop - 17, "I/We hereby declare that the commodities being deposited herein belong to the clients as stated above and are free from any encumbrances. I/We request  you to take custody of the said")
'''Call PrintText(6.5, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop - 10, "commodities and I/We undertake to indemnify you against any losses incurred by you due to any false information provided by me/us. I/We also undertake to pay all required charges as per")
'''Call PrintText(6.5, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop - 3, "your charges schedule from time to time. I have gone throught the attached 'Terms and Conditions'** of storage and agree to abide by the same.")
'''
''''Call CheckForNewPage
''''objPDF.PDFDrawLine mGlobalLeft + 50, mTop, mGlobalLeft + 50, mTop + ((MSHFlexGrid3.Rows - 1) * mRowHeight)
''''objPDF.PDFDrawLine (mGlobalLeft + 120), mTop, (mGlobalLeft + 120), mTop + ((MSHFlexGrid3.Rows - 1) * mRowHeight)
''''objPDF.PDFDrawLine (mGlobalLeft + 190), mTop, (mGlobalLeft + 190), mTop + ((MSHFlexGrid3.Rows) * mRowHeight)
''''objPDF.PDFDrawLine (mGlobalLeft + 256), mTop, (mGlobalLeft + 256), mTop + ((MSHFlexGrid3.Rows + 6) * mRowHeight)
''''objPDF.PDFDrawLine (mGlobalLeft + 356), mTop, (mGlobalLeft + 356), mTop + ((MSHFlexGrid3.Rows - 1) * mRowHeight)
''''
''''
'''''Dim mTempTop As Integer
'''''mTempTop = mTxtTop - 3
'''''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 90, mTempTop + (MSHFlexGrid3.Rows * mRowHeight), "Total")
''''
'''''----Commented on 18-01-2011 by Sagar
''''
'''''''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 60, mTempTop + ((MSHFlexGrid3.Rows + 1) * mRowHeight), "Less Gunny Weight(MT)")
'''''''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 50, mTempTop + ((MSHFlexGrid3.Rows + 2) * mRowHeight), "Less Standard Deductions(MT)")
'''''''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 60, mTempTop + ((MSHFlexGrid3.Rows + 3) * mRowHeight), "Add Adjustments(MT)")
'''''''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 60, mTempTop + ((MSHFlexGrid3.Rows + 4) * mRowHeight), "Electronic Weight(MT)")
''''
''''objPDF.PDFDrawLine (mGlobalLeft), mTempTop + (MSHFlexGrid3.Rows + 4 * mRowHeight), (mGlobalLeft + 256), mTempTop + (MSHFlexGrid3.Rows + 4 * mRowHeight)
''''
'''''pp 'Hologram No
''''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 270, mTempTop + ((MSHFlexGrid3.Rows + 4) * mRowHeight), "Hologram No")
''''
'''''Actual Values
''''Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 200, mTempTop + (MSHFlexGrid3.Rows * mRowHeight), MSHFlexGrid3.TextMatrix(0, 5))
''''
'''''----Commented on 18-01-2011 by Sagar
''''
'''''''Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 200, mTempTop + ((MSHFlexGrid3.Rows + 1) * mRowHeight), MSHFlexGrid4.TextMatrix(0, 8))
'''''''Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 200, mTempTop + ((MSHFlexGrid3.Rows + 2) * mRowHeight), MSHFlexGrid4.TextMatrix(0, 9))
'''''''Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 200, mTempTop + ((MSHFlexGrid3.Rows + 3) * mRowHeight), MSHFlexGrid4.TextMatrix(0, 6))
'''''''Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 200, mTempTop + ((MSHFlexGrid3.Rows + 4) * mRowHeight), MSHFlexGrid4.TextMatrix(0, 10))
'''''Hologram No -- Use Record Set
''''
''''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 370, mTempTop + ((MSHFlexGrid3.Rows + 4) * mRowHeight), " " & RsTxn_CxTF_Details!HologramNo & " ")
''''
'
'For t = 1 To MSHFlexGrid3.Rows - 1
'    If mHeight > objPDF.PDFGetPageHeight Then
'       Call CheckForNewPage
'       mTop = 15
'       mTxtTop = 10
'    End If
'    If t = 1 Then
'       mTemp = mFontSize
'    Else
'       mTemp = mDispTextSize
'    End If
'    mTxtTop = mTxtTop + mRowHeight
'
'    Call PrintText(mTemp, FONT_NORMAL, ALIGN_LEFT, mLeft, mTxtTop, MSHFlexGrid3.TextMatrix(t, 1))
'    Call PrintText(mTemp, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 60, mTxtTop, MSHFlexGrid3.TextMatrix(t, 3))
'    Call PrintText(mTemp, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 130, mTxtTop, MSHFlexGrid3.TextMatrix(t, 4))
'    Call PrintText(mTemp, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 200, mTxtTop, MSHFlexGrid3.TextMatrix(t, 5))
'    Call PrintText(mTemp, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 256, mTxtTop, MSHFlexGrid3.TextMatrix(t, 2))
'    Call PrintText(mTemp, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 356, mTxtTop, MSHFlexGrid3.TextMatrix(t, 6))
'
'    objPDF.PDFDrawLine mGlobalLeft + 50, mTop, mGlobalLeft + 50, mTop + mRowHeight
'    objPDF.PDFDrawLine (mGlobalLeft + 120), mTop, (mGlobalLeft + 120), mTop + mRowHeight
'    objPDF.PDFDrawLine (mGlobalLeft + 190), mTop, (mGlobalLeft + 190), mTop + mRowHeight
'    objPDF.PDFDrawLine (mGlobalLeft + 256), mTop, (mGlobalLeft + 256), mTop + mRowHeight
'    objPDF.PDFDrawLine (mGlobalLeft + 356), mTop, (mGlobalLeft + 356), mTop + mRowHeight
'    objPDF.PDFDrawLine mGlobalLeft, mTop + mRowHeight, mGlobalLeft + 565, mTop + mRowHeight
'
'    mTop = mTop + mRowHeight
'    mHeight = mHeight + mRowHeight
'Next
''Call CheckForNewPage
''Call DrawTable(MSHFlexGrid3.Rows - 1, , mTop, mHeight, mLeft, mRowHeight)
'Call CheckForNewPage
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 90, mTop + mRowHeight - 3, "Total")
'Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 200, mTop + mRowHeight - 3, MSHFlexGrid3.TextMatrix(0, 5))
'objPDF.PDFDrawLine (mGlobalLeft + 190), mTop, (mGlobalLeft + 190), mTop + mRowHeight
'objPDF.PDFDrawLine (mGlobalLeft + 256), mTop, (mGlobalLeft + 256), mTop + mRowHeight * 4
'
'Call CheckForNewPage
'objPDF.PDFDrawLine (mGlobalLeft), mTop + mRowHeight, (mGlobalLeft + 256), mTop + mRowHeight
'
'mTop = mTop + (mRowHeight * 4)
'objPDF.PDFDrawLine mGlobalLeft + 50, mTop + mRowHeight, mGlobalLeft + 50, mTop + (mRowHeight * 3)
'objPDF.PDFDrawLine (mGlobalLeft + 120), mTop + mRowHeight, (mGlobalLeft + 120), mTop + (mRowHeight * 3)
'objPDF.PDFDrawLine (mGlobalLeft + 190), mTop + mRowHeight, (mGlobalLeft + 190), mTop + (mRowHeight * 3)
'objPDF.PDFDrawLine (mGlobalLeft + 256), mTop, (mGlobalLeft + 256), mTop + (mRowHeight * 3)
'objPDF.PDFDrawLine (mGlobalLeft + 356), mTop, (mGlobalLeft + 356), mTop + (mRowHeight * 3)
'objPDF.PDFDrawLine (mGlobalLeft + 420), mTop + mRowHeight, (mGlobalLeft + 420), mTop + (mRowHeight * 3)
'Call CheckForNewPage
'Call PrintText(6, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 265, mTop - 3, "Affix Hologram Here")
'
'If mExchangeTypeID = 2 Then
'   Call CheckForNewPage
'   Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 70, mTop + mRowHeight - 3, "Total No. of Bags")
'   Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 200, mTop + mRowHeight - 3, Format(MSHFlexGrid4.TextMatrix(0, 6)))
'   objPDF.PDFDrawLine (mGlobalLeft + 190), mTop, (mGlobalLeft + 190), mTop + mRowHeight
'End If
'
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 270, mTop + mRowHeight - 3, "Hologram No")
'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 370, mTop + mRowHeight - 3, " " & RsTxn_CxTF_Details!HologramNo & " ")
'Call CheckForNewPage
'Call DrawTable(4, , mTop, mTop, mLeft, mRowHeight)
'
''If mHeight > objPDF.PDFGetPageHeight - 50 Then
''   objPDF.PDFEndPage
''   objPDF.PDFNewPage
''   objPDF.PDFSetFont FONT_ARIAL, 12, FONT_BOLD
''   objPDF.PDFSetDrawColor = vbWhite
''   objPDF.PDFSetTextColor = vbBlack
''   objPDF.PDFSetAlignement = ALIGN_CENTER
''   objPDF.PDFSetBorder = BORDER_ALL
''   objPDF.PDFSetFill = True
''
''   objPDF.PDFCell " ", 15, 15, _
''   objPDF.PDFGetPageWidth - 30, objPDF.PDFGetPageHeight - 20
''   mLeft = mGlobalLeft
''   mTop = 15
''   mHeight = mTop + mRowHeight
''   mTxtFromTop = mRowHeight + 2
''End If
'
'mHeight = mTop
'Call CheckForNewPage
'Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
'Call CheckForNewPage
'mTxtTop = mHeight - (mTxtFromTop)
''''
'
'Call PrintText(5.4, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "NOTE: THIS DOCUMENT IS NOT VALID WITHOUT A HOLOGRAM OF NCMSL AFFIXED ABOVE & THE SERIAL NUMBER IN THE HOLOGRAM SHOULD MATCH WITH THE HOLOGRAM NUMBER PRINTED ABOVE.")
''''
'mHeight = mHeight + 50
''''
'Call CheckForNewPage
'Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
'
''''
'Call CheckForNewPage
'mTxtTop = mHeight - (mTxtFromTop) '+ 15
'Call PrintText(9, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "Name & Signature of CMP Manager/CMP DRC:")
'Call CheckForNewPage
'Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
'Call CheckForNewPage
'mTxtTop = mTxtTop + mRowHeight
'Call PrintText(7, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "Received the above mentioned goods in good condition.")
''mHeight = mHeight + 30
''Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
'
''''
'''Call PrintText(6.5, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop - 17, "I/We hereby declare that the commodities being deposited herein belong to the clients as stated above and are free from any encumbrances. I/We request  you to take custody of the said")
'''Call PrintText(6.5, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop - 10, "commodities and I/We undertake to indemnify you against any losses incurred by you due to any false information provided by me/us. I/We also undertake to pay all required charges as per")
'''Call PrintText(6.5, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop - 3, "your charges schedule from time to time. I have gone throught the attached 'Terms and Conditions'** of storage and agree to abide by the same.")
''''''
'
'mHeight = mHeight + 70
'
'Call CheckForNewPage
'
'Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
'Call CheckForNewPage
'mTxtTop = mHeight - (mTxtFromTop)
'Call PrintText(9, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "Name & Signature of Withdrawer/Agent of Withdrawer:")
'
'mHeight = mHeight + 50
''Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
'
'Call CheckForNewPage
'mTxtTop = mHeight - (mTxtFromTop)
'Call PrintText(8, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "**'Terms and Conditions' can be viewed in http://www.ncmsl.com/whrtNc.html")
'
'objPDF.PDFEndDoc
'
'End Sub
Private Sub PDF_Print_Rpt()
Totpages = GetTotalPDFPages()
Set objPDF = New mjwPDF

' Create a simple PDF file using the mjwPDF class

Gen_mTimeStamp = GetTimeStamp
'Call Create_Xls(Gen_mTimeStamp & Gen_UserName & "_CDTF-" & Trim(TxtPreFix) & "" & Trim(TxtCxTFNo) & ".xls")
    
    
' Set the PDF title and filename
objPDF.PDFTitle = "CITF PDF Document"
objPDF.PDFFileName = App.Path & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "_CITF-" & Trim(TxtPreFix) & "" & Trim(TxtCxTFNo) & ".pdf"
mPDFFileName = App.Path & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "_CITF-" & Trim(TxtPreFix) & "" & Trim(TxtCxTFNo) & ".pdf"
' We must tell the class where the PDF fonts are located
objPDF.PDFLoadAfm = App.Path & "\Fonts"

' Set the file properties
objPDF.PDFSetLayoutMode = LAYOUT_DEFAULT
objPDF.PDFFormatPage = FORMAT_A4
objPDF.PDFOrientation = ORIENT_PORTRAIT
objPDF.PDFSetUnit = UNIT_PT
objPDF.PDFSetViewerPreferences = VIEW_HIDEMENUBAR & VIEW_HIDETOOLBAR & VIEW_HIDEWINDOWUI & VIEW_HIDERIGHTCLICK '& VIEW_DISABLEFILEMENU


' Lets us set see the bookmark pane when we view the PDF
'objPDF.PDFUseOutlines = True

' View the PDF file after we create it
objPDF.PDFView = True

' Begin our PDF document

objPDF.PDFBeginDoc
    ' Lets add a heading
objPDF.PDFSetFont FONT_ARIAL, 12, FONT_BOLD
objPDF.PDFSetDrawColor = vbWhite
objPDF.PDFSetTextColor = vbBlack
objPDF.PDFSetAlignement = ALIGN_CENTER
objPDF.PDFSetBorder = BORDER_ALL
objPDF.PDFSetFill = True

objPDF.PDFCell " ", 15, 15, _
objPDF.PDFGetPageWidth - 30, objPDF.PDFGetPageHeight - 20

If PntPrv = "Y" Then
   objPDF.PDFWatermark = "Preview"
Else
   objPDF.PDFWatermark = "NCMSL"
End If

objPDF.PDFSetFont FONT_ARIAL, 12, FONT_BOLD
objPDF.PDFSetDrawColor = vbWhite
objPDF.PDFSetTextColor = vbBlack
objPDF.PDFSetAlignement = ALIGN_CENTER
objPDF.PDFSetBorder = BORDER_ALL
objPDF.PDFSetFill = True


pgval = ((objPDF.PDFGetPageWidth - 18))

'''objPDF.PDFDrawPolygon Array(12, 15, objPDF.PDFGetPageWidth - 18, 15, objPDF.PDFGetPageWidth - 18, 30, 12, 30)
''
''mRowHeight = 15
''mGlobalLeft = 12
'''objPDF.PDFDrawPolygon Array(12, 15, objPDF.PDFGetPageWidth - 18, 15, objPDF.PDFGetPageWidth - 18, 30, 12, 30)
''mLeft = mGlobalLeft
''mTop = 15
''mHeight = mTop + mRowHeight
''mTxtFromTop = mRowHeight + 2
''
'' 'Call DrawTable(10, 2, 15, 30, 12)
''
''Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
''
''mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
''mStrLength = Len("NCMSL " & TxtCMPID & "")
''mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
''mTxtLeft = mPrintPoint - (mStrLength * 2) '* 5) '270
''mFontSize = 12
''
''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_CENTER, mTxtLeft, mTxtTop, "NCMSL " & TxtCMPID.Text & "")
''
''mHeight = mTop + mRowHeight
''mLeft = mGlobalLeft
''
''
''Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
''
''mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
''mStrLength = Len("Commodity Issue Transaction Form (CITF)")
''mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
''mTxtLeft = mPrintPoint - (mStrLength * 2) '* 5) '270
''mFontSize = 10
''
''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_CENTER, mTxtLeft, mTxtTop, "Commodity Issue Transaction Form (CITF)")
''
''mHeight = mTop + mRowHeight
''mLeft = mGlobalLeft
''
''Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
''
''
''mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
''mStrLength = Len("(This is NOT an instrument of title of the goods)")
''mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
''mTxtLeft = mPrintPoint '- (mStrLength) '* 5) '270
''mFontSize = 6
''
''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_CENTER, mTxtLeft - 30, mTxtTop - 3, "(This is NOT an instrument of title of the goods)")
''
'''--Code added By Kishor
'''Row 1
''
''
''mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
''mStrLength = Len("Depositor Details")
''mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
''mTxtLeft = mPrintPoint '- (mStrLength) '* 5) '270
''mFontSize = 7
''Dim mDispTextSize  As Integer
''mDispTextSize = 6
''
''mTxtTop = mTxtTop + mRowHeight
''
''Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop, "CITF No")
''Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop, TxtCxTFNo)
'''--
''Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 280, mTxtTop, "Date")
''Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 430, mTxtTop, Format(TxtCxTFDate.Value, "dd-MMM-yyyy"))
'''--
''mTxtTop = mTxtTop + mRowHeight
''Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop, "Location")
''Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop, GetLocationFromCMP(Val(mCMPID)))
'''--
''Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 280, mTxtTop, "State")
''Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 430, mTxtTop, GetStateFromLocation(GetLocationIDFromCMP(Val(mCMPID))))
'''--
''mTxtTop = mTxtTop + mRowHeight
''Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop, "CMPID & Address")
''Dim mCMPIDnAddr As String
''mCMPIDnAddr = mCMPID & "-" & GetCMPAddress(Val(mCMPID), "A")
''Dim mLen As Integer, mMax As Integer, mPara As Integer
''Dim t1 As Double
''mMax = 35
''mPara = 0
''t1 = 1
''mLen = Len(mCMPIDnAddr)
''mLen = Round(mLen / mMax) '+ 1
''
''For t = 1 To mLen
''    If t > 1 Then
''       t1 = mMax + t1
''       'mMax '= mMax + 35
''       mPara = mPara + 7
''    End If
''    Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop + mPara, Mid$(mCMPIDnAddr, t1, mMax))
''Next
''
'''--
''Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 280, mTxtTop, "Commodity")
''Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 430, mTxtTop, TxtCommodityID)
''
''Call DrawTable(2, 4, mTop, mHeight, mLeft, mRowHeight)
''mHeight = mHeight + 15
''Call DrawTable(, 4, mTop, mHeight, mLeft, mRowHeight)
''
''
'''Row 2
''
''mTxtTop = mHeight - (mTxtFromTop) + 15 'mTop ' + 15
''mStrLength = Len("Transaction Tyepe")
''mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
''mTxtLeft = mPrintPoint '- (mStrLength) '* 5) '270
''mFontSize = 7
''
''Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop, "Transaction Type")
''Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop, CmbExchangeTypeID)
''
''mTxtTop = mTxtTop + mRowHeight
''
''Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop, "Name Of Agent")
''Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop, TxtAgent)
''
''objPDF.PDFDrawLine mGlobalLeft + 132, mTop - 4 * mRowHeight, mGlobalLeft + 132, mHeight + mRowHeight
''Call DrawTable(2, , mTop, mHeight, mLeft, mRowHeight)
''
''
'''Row 3

mRowHeight = 15
mGlobalLeft = 12
'objPDF.PDFDrawPolygon Array(12, 15, objPDF.PDFGetPageWidth - 18, 15, objPDF.PDFGetPageWidth - 18, 30, 12, 30)
mLeft = mGlobalLeft
mTop = 15
mHeight = mTop + mRowHeight
mTxtFromTop = mRowHeight + 2
 
 
 
 'Call DrawTable(10, 2, 15, 30, 12)

Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)

mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
mStrLength = Len("NCMSL " & TxtCMPID & "")
mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
mTxtLeft = mPrintPoint - (mStrLength * 2) '* 5) '270
mFontSize = 12

Call PrintText(mFontSize, FONT_NORMAL, ALIGN_CENTER, mTxtLeft, mTxtTop, "NCMSL " & TxtCMPID.Text & "")

If Totpages > 1 Then
   Call PrintText(6, FONT_NORMAL, ALIGN_RIGHT, objPDF.PDFGetPageWidth - 70, mTxtTop - 5, "Page " & objPDF.PDFPageNumber & " out of " & Totpages)
Else
   
End If
mHeight = mTop + mRowHeight
mLeft = mGlobalLeft


Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)

mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
mStrLength = Len("Commodity Deposit Transaction Form (CDTF)")
mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
mTxtLeft = mPrintPoint - (mStrLength * 2) '* 5) '270
mFontSize = 10

Call PrintText(mFontSize, FONT_NORMAL, ALIGN_CENTER, mTxtLeft, mTxtTop, "Commodity Issue Transaction Form (CITF)")

mHeight = mTop + mRowHeight
mLeft = mGlobalLeft

Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
 
tmp1 = mTop
mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
mStrLength = Len("(This is NOT an instrument of title of the goods)")
mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
mTxtLeft = mPrintPoint '- (mStrLength) '* 5) '270
mFontSize = 6

Call PrintText(mFontSize, FONT_NORMAL, ALIGN_CENTER, mTxtLeft - 30, mTxtTop - 3, "(This is NOT an instrument of title of the goods)")

'--Code added By Kishor
'Row 1

mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
mStrLength = Len("Depositor Details")
mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
mTxtLeft = mPrintPoint '- (mStrLength) '* 5) '270
mFontSize = 7
Dim mDispTextSize  As Integer
mDispTextSize = 6

mTxtTop = mTxtTop + mRowHeight

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop, "CITF No")
Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop, TxtCxTFNo)
'--
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 350, mTxtTop, "Date")
Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 450, mTxtTop, Format(TxtCxTFDate, "dd-MMM-yyyy"))

mTxtTop = mTxtTop + mRowHeight
objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop, "Location")
Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop, GetLocationFromCMP(mCMPID))
'--
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 350, mTxtTop, "State")
Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 450, mTxtTop, GetStateFromLocation(GetLocationIDFromCMP(mCMPID)))
'--
mTxtTop = mTxtTop + mRowHeight
objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop, "CMPID & Address")
Dim mCMPIDnAddr As String
mCMPIDnAddr = mCMPID & "-" & GetCMPAddress(mCMPID, "A")
Dim mLen As Integer, mMax As Integer, mPara As Integer
Dim t1 As Double
mMax = 45
mPara = 0
t1 = 1
mLen = Len(mCMPIDnAddr)
mLen = Round(mLen / mMax) '+ 1

For T = 1 To mLen
    If T > 1 Then
       t1 = mMax + t1
       'mMax '= mMax + 35
       mPara = mPara + 7
    End If
    Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop + mPara, mID$(mCMPIDnAddr, t1, mMax))
Next

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 350, mTxtTop, "Commodity")
Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 450, mTxtTop, TxtCommodityID)

mTxtTop = mTxtTop + mPara + mRowHeight

objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10

objPDF.PDFDrawLine mGlobalLeft + 345, Val(tmp1), mGlobalLeft + 345, mTxtTop - 10
objPDF.PDFDrawLine mGlobalLeft + 445, Val(tmp1), mGlobalLeft + 445, mTxtTop - 10

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop, "Transaction Type")
Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop, CmbExchangeTypeID)

mTxtTop = mTxtTop + mRowHeight
objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop, "Name Of Agent")
Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop, TxtAgent)

mTxtTop = mTxtTop + mRowHeight
objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10
objPDF.PDFDrawLine mGlobalLeft + 135, Val(tmp1), mGlobalLeft + 135, mTxtTop - 10

mTxtTop = mTxtTop + mRowHeight '- (mTxtFromTop) 'mTop ' + 15
mStrLength = Len("Withdrawer Details")
mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
mTxtLeft = mPrintPoint '- (mStrLength) '* 5) '270
mFontSize = 8
mHeight = mTxtTop
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mLeft, mTxtTop - 2, "Withdrawer Details")

Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)

objPDF.PDFDrawLine mGlobalLeft + 72, mTop, mGlobalLeft + 72, mTop + ((MSHFlexGrid2.Rows - 2) * mRowHeight)
objPDF.PDFDrawLine mGlobalLeft + 142, mTop, mGlobalLeft + 142, mTop + ((MSHFlexGrid2.Rows - 2) * mRowHeight)
If mExchangeTypeID <> 2 Then
   objPDF.PDFDrawLine 2.8 * (mGlobalLeft + 142), mTop, 2.8 * (mGlobalLeft + 142), mTop + ((MSHFlexGrid2.Rows - 1) * mRowHeight)
End If
objPDF.PDFDrawLine 3.3 * (mGlobalLeft + 142), mTop, 3.3 * (mGlobalLeft + 142), mTop + ((MSHFlexGrid2.Rows - 1) * mRowHeight)

Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)


mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
mStrLength = Len("Depositor Details")
mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
mTxtLeft = mPrintPoint '- (mStrLength) '* 5) '270
mFontSize = 7

If CmbExchangeTypeID.Text = "NCDEX" Then
   If Check1.Value = 0 Then
      Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mLeft, mTxtTop, "RRN No")
   Else
      Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mLeft, mTxtTop, "Rejected Withdrawal")
   End If
ElseIf CmbExchangeTypeID.Text = "Non NCDEX" Then
   
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mLeft, mTxtTop, "WHR No")
ElseIf CmbExchangeTypeID.Text = "SPOT" Then
   
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mLeft, mTxtTop, "CDC No")
End If

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mLeft + 80, mTxtTop, "Client ID")
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 150, mTxtTop, "Client Name")
If mExchangeTypeID <> 2 Then
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 430, mTxtTop, "No of Bags")
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 500, mTxtTop - 6, "Electronic Weight")
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 500, mTxtTop, "in MT")
 
Else
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 500, mTxtTop - 6, "Electronic Weight")
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 500, mTxtTop, "in MT")
 
End If

objPDF.PDFDrawLine mGlobalLeft + 72, mTop, mGlobalLeft + 72, mTop + ((MSHFlexGrid2.Rows - 2) * mRowHeight)
objPDF.PDFDrawLine mGlobalLeft + 142, mTop, mGlobalLeft + 142, mTop + ((MSHFlexGrid2.Rows - 2) * mRowHeight)
If mExchangeTypeID <> 2 Then
   objPDF.PDFDrawLine 2.8 * (mGlobalLeft + 142), mTop, 2.8 * (mGlobalLeft + 142), mTop + ((MSHFlexGrid2.Rows - 1) * mRowHeight)
End If
objPDF.PDFDrawLine 3.3 * (mGlobalLeft + 142), mTop, 3.3 * (mGlobalLeft + 142), mTop + ((MSHFlexGrid2.Rows - 1) * mRowHeight)
objPDF.PDFDrawLine mGlobalLeft, mTop + mRowHeight, mGlobalLeft + 560, mTop + mRowHeight
Call DrawTable(1, , mTop, mHeight, mLeft, mRowHeight)
 
'FILL GRID WITH DETAILS
Dim mTemp As Integer
For T = 2 To MSHFlexGrid2.Rows - 1
    If mHeight > objPDF.PDFGetPageHeight Then
       Call CheckForNewPage
       mTop = 15
       mTxtTop = 10
    End If
    If T = 1 Then
       mTemp = mFontSize
    Else
       mTemp = mDispTextSize
    End If
    mTxtTop = mTxtTop + mRowHeight
   
    Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mLeft, mTxtTop, MSHFlexGrid2.TextMatrix(T, 1))
    Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 80, mTxtTop, MSHFlexGrid2.TextMatrix(T, 3))
    Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 150, mTxtTop, MSHFlexGrid2.TextMatrix(T, 4))
    If mExchangeTypeID <> 2 Then
       Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 430, mTxtTop, Format(MSHFlexGrid2.TextMatrix(T, 15), "#######0.000"))
    End If
    Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 500, mTxtTop, Format(MSHFlexGrid2.TextMatrix(T, 9), "#######0.000"))
    
    objPDF.PDFDrawLine mGlobalLeft + 72, mTop, mGlobalLeft + 72, mTop + mRowHeight
    objPDF.PDFDrawLine mGlobalLeft + 142, mTop, mGlobalLeft + 142, mTop + mRowHeight
    If mExchangeTypeID <> 2 Then
       objPDF.PDFDrawLine 2.8 * (mGlobalLeft + 142), mTop, 2.8 * (mGlobalLeft + 142), mTop + mRowHeight
    End If
    objPDF.PDFDrawLine 3.3 * (mGlobalLeft + 142), mTop, 3.3 * (mGlobalLeft + 142), mTop + mRowHeight

    'Call DrawTable(MSHFlexGrid2.Rows - 1, , mTop, mHeight, mLeft, mRowHeight)
    objPDF.PDFDrawLine mGlobalLeft, mTop + mRowHeight, mGlobalLeft + 565, mTop + mRowHeight

    mTop = mTop + mRowHeight
    mHeight = mHeight + mRowHeight
Next

Call CheckForNewPage
mTxtTop = mTxtTop + mRowHeight
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 220, mTxtTop, "Total")
If mExchangeTypeID <> 2 Then
   Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 430, mTxtTop, Format(MSHFlexGrid2.TextMatrix(0, 15), "#######0.000"))
End If
Call CheckForNewPage
Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 500, mTxtTop, Format(MSHFlexGrid2.TextMatrix(0, 9), "#######0.000"))
Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)

'mTxtTop = mTxtTop + mRowHeight
mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
mStrLength = Len("Truck-wise Withdrawal Details")
mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
mTxtLeft = mPrintPoint '- (mStrLength) '* 5) '270
mFontSize = 7
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mLeft, mTxtTop - 2, "Truck-wise Withdrawal Details")

''objPDF.PDFDrawLine mGlobalLeft, mTxtTop + mRowHeight, mGlobalLeft + 565, mTxtTop + mRowHeight
''objPDF.PDFDrawLine mGlobalLeft + 50, mTop, mGlobalLeft + 50, mTop + ((MSHFlexGrid3.Rows - 1) * mRowHeight)
''objPDF.PDFDrawLine (mGlobalLeft + 120), mTop, (mGlobalLeft + 120), mTop + ((MSHFlexGrid3.Rows - 1) * mRowHeight)
''objPDF.PDFDrawLine (mGlobalLeft + 190), mTop, (mGlobalLeft + 190), mTop + ((MSHFlexGrid3.Rows + 6) * mRowHeight)
''objPDF.PDFDrawLine (mGlobalLeft + 256), mTop, (mGlobalLeft + 256), mTop + ((MSHFlexGrid3.Rows + 6) * mRowHeight)
''objPDF.PDFDrawLine (mGlobalLeft + 356), mTop, (mGlobalLeft + 356), mTop + ((MSHFlexGrid3.Rows - 1) * mRowHeight)
''
''Dim mTempTop As Integer
''mTempTop = mTxtTop - 3
''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 90, mTempTop + (MSHFlexGrid3.Rows * mRowHeight), "Total")
''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 60, mTempTop + ((MSHFlexGrid3.Rows + 1) * mRowHeight), "Less Gunny Weight(MT)")
''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 50, mTempTop + ((MSHFlexGrid3.Rows + 2) * mRowHeight), "Less Standard Deductions(MT)")
''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 60, mTempTop + ((MSHFlexGrid3.Rows + 3) * mRowHeight), "Add Adjustments(MT)")
''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 60, mTempTop + ((MSHFlexGrid3.Rows + 4) * mRowHeight), "Electronic Weight(MT)")
'''pp 'Hologram No
''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 270, mTempTop + ((MSHFlexGrid3.Rows + 4) * mRowHeight), "Hologram No")
''
'''Actual Values
''Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 200, mTempTop + (MSHFlexGrid3.Rows * mRowHeight), MSHFlexGrid3.TextMatrix(0, 5))
''Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 200, mTempTop + ((MSHFlexGrid3.Rows + 1) * mRowHeight), MSHFlexGrid4.TextMatrix(0, 8))
''Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 200, mTempTop + ((MSHFlexGrid3.Rows + 2) * mRowHeight), MSHFlexGrid4.TextMatrix(0, 9))
''Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 200, mTempTop + ((MSHFlexGrid3.Rows + 3) * mRowHeight), MSHFlexGrid4.TextMatrix(0, 6))
''Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 200, mTempTop + ((MSHFlexGrid3.Rows + 4) * mRowHeight), MSHFlexGrid4.TextMatrix(0, 10))
'''Hologram No -- Use Record Set
''
''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 370, mTempTop + ((MSHFlexGrid3.Rows + 4) * mRowHeight), " " & RsTxn_CxTF_Details!HologramNo & " ")
''
''Dim mTemp As Integer
''For t = 1 To MSHFlexGrid3.Rows - 1
''    If t = 1 Then
''       mTemp = mFontSize
''    Else
''       mTemp = mDispTextSize
''    End If
''    mTxtTop = mTxtTop + mRowHeight
''    Call PrintText(mTemp, FONT_NORMAL, ALIGN_LEFT, mLeft, mTxtTop, MSHFlexGrid3.TextMatrix(t, 1))
''    Call PrintText(mTemp, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 60, mTxtTop, MSHFlexGrid3.TextMatrix(t, 3))
''    Call PrintText(mTemp, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 130, mTxtTop, MSHFlexGrid3.TextMatrix(t, 4))
''    Call PrintText(mTemp, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 200, mTxtTop, MSHFlexGrid3.TextMatrix(t, 5))
''    Call PrintText(mTemp, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 256, mTxtTop, MSHFlexGrid3.TextMatrix(t, 2))
''    Call PrintText(mTemp, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 356, mTxtTop, MSHFlexGrid3.TextMatrix(t, 6))
''Next
''
''Call DrawTable(MSHFlexGrid3.Rows + 6, , mTop, mHeight, mLeft, mRowHeight)
'''Clear Line
''mTxtTop = mHeight - (mTxtFromTop)
''objPDF.PDFSetTextColor = vbWhite
''
''Call PrintText(15, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 255.2, mTxtTop - (6 * mRowHeight), "_____________________________________")
''Call PrintText(15, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 255.2, mTxtTop - (5 * mRowHeight), "_____________________________________")
''Call PrintText(15, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 255.2, mTxtTop - (4 * mRowHeight), "_____________________________________")
''
''objPDF.PDFSetTextColor = vbBlack
''Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 256, mTxtTop - (3 * mRowHeight), "Affix Hologram Here")
'''--
''
''objPDF.PDFDrawLine (mGlobalLeft + 50), mTop - (2 * mRowHeight), (mGlobalLeft + 50), mTop
''objPDF.PDFDrawLine (mGlobalLeft + 120), mTop - (2 * mRowHeight), (mGlobalLeft + 120), mTop
''
''objPDF.PDFDrawLine (mGlobalLeft + 360), mTop - (3 * mRowHeight), (mGlobalLeft + 360), mTop
''objPDF.PDFDrawLine (mGlobalLeft + 420), mTop - (2 * mRowHeight), (mGlobalLeft + 420), mTop
''
''If mHeight > objPDF.PDFGetPageHeight - 50 Then
''   objPDF.PDFEndPage
''   objPDF.PDFNewPage
''   objPDF.PDFSetFont FONT_ARIAL, 12, FONT_BOLD
''   objPDF.PDFSetDrawColor = vbWhite
''   objPDF.PDFSetTextColor = vbBlack
''   objPDF.PDFSetAlignement = ALIGN_CENTER
''   objPDF.PDFSetBorder = BORDER_ALL
''   objPDF.PDFSetFill = True
''
''   objPDF.PDFCell " ", 15, 15, _
''   objPDF.PDFGetPageWidth - 30, objPDF.PDFGetPageHeight - 20
''   mLeft = mGlobalLeft
''   mTop = 15
''   mHeight = mTop + mRowHeight
''   mTxtFromTop = mRowHeight + 2
''End If
''
''Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
''mTxtTop = mHeight - (mTxtFromTop)
''
''Call PrintText(5.4, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "NOTE: THIS DOCUMENT IS NOT VALID WITHOUT A HOLOGRAM OF NCMSL AFFIXED ABOVE & THE SERIAL NUMBER IN THE HOLOGRAM SHOULD MATCH WITH THE HOLOGRAM NUMBER PRINTED ABOVE.")
''
''mHeight = mHeight + 50
''
''Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
''mTxtTop = mHeight - (mTxtFromTop) '+ 15
''
''Call PrintText(9, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "Name & Signature of CMP Manager/CMP DRC:")
''
''mHeight = mHeight + 30
''Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
''mTxtTop = mHeight - (mTxtFromTop)
''
''Call PrintText(6.5, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop - 17, "I/We hereby declare that the commodities being deposited herein belong to the clients as stated above and are free from any encumbrances. I/We request  you to take custody of the said")
''Call PrintText(6.5, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop - 10, "commodities and I/We undertake to indemnify you against any losses incurred by you due to any false information provided by me/us. I/We also undertake to pay all required charges as per")
''Call PrintText(6.5, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop - 3, "your charges schedule from time to time. I have gone throught the attached 'Terms and Conditions'** of storage and agree to abide by the same.")
''
'''Call CheckForNewPage
'''objPDF.PDFDrawLine mGlobalLeft + 50, mTop, mGlobalLeft + 50, mTop + ((MSHFlexGrid3.Rows - 1) * mRowHeight)
'''objPDF.PDFDrawLine (mGlobalLeft + 120), mTop, (mGlobalLeft + 120), mTop + ((MSHFlexGrid3.Rows - 1) * mRowHeight)
'''objPDF.PDFDrawLine (mGlobalLeft + 190), mTop, (mGlobalLeft + 190), mTop + ((MSHFlexGrid3.Rows) * mRowHeight)
'''objPDF.PDFDrawLine (mGlobalLeft + 256), mTop, (mGlobalLeft + 256), mTop + ((MSHFlexGrid3.Rows + 6) * mRowHeight)
'''objPDF.PDFDrawLine (mGlobalLeft + 356), mTop, (mGlobalLeft + 356), mTop + ((MSHFlexGrid3.Rows - 1) * mRowHeight)
'''
'''
''''Dim mTempTop As Integer
''''mTempTop = mTxtTop - 3
''''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 90, mTempTop + (MSHFlexGrid3.Rows * mRowHeight), "Total")
'''
''''----Commented on 18-01-2011 by Sagar
'''
''''''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 60, mTempTop + ((MSHFlexGrid3.Rows + 1) * mRowHeight), "Less Gunny Weight(MT)")
''''''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 50, mTempTop + ((MSHFlexGrid3.Rows + 2) * mRowHeight), "Less Standard Deductions(MT)")
''''''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 60, mTempTop + ((MSHFlexGrid3.Rows + 3) * mRowHeight), "Add Adjustments(MT)")
''''''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 60, mTempTop + ((MSHFlexGrid3.Rows + 4) * mRowHeight), "Electronic Weight(MT)")
'''
'''objPDF.PDFDrawLine (mGlobalLeft), mTempTop + (MSHFlexGrid3.Rows + 4 * mRowHeight), (mGlobalLeft + 256), mTempTop + (MSHFlexGrid3.Rows + 4 * mRowHeight)
'''
''''pp 'Hologram No
'''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 270, mTempTop + ((MSHFlexGrid3.Rows + 4) * mRowHeight), "Hologram No")
'''
''''Actual Values
'''Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 200, mTempTop + (MSHFlexGrid3.Rows * mRowHeight), MSHFlexGrid3.TextMatrix(0, 5))
'''
''''----Commented on 18-01-2011 by Sagar
'''
''''''Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 200, mTempTop + ((MSHFlexGrid3.Rows + 1) * mRowHeight), MSHFlexGrid4.TextMatrix(0, 8))
''''''Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 200, mTempTop + ((MSHFlexGrid3.Rows + 2) * mRowHeight), MSHFlexGrid4.TextMatrix(0, 9))
''''''Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 200, mTempTop + ((MSHFlexGrid3.Rows + 3) * mRowHeight), MSHFlexGrid4.TextMatrix(0, 6))
''''''Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 200, mTempTop + ((MSHFlexGrid3.Rows + 4) * mRowHeight), MSHFlexGrid4.TextMatrix(0, 10))
''''Hologram No -- Use Record Set
'''
'''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 370, mTempTop + ((MSHFlexGrid3.Rows + 4) * mRowHeight), " " & RsTxn_CxTF_Details!HologramNo & " ")
'''
'mTxtTop = mTxtTop + mRowHeight
mTop = mTop - mRowHeight
mHeight = mHeight - mRowHeight
For T = 1 To MSHFlexGrid3.Rows - 1
    If T = 1 Then
       mTemp = mFontSize
    Else
       mTemp = mDispTextSize
    End If
    
    If mHeight > objPDF.PDFGetPageHeight - 50 Then
       Call CheckForNewPage
       mTop = 15
       mTxtTop = 10
       
       
      'objPDF.PDFDrawPolygon Array(12, 15, objPDF.PDFGetPageWidth - 18, 15, objPDF.PDFGetPageWidth - 18, 30, 12, 30)
      mLeft = mGlobalLeft
      mTop = 15
      mHeight = mTop + mRowHeight
      mTxtFromTop = mRowHeight + 2
      
      'Call DrawTable(10, 2, 15, 30, 12)
      
      Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
      
      
      mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
      mStrLength = Len("NCMSL " & TxtCMPID & "")
      mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
      mTxtLeft = mPrintPoint - (mStrLength * 2) '* 5) '270
      mFontSize = 12
      
      Call PrintText(mFontSize, FONT_NORMAL, ALIGN_CENTER, mTxtLeft, mTxtTop, "NCMSL " & TxtCMPID.Text & "")
      
      mHeight = mTop + mRowHeight
      mLeft = mGlobalLeft
      
      
      Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
      
      mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
      mStrLength = Len("Commodity Deposit Transaction Form (CDTF)")
      mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
      mTxtLeft = mPrintPoint - (mStrLength * 2) '* 5) '270
      mFontSize = 10
      
      Call PrintText(mFontSize, FONT_NORMAL, ALIGN_CENTER, mTxtLeft, mTxtTop, "Commodity Issue Transaction Form (CITF)")
      
      mHeight = mTop + mRowHeight
      mLeft = mGlobalLeft
      
      Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
      
      tmp1 = mTop
      mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
      mStrLength = Len("(This is NOT an instrument of title of the goods)")
      mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
      mTxtLeft = mPrintPoint '- (mStrLength) '* 5) '270
      mFontSize = 6
      
      Call PrintText(mFontSize, FONT_NORMAL, ALIGN_CENTER, mTxtLeft - 30, mTxtTop - 3, "(This is NOT an instrument of title of the goods)")
      
      '--Code added By Kishor
      'Row 1
      
      mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
      mStrLength = Len("Depositor Details")
      mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
      mTxtLeft = mPrintPoint '- (mStrLength) '* 5) '270
      mFontSize = 7
      'Dim mDispTextSize  As Integer
      mDispTextSize = 6
      
      mTxtTop = mTxtTop + mRowHeight
      mTop = mTop + mRowHeight
      Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop, "CITF No")
      Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop, TxtCxTFNo)
      '--
      Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 350, mTxtTop, "Date")
      Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 450, mTxtTop, Format(TxtCxTFDate, "dd-MMM-yyyy"))
      
      mTxtTop = mTxtTop + mRowHeight
      mTop = mTop + mRowHeight
      objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10
      
      Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop, "Location")
      Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop, GetLocationFromCMP(mCMPID))
      '--
      Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 350, mTxtTop, "State")
      Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 450, mTxtTop, GetStateFromLocation(GetLocationIDFromCMP(mCMPID)))
      '--
      mTxtTop = mTxtTop + mRowHeight
      mTop = mTop + mRowHeight
      objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10

     
      Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop, "Transaction Type")
      Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop, CmbExchangeTypeID)
      
      Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 350, mTxtTop, "Commodity")
      Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 450, mTxtTop, TxtCommodityID)
      
      mTxtTop = mTxtTop + mRowHeight
    '  mTop = mTop + mRowHeight
      objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10

      objPDF.PDFDrawLine mGlobalLeft + 345, Val(tmp1), mGlobalLeft + 345, mTxtTop - 10
      objPDF.PDFDrawLine mGlobalLeft + 445, Val(tmp1), mGlobalLeft + 445, mTxtTop - 10
      objPDF.PDFDrawLine mGlobalLeft + 135, Val(tmp1), mGlobalLeft + 135, mTxtTop - 10
      
      mTxtTop = mTxtTop + (mRowHeight * 2)
      mTop = mTop + (mRowHeight * 2) + 3
      
      objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10
      
      Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mLeft, mTxtTop, MSHFlexGrid3.TextMatrix(1, 1))
      Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 60, mTxtTop, MSHFlexGrid3.TextMatrix(1, 3))
      Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 130, mTxtTop, MSHFlexGrid3.TextMatrix(1, 4))
      Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 200, mTxtTop, MSHFlexGrid3.TextMatrix(1, 5))
      Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 256, mTxtTop, MSHFlexGrid3.TextMatrix(1, 2))
      Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 356, mTxtTop, MSHFlexGrid3.TextMatrix(1, 6))
      
      objPDF.PDFDrawLine mGlobalLeft + 50, mTop, mGlobalLeft + 50, mTop + mRowHeight
      objPDF.PDFDrawLine (mGlobalLeft + 120), mTop, (mGlobalLeft + 120), mTop + mRowHeight
      objPDF.PDFDrawLine (mGlobalLeft + 190), mTop, (mGlobalLeft + 190), mTop + mRowHeight
      objPDF.PDFDrawLine (mGlobalLeft + 256), mTop, (mGlobalLeft + 256), mTop + mRowHeight
      objPDF.PDFDrawLine (mGlobalLeft + 356), mTop, (mGlobalLeft + 356), mTop + mRowHeight
      objPDF.PDFDrawLine mGlobalLeft, mTop + mRowHeight, mGlobalLeft + 565, mTop + mRowHeight
   
     ' mTop = mTop + mRowHeight
      'mTxtTop = mTxtTop + mRowHeight
      
      mHeight = mTxtTop
      'mTxtTop = mTxtTop + mRowHeight
      'mTop = mTxtTop
    End If
 
        
    mTxtTop = mTxtTop + mRowHeight
    mTop = mTop + mRowHeight
    mHeight = mHeight + mRowHeight
    
    Call PrintText(mTemp, FONT_NORMAL, ALIGN_LEFT, mLeft, mTxtTop, MSHFlexGrid3.TextMatrix(T, 1))
    Call PrintText(mTemp, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 60, mTxtTop, MSHFlexGrid3.TextMatrix(T, 3))
    Call PrintText(mTemp, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 130, mTxtTop, MSHFlexGrid3.TextMatrix(T, 4))
    Call PrintText(mTemp, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 200, mTxtTop, MSHFlexGrid3.TextMatrix(T, 5))
    Call PrintText(mTemp, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 256, mTxtTop, MSHFlexGrid3.TextMatrix(T, 2))
    Call PrintText(mTemp, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 356, mTxtTop, MSHFlexGrid3.TextMatrix(T, 6))
    
    objPDF.PDFDrawLine mGlobalLeft + 50, mTop, mGlobalLeft + 50, mTop + mRowHeight
    objPDF.PDFDrawLine (mGlobalLeft + 120), mTop, (mGlobalLeft + 120), mTop + mRowHeight
    objPDF.PDFDrawLine (mGlobalLeft + 190), mTop, (mGlobalLeft + 190), mTop + mRowHeight
    objPDF.PDFDrawLine (mGlobalLeft + 256), mTop, (mGlobalLeft + 256), mTop + mRowHeight
    objPDF.PDFDrawLine (mGlobalLeft + 356), mTop, (mGlobalLeft + 356), mTop + mRowHeight
    objPDF.PDFDrawLine mGlobalLeft, mTop + mRowHeight, mGlobalLeft + 565, mTop + mRowHeight


Next
'Call CheckForNewPage
'Call DrawTable(MSHFlexGrid3.Rows - 1, , mTop, mHeight, mLeft, mRowHeight)
Call CheckForNewPage
Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 90, mTop + mRowHeight - 3, "Total")
Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 200, mTop + mRowHeight - 3, MSHFlexGrid3.TextMatrix(0, 5))
objPDF.PDFDrawLine (mGlobalLeft + 190), mTop, (mGlobalLeft + 190), mTop + mRowHeight
objPDF.PDFDrawLine (mGlobalLeft + 256), mTop, (mGlobalLeft + 256), mTop + mRowHeight * 4

Call CheckForNewPage
objPDF.PDFDrawLine (mGlobalLeft), mTop + mRowHeight, (mGlobalLeft + 256), mTop + mRowHeight

mTop = mTop + (mRowHeight * 4)
objPDF.PDFDrawLine mGlobalLeft + 50, mTop + mRowHeight, mGlobalLeft + 50, mTop + (mRowHeight * 3)
objPDF.PDFDrawLine (mGlobalLeft + 120), mTop + mRowHeight, (mGlobalLeft + 120), mTop + (mRowHeight * 3)
objPDF.PDFDrawLine (mGlobalLeft + 190), mTop + mRowHeight, (mGlobalLeft + 190), mTop + (mRowHeight * 3)
objPDF.PDFDrawLine (mGlobalLeft + 256), mTop, (mGlobalLeft + 256), mTop + (mRowHeight * 3)
objPDF.PDFDrawLine (mGlobalLeft + 356), mTop, (mGlobalLeft + 356), mTop + (mRowHeight * 3)
objPDF.PDFDrawLine (mGlobalLeft + 420), mTop + mRowHeight, (mGlobalLeft + 420), mTop + (mRowHeight * 3)
Call CheckForNewPage
Call PrintText(6, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 265, mTop - 3, "Affix Hologram Here")

If mExchangeTypeID = 2 Then
   Call CheckForNewPage
   Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 70, mTop + mRowHeight - 3, "Total No. of Bags")
   Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 200, mTop + mRowHeight - 3, Format(MSHFlexGrid4.TextMatrix(0, 6)))
   objPDF.PDFDrawLine (mGlobalLeft + 190), mTop, (mGlobalLeft + 190), mTop + mRowHeight
End If

Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 270, mTop + mRowHeight - 3, "Hologram No")
Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 370, mTop + mRowHeight - 3, " " & RsTxn_CxTF_Details!HologramNo & " ")
Call CheckForNewPage
Call DrawTable(4, , mTop, mTop, mLeft, mRowHeight)

'If mHeight > objPDF.PDFGetPageHeight - 50 Then
'   objPDF.PDFEndPage
'   objPDF.PDFNewPage
'   objPDF.PDFSetFont FONT_ARIAL, 12, FONT_BOLD
'   objPDF.PDFSetDrawColor = vbWhite
'   objPDF.PDFSetTextColor = vbBlack
'   objPDF.PDFSetAlignement = ALIGN_CENTER
'   objPDF.PDFSetBorder = BORDER_ALL
'   objPDF.PDFSetFill = True
'
'   objPDF.PDFCell " ", 15, 15, _
'   objPDF.PDFGetPageWidth - 30, objPDF.PDFGetPageHeight - 20
'   mLeft = mGlobalLeft
'   mTop = 15
'   mHeight = mTop + mRowHeight
'   mTxtFromTop = mRowHeight + 2
'End If

mHeight = mTop
Call CheckForNewPage
Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
Call CheckForNewPage
mTxtTop = mHeight - (mTxtFromTop)
'''

Call PrintText(5.4, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "NOTE: THIS DOCUMENT IS NOT VALID WITHOUT A HOLOGRAM OF NCMSL AFFIXED ABOVE & THE SERIAL NUMBER IN THE HOLOGRAM SHOULD MATCH WITH THE HOLOGRAM NUMBER PRINTED ABOVE.")
'''
mHeight = mHeight + 50
'''
Call CheckForNewPage
Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)

'''
Call CheckForNewPage
mTxtTop = mHeight - (mTxtFromTop) '+ 15
Call PrintText(9, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "Name & Signature of CMP Manager/CMP DRC:")
Call CheckForNewPage
Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
Call CheckForNewPage
mTxtTop = mTxtTop + mRowHeight
Call PrintText(7, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "Received the above mentioned goods in good condition.")
'mHeight = mHeight + 30
'Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)

'''
''Call PrintText(6.5, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop - 17, "I/We hereby declare that the commodities being deposited herein belong to the clients as stated above and are free from any encumbrances. I/We request  you to take custody of the said")
''Call PrintText(6.5, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop - 10, "commodities and I/We undertake to indemnify you against any losses incurred by you due to any false information provided by me/us. I/We also undertake to pay all required charges as per")
''Call PrintText(6.5, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop - 3, "your charges schedule from time to time. I have gone throught the attached 'Terms and Conditions'** of storage and agree to abide by the same.")
'''''

mHeight = mHeight + 70

Call CheckForNewPage

Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
Call CheckForNewPage
mTxtTop = mHeight - (mTxtFromTop)
Call PrintText(9, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "Name & Signature of Withdrawer/Agent of Withdrawer:")

mHeight = mHeight + 50
'Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)

Call CheckForNewPage
mTxtTop = mHeight - (mTxtFromTop)
Call PrintText(8, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "**'Terms and Conditions' can be viewed in http://www.ncmsl.com/whrtNc.html")


If mExchangeTypeID = 1 Then
   Call PrintText(7, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 510, mTxtTop + (20 * 11), "F&F 118")
ElseIf mExchangeTypeID = 2 Then
   Call PrintText(7, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 510, mTxtTop + (20 * 11), "F&F 117")
End If

objPDF.PDFEndDoc

End Sub




Private Sub CheckForNewPage()

If mHeight > objPDF.PDFGetPageHeight - 50 Then
   objPDF.PDFWatermark = ""
   objPDF.PDFEndPage
   objPDF.PDFNewPage
   objPDF.PDFSetFont FONT_ARIAL, 12, FONT_BOLD
   objPDF.PDFSetDrawColor = vbWhite
   objPDF.PDFSetTextColor = vbBlack
   objPDF.PDFSetAlignement = ALIGN_CENTER
   objPDF.PDFSetBorder = BORDER_ALL
   objPDF.PDFSetFill = True

   objPDF.PDFCell " ", 15, 15, _
   objPDF.PDFGetPageWidth - 30, objPDF.PDFGetPageHeight - 20
   If PntPrv = "Y" Then
      objPDF.PDFWatermark = "Preview"
   Else
      objPDF.PDFWatermark = "NCMSL"
   End If
   
    objPDF.PDFSetFont FONT_ARIAL, 12, FONT_BOLD
    objPDF.PDFSetDrawColor = vbWhite
    objPDF.PDFSetTextColor = vbBlack
    objPDF.PDFSetAlignement = ALIGN_CENTER
    objPDF.PDFSetBorder = BORDER_ALL
    objPDF.PDFSetFill = True

   mLeft = mGlobalLeft
   mTop = 30
   mHeight = mTop + mRowHeight
   mTxtFromTop = mRowHeight + 2
   mTxtTop = mHeight - (mTxtFromTop)
   If Totpages > 1 Then
      Call PrintText(6, FONT_NORMAL, ALIGN_RIGHT, objPDF.PDFGetPageWidth - 70, mTxtTop - 5, "Page " & objPDF.PDFPageNumber & " out of " & Totpages)
   Else
   End If
   
End If

End Sub



'Private Sub CheckForNewPage()
'
'If mHeight > objPDF.PDFGetPageHeight Then
'   objPDF.PDFWatermark = ""
'   objPDF.PDFEndPage
'   objPDF.PDFNewPage
'   objPDF.PDFSetFont FONT_ARIAL, 12, FONT_BOLD
'   objPDF.PDFSetDrawColor = vbWhite
'   objPDF.PDFSetTextColor = vbBlack
'   objPDF.PDFSetAlignement = ALIGN_CENTER
'   objPDF.PDFSetBorder = BORDER_ALL
'   objPDF.PDFSetFill = True
'
'   objPDF.PDFCell " ", 15, 15, _
'   objPDF.PDFGetPageWidth - 30, objPDF.PDFGetPageHeight - 20
'   If PntPrv = "Y" Then
'      objPDF.PDFWatermark = "Preview"
'   Else
'      objPDF.PDFWatermark = "NCMSL"
'   End If
'
'    objPDF.PDFSetFont FONT_ARIAL, 12, FONT_BOLD
'    objPDF.PDFSetDrawColor = vbWhite
'    objPDF.PDFSetTextColor = vbBlack
'    objPDF.PDFSetAlignement = ALIGN_CENTER
'    objPDF.PDFSetBorder = BORDER_ALL
'    objPDF.PDFSetFill = True
'
'   mLeft = mGlobalLeft
'   mTop = 30
'   mHeight = mTop + mRowHeight
'   mTxtFromTop = mRowHeight + 2
'End If
'
'End Sub

'Private Function CalculateDepoNoOfCycle(mFromDate_ As Date, mToDate_ As Date, mBillingCycleID_ As Variant, mNoOfDays_ As Variant)
'Dim mCycleCount As Variant
'Dim MonthCount, DepositMonthCycles, WithdrawalMonthCycles As Variant
''D=1 Daily
'If mBillingCycleID_ = 1 Then
'   mCycleCount = DateDiff("d", mFromDate_, mToDate_) + 1
''w=2 'Weekly
'ElseIf mBillingCycleID_ = 2 Then
'   MonthCount = DateDiff("M", mFromDate_, mToDate_) - 1
''   DepositMonthCycles = 4 - Left((DatePart("d", mFromDate_) / mNoOfDays_), 2)
'
'   If Format(DatePart("d", mFromDate_) / mNoOfDays_) = 1 Then
'      DepositMonthCycles = 4 '- 0 'Left(Format(DatePart("d", mFromDate_) / mNoOfDays_), 1)
'   ElseIf Format(DatePart("d", mFromDate_) / mNoOfDays_) > 4 Then
'      DepositMonthCycles = 4 - (Left(Format(DatePart("d", mFromDate_) / mNoOfDays_), 1) - 1)
'   Else
'
'      If DatePart("d", mFromDate_) Mod mNoOfDays_ = 0 Then
'         DepositMonthCycles = 4 - Left(Format(DatePart("d", mFromDate_) / mNoOfDays_), 1) + 1
'      Else
'         DepositMonthCycles = 4 - Left(Format(DatePart("d", mFromDate_) / mNoOfDays_), 1)
'      End If
''      DepositMonthCycles = 4 - Left(Format(DatePart("d", mFromDate_) / mNoOfDays_), 1)
'   End If
'
''   WithdrawalMonthCycles = RoundUp(DatePart("d", mToDate_) / mNoOfDays_)
'   mCycleCount = DepositMonthCycles '+ WithdrawalMonthCycles
''F=3 'Fortnightly
'ElseIf mBillingCycleID_ = 3 Then
'   'MonthCount = DateDiff("M", mFromDate_, mToDate_) - 1
'   DepositMonthCycles = IIf(DatePart("d", mFromDate_) - mNoOfDays_ >= 0, 1, 2)
''   WithdrawalMonthCycles = IIf(DatePart("d", mToDate_) - mNoOfDays_ >= 0, 2, 1)
'   mCycleCount = DepositMonthCycles ' + WithdrawalMonthCycles
''M=4 'Monthly
'ElseIf mBillingCycleID_ = 4 Then
'   mCycleCount = DateDiff("M", mFromDate_, mToDate_) + 1
'End If
'
'CalculateDepoNoOfCycle = mCycleCount
'
'End Function

Private Function GetLastDateOfMonth(mDate_ As Date)
Dim mSMonth, mSYear, mSTotalDays As Variant
mSMonth = Month(mDate_)
mSYear = Year(mDate_)
mSTotalDays = Day(DateSerial(mSYear, mSMonth + 1, 0))
GetLastDateOfMonth = CDate(mSYear & "-" & mSMonth & "-" & mSTotalDays)
End Function

Private Sub FillBillingDetails(Row_ As Variant, mFromDate As Date, mToDate As Date)
Dim mMonth, mYear As Variant

If Val(MSHFlexGrid6.TextMatrix(Row_, 11)) <> 0 Then
   MSHFlexGrid6.TextMatrix(Row_, 7) = MSHFlexGrid6.TextMatrix(Row_, 7) & "-Reservation"
   Exit Sub
End If
', MBC.BillingCycle 'WBillingCycle', TCG.BillingUnit, TCG.Rate, '' 'StorageAmount' , MBC.NoofDays,MBC1.BillingCycle 'DepBillingCycle' ,MBC1.NoOfDays 'DepoNoOfDays' ,TDCD.DepositFinalPrice "

   MSHFlexGrid6.TextMatrix(Row_, 12) = IIf(IsNull(TmpRs!WBillingCycle), "", TmpRs!WBillingCycle) 'TmpRs2!BillingCycle
   MSHFlexGrid6.TextMatrix(Row_, 13) = IIf(TmpRs!BillingUnit = "B", "Bag", "Qty") 'IIf(TmpRs2!BillingUnit = "B", "Bag", "Qty")
   MSHFlexGrid6.TextMatrix(Row_, 15) = IIf(IsNull(TmpRs!Rate), 0, TmpRs!Rate) 'TmpRs1!RatePErBag
   MSHFlexGrid6.TextMatrix(Row_, 19) = IIf(IsNull(TmpRs!DepositFinalPrice), 0, TmpRs!DepositFinalPrice) 'TmpRs1!RatePErBag
   
   If mChkSameMonth = True Then
      MSHFlexGrid6.TextMatrix(Row_, 12) = IIf(IsNull(TmpRs!DepBillingCycle), "", TmpRs!DepBillingCycle) ' TmpRs2!DepoBillingCycle
   End If
    
   mLastDateofCDTF = GetLastDateOfMonth(mFromDate)
   mMonth = Month(mFromDate)
   mYear = Year(mFromDate)
            
            
   If mMonth = "12" Then
      mMonth = "01"
      mYear = mYear + 1
   Else
      mMonth = mMonth + 1
   End If
   
     
   mPostDepoFromDate = CDate(mMonth & "-" & "01" & "-" & mYear)
   If mChkSameMonth = False Then
      MSHFlexGrid6.TextMatrix(Row_, 14) = CalculateNoOfCycle(mFromDate, mToDate, TmpRs!BillingCycleID, TmpRs!NoOfDays) 'CalculateNoOfCycle(mFromDate, mToDate, TmpRs2!BillingCycleID, TmpRs2!NoOfDays)
   Else
      MSHFlexGrid6.TextMatrix(Row_, 14) = CalculateDepoNoOfCycle(mFromDate, mLastDateofCDTF, TmpRs!DepoBillingCycleID, TmpRs!DepoNoofDays) 'CalculateDepoNoOfCycle(mFromDate, mLastDateofCDTF, TmpRs2!DepoBillingCycleID, TmpRs2!DepoNoofDays)
   End If

End Sub

Private Sub TxtTTDNoOfBags_LostFocus()
If TxtTTDNoOfBags = "" Then Exit Sub
'If Not IsNumeric(TxtTTDNoOfBags) Then
'   MsgBox "Invalid numeric format"
'   TxtTTDNoOfBags.SetFocus
'   Exit Sub
'End If
'
'If Val(TxtTTDNoOfBags) < 0 Then
'   MsgBox "Invalid numeric format"
'   TxtTTDNoOfBags.SetFocus
'   Exit Sub
'End If

If IsNumeric(TxtTTDNoOfBags) = False Or Val(TxtTTDNoOfBags) < 0 Then
   MsgBox "Invalid numeric format"
   TxtTTDNoOfBags.SetFocus
   Exit Sub
End If

TxtTTDNoOfBags = Format(TxtTTDNoOfBags, "#######0")

End Sub

Private Function CheckClientOutStanding() As Boolean
Dim Retval As Boolean
'Dim mTotOutstAmt As Variant
'Dim mTotCommPrice, mTotBalQty, mtimes As Variant
Retval = True
mTotOutstAmt = 0

''mTotOutstAmt = GetClientOutstanding()
''Call FillGSLBalanceStockWithPrice

If LCase(Gen_UserRole) = "power" Or LCase(Gen_UserRole) = "admin" Or LCase(Gen_UserRole) = "superadmin" Or LCase(Gen_UserRole) = "headcmg" Then
   mtimes = 1
ElseIf LCase(Gen_UserRole) = "amsandrms" Then
   mtimes = 5
Else
   mtimes = 10
End If

'''-----Query to get GSL Balance
''tmp = "Select isnull(sum(BalanceWeight),0)BalanceWeight from Mst_GSL Where GSLID in (Select GSLID"
''tmp = tmp & " From Txn_CxTF_GSL Where SM_Prefix='" & TxtPreFix & "' And CxTFNo=" & Val(TxtCxTFNo) & " And TxnType='W') And SM_Prefix='" & TxtPreFix & "'"
''
''Call Tmp4Table
''
''If TmpRs4.RecordCount > 0 Then
''   If TmpRs4!BalanceWeight > 0 Then
''      mCommPrice = GetCommodityLatestRate
''      mTotCommPrice = TmpRs4!BalanceWeight * mCommPrice
''      If mTotCommPrice < (mTotOutstAmt * mtimes) Then
''         MsgBox "Balance Commodity Price (" & mTotCommPrice & ") is less than " & mtimes & " times of Outstanding amount(" & mTotOutstAmt & ")"
''         Retval = False
''      End If
''   Else
''      If mTotOutstAmt > 0 Then
''         MsgBox "Outstanding amount (" & mTotOutstAmt & ") is greater than zero."
''         Retval = False
''      End If
''   End If
''End If
mTotBalQty = Val(MSHFlexGrid8.TextMatrix(0, 4))
mTotCommPrice = Val(MSHFlexGrid8.TextMatrix(0, 6))
mTotOutstAmt = Val(MSHFlexGrid9.TextMatrix(0, 12))
If mTotBalQty > 0 Then
   If mTotCommPrice < (mTotOutstAmt * mtimes) Then
      MsgBox "Value of remaining commodity of the client is Rs." & mTotCommPrice & " which is less than " & mtimes & " times the value of outstanding invoice(s)amount of Rs." & mTotOutstAmt & ""
      Retval = False
   End If
Else
   If mTotOutstAmt > 0 Then
      MsgBox "The client has withdrawn all stock but there are still outstanding invoices to the extend of Rs." & mTotOutstAmt & "."
      Retval = False
   End If
End If

CheckClientOutStanding = Retval
End Function

'------ Function to Get outstanding Amount of Clients of this CITF
Private Function GetClientOutstanding()
Dim mOutStand As Variant
mOutStand = 0

''tmp = "Select TIID,TINo,TIDate,PostInvoiceNo,MC.ClientID+'~'+MC.ClientName as ClientName,TI.NAVCode,TIAmount,TISTAmount,"
''tmp = tmp & " (Select isnull(Sum(TCRT.RecAmount),0) as RecAmount From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.CRID= TCRT.CRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y' And TCR.RejectFlag<>'Y' And TIID=TI.TIID And TCR.FAGPostingDate Is Not Null) +isnull(TI.AdjAmountFAG,0) RecAmt,"
''tmp = tmp & " (Select isnull(Sum(TCRT.RecTDS),0) as RecTDS From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.CRID= TCRT.CRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y' And TCR.RejectFlag<>'Y' And TIID=TI.TIID And TCR.FAGPostingDate Is Not Null)+ (Select isnull(Sum(TCRT.RecTDS),0) as RecTDS From Txn_CRTDSAdjustmentDetail TCRT Inner Join Txn_CRTDSAdjustment TCR on TCR.TDSAID= TCRT.TDSAID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.Flag='A' And TIID=TI.TIID) RecTDS, "
''tmp = tmp & " (Select isnull(Sum(TCRT.RecOtherDed),0) as RecOtherDed From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.CRID= TCRT.CRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y' And TCR.RejectFlag<>'Y' And TIID=TI.TIID And TCR.FAGPostingDate Is Not Null)+ (Select isnull(Sum(TCRT.RecOtherDed),0) as RecOtherDed From Txn_CRTDSAdjustmentDetail TCRT Inner Join Txn_CRTDSAdjustment TCR on TCR.TDSAID= TCRT.TDSAID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.Flag='A' And TIID=TI.TIID) as RecOtherDed"
''tmp = tmp & " From  Txn_TIDetails TI "
''tmp = tmp & " Left Join Mst_Client MC on TI.ClientIDRow=MC.ClientIDRow"
''tmp = tmp & " Where (TI.NAVCode in (Select distinct NAVCode From Mst_Client Where ClientIDRow in (Select ClientIDRow from Txn_CxTF_ClientDetail Where SM_Prefix='" & TxtPreFix & "' And CxTFNo=" & Val(TxtCxTFNo) & " And TxnType='W')) Or TI.NAVCode in (Select distinct RNAVCode From Mst_Client Where ClientIDRow in (Select ClientIDRow from Txn_CxTF_ClientDetail Where SM_Prefix='" & TxtPreFix & "' And CxTFNo=" & Val(TxtCxTFNo) & " And TxnType='W')))"
''tmp = tmp & " And (TI.LocationID=" & mLocationID & " )" ''OR LocationID is null

'============Added bacth id > 0 on 7-June 2012 as instruction from Sankar sir
tmp = "Select TIID,TINo,TIDate,PostInvoiceNo,MC.ClientIDRow+'~'+MC.ClientName as ClientName,TI.NAVCode,TIAmount,TISTAmount,"
tmp = tmp & " (Select isnull(Sum(TCRT.RecAmount),0) as RecAmount From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.NewCRID= TCRT.NewCRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y' And TCR.RejectFlag<>'Y' And TIID=TI.TIID And TCR.FAGPostingDate Is Not Null And TCR.BatchID > 0) +isnull(TI.AdjAmountFAG,0) RecAmt,"
tmp = tmp & " (Select isnull(Sum(TCRT.RecTDS),0) as RecTDS From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.NewCRID= TCRT.NewCRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y' And TCR.RejectFlag<>'Y' And TIID=TI.TIID And TCR.FAGPostingDate Is Not Null And TCR.BatchID > 0)+ (Select isnull(Sum(TCRT.RecTDS),0) as RecTDS From Txn_CRTDSAdjustmentDetail TCRT Inner Join Txn_CRTDSAdjustment TCR on TCR.CRTDSAdjustmentID= TCRT.CRTDSAdjustmentID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.Flag='A' And TIID=TI.TIID And TCR.BatchID > 0) RecTDS, "
tmp = tmp & " (Select isnull(Sum(TCRT.RecOtherDed),0) as RecOtherDed From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.NewCRID= TCRT.NewCRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y' And TCR.RejectFlag<>'Y' And TIID=TI.TIID And TCR.FAGPostingDate Is Not Null And TCR.BatchID > 0)+ (Select isnull(Sum(TCRT.RecOtherDed),0) as RecOtherDed From Txn_CRTDSAdjustmentDetail TCRT Inner Join Txn_CRTDSAdjustment TCR on TCR.CRTDSAdjustmentID= TCRT.CRTDSAdjustmentID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.Flag='A' And TIID=TI.TIID And TCR.BatchID > 0) as RecOtherDed"
tmp = tmp & " From  Txn_TIDetails TI "
tmp = tmp & " Left Join Mst_Client MC on TI.ClientID=MC.ClientID"
'Mst_ClientStateWiseGST
'Commeneted on 22 March 2018 by Parag for fetching NAV Code from GST Mapping
'tmp = tmp & " Where (TI.NAVCode in (Select distinct NAVCode From Mst_Client Where ClientID in (Select ClientIDRow from Txn_CxTF_ClientDetail Where SM_Prefix='" & TxtPreFix & "' And CxTFNo=" & Val(TxtCxTFNo) & " And TxnType='W')) Or TI.NAVCode in (Select distinct RNAVCode From Mst_Client Where ClientIDRow in (Select ClientIDRow from Txn_CxTF_ClientDetail Where SM_Prefix='" & TxtPreFix & "' And CxTFNo=" & Val(TxtCxTFNo) & " And TxnType='W')))"
tmp = tmp & " Where (TI.NAVCode in (Select distinct NAVCode From Mst_Client Where ClientID in (Select ClientID from Txn_CxTF_ClientDetail Where SM_Prefix='" & TxtPreFix & "' And CxTF_DetailsID=" & Val(TxtCxTFNo) & " And TxnType='W')) Or TI.NAVCode in (Select distinct RNAVCode From Mst_Client Where ClientID in (Select ClientID from Txn_CxTF_ClientDetail Where SM_Prefix='" & TxtPreFix & "' And CxTF_DetailsID=" & Val(TxtCxTFNo) & " And TxnType='W')) Or TI.NAVCode in (Select distinct GSTNAVCode From Mst_ClientStateWiseGST Where ClientID in (Select ClientID from Txn_CxTF_ClientDetail Where SM_Prefix='" & TxtPreFix & "' And CxTF_DetailsID=" & Val(TxtCxTFNo) & " And TxnType='W')) Or TI.NAVCode in (Select distinct RGSTNAVCode From Mst_ClientStateWiseGST Where ClientID in (Select ClientID from Txn_CxTF_ClientDetail Where SM_Prefix='" & TxtPreFix & "' And CxTF_DetailsID=" & Val(TxtCxTFNo) & " And TxnType='W')))"
tmp = tmp & " And (TI.LocationID=" & mLocationID & " )" ''OR LocationID is null

Call Tmp3Table
Call Grid_Head_OutStand
ToTIAmt = 0
TotRecAmt = 0
If TmpRs3.RecordCount > 0 Then
   For I = 1 To TmpRs3.RecordCount
       If Val(TmpRs3!TIAmount - (TmpRs3!RecAmt + TmpRs3!RecTDS + TmpRs3!RecOtherDed)) > 0 Then
          r = MSHFlexGrid9.Rows - 1
          MSHFlexGrid9.TextMatrix(r, 0) = TmpRs3!TIID
          MSHFlexGrid9.TextMatrix(r, 1) = TmpRs3!TINo
          MSHFlexGrid9.TextMatrix(r, 2) = Format(TmpRs3!TIDate, "dd-MMM-yyyy")
          MSHFlexGrid9.TextMatrix(r, 3) = TmpRs3!PostInvoiceNo
          MSHFlexGrid9.TextMatrix(r, 4) = IIf(IsNull(TmpRs3!ClientName), "", TmpRs3!ClientName)
          MSHFlexGrid9.TextMatrix(r, 5) = TmpRs3!NavCode
          MSHFlexGrid9.TextMatrix(r, 6) = TmpRs3!TIAmount
          MSHFlexGrid9.TextMatrix(r, 7) = TmpRs3!TISTAmount
          MSHFlexGrid9.TextMatrix(r, 8) = TmpRs3!RecAmt
          MSHFlexGrid9.TextMatrix(r, 9) = TmpRs3!RecTDS
          MSHFlexGrid9.TextMatrix(r, 10) = TmpRs3!RecOtherDed
          MSHFlexGrid9.TextMatrix(r, 11) = Format((TmpRs3!RecAmt + TmpRs3!RecTDS + TmpRs3!RecOtherDed), "#####0.00")
          MSHFlexGrid9.TextMatrix(r, 12) = Format(TmpRs3!TIAmount - (TmpRs3!RecAmt + TmpRs3!RecTDS + TmpRs3!RecOtherDed), "#####0.00")
       
          MSHFlexGrid9.TextMatrix(0, 6) = Val(MSHFlexGrid9.TextMatrix(0, 6)) + Val(MSHFlexGrid9.TextMatrix(r, 6))
          MSHFlexGrid9.TextMatrix(0, 7) = Val(MSHFlexGrid9.TextMatrix(0, 7)) + Val(MSHFlexGrid9.TextMatrix(r, 7))
          MSHFlexGrid9.TextMatrix(0, 8) = Val(MSHFlexGrid9.TextMatrix(0, 8)) + Val(MSHFlexGrid9.TextMatrix(r, 8))
          MSHFlexGrid9.TextMatrix(0, 9) = Val(MSHFlexGrid9.TextMatrix(0, 9)) + Val(MSHFlexGrid9.TextMatrix(r, 9))
          MSHFlexGrid9.TextMatrix(0, 10) = Val(MSHFlexGrid9.TextMatrix(0, 10)) + Val(MSHFlexGrid9.TextMatrix(r, 10))
          MSHFlexGrid9.TextMatrix(0, 11) = Val(MSHFlexGrid9.TextMatrix(0, 11)) + Val(MSHFlexGrid9.TextMatrix(r, 11))
          MSHFlexGrid9.TextMatrix(0, 12) = Val(MSHFlexGrid9.TextMatrix(0, 12)) + Val(MSHFlexGrid9.TextMatrix(r, 12))
       
          MSHFlexGrid9.Rows = MSHFlexGrid9.Rows + 1
        '  ToTIAmt = ToTIAmt + TmpRs3!TIAmount
       '   TotRecAmt = TotRecAmt + TmpRs3!RecAmt
         
       End If
       TmpRs3.MoveNext
   Next
   'mOutStand = Format(ToTIAmt - TotRecAmt, "#####0.00")
End If

'GetClientOutstanding = mOutStand
End Function
'----------------------

Private Function GetCommodityLatestRate(mCommodityID_ As Variant) As Variant

Dim mComRate, mStateID As Variant

tmp = " Select StateID from mst_Location Where LocationID=" & mLocationID & ""
Call Tmp2Table
mStateID = 0
If TmpRs2.RecordCount > 0 Then
   mStateID = TmpRs2!StateID
End If

'tmp = " Select MCP.YearMonth, MCP.Day28,MCP.Day29,MCP.Day30,MCP.Day31 from Mst_CommodityPrice MCP"
'tmp = tmp & " Where MCP.StateID = " & mStateID & " And MCP.CommodityID = " & mCommodityID_ & ""
'tmp = tmp & " Order By MCP.YearMonth "

tmp = " Select MCP.StateID,MCP.CommodityID,MCP.YearMonth,"
tmp = tmp & " Case When Convert(date,MCP.YearMonth+'-01') = Convert(date,Convert(Varchar,DATEPART(Year,GETDATE()))+'-'+Convert(Varchar,DATEPART(Month,GETDATE()))+'-01')"
tmp = tmp & " Then"
tmp = tmp & "   COALESCE(MCP.Day31,MCP.Day30,MCP.Day29,MCP.Day28,MCP.Day27,MCP.Day26,MCP.Day25,MCP.Day24,MCP.Day23,MCP.Day22,MCP.Day21,MCP.Day20,MCP.Day19,MCP.Day18,MCP.Day17,MCP.Day16,MCP.Day15,MCP.Day14,MCP.Day13,MCP.Day12,MCP.Day11,MCP.Day10,MCP.Day9,MCP.Day8,MCP.Day7,MCP.Day6,MCP.Day5,MCP.Day4,MCP.Day3,MCP.Day2,MCP.Day1)"
tmp = tmp & " Else"
tmp = tmp & "   COALESCE(MCP.Day31,MCP.Day30,MCP.Day29,MCP.Day28,MCP.Day27,MCP.Day26,MCP.Day25,MCP.Day24,MCP.Day23,MCP.Day22,MCP.Day21,MCP.Day20,MCP.Day19,MCP.Day18,MCP.Day17,MCP.Day16,MCP.Day15,MCP.Day14,MCP.Day13,MCP.Day12,MCP.Day11,MCP.Day10,MCP.Day9,MCP.Day8,MCP.Day7,MCP.Day6,MCP.Day5,MCP.Day4,MCP.Day3,MCP.Day2,MCP.Day1)"
tmp = tmp & " End [Rate]"
tmp = tmp & " from Mst_CommodityPrice MCP"
tmp = tmp & " Where StateID = " & mStateID & " And CommodityID = " & mCommodityID_ & "  And Convert(Varchar,StateID)+'~'+Convert(Varchar,CommodityID)+'~'+Convert(Varchar,YearMonth) In  (Select Convert(Varchar,StateID)+'~'+Convert(Varchar,CommodityID)+'~'+Convert(Varchar,MAX(YearMonth)) From Mst_CommodityPrice Group By StateID,CommodityID )"


Call Tmp2Table

mComRate = 0
'If TmpRs2.RecordCount > 0 Then
'   TmpRs2.MoveLast
'   If Right(TmpRs2!YearMonth, 2) = "01" Or Right(TmpRs2!YearMonth, 2) = "03" Or Right(TmpRs2!YearMonth, 2) = "05" Or Right(TmpRs2!YearMonth, 2) = "07" Or Right(TmpRs2!YearMonth, 2) = "08" Or Right(TmpRs2!YearMonth, 2) = "10" Or Right(TmpRs2!YearMonth, 2) = "12" Then
'      mComRate = IIf(IsNull(TmpRs2!Day31), 0, TmpRs2!Day31)
'   ElseIf Right(TmpRs2!YearMonth, 2) = "04" Or Right(TmpRs2!YearMonth, 2) = "06" Or Right(TmpRs2!YearMonth, 2) = "09" Or Right(TmpRs2!YearMonth, 2) = "11" Then
'      mComRate = IIf(IsNull(TmpRs2!Day30), 0, TmpRs2!Day30) 'TmpRs2!Day30
'   Else
'      If IsNull(TmpRs2!Day29) = False Then
'         mComRate = IIf(IsNull(TmpRs2!Day29), 0, TmpRs2!Day29) 'TmpRs2!Day29
'      Else
'         mComRate = IIf(IsNull(TmpRs2!Day28), 0, TmpRs2!Day28) 'TmpRs2!Day28
'      End If
'   End If
'End If


If TmpRs2.RecordCount > 0 Then
   mComRate = IIf(IsNull(TmpRs2!Rate), 0, TmpRs2!Rate)
End If
GetCommodityLatestRate = mComRate

End Function

Private Sub Grid_Head_Stock()

MSHFlexGrid8.Clear
MSHFlexGrid8.Rows = 3
MSHFlexGrid8.FixedRows = 2
MSHFlexGrid8.Cols = 7

MSHFlexGrid8.TextMatrix(1, 0) = "Godown No"
MSHFlexGrid8.TextMatrix(1, 1) = "Commodity"
MSHFlexGrid8.TextMatrix(1, 2) = "Client Name"
MSHFlexGrid8.TextMatrix(1, 3) = "Balance Bags"
MSHFlexGrid8.TextMatrix(1, 4) = "Balance Weight"
MSHFlexGrid8.TextMatrix(1, 5) = "Commodity Price"
MSHFlexGrid8.TextMatrix(1, 6) = "Total Price"

MSHFlexGrid8.ColWidth(0) = 1200
MSHFlexGrid8.ColWidth(1) = 1400
MSHFlexGrid8.ColWidth(2) = 2500
MSHFlexGrid8.ColWidth(3) = 1400
MSHFlexGrid8.ColWidth(4) = 1400
MSHFlexGrid8.ColWidth(5) = 1400
MSHFlexGrid8.ColWidth(6) = 1400

MSHFlexGrid8.RowHeight(1) = 600

End Sub

Private Sub Grid_Head_OutStand()

MSHFlexGrid9.Clear
MSHFlexGrid9.Rows = 3
MSHFlexGrid9.FixedRows = 2
MSHFlexGrid9.Cols = 13
MSHFlexGrid9.FixedCols = 2


MSHFlexGrid9.TextMatrix(1, 0) = "TIID"
MSHFlexGrid9.TextMatrix(1, 1) = "TI No"
MSHFlexGrid9.TextMatrix(1, 2) = "TI Date"
MSHFlexGrid9.TextMatrix(1, 3) = "Post Invoice No"
MSHFlexGrid9.TextMatrix(1, 4) = "Client Name"
MSHFlexGrid9.TextMatrix(1, 5) = "NAVCode"
MSHFlexGrid9.TextMatrix(1, 6) = "TI Amount"
MSHFlexGrid9.TextMatrix(1, 7) = "ST Amount"
MSHFlexGrid9.TextMatrix(1, 8) = "Rec Amount"
MSHFlexGrid9.TextMatrix(1, 9) = "Rec TDS"
MSHFlexGrid9.TextMatrix(1, 10) = "Rec Other Ded"
MSHFlexGrid9.TextMatrix(1, 11) = "Total Received"
MSHFlexGrid9.TextMatrix(1, 12) = "Balance"

MSHFlexGrid9.ColWidth(0) = 0
MSHFlexGrid9.ColWidth(1) = 1400
MSHFlexGrid9.ColWidth(2) = 1200
MSHFlexGrid9.ColWidth(3) = 1400
MSHFlexGrid9.ColWidth(4) = 2000
MSHFlexGrid9.ColWidth(5) = 1200
MSHFlexGrid9.ColWidth(6) = 1200
MSHFlexGrid9.ColWidth(7) = 1100
MSHFlexGrid9.ColWidth(8) = 1200
MSHFlexGrid9.ColWidth(9) = 1200
MSHFlexGrid9.ColWidth(10) = 1100
MSHFlexGrid9.ColWidth(11) = 1200

MSHFlexGrid9.RowHeight(1) = 600

End Sub

Private Sub FillGSLBalanceStockWithPrice()

Call Grid_Head_Stock
'Call TableTxn_CxTF_GSL("  Where  sm_prefix = '" & TxtPreFix & "' And CxTFNo = " & TxtCxTFNo & " And TxnType = 'W'")
'Call TableMst_GSL("  Where  sm_prefix = '" & TxtPreFix & "' And ClientIDRow in (Select ClientIDRow from Txn_CxTF_ClientDetail Where SM_Prefix='" & TxtPreFix & "' And CxTFNo=" & Val(TxtCxTFNo) & " And TxnType='W') And BalanceWeight>0")

tmp = "Select MG.GodownNo,GSL.CommodityID,MC.Commodity,sum(GSL.BalanceBags)BalanceBags,sum(GSL.BalanceWeight)BalanceWeight,MCL.ClientIDRow+'~'+MCL.ClientName as ClientName"
tmp = tmp & " from Mst_GSL GSL"
tmp = tmp & " Inner Join Mst_Godown MG on MG.GodownID=GSL.GodownID"
tmp = tmp & " Inner Join Mst_Commodity MC on MC.CommodityID=GSL.CommodityID"
tmp = tmp & " Inner Join Mst_Client MCL on MCL.ClientID=GSL.ClientID"
tmp = tmp & " Where  GSL.ExchangeTypeID = 1 And GSL.sm_prefix = '" & TxtPreFix & "'And GSL.ClientID in (Select ClientID from Mst_Client Where NAVCode in(Select distinct NAVCode From Mst_Client Where ClientID in (Select ClientID from Txn_CxTF_ClientDetail Where SM_Prefix='" & TxtPreFix & "' And CxTF_DetailsID=" & Val(TxtCxTFNo) & " And TxnType='W' )And isnull(NAVCode,'')<>'')"
tmp = tmp & " or RNAVCode in (Select distinct RNAVCode From Mst_Client Where ClientID in (Select ClientID from Txn_CxTF_ClientDetail Where SM_Prefix='" & TxtPreFix & "' And CxTF_DetailsID=" & Val(TxtCxTFNo) & " And TxnType='W' )And isnull(RNAVCode,'')<>'')) And GSL.Status not in ('E','Z') And GSL.BalanceBags > 0 And GSL.BalanceWeight > 0"
tmp = tmp & " Group by MG.GodownNo,GSL.CommodityID,MC.Commodity,MCL.ClientIDRow,MCL.ClientName"

Call Tmp5Table

MSHFlexGrid8.Rows = MSHFlexGrid8.Rows + TmpRs5.RecordCount

If TmpRs5.RecordCount = 0 Then Exit Sub

For I = 2 To MSHFlexGrid8.Rows - 2
    MSHFlexGrid8.TextMatrix(I, 0) = TmpRs5!GodownNo
    MSHFlexGrid8.TextMatrix(I, 1) = TmpRs5!Commodity
    MSHFlexGrid8.TextMatrix(I, 2) = TmpRs5!ClientName
    MSHFlexGrid8.TextMatrix(I, 3) = TmpRs5!BalanceBags
    MSHFlexGrid8.TextMatrix(I, 4) = TmpRs5!BalanceWeight
    mCommPrice = GetCommodityLatestRate(TmpRs5!CommodityID)
    MSHFlexGrid8.TextMatrix(I, 5) = mCommPrice
    MSHFlexGrid8.TextMatrix(I, 6) = Format(mCommPrice * TmpRs5!BalanceWeight, "#####0.00")
    MSHFlexGrid8.TextMatrix(0, 3) = Val(MSHFlexGrid8.TextMatrix(0, 3)) + Val(MSHFlexGrid8.TextMatrix(I, 3))
    MSHFlexGrid8.TextMatrix(0, 4) = Val(MSHFlexGrid8.TextMatrix(0, 4)) + Val(MSHFlexGrid8.TextMatrix(I, 4))
    MSHFlexGrid8.TextMatrix(0, 6) = Val(MSHFlexGrid8.TextMatrix(0, 6)) + Val(MSHFlexGrid8.TextMatrix(I, 6))
    TmpRs5.MoveNext
Next

End Sub

Private Function GetBOEAmount() As Variant
Dim Retval As Variant
Retval = 0
tmp = "Select isnull(sum(BOEAmount),0)BOEAmount From  Txn_CxTF_BOE Where CxTF_DetailsID=" & Val(TxtCxTFNo) & " And SM_Prefix='" & TxtPreFix & "'"
Call Tmp4Table
If TmpRs4.RecordCount > 0 Then
   Retval = TmpRs4!BOEAmount
End If
GetBOEAmount = Retval
End Function

Function CheckDuplicateTruck() As Boolean
Dim Retval As Boolean
Retval = False
If MSHFlexGrid3.Rows > 3 Then
   For u = 1 To MSHFlexGrid3.Rows - 1
       If LCase(MSHFlexGrid3.TextMatrix(u, 1)) = LCase(TxtTTDTruckNo.Text) And LCase(MSHFlexGrid3.TextMatrix(u, 2)) = LCase(TxtWeighBridgNo.Text) Then
          If LCase(CmdTTDAdd2List.Caption) = "add to list" Then
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

CheckDuplicateTruck = Retval
End Function

Private Sub CmbGSLTruckNo_GotFocus()
tmp = CmbGSLTruckNo.Text
CmbGSLTruckNo.Clear
If MSHFlexGrid3.Rows < 4 Then Exit Sub

For I = 2 To MSHFlexGrid3.Rows - 2
    CmbGSLTruckNo.AddItem MSHFlexGrid3.TextMatrix(I, 1) & "~" & MSHFlexGrid3.TextMatrix(I, 2)
Next I
CmbGSLTruckNo.Text = tmp
End Sub

Private Sub CmbGSLTruckNo_LostFocus()
If CmbGSLTruckNo.Text = "" Then Exit Sub
For I = 2 To MSHFlexGrid3.Rows - 2
    If CmbGSLTruckNo.Text = MSHFlexGrid3.TextMatrix(I, 1) & "~" & MSHFlexGrid3.TextMatrix(I, 2) Then
       'mTruckNo = MSHFlexGrid3.TextMatrix(I, 11)
       mTruckWtRatio = Val(MSHFlexGrid3.TextMatrix(I, 5)) / Val(MSHFlexGrid3.TextMatrix(I, 6))
       TxtGCTDNoofBags = getRemainingTruckBag(CmbGSLTruckNo.Text)
       If Val(TxtGCTDNoofBags) > Val(TxtGSLAvailableNofBags) Then
          TxtGCTDNoofBags = TxtGSLAvailableNofBags
       End If
       TxtGCTDGrossWt = Format(getRemainingTruckWt(CmbGSLTruckNo.Text), "#######0.000")
       If Val(TxtGCTDGrossWt) > Val(TxtGSLAvailableGrossWgt) Then
          TxtGCTDGrossWt = TxtGSLAvailableGrossWgt
       End If
       Call CalDmatWt
       
       'Add BY Prajakta to get moisture %
       For j = 2 To MSHFlexGrid3.Rows - 1
           If CmbGSLTruckNo.Text = MSHFlexGrid3.TextMatrix(j, 1) & "~" & MSHFlexGrid3.TextMatrix(j, 2) Then
              TxtMoisturePer1.Text = Val(MSHFlexGrid3.TextMatrix(j, 8))
              Exit For
           End If
       Next

       Exit Sub
    End If
Next I

MsgBox "Invalid selection!!!"
CmbGSLTruckNo.SetFocus

End Sub

Private Function getRemainingTruckWt(mTruckNo_ As String) As Double
Dim mTotTruckwt, mTotUseWt As Double
Dim mstr() As String
mTotTruckwt = 0
mTotUseWt = 0
mstr = Split(mTruckNo_, "~")
For I = 2 To MSHFlexGrid3.Rows - 2
    If MSHFlexGrid3.TextMatrix(I, 1) = mstr(0) And MSHFlexGrid3.TextMatrix(I, 2) = mstr(1) Then
       mTotTruckwt = mTotTruckwt + Val(MSHFlexGrid3.TextMatrix(I, 5))
    End If
Next I

For I = 2 To MSHFlexGrid4.Rows - 2
    If MSHFlexGrid4.TextMatrix(I, 19) = mTruckNo_ Then
       If LCase(CmdGCTDAdd2List.Caption) = "add to list" Then
          mTotUseWt = mTotUseWt + Val(MSHFlexGrid4.TextMatrix(I, 7))
       Else
          If I <> MSHFlexGrid4.RowSel Then
             mTotUseWt = mTotUseWt + Val(MSHFlexGrid4.TextMatrix(I, 7))
          End If
       End If
    End If
Next I
getRemainingTruckWt = mTotTruckwt - mTotUseWt
End Function

Private Function getRemainingTruckBag(mTruckNo_ As String) As Double
Dim mTotTruckBag, mTotUseBag As Double
Dim mstr() As String
mTotTruckBag = 0
mTotUseBag = 0
mstr = Split(mTruckNo_, "~")
For I = 2 To MSHFlexGrid3.Rows - 2
    If MSHFlexGrid3.TextMatrix(I, 1) = mstr(0) And MSHFlexGrid3.TextMatrix(I, 2) = mstr(1) Then
       mTotTruckBag = mTotTruckBag + Val(MSHFlexGrid3.TextMatrix(I, 6))
    End If
Next I

For I = 2 To MSHFlexGrid4.Rows - 2
    If MSHFlexGrid4.TextMatrix(I, 19) = mTruckNo_ Then
       If LCase(CmdGCTDAdd2List.Caption) = "add to list" Then
          mTotUseBag = mTotUseBag + Val(MSHFlexGrid4.TextMatrix(I, 6))
       Else
          If I <> MSHFlexGrid4.RowSel Then
             mTotUseBag = mTotUseBag + Val(MSHFlexGrid4.TextMatrix(I, 6))
          End If
       End If
    End If
Next I
getRemainingTruckBag = mTotTruckBag - mTotUseBag
End Function

Private Function getRemainingGSLBag(mGSLID_ As Variant) As Double
Dim mTotWidBag As Double
mTotWidBag = 0

For I = 2 To MSHFlexGrid4.Rows - 2
    If MSHFlexGrid4.TextMatrix(I, 14) = mGSLID_ Then
       If LCase(CmdGCTDAdd2List.Caption) = "add to list" Then
          mTotWidBag = mTotWidBag + Val(MSHFlexGrid4.TextMatrix(I, 6))
       Else
          If I <> MSHFlexGrid4.RowSel Then
             mTotWidBag = mTotWidBag + Val(MSHFlexGrid4.TextMatrix(I, 6))
          End If
       End If
    End If
Next I

getRemainingGSLBag = mTotWidBag
End Function
Private Function getRemainingGSLWt(mGSLID_ As Variant) As Double
Dim mTotWidwt As Double
mTotWidwt = 0

For I = 2 To MSHFlexGrid4.Rows - 2
    If MSHFlexGrid4.TextMatrix(I, 14) = mGSLID_ Then
       If LCase(CmdGCTDAdd2List.Caption) = "add to list" Then
          mTotWidwt = mTotWidwt + Val(MSHFlexGrid4.TextMatrix(I, 7))
       Else
          If I <> MSHFlexGrid4.RowSel Then
             mTotWidwt = mTotWidwt + Val(MSHFlexGrid4.TextMatrix(I, 7))
          End If
       End If
    End If
Next I

getRemainingGSLWt = mTotWidwt
End Function

Private Function CheckCDTFPrintPending(mClientID_ As Variant) As Boolean
Dim Retval As Boolean
Retval = False
tmp = "Select TCD.CxTF_DetailsID,TCD.Flag from Txn_CxTF_Details  TCD"
tmp = tmp & " Inner join Txn_CxTF_ClientDetail TCC on TCC.CxTF_DetailsID=TCD.CxTF_DetailsID And TCC.SM_Prefix=TCD.SM_Prefix And TCC.TxnType=TCD.TxnType"
tmp = tmp & " Where TCD.SM_Prefix='" & TxtPreFix & "' And TCD.TxnType='D' and TCD.Flag is null And TCC.ClientID=" & mClientID_

Call TmpTable
If TmpRs.RecordCount > 0 Then
   Retval = True
End If
CheckCDTFPrintPending = Retval
End Function

Private Function CheckForBOEAmount() As Boolean
Dim Retval As Boolean
Retval = False

If mCommodityYearlyFlag = 0 Or ChkGodownYearlyFlag = False Then
   tmp = "Select TC.CxTFDate,GSL.CDTFDate,day(TC.CxTFDate) NoOfDay,GSL.ClientID,MC.ClientName,TCG.Mst_GSL_ID,MG.GodownNo+'~'+"
   tmp = tmp & " GSL.StackNo +'~'+ GSL.LotNo As GSL,GSL.SpillageFlag,TCG.NoOfBags,TCG.DematWeight,(TCG.DematWeight/TCG.NoOfBags)*1000 as 'BagSize',GSL.NContractID"
   tmp = tmp & " From  Txn_CxTF_GSL TCG"
   tmp = tmp & " Inner Join Txn_CxTF_Details TC on TCG.SM_Prefix=TC.SM_Prefix And TCG.CxTF_DetailsID=TC.CxTF_DetailsID And TCG.TxnType=TC.TxnType"
   tmp = tmp & " left Join Mst_GSL GSL on TCG.Mst_GSL_ID=GSL.Mst_GSL_ID And TCG.SM_Prefix=GSL.SM_Prefix"
   tmp = tmp & " Inner Join Mst_Godown MG on GSL.GodownID=MG.GodownID"
   tmp = tmp & " Inner Join Mst_Client MC on GSL.ClientID=MC.ClientID"
   tmp = tmp & " Where TCG.NoOfBags > 0 And TC.ExchangeTypeID=1 And TC.CxTF_DetailsID=" & Val(TxtCxTFNo) & " And TC.TxnType='W' And TC.SM_Prefix='" & TxtPreFix.Text & "' And GSL.SpillageFlag = 0 And isnull(GSL.NContractID,0) = 0 And GSL.MadeupFlag = 0  "
'tmp = tmp & " And DateDiff(MM,TC.CxTFDate,GSL.CDTFDate) > 0 "
 
   Call TmpTable
   If TmpRs.RecordCount > 0 Then
      Retval = True
   End If
End If
CheckForBOEAmount = Retval

End Function

Private Sub CmbALocationID_GotFocus()
tmp = CmbALocationID.Text
If Gen_WcLocation = "" Then
   Call TableMst_Location("")
Else
   Call TableMst_Location("where " & Gen_WcLocation)
End If



CmbALocationID.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbALocationID.AddItem RsMst_Location!Locationname
    RsMst_Location.MoveNext
Next
CmbALocationID.Text = tmp
End Sub

Private Sub CmbALocationID_LostFocus()
If CmbALocationID.Text = "" Then
''   mLocationID = Null
''   mSM_Prefix = Null
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbALocationID.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection "
   CmbALocationID.SetFocus
   Exit Sub
End If

If tmp <> CmbALocationID.Text Then
   CmbACMPID.Text = ""
End If
mALocationID = RsMst_Location!LocationID
mALocationName = RsMst_Location!Locationname
mASM_prefix = RsMst_Location!SM_Prefix
End Sub

Private Sub CmbACMPID_GotFocus()
tmp = CmbACMPID.Text
If mASM_prefix = "" Then
   MsgBox "Select Location !!!"
   CmbALocationID.SetFocus
   Exit Sub
End If
Call TableMst_CMP(" Where LocationID = " & mALocationID & "")
CmbACMPID.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "No record found!!! "
   Exit Sub
End If
For I = 1 To RsMst_CMP.RecordCount
    CmbACMPID.AddItem RsMst_CMP!CMPName
    RsMst_CMP.MoveNext
Next
CmbACMPID.Text = tmp
End Sub

Private Sub CmbACMPID_LostFocus()
If CmbACMPID.Text = "" Then
   Exit Sub
End If
RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPName = '" & CmbACMPID.Text & "'"
If RsMst_CMP.EOF Then
   MsgBox "Invalid selection!!!"
   CmbACMPID.SetFocus
   Exit Sub
End If
mACMPName = RsMst_CMP!CMPID
End Sub

Private Sub CmbACommodityID_GotFocus()
tmp = CmbACommodityID.Text
Call TableMst_Commodity
CmbACommodityID.Clear
If RsMst_Commodity.RecordCount = 0 Then
   MsgBox "No record found!!! "
   Exit Sub
End If
For I = 1 To RsMst_Commodity.RecordCount
 CmbACommodityID.AddItem RsMst_Commodity!Commodity
 RsMst_Commodity.MoveNext
Next
CmbACommodityID.Text = tmp
End Sub

Private Sub CmbACommodityID_LostFocus()
If CmbACommodityID.Text = "" Then
   Exit Sub
End If
RsMst_Commodity.MoveFirst
RsMst_Commodity.Find "Commodity = '" & CmbACommodityID.Text & "'"
If RsMst_Commodity.EOF Then
   MsgBox "Invalid selection!!!"
   CmbACommodityID.SetFocus
   Exit Sub
End If

mACommodityID = RsMst_Commodity!CommodityID

End Sub

Private Sub Grid_Head_ApprovaPrint()
MSHFlexGrid10.Clear
MSHFlexGrid10.Rows = 2
MSHFlexGrid10.Cols = 18
MSHFlexGrid10.TextMatrix(0, 0) = "CITF No"
MSHFlexGrid10.TextMatrix(0, 1) = "CITF Date"
MSHFlexGrid10.TextMatrix(0, 2) = "Exchange Type"
MSHFlexGrid10.TextMatrix(0, 3) = "CMP"
MSHFlexGrid10.TextMatrix(0, 4) = "Location ID"
MSHFlexGrid10.TextMatrix(0, 5) = "Commodity"
MSHFlexGrid10.TextMatrix(0, 6) = "Flag"
MSHFlexGrid10.TextMatrix(0, 7) = "Agent"
MSHFlexGrid10.TextMatrix(0, 8) = "Remark"
MSHFlexGrid10.TextMatrix(0, 9) = "Hologram No"
MSHFlexGrid10.TextMatrix(0, 10) = "Agent Name"
MSHFlexGrid10.TextMatrix(0, 11) = "Manual CITF No"
MSHFlexGrid10.TextMatrix(0, 12) = "No of Bags"
MSHFlexGrid10.TextMatrix(0, 13) = "Gross Wt."
MSHFlexGrid10.TextMatrix(0, 14) = "Guid"
MSHFlexGrid10.TextMatrix(0, 15) = "Rejected"
MSHFlexGrid10.TextMatrix(0, 16) = "Balance Stock"
MSHFlexGrid10.TextMatrix(0, 17) = "Outstanding Amt"

MSHFlexGrid1.ColWidth(0) = 1000
MSHFlexGrid1.ColWidth(1) = 1000
MSHFlexGrid1.ColWidth(2) = 1400
MSHFlexGrid1.ColWidth(3) = 1600
MSHFlexGrid1.ColWidth(4) = 1200
MSHFlexGrid1.ColWidth(5) = 1200
MSHFlexGrid1.ColWidth(6) = 1200
MSHFlexGrid1.ColWidth(7) = 1200
MSHFlexGrid1.ColWidth(8) = 1600
MSHFlexGrid1.ColWidth(9) = 0
MSHFlexGrid1.ColWidth(10) = 1200
MSHFlexGrid1.ColWidth(11) = 1200
MSHFlexGrid1.ColWidth(14) = 0
MSHFlexGrid1.ColWidth(15) = 1000

MSHFlexGrid1.RowHeight(0) = 700

End Sub

Private Sub DeleteFromGSLDetail()
I = MSHFlexGrid3.RowSel
Dim mDelRows As Variant
If I > 0 Then
   If MSHFlexGrid3.TextMatrix(I, 1) <> "" Then
      mDelRows = 0
Restart:
        For r = 2 To MSHFlexGrid4.Rows - 1
            If LCase(MSHFlexGrid4.TextMatrix(r, 19)) = LCase(MSHFlexGrid3.TextMatrix(I, 1) & "~" & MSHFlexGrid3.TextMatrix(I, 2)) Then
               MSHFlexGrid4.TextMatrix(0, 4) = Val(MSHFlexGrid4.TextMatrix(0, 4)) - Val(MSHFlexGrid4.TextMatrix(r, 4))
               MSHFlexGrid4.TextMatrix(0, 5) = Val(MSHFlexGrid4.TextMatrix(0, 5)) - Val(MSHFlexGrid4.TextMatrix(r, 5))
               MSHFlexGrid4.TextMatrix(0, 6) = Val(MSHFlexGrid4.TextMatrix(0, 6)) - Val(MSHFlexGrid4.TextMatrix(r, 6))
               MSHFlexGrid4.TextMatrix(0, 7) = Val(MSHFlexGrid4.TextMatrix(0, 7)) - Val(MSHFlexGrid4.TextMatrix(r, 7))
               MSHFlexGrid4.TextMatrix(0, 8) = Val(MSHFlexGrid4.TextMatrix(0, 8)) - Val(MSHFlexGrid4.TextMatrix(r, 8))
               MSHFlexGrid4.TextMatrix(0, 9) = Val(MSHFlexGrid4.TextMatrix(0, 9)) - Val(MSHFlexGrid4.TextMatrix(r, 9))
               MSHFlexGrid4.TextMatrix(0, 10) = Val(MSHFlexGrid4.TextMatrix(0, 10)) - Val(MSHFlexGrid4.TextMatrix(r, 10))
               MSHFlexGrid4.TextMatrix(0, 11) = Val(MSHFlexGrid4.TextMatrix(0, 11)) - Val(MSHFlexGrid4.TextMatrix(r, 11))
               Call UpdateWHRQty(r)
               For RO = r To MSHFlexGrid4.Rows - 2
                   For C = 0 To MSHFlexGrid4.Cols - 1
                       MSHFlexGrid4.TextMatrix(RO, C) = MSHFlexGrid4.TextMatrix(RO + 1, C)
                   Next
               Next
               MSHFlexGrid4.Rows = MSHFlexGrid4.Rows - 1
               GoTo Restart
            End If
        Next
         

   End If
End If

End Sub

Private Function getClientNAVCodeFromGrid() As String
Dim Retval As String
Retval = ""

For I = 2 To MSHFlexGrid2.Rows - 2
    If LCase(CmdAdd2ListRRN.Caption) = "add to list" Then
       Retval = MSHFlexGrid2.TextMatrix(I, 18)
       If Retval <> "" Then Exit For
    Else
       If I <> MSHFlexGrid2.RowSel Then
          Retval = MSHFlexGrid2.TextMatrix(I, 18)
          If Retval <> "" Then Exit For
       End If
    End If
Next I

getClientNAVCodeFromGrid = Retval
End Function

Private Function getClientRNAVCodeFromGrid() As String
Dim Retval As String
Retval = ""

For I = 2 To MSHFlexGrid2.Rows - 2
    If LCase(CmdAdd2ListRRN.Caption) = "add to list" Then
       Retval = MSHFlexGrid2.TextMatrix(I, 19)
       If Retval <> "" Then Exit For
    Else
       If I <> MSHFlexGrid2.RowSel Then
          Retval = MSHFlexGrid2.TextMatrix(I, 19)
          If Retval <> "" Then Exit For
       End If
    End If
Next I

getClientRNAVCodeFromGrid = Retval
End Function

Private Sub UpdateWHRQty(mROw_ As Variant)

OAdjBags = Val(MSHFlexGrid4.TextMatrix(mROw_, 6))
OAdjWt = Val(MSHFlexGrid4.TextMatrix(mROw_, 10))
   
For ir = 2 To MSHFlexGrid2.Rows - 2
    If MSHFlexGrid4.TextMatrix(mROw_, 12) = MSHFlexGrid2.TextMatrix(ir, 1) & " ~ " & MSHFlexGrid2.TextMatrix(ir, 4) Then
       MSHFlexGrid2.TextMatrix(0, 15) = (Val(MSHFlexGrid2.TextMatrix(0, 15)) - OAdjBags)
       MSHFlexGrid2.TextMatrix(0, 9) = (Val(MSHFlexGrid2.TextMatrix(0, 9)) - OAdjWt)
     
       MSHFlexGrid2.TextMatrix(ir, 15) = (Val(MSHFlexGrid2.TextMatrix(ir, 15)) - OAdjBags)
       MSHFlexGrid2.TextMatrix(ir, 9) = (Val(MSHFlexGrid2.TextMatrix(ir, 9)) - OAdjWt)
       
       If Val(MSHFlexGrid4.TextMatrix(ir, 17)) = 1 Or Val(MSHFlexGrid4.TextMatrix(ir, 20)) = 1 Then
          MSHFlexGrid2.TextMatrix(0, 16) = (Val(MSHFlexGrid2.TextMatrix(0, 16)) - OAdjBags)
          MSHFlexGrid2.TextMatrix(0, 17) = (Val(MSHFlexGrid2.TextMatrix(0, 17)) - OAdjWt)
     
          MSHFlexGrid2.TextMatrix(ir, 16) = (Val(MSHFlexGrid2.TextMatrix(ir, 16)) - OAdjBags)
          MSHFlexGrid2.TextMatrix(ir, 17) = (Val(MSHFlexGrid2.TextMatrix(ir, 17)) - OAdjWt)
       End If
       Exit For
    End If
  
Next
End Sub

Private Sub AssignMultOf()
tmp = "select MultOfAll,MultOfAM from Mst_ExchangeType where ExchangeTypeID=" & mExchangeTypeID
Call Tmp1Table
If TmpRs1.RecordCount > 0 Then
   mMultOfAll = IIf(IsNull(TmpRs1!MultOfAll), 0, TmpRs1!MultOfAll)
   mMultOfAM = IIf(IsNull(TmpRs1!MultOfAM), 0, TmpRs1!MultOfAM)
End If

End Sub
Private Function CheckVoucheNo(mVouchNo_ As String) As Boolean
Dim Retval As Boolean
Dim mTIDate As Variant
Retval = True
''For Ln = 2 To MSHFlexGrid9.Rows - 1
''    If LCase(MSHFlexGrid9.TextMatrix(Ln, 1)) = LCase(mVouchNo_) Then
''       RetVal = True
''       mTIDate = CDate(MSHFlexGrid9.TextMatrix(Ln, 2))
''       Exit For
''    End If
''Next

tmp = "Select TIID,TINo,TIDate,PostInvoiceNo,MC.ClientID+'~'+MC.ClientName as ClientName,TI.NAVCode,TIAmount,TISTAmount,"
tmp = tmp & " (Select isnull(Sum(TCRT.RecAmount),0) as RecAmount From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.NewCRID= TCRT.NewCRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y' And TCR.RejectFlag<>'Y' And TIID=TI.TIID And TCR.FAGPostingDate Is Not Null And TCR.BatchID > 0) +isnull(TI.AdjAmountFAG,0) RecAmt,"
tmp = tmp & " (Select isnull(Sum(TCRT.RecTDS),0) as RecTDS From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.NewCRID= TCRT.NewCRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y' And TCR.RejectFlag<>'Y' And TIID=TI.TIID And TCR.FAGPostingDate Is Not Null And TCR.BatchID > 0)+ (Select isnull(Sum(TCRT.RecTDS),0) as RecTDS From Txn_CRTDSAdjustmentDetail TCRT Inner Join Txn_CRTDSAdjustment TCR on TCR.CRTDSAdjustmentID= TCRT.CRTDSAdjustmentID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.Flag='A' And TIID=TI.TIID And TCR.BatchID > 0) RecTDS, "
tmp = tmp & " (Select isnull(Sum(TCRT.RecOtherDed),0) as RecOtherDed From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.NewCRID= TCRT.NewCRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y' And TCR.RejectFlag<>'Y' And TIID=TI.TIID And TCR.FAGPostingDate Is Not Null And TCR.BatchID > 0)+ (Select isnull(Sum(TCRT.RecOtherDed),0) as RecOtherDed From Txn_CRTDSAdjustmentDetail TCRT Inner Join Txn_CRTDSAdjustment TCR on TCR.CRTDSAdjustmentID= TCRT.CRTDSAdjustmentID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.Flag='A' And TIID=TI.TIID And TCR.BatchID > 0) as RecOtherDed"
tmp = tmp & " From  Txn_TIDetails TI "
tmp = tmp & " Left Join Mst_Client MC on TI.ClientID=MC.ClientID"
tmp = tmp & " Where (TI.NAVCode in (Select distinct NAVCode From Mst_Client Where ClientID in (Select ClientID from Txn_CxTF_ClientDetail Where SM_Prefix='" & TxtPreFix & "' And CxTF_DetailsID=" & Val(TxtCxTFNo) & " And TxnType='W')) Or TI.NAVCode in (Select distinct RNAVCode From Mst_Client Where ClientID in (Select ClientID from Txn_CxTF_ClientDetail Where SM_Prefix='" & TxtPreFix & "' And CxTF_DetailsID=" & Val(TxtCxTFNo) & " And TxnType='W')))"
tmp = tmp & " And (TI.LocationID=" & mLocationID & " ) And "
If OptPostNo.Value = True Then
   tmp = tmp & " TI.PostInvoiceNo='" & Trim(mVouchNo_) & "'"
Else
   tmp = tmp & " TI.TINo='" & Trim(mVouchNo_) & "'"
End If


Call Tmp3Table
If TmpRs3.RecordCount > 0 Then
   If TmpRs3!TIDate < TxtCxTFDate.Value Then
      MsgBox "TI Date for this invoice is less than CITF Date!!"
      Retval = False
   End If
   mTIID = TmpRs3!TIID
Else
   If OptPostNo.Value = True Then
      MsgBox "Post Invoice No " & mVouchNo_ & " not found!!"
   Else
      MsgBox "TI No " & mVouchNo_ & " not found!!"
   End If
   Retval = False
End If

CheckVoucheNo = Retval
End Function

Private Sub GetClientDetailsWHRAdd(mClientIDRow_ As Double)

If mClientIDRow_ = 0 Or IsNull(mClientIDRow_) Then Exit Sub
ans = GetClientNameByRow(mClientIDRow_)
If ans(0) = "" Then
   MsgBox "Not found!!!"
   Exit Sub
End If
TxtClientID = ans(0)
TxtClientName = ans(1)
TxtDPID = ans(2)
TxtDPName = ans(3)
mNAVCode = ans(6)
mRNAVCode = ans(7)

If mNAVCode <> "" And getClientNAVCodeFromGrid() <> "" Then
   If mNAVCode <> getClientNAVCodeFromGrid() Then
      MsgBox "Please select client having same NAV Code!!"
      TxtRRNNo.SetFocus
      Exit Sub
   End If
End If

If mRNAVCode <> "" And getClientRNAVCodeFromGrid() <> "" Then
   If mRNAVCode <> getClientRNAVCodeFromGrid() Then
      MsgBox "Please select client having same RNAV Code!!"
      TxtRRNNo.SetFocus
      Exit Sub
   End If
End If

If CheckCDTFPrintPending(mClientIDRow_) = True Then
   MsgBox "Print pending for one or more CDTF of this client!!"
   TxtRRNNo.SetFocus
   Exit Sub
End If

End Sub

Private Sub ApproveCITFPrint()

If LCase(CmbExchangeTypeID.Text) = "non ncdex" Then

   mTotBalQty = Val(MSHFlexGrid8.TextMatrix(0, 4))
   mTotCommPrice = Val(MSHFlexGrid8.TextMatrix(0, 6))
   mTotOutstAmt = mTotOutstAmt + Val(MSHFlexGrid9.TextMatrix(0, 12))
   mGridOutAmt = Val(MSHFlexGrid9.TextMatrix(0, 12))
   
   If mTotBalQty > 0 Then
      If mTotCommPrice < (mTotOutstAmt * mtimes) Then
         If mtimes <> 1 Then
            MsgBox "Value of remaining commodity of the client is Rs." & mTotCommPrice & " which is less than " & mtimes & " times the value of outstanding invoice(s)amount of Rs." & mTotOutstAmt & ", hence this withdrawal not approve."
            Frame2.Visible = False
            Exit Sub
         Else
            If MsgBox("Value of remaining commodity of the client is Rs." & mTotCommPrice & " which is less than the value of outstanding invoice(s)amount of Rs." & mTotOutstAmt & ". Still you want to approve print?", vbYesNo) = vbNo Then
               Frame2.Visible = False
               Exit Sub
            End If
         End If
      End If
   Else
      If mTotOutstAmt > 0 Then
         If mtimes <> 1 Then
            If mGridOutAmt > 0 Then
               MsgBox "The client has withdrawn all stock but there are still outstanding invoices to the extend of Rs." & mTotOutstAmt & " hence this withdrawal not approve."
               Exit Sub
            End If
         Else
            If MsgBox("The client has withdrawn all stock but there are still outstanding invoices to the extend of Rs." & mTotOutstAmt & ". Still you want to approve print?", vbYesNo) = vbNo Then
               Frame2.Visible = False
               Exit Sub
            End If
         End If
      End If
   End If
Else
   mTotBalQty = 0
   mTotCommPrice = 0
   mTotOutstAmt = 0
   mGridOutAmt = 0
   mTIID = 0
End If

If TxtARemark = "" Then
   MsgBox "Blank Remark not allowed!!!"
   TxtARemark.SetFocus
   Exit Sub
End If

Frame2.Visible = False
RsTxn_CxTF_Details!ApprovePrintFlag = "Y"
RsTxn_CxTF_Details!ApprovePrintBy = Gen_UserLoginID
RsTxn_CxTF_Details!ApprovePrintRemark = TxtARemark.Text
RsTxn_CxTF_Details!ApprovePrintDate = Now
RsTxn_CxTF_Details!OutstandingAmt = mTotOutstAmt
RsTxn_CxTF_Details!CommodityValue = mTotCommPrice
RsTxn_CxTF_Details!BalQty = mTotBalQty
RsTxn_CxTF_Details!TIID = Val(mTIID)
RsTxn_CxTF_Details.Update
RsTxn_CxTF_Details.Requery
If RsTxn_CxTF_Details.RecordCount > 0 Then
   RsTxn_CxTF_Details.MoveFirst
End If

 
RsTxn_CxTF_Details.Find "G_UID='" & TxtGUID & "'"

CmdAPrint.Enabled = False

End Sub
'Private Sub SaveCITFBOE()
'
'Dim mVar As Variant
'Dim mTotDays, mInvAmt, mLotSize, mMinLotSize, m7DaysFlag As Variant
'
'tmp = "Select Collect7Days from Mst_Location where LocationID=" & mLocationID
'Call TmpTable
'If TmpRs.RecordCount > 0 Then
'   m7DaysFlag = IIf(IsNull(TmpRs!Collect7Days), 0, TmpRs!Collect7Days)
'End If
'
'Call TableTxn_CDTFBOE(" Where SM_Prefix='" & TxtPreFix & "' And CDTFNo=" & Val(TxtCxTFNo) & " And TxnType='W'")
'
'If RsTxn_CDTFBOE.RecordCount > 0 Then Exit Sub
'
'Call TableTxn_CDTFBOE_GSL(" Where SM_Prefix='" & TxtPreFix & "'")
'
'tmp = "Select GodownID,Rate from Mst_NCDEXStorageRateCardDetail MNRC"
'tmp = tmp & " inner Join Mst_NCDEXStorageRateCard MNR on MNRC.NSRCID=MNR.NSRCID"
'tmp = tmp & " Inner Join Mst_CommodityGroup MCG On MNR.CommodityGroupID = MCG.CommodityGroupID"
'tmp = tmp & " Inner Join Mst_Commodity MComm On MCG.CommodityGroupID = MComm.CommodityGroupID"
'tmp = tmp & " Where MComm.CommodityID=" & mCommodityID & "  And MNR.LocationID=" & mLocationID
'
'Call Tmp2Table
'
'tmp = "select god.GodownNo 'GodownNo',gsl.StackNo,gsl.LotNo,GSL.CDTFDate,GSL.LCDTF_Date,"
'tmp = tmp & " sum(tgsl.NoofBags)NoofBags,sum(tgsl.GrossWeight)GrossWeight,sum(tgsl.AdjNClientWeight)AdjNClientWeight,"
'tmp = tmp & " sum(tgsl.AdjNCMSLWeight)AdjNCMSLWeight,sum(tgsl.GunnyWeight)GunnyWeight,sum(tgsl.StdDeduction)StdDeduction,"
'tmp = tmp & " sum(tgsl.DematWeight)DematWeight,c.ClientName + ' ' + d.DepositoryName ClientName,tgsl.Remarks,tgsl.GSLID,"
'tmp = tmp & " god.CloseDate , god.GodownID, gsl.ClientIDRow, gsl.ISINID, gsl.FGSLID, god.Address"
'tmp = tmp & " From"
'tmp = tmp & " Txn_CxTF_GSL tgsl"
'tmp = tmp & " inner join mst_gsl gsl on gsl.GSLID = tgsl.GSLID and gsl.SM_Prefix = tgsl.SM_Prefix"
'tmp = tmp & " inner join Mst_Godown god on gsl.GodownID = god.GodownID"
'tmp = tmp & " inner join mst_client c on gsl.ClientIDRow = c.ClientIDRow"
'tmp = tmp & " inner join mst_depository d on d.DepositoryID = c.DepositoryID  and d.ExchangeTypeID =" & mExchangeTypeID
'tmp = tmp & " Where tgsl.TxnType = 'W'  And  tgsl.CxTFNo =  " & TxtCxTFNo & "  And tgsl.SM_Prefix = '" & TxtPreFix & "' And MadeupFlag = '0' "
'tmp = tmp & " Group By god.GodownNo,gsl.StackNo,gsl.LotNo,c.ClientName ,d.DepositoryName,"
'tmp = tmp & " tgsl.Remarks,tgsl.GSLID,god.CloseDate,god.GodownID,gsl.ClientIDRow,gsl.ISINID,gsl.FGSLID,GSL.CDTFDate,GSL.LCDTF_Date,"
'tmp = tmp & " god.Address"
'
'Call Tmp3Table
'
'mRejInvIDs = ""
' For I = 1 To TmpRs3.RecordCount
'
'    If TmpRs3!GSLID <> "" Then
'      If TmpRs2.RecordCount > 0 Then TmpRs2.MoveFirst
'      TmpRs2.Find "GodownID=" & TmpRs3!godownid
'      If TmpRs2.EOF = False Then
'         mRate = TmpRs2!Rate
'      Else
'         mRate = 0
'      End If
'
'
'      If m7DaysFlag = 1 Then
'         If Check7DaysInvoice(TmpRs3!GSLID) = True Then
'            mTotDays = DateDiff("d", DateAdd("d", 7, TmpRs3!LCDTF_Date), TxtCxTFDate.Value) + 1
'         Else
'            mTotDays = DateDiff("d", TmpRs3!CDTFDate, TxtCxTFDate.Value) + 1
'         End If
'      Else
'         mTotDays = DateDiff("d", TmpRs3!CDTFDate, TxtCxTFDate.Value) + 1
'      End If
'      If mTotDays > 0 Then
'         If RsTxn_CDTFBOE.RecordCount > 0 Then RsTxn_CDTFBOE.MoveFirst
'         RsTxn_CDTFBOE.Find "ClientIDRow=" & TmpRs3!ClientIDRow
'
'         If RsTxn_CDTFBOE.EOF = True Then
'            Call AddCDTFBOE(Val(TmpRs3!ClientIDRow))
'         End If
'
'         Call SaveCDTFBOEGSL(Val(TmpRs3!GSLID), TmpRs3!NoofBags, TmpRs3!DematWeight, mTotDays, mRate)
'       End If
'    End If
'    TmpRs3.MoveNext
'Next
'
'End Sub
Private Sub SaveYearlyBilling()

Dim mVar As Variant
Dim mTotDays, mInvAmt, mLotSize, mMinLotSize, m7DaysFlag As Variant

Call TableTxn_CDTFBOE(" Where SM_Prefix='" & TxtPreFix & "' And CxTF_DetailsID=" & Val(TxtCxTFNo) & " And TxnType='Y'")

If RsTxn_CDTFBOE.RecordCount > 0 Then Exit Sub

Call TableTxn_CDTFBOE_GSL(" Where SM_Prefix='" & TxtPreFix & "'")

tmp = "Select GodownID,Rate from Mst_YearlyStorageRateCardDetail MNRC"
tmp = tmp & " inner Join Mst_YearlyStorageRateCard MNR on MNRC.YSRCID=MNR.YSRCID"
tmp = tmp & " Where MNR.CommodityID=" & mCommodityID & "  And MNR.LocationID=" & mLocationID

Call Tmp2Table

'mRate = IIf(IsNull(TmpRs3!NNNRYearBillingRate), 1, TmpRs3!NNNRYearBillingRate)

tmp = "select god.GodownNo 'GodownNo',gsl.StackNo,gsl.LotNo,GSL.CDTFDate,GSL.LCDTF_Date,"
tmp = tmp & " sum(tgsl.NoofBags)NoofBags,sum(tgsl.GrossWeight)GrossWeight,sum(tgsl.AdjNClientWeight)AdjNClientWeight,"
tmp = tmp & " sum(tgsl.AdjNCMSLWeight)AdjNCMSLWeight,sum(tgsl.GunnyWeight)GunnyWeight,sum(tgsl.StdDeduction)StdDeduction,"
tmp = tmp & " sum(tgsl.DematWeight)DematWeight,c.ClientName + ' ' + d.DepositoryName ClientName,tgsl.Remarks,tgsl.Mst_GSL_ID,"
tmp = tmp & " god.CloseDate , god.GodownID, gsl.ClientID, gsl.ISINID, gsl.FGSLID_Mst_GSL_ID, god.Address"
tmp = tmp & " From"
tmp = tmp & " Txn_CxTF_GSL tgsl"
tmp = tmp & " inner join mst_gsl gsl on gsl.Mst_GSL_ID = tgsl.Mst_GSL_ID and gsl.SM_Prefix = tgsl.SM_Prefix"
tmp = tmp & " inner join Mst_Godown god on gsl.GodownID = god.GodownID"
tmp = tmp & " inner join mst_client c on gsl.ClientID = c.ClientID"
tmp = tmp & " inner join mst_depository d on d.DepositoryID = c.DepositoryID  and d.ExchangeTypeID =" & mExchangeTypeID
tmp = tmp & " Where tgsl.TxnType = 'W'  And  tgsl.CxTF_DetailsID =  " & TxtCxTFNo & "  And tgsl.SM_Prefix = '" & TxtPreFix & "' and IsNull(gsl.MadeupFlag,0) = 0"
tmp = tmp & " Group By god.GodownNo,gsl.StackNo,gsl.LotNo,c.ClientName ,d.DepositoryName,"
tmp = tmp & " tgsl.Remarks,tgsl.Mst_GSL_ID,god.CloseDate,god.GodownID,gsl.ClientID,gsl.ISINID,gsl.FGSLID_Mst_GSL_ID,GSL.CDTFDate,GSL.LCDTF_Date,"
tmp = tmp & " god.Address"
  
Call Tmp3Table

mRejInvIDs = ""
For I = 1 To TmpRs3.RecordCount
    If TmpRs3!Mst_GSL_ID <> "" Then
       
       If TmpRs2.RecordCount > 0 Then TmpRs2.MoveFirst
       TmpRs2.Find "GodownID=" & TmpRs3!GodownID
       If TmpRs2.EOF = False Then
          mRate = TmpRs2!Rate
       Else
          mRate = 0
       End If
      
       If RsTxn_CDTFBOE.RecordCount > 0 Then RsTxn_CDTFBOE.MoveFirst
       RsTxn_CDTFBOE.Find "ClientID=" & TmpRs3!ClientID
       If RsTxn_CDTFBOE.EOF = True Then
          mInvAmountOnPrint = 0
          mBOEIDNew = 0
          Call AddCDTFBOE(Val(TmpRs3!ClientID), "Y")
       End If
       Call SaveCDTFBOEGSL(Val(TmpRs3!Mst_GSL_ID), TmpRs3!NoofBags, TmpRs3!DematWeight, 1, mRate, Format(1 * TmpRs3!NoofBags * mRate, "#######0"))
       
       tmp = " Update  Txn_CDTFBOE Set Amount = " & Val(mInvAmountOnPrint) & " Where BOEID =  " & Val(mBOEIDNew) & "   "
       Call Tmp7Table
       
    End If
    TmpRs3.MoveNext
Next

End Sub
Private Sub SaveCITFBOE()

Dim mVar As Variant
Dim mTotDays, mInvAmt, mLotSize, mMinLotSize, m7DaysFlag As Variant

tmp = "Select Collect7Days from Mst_Location where LocationID=" & mLocationID
Call TmpTable
If TmpRs.RecordCount > 0 Then
   m7DaysFlag = IIf(IsNull(TmpRs!Collect7Days), 0, TmpRs!Collect7Days)
End If

Call TableTxn_CDTFBOE(" Where SM_Prefix='" & TxtPreFix & "' And CxTF_DetailsID=" & Val(TxtCxTFNo) & " And TxnType='W'")

If RsTxn_CDTFBOE.RecordCount > 0 Then Exit Sub

Call TableTxn_CDTFBOE_GSL(" Where SM_Prefix='" & TxtPreFix & "'")

tmp = "Select GodownID,Rate from Mst_NCDEXStorageRateCardDetail MNRC"
tmp = tmp & " inner Join Mst_NCDEXStorageRateCard MNR on MNRC.NSRCID=MNR.NSRCID"
tmp = tmp & " Inner Join Mst_CommodityGroup MCG On MNR.CommodityGroupID = MCG.CommodityGroupID"
tmp = tmp & " Inner Join Mst_Commodity MComm On MCG.CommodityGroupID = MComm.CommodityGroupID"
tmp = tmp & " Where MComm.CommodityID=" & mCommodityID & "  And MNR.LocationID=" & mLocationID

Call Tmp2Table

tmp = "select god.GodownNo 'GodownNo',gsl.StackNo,gsl.LotNo,GSL.CDTFDate,GSL.LCDTF_Date,"
tmp = tmp & " sum(tgsl.NoofBags)NoofBags,sum(tgsl.GrossWeight)GrossWeight,sum(tgsl.AdjNClientWeight)AdjNClientWeight,"
tmp = tmp & " sum(tgsl.AdjNCMSLWeight)AdjNCMSLWeight,sum(tgsl.GunnyWeight)GunnyWeight,sum(tgsl.StdDeduction)StdDeduction,"
tmp = tmp & " sum(tgsl.DematWeight)DematWeight,c.ClientName + ' ' + d.DepositoryName ClientName,tgsl.Remarks,tgsl.Mst_GSL_ID,"
tmp = tmp & " god.CloseDate , god.GodownID, gsl.ClientID, gsl.ISINID, gsl.FGSLID_Mst_GSL_ID, god.Address"
tmp = tmp & " From"
tmp = tmp & " Txn_CxTF_GSL tgsl"
tmp = tmp & " inner join mst_gsl gsl on gsl.Mst_GSL_ID = tgsl.Mst_GSL_ID and gsl.SM_Prefix = tgsl.SM_Prefix"
tmp = tmp & " inner join Mst_Godown god on gsl.GodownID = god.GodownID"
tmp = tmp & " inner join mst_client c on gsl.ClientID = c.ClientID"
tmp = tmp & " inner join mst_depository d on d.DepositoryID = c.DepositoryID  and d.ExchangeTypeID =" & mExchangeTypeID
tmp = tmp & " Where tgsl.TxnType = 'W'  And  tgsl.CxTF_DetailsID =  " & TxtCxTFNo & "  And tgsl.SM_Prefix = '" & TxtPreFix & "' and IsNull(gsl.MadeupFlag,0) = 0"
tmp = tmp & " Group By god.GodownNo,gsl.StackNo,gsl.LotNo,c.ClientName ,d.DepositoryName,"
tmp = tmp & " tgsl.Remarks,tgsl.Mst_GSL_ID,god.CloseDate,god.GodownID,gsl.ClientID,gsl.ISINID,gsl.FGSLID_Mst_GSL_ID,GSL.CDTFDate,GSL.LCDTF_Date,"
tmp = tmp & " god.Address"
  
Call Tmp3Table

mRejInvIDs = ""
For I = 1 To TmpRs3.RecordCount

    If TmpRs3!Mst_GSL_ID <> "" Then
      If TmpRs2.RecordCount > 0 Then TmpRs2.MoveFirst
      TmpRs2.Find "GodownID=" & TmpRs3!GodownID
      If TmpRs2.EOF = False Then
         mRate = TmpRs2!Rate
      Else
         mRate = 0
      End If
   
      
      If m7DaysFlag = True Then
      
         If Check7DaysInvoice(TmpRs3!Mst_GSL_ID) = True Then
            mTotDays = DateDiff("d", DateAdd("d", 7, TmpRs3!LCDTF_Date), TxtCxTFDate.Value) + 1
         Else
            mBGSLIDs = ""
            Call GetBGSLIDs(TxtPreFix, TmpRs3!Mst_GSL_ID)
            If mBGSLIDs <> "" Then
               If Check7DaysInvoice(mBGSLIDs) = True Then
                  mTotDays = DateDiff("d", DateAdd("d", 7, TmpRs3!LCDTF_Date), TxtCxTFDate.Value) + 1
               Else
                  mTotDays = DateDiff("d", TmpRs3!CDTFDate, TxtCxTFDate.Value) + 1
               End If
            Else
               mTotDays = DateDiff("d", TmpRs3!CDTFDate, TxtCxTFDate.Value) + 1
            End If
         End If
      Else
         mTotDays = DateDiff("d", TmpRs3!CDTFDate, TxtCxTFDate.Value) + 1
      End If
      If mTotDays > 0 Then
         If RsTxn_CDTFBOE.RecordCount > 0 Then RsTxn_CDTFBOE.MoveFirst
         RsTxn_CDTFBOE.Find "ClientID=" & TmpRs3!ClientID
         
         If RsTxn_CDTFBOE.EOF = True Then
            mInvAmountOnPrint = 0
            mBOEIDNew = 0
            Call AddCDTFBOE(Val(TmpRs3!ClientID), "W")
         End If
         
         Call SaveCDTFBOEGSL(Val(TmpRs3!Mst_GSL_ID), TmpRs3!NoofBags, TmpRs3!DematWeight, mTotDays, mRate, Format(mTotDays * TmpRs3!DematWeight * mRate, "#####0.00"))
       
         tmp = " Update  Txn_CDTFBOE Set Amount = " & Val(mInvAmountOnPrint) & " Where BOEID =  " & Val(mBOEIDNew) & "   "
         Call Tmp7Table
       End If
    End If
    TmpRs3.MoveNext
Next

End Sub

'Private Sub AddCDTFBOE(mClientID_ As Variant)
'
'RsTxn_CDTFBOE.AddNew
'RsTxn_CDTFBOE!BOEID = getCDTFBOE_Max
'RsTxn_CDTFBOE!LocationID = mLocationID
'RsTxn_CDTFBOE!SM_Prefix = TxtPreFix
'RsTxn_CDTFBOE!CdtfNo = TxtCxTFNo
'RsTxn_CDTFBOE!CDTFDate = TxtCxTFDate.Value
'RsTxn_CDTFBOE!ClientIDRow = mClientID_
'RsTxn_CDTFBOE!InvoiceDate = TxtCxTFDate.Value
'RsTxn_CDTFBOE!BillingMonth = CDate(Year(TxtCxTFDate.Value) & "-" & Month(TxtCxTFDate.Value) & "-01")
'RsTxn_CDTFBOE!Discount = 0
'RsTxn_CDTFBOE!Premium = 0
'RsTxn_CDTFBOE!BatchID = 0
'RsTxn_CDTFBOE!RBatchID = 0
'RsTxn_CDTFBOE!Flag = "P"
'RsTxn_CDTFBOE!TxnType = "W"
'RsTxn_CDTFBOE!G_UID = GetGUID
'RsTxn_CDTFBOE!CreatedDate = Now
'RsTxn_CDTFBOE!CreatedBy = Gen_UserLoginID
'RsTxn_CDTFBOE.Update
'
'If mRejInvIDs = "" Then
'   mRejInvIDs = RsTxn_CDTFBOE!BOEID
'Else
'   mRejInvIDs = mRejInvIDs & "," & RsTxn_CDTFBOE!BOEID
'End If
'
'End Sub
Private Sub AddCDTFBOE(mClientID_ As Variant, TxnType_ As Variant)

RsTxn_CDTFBOE.AddNew
'RsTxn_CDTFBOE!BOEID = getCDTFBOE_Max
'mBOEIDNew = RsTxn_CDTFBOE!BOEID
RsTxn_CDTFBOE!LocationID = mLocationID
RsTxn_CDTFBOE!SM_Prefix = TxtPreFix
RsTxn_CDTFBOE!CxTF_DetailsID = TxtCxTFNo
RsTxn_CDTFBOE!CDTFDate = TxtCxTFDate.Value
RsTxn_CDTFBOE!ClientID = mClientID_
RsTxn_CDTFBOE!InvoiceDate = TxtCxTFDate.Value
RsTxn_CDTFBOE!BillingMonth = CDate(Year(TxtCxTFDate.Value) & "-" & Month(TxtCxTFDate.Value) & "-01")
RsTxn_CDTFBOE!Discount = 0
RsTxn_CDTFBOE!Premium = 0
RsTxn_CDTFBOE!BatchID = 0
RsTxn_CDTFBOE!RBatchID = 0
RsTxn_CDTFBOE!Flag = "P"
RsTxn_CDTFBOE!TxnType = TxnType_
RsTxn_CDTFBOE!G_UID = GetGUID
mInvG_UID = RsTxn_CDTFBOE!G_UID
RsTxn_CDTFBOE!CreatedDate = Now
RsTxn_CDTFBOE!CreatedBy = Gen_UserLoginID
RsTxn_CDTFBOE.Update

tmp = "Select BOEID from Txn_CDTFBOE Where G_UID = '" & mInvG_UID & "' "
Call TmpTable

If TmpRs.RecordCount > 0 Then
   mBOEIDNew = TmpRs!BOEID
End If

If mRejInvIDs = "" Then
   mRejInvIDs = mBOEIDNew 'RsTxn_CDTFBOE!BOEID
Else
   mRejInvIDs = mRejInvIDs & "," & mBOEIDNew 'RsTxn_CDTFBOE!BOEID
End If

End Sub


'Private Sub SaveCDTFBOEGSL(mGSLID_ As Variant, mNoofBags_ As Variant, mDematWeight_ As Variant, mNoOfDays_ As Variant, mRate_ As Variant)
'
'RsTxn_CDTFBOE_GSL.AddNew
'RsTxn_CDTFBOE_GSL!BOEGSLID = getCDTFBOE_GSL_Max
'RsTxn_CDTFBOE_GSL!BOEID = RsTxn_CDTFBOE!BOEID
'RsTxn_CDTFBOE_GSL!SM_Prefix = TxtPreFix
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
'tmp = "Update Txn_CxTF_GSL Set BillFlag='I' Where SM_Prefix='" & TxtPreFix & "' And CxTFNo=" & TxtCxTFNo & " And TxnType='W' And GSLID=" & mGSLID_ & " And BillFlag is null"
'Call Tmp5Table
'
'End Sub

Private Sub SaveCDTFBOEGSL(mGSLID_ As Variant, mNoofBags_ As Variant, mDematWeight_ As Variant, mNoOfDays_ As Variant, mRate_ As Variant, amt_ As Variant)

RsTxn_CDTFBOE_GSL.AddNew
'RsTxn_CDTFBOE_GSL!BOEGSLID = getCDTFBOE_GSL_Max
RsTxn_CDTFBOE_GSL!BOEID = mBOEIDNew 'RsTxn_CDTFBOE!BOEID
RsTxn_CDTFBOE_GSL!SM_Prefix = TxtPreFix
RsTxn_CDTFBOE_GSL!Mst_GSL_ID = mGSLID_
RsTxn_CDTFBOE_GSL!CDTFBags = mNoofBags_
RsTxn_CDTFBOE_GSL!CDTFQty = mDematWeight_
RsTxn_CDTFBOE_GSL!NoOfDays = mNoOfDays_
RsTxn_CDTFBOE_GSL!Rate = mRate_ 'TmpRs2!Rate
If mExchangeTypeID = 1 Then
   RsTxn_CDTFBOE_GSL!BillingUnit = "B" 'TmpRs2!Unit
   RsTxn_CDTFBOE_GSL!BillingCycleID = 6 ''TmpRs2!BillingCycleID
Else
   RsTxn_CDTFBOE_GSL!BillingUnit = "Q" 'TmpRs2!Unit
   RsTxn_CDTFBOE_GSL!BillingCycleID = 1 ''TmpRs2!BillingCycleID
End If
RsTxn_CDTFBOE_GSL!Amount = amt_ 'Format(mNoOfDays_ * mDematWeight_ * mRate_, "#####0.00")
RsTxn_CDTFBOE_GSL!NCDEXSpaceCharges = 0 'Val(mSpaceChargesForInvoice)
RsTxn_CDTFBOE_GSL.Update

mInvAmountOnPrint = Val(mInvAmountOnPrint) + RsTxn_CDTFBOE_GSL!Amount
'RsTxn_CDTFBOE.Update

tmp = "Update Txn_CxTF_GSL Set BillFlag='I' Where SM_Prefix='" & TxtPreFix & "' And CxTF_DetailsID=" & TxtCxTFNo & " And TxnType='W' And Mst_GSL_ID=" & mGSLID_ & " And BillFlag is null"
Call Tmp5Table

End Sub



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

Private Function CheckNCDEXRate() As Boolean
Dim Retval As Boolean
Retval = True
tmp = "Select GodownID,Rate from Mst_NCDEXStorageRateCardDetail MNRC"
tmp = tmp & " inner Join Mst_NCDEXStorageRateCard MNR on MNRC.NSRCID=MNR.NSRCID"
tmp = tmp & " Inner Join Mst_CommodityGroup MCG On MNR.CommodityGroupID = MCG.CommodityGroupID"
tmp = tmp & " Inner Join Mst_Commodity MComm On MCG.CommodityGroupID = MComm.CommodityGroupID"
tmp = tmp & " Where MComm.CommodityID=" & mCommodityID & "  And MNR.LocationID=" & mLocationID
 
Call Tmp2Table

tmp = "Select distinct gsl.GodownID From"
tmp = tmp & " Txn_CxTF_GSL tgsl"
tmp = tmp & " inner join mst_gsl gsl on gsl.Mst_GSL_ID = tgsl.Mst_GSL_ID and gsl.SM_Prefix = tgsl.SM_Prefix"
tmp = tmp & " Where tgsl.TxnType = 'W'  And  tgsl.CxTF_DetailsID =  " & TxtCxTFNo & "  And tgsl.SM_Prefix = '" & TxtPreFix & "'"

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

Private Function GetBalAmt(InvoiceID_ As Double, Mod_ As Variant) As Variant
Dim mBalAmt As Variant
mBalAmt = 0

tmp = "Select TIID,TINo,TIDate,PostInvoiceNo,MC.ClientIDRow+'~'+MC.ClientName as ClientName,TI.NAVCode,TIAmount,TISTAmount,"
tmp = tmp & " (Select isnull(Sum(TCRT.RecAmount),0) as RecAmount From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.NewCRID= TCRT.NewCRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y' And TCR.RejectFlag<>'Y' And TIID=TI.TIID And TCR.Flag ='P') +isnull(TI.AdjAmountFAG,0) RecAmt,"
tmp = tmp & " (Select isnull(Sum(TCRT.RecTDS),0) as RecTDS From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.NewCRID= TCRT.NewCRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y' And TCR.RejectFlag<>'Y' And TIID=TI.TIID And TCR.Flag ='P')+ (Select isnull(Sum(TCRT.RecTDS),0) as RecTDS From Txn_CRTDSAdjustmentDetail TCRT Inner Join Txn_CRTDSAdjustment TCR on TCR.CRTDSAdjustmentID= TCRT.CRTDSAdjustmentID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.Flag='A' And TIID=TI.TIID ) RecTDS, "
tmp = tmp & " (Select isnull(Sum(TCRT.RecOtherDed),0) as RecOtherDed From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.NewCRID= TCRT.NewCRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y' And TCR.RejectFlag<>'Y' And TIID=TI.TIID And TCR.Flag ='P')+ (Select isnull(Sum(TCRT.RecOtherDed),0) as RecOtherDed From Txn_CRTDSAdjustmentDetail TCRT Inner Join Txn_CRTDSAdjustment TCR on TCR.CRTDSAdjustmentID= TCRT.CRTDSAdjustmentID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.Flag='A' And TIID=TI.TIID ) as RecOtherDed"
tmp = tmp & " From  Txn_TIDetails TI "
tmp = tmp & " Left Join Mst_Client MC on TI.ClientID=MC.ClientID"
If LCase(Mod_) = "w" Then
   tmp = tmp & " Where TI.PIID=" & InvoiceID_ & " And TI.PIFlag='RJ'"
ElseIf LCase(Mod_) = "r" Then
   tmp = tmp & " Where TI.PIID=" & InvoiceID_ & " And TI.PIFlag='RRN'"
ElseIf LCase(Mod_) = "y" Then
   tmp = tmp & " Where TI.PIID=" & InvoiceID_ & " And TI.PIFlag='NNNRY'"
End If

Call Tmp2Table

If TmpRs2.RecordCount > 0 Then
   mBalAmt = Format(TmpRs2!TIAmount - (TmpRs2!RecAmt + TmpRs2!RecTDS + TmpRs2!RecOtherDed), "#####0.00")
Else
   mBalAmt = "NA"
End If
GetBalAmt = mBalAmt

End Function
'Private Function Check7DaysInvoice(gslID_ As Variant) As Boolean
'Dim Retval As Boolean
'Retval = False
'tmp = "Select TC.BOEID,TCG.GSLID  from  Txn_CDTFBOE_GSL  TCG"
'tmp = tmp & " Inner join Txn_CDTFBOE TC on TCG.BOEID=TC.BOEID"
'tmp = tmp & " Where TCG.GSLID=" & gslID_ & " And TCG.SM_Prefix='" & TxtPreFix & "' And TC.TxnType='D'"
'
'Call Tmp6Table
'
'If TmpRs6.RecordCount > 0 Then
'   Retval = True
'End If
'Check7DaysInvoice = Retval
'End Function
Private Function Check7DaysInvoice(gslID_ As Variant) As Boolean
Dim Retval As Boolean
Retval = False
tmp = "Select TC.BOEID,TCG.Mst_GSL_ID  from  Txn_CDTFBOE_GSL  TCG"
tmp = tmp & " Inner join Txn_CDTFBOE TC on TCG.BOEID=TC.BOEID"
tmp = tmp & " Where TCG.Mst_GSL_ID in (" & gslID_ & ") And TCG.SM_Prefix='" & TxtPreFix & "' And TC.TxnType='D'"

Call Tmp6Table
 
If TmpRs6.RecordCount > 0 Then
   Retval = True
End If
Check7DaysInvoice = Retval
End Function
Private Sub GetBGSLIDs(SM_Prefix_ As Variant, gslID_ As Variant)
tmp = "Select Mst_GSL_ID,Status,BGSLID_Mst_GSL_ID From Mst_GSL Where SM_Prefix = '" & SM_Prefix_ & "' And Mst_GSL_ID = " & gslID_ & " "
Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   If IsNull(TmpRs1!BGSLID_Mst_GSL_ID) = False Then
      If mBGSLIDs = "" Then
         mBGSLIDs = TmpRs1!Mst_GSL_ID
      Else
         mBGSLIDs = mBGSLIDs & "," & TmpRs1!Mst_GSL_ID
      End If
      Call GetBGSLIDs(SM_Prefix_, TmpRs1!BGSLID_Mst_GSL_ID)
   Else
      If mBGSLIDs = "" Then
         mBGSLIDs = TmpRs1!Mst_GSL_ID
      Else
         mBGSLIDs = mBGSLIDs & "," & TmpRs1!Mst_GSL_ID
      End If
      Exit Sub
   End If
End If

End Sub



Private Function ChkMadeUpLots() As Boolean
Dim Retval As Boolean
Retval = True

For x = 2 To MSHFlexGrid4.Rows - 1
    If MSHFlexGrid4.TextMatrix(x, 1) = "" Then Exit For
    If MSHFlexGrid4.TextMatrix(x, 20) = 0 Then
       Retval = False
       Exit For
    End If
Next
ChkMadeUpLots = Retval

End Function

Private Function GetTotalPDFPages() As Double

Set objPDF = New mjwPDF

' Create a simple PDF file using the mjwPDF class

Gen_mTimeStamp = GetTimeStamp
'Call Create_Xls(Gen_mTimeStamp & Gen_UserName & "_CDTF-" & Trim(TxtPreFix) & "" & Trim(TxtCxTFNo) & ".xls")
    
    
' Set the PDF title and filename
objPDF.PDFTitle = "CDTF PDF Document"
objPDF.PDFFileName = App.Path & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "_CITF-" & Trim(TxtPreFix) & "" & Trim(TxtCxTFNo) & ".pdf"
mPDFFileName = App.Path & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "_CITF-" & Trim(TxtPreFix) & "" & Trim(TxtCxTFNo) & ".pdf"
' We must tell the class where the PDF fonts are located
objPDF.PDFLoadAfm = App.Path & "\Fonts"

' Set the file properties
objPDF.PDFSetLayoutMode = LAYOUT_DEFAULT
objPDF.PDFFormatPage = FORMAT_A4
objPDF.PDFOrientation = ORIENT_PORTRAIT
objPDF.PDFSetUnit = UNIT_PT
objPDF.PDFSetViewerPreferences = VIEW_HIDEMENUBAR & VIEW_HIDETOOLBAR & VIEW_HIDEWINDOWUI & VIEW_HIDERIGHTCLICK '& VIEW_DISABLEFILEMENU


' Lets us set see the bookmark pane when we view the PDF
'objPDF.PDFUseOutlines = True

' View the PDF file after we create it
objPDF.PDFView = False

' Begin our PDF document

objPDF.PDFBeginDoc
    ' Lets add a heading
objPDF.PDFSetFont FONT_ARIAL, 12, FONT_BOLD
objPDF.PDFSetDrawColor = vbWhite
objPDF.PDFSetTextColor = vbBlack
objPDF.PDFSetAlignement = ALIGN_CENTER
objPDF.PDFSetBorder = BORDER_ALL
objPDF.PDFSetFill = True

objPDF.PDFCell " ", 15, 15, _
objPDF.PDFGetPageWidth - 30, objPDF.PDFGetPageHeight - 20

If PntPrv = "Y" Then
   objPDF.PDFWatermark = "Preview"
Else
   objPDF.PDFWatermark = "NCMSL"
End If

objPDF.PDFSetFont FONT_ARIAL, 12, FONT_BOLD
objPDF.PDFSetDrawColor = vbWhite
objPDF.PDFSetTextColor = vbBlack
objPDF.PDFSetAlignement = ALIGN_CENTER
objPDF.PDFSetBorder = BORDER_ALL
objPDF.PDFSetFill = True


pgval = ((objPDF.PDFGetPageWidth - 18))

'''objPDF.PDFDrawPolygon Array(12, 15, objPDF.PDFGetPageWidth - 18, 15, objPDF.PDFGetPageWidth - 18, 30, 12, 30)
''
''mRowHeight = 15
''mGlobalLeft = 12
'''objPDF.PDFDrawPolygon Array(12, 15, objPDF.PDFGetPageWidth - 18, 15, objPDF.PDFGetPageWidth - 18, 30, 12, 30)
''mLeft = mGlobalLeft
''mTop = 15
''mHeight = mTop + mRowHeight
''mTxtFromTop = mRowHeight + 2
''
'' 'Call DrawTable(10, 2, 15, 30, 12)
''
''Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
''
''mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
''mStrLength = Len("NCMSL " & TxtCMPID & "")
''mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
''mTxtLeft = mPrintPoint - (mStrLength * 2) '* 5) '270
''mFontSize = 12
''
''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_CENTER, mTxtLeft, mTxtTop, "NCMSL " & TxtCMPID.Text & "")
''
''mHeight = mTop + mRowHeight
''mLeft = mGlobalLeft
''
''
''Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
''
''mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
''mStrLength = Len("Commodity Issue Transaction Form (CITF)")
''mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
''mTxtLeft = mPrintPoint - (mStrLength * 2) '* 5) '270
''mFontSize = 10
''
''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_CENTER, mTxtLeft, mTxtTop, "Commodity Issue Transaction Form (CITF)")
''
''mHeight = mTop + mRowHeight
''mLeft = mGlobalLeft
''
''Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
''
''
''mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
''mStrLength = Len("(This is NOT an instrument of title of the goods)")
''mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
''mTxtLeft = mPrintPoint '- (mStrLength) '* 5) '270
''mFontSize = 6
''
''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_CENTER, mTxtLeft - 30, mTxtTop - 3, "(This is NOT an instrument of title of the goods)")
''
'''--Code added By Kishor
'''Row 1
''
''
''mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
''mStrLength = Len("Depositor Details")
''mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
''mTxtLeft = mPrintPoint '- (mStrLength) '* 5) '270
''mFontSize = 7
''Dim mDispTextSize  As Integer
''mDispTextSize = 6
''
''mTxtTop = mTxtTop + mRowHeight
''
''Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop, "CITF No")
''Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop, TxtCxTFNo)
'''--
''Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 280, mTxtTop, "Date")
''Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 430, mTxtTop, Format(TxtCxTFDate.Value, "dd-MMM-yyyy"))
'''--
''mTxtTop = mTxtTop + mRowHeight
''Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop, "Location")
''Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop, GetLocationFromCMP(Val(mCMPID)))
'''--
''Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 280, mTxtTop, "State")
''Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 430, mTxtTop, GetStateFromLocation(GetLocationIDFromCMP(Val(mCMPID))))
'''--
''mTxtTop = mTxtTop + mRowHeight
''Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop, "CMPID & Address")
''Dim mCMPIDnAddr As String
''mCMPIDnAddr = mCMPID & "-" & GetCMPAddress(Val(mCMPID), "A")
''Dim mLen As Integer, mMax As Integer, mPara As Integer
''Dim t1 As Double
''mMax = 35
''mPara = 0
''t1 = 1
''mLen = Len(mCMPIDnAddr)
''mLen = Round(mLen / mMax) '+ 1
''
''For t = 1 To mLen
''    If t > 1 Then
''       t1 = mMax + t1
''       'mMax '= mMax + 35
''       mPara = mPara + 7
''    End If
''    Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop + mPara, Mid$(mCMPIDnAddr, t1, mMax))
''Next
''
'''--
''Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 280, mTxtTop, "Commodity")
''Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 430, mTxtTop, TxtCommodityID)
''
''Call DrawTable(2, 4, mTop, mHeight, mLeft, mRowHeight)
''mHeight = mHeight + 15
''Call DrawTable(, 4, mTop, mHeight, mLeft, mRowHeight)
''
''
'''Row 2
''
''mTxtTop = mHeight - (mTxtFromTop) + 15 'mTop ' + 15
''mStrLength = Len("Transaction Tyepe")
''mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
''mTxtLeft = mPrintPoint '- (mStrLength) '* 5) '270
''mFontSize = 7
''
''Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop, "Transaction Type")
''Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop, CmbExchangeTypeID)
''
''mTxtTop = mTxtTop + mRowHeight
''
''Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop, "Name Of Agent")
''Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop, TxtAgent)
''
''objPDF.PDFDrawLine mGlobalLeft + 132, mTop - 4 * mRowHeight, mGlobalLeft + 132, mHeight + mRowHeight
''Call DrawTable(2, , mTop, mHeight, mLeft, mRowHeight)
''
''
'''Row 3

mRowHeight = 15
mGlobalLeft = 12
'objPDF.PDFDrawPolygon Array(12, 15, objPDF.PDFGetPageWidth - 18, 15, objPDF.PDFGetPageWidth - 18, 30, 12, 30)
mLeft = mGlobalLeft
mTop = 15
mHeight = mTop + mRowHeight
mTxtFromTop = mRowHeight + 2
 
 
 
 'Call DrawTable(10, 2, 15, 30, 12)

Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)

mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
mStrLength = Len("NCMSL " & TxtCMPID & "")
mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
mTxtLeft = mPrintPoint - (mStrLength * 2) '* 5) '270
mFontSize = 12

Call PrintText(mFontSize, FONT_NORMAL, ALIGN_CENTER, mTxtLeft, mTxtTop, "NCMSL " & TxtCMPID.Text & "")

If Totpages > 1 Then
   Call PrintText(6, FONT_NORMAL, ALIGN_RIGHT, objPDF.PDFGetPageWidth - 70, mTxtTop - 5, "Page " & objPDF.PDFPageNumber & " out of " & Totpages)
Else
   
End If
mHeight = mTop + mRowHeight
mLeft = mGlobalLeft


Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)

mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
mStrLength = Len("Commodity Deposit Transaction Form (CDTF)")
mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
mTxtLeft = mPrintPoint - (mStrLength * 2) '* 5) '270
mFontSize = 10

Call PrintText(mFontSize, FONT_NORMAL, ALIGN_CENTER, mTxtLeft, mTxtTop, "Commodity Issue Transaction Form (CITF)")

mHeight = mTop + mRowHeight
mLeft = mGlobalLeft

Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
 
tmp1 = mTop
mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
mStrLength = Len("(This is NOT an instrument of title of the goods)")
mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
mTxtLeft = mPrintPoint '- (mStrLength) '* 5) '270
mFontSize = 6

Call PrintText(mFontSize, FONT_NORMAL, ALIGN_CENTER, mTxtLeft - 30, mTxtTop - 3, "(This is NOT an instrument of title of the goods)")

'--Code added By Kishor
'Row 1

mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
mStrLength = Len("Depositor Details")
mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
mTxtLeft = mPrintPoint '- (mStrLength) '* 5) '270
mFontSize = 7
Dim mDispTextSize  As Integer
mDispTextSize = 6

mTxtTop = mTxtTop + mRowHeight

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop, "CITF No")
Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop, TxtCxTFNo)
'--
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 350, mTxtTop, "Date")
Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 450, mTxtTop, Format(TxtCxTFDate, "dd-MMM-yyyy"))

mTxtTop = mTxtTop + mRowHeight
objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop, "Location")
Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop, GetLocationFromCMP(mCMPID))
'--
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 350, mTxtTop, "State")
Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 450, mTxtTop, GetStateFromLocation(GetLocationIDFromCMP(mCMPID)))
'--
mTxtTop = mTxtTop + mRowHeight
objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop, "CMPID & Address")
Dim mCMPIDnAddr As String
mCMPIDnAddr = mCMPID & "-" & GetCMPAddress(mCMPID, "A")
Dim mLen As Integer, mMax As Integer, mPara As Integer
Dim t1 As Double
mMax = 45
mPara = 0
t1 = 1
mLen = Len(mCMPIDnAddr)
mLen = Round(mLen / mMax) '+ 1

For T = 1 To mLen
    If T > 1 Then
       t1 = mMax + t1
       'mMax '= mMax + 35
       mPara = mPara + 7
    End If
    Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop + mPara, mID$(mCMPIDnAddr, t1, mMax))
Next

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 350, mTxtTop, "Commodity")
Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 450, mTxtTop, TxtCommodityID)

mTxtTop = mTxtTop + mPara + mRowHeight

objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10

objPDF.PDFDrawLine mGlobalLeft + 345, Val(tmp1), mGlobalLeft + 345, mTxtTop - 10
objPDF.PDFDrawLine mGlobalLeft + 445, Val(tmp1), mGlobalLeft + 445, mTxtTop - 10

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop, "Transaction Type")
Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop, CmbExchangeTypeID)

mTxtTop = mTxtTop + mRowHeight
objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop, "Name Of Agent")
Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop, TxtAgent)

mTxtTop = mTxtTop + mRowHeight
objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10
objPDF.PDFDrawLine mGlobalLeft + 135, Val(tmp1), mGlobalLeft + 135, mTxtTop - 10

mTxtTop = mTxtTop + mRowHeight '- (mTxtFromTop) 'mTop ' + 15
mStrLength = Len("Withdrawer Details")
mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
mTxtLeft = mPrintPoint '- (mStrLength) '* 5) '270
mFontSize = 8
mHeight = mTxtTop
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mLeft, mTxtTop - 2, "Withdrawer Details")

Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)

objPDF.PDFDrawLine mGlobalLeft + 72, mTop, mGlobalLeft + 72, mTop + ((MSHFlexGrid2.Rows - 2) * mRowHeight)
objPDF.PDFDrawLine mGlobalLeft + 142, mTop, mGlobalLeft + 142, mTop + ((MSHFlexGrid2.Rows - 2) * mRowHeight)
If mExchangeTypeID <> 2 Then
   objPDF.PDFDrawLine 2.8 * (mGlobalLeft + 142), mTop, 2.8 * (mGlobalLeft + 142), mTop + ((MSHFlexGrid2.Rows - 1) * mRowHeight)
End If
objPDF.PDFDrawLine 3.3 * (mGlobalLeft + 142), mTop, 3.3 * (mGlobalLeft + 142), mTop + ((MSHFlexGrid2.Rows - 1) * mRowHeight)

Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)


mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
mStrLength = Len("Depositor Details")
mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
mTxtLeft = mPrintPoint '- (mStrLength) '* 5) '270
mFontSize = 7

If CmbExchangeTypeID.Text = "NCDEX" Then
   If Check1.Value = 0 Then
      Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mLeft, mTxtTop, "RRN No")
   Else
      Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mLeft, mTxtTop, "Rejected Withdrawal")
   End If
ElseIf CmbExchangeTypeID.Text = "Non NCDEX" Then
   
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mLeft, mTxtTop, "WHR No")
ElseIf CmbExchangeTypeID.Text = "SPOT" Then
   
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mLeft, mTxtTop, "CDC No")
End If

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mLeft + 80, mTxtTop, "Client ID")
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 150, mTxtTop, "Client Name")
If mExchangeTypeID <> 2 Then
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 430, mTxtTop, "No of Bags")
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 500, mTxtTop - 6, "Electronic Weight")
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 500, mTxtTop, "in MT")
 
Else
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 500, mTxtTop - 6, "Electronic Weight")
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 500, mTxtTop, "in MT")
 
End If

objPDF.PDFDrawLine mGlobalLeft + 72, mTop, mGlobalLeft + 72, mTop + ((MSHFlexGrid2.Rows - 2) * mRowHeight)
objPDF.PDFDrawLine mGlobalLeft + 142, mTop, mGlobalLeft + 142, mTop + ((MSHFlexGrid2.Rows - 2) * mRowHeight)
If mExchangeTypeID <> 2 Then
   objPDF.PDFDrawLine 2.8 * (mGlobalLeft + 142), mTop, 2.8 * (mGlobalLeft + 142), mTop + ((MSHFlexGrid2.Rows - 1) * mRowHeight)
End If
objPDF.PDFDrawLine 3.3 * (mGlobalLeft + 142), mTop, 3.3 * (mGlobalLeft + 142), mTop + ((MSHFlexGrid2.Rows - 1) * mRowHeight)
objPDF.PDFDrawLine mGlobalLeft, mTop + mRowHeight, mGlobalLeft + 560, mTop + mRowHeight
Call DrawTable(1, , mTop, mHeight, mLeft, mRowHeight)
 
'FILL GRID WITH DETAILS
Dim mTemp As Integer
For T = 2 To MSHFlexGrid2.Rows - 1
    If mHeight > objPDF.PDFGetPageHeight Then
       Call CheckForNewPage
       mTop = 15
       mTxtTop = 10
    End If
    If T = 1 Then
       mTemp = mFontSize
    Else
       mTemp = mDispTextSize
    End If
    mTxtTop = mTxtTop + mRowHeight
   
    Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mLeft, mTxtTop, MSHFlexGrid2.TextMatrix(T, 1))
    Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 80, mTxtTop, MSHFlexGrid2.TextMatrix(T, 3))
    Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 150, mTxtTop, MSHFlexGrid2.TextMatrix(T, 4))
    If mExchangeTypeID <> 2 Then
       Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 430, mTxtTop, Format(MSHFlexGrid2.TextMatrix(T, 15), "#######0.000"))
    End If
    Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 500, mTxtTop, Format(MSHFlexGrid2.TextMatrix(T, 9), "#######0.000"))
    
    objPDF.PDFDrawLine mGlobalLeft + 72, mTop, mGlobalLeft + 72, mTop + mRowHeight
    objPDF.PDFDrawLine mGlobalLeft + 142, mTop, mGlobalLeft + 142, mTop + mRowHeight
    If mExchangeTypeID <> 2 Then
       objPDF.PDFDrawLine 2.8 * (mGlobalLeft + 142), mTop, 2.8 * (mGlobalLeft + 142), mTop + mRowHeight
    End If
    objPDF.PDFDrawLine 3.3 * (mGlobalLeft + 142), mTop, 3.3 * (mGlobalLeft + 142), mTop + mRowHeight

    'Call DrawTable(MSHFlexGrid2.Rows - 1, , mTop, mHeight, mLeft, mRowHeight)
    objPDF.PDFDrawLine mGlobalLeft, mTop + mRowHeight, mGlobalLeft + 565, mTop + mRowHeight

    mTop = mTop + mRowHeight
    mHeight = mHeight + mRowHeight
Next

Call CheckForNewPage
mTxtTop = mTxtTop + mRowHeight
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 220, mTxtTop, "Total")
If mExchangeTypeID <> 2 Then
   Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 430, mTxtTop, Format(MSHFlexGrid2.TextMatrix(0, 15), "#######0.000"))
End If
Call CheckForNewPage
Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 500, mTxtTop, Format(MSHFlexGrid2.TextMatrix(0, 9), "#######0.000"))
Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)

'mTxtTop = mTxtTop + mRowHeight
mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
mStrLength = Len("Truck-wise Withdrawal Details")
mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
mTxtLeft = mPrintPoint '- (mStrLength) '* 5) '270
mFontSize = 7
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mLeft, mTxtTop - 2, "Truck-wise Withdrawal Details")

''objPDF.PDFDrawLine mGlobalLeft, mTxtTop + mRowHeight, mGlobalLeft + 565, mTxtTop + mRowHeight
''objPDF.PDFDrawLine mGlobalLeft + 50, mTop, mGlobalLeft + 50, mTop + ((MSHFlexGrid3.Rows - 1) * mRowHeight)
''objPDF.PDFDrawLine (mGlobalLeft + 120), mTop, (mGlobalLeft + 120), mTop + ((MSHFlexGrid3.Rows - 1) * mRowHeight)
''objPDF.PDFDrawLine (mGlobalLeft + 190), mTop, (mGlobalLeft + 190), mTop + ((MSHFlexGrid3.Rows + 6) * mRowHeight)
''objPDF.PDFDrawLine (mGlobalLeft + 256), mTop, (mGlobalLeft + 256), mTop + ((MSHFlexGrid3.Rows + 6) * mRowHeight)
''objPDF.PDFDrawLine (mGlobalLeft + 356), mTop, (mGlobalLeft + 356), mTop + ((MSHFlexGrid3.Rows - 1) * mRowHeight)
''
''Dim mTempTop As Integer
''mTempTop = mTxtTop - 3
''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 90, mTempTop + (MSHFlexGrid3.Rows * mRowHeight), "Total")
''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 60, mTempTop + ((MSHFlexGrid3.Rows + 1) * mRowHeight), "Less Gunny Weight(MT)")
''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 50, mTempTop + ((MSHFlexGrid3.Rows + 2) * mRowHeight), "Less Standard Deductions(MT)")
''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 60, mTempTop + ((MSHFlexGrid3.Rows + 3) * mRowHeight), "Add Adjustments(MT)")
''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 60, mTempTop + ((MSHFlexGrid3.Rows + 4) * mRowHeight), "Electronic Weight(MT)")
'''pp 'Hologram No
''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 270, mTempTop + ((MSHFlexGrid3.Rows + 4) * mRowHeight), "Hologram No")
''
'''Actual Values
''Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 200, mTempTop + (MSHFlexGrid3.Rows * mRowHeight), MSHFlexGrid3.TextMatrix(0, 5))
''Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 200, mTempTop + ((MSHFlexGrid3.Rows + 1) * mRowHeight), MSHFlexGrid4.TextMatrix(0, 8))
''Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 200, mTempTop + ((MSHFlexGrid3.Rows + 2) * mRowHeight), MSHFlexGrid4.TextMatrix(0, 9))
''Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 200, mTempTop + ((MSHFlexGrid3.Rows + 3) * mRowHeight), MSHFlexGrid4.TextMatrix(0, 6))
''Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 200, mTempTop + ((MSHFlexGrid3.Rows + 4) * mRowHeight), MSHFlexGrid4.TextMatrix(0, 10))
'''Hologram No -- Use Record Set
''
''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 370, mTempTop + ((MSHFlexGrid3.Rows + 4) * mRowHeight), " " & RsTxn_CxTF_Details!HologramNo & " ")
''
''Dim mTemp As Integer
''For t = 1 To MSHFlexGrid3.Rows - 1
''    If t = 1 Then
''       mTemp = mFontSize
''    Else
''       mTemp = mDispTextSize
''    End If
''    mTxtTop = mTxtTop + mRowHeight
''    Call PrintText(mTemp, FONT_NORMAL, ALIGN_LEFT, mLeft, mTxtTop, MSHFlexGrid3.TextMatrix(t, 1))
''    Call PrintText(mTemp, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 60, mTxtTop, MSHFlexGrid3.TextMatrix(t, 3))
''    Call PrintText(mTemp, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 130, mTxtTop, MSHFlexGrid3.TextMatrix(t, 4))
''    Call PrintText(mTemp, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 200, mTxtTop, MSHFlexGrid3.TextMatrix(t, 5))
''    Call PrintText(mTemp, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 256, mTxtTop, MSHFlexGrid3.TextMatrix(t, 2))
''    Call PrintText(mTemp, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 356, mTxtTop, MSHFlexGrid3.TextMatrix(t, 6))
''Next
''
''Call DrawTable(MSHFlexGrid3.Rows + 6, , mTop, mHeight, mLeft, mRowHeight)
'''Clear Line
''mTxtTop = mHeight - (mTxtFromTop)
''objPDF.PDFSetTextColor = vbWhite
''
''Call PrintText(15, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 255.2, mTxtTop - (6 * mRowHeight), "_____________________________________")
''Call PrintText(15, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 255.2, mTxtTop - (5 * mRowHeight), "_____________________________________")
''Call PrintText(15, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 255.2, mTxtTop - (4 * mRowHeight), "_____________________________________")
''
''objPDF.PDFSetTextColor = vbBlack
''Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 256, mTxtTop - (3 * mRowHeight), "Affix Hologram Here")
'''--
''
''objPDF.PDFDrawLine (mGlobalLeft + 50), mTop - (2 * mRowHeight), (mGlobalLeft + 50), mTop
''objPDF.PDFDrawLine (mGlobalLeft + 120), mTop - (2 * mRowHeight), (mGlobalLeft + 120), mTop
''
''objPDF.PDFDrawLine (mGlobalLeft + 360), mTop - (3 * mRowHeight), (mGlobalLeft + 360), mTop
''objPDF.PDFDrawLine (mGlobalLeft + 420), mTop - (2 * mRowHeight), (mGlobalLeft + 420), mTop
''
''If mHeight > objPDF.PDFGetPageHeight - 50 Then
''   objPDF.PDFEndPage
''   objPDF.PDFNewPage
''   objPDF.PDFSetFont FONT_ARIAL, 12, FONT_BOLD
''   objPDF.PDFSetDrawColor = vbWhite
''   objPDF.PDFSetTextColor = vbBlack
''   objPDF.PDFSetAlignement = ALIGN_CENTER
''   objPDF.PDFSetBorder = BORDER_ALL
''   objPDF.PDFSetFill = True
''
''   objPDF.PDFCell " ", 15, 15, _
''   objPDF.PDFGetPageWidth - 30, objPDF.PDFGetPageHeight - 20
''   mLeft = mGlobalLeft
''   mTop = 15
''   mHeight = mTop + mRowHeight
''   mTxtFromTop = mRowHeight + 2
''End If
''
''Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
''mTxtTop = mHeight - (mTxtFromTop)
''
''Call PrintText(5.4, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "NOTE: THIS DOCUMENT IS NOT VALID WITHOUT A HOLOGRAM OF NCMSL AFFIXED ABOVE & THE SERIAL NUMBER IN THE HOLOGRAM SHOULD MATCH WITH THE HOLOGRAM NUMBER PRINTED ABOVE.")
''
''mHeight = mHeight + 50
''
''Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
''mTxtTop = mHeight - (mTxtFromTop) '+ 15
''
''Call PrintText(9, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "Name & Signature of CMP Manager/CMP DRC:")
''
''mHeight = mHeight + 30
''Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
''mTxtTop = mHeight - (mTxtFromTop)
''
''Call PrintText(6.5, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop - 17, "I/We hereby declare that the commodities being deposited herein belong to the clients as stated above and are free from any encumbrances. I/We request  you to take custody of the said")
''Call PrintText(6.5, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop - 10, "commodities and I/We undertake to indemnify you against any losses incurred by you due to any false information provided by me/us. I/We also undertake to pay all required charges as per")
''Call PrintText(6.5, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop - 3, "your charges schedule from time to time. I have gone throught the attached 'Terms and Conditions'** of storage and agree to abide by the same.")
''
'''Call CheckForNewPage
'''objPDF.PDFDrawLine mGlobalLeft + 50, mTop, mGlobalLeft + 50, mTop + ((MSHFlexGrid3.Rows - 1) * mRowHeight)
'''objPDF.PDFDrawLine (mGlobalLeft + 120), mTop, (mGlobalLeft + 120), mTop + ((MSHFlexGrid3.Rows - 1) * mRowHeight)
'''objPDF.PDFDrawLine (mGlobalLeft + 190), mTop, (mGlobalLeft + 190), mTop + ((MSHFlexGrid3.Rows) * mRowHeight)
'''objPDF.PDFDrawLine (mGlobalLeft + 256), mTop, (mGlobalLeft + 256), mTop + ((MSHFlexGrid3.Rows + 6) * mRowHeight)
'''objPDF.PDFDrawLine (mGlobalLeft + 356), mTop, (mGlobalLeft + 356), mTop + ((MSHFlexGrid3.Rows - 1) * mRowHeight)
'''
'''
''''Dim mTempTop As Integer
''''mTempTop = mTxtTop - 3
''''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 90, mTempTop + (MSHFlexGrid3.Rows * mRowHeight), "Total")
'''
''''----Commented on 18-01-2011 by Sagar
'''
''''''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 60, mTempTop + ((MSHFlexGrid3.Rows + 1) * mRowHeight), "Less Gunny Weight(MT)")
''''''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 50, mTempTop + ((MSHFlexGrid3.Rows + 2) * mRowHeight), "Less Standard Deductions(MT)")
''''''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 60, mTempTop + ((MSHFlexGrid3.Rows + 3) * mRowHeight), "Add Adjustments(MT)")
''''''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 60, mTempTop + ((MSHFlexGrid3.Rows + 4) * mRowHeight), "Electronic Weight(MT)")
'''
'''objPDF.PDFDrawLine (mGlobalLeft), mTempTop + (MSHFlexGrid3.Rows + 4 * mRowHeight), (mGlobalLeft + 256), mTempTop + (MSHFlexGrid3.Rows + 4 * mRowHeight)
'''
''''pp 'Hologram No
'''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 270, mTempTop + ((MSHFlexGrid3.Rows + 4) * mRowHeight), "Hologram No")
'''
''''Actual Values
'''Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 200, mTempTop + (MSHFlexGrid3.Rows * mRowHeight), MSHFlexGrid3.TextMatrix(0, 5))
'''
''''----Commented on 18-01-2011 by Sagar
'''
''''''Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 200, mTempTop + ((MSHFlexGrid3.Rows + 1) * mRowHeight), MSHFlexGrid4.TextMatrix(0, 8))
''''''Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 200, mTempTop + ((MSHFlexGrid3.Rows + 2) * mRowHeight), MSHFlexGrid4.TextMatrix(0, 9))
''''''Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 200, mTempTop + ((MSHFlexGrid3.Rows + 3) * mRowHeight), MSHFlexGrid4.TextMatrix(0, 6))
''''''Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 200, mTempTop + ((MSHFlexGrid3.Rows + 4) * mRowHeight), MSHFlexGrid4.TextMatrix(0, 10))
''''Hologram No -- Use Record Set
'''
'''Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 370, mTempTop + ((MSHFlexGrid3.Rows + 4) * mRowHeight), " " & RsTxn_CxTF_Details!HologramNo & " ")
'''
'mTxtTop = mTxtTop + mRowHeight
mTop = mTop - mRowHeight
mHeight = mHeight - mRowHeight
For T = 1 To MSHFlexGrid3.Rows - 1
    If T = 1 Then
       mTemp = mFontSize
    Else
       mTemp = mDispTextSize
    End If
    
    If mHeight > objPDF.PDFGetPageHeight - 50 Then
       Call CheckForNewPage
       mTop = 15
       mTxtTop = 10
       
       
      'objPDF.PDFDrawPolygon Array(12, 15, objPDF.PDFGetPageWidth - 18, 15, objPDF.PDFGetPageWidth - 18, 30, 12, 30)
      mLeft = mGlobalLeft
      mTop = 15
      mHeight = mTop + mRowHeight
      mTxtFromTop = mRowHeight + 2
      
      'Call DrawTable(10, 2, 15, 30, 12)
      
      Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
      
      
      mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
      mStrLength = Len("NCMSL " & TxtCMPID & "")
      mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
      mTxtLeft = mPrintPoint - (mStrLength * 2) '* 5) '270
      mFontSize = 12
      
      Call PrintText(mFontSize, FONT_NORMAL, ALIGN_CENTER, mTxtLeft, mTxtTop, "NCMSL " & TxtCMPID.Text & "")
      
      mHeight = mTop + mRowHeight
      mLeft = mGlobalLeft
      
      
      Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
      
      mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
      mStrLength = Len("Commodity Deposit Transaction Form (CDTF)")
      mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
      mTxtLeft = mPrintPoint - (mStrLength * 2) '* 5) '270
      mFontSize = 10
      
      Call PrintText(mFontSize, FONT_NORMAL, ALIGN_CENTER, mTxtLeft, mTxtTop, "Commodity Issue Transaction Form (CITF)")
      
      mHeight = mTop + mRowHeight
      mLeft = mGlobalLeft
      
      Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
      
      tmp1 = mTop
      mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
      mStrLength = Len("(This is NOT an instrument of title of the goods)")
      mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
      mTxtLeft = mPrintPoint '- (mStrLength) '* 5) '270
      mFontSize = 6
      
      Call PrintText(mFontSize, FONT_NORMAL, ALIGN_CENTER, mTxtLeft - 30, mTxtTop - 3, "(This is NOT an instrument of title of the goods)")
      
      '--Code added By Kishor
      'Row 1
      
      mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
      mStrLength = Len("Depositor Details")
      mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
      mTxtLeft = mPrintPoint '- (mStrLength) '* 5) '270
      mFontSize = 7
      'Dim mDispTextSize  As Integer
      mDispTextSize = 6
      
      mTxtTop = mTxtTop + mRowHeight
      mTop = mTop + mRowHeight
      Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop, "CITF No")
      Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop, TxtCxTFNo)
      '--
      Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 350, mTxtTop, "Date")
      Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 450, mTxtTop, Format(TxtCxTFDate, "dd-MMM-yyyy"))
      
      mTxtTop = mTxtTop + mRowHeight
      mTop = mTop + mRowHeight
      objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10
      
      Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop, "Location")
      Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop, GetLocationFromCMP(mCMPID))
      '--
      Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 350, mTxtTop, "State")
      Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 450, mTxtTop, GetStateFromLocation(GetLocationIDFromCMP(mCMPID)))
      '--
      mTxtTop = mTxtTop + mRowHeight
      mTop = mTop + mRowHeight
      objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10

     
      Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop, "Transaction Type")
      Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop, CmbExchangeTypeID)
      
      Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 350, mTxtTop, "Commodity")
      Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 450, mTxtTop, TxtCommodityID)
      
      mTxtTop = mTxtTop + mRowHeight
    '  mTop = mTop + mRowHeight
      objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10

      objPDF.PDFDrawLine mGlobalLeft + 345, Val(tmp1), mGlobalLeft + 345, mTxtTop - 10
      objPDF.PDFDrawLine mGlobalLeft + 445, Val(tmp1), mGlobalLeft + 445, mTxtTop - 10
      objPDF.PDFDrawLine mGlobalLeft + 135, Val(tmp1), mGlobalLeft + 135, mTxtTop - 10
      
      mTxtTop = mTxtTop + (mRowHeight * 2)
      mTop = mTop + (mRowHeight * 2) + 3
      
      objPDF.PDFDrawLine mGlobalLeft, mTxtTop - 10, (mGlobalLeft + 565), mTxtTop - 10
      
      Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mLeft, mTxtTop, MSHFlexGrid3.TextMatrix(1, 1))
      Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 60, mTxtTop, MSHFlexGrid3.TextMatrix(1, 3))
      Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 130, mTxtTop, MSHFlexGrid3.TextMatrix(1, 4))
      Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 200, mTxtTop, MSHFlexGrid3.TextMatrix(1, 5))
      Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 256, mTxtTop, MSHFlexGrid3.TextMatrix(1, 2))
      Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 356, mTxtTop, MSHFlexGrid3.TextMatrix(1, 6))
      
      objPDF.PDFDrawLine mGlobalLeft + 50, mTop, mGlobalLeft + 50, mTop + mRowHeight
      objPDF.PDFDrawLine (mGlobalLeft + 120), mTop, (mGlobalLeft + 120), mTop + mRowHeight
      objPDF.PDFDrawLine (mGlobalLeft + 190), mTop, (mGlobalLeft + 190), mTop + mRowHeight
      objPDF.PDFDrawLine (mGlobalLeft + 256), mTop, (mGlobalLeft + 256), mTop + mRowHeight
      objPDF.PDFDrawLine (mGlobalLeft + 356), mTop, (mGlobalLeft + 356), mTop + mRowHeight
      objPDF.PDFDrawLine mGlobalLeft, mTop + mRowHeight, mGlobalLeft + 565, mTop + mRowHeight
   
     ' mTop = mTop + mRowHeight
      'mTxtTop = mTxtTop + mRowHeight
      
      mHeight = mTxtTop
      'mTxtTop = mTxtTop + mRowHeight
      'mTop = mTxtTop
    End If
 
        
    mTxtTop = mTxtTop + mRowHeight
    mTop = mTop + mRowHeight
    mHeight = mHeight + mRowHeight
    
    Call PrintText(mTemp, FONT_NORMAL, ALIGN_LEFT, mLeft, mTxtTop, MSHFlexGrid3.TextMatrix(T, 1))
    Call PrintText(mTemp, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 60, mTxtTop, MSHFlexGrid3.TextMatrix(T, 3))
    Call PrintText(mTemp, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 130, mTxtTop, MSHFlexGrid3.TextMatrix(T, 4))
    Call PrintText(mTemp, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 200, mTxtTop, MSHFlexGrid3.TextMatrix(T, 5))
    Call PrintText(mTemp, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 256, mTxtTop, MSHFlexGrid3.TextMatrix(T, 2))
    Call PrintText(mTemp, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 356, mTxtTop, MSHFlexGrid3.TextMatrix(T, 6))
    
    objPDF.PDFDrawLine mGlobalLeft + 50, mTop, mGlobalLeft + 50, mTop + mRowHeight
    objPDF.PDFDrawLine (mGlobalLeft + 120), mTop, (mGlobalLeft + 120), mTop + mRowHeight
    objPDF.PDFDrawLine (mGlobalLeft + 190), mTop, (mGlobalLeft + 190), mTop + mRowHeight
    objPDF.PDFDrawLine (mGlobalLeft + 256), mTop, (mGlobalLeft + 256), mTop + mRowHeight
    objPDF.PDFDrawLine (mGlobalLeft + 356), mTop, (mGlobalLeft + 356), mTop + mRowHeight
    objPDF.PDFDrawLine mGlobalLeft, mTop + mRowHeight, mGlobalLeft + 565, mTop + mRowHeight


Next
'Call CheckForNewPage
'Call DrawTable(MSHFlexGrid3.Rows - 1, , mTop, mHeight, mLeft, mRowHeight)
Call CheckForNewPage
Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 90, mTop + mRowHeight - 3, "Total")
Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 200, mTop + mRowHeight - 3, MSHFlexGrid3.TextMatrix(0, 5))
objPDF.PDFDrawLine (mGlobalLeft + 190), mTop, (mGlobalLeft + 190), mTop + mRowHeight
objPDF.PDFDrawLine (mGlobalLeft + 256), mTop, (mGlobalLeft + 256), mTop + mRowHeight * 4

Call CheckForNewPage
objPDF.PDFDrawLine (mGlobalLeft), mTop + mRowHeight, (mGlobalLeft + 256), mTop + mRowHeight

mTop = mTop + (mRowHeight * 4)
objPDF.PDFDrawLine mGlobalLeft + 50, mTop + mRowHeight, mGlobalLeft + 50, mTop + (mRowHeight * 3)
objPDF.PDFDrawLine (mGlobalLeft + 120), mTop + mRowHeight, (mGlobalLeft + 120), mTop + (mRowHeight * 3)
objPDF.PDFDrawLine (mGlobalLeft + 190), mTop + mRowHeight, (mGlobalLeft + 190), mTop + (mRowHeight * 3)
objPDF.PDFDrawLine (mGlobalLeft + 256), mTop, (mGlobalLeft + 256), mTop + (mRowHeight * 3)
objPDF.PDFDrawLine (mGlobalLeft + 356), mTop, (mGlobalLeft + 356), mTop + (mRowHeight * 3)
objPDF.PDFDrawLine (mGlobalLeft + 420), mTop + mRowHeight, (mGlobalLeft + 420), mTop + (mRowHeight * 3)
Call CheckForNewPage
Call PrintText(6, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 265, mTop - 3, "Affix Hologram Here")

If mExchangeTypeID = 2 Then
   Call CheckForNewPage
   Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 70, mTop + mRowHeight - 3, "Total No. of Bags")
   Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 200, mTop + mRowHeight - 3, Format(MSHFlexGrid4.TextMatrix(0, 6)))
   objPDF.PDFDrawLine (mGlobalLeft + 190), mTop, (mGlobalLeft + 190), mTop + mRowHeight
End If

Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 270, mTop + mRowHeight - 3, "Hologram No")
Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 370, mTop + mRowHeight - 3, " " & RsTxn_CxTF_Details!HologramNo & " ")
Call CheckForNewPage
Call DrawTable(4, , mTop, mTop, mLeft, mRowHeight)

'If mHeight > objPDF.PDFGetPageHeight - 50 Then
'   objPDF.PDFEndPage
'   objPDF.PDFNewPage
'   objPDF.PDFSetFont FONT_ARIAL, 12, FONT_BOLD
'   objPDF.PDFSetDrawColor = vbWhite
'   objPDF.PDFSetTextColor = vbBlack
'   objPDF.PDFSetAlignement = ALIGN_CENTER
'   objPDF.PDFSetBorder = BORDER_ALL
'   objPDF.PDFSetFill = True
'
'   objPDF.PDFCell " ", 15, 15, _
'   objPDF.PDFGetPageWidth - 30, objPDF.PDFGetPageHeight - 20
'   mLeft = mGlobalLeft
'   mTop = 15
'   mHeight = mTop + mRowHeight
'   mTxtFromTop = mRowHeight + 2
'End If

mHeight = mTop
Call CheckForNewPage
Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
Call CheckForNewPage
mTxtTop = mHeight - (mTxtFromTop)
'''

Call PrintText(5.4, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "NOTE: THIS DOCUMENT IS NOT VALID WITHOUT A HOLOGRAM OF NCMSL AFFIXED ABOVE & THE SERIAL NUMBER IN THE HOLOGRAM SHOULD MATCH WITH THE HOLOGRAM NUMBER PRINTED ABOVE.")
'''
mHeight = mHeight + 50
'''
Call CheckForNewPage
Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)

'''
Call CheckForNewPage
mTxtTop = mHeight - (mTxtFromTop) '+ 15
Call PrintText(9, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "Name & Signature of CMP Manager/CMP DRC:")
Call CheckForNewPage
Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
Call CheckForNewPage
mTxtTop = mTxtTop + mRowHeight
Call PrintText(7, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "Received the above mentioned goods in good condition.")
'mHeight = mHeight + 30
'Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)

'''
''Call PrintText(6.5, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop - 17, "I/We hereby declare that the commodities being deposited herein belong to the clients as stated above and are free from any encumbrances. I/We request  you to take custody of the said")
''Call PrintText(6.5, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop - 10, "commodities and I/We undertake to indemnify you against any losses incurred by you due to any false information provided by me/us. I/We also undertake to pay all required charges as per")
''Call PrintText(6.5, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop - 3, "your charges schedule from time to time. I have gone throught the attached 'Terms and Conditions'** of storage and agree to abide by the same.")
'''''

mHeight = mHeight + 70

Call CheckForNewPage

Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
Call CheckForNewPage
mTxtTop = mHeight - (mTxtFromTop)
Call PrintText(9, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "Name & Signature of Withdrawer/Agent of Withdrawer:")

mHeight = mHeight + 50

Call CheckForNewPage
mTxtTop = mHeight - (mTxtFromTop)
Call PrintText(8, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop, "**'Terms and Conditions' can be viewed in http://www.ncmsl.com/whrtNc.html")

objPDF.PDFEndDoc

GetTotalPDFPages = objPDF.PDFPageNumber

End Function

Private Sub Grid_Head_RRNDetail()

MSHFlexGrid10.Clear
MSHFlexGrid10.Rows = 3
MSHFlexGrid10.FixedRows = 2
MSHFlexGrid10.Cols = 7

MSHFlexGrid10.TextMatrix(1, 0) = "RRNID"
MSHFlexGrid10.TextMatrix(1, 1) = "ISINID"
MSHFlexGrid10.TextMatrix(1, 2) = "Balance Remat Qty"
MSHFlexGrid10.TextMatrix(1, 3) = "Withdrawal Qty"
MSHFlexGrid10.TextMatrix(1, 4) = "RRNISINID"
MSHFlexGrid10.TextMatrix(1, 5) = "Withdrawal Bags"
MSHFlexGrid10.TextMatrix(1, 6) = "RRN No"

MSHFlexGrid10.ColWidth(0) = 1200
MSHFlexGrid10.ColWidth(1) = 1400
MSHFlexGrid10.ColWidth(2) = 1400
MSHFlexGrid10.ColWidth(3) = 1400
MSHFlexGrid10.ColWidth(4) = 1400

MSHFlexGrid10.RowHeight(1) = 600

End Sub

Private Sub FillRRNDetails(mRRNID_ As Variant)
tmp = "Select MRD.RRNID,MRD.ISINID, MG.BalanceWeight,MR.ISINID as RRNISINID,MR.RRNNo"
tmp = tmp & " from Mst_RRN_LotDetail MRD"
tmp = tmp & " Inner join MSt_GSL MG on MRD.ISINID=MG.ISINID"
tmp = tmp & " Inner join Mst_RRN MR on MRD.RRNID=MR.RRNID"
tmp = tmp & " Where MRD.RRNID in (" & mRRNID_ & ")"
Call Tmp5Table

'MSHFlexGrid10.Rows = MSHFlexGrid10.Rows + TmpRs5.RecordCount

If TmpRs5.RecordCount = 0 Then Exit Sub

For Ln = 1 To TmpRs5.RecordCount
    I = MSHFlexGrid10.Rows - 1
    MSHFlexGrid10.TextMatrix(I, 0) = IIf(IsNull(TmpRs5!RRNID), "", TmpRs5!RRNID)
    MSHFlexGrid10.TextMatrix(I, 1) = IIf(IsNull(TmpRs5!ISINID), "", TmpRs5!ISINID)
    MSHFlexGrid10.TextMatrix(I, 2) = IIf(IsNull(TmpRs5!BalanceWeight), "", TmpRs5!BalanceWeight)
    MSHFlexGrid10.TextMatrix(I, 4) = IIf(IsNull(TmpRs5!RRNISINID), "", TmpRs5!RRNISINID)
    MSHFlexGrid10.TextMatrix(I, 6) = IIf(IsNull(TmpRs5!RRNNo), "", TmpRs5!RRNNo)
    
    MSHFlexGrid10.TextMatrix(0, 2) = Val(MSHFlexGrid10.TextMatrix(0, 2)) + Val(MSHFlexGrid10.TextMatrix(I, 2))
    TmpRs5.MoveNext
    MSHFlexGrid10.Rows = MSHFlexGrid10.Rows + 1
Next

End Sub

Function CheckISINinRRNDetailTab() As Boolean
Dim Retval As Boolean
Dim TxtGSLISINno_ As Variant
Retval = True
If Check2.Value = 0 Then
   TxtGSLISINno_ = CmbGCTDClient.Text ' TxtGSLISINno
Else
   TxtGSLISINno_ = CmbGCTDClient.Text
End If

If MSHFlexGrid10.Rows > 3 Then
   For u = 2 To MSHFlexGrid10.Rows - 1
       If MSHFlexGrid10.TextMatrix(u, 1) <> "" Then
          'If LCase(TxtGSLISINno_) = LCase(MSHFlexGrid10.TextMatrix(u, 1)) & " ~ " & LCase(MSHFlexGrid10.TextMatrix(u, 4)) Then
           If mExchangeTypeID = 2 Then
               If LCase(TxtGSLISINno) = LCase(MSHFlexGrid10.TextMatrix(u, 1)) Then
                  'mRRNinGCTD = MSHFlexGrid2.TextMatrix(u, 1)
                  Retval = False
                  Exit For
               End If
           Else
              Retval = True
           End If
       End If
   Next
End If

CheckISINinRRNDetailTab = Retval

End Function

Function UpdateRRNDetailTab(GSLISINno_ As Variant, wt_ As Variant, bags_ As Variant, Flag_ As Variant) As String
Dim Retval As String
Dim TxtGSLISINno_ As Variant
Retval = True
If Check2.Value = 0 Then
   TxtGSLISINno_ = CmbGCTDClient.Text ' TxtGSLISINno
Else
   TxtGSLISINno_ = CmbGCTDClient.Text
End If

If MSHFlexGrid10.Rows > 3 Then
   For u = 2 To MSHFlexGrid10.Rows - 1
       If MSHFlexGrid10.TextMatrix(u, 1) <> "" Then
          If mExchangeTypeID = 2 Then
             If LCase(GSLISINno_) = LCase(MSHFlexGrid10.TextMatrix(u, 6)) & " ~ " & LCase(MSHFlexGrid10.TextMatrix(u, 4)) Then
                If Flag_ = "A" Then
                   MSHFlexGrid10.TextMatrix(u, 3) = Val(MSHFlexGrid10.TextMatrix(u, 3)) + Val(wt_)
                   MSHFlexGrid10.TextMatrix(u, 5) = Val(MSHFlexGrid10.TextMatrix(u, 5)) + Val(bags_)
                Else
                   MSHFlexGrid10.TextMatrix(u, 3) = Val(MSHFlexGrid10.TextMatrix(u, 3)) - Val(wt_)
                   MSHFlexGrid10.TextMatrix(u, 5) = Val(MSHFlexGrid10.TextMatrix(u, 5)) - Val(bags_)
                End If
                Retval = MSHFlexGrid10.TextMatrix(u, 4)
                Exit For
             End If
          End If
       End If
   Next
End If

UpdateRRNDetailTab = Retval

End Function


Function CheckWtRRNDetailTab() As Boolean
Dim Retval As Boolean
Retval = True

If MSHFlexGrid10.Rows > 3 Then
   For u = 2 To MSHFlexGrid10.Rows - 1
       If MSHFlexGrid10.TextMatrix(u, 1) <> "" Then
          If Val(MSHFlexGrid10.TextMatrix(u, 3)) > Val(MSHFlexGrid10.TextMatrix(u, 2)) Then
             Retval = False
             Exit For
          End If
       End If
   Next
End If

CheckWtRRNDetailTab = Retval

End Function

Private Sub DeleteFromRRNDetail(mRRNID_ As Variant)
Dim mDelRows As Variant
If I > 0 Then
   If mRRNID_ <> "" Then
      mDelRows = 0
Restart:
        For r = 2 To MSHFlexGrid10.Rows - 1
            If MSHFlexGrid10.TextMatrix(r, 0) = mRRNID_ Then
               MSHFlexGrid10.TextMatrix(0, 2) = Val(MSHFlexGrid10.TextMatrix(0, 2)) + Val(MSHFlexGrid10.TextMatrix(r, 2))
               For RO = r To MSHFlexGrid10.Rows - 2
                   For C = 0 To MSHFlexGrid10.Cols - 1
                       MSHFlexGrid10.TextMatrix(RO, C) = MSHFlexGrid10.TextMatrix(RO + 1, C)
                   Next
               Next
               MSHFlexGrid10.Rows = MSHFlexGrid10.Rows - 1
               GoTo Restart
            End If
        Next
   End If
End If

End Sub

Function GetHeaderISINRRN(GSLISINno_ As Variant) As String
Dim Retval As String

If MSHFlexGrid10.Rows > 3 Then
   For u = 2 To MSHFlexGrid10.Rows - 1
       If MSHFlexGrid10.TextMatrix(u, 1) <> "" Then
          If mExchangeTypeID = 2 Then
             If LCase(GSLISINno_) = LCase(MSHFlexGrid10.TextMatrix(u, 1)) Then
                Retval = MSHFlexGrid10.TextMatrix(u, 6) & " ~ " & MSHFlexGrid10.TextMatrix(u, 4)
                Exit For
             End If
          End If
       End If
   Next
End If

GetHeaderISINRRN = Retval

End Function

Private Sub CreateRRNNCDEXInvoice()

If TxtCxTFDate.Value < CDate("10/01/2013") Then Exit Sub
       
Call TableTxn_CDTFBOE(" Where SM_Prefix='" & TxtPreFix & "' And CxTF_DetailsID=" & Val(TxtCxTFNo) & " And TxnType='R'")

If RsTxn_CDTFBOE.RecordCount > 0 Then Exit Sub

tmp = "Select CommodityGroupID From Mst_Commodity Where CommodityID=" & mCommodityID
Call TmpTable
If TmpRs.RecordCount > 0 Then
   mCommodityGroupId = TmpRs!CommodityGroupID
End If

mLocationID = GetLocationIDFromCMP(Val(mCMPID))

tmp = "Select  Max(Rate) as MaxRate From MSt_NcDEXStorageRateCardDetail Where NSRCID in("
tmp = tmp & " Select NSRCID From MSt_NcDEXStorageRateCard Where LocationID = " & mLocationID & " And CommodityGroupID =" & mCommodityGroupId & ")" ' And GodownID in (Select GodownID from MSt_Godown Where CMPID=" & mCMPID & ")"
Call Tmp3Table
If Not TmpRs3.EOF Then
   mMaxRate = TmpRs3!MaxRate
Else
   mRateExist = False
End If

Call TableTxn_CDTFBOE_GSL(" Where SM_Prefix='" & TxtPreFix & "'")

tmp = "Select MR.RRNID,MR.RRNNo,MR.RRNDate,isnull(ISIN.ISINID,ISIN1.ISINID)ISINID,isnull(ISIN.FED,ISIN1.FED)FED,isnull(ISIN.NE_Date,ISIN1.NE_Date)NE_Date, isnull(ISIN.Grade,ISIN1.Grade)Grade,MR.Flag, MR.CommTrackFlag,sum(TCG.DematWeight)DematWeight,sum(TCG.NoOfBags)NoOfBags,MR.ClientID,TCG.Mst_GSL_ID"
tmp = tmp & " from Txn_CxTF_GSL TCG"
tmp = tmp & " inner join mst_RRN MR on MR.RRNID=TCG.WHRNo"
tmp = tmp & " inner join mst_GSL GSL on GSL.Mst_GSL_ID=TCG.Mst_GSL_ID and GSL.SM_prefix=TCG.SM_prefix"
tmp = tmp & " left Join Mst_ISIN as ISIN on ISIN.ISINID = GSl.ISINID"
tmp = tmp & " left Join Mst_ISIN as ISIN1 on ISIN1.ISINID = MR.ISINID"
tmp = tmp & " Where TCG.CxTF_DetailsID=" & Val(TxtCxTFNo) & "And TCG.SM_prefix='" & TxtPreFix & "' And TCG.TxnType='W'"
tmp = tmp & " Group by MR.RRNID,MR.RRNNo,MR.RRNDate,ISIN.ISINID,ISIN.FED, ISIN.NE_Date, ISIN.Grade,MR.Flag, MR.CommTrackFlag,MR.ClientID,TCG.Mst_GSL_ID,ISIN1.ISINID,ISIN1.FED, ISIN1.NE_Date, ISIN1.Grade"
Call Tmp2Table

Call Grid_HeadRRNNCDEXInvoice
mSpaceChargesForInvoice = 0
If TmpRs2.RecordCount > 0 Then
   
   mMonthLimit = Gen_ReadNFinMailIDINI("NDCEXSTORAGEMonths")
   mStorageRate1 = Gen_ReadNFinMailIDINI("NDCEXSTORAGERate1")
   mStorageRate2 = Gen_ReadNFinMailIDINI("NDCEXSTORAGERate2")
   mSpaceRate1 = Gen_ReadNFinMailIDINI("NCDEXSpaceCharges1")
   mSpaceRate2 = Gen_ReadNFinMailIDINI("NCDEXSpaceCharges2")
   MSHFlexGrid11.Rows = MSHFlexGrid11.Rows + TmpRs2.RecordCount
   mInvAmountOnPrint = 0
   For I = 1 To TmpRs2.RecordCount
       
       mFEDCITF = 0
        mNewRRNDate = TmpRs2!RRNDate + 1
        If TmpRs2!RRNDate < CDate("10/01/2013") Then
           If TmpRs2!CommTrackFlag = "1" Then
              mNewRRNDate = CDate("10/01/2013")
           End If
        End If
        If TmpRs2!CommTrackFlag = "1" Then
           mLastFEDDate = GetLastDateOfMonth(CDate(TmpRs2!FED))
           If TxtCxTFDate.Value >= CDate(mLastFEDDate) Then
              mLastFEDDate = GetLastDateOfMonth(CDate(TmpRs2!FED))
              
              If CDate(mNewRRNDate) > CDate(mLastFEDDate) Then
'Commented on 20 Oct 2016 after issue of Bikaner Invoice raised and RRN Date is after FED so CP will calculate charges by it self
'                 mTotDays1 = DateDiff("d", CDate(mLastFEDDate) + 1, CDate(mNewRRNDate) - 1) + 1
'                 mRRNFED = mTotDays1 * mMaxRate * Val(TmpRs2!DematWeight)
'                 mNextToFED = CDate(mLastFEDDate) + 1
                 mTotDays = DateDiff("d", CDate(mNewRRNDate), TxtCxTFDate.Value) + 1
                 
                 'mRRNDateAdded = DateAdd("M", mMonthLimit, CDate(mNewRRNDate))
                 mRRNDateAdded = DateAdd("M", mMonthLimit, CDate(TmpRs2!FED))
                 mTotDays = DateDiff("d", CDate(mNewRRNDate), TxtCxTFDate.Value) + 1
                 'mTotDays2 = DateDiff("d", CDate(mNewRRNDate), CDate(mRRNDateAdded)) + 1 'mTotDays = DateDiff("d", CDate(mNextToFED), TxtCxTFDate.Value) + 1
                 mTotDays2 = DateDiff("d", CDate(mNewRRNDate), CDate(mRRNDateAdded)) + 1
''''Commented on 05 Jan 2017 after modification required for Jodhpur issue
'                 If mTotDays < mTotDays2 Then
'                    mFEDCITF = Val(mFEDCITF) + (mTotDays * mMaxRate * Val(TmpRs2!DematWeight)) * Val(mStorageRate1)
'                    mTotDays3 = 0
'                 Else
'                    mFEDCITF = Val(mFEDCITF) + (mTotDays2 * mMaxRate * Val(TmpRs2!DematWeight)) * Val(mStorageRate1)
'                    mRRNDateAdded = CDate(mRRNDateAdded) + 1
'                    mTotDays3 = DateDiff("d", CDate(mRRNDateAdded), TxtCxTFDate.Value) + 1
'                    mFEDCITF = Val(mFEDCITF) + (mTotDays3 * mMaxRate * Val(TmpRs2!DematWeight)) * Val(mStorageRate2)
'                 End If
                 If CDate(mNewRRNDate) <= CDate(TmpRs2!FED) Then
                    If mTotDays < mTotDays2 Then
                       mFEDCITF = Val(mFEDCITF) + (mTotDays * mMaxRate * Val(TmpRs2!DematWeight)) * Val(mStorageRate1)
                       mTotDays3 = 0
                    Else
                       mFEDCITF = Val(mFEDCITF) + (mTotDays2 * mMaxRate * Val(TmpRs2!DematWeight)) * Val(mStorageRate1)
                       mRRNDateAdded = CDate(mRRNDateAdded) + 1
                       mTotDays3 = DateDiff("d", CDate(mRRNDateAdded), TxtCxTFDate.Value) + 1
                       mFEDCITF = Val(mFEDCITF) + (mTotDays3 * mMaxRate * Val(TmpRs2!DematWeight)) * Val(mStorageRate2)
                    End If
                 Else
                    mFEDCITF = Val(mFEDCITF) + (mTotDays * mMaxRate * Val(TmpRs2!DematWeight)) * Val(mStorageRate1) 'Commented on 29 Dec 2017 after mail from nagina Val(mStorageRate2)
                    mTotDays3 = 0
                 End If
                 
                 
                 mAUM = Val(TxtSpotPrice) * Val(TmpRs2!DematWeight)
                 mSpaceIncome = 0
                 mSpaceChargesForInvoice = 0
                 mCITFMonth = Month(TxtCxTFDate)
''''Commented on 05 Jan 2017 after modification required for Jodhpur issue
'                 mSpaceDateDiff = DateDiff("m", CDate(mNewRRNDate), TxtCxTFDate.Value)
'                 For x = 1 To mSpaceDateDiff 'Month(TxtCxTFDate) 'For x = Month(mNextToFED) To Month(TxtCxTFDate)
'                     If x > Val(mMonthLimit) Then
'                        mSpaceIncome = Val(mSpaceIncome) + (Val(mAUM) * (mSpaceRate2 / 100))
'                     Else
'                        mSpaceIncome = Val(mSpaceIncome) + (Val(mAUM) * (mSpaceRate1 / 100))
'                     End If
'                 Next
                 
                 If CDate(mNewRRNDate) <= CDate(TmpRs2!FED) Then
                    mSpaceDateDiff = DateDiff("m", CDate(mNewRRNDate), TxtCxTFDate.Value)
                    For x = 1 To mSpaceDateDiff 'Month(TxtCxTFDate) 'For x = Month(mNextToFED) To Month(TxtCxTFDate)
                        If x > Val(mMonthLimit) Then
                           mSpaceIncome = Val(mSpaceIncome) + (Val(mAUM) * (mSpaceRate2 / 100))
                        Else
                           mSpaceIncome = Val(mSpaceIncome) + (Val(mAUM) * (mSpaceRate1 / 100))
                        End If
                    Next
                 Else
                    mSpaceIncome = Val(mSpaceIncome) + (Val(mAUM) * (mSpaceRate1 / 100)) 'Commented on 29 Dec 2017 after mail from Nagina Val(mSpaceIncome) + (Val(mAUM) * (mSpaceRate2 / 100))
                 End If
                 
                 mSCToBerecovered = Val(mFEDCITF) '+ Val(mSpaceIncome) 'Val(mRRNFED) + Val(mFEDCITF) + Val(mSpaceIncome)
                 mSpaceChargesForInvoice = 0 'Val(mSpaceChargesForInvoice) + Val(mSpaceIncome)
                 mTotDays = mTotDays
                 
'                'Introduced this line for month comparison on 13 Dec 2013 after mail from Mayur Sir and Dharmendra Sir
'                 If DateDiff("m", CDate(TmpRs2!FED), TxtCxTFDate.Value) > 0 Then
'                    If DateDiff("m", CDate(TmpRs2!FED), TxtCxTFDate.Value) > Val(mMonthLimit) Then
'                       mFEDCITF = (mTotDays * mMaxRate * Val(TmpRs2!DematWeight)) * Val(mStorageRate2) '2
'                    Else
'                       mFEDCITF = (mTotDays * mMaxRate * Val(TmpRs2!DematWeight)) * Val(mStorageRate1) '2
'                    End If
'                 Else
'                    mFEDCITF = (mTotDays * mMaxRate * Val(TmpRs2!DematWeight))
'                 End If
'
'                'mFEDCITF = (mTotDays * mMaxRate * Val(TmpRs2!DematWeight)) * 2
'                 mAUM = Val(TxtSpotPrice) * Val(TmpRs2!DematWeight)
'                 If DateDiff("m", CDate(TmpRs2!FED), TxtCxTFDate.Value) > Val(mMonthLimit) Then
'                    mSpaceIncome = Val(mAUM) * (mSpaceRate2 / 100)
'                 Else
'                    mSpaceIncome = Val(mAUM) * (mSpaceRate1 / 100)
'                 End If
'                 mSCToBerecovered = mRRNFED + mFEDCITF + mSpaceIncome
              Else
                 mFEDCITF = 0
                 mTotDays1 = DateDiff("d", CDate(mNewRRNDate), CDate(mLastFEDDate)) + 1 ' RRN Date to FED Date charges
                 mRRNFED = mTotDays1 * mMaxRate * Val(TmpRs2!DematWeight)
                 mNextToFED = CDate(mLastFEDDate) + 1
                 mRRNDateAdded = DateAdd("M", mMonthLimit, CDate(TmpRs2!FED)) ' CDate(mNextToFED))   'EDD + 3 Months
                 mTotDays = DateDiff("d", CDate(mNextToFED), TxtCxTFDate.Value) + 1
                 mTotDays2 = DateDiff("d", CDate(mNextToFED), CDate(mRRNDateAdded)) + 1 'mTotDays = DateDiff("d", CDate(mNextToFED), TxtCxTFDate.Value) + 1
                 If mTotDays < mTotDays2 Then
                    mTotDays2 = DateDiff("d", CDate(mNextToFED), TxtCxTFDate.Value) + 1
                    mFEDCITF = Val(mFEDCITF) + (mTotDays2 * mMaxRate * Val(TmpRs2!DematWeight)) * Val(mStorageRate1)
                 Else
                    mFEDCITF = Val(mFEDCITF) + (mTotDays2 * mMaxRate * Val(TmpRs2!DematWeight)) * Val(mStorageRate1)
                    mRRNDateAdded = CDate(mRRNDateAdded) + 1
                    mTotDays3 = DateDiff("d", CDate(mRRNDateAdded), TxtCxTFDate.Value) + 1
                    mFEDCITF = Val(mFEDCITF) + (mTotDays3 * mMaxRate * Val(TmpRs2!DematWeight)) * Val(mStorageRate2)
                 End If
                 
                 mAUM = Val(TxtSpotPrice) * Val(TmpRs2!DematWeight)
                 mSpaceIncome = 0
                 mSpaceChargesForInvoice = 0
                 mCITFMonth = Month(TxtCxTFDate)
                 
                 mSpaceDateDiff = DateDiff("m", CDate(TmpRs2!FED), TxtCxTFDate.Value) 'DateDiff("m", CDate(TmpRs2!FED), TxtCxTFDate.Value)
                 For x = 1 To mSpaceDateDiff 'Month(TxtCxTFDate) 'For x = Month(mNextToFED) To Month(TxtCxTFDate)
                     If x >= Val(mMonthLimit) Then
                        mSpaceIncome = Val(mSpaceIncome) + (Val(mAUM) * (mSpaceRate2 / 100))
                     Else
                        mSpaceIncome = Val(mSpaceIncome) + (Val(mAUM) * (mSpaceRate1 / 100))
                     End If
                 Next
                 
                 mSCToBerecovered = Val(mRRNFED) + Val(mFEDCITF) '+ Val(mSpaceIncome)
                 mSpaceChargesForInvoice = 0 'Val(mSpaceChargesForInvoice) + Val(mSpaceIncome)
                 
                 
'Commented on 21 July 2016
'''''                'Introduced this line for month comparison on 13 Dec 2013 after mail from Mayur Sir and Dharmendra Sir
'''''                 If DateDiff("m", CDate(TmpRs2!FED), TxtCxTFDate.Value) > 0 Then
''''''                    If DateDiff("m", CDate(TmpRs2!FED), TxtCxTFDate.Value) > Val(mMonthLimit) Then
''''''                       mFEDCITF = (mTotDays * mMaxRate * Val(TmpRs2!DematWeight)) * Val(mStorageRate2) '* 2
''''''                    Else
''''''                       mFEDCITF = (mTotDays * mMaxRate * Val(TmpRs2!DematWeight)) * Val(mStorageRate1) '* 2
''''''                     End If
'''''                    mTotalMonthforLoop = DateDiff("m", CDate(TmpRs2!FED), TxtCxTFDate.Value)
'''''                    mInvCh = 1
'''''                    mFEDCITF = 0
'''''                    For x = 1 To mTotalMonthforLoop 'For x = Month(mNextToFED) To Month(TxtCxTFDate)
'''''                        If mInvCh > Val(mMonthLimit) Then
'''''                           mFEDCITF = Val(mFEDCITF) + (mTotDays * mMaxRate * Val(TmpRs2!DematWeight)) * Val(mStorageRate2)
'''''                        Else
'''''                           mFEDCITF = Val(mFEDCITF) + (mTotDays * mMaxRate * Val(TmpRs2!DematWeight)) * Val(mStorageRate1) '* 2
'''''                        End If
'''''                        mInvCh = Val(mInvCh) + 1
'''''                   Next
'''''                 Else
'''''                   mFEDCITF = (mTotDays * mMaxRate * Val(TmpRs2!DematWeight))
'''''                 End If
'''''                'mFEDCITF = (mTotDays * mMaxRate * Val(TmpRs2!DematWeight)) * 2
                 
                 
                 
'                 If DateDiff("m", CDate(TmpRs2!FED), TxtCxTFDate.Value) > Val(mMonthLimit) Then
'                    'mSpaceIncome = Val(mAUM) * (mSpaceRate2 / 100)
'                    For x = Month(mNextToFED) To Month(TxtCxTFDate)
'                        If x = mCITFMonth Then
'                           mLasTCITFMonth = Day(GetLastDateOfMonth(CDate(TxtCxTFDate)))
'                           mDayCITFMonth = Day(TxtCxTFDate)
'                           mSpaceIncome = Val(mSpaceIncome) + Val((Val(mAUM) * (mSpaceRate2 / 100)) * (mDayCITFMonth / mLasTCITFMonth))
'                        Else
'                           mSpaceIncome = Val(mAUM) * (mSpaceRate2 / 100)
'                        End If
'                    Next
'                 Else
'                    For x = Month(mNextToFED) To Month(TxtCxTFDate)
'                        If x = mCITFMonth Then
'                           mLasTCITFMonth = Day(GetLastDateOfMonth(CDate(TxtCxTFDate)))
'                           mDayCITFMonth = Day(TxtCxTFDate)
'                           mSpaceIncome = Val(mSpaceIncome) + Val((Val(mAUM) * (mSpaceRate1 / 100)) * (mDayCITFMonth / mLasTCITFMonth))
'                        Else
'                           mSpaceIncome = Val(mAUM) * (mSpaceRate1 / 100)
'                        End If
'                    Next
'                    'mSpaceIncome = Val(mAUM) * (mSpaceRate1 / 100)
'
'                 End If
              
                  mTotDays = mTotDays + mTotDays1
              End If
              
           ElseIf TxtCxTFDate.Value <= CDate(mLastFEDDate) Then 'CDate(TmpRs2!FED) Then
           'Single Charges From RRNDate to CITFDate then
              mSpaceChargesForInvoice = 0
              mTotDays = DateDiff("d", CDate(mNewRRNDate), TxtCxTFDate.Value) + 1
              mSCToBerecovered = (mTotDays * mMaxRate * Val(TmpRs2!DematWeight))
           End If
        Else
            If TxtCxTFDate.Value <= CDate(TmpRs2!FED) Then
               
            ElseIf TxtCxTFDate.Value >= CDate(TmpRs2!FED) Then
               tmp = "Select * From Txn_RRNPWR Where RRNID =" & TmpRs2!RRNID
               Call Tmp3Table
               If Not TmpRs3.EOF Then
                  mTotDays1 = DateDiff("d", TmpRs3!PWRDate, CDate(TmpRs2!FED)) + 1
                  mPWRFED = mTotDays1 * mMaxRate * Val(TmpRs2!DematWeight)
                  
                  mTotDays = DateDiff("d", TmpRs3!PWRDate, TxtCxTFDate.Value) + 1
                  
                  If DateDiff("m", TmpRs3!PWRDate, TxtCxTFDate.Value) > 0 Then
                     mFEDCITF = (mTotDays * mMaxRate * Val(TmpRs2!DematWeight)) * 2
                  Else
                     mFEDCITF = (mTotDays * mMaxRate * Val(TmpRs2!DematWeight))
                  End If
                  mSCToBerecovered = mPWRFED + mFEDCITF
                  mTotDays = mTotDays + mTotDays1
               Else
                  mTotDays = DateDiff("d", CDate(TmpRs2!FED), TxtCxTFDate.Value) + 1
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
        MSHFlexGrid11.TextMatrix(r, 5) = Format(TxtCxTFDate.Value, "dd-MMM-yyyy")
        MSHFlexGrid11.TextMatrix(r, 6) = TmpRs2!DematWeight
        MSHFlexGrid11.TextMatrix(r, 7) = TxtCommodityID
        MSHFlexGrid11.TextMatrix(r, 8) = mSCToBerecovered
        MSHFlexGrid11.TextMatrix(r, 9) = mLocationID
        MSHFlexGrid11.TextMatrix(r, 10) = TmpRs2!ClientID
        MSHFlexGrid11.TextMatrix(r, 11) = mCommodityID
        MSHFlexGrid11.TextMatrix(r, 12) = mServiceTaxID
        MSHFlexGrid11.TextMatrix(r, 13) = mMaxRate
        MSHFlexGrid11.TextMatrix(r, 14) = TmpRs2!Mst_GSL_ID
        MSHFlexGrid11.TextMatrix(r, 15) = TmpRs2!NoofBags
        MSHFlexGrid11.TextMatrix(r, 16) = mTotDays
        MSHFlexGrid11.TextMatrix(0, 8) = Val(MSHFlexGrid11.TextMatrix(0, 8)) + Val(MSHFlexGrid11.TextMatrix(r, 8))
        MSHFlexGrid11.TextMatrix(0, 6) = Val(MSHFlexGrid11.TextMatrix(0, 6)) + Val(MSHFlexGrid11.TextMatrix(r, 6))
        
        If RsTxn_CDTFBOE.RecordCount > 0 Then RsTxn_CDTFBOE.MoveFirst
        RsTxn_CDTFBOE.Find "ClientID=" & TmpRs2!ClientID
        If RsTxn_CDTFBOE.EOF = True Then
           mInvAmountOnPrint = 0
           mBOEIDNew = 0
           Call AddCDTFBOE(Val(MSHFlexGrid11.TextMatrix(r, 10)), "R")
        End If
        
        Call SaveCDTFBOEGSL(Val(MSHFlexGrid11.TextMatrix(r, 14)), Val(MSHFlexGrid11.TextMatrix(r, 15)), Val(MSHFlexGrid11.TextMatrix(r, 6)), Val(MSHFlexGrid11.TextMatrix(r, 16)), Val(MSHFlexGrid11.TextMatrix(r, 13)), Val(MSHFlexGrid11.TextMatrix(r, 8)))
            
        tmp = " Update  Txn_CDTFBOE Set Amount = " & Val(mInvAmountOnPrint) & " Where BOEID =  " & Val(mBOEIDNew) & "   "
        Call Tmp7Table
            
        TmpRs2.MoveNext
    Next
End If
End Sub

Private Sub Grid_HeadRRNNCDEXInvoice()
MSHFlexGrid11.Clear
MSHFlexGrid11.Rows = 3
MSHFlexGrid11.FixedCols = 0
MSHFlexGrid11.Cols = 17

MSHFlexGrid11.TextMatrix(1, 0) = "RRN No"
MSHFlexGrid11.TextMatrix(1, 1) = "CMSE Lot No" '"ISIN No"
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
''
''MSHFlexGrid11.ColWidth(0) = 1200
''MSHFlexGrid11.ColWidth(1) = 1200
''MSHFlexGrid11.ColWidth(2) = 1000
''MSHFlexGrid11.ColWidth(3) = 1200
''MSHFlexGrid11.ColWidth(4) = 1500
''MSHFlexGrid11.ColWidth(5) = 1000
''MSHFlexGrid11.ColWidth(6) = 1000
''MSHFlexGrid11.ColWidth(7) = 1000
''MSHFlexGrid11.ColWidth(8) = 1200
''MSHFlexGrid11.ColWidth(9) = 2500
''MSHFlexGrid11.ColWidth(10) = 2500
'''MSHFlexGrid11.ColWidth(10) = 1200
'''MSHFlexGrid11.ColWidth(11) = 1200
'''MSHFlexGrid11.ColWidth(12) = 1500
'''MSHFlexGrid11.ColWidth(13) = 1500
'''MSHFlexGrid11.ColWidth(14) = 1500
'''MSHFlexGrid11.ColWidth(15) = 1000
''MSHFlexGrid11.ColWidth(11) = 0
''MSHFlexGrid11.ColWidth(12) = 0
''MSHFlexGrid11.ColWidth(13) = 0
''MSHFlexGrid11.ColWidth(14) = 0
''MSHFlexGrid11.ColWidth(15) = 0
''MSHFlexGrid11.ColWidth(16) = 0

End Sub
Public Function CheckMoisturePer(RowVal As Long) As Double
Dim Retval As Double
Retval = 0

If RsMst_GSL!Status = "B" Then
   Retval = Val(MSHFlexGrid4.TextMatrix(RowVal, 21))
Else
   If (RsMst_GSL!BalanceWeight + Val(MSHFlexGrid4.TextMatrix(RowVal, 10))) > 0 Then
      Retval = (RsMst_GSL!BalanceWeight * IIf(IsNull(RsMst_GSL!CITF_MoisturePer), 0, RsMst_GSL!CITF_MoisturePer)) + (Val(MSHFlexGrid4.TextMatrix(RowVal, 10)) * Val(MSHFlexGrid4.TextMatrix(RowVal, 21)))
      Retval = Retval / (RsMst_GSL!BalanceWeight + Val(MSHFlexGrid4.TextMatrix(RowVal, 10)))
   End If
End If
   
CheckMoisturePer = Retval

End Function

Private Function ChkGodownYearlyFlag() As Boolean
Dim Retval As Boolean

Retval = False

For I = 2 To MSHFlexGrid4.Rows - 1
    If MSHFlexGrid4.TextMatrix(I, 22) = "" Then Exit For
    If MSHFlexGrid4.TextMatrix(I, 22) = True Then
       Retval = True
    End If
Next

ChkGodownYearlyFlag = Retval

End Function

Private Sub CmbVehicleTypeID_GotFocus()
tmp = CmbVehicleTypeID.Text
Call TableMst_VehicleType
CmbVehicleTypeID.Clear
If RsMst_VehicleType.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_VehicleType.RecordCount
    CmbVehicleTypeID.AddItem RsMst_VehicleType!VehicleType
    RsMst_VehicleType.MoveNext
Next
CmbVehicleTypeID.Text = tmp
End Sub
'----
Private Sub CmbVehicleTypeID_LostFocus()
If CmbVehicleTypeID.Text = "" Then
   Exit Sub
End If
RsMst_VehicleType.MoveFirst

If tmp <> CmbVehicleTypeID.Text Then
   TxtMinimum = ""
   TxtMaximum = ""
End If

RsMst_VehicleType.Find "VehicleType = '" & CmbVehicleTypeID.Text & "'"
If RsMst_VehicleType.EOF Then
   MsgBox "Invalid selection "
   CmbVehicleTypeID.SetFocus
   Exit Sub
End If
mVehicleTypeID = RsMst_VehicleType!VehicleTypeID
TxtMinimum = IIf(IsNull(RsMst_VehicleType!MinTareWt), 0, RsMst_VehicleType!MinTareWt)
TxtMaximum = IIf(IsNull(RsMst_VehicleType!MaxTareWt), 0, RsMst_VehicleType!MaxTareWt)

End Sub


Private Sub Gen_MailTareWeightDiscripency()

Dim TOADD As Variant
Dim TOCC, mMsg, mGridMsg As Variant
Dim strSplitString() As String
Dim EmpNo_, EmpName_ As String

mMsg = ""
TOADD = ""
TOCC = ""

mMsg = "The Tare weight entered for the following vehicles are beyond the approved range. Please look into the matter on priority basis"
mMsg = mMsg & "<br><br><br>"
mMsg = mMsg & "<table border = 1 face= Candara><tr><td colspan=2>CITF Details (From Store-Man)</td></tr>"
mMsg = mMsg & "<tr><td>CDTF No</td><td>" & TxtCxTFNo.Text & "</td> </tr>"
mMsg = mMsg & "<tr><td>CITF Date</td><td>" & Format(TxtCxTFDate.Value, "dd-MMM-yyyy") & "</td></tr>"
mMsg = mMsg & "<tr><td>Location</td><td>" & GetLocationFromCMP(RsTxn_CxTF_Details!CMPID) & "</td> </tr>"
mMsg = mMsg & "<tr><td>State</td><td>" & GetStateFromLocation(GetLocationIDFromCMP(mCMPID)) & "</td></tr>"
mMsg = mMsg & "<tr><td>CMPID & Address</td><td>" & mCMPID & "-" & GetCMPAddress(mCMPID, "A") & "</td></tr>"
mMsg = mMsg & "<tr><td>Commodity</td><td>" & CmbCommodityID.Text & "</td></tr>"
mMsg = mMsg & "<tr><td>Transaction Type</td><td>" & CmbExchangeTypeID.Text & "</td></tr>"
mMsg = mMsg & "<tr><td>Name Of Agent </td><td>" & CmbAgent.Text & "</td></tr>"
mMsg = mMsg & "<tr><td>Total No of Bags</td><td>" & MSHFlexGrid4.TextMatrix(0, 6) & "</td></tr>"
mMsg = mMsg & "<tr><td>Total Electronic Weight in MT </td><td>" & MSHFlexGrid4.TextMatrix(0, 10) & "</td></tr>"

mGridMsg = ""

For I = 1 To MSHFlexGrid3.Rows - 1
    If MSHFlexGrid3.TextMatrix(I, 1) <> "" Then
       If I = 1 Then
          mGridMsg = mGridMsg & "<tr><td> " & MSHFlexGrid3.TextMatrix(I, 9) & " </td><td> " & MSHFlexGrid3.TextMatrix(I, 1) & "~" & MSHFlexGrid3.TextMatrix(I, 2) & " </td><td>Approved Tare Weight Range</td><td> " & MSHFlexGrid3.TextMatrix(I, 4) & " </td><td>Remarks</td></tr>"
       Else
          mGridMsg = mGridMsg & "<tr><td> " & MSHFlexGrid3.TextMatrix(I, 9) & " </td><td> " & MSHFlexGrid3.TextMatrix(I, 1) & "~" & MSHFlexGrid3.TextMatrix(I, 2) & " </td><td> " & MSHFlexGrid3.TextMatrix(I, 11) & "~" & MSHFlexGrid3.TextMatrix(I, 12) & " </td><td> " & MSHFlexGrid3.TextMatrix(I, 4) & " </td><td> " & MSHFlexGrid3.TextMatrix(I, 7) & " </td></tr>"
       End If
    End If
Next

mMsg = mMsg & "<tr><td> TTD Details </td><td><table border = 1 face= Candara>" & mGridMsg & "</table></td></table>"


tmp = " Select EMP.EmployeeID,EMP.EmployeeNo,EMP.EmployeeName,EMP.EmployeeEmailID,EMP1.EmployeeID as 'AMEmpID',EMP1.EmployeeNo as 'AMEmpNo',EMP1.EmployeeName as 'AMName',EMP1.EmployeeEmailID as 'AMEmailID'" & _
      " From Mst_EmpCMPMapping EGM" & _
      " Inner Join Mst_Employee EMP on EGM.EmployeeID = EMP.EmployeeID" & _
      " Inner Join Mst_Employee EMP1 on EMP.DirectReportingTo = EMP1.EmployeeID" & _
      " Where EMP.Flag = 'Y' And EGM.CMPID = " & mCMPID
'      " Group By EMP.EmployeeID,EMP.EmployeeNo,EMP.EmployeeName,EMP.EmployeeEmailID"

Call Tmp3Table

If TmpRs3.RecordCount > 0 Then
   TOADD = TmpRs3!EmployeeEmailID
   TOCC = IIf(IsNull(TmpRs3!AMEmailID), "", TmpRs3!AMEmailID) '""
End If

If TOADD = "" Then
   TOADD = "kuldip.s@ncml.com"
End If
 
 '"cmg@ncml.com" & "," & Gen_UserEmailID  'Gen_UserEmailID
If TOCC <> "" Then
   TOCC = TOCC & ",cmg@ncml.com," & Gen_UserEmailID
Else
   TOCC = Gen_UserEmailID
End If


mMsgSubject = "CITF No - " & Val(TxtCxTFNo.Text) & " is having inconsistency in the tare weight entered in StoreMan at Location - " & GetLocationFromCMP(RsTxn_CxTF_Details!CMPID) & ""

Call SendMail(TOADD, TOCC, mMsgSubject, mMsg, "D")

End Sub


Private Function ChkTareWeightDiscripency() As Boolean
Dim Retval As Boolean
Retval = False

For I = 2 To MSHFlexGrid3.Rows - 1
    If MSHFlexGrid3.TextMatrix(I, 1) <> "" Then
       If Val(MSHFlexGrid3.TextMatrix(I, 4)) < Val(MSHFlexGrid3.TextMatrix(I, 11)) Or Val(MSHFlexGrid3.TextMatrix(I, 4)) > Val(MSHFlexGrid3.TextMatrix(I, 12)) Then
          Retval = True
          Exit For
       End If
    End If
Next
ChkTareWeightDiscripency = Retval

End Function

Public Sub GetClientFeedbackBlockingDays()

Dim objFS As Scripting.FileSystemObject
Dim objFile As File
Dim objTS As TextStream
Dim lngFreefile As Long
Dim sFileLine As Variant

Set objFS = New Scripting.FileSystemObject
Set objFile = objFS.GetFile(App.Path & "\" & "ClientFeedbackBlockingDays.ini")
Set objTS = objFile.OpenAsTextStream(ForReading, TristateUseDefault)

sFileLine = objTS.ReadLine
mBlockingDays = Replace(LCase(sFileLine), "days", "")
objTS.Close

End Sub
