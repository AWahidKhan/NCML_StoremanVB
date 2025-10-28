VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmMst_SecurityRequest 
   Caption         =   "Security Request"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Enabled         =   0   'False
      Height          =   9780
      Left            =   0
      TabIndex        =   37
      Top             =   -120
      Width           =   15060
      Begin VB.Frame Frame0 
         Height          =   2055
         Left            =   6840
         TabIndex        =   38
         Top             =   6240
         Visible         =   0   'False
         Width           =   6615
         Begin VB.ComboBox CmbSSecurityDesignation 
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
            ItemData        =   "FrmMst_SecurityRequest.frx":0000
            Left            =   120
            List            =   "FrmMst_SecurityRequest.frx":0002
            Sorted          =   -1  'True
            TabIndex        =   107
            Top             =   1200
            Width           =   2655
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
            ItemData        =   "FrmMst_SecurityRequest.frx":0004
            Left            =   6960
            List            =   "FrmMst_SecurityRequest.frx":0011
            Sorted          =   -1  'True
            TabIndex        =   106
            Top             =   1200
            Width           =   2175
         End
         Begin VB.ComboBox CmbSRequestTo 
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
            ItemData        =   "FrmMst_SecurityRequest.frx":0032
            Left            =   2880
            List            =   "FrmMst_SecurityRequest.frx":0034
            Sorted          =   -1  'True
            TabIndex        =   33
            Top             =   1200
            Width           =   3975
         End
         Begin VB.ComboBox CmbSSecurityType 
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
            ItemData        =   "FrmMst_SecurityRequest.frx":0036
            Left            =   11040
            List            =   "FrmMst_SecurityRequest.frx":0038
            Sorted          =   -1  'True
            TabIndex        =   32
            Top             =   360
            Width           =   3615
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
            Left            =   2790
            Sorted          =   -1  'True
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   353
            Width           =   3180
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
            Height          =   360
            Left            =   13635
            TabIndex        =   35
            Top             =   840
            Width           =   1000
         End
         Begin VB.CommandButton CmdGo 
            Caption         =   "Go"
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
            Left            =   12615
            TabIndex        =   34
            Top             =   840
            Width           =   1000
         End
         Begin VB.ComboBox CmbSSecurityAgency 
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
            Left            =   6030
            Sorted          =   -1  'True
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   360
            Width           =   4905
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   1455
            Left            =   75
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   1710
            Width           =   14475
            _ExtentX        =   25532
            _ExtentY        =   2566
            _Version        =   393216
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSComCtl2.DTPicker STo 
            Height          =   360
            Left            =   1450
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
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
            Format          =   112132097
            CurrentDate     =   36494
         End
         Begin MSComCtl2.DTPicker SFrom 
            Height          =   360
            Left            =   105
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
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
            Format          =   112132097
            CurrentDate     =   36494
         End
         Begin VB.Label Label36 
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
            Left            =   7770
            TabIndex        =   109
            Top             =   900
            Width           =   555
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Security Designation"
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
            Left            =   525
            TabIndex        =   108
            Top             =   900
            Width           =   1845
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            Caption         =   "Request To"
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
            Left            =   4335
            TabIndex        =   103
            Top             =   900
            Width           =   1065
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            Caption         =   "Security Type"
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
            Left            =   12225
            TabIndex        =   102
            Top             =   135
            Width           =   1245
         End
         Begin VB.Label Label31 
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
            Left            =   3998
            TabIndex        =   45
            Top             =   135
            Width           =   765
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
            Left            =   9255
            TabIndex        =   44
            Top             =   900
            Width           =   2415
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
            Left            =   11745
            TabIndex        =   43
            Top             =   900
            Width           =   90
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
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
            Height          =   240
            Left            =   1815
            TabIndex        =   42
            Top             =   120
            Width           =   735
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            Caption         =   "Security Agency"
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
            TabIndex        =   41
            Top             =   135
            Width           =   1455
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
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
            Height          =   240
            Left            =   360
            TabIndex        =   40
            Top             =   120
            Width           =   945
         End
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   4755
         Left            =   120
         TabIndex        =   46
         Top             =   3885
         Width           =   14835
         _ExtentX        =   26167
         _ExtentY        =   8387
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Godown Mapping "
         TabPicture(0)   =   "FrmMst_SecurityRequest.frx":003A
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "MSHFlexGrid2"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Frame3"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Log History"
         TabPicture(1)   =   "FrmMst_SecurityRequest.frx":0056
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "TxtCreated"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "TxtUpdated"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "TxtApprovedBy"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "MSHFlexGrid3"
         Tab(1).Control(4)=   "Label17"
         Tab(1).Control(5)=   "Label15"
         Tab(1).Control(6)=   "Label35"
         Tab(1).ControlCount=   7
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
            Left            =   -73350
            Locked          =   -1  'True
            TabIndex        =   112
            TabStop         =   0   'False
            Top             =   480
            Width           =   4590
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
            Left            =   -67515
            Locked          =   -1  'True
            TabIndex        =   111
            TabStop         =   0   'False
            Top             =   480
            Width           =   4590
         End
         Begin VB.TextBox TxtApprovedBy 
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
            Left            =   -73350
            Locked          =   -1  'True
            TabIndex        =   110
            TabStop         =   0   'False
            Top             =   1080
            Width           =   4590
         End
         Begin VB.Frame Frame3 
            Enabled         =   0   'False
            Height          =   2220
            Left            =   120
            TabIndex        =   72
            Top             =   300
            Width           =   14655
            Begin VB.CommandButton CmdDeleteFromList 
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
               Left            =   12885
               TabIndex        =   17
               Top             =   1725
               Width           =   1710
            End
            Begin VB.CommandButton CmdAddtoList 
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
               Left            =   11535
               TabIndex        =   16
               Top             =   1725
               Width           =   1335
            End
            Begin VB.TextBox TxtRemarksD 
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
               Height          =   525
               Left            =   7500
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   15
               Top             =   1080
               Width           =   7095
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
               Left            =   135
               Sorted          =   -1  'True
               TabIndex        =   13
               Top             =   397
               Width           =   4125
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
               Left            =   4320
               Sorted          =   -1  'True
               TabIndex        =   14
               Top             =   397
               Width           =   3105
            End
            Begin VB.TextBox TxtBalanceBags 
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               Height          =   375
               Left            =   7470
               Locked          =   -1  'True
               TabIndex        =   76
               TabStop         =   0   'False
               Top             =   390
               Width           =   2130
            End
            Begin VB.TextBox TxtBalanceQty 
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               Height          =   375
               Left            =   9630
               Locked          =   -1  'True
               TabIndex        =   75
               TabStop         =   0   'False
               Top             =   390
               Width           =   1965
            End
            Begin VB.TextBox TxtAUM 
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               Height          =   375
               Left            =   11640
               Locked          =   -1  'True
               TabIndex        =   74
               TabStop         =   0   'False
               Top             =   390
               Width           =   2955
            End
            Begin VB.TextBox TxtGodownAddress 
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               Height          =   510
               Left            =   120
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   73
               TabStop         =   0   'False
               Top             =   1080
               Width           =   7305
            End
            Begin VB.Label Label29 
               AutoSize        =   -1  'True
               Caption         =   "Display of Message when same godown will be mapped again"
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
               TabIndex        =   97
               Top             =   1755
               Visible         =   0   'False
               Width           =   6450
            End
            Begin VB.Label Label10 
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
               Left            =   10635
               TabIndex        =   83
               Top             =   810
               Width           =   825
            End
            Begin VB.Label Label11 
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
               Height          =   240
               Left            =   1980
               TabIndex        =   82
               Top             =   150
               Width           =   495
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               Caption         =   "Godown"
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
               Left            =   5490
               TabIndex        =   81
               Top             =   150
               Width           =   855
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               Caption         =   "Balance Bags"
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
               Left            =   7898
               TabIndex        =   80
               Top             =   120
               Width           =   1275
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
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
               Height          =   240
               Left            =   10065
               TabIndex        =   79
               Top             =   120
               Width           =   1095
            End
            Begin VB.Label Label19 
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
               Left            =   12892
               TabIndex        =   78
               Top             =   120
               Width           =   450
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               Caption         =   "Godown Address"
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
               Left            =   3007
               TabIndex        =   77
               Top             =   810
               Width           =   1560
            End
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
            Height          =   2070
            Left            =   120
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   2565
            Width           =   14625
            _ExtentX        =   25797
            _ExtentY        =   3651
            _Version        =   393216
            FixedCols       =   0
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
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
            Height          =   4185
            Left            =   -71280
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   1320
            Visible         =   0   'False
            Width           =   14610
            _ExtentX        =   25770
            _ExtentY        =   7382
            _Version        =   393216
            FixedCols       =   0
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
         Begin VB.Label Label17 
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
            TabIndex        =   115
            Top             =   540
            Width           =   855
         End
         Begin VB.Label Label15 
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
            Left            =   -68565
            TabIndex        =   114
            Top             =   540
            Width           =   885
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            Caption         =   "Approved By :"
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
            TabIndex        =   113
            Top             =   1140
            Width           =   1275
         End
      End
      Begin VB.Frame ButtonFrm 
         Height          =   1080
         Left            =   120
         TabIndex        =   60
         Top             =   8625
         Width           =   14865
         Begin VB.CommandButton AddCmd 
            Caption         =   "&Add New"
            Height          =   400
            Left            =   105
            TabIndex        =   18
            Top             =   195
            Width           =   3060
         End
         Begin VB.CommandButton NextCmd 
            Caption         =   "&Next"
            Height          =   400
            Left            =   105
            TabIndex        =   23
            Top             =   600
            Width           =   3060
         End
         Begin VB.CommandButton SaveCmd 
            Caption         =   "Save"
            Height          =   400
            Left            =   3165
            TabIndex        =   19
            Top             =   195
            Width           =   3060
         End
         Begin VB.CommandButton PreviousCmd 
            Caption         =   "&Previous"
            Height          =   400
            Left            =   3165
            TabIndex        =   24
            Top             =   600
            Width           =   3060
         End
         Begin VB.CommandButton EditCmd 
            Caption         =   "&Edit"
            Height          =   400
            Left            =   6225
            TabIndex        =   20
            Top             =   195
            Width           =   2940
         End
         Begin VB.CommandButton FirstCmd 
            Caption         =   "&First"
            Height          =   400
            Left            =   6225
            TabIndex        =   25
            Top             =   600
            Width           =   2940
         End
         Begin VB.CommandButton DeleteCmd 
            Caption         =   "&Delete"
            Height          =   400
            Left            =   9165
            TabIndex        =   21
            Top             =   195
            Width           =   2940
         End
         Begin VB.CommandButton LastCmd 
            Caption         =   "&Last"
            Height          =   400
            Left            =   9165
            TabIndex        =   26
            Top             =   600
            Width           =   2940
         End
         Begin VB.CommandButton SearchCmd 
            Caption         =   "Search"
            Height          =   400
            Left            =   12105
            TabIndex        =   22
            Top             =   195
            Width           =   2700
         End
         Begin VB.CommandButton ExitCmd 
            Caption         =   "E&xit"
            Height          =   400
            Left            =   12105
            TabIndex        =   27
            Top             =   600
            Width           =   2700
         End
      End
      Begin VB.Frame Frame1 
         Enabled         =   0   'False
         Height          =   3750
         Left            =   120
         TabIndex        =   49
         Top             =   105
         Width           =   14880
         Begin VB.CommandButton CmdGetReport 
            Caption         =   "Godown Mapped Report"
            Height          =   360
            Left            =   12675
            TabIndex        =   105
            Top             =   195
            Width           =   2100
         End
         Begin VB.CommandButton CmdLocation 
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
            Left            =   12240
            TabIndex        =   52
            TabStop         =   0   'False
            Top             =   225
            Width           =   375
         End
         Begin MSComCtl2.DTPicker TxtRequestDate 
            Height          =   375
            Left            =   5370
            TabIndex        =   53
            TabStop         =   0   'False
            Top             =   188
            Width           =   1485
            _ExtentX        =   2619
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
            Format          =   112132097
            CurrentDate     =   39884
         End
         Begin VB.CommandButton CmdRequestTo 
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
            Left            =   7005
            TabIndex        =   84
            TabStop         =   0   'False
            Top             =   675
            Width           =   375
         End
         Begin VB.CommandButton CmdSecurityAgency 
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
            Left            =   14400
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   675
            Width           =   375
         End
         Begin MSComCtl2.DTPicker TxtAgreeEndDate 
            Height          =   375
            Left            =   5340
            TabIndex        =   5
            Top             =   1200
            Width           =   1515
            _ExtentX        =   2672
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
            Format          =   112132097
            CurrentDate     =   39884
         End
         Begin VB.TextBox TxtNoofopenGodown 
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
            Left            =   8700
            Locked          =   -1  'True
            TabIndex        =   61
            TabStop         =   0   'False
            Top             =   1200
            Width           =   2520
         End
         Begin VB.TextBox TxtNoofGodownsMapped 
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
            Left            =   5325
            Locked          =   -1  'True
            TabIndex        =   87
            TabStop         =   0   'False
            Top             =   1740
            Width           =   1530
         End
         Begin VB.TextBox TxtTotalAUM 
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
            Left            =   12270
            Locked          =   -1  'True
            TabIndex        =   91
            TabStop         =   0   'False
            Top             =   1740
            Width           =   2505
         End
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
            Left            =   13335
            Locked          =   -1  'True
            TabIndex        =   51
            TabStop         =   0   'False
            Top             =   2670
            Visible         =   0   'False
            Width           =   1440
         End
         Begin VB.ComboBox CmbSecurityType 
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
            ItemData        =   "FrmMst_SecurityRequest.frx":0072
            Left            =   1980
            List            =   "FrmMst_SecurityRequest.frx":0074
            Sorted          =   -1  'True
            TabIndex        =   7
            Top             =   2190
            Width           =   4335
         End
         Begin MSComCtl2.DTPicker TxtAgreeStartDate 
            Height          =   375
            Left            =   1980
            TabIndex        =   4
            Top             =   1200
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
            Format          =   112132097
            CurrentDate     =   39884
         End
         Begin VB.ComboBox CmbRequestTo 
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
            Left            =   1980
            Sorted          =   -1  'True
            TabIndex        =   2
            Top             =   645
            Width           =   4875
         End
         Begin VB.TextBox TxtRequestID 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   375
            Left            =   1980
            Locked          =   -1  'True
            TabIndex        =   99
            TabStop         =   0   'False
            Top             =   188
            Width           =   1575
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
            Height          =   585
            Left            =   1980
            MaxLength       =   2000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   11
            Top             =   3105
            Width           =   4875
         End
         Begin VB.TextBox TxtNoofGuards 
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
            Left            =   1980
            Locked          =   -1  'True
            TabIndex        =   101
            TabStop         =   0   'False
            Top             =   1740
            Width           =   1575
         End
         Begin VB.TextBox TxtRatePerMonth 
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
            Left            =   1980
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   2663
            Width           =   4335
         End
         Begin VB.TextBox TxtRequestTo 
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
            Left            =   1980
            Locked          =   -1  'True
            TabIndex        =   100
            TabStop         =   0   'False
            Top             =   645
            Width           =   4875
         End
         Begin VB.ComboBox CmbSecurityAgency 
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
            Left            =   8700
            Sorted          =   -1  'True
            TabIndex        =   3
            Top             =   645
            Width           =   5655
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
            ItemData        =   "FrmMst_SecurityRequest.frx":0076
            Left            =   8700
            List            =   "FrmMst_SecurityRequest.frx":0083
            Sorted          =   -1  'True
            TabIndex        =   10
            Top             =   2670
            Width           =   2175
         End
         Begin VB.ComboBox CmbSecurityDesignation 
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
            ItemData        =   "FrmMst_SecurityRequest.frx":00A4
            Left            =   8700
            List            =   "FrmMst_SecurityRequest.frx":00A6
            Sorted          =   -1  'True
            TabIndex        =   8
            Top             =   2190
            Width           =   5655
         End
         Begin VB.TextBox TxtTotalCost 
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
            Left            =   8700
            Locked          =   -1  'True
            TabIndex        =   89
            TabStop         =   0   'False
            Top             =   1740
            Width           =   2130
         End
         Begin MSComCtl2.DTPicker TxtSalaryPaidUpto 
            Height          =   375
            Left            =   12660
            TabIndex        =   6
            Top             =   1680
            Visible         =   0   'False
            Width           =   1395
            _ExtentX        =   2461
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
            Format          =   112132097
            CurrentDate     =   39884
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
            Left            =   8700
            Sorted          =   -1  'True
            TabIndex        =   1
            Top             =   195
            Width           =   3495
         End
         Begin VB.TextBox TxtApproveRemarks 
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
            Height          =   585
            Left            =   8700
            Locked          =   -1  'True
            MaxLength       =   2000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   12
            Top             =   3120
            Width           =   6075
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
            Left            =   8700
            Locked          =   -1  'True
            TabIndex        =   95
            TabStop         =   0   'False
            Top             =   2670
            Width           =   2175
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
            Left            =   8700
            Locked          =   -1  'True
            TabIndex        =   98
            TabStop         =   0   'False
            Top             =   195
            Width           =   3495
         End
         Begin VB.TextBox TxtSecurityAgency 
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
            Left            =   8700
            Locked          =   -1  'True
            TabIndex        =   54
            TabStop         =   0   'False
            Top             =   645
            Width           =   5655
         End
         Begin VB.TextBox TxtSecurityType 
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
            Left            =   1980
            Locked          =   -1  'True
            TabIndex        =   93
            TabStop         =   0   'False
            Top             =   2190
            Width           =   4335
         End
         Begin VB.TextBox TxtSecurityDesignation 
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
            Left            =   8700
            Locked          =   -1  'True
            TabIndex        =   94
            TabStop         =   0   'False
            Top             =   2190
            Width           =   5655
         End
         Begin VB.CommandButton CmdSecurityDesignation 
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
            Left            =   14400
            TabIndex        =   67
            TabStop         =   0   'False
            Top             =   2220
            Width           =   375
         End
         Begin VB.CommandButton CmdSecurityType 
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
            Left            =   6480
            TabIndex        =   68
            TabStop         =   0   'False
            Top             =   2220
            Width           =   375
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            Caption         =   "Approval Remarks"
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
            Left            =   6975
            TabIndex        =   104
            Top             =   3285
            Width           =   1695
         End
         Begin VB.Label Label12 
            Caption         =   "G_UID"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   13260
            TabIndex        =   57
            Top             =   285
            Visible         =   0   'False
            Width           =   630
         End
         Begin VB.Label Label28 
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
            Left            =   7665
            TabIndex        =   59
            Top             =   255
            Width           =   900
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Request Date"
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
            Left            =   3660
            TabIndex        =   58
            Top             =   255
            Width           =   1245
         End
         Begin VB.Label LbClaimID 
            AutoSize        =   -1  'True
            Caption         =   "Request ID"
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
            TabIndex        =   55
            Top             =   255
            Width           =   1005
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Request to"
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
            TabIndex        =   85
            Top             =   705
            Width           =   1140
         End
         Begin VB.Label Label4 
            Caption         =   "Security Agency"
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
            Left            =   7680
            TabIndex        =   56
            Top             =   585
            Width           =   1140
         End
         Begin VB.Label Label5 
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
            Left            =   135
            TabIndex        =   70
            Top             =   3277
            Width           =   825
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "Total AUM"
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
            Left            =   11175
            TabIndex        =   92
            Top             =   1807
            Width           =   960
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Total Cost"
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
            Left            =   6975
            TabIndex        =   90
            Top             =   1807
            Width           =   915
         End
         Begin VB.Label Label25 
            Caption         =   "No of Godowns Mapped"
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
            Left            =   3675
            TabIndex        =   88
            Top             =   1687
            Width           =   1470
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "No of Guards"
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
            TabIndex        =   86
            Top             =   1807
            Width           =   1185
         End
         Begin VB.Label Label7 
            Caption         =   "Salary Payable Upto"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   10935
            TabIndex        =   66
            Top             =   1590
            Visible         =   0   'False
            Width           =   1680
         End
         Begin VB.Label Label6 
            Caption         =   "No of Open Godowns"
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
            Left            =   7665
            TabIndex        =   65
            Top             =   1147
            Width           =   1320
         End
         Begin VB.Label Label1 
            Caption         =   "Security End Date"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   510
            Left            =   3675
            TabIndex        =   63
            Top             =   1132
            Width           =   1200
         End
         Begin VB.Label Label8 
            Caption         =   "Security Start Date"
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
            Left            =   135
            TabIndex        =   62
            Top             =   1140
            Width           =   1260
         End
         Begin VB.Label Label9 
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
            Left            =   6975
            TabIndex        =   96
            Top             =   2730
            Width           =   660
         End
         Begin VB.Label Label2 
            Caption         =   "Rate per month (Excluding Service Tax)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   135
            TabIndex        =   64
            Top             =   2543
            Width           =   1710
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "Security Type"
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
            Left            =   135
            TabIndex        =   71
            Top             =   2250
            Width           =   1455
         End
         Begin VB.Label Label3 
            Caption         =   "Security Designation"
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
            Left            =   6975
            TabIndex        =   69
            Top             =   2130
            Width           =   1710
         End
      End
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      TabIndex        =   36
      Top             =   10320
      Width           =   10935
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   9015
      Left            =   15000
      TabIndex        =   0
      Top             =   0
      Width           =   255
   End
End
Attribute VB_Name = "FrmMst_SecurityRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mLocationID, mSecurityAgencyID, mSecurityTypeID, mSecurityDesignationID, mRequestToID, mCMPID, mGodownID As Variant
Dim Edit_Flg, mStateID, mSAID As Variant
Dim mAgreeStartDt, mAgreeEndDt As Variant


Private Sub Grid_HeadGodownMapping()
With MSHFlexGrid2
    .Clear
    .Rows = 3
    .FixedRows = 2
    .Cols = 10
    
    .TextMatrix(1, 0) = "CMP Name"
    .TextMatrix(1, 1) = "Godown No"
    .TextMatrix(1, 2) = "Address"
    .TextMatrix(1, 3) = "Balance Bags"
    .TextMatrix(1, 4) = "Balance Stock"
    .TextMatrix(1, 5) = "AUM"
    .TextMatrix(1, 6) = "Remarks"
    .TextMatrix(1, 7) = "CMP ID"
    .TextMatrix(1, 8) = "Godown ID"
    .TextMatrix(1, 9) = "Security Detail ID"
    
    .ColWidth(0) = 1800
    .ColWidth(1) = 1600
    .ColWidth(2) = 4500
    .ColWidth(3) = 1200
    .ColWidth(4) = 1200
    .ColWidth(5) = 1600
    .ColWidth(6) = 3500
    .ColWidth(7) = 0
    .ColWidth(8) = 0
    .ColWidth(9) = 0
    
    .RowHeight(1) = 650
    .WordWrap = True
End With
End Sub

Private Sub Grid_HeadGodownInfo()
With MSHFlexGrid3
    .Clear
    .Rows = 3
    .FixedRows = 2
    .Cols = 13

    .TextMatrix(1, 0) = "Security Type ~ SecurityID"
    .TextMatrix(1, 1) = "CMP Name"
    .TextMatrix(1, 2) = "Godown No"
    .TextMatrix(1, 3) = "Address"
    .TextMatrix(1, 4) = "Balance Bags"
    .TextMatrix(1, 5) = "Balance Qty"
    .TextMatrix(1, 6) = "AUM"
    .TextMatrix(1, 7) = "Remarks"
    .TextMatrix(1, 8) = "DetailID" '"Security TypeID"
    .TextMatrix(1, 9) = "CMPID"
    .TextMatrix(1, 10) = "GodownID"
    .TextMatrix(1, 11) = "GodMapDID"
    .TextMatrix(1, 12) = "Security Designation"
    
    .ColWidth(0) = 2500
    .ColWidth(1) = 1800
    .ColWidth(2) = 1000
    .ColWidth(3) = 4500
    .ColWidth(4) = 1200
    .ColWidth(5) = 1200
    .ColWidth(6) = 1200
    .ColWidth(7) = 3500
    .ColWidth(8) = 0
    .ColWidth(9) = 0
    .ColWidth(10) = 0
    .ColWidth(11) = 0
    .ColWidth(12) = 1500
    
    .RowHeight(1) = 650
    .WordWrap = True
End With
'With MSHFlexGrid3
'    .Clear
'    .Rows = 3
'    .FixedRows = 2
'    .Cols = 8
'
'    .TextMatrix(1, 0) = "CMP Name"
'    .TextMatrix(1, 1) = "Godown No"
'    .TextMatrix(1, 2) = "Address"
'    .TextMatrix(1, 3) = "Balance Bags"
'    .TextMatrix(1, 4) = "Balance Stock"
'    .TextMatrix(1, 5) = "AUM"
'    .TextMatrix(1, 6) = "Security Type"
'    .TextMatrix(1, 7) = "Cost"
'
'    .ColWidth(0) = 1800
'    .ColWidth(1) = 1600
'    .ColWidth(2) = 4000
'    .ColWidth(3) = 1200
'    .ColWidth(4) = 1200
'    .ColWidth(5) = 1600
'    .ColWidth(6) = 3600
'    .ColWidth(7) = 1800
'
'    .RowHeight(1) = 650
'    .WordWrap = True
'End With
End Sub

Private Sub AddCmd_Click()
Edit_Flg = "A"
Frame1.Enabled = True
Frame3.Enabled = True
Call GButtonLock(Me, False)

Call ClearControl
Call ComboVisible(True)
TxtStatus.Visible = True
CmbStatus.Visible = False

If RsMst_SecurityRequest.RecordCount > 0 Then
   CmbLocation.SetFocus
Else
   SearchCmd.Enabled = False
End If
End Sub

Private Sub ComboVisible(var1 As Boolean)

CmbLocation.Visible = var1
CmbRequestTo.Visible = var1
CmbSecurityAgency.Visible = var1
CmbSecurityType.Visible = var1
CmbSecurityDesignation.Visible = var1
CmbStatus.Visible = var1

CmdLocation.Visible = var1
CmdRequestTo.Visible = var1
CmdSecurityAgency.Visible = var1
CmdSecurityType.Visible = var1
CmdSecurityDesignation.Visible = var1

TxtLocation.Visible = IIf(var1 = True, False, True)
TxtRequestTo.Visible = IIf(var1 = True, False, True)
TxtSecurityAgency.Visible = IIf(var1 = True, False, True)
TxtSecurityType.Visible = IIf(var1 = True, False, True)
TxtSecurityDesignation.Visible = IIf(var1 = True, False, True)
TxtStatus.Visible = IIf(var1 = True, False, True)

End Sub

Private Sub ClearControl()

TxtRequestDate.Value = Date
TxtAgreeStartDate.Value = Date - 30
TxtAgreeEndDate.Value = Date + 30
TxtSalaryPaidUpto.Value = Date + 30

TxtRequestID.Text = ""
TxtLocation.Text = ""
TxtRequestTo.Text = ""
TxtSecurityAgency.Text = ""
TxtNoofopenGodown.Text = ""
TxtNoofGuards.Text = ""
TxtNoofGodownsMapped.Text = ""
TxtTotalCost.Text = ""
TxtTotalAUM.Text = ""
TxtSecurityType.Text = ""
TxtSecurityDesignation.Text = ""
TxtRatePerMonth.Text = ""
TxtStatus.Text = "Pending"
CmbStatus.Text = "Pending"
TxtRemarks.Text = ""
TxtApproveRemarks.Text = ""
TxtCreated.Text = ""
TxtUpdated.Text = ""
TxtApprovedBy.Text = ""

CmbLocation.Text = ""
CmbRequestTo.Text = ""
CmbSecurityAgency.Text = ""
CmbSecurityType.Text = ""
CmbSecurityDesignation.Text = ""

Call Grid_HeadGodownInfo
Call Grid_HeadGodownMapping
Call ClearGodownMapping

End Sub

Private Sub ClearGodownMapping()
CmbCMP.Enabled = True
CmbGodown.Enabled = True
CmbCMP.Text = ""
CmbGodown.Text = ""
TxtBalanceBags.Text = ""
TxtBalanceQty.Text = ""
TxtAUM.Text = ""
TxtGodownAddress.Text = ""
TxtRemarksD.Text = ""
CmdAddtoList.Caption = "Add to List"
CmdDeletefromList.Enabled = False
End Sub

Private Sub CmbCMP_GotFocus()
If CmbLocation.Text = "" Then
    MsgBox "Select Location First!!!"
    CmbLocation.SetFocus
    Exit Sub
End If
tmp = CmbCMP.Text
Call TableMst_CMP(" where CloseDate is null And LocationID = " & mLocationID)

CmbCMP.Clear
If RsMst_CMP.RecordCount = 0 Then
    MsgBox "No Record found !!!! "
    SSTab1.Tab = 0
    'CmbCmp.SetFocus
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
    TxtBalanceBags.Text = ""
    TxtBalanceQty.Text = ""
    TxtGodownAddress.Text = ""
    TxtAUM.Text = ""
    CmbGodown.Text = ""
    Exit Sub
End If

RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPName = '" & CmbCMP.Text & "'"
If RsMst_CMP.EOF Then
    MsgBox "Invalid selection "
    CmbCMP.SetFocus
    SSTab1.Tab = 0
    Exit Sub
End If
If tmp <> CmbCMP.Text Then
    TxtBalanceBags.Text = ""
    TxtBalanceQty.Text = ""
    TxtGodownAddress.Text = ""
    TxtAUM.Text = ""
    CmbGodown.Text = ""
End If
mCMPID = RsMst_CMP!CMPID

End Sub

Private Sub CmbGodown_GotFocus()
If CmbCMP.Text = "" Then
    MsgBox "Select CMP First!!!"
    CmbCMP.SetFocus
    SSTab1.Tab = 0
    Exit Sub
End If
tmp = CmbGodown.Text
Call TableMst_Godown(" where CloseDate is null And CMPID = " & mCMPID)

CmbGodown.Clear
If RsMst_Godown.RecordCount = 0 Then
    MsgBox "No Record found !!!! "
    'CmbCmp.SetFocus
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
    TxtBalanceBags.Text = ""
    TxtBalanceQty.Text = ""
    TxtGodownAddress.Text = ""
    TxtAUM.Text = ""
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
'TxtBalanceBags.Text = IIf(IsNull(RsMst_Godown!BalanceBags), 0, RsMst_Godown!BalanceBags)
'TxtBalanceQty.Text = IIf(IsNull(RsMst_Godown!BalanceQty), 0, RsMst_Godown!BalanceQty)
TxtGodownAddress.Text = IIf(IsNull(RsMst_Godown!Address), "", RsMst_Godown!Address)
'TxtAUM.Text = ""
Call GetGodownDetail

End Sub

Private Sub GetGodownDetail()

'tmp = " Select TIMG.CommodityID, sum(TIMG.BalanceBags) BalanceBags, sum(TIMG.BalanceWeight) BalanceQty"
'tmp = tmp & " From Txn_InvMonGSLDetail TIMG"
'tmp = tmp & " Inner Join Mst_Godown MG on MG.GodownID = TIMG.GodownID"
'tmp = tmp & " Inner Join Mst_CMP MCMP on MCMP.CMPID = MG.CMPID"
'tmp = tmp & " Where TIMG.BalanceWeight > 0"
'tmp = tmp & " And TIMG.GodownID = " & mGodownID
'tmp = tmp & " Group by TIMG.CommodityID"


tmp = " Select TIMG.CommodityID, sum(TIMG.BalanceBags) BalanceBags, sum(TIMG.BalanceWeight) BalanceQty"
tmp = tmp & " From Mst_GSL TIMG"
tmp = tmp & " Inner Join Mst_Godown MG on MG.GodownID = TIMG.GodownID"
tmp = tmp & " Inner Join Mst_CMP MCMP on MCMP.CMPID = MG.CMPID"
tmp = tmp & " Where TIMG.Flag not in ('Z','E','B') And TIMG.CommodityID is not null  "
tmp = tmp & " And TIMG.GodownID = " & mGodownID
tmp = tmp & " Group by TIMG.CommodityID"

Call TmpTable

mBalanceBags = 0
mBalanceQty = 0
mTotalAUM = 0

If TmpRs.RecordCount > 0 Then
   For j = 1 To TmpRs.RecordCount
       mBalanceBags = Val(mBalanceBags) + IIf(IsNull(TmpRs!BalanceBags), 0, TmpRs!BalanceBags)
       mBalanceQty = Val(mBalanceQty) + IIf(IsNull(TmpRs!BalanceQty), 0, TmpRs!BalanceQty)
       
       tmp = " Select MCP.* From Mst_CommodityPrice MCP"
       tmp = tmp & " Inner Join Mst_State MS on MS.StateID = MCP.StateID"
       tmp = tmp & " Inner Join Mst_Location ML on ML.StateID = MS.StateID"
       tmp = tmp & " Where ML.LocationID = " & mLocationID & " And MCP.CommodityID = " & TmpRs!CommodityID
       tmp = tmp & " And YearMonth = ("
       tmp = tmp & " Select Max(YearMonth) From Mst_CommodityPrice MCP"
       tmp = tmp & " Inner Join Mst_State MS on MS.StateID = MCP.StateID"
       tmp = tmp & " Inner Join Mst_Location ML on ML.StateID = MS.StateID"
       tmp = tmp & " Where ML.LocationID = " & mLocationID & " And MCP.CommodityID = " & TmpRs!CommodityID
       tmp = tmp & " )"
       
       Call Tmp1Table
        
       If TmpRs1.RecordCount > 0 Then
          If IsNull(TmpRs1!Day31) = False Then
             mTotalAUM = Val(mTotalAUM) + ((IIf(IsNull(TmpRs!BalanceQty), 0, TmpRs!BalanceQty)) * TmpRs1!Day31)
          ElseIf IsNull(TmpRs1!Day30) = False Then
             mTotalAUM = Val(mTotalAUM) + ((IIf(IsNull(TmpRs!BalanceQty), 0, TmpRs!BalanceQty)) * TmpRs1!Day30)
          ElseIf IsNull(TmpRs1!Day29) = False Then
             mTotalAUM = Val(mTotalAUM) + ((IIf(IsNull(TmpRs!BalanceQty), 0, TmpRs!BalanceQty)) * TmpRs1!Day29)
          ElseIf IsNull(TmpRs1!Day28) = False Then
             mTotalAUM = Val(mTotalAUM) + ((IIf(IsNull(TmpRs!BalanceQty), 0, TmpRs!BalanceQty)) * TmpRs1!Day28)
          End If
       End If
       TmpRs.MoveNext
    Next
End If

TxtBalanceBags.Text = Val(mBalanceBags)
TxtBalanceQty.Text = Val(mBalanceQty)
TxtAUM.Text = Val(mTotalAUM)

End Sub

Private Sub CmbLocation_GotFocus()
tmp = CmbLocation.Text
If Gen_WcLocation = "" Then
   Call TableMst_Location
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
    CmbLocation.AddItem RsMst_Location!Locationname
    RsMst_Location.MoveNext
Next
CmbLocation.Text = tmp

End Sub

Private Sub CmbLocation_LostFocus()
If CmbLocation.Text = "" Then
   TxtNoofopenGodown.Text = ""
   CmbSecurityAgency.Text = ""
   TxtNoofGodownsMapped.Text = ""
   TxtNoofGuards.Text = ""
   TxtTotalAUM.Text = ""
   TxtTotalCost.Text = ""
   Exit Sub
End If

RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocation.Text & "'"

If RsMst_Location.EOF Then
   MsgBox "Invalid Selection!!!"
   CmbLocation.SetFocus
   Exit Sub
End If

If tmp <> CmbLocation.Text Then
   Call Grid_HeadGodownMapping
   Call ClearGodownMapping
   Call Grid_HeadGodownInfo
   TxtNoofopenGodown.Text = ""
   CmbSecurityAgency.Text = ""
   TxtNoofGodownsMapped.Text = ""
   TxtNoofGuards.Text = ""
   TxtTotalAUM.Text = ""
   TxtTotalCost.Text = ""
End If
mLocationID = RsMst_Location!LocationID
mStateID = RsMst_Location!StateID
Call GetNoofOpenGodown
Call GetGodownCostDetails
Call GetGodownInfoDetails
End Sub

Private Sub GetNoofOpenGodown()
tmp = "Select Count(MG.GodownID) NoofGodown From Mst_Godown MG Inner Join Mst_CMP MCMP ON MCMP.CMPID = MG.CMPID Where MG.CloseDate is null And MCMP.LocationID = " & mLocationID
Call TmpTable
If TmpRs.RecordCount > 0 Then
   TxtNoofopenGodown.Text = TmpRs!NoofGodown
Else
   TxtNoofopenGodown.Text = ""
End If
End Sub

'Private Sub CmbRequestTo_GotFocus()
'tmp = CmbRequestTo.Text
'Call TableMst_Employee(" Where Flag = 'Y'")
'
'CmbRequestTo.Clear
'If RsMst_Employee.RecordCount = 0 Then
'   MsgBox "No Record found!!!"
'   'CmbRequestTo.SetFocus
'   Exit Sub
'End If
'For I = 1 To RsMst_Employee.RecordCount
'    CmbRequestTo.AddItem RsMst_Employee!EmployeeName & "~" & RsMst_Employee!EmployeeNo
'    RsMst_Employee.MoveNext
'Next
'CmbRequestTo.Text = tmp
'
'End Sub
'
'Private Sub CmbRequestTo_LostFocus()
'If CmbRequestTo.Text = "" Then Exit Sub
'
'Dim strSplitString() As String
'Dim EmpNo_, EmpName_ As String
'
'strSplitString = Split(CmbRequestTo, "~")
'If UBound(strSplitString) = 0 Then
'   MsgBox "Invalid Selection!!!"
'   CmbRequestTo.SetFocus
'   Exit Sub
'End If
'EmpName_ = Trim(strSplitString(0))
'EmpNo_ = Trim(strSplitString(1))
'
'RsMst_Employee.MoveFirst
'RsMst_Employee.Find "EmployeeNo = '" & EmpNo_ & "'"
'
'If RsMst_Employee.EOF Then
'   MsgBox "Invalid Selection!!!"
'   CmbRequestTo.SetFocus
'   Exit Sub
'End If
'mRequestToID = RsMst_Employee!EmployeeID
'
'End Sub

Private Sub CmbRequestTo_GotFocus()
tmp1 = CmbRequestTo.Text
Call TableMst_Employee("Where  Flag = 'Y' ")
CmbRequestTo.Clear

tmp = "Select DirectReportingTo from Mst_Employee Where EmployeeID=" & Gen_EmployeeID
Call Tmp2Table
If TmpRs2.RecordCount > 0 Then
   Call UpperHierarchy(TmpRs2!DirectReportingTo)
End If

tmp = "Select EmployeeID,EmployeeName ,EmployeeNo  from Mst_Employee Where DesignationID=1 And Flag='Y'"
Call Tmp2Table

For I = 1 To TmpRs2.RecordCount
    CmbRequestTo.AddItem TmpRs2!EmployeeName & "~" & TmpRs2!EmployeeNo
    TmpRs2.MoveNext
Next

CmbRequestTo.Text = tmp1
End Sub

Private Sub CmbRequestTo_LostFocus()
Dim strSplitString() As String
Dim EmpNo_, EmpName_ As String

If CmbRequestTo.Text = "" Then Exit Sub

strSplitString = Split(CmbRequestTo, "~")
If UBound(strSplitString) < 1 Then
   MsgBox "Invalid Selection "
   Exit Sub
End If
EmpName_ = Trim(strSplitString(0))
EmpNo_ = Trim(strSplitString(1))

RsMst_Employee.MoveFirst
RsMst_Employee.Find "EmployeeNo = '" & EmpNo_ & "'"

If RsMst_Employee.EOF Then
   MsgBox "Invalid Selection "
   CmbRequestTo.SetFocus
   Exit Sub
End If
mRequestToID = RsMst_Employee!EmployeeID
'TxtFWDDesignation = GetDesignaion(RsMst_Employee!DesignationID)
'mForwardToEmailID = RsMst_Employee!EmployeeEmailID
End Sub

Private Sub UpperHierarchy(mEmployeeID_ As Variant)
tmp = "Select EmployeeID,EmployeeName + '~' + EmployeeNo As EmployeeName,DirectReportingTo From Mst_Employee Where EmployeeID=" & mEmployeeID_ & " And Flag='Y' And DesignationID<>1"
Call TmpTable
If TmpRs.RecordCount > 0 Then
   If Not IsNull(TmpRs!DirectReportingTo) Then
      CmbRequestTo.AddItem TmpRs!EmployeeName
      Call UpperHierarchy(TmpRs!DirectReportingTo)
   End If
End If
End Sub

Private Sub CmbSecurityAgency_GotFocus()
If CmbLocation.Text = "" Then
   MsgBox "Select Location First!!!"
   CmbLocation.SetFocus
   Exit Sub
End If
tmp = CmbSecurityAgency.Text

'Call TableMst_SecurityAgency("Where Flag = 'Y'")
Call TableMst_SecurityAgency("Where Flag = 'Y' and AgencyID in (Select Distinct MSA.SecurityAgencyID " & _
" From Txn_SecurityAgreement TSA " & _
" Inner join Mst_SecurityAgency MSA on TSA.SecurityAgencyID=MSA.SecurityAgencyID " & _
" Inner join Mst_State MS on TSA.StateID=MS.StateID " & _
" Where TSA.Flag='A' And MS.StateID=" & mStateID & ")")

CmbSecurityAgency.Clear
If RsMst_SecurityAgency.RecordCount = 0 Then
   MsgBox "No Record found!!!"
   'CmbSecurityAgency.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_SecurityAgency.RecordCount
    CmbSecurityAgency.AddItem RsMst_SecurityAgency!AgencyName
    RsMst_SecurityAgency.MoveNext
Next
CmbSecurityAgency.Text = tmp
End Sub

Private Sub CmbSecurityAgency_LostFocus()
'If CmbSecurityAgency.Text = "" Then
'    TxtNoofGodownsMapped.Text = ""
'    TxtNoofGuards.Text = ""
'    TxtTotalCost.Text = ""
'    Call Grid_HeadGodownInfo
'    Exit Sub
'End If
   
If CmbSecurityAgency.Text = "" Then Exit Sub

RsMst_SecurityAgency.MoveFirst
RsMst_SecurityAgency.Find "AgencyName = '" & CmbSecurityAgency.Text & "'"

If RsMst_SecurityAgency.EOF Then
   MsgBox "Invalid Selection!!!"
   CmbSecurityAgency.SetFocus
   Exit Sub
End If
mSecurityAgencyID = RsMst_SecurityAgency!SecurityAgencyID
Call GetAgreementStartDate
'Call GetGodownCostDetails
'Call GetGodownInfoDetails
End Sub

'Private Sub GetAgreementStartDate()
'tmp = " Select AgreementStartDate, AgreementEndDate From Mst_SecLocMapping Where LocationID = " & mLocationID
'tmp = tmp & " And SecurityAgencyID = " & mSecurityAgencyID
'Call TmpTable
'If TmpRs.RecordCount > 0 Then
'    TxtAgreeStartDate.Value = IIf(IsNull(TmpRs!AgreementStartDate), Date - 30, TmpRs!AgreementStartDate)
'    TxtAgreeEndDate.Value = IIf(IsNull(TmpRs!AgreementEndDate), Date + 30, TmpRs!AgreementEndDate)
'End If
'End Sub

Private Sub GetGodownInfoDetails()
Call Grid_HeadGodownInfo
tmp = " select "
'tmp = tmp & " MSLGMD.GodMapDID, MSS.SecurityShifts, MCMP.CMPName, MG.GodownNo, MG.Address, MSLGMD.BalanceBags,"
'tmp = tmp & " MSLGMD.BalanceQty , MSLGMD.AUM, MSLGMD.Remarks, MSLGMD.SecurityTypeID, MSLGMD.CMPID, MSLGMD.GodownID"
'tmp = tmp & " from Mst_SecLocGodMappingDetail MSLGMD"
'tmp = tmp & " Inner Join Mst_SecLocMapping MSLM  on MSLM.TxnID = MSLGMD.TxnID"
'tmp = tmp & " Inner Join Mst_securityShifts MSS on MSS.SecurityShiftsID = MSLGMD.SecurityTypeID"
'tmp = tmp & " Inner Join Mst_CMP MCMP on MCMP.CMPID =  MSLGMD.CMPID"
'tmp = tmp & " Inner Join Mst_Godown MG on MG.GodownID =  MSLGMD.GodownID"
'tmp = tmp & " Where MSLM.LocatioNID = " & mLocationID & " And MSLM.SecurityAgencyID  = " & mSecurityAgencyID
'Call TmpTable

tmp = " select "
tmp = tmp & " MSLGMD.GodMapDID, MSS.SecurityShifts, MCMP.CMPName, MG.GodownNo, MG.Address, MSLGMD.BalanceBags, MSD.SecurityDesignation,"
tmp = tmp & " MSLGMD.BalanceQty , MSLGMD.AUM, MSLGMD.Remarks, MSLGMD.TxnDetailID, MSLGMD.CMPID, MSLGMD.GodownID" 'MSLGMD.SecurityTypeID,
tmp = tmp & " from Mst_SecLocGodMappingDetail MSLGMD"
tmp = tmp & " Inner Join Mst_SecLocMapping MSLM  on MSLM.TxnID = MSLGMD.TxnID"
tmp = tmp & " Inner Join Mst_SecLocMappingDetail MSLMD on MSLMD.TxnDetailID = MSLGMD.TxnDetailID"
tmp = tmp & " Inner Join Mst_securityShifts MSS on MSS.SecurityShiftsID = MSLMD.SecurityTypeID"
tmp = tmp & " Inner Join Mst_securityDesignation MSD on MSD.SecurityDesignationID = MSLMD.SecurityDesignationID"
tmp = tmp & " Inner Join Mst_CMP MCMP on MCMP.CMPID =  MSLGMD.CMPID"
tmp = tmp & " Inner Join Mst_Godown MG on MG.GodownID =  MSLGMD.GodownID"
tmp = tmp & " Where MSLM.LocatioNID = " & mLocationID
'tmp = tmp & " And MSLM.SecurityAgencyID  = " & mSecurityAgencyID
Call TmpTable


If TmpRs.RecordCount > 0 Then
    MSHFlexGrid3.Rows = MSHFlexGrid3.Rows + TmpRs.RecordCount
    With MSHFlexGrid3
        For I = 2 To TmpRs.RecordCount + 1
            '.TextMatrix(I, 0) = IIf(IsNull(TmpRs!SecurityShifts), "", TmpRs!SecurityShifts)
            .TextMatrix(I, 0) = IIf(IsNull(TmpRs!SecurityShifts), "", TmpRs!SecurityShifts) & "~" & IIf(IsNull(TmpRs!TxnDetailID), "", TmpRs!TxnDetailID)
            .TextMatrix(I, 1) = IIf(IsNull(TmpRs!CMPName), "", TmpRs!CMPName)
            .TextMatrix(I, 2) = IIf(IsNull(TmpRs!GodownNo), "", TmpRs!GodownNo)
            .TextMatrix(I, 3) = IIf(IsNull(TmpRs!Address), "", TmpRs!Address)
            .TextMatrix(I, 4) = IIf(IsNull(TmpRs!BalanceBags), "", TmpRs!BalanceBags)
            .TextMatrix(I, 5) = IIf(IsNull(TmpRs!BalanceQty), "", TmpRs!BalanceQty)
            .TextMatrix(I, 6) = IIf(IsNull(TmpRs!AUM), "", TmpRs!AUM)
            .TextMatrix(I, 7) = IIf(IsNull(TmpRs!Remarks), "", TmpRs!Remarks)
            '.TextMatrix(I, 8) = IIf(IsNull(TmpRs!SecurityTypeID), 0, TmpRs!SecurityTypeID)
            .TextMatrix(I, 8) = IIf(IsNull(TmpRs!TxnDetailID), 0, TmpRs!TxnDetailID)
            .TextMatrix(I, 9) = IIf(IsNull(TmpRs!CMPID), 0, TmpRs!CMPID)
            .TextMatrix(I, 10) = IIf(IsNull(TmpRs!GodownID), 0, TmpRs!GodownID)
            .TextMatrix(I, 11) = IIf(IsNull(TmpRs!GodMapDID), 0, TmpRs!GodMapDID)
            .TextMatrix(I, 12) = IIf(IsNull(TmpRs!SecurityDesignation), 0, TmpRs!SecurityDesignation)
            .TextMatrix(0, 4) = Val(.TextMatrix(I, 4)) + Val(.TextMatrix(0, 4))
            .TextMatrix(0, 5) = Val(.TextMatrix(I, 5)) + Val(.TextMatrix(0, 5))
            .TextMatrix(0, 6) = Val(.TextMatrix(I, 6)) + Val(.TextMatrix(0, 6))
            TmpRs.MoveNext
        Next
    End With
End If
End Sub

Private Sub GetGodownCostDetails()
tmp = " Select Count(MSLMD.TxnDetailID) NoofGuards, sum(MSLMD.RatePerMonth) TotalCost From Mst_SecLocMappingDetail MSLMD" 'Sum(NoofGodown) NoofGodownMapped,
tmp = tmp & " Inner Join Mst_SecLocMapping MSLM on MSLM.TxnID = MSLMD.TxnID"
tmp = tmp & " Where MSLM.LocationID =  " & mLocationID
'tmp = tmp & " And MSLM.SecurityAgencyID = " & mSecurityAgencyID
Call TmpTable

If TmpRs.RecordCount > 0 Then
    TxtNoofGuards.Text = IIf(IsNull(TmpRs!NoofGuards), 0, TmpRs!NoofGuards)
    TxtTotalCost.Text = IIf(IsNull(TmpRs!TotalCost), 0, TmpRs!TotalCost)
Else
    TxtNoofGuards.Text = ""
    TxtTotalCost.Text = ""
End If

tmp = " Select"
tmp = tmp & " Count(Distinct MSLGMD.GodownID) NoofGodownMapped"
tmp = tmp & " From Mst_SecLocGodMappingDetail MSLGMD"
tmp = tmp & " Inner Join Mst_SecLocMapping MSLM on MSLM.TxnID = MSLGMD.TxnID"
tmp = tmp & " Where MSLM.LocationID =  " & mLocationID
'tmp = tmp & "  And MSLM.SecurityAgencyID = " & mSecurityAgencyID
Call TmpTable

If TmpRs.RecordCount > 0 Then
    TxtNoofGodownsMapped.Text = IIf(IsNull(TmpRs!NoofGodownMapped), "", TmpRs!NoofGodownMapped)
Else
    TxtNoofGodownsMapped.Text = ""
End If
End Sub

Private Sub CmbSecurityDesignation_GotFocus()
'tmp = CmbSecurityDesignation.Text
'Call TableMst_SecurityDesignation("")
'
'CmbSecurityDesignation.Clear
'If RsMst_SecurityDesignation.RecordCount = 0 Then
'   MsgBox "No Record found!!!"
'   'CmbSecurityType.SetFocus
'   Exit Sub
'End If
'For i = 1 To RsMst_SecurityDesignation.RecordCount
'    CmbSecurityDesignation.AddItem RsMst_SecurityDesignation!SecurityDesignation
'    RsMst_SecurityDesignation.MoveNext
'Next
'CmbSecurityDesignation.Text = tmp

If CmbSecurityType.Text = "" Then
   MsgBox "Please select security Type!!"
   CmbSecurityType.SetFocus
   Exit Sub
End If
tmp1 = CmbSecurityDesignation.Text
'Call Mst_SecurityDesignation("")

CmbSecurityDesignation.Clear

tmp = "Select distinct MSS.SecurityDesignationID,SecurityDesignation From Txn_SecurityAgreementDetail TSAD"
tmp = tmp & " Inner join Mst_SecurityDesignation MSS on MSS.SecurityDesignationID=TSAD.SecurityDesignationID"
tmp = tmp & " Where isnull(MSS.Flag,'A' )= 'A' And TSAD.SAID=" & mSAID & " And TSAD.SecurityTypeID=" & mSecurityTypeID

Call TmpTable

If TmpRs.RecordCount = 0 Then
   MsgBox "No Record found!!!"
   'CmbSecurityType.SetFocus
   Exit Sub
End If
For I = 1 To TmpRs.RecordCount
    CmbSecurityDesignation.AddItem TmpRs!SecurityDesignation
    TmpRs.MoveNext
Next
CmbSecurityDesignation.Text = tmp1


End Sub

Private Sub CmbSecurityDesignation_LostFocus()
'If CmbSecurityDesignation.Text = "" Then Exit Sub
'
'RsMst_SecurityDesignation.MoveFirst
'RsMst_SecurityDesignation.Find "SecurityDesignation = '" & CmbSecurityDesignation.Text & "'"
'
'If RsMst_SecurityDesignation.EOF Then
'   MsgBox "Invalid Selection!!!"
'   CmbSecurityDesignation.SetFocus
'   Exit Sub
'End If
'mSecurityDesignationID = RsMst_SecurityDesignation!SecurityDesignationID


If CmbSecurityDesignation.Text = "" Then Exit Sub
      
TmpRs.MoveFirst
TmpRs.Find "SecurityDesignation = '" & CmbSecurityDesignation.Text & "'"

If TmpRs.EOF Then
   MsgBox "Invalid Selection!!!"
   CmbSecurityDesignation.SetFocus
   Exit Sub
End If
mSecurityDesignationID = TmpRs!SecurityDesignationID

If CmbSecurityType.Text <> "" Then
   TxtRatePerMonth = GetSecRate()
End If


End Sub

Private Function GetSecRate() As Variant
Dim Retval As Variant

tmp = "select MonthlyRate From Txn_SecurityAgreementDetail Where SAID=" & mSAID & " And SecurityTypeID=" & mSecurityTypeID & " And"
tmp = tmp & " SecurityDesignationID=" & mSecurityDesignationID

Call Tmp3Table

Retval = 0

If TmpRs3.RecordCount > 0 Then
   Retval = TmpRs3!MonthlyRate
End If
GetSecRate = Retval
End Function

Private Sub CmbSecurityType_GotFocus()
'tmp = CmbSecurityType.Text
'Call TableMst_SecurityShifts("")
'
'CmbSecurityType.Clear
'If RsMst_SecurityShifts.RecordCount = 0 Then
'   MsgBox "No Record found!!!"
'   'CmbSecurityType.SetFocus
'   Exit Sub
'End If
'For i = 1 To RsMst_SecurityShifts.RecordCount
'    CmbSecurityType.AddItem RsMst_SecurityShifts!SecurityShifts
'    RsMst_SecurityShifts.MoveNext
'Next
'CmbSecurityType.Text = tmp


If CmbSecurityAgency.Text = "" Then
   MsgBox "Select Security Agency First!!!"
   CmbSecurityAgency.SetFocus
   Exit Sub
End If

tmp1 = CmbSecurityType.Text
'Call Mst_SecurityShiftsTable("")

tmp = "Select distinct SecurityShiftsID,SecurityShifts From Txn_SecurityAgreementDetail TSAD"
tmp = tmp & " Inner join Mst_SecurityShifts MSS on MSS.SecurityShiftsID=TSAD.SecurityTypeID"
tmp = tmp & " Where TSAD.SAID=" & mSAID

Call TmpTable

CmbSecurityType.Clear

If TmpRs.RecordCount = 0 Then
   MsgBox "No Record found!!!"
   'CmbSecurityType.SetFocus
   Exit Sub
End If
For I = 1 To TmpRs.RecordCount
    CmbSecurityType.AddItem TmpRs!SecurityShifts
    TmpRs.MoveNext
Next
CmbSecurityType.Text = tmp1


End Sub

Private Sub CmbSecurityType_LostFocus()

'If CmbSecurityType.Text = "" Then Exit Sub
'
'RsMst_SecurityShifts.MoveFirst
'RsMst_SecurityShifts.Find "SecurityShifts = '" & CmbSecurityType.Text & "'"
'
'If RsMst_SecurityShifts.EOF Then
'   MsgBox "Invalid Selection!!!"
'   CmbSecurityType.SetFocus
'   Exit Sub
'End If
'mSecurityTypeID = RsMst_SecurityShifts!SecurityShiftsID


If CmbSecurityType.Text = "" Then Exit Sub

TmpRs.MoveFirst
TmpRs.Find "SecurityShifts = '" & CmbSecurityType.Text & "'"

If TmpRs.EOF Then
   MsgBox "Invalid Selection!!!"
   CmbSecurityType.SetFocus
   Exit Sub
End If
If TmpRs!SecurityShiftsID <> mSecurityTypeID Then
   CmbSecurityDesignation.Text = ""
   mSecurityDesignationID = 0
End If

mSecurityTypeID = TmpRs!SecurityShiftsID
If CmbSecurityDesignation.Text <> "" Then
   TxtRatePerMonth = GetSecRate()
End If



End Sub

Private Sub CmbSLocation_GotFocus()
tmp = CmbSLocation.Text
If Gen_WcLocation = "" Then
    Call TableMst_Location
Else
    Call TableMst_Location(" WHERE " & Gen_WcLocation)
End If
CmbSLocation.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Record found!!!"
   'CmbSLocation.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbSLocation.AddItem RsMst_Location!Locationname
    RsMst_Location.MoveNext
Next
CmbSLocation.Text = tmp

End Sub

Private Sub CmbSLocation_LostFocus()
If CmbSLocation.Text = "" Then Exit Sub

RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbSLocation.Text & "'"

If RsMst_Location.EOF Then
    MsgBox "Invalid Selection!!!"
    CmbSLocation.SetFocus
    Exit Sub
End If

mLocationID = RsMst_Location!LocationID

End Sub

Private Sub CmbSRequestTo_GotFocus()
tmp = CmbSRequestTo.Text
Call TableMst_Employee("")

CmbSRequestTo.Clear
If RsMst_Employee.RecordCount = 0 Then
   MsgBox "No Record found!!!"
   'CmbSRequestTo.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Employee.RecordCount
    CmbSRequestTo.AddItem RsMst_Employee!EmployeeName & "~" & RsMst_Employee!EmployeeNo
    RsMst_Employee.MoveNext
Next
CmbSRequestTo.Text = tmp
End Sub

Private Sub CmbSRequestTo_LostFocus()
If CmbSRequestTo.Text = "" Then Exit Sub

Dim strSplitString() As String
Dim EmpNo_, EmpName_ As String

strSplitString = Split(CmbSRequestTo, "~")
If UBound(strSplitString) = 0 Then
   MsgBox "Invalid Selection!!!"
   CmbSRequestTo.SetFocus
   Exit Sub
End If
EmpName_ = Trim(strSplitString(0))
EmpNo_ = Trim(strSplitString(1))

RsMst_Employee.MoveFirst
RsMst_Employee.Find "EmployeeNo = '" & EmpNo_ & "'"

If RsMst_Employee.EOF Then
   MsgBox "Invalid Selection!!!"
   CmbSRequestTo.SetFocus
   Exit Sub
End If
mRequestToID = RsMst_Employee!EmployeeID
End Sub

Private Sub CmbSSecurityAgency_GotFocus()
tmp = CmbSSecurityAgency.Text
Call TableMst_SecurityAgency("Where Flag = 'Y'")
CmbSSecurityAgency.Clear
If RsMst_SecurityAgency.RecordCount = 0 Then
   MsgBox "No Record found!!!"
   'CmbSSecurityAgency.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_SecurityAgency.RecordCount
    CmbSSecurityAgency.AddItem RsMst_SecurityAgency!AgencyName
    RsMst_SecurityAgency.MoveNext
Next
CmbSSecurityAgency.Text = tmp
End Sub

Private Sub CmbSSecurityAgency_LostFocus()
If CmbSSecurityAgency.Text = "" Then Exit Sub
   
RsMst_SecurityAgency.MoveFirst
RsMst_SecurityAgency.Find "AgencyName = '" & CmbSSecurityAgency.Text & "'"

If RsMst_SecurityAgency.EOF Then
   MsgBox "Invalid Selection!!!"
   CmbSSecurityAgency.SetFocus
   Exit Sub
End If
mSecurityAgencyID = RsMst_SecurityAgency!SecurityAgencyID

End Sub

Private Sub CmbSSecurityDesignation_GotFocus()
tmp = CmbSSecurityDesignation.Text
Call TableMst_SecurityDesignation("")

CmbSSecurityDesignation.Clear
If RsMst_SecurityDesignation.RecordCount = 0 Then
   MsgBox "No Record found!!!"
   'CmbSSecurityType.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_SecurityDesignation.RecordCount
    CmbSSecurityDesignation.AddItem RsMst_SecurityDesignation!SecurityDesignation
    RsMst_SecurityDesignation.MoveNext
Next
CmbSSecurityDesignation.Text = tmp

End Sub

Private Sub CmbSSecurityDesignation_LostFocus()
If CmbSSecurityDesignation.Text = "" Then Exit Sub
      
RsMst_SecurityDesignation.MoveFirst
RsMst_SecurityDesignation.Find "SecurityDesignation = '" & CmbSSecurityDesignation.Text & "'"

If RsMst_SecurityDesignation.EOF Then
   MsgBox "Invalid Selection!!!"
   CmbSSecurityDesignation.SetFocus
   Exit Sub
End If
mSecurityDesignationID = RsMst_SecurityDesignation!SecurityDesignationID
End Sub

Private Sub CmbSSecurityType_GotFocus()
tmp = CmbSSecurityType.Text
Call TableMst_SecurityShifts("")

CmbSSecurityType.Clear
If RsMst_SecurityShifts.RecordCount = 0 Then
   MsgBox "No Record found!!!"
   'CmbSSecurityType.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_SecurityShifts.RecordCount
    CmbSSecurityType.AddItem RsMst_SecurityShifts!SecurityShifts
    RsMst_SecurityShifts.MoveNext
Next
CmbSSecurityType.Text = tmp

End Sub

Private Sub CmbSSecurityType_LostFocus()
If CmbSSecurityType.Text = "" Then Exit Sub

RsMst_SecurityShifts.MoveFirst
RsMst_SecurityShifts.Find "SecurityShifts = '" & CmbSSecurityType.Text & "'"

If RsMst_SecurityShifts.EOF Then
   MsgBox "Invalid Selection!!!"
   CmbSSecurityType.SetFocus
   Exit Sub
End If
mSecurityTypeID = RsMst_SecurityShifts!SecurityShiftsID
End Sub

Private Sub CmbSStatus_LostFocus()
If Trim(CmbSStatus.Text) = "" Then Exit Sub

If LCase(CmbSStatus.Text) <> "pending" And LCase(CmbSStatus.Text) <> "approved" And LCase(CmbSStatus.Text) <> "rejected" Then
    MsgBox "Invalid Selection!!!"
    CmbSStatus.SetFocus
End If
End Sub

Private Sub CmbStatus_LostFocus()
If Trim(CmbStatus.Text) = "" Then Exit Sub

If LCase(CmbStatus.Text) <> "pending" And LCase(CmbStatus.Text) <> "approved" And LCase(CmbStatus.Text) <> "rejected" Then
    MsgBox "Invalid Selection!!!"
    CmbStatus.SetFocus
End If
If LCase(CmbStatus.Text) = "pending" Then
    TxtApproveRemarks.Locked = True
    TxtApproveRemarks.Text = ""
Else
    If LCase(Gen_UserRole) <> "headcmg" And LCase(Gen_UserRole) <> "admin" And LCase(Gen_UserRole) <> "superadmin" And LCase(Gen_UserRole) <> "power" Then
    'If LCase(CmbStatus.Text) = "approved" Or LCase(CmbStatus.Text) = "rejected" Then
        MsgBox "Please forward this claim to CMG user, Only CMG user have access to approve!!!"
        CmbStatus.SetFocus
        Exit Sub
    'End If
    End If

    TxtApproveRemarks.Locked = False
End If
End Sub

Private Sub CmdAddtoList_Click()
If CmbCMP.Text = "" Then
    MsgBox "Blank CMP not Allowed!!!"
    CmbCMP.SetFocus
    Exit Sub
End If
If CmbGodown.Text = "" Then
    MsgBox "Blank Godown not Allowed!!!"
    CmbGodown.SetFocus
    Exit Sub
End If
For I = 2 To MSHFlexGrid2.Rows - 2
    If CmdAddtoList.Caption <> "Add to List" Then
        If I <> MSHFlexGrid2.RowSel Then
            If MSHFlexGrid2.TextMatrix(I, 0) = CmbCMP.Text And MSHFlexGrid2.TextMatrix(I, 1) = CmbGodown.Text Then
                MsgBox "Duplicate Godown not Allowed!!!"
                Exit Sub
            End If
        End If
    Else
        If MSHFlexGrid2.TextMatrix(I, 0) = CmbCMP.Text And MSHFlexGrid2.TextMatrix(I, 1) = CmbGodown.Text Then
            MsgBox "Duplicate Godown not Allowed!!!"
            Exit Sub
        End If
    End If
Next

If CmdAddtoList.Caption <> "Add to List" Then
    I = MSHFlexGrid2.RowSel
    MSHFlexGrid2.TextMatrix(0, 3) = Val(MSHFlexGrid2.TextMatrix(0, 3)) - Val(MSHFlexGrid2.TextMatrix(I, 3))
    MSHFlexGrid2.TextMatrix(0, 4) = Val(MSHFlexGrid2.TextMatrix(0, 4)) - Val(MSHFlexGrid2.TextMatrix(I, 4))
    MSHFlexGrid2.TextMatrix(0, 5) = Val(MSHFlexGrid2.TextMatrix(0, 5)) - Val(MSHFlexGrid2.TextMatrix(I, 5))
Else
    I = MSHFlexGrid2.Rows - 1
    MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
End If

MSHFlexGrid2.TextMatrix(I, 0) = CmbCMP.Text
MSHFlexGrid2.TextMatrix(I, 1) = CmbGodown.Text
MSHFlexGrid2.TextMatrix(I, 2) = TxtGodownAddress.Text
MSHFlexGrid2.TextMatrix(I, 3) = TxtBalanceBags.Text
MSHFlexGrid2.TextMatrix(I, 4) = TxtBalanceQty.Text
MSHFlexGrid2.TextMatrix(I, 5) = TxtAUM.Text
MSHFlexGrid2.TextMatrix(I, 6) = TxtRemarksD.Text
MSHFlexGrid2.TextMatrix(I, 7) = mCMPID
MSHFlexGrid2.TextMatrix(I, 8) = mGodownID

MSHFlexGrid2.TextMatrix(0, 3) = Val(MSHFlexGrid2.TextMatrix(0, 3)) + Val(MSHFlexGrid2.TextMatrix(I, 3))
MSHFlexGrid2.TextMatrix(0, 4) = Val(MSHFlexGrid2.TextMatrix(0, 4)) + Val(MSHFlexGrid2.TextMatrix(I, 4))
MSHFlexGrid2.TextMatrix(0, 5) = Val(MSHFlexGrid2.TextMatrix(0, 5)) + Val(MSHFlexGrid2.TextMatrix(I, 5))
TxtTotalAUM.Text = MSHFlexGrid2.TextMatrix(0, 5)
Call ClearGodownMapping
End Sub

Private Sub CmdDeleteFromList_Click()
If CmbCMP.Text = "" Then Exit Sub

I = MSHFlexGrid2.RowSel
If I > 0 Then
   If MSHFlexGrid2.TextMatrix(I, 1) <> "" Then
      
      MSHFlexGrid2.TextMatrix(0, 3) = Val(MSHFlexGrid2.TextMatrix(0, 3)) - Val(MSHFlexGrid2.TextMatrix(I, 3))
      MSHFlexGrid2.TextMatrix(0, 4) = Val(MSHFlexGrid2.TextMatrix(0, 4)) - Val(MSHFlexGrid2.TextMatrix(I, 4))
      MSHFlexGrid2.TextMatrix(0, 5) = Val(MSHFlexGrid2.TextMatrix(0, 5)) - Val(MSHFlexGrid2.TextMatrix(I, 5))
      For RO = I To MSHFlexGrid2.Rows - 2
          For C = 0 To MSHFlexGrid2.Cols - 1
              MSHFlexGrid2.TextMatrix(RO, C) = MSHFlexGrid2.TextMatrix(RO + 1, C)
          Next
      Next
      MSHFlexGrid2.Rows = MSHFlexGrid2.Rows - 1
   End If
End If
TxtTotalAUM.Text = MSHFlexGrid2.TextMatrix(0, 5)
Call ClearGodownMapping
End Sub

Private Sub Grid_Head1()
CmdExcel.Enabled = False
Label22.Caption = ""
With MSHFlexGrid1
    .Clear
    .Rows = 2
    .Cols = 18
    
    .TextMatrix(0, 0) = "Request ID"
    .TextMatrix(0, 1) = "Request Date"
    .TextMatrix(0, 2) = "Location"
    .TextMatrix(0, 3) = "Request To"
    .TextMatrix(0, 4) = "Security Agency"
    .TextMatrix(0, 5) = "Agreement Start Date"
    .TextMatrix(0, 6) = "Agreement End Date"
    .TextMatrix(0, 7) = "Salary Payable Upto"
    .TextMatrix(0, 8) = "No Of Open Godowns"
    .TextMatrix(0, 9) = "No Of Guards"
    .TextMatrix(0, 10) = "No Of Godowns Mapped"
    .TextMatrix(0, 11) = "Total Cost"
    .TextMatrix(0, 12) = "Total AUM"
    .TextMatrix(0, 13) = "Security Type"
    .TextMatrix(0, 14) = "Security Designation"
    .TextMatrix(0, 15) = "Rate Per Month"
    .TextMatrix(0, 16) = "Status"
    .TextMatrix(0, 17) = "Remarks"
        
    .ColWidth(0) = 1000
    .ColWidth(1) = 1500
    .ColWidth(2) = 1900
    .ColWidth(3) = 2000
    .ColWidth(4) = 2000
    .ColWidth(5) = 1500
    .ColWidth(6) = 1500
    .ColWidth(7) = 1500
    .ColWidth(8) = 1500
    .ColWidth(9) = 1500
    .ColWidth(10) = 1500
    .ColWidth(11) = 1500
    .ColWidth(12) = 1500
    .ColWidth(13) = 2200
    .ColWidth(14) = 2000
    .ColWidth(15) = 1500
    .ColWidth(16) = 1500
    .ColWidth(17) = 3500
    
    .RowHeight(0) = 650
    .WordWrap = True
End With
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Mst_SecurityRequest.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Mst_SecurityRequest.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1

oWS.Range("a:a").ColumnWidth = 8
oWS.Range("b:b").ColumnWidth = 9.43
oWS.Range("c:c").ColumnWidth = 15
oWS.Range("d:d").ColumnWidth = 20
oWS.Range("e:e").ColumnWidth = 20
oWS.Range("f:f").ColumnWidth = 10
oWS.Range("g:g").ColumnWidth = 10.14
oWS.Range("h:h").ColumnWidth = 12
oWS.Range("i:i").ColumnWidth = 14
oWS.Range("j:j").ColumnWidth = 11.14
oWS.Range("k:k").ColumnWidth = 12.57
oWS.Range("l:l").ColumnWidth = 11
oWS.Range("m:m").ColumnWidth = 10
oWS.Range("n:n").ColumnWidth = 20
oWS.Range("o:o").ColumnWidth = 20
oWS.Range("p:p").ColumnWidth = 10
oWS.Range("q:q").ColumnWidth = 10
oWS.Range("r:r").ColumnWidth = 20

tmp = "a" & mRow & ":" & "r" & mRow

oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "Security Request"
oWS.Cells(mRow, C + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 14
oWS.Cells(mRow, C + 1).Interior.Color = &HFF9F9F
'&H8080BF

For I = 0 To MSHFlexGrid1.Rows - 2
   mRow = mRow + 1 ' + 6
   For C = 0 To MSHFlexGrid1.Cols - 1
      oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
      If I = 0 Then
         oWS.Cells(mRow, C + 1).Interior.Color = &H55FFFF '&HD9D9FF
         oWS.Cells(mRow, C + 1).Font.Bold = True
         oWS.Cells(mRow, C + 1).RowHeight = 30
      'ElseIf I = 1 Then
      '   oWS.Cells(mRow, c + 1).Interior.Color = &H55FFFF
      '   oWS.Cells(mRow, c + 1).Font.Bold = True
      '   oWS.Cells(mRow, c + 1).RowHeight = 45
      End If
      oWS.Cells(mRow, C + 1).WrapText = True
      oWS.Cells(mRow, C + 1).Borders.Color = RGB(0, 0, 0)
   Next
Next

oWB.Save
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub


Private Sub CmdGetReport_Click()
If CmbLocation.Text = "" Then
    MsgBox "Blank Location not Allowed!!!"
    Exit Sub
End If
Gen_SecLocation = CmbLocation.Text
Gen_SecLocationID = mLocationID
FrmRpt_SecurityRequestGodownMapping.Show
Gen_SecLocation = ""
Gen_SecLocationID = ""
End Sub

Private Sub CmdGo_Click()
Call Grid_Head1
Dim wc As Variant

If Gen_DBType = "MDB" Then
    wc = " Where MSR.RequestDate Between #" & SFrom.Value & "# And #" & STo.Value & "#"
Else
    wc = " Where MSR.RequestDate Between '" & SFrom.Value & "' And '" & STo.Value & "'"
End If
If CmbSLocation.Text <> "" Then
    wc = wc & " And MSR.LocationID = " & mLocationID
Else
    If Gen_WcLocation <> "" Then
        wc = wc & " And MSR." & Gen_WcLocation
    End If
End If
If CmbSSecurityAgency.Text <> "" Then
    wc = wc & " And MSR.SecurityAgencyID = " & mSecurityAgencyID
End If
If CmbSSecurityType.Text <> "" Then
    wc = wc & " And MSR.SecurityTypeID = " & mSecurityTypeID
End If
If CmbSSecurityDesignation.Text <> "" Then
    wc = wc & " And MSR.SecurityDesignationID = " & mSecurityDesignationID
End If
If CmbSRequestTo.Text <> "" Then
    wc = wc & " And MSR.RequestToID= " & mRequestToID
End If
If CmbSStatus.Text <> "" Then
    wc = wc & " And MSR.Status = '" & Left$(CmbSStatus.Text, 1) & "'"
End If

tmp = " select"
tmp = tmp & " MSR.SecurityRequestID, MSR.RequestDate, MSR.LocationID, MSR.RequestToID, MSR.SecurityAgencyID, MSR.AgreeStartDate,"
tmp = tmp & " MSR.AgreeEndDate, MSR.SalaryPayableUpto, MSR.SecurityTypeID, MSR.SecurityDesignationID, MSR.RatePerMonth,"
tmp = tmp & " MSR.Status, MSR.Remarks, MSR.G_UID, MSR.CreatedDate, MSR.CreatedBy, MSR.UpdatedDate, MSR.UpdatedBy,"
tmp = tmp & " ML.LocationName, ME.EmployeeName+'~'+ME.EmployeeNo as RequestTo, MSA.AgencyName, MSS.SecurityShifts, MSD.Securitydesignation"
tmp = tmp & " From Mst_SecurityRequest MSR"
tmp = tmp & " Inner Join Mst_Location ML On ML.LocatioNID = MSR.LocationID"
tmp = tmp & " Inner Join Mst_Employee ME On ME.EmployeeID = MSR.RequestToID"
tmp = tmp & " Inner Join Mst_SecurityAgency MSA ON MSA.SecurityAgencyID = MSR.SecurityAgencyID"
tmp = tmp & " Inner Join Mst_SecurityShifts MSS ON MSS.SecurityShiftsID = MSR.SecurityTypeID"
tmp = tmp & " Inner Join Mst_SecurityDesignation MSD ON MSD.SecurityDesignationID = MSR.SecurityDesignationID"
tmp = tmp & wc
Call TmpTable

If TmpRs.RecordCount > 0 Then
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + TmpRs.RecordCount
    With MSHFlexGrid1
        For I = 1 To TmpRs.RecordCount
            .TextMatrix(I, 0) = IIf(IsNull(TmpRs!SecurityRequestID), 0, TmpRs!SecurityRequestID)
            .TextMatrix(I, 1) = IIf(IsNull(TmpRs!RequestDate), "", Format(TmpRs!RequestDate, "dd-MMM-yyyy"))
            .TextMatrix(I, 2) = IIf(IsNull(TmpRs!Locationname), "", TmpRs!Locationname)
            .TextMatrix(I, 3) = IIf(IsNull(TmpRs!RequestTo), "", TmpRs!RequestTo)
            .TextMatrix(I, 4) = IIf(IsNull(TmpRs!AgencyName), "", TmpRs!AgencyName)
            .TextMatrix(I, 5) = IIf(IsNull(TmpRs!AgreeStartDate), "", Format(TmpRs!AgreeStartDate, "dd-MMM-yyyy"))
            .TextMatrix(I, 6) = IIf(IsNull(TmpRs!AgreeEndDate), "", Format(TmpRs!AgreeEndDate, "dd-MMM-yyyy"))
            .TextMatrix(I, 7) = IIf(IsNull(TmpRs!SalaryPayableUpto), "", Format(TmpRs!SalaryPayableUpto, "dd-MMM-yyyy"))
            .TextMatrix(I, 8) = "No Of Open Godowns"
            .TextMatrix(I, 9) = "No Of Guards"
            .TextMatrix(I, 10) = "No Of Godowns Mapped"
            .TextMatrix(I, 11) = "Total Cost"
            .TextMatrix(I, 12) = "Total AUM"
            .TextMatrix(I, 13) = IIf(IsNull(TmpRs!SecurityShifts), "", TmpRs!SecurityShifts)
            .TextMatrix(I, 14) = IIf(IsNull(TmpRs!SecurityDesignation), "", TmpRs!SecurityDesignation)
            .TextMatrix(I, 15) = IIf(IsNull(TmpRs!RatePerMonth), "", TmpRs!RatePerMonth)
            .TextMatrix(I, 16) = IIf(IsNull(TmpRs!Status), "", TmpRs!Status)
            If LCase(.TextMatrix(I, 16)) = "r" Then
                .TextMatrix(I, 16) = "Rejected"
            ElseIf LCase(.TextMatrix(I, 16)) = "a" Then
                .TextMatrix(I, 16) = "Approved"
            Else
                .TextMatrix(I, 16) = "Pending"
            End If
            .TextMatrix(I, 17) = IIf(IsNull(TmpRs!Remarks), "", TmpRs!Remarks)
            TmpRs.MoveNext
        Next
    End With
    CmdExcel.Enabled = True
    Label22.Caption = TmpRs.RecordCount
End If
End Sub

Private Sub CmdLocation_Click()
FrmMst_Location.Show
End Sub

Private Sub CmdRequestTo_Click()
FrmMst_Employee.Show
End Sub

Private Sub CmdSecurityAgency_Click()
FrmMst_SecurityAgency.Show
End Sub

Private Sub CmdSecurityDesignation_Click()
FrmMst_SecurityDesignation.Show
End Sub

Private Sub CmdSecurityType_Click()
FrmMst_SecurityShifts.Show
End Sub


Private Sub Command1_Click()

End Sub

Private Sub DeleteCmd_Click()
Dim ans As Variant
On Error GoTo msgError

If TxtRequestID.Text = "" Then Exit Sub
If LCase(TxtStatus.Text) <> "pending" Then
    MsgBox "Security Request is " & TxtStatus & ", You can not Delete!!!"
    Exit Sub
End If
'tmp = "SELECT MonthwiseTxnID FROM Txn_MonthwiseSecLoc WHERE SecurityAgencyID = " & mSecurityAgencyID & " And LocationID = " & mLocationID
'Call TmpTable'

'If TmpRs.RecordCount > 0 Then
'    MsgBox "Month wise details for Security Agency '" & TxtSecurityAgency.Text & "' and Location '" & TxtLocation.Text & "' already exists."
'    Exit Sub
'End If

ans = MsgBox("Do you really want to DELETE this record???", vbYesNo)
If ans = vbYes Then
   tmp = "Delete from Mst_SecurityRequestGodownMapping where SecurityRequestID = " & Val(TxtRequestID)
   Call TmpTable
   RsMst_SecurityRequest.Delete
   RsMst_SecurityRequest.Update
   If RsMst_SecurityRequest.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_SecurityRequest.MoveNext
   If RsMst_SecurityRequest.EOF() Then
      RsMst_SecurityRequest.MoveLast
   End If
   Call Indata
End If
Exit Sub
msgError:
MsgBox "Child record Present "
RsMst_SecurityRequest.CancelUpdate
End Sub

Private Sub EditCmd_Click()
If LCase(TxtStatus.Text) <> "pending" Then
    MsgBox "Security Request is " & TxtStatus & ", You can not Edit!!!"
    Exit Sub
End If
Edit_Flg = "E"
Frame1.Enabled = True
Frame3.Enabled = True
Call GButtonLock(Me, False)

'Call ComboVisible(False)

CmbRequestTo.Visible = True
TxtRequestTo.Visible = False

CmbStatus.Visible = True
TxtStatus.Visible = False

End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub FirstCmd_Click()
RsMst_SecurityRequest.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub

Private Sub Form_Load()
'Call Grid_HeadGodownInfo
'Call Grid_HeadGodownMapping

Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Call GButtonLockAD(Me)
   MsgBox "Access Denied !!!!!!!!!"
   Exit Sub
End If
Call Grid_Head1

SFrom.Value = CDate(From_GenDate) - 900
STo.Value = Date 'To_GenDate
CmbSSecurityAgency.Text = ""
CmbSLocation.Text = ""
Label22.Caption = ""

Frame4.Enabled = True
If Gen_WcLocation <> "" Then
    Call TableMst_SecurityRequest(" Where " & Gen_WcLocation & " Order by SecurityRequestID ")
Else
    Call TableMst_SecurityRequest(" Order by SecurityRequestID ")
End If

If RsMst_SecurityRequest.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
RsMst_SecurityRequest.MoveLast
Call Indata
Call GButtonLock(Me, True)
End Sub


Private Sub LastCmd_Click()
RsMst_SecurityRequest.MoveLast
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
RsMst_SecurityRequest.MoveFirst
RsMst_SecurityRequest.Find "SecurityRequestID = '" & MSHFlexGrid1.TextMatrix(I, 0) & "'"
If Not RsMst_SecurityRequest.EOF Then
   Call Indata
   Frame0.Visible = False
End If
End Sub

Private Sub MSHFlexGrid2_Click()
If SaveCmd.Enabled = False Then Exit Sub
Call ClearGodownMapping
I = MSHFlexGrid2.RowSel
If MSHFlexGrid2.TextMatrix(I, 0) = "" Then Exit Sub
CmbCMP.Text = MSHFlexGrid2.TextMatrix(I, 0)
CmbGodown.Text = MSHFlexGrid2.TextMatrix(I, 1)
TxtGodownAddress.Text = MSHFlexGrid2.TextMatrix(I, 2)
TxtBalanceBags.Text = MSHFlexGrid2.TextMatrix(I, 3)
TxtBalanceQty.Text = MSHFlexGrid2.TextMatrix(I, 4)
TxtAUM.Text = MSHFlexGrid2.TextMatrix(I, 5)
TxtRemarksD = MSHFlexGrid2.TextMatrix(I, 6)
mCMPID = MSHFlexGrid2.TextMatrix(I, 7)
mGodownID = MSHFlexGrid2.TextMatrix(I, 8)
CmdAddtoList.Caption = "Update List"
CmbCMP.Enabled = False
CmbGodown.Enabled = False
CmdDeletefromList.Enabled = True
End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_SecurityRequest.MoveNext
If RsMst_SecurityRequest.EOF Then
   RsMst_SecurityRequest.MoveLast
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
RsMst_SecurityRequest.MovePrevious
If RsMst_SecurityRequest.BOF Then
   RsMst_SecurityRequest.MoveFirst
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
If CmbLocation.Text = "" Then
    MsgBox "Blank Location not Allowed!!!"
    CmbLocation.SetFocus
    Exit Sub
End If
If CmbRequestTo.Text = "" Then
    MsgBox "Blank Request to not Allowed!!!"
    CmbRequestTo.SetFocus
    Exit Sub
End If
If CmbSecurityAgency.Text = "" Then
    MsgBox "Blank Security Agency not Allowed!!!"
    CmbSecurityAgency.SetFocus
    Exit Sub
End If
If TxtAgreeStartDate.Value > TxtAgreeEndDate.Value Then
   MsgBox "Agreement 'Start Date' must be less than Agreement 'End Date'!!!"
   TxtAgreeStartDate.SetFocus
   Exit Sub
End If
'If TxtSalaryPaidUpto.Value < TxtAgreeEndDate.Value Then
'   MsgBox "'Salary Payable Date' must be greater than Agreement 'End Date'!!!"
'   TxtSalaryPaidUpto.SetFocus
'   Exit Sub
'End If
If CmbSecurityType.Text = "" Then
    MsgBox "Blank Security Type not Allowed!!!"
    CmbSecurityType.SetFocus
    Exit Sub
End If
If CmbSecurityDesignation.Text = "" Then
    MsgBox "Blank Security Designation not Allowed!!!"
    CmbSecurityDesignation.SetFocus
    Exit Sub
End If
If TxtRatePerMonth.Text = "" Then
    MsgBox "Blank Rate Per Month not Allowed!!!"
    TxtRatePerMonth.SetFocus
    Exit Sub
End If

If CmbStatus.Text = "" Then
    MsgBox "Blank Status not Allowed!!!"
    CmbStatus.SetFocus
    Exit Sub
End If
If LCase(CmbStatus.Text) <> "pending" And TxtApproveRemarks.Text = "" Then
    MsgBox "Blank Approval Remarks not Allowed!!!"
    TxtApproveRemarks.SetFocus
    Exit Sub
End If
If MSHFlexGrid2.Rows = 3 Then
    MsgBox "Map Atleast one Godown!!!"
    SSTab1.Tab = 0
    CmbCMP.SetFocus
    Exit Sub
End If


If LCase(CmbStatus.Text) <> "rejected" Then
   If CheckSecurityAgencyAndLocation(mLocationID, mSecurityAgencyID, TxtAgreeStartDate.Value, TxtAgreeEndDate.Value) = False Then
      MsgBox "Monthly Attendance process already executed for seleceted Start Date and End Date!!!"
      Exit Sub
   End If

   If CheckMonthwiseTxnID = False Then
      MsgBox "Security Agency '" & CmbSecurityAgency.Text & "' for '" & CmbLocation.Text & "' is already processed!!!"
      Exit Sub
   End If
End If


If Edit_Flg = "A" Then
    RsMst_SecurityRequest.AddNew
    RsMst_SecurityRequest!SecurityRequestID = GetMaxRequestID
    RsMst_SecurityRequest!G_UID = GetGUID
    TxtRequestID = RsMst_SecurityRequest!SecurityRequestID
    TxtG_UID = RsMst_SecurityRequest!G_UID
Else
    RsMst_SecurityRequest.MoveFirst
    RsMst_SecurityRequest.Find "SecurityRequestID= " & TxtRequestID & ""
End If
RsMst_SecurityRequest!RequestDate = TxtRequestDate.Value
RsMst_SecurityRequest!LocationID = mLocationID
RsMst_SecurityRequest!RequestToID = mRequestToID
RsMst_SecurityRequest!SecurityAgencyID = mSecurityAgencyID
RsMst_SecurityRequest!NoofOpenGodowns = Val(TxtNoofopenGodown)
RsMst_SecurityRequest!NoofGuards = Val(TxtNoofGuards)
RsMst_SecurityRequest!NoofGodownsMapped = Val(TxtNoofGodownsMapped)
RsMst_SecurityRequest!TotalCost = Val(TxtTotalCost)
RsMst_SecurityRequest!TotalAUM = Val(TxtTotalAUM)
RsMst_SecurityRequest!AgreeStartDate = TxtAgreeStartDate.Value
RsMst_SecurityRequest!AgreeEndDate = TxtAgreeEndDate.Value
RsMst_SecurityRequest!SalaryPayableUpto = TxtSalaryPaidUpto.Value
RsMst_SecurityRequest!SecurityTypeID = mSecurityTypeID
RsMst_SecurityRequest!SecurityDesignationID = mSecurityDesignationID
RsMst_SecurityRequest!RatePerMonth = Val(TxtRatePerMonth)
RsMst_SecurityRequest!Status = Left$(CmbStatus.Text, 1)
RsMst_SecurityRequest!Remarks = TxtRemarks.Text
If LCase(CmbStatus.Text) <> "pending" Then
    RsMst_SecurityRequest!ApprovedRemarks = TxtApproveRemarks.Text
    RsMst_SecurityRequest!ApprovedBy = Gen_UserLoginID
    RsMst_SecurityRequest!ApprovedDate = Now
End If
If IsNull(RsMst_SecurityRequest!CreatedBy) = True Or RsMst_SecurityRequest!CreatedBy = "" Then
   RsMst_SecurityRequest!CreatedBy = Gen_UserLoginID
   RsMst_SecurityRequest!CreatedDate = Now
Else
    If LCase(CmbStatus.Text) = "pending" Then
        RsMst_SecurityRequest!UpdatedBy = Gen_UserLoginID
        RsMst_SecurityRequest!UpdatedDate = Now
    End If
End If

RsMst_SecurityRequest.Update
RsMst_SecurityRequest.Requery

Call SaveSecurityRequestGodownMappingDetails

If LCase(CmbStatus.Text) = "approved" Then
    tmp = "Select * From Mst_SecLocMapping Where LocationID = " & mLocationID & " And SecurityAgencyID = " & mSecurityAgencyID
    Call Tmp1Table
    If TmpRs1.RecordCount > 0 Then
        mTxnID = TmpRs1!TxnID
    Else
        TmpRs1.AddNew
'        TmpRs1!TxnID = GetMaxTxnID
        TmpRs1!G_UID = GetGUID
        mTxnID = TmpRs1!TxnID
        TmpRs1!TxnDate = Date ' Now
        TmpRs1!SecurityAgencyID = mSecurityAgencyID
        TmpRs1!LocationID = mLocationID
        TmpRs1!CreatedBy = Gen_UserLoginID
        TmpRs1!CreatedDate = Now
    End If
    TmpRs1!AgreementStartDate = TxtAgreeStartDate.Value
    TmpRs1!AgreementEndDate = TxtAgreeEndDate.Value
    TmpRs1!UpdatedBy = Gen_UserLoginID
    TmpRs1!UpdatedDate = Now
    TmpRs1.Update
        
    
    tmp = "Select * From Mst_SecLocMappingDetail Where TxnID = " & mTxnID
    Call Tmp2Table
    
    TmpRs2.AddNew
    TmpRs2!TxnDetailID = GetMaxTxnDetailsID
    mDetailID = TmpRs2!TxnDetailID
    TmpRs2!TxnID = mTxnID 'TmpRs1!TxnID
    TmpRs2!SecurityTypeID = mSecurityTypeID
    TmpRs2!SecurityDesignationID = mSecurityDesignationID
    TmpRs2!StartDate = TxtAgreeStartDate.Value
    TmpRs2!EndDate = TxtAgreeEndDate.Value
    TmpRs2!NoofGodown = MSHFlexGrid2.Rows - 3
    TmpRs2!RatePerMonth = TxtRatePerMonth
    TmpRs2!Remarks = TxtRemarks.Text
    TmpRs2!SalaryPaidUpto = TxtSalaryPaidUpto.Value
    TmpRs2!CreatedBy = Gen_UserLoginID
    TmpRs2!CreatedDate = Now
    TmpRs2!RequestID = TxtRequestID
    TmpRs2.Update
    
    tmp = "Select * From Mst_SecLocGodMappingDetail Where TxnID = " & mTxnID 'TmpRs1!TxnID
    Call Tmp2Table
    
    For I = 2 To MSHFlexGrid2.Rows - 2
        TmpRs2.AddNew
        TmpRs2!GodMapDID = GetMaxGodMapDID
        TmpRs2!TxnID = mTxnID 'TmpRs1!TxnID
        'TmpRs2!SecurityTypeID = mSecurityTypeID
        TmpRs2!TxnDetailID = mDetailID
        TmpRs2!CMPID = Val(MSHFlexGrid2.TextMatrix(I, 7))
        TmpRs2!GodownID = Val(MSHFlexGrid2.TextMatrix(I, 8))
        TmpRs2!BalanceBags = Val(MSHFlexGrid2.TextMatrix(I, 3))
        TmpRs2!BalanceQty = Val(MSHFlexGrid2.TextMatrix(I, 4))
        TmpRs2!AUM = Val(MSHFlexGrid2.TextMatrix(I, 5))
        TmpRs2!Remarks = MSHFlexGrid2.TextMatrix(I, 6)
        TmpRs2!CreatedBy = Gen_UserLoginID
        TmpRs2!CreatedDate = Now
        TmpRs2.Update
    Next
End If

RsMst_SecurityRequest.MoveFirst
RsMst_SecurityRequest.Find "SecurityRequestID = " & TxtRequestID
Call Indata
End Sub

Private Function GetMaxTxnID() As Double
Dim Retval As Double
tmp = "Select max(TxnID) from Mst_SecLocMapping "
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetMaxTxnID = Retval
End Function

Function GetMaxTxnDetailsID() As Double
Dim Retval As Double
tmp = "Select max(TxnDetailID) From Mst_SecLocMappingDetail " 'Where ClaimID = " & TxtClaimID & " And SM_prefix = '" & TxtPreFix & "'"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetMaxTxnDetailsID = Retval
End Function

Function GetMaxGodMapDID() As Double
Dim Retval As Double
tmp = "Select max(GodMapDID) From Mst_SecLocGodMappingDetail " 'Where ClaimID = " & TxtClaimID & " And SM_prefix = '" & TxtPreFix & "'"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetMaxGodMapDID = Retval
End Function

Private Sub Indata()

Edit_Flg = ""
Frame1.Enabled = False
Frame3.Enabled = False
Call GButtonLock(Me, True)
Call ComboVisible(False)
TxtApproveRemarks.Locked = True

tmp = " select"
tmp = tmp & " MSR.SecurityRequestID, MSR.RequestDate, MSR.LocationID, MSR.RequestToID, MSR.SecurityAgencyID, MSR.AgreeStartDate,"
tmp = tmp & " MSR.AgreeEndDate, MSR.SalaryPayableUpto, MSR.SecurityTypeID, MSR.SecurityDesignationID, MSR.RatePerMonth,"
tmp = tmp & " MSR.Status, MSR.Remarks, MSR.ApprovedRemarks, MSR.ApprovedDate, MSR.ApprovedBy, MSR.G_UID, MSR.CreatedDate, MSR.CreatedBy, MSR.UpdatedDate, MSR.UpdatedBy,"
tmp = tmp & " ML.LocationName, ME.EmployeeName+'~'+ME.EmployeeNo as RequestTo, MSA.AgencyName, MSS.SecurityShifts, MSD.Securitydesignation,"
tmp = tmp & " MSR.NoofOpenGodowns, MSR.NoofGuards, MSR.NoofGodownsMapped, MSR.TotalCost, MSR.TotalAUM,ML.StateID "
tmp = tmp & " From Mst_SecurityRequest MSR"
tmp = tmp & " Inner Join Mst_Location ML On ML.LocatioNID = MSR.LocationID"
tmp = tmp & " Inner Join Mst_Employee ME On ME.EmployeeID = MSR.RequestToID"
tmp = tmp & " Inner Join Mst_SecurityAgency MSA ON MSA.SecurityAgencyID = MSR.SecurityAgencyID"
tmp = tmp & " Inner Join Mst_SecurityShifts MSS ON MSS.SecurityShiftsID = MSR.SecurityTypeID"
tmp = tmp & " Inner Join Mst_SecurityDesignation MSD ON MSD.SecurityDesignationID = MSR.SecurityDesignationID"
tmp = tmp & " Where SecurityRequestID = " & Val(RsMst_SecurityRequest!SecurityRequestID)
Call TmpTable

If TmpRs.RecordCount > 0 Then
    TxtRequestID = RsMst_SecurityRequest!SecurityRequestID
    TxtRequestDate.Value = IIf(IsNull(TmpRs!RequestDate), Date, TmpRs!RequestDate)
    TxtAgreeStartDate.Value = IIf(IsNull(TmpRs!AgreeStartDate), Date, TmpRs!AgreeStartDate)
    TxtAgreeEndDate.Value = IIf(IsNull(TmpRs!AgreeEndDate), Date, TmpRs!AgreeEndDate)
    TxtSalaryPaidUpto.Value = IIf(IsNull(TmpRs!SalaryPayableUpto), Date, TmpRs!SalaryPayableUpto)
    TxtRatePerMonth.Text = IIf(IsNull(TmpRs!RatePerMonth), "", TmpRs!RatePerMonth)
    TxtRemarks.Text = IIf(IsNull(TmpRs!Remarks), "", TmpRs!Remarks)
    TxtApproveRemarks.Text = IIf(IsNull(TmpRs!ApprovedRemarks), "", TmpRs!ApprovedRemarks)
    TxtG_UID.Text = IIf(IsNull(TmpRs!G_UID), "", TmpRs!G_UID)
    TxtNoofopenGodown = IIf(IsNull(TmpRs!NoofOpenGodowns), "", TmpRs!NoofOpenGodowns)
    TxtNoofGuards = IIf(IsNull(TmpRs!NoofGuards), "", TmpRs!NoofGuards)
    TxtNoofGodownsMapped = IIf(IsNull(TmpRs!NoofGodownsMapped), "", TmpRs!NoofGodownsMapped)
    TxtTotalCost = IIf(IsNull(TmpRs!TotalCost), "", TmpRs!TotalCost)
    TxtTotalAUM = IIf(IsNull(TmpRs!TotalAUM), "", TmpRs!TotalAUM)
    If LCase(TmpRs!Status) = "r" Then
        CmbStatus.Text = "Rejected"
    ElseIf LCase(TmpRs!Status) = "a" Then
        CmbStatus.Text = "Approved"
    Else
        CmbStatus.Text = "Pending"
    End If
    TxtCreated = IIf(IsNull(TmpRs!CreatedBy), "", TmpRs!CreatedBy) & " :- " & IIf(IsNull(TmpRs!CreatedDate), "", TmpRs!CreatedDate)
    TxtUpdated = IIf(IsNull(TmpRs!UpdatedBy), "", TmpRs!UpdatedBy) & " :- " & IIf(IsNull(TmpRs!UpdatedDate), "", TmpRs!UpdatedDate)
    TxtApprovedBy = IIf(IsNull(TmpRs!ApprovedBy), "", TmpRs!ApprovedBy) & " :- " & IIf(IsNull(TmpRs!ApprovedDate), "", TmpRs!ApprovedDate)
    CmbSecurityAgency.Text = IIf(IsNull(TmpRs!AgencyName), "", TmpRs!AgencyName)
    CmbLocation.Text = IIf(IsNull(TmpRs!Locationname), "", TmpRs!Locationname)
    CmbRequestTo.Text = IIf(IsNull(TmpRs!RequestTo), "", TmpRs!RequestTo)
    CmbSecurityDesignation.Text = IIf(IsNull(TmpRs!SecurityDesignation), "", TmpRs!SecurityDesignation)
    CmbSecurityType.Text = IIf(IsNull(TmpRs!SecurityShifts), "", TmpRs!SecurityShifts)
    
    TxtSecurityAgency.Text = CmbSecurityAgency.Text
    TxtLocation.Text = CmbLocation.Text
    TxtRequestTo.Text = CmbRequestTo.Text
    TxtSecurityDesignation.Text = CmbSecurityDesignation.Text
    TxtSecurityType = CmbSecurityType.Text
    TxtStatus.Text = CmbStatus.Text
    
    mSecurityAgencyID = IIf(IsNull(TmpRs!SecurityAgencyID), 0, TmpRs!SecurityAgencyID)
    mLocationID = IIf(IsNull(TmpRs!LocationID), 0, TmpRs!LocationID)
    
    mSecurityTypeID = IIf(IsNull(TmpRs!SecurityTypeID), 0, TmpRs!SecurityTypeID)
    mSecurityDesignationID = IIf(IsNull(TmpRs!SecurityDesignationID), 0, TmpRs!SecurityDesignationID)
    mRequestToID = IIf(IsNull(TmpRs!RequestToID), 0, TmpRs!RequestToID)
    mStateID = IIf(IsNull(TmpRs!StateID), 0, TmpRs!StateID)
    
    Call GetAgreementStartDateIndata
    
    'Call GetNoofOpenGodown
    'Call GetGodownCostDetails
 
End If

Call InDataSecurityRequestGodownMappingDetails
Call GetGodownInfoDetails

End Sub

Private Sub InDataSecurityRequestGodownMappingDetails()
Call ClearGodownMapping
Call Grid_HeadGodownMapping

tmp = " Select MSRGM.SecurityRequestGodownMappingID, MSRGM.SecurityRequestID, MSRGM.CMPID, MSRGM.GodownID, MSRGM.Remarks,"
tmp = tmp & " MCMP.CMPName , MG.GodownNo, MG.Address, MSRGM.BalanceBags, MSRGM.BalanceQty, MSRGM.AUM"
tmp = tmp & " From Mst_SecurityRequestGodownMapping MSRGM"
tmp = tmp & " Inner Join Mst_CMP MCMP on MCMP.CMPID = MSRGM.CMPID"
tmp = tmp & " Inner Join Mst_Godown MG on MG.GodownID = MSRGM.GodownID"
tmp = tmp & " Where MSRGM.SecurityRequestID = " & Val(TxtRequestID)
Call TmpTable

If TmpRs.RecordCount > 0 Then
    MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + TmpRs.RecordCount
    For I = 2 To TmpRs.RecordCount + 1
        With MSHFlexGrid2
            .TextMatrix(I, 0) = IIf(IsNull(TmpRs!CMPName), "", TmpRs!CMPName)
            .TextMatrix(I, 1) = IIf(IsNull(TmpRs!GodownNo), "", TmpRs!GodownNo)
            .TextMatrix(I, 2) = IIf(IsNull(TmpRs!Address), "", TmpRs!Address)
            .TextMatrix(I, 3) = IIf(IsNull(TmpRs!BalanceBags), 0, TmpRs!BalanceBags)
            .TextMatrix(I, 4) = IIf(IsNull(TmpRs!BalanceQty), 0, TmpRs!BalanceQty)
            .TextMatrix(I, 5) = IIf(IsNull(TmpRs!AUM), 0, TmpRs!AUM)
            .TextMatrix(I, 6) = IIf(IsNull(TmpRs!Remarks), "", TmpRs!Remarks)
            .TextMatrix(I, 7) = IIf(IsNull(TmpRs!CMPID), 0, TmpRs!CMPID)
            .TextMatrix(I, 8) = IIf(IsNull(TmpRs!GodownID), 0, TmpRs!GodownID)
            .TextMatrix(I, 9) = IIf(IsNull(TmpRs!SecurityRequestGodownMappingID), 0, TmpRs!SecurityRequestGodownMappingID)
            .TextMatrix(0, 3) = Val(.TextMatrix(I, 3)) + Val(.TextMatrix(0, 3))
            .TextMatrix(0, 4) = Val(.TextMatrix(I, 4)) + Val(.TextMatrix(0, 4))
            .TextMatrix(0, 5) = Val(.TextMatrix(I, 5)) + Val(.TextMatrix(0, 5))
        End With
        TmpRs.MoveNext
    Next
End If
'TxtTotalAUM.Text = MSHFlexGrid2.TextMatrix(0, 5)

End Sub

Private Sub SaveSecurityRequestGodownMappingDetails()

tmp1 = "(0"
For j = 2 To MSHFlexGrid2.Rows - 2
    If MSHFlexGrid2.TextMatrix(j, 9) <> "" Then
        tmp1 = tmp1 & "," & MSHFlexGrid2.TextMatrix(j, 9)
    End If
Next
tmp1 = tmp1 & ")"
tmp = "Delete From Mst_SecurityRequestGodownMapping Where SecurityRequestID = " & TxtRequestID & " And SecurityRequestGodownMappingID not in " & tmp1
Call TmpTable


Call TableMst_SecurityRequestGodownMapping(" Where SecurityRequestID = " & TxtRequestID)
For j = 2 To MSHFlexGrid2.Rows - 2
    If MSHFlexGrid2.TextMatrix(j, 9) = "" Then
        RsMst_SecurityRequestGodownMapping.AddNew
        RsMst_SecurityRequestGodownMapping!SecurityRequestGodownMappingID = GetMaxSecurityRequestGodownMappingID
        RsMst_SecurityRequestGodownMapping!G_UID = GetGUID
    Else
        RsMst_SecurityRequestGodownMapping.MoveFirst
        RsMst_SecurityRequestGodownMapping.Find "SecurityRequestGodownMappingID = " & MSHFlexGrid2.TextMatrix(j, 9)
    End If
    RsMst_SecurityRequestGodownMapping!SecurityRequestID = TxtRequestID
    RsMst_SecurityRequestGodownMapping!BalanceBags = Val(MSHFlexGrid2.TextMatrix(j, 3))
    RsMst_SecurityRequestGodownMapping!BalanceQty = Val(MSHFlexGrid2.TextMatrix(j, 4))
    RsMst_SecurityRequestGodownMapping!AUM = Val(MSHFlexGrid2.TextMatrix(j, 5))
    RsMst_SecurityRequestGodownMapping!CMPID = Val(MSHFlexGrid2.TextMatrix(j, 7))
    RsMst_SecurityRequestGodownMapping!GodownID = Val(MSHFlexGrid2.TextMatrix(j, 8))
    RsMst_SecurityRequestGodownMapping!Remarks = MSHFlexGrid2.TextMatrix(j, 6)
    If IsNull(RsMst_SecurityRequestGodownMapping!CreatedBy) = True Or RsMst_SecurityRequestGodownMapping!CreatedBy = "" Then
       RsMst_SecurityRequestGodownMapping!CreatedBy = Gen_UserLoginID
       RsMst_SecurityRequestGodownMapping!CreatedDate = Now
    Else
       RsMst_SecurityRequestGodownMapping!UpdatedBy = Gen_UserLoginID
       RsMst_SecurityRequestGodownMapping!UpdatedDate = Now
    End If
    RsMst_SecurityRequestGodownMapping.Update
Next
End Sub

Private Function GetMaxRequestID() As Double
Dim Retval As Double
tmp = "Select max(SecurityRequestID) from Mst_SecurityRequest "
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetMaxRequestID = Retval
End Function

Private Function GetMaxSecurityRequestGodownMappingID() As Double
Dim Retval As Double
tmp = "Select max(SecurityRequestGodownMappingID) from Mst_SecurityRequestGodownMapping"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetMaxSecurityRequestGodownMappingID = Retval
End Function

Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   If RsMst_SecurityRequest.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   RsMst_SecurityRequest.MoveFirst
   RsMst_SecurityRequest.Find "SecurityRequestID= " & Val(TxtRequestID) & ""
   If RsMst_SecurityRequest.EOF Then
      RsMst_SecurityRequest.MoveLast
   End If
   Call Indata
   Exit Sub
End If

Frame0.Height = Frame4.Height - 200 'Me.Height - 500
Frame0.Left = Frame1.Left - 100
Frame0.Top = Frame1.Top
Frame0.Width = Frame4.Width - 200 'Me.Width - 200
Frame0.Visible = True

MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 1800

RsMst_SecurityRequest.MoveFirst

''===========comment on 7 july 2012 for not clearing search grid=============
'Call Grid_Head1
'
'SFrom.Value = CDate(From_GenDate) - 900
'STo.Value = Date 'To_GenDate
'CmbSSecurityAgency.Text = ""
'CmbSLocation.Text = ""
'Label22.Caption = ""
''===========comment on 7 july 2012 for not clearing search grid=============

End Sub

Private Sub TxtAgreeEndDate_LostFocus()

If IsNull(mAgreeEndDt) = False Then
   If TxtAgreeEndDate.Value > CDate(mAgreeEndDt) Then
      MsgBox "Sec. Appointed Date should not greater than Agreement Expiry Date(" & mAgreeEndDt & " ) "
      TxtAgreeEndDate.SetFocus
      Exit Sub
   End If
End If

TxtSalaryPaidUpto.Value = TxtAgreeEndDate.Value

End Sub

Private Sub TxtAgreeStartDate_LostFocus()

If Format(Gen_BlockingDate, "yyyy-mm") >= Format(TxtAgreeStartDate.Value, "yyyy-mm") Then
   MsgBox "Agreement Start Date allowed for next month of " & Format(Gen_BlockingDate, "MMM-yyyy") & "!!! "
   TxtAgreeStartDate.SetFocus
   Exit Sub
End If

If IsNull(mAgreeStartDt) = False Then
   If TxtAgreeStartDate.Value < CDate(mAgreeStartDt) Then
      MsgBox "Sec. Appointed Date should not less than Agreement Start Date"
      TxtAgreeStartDate.SetFocus
      Exit Sub
   End If
End If


End Sub

Private Sub TxtRatePerMonth_LostFocus()
If TxtRatePerMonth.Text = "" Then Exit Sub
If IsNumeric(TxtRatePerMonth) = False Then
   MsgBox "Invalid Numeric Format !!!!"
   TxtRatePerMonth.SetFocus
   Exit Sub
End If
TxtRatePerMonth = Val(TxtRatePerMonth)
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

Private Function CheckSecurityAgencyAndLocation(ByVal mLocationID_ As Variant, ByVal mAgencyID_ As Variant, ByVal mStartDate_ As Variant, ByVal mEndDate_ As Variant) As Boolean

CheckSecurityAgencyAndLocation = True


mStartDate_ = CDate(Year(CDate(mStartDate_)) & "-" & Month(CDate(mStartDate_)) & "-01")

mSMonth = Month(CDate(mEndDate_))
mSYear = Year(CDate(mEndDate_))
mSTotalDays = Day(DateSerial(mSYear, mSMonth + 1, 0))
mEndDate_ = CDate(mSYear & "-" & mSMonth & "-" & mSTotalDays)

'If Month(mStartDate_) < Month(mEndDate_) Then
'   tmp = " select MonthWiseTxnID from Txn_MonthwiseSecLoc"
'   tmp = tmp & " Where SecurityAgencyID = " & mAgencyID_ & " And LocationID = " & mLocationID_
'   tmp = tmp & " And month(MonthYear) between " & Month(mStartDate_) & " and " & Month(mEndDate_)
 '  tmp = tmp & " And Year(MonthYear) between " & Year(mStartDate_) & " and " & Year(mEndDate_)
'Else
'   tmp = " select MonthWiseTxnID from Txn_MonthwiseSecLoc"
'   tmp = tmp & " Where SecurityAgencyID = " & mAgencyID_ & " And LocationID = " & mLocationID_
'   tmp = tmp & " And month(MonthYear) between " & Month(mEndDate_) & " and " & Month(mStartDate_)
'   tmp = tmp & " And Year(MonthYear) between " & Year(mStartDate_) & " and " & Year(mEndDate_)
'End If


tmp = " select MonthWiseTxnID from Txn_MonthwiseSecLoc"
tmp = tmp & " Where SecurityAgencyID = " & mAgencyID_ & " And LocationID = " & mLocationID_
tmp = tmp & " And MonthYear between '" & Format(mStartDate_, Gen_DatFormat) & "' and '" & Format(mEndDate_, Gen_DatFormat) & "'"

Call TmpTable

If TmpRs.RecordCount > 0 Then
   CheckSecurityAgencyAndLocation = False
End If

End Function

Private Function CheckMonthwiseTxnID() As Boolean
CheckMonthwiseTxnID = True
tmp = "select * from Txn_MonthwiseSecLoc where monthyear between '" & Format(TxtAgreeStartDate.Value, "yyyy-mm-dd") & "' and '" & Format(TxtAgreeEndDate.Value, "yyyy-mm-dd") & "' and securityagencyid = " & mSecurityAgencyID & "  and locationid = " & mLocationID
Call Tmp7Table
If TmpRs7.RecordCount > 0 Then
   CheckMonthwiseTxnID = False
End If
End Function

Private Sub GetAgreementStartDate()

tmp = "Select TSA.AgreeStartDate,TSA.AgreeExpiryDate,TSA.SAID"
tmp = tmp & " From Txn_SecurityAgreement TSA"
tmp = tmp & " Inner join Mst_State MS on TSA.StateID=MS.StateID"
tmp = tmp & " Where TSA.Flag='A' And MS.StateID =" & mStateID & " And TSA.SecurityAgencyID = " & mSecurityAgencyID
'tmp = tmp & " And TSA.AgreeStartDate<='" & Format(TxtAgreeStartDate.Value, Gen_DatFormat) & "' And TSA.AgreeExpiryDate>='" & Format(TxtAgreeStartDate.Value, Gen_DatFormat) & "'"
tmp = tmp & " And TSA.AgreeStartDate<='" & Format(TxtAgreeEndDate.Value, Gen_DatFormat) & "' And TSA.AgreeExpiryDate>='" & Format(TxtAgreeStartDate.Value, Gen_DatFormat) & "'"
tmp = tmp & " order by TSA.AgreeStartDate Desc"

'tmp = " Select AgreementStartDate, AgreementEndDate From Mst_SecLocMapping Where LocationID = " & mLocationID
'tmp = tmp & " And SecurityAgencyID = " & mSecurityAgencyID

Call TmpTable

If TmpRs.RecordCount > 0 Then
    
    TxtAgreeStartDate.Value = IIf(IsNull(TmpRs!AgreeStartDate), Date - 30, TmpRs!AgreeStartDate)
    TxtAgreeEndDate.Value = IIf(IsNull(TmpRs!AgreeExpiryDate), Date + 30, TmpRs!AgreeExpiryDate)
    
    mAgreeStartDt = IIf(IsNull(TmpRs!AgreeStartDate), Null, TmpRs!AgreeStartDate)
    mAgreeEndDt = IIf(IsNull(TmpRs!AgreeExpiryDate), Null, TmpRs!AgreeExpiryDate)
    
    mSAID = TmpRs!SAID
    
    TxtSalaryPaidUpto.Value = TxtAgreeEndDate.Value
Else
    MsgBox "No live agreement available for selected security agency!!"
    CmbSecurityAgency.SetFocus
    Exit Sub
End If

End Sub


Private Sub GetAgreementStartDateIndata()

tmp = "Select TSA.AgreeStartDate,TSA.AgreeExpiryDate,TSA.SAID"
tmp = tmp & " From Txn_SecurityAgreement TSA"
tmp = tmp & " Inner join Mst_State MS on TSA.StateID=MS.StateID"
tmp = tmp & " Where TSA.Flag='A' And MS.StateID =" & mStateID & " And TSA.SecurityAgencyID = " & mSecurityAgencyID
'tmp = tmp & " And TSA.AgreeStartDate<='" & Format(TxtAgreeStartDate.Value, Gen_DatFormat) & "' And TSA.AgreeExpiryDate>='" & Format(TxtAgreeStartDate.Value, Gen_DatFormat) & "'"
tmp = tmp & " And TSA.AgreeStartDate<='" & Format(TxtAgreeEndDate.Value, Gen_DatFormat) & "' And TSA.AgreeExpiryDate>='" & Format(TxtAgreeStartDate.Value, Gen_DatFormat) & "'"
tmp = tmp & " order by TSA.AgreeStartDate Desc"

'tmp = " Select AgreementStartDate, AgreementEndDate From Mst_SecLocMapping Where LocationID = " & mLocationID
'tmp = tmp & " And SecurityAgencyID = " & mSecurityAgencyID

Call TmpTable

If TmpRs.RecordCount > 0 Then
    mAgreeStartDt = IIf(IsNull(TmpRs!AgreeStartDate), Null, TmpRs!AgreeStartDate)
    mAgreeEndDt = IIf(IsNull(TmpRs!AgreeExpiryDate), Null, TmpRs!AgreeExpiryDate)
    mSAID = TmpRs!SAID
End If

End Sub

