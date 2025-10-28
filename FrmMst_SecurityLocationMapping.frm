VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmMst_SecurityLocationMapping 
   Caption         =   "Security Location Mapping"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.VScrollBar VScroll1 
      Height          =   9015
      Left            =   15000
      TabIndex        =   42
      Top             =   0
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      TabIndex        =   41
      Top             =   10320
      Width           =   10935
   End
   Begin VB.Frame Frame4 
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   9780
      Left            =   0
      TabIndex        =   18
      Top             =   -120
      Width           =   15180
      Begin VB.Frame Frame0 
         Height          =   1335
         Left            =   120
         TabIndex        =   32
         Top             =   5040
         Visible         =   0   'False
         Width           =   14655
         Begin VB.TextBox TxtSTxnID 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   10800
            TabIndex        =   15
            Top             =   360
            Width           =   1305
         End
         Begin VB.ComboBox CmbSSecurityAgency 
            Height          =   315
            Left            =   2795
            Sorted          =   -1  'True
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   353
            Width           =   4065
         End
         Begin VB.CommandButton CmdGo 
            Caption         =   "Go"
            Height          =   350
            Left            =   12375
            TabIndex        =   16
            Top             =   705
            Width           =   1000
         End
         Begin VB.CommandButton CmdExcel 
            Caption         =   "Excel"
            Height          =   350
            Left            =   13395
            TabIndex        =   17
            Top             =   705
            Width           =   1000
         End
         Begin VB.ComboBox CmbSLocation 
            Height          =   315
            Left            =   6870
            Sorted          =   -1  'True
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   353
            Width           =   3780
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   1455
            Left            =   75
            TabIndex        =   33
            Top             =   1110
            Width           =   14355
            _ExtentX        =   25321
            _ExtentY        =   2566
            _Version        =   393216
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSComCtl2.DTPicker STo 
            Height          =   300
            Left            =   1450
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   360
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
            Format          =   109969409
            CurrentDate     =   36494
         End
         Begin MSComCtl2.DTPicker SFrom 
            Height          =   300
            Left            =   105
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   360
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
            Format          =   109969409
            CurrentDate     =   36494
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "TxnID"
            Height          =   195
            Left            =   11235
            TabIndex        =   95
            Top             =   135
            Width           =   435
         End
         Begin VB.Label Label44 
            Caption         =   "From Date"
            Height          =   255
            Left            =   382
            TabIndex        =   39
            Top             =   120
            Width           =   780
         End
         Begin VB.Label Label37 
            Caption         =   "Security Agency"
            Height          =   225
            Left            =   4212
            TabIndex        =   38
            Top             =   135
            Width           =   1230
         End
         Begin VB.Label Label45 
            Caption         =   "To Date"
            Height          =   255
            Left            =   1810
            TabIndex        =   37
            Top             =   120
            Width           =   615
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
            Left            =   9465
            TabIndex        =   36
            Top             =   780
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
            Left            =   7095
            TabIndex        =   35
            Top             =   780
            Width           =   2415
         End
         Begin VB.Label Label31 
            Caption         =   "Location"
            Height          =   225
            Left            =   8400
            TabIndex        =   34
            Top             =   135
            Width           =   720
         End
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   6120
         Left            =   120
         TabIndex        =   47
         Top             =   2520
         Width           =   14835
         _ExtentX        =   26167
         _ExtentY        =   10795
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Security Guard Details"
         TabPicture(0)   =   "FrmMst_SecurityLocationMapping.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "MSHFlexGrid2"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Frame2"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Godown Mapping Details"
         TabPicture(1)   =   "FrmMst_SecurityLocationMapping.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "MSHFlexGrid3"
         Tab(1).Control(1)=   "Frame3"
         Tab(1).ControlCount=   2
         Begin VB.Frame Frame3 
            Height          =   2220
            Left            =   -74880
            TabIndex        =   71
            Top             =   360
            Width           =   14655
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
               Left            =   1020
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   81
               Top             =   1620
               Width           =   10215
            End
            Begin VB.ComboBox CmbSecurityType1 
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
               ItemData        =   "FrmMst_SecurityLocationMapping.frx":0038
               Left            =   120
               List            =   "FrmMst_SecurityLocationMapping.frx":003A
               Sorted          =   -1  'True
               TabIndex        =   80
               Top             =   390
               Width           =   7335
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
               Left            =   7485
               Sorted          =   -1  'True
               TabIndex        =   79
               Top             =   390
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
               Left            =   11640
               Sorted          =   -1  'True
               TabIndex        =   78
               Top             =   390
               Width           =   2970
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
               TabIndex        =   77
               Top             =   1680
               Width           =   1335
            End
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
               TabIndex        =   76
               Top             =   1680
               Width           =   1710
            End
            Begin VB.TextBox TxtBalanceBags 
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               Height          =   375
               Left            =   7470
               Locked          =   -1  'True
               TabIndex        =   75
               TabStop         =   0   'False
               Top             =   1065
               Width           =   2130
            End
            Begin VB.TextBox TxtBalanceQty 
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               Height          =   375
               Left            =   9630
               Locked          =   -1  'True
               TabIndex        =   74
               TabStop         =   0   'False
               Top             =   1065
               Width           =   1965
            End
            Begin VB.TextBox TxtAUM 
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               Height          =   375
               Left            =   11640
               Locked          =   -1  'True
               TabIndex        =   73
               TabStop         =   0   'False
               Top             =   1065
               Width           =   2955
            End
            Begin VB.TextBox TxtGodownAddress 
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               Height          =   510
               Left            =   135
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   72
               TabStop         =   0   'False
               Top             =   1065
               Width           =   7305
            End
            Begin VB.Label Label9 
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
               Left            =   3060
               TabIndex        =   89
               Top             =   120
               Width           =   1455
            End
            Begin VB.Label Label10 
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
               TabIndex        =   88
               Top             =   1755
               Width           =   945
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
               Left            =   9300
               TabIndex        =   87
               Top             =   120
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
               Left            =   12698
               TabIndex        =   86
               Top             =   120
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
               Left            =   7905
               TabIndex        =   85
               Top             =   780
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
               TabIndex        =   84
               Top             =   780
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
               Left            =   12885
               TabIndex        =   83
               Top             =   780
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
               Left            =   3000
               TabIndex        =   82
               Top             =   780
               Width           =   1560
            End
         End
         Begin VB.Frame Frame2 
            Height          =   1815
            Left            =   120
            TabIndex        =   48
            Top             =   360
            Width           =   14625
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
               Left            =   13170
               Locked          =   -1  'True
               TabIndex        =   57
               Top             =   645
               Width           =   1305
            End
            Begin VB.TextBox TxtNoofGodown 
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
               Left            =   11970
               Locked          =   -1  'True
               TabIndex        =   56
               Top             =   645
               Width           =   1155
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
               Left            =   12765
               TabIndex        =   55
               Top             =   1260
               Visible         =   0   'False
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
               Left            =   11415
               TabIndex        =   54
               Top             =   1260
               Visible         =   0   'False
               Width           =   1335
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
               ItemData        =   "FrmMst_SecurityLocationMapping.frx":003C
               Left            =   60
               List            =   "FrmMst_SecurityLocationMapping.frx":003E
               Sorted          =   -1  'True
               TabIndex        =   53
               Top             =   652
               Width           =   4095
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
               Left            =   1215
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   52
               Top             =   1095
               Width           =   10095
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
               ItemData        =   "FrmMst_SecurityLocationMapping.frx":0040
               Left            =   4635
               List            =   "FrmMst_SecurityLocationMapping.frx":0042
               Sorted          =   -1  'True
               TabIndex        =   51
               Top             =   652
               Width           =   3975
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
               Left            =   8640
               TabIndex        =   50
               TabStop         =   0   'False
               Top             =   675
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
               Left            =   4185
               TabIndex        =   49
               TabStop         =   0   'False
               Top             =   682
               Width           =   375
            End
            Begin MSComCtl2.DTPicker TxtStartDate 
               Height          =   375
               Left            =   9090
               TabIndex        =   58
               Top             =   645
               Width           =   1395
               _ExtentX        =   2461
               _ExtentY        =   661
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
               Format          =   109969409
               CurrentDate     =   39884
            End
            Begin MSComCtl2.DTPicker TxtEndDate 
               Height          =   375
               Left            =   10530
               TabIndex        =   59
               Top             =   645
               Width           =   1395
               _ExtentX        =   2461
               _ExtentY        =   661
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
               Format          =   109969409
               CurrentDate     =   39884
            End
            Begin MSComCtl2.DTPicker TxtSalaryPaidUpto 
               Height          =   375
               Left            =   585
               TabIndex        =   60
               Top             =   1320
               Visible         =   0   'False
               Width           =   1395
               _ExtentX        =   2461
               _ExtentY        =   661
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
               Format          =   109969409
               CurrentDate     =   39884
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
               Height          =   300
               Left            =   5512
               TabIndex        =   68
               Top             =   255
               Width           =   2220
            End
            Begin VB.Label Label5 
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
               TabIndex        =   67
               Top             =   1320
               Width           =   945
            End
            Begin VB.Label Lbl 
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
               Height          =   300
               Left            =   1357
               TabIndex        =   66
               Top             =   255
               Width           =   1500
            End
            Begin VB.Label Label8 
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
               Left            =   9120
               TabIndex        =   65
               Top             =   285
               Width           =   1050
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "End Date"
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
               Left            =   10605
               TabIndex        =   64
               Top             =   285
               Width           =   975
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               Caption         =   "Rate per month"
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
               Left            =   13357
               TabIndex        =   63
               Top             =   150
               Width           =   930
            End
            Begin VB.Label Label6 
               Alignment       =   2  'Center
               Caption         =   "No of Godown"
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
               Left            =   12045
               TabIndex        =   62
               Top             =   150
               Width           =   885
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
               Height          =   255
               Left            =   210
               TabIndex        =   61
               Top             =   1080
               Visible         =   0   'False
               Width           =   2265
            End
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
            Height          =   3585
            Left            =   120
            TabIndex        =   69
            Top             =   2325
            Width           =   14625
            _ExtentX        =   25797
            _ExtentY        =   6324
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
            Height          =   3345
            Left            =   -74880
            TabIndex        =   70
            Top             =   2625
            Width           =   14610
            _ExtentX        =   25770
            _ExtentY        =   5900
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
      End
      Begin VB.Frame Frame1 
         Height          =   2265
         Left            =   120
         TabIndex        =   19
         Top             =   105
         Width           =   14880
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
            Left            =   1245
            Sorted          =   -1  'True
            TabIndex        =   93
            Top             =   690
            Width           =   10455
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
            Left            =   1245
            Locked          =   -1  'True
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   1695
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
            Left            =   7110
            Locked          =   -1  'True
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   1680
            Width           =   4590
         End
         Begin VB.TextBox TxtTxnID 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   375
            Left            =   1245
            Locked          =   -1  'True
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   263
            Width           =   1455
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
            Left            =   11760
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   735
            Width           =   375
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
            Left            =   13365
            Locked          =   -1  'True
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   240
            Visible         =   0   'False
            Width           =   1440
         End
         Begin VB.CommandButton CmdLocationID 
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
            Left            =   11760
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   1140
            Width           =   375
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
            Left            =   1245
            Sorted          =   -1  'True
            TabIndex        =   0
            Top             =   1155
            Width           =   10455
         End
         Begin MSComCtl2.DTPicker TxtTxnDate 
            Height          =   375
            Left            =   3900
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   263
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   109969409
            CurrentDate     =   39884
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
            Left            =   1245
            Locked          =   -1  'True
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   690
            Width           =   10455
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
            Left            =   1245
            Locked          =   -1  'True
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   1155
            Width           =   10455
         End
         Begin MSComCtl2.DTPicker TxtAgreeStartDate 
            Height          =   375
            Left            =   6960
            TabIndex        =   90
            Top             =   270
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   661
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   109969409
            CurrentDate     =   39884
         End
         Begin MSComCtl2.DTPicker TxtAgreeEndDate 
            Height          =   375
            Left            =   10305
            TabIndex        =   94
            Top             =   270
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   661
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   109969409
            CurrentDate     =   39884
         End
         Begin VB.Label Label24 
            Caption         =   "Agreement End Date"
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
            Left            =   8640
            TabIndex        =   92
            Top             =   210
            Width           =   1395
         End
         Begin VB.Label Label21 
            Caption         =   "Agreement Start Date"
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
            Left            =   5280
            TabIndex        =   91
            Top             =   210
            Width           =   1560
         End
         Begin VB.Label Label17 
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
            Left            =   105
            TabIndex        =   46
            Top             =   1755
            Width           =   945
         End
         Begin VB.Label Label15 
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
            Left            =   6060
            TabIndex        =   45
            Top             =   1755
            Width           =   915
         End
         Begin VB.Label LbClaimID 
            Caption         =   "Txn ID"
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
            TabIndex        =   31
            Top             =   300
            Width           =   900
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
            Left            =   120
            TabIndex        =   30
            Top             =   630
            Width           =   1140
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
            Left            =   12420
            TabIndex        =   29
            Top             =   285
            Visible         =   0   'False
            Width           =   630
         End
         Begin VB.Label Label18 
            Caption         =   "Txn Date"
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
            Left            =   2985
            TabIndex        =   28
            Top             =   315
            Width           =   900
         End
         Begin VB.Label Label28 
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
            Height          =   270
            Left            =   105
            TabIndex        =   27
            Top             =   1200
            Width           =   1230
         End
      End
      Begin VB.Frame ButtonFrm 
         Height          =   1080
         Left            =   120
         TabIndex        =   40
         Top             =   8625
         Width           =   14865
         Begin VB.CommandButton ExitCmd 
            Caption         =   "E&xit"
            Height          =   400
            Left            =   12105
            TabIndex        =   6
            Top             =   600
            Width           =   2700
         End
         Begin VB.CommandButton SearchCmd 
            Caption         =   "Search"
            Height          =   400
            Left            =   12105
            TabIndex        =   5
            Top             =   195
            Width           =   2700
         End
         Begin VB.CommandButton LastCmd 
            Caption         =   "&Last"
            Height          =   400
            Left            =   9165
            TabIndex        =   7
            Top             =   600
            Width           =   2940
         End
         Begin VB.CommandButton DeleteCmd 
            Caption         =   "&Delete"
            Height          =   400
            Left            =   9165
            TabIndex        =   4
            Top             =   195
            Width           =   2940
         End
         Begin VB.CommandButton FirstCmd 
            Caption         =   "&First"
            Height          =   400
            Left            =   6225
            TabIndex        =   8
            Top             =   600
            Width           =   2940
         End
         Begin VB.CommandButton EditCmd 
            Caption         =   "&Edit"
            Height          =   400
            Left            =   6225
            TabIndex        =   3
            Top             =   195
            Width           =   2940
         End
         Begin VB.CommandButton PreviousCmd 
            Caption         =   "&Previous"
            Height          =   400
            Left            =   3165
            TabIndex        =   9
            Top             =   600
            Width           =   3060
         End
         Begin VB.CommandButton SaveCmd 
            Caption         =   "Save"
            Height          =   400
            Left            =   3165
            TabIndex        =   2
            Top             =   195
            Width           =   3060
         End
         Begin VB.CommandButton NextCmd 
            Caption         =   "&Next"
            Height          =   400
            Left            =   105
            TabIndex        =   10
            Top             =   600
            Width           =   3060
         End
         Begin VB.CommandButton AddCmd 
            Caption         =   "&Add New"
            Height          =   400
            Left            =   105
            TabIndex        =   1
            Top             =   195
            Width           =   3060
         End
      End
   End
End
Attribute VB_Name = "FrmMst_SecurityLocationMapping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg, mSecurityAgencyID, mSSecurityAgencyID, mSecurityTypeID, mSecurityDesignationID As Variant
Dim mLocationID, mSLocationID, mCMPID, mGodownID, mDetailDID  As Double
Dim mSalaryProcessed As Variant

Private Sub Grid_HeadGodown()
With MSHFlexGrid3
    .Clear
    .Rows = 3
    .FixedRows = 2
    .Cols = 12

    .TextMatrix(1, 0) = "Security Type ~ Security ID "
    .TextMatrix(1, 1) = "CMP Name"
    .TextMatrix(1, 2) = "Godown No"
    .TextMatrix(1, 3) = "Address"
    .TextMatrix(1, 4) = "Balance Bags"
    .TextMatrix(1, 5) = "Balance Qty"
    .TextMatrix(1, 6) = "AUM"
    .TextMatrix(1, 7) = "Remarks"
    .TextMatrix(1, 8) = "DetailDID" '"Security TypeID"
    .TextMatrix(1, 9) = "CMPID"
    .TextMatrix(1, 10) = "GodownID"
    .TextMatrix(1, 11) = "GodMapDID"
    
    .ColWidth(0) = 1800
    .ColWidth(1) = 1800
    .ColWidth(2) = 1000
    .ColWidth(3) = 2500
    .ColWidth(4) = 1200
    .ColWidth(5) = 1200
    .ColWidth(6) = 1200
    .ColWidth(7) = 3500
    .ColWidth(8) = 0
    .ColWidth(9) = 0
    .ColWidth(10) = 0
    .ColWidth(11) = 0
    
    .RowHeight(1) = 650
    .WordWrap = True
End With
End Sub

Private Sub Grid_Head()

MSHFlexGrid2.Clear
MSHFlexGrid2.Rows = 3
MSHFlexGrid2.FixedRows = 2
MSHFlexGrid2.Cols = 15

MSHFlexGrid2.TextMatrix(1, 0) = "Security ID"
MSHFlexGrid2.TextMatrix(1, 1) = "Security Type"
MSHFlexGrid2.TextMatrix(1, 2) = "Security TypeID"
MSHFlexGrid2.TextMatrix(1, 3) = "Security Designation"
MSHFlexGrid2.TextMatrix(1, 4) = "Security DesignationID"
MSHFlexGrid2.TextMatrix(1, 5) = "Start Date"
MSHFlexGrid2.TextMatrix(1, 6) = "End Date"
MSHFlexGrid2.TextMatrix(1, 7) = "Salary Payable Upto"
MSHFlexGrid2.TextMatrix(1, 8) = "No of Godown"
MSHFlexGrid2.TextMatrix(1, 9) = "Rate per month"
MSHFlexGrid2.TextMatrix(1, 10) = "Remarks"
MSHFlexGrid2.TextMatrix(1, 11) = "Salary Processed Upto"
MSHFlexGrid2.TextMatrix(1, 12) = "Attendance Marked Upto"
MSHFlexGrid2.TextMatrix(1, 13) = "Attendance Started Date"
MSHFlexGrid2.TextMatrix(1, 14) = "Request ID"

MSHFlexGrid2.ColWidth(0) = 700
MSHFlexGrid2.ColWidth(1) = 1800
MSHFlexGrid2.ColWidth(2) = 0
MSHFlexGrid2.ColWidth(3) = 2500
MSHFlexGrid2.ColWidth(4) = 0
MSHFlexGrid2.ColWidth(5) = 1200
MSHFlexGrid2.ColWidth(6) = 1200
MSHFlexGrid2.ColWidth(7) = 1000
MSHFlexGrid2.ColWidth(8) = 1200
MSHFlexGrid2.ColWidth(9) = 1200
MSHFlexGrid2.ColWidth(10) = 3000
MSHFlexGrid2.ColWidth(11) = 1000
MSHFlexGrid2.ColWidth(12) = 1000
MSHFlexGrid2.ColWidth(13) = 0
MSHFlexGrid2.RowHeight(1) = 650
MSHFlexGrid2.WordWrap = True
End Sub
'Private Sub Grid_HeadGodownMapping()
'
'MSHFlexGrid3.Clear
'MSHFlexGrid3.Rows = 3
'MSHFlexGrid3.FixedRows = 2
'MSHFlexGrid3.Cols = 9
'
'MSHFlexGrid3.TextMatrix(0, 0) = "CMP Name"
'MSHFlexGrid3.TextMatrix(0, 1) = "Godown No"
'MSHFlexGrid3.TextMatrix(0, 2) = "Address"
'MSHFlexGrid3.TextMatrix(0, 3) = "Balance Bags"
'MSHFlexGrid3.TextMatrix(0, 4) = "Balance Stock"
'MSHFlexGrid3.TextMatrix(0, 5) = "AUM"
'MSHFlexGrid3.TextMatrix(0, 6) = "Security Type"
'MSHFlexGrid3.TextMatrix(0, 7) = "Godown ID"
'MSHFlexGrid3.TextMatrix(0, 8) = "Security Detail ID"
'
'MSHFlexGrid3.ColWidth(0) = 1800
'MSHFlexGrid3.ColWidth(1) = 1600
'MSHFlexGrid3.ColWidth(2) = 3000
'MSHFlexGrid3.ColWidth(3) = 1200
'MSHFlexGrid3.ColWidth(4) = 1200
'MSHFlexGrid3.ColWidth(5) = 1600
'MSHFlexGrid3.ColWidth(6) = 3000
'MSHFlexGrid3.ColWidth(7) = 0
'MSHFlexGrid3.ColWidth(8) = 0
'
'MSHFlexGrid3.RowHeight(0) = 650
'MSHFlexGrid3.WordWrap = True
'End Sub

Private Sub AddCmd_Click()

Edit_Flg = "A"
Call GButtonLock(Me, False)

Call Grid_Head
Call Grid_HeadGodown

TxtTxnID.Text = ""
TxtTxnDate.Value = Date
TxtSecurityAgency.Visible = False
CmbSecurityAgency.Visible = True
CmbSecurityAgency.Text = ""
CmdSecurityAgency.Visible = True

TxtLocation.Visible = False
CmbLocationID.Visible = True
CmbLocationID.Text = ""
CmdLocationID.Visible = True

TxtCreated = ""
TxtUpdated = ""
mSalaryProcessed = ""

Frame2.Enabled = True
Frame3.Enabled = True
Call ClearFrame

If RsMst_SecLocMapping.RecordCount > 0 Then
   CmbSecurityAgency.SetFocus
End If
End Sub

Private Sub ClearFrame()
CmbSecurityType.Text = ""
CmbSecurityDesignation.Text = ""
TxtStartDate.Value = Date
TxtEndDate.Value = Date + 30
TxtSalaryPaidUpto.Value = Date + 60
TxtNoofGodown = ""
TxtRatePerMonth = ""
TxtRemarks.Text = ""
mSalaryProcessed = ""
CmdAdd2List.Caption = "Add to List"
'''edit facility added on 18 june 2012''''''
CmbSecurityType.Enabled = True
CmbSecurityDesignation.Enabled = True
CmdAdd2List.Visible = False
'''edit facility added on 18 june 2012''''''

End Sub

Private Sub CmbCMP_GotFocus()
If CmbLocationID.Text = "" And TxtLocation.Text = "" Then
    MsgBox "Select Location First!!!"
    CmbLocationID.SetFocus
    Exit Sub
End If
tmp = CmbCMP.Text
Call TableMst_CMP(" where CloseDate is null And LocationID = " & mLocationID)

CmbCMP.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   'CmbCMP.SetFocus
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
If tmp <> CmbCMP.Text Then
    CmbGodown.Text = ""
End If
mCMPID = RsMst_CMP!CMPID

End Sub

Private Sub CmbGodown_GotFocus()
If CmbCMP.Text = "" Then
    MsgBox "Select CMP First!!!"
    CmbCMP.SetFocus
    Exit Sub
End If
tmp = CmbGodown.Text
Call TableMst_Godown(" where CloseDate is null And CMPID = " & mCMPID)

CmbGodown.Clear
If RsMst_Godown.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   'CmbGodown.SetFocus
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
tmp = " Select TIMG.CommodityID, sum(TIMG.BalanceBags) BalanceBags, sum(TIMG.BalanceWeight) BalanceQty"
tmp = tmp & " From Txn_InvMonGSLDetail TIMG"
tmp = tmp & " Inner Join Mst_Godown MG on MG.GodownID = TIMG.GodownID"
tmp = tmp & " Inner Join Mst_CMP MCMP on MCMP.CMPID = MG.CMPID"
tmp = tmp & " Where TIMG.BalanceWeight > 0"
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


Private Sub CmbSecurityAgency_GotFocus()
tmp = CmbSecurityAgency.Text
Call TableMst_SecurityAgency("Where Flag = 'Y'")
CmbSecurityAgency.Clear
If RsMst_SecurityAgency.RecordCount = 0 Then
   MsgBox "No Security Agency Found!!!"
   CmbSecurityAgency.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_SecurityAgency.RecordCount
    CmbSecurityAgency.AddItem RsMst_SecurityAgency!AgencyName
    RsMst_SecurityAgency.MoveNext
Next
CmbSecurityAgency.Text = tmp

End Sub

Private Sub CmbSecurityAgency_LostFocus()
If CmbSecurityAgency.Text = "" Then
   mSecurityAgencyID = Null
   Call ChkDupSecurityLocMapping
   Exit Sub
End If

RsMst_SecurityAgency.MoveFirst
RsMst_SecurityAgency.Find "AgencyName = '" & CmbSecurityAgency.Text & "'"

If RsMst_SecurityAgency.EOF Then
   MsgBox "Invalid Selection "
   CmbSecurityAgency.SetFocus
   Exit Sub
End If
mSecurityAgencyID = RsMst_SecurityAgency!SecurityAgencyID

If LCase(tmp) <> LCase(RsMst_SecurityAgency!AgencyName) Then
   Call ChkDupSecurityLocMapping
End If

End Sub

Private Sub CmbSecurityDesignation_GotFocus()
tmp = CmbSecurityDesignation.Text
Call TableMst_SecurityDesignation("")

CmbSecurityDesignation.Clear
If RsMst_SecurityDesignation.RecordCount = 0 Then
   MsgBox "No Security Designation found!!!"
   CmbSecurityType.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_SecurityDesignation.RecordCount
    CmbSecurityDesignation.AddItem RsMst_SecurityDesignation!SecurityDesignation
    RsMst_SecurityDesignation.MoveNext
Next
CmbSecurityDesignation.Text = tmp
End Sub

Private Sub CmbSecurityDesignation_LostFocus()
If CmbSecurityDesignation.Text = "" Then
   mSecurityDesignationID = Null
   Exit Sub
End If

RsMst_SecurityDesignation.MoveFirst
RsMst_SecurityDesignation.Find "SecurityDesignation = '" & CmbSecurityDesignation.Text & "'"

If RsMst_SecurityDesignation.EOF Then
   MsgBox "Invalid Designation Selection "
   CmbSecurityDesignation.SetFocus
   Exit Sub
End If
mSecurityDesignationID = RsMst_SecurityDesignation!SecurityDesignationID
End Sub

Private Sub CmbSecurityType_GotFocus()
tmp = CmbSecurityType.Text
Call TableMst_SecurityShifts("")

CmbSecurityType.Clear
If RsMst_SecurityShifts.RecordCount = 0 Then
   MsgBox "No Security Type found!!!"
   CmbSecurityType.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_SecurityShifts.RecordCount
    CmbSecurityType.AddItem RsMst_SecurityShifts!SecurityShifts
    RsMst_SecurityShifts.MoveNext
Next
CmbSecurityType.Text = tmp
End Sub

Private Sub CmbSecurityType_LostFocus()
If CmbSecurityType.Text = "" Then
   mSecurityTypeID = Null
   Exit Sub
End If

RsMst_SecurityShifts.MoveFirst
RsMst_SecurityShifts.Find "SecurityShifts = '" & CmbSecurityType.Text & "'"

If RsMst_SecurityShifts.EOF Then
   MsgBox "Invalid Shift Selection "
   CmbSecurityType.SetFocus
   Exit Sub
End If
mSecurityTypeID = RsMst_SecurityShifts!SecurityShiftsID

End Sub

Private Sub CmbSecurityType1_GotFocus()
'tmp = CmbSecurityType1.Text
'Call TableMst_SecurityShifts("")
'CmbSecurityType1.Clear
'If RsMst_SecurityShifts.RecordCount = 0 Then
'   MsgBox "No Security Type found!!!"
'   CmbSecurityType1.SetFocus
'   Exit Sub
'End If
'For I = 1 To RsMst_SecurityShifts.RecordCount
'    CmbSecurityType1.AddItem RsMst_SecurityShifts!SecurityShifts
'    RsMst_SecurityShifts.MoveNext
'Next
'CmbSecurityType1.Text = tmp

tmp = CmbSecurityType1.Text
CmbSecurityType1.Clear
If MSHFlexGrid2.Rows = 3 Then
    MsgBox "No Security Type found in Guard Details"
    SSTab1.Tab = 0
    CmbSecurityType.SetFocus
End If
For I = 2 To MSHFlexGrid2.Rows - 2
    CmbSecurityType1.AddItem MSHFlexGrid2.TextMatrix(I, 1) & "~" & MSHFlexGrid2.TextMatrix(I, 0)
Next

CmbSecurityType1.Text = tmp

End Sub

Private Sub CmbSecurityType1_LostFocus()
If CmbSecurityType1.Text = "" Then Exit Sub

'RsMst_SecurityShifts.MoveFirst
'RsMst_SecurityShifts.Find "SecurityShifts = '" & CmbSecurityType1.Text & "'"
'
'If RsMst_SecurityShifts.EOF Then
'   MsgBox "Invalid Selection "
'   CmbSecurityType1.SetFocus
'   Exit Sub
'End If
'mSecurityTypeID = RsMst_SecurityShifts!SecurityShiftsID

For I = 2 To MSHFlexGrid2.Rows - 2
    If LCase(CmbSecurityType1.Text) = LCase(MSHFlexGrid2.TextMatrix(I, 1)) & "~" & LCase(MSHFlexGrid2.TextMatrix(I, 0)) Then
        mDetailDID = MSHFlexGrid2.TextMatrix(I, 0)
        Exit Sub
    End If
Next
MsgBox "Invalid Selection"

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
   MsgBox "No Location found"
   CmbSLocationID.SetFocus
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
   Exit Sub
End If

RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbSLocation.Text & "'"

If RsMst_Location.EOF Then
   MsgBox "Invalid Location  selection "
   CmbSLocation.SetFocus
   Exit Sub
End If
mSLocationID = RsMst_Location!LocationID
End Sub

Private Sub CmbSSecurityAgency_GotFocus()
Call TableMst_SecurityAgency("Where Flag = 'Y'")

CmbSSecurityAgency.Clear
If RsMst_SecurityAgency.RecordCount = 0 Then
   MsgBox "No Security Agency Found!!!"
   CmbSSecurityAgency.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_SecurityAgency.RecordCount
    CmbSSecurityAgency.AddItem RsMst_SecurityAgency!AgencyName
    RsMst_SecurityAgency.MoveNext
Next
End Sub

Private Sub CmbSSecurityAgency_LostFocus()
If CmbSSecurityAgency.Text = "" Then
   mSSecurityAgencyID = Null
   Exit Sub
End If

RsMst_SecurityAgency.MoveFirst
RsMst_SecurityAgency.Find "AgencyName = '" & CmbSSecurityAgency.Text & "'"

If RsMst_SecurityAgency.EOF Then
   MsgBox "Invalid Selection "
   CmbSSecurityAgency.SetFocus
   Exit Sub
End If
mSSecurityAgencyID = RsMst_SecurityAgency!SecurityAgencyID
End Sub

Private Sub CmdAdd2List_Click()
Dim r As Variant

If CmbSecurityType.Text = "" Then
   MsgBox "Blank Security Type not allowed!!!"
   CmbSecurityType.SetFocus
   Exit Sub
End If

If CmbSecurityDesignation.Text = "" Then
   MsgBox "Blank Security Designation not allowed!!!"
   CmbSecurityDesignation.SetFocus
   Exit Sub
End If

If TxtStartDate.Value > TxtEndDate.Value Then
   MsgBox "'Start date ' must be less than 'End Date'!!!"
   TxtStartDate.SetFocus
   Exit Sub
End If

If TxtSalaryPaidUpto.Value < TxtEndDate.Value Then
   MsgBox "'Salary Payable date ' must be greater than 'End Date'!!!"
'   TxtSalaryPaidUpto.SetFocus
   Exit Sub
End If


If mSalaryProcessed <> "" Then
   'If CDate(Format(mSalaryProcessed, "yyyymm")) > Format(TxtSalaryPaidUpto.Value, "yyyymm") Then
   '   MsgBox "'Salary Payable date ' must be greater than 'Salary Processed' from Attendance Screen!!!"
   '   TxtSalaryPaidUpto.SetFocus
   '   Exit Sub
   'End If
   If CDate(Format(mSalaryProcessed, "yyyymm")) > Format(TxtEndDate.Value, "yyyymm") Then
      MsgBox "'End date ' must be greater than 'Salary Processed' from Attendance Screen!!!"
      TxtEndDate.SetFocus
      Exit Sub
   End If
End If


If TxtNoofGodown = "" Then
   MsgBox "Blank No Of Godowns not allowed!!!"
   TxtNoofGodown.SetFocus
   Exit Sub
End If

If TxtRatePerMonth = "" Then
   MsgBox "Blank Rate per month not allowed!!!"
   TxtRatePerMonth.SetFocus
   Exit Sub
End If

'If TxtRemarks = "" Then
'   MsgBox "Blank Remarks not allowed!!!"
'   TxtRemarks.SetFocus
'   Exit Sub
'End If

If CmdAdd2List.Caption = "Update List" Then
   r = MSHFlexGrid2.RowSel
   'tmp = " Select max(AttendanceDate) AttendanceDate From Txn_SecurityAttendance Where TxnDetailID =  " & Val(MSHFlexGrid2.TextMatrix(r, 0))
   'tmp = tmp & " And AttendanceDate >= '" & TxtSalaryPaidUpto.Value & "' "
   'Call TmpTable
   'If TmpRs.RecordCount > 0 Then
   '   If IsNull(TmpRs!AttendanceDate) = False Then
   '      MsgBox "Salary Payable Upto Date should not less than Last Attendance Date (" & Format(TmpRs!AttendanceDate, "dd-MMM-yyyy") & ")"
   '      Exit Sub
   '   End If
   'End If
   If MSHFlexGrid2.TextMatrix(r, 12) <> "" Then
      If Format(TxtEndDate.Value, "yyyy-mm-dd") < Format(CDate(MSHFlexGrid2.TextMatrix(r, 12)), "yyyy-mm-dd") Then
         MsgBox "End Date should not less than Last Attendance Date (" & Format(CDate(MSHFlexGrid2.TextMatrix(r, 12)), "dd-MMM-yyyy") & ")"
         Exit Sub
      End If
   End If
   If MSHFlexGrid2.TextMatrix(r, 13) <> "" Then
      If Format(TxtStartDate.Value, "yyyy-mm-dd") > Format(CDate(MSHFlexGrid2.TextMatrix(r, 13)), "yyyy-mm-dd") Then
         MsgBox "Start Date should not greater than First Attendance Date (" & Format(CDate(MSHFlexGrid2.TextMatrix(r, 13)), "dd-MMM-yyyy") & ")"
         Exit Sub
      End If
   End If
   MSHFlexGrid2.TextMatrix(0, 8) = Val(MSHFlexGrid2.TextMatrix(0, 8)) - Val(MSHFlexGrid2.TextMatrix(r, 8))
   'TxtNoofGodown = Val(TxtNoofGodown) - Val(MSHFlexGrid2.TextMatrix(r, 5))
Else
   r = MSHFlexGrid2.Rows - 1
   MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
   MSHFlexGrid2.TextMatrix(r, 1) = ""
   MSHFlexGrid2.TextMatrix(r, 0) = ""
End If


MSHFlexGrid2.TextMatrix(r, 1) = CmbSecurityType.Text
MSHFlexGrid2.TextMatrix(r, 2) = mSecurityTypeID
MSHFlexGrid2.TextMatrix(r, 3) = CmbSecurityDesignation.Text
MSHFlexGrid2.TextMatrix(r, 4) = mSecurityDesignationID
MSHFlexGrid2.TextMatrix(r, 5) = TxtStartDate.Value
MSHFlexGrid2.TextMatrix(r, 6) = TxtEndDate.Value
MSHFlexGrid2.TextMatrix(r, 7) = TxtSalaryPaidUpto.Value
MSHFlexGrid2.TextMatrix(r, 8) = Val(TxtNoofGodown)
MSHFlexGrid2.TextMatrix(r, 9) = Val(TxtRatePerMonth)
MSHFlexGrid2.TextMatrix(r, 10) = TxtRemarks

MSHFlexGrid2.TextMatrix(0, 8) = Val(MSHFlexGrid2.TextMatrix(0, 8)) + Val(MSHFlexGrid2.TextMatrix(r, 8))

Call ClearFrame
CmbSecurityType.SetFocus
End Sub

Private Sub CmdAddtoList_Click()
If CmbSecurityType1.Text = "" Then
    MsgBox "Blank Security Type Allowed!!!"
    CmbSecurityType1.SetFocus
    Exit Sub
End If
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
For I = 2 To MSHFlexGrid3.Rows - 2
    If CmdAddtoList.Caption <> "Add to List" Then
        If I <> MSHFlexGrid3.RowSel Then
            If MSHFlexGrid3.TextMatrix(I, 0) = CmbSecurityType1.Text And MSHFlexGrid3.TextMatrix(I, 1) = CmbCMP.Text And MSHFlexGrid3.TextMatrix(I, 2) = CmbGodown.Text Then
                MsgBox "Duplicate Godown not Allowed!!!"
                Exit Sub
            End If
        End If
    Else
        If MSHFlexGrid3.TextMatrix(I, 0) = CmbSecurityType1.Text And MSHFlexGrid3.TextMatrix(I, 1) = CmbCMP.Text And MSHFlexGrid3.TextMatrix(I, 2) = CmbGodown.Text Then
            MsgBox "Duplicate Godown not Allowed!!!"
            Exit Sub
        End If
    End If
Next

If CmdAddtoList.Caption <> "Add to List" Then
    I = MSHFlexGrid3.RowSel
    MSHFlexGrid3.TextMatrix(0, 4) = Val(MSHFlexGrid3.TextMatrix(0, 4)) - Val(MSHFlexGrid3.TextMatrix(I, 4))
    MSHFlexGrid3.TextMatrix(0, 5) = Val(MSHFlexGrid3.TextMatrix(0, 5)) - Val(MSHFlexGrid3.TextMatrix(I, 5))
    MSHFlexGrid3.TextMatrix(0, 6) = Val(MSHFlexGrid3.TextMatrix(0, 6)) - Val(MSHFlexGrid3.TextMatrix(I, 6))
Else
    I = MSHFlexGrid3.Rows - 1
    MSHFlexGrid3.Rows = MSHFlexGrid3.Rows + 1
End If

MSHFlexGrid3.TextMatrix(I, 0) = CmbSecurityType1.Text
MSHFlexGrid3.TextMatrix(I, 1) = CmbCMP.Text
MSHFlexGrid3.TextMatrix(I, 2) = CmbGodown.Text
MSHFlexGrid3.TextMatrix(I, 3) = TxtGodownAddress.Text
MSHFlexGrid3.TextMatrix(I, 4) = TxtBalanceBags.Text
MSHFlexGrid3.TextMatrix(I, 5) = TxtBalanceQty.Text
MSHFlexGrid3.TextMatrix(I, 6) = TxtAUM.Text
MSHFlexGrid3.TextMatrix(I, 7) = TxtRemarksD.Text
MSHFlexGrid3.TextMatrix(I, 8) = mDetailDID
MSHFlexGrid3.TextMatrix(I, 9) = mCMPID
MSHFlexGrid3.TextMatrix(I, 10) = mGodownID

MSHFlexGrid3.TextMatrix(0, 4) = Val(MSHFlexGrid3.TextMatrix(0, 4)) + Val(MSHFlexGrid3.TextMatrix(I, 4))
MSHFlexGrid3.TextMatrix(0, 5) = Val(MSHFlexGrid3.TextMatrix(0, 5)) + Val(MSHFlexGrid3.TextMatrix(I, 5))
MSHFlexGrid3.TextMatrix(0, 6) = Val(MSHFlexGrid3.TextMatrix(0, 6)) + Val(MSHFlexGrid3.TextMatrix(I, 6))
'TxtTotalAUM.Text = MSHFlexGrid2.TextMatrix(0, 5)

For I = 2 To MSHFlexGrid2.Rows - 2
    If Val(MSHFlexGrid2.TextMatrix(I, 0)) = Val(mDetailDID) Then
        MSHFlexGrid2.TextMatrix(I, 8) = Val(MSHFlexGrid2.TextMatrix(I, 8)) + 1
        MSHFlexGrid2.TextMatrix(0, 8) = Val(MSHFlexGrid2.TextMatrix(0, 8)) + 1
    End If
Next

Call ClearGodownMapping
End Sub

Private Sub ClearGodownMapping()
CmbSecurityType1.Enabled = True
CmbCMP.Enabled = True
CmbGodown.Enabled = True
CmbSecurityType1.Text = ""
CmbCMP.Text = ""
CmbGodown.Text = ""
TxtBalanceBags.Text = ""
TxtBalanceQty.Text = ""
TxtAUM.Text = ""
TxtGodownAddress.Text = ""
TxtRemarksD.Text = ""
CmdAddtoList.Caption = "Add to List"
CmdDeleteFromList.Enabled = False
End Sub

Private Sub CmdDeleteFromList_Click()
If CmbCMP.Text = "" Then Exit Sub

I = MSHFlexGrid3.RowSel
If I > 0 Then
   If MSHFlexGrid3.TextMatrix(I, 1) <> "" Then
      mDetailDID = MSHFlexGrid3.TextMatrix(I, 8)
      MSHFlexGrid3.TextMatrix(0, 4) = Val(MSHFlexGrid3.TextMatrix(0, 4)) - Val(MSHFlexGrid3.TextMatrix(I, 4))
      MSHFlexGrid3.TextMatrix(0, 5) = Val(MSHFlexGrid3.TextMatrix(0, 5)) - Val(MSHFlexGrid3.TextMatrix(I, 5))
      MSHFlexGrid3.TextMatrix(0, 6) = Val(MSHFlexGrid3.TextMatrix(0, 6)) - Val(MSHFlexGrid3.TextMatrix(I, 6))
      For RO = I To MSHFlexGrid3.Rows - 2
          For C = 0 To MSHFlexGrid3.Cols - 1
              MSHFlexGrid3.TextMatrix(RO, C) = MSHFlexGrid3.TextMatrix(RO + 1, C)
          Next
      Next
      MSHFlexGrid3.Rows = MSHFlexGrid3.Rows - 1
   End If
End If
'TxtTotalAUM.Text = MSHFlexGrid2.TextMatrix(0, 5)
For I = 2 To MSHFlexGrid2.Rows - 2
    If Val(MSHFlexGrid2.TextMatrix(I, 0)) = Val(mDetailDID) Then
        MSHFlexGrid2.TextMatrix(I, 8) = Val(MSHFlexGrid2.TextMatrix(I, 8)) - 1
        MSHFlexGrid2.TextMatrix(0, 8) = Val(MSHFlexGrid2.TextMatrix(0, 8)) - 1
    End If
Next

Call ClearGodownMapping

End Sub

Private Sub CmdDeleteList_Click()

If CmbSecurityType.Text = "" Or CmbSecurityDesignation.Text = "" Then Exit Sub


I = MSHFlexGrid2.RowSel
If I > 0 Then
   If MSHFlexGrid2.TextMatrix(I, 2) <> "" Then
      TxtNoofGodown = Val(TxtNoofGodown) - Val(MSHFlexGrid2.TextMatrix(I, 8))
      MSHFlexGrid2.TextMatrix(0, 8) = Val(MSHFlexGrid2.TextMatrix(0, 8)) - Val(MSHFlexGrid2.TextMatrix(I, 8))
      For RO = I To MSHFlexGrid2.Rows - 2
          For C = 0 To MSHFlexGrid2.Cols - 1
              MSHFlexGrid2.TextMatrix(RO, C) = MSHFlexGrid2.TextMatrix(RO + 1, C)
          Next
      Next
      MSHFlexGrid2.Rows = MSHFlexGrid2.Rows - 1
   End If
End If

Call ClearFrame
End Sub

Private Sub CmdGo_Click()
Dim wc As Variant

If SFrom.Value > STo.Value Then
   MsgBox "'From Date' must be less than 'To date'!!!"
   SFrom.SetFocus
   Exit Sub
End If

wc = ""

tmp = "Select MSLM.TxnID,MSLM.TxnDate, MSA.AgencyName, ML.LocationName "
tmp = tmp & " From Mst_SecLocMapping MSLM "
tmp = tmp & " Inner Join Mst_SecurityAgency MSA On MSA.SecurityAgencyID = MSLM.SecurityAgencyID "
tmp = tmp & " Inner Join Mst_Location ML On ML.LocationID = MSLM.LocationID"

If Gen_DBType = "MDB" Then
   tmp = tmp & " Where TxnDate Between #" & Format(SFrom, Gen_DatFormat) & "# And #" & Format(STo, Gen_DatFormat) & "#) "
Else
   tmp = tmp & " Where TxnDate Between '" & Format(SFrom, Gen_DatFormat) & "' And '" & Format(STo, Gen_DatFormat) & "' "
End If

If CmbSSecurityAgency.Text <> "" Then
   tmp = tmp & " AND MSA.AgencyName = '" & CmbSSecurityAgency.Text & "' "
End If

If CmbSLocation.Text <> "" Then
   tmp = tmp & " AND ML.LocationName = '" & CmbSLocation.Text & "' "
Else
   If Gen_WcLocation <> "" Then
      tmp = tmp & " AND ML." & Gen_WcLocation
   End If
End If
If TxtSTxnID.Text <> "" Then
    tmp = tmp & " And MSLM.TxnID like '%" & Val(TxtSTxnID) & "%'"
End If
Call TmpTable

Call Grid_Head1
'RsMst_SecLocMapping.Filter = "" & wc & ""

For iu = 1 To TmpRs.RecordCount
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
    For C = 0 To MSHFlexGrid1.Cols - 1
        MSHFlexGrid1.TextMatrix(iu, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
    Next
    TmpRs.MoveNext
Next

Label22.Refresh
Label22.Caption = TmpRs.RecordCount


End Sub

Private Sub CmdLocationID_Click()
FrmMst_Location.Show
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

If TxtTxnID.Text = "" Then Exit Sub

tmp = "SELECT MonthwiseTxnID FROM Txn_MonthwiseSecLoc WHERE SecurityAgencyID = " & mSecurityAgencyID & " And LocationID = " & mLocationID
Call TmpTable

If TmpRs.RecordCount > 0 Then
    MsgBox "Month wise details for Security Agency '" & TxtSecurityAgency.Text & "' and Location '" & TxtLocation.Text & "' already exists."
    Exit Sub
End If

ans = MsgBox("Do you really want to DELETE this record???", vbYesNo)
If ans = vbYes Then
   tmp = "Delete from Mst_SecLocMappingDetail where TxnID = " & Val(TxtTxnID)
   Call TmpTable
   RsMst_SecLocMapping.Delete
   RsMst_SecLocMapping.Update
   If RsMst_SecLocMapping.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_SecLocMapping.MoveNext
   If RsMst_SecLocMapping.EOF() Then
      RsMst_SecLocMapping.MoveLast
   End If
   Call Indata
End If
Exit Sub
msgError:
MsgBox "Child record Present "
RsMst_SecLocMapping.CancelUpdate
End Sub

Private Sub EditCmd_Click()

Edit_Flg = "E"

'tmp = "SELECT MonthwiseTxnID FROM Txn_MonthwiseSecLoc WHERE SecurityAgencyID = " & mSecurityAgencyID & " And LocationID = " & mLocationID
'Call TmpTable

'If TmpRs.RecordCount > 0 Then
'    MsgBox "Month wise details for Security Agency '" & TxtSecurityAgency.Text & "' and Location '" & TxtLocation.Text & "' already exists."
'    Exit Sub
'End If

Call GButtonLock(Me, False)
'TxtSecurityAgency.Visible = False
'CmbSecurityAgency.Visible = True
CmbSecurityAgency.Text = TxtSecurityAgency
'CmdSecurityAgency.Visible = True'

'TxtLocation.Visible = False
'CmbLocationID.Visible = True
CmbLocationID.Text = TxtLocation
'CmdLocationID.Visible = True

Call ClearFrame
Frame2.Enabled = True
Frame3.Enabled = True
'CmbSecurityAgency.SetFocus
End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim wc As Variant

Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Call GButtonLockAD(Me)
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If

Call Grid_Head
Call Grid_Head1
SFrom.Value = CDate(From_GenDate) - 900
STo.Value = Date 'To_GenDate
CmbSSecurityAgency.Text = ""
CmbSLocation.Text = ""
Label22.Caption = ""

Call TableMst_SecLocMapping(" Order by TxnID")

If RsMst_SecLocMapping.RecordCount = 0 Then
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

Private Sub CmbLocationID_GotFocus()

tmp = CmbLocationID.Text
If Gen_WcLocation = "" Then
   Call TableMst_Location
Else
    Call TableMst_Location(" WHERE " & Gen_WcLocation)
End If
CmbLocationID.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Location found"
  CmbLocationID.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbLocationID.AddItem RsMst_Location!LocationName
    RsMst_Location.MoveNext
Next
CmbLocationID.Text = tmp

End Sub

Private Sub CmbLocationID_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "CmbLocationID")
End If
End Sub

Private Sub CmbLocationID_LostFocus()

If CmbLocationID.Text = "" Then
   mLocationID = 0
   Call ChkDupSecurityLocMapping
   Exit Sub
End If

RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocationID.Text & "'"

If RsMst_Location.EOF Then
   MsgBox "Invalid Location  selection "
   CmbLocationID.SetFocus
   Exit Sub
End If
mLocationID = RsMst_Location!LocationID
If LCase(tmp) <> LCase(RsMst_Location!LocationName) Then
   Call ChkDupSecurityLocMapping
End If

End Sub

Private Sub Form_Resize()
If Me.WindowState = vbMaximized Then
   SetScrollBars
End If
End Sub

Private Sub MSHFlexGrid1_Click()
Dim I As Variant
I = MSHFlexGrid1.RowSel
If I <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(I, 0) = "" Then Exit Sub
RsMst_SecLocMapping.MoveFirst
RsMst_SecLocMapping.Find "TxnID = '" & MSHFlexGrid1.TextMatrix(I, 0) & "'"
If Not RsMst_SecLocMapping.EOF Then
   Call Indata
   Frame0.Visible = False
End If
End Sub

Private Sub MSHFlexGrid2_Click()
If SaveCmd.Enabled = False Then Exit Sub


Dim r As Variant
r = MSHFlexGrid2.RowSel
If r < 2 Then Exit Sub
If MSHFlexGrid2.TextMatrix(r, 2) = "" Then Exit Sub

If MSHFlexGrid2.TextMatrix(r, 0) <> "" Then
   tmp = " SELECT TMSD.MonthwiseTxnID "
   tmp = tmp & " FROM Txn_MonthwiseSecLocDetail TMSD"
   tmp = tmp & " Inner Join Txn_MonthwiseSecLoc TMS On TMSD.MonthwiseTxnID = TMS.MonthwiseTxnID "
   tmp = tmp & " Where TMS.SecurityAgencyID = " & mSecurityAgencyID & " And TMS.LocationID = " & mLocationID & " And TMSD.SecLocMapDetailID = " & MSHFlexGrid2.TextMatrix(r, 0) & ""
   
   Call TmpTable
   CmdDeleteList.Enabled = True
    If TmpRs.RecordCount > 0 Then
      'MsgBox "Month wise details for Security Agency '" & TxtSecurityAgency.Text & "' and Location '" & TxtLocation.Text & "' already exists."
      CmdDeleteList.Enabled = False
      'Exit Sub
   End If
End If

CmbSecurityType.Text = MSHFlexGrid2.TextMatrix(r, 1)
mSecurityTypeID = MSHFlexGrid2.TextMatrix(r, 2)
CmbSecurityDesignation.Text = MSHFlexGrid2.TextMatrix(r, 3)
mSecurityDesignationID = MSHFlexGrid2.TextMatrix(r, 4)
TxtStartDate.Value = IIf(MSHFlexGrid2.TextMatrix(r, 5) = "", Date, MSHFlexGrid2.TextMatrix(r, 5))
TxtEndDate.Value = IIf(MSHFlexGrid2.TextMatrix(r, 6) = "", Date, MSHFlexGrid2.TextMatrix(r, 6))
TxtNoofGodown = Val(MSHFlexGrid2.TextMatrix(r, 8))
TxtRatePerMonth = Val(MSHFlexGrid2.TextMatrix(r, 9))
TxtRemarks.Text = MSHFlexGrid2.TextMatrix(r, 10)
TxtSalaryPaidUpto.Value = IIf(MSHFlexGrid2.TextMatrix(r, 7) = "", Date, MSHFlexGrid2.TextMatrix(r, 7))
mSalaryProcessed = MSHFlexGrid2.TextMatrix(r, 11)

CmdAdd2List.Caption = "Update List"
'CmbSecurityType.SetFocus
CmbSecurityType.Enabled = False
CmbSecurityDesignation.Enabled = False
'''edit facility added on 18 june 2012''''''
CmdAdd2List.Visible = True
'''edit facility added on 18 june 2012''''''
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

Private Sub MSHFlexGrid3_Click()
If SaveCmd.Enabled = False Then Exit Sub
Call ClearGodownMapping
I = MSHFlexGrid3.RowSel
If MSHFlexGrid3.TextMatrix(I, 0) = "" Then Exit Sub
CmbSecurityType1.Text = MSHFlexGrid3.TextMatrix(I, 0)
CmbCMP.Text = MSHFlexGrid3.TextMatrix(I, 1)
CmbGodown.Text = MSHFlexGrid3.TextMatrix(I, 2)
TxtGodownAddress.Text = MSHFlexGrid3.TextMatrix(I, 3)
TxtBalanceBags.Text = MSHFlexGrid3.TextMatrix(I, 4)
TxtBalanceQty.Text = MSHFlexGrid3.TextMatrix(I, 5)
TxtAUM.Text = MSHFlexGrid3.TextMatrix(I, 6)
TxtRemarksD = MSHFlexGrid3.TextMatrix(I, 7)
'mSecurityTypeID = MSHFlexGrid3.TextMatrix(I, 8)
mDetailDID = MSHFlexGrid3.TextMatrix(I, 8)
mCMPID = MSHFlexGrid3.TextMatrix(I, 9)
mGodownID = MSHFlexGrid3.TextMatrix(I, 10)
CmdAddtoList.Caption = "Update List"
CmbSecurityType1.Enabled = False
CmbCMP.Enabled = False
CmbGodown.Enabled = False
CmdDeleteFromList.Enabled = True

End Sub

Private Sub SaveCmd_Click()

If CmbSecurityAgency.Text = "" Then
   MsgBox "Blank Security Agency not allowed!!!"
   CmbSecurityAgency.SetFocus
   Exit Sub
End If

If CmbLocationID.Text = "" Then
   MsgBox "Blank Location not allowed!!!"
   CmbLocationID.SetFocus
   Exit Sub
End If

If MSHFlexGrid2.Rows <= 3 Then
   MsgBox "No details found in grid!!!"
   CmbSecurityType.SetFocus
   Exit Sub
End If

If Edit_Flg = "A" Then
   RsMst_SecLocMapping.AddNew
   RsMst_SecLocMapping!TxnID = GetMaxTxnID
   RsMst_SecLocMapping!G_UID = GetGUID
   TxtTxnID = RsMst_SecLocMapping!TxnID
   TxtG_UID = RsMst_SecLocMapping!G_UID
Else
   RsMst_SecLocMapping.MoveFirst
   RsMst_SecLocMapping.Find "TxnID= " & TxtTxnID & ""
End If

RsMst_SecLocMapping!TxnDate = TxtTxnDate.Value
RsMst_SecLocMapping!SecurityAgencyID = mSecurityAgencyID
RsMst_SecLocMapping!LocationID = mLocationID

If IsNull(RsMst_SecLocMapping!CreatedBy) = True Or RsMst_SecLocMapping!CreatedBy = "" Then
   RsMst_SecLocMapping!CreatedBy = Gen_UserLoginID
   RsMst_SecLocMapping!CreatedDate = Now
Else
   RsMst_SecLocMapping!UpdatedBy = Gen_UserLoginID
   RsMst_SecLocMapping!UpdatedDate = Now
End If
RsMst_SecLocMapping.Update

TxtCreated = IIf(IsNull(RsMst_SecLocMapping!CreatedBy), "", RsMst_SecLocMapping!CreatedBy) & " :- " & IIf(IsNull(RsMst_SecLocMapping!CreatedDate), "", RsMst_SecLocMapping!CreatedDate)
TxtUpdated = IIf(IsNull(RsMst_SecLocMapping!UpdatedBy), "", RsMst_SecLocMapping!UpdatedBy) & " :- " & IIf(IsNull(RsMst_SecLocMapping!UpdatedDate), "", RsMst_SecLocMapping!UpdatedDate)

Call SaveSecLocMappingDetails
Call SaveSecLocGodMappingDetails

'RsMst_SecLocMapping.Filter = ""
RsMst_SecLocMapping.Requery
RsMst_SecLocMapping.MoveFirst
RsMst_SecLocMapping.Find "TxnID= " & TxtTxnID & ""

Call Indata

End Sub

Private Sub SaveSecLocGodMappingDetails()
tmp1 = "(0"
For j = 2 To MSHFlexGrid3.Rows - 2
    If MSHFlexGrid3.TextMatrix(j, 11) <> "" Then
        tmp1 = tmp1 & "," & MSHFlexGrid3.TextMatrix(j, 11)
    End If
Next
tmp1 = tmp1 & ")"
tmp = "Delete From Mst_SecLocGodMappingDetail Where TxnId = " & TxtTxnID & " And GodMapDID not in " & tmp1
Call TmpTable

tmp = " Select * From Mst_SecLocGodMappingDetail Where TxnId = " & TxtTxnID
Call Tmp1Table

For j = 2 To MSHFlexGrid3.Rows - 2
    If MSHFlexGrid3.TextMatrix(j, 11) = "" Then
        TmpRs1.AddNew
        TmpRs1!GodMapDID = GetMaxGodMapDID
        'TmpRs1!G_UID = GetGUID
    Else
        TmpRs1.MoveFirst
        TmpRs1.Find "GodMapDID = " & MSHFlexGrid3.TextMatrix(j, 11)
    End If
    TmpRs1!TxnID = TxtTxnID
    TmpRs1!BalanceBags = Val(MSHFlexGrid3.TextMatrix(j, 4))
    TmpRs1!BalanceQty = Val(MSHFlexGrid3.TextMatrix(j, 5))
    TmpRs1!AUM = Val(MSHFlexGrid3.TextMatrix(j, 6))
    TmpRs1!Remarks = MSHFlexGrid3.TextMatrix(j, 7)
    'TmpRs1!SecurityTypeID = Val(MSHFlexGrid3.TextMatrix(j, 8))
    TmpRs1!TxnDetailID = Val(MSHFlexGrid3.TextMatrix(j, 8))
    TmpRs1!CMPID = Val(MSHFlexGrid3.TextMatrix(j, 9))
    TmpRs1!GodownID = Val(MSHFlexGrid3.TextMatrix(j, 10))
    TmpRs1!CreatedBy = Gen_UserLoginID
    TmpRs1!CreatedDate = Now
    TmpRs1.Update
Next
End Sub

Private Function GetMaxGodMapDID() As Double
Dim Retval As Double
tmp = "Select max(GodMapDID) from Mst_SecLocGodMappingDetail"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetMaxGodMapDID = Retval
End Function


Private Sub SaveSecLocMappingDetails()

Call DeleteSecLocMappingDetails

Call TableMst_SecLocMappingDetails("") '(" Where ClaimID = " & TxtClaimID & " AND SM_Prefix = '" & TxtPreFix & "'")

For I = 2 To MSHFlexGrid2.Rows - 1
    If MSHFlexGrid2.TextMatrix(I, 2) = "" Then Exit For
    If MSHFlexGrid2.TextMatrix(I, 0) <> "" Then
       RsMst_SecLocMappingDetails.MoveFirst
       RsMst_SecLocMappingDetails.Find " TxnDetailID =  " & MSHFlexGrid2.TextMatrix(I, 0) & " "
    Else
       RsMst_SecLocMappingDetails.AddNew
       RsMst_SecLocMappingDetails!TxnDetailID = GetMaxTxnDetailsID
       RsMst_SecLocMappingDetails!CreatedBy = Gen_UserLoginID
       RsMst_SecLocMappingDetails!CreatedDate = Now
       RsMst_SecLocMappingDetails!TxnID = TxtTxnID
    End If
    RsMst_SecLocMappingDetails!SecurityTypeID = MSHFlexGrid2.TextMatrix(I, 2)
    RsMst_SecLocMappingDetails!SecurityDesignationID = MSHFlexGrid2.TextMatrix(I, 4)
    RsMst_SecLocMappingDetails!StartDate = MSHFlexGrid2.TextMatrix(I, 5)
    RsMst_SecLocMappingDetails!EndDate = MSHFlexGrid2.TextMatrix(I, 6)
    RsMst_SecLocMappingDetails!NoofGodown = Val(MSHFlexGrid2.TextMatrix(I, 8))
    RsMst_SecLocMappingDetails!RatePerMonth = Val(MSHFlexGrid2.TextMatrix(I, 9))
    RsMst_SecLocMappingDetails!Remarks = Trim(MSHFlexGrid2.TextMatrix(I, 10))
    RsMst_SecLocMappingDetails!SalaryPaidUpto = MSHFlexGrid2.TextMatrix(I, 7)
    RsMst_SecLocMappingDetails.Update
    RsMst_SecLocMappingDetails.Requery
    
Next
End Sub

Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   If RsMst_SecLocMapping.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   RsMst_SecLocMapping.MoveFirst
   RsMst_SecLocMapping.Find "TxnID= " & Val(TxtTxnID) & ""
   If RsMst_SecLocMapping.EOF Then
      RsMst_SecLocMapping.MoveLast
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
MSHFlexGrid1.Height = Frame0.Height - 1400

RsMst_SecLocMapping.MoveFirst

''===========comment on 7 july 2012 for not clearing search grid=============
'Call Grid_Head1
'SFrom.Value = CDate(From_GenDate) - 900
'STo.Value = Date 'To_GenDate
'CmbSSecurityAgency.Text = ""
'CmbSLocation.Text = ""
'Label22.Caption = ""
''===========comment on 7 july 2012 for not clearing search grid=============

End Sub

Private Sub TxtEndDate_LostFocus()
TxtSalaryPaidUpto = TxtEndDate.Value
End Sub

Private Sub TxtNoofGodown_LostFocus()
If TxtNoofGodown = "" Then Exit Sub
If IsNumeric(TxtNoofGodown) = False Or Val(TxtNoofGodown) <= 0 Then
   MsgBox "Invalid Format!!!"
   TxtNoofGodown.SetFocus
   Exit Sub
End If
End Sub

Private Sub TxtRatePerMonth_LostFocus()
If TxtRatePerMonth = "" Then Exit Sub
If IsNumeric(TxtRatePerMonth) = False Or Val(TxtRatePerMonth) < 0 Then
   MsgBox "Invalid Format!!!"
   TxtRatePerMonth.SetFocus
   Exit Sub
End If

TxtRatePerMonth = Format(TxtRatePerMonth, "########0.00")
End Sub

Private Sub DeleteSecLocMappingDetails()

mDetailID = ""
For I = 2 To MSHFlexGrid2.Rows - 1
    If MSHFlexGrid2.TextMatrix(I, 2) <> "" Then
       If MSHFlexGrid2.TextMatrix(I, 0) <> "" Then
          If mDetailID = "" Then
             mDetailID = "(" & MSHFlexGrid2.TextMatrix(I, 0)
          Else
             mDetailID = mDetailID & "," & MSHFlexGrid2.TextMatrix(I, 0)
          End If
       End If
    End If
    If I = MSHFlexGrid2.Rows - 1 Then
       If mDetailID <> "" Then
          mDetailID = mDetailID & ")"
       End If
    End If
Next

If mDetailID = "" Then
   tmp = "Delete From Mst_SecLocMappingDetail Where TxnID=" & TxtTxnID & " "
Else
   tmp = "Delete From Mst_SecLocMappingDetail Where TxnID=" & TxtTxnID & " And TxnDetailID not in " & mDetailID & ""
End If

Call Tmp4Table


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
Retval = 0
tmp = "Select max(TxnDetailID) From Mst_SecLocMappingDetail " 'Where ClaimID = " & TxtClaimID & " And SM_prefix = '" & TxtPreFix & "'"
Call TmpTable
If TmpRs.RecordCount > 0 Then
   Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
End If
Retval = Retval + 1
TmpRs.Close
GetMaxTxnDetailsID = Retval

End Function
Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_SecLocMapping.MoveNext
If RsMst_SecLocMapping.EOF Then
   RsMst_SecLocMapping.MoveLast
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
RsMst_SecLocMapping.MovePrevious
If RsMst_SecLocMapping.BOF Then
   RsMst_SecLocMapping.MoveFirst
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
RsMst_SecLocMapping.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsMst_SecLocMapping.MoveLast
Call Indata
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False
End Sub
Public Sub Indata()

Call ClearFrame
Frame2.Enabled = False
Frame3.Enabled = False

TxtSecurityAgency.Visible = True
CmbSecurityAgency.Visible = False
CmdSecurityAgency.Visible = Fasle

TxtLocation.Visible = True
CmbLocationID.Visible = False
CmdLocationID.Visible = False

TxtTxnID = RsMst_SecLocMapping!TxnID
TxtTxnDate.Value = RsMst_SecLocMapping!TxnDate
If IsNull(RsMst_SecLocMapping!AgreementStartDate) = False Then
    TxtAgreeStartDate.Value = RsMst_SecLocMapping!AgreementStartDate
    TxtAgreeEndDate.Value = RsMst_SecLocMapping!AgreementEndDate
End If
mSecurityAgencyID = RsMst_SecLocMapping!SecurityAgencyID
TxtSecurityAgency = GetSecurityAgency(RsMst_SecLocMapping!SecurityAgencyID, "N")
mLocationID = RsMst_SecLocMapping!LocationID
TxtLocation = GetLocationName(RsMst_SecLocMapping!LocationID)
TxtG_UID = IIf(IsNull(RsMst_SecLocMapping!G_UID), "", RsMst_SecLocMapping!G_UID)
TxtCreated = IIf(IsNull(RsMst_SecLocMapping!CreatedBy), "", RsMst_SecLocMapping!CreatedBy) & " :- " & IIf(IsNull(RsMst_SecLocMapping!CreatedDate), "", RsMst_SecLocMapping!CreatedDate)
TxtUpdated = IIf(IsNull(RsMst_SecLocMapping!UpdatedBy), "", RsMst_SecLocMapping!UpdatedBy) & " :- " & IIf(IsNull(RsMst_SecLocMapping!UpdatedDate), "", RsMst_SecLocMapping!UpdatedDate)

Call InDataSecLocMappingDetails
Call InDataSecLocGodMappingDetails

If RsMst_SecLocMapping.RecordCount = 1 Then
   Call GButtonLock_1(Me, True)
Else
   Call GButtonLock(Me, True)
End If

End Sub

Private Sub InDataSecLocGodMappingDetails()
Call Grid_HeadGodown
tmp = " select "
tmp = tmp & " MSLGMD.GodMapDID, MSS.SecurityShifts, MCMP.CMPName, MG.GodownNo, MG.Address, MSLGMD.BalanceBags,"
tmp = tmp & " MSLGMD.BalanceQty , MSLGMD.AUM, MSLGMD.Remarks, MSLGMD.TxnDetailID, MSLGMD.CMPID, MSLGMD.GodownID" 'MSLGMD.SecurityTypeID,
tmp = tmp & " from Mst_SecLocGodMappingDetail MSLGMD"
tmp = tmp & " Inner Join Mst_SecLocMappingDetail MSLMD on MSLMD.TxnDetailID = MSLGMD.TxnDetailID"
tmp = tmp & " Inner Join Mst_securityShifts MSS on MSS.SecurityShiftsID = MSLMD.SecurityTypeID"
tmp = tmp & " Inner Join Mst_CMP MCMP on MCMP.CMPID =  MSLGMD.CMPID"
tmp = tmp & " Inner Join Mst_Godown MG on MG.GodownID =  MSLGMD.GodownID"
tmp = tmp & " Where MSLGMD.TxnID = " & TxtTxnID
Call TmpTable

If TmpRs.RecordCount > 0 Then
    MSHFlexGrid3.Rows = MSHFlexGrid3.Rows + TmpRs.RecordCount
    With MSHFlexGrid3
        For I = 2 To TmpRs.RecordCount + 1
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
            .TextMatrix(0, 4) = Val(.TextMatrix(I, 4)) + Val(.TextMatrix(0, 4))
            .TextMatrix(0, 5) = Val(.TextMatrix(I, 5)) + Val(.TextMatrix(0, 5))
            .TextMatrix(0, 6) = Val(.TextMatrix(I, 6)) + Val(.TextMatrix(0, 6))
            TmpRs.MoveNext
        Next
    End With
End If
End Sub
Private Sub InDataSecLocMappingDetails()

Call Grid_Head

'Call TableMst_SecLocMappingDetails(" Where TxnID= " & TxtTxnID & " ") 'AND SM_Prefix = '" & TxtPreFix & "'")

tmp = "Select MSLMD.TxnDetailID,MST.SecurityShifts,MSLMD.SecurityTypeID "
tmp = tmp & " ,MSD.SecurityDesignation,"
tmp = tmp & " MSLMD.SecurityDesignationID,MSLMD.StartDate,MSLMD.EndDate,MSLMD.NoofGodown,MSLMD.RatePerMonth,MSLMD.Remarks,"
tmp = tmp & " MSLMD.SalaryPaidUpto,Max(TMSL.MonthYear) as 'SalaryProcessed',MSLMD.RequestID "
tmp = tmp & " from Mst_SecLocMappingDetail MSLMD"
tmp = tmp & " Inner Join Mst_SecurityShifts MST On MSLMD.SecurityTypeID = MST.SecurityShiftsID"
tmp = tmp & " Inner Join Mst_SecurityDesignation MSD On MSLMD.SecurityDesignationID = MSD.SecurityDesignationID"
tmp = tmp & " Left Join  Txn_MonthwiseSecLocDetail TMSLD on MSLMD.TxnDetailID  = TMSLD.SecLocMapDetailID"
tmp = tmp & " Left Join  Txn_MonthwiseSecLoc TMSL On TMSLD.MonthwiseTxnID = TMSL.MonthwiseTxnID"
tmp = tmp & " Where MSLMD.TxnID = " & TxtTxnID & ""
tmp = tmp & " Group By MSLMD.TxnDetailID,MST.SecurityShifts,MSLMD.SecurityTypeID,"
tmp = tmp & " MSD.SecurityDesignation,"
tmp = tmp & " MSLMD.SecurityDesignationID,MSLMD.StartDate,MSLMD.EndDate,MSLMD.NoofGodown,MSLMD.RatePerMonth,MSLMD.Remarks,"
tmp = tmp & " MSLMD.SalaryPaidUpto,MSLMD.RequestID  Order By MSLMD.TxnDetailID"

Call TmpTable

tmp = " Select min(TSA.AttendanceDate) minAttendanceDate , max(TSA.AttendanceDate) AttendanceDate ,TSA.TxnDetailID  From Txn_SecurityAttendance TSA"
tmp = tmp & " Inner Join Mst_SecLocMappingDetail MSLMD on MSLMD.TxnDetailID =  TSA.TxnDetailID"
tmp = tmp & " Where MSLMD.TxnID = " & Val(TxtTxnID) & " And TSA.Status = 'A' "
tmp = tmp & " Group by TSA.TxnDetailID"
Call Tmp1Table

If TmpRs.RecordCount > 0 Then 'If RsMst_SecLocMappingDetails.RecordCount > 0 Then
    MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + TmpRs.RecordCount
    For I = 2 To MSHFlexGrid2.Rows - 2
        MSHFlexGrid2.TextMatrix(I, 1) = TmpRs!SecurityShifts 'GetSecurityType(RsMst_SecLocMappingDetails!SecurityTypeID)
        MSHFlexGrid2.TextMatrix(I, 2) = TmpRs!SecurityTypeID
        MSHFlexGrid2.TextMatrix(I, 3) = TmpRs!SecurityDesignation 'GetSecurityDesignation(RsMst_SecLocMappingDetails!SecurityDesignationID)
        MSHFlexGrid2.TextMatrix(I, 4) = TmpRs!SecurityDesignationID
        MSHFlexGrid2.TextMatrix(I, 5) = TmpRs!StartDate
        MSHFlexGrid2.TextMatrix(I, 6) = TmpRs!EndDate
        MSHFlexGrid2.TextMatrix(I, 8) = TmpRs!NoofGodown
        MSHFlexGrid2.TextMatrix(0, 8) = Val(MSHFlexGrid2.TextMatrix(0, 8)) + Val(TmpRs!NoofGodown)
        MSHFlexGrid2.TextMatrix(I, 9) = TmpRs!RatePerMonth
        'MSHFlexGrid2.TextMatrix(0, 7) = Val(MSHFlexGrid2.TextMatrix(0, 7)) + Val(MSHFlexGrid2.TextMatrix(i, 7))
        MSHFlexGrid2.TextMatrix(I, 10) = Trim(TmpRs!Remarks)
        MSHFlexGrid2.TextMatrix(I, 0) = TmpRs!TxnDetailID
        MSHFlexGrid2.TextMatrix(I, 7) = TmpRs!SalaryPaidUpto
        MSHFlexGrid2.TextMatrix(I, 11) = IIf(IsNull(TmpRs!SalaryProcessed), "", TmpRs!SalaryProcessed)
        MSHFlexGrid2.TextMatrix(I, 14) = IIf(IsNull(TmpRs!RequestID), "", TmpRs!RequestID)
        If TmpRs1.RecordCount > 0 Then
            TmpRs1.MoveFirst
            TmpRs1.Find "TxnDetailID = " & Val(MSHFlexGrid2.TextMatrix(I, 0))
            If Not TmpRs1.EOF Then
                MSHFlexGrid2.TextMatrix(I, 12) = IIf(IsNull(TmpRs1!AttendanceDate), "", TmpRs1!AttendanceDate)
                MSHFlexGrid2.TextMatrix(I, 13) = IIf(IsNull(TmpRs1!minAttendanceDate), "", TmpRs1!minAttendanceDate)
            End If
        End If
        TmpRs.MoveNext
        
    Next
End If
End Sub
Private Sub Grid_Head1()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.FixedRows = 1
MSHFlexGrid1.Cols = 4

MSHFlexGrid1.TextMatrix(0, 0) = "Txn ID"
MSHFlexGrid1.TextMatrix(0, 1) = "Txn Date"
MSHFlexGrid1.TextMatrix(0, 2) = "Security Agency"
MSHFlexGrid1.TextMatrix(0, 3) = "Location Name"


MSHFlexGrid1.ColWidth(0) = 1000
MSHFlexGrid1.ColWidth(1) = 1200
MSHFlexGrid1.ColWidth(2) = 3000
MSHFlexGrid1.ColWidth(3) = 3500
   
End Sub
Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Mst_SecLocMapping.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Mst_SecLocMapping.xls")
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


Private Sub TxtSTxnID_LostFocus()
If TxtSTxnID = "" Then Exit Sub
If IsNumeric(TxtSTxnID) = False Or Val(TxtSTxnID) < 0 Then
   MsgBox "Invalid Format!!!"
   TxtSTxnID.SetFocus
   Exit Sub
End If
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

Private Sub ChkDupSecurityLocMapping()
Dim Retval As Boolean
Retval = True


If CmbSecurityAgency.Text = "" Or CmbLocationID.Text = "" Then
   Call Grid_Head
   Call Grid_HeadGodown
   Exit Sub
End If
'   MsgBox "Either Security agency or Location Name is blank."
'   CmbSecurityAgency.SetFocus
'   Exit Sub
'End If

'tmp = "SecurityAgencyID = " & mSecurityAgencyID & " And LocationID = " & mLocationID
'RsMst_SecLocMapping.Filter = tmp
tmp = " SELECT MSLM.TxnID, MSLM.TxnDate, AgencyName, LocationName, MSLM.G_UID, MSLM.CreatedBy, MSLM.UpdatedBy, MSLM.CreatedDate, MSLM.UpdatedDate "
tmp = tmp & " FROM Mst_SecLocMapping MSLM"
tmp = tmp & " Inner Join Mst_SecurityAgency MSA On MSA.SecurityAgencyID = MSLM.SecurityAgencyID "
tmp = tmp & " Inner Join Mst_Location ML On ML.LocationID = MSLM.LocationID"
tmp = tmp & " Where MSLM.SecurityAgencyID = " & mSecurityAgencyID & " And ML.LocationID = " & mLocationID & ""

Call TmpTable


If TmpRs.RecordCount > 0 Then
   TxtTxnID = TmpRs!TxnID
   TxtTxnDate.Value = TmpRs!TxnDate
   TxtSecurityAgency = TmpRs!AgencyName
   TxtLocation = TmpRs!LocationName
   TxtG_UID = IIf(IsNull(TmpRs!G_UID), "", TmpRs!G_UID)
   TxtCreated = IIf(IsNull(TmpRs!CreatedBy), "", TmpRs!CreatedBy) & " :- " & IIf(IsNull(TmpRs!CreatedDate), "", TmpRs!CreatedDate)
   TxtUpdated = IIf(IsNull(TmpRs!UpdatedBy), "", TmpRs!UpdatedBy) & " :- " & IIf(IsNull(TmpRs!UpdatedDate), "", TmpRs!UpdatedDate)
    
   Call InDataSecLocMappingDetails
   Call InDataSecLocGodMappingDetails
   Edit_Flg = "E"
Else
   Edit_Flg = "A"
   TxtTxnID.Text = ""
   TxtTxnDate.Value = Date
   TxtCreated = ""
   TxtUpdated = ""
   Call ClearFrame
   Call Grid_Head
   Call Grid_HeadGodown
End If

End Sub

