VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmTxn_GodownApproval 
   Caption         =   "Godown Rent Claim"
   ClientHeight    =   7065
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10785
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   7065
   ScaleWidth      =   10785
   WindowState     =   2  'Maximized
   Begin VB.VScrollBar VScroll1 
      Height          =   10095
      Left            =   14925
      TabIndex        =   102
      Top             =   0
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      TabIndex        =   101
      Top             =   9885
      Width           =   10935
   End
   Begin VB.Frame Frame4 
      Height          =   9735
      Left            =   45
      TabIndex        =   45
      Top             =   30
      Width           =   14940
      Begin VB.Frame Frame0 
         Height          =   3015
         Left            =   360
         TabIndex        =   46
         Top             =   4080
         Visible         =   0   'False
         Width           =   14745
         Begin VB.TextBox TxtSBatchID 
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
            TabIndex        =   64
            TabStop         =   0   'False
            Top             =   1260
            Width           =   1260
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
            Height          =   345
            Left            =   10395
            TabIndex        =   67
            Top             =   1260
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
            Left            =   11520
            TabIndex        =   68
            Top             =   1260
            Width           =   1125
         End
         Begin VB.CheckBox ChkBatchNo 
            Appearance      =   0  'Flat
            Caption         =   "Check1"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3300
            TabIndex        =   62
            Top             =   1320
            Width           =   225
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
            ItemData        =   "FrmTxn_GodownApproval.frx":0000
            Left            =   8820
            List            =   "FrmTxn_GodownApproval.frx":0013
            Sorted          =   -1  'True
            TabIndex        =   66
            Top             =   1260
            Width           =   1485
         End
         Begin VB.TextBox TxtSClaimID 
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
            Left            =   120
            TabIndex        =   63
            TabStop         =   0   'False
            Top             =   1275
            Width           =   1260
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
            Left            =   6990
            Sorted          =   -1  'True
            TabIndex        =   61
            TabStop         =   0   'False
            Top             =   570
            Width           =   3870
         End
         Begin VB.ComboBox CmbSForwardTo 
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
            Sorted          =   -1  'True
            TabIndex        =   65
            TabStop         =   0   'False
            Top             =   1275
            Width           =   3795
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
            Left            =   3120
            Sorted          =   -1  'True
            TabIndex        =   60
            Top             =   585
            Width           =   3765
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   1455
            Left            =   90
            TabIndex        =   70
            Top             =   1740
            Width           =   14355
            _ExtentX        =   25321
            _ExtentY        =   2566
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
         Begin MSComCtl2.DTPicker STo 
            Height          =   360
            Left            =   1620
            TabIndex        =   59
            TabStop         =   0   'False
            Top             =   585
            Width           =   1380
            _ExtentX        =   2434
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
            Format          =   106037249
            CurrentDate     =   36494
         End
         Begin MSComCtl2.DTPicker SFrom 
            Height          =   360
            Left            =   60
            TabIndex        =   58
            TabStop         =   0   'False
            Top             =   570
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   635
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
            Format          =   106037249
            CurrentDate     =   36494
         End
         Begin VB.Label Label9 
            Caption         =   "Batch ID"
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
            Left            =   3825
            TabIndex        =   103
            Top             =   960
            Width           =   810
         End
         Begin VB.Label Label44 
            Caption         =   "Claim Entered From"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   135
            TabIndex        =   79
            Top             =   105
            Width           =   1290
         End
         Begin VB.Label Label45 
            Caption         =   "Claim Entered Upto"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   1620
            TabIndex        =   78
            Top             =   105
            Width           =   1335
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
            Left            =   13605
            TabIndex        =   77
            Top             =   615
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
            Left            =   11085
            TabIndex        =   76
            Top             =   615
            Width           =   2415
         End
         Begin VB.Label Label30 
            Caption         =   "Approved By FAG"
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
            Left            =   1590
            TabIndex        =   75
            Top             =   1320
            Width           =   1650
         End
         Begin VB.Label Label16 
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
            Height          =   300
            Left            =   9240
            TabIndex        =   74
            Top             =   975
            Width           =   600
         End
         Begin VB.Label Label19 
            Caption         =   "Claim ID"
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
            Left            =   345
            TabIndex        =   73
            Top             =   975
            Width           =   810
         End
         Begin VB.Label Label37 
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
            Height          =   225
            Left            =   8235
            TabIndex        =   72
            Top             =   195
            Width           =   1260
         End
         Begin VB.Label Label31 
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
            Height          =   225
            Left            =   5880
            TabIndex        =   71
            Top             =   960
            Width           =   1275
         End
         Begin VB.Label Label36 
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
            Left            =   4425
            TabIndex        =   69
            Top             =   180
            Width           =   960
         End
      End
      Begin VB.Frame Frame1 
         Height          =   8445
         Left            =   120
         TabIndex        =   81
         Top             =   135
         Width           =   14730
         Begin VB.ComboBox CmbBillingState 
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
            Left            =   10830
            Sorted          =   -1  'True
            TabIndex        =   16
            Top             =   1920
            Width           =   3735
         End
         Begin VB.TextBox TxtLessorID 
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
            Left            =   1485
            Locked          =   -1  'True
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   666
            Width           =   7335
         End
         Begin VB.TextBox TxtClaimID 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   375
            Left            =   1485
            Locked          =   -1  'True
            TabIndex        =   0
            TabStop         =   0   'False
            Top             =   240
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
            Height          =   720
            Left            =   1500
            MaxLength       =   200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   15
            Top             =   1905
            Width           =   7320
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
            Left            =   13440
            Locked          =   -1  'True
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   240
            Visible         =   0   'False
            Width           =   1080
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
            ItemData        =   "FrmTxn_GodownApproval.frx":0045
            Left            =   1485
            List            =   "FrmTxn_GodownApproval.frx":0055
            Sorted          =   -1  'True
            TabIndex        =   7
            Top             =   1077
            Width           =   3000
         End
         Begin VB.TextBox TxtClaimAmount 
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
            Left            =   10830
            Locked          =   -1  'True
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   666
            Width           =   1710
         End
         Begin VB.TextBox TxtFWDDesignation 
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
            Left            =   10830
            Locked          =   -1  'True
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   1488
            Width           =   3720
         End
         Begin VB.CommandButton CmdForwardTo 
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
            Left            =   8880
            TabIndex        =   12
            Top             =   1530
            Width           =   375
         End
         Begin VB.ComboBox CmbForwardTo 
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
            Left            =   1485
            Sorted          =   -1  'True
            TabIndex        =   11
            Top             =   1488
            Width           =   7335
         End
         Begin VB.TextBox TxtServiceTax 
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
            Left            =   10830
            Locked          =   -1  'True
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   1077
            Width           =   3720
         End
         Begin VB.TextBox TxtApprovalRemarks 
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
            Left            =   1500
            Locked          =   -1  'True
            MaxLength       =   200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   17
            Top             =   2670
            Width           =   13080
         End
         Begin TabDlg.SSTab SSTab1 
            Height          =   5295
            Left            =   120
            TabIndex        =   18
            Top             =   3075
            Width           =   14535
            _ExtentX        =   25638
            _ExtentY        =   9340
            _Version        =   393216
            Tab             =   1
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
            TabCaption(0)   =   "Expense Details"
            TabPicture(0)   =   "FrmTxn_GodownApproval.frx":007F
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "Frame2"
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Payment Details"
            TabPicture(1)   =   "FrmTxn_GodownApproval.frx":009B
            Tab(1).ControlEnabled=   -1  'True
            Tab(1).Control(0)=   "Frame3"
            Tab(1).Control(0).Enabled=   0   'False
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "Log History"
            TabPicture(2)   =   "FrmTxn_GodownApproval.frx":00B7
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "TxtUpdatedByFAG"
            Tab(2).Control(0).Enabled=   0   'False
            Tab(2).Control(1)=   "TxtCreated"
            Tab(2).Control(1).Enabled=   0   'False
            Tab(2).Control(2)=   "TxtUpdated"
            Tab(2).Control(2).Enabled=   0   'False
            Tab(2).Control(3)=   "MSHFlexGrid4"
            Tab(2).Control(4)=   "Label24"
            Tab(2).Control(5)=   "Label21"
            Tab(2).Control(6)=   "Label20"
            Tab(2).ControlCount=   7
            Begin VB.Frame Frame2 
               Height          =   4860
               Left            =   -74925
               TabIndex        =   83
               Top             =   315
               Width           =   14400
               Begin VB.TextBox TxtArrears 
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
                  Left            =   240
                  TabIndex        =   20
                  TabStop         =   0   'False
                  Top             =   960
                  Visible         =   0   'False
                  Width           =   2175
               End
               Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
                  Height          =   4530
                  Left            =   120
                  TabIndex        =   19
                  Top             =   240
                  Width           =   14175
                  _ExtentX        =   25003
                  _ExtentY        =   7990
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
            Begin VB.TextBox TxtUpdatedByFAG 
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
               Left            =   -72615
               Locked          =   -1  'True
               TabIndex        =   43
               TabStop         =   0   'False
               Top             =   1320
               Width           =   9855
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
               Left            =   -72615
               Locked          =   -1  'True
               TabIndex        =   41
               TabStop         =   0   'False
               Top             =   480
               Width           =   9855
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
               Left            =   -72615
               Locked          =   -1  'True
               TabIndex        =   42
               TabStop         =   0   'False
               Top             =   900
               Width           =   9855
            End
            Begin VB.Frame Frame3 
               Height          =   4815
               Left            =   120
               TabIndex        =   82
               Top             =   345
               Width           =   14295
               Begin VB.Frame FrmFAG 
                  Height          =   3360
                  Left            =   90
                  TabIndex        =   110
                  Top             =   1395
                  Width           =   14115
                  Begin VB.TextBox TxtPaidDate 
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
                     Left            =   12210
                     TabIndex        =   40
                     TabStop         =   0   'False
                     Top             =   2895
                     Width           =   1815
                  End
                  Begin VB.TextBox TxtBankID 
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
                     Left            =   2820
                     TabIndex        =   36
                     TabStop         =   0   'False
                     Top             =   2475
                     Width           =   4815
                  End
                  Begin VB.TextBox TxtTDS 
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
                     Left            =   8850
                     TabIndex        =   39
                     TabStop         =   0   'False
                     Top             =   2895
                     Width           =   1815
                  End
                  Begin VB.TextBox TxtPaidAmt 
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
                     Left            =   2820
                     TabIndex        =   38
                     TabStop         =   0   'False
                     Top             =   2895
                     Width           =   4815
                  End
                  Begin VB.TextBox TxtChequeNo 
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
                     Left            =   8850
                     TabIndex        =   37
                     TabStop         =   0   'False
                     Top             =   2475
                     Width           =   5175
                  End
                  Begin VB.TextBox TxtExportedDate 
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
                     Left            =   12240
                     Locked          =   -1  'True
                     MaxLength       =   18
                     TabIndex        =   35
                     TabStop         =   0   'False
                     Top             =   2070
                     Width           =   1785
                  End
                  Begin VB.TextBox TxtEntryNo 
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
                     Left            =   8850
                     Locked          =   -1  'True
                     MaxLength       =   18
                     TabIndex        =   34
                     TabStop         =   0   'False
                     Top             =   2070
                     Width           =   1860
                  End
                  Begin VB.TextBox TxtDocNo 
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
                     Left            =   2820
                     Locked          =   -1  'True
                     MaxLength       =   18
                     TabIndex        =   33
                     TabStop         =   0   'False
                     Top             =   2070
                     Width           =   4815
                  End
                  Begin VB.ComboBox CmbModeOfPay 
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
                     ItemData        =   "FrmTxn_GodownApproval.frx":00D3
                     Left            =   2820
                     List            =   "FrmTxn_GodownApproval.frx":00E3
                     TabIndex        =   27
                     Top             =   810
                     Width           =   1815
                  End
                  Begin VB.ComboBox CmbFAGBank 
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
                     ItemData        =   "FrmTxn_GodownApproval.frx":010A
                     Left            =   7125
                     List            =   "FrmTxn_GodownApproval.frx":010C
                     TabIndex        =   29
                     Top             =   810
                     Width           =   4575
                  End
                  Begin VB.TextBox TxtFAGDisbAmt 
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
                     Left            =   2820
                     MaxLength       =   18
                     TabIndex        =   25
                     Top             =   390
                     Width           =   1785
                  End
                  Begin VB.TextBox TxtFAGRemark 
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
                     Left            =   2820
                     MaxLength       =   250
                     MultiLine       =   -1  'True
                     ScrollBars      =   2  'Vertical
                     TabIndex        =   31
                     Top             =   1275
                     Width           =   4800
                  End
                  Begin VB.TextBox TxtDeductionRemark 
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
                     Left            =   8850
                     MaxLength       =   1000
                     MultiLine       =   -1  'True
                     ScrollBars      =   2  'Vertical
                     TabIndex        =   32
                     Top             =   1275
                     Width           =   5160
                  End
                  Begin MSComCtl2.DTPicker TxtFagDisbDate 
                     Height          =   375
                     Left            =   7125
                     TabIndex        =   26
                     Top             =   383
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
                     Format          =   106037249
                     CurrentDate     =   40255
                  End
                  Begin VB.TextBox TxtFAGBank 
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
                     Left            =   7125
                     Locked          =   -1  'True
                     TabIndex        =   30
                     Top             =   810
                     Width           =   4575
                  End
                  Begin VB.TextBox TxtModeOfPayment 
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
                     Left            =   2820
                     Locked          =   -1  'True
                     TabIndex        =   28
                     Top             =   810
                     Width           =   1800
                  End
                  Begin VB.Label Label8 
                     Caption         =   "Paid Date"
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
                     Left            =   10845
                     TabIndex        =   125
                     Top             =   2940
                     Width           =   900
                  End
                  Begin VB.Label Label7 
                     Caption         =   "TDS"
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
                     Left            =   7725
                     TabIndex        =   124
                     Top             =   2955
                     Width           =   900
                  End
                  Begin VB.Label Label6 
                     Caption         =   "Paid Amount "
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
                     Left            =   165
                     TabIndex        =   123
                     Top             =   2940
                     Width           =   2445
                  End
                  Begin VB.Label Label5 
                     Caption         =   "Cheque No"
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
                     Left            =   7725
                     TabIndex        =   122
                     Top             =   2520
                     Width           =   1140
                  End
                  Begin VB.Label Label1 
                     Caption         =   "Bank Name"
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
                     Left            =   165
                     TabIndex        =   121
                     Top             =   2565
                     Width           =   1140
                  End
                  Begin VB.Label Label77 
                     AutoSize        =   -1  'True
                     Caption         =   "Exported Date"
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
                     Left            =   10845
                     TabIndex        =   120
                     Top             =   2130
                     Width           =   1290
                  End
                  Begin VB.Label Label76 
                     AutoSize        =   -1  'True
                     Caption         =   "Entry No"
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
                     Left            =   7725
                     TabIndex        =   119
                     Top             =   2130
                     Width           =   765
                  End
                  Begin VB.Label Label74 
                     AutoSize        =   -1  'True
                     Caption         =   "Doc No"
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
                     Left            =   165
                     TabIndex        =   118
                     Top             =   2130
                     Width           =   690
                  End
                  Begin VB.Label Label70 
                     AutoSize        =   -1  'True
                     Caption         =   "Mode of Payment"
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
                     Left            =   165
                     TabIndex        =   117
                     Top             =   870
                     Width           =   1575
                  End
                  Begin VB.Label Label64 
                     AutoSize        =   -1  'True
                     Caption         =   "Bank - A/C No."
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
                     Left            =   5205
                     TabIndex        =   116
                     Top             =   870
                     Width           =   1305
                  End
                  Begin VB.Label Label63 
                     AutoSize        =   -1  'True
                     Caption         =   "Disbursement Date"
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
                     Left            =   5205
                     TabIndex        =   115
                     Top             =   450
                     Width           =   1725
                  End
                  Begin VB.Label Label62 
                     AutoSize        =   -1  'True
                     Caption         =   "Amount Disbursed by FAG"
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
                     Left            =   165
                     TabIndex        =   114
                     Top             =   450
                     Width           =   2370
                  End
                  Begin VB.Label Label61 
                     AutoSize        =   -1  'True
                     Caption         =   "Disbursement Remark"
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
                     Left            =   165
                     TabIndex        =   113
                     Top             =   1515
                     Width           =   2010
                  End
                  Begin VB.Label Label60 
                     AutoSize        =   -1  'True
                     Caption         =   "FAG Payment Details :"
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
                     TabIndex        =   112
                     Top             =   120
                     Width           =   2340
                  End
                  Begin VB.Label Label49 
                     Caption         =   "Deduction Remark"
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
                     Left            =   7725
                     TabIndex        =   111
                     Top             =   1350
                     Width           =   1530
                  End
               End
               Begin VB.Frame Frame5 
                  Height          =   1290
                  Left            =   120
                  TabIndex        =   104
                  Top             =   120
                  Width           =   14145
                  Begin VB.TextBox TxtBankName 
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
                     Left            =   1380
                     Locked          =   -1  'True
                     MaxLength       =   50
                     TabIndex        =   21
                     Top             =   420
                     Width           =   5625
                  End
                  Begin VB.TextBox TxtBranchName 
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
                     Left            =   8700
                     Locked          =   -1  'True
                     MaxLength       =   50
                     TabIndex        =   22
                     Top             =   420
                     Width           =   5145
                  End
                  Begin VB.TextBox TxtACNo 
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
                     Left            =   1380
                     Locked          =   -1  'True
                     MaxLength       =   25
                     TabIndex        =   23
                     Top             =   825
                     Width           =   5625
                  End
                  Begin VB.TextBox TxtIFSCCode 
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
                     Left            =   8700
                     Locked          =   -1  'True
                     MaxLength       =   25
                     TabIndex        =   24
                     Top             =   825
                     Width           =   5145
                  End
                  Begin VB.Label Label69 
                     AutoSize        =   -1  'True
                     Caption         =   "Bank Name"
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
                     Left            =   180
                     TabIndex        =   109
                     Top             =   480
                     Width           =   1065
                  End
                  Begin VB.Label Label68 
                     AutoSize        =   -1  'True
                     Caption         =   "Branch Name"
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
                     Left            =   7380
                     TabIndex        =   108
                     Top             =   480
                     Width           =   1230
                  End
                  Begin VB.Label Label67 
                     AutoSize        =   -1  'True
                     Caption         =   "A/C No"
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
                     Left            =   180
                     TabIndex        =   107
                     Top             =   885
                     Width           =   645
                  End
                  Begin VB.Label Label66 
                     AutoSize        =   -1  'True
                     Caption         =   "IFSC Code"
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
                     Left            =   7380
                     TabIndex        =   106
                     Top             =   825
                     Width           =   975
                  End
                  Begin VB.Label Label65 
                     AutoSize        =   -1  'True
                     Caption         =   "Vendor Bank A/C  Details :"
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
                     TabIndex        =   105
                     Top             =   120
                     Width           =   2760
                  End
               End
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid4 
               Height          =   3375
               Left            =   -74880
               TabIndex        =   44
               Top             =   1800
               Width           =   14220
               _ExtentX        =   25083
               _ExtentY        =   5953
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
               Caption         =   "Updated By FAG/Power"
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
               TabIndex        =   86
               Top             =   1380
               Width           =   2175
            End
            Begin VB.Label Label21 
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
               TabIndex        =   85
               Top             =   540
               Width           =   1695
            End
            Begin VB.Label Label20 
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
               TabIndex        =   84
               Top             =   975
               Width           =   1695
            End
         End
         Begin MSComCtl2.DTPicker TxtClaimDate 
            Height          =   375
            Left            =   4440
            TabIndex        =   1
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Format          =   106037249
            CurrentDate     =   39884
         End
         Begin MSComCtl2.DTPicker TxtExpenseFrom 
            Height          =   375
            Left            =   7365
            TabIndex        =   2
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
            Format          =   106037249
            CurrentDate     =   39884
         End
         Begin MSComCtl2.DTPicker TxtExpenseTo 
            Height          =   375
            Left            =   10830
            TabIndex        =   3
            Top             =   240
            Width           =   1440
            _ExtentX        =   2540
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
            Format          =   106037249
            CurrentDate     =   39884
         End
         Begin MSComCtl2.DTPicker TxtApprovalDate 
            Height          =   375
            Left            =   7080
            TabIndex        =   9
            Top             =   1065
            Visible         =   0   'False
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   661
            _Version        =   393216
            Format          =   106037249
            CurrentDate     =   39884
         End
         Begin VB.TextBox TxtStatus 
            Appearance      =   0  'Flat
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
            Left            =   1485
            Locked          =   -1  'True
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   1077
            Width           =   3000
         End
         Begin VB.TextBox TxtForwardTo 
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
            Left            =   1485
            Locked          =   -1  'True
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   1488
            Width           =   7335
         End
         Begin VB.TextBox TxtBillingState 
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
            Left            =   10830
            Locked          =   -1  'True
            TabIndex        =   127
            TabStop         =   0   'False
            Top             =   1920
            Width           =   3720
         End
         Begin VB.Label Label10 
            Caption         =   "Vendor GST State on Bill"
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
            Left            =   9360
            TabIndex        =   126
            Top             =   1860
            Width           =   1365
         End
         Begin VB.Label LbClaimID 
            Caption         =   "Rent ID"
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
            TabIndex        =   100
            Top             =   277
            Width           =   900
         End
         Begin VB.Label Label41 
            Caption         =   "Claim Date"
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
            Left            =   3270
            TabIndex        =   99
            Top             =   285
            Width           =   1095
         End
         Begin VB.Label Label4 
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
            Height          =   270
            Left            =   120
            TabIndex        =   98
            Top             =   705
            Width           =   1305
         End
         Begin VB.Label Label2 
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
            Left            =   135
            TabIndex        =   97
            Top             =   2145
            Width           =   1080
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
            Left            =   12660
            TabIndex        =   96
            Top             =   285
            Visible         =   0   'False
            Width           =   630
         End
         Begin VB.Label Label14 
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
            Left            =   120
            TabIndex        =   95
            Top             =   1130
            Width           =   960
         End
         Begin VB.Label Label15 
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
            Height          =   270
            Left            =   9360
            TabIndex        =   94
            Top             =   720
            Width           =   1260
         End
         Begin VB.Label Label17 
            Caption         =   "Rent From "
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
            Left            =   6240
            TabIndex        =   93
            Top             =   285
            Width           =   1095
         End
         Begin VB.Label Label18 
            Caption         =   "Rent To"
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
            Left            =   9360
            TabIndex        =   92
            Top             =   285
            Width           =   1095
         End
         Begin VB.Label Label27 
            Caption         =   "Booked in NAV Date"
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
            Left            =   5160
            TabIndex        =   91
            Top             =   1125
            Visible         =   0   'False
            Width           =   1980
         End
         Begin VB.Label Label29 
            Caption         =   "Designation"
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
            Left            =   9360
            TabIndex        =   90
            Top             =   1530
            Width           =   1065
         End
         Begin VB.Label Label28 
            Caption         =   "Forward To"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   105
            TabIndex        =   89
            Top             =   1530
            Width           =   1230
         End
         Begin VB.Label Label3 
            Caption         =   "Tax"
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
            Left            =   9360
            TabIndex        =   88
            Top             =   1125
            Width           =   1410
         End
         Begin VB.Label Label26 
            Caption         =   "FAG  Remarks"
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
            Left            =   135
            TabIndex        =   87
            Top             =   2708
            Width           =   1350
         End
      End
      Begin VB.Frame ButtonFrm 
         Height          =   1125
         Left            =   120
         TabIndex        =   80
         Top             =   8550
         Width           =   14745
         Begin VB.CommandButton ExitCmd 
            Caption         =   "E&xit"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   9465
            TabIndex        =   56
            Top             =   600
            Width           =   2340
         End
         Begin VB.CommandButton SearchCmd 
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
            Height          =   400
            Left            =   9465
            TabIndex        =   51
            Top             =   195
            Width           =   2340
         End
         Begin VB.CommandButton LastCmd 
            Caption         =   "&Last"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   7125
            TabIndex        =   55
            Top             =   600
            Width           =   2340
         End
         Begin VB.CommandButton DeleteCmd 
            Caption         =   "&Delete"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   7125
            TabIndex        =   50
            Top             =   195
            Width           =   2340
         End
         Begin VB.CommandButton FirstCmd 
            Caption         =   "&First"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   4785
            TabIndex        =   54
            Top             =   600
            Width           =   2340
         End
         Begin VB.CommandButton EditCmd 
            Caption         =   "&Edit"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   4785
            TabIndex        =   49
            Top             =   195
            Width           =   2340
         End
         Begin VB.CommandButton PreviousCmd 
            Caption         =   "&Previous"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   2445
            TabIndex        =   53
            Top             =   600
            Width           =   2340
         End
         Begin VB.CommandButton SaveCmd 
            Caption         =   "Save"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   2445
            TabIndex        =   48
            Top             =   195
            Width           =   2340
         End
         Begin VB.CommandButton NextCmd 
            Caption         =   "&Next"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   105
            TabIndex        =   52
            Top             =   600
            Width           =   2340
         End
         Begin VB.CommandButton AddCmd 
            Caption         =   "&Add New"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   105
            TabIndex        =   47
            Top             =   195
            Width           =   2340
         End
         Begin VB.CommandButton CmdPrint 
            Caption         =   "Print"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Left            =   11805
            TabIndex        =   57
            Top             =   195
            Width           =   2700
         End
      End
   End
End
Attribute VB_Name = "FrmTxn_GodownApproval"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg, mExpenseTypeID, mLessorID, mSLessorID, mStatus, mSStatus, mLessorEmailID As Variant
Dim mLocationID, mSLocationID As Double
Dim mBatchID As Double
Dim mForwardToID, mForwardToEmailID, mSForwardToID, mEmployeeEmailID As Variant
Dim oSMTP As SMTPSession
Dim mSelRow As Variant
Dim mSelCol As Variant
Dim mFAGBankID, mPBatchID As Variant

Dim mBillingStateID As Double   ''-- Ullhas Patil 26 Oct 2017

Private Sub AddCmd_Click()
    ''--- Ullhas Patil 25 Oct 2017
    CmbBillingState.Visible = True: CmbBillingState = ""
    TxtBillingState.Visible = False: TxtBillingState = ""
    ''-----------------------------
End Sub

''-- Ullhas Patil 26 Oct 2017
Private Sub CmbBillingState_GotFocus()
tmp = CmbBillingState.Text
Call TableMst_StateIndia
CmbBillingState.Clear
If RsMst_StateIndia.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_StateIndia.RecordCount
    CmbBillingState.AddItem RsMst_StateIndia!StateIndiaName
    RsMst_StateIndia.MoveNext
Next
CmbBillingState.Text = tmp
End Sub
Private Sub CmbBillingState_LostFocus()
If CmbBillingState.Text = "" Then
   Exit Sub
End If
RsMst_StateIndia.MoveFirst
RsMst_StateIndia.Find "StateIndiaName = '" & CmbBillingState.Text & "'"
If RsMst_StateIndia.EOF Then
   MsgBox "Invalid selection "
   CmbBillingState.SetFocus
   Exit Sub
End If
mBillingStateID = RsMst_StateIndia!StateIndiaID
End Sub
''-------------------------------------


Private Sub ChkBatchNo_Click()
TxtSBatchID.Locked = True
TxtSBatchID = ""
If ChkBatchNo.Value = 1 Then
   TxtSBatchID.Locked = False
End If

End Sub

Private Sub CmbFAGBank_GotFocus()
tmp = CmbFAGBank.Text
Call TableMst_FAGBank("")
CmbFAGBank.Clear
If RsMst_FAGBank.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_FAGBank.RecordCount
    CmbFAGBank.AddItem RsMst_FAGBank!BankName & "~" & RsMst_FAGBank!AccountNo
    RsMst_FAGBank.MoveNext
Next
CmbFAGBank.Text = tmp
End Sub
Private Sub CmbFAGBank_LostFocus()
Dim strSplitString() As String
Dim AcNo_, BankName_ As String

If CmbFAGBank.Text = "" Then
   Exit Sub
End If
strSplitString = Split(CmbFAGBank, "~")
If UBound(strSplitString) <> 1 Then
   MsgBox "Invalid selection !!!!! "
   CmbFAGBank.SetFocus
   Exit Sub
End If
BankName_ = Trim(strSplitString(0))
AcNo_ = Trim(strSplitString(1))

RsMst_FAGBank.MoveFirst
RsMst_FAGBank.Find "AccountNo='" & AcNo_ & "'"
If RsMst_FAGBank.EOF Then
   MsgBox "Invalid selection "
   CmbFAGBank.SetFocus
   Exit Sub
End If
mFAGBankID = RsMst_FAGBank!FAGBankID
End Sub

Private Sub CmbModeOfPay_LostFocus()
If CmbModeOfPay.Text = "" Then Exit Sub
If LCase(CmbModeOfPay.Text) <> "rtgs" And LCase(CmbModeOfPay.Text) <> "neft" And LCase(CmbModeOfPay.Text) <> "bank transfer" And LCase(CmbModeOfPay.Text) <> "cheque" Then
   MsgBox "Invalid Selection!!!"
   CmbModeOfPay.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmbSForwardTo_GotFocus()
tmp = CmbSForwardTo.Text
If Gen_WcEmployeeID <> "" Then
   Call TableMst_Employee("Where " & Gen_WcEmployeeID)
Else
   Call TableMst_Employee '("Where Flag = 'Y'")
End If
CmbSForwardTo.Clear
If RsMst_Employee.RecordCount = 0 Then
   MsgBox "No Employee Found!!!"
   CmbSForwardTo.SetFocus
   'RsMst_Employee.MoveFirst
   Exit Sub
End If
For I = 1 To RsMst_Employee.RecordCount
    CmbSForwardTo.AddItem RsMst_Employee!EmployeeName & "~" & RsMst_Employee!EmployeeNo
    RsMst_Employee.MoveNext
Next
CmbSForwardTo.Text = tmp

End Sub


Private Sub CmbSLessor_GotFocus()
tmp = CmbSLessor.Text
mSLessorID = ""
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
   mSStatus = Null
   Exit Sub
End If
If LCase(CmbSStatus.Text) = "approved" Then
   mSStatus = "A"
ElseIf LCase(CmbSStatus.Text) = "pending" Then
   mSStatus = "P"
ElseIf LCase(CmbSStatus.Text) = "rejected" Then
   mSStatus = "R"
ElseIf LCase(CmbSStatus.Text) = "settled" Then
   mSStatus = "S"
ElseIf LCase(CmbSStatus.Text) = "frozen" Then
   mSStatus = "F"
Else
   MsgBox "Invalid selection!!!"
   CmbSStatus.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmbStatus_LostFocus()

If CmbStatus.Text = "" Then
   TxtApprovalRemarks.Locked = True
   TxtApprovalDate.Visible = False
   Label27.Visible = False
   Exit Sub
End If

If mStatus = "A" Then
   If mBatchID = 0 Then
      If LCase(CmbStatus.Text) = "pending" Then
         MsgBox "This claim is approved already . You can't mark it as " & CmbStatus.Text & " status. !!!"
         CmbStatus.SetFocus
         Exit Sub
      End If
   ElseIf mBatchID <> 0 Then
      If LCase(CmbStatus.Text) = "pending" Or LCase(CmbStatus.Text) = "rejected" Or LCase(CmbStatus.Text) = "frozen" Then
         MsgBox "This claim is approved already . You can't mark it as " & CmbStatus.Text & " status. !!!"
         CmbStatus.SetFocus
         Exit Sub
      End If
   End If
End If

If LCase(CmbStatus.Text) = "approved" Then
   If LCase(Gen_UserRole) <> "fag" And LCase(Gen_UserRole) <> "fagadmin" And LCase(Gen_UserRole) <> "power" Then
      MsgBox "Only FAG can Approved Status !!!"
      'CmbStatus.Text = ""
      CmbStatus.SetFocus
      Exit Sub
   End If
   mStatus = "A"
   TxtApprovalRemarks.Locked = False
   TxtApprovalDate.Visible = True
   Label27.Visible = True
   TxtApprovalDate.Enabled = True
   TxtApprovalDate.Value = Date
   TxtFAGRemark = ""
ElseIf LCase(CmbStatus.Text) = "settled" Then
   mStatus = "S"
   Frame3.Enabled = True
   FrmFAG.Enabled = True
   CmbFAGBank.Visible = True
   TxtFAGBank.Visible = False
   CmbFAGBank.Text = TxtFAGBank
   CmbModeOfPay.Text = TxtModeOfPayment.Text
   CmbModeOfPay.Visible = True
   TxtModeOfPayment.Visible = False
   TxtFAGDisbAmt.Locked = False
   TxtFAGDisbAmt.Text = Val(TxtClaimAmount)
   TxtDeductionRemark = "TDS@"
   TxtFAGRemark = Trim(TxtApprovalRemarks)
   Call TableMst_FAGBank(" Where FAGBankID in (" & Gen_FAGBankIDForPayment & ") ")
   If RsMst_FAGBank.RecordCount > 0 Then
      CmbFAGBank.Text = RsMst_FAGBank!BankName & "~" & RsMst_FAGBank!AccountNo
      mFAGBankID = RsMst_FAGBank!FAGBankID
   End If
ElseIf LCase(CmbStatus.Text) = "pending" Then
   mStatus = "P"
   TxtApprovalRemarks.Locked = True
   TxtApprovalDate.Visible = False
   Label27.Visible = False
   TxtFAGRemark = ""
ElseIf LCase(CmbStatus.Text) = "rejected" Then
   mStatus = "R"
   TxtApprovalRemarks.Locked = False
   TxtApprovalDate.Visible = True
   Label27.Visible = True
   TxtApprovalDate.Enabled = True
   TxtApprovalDate.Value = Date
   TxtFAGRemark = ""
ElseIf LCase(CmbStatus.Text) = "frozen" Then
   mStatus = "F"
   TxtApprovalRemarks.Locked = False
   TxtApprovalDate.Visible = True
   Label27.Visible = True
   TxtApprovalDate.Enabled = True
   TxtApprovalDate.Value = Date
   TxtFAGRemark = ""
Else
   MsgBox "Invalid selection!!!"
   CmbStatus.SetFocus
   TxtApprovalRemarks.Locked = True
   TxtApprovalDate.Visible = False
   Label27.Visible = False
   Exit Sub
End If


If LCase(Gen_UserRole) <> "fag" And LCase(Gen_UserRole) <> "fagadmin" And LCase(Gen_UserRole) <> "power" Then
   If mStatus = "A" Or mStatus = "S" Then
      MsgBox "Please forward this claim to FAG user, Only FAG user have access to approve or settle claim '!!!"
      If FrmFAG.Enabled = True Then
         TxtFAGDisbAmt.Text = ""
         FrmFAG.Enabled = False
      End If
      CmbStatus.SetFocus
      Exit Sub
   End If
Else
   If mStatus = "S" Then
      If mBatchID = 0 Then
         MsgBox "Claim not booked in NAVISION, please export batch from Store-Man and upload in NAVISION!!!"
         If FrmFAG.Enabled = True Then
            TxtFAGDisbAmt.Text = ""
            FrmFAG.Enabled = False
         End If
         CmbStatus.SetFocus
         Exit Sub
      End If
   End If
End If
End Sub

Private Sub CmdForwardTo_Click()
FrmMst_Employee.Show
End Sub

Private Sub CmdGo_Click()
Dim strSplitString() As String
Dim EmpNo_, EmpName_ As String
Dim wc As Variant

Call Grid_Head1

If SFrom.Value > STo.Value Then
   MsgBox "'From Date' must be less than 'To date'!!!"
   SFrom.SetFocus
   Exit Sub
End If
Gen_DatFormat = "yyyy/mm/dd"

wc = ""
RsTxn_GodownRent.Filter = ""  'VC TGR

tmp = "Select TGR.ClaimID,TGR.ClaimDate,TGR.ExpenseFrom,TGR.ExpenseTo, "
tmp = tmp & "  (Select Max(LocationName) from  Txn_GodownRentDetail TG"
tmp = tmp & "  Inner join Mst_Godown MG on TG.GodownID=MG.GodownID"
tmp = tmp & "  Inner join Mst_CMP MC on MG.CMPID=MC.CMPID"
tmp = tmp & "  Inner join Mst_Location ML on MC.LocationID=ML.LocationID"
tmp = tmp & "  Where  TG.ClaimID= TGR.ClaimID) as LocationName,"
tmp = tmp & "  ML.LessorName,TGR.Amount,EMP.EmployeeName +'~'+ EMP.EmployeeNo as ForwardedTo ,TGR.Flag, "
tmp = tmp & "  TGR.ApprovalDate,TGR.ApprovalRemarks, "
tmp = tmp & "  TGR.Remarks,TGR.BatchID,TGR.G_UID,TGR.BankName,TGR.ChequeNo,TGR.PaidAmount,TGR.TDS,TGR.PaidDate, "
tmp = tmp & "  TGR.DisbuDate,TGR.DisbuRemark,TGR.DisbuAmount,TGR.ModeOfPay,MFB.BankName+'~'+MFB.AccountNo,TGR.PBatchID,TGR.DocNo,TGR.EntryNo,TGR.FAGPayExportedDate,TGR.DeductionRemark,TGR.LetterPrintFlag,MSII.StateIndiaName[BillingStateName]  "
tmp = tmp & "  From Txn_GodownRent TGR "
tmp = tmp & "  Inner Join Mst_Lessor ML On TGR.LessorID = ML.LessorID "
tmp = tmp & "  Left Outer Join Mst_Employee EMP On TGR.ForwardToID = EMP.EmployeeID "
tmp = tmp & "  Left Join Mst_FAGBank MFB On TGR.FAGBankID = MFB.FAGBankID "

tmp = tmp & " Left join Mst_StateIndia MSII On TGR.BillingStateIndiaID = MSII.StateIndiaID "     ''--- Ullhas Patil 25 Oct 2017

If Gen_DBType = "MDB" Then
   tmp = tmp & " Where ClaimDate Between #" & Format(SFrom, Gen_DatFormat) & "# And #" & Format(STo, Gen_DatFormat) & "#) "
   wc = " ClaimDate >= #" & Format(SFrom, Gen_DatFormat) & "# And ClaimDate <= #" & Format(STo, Gen_DatFormat) & "#) "
Else
   tmp = tmp & " Where ClaimDate Between '" & Format(SFrom, Gen_DatFormat) & "' And '" & Format(STo, Gen_DatFormat) & "' "
   wc = " ClaimDate >= '" & Format(SFrom, Gen_DatFormat) & "' And ClaimDate <= '" & Format(STo, Gen_DatFormat) & "' "
End If


If CmbSLocation.Text <> "" Then
   tmp = tmp & " And ClaimID in (Select Distinct ClaimID from  Txn_GodownRentDetail TG"
   tmp = tmp & " Inner join Mst_Godown MG on TG.GodownID=MG.GodownID"
   tmp = tmp & " Inner join Mst_CMP MC on MG.CMPID=MC.CMPID"
   tmp = tmp & " Where  MC.LocationID=" & mSLocationID & ")"
End If

If CmbSLessor.Text <> "" Then
   tmp = tmp & " AND ML.LessorName = '" & CmbSLessor.Text & "' "
   wc = wc & " And LessorID = " & mSLessorID & ""
End If

If CmbSStatus.Text <> "" Then
   tmp = tmp & " AND TGR.Flag = '" & mSStatus & "' "
   wc = wc & " And Flag = '" & mSStatus & "'"
End If

If CmbSForwardTo.Text <> "" Then
   strSplitString = Split(CmbSForwardTo, "~")
   EmpName_ = Trim(strSplitString(0))
   EmpNo_ = Trim(strSplitString(1))
 
   tmp = tmp & " AND EMP.EmployeeNo = '" & EmpNo_ & "' "
   wc = wc & " And ForwardToID = " & mSForwardToID & ""
'Else
'   If Gen_WcEmployeeID <> "" Then
'      tmp = tmp & " And EMP." & Gen_WcEmployeeID & ""
''      Wc = Wc & " And ForwardToID = " & mSForwardToID & ""
'   End If
End If

If TxtSClaimID <> "" Then
   tmp = tmp & " AND TGR.ClaimID = " & TxtSClaimID & " "
   wc = wc & " And ClaimID = " & TxtSClaimID & ""
End If

If ChkBatchNo.Value = 0 Then
   tmp = tmp & " AND TGR.BatchID = 0 "
   wc = wc & " And BatchID = 0 "
End If

If TxtSBatchID <> "" And ChkBatchNo.Value = 1 Then
   tmp = tmp & " AND TGR.BatchID = " & TxtSBatchID & " "
   wc = wc & " And BatchID = " & TxtSBatchID & ""
End If

Call TmpTable

'Here Filter is used to show records in indata which are searched
RsTxn_GodownRent.Filter = "" & wc & ""

For iu = 1 To TmpRs.RecordCount
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
    For C = 0 To MSHFlexGrid1.Cols - 1
        If InStr(TmpRs.Fields(C), "/") > 0 Then
           If IsDate(TmpRs.Fields(C)) Then
              MSHFlexGrid1.TextMatrix(iu, C) = Format(TmpRs.Fields(C), "dd-MMM-yyyy")
           Else
              MSHFlexGrid1.TextMatrix(iu, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
           End If
        Else
           MSHFlexGrid1.TextMatrix(iu, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
        End If
''        If c = 5 Then
''           MSHFlexGrid1.TextMatrix(iu, c) = Format(IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c)), "########0.00")
''        End If
    Next

    If LCase(MSHFlexGrid1.TextMatrix(iu, 8)) = "a" Then
       MSHFlexGrid1.TextMatrix(iu, 8) = "Approved"
    ElseIf LCase(MSHFlexGrid1.TextMatrix(iu, 8)) = "p" Then
       MSHFlexGrid1.TextMatrix(iu, 8) = "Pending"
    ElseIf LCase(MSHFlexGrid1.TextMatrix(iu, 8)) = "r" Then
       MSHFlexGrid1.TextMatrix(iu, 8) = "Rejected"
    ElseIf LCase(MSHFlexGrid1.TextMatrix(iu, 8)) = "s" Then
       MSHFlexGrid1.TextMatrix(iu, 8) = "Settled"
    ElseIf LCase(MSHFlexGrid1.TextMatrix(iu, 8)) = "f" Then
       MSHFlexGrid1.TextMatrix(iu, 8) = "Frozen"
    End If
    
    If LCase(MSHFlexGrid1.TextMatrix(iu, 29)) = "n" Then
       MSHFlexGrid1.TextMatrix(iu, 29) = "No"
    ElseIf LCase(MSHFlexGrid1.TextMatrix(iu, 29)) = "y" Then
       MSHFlexGrid1.TextMatrix(iu, 29) = "Yes"
    End If
    
    TmpRs.MoveNext
Next


Label22.Refresh
Label22.Caption = TmpRs.RecordCount

End Sub

Private Sub CmdPrint_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "GodownRent_" & TxtClaimID & ".xls")
Call Xls_Detail_Rpt
End Sub
Public Sub Xls_Detail_Rpt()

Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim oWS As Excel.Worksheet
Dim mRow, mRow1 As Integer
Dim mCol As Integer
Dim C As Variant
Set oXL = New Excel.Application
Dim Pat As String
Dim mAmt, mSTAmt, mTotAmt As Double
Pat = App.Path

Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "GodownRent_" & TxtClaimID & ".xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows

oWS.Shapes.AddPicture(App.Path & "\logo.gif", msoTrue, msoTrue, 10, 10, 40, 40).Select

oWS.Range("d:d").ColumnWidth = 12.86
oWS.Range("e:e").ColumnWidth = 18.86

tmp = "B" & 1 & ":f" & 1
oWS.Range(tmp).Merge
oWS.Cells(1, 2) = "National Collateral Management Services Limited"
oWS.Cells(1, 2).HorizontalAlignment = -4108
oWS.Range(tmp).Font.Bold = True

tmp = "B" & 2 & ":f" & 2
oWS.Range(tmp).Merge
oWS.Cells(2, 2) = "Lodha Supremus, Unit No 505 to 509 , 5th Floor,"
oWS.Cells(2, 2).HorizontalAlignment = -4108

tmp = "B" & 3 & ":f" & 3
oWS.Range(tmp).Merge
oWS.Cells(3, 2) = "Kanjurmarg (East), Mumbai-400042"
oWS.Cells(3, 2).HorizontalAlignment = -4108

tmp = "B" & 4 & ":f" & 4
oWS.Range(tmp).Merge
oWS.Cells(4, 2) = "Tel.No-022 40419191, Fax. No-022 40419193"
oWS.Cells(4, 2).HorizontalAlignment = -4108

tmp = "a" & 7 & ":f" & 7
oWS.Range(tmp).Merge
oWS.Cells(7, 1) = ""
oWS.Cells(8, 1) = "To :"

tmp = "a" & 9 & ":f" & 9
oWS.Range(tmp).Merge
oWS.Cells(9, 1) = "      " & TxtLessorID.Text
oWS.Cells(9, 1).Font.Bold = True
tmp = "a" & 10 & ":d" & 10
oWS.Range(tmp).Merge
oWS.Cells(10, 1) = "" & GetLessorAddress(mLessorID) '""  'Address
oWS.Range(tmp).WrapText = True
oWS.Cells(10, 1).RowHeight = 45
oWS.Cells(11, 1) = "Dear Sir,"
tmp = "a" & 13 & ":f" & 14
oWS.Range(tmp).Merge
oWS.Range(tmp).WrapText = True
oWS.Cells(13, 1) = "Please find enclosed Rent paid details for the month " & _
                   " of " & MonthName(Month(TxtExpenseFrom.Value)) & " " & Year(TxtExpenseFrom.Value) & ". The details are given below."

oWS.Cells(16, 1) = "Agreement ID"
oWS.Cells(16, 2) = "Location"
oWS.Cells(16, 3) = "Godown No."
oWS.Cells(16, 4) = "Area in SqFt"
oWS.Cells(16, 5) = "Agree. Expiry Date"
oWS.Cells(16, 6) = "Gross Amount"
tmp = "a" & 16 & ":f" & 16
oWS.Range(tmp).Font.Bold = True

tmp = "Select TLA.LAID,ML.LocationName,MG.GodownNo,MG.Area,TLA.ExpiryDate,TGRD.Amount"
tmp = tmp & " from Txn_GodownRent TGR"
tmp = tmp & " Inner Join Txn_GodownRentDetail TGRD On TGR.ClaimID = TGRD.ClaimID"
tmp = tmp & " Inner Join Mst_Godown MG On MG.GodownID=TGRD.GodownID"
tmp = tmp & " Inner Join Mst_CMP MC On MC.CMPID = MG.CMPID"
tmp = tmp & " inner join Mst_Location ML on MC.LocationID=ML.LocationID"
tmp = tmp & " Inner Join Txn_LessorAgreement TLA on TGRD.LAID=TLA.LAID"
tmp = tmp & " Where TGR.ClaimID=" & Val(TxtClaimID)

Call TmpTable

tmp = "Select PANNo,RTGSNo From Mst_Lessor Where LessorID=" & mLessorID
Call Tmp1Table

mRow = 17

mAmt = 0
mSTAmt = 0
mTotAmt = 0
If TmpRs.RecordCount > 0 Then
   For I = 1 To TmpRs.RecordCount
       C = 1
       For Col = 0 To TmpRs.Fields.Count - 1
           If InStr(TmpRs.Fields(Col), "/") > 0 Then
              If IsDate(TmpRs.Fields(Col)) Then
                 oWS.Cells(mRow, C) = Format(TmpRs.Fields(Col), "dd-MMM-yyyy")
              Else
                 oWS.Cells(mRow, C) = TmpRs.Fields(Col)
              End If
           Else
              If Col = 3 Then
                 oWS.Cells(mRow, C) = TmpRs.Fields(Col)
                 mAmt = mAmt + Val(TmpRs.Fields(Col))
              ElseIf Col = 4 Then
                 oWS.Cells(mRow, C) = TmpRs.Fields(Col)
              ElseIf Col = 5 Then
                 oWS.Cells(mRow, C) = TmpRs.Fields(Col)
                 mTotAmt = mTotAmt + Val(TmpRs.Fields(Col))
              Else
                 oWS.Cells(mRow, C) = TmpRs.Fields(Col)
              End If
           End If
           C = C + 1
       Next
       TmpRs.MoveNext
       mRow = mRow + 1
   Next
End If
tmp = "a" & 16 & ":f" & mRow
oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)

tmp = "a" & mRow & ":c" & mRow
oWS.Range(tmp).Merge
oWS.Range(tmp).Font.Bold = True
oWS.Cells(mRow, 1) = "Total"
oWS.Cells(mRow, 4) = Format(mAmt, "#####0.00")
oWS.Cells(mRow, 6) = Format(mTotAmt, "#####0.00")

tmp = "d" & mRow & ":f" & mRow
oWS.Range(tmp).Font.Bold = True

tmp = "b" & mRow + 2 & ":b" & mRow + 9
oWS.Range(tmp).Font.Bold = True

tmp = "b" & mRow + 2 & ":e" & mRow + 9
oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)

tmp = "b" & mRow + 2 & ":c" & mRow + 2
oWS.Range(tmp).Merge
oWS.Range(tmp).WrapText = True
oWS.Cells(mRow + 2, 2) = "Bank Name"

tmp = "d" & mRow + 2 & ":e" & mRow + 2
oWS.Range(tmp).Merge
oWS.Range(tmp).WrapText = True
oWS.Cells(mRow + 2, 4) = TxtBankID.Text

tmp = "b" & mRow + 3 & ":c" & mRow + 3
oWS.Range(tmp).Merge
oWS.Range(tmp).WrapText = True
oWS.Cells(mRow + 3, 2) = "Paid Amount"

tmp = "d" & mRow + 3 & ":e" & mRow + 3
oWS.Range(tmp).Merge
oWS.Range(tmp).WrapText = True
oWS.Cells(mRow + 3, 4) = Format(Val(TxtPaidAmt.Text), "#####0.00")


tmp = "b" & mRow + 4 & ":c" & mRow + 4
oWS.Range(tmp).Merge
oWS.Range(tmp).WrapText = True
oWS.Cells(mRow + 4, 2) = "TDS"

tmp = "d" & mRow + 4 & ":e" & mRow + 4
oWS.Range(tmp).Merge
oWS.Range(tmp).WrapText = True
oWS.Cells(mRow + 4, 4) = Format(Val(TxtTDS.Text), "#####0.00")

tmp = "b" & mRow + 5 & ":c" & mRow + 5
oWS.Range(tmp).Merge
oWS.Range(tmp).WrapText = True
oWS.Cells(mRow + 5, 2) = "Net Amount"

tmp = "d" & mRow + 5 & ":e" & mRow + 5
oWS.Range(tmp).Merge
oWS.Range(tmp).WrapText = True
oWS.Cells(mRow + 5, 4) = Format(Val(TxtPaidAmt) + Val(TxtTDS), "#####0.00")


tmp = "b" & mRow + 6 & ":c" & mRow + 6
oWS.Range(tmp).Merge
oWS.Range(tmp).WrapText = True
oWS.Cells(mRow + 6, 2) = "Cheque No."

tmp = "d" & mRow + 6 & ":e" & mRow + 6
oWS.Range(tmp).Merge
oWS.Range(tmp).WrapText = True
oWS.Cells(mRow + 6, 4) = "'" & TxtChequeNo.Text

tmp = "b" & mRow + 7 & ":c" & mRow + 7
oWS.Range(tmp).Merge
oWS.Range(tmp).WrapText = True
oWS.Cells(mRow + 7, 2) = "Paid Date"

tmp = "d" & mRow + 7 & ":e" & mRow + 7
oWS.Range(tmp).Merge
oWS.Range(tmp).WrapText = True
oWS.Cells(mRow + 7, 4) = "'" & TxtPaidDate.Text

tmp = "b" & mRow + 8 & ":c" & mRow + 8
oWS.Range(tmp).Merge
oWS.Range(tmp).WrapText = True
oWS.Cells(mRow + 8, 2) = "RTGS Details"


tmp = "d" & mRow + 8 & ":e" & mRow + 8
oWS.Range(tmp).Merge
oWS.Range(tmp).WrapText = True
If LCase(TxtChequeNo) = "trf" Then
   oWS.Cells(mRow + 8, 4) = "'" & IIf(IsNull(TmpRs1!RTGSNo), "", TmpRs1!RTGSNo)
Else
   oWS.Cells(mRow + 8, 4) = ""
End If
tmp = "b" & mRow + 9 & ":c" & mRow + 9
oWS.Range(tmp).Merge
oWS.Range(tmp).WrapText = True
oWS.Cells(mRow + 9, 2) = "PAN Card No."

tmp = "d" & mRow + 9 & ":e" & mRow + 9
oWS.Range(tmp).Merge
oWS.Range(tmp).WrapText = True
oWS.Cells(mRow + 9, 4) = "'" & IIf(IsNull(TmpRs1!PANNO), "", TmpRs1!PANNO)

tmp = "a" & mRow + 12 & ":f" & mRow + 12
oWS.Range(tmp).Merge
oWS.Cells(mRow + 12, 1) = "For National Collateral Management Services Ltd."

tmp = "a" & mRow + 15 & ":f" & mRow + 15
oWS.Range(tmp).Merge
oWS.Cells(mRow + 15, 1) = "(Authorised Signatory)"

If Gen_ExcelPassword <> "" Then
   oWS.Protect Gen_ExcelPassword, , , , , , True, True
End If

oWB.Save
MsgBox ("Done!!")
oXL.Visible = True

End Sub

Private Sub EditCmd_Click()


'If LCase(TxtStatus) = "approved" Or LCase(TxtStatus) = "rejected" Then
'   MsgBox "This Claim is " & TxtStatus & ". You can't edit this claim any more!!!"
'   Exit Sub
'End If

If Format(Gen_BlockingDate, "yyyy-mm") >= Format(TxtExpenseFrom.Value, "yyyy-mm") Then
   MsgBox "Expense month is previous month from " & Format(Gen_BlockingDate, "MMM-yyyy") & ", you can not edit!!! "
   Exit Sub
End If
If LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "fag" And LCase(Gen_UserRole) <> "fagadmin" Then
   If LCase(TxtStatus) = "settled" Or LCase(TxtStatus) = "approved" Or LCase(TxtStatus) = "rejected" Or LCase(TxtStatus) = "frozen" Then
      MsgBox "This Claim is " & TxtStatus & ". You can't edit this claim any more!!!"
      Exit Sub
   End If
Else
   If LCase(TxtStatus) = "rejected" Or LCase(TxtStatus) = "settled" Or LCase(TxtStatus) = "frozen" Then
      MsgBox "This Claim is " & TxtStatus & ". You can't edit this claim any more!!!"
      Exit Sub
   End If
End If

If mPBatchID <> 0 Then
   MsgBox "Payment batch is generated against this Claim. You can't edit this claim any more!!!"
   Exit Sub
End If

Edit_Flg = "E"
Call GButtonLock(Me, False)
'TxtClaimDate.Enabled = True
CmdPrint.Enabled = False
TxtForwardTo.Visible = False
CmbForwardTo.Visible = True
CmbForwardTo.Text = TxtForwardTo

'If LCase(Gen_UserRole) = "fag" Then ''Or LCase(Gen_UserRole) = "power"
TxtStatus.Visible = False
CmbStatus.Visible = True
CmbStatus.Text = TxtStatus

'End If

If mStatus = "p" Then
   TxtApprovalDate.Value = Date
End If

If LCase(mStatus) <> "s" And mBatchID = 0 Then
   TxtRemark.Locked = False
   Frame2.Enabled = True
   TxtApprovalRemarks.Locked = False
   CmdForwardTo.Visible = True
End If

''--- Ullhas Patil 26 Oct 2017
CmbBillingState.Visible = True
TxtBillingState.Visible = False
CmbBillingState = TxtBillingState
''-----------------------------

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
   Frame0.Enabled = False
   Frame1.Enabled = False
   Frame2.Enabled = False
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If
Frame0.Enabled = True
Frame1.Enabled = True
Frame2.Enabled = True
Call Grid_Head1

SFrom.Value = CDate(From_GenDate) - 900
STo.Value = Date 'To_GenDate
CmbSLessor.Text = ""
CmbSStatus.Text = ""
CmbSForwardTo.Text = ""
CmbSLocation.Text = ""
ChkBatchNo.Value = 0
TxtSBatchID.Locked = True
TxtSClaimID = ""
Label22.Caption = ""

tmp = "Select EmployeeID,EmployeeName +'~'+ EmployeeNo as Employee,DesignationID,EmployeeEmailID From Mst_Employee Where EmployeeNo = '" & Gen_UserLoginID & "'"
Call TmpTable

If TmpRs.RecordCount = 1 Then
   CmbSForwardTo.Text = TmpRs!Employee
   mSForwardToID = TmpRs!EmployeeID
Else
   CmbSForwardTo.Text = ""
   mSForwardToID = Null
End If

Call Grid_Head
mEID = Replace(Gen_WcEmployeeID, "EmployeeID", "CMPMID")


tmp = "Select Distinct(EMP.EmployeeID) " ', EMP.EmployeeNO, EMP.EmployeeName"
tmp = tmp & " From Mst_Employee EMP"
tmp = tmp & " Left Outer Join Mst_Employee EMP1 On EMP.DirectReportingTo = EMP1.EmployeeID"
tmp = tmp & " Left Outer Join Mst_Employee EMP2 On EMP1.DirectReportingTo = EMP2.EmployeeID"
tmp = tmp & " Left Outer Join Mst_Employee EMP3 On EMP2.DirectReportingTo = EMP3.EmployeeID"
tmp = tmp & " Left Outer Join Mst_Employee EMP4 On EMP3.DirectReportingTo = EMP4.EmployeeID"
tmp = tmp & " Left Outer Join Mst_Employee EMP5 On EMP4.DirectReportingTo = EMP5.EmployeeID"
tmp = tmp & " Left Outer Join Mst_Employee EMP6 On EMP5.DirectReportingTo = EMP6.EmployeeID"
tmp = tmp & " Left Outer Join Mst_Employee EMP7 On EMP6.DirectReportingTo = EMP7.EmployeeID"
tmp = tmp & " Left Outer Join Mst_Employee EMP8 On EMP5.DirectReportingTo = EMP8.EmployeeID"
tmp = tmp & " Left Outer Join Mst_Employee EMP9 On EMP8.DirectReportingTo = EMP9.EmployeeID"

tmp = tmp & " Where (EMP.EmployeeID = " & Gen_EmployeeID & " Or EMP1.EmployeeID = " & Gen_EmployeeID & " Or EMP2.EmployeeID =" & Gen_EmployeeID & " Or"
tmp = tmp & " EMP3.EmployeeID = " & Gen_EmployeeID & " Or EMP4.EmployeeID = " & Gen_EmployeeID & " or EMP5.EmployeeID=" & Gen_EmployeeID & " Or"
tmp = tmp & " EMP6.EmployeeID = " & Gen_EmployeeID & " Or EMP7.EmployeeID = " & Gen_EmployeeID & " Or EMP8.EmployeeID = " & Gen_EmployeeID & " Or"
tmp = tmp & " EMP9.EmployeeID = " & Gen_EmployeeID & " )"
tmp = tmp & " and Emp.Flag='N' "

Call Tmp4Table

If TmpRs4.RecordCount > 0 Then
   mDeactiveEID = "CMPMID in ("
   For I = 1 To TmpRs4.RecordCount
       If I = 1 Then
          mDeactiveEID = mDeactiveEID & TmpRs4.Fields(0) & " "
       Else
          mDeactiveEID = mDeactiveEID & "," & TmpRs4.Fields(0) & " "
       End If
       TmpRs4.MoveNext
   Next
   mDeactiveEID = mDeactiveEID & " )"
End If

If LCase(Gen_UserRole) = "power" Or LCase(Gen_UserRole) = "fag" Or LCase(Gen_UserRole) = "fagadmin" Or Gen_DesignationID = 1 Then
   Call TableTxn_GodownRent
Else
   If mDeactiveEID <> "" Then
      Call TableTxn_GodownRent(" Where " & mEID & " Or " & mDeactiveEID & "")
   Else
      Call TableTxn_GodownRent(" Where " & mEID & " ")
   End If
End If


'If LCase(Gen_UserRole) = "power" Or LCase(Gen_UserRole) = "fag" Or Gen_DesignationID = 1 Then
'   Call TableTxn_GodownRent
'Else
'   Call TableTxn_GodownRent(" Where " & mEID & "")
'End If

If RsTxn_GodownRent.RecordCount = 0 Then
   If UserType = "Auditor" Then
      Call GButtonLock(Me, False)
      SaveCmd.Enabled = False
      SearchCmd.Enabled = False
      Exit Sub
   End If
   Frame0.Enabled = False
   Frame1.Enabled = False
   
   'ButtonFrm.Enabled = False
   Call GButtonLockAD(Me)
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

Private Sub MSHFlexGrid1_Click()
Dim I As Variant
I = MSHFlexGrid1.RowSel
If I <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(I, 0) = "" Then Exit Sub
RsTxn_GodownRent.MoveFirst
RsTxn_GodownRent.Find "G_UID = '" & MSHFlexGrid1.TextMatrix(I, 13) & "'"
If Not RsTxn_GodownRent.EOF Then
   Call Indata
   Frame0.Visible = False
End If
End Sub

Private Sub MSHFlexGrid2_Click()
If SaveCmd.Enabled = False Then Exit Sub
Dim r As Variant
r = MSHFlexGrid2.RowSel
If r < 2 Then Exit Sub
If MSHFlexGrid2.TextMatrix(r, 1) = "" Then Exit Sub

If LCase(Gen_UserRole) = "power" Or LCase(Gen_UserRole) = "superadmin" Then
   mSelRow = MSHFlexGrid2.row
   mSelCol = MSHFlexGrid2.Col
   If mSelCol <> 7 Then Exit Sub
   TxtArrears.Visible = True
   TxtArrears.Text = MSHFlexGrid2.TextMatrix(MSHFlexGrid2.row, 7)
   With TxtArrears
        .Left = (MSHFlexGrid2.Left + MSHFlexGrid2.CellLeft) - 30
        .Top = (MSHFlexGrid2.Top + MSHFlexGrid2.CellTop) - 30
        .Width = MSHFlexGrid2.CellWidth + 35
        .Height = MSHFlexGrid2.CellHeight '- 200
        .ZOrder
        .SetFocus
   End With
End If
End Sub
Private Sub SaveCmd_Click()

If TxtExpenseFrom.Value > TxtExpenseTo.Value Then
   MsgBox "'Expense From ' must be less than 'Expense To'!!!"
   TxtExpenseFrom.SetFocus
   Exit Sub
End If

If CmbForwardTo.Text = "" Then
   MsgBox "Blank Forward to not allowed!!!"
   CmbForwardTo.SetFocus
   Exit Sub
End If

''-- Ullhas Patil 26 Oct 2017
If CmbBillingState.Text = "" Then
   MsgBox "Blank Billing State not allowed!!!"
   CmbBillingState.SetFocus
   Exit Sub
End If
''---------------------------

If Gen_DesignationID = 6 Then
   tmp = "Select UG.GroupName from Mst_Users MU"
   tmp = tmp & " Inner Join  Mst_UserRole UG On UG.UserGroupID=MU.UserGroupID"
   tmp = tmp & " Inner Join  Mst_Employee ME On ME.EmployeeNO=MU.EmployeeNo"
   tmp = tmp & " Where Me.EmployeeID = " & mForwardToID & ""
   
   Call TmpTable
   If TmpRs.RecordCount > 0 Then
      If LCase(TmpRs!GroupName) = "fag" Or LCase(TmpRs!GroupName) = "fagadmin" Then
         MsgBox "You Can't Forward this Claim to FAG  !!!!"
         Exit Sub
      End If
   End If
End If

If Edit_Flg = "E" Then
   'If LCase(Gen_UserRole) = "fag" Then ''' Or LCase(Gen_UserRole) = "power"
   If CmbStatus.Text = "" Then
      MsgBox "Blank Status not allowed!!!"
      CmbStatus.SetFocus
      Exit Sub
   End If
   If LCase(Gen_UserRole) = "fag" Or LCase(Gen_UserRole) = "fagadmin" Then
      If LCase(CmbStatus.Text) = "approved" Or LCase(CmbStatus.Text) = "rejected" Or LCase(CmbStatus.Text) = "settled" Then
         If TxtApprovalRemarks = "" Then
            MsgBox "Blank FAG Remarks not allowed!!!"
            TxtApprovalRemarks.SetFocus
            Exit Sub
         End If
         If TxtApprovalDate.Value > Date Then
            MsgBox "Future NAV Booked date not allowed!!!"
            TxtApprovalDate.SetFocus
            Exit Sub
         End If
         If Format(Gen_BlockingDate, "yyyy-mm") >= Format(CDate(TxtApprovalDate.Value), "yyyy-mm") Then
            MsgBox "Select next month from " & Format(Gen_BlockingDate, "MMM-yyyy") & "!!! "
            Exit Sub
         End If
    End If
   End If
   If LCase(CmbStatus.Text) = "settled" Then
      If Val(TxtFAGDisbAmt.Text) > 0 Then
         If TxtFagDisbDate.Value > Date Then
            MsgBox "Future date not allowed!!"
            TxtFagDisbDate.SetFocus
            SSTab1.Tab = 3
            Exit Sub
         End If
         If TxtFagDisbDate.Value < TxtClaimDate.Value Then
            MsgBox "Disbursement Date should not less than Claim Date!!"
            TxtFagDisbDate.SetFocus
            SSTab1.Tab = 1
            Exit Sub
         End If
         If CmbFAGBank.Text = "" Then
            MsgBox "Blank Bank-A/C No not allowd !!"
            If CmbFAGBank.Visible = True Then
               SSTab1.Tab = 1
               CmbFAGBank.SetFocus
            End If
            Exit Sub
         End If
         If TxtFAGRemark.Text = "" Then
            MsgBox "Blank  Disbursement Remark not allowed!!"
            SSTab1.Tab = 1
            TxtFAGRemark.SetFocus
            Exit Sub
         End If
         If Len(TxtFAGRemark) > 250 Then
            MsgBox "Length of disbursment remark is more than 250 characters , please check!!!"
            TxtFAGRemark.SetFocus
            Exit Sub
         End If
         
         If CmbModeOfPay.Text = "" Then
            MsgBox "Blank Mode of Payment not allowed!!"
            CmbModeOfPay.SetFocus
            Exit Sub
         End If
         
         If LCase(CmbModeOfPay.Text) = "bank transfer" Or LCase(CmbModeOfPay.Text) = "neft" Or LCase(CmbModeOfPay.Text) = "rtgs" Then
            If TxtACNo.Text = "" Then
               MsgBox "Blank Account Details not allowed for 'Bank Transfer', Contact FAG to update Vendor Account details!!"
               CmbModeOfPay.SetFocus
               Exit Sub
            End If
         End If
         
         If Val(TxtFAGDisbAmt) < Val(TxtClaimAmount) Then
            If Trim(TxtDeductionRemark) = "" Then
               MsgBox "As disbursment amount is less than claim amount please provide deduction remarks!!!"
               TxtDeductionRemark.SetFocus
               Exit Sub
            End If
         End If

      End If
   End If
End If

If TxtRemark = "" Then
   MsgBox "Blank remarks not allowed!!!"
   TxtRemark.SetFocus
   Exit Sub
End If

If Left(LCase(CmbStatus.Text), 1) <> "r" Then
   If MSHFlexGrid2.Rows < 4 Then
      MsgBox "No details found in grid!!!"
      CmbStatus.SetFocus
      Exit Sub
   End If
End If

If Edit_Flg = "A" Then
   RsTxn_GodownRent.AddNew
   RsTxn_GodownRent!ClaimID = GetMaxClaimID
   RsTxn_GodownRent!G_UID = GetGUID
   TxtClaimID = RsTxn_GodownRent!ClaimID
   TxtG_UID = RsTxn_GodownRent!G_UID
Else
   RsTxn_GodownRent.MoveFirst
   RsTxn_GodownRent.Find "G_UID= '" & TxtG_UID & "'"
End If

RsTxn_GodownRent!ClaimDate = TxtClaimDate.Value
'RsTxn_GodownRent!LessorID = mLessorID

RsTxn_GodownRent!Amount = Val(TxtClaimAmount)
'RsTxn_GodownRent!Amount = Val(MSHFlexGrid2.TextMatrix(0, 8))

RsTxn_GodownRent!Flag = mStatus
RsTxn_GodownRent!Remarks = TxtRemark
If TxtApprovalRemarks.Locked = False Then
   RsTxn_GodownRent!ApprovalRemarks = Trim(TxtApprovalRemarks)
End If
If TxtApprovalDate.Visible = True Then
   RsTxn_GodownRent!ApprovalDate = TxtApprovalDate.Value
End If
RsTxn_GodownRent!ForwardToID = mForwardToID
'RsTxn_GodownRent!G_UID = TxtG_UID
RsTxn_GodownRent!ExpenseFrom = TxtExpenseFrom.Value
RsTxn_GodownRent!ExpenseTo = TxtExpenseTo.Value
RsTxn_GodownRent!BatchID = mBatchID

If LCase(CmbStatus.Text) = "settled" Then
   RsTxn_GodownRent!DisbuDate = TxtFagDisbDate.Value
   RsTxn_GodownRent!DisbuRemark = TxtFAGRemark.Text
   RsTxn_GodownRent!DisbuAmount = TxtFAGDisbAmt.Text
   RsTxn_GodownRent!ModeOfPay = CmbModeOfPay.Text
   RsTxn_GodownRent!FAGBankID = mFAGBankID
   RsTxn_GodownRent!DeductionRemark = Trim(TxtDeductionRemark)
Else
   RsTxn_GodownRent!DisbuDate = Null
   RsTxn_GodownRent!DisbuRemark = Null
   RsTxn_GodownRent!DisbuAmount = Null
   RsTxn_GodownRent!ModeOfPay = Null
   RsTxn_GodownRent!FAGBankID = Null
   RsTxn_GodownRent!DeductionRemark = Null
End If

RsTxn_GodownRent!BillingStateIndiaID = mBillingStateID      ''-- Ullhas Patil 26 Oct 2017

If IsNull(RsTxn_GodownRent!CreatedBy) = True Or RsTxn_GodownRent!CreatedBy = "" Then
   RsTxn_GodownRent!CreatedBy = Gen_UserLoginID
   RsTxn_GodownRent!CreatedDate = Now
Else
   RsTxn_GodownRent!UpdatedBy = Gen_UserLoginID
   RsTxn_GodownRent!UpdatedDate = Now
End If

If LCase(Gen_UserRole) = "fag" Or LCase(Gen_UserRole) = "power" Or LCase(Gen_UserRole) = "fagadmin" Then
   If LCase(mStatus) = "a" Or LCase(mStatus) = "s" Then
      RsTxn_GodownRent!UpdatedByFAG = Gen_UserLoginID
      RsTxn_GodownRent!UpdatedDateFAG = Now
   End If
End If

RsTxn_GodownRent.Update

TxtCreated = IIf(IsNull(RsTxn_GodownRent!CreatedBy), "", RsTxn_GodownRent!CreatedBy) & " :- " & IIf(IsNull(RsTxn_GodownRent!CreatedDate), "", RsTxn_GodownRent!CreatedDate)
TxtUpdated = IIf(IsNull(RsTxn_GodownRent!UpdatedBy), "", RsTxn_GodownRent!UpdatedBy) & " :- " & IIf(IsNull(RsTxn_GodownRent!UpdatedDate), "", RsTxn_GodownRent!UpdatedDate)

Call SaveLessorArrearsDetails

If LCase(mStatus) = "r" Then
   Call UpdateRentPaidUpTo
End If


'If LCase(mStatus) <> "f" Then
  Call Gen_Email
'End If
RsTxn_GodownRent.Filter = ""
RsTxn_GodownRent.Requery
If RsTxn_GodownRent.RecordCount > 0 Then
   RsTxn_GodownRent.MoveFirst
   RsTxn_GodownRent.Find "G_UID = '" & TxtG_UID & "'"
   Call Indata
Else
   MsgBox "Data saved!!!"
   
   MsgBox "No Pending Claim under your ID!!!!"
   'Call Grid_Head
   'Call GButtonLock_1(Me, True)
   Unload Me
   
End If
End Sub

Public Sub UpdateRentPaidUpTo()
For I = 2 To MSHFlexGrid2.Rows - 1
    If MSHFlexGrid2.TextMatrix(I, 1) <> "" Then
       tmp = " Select MAX(TGR.ExpenseTo) as ExpenseTo From Txn_GodownRentDetail TGRD"
       tmp = tmp & " Inner join Txn_GodownRent TGR on TGRD.ClaimID=TGR.ClaimID"
       tmp = tmp & " Where TGR.Flag <> 'R' And TGRD.LAID=" & Val(MSHFlexGrid2.TextMatrix(I, 1)) & " AND TGRD.GodownID =" & Val(MSHFlexGrid2.TextMatrix(I, 14)) & " "
       
       Call Tmp1Table
       If TmpRs1.RecordCount > 0 And Not IsNull(TmpRs1!ExpenseTo) Then
          If Format(TmpRs1!ExpenseTo, "MM/dd/yyyy") < Format(TxtExpenseTo.Value, "MM/dd/yyyy") Then
             Call TableTxn_GodownLessorAgreeDetail("Where LAID=" & Val(MSHFlexGrid2.TextMatrix(I, 1)) & " AND GodownID =" & Val(MSHFlexGrid2.TextMatrix(I, 14)) & "")
             If RsTxn_GodownLessorAgreeDetail.RecordCount > 0 Then
                RsTxn_GodownLessorAgreeDetail!RentPaidUpTo = TmpRs1!ExpenseTo
                RsTxn_GodownLessorAgreeDetail.Update
             End If
          End If
       End If
    End If
Next
End Sub

Private Sub InDataLogHistory()

Call Grid_HeadHistory


tmp = "Select ClaimID , OldForwardToID, NewForwardToID, UpdatedBy, Concurrency"
tmp = tmp & " From Log_GodownRentHistory Where ClaimID=" & Val(TxtClaimID) & ""

Call TmpTable
MSHFlexGrid4.Rows = MSHFlexGrid4.Rows + TmpRs.RecordCount
For I = 1 To MSHFlexGrid4.Rows - 2
   mTmpEmpName = ""
   MSHFlexGrid4.TextMatrix(I, 0) = IIf(IsNull(TmpRs!ClaimID), 0, TmpRs!ClaimID)
   mTmpEmpName = GetEmployeeNonName(IIf(IsNull(TmpRs!OldForwardToID), "", TmpRs!OldForwardToID))
   MSHFlexGrid4.TextMatrix(I, 1) = mTmpEmpName
   mTmpEmpName = GetEmployeeNonName(IIf(IsNull(TmpRs!NewForwardToID), 0, TmpRs!NewForwardToID))
   MSHFlexGrid4.TextMatrix(I, 2) = mTmpEmpName

   MSHFlexGrid4.TextMatrix(I, 3) = IIf(IsNull(TmpRs!UpdatedBy), "", TmpRs!UpdatedBy)
   MSHFlexGrid4.TextMatrix(I, 4) = Format(IIf(IsNull(TmpRs!Concurrency), "", TmpRs!Concurrency), "dd-mmm-yyyy hh:mm:ss")
   TmpRs.MoveNext
Next
End Sub
Private Sub Grid_HeadHistory()
With MSHFlexGrid4
     .Clear
     .Rows = 2
     .Cols = 5
     .AllowUserResizing = flexResizeBoth
     .TextMatrix(0, 0) = "Rent ID"
     .TextMatrix(0, 1) = "Old Forward ID"
     .TextMatrix(0, 2) = "New Forward ID"
     .TextMatrix(0, 3) = "Updated By"
     .TextMatrix(0, 4) = "Updated On"
     .ColWidth(0) = 1000
     .ColWidth(1) = 3000
     .ColWidth(2) = 3000
     .ColWidth(3) = 2000
     .ColWidth(4) = 2300
     
End With
End Sub
Private Sub Gen_Email()
Dim TOADD As Variant
Dim TOCC, mMsg, mGridMsg As Variant
Dim strSplitString() As String
Dim EmpNo_, EmpName_ As String
Dim mPANno, mAccNo, mRTGSNo, mBankName, mBranchName As Variant
Dim mRow, mCol, mLocationName, mModeOfPayment, mData As Variant

mLocationName = GetLocationFromCMPName(MSHFlexGrid2.TextMatrix(2, 2))

tmp = "Select ML.LessorName,ML.PANNo,ML.AccountNo,ML.RTGSNo,"
tmp = tmp & " MB.BankName , MBR.BranchName"
tmp = tmp & " From Mst_Lessor ML"
tmp = tmp & " Left Join Mst_Bank MB on MB.BankID= ML.BankID"
tmp = tmp & " Left Join Mst_Branch MBR on MBR.BranchID=ML.BranchID"
tmp = tmp & " Where ML.LessorID = " & mLessorID & ""
Call TmpTable


If TmpRs.RecordCount > 0 Then
   mPANno = IIf(IsNull(TmpRs!PANNO), "", TmpRs!PANNO)
   mAccNo = IIf(IsNull(TmpRs!AccountNo), "", TmpRs!AccountNo)
   mRTGSNo = IIf(IsNull(TmpRs!RTGSNo), "", TmpRs!RTGSNo)
   mBankName = IIf(IsNull(TmpRs!BankName), "", TmpRs!BankName)
   mBranchName = IIf(IsNull(TmpRs!BranchName), "", TmpRs!BranchName)
End If
mPANno = IIf(mPANno = "", "&nbsp", mPANno)
mAccNo = IIf(mAccNo = "", "&nbsp", mAccNo)
mRTGSNo = IIf(mRTGSNo = "", "&nbsp", mRTGSNo)
mBankName = IIf(mBankName = "", "&nbsp", mBankName)
mBranchName = IIf(mBranchName = "", "&nbsp", mBranchName)
mModeOfPayment = IIf(mModeOfPayment = "", "&nbsp", mModeOfPayment)


EmpNo_ = Split(TxtCreated.Text, ":")
mMsg = ""
Call TableMst_Employee("Where EmployeeNo='" & Trim(EmpNo_(0)) & "'")
If Not RsMst_Employee.EOF Then
   EmpName_ = RsMst_Employee!EmployeeName & "~" & RsMst_Employee!EmployeeNo
   
   tmp = "Select EmailID from Mst_Users Where EmployeeNo='" & EmpNo_(0) & "'"
   Call TmpTable
   If Not TmpRs.EOF Then
      mEmployeeEmailID = TmpRs!EmailID
   Else
      mEmployeeEmailID = ""
   End If
Else
   EmpName_ = ""
End If

mMsg = "<tr>"
mMsg = mMsg & " <td colspan=" & 6 & " >"
mMsg = mMsg & "   <C> <b>Godown Rent Claim Details</b></C><br /></td>"
mMsg = mMsg & "</tr>"
mMsg = mMsg & "<tr>"
mMsg = mMsg & "    <td ><b>Rent Payment ID</b></td>"
mMsg = mMsg & "    <td >" & TxtClaimID & "</td>"
mMsg = mMsg & "    <td ><b>Date</b></td>"
mMsg = mMsg & "    <td >" & Format(TxtClaimDate.Value, "dd-mmm-yyyy") & "</td>"
mMsg = mMsg & "    <td ><b>Location</b></td>"
mMsg = mMsg & "    <td >" & mLocationName & "</td>"
mMsg = mMsg & "</tr>"
mMsg = mMsg & "<tr>"
mMsg = mMsg & "    <td ><b>Initiated by</b></td>"
mMsg = mMsg & "    <td >" & EmpName_ & "</td>"  '" & TxtCreated.Text & "
mMsg = mMsg & "    <td ><b>Currently with</b></td>"
mMsg = mMsg & "    <td >" & TxtForwardTo.Text & "</td>"
mMsg = mMsg & "    <td ><b>Status</b></td>"
mMsg = mMsg & "    <td >" & CmbStatus.Text & "</td>"
mMsg = mMsg & "</tr>"
mMsg = mMsg & "<tr>"
mMsg = mMsg & "    <td ><b>Lessor Name</b></td>"
mMsg = mMsg & "    <td >" & TxtLessorID.Text & "</td>"
mMsg = mMsg & "    <td ><b>PAN No</b></td>"
mMsg = mMsg & "    <td >" & mPANno & "</td>"
mMsg = mMsg & "    <td ><b>Rent Period</b></td>"
mMsg = mMsg & "    <td >" & Format(TxtExpenseFrom.Value, "dd-MMM-yyyy") & " " & " To   " & Format(TxtExpenseTo.Value, "dd-MMM-yyyy") & " </td>"
mMsg = mMsg & "</tr>"
mMsg = mMsg & "<tr>"
mMsg = mMsg & "    <td ><b>Mode of Payment</b></td>"
mMsg = mMsg & "    <td >" & mModeOfPayment & " </td>"
mMsg = mMsg & "    <td ><b>Total Amount</b></td>"
mMsg = mMsg & "    <td >" & Format(TxtClaimAmount.Text, "#######0.00") & "</td>"
mMsg = mMsg & "    <td >&nbsp</td>"
mMsg = mMsg & "    <td >&nbsp</td>"
mMsg = mMsg & "</tr>"
mMsg = mMsg & "<tr>"
mMsg = mMsg & "    <td ><b>Bank Name</b></td>"
mMsg = mMsg & "    <td >" & mBankName & "</td>"
mMsg = mMsg & "    <td ><b>Branch Name</b></td>"
mMsg = mMsg & "    <td >" & mBranchName & "</td>"
mMsg = mMsg & "    <td ><b>Account No</b></td>"
mMsg = mMsg & "    <td >" & mAccNo & "</td>"
mMsg = mMsg & "</tr>"
mMsg = mMsg & "<tr>"
mMsg = mMsg & "    <td ><b>RTGS Code</b></td>"
mMsg = mMsg & "    <td > " & mRTGSNo & "</td>"
mMsg = mMsg & "    <td >&nbsp</td>"
mMsg = mMsg & "    <td >&nbsp</td>"
mMsg = mMsg & "    <td >&nbsp</td>"
mMsg = mMsg & "    <td >&nbsp</td>"
mMsg = mMsg & "</tr>"


mMsg = mMsg & " <tr> <td colspan=" & 6 & " >"

'''strSplitString = Split(CmbForwardTo.Text, "~")
'''EmpName_ = Trim(strSplitString(0))
'''EmpNo_ = Trim(strSplitString(1))
'''mMsg = mMsg & "<tr><td> Forwarded To </td><td> " & EmpName_ & " " & " ~ " & " " & EmpNo_ & " </td></tr>"

mMsg = mMsg & "<P><b> Godown wise Details </b> </P>"

mGridMsg = "<table border = 1 cellpadding=0  cellSpacing=0>"


For mRow = 1 To MSHFlexGrid2.Rows - 2
    If mRow = 1 Then
       mGridMsg = mGridMsg & "<tr bgcolor=" & "#E9CFEC" & "   style=" & "font:bold" & " >"
    Else
       mGridMsg = mGridMsg & "<tr>"
    End If
    For mCol = 1 To MSHFlexGrid2.Cols - 2
        If MSHFlexGrid2.TextMatrix(mRow, 1) <> "" Then
           mData = IIf(MSHFlexGrid2.TextMatrix(mRow, mCol) = "", "&nbsp", MSHFlexGrid2.TextMatrix(mRow, mCol))
           mGridMsg = mGridMsg & "<td> " & mData & " </td> "
        End If
    Next
    mGridMsg = mGridMsg & "</tr>"
Next


mMsg = mMsg & mGridMsg & "</table> </td> </tr>"

TOADD = mForwardToEmailID ' mLessorEmailID
TOCC = Gen_UserEmailID  '& "," & mForwardToEmailID

If LCase(mStatus) = "p" Then
   mMsgSubject = "Godown Rent Claim Pending"
ElseIf LCase(mStatus) = "a" Then
   mMsgSubject = "Godown Rent Claim Approved"
ElseIf LCase(mStatus) = "r" Then
   mMsgSubject = "Godown Rent Claim Rejected"
ElseIf LCase(mStatus) = "s" Then
   mMsgSubject = "Godown Rent Claim Settled"

End If

Call SendMail(TOADD, TOCC, mMsgSubject, mMsg)


End Sub
Private Function SendMail(ToADD_ As Variant, TOCC_ As Variant, MsgSubject_ As Variant, Message1_ As Variant)
Dim htmlStringhead As String
Dim htmlStringtail As String
Set oSMTP = New OSSMTP.SMTPSession

'servertxt = "10.0.0.13"

'htmlStringhead = "<html><body><font size= 3 face=Zurich BT color = Blue><table border = 1 cellpadding=1 cellspacing =1>"
'htmlStringtail = "</table></Font></body></html>"
'
'With oSMTP
'  'connection
'  .Server = servertxt
'  'message
'  .MailFrom = TOCC_ '"mayur.d@ncmsl.com"
'  If IsNull(Trim(ToADD_)) = True Or Trim(ToADD_) = "" Then
'     ToADD_ = "taufique.k@ncml.com"
'  End If
'  .SendTo = ToADD_
'  TOCC_ = TOCC_ & "," & mEmployeeEmailID  '& "," & mLessorEmailID
'  If ToADD_ <> TOCC_ Then
'     .CC = TOCC_
'  End If
'  '.BCC = "mayur.d@ncmsl.com"
'  '.BCC = "parag.r@ncmsl.com" '"mayur.d@ncmsl.com"
'  .MessageSubject = MsgSubject_ & " (from Store-Man)" '"Confidential Mail"
'  .MessageHTML = htmlStringhead & Message1_ & htmlStringtail
'  .TimeStamp = Now()
'  If LCase(Gen_EmailConfigFlag) = "y" Then
'     .SendEmail
'  End If
'End With


htmlStringhead = "<html><body><font size= 3 face= Candara color = Blue>" '<table border = 1>"
htmlStringtail = "</Font></body></html>" '</table>


Dim objSMTP
Set objSMTP = CreateObject("MailBee.SMTP")
objSMTP.LicenseKey = "MBC900-5F470FAD67-D8E202A2E22927335F24F5A667BFD196"
objSMTP.ServerName = Gen_ReadMailServerIPINI("MailIP") ' = "mail.ncml.com" '"mail.ncml.com"

objSMTP.FromAddr = TOCC_
If IsNull(Trim(ToADD_)) = True Or Trim(ToADD_) = "" Then
   ToADD_ = "taufique.k@ncml.com"
End If
objSMTP.ToAddr = ToADD_
If ToADD_ <> TOCC_ Then
   objSMTP.CCAddr = TOCC_ & "," & mEmployeeEmailID  '& "," & mLessorEmailID
End If

objSMTP.Subject = MsgSubject_ & " (from Store-Man)" 'MsgSubject_ '"Message in HTML format"
objSMTP.BodyFormat = 1
objSMTP.BodyText = htmlStringhead & Message1_ & htmlStringtail 'htmlStringhead & MsgSubject_ & " :- " & Message1_ & htmlStringtail '"<html>This word is <b>bold</b></html>"
If LCase(Gen_EmailConfigFlag) = "y" Then
   objSMTP.Send
End If


End Function

Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   If RsTxn_GodownRent.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   RsTxn_GodownRent.MoveFirst
 
   RsTxn_GodownRent.Find "G_UID = '" & TxtG_UID & "'"
   If RsTxn_GodownRent.EOF Then
      RsTxn_GodownRent.MoveLast
   End If
   Call Indata
   Exit Sub
End If

'Frame0.Height = Me.Height - 500
'Frame0.Left = Frame1.Left - 100
'Frame0.Top = Frame1.Top
'Frame0.Width = Me.Width - 200
'Frame0.Visible = True

Frame0.Height = Frame4.Height - 100
Frame0.Left = Frame1.Left '- 100
Frame0.Top = Frame1.Top
Frame0.Width = Frame4.Width '- 100
Frame0.Visible = True

MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 2000

RsTxn_GodownRent.MoveFirst

''===========comment on 7 july 2012 for not clearing search grid=============
'Call Grid_Head1
'
'SFrom.Value = CDate(From_GenDate) - 900
'STo.Value = Date 'To_GenDate
'CmbSLessor.Text = ""
'CmbSStatus.Text = ""
'CmbSForwardTo.Text = ""
'CmbSLocation.Text = ""
'ChkBatchNo.Value = 0
'TxtSClaimID = ""
'Label22.Caption = ""
'
'tmp = "Select EmployeeID,EmployeeName +'~'+ EmployeeNo as Employee,DesignationID,EmployeeEmailID From Mst_Employee Where EmployeeNo = '" & Gen_UserLoginID & "'"
'Call TmpTable
'
'If TmpRs.RecordCount = 1 Then
'   CmbSForwardTo.Text = TmpRs!Employee
'   mSForwardToID = TmpRs!EmployeeID
'Else
'   CmbSForwardTo.Text = ""
'   mSForwardToID = Null
'End If
''===========comment on 7 july 2012 for not clearing search grid=============

End Sub


Private Function GetMaxClaimID() As Double
Dim Retval As Double
tmp = "Select max(ClaimID) from Txn_GodownRentClaim "
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetMaxClaimID = Retval
End Function

Function GetMaxLessorClaimDetails() As Double
Dim Retval As Double
Retval = 0
tmp = "Select max(ClaimDetailID) From Txn_GodownRentClaimDetail " 'where ClaimID = " & TxtClaimID & " And SM_prefix = '" & TxtPreFix & "'"
Call TmpTable
If TmpRs.RecordCount > 0 Then
   Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
End If
Retval = Retval + 1
TmpRs.Close
GetMaxLessorClaimDetails = Retval

End Function
Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsTxn_GodownRent.MoveNext
If RsTxn_GodownRent.EOF Then
   RsTxn_GodownRent.MoveLast
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
RsTxn_GodownRent.MovePrevious
If RsTxn_GodownRent.BOF Then
   RsTxn_GodownRent.MoveFirst
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
RsTxn_GodownRent.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsTxn_GodownRent.MoveLast
Call Indata
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False
End Sub
Public Sub Indata()

TxtApprovalDate.Visible = False
Label27.Visible = False

mStatus = ""
FrmFAG.Enabled = False
TxtForwardTo.Visible = True
CmbForwardTo.Visible = False
CmdForwardTo.Visible = False

TxtStatus.Visible = True
CmbStatus.Visible = False

TxtExpenseFrom.Enabled = False
TxtExpenseTo.Enabled = False
TxtClaimDate.Enabled = False

TxtClaimID = RsTxn_GodownRent!ClaimID
TxtClaimDate.Value = RsTxn_GodownRent!ClaimDate
TxtExpenseFrom.Value = RsTxn_GodownRent!ExpenseFrom
TxtExpenseTo.Value = RsTxn_GodownRent!ExpenseTo
mLessorID = RsTxn_GodownRent!LessorID
Call ShowAccountDetails(mLessorID)
TxtLessorID = GetLessorName(RsTxn_GodownRent!LessorID)
'mLessorEmailID = GetLessorName(mLessorID)
TxtClaimAmount = RsTxn_GodownRent!Amount 'a 'Val(MSHFlexGrid2.TextMatrix(0, 4))

CmdPrint.Enabled = True

mStatus = RsTxn_GodownRent!Flag
mBatchID = IIf(IsNull(RsTxn_GodownRent!BatchID), 0, RsTxn_GodownRent!BatchID)
mPBatchID = IIf(IsNull(RsTxn_GodownRent!PBatchID), 0, RsTxn_GodownRent!PBatchID)

mForwardToID = IIf(IsNull(RsTxn_GodownRent!ForwardToID), "", RsTxn_GodownRent!ForwardToID)

If mForwardToID <> "" Then
   TxtForwardTo = GetEmployeeNonName(RsTxn_GodownRent!ForwardToID)
   mForwardToEmailID = GetEmployeeEmailID(mForwardToID)
Else
   TxtForwardTo = "" 'GetLessorNonName(RsTxn_GodownRent!ForwardToID)
   mForwardToEmailID = Null 'GetLessorEmailID(mForwardToID)
End If
TxtFWDDesignation = GetDesignationFromEmp(mForwardToID)

If LCase(mStatus) = "a" Then
   TxtStatus = "Approved"
ElseIf LCase(mStatus) = "p" Then
   TxtStatus = "Pending"
ElseIf LCase(mStatus) = "r" Then
   TxtStatus = "Rejected"
ElseIf LCase(mStatus) = "s" Then
   TxtStatus = "Settled"
ElseIf LCase(mStatus) = "f" Then
   TxtStatus = "Frozen"
Else
   TxtStatus = ""
End If

TxtRemark = IIf(IsNull(RsTxn_GodownRent!Remarks), "", RsTxn_GodownRent!Remarks)
TxtApprovalRemarks = IIf(IsNull(RsTxn_GodownRent!ApprovalRemarks), "", RsTxn_GodownRent!ApprovalRemarks)

If IsNull(RsTxn_GodownRent!ApprovalDate) = False Then
   TxtApprovalDate.Visible = True
   Label27.Visible = True
   TxtApprovalDate.Value = RsTxn_GodownRent!ApprovalDate
   TxtApprovalDate.Enabled = False
End If

TxtG_UID = IIf(IsNull(RsTxn_GodownRent!G_UID), "", RsTxn_GodownRent!G_UID)

TxtRemark.Locked = True
TxtApprovalRemarks.Locked = True
TxtApprovalRemarks.Locked = True

CmbModeOfPay.Visible = False
TxtModeOfPayment.Visible = True

CmbFAGBank.Visible = False
TxtFAGBank.Visible = True

''--- Ullhas Patil 26 Oct 2017
CmbBillingState.Visible = False: CmbBillingState = ""
TxtBillingState.Visible = True: TxtBillingState = ""
''-----------------------------

TxtFAGDisbAmt = IIf(IsNull(RsTxn_GodownRent!DisbuAmount), "", RsTxn_GodownRent!DisbuAmount)
TxtFAGRemark = IIf(IsNull(RsTxn_GodownRent!DisbuRemark), "", RsTxn_GodownRent!DisbuRemark)
TxtFagDisbDate.Value = IIf(IsNull(RsTxn_GodownRent!DisbuDate), Date, RsTxn_GodownRent!DisbuDate)
mFAGBankID = IIf(IsNull(RsTxn_GodownRent!FAGBankID), 0, RsTxn_GodownRent!FAGBankID)
TxtFAGBank.Text = GetFAGBankAcNo(Val(mFAGBankID))
TxtModeOfPayment.Text = IIf(IsNull(RsTxn_GodownRent!ModeOfPay), "", RsTxn_GodownRent!ModeOfPay)
TxtDeductionRemark = IIf(IsNull(RsTxn_GodownRent!DeductionRemark), "", RsTxn_GodownRent!DeductionRemark)

TxtDocNo = IIf(IsNull(RsTxn_GodownRent!DocNo), "", RsTxn_GodownRent!DocNo)
TxtEntryNo = IIf(IsNull(RsTxn_GodownRent!EntryNo), "", RsTxn_GodownRent!EntryNo)
TxtExportedDate = IIf(IsNull(RsTxn_GodownRent!FAGPayExportedDate), "", Format(RsTxn_GodownRent!FAGPayExportedDate, "dd-MMM-yyyy"))

TxtBankID.Text = IIf(IsNull(RsTxn_GodownRent!BankName), "", RsTxn_GodownRent!BankName)
TxtChequeNo.Text = IIf(IsNull(RsTxn_GodownRent!ChequeNo), "", RsTxn_GodownRent!ChequeNo)
TxtPaidAmt.Text = IIf(IsNull(RsTxn_GodownRent!PaidAmount), "", Format(RsTxn_GodownRent!PaidAmount, "#####0.00"))
TxtTDS.Text = IIf(IsNull(RsTxn_GodownRent!TDS), "", Format(RsTxn_GodownRent!TDS, "#####0.00"))
TxtPaidDate.Text = IIf(IsNull(RsTxn_GodownRent!PaidDate), "", Format(RsTxn_GodownRent!PaidDate, "dd-MMM-yyyy"))

Frame3.Enabled = False

TxtCreated = IIf(IsNull(RsTxn_GodownRent!CreatedBy), "", RsTxn_GodownRent!CreatedBy) & " :- " & IIf(IsNull(RsTxn_GodownRent!CreatedDate), "", RsTxn_GodownRent!CreatedDate)
TxtUpdated = IIf(IsNull(RsTxn_GodownRent!UpdatedBy), "", RsTxn_GodownRent!UpdatedBy) & " :- " & IIf(IsNull(RsTxn_GodownRent!UpdatedDate), "", RsTxn_GodownRent!UpdatedDate)
TxtUpdatedByFAG = IIf(IsNull(RsTxn_GodownRent!UpdatedByFAG), "", RsTxn_GodownRent!UpdatedByFAG) & " :- " & IIf(IsNull(RsTxn_GodownRent!UpdatedDateFAG), "", RsTxn_GodownRent!UpdatedDateFAG)
''--- Ullhas Patil 26 Oct 2017
mBillingStateID = IIf(IsNull(RsTxn_GodownRent!BillingStateIndiaID), 0, RsTxn_GodownRent!BillingStateIndiaID)
TxtBillingState = GetStateIndiaName(mBillingStateID)
TxtBillingState = IIf(Trim(TxtBillingState) = "0", "", TxtBillingState)
''-----------------------------

Call InDataClaimDetails
Call InDataLogHistory

If RsTxn_GodownRent.RecordCount = 1 Then
   Call GButtonLock_1(Me, True)
Else
   Call GButtonLock(Me, True)
End If

SSTab1.Tab = 0

End Sub
Private Sub InDataClaimDetails()
Dim mSerTax As Variant
Dim mStock As Variant

Call Grid_Head

tmp = " Select GodownID,Sum(WHRWeight) as 'Stock' From"
'Parag
'tmp = tmp & " Mst_gsl  WHERE (Flag <> 'E') AND (Flag <> 'Z') group by GodownID"
tmp = tmp & " Mst_gsl  WHERE (Status <> 'E') AND (Status <> 'Z') group by GodownID"

Call Tmp1Table
'MS.Description,

tmp = "Select TGRD.ClaimDetailID,TLA.LAID,MC.CMPName,MG.GodownNo,MG.Address,TLA.ExpiryDate,"
tmp = tmp & " MG.CloseDate,MG.Area ,"
tmp = tmp & " TGR.Flag , TGRD.Amount,TLA.AdvanceGiven,TLA.Remark, MG.GodownID, TGR.LessorID,TLA.ServiceTaxID,TGRD.Arrears"
tmp = tmp & " from Txn_GodownRent TGR"
tmp = tmp & " Inner Join Txn_GodownRentDetail TGRD On TGR.ClaimID = TGRD.ClaimID"
tmp = tmp & " Inner Join Mst_Godown MG On MG.GodownID=TGRD.GodownID"
tmp = tmp & " Inner Join Mst_CMP MC On MC.CMPID = MG.CMPID"
tmp = tmp & " Inner Join Txn_LessorAgreement TLA on TGRD.LAID=TLA.LAID"

'tmp = tmp & " Inner Join Txn_GodownLessorAgreeDetail TGLA on TGLA.LessorID=TGR.LessorID and TGLA.GodownID=TGRD.GodownID"

'tmp = tmp & " Inner Join Mst_ServiceTax MS on MS.ServiceTaxID=TGLA.ServiceTaxID"
tmp = tmp & " Where TGR.ClaimID=" & TxtClaimID & "  "
Call TmpTable
TxtServiceTax.Text = ""
If TmpRs.RecordCount > 0 Then
    mSerTax = IIf(IsNull(TmpRs!ServiceTaxID), 0, TmpRs!ServiceTaxID)
    If mSerTax <> 0 Then
       TxtServiceTax.Text = GetServiceTaxDescription(mSerTax)
    End If
    For I = 2 To TmpRs.RecordCount + 1
        MSHFlexGrid2.TextMatrix(I, 0) = TmpRs!ClaimDetailID
        MSHFlexGrid2.TextMatrix(I, 1) = TmpRs!LAID
        MSHFlexGrid2.TextMatrix(I, 2) = TmpRs!CMPName
        MSHFlexGrid2.TextMatrix(I, 3) = GetGodownName(TmpRs!GodownID, "N")
        MSHFlexGrid2.TextMatrix(I, 4) = GetGodownName(TmpRs!GodownID, "")
        
        MSHFlexGrid2.TextMatrix(I, 5) = Format(TmpRs!AdvanceGiven, "##########0")
        MSHFlexGrid2.TextMatrix(I, 6) = Format(TmpRs!Amount, "##########0")
        MSHFlexGrid2.TextMatrix(0, 6) = Val(MSHFlexGrid2.TextMatrix(0, 6)) + Val(TmpRs!Amount)
        
        MSHFlexGrid2.TextMatrix(I, 7) = Format(IIf(IsNull(TmpRs!Arrears), 0, TmpRs!Arrears), "##########0.00")
        MSHFlexGrid2.TextMatrix(0, 7) = Val(MSHFlexGrid2.TextMatrix(0, 7)) + Val(IIf(IsNull(TmpRs!Arrears), 0, TmpRs!Arrears))
        
        MSHFlexGrid2.TextMatrix(I, 8) = Val(MSHFlexGrid2.TextMatrix(I, 6)) + Val(MSHFlexGrid2.TextMatrix(I, 7)) 'Format(TmpRs!Arrears, "##########0.00")
        MSHFlexGrid2.TextMatrix(0, 8) = Val(MSHFlexGrid2.TextMatrix(0, 8)) + Val(MSHFlexGrid2.TextMatrix(I, 8))
        
        TmpRs1.MoveFirst
        TmpRs1.Find "GodownID=" & TmpRs!GodownID & ""
        If Not TmpRs1.EOF Then
           MSHFlexGrid2.TextMatrix(I, 9) = TmpRs1!Stock
        End If
        
        MSHFlexGrid2.TextMatrix(I, 10) = TmpRs!Area
        MSHFlexGrid2.TextMatrix(I, 11) = IIf(IsNull(TmpRs!ExpiryDate), "", Format(TmpRs!ExpiryDate, "dd-MMM-yyyy"))
        MSHFlexGrid2.TextMatrix(I, 12) = IIf(IsNull(TmpRs!CloseDate), "", Format(TmpRs!CloseDate, "dd-MMM-yyyy"))
      
        MSHFlexGrid2.TextMatrix(I, 13) = TmpRs!Remark
        MSHFlexGrid2.TextMatrix(I, 14) = TmpRs!GodownID
        MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
        TmpRs.MoveNext
    Next
    TxtClaimAmount = Val(MSHFlexGrid2.TextMatrix(0, 8))
End If
End Sub
Private Sub Grid_Head()
MSHFlexGrid2.Clear
MSHFlexGrid2.Rows = 3
MSHFlexGrid2.FixedRows = 2
MSHFlexGrid2.Cols = 15

'MSHFlexGrid2.TextMatrix(1, 0) = "Select"
MSHFlexGrid2.TextMatrix(1, 0) = "Row ID"
MSHFlexGrid2.TextMatrix(1, 1) = "Agreement ID"
MSHFlexGrid2.TextMatrix(1, 2) = "CMP Name"
MSHFlexGrid2.TextMatrix(1, 3) = "Godown No"
MSHFlexGrid2.TextMatrix(1, 4) = "Godown Address"

MSHFlexGrid2.TextMatrix(1, 5) = "Security Deposit"
MSHFlexGrid2.TextMatrix(1, 6) = "Rent Amount"
MSHFlexGrid2.TextMatrix(1, 7) = "Arrears"
MSHFlexGrid2.TextMatrix(1, 8) = "Total Amount"
MSHFlexGrid2.TextMatrix(1, 9) = "Current Stock"

MSHFlexGrid2.TextMatrix(1, 10) = "Area in Sq. ft."
MSHFlexGrid2.TextMatrix(1, 11) = "Agreement Expiry Date"
MSHFlexGrid2.TextMatrix(1, 12) = "Godown Close Date"


MSHFlexGrid2.TextMatrix(1, 13) = "Remarks"
MSHFlexGrid2.TextMatrix(1, 14) = "GodownID"

MSHFlexGrid2.ColWidth(0) = 800
MSHFlexGrid2.ColWidth(1) = 1000
MSHFlexGrid2.ColWidth(2) = 1500
MSHFlexGrid2.ColWidth(3) = 900
MSHFlexGrid2.ColWidth(4) = 3000

MSHFlexGrid2.ColWidth(5) = 1200
MSHFlexGrid2.ColWidth(6) = 1500
MSHFlexGrid2.ColWidth(7) = 1200
MSHFlexGrid2.ColWidth(8) = 1500
MSHFlexGrid2.ColWidth(9) = 1500
MSHFlexGrid2.ColWidth(10) = 1700
MSHFlexGrid2.ColWidth(11) = 1500
'MSHFlexGrid2.ColWidth(11) = 2300
MSHFlexGrid2.ColWidth(12) = 2500
MSHFlexGrid2.ColWidth(13) = 2500
MSHFlexGrid2.ColWidth(14) = 0

MSHFlexGrid2.RowHeight(1) = 500

End Sub
Private Sub Grid_Head1()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.FixedRows = 1
MSHFlexGrid1.Cols = 31

MSHFlexGrid1.TextMatrix(0, 0) = "Claim ID"
MSHFlexGrid1.TextMatrix(0, 1) = "Claim Date"
MSHFlexGrid1.TextMatrix(0, 2) = "Expense From"
MSHFlexGrid1.TextMatrix(0, 3) = "Expense To"
MSHFlexGrid1.TextMatrix(0, 4) = "Location"
MSHFlexGrid1.TextMatrix(0, 5) = "Lessor Name"
MSHFlexGrid1.TextMatrix(0, 6) = "Rent Amount"
MSHFlexGrid1.TextMatrix(0, 7) = "Claim Forwarded To"
MSHFlexGrid1.TextMatrix(0, 8) = "Status"
MSHFlexGrid1.TextMatrix(0, 9) = "Booked in NAV Date"
MSHFlexGrid1.TextMatrix(0, 10) = "FAG Remarks"
MSHFlexGrid1.TextMatrix(0, 11) = "Remarks"
MSHFlexGrid1.TextMatrix(0, 12) = "Batch No"
MSHFlexGrid1.TextMatrix(0, 13) = "G_UID"
MSHFlexGrid1.TextMatrix(0, 14) = "Bank Name"
MSHFlexGrid1.TextMatrix(0, 15) = "Cheque No"
MSHFlexGrid1.TextMatrix(0, 16) = "Paid Amount"
MSHFlexGrid1.TextMatrix(0, 17) = "TDS "
MSHFlexGrid1.TextMatrix(0, 18) = "Paid Date"
MSHFlexGrid1.TextMatrix(0, 19) = "Disbursment Date"
MSHFlexGrid1.TextMatrix(0, 20) = "Disbursment Remarks"
MSHFlexGrid1.TextMatrix(0, 21) = "Disbursment Amount"
MSHFlexGrid1.TextMatrix(0, 22) = "Mode for Payment"
MSHFlexGrid1.TextMatrix(0, 23) = "Disbursment through Bank"
MSHFlexGrid1.TextMatrix(0, 24) = "Payment Batch ID"
MSHFlexGrid1.TextMatrix(0, 25) = "Payment Doc No"
MSHFlexGrid1.TextMatrix(0, 26) = "Payment Entry No"
MSHFlexGrid1.TextMatrix(0, 27) = "FAG Payment exported Date"
MSHFlexGrid1.TextMatrix(0, 28) = "Deduction Remark"
MSHFlexGrid1.TextMatrix(0, 29) = "Letter Print Flag"
MSHFlexGrid1.TextMatrix(0, 30) = "Billing State India"

MSHFlexGrid1.ColWidth(0) = 600
MSHFlexGrid1.ColWidth(1) = 1400
MSHFlexGrid1.ColWidth(2) = 1400
MSHFlexGrid1.ColWidth(3) = 1400
MSHFlexGrid1.ColWidth(4) = 2200
MSHFlexGrid1.ColWidth(5) = 3000
MSHFlexGrid1.ColWidth(6) = 1400
MSHFlexGrid1.ColWidth(7) = 3000
MSHFlexGrid1.ColWidth(8) = 1000
MSHFlexGrid1.ColWidth(9) = 1400
MSHFlexGrid1.ColWidth(10) = 2600
MSHFlexGrid1.ColWidth(11) = 2600
MSHFlexGrid1.ColWidth(12) = 1200
MSHFlexGrid1.ColWidth(13) = 0
MSHFlexGrid1.ColWidth(14) = 2000
MSHFlexGrid1.ColWidth(15) = 1800
MSHFlexGrid1.ColWidth(16) = 1400
MSHFlexGrid1.ColWidth(17) = 1400
MSHFlexGrid1.ColWidth(18) = 1400
MSHFlexGrid1.RowHeight(0) = 500

End Sub
Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Txn_GodownRentApproval.xls")
Call Xls_Detail_Rpt1
End Sub
Public Sub Xls_Detail_Rpt1()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Integer
Dim mCol As Integer
MsgBox ("Wait till next message")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Txn_GodownRentApproval.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = I + 1 ' + 6
    For C = 0 To MSHFlexGrid1.Cols - 2
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
Private Sub CmbForwardTo_GotFocus()

tmp = CmbForwardTo.Text
''Only AM,RM,FAG and Management
Call TableMst_Employee("Where  Flag = 'Y' and DesignationID in (1,2,3)")
CmbForwardTo.Clear
If RsMst_Employee.RecordCount = 0 Then
   MsgBox "No Employee Found!!!"
   CmbForwardTo.SetFocus
   'RsMst_Employee.MoveFirst
   Exit Sub
End If
For I = 1 To RsMst_Employee.RecordCount
    CmbForwardTo.AddItem RsMst_Employee!EmployeeName & "~" & RsMst_Employee!EmployeeNo
    RsMst_Employee.MoveNext
Next

CmbForwardTo.Text = tmp

End Sub

Private Sub CmbForwardTo_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "CmbForwardTo")
End If
End Sub

Private Sub CmbForwardTo_LostFocus()
Dim strSplitString() As String
Dim EmpNo_, EmpName_ As String

If CmbForwardTo.Text = "" Then
   mForwardToID = Null
   Exit Sub
End If

strSplitString = Split(CmbForwardTo, "~")
EmpName_ = Trim(strSplitString(0))
EmpNo_ = Trim(strSplitString(1))


RsMst_Employee.MoveFirst
RsMst_Employee.Find "EmployeeNo = '" & EmpNo_ & "'"

If RsMst_Employee.EOF Then
   MsgBox "Invalid Selection "
   CmbForwardTo.SetFocus
   Exit Sub
End If
mForwardToID = RsMst_Employee!EmployeeID
TxtFWDDesignation = GetDesignaion(RsMst_Employee!DesignationID)
mForwardToEmailID = RsMst_Employee!EmployeeEmailID

End Sub


Private Sub CmbSForwardTo_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "CmbSForwardTo")
End If
End Sub

Private Sub CmbSForwardTo_LostFocus()
Dim strSplitString() As String
Dim EmpNo_, EmpName_ As String

If CmbSForwardTo.Text = "" Then
   mSForwardToID = Null
   Exit Sub
End If

strSplitString = Split(CmbSForwardTo, "~")
EmpName_ = Trim(strSplitString(0))
EmpNo_ = Trim(strSplitString(1))

RsMst_Employee.MoveFirst
RsMst_Employee.Find "EmployeeNo = '" & EmpNo_ & "'"

If RsMst_Employee.EOF Then
   MsgBox "Invalid Selection "
   CmbSForwardTo.SetFocus
   Exit Sub
End If
mSForwardToID = RsMst_Employee!EmployeeID

End Sub


Private Sub SaveLessorArrearsDetails()

Call TableTxn_GodownRentDetail(" Where ClaimID = " & TxtClaimID & "")
For I = 2 To MSHFlexGrid2.Rows - 1
    If MSHFlexGrid2.TextMatrix(I, 1) <> "" Then
       RsTxn_GodownRentDetail.Find "ClaimDetailID = " & MSHFlexGrid2.TextMatrix(I, 0) & ""
       If RsTxn_GodownRentDetail.EOF = False Then
          RsTxn_GodownRentDetail!Arrears = Val(MSHFlexGrid2.TextMatrix(I, 7))
          RsTxn_GodownRentDetail.Update
          RsTxn_GodownRentDetail.Requery
       End If
    End If
Next
End Sub


Private Sub TxtArrears_LostFocus()

If TxtArrears = "" Then Exit Sub
If IsNumeric(TxtArrears) = False Or Val(TxtArrears) < 0 Then
   MsgBox "Invalid Format!!!"
   TxtArrears.SetFocus
   Exit Sub
End If

MSHFlexGrid2.TextMatrix(mSelRow, mSelCol) = Val(TxtArrears)

MSHFlexGrid2.TextMatrix(0, 8) = Val(MSHFlexGrid2.TextMatrix(0, 8)) - Val(MSHFlexGrid2.TextMatrix(mSelRow, 8))
MSHFlexGrid2.TextMatrix(0, 7) = Val(MSHFlexGrid2.TextMatrix(0, 7)) - Val(MSHFlexGrid2.TextMatrix(mSelRow, 7))


MSHFlexGrid2.TextMatrix(mSelRow, 8) = Val(MSHFlexGrid2.TextMatrix(mSelRow, 6)) + Val(TxtArrears)

MSHFlexGrid2.TextMatrix(0, 7) = Val(MSHFlexGrid2.TextMatrix(0, 7)) + Val(MSHFlexGrid2.TextMatrix(mSelRow, 7))

MSHFlexGrid2.TextMatrix(0, 8) = Val(MSHFlexGrid2.TextMatrix(0, 8)) + Val(MSHFlexGrid2.TextMatrix(mSelRow, 8))


TxtArrears.Visible = False
MSHFlexGrid2.SetFocus

TxtClaimAmount = Val(MSHFlexGrid2.TextMatrix(0, 8))

End Sub
Private Sub CmbSLocation_GotFocus()
tmp = CmbSLocation.Text
Call TableMst_Location(" Where LocationId > 0 ")
CmbSLocation.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No record found !!!! "
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
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbSLocation.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection "
   CmbSLocation.SetFocus
   Exit Sub
End If
mSLocationID = RsMst_Location!LocationID
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
VScroll1.Max = Me.ScaleHeight - Frame1.Height '4
VScroll1.LargeChange = Me.ScaleHeight
VScroll1.SmallChange = Me.ScaleHeight / 5
HScroll1.Max = Me.ScaleWidth - Frame1.Width '4
HScroll1.LargeChange = Me.ScaleWidth
HScroll1.SmallChange = Me.ScaleWidth / 5
If Frame1.Height >= Me.Height Then VScroll1.Enabled = True ''4
If Frame1.Height < Me.Height Then VScroll1.Enabled = False '4
If Frame1.Width >= Me.Width Then HScroll1.Enabled = True '4
If Frame1.Width < Me.Width Then HScroll1.Enabled = False '4
End Sub

Private Sub TxtSBatchID_LostFocus()
If TxtSBatchID = "" Then Exit Sub

If IsNumeric(TxtSBatchID) = False Or Val(TxtSBatchID) < 0 Then
   MsgBox "Invalid Format!!!"
   TxtSBatchID.SetFocus
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

Private Sub ShowAccountDetails(mLessorID_ As Variant)

tmp = "Select BankID ,BranchID,AccountNo,RTGSNo from Mst_Lessor Where LessorID=" & mLessorID_
Call TmpTable

If TmpRs.RecordCount > 0 Then
   TxtBankName.Text = GetBankName(IIf(IsNull(TmpRs!BankID), 0, TmpRs!BankID), "N")
   TxtBranchName = GetBranchName(IIf(IsNull(TmpRs!BranchID), 0, TmpRs!BranchID), "N")  ' IIf(IsNull(TmpRs!BankBranch), "", TmpRs!BankBranch)
   TxtACNo = IIf(IsNull(TmpRs!AccountNo), "", TmpRs!AccountNo)
   TxtIFSCCode = IIf(IsNull(TmpRs!RTGSNo), "", TmpRs!RTGSNo)
End If

End Sub

Private Sub TxtFAGDisbAmt_LostFocus()
If TxtFAGDisbAmt = "" Then Exit Sub
If IsNumeric(TxtFAGDisbAmt) = False Or Val(TxtFAGDisbAmt) < 0 Then
   MsgBox "Invalid format!!!"
   TxtFAGDisbAmt.SetFocus
   Exit Sub
End If
If Val(TxtFAGDisbAmt) > Val(TxtClaimAmount) Then
   MsgBox "Disbusment amount should be less than or equal to Claim Amount!!!"
   TxtFAGDisbAmt.SetFocus
   Exit Sub
End If
End Sub



