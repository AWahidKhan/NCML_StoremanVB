VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{A4F98396-D7CE-459C-BA85-E874A5B1E44F}#3.3#0"; "OSSMTP.ocx"
Begin VB.Form FrmTxn_ClaimSubmission 
   Caption         =   "Claim Submission"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      TabIndex        =   113
      Top             =   10320
      Width           =   10935
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   9015
      Left            =   15000
      TabIndex        =   112
      Top             =   0
      Width           =   255
   End
   Begin VB.Frame Frame4 
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   9855
      Left            =   0
      TabIndex        =   1
      Top             =   -120
      Width           =   15135
      Begin VB.Frame Frame0 
         Height          =   1935
         Left            =   840
         TabIndex        =   99
         Top             =   6480
         Visible         =   0   'False
         Width           =   13560
         Begin VB.ComboBox CmbSLocation 
            Height          =   315
            Left            =   6885
            Sorted          =   -1  'True
            TabIndex        =   34
            Top             =   345
            Width           =   2325
         End
         Begin VB.ComboBox CmbSEmployee 
            Height          =   315
            Left            =   2835
            Sorted          =   -1  'True
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   345
            Width           =   3990
         End
         Begin VB.CommandButton CmdGo 
            Caption         =   "Go"
            Height          =   350
            Left            =   12930
            TabIndex        =   39
            Top             =   750
            Width           =   1000
         End
         Begin VB.CommandButton CmdExcel 
            Caption         =   "Excel"
            Height          =   350
            Left            =   13995
            TabIndex        =   40
            Top             =   750
            Width           =   1000
         End
         Begin VB.ComboBox CmbSStatus 
            Height          =   315
            ItemData        =   "FrmTxn_ClaimSubmission.frx":0000
            Left            =   9240
            List            =   "FrmTxn_ClaimSubmission.frx":0010
            Sorted          =   -1  'True
            TabIndex        =   35
            Top             =   345
            Width           =   2115
         End
         Begin VB.TextBox TxtSClaimID 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   1440
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   675
            Width           =   1335
         End
         Begin VB.CheckBox ChkBatchNo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Check1"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3840
            TabIndex        =   38
            Top             =   720
            Width           =   225
         End
         Begin VB.ComboBox CmbSForwardTo 
            Height          =   315
            Left            =   11400
            Sorted          =   -1  'True
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   345
            Width           =   3630
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   1455
            Left            =   75
            TabIndex        =   100
            Top             =   1170
            Width           =   14115
            _ExtentX        =   24897
            _ExtentY        =   2566
            _Version        =   393216
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSComCtl2.DTPicker STo 
            Height          =   300
            Left            =   1440
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   330
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
            Format          =   112066561
            CurrentDate     =   36494
         End
         Begin MSComCtl2.DTPicker SFrom 
            Height          =   300
            Left            =   120
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   330
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
            Format          =   112066561
            CurrentDate     =   36494
         End
         Begin VB.Label Label44 
            Caption         =   "From Date"
            Height          =   255
            Left            =   360
            TabIndex        =   110
            Top             =   120
            Width           =   855
         End
         Begin VB.Label Label36 
            Caption         =   "Location"
            Height          =   300
            Left            =   7687
            TabIndex        =   109
            Top             =   120
            Width           =   720
         End
         Begin VB.Label Label37 
            Caption         =   "Employee Name"
            Height          =   225
            Left            =   4193
            TabIndex        =   108
            Top             =   150
            Width           =   1275
         End
         Begin VB.Label Label45 
            Caption         =   "To Date"
            Height          =   255
            Left            =   1680
            TabIndex        =   107
            Top             =   120
            Width           =   855
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
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
            Left            =   9315
            TabIndex        =   106
            Top             =   795
            Width           =   75
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
            Left            =   6840
            TabIndex        =   105
            Top             =   780
            Width           =   2415
         End
         Begin VB.Label Label16 
            Caption         =   "Status"
            Height          =   225
            Left            =   9937
            TabIndex        =   104
            Top             =   120
            Width           =   720
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
            Left            =   495
            TabIndex        =   103
            Top             =   727
            Width           =   900
         End
         Begin VB.Label Label30 
            Caption         =   "Batch No"
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
            Left            =   2880
            TabIndex        =   102
            Top             =   720
            Width           =   870
         End
         Begin VB.Label Label31 
            Caption         =   "Forwarded To"
            Height          =   225
            Left            =   12578
            TabIndex        =   101
            Top             =   120
            Width           =   1275
         End
      End
      Begin VB.Frame Frame1 
         Height          =   8595
         Left            =   120
         TabIndex        =   41
         Top             =   120
         Width           =   15015
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
            TabIndex        =   5
            Top             =   1080
            Width           =   10440
         End
         Begin VB.TextBox TxtClaimID 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   330
            Left            =   1245
            Locked          =   -1  'True
            TabIndex        =   77
            TabStop         =   0   'False
            Top             =   240
            Width           =   1455
         End
         Begin VB.ComboBox CmbEmployeeID 
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
            TabIndex        =   4
            Top             =   675
            Width           =   10455
         End
         Begin VB.CommandButton CmdEmployeeID 
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
            Left            =   11715
            TabIndex        =   76
            Top             =   690
            Width           =   375
         End
         Begin VB.TextBox TxtDesignation 
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
            Left            =   13365
            Locked          =   -1  'True
            TabIndex        =   75
            TabStop         =   0   'False
            Top             =   675
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
            Height          =   405
            Left            =   1245
            MaxLength       =   200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   8
            Top             =   2340
            Width           =   13560
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
            Left            =   13380
            Locked          =   -1  'True
            TabIndex        =   74
            TabStop         =   0   'False
            Top             =   255
            Visible         =   0   'False
            Width           =   1440
         End
         Begin VB.TextBox TxtPreFix 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Enabled         =   0   'False
            Height          =   360
            Left            =   13365
            Locked          =   -1  'True
            TabIndex        =   73
            TabStop         =   0   'False
            Top             =   1080
            Width           =   1470
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
            ItemData        =   "FrmTxn_ClaimSubmission.frx":0039
            Left            =   1245
            List            =   "FrmTxn_ClaimSubmission.frx":0049
            Sorted          =   -1  'True
            TabIndex        =   72
            Top             =   1500
            Width           =   3600
         End
         Begin VB.TextBox TxtClaimAmount 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   375
            Left            =   10080
            Locked          =   -1  'True
            TabIndex        =   71
            TabStop         =   0   'False
            Top             =   1500
            Width           =   1635
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
            Height          =   390
            Left            =   1245
            Locked          =   -1  'True
            MaxLength       =   200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   9
            Top             =   2775
            Width           =   10440
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
            Left            =   1245
            Sorted          =   -1  'True
            TabIndex        =   7
            Top             =   1935
            Width           =   10455
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
            Left            =   11715
            TabIndex        =   43
            Top             =   1965
            Width           =   375
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
            Left            =   13365
            Locked          =   -1  'True
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   1920
            Width           =   1455
         End
         Begin TabDlg.SSTab SSTab1 
            Height          =   5325
            Left            =   105
            TabIndex        =   44
            Top             =   3195
            Width           =   14715
            _ExtentX        =   25956
            _ExtentY        =   9393
            _Version        =   393216
            TabHeight       =   520
            TabCaption(0)   =   "Expense Details"
            TabPicture(0)   =   "FrmTxn_ClaimSubmission.frx":0072
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "MSHFlexGrid2"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "Frame2"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).ControlCount=   2
            TabCaption(1)   =   "Log History"
            TabPicture(1)   =   "FrmTxn_ClaimSubmission.frx":008E
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Label24"
            Tab(1).Control(1)=   "Label21"
            Tab(1).Control(2)=   "Label20"
            Tab(1).Control(3)=   "TxtUpdatedByFAG"
            Tab(1).Control(3).Enabled=   0   'False
            Tab(1).Control(4)=   "TxtCreated"
            Tab(1).Control(4).Enabled=   0   'False
            Tab(1).Control(5)=   "TxtUpdated"
            Tab(1).Control(5).Enabled=   0   'False
            Tab(1).ControlCount=   6
            TabCaption(2)   =   "Claim Forward History Details"
            TabPicture(2)   =   "FrmTxn_ClaimSubmission.frx":00AA
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "MSHFlexGrid3"
            Tab(2).ControlCount=   1
            Begin VB.Frame Frame2 
               Height          =   2340
               Left            =   75
               TabIndex        =   48
               Top             =   330
               Width           =   14535
               Begin VB.ComboBox CmbNonNcdexClient 
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
                  Sorted          =   -1  'True
                  TabIndex        =   17
                  Top             =   960
                  Visible         =   0   'False
                  Width           =   495
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
                  Left            =   4920
                  Locked          =   -1  'True
                  TabIndex        =   53
                  TabStop         =   0   'False
                  Top             =   1320
                  Width           =   3855
               End
               Begin VB.ComboBox CmbClientDP 
                  Height          =   315
                  Left            =   8760
                  Sorted          =   -1  'True
                  TabIndex        =   18
                  Top             =   1320
                  Visible         =   0   'False
                  Width           =   5685
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
                  Left            =   3480
                  TabIndex        =   16
                  Top             =   1320
                  Width           =   1455
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
                  Left            =   8760
                  Locked          =   -1  'True
                  TabIndex        =   52
                  TabStop         =   0   'False
                  Top             =   1320
                  Width           =   1335
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
                  Left            =   10080
                  Locked          =   -1  'True
                  TabIndex        =   51
                  TabStop         =   0   'False
                  Top             =   1320
                  Width           =   4335
               End
               Begin VB.CommandButton CmdClient 
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
                  Height          =   270
                  Left            =   4695
                  TabIndex        =   50
                  TabStop         =   0   'False
                  Top             =   1035
                  Width           =   375
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
                  ItemData        =   "FrmTxn_ClaimSubmission.frx":00C6
                  Left            =   1530
                  List            =   "FrmTxn_ClaimSubmission.frx":00C8
                  Sorted          =   -1  'True
                  TabIndex        =   15
                  Top             =   1320
                  Width           =   1935
               End
               Begin VB.CommandButton CmdExpenseTypeID 
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
                  Left            =   3885
                  TabIndex        =   49
                  TabStop         =   0   'False
                  Top             =   150
                  Width           =   375
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
                  Left            =   11280
                  TabIndex        =   20
                  Top             =   1845
                  Width           =   1455
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
                  Left            =   12735
                  TabIndex        =   21
                  Top             =   1845
                  Width           =   1710
               End
               Begin VB.TextBox TxtAmount 
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
                  Left            =   4560
                  MaxLength       =   12
                  TabIndex        =   11
                  Top             =   480
                  Width           =   1725
               End
               Begin VB.CheckBox ChkRecoverable 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  Caption         =   "Check1"
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Left            =   645
                  TabIndex        =   14
                  Top             =   1350
                  Width           =   225
               End
               Begin VB.TextBox TxtParticulars 
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
                  Left            =   6315
                  MaxLength       =   250
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   12
                  Top             =   480
                  Width           =   6270
               End
               Begin VB.ComboBox CmbExpenseTypeID 
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
                  TabIndex        =   10
                  Top             =   480
                  Width           =   4440
               End
               Begin VB.TextBox TxtRecoverableAmount 
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
                  Left            =   1530
                  MaxLength       =   12
                  TabIndex        =   19
                  Top             =   1920
                  Width           =   1725
               End
               Begin VB.TextBox TxtQty 
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
                  Left            =   12660
                  MaxLength       =   12
                  TabIndex        =   13
                  Top             =   480
                  Width           =   1710
               End
               Begin VB.Label Label8 
                  Caption         =   "Client Name "
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
                  Left            =   6180
                  TabIndex        =   65
                  Top             =   1050
                  Width           =   1095
               End
               Begin VB.Label Label9 
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
                  Left            =   3915
                  TabIndex        =   64
                  Top             =   1050
                  Width           =   840
               End
               Begin VB.Label Label10 
                  Caption         =   "DP ID"
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
                  Left            =   9135
                  TabIndex        =   63
                  Top             =   1050
                  Width           =   585
               End
               Begin VB.Label Label11 
                  Caption         =   "DP Name "
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
                  Left            =   12180
                  TabIndex        =   62
                  Top             =   1050
                  Width           =   975
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
                  Left            =   11505
                  TabIndex        =   61
                  Top             =   1050
                  Visible         =   0   'False
                  Width           =   1035
               End
               Begin VB.Label Label7 
                  Caption         =   "Client Type"
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
                  Left            =   1680
                  TabIndex        =   60
                  Top             =   1050
                  Width           =   1695
               End
               Begin VB.Label LblInvoiceAmount 
                  Caption         =   "Amount"
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
                  Left            =   4980
                  TabIndex        =   59
                  Top             =   165
                  Width           =   855
               End
               Begin VB.Label Label6 
                  Caption         =   "Recoverable from Client?"
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
                  TabIndex        =   58
                  Top             =   840
                  Width           =   1395
               End
               Begin VB.Label Label5 
                  Caption         =   "Particulars"
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
                  Left            =   9000
                  TabIndex        =   57
                  Top             =   165
                  Width           =   1200
               End
               Begin VB.Label Label3 
                  Caption         =   "Expense Type"
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
                  Left            =   1320
                  TabIndex        =   56
                  Top             =   150
                  Width           =   1530
               End
               Begin VB.Label Label25 
                  Caption         =   "Recoverable Amount "
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
                  Left            =   1440
                  TabIndex        =   55
                  Top             =   1680
                  Width           =   2190
               End
               Begin VB.Label Label32 
                  Caption         =   "Quantity"
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
                  Left            =   13140
                  TabIndex        =   54
                  Top             =   165
                  Width           =   855
               End
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
               TabIndex        =   47
               TabStop         =   0   'False
               Top             =   1635
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
               TabIndex        =   46
               TabStop         =   0   'False
               Top             =   840
               Width           =   9855
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
               TabIndex        =   45
               TabStop         =   0   'False
               Top             =   2520
               Width           =   9855
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
               Height          =   2565
               Left            =   75
               TabIndex        =   66
               Top             =   2670
               Width           =   14535
               _ExtentX        =   25638
               _ExtentY        =   4524
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
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
               Height          =   4575
               Left            =   -74880
               TabIndex        =   67
               Top             =   480
               Width           =   14475
               _ExtentX        =   25532
               _ExtentY        =   8070
               _Version        =   393216
               WordWrap        =   -1  'True
               AllowUserResizing=   3
               _NumberOfBands  =   1
               _Band(0).Cols   =   2
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
               TabIndex        =   70
               Top             =   1695
               Width           =   1695
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
               TabIndex        =   69
               Top             =   900
               Width           =   1695
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
               TabIndex        =   68
               Top             =   2580
               Width           =   2175
            End
         End
         Begin MSComCtl2.DTPicker TxtClaimDate 
            Height          =   375
            Left            =   4200
            TabIndex        =   2
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   112066561
            CurrentDate     =   39884
         End
         Begin MSComCtl2.DTPicker TxtExpenseFrom 
            Height          =   375
            Left            =   7440
            TabIndex        =   3
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Format          =   112066561
            CurrentDate     =   39884
         End
         Begin MSComCtl2.DTPicker TxtExpenseTo 
            Height          =   375
            Left            =   10335
            TabIndex        =   81
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   112066561
            CurrentDate     =   39884
         End
         Begin MSComCtl2.DTPicker TxtApprovalDate 
            Height          =   375
            Left            =   6960
            TabIndex        =   6
            Top             =   1500
            Visible         =   0   'False
            Width           =   1710
            _ExtentX        =   3016
            _ExtentY        =   661
            _Version        =   393216
            Format          =   112066561
            CurrentDate     =   39884
         End
         Begin VB.TextBox TxtEmployeeID 
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
            TabIndex        =   79
            TabStop         =   0   'False
            Top             =   675
            Width           =   10455
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
            Left            =   1245
            Locked          =   -1  'True
            TabIndex        =   78
            TabStop         =   0   'False
            Top             =   1080
            Width           =   10440
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
            Left            =   1245
            Locked          =   -1  'True
            TabIndex        =   82
            TabStop         =   0   'False
            Top             =   1935
            Width           =   10455
         End
         Begin VB.TextBox TxtStatus 
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
            Left            =   1245
            Locked          =   -1  'True
            TabIndex        =   80
            TabStop         =   0   'False
            Top             =   1500
            Width           =   3600
         End
         Begin VB.Label LblLocationID 
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
            Height          =   255
            Left            =   120
            TabIndex        =   98
            Top             =   1133
            Width           =   960
         End
         Begin VB.Label LbClaimID 
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
            Height          =   300
            Left            =   120
            TabIndex        =   97
            Top             =   300
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
            Left            =   3030
            TabIndex        =   96
            Top             =   292
            Width           =   1095
         End
         Begin VB.Label Label4 
            Caption         =   "Employee"
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
            Left            =   120
            TabIndex        =   95
            Top             =   727
            Width           =   1065
         End
         Begin VB.Label Label1 
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
            Left            =   12240
            TabIndex        =   94
            Top             =   705
            Width           =   1065
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
            Left            =   120
            TabIndex        =   93
            Top             =   2407
            Width           =   1800
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
            Left            =   12540
            TabIndex        =   92
            Top             =   285
            Visible         =   0   'False
            Width           =   630
         End
         Begin VB.Label Label13 
            Caption         =   "Prefix"
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
            Left            =   12390
            TabIndex        =   91
            Top             =   1140
            Width           =   720
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
            Height          =   300
            Left            =   135
            TabIndex        =   90
            Top             =   1530
            Width           =   960
         End
         Begin VB.Label Label15 
            Caption         =   "Claim Amount"
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
            Left            =   8760
            TabIndex        =   89
            Top             =   1537
            Width           =   1260
         End
         Begin VB.Label Label17 
            Caption         =   "Expense From "
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
            Left            =   6120
            TabIndex        =   88
            Top             =   285
            Width           =   1335
         End
         Begin VB.Label Label18 
            Caption         =   "Expense To"
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
            Left            =   9195
            TabIndex        =   87
            Top             =   292
            Width           =   1095
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
            Height          =   465
            Left            =   120
            TabIndex        =   86
            Top             =   2715
            Width           =   1110
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
            Left            =   5040
            TabIndex        =   85
            Top             =   1560
            Visible         =   0   'False
            Width           =   1980
         End
         Begin VB.Label Label28 
            Caption         =   "Forward To"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   90
            TabIndex        =   84
            Top             =   1980
            Width           =   1230
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
            Left            =   12240
            TabIndex        =   83
            Top             =   1950
            Width           =   1065
         End
      End
      Begin VB.Frame ButtonFrm 
         Height          =   1080
         Left            =   120
         TabIndex        =   111
         Top             =   8655
         Width           =   15015
         Begin VB.CommandButton ExitCmd 
            Caption         =   "E&xit"
            Height          =   400
            Left            =   12105
            TabIndex        =   26
            Top             =   600
            Width           =   2700
         End
         Begin VB.CommandButton SearchCmd 
            Caption         =   "Search"
            Height          =   400
            Left            =   12105
            TabIndex        =   25
            Top             =   195
            Width           =   2700
         End
         Begin VB.CommandButton LastCmd 
            Caption         =   "&Last"
            Height          =   400
            Left            =   9165
            TabIndex        =   27
            Top             =   600
            Width           =   2940
         End
         Begin VB.CommandButton DeleteCmd 
            Caption         =   "&Delete"
            Height          =   400
            Left            =   9165
            TabIndex        =   24
            Top             =   195
            Width           =   2940
         End
         Begin VB.CommandButton FirstCmd 
            Caption         =   "&First"
            Height          =   400
            Left            =   6225
            TabIndex        =   28
            Top             =   600
            Width           =   2940
         End
         Begin VB.CommandButton EditCmd 
            Caption         =   "&Edit"
            Height          =   400
            Left            =   6225
            TabIndex        =   23
            Top             =   195
            Width           =   2940
         End
         Begin VB.CommandButton PreviousCmd 
            Caption         =   "&Previous"
            Height          =   400
            Left            =   3165
            TabIndex        =   29
            Top             =   600
            Width           =   3060
         End
         Begin VB.CommandButton SaveCmd 
            Caption         =   "Save"
            Height          =   400
            Left            =   3165
            TabIndex        =   22
            Top             =   195
            Width           =   3060
         End
         Begin VB.CommandButton NextCmd 
            Caption         =   "&Next"
            Height          =   400
            Left            =   105
            TabIndex        =   30
            Top             =   600
            Width           =   3060
         End
         Begin VB.CommandButton AddCmd 
            Caption         =   "&Add New"
            Height          =   400
            Left            =   105
            TabIndex        =   0
            Top             =   195
            Width           =   3060
         End
      End
   End
   Begin OSSMTPControl.SMTP SMTP 
      Left            =   120
      Top             =   10080
      _ExtentX        =   1296
      _ExtentY        =   873
      AuthenticationType=   0
      BCC             =   ""
      CC              =   ""
      Charset         =   "us-ascii"
      ClientHostName  =   ""
      ContentTransferEncoding=   "7bit"
      ContentType     =   "text/plain"
      ExpiresAfter    =   ""
      Importance      =   0
      LastSMTPCommand =   ""
      MailFrom        =   ""
      MessageHTML     =   ""
      MessageSubject  =   ""
      MessageText     =   ""
      Notification    =   0
      Password        =   ""
      POPServer       =   ""
      Port            =   25
      RaiseError      =   0   'False
      ReplyTo         =   ""
      SendTo          =   ""
      Sensitivity     =   0
      Server          =   ""
      Status          =   ""
      Timeout         =   10000
      TimeStamp       =   ""
      Username        =   ""
   End
End
Attribute VB_Name = "FrmTxn_ClaimSubmission"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg, mSM_Prefix, mAgentID, mExpenseTypeID, mEmployeeID, mSEmployeeID, mStatus, mSStatus, mEmployeeEmailID As Variant
Dim mLocationID, mClientIDRow, mSLocationID As Double
Dim mExchangeTypeID As Double
Dim mBatchID As Double
Dim mForwardToID, mForwardToEmailID, mSForwardToID, mEmpDesignationID As Variant
Dim mMonth, mYear, mTotDays As Variant
'Dim oSMTP As SMTPSession

Private Sub Grid_Head()

MSHFlexGrid2.Clear
MSHFlexGrid2.Rows = 3
MSHFlexGrid2.FixedRows = 2
MSHFlexGrid2.Cols = 14


MSHFlexGrid2.TextMatrix(1, 0) = "Row ID"
MSHFlexGrid2.TextMatrix(1, 1) = "Expense Type"
MSHFlexGrid2.TextMatrix(1, 2) = "Particulars"
MSHFlexGrid2.TextMatrix(1, 3) = "Amount"
MSHFlexGrid2.TextMatrix(1, 4) = "Quantity"
MSHFlexGrid2.TextMatrix(1, 5) = "Recoverable"
MSHFlexGrid2.TextMatrix(1, 6) = "Exchange"
MSHFlexGrid2.TextMatrix(1, 7) = "Client ID"
MSHFlexGrid2.TextMatrix(1, 8) = "Client Name"
MSHFlexGrid2.TextMatrix(1, 9) = "DP ID"
MSHFlexGrid2.TextMatrix(1, 10) = "DP Name"
MSHFlexGrid2.TextMatrix(1, 11) = "ClientIDRow"
MSHFlexGrid2.TextMatrix(1, 12) = "ExpenseTypeID"
MSHFlexGrid2.TextMatrix(1, 13) = "Recoverable Amount"


MSHFlexGrid2.ColWidth(0) = 500
MSHFlexGrid2.ColWidth(1) = 3000
MSHFlexGrid2.ColWidth(2) = 6500
MSHFlexGrid2.ColWidth(3) = 800
MSHFlexGrid2.ColWidth(4) = 800
MSHFlexGrid2.ColWidth(5) = 800
MSHFlexGrid2.ColWidth(6) = 800
MSHFlexGrid2.ColWidth(7) = 1000
MSHFlexGrid2.ColWidth(8) = 2000
MSHFlexGrid2.ColWidth(9) = 1000
MSHFlexGrid2.ColWidth(10) = 2000
MSHFlexGrid2.ColWidth(11) = 0
MSHFlexGrid2.ColWidth(12) = 0
MSHFlexGrid2.ColWidth(13) = 1600

MSHFlexGrid2.RowHeight(1) = 650

End Sub

Private Sub Grid_Head2()

MSHFlexGrid3.Clear
MSHFlexGrid3.Rows = 3
MSHFlexGrid3.FixedRows = 1
MSHFlexGrid3.Cols = 5


MSHFlexGrid3.TextMatrix(0, 0) = "Cliam Row ID"
MSHFlexGrid3.TextMatrix(0, 1) = "Old Employee to Whom Claim was Forwarded"
MSHFlexGrid3.TextMatrix(0, 2) = "New Employee to Whom Claim is Forwarded"
MSHFlexGrid3.TextMatrix(0, 3) = "Updated By Employee"
MSHFlexGrid3.TextMatrix(0, 4) = "Concurrency"



MSHFlexGrid3.ColWidth(0) = 0 '500
MSHFlexGrid3.ColWidth(1) = 3800
MSHFlexGrid3.ColWidth(2) = 3800
MSHFlexGrid3.ColWidth(3) = 2000
MSHFlexGrid3.ColWidth(4) = 2500

MSHFlexGrid2.RowHeight(1) = 650

End Sub

Private Sub AddCmd_Click()

Edit_Flg = "A"
Call GButtonLock(Me, False)

Call Grid_Head
Call Grid_Head2
TxtEmployeeID.Visible = False
CmbEmployeeID.Visible = True
CmbEmployeeID.Text = ""
CmdEmployeeID.Visible = True

TxtLocationID.Visible = False
CmbLocationID.Visible = True
CmbLocationID.Text = ""
'TxtClaimDate.Enabled = True


TxtForwardTo.Visible = False
CmbForwardTo.Visible = True
CmbForwardTo.Text = ""
CmdForwardTo.Visible = True


TxtRemark.Locked = False
TxtClaimID = ""
TxtClaimDate.Value = Date


TxtExpenseFrom.Enabled = True
'TxtExpenseTo.Enabled = True

CmbLocationID.Enabled = True

TxtStatus.Visible = True
CmbStatus.Visible = False
TxtStatus = "Pending"
CmbStatus = ""
mStatus = "P"

TxtDesignation = ""
TxtPreFix = ""
TxtClaimAmount = ""
TxtRemark = ""

mEmployeeID = Null

tmp = "Select EmployeeID,EmployeeName +'~'+ EmployeeNo as Employee,DesignationID,EmployeeEmailID From Mst_Employee Where EmployeeNo = '" & Gen_UserLoginID & "'"
Call TmpTable

If TmpRs.RecordCount = 1 Then
   CmbEmployeeID.Text = TmpRs!Employee
   TxtDesignation = GetDesignaion(TmpRs!DesignationID)
   mEmployeeID = TmpRs!EmployeeID
   mEmployeeEmailID = TmpRs!EmployeeEmailID
   mEmpDesignationID = TmpRs!DesignationID
Else
   CmbEmployeeID.Text = ""
   TxtDesignation = ""
   mEmployeeID = Null
   mEmpDesignationID = Null
End If

mBatchID = 0
TxtCreated = ""
TxtUpdated = ""
TxtUpdatedByFAG = ""
TxtApprovalRemarks = ""

Call ClearFrame
Frame2.Enabled = True

ChkRecoverable.Enabled = True

If RsTxn_ClaimSubmission.RecordCount > 0 Then
   TxtExpenseFrom.SetFocus
End If
'
'TxtExpenseFrom.Value = Date
'TxtExpenseTo.Value = Date
mMonth = Month(Date)
mYear = Year(Date)
TxtExpenseFrom.Value = mYear & "-" & mMonth & "- 01"
mTotDays = Day(DateSerial(mYear, mMonth + 1, 0))
TxtExpenseTo.Value = mYear & "-" & mMonth & "-" & mTotDays

SSTab1.Tab = 0

End Sub

Private Sub ClearFrame()

TxtQty.Text = ""
CmbExpenseTypeID.Text = ""
'CmbExchangeTypeID.Text = ""
'TxtClientID = ""
'TxtClientName = ""
'TxtDPID = ""
'TxtDPName = ""
'CmbNonNcdexClient.Text = ""
'CmbClientDP.Text = ""
TxtParticulars = ""
TxtAmount = ""
'ChkSupporting.Value = 1
ChkRecoverable.Value = 0
Call ChkRecoverable_Click
CmdAdd2List.Caption = "Add to List"

End Sub

Private Sub ChkRecoverable_Click()
If ChkRecoverable.Value = 0 Then
   CmbExchangeTypeID.Text = ""
   mExchangeTypeID = 0
   TxtClientID = ""
   TxtClientName = ""
   TxtDPID = ""
   TxtDPName = ""
   mClientIDRow = 0
   CmbClientDP.Text = ""
   TxtRecoverableAmount = ""
   CmbExchangeTypeID.Enabled = False
   CmdClient.Enabled = Flase
   TxtClientID.Enabled = False
   TxtRecoverableAmount.Locked = True
   If CmbNonNcdexClient.Visible = True Then
      Label9.Visible = True
      Label10.Visible = True
      Label11.Visible = True
      CmbNonNcdexClient.Visible = False
   End If
Else
   CmbExchangeTypeID.Enabled = True
   TxtClientID.Enabled = True
   TxtRecoverableAmount.Locked = False
   CmdClient.Enabled = True
End If

End Sub
Private Sub CmbSEmployee_GotFocus()
tmp = CmbSEmployee.Text
CmbSEmployee.Clear
Call GetSearchEmployeeDataonGotFocus
CmbSEmployee.Text = tmp
End Sub
Private Sub GetSearchEmployeeDataonGotFocus()

Call TableMst_Employee '("Where Flag = 'Y'")
CmbEmployeeID.Clear
If RsMst_Employee.RecordCount = 0 Then
   MsgBox "No Employee Found!!!"
   CmbSEmployee.SetFocus
   'RsMst_Employee.MoveFirst
   Exit Sub
End If
For I = 1 To RsMst_Employee.RecordCount
    CmbSEmployee.AddItem RsMst_Employee!EmployeeName & "~" & RsMst_Employee!EmployeeNo
    RsMst_Employee.MoveNext
Next

End Sub

Private Sub CmbSEmployee_LostFocus()

Call GetSearchEmployeeDataOnLostFocus

End Sub


Private Sub CmbSForwardTo_GotFocus()
tmp = CmbSForwardTo.Text
Call TableMst_Employee '("Where Flag = 'Y'")
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

Private Sub CmbSStatus_LostFocus()
If CmbSStatus.Text = "" Then Exit Sub

If LCase(CmbSStatus.Text) = "approved" Then
   mSStatus = "A"
ElseIf LCase(CmbSStatus.Text) = "pending" Then
   mSStatus = "P"
ElseIf LCase(CmbSStatus.Text) = "rejected" Then
   mSStatus = "R"
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

If LCase(CmbStatus.Text) = "approved" Then
   mStatus = "A"
   TxtApprovalRemarks.Locked = False
   TxtApprovalDate.Visible = True
   Label27.Visible = True
   TxtApprovalDate.Enabled = True

ElseIf LCase(CmbStatus.Text) = "pending" Then
   mStatus = "P"
   TxtApprovalRemarks.Locked = True
   TxtApprovalDate.Visible = False
   Label27.Visible = False
ElseIf LCase(CmbStatus.Text) = "rejected" Then
   mStatus = "R"
   TxtApprovalRemarks.Locked = False
   TxtApprovalDate.Visible = True
   Label27.Visible = True
   TxtApprovalDate.Enabled = True
ElseIf LCase(CmbStatus.Text) = "frozen" Then
   mStatus = "F"
   TxtApprovalRemarks.Locked = False
   TxtApprovalDate.Visible = True
   Label27.Visible = True
   TxtApprovalDate.Enabled = True
Else
   MsgBox "Invalid selection!!!"
   CmbStatus.SetFocus
   TxtApprovalRemarks.Locked = True
   TxtApprovalDate.Visible = False
   Label27.Visible = False
   Exit Sub
End If

If LCase(Gen_UserRole) <> "fag" And LCase(Gen_UserRole) <> "fagadmin" And LCase(Gen_UserRole) <> "power" Then
   If mStatus = "F" Or mStatus = "A" Then
      MsgBox "Please forward this claim to FAG user, Only FAG user have access to approve claim '!!!"
      CmbStatus.SetFocus
      Exit Sub
   End If
End If

End Sub

Private Sub CmdAdd2List_Click()
Dim r As Variant

If CmbExpenseTypeID.Text = "" Then
   MsgBox "Blank Expense Type not allowed!!!"
   CmbExpenseTypeID.SetFocus
   Exit Sub
End If

If LCase(CmbExpenseTypeID.Text) = "assets to be capitalised" Then
   If TxtQty.Text = "" Then
      MsgBox "Blank Quantity not allowed for 'Assets to be capitalised'."
      TxtQty.SetFocus
      Exit Sub
   End If
End If

If ChkRecoverable.Value = 1 Then
   If CmbExchangeTypeID.Text = "" Then
      MsgBox "Blank Exchange Type not allowed!!!"
      CmbExchangeTypeID.SetFocus
      Exit Sub
   End If
   If CmbClientDP.Visible = True Then
      MsgBox "Invalid selection !!! "
      Exit Sub
   End If
   If CmbNonNcdexClient.Visible = True Then
      If CmbNonNcdexClient.Text = "" Then
         MsgBox "Invalid selection !!! "
         CmbNonNcdexClient.SetFocus
         Exit Sub
      End If
   End If
   If TxtClientID = "" Then
      MsgBox "Invalid Client ID !!! "
      Exit Sub
   End If
   If TxtRecoverableAmount = "" Then
      MsgBox "Recoverable Amount can't be blank!!!"
      TxtRecoverableAmount.SetFocus
      Exit Sub
   End If
End If

If Val(TxtAmount) = 0 Then
   MsgBox "Amount should be more than '0'!!!"
   TxtAmount.SetFocus
   Exit Sub
End If


'If ChkSupporting.Value = 1 Then
   If TxtParticulars = "" Then
      MsgBox "Blank particulas not allowed!!!"
      TxtParticulars.SetFocus
      Exit Sub
   End If
'End If

'If ChkDuplicateExpenseType = False Then
'   MsgBox "Duplicate expense type not allowed"
'   CmbExpenseTypeID.SetFocus
'   Exit Sub
'End If

If CmdAdd2List.Caption = "Update List" Then
   r = MSHFlexGrid2.RowSel
   MSHFlexGrid2.TextMatrix(0, 3) = Val(MSHFlexGrid2.TextMatrix(0, 3)) - Val(MSHFlexGrid2.TextMatrix(r, 3))
   TxtClaimAmount = Val(TxtClaimAmount) - Val(MSHFlexGrid2.TextMatrix(r, 3))
Else
   r = MSHFlexGrid2.Rows - 1
   MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
   MSHFlexGrid2.TextMatrix(r, 0) = ""
End If


MSHFlexGrid2.TextMatrix(r, 1) = CmbExpenseTypeID.Text
'If ChkSupporting.Value = 1 Then
MSHFlexGrid2.TextMatrix(r, 2) = TxtParticulars
'Else
'   MSHFlexGrid2.TextMatrix(r, 2) = "" 'IIf(IsNull(TxtParticulars), "", TxtParticulars)
'End If
'MSHFlexGrid2.TextMatrix(r, 3) = "" 'IIf(ChkSupporting.Value = 0, "N", "Y")
MSHFlexGrid2.TextMatrix(r, 3) = Val(TxtAmount)
MSHFlexGrid2.TextMatrix(r, 4) = TxtQty.Text
MSHFlexGrid2.TextMatrix(r, 5) = IIf(ChkRecoverable.Value = 0, "N", "Y")
If ChkRecoverable.Value = 1 Then
   MSHFlexGrid2.TextMatrix(r, 6) = CmbExchangeTypeID.Text
   MSHFlexGrid2.TextMatrix(r, 7) = TxtClientID
   MSHFlexGrid2.TextMatrix(r, 8) = TxtClientName
   MSHFlexGrid2.TextMatrix(r, 9) = TxtDPID
   MSHFlexGrid2.TextMatrix(r, 10) = TxtDPName
   MSHFlexGrid2.TextMatrix(r, 11) = mClientIDRow
   MSHFlexGrid2.TextMatrix(r, 13) = Val(TxtRecoverableAmount)
Else
   MSHFlexGrid2.TextMatrix(r, 6) = "" 'CmbExchangeTypeID.Text
   MSHFlexGrid2.TextMatrix(r, 7) = "" 'TxtClientID
   MSHFlexGrid2.TextMatrix(r, 8) = "" 'TxtClientName
   MSHFlexGrid2.TextMatrix(r, 9) = "" 'TxtDPID
   MSHFlexGrid2.TextMatrix(r, 10) = "" 'TxtDPName
   MSHFlexGrid2.TextMatrix(r, 11) = "" 'mClientIDRow
   MSHFlexGrid2.TextMatrix(r, 13) = "" 'Val(TxtRecoverableAmount)
End If
MSHFlexGrid2.TextMatrix(r, 12) = mExpenseTypeID
 
MSHFlexGrid2.TextMatrix(0, 3) = Val(MSHFlexGrid2.TextMatrix(0, 3)) + Val(MSHFlexGrid2.TextMatrix(r, 3))
MSHFlexGrid2.TextMatrix(0, 13) = Val(MSHFlexGrid2.TextMatrix(0, 13)) + Val(MSHFlexGrid2.TextMatrix(r, 13))
TxtClaimAmount = Val(TxtClaimAmount) + Val(MSHFlexGrid2.TextMatrix(r, 3))
Call ClearFrame
CmbExpenseTypeID.SetFocus

End Sub

Private Sub CmdClient_Click()
FrmMst_Client.Show
End Sub

Private Sub CmdDeleteList_Click()
I = MSHFlexGrid2.RowSel
If I > 0 Then
   If MSHFlexGrid2.TextMatrix(I, 1) <> "" Then
      TxtClaimAmount = Val(TxtClaimAmount) - Val(MSHFlexGrid2.TextMatrix(I, 3))
      MSHFlexGrid2.TextMatrix(0, 3) = Val(MSHFlexGrid2.TextMatrix(0, 3)) - Val(MSHFlexGrid2.TextMatrix(I, 3))
      MSHFlexGrid2.TextMatrix(0, 12) = Val(MSHFlexGrid2.TextMatrix(0, 12)) - Val(MSHFlexGrid2.TextMatrix(I, 12))
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

Private Sub CmdEmployeeID_Click()
FrmMst_Employee.Show
End Sub

Private Sub CmdExpenseTypeID_Click()
FrmMst_ExpenseType.Show
End Sub

Private Sub CmdForwardTo_Click()
FrmMst_Employee.Show
End Sub

Private Sub CmdGo_Click()
Dim strSplitString() As String
Dim EmpNo_, EmpName_ As String
Dim wc As Variant

If SFrom.Value > STo.Value Then
   MsgBox "'From Date' must be less than 'To date'!!!"
   SFrom.SetFocus
   Exit Sub
End If



wc = ""
RsTxn_ClaimSubmission.Filter = ""

tmp = "Select CS.ClaimRowID,CS.SM_Prefix,CS.ClaimID,CS.ClaimDate,CS.ExpenseFrom,CS.ExpenseTo,EMP.EmployeeName +'~'+ EMP.EmployeeNo as Employee,MD.Designation,ML.LocationName,CS.Amount, " & _
      " CS.Remarks ,CS.Flag, CS.G_UID,CS.ApprovalRemarks,CS.ApprovalDate,EMP1.EmployeeName +'~'+ EMP1.EmployeeNo as ForwardedTo,CS.BatchID " & _
      " From Txn_ClaimSubmission CS " & _
      " Inner Join Mst_Employee EMP On CS.EmployeeID = EMP.EmployeeID " & _
      " Left Outer Join Mst_Employee EMP1 On CS.ForwardToID = EMP1.EmployeeID " & _
      " Left Join Mst_Designation MD On CS.DesignationID = MD.DesignationID " & _
      " Left Join Mst_Location ML On CS.LocationID = ML.LocationID "

'tmp = tmp & " Where CS.ClaimDate between '2009-08-20 00:00:00' and '2009-08-20 00:00:00'"

If Gen_DBType = "MDB" Then
   tmp = tmp & " Where ClaimDate Between #" & Format(SFrom, Gen_DatFormat) & "# And #" & Format(STo, Gen_DatFormat) & "#) "
   wc = " ClaimDate >= #" & Format(SFrom, Gen_DatFormat) & "# And ClaimDate <= #" & Format(STo, Gen_DatFormat) & "#) "
Else
   tmp = tmp & " Where ClaimDate Between '" & Format(SFrom, Gen_DatFormat) & "' And '" & Format(STo, Gen_DatFormat) & "' "
   wc = " ClaimDate >= '" & Format(SFrom, Gen_DatFormat) & "' And ClaimDate <= '" & Format(STo, Gen_DatFormat) & "' "
End If

If CmbSLocation.Text <> "" Then
   tmp = tmp & " AND ML.LocationName = '" & CmbSLocation.Text & "' "
   wc = wc & " And LocationID = " & mSLocationID & ""
End If

If CmbSEmployee.Text <> "" Then
   strSplitString = Split(CmbSEmployee, "~")
   EmpName_ = Trim(strSplitString(0))
   EmpNo_ = Trim(strSplitString(1))
   
   tmp = tmp & " AND EMP.EmployeeNo = '" & EmpNo_ & "' "
   wc = wc & " And EmployeeID = " & mSEmployeeID & ""
End If

If CmbSStatus.Text <> "" Then
   tmp = tmp & " AND CS.Flag = '" & mSStatus & "' "
   wc = wc & " And Flag = '" & mSStatus & "'"
End If

If CmbSForwardTo.Text <> "" Then
   strSplitString = Split(CmbSForwardTo, "~")
   EmpName_ = Trim(strSplitString(0))
   EmpNo_ = Trim(strSplitString(1))
 
   tmp = tmp & " AND EMP1.EmployeeNo = '" & EmpNo_ & "' "
   wc = wc & " And ForwardToID = " & mSForwardToID & ""
End If

If TxtSClaimID <> "" Then
   tmp = tmp & " AND CS.ClaimID Like " & TxtSClaimID & " "
   wc = wc & " And ClaimID = " & TxtSClaimID & ""
End If

If ChkBatchNo.Value = 0 Then
   tmp = tmp & " AND CS.BatchID = 0 "
   wc = wc & " And BatchID = 0 "
End If

Call TmpTable

Call Grid_Head1
RsTxn_ClaimSubmission.Filter = "" & wc & ""

For iu = 1 To TmpRs.RecordCount
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
    For C = 0 To MSHFlexGrid1.Cols - 1
        MSHFlexGrid1.TextMatrix(iu, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
        If InStr(MSHFlexGrid1.TextMatrix(iu, C), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(iu, C)) Then
              MSHFlexGrid1.TextMatrix(iu, C) = Format(MSHFlexGrid1.TextMatrix(iu, C), "dd-mmm-yyyy")
           Else
              MSHFlexGrid1.TextMatrix(iu, C) = MSHFlexGrid1.TextMatrix(iu, C)
           End If
        Else
          MSHFlexGrid1.TextMatrix(iu, C) = MSHFlexGrid1.TextMatrix(iu, C)
        End If
    Next
    
    If LCase(MSHFlexGrid1.TextMatrix(iu, 11)) = "a" Then
       MSHFlexGrid1.TextMatrix(iu, 11) = "Approved"
    ElseIf LCase(MSHFlexGrid1.TextMatrix(iu, 11)) = "p" Then
       MSHFlexGrid1.TextMatrix(iu, 11) = "Pending"
    ElseIf LCase(MSHFlexGrid1.TextMatrix(iu, 11)) = "r" Then
       MSHFlexGrid1.TextMatrix(iu, 11) = "Rejected"
    ElseIf LCase(MSHFlexGrid1.TextMatrix(iu, 11)) = "f" Then
       MSHFlexGrid1.TextMatrix(iu, 11) = "Frozen"
    End If
    TmpRs.MoveNext
Next


Label22.Refresh
Label22.Caption = TmpRs.RecordCount


End Sub



Private Sub EditCmd_Click()
If LCase(TxtStatus) = "approved" Or LCase(TxtStatus) = "rejected" Or LCase(TxtStatus) = "frozen" Then
   MsgBox "This Claim is " & TxtStatus & ". You can't edit this claim any more!!!"
   Exit Sub
End If

Edit_Flg = "E"
Call GButtonLock(Me, False)
'TxtClaimDate.Enabled = True
TxtEmployeeID.Visible = False
CmbEmployeeID.Visible = True
CmbEmployeeID.Text = TxtEmployeeID
CmdEmployeeID.Visible = True

TxtForwardTo.Visible = False
CmbForwardTo.Visible = True
CmbForwardTo.Text = TxtForwardTo
CmdForwardTo.Visible = True


TxtLocationID.Visible = False
CmbLocationID.Visible = True
CmbLocationID.Text = TxtLocationID

'If LCase(Gen_UserRole) = "fag" Or LCase(Gen_UserRole) = "power" Then
'   TxtStatus.Visible = False
'   CmbStatus.Visible = True
'   CmbStatus.Text = TxtStatus
'End If
TxtStatus.Visible = False
CmbStatus.Visible = True
CmbStatus.Text = TxtStatus

TxtApprovalDate.Value = Date
TxtExpenseFrom.Enabled = True
'TxtExpenseTo.Enabled = True
TxtRemark.Locked = False
CmbLocationID.Enabled = False
Call ClearFrame
Frame2.Enabled = True
ChkRecoverable.Enabled = True

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
Call TableTxn_ClaimSubmission(" Order by ClaimRowID")
If RsTxn_ClaimSubmission.RecordCount = 0 Then
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
Call TableMst_Location("")
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
   TxtPreFix = ""
   mSM_Prefix = Null
   Exit Sub
End If
' Code for clearing Godown combox in case of change
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocationID.Text & "'"

If RsMst_Location.EOF Then
   MsgBox "Invalid Location  selection "
   CmbLocationID.SetFocus
   Exit Sub
End If
mLocationID = RsMst_Location!LocationID
mSM_Prefix = RsMst_Location!SM_Prefix
TxtPreFix = RsMst_Location!SM_Prefix
End Sub


Private Sub GetEmployeeinList()

tmp = CmbEmployeeID.Text
Call TableMst_Employee("Where Flag = 'Y'")
CmbEmployeeID.Clear
If RsMst_Employee.RecordCount = 0 Then
   MsgBox "No Employee Found!!!"
   CmbEmployeeID.SetFocus
   'RsMst_Employee.MoveFirst
   Exit Sub
End If
For I = 1 To RsMst_Employee.RecordCount
    CmbEmployeeID.AddItem RsMst_Employee!EmployeeName & "~" & RsMst_Employee!EmployeeNo
    RsMst_Employee.MoveNext
Next

CmbEmployeeID.Text = tmp


End Sub

Private Sub CmbEmployeeID_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "EmployeeID")
End If
End Sub
Private Sub GetEmployeeDataOnLostFocus()
Dim strSplitString() As String
Dim EmpNo_, EmpName_ As String

If CmbEmployeeID.Text = "" Then
   mEmployeeID = Null
   Exit Sub
End If

strSplitString = Split(CmbEmployeeID, "~")
EmpName_ = Trim(strSplitString(0))
EmpNo_ = Trim(strSplitString(1))


RsMst_Employee.MoveFirst
RsMst_Employee.Find "EmployeeNo = '" & EmpNo_ & "'"

If RsMst_Employee.EOF Then
   MsgBox "Invalid Selection "
   CmbEmployeeID.SetFocus
   Exit Sub
End If
mEmployeeID = RsMst_Employee!EmployeeID
TxtDesignation = GetDesignaion(RsMst_Employee!DesignationID)
mEmployeeEmailID = RsMst_Employee!EmployeeEmailID
mEmpDesignationID = RsMst_Employee!DesignationID
End Sub
Private Sub CmbEmployeeID_LostFocus()
Call GetEmployeeDataOnLostFocus
End Sub
Private Sub CmbExpenseTypeID_GotFocus()
Dim wc As String
tmp = CmbExpenseTypeID.Text
Call TableMst_ExpenseType(" Where (AllocateTo = 'B' Or AllocateTo = 'E') And Flag = 'A' And OtherDeductionFlag = 'N' ")
CmbExpenseTypeID.Clear
If RsMst_ExpenseType.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_ExpenseType.RecordCount
    CmbExpenseTypeID.AddItem RsMst_ExpenseType!ExpenseType
    RsMst_ExpenseType.MoveNext
Next
CmbExpenseTypeID.Text = tmp
End Sub
'----
Private Sub CmbExpenseTypeID_LostFocus()
If CmbExpenseTypeID.Text = "" Then
   mExpenseTypeID = Null
   Exit Sub
End If
RsMst_ExpenseType.MoveFirst
RsMst_ExpenseType.Find "ExpenseType = '" & CmbExpenseTypeID.Text & "'"
If RsMst_ExpenseType.EOF Then
   MsgBox "Invalid selection "
   CmbExpenseTypeID.SetFocus
   Exit Sub
End If
mExpenseTypeID = RsMst_ExpenseType!ExpenseTypeID
End Sub

Private Sub Form_Resize()
If Me.WindowState = vbMaximized Then
   SetScrollBars
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Gen_ClaimApplicant = ""
End Sub

Private Sub MSHFlexGrid1_Click()
Dim I As Variant
I = MSHFlexGrid1.RowSel
If I <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(I, 0) = "" Then Exit Sub
RsTxn_ClaimSubmission.MoveFirst
RsTxn_ClaimSubmission.Find "G_UID = '" & MSHFlexGrid1.TextMatrix(I, 12) & "'"
If Not RsTxn_ClaimSubmission.EOF Then
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

CmbExpenseTypeID.Text = MSHFlexGrid2.TextMatrix(r, 1)
mExpenseTypeID = MSHFlexGrid2.TextMatrix(r, 11)

'If MSHFlexGrid2.TextMatrix(r, 3) = "Y" Then
'   ChkSupporting.Value = 1
TxtParticulars = MSHFlexGrid2.TextMatrix(r, 2)
'ElseIf MSHFlexGrid2.TextMatrix(r, 3) = "N" Then
'   ChkSupporting.Value = 0
'   TxtParticulars = "" 'MSHFlexGrid2.TextMatrix(i, 2)
'End If

TxtAmount = Val(MSHFlexGrid2.TextMatrix(r, 3))
TxtQty.Text = MSHFlexGrid2.TextMatrix(r, 4)
If MSHFlexGrid2.TextMatrix(r, 5) = "Y" Then
   ChkRecoverable.Value = 1
   CmbExchangeTypeID.Text = MSHFlexGrid2.TextMatrix(r, 6)
   mExchangeTypeID = GetExchangeID(CmbExchangeTypeID.Text)
   TxtClientID = MSHFlexGrid2.TextMatrix(r, 7)
   TxtClientName = MSHFlexGrid2.TextMatrix(r, 8)
   TxtDPID = MSHFlexGrid2.TextMatrix(r, 9)
   TxtDPName = MSHFlexGrid2.TextMatrix(r, 10)
   If mExchangeTypeID = 1 Then
      CmbNonNcdexClient.Width = 10095
      CmbNonNcdexClient.Top = TxtClientID.Top
      CmbNonNcdexClient.Left = TxtClientID.Left
      Label9.Visible = False
      Label10.Visible = False
      Label11.Visible = False
      CmbNonNcdexClient.Visible = True
   Else
      Label9.Visible = True
      Label10.Visible = True
      Label11.Visible = True
      CmbNonNcdexClient.Visible = False
   End If
   
   If CmbNonNcdexClient.Visible = True Then
      CmbNonNcdexClient.Text = MSHFlexGrid2.TextMatrix(r, 8)
   End If
   mClientIDRow = MSHFlexGrid2.TextMatrix(r, 11)
   TxtRecoverableAmount = Val(MSHFlexGrid2.TextMatrix(r, 13))
ElseIf MSHFlexGrid2.TextMatrix(r, 5) = "N" Then
   ChkRecoverable.Value = 0
'   TxtParticulars = "" 'MSHFlexGrid2.TextMatrix(i, 2)
End If

CmdAdd2List.Caption = "Update List"
CmbExpenseTypeID.SetFocus

End Sub
Private Sub CmbExchangeTypeID_GotFocus()
tmp = CmbExchangeTypeID.Text
Call TableMst_ExchangeType
CmbExchangeTypeID.Clear
If RsMst_ExchangeType.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
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
   MsgBox "Invalid selection "
   CmbExchangeTypeID.SetFocus
   Exit Sub
End If

If mExchangeTypeID <> RsMst_ExchangeType!ExchangeTypeID Then
   TxtClientID = ""
   TxtClientName = ""
   TxtDPID = ""
   TxtDPName = ""
   CmbNonNcdexClient.Text = ""
   CmbClientDP.Text = ""
End If
mExchangeTypeID = RsMst_ExchangeType!ExchangeTypeID

CmdClient.Left = 4695
If mExchangeTypeID = 1 Then
   CmbNonNcdexClient.Width = 10900
   CmbNonNcdexClient.Top = TxtClientID.Top
   CmbNonNcdexClient.Left = TxtClientID.Left
   Label9.Visible = False
   Label10.Visible = False
   Label11.Visible = False
   CmdClient.Left = 7320
   CmbNonNcdexClient.Visible = True
   CmbNonNcdexClient.SetFocus
Else
   Label9.Visible = True
   Label10.Visible = True
   Label11.Visible = True
   CmbNonNcdexClient.Visible = False
   CmdClient.Left = 4695
End If
End Sub

Private Sub SaveCmd_Click()

If TxtExpenseFrom.Value > TxtExpenseTo.Value Then
   MsgBox "'Expense From ' must be less than 'Expense To'!!!"
   TxtExpenseFrom.SetFocus
   Exit Sub
End If

If TxtExpenseFrom.Value > Date Then
   MsgBox "'Expense From' must be less than 'Claim Date'!!!"
   TxtExpenseFrom.SetFocus
   Exit Sub
End If

If Format(TxtExpenseFrom.Value, "yyyymm") > "201103" Then
   MsgBox "'Expense From' should be upto 'March 2011' only!!!"
   TxtExpenseFrom.SetFocus
   Exit Sub
End If

If CmbEmployeeID.Text = "" Then
   MsgBox "Blank Employee not allowed!!!"
   CmbEmployeeID.SetFocus
   Exit Sub
End If

If CmbLocationID.Text = "" Then
   MsgBox "Blank Location not allowed!!!"
   CmbLocationID.SetFocus
   Exit Sub
End If

If Edit_Flg = "E" Then
   If LCase(Gen_UserRole) = "fag" Or LCase(Gen_UserRole) = "power" Or LCase(Gen_UserRole) = "fagadmin" Then
      If CmbStatus.Text = "" Then
         MsgBox "Blank Status not allowed!!!"
         CmbStatus.SetFocus
         Exit Sub
      End If
      If LCase(CmbStatus.Text) = "approved" Or LCase(CmbStatus.Text) = "rejected" Or LCase(CmbStatus.Text) = "frozen" Then
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
      End If
   End If

End If

If CmbForwardTo.Text = "" Then
   MsgBox "Blank Forward to not allowed!!!"
   CmbForwardTo.SetFocus
   Exit Sub
End If

If mEmployeeID = mForwardToID Then
   MsgBox "Employee and Forward To can't be same!!!"
   CmbForwardTo.SetFocus
   Exit Sub
End If


If TxtRemark = "" Then
   MsgBox "Blank remarks not allowed!!!"
   TxtRemark.SetFocus
   Exit Sub
End If

If MSHFlexGrid2.Rows < 4 Then
   MsgBox "No details found in grid!!!"
   CmbExpenseTypeID.SetFocus
   Exit Sub
End If

If Edit_Flg = "A" Then
   RsTxn_ClaimSubmission.AddNew
   RsTxn_ClaimSubmission!ClaimRowID = GetMaxClaimRowID
   RsTxn_ClaimSubmission!G_UID = GetGUID
   TxtClaimID.Text = RsTxn_ClaimSubmission!ClaimRowID
   TxtG_UID = RsTxn_ClaimSubmission!G_UID
   RsTxn_ClaimSubmission!ClaimID = TxtClaimID
   RsTxn_ClaimSubmission!SM_Prefix = TxtPreFix
   RsTxn_ClaimSubmission!DesignationID = mEmpDesignationID
Else
   RsTxn_ClaimSubmission.MoveFirst
   RsTxn_ClaimSubmission.Find "G_UID= '" & TxtG_UID & "'"
End If
'RsTxn_ClaimSubmission!ClaimRowID = TxtClaimID
RsTxn_ClaimSubmission!ClaimDate = TxtClaimDate.Value
RsTxn_ClaimSubmission!EmployeeID = mEmployeeID
RsTxn_ClaimSubmission!LocationID = mLocationID
RsTxn_ClaimSubmission!Amount = Val(MSHFlexGrid2.TextMatrix(0, 3)) 'Val(TxtClaimAmount)
RsTxn_ClaimSubmission!Flag = mStatus

TxtRemark = Replace(TxtRemark, ",", "-")
RsTxn_ClaimSubmission!Remarks = TxtRemark

'Commented on 6 April 2010
'If TxtApprovalRemarks.Locked = False Then
'   RsTxn_ClaimSubmission!ApprovalRemarks = TxtApprovalRemarks
'End If
'If TxtApprovalDate.Visible = True Then
'   RsTxn_ClaimSubmission!ApprovalDate = TxtApprovalDate.Value
'End If

If LCase(mStatus) <> "p" Then
   TxtApprovalRemarks = Replace(TxtApprovalRemarks, ",", "-")
   RsTxn_ClaimSubmission!ApprovalRemarks = TxtApprovalRemarks
   RsTxn_ClaimSubmission!ApprovalDate = TxtApprovalDate.Value
Else
   RsTxn_ClaimSubmission!ApprovalDate = Null
End If

RsTxn_ClaimSubmission!ForwardToID = mForwardToID
RsTxn_ClaimSubmission!G_UID = TxtG_UID
RsTxn_ClaimSubmission!ExpenseFrom = TxtExpenseFrom.Value
RsTxn_ClaimSubmission!ExpenseTo = TxtExpenseTo.Value
RsTxn_ClaimSubmission!BatchID = mBatchID


If IsNull(RsTxn_ClaimSubmission!CreatedBy) = True Or RsTxn_ClaimSubmission!CreatedBy = "" Then
   RsTxn_ClaimSubmission!CreatedBy = Gen_UserLoginID
   RsTxn_ClaimSubmission!CreatedDate = Now
Else
   RsTxn_ClaimSubmission!UpdatedBy = Gen_UserLoginID
   RsTxn_ClaimSubmission!UpdatedDate = Now
End If

If LCase(Gen_UserRole) = "fag" Or LCase(Gen_UserRole) = "power" Or LCase(Gen_UserRole) = "fagadmin" Then
   If LCase(mStatus) = "a" Then
      RsTxn_ClaimSubmission!UpdatedByFAG = Gen_UserLoginID
      RsTxn_ClaimSubmission!UpdatedDateFAG = Now
   End If
End If

RsTxn_ClaimSubmission.Update

TxtCreated = IIf(IsNull(RsTxn_ClaimSubmission!CreatedBy), "", RsTxn_ClaimSubmission!CreatedBy) & " :- " & IIf(IsNull(RsTxn_ClaimSubmission!CreatedDate), "", RsTxn_ClaimSubmission!CreatedDate)
If Edit_Flg = "E" Then
   TxtUpdated = IIf(IsNull(RsTxn_ClaimSubmission!UpdatedBy), "", RsTxn_ClaimSubmission!UpdatedBy) & " :- " & IIf(IsNull(RsTxn_ClaimSubmission!UpdatedDate), "", RsTxn_ClaimSubmission!UpdatedDate)
End If

Call SaveClaimSubmissionDetails

If LCase(mStatus) <> "f" Then
   Call Gen_Email
End If
RsTxn_ClaimSubmission.Filter = ""
RsTxn_ClaimSubmission.Requery
RsTxn_ClaimSubmission.MoveFirst
RsTxn_ClaimSubmission.Find "G_UID = '" & TxtG_UID & "'"

Call Indata

End Sub
Private Sub SaveClaimSubmissionDetails()
Call DeleteClaimSubmissionDetails
Call TableTxn_ClaimSubmissionDetails("") '(" Where ClaimID = " & TxtClaimID & " AND SM_Prefix = '" & TxtPreFix & "'")
For I = 2 To MSHFlexGrid2.Rows - 1
    If MSHFlexGrid2.TextMatrix(I, 1) <> "" Then
       RsTxn_ClaimSubmissionDetails.AddNew
       RsTxn_ClaimSubmissionDetails!ClaimDetailID = GetMaxClaimSubmissionDetails
       RsTxn_ClaimSubmissionDetails!ClaimRowID = TxtClaimID
       RsTxn_ClaimSubmissionDetails!ClaimID = RsTxn_ClaimSubmission!ClaimID
       RsTxn_ClaimSubmissionDetails!SM_Prefix = TxtPreFix
       RsTxn_ClaimSubmissionDetails!ExpenseTypeID = MSHFlexGrid2.TextMatrix(I, 12)
       RsTxn_ClaimSubmissionDetails!Particulars = MSHFlexGrid2.TextMatrix(I, 2)
       RsTxn_ClaimSubmissionDetails!Amount = Val(MSHFlexGrid2.TextMatrix(I, 3))
       If MSHFlexGrid2.TextMatrix(I, 4) <> "" Then
          RsTxn_ClaimSubmissionDetails!Quantity = Val(MSHFlexGrid2.TextMatrix(I, 4))
       End If
       RsTxn_ClaimSubmissionDetails!Recoverable = MSHFlexGrid2.TextMatrix(I, 5)
       If MSHFlexGrid2.TextMatrix(I, 5) <> "N" Then
          RsTxn_ClaimSubmissionDetails!ExchangeTypeID = GetExchangeID(MSHFlexGrid2.TextMatrix(I, 6))
          RsTxn_ClaimSubmissionDetails!ClientIDRow = MSHFlexGrid2.TextMatrix(I, 11)
          RsTxn_ClaimSubmissionDetails!RecoverableAmount = Val(MSHFlexGrid2.TextMatrix(I, 13))
       End If
'      If IsNull(RsMst_EmpCMPMap!CreatedBy) = True Or RsMst_EmpCMPMap!CreatedBy = "" Then
'          RsMst_EmpCMPMap!CreatedBy = Gen_UserLoginID
'          RsMst_EmpCMPMap!CreatedDate = Now
'       Else
'          RsMst_EmpCMPMap!UpdatedBy = Gen_UserLoginID
'          RsMst_EmpCMPMap!UpdatedDate = Now
'       End If
       
       RsTxn_ClaimSubmissionDetails.Update
       RsTxn_ClaimSubmissionDetails.Requery
    End If
Next
End Sub
Private Sub Gen_Email()
Dim TOADD As Variant
Dim TOCC, mMsg, mGridMsg As Variant
Dim strSplitString() As String
Dim EmpNo_, EmpName_ As String

mMsg = ""
mMsg = "<tr>Claim Submission (From Store-Man)</tr>"
mMsg = mMsg & "<tr><td> Claim ID </td><td> " & TxtPreFix & " - " & TxtClaimID & " </td></tr>"
mMsg = mMsg & "<tr><td> Claim Date </td><td> " & Format(TxtClaimDate.Value, "dd-mmm-yyyy") & "  </td></tr>"
mMsg = mMsg & "<tr><td> Expense From </td><td> " & Format(TxtExpenseFrom.Value, "dd-mmm-yyyy") & "  </td></tr>"
mMsg = mMsg & "<tr><td> Expense To </td><td> " & Format(TxtExpenseTo.Value, "dd-mmm-yyyy") & "  </td></tr>"
If CmbStatus.Visible = True Then
   mMsg = mMsg & "<tr><td> Status </td><td> " & CmbStatus.Text & "  </td></tr>"
Else
   mMsg = mMsg & "<tr><td> Status </td><td> " & TxtStatus & "  </td></tr>"
End If

strSplitString = Split(CmbEmployeeID.Text, "~")
EmpName_ = Trim(strSplitString(0))
EmpNo_ = Trim(strSplitString(1))
 
mMsg = mMsg & "<tr><td> Employee Name </td><td> " & EmpName_ & "  </td></tr>"
mMsg = mMsg & "<tr><td> Employee No </td><td> " & EmpNo_ & "  </td></tr>"

mMsg = mMsg & "<tr><td> Location </td><td> " & CmbLocationID.Text & "  </td></tr>"
mMsg = mMsg & "<tr><td> Entered By </td><td> " & TxtCreated & "  </td></tr>"
If Edit_Flg = "E" Then
   mMsg = mMsg & "<tr><td> Updated By </td><td> " & TxtUpdated & "  </td></tr>"
End If
mMsg = mMsg & "<tr><td> Total Claim Amount </td><td> " & TxtClaimAmount & "  </td></tr>"
mMsg = mMsg & "<tr><td> Remark </td><td> " & TxtRemark & "   </td></tr>"

strSplitString = Split(CmbForwardTo.Text, "~")
EmpName_ = Trim(strSplitString(0))
EmpNo_ = Trim(strSplitString(1))
mMsg = mMsg & "<tr><td> Forwarded To </td><td> " & EmpName_ & " " & " ~ " & " " & EmpNo_ & " </td></tr>"


If TxtApprovalRemarks <> "" Then
   mMsg = mMsg & "<tr><td> FAG Remarks </td><td> " & TxtApprovalRemarks & " </td></tr>"
   mMsg = mMsg & "<tr><td> FAG Process Date </td><td> " & Format(TxtApprovalDate.Value, "dd-mmm-yyyy") & " </td></tr>"
End If
'mGridMsg = "<tr><td>Expense Type</td><td>Claim Amount</td><td>Particulars</td></tr>"
For I = 1 To MSHFlexGrid2.Rows - 1
    If MSHFlexGrid2.TextMatrix(I, 1) <> "" Then
       mGridMsg = mGridMsg & "<tr><td> " & MSHFlexGrid2.TextMatrix(I, 1) & " </td><td> " & MSHFlexGrid2.TextMatrix(I, 2) & " </td><td> " & MSHFlexGrid2.TextMatrix(I, 3) & " </td></tr>"
    End If
Next

mMsg = mMsg & "<tr><td> Details </td><td><table border = 1>" & mGridMsg & "</table></td>"


'TOADD = mForwardToEmailID ' mEmployeeEmailID
TOCC = Gen_UserEmailID '& "," & mForwardToEmailID

If LCase(mStatus) = "p" Then
   TOADD = mForwardToEmailID
   If LCase(TOCC) <> LCase(mEmployeeEmailID) Then
      TOCC = TOCC & "," & mEmployeeEmailID
   End If
   mMsgSubject = "Claim Pending"
ElseIf LCase(mStatus) = "a" Then
   TOADD = mEmployeeEmailID
   If LCase(TOCC) <> LCase(mForwardToEmailID) Then
      TOCC = TOCC & "," & mForwardToEmailID
   End If
   mMsgSubject = "Claim Approved"
ElseIf LCase(mStatus) = "r" Then
   TOADD = mEmployeeEmailID
   If LCase(TOCC) <> LCase(mForwardToEmailID) Then
      TOCC = TOCC & "," & mForwardToEmailID
   End If
   mMsgSubject = "Claim Rejected"
End If

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
'  'connection
'  .Server = servertxt
'  'message
'  If Edit_Flg = "A" Then
'     .MailFrom = mEmployeeEmailID
'  Else
'     .MailFrom = Gen_UserEmailID
'  End If
'  If IsNull(Trim(ToADD_)) = True Or Trim(ToADD_) = "" Then
'     ToADD_ = "bhagwat.k@ncml.com"
'  End If
'  .SendTo = ToADD_
'  If ToADD_ <> TOCC_ Then
'     .CC = TOCC_
'  End If
'  .BCC = "mayur.d@ncml.com"
'  .MessageSubject = MsgSubject_ & " (from Store-Man)" '"Confidential Mail"
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
   ToADD_ = "vaishali.k@ncml.com"
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

Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   If RsTxn_ClaimSubmission.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   RsTxn_ClaimSubmission.MoveFirst
 
   RsTxn_ClaimSubmission.Find "G_UID = '" & TxtG_UID & "'"
   If RsTxn_ClaimSubmission.EOF Then
      RsTxn_ClaimSubmission.MoveLast
   End If
   Call Indata
   Exit Sub
End If

Frame0.Height = Frame4.Height - 300 'Me.Height - 500
Frame0.Left = Frame1.Left - 100
Frame0.Top = Frame1.Top
Frame0.Width = Frame1.Width + 100 'Me.Width - 200
Frame0.Visible = True

MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 1400

RsTxn_ClaimSubmission.MoveFirst

Call Grid_Head1

SFrom.Value = CDate(From_GenDate) - 900
STo.Value = Date 'To_GenDate
CmbSLocation.Text = ""
CmbSEmployee.Text = ""
CmbSStatus.Text = "Pending"
CmbSForwardTo.Text = ""
ChkBatchNo.Value = 0
TxtSClaimID = ""
Label22.Caption = ""
mSStatus = "P"

tmp = "Select EmployeeID,EmployeeName +'~'+ EmployeeNo as Employee,DesignationID,EmployeeEmailID From Mst_Employee Where EmployeeNo = '" & Gen_UserLoginID & "'"
Call TmpTable

If TmpRs.RecordCount = 1 Then
   CmbSForwardTo.Text = TmpRs!Employee
   mSForwardToID = TmpRs!EmployeeID
Else
   CmbSForwardTo.Text = ""
   mSForwardToID = Null
End If



End Sub

Private Sub TxtAmount_LostFocus()
If TxtAmount = "" Then Exit Sub
If IsNumeric(TxtAmount) = False Or Val(TxtAmount) < 0 Then
   MsgBox "Invalid Format!!!"
   TxtAmount.SetFocus
   Exit Sub
End If

TxtAmount = Format(TxtAmount, "########0.00")
End Sub

Private Sub TxtClientID_LostFocus()
CmbClientDP.Visible = False
If TxtClientID = "" Then Exit Sub
If CmbExchangeTypeID.Text = "" Then Exit Sub

Dim ans As Variant
ans = GetClientNameDtl(TxtClientID, mExchangeTypeID)

If ans(0, 0) = False Then
   MsgBox "Not Found !!!"
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
End If
End Sub

Private Sub CmbClientDP_LostFocus()

Dim Str_() As String
Str_ = Split(CmbClientDP.Text, "~")
If UBound(Str_) <> 1 Then
   MsgBox "Invalid selection !!!!! "
   If CmbClientDP.Visible = True Then
      CmbClientDP.SetFocus
   Else
      TxtClientID.SetFocus
   End If
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
End Sub
Private Sub CmbNonNcdexClient_GotFocus()
tmp = CmbNonNcdexClient.Text
Call TableMst_Client("where ExchangeTypeId = " & mExchangeTypeID)
CmbNonNcdexClient.Clear
If RsMst_Client.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Client.RecordCount
    CmbNonNcdexClient.AddItem RsMst_Client!ClientName
    RsMst_Client.MoveNext
Next
CmbNonNcdexClient.Text = tmp
End Sub
Private Sub CmbNonNcdexClient_LostFocus()
If CmbNonNcdexClient.Text = "" Then
   Exit Sub
End If
RsMst_Client.MoveFirst
RsMst_Client.Find "ClientName = '" & CmbNonNcdexClient.Text & "'"
If RsMst_Client.EOF Then
   MsgBox "Invalid selection "
   CmbNonNcdexClient.SetFocus
   Exit Sub
End If
mClientIDRow = RsMst_Client!ClientIDRow
Dim ans As Variant
ans = GetClientNameByRow(RsMst_Client!ClientIDRow)
TxtClientID = ans(0)
TxtClientName = ans(1)
TxtDPID = ans(2)
TxtDPName = ans(3)

CmdAdd2List.SetFocus

End Sub

Function ChkDuplicateExpenseType() As Boolean
Dim Retval As Boolean
Retval = True
If MSHFlexGrid2.Rows > 2 Then
   For u = 2 To MSHFlexGrid2.Rows - 1
       If MSHFlexGrid2.TextMatrix(u, 1) = CmbExpenseTypeID.Text Then
          If LCase(CmdAdd2List.Caption) = "add to list" Then
             Retval = False
             Exit For
          Else
             If u <> MSHFlexGrid2.RowSel Then
                Retval = False
             End If
          End If
       End If
   Next
End If

ChkDuplicateExpenseType = Retval
End Function

Private Sub DeleteClaimSubmissionDetails()
tmp = "Delete From Txn_ClaimSubmissionDetail Where ClaimRowID=" & TxtClaimID & " "
Call Tmp4Table
End Sub
Private Function GetMaxClaimRowID() As Double
Dim Retval As Double
tmp = "Select max(ClaimRowID) from Txn_ClaimSubmission "
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetMaxClaimRowID = Retval
End Function

Function GetMaxClaimSubmissionDetails() As Double
Dim Retval As Double
Retval = 0
tmp = "Select max(ClaimDetailID) from Txn_ClaimSubmissionDetail " 'where ClaimID = " & TxtClaimID & " And SM_prefix = '" & TxtPreFix & "'"
Call TmpTable
If TmpRs.RecordCount > 0 Then
   Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
End If
Retval = Retval + 1
TmpRs.Close
GetMaxClaimSubmissionDetails = Retval

End Function


Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsTxn_ClaimSubmission.MoveNext
If RsTxn_ClaimSubmission.EOF Then
   RsTxn_ClaimSubmission.MoveLast
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
RsTxn_ClaimSubmission.MovePrevious
If RsTxn_ClaimSubmission.BOF Then
   RsTxn_ClaimSubmission.MoveFirst
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
RsTxn_ClaimSubmission.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsTxn_ClaimSubmission.MoveLast
Call Indata
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False
End Sub


Public Sub Indata()

Call ClearFrame
Frame2.Enabled = False
TxtApprovalDate.Visible = False
Label27.Visible = False

mStatus = ""
TxtEmployeeID.Visible = True
CmbEmployeeID.Visible = False
CmdEmployeeID.Visible = Fasle

TxtForwardTo.Visible = True
CmbForwardTo.Visible = False
CmdForwardTo.Visible = False

TxtLocationID.Visible = True
CmbLocationID.Visible = False

TxtStatus.Visible = True
CmbStatus.Visible = False

TxtExpenseFrom.Enabled = False
TxtExpenseTo.Enabled = False
TxtClaimDate.Enabled = False

TxtClaimID = RsTxn_ClaimSubmission!ClaimRowID
TxtClaimDate.Value = RsTxn_ClaimSubmission!ClaimDate
TxtExpenseFrom.Value = RsTxn_ClaimSubmission!ExpenseFrom
TxtExpenseTo.Value = RsTxn_ClaimSubmission!ExpenseTo
mEmployeeID = RsTxn_ClaimSubmission!EmployeeID
TxtEmployeeID = GetEmployeeNonName(RsTxn_ClaimSubmission!EmployeeID)
mEmployeeEmailID = GetEmployeeEmailID(mEmployeeID)

'Commented on 21 May 2010 by Parag ( As mail by Mayur Sir on 20 May for Modification in Store-Man Claim)
'TxtDesignation = GetDesignationFromEmp(RsTxn_ClaimSubmission!EmployeeID)


mEmpDesignationID = IIf(IsNull(RsTxn_ClaimSubmission!DesignationID), 0, RsTxn_ClaimSubmission!DesignationID)
TxtDesignation = GetDesignaion(mEmpDesignationID)


mLocationID = IIf(IsNull(RsTxn_ClaimSubmission!LocationID), 0, RsTxn_ClaimSubmission!LocationID)
TxtLocationID = GetLocationName(RsTxn_ClaimSubmission!LocationID)

TxtPreFix = RsTxn_ClaimSubmission!SM_Prefix
mStatus = RsTxn_ClaimSubmission!Flag
mBatchID = IIf(IsNull(RsTxn_ClaimSubmission!BatchID), 0, RsTxn_ClaimSubmission!BatchID)
mForwardToID = IIf(IsNull(RsTxn_ClaimSubmission!ForwardToID), "", RsTxn_ClaimSubmission!ForwardToID)

If mForwardToID <> "" Then
   TxtForwardTo = GetEmployeeNonName(RsTxn_ClaimSubmission!ForwardToID)
   mForwardToEmailID = GetEmployeeEmailID(mForwardToID)
Else
   TxtForwardTo = "" 'GetEmployeeNonName(RsTxn_ClaimSubmission!ForwardToID)
   mForwardToEmailID = Null 'GetEmployeeEmailID(mForwardToID)
End If

If LCase(mStatus) = "a" Then
   TxtStatus = "Approved"
ElseIf LCase(mStatus) = "p" Then
   TxtStatus = "Pending"
ElseIf LCase(mStatus) = "r" Then
   TxtStatus = "Rejected"
ElseIf LCase(mStatus) = "f" Then
   TxtStatus = "Frozen"
Else
   TxtStatus = ""
End If
TxtClaimAmount = RsTxn_ClaimSubmission!Amount
TxtRemark = IIf(IsNull(RsTxn_ClaimSubmission!Remarks), "", RsTxn_ClaimSubmission!Remarks)
TxtApprovalRemarks = IIf(IsNull(RsTxn_ClaimSubmission!ApprovalRemarks), "", RsTxn_ClaimSubmission!ApprovalRemarks)

If IsNull(RsTxn_ClaimSubmission!ApprovalDate) = False Then
   TxtApprovalDate.Visible = True
   Label27.Visible = True
   TxtApprovalDate.Value = RsTxn_ClaimSubmission!ApprovalDate
   TxtApprovalDate.Enabled = False
End If

TxtG_UID = IIf(IsNull(RsTxn_ClaimSubmission!G_UID), "", RsTxn_ClaimSubmission!G_UID)

TxtLocationID.Locked = True
TxtRemark.Locked = True
TxtApprovalRemarks.Locked = True

TxtCreated = IIf(IsNull(RsTxn_ClaimSubmission!CreatedBy), "", RsTxn_ClaimSubmission!CreatedBy) & " :- " & IIf(IsNull(RsTxn_ClaimSubmission!CreatedDate), "", RsTxn_ClaimSubmission!CreatedDate)
TxtUpdated = IIf(IsNull(RsTxn_ClaimSubmission!UpdatedBy), "", RsTxn_ClaimSubmission!UpdatedBy) & " :- " & IIf(IsNull(RsTxn_ClaimSubmission!UpdatedDate), "", RsTxn_ClaimSubmission!UpdatedDate)
TxtUpdatedByFAG = IIf(IsNull(RsTxn_ClaimSubmission!UpdatedByFAG), "", RsTxn_ClaimSubmission!UpdatedByFAG) & " :- " & IIf(IsNull(RsTxn_ClaimSubmission!UpdatedDateFAG), "", RsTxn_ClaimSubmission!UpdatedDateFAG)

Call InDataClaimDetails

Call InDataClaimForwardHistory


If RsTxn_ClaimSubmission.RecordCount = 1 Then
   Call GButtonLock_1(Me, True)
Else
   Call GButtonLock(Me, True)
End If

SSTab1.Tab = 0

End Sub
Private Sub InDataClaimForwardHistory()

Call Grid_Head2

tmp = "Select TCS.ClaimRowID,MU.EmployeeName+'~'+ MU.EmployeeNo,MU1.EmployeeName+'~'+ MU1.EmployeeNo, " & _
      " TCS.ForwardToID,EMP.EmployeeName+'~'+ EMP.EmployeeNo as 'OldFwdTo', " & _
      " EMP1.EmployeeName+'~'+ EMP1.EmployeeNo as 'NewFwdTo',LCFTH.UpdatedBy,LCFTH.Concurrency,TCS.CreatedDate,TCS.UpdatedDate  " & _
      " From Log_ClaimForwardToHistory LCFTH  " & _
      " Left Join Mst_Employee EMP On LCFTH.OldForwardToID = EMP.EmployeeID  " & _
      " Left Join Mst_Employee EMP1 On LCFTH.NewForwardToID = EMP1.EmployeeID " & _
      " Right Join Txn_ClaimSubmission TCS On LCFTH.ClaimRowID = TCS.ClaimRowID  " & _
      " Inner Join Mst_Employee EMP2 On TCS.EmployeeID = EMP2.EmployeeID  " & _
      " Inner Join Mst_Users MU On TCS.CreatedBy = MU.EmployeeNo  " & _
      " Left Join Mst_Users MU1 On TCS.UpdatedBy = MU1.EmployeeNo  " & _
      " Where TCS.ClaimRowID = " & TxtClaimID & " " & _
      " Order by TCS.ClaimRowID "

Call TmpTable
For iu = 1 To TmpRs.RecordCount
    MSHFlexGrid3.Rows = TmpRs.RecordCount + 1
    For C = 0 To MSHFlexGrid3.Cols - 1
         MSHFlexGrid3.TextMatrix(iu, 0) = IIf(IsNull(TmpRs!ClaimRowID), "", TmpRs!ClaimRowID)
         MSHFlexGrid3.TextMatrix(iu, 1) = IIf(IsNull(TmpRs.Fields(4)), TmpRs.Fields(1) & ":- " & Format(TmpRs!CreatedDate, "dd-mmm-yyyy"), TmpRs.Fields(4))
         MSHFlexGrid3.TextMatrix(iu, 2) = IIf(IsNull(TmpRs.Fields(5)), TmpRs.Fields(2) & ":- " & Format(TmpRs!UpdatedDate, "dd-mmm-yyyy"), TmpRs.Fields(5))
         MSHFlexGrid3.TextMatrix(iu, 3) = IIf(IsNull(TmpRs!UpdatedBy), "", TmpRs!UpdatedBy)
         MSHFlexGrid3.TextMatrix(iu, 4) = IIf(IsNull(TmpRs!Concurrency), "", TmpRs!Concurrency)
         MSHFlexGrid3.TextMatrix(iu, 4) = Format(MSHFlexGrid3.TextMatrix(iu, 4), "dd-mmm-yyyy")
    Next
    TmpRs.MoveNext
Next

End Sub

Private Sub InDataClaimDetails()

Call Grid_Head
Call TableTxn_ClaimSubmissionDetails(" Where ClaimRowID= " & TxtClaimID & " ") 'AND SM_Prefix = '" & TxtPreFix & "'")
If RsTxn_ClaimSubmissionDetails.RecordCount > 0 Then
    MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + RsTxn_ClaimSubmissionDetails.RecordCount
    For I = 2 To MSHFlexGrid2.Rows - 2
        MSHFlexGrid2.TextMatrix(I, 0) = RsTxn_ClaimSubmissionDetails!ClaimDetailID
        MSHFlexGrid2.TextMatrix(I, 1) = GetExpenseType(RsTxn_ClaimSubmissionDetails!ExpenseTypeID)
        MSHFlexGrid2.TextMatrix(I, 2) = RsTxn_ClaimSubmissionDetails!Particulars
        'MSHFlexGrid2.TextMatrix(i, 3) = IIf(IsNull(RsTxn_ClaimSubmissionDetails!Supporting), "N", RsTxn_ClaimSubmissionDetails!Supporting)
        MSHFlexGrid2.TextMatrix(I, 3) = Val(RsTxn_ClaimSubmissionDetails!Amount)
        MSHFlexGrid2.TextMatrix(I, 4) = IIf(IsNull(RsTxn_ClaimSubmissionDetails!Quantity), "", RsTxn_ClaimSubmissionDetails!Quantity)
        MSHFlexGrid2.TextMatrix(I, 5) = IIf(IsNull(RsTxn_ClaimSubmissionDetails!Recoverable), "N", RsTxn_ClaimSubmissionDetails!Recoverable)
        If MSHFlexGrid2.TextMatrix(I, 5) = "Y" Then
           MSHFlexGrid2.TextMatrix(I, 6) = GetExchangeName(RsTxn_ClaimSubmissionDetails!ExchangeTypeID)
           ans = GetClientNameByRow(RsTxn_ClaimSubmissionDetails!ClientIDRow)
           MSHFlexGrid2.TextMatrix(I, 7) = ans(0) 'RsTxn_ClaimSubmissionDetails!Particulars
           MSHFlexGrid2.TextMatrix(I, 8) = ans(1) 'RsTxn_ClaimSubmissionDetails!Particulars
           MSHFlexGrid2.TextMatrix(I, 9) = ans(2)
           MSHFlexGrid2.TextMatrix(I, 10) = ans(3)
           MSHFlexGrid2.TextMatrix(I, 11) = RsTxn_ClaimSubmissionDetails!ClientIDRow
           MSHFlexGrid2.TextMatrix(I, 13) = RsTxn_ClaimSubmissionDetails!RecoverableAmount
        Else
           MSHFlexGrid2.TextMatrix(I, 6) = "" 'GetExchangeName(RsTxn_ClaimSubmissionDetails!ExchangeTypeID)
           'Ans = GetClientNameByRow(RsTxn_ClaimSubmissionDetails!ClientIDRow)
           MSHFlexGrid2.TextMatrix(I, 7) = "" 'Ans(0) 'RsTxn_ClaimSubmissionDetails!Particulars
           MSHFlexGrid2.TextMatrix(I, 8) = "" 'Ans(1) 'RsTxn_ClaimSubmissionDetails!Particulars
           MSHFlexGrid2.TextMatrix(I, 9) = "" 'Ans(2)
           MSHFlexGrid2.TextMatrix(I, 10) = "" 'Ans(3)
           MSHFlexGrid2.TextMatrix(I, 11) = "" 'RsTxn_ClaimSubmissionDetails!ClientIDRow
           MSHFlexGrid2.TextMatrix(I, 13) = "" 'RsTxn_ClaimSubmissionDetails!RecoverableAmount
        End If
        MSHFlexGrid2.TextMatrix(I, 12) = RsTxn_ClaimSubmissionDetails!ExpenseTypeID
        MSHFlexGrid2.TextMatrix(0, 3) = Val(MSHFlexGrid2.TextMatrix(0, 3)) + Val(RsTxn_ClaimSubmissionDetails!Amount)
        If IsNull(RsTxn_ClaimSubmissionDetails!RecoverableAmount) = False Then
           MSHFlexGrid2.TextMatrix(0, 13) = Val(MSHFlexGrid2.TextMatrix(0, 13)) + Val(RsTxn_ClaimSubmissionDetails!RecoverableAmount)
        End If
        RsTxn_ClaimSubmissionDetails.MoveNext
    Next
    'TxtClaimAmount = Val(MSHFlexGrid2.TextMatrix(0, 3))
End If
End Sub


Private Sub Grid_Head1()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.FixedRows = 1
MSHFlexGrid1.Cols = 17

MSHFlexGrid1.TextMatrix(0, 0) = "New Claim ID"
MSHFlexGrid1.TextMatrix(0, 1) = "SM_Prefix"
MSHFlexGrid1.TextMatrix(0, 2) = "Claim ID"
MSHFlexGrid1.TextMatrix(0, 3) = "Claim Date"
MSHFlexGrid1.TextMatrix(0, 4) = "Expense From"
MSHFlexGrid1.TextMatrix(0, 5) = "Expense To"
MSHFlexGrid1.TextMatrix(0, 6) = "Employee"
MSHFlexGrid1.TextMatrix(0, 7) = "Designation"
MSHFlexGrid1.TextMatrix(0, 8) = "Location"
MSHFlexGrid1.TextMatrix(0, 9) = "Claim Amount"
MSHFlexGrid1.TextMatrix(0, 10) = "Remarks"
MSHFlexGrid1.TextMatrix(0, 11) = "Status"
MSHFlexGrid1.TextMatrix(0, 12) = "G_UID"
MSHFlexGrid1.TextMatrix(0, 13) = "FAG Remarks"
MSHFlexGrid1.TextMatrix(0, 14) = "Booked in NAV Date"
MSHFlexGrid1.TextMatrix(0, 15) = "Claim Forwarded To"
MSHFlexGrid1.TextMatrix(0, 16) = "Batch No"


MSHFlexGrid1.ColWidth(0) = 1000
MSHFlexGrid1.ColWidth(1) = 1000
MSHFlexGrid1.ColWidth(2) = 1000
MSHFlexGrid1.ColWidth(3) = 1100
MSHFlexGrid1.ColWidth(4) = 1100
MSHFlexGrid1.ColWidth(5) = 1200
MSHFlexGrid1.ColWidth(6) = 3200
MSHFlexGrid1.ColWidth(7) = 1400
MSHFlexGrid1.ColWidth(8) = 1400
MSHFlexGrid1.ColWidth(9) = 1600
MSHFlexGrid1.ColWidth(10) = 3200
MSHFlexGrid1.ColWidth(12) = 0
MSHFlexGrid1.ColWidth(13) = 3000
MSHFlexGrid1.ColWidth(14) = 1600
MSHFlexGrid1.ColWidth(15) = 3200
MSHFlexGrid1.ColWidth(16) = 800
   
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
    CmbSLocation.AddItem RsMst_Location!LocationName
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
Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Txn_ClaimSubmission.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Txn_ClaimSubmission.xls")
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

Private Sub TxtExpenseFrom_LostFocus()

mMonth = Month(TxtExpenseFrom.Value)
mYear = Year(TxtExpenseFrom.Value)
TxtExpenseFrom.Value = mYear & "-" & mMonth & "- 01"
mTotDays = Day(DateSerial(mYear, mMonth + 1, 0))
TxtExpenseTo.Value = mYear & "-" & mMonth & "-" & mTotDays

End Sub
Private Sub TxtExpenseTo_LostFocus()
mMonth = Month(TxtExpenseFrom.Value)
mYear = Year(TxtExpenseFrom.Value)
mTotDays = Day(DateSerial(mYear, mMonth + 1, 0))
TxtExpenseTo.Value = mYear & "-" & mMonth & "-" & mTotDays
End Sub

Private Sub TxtQty_LostFocus()
If TxtQty = "" Then Exit Sub
If IsNumeric(TxtQty) = False Or Val(TxtQty) <= 0 Then
   MsgBox "Invalid Format!!!"
   TxtQty.SetFocus
   Exit Sub
End If
TxtQty = Format(TxtQty, "########0")
End Sub

Private Sub TxtRecoverableAmount_LostFocus()
If TxtRecoverableAmount = "" Then Exit Sub
If IsNumeric(TxtRecoverableAmount) = False Or Val(TxtRecoverableAmount) < 0 Then
   MsgBox "Invalid Format!!!"
   TxtRecoverableAmount.SetFocus
   Exit Sub
End If

TxtRecoverableAmount = Format(TxtRecoverableAmount, "########0.00")
End Sub


Private Sub CmbForwardTo_GotFocus()

tmp = CmbForwardTo.Text
Call TableMst_Employee("Where Flag = 'Y'")
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
If UBound(strSplitString) <> 1 Then
   MsgBox "Invalid selection!!!"
   CmbForwardTo.SetFocus
   Exit Sub
End If
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

Private Sub CmbEmployeeID_GotFocus()
Call GetEmployeeDataonGotFocus

End Sub
Private Sub GetEmployeeDataonGotFocus()
tmp = CmbEmployeeID.Text
Call TableMst_Employee("Where Flag = 'Y'")
CmbEmployeeID.Clear
If RsMst_Employee.RecordCount = 0 Then
   MsgBox "No Employee Found!!!"
   CmbEmployeeID.SetFocus
   'RsMst_Employee.MoveFirst
   Exit Sub
End If
For I = 1 To RsMst_Employee.RecordCount
    CmbEmployeeID.AddItem RsMst_Employee!EmployeeName & "~" & RsMst_Employee!EmployeeNo
    RsMst_Employee.MoveNext
Next

CmbEmployeeID.Text = tmp

End Sub

Private Sub GetSearchEmployeeDataOnLostFocus()
Dim strSplitString() As String
Dim EmpNo_, EmpName_ As String

If CmbSEmployee.Text = "" Then
   mSEmployeeID = Null
   Exit Sub
End If

strSplitString = Split(CmbSEmployee, "~")
EmpName_ = Trim(strSplitString(0))
EmpNo_ = Trim(strSplitString(1))


RsMst_Employee.MoveFirst
RsMst_Employee.Find "EmployeeNo = '" & EmpNo_ & "'"

If RsMst_Employee.EOF Then
   MsgBox "Invalid Selection "
   CmbSEmployee.SetFocus
   Exit Sub
End If
mSEmployeeID = RsMst_Employee!EmployeeID

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

