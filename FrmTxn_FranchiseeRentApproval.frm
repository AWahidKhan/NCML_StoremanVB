VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmTxn_FranchiseeRentApproval 
   Caption         =   "Franchisee Fee Claim"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      Height          =   9735
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   14940
      Begin VB.Frame Frame0 
         Height          =   2775
         Left            =   -120
         TabIndex        =   14
         Top             =   4560
         Visible         =   0   'False
         Width           =   13905
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
            Left            =   3030
            Sorted          =   -1  'True
            TabIndex        =   89
            Top             =   570
            Width           =   2565
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
            Left            =   9795
            Sorted          =   -1  'True
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   570
            Width           =   3975
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
            Left            =   5670
            Sorted          =   -1  'True
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   570
            Width           =   4110
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
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   1275
            Width           =   2835
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
            ItemData        =   "FrmTxn_FranchiseeRentApproval.frx":0000
            Left            =   5505
            List            =   "FrmTxn_FranchiseeRentApproval.frx":0013
            Sorted          =   -1  'True
            TabIndex        =   18
            Top             =   1305
            Width           =   1485
         End
         Begin VB.CheckBox ChkBatchNo 
            Appearance      =   0  'Flat
            Caption         =   "Check1"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   4815
            TabIndex        =   17
            Top             =   1320
            Width           =   225
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
            Left            =   12720
            TabIndex        =   16
            Top             =   1260
            Width           =   1125
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
            Left            =   11595
            TabIndex        =   15
            Top             =   1260
            Width           =   1125
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   1455
            Left            =   90
            TabIndex        =   22
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
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   570
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
            Format          =   118882305
            CurrentDate     =   36494
         End
         Begin MSComCtl2.DTPicker SFrom 
            Height          =   360
            Left            =   60
            TabIndex        =   24
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
            Format          =   118882305
            CurrentDate     =   36494
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
            Left            =   3720
            TabIndex        =   90
            Top             =   180
            Width           =   960
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
            Left            =   11205
            TabIndex        =   33
            Top             =   240
            Width           =   1275
         End
         Begin VB.Label Label37 
            Caption         =   "Franchisee Name"
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
            Left            =   6960
            TabIndex        =   32
            Top             =   195
            Width           =   1620
         End
         Begin VB.Label Label19 
            Caption         =   "Fee ID"
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
            Left            =   1215
            TabIndex        =   31
            Top             =   975
            Width           =   645
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
            Left            =   5910
            TabIndex        =   30
            Top             =   975
            Width           =   600
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
            Left            =   3030
            TabIndex        =   29
            Top             =   1320
            Width           =   1710
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
            Left            =   7560
            TabIndex        =   28
            Top             =   1320
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
            Left            =   9960
            TabIndex        =   27
            Top             =   1320
            Width           =   90
         End
         Begin VB.Label Label45 
            Caption         =   "Fee Entered Upto"
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
            TabIndex        =   26
            Top             =   105
            Width           =   1335
         End
         Begin VB.Label Label44 
            Caption         =   "Fee Entered From"
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
            TabIndex        =   25
            Top             =   105
            Width           =   1290
         End
      End
      Begin VB.Frame ButtonFrm 
         Height          =   1125
         Left            =   120
         TabIndex        =   70
         Top             =   8550
         Width           =   14745
         Begin VB.CommandButton CmdPrint 
            Caption         =   "Print"
            Height          =   810
            Left            =   11805
            TabIndex        =   81
            Top             =   195
            Width           =   2700
         End
         Begin VB.CommandButton AddCmd 
            Caption         =   "&Add New"
            Height          =   400
            Left            =   105
            TabIndex        =   80
            Top             =   195
            Width           =   2340
         End
         Begin VB.CommandButton NextCmd 
            Caption         =   "&Next"
            Height          =   400
            Left            =   105
            TabIndex        =   79
            Top             =   600
            Width           =   2340
         End
         Begin VB.CommandButton SaveCmd 
            Caption         =   "Save"
            Height          =   400
            Left            =   2445
            TabIndex        =   78
            Top             =   195
            Width           =   2340
         End
         Begin VB.CommandButton PreviousCmd 
            Caption         =   "&Previous"
            Height          =   400
            Left            =   2445
            TabIndex        =   77
            Top             =   600
            Width           =   2340
         End
         Begin VB.CommandButton EditCmd 
            Caption         =   "&Edit"
            Height          =   400
            Left            =   4785
            TabIndex        =   76
            Top             =   195
            Width           =   2340
         End
         Begin VB.CommandButton FirstCmd 
            Caption         =   "&First"
            Height          =   400
            Left            =   4785
            TabIndex        =   75
            Top             =   600
            Width           =   2340
         End
         Begin VB.CommandButton DeleteCmd 
            Caption         =   "&Delete"
            Height          =   400
            Left            =   7125
            TabIndex        =   74
            Top             =   195
            Width           =   2340
         End
         Begin VB.CommandButton LastCmd 
            Caption         =   "&Last"
            Height          =   400
            Left            =   7125
            TabIndex        =   73
            Top             =   600
            Width           =   2340
         End
         Begin VB.CommandButton SearchCmd 
            Caption         =   "Search"
            Height          =   400
            Left            =   9465
            TabIndex        =   72
            Top             =   195
            Width           =   2340
         End
         Begin VB.CommandButton ExitCmd 
            Caption         =   "E&xit"
            Height          =   400
            Left            =   9465
            TabIndex        =   71
            Top             =   600
            Width           =   2340
         End
      End
      Begin VB.Frame Frame1 
         Height          =   8445
         Left            =   120
         TabIndex        =   34
         Top             =   120
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
            Left            =   6720
            Sorted          =   -1  'True
            TabIndex        =   2
            Top             =   1020
            Width           =   2355
         End
         Begin VB.TextBox TxtFranchInvNo 
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
            Left            =   10710
            MaxLength       =   20
            TabIndex        =   85
            Top             =   1425
            Width           =   1350
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
            Height          =   360
            Left            =   1845
            Locked          =   -1  'True
            TabIndex        =   83
            TabStop         =   0   'False
            Top             =   1020
            Width           =   3015
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
            ItemData        =   "FrmTxn_FranchiseeRentApproval.frx":0044
            Left            =   10695
            List            =   "FrmTxn_FranchiseeRentApproval.frx":0046
            TabIndex        =   3
            Top             =   1020
            Visible         =   0   'False
            Width           =   3930
         End
         Begin VB.TextBox TxtActualFee 
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
            Left            =   13275
            MaxLength       =   18
            TabIndex        =   1
            Top             =   615
            Width           =   1350
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
            Left            =   1845
            Locked          =   -1  'True
            MaxLength       =   200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   7
            Top             =   2640
            Width           =   12735
         End
         Begin VB.TextBox TxtServiceTaxID 
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
            Left            =   10710
            Locked          =   -1  'True
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   1020
            Width           =   3900
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
            Left            =   1845
            Sorted          =   -1  'True
            TabIndex        =   5
            Top             =   1830
            Width           =   6855
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
            Left            =   8760
            TabIndex        =   40
            Top             =   1860
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
            Left            =   10710
            Locked          =   -1  'True
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   1830
            Width           =   3885
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
            Left            =   10710
            Locked          =   -1  'True
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   615
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
            ItemData        =   "FrmTxn_FranchiseeRentApproval.frx":0048
            Left            =   1845
            List            =   "FrmTxn_FranchiseeRentApproval.frx":005B
            Sorted          =   -1  'True
            TabIndex        =   4
            Top             =   1425
            Width           =   3000
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
            Left            =   13260
            Locked          =   -1  'True
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   195
            Visible         =   0   'False
            Width           =   1350
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
            Height          =   360
            Left            =   1845
            MaxLength       =   200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   6
            Top             =   2235
            Width           =   12720
         End
         Begin VB.TextBox TxtClaimID 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   375
            Left            =   1845
            Locked          =   -1  'True
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   195
            Width           =   1455
         End
         Begin VB.TextBox TxtVendorID 
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
            Left            =   1845
            Locked          =   -1  'True
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   615
            Width           =   7215
         End
         Begin TabDlg.SSTab SSTab1 
            Height          =   5175
            Left            =   120
            TabIndex        =   42
            Top             =   3075
            Width           =   14535
            _ExtentX        =   25638
            _ExtentY        =   9128
            _Version        =   393216
            Tab             =   2
            TabHeight       =   520
            TabCaption(0)   =   "Expense Details"
            TabPicture(0)   =   "FrmTxn_FranchiseeRentApproval.frx":008C
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "CmdExcelDetail"
            Tab(0).Control(1)=   "Frame2"
            Tab(0).ControlCount=   2
            TabCaption(1)   =   "Payment Details"
            TabPicture(1)   =   "FrmTxn_FranchiseeRentApproval.frx":00A8
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Frame3"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "Log History"
            TabPicture(2)   =   "FrmTxn_FranchiseeRentApproval.frx":00C4
            Tab(2).ControlEnabled=   -1  'True
            Tab(2).Control(0)=   "Label24"
            Tab(2).Control(0).Enabled=   0   'False
            Tab(2).Control(1)=   "Label21"
            Tab(2).Control(1).Enabled=   0   'False
            Tab(2).Control(2)=   "Label20"
            Tab(2).Control(2).Enabled=   0   'False
            Tab(2).Control(3)=   "MSHFlexGrid4"
            Tab(2).Control(3).Enabled=   0   'False
            Tab(2).Control(4)=   "TxtUpdatedByFAG"
            Tab(2).Control(4).Enabled=   0   'False
            Tab(2).Control(5)=   "TxtCreated"
            Tab(2).Control(5).Enabled=   0   'False
            Tab(2).Control(6)=   "TxtUpdated"
            Tab(2).Control(6).Enabled=   0   'False
            Tab(2).ControlCount=   7
            Begin VB.CommandButton CmdExcelDetail 
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
               Height          =   240
               Left            =   -71400
               TabIndex        =   91
               Top             =   120
               Width           =   1125
            End
            Begin VB.Frame Frame3 
               Height          =   4695
               Left            =   -74880
               TabIndex        =   49
               Top             =   375
               Width           =   14295
               Begin VB.Frame FrmFAG 
                  Height          =   3240
                  Left            =   120
                  TabIndex        =   102
                  Top             =   1380
                  Width           =   14115
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
                     ItemData        =   "FrmTxn_FranchiseeRentApproval.frx":00E0
                     Left            =   7125
                     List            =   "FrmTxn_FranchiseeRentApproval.frx":00E2
                     TabIndex        =   112
                     Text            =   "CmbFAGBank"
                     Top             =   810
                     Width           =   6855
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
                     TabIndex        =   117
                     Top             =   810
                     Width           =   6855
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
                     Height          =   615
                     Left            =   8850
                     MaxLength       =   1000
                     MultiLine       =   -1  'True
                     ScrollBars      =   2  'Vertical
                     TabIndex        =   115
                     Top             =   1275
                     Width           =   5160
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
                     Height          =   615
                     Left            =   2820
                     MaxLength       =   250
                     MultiLine       =   -1  'True
                     ScrollBars      =   2  'Vertical
                     TabIndex        =   114
                     Top             =   1275
                     Width           =   4800
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
                     TabIndex        =   113
                     Top             =   390
                     Width           =   1785
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
                     ItemData        =   "FrmTxn_FranchiseeRentApproval.frx":00E4
                     Left            =   2820
                     List            =   "FrmTxn_FranchiseeRentApproval.frx":00F4
                     TabIndex        =   111
                     Top             =   810
                     Width           =   1815
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
                     TabIndex        =   110
                     TabStop         =   0   'False
                     Top             =   1950
                     Width           =   4815
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
                     TabIndex        =   109
                     TabStop         =   0   'False
                     Top             =   1950
                     Width           =   1860
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
                     TabIndex        =   108
                     TabStop         =   0   'False
                     Top             =   1950
                     Width           =   1785
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
                     Locked          =   -1  'True
                     TabIndex        =   107
                     TabStop         =   0   'False
                     Top             =   2355
                     Width           =   5175
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
                     Locked          =   -1  'True
                     TabIndex        =   106
                     TabStop         =   0   'False
                     Top             =   2775
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
                     Locked          =   -1  'True
                     TabIndex        =   105
                     TabStop         =   0   'False
                     Top             =   2775
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
                     Locked          =   -1  'True
                     TabIndex        =   104
                     TabStop         =   0   'False
                     Top             =   2355
                     Width           =   4815
                  End
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
                     Locked          =   -1  'True
                     TabIndex        =   103
                     TabStop         =   0   'False
                     Top             =   2775
                     Width           =   1815
                  End
                  Begin MSComCtl2.DTPicker TxtFagDisbDate 
                     Height          =   375
                     Left            =   7125
                     TabIndex        =   116
                     Top             =   390
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
                     Format          =   118882305
                     CurrentDate     =   40255
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
                     TabIndex        =   118
                     Top             =   810
                     Width           =   1800
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
                     TabIndex        =   133
                     Top             =   1350
                     Width           =   1530
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
                     TabIndex        =   132
                     Top             =   120
                     Width           =   2340
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
                     TabIndex        =   131
                     Top             =   1395
                     Width           =   2010
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
                     TabIndex        =   130
                     Top             =   450
                     Width           =   2370
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
                     TabIndex        =   129
                     Top             =   450
                     Width           =   1725
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
                     TabIndex        =   128
                     Top             =   870
                     Width           =   1305
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
                     TabIndex        =   127
                     Top             =   870
                     Width           =   1575
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
                     TabIndex        =   126
                     Top             =   2010
                     Width           =   690
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
                     TabIndex        =   125
                     Top             =   2010
                     Width           =   765
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
                     TabIndex        =   124
                     Top             =   2010
                     Width           =   1290
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
                     TabIndex        =   123
                     Top             =   2445
                     Width           =   1140
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
                     Top             =   2400
                     Width           =   1140
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
                     TabIndex        =   121
                     Top             =   2820
                     Width           =   2445
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
                     TabIndex        =   120
                     Top             =   2835
                     Width           =   900
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
                     TabIndex        =   119
                     Top             =   2820
                     Width           =   900
                  End
               End
               Begin VB.Frame Frame5 
                  Height          =   1290
                  Left            =   120
                  TabIndex        =   92
                  Top             =   100
                  Width           =   14025
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
                     TabIndex        =   96
                     Top             =   825
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
                     TabIndex        =   95
                     Top             =   825
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
                     TabIndex        =   94
                     Top             =   420
                     Width           =   5145
                  End
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
                     TabIndex        =   93
                     Top             =   420
                     Width           =   5625
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
                     TabIndex        =   101
                     Top             =   120
                     Width           =   2760
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
                     TabIndex        =   100
                     Top             =   825
                     Width           =   975
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
                     TabIndex        =   99
                     Top             =   885
                     Width           =   645
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
                     TabIndex        =   98
                     Top             =   480
                     Width           =   1230
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
                     TabIndex        =   97
                     Top             =   480
                     Width           =   1065
                  End
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
               Left            =   2385
               Locked          =   -1  'True
               TabIndex        =   48
               TabStop         =   0   'False
               Top             =   900
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
               Left            =   2385
               Locked          =   -1  'True
               TabIndex        =   47
               TabStop         =   0   'False
               Top             =   480
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
               Left            =   2385
               Locked          =   -1  'True
               TabIndex        =   46
               TabStop         =   0   'False
               Top             =   1320
               Width           =   9855
            End
            Begin VB.Frame Frame2 
               Height          =   4740
               Left            =   -74925
               TabIndex        =   43
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
                  TabIndex        =   44
                  TabStop         =   0   'False
                  Top             =   960
                  Visible         =   0   'False
                  Width           =   2175
               End
               Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
                  Height          =   4410
                  Left            =   120
                  TabIndex        =   45
                  Top             =   240
                  Width           =   14175
                  _ExtentX        =   25003
                  _ExtentY        =   7779
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
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid4 
               Height          =   3015
               Left            =   120
               TabIndex        =   50
               Top             =   1800
               Width           =   14100
               _ExtentX        =   24871
               _ExtentY        =   5318
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
               Left            =   120
               TabIndex        =   53
               Top             =   975
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
               Left            =   120
               TabIndex        =   52
               Top             =   540
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
               Left            =   120
               TabIndex        =   51
               Top             =   1380
               Width           =   2175
            End
         End
         Begin MSComCtl2.DTPicker TxtClaimDate 
            Height          =   375
            Left            =   4560
            TabIndex        =   9
            Top             =   195
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Format          =   118882305
            CurrentDate     =   39884
         End
         Begin MSComCtl2.DTPicker TxtExpenseFrom 
            Height          =   375
            Left            =   7245
            TabIndex        =   10
            Top             =   195
            Width           =   1815
            _ExtentX        =   3201
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
            Format          =   118882305
            CurrentDate     =   39884
         End
         Begin MSComCtl2.DTPicker TxtExpenseTo 
            Height          =   375
            Left            =   10710
            TabIndex        =   11
            Top             =   195
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
            Format          =   118882305
            CurrentDate     =   39884
         End
         Begin MSComCtl2.DTPicker TxtApprovalDate 
            Height          =   375
            Left            =   7200
            TabIndex        =   8
            Top             =   1425
            Visible         =   0   'False
            Width           =   1900
            _ExtentX        =   3360
            _ExtentY        =   661
            _Version        =   393216
            Format          =   118882305
            CurrentDate     =   39884
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
            Left            =   1845
            Locked          =   -1  'True
            TabIndex        =   54
            TabStop         =   0   'False
            Top             =   1425
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
            Left            =   1845
            Locked          =   -1  'True
            TabIndex        =   55
            TabStop         =   0   'False
            Top             =   1830
            Width           =   6855
         End
         Begin MSComCtl2.DTPicker TxtFranchInvDate 
            Height          =   375
            Left            =   13200
            TabIndex        =   88
            Top             =   1440
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
            Format          =   118882305
            CurrentDate     =   39884
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
            Left            =   6720
            Locked          =   -1  'True
            TabIndex        =   135
            TabStop         =   0   'False
            Top             =   1020
            Width           =   2350
         End
         Begin VB.Label Label25 
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
            Left            =   5040
            TabIndex        =   134
            Top             =   953
            Width           =   1725
         End
         Begin VB.Label Label13 
            Caption         =   "Invoice Date"
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
            Left            =   12210
            TabIndex        =   87
            Top             =   1335
            Width           =   1050
         End
         Begin VB.Label Label11 
            Caption         =   "Franch Invoice No"
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
            Left            =   9240
            TabIndex        =   86
            Top             =   1395
            Width           =   1065
         End
         Begin VB.Label Label10 
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
            TabIndex        =   84
            Top             =   1080
            Width           =   960
         End
         Begin VB.Label Label9 
            Caption         =   "Actual Fee"
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
            Left            =   12240
            TabIndex        =   82
            Top             =   667
            Width           =   990
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
            Left            =   150
            TabIndex        =   69
            Top             =   2715
            Width           =   1350
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
            Left            =   9240
            TabIndex        =   68
            Top             =   1050
            Width           =   1410
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
            Height          =   270
            Left            =   135
            TabIndex        =   67
            Top             =   1875
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
            Left            =   9240
            TabIndex        =   66
            Top             =   1867
            Width           =   1065
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
            TabIndex        =   65
            Top             =   1485
            Visible         =   0   'False
            Width           =   1980
         End
         Begin VB.Label Label18 
            Caption         =   "Fee To"
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
            Left            =   9240
            TabIndex        =   64
            Top             =   247
            Width           =   1095
         End
         Begin VB.Label Label17 
            Caption         =   "Fee From "
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
            TabIndex        =   63
            Top             =   247
            Width           =   1095
         End
         Begin VB.Label Label15 
            Caption         =   "Fee Amount"
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
            Left            =   9240
            TabIndex        =   62
            Top             =   667
            Width           =   1260
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
            TabIndex        =   61
            Top             =   1478
            Width           =   960
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
            Left            =   12240
            TabIndex        =   60
            Top             =   240
            Visible         =   0   'False
            Width           =   630
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
            TabIndex        =   59
            Top             =   2325
            Width           =   1080
         End
         Begin VB.Label Label4 
            Caption         =   "Franchisee Name"
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
            TabIndex        =   58
            Top             =   660
            Width           =   1785
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
            Left            =   3390
            TabIndex        =   57
            Top             =   247
            Width           =   1095
         End
         Begin VB.Label LbClaimID 
            Caption         =   "Fee ID"
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
            TabIndex        =   56
            Top             =   277
            Width           =   900
         End
      End
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   9885
      Width           =   10935
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   10095
      Left            =   15000
      TabIndex        =   0
      Top             =   120
      Width           =   255
   End
End
Attribute VB_Name = "FrmTxn_FranchiseeRentApproval"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg, mExpenseTypeID, mVendorID, mSVendorID, mStatus, mSStatus, mLessorEmailID As Variant
Dim mLocationID, mSLocationID As Double
Dim mBatchID As Double
Dim mForwardToID, mForwardToEmailID, mSForwardToID, mEmployeeEmailID As Variant
Dim oSMTP As SMTPSession
Dim mSelRow As Variant
Dim mSelCol As Variant
Dim mServiceTaxID As Variant
Dim mAllowEdit As Boolean

Dim mFAGBankID, mPBatchID, mSaveStatus As Variant

Dim mBillingStateID As Double   ''-- Ullhas Patil 25 Oct 2017

Private Sub AddCmd_Click()
''--- Ullhas Patil 25 Oct 2017
CmbBillingState.Visible = True: CmbBillingState = ""
TxtBillingState.Visible = False: TxtBillingState = ""
''-----------------------------
End Sub

''-- Ullhas Patil 25 Oct 2017
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
mSVendorID = ""
Call TableMst_Vendor("")
CmbSLessor.Clear
If RsMst_Vendor.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Vendor.RecordCount
    CmbSLessor.AddItem RsMst_Vendor!VendorName
    RsMst_Vendor.MoveNext
Next
CmbSLessor.Text = tmp
End Sub

Private Sub CmbSLessor_LostFocus()
If CmbSLessor.Text = "" Then
   Exit Sub
End If
RsMst_Vendor.MoveFirst
RsMst_Vendor.Find "VendorName = '" & CmbSLessor.Text & "'"
If RsMst_Vendor.EOF Then
   MsgBox "Invalid selection "
   CmbSLessor.SetFocus
   Exit Sub
End If
mSVendorID = RsMst_Vendor!VendorID
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
ElseIf LCase(CmbSStatus.Text) = "freez" Then
   mSStatus = "F"
''-- 03 Jan 2017 Ullhas Patil
ElseIf LCase(CmbSStatus.Text) = LCase("Settled") Then
   mSStatus = "S"
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

''-- 03 Jan 2017 Ullhas Patil
If mSaveStatus = "A" Then
'If mStatus = "A" Then
   If mBatchID = 0 Then
      If LCase(CmbStatus.Text) = "pending" Then
         MsgBox "This claim is approved already. You can't mark it as " & CmbStatus.Text & " status. !!!"
         CmbStatus.SetFocus
         Exit Sub
      End If
   ElseIf mBatchID <> 0 Then
      If LCase(CmbStatus.Text) = "pending" Or LCase(CmbStatus.Text) = "rejected" Or LCase(CmbStatus.Text) = "freez" Then
         MsgBox "This claim is approved already. You can't mark it as " & CmbStatus.Text & " status. !!!"
         CmbStatus.SetFocus
         Exit Sub
      End If
   End If
'End If
End If
If LCase(CmbStatus.Text) = "approved" Then
   If LCase(Gen_UserRole) <> "fag" And LCase(Gen_UserRole) <> "fagadmin" And LCase(Gen_UserRole) <> "power" Then
      MsgBox "Only FAG can Approved Status !!!"
      CmbStatus.SetFocus
      Exit Sub
   End If
   mStatus = "A"
   TxtApprovalRemarks.Locked = False
   TxtApprovalDate.Visible = True
   Label27.Visible = True
   TxtApprovalDate.Enabled = True
   TxtFAGRemark = ""
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
   TxtFAGRemark = ""
ElseIf LCase(CmbStatus.Text) = "freez" Then
   mStatus = "F"
   TxtApprovalRemarks.Locked = False
   TxtApprovalDate.Visible = True
   Label27.Visible = True
   TxtApprovalDate.Enabled = True
   TxtFAGRemark = ""
''-- 03 Jan 2017 Ullhas Patil
ElseIf LCase(CmbStatus.Text) = LCase("Settled") Then
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
   TxtFAGDisbAmt.Text = Val(TxtActualFee)
   TxtDeductionRemark = "TDS@"
   TxtFAGRemark = Trim(TxtApprovalRemarks)
   Call TableMst_FAGBank(" Where FAGBankID in (" & Gen_FAGBankIDForPayment & ") ")
   If RsMst_FAGBank.RecordCount > 0 Then
      CmbFAGBank.Text = RsMst_FAGBank!BankName & "~" & RsMst_FAGBank!AccountNo
      mFAGBankID = RsMst_FAGBank!FAGBankID
   End If
Else
   MsgBox "Invalid selection!!!"
   CmbStatus.SetFocus
   TxtApprovalRemarks.Locked = True
   TxtApprovalDate.Visible = False
   Label27.Visible = False
   Exit Sub
End If

''-- 03 Jan 2017 Ullhas Patil
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


wc = ""
RsTxn_FranchiseeRent.Filter = ""  'VC TGR

''tmp = "Select TGR.ClaimID,TGR.ClaimDate,TGR.ExpenseFrom,TGR.ExpenseTo, "
''tmp = tmp & "  MV.VendorName,ML.LocationName,TGR.Amount,TGR.ActualFee,EMP.EmployeeName +'~'+ EMP.EmployeeNo as ForwardedTo ,TGR.Flag, "
''tmp = tmp & "  TGR.ApprovalDate,TGR.ApprovalRemarks, "
''tmp = tmp & "  TGR.Remarks,TGR.BatchID,TGR.G_UID,TGR.BankName,TGR.ChequeNo,TGR.PaidAmount,TGR.TDS,TGR.PaidDate "
''tmp = tmp & "  From Txn_FranchiseeRent TGR "
''tmp = tmp & "  Inner Join Mst_Vendor MV On TGR.VendorID = MV.VendorID "
''tmp = tmp & "  Left Join Mst_Location ML On TGR.LocationID = ML.LocationID "
''tmp = tmp & "  Left Outer Join Mst_Employee EMP On TGR.ForwardToID = EMP.EmployeeID "

tmp = " Select TGR.ClaimID,TGR.ClaimDate,TGR.ExpenseFrom,TGR.ExpenseTo,MV.VendorName,ML.LocationName,"
tmp = tmp & " TGR.Amount,TGR.ActualFee,EMP.EmployeeName +'~'+ EMP.EmployeeNo as ForwardedTo ,TGR.Flag,"
tmp = tmp & " TGR.ApprovedDate , TGR.ApprovalRemarks, TGR.Remarks, TGR.BatchID, TGR.G_UID, TGR.BankName, TGR.ChequeNo, TGR.PaidAmount, TGR.TDS, TGR.PaidDate"
tmp = tmp & " ,TGR.DisbuDate,TGR.DisbuRemark,TGR.DisbuAmount,TGR.ModeOfPay,TGR.FAGBankID"
''tmp = tmp & "  --MFB.BankName+'~'+MFB.AccountNo[]"
tmp = tmp & " ,TGR.PBatchID,TGR.DocNo , TGR.EntryNo, TGR.FAGPayExportedDate, TGR.DeductionRemark, TGR.LetterPrintFlag,MSII.StateIndiaName[BillingStateName]"
tmp = tmp & " From Txn_FranchiseeRent TGR"
tmp = tmp & " Inner Join Mst_Vendor MV On TGR.VendorID = MV.VendorID"
tmp = tmp & " Left Join Mst_Location ML On TGR.LocationID = ML.LocationID"
tmp = tmp & " Left Outer Join Mst_Employee EMP On TGR.ForwardToID = EMP.EmployeeID"
tmp = tmp & " Left Join Mst_FAGBank MFB On TGR.FAGBankID = MFB.FAGBankID"
tmp = tmp & " Left join Mst_StateIndia MSII On TGR.BillingStateIndiaID = MSII.StateIndiaID "     ''--- Ullhas Patil 25 Oct 2017

If Gen_DBType = "MDB" Then
   tmp = tmp & " Where ClaimDate Between #" & Format(SFrom, Gen_DatFormat) & "# And #" & Format(STo, Gen_DatFormat) & "#) "
   wc = " ClaimDate >= #" & Format(SFrom, Gen_DatFormat) & "# And ClaimDate <= #" & Format(STo, Gen_DatFormat) & "#) "
Else
   tmp = tmp & " Where ClaimDate Between '" & Format(SFrom, Gen_DatFormat) & "' And '" & Format(STo, Gen_DatFormat) & "' "
   wc = " ClaimDate >= '" & Format(SFrom, Gen_DatFormat) & "' And ClaimDate <= '" & Format(STo, Gen_DatFormat) & "' "
End If

If CmbSLessor.Text <> "" Then
   tmp = tmp & " AND MV.VendorName = '" & CmbSLessor.Text & "' "
   wc = wc & " And VendorID = " & mSVendorID & ""
End If

If CmbSLocation.Text <> "" Then
   tmp = tmp & " AND ML.LocationName = '" & CmbSLocation.Text & "' "
   wc = wc & " And LocationID = " & mSLocationID & ""
Else
   If Gen_WcLocation <> "" Then
      tmp = tmp & " And ML." & Gen_WcLocation
   End If
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

Call TmpTable

'Here Filter is used to show records in indata which are searched
RsTxn_FranchiseeRent.Filter = "" & wc & ""

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

    If LCase(MSHFlexGrid1.TextMatrix(iu, 9)) = "a" Then
       MSHFlexGrid1.TextMatrix(iu, 9) = "Approved"
    ElseIf LCase(MSHFlexGrid1.TextMatrix(iu, 9)) = "p" Then
       MSHFlexGrid1.TextMatrix(iu, 9) = "Pending"
    ElseIf LCase(MSHFlexGrid1.TextMatrix(iu, 9)) = "r" Then
       MSHFlexGrid1.TextMatrix(iu, 9) = "Rejected"
    ElseIf LCase(MSHFlexGrid1.TextMatrix(iu, 9)) = "f" Then
       MSHFlexGrid1.TextMatrix(iu, 9) = "Freez"
    ''-- 03rd Jan 2017 Ullhas Patil
    ElseIf LCase(MSHFlexGrid1.TextMatrix(iu, 9)) = "s" Then
       MSHFlexGrid1.TextMatrix(iu, 9) = "Settled"
    End If
    
    If Val(MSHFlexGrid1.TextMatrix(iu, 24)) > 0 Then
        MSHFlexGrid1.TextMatrix(iu, 24) = GetFAGBankAcNo(MSHFlexGrid1.TextMatrix(iu, 24))
    End If
    
    ''-- 03rd Jan 2017 Ullhas Patil
    If LCase(MSHFlexGrid1.TextMatrix(iu, 30)) = "n" Then
       MSHFlexGrid1.TextMatrix(iu, 30) = "No"
    ElseIf LCase(MSHFlexGrid1.TextMatrix(iu, 30)) = "y" Then
       MSHFlexGrid1.TextMatrix(iu, 30) = "Yes"
    End If
    
    TmpRs.MoveNext
Next

Label22.Refresh
Label22.Caption = TmpRs.RecordCount

End Sub

Private Sub CmdPrint_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "FranchiseeRent_" & TxtClaimID & ".xls")
Call Xls_Detail_Rpt
End Sub
'Public Sub Xls_Detail_Rpt()
'
'Dim oXL As Excel.Application
'Dim oWB As Excel.Workbook
'Dim oWS As Excel.Worksheet
'Dim mRow, mRow1 As Integer
'Dim mCol As Integer
'Dim c As Variant
'Set oXL = New Excel.Application
'Dim Pat As String
'Dim mAmt, mSTAmt, mTotAmt As Double
'Pat = App.Path
'
'Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "FranchiseeRent_" & TxtClaimID & ".xls")
'Set oWS = oWB.Worksheets("Sheet1")
'oWS.ClearArrows
'
'oWS.Shapes.AddPicture(App.Path & "\logo.gif", msoTrue, msoTrue, 10, 10, 40, 40).Select
'
'oWS.Range("d:d").ColumnWidth = 12.86
'oWS.Range("e:e").ColumnWidth = 18.86
'
'tmp = "B" & 1 & ":f" & 1
'oWS.Range(tmp).Merge
'oWS.Cells(1, 2) = "National Collateral Management Services Limited"
'oWS.Cells(1, 2).HorizontalAlignment = -4108
'oWS.Range(tmp).Font.Bold = True
'
'tmp = "B" & 2 & ":f" & 2
'oWS.Range(tmp).Merge
'oWS.Cells(2, 2) = "Gayatri Towers,954,Appasaheb Marathe Marg,"
'oWS.Cells(2, 2).HorizontalAlignment = -4108
'
'tmp = "B" & 3 & ":f" & 3
'oWS.Range(tmp).Merge
'oWS.Cells(3, 2) = "Prabhadevi, Mumbai-400025"
'oWS.Cells(3, 2).HorizontalAlignment = -4108
'
'tmp = "B" & 4 & ":f" & 4
'oWS.Range(tmp).Merge
'oWS.Cells(4, 2) = "Tel.No-022 40419191, Fax. No-022 40419193"
'oWS.Cells(4, 2).HorizontalAlignment = -4108
'
'tmp = "a" & 7 & ":f" & 7
'oWS.Range(tmp).Merge
'oWS.Cells(7, 1) = ""
'oWS.Cells(8, 1) = "To :"
'
'tmp = "a" & 9 & ":f" & 9
'oWS.Range(tmp).Merge
'oWS.Cells(9, 1) = "      " & TxtVendorID.Text
'oWS.Cells(9, 1).Font.Bold = True
'tmp = "a" & 10 & ":d" & 10
'oWS.Range(tmp).Merge
'oWS.Cells(10, 1) = "" & GetVendorName(mVendorID, "A") '""  'Address
'oWS.Range(tmp).WrapText = True
'oWS.Cells(10, 1).RowHeight = 45
'oWS.Cells(11, 1) = "Dear Sir,"
'tmp = "a" & 13 & ":f" & 14
'oWS.Range(tmp).Merge
'oWS.Range(tmp).WrapText = True
'oWS.Cells(13, 1) = "Please find enclosed Rent paid details for the month " & _
'                   " of " & MonthName(Month(TxtExpenseFrom.Value)) & " " & Year(TxtExpenseFrom.Value) & ". The details are given below."
'
'''oWS.Cells(16, 1) = "Agreement ID"
'''oWS.Cells(16, 2) = "Location"
'''oWS.Cells(16, 3) = "Godown No."
'''oWS.Cells(16, 4) = "Area in SqFt"
'''oWS.Cells(16, 5) = "Agree. Expiry Date"
'''oWS.Cells(16, 6) = "Gross Amount"
'''tmp = "a" & 16 & ":f" & 16
'''oWS.Range(tmp).Font.Bold = True
'
'''tmp = "Select TLA.LAID,ML.LocationName,MG.GodownNo,MG.Area,TLA.ExpiryDate,TGRD.Amount"
'''tmp = tmp & " from Txn_FranchiseeRent TGR"
'''tmp = tmp & " Inner Join Txn_FranchiseeRentDetail TGRD On TGR.ClaimID = TGRD.ClaimID"
'''tmp = tmp & " Inner Join Mst_Godown MG On MG.GodownID=TGRD.GodownID"
'''tmp = tmp & " Inner Join Mst_CMP MC On MC.CMPID = MG.CMPID"
'''tmp = tmp & " inner join Mst_Location ML on MC.LocationID=ML.LocationID"
'''tmp = tmp & " Inner Join Txn_LessorAgreement TLA on TGRD.LAID=TLA.LAID"
'''tmp = tmp & " Where TGR.ClaimID=" & Val(TxtClaimID)
'
'''Call TmpTable
'
'''tmp = "Select PANNo,RTGSNo From Mst_Lessor Where VendorID=" & mVendorID
'''Call Tmp1Table
'
'mRow = 17
'
'mAmt = 0
'mSTAmt = 0
'mTotAmt = 0
'
'   For I = 1 To MSHFlexGrid2.Rows - 2
'       c = 1
'       For Col = 0 To MSHFlexGrid2.Cols - 1
'           If InStr(MSHFlexGrid2.TextMatrix(I, Col), "/") > 0 Then
'              If IsDate(MSHFlexGrid2.TextMatrix(I, Col)) Then
'                 oWS.Cells(mRow, c) = Format(MSHFlexGrid2.TextMatrix(I, Col), "dd-MMM-yyyy")
'              Else
'                 oWS.Cells(mRow, c) = MSHFlexGrid2.TextMatrix(I, Col)
'              End If
'           Else
'              oWS.Cells(mRow, c) = MSHFlexGrid2.TextMatrix(I, Col)
'           End If
'           c = c + 1
'       Next
'       mRow = mRow + 1
'   Next
'
'tmp = "a" & 16 & ":f" & mRow
''oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
'
'tmp = "a" & mRow & ":c" & mRow
'oWS.Range(tmp).Merge
'oWS.Range(tmp).Font.Bold = True
'oWS.Cells(mRow, 1) = "Total"
'oWS.Cells(mRow, 4) = Format(mAmt, "#####0.00")
'oWS.Cells(mRow, 6) = Format(mTotAmt, "#####0.00")
'
'tmp = "d" & mRow & ":f" & mRow
'oWS.Range(tmp).Font.Bold = True
'
'tmp = "b" & mRow + 2 & ":b" & mRow + 9
'oWS.Range(tmp).Font.Bold = True
'
'tmp = "b" & mRow + 2 & ":e" & mRow + 9
'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
'
'tmp = "b" & mRow + 2 & ":c" & mRow + 2
'oWS.Range(tmp).Merge
'oWS.Range(tmp).WrapText = True
'oWS.Cells(mRow + 2, 2) = "Bank Name"
'
'tmp = "d" & mRow + 2 & ":e" & mRow + 2
'oWS.Range(tmp).Merge
'oWS.Range(tmp).WrapText = True
'oWS.Cells(mRow + 2, 4) = TxtBankID.Text
'
'tmp = "b" & mRow + 3 & ":c" & mRow + 3
'oWS.Range(tmp).Merge
'oWS.Range(tmp).WrapText = True
'oWS.Cells(mRow + 3, 2) = "Paid Amount"
'
'tmp = "d" & mRow + 3 & ":e" & mRow + 3
'oWS.Range(tmp).Merge
'oWS.Range(tmp).WrapText = True
'oWS.Cells(mRow + 3, 4) = Format(Val(TxtPaidAmt.Text), "#####0.00")
'
'tmp = "b" & mRow + 4 & ":c" & mRow + 4
'oWS.Range(tmp).Merge
'oWS.Range(tmp).WrapText = True
'oWS.Cells(mRow + 4, 2) = "TDS"
'
'tmp = "d" & mRow + 4 & ":e" & mRow + 4
'oWS.Range(tmp).Merge
'oWS.Range(tmp).WrapText = True
'oWS.Cells(mRow + 4, 4) = Format(Val(TxtTDS.Text), "#####0.00")
'
'tmp = "b" & mRow + 5 & ":c" & mRow + 5
'oWS.Range(tmp).Merge
'oWS.Range(tmp).WrapText = True
'oWS.Cells(mRow + 5, 2) = "Net Amount"
'
'tmp = "d" & mRow + 5 & ":e" & mRow + 5
'oWS.Range(tmp).Merge
'oWS.Range(tmp).WrapText = True
'oWS.Cells(mRow + 5, 4) = Format(Val(TxtPaidAmt) + Val(TxtTDS), "#####0.00")
'
'tmp = "b" & mRow + 6 & ":c" & mRow + 6
'oWS.Range(tmp).Merge
'oWS.Range(tmp).WrapText = True
'oWS.Cells(mRow + 6, 2) = "Cheque No."
'
'tmp = "d" & mRow + 6 & ":e" & mRow + 6
'oWS.Range(tmp).Merge
'oWS.Range(tmp).WrapText = True
'oWS.Cells(mRow + 6, 4) = "'" & TxtChequeNo.Text
'
'tmp = "b" & mRow + 7 & ":c" & mRow + 7
'oWS.Range(tmp).Merge
'oWS.Range(tmp).WrapText = True
'oWS.Cells(mRow + 7, 2) = "Paid Date"
'
'tmp = "d" & mRow + 7 & ":e" & mRow + 7
'oWS.Range(tmp).Merge
'oWS.Range(tmp).WrapText = True
'oWS.Cells(mRow + 7, 4) = "'" & TxtPaidDate.Text
'
'tmp = "b" & mRow + 8 & ":c" & mRow + 8
'oWS.Range(tmp).Merge
'oWS.Range(tmp).WrapText = True
'oWS.Cells(mRow + 8, 2) = "RTGS Details"
'
'''tmp = "d" & mRow + 8 & ":e" & mRow + 8
'''oWS.Range(tmp).Merge
'''oWS.Range(tmp).WrapText = True
'''If LCase(TxtChequeNo) = "trf" Then
'''   oWS.Cells(mRow + 8, 4) = "'" & IIf(IsNull(TmpRs1!RTGSNo), "", TmpRs1!RTGSNo)
'''Else
'''   oWS.Cells(mRow + 8, 4) = ""
'''End If
'''tmp = "b" & mRow + 9 & ":c" & mRow + 9
'''oWS.Range(tmp).Merge
'''oWS.Range(tmp).WrapText = True
'''oWS.Cells(mRow + 9, 2) = "PAN Card No."
'''
'''tmp = "d" & mRow + 9 & ":e" & mRow + 9
'''oWS.Range(tmp).Merge
'''oWS.Range(tmp).WrapText = True
'''oWS.Cells(mRow + 9, 4) = "'" & IIf(IsNull(TmpRs1!PANNo), "", TmpRs1!PANNo)
'
'tmp = "a" & mRow + 12 & ":f" & mRow + 12
'oWS.Range(tmp).Merge
'oWS.Cells(mRow + 12, 1) = "For National Collateral Management Services Ltd."
'
'tmp = "a" & mRow + 15 & ":f" & mRow + 15
'oWS.Range(tmp).Merge
'oWS.Cells(mRow + 15, 1) = "(Authorised Signatory)"
'
'If Gen_ExcelPassword <> "" Then
'   oWS.Protect Gen_ExcelPassword, , , , , , True, True
'End If
'
'oWB.Save
'MsgBox ("Done!!")
'oXL.Visible = True
'
'End Sub

Public Sub Xls_Detail_Rpt()

Dim ans As Variant
MsgBox ("Wait till Next Message")
Set oXL = New Excel.Application
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "FranchiseeRent_" & TxtClaimID & ".xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.Name = "Bill Voucher"
oWS.ClearArrows
mRow = 1

oWS.Range("a:a").ColumnWidth = 11
oWS.Range("b:b").ColumnWidth = 11.43
oWS.Range("c:c").ColumnWidth = 10.86
oWS.Range("d:d").ColumnWidth = 8.29
oWS.Range("e:e").ColumnWidth = 7.43
oWS.Range("f:f").ColumnWidth = 7.71
oWS.Range("g:g").ColumnWidth = 7
oWS.Range("h:h").ColumnWidth = 10.57
oWS.Range("i:i").ColumnWidth = 10.14
  
tmp = "a" & mRow & ":" & "i" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "NATIONAL COLLATERAL MANAGEMENT SERVICES LIMITED"
oWS.Cells(mRow, C + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 10

mRow = mRow + 1
tmp = "a" & mRow & ":" & "i" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "Franchises BILL VOUCHER"
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 10
oWS.Cells(mRow, C + 1).HorizontalAlignment = -4108

mRow = mRow + 1
tmp = "a" & mRow & ":" & "b" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "Business Segment :"
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 10

tmp = "c" & mRow & ":" & "d" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 3) = "CMP"
oWS.Cells(mRow, C + 3).Font.Size = 10
oWS.Cells(mRow, C + 3).HorizontalAlignment = -4108

tmp = "e" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 5) = "Store-Man ID"
oWS.Cells(mRow, C + 5).Font.Bold = True
oWS.Cells(mRow, C + 5).Font.Size = 10

tmp = "g" & mRow & ":" & "g" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 7) = "'" & TxtClaimID
oWS.Cells(mRow, C + 7).Font.Bold = True
oWS.Cells(mRow, C + 7).Font.Size = 10
   
tmp = "h" & mRow & ":" & "h" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 8) = "Claim Date"
oWS.Cells(mRow, C + 8).Font.Bold = True
oWS.Cells(mRow, C + 8).Font.Size = 10
   
tmp = "i" & mRow & ":" & "i" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 9) = "'" & Format(TxtClaimDate, "dd-mmm-yyyy")
oWS.Cells(mRow, C + 9).Font.Bold = True
oWS.Cells(mRow, C + 9).Font.Size = 10
   
mRow = mRow + 1
tmp = "a" & mRow & ":" & "b" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "Franchise Vendor Name :"
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 10

tmp = "c" & mRow & ":" & "i" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 3) = TxtVendorID
oWS.Cells(mRow, C + 3).Font.Bold = True
oWS.Cells(mRow, C + 3).Font.Size = 10
   
mRow = mRow + 1
tmp = "a" & mRow & ":" & "b" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "Location:"
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 10

tmp = "c" & mRow & ":" & "i" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 3) = TxtLocationID
oWS.Cells(mRow, C + 3).Font.Bold = True
oWS.Cells(mRow, C + 3).Font.Size = 10

mRow = mRow + 1
tmp = "a" & mRow & ":" & "b" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "Month/Year"
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 10

tmp = "c" & mRow & ":" & "i" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 3) = "'" & Format(TxtExpenseFrom, "MMM-yyyy")
oWS.Cells(mRow, C + 3).Font.Bold = True
oWS.Cells(mRow, C + 3).Font.Size = 10

mRow = mRow + 1
tmp = "a" & mRow & ":" & "b" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "Invoice No"
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 10

tmp = "c" & mRow & ":" & "d" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 3) = TxtFranchInvNo
oWS.Cells(mRow, C + 3).Font.Size = 10

tmp = "e" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 5) = "Invoice Date"
oWS.Cells(mRow, C + 5).Font.Bold = True
oWS.Cells(mRow, C + 5).Font.Size = 10

tmp = "g" & mRow & ":" & "h" & mRow
oWS.Range(tmp).Merge
If TxtFranchInvNo <> "" Then
   oWS.Cells(mRow, C + 7) = "'" & Format(TxtFranchInvDate, "dd-MMM-yyyy")
End If
oWS.Cells(mRow, C + 7).Font.Size = 10

mRow = mRow + 1
tmp = "a" & mRow & ":" & "i" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "Particulars of Franchisee Fees Expenditure : "
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 10

'mRow = mRow + 1
'tmp = "a" & mRow
'oWS.Range(tmp).Merge
'oWS.Cells(mRow, c + 1) = "CMP Name"
'oWS.Cells(mRow, c + 1).Font.Bold = True
'oWS.Cells(mRow, c + 1).Font.Size = 10
'oWS.Cells(mRow, c + 1).RowHeight = 27.75
'oWS.Cells(mRow, c + 1).WrapText = True
'
'tmp = "b" & mRow
'oWS.Range(tmp).Merge
'oWS.Cells(mRow, c + 2) = "GSL"
'oWS.Cells(mRow, c + 2).Font.Bold = True
'oWS.Cells(mRow, c + 2).Font.Size = 10
'oWS.Cells(mRow, c + 2).RowHeight = 27.75
'oWS.Cells(mRow, c + 2).WrapText = True
'
'tmp = "c" & mRow & ":" & "f" & mRow
'oWS.Range(tmp).Merge
'oWS.Cells(mRow, c + 3) = "Expense Type"
'oWS.Cells(mRow, c + 3).Font.Bold = True
'oWS.Cells(mRow, c + 3).Font.Size = 10
'oWS.Cells(mRow, c + 3).RowHeight = 27.75
'
'tmp = "g" & mRow & ":" & "h" & mRow
'oWS.Range(tmp).Merge
'oWS.Cells(mRow, c + 7) = "Particulars"
'oWS.Cells(mRow, c + 7).Font.Bold = True
'oWS.Cells(mRow, c + 7).Font.Size = 10
'oWS.Cells(mRow, c + 7).RowHeight = 27.75
'
'tmp = "i" & mRow
'oWS.Range(tmp).Merge
'oWS.Cells(mRow, c + 9) = "Amount (Without ST)"
'oWS.Cells(mRow, c + 9).Font.Bold = True
'oWS.Cells(mRow, c + 9).Font.Size = 10
'oWS.Cells(mRow, c + 9).RowHeight = 27.75
'oWS.Cells(mRow, c + 9).WrapText = True
'
'For I = 2 To MSHFlexGrid2.Rows - 2
'   mRow = mRow + 1
'   tmp = "a" & mRow
'   oWS.Range(tmp).Merge
'   oWS.Cells(mRow, c + 1) = MSHFlexGrid2.TextMatrix(I, 0)
'   oWS.Cells(mRow, c + 1).Font.Size = 10
''   oWS.Cells(mRow, c + 1).RowHeight = 30
'   oWS.Cells(mRow, c + 1).WrapText = True
'
'   tmp = "b" & mRow
'   oWS.Range(tmp).Merge
'   oWS.Cells(mRow, c + 2) = MSHFlexGrid2.TextMatrix(I, 2) & "~" & MSHFlexGrid2.TextMatrix(I, 3) & "~" & MSHFlexGrid2.TextMatrix(I, 4)
'   oWS.Cells(mRow, c + 2).Font.Size = 10
' '  oWS.Cells(mRow, c + 2).RowHeight = 30
'   oWS.Cells(mRow, c + 2).WrapText = True
'
'   tmp = "c" & mRow & ":" & "f" & mRow
'   oWS.Range(tmp).Merge
'   oWS.Cells(mRow, c + 3) = "Franchisee Fees"
'   'oWS.Cells(mRow, c + 3).Font.Bold = True
'   oWS.Cells(mRow, c + 3).Font.Size = 10
'   ''oWS.Cells(mRow, c + 3).RowHeight = 27.75
'
'   tmp = "g" & mRow & ":" & "h" & mRow
'   oWS.Range(tmp).Merge
'   oWS.Cells(mRow, c + 7) = "Particulars"
'   'oWS.Cells(mRow, c + 4).Font.Bold = True
'   oWS.Cells(mRow, c + 7).Font.Size = 10
'   'oWS.Cells(mRow, c + 4).RowHeight = 27.75
'
'   tmp = "i" & mRow
'   oWS.Range(tmp).Merge
'   oWS.Cells(mRow, c + 9) = MSHFlexGrid2.TextMatrix(I, 14)
'   'oWS.Cells(mRow, c + 9).Font.Bold = True
'   oWS.Cells(mRow, c + 9).Font.Size = 10
'   'oWS.Cells(mRow, c + 9).RowHeight = 27.75
'   oWS.Cells(mRow, c + 9).WrapText = True
'Next I

mRow = mRow + 1
tmp = "a" & mRow & ":" & "h" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "Total Amount"
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 10
oWS.Cells(mRow, C + 1).HorizontalAlignment = -4108

tmp = "i" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 9) = Format(TxtActualFee, "#####0.00") 'Format(MSHFlexGrid2.TextMatrix(0, 14), "#####0.00")
oWS.Cells(mRow, C + 9).Font.Bold = True
oWS.Cells(mRow, C + 9).Font.Size = 10
    
mRow = mRow + 1
tmp = "a" & mRow & ":" & "h" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "ST Amount"
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 10
oWS.Cells(mRow, C + 1).HorizontalAlignment = -4108

tmp = "i" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 9) = 0 'Format(TxtSTAmount.Text, "#####0.00")
oWS.Cells(mRow, C + 9).Font.Bold = True
oWS.Cells(mRow, C + 9).Font.Size = 10
    
mRow = mRow + 1
tmp = "a" & mRow & ":" & "h" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "Total Payable"
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 10
oWS.Cells(mRow, C + 1).HorizontalAlignment = -4108

tmp = "i" & mRow
oWS.Range(tmp).Merge
'oWS.Cells(mRow, c + 9) = Format(MSHFlexGrid1.TextMatrix(0, 12) + Val(TxtSTAmount), "#####0.00")
oWS.Cells(mRow, C + 9) = Format(TxtActualFee, "#####0.00") ' Format(MSHFlexGrid2.TextMatrix(0, 14), "#####0.00")
oWS.Cells(mRow, C + 9).Font.Bold = True
oWS.Cells(mRow, C + 9).Font.Size = 10
    
mRow = mRow + 1
tmp = "a" & mRow & ":" & "i" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "Official recommending the expenditure"
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 10
oWS.Cells(mRow, C + 1).HorizontalAlignment = -4108
'oWS.Cells(mRow, c + j).Font.Size = 10

mRow = mRow + 1
tmp = "a" & mRow & ":" & "e" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "Above expenditure has been incurred. Material/Services has been received & checked for quality/quantity"
oWS.Cells(mRow, C + 1).Font.Size = 10
oWS.Cells(mRow, C + 1).RowHeight = 27
oWS.Cells(mRow, C + 1).WrapText = True
'oWS.Cells(mRow, c + j).Font.Size = 10

tmp = "f" & mRow & ":" & "i" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 6) = "Sanctioning Power vide Delegation of Powers"
oWS.Cells(mRow, C + 6).Font.Size = 10
oWS.Cells(mRow, C + 6).RowHeight = 27
oWS.Cells(mRow, C + 6).WrapText = True
'oWS.Cells(mRow, c + j).Font.Size = 10

mRow = mRow + 1
tmp = "a" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "Signature"
oWS.Cells(mRow, C + 1).Font.Size = 10
oWS.Cells(mRow, C + 1).RowHeight = 27
'oWS.Cells(mRow, c + j).Font.Size = 10

tmp = "b" & mRow
oWS.Range(tmp).Merge
'oWS.Cells(mRow, c + j).Font.Size = 10

tmp = "c" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 3) = "Signature"
oWS.Cells(mRow, C + 3).Font.Size = 10
oWS.Cells(mRow, C + 3).RowHeight = 27
'oWS.Cells(mRow, c + j).Font.Size = 10

tmp = "d" & mRow & ":e" & mRow
oWS.Range(tmp).Merge
'oWS.Cells(mRow, c + j).Font.Size = 10

tmp = "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 6) = "Signature"
oWS.Cells(mRow, C + 6).Font.Size = 10
oWS.Cells(mRow, C + 6).RowHeight = 27
'oWS.Cells(mRow, c + j).Font.Size = 10

tmp = "g" & mRow & ":i" & mRow
oWS.Range(tmp).Merge
'oWS.Cells(mRow, c + j).Font.Size = 10


mRow = mRow + 1
tmp = "a" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "Name of Area Manager/State Head"
oWS.Cells(mRow, C + 1).Font.Size = 10
oWS.Cells(mRow, C + 1).RowHeight = 39
oWS.Cells(mRow, C + 1).WrapText = True
'oWS.Cells(mRow, c + j).Font.Size = 10

tmp = "b" & mRow
oWS.Range(tmp).Merge
'oWS.Cells(mRow, c + j).Font.Size = 10

tmp = "c" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 3) = "Name of CMG Manager"
oWS.Cells(mRow, C + 3).Font.Size = 10
oWS.Cells(mRow, C + 3).RowHeight = 39
oWS.Cells(mRow, C + 3).WrapText = True
'oWS.Cells(mRow, c + j).Font.Size = 10

tmp = "d" & mRow & ":e" & mRow
oWS.Range(tmp).Merge
'oWS.Cells(mRow, c + j).Font.Size = 10

tmp = "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 6) = "Name : "
oWS.Cells(mRow, C + 6).Font.Size = 10
oWS.Cells(mRow, C + 6).RowHeight = 39
'oWS.Cells(mRow, c + j).Font.Size = 10

tmp = "g" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 7) = "Mr Suresh Chandra Keshari"
oWS.Cells(mRow, C + 7).Font.Size = 10
oWS.Cells(mRow, C + 7).RowHeight = 39
oWS.Cells(mRow, C + 7).WrapText = True
'oWS.Cells(mRow, c + j).Font.Size = 10

tmp = "h" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 8) = "Designation :"
oWS.Cells(mRow, C + 8).Font.Size = 10
oWS.Cells(mRow, C + 8).RowHeight = 39
'oWS.Cells(mRow, c + j).Font.Size = 10

tmp = "i" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 9) = "Zonal Business Head"
oWS.Cells(mRow, C + 9).Font.Size = 10
oWS.Cells(mRow, C + 9).RowHeight = 39
'oWS.Cells(mRow, c + j).Font.Size = 10

mRow = mRow + 1
tmp = "a" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "Date :"
oWS.Cells(mRow, C + 1).Font.Size = 10
'oWS.Cells(mRow, c + j).Font.Size = 10

tmp = "b" & mRow
oWS.Range(tmp).Merge
'oWS.Cells(mRow, c + j).Font.Size = 10
    
tmp = "c" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 3) = "Date :"
oWS.Cells(mRow, C + 3).Font.Size = 10
'oWS.Cells(mRow, c + j).Font.Size = 10

tmp = "d" & mRow & ":e" & mRow
oWS.Range(tmp).Merge
'oWS.Cells(mRow, c + j).Font.Size = 10

tmp = "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 6) = "Date :"
oWS.Cells(mRow, C + 6).Font.Size = 10
'oWS.Cells(mRow, c + j).Font.Size = 10

tmp = "g" & mRow & ":i" & mRow
oWS.Range(tmp).Merge
'oWS.Cells(mRow, c + j).Font.Size = 10

mRow = mRow + 1
tmp = "a" & mRow & ":i" & mRow
oWS.Range(tmp).Merge
'oWS.Cells(mRow, c + j).Font.Size = 10
    
mRow = mRow + 1
tmp = "a" & mRow & ":" & "i" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "Remarks"
oWS.Cells(mRow, C + 1).Font.Size = 10
oWS.Cells(mRow, C + 1).Font.Bold = True
'oWS.Cells(mRow, c + j).Font.Size = 10

mRow = mRow + 1
tmp = "a" & mRow & ":i" & mRow
oWS.Range(tmp).Merge
'oWS.Cells(mRow, c + j).Font.Size = 10
oWS.Cells(mRow, C + 1).RowHeight = 60

mRow = mRow + 1
tmp = "a" & mRow & ":" & "e" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "Above bill/invoice/DN and calculations have been scrutinized. The payment may be made after deductions as above."
oWS.Cells(mRow, C + 1).Font.Size = 10
oWS.Cells(mRow, C + 1).WrapText = True
oWS.Cells(mRow, C + 1).RowHeight = 27
'oWS.Cells(mRow, c + j).Font.Size = 10

tmp = "f" & mRow & ":" & "i" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 6) = "Above payment received by Cash/Cheque"
oWS.Cells(mRow, C + 6).Font.Size = 10
oWS.Cells(mRow, C + 6).RowHeight = 27
'oWS.Cells(mRow, c + j).Font.Size = 10

mRow = mRow + 1
tmp = "a" & mRow & ":i" & mRow
oWS.Range(tmp).Merge
'oWS.Cells(mRow, c + j).Font.Size = 10
    
mRow = mRow + 1
tmp = "a" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "Accts."
oWS.Cells(mRow, C + 1).Font.Size = 10
'oWS.Cells(mRow, c + j).Font.Size = 10

tmp = "b" & mRow
oWS.Range(tmp).Merge
'oWS.Cells(mRow, c + j).Font.Size = 10

tmp = "c" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 3) = "Officer"
oWS.Cells(mRow, C + 3).Font.Size = 10
'oWS.Cells(mRow, c + j).Font.Size = 10

tmp = "d" & mRow
oWS.Range(tmp).Merge
'oWS.Cells(mRow, c + j).Font.Size = 10


tmp = "e" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 5) = "Approval"
oWS.Cells(mRow, C + 5).Font.Size = 10
'oWS.Cells(mRow, c + j).Font.Size = 10

tmp = "f" & mRow & ":" & "i" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 6) = "Receiver's Signature"
oWS.Cells(mRow, C + 6).Font.Size = 10
'oWS.Cells(mRow, c + j).Font.Size = 10
For j = 1 To mRow
    tmp = "a" & j & ":i" & j
    oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
Next j
'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
oWS.PageSetup.Orientation = xlPortrait

Set oWS = oWB.Worksheets("Sheet2")
oWS.Name = "Annexure"
oWS.ClearArrows
mRow = 1

'oWS.Range("a1:i1").Merge

oWS.Range("a:a").ColumnWidth = 6.29 '.157 '10.86
oWS.Range("b:b").ColumnWidth = 7.43 '17.71
oWS.Range("c:c").ColumnWidth = 16.57
oWS.Range("d:d").ColumnWidth = 6.57 '12.86
oWS.Range("e:e").ColumnWidth = 8.43 '33 '32.14
oWS.Range("f:f").ColumnWidth = 11.71
oWS.Range("g:g").ColumnWidth = 10
oWS.Range("h:h").ColumnWidth = 6.86
oWS.Range("i:i").ColumnWidth = 6.14
  

'oWS.Range("A1:E1").Borders.Color = RGB(0, 0, 0)
tmp = "a" & mRow & ":" & "i" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "NATIONAL COLLATERAL MANAGEMENT SERVICES LIMITED"
oWS.Cells(mRow, C + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 10
'oWS.Cells(mRow, c + j).Font.Size = 10

'oWS.Cells(1, 1).Font.ColorIndex = 4
mRow = mRow + 1
tmp = "a" & mRow & ":" & "i" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "Franchise BILL VOUCHER"
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 10
oWS.Cells(mRow, C + 1).HorizontalAlignment = -4108
'oWS.Cells(mRow, c + j).Font.Size = 10

mRow = mRow + 1
tmp = "a" & mRow & ":" & "b" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "Business Segment :"
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 10
'oWS.Cells(mRow, c + j).Font.Size = 10

tmp = "c" & mRow & ":" & "d" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 3) = "CMP"
oWS.Cells(mRow, C + 3).Font.Size = 10
oWS.Cells(mRow, C + 3).HorizontalAlignment = -4108
'oWS.Cells(mRow, c + j).Font.Size = 10

tmp = "e" & mRow & ":" & "g" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 5) = "Store-Man ID"
oWS.Cells(mRow, C + 5).Font.Bold = True
oWS.Cells(mRow, C + 5).Font.Size = 10
'oWS.Cells(mRow, c + j).Font.Size = 10

tmp = "h" & mRow & ":" & "i" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 8) = "'" & TxtClaimID
oWS.Cells(mRow, C + 8).Font.Bold = True
oWS.Cells(mRow, C + 8).Font.Size = 10
'oWS.Cells(mRow, c + j).Font.Size = 10
   
mRow = mRow + 1
tmp = "a" & mRow & ":" & "b" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "Franchise Vendor Name :"
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 10
'oWS.Cells(mRow, c + j).Font.Size = 10

tmp = "c" & mRow & ":" & "i" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 3) = TxtVendorID
oWS.Cells(mRow, C + 3).Font.Bold = True
oWS.Cells(mRow, C + 3).Font.Size = 10
'oWS.Cells(mRow, c + j).Font.Size = 10
   
mRow = mRow + 1
tmp = "a" & mRow & ":" & "b" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "Month/Year"
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 10
'oWS.Cells(mRow, c + j).Font.Size = 10

tmp = "c" & mRow & ":" & "i" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 3) = "'" & Format(TxtExpenseFrom, "MMM-yyyy")
oWS.Cells(mRow, C + 3).Font.Bold = True
oWS.Cells(mRow, C + 3).Font.Size = 10
'oWS.Cells(mRow, c + j).Font.Size = 10

mRow = mRow + 1
tmp = "a" & mRow & ":" & "i" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "Annexure of Earlier Payment Paid"
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 10
'oWS.Cells(mRow, c + j).Font.Size = 10


mRow = mRow + 1
tmp = "a" & mRow & ":" & "i" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "Details of Last payments"
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 10
'oWS.Cells(mRow, c + j).Font.Size = 10

mRow = mRow + 1
tmp = "a" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "Month"
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 10
oWS.Cells(mRow, C + 1).RowHeight = 15#
'oWS.Cells(mRow, c + j).Font.Size = 10

tmp = "b" & mRow & ":" & "c" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 2) = "Amount Paid"
oWS.Cells(mRow, C + 2).Font.Bold = True
oWS.Cells(mRow, C + 2).Font.Size = 10
oWS.Cells(mRow, C + 2).HorizontalAlignment = -4108
'oWS.Cells(mRow, c + j).Font.Size = 10

tmp = "d" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 4) = "Month"
oWS.Cells(mRow, C + 4).Font.Bold = True
oWS.Cells(mRow, C + 4).Font.Size = 10
oWS.Cells(mRow, C + 4).RowHeight = 15#
'oWS.Cells(mRow, c + j).Font.Size = 10

tmp = "e" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 5).RowHeight = 15#
'oWS.Cells(mRow, c + j).Font.Size = 10

tmp = "f" & mRow & ":" & "g" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 6) = "Amount Paid"
oWS.Cells(mRow, C + 6).Font.Bold = True
oWS.Cells(mRow, C + 6).Font.Size = 10
oWS.Cells(mRow, C + 6).HorizontalAlignment = -4108
'oWS.Cells(mRow, c + j).Font.Size = 10

tmp = "h" & mRow
oWS.Range(tmp).Merge
'oWS.Cells(mRow, c + j).Font.Size = 10

tmp = "i" & mRow
oWS.Range(tmp).Merge
'oWS.Cells(mRow, c + j).Font.Size = 10
    
mRow = mRow + 1
tmp = "b" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 2) = "Amount"
oWS.Cells(mRow, C + 2).Font.Bold = True
oWS.Cells(mRow, C + 2).Font.Size = 10
'oWS.Cells(mRow, c + j).Font.Size = 10

tmp = "c" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 3) = "Cheque No"
oWS.Cells(mRow, C + 3).Font.Bold = True
oWS.Cells(mRow, C + 3).Font.Size = 10
'oWS.Cells(mRow, c + j).Font.Size = 10

tmp = "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 6) = "Amount"
oWS.Cells(mRow, C + 6).Font.Bold = True
oWS.Cells(mRow, C + 6).Font.Size = 10
'oWS.Cells(mRow, c + j).Font.Size = 10

tmp = "g" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 7) = "Cheque No"
oWS.Cells(mRow, C + 7).Font.Bold = True
oWS.Cells(mRow, C + 7).Font.Size = 10
'oWS.Cells(mRow, c + j).Font.Size = 10


tmp = "h" & mRow & ":i" & mRow
oWS.Range(tmp).Merge
'oWS.Cells(mRow, c + j).Font.Size = 10

mRow = mRow + 1
tmp = "a" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "April - " & VBA.Year(Date)
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 10
'oWS.Cells(mRow, c + j).Font.Size = 10

tmp = "b" & mRow
oWS.Range(tmp).Merge
'oWS.Cells(mRow, c + j).Font.Size = 10

tmp = "c" & mRow
oWS.Range(tmp).Merge
'oWS.Cells(mRow, c + j).Font.Size = 10

tmp = "d" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 4) = "Oct - " & VBA.Year(Date)
oWS.Cells(mRow, C + 4).Font.Bold = True
oWS.Cells(mRow, C + 4).Font.Size = 10
'oWS.Cells(mRow, c + j).Font.Size = 10
    
tmp = "e" & mRow
oWS.Range(tmp).Merge
'oWS.Cells(mRow, c + j).Font.Size = 10
    
tmp = "f" & mRow
oWS.Range(tmp).Merge
'oWS.Cells(mRow, c + j).Font.Size = 10
    
tmp = "g" & mRow
oWS.Range(tmp).Merge
'oWS.Cells(mRow, c + j).Font.Size = 10
    
tmp = "h" & mRow
oWS.Range(tmp).Merge
'oWS.Cells(mRow, c + j).Font.Size = 10
    
tmp = "i" & mRow
oWS.Range(tmp).Merge
'oWS.Cells(mRow, c + j).Font.Size = 10
    
    
mRow = mRow + 1
tmp = "a" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "May - " & VBA.Year(Date)
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 10
'oWS.Cells(mRow, c + j).Font.Size = 10

tmp = "b" & mRow
oWS.Range(tmp).Merge
'oWS.Cells(mRow, c + j).Font.Size = 10

tmp = "c" & mRow
oWS.Range(tmp).Merge
'oWS.Cells(mRow, c + j).Font.Size = 10

tmp = "d" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 4) = "Nov - " & VBA.Year(Date)
oWS.Cells(mRow, C + 4).Font.Bold = True
oWS.Cells(mRow, C + 4).Font.Size = 10
'oWS.Cells(mRow, c + j).Font.Size = 10
    
tmp = "e" & mRow
oWS.Range(tmp).Merge
'oWS.Cells(mRow, c + j).Font.Size = 10
    
tmp = "f" & mRow
oWS.Range(tmp).Merge
'oWS.Cells(mRow, c + j).Font.Size = 10
    
tmp = "g" & mRow
oWS.Range(tmp).Merge
'oWS.Cells(mRow, c + j).Font.Size = 10
    
tmp = "h" & mRow
oWS.Range(tmp).Merge
'oWS.Cells(mRow, c + j).Font.Size = 10
    
tmp = "i" & mRow
oWS.Range(tmp).Merge
'oWS.Cells(mRow, c + j).Font.Size = 10
    

mRow = mRow + 1
tmp = "a" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "June - " & VBA.Year(Date)
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 10
'oWS.Cells(mRow, c + j).Font.Size = 10

tmp = "b" & mRow
oWS.Range(tmp).Merge
'oWS.Cells(mRow, c + j).Font.Size = 10

tmp = "c" & mRow
oWS.Range(tmp).Merge
'oWS.Cells(mRow, c + j).Font.Size = 10

tmp = "d" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 4) = "Dec - " & VBA.Year(Date)
oWS.Cells(mRow, C + 4).Font.Bold = True
oWS.Cells(mRow, C + 4).Font.Size = 10
'oWS.Cells(mRow, c + j).Font.Size = 10
    
tmp = "e" & mRow
oWS.Range(tmp).Merge
'oWS.Cells(mRow, c + j).Font.Size = 10
    
tmp = "f" & mRow
oWS.Range(tmp).Merge
'oWS.Cells(mRow, c + j).Font.Size = 10
    
tmp = "g" & mRow
oWS.Range(tmp).Merge
'oWS.Cells(mRow, c + j).Font.Size = 10
    
tmp = "h" & mRow
oWS.Range(tmp).Merge
'oWS.Cells(mRow, c + j).Font.Size = 10
    
tmp = "i" & mRow
oWS.Range(tmp).Merge
'oWS.Cells(mRow, c + j).Font.Size = 10
    

mRow = mRow + 1
tmp = "a" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "July - " & VBA.Year(Date)
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 10
'oWS.Cells(mRow, c + j).Font.Size = 10

tmp = "b" & mRow
oWS.Range(tmp).Merge
'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)

tmp = "c" & mRow
oWS.Range(tmp).Merge
'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)

tmp = "d" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 4) = "Jan - " & VBA.Year(Date) + 1
oWS.Cells(mRow, C + 4).Font.Bold = True
oWS.Cells(mRow, C + 4).Font.Size = 10
'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
    
tmp = "e" & mRow
oWS.Range(tmp).Merge
'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
    
tmp = "f" & mRow
oWS.Range(tmp).Merge
'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
    
tmp = "g" & mRow
oWS.Range(tmp).Merge
'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
    
tmp = "h" & mRow
oWS.Range(tmp).Merge
'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
    
tmp = "i" & mRow
oWS.Range(tmp).Merge
'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
    

mRow = mRow + 1
tmp = "a" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "Aug - " & VBA.Year(Date)
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 10
'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)

tmp = "b" & mRow
oWS.Range(tmp).Merge
'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)

tmp = "c" & mRow
oWS.Range(tmp).Merge
'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)

tmp = "d" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 4) = "Feb - " & VBA.Year(Date) + 1
oWS.Cells(mRow, C + 4).Font.Bold = True
oWS.Cells(mRow, C + 4).Font.Size = 10
'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
    
tmp = "e" & mRow
oWS.Range(tmp).Merge
'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
    
tmp = "f" & mRow
oWS.Range(tmp).Merge
'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
    
tmp = "g" & mRow
oWS.Range(tmp).Merge
'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
    
tmp = "h" & mRow
oWS.Range(tmp).Merge
'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
    
tmp = "i" & mRow
oWS.Range(tmp).Merge
'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
    

mRow = mRow + 1
tmp = "a" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "Sept - " & VBA.Year(Date)
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 10
'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)

tmp = "b" & mRow
oWS.Range(tmp).Merge
'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)

tmp = "c" & mRow
oWS.Range(tmp).Merge
'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)

tmp = "d" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 4) = "Mar - " & VBA.Year(Date) + 1
oWS.Cells(mRow, C + 4).Font.Bold = True
oWS.Cells(mRow, C + 4).Font.Size = 10
'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
    
tmp = "e" & mRow
oWS.Range(tmp).Merge
'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
    
tmp = "f" & mRow
oWS.Range(tmp).Merge
'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
    
tmp = "g" & mRow
oWS.Range(tmp).Merge
'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
    
tmp = "h" & mRow
oWS.Range(tmp).Merge
'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
    
tmp = "i" & mRow
oWS.Range(tmp).Merge
'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
    

mRow = mRow + 2
tmp = "a" & mRow & ":i" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "For Office USE ONLY -- BY CMG"
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 10
'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)

mRow = mRow + 1
tmp = "a" & mRow & ":i" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "Particulars"
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 10
oWS.Cells(mRow, C + 1).HorizontalAlignment = -4108
'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)

mRow = mRow + 1
tmp = "a" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "S.No"
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 10
'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)

tmp = "b" & mRow & ":c" & mRow
oWS.Range(tmp).Merge
'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)

tmp = "d" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 4) = "Yes/No"
oWS.Cells(mRow, C + 4).Font.Bold = True
oWS.Cells(mRow, C + 4).Font.Size = 10
'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)

tmp = "e" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 5) = "S.No"
oWS.Cells(mRow, C + 5).Font.Bold = True
oWS.Cells(mRow, C + 5).Font.Size = 10
'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)

tmp = "f" & mRow & ":h" & mRow
oWS.Range(tmp).Merge
'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)

tmp = "i" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 9) = "Yes/No"
oWS.Cells(mRow, C + 9).Font.Bold = True
oWS.Cells(mRow, C + 9).Font.Size = 10
'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)

mRow = mRow + 1
tmp = "a" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "1"
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 10
'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)

tmp = "b" & mRow & ":c" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 2) = "AM Approval"
oWS.Cells(mRow, C + 2).Font.Bold = True
oWS.Cells(mRow, C + 2).Font.Size = 10
'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)

tmp = "d" & mRow
oWS.Range(tmp).Merge
'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)

tmp = "e" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 5) = "5"
oWS.Cells(mRow, C + 5).Font.Bold = True
oWS.Cells(mRow, C + 5).Font.Size = 10
'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)

tmp = "f" & mRow & ":h" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 6) = "Report Attached"
oWS.Cells(mRow, C + 6).Font.Bold = True
oWS.Cells(mRow, C + 6).Font.Size = 10
'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)

tmp = "i" & mRow
oWS.Range(tmp).Merge
'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)

mRow = mRow + 1
tmp = "a" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "2"
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 10
'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)

tmp = "b" & mRow & ":c" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 2) = "RM Approval"
oWS.Cells(mRow, C + 2).Font.Bold = True
oWS.Cells(mRow, C + 2).Font.Size = 10
'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)

tmp = "d" & mRow
oWS.Range(tmp).Merge
'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)

tmp = "e" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 5) = "6"
oWS.Cells(mRow, C + 5).Font.Bold = True
oWS.Cells(mRow, C + 5).Font.Size = 10
'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)

tmp = "f" & mRow & ":h" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 6) = "Fumigation Confirmation"
oWS.Cells(mRow, C + 6).Font.Bold = True
oWS.Cells(mRow, C + 6).Font.Size = 10
'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)

tmp = "i" & mRow
oWS.Range(tmp).Merge
'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)

mRow = mRow + 1
tmp = "a" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "3"
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 10
'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)

tmp = "b" & mRow & ":c" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 2) = "Stock Verified as per Store-Man"
oWS.Cells(mRow, C + 2).Font.Bold = True
oWS.Cells(mRow, C + 2).Font.Size = 10
'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)

tmp = "d" & mRow
oWS.Range(tmp).Merge
'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)

tmp = "e" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 5) = "7"
oWS.Cells(mRow, C + 5).Font.Bold = True
oWS.Cells(mRow, C + 5).Font.Size = 10
'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)

tmp = "f" & mRow & ":h" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 6) = "Rate Verified"
oWS.Cells(mRow, C + 6).Font.Bold = True
oWS.Cells(mRow, C + 6).Font.Size = 10
'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)

tmp = "i" & mRow
oWS.Range(tmp).Merge
'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)

mRow = mRow + 1
tmp = "a" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "4"
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 10
'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)

tmp = "b" & mRow & ":c" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 2) = "Stock Statement Attached"

oWS.Cells(mRow, C + 2).Font.Bold = True
oWS.Cells(mRow, C + 2).Font.Size = 10
'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)

tmp = "d" & mRow
oWS.Range(tmp).Merge
'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)

tmp = "e" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 5) = "8"
oWS.Cells(mRow, C + 5).Font.Bold = True
oWS.Cells(mRow, C + 5).Font.Size = 10
'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)

tmp = "f" & mRow & ":h" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 6) = "Calculation Checked"
oWS.Cells(mRow, C + 6).Font.Bold = True
oWS.Cells(mRow, C + 6).Font.Size = 10
'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)

tmp = "i" & mRow
oWS.Range(tmp).Merge

For j = 1 To mRow
    tmp = "a" & j & ":i" & j
    oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
Next j
'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
oWS.PageSetup.Orientation = xlPortrait
If Gen_ExcelPassword <> "" Then
oWS.Protect Gen_ExcelPassword, , , , , , True, True
End If

'Next I

oWB.Save
'oWB.Close
'oXL.Quit
MsgBox ("Excel File Created !!! ")
oXL.Visible = True

End Sub


Private Sub EditCmd_Click()
Dim ans As Variant

'If LCase(TxtStatus) = "approved" Or LCase(TxtStatus) = "rejected" Or LCase(TxtStatus) = "freez" Then
'   MsgBox "This Claim is " & TxtStatus & ". You can't edit this claim any more!!!"
'   Exit Sub
'End If

''-- 3rd Jan 2017 Ullhas Patil
If LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "fag" And LCase(Gen_UserRole) <> "fagadmin" Then
   If LCase(TxtStatus) = "settled" Or LCase(TxtStatus) = "approved" Or LCase(TxtStatus) = "rejected" Or LCase(TxtStatus) = "frozen" Then
      MsgBox "This Claim is " & TxtStatus & ". You can't edit this claim any more!!!"
      Exit Sub
   End If
Else
   If LCase(TxtStatus) = "rejected" Or LCase(TxtStatus) = "settled" Or LCase(TxtStatus) = "freez" Then
      MsgBox "This Claim is " & TxtStatus & ". You can't edit this claim any more!!!"
      Exit Sub
   End If
End If

mAllowEdit = True
''-- 3rd Jan 2017 Ullhas Patil
If mStatus = "A" Then
    mAllowEdit = False
End If

If Format(Gen_BlockingDate, "yyyy-mm") >= Format(TxtExpenseFrom.Value, "yyyy-mm") Then
   ans = MsgBox("Expense month is previous month from " & Format(Gen_BlockingDate, "MMM-yyyy") & ", Still you want to edit!!!", vbYesNo)
   If ans = vbNo Then
      Exit Sub
   Else
      mAllowEdit = False
   End If
End If

Edit_Flg = "E"
Call GButtonLock(Me, False)
'TxtClaimDate.Enabled = True

TxtForwardTo.Visible = False
CmbForwardTo.Visible = True
CmbForwardTo.Text = TxtForwardTo
CmdForwardTo.Visible = True

''--- 04 Jan 2017 Ullhas Patil
If mStatus = "A" Or mStatus = "S" Then
    TxtFranchInvNo.Locked = True
Else
    TxtFranchInvNo.Locked = False
End If

Label13.Visible = True
TxtFranchInvDate.Visible = True

''--- 04 Jan 2017 Ullhas Patil
If mStatus = "A" Then
    TxtServiceTaxID.Visible = True
    CmbServiceTaxID.Visible = False
    CmbServiceTaxID.Text = IIf(TxtServiceTaxID = "0", "", TxtServiceTaxID)
Else
    TxtServiceTaxID.Visible = False
    CmbServiceTaxID.Visible = True
    CmbServiceTaxID.Text = IIf(TxtServiceTaxID = "0", "", TxtServiceTaxID)
End If

CmdPrint.Enabled = False
'If LCase(Gen_UserRole) = "fag" Then ''Or LCase(Gen_UserRole) = "power"
TxtStatus.Visible = False
CmbStatus.Visible = True
CmbStatus.Text = TxtStatus
TxtApprovalRemarks.Locked = False
'End If

TxtApprovalDate.Value = Date
'TxtExpenseFrom.Enabled = True
'TxtExpenseTo.Enabled = True
TxtRemark.Locked = False
TxtActualFee.Locked = True

''--- Ullhas Patil 25 Oct 2017
CmbBillingState.Visible = True
TxtBillingState.Visible = False
CmbBillingState = TxtBillingState
''-----------------------------

If mAllowEdit = True Then
   TxtActualFee.Locked = False
Else
   TxtActualFee.Locked = True
End If

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

Call Grid_Head
Call Grid_Head1

SFrom.Value = CDate(From_GenDate) - 900
STo.Value = Date 'To_GenDate
CmbSLessor.Text = ""
CmbSStatus.Text = ""
CmbSForwardTo.Text = ""
CmbSLocation.Text = ""
ChkBatchNo.Value = 0
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

'mEID = Replace(Gen_WcEmployeeID, "EmployeeID", "CMPMID")
'
'
'tmp = "Select Distinct(EMP.EmployeeID) " ', EMP.EmployeeNO, EMP.EmployeeName"
'tmp = tmp & " From Mst_Employee EMP"
'tmp = tmp & " Left Outer Join Mst_Employee EMP1 On EMP.DirectReportingTo = EMP1.EmployeeID"
'tmp = tmp & " Left Outer Join Mst_Employee EMP2 On EMP1.DirectReportingTo = EMP2.EmployeeID"
'tmp = tmp & " Left Outer Join Mst_Employee EMP3 On EMP2.DirectReportingTo = EMP3.EmployeeID"
'tmp = tmp & " Left Outer Join Mst_Employee EMP4 On EMP3.DirectReportingTo = EMP4.EmployeeID"
'tmp = tmp & " Left Outer Join Mst_Employee EMP5 On EMP4.DirectReportingTo = EMP5.EmployeeID"
'tmp = tmp & " Left Outer Join Mst_Employee EMP6 On EMP5.DirectReportingTo = EMP6.EmployeeID"
'tmp = tmp & " Left Outer Join Mst_Employee EMP7 On EMP6.DirectReportingTo = EMP7.EmployeeID"
'tmp = tmp & " Left Outer Join Mst_Employee EMP8 On EMP5.DirectReportingTo = EMP8.EmployeeID"
'tmp = tmp & " Left Outer Join Mst_Employee EMP9 On EMP8.DirectReportingTo = EMP9.EmployeeID"
'
'tmp = tmp & " Where (EMP.EmployeeID = " & Gen_EmployeeID & " Or EMP1.EmployeeID = " & Gen_EmployeeID & " Or EMP2.EmployeeID =" & Gen_EmployeeID & " Or"
'tmp = tmp & " EMP3.EmployeeID = " & Gen_EmployeeID & " Or EMP4.EmployeeID = " & Gen_EmployeeID & " or EMP5.EmployeeID=" & Gen_EmployeeID & " Or"
'tmp = tmp & " EMP6.EmployeeID = " & Gen_EmployeeID & " Or EMP7.EmployeeID = " & Gen_EmployeeID & " Or EMP8.EmployeeID = " & Gen_EmployeeID & " Or"
'tmp = tmp & " EMP9.EmployeeID = " & Gen_EmployeeID & " )"
'tmp = tmp & " and Emp.Flag='N' "
'
'Call Tmp4Table
'
'If TmpRs4.RecordCount > 0 Then
'      mDeactiveEID = "CMPMID in ("
'      For i = 1 To TmpRs4.RecordCount
'          If i = 1 Then
'             mDeactiveEID = mDeactiveEID & TmpRs4.Fields(0) & " "
'          Else
'             mDeactiveEID = mDeactiveEID & "," & TmpRs4.Fields(0) & " "
'          End If
'          TmpRs4.MoveNext
'      Next
'      mDeactiveEID = mDeactiveEID & " )"
'End If

If LCase(Gen_UserRole) = "power" Or LCase(Gen_UserRole) = "fag" Or LCase(Gen_UserRole) = "fagadmin" Or Gen_DesignationID = 1 Or Gen_DesignationID = 7 Then
   Call TableTxn_FranchiseeRent
Else
'   If mDeactiveEID <> "" Then
'      Call TableTxn_FranchiseeRent(" Where (" & mEID & " Or " & mDeactiveEID & ") And " & Gen_WcLocation & " ")
'   Else
    Call TableTxn_FranchiseeRent(" Where " & Gen_WcLocation & " ")
'   End If
End If


'If LCase(Gen_UserRole) = "power" Or LCase(Gen_UserRole) = "fag" Or Gen_DesignationID = 1 Then
'   Call TableTxn_FranchiseeRent
'Else
'   Call TableTxn_FranchiseeRent(" Where " & mEID & "")
'End If

If RsTxn_FranchiseeRent.RecordCount = 0 Then
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
RsTxn_FranchiseeRent.MoveFirst
RsTxn_FranchiseeRent.Find "G_UID = '" & MSHFlexGrid1.TextMatrix(I, 14) & "'"
If Not RsTxn_FranchiseeRent.EOF Then
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

''If LCase(Gen_UserRole) = "power" Or LCase(Gen_UserRole) = "superadmin" Then
''   mSelRow = MSHFlexGrid2.Row
''   mSelCol = MSHFlexGrid2.Col
''   If mSelCol <> 7 Then Exit Sub
''   TxtArrears.Visible = True
''   TxtArrears.Text = MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Row, 7)
''   With TxtArrears
''        .Left = (MSHFlexGrid2.Left + MSHFlexGrid2.CellLeft) - 30
''        .Top = (MSHFlexGrid2.Top + MSHFlexGrid2.CellTop) - 30
''        .Width = MSHFlexGrid2.CellWidth + 35
''        .Height = MSHFlexGrid2.CellHeight '- 200
''        .ZOrder
''        .SetFocus
''   End With
''End If


End Sub
Private Sub SaveCmd_Click()

If TxtExpenseFrom.Value > TxtExpenseTo.Value Then
   MsgBox "'Fee From ' must be less than 'Fee To'!!!"
   TxtExpenseFrom.SetFocus
   Exit Sub
End If

If CmbForwardTo.Text = "" Then
   MsgBox "Blank Forward to not allowed!!!"
   CmbForwardTo.SetFocus
   Exit Sub
End If

If TxtActualFee = "" Then
   MsgBox "Blank Actual fee not allowed!!"
   TxtActualFee.SetFocus
   Exit Sub
End If

''-- Ullhas Patil 25 Oct 2017
If CmbBillingState.Text = "" Then
   MsgBox "Blank Billing State not allowed!!!"
   CmbBillingState.SetFocus
   Exit Sub
End If
''---------------------------

If Gen_DesignationID = 6 Then
   tmp = "Select UG.GroupName from Mst_Users MU"
   tmp = tmp & " Inner Join  Mst_UserRole UG On UG.UserRoleID=MU.UserRoleID"
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

If TxtFranchInvNo.Text <> "" Then
   If CheckDuplicateBill() = True Then
      MsgBox " Duplicate Bill No & Bill Date not allowed for selected Vendor!!"
      TxtFranchInvNo.SetFocus
      Exit Sub
   End If
End If

If Edit_Flg = "E" Then
   'If LCase(Gen_UserRole) = "fag" Then ''' Or LCase(Gen_UserRole) = "power"
   If CmbStatus.Text = "" Then
      MsgBox "Blank Status not allowed!!!"
      CmbStatus.SetFocus
      Exit Sub
   End If
   If LCase(Gen_UserRole) = "fag" Or LCase(Gen_UserRole) = "fagadmin" Or LCase(Gen_UserRole) = "power" Then
      If LCase(CmbStatus.Text) = "approved" Or LCase(CmbStatus.Text) = "rejected" Or LCase(CmbStatus.Text) = "frozen" Or LCase(CmbStatus.Text) = "settled" Then
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
         If Format(Gen_BlockingDate, "yyyy-mm") >= Format(TxtApprovalDate.Value, "yyyy-mm") Then
            MsgBox "NAV Booked date should be greater than blocked date!!!"
            TxtApprovalDate.SetFocus
         End If
      End If
   End If
   ''-- 03 Jan 2017 Ullhas Patil
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
          
          If Val(TxtFAGDisbAmt) < Val(TxtActualFee) Then
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

If Edit_Flg = "A" Then
   RsTxn_FranchiseeRent.AddNew
'   RsTxn_FranchiseeRent!ClaimID = GetMaxClaimID
   RsTxn_FranchiseeRent!G_UID = GetGUID
   TxtClaimID = RsTxn_FranchiseeRent!ClaimID
   TxtG_UID = RsTxn_FranchiseeRent!G_UID
Else
   RsTxn_FranchiseeRent.MoveFirst
   RsTxn_FranchiseeRent.Find "G_UID= '" & TxtG_UID & "'"
End If

RsTxn_FranchiseeRent!ClaimDate = TxtClaimDate.Value
RsTxn_FranchiseeRent!Amount = Val(TxtClaimAmount)
RsTxn_FranchiseeRent!Flag = mStatus
RsTxn_FranchiseeRent!ServiceTaxID = mServiceTaxID
RsTxn_FranchiseeRent!Remarks = TxtRemark
If TxtFranchInvNo.Text <> "" Then
   RsTxn_FranchiseeRent!FranchInvoiceNo = TxtFranchInvNo.Text
   RsTxn_FranchiseeRent!FranchInvoiceDate = TxtFranchInvDate.Value
Else
   RsTxn_FranchiseeRent!FranchInvoiceNo = Null
   RsTxn_FranchiseeRent!FranchInvoiceDate = Null
End If
If TxtApprovalRemarks.Locked = False Then
   RsTxn_FranchiseeRent!ApprovalRemarks = Trim(TxtApprovalRemarks)
End If

If TxtApprovalDate.Visible = True Then
   RsTxn_FranchiseeRent!ApprovedDate = TxtApprovalDate.Value
End If

RsTxn_FranchiseeRent!ActualFee = Val(TxtActualFee)
RsTxn_FranchiseeRent!ForwardToID = mForwardToID
RsTxn_FranchiseeRent!ExpenseFrom = TxtExpenseFrom.Value
RsTxn_FranchiseeRent!ExpenseTo = TxtExpenseTo.Value
RsTxn_FranchiseeRent!BatchID = mBatchID

If LCase(CmbStatus.Text) = "settled" Then
   RsTxn_FranchiseeRent!DisbuDate = TxtFagDisbDate.Value
   RsTxn_FranchiseeRent!DisbuRemark = TxtFAGRemark.Text
   RsTxn_FranchiseeRent!DisbuAmount = TxtFAGDisbAmt.Text
   RsTxn_FranchiseeRent!ModeOfPay = CmbModeOfPay.Text
   RsTxn_FranchiseeRent!FAGBankID = mFAGBankID
   RsTxn_FranchiseeRent!DeductionRemark = Trim(TxtDeductionRemark)
Else
   RsTxn_FranchiseeRent!DisbuDate = Null
   RsTxn_FranchiseeRent!DisbuRemark = Null
   RsTxn_FranchiseeRent!DisbuAmount = Null
   RsTxn_FranchiseeRent!ModeOfPay = Null
   RsTxn_FranchiseeRent!FAGBankID = Null
   RsTxn_FranchiseeRent!DeductionRemark = Null
End If

RsTxn_FranchiseeRent!BillingStateIndiaID = mBillingStateID      ''-- Ullhas Patil 25 Oct 2017

If IsNull(RsTxn_FranchiseeRent!CreatedBy) = True Or RsTxn_FranchiseeRent!CreatedBy = "" Then
   RsTxn_FranchiseeRent!CreatedBy = Gen_UserLoginID
   RsTxn_FranchiseeRent!CreatedDate = Now
Else
   RsTxn_FranchiseeRent!UpdatedBy = Gen_UserLoginID
   RsTxn_FranchiseeRent!UpdatedDate = Now
End If

If LCase(Gen_UserRole) = "fag" Or LCase(Gen_UserRole) = "power" Or LCase(Gen_UserRole) = "fagadmin" Then
   If LCase(mStatus) = "a" Then
      RsTxn_FranchiseeRent!UpdatedByFAG = Gen_UserLoginID
      RsTxn_FranchiseeRent!UpdatedDateFAG = Now
   End If
End If

RsTxn_FranchiseeRent.Update

TxtCreated = IIf(IsNull(RsTxn_FranchiseeRent!CreatedBy), "", RsTxn_FranchiseeRent!CreatedBy) & " :- " & IIf(IsNull(RsTxn_FranchiseeRent!CreatedDate), "", RsTxn_FranchiseeRent!CreatedDate)
TxtUpdated = IIf(IsNull(RsTxn_FranchiseeRent!UpdatedBy), "", RsTxn_FranchiseeRent!UpdatedBy) & " :- " & IIf(IsNull(RsTxn_FranchiseeRent!UpdatedDate), "", RsTxn_FranchiseeRent!UpdatedDate)

If LCase(mStatus) = "r" Or LCase(mStatus) = "f" Then
   Call SaveLogFranchClaimDetail
   Call UpdateFranchClaimID
End If

'If LCase(mStatus) <> "f" Then
  Call Gen_Email
'End If
RsTxn_FranchiseeRent.Filter = ""
RsTxn_FranchiseeRent.Requery
If RsTxn_FranchiseeRent.RecordCount > 0 Then
   RsTxn_FranchiseeRent.MoveFirst
   RsTxn_FranchiseeRent.Find "G_UID = '" & TxtG_UID & "'"
   
   Call Indata
Else
   MsgBox "Data saved!!!"
   
   MsgBox "No Pending Claim under your ID!!!!"
   'Call Grid_Head
   'Call GButtonLock_1(Me, True)
   Unload Me
   
End If

End Sub
Private Sub InDataLogHistory()

Call Grid_HeadHistory

tmp = "Select ClaimID ,OldActualFee,NewActualFee, OldForwardToID, NewForwardToID, UpdatedBy, Concurrency"
tmp = tmp & " From Log_FranchiseeFeeHistory Where ClaimID=" & Val(TxtClaimID) & ""

Call TmpTable
MSHFlexGrid4.Rows = MSHFlexGrid4.Rows + TmpRs.RecordCount
For I = 1 To MSHFlexGrid4.Rows - 2
   mTmpEmpName = ""
   MSHFlexGrid4.TextMatrix(I, 0) = IIf(IsNull(TmpRs!ClaimID), 0, TmpRs!ClaimID)
   MSHFlexGrid4.TextMatrix(I, 1) = IIf(IsNull(TmpRs!OldActualFee), 0, TmpRs!OldActualFee)
   MSHFlexGrid4.TextMatrix(I, 2) = IIf(IsNull(TmpRs!NewActualFee), 0, TmpRs!NewActualFee)
   mTmpEmpName = GetEmployeeNonName(IIf(IsNull(TmpRs!OldForwardToID), "", TmpRs!OldForwardToID))
   MSHFlexGrid4.TextMatrix(I, 3) = mTmpEmpName
   mTmpEmpName = GetEmployeeNonName(IIf(IsNull(TmpRs!NewForwardToID), 0, TmpRs!NewForwardToID))
   MSHFlexGrid4.TextMatrix(I, 4) = mTmpEmpName

   MSHFlexGrid4.TextMatrix(I, 5) = IIf(IsNull(TmpRs!UpdatedBy), "", TmpRs!UpdatedBy)
   MSHFlexGrid4.TextMatrix(I, 6) = Format(IIf(IsNull(TmpRs!Concurrency), "", TmpRs!Concurrency), "dd-mmm-yyyy hh:mm:ss")
   TmpRs.MoveNext
Next
End Sub
Private Sub Grid_HeadHistory()
With MSHFlexGrid4
     .Clear
     .Rows = 2
     .Cols = 7
     .AllowUserResizing = flexResizeBoth
     .TextMatrix(0, 0) = "Rent ID"
     .TextMatrix(0, 1) = "Old Actual Fee"
     .TextMatrix(0, 2) = "New Actual Fee"
     .TextMatrix(0, 3) = "Old Forward ID"
     .TextMatrix(0, 4) = "New Forward ID"
     .TextMatrix(0, 5) = "Updated By"
     .TextMatrix(0, 6) = "Updated On"
     .ColWidth(0) = 1000
     .ColWidth(1) = 1400
     .ColWidth(2) = 1400
     .ColWidth(3) = 3000
     .ColWidth(4) = 3000
     .ColWidth(5) = 2000
     .ColWidth(6) = 2300
     
End With
End Sub
Private Sub Gen_Email()
Dim TOADD As Variant
Dim TOCC, mMsg, mGridMsg As Variant
Dim strSplitString() As String
Dim EmpNo_, EmpName_ As String
Dim mPANno, mAccNo, mRTGSNo, mBankName, mBranchName As Variant
Dim mRow, mCol, mLocationName, mModeOfPayment, mData As Variant

mLocationName = TxtLocationID 'GetLocationFromCMPName(MSHFlexGrid2.TextMatrix(2, 2))

''tmp = "Select ML.LessorName,ML.PANNo,ML.AccountNo,ML.RTGSNo,"
''tmp = tmp & " MB.BankName , MBR.BranchName"
''tmp = tmp & " From Mst_Lessor ML"
''tmp = tmp & " Left Join Mst_Bank MB on MB.BankID= ML.BankID"
''tmp = tmp & " Left Join Mst_Branch MBR on MBR.BranchID=ML.BranchID"
''tmp = tmp & " Where ML.VendorID = " & mVendorID & ""
''Call TmpTable


''If TmpRs.RecordCount > 0 Then
''   mPANno = IIf(IsNull(TmpRs!PANNo), "", TmpRs!PANNo)
''   mAccNo = IIf(IsNull(TmpRs!AccountNo), "", TmpRs!AccountNo)
''   mRTGSNo = IIf(IsNull(TmpRs!RTGSNo), "", TmpRs!RTGSNo)
''   mBankName = IIf(IsNull(TmpRs!BankName), "", TmpRs!BankName)
''   mBranchName = IIf(IsNull(TmpRs!BranchName), "", TmpRs!BranchName)
''End If
''mPANno = IIf(mPANno = "", "&nbsp", mPANno)
''mAccNo = IIf(mAccNo = "", "&nbsp", mAccNo)
''mRTGSNo = IIf(mRTGSNo = "", "&nbsp", mRTGSNo)
''mBankName = IIf(mBankName = "", "&nbsp", mBankName)
''mBranchName = IIf(mBranchName = "", "&nbsp", mBranchName)
''mModeOfPayment = IIf(mModeOfPayment = "", "&nbsp", mModeOfPayment)


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

mMsg = "<table border = 1 cellpadding=0  cellSpacing=0><tr>"
mMsg = mMsg & " <td colspan=" & 6 & " >"
mMsg = mMsg & "   <C> <b>Franchisee Rent Claim Details</b></C><br /></td>"
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
mMsg = mMsg & "    <td ><b>Vendor Name</b></td>"
mMsg = mMsg & "    <td >" & TxtVendorID.Text & "</td>"
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


mMsg = mMsg & mGridMsg & "</table> </td> </tr></table>"

TOADD = mForwardToEmailID ' mLessorEmailID
TOCC = Gen_UserEmailID  '& "," & mForwardToEmailID

If LCase(mStatus) = "p" Then
   mMsgSubject = "Franchisee Rent Claim Pending"
ElseIf LCase(mStatus) = "a" Then
   mMsgSubject = "Franchisee Rent Claim Approved"
ElseIf LCase(mStatus) = "r" Then
   mMsgSubject = "Franchisee Rent Claim Rejected"
ElseIf LCase(mStatus) = "f" Then
   mMsgSubject = "Franchisee Rent Claim Freezed"
End If

Call SendMail(TOADD, TOCC, mMsgSubject, mMsg)


End Sub
Private Function SendMail(ToADD_ As Variant, TOCC_ As Variant, MsgSubject_ As Variant, Message1_ As Variant)
Dim htmlStringhead As String
Dim htmlStringtail As String
Set oSMTP = New OSSMTP.SMTPSession

servertxt = "10.0.0.13"

'htmlStringhead = "<html><body><font size= 3 face=Zurich BT color = Blue><table border = 1 cellpadding=1 cellspacing =1>"
'htmlStringtail = "</table></Font></body></html>"

   
'With oSMTP
'  'connection
'  .Server = servertxt
'  'message
'  .MailFrom = TOCC_ '"mayur.d@ncml.com"
'  If IsNull(Trim(ToADD_)) = True Or Trim(ToADD_) = "" Then
'     ToADD_ = "bhagwat.k@ncml.com"
'  End If
'  .SendTo = ToADD_
'  TOCC_ = TOCC_ & "," & mEmployeeEmailID  '& "," & mLessorEmailID
'  If ToADD_ <> TOCC_ Then
'     .CC = TOCC_
'  End If
''  .BCC = "mayur.d@ncml.com"
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

Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   If RsTxn_FranchiseeRent.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   RsTxn_FranchiseeRent.MoveFirst
 
   RsTxn_FranchiseeRent.Find "G_UID = '" & TxtG_UID & "'"
   If RsTxn_FranchiseeRent.EOF Then
      RsTxn_FranchiseeRent.MoveLast
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

RsTxn_FranchiseeRent.MoveFirst

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
tmp = "Select max(ClaimID) from Txn_FranchiseeRentClaim "
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetMaxClaimID = Retval
End Function

Function GetMaxLessorClaimDetails() As Double
Dim Retval As Double
Retval = 0
tmp = "Select max(ClaimDetailID) From Txn_FranchiseeRentClaimDetail " 'where ClaimID = " & TxtClaimID & " And SM_prefix = '" & TxtPreFix & "'"
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
RsTxn_FranchiseeRent.MoveNext
If RsTxn_FranchiseeRent.EOF Then
   RsTxn_FranchiseeRent.MoveLast
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
RsTxn_FranchiseeRent.MovePrevious
If RsTxn_FranchiseeRent.BOF Then
   RsTxn_FranchiseeRent.MoveFirst
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
RsTxn_FranchiseeRent.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsTxn_FranchiseeRent.MoveLast
Call Indata
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False
End Sub
Public Sub Indata()

TxtApprovalDate.Visible = False
Label27.Visible = False
Label13.Visible = False
mStatus = ""
TxtForwardTo.Visible = True
CmbForwardTo.Visible = False
CmdForwardTo.Visible = False

TxtServiceTaxID.Visible = True
CmbServiceTaxID.Visible = False

TxtStatus.Visible = True
CmbStatus.Visible = False

TxtExpenseFrom.Enabled = False
TxtExpenseTo.Enabled = False
TxtClaimDate.Enabled = False
TxtActualFee.Locked = True
TxtFranchInvNo.Locked = True

''--- Ullhas Patil 25 Oct 2017
CmbBillingState.Visible = False: CmbBillingState = ""
TxtBillingState.Visible = True: TxtBillingState = ""
''-----------------------------

TxtClaimID = RsTxn_FranchiseeRent!ClaimID
TxtClaimDate.Value = RsTxn_FranchiseeRent!ClaimDate
TxtExpenseFrom.Value = RsTxn_FranchiseeRent!ExpenseFrom
TxtExpenseTo.Value = RsTxn_FranchiseeRent!ExpenseTo
mVendorID = RsTxn_FranchiseeRent!VendorID
TxtVendorID = GetVendorName(RsTxn_FranchiseeRent!VendorID, "N")
TxtLocationID = GetLocationName(IIf(IsNull(RsTxn_FranchiseeRent!LocationID), 0, RsTxn_FranchiseeRent!LocationID))
'mLessorEmailID = GetLessorName(mVendorID)
TxtClaimAmount = RsTxn_FranchiseeRent!Amount 'a 'Val(MSHFlexGrid2.TextMatrix(0, 4))
TxtActualFee = IIf(IsNull(RsTxn_FranchiseeRent!ActualFee), "", RsTxn_FranchiseeRent!ActualFee)
TxtServiceTaxID = GetServiceTaxDescription(IIf(IsNull(RsTxn_FranchiseeRent!ServiceTaxID), 0, RsTxn_FranchiseeRent!ServiceTaxID))
mServiceTaxID = IIf(IsNull(RsTxn_FranchiseeRent!ServiceTaxID), 0, RsTxn_FranchiseeRent!ServiceTaxID)

TxtFranchInvNo = IIf(IsNull(RsTxn_FranchiseeRent!FranchInvoiceNo), "", RsTxn_FranchiseeRent!FranchInvoiceNo)
If IsNull(RsTxn_FranchiseeRent!FranchInvoiceDate) = False Then
   Label13.Visible = True
   TxtFranchInvDate.Visible = True
   TxtFranchInvDate.Value = RsTxn_FranchiseeRent!FranchInvoiceDate
   TxtFranchInvDate.Enabled = False
Else
   TxtFranchInvDate.Value = Date
   TxtFranchInvDate.Visible = False
   Label13.Visible = False
End If
mStatus = RsTxn_FranchiseeRent!Flag
mBatchID = IIf(IsNull(RsTxn_FranchiseeRent!BatchID), 0, RsTxn_FranchiseeRent!BatchID)
mForwardToID = IIf(IsNull(RsTxn_FranchiseeRent!ForwardToID), "", RsTxn_FranchiseeRent!ForwardToID)

If mForwardToID <> "" Then
   TxtForwardTo = GetEmployeeNonName(RsTxn_FranchiseeRent!ForwardToID)
   mForwardToEmailID = GetEmployeeEmailID(mForwardToID)
Else
   TxtForwardTo = "" 'GetLessorNonName(RsTxn_FranchiseeRent!ForwardToID)
   mForwardToEmailID = Null 'GetLessorEmailID(mForwardToID)
End If
TxtFWDDesignation = GetDesignationFromEmp(mForwardToID)
mSaveStatus = mStatus
If LCase(mStatus) = "a" Then
   TxtStatus = "Approved"
ElseIf LCase(mStatus) = "p" Then
   TxtStatus = "Pending"
ElseIf LCase(mStatus) = "r" Then
   TxtStatus = "Rejected"
ElseIf LCase(mStatus) = "f" Then
   TxtStatus = "Freez"
''-- 3rd Jan 2017 Ullhas Patil
ElseIf LCase(mStatus) = "s" Then
   TxtStatus = "Settled"
Else
   TxtStatus = ""
End If

TxtRemark = IIf(IsNull(RsTxn_FranchiseeRent!Remarks), "", RsTxn_FranchiseeRent!Remarks)
TxtApprovalRemarks = IIf(IsNull(RsTxn_FranchiseeRent!ApprovalRemarks), "", RsTxn_FranchiseeRent!ApprovalRemarks)

If IsNull(RsTxn_FranchiseeRent!ApprovedDate) = False Then
   TxtApprovalDate.Visible = True
   Label27.Visible = True
   TxtApprovalDate.Value = RsTxn_FranchiseeRent!ApprovedDate
   TxtApprovalDate.Enabled = False
End If

TxtG_UID = IIf(IsNull(RsTxn_FranchiseeRent!G_UID), "", RsTxn_FranchiseeRent!G_UID)
CmdPrint.Enabled = True
TxtRemark.Locked = True
TxtApprovalRemarks.Locked = True
TxtApprovalRemarks.Locked = True

''-- 3rd Jan 2017 Ullhas Patil
CmbModeOfPay.Visible = False
TxtModeOfPayment.Visible = True

CmbFAGBank.Visible = False
TxtFAGBank.Visible = True
Call VendorBankDetails(CDbl(mVendorID))
TxtFAGDisbAmt = IIf(IsNull(RsTxn_FranchiseeRent!DisbuAmount), "", RsTxn_FranchiseeRent!DisbuAmount)
TxtFAGRemark = IIf(IsNull(RsTxn_FranchiseeRent!DisbuRemark), "", RsTxn_FranchiseeRent!DisbuRemark)
TxtFagDisbDate.Value = IIf(IsNull(RsTxn_FranchiseeRent!DisbuDate), Date, RsTxn_FranchiseeRent!DisbuDate)
mFAGBankID = IIf(IsNull(RsTxn_FranchiseeRent!FAGBankID), 0, RsTxn_FranchiseeRent!FAGBankID)
TxtFAGBank.Text = GetFAGBankAcNo(Val(mFAGBankID))
TxtModeOfPayment.Text = IIf(IsNull(RsTxn_FranchiseeRent!ModeOfPay), "", RsTxn_FranchiseeRent!ModeOfPay)
TxtDeductionRemark = IIf(IsNull(RsTxn_FranchiseeRent!DeductionRemark), "", RsTxn_FranchiseeRent!DeductionRemark)

TxtDocNo = IIf(IsNull(RsTxn_FranchiseeRent!DocNo), "", RsTxn_FranchiseeRent!DocNo)
TxtEntryNo = IIf(IsNull(RsTxn_FranchiseeRent!EntryNo), "", RsTxn_FranchiseeRent!EntryNo)
TxtExportedDate = IIf(IsNull(RsTxn_FranchiseeRent!FAGPayExportedDate), "", Format(RsTxn_FranchiseeRent!FAGPayExportedDate, "dd-MMM-yyyy"))
''------------------------------

TxtBankID.Text = IIf(IsNull(RsTxn_FranchiseeRent!BankName), "", RsTxn_FranchiseeRent!BankName)
TxtChequeNo.Text = IIf(IsNull(RsTxn_FranchiseeRent!ChequeNo), "", RsTxn_FranchiseeRent!ChequeNo)
TxtPaidAmt.Text = IIf(IsNull(RsTxn_FranchiseeRent!PaidAmount), "", Format(RsTxn_FranchiseeRent!PaidAmount, "#####0.00"))
TxtTDS.Text = IIf(IsNull(RsTxn_FranchiseeRent!TDS), "", Format(RsTxn_FranchiseeRent!TDS, "#####0.00"))
TxtPaidDate.Text = IIf(IsNull(RsTxn_FranchiseeRent!PaidDate), "", Format(RsTxn_FranchiseeRent!PaidDate, "dd-MMM-yyyy"))

Frame3.Enabled = False

''--- Ullhas Patil 25 Oct 2017
mBillingStateID = IIf(IsNull(RsTxn_FranchiseeRent!BillingStateIndiaID), 0, RsTxn_FranchiseeRent!BillingStateIndiaID)
TxtBillingState = GetStateIndiaName(mBillingStateID)
TxtBillingState = IIf(Trim(TxtBillingState) = "0", "", TxtBillingState)
''-----------------------------

TxtCreated = IIf(IsNull(RsTxn_FranchiseeRent!CreatedBy), "", RsTxn_FranchiseeRent!CreatedBy) & " :- " & IIf(IsNull(RsTxn_FranchiseeRent!CreatedDate), "", RsTxn_FranchiseeRent!CreatedDate)
TxtUpdated = IIf(IsNull(RsTxn_FranchiseeRent!UpdatedBy), "", RsTxn_FranchiseeRent!UpdatedBy) & " :- " & IIf(IsNull(RsTxn_FranchiseeRent!UpdatedDate), "", RsTxn_FranchiseeRent!UpdatedDate)
TxtUpdatedByFAG = IIf(IsNull(RsTxn_FranchiseeRent!UpdatedByFAG), "", RsTxn_FranchiseeRent!UpdatedByFAG) & " :- " & IIf(IsNull(RsTxn_FranchiseeRent!UpdatedDateFAG), "", RsTxn_FranchiseeRent!UpdatedDateFAG)

If LCase(mStatus) = "r" Then
   Call InDataLogFranchClaimDetail
Else
   Call IndataGSLDetails
End If
Call InDataLogHistory

'If RsTxn_FranchiseeRent.RecordCount = 1 Then
'   Call GButtonLock_1(Me, True)
'Else
Call GButtonLock(Me, True)
'End If

SSTab1.Tab = 0

End Sub
''Private Sub Grid_Head()
''MSHFlexGrid2.Clear
''MSHFlexGrid2.Rows = 3
''MSHFlexGrid2.FixedRows = 2
''MSHFlexGrid2.Cols = 15
''
'''MSHFlexGrid2.TextMatrix(1, 0) = "Select"
''MSHFlexGrid2.TextMatrix(1, 0) = "Row ID"
''MSHFlexGrid2.TextMatrix(1, 1) = "Agreement ID"
''MSHFlexGrid2.TextMatrix(1, 2) = "CMP Name"
''MSHFlexGrid2.TextMatrix(1, 3) = "Godown No"
''MSHFlexGrid2.TextMatrix(1, 4) = "Godown Address"
''
''MSHFlexGrid2.TextMatrix(1, 5) = "Security Deposit"
''MSHFlexGrid2.TextMatrix(1, 6) = "Rent Amount"
''MSHFlexGrid2.TextMatrix(1, 7) = "Arrears"
''MSHFlexGrid2.TextMatrix(1, 8) = "Total Amount"
''MSHFlexGrid2.TextMatrix(1, 9) = "Current Stock"
''
''MSHFlexGrid2.TextMatrix(1, 10) = "Area in Sq. ft."
''MSHFlexGrid2.TextMatrix(1, 11) = "Agreement Expiry Date"
''MSHFlexGrid2.TextMatrix(1, 12) = "Godown Close Date"
''
''
''MSHFlexGrid2.TextMatrix(1, 13) = "Remarks"
''MSHFlexGrid2.TextMatrix(1, 14) = "GodownID"
''
''MSHFlexGrid2.ColWidth(0) = 800
''MSHFlexGrid2.ColWidth(1) = 1000
''MSHFlexGrid2.ColWidth(2) = 1500
''MSHFlexGrid2.ColWidth(3) = 900
''MSHFlexGrid2.ColWidth(4) = 3000
''
''MSHFlexGrid2.ColWidth(5) = 1200
''MSHFlexGrid2.ColWidth(6) = 1500
''MSHFlexGrid2.ColWidth(7) = 1200
''MSHFlexGrid2.ColWidth(8) = 1500
''MSHFlexGrid2.ColWidth(9) = 1500
''MSHFlexGrid2.ColWidth(10) = 1700
''MSHFlexGrid2.ColWidth(11) = 1500
'''MSHFlexGrid2.ColWidth(11) = 2300
''MSHFlexGrid2.ColWidth(12) = 2500
''MSHFlexGrid2.ColWidth(13) = 2500
''MSHFlexGrid2.ColWidth(14) = 0
''
''MSHFlexGrid2.RowHeight(1) = 500
''
''End Sub

Private Sub Grid_Head()

MSHFlexGrid2.Clear
MSHFlexGrid2.Rows = 3
MSHFlexGrid2.FixedRows = 2

If CDate(TxtExpenseFrom) < "5/01/2022" Then
   MSHFlexGrid2.Cols = 18
Else
   MSHFlexGrid2.Cols = 19
End If

MSHFlexGrid2.TextMatrix(1, 0) = "CMP Name"
MSHFlexGrid2.TextMatrix(1, 1) = "GSLID"
MSHFlexGrid2.TextMatrix(1, 2) = "Godown No"
MSHFlexGrid2.TextMatrix(1, 3) = "Stack No"
MSHFlexGrid2.TextMatrix(1, 4) = "Lot No"
MSHFlexGrid2.TextMatrix(1, 5) = "Commodity"
MSHFlexGrid2.TextMatrix(1, 6) = "No Of Bags"
MSHFlexGrid2.TextMatrix(1, 7) = "Quantity"
MSHFlexGrid2.TextMatrix(1, 8) = "Transaction Flag"
MSHFlexGrid2.TextMatrix(1, 9) = "CDTF No/CITF No"
MSHFlexGrid2.TextMatrix(1, 10) = "Franchisee Billing Cycle"
MSHFlexGrid2.TextMatrix(1, 11) = "Franchisee Billing Unit"
MSHFlexGrid2.TextMatrix(1, 12) = "No of Cycle"
MSHFlexGrid2.TextMatrix(1, 13) = "Franchisee Revenue Share"
MSHFlexGrid2.TextMatrix(1, 14) = "Frachisee Indiviual Rent Amount"
MSHFlexGrid2.TextMatrix(1, 15) = "Invoice Amount for this GSL"
MSHFlexGrid2.TextMatrix(1, 16) = "Franchisee Share for this GSL"
MSHFlexGrid2.TextMatrix(1, 17) = "Franchisee Indiviual Revenue Share"

If CDate(TxtExpenseFrom) >= "5/01/2022" Then
   MSHFlexGrid2.TextMatrix(1, 18) = "Insurance Amount"
End If

MSHFlexGrid2.RowHeight(1) = 700
MSHFlexGrid2.WordWrap = True

MSHFlexGrid2.ColWidth(0) = 1800
MSHFlexGrid2.ColWidth(1) = 0
MSHFlexGrid2.ColWidth(2) = 1000
MSHFlexGrid2.ColWidth(3) = 1000
MSHFlexGrid2.ColWidth(4) = 1000
MSHFlexGrid2.ColWidth(5) = 1600
MSHFlexGrid2.ColWidth(6) = 1000
MSHFlexGrid2.ColWidth(7) = 1000
MSHFlexGrid2.ColWidth(8) = 1000
MSHFlexGrid2.ColWidth(9) = 1000
MSHFlexGrid2.ColWidth(10) = 1000
MSHFlexGrid2.ColWidth(11) = 1000
MSHFlexGrid2.ColWidth(12) = 1000
MSHFlexGrid2.ColWidth(13) = 1000
MSHFlexGrid2.ColWidth(14) = 1000
MSHFlexGrid2.ColWidth(15) = 1000

End Sub

Private Sub Grid_Head1()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.FixedRows = 1
MSHFlexGrid1.Cols = 32

MSHFlexGrid1.TextMatrix(0, 0) = "Claim ID":                     MSHFlexGrid1.ColWidth(0) = 600
MSHFlexGrid1.TextMatrix(0, 1) = "Claim Date":                   MSHFlexGrid1.ColWidth(1) = 1000
MSHFlexGrid1.TextMatrix(0, 2) = "Expense From":                 MSHFlexGrid1.ColWidth(2) = 1000
MSHFlexGrid1.TextMatrix(0, 3) = "Expense To":                   MSHFlexGrid1.ColWidth(3) = 1000
MSHFlexGrid1.TextMatrix(0, 4) = "Franchisee Name":              MSHFlexGrid1.ColWidth(4) = 2000
MSHFlexGrid1.TextMatrix(0, 5) = "Location Name":                MSHFlexGrid1.ColWidth(5) = 2000
MSHFlexGrid1.TextMatrix(0, 6) = "Rent Amount":                  MSHFlexGrid1.ColWidth(6) = 1000
MSHFlexGrid1.TextMatrix(0, 7) = "Actual Amount":                MSHFlexGrid1.ColWidth(7) = 1000
MSHFlexGrid1.TextMatrix(0, 8) = "Claim Forwarded To":           MSHFlexGrid1.ColWidth(8) = 1800
MSHFlexGrid1.TextMatrix(0, 9) = "Status":                       MSHFlexGrid1.ColWidth(9) = 1000
MSHFlexGrid1.TextMatrix(0, 10) = "Booked in NAV Date":          MSHFlexGrid1.ColWidth(10) = 1000
MSHFlexGrid1.TextMatrix(0, 11) = "FAG Remarks":                 MSHFlexGrid1.ColWidth(11) = 2000
MSHFlexGrid1.TextMatrix(0, 12) = "Remarks":                     MSHFlexGrid1.ColWidth(12) = 2000
MSHFlexGrid1.TextMatrix(0, 13) = "Batch No":                    MSHFlexGrid1.ColWidth(13) = 1000
MSHFlexGrid1.TextMatrix(0, 14) = "G_UID":                       MSHFlexGrid1.ColWidth(14) = 0
MSHFlexGrid1.TextMatrix(0, 15) = "Bank Name":                   MSHFlexGrid1.ColWidth(15) = 2000
MSHFlexGrid1.TextMatrix(0, 16) = "Cheque No":                   MSHFlexGrid1.ColWidth(16) = 1000
MSHFlexGrid1.TextMatrix(0, 17) = "Paid Amount":                 MSHFlexGrid1.ColWidth(17) = 1000
MSHFlexGrid1.TextMatrix(0, 18) = "TDS ":                        MSHFlexGrid1.ColWidth(18) = 1000
MSHFlexGrid1.TextMatrix(0, 19) = "Paid Date":                   MSHFlexGrid1.ColWidth(19) = 1000
''-- 3rd Jan 2017 Ullhas Patil
MSHFlexGrid1.TextMatrix(0, 20) = "Disbursment Date":            MSHFlexGrid1.ColWidth(20) = 1000
MSHFlexGrid1.TextMatrix(0, 21) = "Disbursment Remarks":         MSHFlexGrid1.ColWidth(21) = 2000
MSHFlexGrid1.TextMatrix(0, 22) = "Disbursment Amount":          MSHFlexGrid1.ColWidth(22) = 1000
MSHFlexGrid1.TextMatrix(0, 23) = "Mode for Payment":            MSHFlexGrid1.ColWidth(23) = 1000
MSHFlexGrid1.TextMatrix(0, 24) = "Disbursment through Bank":    MSHFlexGrid1.ColWidth(24) = 1500
MSHFlexGrid1.TextMatrix(0, 25) = "Payment Batch ID":            MSHFlexGrid1.ColWidth(25) = 1000
MSHFlexGrid1.TextMatrix(0, 26) = "Payment Doc No":              MSHFlexGrid1.ColWidth(26) = 1000
MSHFlexGrid1.TextMatrix(0, 27) = "Payment Entry No":            MSHFlexGrid1.ColWidth(27) = 1000
MSHFlexGrid1.TextMatrix(0, 28) = "FAG Payment exported Date":   MSHFlexGrid1.ColWidth(28) = 1000
MSHFlexGrid1.TextMatrix(0, 29) = "Deduction Remark":            MSHFlexGrid1.ColWidth(29) = 2000
MSHFlexGrid1.TextMatrix(0, 30) = "Letter Print Flag":           MSHFlexGrid1.ColWidth(30) = 1500
''-- Ullhas Patil 25 Oct 2017
MSHFlexGrid1.TextMatrix(0, 31) = "Billing State India":          MSHFlexGrid1.ColWidth(31) = 2000
MSHFlexGrid1.RowHeight(0) = 500

End Sub
Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Txn_FranchiseeRentApproval-" & TxtClaimID & ".xls")
Call Xls_Detail_Rpt1(MSHFlexGrid1)
End Sub



Private Sub CmdExcelDetail_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Txn_FranchiseeRentApproval-" & TxtClaimID & ".xls")
Call Xls_Detail_Rpt1(MSHFlexGrid2)
End Sub


Private Sub Xls_Detail_Rpt1(MSH_Grid As MSHFlexGrid)
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Integer
Dim mCol As Integer
MsgBox ("Wait till next message")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Txn_FranchiseeRentApproval-" & TxtClaimID & ".xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSH_Grid.Rows - 1
    mRow = I + 1 ' + 6
    For C = 0 To MSH_Grid.Cols - 1
        If InStr(MSH_Grid.TextMatrix(I, C), "/") > 0 Then
           If IsDate(MSH_Grid.TextMatrix(I, C)) Then
              oWS.Cells(mRow, C + 1) = Format(MSH_Grid.TextMatrix(I, C), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, C + 1) = MSH_Grid.TextMatrix(I, C)
           End If
        Else
           oWS.Cells(mRow, C + 1) = MSH_Grid.TextMatrix(I, C)
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


'Private Sub SaveLessorArrearsDetails()
'
'Call TableTxn_FranchiseeRentDetail(" Where ClaimID = " & TxtClaimID & "")
'For I = 2 To MSHFlexGrid2.Rows - 1
'    If MSHFlexGrid2.TextMatrix(I, 1) <> "" Then
'       RsTxn_FranchiseeRentDetail.Find "ClaimDetailID = " & MSHFlexGrid2.TextMatrix(I, 0) & ""
'       If RsTxn_FranchiseeRentDetail.EOF = False Then
'          RsTxn_FranchiseeRentDetail!Arrears = Val(MSHFlexGrid2.TextMatrix(I, 7))
'          RsTxn_FranchiseeRentDetail.Update
'          RsTxn_FranchiseeRentDetail.Requery
'       End If
'    End If
'Next
'End Sub


Private Sub TxtActualFee_LostFocus()
If TxtActualFee = "" Then Exit Sub
If Not IsNumeric(TxtActualFee) Or Val(TxtActualFee) <= 0 Then
   MsgBox "Invalid numeric format !!!! "
   TxtActualFee.SetFocus
End If

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
Private Sub IndataGSLDetails()
Dim mYear, mMonth, mFromDate As Variant

Call Grid_Head

If CDate(RsTxn_FranchiseeRent!CreatedDate) < "9/24/2021" Then

    tmp = "Select MC.CMPName,TIMGD.Mst_GSL_ID,MG.GodownNo,TIMGD.StackNo,TIMGD.LotNo,MComm.Commodity,  TIMGD.BalanceBags,TIMGD.BalanceWeight,TIMGD.TFlag,TIMGD.CxTFNo,BillingCycle,TIMGD.FranchBillingUnit,TIMGD.FranchNoofDays,TIMGD.FranchiseeRate,  TIMGD.FranchInvoiceAmount,TIMGD.InvoiceAmount  " & _
          " From Txn_InvMonGSLDetail TIMGD " & _
          " Left Join Mst_CMP MC On TIMGD.CMPID = MC.CMPID" & _
          " Left Join Mst_Godown MG On TIMGD.GodownID = MG.GodownID" & _
          " Left Join Mst_Commodity MComm On TIMGD.CommodityID = MComm.CommodityID" & _
          " Left Join Mst_BillingCycle MBC On TIMGD.FranchBillingCycleID = MBC.BillingCycleID" & _
          " Where TIMGD.FranchClaimID = " & TxtClaimID & _
          " Order By MC.CMPName,MComm.Commodity,TIMGD.Mst_GSL_ID,MG.GodownNo,TIMGD.StackNo,TIMGD.LotNo "
ElseIf CDate(TxtExpenseFrom) < "6/01/2022" Then
    tmp = "Select MC.CMPName,TIMGD.Mst_GSL_ID,MG.GodownNo,TIMGD.StackNo,TIMGD.LotNo,MComm.Commodity,  TIMGD.BalanceBags,TIMGD.BalanceWeight,TIMGD.TFlag,TIMGD.CxTF_DetailsID,BillingCycle,TIMGD.FranchBillingUnit,TIMGD.FranchNoofDays,TIMGD.FranchiseeRate,  TMFRS.[ShareAmount],TIMGD.InvoiceAmount,TIMGD.FranchInvoiceAmount ,TMFRS.[VendorShare] "
    tmp = tmp & " From Txn_InvMonGSLDetail TIMGD "
    tmp = tmp & " Left Join Mst_CMP MC On TIMGD.CMPID = MC.CMPID"
    tmp = tmp & " Left Join Mst_Godown MG On TIMGD.GodownID = MG.GodownID"
    tmp = tmp & "  Left Join Mst_Commodity MComm On TIMGD.CommodityID = MComm.CommodityID"
    tmp = tmp & "  Left Join Mst_BillingCycle MBC On TIMGD.FranchBillingCycleID = MBC.BillingCycleID"
    tmp = tmp & "  Left Join Txn_MonthlyFranchiseeRentShare TMFRS On TMFRS.MonG_UID = TIMGD.MonG_UID And TMFRS.YearMonth = TIMGD.YearMonth"
    tmp = tmp & "  Where TMFRS.ClaimID = " & TxtClaimID & "  And TMFRS.VendorID = " & mVendorID & ""
    tmp = tmp & "  Order By MC.CMPName,MComm.Commodity,TIMGD.Mst_GSL_ID,MG.GodownNo,TIMGD.StackNo,TIMGD.LotNo "
Else
    tmp = "Select MC.CMPName,TIMGD.Mst_GSL_ID,MG.GodownNo,TIMGD.StackNo,TIMGD.LotNo,MComm.Commodity,  TIMGD.BalanceBags,TIMGD.BalanceWeight,TIMGD.TFlag,TIMGD.CxTF_DetailsID,BillingCycle,TIMGD.FranchBillingUnit,TIMGD.FranchNoofDays,TIMGD.FranchiseeRate,  TMFRS.[ShareAmount],TIMGD.InvoiceAmtForFranch,TIMGD.FranchInvoiceAmount ,TMFRS.[VendorShare],TIMGD.InsuranceCostDeduction "
    tmp = tmp & " From Txn_InvMonGSLDetail TIMGD "
    tmp = tmp & " Left Join Mst_CMP MC On TIMGD.CMPID = MC.CMPID"
    tmp = tmp & " Left Join Mst_Godown MG On TIMGD.GodownID = MG.GodownID"
    tmp = tmp & "  Left Join Mst_Commodity MComm On TIMGD.CommodityID = MComm.CommodityID"
    tmp = tmp & "  Left Join Mst_BillingCycle MBC On TIMGD.FranchBillingCycleID = MBC.BillingCycleID"
    tmp = tmp & "  Left Join Txn_MonthlyFranchiseeRentShare TMFRS On TMFRS.MonG_UID = TIMGD.MonG_UID And TMFRS.YearMonth = TIMGD.YearMonth"
    tmp = tmp & "  Where TMFRS.ClaimID = " & TxtClaimID & "  And TMFRS.VendorID = " & mVendorID & ""
    tmp = tmp & "  Order By MC.CMPName,MComm.Commodity,TIMGD.Mst_GSL_ID,MG.GodownNo,TIMGD.StackNo,TIMGD.LotNo "
End If

Call TmpTable

If TmpRs.RecordCount > 0 Then

   x = 2
   For I = 1 To TmpRs.RecordCount
       MSHFlexGrid2.Rows = TmpRs.RecordCount + 2
       For C = 0 To TmpRs.Fields.Count - 1
           MSHFlexGrid2.TextMatrix(x, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))

       Next
       If LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "superadmin" And LCase(Gen_UserRole) <> "headcmg" Then
           MSHFlexGrid2.TextMatrix(x, 18) = 0
       End If
       If LCase(MSHFlexGrid2.TextMatrix(x, 8)) = "o" Then
          MSHFlexGrid2.TextMatrix(x, 8) = "Op. Balance"
       ElseIf LCase(MSHFlexGrid2.TextMatrix(x, 8)) = "d" Then
          MSHFlexGrid2.TextMatrix(x, 8) = "Deposit"
       ElseIf LCase(MSHFlexGrid2.TextMatrix(x, 8)) = "w" Then
          MSHFlexGrid2.TextMatrix(x, 8) = "Withdrawal"
       End If
       
       If LCase(MSHFlexGrid2.TextMatrix(x, 11)) = "q" Then
          MSHFlexGrid2.TextMatrix(x, 11) = "Quantity"
       ElseIf LCase(MSHFlexGrid2.TextMatrix(x, 11)) = "b" Then
          MSHFlexGrid2.TextMatrix(x, 11) = "Bags"
       End If
       
       'InsuranceCostDeduction
       
       MSHFlexGrid2.TextMatrix(0, 6) = Val(MSHFlexGrid2.TextMatrix(0, 6)) + Val(MSHFlexGrid2.TextMatrix(x, 6))
       MSHFlexGrid2.TextMatrix(0, 7) = Val(MSHFlexGrid2.TextMatrix(0, 7)) + Val(MSHFlexGrid2.TextMatrix(x, 7))
       
       MSHFlexGrid2.TextMatrix(0, 14) = Val(MSHFlexGrid2.TextMatrix(0, 14)) + Val(MSHFlexGrid2.TextMatrix(x, 14))
       MSHFlexGrid2.TextMatrix(0, 15) = Val(MSHFlexGrid2.TextMatrix(0, 15)) + Val(MSHFlexGrid2.TextMatrix(x, 15))
       

       
       If CDate(RsTxn_FranchiseeRent!ExpenseFrom) >= "5/01/2022" Then
          MSHFlexGrid2.TextMatrix(0, 18) = Val(MSHFlexGrid2.TextMatrix(0, 18)) + Val(MSHFlexGrid2.TextMatrix(x, 18))
       End If
       
       x = x + 1
       TmpRs.MoveNext
   Next

End If

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
If CmbServiceTaxID.Text = "" Then
   mServiceTaxID = Null
   Exit Sub
End If
RsMst_ServiceTax.MoveFirst
RsMst_ServiceTax.Find "Description = '" & CmbServiceTaxID.Text & "'"
If RsMst_ServiceTax.EOF Then
   MsgBox "Invalid selection "
   CmbServiceTaxID.SetFocus
   Exit Sub
End If
mServiceTaxID = RsMst_ServiceTax!ServiceTaxID

End Sub

Private Sub UpdateFranchClaimID()
If LCase(mStatus) = "r" Then
   If CDate(RsTxn_FranchiseeRent!CreatedDate) < "9/24/2021" Then
      tmp = "Update Txn_InvMonGSLDetail set FranchClaimID = null Where FranchClaimID = " & TxtClaimID
   Else
      tmp = "Delete Txn_MonthlyFranchiseeRentShare Where ClaimID = " & TxtClaimID
   End If
   Call TmpTable
ElseIf LCase(mStatus) = "f" Then
   If CDate(RsTxn_FranchiseeRent!CreatedDate) < "9/24/2021" Then
      tmp = "Update Txn_InvMonGSLDetail set FranchInvoiceAmount = 0 Where FranchClaimID = " & TxtClaimID
   Else
      tmp = "Update Txn_MonthlyFranchiseeRentShare set ShareAmount = 0 Where ClaimID = " & TxtClaimID
   End If
   Call TmpTable
End If

End Sub

Private Sub SaveLogFranchClaimDetail()
      
Call TableLog_FranchiseeFeeRejectHistory

For I = 2 To MSHFlexGrid2.Rows - 1
    If MSHFlexGrid2.TextMatrix(I, 0) <> "" Then
       RsLog_FranchiseeFeeRejectHistory.AddNew
       RsLog_FranchiseeFeeRejectHistory!ClaimID = TxtClaimID
       RsLog_FranchiseeFeeRejectHistory!CMPName = MSHFlexGrid2.TextMatrix(I, 0)
       RsLog_FranchiseeFeeRejectHistory!GSLID = MSHFlexGrid2.TextMatrix(I, 1)
       RsLog_FranchiseeFeeRejectHistory!GodownNo = MSHFlexGrid2.TextMatrix(I, 2)
       RsLog_FranchiseeFeeRejectHistory!StackNo = MSHFlexGrid2.TextMatrix(I, 3)
       RsLog_FranchiseeFeeRejectHistory!LotNo = MSHFlexGrid2.TextMatrix(I, 4)
       RsLog_FranchiseeFeeRejectHistory!Commodity = MSHFlexGrid2.TextMatrix(I, 5)
       RsLog_FranchiseeFeeRejectHistory!NoofBags = MSHFlexGrid2.TextMatrix(I, 6)
       RsLog_FranchiseeFeeRejectHistory!Quantity = MSHFlexGrid2.TextMatrix(I, 7)
       RsLog_FranchiseeFeeRejectHistory!TransactionFlag = Left(MSHFlexGrid2.TextMatrix(I, 8), 1)
       RsLog_FranchiseeFeeRejectHistory!CxTFNO = MSHFlexGrid2.TextMatrix(I, 9)
       RsLog_FranchiseeFeeRejectHistory!BillingCycle = MSHFlexGrid2.TextMatrix(I, 10)
       RsLog_FranchiseeFeeRejectHistory!BillingUnit = MSHFlexGrid2.TextMatrix(I, 11)
       RsLog_FranchiseeFeeRejectHistory!NoofCycle = Val(MSHFlexGrid2.TextMatrix(I, 12))
       RsLog_FranchiseeFeeRejectHistory!RateperCycle = Val(MSHFlexGrid2.TextMatrix(I, 13))
       RsLog_FranchiseeFeeRejectHistory!Amount = Val(MSHFlexGrid2.TextMatrix(I, 14))
       'Added on 30 Dec 2013 after mail from Kuldip Sir to show Invoice amount in grid
       RsLog_FranchiseeFeeRejectHistory!InvoiceAmount = Val(MSHFlexGrid2.TextMatrix(I, 15))
       
       RsLog_FranchiseeFeeRejectHistory.Update
       RsLog_FranchiseeFeeRejectHistory.Requery
    End If
Next
End Sub


Private Sub InDataLogFranchClaimDetail()
Call Grid_Head
Call TableLog_FranchiseeFeeRejectHistory(" Where ClaimID= " & Val(TxtClaimID))
If RsLog_FranchiseeFeeRejectHistory.RecordCount > 0 Then
    MSHFlexGrid2.Rows = RsLog_FranchiseeFeeRejectHistory.RecordCount + 2
    For I = 1 To RsLog_FranchiseeFeeRejectHistory.RecordCount
        MSHFlexGrid2.TextMatrix(I + 1, 0) = RsLog_FranchiseeFeeRejectHistory!CMPName
        MSHFlexGrid2.TextMatrix(I + 1, 1) = RsLog_FranchiseeFeeRejectHistory!GSLID
        MSHFlexGrid2.TextMatrix(I + 1, 2) = RsLog_FranchiseeFeeRejectHistory!GodownNo
        MSHFlexGrid2.TextMatrix(I + 1, 3) = RsLog_FranchiseeFeeRejectHistory!StackNo
        MSHFlexGrid2.TextMatrix(I + 1, 4) = RsLog_FranchiseeFeeRejectHistory!LotNo
        MSHFlexGrid2.TextMatrix(I + 1, 5) = RsLog_FranchiseeFeeRejectHistory!Commodity
        MSHFlexGrid2.TextMatrix(I + 1, 6) = RsLog_FranchiseeFeeRejectHistory!NoofBags
        MSHFlexGrid2.TextMatrix(I + 1, 7) = RsLog_FranchiseeFeeRejectHistory!Quantity
         
        MSHFlexGrid2.TextMatrix(I + 1, 8) = RsLog_FranchiseeFeeRejectHistory!TransactionFlag
        MSHFlexGrid2.TextMatrix(I + 1, 9) = IIf(IsNull(RsLog_FranchiseeFeeRejectHistory!CxTFNO), "", RsLog_FranchiseeFeeRejectHistory!CxTFNO)
        MSHFlexGrid2.TextMatrix(I + 1, 10) = RsLog_FranchiseeFeeRejectHistory!BillingCycle
        MSHFlexGrid2.TextMatrix(I + 1, 11) = RsLog_FranchiseeFeeRejectHistory!BillingUnit
        MSHFlexGrid2.TextMatrix(I + 1, 12) = RsLog_FranchiseeFeeRejectHistory!NoofCycle
        MSHFlexGrid2.TextMatrix(I + 1, 13) = RsLog_FranchiseeFeeRejectHistory!RateperCycle
        MSHFlexGrid2.TextMatrix(I + 1, 14) = RsLog_FranchiseeFeeRejectHistory!Amount
        MSHFlexGrid2.TextMatrix(I + 1, 15) = IIf(IsNull(RsLog_FranchiseeFeeRejectHistory!InvoiceAmount), "", RsLog_FranchiseeFeeRejectHistory!InvoiceAmount)
        
        MSHFlexGrid2.TextMatrix(0, 14) = Val(MSHFlexGrid2.TextMatrix(0, 14)) + Val(MSHFlexGrid2.TextMatrix(I + 1, 14))
        MSHFlexGrid2.TextMatrix(0, 15) = Val(MSHFlexGrid2.TextMatrix(0, 15)) + Val(MSHFlexGrid2.TextMatrix(I + 1, 15))
        
        RsLog_FranchiseeFeeRejectHistory.MoveNext
    Next
End If
End Sub

Private Function CheckDuplicateBill() As Boolean
Dim Retval As Boolean
Retval = False
tmp = " Select * from Txn_FranchiseeRent Where FranchInvoiceNo='" & TxtFranchInvNo.Text & "' and FranchInvoiceDate='" & TxtFranchInvDate.Value & "' And Flag Not in ('R','F') And VendorID=" & mVendorID
If Edit_Flg = "E" Then
   tmp = tmp & " And ClaimID<>" & Val(TxtClaimID) & ""
End If
Call TmpTable
If TmpRs.RecordCount > 0 Then
   Retval = True
End If
CheckDuplicateBill = Retval
End Function

''-- 03rd Jan 2017------
Private Sub TxtFAGDisbAmt_LostFocus()
If TxtFAGDisbAmt = "" Then Exit Sub
If IsNumeric(TxtFAGDisbAmt) = False Or Val(TxtFAGDisbAmt) < 0 Then
   MsgBox "Invalid format!!!"
   TxtFAGDisbAmt.SetFocus
   Exit Sub
End If
If Val(TxtFAGDisbAmt) > Val(TxtActualFee) Then
   MsgBox "Disbusment amount should be less than or equal to Claim Amount!!!"
   TxtFAGDisbAmt.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmbModeOfPay_LostFocus()
If CmbModeOfPay.Text = "" Then Exit Sub
If LCase(CmbModeOfPay.Text) <> "rtgs" And LCase(CmbModeOfPay.Text) <> "neft" And LCase(CmbModeOfPay.Text) <> "bank transfer" And LCase(CmbModeOfPay.Text) <> "cheque" Then
   MsgBox "Invalid Selection!!!"
   CmbModeOfPay.SetFocus
   Exit Sub
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
   CmbEmployeeID.SetFocus
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

Private Sub VendorBankDetails(mVendorID_ As Double)
TxtBankName = ""
TxtBranchName = ""
TxtACNo = ""
TxtIFSCCode = ""
If mVendorID_ > 0 Then
    tmp = " SELECT MV.VendorID,MV.VendorName,MV.BankID,MV.BranchID,MV.BankAcNo,MV.IFSCCode,MBB.BranchName,MB.BankName"
    tmp = tmp & " FROM Mst_Vendor MV"
    tmp = tmp & " Left Join Mst_Branch MBB On MV.BranchID = MBB.BranchId"
    tmp = tmp & " Left Join Mst_Bank MB On MV.BankID = MB.BankId"
    tmp = tmp & " Where MV.VendorID = " & mVendorID_ & " "
    Call TmpTable
    If TmpRs.RecordCount > 0 Then
        TxtBankName = IIf(IsNull(TmpRs!BankName), "", TmpRs!BankName)
        TxtBranchName = IIf(IsNull(TmpRs!BranchName), "", TmpRs!BranchName)
        TxtACNo = IIf(IsNull(TmpRs!BankAcNo), "", TmpRs!BankAcNo)
        TxtIFSCCode = IIf(IsNull(TmpRs!IFSCCode), "", TmpRs!IFSCCode)
    End If
End If
End Sub
