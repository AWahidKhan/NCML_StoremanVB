VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmTxn_CRAdvanceAllocation 
   Caption         =   "CR Advance Allocation"
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
      TabIndex        =   77
      Top             =   9720
      Width           =   10935
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   9015
      Left            =   15240
      TabIndex        =   76
      Top             =   1440
      Width           =   255
   End
   Begin VB.Frame Frame4 
      Height          =   9495
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   15375
      Begin VB.Frame Frame0 
         Height          =   2085
         Left            =   2400
         TabIndex        =   4
         Top             =   4440
         Visible         =   0   'False
         Width           =   14880
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
            Left            =   9090
            TabIndex        =   12
            Top             =   345
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
            Left            =   8055
            TabIndex        =   11
            Top             =   345
            Width           =   1000
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
            ItemData        =   "FrmTxn_CRAdvanceAllocation.frx":0000
            Left            =   2845
            List            =   "FrmTxn_CRAdvanceAllocation.frx":0002
            Sorted          =   -1  'True
            TabIndex        =   10
            Top             =   375
            Width           =   3165
         End
         Begin VB.TextBox TxtSCRID 
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
            Left            =   6120
            MaxLength       =   15
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   360
            Width           =   1005
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   1095
            Left            =   120
            TabIndex        =   13
            Top             =   840
            Width           =   10215
            _ExtentX        =   18018
            _ExtentY        =   1931
            _Version        =   393216
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSComCtl2.DTPicker STo 
            Height          =   360
            Left            =   1475
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   375
            Width           =   1305
            _ExtentX        =   2302
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
            Format          =   8388609
            CurrentDate     =   36494
         End
         Begin MSComCtl2.DTPicker SFrom 
            Height          =   360
            Left            =   105
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   375
            Width           =   1305
            _ExtentX        =   2302
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
            Format          =   8388609
            CurrentDate     =   36494
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
            Left            =   1800
            TabIndex        =   19
            Top             =   120
            Width           =   855
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
            Left            =   4080
            TabIndex        =   18
            Top             =   120
            Width           =   960
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
            Left            =   360
            TabIndex        =   17
            Top             =   120
            Width           =   855
         End
         Begin VB.Label Label19 
            Caption         =   "CR ID"
            Height          =   255
            Left            =   6480
            TabIndex        =   16
            Top             =   120
            Width           =   525
         End
      End
      Begin VB.Frame ButtonFrm 
         Height          =   1095
         Left            =   135
         TabIndex        =   43
         Top             =   8310
         Width           =   14985
         Begin VB.CommandButton CmdReturn 
            Caption         =   "Return"
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
            Height          =   810
            Left            =   16890
            TabIndex        =   91
            Top             =   300
            Visible         =   0   'False
            Width           =   1275
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
            Left            =   225
            TabIndex        =   55
            Top             =   585
            Width           =   2835
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
            Left            =   3060
            TabIndex        =   54
            Top             =   585
            Width           =   2835
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
            Left            =   5895
            TabIndex        =   53
            Top             =   585
            Width           =   2835
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
            Left            =   8730
            TabIndex        =   52
            Top             =   585
            Width           =   2835
         End
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
            Left            =   11565
            TabIndex        =   51
            Top             =   585
            Width           =   2835
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
            Left            =   225
            TabIndex        =   50
            Top             =   165
            Width           =   2835
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
            Left            =   3060
            TabIndex        =   49
            Top             =   165
            Width           =   2835
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
            Left            =   5895
            TabIndex        =   48
            Top             =   165
            Width           =   2835
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
            Left            =   8730
            TabIndex        =   47
            Top             =   165
            Width           =   2835
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
            Left            =   11565
            TabIndex        =   46
            Top             =   165
            Width           =   2835
         End
         Begin VB.CommandButton CmdPrint 
            Caption         =   "&Print"
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
            Height          =   400
            Left            =   14640
            TabIndex        =   45
            Top             =   705
            Visible         =   0   'False
            Width           =   2235
         End
         Begin VB.CommandButton CmdPrintPreview 
            Caption         =   "Print Preview"
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
            Height          =   400
            Left            =   14640
            TabIndex        =   44
            Top             =   285
            Visible         =   0   'False
            Width           =   2235
         End
      End
      Begin VB.Frame Frame1 
         Height          =   4080
         Left            =   120
         TabIndex        =   20
         Top             =   120
         Width           =   14895
         Begin VB.TextBox TxtChequeReturnChargesAmount 
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
            Left            =   8040
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   99
            TabStop         =   0   'False
            Top             =   1620
            Width           =   1560
         End
         Begin VB.TextBox TxtCRG_UID 
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
            Left            =   13080
            Locked          =   -1  'True
            TabIndex        =   97
            TabStop         =   0   'False
            Top             =   1200
            Visible         =   0   'False
            Width           =   1560
         End
         Begin VB.TextBox TxtUnAllocAmt 
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
            Left            =   13320
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   94
            TabStop         =   0   'False
            Top             =   3510
            Width           =   1440
         End
         Begin VB.TextBox TxtAllocAmt 
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
            Left            =   13320
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   92
            TabStop         =   0   'False
            Top             =   3060
            Width           =   1440
         End
         Begin VB.TextBox TxtInstrumentDate 
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
            Left            =   11400
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   89
            TabStop         =   0   'False
            Top             =   1170
            Width           =   1605
         End
         Begin VB.TextBox TxtCRDate 
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
            Left            =   8040
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   88
            TabStop         =   0   'False
            Top             =   720
            Width           =   1560
         End
         Begin VB.TextBox TxtCRAAID 
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
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   87
            TabStop         =   0   'False
            Top             =   240
            Width           =   1605
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
            Left            =   6120
            Locked          =   -1  'True
            TabIndex        =   86
            TabStop         =   0   'False
            Top             =   3060
            Width           =   4860
         End
         Begin VB.TextBox TxtInstrumentNo 
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
            Left            =   11400
            Locked          =   -1  'True
            MaxLength       =   25
            TabIndex        =   85
            TabStop         =   0   'False
            Top             =   720
            Width           =   1605
         End
         Begin VB.TextBox TxtRemark 
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
            Height          =   435
            Left            =   1890
            Locked          =   -1  'True
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   84
            TabStop         =   0   'False
            Top             =   2535
            Width           =   9360
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
            Left            =   1890
            Locked          =   -1  'True
            MaxLength       =   150
            TabIndex        =   83
            TabStop         =   0   'False
            Top             =   2100
            Width           =   12885
         End
         Begin VB.TextBox TxtReceivedFrom 
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
            Left            =   1890
            Locked          =   -1  'True
            MaxLength       =   100
            ScrollBars      =   2  'Vertical
            TabIndex        =   82
            TabStop         =   0   'False
            Top             =   1170
            Width           =   4380
         End
         Begin VB.ComboBox CmbCRID 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            ItemData        =   "FrmTxn_CRAdvanceAllocation.frx":0004
            Left            =   1890
            List            =   "FrmTxn_CRAdvanceAllocation.frx":0006
            TabIndex        =   2
            Top             =   720
            Width           =   1605
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
            Left            =   11850
            TabIndex        =   79
            TabStop         =   0   'False
            Top             =   270
            Width           =   375
         End
         Begin VB.TextBox TxtInstrumentAmount 
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
            Left            =   4830
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   720
            Width           =   1440
         End
         Begin VB.TextBox TxtPreFix 
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
            Left            =   13320
            Locked          =   -1  'True
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   240
            Width           =   1440
         End
         Begin VB.TextBox TxtDrawnOn 
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
            Left            =   1890
            Locked          =   -1  'True
            MaxLength       =   150
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   1620
            Width           =   4350
         End
         Begin VB.TextBox TxtDepositeDate 
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
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   3060
            Width           =   1560
         End
         Begin VB.TextBox TxtFAGPostingDate 
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
            Left            =   13320
            Locked          =   -1  'True
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   2535
            Width           =   1440
         End
         Begin VB.TextBox TxtRemark1 
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
            Height          =   435
            Left            =   1890
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   3510
            Width           =   9360
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
            ItemData        =   "FrmTxn_CRAdvanceAllocation.frx":0008
            Left            =   8040
            List            =   "FrmTxn_CRAdvanceAllocation.frx":000A
            Sorted          =   -1  'True
            TabIndex        =   1
            Top             =   240
            Width           =   3795
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
            Left            =   8040
            Locked          =   -1  'True
            TabIndex        =   78
            TabStop         =   0   'False
            Top             =   240
            Width           =   3780
         End
         Begin VB.TextBox TxtCRID 
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
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   80
            TabStop         =   0   'False
            Top             =   720
            Width           =   1605
         End
         Begin VB.TextBox TxtInstrumentType 
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
            Left            =   8040
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   75
            TabStop         =   0   'False
            Top             =   1170
            Width           =   1560
         End
         Begin MSComCtl2.DTPicker TxtAllocDate 
            Height          =   360
            Left            =   4830
            TabIndex        =   0
            TabStop         =   0   'False
            Top             =   240
            Width           =   1440
            _ExtentX        =   2540
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
            Format          =   8388609
            CurrentDate     =   36494
         End
         Begin VB.Label Label52 
            Caption         =   "Cheque Return Charges Amount"
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
            Left            =   6480
            TabIndex        =   98
            Top             =   1553
            Width           =   1800
         End
         Begin VB.Label Label10 
            Caption         =   "Unallocated Amount"
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
            Left            =   11400
            TabIndex        =   95
            Top             =   3585
            Width           =   1815
         End
         Begin VB.Label Label3 
            Caption         =   "Allocated Amount"
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
            Left            =   11370
            TabIndex        =   93
            Top             =   3120
            Width           =   1815
         End
         Begin VB.Label Label14 
            Caption         =   "Alloc Date"
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
            Top             =   255
            Width           =   1080
         End
         Begin VB.Label Label9 
            Caption         =   "Adv Alloc ID"
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
            Left            =   240
            TabIndex        =   81
            Top             =   270
            Width           =   1440
         End
         Begin VB.Label LblLocationID 
            Caption         =   "Receipt No"
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
            Left            =   240
            TabIndex        =   42
            Top             =   720
            Width           =   1440
         End
         Begin VB.Label LblStateID 
            Caption         =   "Drawn On"
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
            Left            =   240
            TabIndex        =   41
            Top             =   1635
            Width           =   1080
         End
         Begin VB.Label LblFlag 
            Caption         =   "Receipt Date"
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
            Left            =   6480
            TabIndex        =   40
            Top             =   750
            Width           =   1200
         End
         Begin VB.Label Label1 
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
            Height          =   255
            Left            =   240
            TabIndex        =   39
            Top             =   2160
            Width           =   1320
         End
         Begin VB.Label Label2 
            Caption         =   "Instrument Type"
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
            Left            =   6480
            TabIndex        =   38
            Top             =   1200
            Width           =   1440
         End
         Begin VB.Label Label4 
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
            Left            =   240
            TabIndex        =   37
            Top             =   2595
            Width           =   840
         End
         Begin VB.Label Label5 
            Caption         =   "Instrument No"
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
            Left            =   9870
            TabIndex        =   36
            Top             =   750
            Width           =   1320
         End
         Begin VB.Label Label6 
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
            Height          =   300
            Left            =   3720
            TabIndex        =   35
            Top             =   750
            Width           =   960
         End
         Begin VB.Label Label7 
            Caption         =   "Instrument Date"
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
            Left            =   9855
            TabIndex        =   34
            Top             =   1215
            Width           =   1440
         End
         Begin VB.Label Label8 
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
            Height          =   285
            Left            =   6510
            TabIndex        =   33
            Top             =   255
            Width           =   960
         End
         Begin VB.Label Label29 
            Caption         =   "Received from"
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
            Left            =   240
            TabIndex        =   32
            Top             =   1155
            Width           =   1455
         End
         Begin VB.Label Label40 
            Caption         =   "Deposited Date"
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
            Left            =   240
            TabIndex        =   31
            Top             =   3120
            Width           =   1695
         End
         Begin VB.Label Label15 
            Caption         =   "FAG Posting Date"
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
            Left            =   11400
            TabIndex        =   30
            Top             =   2640
            Width           =   1815
         End
         Begin VB.Label Label39 
            Caption         =   "Deposited Bank Name"
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
            Left            =   3840
            TabIndex        =   29
            Top             =   3120
            Width           =   2175
         End
         Begin VB.Label Label35 
            Caption         =   "Deposit Remark "
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
            Left            =   240
            TabIndex        =   28
            Top             =   3600
            Width           =   1695
         End
         Begin VB.Label Label16 
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
            Height          =   240
            Left            =   12480
            TabIndex        =   27
            Top             =   270
            Width           =   975
         End
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   3990
         Left            =   120
         TabIndex        =   56
         Top             =   4320
         Width           =   14895
         _ExtentX        =   26273
         _ExtentY        =   7038
         _Version        =   393216
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
         TabCaption(0)   =   "Client Details"
         TabPicture(0)   =   "FrmTxn_CRAdvanceAllocation.frx":000C
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "MSHFlexGrid2"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "TI Details"
         TabPicture(1)   =   "FrmTxn_CRAdvanceAllocation.frx":0028
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame2"
         Tab(1).Control(1)=   "MSHFlexGrid3"
         Tab(1).ControlCount=   2
         TabCaption(2)   =   "Log Details"
         TabPicture(2)   =   "FrmTxn_CRAdvanceAllocation.frx":0044
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "TxtCreated"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "TxtUpdated"
         Tab(2).Control(1).Enabled=   0   'False
         Tab(2).Control(2)=   "Label13"
         Tab(2).Control(3)=   "Label12"
         Tab(2).ControlCount=   4
         Begin VB.Frame Frame2 
            Height          =   1065
            Left            =   -74880
            TabIndex        =   59
            Top             =   360
            Width           =   14655
            Begin VB.TextBox TxtTILocation 
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
               Left            =   10290
               Locked          =   -1  'True
               TabIndex        =   100
               TabStop         =   0   'False
               Top             =   540
               Width           =   2625
            End
            Begin VB.TextBox TxtNAVCode 
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
               Left            =   14280
               Locked          =   -1  'True
               TabIndex        =   96
               TabStop         =   0   'False
               Top             =   240
               Visible         =   0   'False
               Width           =   1680
            End
            Begin VB.TextBox TxtTDSAmt 
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
               Left            =   14040
               MaxLength       =   15
               TabIndex        =   6
               Top             =   540
               Visible         =   0   'False
               Width           =   2010
            End
            Begin VB.TextBox TxtBalAmount 
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
               Left            =   4020
               TabIndex        =   63
               Top             =   540
               Width           =   1860
            End
            Begin VB.CommandButton CmdAdd2List 
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
               Height          =   375
               Left            =   13200
               TabIndex        =   7
               Top             =   480
               Width           =   1335
            End
            Begin VB.TextBox TxtRecAmt 
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
               Left            =   6000
               MaxLength       =   15
               TabIndex        =   5
               Top             =   540
               Width           =   2040
            End
            Begin VB.TextBox TxtBalRec 
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
               Left            =   8160
               Locked          =   -1  'True
               TabIndex        =   62
               TabStop         =   0   'False
               Top             =   540
               Width           =   2025
            End
            Begin VB.TextBox TxtTINo 
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
               Left            =   90
               Locked          =   -1  'True
               TabIndex        =   61
               TabStop         =   0   'False
               Top             =   540
               Width           =   1800
            End
            Begin VB.TextBox TxtTIDate 
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
               Left            =   2010
               Locked          =   -1  'True
               TabIndex        =   60
               TabStop         =   0   'False
               Top             =   540
               Width           =   1890
            End
            Begin VB.TextBox TxtOtherDed 
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
               Left            =   13650
               MaxLength       =   15
               TabIndex        =   8
               Top             =   540
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.Label Label56 
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
               Left            =   11205
               TabIndex        =   101
               Top             =   210
               Width           =   765
            End
            Begin VB.Label Label46 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               Caption         =   "TDS Deducted"
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
               Left            =   14310
               TabIndex        =   70
               Top             =   210
               Visible         =   0   'False
               Width           =   1560
            End
            Begin VB.Label LblPIDate 
               AutoSize        =   -1  'True
               Caption         =   "TI Date"
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
               Left            =   2640
               TabIndex        =   69
               Top             =   210
               Width           =   1020
            End
            Begin VB.Label LblInvoiceAmount 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               Caption         =   "Balance Amt."
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
               Left            =   4200
               TabIndex        =   68
               Top             =   210
               Width           =   1635
            End
            Begin VB.Label LblPINo 
               AutoSize        =   -1  'True
               Caption         =   "TI No."
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
               Left            =   735
               TabIndex        =   67
               Top             =   210
               Width           =   645
            End
            Begin VB.Label LblTDSAmt 
               AutoSize        =   -1  'True
               Caption         =   "Received Amt"
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
               TabIndex        =   66
               Top             =   210
               Width           =   1590
            End
            Begin VB.Label LblChequeAmt 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               Caption         =   "Balance Receivable"
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
               Left            =   8235
               TabIndex        =   65
               Top             =   210
               Width           =   1875
            End
            Begin VB.Label Label17 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               Caption         =   "Other Deduction"
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
               Left            =   13800
               TabIndex        =   64
               Top             =   210
               Visible         =   0   'False
               Width           =   1440
            End
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
            Left            =   -73935
            Locked          =   -1  'True
            TabIndex        =   58
            TabStop         =   0   'False
            Top             =   600
            Width           =   8970
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
            Left            =   -73935
            Locked          =   -1  'True
            TabIndex        =   57
            TabStop         =   0   'False
            Top             =   1185
            Width           =   8970
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
            Height          =   3330
            Left            =   120
            TabIndex        =   71
            Top             =   480
            Width           =   14670
            _ExtentX        =   25876
            _ExtentY        =   5874
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
            Height          =   2415
            Left            =   -74880
            TabIndex        =   72
            Top             =   1440
            Width           =   14670
            _ExtentX        =   25876
            _ExtentY        =   4260
            _Version        =   393216
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
         Begin VB.Label Label13 
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
            TabIndex        =   74
            Top             =   660
            Width           =   1695
         End
         Begin VB.Label Label12 
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
            TabIndex        =   73
            Top             =   1245
            Width           =   975
         End
      End
   End
End
Attribute VB_Name = "FrmTxn_CRAdvanceAllocation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mAgentID, mClientIDRow, mSDepoBankID As Variant
Dim mExchangeTypeID As Double
Dim mLocationID As Double
Dim Edit_Flg As Variant
Dim mCRID As Double
Dim mRow As Variant
Dim Wc1 As String
Dim mGSTNAVCode As Variant
Private Sub Grid_Head1()
MSHFlexGrid2.Clear
MSHFlexGrid2.Rows = 3
MSHFlexGrid2.FixedRows = 2
MSHFlexGrid2.Cols = 14

MSHFlexGrid2.TextMatrix(1, 0) = "ExchangeTypeID"
MSHFlexGrid2.TextMatrix(1, 1) = "Exchange Type"
MSHFlexGrid2.TextMatrix(1, 2) = "Client ID"
MSHFlexGrid2.TextMatrix(1, 3) = "Client Name"
MSHFlexGrid2.TextMatrix(1, 4) = "DP ID"
MSHFlexGrid2.TextMatrix(1, 5) = "DP Name"
MSHFlexGrid2.TextMatrix(1, 6) = "Amount"
MSHFlexGrid2.TextMatrix(1, 7) = "Remark"
MSHFlexGrid2.TextMatrix(1, 8) = "ClientRowID"
MSHFlexGrid2.TextMatrix(1, 9) = "ShowDP"
MSHFlexGrid2.TextMatrix(1, 10) = "NAVCode"
MSHFlexGrid2.TextMatrix(1, 11) = "RNAVCode"
MSHFlexGrid2.TextMatrix(1, 12) = "Alloc. Amt"
MSHFlexGrid2.TextMatrix(1, 13) = "Rem. Amt"

MSHFlexGrid2.ColWidth(0) = 0
MSHFlexGrid2.ColWidth(1) = 2000
MSHFlexGrid2.ColWidth(2) = 1400
MSHFlexGrid2.ColWidth(3) = 4200
MSHFlexGrid2.ColWidth(4) = 1400
MSHFlexGrid2.ColWidth(5) = 2200

MSHFlexGrid2.ColWidth(7) = 3800
MSHFlexGrid2.ColWidth(8) = 0
MSHFlexGrid2.ColWidth(9) = 0
End Sub
Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)
Frame1.Enabled = True
Frame2.Enabled = True
MSHFlexGrid2.Enabled = True
Call Grid_Head1
Call Grid_Head2
TxtInstrumentDate.Enabled = True
TxtInstrumentType.Visible = True
TxtG_UID = ""
CmbLocationID.Visible = True
TxtLocationID.Visible = False
'TxtCRID.Locked = True
TxtInstrumentNo.Locked = False
TxtDrawnOn.Locked = False
TxtBranchName.Locked = False
TxtRemark.Locked = False
CmbCRID.Visible = True
TxtCRID.Visible = False
SSTab1.Tab = 0
CmbLocationID.Text = ""
Call ClearHeader
Call ClearTITabControl
TxtCreated = ""
TxtUpdated = ""
TxtPreFix = ""
TxtCRAAID = ""
TxtAllocDate.Value = Date
CmdPrint.Enabled = False

If RsTxn_CRAdvanceAllocation.RecordCount > 0 Then
   TxtAllocDate.SetFocus
End If

End Sub

 
Private Sub CmbCRID_GotFocus()
If CmbLocationID.Text = "" Then
   MsgBox "Please Select Location First"
   CmbLocationID.SetFocus
   Exit Sub
End If

tmp1 = CmbCRID.Text

tmp = "Select distinct TCR.CRID,TCR.SM_Prefix from Txn_NewCashReceipt TCR"
tmp = tmp & " Inner join Txn_NewCashReceiptClientDetail TCRD on TCR.CRID=TCRD.NewCRID and TCR.SM_Prefix=TCRD.SM_Prefix"
tmp = tmp & " Where TCR.AdvanceFlag = 1 And TCR.RejectFlag <> 'Y' And TCR.ReturnFlag<>'Y' and TCR.SM_Prefix = '" & TxtPreFix.Text & "' And TCR.Flag = 'P' "  'And TCR.BatchID > 0 And TCRD.EntryNo > 0
tmp = tmp & " and (TCR.InstrumentAmount-(Select isnull(sum(RecAmount),0) from Txn_NewCashReceiptTIDetail Where NewCRID= TCR.NewCRID and SM_Prefix='" & TxtPreFix & "'))>0"

Call Tmp3Table

CmbNewCRID.Clear

If TmpRs3.RecordCount = 0 Then
   MsgBox "No record found !!!! "
   Exit Sub
End If

For I = 1 To TmpRs3.RecordCount
    CmbCRID.AddItem TmpRs3!NewCRID
    TmpRs3.MoveNext
Next

CmbCRID.Text = tmp1

End Sub

Private Sub CmbCRID_LostFocus()
If CmbCRID.Text = "" Then
   Call ClearHeader
   Call Grid_Head1
   Call Grid_Head2
   Call ClearTITabControl

   Exit Sub
End If
TmpRs3.MoveFirst
TmpRs3.Find "NewCRID = " & Val(CmbCRID.Text) & ""
If TmpRs3.EOF Then
   MsgBox "Invalid selection "
   CmbCRID.SetFocus
   Exit Sub
End If

mCRID = TmpRs3!NewCRID
Call GetData
Call GetClientDetails
'Call AddTIDetails
Call GetAllocatedAmt

End Sub

Private Sub CmdAdd2List_Click()
If TxtTINo = "" Then Exit Sub
If TxtRecAmt = "" Then
   MsgBox "Blank 'Received Amount' not allowed!!"
   TxtRecAmt.SetFocus
   Exit Sub
End If

''If TxtTDSAmt = "" Then
''   MsgBox "Blank 'TDS Deducted' not allowed!!"
''   TxtTDSAmt.SetFocus
''   Exit Sub
''End If

r = MSHFlexGrid3.RowSel

MSHFlexGrid3.TextMatrix(0, 8) = Val(MSHFlexGrid3.TextMatrix(0, 8)) - Val(MSHFlexGrid3.TextMatrix(r, 8))
MSHFlexGrid3.TextMatrix(0, 7) = Val(MSHFlexGrid3.TextMatrix(0, 7)) - Val(MSHFlexGrid3.TextMatrix(r, 7))
MSHFlexGrid3.TextMatrix(0, 9) = Val(MSHFlexGrid3.TextMatrix(0, 9)) - Val(MSHFlexGrid3.TextMatrix(r, 9))

MSHFlexGrid3.TextMatrix(r, 7) = TxtRecAmt.Text
MSHFlexGrid3.TextMatrix(r, 8) = TxtTDSAmt.Text
MSHFlexGrid3.TextMatrix(r, 9) = TxtOtherDed.Text
MSHFlexGrid3.TextMatrix(r, 10) = TxtBalRec.Text

MSHFlexGrid3.TextMatrix(0, 8) = Val(MSHFlexGrid3.TextMatrix(0, 8)) + Val(MSHFlexGrid3.TextMatrix(r, 8))
MSHFlexGrid3.TextMatrix(0, 7) = Val(MSHFlexGrid3.TextMatrix(0, 7)) + Val(MSHFlexGrid3.TextMatrix(r, 7))
MSHFlexGrid3.TextMatrix(0, 9) = Val(MSHFlexGrid3.TextMatrix(0, 9)) + Val(MSHFlexGrid3.TextMatrix(r, 9))
Call ClearTITabControl
End Sub

Private Sub AddTIDetails(mNAVCode_ As Variant)

''tmp = " Select TINo,TIDate,TIAmount,TIID,TI.ClientIDRow,ClientID,ClientName,TI.DepositoryID,DepositoryName,TI.PostInvoiceNo,TI.EntryNo,TI.NAVCode From  Txn_TIDetails TI "
''tmp = tmp & " Left Join Mst_Client MC on TI.ClientIDRow=MC.ClientIDRow"
''tmp = tmp & " Left Join Mst_Depository MD on TI.DepositoryID=MD.DepositoryID"
''tmp = tmp & " Where  TINo <> '' And TIAmount-((select isnull(Sum(RecAmount)+Sum(RecTDS)+Sum(RecOtherDed),0) From Txn_NewCashReceiptTIDetail TCRT"
''tmp = tmp & " Inner Join Txn_NewCashReceipt TCR on TCR.CRID= TCRT.CRID and TCR.SM_Prefix=TCRT.SM_Prefix"
''tmp = tmp & " Where TCRT.TIID=TI.TIID And TCR.ReturnFlag<>'Y'))>0 "
''
''
''If LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "admin" And LCase(Gen_UserRole) <> "superadmin" And LCase(Gen_UserRole) <> "fag" And LCase(Gen_UserRole) <> "head cmg" Then
''   tmp = tmp & " And (LocationID=" & mLocationID & " or LocationID is null)"
''End If
''
''tmp = tmp & " And (TI.NAVCode in "
''tmp = tmp & " (Select distinct NAVCode from Txn_NewCashReceiptClientDetail TCRD"
''tmp = tmp & " inner join Mst_Client MC on MC.ClientIDRow=TCRD.ClientIDRow"
''tmp = tmp & " Where TCRD.CRID =  " & Val(mCRID) & " and TCRD.SM_Prefix= '" & TxtPreFix.Text & "') OR"
''tmp = tmp & " TI.NAVCode in "
''tmp = tmp & " (Select distinct RNAVCode from Txn_NewCashReceiptClientDetail TCRD"
''tmp = tmp & " inner join Mst_Client MC on MC.ClientIDRow=TCRD.ClientIDRow"
''tmp = tmp & " Where TCRD.CRID =  " & Val(mCRID) & " and TCRD.SM_Prefix= '" & TxtPreFix.Text & "') OR"
''tmp = tmp & " TI.NAVCode in "
''tmp = tmp & " (Select distinct NAVCode from Txn_NewCashReceiptClientDetail TCRD"
''tmp = tmp & " inner join Mst_Depository MC on MC.DepositoryID=TCRD.DPID"
''tmp = tmp & " Where TCRD.CRID =  " & Val(mCRID) & " and TCRD.SM_Prefix= '" & TxtPreFix.Text & "'))"
''
''Call Tmp3Table
''
''tmp = " Select TCRT.TIID,Sum(TCRT.RecAmount) as RecAmount,Sum(RecTDS) as RecTDSAmount,Sum(RecOtherDed)as RecOtherDed  From Txn_NewCashReceiptTIDetail TCRT"
''tmp = tmp & " Inner Join Txn_NewCashReceipt TCR on TCR.CRID= TCRT.CRID and TCR.SM_Prefix=TCRT.SM_Prefix"
''tmp = tmp & " Where TCR.ReturnFlag<>'Y'"
''tmp = tmp & " Group By  TCRT.TIID"
''
''Call Tmp4Table


Wc1 = ""

If LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "admin" And LCase(Gen_UserRole) <> "superadmin" And LCase(Gen_UserRole) <> "fag" And LCase(Gen_UserRole) <> "fagadmin" And LCase(Gen_UserRole) <> "headcmg" Then
   Wc1 = Wc1 & " And (TI.LocationID=" & mLocationID & " or TI.LocationID is null)"
End If


Wc1 = Wc1 & " And (TI.NAVCode='" & mNAVCode_ & "' Or TI.NAVCode = '" & mGSTNAVCode & "') "

'wc1 = wc1 & " And (TI.NAVCode in "
'wc1 = wc1 & " (Select distinct NAVCode from Txn_NewCashReceiptClientDetail TCRD"
'wc1 = wc1 & " inner join Mst_Client MC on MC.ClientIDRow=TCRD.ClientIDRow"
'wc1 = wc1 & " Where TCRD.CRID =  " & Val(mCRID) & " and TCRD.SM_Prefix= '" & TxtPreFix.Text & "') OR"
'wc1 = wc1 & " TI.NAVCode in "
'wc1 = wc1 & " (Select distinct RNAVCode from Txn_NewCashReceiptClientDetail TCRD"
'wc1 = wc1 & " inner join Mst_Client MC on MC.ClientIDRow=TCRD.ClientIDRow"
'wc1 = wc1 & " Where TCRD.CRID =  " & Val(mCRID) & " and TCRD.SM_Prefix= '" & TxtPreFix.Text & "') OR"
'wc1 = wc1 & " TI.NAVCode in "
'wc1 = wc1 & " (Select distinct NAVCode from Txn_NewCashReceiptClientDetail TCRD"
'wc1 = wc1 & " inner join Mst_Depository MC on MC.DepositoryID=TCRD.DPID"
'wc1 = wc1 & " Where TCRD.CRID =  " & Val(mCRID) & " and TCRD.SM_Prefix= '" & TxtPreFix.Text & "'))"

Call GetClientsRemInvoices(Wc1)

'Call Grid_Head2

If TmpRs3.RecordCount > 0 Then
   For gr = 1 To TmpRs.RecordCount
       If CheckDuplicateTINo(TmpRs!EntryNo) = False Then
           rn = MSHFlexGrid3.Rows - 1
    
           MSHFlexGrid3.TextMatrix(rn, 0) = TmpRs!TINo
           MSHFlexGrid3.TextMatrix(rn, 1) = IIf(IsNull(TmpRs!TIDate), "", Format(TmpRs!TIDate, "dd-MMM-yyyy"))
           MSHFlexGrid3.TextMatrix(rn, 2) = IIf(IsNull(TmpRs!TIAmount), 0, TmpRs!TIAmount)
           MSHFlexGrid3.TextMatrix(rn, 3) = IIf(IsNull(TmpRs!RecAmount), 0, TmpRs!RecAmount)
           MSHFlexGrid3.TextMatrix(rn, 3) = Val(MSHFlexGrid3.TextMatrix(rn, 3)) + IIf(IsNull(TmpRs!AdjAmountFAG), 0, TmpRs!AdjAmountFAG)
           MSHFlexGrid3.TextMatrix(rn, 4) = IIf(IsNull(TmpRs!RecTDS), 0, TmpRs!RecTDS)
           MSHFlexGrid3.TextMatrix(rn, 5) = IIf(IsNull(TmpRs!RecOtherDed), 0, TmpRs!RecOtherDed)
                
           MSHFlexGrid3.TextMatrix(rn, 6) = Format(Val(MSHFlexGrid3.TextMatrix(rn, 2)) - (Val(MSHFlexGrid3.TextMatrix(rn, 3)) + Val(MSHFlexGrid3.TextMatrix(rn, 4)) + Val(MSHFlexGrid3.TextMatrix(rn, 5))), "#####0.00")
           
           MSHFlexGrid3.TextMatrix(rn, 11) = TmpRs!TIID
           
           tmp = "Select ShowDP,DPID,TCRD.ClientIDRow,ClientID,ClientName,DepositoryName,Amount"
           tmp = tmp & " From Txn_NewCashReceiptClientDetail TCRD"
           tmp = tmp & " left join Mst_Client MC on MC.ClientIDRow=TCRD.ClientIDRow"
           tmp = tmp & " Left join Mst_Depository MD on MD.DepositoryID=TCRD.DPID"
           tmp = tmp & " Where  TCRD.CRID = " & mCRID & " and TCRD.SM_Prefix= '" & TxtPreFix & "' And ( MC.NAVCode='" & TmpRs!NavCode & "' or MC.RNAVCode='" & TmpRs!NavCode & "' or MD.NAVCode='" & TmpRs!NavCode & "')"
           
           Call Tmp6Table
           
           If TmpRs6.RecordCount > 0 Then
              If TmpRs6!ShowDP = 0 Then
                 MSHFlexGrid3.TextMatrix(rn, 12) = TmpRs6!ClientIDRow
                 MSHFlexGrid3.TextMatrix(rn, 13) = TmpRs6!ClientID
                 MSHFlexGrid3.TextMatrix(rn, 14) = TmpRs6!ClientName
                 MSHFlexGrid3.TextMatrix(rn, 17) = "N"
               Else
                 MSHFlexGrid3.TextMatrix(rn, 12) = TmpRs6!DPID
                 MSHFlexGrid3.TextMatrix(rn, 13) = TmpRs6!DPID
                 MSHFlexGrid3.TextMatrix(rn, 14) = TmpRs6!DepositoryName
                 MSHFlexGrid3.TextMatrix(rn, 17) = "Y"
               End If
           End If
           
           MSHFlexGrid3.TextMatrix(rn, 15) = IIf(IsNull(TmpRs!PostInvoiceNo), 0, TmpRs!PostInvoiceNo)
           MSHFlexGrid3.TextMatrix(rn, 16) = IIf(IsNull(TmpRs!EntryNo), 0, TmpRs!EntryNo)
           MSHFlexGrid3.TextMatrix(rn, 18) = IIf(IsNull(TmpRs!NavCode), "", TmpRs!NavCode)
           MSHFlexGrid3.TextMatrix(rn, 19) = IIf(IsNull(TmpRs!Locationname), "", TmpRs!Locationname)
           MSHFlexGrid3.Rows = MSHFlexGrid3.Rows + 1
       End If
       TmpRs.MoveNext
   Next
End If
End Sub
Private Function CheckDuplicateTINo(EntryNo_ As Double) As Boolean
Dim Retval As Boolean
Retval = False
For I = 2 To MSHFlexGrid3.Rows - 1
    If MSHFlexGrid3.TextMatrix(I, 0) <> "" Then
       If Val(MSHFlexGrid3.TextMatrix(I, 16)) = EntryNo_ Then
          Retval = True
          Exit For
       End If
    End If
Next
CheckDuplicateTINo = Retval
End Function
Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Txn_CRAdvanceAllocation.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Txn_CRAdvanceAllocation.xls")
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


Private Sub CmdLocationID_Click()
FrmMst_Location.Show
End Sub





Private Sub CmdSearch_Click()

If SFrom.Value > STo.Value Then
   MsgBox "'From Date' must be less than 'To date'!!!"
   SFrom.SetFocus
   Exit Sub
End If

tmp = "Select CRA.CRAAID,CRA.AllocDate,CR.CRID,LM.LocationName,CR.CRDate,CR.InstrumentType,CR.DrawnOn,CR.BranchName,CR.InstrumentNo,CR.InstrumentDate,"
tmp = tmp & " CR.InstrumentAmount ,CR.ReceivedFrom, CR.Remark, CR.Flag, CR.HologramNo,CRA.G_UID,CR.DepositDate,CR.FAGPostingDate,CR.FAGRemarks 'DepositRemarks',MFB.BankName +'~'+ MFB.AccountNo as 'Bank',CR.AdvanceFlag,CR.ReturnFlag,CRA.BatchID,CRA.CreatedBy,CRA.CreatedDate,CRA.UpdatedBy,CRA.UpdatedDate"
tmp = tmp & " From Txn_CRAdvanceAllocation CRA"
tmp = tmp & " Inner Join Txn_NewCashReceipt CR on CR.CRID=CRA.CRID and CR.SM_Prefix=CRA.SM_Prefix"
tmp = tmp & " Inner Join Mst_Location LM On CR.LocationID = LM.LocationID"
tmp = tmp & " Left Join Mst_FAGBank MFB On CR.FAGBankID = MFB.FAGBankID "

If Gen_DBType = "MDB" Then
   tmp = tmp & " Where CRA.AllocDate Between #" & Format(SFrom, Gen_DatFormat) & "# And #" & Format(STo, Gen_DatFormat) & "#) "
Else
    tmp = tmp & " where CRA.AllocDate Between '" & Format(SFrom, Gen_DatFormat) & "' And '" & Format(STo, Gen_DatFormat) & "' "
End If

If CmbSLocation.Text <> "" Then
   tmp = tmp & " AND LM.LocationName = '" & CmbSLocation.Text & "' "
Else
   If Gen_WcLocation <> "" Then
      tmp = tmp & " AND LM." & Trim(Gen_WcLocation)
   End If
End If
If TxtSCRID <> "" Then
   tmp = tmp & " And CR.CRID = " & Val(TxtSCRID)
End If

Call TmpTable
Call Grid_Head
For iu = 1 To TmpRs.RecordCount
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
    For C = 0 To MSHFlexGrid1.Cols - 1
       MSHFlexGrid1.TextMatrix(iu, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
    Next
    MSHFlexGrid1.TextMatrix(iu, 1) = Format(MSHFlexGrid1.TextMatrix(iu, 1), "dd-mmm-yyyy")
    MSHFlexGrid1.TextMatrix(iu, 4) = Format(MSHFlexGrid1.TextMatrix(iu, 4), "dd-mmm-yyyy")
    
    MSHFlexGrid1.TextMatrix(iu, 24) = Format(MSHFlexGrid1.TextMatrix(iu, 24), "dd-mmm-yyyy hh:mm:ss")
    MSHFlexGrid1.TextMatrix(iu, 26) = Format(MSHFlexGrid1.TextMatrix(iu, 26), "dd-mmm-yyyy hh:mm:ss")
    
    If MSHFlexGrid1.TextMatrix(iu, 9) <> "" Then
       MSHFlexGrid1.TextMatrix(iu, 9) = Format(MSHFlexGrid1.TextMatrix(iu, 9), "dd-mmm-yyyy")
    End If
    If MSHFlexGrid1.TextMatrix(iu, 17) <> "" Then
       MSHFlexGrid1.TextMatrix(iu, 17) = Format(MSHFlexGrid1.TextMatrix(iu, 17), "dd-mmm-yyyy")
    End If
    If MSHFlexGrid1.TextMatrix(iu, 20) = "1" Then
       MSHFlexGrid1.TextMatrix(iu, 20) = "Yes"
    Else
       MSHFlexGrid1.TextMatrix(iu, 20) = "No"
    End If
    If MSHFlexGrid1.TextMatrix(iu, 21) = "Y" Then
       MSHFlexGrid1.TextMatrix(iu, 21) = "Yes"
    Else
       MSHFlexGrid1.TextMatrix(iu, 21) = "No"
    End If
    
    TmpRs.MoveNext
Next

End Sub



Private Sub DeleteCmd_Click()
If RsTxn_CRAdvanceAllocation!BatchID > 0 Then
   MsgBox "Batch ID is generated already. You can't delete this receipt!!!! "
   Exit Sub
End If

Dim ans As Variant
On Error GoTo msgError

ans = MsgBox("Do you really want to DELETE this Record??", vbYesNo)
If ans = vbYes Then
   
   Call DeleteTIDetail
   
   RsTxn_CRAdvanceAllocation.Delete
   RsTxn_CRAdvanceAllocation.Update
   RsTxn_CRAdvanceAllocation.Requery
   If RsTxn_CRAdvanceAllocation.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsTxn_CRAdvanceAllocation.MoveNext
   If RsTxn_CRAdvanceAllocation.EOF() Then
      RsTxn_CRAdvanceAllocation.MoveLast
   End If
   Call Indata
End If
Exit Sub
msgError:
MsgBox "Child record Present "
RsTxn_CRAdvanceAllocation.CancelUpdate

End Sub

Private Sub EditCmd_Click()

If RsTxn_CRAdvanceAllocation!BatchID <> 0 Then
   MsgBox "Batch is generated already. You can't edit this Record!!!! "
   Exit Sub
End If
   
Edit_Flg = "E"

Call GButtonLock(Me, False)
CmbLocationID.Text = TxtLocationID
CmbExchangeTypeID = TxtExchangeTypeID
CmbCRID = TxtCRID

Call AppendTIDetails
Frame2.Enabled = True
End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Call GButtonLockAD(Me)
   Frame1.Enabled = False
   SSTab1.Enabled = False
   CmdPrint.Enabled = False
   CmdPrintPreview.Enabled = False
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If
Frame1.Enabled = True
SSTab1.Enabled = True
Call Grid_Head1
Call Grid_Head2
wc = ""
If Gen_WcLocation <> "" Then
   wc = "Where " & Gen_WcLocation
End If

Call Grid_Head
SFrom.Value = Date - 730
STo.Value = Date
CmbSLocation.Text = ""
TxtSCRID = ""

Call TableTxn_CRAdvanceAllocation(" " & wc & "  Order by CRAAID")
If RsTxn_CRAdvanceAllocation.RecordCount = 0 Then
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
   Call TableMst_Location("")
Else
   Call TableMst_Location("where " & Gen_WcLocation)
End If

CmbLocationID.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbLocationID.AddItem RsMst_Location!Locationname
    RsMst_Location.MoveNext
Next
CmbLocationID.Text = tmp
End Sub
'----
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

If tmp <> CmbLocationID.Text Then
   CmbCRID.Text = ""
   Call ClearHeader
   Call Grid_Head1
   Call Grid_Head2
   Call ClearTITabControl
End If

mLocationID = RsMst_Location!LocationID
TxtPreFix = GetPreFixByLocation(RsMst_Location!LocationID, "P")
End Sub

Private Sub MSHFlexGrid1_Click()
Dim I As Variant
I = MSHFlexGrid1.RowSel
If I <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(I, 0) = "" Then Exit Sub
RsTxn_CRAdvanceAllocation.MoveFirst
RsTxn_CRAdvanceAllocation.Find "CRAAID= " & Val(MSHFlexGrid1.TextMatrix(I, 0)) & ""
If Not RsTxn_CRAdvanceAllocation.EOF Then
   Call Indata
   Frame0.Visible = False
End If


End Sub



Private Sub MSHFlexGrid3_Click()
If SaveCmd.Enabled = False Then Exit Sub
Call ClearTITabControl
r = MSHFlexGrid3.RowSel
If MSHFlexGrid3.TextMatrix(r, 0) = "" Then Exit Sub
If r <= 1 Then Exit Sub
mRow = r
TxtTINo.Text = MSHFlexGrid3.TextMatrix(r, 0)
TxtTIDate.Text = MSHFlexGrid3.TextMatrix(r, 1)
TxtBalAmount = MSHFlexGrid3.TextMatrix(r, 6)
TxtRecAmt = MSHFlexGrid3.TextMatrix(r, 7)
TxtTDSAmt = MSHFlexGrid3.TextMatrix(r, 8)
TxtOtherDed = MSHFlexGrid3.TextMatrix(r, 9)
TxtBalRec = MSHFlexGrid3.TextMatrix(r, 8)
TxtNAVCode = MSHFlexGrid3.TextMatrix(r, 18)
TxtTILocation = MSHFlexGrid3.TextMatrix(r, 19)
TxtRecAmt.SetFocus
End Sub

Private Sub SaveCmd_Click()

If TxtAllocDate.Value > Date Then
   MsgBox "FutureDate not Allowed!!"
   TxtAllocDate.SetFocus
   Exit Sub
End If
If CmbLocationID.Text = "" Then
   MsgBox "Blank Location Name not allowed !!! "
   CmbLocationID.SetFocus
   Exit Sub
End If

If CmbCRID.Text = "" Then
   MsgBox "Blank Cash Receipt ID  not allowed !!! "
   CmbCRID.SetFocus
   Exit Sub
End If

If MSHFlexGrid2.Rows < 4 Then
   MsgBox "Blank Client Details not allowed !!! "
   SSTab1.Tab = 0
   Exit Sub
End If

If MSHFlexGrid3.Rows < 4 Then
   MsgBox "Blank TI Details not allowed !!! "
   SSTab1.Tab = 1
   Exit Sub
End If

If Edit_Flg = "A" Then
   RsTxn_CRAdvanceAllocation.AddNew
   RsTxn_CRAdvanceAllocation!CRAAID = GetCRAAID
   RsTxn_CRAdvanceAllocation!G_UID = GetGUID
   TxtCRAAID.Text = RsTxn_CRAdvanceAllocation!CRAAID
   TxtG_UID = RsTxn_CRAdvanceAllocation!G_UID
Else
   RsTxn_CRAdvanceAllocation.MoveFirst
   RsTxn_CRAdvanceAllocation.Find "CRAAID= " & TxtCRAAID & ""
End If
RsTxn_CRAdvanceAllocation!CRID = mCRID
RsTxn_CRAdvanceAllocation!SM_Prefix = TxtPreFix
RsTxn_CRAdvanceAllocation!AllocDate = TxtAllocDate.Value
RsTxn_CRAdvanceAllocation!LocationID = mLocationID
'RsTxn_CRAdvanceAllocation!Remark = TxtRemark
RsTxn_CRAdvanceAllocation!BatchID = 0
   
If IsNull(RsTxn_CRAdvanceAllocation!CreatedBy) = True Or RsTxn_CRAdvanceAllocation!CreatedBy = "" Then
   RsTxn_CRAdvanceAllocation!CreatedBy = Gen_UserLoginID
   RsTxn_CRAdvanceAllocation!CreatedDate = Now
Else
   RsTxn_CRAdvanceAllocation!UpdatedBy = Gen_UserLoginID
   RsTxn_CRAdvanceAllocation!UpdatedDate = Now
End If

RsTxn_CRAdvanceAllocation.Update

Call SaveTIDetail

RsTxn_CRAdvanceAllocation.Requery
RsTxn_CRAdvanceAllocation.MoveFirst
RsTxn_CRAdvanceAllocation.Find "CRAAID= " & TxtCRAAID.Text & ""
If RsTxn_CRAdvanceAllocation.EOF Then
   RsTxn_CRAdvanceAllocation.MoveLast
End If

Call Indata

End Sub

Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   If RsTxn_CRAdvanceAllocation.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   RsTxn_CRAdvanceAllocation.MoveFirst
   RsTxn_CRAdvanceAllocation.Find "CRAAID= " & Val(TxtCRAAID) & ""
   If RsTxn_CRAdvanceAllocation.EOF Then
      RsTxn_CRAdvanceAllocation.MoveLast
   End If
   Call Indata
   Exit Sub
End If
Frame0.Height = Frame4.Height - 200
Frame0.Left = Frame1.Left - 100
Frame0.Top = Frame1.Top - 100
Frame0.Width = Frame4.Width - 150
Frame0.Visible = True
MSHFlexGrid1.Width = Frame0.Width - 400
MSHFlexGrid1.Height = Frame0.Height - 1700
RsTxn_CRAdvanceAllocation.Filter = ""
RsTxn_CRAdvanceAllocation.MoveFirst

''===========comment on 7 july 2012 for not clearing search grid=============
'MSHFlexGrid1.Clear
'MSHFlexGrid1.Cols = 11
'Call Grid_Head
'SFrom.Value = Date - 730
'STo.Value = Date
'CmbSLocation.Text = ""
'TxtSCRID = ""
''===========comment on 7 july 2012 for not clearing search grid=============

End Sub

Sub ClearTITabControl()
TxtTINo = ""
TxtTIDate = ""
TxtBalAmount = ""
TxtRecAmt = ""
TxtTDSAmt = ""
TxtBalRec = ""
TxtOtherDed = ""
TxtNAVCode = ""
TxtTILocation = ""
mRow = 0
End Sub

Private Sub TxtAllocDate_LostFocus()

If Format(Gen_BlockingDate, "yyyy-mm") >= Format(TxtAllocDate.Value, "yyyy-mm") Then
   MsgBox "Cash Receipt Allocation Date allowed for next month of " & Format(Gen_BlockingDate, "MMM-yyyy") & "!!! "
   TxtAllocDate.SetFocus
   Exit Sub
End If

End Sub

Private Sub TxtRecAmt_LostFocus()
If TxtRecAmt.Text = "" Then Exit Sub
If IsNumeric(TxtRecAmt.Text) = False Or Val(TxtRecAmt.Text) < 0 Then
   MsgBox "Invalid Numeric Format !!!"
   TxtRecAmt.SetFocus
   Exit Sub
End If
'TxtRecAmt = Round(Val(TxtRecAmt), 2)
If Val(TxtRecAmt) > Val(TxtBalAmount) Then
   MsgBox "Received Amount should be less than or equal to Balance Amount!"
   TxtRecAmt.SetFocus
   Exit Sub
End If
If Val(TxtRecAmt) > Val(TxtUnAllocAmt) Then
   MsgBox "Received Amount should be less than or equal to Unallocated Amount!"
   TxtRecAmt.SetFocus
   Exit Sub
End If
If Val(TxtRecAmt.Text) > GetRemClientAdvAmt(TxtNAVCode, mRow) Then
   MsgBox "Received Amount should be less than or equal to Clients Adv. Amount!"
   TxtRecAmt.SetFocus
   Exit Sub
End If
'TxtBalRec.Text = Format(Val(TxtBalAmount) - (Val(TxtRecAmt) + Val(TxtTDSAmt)), "#####0.00")
TxtBalRec.Text = Format(Val(TxtBalAmount) - (Val(TxtRecAmt) + Val(TxtTDSAmt) + Val(TxtOtherDed)), "#####0.00")
If Val(TxtBalRec.Text) < 0 Then
   MsgBox "Balance Receivable cannot be less than zero. "
   TxtRecAmt.SetFocus
End If
End Sub

Function getCRTIDetailMaxID() As Double
Dim Retval As Double
Retval = 0
tmp = "select max(CRTIDID) from Txn_NewCashReceiptTIDetail "
Call TmpTable
If TmpRs.RecordCount > 0 Then
   Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
End If
Retval = Retval + 1
TmpRs.Close
getCRTIDetailMaxID = Retval

End Function

Public Sub Indata()

Edit_Flg = ""
Frame1.Enabled = False

Frame2.Enabled = False
Call ClearTITabControl
CmdReturn.Enabled = False
CmbLocationID.Visible = False
CmdLocationID.Visible = False
TxtLocationID.Visible = True
CmbLocationID.Text = ""

CmbCRID.Visible = False
TxtCRID.Visible = True
CmbCRID.Text = ""

TxtCRAAID = RsTxn_CRAdvanceAllocation!CRAAID
TxtCRID = RsTxn_CRAdvanceAllocation!CRID
mCRID = RsTxn_CRAdvanceAllocation!CRID
TxtPreFix = RsTxn_CRAdvanceAllocation!SM_Prefix
mLocationID = IIf(IsNull(RsTxn_CRAdvanceAllocation!LocationID), 0, RsTxn_CRAdvanceAllocation!LocationID)
TxtLocationID = GetLocationName(RsTxn_CRAdvanceAllocation!LocationID)
TxtAllocDate = RsTxn_CRAdvanceAllocation!AllocDate
Call GetData
Call GetClientDetails

TxtCreated = IIf(IsNull(RsTxn_CRAdvanceAllocation!CreatedBy), "", RsTxn_CRAdvanceAllocation!CreatedBy) & " :- " & IIf(IsNull(RsTxn_CRAdvanceAllocation!CreatedDate), "", Format(RsTxn_CRAdvanceAllocation!CreatedDate, "dd-mmm-yyyy hh:mm:ss"))
TxtUpdated = IIf(IsNull(RsTxn_CRAdvanceAllocation!UpdatedBy), "", RsTxn_CRAdvanceAllocation!UpdatedBy) & " :- " & IIf(IsNull(RsTxn_CRAdvanceAllocation!UpdatedDate), "", Format(RsTxn_CRAdvanceAllocation!UpdatedDate, "dd-mmm-yyyy hh:mm:ss"))

Call GetClientDetails
Call InDataTIDetails
Call GetAllocatedAmt

If RsTxn_CRAdvanceAllocation.RecordCount = 1 Then
   Call GButtonLock_1(Me, True)
Else
   Call GButtonLock(Me, True)
End If
CmdPrint.Enabled = True


End Sub

Private Sub GetClientDetails()
Dim ans As Variant
Call Grid_Head1
Call Grid_Head2
'Call TableTxn_NewCashReceiptClientDetail(" Where CRID = " & Val(mCRID) & " And SM_Prefix='" & TxtPreFix & "'")

'''tmp = "Select TCRD.ExchangeTypeID,ExchangeType,ShowDP,TCRD.ClientIDRow,MC.ClientID,MC.ClientName,DP1.DepositoryID,DP1.DepositoryName as CDepositoryName,TCRD.DPID,DP.DepositoryName,DP.NAVCode as DPNAVCode,MC.NAVCode,MC.RNAVCode,Amount,Remark"
'''tmp = tmp & " From Txn_NewCashReceiptClientDetail TCRD"
'''tmp = tmp & " Inner join Mst_ExchangeType ME on ME.ExchangeTypeID=TCRD.ExchangeTypeID"
'''tmp = tmp & " left join Mst_Client MC on MC.ClientIDRow=TCRD.ClientIDRow"
'''tmp = tmp & " Left join Mst_Depository DP1 on DP1.DepositoryID=MC.DepositoryID"
'''tmp = tmp & " Left join Mst_Depository DP on DP.DepositoryID=TCRD.DPID"
'''tmp = tmp & "  Where TCRD.CRID = " & Val(mCRID) & " And TCRD.Srtu8M_Prefix='" & TxtPreFix & "'"

tmp = " Select TCRD.ExchangeTypeID,ExchangeType,ShowDP,TCRD.ClientID,MC.ClientIDRow,MC.ClientName,DP1.DepositoryID,DP1.DepositoryName as CDepositoryName,TCRD.DPID,DP.DepositoryName,isnull(TCRD.AdvNAVCode,DP.NAVCode) as DPNAVCode,isnull(TCRD.AdvNAVCode,MC.NAVCode)as NAVCode,MC.RNAVCode,TCRD.NAVCode 'GSTNAVCode',Amount,Remark,sum(RecAmount) RecAmt"
tmp = tmp & " From Txn_NewCashReceiptClientDetail TCRD"
tmp = tmp & " Inner join Mst_ExchangeType ME on ME.ExchangeTypeID=TCRD.ExchangeTypeID"
tmp = tmp & " left join Txn_NewCashReceiptTIDetail TCRT on TCRT.NewCRID=TCRD.NewCRID and TCRT.SM_Prefix=TCRD.SM_Prefix And (TCRD.ClientID=TCRT.ClientID OR TCRD.DPID=TCRT.DPID) And TCRT.CRAAID<>" & Val(TxtCRAAID) & ""
tmp = tmp & " left join Mst_Client MC on MC.ClientID=TCRD.ClientID"
tmp = tmp & " Left join Mst_Depository DP1 on DP1.DepositoryID=MC.DepositoryID"
tmp = tmp & " Left join Mst_Depository DP on DP.DepositoryID=TCRD.DPID"
tmp = tmp & " Where TCRD.NewCRID=" & Val(mCRID) & "  And TCRD.SM_Prefix='" & TxtPreFix & "' "
tmp = tmp & " Group by TCRD.ExchangeTypeID,TCRD.NAVCode,ExchangeType,ShowDP,TCRD.ClientID,MC.ClientIDRow,MC.ClientName,DP1.DepositoryID,DP1.DepositoryName ,TCRD.DPID,DP.DepositoryName,DP.NAVCode,MC.NAVCode,MC.RNAVCode,Amount,Remark,TCRD.AdvNAVCode"

Call Tmp2Table

MSHFlexGrid2.Rows = TmpRs2.RecordCount + 3
For I = 2 To MSHFlexGrid2.Rows - 2
    MSHFlexGrid2.TextMatrix(I, 0) = TmpRs2!ExchangeTypeID
    MSHFlexGrid2.TextMatrix(I, 1) = TmpRs2!ExchangeType
    MSHFlexGrid2.TextMatrix(I, 9) = IIf(IsNull(TmpRs2!ShowDP), 0, TmpRs2!ShowDP)
    mGSTNAVCode = IIf(IsNull(TmpRs2!GSTNavCode), 0, TmpRs2!GSTNavCode)
    MSHFlexGrid2.TextMatrix(I, 6) = IIf(IsNull(TmpRs2!Amount), "0", TmpRs2!Amount)
    If MSHFlexGrid2.TextMatrix(I, 9) = 0 Then
       MSHFlexGrid2.TextMatrix(I, 2) = TmpRs2!ClientIDRow
       MSHFlexGrid2.TextMatrix(I, 3) = TmpRs2!ClientName
       MSHFlexGrid2.TextMatrix(I, 4) = TmpRs2!DepositoryID
       MSHFlexGrid2.TextMatrix(I, 5) = TmpRs2!CDepositoryName
       MSHFlexGrid2.TextMatrix(I, 8) = TmpRs2!ClientID
       MSHFlexGrid2.TextMatrix(I, 10) = IIf(IsNull(TmpRs2!NavCode), "", TmpRs2!NavCode)
       MSHFlexGrid2.TextMatrix(I, 11) = IIf(IsNull(TmpRs2!RNAVCode), "", TmpRs2!RNAVCode)
       
       If Edit_Flg = "A" Then
          If Val(MSHFlexGrid2.TextMatrix(I, 6)) > 0 Then
             If MSHFlexGrid2.TextMatrix(I, 10) <> "" Then
                Call AddTIDetails(MSHFlexGrid2.TextMatrix(I, 10))
             ElseIf MSHFlexGrid2.TextMatrix(I, 11) <> "" Then
                Call AddTIDetails(MSHFlexGrid2.TextMatrix(I, 11))
             End If
          End If
       End If
    Else
       MSHFlexGrid2.TextMatrix(I, 4) = TmpRs2!DPID
       MSHFlexGrid2.TextMatrix(I, 5) = TmpRs2!DepositoryName
       MSHFlexGrid2.TextMatrix(I, 8) = TmpRs2!DPID
       MSHFlexGrid2.TextMatrix(I, 10) = IIf(IsNull(TmpRs2!DPNAVCode), "", TmpRs2!DPNAVCode)
       If Edit_Flg = "A" Then
          If Val(MSHFlexGrid2.TextMatrix(I, 6)) > 0 Then
             If MSHFlexGrid2.TextMatrix(I, 10) <> "" Then
                Call AddTIDetails(MSHFlexGrid2.TextMatrix(I, 10))
             End If
          End If
       End If
    End If
    
    MSHFlexGrid2.TextMatrix(I, 12) = IIf(IsNull(TmpRs2!RecAmt), "0", TmpRs2!RecAmt)
    MSHFlexGrid2.TextMatrix(I, 13) = Format(Val(MSHFlexGrid2.TextMatrix(I, 6)) - Val(MSHFlexGrid2.TextMatrix(I, 12)), "######0.00")
    MSHFlexGrid2.TextMatrix(I, 7) = TmpRs2!Remark
    MSHFlexGrid2.TextMatrix(0, 6) = Val(MSHFlexGrid2.TextMatrix(0, 6)) + Val(MSHFlexGrid2.TextMatrix(I, 6))
    MSHFlexGrid2.TextMatrix(0, 12) = Val(MSHFlexGrid2.TextMatrix(0, 12)) + Val(MSHFlexGrid2.TextMatrix(I, 12))
    MSHFlexGrid2.TextMatrix(0, 13) = Val(MSHFlexGrid2.TextMatrix(0, 13)) + Val(MSHFlexGrid2.TextMatrix(I, 13))
    TmpRs2.MoveNext
Next

End Sub

Private Sub InDataTIDetails()
Dim ans As Variant
Call Grid_Head2

''tmp = " Select TTI.TINo,TTI.TIDate,TTI.TIAmount,"
''tmp = tmp & " (select Sum(TCRT.RecAmount) From Txn_NewCashReceiptTIDetail TCRT"
''tmp = tmp & " Inner Join Txn_NewCashReceipt TCR on TCR.CRID= TCRT.CRID and TCR.SM_Prefix=TCRT.SM_Prefix"
''tmp = tmp & " Where TCRT.TIID=TI.TIID And CRTIDID<>TI.CRTIDID And TCR.ReturnFlag<>'Y') as TotRecAmt"
''tmp = tmp & " ,(select Sum(RecTDS) From Txn_NewCashReceiptTIDetail TCRT"
''tmp = tmp & " Inner Join Txn_NewCashReceipt TCR on TCR.CRID= TCRT.CRID and TCR.SM_Prefix=TCRT.SM_Prefix"
''tmp = tmp & " Where TCRT.TIID=TI.TIID And CRTIDID<>TI.CRTIDID And TCR.ReturnFlag<>'Y') as TotTDSAmt"
''tmp = tmp & " ,(select Sum(RecOtherDed) From Txn_NewCashReceiptTIDetail TCRT"
''tmp = tmp & " Inner Join Txn_NewCashReceipt TCR on TCR.CRID= TCRT.CRID and TCR.SM_Prefix=TCRT.SM_Prefix"
''tmp = tmp & " Where TIID=TI.TIID And CRTIDID<>TI.CRTIDID And TCR.ReturnFlag<>'Y') as TotOtherDed"
''tmp = tmp & " ,TI.RecAmount,TI.RecTDS,TI.RecOtherDed,TI.TIID,TTI.ClientIDRow,ClientID,ClientName,TTI.DepositoryID,DepositoryName,TTI.PostInvoiceNo,TTI.EntryNo,TTI.NAVCode"
''tmp = tmp & " From Txn_NewCashReceiptTIDetail TI"
''tmp = tmp & " inner Join  Txn_TIDetails TTI on TI.TIID=TTI.TIID"
''tmp = tmp & " Left Join Mst_Client MC on TTI.ClientIDRow=MC.ClientIDRow"
''tmp = tmp & " Left Join Mst_Depository MD on TTI.DepositoryID=MD.DepositoryID"
''tmp = tmp & " Where TI.CRID = " & Val(mCRID) & " And SM_Prefix='" & TxtPreFix & "' And CRAAID=" & Val(TxtCRAAID)

tmp = " Select TTI.TINo,TTI.TIDate,TTI.TIAmount,"
tmp = tmp & " TI.RecAmount,TI.RecTDS,TI.RecOtherDed,TI.TIID,TI.ClientID,ClientIDRow,ClientName,TI.DPID,DepositoryName,TTI.PostInvoiceNo,TTI.EntryNo,TTI.NAVCode,ML.LocationName"
tmp = tmp & " From Txn_NewCashReceiptTIDetail TI"
tmp = tmp & " inner Join  Txn_TIDetails TTI on TI.TIID=TTI.TIID"
tmp = tmp & " Left Join Mst_Client MC on TI.ClientID=MC.ClientID"
tmp = tmp & " Left Join Mst_Depository MD on TI.DPID=MD.DepositoryID"
tmp = tmp & " Left Join Mst_Location ML on TTI.LocationID=ML.LocationID"
tmp = tmp & " Where TI.NewCRID = " & Val(mCRID) & " And TI.SM_Prefix='" & TxtPreFix & "' And CRAAID=" & Val(TxtCRAAID)

Call Tmp4Table

Wc1 = ""
Wc1 = " And TI.TIID in (Select TIID from Txn_NewCashReceiptTIDetail TCRT"
Wc1 = Wc1 & " Where TCRT.CRID =  " & Val(TxtCRID) & " and TCRT.SM_Prefix= '" & TxtPreFix.Text & "' And CRAAID=" & Val(TxtCRAAID) & ")"

Call GetClientsRemInvoices(Wc1, "", "", Val(TxtCRAAID), "show")

MSHFlexGrid3.Rows = TmpRs4.RecordCount + 3
For I = 2 To MSHFlexGrid3.Rows - 2
    MSHFlexGrid3.TextMatrix(I, 0) = TmpRs4!TINo
    MSHFlexGrid3.TextMatrix(I, 1) = Format(TmpRs4!TIDate, "dd-MMM-yyyy")
    MSHFlexGrid3.TextMatrix(I, 2) = TmpRs4!TIAmount
    If TmpRs.RecordCount > 0 Then TmpRs.MoveFirst
    TmpRs.Find "TIID=" & TmpRs4!TIID
    If TmpRs.EOF = False Then
       MSHFlexGrid3.TextMatrix(I, 3) = IIf(IsNull(TmpRs!RecAmount), 0, TmpRs!RecAmount)
       MSHFlexGrid3.TextMatrix(I, 4) = IIf(IsNull(TmpRs!RecTDS), 0, TmpRs!RecTDS)
       MSHFlexGrid3.TextMatrix(I, 5) = IIf(IsNull(TmpRs!RecOtherDed), 0, TmpRs!RecOtherDed)
    Else
       MSHFlexGrid3.TextMatrix(I, 3) = 0
       MSHFlexGrid3.TextMatrix(I, 4) = 0
       MSHFlexGrid3.TextMatrix(I, 5) = 0
    End If
    MSHFlexGrid3.TextMatrix(I, 6) = Format(Val(MSHFlexGrid3.TextMatrix(I, 2)) - (Val(MSHFlexGrid3.TextMatrix(I, 3)) + Val(MSHFlexGrid3.TextMatrix(I, 4)) + Val(MSHFlexGrid3.TextMatrix(I, 5))), "#####0.00")
    MSHFlexGrid3.TextMatrix(I, 7) = TmpRs4!RecAmount
    MSHFlexGrid3.TextMatrix(I, 8) = TmpRs4!RecTDS
    MSHFlexGrid3.TextMatrix(I, 9) = IIf(IsNull(TmpRs4!RecOtherDed), 0, TmpRs4!RecOtherDed)
    MSHFlexGrid3.TextMatrix(I, 10) = Format(Val(MSHFlexGrid3.TextMatrix(I, 6)) - (Val(MSHFlexGrid3.TextMatrix(I, 7)) + Val(MSHFlexGrid3.TextMatrix(I, 8)) + Val(MSHFlexGrid3.TextMatrix(I, 9))), "#####0.00")
    MSHFlexGrid3.TextMatrix(I, 11) = TmpRs4!TIID
    If IsNull(TmpRs4!ClientID) = True Then
       MSHFlexGrid3.TextMatrix(I, 12) = IIf(IsNull(TmpRs4!DPID), "", TmpRs4!DPID)
       MSHFlexGrid3.TextMatrix(I, 13) = IIf(IsNull(TmpRs4!DPID), "", TmpRs4!DPID)
       MSHFlexGrid3.TextMatrix(I, 14) = IIf(IsNull(TmpRs4!DepositoryName), "", DepositoryName)
       MSHFlexGrid3.TextMatrix(I, 17) = "Y"
    Else
       MSHFlexGrid3.TextMatrix(I, 12) = IIf(IsNull(TmpRs4!ClientID), "", TmpRs4!ClientID)
       MSHFlexGrid3.TextMatrix(I, 13) = IIf(IsNull(TmpRs4!ClientID), "", TmpRs4!ClientID)
       MSHFlexGrid3.TextMatrix(I, 14) = IIf(IsNull(TmpRs4!ClientName), "", TmpRs4!ClientName)
       MSHFlexGrid3.TextMatrix(I, 17) = "N"
    End If
    
    MSHFlexGrid3.TextMatrix(I, 15) = IIf(IsNull(TmpRs4!PostInvoiceNo), 0, TmpRs4!PostInvoiceNo)
    MSHFlexGrid3.TextMatrix(I, 16) = IIf(IsNull(TmpRs4!EntryNo), 0, TmpRs4!EntryNo)
    MSHFlexGrid3.TextMatrix(I, 18) = IIf(IsNull(TmpRs4!NavCode), "", TmpRs4!NavCode)
    MSHFlexGrid3.TextMatrix(I, 19) = IIf(IsNull(TmpRs4!Locationname), "", TmpRs4!Locationname)
    MSHFlexGrid3.TextMatrix(0, 8) = Val(MSHFlexGrid3.TextMatrix(0, 8)) + Val(MSHFlexGrid3.TextMatrix(I, 8))
    MSHFlexGrid3.TextMatrix(0, 7) = Val(MSHFlexGrid3.TextMatrix(0, 7)) + Val(MSHFlexGrid3.TextMatrix(I, 7))
    MSHFlexGrid3.TextMatrix(0, 9) = Val(MSHFlexGrid3.TextMatrix(0, 9)) + Val(MSHFlexGrid3.TextMatrix(I, 9))
    TmpRs4.MoveNext
Next

End Sub


Private Function GetCRID() As Double
Dim Retval As Double
tmp = "Select max(CRID) from Txn_NewCashReceipt Where SM_Prefix='" & TxtPreFix & "'"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetCRID = Retval
End Function
Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsTxn_CRAdvanceAllocation.MoveNext
If RsTxn_CRAdvanceAllocation.EOF Then
   RsTxn_CRAdvanceAllocation.MoveLast
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
RsTxn_CRAdvanceAllocation.MovePrevious
If RsTxn_CRAdvanceAllocation.BOF Then
   RsTxn_CRAdvanceAllocation.MoveFirst
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
RsTxn_CRAdvanceAllocation.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsTxn_CRAdvanceAllocation.MoveLast
Call Indata
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False
End Sub
Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub Grid_Head()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 27
MSHFlexGrid1.TextMatrix(0, 0) = "Allocation ID"
MSHFlexGrid1.TextMatrix(0, 1) = "Allocation Date"
MSHFlexGrid1.TextMatrix(0, 2) = "Receipt ID"
MSHFlexGrid1.TextMatrix(0, 3) = "Location"
MSHFlexGrid1.TextMatrix(0, 4) = "Receipt Date"
MSHFlexGrid1.TextMatrix(0, 5) = "Type"
MSHFlexGrid1.TextMatrix(0, 6) = "Drawn On"
MSHFlexGrid1.TextMatrix(0, 7) = "Branch Name"
MSHFlexGrid1.TextMatrix(0, 8) = "Instrument No"
MSHFlexGrid1.TextMatrix(0, 9) = "Instrument Date"
MSHFlexGrid1.TextMatrix(0, 10) = "Instrument Amount"
MSHFlexGrid1.TextMatrix(0, 11) = "Received From"
MSHFlexGrid1.TextMatrix(0, 12) = "Remark"
MSHFlexGrid1.TextMatrix(0, 13) = "Print Flag"
MSHFlexGrid1.TextMatrix(0, 14) = "Hologram No"
MSHFlexGrid1.TextMatrix(0, 15) = "G_UID"
MSHFlexGrid1.TextMatrix(0, 16) = "Deposit Date"
MSHFlexGrid1.TextMatrix(0, 17) = "FAG Posting Date"
MSHFlexGrid1.TextMatrix(0, 18) = "Deposit Remarks"
MSHFlexGrid1.TextMatrix(0, 19) = "Bank"
MSHFlexGrid1.TextMatrix(0, 20) = "Advance Flag"
MSHFlexGrid1.TextMatrix(0, 21) = "Return Flag"
MSHFlexGrid1.TextMatrix(0, 22) = "Batch ID"
MSHFlexGrid1.TextMatrix(0, 23) = "Created By"
MSHFlexGrid1.TextMatrix(0, 24) = "Created Date"
MSHFlexGrid1.TextMatrix(0, 25) = "Updated By"
MSHFlexGrid1.TextMatrix(0, 26) = "Updated Date"

MSHFlexGrid1.ColWidth(0) = 1200
MSHFlexGrid1.ColWidth(1) = 1400
MSHFlexGrid1.ColWidth(2) = 1200
MSHFlexGrid1.ColWidth(3) = 1000
MSHFlexGrid1.ColWidth(4) = 1200
MSHFlexGrid1.ColWidth(5) = 1400
MSHFlexGrid1.ColWidth(6) = 1800
MSHFlexGrid1.ColWidth(7) = 1800
MSHFlexGrid1.ColWidth(8) = 1600
MSHFlexGrid1.ColWidth(9) = 1400
MSHFlexGrid1.ColWidth(10) = 1800
MSHFlexGrid1.ColWidth(11) = 1800
MSHFlexGrid1.ColWidth(12) = 3400
MSHFlexGrid1.ColWidth(13) = 1400
MSHFlexGrid1.ColWidth(14) = 1400
MSHFlexGrid1.ColWidth(15) = 0
MSHFlexGrid1.ColWidth(16) = 1400
MSHFlexGrid1.ColWidth(17) = 1400
MSHFlexGrid1.ColWidth(18) = 1400
MSHFlexGrid1.ColWidth(19) = 5000
MSHFlexGrid1.ColWidth(20) = 1000
MSHFlexGrid1.ColWidth(21) = 1000
MSHFlexGrid1.ColWidth(22) = 1000
MSHFlexGrid1.ColWidth(23) = 1000
MSHFlexGrid1.ColWidth(24) = 1000
MSHFlexGrid1.ColWidth(25) = 1000
MSHFlexGrid1.ColWidth(26) = 1000

MSHFlexGrid1.WordWrap = True
MSHFlexGrid1.RowHeight(0) = 600

End Sub

Private Sub CmbSLocation_GotFocus()
tmp = CmbSLocation.Text
'Call TableMst_Location(" Where LocationId > 0 ")

If Gen_WcLocation = "" Then
   Call TableMst_Location("")
Else
   Call TableMst_Location("where " & Gen_WcLocation)
End If

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
End Sub

Private Sub DeleteTIDetail()
tmp = "Select TIID,RecAmount,RecTDs,RecOtherDed From Txn_NewCashReceiptTIDetail Where CRID =  " & Val(mCRID) & " and SM_Prefix= '" & TxtPreFix.Text & "' And CRAAID=" & Val(TxtCRAAID)
Call TmpTable
Call TableTxn_TIDetails(" where TIID in (Select TIID From Txn_NewCashReceiptTIDetail Where CRID =  " & Val(mCRID) & " and SM_Prefix= '" & TxtPreFix.Text & "' And CRAAID=" & Val(TxtCRAAID) & ")")
For I = 1 To TmpRs.RecordCount
    If RsTxn_TIDetails.RecordCount > 0 Then RsTxn_TIDetails.MoveFirst
    RsTxn_TIDetails.Find "TIID=" & TmpRs!TIID
    If RsTxn_TIDetails.EOF = False Then
       RsTxn_TIDetails!RecAmount = IIf(IsNull(RsTxn_TIDetails!RecAmount), 0, RsTxn_TIDetails!RecAmount) - IIf(IsNull(TmpRs!RecAmount), 0, TmpRs!RecAmount)
       RsTxn_TIDetails!RecTDSAmount = IIf(IsNull(RsTxn_TIDetails!RecTDSAmount), 0, RsTxn_TIDetails!RecTDSAmount) - IIf(IsNull(TmpRs!RecTDS), 0, TmpRs!RecTDS)
       RsTxn_TIDetails!RecOtherDed = IIf(IsNull(RsTxn_TIDetails!RecOtherDed), 0, RsTxn_TIDetails!RecOtherDed) - IIf(IsNull(TmpRs!RecOtherDed), 0, TmpRs!RecOtherDed)
       RsTxn_TIDetails.Update
    End If
    TmpRs.MoveNext
Next
tmp = "Delete From Txn_NewCashReceiptTIDetail Where CRID =  " & Val(mCRID) & " and SM_Prefix= '" & TxtPreFix.Text & "' And CRAAID=" & Val(TxtCRAAID)
Call Tmp4Table
End Sub

Private Sub SaveTIDetail()
Call DeleteTIDetail
Call TableTxn_NewCashReceiptTIDetail(" Where CRID= " & Val(mCRID))
Call TableTxn_TIDetails
For I = 2 To MSHFlexGrid3.Rows - 1
    If MSHFlexGrid3.TextMatrix(I, 0) <> "" And (Val(MSHFlexGrid3.TextMatrix(I, 7)) > 0 Or Val(MSHFlexGrid3.TextMatrix(I, 8)) > 0 Or Val(MSHFlexGrid3.TextMatrix(I, 9)) > 0) Then
       RsTxn_NewCashReceiptTIDetail.AddNew
       RsTxn_NewCashReceiptTIDetail!CRTIDID = getCRTIDetailMaxID
       RsTxn_NewCashReceiptTIDetail!CRID = Val(mCRID)
       RsTxn_NewCashReceiptTIDetail!CRAAID = Val(TxtCRAAID)
       RsTxn_NewCashReceiptTIDetail!SM_Prefix = TxtPreFix
       RsTxn_NewCashReceiptTIDetail!G_UID = GetGUID
       RsTxn_NewCashReceiptTIDetail!TIID = Val(MSHFlexGrid3.TextMatrix(I, 11))
       RsTxn_NewCashReceiptTIDetail!RecAmount = Val(MSHFlexGrid3.TextMatrix(I, 7))
       RsTxn_NewCashReceiptTIDetail!RecTDS = Val(MSHFlexGrid3.TextMatrix(I, 8))
       RsTxn_NewCashReceiptTIDetail!RecOtherDed = Val(MSHFlexGrid3.TextMatrix(I, 9))
       If MSHFlexGrid3.TextMatrix(I, 17) = "Y" Then
          RsTxn_NewCashReceiptTIDetail!DPID = MSHFlexGrid3.TextMatrix(I, 12)
       Else
          RsTxn_NewCashReceiptTIDetail!ClientID = Val(MSHFlexGrid3.TextMatrix(I, 12))
       End If
       RsTxn_NewCashReceiptTIDetail.Update
       RsTxn_NewCashReceiptTIDetail.Requery
       If RsTxn_TIDetails.RecordCount > 0 Then RsTxn_TIDetails.MoveFirst
       RsTxn_TIDetails.Find "TIID=" & Val(MSHFlexGrid3.TextMatrix(I, 11))
       If RsTxn_TIDetails.EOF = False Then
          RsTxn_TIDetails!RecAmount = IIf(IsNull(RsTxn_TIDetails!RecAmount), 0, RsTxn_TIDetails!RecAmount) + Val(MSHFlexGrid3.TextMatrix(I, 7))
          RsTxn_TIDetails!RecTDSAmount = IIf(IsNull(RsTxn_TIDetails!RecTDSAmount), 0, RsTxn_TIDetails!RecTDSAmount) + Val(MSHFlexGrid3.TextMatrix(I, 8))
          RsTxn_TIDetails!RecOtherDed = IIf(IsNull(RsTxn_TIDetails!RecOtherDed), 0, RsTxn_TIDetails!RecOtherDed) + Val(MSHFlexGrid3.TextMatrix(I, 9))
          RsTxn_TIDetails.Update
       End If
    End If
Next
End Sub

Private Sub TxtTDSAmt_LostFocus()
If TxtTDSAmt.Text = "" Then Exit Sub
If IsNumeric(TxtTDSAmt.Text) = False Or Val(TxtTDSAmt.Text) < 0 Then
   MsgBox "Invalid Numeric Format !!!"
   TxtTDSAmt.SetFocus
   Exit Sub
End If
TxtTDSAmt = Round(Val(TxtTDSAmt), 2)
'TxtBalRec.Text = Format(Val(TxtBalAmount) - (Val(TxtRecAmt) + Val(TxtTDSAmt)), "#####0.00")
TxtBalRec.Text = Format(Val(TxtBalAmount) - (Val(TxtRecAmt) + Val(TxtTDSAmt) + Val(TxtOtherDed)), "#####0.00")
If Val(TxtBalRec.Text) < 0 Then
   MsgBox "Balance Receivable cannot be less than zero. "
   TxtTDSAmt.Text = 0
   TxtTDSAmt.SetFocus
End If
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




Private Sub CmdPrintPreview_Click()
'Gen_mTimeStamp = GetTimeStamp
'Call Create_Xls(Gen_mTimeStamp & Gen_UserName & "_NewCashReceipt-" & Trim(TxtPreFix) & "" & Trim(TxtCRID) & ".xls")
'
''Call Xls_Print_Rpt
'
'Call Xls_Print_Rpt_New

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
oBook.SaveAs (Pat & "\Excel\" & FileName)

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

Public Sub Xls_Print_Rpt_New()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim oWS As Excel.Worksheet
Dim mRow As Double
'Dim mCol As Integer
Dim ans As Variant
MsgBox ("Wait till next message")
Set oXL = New Excel.Application
'Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "_NewCashReceipt-" & Trim(TxtPreFix) & "" & Trim(TxtCRID) & ".xls")

I = 2

Set oWS = oWB.Worksheets("Sheet" & 1)
''oWS.PageSetup.RightMargin = 0.05
'oWS.PageSetup.TopMargin = 0.05
'oWS.PageSetup.RightMargin = 0.05
'oWS.PageSetup.TopMargin = 4.6
oWS.Name = TxtPreFix & "-" & TxtCRID  'MSHFlexGrid2.TextMatrix(i, 23)
oWS.ClearArrows
mRow = 6

oWS.Range("a1:g1").Merge
oWS.Range("a:a").ColumnWidth = 10.86
oWS.Range("b:b").ColumnWidth = 12.86
oWS.Range("c:c").ColumnWidth = 12.86
oWS.Range("d:d").ColumnWidth = 12.86
oWS.Range("e:e").ColumnWidth = 12.86 '32.14
oWS.Range("f:f").ColumnWidth = 12.86
oWS.Range("g:g").ColumnWidth = 12.86

tmp = "a" & mRow & ":" & "h" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "NATIONAL COLLATERAL MANAGEMENT SERVICES LIMITED"
oWS.Cells(mRow, C + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 10

'oWS.Cells(1, 1).Font.ColorIndex = 4
mRow = mRow + 1

tmp = "a" & mRow & ":" & "h" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "Location -" & TxtLocationID
oWS.Cells(mRow, C + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 10

mRow = mRow + 1

tmp = "a" & mRow & ":" & "h" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = TxtInstrumentType.Text & " Receipt"
oWS.Cells(mRow, C + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 10

mRow = mRow + 1

oWS.Cells(mRow, C + 1) = "Receipt No"
oWS.Cells(mRow, C + 1).Font.Size = 7
oWS.Cells(mRow, C + 1).Font.Bold = True

tmp = "b" & mRow & ":" & "c" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 2) = TxtPreFix & "-" & TxtCRID
oWS.Cells(mRow, C + 2).Font.Size = 7
oWS.Cells(mRow, C + 2).HorizontalAlignment = -4108

oWS.Cells(mRow, C + 4) = "Receipt Date"
oWS.Cells(mRow, C + 4).Font.Size = 7
oWS.Cells(mRow, C + 4).Font.Bold = True

tmp = "e" & mRow & ":" & "h" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 5) = " " & Format(TxtCRDate, "dd-mmm-yyyy")
oWS.Cells(mRow, C + 5).Font.Size = 7
oWS.Cells(mRow, C + 5).HorizontalAlignment = -4108


mRow = mRow + 1

Dim Rsp As String
Rsp = NumStrConv(Val(TxtInstrumentAmount))
tmp = "a" & mRow & ":" & "h" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "Received with thanks from Shri " & TxtReceivedFrom & " a sum of Rs " & TxtInstrumentAmount & " ( Rupees " & Rsp & " )" & _
                         " as per the details enclosed, by " & CmbInstrumentType & " "
'No:................. (instrument no) dated ................. ( instrument date) drawn on ............................ (drawn on Bank) ........................ branch (drawn on branch)."
If TxtInstrumentType.Text <> "Cash" Then
   oWS.Cells(mRow, C + 1) = oWS.Cells(mRow, C + 1) & " No: " & TxtInstrumentNo & " dated " & TxtInstrumentDate & " drawn on " & TxtDrawnOn & " " & TxtBranchName & " branch (subject to realisation). "
End If
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).WrapText = True
oWS.Cells(mRow, C + 1).HorizontalAlignment = 2
oWS.Cells(mRow, C + 1).Font.Size = 10
oWS.Cells(mRow, C + 1).RowHeight = 60

mRow = mRow + 1
'mCol = 5

tmp = "a" & mRow & ":" & "d" & mRow
oWS.Range(tmp).Merge

tmp = "e" & mRow & ":" & "h" & mRow
oWS.Range(tmp).Merge

oWS.Cells(mRow, C + 5) = "Affix Hologram here"
oWS.Cells(mRow, C + 5).Font.Size = 7
oWS.Cells(mRow, C + 5).Font.Bold = True
oWS.Cells(mRow, C + 5).RowHeight = 45

mRow = mRow + 1

tmp = "a" & mRow & ":" & "d" & mRow
oWS.Range(tmp).Merge

oWS.Cells(mRow, C + 1) = "Name and Signature of authorised representative of NCMSL"
oWS.Cells(mRow, C + 1).Font.Size = 7
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).WrapText = True

tmp = "e" & mRow & ":" & "h" & mRow
oWS.Range(tmp).Merge

If IsNull(RsTxn_NewCashReceipt!HologramNo) Then
   oWS.Cells(mRow, C + 5) = "Hologram No:- Preview"
Else
  oWS.Cells(mRow, C + 5) = "Hologram No:- " & RsTxn_NewCashReceipt!HologramNo
End If
oWS.Cells(mRow, C + 5).Font.Size = 7
oWS.Cells(mRow, C + 5).Font.Bold = True
oWS.Cells(mRow, C + 5).RowHeight = 40

mRow = mRow + 1

oWS.Cells(mRow, C + 1) = "Created By :-"
oWS.Cells(mRow, C + 1).Font.Size = 7
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).WrapText = True

tmp = "b" & mRow & ":" & "c" & mRow
oWS.Range(tmp).Merge

Dim strSplitString() As String
Dim EmpNo_, CreatedDateTime_ As String

If TxtCreated <> "" Then
   strSplitString = Split(TxtCreated, ":-")
   EmpNo_ = Trim(strSplitString(0))
   CreatedDateTime_ = Trim(strSplitString(1))
Else
   EmpNo_ = "" 'Trim(strSplitString(0))
   CreatedDateTime_ = "" 'Trim(strSplitString(1))

End If

oWS.Cells(mRow, C + 2) = EmpNo_ & "-" & GetCreatedByName(EmpNo_)
oWS.Cells(mRow, C + 2).Font.Size = 7
oWS.Cells(mRow, C + 2).Font.Bold = True
oWS.Cells(mRow, C + 2).WrapText = True
oWS.Cells(mRow, C + 2).HorizontalAlignment = -4108

oWS.Cells(mRow, C + 4) = "Created On :-"
oWS.Cells(mRow, C + 4).Font.Size = 7
oWS.Cells(mRow, C + 4).Font.Bold = True
oWS.Cells(mRow, C + 4).WrapText = True

tmp = "e" & mRow & ":" & "h" & mRow
oWS.Range(tmp).Merge

oWS.Cells(mRow, C + 5) = CreatedDateTime_
oWS.Cells(mRow, C + 5).Font.Size = 7
oWS.Cells(mRow, C + 5).Font.Bold = True
oWS.Cells(mRow, C + 5).WrapText = True
oWS.Cells(mRow, C + 5).HorizontalAlignment = -4108

If LCase(TxtInstrumentType.Text) = "cheque" Then
   mRow = mRow + 1

   tmp = "a" & mRow & ":" & "h" & mRow
   oWS.Range(tmp).Merge
   oWS.Cells(mRow, C + 1) = "Note :cheque dishonour would attract charges of RS. 500."
   oWS.Cells(mRow, C + 1).Font.Size = 7
   oWS.Cells(mRow, C + 1).Font.Bold = True
   oWS.Cells(mRow, C + 1).WrapText = True
End If
mRow = mRow + 1
tmp = "a" & mRow & ":" & "h" & mRow
oWS.Range(tmp).Merge


k = mRow

For Cell = 6 To k
    tmp = "a" & Cell & ":" & "h" & Cell
    oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
    oWS.Range(tmp).Font.Name = "Zurich BT"
    'oWS.Range(tmp).Font.Size = 7
    oWS.Range(tmp).WrapText = True
Next Cell

''
''If Gen_ExcelPassword <> "" Then
''   oWS.Protect Gen_ExcelPassword, , , , , , True, True
''End If



'Set oWS = oWB.Worksheets("Sheet" & 2)


'oWS.Name = "Cash Receipt Details (Invoice-CDTF related)"
mRow = mRow + 2

If ChkAdv.Value = 1 Then
   k = mRow
   'oWS.Name = "Advance Details"
   oWS.Range("a1:g1").Merge
   oWS.Range("a:a").ColumnWidth = 11
   oWS.Range("b:b").ColumnWidth = 11
   oWS.Range("c:c").ColumnWidth = 23
   oWS.Range("d:d").ColumnWidth = 11
   oWS.Range("e:e").ColumnWidth = 22 '32.14
   oWS.Range("f:f").ColumnWidth = 10
    
    mStart = mRow
   tmp = "a" & mRow & ":" & "f" & mRow
   oWS.Range(tmp).Merge
   oWS.Cells(mRow, C + 1) = "Cash Receipt Details (Advance related) for Receipt No: " & TxtPreFix & "-" & TxtCRID & " Dated " & Format(TxtCRDate, "dd-mmm-yyyy")
   'oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
   oWS.Cells(mRow, C + 1).Font.Bold = True
   oWS.Cells(mRow, C + 1).Font.Size = 10
    
   mRow = mRow + 1
    
   oWS.Cells(mRow, C + 1) = "Exchange Type"
   oWS.Cells(mRow, C + 1).Font.Size = 7
   oWS.Cells(mRow, C + 1).Font.Bold = True
   
   oWS.Cells(mRow, C + 2) = "Client ID"
   oWS.Cells(mRow, C + 2).Font.Size = 7
   oWS.Cells(mRow, C + 2).Font.Bold = True
   
   oWS.Cells(mRow, C + 3) = "Client Name"
   oWS.Cells(mRow, C + 3).Font.Size = 7
   oWS.Cells(mRow, C + 3).Font.Bold = True
   
   oWS.Cells(mRow, C + 4) = "DP ID"
   oWS.Cells(mRow, C + 4).Font.Size = 7
   oWS.Cells(mRow, C + 4).Font.Bold = True
   
   oWS.Cells(mRow, C + 5) = "DP Name"
   oWS.Cells(mRow, C + 5).Font.Size = 7
   oWS.Cells(mRow, C + 5).Font.Bold = True
   
   oWS.Cells(mRow, C + 6) = "Amount"
   oWS.Cells(mRow, C + 6).Font.Size = 7
   oWS.Cells(mRow, C + 6).Font.Bold = True
   
   mRow = mRow + 1
    
   For k = 2 To MSHFlexGrid2.Rows - 2
       oWS.Cells(mRow, C + 1) = MSHFlexGrid2.TextMatrix(k, 1)
       oWS.Cells(mRow, C + 1).Font.Size = 7
       
       oWS.Cells(mRow, C + 2) = MSHFlexGrid2.TextMatrix(k, 2)
       oWS.Cells(mRow, C + 2).Font.Size = 7
        
       oWS.Cells(mRow, C + 3) = MSHFlexGrid2.TextMatrix(k, 3)
       oWS.Cells(mRow, C + 3).Font.Size = 7
       
       oWS.Cells(mRow, C + 4) = MSHFlexGrid2.TextMatrix(k, 4)
       oWS.Cells(mRow, C + 4).Font.Size = 7
       
       oWS.Cells(mRow, C + 5) = MSHFlexGrid2.TextMatrix(k, 5)
       oWS.Cells(mRow, C + 5).Font.Size = 7
       
       oWS.Cells(mRow, C + 6) = MSHFlexGrid2.TextMatrix(k, 6)
       oWS.Cells(mRow, C + 6).Font.Size = 7
      
       mRow = mRow + 1
       
   Next k

   k = mRow

    For Cell = mStart To k
        tmp = "a" & Cell & ":" & "f" & Cell
        oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
        oWS.Range(tmp).Font.Name = "Zurich BT"
        'oWS.Range(tmp).Font.Size = 7
        oWS.Range(tmp).WrapText = True
    Next Cell


Else
    'oWS.Name = "TI Details"
   ' oWS.Range("a1:g1").Merge
    oWS.Range("a:a").ColumnWidth = 15
    oWS.Range("b:b").ColumnWidth = 22
    oWS.Range("c:c").ColumnWidth = 15
    oWS.Range("d:d").ColumnWidth = 12
    oWS.Range("e:e").ColumnWidth = 13
    oWS.Range("f:f").ColumnWidth = 13
    oWS.Range("g:g").ColumnWidth = 10 '32.14
    oWS.Range("h:h").ColumnWidth = 13
    
    mStart = mRow
    tmp = "a" & mRow & ":" & "h" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 1) = "Cash Receipt Details (Invoice related) for Receipt No: " & TxtPreFix & "-" & TxtCRID & " Dated " & Format(TxtCRDate, "dd-mmm-yyyy")
    'oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
    oWS.Cells(mRow, C + 1).Font.Bold = True
    oWS.Cells(mRow, C + 1).Font.Size = 10
    
    mRow = mRow + 1
    'mCol = 5
    
    ''New
    oWS.Cells(mRow, C + 1) = "ClientID /DPID"
    oWS.Cells(mRow, C + 1).Font.Size = 7
    oWS.Cells(mRow, C + 1).Font.Bold = True
    
    oWS.Cells(mRow, C + 2) = "Client Name"
    oWS.Cells(mRow, C + 2).Font.Size = 7
    oWS.Cells(mRow, C + 2).Font.Bold = True
    
    oWS.Cells(mRow, C + 3) = "Invoice No"
    oWS.Cells(mRow, C + 3).Font.Size = 7
    oWS.Cells(mRow, C + 3).Font.Bold = True
    
    oWS.Cells(mRow, C + 4) = "Invoice Date"
    oWS.Cells(mRow, C + 4).Font.Size = 7
    oWS.Cells(mRow, C + 4).Font.Bold = True
    
    oWS.Cells(mRow, C + 5) = "Invoice Amount"
    oWS.Cells(mRow, C + 5).Font.Size = 7
    oWS.Cells(mRow, C + 5).Font.Bold = True
    
    oWS.Cells(mRow, C + 6) = "Received Amount"
    oWS.Cells(mRow, C + 6).Font.Size = 7
    oWS.Cells(mRow, C + 6).Font.Bold = True
    
    oWS.Cells(mRow, C + 7) = "TDS Amount"
    oWS.Cells(mRow, C + 7).Font.Size = 7
    oWS.Cells(mRow, C + 7).Font.Bold = True
    
    oWS.Cells(mRow, C + 8) = "Other Deduction"
    oWS.Cells(mRow, C + 8).Font.Size = 7
    oWS.Cells(mRow, C + 8).Font.Bold = True
    
    mRow = mRow + 1
    
    For k = 2 To MSHFlexGrid3.Rows - 2
    'k = 1
        oWS.Cells(mRow, C + 1) = MSHFlexGrid3.TextMatrix(k, 13)
        oWS.Cells(mRow, C + 1).Font.Size = 7
        
        oWS.Cells(mRow, C + 2) = MSHFlexGrid3.TextMatrix(k, 14)
        oWS.Cells(mRow, C + 2).Font.Size = 7
        
        oWS.Cells(mRow, C + 3) = MSHFlexGrid3.TextMatrix(k, 0)
        oWS.Cells(mRow, C + 3).Font.Size = 7
        
        oWS.Cells(mRow, C + 4) = " " & Format(MSHFlexGrid3.TextMatrix(k, 1), "dd-mmm-yyyy")
        oWS.Cells(mRow, C + 4).Font.Size = 7
        
        oWS.Cells(mRow, C + 5) = Val(MSHFlexGrid3.TextMatrix(k, 2))
        oWS.Cells(mRow, C + 5).Font.Size = 7
        
        oWS.Cells(mRow, C + 6) = Val(MSHFlexGrid3.TextMatrix(k, 7))
        oWS.Cells(mRow, C + 6).Font.Size = 7
        
        oWS.Cells(mRow, C + 7) = Val(MSHFlexGrid3.TextMatrix(k, 8))
        oWS.Cells(mRow, C + 7).Font.Size = 7
        
        oWS.Cells(mRow, C + 8) = Val(MSHFlexGrid3.TextMatrix(k, 9))
        oWS.Cells(mRow, C + 8).Font.Size = 7
        
        mRow = mRow + 1
    Next k
    For Cell = mStart To mRow
        tmp = "a" & Cell & ":" & "h" & Cell
        oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
        oWS.Range(tmp).Font.Name = "Zurich BT"
        'oWS.Range(tmp).Font.Size = 7
        oWS.Range(tmp).WrapText = True
    Next Cell
End If

If Gen_ExcelPassword <> "" Then
   oWS.Protect Gen_ExcelPassword, , , , , , True, True
End If
oWB.Save

Call Log_Print_History(TxtPreFix, "CR", TxtCRID, Gen_FLName)

MsgBox ("Excel file created !!! ")
oXL.Visible = True

End Sub


Private Sub CmdPrint_Click()
''If RsTxn_NewCashReceipt!HologramNo = "" Or IsNull(RsTxn_NewCashReceipt!HologramNo) Then
''   If LCase(UserType) = "power" Then
''      Check2.Visible = True
''   End If
''   Frame3.Visible = True
''   TxtHologramNo.SetFocus
''   Call TableMst_Hologram
''   Exit Sub
''Else
''   TxtHologramNo = ""
''   Frame3.Visible = False
''End If
''RsTxn_NewCashReceipt!Flag = "P"
''RsTxn_NewCashReceipt.Update
''Gen_mTimeStamp = GetTimeStamp
''Call Create_Xls(Gen_mTimeStamp & Gen_UserName & "_NewCashReceipt-" & Trim(TxtPreFix) & "" & Trim(TxtCRID) & ".xls")
'''Call Xls_Print_Rpt
''Call Xls_Print_Rpt_New
End Sub
Private Sub TxtOtherDed_LostFocus()
If TxtOtherDed.Text = "" Then Exit Sub
If IsNumeric(TxtOtherDed.Text) = False Or Val(TxtOtherDed.Text) < 0 Then
   MsgBox "Invalid Numeric Format !!!"
   TxtOtherDed.SetFocus
   Exit Sub
End If
TxtOtherDed = Round(Val(TxtOtherDed), 2)
TxtBalRec.Text = Format(Val(TxtBalAmount) - (Val(TxtRecAmt) + Val(TxtTDSAmt) + Val(TxtOtherDed)), "#####0.00")
If Val(TxtBalRec.Text) < 0 Then
   MsgBox "Balance Receivable cannot be less than zero."
   TxtOtherDed.Text = 0
   TxtOtherDed.SetFocus
End If
End Sub


Public Sub GetData()

Call ClearTITabControl

tmp = "Select CR.CRID,CR.SM_Prefix,Cr.CreatedBy,Cr.UpdatedBY,Cr.CreatedDate,CR.UpdatedDate,CR.SightedDate,LM.LocationName,CR.CRDate,CR.InstrumentType,CR.DrawnOn,CR.BranchName,CR.InstrumentNo,CR.InstrumentDate,"
tmp = tmp & " CR.InstrumentAmount ,CR.ReceivedFrom, CR.Remark, CR.Flag, CR.HologramNo,CR.G_UID,CR.DepositDate,CR.FAGRemarks 'DepositRemarks',MFB.BankName +'~'+ MFB.AccountNo as 'Bank',CR.AdvanceFlag,CR.ReturnFlag,CR.FAGPostingDate,CR.LocationID, CR.ChequeReturnChargesAmount"
tmp = tmp & " From Txn_NewCashReceipt CR"
tmp = tmp & " Inner Join Mst_Location LM On CR.LocationID = LM.LocationID"
tmp = tmp & " Left Join Mst_FAGBank MFB On CR.FAGBankID = MFB.FAGBankID "
tmp = tmp & " Where CRID =  " & mCRID & "  "
tmp = tmp & " and CR.SM_Prefix= '" & TxtPreFix.Text & "' "

Call TmpTable

TxtCRDate = Format(TmpRs!CRDate, "dd-MMM-yyyy")
TxtInstrumentType = TmpRs!InstrumentType
If LCase(TxtInstrumentType.Text) = "cheque" Then
   Label5.Caption = "Cheque No"
   Label7.Caption = "Cheque Date"
ElseIf LCase(TxtInstrumentType.Text) = "dd" Then
   Label5.Caption = "DD No"
   Label7.Caption = "DD Date"
ElseIf LCase(TxtInstrumentType.Text) = "rtgs" Then
   Label5.Caption = "RTGS No"
   Label7.Caption = "RTGS Date"
Else
   Label5.Caption = "Instrument No"
   Label7.Caption = "Instrument Date"
End If

TxtDrawnOn = TmpRs!DrawnOn
TxtBranchName = TmpRs!BranchName
TxtInstrumentNo = TmpRs!InstrumentNo
TxtInstrumentDate = Format(TmpRs!InstrumentDate, "dd-MMM-yyyy")
TxtInstrumentAmount = TmpRs!InstrumentAmount
TxtChequeReturnChargesAmount = IIf(IsNull(TmpRs!ChequeReturnChargesAmount), 0, TmpRs!ChequeReturnChargesAmount)
TxtReceivedFrom = TmpRs!ReceivedFrom
TxtRemark = TmpRs!Remark
TxtPreFix = TmpRs!SM_Prefix
TxtDepositeDate = IIf(IsNull(TmpRs!DepositDate), "", Format(TmpRs!DepositDate, "dd-MMM-yyyy"))
TxtFAGPostingDate.Text = IIf(IsNull(TmpRs!DepositDate), "", Format(TmpRs!FAGPostingDate, "dd-MMM-yyyy"))
TxtBankName = IIf(IsNull(TmpRs!Bank), "", TmpRs!Bank)
TxtRemark1 = IIf(IsNull(TmpRs!DepositRemarks), "", TmpRs!DepositRemarks)
TxtCRG_UID = TmpRs!G_UID
End Sub

Private Function GetCRAAID() As Double
Dim Retval As Double
tmp = "Select max(CRAAID) from Txn_CRAdvanceAllocation"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetCRAAID = Retval
End Function


Private Sub Grid_Head2()

MSHFlexGrid3.Clear
MSHFlexGrid3.Rows = 3
MSHFlexGrid3.Cols = 20
MSHFlexGrid3.FixedRows = 2
MSHFlexGrid3.WordWrap = True

MSHFlexGrid3.TextMatrix(1, 0) = "TI No"
MSHFlexGrid3.TextMatrix(1, 1) = "TI Date"
MSHFlexGrid3.TextMatrix(1, 2) = "TI Amount"
MSHFlexGrid3.TextMatrix(1, 3) = "Received Amt"
MSHFlexGrid3.TextMatrix(1, 4) = "TDS Deducted"
MSHFlexGrid3.TextMatrix(1, 5) = "Other Deduction"
MSHFlexGrid3.TextMatrix(1, 6) = "Balance Amt"
MSHFlexGrid3.TextMatrix(1, 7) = "Received Amt against this CR"
MSHFlexGrid3.TextMatrix(1, 8) = "TDS Deducted against this CR"
MSHFlexGrid3.TextMatrix(1, 9) = "Other Deduction against this CR"
MSHFlexGrid3.TextMatrix(1, 10) = "Balance Receivable after this CR"
MSHFlexGrid3.TextMatrix(1, 11) = "TIID"
MSHFlexGrid3.TextMatrix(1, 12) = "ClientIDRow/DPID"
MSHFlexGrid3.TextMatrix(1, 13) = "ClientID/DPID"
MSHFlexGrid3.TextMatrix(1, 14) = "Client Name"
MSHFlexGrid3.TextMatrix(1, 15) = "PostInvoice No"
MSHFlexGrid3.TextMatrix(1, 16) = "Entry No"
MSHFlexGrid3.TextMatrix(1, 17) = "DPFlag"
MSHFlexGrid3.TextMatrix(1, 18) = "NAV Code"
MSHFlexGrid3.TextMatrix(1, 19) = "Location"

MSHFlexGrid3.ColWidth(0) = 1600
MSHFlexGrid3.ColWidth(1) = 1400
MSHFlexGrid3.ColWidth(2) = 1200
MSHFlexGrid3.ColWidth(3) = 1200
MSHFlexGrid3.ColWidth(4) = 1200
MSHFlexGrid3.ColWidth(5) = 1200
MSHFlexGrid3.ColWidth(6) = 1200
MSHFlexGrid3.ColWidth(7) = 1300
MSHFlexGrid3.ColWidth(8) = 0
MSHFlexGrid3.ColWidth(9) = 0
MSHFlexGrid3.ColWidth(10) = 1300
MSHFlexGrid3.ColWidth(11) = 0
MSHFlexGrid3.ColWidth(12) = 0
MSHFlexGrid3.ColWidth(13) = 0
MSHFlexGrid3.ColWidth(14) = 0
MSHFlexGrid3.ColWidth(19) = 1800

MSHFlexGrid3.RowHeight(1) = 800

End Sub

Private Sub ClearHeader()

TxtCRID = ""
TxtInstrumentNo = ""
TxtInstrumentType.Text = ""
TxtReceivedFrom.Text = ""
TxtInstrumentAmount = ""
TxtChequeReturnChargesAmount = ""
TxtDrawnOn = ""
TxtBranchName = ""
TxtRemark = ""
TxtDepositeDate.Text = ""
TxtBankName.Text = ""
TxtRemark1.Text = ""
TxtFAGPostingDate = ""
TxtReturnFlag = ""
TxtCRDate = ""
TxtInstrumentDate = ""
TxtAllocAmt = ""
TxtUnAllocAmt = ""
TxtCRG_UID = ""
End Sub
Private Sub GetAllocatedAmt()

tmp = "Select sum(RecAmount) from Txn_NewCashReceiptTIDetail Where CRID=" & mCRID & " and SM_Prefix='" & TxtPreFix & "' "
If Val(TxtCRAAID) > 0 Then
   tmp = tmp & " And isnull(CRAAID,0)<>" & Val(TxtCRAAID)
End If
Call TmpTable

If TmpRs.RecordCount > 0 Then
   TxtAllocAmt = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
End If

'TxtUnAllocAmt = Val(TxtInstrumentAmount) - Val(TxtAllocAmt)
TxtUnAllocAmt = Val(TxtInstrumentAmount) - Val(TxtChequeReturnChargesAmount) - Val(TxtAllocAmt)
TxtUnAllocAmt = Round(TxtUnAllocAmt, 2)
End Sub

Private Function GetRemClientAdvAmt(mNAVCode_ As String, mROw_ As Variant) As Double
Dim mTotClientAdv, mTotAllocAdv As Variant
Dim mRetval As Double
mTotClientAdv = 0
mTotAllocAdv = 0
For I = 2 To MSHFlexGrid2.Rows - 1
    If Trim(MSHFlexGrid2.TextMatrix(I, 10)) = mNAVCode_ Or MSHFlexGrid2.TextMatrix(I, 11) = mNAVCode_ Then
       mTotClientAdv = mTotClientAdv + Val(MSHFlexGrid2.TextMatrix(I, 13))
    End If
Next

For I = 2 To MSHFlexGrid3.Rows - 1
    If Trim(MSHFlexGrid3.TextMatrix(I, 18)) = mNAVCode_ Then
       If I <> mROw_ Then
          mTotAllocAdv = mTotAllocAdv + Val(MSHFlexGrid3.TextMatrix(I, 7))
       End If
    End If
Next
mRetval = Format(mTotClientAdv - mTotAllocAdv, "#####0.00")
GetRemClientAdvAmt = mRetval
End Function


Private Sub AppendTIDetails()

''tmp = " Select TINo,TIDate,TIAmount,TIID,TI.ClientIDRow,ClientID,ClientName,TI.DepositoryID,DepositoryName,TI.PostInvoiceNo,TI.EntryNo,TI.NAVCode From  Txn_TIDetails TI "
''tmp = tmp & " Left Join Mst_Client MC on TI.ClientIDRow=MC.ClientIDRow"
''tmp = tmp & " Left Join Mst_Depository MD on TI.DepositoryID=MD.DepositoryID"
''tmp = tmp & " Where  TINo <> '' And TIAmount-((select isnull(Sum(RecAmount)+Sum(RecTDS)+Sum(RecOtherDed),0) From Txn_NewCashReceiptTIDetail TCRT"
''tmp = tmp & " Inner Join Txn_NewCashReceipt TCR on TCR.CRID= TCRT.CRID and TCR.SM_Prefix=TCRT.SM_Prefix"
''tmp = tmp & " Where TCRT.TIID=TI.TIID And TCR.ReturnFlag<>'Y'))>0 "
''
''If LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "admin" And LCase(Gen_UserRole) <> "superadmin" And LCase(Gen_UserRole) <> "fag" And LCase(Gen_UserRole) <> "head cmg" Then
''   tmp = tmp & " And (LocationID=" & mLocationID & " or LocationID is null)"
''End If
''
''tmp = tmp & " And (TI.NAVCode in "
''tmp = tmp & " (Select distinct NAVCode from Txn_NewCashReceiptClientDetail TCRD"
''tmp = tmp & " inner join Mst_Client MC on MC.ClientIDRow=TCRD.ClientIDRow"
''tmp = tmp & " Where TCRD.CRID =  " & Val(mCRID) & " and TCRD.SM_Prefix= '" & TxtPreFix.Text & "') OR"
''tmp = tmp & " TI.NAVCode in "
''tmp = tmp & " (Select distinct RNAVCode from Txn_NewCashReceiptClientDetail TCRD"
''tmp = tmp & " inner join Mst_Client MC on MC.ClientIDRow=TCRD.ClientIDRow"
''tmp = tmp & " Where TCRD.CRID =  " & Val(mCRID) & " and TCRD.SM_Prefix= '" & TxtPreFix.Text & "') OR"
''tmp = tmp & " TI.NAVCode in "
''tmp = tmp & " (Select distinct NAVCode from Txn_NewCashReceiptClientDetail TCRD"
''tmp = tmp & " inner join Mst_Depository MC on MC.DepositoryID=TCRD.DPID"
''tmp = tmp & " Where TCRD.CRID =  " & Val(mCRID) & " and TCRD.SM_Prefix= '" & TxtPreFix.Text & "'))"
''
''tmp = tmp & " And TI.TIID not in (Select TIID from Txn_NewCashReceiptTIDetail TCRT"
''tmp = tmp & " Where TCRT.CRID =  " & Val(mCRID) & " and TCRT.SM_Prefix= '" & TxtPreFix.Text & "' And TCRT.CRAAID=" & TxtCRAAID & ")"
''
''Call Tmp3Table
''
''tmp = " Select TCRT.TIID,Sum(TCRT.RecAmount) as RecAmount,Sum(RecTDS) as RecTDSAmount,Sum(RecOtherDed)as RecOtherDed  From Txn_NewCashReceiptTIDetail TCRT"
''tmp = tmp & " Inner Join Txn_NewCashReceipt TCR on TCR.CRID= TCRT.CRID and TCR.SM_Prefix=TCRT.SM_Prefix"
''tmp = tmp & " Where TCR.G_UID<>'" & TxtCRG_UID & "' And TCR.ReturnFlag<>'Y'"
''tmp = tmp & " Group By  TCRT.TIID"
''
''Call Tmp4Table
''
''tmp = " Select TCRT.TIID,Sum(TCRT.RecAmount) as RecAmount,Sum(RecTDS) as RecTDSAmount,Sum(RecOtherDed)as RecOtherDed  From Txn_NewCashReceiptTIDetail TCRT"
''tmp = tmp & " Inner Join Txn_NewCashReceipt TCR on TCR.CRID= TCRT.CRID and TCR.SM_Prefix=TCRT.SM_Prefix"
''tmp = tmp & " Where TCR.G_UID='" & TxtCRG_UID & "' And TCR.ReturnFlag<>'Y' And TCRT.CRAAID<>" & Val(TxtCRAAID)
''tmp = tmp & " Group By  TCRT.TIID"
''
''Call Tmp1Table


Wc1 = ""
If LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "admin" And LCase(Gen_UserRole) <> "superadmin" And LCase(Gen_UserRole) <> "fag" And LCase(Gen_UserRole) <> "fagadmin" And LCase(Gen_UserRole) <> "headcmg" Then
   Wc1 = Wc1 & " And (TI.LocationID=" & mLocationID & " or TI.LocationID is null)"
End If

Wc1 = Wc1 & " And (TI.NAVCode in "
Wc1 = Wc1 & " (Select distinct TCRD.NAVCode from Txn_NewCashReceiptClientDetail TCRD"
Wc1 = Wc1 & " inner join Mst_Client MC on MC.ClientID=TCRD.ClientID"
Wc1 = Wc1 & " Where TCRD.CRID =  " & Val(mCRID) & " and TCRD.SM_Prefix= '" & TxtPreFix.Text & "') OR"
Wc1 = Wc1 & " TI.NAVCode in "
Wc1 = Wc1 & " (Select distinct RNAVCode from Txn_NewCashReceiptClientDetail TCRD"
Wc1 = Wc1 & " inner join Mst_Client MC on MC.ClientID=TCRD.ClientID"
Wc1 = Wc1 & " Where TCRD.CRID =  " & Val(mCRID) & " and TCRD.SM_Prefix= '" & TxtPreFix.Text & "') OR"
Wc1 = Wc1 & " TI.NAVCode in "
Wc1 = Wc1 & " (Select distinct TCRD.NAVCode from Txn_NewCashReceiptClientDetail TCRD"
Wc1 = Wc1 & " inner join Mst_Depository MC on MC.DepositoryID=TCRD.DPID"
Wc1 = Wc1 & " Where TCRD.CRID =  " & Val(mCRID) & " and TCRD.SM_Prefix= '" & TxtPreFix.Text & "'))"

Wc1 = Wc1 & " And TI.TIID not in (Select TIID from Txn_NewCashReceiptTIDetail TCRT"
Wc1 = Wc1 & " Where TCRT.CRID =  " & Val(mCRID) & " and TCRT.SM_Prefix= '" & TxtPreFix.Text & "' And TCRT.CRAAID=" & TxtCRAAID & ")"

Call GetClientsRemInvoices(Wc1)

If TmpRs.RecordCount > 0 Then
   For gr = 1 To TmpRs.RecordCount
       rn = MSHFlexGrid3.Rows - 1

       MSHFlexGrid3.TextMatrix(rn, 0) = TmpRs!TINo
       MSHFlexGrid3.TextMatrix(rn, 1) = IIf(IsNull(TmpRs!TIDate), "", Format(TmpRs!TIDate, "dd-MMM-yyyy"))
       MSHFlexGrid3.TextMatrix(rn, 2) = IIf(IsNull(TmpRs!TIAmount), 0, TmpRs!TIAmount)
       MSHFlexGrid3.TextMatrix(rn, 3) = IIf(IsNull(TmpRs!RecAmount), 0, TmpRs!RecAmount)
       MSHFlexGrid3.TextMatrix(rn, 4) = IIf(IsNull(TmpRs!RecTDS), 0, TmpRs!RecTDS)
       MSHFlexGrid3.TextMatrix(rn, 5) = IIf(IsNull(TmpRs!RecOtherDed), 0, TmpRs!RecOtherDed)
            
       MSHFlexGrid3.TextMatrix(rn, 6) = Format(Val(MSHFlexGrid3.TextMatrix(rn, 2)) - (Val(MSHFlexGrid3.TextMatrix(rn, 3)) + Val(MSHFlexGrid3.TextMatrix(rn, 4)) + Val(MSHFlexGrid3.TextMatrix(rn, 5))), "#####0.00")
       
       MSHFlexGrid3.TextMatrix(rn, 11) = TmpRs!TIID
       
       tmp = "Select ShowDP,DPID,TCRD.ClientIDRow,ClientID,ClientName,DepositoryName,Amount"
       tmp = tmp & " From Txn_NewCashReceiptClientDetail TCRD"
       tmp = tmp & " left join Mst_Client MC on MC.ClientID=TCRD.ClientID"
       tmp = tmp & " Left join Mst_Depository MD on MD.DepositoryID=TCRD.DPID"
       tmp = tmp & " Where  TCRD.CRID = " & mCRID & " and TCRD.SM_Prefix= '" & TxtPreFix & "' And ( MC.NAVCode='" & TmpRs!NavCode & "' or MC.RNAVCode='" & TmpRs!NavCode & "' or MD.NAVCode='" & TmpRs!NavCode & "')"
       
       Call Tmp6Table
       
       If TmpRs6.RecordCount > 0 Then
          If TmpRs6!ShowDP = 0 Then
              MSHFlexGrid3.TextMatrix(rn, 12) = TmpRs6!ClientID
              MSHFlexGrid3.TextMatrix(rn, 13) = TmpRs6!ClientID
              MSHFlexGrid3.TextMatrix(rn, 14) = TmpRs6!ClientName
              MSHFlexGrid3.TextMatrix(rn, 17) = "N"
           Else
              MSHFlexGrid3.TextMatrix(rn, 12) = TmpRs6!DPID
              MSHFlexGrid3.TextMatrix(rn, 13) = TmpRs6!DPID
              MSHFlexGrid3.TextMatrix(rn, 14) = TmpRs6!DepositoryName
              MSHFlexGrid3.TextMatrix(rn, 17) = "Y"
           End If
       End If
       
       MSHFlexGrid3.TextMatrix(rn, 15) = IIf(IsNull(TmpRs!PostInvoiceNo), 0, TmpRs!PostInvoiceNo)
       MSHFlexGrid3.TextMatrix(rn, 16) = IIf(IsNull(TmpRs!EntryNo), 0, TmpRs!EntryNo)
       MSHFlexGrid3.TextMatrix(rn, 18) = IIf(IsNull(TmpRs!NavCode), "", TmpRs!NavCode)
       MSHFlexGrid3.TextMatrix(rn, 19) = IIf(IsNull(TmpRs!Locationname), "", TmpRs!Locationname)
       TmpRs.MoveNext
       MSHFlexGrid3.Rows = MSHFlexGrid3.Rows + 1
   Next
End If
End Sub

