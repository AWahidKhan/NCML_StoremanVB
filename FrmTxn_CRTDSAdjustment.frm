VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmTxn_CRTDSAdjustment 
   Caption         =   "TDS Adjustment"
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10170
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   6810
   ScaleWidth      =   10170
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      Height          =   9855
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   15375
      Begin VB.Frame Frame0 
         Height          =   1365
         Left            =   600
         TabIndex        =   19
         Top             =   6720
         Visible         =   0   'False
         Width           =   15240
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
            ItemData        =   "FrmTxn_CRTDSAdjustment.frx":0000
            Left            =   6810
            List            =   "FrmTxn_CRTDSAdjustment.frx":000D
            TabIndex        =   90
            Top             =   360
            Width           =   2055
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
            Left            =   13785
            TabIndex        =   23
            Top             =   360
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
            Left            =   12750
            TabIndex        =   22
            Top             =   360
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
            ItemData        =   "FrmTxn_CRTDSAdjustment.frx":002E
            Left            =   2880
            List            =   "FrmTxn_CRTDSAdjustment.frx":0030
            Sorted          =   -1  'True
            TabIndex        =   21
            Top             =   360
            Width           =   3885
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
            Left            =   8925
            MaxLength       =   15
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   360
            Width           =   1245
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   1575
            Left            =   120
            TabIndex        =   24
            Top             =   840
            Width           =   9735
            _ExtentX        =   17171
            _ExtentY        =   2778
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
         Begin MSComCtl2.DTPicker STo 
            Height          =   360
            Left            =   1475
            TabIndex        =   25
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
            Format          =   118882305
            CurrentDate     =   36494
         End
         Begin MSComCtl2.DTPicker SFrom 
            Height          =   360
            Left            =   105
            TabIndex        =   26
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
            Format          =   118882305
            CurrentDate     =   36494
         End
         Begin VB.Label Label3 
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
            Left            =   7440
            TabIndex        =   91
            Top             =   120
            Width           =   720
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
            TabIndex        =   30
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
            TabIndex        =   29
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
            TabIndex        =   28
            Top             =   120
            Width           =   855
         End
         Begin VB.Label Label19 
            Caption         =   "ID"
            Height          =   255
            Left            =   9330
            TabIndex        =   27
            Top             =   135
            Width           =   405
         End
      End
      Begin VB.Frame ButtonFrm 
         Height          =   1095
         Left            =   120
         TabIndex        =   43
         Top             =   8670
         Width           =   14985
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
            Width           =   2355
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
            Left            =   2580
            TabIndex        =   54
            Top             =   585
            Width           =   2355
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
            Left            =   4935
            TabIndex        =   53
            Top             =   585
            Width           =   2355
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
            Left            =   7290
            TabIndex        =   52
            Top             =   585
            Width           =   2355
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
            Left            =   9645
            TabIndex        =   51
            Top             =   585
            Width           =   2355
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
            Width           =   2355
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
            Left            =   2580
            TabIndex        =   49
            Top             =   165
            Width           =   2355
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
            Left            =   4935
            TabIndex        =   48
            Top             =   165
            Width           =   2355
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
            Left            =   7290
            TabIndex        =   47
            Top             =   165
            Width           =   2355
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
            Left            =   9645
            TabIndex        =   46
            Top             =   165
            Width           =   2355
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
            Left            =   12000
            TabIndex        =   45
            Top             =   585
            Width           =   2355
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
            Left            =   12000
            TabIndex        =   44
            Top             =   165
            Width           =   2355
         End
      End
      Begin VB.Frame Frame1 
         Height          =   3705
         Left            =   120
         TabIndex        =   31
         Top             =   120
         Width           =   14895
         Begin VB.TextBox TxtNAVCode 
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
            Left            =   2400
            TabIndex        =   98
            Top             =   1012
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.CheckBox ChkNAVCode 
            Enabled         =   0   'False
            Height          =   255
            Left            =   2040
            TabIndex        =   96
            Top             =   1065
            Width           =   255
         End
         Begin VB.ComboBox CmbCNonNcdexClient 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   3435
            Sorted          =   -1  'True
            TabIndex        =   4
            Top             =   1019
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.ComboBox CmbDPName 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   4275
            Sorted          =   -1  'True
            TabIndex        =   5
            Top             =   1035
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox TxtFAGRemarks 
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
            Height          =   435
            Left            =   2010
            MaxLength       =   200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   8
            Top             =   2760
            Width           =   12720
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
            ItemData        =   "FrmTxn_CRTDSAdjustment.frx":0032
            Left            =   2010
            List            =   "FrmTxn_CRTDSAdjustment.frx":003F
            TabIndex        =   6
            Top             =   1858
            Width           =   4095
         End
         Begin VB.TextBox TxtCDPName 
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
            Left            =   8235
            Locked          =   -1  'True
            TabIndex        =   60
            TabStop         =   0   'False
            Top             =   1425
            Width           =   3375
         End
         Begin VB.TextBox TxtCDPID 
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
            Left            =   7035
            Locked          =   -1  'True
            TabIndex        =   59
            TabStop         =   0   'False
            Top             =   1431
            Width           =   1215
         End
         Begin VB.TextBox TxtCClientName 
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
            Left            =   3555
            Locked          =   -1  'True
            TabIndex        =   58
            TabStop         =   0   'False
            Top             =   1431
            Width           =   3495
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Show DP "
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
            Left            =   13200
            TabIndex        =   14
            Top             =   645
            Width           =   1140
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
            ItemData        =   "FrmTxn_CRTDSAdjustment.frx":0060
            Left            =   10080
            List            =   "FrmTxn_CRTDSAdjustment.frx":0062
            Sorted          =   -1  'True
            TabIndex        =   2
            Top             =   607
            Width           =   2775
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
            Height          =   435
            Left            =   2010
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   7
            Top             =   2270
            Width           =   12720
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
            Height          =   360
            Left            =   10035
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   35
            Top             =   1858
            Width           =   1560
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
            Left            =   7575
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   615
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
            Left            =   2010
            Sorted          =   -1  'True
            TabIndex        =   1
            Top             =   607
            Width           =   5565
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
            Left            =   10080
            Locked          =   -1  'True
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   195
            Width           =   1440
         End
         Begin VB.TextBox TxtAdjID 
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
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   195
            Width           =   2280
         End
         Begin MSComCtl2.DTPicker TxtAdjDate 
            Height          =   375
            Left            =   6120
            TabIndex        =   0
            Top             =   195
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
            Format          =   118882305
            CurrentDate     =   39884
         End
         Begin MSComCtl2.DTPicker TxtFAGPostingDate 
            Height          =   375
            Left            =   2010
            TabIndex        =   9
            Top             =   3240
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
            Format          =   118882305
            CurrentDate     =   39884
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
            Left            =   2010
            Locked          =   -1  'True
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   607
            Width           =   5550
         End
         Begin VB.TextBox TxtStatus 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
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
            TabIndex        =   88
            Top             =   1858
            Width           =   4095
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
            Left            =   7035
            Sorted          =   -1  'True
            TabIndex        =   66
            Top             =   1431
            Width           =   4575
         End
         Begin VB.TextBox TxtExchangeTypeID 
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
            Left            =   10080
            Locked          =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   89
            TabStop         =   0   'False
            Top             =   607
            Width           =   2775
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
            Left            =   2010
            TabIndex        =   3
            Top             =   1425
            Width           =   1575
         End
         Begin VB.Label lblNAVCode 
            AutoSize        =   -1  'True
            Caption         =   "Search by NAV Code"
            Height          =   195
            Left            =   240
            TabIndex        =   97
            Top             =   1095
            Width           =   1515
         End
         Begin VB.Label Label5 
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
            Height          =   300
            Left            =   240
            TabIndex        =   87
            Top             =   3277
            Width           =   1620
         End
         Begin VB.Label Label2 
            Caption         =   "FAG Remark"
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
            TabIndex        =   86
            Top             =   2857
            Width           =   1800
         End
         Begin VB.Label Label7 
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
            Height          =   300
            Left            =   240
            TabIndex        =   85
            Top             =   1888
            Width           =   1800
         End
         Begin VB.Label Label28 
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
            Left            =   240
            TabIndex        =   65
            Top             =   1491
            Width           =   975
         End
         Begin VB.Label Label27 
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
            Left            =   4995
            TabIndex        =   64
            Top             =   1065
            Width           =   1095
         End
         Begin VB.Label Label26 
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
            Left            =   7320
            TabIndex        =   63
            Top             =   1080
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label Label25 
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
            Left            =   9435
            TabIndex        =   62
            Top             =   1065
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label Label24 
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
            Left            =   7995
            TabIndex        =   61
            Top             =   1095
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label14 
            Caption         =   "Exchange Type"
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
            Left            =   8280
            TabIndex        =   57
            Top             =   660
            Width           =   1695
         End
         Begin VB.Label LblLocationID 
            Caption         =   "Adjustment No"
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
            TabIndex        =   42
            Top             =   225
            Width           =   1680
         End
         Begin VB.Label LblFlag 
            Caption         =   "Adj. Date"
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
            Left            =   4620
            TabIndex        =   41
            Top             =   225
            Width           =   1800
         End
         Begin VB.Label Label4 
            Caption         =   "Remark"
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
            TabIndex        =   40
            Top             =   2337
            Width           =   1800
         End
         Begin VB.Label Label6 
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
            Height          =   300
            Left            =   8280
            TabIndex        =   39
            Top             =   1888
            Width           =   960
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
            Left            =   240
            TabIndex        =   38
            Top             =   645
            Width           =   1440
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
            Left            =   8280
            TabIndex        =   37
            Top             =   255
            Width           =   975
         End
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   4830
         Left            =   120
         TabIndex        =   56
         Top             =   3840
         Width           =   14895
         _ExtentX        =   26273
         _ExtentY        =   8520
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
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
         TabCaption(0)   =   "TI Details"
         TabPicture(0)   =   "FrmTxn_CRTDSAdjustment.frx":0064
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "MSHFlexGrid3"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Frame2"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Log History"
         TabPicture(1)   =   "FrmTxn_CRTDSAdjustment.frx":0080
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "TxtCreated"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "TxtUpdated"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "Label13"
         Tab(1).Control(3)=   "Label12"
         Tab(1).ControlCount=   4
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
            TabIndex        =   81
            TabStop         =   0   'False
            Top             =   480
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
            TabIndex        =   80
            TabStop         =   0   'False
            Top             =   1065
            Width           =   8970
         End
         Begin VB.Frame Frame2 
            Height          =   1785
            Left            =   120
            TabIndex        =   67
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
               Left            =   10200
               Locked          =   -1  'True
               TabIndex        =   101
               TabStop         =   0   'False
               Top             =   1290
               Width           =   2625
            End
            Begin VB.TextBox TxtSearchBy 
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
               MaxLength       =   50
               TabIndex        =   99
               Top             =   1290
               Width           =   3885
            End
            Begin VB.TextBox TxtTISTAmount 
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
               Left            =   2280
               Locked          =   -1  'True
               TabIndex        =   94
               TabStop         =   0   'False
               Top             =   1290
               Width           =   2085
            End
            Begin VB.TextBox TxtTIAmount 
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
               TabIndex        =   92
               TabStop         =   0   'False
               Top             =   1290
               Width           =   2085
            End
            Begin VB.ComboBox CmbExpenseTypeID 
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
               ItemData        =   "FrmTxn_CRTDSAdjustment.frx":009C
               Left            =   8385
               List            =   "FrmTxn_CRTDSAdjustment.frx":009E
               Sorted          =   -1  'True
               TabIndex        =   12
               Top             =   600
               Width           =   2895
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
               Left            =   6690
               MaxLength       =   15
               TabIndex        =   11
               Top             =   600
               Width           =   1650
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
               Left            =   1770
               Locked          =   -1  'True
               TabIndex        =   72
               TabStop         =   0   'False
               Top             =   600
               Width           =   1650
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
               TabIndex        =   71
               TabStop         =   0   'False
               Top             =   600
               Width           =   1680
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
               Left            =   14520
               TabIndex        =   70
               Top             =   525
               Visible         =   0   'False
               Width           =   1710
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
               Left            =   11280
               Locked          =   -1  'True
               TabIndex        =   69
               TabStop         =   0   'False
               Top             =   600
               Width           =   1905
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
               TabIndex        =   13
               Top             =   600
               Width           =   1335
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
               Left            =   3420
               Locked          =   -1  'True
               TabIndex        =   68
               TabStop         =   0   'False
               Top             =   600
               Width           =   1620
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
               Left            =   5040
               MaxLength       =   15
               TabIndex        =   10
               Top             =   600
               Width           =   1650
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
               Left            =   11115
               TabIndex        =   102
               Top             =   990
               Width           =   765
            End
            Begin VB.Label Label11 
               Caption         =   "Search TI No"
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
               Left            =   4800
               TabIndex        =   100
               Top             =   1320
               Width           =   1440
            End
            Begin VB.Label Label10 
               Caption         =   "TI Tax Amount."
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
               Left            =   2527
               TabIndex        =   95
               Top             =   1005
               Width           =   1590
            End
            Begin VB.Label Label9 
               Caption         =   "TI Amount."
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
               Left            =   555
               TabIndex        =   93
               Top             =   1005
               Width           =   1110
            End
            Begin VB.Label Label1 
               Caption         =   "Expense Ac."
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
               Left            =   9285
               TabIndex        =   84
               Top             =   210
               Width           =   1215
            End
            Begin VB.Label Label17 
               Caption         =   "Other Deduction     (with Tax)"
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
               Left            =   6705
               TabIndex        =   78
               Top             =   105
               Width           =   1575
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
               Left            =   11280
               TabIndex        =   77
               Top             =   210
               Width           =   1875
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
               Left            =   615
               TabIndex        =   76
               Top             =   210
               Width           =   645
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
               Left            =   3600
               TabIndex        =   75
               Top             =   210
               Width           =   1395
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
               Left            =   2160
               TabIndex        =   74
               Top             =   210
               Width           =   780
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
               Left            =   5040
               TabIndex        =   73
               Top             =   210
               Width           =   1680
            End
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
            Height          =   2535
            Left            =   120
            TabIndex        =   79
            Top             =   2160
            Width           =   14670
            _ExtentX        =   25876
            _ExtentY        =   4471
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
            TabIndex        =   83
            Top             =   540
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
            TabIndex        =   82
            Top             =   1125
            Width           =   975
         End
      End
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
      Left            =   15120
      Locked          =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1320
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   480
      TabIndex        =   16
      Top             =   10080
      Width           =   10935
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   9015
      Left            =   15120
      TabIndex        =   15
      Top             =   1440
      Width           =   255
   End
End
Attribute VB_Name = "FrmTxn_CRTDSAdjustment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mAgentID, mClientIDRow, mSDepoBankID As Variant
Dim mExchangeTypeID As Double
Dim mLocationID As Double
Dim mExpenseTypeID As Double
Dim Edit_Flg As Variant
Dim mClientNAVCode As String
Dim mClientRNAVCode As String
Dim mDPNAVCode As String
Dim mStatus, mDPID As Variant
Dim Wc1  As String
Dim mSTAmount As Variant

Dim mSearchClientIDRow As Variant
Dim mSearchbyNAV As Boolean

Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)
Frame1.Enabled = True
Frame2.Enabled = True
mSearchbyNAV = False
Call Grid_Head2
TxtAdjDate.Enabled = True
TxtG_UID = ""
CmbLocationID.Visible = True
TxtLocationID.Visible = False
TxtAdjID.Locked = True
TxtRemark.Locked = False

SSTab1.Tab = 0

TxtAdjID = ""

TxtAdjDate.Value = Date
CmbLocationID.Text = ""
Call ClearTITabControl
TxtRemark = ""
TxtCreated = ""
TxtUpdated = ""
TxtPreFix = ""
CmdPrint.Enabled = False

CmbStatus.Visible = False
TxtStatus.Visible = True
CmbStatus.Text = "Pending"
TxtStatus.Text = "Pending"
mStatusFlag = "P"

TxtClientID.Locked = False
TxtClientID = ""
TxtCClientName = ""
TxtCDPID = ""
TxtCDPName = ""
CmbClientDP.Text = ""
CmbClientDP.Visible = False

mExchangeTypeID = 0

CmbExchangeTypeID.Visible = True
TxtExchangeTypeID.Visible = False
CmbExchangeTypeID.Text = ""

CmbCNonNcdexClient.Text = ""
CmbDPName.Text = ""
CmbDPName.Enabled = True
TxtFAGRemarks.Locked = True
TxtFAGPostingDate.Visible = False
Label5.Visible = False
TxtAmount = ""
TxtFAGRemarks = ""
Check3.Enabled = True
Check3.Value = 0
ChkNAVCode.Enabled = True
ChkNAVCode.Value = 0
If RsTxn_CRTDSAdjustment.RecordCount > 0 Then
   TxtAdjDate.SetFocus
End If

End Sub

Private Sub ChkNAVCode_Click()
If CmbExchangeTypeID.Text <> "" Then
    If ChkNAVCode.Value = vbChecked Then
        TxtNAVCode.Visible = True
    Else
        TxtNAVCode.Visible = False
    End If
Else
    MsgBox "Select Exchange Type first!!!"
    ChkNAVCode.Value = vbUnchecked
    CmbExchangeTypeID.SetFocus
End If
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
    Call Grid_Head2
   ' Call ClearClientTabControl
End If

mLocationID = RsMst_Location!LocationID
TxtPreFix = GetPreFixByLocation(RsMst_Location!LocationID, "P")
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
   ChkNAVCode.Value = 0
   TxtNAVCode.Visible = False
   TxtNAVCode.Text = ""
   mExchangeTypeID = 0
   CmbCNonNcdexClient.Clear
   CmbClientDP.Clear
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
   CmbCNonNcdexClient.Clear
   CmbClientDP.Clear
   CmbDPName.Clear
End If
mExchangeTypeID = RsMst_ExchangeType!ExchangeTypeID

If mExchangeTypeID = 1 Then
   CmbCNonNcdexClient.Width = 9600
   CmbCNonNcdexClient.Top = TxtClientID.Top
   CmbCNonNcdexClient.Left = TxtClientID.Left
   CmbCNonNcdexClient.Visible = True
   Label25.Visible = False
   Label26.Visible = False
   Label28.Visible = False
   Check3.Value = 0
   Check3.Enabled = False
Else
   Label25.Visible = True
   Label26.Visible = True
   Label28.Visible = True
   CmbCNonNcdexClient.Visible = False
   Check3.Enabled = True
   Call Check3_Click
End If

End Sub

Private Sub Check3_Click()
If Check3.Value = 1 Then
   Label25.Visible = False
   Label26.Visible = False
   Label28.Visible = False
   CmbDPName.Visible = True
   CmbDPName.Width = 9600 'CmbNonNcdexClient.Width
   CmbDPName.Top = TxtClientID.Top 'CmbNonNcdexClient.Top
   CmbDPName.Left = TxtClientID.Left 'CmbNonNcdexClient.Left
   CmbCNonNcdexClient.Visible = False
   SSTab1.Tab = 0
Else
   Label25.Visible = True
   Label26.Visible = True
   Label28.Visible = True
'   CmbNonNcdexClient.Visible = True
'   CmbNonNcdexClient.Width = CmbDPName.Width
'   CmbNonNcdexClient.Top = CmbDPName.Top
'   CmbNonNcdexClient.Left = CmbDPName.Left
   CmbDPName.Visible = False
End If

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


Private Sub CmbSStatus_LostFocus()
If CmbSStatus.Text = "" Then
   Exit Sub
End If

If LCase(CmbSStatus.Text) <> "approved" And LCase(CmbSStatus.Text) <> "pending" And LCase(CmbSStatus.Text) <> "rejected" Then
   MsgBox "Invalid selection!!!"
   CmbSStatus.SetFocus
   Exit Sub
End If

End Sub
Private Sub CmdAdd2List_Click()

If TxtTINo = "" Then Exit Sub

If TxtTDSAmt = "" Then
   MsgBox "Blank 'TDS Deducted' not allowed!!"
   TxtTDSAmt.SetFocus
   Exit Sub
End If

If Val(TxtOtherDed) > 0 And CmbExpenseTypeID.Text = "" Then
   MsgBox "Expense Acc. is Compulsory for Other Deduction!!"
   CmbExpenseTypeID.SetFocus
   Exit Sub
End If

If Val(TxtOtherDed) <= 0 And CmbExpenseTypeID.Text <> "" Then
   MsgBox "'Other Deduction' is Compulsory for Expense Acc.!!"
   TxtOtherDed.SetFocus
   Exit Sub
End If

r = MSHFlexGrid3.RowSel

MSHFlexGrid3.TextMatrix(0, 8) = Val(MSHFlexGrid3.TextMatrix(0, 8)) - Val(MSHFlexGrid3.TextMatrix(r, 8))
MSHFlexGrid3.TextMatrix(0, 7) = Val(MSHFlexGrid3.TextMatrix(0, 7)) - Val(MSHFlexGrid3.TextMatrix(r, 7))
MSHFlexGrid3.TextMatrix(0, 9) = Val(MSHFlexGrid3.TextMatrix(0, 9)) - Val(MSHFlexGrid3.TextMatrix(r, 9))
MSHFlexGrid3.TextMatrix(0, 15) = Val(MSHFlexGrid3.TextMatrix(0, 15)) - Val(MSHFlexGrid3.TextMatrix(r, 15))
MSHFlexGrid3.TextMatrix(0, 16) = Val(MSHFlexGrid3.TextMatrix(0, 16)) - Val(MSHFlexGrid3.TextMatrix(r, 16))

MSHFlexGrid3.TextMatrix(r, 7) = TxtTDSAmt.Text
MSHFlexGrid3.TextMatrix(r, 8) = TxtOtherDed.Text
MSHFlexGrid3.TextMatrix(r, 9) = TxtBalRec.Text
MSHFlexGrid3.TextMatrix(r, 10) = CmbExpenseTypeID.Text
MSHFlexGrid3.TextMatrix(r, 11) = mExpenseTypeID
If Val(TxtOtherDed.Text) <> 0 Then
   MSHFlexGrid3.TextMatrix(r, 16) = CalSTAmount(TxtTIAmount, TxtTISTAmount, TxtOtherDed)
   'MSHFlexGrid3.TextMatrix(r, 16) = CalSTAmount(Val(TxtTIAmount) - Val(TxtTISTAmount), TxtTISTAmount, TxtOtherDed)
Else
   MSHFlexGrid3.TextMatrix(r, 16) = 0
End If
MSHFlexGrid3.TextMatrix(0, 8) = Val(MSHFlexGrid3.TextMatrix(0, 8)) + Val(MSHFlexGrid3.TextMatrix(r, 8))
MSHFlexGrid3.TextMatrix(0, 7) = Val(MSHFlexGrid3.TextMatrix(0, 7)) + Val(MSHFlexGrid3.TextMatrix(r, 7))
MSHFlexGrid3.TextMatrix(0, 9) = Val(MSHFlexGrid3.TextMatrix(0, 9)) + Val(MSHFlexGrid3.TextMatrix(r, 9))
MSHFlexGrid3.TextMatrix(0, 15) = Val(MSHFlexGrid3.TextMatrix(0, 15)) + Val(MSHFlexGrid3.TextMatrix(r, 15))
MSHFlexGrid3.TextMatrix(0, 16) = Val(MSHFlexGrid3.TextMatrix(0, 16)) + Val(MSHFlexGrid3.TextMatrix(r, 16))

Call ClearTITabControl

'TxtAmount.Text = Val(TxtAmount.Text) + Val(MSHFlexGrid3.TextMatrix(0, 7)) + Val(MSHFlexGrid3.TextMatrix(0, 8))

TxtAmount.Text = Val(MSHFlexGrid3.TextMatrix(0, 7)) + Val(MSHFlexGrid3.TextMatrix(0, 8))

End Sub



Private Sub CmdSearch_Click()

If SFrom.Value > STo.Value Then
   MsgBox "'From Date' must be less than 'To date'!!!"
   SFrom.SetFocus
   Exit Sub
End If

tmp = "Select CR.TDSAID,LM.LocationName,CR.AdjDate,ME.ExchangeType,ClientName+'~'+MC.ClientID,DepositoryName+'~'+MD.DepositoryID, CR.Remark, CR.Flag ,CR.BatchID,CR.G_UID"
tmp = tmp & " From Txn_CRTDSAdjustment CR "
tmp = tmp & " Inner Join Mst_Location LM On CR.LocationID = LM.LocationID"
tmp = tmp & " Inner Join Mst_ExchangeType ME On CR.ExchangeTypeID= ME.ExchangeTypeID"
tmp = tmp & " Left Join Mst_Client MC On CR.ClientIDRow = MC.ClientIDRow"
tmp = tmp & " left Join Mst_Depository MD On CR.DPID = MD.DepositoryID"
If Gen_DBType = "MDB" Then
   tmp = tmp & " Where CR.AdjDate Between #" & Format(SFrom, Gen_DatFormat) & "# And #" & Format(STo, Gen_DatFormat) & "#) "
Else
    tmp = tmp & " where CR.AdjDate Between '" & Format(SFrom, Gen_DatFormat) & "' And '" & Format(STo, Gen_DatFormat) & "' "
End If

If CmbSLocation.Text <> "" Then
   tmp = tmp & " AND LM.LocationName = '" & CmbSLocation.Text & "' "
Else
   If Gen_WcLocation <> "" Then
      tmp = tmp & " AND LM." & Trim(Gen_WcLocation)
   End If
End If
If TxtSCRID <> "" Then
   tmp = tmp & " And CR.TDSAID = " & Val(TxtSCRID)
End If

If CmbSStatus.Text <> "" Then
    tmp = tmp & " And  CR.Flag = '" & Left(CmbSStatus.Text, 1) & "'"
End If

Call TmpTable
Call Grid_Head
For iu = 1 To TmpRs.RecordCount
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
    For C = 0 To MSHFlexGrid1.Cols - 1
       MSHFlexGrid1.TextMatrix(iu, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
    Next
    MSHFlexGrid1.TextMatrix(iu, 2) = Format(MSHFlexGrid1.TextMatrix(iu, 2), "dd-mmm-yyyy")
    If MSHFlexGrid1.TextMatrix(iu, 7) = "P" Then
       MSHFlexGrid1.TextMatrix(iu, 7) = "Pending"
    ElseIf MSHFlexGrid1.TextMatrix(iu, 7) = "A" Then
       MSHFlexGrid1.TextMatrix(iu, 7) = "Approved"
    ElseIf MSHFlexGrid1.TextMatrix(iu, 7) = "R" Then
       MSHFlexGrid1.TextMatrix(iu, 7) = "Rejected"
    End If
    TmpRs.MoveNext
Next

End Sub

Private Sub DeleteCmd_Click()
If LCase(TxtStatus) = "approved" Or LCase(TxtStatus) = "rejected" Then
   MsgBox "TDS Adjustment " & TxtStatus & ". Can't Delete!!!"
   Exit Sub
End If


Dim ans As Variant
On Error GoTo msgError

ans = MsgBox("Do you really want to DELETE this Record??", vbYesNo)
If ans = vbYes Then
   
   Call DeleteTIDetail
   
   RsTxn_CRTDSAdjustment.Delete
   RsTxn_CRTDSAdjustment.Update
   RsTxn_CRTDSAdjustment.Requery
   If RsTxn_CRTDSAdjustment.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsTxn_CRTDSAdjustment.MoveNext
   If RsTxn_CRTDSAdjustment.EOF() Then
      RsTxn_CRTDSAdjustment.MoveLast
   End If
   Call Indata
End If
Exit Sub
msgError:
MsgBox "Child record Present "
RsTxn_CRTDSAdjustment.CancelUpdate

End Sub

Private Sub EditCmd_Click()

If LCase(TxtStatus) = "approved" Or LCase(TxtStatus) = "rejected" Then
   MsgBox "TDS Adjustment " & TxtStatus & ". Can't edit!!!"
   Exit Sub
End If

Edit_Flg = "E"

Call GButtonLock(Me, False)
CmbLocationID.Text = TxtLocationID
CmbExchangeTypeID = TxtExchangeTypeID
CmbStatus = TxtStatus
TxtStatus.Visible = False
CmbStatus.Visible = True
TxtRemark.Locked = False

Call AppendTIDetails

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

If LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "fag" And LCase(Gen_UserRole) <> "fagadmin" And LCase(Gen_UserRole) <> "super admin" And LCase(Gen_UserRole) <> "admin" And LCase(Gen_UserRole) <> "headcmg" Then
    lblNAVCode.Visible = False
    ChkNAVCode.Visible = False
End If
Frame1.Enabled = True
SSTab1.Enabled = True
Call Grid_Head2
'CmdPrint.Enabled = True
'CmdPrintPreview.Enabled = True
wc = ""
If Gen_WcLocation <> "" Then
   wc = "Where " & Gen_WcLocation
End If
Call Grid_Head
SFrom.Value = Date - 730
STo.Value = Date
CmbSLocation.Text = ""
TxtSCRID = ""
CmbSStatus.Text = ""

Call TableTxn_CRTDSAdjustment(" " & wc & "  Order by SM_Prefix,TDSAID")
If RsTxn_CRTDSAdjustment.RecordCount = 0 Then
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

Private Sub MSHFlexGrid1_Click()
Dim I As Variant
I = MSHFlexGrid1.RowSel
If I <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(I, 0) = "" Then Exit Sub
RsTxn_CRTDSAdjustment.MoveFirst
RsTxn_CRTDSAdjustment.Find "G_UID= '" & MSHFlexGrid1.TextMatrix(I, 9) & "'"
If Not RsTxn_CRTDSAdjustment.EOF Then
   Call Indata
   Frame0.Visible = False
End If

End Sub

Private Sub MSHFlexGrid3_Click()
If SaveCmd.Enabled = False Then Exit Sub
r = MSHFlexGrid3.RowSel
If MSHFlexGrid3.TextMatrix(r, 0) = "" Then Exit Sub
If r <= 1 Then Exit Sub

TxtTINo.Text = MSHFlexGrid3.TextMatrix(r, 0)
TxtTIDate.Text = MSHFlexGrid3.TextMatrix(r, 1)
TxtBalAmount = MSHFlexGrid3.TextMatrix(r, 6)
TxtTDSAmt = MSHFlexGrid3.TextMatrix(r, 7)
TxtOtherDed = MSHFlexGrid3.TextMatrix(r, 8)
TxtBalRec = MSHFlexGrid3.TextMatrix(r, 9)
TxtTIAmount = MSHFlexGrid3.TextMatrix(r, 2)
TxtTISTAmount = MSHFlexGrid3.TextMatrix(r, 15)
TxtTILocation = MSHFlexGrid3.TextMatrix(r, 17)
CmbExpenseTypeID.Text = MSHFlexGrid3.TextMatrix(r, 10)
mExpenseTypeID = Val(MSHFlexGrid3.TextMatrix(r, 11))
TxtTDSAmt.SetFocus

End Sub

Private Sub SaveCmd_Click()

If CmbLocationID.Text = "" Then
   MsgBox "Blank Location not allowed!"
   CmbLocationID.SetFocus
   Exit Sub
End If

If CmbExchangeTypeID.Text = "" Then
   MsgBox "Blank Exchange Type not allowed!"
   CmbExchangeTypeID.SetFocus
   Exit Sub
End If

If Check3.Value = 0 Then
   If CmbClientDP.Visible = True Then
      MsgBox "Invalid selection !!! "
      Exit Sub
   End If
   If CmbCNonNcdexClient.Visible = True Then
      If CmbCNonNcdexClient.Text = "" Then
         MsgBox "Invalid selection !!! "
         CmbCNonNcdexClient.SetFocus
         Exit Sub
      End If
   End If
   If TxtClientID = "" Then
      MsgBox "Invalid Client ID !!! "
      Exit Sub
   End If
Else
   If CmbDPName.Visible = True Then
      If CmbDPName.Text = "" Then
         MsgBox "Blank DP Name Not allowed !!! "
         CmbDPName.SetFocus
         Exit Sub
      End If
   End If
End If

If TxtRemark = "" Then
   MsgBox " Blank Remark not allowed!!! "
   TxtRemark.SetFocus
   Exit Sub
End If
   
If Val(TxtAmount.Text) = 0 Then
   MsgBox " No grid updated from grid!!! "
   Exit Sub
End If
   
If CheckSTAmount = False Then
   MsgBox "Select invoices either Tax or without tax!!"
   Exit Sub
End If
   
If TxtAdjDate.Value > Date Then
   MsgBox "Future date not allowed!!!"
   TxtAdjDate.SetFocus
   Exit Sub
End If
   
   
If CheckCalSTAmount = False Then
   MsgBox "Issue with tax amount calculation for some invoice(s). Please check by updating invoices in grid!!!"
   Exit Sub
End If
  
   
If LCase(CmbStatus.Text) <> "pending" Then
   
   If TxtFAGRemarks = "" Then
      MsgBox " Blank FAG Remarks not allowed!!! "
      TxtFAGRemarks.SetFocus
      Exit Sub
   End If
   
   If TxtFAGPostingDate.Value > Date Then
      MsgBox "Future date not allowed!!!"
      TxtFAGPostingDate.SetFocus
      Exit Sub
   End If
      
   If TxtFAGPostingDate.Value < TxtAdjDate.Value Then
      MsgBox "'FAG Posting Date' must be greater than 'Adj Date'"
      TxtFAGPostingDate.SetFocus
      Exit Sub
   End If

End If

If Edit_Flg = "A" Then
   RsTxn_CRTDSAdjustment.AddNew
   RsTxn_CRTDSAdjustment!TDSAID = GetMaxTDSAID
   RsTxn_CRTDSAdjustment!G_UID = GetGUID
   TxtAdjID.Text = RsTxn_CRTDSAdjustment!TDSAID
   TxtG_UID = RsTxn_CRTDSAdjustment!G_UID
Else
   RsTxn_CRTDSAdjustment.MoveFirst
   RsTxn_CRTDSAdjustment.Find "G_UID= '" & TxtG_UID & "'"
End If

RsTxn_CRTDSAdjustment!AdjDate = Format(TxtAdjDate.Value, Gen_DatFormat)
RsTxn_CRTDSAdjustment!LocationID = mLocationID
RsTxn_CRTDSAdjustment!SM_Prefix = TxtPreFix
RsTxn_CRTDSAdjustment!ExchangeTypeID = mExchangeTypeID
RsTxn_CRTDSAdjustment!ShowDP = Check3.Value
If Check3.Value = 1 Then
   RsTxn_CRTDSAdjustment!DPID = mDPID
Else
   RsTxn_CRTDSAdjustment!ClientIDRow = mClientIDRow
End If
RsTxn_CRTDSAdjustment!BatchID = 0
'RsTxn_CRTDSAdjustment!Amount = Format(Val(TxtAmount.Text), "#####0.00")

RsTxn_CRTDSAdjustment!Flag = Left(CmbStatus.Text, 1)
RsTxn_CRTDSAdjustment!Remark = TxtRemark.Text

If LCase(CmbStatus) <> "pending" Then
   RsTxn_CRTDSAdjustment!FAGRemarks = TxtFAGRemarks.Text
   RsTxn_CRTDSAdjustment!FAGPostingDate = TxtFAGPostingDate.Value
Else
   RsTxn_CRTDSAdjustment!FAGRemarks = Null
   RsTxn_CRTDSAdjustment!FAGPostingDate = Null
End If

If IsNull(RsTxn_CRTDSAdjustment!CreatedBy) = True Or RsTxn_CRTDSAdjustment!CreatedBy = "" Then
   RsTxn_CRTDSAdjustment!CreatedBy = Gen_UserLoginID
   RsTxn_CRTDSAdjustment!CreatedDate = Now
Else
   RsTxn_CRTDSAdjustment!UpdatedBy = Gen_UserLoginID
   RsTxn_CRTDSAdjustment!UpdatedDate = Now
End If

RsTxn_CRTDSAdjustment.Update

Call SaveTIDetail

RsTxn_CRTDSAdjustment.Requery
RsTxn_CRTDSAdjustment.MoveFirst
RsTxn_CRTDSAdjustment.Find "G_UID= '" & TxtG_UID & "'"
If RsTxn_CRTDSAdjustment.EOF Then
   RsTxn_CRTDSAdjustment.MoveLast
End If
Call GButtonLock(Me, True)
Call Indata


End Sub

Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   If RsTxn_CRTDSAdjustment.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   RsTxn_CRTDSAdjustment.MoveFirst
   RsTxn_CRTDSAdjustment.Find "G_UID= '" & TxtG_UID & "'"
   If RsTxn_CRTDSAdjustment.EOF Then
      RsTxn_CRTDSAdjustment.MoveLast
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
RsTxn_CRTDSAdjustment.Filter = ""
RsTxn_CRTDSAdjustment.MoveFirst

''===========comment on 7 july 2012 for not clearing search grid=============
'MSHFlexGrid1.Clear
'MSHFlexGrid1.Cols = 11
'Call Grid_Head
'SFrom.Value = Date - 730
'STo.Value = Date
'CmbSLocation.Text = ""
'TxtSCRID = ""
'CmbSStatus.Text = ""
''===========comment on 7 july 2012 for not clearing search grid=============

End Sub

Private Sub TxtAdjDate_LostFocus()

If Format(Gen_BlockingDate, "yyyy-mm") >= Format(TxtAdjDate.Value, "yyyy-mm") Then
   MsgBox "Cash Receipt TDS Adjsutment Date allowed for next month of " & Format(Gen_BlockingDate, "MMM-yyyy") & "!!! "
   TxtAdjDate.SetFocus
   Exit Sub
End If

End Sub

Private Sub TxtClientID_LostFocus()
CmbClientDP.Visible = False
If TxtClientID = "" Then Exit Sub
If CmbExchangeTypeID.Text = "" Then Exit Sub
If Edit_Flg = "E" Then Exit Sub
If SaveCmd.Enabled = False Then Exit Sub
Dim ans As Variant
If mSearchbyNAV = True Then
   ans = GetClientNameDtlFrmCRowID(mSearchClientIDRow)
Else
   ans = GetClientNameDtl(TxtClientID, mExchangeTypeID)
End If

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

'mSearchClientIDRow

If mSearchbyNAV = False Then
   tmp = "SELECT Mst_Depository.DepositoryID, Mst_Depository.DepositoryName, Mst_Client.ClientID, " & _
         "Mst_Client.ClientName,Mst_Client.ClientIDRow,Mst_Client.NAVCode,Mst_Client.RNAVCode " & _
         "FROM Mst_Depository INNER JOIN Mst_Client ON Mst_Depository.DepositoryID = Mst_Client.DepositoryID " & _
         "WHERE Mst_Depository.ExchangeTypeID = " & mExchangeTypeID & _
         "  And  (((Mst_Client.ExchangeTypeID)=" & mExchangeTypeID & ") AND ((Mst_Client.ClientID)='" & TxtClientID & "'));"
Else
   tmp = "SELECT Mst_Depository.DepositoryID, Mst_Depository.DepositoryName, Mst_Client.ClientID, " & _
         "Mst_Client.ClientName,Mst_Client.ClientIDRow,Mst_Client.NAVCode,Mst_Client.RNAVCode " & _
         "FROM Mst_Depository INNER JOIN Mst_Client ON Mst_Depository.DepositoryID = Mst_Client.DepositoryID " & _
         "WHERE Mst_Depository.ExchangeTypeID = " & mExchangeTypeID & _
         "  And  (((Mst_Client.ExchangeTypeID)=" & mExchangeTypeID & ") AND ((Mst_Client.ClientIDRow)=" & mSearchClientIDRow & "));"
End If

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

Private Sub CmbCNonNcdexClient_GotFocus()

tmp = CmbCNonNcdexClient.Text
Call TableMst_Client("where ExchangeTypeId = " & mExchangeTypeID)
CmbCNonNcdexClient.Clear
If RsMst_Client.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Client.RecordCount
    CmbCNonNcdexClient.AddItem RsMst_Client!ClientName
    RsMst_Client.MoveNext
Next
CmbCNonNcdexClient.Text = tmp
End Sub

Private Sub CmbCNonNcdexClient_LostFocus()
If CmbCNonNcdexClient.Text = "" Then
   Exit Sub
End If
RsMst_Client.MoveFirst
RsMst_Client.Find "ClientName = '" & CmbCNonNcdexClient.Text & "'"
If RsMst_Client.EOF Then
   MsgBox "Invalid selection "
   CmbCNonNcdexClient.SetFocus
   Exit Sub
End If
mClientIDRow = RsMst_Client!ClientIDRow
mClientNAVCode = IIf(IsNull(RsMst_Client!NavCode), "", RsMst_Client!NavCode)
mClientRNAVCode = IIf(IsNull(RsMst_Client!RNAVCode), "", RsMst_Client!RNAVCode)

Dim ans As Variant
ans = GetClientNameByRow(RsMst_Client!ClientIDRow)
TxtClientID = ans(0)
TxtCClientName = ans(1)
TxtCDPID = ans(2)
TxtCDPName = ans(3)


tmp = "Select distinct GSTNavCode from Mst_ClientStateWiseGST Where ClientIDRow = '" & mClientIDRow & "' And isnull(GSTNavCode,'') <> '' "
Call TmpTable

If TmpRs.RecordCount > 1 Then
   If mClientNAVCode = "" Then
      For x = 1 To TmpRs.RecordCount
          If x = 1 Then
             mClientNAVCode = "('" & TmpRs!GSTNavCode
          ElseIf x = TmpRs.RecordCount Then
             mClientNAVCode = mClientNAVCode & "','" & TmpRs!GSTNavCode & "')"
          Else
             mClientNAVCode = mClientNAVCode & "','" & TmpRs!GSTNavCode
          End If
          TmpRs.MoveNext
      Next
   Else
      For x = 1 To TmpRs.RecordCount
          If x = 1 Then
             mClientNAVCode = "('" & mClientNAVCode & "','" & TmpRs!GSTNavCode
          ElseIf x = TmpRs.RecordCount Then
             mClientNAVCode = mClientNAVCode & "','" & TmpRs!GSTNavCode & "')"
          Else
             mClientNAVCode = mClientNAVCode & "','" & TmpRs!GSTNavCode
          End If
          TmpRs.MoveNext
      Next
   End If
ElseIf TmpRs.RecordCount = 1 Then
   If mClientNAVCode = "" Then
      mClientNAVCode = "('" & TmpRs!GSTNavCode & "')"
   Else
      mClientNAVCode = "('" & mClientNAVCode & "','" & TmpRs!GSTNavCode & "')"
   End If
Else
   If mClientNAVCode <> "" Then
      mClientNAVCode = "('" & mClientNAVCode & "')"
   End If
End If

tmp = "Select distinct RGSTNavCode as 'GSTNavCode' from Mst_ClientStateWiseGST Where ClientIDRow = '" & mClientIDRow & "' And isnull(RGSTNavCode,'') <> '' "
Call TmpTable

If TmpRs.RecordCount > 1 Then
   If mClientRNAVCode = "" Then
      For x = 1 To TmpRs.RecordCount
          If x = 1 Then
             mClientRNAVCode = "('" & TmpRs!GSTNavCode
          ElseIf x = TmpRs.RecordCount Then
             mClientRNAVCode = mClientRNAVCode & "','" & TmpRs!GSTNavCode & "')"
          Else
             mClientRNAVCode = mClientRNAVCode & "','" & TmpRs!GSTNavCode
          End If
          TmpRs.MoveNext
      Next
   Else
      For x = 1 To TmpRs.RecordCount
          If x = 1 Then
             mClientRNAVCode = "('" & mClientRNAVCode & "','" & TmpRs!GSTNavCode
          ElseIf x = TmpRs.RecordCount Then
             mClientRNAVCode = mClientRNAVCode & "','" & TmpRs!GSTNavCode & "')"
          Else
             mClientRNAVCode = mClientRNAVCode & "','" & TmpRs!GSTNavCode
          End If
          TmpRs.MoveNext
      Next
   End If
ElseIf TmpRs.RecordCount = 1 Then
   If mClientRNAVCode = "" Then
      mClientRNAVCode = "('" & TmpRs!GSTNavCode & "')"
   Else
      mClientRNAVCode = "('" & mClientRNAVCode & "','" & TmpRs!GSTNavCode & "')"
   End If
End If

Call AddTIDetails

End Sub

Private Sub CmbClientDP_LostFocus()
Dim Str_() As String
Str_ = Split(CmbClientDP.Text, "~")
If UBound(Str_) <> 1 Then
   MsgBox "Invalid selection !!!!! "
   If CmbClientDP.Visible = True Then
      CmbClientDP.SetFocus
   Else
      CmbClientDP.SetFocus
   End If
   Exit Sub
End If
TmpRs1.MoveFirst

TmpRs1.Find "DepositoryID  = '" & Str_(0) & "'"
If TmpRs1.EOF Then
   MsgBox "Invalid selection !!!!! "
   CmbClientDP.SetFocus
   Exit Sub
End If

TxtCClientName = TmpRs1!ClientName
TxtCDPID = TmpRs1!DepositoryID
TxtCDPName = TmpRs1!DepositoryName
CmbClientDP.Visible = False
mClientIDRow = TmpRs1!ClientIDRow
mClientNAVCode = IIf(IsNull(TmpRs1!NavCode), "", TmpRs1!NavCode)
mClientRNAVCode = IIf(IsNull(TmpRs1!RNAVCode), "", TmpRs1!RNAVCode)

tmp = "Select distinct GSTNavCode from Mst_ClientStateWiseGST Where ClientIDRow = '" & mClientIDRow & "' And isnull(GSTNavCode,'') <> '' "
Call TmpTable

If TmpRs.RecordCount > 1 Then
   If mClientNAVCode = "" Then
      For x = 1 To TmpRs.RecordCount
          If x = 1 Then
             mClientNAVCode = "('" & TmpRs!GSTNavCode
          ElseIf x = TmpRs.RecordCount Then
             mClientNAVCode = mClientNAVCode & "','" & TmpRs!GSTNavCode & "')"
          Else
             mClientNAVCode = mClientNAVCode & "','" & TmpRs!GSTNavCode
          End If
          TmpRs.MoveNext
      Next
   Else
      For x = 1 To TmpRs.RecordCount
          If x = 1 Then
             mClientNAVCode = "('" & mClientNAVCode & "','" & TmpRs!GSTNavCode
          ElseIf x = TmpRs.RecordCount Then
             mClientNAVCode = mClientNAVCode & "','" & TmpRs!GSTNavCode & "')"
          Else
             mClientNAVCode = mClientNAVCode & "','" & TmpRs!GSTNavCode
          End If
          TmpRs.MoveNext
      Next
   End If
ElseIf TmpRs.RecordCount = 1 Then
   If mClientNAVCode = "" Then
      mClientNAVCode = "('" & TmpRs!GSTNavCode & "')"
   Else
      mClientNAVCode = "('" & mClientNAVCode & "','" & TmpRs!GSTNavCode & "')"
   End If
End If


tmp = "Select distinct RGSTNavCode as 'GSTNavCode' from Mst_ClientStateWiseGST Where ClientIDRow = '" & mClientIDRow & "' And isnull(RGSTNavCode,'') <> '' "
Call TmpTable

If TmpRs.RecordCount > 1 Then
   If mClientRNAVCode = "" Then
      For x = 1 To TmpRs.RecordCount
          If x = 1 Then
             mClientRNAVCode = "('" & TmpRs!GSTNavCode
          ElseIf x = TmpRs.RecordCount Then
             mClientRNAVCode = mClientRNAVCode & "','" & TmpRs!GSTNavCode & "')"
          Else
             mClientRNAVCode = mClientRNAVCode & "','" & TmpRs!GSTNavCode
          End If
          TmpRs.MoveNext
      Next
   Else
      For x = 1 To TmpRs.RecordCount
          If x = 1 Then
             mClientRNAVCode = "('" & mClientRNAVCode & "','" & TmpRs!GSTNavCode
          ElseIf x = TmpRs.RecordCount Then
             mClientRNAVCode = mClientRNAVCode & "','" & TmpRs!GSTNavCode & "')"
          Else
             mClientRNAVCode = mClientRNAVCode & "','" & TmpRs!GSTNavCode
          End If
          TmpRs.MoveNext
      Next
   End If
ElseIf TmpRs.RecordCount = 1 Then
   If mClientRNAVCode = "" Then
      mClientRNAVCode = "('" & TmpRs!GSTNavCode & "')"
   Else
      mClientRNAVCode = "('" & mClientRNAVCode & "','" & TmpRs!GSTNavCode & "')"
   End If
End If

Call AddTIDetails

End Sub

Private Sub CmbDPName_GotFocus()
If CmbExchangeTypeID.Text = "" Then
   MsgBox "Please select Exchange Type!!!"
   CmbDPName.Clear
   CmbExchangeTypeID.SetFocus
   Exit Sub
End If
tmp = CmbDPName.Text

Call TableMst_Depository("Where ExchangeTypeId = " & mExchangeTypeID & " ")

CmbDPName.Clear
If RsMst_Depository.RecordCount = 0 Then
   MsgBox "No record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Depository.RecordCount
    CmbDPName.AddItem RsMst_Depository!DepositoryName & "~" & RsMst_Depository!DepositoryID
    RsMst_Depository.MoveNext
Next
CmbDPName.Text = tmp

End Sub

Private Sub CmbDPName_LostFocus()

Dim strSplitString() As String
Dim DPID_, DPName_ As String

If CmbDPName.Text = "" Then
   TxtCDPID = ""
   TxtCDPName = ""
   Exit Sub
End If
strSplitString = Split(CmbDPName, "~")
DPName_ = Trim(strSplitString(0))
DPID_ = Trim(strSplitString(1))

RsMst_Depository.MoveFirst
RsMst_Depository.Find "DepositoryID = '" & DPID_ & "'"
If RsMst_Depository.EOF Then
   MsgBox "Invalid selection "
   CmbDPName.SetFocus
   Exit Sub
End If

mDPID = RsMst_Depository!DepositoryID
mDPNAVCode = IIf(IsNull(RsMst_Depository!NavCode), "", RsMst_Depository!NavCode)

tmp = "select distinct GSTNavCode from Mst_DepositoryStateWiseGST Where DepositoryID = '" & mDPID & "' "
Call TmpTable

If TmpRs.RecordCount > 1 Then
   If mDPNAVCode = "" Then
      For x = 1 To TmpRs.RecordCount
          If x = 1 Then
             mDPNAVCode = "('" & TmpRs!GSTNavCode
          ElseIf x = TmpRs.RecordCount Then
             mDPNAVCode = mDPNAVCode & "','" & TmpRs!GSTNavCode & "')"
          Else
             mDPNAVCode = mDPNAVCode & "','" & TmpRs!GSTNavCode
          End If
          TmpRs.MoveNext
      Next
   Else
      For x = 1 To TmpRs.RecordCount
          If x = 1 Then
             mDPNAVCode = "('" & mDPNAVCode & "','" & TmpRs!GSTNavCode
          ElseIf x = TmpRs.RecordCount Then
             mDPNAVCode = mDPNAVCode & "','" & TmpRs!GSTNavCode & "')"
          Else
             mDPNAVCode = mDPNAVCode & "','" & TmpRs!GSTNavCode
          End If
          TmpRs.MoveNext
      Next
   End If
ElseIf TmpRs.RecordCount = 1 Then
   If mDPNAVCode = "" Then
      mDPNAVCode = "('" & TmpRs!GSTNavCode & "')"
   Else
      mDPNAVCode = "('" & mDPNAVCode & "','" & TmpRs!GSTNavCode & "')"
   End If
End If

TxtCDPID = mDPID
TxtCDPName = DPName_ 'CmbDPName.Text
Call AddTIDetails
End Sub

Private Sub CmbStatus_LostFocus()
If CmbStatus.Text = "" Then
   TxtFAGRemarks.Locked = True
   TxtFAGPostingDate.Visible = False
   Label5.Visible = False
   Exit Sub
End If

If LCase(CmbStatus.Text) = "approved" Then
   mStatus = "A"
   TxtFAGRemarks.Locked = False
   TxtFAGRemarks = Trim(TxtRemark)
   TxtFAGPostingDate.Visible = True
   Label5.Visible = True
   TxtFAGPostingDate.Enabled = True
   TxtFAGPostingDate.Value = Date
ElseIf LCase(CmbStatus.Text) = "pending" Then
   mStatus = "P"
   TxtFAGRemarks.Locked = True
   TxtFAGRemarks = ""
   TxtFAGPostingDate.Visible = False
   Label5.Visible = False
ElseIf LCase(CmbStatus.Text) = "rejected" Then
   mStatus = "R"
   TxtFAGRemarks.Locked = False
   TxtFAGPostingDate.Visible = True
   Label5.Visible = True
   TxtFAGPostingDate.Enabled = True
   TxtFAGPostingDate.Value = Date
Else
   MsgBox "Invalid selection!!!"
   CmbStatus.SetFocus
   TxtFAGRemarks.Locked = True
   TxtFAGRemarks = ""
   TxtFAGPostingDate.Visible = False
   Label5.Visible = False
   Exit Sub
End If

If LCase(Gen_UserRole) <> "fag" And LCase(Gen_UserRole) <> "fagadmin" And LCase(Gen_UserRole) <> "power" Then
   If mStatus = "A" Then
      MsgBox "Please forward this claim to FAG user, Only FAG user have access to approve claim '!!!"
      CmbStatus.SetFocus
      Exit Sub
   End If
End If

End Sub

Private Sub Grid_Head2()

MSHFlexGrid3.Clear
MSHFlexGrid3.Rows = 3
MSHFlexGrid3.Cols = 19
MSHFlexGrid3.FixedRows = 2
MSHFlexGrid3.WordWrap = True

MSHFlexGrid3.TextMatrix(1, 0) = "TI No"
MSHFlexGrid3.TextMatrix(1, 1) = "TI Date"
MSHFlexGrid3.TextMatrix(1, 2) = "TI Amount"
MSHFlexGrid3.TextMatrix(1, 3) = "Received Amt"
MSHFlexGrid3.TextMatrix(1, 4) = "TDS Deducted"
MSHFlexGrid3.TextMatrix(1, 5) = "Other Deduction"
MSHFlexGrid3.TextMatrix(1, 6) = "Balance Amt"
MSHFlexGrid3.TextMatrix(1, 7) = "TDS Deducted against this Adj."
MSHFlexGrid3.TextMatrix(1, 8) = "Other Deduction against this Adj."
MSHFlexGrid3.TextMatrix(1, 9) = "Balance Receivable after this Adj."
MSHFlexGrid3.TextMatrix(1, 10) = "Expense Ac."
MSHFlexGrid3.TextMatrix(1, 11) = "Expense Ac ID"
MSHFlexGrid3.TextMatrix(1, 12) = "TIID"
MSHFlexGrid3.TextMatrix(1, 13) = "PostInvoice No"
MSHFlexGrid3.TextMatrix(1, 14) = "Entry No"
MSHFlexGrid3.TextMatrix(1, 15) = "TI Tax Amount"
MSHFlexGrid3.TextMatrix(1, 16) = "Tax Amount"
MSHFlexGrid3.TextMatrix(1, 17) = "Location"
MSHFlexGrid3.TextMatrix(1, 18) = "NAV Code"

MSHFlexGrid3.ColWidth(0) = 1600
MSHFlexGrid3.ColWidth(1) = 1400
MSHFlexGrid3.ColWidth(2) = 1200
MSHFlexGrid3.ColWidth(3) = 1200
MSHFlexGrid3.ColWidth(4) = 1200
MSHFlexGrid3.ColWidth(5) = 1200
MSHFlexGrid3.ColWidth(6) = 1200
MSHFlexGrid3.ColWidth(7) = 1500
MSHFlexGrid3.ColWidth(8) = 1500
MSHFlexGrid3.ColWidth(9) = 1500
MSHFlexGrid3.ColWidth(10) = 1200
MSHFlexGrid3.ColWidth(11) = 0
MSHFlexGrid3.ColWidth(12) = 0
MSHFlexGrid3.ColWidth(17) = 1800

MSHFlexGrid3.RowHeight(1) = 800

End Sub

Private Sub AddTIDetails()

''tmp = "Select TINo,TIDate,TIAmount,TIID,TI.PostInvoiceNo,TI.EntryNo,(Select Sum(TCRT.RecAmount) as RecAmount From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.CRID= TCRT.CRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'"
''tmp = tmp & " And TIID=TI.TIID Group By  TCRT.TIID)RecAmount,(Select Sum(TCRT.RecTDS) as RecTDS From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.CRID= TCRT.CRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'"
''tmp = tmp & " And TIID=TI.TIID Group By  TCRT.TIID)RecTDS,(Select Sum(TCRT.RecOtherDed) as RecOtherDed From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.CRID= TCRT.CRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'"
''tmp = tmp & " And TIID=TI.TIID Group By  TCRT.TIID)RecOtherDed"
''tmp = tmp & " From  Txn_TIDetails TI"
''tmp = tmp & " Where  TINo<>'' And TIAmount-((select isnull(Sum(RecAmount)+Sum(RecTDS)+Sum(RecOtherDed),0) From Txn_NewCashReceiptTIDetail TCRT"
''tmp = tmp & " Inner Join Txn_NewCashReceipt TCR on TCR.CRID= TCRT.CRID and TCR.SM_Prefix=TCRT.SM_Prefix"
''tmp = tmp & " Where TCRT.TIID=TI.TIID And TCR.ReturnFlag<>'Y')+(Select isnull(Sum(RecTDS)+Sum(RecOtherDed),0)  From Txn_CRTDSAdjustmentDetail TCRT"
''tmp = tmp & " Inner Join Txn_CRTDSAdjustment TCR on TCR.TDSAID = TCRT.TDSAID and TCR.SM_Prefix=TCRT.SM_Prefix"
''tmp = tmp & " Where TCRT.TIID=TI.TIID And TCR.Flag <> 'R' and TCR.G_UID<>'" & TxtG_UID & "' ))>0 "
''
''
''tmp = tmp & " And (LocationID=" & mLocationID & " or LocationID is null)"
''If Check3.Value = 0 Then
''   If mClientRNAVCode <> "" Then
''      tmp = tmp & " And (TI.NAVCode = '" & mClientNAVCode & "' or TI.NAVCode = '" & mClientRNAVCode & "')"
''   Else
''      tmp = tmp & " And TI.NAVCode = '" & mClientNAVCode & "'"
''   End If
''Else
''   tmp = tmp & " And TI.NAVCode='" & mDPNAVCode & "'"
''End If
''
''Call Tmp3Table

Call Grid_Head2

Wc1 = ""

If LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "admin" And LCase(Gen_UserRole) <> "superadmin" And LCase(Gen_UserRole) <> "fag" And LCase(Gen_UserRole) <> "fagadmin" And LCase(Gen_UserRole) <> "headcmg" Then
   Wc1 = " And (TI.LocationID=" & mLocationID & " or TI.LocationID is null)"
End If

If Check3.Value = 0 Then
   If InStr(1, mClientNAVCode, ")") = 0 And InStr(1, mClientRNAVCode, ")") = 0 Then
      If mClientRNAVCode <> "" Then
         Wc1 = Wc1 & " And (TI.NAVCode = '" & mClientNAVCode & "' or TI.NAVCode = '" & mClientRNAVCode & "')"
      Else
         Wc1 = Wc1 & " And TI.NAVCode = '" & mClientNAVCode & "'"
      End If
   Else
      If mClientRNAVCode <> "" Then
         If mClientNAVCode <> "" Then
            Wc1 = Wc1 & " And (TI.NAVCode in " & mClientNAVCode & " or TI.NAVCode in " & mClientRNAVCode & ")"
         Else
            Wc1 = Wc1 & " And (TI.NAVCode in " & mClientRNAVCode & ")"
         End If
      Else
         Wc1 = Wc1 & " And TI.NAVCode in " & mClientNAVCode & ""
      End If
   End If
Else
   If InStr(1, mDPNAVCode, ")") > 0 Then
      Wc1 = Wc1 & " And TI.NAVCode in " & mDPNAVCode & ""
   Else
      Wc1 = Wc1 & " And TI.NAVCode = '" & mDPNAVCode & "' "
   End If
End If


Call GetClientsRemInvoices(Wc1)

''tmp = " Select TCRT.TIID,Sum(RecTDS) as RecTDSAmount,Sum(RecOtherDed)as RecOtherDed  From Txn_CRTDSAdjustmentDetail TCRT"
''tmp = tmp & " Inner Join Txn_CRTDSAdjustment TCR on TCR.TDSAID = TCRT.TDSAID and TCR.SM_Prefix=TCRT.SM_Prefix"
''tmp = tmp & " Where TCR.Flag <> 'R' and TCR.G_UID<>'" & TxtG_UID & "'"
''tmp = tmp & " Group By  TCRT.TIID"
''
''Call Tmp5Table

If TmpRs.RecordCount > 0 Then
   For gr = 1 To TmpRs.RecordCount
       rn = MSHFlexGrid3.Rows - 1

       MSHFlexGrid3.TextMatrix(rn, 0) = TmpRs!TINo
       MSHFlexGrid3.TextMatrix(rn, 1) = IIf(IsNull(TmpRs!TIDate), "", Format(TmpRs!TIDate, "dd-MMM-yyyy"))
       MSHFlexGrid3.TextMatrix(rn, 2) = IIf(IsNull(TmpRs!TIAmount), 0, TmpRs!TIAmount)
       MSHFlexGrid3.TextMatrix(rn, 3) = IIf(IsNull(TmpRs!RecAmount), 0, TmpRs!RecAmount)
       MSHFlexGrid3.TextMatrix(rn, 3) = Val(MSHFlexGrid3.TextMatrix(rn, 3)) + IIf(IsNull(TmpRs!AdjAmountFAG), 0, TmpRs!AdjAmountFAG)
       MSHFlexGrid3.TextMatrix(rn, 4) = IIf(IsNull(TmpRs!RecTDS), 0, TmpRs!RecTDS)
       MSHFlexGrid3.TextMatrix(rn, 5) = IIf(IsNull(TmpRs!RecOtherDed), 0, TmpRs!RecOtherDed)

''       If TmpRs5.RecordCount > 0 Then TmpRs5.MoveFirst
''       TmpRs5.Find "TIID=" & TmpRs3!TIID
''       If TmpRs5.EOF = False Then
''          MSHFlexGrid3.TextMatrix(rn, 4) = Val(MSHFlexGrid3.TextMatrix(rn, 4)) + IIf(IsNull(TmpRs5!RecTDSAmount), 0, TmpRs5!RecTDSAmount)
''           MSHFlexGrid3.TextMatrix(rn, 5) = Val(MSHFlexGrid3.TextMatrix(rn, 5)) + IIf(IsNull(TmpRs5!RecOtherDed), 0, TmpRs5!RecOtherDed)
''       End If

       MSHFlexGrid3.TextMatrix(rn, 6) = Format(Val(MSHFlexGrid3.TextMatrix(rn, 2)) - (Val(MSHFlexGrid3.TextMatrix(rn, 3)) + Val(MSHFlexGrid3.TextMatrix(rn, 4)) + Val(MSHFlexGrid3.TextMatrix(rn, 5))), "#####0.00")
       
       MSHFlexGrid3.TextMatrix(rn, 12) = TmpRs!TIID
       MSHFlexGrid3.TextMatrix(rn, 13) = IIf(IsNull(TmpRs!PostInvoiceNo), 0, TmpRs!PostInvoiceNo)
       MSHFlexGrid3.TextMatrix(rn, 14) = IIf(IsNull(TmpRs!EntryNo), 0, TmpRs!EntryNo)
       MSHFlexGrid3.TextMatrix(rn, 15) = IIf(IsNull(TmpRs!TISTAmount), 0, TmpRs!TISTAmount)
       MSHFlexGrid3.TextMatrix(rn, 17) = IIf(IsNull(TmpRs!Locationname), "", TmpRs!Locationname)
       MSHFlexGrid3.TextMatrix(rn, 18) = IIf(IsNull(TmpRs!NavCode), "", TmpRs!NavCode)
       TmpRs.MoveNext
       MSHFlexGrid3.Rows = MSHFlexGrid3.Rows + 1
   Next
Else
   MsgBox "TI Details not found!!"
End If
End Sub

Private Sub TxtFAGPostingDate_LostFocus()

If Format(Gen_BlockingDate, "yyyy-mm") >= Format(TxtFAGPostingDate.Value, "yyyy-mm") Then
   MsgBox "Date allowed for next month of " & Format(Gen_BlockingDate, "MMM-yyyy") & "!!! "
   TxtFAGPostingDate.SetFocus
   Exit Sub
End If


End Sub

Private Sub TxtNAVCode_LostFocus()
If TxtNAVCode.Text = "" Then Exit Sub
mSearchbyNAV = True
If Check3.Value = vbChecked Then
    tmp = "Select * from Mst_Depository Where NAVCode = '" & TxtNAVCode & "' And ExchangeTypeID = " & mExchangeTypeID
    Call TmpTable
    If TmpRs.RecordCount > 0 Then
        CmbDPName.Text = TmpRs!DepositoryName & "~" & TmpRs!DepositoryID
    Else
        CmbDPName.Text = ""
    End If
    Call CmbDPName_GotFocus
    Call CmbDPName_LostFocus
Else
    tmp = "Select * from Mst_Client Where (NAVCode = '" & TxtNAVCode & "' Or RNAVCode = '" & TxtNAVCode & "') And ExchangeTypeID = " & mExchangeTypeID
    Call TmpTable
    If TmpRs.RecordCount > 0 Then
        If mExchangeTypeID = 1 Then
            CmbCNonNcdexClient.Text = TmpRs!ClientName
            Call CmbCNonNcdexClient_GotFocus
            Call CmbCNonNcdexClient_LostFocus
        Else
            TxtClientID.Text = TmpRs!ClientID
            mSearchClientIDRow = TmpRs!ClientIDRow
            Call TxtClientID_LostFocus
        End If
    Else
        CmbCNonNcdexClient.Text = ""
        TxtClientID.Text = ""
        Call TxtClientID_LostFocus
    End If
End If
End Sub

Private Sub TxtTDSAmt_LostFocus()
If TxtTDSAmt.Text = "" Then
   TxtTDSAmt = 0
   Exit Sub
End If
If IsNumeric(TxtTDSAmt.Text) = False Or Val(TxtTDSAmt.Text) < 0 Then
   MsgBox "Invalid Numeric Format !!!"
   TxtTDSAmt.SetFocus
   Exit Sub
End If
TxtTDSAmt = Round(Val(TxtTDSAmt), 2)
'TxtBalRec.Text = Format(Val(TxtBalAmount) - (Val(TxtRecAmt) + Val(TxtTDSAmt)), "#####0.00")
TxtBalRec.Text = Format(Val(TxtBalAmount) - (Val(TxtTDSAmt) + Val(TxtOtherDed)), "#####0.00")
If Val(TxtBalRec.Text) < 0 Then
   MsgBox "Balance Receivable cannot be less than zero."
   TxtTDSAmt.Text = 0
   TxtTDSAmt.SetFocus
End If
End Sub

Sub ClearTITabControl()

TxtTINo = ""
TxtTIDate = ""
TxtBalAmount = ""
TxtTDSAmt = ""
TxtBalRec = ""
TxtOtherDed = ""
CmbExpenseTypeID.Text = ""
TxtTIAmount = ""
TxtTISTAmount = ""
TxtTILocation = ""
TxtSearchBy = ""

End Sub

Private Sub TxtOtherDed_LostFocus()
If TxtOtherDed.Text = "" Then
   TxtOtherDed = 0
   Exit Sub
End If
If IsNumeric(TxtOtherDed.Text) = False Or Val(TxtOtherDed.Text) < 0 Then
   MsgBox "Invalid Numeric Format !!!"
   TxtOtherDed.SetFocus
   Exit Sub
End If
TxtOtherDed = Round(Val(TxtOtherDed), 2)
TxtBalRec.Text = Format(Val(TxtBalAmount) - (Val(TxtTDSAmt) + Val(TxtOtherDed)), "#####0.00")
If Val(TxtBalRec.Text) < 0 Then
   MsgBox "Balance Receivable cannot be less than zero."
   TxtOtherDed.Text = 0
   TxtOtherDed.SetFocus
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

Private Sub CmbExpenseTypeID_GotFocus()
Dim wc As String
tmp = CmbExpenseTypeID.Text
Call TableMst_ExpenseType(" Where (AllocateTo = 'B' Or AllocateTo = 'E') And Flag = 'A' And OtherDeductionFlag = 'Y' ")
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
   mExpenseTypeID = 0
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


Private Function GetMaxTDSAID() As Double
Dim Retval As Double
tmp = "Select max(TDSAID) from Txn_CRTDSAdjustment Where SM_Prefix='" & TxtPreFix & "'"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetMaxTDSAID = Retval
End Function

Private Sub DeleteTIDetail()
tmp = "Select TIID,RecTDs,RecOtherDed From Txn_CRTDSAdjustmentDetail Where TDSAID= " & Val(TxtAdjID) & " and SM_Prefix='" & TxtPreFix & "'"
Call TmpTable
Call TableTxn_TIDetails(" where TIID in (Select TIID From Txn_CRTDSAdjustmentDetail Where TDSAID= " & Val(TxtAdjID) & "and SM_Prefix='" & TxtPreFix & "')")
For I = 1 To TmpRs.RecordCount
    If RsTxn_TIDetails.RecordCount > 0 Then RsTxn_TIDetails.MoveFirst
    RsTxn_TIDetails.Find "TIID=" & TmpRs!TIID
    If RsTxn_TIDetails.EOF = False Then
       RsTxn_TIDetails!RecTDSAmount = IIf(IsNull(RsTxn_TIDetails!RecTDSAmount), 0, RsTxn_TIDetails!RecTDSAmount) - IIf(IsNull(TmpRs!RecTDS), 0, TmpRs!RecTDS)
       RsTxn_TIDetails!RecOtherDed = IIf(IsNull(RsTxn_TIDetails!RecOtherDed), 0, RsTxn_TIDetails!RecOtherDed) - IIf(IsNull(TmpRs!RecOtherDed), 0, TmpRs!RecOtherDed)
       RsTxn_TIDetails.Update
    End If
    TmpRs.MoveNext
Next
tmp = "Delete From Txn_CRTDSAdjustmentDetail Where TDSAID=" & Val(TxtAdjID) & " And SM_Prefix='" & TxtPreFix & "'"
Call Tmp4Table
End Sub

Private Sub SaveTIDetail()
Dim Wc1 As String
Wc1 = ""
Call DeleteTIDetail
Call TableTxn_CRTDSAdjustmentDetail(" Where TDSAID= " & Val(TxtAdjID))
If Check3.Value = 0 Then
   If InStr(1, mClientNAVCode, ")") = 0 And InStr(1, mClientRNAVCode, ")") = 0 Then
      If mClientRNAVCode <> "" Then
         Wc1 = Wc1 & " Where (NAVCode = '" & mClientNAVCode & "' or NAVCode = '" & mClientRNAVCode & "')"
      Else
         Wc1 = Wc1 & " Where NAVCode = '" & mClientNAVCode & "'"
      End If
   Else
      If mClientRNAVCode <> "" Then
         If mClientNAVCode <> "" Then
'            Wc1 = Wc1 & " Where (NAVCode in " & mClientNAVCode & " or NAVCode in " & mClientRNAVCode & ")"
         
            If InStr(1, mClientNAVCode, ")") = 0 Then
               Wc1 = Wc1 & " Where (NAVCode = '" & mClientNAVCode & "' or NAVCode in " & mClientRNAVCode & ")"
            Else
               Wc1 = Wc1 & " Where (NAVCode in " & mClientNAVCode & " or NAVCode in " & mClientRNAVCode & ")"
            End If
         Else
            Wc1 = Wc1 & " Where (NAVCode in " & mClientRNAVCode & ")"
         End If
      Else
         Wc1 = Wc1 & " Where NAVCode in " & mClientNAVCode & ""
      End If
   End If
Else
   If InStr(1, mDPNAVCode, ")") > 0 Then
      Wc1 = Wc1 & " Where NAVCode in " & mDPNAVCode & ""
   Else
      Wc1 = Wc1 & " Where NAVCode = '" & mDPNAVCode & "' "
   End If
End If

Call TableTxn_TIDetails(Wc1)

For I = 2 To MSHFlexGrid3.Rows - 1
    If MSHFlexGrid3.TextMatrix(I, 0) <> "" And (Val(MSHFlexGrid3.TextMatrix(I, 7)) > 0 Or Val(MSHFlexGrid3.TextMatrix(I, 8)) > 0) Then
       RsTxn_CRTDSAdjustmentDetail.AddNew
       RsTxn_CRTDSAdjustmentDetail!TDSADID = GetMaxTDSADID
       RsTxn_CRTDSAdjustmentDetail!TDSAID = Val(TxtAdjID)
       RsTxn_CRTDSAdjustmentDetail!SM_Prefix = TxtPreFix
       RsTxn_CRTDSAdjustmentDetail!G_UID = GetGUID
       RsTxn_CRTDSAdjustmentDetail!TIID = Val(MSHFlexGrid3.TextMatrix(I, 12))
       If Val(MSHFlexGrid3.TextMatrix(I, 11)) > 0 Then
          RsTxn_CRTDSAdjustmentDetail!ExpenseAccID = Val(MSHFlexGrid3.TextMatrix(I, 11))
       End If
       RsTxn_CRTDSAdjustmentDetail!RecTDS = Val(MSHFlexGrid3.TextMatrix(I, 7))
       RsTxn_CRTDSAdjustmentDetail!RecOtherDed = Val(MSHFlexGrid3.TextMatrix(I, 8))
       
       'Added on 28th Dec 2011 Parag
       RsTxn_CRTDSAdjustmentDetail!STAmount = Val(MSHFlexGrid3.TextMatrix(I, 16))
       
       RsTxn_CRTDSAdjustmentDetail.Update
       
       RsTxn_CRTDSAdjustmentDetail.Requery
       If RsTxn_TIDetails.RecordCount > 0 Then RsTxn_TIDetails.MoveFirst
       RsTxn_TIDetails.Find "TIID=" & Val(MSHFlexGrid3.TextMatrix(I, 12))
       If RsTxn_TIDetails.EOF = False Then
          RsTxn_TIDetails!RecTDSAmount = IIf(IsNull(RsTxn_TIDetails!RecTDSAmount), 0, RsTxn_TIDetails!RecTDSAmount) + Val(MSHFlexGrid3.TextMatrix(I, 7))
          RsTxn_TIDetails!RecOtherDed = IIf(IsNull(RsTxn_TIDetails!RecOtherDed), 0, RsTxn_TIDetails!RecOtherDed) + Val(MSHFlexGrid3.TextMatrix(I, 8))
          RsTxn_TIDetails.Update
       End If
    End If
Next
End Sub

Private Function GetMaxTDSADID() As Double
Dim Retval As Double
tmp = "Select max(TDSADID) from Txn_CRTDSAdjustmentDetail"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetMaxTDSADID = Retval
End Function

Private Sub Indata()

Call GButtonLock(Me, True)

Call ClearTITabControl
Check3.Enabled = False

ChkNAVCode.Value = 0
ChkNAVCode.Enabled = False

tmp = " Select TTI.TDSAID,TTI.AdjDate,TTI.ExchangeTypeID,TTI.LocationID,TTI.SM_Prefix,TTI.ClientIDRow,TTI.DPID,TTI.ShowDP,"
tmp = tmp & " TTI.Flag,TTI.BatchID,TTI.Remark,TTI.FAGRemarks,TTI.FAGPostingDate,"
tmp = tmp & " Me.ExchangeType, ML.LocationName, MC.ClientName, MD.DepositoryName, TTI.G_UID"
tmp = tmp & " From Txn_CRTDSAdjustment TTI"
tmp = tmp & " Inner Join Mst_ExchangeType ME On TTI.ExchangeTypeID = ME.ExchangeTypeID"
tmp = tmp & " left Join Mst_Client MC On TTI.ClientIDRow = MC.ClientIDRow"
tmp = tmp & " left Join Mst_Depository MD On TTI.DPID = MD.DepositoryID"
tmp = tmp & " Inner Join Mst_Location ML On TTI.LocationID = ML.LocationID"
tmp = tmp & " Where TTI.TDSAID = " & RsTxn_CRTDSAdjustment!TDSAID & " And TTI.SM_Prefix='" & RsTxn_CRTDSAdjustment!SM_Prefix & "'"

Call TmpTable
mSearchbyNAV = False
TxtAdjID = TmpRs!TDSAID
TxtAdjDate.Value = TmpRs!AdjDate
TxtAdjDate.Enabled = False

CmbExchangeTypeID.Visible = False
TxtExchangeTypeID.Visible = True
TxtExchangeTypeID.Text = TmpRs!ExchangeType
mExchangeTypeID = TmpRs!ExchangeTypeID

CmbLocationID.Visible = False
TxtLocationID.Visible = True
TxtLocationID.Text = TmpRs!Locationname
CmdLocationID.Visible = False
mLocationID = TmpRs!LocationID
TxtPreFix = TmpRs!SM_Prefix
CmbStatus.Visible = False
TxtStatus.Visible = True
mStatusFlag = TmpRs!Flag

If LCase(mStatusFlag) = "p" Then
   TxtStatus.Text = "Pending"
ElseIf LCase(mStatusFlag) = "a" Then
   TxtStatus.Text = "Approved"
ElseIf LCase(mStatusFlag) = "r" Then
   TxtStatus.Text = "Rejected"
End If
'CmdClient.Visible = False
TxtClientID.Locked = True
Check3.Value = IIf(IsNull(TmpRs!ShowDP), 0, TmpRs!ShowDP)
If TmpRs!ShowDP = 0 Then
   mClientIDRow = TmpRs!ClientIDRow
   ans = GetClientNameByRow(Val(mClientIDRow))
   TxtClientID = ans(0)
   TxtCClientName = ans(1)  'TmpRs!ClientName
   TxtCDPID = ans(2)  'TmpRs!DepositoryID
   TxtCDPName = ans(3) 'TmpRs!DepositoryName

   CmbClientDP.Text = ""
   CmbClientDP.Visible = False
   CmbCNonNcdexClient.Visible = False
Else
   CmbDPName.Text = TmpRs!DepositoryName & "~" & TmpRs!DPID
   mDPID = TmpRs!DPID
   CmbDPName.Enabled = False
End If

TxtAmount.Locked = True
TxtRemark.Locked = True
TxtFAGRemarks.Locked = True

TxtRemark = TmpRs!Remark
TxtFAGRemarks = IIf(IsNull(TmpRs!FAGRemarks), "", TmpRs!FAGRemarks)
If IsNull(TmpRs!FAGPostingDate) = False Then
   TxtFAGPostingDate.Value = TmpRs!FAGPostingDate
   TxtFAGPostingDate.Enabled = False
Else
   TxtFAGPostingDate.Visible = False
   Label5.Visible = False
End If

TxtBatchNo = IIf(IsNull(TmpRs!BatchID), 0, TmpRs!BatchID)
TxtG_UID = TmpRs!G_UID
TxtCreated = IIf(IsNull(RsTxn_CRTDSAdjustment!CreatedBy), "", RsTxn_CRTDSAdjustment!CreatedBy) & " :- " & IIf(IsNull(RsTxn_CRTDSAdjustment!CreatedDate), "", Format(RsTxn_CRTDSAdjustment!CreatedDate, "dd-mmm-yyyy hh:mm:ss"))
TxtUpdated = IIf(IsNull(RsTxn_CRTDSAdjustment!UpdatedBy), "", RsTxn_CRTDSAdjustment!UpdatedBy) & " :- " & IIf(IsNull(RsTxn_CRTDSAdjustment!UpdatedDate), "", Format(RsTxn_CRTDSAdjustment!UpdatedDate, "dd-mmm-yyyy hh:mm:ss"))

Call InDataTIDetails

If RsTxn_CRTDSAdjustment.RecordCount = 1 Then
   Call GButtonLock_1(Me, True)
Else
   Call GButtonLock(Me, True)
End If

End Sub
Private Sub InDataTIDetails()
Dim ans As Variant
Call Grid_Head2

tmp = " Select TTI.TINo,TTI.TIDate,TTI.TIAmount,TTI.TISTAmount, "
tmp = tmp & " TI.RecTDS,TI.RecOtherDed,TI.TIID,TTI.PostInvoiceNo,TTI.EntryNo,TI.ExpenseAccID,ME.ExpenseType,TI.STAmount,ML.LocationName,TTI.NAVCode "
tmp = tmp & " From Txn_CRTDSAdjustmentDetail TI"
tmp = tmp & " inner Join  Txn_TIDetails TTI on TI.TIID=TTI.TIID"
tmp = tmp & " left Join  Mst_Location ML on ML.LocationID=TTI.LocationID"
tmp = tmp & " left Join  Mst_ExpenseType ME on TI.ExpenseAccID=ME.ExpenseTypeID"
tmp = tmp & " Where TI.TDSAID = " & Val(TxtAdjID) & " And TI.SM_Prefix='" & TxtPreFix & "'"

Call Tmp4Table

Wc1 = ""
Wc1 = " And TI.TIID in (Select TIID from Txn_CRTDSAdjustmentDetail TCRT"
Wc1 = Wc1 & " Where TCRT.TDSAID =  " & Val(TxtAdjID) & " and TCRT.SM_Prefix= '" & TxtPreFix.Text & "')"

Call GetClientsRemInvoices(Wc1, , TxtG_UID, , "show")

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
    MSHFlexGrid3.TextMatrix(I, 7) = TmpRs4!RecTDS
    MSHFlexGrid3.TextMatrix(I, 8) = IIf(IsNull(TmpRs4!RecOtherDed), 0, TmpRs4!RecOtherDed)
    MSHFlexGrid3.TextMatrix(I, 9) = Format(Val(MSHFlexGrid3.TextMatrix(I, 6)) - (Val(MSHFlexGrid3.TextMatrix(I, 7)) + Val(MSHFlexGrid3.TextMatrix(I, 8))), "#####0.00")
    
    MSHFlexGrid3.TextMatrix(I, 11) = IIf(IsNull(TmpRs4!ExpenseAccID), 0, TmpRs4!ExpenseAccID)
    If Val(MSHFlexGrid3.TextMatrix(I, 11)) > 0 Then
       MSHFlexGrid3.TextMatrix(I, 10) = GetExpenseType(Val(MSHFlexGrid3.TextMatrix(I, 11)))
    End If
    
    MSHFlexGrid3.TextMatrix(I, 12) = TmpRs4!TIID
    
    MSHFlexGrid3.TextMatrix(I, 13) = IIf(IsNull(TmpRs4!PostInvoiceNo), 0, TmpRs4!PostInvoiceNo)
    MSHFlexGrid3.TextMatrix(I, 14) = IIf(IsNull(TmpRs4!EntryNo), 0, TmpRs4!EntryNo)
    MSHFlexGrid3.TextMatrix(I, 15) = IIf(IsNull(TmpRs4!TISTAmount), 0, TmpRs4!TISTAmount)
    MSHFlexGrid3.TextMatrix(I, 16) = IIf(IsNull(TmpRs4!STAmount), 0, TmpRs4!STAmount)
    MSHFlexGrid3.TextMatrix(0, 8) = Val(MSHFlexGrid3.TextMatrix(0, 8)) + Val(MSHFlexGrid3.TextMatrix(I, 8))
    MSHFlexGrid3.TextMatrix(0, 7) = Val(MSHFlexGrid3.TextMatrix(0, 7)) + Val(MSHFlexGrid3.TextMatrix(I, 7))
    MSHFlexGrid3.TextMatrix(0, 15) = Val(MSHFlexGrid3.TextMatrix(0, 15)) + Val(MSHFlexGrid3.TextMatrix(I, 15))
    MSHFlexGrid3.TextMatrix(0, 16) = Val(MSHFlexGrid3.TextMatrix(0, 16)) + Val(MSHFlexGrid3.TextMatrix(I, 16))
    MSHFlexGrid3.TextMatrix(I, 17) = IIf(IsNull(TmpRs4!Locationname), "", TmpRs4!Locationname)
    MSHFlexGrid3.TextMatrix(I, 18) = IIf(IsNull(TmpRs4!NavCode), "", TmpRs4!NavCode)
    TmpRs4.MoveNext
Next
TxtAmount.Text = Val(MSHFlexGrid3.TextMatrix(0, 7)) + Val(MSHFlexGrid3.TextMatrix(0, 8))

End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsTxn_CRTDSAdjustment.MoveNext
If RsTxn_CRTDSAdjustment.EOF Then
   RsTxn_CRTDSAdjustment.MoveLast
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
RsTxn_CRTDSAdjustment.MovePrevious
If RsTxn_CRTDSAdjustment.BOF Then
   RsTxn_CRTDSAdjustment.MoveFirst
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
RsTxn_CRTDSAdjustment.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsTxn_CRTDSAdjustment.MoveLast
Call Indata
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False
End Sub
Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub AppendTIDetails()

If Check3.Value = 1 Then
   tmp = "Select NAVCode from Mst_Depository where DepositoryID='" & mDPID & "'"
   Call TmpTable
   If TmpRs.RecordCount > 0 Then
      mDPNAVCode = IIf(IsNull(TmpRs!NavCode), "", TmpRs!NavCode)
   End If

    tmp = "select distinct GSTNavCode from Mst_DepositoryStateWiseGST Where DepositoryID = '" & mDPID & "' "
    Call TmpTable
    
    If TmpRs.RecordCount > 1 Then
       If mDPNAVCode = "" Then
          For x = 1 To TmpRs.RecordCount
              If x = 1 Then
                 mDPNAVCode = "('" & TmpRs!GSTNavCode
              ElseIf x = TmpRs.RecordCount Then
                 mDPNAVCode = mDPNAVCode & "','" & TmpRs!GSTNavCode & "')"
              Else
                 mDPNAVCode = mDPNAVCode & "','" & TmpRs!GSTNavCode
              End If
              TmpRs.MoveNext
          Next
       Else
          For x = 1 To TmpRs.RecordCount
              If x = 1 Then
                 mDPNAVCode = "('" & mDPNAVCode & "','" & TmpRs!GSTNavCode
              ElseIf x = TmpRs.RecordCount Then
                 mDPNAVCode = mDPNAVCode & "','" & TmpRs!GSTNavCode & "')"
              Else
                 mDPNAVCode = mDPNAVCode & "','" & TmpRs!GSTNavCode
              End If
              TmpRs.MoveNext
          Next
       End If
    ElseIf TmpRs.RecordCount = 1 Then
       If mDPNAVCode = "" Then
          mDPNAVCode = "('" & TmpRs!GSTNavCode & "')"
       Else
          mDPNAVCode = "('" & mDPNAVCode & "','" & TmpRs!GSTNavCode & "')"
       End If
    End If

Else
   tmp = "Select NAVCode,RNAVCode from Mst_Client Where ClientIDRow=" & mClientIDRow & ""
   Call TmpTable
   If TmpRs.RecordCount > 0 Then
      mClientNAVCode = IIf(IsNull(TmpRs!NavCode), "", TmpRs!NavCode)
      mClientRNAVCode = IIf(IsNull(TmpRs!RNAVCode), "", TmpRs!RNAVCode)
   End If


    tmp = "Select distinct GSTNavCode from Mst_ClientStateWiseGST Where ClientIDRow = '" & mClientIDRow & "' And isnull(GSTNavCode,'') <> '' "
    Call TmpTable
    
    If TmpRs.RecordCount > 1 Then
       If mClientNAVCode = "" Then
          For x = 1 To TmpRs.RecordCount
              If x = 1 Then
                 mClientNAVCode = "('" & TmpRs!GSTNavCode
              ElseIf x = TmpRs.RecordCount Then
                 mClientNAVCode = mClientNAVCode & "','" & TmpRs!GSTNavCode & "')"
              Else
                 mClientNAVCode = mClientNAVCode & "','" & TmpRs!GSTNavCode
              End If
              TmpRs.MoveNext
          Next
       Else
          For x = 1 To TmpRs.RecordCount
              If x = 1 Then
                 mClientNAVCode = "('" & mClientNAVCode & "','" & TmpRs!GSTNavCode
              ElseIf x = TmpRs.RecordCount Then
                 mClientNAVCode = mClientNAVCode & "','" & TmpRs!GSTNavCode & "')"
              Else
                 mClientNAVCode = mClientNAVCode & "','" & TmpRs!GSTNavCode
              End If
              TmpRs.MoveNext
          Next
       End If
    ElseIf TmpRs.RecordCount = 1 Then
       If mClientNAVCode = "" Then
          mClientNAVCode = "('" & TmpRs!GSTNavCode & "')"
       Else
          mClientNAVCode = "('" & mClientNAVCode & "','" & TmpRs!GSTNavCode & "')"
       End If
    End If
    
    
    tmp = "Select distinct RGSTNavCode as 'GSTNavCode' from Mst_ClientStateWiseGST Where ClientIDRow = '" & mClientIDRow & "' And isnull(RGSTNavCode,'') <> '' "
    Call TmpTable
    
    If TmpRs.RecordCount > 1 Then
       If mClientRNAVCode = "" Then
          For x = 1 To TmpRs.RecordCount
              If x = 1 Then
                 mClientRNAVCode = "('" & TmpRs!GSTNavCode
              ElseIf x = TmpRs.RecordCount Then
                 mClientRNAVCode = mClientRNAVCode & "','" & TmpRs!GSTNavCode & "')"
              Else
                 mClientRNAVCode = mClientRNAVCode & "','" & TmpRs!GSTNavCode
              End If
              TmpRs.MoveNext
          Next
       Else
          For x = 1 To TmpRs.RecordCount
              If x = 1 Then
                 mClientRNAVCode = "('" & mClientRNAVCode & "','" & TmpRs!GSTNavCode
              ElseIf x = TmpRs.RecordCount Then
                 mClientRNAVCode = mClientRNAVCode & "','" & TmpRs!GSTNavCode & "')"
              Else
                 mClientRNAVCode = mClientRNAVCode & "','" & TmpRs!GSTNavCode
              End If
              TmpRs.MoveNext
          Next
       End If
    ElseIf TmpRs.RecordCount = 1 Then
       If mClientRNAVCode = "" Then
          mClientRNAVCode = "('" & TmpRs!GSTNavCode & "')"
       Else
          mClientRNAVCode = "('" & mClientRNAVCode & "','" & TmpRs!GSTNavCode & "')"
       End If
    End If


End If

''tmp = "Select TINo,TIDate,TIAmount,TIID,TI.PostInvoiceNo,TI.EntryNo,(Select Sum(TCRT.RecAmount) as RecAmount From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.CRID= TCRT.CRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'"
''tmp = tmp & " And TIID=TI.TIID Group By  TCRT.TIID)RecAmount,(Select Sum(TCRT.RecTDS) as RecTDS From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.CRID= TCRT.CRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'"
''tmp = tmp & " And TIID=TI.TIID Group By  TCRT.TIID)RecTDS,(Select Sum(TCRT.RecOtherDed) as RecOtherDed From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.CRID= TCRT.CRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'"
''tmp = tmp & " And TIID=TI.TIID Group By  TCRT.TIID)RecOtherDed"
''tmp = tmp & " From  Txn_TIDetails TI"
''tmp = tmp & " Where  TINo<>'' And TIAmount-((select isnull(Sum(RecAmount)+Sum(RecTDS)+Sum(RecOtherDed),0) From Txn_NewCashReceiptTIDetail TCRT"
''tmp = tmp & " Inner Join Txn_NewCashReceipt TCR on TCR.CRID= TCRT.CRID and TCR.SM_Prefix=TCRT.SM_Prefix"
''tmp = tmp & " Where TCRT.TIID=TI.TIID And TCR.ReturnFlag<>'Y')+(Select isnull(Sum(RecTDS)+Sum(RecOtherDed),0)  From Txn_CRTDSAdjustmentDetail TCRT"
''tmp = tmp & " Inner Join Txn_CRTDSAdjustment TCR on TCR.TDSAID = TCRT.TDSAID and TCR.SM_Prefix=TCRT.SM_Prefix"
''tmp = tmp & " Where TCRT.TIID=TI.TIID And TCR.Flag <> 'R' and TCR.G_UID<>'" & TxtG_UID & "' ))>0 And "
''tmp = tmp & " (LocationID=" & mLocationID & " or LocationID is null)"
''If Check3.Value = 0 Then
''   If mClientRNAVCode <> "" Then
''      tmp = tmp & " And (TI.NAVCode = '" & mClientNAVCode & "' or TI.NAVCode = '" & mClientRNAVCode & "')"
''   Else
''      tmp = tmp & " And TI.NAVCode = '" & mClientNAVCode & "'"
''   End If
''Else
''   tmp = tmp & " And TI.NAVCode='" & mDPNAVCode & "'"
''End If
''tmp = tmp & " And TI.TIID not in (Select TIID from Txn_CRTDSAdjustmentDetail TCRT"
''tmp = tmp & " Inner Join Txn_CRTDSAdjustment TCR on TCR.TDSAID = TCRT.TDSAID and TCR.SM_Prefix=TCRT.SM_Prefix"
''tmp = tmp & " Where TCR.Flag <> 'R' and TCR.G_UID='" & TxtG_UID & "')"
''
''Call Tmp3Table

''tmp = " Select TCRT.TIID,Sum(RecTDS) as RecTDSAmount,Sum(RecOtherDed)as RecOtherDed  From Txn_CRTDSAdjustmentDetail TCRT"
''tmp = tmp & " Inner Join Txn_CRTDSAdjustment TCR on TCR.TDSAID = TCRT.TDSAID and TCR.SM_Prefix=TCRT.SM_Prefix"
''tmp = tmp & " Where TCR.Flag <> 'R' and TCR.G_UID<>'" & TxtG_UID & "'"
''tmp = tmp & " Group By  TCRT.TIID"
''
''Call Tmp5Table

Wc1 = ""
If LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "admin" And LCase(Gen_UserRole) <> "superadmin" And LCase(Gen_UserRole) <> "fag" And LCase(Gen_UserRole) <> "fagadmin" And LCase(Gen_UserRole) <> "headcmg" Then
   Wc1 = " And (TI.LocationID=" & mLocationID & " or TI.LocationID is null)"
End If

If Check3.Value = 0 Then
   If InStr(1, mClientNAVCode, ")") = 0 And InStr(1, mClientRNAVCode, ")") = 0 Then
      If mClientRNAVCode <> "" Then
         Wc1 = Wc1 & " And (TI.NAVCode = '" & mClientNAVCode & "' or TI.NAVCode = '" & mClientRNAVCode & "')"
      Else
         Wc1 = Wc1 & " And TI.NAVCode = '" & mClientNAVCode & "'"
      End If
   Else
      If mClientRNAVCode <> "" Then
         If mClientNAVCode <> "" Then
            If InStr(1, mClientNAVCode, ")") = 0 Then
               Wc1 = Wc1 & " And (TI.NAVCode = '" & mClientNAVCode & "' or TI.NAVCode in " & mClientRNAVCode & ")"
            Else
               Wc1 = Wc1 & " And (TI.NAVCode in " & mClientNAVCode & " or TI.NAVCode in " & mClientRNAVCode & ")"
            End If
         Else
            Wc1 = Wc1 & " And (TI.NAVCode in " & mClientRNAVCode & ")"
         End If
      Else
         Wc1 = Wc1 & " And TI.NAVCode in " & mClientNAVCode & ""
      End If
   End If
Else
   If InStr(1, mDPNAVCode, ")") > 0 Then
      Wc1 = Wc1 & " And TI.NAVCode in " & mDPNAVCode & ""
   Else
      Wc1 = Wc1 & " And TI.NAVCode = '" & mDPNAVCode & "' "
   End If
End If

Wc1 = Wc1 & " And TI.TIID not in (Select TIID from Txn_CRTDSAdjustmentDetail TCRT"
Wc1 = Wc1 & " Inner Join Txn_CRTDSAdjustment TCR on TCR.TDSAID = TCRT.TDSAID and TCR.SM_Prefix=TCRT.SM_Prefix"
Wc1 = Wc1 & " Where TCR.Flag <> 'R' and TCR.G_UID='" & TxtG_UID & "')"

Call GetClientsRemInvoices(Wc1, "", TxtG_UID)

If TmpRs.RecordCount > 0 Then
   For gr = 1 To TmpRs.RecordCount
       rn = MSHFlexGrid3.Rows - 1
       MSHFlexGrid3.TextMatrix(rn, 0) = TmpRs!TINo
       MSHFlexGrid3.TextMatrix(rn, 1) = IIf(IsNull(TmpRs!TIDate), "", Format(TmpRs!TIDate, "dd-MMM-yyyy"))
       MSHFlexGrid3.TextMatrix(rn, 2) = IIf(IsNull(TmpRs!TIAmount), 0, TmpRs!TIAmount)
       MSHFlexGrid3.TextMatrix(rn, 3) = IIf(IsNull(TmpRs!RecAmount), 0, TmpRs!RecAmount)
       MSHFlexGrid3.TextMatrix(rn, 4) = IIf(IsNull(TmpRs!RecTDS), 0, TmpRs!RecTDS)
       MSHFlexGrid3.TextMatrix(rn, 5) = IIf(IsNull(TmpRs!RecOtherDed), 0, TmpRs!RecOtherDed)
''       If TmpRs5.RecordCount > 0 Then TmpRs5.MoveFirst
''       TmpRs5.Find "TIID=" & TmpRs3!TIID
''       If TmpRs5.EOF = False Then
''          MSHFlexGrid3.TextMatrix(rn, 4) = Val(MSHFlexGrid3.TextMatrix(rn, 4)) + IIf(IsNull(TmpRs5!RecTDSAmount), 0, TmpRs5!RecTDSAmount)
''           MSHFlexGrid3.TextMatrix(rn, 5) = Val(MSHFlexGrid3.TextMatrix(rn, 5)) + IIf(IsNull(TmpRs5!RecOtherDed), 0, TmpRs5!RecOtherDed)
''       End If

       MSHFlexGrid3.TextMatrix(rn, 6) = Format(Val(MSHFlexGrid3.TextMatrix(rn, 2)) - (Val(MSHFlexGrid3.TextMatrix(rn, 3)) + Val(MSHFlexGrid3.TextMatrix(rn, 4)) + Val(MSHFlexGrid3.TextMatrix(rn, 5))), "#####0.00")
       
       MSHFlexGrid3.TextMatrix(rn, 12) = TmpRs!TIID
       MSHFlexGrid3.TextMatrix(rn, 13) = IIf(IsNull(TmpRs!PostInvoiceNo), 0, TmpRs!PostInvoiceNo)
       MSHFlexGrid3.TextMatrix(rn, 14) = IIf(IsNull(TmpRs!EntryNo), 0, TmpRs!EntryNo)
       MSHFlexGrid3.TextMatrix(rn, 17) = IIf(IsNull(TmpRs!Locationname), "", TmpRs!Locationname)
       MSHFlexGrid3.TextMatrix(rn, 18) = IIf(IsNull(TmpRs!NavCode), "", TmpRs!NavCode)
       TmpRs.MoveNext
       MSHFlexGrid3.Rows = MSHFlexGrid3.Rows + 1
   Next
End If
End Sub

Private Sub Grid_Head()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 10
MSHFlexGrid1.TextMatrix(0, 0) = "Adjustment ID"
MSHFlexGrid1.TextMatrix(0, 1) = "Location "
MSHFlexGrid1.TextMatrix(0, 2) = "Adj. Date"
MSHFlexGrid1.TextMatrix(0, 3) = "Exchange Type"
MSHFlexGrid1.TextMatrix(0, 4) = "Client Name"
MSHFlexGrid1.TextMatrix(0, 5) = "Depository Name"
MSHFlexGrid1.TextMatrix(0, 6) = "Remark"
MSHFlexGrid1.TextMatrix(0, 7) = "Flag"
MSHFlexGrid1.TextMatrix(0, 8) = "BatchID"
MSHFlexGrid1.TextMatrix(0, 9) = "G_UID"

MSHFlexGrid1.ColWidth(0) = 1200
MSHFlexGrid1.ColWidth(1) = 2000
MSHFlexGrid1.ColWidth(2) = 1400
MSHFlexGrid1.ColWidth(3) = 1600
MSHFlexGrid1.ColWidth(4) = 2400
MSHFlexGrid1.ColWidth(5) = 2400
MSHFlexGrid1.ColWidth(6) = 1800
MSHFlexGrid1.ColWidth(7) = 1000
MSHFlexGrid1.ColWidth(8) = 800
MSHFlexGrid1.ColWidth(9) = 0

MSHFlexGrid1.WordWrap = True
MSHFlexGrid1.RowHeight(0) = 600

End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Txn_CRTDSAdjustment.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Txn_CRTDSAdjustment.xls")
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

Private Function CalSTAmount(mTotamt_ As Variant, mSTAmt_ As Variant, mBalAmt_ As Variant) As Variant
Dim Retval As Variant

Retval = 0
mSTPercent = 0
If Val(mSTAmt_) > 0 Then
   mSTPercent = (Val(mTotamt_)) / Val(mSTAmt_)
   mRatio = mBalAmt_ / mSTPercent
   Retval = (Val(mRatio))  'mBalAmt_
   'mSTPercent = Round(mSTPercent, 3) ', "#######0.00")
End If


If Val(mSTAmt_) > 0 Then
   If Round(Retval, 0) = 0 Then
      CalSTAmount = Round(Retval, 1) '"#######0.00")
   Else
      CalSTAmount = Round(Retval, 0)
   End If
Else
   CalSTAmount = Round(Retval, 0) '"#######0.00")
End If
End Function

Private Function CheckSTAmount() As Boolean
Dim mSTFlag, mWSTFlag As Variant
Dim Retval As Boolean
mSTFlag = "N"
mWSTFlag = "N"
Retval = True
For I = 2 To MSHFlexGrid3.Rows - 2
    If Val(MSHFlexGrid3.TextMatrix(I, 8)) > 0 Then
       If Val(MSHFlexGrid3.TextMatrix(I, 16)) > 0 Then
          mSTFlag = "Y"
       Else
         mWSTFlag = "Y"
       End If
    End If
Next
If mSTFlag = "Y" And mWSTFlag = "Y" Then
   Retval = False
End If
CheckSTAmount = Retval
End Function

Private Sub TxtSearchBy_LostFocus()


If TxtSearchBy = "" Then
   MSHFlexGrid3.TopRow = 2
   Exit Sub
End If


For I = 2 To MSHFlexGrid3.Rows - 1
    MSHFlexGrid3.row = I
    MSHFlexGrid3.Col = 0
    If LCase(Trim(TxtSearchBy)) = LCase(Trim(MSHFlexGrid3.TextMatrix(I, 0))) Then
       MSHFlexGrid3.TopRow = I
       Exit Sub
    End If
Next

End Sub


Private Function CheckCalSTAmount() As Boolean
Dim mSTFlag, mWSTFlag As Variant
Dim Retval As Boolean
mSTFlag = "N"
mWSTFlag = "N"
Retval = True
For I = 2 To MSHFlexGrid3.Rows - 2
    If Val(MSHFlexGrid3.TextMatrix(I, 8)) > 0 Then
       If Val(MSHFlexGrid3.TextMatrix(I, 16)) > 0 And Val(MSHFlexGrid3.TextMatrix(I, 15)) > 0 Then
          Retval = True
       ElseIf Val(MSHFlexGrid3.TextMatrix(I, 16)) = 0 And Val(MSHFlexGrid3.TextMatrix(I, 15)) = 0 Then
          Retval = True
       Else
          If Val(MSHFlexGrid3.TextMatrix(I, 8)) > 1 Then
             Retval = False
             Exit For
          End If
       End If
    End If
Next
CheckCalSTAmount = Retval
End Function

