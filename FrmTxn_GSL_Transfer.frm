VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmTxn_GSL_Transfer 
   Caption         =   "GSL Transfer"
   ClientHeight    =   9705
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10875
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9705
   ScaleWidth      =   10875
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   2325
      Left            =   120
      TabIndex        =   88
      Top             =   9600
      Visible         =   0   'False
      Width           =   15060
      Begin VB.ComboBox CmbSNewLocation 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1560
         Sorted          =   -1  'True
         TabIndex        =   93
         Top             =   840
         Width           =   3360
      End
      Begin VB.ComboBox CmbSOldCMP 
         BeginProperty Font 
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
         Sorted          =   -1  'True
         TabIndex        =   92
         Top             =   400
         Width           =   4485
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "Clear"
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
         Left            =   13890
         TabIndex        =   97
         Top             =   330
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
         Left            =   11880
         TabIndex        =   95
         Top             =   330
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
         Left            =   6000
         Sorted          =   -1  'True
         TabIndex        =   94
         Top             =   840
         Width           =   4485
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
         Left            =   12885
         TabIndex        =   96
         Top             =   330
         Width           =   1000
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
         Left            =   3030
         Sorted          =   -1  'True
         TabIndex        =   91
         Top             =   400
         Width           =   2865
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1215
         Left            =   120
         TabIndex        =   99
         Top             =   1320
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   2143
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSComCtl2.DTPicker STo 
         Height          =   360
         Left            =   1545
         TabIndex        =   90
         TabStop         =   0   'False
         Top             =   400
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
         Format          =   117964801
         CurrentDate     =   36494
      End
      Begin MSComCtl2.DTPicker SFrom 
         Height          =   360
         Left            =   75
         TabIndex        =   89
         TabStop         =   0   'False
         Top             =   400
         Width           =   1395
         _ExtentX        =   2461
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
         Format          =   117964801
         CurrentDate     =   36494
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "New Location"
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
         TabIndex        =   131
         Top             =   900
         Width           =   1215
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         Caption         =   "Old CMP"
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
         Left            =   7830
         TabIndex        =   103
         Top             =   165
         Width           =   825
      End
      Begin VB.Label Label44 
         Alignment       =   2  'Center
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
         Height          =   375
         Left            =   120
         TabIndex        =   102
         Top             =   150
         Width           =   1215
      End
      Begin VB.Label Label45 
         Alignment       =   2  'Center
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
         Left            =   1808
         TabIndex        =   101
         Top             =   150
         Width           =   915
      End
      Begin VB.Label Label37 
         Caption         =   "New CMP"
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
         Left            =   5040
         TabIndex        =   100
         Top             =   915
         Width           =   1065
      End
      Begin VB.Label Label34 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Old Location"
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
         Left            =   3900
         TabIndex        =   98
         Top             =   165
         Width           =   1125
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "New GSL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4125
      Left            =   120
      TabIndex        =   104
      Top             =   4395
      Width           =   15015
      Begin VB.TextBox TxtTransferedMadeupBags 
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   151
         TabStop         =   0   'False
         Top             =   3060
         Visible         =   0   'False
         Width           =   1875
      End
      Begin VB.ComboBox CmbNewCMPName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8055
         Sorted          =   -1  'True
         TabIndex        =   18
         Top             =   195
         Width           =   6375
      End
      Begin VB.TextBox TxtNewCommodityID 
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
         Left            =   13785
         Locked          =   -1  'True
         TabIndex        =   138
         TabStop         =   0   'False
         Top             =   1920
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.TextBox TxtNewExchangeTypeID 
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
         Left            =   13785
         Locked          =   -1  'True
         TabIndex        =   137
         TabStop         =   0   'False
         Top             =   2760
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.TextBox TxtNewGSLAvailableWgt 
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
         Left            =   13785
         Locked          =   -1  'True
         TabIndex        =   136
         TabStop         =   0   'False
         Top             =   2340
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.TextBox TxtNewGSLAvailableNofBags 
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
         Left            =   13785
         Locked          =   -1  'True
         TabIndex        =   135
         TabStop         =   0   'False
         Top             =   3180
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.TextBox TxtNewClientID 
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
         Left            =   13785
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   134
         TabStop         =   0   'False
         Top             =   3615
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.CommandButton CmdNewLocationID 
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
         Height          =   315
         Left            =   6275
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   218
         Width           =   375
      End
      Begin VB.TextBox TxtNewFlag 
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
         Left            =   2750
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   109
         TabStop         =   0   'False
         Top             =   1995
         Width           =   705
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
         Left            =   9060
         Locked          =   -1  'True
         TabIndex        =   107
         TabStop         =   0   'False
         Top             =   3630
         Width           =   5805
      End
      Begin VB.TextBox TxtTransferedQty 
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
         Left            =   4950
         Locked          =   -1  'True
         TabIndex        =   105
         TabStop         =   0   'False
         Top             =   2535
         Width           =   1695
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
         Left            =   4950
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   1995
         Width           =   1695
      End
      Begin VB.CommandButton CmdNewGodownNo 
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
         Height          =   315
         Left            =   6275
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   623
         Width           =   375
      End
      Begin VB.CommandButton CmdNewCMPName 
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
         Height          =   315
         Left            =   14450
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   218
         Width           =   375
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
         Height          =   945
         Left            =   1575
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   28
         Top             =   1000
         Width           =   13320
      End
      Begin VB.TextBox TxtTransferedBags 
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
         Left            =   1575
         Locked          =   -1  'True
         TabIndex        =   106
         TabStop         =   0   'False
         Top             =   2535
         Width           =   1875
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
         Left            =   1575
         Locked          =   -1  'True
         TabIndex        =   108
         TabStop         =   0   'False
         Top             =   3630
         Width           =   5115
      End
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1575
         TabIndex        =   29
         Top             =   2010
         Width           =   210
      End
      Begin VB.ComboBox CmbNewLocationID 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1575
         Sorted          =   -1  'True
         TabIndex        =   15
         Top             =   195
         Width           =   4650
      End
      Begin VB.ComboBox CmbNewGodownNo 
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
         Left            =   1575
         Sorted          =   -1  'True
         TabIndex        =   21
         Top             =   600
         Width           =   4650
      End
      Begin VB.ComboBox CmbNewLot 
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
         Left            =   12240
         Sorted          =   -1  'True
         TabIndex        =   26
         Top             =   600
         Width           =   2655
      End
      Begin VB.ComboBox CmbNewStack 
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
         Left            =   8055
         Sorted          =   -1  'True
         TabIndex        =   24
         Top             =   600
         Width           =   2655
      End
      Begin VB.TextBox TxtNewLocationID 
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
         Left            =   1575
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   195
         Visible         =   0   'False
         Width           =   5115
      End
      Begin VB.TextBox TxtNewCMPName 
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
         Left            =   8055
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   195
         Visible         =   0   'False
         Width           =   6825
      End
      Begin VB.TextBox TxtNewGodownNo 
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
         Left            =   1575
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   600
         Visible         =   0   'False
         Width           =   5115
      End
      Begin VB.TextBox TxtNewStack 
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
         Left            =   8055
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   600
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.TextBox TxtNewLot 
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
         Left            =   12240
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   600
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.TextBox TxtNewSMPrefix 
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
         Left            =   13785
         Locked          =   -1  'True
         TabIndex        =   132
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.TextBox TxtNewGSLID 
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
         Left            =   13785
         Locked          =   -1  'True
         TabIndex        =   133
         TabStop         =   0   'False
         Top             =   660
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.TextBox TxtNewGSLGUID 
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
         Left            =   13785
         Locked          =   -1  'True
         TabIndex        =   139
         TabStop         =   0   'False
         Top             =   1050
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.TextBox TxtNewCurrentDate 
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
         Left            =   13785
         Locked          =   -1  'True
         TabIndex        =   140
         TabStop         =   0   'False
         Top             =   1455
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.Label Label39 
         Caption         =   "Transferred Madeup Bags"
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
         TabIndex        =   152
         Top             =   2993
         Visible         =   0   'False
         Width           =   1470
      End
      Begin VB.Label Label29 
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
         Height          =   300
         Left            =   135
         TabIndex        =   130
         Top             =   225
         Width           =   1560
      End
      Begin VB.Label Label3 
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
         Height          =   300
         Left            =   12360
         TabIndex        =   128
         Top             =   3645
         Visible         =   0   'False
         Width           =   1680
      End
      Begin VB.Label Label5 
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
         Left            =   12360
         TabIndex        =   126
         Top             =   3240
         Visible         =   0   'False
         Width           =   1680
      End
      Begin VB.Label Label6 
         Caption         =   "Available Qty"
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
         Left            =   12360
         TabIndex        =   125
         Top             =   2400
         Visible         =   0   'False
         Width           =   1680
      End
      Begin VB.Label LblNewExchangeTypeID 
         Caption         =   "Exchange Type"
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
         Left            =   12360
         TabIndex        =   122
         Top             =   2790
         Visible         =   0   'False
         Width           =   1680
      End
      Begin VB.Label Label14 
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
         Left            =   12360
         TabIndex        =   121
         Top             =   1920
         Visible         =   0   'False
         Width           =   1680
      End
      Begin VB.Label LblGSLNewGUID 
         Caption         =   "GSL GUID"
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
         Left            =   12360
         TabIndex        =   116
         Top             =   1080
         Visible         =   0   'False
         Width           =   1680
      End
      Begin VB.Label Label23 
         Caption         =   "CurrentDate"
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
         Left            =   12360
         TabIndex        =   115
         Top             =   1440
         Visible         =   0   'False
         Width           =   1680
      End
      Begin VB.Label LblNewGodownNo 
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
         Left            =   135
         TabIndex        =   124
         Top             =   630
         Width           =   1335
      End
      Begin VB.Label LblNewCMPName 
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
         Height          =   300
         Left            =   6855
         TabIndex        =   123
         Top             =   225
         Width           =   1680
      End
      Begin VB.Label lblNewLot 
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
         Left            =   11655
         TabIndex        =   119
         Top             =   660
         Width           =   615
      End
      Begin VB.Label LblNewStack 
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
         Left            =   6855
         TabIndex        =   118
         Top             =   660
         Width           =   855
      End
      Begin VB.Label LblRemark 
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
         Left            =   135
         TabIndex        =   117
         Top             =   1322
         Width           =   1800
      End
      Begin VB.Label Label2 
         Caption         =   "Flag"
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
         Height          =   270
         Left            =   1935
         TabIndex        =   129
         Top             =   2040
         Width           =   630
      End
      Begin VB.Label Label17 
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
         Height          =   240
         Left            =   135
         TabIndex        =   120
         Top             =   2010
         Width           =   1560
      End
      Begin VB.Label Label24 
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
         Left            =   135
         TabIndex        =   114
         Top             =   3690
         Width           =   1695
      End
      Begin VB.Label Label25 
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
         Left            =   7695
         TabIndex        =   113
         Top             =   3705
         Width           =   975
      End
      Begin VB.Label LblQty 
         Caption         =   "Qty to be transferred"
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
         Left            =   3510
         TabIndex        =   112
         Top             =   1935
         Width           =   1470
      End
      Begin VB.Label Label27 
         Caption         =   "Transferred Bags"
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
         Left            =   135
         TabIndex        =   111
         Top             =   2408
         Width           =   1350
      End
      Begin VB.Label Label28 
         Caption         =   "Transferred Qty"
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
         Left            =   3615
         TabIndex        =   110
         Top             =   2468
         Width           =   1350
      End
      Begin VB.Label LblNewGSLID 
         Caption         =   "GSL ID"
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
         Left            =   12360
         TabIndex        =   127
         Top             =   720
         Visible         =   0   'False
         Width           =   1680
      End
   End
   Begin VB.Frame Frame3 
      Height          =   705
      Left            =   120
      TabIndex        =   67
      Top             =   -15
      Width           =   15015
      Begin VB.TextBox TxtTransferID 
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
         Left            =   1575
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   240
         Width           =   1875
      End
      Begin MSComCtl2.DTPicker TxtTransferDate 
         Height          =   375
         Left            =   4920
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   240
         Width           =   1560
         _ExtentX        =   2752
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
         Format          =   117964801
         CurrentDate     =   39884
      End
      Begin VB.Label Label8 
         Caption         =   "Transfer Date"
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
         Left            =   3600
         TabIndex        =   69
         Top             =   270
         Width           =   1800
      End
      Begin VB.Label Label4 
         Caption         =   "Transfer ID"
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
         TabIndex        =   68
         Top             =   270
         Width           =   1800
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Original GSL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3660
      Left            =   120
      TabIndex        =   41
      Top             =   700
      Width           =   15015
      Begin VB.TextBox TxtBalanceMadeupBags 
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
         TabIndex        =   149
         TabStop         =   0   'False
         Top             =   3120
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.TextBox TxtMadeupBags 
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
         Left            =   10575
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   3120
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.TextBox TxtMadeupQty 
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
         Left            =   13215
         Locked          =   -1  'True
         TabIndex        =   148
         TabStop         =   0   'False
         Top             =   3120
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.TextBox TxtBalanceBags 
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
         TabIndex        =   145
         TabStop         =   0   'False
         Top             =   2460
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.TextBox TxtTransferBags 
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
         Left            =   10560
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   2460
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.TextBox TxtTransferQty 
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
         Left            =   13200
         Locked          =   -1  'True
         TabIndex        =   143
         TabStop         =   0   'False
         Top             =   2460
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1560
         TabIndex        =   14
         Top             =   3180
         Width           =   210
      End
      Begin VB.TextBox TxtCurrentDate 
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
         Left            =   14400
         Locked          =   -1  'True
         TabIndex        =   86
         TabStop         =   0   'False
         Top             =   3120
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.TextBox TxtBGSLID 
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
         Left            =   1575
         Locked          =   -1  'True
         TabIndex        =   82
         TabStop         =   0   'False
         Top             =   3120
         Visible         =   0   'False
         Width           =   1875
      End
      Begin VB.TextBox TxtCDTFDate 
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
         Left            =   4950
         Locked          =   -1  'True
         TabIndex        =   79
         TabStop         =   0   'False
         Top             =   2700
         Width           =   1725
      End
      Begin VB.TextBox TxtCDTFNo 
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
         Left            =   1575
         Locked          =   -1  'True
         TabIndex        =   78
         TabStop         =   0   'False
         Top             =   2700
         Width           =   1875
      End
      Begin VB.TextBox TxtPledgeWeight 
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
         Left            =   4950
         Locked          =   -1  'True
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   2260
         Width           =   1725
      End
      Begin VB.TextBox TxtPledgeBags 
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
         Left            =   1575
         Locked          =   -1  'True
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   2260
         Width           =   1875
      End
      Begin VB.TextBox TxtPledgeNo 
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
         Left            =   4965
         Locked          =   -1  'True
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   1850
         Width           =   1725
      End
      Begin VB.TextBox TxtISINID 
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   1850
         Width           =   1875
      End
      Begin VB.TextBox TxtGSLGUID 
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
         Left            =   14385
         Locked          =   -1  'True
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   2700
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.TextBox TxtGSLID 
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
         Left            =   14385
         Locked          =   -1  'True
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   2260
         Visible         =   0   'False
         Width           =   1305
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
         Height          =   360
         Left            =   8040
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   1425
         Width           =   6825
      End
      Begin VB.TextBox TxtFlag 
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
         Left            =   4950
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   3120
         Width           =   1005
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
         Height          =   360
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   1425
         Width           =   1875
      End
      Begin VB.TextBox TxtGSLAvailableWgt 
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
         Left            =   4965
         Locked          =   -1  'True
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   1425
         Width           =   1725
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
         Left            =   12210
         Sorted          =   -1  'True
         TabIndex        =   10
         Top             =   610
         Width           =   2655
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
         Left            =   8040
         Sorted          =   -1  'True
         TabIndex        =   8
         Top             =   615
         Width           =   2655
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
         Left            =   1560
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   610
         Width           =   5115
      End
      Begin VB.ComboBox CmbCMPName 
         BeginProperty Font 
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
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   200
         Width           =   6840
      End
      Begin VB.TextBox TxtExchangeTypeID 
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   1015
         Width           =   5115
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
         Left            =   8040
         Locked          =   -1  'True
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   1015
         Width           =   6840
      End
      Begin VB.TextBox TxtPreFix 
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
         Left            =   14385
         Locked          =   -1  'True
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   1850
         Visible         =   0   'False
         Width           =   1305
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
         Left            =   1560
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   200
         Width           =   5115
      End
      Begin VB.TextBox TxtLocationID 
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   200
         Visible         =   0   'False
         Width           =   5115
      End
      Begin VB.TextBox TxtCMPName 
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
         Left            =   8040
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   200
         Visible         =   0   'False
         Width           =   6825
      End
      Begin VB.TextBox TxtGodownNo 
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   610
         Visible         =   0   'False
         Width           =   5115
      End
      Begin VB.TextBox TxtLot 
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
         Left            =   12210
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   610
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.TextBox TxtStack 
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
         Left            =   8040
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   618
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.TextBox TxtFGSLID 
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
         Left            =   4950
         Locked          =   -1  'True
         TabIndex        =   83
         TabStop         =   0   'False
         Top             =   3120
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.Label Label38 
         Caption         =   "Balance Madeup Bags"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   6720
         TabIndex        =   150
         Top             =   3030
         Visible         =   0   'False
         Width           =   1440
      End
      Begin VB.Label Label36 
         Caption         =   "Madeup Qty"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   12255
         TabIndex        =   147
         Top             =   3030
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Label Label35 
         Caption         =   "Madeup Bags"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   9600
         TabIndex        =   146
         Top             =   3030
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.Label Label33 
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
         Height          =   540
         Left            =   6705
         TabIndex        =   144
         Top             =   2370
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.Label Label32 
         Caption         =   "Transfer Qty"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   12240
         TabIndex        =   142
         Top             =   2370
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Label Label31 
         Caption         =   "Transfer Bags"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   9585
         TabIndex        =   141
         Top             =   2370
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.Label Label22 
         Caption         =   "CurrentDate"
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
         Left            =   13200
         TabIndex        =   87
         Top             =   3150
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Label Label18 
         Caption         =   "CDTF No"
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
         TabIndex        =   81
         Top             =   2730
         Width           =   1560
      End
      Begin VB.Label Label16 
         Caption         =   "CDTF Date"
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
         Left            =   3550
         TabIndex        =   80
         Top             =   2730
         Width           =   1485
      End
      Begin VB.Label Label15 
         Caption         =   "Pledge Bags"
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
         TabIndex        =   77
         Top             =   2290
         Width           =   1560
      End
      Begin VB.Label Label9 
         Caption         =   "ISIN ID"
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
         TabIndex        =   76
         Top             =   1903
         Width           =   1560
      End
      Begin VB.Label Label11 
         Caption         =   "Pledge Weight"
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
         Left            =   3550
         TabIndex        =   75
         Top             =   2310
         Width           =   1485
      End
      Begin VB.Label Label10 
         Caption         =   "Pledge No"
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
         Left            =   3550
         TabIndex        =   72
         Top             =   1905
         Width           =   1485
      End
      Begin VB.Label LblGSLGUID 
         Caption         =   "GSL GUID"
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
         Left            =   13185
         TabIndex        =   66
         Top             =   2730
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Label Label1 
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
         Height          =   390
         Left            =   120
         TabIndex        =   63
         Top             =   3165
         Width           =   840
      End
      Begin VB.Label LblFlag 
         Caption         =   "Flag"
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
         Height          =   255
         Left            =   3555
         TabIndex        =   62
         Top             =   3180
         Width           =   645
      End
      Begin VB.Label LblClientID 
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
         Height          =   300
         Left            =   6825
         TabIndex        =   61
         Top             =   1455
         Width           =   1800
      End
      Begin VB.Label LblGSLID 
         Caption         =   "GSL ID"
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
         Left            =   13185
         TabIndex        =   60
         Top             =   2295
         Visible         =   0   'False
         Width           =   1080
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
         Left            =   105
         TabIndex        =   56
         Top             =   1500
         Width           =   1560
      End
      Begin VB.Label Label13 
         Caption         =   "Available Qty"
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
         Left            =   3550
         TabIndex        =   55
         Top             =   1485
         Width           =   1485
      End
      Begin VB.Label LblLot 
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
         Left            =   11520
         TabIndex        =   52
         Top             =   660
         Width           =   615
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
         Left            =   6825
         TabIndex        =   51
         Top             =   660
         Width           =   855
      End
      Begin VB.Label lblGodownNo 
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
         Left            =   105
         TabIndex        =   50
         Top             =   645
         Width           =   1560
      End
      Begin VB.Label LblCMPName 
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
         Height          =   300
         Left            =   6825
         TabIndex        =   49
         Top             =   285
         Width           =   1800
      End
      Begin VB.Label LblExchangeTypeID 
         Caption         =   "Exchange Type"
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
         Left            =   105
         TabIndex        =   48
         Top             =   1035
         Width           =   1560
      End
      Begin VB.Label LblCommodityID 
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
         Left            =   6825
         TabIndex        =   47
         Top             =   1035
         Width           =   1080
      End
      Begin VB.Label Label7 
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
         Height          =   300
         Left            =   13200
         TabIndex        =   46
         Top             =   1875
         Visible         =   0   'False
         Width           =   1200
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
         Height          =   300
         Left            =   105
         TabIndex        =   42
         Top             =   270
         Width           =   1560
      End
      Begin VB.Label Label19 
         Caption         =   "Fwd GSL ID"
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
         Left            =   3855
         TabIndex        =   84
         Top             =   3195
         Visible         =   0   'False
         Width           =   1800
      End
      Begin VB.Label Label20 
         Caption         =   "Back GSL ID"
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
         Left            =   240
         TabIndex        =   85
         Top             =   3165
         Visible         =   0   'False
         Width           =   1560
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1120
      Left            =   120
      TabIndex        =   64
      Top             =   8520
      Width           =   15015
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
         TabIndex        =   31
         Top             =   195
         Width           =   2940
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
         TabIndex        =   36
         Top             =   600
         Width           =   2940
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
         Left            =   3045
         TabIndex        =   32
         Top             =   195
         Width           =   3060
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
         Left            =   3045
         TabIndex        =   37
         Top             =   600
         Width           =   3060
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
         Left            =   6105
         TabIndex        =   33
         Top             =   195
         Width           =   2940
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
         Left            =   6105
         TabIndex        =   38
         Top             =   600
         Width           =   2940
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
         Left            =   9045
         TabIndex        =   34
         Top             =   195
         Width           =   2940
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
         Left            =   9045
         TabIndex        =   39
         Top             =   600
         Width           =   2940
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
         Left            =   11985
         TabIndex        =   35
         Top             =   195
         Width           =   2820
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
         Left            =   11985
         TabIndex        =   40
         Top             =   600
         Width           =   2820
      End
   End
End
Attribute VB_Name = "FrmTxn_GSL_Transfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mLocationID, mCMPID, mCommodity As Double
Dim mNewCMPID, mNewCommodity As Double

Dim mExchageTypeId As Double
Dim mClientIDRowGsl As Double
Dim mNewClientIDRowGsl As Double
Dim mNewExchageTypeId As Double
Dim mCMPCurrentDate, mNewCMPCurrentDate As Date

Dim mGodownID, mCmbStack, mCmbLot, mNewGodownID, mCmbNewStack, mCmbNewLot, mGSLID, mNewGSLID, Edit_Flg As Variant
Dim mSLocationID As Double, mSNewLocationID As Double

Dim mDepoBags, mDepoWt As Variant
Dim mConnectedGSLID, mNewMGSLID, mMadeupFlag, tmp1 As Variant
Dim mNewLocationID As Double
Dim mNewSM_Prefix As String
Dim mContractID As Variant
Dim mOldGodownYearlyFlag, mNewGodownYearlyFlag As Variant

''-- 10 Aug 2016
Dim mBagSize, mOriginalDepositFlag As Variant
Dim mSplitBagsFlag, mOrignalFullTransfer, mOriginalDeposit As Boolean

Private Sub CmbNewLocationID_GotFocus()
tmp = CmbNewLocationID.Text
If LCase(Gen_UserRole) = "power" Or LCase(Gen_UserRole) = "headcmg" Then
   If Gen_WcLocation <> "" Then
      Call TableMst_Location(" Where " & Gen_WcLocation)
   Else
      Call TableMst_Location("")
   End If
Else
   If Trim(CmbLocationID.Text) = "" Then
      MsgBox "Select Old Location first!!!"
      CmbLocationID.SetFocus
      Exit Sub
   End If
   Call TableMst_Location(" where LocationID = " & mLocationID)
End If
CmbNewLocationID.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Location found!!!"
   CmbNewLocationID.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
   CmbNewLocationID.AddItem RsMst_Location!Locationname
   RsMst_Location.MoveNext
Next
CmbNewLocationID.Text = tmp
End Sub

Private Sub CmbNewLocationID_LostFocus()
If CmbNewLocationID.Text = "" Then
   mNewLocationID = 0
   TxtNewPreFix = ""
   mNewSM_Prefix = ""
   CmbNewCMPName.Text = ""
   CmbNewGodownNo.Text = ""
   CmbNewStack.Text = ""
   CmbNewLot.Text = ""
   TxtNewCurrentDate = ""
   Call ClearNewGSLDetail
   Exit Sub
End If
' Code for clearing Godown combox in case of change
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbNewLocationID.Text & "'"

If RsMst_Location.EOF Then
   MsgBox "Invalid Location selection!!!"
   CmbNewLocationID.SetFocus
   Exit Sub
End If
If mNewLocationID <> RsMst_Location!LocationID Then
   CmbNewCMPName.Text = ""
   CmbNewGodownNo.Text = ""
   CmbNewStack.Text = ""
   CmbNewLot.Text = ""
   TxtNewCurrentDate = ""
   Call ClearNewGSLDetail
End If
mNewLocationID = RsMst_Location!LocationID
mNewSM_Prefix = RsMst_Location!SM_Prefix

End Sub

Private Sub CmbSNewLocation_GotFocus()
tmp = CmbSNewLocation.Text

If Gen_WcLocation = "" Then
   Call TableMst_Location("")
Else
   Call TableMst_Location("where " & Gen_WcLocation)
End If

CmbSNewLocation.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbSNewLocation.AddItem RsMst_Location!Locationname
    RsMst_Location.MoveNext
Next
CmbSNewLocation.Text = tmp
End Sub

Private Sub CmbSNewLocation_LostFocus()

If CmbSNewLocation.Text = "" Then
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbSNewLocation.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection "
   CmbSNewLocation.SetFocus
   Exit Sub
End If

'mNewSM_Prefix = RsMst_Location!SM_Prefix
mSNewLocationID = RsMst_Location!LocationID
End Sub

Private Sub CmdClear_Click()

SFrom.Value = Date - 365
STo.Value = Date
CmbPreFix.Text = ""
CmbSNewLocation.Text = ""
CmbSCmp.Text = ""
CmbSOldCMP.Text = ""
Call Grid_Head1

End Sub

Private Sub CmdNewLocationID_Click()
FrmMst_Location.Show
End Sub

Private Sub CmdSearch_Click()
Dim wc As String
Dim tmp1 As String
Call Grid_Head1

If STo.Value < SFrom.Value Then
   MsgBox "'To date' must be always greater than 'From date'!!!"
   STo.SetFocus
   Exit Sub
End If

tmp = "Select TGT.GSLTransferID,TGT.GSLTransferDate,ML1.LocationName,MC1.CMPName as 'NewCMP', "
tmp = tmp & " MG1.GodownNo +'~'+ MGSL1.StackNo +'~'+ MGSL1.LotNo as 'NewGSL',MGSL1.BalanceBags,MGSL1.BalanceWeight,MGSL1.Status,"
tmp = tmp & " MGSL1.SpillageFlag,MC2.CMPName as 'OldCMP',MG2.GodownNo +'~'+ MGSL2.StackNo +'~'+ MGSL2.LotNo as 'OldGSL',MGSL2.SpillageFlag,"
tmp = tmp & " TGT.Remarks , TGT.G_UID,TGT.TransferedBags,TGT.TransferedWeight, ML2.LocationName NewLocation,TGT.MadeupBags,TGT.MadeupNew_Mst_GSL_ID,TGT.MadeupOLD_Mst_GSL_ID"
tmp = tmp & " From Txn_GSL_Transfer  TGT"
tmp = tmp & " Inner Join Mst_GSL MGSL1 On TGT.New_Mst_GSL_ID = MGSL1.Mst_GSL_ID And TGT.NewSM_Prefix = MGSL1.SM_Prefix"
tmp = tmp & " Inner Join Mst_GSL MGSL2 On TGT.Old_Mst_GSL_ID = MGSL2.Mst_GSL_ID And TGT.SM_Prefix = MGSL2.SM_Prefix"
tmp = tmp & " Inner Join Mst_CMP MC1 On MGSL1.CMPID = MC1.CMPID"
tmp = tmp & " Inner Join Mst_CMP MC2 On MGSL2.CMPID = MC2.CMPID"
tmp = tmp & " Inner Join Mst_Godown MG1 On MGSL1.GodownID = MG1.GodownID"
tmp = tmp & " Inner Join Mst_Godown MG2 On MGSL2.GodownID = MG2.GodownID"
tmp = tmp & " Inner Join Mst_Location ML1 On TGT.SM_Prefix = ML1.SM_Prefix"
tmp = tmp & " Inner Join Mst_Location ML2 On TGT.NewSM_Prefix = ML2.SM_Prefix"

wc = ""

If Gen_DBType = "MDB" Then
   wc = " Where TGT.GSLTransferDate Between #" & Format(SFrom, Gen_DatFormat) & "# And #" & Format(STo, Gen_DatFormat) & "#) "
Else
   wc = " Where TGT.GSLTransferDate Between '" & Format(SFrom, Gen_DatFormat) & "' And '" & Format(STo, Gen_DatFormat) & "' "
End If

If CmbPreFix.Text <> "" Then
   wc = GenWc(wc, CmbPreFix.Text, "ML1.LocationName", "S", "N")
Else
   If Gen_WcLocation <> "" Then
      tmp1 = Replace(Gen_WcLocation, "LocationID", "ML1.LocationID")
      wc = GenWc(wc, "", tmp1, "N", "I")
   End If
End If

If CmbSNewLocation.Text <> "" Then
   wc = GenWc(wc, CmbSNewLocation.Text, "ML2.LocationName", "S", "N")
Else
   If Gen_WcLocation <> "" Then
      tmp1 = Replace(Gen_WcLocation, "LocationID", "ML2.LocationID")
      wc = GenWc(wc, "", tmp1, "N", "I")
   End If
End If

If CmbSCmp.Text <> "" Then
   wc = GenWc(wc, CmbSCmp.Text, "MC1.CMPName", "S", "N")
End If

If CmbSOldCMP.Text <> "" Then
   wc = GenWc(wc, CmbSOldCMP.Text, "MC2.CMPName", "S", "N")
End If


tmp = tmp & " " & wc & " " & " Order by GSLTransferID,TGT.SM_Prefix"

Call TmpTable

Call Grid_Head1


For iu = 1 To TmpRs.RecordCount
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
    MSHFlexGrid1.TextMatrix(iu, 0) = TmpRs!GSLTransferID
    MSHFlexGrid1.TextMatrix(iu, 1) = Format(TmpRs!GSLTransferDate, "dd-mmm-yyyy")
    MSHFlexGrid1.TextMatrix(iu, 2) = TmpRs!NewLocation
    MSHFlexGrid1.TextMatrix(iu, 3) = TmpRs!NewCMP 'Format(TmpRs!PledgeReleaseDate, "dd-mmm-yyyy")
    MSHFlexGrid1.TextMatrix(iu, 4) = TmpRs!NewGSL
    MSHFlexGrid1.TextMatrix(iu, 5) = TmpRs!BalanceBags
    MSHFlexGrid1.TextMatrix(iu, 6) = TmpRs!BalanceWeight
    MSHFlexGrid1.TextMatrix(iu, 7) = TmpRs!Status
    MSHFlexGrid1.TextMatrix(iu, 8) = TmpRs!SpillageFlag
    MSHFlexGrid1.TextMatrix(iu, 9) = TmpRs!OldCMP
    MSHFlexGrid1.TextMatrix(iu, 10) = TmpRs!OLDGSL
    MSHFlexGrid1.TextMatrix(iu, 11) = TmpRs!SpillageFlag 'IIf(IsNull(TmpRs!ReleaseBags), 0, TmpRs!ReleaseBags)
    MSHFlexGrid1.TextMatrix(iu, 12) = IIf(IsNull(TmpRs!Remarks), "", TmpRs!Remarks) 'IIf(IsNull(TmpRs!ReleaseQty), 0, TmpRs!ReleaseQty)
    MSHFlexGrid1.TextMatrix(iu, 13) = TmpRs!G_UID
    MSHFlexGrid1.TextMatrix(iu, 14) = IIf(IsNull(TmpRs!TransferedBags), "", TmpRs!TransferedBags)
    
    MSHFlexGrid1.TextMatrix(iu, 15) = IIf(IsNull(TmpRs!MadeupBags), "", TmpRs!MadeupBags)
    MSHFlexGrid1.TextMatrix(iu, 16) = IIf(IsNull(TmpRs!MadeupNew_Mst_GSL_ID), "", TmpRs!MadeupNew_Mst_GSL_ID)
    MSHFlexGrid1.TextMatrix(iu, 17) = IIf(IsNull(TmpRs!MadeupOLD_Mst_GSL_ID), "", TmpRs!MadeupOLD_Mst_GSL_ID)
    
    MSHFlexGrid1.TextMatrix(iu, 18) = IIf(IsNull(TmpRs!TransferedWeight), "", TmpRs!TransferedWeight)
    MSHFlexGrid1.TextMatrix(iu, 19) = TmpRs!Locationname
    TmpRs.MoveNext
Next


End Sub

Private Sub EditCmd_Click()
Edit_Flg = "E"
TxtRemark.Locked = False
End Sub

Private Sub FirstCmd_Click()
RsTxn_GSL_Transfer.MoveFirst
Call Indata
'TmpRs.MoveLast
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
Call CmdClear_Click
If Gen_WcLocation <> "" Then
   Call TableTxn_GSL_Transfer(" Where (SM_Prefix in (Select Distinct SM_Prefix from Mst_Location Where " & Gen_WcLocation & ") Or NewSM_Prefix in (Select Distinct SM_Prefix from Mst_Location Where " & Gen_WcLocation & ")) ")
Else
   Call TableTxn_GSL_Transfer
End If
If RsTxn_GSL_Transfer.RecordCount = 0 Then
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
Private Sub AddCmd_Click()

Edit_Flg = "A"
Call GButtonLock(Me, False)
TxtTransferDate.Enabled = False
CmbLocationID.Visible = True
TxtLocationID.Visible = False
CmbCMPName.Visible = True
TxtCMPName.Visible = False
CmbGodownNo.Visible = True
TxtGodownNo.Visible = False
CmbStack.Visible = True
TxtStack.Visible = False
CmbLot.Visible = True
TxtLot.Visible = False
'TxtTransferDate.Value = Date

''-- 10 Aug 2016 For GSL Split
Label31.Visible = False: TxtTransferBags = "": TxtTransferBags.Visible = False
Label32.Visible = False: TxtTransferQty = "":  TxtTransferQty.Visible = False
Label33.Visible = False: TxtBalanceBags = "":  TxtBalanceBags.Visible = False
Label35.Visible = False: TxtMadeupBags = "":   TxtMadeupBags.Visible = False
Label38.Visible = False: TxtBalanceMadeupBags = "":  TxtBalanceMadeupBags.Visible = False
Label39.Visible = False: TxtTransferedMadeupBags = "": TxtTransferedMadeupBags.Visible = False


CmbLocationID.Text = TxtLocationID
If CmbLocationID.Text <> "" Then
   CmbLocationID_GotFocus
   CmbLocationID_LostFocus
End If

CmbCMPName.Text = TxtCMPName
If CmbCMPName.Text <> "" Then
   CmbCMPName_GotFocus
   CmbCMPName_LostFocus
End If
CmbGodownNo.Text = TxtGodownNo
If CmbGodownNo.Text <> "" Then
   CmbGodownNo_GotFocus
   CmbGodownNo_LostFocus
End If
CmbStack.Text = TxtStack
If CmbStack.Text <> "" Then
   CmbStack_GotFocus
   CmbStack_LostFocus
End If

CmbCMPName.Text = TxtCMPName

CmbGodownNo = TxtGodownNo
CmbStack = TxtStack

CmbLot = ""

'TxtPreFix = ""
Call ClearOrignalGSLDetail

CmbNewLocationID.Visible = True
CmdNewLocationID.Visible = True
TxtNewLocationID.Visible = False

CmbNewCMPName.Visible = True
CmdNewCMPName.Visible = True
TxtNewCMPName.Visible = False

CmbNewGodownNo.Visible = True
CmdNewGodownNo.Visible = True
txtNewGodownNo.Visible = False

CmbNewStack.Visible = True
TxtNewStack.Visible = False

CmbNewLot.Visible = True
TxtNewLot.Visible = False

Call ClearNewGSLDetail
TxtRemark.Locked = False
TxtRemark = ""

TxtCreated = ""
TxtUpdated = ""

'CmbNewCMPName = ""
'CmbNewGodownNo = ""
'CmbNewStack = ""
CmbNewLot = ""
TxtTransferID = ""
'txtTransferDate = Date

If RsTxn_GSL_Transfer.RecordCount > 0 Then
   CmbLocationID.SetFocus
End If

End Sub

Private Sub CmbLocationID_GotFocus()
tmp = CmbLocationID.Text
If Gen_WcLocation = "" Then
   Call TableMst_Location(" Where LocationID <> 0 ")
Else
   Call TableMst_Location(" Where " & Gen_WcLocation & " ")
End If
CmbLocationID.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbLocationID.AddItem RsMst_Location!Locationname
    RsMst_Location.MoveNext
Next
CmbLocationID.Text = tmp
End Sub

Private Sub CmbLocationID_LostFocus()
If CmbLocationID.Text = "" Then
   CmbCMPName.Text = ""
   CmbGodownNo.Text = ""
   CmbStack.Text = ""
   CmbLot.Text = ""
   mLocationID = 0
   CmbNewLocationID.Text = ""
   CmbNewCMPName.Text = ""
   CmbNewGodownNo.Text = ""
   TxtNewSMPrefix = ""
   Call ClearOrignalGSLDetail
   Call ClearNewGSLDetail
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocationID.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection "
   CmbLocationID.SetFocus
   Exit Sub
End If

If mLocationID <> RsMst_Location!LocationID Then
   CmbCMPName.Text = ""
   CmbGodownNo.Text = ""
   CmbStack.Text = ""
   CmbLot.Text = ""
   TxtPreFix = ""
   CmbNewLocationID.Text = ""
   CmbNewCMPName.Text = ""
   CmbNewGodownNo.Text = ""
   TxtNewSMPrefix = ""
   Call ClearOrignalGSLDetail
   Call ClearNewGSLDetail
End If

mLocationID = RsMst_Location!LocationID
TxtPreFix = RsMst_Location!SM_Prefix
End Sub

Private Sub CmbCMPName_GotFocus()
If CmbLocationID.Text = "" Then
   MsgBox "Select Location!!!"
   If RsMst_Location.RecordCount > 0 Then
      CmbLocationID.SetFocus
   End If
   Exit Sub
End If
tmp = CmbCMPName.Text

Call TableMst_CMP(" Where ((CloseDate is null And LocationID = " & mLocationID & " ) Or CMPID = " & Val(mCMPID) & " )")
CmbCMPName.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_CMP.RecordCount
    CmbCMPName.AddItem RsMst_CMP!CMPName
    RsMst_CMP.MoveNext
Next
CmbCMPName.Text = tmp

End Sub

Private Sub CmbCMPName_LostFocus()
If CmbCMPName.Text = "" Then
   CmbGodownNo.Text = ""
   CmbStack.Text = ""
   CmbLot.Text = ""
   TxtCurrentDate = ""
   mCMPID = 0
   Exit Sub
End If

RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPName = '" & CmbCMPName.Text & "'"
If RsMst_CMP.EOF Then
   TxtGodownNo.Text = ""
   TxtStack.Text = ""
   TxtLot.Text = ""
   MsgBox "Invalid selection "
   CmbCMPName.SetFocus
   Exit Sub
End If
If mCMPID <> RsMst_CMP!CMPID Then
   CmbGodownNo.Text = ""
   CmbStack.Text = ""
   CmbLot.Text = ""
   TxtCurrentDate = ""
   'TxtTransferDate = CDate(TxtCurrentDate) + 1
   Call ClearOrignalGSLDetail
End If
mCMPID = RsMst_CMP!CMPID
TxtCurrentDate = IIf(IsNull(RsMst_CMP!CurrentDate), "", RsMst_CMP!CurrentDate)
TxtTransferDate = CDate(TxtCurrentDate) + 1
'TxtPreFix = GetPreFix(RsMst_CMP!CMPID, "P")
End Sub

Private Sub CmbGodownNo_GotFocus()
If CmbCMPName.Text = "" Then
   MsgBox "Select CMP !!!! "
   CmbCMPName.SetFocus
   Exit Sub
End If
tmp = CmbGodownNo.Text
Call TableMst_Godown("Where CloseDate is null And CMPID = " & Val(mCMPID) & "")
CmbGodownNo.Clear
If RsMst_Godown.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Godown.RecordCount
    CmbGodownNo.AddItem RsMst_Godown!GodownNo
    RsMst_Godown.MoveNext
Next
CmbGodownNo.Text = tmp
End Sub

Private Sub CmbGodownNo_LostFocus()
If CmbGodownNo.Text = "" Then
   CmbStack.Text = ""
   CmbLot.Text = ""
   Check1.Value = 0
   mGodownID = Null
   Exit Sub
End If
RsMst_Godown.MoveFirst
RsMst_Godown.Find "GodownNo = '" & CmbGodownNo.Text & "'"
If RsMst_Godown.EOF Then
   MsgBox "Invalid selection "
   If CmbGodownNo.Enabled = True Then
      CmbGodownNo.SetFocus
   End If
   Exit Sub
End If
If mGodownID <> RsMst_Godown!GodownID Then
   CmbStack.Text = ""
   CmbLot.Text = ""
   Call ClearOrignalGSLDetail
End If
mGodownID = RsMst_Godown!GodownID
mOldGodownYearlyFlag = IIf(IsNull(RsMst_Godown!YearlyBillingFlg), 0, RsMst_Godown!YearlyBillingFlg)
End Sub

Private Sub CmbLot_GotFocus()
If CmbStack.Text = "" Then
  MsgBox "Please Select Stack No !!!"
  CmbStack.SetFocus
  Exit Sub
End If
tmp1 = CmbLot.Text
Call TableMst_GSL(" Where " & Gen_ExpireGSLFlag & " And  CMPID = " & mCMPID & " And GodownID = " & mGodownID & " And StackNo = '" & CmbStack.Text & "' And FGSLID_Mst_GSL_ID is Null And MadeupFlag = '0'")
tmp = "Select * from Mst_GSL  Where  " & Gen_ExpireGSLFlag & "   And  CMPID = " & mCMPID & " And GodownID = " & mGodownID & " "
tmp = tmp & " And StackNo = '" & CmbStack.Text & "' And FGSLID_Mst_GSL_ID is Null And MadeupFlag = '0' "  'Commeneted on 10 Nov 2016 for allowing pledge lot as full transfer after approval from Sankar Sir' And PledgeBags = 0 "
tmp = tmp & " Union All"
tmp = tmp & " Select * from Mst_GSL  Where  " & Gen_ExpireGSLFlag & "  And  CMPID = " & mCMPID & " And GodownID = " & mGodownID & " "
tmp = tmp & " And StackNo = '" & CmbStack.Text & "' And FGSLID_Mst_GSL_ID is Null And MadeupFlag = '1'   And Connected_Mst_GSL_ID in (Select Mst_GSL_ID From Mst_GSL Where Status = 'E' And  CMPID = " & mCMPID & " And GodownID = " & mGodownID & " And StackNo = '" & CmbStack.Text & "'  )"
'Commeneted on 10 Nov 2016 for allowing pledge lot as full transfer after approval from Sankar Sir'
'And PledgeBags = 0
Call TmpTable

CmbLot.Clear
'If RsMst_GSL.RecordCount = 0 Then
'   MsgBox "No Record found !!!! "
'   Exit Sub
'End If
'
'For I = 1 To RsMst_GSL.RecordCount
'    CmbLot.AddItem RsMst_GSL!LotNo
'    RsMst_GSL.MoveNext
'Next I

If TmpRs.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If

For I = 1 To TmpRs.RecordCount
    CmbLot.AddItem TmpRs!LotNo
    TmpRs.MoveNext
Next I

CmbLot.Text = tmp1
End Sub

Private Sub CmbLot_LostFocus()
If CmbLot.Text = "" Then
   LblQty.Visible = False
   TxtQty = ""
   TxtQty.Visible = False
   Exit Sub
End If
'RsMst_GSL.MoveFirst
'RsMst_GSL.Find "LotNo = '" & CmbLot.Text & "'"
'If RsMst_GSL.EOF Then
'   MsgBox "Invalid selection "
'   CmbLot.SetFocus
'   Exit Sub
'End If
'If tmp <> RsMst_GSL!LotNo Then
'   Call ClearOrignalGSLDetail
'End If
'
'mGSLID = RsMst_GSL!GSLID

TmpRs.MoveFirst
TmpRs.Find "LotNo = '" & CmbLot.Text & "'"
If TmpRs.EOF Then
   MsgBox "Invalid selection "
   CmbLot.SetFocus
   Exit Sub
End If
If tmp1 <> TmpRs!LotNo Then
   Call ClearOrignalGSLDetail
End If

mGSLID = TmpRs!Mst_GSL_ID
mMadeupFlag = IIf(IsNull(TmpRs!MadeupFlag), 0, TmpRs!MadeupFlag)
mConnectedGSLID = IIf(IsNull(TmpRs!Connected_Mst_GSL_ID), 0, TmpRs!Connected_Mst_GSL_ID)
mContractID = IIf(IsNull(TmpRs!NContractID), 0, TmpRs!NContractID)
mOriginalDepositFlag = IIf(IsNull(TmpRs!OriginalDepositFlag), "N", TmpRs!OriginalDepositFlag)
mOriginalDeposit = False
If LCase(mOriginalDepositFlag) = "y" Then
    If Val(IIf(IsNull(TmpRs!DepositeBags), 0, TmpRs!DepositeBags)) > 0 Then
        If Val(IIf(IsNull(TmpRs!DepositeBags), 0, TmpRs!DepositeBags)) = Val(IIf(IsNull(TmpRs!BalanceBags), 0, TmpRs!BalanceBags)) Then
            mOriginalDeposit = True
        Else
            mOriginalDeposit = False
        End If
    Else
        mOriginalDeposit = False
    End If
End If

If mMadeupFlag = 0 Then
   tmp = " Select TCD.* from Txn_CxTF_Details TCD"
   tmp = tmp & "  inner join Txn_CxTF_GSL TCG on TCG.CxTF_DetailsID = TCD.CxTF_DetailsID And TCG.TxnType = TCD.TxnType And TCG.SM_Prefix = TCD.SM_Prefix"
   tmp = tmp & " where isnull(TCD.flag,'') <> 'P' and TCG.Mst_GSL_ID = " & mGSLID & " and TCG.SM_prefix = '" & TxtPreFix & "'"
ElseIf mMadeupFlag = 1 Then
   tmp = " Select TCD.* from Txn_CxTF_Details TCD"
   tmp = tmp & "  inner join Txn_CxTF_GSL TCG on TCG.CxTF_DetailsID = TCD.CxTF_DetailsID And TCG.TxnType = TCD.TxnType And TCG.SM_Prefix = TCD.SM_Prefix"
   tmp = tmp & " where isnull(TCD.flag,'') <> 'P' and (TCG.Mst_GSL_ID = " & mConnectedGSLID & " Or TCG.Mst_GSL_ID = " & mGSLID & ") and TCG.SM_prefix = '" & TxtPreFix & "'"
End If
Call Tmp5Table

If TmpRs5.RecordCount > 0 Then
   If mMadeupFlag = 0 Then
      MsgBox "One or more CDTF/CITFs for this lot not printed. Can't transfer!!!"
   ElseIf mMadeupFlag = 1 Then
      MsgBox "One or more CDTF/CITFs for Original lot of madeup lot not printed. Can't transfer!!!"
   End If
   CmbLot.SetFocus
   Exit Sub
End If

TxtGSLGUID.Text = IIf(IsNull(TmpRs!G_UID), "", TmpRs!G_UID)
TxtGSLID.Text = mGSLID
If Not IsNull(TmpRs!ExchangeTypeID) Then
   TxtExchangeTypeID.Text = GetExchangeName(TmpRs!ExchangeTypeID)
   ''-- 10 Aug 2016 For GSL Split
   ''-- Visible only When Non-NCDEX Exchange is available otherwise Hide
   If IIf(IsNull(TmpRs!ExchangeTypeID), 0, TmpRs!ExchangeTypeID) = 1 Then
      If TmpRs!PledgeBags = 0 Then
         Label33.Visible = True: TxtBalanceBags = "":  TxtBalanceBags.Visible = True
         mSplitBagsFlag = True
         If mMadeupFlag = 0 Then ''-- MadeupFlag = 0 then Given ConnectedGSLID Is Actual Madeup Lot
            Label31.Visible = True: TxtTransferBags = "": TxtTransferBags.Visible = True
            Label32.Visible = True: TxtTransferQty = "":  TxtTransferQty.Visible = True
            If mConnectedGSLID > 0 Then ''-- If Madeup lot is given then allow to Tranfer
                Label35.Visible = True: TxtMadeupBags = "":   TxtMadeupBags.Visible = True
                Label38.Visible = True: TxtBalanceMadeupBags = "":  TxtBalanceMadeupBags.Visible = True
            Else
                Label35.Visible = False: TxtMadeupBags = "":   TxtMadeupBags.Visible = False
                Label38.Visible = False: TxtBalanceMadeupBags = "":  TxtBalanceMadeupBags.Visible = False
            End If
         ElseIf mMadeupFlag = 1 Then ''-- MadeupFlag = 1 then Given ConnectedGSLID Is Actual Lot
             Label32.Visible = False: TxtTransferQty = "":  TxtTransferQty.Visible = False
             Label33.Visible = False: TxtBalanceBags = "":  TxtBalanceBags.Visible = False
             Label35.Visible = False: TxtMadeupBags = "":   TxtMadeupBags.Visible = False
             Label38.Visible = False: TxtBalanceMadeupBags = "":  TxtBalanceMadeupBags.Visible = False
         End If
      Else
         mSplitBagsFlag = False
      End If
   Else
      mSplitBagsFlag = False
   End If
Else
    mSplitBagsFlag = False
End If
mCommodity = IIf(IsNull(TmpRs!CommodityID), 0, TmpRs!CommodityID)
TxtCommodityID.Text = GetCommodityName(IIf(IsNull(TmpRs!CommodityID), 0, TmpRs!CommodityID))
TxtGSLAvailableNofBags.Text = TmpRs!DepositeBags - TmpRs!CDTFBags
TxtGSLAvailableWgt.Text = TmpRs!DepositeWeight - TmpRs!CDTFWeight

mDepoBags = IIf(IsNull(TmpRs!DepositeBags), 0, TmpRs!DepositeBags)
mDepoWt = IIf(IsNull(TmpRs!DepositeWeight), 0, TmpRs!DepositeWeight)

If Not IsNull(TmpRs!ClientID) Then
   TxtClientID.Text = GetClientName(TmpRs!ClientID)
End If

If TmpRs!ExchangeTypeID = 3 Then
   Label9.Caption = "CDC No"
Else
   Label9.Caption = "ISIN ID"
End If
TxtISINID.Text = IIf(IsNull(TmpRs!ISINID), "", TmpRs!ISINID)
TxtPledgeNo = IIf(IsNull(TmpRs!PledgeNo), "", TmpRs!PledgeNo)
TxtPledgeBags = IIf(IsNull(TmpRs!PledgeBags), "", TmpRs!PledgeBags)
TxtPledgeWeight = IIf(IsNull(TmpRs!PledgeWeight), "", TmpRs!PledgeWeight)
TxtCDTFNo = IIf(IsNull(TmpRs!CxTF_DetailsID), "", TmpRs!CxTF_DetailsID)
TxtCDTFDate = IIf(IsNull(TmpRs!CDTFDate), "", TmpRs!CDTFDate)
TxtBGSLID = IIf(IsNull(TmpRs!BGSLID_Mst_GSL_ID), "", TmpRs!BGSLID_Mst_GSL_ID)
TxtFGSLID = IIf(IsNull(TmpRs!FGSLID_Mst_GSL_ID), "", TmpRs!FGSLID_Mst_GSL_ID)
Check1.Value = TmpRs!SpillageFlag
TxtFlag = TmpRs!Status

If mMadeupFlag = 0 Then
   LblQty.Visible = False
   TxtQty = ""
   TxtQty.Visible = False
   TxtQty.Locked = True
ElseIf mMadeupFlag = 1 Then
   LblQty.Visible = True
   TxtQty = ""
   TxtQty.Visible = True
   TxtQty.Locked = False
End If

If IIf(IsNull(TmpRs!ExchangeTypeID), 0, TmpRs!ExchangeTypeID) = 1 Then
    TxtBalanceBags = IIf(IsNull(TmpRs!BalanceBags), 0, TmpRs!BalanceBags)
    TxtTransferBags = IIf(IsNull(TmpRs!BalanceBags), 0, TmpRs!BalanceBags)
    mBagSize = IIf(IsNull(TmpRs!BagSize), "", TmpRs!BagSize)
    If mBagSize = "" Then
    ''--BagSize = (DepositeWeight/DepositeBags)*1000
       mBagSize = (IIf(IsNull(TmpRs!BalanceWeight), 0, TmpRs!BalanceWeight) / TxtBalanceBags) * 1000
    End If
    'TxtTransferQty = TxtGSLAvailableWgt
    TxtTransferQty = Format(((TxtTransferBags * mBagSize) / 1000), "#####0.000")
    If mMadeupFlag = 0 And mConnectedGSLID > 0 Then
        tmp = " SELECT * FROM Mst_GSL Where Mst_GSL_ID = " & mConnectedGSLID & " And Sm_Prefix  = '" & TxtPreFix & "' And CMPID = " & mCMPID & " And CommodityID = " & mCommodity & " AND GodownID = " & mGodownID & " And StackNo = '" & CmbStack.Text & "' And FGSLID_Mst_GSL_ID is Null And MadeupFlag = '1' And Connected_Mst_GSL_ID = " & mGSLID & " "
        Call Tmp1Table
        If TmpRs1.RecordCount > 0 Then
            TxtMadeupBags = IIf(IsNull(TmpRs1!BalanceBags), 0, TmpRs1!BalanceBags)
            TxtMadeupQty = ""
            TxtBalanceMadeupBags = IIf(IsNull(TmpRs1!BalanceBags), 0, TmpRs1!BalanceBags)
        End If
    End If
End If

End Sub

Private Sub CmbStack_GotFocus()
If CmbGodownNo.Text = "" Or mGodownID = "" Then
  MsgBox "Please Select Godown No!!!"
  Exit Sub
End If
tmp = CmbStack.Text
Call TableMst_GSL(" Where CMPID = " & mCMPID & " And GodownID = " & mGodownID, "D")
CmbStack.Clear
If RsMst_GSL.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_GSL.RecordCount
    CmbStack.AddItem RsMst_GSL!StackNo
    RsMst_GSL.MoveNext
Next
CmbStack.Text = tmp
End Sub

Private Sub CmbStack_LostFocus()
If CmbStack.Text = "" Then
   CmbLot.Text = ""
   Exit Sub
End If
RsMst_GSL.MoveFirst
RsMst_GSL.Find "StackNo = '" & CmbStack.Text & "'"
If RsMst_GSL.EOF Then
   MsgBox "Invalid selection "
   CmbStack.SetFocus
   Exit Sub
End If
If tmp <> CmbStack.Text Then
   CmbLot.Text = ""
   Call ClearOrignalGSLDetail
End If
mCmbStack = CmbStack.Text
End Sub
Private Sub ClearOrignalGSLDetail()

TxtGSLID = ""
TxtGSLGUID = ""
TxtExchangeTypeID = ""
TxtCommodityID = ""
TxtGSLAvailableNofBags = ""
TxtGSLAvailableWgt = ""
TxtDepositBags = ""
TxtDepositWeight = ""
TxtClientID = ""
TxtISINID.Text = ""
TxtPledgeNo = ""
TxtPledgeBags = ""
TxtPledgeWeight = ""
TxtCDTFNo = ""
TxtCDTFDate = ""
TxtBGSLID = ""
TxtFGSLID = ""
TxtFlag = ""
Check1.Value = 0

''-- 10 Aug 2016 For GSL Split
Label31.Visible = False: TxtTransferBags = "": TxtTransferBags.Visible = False
Label32.Visible = False: TxtTransferQty = "":  TxtTransferQty.Visible = False
Label33.Visible = False: TxtBalanceBags = "":  TxtBalanceBags.Visible = False
Label35.Visible = False: TxtMadeupBags = "":   TxtMadeupBags.Visible = False
'Label36.Visible = False: TxtMadeupQty = "":    TxtMadeupQty.Visible = False
Label38.Visible = False: TxtBalanceMadeupBags = "":  TxtBalanceMadeupBags.Visible = False
''--------------
End Sub
Private Sub ClearNewGSLDetail()
TxtNewGSLID = ""
TxtNewGSLGUID = ""
TxtNewExchangeTypeID = ""
TxtNewCommodityID = ""
TxtNewGSLAvailableNofBags = 0
TxtNewGSLAvailableWgt = 0
TxtNewClientID = ""
Check2.Value = 0
TxtNewFlag = ""
TxtTransferedBags = ""
TxtTransferedQty = ""
End Sub

Private Sub CmbNewCMPName_GotFocus()
If CmbNewLocationID.Text = "" Then
   MsgBox "Select Location!!!"
   'CmbNewLocationID.SetFocus
   Exit Sub
End If
tmp = CmbNewCMPName.Text
Call TableMst_CMP(" Where CloseDate is null And LocationID = " & mNewLocationID)
'Call TableMst_CMP
If RsMst_CMP.RecordCount > 0 Then
   RsMst_CMP.MoveFirst
   CmbNewCMPName.Clear
   If RsMst_CMP.RecordCount = 0 Then
      MsgBox "No Record found !!!! "
      Exit Sub
   End If
   For I = 1 To RsMst_CMP.RecordCount
       CmbNewCMPName.AddItem RsMst_CMP!CMPName
       RsMst_CMP.MoveNext
   Next
   CmbNewCMPName.Text = tmp
Else
   MsgBox "No record found!!!"
   CmbNewCMPName.SetFocus
   Exit Sub
End If
End Sub
'----
Private Sub CmbNewCMPName_LostFocus()
If CmbNewCMPName.Text = "" Then
   CmbNewGodownNo.Text = ""
   CmbNewStack.Text = ""
   CmbNewLot.Text = ""
   TxtNewCurrentDate = ""
    mNewCMPID = 0
   Exit Sub
End If

RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPName = '" & CmbNewCMPName.Text & "'"
If RsMst_CMP.EOF Then
   MsgBox "Invalid selection "
   CmbNewCMPName.SetFocus
   Exit Sub
End If
If mNewCMPID <> RsMst_CMP!CMPID Then
   CmbNewGodownNo.Text = ""
   CmbNewStack.Text = ""
   CmbNewLot.Text = ""
   TxtNewCurrentDate = ""
   Call ClearNewGSLDetail
End If
mNewCMPID = RsMst_CMP!CMPID
TxtNewCurrentDate = IIf(IsNull(RsMst_CMP!CurrentDate), "", RsMst_CMP!CurrentDate)
'TxtNewPreFix = GetPreFix(RsMst_CMP!CMPID, "P")
End Sub

Private Sub CmbNewGodownNo_GotFocus()
If CmbNewCMPName.Text = "" Then
   MsgBox "Select CMP !!!! "
   CmbNewCMPName.SetFocus
   Exit Sub
End If
tmp = CmbNewGodownNo.Text
Call TableMst_Godown("Where CloseDate is null And CMPID = " & mNewCMPID)
CmbNewGodownNo.Clear
If RsMst_Godown.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Godown.RecordCount
    CmbNewGodownNo.AddItem RsMst_Godown!GodownNo
    RsMst_Godown.MoveNext
Next
CmbNewGodownNo.Text = tmp

End Sub
'----
Private Sub CmbNewGodownNo_LostFocus()
If CmbNewGodownNo.Text = "" Then
   mNewGodownID = Null
   CmbNewStack.Text = ""
   CmbNewLot.Text = ""
   Exit Sub
End If
RsMst_Godown.MoveFirst
RsMst_Godown.Find "GodownNo = '" & CmbNewGodownNo.Text & "'"
If RsMst_Godown.EOF Then
   MsgBox "Invalid selection "
   If CmbNewGodownNo.Enabled = True Then
      CmbNewGodownNo.SetFocus
   End If
   Exit Sub
End If
If mNewGodownID <> RsMst_Godown!GodownID Then
   Call ClearNewGSLDetail
   CmbNewStack.Text = ""
   CmbNewLot.Text = ""
End If
mNewGodownID = RsMst_Godown!GodownID
mNewGodownYearlyFlag = IIf(IsNull(RsMst_Godown!YearlyBillingFlg), 0, RsMst_Godown!YearlyBillingFlg)
End Sub
Private Sub CmbNewStack_GotFocus()
If CmbNewGodownNo.Text = "" Or mNewGodownID = "" Then
  MsgBox "Please Select Godown No!!!"
  Exit Sub
End If
tmp = CmbNewStack.Text
'Call TableMst_GSL(" Where CMPID = " & mNewCMPID & " And GodownID = " & mNewGodownID, "D")

If mMadeupFlag = 0 Then
   Call TableMst_GSL(" Where (SpillageFlag = 1 or Status = 'B') And  CMPID = " & mNewCMPID & " And FGSLID_Mst_GSL_ID is Null And GodownID = " & mNewGodownID, "D")
ElseIf mMadeupFlag = 1 Then
   Call TableMst_GSL(" Where SpillageFlag = 1 And CMPID = " & mNewCMPID & " And FGSLID_Mst_GSL_ID is Null And GodownID = " & mNewGodownID, "D")
End If
CmbNewStack.Clear
If RsMst_GSL.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_GSL.RecordCount
    CmbNewStack.AddItem RsMst_GSL!StackNo
    RsMst_GSL.MoveNext
Next
CmbNewStack.Text = tmp
End Sub

Private Sub CmbNewStack_LostFocus()
If CmbNewStack.Text = "" Then
   CmbNewLot.Text = ""
   Exit Sub
End If
RsMst_GSL.MoveFirst
RsMst_GSL.Find "StackNo = '" & CmbNewStack.Text & "'"
If RsMst_GSL.EOF Then
   MsgBox "Invalid selection "
   CmbNewStack.SetFocus
   Exit Sub
End If
If tmp <> CmbNewStack.Text Then
   Call ClearNewGSLDetail
   CmbNewLot.Text = ""
End If
mCmbNewStack = CmbNewStack.Text
End Sub

Private Sub CmbNewLot_GotFocus()
If CmbNewStack.Text = "" Then
  MsgBox "Please Select Stack No !!!"
  CmbNewStack.SetFocus
  Exit Sub
End If
'And FGSLID is Null
'And BGSLID is Null
tmp = CmbNewLot.Text
If mMadeupFlag = 0 Then
   Call TableMst_GSL(" Where (SpillageFlag = 1 or Status = 'B') And  CMPID = " & mNewCMPID & " And GodownID = " & mNewGodownID & " And StackNo = '" & CmbNewStack.Text & "' And FGSLID_Mst_GSL_ID is Null")
ElseIf mMadeupFlag = 1 Then
   Call TableMst_GSL(" Where SpillageFlag = 1 And  CMPID = " & mNewCMPID & " And GodownID = " & mNewGodownID & " And StackNo = '" & CmbNewStack.Text & "' And FGSLID_Mst_GSL_ID is Null")
End If
CmbNewLot.Clear
If RsMst_GSL.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If

For I = 1 To RsMst_GSL.RecordCount
    CmbNewLot.AddItem RsMst_GSL!LotNo
    RsMst_GSL.MoveNext
Next I

CmbNewLot.Text = tmp
End Sub
Private Sub CmbNewLot_LostFocus()
If CmbNewLot.Text = "" Then
   Exit Sub
End If
RsMst_GSL.MoveFirst
RsMst_GSL.Find "LotNo = '" & CmbNewLot.Text & "'"
If RsMst_GSL.EOF Then
   MsgBox "Invalid selection "
   CmbNewLot.SetFocus
   Exit Sub
End If

If tmp <> RsMst_GSL!LotNo Then
   Call ClearNewGSLDetail
End If

mNewGSLID = RsMst_GSL!Mst_GSL_ID

TxtNewGSLGUID.Text = IIf(IsNull(RsMst_GSL!G_UID), "", RsMst_GSL!G_UID)
TxtNewGSLID.Text = mNewGSLID
If RsMst_GSL!ExchangeTypeID <> 0 Then
   TxtNewExchangeTypeID.Text = GetExchangeName(RsMst_GSL!ExchangeTypeID)
End If
If IsNull(RsMst_GSL!CommodityID) = False Then
   mNewCommodity = RsMst_GSL!CommodityID
   TxtNewCommodityID.Text = GetCommodityName(RsMst_GSL!CommodityID)
Else
   mNewCommodity = 0
End If
TxtNewGSLAvailableNofBags.Text = RsMst_GSL!DepositeBags - RsMst_GSL!CDTFBags
TxtNewGSLAvailableWgt.Text = RsMst_GSL!DepositeWeight - RsMst_GSL!CDTFWeight
If IsNull(RsMst_GSL!ClientID) = False Or RsMst_GSL!ClientID <> 0 Then
   TxtNewClientID.Text = GetClientName(RsMst_GSL!ClientID)
Else
   TxtNewClientID.Text = 0
End If
Check2.Value = RsMst_GSL!SpillageFlag
TxtNewFlag = RsMst_GSL!Status

End Sub
Private Sub CmdCMPName_Click()
FrmMst_CMP.Show
End Sub

Private Sub CmdGodownNo_Click()
Gen_GSLID(0) = False
Gen_GSLID(1) = mCMPID
Gen_GSLID(2) = mGodownID
Gen_GSLID(3) = CmbStack.Text
Gen_GSLID(4) = CmbLot.Text
Gen_GSLID(5) = CmbCMPName.Text
Gen_GSLID(6) = CmbGodownNo.Text
FrmMst_GSL.Show
End Sub

Private Sub CmdLocationID_Click()
FrmMst_Location.Show
End Sub

Private Sub CmdNewCMPName_Click()
FrmMst_CMP.Show
End Sub

Private Sub CmdNewGodownNo_Click()
Gen_GSLID(0) = False
Gen_GSLID(1) = mNewCMPID
Gen_GSLID(2) = mNewGodownID
Gen_GSLID(3) = CmbNewStack.Text
Gen_GSLID(4) = CmbNewLot.Text
Gen_GSLID(5) = CmbNewCMPName.Text
Gen_GSLID(6) = CmbNewGodownNo.Text
FrmMst_GSL.Show
End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub LastCmd_Click()
RsTxn_GSL_Transfer.MoveLast
Call Indata
'TmpRs.MoveLast
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
RsTxn_GSL_Transfer.MoveFirst
RsTxn_GSL_Transfer.Find "G_UID = '" & MSHFlexGrid1.TextMatrix(I, 13) & "'"
If Not RsTxn_GSL_Transfer.EOF Then
   Call Indata
   Frame0.Visible = False
End If
End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsTxn_GSL_Transfer.MoveNext
If RsTxn_GSL_Transfer.EOF Then
   RsTxn_GSL_Transfer.MoveLast
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
RsTxn_GSL_Transfer.MovePrevious

If RsTxn_GSL_Transfer.BOF Then
   RsTxn_GSL_Transfer.MoveFirst
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

'If TxtTransferDate.Value > Date Then
'   MsgBox "Future Transfer date not allowed!!!"
'   TxtTransferDate.SetFocus
'   Exit Sub
'End If
'If TxtTransferDate.Value > CDate(TxtCurrentDate) Then
'   MsgBox "Transfer date must be less than or same as BOD of Original CMP!!!"
'   TxtTransferDate.SetFocus
'   Exit Sub
'End If
'If TxtTransferDate.Value > CDate(TxtNewCurrentDate) + 1 Then
'   MsgBox "Transfer date must be less than or one day more than BOD of New CMP!!!"
'   TxtTransferDate.SetFocus
'   Exit Sub
'End If

If CmbLocationID.Text = "" Then
   MsgBox "Blank Location Not Allowed!!!"
   CmbLocationID.SetFocus
   Exit Sub
End If
If CmbCMPName.Text = "" Then
   MsgBox "Blank Original CMP Name Not Allowed!!!"
   CmbCMPName.SetFocus
   Exit Sub
End If
If CmbGodownNo.Text = "" Then
   MsgBox "Blank Original Godown Not Allowed!!!"
   CmbGodownNo.SetFocus
   Exit Sub
End If
If CmbStack.Text = "" Then
   MsgBox "Blank Original Stack Not Allowed!!!"
   CmbStack.SetFocus
   Exit Sub
End If
If CmbLot.Text = "" Then
   MsgBox "Blank Original Lot Not Allowed!!!"
   CmbLot.SetFocus
   Exit Sub
End If
If Val(TxtGSLAvailableNofBags) = 0 And Val(TxtGSLAvailableWgt) = 0 Then
   MsgBox "No Bags and Qty available for transfer!!! "
   CmbLot.SetFocus
   Exit Sub
End If

'If Val(TxtGSLAvailableWgt) = 0 Then
'   MsgBox "No Quantity available for transfer!!! "
'   CmbLot.SetFocus
'   Exit Sub
'End If
'If Val(TxtDepositBags.Text) = 0 Then
'   MsgBox "Blank No Of Bags Not Allowed!!! "
'   TxtDepositBags.SetFocus
'   Exit Sub
'End If
'If Val(TxtDepositWeight.Text) = 0 Then
'   MsgBox "Blank Original Quantity Not Allowed!!! "
'   TxtDepositWeight.SetFocus
'   Exit Sub
'End If

If CmbNewLocationID.Text = "" Then
   MsgBox "Blank New Location Not Allowed!!!"
   CmbNewLocationID.SetFocus
   Exit Sub
End If
If CmbNewCMPName.Text = "" Then
   MsgBox "Blank New CMP Name Not Allowed!!!"
   CmbNewCMPName.SetFocus
   Exit Sub
End If
If CmbNewGodownNo.Text = "" Then
   MsgBox "Blank New Godown Not Allowed!!!"
   CmbNewGodownNo.SetFocus
   Exit Sub
End If
If CmbNewStack.Text = "" Then
   MsgBox "Blank New Stack Not Allowed!!!"
   CmbNewStack.SetFocus
   Exit Sub
End If
If CmbNewLot.Text = "" Then
   MsgBox "Blank New Lot Not Allowed!!!"
   CmbNewLot.SetFocus
   Exit Sub
End If
If TxtNewExchangeTypeID <> "" Then
   MsgBox "New GSL must be Empty!!!"
   CmbNewLot.SetFocus
   Exit Sub
End If

''-- 10 Aug 2016
If mSplitBagsFlag = True Then
    If Val(TxtTransferBags) <= 0 Then
        MsgBox "Empty/Nagative Transfer Bags not allowed."
        TxtTransferBags.SetFocus
        Exit Sub
    End If
    If Val(TxtTransferBags) > Val(TxtBalanceBags) Then
        MsgBox "More than Balance Bags transfer not allowed."
        TxtTransferBags.SetFocus
        Exit Sub
    End If
    If mMadeupFlag = 0 And mConnectedGSLID > 0 Then
        If Val(TxtMadeupBags) < 0 Then
            MsgBox "Nagative Transfer Madeup Bags not allowed."
            TxtMadeupBags.SetFocus
            Exit Sub
        End If
        If Val(TxtMadeupBags) > Val(TxtBalanceMadeupBags) Then
            MsgBox "More than Balance Madeup Bags transfer not allowed."
            TxtMadeupBags.SetFocus
            Exit Sub
        End If
        
        ''-- If Bags are full Transfer and Madeup bags are connected to it.
        ''-- Then User has to Tarnsfer all Madeup lots otherwise partial Transfer are allowed without madeup bags
        If Val(TxtBalanceBags) = Val(TxtTransferBags) And (Val(TxtBalanceMadeupBags) > 0 And Val(TxtMadeupBags) < Val(TxtBalanceMadeupBags)) Then
            MsgBox "Transfer all the Balance Bags without Madeup Bags or Less than Balance Madeup Bags are not allowed."
            TxtMadeupBags.SetFocus
            Exit Sub
        End If
        ''-- If Normal Bags are not get Transfer Only Madeup is Try to Transfer Then not Allow
        If Val(TxtTransferBags) = 0 And (Val(TxtBalanceMadeupBags) > 0 And Val(TxtMadeupBags) > 0) Then
            MsgBox "Only Made Up Bags Transfer not allowed."
            TxtTransferBags.SetFocus
            Exit Sub
        End If
    End If
    ''-- If New Blank GSL are Spillage then All Bags are get to Transfer , Partial Trnsfer not allowed
    If Check2.Value = vbChecked Then
        If (Val(TxtBalanceBags) <> Val(TxtTransferBags)) Or (Val(TxtBalanceMadeupBags) <> Val(TxtMadeupBags)) Then
            MsgBox "Your New GSL is Spillage so,Partially Transfer not allowed."
            TxtTransferBags.SetFocus
            Exit Sub
        End If
    End If
End If

If TxtNewCommodityID <> "" Then
   If Check1.Value = 1 And Check2.Value = 1 Then
      If mCommodity <> mNewCommodity Then
         MsgBox "New GSL must contain of same commodity !!!"
         CmbNewLot.SetFocus
         Exit Sub
      End If
   End If
   If Check1.Value = 0 And Check2.Value = 1 Then
      If mCommodity <> mNewCommodity Then
         MsgBox "New GSL must contain of same commodity !!!"
         CmbNewLot.SetFocus
         Exit Sub
      End If
   End If
End If
If TxtNewClientID <> 0 Then
   MsgBox "New GSL must be Empty!!!"
   CmbNewLot.SetFocus
   Exit Sub
End If

If Check2.Value = 0 Then
   If Val(TxtNewGSLAvailableNofBags) > 0 Then
      MsgBox "New GSL must be Empty!!!"
      CmbNewLot.SetFocus
      Exit Sub
   End If
   If Val(TxtNewGSLAvailableNofBags) > 0 Then
      MsgBox "New GSL must be Empty!!!"
      CmbNewLot.SetFocus
      Exit Sub
   End If
End If
If TxtGSLGUID = TxtNewGSLGUID Then
   MsgBox "Transfer to same GSL not Allowed!!!"
   CmbNewLot.SetFocus
   Exit Sub
End If

If Check1.Value <> Check2.Value Then
   If Check1.Value = 1 And Check2.Value = 0 Then
      MsgBox "Transfer from spillage to Non spillage lot not allowed!!!"
      CmbNewLot.SetFocus
      Exit Sub
   End If
End If

If LCase(TxtExchangeTypeID) = "non ncdex" Then
   If mMadeupFlag = 0 Then
      If Check1.Value = 0 And Check2.Value = 1 Then
         If LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "headcmg" And LCase(Gen_UserRole) <> "superadmin" Then
            If Val(TxtGSLAvailableNofBags) >= Format((Val(mDepoBags) * (Val(Val(Gen_BlockValueNNTransfer)) / 100)), "#######0") Then
               MsgBox "Transfer Bags are " & TxtGSLAvailableNofBags & " which is more than " & Gen_BlockValueNNTransfer & "% of deposit bags " & mDepoBags & " . Can't transfer!!!"
               Exit Sub
            End If
         End If
      End If
   End If
End If

If mNewGodownYearlyFlag <> mOldGodownYearlyFlag Then
   MsgBox "Yearly billing flag should be same for GSL Transfer!!!"
   Exit Sub
End If

If LCase(TxtExchangeTypeID) = "ncdex" Then
   If mMadeupFlag = 0 Then
      If Check1.Value = 0 And Check2.Value = 1 Then
         If LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "headcmg" And LCase(Gen_UserRole) <> "superadmin" Then
            If Val(TxtGSLAvailableNofBags) >= Format((Val(mDepoBags) * (Val(Gen_BlockValueNCDEXTransfer) / 100)), "#######0") Then
               MsgBox "Transfer Bags are " & TxtGSLAvailableNofBags & " which is more than " & Gen_BlockValueNCDEXTransfer & "% of deposit bags " & mDepoBags & " . Can't transfer!!!"
               Exit Sub
            End If
         End If
      End If
   End If
End If

If mMadeupFlag = 1 Then
   If TxtQty.Visible = True Then
      If Val(TxtQty) <= 0 Then
         MsgBox "Quantity to be transferred should be greater than zero!!!"
         TxtQty.SetFocus
         Exit Sub
      End If
   End If
End If
If TxtRemark.Text = "" Then
   MsgBox "Blank Remark Not Allowed!!!"
   TxtRemark.SetFocus
   Exit Sub
End If

If mGodownID <> mNewGodownID Then
   If CheckMappedReservationContract = False Then
      MsgBox "Transferee godown  is mapped to Reservation Contract(s). Kindly inform CMG team to map  the same to new godown before saving this Transfer entry!!!"
      Exit Sub
   End If
End If



mOrignalFullTransfer = False
If mOriginalDeposit = True Then
    If mSplitBagsFlag = True Then
        If Val(TxtTransferBags) - Val(TxtBalanceBags) = 0 Then
            mOrignalFullTransfer = True
        Else
            mOrignalFullTransfer = False
        End If
    Else
        mOrignalFullTransfer = False
    End If
End If

If Edit_Flg = "A" Then
   RsTxn_GSL_Transfer.AddNew
'   TxtTransferID = GetGSLTransferID
   RsTxn_GSL_Transfer!G_UID = GetGUID
Else
'   RsTxn_ISINRevalidation.MoveFirst
'   RsTxn_ISINRevalidation.Find " ISINRevalID= " & TxtISINRevalReqNo.Text
   'TxtRevalidationID.Text = RsTxn_ISINRevalidation!ISINRevalID
   'If Edit_Flg = "E" Then
    '  Call DeleteGridData
   'End If
End If

'RsTxn_GSL_Transfer!GSLTransferID = TxtTransferID 'GetGSLTransferID
RsTxn_GSL_Transfer!GSLTransferDate = TxtTransferDate.Value
RsTxn_GSL_Transfer!Old_Mst_GSL_ID = TxtGSLID.Text
RsTxn_GSL_Transfer!New_Mst_GSL_ID = TxtNewGSLID.Text
RsTxn_GSL_Transfer!SM_Prefix = TxtPreFix
''commented by rakesh on 15/03/2013(as new location added)
'RsTxn_GSL_Transfer!NewSM_Prefix = TxtPreFix
RsTxn_GSL_Transfer!NewSM_Prefix = mNewSM_Prefix

''-- On 16 Aug 2016 For Split Transfer
If mSplitBagsFlag = False Then
    RsTxn_GSL_Transfer!TransferedBags = Val(TxtGSLAvailableNofBags)
    If TxtQty.Visible = True And mMadeupFlag = 1 Then
       RsTxn_GSL_Transfer!TransferedWeight = Val(TxtQty) 'TxtQty = IIf(IsNull(RsTxn_GSL_Transfer!TransferedWeight), 0, RsTxn_GSL_Transfer!TransferedWeight)
    Else
       RsTxn_GSL_Transfer!TransferedWeight = Val(TxtGSLAvailableWgt)
    End If
Else
    RsTxn_GSL_Transfer!TransferedBags = Val(TxtTransferBags)
    If TxtQty.Visible = True And mMadeupFlag = 1 Then
       RsTxn_GSL_Transfer!TransferedWeight = Val(TxtQty)
    Else
       RsTxn_GSL_Transfer!TransferedWeight = Val(TxtTransferQty)
    End If
    If mMadeupFlag = 0 And mConnectedGSLID > 0 And Val(TxtMadeupBags) > 0 Then
        RsTxn_GSL_Transfer!MadeupBags = Val(TxtMadeupBags)
    End If
End If

RsTxn_GSL_Transfer!Flag = "T" 'TxtFlag.Text
RsTxn_GSL_Transfer!Remarks = TxtRemark.Text

If IsNull(RsTxn_GSL_Transfer!CreatedBy) = True Or RsTxn_GSL_Transfer!CreatedBy = "" Then
   RsTxn_GSL_Transfer!CreatedBy = Gen_UserLoginID
   RsTxn_GSL_Transfer!CreatedDate = Now
Else
   RsTxn_GSL_Transfer!UpdatedBy = Gen_UserLoginID
   RsTxn_GSL_Transfer!UpdatedDate = Now
End If
RsTxn_GSL_Transfer.Update
Call SaveGSL
If RsTxn_GSL_Transfer.RecordCount > 0 Then
   RsTxn_GSL_Transfer.MoveLast
End If

End Sub

Private Function GetGSLTransferID() As Double
Dim Retval As Double
tmp = "Select Max(GSLTransferID) from Txn_GSL_Transfer"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
TmpRs.Close
GetGSLTransferID = Retval
End Function

Private Sub SaveGSL()
Call TableMst_GSL
'RsMst_GSL.MoveFirst

RsMst_GSL.Find ("G_UID= '" & TxtNewGSLGUID & "'")
If Not RsMst_GSL.EOF Then
   tmp = "Select * from Mst_GSL Where G_UID = '" & TxtGSLGUID & "'"
   Call Tmp4Table
   If TmpRs4.RecordCount > 0 Then
      ''-- On 10 Aug 2016 ''-- As per New Requirement if Non-Ncdex Bags are Transfer/Split
      If mSplitBagsFlag = True Then
      
        RsMst_GSL!CommodityID = TmpRs4!CommodityID
        'RsMst_GSL!CDTFBags = IIf(IsNull(RsMst_GSL!CDTFBags), 0, RsMst_GSL!CDTFBags) + Val(TxtGSLAvailableNofBags.Text)
        If (Val(TxtBalanceBags) - Val(TxtTransferBags) = 0) Then
            RsMst_GSL!DepositeWeight = IIf(IsNull(RsMst_GSL!DepositeWeight), 0, RsMst_GSL!DepositeWeight) + Val(TmpRs4!DepositeWeight)
            RsMst_GSL!DepositeBags = IIf(IsNull(RsMst_GSL!DepositeBags), 0, RsMst_GSL!DepositeBags) + Val(TmpRs4!DepositeBags)
            RsMst_GSL!CDTFBags = IIf(IsNull(RsMst_GSL!CDTFBags), 0, RsMst_GSL!CDTFBags) + Val(TmpRs4!CDTFBags)
        Else
            RsMst_GSL!DepositeWeight = IIf(IsNull(RsMst_GSL!DepositeWeight), 0, RsMst_GSL!DepositeWeight) + Val(TxtTransferQty)
            RsMst_GSL!DepositeBags = IIf(IsNull(RsMst_GSL!DepositeBags), 0, RsMst_GSL!DepositeBags) + Val(TxtTransferBags)
            RsMst_GSL!CDTFBags = 0 'IIf(IsNull(RsMst_GSL!CDTFBags), 0, RsMst_GSL!CDTFBags) + Val(TmpRs4!CDTFBags)
        End If
        'RsMst_GSL!CDTFBags = 0 'IIf(IsNull(RsMst_GSL!CDTFBags), 0, RsMst_GSL!CDTFBags) + Val(TmpRs4!CDTFBags)
        RsMst_GSL!BalanceBags = IIf(IsNull(RsMst_GSL!BalanceBags), 0, RsMst_GSL!BalanceBags) + Val(TxtTransferBags)
        
        If Check2.Value = 1 Then
           If TmpRs4!MadeupFlag = 0 Then
              RsMst_GSL!CDTFWeight = IIf(IsNull(RsMst_GSL!CDTFWeight), 0, RsMst_GSL!CDTFWeight) + Val(TmpRs4!CDTFWeight) 'TmpRs4!CDTFWeight
              RsMst_GSL!BalanceWeight = IIf(IsNull(RsMst_GSL!BalanceWeight), 0, RsMst_GSL!BalanceWeight) + Val(TmpRs4!BalanceWeight) 'TmpRs4!BalanceWeight
           Else
              RsMst_GSL!DepositeWeight = IIf(IsNull(RsMst_GSL!DepositeWeight), 0, RsMst_GSL!DepositeWeight) + Val(TxtQty) 'TmpRs4!CDTFWeight
              RsMst_GSL!BalanceWeight = IIf(IsNull(RsMst_GSL!BalanceWeight), 0, RsMst_GSL!BalanceWeight) + Val(TxtQty) 'TmpRs4!BalanceWeight
           End If
        Else
           If (Val(TxtBalanceBags) - Val(TxtTransferBags) = 0) Then
               RsMst_GSL!CDTFWeight = IIf(IsNull(RsMst_GSL!CDTFWeight), 0, RsMst_GSL!CDTFWeight) + Val(TmpRs4!CDTFWeight)
           Else
              RsMst_GSL!CDTFWeight = 0 'IIf(IsNull(RsMst_GSL!CDTFWeight), 0, RsMst_GSL!CDTFWeight) + Val(TmpRs4!CDTFWeight)             'TmpRs4!CDTFWeight
           End If
           RsMst_GSL!BalanceWeight = IIf(IsNull(RsMst_GSL!BalanceWeight), 0, RsMst_GSL!BalanceWeight) + Val(TxtTransferQty) ''Val(TmpRs4!BalanceWeight) 'TmpRs4!BalanceWeight
        End If
        RsMst_GSL!Status = TmpRs4!Status
        RsMst_GSL!SCMFlag = TmpRs4!SCMFlag
        If Check2.Value = 0 Then
           
           ''-- Commented for Split Transfer on 16 Aug 2016
           ''RsMst_GSL!WHRBags = IIf(IsNull(RsMst_GSL!WHRBags), 0, RsMst_GSL!WHRBags) + Val(TmpRs4!WHRBags) 'TmpRs4!WHRBags
           ''RsMst_GSL!WHRWeight = IIf(IsNull(RsMst_GSL!WHRWeight), 0, RsMst_GSL!WHRWeight) + Val(TmpRs4!WHRWeight) 'TmpRs4!WHRWeight
           
           ''-- WHR Bags/Weight are 0 if Lots are Split
           ''-- WHR Bags/Weight are as same as Orignal Bags if the All Balance Bags are Transefed
           If (Val(TxtBalanceBags) - Val(TxtTransferBags) = 0) Then
              RsMst_GSL!WHRBags = IIf(IsNull(RsMst_GSL!WHRBags), 0, RsMst_GSL!WHRBags) + Val(TmpRs4!WHRBags) 'TmpRs4!WHRBags
              RsMst_GSL!WHRWeight = IIf(IsNull(RsMst_GSL!WHRWeight), 0, RsMst_GSL!WHRWeight) + Val(TmpRs4!WHRWeight) 'TmpRs4!WHRWeight
           Else
              If TmpRs4!WHRBags = 0 Then
                 RsMst_GSL!WHRBags = 0
                 RsMst_GSL!WHRWeight = 0
              Else
                 RsMst_GSL!WHRBags = Val(TxtTransferBags)
                 RsMst_GSL!WHRWeight = Val(TxtTransferQty)
              End If
           End If
           
           RsMst_GSL!ExchangeTypeID = TmpRs4!ExchangeTypeID
           RsMst_GSL!ClientID = TmpRs4!ClientID
           RsMst_GSL!PledgeNo = TmpRs4!PledgeNo
           RsMst_GSL!ISINID = TmpRs4!ISINID
           RsMst_GSL!CxTF_DetailsID = TmpRs4!CxTF_DetailsID
           RsMst_GSL!CDTFDate = TmpRs4!CDTFDate
           RsMst_GSL!FirstISINID = TmpRs4!FirstISINID
           RsMst_GSL!PledgeBags = TmpRs4!PledgeBags
           RsMst_GSL!PledgeWeight = TmpRs4!PledgeWeight
           RsMst_GSL!DepositAdjWt = TmpRs4!DepositAdjWt
           RsMst_GSL!WithdrawalAdjWt = TmpRs4!WithdrawalAdjWt
           RsMst_GSL!Flag1 = TmpRs4!Flag1
           RsMst_GSL!LCDTF_Date = TmpRs4!LCDTF_Date
           RsMst_GSL!GrossWt = TmpRs4!GrossWt
           RsMst_GSL!KRNo = TmpRs4!KRNo
           RsMst_GSL!KRNoBatchID = TmpRs4!KRNoBatchID
           RsMst_GSL!KRDate = TmpRs4!KRDate
           RsMst_GSL!ContractID = TmpRs4!ContractID
           RsMst_GSL!NContractID = TmpRs4!NContractID
           RsMst_GSL!BGSLID_Mst_GSL_ID = TxtGSLID.Text
           RsMst_GSL!BSM_Prefix = TxtPreFix
           RsMst_GSL!CMSPBatchID = TmpRs4!CMSPBatchID
           RsMst_GSL!CMSPBatchGeneratedBy = TmpRs4!CMSPBatchGeneratedBy
           RsMst_GSL!CMSPBatchGeneratedDate = TmpRs4!CMSPBatchGeneratedDate
           RsMst_GSL!CMSPUpdatedBy = TmpRs4!CMSPUpdatedBy
           RsMst_GSL!CMSPUpdatedDate = TmpRs4!CMSPUpdatedDate
           ''-- Added on 16 Aug 2016
           RsMst_GSL!BagSize = TmpRs4!BagSize
           RsMst_GSL!BagEquivalent = TmpRs4!BagEquivalent
           RsMst_GSL!WtQty = TmpRs4!WtQty
           RsMst_GSL!MoisturePer = TmpRs4!MoisturePer
           RsMst_GSL!CITF_MoisturePer = TmpRs4!CITF_MoisturePer
           RsMst_GSL!MoistureValue = TmpRs4!MoistureValue
           RsMst_GSL!TransactionType = TmpRs4!TransactionType
           RsMst_GSL!TransactionID = TmpRs4!TransactionID
           RsMst_GSL!TransactionDate = TmpRs4!TransactionDate
           RsMst_GSL!MAWWt = TmpRs4!MAWWt
           RsMst_GSL!WHRNo = TmpRs4!WHRNo
           'Txn_WHR_Detail_ID
           RsMst_GSL!Txn_WHR_Detail_ID = TmpRs4!Txn_WHR_Detail_ID
           RsMst_GSL!RateCardType = TmpRs4!RateCardType
           RsMst_GSL!BillingCycleID = TmpRs4!BillingCycleID
           RsMst_GSL!BillingUnit = TmpRs4!BillingUnit
           RsMst_GSL!Rate = TmpRs4!Rate
           RsMst_GSL!RateForDailyPnL = TmpRs4!RateForDailyPnL
           RsMst_GSL!DepConDID = TmpRs4!DepConDID
           RsMst_GSL!NoOfDays = TmpRs4!NoOfDays


           'RsMst_GSL!NFinFlag = TmpRs4!NFinFlag
           ''-----
           If (Val(TxtBalanceBags) - Val(TxtTransferBags) = 0) Then
                RsMst_GSL!OriginalDepositFlag = "Y"
           Else
                RsMst_GSL!OriginalDepositFlag = "N"
           End If
                  
           If TmpRs4!CxTF_DetailsID < 0 Then
              tmp = "Update Txn_RRNPWRGSL Set NewTransferSM_Prefix = '" & TxtPreFix & "',NewTransferMst_GSL_ID = " & RsMst_GSL!Mst_GSL_ID & " Where SM_Prefix = '" & TxtPreFix & "' And (New_Mst_GSL_ID = " & TmpRs4!Mst_GSL_ID & " Or NewTransferMst_GSL_ID = " & TmpRs4!Mst_GSL_ID & ")"
              Call Tmp3Table
           End If
        End If
        RsMst_GSL.Update
      
      Else
      
        RsMst_GSL!CommodityID = TmpRs4!CommodityID
        'RsMst_GSL!CDTFBags = IIf(IsNull(RsMst_GSL!CDTFBags), 0, RsMst_GSL!CDTFBags) + Val(TxtGSLAvailableNofBags.Text)
        RsMst_GSL!DepositeWeight = IIf(IsNull(RsMst_GSL!DepositeWeight), 0, RsMst_GSL!DepositeWeight) + Val(TmpRs4!DepositeWeight) 'TmpRs4!DepositeWeight
        RsMst_GSL!DepositeBags = IIf(IsNull(RsMst_GSL!DepositeBags), 0, RsMst_GSL!DepositeBags) + Val(TmpRs4!DepositeBags) 'TmpRs4!DepositeBags
        RsMst_GSL!CDTFBags = IIf(IsNull(RsMst_GSL!CDTFBags), 0, RsMst_GSL!CDTFBags) + Val(TmpRs4!CDTFBags) 'TmpRs4!CDTFBags
        RsMst_GSL!BalanceBags = IIf(IsNull(RsMst_GSL!BalanceBags), 0, RsMst_GSL!BalanceBags) + Val(TmpRs4!BalanceBags) 'TmpRs4!BalanceBags
        If Check2.Value = 1 Then
           If TmpRs4!MadeupFlag = 0 Then
              RsMst_GSL!CDTFWeight = IIf(IsNull(RsMst_GSL!CDTFWeight), 0, RsMst_GSL!CDTFWeight) + Val(TmpRs4!CDTFWeight) 'TmpRs4!CDTFWeight
              RsMst_GSL!BalanceWeight = IIf(IsNull(RsMst_GSL!BalanceWeight), 0, RsMst_GSL!BalanceWeight) + Val(TmpRs4!BalanceWeight) 'TmpRs4!BalanceWeight
           Else
              RsMst_GSL!DepositeWeight = IIf(IsNull(RsMst_GSL!DepositeWeight), 0, RsMst_GSL!DepositeWeight) + Val(TxtQty) 'TmpRs4!CDTFWeight
              RsMst_GSL!BalanceWeight = IIf(IsNull(RsMst_GSL!BalanceWeight), 0, RsMst_GSL!BalanceWeight) + Val(TxtQty) 'TmpRs4!BalanceWeight
           End If
        Else
           RsMst_GSL!CDTFWeight = IIf(IsNull(RsMst_GSL!CDTFWeight), 0, RsMst_GSL!CDTFWeight) + Val(TmpRs4!CDTFWeight) 'TmpRs4!CDTFWeight
           RsMst_GSL!BalanceWeight = IIf(IsNull(RsMst_GSL!BalanceWeight), 0, RsMst_GSL!BalanceWeight) + Val(TmpRs4!BalanceWeight) 'TmpRs4!BalanceWeight
        End If
        RsMst_GSL!Status = TmpRs4!Status
        RsMst_GSL!SCMFlag = TmpRs4!SCMFlag
        If Check2.Value = 0 Then
           RsMst_GSL!WHRNo = TmpRs4!WHRNo
           'Txn_WHR_Detail_ID
           RsMst_GSL!Txn_WHR_Detail_ID = TmpRs4!Txn_WHR_Detail_ID
           'RsMst_GSL!Txn_WHR_Detail_ID = TmpRs4!Txn_WHR_Detail_ID
           RsMst_GSL!WHRBags = IIf(IsNull(RsMst_GSL!WHRBags), 0, RsMst_GSL!WHRBags) + Val(TmpRs4!WHRBags) 'TmpRs4!WHRBags
           RsMst_GSL!WHRWeight = IIf(IsNull(RsMst_GSL!WHRWeight), 0, RsMst_GSL!WHRWeight) + Val(TmpRs4!WHRWeight) 'TmpRs4!WHRWeight
           RsMst_GSL!ExchangeTypeID = TmpRs4!ExchangeTypeID
           RsMst_GSL!ClientID = TmpRs4!ClientID
           RsMst_GSL!PledgeNo = TmpRs4!PledgeNo
           RsMst_GSL!ISINID = TmpRs4!ISINID
           RsMst_GSL!CxTF_DetailsID = TmpRs4!CxTF_DetailsID
           RsMst_GSL!CDTFDate = TmpRs4!CDTFDate
           RsMst_GSL!FirstISINID = TmpRs4!FirstISINID
           RsMst_GSL!PledgeBags = TmpRs4!PledgeBags
           RsMst_GSL!PledgeWeight = TmpRs4!PledgeWeight
           RsMst_GSL!DepositAdjWt = TmpRs4!DepositAdjWt
           RsMst_GSL!WithdrawalAdjWt = TmpRs4!WithdrawalAdjWt
           RsMst_GSL!Flag1 = TmpRs4!Flag1
           RsMst_GSL!LCDTF_Date = TmpRs4!LCDTF_Date
           RsMst_GSL!GrossWt = TmpRs4!GrossWt
           RsMst_GSL!KRNo = TmpRs4!KRNo
           RsMst_GSL!KRNoBatchID = TmpRs4!KRNoBatchID
           RsMst_GSL!KRDate = TmpRs4!KRDate
           RsMst_GSL!ContractID = TmpRs4!ContractID
           RsMst_GSL!NContractID = TmpRs4!NContractID
           RsMst_GSL!BGSLID_Mst_GSL_ID = TxtGSLID.Text
           RsMst_GSL!BSM_Prefix = TxtPreFix
           RsMst_GSL!CMSPBatchID = TmpRs4!CMSPBatchID
           RsMst_GSL!CMSPBatchGeneratedBy = TmpRs4!CMSPBatchGeneratedBy
           RsMst_GSL!CMSPBatchGeneratedDate = TmpRs4!CMSPBatchGeneratedDate
           RsMst_GSL!CMSPUpdatedBy = TmpRs4!CMSPUpdatedBy
           RsMst_GSL!CMSPUpdatedDate = TmpRs4!CMSPUpdatedDate
           ''-- Added on 16 Aug 2016
           RsMst_GSL!BagSize = TmpRs4!BagSize
           RsMst_GSL!BagEquivalent = TmpRs4!BagEquivalent
           RsMst_GSL!WtQty = TmpRs4!WtQty
           RsMst_GSL!MoisturePer = TmpRs4!MoisturePer
           RsMst_GSL!CITF_MoisturePer = TmpRs4!CITF_MoisturePer
           RsMst_GSL!MoistureValue = TmpRs4!MoistureValue
           RsMst_GSL!TransactionType = TmpRs4!TransactionType
           RsMst_GSL!TransactionID = TmpRs4!TransactionID
           RsMst_GSL!TransactionDate = TmpRs4!TransactionDate
           RsMst_GSL!MAWWt = TmpRs4!MAWWt
           RsMst_GSL!NFinFlag = TmpRs4!NFinFlag
           
           RsMst_GSL!RateCardType = TmpRs4!RateCardType
           RsMst_GSL!BillingCycleID = TmpRs4!BillingCycleID
           RsMst_GSL!BillingUnit = TmpRs4!BillingUnit
           RsMst_GSL!Rate = TmpRs4!Rate
           RsMst_GSL!RateForDailyPnL = TmpRs4!RateForDailyPnL
           RsMst_GSL!DepConDID = TmpRs4!DepConDID
           RsMst_GSL!NoOfDays = TmpRs4!NoOfDays
           ''------
           If TmpRs4!CxTF_DetailsID < 0 Then
              tmp = "Update Txn_RRNPWRGSL Set NewTransferSM_Prefix = '" & TxtPreFix & "',NewTransferMst_GSL_ID = " & RsMst_GSL!Mst_GSL_ID & " Where SM_Prefix = '" & TxtPreFix & "' And (New_Mst_GSL_ID = " & TmpRs4!Mst_GSL_ID & " Or NewTransferMst_GSL_ID = " & TmpRs4!Mst_GSL_ID & ")"
              Call Tmp3Table
           End If
        End If
        RsMst_GSL.Update
      
      End If
   End If
   
   RsMst_GSL.MoveFirst
   RsMst_GSL.Find ("G_UID= '" & TxtGSLGUID & "'")
   If Not RsMst_GSL.EOF Then
          ''-- Split Transfer
          If mSplitBagsFlag = False Then
            RsMst_GSL!CDTFBags = IIf(IsNull(RsMst_GSL!CDTFBags), 0, RsMst_GSL!CDTFBags) + Val(TxtGSLAvailableNofBags.Text)
            RsMst_GSL!BalanceBags = IIf(IsNull(RsMst_GSL!BalanceBags), 0, RsMst_GSL!BalanceBags) - Val(TxtGSLAvailableNofBags.Text)
          Else
            RsMst_GSL!CDTFBags = IIf(IsNull(RsMst_GSL!CDTFBags), 0, RsMst_GSL!CDTFBags) + Val(TxtTransferBags)
            RsMst_GSL!BalanceBags = IIf(IsNull(RsMst_GSL!BalanceBags), 0, RsMst_GSL!BalanceBags) - Val(TxtTransferBags)
            If RsMst_GSL!WHRBags > 0 Then
               RsMst_GSL!WHRBags = IIf(IsNull(RsMst_GSL!WHRBags), 0, RsMst_GSL!WHRBags) - Val(TxtTransferBags)
            End If
          End If

          If RsMst_GSL!MadeupFlag = 0 Then
             If mSplitBagsFlag = False Then
                RsMst_GSL!CDTFWeight = IIf(IsNull(RsMst_GSL!CDTFWeight), 0, RsMst_GSL!CDTFWeight) + Val(TxtGSLAvailableWgt.Text)
                RsMst_GSL!BalanceWeight = IIf(IsNull(RsMst_GSL!BalanceWeight), 0, RsMst_GSL!BalanceWeight) - Val(TxtGSLAvailableWgt.Text)
             Else
                RsMst_GSL!CDTFWeight = IIf(IsNull(RsMst_GSL!CDTFWeight), 0, RsMst_GSL!CDTFWeight) + Val(TxtTransferQty)
                RsMst_GSL!BalanceWeight = IIf(IsNull(RsMst_GSL!BalanceWeight), 0, RsMst_GSL!BalanceWeight) - Val(TxtTransferQty)
                If RsMst_GSL!WHRWeight > 0 Then
                   RsMst_GSL!WHRWeight = IIf(IsNull(RsMst_GSL!WHRWeight), 0, RsMst_GSL!WHRWeight) - Val(TxtTransferQty)
                End If
             End If
          Else
             RsMst_GSL!CDTFWeight = IIf(IsNull(RsMst_GSL!CDTFWeight), 0, RsMst_GSL!CDTFWeight) + Val(TxtQty)
             RsMst_GSL!BalanceWeight = IIf(IsNull(RsMst_GSL!BalanceWeight), 0, RsMst_GSL!BalanceWeight) - Val(TxtQty)
          End If
          
          ''-- If Transfer is full then
          If mSplitBagsFlag = True Then
            If (Val(TxtBalanceBags) - Val(TxtTransferBags) = 0) Then
                RsMst_GSL!FGSLID_Mst_GSL_ID = TxtNewGSLID
                RsMst_GSL!FSM_Prefix = mNewSM_Prefix
            End If
          Else
            RsMst_GSL!FGSLID_Mst_GSL_ID = TxtNewGSLID
            RsMst_GSL!FSM_Prefix = mNewSM_Prefix
          End If
          
          ''-- If Transfer is full then 'Z' otherwise 'O'
          If mSplitBagsFlag = True And (Val(TxtBalanceBags) - Val(TxtTransferBags) = 0) Then
            RsMst_GSL!Status = "Z"
          ElseIf (Val(TxtBalanceBags) - Val(TxtTransferBags) = 0) Then
            RsMst_GSL!Status = "Z"
          Else
            RsMst_GSL!Status = "O"
          End If
          RsMst_GSL.Update
   End If
   
   If mMadeupFlag <> 1 Then
      If mConnectedGSLID <> 0 Then
         If Check2.Value = 0 Then
            If mSplitBagsFlag = True Then
                If Val(TxtMadeupBags) > 0 Then
                    Call CreateMadeUpGSL
                End If
            Else
                Call CreateMadeUpGSL
            End If
         Else
            tmp = "Select * from Mst_GSL Where SM_Prefix = '" & TxtPreFix & "' And Mst_GSL_ID = " & mConnectedGSLID & " And MadeupFlag = '1' " 'G_UID = '" & TxtGSLGUID & "'"
            Call Tmp4Table
            RsMst_GSL.MoveFirst
            RsMst_GSL.Find ("G_UID= '" & TxtNewGSLGUID & "'")
            If Not RsMst_GSL.EOF Then
               RsMst_GSL!DepositeBags = IIf(IsNull(RsMst_GSL!DepositeBags), 0, RsMst_GSL!DepositeBags) + Val(TmpRs4!DepositeBags) 'TmpRs4!DepositeBags
               RsMst_GSL!CDTFBags = IIf(IsNull(RsMst_GSL!CDTFBags), 0, RsMst_GSL!CDTFBags) + Val(TmpRs4!CDTFBags) 'TmpRs4!CDTFBags
               RsMst_GSL!BalanceBags = IIf(IsNull(RsMst_GSL!BalanceBags), 0, RsMst_GSL!BalanceBags) + Val(TmpRs4!BalanceBags) 'TmpRs4!BalanceBags
               RsMst_GSL.Update
            End If
            TmpRs4!CDTFBags = IIf(IsNull(TmpRs4!CDTFBags), 0, TmpRs4!CDTFBags) + Val(TmpRs4!BalanceBags)
            TmpRs4!BalanceBags = IIf(IsNull(TmpRs4!BalanceBags), 0, TmpRs4!BalanceBags) - Val(TmpRs4!BalanceBags)
            TmpRs4!FGSLID_Mst_GSL_ID = TxtNewGSLID
            TmpRs4!FSM_Prefix = mNewSM_Prefix
            TmpRs4!Status = "Z"
            TmpRs4.Update
         End If
      End If
   End If
End If
    
Call Indata

End Sub

Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   If RsTxn_GSL_Transfer.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   Call Indata
   Exit Sub
End If

Frame0.Height = Me.Height - 500
Frame0.Left = Frame3.Left
Frame0.Top = Frame3.Top
Frame0.Width = Me.Width - 300
Frame0.Visible = True
MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 1400

''===========comment on 7 july 2012 for not clearing search grid=============
'Call CmdClear_Click
''===========comment on 7 july 2012 for not clearing search grid=============

End Sub

Public Sub Indata()
''-- 10 Aug 2016 For GSL Split
Label31.Visible = False: TxtTransferBags = "": TxtTransferBags.Visible = False
Label32.Visible = False: TxtTransferQty = "":  TxtTransferQty.Visible = False
Label33.Visible = False: TxtBalanceBags = "":  TxtBalanceBags.Visible = False
Label35.Visible = False: TxtMadeupBags = "":   TxtMadeupBags.Visible = False
Label38.Visible = False: TxtBalanceMadeupBags = "":  TxtBalanceMadeupBags.Visible = False
Label39.Visible = False: TxtTransferedMadeupBags = "":  TxtTransferedMadeupBags.Visible = False
''--------------

TxtRemark.Locked = True
TxtTransferID = RsTxn_GSL_Transfer!GSLTransferID
TxtTransferDate = RsTxn_GSL_Transfer!GSLTransferDate
TxtTransferDate.Enabled = False
TxtRemark = IIf(IsNull(RsTxn_GSL_Transfer!Remarks), "", RsTxn_GSL_Transfer!Remarks)
TxtTransferedBags = IIf(IsNull(RsTxn_GSL_Transfer!TransferedBags), "", RsTxn_GSL_Transfer!TransferedBags)
TxtTransferedQty = IIf(IsNull(RsTxn_GSL_Transfer!TransferedWeight), "", RsTxn_GSL_Transfer!TransferedWeight)

If Val(IIf(IsNull(RsTxn_GSL_Transfer!MadeupBags), 0, RsTxn_GSL_Transfer!MadeupBags)) > 0 Then
    Label39.Visible = True:
    TxtTransferedMadeupBags = IIf(IsNull(RsTxn_GSL_Transfer!MadeupBags), "", RsTxn_GSL_Transfer!MadeupBags)
    TxtTransferedMadeupBags.Visible = True
End If

TxtCreated = IIf(IsNull(RsTxn_GSL_Transfer!CreatedBy), "", RsTxn_GSL_Transfer!CreatedBy) & " :- " & IIf(IsNull(RsTxn_GSL_Transfer!CreatedDate), "", RsTxn_GSL_Transfer!CreatedDate)
TxtUpdated = IIf(IsNull(RsTxn_GSL_Transfer!UpdatedBy), "", RsTxn_GSL_Transfer!UpdatedBy) & " :- " & IIf(IsNull(RsTxn_GSL_Transfer!UpdatedDate), "", RsTxn_GSL_Transfer!UpdatedDate)

CmbLocationID.Visible = False
TxtLocationID.Visible = True

CmbCMPName.Visible = False
TxtCMPName.Visible = True

CmbGodownNo.Visible = False

TxtGodownNo.Visible = True

CmbStack.Visible = False
TxtStack.Visible = True

CmbLot.Visible = False
TxtLot.Visible = True

CmbNewLocationID.Visible = False
CmdNewLocationID.Visible = False
TxtNewLocationID.Visible = True

CmbNewCMPName.Visible = False
CmdNewCMPName.Visible = False
TxtNewCMPName.Visible = True

CmbNewGodownNo.Visible = False
CmdNewGodownNo.Visible = False
txtNewGodownNo.Visible = True

CmbNewStack.Visible = False
TxtNewStack.Visible = True

CmbNewLot.Visible = False
TxtNewLot.Visible = True

TxtNewLocationID = ""
TxtNewCMPName = ""
txtNewGodownNo = ""
TxtNewStack = ""
TxtNewLot = ""
TxtNewExchangeTypeID.Text = ""
TxtNewCommodityID = ""
TxtNewGSLAvailableNofBags = ""
TxtNewGSLAvailableWgt = ""
TxtNewDepositBags = ""
TxtNewDepositWeight = ""
TxtNewClientID = ""

Call TableMst_GSL(" Where SM_Prefix = '" & RsTxn_GSL_Transfer!SM_Prefix & "' ")
RsMst_GSL.Find ("Mst_GSL_ID = " & RsTxn_GSL_Transfer!Old_Mst_GSL_ID)
If Not RsMst_GSL.EOF = True Then
   TxtLocationID = GetLocationFromCMP(RsMst_GSL!CMPID)
   TxtCMPName = GetCMPName(RsMst_GSL!CMPID)
   TxtGodownNo = GetGodownName(RsMst_GSL!GodownID, "N")
   TxtStack = RsMst_GSL!StackNo
   TxtLot = RsMst_GSL!LotNo
   If Not IsNull(RsMst_GSL!ExchangeTypeID) Then
      TxtExchangeTypeID.Text = GetExchangeName(RsMst_GSL!ExchangeTypeID)
   End If
   TxtCommodityID = GetCommodityName(RsMst_GSL!CommodityID)
   TxtGSLAvailableNofBags = 0 'RsMst_GSL!BalanceBags
   TxtGSLAvailableWgt = 0 'RsMst_GSL!BalanceWeight
   TxtDepositBags = 0
   TxtDepositWeight = 0
   If Not IsNull(RsMst_GSL!ClientID) Then
      TxtClientID = GetClientName(RsMst_GSL!ClientID)
   End If
   TxtISINID.Text = IIf(IsNull(RsMst_GSL!ISINID), "", RsMst_GSL!ISINID)
   TxtPledgeNo = IIf(IsNull(RsMst_GSL!PledgeNo), "", RsMst_GSL!PledgeNo)
   TxtPledgeBags = IIf(IsNull(RsMst_GSL!PledgeBags), "", RsMst_GSL!PledgeBags)
   TxtPledgeWeight = IIf(IsNull(RsMst_GSL!PledgeWeight), "", RsMst_GSL!PledgeWeight)
   TxtCDTFNo = IIf(IsNull(RsMst_GSL!CxTF_DetailsID), "", RsMst_GSL!CxTF_DetailsID)
   TxtCDTFDate = IIf(IsNull(RsMst_GSL!CDTFDate), "", RsMst_GSL!CDTFDate)
   TxtBGSLID = IIf(IsNull(RsMst_GSL!BGSLID_Mst_GSL_ID), "", RsMst_GSL!BGSLID_Mst_GSL_ID)
   TxtFGSLID = IIf(IsNull(RsMst_GSL!FGSLID_Mst_GSL_ID), "", RsMst_GSL!FGSLID_Mst_GSL_ID)
   Check1.Value = RsMst_GSL!SpillageFlag
   TxtFlag = RsMst_GSL!Status

   mMadeupFlag = IIf(IsNull(RsMst_GSL!MadeupFlag), 0, RsMst_GSL!MadeupFlag)
    
   If mMadeupFlag = 0 Then
      LblQty.Visible = False
      TxtQty = ""
      TxtQty.Visible = False
   ElseIf mMadeupFlag = 1 Then
      LblQty.Visible = True
      TxtQty = ""
      TxtQty.Visible = True
   End If

End If

Call TableMst_GSL(" Where SM_Prefix = '" & RsTxn_GSL_Transfer!NewSM_Prefix & "' ")
RsMst_GSL.Find ("Mst_GSL_ID = " & RsTxn_GSL_Transfer!New_Mst_GSL_ID)


If Not RsMst_GSL.EOF = True Then
   TxtNewLocationID = GetLocationFromCMP(RsMst_GSL!CMPID)
   TxtNewCMPName = GetCMPName(RsMst_GSL!CMPID)
   txtNewGodownNo = GetGodownName(RsMst_GSL!GodownID, "N")
   TxtNewStack = RsMst_GSL!StackNo
   TxtNewLot = RsMst_GSL!LotNo
   If Not IsNull(RsMst_GSL!ExchangeTypeID) Then
      TxtNewExchangeTypeID.Text = GetExchangeName(RsMst_GSL!ExchangeTypeID)
   End If
   TxtNewCommodityID = GetCommodityName(RsMst_GSL!CommodityID)
   TxtNewGSLAvailableNofBags = RsMst_GSL!BalanceBags
   TxtNewGSLAvailableWgt = RsMst_GSL!BalanceWeight
   If Not IsNull(RsMst_GSL!ClientID) Then
      TxtNewClientID = GetClientName(RsMst_GSL!ClientID)
   End If
   Check2.Value = RsMst_GSL!SpillageFlag
   TxtNewFlag = RsMst_GSL!Status
   If TxtQty.Visible = True Then
      TxtQty = IIf(IsNull(RsTxn_GSL_Transfer!TransferedWeight), 0, RsTxn_GSL_Transfer!TransferedWeight)
      TxtQty.Locked = True
   End If
End If


If RsTxn_GSL_Transfer.RecordCount = 1 Then
   Call GButtonLock_1(Me, True)
Else
   Call GButtonLock(Me, True)
End If

DeleteCmd.Enabled = False
'SearchCmd.Enabled = False

EditCmd.Enabled = False

End Sub
Private Sub Grid_Head1()

MSHFlexGrid1.Clear
MSHFlexGrid1.Cols = 20
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.FixedRows = 1

MSHFlexGrid1.TextMatrix(0, 0) = "Transfer ID"
MSHFlexGrid1.TextMatrix(0, 1) = "Transfer Date"
MSHFlexGrid1.TextMatrix(0, 2) = "New Location Name"
MSHFlexGrid1.TextMatrix(0, 3) = "New CMP Name"
MSHFlexGrid1.TextMatrix(0, 4) = "New G~S~L"
MSHFlexGrid1.TextMatrix(0, 5) = "Balance Bags"
MSHFlexGrid1.TextMatrix(0, 6) = "Balance Weight"
MSHFlexGrid1.TextMatrix(0, 7) = "New GSL Flag"
MSHFlexGrid1.TextMatrix(0, 8) = "New GSL Spillage Flag"
MSHFlexGrid1.TextMatrix(0, 9) = "Old CMP Name"
MSHFlexGrid1.TextMatrix(0, 10) = "Old G~S~L"
MSHFlexGrid1.TextMatrix(0, 11) = "Old GSL Spillage Flag"
MSHFlexGrid1.TextMatrix(0, 12) = "Remarks"
MSHFlexGrid1.TextMatrix(0, 13) = "G_UID"
MSHFlexGrid1.TextMatrix(0, 14) = "No Of Bags"
''-- For GSL Partial Trnf
MSHFlexGrid1.TextMatrix(0, 15) = "No of Madeup Bags"
MSHFlexGrid1.TextMatrix(0, 16) = "Madeup New GSL ID"
MSHFlexGrid1.TextMatrix(0, 17) = "Madeup Old GSL ID"
''---
MSHFlexGrid1.TextMatrix(0, 18) = "Quantity"
MSHFlexGrid1.TextMatrix(0, 19) = "Old Location"

MSHFlexGrid1.ColWidth(0) = 800
MSHFlexGrid1.ColWidth(1) = 1000
MSHFlexGrid1.ColWidth(2) = 1400
MSHFlexGrid1.ColWidth(3) = 1400
MSHFlexGrid1.ColWidth(4) = 1400
MSHFlexGrid1.ColWidth(5) = 800
MSHFlexGrid1.ColWidth(6) = 800
MSHFlexGrid1.ColWidth(7) = 800
MSHFlexGrid1.ColWidth(8) = 800
MSHFlexGrid1.ColWidth(9) = 1400
MSHFlexGrid1.ColWidth(10) = 1400
MSHFlexGrid1.ColWidth(11) = 800
MSHFlexGrid1.ColWidth(12) = 1800
MSHFlexGrid1.ColWidth(13) = 0
MSHFlexGrid1.ColWidth(14) = 1500
MSHFlexGrid1.ColWidth(15) = 1500
MSHFlexGrid1.ColWidth(16) = 1800
MSHFlexGrid1.ColWidth(17) = 1800
MSHFlexGrid1.ColWidth(18) = 1800
MSHFlexGrid1.ColWidth(19) = 1800


MSHFlexGrid1.RowHeight(0) = 600

End Sub


Private Sub CmbPreFix_GotFocus()
tmp = CmbPreFix.Text

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

mSLocationID = RsMst_Location!LocationID
'mLocationName = RsMst_Location!LocationName
'msSM_prefix = RsMst_Location!SM_Prefix

End Sub


Private Sub CmbSCmp_GotFocus()
If CmbSNewLocation.Text = "" Then
   MsgBox "Plase Select Location !!!"
   CmbSNewLocation.SetFocus
   Exit Sub
End If

tmp1 = ""
If Gen_WcCMP <> "" Then
   tmp1 = "And " & Gen_WcCMP
End If

tmp = CmbSCmp.Text
Call TableMst_CMP(" Where LocationID = " & mSNewLocationID & "" & tmp1)
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
mCMPName = RsMst_CMP!CMPID
End Sub

Private Sub CmbSOldCmp_GotFocus()
If CmbPreFix.Text = "" Then
   MsgBox "Plase Select Location !!!"
   CmbPreFix.SetFocus
   Exit Sub
End If

tmp1 = ""
If Gen_WcCMP <> "" Then
   tmp1 = "And " & Gen_WcCMP
End If

tmp = CmbSOldCMP.Text
Call TableMst_CMP(" Where LocationID = " & mSLocationID & "" & tmp1)
CmbSOldCMP.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_CMP.RecordCount
    CmbSOldCMP.AddItem RsMst_CMP!CMPName
    RsMst_CMP.MoveNext
Next
CmbSOldCMP.Text = tmp
End Sub
Private Sub CmbSOldCmp_LostFocus()
If CmbSOldCMP.Text = "" Then
   Exit Sub
End If
RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPName = '" & CmbSOldCMP.Text & "'"
If RsMst_CMP.EOF Then
   MsgBox "Invalid selection "
   CmbSOldCMP.SetFocus
   Exit Sub
End If
End Sub


Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Txn_GSLTransfer.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Txn_GSLTransfer.xls")
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

Private Function GetMaxGSLID(mPrefix_ As Variant) As Double
Dim Retval As Double
tmp = "Select max(GSLID) from Mst_GSL Where SM_Prefix = '" & mPrefix_ & "'"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetMaxGSLID = Retval
End Function

Private Sub CreateMadeUpGSL()
Dim mTransferMSMPrefix, mTransferMGSLID, mTransferMGUID As Variant

'-- On 19 Aug  2016 for Non-Ncdex Split Transfer With Madeup Bags
If mSplitBagsFlag = True Then
    If mConnectedGSLID > 0 And mMadeupFlag = 0 Then
        If Val(TxtMadeupBags) > 0 Then
        
            RsMst_GSL.MoveFirst
            RsMst_GSL.Filter = (" SM_Prefix = '" & TxtPreFix & "' And Mst_GSL_ID = " & mConnectedGSLID & " ")
            'RsMst_GSL.Find ("G_UID= '" & TxtGSLGUID & "'")
                
            ''-- TO Get New SM Prefix
            tmp = "Select * from Mst_GSL Where G_UID = '" & TxtNewGSLGUID & "'"
            Call Tmp1Table
            If TmpRs1.RecordCount > 0 Then
                mTransferSMPrefix = TmpRs1!SM_Prefix
            End If
            
            tmp = "Select * from Mst_GSL Where SM_Prefix = '" & mTransferSMPrefix & "'"
            Call Tmp1Table
            
            TmpRs1.AddNew
'            TmpRs1!GSLID = GetMaxGSLID(mTransferSMPrefix)
            TmpRs1!G_UID = GetGUID
            mTransferMGUID = TmpRs1!G_UID
'            mTransferMGSLID = TmpRs1!GSLID
            TmpRs1!MadeupFlag = 1
            TmpRs1!Connected_Mst_GSL_ID = TxtNewGSLID
            TmpRs1!SM_Prefix = mTransferSMPrefix
            TmpRs1!CMPID = mNewCMPID
            TmpRs1!GodownID = mNewGodownID
            TmpRs1!StackNo = CmbNewStack.Text
            TmpRs1!LotNo = "MUB" & CmbNewLot.Text
            TmpRs1!ExchangeTypeID = RsMst_GSL!ExchangeTypeID
            TmpRs1!ClientID = RsMst_GSL!ClientID
            TmpRs1!CommodityID = RsMst_GSL!CommodityID
            TmpRs1!SpillageFlag = RsMst_GSL!SpillageFlag
            TmpRs1!Status = RsMst_GSL!Status
            TmpRs1!SCMFlag = RsMst_GSL!SCMFlag
            TmpRs1!PledgeNo = RsMst_GSL!PledgeNo
            TmpRs1!CDTFDate = RsMst_GSL!CDTFDate
            TmpRs1!SpillageFlag = RsMst_GSL!SpillageFlag
            TmpRs1!LCDTF_Date = RsMst_GSL!LCDTF_Date
            TmpRs1!ContractID = RsMst_GSL!ContractID
            TmpRs1!BagSize = RsMst_GSL!BagSize
            TmpRs1!BagEquivalent = RsMst_GSL!BagEquivalent
            TmpRs1!CRAmount = RsMst_GSL!CRAmount
            TmpRs1!CRDate = RsMst_GSL!CRDate
            TmpRs1!NContractID = RsMst_GSL!NContractID
            
            'TmpRs1!DepositeWeight = IIf(IsNull(RsMst_GSL!DepositeWeight), 0, RsMst_GSL!DepositeWeight) + Val(TxtMadeupQty)
            TmpRs1!DepositeWeight = 0#
            TmpRs1!DepositeBags = Val(TxtMadeupBags)
            'TmpRs1!BalanceWeight = IIf(IsNull(RsMst_GSL!BalanceWeight), 0, RsMst_GSL!BalanceWeight) + Val(RsMst_GSL!BalanceWeight)
            TmpRs1!BalanceBags = 0#
            TmpRs1!BalanceBags = Val(TxtMadeupBags)
            TmpRs1!CDTFBags = 0#
            TmpRs1!CDTFWeight = 0#
            'TmpRs1!CDTFBags = IIf(IsNull(RsMst_GSL!CDTFBags), 0, RsMst_GSL!CDTFBags)
            'TmpRs1!CDTFWeight = IIf(IsNull(RsMst_GSL!CDTFWeight), 0, RsMst_GSL!CDTFWeight)
            TmpRs1!Status = RsMst_GSL!Status
            TmpRs1!SCMFlag = RsMst_GSL!SCMFlag
            TmpRs1!SpillageFlag = RsMst_GSL!SpillageFlag
            
            TmpRs1!RateCardType = RsMst_GSL!RateCardType
            TmpRs1!BillingCycleID = RsMst_GSL!BillingCycleID
            TmpRs1!BillingUnit = RsMst_GSL!BillingUnit
            TmpRs1!Rate = RsMst_GSL!Rate
            TmpRs1!RateForDailyPnL = RsMst_GSL!RateForDailyPnL
            TmpRs1!DepConDID = RsMst_GSL!DepConDID
            TmpRs1!NoOfDays = RsMst_GSL!NoOfDays
            
            If Val(TxtBalanceMadeupBags) = Val(TxtMadeupBags) And Val(TxtMadeupBags) > 0 Then
                TmpRs1!WHRBags = IIf(IsNull(RsMst_GSL!WHRBags), 0, RsMst_GSL!WHRBags)
                TmpRs1!WHRWeight = IIf(IsNull(RsMst_GSL!WHRWeight), 0, RsMst_GSL!WHRWeight)
            Else
                TmpRs1!WHRBags = 0
                TmpRs1!WHRWeight = 0
            End If
            
            TmpRs1!ISINID = RsMst_GSL!ISINID
            TmpRs1!FirstISINID = RsMst_GSL!FirstISINID
            TmpRs1!LCDTF_Date = RsMst_GSL!LCDTF_Date
            TmpRs1!KRNo = RsMst_GSL!KRNo
            TmpRs1!KRNoBatchID = RsMst_GSL!KRNoBatchID
            TmpRs1!KRDate = RsMst_GSL!KRDate
            TmpRs1!ContractID = RsMst_GSL!ContractID
            TmpRs1!NContractID = RsMst_GSL!NContractID
            TmpRs1!BGSLID_Mst_GSL_ID = Val(mConnectedGSLID)
            TmpRs1!BSM_Prefix = TxtPreFix
            
            If IsNull(TmpRs1!CreatedBy) = True Or TmpRs1!CreatedBy = "" Then
               TmpRs1!CreatedBy = Gen_UserLoginID
               TmpRs1!CreatedDate = Now
            Else
               TmpRs1!UpdatedBy = Gen_UserLoginID
               TmpRs1!UpdatedDate = Now
            End If
            
            TmpRs1.Update
            'TmpRs1.Requery
            TmpRs1.Close
            
            If Val(TxtBalanceMadeupBags) > 0 And Val(TxtMadeupBags) > 0 Then
            
            '            mTransferMGSLID = TmpRs1!GSLID
                tmp = "Select Mst_GSL_ID From Mst_GSL Where G_UID='" & mTransferMGUID & "'"
                Call Tmp2Table
    
                If TmpRs2.RecordCount > 0 Then
                   mTransferMGSLID = TmpRs2!Mst_GSL_ID
                End If
            
                tmp = " UPDATE Txn_GSL_Transfer SET MadeupNew_Mst_GSL_ID = " & Val(mTransferMGSLID) & ",MadeupOLD_Mst_GSL_ID = " & Val(mConnectedGSLID) & " Where GSLTransferID = " & Val(TxtTransferID) & " "
                Call Tmp2Table
            End If
            
            'tmp = "Select * from Mst_GSL Where SM_Prefix = '" & mTransferSMPrefix & "' And GSLID = " & mTransferMGSLID & " And MadeupFlag = '1' And G_UID = '" & mTransferMGUID & "' "
            'Call Tmp1Table
            
            tmp = " Select * from Mst_GSL Where MadeupFlag = '1' And SM_Prefix = '" & TxtPreFix & "' And Mst_GSL_ID = " & mConnectedGSLID & " "
            Call Tmp2Table
               
            If Not TmpRs2.EOF Then
                   TmpRs2!CDTFBags = IIf(IsNull(TmpRs2!CDTFBags), 0, TmpRs2!CDTFBags) + Val(TxtMadeupBags)
                   TmpRs2!CDTFWeight = 0
                   TmpRs2!BalanceBags = IIf(IsNull(TmpRs2!BalanceBags), 0, TmpRs2!BalanceBags) - Val(TxtMadeupBags)
                   TmpRs2!BalanceWeight = 0 ' IIf(IsNull(TmpRs2!BalanceWeight), 0, TmpRs2!BalanceWeight) - Val(TxtGSLAvailableWgt.Text)
                   If Val(TxtBalanceMadeupBags) - Val(TxtMadeupBags) = 0 Then
                    TmpRs2!FGSLID_Mst_GSL_ID = mTransferMGSLID
                    TmpRs2!FSM_Prefix = mTransferSMPrefix
                    TmpRs2!Status = "Z"
                   Else
                    TmpRs2!Status = "O"
                   End If
                   TmpRs2.Update
            End If
            
            'commented on 15/03/2013 by rakesh (as added new location)
            tmp = "Update Mst_GSL Set Connected_Mst_GSL_ID = " & mTransferMGSLID & " Where SM_Prefix = '" & mTransferSMPrefix & "' And Mst_GSL_ID = " & TxtNewGSLID & " And G_UID = '" & TxtNewGSLGUID & "'"
            Call Tmp1Table
        
        End If
    End If
    
Else

    tmp = "Select * from Mst_GSL Where SM_Prefix = '" & TxtPreFix & "'"
    Call Tmp1Table
    
    TmpRs1.AddNew
'    TmpRs1!GSLID = GetMaxGSLID(TxtPreFix)
    TmpRs1!G_UID = GetGUID
    mNewG_UID = TmpRs1!G_UID
'    mNewMGSLID = TmpRs1!GSLID
    TmpRs1!Connected_Mst_GSL_ID = TxtNewGSLID
    'added on 15/03/2013 by rakesh( as adding new location)
    'TmpRs1!SM_Prefix = TxtPreFix
    TmpRs1!SM_Prefix = mNewSM_Prefix
    TmpRs1!CMPID = mNewCMPID
    TmpRs1!GodownID = mNewGodownID
    TmpRs1!StackNo = CmbNewStack.Text
    TmpRs1!LotNo = "MUB" & CmbNewLot.Text
    TmpRs1!ExchangeTypeID = RsMst_GSL!ExchangeTypeID
    TmpRs1!ClientID = RsMst_GSL!ClientID
    TmpRs1!CommodityID = RsMst_GSL!CommodityID 'mCommodityID
    TmpRs1!SpillageFlag = RsMst_GSL!SpillageFlag
    TmpRs1!MadeupFlag = 1
    TmpRs1!Status = RsMst_GSL!Status
    TmpRs1!SCMFlag = RsMst_GSL!SCMFlag
    TmpRs1!PledgeNo = RsMst_GSL!PledgeNo
    'TmpRs1!ISINID = RsMst_GSL!ISINID
    'TmpRs1!CdtfNo = RsMst_GSL!CdtfNo
    TmpRs1!CDTFDate = RsMst_GSL!CDTFDate
    'TmpRs1!FirstISINID = RsMst_GSL!FirstISINID
    TmpRs1!SpillageFlag = RsMst_GSL!SpillageFlag
    'TmpRs1!DepositAdjWt = RsMst_GSL!DepositAdjWt
    'TmpRs1!WithdrawalAdjWt = RsMst_GSL!WithdrawalAdjWt
    TmpRs1!LCDTF_Date = RsMst_GSL!LCDTF_Date
    'TmpRs1!GrossWt = RsMst_GSL!GrossWt
    'TmpRs1!KRNo = RsMst_GSL!KRNo
    'TmpRs1!KRNoBatchID = RsMst_GSL!KRNoBatchID
    'TmpRs1!KRDate = RsMst_GSL!KRDate
    TmpRs1!ContractID = RsMst_GSL!ContractID
    TmpRs1!BagSize = RsMst_GSL!BagSize
    TmpRs1!BagEquivalent = RsMst_GSL!BagEquivalent
    'TmpRs1!WtQty = RsMst_GSL!WtQty
    TmpRs1!CRAmount = RsMst_GSL!CRAmount
    TmpRs1!CRDate = RsMst_GSL!CRDate
    TmpRs1!NContractID = RsMst_GSL!NContractID
    
    TmpRs1!RateCardType = RsMst_GSL!RateCardType
    TmpRs1!BillingCycleID = RsMst_GSL!BillingCycleID
    TmpRs1!BillingUnit = RsMst_GSL!BillingUnit
    TmpRs1!Rate = RsMst_GSL!Rate
    TmpRs1!RateForDailyPnL = RsMst_GSL!RateForDailyPnL
    TmpRs1!DepConDID = RsMst_GSL!DepConDID
    TmpRs1!NoOfDays = RsMst_GSL!NoOfDays
    
    
    If IsNull(TmpRs1!CreatedBy) = True Or TmpRs1!CreatedBy = "" Then
       TmpRs1!CreatedBy = Gen_UserLoginID
       TmpRs1!CreatedDate = Now
    Else
       TmpRs1!UpdatedBy = Gen_UserLoginID
       TmpRs1!UpdatedDate = Now
    End If
    
    TmpRs1.Update
    TmpRs1.Requery
    TmpRs1.Close
    
    tmp = "Select Mst_GSL_ID From Mst_GSL Where G_UID='" & mNewG_UID & "'"
    Call Tmp1Table
    
    If TmpRs1.RecordCount > 0 Then
    '   mCxTF_DetailsID = TmpRs!CxTF_DetailsID
       mNewMGSLID = TmpRs1!Mst_GSL_ID
       
    End If
    
    ''commented on 15/03/2013 by rakesh
    'tmp = "Select * from Mst_GSL Where SM_Prefix = '" & TxtPreFix & "' And GSLID = " & mNewMGSLID & " And MadeupFlag = '1' " 'G_UID = '" & TxtGSLGUID & "'"
    tmp = "Select * from Mst_GSL Where SM_Prefix = '" & mNewSM_Prefix & "' And Mst_GSL_ID = " & mNewMGSLID & " And MadeupFlag = '1' " 'G_UID = '" & TxtGSLGUID & "'"
    Call Tmp1Table
    
    tmp = "Select * from Mst_GSL Where SM_Prefix = '" & TxtPreFix & "' And Mst_GSL_ID = " & mConnectedGSLID & " And MadeupFlag = '1' " 'G_UID = '" & TxtGSLGUID & "'"
    Call Tmp4Table
    If TmpRs4.RecordCount > 0 Then
    '   TmpRs1!CommodityID = TmpRs4!CommodityID
       'RsMst_GSL!CDTFBags = IIf(IsNull(RsMst_GSL!CDTFBags), 0, RsMst_GSL!CDTFBags) + Val(TxtGSLAvailableNofBags.Text)
       TmpRs1!DepositeWeight = IIf(IsNull(TmpRs1!DepositeWeight), 0, TmpRs1!DepositeWeight) + Val(TmpRs4!DepositeWeight) 'TmpRs4!DepositeWeight
       TmpRs1!DepositeBags = IIf(IsNull(TmpRs1!DepositeBags), 0, TmpRs1!DepositeBags) + Val(TmpRs4!DepositeBags) 'TmpRs4!DepositeBags
       TmpRs1!CDTFBags = IIf(IsNull(TmpRs1!CDTFBags), 0, TmpRs1!CDTFBags) + Val(TmpRs4!CDTFBags) 'TmpRs4!CDTFBags
       TmpRs1!BalanceBags = IIf(IsNull(TmpRs1!BalanceBags), 0, TmpRs1!BalanceBags) + Val(TmpRs4!BalanceBags) 'TmpRs4!BalanceBags
       TmpRs1!CDTFWeight = IIf(IsNull(TmpRs1!CDTFWeight), 0, TmpRs1!CDTFWeight) + Val(TmpRs4!CDTFWeight) 'TmpRs4!CDTFWeight
       TmpRs1!BalanceWeight = IIf(IsNull(TmpRs1!BalanceWeight), 0, TmpRs1!BalanceWeight) + Val(TmpRs4!BalanceWeight) 'TmpRs4!BalanceWeight
       TmpRs1!Status = TmpRs4!Status
       TmpRs1!SCMFlag = TmpRs4!SCMFlag
       TmpRs1!SpillageFlag = TmpRs4!SpillageFlag
       If Check2.Value = 0 Then
          TmpRs1!WHRBags = IIf(IsNull(TmpRs1!WHRBags), 0, TmpRs1!WHRBags) + Val(TmpRs4!WHRBags) 'TmpRs4!WHRBags
          TmpRs1!WHRWeight = IIf(IsNull(TmpRs1!WHRWeight), 0, TmpRs1!WHRWeight) + Val(TmpRs4!WHRWeight) 'TmpRs4!WHRWeight
          TmpRs1!ExchangeTypeID = TmpRs4!ExchangeTypeID
          TmpRs1!ClientID = TmpRs4!ClientID
          TmpRs1!PledgeNo = TmpRs4!PledgeNo
          TmpRs1!ISINID = TmpRs4!ISINID
    '      TmpRs1!CdtfNo = TmpRs4!CdtfNo
    '      TmpRs1!CDTFDate = TmpRs4!CDTFDate
          TmpRs1!FirstISINID = TmpRs4!FirstISINID
    '      TmpRs1!PledgeBags = TmpRs4!PledgeBags
    '      TmpRs1!PledgeWeight = TmpRs4!PledgeWeight
    '      TmpRs1!DepositAdjWt = TmpRs4!DepositAdjWt
    '      TmpRs1!WithdrawalAdjWt = TmpRs4!WithdrawalAdjWt
    '      TmpRs1!Flag1 = TmpRs4!Flag1
          TmpRs1!LCDTF_Date = TmpRs4!LCDTF_Date
    '      TmpRs1!GrossWt = TmpRs4!GrossWt
          TmpRs1!KRNo = TmpRs4!KRNo
          TmpRs1!KRNoBatchID = TmpRs4!KRNoBatchID
          TmpRs1!KRDate = TmpRs4!KRDate
          TmpRs1!ContractID = TmpRs4!ContractID
          TmpRs1!NContractID = TmpRs4!NContractID
          TmpRs1!BGSLID_Mst_GSL_ID = Val(mConnectedGSLID)
          TmpRs1!BSM_Prefix = TxtPreFix
       End If
       If IsNull(TmpRs1!CreatedBy) = True Or TmpRs1!CreatedBy = "" Then
          TmpRs1!CreatedBy = Gen_UserLoginID
          TmpRs1!CreatedDate = Now
       Else
          TmpRs1!UpdatedBy = Gen_UserLoginID
          TmpRs1!UpdatedDate = Now
       End If
       TmpRs1.Update
    End If
    tmp = " Select * from Mst_GSL Where MadeUPFLag = '1' ANd SM_Prefix = '" & TxtPreFix & "' And Mst_GSL_ID = " & mConnectedGSLID & " "
    Call Tmp2Table
       
    If Not TmpRs2.EOF Then
           TmpRs2!CDTFBags = IIf(IsNull(TmpRs2!CDTFBags), 0, TmpRs2!CDTFBags) + Val(TmpRs1!BalanceBags)
           TmpRs2!CDTFWeight = 0
           TmpRs2!BalanceBags = IIf(IsNull(TmpRs2!BalanceBags), 0, TmpRs2!BalanceBags) - Val(TmpRs1!BalanceBags)
           TmpRs2!BalanceWeight = 0 ' IIf(IsNull(TmpRs2!BalanceWeight), 0, TmpRs2!BalanceWeight) - Val(TxtGSLAvailableWgt.Text)
           TmpRs2!FGSLID_Mst_GSL_ID = mNewMGSLID
           TmpRs2!FSM_Prefix = mNewSM_Prefix
           TmpRs2!Status = "Z"
           TmpRs2.Update
    End If
    
    'commented on 15/03/2013 by rakesh (as added new location)
    'tmp = "Update Mst_GSL Set ConnectedGSLID = " & mNewMGSLID & " Where SM_Prefix = '" & TxtPreFix & "' And GSLID = " & TxtNewGSLID & " And G_UID = '" & TxtNewGSLGUID & "'"
    tmp = "Update Mst_GSL Set Connected_Mst_GSL_ID = " & mNewMGSLID & " Where SM_Prefix = '" & mNewSM_Prefix & "' And Mst_GSL_ID = " & TxtNewGSLID & " And G_UID = '" & TxtNewGSLGUID & "'"
    Call Tmp1Table

End If
End Sub

Private Sub TxtMadeUpBags_LostFocus()
    If Trim(TxtMadeupBags) = "" Then
       TxtMadeupQty = ""
       Exit Sub
    End If
    If IsNumeric(TxtMadeupBags) = False Then
        MsgBox "Invalid No of Bags"
        TxtMadeupBags.SetFocus
        Exit Sub
    End If
    If Val(TxtMadeupBags) <= 0 Then
        MsgBox "Zero/Negative Bags are not allowed."
        TxtMadeupBags.SetFocus
        Exit Sub
    End If
    TxtMadeupBags = Format(TxtMadeupBags, "###0")
    If TxtMadeupBags > TxtBalanceMadeupBags Then
        MsgBox "More than Balance Madeup Bags are not allowed."
        TxtMadeupBags.SetFocus
        Exit Sub
    End If
End Sub

Private Sub TxtQty_LostFocus()

If TxtQty.Text = "" Then Exit Sub
If IsNumeric(TxtQty.Text) = False Or Val(TxtQty.Text) < 0 Then
   MsgBox "Invalid Numeric Format !!!"
   TxtQty.SetFocus
   Exit Sub
End If

End Sub

Private Function CheckMappedReservationContract() As Boolean
Retval = True

'tmp = "Select Distinct ContractID from Txn_SRGodownMapDetails a "
'tmp = tmp & " inner Join Txn_SRGodownMapping b on a.SRGodownMappingID = b.SRGodownMappingID "
'tmp = tmp & " Where a.GodownID = " & mGodownID & " "
'
'Call TmpTable
'
'If TmpRs.RecordCount > 0 Then
'   For X = 1 To TmpRs.RecordCount
If mContractID <> 0 Then
   tmp = "Select Distinct a.GodownID from Txn_SRGodownMapDetails a "
   tmp = tmp & " inner Join Txn_SRGodownMapping b on a.SRGodownMappingID = b.SRGodownMappingID "
   tmp = tmp & " Where b.ContractID = " & mContractID & " And a.GodownID = " & mNewGodownID & " "
       
   Call Tmp1Table
   If TmpRs1.RecordCount = 0 Then
      Retval = False
      'Exit For
   End If
'       TmpRs.MoveNext
'   Next
End If

CheckMappedReservationContract = Retval

End Function

Private Sub TxtTransferBags_LostFocus()
    If Trim(TxtTransferBags) = "" Then
       TxtTransferQty = ""
       Exit Sub
    End If
    If IsNumeric(TxtTransferBags) = False Then
        MsgBox "Invalid No of Bags"
        TxtTransferBags.SetFocus
        Exit Sub
    End If
    If Val(TxtTransferBags) <= 0 Then
        MsgBox "Zero/Negative Bags are not allowed."
        TxtTransferBags.SetFocus
        Exit Sub
    End If
    TxtTransferBags = Format(TxtTransferBags, "###0")
    If Val(TxtTransferBags) > Val(TxtBalanceBags) Then
        MsgBox "More than Balance Bags are not allowed."
        TxtTransferBags.SetFocus
        Exit Sub
    End If
    TxtTransferQty = Format(((TxtTransferBags * mBagSize) / 1000), "#####0.000")
End Sub
