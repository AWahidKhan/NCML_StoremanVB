VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmMst_FranchiseeGodownRateCard 
   Caption         =   "Franchisee Godown Rate Card"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   2295
      Left            =   0
      TabIndex        =   63
      Top             =   5880
      Visible         =   0   'False
      Width           =   15120
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
         ItemData        =   "FrmMst_FranchiseeGodownRateCard.frx":0000
         Left            =   6600
         List            =   "FrmMst_FranchiseeGodownRateCard.frx":000A
         Sorted          =   -1  'True
         TabIndex        =   117
         Top             =   690
         Width           =   2130
      End
      Begin VB.ComboBox CmbSLessorName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6600
         Sorted          =   -1  'True
         TabIndex        =   25
         Top             =   240
         Width           =   8160
      End
      Begin VB.ComboBox CmbSCommodityGroupID 
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
         Left            =   1320
         Sorted          =   -1  'True
         TabIndex        =   26
         Top             =   675
         Width           =   3240
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
         Left            =   13725
         TabIndex        =   28
         Top             =   720
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
         Height          =   350
         Left            =   12720
         TabIndex        =   27
         Top             =   720
         Width           =   1000
      End
      Begin VB.ComboBox CmbSLocationID 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1320
         Sorted          =   -1  'True
         TabIndex        =   24
         Top             =   240
         Width           =   3225
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   855
         Left            =   120
         TabIndex        =   64
         Top             =   1125
         Width           =   14655
         _ExtentX        =   25850
         _ExtentY        =   1508
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
      Begin VB.Label Label25 
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
         Left            =   4680
         TabIndex        =   118
         Top             =   750
         Width           =   555
      End
      Begin VB.Label Label7 
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
         Height          =   315
         Left            =   4680
         TabIndex        =   69
         Top             =   270
         Width           =   1770
      End
      Begin VB.Label Label2 
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
         Height          =   240
         Left            =   120
         TabIndex        =   68
         Top             =   735
         Width           =   1215
      End
      Begin VB.Label LblTotalRecords 
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
         Left            =   12000
         TabIndex        =   67
         Top             =   810
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
         Left            =   9600
         TabIndex        =   66
         Top             =   810
         Width           =   2415
      End
      Begin VB.Label Label6 
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
         Height          =   210
         Left            =   120
         TabIndex        =   65
         Top             =   315
         Width           =   960
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFBagTable 
      Height          =   2370
      Left            =   9120
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   240
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   4180
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame FraGodownDetails 
      Height          =   2460
      Left            =   120
      TabIndex        =   45
      Top             =   2430
      Width           =   8895
      Begin VB.CommandButton CmdAddGodownDetails 
         Caption         =   "Update List"
         Height          =   375
         Left            =   6960
         TabIndex        =   11
         Top             =   1995
         Width           =   1215
      End
      Begin VB.TextBox TxtGodownAddress 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   480
         Left            =   1800
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   1050
         Width           =   6615
      End
      Begin VB.TextBox TxtCMPName 
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   240
         Width           =   6375
      End
      Begin VB.TextBox TxtGodownName 
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   645
         Width           =   6375
      End
      Begin VB.TextBox TxtRateOnQty 
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
         Left            =   5370
         TabIndex        =   10
         Top             =   1995
         Width           =   990
      End
      Begin VB.ComboBox CmbBillingUnit 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "FrmMst_FranchiseeGodownRateCard.frx":0020
         Left            =   1800
         List            =   "FrmMst_FranchiseeGodownRateCard.frx":0022
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   1995
         Width           =   2010
      End
      Begin VB.TextBox TxtTotDays 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   375
         Left            =   7170
         Locked          =   -1  'True
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   1575
         Width           =   990
      End
      Begin VB.ComboBox CmbBillingCycle 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1800
         Sorted          =   -1  'True
         TabIndex        =   8
         Top             =   1575
         Width           =   2745
      End
      Begin VB.CommandButton CmdBillingCycle 
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
         Left            =   4620
         TabIndex        =   46
         Top             =   1605
         Width           =   375
      End
      Begin VB.TextBox TxtBillingCycle 
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1575
         Width           =   2730
      End
      Begin VB.TextBox TxtBillingUnit 
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1995
         Width           =   2010
      End
      Begin VB.Label Label1 
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
         Height          =   540
         Left            =   240
         TabIndex        =   57
         Top             =   1020
         Width           =   1320
      End
      Begin VB.Label LblCMPName 
         Caption         =   "CMP Name"
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
         TabIndex        =   55
         Top             =   270
         Width           =   1320
      End
      Begin VB.Label LblGodownID 
         Caption         =   "Godown Name"
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
         Left            =   255
         TabIndex        =   53
         Top             =   675
         Width           =   1440
      End
      Begin VB.Label LblRatePerBag 
         Caption         =   "Rate on Units"
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
         Left            =   4020
         TabIndex        =   51
         Top             =   2055
         Width           =   1320
      End
      Begin VB.Label Label11 
         Caption         =   "Billing Unit"
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
         TabIndex        =   50
         Top             =   2048
         Width           =   1365
      End
      Begin VB.Label Label14 
         Caption         =   "Total Days"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6045
         TabIndex        =   49
         Top             =   1590
         Width           =   1140
      End
      Begin VB.Label Label8 
         Caption         =   "Billing Cycle"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   48
         Top             =   1605
         Width           =   1440
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1095
      Left            =   120
      TabIndex        =   38
      Top             =   8580
      Width           =   8925
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New "
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
         TabIndex        =   0
         Top             =   240
         Width           =   1695
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
         Height          =   375
         Left            =   1815
         TabIndex        =   15
         Top             =   240
         Width           =   1695
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
         Height          =   375
         Left            =   3510
         TabIndex        =   16
         Top             =   240
         Width           =   1815
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
         Height          =   375
         Left            =   5325
         TabIndex        =   17
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "&Search"
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
         Left            =   7020
         TabIndex        =   18
         Top             =   240
         Width           =   1695
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
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   615
         Width           =   1695
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
         Height          =   375
         Left            =   1800
         TabIndex        =   22
         Top             =   615
         Width           =   1695
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
         Height          =   375
         Left            =   3510
         TabIndex        =   21
         Top             =   615
         Width           =   1815
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
         Height          =   375
         Left            =   5325
         TabIndex        =   20
         Top             =   615
         Width           =   1695
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
         Height          =   375
         Left            =   7020
         TabIndex        =   19
         Top             =   615
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2310
      Left            =   120
      TabIndex        =   33
      Top             =   120
      Width           =   8895
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
         ItemData        =   "FrmMst_FranchiseeGodownRateCard.frx":0024
         Left            =   6000
         List            =   "FrmMst_FranchiseeGodownRateCard.frx":002E
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   1830
         Width           =   2130
      End
      Begin VB.CommandButton CmdExchangeTypeID 
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
         Left            =   4440
         TabIndex        =   100
         TabStop         =   0   'False
         Top             =   1830
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
         ItemData        =   "FrmMst_FranchiseeGodownRateCard.frx":0044
         Left            =   1875
         List            =   "FrmMst_FranchiseeGodownRateCard.frx":0046
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   1830
         Width           =   2535
      End
      Begin VB.ComboBox CmbCommodityGroupID 
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
         TabIndex        =   5
         Top             =   1425
         Width           =   6255
      End
      Begin VB.CommandButton CmdCommodityGroupID 
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
         Left            =   8235
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   1425
         Width           =   375
      End
      Begin VB.ComboBox CmbLessorName 
         BeginProperty Font 
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
         TabIndex        =   3
         Top             =   1020
         Width           =   6255
      End
      Begin VB.CommandButton CmdLessorName 
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
         Left            =   8235
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   1020
         Width           =   375
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
         Height          =   315
         Left            =   8235
         TabIndex        =   35
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
         Left            =   1875
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   615
         Width           =   6255
      End
      Begin VB.TextBox TxtFGRCID 
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
         Left            =   1875
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   195
         Width           =   1320
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
         Left            =   1875
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   615
         Width           =   6255
      End
      Begin VB.TextBox TxtLessorName 
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
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1020
         Width           =   6255
      End
      Begin VB.TextBox TxtCommodityGroupID 
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
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1425
         Width           =   6255
      End
      Begin VB.TextBox TxtExchangeTypeID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BeginProperty Font 
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
         Locked          =   -1  'True
         TabIndex        =   98
         Top             =   1830
         Width           =   2535
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
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   115
         Top             =   1830
         Visible         =   0   'False
         Width           =   2100
      End
      Begin VB.Label Label24 
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
         Left            =   5115
         TabIndex        =   116
         Top             =   1890
         Width           =   660
      End
      Begin VB.Label Label10 
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
         Left            =   225
         TabIndex        =   99
         Top             =   1860
         Width           =   1695
      End
      Begin VB.Label Label4 
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
         Height          =   240
         Left            =   225
         TabIndex        =   62
         Top             =   1485
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Franchisee Name"
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
         Left            =   225
         TabIndex        =   60
         Top             =   960
         Width           =   1530
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
         Left            =   225
         TabIndex        =   37
         Top             =   645
         Width           =   1440
      End
      Begin VB.Label LblFGRCID 
         Caption         =   "Rate Card ID"
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
         Left            =   225
         TabIndex        =   36
         Top             =   255
         Width           =   1320
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3645
      Left            =   120
      TabIndex        =   70
      Top             =   4920
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   6429
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Godown Details"
      TabPicture(0)   =   "FrmMst_FranchiseeGodownRateCard.frx":0048
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "MSHFGodownDetails"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Log History"
      TabPicture(1)   =   "FrmMst_FranchiseeGodownRateCard.frx":0064
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label13"
      Tab(1).Control(1)=   "Label12"
      Tab(1).Control(2)=   "TxtCreated"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "TxtUpdated"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
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
         Left            =   -73080
         Locked          =   -1  'True
         TabIndex        =   95
         TabStop         =   0   'False
         Top             =   1635
         Width           =   6375
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
         Left            =   -73080
         Locked          =   -1  'True
         TabIndex        =   94
         TabStop         =   0   'False
         Top             =   1080
         Width           =   6375
      End
      Begin VB.Frame Frame2 
         Height          =   3810
         Left            =   -74880
         TabIndex        =   71
         Top             =   360
         Width           =   12210
         Begin VB.CommandButton CmdAssetType 
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
            Height          =   360
            Left            =   3960
            TabIndex        =   81
            Top             =   270
            Width           =   375
         End
         Begin VB.CommandButton CmdAdd2GridAsset 
            Caption         =   "Add To Grid"
            Height          =   375
            Left            =   9105
            TabIndex        =   80
            Top             =   1200
            Width           =   1455
         End
         Begin VB.ComboBox CmbAssetType 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1335
            Sorted          =   -1  'True
            TabIndex        =   79
            Top             =   270
            Width           =   2535
         End
         Begin VB.CommandButton CmdDeleteFromGridAsset 
            Caption         =   "Remove From Grid"
            Height          =   375
            Left            =   10575
            TabIndex        =   78
            Top             =   1200
            Width           =   1575
         End
         Begin VB.TextBox TxtAssetDescription 
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
            Left            =   3600
            MaxLength       =   500
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   77
            Top             =   720
            Width           =   5415
         End
         Begin VB.CommandButton cmdAssetStatus 
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
            Height          =   360
            Left            =   9120
            TabIndex        =   76
            Top             =   270
            Width           =   375
         End
         Begin VB.ComboBox CmbAssetStatus 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   5700
            Sorted          =   -1  'True
            TabIndex        =   75
            Top             =   270
            Width           =   3345
         End
         Begin VB.TextBox TxtAssetQty 
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
            Left            =   1320
            TabIndex        =   74
            Top             =   720
            Width           =   975
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
            Height          =   375
            Left            =   3585
            MaxLength       =   1000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   73
            Top             =   1200
            Width           =   5400
         End
         Begin VB.TextBox TxtAMCAmount 
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
            Left            =   1320
            TabIndex        =   72
            Top             =   1200
            Width           =   975
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
            Height          =   2100
            Left            =   120
            TabIndex        =   82
            Top             =   1635
            Width           =   12030
            _ExtentX        =   21220
            _ExtentY        =   3704
            _Version        =   393216
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSComCtl2.DTPicker AssetVerifiedon 
            Height          =   330
            Left            =   10680
            TabIndex        =   83
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   582
            _Version        =   393216
            Format          =   51118081
            CurrentDate     =   39850
         End
         Begin MSComCtl2.DTPicker TxtAMCExpDate 
            Height          =   330
            Left            =   10695
            TabIndex        =   84
            Top             =   720
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   582
            _Version        =   393216
            Format          =   51118081
            CurrentDate     =   39850
         End
         Begin VB.Label Label9 
            Caption         =   "Asset Type"
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
            Left            =   135
            TabIndex        =   92
            Top             =   263
            Width           =   1095
         End
         Begin VB.Label Label15 
            Caption         =   "Description"
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
            Left            =   2415
            TabIndex        =   91
            Top             =   757
            Width           =   1095
         End
         Begin VB.Label Label16 
            Caption         =   "Asset Status"
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
            Left            =   4560
            TabIndex        =   90
            Top             =   263
            Width           =   1215
         End
         Begin VB.Label Label17 
            Caption         =   "Verified On"
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
            Left            =   9645
            TabIndex        =   89
            Top             =   240
            Width           =   1035
         End
         Begin VB.Label Label18 
            Caption         =   "Asset Qty"
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
            TabIndex        =   88
            Top             =   772
            Width           =   1095
         End
         Begin VB.Label Label27 
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
            Height          =   300
            Left            =   2400
            TabIndex        =   87
            Top             =   1230
            Width           =   1095
         End
         Begin VB.Label Label28 
            Caption         =   "AMC Exp. Date"
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
            TabIndex        =   86
            Top             =   750
            Width           =   1395
         End
         Begin VB.Label Label29 
            Caption         =   "AMC Amount"
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
            TabIndex        =   85
            Top             =   1245
            Width           =   1200
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFGodownDetails 
         Height          =   3210
         Left            =   60
         TabIndex        =   93
         Top             =   375
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   5662
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
      Begin VB.Label Label12 
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
         Left            =   -74640
         TabIndex        =   97
         Top             =   1695
         Width           =   885
      End
      Begin VB.Label Label13 
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
         Left            =   -74640
         TabIndex        =   96
         Top             =   1140
         Width           =   855
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid6 
      Height          =   2250
      Left            =   9120
      TabIndex        =   112
      TabStop         =   0   'False
      Top             =   2880
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   3969
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame Frame4 
      Caption         =   "Fixed Rate Card Details"
      Height          =   4455
      Left            =   9120
      TabIndex        =   101
      Top             =   5280
      Width           =   6015
      Begin VB.Frame Frame5 
         Height          =   960
         Left            =   90
         TabIndex        =   102
         Top             =   240
         Width           =   5865
         Begin VB.TextBox TxtFixAmount 
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
            Left            =   2880
            TabIndex        =   107
            Top             =   390
            Width           =   1440
         End
         Begin VB.CommandButton CmdAdd2ListFixAmount 
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
            Height          =   345
            Left            =   4425
            TabIndex        =   109
            Top             =   375
            Width           =   1335
         End
         Begin VB.CommandButton CmdRemoveFixAmount 
            Caption         =   "Remove"
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
            Left            =   5640
            TabIndex        =   103
            Top             =   525
            Visible         =   0   'False
            Width           =   1335
         End
         Begin MSComCtl2.DTPicker TxtFromDate 
            Height          =   375
            Left            =   75
            TabIndex        =   104
            Top             =   390
            Width           =   1380
            _ExtentX        =   2434
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
            CustomFormat    =   "MM/dd/yyyy"
            Format          =   51118083
            CurrentDate     =   39850
            MinDate         =   38353
         End
         Begin MSComCtl2.DTPicker TxtToDate 
            Height          =   375
            Left            =   1485
            TabIndex        =   105
            Top             =   390
            Width           =   1380
            _ExtentX        =   2434
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
            CustomFormat    =   "MM/dd/yyyy"
            Format          =   51118083
            CurrentDate     =   39850
            MinDate         =   38353
         End
         Begin VB.Label Label20 
            Caption         =   "Fix Amount"
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
            Left            =   3060
            TabIndex        =   110
            Top             =   120
            Width           =   1215
         End
         Begin VB.Label Label19 
            Caption         =   "Start Date"
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
            Left            =   300
            TabIndex        =   108
            Top             =   120
            Width           =   975
         End
         Begin VB.Label Label26 
            Caption         =   "End Date"
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
            Left            =   1770
            TabIndex        =   106
            Top             =   120
            Width           =   855
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid5 
         Height          =   3105
         Left            =   120
         TabIndex        =   111
         Top             =   1215
         Width           =   5835
         _ExtentX        =   10292
         _ExtentY        =   5477
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
   End
   Begin VB.Frame FraBagDetails 
      Caption         =   "Rate Details As Per Bags"
      Height          =   4395
      Left            =   9120
      TabIndex        =   39
      Top             =   5280
      Width           =   6015
      Begin VB.Frame Frame3 
         Height          =   975
         Left            =   120
         TabIndex        =   40
         Top             =   240
         Width           =   5775
         Begin VB.TextBox TxtRatePerBag 
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
            Left            =   2925
            TabIndex        =   31
            Top             =   540
            Width           =   1215
         End
         Begin VB.TextBox TxtFromBagSize 
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
            Left            =   225
            MaxLength       =   100
            TabIndex        =   29
            Top             =   540
            Width           =   1215
         End
         Begin VB.TextBox TxtToBagSize 
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
            Left            =   1575
            MaxLength       =   50
            TabIndex        =   30
            Top             =   540
            Width           =   1215
         End
         Begin VB.CommandButton CmdAdd2List 
            Caption         =   "Add to List"
            Height          =   375
            Left            =   4320
            TabIndex        =   32
            Top             =   525
            Width           =   1215
         End
         Begin VB.Label Label5 
            Caption         =   "Rate Per Bag per Cycle"
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
            Left            =   2925
            TabIndex        =   43
            Top             =   120
            Width           =   1245
         End
         Begin VB.Label LblBagSizeFrom 
            Caption         =   "Bag Size From"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   225
            TabIndex        =   42
            Top             =   180
            Width           =   1260
         End
         Begin VB.Label LblBagsSizeTo 
            Caption         =   "Bag Size To"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1620
            TabIndex        =   41
            Top             =   180
            Width           =   1095
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFBagDetails 
         Height          =   3015
         Left            =   120
         TabIndex        =   44
         Top             =   1290
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   5318
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
   End
   Begin VB.Label Label22 
      Caption         =   "Fixed Rate Details as Per Date"
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
      Left            =   9120
      TabIndex        =   114
      Top             =   2595
      Width           =   3360
   End
   Begin VB.Label Label21 
      Caption         =   "Rate Details as Per Bag"
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
      Left            =   9120
      TabIndex        =   113
      Top             =   0
      Width           =   2280
   End
End
Attribute VB_Name = "FrmMst_FranchiseeGodownRateCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mLocationID, mCMPID, mGodownID As Variant
Dim mLessorID, mCommodityGroupId, mExchangeTypeID, mBillingCycleID, mVendorID, mSVendorID As Variant
Dim mSLocationID, mSLessorID, mSCommodityGroupId As Variant
Dim Edit_Flg As Variant
Dim mSelGodownUnitTypeID As Double
Const strChecked = "þ"
Const strUnChecked = "q"

Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)
Frame1.Enabled = True
FraGodownDetails.Enabled = False

TxtFGRCID.Text = ""

CmbLocationID.Visible = True
TxtLocationID.Visible = False
CmbLocationID.Text = ""
CmdLocationID.Visible = True

CmbLessorName.Visible = True
CmdLessorName.Visible = True
TxtLessorName.Visible = False
CmbLessorName.Text = ""

CmbCommodityGroupID.Visible = True
CmbCommodityGroupID.Text = ""
TxtCommodityGroupID.Visible = False
CmdCommodityGroupID.Visible = True

CmbExchangeTypeID.Visible = True
CmbExchangeTypeID.Text = ""
TxtExchangeTypeID.Visible = False
CmdExchangeTypeID.Visible = True

CmbStatus.Visible = True
TxtStatus.Visible = False
TxtStatus.Text = ""
CmbStatus.Text = ""

TxtCreated.Text = ""
TxtUpdated.Text = ""
TxtCMPName.Text = ""
TxtGodownName.Text = ""
TxtGodownAddress.Text = ""

CmbBillingCycle.Visible = True
TxtBillingCycle.Visible = False
CmbBillingCycle.Text = ""
CmdBillingCycle.Visible = True

CmbBillingUnit.Visible = True
TxtBillingUnit.Visible = False
CmbBillingUnit.Text = ""
TxtTotDays.Text = ""
TxtRateOnQty.Text = ""

Call Grid_HeadGodowndetails
Call Grid_HeadBagTable
Call Grid_HeadBagDetails
Call Grid_HeadFixAmount
Call Grid_HeadFixTable
If RsMst_FF_RateCard.RecordCount > 0 Then
   CmbLocationID.SetFocus
End If

End Sub

Private Sub CmbLocationID_GotFocus()
tmp = CmbLocationID.Text

If Gen_WcLocation = "" Then
   Call TableMst_Location
Else
   Call TableMst_Location(" Where " & Gen_WcLocation & "")
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
If tmp = "" Then
   mLocationID = Null
End If
End Sub
'----
Private Sub CmbLocationID_LostFocus()
Dim DoUpdate As Boolean
DoUpdate = False
If CmbLocationID.Text = "" Then
   CmbLessorName.Text = ""
   CmbCommodityGroupID.Text = ""
   CmbExchangeTypeID.Text = ""
   Call ClearGodownDetails
   Call ClearBagDetail
   Call Grid_HeadBagDetails
   Call Grid_HeadBagTable
   Call Grid_HeadFixTable
   Call Grid_HeadFixAmount
   Call Fill_CMPGodownDetails

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
   DoUpdate = True
End If

mLocationID = RsMst_Location!LocationID

If MSHFGodownDetails.Rows > 2 Then
   If DoUpdate = True Then
      If MsgBox(" Data in the list will be lost. Do you want to Continue.", vbYesNo) = vbYes Then
'         mLocationID = Null
         CmbLessorName.Text = ""
         CmbCommodityGroupID.Text = ""
         CmbExchangeTypeID.Text = ""
         Call ClearGodownDetails
         Call ClearBagDetail
         Call Grid_HeadBagDetails
         Call Grid_HeadBagTable
         Call Grid_HeadFixTable
         Call Grid_HeadFixAmount
         Call Fill_CMPGodownDetails
      End If
   End If
Else
  Call Fill_CMPGodownDetails
End If

End Sub
'Private Sub CmbLessorName_GotFocus()
'tmp = CmbLessorName.Text
'mLessorID = ""
'Call TableMst_Lessor
'CmbLessorName.Clear
'If RsMst_Lessor.RecordCount = 0 Then
'   MsgBox "No Record found !!!! "
'   Exit Sub
'End If
'For I = 1 To RsMst_Lessor.RecordCount
'    CmbLessorName.AddItem RsMst_Lessor!LessorName
'    RsMst_Lessor.MoveNext
'Next
'CmbLessorName.Text = tmp
'End Sub
'
'Private Sub CmbLessorName_LostFocus()
'If CmbLessorName.Text = "" Then
'   Exit Sub
'End If
'RsMst_Lessor.MoveFirst
'RsMst_Lessor.Find "LessorName = '" & CmbLessorName.Text & "'"
'If RsMst_Lessor.EOF Then
'   MsgBox "Invalid selection "
'   CmbLessorName.SetFocus
'   Exit Sub
'End If
'mLessorID = RsMst_Lessor!LessorID
'End Sub

Private Sub CmbLessorName_GotFocus()
tmp = CmbLessorName.Text
mLessorID = ""
Call TableMst_Vendor
CmbLessorName.Clear

If RsMst_Vendor.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Vendor.RecordCount
    CmbLessorName.AddItem RsMst_Vendor!VendorName
    RsMst_Vendor.MoveNext
Next
CmbLessorName.Text = tmp
End Sub

Private Sub CmbLessorName_LostFocus()
If CmbLessorName.Text = "" Then
   Exit Sub
End If
RsMst_Vendor.MoveFirst
RsMst_Vendor.Find "VendorName = '" & CmbLessorName.Text & "'"
If RsMst_Vendor.EOF Then
   MsgBox "Invalid selection "
   CmbLessorName.SetFocus
   Exit Sub
End If
mVendorID = RsMst_Vendor!VendorID
End Sub

Private Sub CmbCommodityGroupID_GotFocus()
tmp = CmbCommodityGroupID.Text
Call TableMst_CommodityGroup
CmbCommodityGroupID.Clear
If RsMst_CommodityGroup.RecordCount = 0 Then
  MsgBox "No Record found !!!! "
  Exit Sub
End If
For I = 1 To RsMst_CommodityGroup.RecordCount
   CmbCommodityGroupID.AddItem RsMst_CommodityGroup!COmmodityGroup
   RsMst_CommodityGroup.MoveNext
Next
CmbCommodityGroupID.Text = tmp
    
End Sub

'----
Private Sub CmbCommodityGroupID_LostFocus()
Dim DoUpdate As Boolean
DoUpdate = False
If CmbCommodityGroupID.Text = "" Then
   Exit Sub
End If

RsMst_CommodityGroup.MoveFirst
RsMst_CommodityGroup.Find "CommodityGroup = '" & CmbCommodityGroupID.Text & "'"
If RsMst_CommodityGroup.EOF Then
   MsgBox "Invalid selection "
   CmbCommodityGroupID.SetFocus
   Exit Sub
End If
If mCommodityGroupId <> RsMst_CommodityGroup!CommodityGroupID Then
   DoUpdate = True
End If

mCommodityGroupId = RsMst_CommodityGroup!CommodityGroupID

If MSHFGodownDetails.Rows > 2 Then
   If DoUpdate = True Then
      If MsgBox(" Data in the list will be lost. Do you want to Continue.", vbYesNo) = vbYes Then
         Call ClearGodownDetails
         Call ClearBagDetail
         Call Grid_HeadBagDetails
         Call Grid_HeadBagTable
         Call Grid_HeadFixTable
         Call Grid_HeadFixAmount
         
         Call Fill_CMPGodownDetails
      End If
   End If
Else
  Call Fill_CMPGodownDetails
End If


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

Dim DoUpdate As Boolean
DoUpdate = False

If CmbExchangeTypeID.Text = "" Then
   mExchangeTypeID = 0
   Call ClearGodownDetails
   Call ClearBagDetail
   Call Grid_HeadBagDetails
   Call Grid_HeadBagTable
   Call Grid_HeadFixTable
   Call Grid_HeadFixAmount
   Call Fill_CMPGodownDetails
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
   DoUpdate = True
End If

mExchangeTypeID = RsMst_ExchangeType!ExchangeTypeID

If MSHFGodownDetails.Rows > 2 Then
   If DoUpdate = True Then
      If MsgBox(" Data in the list will be lost. Do you want to Continue.", vbYesNo) = vbYes Then
         Call ClearGodownDetails
         Call ClearBagDetail
         Call Grid_HeadBagDetails
         Call Grid_HeadBagTable
         Call Grid_HeadFixTable
         Call Grid_HeadFixAmount
         
         Call Fill_CMPGodownDetails
      End If
   End If
Else
  Call Fill_CMPGodownDetails
End If

End Sub


Private Sub Fill_CMPGodownDetails()

If CmbLocationID.Text = "" Then Exit Sub
If CmbCommodityGroupID.Text = "" Then Exit Sub
If CmbExchangeTypeID.Text = "" Then Exit Sub

If Edit_Flg = "E" Then
   tmp = "Select MC.CMPName,MG.GodownNo,MG.Address,MC.CMPID,MG.GodownID,MG.UnitTypeID,MG.AgreementEndDate"
   tmp = tmp & " From Mst_Location ML"
   tmp = tmp & " Inner Join Mst_CMP MC On MC.LocationID = ML.LocationID"
   tmp = tmp & " Inner Join Mst_Godown MG On MG.CMPID=MC.CMPID"
   tmp = tmp & " Where Convert(Varchar,ML.LocationID)+'~'+Convert(Varchar,MC.CMPID)+'~'+Convert(Varchar,MG.GodownID) Not In"
   tmp = tmp & " (Select Convert(Varchar,A.LocationID)+'~'+Convert(Varchar,B.CMPID)+'~'+Convert(Varchar,B.GodownID)"
   tmp = tmp & " From Mst_FF_RateCard A"
   tmp = tmp & " Inner Join Mst_FF_RateCard_Godown B On A.FF_RateCardID=B.FF_RateCardID Where ((A.ExchangeTypeID = " & Val(mExchangeTypeID) & " And A.CommodityGroupID = " & mCommodityGroupId & ") or B.BillingUnit='F') And A.Flag = 'A')"
   tmp = tmp & " And ML.LocationID=" & mLocationID & " And MG.UnitTypeID in (2,6) And MG.CloseDate Is Null "
   tmp = tmp & " Order By MC.CMPName,MG.GodownNo,MG.Address"
Else
   Call Grid_HeadGodowndetails
   Call Grid_HeadBagTable
   Call Grid_HeadFixTable
   tmp = "Select MC.CMPName,MG.GodownNo,MG.Address,MC.CMPID,MG.GodownID,MG.UnitTypeID,MG.AgreementEndDate"
   tmp = tmp & " From Mst_Location ML"
   tmp = tmp & " Inner Join Mst_CMP MC On MC.LocationID =ML.LocationID"
   tmp = tmp & " Inner Join Mst_Godown MG On MG.CMPID=MC.CMPID"
   tmp = tmp & " Where ML.LocationID=" & mLocationID & " And MG.UnitTypeID in (2,6) and MG.CloseDate Is Null"
   tmp = tmp & " And Convert(Varchar,ML.LocationID)+'~'+Convert(Varchar,MC.CMPID)+'~'+Convert(Varchar,MG.GodownID) Not In"
   tmp = tmp & " (Select Convert(Varchar,A.LocationID)+'~'+Convert(Varchar,B.CMPID)+'~'+Convert(Varchar,B.GodownID)"
   tmp = tmp & " From Mst_FF_RateCard A"
   tmp = tmp & " Inner Join Mst_FF_RateCard_Godown B On A.FF_RateCardID=B.FF_RateCardID Where ((A.ExchangeTypeID = " & Val(mExchangeTypeID) & " And A.CommodityGroupID = " & mCommodityGroupId & ") or B.BillingUnit='F') And A.Flag = 'A' )" '
   tmp = tmp & " Order By MC.CMPName,MG.GodownNo,MG.Address"
   
End If
Call TmpTable
x = MSHFGodownDetails.Rows - 1
If TmpRs.RecordCount > 0 Then
   For I = x To TmpRs.RecordCount + (x - 1)
       MSHFGodownDetails.TextMatrix(I, 1) = TmpRs!CMPName
       MSHFGodownDetails.TextMatrix(I, 2) = TmpRs!GodownNo
       MSHFGodownDetails.TextMatrix(I, 3) = TmpRs!Address
       MSHFGodownDetails.TextMatrix(I, 7) = TmpRs!CMPID
       MSHFGodownDetails.TextMatrix(I, 8) = TmpRs!GodownID
       MSHFGodownDetails.TextMatrix(I, 10) = IIf(IsNull(TmpRs!AgreementEndDate), "", Format(TmpRs!AgreementEndDate, "dd-mmm-yyyy"))
       MSHFGodownDetails.row = I
       MSHFGodownDetails.Col = 0
       MSHFGodownDetails.CellFontName = "Wingdings"
       MSHFGodownDetails.CellAlignment = flexAlignCenterCenter
       MSHFGodownDetails.CellFontSize = 12
       MSHFGodownDetails.Text = strUnChecked
       TmpRs.MoveNext
       MSHFGodownDetails.Rows = MSHFGodownDetails.Rows + 1
   Next
Else
   If CmbLocationID.Visible = True Then
      MsgBox "No Franchisee Godown Found For Selected Location !!!"
      'CmbLocationID.SetFocus
   End If
End If

End Sub


Private Sub CmdAdd2ListFixAmount_Click()

If TxtFixAmount.Text = "" Then
   MsgBox "Blank Fix Amount Not Allowed !!"
   TxtFixAmount.SetFocus
   Exit Sub
End If

If TxtToDate.Value < TxtFromDate.Value Then
   MsgBox " 'Start Date' must be less than 'To Date'!!! "
   TxtFromDate.SetFocus
   Exit Sub
End If

If CheckDuplicateFixAmountDateRange = True Then
   MsgBox "'Start Date' or 'End Date' already in one of the range from grid!!! "
   TxtFromDate.SetFocus
   Exit Sub
End If

If CmdAdd2ListFixAmount.Caption = "Add to List" Then
   I = MSHFlexGrid5.Rows - 1
   MSHFlexGrid5.TextMatrix(I, 5) = I
   MSHFlexGrid5.Rows = MSHFlexGrid5.Rows + 1
Else
   I = MSHFlexGrid5.RowSel
End If

MSHFlexGrid5.TextMatrix(I, 0) = Format(TxtFromDate.Value, "dd-MMM-yyyy")
MSHFlexGrid5.TextMatrix(I, 1) = Format(TxtToDate.Value, "dd-MMM-yyyy")
MSHFlexGrid5.TextMatrix(I, 2) = TxtFixAmount
MSHFlexGrid5.TextMatrix(I, 3) = MSHFGodownDetails.TextMatrix(MSHFGodownDetails.RowSel, 7)
MSHFlexGrid5.TextMatrix(I, 4) = MSHFGodownDetails.TextMatrix(MSHFGodownDetails.RowSel, 8)

Call ClearFixAmountDetails

TxtFromDate.Value = CDate(MSHFlexGrid5.TextMatrix(I, 1)) + 1
TxtToDate.Value = TxtFromDate.Value + 365

TxtFromDate.SetFocus


End Sub

Private Sub CmdBillingCycle_Click()
FrmMst_BillingCycle.Show
End Sub

Private Sub CmdCommodityGroupID_Click()
FrmMst_CommodityGroup.Show
End Sub

Private Sub CmdExchangeTypeID_Click()
FrmMst_ExchangeType.Show
End Sub

Private Sub CmdGo_Click()

Dim wc As String
wc = ""
Call Grid_HeadSearch

If CmbSLocationID.Text <> "" Then
   wc = GenWc(wc, mSLocationID, "MF.LocationID", "N", "N")
End If

'If CmbSLessorName.Text <> "" Then
'   Wc = GenWc(Wc, mSLessorID, "MF.LessorID", "N", "N")
'End If

If CmbSLessorName.Text <> "" Then
   wc = GenWc(wc, mSVendorID, "MF.VendorID", "N", "N")
End If


If CmbSCommodityGroupID.Text <> "" Then
   wc = GenWc(wc, mSCommodityGroupId, "MF.CommodityGroupID", "N", "N")
End If

If CmbSStatus.Text <> "" Then
   wc = GenWc(wc, Left$(CmbSStatus.Text, 1), "MF.Flag", "S", "N")
End If

tmp = " Select MF.FF_RateCardID,ML.LocationName,MV.VendorName,MC.CommodityGroup,MF.Flag"
tmp = tmp & " From Mst_FF_RateCard MF"
tmp = tmp & " Inner Join Mst_Location ML On  ML.LocationID= MF.LocationID"
tmp = tmp & " Left Join Mst_Vendor MV On MV.VendorID = MF.VendorID"
tmp = tmp & " Inner Join Mst_CommodityGroup MC On MC.CommodityGroupID=MF.CommodityGroupID "
tmp = tmp & wc

Call TmpTable

If TmpRs.RecordCount > 0 Then
   For I = 1 To TmpRs.RecordCount
       For C = 0 To TmpRs.Fields.Count - 1
           MSHFlexGrid1.TextMatrix(I, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
       Next
       If LCase(MSHFlexGrid1.TextMatrix(I, 4)) = "a" Then
          MSHFlexGrid1.TextMatrix(I, 4) = "Active"
       Else
          MSHFlexGrid1.TextMatrix(I, 4) = "Deactive"
       End If
    
       TmpRs.MoveNext
       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
   Next
Else
   MsgBox "No Record Found!!!"
End If
LblTotalRecords.Caption = TmpRs.RecordCount
End Sub

Private Sub Grid_HeadSearch()
With MSHFlexGrid1
     .Clear
     .Rows = 2
     .Cols = 5
     .Font.Size = 10
     .AllowUserResizing = flexResizeBoth
     .WordWrap = True
     .TextMatrix(0, 0) = "FF Rate Card ID"
     .TextMatrix(0, 1) = "Location Name"
     .TextMatrix(0, 2) = "Franchisee Name"
     .TextMatrix(0, 3) = "Commodity"
     .TextMatrix(0, 4) = "Status"
     
     .ColWidth(0) = 1200
     .ColWidth(1) = 3500
     .ColWidth(2) = 3500
     .ColWidth(3) = 3500
     .ColWidth(4) = 1500
     
     .RowHeight(0) = 500
End With
End Sub

Private Sub CmdLessorName_Click()
'FrmMst_Lessor.Show
FrmMst_Vendor.Show

End Sub

Private Sub CmdLocationID_Click()
FrmMst_Location.Show
End Sub

Private Sub CmdRemoveFixAmount_Click()

If MSHFlexGrid5.Rows < 3 Then Exit Sub
Dim ans As Variant
If CmdAdd2ListFixAmount.Caption = "Add to List" Then Exit Sub
If TxtFixAmount = "" Then
   MsgBox "Please Select atleast one row from grid!!!"
   Exit Sub
End If


I = MSHFlexGrid5.RowSel
If I > 0 Then
   If MSHFlexGrid5.TextMatrix(I, 0) <> "" Then
      For RO = I To MSHFlexGrid5.Rows - 2
          For C = 0 To MSHFlexGrid5.Cols - 1
              MSHFlexGrid5.TextMatrix(RO, C) = MSHFlexGrid5.TextMatrix(RO + 1, C)
          Next
      Next
      MSHFlexGrid5.Rows = MSHFlexGrid5.Rows - 1
   End If
End If
Call ClearFixAmountDetails
TxtFromDate.SetFocus


End Sub

Private Sub DeleteCmd_Click()
Dim ans As Variant
On Error GoTo msgError

tmp = "Select * from Txn_InvMonGSLDetail Where FF_RateCardID = " & TxtFGRCID & " "
Call TmpTable

If TmpRs.RecordCount > 0 Then
   MsgBox "Rate Card used in Income module. Can't delete!!!"
   Exit Sub
End If

ans = MsgBox("Are you Sure you want to Delete this Record ", vbYesNo)
If ans = vbYes Then
   Call DeleteChildTableDetail
   RsMst_FF_RateCard.Delete
   RsMst_FF_RateCard.Update
   If RsMst_FF_RateCard.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_FF_RateCard.MoveNext
   If RsMst_FF_RateCard.EOF() Then
      RsMst_FF_RateCard.MoveLast
   End If
   Call Indata
End If

Exit Sub
msgError:
MsgBox "Child Record Present "
RsMst_FF_RateCard.CancelUpdate

End Sub

Private Sub EditCmd_Click()
Edit_Flg = "E"
Call GButtonLock(Me, False)

'CmbLocationID.Visible = True
'TxtLocationID.Visible = True
CmbLocationID.Text = TxtLocationID.Text
CmdLocationID.Visible = True

'CmbLessorName.Visible = True
'TxtLessorName.Visible = True
CmbLessorName.Text = TxtLessorName.Text
CmdLessorName.Visible = True

'CmbCommodityGroupID.Visible = True
'TxtCommodityGroupID.Visible = True
CmbCommodityGroupID.Text = TxtCommodityGroupID.Text
CmdCommodityGroupID.Visible = True
CmdBillingCycle.Visible = True

'CmbExchangeTypeID.Visible = True
CmbExchangeTypeID.Text = TxtExchangeTypeID.Text
'TxtExchangeTypeID.Visible = False
CmdExchangeTypeID.Visible = True

TxtStatus.Visible = False
CmbStatus.Visible = True
CmbStatus.Text = TxtStatus.Text

Call Fill_CMPGodownDetails

End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Call GButtonLockAD(Me)
   Frame1.Enabled = False
   FraGodownDetails.Enabled = False
   FraBagDetails.Enabled = False
   Frame4.Enabled = False
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If

Frame1.Enabled = True
FraGodownDetails.Enabled = False
FraBagDetails.Enabled = False
Frame4.Enabled = False

Call Grid_HeadGodowndetails
Call Grid_HeadBagDetails
Call Grid_HeadBagTable
Call Grid_HeadSearch
Call Grid_HeadFixAmount
Call Grid_HeadFixTable

CmbSLocationID.Text = ""
CmbSLessorName.Text = ""
CmbCommodityGroupID.Text = ""
LblTotalRecords.Caption = "-"
'CmbSLocationID.SetFocus

Call TableMst_FF_RateCard
Call TableMst_FF_RateCard_Godown
Call TableMst_FF_RateCard_Godown_Bag

If RsMst_FF_RateCard.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If

Call GButtonLock(Me, True)

Call Indata
End Sub
Private Sub Indata()
CmbLocationID.Visible = False
CmdLocationID.Visible = False
TxtLocationID.Visible = True

CmbLessorName.Visible = False
CmdLessorName.Visible = False
TxtLessorName.Visible = True

CmbCommodityGroupID.Visible = False
CmdCommodityGroupID.Visible = False
TxtCommodityGroupID.Visible = True

CmbExchangeTypeID.Visible = False
TxtExchangeTypeID.Visible = True
CmdExchangeTypeID.Visible = False

CmbStatus.Visible = False
TxtStatus.Visible = True

CmdBillingCycle.Visible = False

TxtFGRCID.Text = RsMst_FF_RateCard!FF_RateCardID
mLocationID = RsMst_FF_RateCard!LocationID
TxtLocationID.Text = GetLocationName(RsMst_FF_RateCard!LocationID)
mCommodityGroupId = RsMst_FF_RateCard!CommodityGroupID
TxtCommodityGroupID.Text = GetCommodityGroupName(RsMst_FF_RateCard!CommodityGroupID)
'mLessorID = RsMst_FF_RateCard!LessorID
'TxtLessorName.Text = GetLessorName(RsMst_FF_RateCard!LessorID)

If IsNull(RsMst_FF_RateCard!VendorID) = False Then
   mVendorID = RsMst_FF_RateCard!VendorID
   TxtLessorName.Text = GetVendorName(RsMst_FF_RateCard!VendorID, "N")
Else
   'mVendorD = RsMst_FF_RateCard!VendorID
   TxtLessorName.Text = "" 'GetVendorName(RsMst_FF_RateCard!VendorID, "N")
End If

mExchangeTypeID = RsMst_FF_RateCard!ExchangeTypeID
TxtExchangeTypeID.Text = GetExchangeName(RsMst_FF_RateCard!ExchangeTypeID)

If LCase(IIf(IsNull(RsMst_FF_RateCard!Flag), "", RsMst_FF_RateCard!Flag)) = "a" Then
   TxtStatus.Text = "Active"
ElseIf LCase(IIf(IsNull(RsMst_FF_RateCard!Flag), "", RsMst_FF_RateCard!Flag)) = "d" Then
   TxtStatus.Text = "Deactive"
End If

TxtCreated = IIf(IsNull(RsMst_FF_RateCard!CreatedBy), "", RsMst_FF_RateCard!CreatedBy) & " :- " & IIf(IsNull(RsMst_FF_RateCard!CreatedDate), "", RsMst_FF_RateCard!CreatedDate)
TxtUpdated = IIf(IsNull(RsMst_FF_RateCard!UpdatedBy), "", RsMst_FF_RateCard!UpdatedBy) & " :- " & IIf(IsNull(RsMst_FF_RateCard!UpdatedDate), "", RsMst_FF_RateCard!UpdatedDate)

FraBagDetails.Enabled = False
Frame4.Enabled = False

Call ClearGodownDetails
Call ClearBagDetail
Call Grid_HeadGodowndetails
Call Grid_HeadBagDetails
Call Grid_HeadBagTable
Call Grid_HeadFixAmount
Call Grid_HeadFixTable

Call IndataGodown_BagDetails

End Sub
Private Sub IndataGodown_BagDetails()
Dim x As Variant


Call TableMst_FF_RateCard_Godown(" Where FF_RateCardID=" & Val(TxtFGRCID) & "")
For I = 1 To RsMst_FF_RateCard_Godown.RecordCount
    MSHFGodownDetails.TextMatrix(I, 1) = GetCMPName(RsMst_FF_RateCard_Godown!CMPID) '"CmpName"
    MSHFGodownDetails.TextMatrix(I, 2) = GetGodownName(RsMst_FF_RateCard_Godown!GodownID, "N") '"Godown Name"
    MSHFGodownDetails.TextMatrix(I, 3) = GetGodownName(RsMst_FF_RateCard_Godown!GodownID, "A") '"Address"
    MSHFGodownDetails.TextMatrix(I, 4) = GetBillingCycle(RsMst_FF_RateCard_Godown!BillingCycleID, "N") '"billing Cycle"
     
    MSHFGodownDetails.TextMatrix(I, 5) = RsMst_FF_RateCard_Godown!BillingUnit '"Billing Unit"
    If LCase(MSHFGodownDetails.TextMatrix(I, 5)) = "b" Then
       MSHFGodownDetails.TextMatrix(I, 5) = "Bags"
       Call TableMst_FF_RateCard_Godown_Bag(" Where FF_RateCard_GodownID=" & RsMst_FF_RateCard_Godown!FF_RateCard_GodownID & "")
       
       x = MSHFBagTable.Rows - 1
       For j = x To RsMst_FF_RateCard_Godown_Bag.RecordCount + (x - 1)
           MSHFBagTable.TextMatrix(j, 0) = RsMst_FF_RateCard_Godown_Bag!FF_RateCard_Godown_BagID
           MSHFBagTable.TextMatrix(j, 1) = MSHFGodownDetails.TextMatrix(I, 1) '"CMPName"
           MSHFBagTable.TextMatrix(j, 2) = MSHFGodownDetails.TextMatrix(I, 2) '"GodownName"
           MSHFBagTable.TextMatrix(j, 3) = RsMst_FF_RateCard_Godown_Bag!BagSizeFrom '"from"
           MSHFBagTable.TextMatrix(j, 4) = RsMst_FF_RateCard_Godown_Bag!BagSizeTo '"To"
           MSHFBagTable.TextMatrix(j, 5) = Format(RsMst_FF_RateCard_Godown_Bag!RatePerBag, "####.0000") '"Rate"
           MSHFBagTable.TextMatrix(j, 6) = RsMst_FF_RateCard_Godown!CMPID '"CMP ID"
           MSHFBagTable.TextMatrix(j, 7) = RsMst_FF_RateCard_Godown!GodownID '"GodownID"
           MSHFBagTable.TextMatrix(j, 8) = j '"ID"
           MSHFBagTable.Rows = MSHFBagTable.Rows + 1
           RsMst_FF_RateCard_Godown_Bag.MoveNext
       Next
       
    ElseIf LCase(MSHFGodownDetails.TextMatrix(I, 5)) = "q" Then
       MSHFGodownDetails.TextMatrix(I, 5) = "Quantity"
    ElseIf LCase(MSHFGodownDetails.TextMatrix(I, 5)) = "f" Then
       MSHFGodownDetails.TextMatrix(I, 5) = "Fixed"
       Call TableMst_FF_RateCard_Godown_Fix(" Where FF_RateCard_GodownID=" & RsMst_FF_RateCard_Godown!FF_RateCard_GodownID)

       x = MSHFlexGrid6.Rows - 1
       For j = x To RsMst_FF_RateCard_Godown_Fix.RecordCount + (x - 1)
           MSHFlexGrid6.TextMatrix(j, 0) = RsMst_FF_RateCard_Godown_Fix!FF_RateCard_Godown_FixID
           MSHFlexGrid6.TextMatrix(j, 1) = MSHFGodownDetails.TextMatrix(I, 1) '"CMPName"
           MSHFlexGrid6.TextMatrix(j, 2) = MSHFGodownDetails.TextMatrix(I, 2) '"GodownName"
           MSHFlexGrid6.TextMatrix(j, 3) = Format(RsMst_FF_RateCard_Godown_Fix!StartDate, "dd-MMM-yyyy")
           MSHFlexGrid6.TextMatrix(j, 4) = Format(RsMst_FF_RateCard_Godown_Fix!EndDate, "dd-MMM-yyyy")
           MSHFlexGrid6.TextMatrix(j, 5) = Format(RsMst_FF_RateCard_Godown_Fix!FixAmount, "####.00") '"Rate"
           MSHFlexGrid6.TextMatrix(j, 6) = RsMst_FF_RateCard_Godown!CMPID '"CMP ID"
           MSHFlexGrid6.TextMatrix(j, 7) = RsMst_FF_RateCard_Godown!GodownID '"GodownID"
           MSHFlexGrid6.TextMatrix(j, 8) = j '"ID"
           MSHFlexGrid6.Rows = MSHFlexGrid6.Rows + 1
           RsMst_FF_RateCard_Godown_Fix.MoveNext
       Next
       
    ElseIf LCase(MSHFGodownDetails.TextMatrix(I, 5)) = "p" Then
       MSHFGodownDetails.TextMatrix(I, 5) = "Percentage"
    End If
    
    MSHFGodownDetails.TextMatrix(I, 6) = Format(RsMst_FF_RateCard_Godown!RateOnQty, "####.0000") '"Rate"
    MSHFGodownDetails.TextMatrix(I, 7) = RsMst_FF_RateCard_Godown!CMPID '"CMPID"
    MSHFGodownDetails.TextMatrix(I, 8) = RsMst_FF_RateCard_Godown!GodownID '"GodownID"
    MSHFGodownDetails.TextMatrix(I, 9) = RsMst_FF_RateCard_Godown!BillingCycleID '"BillingCycleID"
    MSHFGodownDetails.TextMatrix(I, 10) = GetGodownName(RsMst_FF_RateCard_Godown!GodownID, "ED") '"Agree Exp Date"
    MSHFGodownDetails.TextMatrix(I, 10) = Format(MSHFGodownDetails.TextMatrix(I, 10), "dd-mmm-yyyy")
    MSHFGodownDetails.row = I
    MSHFGodownDetails.Col = 0
    MSHFGodownDetails.CellFontName = "Wingdings"
    MSHFGodownDetails.CellFontSize = 12
    MSHFGodownDetails.CellAlignment = flexAlignCenterCenter
    MSHFGodownDetails.Text = strChecked
    MSHFGodownDetails.Rows = MSHFGodownDetails.Rows + 1
    RsMst_FF_RateCard_Godown.MoveNext
Next

End Sub

Private Sub Grid_HeadBagTable()
With MSHFBagTable
     .Clear
     .Cols = 9
     .Rows = 2
     .AllowUserResizing = flexResizeColumns
     .Font.Size = 9
     .WordWrap = True
     .TextMatrix(0, 0) = "FGRCID"
     .TextMatrix(0, 1) = "CMP"
     .TextMatrix(0, 2) = "Godown"
     .TextMatrix(0, 3) = "Bag Size From"
     .TextMatrix(0, 4) = "Bag Size To"
     .TextMatrix(0, 5) = "Rate Per Bag"
     .TextMatrix(0, 6) = "CMPID"
     .TextMatrix(0, 7) = "GodownID"
     .TextMatrix(0, 8) = "ID"
     .ColWidth(0) = 0
     .ColWidth(1) = 1600
     .ColWidth(2) = 1400
     .ColWidth(3) = 900
     .ColWidth(4) = 900
     .ColWidth(5) = 900
     .ColWidth(6) = 0
     .ColWidth(7) = 0
     .ColWidth(8) = 0
     .RowHeight(0) = 500
End With
End Sub

Private Sub Grid_HeadGodowndetails()
With MSHFGodownDetails
     .Clear
     .Rows = 2
     .Cols = 11 '10
     .WordWrap = True
     .AllowUserResizing = flexResizeBoth ' flexResizeColumns
     .Font.Size = 10
     .TextMatrix(0, 0) = "Select"
     .TextMatrix(0, 1) = "CMP Name"
     .TextMatrix(0, 2) = "Godown No"
     .TextMatrix(0, 3) = "Godown Address"
     .TextMatrix(0, 4) = "Billing Cycle"
     .TextMatrix(0, 5) = "Billing Unit"
     .TextMatrix(0, 6) = "Rate on Quantity"
     .TextMatrix(0, 7) = "CMPID"
     .TextMatrix(0, 8) = "GodownID"
     .TextMatrix(0, 9) = "BillingCycleID"
     .TextMatrix(0, 10) = "Agree End Date"
     
     .ColWidth(0) = 700
     .ColWidth(1) = 1800
     .ColWidth(2) = 1400
     .ColWidth(3) = 1800
     .ColWidth(4) = 1200
     .ColWidth(5) = 1000
     .ColWidth(6) = 900
     .ColWidth(7) = 0
     .ColWidth(8) = 0
     .ColWidth(9) = 0
     .ColWidth(10) = 1000
     
     .FixedCols = 0
     .ColAlignment(2) = vbLeftJustify
     .RowHeight(0) = 500
End With
End Sub

Private Sub Grid_HeadBagDetails()
With MSHFBagDetails
     .Clear
     .Cols = 6
     .Rows = 2
     .AllowUserResizing = flexResizeColumns
     .Font.Size = 10
     .TextMatrix(0, 0) = "Bag Size From"
     .TextMatrix(0, 1) = "Bag Size To"
     .TextMatrix(0, 2) = "Rate Per Bag"
     .TextMatrix(0, 3) = "CMPID"
     .TextMatrix(0, 4) = "GodownID"
     .TextMatrix(0, 5) = "ID"
     .ColWidth(0) = 1400
     .ColWidth(1) = 1600
     .ColWidth(2) = 1400
     .ColWidth(3) = 0
     .ColWidth(4) = 0
     .ColWidth(5) = 0
     .FixedCols = 0
End With
End Sub

Private Sub FraGodownDetails_Click()
If TxtCMPName.Text = "" And TxtGodownName.Text = "" Then
   FraGodownDetails.Enabled = False
Else
   FraGodownDetails.Enabled = True
End If
End Sub

Private Sub MSHFBagDetails_Click()
If SaveCmd.Enabled = False Then Exit Sub
Dim r As Variant
r = MSHFBagDetails.RowSel
If r < 1 Or r = MSHFBagDetails.Rows - 1 Then
   Call ClearBagDetail
   Exit Sub
End If
If MSHFBagDetails.TextMatrix(r, 0) = "" Then
   Exit Sub
End If

TxtFromBagSize = MSHFBagDetails.TextMatrix(r, 0)
TxtToBagSize = MSHFBagDetails.TextMatrix(r, 1)
TxtRatePerBag = MSHFBagDetails.TextMatrix(r, 2)
CmdAdd2List.Caption = "Update List"
TxtFromBagSize.SetFocus
End Sub



Private Sub MSHFGodownDetails_Click()
Dim x As Variant

If SaveCmd.Enabled = False Then Exit Sub
Call ClearGodownDetails
Call ClearBagDetail
Call ClearFixAmountDetails
Call Grid_HeadBagDetails
Call Grid_HeadFixAmount
x = MSHFGodownDetails.RowSel

MSHFGodownDetails.row = x

If MSHFGodownDetails.TextMatrix(x, 0) = strChecked Then

   If MSHFGodownDetails.ColSel <> 0 Then
      GoTo ExecuteProc
      Exit Sub
   Else
      If LCase(MSHFGodownDetails.TextMatrix(x, 5)) = "bags" Then
         ans = MsgBox("Data entered for bags will be lost, do you want to continue ?", vbYesNo)
         If ans = vbNo Then
            Exit Sub
         End If
      ElseIf LCase(MSHFGodownDetails.TextMatrix(x, 5)) = "fixed" Then
         ans = MsgBox("Data entered for Fixed Amounr will be lost, do you want to continue ?", vbYesNo)
         If ans = vbNo Then
            Exit Sub
         End If
      End If
      
   End If
   MSHFGodownDetails.TextMatrix(x, 0) = strUnChecked
   FraGodownDetails.Enabled = False
   For I = 1 To MSHFBagTable.Rows - 2
       If I >= MSHFBagTable.Rows Then
          Exit For
       End If
       If MSHFBagTable.TextMatrix(I, 6) = MSHFGodownDetails.TextMatrix(x, 7) And MSHFBagTable.TextMatrix(I, 7) = MSHFGodownDetails.TextMatrix(x, 8) Then
          If MSHFBagTable.TextMatrix(I, 1) <> "" Then
             For RO = I To MSHFBagTable.Rows - 2
                 For C = 0 To MSHFBagTable.Cols - 1
                     MSHFBagTable.TextMatrix(RO, C) = MSHFBagTable.TextMatrix(RO + 1, C)
                 Next
             Next
             MSHFBagTable.Rows = MSHFBagTable.Rows - 1
             I = I - 1
          End If
       End If
   Next
   
   For I = 1 To MSHFlexGrid6.Rows - 2
       If I >= MSHFlexGrid6.Rows Then
          Exit For
       End If
       If MSHFlexGrid6.TextMatrix(I, 6) = MSHFGodownDetails.TextMatrix(x, 7) And MSHFlexGrid6.TextMatrix(I, 7) = MSHFGodownDetails.TextMatrix(x, 8) Then
          If MSHFlexGrid6.TextMatrix(I, 1) <> "" Then
             For RO = I To MSHFlexGrid6.Rows - 2
                 For C = 0 To MSHFlexGrid6.Cols - 1
                     MSHFlexGrid6.TextMatrix(RO, C) = MSHFlexGrid6.TextMatrix(RO + 1, C)
                 Next
             Next
             MSHFlexGrid6.Rows = MSHFlexGrid6.Rows - 1
             I = I - 1
          End If
       End If
   Next
   
   MSHFGodownDetails.TextMatrix(x, 4) = ""
   MSHFGodownDetails.TextMatrix(x, 5) = ""
   MSHFGodownDetails.TextMatrix(x, 6) = ""
ElseIf MSHFGodownDetails.TextMatrix(x, 0) = strUnChecked Then
ExecuteProc:
   MSHFGodownDetails.TextMatrix(x, 0) = strChecked
   TxtCMPName.Text = MSHFGodownDetails.TextMatrix(x, 1)
   TxtGodownName.Text = MSHFGodownDetails.TextMatrix(x, 2)
   TxtGodownAddress.Text = MSHFGodownDetails.TextMatrix(x, 3)
   mSelGodownUnitTypeID = GetUnitTypeID(Val(MSHFGodownDetails.TextMatrix(x, 8)))
   Call CmbBillingCycle_GotFocus
   CmbBillingCycle.Text = MSHFGodownDetails.TextMatrix(x, 4)
   Call CmbBillingCycle_LostFocus
   Call CmbBillingUnit_GotFocus
   CmbBillingUnit.Text = MSHFGodownDetails.TextMatrix(x, 5)
   TxtRateOnQty.Text = Format(MSHFGodownDetails.TextMatrix(x, 6), "####.0000")
   FraGodownDetails.Enabled = True
   Call CmbBillingUnit_LostFocus
End If
End Sub
Private Sub ClearGodownDetails()
TxtCMPName.Text = ""
TxtGodownName.Text = ""
TxtGodownAddress.Text = ""
CmbBillingCycle.Text = ""
CmbBillingUnit.Text = ""
TxtTotDays.Text = ""
TxtRateOnQty.Text = ""
mSelGodownUnitTypeID = 0
'CmdAddGodownDetails.Caption = "Add To List"
End Sub
Private Sub CmbBillingCycle_GotFocus()
tmp = CmbBillingCycle.Text
CmbBillingCycle.Clear
Call TableMst_BillingCycle

If RsMst_BillingCycle.RecordCount = 0 Then
   MsgBox "No Record Found!!"
   Exit Sub
End If

For I = 1 To RsMst_BillingCycle.RecordCount
    CmbBillingCycle.AddItem RsMst_BillingCycle!BillingCycle
    RsMst_BillingCycle.MoveNext
Next
CmbBillingCycle.Text = tmp
End Sub

Private Sub CmbBillingCycle_LostFocus()
If CmbBillingCycle.Text = "" Then
   mBillingCycleID = Null
   Exit Sub
End If
RsMst_BillingCycle.MoveFirst
RsMst_BillingCycle.Find "BillingCycle='" & CmbBillingCycle.Text & "'"
If RsMst_BillingCycle.EOF Then
   MsgBox "Invalid Selection !!!"
   CmbBillingCycle.SetFocus
   Exit Sub
End If
mBillingCycleID = RsMst_BillingCycle!BillingCycleID
TxtTotDays = RsMst_BillingCycle!NoOfDays
End Sub

Private Sub CmbBillingUnit_GotFocus()
tmp = CmbBillingUnit.Text
CmbBillingUnit.Clear
CmbBillingUnit.AddItem "Bags"
CmbBillingUnit.AddItem "Quantity"
CmbBillingUnit.AddItem "Fixed"
CmbBillingUnit.AddItem "Percentage"
CmbBillingUnit.Text = tmp
End Sub

Private Sub CmbBillingUnit_LostFocus()
If CmbBillingUnit.Text = "" Then
   Exit Sub
End If
If LCase(CmbBillingUnit.Text) <> "bags" And LCase(CmbBillingUnit.Text) <> "quantity" And LCase(CmbBillingUnit.Text) <> "fixed" And LCase(CmbBillingUnit.Text) <> "percentage" Then
   MsgBox "Invalid Selection !!!"
   CmbBillingUnit.SetFocus
   Exit Sub
End If

If LCase(CmbBillingUnit.Text) = "bags" Then
   FraBagDetails.Visible = True
   FraBagDetails.Enabled = True
   Frame4.Visible = False
'   FraBagDetails.Visible = True
'   FraBagDetails.Left = 3480
'   FraBagDetails.Top = 2520
   TxtRateOnQty = ""
   TxtRateOnQty.Enabled = False
   If MSHFBagDetails.Rows <= 2 Then
      Call Grid_HeadBagDetails '1
      Call InDataBagSize
   End If
   Call Grid_HeadFixAmount '1
   TxtFromBagSize.SetFocus
ElseIf LCase(CmbBillingUnit.Text) = "fixed" Then
   If mSelGodownUnitTypeID <> 2 Then 'If mSelGodownUnitTypeID <> 6 Then
      MsgBox "Fixed Unit type is applicable for Franchisee godown only!!"
      CmbBillingUnit.SetFocus
      Exit Sub
   End If
   Frame4.Visible = True
   Frame4.Enabled = True
   FraBagDetails.Visible = False
   TxtRateOnQty = ""
   TxtRateOnQty.Enabled = False
   If MSHFBagDetails.Rows <= 2 Then
      Call Grid_HeadFixAmount '1
      Call InDataFixAmount
   End If
   Call Grid_HeadBagDetails
   Call ClearFixAmountDetails
   TxtFromDate.SetFocus
Else
   'FraBagDetails.Visible = False
   If MSHFBagDetails.Rows > 2 Then
      If MsgBox("Data in Bag Size Details will be lost. Are you want to Continue.", vbYesNo) = vbNo Then
         CmbBillingUnit.SetFocus
         Exit Sub
      End If
   End If
   
   If MSHFlexGrid5.Rows > 2 Then
      If MsgBox("Data in Fixed Rate Details will be lost. Are you want to Continue.", vbYesNo) = vbNo Then
         CmbBillingUnit.SetFocus
         Exit Sub
      End If
   End If
   
   Call Grid_HeadBagDetails '1
   Call ClearBagDetail
   'Frame2.Enabled = False
   FraBagDetails.Enabled = False
   Frame4.Enabled = False
   Frame4.Visible = False
   FraBagDetails.Visible = False
   TxtRateOnQty.Enabled = True
   TxtRateOnQty.SetFocus
End If
End Sub
Private Sub InDataBagSize()
'Call Grid_HeadBagDetails
Dim x As Variant
x = MSHFGodownDetails.RowSel
For I = 1 To MSHFBagTable.Rows - 2
    If MSHFGodownDetails.TextMatrix(x, 7) = MSHFBagTable.TextMatrix(I, 6) And MSHFGodownDetails.TextMatrix(x, 8) = MSHFBagTable.TextMatrix(I, 7) Then
       MSHFBagDetails.TextMatrix(MSHFBagDetails.Rows - 1, 0) = MSHFBagTable.TextMatrix(I, 3)
       MSHFBagDetails.TextMatrix(MSHFBagDetails.Rows - 1, 1) = MSHFBagTable.TextMatrix(I, 4)
       MSHFBagDetails.TextMatrix(MSHFBagDetails.Rows - 1, 2) = MSHFBagTable.TextMatrix(I, 5)
       MSHFBagDetails.TextMatrix(MSHFBagDetails.Rows - 1, 3) = MSHFBagTable.TextMatrix(I, 6)
       MSHFBagDetails.TextMatrix(MSHFBagDetails.Rows - 1, 4) = MSHFBagTable.TextMatrix(I, 7)
       MSHFBagDetails.TextMatrix(MSHFBagDetails.Rows - 1, 5) = MSHFBagTable.TextMatrix(I, 8)
       MSHFBagDetails.Rows = MSHFBagDetails.Rows + 1
    End If
Next
End Sub

Private Sub InDataFixAmount()

Dim x As Variant
x = MSHFGodownDetails.RowSel
For I = 1 To MSHFlexGrid6.Rows - 2
    If MSHFGodownDetails.TextMatrix(x, 7) = MSHFlexGrid6.TextMatrix(I, 6) And MSHFGodownDetails.TextMatrix(x, 8) = MSHFlexGrid6.TextMatrix(I, 7) Then
       MSHFlexGrid5.TextMatrix(MSHFlexGrid5.Rows - 1, 0) = MSHFlexGrid6.TextMatrix(I, 3)
       MSHFlexGrid5.TextMatrix(MSHFlexGrid5.Rows - 1, 1) = MSHFlexGrid6.TextMatrix(I, 4)
       MSHFlexGrid5.TextMatrix(MSHFlexGrid5.Rows - 1, 2) = MSHFlexGrid6.TextMatrix(I, 5)
       MSHFlexGrid5.TextMatrix(MSHFlexGrid5.Rows - 1, 3) = MSHFlexGrid6.TextMatrix(I, 6)
       MSHFlexGrid5.TextMatrix(MSHFlexGrid5.Rows - 1, 4) = MSHFlexGrid6.TextMatrix(I, 7)
       MSHFlexGrid5.TextMatrix(MSHFlexGrid5.Rows - 1, 5) = MSHFlexGrid6.TextMatrix(I, 8)
       MSHFlexGrid5.Rows = MSHFlexGrid5.Rows + 1
    End If
Next
End Sub

Private Sub ClearBagDetail()
TxtFromBagSize.Text = ""
TxtToBagSize.Text = ""
TxtRatePerBag.Text = ""
CmdAdd2List.Caption = "Add to List"
End Sub
Private Sub CmdAdd2List_Click()

If TxtFromBagSize.Text = "" Then
   MsgBox "Blank From Bag Size not Allowed!!!"
   TxtFromBagSize.SetFocus
   Exit Sub
End If

If TxtToBagSize.Text = "" Then
   MsgBox "Blank To Bag Size not Allowed!!!"
   TxtToBagSize.SetFocus
   Exit Sub
End If

If Val(TxtToBagSize.Text) < Val(TxtFromBagSize.Text) Then
   MsgBox "'To Bag Size' must be more than 'From Bag Size'"
   TxtToBagSize.SetFocus
   Exit Sub
End If

If TxtRatePerBag.Text = "" Then
   MsgBox "Blank Rate Per Bag not Allowed!!!"
   TxtRatePerBag.SetFocus
   Exit Sub
End If

If CheckDuplicateBags = True Then
   MsgBox "'Bag Size From' or 'Bag Size To' in already used in range. Please change range!!! "
   TxtFromBagSize.SetFocus
   Exit Sub
End If

If CmdAdd2List.Caption = "Update List" Then
   r = MSHFBagDetails.RowSel
Else
   r = MSHFBagDetails.Rows - 1
   MSHFBagDetails.Rows = MSHFBagDetails.Rows + 1
   MSHFBagDetails.TextMatrix(r, 5) = r ' MSHFBagTable.Rows - 1
End If

MSHFBagDetails.TextMatrix(r, 0) = Val(TxtFromBagSize)
MSHFBagDetails.TextMatrix(r, 1) = Val(TxtToBagSize)
MSHFBagDetails.TextMatrix(r, 2) = Format(Val(TxtRatePerBag.Text), "#####0.0000")
MSHFBagDetails.TextMatrix(r, 3) = MSHFGodownDetails.TextMatrix(MSHFGodownDetails.RowSel, 7)
MSHFBagDetails.TextMatrix(r, 4) = MSHFGodownDetails.TextMatrix(MSHFGodownDetails.RowSel, 8)

Call ClearBagDetail
TxtFromBagSize.SetFocus

End Sub
Function CheckDuplicateBags() As Boolean
Dim Retval As Boolean
Retval = False
If MSHFBagDetails.Rows > 1 Then
   For u = 1 To MSHFBagDetails.Rows - 1
       If (Val(TxtFromBagSize) >= Val(MSHFBagDetails.TextMatrix(u, 0)) And Val(TxtFromBagSize) <= Val(MSHFBagDetails.TextMatrix(u, 1))) Or Val(TxtToBagSize) >= Val(MSHFBagDetails.TextMatrix(u, 0)) And Val(TxtToBagSize) <= Val(MSHFBagDetails.TextMatrix(u, 1)) Or (Val(TxtFromBagSize) <= Val(MSHFBagDetails.TextMatrix(u, 0)) And Val(TxtToBagSize) >= Val(MSHFBagDetails.TextMatrix(u, 1))) Then
          If LCase(CmdAdd2List.Caption) = "add to list" Then
             Retval = True
             Exit For
          Else
             If u <> MSHFBagDetails.RowSel Then
                Retval = True
             End If
          End If
       End If
   Next
End If

CheckDuplicateBags = Retval

End Function
Private Sub CmdAddGodownDetails_Click()

If CmbBillingCycle.Text = "" Then
   MsgBox "Blank Billig Cycle not Allowed!!!"
   CmbBillingCycle.SetFocus
   Exit Sub
End If

If CmbBillingUnit.Text = "" Then
   MsgBox "Blank Billing Unit not Allowed!!!"
   CmbBillingUnit.SetFocus
   Exit Sub
End If

If LCase(CmbBillingUnit.Text) <> "bags" And LCase(CmbBillingUnit.Text) <> "fixed" Then
   If TxtRateOnQty.Text = "" Then
      MsgBox "Blank Rate On Qty not Allowed!!!"
      TxtRateOnQty.SetFocus
      Exit Sub
   End If
ElseIf LCase(CmbBillingUnit.Text) = "bags" Then
   If MSHFBagDetails.Rows = 2 Then
      MsgBox "Please Enter Bag Details !!!"
      FraBagDetails.Visible = True
      Exit Sub
   End If
ElseIf LCase(CmbBillingUnit.Text) = "fixed" Then
   If MSHFlexGrid5.Rows = 2 Then
      MsgBox "Please Enter Fixed Amount Details !!!"
      Frame4.Visible = True
      Exit Sub
   End If
End If

r = MSHFGodownDetails.RowSel ' MSHFGodownDetails.Rows - 1
MSHFGodownDetails.TextMatrix(r, 4) = CmbBillingCycle.Text
MSHFGodownDetails.TextMatrix(r, 5) = CmbBillingUnit.Text
MSHFGodownDetails.TextMatrix(r, 6) = Format(Val(TxtRateOnQty.Text), "#####0.0000")
MSHFGodownDetails.TextMatrix(r, 9) = mBillingCycleID
FraGodownDetails.Enabled = False
Dim mStatus As Boolean
mStatus = False
'MSHFBagTable.Visible = True
'FraBagDetails.Visible = False
If MSHFBagDetails.Rows > 2 Then
   For r = 1 To MSHFBagDetails.Rows - 2
       mStatus = False
       For I = 1 To MSHFBagTable.Rows - 2
           If MSHFBagTable.TextMatrix(I, 6) = MSHFBagDetails.TextMatrix(r, 3) And MSHFBagTable.TextMatrix(I, 7) = MSHFBagDetails.TextMatrix(r, 4) And MSHFBagTable.TextMatrix(I, 8) = MSHFBagDetails.TextMatrix(r, 5) Then
              MSHFBagTable.TextMatrix(I, 3) = MSHFBagDetails.TextMatrix(r, 0)
              MSHFBagTable.TextMatrix(I, 4) = MSHFBagDetails.TextMatrix(r, 1)
              MSHFBagTable.TextMatrix(I, 5) = MSHFBagDetails.TextMatrix(r, 2)
              mStatus = True
              Exit For
           End If
       Next
       If mStatus = False Then
          x = MSHFBagTable.Rows - 1
          MSHFBagTable.Rows = MSHFBagTable.Rows + 1
          MSHFBagTable.TextMatrix(x, 1) = GetCMPName(MSHFBagDetails.TextMatrix(r, 3))
          MSHFBagTable.TextMatrix(x, 2) = GetGodownName(MSHFBagDetails.TextMatrix(r, 4), "N")
          MSHFBagTable.TextMatrix(x, 3) = MSHFBagDetails.TextMatrix(r, 0) 'From
          MSHFBagTable.TextMatrix(x, 4) = MSHFBagDetails.TextMatrix(r, 1) 'To
          MSHFBagTable.TextMatrix(x, 5) = MSHFBagDetails.TextMatrix(r, 2) 'Rate
          MSHFBagTable.TextMatrix(x, 6) = MSHFBagDetails.TextMatrix(r, 3) 'CMPID
          MSHFBagTable.TextMatrix(x, 7) = MSHFBagDetails.TextMatrix(r, 4) 'GodownID
          MSHFBagTable.TextMatrix(x, 8) = MSHFBagDetails.TextMatrix(r, 5) 'ID
       End If
   Next
End If

If MSHFlexGrid5.Rows > 2 Then
   For r = 1 To MSHFlexGrid5.Rows - 2
       mStatus = False
       For I = 1 To MSHFlexGrid6.Rows - 2
           If MSHFlexGrid6.TextMatrix(I, 6) = MSHFlexGrid5.TextMatrix(r, 3) And MSHFlexGrid6.TextMatrix(I, 7) = MSHFlexGrid5.TextMatrix(r, 4) And MSHFlexGrid6.TextMatrix(I, 8) = MSHFlexGrid5.TextMatrix(r, 5) Then
              MSHFlexGrid6.TextMatrix(I, 3) = MSHFlexGrid5.TextMatrix(r, 0)
              MSHFlexGrid6.TextMatrix(I, 4) = MSHFlexGrid5.TextMatrix(r, 1)
              MSHFlexGrid6.TextMatrix(I, 5) = MSHFlexGrid5.TextMatrix(r, 2)
              mStatus = True
              Exit For
           End If
       Next
       If mStatus = False Then
          x = MSHFlexGrid6.Rows - 1
          MSHFlexGrid6.Rows = MSHFlexGrid6.Rows + 1
          MSHFlexGrid6.TextMatrix(x, 1) = GetCMPName(MSHFlexGrid5.TextMatrix(r, 3))
          MSHFlexGrid6.TextMatrix(x, 2) = GetGodownName(MSHFlexGrid5.TextMatrix(r, 4), "N")
          MSHFlexGrid6.TextMatrix(x, 3) = MSHFlexGrid5.TextMatrix(r, 0) 'From
          MSHFlexGrid6.TextMatrix(x, 4) = MSHFlexGrid5.TextMatrix(r, 1) 'To
          MSHFlexGrid6.TextMatrix(x, 5) = MSHFlexGrid5.TextMatrix(r, 2) 'Rate
          MSHFlexGrid6.TextMatrix(x, 6) = MSHFlexGrid5.TextMatrix(r, 3) 'CMPID
          MSHFlexGrid6.TextMatrix(x, 7) = MSHFlexGrid5.TextMatrix(r, 4) 'GodownID
          MSHFlexGrid6.TextMatrix(x, 8) = MSHFlexGrid5.TextMatrix(r, 5) 'ID
       End If
   Next
End If

Call Grid_HeadBagDetails
Call Grid_HeadFixAmount
Call ClearGodownDetails
End Sub

Private Sub MSHFlexGrid1_Click()
Dim x As Variant
x = MSHFlexGrid1.RowSel
If MSHFlexGrid1.TextMatrix(x, 0) = "" Then Exit Sub
RsMst_FF_RateCard.MoveFirst
RsMst_FF_RateCard.Find "FF_RateCardID=" & MSHFlexGrid1.TextMatrix(x, 0) & ""
Call Indata
Frame0.Visible = False
End Sub

Private Sub MSHFlexGrid5_Click()

Dim I As Variant
If SaveCmd.Enabled = False Then Exit Sub
I = MSHFlexGrid5.RowSel

If I < 1 Or I = MSHFlexGrid5.Rows - 1 Then
   Call ClearFixAmountDetails
   Exit Sub
End If

TxtFixAmount = Val(MSHFlexGrid5.TextMatrix(I, 2))

'TxtRemarksFixAmount = Trim(MSHFlexGrid5.TextMatrix(I, 3))
If MSHFlexGrid5.TextMatrix(I, 0) <> "" Then
   TxtFromDate.Value = CDate(MSHFlexGrid5.TextMatrix(I, 0))
Else
   TxtFromDate.Value = Date 'CDate(MSHFlexGrid3.TextMatrix(I, 5))
End If

If MSHFlexGrid5.TextMatrix(I, 1) <> "" Then
   TxtToDate.Value = CDate(MSHFlexGrid5.TextMatrix(I, 1))
Else
   TxtToDate.Value = Date 'CDate(MSHFlexGrid3.TextMatrix(I, 6))
End If

CmdAdd2ListFixAmount.Caption = "Update List"
TxtFromDate.SetFocus


End Sub

Private Sub SaveCmd_Click()

'Write Code For Godown Details Table to check weather details for bags is added or not

If CmbLocationID.Text = "" Then
   MsgBox "Blank Locaiton Name Not Allowed !!"
   CmbLocationID.SetFocus
   Exit Sub
End If

If CmbLessorName.Text = "" Then
   MsgBox "Blank Franchisee Name Not Allowed !!"
   CmbLessorName.SetFocus
   Exit Sub
End If

If CmbCommodityGroupID.Text = "" Then
   MsgBox "Blank Commodity Not Allowed !!"
   CmbCommodityGroupID.SetFocus
   Exit Sub
End If

If CmbExchangeTypeID.Text = "" Then
   MsgBox "Blank Exchange Type Not Allowed !!"
   CmbExchangeTypeID.SetFocus
   Exit Sub
End If

If MSHFGodownDetails.Rows = 2 Then
   MsgBox " No CMP And Godown Details Found !!"
   CmbLocationID.SetFocus
   Exit Sub
End If

If CheckGrid = True Then
   MsgBox "Please Enter Details for CMP and Godown From List !!"
   Exit Sub
End If

If CmbStatus.Text = "" Then
   MsgBox "Blank Status Not Allowed !!! "
   CmbStatus.SetFocus
   Exit Sub
End If

If Edit_Flg = "A" Then
   
   'RsMst_FF_RateCard.Filter = "LocationID=" & mLocationID & "  And CommodityGroupID=" & mCommodityGroupId & " and LessorID=" & mLessorID & " And ExchangeTypeID=" & mExchangeTypeID & ""
   RsMst_FF_RateCard.Filter = "LocationID=" & mLocationID & "  And CommodityGroupID=" & mCommodityGroupId & " and VendorID=" & mVendorID & " And ExchangeTypeID=" & mExchangeTypeID & " And Flag = 'A' "
   If RsMst_FF_RateCard.RecordCount > 0 Then
      RsMst_FF_RateCard.Filter = ""
      MsgBox "Duplicate Record Not Allowed !!"
      Exit Sub
   Else
      RsMst_FF_RateCard.Filter = ""
      RsMst_FF_RateCard.AddNew
'      TxtFGRCID = GetMaxID("FF_RateCardID", "Mst_FF_RateCard")
'      RsMst_FF_RateCard!FF_RateCardID = TxtFGRCID.Text
      RsMst_FF_RateCard!G_UID = GetGUID
      mG_UID = RsMst_FF_RateCard!G_UID
   End If
Else 'Edit
   RsMst_FF_RateCard.MoveFirst
   'RsMst_FF_RateCard.Filter = "LocationID = " & mLocationID & " And LessorID =  " & mLessorID & " And CommodityGroupID = " & mCommodityGroupId & " And ExchangeTypeID=" & mExchangeTypeID & " "
   
   RsMst_FF_RateCard.Filter = "LocationID = " & mLocationID & " And VendorID =  " & mVendorID & " And CommodityGroupID = " & mCommodityGroupId & " And ExchangeTypeID=" & mExchangeTypeID & " And Flag = 'A'  "
   If RsMst_FF_RateCard.RecordCount > 0 Then
      If RsMst_FF_RateCard!FF_RateCardID <> TxtFGRCID Then
         MsgBox "Duplicate Franchisee Godown Rate Card Not allowed !!!"
         RsMst_FF_RateCard.Filter = ""
         Exit Sub
      End If
   End If
   RsMst_FF_RateCard.Filter = ""
   RsMst_FF_RateCard.MoveFirst
   RsMst_FF_RateCard.Find " FF_RateCardID= '" & Val(TxtFGRCID.Text) & "'"
   Call DeleteChildTableDetail
End If
RsMst_FF_RateCard!LocationID = mLocationID
RsMst_FF_RateCard!CommodityGroupID = mCommodityGroupId
'RsMst_FF_RateCard!LessorID = mLessorID
RsMst_FF_RateCard!VendorID = mVendorID
RsMst_FF_RateCard!ExchangeTypeID = mExchangeTypeID
RsMst_FF_RateCard!Flag = Left$(CmbStatus.Text, 1)

If IsNull(RsMst_FF_RateCard!CreatedBy) = True Or RsMst_FF_RateCard!CreatedBy = "" Then
   RsMst_FF_RateCard!CreatedBy = Gen_UserLoginID
   RsMst_FF_RateCard!CreatedDate = Now
Else
   RsMst_FF_RateCard!UpdatedBy = Gen_UserLoginID
   RsMst_FF_RateCard!UpdatedDate = Now
End If


RsMst_FF_RateCard.Update


tmp = "Select FF_RateCardID From Mst_FF_RateCard Where G_UID='" & mG_UID & "'"
Call TmpTable

If TmpRs.RecordCount > 0 Then
'   mCxTF_DetailsID = TmpRs!CxTF_DetailsID
   TxtFGRCID.Text = TmpRs!FF_RateCardID
End If

Call SaveGodown_BagDetails

Call Indata
Call GButtonLock(Me, True)

End Sub
Private Sub DeleteChildTableDetail()

tmp = "Delete Mst_FF_RateCard_Godown_Fix Where FF_RateCard_GodownID In (Select FF_RateCard_GodownID from Mst_FF_RateCard_Godown Where FF_RateCardID=" & Val(TxtFGRCID.Text) & ")"
Call TmpTable

tmp = "Delete Mst_FF_RateCard_Godown_Bag Where FF_RateCard_GodownID In (Select FF_RateCard_GodownID from Mst_FF_RateCard_Godown Where FF_RateCardID=" & Val(TxtFGRCID.Text) & ")"
Call TmpTable

tmp = "Delete Mst_FF_RateCard_Godown Where FF_RateCardID=" & Val(TxtFGRCID.Text) & " "
Call TmpTable
End Sub
Private Sub SaveGodown_BagDetails()
Dim tID As Variant

Call TableMst_FF_RateCard_Godown
Call TableMst_FF_RateCard_Godown_Bag
Call TableMst_FF_RateCard_Godown_Fix
For I = 1 To MSHFGodownDetails.Rows - 2
    If MSHFGodownDetails.TextMatrix(I, 0) = strChecked Then
       RsMst_FF_RateCard_Godown.AddNew
'       tID = GetMaxID("FF_RateCard_GodownID", "Mst_FF_RateCard_Godown")
'       RsMst_FF_RateCard_Godown!FF_RateCard_GodownID = tID
       RsMst_FF_RateCard_Godown!FF_RateCardID = Val(TxtFGRCID.Text)
       RsMst_FF_RateCard_Godown!CMPID = Val(MSHFGodownDetails.TextMatrix(I, 7))
       RsMst_FF_RateCard_Godown!GodownID = Val(MSHFGodownDetails.TextMatrix(I, 8))
       RsMst_FF_RateCard_Godown!BillingCycleID = Val(MSHFGodownDetails.TextMatrix(I, 9))
       RsMst_FF_RateCard_Godown!BillingUnit = Left(MSHFGodownDetails.TextMatrix(I, 5), 1)
       RsMst_FF_RateCard_Godown!RateOnQty = Val(MSHFGodownDetails.TextMatrix(I, 6))
       RsMst_FF_RateCard_Godown!G_UID = GetGUID
       mGGGUID = RsMst_FF_RateCard_Godown!G_UID
       RsMst_FF_RateCard_Godown.Update
       
       
        tmp = "Select FF_RateCard_GodownID From Mst_FF_RateCard_Godown Where G_UID='" & mGGGUID & "'"
        Call TmpTable
        
        If TmpRs.RecordCount > 0 Then
        '   mCxTF_DetailsID = TmpRs!CxTF_DetailsID
           TtID = TmpRs!FF_RateCard_GodownID
        End If
       
       If LCase(MSHFGodownDetails.TextMatrix(I, 5)) = "bags" Then
          For j = 1 To MSHFBagTable.Rows - 2
              If MSHFBagTable.TextMatrix(j, 6) = Val(MSHFGodownDetails.TextMatrix(I, 7)) And MSHFBagTable.TextMatrix(j, 7) = Val(MSHFGodownDetails.TextMatrix(I, 8)) Then
                 RsMst_FF_RateCard_Godown_Bag.AddNew
 '                RsMst_FF_RateCard_Godown_Bag!FF_RateCard_Godown_BagID = GetMaxID("FF_RateCard_Godown_BagID", "Mst_FF_RateCard_Godown_Bag")
                 RsMst_FF_RateCard_Godown_Bag!FF_RateCard_GodownID = TtID
                 RsMst_FF_RateCard_Godown_Bag!BagSizeFrom = Val(MSHFBagTable.TextMatrix(j, 3))
                 RsMst_FF_RateCard_Godown_Bag!BagSizeTo = Val(MSHFBagTable.TextMatrix(j, 4))
                 RsMst_FF_RateCard_Godown_Bag!RatePerBag = Val(MSHFBagTable.TextMatrix(j, 5))
                 RsMst_FF_RateCard_Godown_Bag.Update
              End If
          Next
       ElseIf LCase(MSHFGodownDetails.TextMatrix(I, 5)) = "fixed" Then
          For j = 1 To MSHFlexGrid6.Rows - 2
              If MSHFlexGrid6.TextMatrix(j, 6) = Val(MSHFGodownDetails.TextMatrix(I, 7)) And MSHFlexGrid6.TextMatrix(j, 7) = Val(MSHFGodownDetails.TextMatrix(I, 8)) Then
                 RsMst_FF_RateCard_Godown_Fix.AddNew
 '                RsMst_FF_RateCard_Godown_Fix!FF_RateCard_Godown_FixID = GetMaxID("FF_RateCard_Godown_FixID", "Mst_FF_RateCard_Godown_Fix")
                 RsMst_FF_RateCard_Godown_Fix!FF_RateCard_GodownID = TtID
                 RsMst_FF_RateCard_Godown_Fix!StartDate = CDate(MSHFlexGrid6.TextMatrix(j, 3))
                 RsMst_FF_RateCard_Godown_Fix!EndDate = CDate(MSHFlexGrid6.TextMatrix(j, 4))
                 RsMst_FF_RateCard_Godown_Fix!FixAmount = Val(MSHFlexGrid6.TextMatrix(j, 5))
                 RsMst_FF_RateCard_Godown_Fix.Update
              End If
          Next
       End If
    End If
Next
End Sub
Private Function GetMaxID(mColumn_ As Variant, mTable_ As Variant)
Dim Retval As Double
tmp = "Select Max(" & mColumn_ & ") from " & mTable_ & ""
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 1, TmpRs.Fields(0) + 1)
GetMaxID = Retval
End Function
Private Function CheckGrid() As Boolean
Dim Retval As Boolean
Dim mStatus As Variant
Retval = False
mStatus = 1
For I = 1 To MSHFGodownDetails.Rows - 1
    If MSHFGodownDetails.TextMatrix(I, 0) = strChecked Then
       If MSHFGodownDetails.TextMatrix(I, 4) = "" Then
          Retval = True 'msg
          'mStatus = mStatus + 1
          Exit For
       End If
    ElseIf MSHFGodownDetails.TextMatrix(I, 0) = strUnChecked Then
       mStatus = mStatus + 1
    End If
Next
If mStatus = MSHFGodownDetails.Rows - 1 Then
   Retval = True
End If
CheckGrid = Retval
End Function
Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_FF_RateCard.MoveNext
If RsMst_FF_RateCard.EOF Then
   RsMst_FF_RateCard.MoveLast
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
RsMst_FF_RateCard.MovePrevious
If RsMst_FF_RateCard.BOF Then
   RsMst_FF_RateCard.MoveFirst
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
RsMst_FF_RateCard.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsMst_FF_RateCard.MoveLast
Call Indata
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False
End Sub
Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   If RsMst_FF_RateCard.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   Call Indata
   Exit Sub
End If


Frame0.Left = Frame1.Left
Frame0.Top = Frame1.Top
Frame0.Height = Me.Height - 700
Frame0.Width = Me.Width - 300
MSHFlexGrid1.Height = Frame0.Height - 1300
MSHFlexGrid1.Width = Frame0.Width - 300
Frame0.Visible = True

''===========comment on 7 july 2012 for not clearing search grid=============
'Call Grid_HeadSearch
'CmbSLocationID.Text = ""
'CmbSLessorName.Text = ""
'CmbCommodityGroupID.Text = ""
'LblTotalRecords.Caption = "-"
''CmbSLocationID.SetFocus
''===========comment on 7 july 2012 for not clearing search grid=============

End Sub




Private Sub TxtFromBagSize_LostFocus()
If TxtFromBagSize.Text = "" Then Exit Sub
If IsNumeric(TxtFromBagSize.Text) = False Or Val(TxtFromBagSize.Text) < 0 Then
   MsgBox "Invalid Numeric Format !!!"
   TxtFromBagSize.SetFocus
   Exit Sub
End If
End Sub

Private Sub TxtRateOnQty_LostFocus()
If TxtRateOnQty.Text = "" Then Exit Sub
If IsNumeric(TxtRateOnQty.Text) = False Or Val(TxtRateOnQty.Text) < 0 Then
   MsgBox "Invalid Numeric Format!!!"
   TxtRateOnQty.SetFocus
   Exit Sub
End If
TxtRateOnQty = Round(TxtRateOnQty, 4)
End Sub

Private Sub TxtRatePerBag_LostFocus()
If TxtRatePerBag.Text = "" Then Exit Sub
If IsNumeric(TxtRatePerBag.Text) = False Or Val(TxtRatePerBag.Text) < 0 Then
   MsgBox "Invalid Numeric Format!!!"
   TxtRatePerBag.SetFocus
   Exit Sub
End If
TxtRatePerBag.Text = Format(TxtRatePerBag.Text, "#####0.00")
End Sub

Private Sub TxtToBagSize_LostFocus()
If TxtToBagSize.Text = "" Then Exit Sub
If IsNumeric(TxtToBagSize.Text) = False Or Val(TxtToBagSize.Text) < 0 Then
   MsgBox "Invalid Numeric Format !!!"
   TxtToBagSize.SetFocus
   Exit Sub
End If
If InStr(TxtToBagSize, ".") Then
   MsgBox "Value will get rounded !!!"
End If


TxtToBagSize = Round(TxtToBagSize, 0)
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Mst_FranchiseeGodownRateCard.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Mst_FranchiseeGodownRateCard.xls")
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
Private Sub CmbSLocationID_GotFocus()
tmp = CmbSLocationID.Text
If Gen_WcLocation = "" Then
   Call TableMst_Location
Else
   Call TableMst_Location(" Where " & Gen_WcLocation)
End If

CmbSLocationID.Clear
If RsMst_Location.RecordCount = 0 Then
  MsgBox "No Record found !!!! "
  Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
   CmbSLocationID.AddItem RsMst_Location!Locationname
   RsMst_Location.MoveNext
Next
CmbSLocationID.Text = tmp
    
End Sub
'----
Private Sub CmbSLocationID_LostFocus()
If CmbSLocationID.Text = "" Then
   Exit Sub
End If

RsMst_Location.MoveFirst
RsMst_Location.Find "Locationname = '" & CmbSLocationID.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection "
   CmbSLocationID.SetFocus
   Exit Sub
End If
mSLocationID = RsMst_Location!LocationID

End Sub
'Private Sub CmbSLessorName_GotFocus()
'tmp = CmbSLessorName.Text
'mSLessorID = ""
'Call TableMst_Lessor("")
'CmbSLessorName.Clear
'If RsMst_Lessor.RecordCount = 0 Then
'   MsgBox "No Record found !!!! "
'   Exit Sub
'End If
'For I = 1 To RsMst_Lessor.RecordCount
'    CmbSLessorName.AddItem RsMst_Lessor!LessorName
'    RsMst_Lessor.MoveNext
'Next
'CmbSLessorName.Text = tmp
'End Sub
'
'Private Sub CmbSLessorName_LostFocus()
'If CmbSLessorName.Text = "" Then
'   Exit Sub
'End If
'RsMst_Lessor.MoveFirst
'RsMst_Lessor.Find "LessorName = '" & CmbSLessorName.Text & "'"
'If RsMst_Lessor.EOF Then
'   MsgBox "Invalid selection "
'   CmbSLessorName.SetFocus
'   Exit Sub
'End If
'mSLessorID = RsMst_Lessor!LessorID
'End Sub

Private Sub CmbSLessorName_GotFocus()
tmp = CmbSLessorName.Text
mSVendorID = ""
Call TableMst_Vendor
CmbSLessorName.Clear

If RsMst_Vendor.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Vendor.RecordCount
    CmbSLessorName.AddItem RsMst_Vendor!VendorName
    RsMst_Vendor.MoveNext
Next
CmbSLessorName.Text = tmp
End Sub

Private Sub CmbSLessorName_LostFocus()
If CmbSLessorName.Text = "" Then
   Exit Sub
End If
RsMst_Vendor.MoveFirst
RsMst_Vendor.Find "VendorName = '" & CmbSLessorName.Text & "'"
If RsMst_Vendor.EOF Then
   MsgBox "Invalid selection "
   CmbSLessorName.SetFocus
   Exit Sub
End If
mSVendorID = RsMst_Vendor!VendorID
End Sub

Private Sub CmbSCommodityGroupID_GotFocus()
tmp = CmbSCommodityGroupID.Text
Call TableMst_CommodityGroup
CmbSCommodityGroupID.Clear
If RsMst_CommodityGroup.RecordCount = 0 Then
  MsgBox "No Record found !!!! "
  Exit Sub
End If
For I = 1 To RsMst_CommodityGroup.RecordCount
   CmbSCommodityGroupID.AddItem RsMst_CommodityGroup!COmmodityGroup
   RsMst_CommodityGroup.MoveNext
Next
CmbSCommodityGroupID.Text = tmp
    
End Sub

'----
Private Sub CmbSCommodityGroupID_LostFocus()
If CmbSCommodityGroupID.Text = "" Then
   Exit Sub
End If

RsMst_CommodityGroup.MoveFirst
RsMst_CommodityGroup.Find "CommodityGroup = '" & CmbSCommodityGroupID.Text & "'"
If RsMst_CommodityGroup.EOF Then
   MsgBox "Invalid selection "
   CmbSCommodityGroupID.SetFocus
   Exit Sub
End If
mSCommodityGroupId = RsMst_CommodityGroup!CommodityGroupID
End Sub

Private Sub ClearFixAmountDetails()

TxtFromDate.Value = Date
TxtToDate.Value = Date + 9132
TxtFixAmount = ""
CmdAdd2ListFixAmount.Caption = "Add to List"
End Sub

Private Sub Grid_HeadFixAmount()
With MSHFlexGrid5
     .Clear
     .Rows = 2
     .Cols = 6
     .AllowUserResizing = flexResizeColumns
     .Font.Size = 10
     .TextMatrix(0, 0) = "Start Date"
     .TextMatrix(0, 1) = "End Date"
     .TextMatrix(0, 2) = "Fix Amount"
     .TextMatrix(0, 3) = "CMPID"
     .TextMatrix(0, 4) = "GodownID"
     .TextMatrix(0, 5) = "ID"
     
     .ColWidth(0) = 1700
     .ColWidth(1) = 1700
     .ColWidth(2) = 1500
     .ColWidth(3) = 0
     .ColWidth(4) = 0
     .ColWidth(5) = 0
     
     .RowHeight(0) = 700
End With
End Sub

Function CheckDuplicateFixAmountDateRange() As Boolean
Dim Retval As Boolean
Retval = False
If MSHFlexGrid5.Rows > 1 Then
   For u = 1 To MSHFlexGrid5.Rows - 1
       If MSHFlexGrid5.TextMatrix(u, 0) = "" Then Exit For
       If (TxtFromDate.Value >= CDate(MSHFlexGrid5.TextMatrix(u, 0)) And TxtFromDate.Value <= CDate(MSHFlexGrid5.TextMatrix(u, 1))) Or TxtToDate.Value >= CDate(MSHFlexGrid5.TextMatrix(u, 0)) And TxtToDate.Value <= CDate(MSHFlexGrid5.TextMatrix(u, 1)) Or (TxtFromDate.Value <= CDate(MSHFlexGrid5.TextMatrix(u, 0)) And TxtToDate.Value >= CDate(MSHFlexGrid5.TextMatrix(u, 1))) Then
          If LCase(CmdAdd2ListFixAmount.Caption) = "add to list" Then
             Retval = True
             Exit For
          Else
             If u <> MSHFlexGrid5.RowSel Then
                Retval = True
             End If
          End If
       End If
   Next
End If

CheckDuplicateFixAmountDateRange = Retval

End Function

Private Sub Grid_HeadFixTable()
With MSHFlexGrid6
     .Clear
     .Cols = 9
     .Rows = 2
     .AllowUserResizing = flexResizeColumns
     .Font.Size = 9
     .WordWrap = True
     .TextMatrix(0, 0) = "FGRCID"
     .TextMatrix(0, 1) = "CMP"
     .TextMatrix(0, 2) = "Godown"
     .TextMatrix(0, 3) = "From Date "
     .TextMatrix(0, 4) = "To Date"
     .TextMatrix(0, 5) = "Fixed Amount"
     .TextMatrix(0, 6) = "CMPID"
     .TextMatrix(0, 7) = "GodownID"
     .TextMatrix(0, 8) = "ID"
     
     .ColWidth(0) = 0
     .ColWidth(1) = 1600
     .ColWidth(2) = 1400
     .ColWidth(3) = 900
     .ColWidth(4) = 900
     .ColWidth(5) = 900
     .ColWidth(6) = 0
     .ColWidth(7) = 0
     .ColWidth(8) = 0
     .RowHeight(0) = 500
End With
End Sub

Private Sub TxtFixAmount_LostFocus()
If TxtFixAmount.Text = "" Then Exit Sub
If IsNumeric(TxtFixAmount.Text) = False Or Val(TxtFixAmount.Text) < 0 Then
   MsgBox "Invalid Numeric Format!!!"
   TxtFixAmount.SetFocus
   Exit Sub
End If
TxtFixAmount.Text = Format(Val(TxtFixAmount.Text), "#####0.00")
End Sub

Private Function GetUnitTypeID(GodownID_ As Double) As Double
Dim Retval As Double
Retval = 0
Call TableMst_Godown("Where GodownID = " & GodownID_)
If Not RsMst_Godown.EOF Then
   Retval = RsMst_Godown!UnitTypeID
End If
GetUnitTypeID = Retval
End Function

Private Sub CmbStatus_LostFocus()
If CmbStatus.Text = "" Then Exit Sub
If LCase(CmbStatus.Text) <> "active" And LCase(CmbStatus.Text) <> "deactive" Then
   MsgBox "Invalid Selection!!!"
   CmbStatus.SetFocus
   Exit Sub
End If
End Sub

