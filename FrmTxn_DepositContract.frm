VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmTxn_DepositContract 
   Caption         =   "Deposit Contract"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   FontTransparent =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0FF&
      Height          =   5655
      Left            =   5040
      TabIndex        =   113
      Top             =   3360
      Visible         =   0   'False
      Width           =   5535
      Begin VB.TextBox TxtBillingRateMonthlyMTRate_Min 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   310
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   189
         TabStop         =   0   'False
         Top             =   3240
         Width           =   1410
      End
      Begin VB.TextBox TxtGodownRentPSFeet_Min 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   310
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   188
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1410
      End
      Begin VB.TextBox TxtUtilizationLevel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   310
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   141
         TabStop         =   0   'False
         Top             =   5040
         Width           =   1530
      End
      Begin VB.TextBox TxtFranchiseRatio 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   310
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   140
         TabStop         =   0   'False
         Top             =   4680
         Width           =   1530
      End
      Begin VB.TextBox TxtMonthYear 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   310
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   130
         TabStop         =   0   'False
         Top             =   4320
         Width           =   1410
      End
      Begin VB.TextBox TxtRentPaidDate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   310
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   128
         TabStop         =   0   'False
         Top             =   3960
         Width           =   1410
      End
      Begin VB.TextBox TxtBillingRateMonthlyMTRate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   310
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   125
         TabStop         =   0   'False
         Top             =   3240
         Width           =   1530
      End
      Begin VB.TextBox TxtDivisionFactor 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   310
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   123
         TabStop         =   0   'False
         Top             =   2520
         Width           =   1530
      End
      Begin VB.TextBox TxtBagEquivalent 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   310
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   122
         TabStop         =   0   'False
         Top             =   2160
         Width           =   1530
      End
      Begin VB.TextBox TxtConversionFactor 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   310
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   121
         TabStop         =   0   'False
         Top             =   1800
         Width           =   1530
      End
      Begin VB.TextBox TxtGodownLastPNL 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   310
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   129
         TabStop         =   0   'False
         Top             =   4320
         Width           =   1530
      End
      Begin VB.TextBox TxtGodownLastRentPaid 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   310
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   127
         TabStop         =   0   'False
         Top             =   3960
         Width           =   1530
      End
      Begin VB.TextBox TxtGodownRMMargin 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   310
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   126
         TabStop         =   0   'False
         Top             =   3600
         Width           =   1530
      End
      Begin VB.TextBox TxtFumigationCharges 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   310
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   124
         TabStop         =   0   'False
         Top             =   2880
         Width           =   1530
      End
      Begin VB.TextBox TxtGodownRentPSFeet 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   310
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   120
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1530
      End
      Begin VB.TextBox TxtGodownArea 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   310
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   119
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1530
      End
      Begin VB.TextBox TxtGodownRent 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   310
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   118
         TabStop         =   0   'False
         Top             =   720
         Width           =   1530
      End
      Begin VB.CommandButton CmbCloseDetails 
         Caption         =   "Close"
         Height          =   375
         Left            =   4320
         TabIndex        =   132
         Top             =   4920
         Width           =   855
      End
      Begin VB.Label Label47 
         BackStyle       =   0  'Transparent
         Caption         =   "Godown No :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   144
         Top             =   450
         Width           =   5295
      End
      Begin VB.Label Label46 
         BackStyle       =   0  'Transparent
         Caption         =   "CMP Name :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   143
         Top             =   150
         Width           =   5295
      End
      Begin VB.Label Label43 
         BackStyle       =   0  'Transparent
         Caption         =   "Godown Stock Utilization"
         Height          =   255
         Left            =   120
         TabIndex        =   142
         Top             =   5070
         Width           =   2175
      End
      Begin VB.Label Label42 
         BackStyle       =   0  'Transparent
         Caption         =   "Franchisee Ratio"
         Height          =   255
         Left            =   120
         TabIndex        =   139
         Top             =   4710
         Width           =   2175
      End
      Begin VB.Label Label41 
         BackStyle       =   0  'Transparent
         Caption         =   "Billing Rate Monthly MT Rate"
         Height          =   255
         Left            =   120
         TabIndex        =   138
         Top             =   3270
         Width           =   2175
      End
      Begin VB.Label Label40 
         BackStyle       =   0  'Transparent
         Caption         =   "Division Factor (%)"
         Height          =   255
         Left            =   120
         TabIndex        =   137
         Top             =   2550
         Width           =   2175
      End
      Begin VB.Label Label39 
         BackStyle       =   0  'Transparent
         Caption         =   "Bag Equivalent"
         Height          =   255
         Left            =   120
         TabIndex        =   136
         Top             =   2190
         Width           =   2175
      End
      Begin VB.Label Label38 
         BackStyle       =   0  'Transparent
         Caption         =   "Conversion Factor"
         Height          =   255
         Left            =   120
         TabIndex        =   135
         Top             =   1830
         Width           =   2175
      End
      Begin VB.Label Label37 
         BackStyle       =   0  'Transparent
         Caption         =   "Godown Last PNL Status"
         Height          =   255
         Left            =   120
         TabIndex        =   134
         Top             =   4350
         Width           =   2175
      End
      Begin VB.Label Label36 
         BackStyle       =   0  'Transparent
         Caption         =   "Godown Last Rent Paid"
         Height          =   255
         Left            =   120
         TabIndex        =   133
         Top             =   3990
         Width           =   2175
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         Caption         =   "RM Approval Margin (%)"
         Height          =   255
         Left            =   120
         TabIndex        =   131
         Top             =   3630
         Width           =   2175
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Other + Fumigation Charges"
         Height          =   255
         Left            =   120
         TabIndex        =   117
         Top             =   2910
         Width           =   2175
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "Rent (Rupees Per Sq. Ft.)"
         Height          =   255
         Left            =   120
         TabIndex        =   116
         Top             =   1470
         Width           =   2175
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Godown Area (Sq. Ft.)"
         Height          =   255
         Left            =   120
         TabIndex        =   115
         Top             =   1110
         Width           =   2175
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Godown Rent"
         Height          =   255
         Left            =   120
         TabIndex        =   114
         Top             =   750
         Width           =   2175
      End
   End
   Begin VB.Frame Frame6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   120
      TabIndex        =   163
      Top             =   4080
      Visible         =   0   'False
      Width           =   15030
      Begin VB.TextBox TxtVerifiedRemark 
         Appearance      =   0  'Flat
         Height          =   3480
         Left            =   120
         MaxLength       =   2000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   166
         Top             =   480
         Width           =   14685
      End
      Begin VB.CommandButton CmdVerifiedSave 
         Caption         =   "Verify"
         Height          =   400
         Left            =   12180
         TabIndex        =   165
         Top             =   4080
         Width           =   1380
      End
      Begin VB.CommandButton CmdVerifiedCancel 
         Caption         =   "Cancel"
         Height          =   400
         Left            =   13560
         TabIndex        =   164
         Top             =   4080
         Width           =   1260
      End
      Begin VB.Label Label53 
         Caption         =   "Verified Remark"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   167
         Top             =   180
         Width           =   1815
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Add New Bag Size"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      TabIndex        =   152
      Top             =   3480
      Visible         =   0   'False
      Width           =   14775
      Begin VB.ComboBox CmbGodownID 
         Height          =   315
         Left            =   4800
         Sorted          =   -1  'True
         TabIndex        =   147
         Top             =   645
         Width           =   4605
      End
      Begin VB.ComboBox CmbCMPID 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   146
         Top             =   645
         Width           =   4605
      End
      Begin VB.TextBox TxtMaxNewBagSize 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   310
         Left            =   11040
         MaxLength       =   10
         TabIndex        =   149
         Top             =   645
         Width           =   1395
      End
      Begin VB.TextBox TxtMinNewBagSize 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   310
         Left            =   9480
         MaxLength       =   10
         TabIndex        =   148
         Top             =   645
         Width           =   1395
      End
      Begin VB.CommandButton CmbCloseNewBagSizeFrame 
         Caption         =   "Close"
         Height          =   315
         Left            =   13575
         TabIndex        =   151
         Top             =   645
         Width           =   1080
      End
      Begin VB.CommandButton CmdAddtoListNewBagSize 
         Caption         =   "Add to List"
         Height          =   315
         Left            =   12480
         TabIndex        =   150
         Top             =   645
         Width           =   1095
      End
      Begin VB.Label Label51 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Maximum Bag Size"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11040
         TabIndex        =   156
         Top             =   240
         Width           =   1395
      End
      Begin VB.Label Label50 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Minimum Bag Size"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9480
         TabIndex        =   155
         Top             =   240
         Width           =   1395
      End
      Begin VB.Label Label49 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Godown No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4800
         TabIndex        =   154
         Top             =   360
         Width           =   4605
      End
      Begin VB.Label Label48 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "CMP Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   153
         Top             =   330
         Width           =   4605
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Close Deposit Contract"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   105
      TabIndex        =   110
      Top             =   4080
      Visible         =   0   'False
      Width           =   15030
      Begin VB.CommandButton CmdCloseCancel 
         Caption         =   "Cancel"
         Height          =   400
         Left            =   13560
         TabIndex        =   52
         Top             =   4080
         Width           =   1260
      End
      Begin VB.CommandButton CmdCloseConform 
         Caption         =   "Close"
         Height          =   400
         Left            =   12180
         TabIndex        =   51
         Top             =   4080
         Width           =   1380
      End
      Begin VB.TextBox TxtCloseRemark 
         Appearance      =   0  'Flat
         Height          =   3480
         Left            =   120
         MaxLength       =   8000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   50
         Top             =   480
         Width           =   14685
      End
      Begin VB.Label Label17 
         Caption         =   "Closing Remarks"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   111
         Top             =   300
         Width           =   1815
      End
   End
   Begin VB.Frame Frame0 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2565
      Left            =   105
      TabIndex        =   75
      Top             =   9720
      Visible         =   0   'False
      Width           =   15030
      Begin VB.ComboBox CmbSType 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "FrmTxn_DepositContract.frx":0000
         Left            =   1680
         List            =   "FrmTxn_DepositContract.frx":000A
         Sorted          =   -1  'True
         TabIndex        =   175
         Top             =   1200
         Width           =   1550
      End
      Begin VB.ComboBox CmbSContractType 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "FrmTxn_DepositContract.frx":001E
         Left            =   120
         List            =   "FrmTxn_DepositContract.frx":0028
         Sorted          =   -1  'True
         TabIndex        =   174
         Top             =   1200
         Width           =   1550
      End
      Begin VB.ComboBox CmbSVerified 
         Height          =   315
         ItemData        =   "FrmTxn_DepositContract.frx":003E
         Left            =   13320
         List            =   "FrmTxn_DepositContract.frx":0048
         Sorted          =   -1  'True
         TabIndex        =   180
         Top             =   1200
         Width           =   1500
      End
      Begin VB.ComboBox CmbSStatus 
         Height          =   315
         ItemData        =   "FrmTxn_DepositContract.frx":0055
         Left            =   11430
         List            =   "FrmTxn_DepositContract.frx":0057
         Sorted          =   -1  'True
         TabIndex        =   179
         Top             =   1200
         Width           =   1860
      End
      Begin VB.ComboBox CmbSForwardedTo 
         Height          =   315
         Left            =   7740
         Sorted          =   -1  'True
         TabIndex        =   178
         Top             =   1200
         Width           =   3645
      End
      Begin VB.ComboBox CmbSClient 
         Height          =   315
         Left            =   3240
         Sorted          =   -1  'True
         TabIndex        =   177
         Top             =   1200
         Width           =   4485
      End
      Begin VB.CheckBox Check3 
         Appearance      =   0  'Flat
         Caption         =   "Converted Client"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3240
         TabIndex        =   176
         Top             =   930
         Width           =   1500
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "Search"
         Height          =   350
         Left            =   12810
         TabIndex        =   182
         Top             =   1545
         Width           =   1000
      End
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   350
         Left            =   13830
         TabIndex        =   183
         Top             =   1545
         Width           =   1000
      End
      Begin VB.ComboBox CmbSState 
         Height          =   315
         ItemData        =   "FrmTxn_DepositContract.frx":0059
         Left            =   3240
         List            =   "FrmTxn_DepositContract.frx":005B
         Sorted          =   -1  'True
         TabIndex        =   171
         Top             =   570
         Width           =   3240
      End
      Begin VB.ComboBox CmbSCommodity 
         Height          =   315
         Left            =   10500
         Sorted          =   -1  'True
         TabIndex        =   173
         Top             =   555
         Width           =   4365
      End
      Begin VB.ComboBox CmbSLocation 
         Height          =   315
         ItemData        =   "FrmTxn_DepositContract.frx":005D
         Left            =   6480
         List            =   "FrmTxn_DepositContract.frx":005F
         Sorted          =   -1  'True
         TabIndex        =   172
         Top             =   570
         Width           =   4005
      End
      Begin VB.TextBox TxtSDepositContractID 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   310
         Left            =   7635
         MaxLength       =   12
         TabIndex        =   181
         Top             =   1560
         Width           =   1770
      End
      Begin MSComCtl2.DTPicker STo 
         Height          =   315
         Left            =   1695
         TabIndex        =   170
         TabStop         =   0   'False
         Top             =   555
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
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
         Format          =   50921473
         CurrentDate     =   36494
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
         Height          =   495
         Left            =   120
         TabIndex        =   184
         Top             =   1950
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   873
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
      Begin MSComCtl2.DTPicker SFrom 
         Height          =   315
         Left            =   135
         TabIndex        =   169
         TabStop         =   0   'False
         Top             =   555
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
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
         Format          =   50921473
         CurrentDate     =   36494
      End
      Begin VB.Label Label14 
         Caption         =   "Contract Type"
         Height          =   300
         Index           =   6
         Left            =   355
         TabIndex        =   187
         Top             =   990
         Width           =   1080
      End
      Begin VB.Label Label14 
         Caption         =   "Quantity / Bags"
         Height          =   300
         Index           =   5
         Left            =   1855
         TabIndex        =   186
         Top             =   990
         Width           =   1200
      End
      Begin VB.Label Label14 
         Caption         =   "Verified"
         Height          =   300
         Index           =   4
         Left            =   13680
         TabIndex        =   168
         Top             =   960
         Width           =   600
      End
      Begin VB.Label Label44 
         Caption         =   "From Date"
         Height          =   255
         Left            =   435
         TabIndex        =   90
         Top             =   255
         Width           =   855
      End
      Begin VB.Label Label14 
         Caption         =   "Status"
         Height          =   300
         Index           =   3
         Left            =   12075
         TabIndex        =   89
         Top             =   990
         Width           =   600
      End
      Begin VB.Label Label14 
         Caption         =   "Forwarded To"
         Height          =   300
         Index           =   2
         Left            =   9195
         TabIndex        =   88
         Top             =   990
         Width           =   1440
      End
      Begin VB.Label Label14 
         Caption         =   "Client"
         Height          =   300
         Index           =   1
         Left            =   5182
         TabIndex        =   87
         Top             =   990
         Width           =   600
      End
      Begin VB.Label Label45 
         Caption         =   "To Date"
         Height          =   255
         Left            =   1995
         TabIndex        =   82
         Top             =   255
         Width           =   855
      End
      Begin VB.Label Label23 
         Caption         =   "Total Records Found :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   9720
         TabIndex        =   81
         Top             =   1588
         Width           =   1935
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   11655
         TabIndex        =   80
         Top             =   1618
         Width           =   75
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "State"
         Height          =   195
         Left            =   4485
         TabIndex        =   79
         Top             =   255
         Width           =   375
      End
      Begin VB.Label Label27 
         Caption         =   "Commodity"
         Height          =   300
         Left            =   11235
         TabIndex        =   78
         Top             =   240
         Width           =   1680
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Location"
         Height          =   195
         Index           =   1
         Left            =   7680
         TabIndex        =   77
         Top             =   210
         Width           =   615
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Deposit Contract ID"
         Height          =   195
         Left            =   6105
         TabIndex        =   76
         Top             =   1620
         Width           =   1395
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4575
      Left            =   120
      TabIndex        =   32
      Top             =   4080
      Width           =   15030
      _ExtentX        =   26511
      _ExtentY        =   8070
      _Version        =   393216
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Godown Mapping"
      TabPicture(0)   =   "FrmTxn_DepositContract.frx":0061
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Attachments"
      TabPicture(1)   =   "FrmTxn_DepositContract.frx":007D
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "MSHFlexGrid5"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "FrameAttachment"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Log History"
      TabPicture(2)   =   "FrmTxn_DepositContract.frx":0099
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label32"
      Tab(2).Control(1)=   "Label13(2)"
      Tab(2).Control(2)=   "Approved"
      Tab(2).Control(3)=   "Label52"
      Tab(2).Control(4)=   "MSHFlexGrid3"
      Tab(2).Control(5)=   "TxtCreated"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "TxtUpdated"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "TxtApproved"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "TxtClosed"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).ControlCount=   9
      TabCaption(3)   =   "Orignal Data"
      TabPicture(3)   =   "FrmTxn_DepositContract.frx":00B5
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "MSHFlexGrid6"
      Tab(3).ControlCount=   1
      Begin VB.TextBox TxtClosed 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   310
         Left            =   -66960
         Locked          =   -1  'True
         TabIndex        =   158
         TabStop         =   0   'False
         Top             =   720
         Width           =   6885
      End
      Begin VB.TextBox TxtApproved 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   310
         Left            =   -73935
         Locked          =   -1  'True
         TabIndex        =   97
         TabStop         =   0   'False
         Top             =   720
         Width           =   5895
      End
      Begin VB.PictureBox FrameAttachment 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1185
         Left            =   120
         ScaleHeight     =   1155
         ScaleWidth      =   14745
         TabIndex        =   101
         Top             =   400
         Width           =   14775
         Begin VB.CommandButton CmdGetFile 
            Caption         =   "Get File"
            Height          =   310
            Left            =   12840
            Style           =   1  'Graphical
            TabIndex        =   105
            Top             =   50
            Width           =   1845
         End
         Begin VB.CommandButton CmdAddAttachment 
            Caption         =   "Add Attachment"
            Height          =   310
            Left            =   12840
            TabIndex        =   104
            Top             =   380
            Width           =   1845
         End
         Begin VB.CommandButton CmdRemoveAttachment 
            Caption         =   "Remove"
            Height          =   310
            Left            =   12840
            TabIndex        =   103
            Top             =   700
            Width           =   1845
         End
         Begin VB.TextBox TxtAttachmentRemarks 
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
            Height          =   720
            Left            =   1200
            MaxLength       =   1000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   102
            Top             =   360
            Width           =   11535
         End
         Begin MSComDlg.CommonDialog CDialog 
            Left            =   120
            Top             =   240
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Label lblAttachment 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   310
            Left            =   1200
            TabIndex        =   106
            Top             =   45
            Width           =   11535
         End
         Begin VB.Label Label72 
            Caption         =   "File Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   50
            TabIndex        =   108
            Top             =   105
            Width           =   1215
         End
         Begin VB.Label Label73 
            Caption         =   "Remark"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   45
            TabIndex        =   107
            Top             =   600
            Width           =   1215
         End
      End
      Begin VB.TextBox TxtUpdated 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   310
         Left            =   -66960
         Locked          =   -1  'True
         TabIndex        =   96
         TabStop         =   0   'False
         Top             =   360
         Width           =   6885
      End
      Begin VB.TextBox TxtCreated 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   310
         Left            =   -73935
         Locked          =   -1  'True
         TabIndex        =   95
         TabStop         =   0   'False
         Top             =   360
         Width           =   5895
      End
      Begin VB.Frame Frame2 
         Height          =   4170
         Left            =   -74925
         TabIndex        =   93
         Top             =   315
         Width           =   14850
         Begin VB.CommandButton cmdExportDetails 
            Caption         =   "Export Details In Excel"
            Height          =   315
            Left            =   12960
            TabIndex        =   157
            Top             =   120
            Width           =   1815
         End
         Begin VB.CommandButton CmdAddNewBagSize 
            Caption         =   "Add New Bag Size"
            Height          =   315
            Left            =   1440
            TabIndex        =   145
            Top             =   120
            Width           =   1815
         End
         Begin VB.TextBox TxtBagSize 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            Height          =   310
            Left            =   480
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   1440
            Visible         =   0   'False
            Width           =   1860
         End
         Begin VB.TextBox TxtFinalPrice 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Height          =   310
            Left            =   480
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   1080
            Visible         =   0   'False
            Width           =   1860
         End
         Begin VB.CheckBox Check2 
            Appearance      =   0  'Flat
            Caption         =   "Select All"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   90
            TabIndex        =   33
            Top             =   165
            Width           =   1305
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid4 
            Height          =   3600
            Left            =   120
            TabIndex        =   35
            Top             =   450
            Width           =   14670
            _ExtentX        =   25876
            _ExtentY        =   6350
            _Version        =   393216
            WordWrap        =   -1  'True
            AllowBigSelection=   0   'False
            HighLight       =   0
            AllowUserResizing=   1
            Appearance      =   0
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
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   3600
            Left            =   120
            TabIndex        =   34
            Top             =   450
            Width           =   14670
            _ExtentX        =   25876
            _ExtentY        =   6350
            _Version        =   393216
            WordWrap        =   -1  'True
            AllowBigSelection=   0   'False
            HighLight       =   0
            AllowUserResizing=   1
            Appearance      =   0
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
         Height          =   3330
         Left            =   -74925
         TabIndex        =   98
         Top             =   1155
         Width           =   14850
         _ExtentX        =   26194
         _ExtentY        =   5874
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   1
         Appearance      =   0
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid5 
         Height          =   2775
         Left            =   120
         TabIndex        =   109
         Top             =   1680
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   4895
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         Appearance      =   0
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid6 
         Height          =   4080
         Left            =   -74880
         TabIndex        =   161
         Top             =   360
         Width           =   14790
         _ExtentX        =   26088
         _ExtentY        =   7197
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowBigSelection=   0   'False
         HighLight       =   0
         AllowUserResizing=   1
         Appearance      =   0
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
      Begin VB.Label Label52 
         AutoSize        =   -1  'True
         Caption         =   "Closed :"
         Height          =   195
         Left            =   -67920
         TabIndex        =   159
         Top             =   780
         Width           =   570
      End
      Begin VB.Label Approved 
         AutoSize        =   -1  'True
         Caption         =   "Approved :"
         Height          =   195
         Left            =   -74880
         TabIndex        =   112
         Top             =   780
         Width           =   780
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Created  :"
         Height          =   310
         Index           =   2
         Left            =   -74880
         TabIndex        =   100
         Top             =   420
         Width           =   690
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   "Updated :"
         Height          =   310
         Left            =   -67920
         TabIndex        =   99
         Top             =   420
         Width           =   705
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1125
      Left            =   105
      TabIndex        =   72
      Top             =   8650
      Width           =   15030
      Begin VB.CommandButton CmdVerified 
         Caption         =   "Verified"
         Height          =   400
         Left            =   13260
         TabIndex        =   162
         Top             =   195
         Width           =   1560
      End
      Begin VB.CommandButton CmdClosed 
         Caption         =   "Closed"
         Height          =   400
         Left            =   13260
         TabIndex        =   49
         Top             =   600
         Width           =   1560
      End
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   400
         Left            =   9465
         TabIndex        =   47
         Top             =   600
         Width           =   2340
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "Search"
         Height          =   400
         Left            =   9465
         TabIndex        =   42
         Top             =   195
         Width           =   2340
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   400
         Left            =   7125
         TabIndex        =   46
         Top             =   600
         Width           =   2340
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   400
         Left            =   7125
         TabIndex        =   41
         Top             =   195
         Width           =   2340
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   400
         Left            =   4785
         TabIndex        =   45
         Top             =   600
         Width           =   2340
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   400
         Left            =   4785
         TabIndex        =   40
         Top             =   195
         Width           =   2340
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "Previous"
         Height          =   400
         Left            =   2445
         TabIndex        =   44
         Top             =   600
         Width           =   2340
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "Next"
         Height          =   400
         Left            =   105
         TabIndex        =   43
         Top             =   600
         Width           =   2340
      End
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New"
         Height          =   400
         Left            =   105
         TabIndex        =   73
         Top             =   195
         Width           =   2340
      End
      Begin VB.CommandButton CmdCopy 
         Caption         =   "Copy"
         Height          =   810
         Left            =   11805
         TabIndex        =   48
         Top             =   195
         Width           =   1440
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   400
         Left            =   2445
         TabIndex        =   39
         Top             =   195
         Width           =   2340
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4005
      Left            =   120
      TabIndex        =   38
      Top             =   0
      Width           =   15030
      Begin VB.ComboBox CmbContractType 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "FrmTxn_DepositContract.frx":00D1
         Left            =   4305
         List            =   "FrmTxn_DepositContract.frx":00DB
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   195
         Width           =   1920
      End
      Begin VB.CommandButton CmdRefreshRate 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Refresh Rate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   160
         Top             =   1800
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox TxtQuantum 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   14022
         MaxLength       =   10
         TabIndex        =   17
         Top             =   2235
         Width           =   800
      End
      Begin VB.ComboBox CmbRatePer 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "FrmTxn_DepositContract.frx":00F1
         Left            =   12480
         List            =   "FrmTxn_DepositContract.frx":00FB
         Sorted          =   -1  'True
         TabIndex        =   21
         Top             =   2640
         Width           =   2460
      End
      Begin VB.CheckBox Check4 
         Appearance      =   0  'Flat
         Caption         =   "Calculation On Bags"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1635
         TabIndex        =   12
         Top             =   910
         Width           =   1770
      End
      Begin VB.TextBox TxtLockInPeriod 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   12500
         MaxLength       =   10
         TabIndex        =   16
         Top             =   2235
         Width           =   800
      End
      Begin VB.TextBox TxtRemarks 
         Appearance      =   0  'Flat
         Height          =   580
         Left            =   1560
         MaxLength       =   4000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   24
         Top             =   3350
         Width           =   13365
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Converted Client"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   6945
         TabIndex        =   9
         Top             =   525
         Width           =   1110
      End
      Begin VB.TextBox TxtDepositContractID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   310
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   195
         Width           =   1290
      End
      Begin VB.TextBox TxtTentativeDuration 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   10460
         MaxLength       =   10
         TabIndex        =   15
         Top             =   2235
         Width           =   800
      End
      Begin MSComCtl2.DTPicker TxtStartDate 
         Height          =   315
         Left            =   8085
         TabIndex        =   13
         Top             =   1755
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   556
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
         Format          =   50921473
         CurrentDate     =   39850
      End
      Begin VB.TextBox TxtEmailId 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   310
         Left            =   8085
         Locked          =   -1  'True
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   1300
         Width           =   6870
      End
      Begin VB.TextBox TxtContactNo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   310
         Left            =   8085
         Locked          =   -1  'True
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   910
         Width           =   2970
      End
      Begin VB.TextBox TxtContactDetails 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   310
         Left            =   12480
         Locked          =   -1  'True
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   910
         Width           =   2460
      End
      Begin VB.TextBox TxtAddress 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   1000
         Left            =   1560
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   1245
         Width           =   4650
      End
      Begin MSComCtl2.DTPicker TxtEndDate 
         Height          =   315
         Left            =   8085
         TabIndex        =   14
         Top             =   2205
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   556
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
         Format          =   50921473
         CurrentDate     =   39850
      End
      Begin MSComCtl2.DTPicker TxtDepositContractDate 
         Height          =   315
         Left            =   1560
         TabIndex        =   25
         Top             =   555
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
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
         Format          =   50921473
         CurrentDate     =   39850
      End
      Begin VB.ComboBox CmbStatus 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "FrmTxn_DepositContract.frx":0113
         Left            =   1560
         List            =   "FrmTxn_DepositContract.frx":0115
         Sorted          =   -1  'True
         TabIndex        =   22
         Top             =   3000
         Width           =   4650
      End
      Begin VB.TextBox TxtStatus 
         Appearance      =   0  'Flat
         Height          =   310
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   3000
         Width           =   4650
      End
      Begin VB.ComboBox CmbInsurance 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1560
         Sorted          =   -1  'True
         TabIndex        =   18
         Top             =   2290
         Width           =   4650
      End
      Begin VB.TextBox TxtInsurance 
         Appearance      =   0  'Flat
         Height          =   310
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   2292
         Width           =   4650
      End
      Begin VB.ComboBox CmbBillingCycle 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "FrmTxn_DepositContract.frx":0117
         Left            =   1560
         List            =   "FrmTxn_DepositContract.frx":0119
         TabIndex        =   19
         Top             =   2640
         Width           =   4650
      End
      Begin VB.TextBox TxtBillingCycle 
         Appearance      =   0  'Flat
         Height          =   310
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   2642
         Width           =   4650
      End
      Begin VB.TextBox TxtRatePer 
         Appearance      =   0  'Flat
         Height          =   310
         Left            =   12480
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   2640
         Width           =   2460
      End
      Begin VB.ComboBox CmbClient 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   8085
         Sorted          =   -1  'True
         TabIndex        =   10
         Top             =   555
         Width           =   6855
      End
      Begin VB.TextBox TxtClient 
         Appearance      =   0  'Flat
         Height          =   310
         Left            =   8085
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   555
         Width           =   6855
      End
      Begin VB.ComboBox CmbState 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "FrmTxn_DepositContract.frx":011B
         Left            =   8085
         List            =   "FrmTxn_DepositContract.frx":011D
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   195
         Width           =   2910
      End
      Begin VB.ComboBox CmbCommodity 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4305
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   555
         Width           =   1920
      End
      Begin VB.TextBox TxtState 
         Appearance      =   0  'Flat
         Height          =   310
         Left            =   8085
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   200
         Width           =   2895
      End
      Begin VB.TextBox TxtCommodity 
         Appearance      =   0  'Flat
         Height          =   310
         Left            =   4305
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   557
         Width           =   1920
      End
      Begin VB.ComboBox CmbDepositBillingCycle 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "FrmTxn_DepositContract.frx":011F
         Left            =   8085
         List            =   "FrmTxn_DepositContract.frx":0121
         TabIndex        =   20
         Top             =   2640
         Width           =   2970
      End
      Begin VB.ComboBox CmbLocation 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   11925
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   195
         Width           =   3000
      End
      Begin VB.TextBox TxtLocation 
         Appearance      =   0  'Flat
         Height          =   310
         Left            =   11925
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   195
         Width           =   3000
      End
      Begin VB.TextBox TxtDepositBillingCycle 
         Appearance      =   0  'Flat
         Height          =   310
         Left            =   8085
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   2640
         Width           =   2970
      End
      Begin VB.ComboBox CmbForwardedTo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   8085
         Sorted          =   -1  'True
         TabIndex        =   23
         Top             =   3000
         Width           =   6885
      End
      Begin VB.TextBox TxtForwardedTo 
         Appearance      =   0  'Flat
         Height          =   310
         Left            =   8085
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   3000
         Width           =   6855
      End
      Begin VB.TextBox TxtContractType 
         Appearance      =   0  'Flat
         Height          =   310
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   195
         Width           =   1890
      End
      Begin VB.Label Label54 
         Caption         =   "Contract Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2950
         TabIndex        =   185
         Top             =   225
         Width           =   1335
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "Indicative Quantum"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   13995
         TabIndex        =   94
         Top             =   1750
         Width           =   855
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label7 
         Caption         =   "Contact Details"
         Height          =   255
         Left            =   11115
         TabIndex        =   92
         Top             =   938
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "Email ID "
         Height          =   255
         Left            =   6345
         TabIndex        =   91
         Top             =   1328
         Width           =   1440
      End
      Begin VB.Label Label35 
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   105
         TabIndex        =   86
         Top             =   3030
         Width           =   720
      End
      Begin VB.Label Label34 
         Caption         =   "Forwarded To"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6345
         TabIndex        =   85
         Top             =   3030
         Width           =   1440
      End
      Begin VB.Label Label33 
         Caption         =   "Remarks"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   105
         TabIndex        =   84
         Top             =   3513
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Client"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6345
         TabIndex        =   83
         Top             =   585
         Width           =   1440
      End
      Begin VB.Label Label28 
         Caption         =   "Commodity"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   2950
         TabIndex        =   74
         Top             =   585
         Width           =   1095
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         Caption         =   "Lock - in Period (in months)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   12120
         TabIndex        =   71
         Top             =   1750
         Width           =   1560
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label20 
         Caption         =   "Tentative Dur. of Storage (in months)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9960
         TabIndex        =   70
         Top             =   1750
         Width           =   1800
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label9 
         Caption         =   "Insurance By"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   105
         TabIndex        =   63
         Top             =   2265
         Width           =   1440
      End
      Begin VB.Label Label18 
         Caption         =   "GST"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   11235
         TabIndex        =   69
         Top             =   2670
         Width           =   375
      End
      Begin VB.Label Label16 
         Caption         =   "Deposit/ Withdrawal Billing Cycle"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6345
         TabIndex        =   68
         Top             =   2550
         Width           =   1800
      End
      Begin VB.Label Label15 
         Caption         =   "Contact No"
         Height          =   255
         Left            =   6345
         TabIndex        =   67
         Top             =   938
         Width           =   1440
      End
      Begin VB.Label Label12 
         Caption         =   "Billing Cycle"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   105
         TabIndex        =   66
         Top             =   2670
         Width           =   1335
      End
      Begin VB.Label Label14 
         Caption         =   "Deposit Contract End Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   0
         Left            =   6345
         TabIndex        =   65
         Top             =   2137
         Width           =   1560
      End
      Begin VB.Label Label13 
         Caption         =   "Deposit Contract Start Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   6345
         TabIndex        =   64
         Top             =   1725
         Width           =   1560
      End
      Begin VB.Label Label8 
         Caption         =   "Pin Code"
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
         Left            =   7575
         TabIndex        =   60
         Top             =   5340
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Address"
         Height          =   255
         Left            =   135
         TabIndex        =   58
         Top             =   1365
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Location"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   11025
         TabIndex        =   56
         Top             =   225
         Width           =   1440
      End
      Begin VB.Label Label3 
         Caption         =   "State"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6390
         TabIndex        =   55
         Top             =   255
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Date"
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   615
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Deposit Contract ID"
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   225
         Width           =   1455
      End
   End
End
Attribute VB_Name = "FrmTxn_DepositContract"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg As Variant
Const strChecked = "þ"
Const strUnChecked = "q"
Dim oFs As New FileSystemObject
Dim oFs1 As New FileSystemObject
Dim oFolder As Folder
Dim oFile As File
Dim mBagSizeFlag As Variant
Dim mStateID As Variant, mLocationID As Variant, mClientIDRow As Variant, mSM_Prefix As Variant, mCMPID As Variant, mGodownID As Variant, mCommodityID As Variant
Dim File_Path As Variant, mFileFlag As Variant, mSelCol As Variant, mSelRow As Variant, mBillingCycleID As Variant, mBillingRatio As Variant, mBillingCycleNoofDays As Variant, mSRCID As Variant
Dim mDepositBillingCycleID As Variant, mDepositBillingRatio As Variant, mDepositBillingCycleNoofDays As Variant
Dim mSStateID As Variant, mSLocationID As Variant, mSMPrefix As Variant, mSCommodityID As Variant, mSClientIDRow As Variant, mRateOnQty As Variant
Dim mRatio As Variant, mDepositRatio As Variant, mForwardID As Variant, mStatus As Variant, mSForwardID As Variant, mSStatus As Variant, mInsuranceByID As Variant, mCreatedID As Variant
Dim mCurrentConversionFactor As Variant, mBagEquivalent As Variant, mDivisionFactor As Variant, mFumigationFactorY As Variant, mFumigationFactorN As Variant, mFumigationFlag As Variant
Dim mCMPMtoAMRMApprovalLimit As Variant, mKnownRentMinimumValue As Variant, mDepositContractCMGGroupToBeApproved As Variant
Dim mTmpDepConID As Variant, mValidGodownID As Variant, mSelectGodownID As Variant, tmp1 As Variant, mMinBagSize As Variant, mMaxBagSize As Variant, mSelectStatus As Variant, mFinalStatus As Variant
Dim mUpperHierarchyEmployeeNoList As String, mUpperHierarchyEmployeeIDList As String
Dim mCopyFlag As Boolean, mFlag As Boolean

Dim mBillingOnePercentage, mDepositOnePercentage As Variant
Dim mGodownList As Variant, mNewCMPID As Variant, mNewGodownID As Variant
Dim mExportExcel As Boolean, mConfirmFlag As Boolean
Dim mDepositContractDefaultDaysRange As Integer
Dim mDepositContractInFutureDateFlag As Variant
Dim mCStateName As Variant, mCStateNameID As Variant
Dim mCLocationName As Variant, mCLocationID As Variant
Dim mCCommodityName As Variant, mCCommodityID As Variant
Dim mCClientName As Variant, mCClientID As Variant
Dim mCInsuranceBy As Variant, mCInsuranceID As Variant
Dim mCBags As Boolean
Dim mCTentativeDuration As Variant, mCLockInPeriod As Variant, mCQuantum As Variant
Dim mCBillingCycle As Variant, mCBillingCycleID As Variant
Dim mCDBillingCycle As Variant, mCDBillingCycleID As Variant
Dim mApprovedRights As Boolean
''-- 06 Jan 2018
Dim mDepositContractDefaultAttachmentFlag As Variant
Dim mDepositContractAllowBackDateEntryDays As Integer
''-- 11 Jan 2018
Dim mDepositContractAdditionalGroupToBeApproved As Variant, mDepositContractAdditionalGroupToBeApprovedList As Variant

''-- 1 Feb 2018
Dim mOldBillingCycle As Variant, mOldDBillingCycle As Variant
''-- 02 Feb 2018
Dim mVerifiyFlag As Boolean

''-- 14 Feb 2018
Dim mContractTypeFlag As Variant, mSContractTypeFlag As Variant, mSType As Variant ''-- 24 Jan 2018
Dim mCContractTypeFlag As Variant, mCContractTypeText As Variant

Dim mDepositContractDaysGapStartDateEndDate As Integer

Dim mCurrentBaseRentFactor As Variant
Dim mCalculatedBaseRent As Variant


Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)
Call ClearControl
Call Grid_Head6
CmbStatus.Text = "Forwarded"
CmbBillingCycle.Text = "Monthly"
CmbDepositBillingCycle.Text = "Monthly"
TxtTentativeDuration.Text = "12"
''--- 17 Feb 2018
TxtLockInPeriod.Text = "1"
TxtQuantum.Text = "1"
''-----------
''-- 14 Feb 2018 Ullhas Patil
If LCase(Gen_UserRole) = "power" Or mApprovedRights = True Then
    TxtContractType = "": TxtContractType.Visible = False
    CmbContractType = "Normal": CmbContractType.Visible = True
    mContractTypeFlag = "N"
Else
    TxtContractType = "Normal": TxtContractType.Visible = True
    CmbContractType = "Normal": CmbContractType.Visible = False
    mContractTypeFlag = "N"
End If
''----------------------------

Call GetBillingCycleID(CmbBillingCycle.Text)
Call GetDepositeBillingCycleID(CmbDepositBillingCycle.Text)
cmdExportDetails.Enabled = True
CmdRefreshRate.Visible = False
End Sub


Private Sub Check2_Click()
TxtFinalPrice.Visible = False
TxtBagSize.Visible = False

If Check4.Value = 0 Then
   Call CheckBoxChecked(MSHFlexGrid1)
   ''Call CmbBillingCycle_LostFocus        '''-- for Color Effect
   ''Call CmbDepositBillingCycle_LostFocus '''-- for Color Effect
Else
   Call CheckBoxChecked(MSHFlexGrid4)
   ''Call CmbBillingCycle_LostFocus        '''-- for Color Effect
   ''Call CmbDepositBillingCycle_LostFocus '''-- for Color Effect
End If

''--- Billing Cycle
If CmbBillingCycle.Text <> "" Then
    Call TableMst_BillingCycle(" Where Flag <> 'R' Order By BillingCycleID ")
    RsMst_BillingCycle.MoveFirst
    RsMst_BillingCycle.Find "BillingCycle = '" & CmbBillingCycle.Text & "'"
    If RsMst_BillingCycle.EOF = True Then
       MsgBox "Invalid selection "
       CmbBillingCycle.SetFocus
       Exit Sub
    End If
    
    mBillingCycleID = RsMst_BillingCycle!BillingCycleID
    mBillingRatio = IIf(IsNull(RsMst_BillingCycle!Ratio), 0, RsMst_BillingCycle!Ratio)
    mBillingCycleNoofDays = IIf(IsNull(RsMst_BillingCycle!NoOfDays), 0, RsMst_BillingCycle!NoOfDays)

    If Check4.Value = 0 Then 'Calculation without bags
        With MSHFlexGrid1
            For I = 1 To .Rows - 1
               If .TextMatrix(I, 0) = strChecked Then
                ''-- Lease / Own Godown
                If Val(.TextMatrix(I, 15)) = 1 Or Val(.TextMatrix(I, 15)) = 7 Then
                    .TextMatrix(I, 7) = Format(CalculateRate(.TextMatrix(I, 13), mBillingRatio, CmbBillingCycle.Text), "#0.0000")
                    mBillingOnePercentage = Format(IIf(Val(.TextMatrix(I, 7)) > 0, Val(.TextMatrix(I, 7)) / 100, 0), "#0.0000")
                    .TextMatrix(I, 8) = Format(Val(.TextMatrix(I, 7)) + (mBillingOnePercentage * .TextMatrix(I, 17)), "#0.0000")
                Else
                    .TextMatrix(I, 8) = .TextMatrix(I, 8)
                End If
                Call CheckRateByColor(CInt(I))
               End If
            Next
        End With
    Else 'Calculation on bags
        With MSHFlexGrid4
            For I = 1 To .Rows - 1
               If .TextMatrix(I, 0) = strChecked Then
                  ''-- Lease / Own Godown
                 If Val(.TextMatrix(I, 17)) = 1 Or Val(.TextMatrix(I, 17)) = 7 Then     ''-- If Godown Type is not Lease (ie Rent not available then on Billing Cycle Only Proposed Rate get Calculate)
                     .TextMatrix(I, 9) = Format(CalculateRate(.TextMatrix(I, 15), mBillingRatio, CmbBillingCycle) * Val(.TextMatrix(I, 8)) / 1000, "#0.0000")
                     mBillingOnePercentage = Format(Val(.TextMatrix(I, 9)) / 100, "#0.0000")
                     .TextMatrix(I, 10) = Format(Val(.TextMatrix(I, 9)) + (mBillingOnePercentage * .TextMatrix(I, 19)), "#0.0000")
                  Else
                     .TextMatrix(I, 10) = .TextMatrix(I, 10)
                  End If
                  Call CheckBagRateByColor(CInt(I))
               End If
            Next
        End With
    End If
End If
''--- Deposit Billing Cycle

If CmbDepositBillingCycle.Text <> "" Then
    Call TableMst_BillingCycle(" Where Flag <> 'R' Order By BillingCycleID ")
    RsMst_BillingCycle.MoveFirst
    RsMst_BillingCycle.Find "BillingCycle = '" & CmbDepositBillingCycle.Text & "'"
    If RsMst_BillingCycle.EOF = True Then
       MsgBox "Invalid selection "
       CmbDepositBillingCycle.SetFocus
       Exit Sub
    End If
    
    mDepositBillingCycleID = RsMst_BillingCycle!BillingCycleID
    mDepositBillingRatio = IIf(IsNull(RsMst_BillingCycle!Ratio), 0, RsMst_BillingCycle!Ratio)
    mDepositBillingCycleNoofDays = IIf(IsNull(RsMst_BillingCycle!NoOfDays), 0, RsMst_BillingCycle!NoOfDays)

    If Check4.Value = 0 Then 'Calculation without bags
        With MSHFlexGrid1
            For I = 1 To .Rows - 1
                If .TextMatrix(I, 0) = strChecked Then
                    If Val(.TextMatrix(I, 15)) = 1 Or Val(.TextMatrix(I, 15)) = 7 Then
                        .TextMatrix(I, 11) = Format(CalculateRate(.TextMatrix(I, 14), mDepositBillingRatio, CmbDepositBillingCycle.Text), "#0.0000")
                        mDepositOnePercentage = Format(IIf(Val(.TextMatrix(I, 11)) > 0, Val(.TextMatrix(I, 11)) / 100, 0), "#0.0000")
                        .TextMatrix(I, 12) = Format(Val(.TextMatrix(I, 11)) + (mDepositOnePercentage * .TextMatrix(I, 17)), "#0.0000")
                    Else
                       .TextMatrix(I, 12) = .TextMatrix(I, 12)
                    End If
                    Call CheckRateByColor(CInt(I))
                End If
            Next
        End With
    Else
        With MSHFlexGrid4
            For I = 1 To .Rows - 1
                If .TextMatrix(I, 0) = strChecked Then
                    If Val(.TextMatrix(I, 17)) = 1 Or Val(.TextMatrix(I, 17)) = 7 Then     ''-- If Godown Type is not Lease (ie Rent not available then on Billing Cycle Only Proposed Rate get Calculate)
                        .TextMatrix(I, 13) = Format(CalculateRate(.TextMatrix(I, 16), mDepositBillingRatio, CmbDepositBillingCycle) * Val(.TextMatrix(I, 8)) / 1000, "#0.0000")
                        mDepositOnePercentage = Format(Val(.TextMatrix(I, 13)) / 100, "#0.0000")
                        .TextMatrix(I, 14) = Format(Val(.TextMatrix(I, 13)) + (mDepositOnePercentage * .TextMatrix(I, 19)), "#0.0000")
                    Else
                        .TextMatrix(I, 14) = .TextMatrix(I, 14)
                    End If
                    Call CheckBagRateByColor(CInt(I))
                End If
            Next
        End With
    End If
End If

End Sub


Private Sub Check4_Click()
If Check4.Tag <> "" Then Exit Sub
If Check4.Value = 0 Then
   MSHFlexGrid1.Visible = True
   MSHFlexGrid4.Visible = False
   CmdAddNewBagSize.Visible = False
Else
   MSHFlexGrid1.Visible = False
   MSHFlexGrid4.Visible = True
   CmdAddNewBagSize.Visible = True
   CmdAddNewBagSize.Enabled = True
End If
Call GetGodownDetails
End Sub

Private Sub CmbBillingCycle_GotFocus()
mOldBillingCycle = CmbBillingCycle.Text
'tmp1 = CmbBillingCycle.Text
tmp = CmbBillingCycle.Text
Call TableMst_BillingCycle(" Where Flag <> 'R' Order By BillingCycleID ")
CmbBillingCycle.Clear
If RsMst_BillingCycle.RecordCount = 0 Then
  MsgBox "No Record found !!!! "
  Exit Sub
End If
For I = 1 To RsMst_BillingCycle.RecordCount
   CmbBillingCycle.AddItem RsMst_BillingCycle!BillingCycle
   RsMst_BillingCycle.MoveNext
Next
CmbBillingCycle.Text = tmp
End Sub

Private Sub CmbBillingCycle_LostFocus()
If CmbBillingCycle.Text = "" Then mBillingCycleID = Null:  Exit Sub
RsMst_BillingCycle.MoveFirst
RsMst_BillingCycle.Find "BillingCycle = '" & CmbBillingCycle.Text & "'"
If RsMst_BillingCycle.EOF = True Then
   MsgBox "Invalid selection "
   CmbBillingCycle.SetFocus
   Exit Sub
End If

mBillingCycleID = RsMst_BillingCycle!BillingCycleID
mBillingRatio = IIf(IsNull(RsMst_BillingCycle!Ratio), 0, RsMst_BillingCycle!Ratio)
mBillingCycleNoofDays = IIf(IsNull(RsMst_BillingCycle!NoOfDays), 0, RsMst_BillingCycle!NoOfDays)

If LCase(mOldBillingCycle) = LCase(CmbBillingCycle.Text) Then
    Exit Sub
End If
'If mOldBillingCycle <> "" Then
    If MsgBox("Below Rates will change as the Billing cycle is changed." & vbNewLine & "Do you want to continue?", vbYesNo) = vbNo Then
        CmbBillingCycle.Text = mOldBillingCycle
        CmbBillingCycle.SetFocus
        Exit Sub
    End If
'End If

If Check4.Value = 0 Then 'Calculation without bags
    With MSHFlexGrid1
        For I = 1 To .Rows - 1
           If .TextMatrix(I, 0) = strChecked Then
            ''-- Lease / Own Godown
            If Val(.TextMatrix(I, 15)) = 1 Or Val(.TextMatrix(I, 15)) = 7 Then
                .TextMatrix(I, 7) = Format(CalculateRate(.TextMatrix(I, 13), mBillingRatio, CmbBillingCycle.Text), "#0.0000")
                mBillingOnePercentage = Format(IIf(Val(.TextMatrix(I, 7)) > 0, Val(.TextMatrix(I, 7)) / 100, 0), "#0.0000")
                .TextMatrix(I, 8) = Format(Val(.TextMatrix(I, 7)) + (mBillingOnePercentage * .TextMatrix(I, 17)), "#0.0000")
            Else
                .TextMatrix(I, 8) = .TextMatrix(I, 8)
            End If
            Call CheckRateByColor(CInt(I))
           End If
        Next
    End With
Else 'Calculation on bags
    With MSHFlexGrid4
        For I = 1 To .Rows - 1
           If .TextMatrix(I, 0) = strChecked Then
              ''-- Lease / Own Godown
             If Val(.TextMatrix(I, 17)) = 1 Or Val(.TextMatrix(I, 17)) = 7 Then     ''-- If Godown Type is not Lease (ie Rent not available then on Billing Cycle Only Proposed Rate get Calculate)
                 .TextMatrix(I, 9) = Format(CalculateRate(.TextMatrix(I, 15), mBillingRatio, CmbBillingCycle) * Val(.TextMatrix(I, 8)) / 1000, "#0.0000")
                 mBillingOnePercentage = Format(Val(.TextMatrix(I, 9)) / 100, "#0.0000")
                 .TextMatrix(I, 10) = Format(Val(.TextMatrix(I, 9)) + (mBillingOnePercentage * .TextMatrix(I, 19)), "#0.0000")
              Else
                 .TextMatrix(I, 10) = .TextMatrix(I, 10)
              End If
              Call CheckBagRateByColor(CInt(I))
           End If
        Next
    End With
End If
End Sub

Private Sub CmbClient_GotFocus()
If CmbContractType.Text = "" Then
   MsgBox "Select Contract Type!!!"
   Call Grid_Head
   Call Grid_Head4
   CmbContractType.SetFocus
   Exit Sub
End If
If CmbState.Text = "" Then
   MsgBox "Select State!!!"
   Call Grid_Head
   Call Grid_Head4
   CmbState.SetFocus
   Exit Sub
End If
If CmbLocation.Text = "" Then
   MsgBox "Select Location!!!"
   Call Grid_Head
   Call Grid_Head4
   CmbLocation.SetFocus
   Exit Sub
End If
If CmbCommodity.Text = "" Then
    MsgBox "Select Commodity."
    Call Grid_Head
    Call Grid_Head4
    CmbCommodity.SetFocus
    Exit Sub
End If

CmbClient.Clear
tmp1 = CmbClient.Text
If Check1.Value = vbUnchecked Then
   ''Call TableMst_Client("Where Converted = 0 and status='A' And ExchangeTypeID = " & GetExchangeID("Non NCDEX") & " Order By ClientName ")
   If LCase(CmbContractType.Text) = "revision" Then
        Call TableMst_Client("Where Converted = 0 and Status='A' And ClientID In ( Select Distinct GSL.ClientID From Mst_GSL GSL Where GSL.Flag Not In ('B') And GSL.ExchangeTypeID = 1 And GSL.SM_Prefix = '" & mSM_Prefix & "' And GSL.CommodityID = " & mCommodityID & " ) ")
   ElseIf LCase(CmbContractType.Text) = "normal" Then
        Call TableMst_Client("Where Converted = 0 and Status='A' And ExchangeTypeID = " & GetExchangeID("Non NCDEX") & " ")
   End If
Else
   ''Call TableMst_Client("Where Converted = 1 and status='A' And ExchangeTypeID = " & GetExchangeID("Non NCDEX") & " Order By ClientName ")
   If LCase(CmbContractType.Text) = "revision" Then
        Call TableMst_Client("Where Converted = 1 and status='A' And ClientID In ( Select Distinct GSL.ClientID From Mst_GSL GSL Where GSL.Flag Not In ('B') And GSL.ExchangeTypeID = 1 And GSL.SM_Prefix = '" & mSM_Prefix & "' And GSL.CommodityID = " & mCommodityID & " ) ")
   ElseIf LCase(CmbContractType.Text) = "normal" Then
        Call TableMst_Client("Where Converted = 1 and status='A' And ExchangeTypeID = " & GetExchangeID("Non NCDEX") & " ")
   End If
End If
CmbClient.Clear
If RsMst_Client.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Client.RecordCount
    CmbClient.AddItem UCase(RsMst_Client!ClientName) & "#" & RsMst_Client!ClientID & "#" & RsMst_Client!ClientID
    RsMst_Client.MoveNext
Next
CmbClient.Text = tmp1
End Sub

Private Sub CmbClient_LostFocus()
Dim mR, mSplit As Variant

If CmbClient.Text = "" Then
   TxtAddress.Text = ""
   TxtContactNo.Text = ""
   TxtContactDetails.Text = ""
   TxtEmailId.Text = ""
   Call Grid_Head
   Call Grid_Head4
   Exit Sub
End If
mSplit = Split(CmbClient.Text, "#")
If UBound(mSplit) <> 2 Then
   MsgBox "Invalid selection "
   CmbClient.SetFocus
   Exit Sub
End If
RsMst_Client.MoveFirst
RsMst_Client.Find "ClientID = " & mSplit(2)
If RsMst_Client.EOF = True Then
   MsgBox "Invalid selection "
   CmbClient.SetFocus
   Exit Sub
Else
   If LCase(RsMst_Client!ClientName) <> LCase(mSplit(0)) Then
    MsgBox "Invalid selection "
    CmbClient.SetFocus
    Exit Sub
   End If
End If
mClientIDRow = RsMst_Client!ClientID
Call GetClientDetails(mClientIDRow)

''-- 14 Feb 2018
''-- if Contract Type is revision then start date is min CDTF date of same Location + Commodity + Client
If LCase(CmbContractType.Text) = "revision" Then
    If CmbLocation.Text <> "" And CmbCommodity.Text <> "" And CmbClient.Text <> "" Then
        tmp = " Select MIN(CDTFDate)[CDTFDate] from Mst_GSL Where SM_Prefix = '" & mSM_Prefix & "' And ClientID = " & mClientIDRow & " And CommodityID = " & mCommodityID & " And Flag = 'O' "
        Call TmpTable
        If TmpRs.RecordCount > 0 Then
            TxtStartDate.Value = IIf(IsNull(TmpRs!CDTFDate), Date - 1, TmpRs!CDTFDate)
            TxtEndDate = CDate(Format(Date, "YYYY-MM-01")) - 1    ''-- Last Month Last Date for revision Contract
        End If
    Else
        TxtStartDate.Value = Date - 1
        TxtEndDate = Date
    End If
Else
    TxtStartDate.Value = Date
    TxtEndDate = Date + mDepositContractDefaultDaysRange
End If
''---------------------------------
Call GetGodownDetails
End Sub

Private Sub CmbCloseNewBagSizeFrame_Click()
    Frame5.Visible = False
    Frame4.Visible = False
    Frame1.Enabled = True
    SSTab1.Enabled = True
    ButtonFrm.Enabled = True
    SSTab1.SetFocus
End Sub

Private Sub CmbCMP_GotFocus()
If Trim(CmbLocation.Text) = "" Then
   MsgBox "Select Location!!!"
   CmbLocation.SetFocus
   Exit Sub
End If
tmp = CmbCMP.Text
If Gen_WcCMP = "" Then
   Call TableMst_CMP(" where CloseDate is null and LocationID = " & mLocationID)
Else
   Call TableMst_CMP(" where " & Gen_WcCMP & " And CloseDate is null and LocationID = " & mLocationID)
End If
CmbCMP.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "No Record found!!!"
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
If RsMst_CMP.EOF = True Then
   MsgBox "Invalid selection!!!"
   CmbCMP.SetFocus
   Exit Sub
End If

If tmp <> CmbCMP.Text Then
   CmbGodown.Text = ""
End If

mCMPID = RsMst_CMP!CMPID
End Sub

Private Sub CmbCloseDetails_Click()
    Frame4.Visible = False
    Frame1.Enabled = True
    SSTab1.Enabled = True
    ButtonFrm.Enabled = True
    SSTab1.SetFocus
End Sub


Private Sub CmbCMPID_GotFocus()
If mLocationID = "" Then
   MsgBox "Please Select Location !!!"
   If CmbLocationID.Enabled = True And CmbLocationID.Visible = True Then CmbLocationID.SetFocus
   Exit Sub
End If
CmbGodownID.Clear
tmp = CmbCMPID.Text

Call TableMst_CMP(" Where LocationID = " & mLocationID & " And CMPID In ( SELECT CMPID FROM Mst_Godown Where GodownID In (" & mGodownList & ") ) ")
CmbCMPID.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   CmbCMPID.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_CMP.RecordCount
    CmbCMPID.AddItem RsMst_CMP!CMPName
    RsMst_CMP.MoveNext
Next
CmbCMPID.Text = tmp
End Sub

Private Sub CmbCMPID_LostFocus()
If CmbCMPID.Text = "" Then
   Exit Sub
End If
RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPName = '" & CmbCMPID.Text & "'"
If RsMst_CMP.EOF = True Then
   MsgBox "Invalid selection "
   CmbCMPID.SetFocus
   Exit Sub
End If
mNewCMPID = RsMst_CMP!CMPID
End Sub

Private Sub CmbCommodity_GotFocus()
If Trim(CmbState.Text) = "" Then
   MsgBox "Select State!!!"
   CmbState.SetFocus
   Exit Sub
End If
If Trim(CmbLocation.Text) = "" Then
   MsgBox "Select Location!!!"
   CmbLocation.SetFocus
   Exit Sub
End If
tmp1 = CmbCommodity.Text
Call TableMst_Commodity(" Where CommodityID IN (SELECT Distinct CommodityID FROM Mst_LocationCommodityMapping Where LocationID = " & mLocationID & " ) ")
CmbCommodity.Clear
If RsMst_Commodity.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Commodity.RecordCount
 CmbCommodity.AddItem RsMst_Commodity!Commodity
 RsMst_Commodity.MoveNext
Next
CmbCommodity.Text = tmp1
End Sub

Private Sub CmbCommodity_LostFocus()
If CmbCommodity.Text = "" Then
    If LCase(CmbContractType.Text) = "revision" Then
     CmbClient.Text = ""
    End If
   Exit Sub
End If
RsMst_Commodity.MoveFirst
RsMst_Commodity.Find "Commodity = '" & CmbCommodity.Text & "'"
If RsMst_Commodity.EOF = True Then
   MsgBox "Invalid selection "
   CmbCommodity.SetFocus
   Exit Sub
End If
If RsMst_Commodity!Commodity <> CmbCommodity Then
   Call ClearData
   
If LCase(CmbContractType.Text) = "revision" Then
 CmbClient.Text = ""
End If
   
End If
mCommodityID = RsMst_Commodity!CommodityID
mFumigationFlag = IIf(IsNull(RsMst_Commodity!FumigationFlag), "N", RsMst_Commodity!FumigationFlag)
mMinBagSize = RsMst_Commodity!Minimum
mMaxBagSize = RsMst_Commodity!Maximum
mBagEquivalent = GetCommodityBagEquivalent(mCommodityID)

''-- 14 Feb 2018
''-- if Contract Type is revision then start date is min CDTF date of same Location + Commodity + Client
If LCase(CmbContractType.Text) = "revision" Then
    If CmbLocation.Text <> "" And CmbCommodity.Text <> "" And CmbClient.Text <> "" Then
        tmp = " Select MIN(CDTFDate)[CDTFDate] from Mst_GSL Where SM_Prefix = '" & mSM_Prefix & "' And ClientID = " & mClientIDRow & " And CommodityID = " & mCommodityID & " And Flag = 'O' "
        Call TmpTable
        If TmpRs.RecordCount > 0 Then
            TxtStartDate.Value = IIf(IsNull(TmpRs!CDTFDate), Date, TmpRs!CDTFDate) - 1
            TxtEndDate = CDate(Format(Date, "YYYY-MM-01")) - 1    ''-- Last Month Last Date for revision Contract
        End If
    Else
        TxtStartDate.Value = Date - 1
        TxtEndDate = Date
    End If
Else
    TxtStartDate.Value = Date
    TxtEndDate = Date + mDepositContractDefaultDaysRange
End If
''---------------------------------

If Val(mBagEquivalent) = 0 Then
    MsgBox "Bag Equivalent is not available, Please contact CMG Team for Commodity Bag Equivalent."
    Exit Sub
End If

If tmp1 <> CmbCommodity Then
    Call GetGodownDetail
End If
End Sub

''-- Ullhas Patil 14 Feb 2018
Private Sub CmbContractType_GotFocus()
    tmp1 = CmbContractType.Text
End Sub
Private Sub CmbContractType_LostFocus()
If CmbContractType.Text = "" Then Exit Sub
If LCase(CmbContractType.Text) <> "normal" And LCase(CmbContractType.Text) <> "revision" Then
   MsgBox "Invalid Selection !!!"
   CmbContractType.SetFocus
   Exit Sub
End If
If LCase(CmbContractType.Text) = "normal" Then mContractTypeFlag = "N"
If LCase(CmbContractType.Text) = "revision" Then mContractTypeFlag = "R"

''-- if Contract Type is revision then start date is min CDTF date of same Location + Commodity + Client
If LCase(CmbContractType.Text) = "revision" Then
    If CmbLocation.Text <> "" And CmbCommodity.Text <> "" And CmbClient.Text <> "" Then
        tmp = " Select MIN(CDTFDate)[CDTFDate] from Mst_GSL Where SM_Prefix = '" & mSM_Prefix & "' And ClientID = " & mClientIDRow & " And CommodityID = " & mCommodityID & " And Flag = 'O' "
        Call TmpTable
        If TmpRs.RecordCount > 0 Then
            TxtStartDate.Value = IIf(IsNull(TmpRs!CDTFDate), Date, TmpRs!CDTFDate) - 1
            TxtEndDate.Value = CDate(Format(Date, "YYYY-MM-01")) - 1    ''-- Last Month Last Date for revision Contract
        End If
    Else
        TxtStartDate.Value = Date - 1
        TxtEndDate.Value = Date
    End If
Else
    TxtStartDate.Value = Date
    TxtEndDate = Date + mDepositContractDefaultDaysRange
End If
''---------------------------------
If LCase(tmp1) <> LCase(CmbContractType.Text) Then
    Call GetGodownDetail
End If
End Sub
''-----------------------------------------------

Private Sub CmbDepositBillingCycle_GotFocus()
mOldDBillingCycle = CmbDepositBillingCycle.Text
'tmp1 = CmbDepositBillingCycle.Text
tmp = CmbDepositBillingCycle.Text
Call TableMst_BillingCycle(" Where Flag <> 'R' Order By BillingCycleID")
CmbDepositBillingCycle.Clear
If RsMst_BillingCycle.RecordCount = 0 Then
  MsgBox "No Record found !!!! "
  Exit Sub
End If
For I = 1 To RsMst_BillingCycle.RecordCount
   CmbDepositBillingCycle.AddItem RsMst_BillingCycle!BillingCycle
   RsMst_BillingCycle.MoveNext
Next
CmbDepositBillingCycle.Text = tmp
End Sub

Private Sub CmbDepositBillingCycle_LostFocus()
If CmbDepositBillingCycle.Text = "" Then mDepositBillingCycleID = Null:      Exit Sub

RsMst_BillingCycle.MoveFirst
RsMst_BillingCycle.Find "BillingCycle = '" & CmbDepositBillingCycle.Text & "'"
If RsMst_BillingCycle.EOF = True Then
   MsgBox "Invalid selection "
   CmbDepositBillingCycle.SetFocus
   Exit Sub
End If

mDepositBillingCycleID = RsMst_BillingCycle!BillingCycleID
mDepositBillingRatio = IIf(IsNull(RsMst_BillingCycle!Ratio), 0, RsMst_BillingCycle!Ratio)
mDepositBillingCycleNoofDays = IIf(IsNull(RsMst_BillingCycle!NoOfDays), 0, RsMst_BillingCycle!NoOfDays)

If LCase(mOldDBillingCycle) = LCase(CmbDepositBillingCycle.Text) Then
    Exit Sub
End If
'If mOldDBillingCycle <> "" Then
    If MsgBox("Below Rates will change as the Deposit / Withdrawal Billing cycle is changed." & vbNewLine & "Do you want to continue?", vbYesNo) = vbNo Then
        CmbDepositBillingCycle.Text = mOldDBillingCycle
        CmbDepositBillingCycle.SetFocus
        Exit Sub
    End If
'End If

If Check4.Value = 0 Then 'Calculation without bags
    With MSHFlexGrid1
        For I = 1 To .Rows - 1
            If .TextMatrix(I, 0) = strChecked Then
                If Val(.TextMatrix(I, 15)) = 1 Or Val(.TextMatrix(I, 15)) = 7 Then
                    .TextMatrix(I, 11) = Format(CalculateRate(.TextMatrix(I, 14), mDepositBillingRatio, CmbDepositBillingCycle.Text), "#0.0000")
                    mDepositOnePercentage = Format(IIf(Val(.TextMatrix(I, 11)) > 0, Val(.TextMatrix(I, 11)) / 100, 0), "#0.0000")
                    .TextMatrix(I, 12) = Format(Val(.TextMatrix(I, 11)) + (mDepositOnePercentage * .TextMatrix(I, 17)), "#0.0000")
                Else
                   .TextMatrix(I, 12) = .TextMatrix(I, 12)
                End If
                Call CheckRateByColor(CInt(I))
            End If
        Next
    End With
Else
    With MSHFlexGrid4
        For I = 1 To .Rows - 1
            If .TextMatrix(I, 0) = strChecked Then
                If Val(.TextMatrix(I, 17)) = 1 Or Val(.TextMatrix(I, 17)) = 7 Then     ''-- If Godown Type is not Lease (ie Rent not available then on Billing Cycle Only Proposed Rate get Calculate)
                    .TextMatrix(I, 13) = Format(CalculateRate(.TextMatrix(I, 16), mDepositBillingRatio, CmbDepositBillingCycle) * Val(.TextMatrix(I, 8)) / 1000, "#0.0000")
                    mDepositOnePercentage = Format(Val(.TextMatrix(I, 13)) / 100, "#0.0000")
                    .TextMatrix(I, 14) = Format(Val(.TextMatrix(I, 13)) + (mDepositOnePercentage * .TextMatrix(I, 19)), "#0.0000")
                Else
                    .TextMatrix(I, 14) = .TextMatrix(I, 14)
                End If
                Call CheckBagRateByColor(CInt(I))
            End If
        Next
    End With
End If
End Sub

Private Sub CmbForwardedTo_GotFocus()
tmp1 = CmbForwardedTo.Text
CmbForwardedTo.Clear
    ''-- Rights applyed on 05 Jan 2018
    ''If LCase(Gen_UserRole) = "superadmin" Or LCase(Gen_UserRole) = "head cmg" Or LCase(Gen_UserRole) = "power" Then
    If LCase(Gen_UserRole) = "power" Or mApprovedRights = True Then
        tmp = " Select * from Mst_Employee ME Inner Join Mst_Users MU on Me.EmployeeNo=MU.EmployeeNo Where MU.UserGroupID in (" & IIf(IsEmpty(mDepositContractCMGGroupToBeApproved), 0, mDepositContractCMGGroupToBeApproved) & ") And Me.EmployeeID <> " & Gen_EmployeeID & " And ME.Flag='Y' "
        Call TmpTable
        If TmpRs.RecordCount > 0 Then
           For I = 1 To TmpRs.RecordCount
               CmbForwardedTo.AddItem TmpRs!EmployeeName & "~" & TmpRs!EmployeeNo
               TmpRs.MoveNext
           Next
        End If
    Else
        tmp = "Select * From Mst_Employee Where EmployeeID IN (" & mUpperHierarchyEmployeeIDList & ") And  EmployeeID <> " & Gen_EmployeeID & " And Flag='Y' "
        'tmp = tmp & " Union "
        'tmp = tmp & " Select ME.* from Mst_Employee ME Inner Join Mst_Users MU on Me.EmployeeNo=MU.Employee_ID Where MU.UserGroupID in (" & mDepositContractCMGGroupToBeApproved & ") And Me.EmployeeID <> " & Gen_EmployeeID & " And ME.Flag='Y' "
        Call TmpTable
        If TmpRs.RecordCount > 0 Then
           For I = 1 To TmpRs.RecordCount
               CmbForwardedTo.AddItem TmpRs!EmployeeName & "~" & TmpRs!EmployeeNo
               TmpRs.MoveNext
           Next
        End If
    End If
    
CmbForwardedTo.Text = tmp1
End Sub

Private Sub CmbForwardedTo_LostFocus()
If CmbForwardedTo.Text = "" Then Exit Sub
Dim strSplitString() As String
Dim EmpNo_, EmpName_ As String
strSplitString = Split(CmbForwardedTo, "~")
EmpName_ = Trim(strSplitString(0))
EmpNo_ = Trim(strSplitString(1))
If CmbForwardedTo.Text = "" Then
   mForwardID = Null
   Exit Sub
End If
tmp = " Select EmployeeID from Mst_Employee where EmployeeName= '" & EmpName_ & "' And EmployeeNo = '" & EmpNo_ & "' And Flag='Y'"
Call Tmp1Table
If TmpRs1.EOF = True Then
   MsgBox "Invalid selection "
   CmbForwardedTo.SetFocus
   Exit Sub
End If
mForwardID = TmpRs1!EmployeeID
End Sub

Private Sub CmbGodown_GotFocus()

If Trim(CmbCMP.Text) = "" Then
   MsgBox "Select CMP!!!"
   CmbCMP.SetFocus
   Exit Sub
End If

tmp = CmbGodown.Text
Call TableMst_Godown(" where CloseDate is null and CMPID = " & mCMPID & " And GodownID in (Select GodownId From Mst_StorageRateGodownDetail Where SRCID=" & mSRCID & ")")
CmbGodown.Clear
If RsMst_Godown.RecordCount = 0 Then
   MsgBox "No Record found!!!"
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
   Exit Sub
End If

RsMst_Godown.MoveFirst
RsMst_Godown.Find "GodownNo= '" & CmbGodown.Text & "'"
If RsMst_Godown.EOF = True Then
   MsgBox "Invalid selection!!!"
   CmbGodown.SetFocus
   Exit Sub
End If

mGodownID = RsMst_Godown!GodownID
If tmp <> CmbGodown.Text Then
   TxtCapacity.Text = IIf(IsNull(RsMst_Godown!GodownCapacity), "0", RsMst_Godown!GodownCapacity)
   TxtGodownAddress.Text = IIf(IsNull(RsMst_Godown!Address), "", RsMst_Godown!Address)
   TxtGodownType.Text = GetUnitName(IIf(IsNull(RsMst_Godown!UnitTypeID), 0, RsMst_Godown!UnitTypeID))
End If
End Sub

Private Sub CmbGodownID_GotFocus()
If CmbCMPID.Text = "" Then
   MsgBox "Please Select CMP  !!! "
   CmbGodownID.SetFocus
   Exit Sub
End If
tmp = CmbGodownID.Text
Call TableMst_Godown("Where CMPID = " & mNewCMPID & " And GodownID In (" & mGodownList & " )  ")
CmbGodownID.Clear
If RsMst_Godown.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Godown.RecordCount
    CmbGodownID.AddItem RsMst_Godown!GodownNo
    RsMst_Godown.MoveNext
Next
CmbGodownID.Text = tmp
End Sub
Private Sub CmbGodownID_LostFocus()
If CmbGodownID.Text = "" Then
   Exit Sub
End If
RsMst_Godown.MoveFirst
RsMst_Godown.Find "GodownNo = '" & CmbGodownID.Text & "'"
If RsMst_Godown.EOF = True Then
   MsgBox "Invalid selection "
   CmbGodownID.SetFocus
   Exit Sub
End If
mNewGodownID = RsMst_Godown!GodownID
End Sub

Private Sub CmbInsurance_GotFocus()
tmp = CmbInsurance.Text
Call TableMst_InsuranceBy
CmbInsurance.Clear
If RsMst_InsuranceBy.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_InsuranceBy.RecordCount
 CmbInsurance.AddItem RsMst_InsuranceBy!InsuranceByName
 RsMst_InsuranceBy.MoveNext
Next
CmbInsurance.Text = tmp
End Sub

Private Sub CmbInsurance_LostFocus()
If CmbInsurance.Text = "" Then
   Exit Sub
End If
If RsMst_InsuranceBy.RecordCount > 0 Then RsMst_InsuranceBy.MoveFirst
RsMst_InsuranceBy.Find "InsuranceByName = '" & CmbInsurance.Text & "'"
If RsMst_InsuranceBy.EOF = True Then
   MsgBox "Invalid selection "
   CmbInsurance.SetFocus
   Exit Sub
End If
mInsuranceByID = RsMst_InsuranceBy!InsuranceByID
End Sub


Private Sub CmbLocation_GotFocus()
If Trim(CmbState.Text) = "" Then
   MsgBox "Select State!!!"
   CmbState.SetFocus
   Exit Sub
End If
tmp1 = CmbLocation.Text
tmp = CmbLocation.Text
If Gen_WcLocation = "" Then
   Call TableMst_Location(" Where StateID =" & mStateID & "")
Else
   Call TableMst_Location(" Where " & Gen_WcLocation & " And StateID = " & mStateID & " ")
End If
CmbLocation.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Record found!!!"
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
   CmbCommodity.Text = ""
    If LCase(CmbContractType.Text) = "revision" Then
     CmbClient.Text = ""
    End If
   Exit Sub
End If
If RsMst_Location.RecordCount > 0 Then RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocation.Text & "'"

If RsMst_Location.EOF = True Then
   MsgBox "Invalid Location selection "
   CmbLocation.SetFocus
   Exit Sub
End If
If RsMst_Location!Locationname <> CmbLocation Then
   Call ClearData
   
If LCase(CmbContractType.Text) = "revision" Then
 CmbClient.Text = ""
End If
   
End If
mLocationID = RsMst_Location.Fields("LocationID")
mSM_Prefix = RsMst_Location.Fields("SM_Prefix")

''-- 14 Feb 2018
''-- if Contract Type is revision then start date is min CDTF date of same Location + Commodity + Client
If LCase(CmbContractType.Text) = "revision" Then
    If CmbLocation.Text <> "" And CmbCommodity.Text <> "" And CmbClient.Text <> "" Then
        tmp = " Select MIN(CDTFDate)[CDTFDate] from Mst_GSL Where SM_Prefix = '" & mSM_Prefix & "' And ClientID = " & mClientIDRow & " And CommodityID = " & mCommodityID & " And Flag = 'O' "
        Call TmpTable
        If TmpRs.RecordCount > 0 Then
            TxtStartDate.Value = IIf(IsNull(TmpRs!CDTFDate), Date, TmpRs!CDTFDate) - 1
            TxtEndDate.Value = CDate(Format(Date, "YYYY-MM-01")) - 1    ''-- Last Month Last Date for revision Contract
        End If
    Else
        TxtStartDate.Value = Date - 1
        TxtEndDate.Value = Date
    End If
Else
    TxtStartDate.Value = Date
    TxtEndDate = Date + mDepositContractDefaultDaysRange
End If
''---------------------------------

If LCase(tmp1) <> LCase(CmbLocation.Text) Then
    CmbCommodity.Text = ""
    Call GetGodownDetail
End If
End Sub

Private Sub CmbRatePer_LostFocus()
If CmbRatePer.Text = "" Then Exit Sub
If LCase(CmbRatePer) <> "included" And LCase(CmbRatePer) <> "excluded" Then
   MsgBox "Invalid Selection"
   CmbRatePer.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmbSClient_GotFocus()
tmp = CmbSClient.Text
If Check3.Value = vbUnchecked Then
   Call TableMst_Client("Where Converted = 0 And ExchangeTypeID = " & GetExchangeID("Non NCDEX") & " Order By ClientName ")
Else
   Call TableMst_Client("Where Converted = 1 And ExchangeTypeID = " & GetExchangeID("Non NCDEX") & " Order By ClientName ")
End If
CmbSClient.Clear
If RsMst_Client.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Client.RecordCount
    CmbSClient.AddItem RsMst_Client!ClientName
    RsMst_Client.MoveNext
Next
CmbSClient.Text = tmp
End Sub

Private Sub CmbSClient_LostFocus()
If CmbSClient.Text = "" Then
   Exit Sub
End If
RsMst_Client.MoveFirst
RsMst_Client.Find "ClientName = '" & CmbSClient.Text & "'"
If RsMst_Client.EOF = True Then
   MsgBox "Invalid selection "
   CmbSClient.SetFocus
   Exit Sub
End If
mSClientIDRow = RsMst_Client!ClientID
End Sub

Private Sub CmbSCommodity_GotFocus()
If Trim(CmbSState.Text) = "" Then
   MsgBox "Select State!!!"
   CmbSState.SetFocus
   Exit Sub
End If
If Trim(CmbSLocation.Text) = "" Then
   MsgBox "Select Location!!!"
   CmbSLocation.SetFocus
   Exit Sub
End If
tmp = CmbSCommodity.Text
Call TableMst_Commodity(" Where CommodityID IN ( SELECT Distinct CommodityID FROM Mst_LocationCommodityMapping Where LocationID = " & mSLocationID & " ) ")
CmbSCommodity.Clear
If RsMst_Commodity.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
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
If RsMst_Commodity.EOF = True Then
   MsgBox "Invalid selection "
   CmbSCommodity.SetFocus
   Exit Sub
End If
mSCommodityID = RsMst_Commodity!CommodityID
End Sub

''-- 14 Feb 2018
Private Sub CmbSContractType_LostFocus()
If CmbSContractType.Text = "" Then Exit Sub
If LCase(CmbSContractType.Text) <> "normal" And LCase(CmbSContractType.Text) <> "revision" Then
   MsgBox "Invalid Selection !!!"
   CmbSContractType.SetFocus
   Exit Sub
End If
If LCase(CmbSContractType.Text) = "normal" Then mSContractTypeFlag = "N"
If LCase(CmbSContractType.Text) = "revision" Then mSContractTypeFlag = "R"
End Sub
Private Sub CmbSType_LostFocus()
If CmbSType.Text = "" Then Exit Sub
If LCase(CmbSType.Text) <> "quantity" And LCase(CmbSType.Text) <> "bags" Then
   MsgBox "Invalid Selection !!!"
   CmbSType.SetFocus
   Exit Sub
End If
If LCase(CmbSType.Text) = "quantity" Then mSType = "N"
If LCase(CmbSType.Text) = "bags" Then mSType = "Y"
End Sub
''------------------

Private Sub CmbSForwardedTo_GotFocus()
tmp1 = CmbSForwardedTo.Text
''CmbSForwardedTo.Clear
''tmp = " Select ME.EmployeeID,ME.EmployeeName,ME.EmployeeNo,MUR.GroupName"
''tmp = tmp & " from Mst_Employee ME"
''tmp = tmp & " Inner Join Mst_USers MU on ME.EmployeeNo=MU.Employee_ID"
''tmp = tmp & " Inner Join Mst_UserRole MUR on MU.UserGroupID=MUR.UserGroupID"
''tmp = tmp & " Where ME.Flag = 'Y'"
''Call Tmp2Table
''CmbSForwardedTo.Clear
''If TmpRs2.RecordCount = 0 Then
''   MsgBox "No Record found !!!! "
''   CmbSForwardedTo.SetFocus
''   Exit Sub
''End If
''For I = 1 To TmpRs2.RecordCount
''    CmbSForwardedTo.AddItem TmpRs2!EmployeeName & "~" & TmpRs2!EmployeeNo
''    TmpRs2.MoveNext
''Next

CmbSForwardedTo.Clear
''-- Rights applyed on 05 Jan 2018
''If LCase(Gen_UserRole) = "superadmin" Or LCase(Gen_UserRole) = "head cmg" Or LCase(Gen_UserRole) = "power" Then
If LCase(Gen_UserRole) = "power" Or mApprovedRights = True Then
    tmp = " Select * from Mst_Employee ME Inner Join Mst_Users MU on Me.EmployeeNo=MU.EmployeeNo Where MU.UserGroupID in (" & mDepositContractCMGGroupToBeApproved & ")  Or MU.EmployeeNo in (" & mDepositContractAdditionalGroupToBeApprovedList & ") And ME.Flag='Y'"
    Call TmpTable
    If TmpRs.RecordCount > 0 Then
       For I = 1 To TmpRs.RecordCount
           CmbSForwardedTo.AddItem TmpRs!EmployeeName & "~" & TmpRs!EmployeeNo
           TmpRs.MoveNext
       Next
    End If
Else
    tmp = "Select * From Mst_Employee Where EmployeeID IN (" & mUpperHierarchyEmployeeIDList & ") And  EmployeeID <> " & Gen_EmployeeID & " And Flag='Y' "
    tmp = tmp & " Union "
    tmp = tmp & " Select ME.* from Mst_Employee ME Inner Join Mst_Users MU on Me.EmployeeNo=MU.EmployeeNo Where MU.UserGroupID in (" & mDepositContractCMGGroupToBeApproved & ") And Me.EmployeeID <> " & Gen_EmployeeID & " And ME.Flag='Y' "
    Call TmpTable
    If TmpRs.RecordCount > 0 Then
       For I = 1 To TmpRs.RecordCount
           CmbSForwardedTo.AddItem TmpRs!EmployeeName & "~" & TmpRs!EmployeeNo
           TmpRs.MoveNext
       Next
    End If
End If
CmbSForwardedTo.Text = tmp1

End Sub


Private Sub CmbSForwardedTo_LostFocus()
If CmbSForwardedTo.Text = "" Then Exit Sub

Dim strSplitString() As String
Dim EmpNo_, EmpName_ As String

strSplitString = Split(CmbSForwardedTo, "~")
EmpName_ = Trim(strSplitString(0))
EmpNo_ = Trim(strSplitString(1))

If CmbSForwardedTo.Text = "" Then
   mForwardID = ""
   Exit Sub
End If

TmpRs.MoveFirst
TmpRs.Find "EmployeeName = '" & EmpName_ & "'"
If TmpRs.EOF = True Then
   MsgBox "Invalid selection "
   CmbSForwardedTo.SetFocus
   Exit Sub
End If
mSForwardID = TmpRs!EmployeeID
End Sub


Private Sub CmbSLocation_GotFocus()
If Trim(CmbSState.Text) = "" Then
   MsgBox "Select State!!!"
   CmbSState.SetFocus
   Exit Sub
End If
tmp1 = CmbSLocation.Text
tmp = CmbSLocation.Text
If Gen_WcLocation = "" Then
   Call TableMst_Location(" Where StateID =" & mSStateID & "")
Else
   Call TableMst_Location(" Where " & Gen_WcLocation & " And StateID = " & mSStateID & " ")
End If
CmbSLocation.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Record found!!!"
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
   CmbSCommodity.Text = ""
   Exit Sub
End If
If RsMst_Location.RecordCount > 0 Then RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbSLocation.Text & "'"

If RsMst_Location.EOF = True Then
   MsgBox "Invalid Location selection "
   CmbSLocation.SetFocus
   Exit Sub
End If
mSLocationID = RsMst_Location.Fields("LocationID")
mSMPrefix = RsMst_Location.Fields("SM_Prefix")
If LCase(tmp1) <> LCase(CmbSLocation.Text) Then CmbSCommodity.Text = ""
End Sub


Private Sub CmbSState_GotFocus()
tmp = CmbSState.Text
tmp1 = CmbSState.Text
If Gen_WcState <> "" Then
   Call TableMst_State("Where" & Gen_WcState)
Else
   Call TableMst_State
End If
CmbSState.Clear
If RsMst_State.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   CmbSState.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_State.RecordCount
    CmbSState.AddItem RsMst_State!StateName
    RsMst_State.MoveNext
Next
CmbSState.Text = tmp
End Sub


Private Sub CmbSState_LostFocus()
If CmbSState.Text = "" Then
   CmbSLocation.Text = ""
   CmbSCommodity.Text = ""
   Exit Sub
End If
RsMst_State.MoveFirst
RsMst_State.Find "StateName = '" & CmbSState.Text & "'"

If RsMst_State.EOF = True Then
   MsgBox "Invalid State selection "
   CmbSState.SetFocus
   Exit Sub
End If

mSStateID = RsMst_State.Fields(0)
If tmp1 <> CmbSState.Text Then
    CmbSLocation.Text = ""
    CmbSCommodity.Text = ""
End If
End Sub

Private Sub CmbSStatus_GotFocus()
CmbSStatus.Clear
CmbSStatus.AddItem "System Approved"
CmbSStatus.AddItem "Approved"
CmbSStatus.AddItem "Forwarded"
CmbSStatus.AddItem "Rejected"
CmbSStatus.AddItem "Closed"
End Sub

Private Sub CmbSStatus_LostFocus()
If CmbSStatus.Text = "" Then Exit Sub
If LCase(CmbSStatus.Text) <> "system approved" And LCase(CmbSStatus.Text) <> "approved" And LCase(CmbSStatus.Text) <> "rejected" And LCase(CmbSStatus.Text) <> "forwarded" And LCase(CmbSStatus.Text) <> "closed" Then
   MsgBox "Invalid Selection !!!"
   CmbSStatus.SetFocus
   Exit Sub
End If

If LCase(CmbSStatus.Text) = "system approved" Then mSStatus = "'S'"
If LCase(CmbSStatus.Text) = "approved" Then mSStatus = "'A'"
If LCase(CmbSStatus.Text) = "rejected" Then mSStatus = "'R'"
If LCase(CmbSStatus.Text) = "forwarded" Then mSStatus = "'F'"
If LCase(CmbSStatus.Text) = "closed" Then mSStatus = "'C'"
End Sub

Private Sub CmbState_GotFocus()
tmp = CmbState.Text
tmp1 = tmp
Call TableMst_State(" Where Flag = 'A' " & IIf(Gen_WcState <> "", "And " & Gen_WcState, ""))
CmbState.Clear
If RsMst_State.RecordCount = 0 Then
   MsgBox "No Record Found!"
   Exit Sub
End If
For I = 1 To RsMst_State.RecordCount
    CmbState.AddItem RsMst_State!StateName
    RsMst_State.MoveNext
Next
CmbState.Text = tmp
End Sub

Private Sub CmbState_LostFocus()
If CmbState.Text = "" Then
   CmbLocation.Text = ""
   CmbCommodity.Text = ""
   Exit Sub
End If
RsMst_State.MoveFirst
RsMst_State.Find "StateName = '" & CmbState.Text & "'"
If RsMst_State.EOF = True Then
   MsgBox "Invalid selection "
   CmbState.SetFocus
   Exit Sub
End If
If LCase(CmbState.Text) <> LCase(tmp1) Then
   CmbLocation.Text = ""
   CmbCommodity.Text = ""
   Call GetGodownDetail
End If
mStateID = RsMst_State!StateID
End Sub

Private Sub CmbStatus_GotFocus()
tmp = CmbStatus.Text
CmbStatus.Clear
'If LCase(Gen_UserRole) = "superadmin" Or LCase(Gen_UserRole) = "head cmg" Or LCase(Gen_UserRole) = "power" Then
If mApprovedRights = True Or LCase(Gen_UserRole) = "power" Then
   If Edit_Flg <> "A" Then
      CmbStatus.AddItem "Approved"
      CmbStatus.AddItem "Forwarded"
      CmbStatus.AddItem "Rejected"
   Else
      CmbStatus.AddItem "Forwarded"
      CmbStatus.AddItem "Approved"
   End If
Else
   CmbStatus.AddItem "Forwarded"
   If Edit_Flg <> "A" Then
    ''---- if login User is in Upper Hierarchy of  DP Created User then he can Approve/Reject the DP
    If LoginEmployeeInHierarchy(Gen_UserLoginID) = True Then
        CmbStatus.AddItem "Approved"
        CmbStatus.AddItem "Rejected"
    Else
       If LoginEmployeeInAdditionalGroup(Gen_UserLoginID) = True Then
          CmbStatus.AddItem "Approved"
          CmbStatus.AddItem "Rejected"
       Else
          CmbStatus.AddItem "Rejected"
       End If
    End If
   End If
End If
CmbStatus = tmp
End Sub

Private Sub CmbStatus_LostFocus()
If CmbStatus.Text = "" Then Exit Sub
If LCase(CmbStatus.Text) <> "approved" And LCase(CmbStatus.Text) <> "rejected" And LCase(CmbStatus.Text) <> "forwarded" Then
''And LCase(CmbStatus.Text) <> "closed"
   MsgBox "Invalid Selection !!!"
   CmbStatus.SetFocus
   Exit Sub
End If

''If LCase(CmbStatus.Text) = "forwarded" Then
''   CmbForwardedTo.Enabled = True
''Else
''   CmbForwardedTo.Enabled = False
''End If
mSelectStatus = UCase(Left(CmbStatus.Text, 1))
End Sub


Private Sub CmdAdd2List_Click()

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

If TxtFinalPrice.Text = "" Then
   MsgBox "Blank Final Price not Allowed!!!"
   TxtFinalPrice.SetFocus
   Exit Sub
End If

If CmdAdd2List.Caption = "Update List" Then
   r = MSHFlexGrid1.RowSel
Else
  r = MSHFlexGrid1.Rows - 1
  MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
End If

MSHFlexGrid1.TextMatrix(r, 0) = CmbCMP.Text
MSHFlexGrid1.TextMatrix(r, 1) = CmbGodown.Text
MSHFlexGrid1.TextMatrix(r, 2) = TxtGodownAddress.Text
MSHFlexGrid1.TextMatrix(r, 3) = TxtGodownType.Text
MSHFlexGrid1.TextMatrix(r, 4) = TxtRent.Text
MSHFlexGrid1.TextMatrix(r, 5) = TxtCapacity.Text
MSHFlexGrid1.TextMatrix(r, 6) = TxtFinalPrice
MSHFlexGrid1.TextMatrix(r, 7) = mGodownID

Call ClearFrameControl
End Sub


Private Sub CmdCancel_Click()
Frame3.Visible = False
End Sub

Private Sub CmbSVerified_LostFocus()
If CmbSVerified.Text = "" Then Exit Sub
If LCase(CmbSVerified.Text) <> "yes" And LCase(CmbSVerified.Text) <> "no" Then
   MsgBox "Invalid Selection !!!"
   CmbSVerified.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmdAddAttachment_Click()
Dim r As Variant
With MSHFlexGrid5
If lblAttachment.Caption = "" Then
   MsgBox "Blank Attachment not allowed!!!", vbOKOnly + vbInformation, App.Title
   Exit Sub
End If

If TxtAttachmentRemarks = "" Then
   MsgBox "Blank Remarks not allowed!!!", vbOKOnly + vbInformation, App.Title
   TxtAttachmentRemarks.SetFocus
   Exit Sub
End If

If CmdAddAttachment.Caption = "Update Attachment" Then
   r = .RowSel
   If DuplicateFileName(GetFileFromPath(lblAttachment.Caption), CInt(r)) = True Then
      MsgBox "Duplicate File Name not allowed."
      Exit Sub
   End If
Else
   r = .Rows - 1
   If DuplicateFileName(GetFileFromPath(lblAttachment.Caption)) = True Then
      MsgBox "Duplicate File Name not allowed."
      Exit Sub
   End If
   .Rows = .Rows + 1
End If
.TextMatrix(r, 0) = ""
.TextMatrix(r, 1) = GetFileFromPath(lblAttachment.Caption)
.TextMatrix(r, 2) = TxtAttachmentRemarks
.TextMatrix(r, 3) = lblAttachment.Caption
Call ClearFrameAttachment
End With
End Sub

Private Function DuplicateFileName(FileName_ As String, Optional Index As Integer) As Boolean
Dim mR, mC As Variant
    With MSHFlexGrid5
       For mR = 1 To .Rows - 2
            If Index = 0 Then
                If LCase(FileName_) = LCase(.TextMatrix(mR, 1)) Then
                    DuplicateFileName = True
                    Exit Function
                End If
            Else
                If LCase(FileName_) = LCase(.TextMatrix(mR, 1)) Then
                   If Index <> 0 And Index <> mR Then
                    DuplicateFileName = True
                    Exit Function
                   End If
                End If
            End If
       Next mR
    End With
DuplicateFileName = False
End Function


Private Sub ClearFrameAttachment()
    lblAttachment.Caption = ""
    TxtAttachmentRemarks = ""
    CmdAddAttachment.Caption = "Add Attachment"
End Sub

Private Sub CmdAddNewBagSize_Click()
Dim mR As Double

    With MSHFlexGrid4
        mGodownList = "0"
        For mR = 1 To .Rows - 1
            If .TextMatrix(mR, 0) = strChecked Then mGodownList = Val(.TextMatrix(mR, 11)) & "," & mGodownList
        Next mR
    End With
    If mGodownList = "0" Then
        MsgBox "Select atleast one godown to add more bag size for same."
        Exit Sub
    End If
    
    Frame5.Visible = True
    Frame1.Enabled = False
    SSTab1.Enabled = False
    ButtonFrm.Enabled = False
End Sub

Private Sub CmdAddtoListNewBagSize_Click()
If CmbCMPID.Text = "" Then
    MsgBox "Select CMP Name."
    CmbCMPID.SetFocus
    Exit Sub
End If
If CmbGodownID.Text = "" Then
    MsgBox "Select Godown Name."
    CmbGodownID.SetFocus
    Exit Sub
End If
If Val(TxtMinNewBagSize.Text) = 0 Then
    MsgBox "Zero / Blank Minimum Bag Size not allowed." & vbNewLine & "Minimum Bag size Limit: " & mMinBagSize & vbNewLine & "Maximum Bag size Limit: " & mMaxBagSize
    TxtMinNewBagSize.SetFocus
    Exit Sub
End If
If Val(TxtMaxNewBagSize.Text) = 0 Then
    MsgBox "Zero / Blank Maximum Bag Size not allowed." & vbNewLine & "Minimum Bag size Limit: " & mMinBagSize & vbNewLine & "Maximum Bag size Limit: " & mMaxBagSize
    TxtMaxNewBagSize.SetFocus
    Exit Sub
End If
If Val(TxtMaxNewBagSize) = Val(TxtMinNewBagSize) Then
    MsgBox "Minimum bag size and Maximim bag size should not be same." & vbNewLine & "Minimum Bag size Limit: " & mMinBagSize & vbNewLine & "Maximum Bag size Limit: " & mMaxBagSize
    TxtMaxNewBagSize.SetFocus
    Exit Sub
End If
If Val(TxtMaxNewBagSize) < Val(TxtMinNewBagSize) Then
    MsgBox "Maximum Bag Size less then Minimum Bag Size not allowed." & vbNewLine & "Minimum Bag size Limit: " & mMinBagSize & vbNewLine & "Maximum Bag size Limit: " & mMaxBagSize
    TxtMaxNewBagSize.SetFocus
    Exit Sub
End If
If Val(TxtMinNewBagSize.Text) < Val(mMinBagSize) Or Val(TxtMinNewBagSize.Text) > Val(mMaxBagSize) Then
     MsgBox "Bag Size is not as per define Commodity Minimum-Maximum bag size range." & vbNewLine & "Minimum Bag size Limit: " & mMinBagSize & vbNewLine & "Maximum Bag size Limit: " & mMaxBagSize
     TxtMinNewBagSize.SetFocus
     Exit Sub
End If
If Val(TxtMaxNewBagSize.Text) < Val(mMinBagSize) Or Val(TxtMaxNewBagSize.Text) > Val(mMaxBagSize) Then
     MsgBox "Bag Size is not as per define Commodity Minimum-Maximum bag size range." & vbNewLine & "Minimum Bag size Limit: " & mMinBagSize & vbNewLine & "Maximum Bag size Limit: " & mMaxBagSize
     TxtMaxNewBagSize.SetFocus
     Exit Sub
End If
If BagSizeRangeCheck(mNewGodownID, CDbl(TxtMinNewBagSize.Text), CInt(0)) = False Then
     MsgBox "Minimum Bag size is already define in given Bag size range for same selected godown." & vbNewLine & "Minimum Bag size Limit: " & mMinBagSize & vbNewLine & "Maximum Bag size Limit: " & mMaxBagSize
     TxtMinNewBagSize.SetFocus
     Exit Sub
End If
If BagSizeRangeCheck(mNewGodownID, CDbl(TxtMaxNewBagSize.Text), CInt(0)) = False Then
     MsgBox "Maximum Bag size is already define in given Bag size range for same selected godown." & vbNewLine & "Minimum Bag size Limit: " & mMinBagSize & vbNewLine & "Maximum Bag size Limit: " & mMaxBagSize
     TxtMaxNewBagSize.SetFocus
     Exit Sub
End If

''-- Finally add in Bag Wise main grid
With MSHFlexGrid4
''-- Find the row index of any grid of same godown then move all the cell data in new created row with new bag size
Dim mRID As Double, mLastindex As Double, mNewindex As Double
Dim mR As Double

    For mR = 1 To .Rows - 1
        If Val(.TextMatrix(mR, 11)) = Val(mNewGodownID) Then
            mRID = mR
        End If
    Next mR
    If mRID = 0 Then
        MsgBox "Add new godown not in list."
        Exit Sub
    End If

    mLastindex = .Rows - 1
    .Rows = .Rows + 1
    mNewindex = .Rows - 1
    
    .row = mNewindex
    .Col = 0
    .CellFontName = "Wingdings"
    .CellFontSize = 12
    .CellAlignment = flexAlignCenterCenter
    .Text = strChecked

    .TextMatrix(mNewindex, 1) = .TextMatrix(mRID, 1)
    .TextMatrix(mNewindex, 2) = .TextMatrix(mRID, 2)
    .TextMatrix(mNewindex, 3) = .TextMatrix(mRID, 3)
    .TextMatrix(mNewindex, 4) = .TextMatrix(mRID, 4)
    .TextMatrix(mNewindex, 5) = .TextMatrix(mRID, 5)
    .TextMatrix(mNewindex, 6) = .TextMatrix(mRID, 6)
    .TextMatrix(mNewindex, 7) = TxtMinNewBagSize.Text
    .TextMatrix(mNewindex, 8) = TxtMaxNewBagSize.Text
    .TextMatrix(mNewindex, 9) = .TextMatrix(mRID, 9)
    .TextMatrix(mNewindex, 10) = .TextMatrix(mRID, 10)
    .TextMatrix(mNewindex, 11) = .TextMatrix(mRID, 11)
    .TextMatrix(mNewindex, 12) = ""
    .TextMatrix(mNewindex, 13) = .TextMatrix(mRID, 13)
    .TextMatrix(mNewindex, 14) = .TextMatrix(mRID, 14)
    .TextMatrix(mNewindex, 15) = .TextMatrix(mRID, 15)
    .TextMatrix(mNewindex, 16) = .TextMatrix(mRID, 16)
    .TextMatrix(mNewindex, 17) = .TextMatrix(mRID, 17)
    .TextMatrix(mNewindex, 18) = .TextMatrix(mRID, 18)
    .TextMatrix(mNewindex, 19) = .TextMatrix(mRID, 19)
    .TextMatrix(mNewindex, 20) = "M"
    
    If Val(.TextMatrix(mNewindex, 17)) = 1 Or Val(.TextMatrix(mNewindex, 17)) = 7 Then
        .TextMatrix(mNewindex, 9) = Format(CalculateRate(.TextMatrix(mNewindex, 15), mBillingRatio, CmbBillingCycle) * Val(.TextMatrix(mNewindex, 8)) / 1000, "#0.0000")
        mBillingOnePercentage = Format(Val(.TextMatrix(mNewindex, 9)) / 100, "#0.0000")
        .TextMatrix(mNewindex, 10) = Format(Val(.TextMatrix(mNewindex, 9)) + (mBillingOnePercentage * .TextMatrix(mNewindex, 19)), "#0.0000")

        .TextMatrix(mNewindex, 13) = Format(CalculateRate(.TextMatrix(mNewindex, 16), mDepositBillingRatio, CmbDepositBillingCycle) * Val(.TextMatrix(mNewindex, 8)) / 1000, "#0.0000")
        mDepositOnePercentage = Format(Val(.TextMatrix(mNewindex, 13)) / 100, "#0.0000")
        .TextMatrix(mNewindex, 14) = Format(Val(.TextMatrix(mNewindex, 13)) + (mDepositOnePercentage * .TextMatrix(mNewindex, 19)), "#0.0000")
    Else
        .TextMatrix(mNewindex, 10) = Format((.TextMatrix(mRID, 10) / .TextMatrix(mRID, 8)) * .TextMatrix(mNewindex, 8), "#0.0000")
        .TextMatrix(mNewindex, 14) = Format((.TextMatrix(mRID, 14) / .TextMatrix(mRID, 8)) * .TextMatrix(mNewindex, 8), "#0.0000")
    End If
    
    Call CheckBagRateByColor(CInt(mNewindex))
    TxtMinNewBagSize = ""
    TxtMaxNewBagSize = ""
    
End With

End Sub

Private Sub CmdCloseCancel_Click()
   Frame3.Visible = False
End Sub

Private Sub CmdCloseConform_Click()
Dim mLastProcess As Variant
If Trim(TxtCloseRemark) = "" Then
    MsgBox "Close remark can not blank."
    Exit Sub
End If
    Select Case MsgBox("Do you want to Close Deposit Contract.", vbYesNo)
        Case vbNo
            Exit Sub
        Case vbYes
        
            tmp = " Select Max(DepositContractProcessUpto) [DepositContractProcessUpto] From Mst_Config  "
            Call TmpTable
            mLastProcess = CDate(IIf(IsNull(TmpRs!DepositContractProcessUpto), Date, TmpRs!DepositContractProcessUpto))
        
            tmp = " UPDATE Txn_DepositContract SET ApprovalFlag = 'C',EndDate = '" & Format(mLastProcess + 1, "YYYY-MM-DD") & "',ClosedOn = GETDATE(),ClosedBy= '" & Gen_UserLoginID & "',ClosedRemark = '" & TxtCloseRemark & "',Concurrency = GETDATE(),UpdatedDate = GETDATE(),UpdatedBy = '" & Gen_UserLoginID & "' Where DepConID = " & Val(TxtDepositContractID) & " And ApprovalFlag In ('A','S') "
            Call TmpTable
            
            ''---------------
            Dim TOADD As Variant
            Dim TOCC, mMsg, mGridMsg As Variant
            Dim strSplitString() As String
            Dim EmpNo_, EmpName_ As String
            Dim mR As Variant
            Dim mColor As Variant
            
            mMsg = ""
            TOADD = ""
            TOCC = ""
            
            mMsg = "<table border=1 width='100%' align='center'>"
            mMsg = mMsg & "<tr><td width='40%'><b>Deposit Contract ID</b></td><td width='60%'>" & TxtDepositContractID.Text & "</td></tr>"
            mMsg = mMsg & "<tr><td width='40%'><b>State</b></td><td width='60%'>" & CmbState.Text & "</td></tr>"
            mMsg = mMsg & "<tr><td width='40%'><b>Location</b></td><td width='60%'>" & CmbLocation.Text & "</td></tr>"
            mMsg = mMsg & "<tr><td width='40%'><b>Commodity</b></td><td width='60%'>" & CmbCommodity.Text & "</td></tr>"
            mMsg = mMsg & "<tr><td width='40%'><b>Client Name</b></td><td width='60%'>" & CmbClient.Text & "</td></tr>"
            mMsg = mMsg & "<tr><td width='40%'><b>Start Date</b></td><td width='60%'>" & Format(TxtStartDate, "DD-MMM-yyyy") & "</td></tr>"
            mMsg = mMsg & "<tr><td width='40%'><b>End Date</b></td><td width='60%'>" & Format(mLastProcess + 1, "DD-MMM-YYYY") & "</td></tr>"
            mMsg = mMsg & "<tr><td width='40%'><b>Insured By</b></td><td width='60%'>" & CmbInsurance.Text & "</td></tr>"
            mMsg = mMsg & "<tr><td width='40%'><b>Billing Cycle</b></td><td width='60%'>" & CmbBillingCycle.Text & "</td></tr>"
            mMsg = mMsg & "<tr><td width='40%'><b>Depository Billing Cycle</b></td><td width='60%'>" & CmbDepositBillingCycle.Text & "</td></tr>"
            mMsg = mMsg & "<tr><td width='40%'><b>Status</b></td><td width='60%'> Closed </td></tr>"
            mMsg = mMsg & "<tr><td width='40%'><b>Forwared TO</b></td><td width='60%'>" & CmbForwardedTo.Text & "</td></tr>"
            mMsg = mMsg & "<tr><td width='40%'><b>Remarks</b></td><td width='60%'>" & TxtRemarks & "</td></tr>"
            mMsg = mMsg & "</table>"
            
            If mForwardID <> Gen_UserID Then
                TOADD = GetEmployeeEmailID(mForwardID)
            Else
                tmp = " Select EmployeeID,EmployeeNo,EmployeeEmailID from Mst_Employee Where EmployeeNo = '" & IIf(IsNull(RsTxn_DepositContract!CreatedBy), 0, RsTxn_DepositContract!CreatedBy) & "' And Isnull(Flag,'Y') = 'Y'  "
                Call TmpTable
                If TmpRs.RecordCount > 0 Then
                    TOADD = IIf(IsNull(TmpRs!EmployeeEmailID), "", TmpRs!EmployeeEmailID)
                End If
            End If

            TOCC = Gen_UserEmailID
            
            ''---- Location Wise Email and Reporty
            tmp = " Select Distinct ME.EmployeeEmailID[First],MEE.EmployeeEmailID[Second],MEEE.EmployeeEmailID[Third],MEEEE.EmployeeEmailID[Fourth]"
            tmp = tmp & " From Mst_EmpCMPMapping MEC"
            tmp = tmp & " Inner Join Mst_Employee ME on MEC.EmployeeID = ME.EmployeeID And ISnull(ME.Flag,'Y') = 'Y' And ME.DesignationID <> 1"
            tmp = tmp & " Left Join Mst_Employee MEE On ME.DirectReportingTo = MEE.EmployeeID And ISnull(MEE.Flag,'Y') = 'Y' And MEE.DesignationID <> 1"
            tmp = tmp & " Left Join Mst_Employee MEEE On MEE.DirectReportingTo = MEEE.EmployeeID And ISnull(MEEE.Flag,'Y') = 'Y' And MEEE.EmployeeID <> 20"
            tmp = tmp & " Left Join Mst_Employee MEEEE On MEEE.DirectReportingTo = MEEEE.EmployeeID And ISnull(MEEEE.Flag,'Y') = 'Y'  And MEEE.EmployeeID <> 20"
            tmp = tmp & " Where MEC.LocationID = " & mLocationID & ""
            Call TmpTable
            For I = 1 To TmpRs.RecordCount
                If InStr(LCase(TOCC), LCase(IIf(IsNull(TmpRs!First), "", TmpRs!First))) = 0 Then TOCC = TOCC & "," & IIf(IsNull(TmpRs!First), "", TmpRs!First)
                If InStr(LCase(TOCC), LCase(IIf(IsNull(TmpRs!Second), "", TmpRs!Second))) = 0 Then TOCC = TOCC & "," & IIf(IsNull(TmpRs!Second), "", TmpRs!Second)
                If InStr(LCase(TOCC), LCase(IIf(IsNull(TmpRs!Third), "", TmpRs!Third))) = 0 Then TOCC = TOCC & "," & IIf(IsNull(TmpRs!Third), "", TmpRs!Third)
                If InStr(LCase(TOCC), LCase(IIf(IsNull(TmpRs!Fourth), "", TmpRs!Fourth))) = 0 Then TOCC = TOCC & "," & IIf(IsNull(TmpRs!Fourth), "", TmpRs!Fourth)
                TmpRs.MoveNext
            Next I
            ''---------
            
''            tmp = " Select ME.EmployeeID,ME.EmployeeNo,ME.EmployeeEmailID from Mst_Employee ME Inner Join Mst_Users MU on Me.EmployeeNo=MU.Employee_ID Where MU.UserGroupID in (" & mDepositContractCMGGroupToBeApproved & ") And ME.Flag='Y' "
''            Call TmpTable
''            For i = 1 To TmpRs.RecordCount
''                If InStr(LCase(TOCC), LCase(IIf(IsNull(TmpRs!EmployeeEmailID), "", TmpRs!EmployeeEmailID))) = 0 Then
''                    TOCC = TOCC & "," & IIf(IsNull(TmpRs!EmployeeEmailID), "", TmpRs!EmployeeEmailID)
''                End If
''                TmpRs.MoveNext
''            Next i
            
            ''-- 11 Jan 2018
            If TOCC <> "" Then
                TOCC = TOCC & ",CMG@ncml.com"
            End If
            
            mMsgSubject = "Deposit Contract ID " & TxtDepositContractID
            Call SendMail(TOADD, TOCC, mMsgSubject, mMsg, "")
            ''---------------
                        
            RsTxn_DepositContract.Requery
            If RsTxn_DepositContract.RecordCount > 0 Then
                RsTxn_DepositContract.Find "DepConID = " & Val(TxtDepositContractID)
                Call Indata
                Exit Sub
            End If
            
    End Select
End Sub

Private Sub CmdClosed_Click()
    ''-- for Concurrency Check Only for Update
    tmp = " Select * from Txn_DepositContract Where DepConID = " & Val(TxtDepositContractID) & " And CONVERT(varchar(30),Concurrency,120) > '" & Format(RsTxn_DepositContract!Concurrency, "YYYY-MM-DD hh:mm:ss") & "' "
    Call TmpTable
    If TmpRs.EOF = False Then
        MsgBox "Sorry, System not allow to update the data because same data in updated by another user." & vbNewLine & vbNewLine & "Please reopen your form to get update information and process.", vbInformation + vbOKOnly
        Exit Sub
    End If
    ''--------------------------------
    If LCase(mStatus) <> "a" And LCase(mStatus) <> "s" Then
        MsgBox "Deposit Contract status not 'Approved'/'System Approved' to close."
        Exit Sub
    End If
    Frame3.Width = SSTab1.Width
    Frame3.Height = SSTab1.Height + ButtonFrm.Height + 100
    Frame3.Left = SSTab1.Left
    Frame3.Top = SSTab1.Top - 50
    Frame3.Visible = True
    TxtCloseRemark.SetFocus
End Sub

Private Sub CmdCopy_Click()
mCStateName = TxtState
mCStateNameID = mStateID
mCLocationName = TxtLocation
mCLocationID = mLocationID
mCCommodityName = TxtCommodity
mCCommodityID = mCommodityID
mCClientName = TxtClient
mCClientID = mClientIDRow
mCInsuranceBy = TxtInsurance
mCInsuranceID = mInsuranceByID
mCBags = IIf(Check4.Value = vbChecked, True, False)
mCTentativeDuration = TxtTentativeDuration
mCLockInPeriod = TxtLockInPeriod
mCQuantum = TxtQuantum
mCBillingCycle = TxtBillingCycle
mCBillingCycleID = mBillingCycleID
mCDBillingCycle = TxtDepositBillingCycle
mCDBillingCycleID = mDepositBillingCycleID
mCContractTypeFlag = mContractTypeFlag
mCContractTypeText = TxtContractType

mCopyFlag = True
Edit_Flg = "A"
Call GButtonLock(Me, False)
Call ClearControl
Call Grid_Head6
CmbStatus.Text = "Forwarded"
CmbBillingCycle.Text = "Monthly"
CmbDepositBillingCycle.Text = "Monthly"
TxtTentativeDuration.Text = "12"
''--- 17 Feb 2018
TxtLockInPeriod.Text = "1"
TxtQuantum.Text = "1"
''-----------
Call GetBillingCycleID(CmbBillingCycle.Text)
Call GetDepositeBillingCycleID(CmbDepositBillingCycle.Text)
cmdExportDetails.Enabled = True

CmbState = mCStateName
mStateID = mCStateNameID
CmbLocation = mCLocationName
mLocationID = mCLocationID
CmbCommodity = mCCommodityName
mCommodityID = mCCommodityID
CmbClient = mCClientName
mClientIDRow = mCClientID
Check4.Value = IIf(mCBags = True, vbChecked, vbUnchecked)
CmbInsurance = mCInsuranceBy
mInsuranceByID = mCInsuranceID
TxtTentativeDuration = mCTentativeDuration
TxtLockInPeriod = mCLockInPeriod
TxtQuantum = mCQuantum
CmbBillingCycle = mCBillingCycle
mBillingCycleID = mCBillingCycleID
CmbDepositBillingCycle = mCDBillingCycle
mDepositBillingCycleID = mCDBillingCycleID
mFlag = True

''-- 14 Feb 2018 Ullhas Patil
If LCase(Gen_UserRole) = "power" Or mApprovedRights = True Then
    TxtContractType = "": TxtContractType.Visible = False
    CmbContractType = "Normal": CmbContractType.Visible = True
    mContractTypeFlag = "N"
    CmbContractType = mCContractTypeText
    mContractTypeFlag = mCContractTypeFlag
Else
    TxtContractType = "Normal": TxtContractType.Visible = True
    CmbContractType = "Normal": CmbContractType.Visible = False
    mContractTypeFlag = "N"
End If
''----------------------------

CmbClient.SetFocus

End Sub


Private Sub CmdDeleteFrmList_Click()
If CmdAdd2List.Caption <> "Add to List" Then I = MSHFlexGrid1.RowSel
If MSHFlexGrid1.TextMatrix(I, 0) = "" Then Exit Sub
If MSHFlexGrid1.TextMatrix(I, 0) <> "" Then
   For RO = I To MSHFlexGrid1.Rows - 2
       For C = 0 To MSHFlexGrid1.Cols - 1
           MSHFlexGrid1.TextMatrix(RO, C) = MSHFlexGrid1.TextMatrix(RO + 1, C)
       Next
   Next
       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows - 1
End If
Call ClearFrameControl
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Txn_DepositContract.xls")
Call Xls_Detail_Rpt
End Sub

Private Sub cmdExportDetails_Click()
Dim mR As Double
If SaveCmd.Enabled = False Then Exit Sub
If Check4.Value = vbUnchecked Then  ''-- Quantity Wise
    With MSHFlexGrid1
         If .Rows > 1 And .TextMatrix(1, 0) <> "" Then
            Gen_mTimeStamp = GetTimeStamp
            Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "_Data_Txn_DepositContract.xls")
            Call Xls_GridDetail(MSHFlexGrid1)
         End If
    End With
Else ''-- Bag Wise
    With MSHFlexGrid4
         If .Rows > 1 And .TextMatrix(1, 0) <> "" Then
            Gen_mTimeStamp = GetTimeStamp
            Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "_Data_Txn_DepositContract.xls")
            Call Xls_GridDetail(MSHFlexGrid4)
         End If
    End With
End If
End Sub

Private Sub CmdGetFile_Click()
With cdialog
     .InitDir = App.Path
     .DialogTitle = "Select File"
     .FilterIndex = 1
     .Filter = "*.*"
     .ShowOpen
End With
If cdialog.CancelError Then
   lblAttachment.Caption = ""
Else
   strFilename = cdialog.FileName
   If strFilename = "" Then Exit Sub
   lblAttachment.Refresh
   lblAttachment.Caption = strFilename 'GetFileFromPath(strFilename)
End If
End Sub


Private Sub CmdRefreshRate_Click()
Dim mGList As Variant
Dim mGCount As Variant

''-- for Concurrency Check Only for Update
tmp = " Select * from Txn_DepositContract Where DepConID = " & Val(TxtDepositContractID) & " And CONVERT(varchar(30),Concurrency,120) > '" & Format(RsTxn_DepositContract!Concurrency, "YYYY-MM-DD hh:mm:ss") & "' "
Call TmpTable
If TmpRs.EOF = False Then
    MsgBox "Sorry, System not allow to update the data because same data in updated by another user." & vbNewLine & vbNewLine & "Please reopen your form to get update information and process.", vbInformation + vbOKOnly
    Exit Sub
End If
''--------------------------------
mOwnGodownSqrFeetRate = GetOwnGodownPerSqtFeet(mLocationID)

If CmbBillingCycle.Text = "" Then
   MsgBox "Blank Billing Cycle Not Allowed.", vbOKOnly + vbExclamation, App.Title
   CmbBillingCycle.SetFocus
   Exit Sub
End If
If CmbDepositBillingCycle.Text = "" Then
   MsgBox "Blank Deposit Billing Cycle Not Allowed.", vbOKOnly + vbExclamation, App.Title
   CmbDepositBillingCycle.SetFocus
   Exit Sub
End If

If Check4.Value = 0 Then
   tmp = " Select MC.CMPName,MG.GodownNo,MG.Address,MU.UnitType,sum(M1.RentAmount) 'Rent',MG.Area,"
   tmp = tmp & " MG.GodownID,cast('' As Varchar) DepConDID,MU.UnitTypeID,Isnull(MG.ApprovalLimit,0)ApprovalLimit "
   tmp = tmp & " From Mst_Godown MG"
   tmp = tmp & " Inner Join Mst_CMP MC On MG.CMPID=MC.CMPID"
   tmp = tmp & " Inner Join Mst_UnitType MU on MG.UnitTypeID = MU.UnitTypeID"
   tmp = tmp & " Inner Join Txn_GodownLessorAgreeDetail M1 On MG.GodownID=M1.GodownID"
   tmp = tmp & " inner Join Txn_LessorAgreement TLA on M1.LAID = TLA.LAID And TLA.Flag = 'A'"
   tmp = tmp & " Where MG.CloseDate is null  And MG.StartDate<='" & Format(TxtEndDate, "YYYY-MM-DD") & "' And ( '" & TxtStartDate.Value & "' <= TLA.ExpiryDate And '" & TxtEndDate.Value & "' >= TLA.StartDate )"
   tmp = tmp & " And MC.LocationID = " & mLocationID & " And  MU.UnitTypeID = 1"
   tmp = tmp & " AND MG.GodownID IN ( Select GodownId from Txn_DepositContractGodownDetail Where DepConID = " & Val(TxtDepositContractID) & " ) "

''---- 05 Feb 2018 added for Last Rate
    tmp = tmp & " And Convert(Varchar,MG.GodownID)+''+Convert(Varchar,TLA.ExpiryDate,112) In "
    tmp = tmp & " ( "
    tmp = tmp & "   Select Convert(Varchar,AA.GodownID)+''+Convert(Varchar,Max(BB.ExpiryDate),112) "
    tmp = tmp & "   From Txn_GodownLessorAgreeDetail AA Inner Join Txn_LessorAgreement BB On AA.LAID = BB.LAID "
    tmp = tmp & "   Where Convert(Varchar,AA.LAID)+''+Convert(Varchar,AA.GodownID) "
    tmp = tmp & "   In ( Select Convert(Varchar,E.LAID)+''+Convert(Varchar,D.GodownID) LAIDGodownKey "
    tmp = tmp & "   From Mst_Godown A Inner Join Mst_CMP B On A.CMPID=B.CMPID "
    tmp = tmp & "   Inner Join Mst_UnitType C on A.UnitTypeID = C.UnitTypeID "
    tmp = tmp & "   Inner Join Txn_GodownLessorAgreeDetail D On A.GodownID=D.GodownID "
    tmp = tmp & "   Inner Join Txn_LessorAgreement E on D.LAID = E.LAID And E.Flag = 'A' "
    tmp = tmp & "   Where A.CloseDate is null And A.StartDate<='" & Format(TxtEndDate, "YYYY-MM-DD") & "' And ( '" & TxtStartDate.Value & "' <= E.ExpiryDate And '" & TxtEndDate.Value & "' >= E.StartDate ) "
    tmp = tmp & "   And B.LocationID = " & mLocationID & " And C.UnitTypeID = 1 "
    tmp = tmp & "   AND A.GodownID IN (Select GodownId from Txn_DepositContractGodownDetail Where DepConID = " & Val(TxtDepositContractID) & " ) "
    tmp = tmp & "     ) Group By AA.GodownID "
    tmp = tmp & " ) "
''-----------

   tmp = tmp & " Group by MC.CMPName,MG.GodownNo,MG.Address,MU.UnitType,MG.Area,MG.GodownID,MU.UnitTypeID,Isnull(MG.ApprovalLimit,0) "
   tmp = tmp & " Union All"
   tmp = tmp & " Select MC.CMPName,MG.GodownNo,MG.Address,MU.UnitType,0 'Rent',MG.Area,MG.GodownID,cast('' As Varchar) DepConDID,MU.UnitTypeID,Isnull(MG.ApprovalLimit,0)ApprovalLimit "
   tmp = tmp & " From Mst_Godown MG"
   tmp = tmp & " Inner Join Mst_CMP MC On MG.CMPID=MC.CMPID"
   tmp = tmp & " Inner Join Mst_UnitType MU on MG.UnitTypeID = MU.UnitTypeID"
   tmp = tmp & " Where MG.CloseDate is null  And MG.StartDate<='" & Format(TxtEndDate, "YYYY-MM-DD") & "' And '" & Format(TxtEndDate, "YYYY-MM-DD") & "' <= MG.AgreementEndDate "
   tmp = tmp & " And MC.LocationID = " & mLocationID & " And  MU.UnitTypeID <> 1"
   tmp = tmp & " AND MG.GodownID IN ( Select GodownId from Txn_DepositContractGodownDetail Where DepConID = " & Val(TxtDepositContractID) & " ) "
   tmp = tmp & " Group by MC.CMPName,MG.GodownNo,MG.Address,MU.UnitType,MG.Area,MG.GodownID,MU.UnitTypeID,Isnull(MG.ApprovalLimit,0) "
   Call Tmp2Table
'mCurrentBaseRentFactor,mCalculatedBaseRent
   With MSHFlexGrid1
    If TmpRs2.RecordCount > 0 Then
       For I = 1 To .Rows - 1
           TmpRs2.MoveFirst
           TmpRs2.Find " GodownID = " & Val(.TextMatrix(I, 9))
           If TmpRs2.EOF = True Then GoTo x
           mCalculatedBaseRent = (Val(.TextMatrix(I, 5)) / Val(.TextMatrix(I, 6)))

           If TmpRs2!UnitTypeID = 1 Or TmpRs2!UnitTypeID = 7 Then
              If TmpRs2!UnitTypeID = 1 Then
                 If mCalculatedBaseRent < mCurrentBaseRentFactor Then
                    mCalculatedBaseRent = mCurrentBaseRentFactor
                 End If
              End If
              If LCase(mFumigationFlag) = "y" Then
              '   .TextMatrix(I, 7) = Format((((Val(.TextMatrix(I, 5)) / Val(.TextMatrix(I, 6))) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorY, "#0.0000")
              '   .TextMatrix(I, 11) = Format((((Val(.TextMatrix(I, 5)) / Val(.TextMatrix(I, 6))) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorY, "#0.0000")
              '   .TextMatrix(I, 16) = Format((((Val(.TextMatrix(I, 5)) / Val(.TextMatrix(I, 6))) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorY, "#0.0000")
              
                 .TextMatrix(I, 7) = Format((((Val(mCalculatedBaseRent)) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorY, "#0.0000")
                 .TextMatrix(I, 11) = Format((((Val(mCalculatedBaseRent)) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorY, "#0.0000")
                 .TextMatrix(I, 16) = Format((((Val(mCalculatedBaseRent)) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorY, "#0.0000")
              Else
                 
              '   .TextMatrix(I, 7) = Format((((Val(.TextMatrix(I, 5)) / Val(.TextMatrix(I, 6))) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorN, "#0.0000")
              '   .TextMatrix(I, 11) = Format((((Val(.TextMatrix(I, 5)) / Val(.TextMatrix(I, 6))) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorN, "#0.0000")
              '   .TextMatrix(I, 16) = Format((((Val(.TextMatrix(I, 5)) / Val(.TextMatrix(I, 6))) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorN, "#0.0000")
              
                 .TextMatrix(I, 7) = Format((((Val(mCalculatedBaseRent)) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorN, "#0.0000")
                 .TextMatrix(I, 11) = Format((((Val(mCalculatedBaseRent)) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorN, "#0.0000")
                 .TextMatrix(I, 16) = Format((((Val(mCalculatedBaseRent)) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorN, "#0.0000")
              End If
           Else
           
           
'              .TextMatrix(I, 7) = Val(mCurrentBaseRentFactor) ' "0.0000"
'              .TextMatrix(I, 11) = Val(mCurrentBaseRentFactor) '"0.0000"
'              .TextMatrix(I, 16) = Val(mCurrentBaseRentFactor) '"0.0000"
                 
              mCalculatedBaseRent = mCurrentBaseRentFactor
              If LCase(mFumigationFlag) = "y" Then
                 .TextMatrix(I, 7) = Format((((Val(mCalculatedBaseRent)) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorY, "#0.0000")
                 .TextMatrix(I, 11) = Format((((Val(mCalculatedBaseRent)) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorY, "#0.0000")
                 .TextMatrix(I, 16) = Format((((Val(mCalculatedBaseRent)) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorY, "#0.0000")
              Else
                 .TextMatrix(I, 7) = Format((((Val(mCalculatedBaseRent)) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorN, "#0.0000")
                 .TextMatrix(I, 11) = Format((((Val(mCalculatedBaseRent)) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorN, "#0.0000")
                 .TextMatrix(I, 16) = Format((((Val(mCalculatedBaseRent)) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorN, "#0.0000")
              End If
           End If
           .TextMatrix(I, 13) = Val(.TextMatrix(I, 7))
           .TextMatrix(I, 14) = Val(.TextMatrix(I, 7))
           .TextMatrix(I, 17) = Format(IIf(Val(TmpRs2!ApprovalLimit) = 0, mCMPMtoAMRMApprovalLimit, TmpRs2!ApprovalLimit), "#0.00")

            '''-- Billing Cycle / Deposite Biling Cycle Wise Rate Calculation
            .TextMatrix(I, 7) = Format(CalculateRate(.TextMatrix(I, 13), mBillingRatio, CmbBillingCycle), "#0.0000")
            mBillingOnePercentage = Format(IIf(Val(.TextMatrix(I, 7)) > 0, Val(.TextMatrix(I, 7)) / 100, 0), "#0.0000")
            '.TextMatrix(i, 8) = Format(Val(.TextMatrix(i, 7)) + (mBillingOnePercentage * .TextMatrix(i, 17)), "#0.0000")

            .TextMatrix(I, 11) = Format(CalculateRate(.TextMatrix(I, 14), mDepositBillingRatio, CmbDepositBillingCycle), "#0.0000")
            mDepositOnePercentage = Format(IIf(Val(.TextMatrix(I, 11)) > 0, Val(.TextMatrix(I, 11)) / 100, 0), "#0.0000")
            '.TextMatrix(i, 12) = Format(Val(.TextMatrix(i, 11)) + (mDepositOnePercentage * .TextMatrix(i, 17)), "#0.0000")

            Call CheckRateByColor(CInt(I))

x:       Next I
    Else
       MsgBox "No Godown Found !!!"
       Exit Sub
    End If

   End With
Else

   tmp = " Select MC.CMPName,MG.GodownNo,MG.Address,MU.UnitType,sum(M1.RentAmount) 'Rent',MG.Area,MG.GodownID,cast('' As Varchar) DepConDID,MU.UnitTypeID,Isnull(MG.ApprovalLimit,0)ApprovalLimit"
   tmp = tmp & " From Mst_Godown MG"
   tmp = tmp & " Inner Join Mst_CMP MC On MG.CMPID=MC.CMPID"
   tmp = tmp & " Inner Join Mst_UnitType MU on MG.UnitTypeID = MU.UnitTypeID"
   tmp = tmp & " Inner Join Txn_GodownLessorAgreeDetail M1 On MG.GodownID=M1.GodownID"
   tmp = tmp & " inner Join Txn_LessorAgreement TLA on M1.LAID = TLA.LAID And TLA.Flag = 'A'"
   tmp = tmp & " Where MG.CloseDate is null  And MG.StartDate<='" & Format(TxtEndDate, "YYYY-MM-DD") & "' And ( '" & TxtStartDate.Value & "' <= TLA.ExpiryDate And '" & TxtEndDate.Value & "' >= TLA.StartDate )"
   tmp = tmp & " And MC.LocationID = " & mLocationID & " And  MU.UnitTypeID = 1"
   tmp = tmp & " AND MG.GodownID IN ( Select GodownId from Txn_DepositContractGodownDetail Where DepConID = " & Val(TxtDepositContractID) & " ) "
   
''---- 05 Feb 2018 added for Last Rate
    tmp = tmp & " And Convert(Varchar,MG.GodownID)+''+Convert(Varchar,TLA.ExpiryDate,112) In "
    tmp = tmp & " ( "
    tmp = tmp & "   Select Convert(Varchar,AA.GodownID)+''+Convert(Varchar,Max(BB.ExpiryDate),112) "
    tmp = tmp & "   From Txn_GodownLessorAgreeDetail AA Inner Join Txn_LessorAgreement BB On AA.LAID = BB.LAID "
    tmp = tmp & "   Where Convert(Varchar,AA.LAID)+''+Convert(Varchar,AA.GodownID) "
    tmp = tmp & "   In ( Select Convert(Varchar,E.LAID)+''+Convert(Varchar,D.GodownID) LAIDGodownKey "
    tmp = tmp & "   From Mst_Godown A Inner Join Mst_CMP B On A.CMPID=B.CMPID "
    tmp = tmp & "   Inner Join Mst_UnitType C on A.UnitTypeID = C.UnitTypeID "
    tmp = tmp & "   Inner Join Txn_GodownLessorAgreeDetail D On A.GodownID=D.GodownID "
    tmp = tmp & "   Inner Join Txn_LessorAgreement E on D.LAID = E.LAID And E.Flag = 'A' "
    tmp = tmp & "   Where A.CloseDate is null And A.StartDate<='" & Format(TxtEndDate, "YYYY-MM-DD") & "' And ( '" & TxtStartDate.Value & "' <= E.ExpiryDate And '" & TxtEndDate.Value & "' >= E.StartDate ) "
    tmp = tmp & "   And B.LocationID = " & mLocationID & " And C.UnitTypeID = 1 "
    tmp = tmp & "   AND A.GodownID IN (Select GodownId from Txn_DepositContractGodownDetail Where DepConID = " & Val(TxtDepositContractID) & " ) "
    tmp = tmp & "     ) Group By AA.GodownID "
    tmp = tmp & " ) "
''-----------
   
   tmp = tmp & " Group by MC.CMPName,MG.GodownNo,MG.Address,MU.UnitType,MG.Area,MG.GodownID,MU.UnitTypeID,Isnull(MG.ApprovalLimit,0) "
   tmp = tmp & " Union All"
   tmp = tmp & " Select MC.CMPName,MG.GodownNo,MG.Address,MU.UnitType,0 'Rent',MG.Area,MG.GodownID,cast('' As Varchar) DepConDID,MU.UnitTypeID,Isnull(MG.ApprovalLimit,0)ApprovalLimit"
   tmp = tmp & " From Mst_Godown MG"
   tmp = tmp & " Inner Join Mst_CMP MC On MG.CMPID=MC.CMPID"
   tmp = tmp & " Inner Join Mst_UnitType MU on MG.UnitTypeID = MU.UnitTypeID"
   tmp = tmp & " Where MG.CloseDate is null  And MG.StartDate<='" & Format(TxtEndDate, "YYYY-MM-DD") & "' And '" & Format(TxtEndDate, "YYYY-MM-DD") & "' <= MG.AgreementEndDate "
   tmp = tmp & " And MC.LocationID = " & mLocationID & " And  MU.UnitTypeID <> 1"
   tmp = tmp & " AND MG.GodownID IN ( Select GodownId from Txn_DepositContractGodownDetail Where DepConID = " & Val(TxtDepositContractID) & " ) "
   tmp = tmp & " Group by MC.CMPName,MG.GodownNo,MG.Address,MU.UnitType,MG.Area,MG.GodownID,MU.UnitTypeID,Isnull(MG.ApprovalLimit,0) "
   Call Tmp2Table
   With MSHFlexGrid4
    If TmpRs2.RecordCount > 0 Then
       For I = 1 To .Rows - 1
           TmpRs2.MoveFirst
           TmpRs2.Find " GodownID = " & Val(.TextMatrix(I, 11))
           If TmpRs2.EOF = True Then GoTo y
'Val(mCalculatedBaseRent)
           mCalculatedBaseRent = (Val(.TextMatrix(I, 5)) / Val(.TextMatrix(I, 6)))
           
           If TmpRs2!UnitTypeID = 1 Or TmpRs2!UnitTypeID = 7 Then
              If TmpRs2!UnitTypeID = 1 Then
                 If mCalculatedBaseRent < mCurrentBaseRentFactor Then
                    mCalculatedBaseRent = mCurrentBaseRentFactor
                 End If
              End If
              If LCase(mFumigationFlag) = "y" Then
'                 .TextMatrix(I, 9) = Format((((Val(.TextMatrix(I, 5)) / Val(.TextMatrix(I, 6))) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorY, "#0.0000")
'                 .TextMatrix(I, 13) = Format((((Val(.TextMatrix(I, 5)) / Val(.TextMatrix(I, 6))) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorY, "#0.0000")
'                 .TextMatrix(I, 18) = Format((((Val(.TextMatrix(I, 5)) / Val(.TextMatrix(I, 6))) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorY, "#0.0000")
              
                 .TextMatrix(I, 9) = Format((((Val(mCalculatedBaseRent)) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorY, "#0.0000")
                 .TextMatrix(I, 13) = Format((((Val(mCalculatedBaseRent)) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorY, "#0.0000")
                 .TextMatrix(I, 18) = Format((((Val(mCalculatedBaseRent)) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorY, "#0.0000")
              
              
              Else
'                 .TextMatrix(I, 9) = Format((((Val(.TextMatrix(I, 5)) / Val(.TextMatrix(I, 6))) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorN, "#0.0000")
'                 .TextMatrix(I, 13) = Format((((Val(.TextMatrix(I, 5)) / Val(.TextMatrix(I, 6))) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorN, "#0.0000")
'                 .TextMatrix(I, 18) = Format((((Val(.TextMatrix(I, 5)) / Val(.TextMatrix(I, 6))) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorN, "#0.0000")
              
                 .TextMatrix(I, 9) = Format((((Val(mCalculatedBaseRent)) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorN, "#0.0000")
                 .TextMatrix(I, 13) = Format((((Val(mCalculatedBaseRent)) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorN, "#0.0000")
                 .TextMatrix(I, 18) = Format((((Val(mCalculatedBaseRent)) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorN, "#0.0000")
              
              End If
           Else
'              .TextMatrix(I, 9) = Val(mCurrentBaseRentFactor) '"0.0000"
'              .TextMatrix(I, 13) = Val(mCurrentBaseRentFactor) '"0.0000"
'              .TextMatrix(I, 18) = Val(mCurrentBaseRentFactor) '"0.0000"
           
              mCalculatedBaseRent = mCurrentBaseRentFactor
              If LCase(mFumigationFlag) = "y" Then
                 .TextMatrix(I, 9) = Format((((Val(mCalculatedBaseRent)) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorY, "#0.0000")
                 .TextMatrix(I, 13) = Format((((Val(mCalculatedBaseRent)) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorY, "#0.0000")
                 .TextMatrix(I, 18) = Format((((Val(mCalculatedBaseRent)) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorY, "#0.0000")
              Else
                 .TextMatrix(I, 9) = Format((((Val(mCalculatedBaseRent)) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorN, "#0.0000")
                 .TextMatrix(I, 13) = Format((((Val(mCalculatedBaseRent)) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorN, "#0.0000")
                 .TextMatrix(I, 18) = Format((((Val(mCalculatedBaseRent)) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorN, "#0.0000")
              End If
           
           
           End If

           .TextMatrix(I, 15) = Val(.TextMatrix(I, 9))
           .TextMatrix(I, 16) = Val(.TextMatrix(I, 9))
           .TextMatrix(I, 19) = Format(IIf(Val(TmpRs2!ApprovalLimit) = 0, mCMPMtoAMRMApprovalLimit, TmpRs2!ApprovalLimit), "#0.00")

            '''-- Billing Cycle / Deposite Biling Cycle Wise Rate Calculation
            .TextMatrix(I, 9) = Format(CalculateRate(.TextMatrix(I, 15), mBillingRatio, CmbBillingCycle) * Val(.TextMatrix(I, 8)) / 1000, "#0.0000")
            mBillingOnePercentage = Format(Val(.TextMatrix(I, 9)) / 100, "#0.0000")
            '.TextMatrix(i, 10) = Format(Val(.TextMatrix(i, 9)) + (mBillingOnePercentage * .TextMatrix(i, 19)), "#0.0000")

            .TextMatrix(I, 13) = Format(CalculateRate(.TextMatrix(I, 16), mDepositBillingRatio, CmbDepositBillingCycle) * Val(.TextMatrix(I, 8)) / 1000, "#0.0000")
            mDepositOnePercentage = Format(Val(.TextMatrix(I, 13)) / 100, "#0.0000")
            '.TextMatrix(i, 14) = Format(Val(.TextMatrix(i, 13)) + (mDepositOnePercentage * .TextMatrix(i, 19)), "#0.0000")

            Call CheckBagRateByColor(CInt(I))

y:       Next I
    Else
       MsgBox "No Godown Found !!!"
       Exit Sub
    End If

   End With
End If

End Sub

Private Sub CmdRemoveAttachment_Click()
With MSHFlexGrid5
    I = .RowSel
    If I > 0 Then
       If lblAttachment.Caption <> "" Then
        'If .TextMatrix(I, 4) = "" And .TextMatrix(I, 1) <> "" Then
            For RO = I To .Rows - 2
                For C = 0 To .Cols - 1
                    .TextMatrix(RO, C) = .TextMatrix(RO + 1, C)
                Next
            Next
            .Rows = .Rows - 1
        'End If
       End If
    End If
    Call ClearFrameAttachment
End With
End Sub

Private Sub CmdSearch_Click()
Dim wc As String
Call Grid_Head2
wc = ""
Label30.Caption = "-"
If STo.Value < SFrom.Value Then
   MsgBox "To date must be always greater than from date!!!"
   STo.SetFocus
   Exit Sub
End If

If Gen_DBType = "MDB" Then
   wc = " Where TDC.Date Between #" & Format(SFrom, Gen_DatFormat) & "# And #" & Format(STo, Gen_DatFormat) & "#) "
Else
   wc = " Where TDC.Date Between '" & Format(SFrom, Gen_DatFormat) & "' And '" & Format(STo, Gen_DatFormat) & "' "
End If

If CmbSState.Text <> "" Then
   wc = GenWc(wc, mSStateID, "MS.StateID", "N", "N")
Else
   If Gen_WcState <> "" Then
      tmp = Replace(Gen_WcState, "StateID", "MS.StateID")
'      wc = GenWc(wc, tmp, "", "I", "N")
      wc = wc & " And " & tmp
   End If
End If

If CmbSLocation.Text <> "" Then
   wc = GenWc(wc, mSLocationID, "ML.LocationID", "N", "N")
Else
   If Gen_WcLocation <> "" Then
      tmp = Replace(Gen_WcLocation, "LocationID", "ML.LocationID")
'      wc = GenWc(wc, tmp, "", "I", "N")
      wc = wc & " And " & tmp
   End If
End If

If CmbSCommodity.Text <> "" Then
   wc = GenWc(wc, mSCommodityID, "Mco.CommodityID", "N", "N")
End If

If CmbSClient.Text <> "" Then
   wc = GenWc(wc, mSClientIDRow, "MC.ClientID", "N", "N")
End If

If TxtSDepositContractID.Text <> "" Then
   wc = GenWc(wc, Val(TxtSDepositContractID), "TDC.DepConID", "N", "N")
End If
 
If CmbSForwardedTo.Text <> "" Then
   wc = GenWc(wc, mSForwardID, "ME.EmployeeID", "N", "N")
End If

If CmbSStatus.Text <> "" Then
   If wc = "" Then
      wc = " TDC.ApprovalFlag in (" & mSStatus & ")"
   Else
      wc = wc & " And TDC.ApprovalFlag in (" & mSStatus & ")"
   End If
End If

If CmbSVerified.Text <> "" Then
   If wc = "" Then
      wc = " Isnull(TDC.Verified,'N') = '" & Left(CmbSVerified.Text, 1) & "'"
   Else
      wc = wc & " And Isnull(TDC.Verified,'N') = '" & Left(CmbSVerified.Text, 1) & "'"
   End If
End If

''--- 14 Feb 2018
If CmbSContractType.Text <> "" Then
   If wc = "" Then
      wc = " Isnull(TDC.ContractType,'N') = '" & mSContractTypeFlag & "'"
   Else
      wc = wc & " And Isnull(TDC.ContractType,'N') = '" & mSContractTypeFlag & "'"
   End If
End If
If CmbSType.Text <> "" Then
   If wc = "" Then
      wc = " TDC.BagSizeFlag = '" & mSType & "'"
   Else
      wc = wc & " And TDC.BagSizeFlag = '" & mSType & "'"
   End If
End If
''--------------------

tmp = "Select TDC.DepConID,TDC.Date,MS.StateName,ML.LocationName,Mco.Commodity,MC.ClientName,MC.ClientID,MC.ClientIDRow,"
tmp = tmp & " TDC.QuantumMT,MIB.InsuranceByName,TDC.StartDate,TDC.EndDate,"
tmp = tmp & " MBC.BillingCycle As BillingCycle, MBCDep.BillingCycle As BillingCycleDep,TDC.RatePer,TDC.TentativeDur,TDC.LockPeriod,ME.EmployeeName+'~'+Me.EmployeeNo As ForwardedTo,TDC.ApprovalFlag,TDC.ApproveOn,MEE.EmployeeName+'~'+MEE.EmployeeNo As ApproveBy,TDC.ClosedOn,TDC.ClosedBy,MEEE.EmployeeName+'~'+MEEE.EmployeeNo As CloseBy,TDC.BagSizeFlag,TDC.Verified,TDC.VerifiedDate,MEEEE.EmployeeName+'~'+MEEEE.EmployeeNo as VerifiedBy,TDC.ContractType "
tmp = tmp & " From Txn_DepositContract TDC"
''tmp = tmp & " Inner Join Mst_State MS on TDC.StateID=MS.StateID"
tmp = tmp & " Inner Join MST_Location ML on TDC.LocationID=ML.LocationID"
tmp = tmp & " Inner Join Mst_State MS on ML.StateID=MS.StateID"
tmp = tmp & " Inner Join MST_Commodity Mco on TDC.CommodityID=Mco.CommodityID"
tmp = tmp & " Inner Join MST_Client MC on TDC.ClientID=MC.ClientID"
tmp = tmp & " Inner Join MST_BillingCycle MBC on TDC.BillingCycleID=MBC.BillingCycleID"
tmp = tmp & " Inner Join MST_BillingCycle MBCDep on TDC.DepBillingCycleID=MBCDep.BillingCycleID"
tmp = tmp & " Inner Join Mst_InsuranceBy MIB on TDC.InsuranceByID=MIB.InsuranceByID"
tmp = tmp & " Left Join MST_Employee ME on TDC.EmployeeID=ME.EmployeeID"
tmp = tmp & " Left Join MST_Employee MEE on TDC.ApproveBy=MEE.EmployeeNo"
tmp = tmp & " Left Join MST_Employee MEEE on TDC.ClosedBy=MEEE.EmployeeNo"
tmp = tmp & " Left Join MST_Employee MEEEE on TDC.VerifiedBy=MEEEE.EmployeeNo"

If wc <> "" Then
   tmp = tmp & wc
End If
tmp = tmp & " ORDER BY TDC.DepConID,TDC.Date "
Call TmpTable

MSHFlexGrid2.Rows = TmpRs.RecordCount + 2
For I = 1 To TmpRs.RecordCount
    With MSHFlexGrid2
        .TextMatrix(I, 0) = IIf(IsNull(TmpRs!DepConID), "", TmpRs!DepConID)                                                         ''-- Deposit Contract ID
        .TextMatrix(I, 1) = IIf(IsNull(TmpRs!Date), "", Format(TmpRs!Date, "DD-MMM-YYYY"))                                          ''-- Date
        .TextMatrix(I, 2) = IIf(IsNull(TmpRs!ContractType), "Normal", IIf(LCase(TmpRs!ContractType) = "n", "Normal", "Revision"))   ''-- ContractType   ''-- 14 Feb 2018
        .TextMatrix(I, 3) = IIf(IsNull(TmpRs!StateName), "", TmpRs!StateName)                                                       ''-- State
        .TextMatrix(I, 4) = IIf(IsNull(TmpRs!Locationname), "", TmpRs!Locationname)                                                 ''-- Location
        .TextMatrix(I, 5) = IIf(IsNull(TmpRs!Commodity), "", TmpRs!Commodity)                                                       ''-- Commodity
        .TextMatrix(I, 6) = IIf(IsNull(TmpRs!ClientName), "", TmpRs!ClientName & "#" & TmpRs!ClientIDRow & "#" & TmpRs!ClientID)    ''-- Client
        .TextMatrix(I, 7) = IIf(IsNull(TmpRs!InsuranceByName), "", TmpRs!InsuranceByName)                                           ''-- Insurance By
        .TextMatrix(I, 8) = IIf(IsNull(TmpRs!BagSizeFlag), "", IIf(LCase(TmpRs!BagSizeFlag) = "y", "Bags", "Quantity"))             ''-- Qty/Bags
        .TextMatrix(I, 9) = IIf(IsNull(TmpRs!StartDate), "", Format(TmpRs!StartDate, "DD-MMM-YYYY"))                                ''-- Deposit Contract Start Date
        .TextMatrix(I, 10) = IIf(IsNull(TmpRs!EndDate), "", Format(TmpRs!EndDate, "DD-MMM-YYYY"))                                   ''-- Deposit Contract End Date
        .TextMatrix(I, 11) = IIf(IsNull(TmpRs!BillingCycle), "", TmpRs!BillingCycle)                                                ''-- Billing Cycle
        .TextMatrix(I, 12) = IIf(IsNull(TmpRs!BillingCycleDep), "", TmpRs!BillingCycleDep)                                          ''-- Deposit/ Withdrawal Billing Cycle
        .TextMatrix(I, 13) = IIf(IsNull(TmpRs!RatePer), "", IIf(UCase(TmpRs!RatePer) = "I", "Included", "Excluded"))                ''-- Service Tax (GST)
        .TextMatrix(I, 14) = IIf(IsNull(TmpRs!ForwardedTo), "", TmpRs!ForwardedTo)                                                  ''-- Forwarded To
        .TextMatrix(I, 15) = IIf(IsNull(TmpRs!ApprovalFlag), "", UCase(TmpRs!ApprovalFlag))                                         ''-- Status
            If .TextMatrix(I, 15) = "A" Then .TextMatrix(I, 15) = "Approved"
            If .TextMatrix(I, 15) = "S" Then .TextMatrix(I, 15) = "System Approved"
            If .TextMatrix(I, 15) = "F" Then .TextMatrix(I, 15) = "Forwarded"
            If .TextMatrix(I, 15) = "R" Then .TextMatrix(I, 15) = "Rejected"
            If .TextMatrix(I, 15) = "C" Then .TextMatrix(I, 15) = "Closed"
        .TextMatrix(I, 16) = IIf(IsNull(TmpRs!TentativeDur), "", TmpRs!TentativeDur)                                                ''-- Tentative Duration of Storage (in months)
        .TextMatrix(I, 17) = IIf(IsNull(TmpRs!LockPeriod), "", TmpRs!LockPeriod)                                                    ''-- Lock-in Period (in months)
        .TextMatrix(I, 18) = IIf(IsNull(TmpRs!QuantumMT), "", TmpRs!QuantumMT)                                                      ''-- Indicative Quantum
        
        If LCase(TmpRs!ApprovalFlag) = "a" Or LCase(TmpRs!ApprovalFlag) = "c" Then
            .TextMatrix(I, 19) = IIf(IsNull(TmpRs!ApproveOn), "", Format(TmpRs!ApproveOn, "DD-MMM-YYYY hh:mm:ss"))                  ''-- Approveon
            .TextMatrix(I, 20) = IIf(IsNull(TmpRs!ApproveBy), "", TmpRs!ApproveBy)                                                  ''-- ApproveBy
        End If
        ''-- 02 Feb 2018
        .TextMatrix(I, 21) = IIf(IsNull(TmpRs!Verified), "No", IIf(LCase(TmpRs!Verified) = "y", "Yes", "No"))                       ''-- Verified
        .TextMatrix(I, 22) = IIf(IsNull(TmpRs!VerifiedDate), "", Format(TmpRs!VerifiedDate, "DD-MMM-YYYY hh:mm:ss"))                ''-- Approveon
        .TextMatrix(I, 23) = IIf(IsNull(TmpRs!VerifiedBy), "", TmpRs!VerifiedBy)                                                    ''-- ApproveBy
        ''-----------------
        ''-- Rights applyed on 05 Jan 2018
        ''If LCase(Gen_UserRole) = "head cmg" Or LCase(Gen_UserRole) = "superadmin" Or LCase(Gen_UserRole) = "power" Then
        If LCase(Gen_UserRole) = "power" Or mApprovedRights = True Then
            .TextMatrix(I, 24) = IIf(IsNull(TmpRs!ClosedOn), "", Format(TmpRs!ClosedOn, "DD-MMM-YYYY hh:mm:ss"))                    ''-- ClosedNo
            .TextMatrix(I, 25) = IIf(IsNull(TmpRs!CloseBy), "", TmpRs!CloseBy)                                                      ''-- ClosedBy
        End If
        
    End With
    TmpRs.MoveNext
Next
Label30.Caption = TmpRs.RecordCount
End Sub


Private Sub CmdVerified_Click()
    ''-- 02 Feb 2018 Verification
    ''-- for Concurrency Check Only for Update
    tmp = " Select * from Txn_DepositContract Where DepConID = " & Val(TxtDepositContractID) & " And CONVERT(varchar(30),Concurrency,120) > '" & Format(RsTxn_DepositContract!Concurrency, "YYYY-MM-DD hh:mm:ss") & "' "
    Call TmpTable
    If TmpRs.EOF = False Then
        MsgBox "Sorry, System not allow to update the data because same data in updated by another user." & vbNewLine & vbNewLine & "Please reopen your form to get update information and process.", vbInformation + vbOKOnly
        Exit Sub
    End If
    ''--------------------------------
    If LCase(mStatus) <> "a" And LCase(mStatus) <> "s" Then
        MsgBox " Deposit Contract status not 'Approved'/'System Approved' to Verifiy."
        Exit Sub
    End If
    Frame6.Width = SSTab1.Width
    Frame6.Height = SSTab1.Height + ButtonFrm.Height + 100
    Frame6.Left = SSTab1.Left
    Frame6.Top = SSTab1.Top - 50
    Frame6.Visible = True
    TxtVerifiedRemark.SetFocus

End Sub

Private Sub CmdVerifiedCancel_Click()
   Frame6.Visible = False
End Sub

Private Sub CmdVerifiedSave_Click()
If Trim(TxtVerifiedRemark) = "" Then
    MsgBox "Verified remark can not blank."
    Exit Sub
End If

Select Case MsgBox("Do you want to Verified Deposit Contract.", vbYesNo)
    Case vbNo
        Exit Sub
    Case vbYes
    
        tmp = " UPDATE Txn_DepositContract SET Verified = 'Y',VerifiedDate = GETDATE(),VerifiedBy= '" & Gen_UserLoginID & "',ApprovalRemarks = isnull(ApprovalRemarks,'')+' " & TxtVerifiedRemark.Text & "',Concurrency = GETDATE(),UpdatedDate = GETDATE(),UpdatedBy = '" & Gen_UserLoginID & "' Where DepConID = " & Val(TxtDepositContractID) & " And ApprovalFlag In ('A','S') "
        Call TmpTable
        
        Dim TOADD As Variant
        Dim TOCC, mMsg, mGridMsg As Variant
        Dim strSplitString() As String
        Dim EmpNo_, EmpName_ As String
        Dim mR As Variant
        Dim mColor As Variant
        
        mMsg = ""
        TOADD = ""
        TOCC = ""
        
        mMsg = "<table border=1 width='100%' align='center'>"
        mMsg = mMsg & "<tr><td width='40%'><b>Deposit Contract ID</b></td><td width='60%'>" & TxtDepositContractID.Text & "</td></tr>"
        mMsg = mMsg & "<tr><td width='40%'><b>State</b></td><td width='60%'>" & TxtState.Text & "</td></tr>"
        mMsg = mMsg & "<tr><td width='40%'><b>Location</b></td><td width='60%'>" & TxtLocation.Text & "</td></tr>"
        mMsg = mMsg & "<tr><td width='40%'><b>Commodity</b></td><td width='60%'>" & TxtCommodity.Text & "</td></tr>"
        mMsg = mMsg & "<tr><td width='40%'><b>Client Name</b></td><td width='60%'>" & TxtClient.Text & "</td></tr>"
        mMsg = mMsg & "<tr><td width='40%'><b>Start Date</b></td><td width='60%'>" & Format(TxtStartDate, "DD-MMM-yyyy") & "</td></tr>"
        mMsg = mMsg & "<tr><td width='40%'><b>End Date</b></td><td width='60%'>" & Format(mLastProcess + 1, "DD-MMM-YYYY") & "</td></tr>"
        mMsg = mMsg & "<tr><td width='40%'><b>Insured By</b></td><td width='60%'>" & TxtInsurance.Text & "</td></tr>"
        mMsg = mMsg & "<tr><td width='40%'><b>Billing Cycle</b></td><td width='60%'>" & TxtBillingCycle.Text & "</td></tr>"
        mMsg = mMsg & "<tr><td width='40%'><b>Depository Billing Cycle</b></td><td width='60%'>" & TxtDepositBillingCycle.Text & "</td></tr>"
        mMsg = mMsg & "<tr><td width='40%'><b>Status</b></td><td width='60%'> Verified </td></tr>"
        mMsg = mMsg & "<tr><td width='40%'><b>Forwared TO</b></td><td width='60%'>" & TxtForwardedTo.Text & "</td></tr>"
        mMsg = mMsg & "<tr><td width='40%'><b>Remarks</b></td><td width='60%'>" & TxtRemarks & "</td></tr>"
        mMsg = mMsg & "<tr><td width='40%'><b>Verified Remarks</b></td><td width='60%'>" & TxtVerifiedRemark & "</td></tr>"
        mMsg = mMsg & "</table>"
        
        If mForwardID <> Gen_UserID Then
            TOADD = GetEmployeeEmailID(mForwardID)
        Else
            tmp = " Select EmployeeID,EmployeeNo,EmployeeEmailID from Mst_Employee Where EmployeeNo = '" & IIf(IsNull(RsTxn_DepositContract!CreatedBy), 0, RsTxn_DepositContract!CreatedBy) & "' And Isnull(Flag,'Y') = 'Y'  "
            Call TmpTable
            If TmpRs.RecordCount > 0 Then
                TOADD = IIf(IsNull(TmpRs!EmployeeEmailID), "", TmpRs!EmployeeEmailID)
            End If
        End If

        TOCC = Gen_UserEmailID
        
        ''---- Location Wise Email and Reporty
        tmp = " Select Distinct ME.EmployeeEmailID[First],MEE.EmployeeEmailID[Second],MEEE.EmployeeEmailID[Third],MEEEE.EmployeeEmailID[Fourth]"
        tmp = tmp & " From Mst_EmpCMPMapping MEC"
        tmp = tmp & " Inner Join Mst_Employee ME on MEC.EmployeeID = ME.EmployeeID And ISnull(ME.Flag,'Y') = 'Y' And ME.DesignationID <> 1"
        tmp = tmp & " Left Join Mst_Employee MEE On ME.DirectReportingTo = MEE.EmployeeID And ISnull(MEE.Flag,'Y') = 'Y' And MEE.DesignationID <> 1"
        tmp = tmp & " Left Join Mst_Employee MEEE On MEE.DirectReportingTo = MEEE.EmployeeID And ISnull(MEEE.Flag,'Y') = 'Y' And MEEE.EmployeeID <> 20"
        tmp = tmp & " Left Join Mst_Employee MEEEE On MEEE.DirectReportingTo = MEEEE.EmployeeID And ISnull(MEEEE.Flag,'Y') = 'Y'  And MEEE.EmployeeID <> 20"
        tmp = tmp & " Where MEC.LocationID = " & mLocationID & ""
        Call TmpTable
        For I = 1 To TmpRs.RecordCount
            If InStr(LCase(TOCC), LCase(IIf(IsNull(TmpRs!First), "", TmpRs!First))) = 0 Then TOCC = TOCC & "," & IIf(IsNull(TmpRs!First), "", TmpRs!First)
            If InStr(LCase(TOCC), LCase(IIf(IsNull(TmpRs!Second), "", TmpRs!Second))) = 0 Then TOCC = TOCC & "," & IIf(IsNull(TmpRs!Second), "", TmpRs!Second)
            If InStr(LCase(TOCC), LCase(IIf(IsNull(TmpRs!Third), "", TmpRs!Third))) = 0 Then TOCC = TOCC & "," & IIf(IsNull(TmpRs!Third), "", TmpRs!Third)
            If InStr(LCase(TOCC), LCase(IIf(IsNull(TmpRs!Fourth), "", TmpRs!Fourth))) = 0 Then TOCC = TOCC & "," & IIf(IsNull(TmpRs!Fourth), "", TmpRs!Fourth)
            TmpRs.MoveNext
        Next I
        ''---------
        
        
        If TOCC <> "" Then
            TOCC = TOCC & ",CMG@ncml.com"
        End If
        
        mMsgSubject = "Deposit Contract ID " & TxtDepositContractID & " Verified."
        Call SendMail(TOADD, TOCC, mMsgSubject, mMsg, "")
        
        
        
        RsTxn_DepositContract.Requery
        If RsTxn_DepositContract.RecordCount > 0 Then
            RsTxn_DepositContract.Find "DepConID = " & Val(TxtDepositContractID)
            Call Indata
            Exit Sub
        End If

End Select

End Sub

Private Sub DeleteCmd_Click()
Dim ans As Variant
On Error GoTo msgError

If LCase(mStatus) = "a" Or LCase(mStatus) = "s" Then
    MsgBox "Deposit Contract is approved ,so you could not delete.", vbOKOnly + vbExclamation
    Exit Sub
End If

ans = MsgBox("Do you really want to DELETE this record???", vbYesNo)
If ans = vbYes Then

   tmp = "Delete from Txn_DepositContractGodownDetail Where DepConID = " & Val(TxtDepositContractID.Text) & ""
   Call TmpTable

   RsTxn_DepositContract.Delete
   RsTxn_DepositContract.Update
   
   ''---------------------- Delete Log
   tmp = "Delete from Log_Txn_DepositContract Where DepConID = " & Val(TxtDepositContractID.Text) & ""
   Call TmpTable
   
   tmp = "Delete from Log_Txn_DepositContractGodownDetail Where DepConID = " & Val(TxtDepositContractID.Text) & ""
   Call TmpTable
   ''---------------------- Attachment
   tmp = " Delete From Txn_DepositContractAttachment  Where DepConID = " & Val(TxtDepositContractID.Text) & ""
   Call TmpTableSMAttachment
   ''---------------------------------
   
   If RsTxn_DepositContract.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If

   RsTxn_DepositContract.MoveNext
   If RsTxn_DepositContract.EOF = False Then
      RsTxn_DepositContract.MoveLast
   End If
   Call Indata
End If
Exit Sub
msgError:
MsgBox "Child record Present "
RsTxn_DepositContract.CancelUpdate
End Sub

Private Sub EditCmd_Click()

''-- for Concurrency Check Only for Update
tmp = " Select * from Txn_DepositContract Where DepConID = " & Val(TxtDepositContractID) & " And CONVERT(varchar(30),Concurrency,120) > '" & Format(RsTxn_DepositContract!Concurrency, "YYYY-MM-DD hh:mm:ss") & "' "
Call TmpTable
If TmpRs.EOF = False Then
    MsgBox "Sorry, System not allow to update the data because same data in updated by another user." & vbNewLine & vbNewLine & "Please reopen your form to get update information and process.", vbInformation + vbOKOnly
    Exit Sub
End If
''--------------------------------
Edit_Flg = "E"
mCopyFlag = False

CmdCopy.Enabled = False

If Left$(LCase(TxtStatus.Text), 1) = "r" Or Left$(LCase(TxtStatus.Text), 1) = "c" Then
   MsgBox "Deposit Contract can't modified, Already " & TxtStatus.Text & ".", vbOKOnly + vbInformation, App.Title
   Exit Sub
End If
''-- Rights applyed on 05 Jan 2018
''If LCase(Gen_UserRole) <> "superadmin" And LCase(Gen_UserRole) <> "head cmg" And LCase(Gen_UserRole) <> "power" Then
If LCase(Gen_UserRole) <> "power" And mApprovedRights = False Then
   If Left$(LCase(TxtStatus.Text), 1) = "a" Or Left$(LCase(TxtStatus.Text), 1) = "s" Then
      MsgBox "Deposit Contract can't modified, Already " & TxtStatus.Text & ".", vbOKOnly + vbInformation, App.Title
      Exit Sub
   End If
Else
''   If Left$(LCase(TxtStatus.Text), 1) = "a" Or Left$(LCase(TxtStatus.Text), 1) = "s" Then
''      MsgBox "Deposit Contract can't modified, Already " & TxtStatus.Text & ".", vbOKOnly + vbInformation, App.Title
''      Exit Sub
''      'MSHFlexGrid1.Enabled = False
''      'MSHFlexGrid2.Enabled = False
''      'Exit Sub
''   End If
''-- 01 March 2018 Allowed to update the Attachment after Approved but not verified by given userrole only
   If (Left$(LCase(TxtStatus.Text), 1) = "a" Or Left$(LCase(TxtStatus.Text), 1) = "s") And mVerifiyFlag = True Then
      MsgBox "Deposit Contract can't modified, Already " & TxtStatus.Text & " and Verified", vbOKOnly + vbInformation, App.Title
      Exit Sub
   ElseIf (Left$(LCase(TxtStatus.Text), 1) = "a" Or Left$(LCase(TxtStatus.Text), 1) = "s") And mVerifiyFlag = False Then
      If LCase(Gen_UserRole) = "superadmin" Or LCase(Gen_UserRole) = "headcmg" Or LCase(Gen_UserRole) = "power" Then
         MsgBox "You can edit only attachment because Deposit Contract is already Approved, but not yet verified."
         SSTab1.Tab = 1
         ''----- to Edit only Attachment Frame
         Call GButtonLock(Me, False)
         
         CmbContractType.Text = TxtContractType.Text
         CmbState.Text = TxtState.Text
         CmbLocation.Text = TxtLocation.Text
         CmbCommodity.Text = TxtCommodity.Text
         CmbClient.Text = TxtClient.Text
         CmbBillingCycle.Text = TxtBillingCycle.Text
         CmbDepositBillingCycle.Text = TxtDepositBillingCycle.Text
         CmbRatePer.Text = TxtRatePer.Text
         CmbForwardedTo.Text = TxtForwardedTo.Text
         CmbStatus.Text = TxtStatus.Text
         CmbInsurance.Text = TxtInsurance.Text
         
         FrameAttachment.Enabled = True
         CmdRefreshRate.Visible = False
         CmdClosed.Enabled = False
         CmdVerified.Enabled = False
         Exit Sub
         ''------------------------------
      Else
         MsgBox "Deposit Contract can't modified, Already " & TxtStatus.Text & ".", vbOKOnly + vbInformation, App.Title
         Exit Sub
      End If
   End If

''-----------------------------
End If

Call GButtonLock(Me, False)
If Left$(LCase(TxtStatus.Text), 1) <> "a" And Left$(LCase(TxtStatus.Text), 1) <> "s" Then
   'TxtState.Visible = False:    CmbState.Visible = True
   'TxtLocation.Visible = False:   CmbLocation.Visible = True
   'TxtCommodity.Visible = False:   CmbCommodity.Visible = True
   'TxtClient.Visible = False:   CmbClient.Visible = True
   'TxtBillingCycle.Visible = False:   CmbBillingCycle.Visible = True
   'TxtDepositBillingCycle.Visible = False:   CmbDepositBillingCycle.Visible = True
   'TxtRatePer.Visible = False:   CmbRatePer.Visible = True
   TxtForwardedTo.Visible = False
   CmbForwardedTo.Visible = True
   CmbForwardedTo.Enabled = True
   TxtStatus.Visible = False
   CmbStatus.Visible = True
   'TxtInsurance.Visible = False:   CmbInsurance.Visible = True
   'TxtQuantum.Locked = False:   TxtInsurance.Locked = False
   'TxtTentativeDuration.Locked = False:   TxtLockInPeriod.Locked = False
   Frame2.Enabled = True
   If Check4.Value = vbChecked Then CmdAddNewBagSize.Visible = True: CmdAddNewBagSize.Enabled = True Else CmdAddNewBagSize.Visible = False: CmdAddNewBagSize.Enabled = False
End If

TxtQuantum.Locked = True
TxtInsurance.Locked = True
TxtTentativeDuration.Locked = True
TxtLockInPeriod.Locked = True

TxtStartDate.Enabled = False
TxtEndDate.Enabled = False
   
CmbState.Text = TxtState.Text
CmbLocation.Text = TxtLocation.Text
CmbCommodity.Text = TxtCommodity.Text
CmbClient.Text = TxtClient.Text
CmbBillingCycle.Text = TxtBillingCycle.Text
CmbDepositBillingCycle.Text = TxtDepositBillingCycle.Text
CmbRatePer.Text = TxtRatePer.Text
CmbForwardedTo.Text = TxtForwardedTo.Text
CmbStatus.Text = TxtStatus.Text
CmbInsurance.Text = TxtInsurance.Text

''-- 14 Feb 2018
CmbContractType = TxtContractType
''-------------

CmbBillingCycle.Visible = True: CmbBillingCycle.Enabled = True: TxtBillingCycle.Visible = False
CmbDepositBillingCycle.Visible = True: CmbDepositBillingCycle.Enabled = True: TxtDepositBillingCycle.Visible = False
CmbInsurance.Visible = True: CmbInsurance.Enabled = True: TxtInsurance.Visible = False

Frame1.Enabled = True
FrameAttachment.Enabled = True

Check1.Enabled = False  ''-- converted Client
Check2.Enabled = True   ''-- Select All
Check4.Enabled = True   ''-- BagWise / Quantity Wise

cmdExportDetails.Enabled = True
TxtFinalPrice.Visible = False
'''Call IndataDepositContractGodownDetail_Edit
'''If Check4.Value = 0 Then
'''    Call CheckBaseRate
'''ElseIf Check4.Value = 1 Then
'''    Call CheckBaseRateBags
'''End If
CmdRefreshRate.Visible = True
CmdClosed.Enabled = False
End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub FirstCmd_Click()
RsTxn_DepositContract.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub

Public Sub Grid_Head()
With MSHFlexGrid1
     .Font.Size = 9
     .Clear
     .Rows = 2
     .Cols = 18
     .FixedCols = 0
     .TextMatrix(0, 0) = "Select":                                              .ColWidth(0) = 500
     .TextMatrix(0, 1) = "CMP Name":                                            .ColWidth(1) = 1800
     .TextMatrix(0, 2) = "Godown No":                                           .ColWidth(2) = 900
     .TextMatrix(0, 3) = "Address":                                             .ColWidth(3) = 2000
     .TextMatrix(0, 4) = "Godown Type":                                         .ColWidth(4) = 1000
     .TextMatrix(0, 5) = "Rent":                                                .ColWidth(5) = 1200
     .TextMatrix(0, 6) = "Area in Sq. Ft.":                                     .ColWidth(6) = 1200      ''Capacity in MT
     .TextMatrix(0, 7) = "Billing Rate/MT (As per System)":                     .ColWidth(7) = 1200
     .TextMatrix(0, 8) = "Billing Rate/MT (Proposed/ Approved)":                .ColWidth(8) = 1200
     .TextMatrix(0, 9) = "GodownID":                                            .ColWidth(9) = 0
     .TextMatrix(0, 10) = "ROWID":                                              .ColWidth(10) = 0
     .TextMatrix(0, 11) = "Deposit/ Withdrawal Rate/MT (As per System)":        .ColWidth(11) = 1200
     .TextMatrix(0, 12) = "Deposit/ Withdrawal Rate/MT (Proposed/ Approved)":   .ColWidth(12) = 1200
     .TextMatrix(0, 13) = "Billing Rate":                                       .ColWidth(13) = 0
     .TextMatrix(0, 14) = "Deposit Billing Rate":                               .ColWidth(14) = 0
     .TextMatrix(0, 15) = "UnitTypeID":                                         .ColWidth(15) = 0
     .TextMatrix(0, 16) = "Billing Rate/MT (Monthly Cycle)":                    .ColWidth(16) = 1200
     .TextMatrix(0, 17) = "ApprovalLimit":                                      .ColWidth(17) = 0
     
     .RowHeight(0) = 850
End With
End Sub

Public Sub Grid_Head4()
With MSHFlexGrid4
     .Font.Size = 9
     .Clear
     .Rows = 2
     .Cols = 21
     .FixedCols = 0
     
    .TextMatrix(0, 0) = "Select":                                              .ColWidth(0) = 500
    .TextMatrix(0, 1) = "CMP Name":                                            .ColWidth(1) = 1300
    .TextMatrix(0, 2) = "Godown No":                                           .ColWidth(2) = 900
    .TextMatrix(0, 3) = "Address":                                             .ColWidth(3) = 1200
    .TextMatrix(0, 4) = "Godown Type":                                         .ColWidth(4) = 1000
    .TextMatrix(0, 5) = "Rent":                                                .ColWidth(5) = 1000
    .TextMatrix(0, 6) = "Area in Sq. Ft.":                                     .ColWidth(6) = 1000   ''Capacity in MT
    .TextMatrix(0, 7) = "Bag Size (Minimum)":                                  .ColWidth(7) = 900
    .TextMatrix(0, 8) = "Bag Size (Maximum)":                                  .ColWidth(8) = 900
    .TextMatrix(0, 9) = "Billing Rate/MT (As per System)":                    .ColWidth(9) = 1000
    .TextMatrix(0, 10) = "Billing Rate/MT (Proposed/ Approved)":              .ColWidth(10) = 1200
    .TextMatrix(0, 11) = "GodownID":                                           .ColWidth(11) = 0
    .TextMatrix(0, 12) = "ROWID":                                              .ColWidth(12) = 0
    .TextMatrix(0, 13) = "Deposit/ Withdrawal Rate/MT (As per System)":        .ColWidth(13) = 1200
    .TextMatrix(0, 14) = "Deposit/ Withdrawal Rate/MT (Proposed/ Approved)":   .ColWidth(14) = 1200
    .TextMatrix(0, 15) = "Billing Rate":                                       .ColWidth(15) = 0
    .TextMatrix(0, 16) = "Deposit Billing Rate":                               .ColWidth(16) = 0
    .TextMatrix(0, 17) = "UnitTypeID":                                         .ColWidth(17) = 0
    .TextMatrix(0, 18) = "Billing Rate/MT (Monthly Cycle)":                    .ColWidth(18) = 1200
    .TextMatrix(0, 19) = "ApprovalLimit":                                      .ColWidth(19) = 0
    .TextMatrix(0, 20) = "BagFlag":                                            .ColWidth(20) = 0
    .RowHeight(0) = 850
End With
End Sub

Public Sub Grid_Head6()
    With MSHFlexGrid6
        If LCase(mBagSizeFlag) = "y" Then
             .Font.Size = 9
             .Clear
             .Rows = 2
             .Cols = 21
             .FixedCols = 0
            .TextMatrix(0, 0) = "Select":                                              .ColWidth(0) = 0
            .TextMatrix(0, 1) = "CMP Name":                                            .ColWidth(1) = 1300
            .TextMatrix(0, 2) = "Godown No":                                           .ColWidth(2) = 900
            .TextMatrix(0, 3) = "Address":                                             .ColWidth(3) = 1200
            .TextMatrix(0, 4) = "Godown Type":                                         .ColWidth(4) = 1000
            .TextMatrix(0, 5) = "Rent":                                                .ColWidth(5) = 1000
            .TextMatrix(0, 6) = "Area in Sq. Ft.":                                     .ColWidth(6) = 1000   ''Capacity in MT
            .TextMatrix(0, 7) = "Bag Size (Minimum)":                                  .ColWidth(7) = 900
            .TextMatrix(0, 8) = "Bag Size (Maximum)":                                  .ColWidth(8) = 900
            .TextMatrix(0, 9) = "Billing Rate/MT (As per System)":                    .ColWidth(9) = 1000
            .TextMatrix(0, 10) = "Billing Rate/MT (Proposed/ Approved)":              .ColWidth(10) = 1200
            .TextMatrix(0, 11) = "GodownID":                                           .ColWidth(11) = 0
            .TextMatrix(0, 12) = "ROWID":                                              .ColWidth(12) = 0
            .TextMatrix(0, 13) = "Deposit/ Withdrawal Rate/MT (As per System)":        .ColWidth(13) = 1200
            .TextMatrix(0, 14) = "Deposit/ Withdrawal Rate/MT (Proposed/ Approved)":   .ColWidth(14) = 1200
            .TextMatrix(0, 15) = "Billing Rate":                                       .ColWidth(15) = 0
            .TextMatrix(0, 16) = "Deposit Billing Rate":                               .ColWidth(16) = 0
            .TextMatrix(0, 17) = "UnitTypeID":                                         .ColWidth(17) = 0
            .TextMatrix(0, 18) = "Billing Rate/MT (Monthly Cycle)":                   .ColWidth(18) = 1200
            .TextMatrix(0, 19) = "ApprovalLimit":                                      .ColWidth(19) = 0
            .TextMatrix(0, 20) = "BagFlag":                                            .ColWidth(20) = 0
            .RowHeight(0) = 850
        Else
            .Font.Size = 9
            .Clear
            .Rows = 2
            .Cols = 18
            .FixedCols = 0
            .TextMatrix(0, 0) = "Select":                                              .ColWidth(0) = 0
            .TextMatrix(0, 1) = "CMP Name":                                            .ColWidth(1) = 1800
            .TextMatrix(0, 2) = "Godown No":                                           .ColWidth(2) = 900
            .TextMatrix(0, 3) = "Address":                                             .ColWidth(3) = 2000
            .TextMatrix(0, 4) = "Godown Type":                                         .ColWidth(4) = 1000
            .TextMatrix(0, 5) = "Rent":                                                .ColWidth(5) = 1200
            .TextMatrix(0, 6) = "Area in Sq. Ft.":                                     .ColWidth(6) = 1200      ''Capacity in MT
            .TextMatrix(0, 7) = "Billing Rate/MT (As per System)":                     .ColWidth(7) = 1200
            .TextMatrix(0, 8) = "Billing Rate/MT (Proposed/ Approved)":                .ColWidth(8) = 1200
            .TextMatrix(0, 9) = "GodownID":                                            .ColWidth(9) = 0
            .TextMatrix(0, 10) = "ROWID":                                              .ColWidth(10) = 0
            .TextMatrix(0, 11) = "Deposit/ Withdrawal Rate/MT (As per System)":        .ColWidth(11) = 1200
            .TextMatrix(0, 12) = "Deposit/ Withdrawal Rate/MT (Proposed/ Approved)":   .ColWidth(12) = 1200
            .TextMatrix(0, 13) = "Billing Rate":                                       .ColWidth(13) = 0
            .TextMatrix(0, 14) = "Deposit Billing Rate":                               .ColWidth(14) = 0
            .TextMatrix(0, 15) = "UnitTypeID":                                         .ColWidth(15) = 0
            .TextMatrix(0, 16) = "Billing Rate/MT (Monthly Cycle)":                    .ColWidth(16) = 1200
            .TextMatrix(0, 17) = "ApprovalLimit":                                      .ColWidth(17) = 0
            .RowHeight(0) = 850
        End If
    End With
End Sub


Public Sub Grid_Head2()
With MSHFlexGrid2
     .Font.Size = 9
     .Clear
     .Rows = 2
     .Cols = 24
     ''-- Rights applyed on 05 Jan 2018
     ''If LCase(Gen_UserRole) = "head cmg" Or LCase(Gen_UserRole) = "superadmin" Or LCase(Gen_UserRole) = "power" Then .Cols = 21
     If LCase(Gen_UserRole) = "power" Or mApprovedRights = True Then .Cols = 26
     .FixedCols = 0
     .TextMatrix(0, 0) = "Deposit Contract ID":                             .ColWidth(0) = 800
     .TextMatrix(0, 1) = "Date":                                            .ColWidth(1) = 1000
     .TextMatrix(0, 2) = "Contract Type":                                   .ColWidth(2) = 800 ''-- 14 Feb 2018
     .TextMatrix(0, 3) = "State":                                           .ColWidth(3) = 1500
     .TextMatrix(0, 4) = "Location":                                        .ColWidth(4) = 1500
     .TextMatrix(0, 5) = "Commodity":                                       .ColWidth(5) = 1500
     .TextMatrix(0, 6) = "Client":                                          .ColWidth(6) = 3000
     .TextMatrix(0, 7) = "Insurance By":                                    .ColWidth(7) = 900
     .TextMatrix(0, 8) = "Quantity/ Bag":                                   .ColWidth(8) = 900
     .TextMatrix(0, 9) = "Deposit Contract Start Date":                     .ColWidth(9) = 1200
     .TextMatrix(0, 10) = "Deposit Contract End Date":                      .ColWidth(10) = 1200
     .TextMatrix(0, 11) = "Billing Cycle":                                  .ColWidth(11) = 1000
     .TextMatrix(0, 12) = "Deposit/ Withdrawal Billing Cycle":              .ColWidth(12) = 1000
     .TextMatrix(0, 13) = "GST":                                            .ColWidth(13) = 1000
     .TextMatrix(0, 14) = "Forwarded To":                                   .ColWidth(14) = 1800
     .TextMatrix(0, 15) = "Status":                                         .ColWidth(15) = 1000
     .TextMatrix(0, 16) = "Tentative Duration of Storage (in months)":      .ColWidth(16) = 900
     .TextMatrix(0, 17) = "Lock-in Period (in months)":                     .ColWidth(17) = 900
     .TextMatrix(0, 18) = "Indicative Quantum":                             .ColWidth(18) = 900
     .TextMatrix(0, 19) = "Approved On":                                    .ColWidth(19) = 1800
     .TextMatrix(0, 20) = "Approved By":                                    .ColWidth(20) = 1800
     ''--- 02 Feb 2018 for verified column
        .TextMatrix(0, 21) = "Verified":                                     .ColWidth(21) = 800
        .TextMatrix(0, 22) = "Verified On":                                  .ColWidth(22) = 1800
        .TextMatrix(0, 23) = "Verified By":                                  .ColWidth(23) = 1800
     ''---
     ''-- Rights applyed on 05 Jan 2018
     ''If LCase(Gen_UserRole) = "head cmg" Or LCase(Gen_UserRole) = "superadmin" Or LCase(Gen_UserRole) = "power" Then
     If LCase(Gen_UserRole) = "power" Or mApprovedRights = True Then
        .TextMatrix(0, 24) = "Closed On":                                    .ColWidth(24) = 1800
        .TextMatrix(0, 25) = "Closed By":                                    .ColWidth(25) = 1800
     End If
     .RowHeight(0) = 800
End With
End Sub
Public Sub ClearControl()

mCopyFlag = False
mFlag = False
mStatus = ""

TxtDepositContractID = ""
TxtDepositContractDate = Date
CmbState.Text = ""
TxtState = ""
CmbLocation = ""
TxtLocation = ""
CmbClient = ""
TxtClient = ""
CmbInsurance = ""
TxtInsurance = ""
CmbCommodity = ""
TxtCommodity = ""
TxtAddress = ""
TxtContactDetails = ""
TxtEmailId = ""
TxtPinCode = ""
TxtContactNo = ""
TxtQuantum = ""
TxtStartDate.Enabled = True: TxtStartDate = Date
TxtEndDate.Enabled = True
''If Month(TxtStartDate) <= 9 Then
''   TxtEndDate = Date
''Else
''   TxtEndDate = Date + 365
''End If
TxtEndDate = Date + mDepositContractDefaultDaysRange

CmbBillingCycle = ""
TxtBillingCycle = ""
CmbDepositBillingCycle = ""
TxtDepositBillingCycle = ""
TxtMonthlyMTRate = ""
CmbRatePer = "Excluded"
TxtRatePer = "Excluded"
TxtTentativeDuration = ""
TxtLockInPeriod = ""
TxtRemarks.Text = ""
TxtStatus.Text = ""
TxtForwardedTo.Text = ""
CmbStatus.Text = ""
CmbForwardedTo.Text = ""
TxtCreated.Text = ""
TxtUpdated.Text = ""
TxtApproved.Text = ""
TxtClosed.Text = ""

CmbState.Visible = True
TxtState.Visible = False

CmbLocation.Visible = True
TxtLocation.Visible = False

CmbCommodity.Visible = True
TxtCommodity.Visible = False

CmbClient.Visible = True
TxtClient.Visible = False

CmbBillingCycle.Visible = True
TxtBillingCycle.Visible = False

CmbDepositBillingCycle.Visible = True
TxtDepositBillingCycle.Visible = False

CmbRatePer.Visible = False
TxtRatePer.Visible = True

CmbStatus.Visible = True
TxtStatus.Visible = False

CmbForwardedTo.Visible = True
TxtForwardedTo.Visible = False

CmbInsurance.Visible = True
TxtInsurance.Visible = False

TxtQuantum.Locked = False
TxtTentativeDuration.Locked = False
TxtLockInPeriod.Locked = False
TxtFinalPrice.Visible = False

Call Grid_Head
Call Grid_Head3
Call Grid_Head4
Call Grid_HeadAttachment

Call ClearFrameControl
Check1.Enabled = True: Check1.Value = 0 ''-- Converted Client
Check2.Enabled = True: Check2.Value = 0 ''-- Select All
Check4.Enabled = True: Check4.Value = 0 ''-- Calculation on Bags flag

Check2.Enabled = True ''-- Select All
MSHFlexGrid4.Visible = False

MSHFlexGrid1.Enabled = True
Frame1.Enabled = True
Frame2.Enabled = True
Frame3.Visible = False
Frame4.Visible = False
Frame5.Visible = False
FrameAttachment.Enabled = True
CmdCopy.Enabled = False
CmdClosed.Enabled = False
CmdAddNewBagSize.Visible = False

''-- 02 Feb 2018 -- Verified DC
Frame6.Visible = False
CmdVerified.Enabled = False
TxtVerifiedRemark = ""
''------------------------------

SSTab1.Tab = 0
End Sub

Public Sub ClearFrameControl()
CmbCMP = ""
CmbGodown = ""
TxtFinalPrice = ""
TxtGodownType = ""
TxtRent = ""
TxtCapacity = ""
TxtGodownAddress = ""
End Sub

Public Sub GetClientDetails(mmClientID As Variant)
tmp = " Select * from Mst_Client Where ClientID=" & mmClientID & ""
Call Tmp1Table
If TmpRs1.RecordCount > 0 Then
   TxtAddress.Text = IIf(IsNull(TmpRs1!Address), "", TmpRs1!Address)
   TxtContactDetails.Text = IIf(IsNull(TmpRs1!MobileNo), "", TmpRs1!MobileNo)
   TxtEmailId.Text = IIf(IsNull(TmpRs1!EmailID), "", TmpRs1!EmailID)
   TxtContactNo.Text = IIf(IsNull(TmpRs1!TelephoneNo), "", TmpRs1!TelephoneNo)
End If
End Sub

Private Sub Form_Load()
Dim mSplit As Variant
Dim FrmLoadAccess() As Boolean
Dim mR As Double

FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Call GButtonLockAD(Me)
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If

Call Grid_Head
Call Grid_Head2
Call Grid_Head3
Call Grid_Head4
SFrom = Date - 30
STo = Date

''-- Calculation Factor Values From Email Config File
mCurrentConversionFactor = Val(Gen_DepositContractCalculationFactor("DepositContractCurrentConversionFactor-"))
mDivisionFactor = Val(Gen_DepositContractCalculationFactor("DepositContractDivisionFactor-"))                        ''-- Percentage %
mFumigationFactorY = Val(Gen_DepositContractCalculationFactor("DepositContractFumigationFactorY-"))
mFumigationFactorN = Val(Gen_DepositContractCalculationFactor("DepositContractFumigationFactorN-"))
mCMPMtoAMRMApprovalLimit = Val(Gen_DepositContractCalculationFactor("DepositContractCMPMtoAMRMApprovalLimit-"))      ''-- Percentage %
''-- if Godown Type is Non Leased then (Rent is not available) then PAN India Mini Value From Email Config File
mKnownRentMinimumValue = Val(Gen_DepositContractCalculationFactor("DepositContractNonKnownRentMinimumValue-"))
''-- CMG Group ID to be List in Forward to Additionlly and Approved for Same
mDepositContractCMGGroupToBeApproved = Gen_DepositContractCalculationFactor("DepositContractCMGGroupToBeApproved-")
mDepositContractDefaultDaysRange = Gen_DepositContractCalculationFactor("DepositContractDefaultDaysRange-")
mDepositContractInFutureDateFlag = Gen_DepositContractCalculationFactor("DepositContractInFutureDateFlag-")
''--- Ullhas Patil 06 Jan 2018
mDepositContractDefaultAttachmentFlag = Gen_DepositContractCalculationFactor("DepositContractDefaultAttachmentFlag-")
mDepositContractAllowBackDateEntryDays = Gen_DepositContractCalculationFactor("DepositContractAllowBackDateEntryDays-")
mDepositContractDaysGapStartDateEndDate = Gen_DepositContractCalculationFactor("DepositContractDaysGapStartDateEndDate-")

''--- Ullhas Patil 11 Jan 2018
mDepositContractAdditionalGroupToBeApproved = Gen_DepositContractCalculationFactor("DepositContractAdditionalGroupToBeApproved-")
mDepositContractAdditionalGroupToBeApprovedList = Split(mDepositContractAdditionalGroupToBeApproved, ",")
mDepositContractAdditionalGroupToBeApproved = "'0'"
For mR = 0 To UBound(mDepositContractAdditionalGroupToBeApprovedList)
    mDepositContractAdditionalGroupToBeApproved = mDepositContractAdditionalGroupToBeApproved & ",'" & mDepositContractAdditionalGroupToBeApprovedList(mR) & "'"
Next mR
mDepositContractAdditionalGroupToBeApprovedList = mDepositContractAdditionalGroupToBeApproved
''---------
mCurrentBaseRentFactor = Val(Gen_DepositContractCalculationFactor("DepositContractCurrentBaseRentFactor-"))
mUpperHierarchyEmployeeIDList = ""
Call UpperHierarchyEmployeeID(Gen_EmployeeID)

''-- Rigths Defination add on 05 Jan 2018
mSplit = Split(mDepositContractCMGGroupToBeApproved, ",")
mApprovedRights = False
For mR = 0 To UBound(mSplit)
    If Val(mSplit(mR)) = Gen_UserTypeID Then
        mApprovedRights = True
        Exit For
    End If
Next mR
''----
''--- Ullhas Patil 12 Jan 2018 if the additional User ID as current user then allow as CMG member
mSplit = Split(mDepositContractAdditionalGroupToBeApprovedList, ",")
For mR = 0 To UBound(mSplit)
    If LCase(mSplit(mR)) = LCase("'" & Gen_UserLoginID & "'") Then
        mApprovedRights = True
        Exit For
    End If
Next mR
''---------

If Gen_WcLocation = "" Then
    Call TableTxn_DepositContract
Else
    Call TableTxn_DepositContract(" Where " & Gen_WcLocation & " ")
End If
If RsTxn_DepositContract.RecordCount = 0 Then
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
RsTxn_DepositContract.MoveLast
Call Indata

End Sub

Private Sub LastCmd_Click()
RsTxn_DepositContract.MoveLast
Call Indata
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False
End Sub


Private Sub MSHFlexGrid1_Click()
mSelCol = MSHFlexGrid1.Col
mSelRow = MSHFlexGrid1.row
TxtFinalPrice.Visible = False
TxtBagSize.Visible = False
If SaveCmd.Enabled = False Then Exit Sub
If Edit_Flg = "E" And Val(TxtDepositContractID) > 0 And (LCase(mStatus) = "a" Or LCase(mStatus) = "s") Then
   'If LCase(Gen_UserRole) <> "head cmg" And LCase(Gen_UserRole) <> "power" Then
      Exit Sub
   'End If
End If

If CmbBillingCycle.Text = "" Then
   MsgBox "Blank Billing Cycle Not Allowed.", vbOKOnly + vbExclamation, App.Title
   CmbBillingCycle.SetFocus
   Exit Sub
End If
If CmbDepositBillingCycle.Text = "" Then
   MsgBox "Blank Deposit Billing Cycle Not Allowed.", vbOKOnly + vbExclamation, App.Title
   CmbDepositBillingCycle.SetFocus
   Exit Sub
End If

If LCase(mStatus) <> "a" And LCase(mStatus) <> "s" Then
With MSHFlexGrid1
     If .TextMatrix(mSelRow, 1) = "" Then Exit Sub
     If .TextMatrix(mSelRow, 0) = strChecked Then
        If .Col <> 0 And .Col <> 8 And .Col <> 12 Then
            Call ShowPopDetails(MSHFlexGrid1, CInt(mSelRow))
            Frame1.Enabled = False
            SSTab1.Enabled = False
            ButtonFrm.Enabled = False
            Frame4.Left = (SSTab1.Width / 2) - (Frame4.Width / 2) ''-- Center Align
            Frame4.Visible = True
        End If
     End If
     If .Col = 0 Then
           .Col = 0
           TxtFinalPrice.Visible = False
           If .TextMatrix(.row, .Col) = strUnChecked Then
              .TextMatrix(.row, .Col) = strChecked
              'If Trim(TxtDepositContractID.Text) = "" Then
              '  Call CmbBillingCycle_LostFocus
              '  Call CmbDepositBillingCycle_LostFocus
              'Else
              '  Call CheckRateByColor(.row)
              'End If
              
            ''------------------------------ 01 Feb 2018
            ''-- Lease / Own Godown
            If Val(.TextMatrix(.row, 15)) = 1 Or Val(.TextMatrix(.row, 15)) = 7 Then
                .TextMatrix(.row, 7) = Format(CalculateRate(.TextMatrix(.row, 13), mBillingRatio, CmbBillingCycle.Text), "#0.0000")
                mBillingOnePercentage = Format(IIf(Val(.TextMatrix(.row, 7)) > 0, Val(.TextMatrix(.row, 7)) / 100, 0), "#0.0000")
                .TextMatrix(.row, 8) = Format(Val(.TextMatrix(.row, 7)) + (mBillingOnePercentage * .TextMatrix(.row, 17)), "#0.0000")
            Else
                .TextMatrix(.row, 8) = .TextMatrix(.row, 8)
            End If
            If Val(.TextMatrix(.row, 15)) = 1 Or Val(.TextMatrix(.row, 15)) = 7 Then
                .TextMatrix(.row, 11) = Format(CalculateRate(.TextMatrix(.row, 14), mDepositBillingRatio, CmbDepositBillingCycle.Text), "#0.0000")
                mDepositOnePercentage = Format(IIf(Val(.TextMatrix(.row, 11)) > 0, Val(.TextMatrix(.row, 11)) / 100, 0), "#0.0000")
                .TextMatrix(.row, 12) = Format(Val(.TextMatrix(.row, 11)) + (mDepositOnePercentage * .TextMatrix(.row, 17)), "#0.0000")
            Else
               .TextMatrix(.row, 12) = .TextMatrix(.row, 12)
            End If
            Call CheckRateByColor(.row)
            ''------------------------------
              
           Else
              .TextMatrix(.row, .Col) = strUnChecked
              Call ResetCell(MSHFlexGrid1, 8, 12, .row)
           End If
     ElseIf .Col = 8 Or .Col = 12 Then
         If .TextMatrix(.row, 0) = strUnChecked Then Exit Sub
         
         TxtFinalPrice.Left = .Left + .CellLeft
         TxtFinalPrice.Top = .Top + .CellTop
         TxtFinalPrice.Width = .CellWidth
         TxtFinalPrice.Height = .CellHeight - 10
         TxtFinalPrice.Text = .Text
         TxtFinalPrice.SelStart = 0
         TxtFinalPrice.SelLength = Len(.Text)
         TxtFinalPrice.Visible = True
     Else
         TxtFinalPrice.Visible = False
     End If
End With
End If
End Sub

Private Sub MSHFlexGrid1_Scroll()
If TxtFinalPrice.Visible = True Then TxtFinalPrice.Visible = False
End Sub

Private Sub MSHFlexGrid2_Click()
Dim I As Variant
I = MSHFlexGrid2.RowSel
If I <= 0 Then Exit Sub
If MSHFlexGrid2.TextMatrix(I, 0) = "" Then Exit Sub
RsTxn_DepositContract.MoveFirst
RsTxn_DepositContract.Find "DepConID = " & MSHFlexGrid2.TextMatrix(I, 0)
If Not RsTxn_DepositContract.EOF = True Then
   Call Indata
   Frame0.Visible = False
End If
End Sub

Private Sub MSHFlexGrid4_Click()
Dim RO As Double, CO As Double

mSelCol = MSHFlexGrid4.Col
mSelRow = MSHFlexGrid4.row
TxtFinalPrice.Visible = False
TxtBagSize.Visible = False

If SaveCmd.Enabled = False Then Exit Sub
If Edit_Flg = "E" And Val(TxtDepositContractID) > 0 And (LCase(mStatus) = "a" Or LCase(mStatus) = "s") Then
   'If LCase(Gen_UserRole) <> "head cmg" And LCase(Gen_UserRole) <> "power" Then
      Exit Sub
   'End If
End If

If CmbBillingCycle.Text = "" Then
   MsgBox "Blank Billing Cycle Not Allowed.", vbOKOnly + vbExclamation, App.Title
   CmbBillingCycle.SetFocus
   Exit Sub
End If
If CmbDepositBillingCycle.Text = "" Then
   MsgBox "Blank Deposit Billing Cycle Not Allowed.", vbOKOnly + vbExclamation, App.Title
   CmbDepositBillingCycle.SetFocus
   Exit Sub
End If

If LCase(mStatus) <> "a" And LCase(mStatus) <> "s" Then
With MSHFlexGrid4
     If .TextMatrix(mSelRow, 1) = "" Then Exit Sub
        If .TextMatrix(mSelRow, 0) = strChecked Then
           If .Col <> 0 And .Col <> 7 And .Col <> 8 And .Col <> 10 And .Col <> 14 Then
               Call ShowPopDetails(MSHFlexGrid4, CInt(mSelRow))
               Frame1.Enabled = False
               SSTab1.Enabled = False
               
               Frame4.Left = (SSTab1.Width / 2) - (Frame4.Width / 2) ''-- Center Align
               Frame4.Visible = True
           End If
        End If
     
     If .Col = 0 Then
        .Col = 0
        TxtBagSize.Visible = False
        TxtFinalPrice.Visible = False
        
        If .TextMatrix(.row, .Col) = strUnChecked Then
           .TextMatrix(.row, .Col) = strChecked
           ''Call CheckBagRateByColor(.row)
           
           ''------------------ 01 Feb 2018
              ''-- Lease / Own Godown
                If Val(.TextMatrix(.row, 17)) = 1 Or Val(.TextMatrix(.row, 17)) = 7 Then     ''-- If Godown Type is not Lease (ie Rent not available then on Billing Cycle Only Proposed Rate get Calculate)
                    .TextMatrix(.row, 9) = Format(CalculateRate(.TextMatrix(.row, 15), mBillingRatio, CmbBillingCycle) * Val(.TextMatrix(.row, 8)) / 1000, "#0.0000")
                    mBillingOnePercentage = Format(Val(.TextMatrix(.row, 9)) / 100, "#0.0000")
                    .TextMatrix(.row, 10) = Format(Val(.TextMatrix(.row, 9)) + (mBillingOnePercentage * .TextMatrix(.row, 19)), "#0.0000")
                 Else
                    .TextMatrix(.row, 10) = .TextMatrix(.row, 10)
                 End If
                 If Val(.TextMatrix(.row, 17)) = 1 Or Val(.TextMatrix(.row, 17)) = 7 Then     ''-- If Godown Type is not Lease (ie Rent not available then on Billing Cycle Only Proposed Rate get Calculate)
                    .TextMatrix(.row, 13) = Format(CalculateRate(.TextMatrix(.row, 16), mDepositBillingRatio, CmbDepositBillingCycle) * Val(.TextMatrix(.row, 8)) / 1000, "#0.0000")
                    mDepositOnePercentage = Format(Val(.TextMatrix(.row, 13)) / 100, "#0.0000")
                    .TextMatrix(.row, 14) = Format(Val(.TextMatrix(.row, 13)) + (mDepositOnePercentage * .TextMatrix(.row, 19)), "#0.0000")
                 Else
                    .TextMatrix(.row, 14) = .TextMatrix(.row, 14)
                 End If
                 Call CheckBagRateByColor(.row)
           ''----------------------------------
           
        Else
            ''--- if row is added by system then allow to deselect if row is added by mannual then remove the row
            ''--- system is not allow the modified row added by the mannual bag size
            
            If LCase(.TextMatrix(.row, 20)) = LCase("M") Then
                If mSelRow > 0 Then
                   If .TextMatrix(mSelRow, 1) <> "" Then
                      For RO = mSelRow To .Rows - 2
                          For CO = 0 To .Cols - 1
                               .TextMatrix(RO, CO) = .TextMatrix(RO + 1, CO)
                          Next
                          Call CheckBagRateByColor(CInt(mSelRow))
                      Next
                      .Rows = .Rows - 1
                   End If
                End If
            Else
                .TextMatrix(.row, .Col) = strUnChecked
                Call ResetCell(MSHFlexGrid4, 10, 14, .row)
            End If
        End If
     ElseIf .Col = 7 Or .Col = 8 Then
        If .TextMatrix(.row, 0) = strUnChecked Then Exit Sub
        If LCase(.TextMatrix(.row, 20)) = LCase("M") Then
            MsgBox "System, is not allow to edit to modified the new bag size range." & vbNewLine & "If you want to modified the bag size range, remove the bag size range to deselect and enter again.", vbInformation
            Exit Sub
        End If
        TxtFinalPrice.Visible = False
        TxtBagSize.Visible = True
        TxtBagSize.Left = .Left + .CellLeft
        TxtBagSize.Top = .Top + .CellTop
        TxtBagSize.Width = .CellWidth
        TxtBagSize.Height = .CellHeight - 10
        TxtBagSize.Text = .Text
        TxtBagSize.SelStart = 0
        TxtBagSize.SelLength = Len(.Text)
        TxtBagSize.SetFocus
     ElseIf .Col = 10 Or .Col = 14 Then
        If .TextMatrix(.row, 0) = strUnChecked Then Exit Sub
        TxtBagSize.Visible = False
        TxtFinalPrice.Visible = True
        TxtFinalPrice.Left = .Left + .CellLeft
        TxtFinalPrice.Top = .Top + .CellTop
        TxtFinalPrice.Width = .CellWidth
        TxtFinalPrice.Height = .CellHeight - 10
        TxtFinalPrice.Text = .Text
        TxtFinalPrice.SelStart = 0
        TxtFinalPrice.SelLength = Len(.Text)
        TxtFinalPrice.SetFocus
     Else
         TxtFinalPrice.Visible = False
         TxtBagSize.Visible = False
     End If
End With
End If
End Sub

Private Sub MSHFlexGrid4_Scroll()
    If TxtFinalPrice.Visible = True Then TxtFinalPrice.Visible = False
    If TxtBagSize.Visible = True Then TxtBagSize.Visible = False
End Sub

Private Sub MSHFlexGrid5_Click()
Dim ans, r As Variant
Dim ID As Integer
With MSHFlexGrid5
r = .RowSel
If r = .Rows - 1 Then
    Call ClearFrameAttachment
    Exit Sub
End If
If r < 1 Then Exit Sub
If .TextMatrix(r, 1) = "" Then Exit Sub

If SaveCmd.Enabled = False Then
   ans = MsgBox("Do You want to view this document ", vbYesNo)
   If ans = vbYes Then
      Shell "Explorer.exe " & File_Path, vbMaximizedFocus
        'If Right(App.Path, 1) = "\" Then
        '   Shell "Explorer.exe " & File_Path, vbMaximizedFocus
        'Else
        '   Shell "Explorer.exe " & App.Path & "\Withdrawal\", vbMaximizedFocus
        'End If
   End If
   Exit Sub
End If

lblAttachment.Caption = .TextMatrix(r, 3)
TxtAttachmentRemarks = .TextMatrix(r, 2)
CmdRemoveAttachment.Enabled = True
'If MSHFlexGrid12.TextMatrix(r, 4) <> "" Then
'   CmdRemoveAttachment.Enabled = False
'End If
CmdAddAttachment.Caption = "Update Attachment"
End With
End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsTxn_DepositContract.MoveNext
If RsTxn_DepositContract.EOF Then
   RsTxn_DepositContract.MoveLast
   LF_Flag = "Y"
End If
Call Indata
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
End Sub

Private Sub PreviousCmd_Click()

Dim LF_Flag As Variant
LF_Flag = "N"
RsTxn_DepositContract.MovePrevious
If RsTxn_DepositContract.BOF Then
   RsTxn_DepositContract.MoveFirst
   LF_Flag = "Y"
End If
Call Indata
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub

Private Sub SaveCmd_Click()
Dim mChkApproval As String
Dim mR, mIndex As Variant

''-- for Concurrency Check Only for Update
If Edit_Flg = "E" Then
   tmp = " Select * from Txn_DepositContract Where DepConID = " & Val(TxtDepositContractID) & " And CONVERT(varchar(30),Concurrency,120) > '" & Format(RsTxn_DepositContract!Concurrency, "YYYY-MM-DD hh:mm:ss") & "' "
Call TmpTable
If TmpRs.EOF = False Then
    MsgBox "Sorry, System not allow to update the data because same data in updated by another user." & vbNewLine & vbNewLine & "Please reopen your form to get update information and process.", vbInformation + vbOKOnly
    Exit Sub
End If
End If
''--------------------------------

If CmbState.Text = "" Then
   MsgBox "Blank State Not Allowed.", vbOKOnly + vbExclamation, App.Title
   CmbState.SetFocus
   Exit Sub
End If
If CmbLocation.Text = "" Then
   MsgBox "Blank Location Not Allowed.", vbOKOnly + vbExclamation, App.Title
   CmbLocation.SetFocus
   Exit Sub
End If
If CmbClient.Text = "" Then
   MsgBox "Blank Client Not Allowed.", vbOKOnly + vbExclamation, App.Title
   CmbClient.SetFocus
   Exit Sub
End If
If CmbStatus.Text = "" Then
   MsgBox "Blank Status Not Allowed.", vbOKOnly + vbExclamation, App.Title
   CmbStatus.SetFocus
   Exit Sub
End If
If Edit_Flg = "A" Then
    
''        If LCase(mDepositContractInFutureDateFlag) = "n" Then
''         If CDate(TxtStartDate) > CDate(Date) Then
''          MsgBox "New Deposit Contract can not create in Future Date.", vbInformation + vbOKOnly
''          TxtStartDate.SetFocus
''          Exit Sub
''         End If
''        End If
''        ''-- Rights applyed on 05 Jan 2018
''        ''If LCase(Gen_UserRole) <> "superadmin" And LCase(Gen_UserRole) <> "head cmg" And LCase(Gen_UserRole) <> "power" Then
''        If LCase(Gen_UserRole) <> "power" And mApprovedRights = False Then
''           If CDate(TxtStartDate) < CDate(Date) - mDepositContractAllowBackDateEntryDays Then
''               MsgBox "Start date should not prior more then " & mDepositContractAllowBackDateEntryDays & " day.", vbInformation + vbOKOnly
''               TxtStartDate.SetFocus
''               Exit Sub
''           End If
''        End If
    
    ''-- 14 Feb 2018
    If LCase(CmbContractType.Text) = "normal" Then
        If LCase(mDepositContractInFutureDateFlag) = "n" Then
         If CDate(TxtStartDate) > CDate(Date) Then
          MsgBox "New Deposit Contract can not create in Future Date.", vbInformation + vbOKOnly
          TxtStartDate.SetFocus
          Exit Sub
         End If
        End If
        ''-- Rights applyed on 05 Jan 2018
        ''If LCase(Gen_UserRole) <> "superadmin" And LCase(Gen_UserRole) <> "head cmg" And LCase(Gen_UserRole) <> "power" Then
        If LCase(Gen_UserRole) <> "power" And mApprovedRights = False Then
           If CDate(TxtStartDate) < CDate(Date) - mDepositContractAllowBackDateEntryDays Then
               MsgBox "Start date should not prior more then " & mDepositContractAllowBackDateEntryDays & " day.", vbInformation + vbOKOnly
               TxtStartDate.SetFocus
               Exit Sub
           End If
        End If
    ElseIf LCase(CmbContractType.Text) = "revision" Then
       If CDate(TxtStartDate) >= CDate(Date) Then
            MsgBox "Start Date should be past date for Revision Deposit Contract.", vbInformation + vbOKOnly
            TxtStartDate.SetFocus
            Exit Sub
       End If
    End If
   ''----------
End If

''--14 Feb 2018
If CmbContractType.Text = "" Then
   MsgBox "Blank Contract Type Not Allowed.", vbOKOnly + vbExclamation, App.Title
   CmbContractType.SetFocus
   Exit Sub
End If
''--------------

If CDate(TxtStartDate) > CDate(TxtEndDate) Then
   MsgBox "Start date should be less than End date.", vbOKOnly + vbExclamation, App.Title
   TxtEndDate.SetFocus
   Exit Sub
End If
If CmbInsurance.Text = "" Then
   MsgBox "Blank Insurance Not Allowed.", vbOKOnly + vbExclamation, App.Title
   CmbInsurance.SetFocus
   Exit Sub
End If
If TxtTentativeDuration.Text = "" Then
   MsgBox "Blank Tentative Duration of Storage (in months) Not Allowed.", vbOKOnly + vbExclamation, App.Title
   TxtTentativeDuration.SetFocus
   Exit Sub
End If
If TxtLockInPeriod.Text = "" Then
   MsgBox "Blank Lock-in Period (in months) Not Allowed.", vbOKOnly + vbExclamation, App.Title
   TxtLockInPeriod.SetFocus
   Exit Sub
End If
If TxtQuantum.Text = "" Then
   MsgBox "Blank Quantum Not Allowed.", vbOKOnly + vbExclamation, App.Title
   TxtQuantum.SetFocus
   Exit Sub
End If
If CmbBillingCycle.Text = "" Then
   MsgBox "Blank Billing Cycle Not Allowed.", vbOKOnly + vbExclamation, App.Title
   CmbBillingCycle.SetFocus
   Exit Sub
End If
If CmbDepositBillingCycle.Text = "" Then
   MsgBox "Blank Deposit Billing Cycle Not Allowed.", vbOKOnly + vbExclamation, App.Title
   CmbDepositBillingCycle.SetFocus
   Exit Sub
End If
If CmbRatePer.Text = "" Then
   MsgBox "Blank Rate Per(Excl. S.T.) Not Allowed.", vbOKOnly + vbExclamation, App.Title
   CmbRatePer.SetFocus
   Exit Sub
End If
If CmbForwardedTo.Text = "" Then
   MsgBox "Blank Forwarded To Not Allowed.", vbOKOnly + vbExclamation, App.Title
   CmbForwardedTo.SetFocus
   Exit Sub
End If

If Trim(TxtRemarks) = "" Then
   MsgBox "Blank Remarks Not Allowed.", vbOKOnly + vbExclamation, App.Title
   TxtRemarks.SetFocus
   Exit Sub
End If
If Edit_Flg = "A" Then
   If Val(mBagEquivalent) = 0 Then
    MsgBox "Bag Equivalent is not available, Please contact CMG Team for Commodity Bag Equivalent."
    Exit Sub
   End If
End If


If LCase(CmbStatus.Text) <> "rejected" Then
    If Check4.Value = 1 Then
       If CheckGodownSelectOrNot(MSHFlexGrid4) = False Then
           MsgBox "Please select atleast one godown and define Billing Rate & Deposite/Withdrawal Rate for same.", vbOKOnly + vbInformation, App.Title
           If Edit_Flg = "A" Then Check2.SetFocus
           Exit Sub
       End If
       If CheckForBagSize(MSHFlexGrid4) = False Then
          MsgBox "Bag Size is not given for All selected Godowns.", vbOKOnly + vbInformation, App.Title
          MSHFlexGrid4.SetFocus
          Exit Sub
       End If
       If CheckGodownForRate(MSHFlexGrid4) = False Then
           MsgBox "Billing Rate\Deposite Rate are not define.", vbOKOnly + vbInformation, App.Title
           MSHFlexGrid4.SetFocus
           Exit Sub
       End If
    
    Else
       If CheckGodownSelectOrNot(MSHFlexGrid1) = False Then
           MsgBox "Please select atleast one godown and define Billing Rate & Deposite/Withdrawal Rate for same.", vbOKOnly + vbInformation, App.Title
           If Edit_Flg = "A" Then Check2.SetFocus
           Exit Sub
       End If
       If CheckGodownForRate(MSHFlexGrid1) = False Then
           MsgBox "Billing Rate\Deposite Rate are not define.", vbOKOnly + vbInformation, App.Title
           MSHFlexGrid1.SetFocus
           Exit Sub
       End If
    
    End If
End If
''-- Added on 06 Jan 2018 for Attachment
If LCase(mDepositContractDefaultAttachmentFlag) = LCase("y") Then
   If LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "headcmg" And LCase(Gen_UserRole) <> "superadmin" And LCase(Gen_UserRole) <> "admin" Then
'      If Edit_Flg = "A" Then
        With MSHFlexGrid5
            If .Rows <= 2 Then
                MsgBox "Blank Attachment not allowed."
                SSTab1.Tab = 1
                Exit Sub
            End If
        End With
'      End If
   End If
End If
''---------------------------------------

''-- Sujjested by Sankar Sir on 18 Nov 2017
If Edit_Flg = "A" Then
    Select Case MsgBox("Do you confrim with given Billing Rate/MT (Proposed/Approved) and Deposit/Withdrawal (Proposed/Approved) for listed godowns?" & vbNewLine & vbNewLine & "Click 'No' to cross check with proposed rate or Click 'Yes' to save otherwise.", vbYesNo + vbQuestion)
        Case vbNo
            ''Call cmdExportDetails_Click
            Exit Sub
    End Select
End If
''-----

''-- Check for godown List that the already DC is available or not
mSelectGodownID = "0"
If Check4.Value = 1 Then
   With MSHFlexGrid4
    If .Rows >= 2 Then
       For mR = 1 To .Rows - 1
         If .TextMatrix(mR, 0) = strChecked Then
            mSelectGodownID = mSelectGodownID & "," & .TextMatrix(mR, 11)
         End If
       Next mR
    End If
   End With
Else
   With MSHFlexGrid1
    If .Rows >= 2 Then
       For mR = 1 To .Rows - 1
         If .TextMatrix(mR, 0) = strChecked Then
            mSelectGodownID = mSelectGodownID & "," & .TextMatrix(mR, 9)
         End If
       Next mR
    End If
   End With
End If

''--And TDC.StateID=" & mStateID & "
tmp = " SELECT TDCGD.GodownId,TDCGD.DepConID,TDC.StartDate,TDC.EndDate FROM"
tmp = tmp & " Txn_DepositContractGodownDetail AS TDCGD Inner Join Txn_DepositContract TDC On TDC.DepConID = TDCGD.DepConID"
tmp = tmp & " where TDC.ApprovalFlag Not In ('R','C') And TDC.DepConID<>" & Val(TxtDepositContractID) & " And TDC.LocationID=" & mLocationID & " And TDC.CommodityID=" & mCommodityID & " And TDC.ClientID=" & mClientIDRow & ""
tmp = tmp & " And (('" & TxtStartDate & "' >= StartDate And '" & TxtStartDate & "' <= EndDate) Or '" & TxtEndDate & "' >= StartDate And '" & TxtEndDate & "' <= EndDate Or ('" & TxtStartDate & "' <= StartDate And '" & TxtEndDate & "' >= EndDate))"
''-- 14 Feb 2018
tmp = tmp & " And ContractType = '" & mContractTypeFlag & "'"
''--------------
tmp = tmp & " AND TDCGD.GodownID IN (" & mSelectGodownID & ")"
Call TmpTable
If TmpRs.RecordCount > 0 Then
   MsgBox "Deposit Contract already exist for selected godowns on given start date and end date!!!"
   If TxtStartDate.Enabled = True Then TxtStartDate.SetFocus
   Exit Sub
End If
''----------------------------------

'''''--- Logic to Forward to for Approval
''-- CheckforApproval
''-- if string content "C"-CMG for Approval "U"-Upper Approval "S"-System Approval "N"-Nothing
mFinalStatus = ""

If Check4.Value = 1 Then
   mChkApproval = CheckforApproval(MSHFlexGrid4)
Else
   mChkApproval = CheckforApproval(MSHFlexGrid1)
End If

If InStr(UCase(mChkApproval), "C") > 0 Then
    mFinalStatus = "C"
ElseIf InStr(UCase(mChkApproval), "U") > 0 Then
    mFinalStatus = "U"
ElseIf InStr(UCase(mChkApproval), "S") > 0 Then
    mFinalStatus = "S"
Else
    mFinalStatus = "S"
End If

If Edit_Flg = "E" Then
    ''-- Rights applyed on 05 Jan 2018
    ''If LCase(Gen_UserRole) = "superadmin" Or LCase(Gen_UserRole) = "head cmg" Or LCase(Gen_UserRole) = "power" Then
    If LCase(Gen_UserRole) = "power" Or mApprovedRights = True Then
           If mFinalStatus = "C" And UCase(Left(CmbStatus.Text, 1)) = "A" Then
           
           ElseIf mFinalStatus = "S" And UCase(Left(CmbStatus.Text, 1)) = "F" Then
                CmbStatus.Text = "Approved"
           ElseIf mFinalStatus = "U" And UCase(Left(CmbStatus.Text, 1)) = "F" Then
                CmbStatus.Text = "Forwarded"
           ElseIf mFinalStatus = "C" And UCase(Left(CmbStatus.Text, 1)) = "F" Then
                CmbStatus.Text = "Forwarded"
           End If
    Else
        If LoginEmployeeInHierarchy(Gen_UserLoginID) = True Then
           If mFinalStatus = "C" And UCase(Left(CmbStatus.Text, 1)) = "A" Then
                CmbStatus.Text = "Forwarded"
                MsgBox "Deposite Contract will save with updated Rate but not Approved, Please contact to CMG Team for Approval"
           ElseIf mFinalStatus = "S" And UCase(Left(CmbStatus.Text, 1)) = "F" Then
                CmbStatus.Text = "Approved"
           ElseIf mFinalStatus = "U" And UCase(Left(CmbStatus.Text, 1)) = "F" Then
                CmbStatus.Text = "Forwarded"
           End If
        End If
    End If

End If
'''''------------------------------------

If Edit_Flg = "A" Then
   If RsTxn_DepositContract.RecordCount > 0 Then
      RsTxn_DepositContract.MoveFirst
   End If
   RsTxn_DepositContract.AddNew
   RsTxn_DepositContract!DepConID = GetMaxID("DepConID", "Txn_DepositContract")
   RsTxn_DepositContract!G_UID = GetGUID
   TxtDepositContractID.Text = RsTxn_DepositContract!DepConID
   RsTxn_DepositContract!ContractType = mContractTypeFlag  ''-- 14 Feb 2018
Else
   If LCase(CmbStatus.Text) = "approved" Then
      If mCreatedID = mForwardID Then
         MsgBox "Created by and Approve by should be differnt!!!"
         Exit Sub
      End If
      If mCreatedID = Gen_EmployeeID Then
         MsgBox "Created by and Approved by should be differnt!!!"
         Exit Sub
      End If
   End If
   
   If RsTxn_DepositContract.RecordCount > 0 Then RsTxn_DepositContract.MoveFirst
   RsTxn_DepositContract.MoveFirst
   RsTxn_DepositContract.Find " DepConID= " & TxtDepositContractID.Text
   RsTxn_DepositContract!Concurrency = Now()            ''-- for Concurrency check
End If

RsTxn_DepositContract!DepConID = TxtDepositContractID.Text
RsTxn_DepositContract!Date = TxtDepositContractDate
RsTxn_DepositContract!StateID = mStateID
RsTxn_DepositContract!LocationID = mLocationID
RsTxn_DepositContract!CommodityID = mCommodityID
RsTxn_DepositContract!ClientID = mClientIDRow
RsTxn_DepositContract!QuantumMT = TxtQuantum.Text
RsTxn_DepositContract!InsuranceByID = mInsuranceByID
RsTxn_DepositContract!StartDate = TxtStartDate
RsTxn_DepositContract!EndDate = TxtEndDate
RsTxn_DepositContract!BillingCycleID = mBillingCycleID
RsTxn_DepositContract!DepBillingCycleID = mDepositBillingCycleID
RsTxn_DepositContract!WBillingCycleID = mDepositBillingCycleID  ''-- Withdrawal Billing Cycle for Future if required
RsTxn_DepositContract!RatePer = UCase(Left(CmbRatePer.Text, 1))
RsTxn_DepositContract!TentativeDur = TxtTentativeDuration.Text
RsTxn_DepositContract!LockPeriod = TxtLockInPeriod.Text
RsTxn_DepositContract!BagSizeFlag = IIf(Check4.Value = 1, "Y", "N")
RsTxn_DepositContract!SM_Prefix = mSM_Prefix

If IsNull(RsTxn_DepositContract!CreatedDate) Then
   RsTxn_DepositContract!CreatedDate = Now
   RsTxn_DepositContract!CreatedBy = Gen_UserLoginID
Else
   RsTxn_DepositContract!UpdatedDate = Now
   RsTxn_DepositContract!UpdatedBy = Gen_UserLoginID
End If

If Edit_Flg = "A" Then
    ''-- Rights applyed on 05 Jan 2018
    ''If LCase(Gen_UserRole) = "superadmin" Or LCase(Gen_UserRole) = "head cmg" Or LCase(Gen_UserRole) = "power" Then
    If LCase(Gen_UserRole) = "power" Or mApprovedRights = True Then
        If UCase(Left(CmbStatus.Text, 1)) <> "A" Then
            If UCase(Left(CmbStatus.Text, 1)) <> "F" Then
                RsTxn_DepositContract!ApprovalFlag = "F"
            Else
                'RsTxn_DepositContract!ApprovalFlag = IIf(mFinalStatus = "C", "A", IIf(mFinalStatus = "U", "A", "S"))
                RsTxn_DepositContract!ApprovalFlag = UCase(Left(CmbStatus.Text, 1))
            End If
        Else
            RsTxn_DepositContract!ApprovalFlag = "A"
        End If
    Else
        If LCase(Gen_UserRole) = "account" Then
           RsTxn_DepositContract!ApprovalFlag = "F" 'Commented on 04 Dec 2017 after approval from Sankar Sir and Mayur Sir for temprory basis IIf(mFinalStatus = "C", "F", IIf(mFinalStatus = "U", "F", "S"))
        Else
           RsTxn_DepositContract!ApprovalFlag = IIf(mFinalStatus = "C", "F", IIf(mFinalStatus = "U", "F", "S"))
        End If
    End If
    RsTxn_DepositContract!ApprovalRemarks = TxtRemarks.Text
    RsTxn_DepositContract!ApproveOn = Now
    RsTxn_DepositContract!ApproveBy = Gen_UserLoginID
    RsTxn_DepositContract!EmployeeID = mForwardID
Else
   If LCase(Gen_UserRole) = "account" Then
      RsTxn_DepositContract!ApprovalFlag = "F"
   Else
      RsTxn_DepositContract!ApprovalFlag = UCase(Left(CmbStatus.Text, 1))
   End If
   RsTxn_DepositContract!ApprovalRemarks = TxtRemarks.Text
   RsTxn_DepositContract!ApproveOn = Now
   RsTxn_DepositContract!ApproveBy = Gen_UserLoginID
   RsTxn_DepositContract!EmployeeID = mForwardID
End If

RsTxn_DepositContract.Update
Call SaveDepositContractGodownDetail
Call SaveAttachment
Call Gen_Email
RsTxn_DepositContract.Requery
RsTxn_DepositContract.Find "DepConID =" & Val(TxtDepositContractID)
'RsTxn_DepositContract.MoveLast
Call Indata
End Sub

Private Sub SaveAttachment()
Dim mAttachmentID As Variant
Dim mR, mC As Double
With MSHFlexGrid5
    If .Rows <= 1 Then Exit Sub
    Call TableTxn_DepositContractAttachment(" Where DepConID = " & Val(TxtDepositContractID) & "")
    mAttachmentID = "0"
    For mR = 1 To .Rows - 1
        If .TextMatrix(mR, 1) = "" Then Exit For
         If .TextMatrix(mR, 0) = "" Then
             If .TextMatrix(mR, 4) = "" Then
                RsTxn_DepositContractAttachment.AddNew
                RsTxn_DepositContractAttachment!AttachmentID = GetMaxAttachID
                .TextMatrix(mR, 4) = RsTxn_DepositContractAttachment!AttachmentID
                RsTxn_DepositContractAttachment!G_UID = GetGUID
                RsTxn_DepositContractAttachment!DepConID = Val(TxtDepositContractID)
             Else
                RsTxn_DepositContractAttachment.MoveFirst
                RsTxn_DepositContractAttachment.Find "AttachmentID = " & .TextMatrix(mR, 4) & ""
             End If
            
             mFileFlag = SaveFileToDB(.TextMatrix(mR, 3), RsTxn_DepositContractAttachment, "File")
             mFileFlag = GetFileFromPath(.TextMatrix(mR, 3))
             If RsTxn_DepositContractAttachment!FileName <> mFileFlag Then
                RsTxn_DepositContractAttachment!FileName = TxtDepositContractID & "-" & .TextMatrix(mR, 4) & "-" & mFileFlag
             Else
                If IsNull(RsTxn_DepositContractAttachment!FileName) Then
                    RsTxn_DepositContractAttachment!FileName = TxtDepositContractID & "-" & .TextMatrix(mR, 4) & "-" & mFileFlag
                Else
                    RsTxn_DepositContractAttachment!FileName = mFileFlag
                End If
             End If
             RsTxn_DepositContractAttachment!Remarks = Trim(.TextMatrix(mR, 2))
             If IsNull(RsTxn_DepositContractAttachment!CreatedBy) = True Or RsTxn_DepositContractAttachment!CreatedBy = "" Then
                RsTxn_DepositContractAttachment!CreatedBy = Gen_UserLoginID
                RsTxn_DepositContractAttachment!CreatedDate = Now
             Else
                RsTxn_DepositContractAttachment!UpdatedBy = Gen_UserLoginID
                RsTxn_DepositContractAttachment!UpdatedDate = Now
             End If
             RsTxn_DepositContractAttachment.Update
         End If
        mAttachmentID = mAttachmentID & "," & .TextMatrix(mR, 4)
    Next
    
    tmp = " DELETE FROM Txn_DepositContractAttachment WHERE DepConID = " & Val(TxtDepositContractID) & " AND AttachmentID NOT IN (" & mAttachmentID & ") "
    Call TmpTableSMAttachment
    
End With
End Sub

Private Sub InDataAttachment()
Dim mR, mC As Double
Call Grid_HeadAttachment
Call ClearFrameAttachment
Call TableTxn_DepositContractAttachment(" Where DepConID = " & Val(TxtDepositContractID) & "")
If RsTxn_DepositContractAttachment.RecordCount > 0 Then
   
    If Right(App.Path, 1) = "\" Then
       File_Path = App.Path & "Agreement Attachment\Deposit Contract Documents\" & Gen_UserLoginID
    Else
       File_Path = App.Path & "\Agreement Attachment\Deposit Contract Documents\" & Gen_UserLoginID
    End If
    If oFs.FolderExists(File_Path) = True Then
       Set oFolder = oFs.GetFolder(File_Path)
       For Each oFile In oFolder.Files
           Call oFs.DeleteFile(oFile.Path, True)
       Next
    Else
       If Not oFs.FolderExists(File_Path) Then
          Call oFs.CreateFolder(File_Path)
       End If
    End If
   
   For mR = 1 To RsTxn_DepositContractAttachment.RecordCount
   With MSHFlexGrid5
        .Rows = .Rows + 1
        .TextMatrix(mR, 0) = mR
        .TextMatrix(mR, 1) = IIf(IsNull(RsTxn_DepositContractAttachment!FileName), "", RsTxn_DepositContractAttachment!FileName)
        .TextMatrix(mR, 2) = IIf(IsNull(RsTxn_DepositContractAttachment!Remarks), "", RsTxn_DepositContractAttachment!Remarks)
        .TextMatrix(mR, 3) = IIf(IsNull(RsTxn_DepositContractAttachment!FileName), "", File_Path & "\" & RsTxn_DepositContractAttachment!FileName)
        Call LoadFileFromDB(.TextMatrix(mR, 3), RsTxn_DepositContractAttachment, "File")
        .TextMatrix(mR, 4) = IIf(IsNull(RsTxn_DepositContractAttachment!AttachmentID), "", RsTxn_DepositContractAttachment!AttachmentID)
        RsTxn_DepositContractAttachment.MoveNext
   End With
   Next mR
End If
FrameAttachment.Enabled = False
End Sub

Public Sub Indata()
mFlag = False

CmdCopy.Enabled = True
CmdClosed.Enabled = False
''-- 02 Feb 2018 -- Verified DC
Frame6.Visible = False
CmdVerified.Enabled = False
TxtVerifiedRemark = ""
''------------------------------

TxtState.Visible = True:                CmbState.Visible = False
TxtLocation.Visible = True:             CmbLocation.Visible = False
TxtCommodity.Visible = True:            CmbCommodity.Visible = False
TxtClient.Visible = True:               CmbClient.Visible = False
TxtBillingCycle.Visible = True:         CmbBillingCycle.Visible = False
TxtDepositBillingCycle.Visible = True:  CmbDepositBillingCycle.Visible = False
TxtRatePer.Visible = True:              CmbRatePer.Visible = False
TxtForwardedTo.Visible = True:          CmbForwardedTo.Visible = False
TxtStatus.Visible = True:               CmbStatus.Visible = False
TxtInsurance.Visible = True:            CmbInsurance.Visible = False
TxtContractType.Visible = True:         CmbContractType.Visible = False ''--- 14 Feb 2018

Frame1.Enabled = False
'Frame2.Enabled = False
Check2.Enabled = False
Frame3.Visible = False

tmp = "Select TDC.DepConID,TDC.Date,MS.StateName,ML.StateID,ML.SM_Prefix,ML.LocationName,TDC.LocationID,MC.ClientName,MC.ClientIDRow,MC.ClientID, "
tmp = tmp & " MC.Address,MC.MobileNo,MC.EmailID,MC.TelephoneNo,TDC.QuantumMT,MIB.InsuranceByName,TDC.InsuranceByID,TDC.StartDate,TDC.EndDate,"
tmp = tmp & " MBC.BillingCycle As BillingCycle, MBCDep.BillingCycle As BillingCycleDep,TDC.RatePer,TDC.TentativeDur,TDC.LockPeriod,"  '' ,TDC.SRCID
tmp = tmp & " Mco.Commodity,MCo.FumigationFlag,MCo.Minimum,MCo.Maximum,TDC.CommodityID,TDC.DepBillingCycleID,TDC.BillingCycleID,MBC.Ratio,MBCDep.Ratio As DepositRatio,"
tmp = tmp & " ME.EmployeeName+'~'+ME.EmployeeNo As ForwardedTo,ME.EmployeeID,TDC.ApprovalFlag,TDC.ApprovalRemarks,TDC.BagSizeFlag,TDC.Verified,TDC.ContractType "
''--,MSRC.RateOnQty
tmp = tmp & " From Txn_DepositContract TDC"
''tmp = tmp & " Inner Join Mst_State MS on TDC.StateID=MS.StateID"
tmp = tmp & " Inner Join MST_Location ML on TDC.LocationID=ML.LocationID"
tmp = tmp & " Inner Join Mst_State MS on ML.StateID=MS.StateID"
tmp = tmp & " Inner Join MST_Commodity Mco on TDC.CommodityID=Mco.CommodityID"
tmp = tmp & " Inner Join MST_Client MC on TDC.ClientID=MC.ClientID"
tmp = tmp & " Inner Join MST_BillingCycle MBC on TDC.BillingCycleID=MBC.BillingCycleID"
tmp = tmp & " Inner Join MST_BillingCycle MBCDep on TDC.DepBillingCycleID=MBCDep.BillingCycleID"
tmp = tmp & " Inner Join Mst_InsuranceBy MIB on TDC.InsuranceByID=MIB.InsuranceByID"
tmp = tmp & " Left Join Mst_Employee ME on TDC.EmployeeID=ME.EmployeeID"
tmp = tmp & " Where TDC.DepConID=" & RsTxn_DepositContract!DepConID & ""
Call Tmp1Table
If TmpRs1.RecordCount > 0 Then
   TxtDepositContractID.Text = IIf(IsNull(TmpRs1!DepConID), "", TmpRs1!DepConID)
   TxtDepositContractDate = IIf(IsNull(TmpRs1!Date), "", TmpRs1!Date)
   TxtState.Text = IIf(IsNull(TmpRs1!StateName), "", TmpRs1!StateName)
   TxtLocation.Text = IIf(IsNull(TmpRs1!Locationname), "", TmpRs1!Locationname)
   TxtCommodity.Text = IIf(IsNull(TmpRs1!Commodity), "", TmpRs1!Commodity)
   TxtClient.Text = IIf(IsNull(TmpRs1!ClientName), "", TmpRs1!ClientName) & "#" & IIf(IsNull(TmpRs1!ClientIDRow), "", TmpRs1!ClientIDRow) & "#" & IIf(IsNull(TmpRs1!ClientID), "", TmpRs1!ClientID)
   TxtAddress.Text = IIf(IsNull(TmpRs1!Address), "", TmpRs1!Address)
   TxtContactDetails.Text = IIf(IsNull(TmpRs1!MobileNo), "", TmpRs1!MobileNo)
   TxtEmailId.Text = IIf(IsNull(TmpRs1!EmailID), "", TmpRs1!EmailID)
   TxtContactNo.Text = IIf(IsNull(TmpRs1!TelephoneNo), "", TmpRs1!TelephoneNo)
   TxtQuantum.Text = IIf(IsNull(TmpRs1!QuantumMT), "", TmpRs1!QuantumMT)
   TxtInsurance.Text = IIf(IsNull(TmpRs1!InsuranceByName), "", TmpRs1!InsuranceByName)
   TxtStartDate = IIf(IsNull(TmpRs1!StartDate), "", TmpRs1!StartDate)
   TxtEndDate = IIf(IsNull(TmpRs1!EndDate), "", TmpRs1!EndDate)
   TxtBillingCycle.Text = IIf(IsNull(TmpRs1!BillingCycle), "", TmpRs1!BillingCycle)
   CmbBillingCycle.Text = IIf(IsNull(TmpRs1!BillingCycle), "", TmpRs1!BillingCycle)
   TxtDepositBillingCycle.Text = IIf(IsNull(TmpRs1!BillingCycleDep), "", TmpRs1!BillingCycleDep)
   CmbDepositBillingCycle.Text = IIf(IsNull(TmpRs1!BillingCycleDep), "", TmpRs1!BillingCycleDep)
   TxtRatePer.Text = IIf(IsNull(TmpRs1!RatePer), "", TmpRs1!RatePer)
   TxtRatePer.Text = IIf((TmpRs1!RatePer) = "I", "Included", "Excluded")
   TxtTentativeDuration.Text = IIf(IsNull(TmpRs1!TentativeDur), "", TmpRs1!TentativeDur)
   TxtLockInPeriod.Text = IIf(IsNull(TmpRs1!LockPeriod), "", TmpRs1!LockPeriod)
   mStateID = IIf(IsNull(TmpRs1!StateID), "", TmpRs1!StateID)
   mLocationID = IIf(IsNull(TmpRs1!LocationID), "", TmpRs1!LocationID)
   mSM_Prefix = IIf(IsNull(TmpRs1!SM_Prefix), "", TmpRs1!SM_Prefix)
   mCommodityID = IIf(IsNull(TmpRs1!CommodityID), "", TmpRs1!CommodityID)
   mFumigationFlag = IIf(IsNull(TmpRs1!FumigationFlag), "N", TmpRs1!FumigationFlag)
   mBagEquivalent = GetCommodityBagEquivalent(mCommodityID)
   mClientIDRow = IIf(IsNull(TmpRs1!ClientID), "", TmpRs1!ClientID)
   mDepositBillingCycleID = IIf(IsNull(TmpRs1!DepBillingCycleID), "", TmpRs1!DepBillingCycleID)
   mBillingCycleID = IIf(IsNull(TmpRs1!BillingCycleID), "", TmpRs1!BillingCycleID)
   mInsuranceByID = IIf(IsNull(TmpRs1!InsuranceByID), "", TmpRs1!InsuranceByID)
   mMinBagSize = IIf(IsNull(TmpRs1!Minimum), 0, TmpRs1!Minimum)
   mMaxBagSize = IIf(IsNull(TmpRs1!Maximum), 0, TmpRs1!Maximum)
   mBillingRatio = IIf(IsNull(TmpRs1!Ratio), 0, TmpRs1!Ratio)
   mDepositBillingRatio = IIf(IsNull(TmpRs1!DepositRatio), 0, TmpRs1!DepositRatio)
   TxtForwardedTo.Text = IIf(IsNull(TmpRs1!ForwardedTo), "", TmpRs1!ForwardedTo)
   mForwardID = IIf(IsNull(TmpRs1!EmployeeID), 0, TmpRs1!EmployeeID)
   mStatus = IIf(IsNull(TmpRs1!ApprovalFlag), "", TmpRs1!ApprovalFlag)
   mVerifiyFlag = IIf(IsNull(TmpRs1!Verified), False, IIf(LCase(TmpRs1!Verified) = "y", True, False)) ''-- 02 Feb 2018
   'CreatedBy , CreatedDate, UpdatedBy, UpdatedDate, ApproveBy, ApproveOn, ClosedBy, ClosedOn
   TxtCreated = IIf(IsNull(RsTxn_DepositContract!CreatedBy), "", RsTxn_DepositContract!CreatedBy) & " :- " & IIf(IsNull(RsTxn_DepositContract!CreatedDate), "", RsTxn_DepositContract!CreatedDate)
   TxtUpdated = IIf(IsNull(RsTxn_DepositContract!UpdatedBy), "", RsTxn_DepositContract!UpdatedBy) & " :- " & IIf(IsNull(RsTxn_DepositContract!UpdatedDate), "", RsTxn_DepositContract!UpdatedDate)
   If LCase(mStatus) = "s" Or LCase(mStatus) = "a" Then
    TxtApproved = IIf(IsNull(RsTxn_DepositContract!ApproveBy), "", RsTxn_DepositContract!ApproveBy) & " :- " & IIf(IsNull(RsTxn_DepositContract!ApproveOn), "", RsTxn_DepositContract!ApproveOn)
   Else
    TxtApproved = ""
   End If
   If LCase(mStatus) = "c" Then
     TxtClosed = IIf(IsNull(RsTxn_DepositContract!ClosedBy), "", RsTxn_DepositContract!ClosedBy) & " :- " & IIf(IsNull(RsTxn_DepositContract!ClosedOn), "", RsTxn_DepositContract!ClosedOn)
   Else
     TxtClosed = ""
   End If
   
   If LCase(mStatus) = "f" Then
      TxtStatus.Text = "Forwarded"
   ElseIf LCase(mStatus) = "s" Then
      TxtStatus.Text = "System Approved"
   ElseIf LCase(mStatus) = "r" Then
      TxtStatus.Text = "Rejected"
   ElseIf LCase(mStatus) = "a" Then
      TxtStatus.Text = "Approved"
   ElseIf LCase(mStatus) = "c" Then
      TxtStatus.Text = "Closed"
   End If
   
   ''--- 14 Feb 2018
   mContractTypeFlag = IIf(IsNull(TmpRs1!ContractType), "", TmpRs1!ContractType)
   TxtContractType = IIf(LCase(mContractTypeFlag) = "n", "Normal", IIf(LCase(mContractTypeFlag) = "r", "Revision", "Normal"))
   ''-----------------
   
   TxtRemarks.Text = IIf(IsNull(TmpRs1!ApprovalRemarks), "", TmpRs1!ApprovalRemarks)
   mBagSizeFlag = IIf(IsNull(TmpRs1!BagSizeFlag), "", TmpRs1!BagSizeFlag)
   
   Check4.Tag = "Noclick"
   Check4.Value = vbUnchecked
   Check4.Value = IIf(mBagSizeFlag = "Y", 1, 0)
   Check4.Tag = ""
   
End If

Call IndataDepositContractGodownDetail
Call InDataLogHistory
Call InDataAttachment
Call IndataOrignalData

CmbCMPID.Clear
CmbGodownID.Clear
TxtMinNewBagSize = ""
TxtMaxNewBagSize = ""


''If Check4.Value = 1 Then
''   Call CheckBaseRateBags
''Else
''   Call CheckBaseRate
''End If
SSTab1.Tab = 0

mUpperHierarchyEmployeeNoList = ""
Call GetUpperHierarchyEmployeeNoList(IIf(IsNull(RsTxn_DepositContract!CreatedBy), "", RsTxn_DepositContract!CreatedBy))


CmdClosed.Enabled = False
If LCase(Gen_UserRole) = "headcmg" Or LCase(Gen_UserRole) = "power" Then
    If LCase(mStatus) <> "c" And LCase(mStatus) <> "r" And LCase(mStatus) <> "f" Then
        CmdClosed.Enabled = True
    End If
End If

''-- 02 Feb 2018 Verification
CmdVerified.Enabled = False
If LCase(Gen_UserRole) = "headcmg" Or LCase(Gen_UserRole) = "superadmin" Or LCase(Gen_UserRole) = "power" Then
    If LCase(mStatus) <> "c" And LCase(mStatus) <> "r" And LCase(mStatus) <> "f" Then
        If mVerifiyFlag = False Then
            CmdVerified.Enabled = True
        End If
    End If
End If
''---------------------------

If RsTxn_DepositContract.RecordCount = 1 Then
   Call GButtonLock_1(Me, True)
Else
   Call GButtonLock(Me, True)
End If

If AddCmd.Enabled = False Then
   CmdCopy.Enabled = False
End If

End Sub

Public Sub IndataOrignalData()
Dim mR1 As Double, mC1 As Double
Dim mR2 As Double, mC2 As Double
    Call Grid_Head6
    If LCase(mBagSizeFlag) = "y" Then
        With MSHFlexGrid6
            For mR1 = 0 To MSHFlexGrid4.Rows - 1
                .Rows = MSHFlexGrid4.Rows
                For mC1 = 0 To MSHFlexGrid4.Cols - 1
                    .Cols = MSHFlexGrid4.Cols
                    .TextMatrix(mR1, mC1) = MSHFlexGrid4.TextMatrix(mR1, mC1)
                    MSHFlexGrid4.Col = mC1: MSHFlexGrid4.row = mR1
                    .Col = mC1: .row = mR1
                    .CellBackColor = MSHFlexGrid4.CellBackColor
                Next mC1
            Next mR1
        End With
    Else
        With MSHFlexGrid6
            For mR1 = 0 To MSHFlexGrid1.Rows - 1
                .Rows = MSHFlexGrid1.Rows
                For mC1 = 0 To MSHFlexGrid1.Cols - 1
                    .Cols = MSHFlexGrid1.Cols
                    .TextMatrix(mR1, mC1) = MSHFlexGrid1.TextMatrix(mR1, mC1)
                    MSHFlexGrid1.Col = mC1: MSHFlexGrid1.row = mR1
                    .Col = mC1: .row = mR1
                    .CellBackColor = MSHFlexGrid1.CellBackColor
                Next mC1
            Next mR1
        End With
    End If
End Sub

Public Sub SaveDepositContractGodownDetail()
Dim mDepConDIDList As Variant

Call TableTxn_DepositContractGodownDetail(" Where DepConID=" & Val(TxtDepositContractID.Text) & "")
mDepConDIDList = "0"
If Check4.Value = 0 Then    ''-- Quantity Wise -----------
   For I = 1 To MSHFlexGrid1.Rows - 1
       If MSHFlexGrid1.TextMatrix(I, 0) = strChecked Then
            If Val(MSHFlexGrid1.TextMatrix(I, 10)) = 0 Then
                RsTxn_DepositContractGodownDetail.AddNew
                RsTxn_DepositContractGodownDetail!DepConDID = GetMaxID("DepConDID", "Txn_DepositContractGodownDetail")
                MSHFlexGrid1.TextMatrix(I, 10) = RsTxn_DepositContractGodownDetail!DepConDID
                RsTxn_DepositContractGodownDetail!G_UID = GetGUID
                RsTxn_DepositContractGodownDetail!DepConID = Val(TxtDepositContractID.Text)
                RsTxn_DepositContractGodownDetail!GodownID = Val(MSHFlexGrid1.TextMatrix(I, 9))
            Else
                RsTxn_DepositContractGodownDetail.Find " DepConDID = " & Val(MSHFlexGrid1.TextMatrix(I, 10))
            End If
            RsTxn_DepositContractGodownDetail!FinalPrice = Val(MSHFlexGrid1.TextMatrix(I, 8))
            RsTxn_DepositContractGodownDetail!DepositFinalPrice = Val(MSHFlexGrid1.TextMatrix(I, 12))
            RsTxn_DepositContractGodownDetail!BaseRate = Val(MSHFlexGrid1.TextMatrix(I, 13))
            RsTxn_DepositContractGodownDetail!DepositBaseRate = Val(MSHFlexGrid1.TextMatrix(I, 14))
            RsTxn_DepositContractGodownDetail!OldBillingRate = Val(MSHFlexGrid1.TextMatrix(I, 7))
            RsTxn_DepositContractGodownDetail!OldDepositBillingRate = Val(MSHFlexGrid1.TextMatrix(I, 11))
            ''-- Withdrawal Billing Cycle Rate for Future if required
            RsTxn_DepositContractGodownDetail!WithdrawalFinalPrice = Val(MSHFlexGrid1.TextMatrix(I, 12))
            RsTxn_DepositContractGodownDetail!WithdrawalBaseRate = Val(MSHFlexGrid1.TextMatrix(I, 14))
            RsTxn_DepositContractGodownDetail!OldWithdrawalBillingRate = Val(MSHFlexGrid1.TextMatrix(I, 11))
            
            MSHFlexGrid1.row = I: MSHFlexGrid1.Col = 8
            If MSHFlexGrid1.CellBackColor = vbRed Then
                RsTxn_DepositContractGodownDetail!BaseRateColor = "R"
            ElseIf MSHFlexGrid1.CellBackColor = vbGreen Then
                RsTxn_DepositContractGodownDetail!BaseRateColor = "G"
            ElseIf MSHFlexGrid1.CellBackColor = vbYellow Then
                RsTxn_DepositContractGodownDetail!BaseRateColor = "Y"
            Else
                RsTxn_DepositContractGodownDetail!BaseRateColor = "W"
            End If
            MSHFlexGrid1.row = I: MSHFlexGrid1.Col = 12
            If MSHFlexGrid1.CellBackColor = vbRed Then
                RsTxn_DepositContractGodownDetail!DepositBaseRateColor = "R"
                RsTxn_DepositContractGodownDetail!WithdrawalBaseRateColor = "R"
            ElseIf MSHFlexGrid1.CellBackColor = vbGreen Then
                RsTxn_DepositContractGodownDetail!DepositBaseRateColor = "G"
                RsTxn_DepositContractGodownDetail!WithdrawalBaseRateColor = "G"
            ElseIf MSHFlexGrid1.CellBackColor = vbYellow Then
                RsTxn_DepositContractGodownDetail!DepositBaseRateColor = "Y"
                RsTxn_DepositContractGodownDetail!WithdrawalBaseRateColor = "Y"
            Else
                RsTxn_DepositContractGodownDetail!DepositBaseRateColor = "W"
                RsTxn_DepositContractGodownDetail!WithdrawalBaseRateColor = "W"
            End If
            ''--------------------------------------------------
            RsTxn_DepositContractGodownDetail!SM_Prefix = mSM_Prefix
            RsTxn_DepositContractGodownDetail!Rent = Val(MSHFlexGrid1.TextMatrix(I, 5))
            RsTxn_DepositContractGodownDetail!RatePerMonthPerMT = Val(MSHFlexGrid1.TextMatrix(I, 16))
            If IsNull(RsTxn_DepositContractGodownDetail!CreatedBy) = True Or RsTxn_DepositContractGodownDetail!CreatedBy = "" Then
                RsTxn_DepositContractGodownDetail!CreatedBy = Gen_UserLoginID
                RsTxn_DepositContractGodownDetail!CreatedDate = Now
            Else
                RsTxn_DepositContractGodownDetail!UpdatedBy = Gen_UserLoginID
                RsTxn_DepositContractGodownDetail!UpdatedDate = Now
            End If
            RsTxn_DepositContractGodownDetail.Update
            RsTxn_DepositContractGodownDetail.Requery
            mDepConDIDList = MSHFlexGrid1.TextMatrix(I, 10) & "," & mDepConDIDList
       End If
   Next
   
   tmp = "Delete from Txn_DepositContractGodownDetail Where DepConID = " & Val(TxtDepositContractID.Text) & " And DepConDID NOT IN (" & mDepConDIDList & ") "
   Call TmpTable
  
Else ''-- Bag Size Wise -----------
   For I = 1 To MSHFlexGrid4.Rows - 1
       If MSHFlexGrid4.TextMatrix(I, 0) = strChecked Then
           If Val(MSHFlexGrid4.TextMatrix(I, 12)) = 0 Then
              RsTxn_DepositContractGodownDetail.AddNew
              RsTxn_DepositContractGodownDetail!DepConDID = GetMaxID("DepConDID", "Txn_DepositContractGodownDetail")
              MSHFlexGrid4.TextMatrix(I, 12) = RsTxn_DepositContractGodownDetail!DepConDID
              RsTxn_DepositContractGodownDetail!G_UID = GetGUID
              RsTxn_DepositContractGodownDetail!DepConID = Val(TxtDepositContractID.Text)
              RsTxn_DepositContractGodownDetail!GodownID = Val(MSHFlexGrid4.TextMatrix(I, 11))
           Else
              RsTxn_DepositContractGodownDetail.Find " DepConDID = " & Val(MSHFlexGrid4.TextMatrix(I, 12))
           End If
             RsTxn_DepositContractGodownDetail!FinalPrice = Val(MSHFlexGrid4.TextMatrix(I, 10))
             RsTxn_DepositContractGodownDetail!DepositFinalPrice = Val(MSHFlexGrid4.TextMatrix(I, 14))
             RsTxn_DepositContractGodownDetail!BaseRate = Val(MSHFlexGrid4.TextMatrix(I, 15))
             RsTxn_DepositContractGodownDetail!DepositBaseRate = Val(MSHFlexGrid4.TextMatrix(I, 16))
             RsTxn_DepositContractGodownDetail!OldBillingRate = Val(MSHFlexGrid4.TextMatrix(I, 9))
             RsTxn_DepositContractGodownDetail!OldDepositBillingRate = Val(MSHFlexGrid4.TextMatrix(I, 13))
             
            ''-- Withdrawal Billing Cycle Rate for Future if required
            RsTxn_DepositContractGodownDetail!WithdrawalFinalPrice = Val(MSHFlexGrid4.TextMatrix(I, 14))
            RsTxn_DepositContractGodownDetail!WithdrawalBaseRate = Val(MSHFlexGrid4.TextMatrix(I, 16))
            RsTxn_DepositContractGodownDetail!OldWithdrawalBillingRate = Val(MSHFlexGrid4.TextMatrix(I, 13))
            
            MSHFlexGrid4.row = I: MSHFlexGrid4.Col = 10
            If MSHFlexGrid4.CellBackColor = vbRed Then
                RsTxn_DepositContractGodownDetail!BaseRateColor = "R"
            ElseIf MSHFlexGrid4.CellBackColor = vbGreen Then
                RsTxn_DepositContractGodownDetail!BaseRateColor = "G"
            ElseIf MSHFlexGrid4.CellBackColor = vbYellow Then
                RsTxn_DepositContractGodownDetail!BaseRateColor = "Y"
            Else
                RsTxn_DepositContractGodownDetail!BaseRateColor = "W"
            End If
            MSHFlexGrid4.row = I: MSHFlexGrid4.Col = 14
            If MSHFlexGrid4.CellBackColor = vbRed Then
                RsTxn_DepositContractGodownDetail!DepositBaseRateColor = "R"
                RsTxn_DepositContractGodownDetail!WithdrawalBaseRateColor = "R"
            ElseIf MSHFlexGrid4.CellBackColor = vbGreen Then
                RsTxn_DepositContractGodownDetail!DepositBaseRateColor = "G"
                RsTxn_DepositContractGodownDetail!WithdrawalBaseRateColor = "G"
            ElseIf MSHFlexGrid4.CellBackColor = vbYellow Then
                RsTxn_DepositContractGodownDetail!DepositBaseRateColor = "Y"
                RsTxn_DepositContractGodownDetail!WithdrawalBaseRateColor = "Y"
            Else
                RsTxn_DepositContractGodownDetail!DepositBaseRateColor = "W"
                RsTxn_DepositContractGodownDetail!WithdrawalBaseRateColor = "W"
            End If
            ''--------------------------------------------------
             
             RsTxn_DepositContractGodownDetail!MinBagSize = Val(MSHFlexGrid4.TextMatrix(I, 7))
             RsTxn_DepositContractGodownDetail!MaxBagSize = Val(MSHFlexGrid4.TextMatrix(I, 8))
             RsTxn_DepositContractGodownDetail!SM_Prefix = mSM_Prefix
             RsTxn_DepositContractGodownDetail!Rent = Val(MSHFlexGrid4.TextMatrix(I, 5))
             RsTxn_DepositContractGodownDetail!RatePerMonthPerMT = Val(MSHFlexGrid4.TextMatrix(I, 18))
             If IsNull(RsTxn_DepositContractGodownDetail!CreatedBy) = True Or RsTxn_DepositContractGodownDetail!CreatedBy = "" Then
                RsTxn_DepositContractGodownDetail!CreatedBy = Gen_UserLoginID
                RsTxn_DepositContractGodownDetail!CreatedDate = Now
             Else
                RsTxn_DepositContractGodownDetail!UpdatedBy = Gen_UserLoginID
                RsTxn_DepositContractGodownDetail!UpdatedDate = Now
             End If
             RsTxn_DepositContractGodownDetail.Update
             RsTxn_DepositContractGodownDetail.Requery
             mDepConDIDList = MSHFlexGrid4.TextMatrix(I, 12) & "," & mDepConDIDList
       End If
   Next
   tmp = "Delete from Txn_DepositContractGodownDetail Where DepConID = " & Val(TxtDepositContractID.Text) & " And DepConDID NOT IN (" & mDepConDIDList & ") "
   Call TmpTable
   
End If
End Sub

Public Sub IndataDepositContractGodownDetail()
Call ClearFrameControl
cmdExportDetails.Enabled = False
CmdAddNewBagSize.Visible = False

TxtFinalPrice.Visible = False
If Check4.Value = 0 Then
   MSHFlexGrid1.Visible = True: MSHFlexGrid4.Visible = False
   Call Grid_Head
   tmp = " Select '',MC.CMPName,MG.GodownNo,MG.Address,MU.UnitType,0 'Rent',MG.Area,"
   tmp = tmp & " TDCG.OldBillingRate,TDCG.FinalPrice,MG.GodownID,TDCG.DepConDID,TDCG.OldDepositBillingRate,TDCG.DepositFinalPrice,"
   tmp = tmp & " TDCG.BaseRate,TDCG.DepositBaseRate,MG.UnitTypeID,TDCG.RatePerMonthPerMT,TDCG.Rent,Isnull(MG.ApprovalLimit,0)ApprovalLimit"
   tmp = tmp & " From Txn_DepositContractGodownDetail TDCG"
   tmp = tmp & " Inner Join Mst_Godown MG on TDCG.GodownID=MG.GodownID"
   tmp = tmp & " Inner Join Mst_CMP MC on MG.CMPID=MC.CMPID"
   tmp = tmp & " Inner Join Mst_UnitType MU on MG.UnitTypeID = MU.UnitTypeID "
   tmp = tmp & " Where TDCG.DepConID=" & RsTxn_DepositContract!DepConID & ""
   Call Tmp2Table
   If TmpRs2.RecordCount > 0 Then
      MSHFlexGrid1.Rows = TmpRs2.RecordCount + 1
      For I = 1 To TmpRs2.RecordCount
            With MSHFlexGrid1
               .row = I
               .Col = 0
               .CellFontName = "Wingdings"
               .CellFontSize = 12
               .CellAlignment = flexAlignCenterCenter
               .Text = strChecked
               .TextMatrix(I, 1) = IIf(IsNull(TmpRs2!CMPName), "", TmpRs2!CMPName)
               .TextMatrix(I, 2) = IIf(IsNull(TmpRs2!GodownNo), "", TmpRs2!GodownNo)
               .TextMatrix(I, 3) = IIf(IsNull(TmpRs2!Address), "", TmpRs2!Address)
               .TextMatrix(I, 4) = IIf(IsNull(TmpRs2!UnitType), "", TmpRs2!UnitType)
               .TextMatrix(I, 5) = Format(IIf(IsNull(TmpRs2!Rent), "", TmpRs2!Rent), "#0.00")
               .TextMatrix(I, 6) = IIf(IsNull(TmpRs2!Area), 0, TmpRs2!Area)
               .TextMatrix(I, 7) = Format(IIf(IsNull(TmpRs2!OldBillingRate), 0, TmpRs2!OldBillingRate), "#0.0000")
               .TextMatrix(I, 8) = Format(IIf(IsNull(TmpRs2!FinalPrice), 0, TmpRs2!FinalPrice), "#0.0000")
               .TextMatrix(I, 9) = IIf(IsNull(TmpRs2!GodownID), 0, TmpRs2!GodownID)
               .TextMatrix(I, 10) = IIf(IsNull(TmpRs2!DepConDID), 0, TmpRs2!DepConDID)
               .TextMatrix(I, 11) = Format(IIf(IsNull(TmpRs2!OldDepositBillingRate), 0, TmpRs2!OldDepositBillingRate), "#0.0000")
               .TextMatrix(I, 12) = Format(IIf(IsNull(TmpRs2!DepositFinalPrice), 0, TmpRs2!DepositFinalPrice), "#0.0000")
               .TextMatrix(I, 13) = Format(IIf(IsNull(TmpRs2!BaseRate), 0, TmpRs2!BaseRate), "#0.0000")
               .TextMatrix(I, 14) = Format(IIf(IsNull(TmpRs2!DepositBaseRate), 0, TmpRs2!DepositBaseRate), "#0.0000")
               .TextMatrix(I, 15) = IIf(IsNull(TmpRs2!UnitTypeID), 0, TmpRs2!UnitTypeID)
               If Val(.TextMatrix(I, 15)) = 1 Or Val(.TextMatrix(I, 15)) = 7 Then
                 .TextMatrix(I, 16) = Format(IIf(IsNull(TmpRs2!RatePerMonthPerMT), "#0.00", TmpRs2!RatePerMonthPerMT), "#0.0000")
               Else
                 .TextMatrix(I, 16) = ""
               End If
               .TextMatrix(I, 17) = Format(IIf(Val(TmpRs2!ApprovalLimit) = 0, mCMPMtoAMRMApprovalLimit, TmpRs2!ApprovalLimit), "#0.00")
               Call CheckRateByColor(CInt(I))
            End With
          TmpRs2.MoveNext
      Next
   End If
Else
   MSHFlexGrid4.Visible = True: MSHFlexGrid1.Visible = False
   Call Grid_Head4
   tmp = " Select '',MC.CMPName,MG.GodownNo,MG.Address,MU.UnitType,0,MG.Area,TDCG.MinBagSize,TDCG.MaxBagSize,"
   tmp = tmp & " TDCG.OldBillingRate,TDCG.FinalPrice,MG.GodownID,TDCG.DepConDID,TDCG.OldDepositBillingRate,TDCG.DepositFinalPrice,"
   tmp = tmp & " TDCG.BaseRate,TDCG.DepositBaseRate,MG.UnitTypeID,TDCG.RatePerMonthPerMT,TDCG.Rent,Isnull(MG.ApprovalLimit,0)ApprovalLimit"
   tmp = tmp & " From Txn_DepositContractGodownDetail TDCG"
   tmp = tmp & " Inner Join Mst_Godown MG on TDCG.GodownID=MG.GodownID"
   tmp = tmp & " Inner Join Mst_CMP MC on MG.CMPID=MC.CMPID"
   tmp = tmp & " Inner Join Mst_UnitType MU on MG.UnitTypeID = MU.UnitTypeID "
   tmp = tmp & " Where TDCG.DepConID=" & RsTxn_DepositContract!DepConID & ""
   Call Tmp2Table
   If TmpRs2.RecordCount > 0 Then
      MSHFlexGrid4.Rows = TmpRs2.RecordCount + 1
      For I = 1 To TmpRs2.RecordCount
        With MSHFlexGrid4
           .row = I
           .Col = 0
           .CellFontName = "Wingdings"
           .CellFontSize = 12
           .CellAlignment = flexAlignCenterCenter
           .Text = strChecked
           .TextMatrix(I, 1) = IIf(IsNull(TmpRs2!CMPName), "", TmpRs2!CMPName)
           .TextMatrix(I, 2) = IIf(IsNull(TmpRs2!GodownNo), "", TmpRs2!GodownNo)
           .TextMatrix(I, 3) = IIf(IsNull(TmpRs2!Address), "", TmpRs2!Address)
           .TextMatrix(I, 4) = IIf(IsNull(TmpRs2!UnitType), "", TmpRs2!UnitType)
           .TextMatrix(I, 5) = Format(IIf(IsNull(TmpRs2!Rent), "", TmpRs2!Rent), "#0.0000")
           .TextMatrix(I, 6) = IIf(IsNull(TmpRs2!Area), 0, TmpRs2!Area)
           .TextMatrix(I, 7) = IIf(IsNull(TmpRs2!MinBagSize), 0, TmpRs2!MinBagSize)
           .TextMatrix(I, 8) = IIf(IsNull(TmpRs2!MaxBagSize), 0, TmpRs2!MaxBagSize)
           .TextMatrix(I, 9) = Format(IIf(IsNull(TmpRs2!OldBillingRate), 0, TmpRs2!OldBillingRate), "#0.0000")
           .TextMatrix(I, 10) = Format(IIf(IsNull(TmpRs2!FinalPrice), 0, TmpRs2!FinalPrice), "#0.0000")
           .TextMatrix(I, 11) = IIf(IsNull(TmpRs2!GodownID), 0, TmpRs2!GodownID)
           .TextMatrix(I, 12) = IIf(IsNull(TmpRs2!DepConDID), 0, TmpRs2!DepConDID)
           .TextMatrix(I, 13) = Format(IIf(IsNull(TmpRs2!OldDepositBillingRate), 0, TmpRs2!OldDepositBillingRate), "#0.0000")
           .TextMatrix(I, 14) = Format(IIf(IsNull(TmpRs2!DepositFinalPrice), 0, TmpRs2!DepositFinalPrice), "#0.0000")
           .TextMatrix(I, 15) = Format(IIf(IsNull(TmpRs2!BaseRate), 0, TmpRs2!BaseRate), "#0.0000")
           .TextMatrix(I, 16) = Format(IIf(IsNull(TmpRs2!DepositBaseRate), 0, TmpRs2!DepositBaseRate), "#0.0000")
           .TextMatrix(I, 17) = IIf(IsNull(TmpRs2!UnitTypeID), 0, TmpRs2!UnitTypeID)
           If Val(.TextMatrix(I, 17)) = 1 Or Val(.TextMatrix(I, 17)) = 7 Then
            .TextMatrix(I, 18) = Format(IIf(IsNull(TmpRs2!RatePerMonthPerMT), "#0.0000", TmpRs2!RatePerMonthPerMT), "#0.0000")
           Else
            .TextMatrix(I, 18) = "0.0000"
           End If
           .TextMatrix(I, 19) = Format(IIf(Val(TmpRs2!ApprovalLimit) = 0, mCMPMtoAMRMApprovalLimit, TmpRs2!ApprovalLimit), "#0.00")
           Call CheckBagRateByColor(CInt(I))
        End With
        TmpRs2.MoveNext
     Next
   End If
End If
End Sub

Public Sub IndataDepositContractGodownDetail_Edit()
Call ClearFrameControl
If Check4.Value = 0 Then
   Call Grid_Head
   tmp = " Select '',MC.CMPName,MG.GodownNo,MG.Address,MG.UnitTypeID,0,MG.GodownCapacity,TDCG.OldBillingRate,TDCG.FinalPrice,MG.GodownID,Cast(TDCG.DepConDID As Varchar) DepConDID,TDCG.OldDepositBillingRate,TDCG.DepositFinalPrice,TDCG.BaseRate,TDCG.DepositBaseRate"
   tmp = tmp & " From Txn_DepositContractGodownDetail TDCG"
   tmp = tmp & " Inner Join Mst_Godown MG on TDCG.GodownID=MG.GodownID"
   tmp = tmp & " Inner Join Mst_CMP MC on MG.CMPID=MC.CMPID Where TDCG.DepConID=" & Val(TxtDepositContractID) & ""
   tmp = tmp & " Union All"
   tmp = tmp & " Select '',MC.CMPName,MG.GodownNo,MG.Address,MG.UnitTypeID,0,MG.GodownCapacity,M1.Rate OldBillingRate,M1.Rate AS FinalPrice,MG.GodownID,cast('' As Varchar) DepConDID,M1.Rate OldDepositRate,M1.Rate AS FinalDepositPrice,M1.Rate BillingRate13,M1.Rate DepositBillingRate14"
   tmp = tmp & " From Mst_Godown MG "
   tmp = tmp & " Inner Join Mst_CMP MC On MG.CMPID=MC.CMPID"
   tmp = tmp & " Inner Join Mst_StorageRateGodownDetail M1 On MG.GodownID=M1.GodownID"
   tmp = tmp & " Where MG.CloseDate is null And M1.SRCID=" & Val(mSRCID) & ""
   tmp = tmp & " And MG.GodownID Not in (Select TDCGD.GodownId From Txn_DepositContractGodownDetail TDCGD Inner Join Txn_DepositContract TDC on TDCGD.DepConID=TDC.DepConID Where TDC.SRCID=" & Val(mSRCID) & " And TDC.DepConID=" & Val(TxtDepositContractID) & " "
   tmp = tmp & " And StartDate='" & Format(TxtStartDate, "YYYY-MM-DD") & "' And EndDate='" & Format(TxtEndDate, "YYYY-MM-DD") & "' And ClientID=" & mClientIDRow & ")"
   Call Tmp2Table
   If TmpRs2.RecordCount > 0 Then
      MSHFlexGrid1.Rows = TmpRs2.RecordCount + 2
      For I = 1 To TmpRs2.RecordCount
          For C = 0 To MSHFlexGrid1.Cols - 1
              MSHFlexGrid1.TextMatrix(I, C) = IIf(IsNull(TmpRs2.Fields(C)), "", TmpRs2.Fields(C))
          Next
          
          MSHFlexGrid1.row = I
          MSHFlexGrid1.Col = 0
          MSHFlexGrid1.CellFontName = "Wingdings"
          MSHFlexGrid1.CellFontSize = 12
          MSHFlexGrid1.CellAlignment = flexAlignCenterCenter
          
          If MSHFlexGrid1.TextMatrix(I, 10) <> "" Then
             MSHFlexGrid1.Text = strChecked
          Else
             MSHFlexGrid1.Text = strUnChecked
             MSHFlexGrid1.TextMatrix(I, 8) = CalculateRate(MSHFlexGrid1.TextMatrix(I, 13), mBillingRatio, CmbBillingCycle)
             MSHFlexGrid1.TextMatrix(I, 7) = CalculateRate(MSHFlexGrid1.TextMatrix(I, 13), mBillingRatio, CmbBillingCycle)
             MSHFlexGrid1.TextMatrix(I, 12) = CalculateRate(MSHFlexGrid1.TextMatrix(I, 14), mDepositBillingRatio, CmbDepositBillingCycle)
             MSHFlexGrid1.TextMatrix(I, 11) = CalculateRate(MSHFlexGrid1.TextMatrix(I, 14), mDepositBillingRatio, CmbDepositBillingCycle)
          End If
           
          MSHFlexGrid1.TextMatrix(I, 4) = GetUnitName(IIf(IsNull(MSHFlexGrid1.TextMatrix(I, 4)), 0, MSHFlexGrid1.TextMatrix(I, 4)))
          TmpRs2.MoveNext
      Next
   End If
Else
   Call Grid_Head4
   tmp = " Select '',MC.CMPName,MG.GodownNo,MG.Address,MG.UnitTypeID,0,MG.GodownCapacity,TDCG.BagSize,TDCG.OldBillingRate,TDCG.FinalPrice,MG.GodownID,Cast(TDCG.DepConDID As Varchar) DepConDID,TDCG.OldDepositBillingRate,TDCG.DepositFinalPrice,TDCG.BaseRate,TDCG.DepositBaseRate"
   tmp = tmp & " From Txn_DepositContractGodownDetail TDCG"
   tmp = tmp & " Inner Join Mst_Godown MG on TDCG.GodownID=MG.GodownID"
   tmp = tmp & " Inner Join Mst_CMP MC on MG.CMPID=MC.CMPID"
   tmp = tmp & " Where TDCG.DepConID=" & RsTxn_DepositContract!DepConID & ""
   tmp = tmp & " Union All"
   tmp = tmp & " Select '',MC.CMPName,MG.GodownNo,MG.Address,MG.UnitTypeID,0,MG.GodownCapacity,M2.BagSize,M1.Rate OldBillingRate,M1.Rate FinalPrice,MG.GodownID,Cast('' As Varchar) DepConDID,M1.Rate OldDepositRate,M1.Rate AS DepositFinalPrice,M1.Rate BillingRate13,M1.Rate DepositBillingRate14"
   tmp = tmp & " From Mst_Godown MG Inner Join Mst_CMP MC On MG.CMPID=MC.CMPID"
   tmp = tmp & " Inner Join Mst_StorageRateGodownDetail M1 On MG.GodownID=M1.GodownID"
   tmp = tmp & " Inner Join Mst_StorageRateCardBagSize M2 On M1.SRCID=M2.SRCID"
   tmp = tmp & " Where MG.CloseDate Is Null And M1.SRCID = " & mSRCID & ""
   tmp = tmp & " And Cast(MG.GodownID as Varchar)+'~'+cast(M2.BagSize As Varchar) Not in (Select Cast(TDCGD.GodownId as Varchar)+'~'+cast(TDCGD.BagSize As Varchar) "
   tmp = tmp & " From Txn_DepositContractGodownDetail TDCGD Inner Join Txn_DepositContract TDC on TDCGD.DepConID=TDC.DepConID Where TDC.SRCID=" & mSRCID & " "
   tmp = tmp & " AND (Cast(TDCGD.GodownId as Varchar)+'~'+cast(TDCGD.BagSize As Varchar)) IS NOT NULL And StartDate='" & Format(TxtStartDate, "YYYY-MM-DD") & "' And EndDate='" & Format(TxtEndDate, "YYYY-MM-DD") & "')"
   Call Tmp2Table
   If TmpRs2.RecordCount > 0 Then
      
      For I = 1 To TmpRs2.RecordCount
          iCtr = Val(MSHFlexGrid4.Rows - 1)
          For C = 0 To MSHFlexGrid4.Cols - 1
              MSHFlexGrid4.TextMatrix(iCtr, C) = IIf(IsNull(TmpRs2.Fields(C)), "", TmpRs2.Fields(C))
          Next
          MSHFlexGrid4.row = Val(iCtr)
          MSHFlexGrid4.Col = 0
          MSHFlexGrid4.CellFontName = "Wingdings"
          MSHFlexGrid4.CellFontSize = 12
          MSHFlexGrid4.CellAlignment = flexAlignCenterCenter
          MSHFlexGrid4.Text = strUnChecked

          If MSHFlexGrid4.TextMatrix(I, 11) <> "" Then
             MSHFlexGrid4.Text = strChecked
          Else
             MSHFlexGrid4.TextMatrix(iCtr, 9) = CalculateRate(MSHFlexGrid4.TextMatrix(I, 14), mBillingRatio, CmbBillingCycle) * Val(MSHFlexGrid4.TextMatrix(iCtr, 7)) / 1000
             MSHFlexGrid4.TextMatrix(iCtr, 8) = CalculateRate(MSHFlexGrid4.TextMatrix(I, 14), mBillingRatio, CmbBillingCycle) * Val(MSHFlexGrid4.TextMatrix(iCtr, 7)) / 1000
             MSHFlexGrid4.TextMatrix(iCtr, 13) = CalculateRate(MSHFlexGrid4.TextMatrix(I, 15), mDepositBillingRatio, CmbDepositBillingCycle) * Val(MSHFlexGrid4.TextMatrix(iCtr, 7)) / 1000
             MSHFlexGrid4.TextMatrix(iCtr, 12) = CalculateRate(MSHFlexGrid4.TextMatrix(I, 15), mDepositBillingRatio, CmbDepositBillingCycle) * Val(MSHFlexGrid4.TextMatrix(iCtr, 7)) / 1000
          End If
          
          MSHFlexGrid4.TextMatrix(iCtr, 4) = GetUnitName(IIf(IsNull(MSHFlexGrid4.TextMatrix(iCtr, 4)), 0, MSHFlexGrid4.TextMatrix(iCtr, 4)))
          MSHFlexGrid4.Rows = MSHFlexGrid4.Rows + 1
          TmpRs2.MoveNext
          
      Next
   End If
End If
End Sub

Private Sub SearchCmd_Click()
Dim x As Integer
If SearchCmd.Caption = "Cancel" Then
   If RsTxn_DepositContract.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   Call Indata
   Exit Sub
End If

Frame0.Left = Frame1.Left
Frame0.Top = Frame1.Top
Frame0.Height = Me.Height - 800
Frame0.Width = Me.Width - 500
Frame0.Visible = True

MSHFlexGrid2.Width = Frame0.Width - 200
MSHFlexGrid2.Height = Frame0.Height - 2000

End Sub

Public Sub GetStorageRateCard()
tmp = " Select MSRC.SRCID,MSRC.RateOnQty,MSRC.BillingCycleID,MSRC.DepoBillingCycleID,MBC.BillingCycle,MBCC.BillingCycle As DepoBillingCycle,MBC.Ratio,MBCC.Ratio As DepositRatio"
tmp = tmp & " From Mst_StorageRateCard MSRC"
tmp = tmp & " Inner Join Mst_BillingCycle MBC on MSRC.BillingCycleID=MBC.BillingCycleID"
tmp = tmp & " Inner Join Mst_BillingCycle MBCC on MSRC.DepoBillingCycleID=MBCC.BillingCycleID"
tmp = tmp & " Where LocationID=" & mLocationID & " And CommodityID=" & mCommodityID & " AND MSRC.ExchangeTypeID = 1 And DefaultRate=1 And StatusFlag = 'N'"
Call Tmp1Table
If TmpRs1.RecordCount > 0 Then
   mSRCID = IIf(IsNull(TmpRs1!SRCID), 0, TmpRs1!SRCID)
   mRateOnQty = IIf(IsNull(TmpRs1!RateOnQty), 0, TmpRs1!RateOnQty)
   TxtBaseRate.Text = Val(mRateOnQty)
   
   Call CmbBillingCycle_GotFocus
   Call CmbDepositBillingCycle_GotFocus
   
   CmbBillingCycle.Text = IIf(IsNull(TmpRs1!BillingCycle), 0, TmpRs1!BillingCycle)
   mBillingCycleID = IIf(IsNull(TmpRs1!BillingCycleID), 0, TmpRs1!BillingCycleID)
   CmbDepositBillingCycle.Text = IIf(IsNull(TmpRs1!DepoBillingCycle), 0, TmpRs1!DepoBillingCycle)
   mDepositBillingCycleID = IIf(IsNull(TmpRs1!DepoBillingCycleID), 0, TmpRs1!DepoBillingCycleID)
   mBillingRatio = IIf(IsNull(TmpRs1!Ratio), 0, TmpRs1!Ratio)
   mDepositBillingRatio = IIf(IsNull(TmpRs1!DepositRatio), 0, TmpRs1!DepositRatio)
Else
   mSRCID = Null
   TxtBaseRate.Text = ""
End If
End Sub

Public Sub GetGodownDetails()
Dim mR, mOwnGodownSqrFeetRate As Variant

mTmpDepConID = ""
mValidGodownID = ""

If CmbLocation.Text = "" Or CmbCommodity.Text = "" Or CmbClient.Text = "" Then
    Call Grid_Head
    Call Grid_Head4
    Exit Sub
End If

If CmbClient.Text <> "" Then
    ''And StateID=" & mStateID & "
    If CmbLocation.Text <> "" Then mOwnGodownSqrFeetRate = GetOwnGodownPerSqtFeet(mLocationID)
    '(('" & TxtStartDate & "' >= StartDate And '" & TxtStartDate & "' <= EndDate) Or '" & TxtEndDate & "' >= StartDate And '" & TxtEndDate & "' <= EndDate Or ('" & TxtStartDate & "' <= StartDate And '" & TxtEndDate & "' >= EndDate))
    tmp = " Select StartDate,EndDate,DepConID From Txn_DepositContract where ApprovalFlag not In ('C','R') And DepConID<>" & Val(TxtDepositContractID) & " And LocationID=" & mLocationID & " And  CommodityID=" & mCommodityID & " And ClientID=" & mClientIDRow & ""
    ''-- 14 Feb 2018
    tmp = tmp & " And ContractType = '" & mContractTypeFlag & "' "
    ''--------------
    tmp = tmp & "  And (('" & TxtStartDate & "' >= StartDate And '" & TxtStartDate & "' <= EndDate) Or '" & TxtEndDate & "' >= StartDate And '" & TxtEndDate & "' <= EndDate Or ('" & TxtStartDate & "' <= StartDate And '" & TxtEndDate & "' >= EndDate))"
    Call Tmp2Table
    mTmpDepConID = ""
    If TmpRs2.RecordCount > 0 Then
       For mR = 1 To TmpRs2.RecordCount
           mTmpDepConID = mTmpDepConID & "," & TmpRs2!DepConID
           TmpRs2.MoveNext
       Next mR
       If mTmpDepConID <> "" Then
        mTmpDepConID = Right(mTmpDepConID, Len(mTmpDepConID) - 1)
       End If
    End If
    If mTmpDepConID <> "" Then
        tmp = " SELECT GodownId FROM Txn_DepositContractGodownDetail Where DepConID in (" & mTmpDepConID & ") "
        Call Tmp2Table
        mValidGodownID = ""
        If TmpRs2.RecordCount > 0 Then
           For mR = 1 To TmpRs2.RecordCount
             mValidGodownID = mValidGodownID & "," & TmpRs2!GodownID
             TmpRs2.MoveNext
           Next mR
           If mValidGodownID <> "" Then
             mValidGodownID = Right(mValidGodownID, Len(mValidGodownID) - 1)
           End If
        End If
    End If
End If

Check2.Value = Unchecked
If Check4.Value = 0 Then
   Call Grid_Head
   
'''--05 Feb 2018 As per Sanjeev sir discuss and sugesstion on Tuesday, January 30, 2018
'''-- Last Rate need to capture Godown Wise

'''   tmp = " Select MC.CMPName,MG.GodownNo,MG.Address,MU.UnitType,sum(M1.RentAmount) 'Rent',MG.Area,"
'''   tmp = tmp & " MG.GodownID,cast('' As Varchar) DepConDID,MU.UnitTypeID,Isnull(MG.ApprovalLimit,0)ApprovalLimit "
'''   tmp = tmp & " From Mst_Godown MG"
'''   tmp = tmp & " Inner Join Mst_CMP MC On MG.CMPID=MC.CMPID"
'''   tmp = tmp & " Inner Join Mst_UnitType MU on MG.UnitTypeID = MU.UnitTypeID"
'''   tmp = tmp & " Inner Join Txn_GodownLessorAgreeDetail M1 On MG.GodownID=M1.GodownID"
'''   tmp = tmp & " inner Join Txn_LessorAgreement TLA on M1.LAID = TLA.LAID And TLA.Flag = 'A'"
'''   tmp = tmp & " Where MG.CloseDate is null  And MG.StartDate<='" & Format(TxtEndDate, "YYYY-MM-DD") & "' And ( '" & TxtStartDate.Value & "' <= TLA.ExpiryDate And '" & TxtEndDate.Value & "' >= TLA.StartDate )"
'''   tmp = tmp & " And MC.LocationID = " & mLocationID & " And  MU.UnitTypeID = 1"
'''   If mTmpDepConID <> "" And mValidGodownID <> "" Then
'''      tmp = tmp & " AND MG.GodownID NOT IN (" & mValidGodownID & ") "
'''   End If
'''   tmp = tmp & " Group by MC.CMPName,MG.GodownNo,MG.Address,MU.UnitType,MG.Area,MG.GodownID,MU.UnitTypeID,Isnull(MG.ApprovalLimit,0) "
'''   tmp = tmp & " Union All"
'''   tmp = tmp & " Select MC.CMPName,MG.GodownNo,MG.Address,MU.UnitType,0 'Rent',MG.Area,MG.GodownID,cast('' As Varchar) DepConDID,MU.UnitTypeID,Isnull(MG.ApprovalLimit,0)ApprovalLimit "
'''   tmp = tmp & " From Mst_Godown MG"
'''   tmp = tmp & " Inner Join Mst_CMP MC On MG.CMPID=MC.CMPID"
'''   tmp = tmp & " Inner Join Mst_UnitType MU on MG.UnitTypeID = MU.UnitTypeID"
'''   tmp = tmp & " Where MG.CloseDate is null  And MG.StartDate<='" & Format(TxtEndDate, "YYYY-MM-DD") & "' And '" & Format(TxtEndDate, "YYYY-MM-DD") & "' <= MG.AgreementEndDate "
'''   tmp = tmp & " And MC.LocationID = " & mLocationID & " And  MU.UnitTypeID <> 1"
'''   If mTmpDepConID <> "" And mValidGodownID <> "" Then
'''      tmp = tmp & " AND MG.GodownID NOT IN (" & mValidGodownID & ") "
'''   End If
'''   tmp = tmp & " Group by MC.CMPName,MG.GodownNo,MG.Address,MU.UnitType,MG.Area,MG.GodownID,MU.UnitTypeID,Isnull(MG.ApprovalLimit,0) "

   tmp = " Select MC.CMPName,MG.GodownNo,MG.Address,MU.UnitType,sum(M1.RentAmount) 'Rent',MG.Area,"
   tmp = tmp & " MG.GodownID,cast('' As Varchar) DepConDID,MU.UnitTypeID,Isnull(MG.ApprovalLimit,0)ApprovalLimit "
   tmp = tmp & " From Mst_Godown MG"
   tmp = tmp & " Inner Join Mst_CMP MC On MG.CMPID=MC.CMPID"
   tmp = tmp & " Inner Join Mst_UnitType MU on MG.UnitTypeID = MU.UnitTypeID"
   tmp = tmp & " Inner Join Txn_GodownLessorAgreeDetail M1 On MG.GodownID=M1.GodownID"
   tmp = tmp & " inner Join Txn_LessorAgreement TLA on M1.LAID = TLA.LAID And TLA.Flag = 'A'"
   tmp = tmp & " Where MG.CloseDate is null  And MG.StartDate<='" & Format(TxtEndDate, "YYYY-MM-DD") & "' And ( '" & TxtStartDate.Value & "' <= TLA.ExpiryDate And '" & TxtEndDate.Value & "' >= TLA.StartDate )"
   tmp = tmp & " And MC.LocationID = " & mLocationID & " And  MU.UnitTypeID = 1"
   If mTmpDepConID <> "" And mValidGodownID <> "" Then
      tmp = tmp & " AND MG.GodownID NOT IN (" & mValidGodownID & ") "
   End If
   
        tmp = tmp & " And Convert(Varchar,MG.GodownID)+''+Convert(Varchar,TLA.ExpiryDate,112) In "
        tmp = tmp & " ( "
        tmp = tmp & "   Select Convert(Varchar,AA.GodownID)+''+Convert(Varchar,Max(BB.ExpiryDate),112) "
        tmp = tmp & "   From Txn_GodownLessorAgreeDetail AA Inner Join Txn_LessorAgreement BB On AA.LAID = BB.LAID "
        tmp = tmp & "   Where Convert(Varchar,AA.LAID)+''+Convert(Varchar,AA.GodownID) "
        tmp = tmp & "   In ( Select Convert(Varchar,E.LAID)+''+Convert(Varchar,D.GodownID) LAIDGodownKey "
        tmp = tmp & "   From Mst_Godown A Inner Join Mst_CMP B On A.CMPID=B.CMPID "
        tmp = tmp & "   Inner Join Mst_UnitType C on A.UnitTypeID = C.UnitTypeID "
        tmp = tmp & "   Inner Join Txn_GodownLessorAgreeDetail D On A.GodownID=D.GodownID "
        tmp = tmp & "   Inner Join Txn_LessorAgreement E on D.LAID = E.LAID And E.Flag = 'A' "
        tmp = tmp & "   Where A.CloseDate is null And A.StartDate<='" & Format(TxtEndDate, "YYYY-MM-DD") & "' And ( '" & TxtStartDate.Value & "' <= E.ExpiryDate And '" & TxtEndDate.Value & "' >= E.StartDate ) "
        tmp = tmp & "   And B.LocationID = " & mLocationID & " And C.UnitTypeID = 1 "
            If mTmpDepConID <> "" And mValidGodownID <> "" Then
               tmp = tmp & " AND A.GodownID NOT IN (" & mValidGodownID & ") "
            End If
        tmp = tmp & "     ) Group By AA.GodownID "
        tmp = tmp & " ) "
   
   tmp = tmp & " Group by MC.CMPName,MG.GodownNo,MG.Address,MU.UnitType,MG.Area,MG.GodownID,MU.UnitTypeID,Isnull(MG.ApprovalLimit,0) "
   tmp = tmp & " Union All"
   tmp = tmp & " Select MC.CMPName,MG.GodownNo,MG.Address,MU.UnitType,0 'Rent',MG.Area,MG.GodownID,cast('' As Varchar) DepConDID,MU.UnitTypeID,Isnull(MG.ApprovalLimit,0)ApprovalLimit "
   tmp = tmp & " From Mst_Godown MG"
   tmp = tmp & " Inner Join Mst_CMP MC On MG.CMPID=MC.CMPID"
   tmp = tmp & " Inner Join Mst_UnitType MU on MG.UnitTypeID = MU.UnitTypeID"
   tmp = tmp & " Where MG.CloseDate is null  And MG.StartDate<='" & Format(TxtEndDate, "YYYY-MM-DD") & "' And '" & Format(TxtEndDate, "YYYY-MM-DD") & "' <= MG.AgreementEndDate "
   tmp = tmp & " And MC.LocationID = " & mLocationID & " And  MU.UnitTypeID <> 1"
   If mTmpDepConID <> "" And mValidGodownID <> "" Then
      tmp = tmp & " AND MG.GodownID NOT IN (" & mValidGodownID & ") "
   End If
   tmp = tmp & " Group by MC.CMPName,MG.GodownNo,MG.Address,MU.UnitType,MG.Area,MG.GodownID,MU.UnitTypeID,Isnull(MG.ApprovalLimit,0) "
   
'''-----------------------

   Call Tmp2Table
   With MSHFlexGrid1
    If TmpRs2.RecordCount > 0 Then
       .Rows = TmpRs2.RecordCount + 1
       For I = 1 To TmpRs2.RecordCount
           .row = I
           .Col = 0
           .CellFontName = "Wingdings"
           .CellFontSize = 12
           .CellAlignment = flexAlignCenterCenter
           .Text = strUnChecked
           .TextMatrix(I, 1) = IIf(IsNull(TmpRs2!CMPName), 0, TmpRs2!CMPName)
           .TextMatrix(I, 2) = IIf(IsNull(TmpRs2!GodownNo), 0, TmpRs2!GodownNo)
           .TextMatrix(I, 3) = IIf(IsNull(TmpRs2!Address), 0, TmpRs2!Address)
           .TextMatrix(I, 4) = IIf(IsNull(TmpRs2!UnitType), 0, TmpRs2!UnitType)
           .TextMatrix(I, 5) = Format(IIf(IsNull(TmpRs2!Rent), 0, TmpRs2!Rent), "#0.00")
           .TextMatrix(I, 6) = Format(IIf(IsNull(TmpRs2!Area), 0, TmpRs2!Area), "#0.000")
           
           If TmpRs2!UnitTypeID = 7 And Val(.TextMatrix(I, 5)) = 0 Then
                .TextMatrix(I, 5) = Format(Val(.TextMatrix(I, 6)) * mOwnGodownSqrFeetRate, "#0.00")
           End If
           mCalculatedBaseRent = (Val(.TextMatrix(I, 5)) / Val(.TextMatrix(I, 6)))
           If TmpRs2!UnitTypeID = 1 Or TmpRs2!UnitTypeID = 7 Then
'Commented on 03 Dec 2018 After Mail from MD Sir and Ganesh Sir on 03 Dec 2018 for change of formuls
'              If LCase(mFumigationFlag) = "y" Then
'                 .TextMatrix(I, 7) = Format((((Val(.TextMatrix(I, 5)) / Val(.TextMatrix(I, 6))) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorY, "#0.0000")
'                 .TextMatrix(I, 11) = Format((((Val(.TextMatrix(I, 5)) / Val(.TextMatrix(I, 6))) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorY, "#0.0000")
'                 .TextMatrix(I, 16) = Format((((Val(.TextMatrix(I, 5)) / Val(.TextMatrix(I, 6))) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorY, "#0.0000")
'              Else
'                 .TextMatrix(I, 7) = Format((((Val(.TextMatrix(I, 5)) / Val(.TextMatrix(I, 6))) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorN, "#0.0000")
'                 .TextMatrix(I, 11) = Format((((Val(.TextMatrix(I, 5)) / Val(.TextMatrix(I, 6))) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorN, "#0.0000")
'                 .TextMatrix(I, 16) = Format((((Val(.TextMatrix(I, 5)) / Val(.TextMatrix(I, 6))) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorN, "#0.0000")
'              End If
              
              If TmpRs2!UnitTypeID = 1 Then
                 If mCalculatedBaseRent < mCurrentBaseRentFactor Then
                    mCalculatedBaseRent = mCurrentBaseRentFactor
                 End If
              ElseIf TmpRs2!UnitTypeID = 2 Then
                a = a
              End If
              If LCase(mFumigationFlag) = "y" Then
'                 .TextMatrix(I, 7) = Format((((Val(.TextMatrix(I, 5)) / Val(.TextMatrix(I, 6))) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorY, "#0.0000")
'                 .TextMatrix(I, 11) = Format((((Val(.TextMatrix(I, 5)) / Val(.TextMatrix(I, 6))) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorY, "#0.0000")
'                 .TextMatrix(I, 16) = Format((((Val(.TextMatrix(I, 5)) / Val(.TextMatrix(I, 6))) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorY, "#0.0000")
              
                 .TextMatrix(I, 7) = Format((((Val(mCalculatedBaseRent)) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorY, "#0.0000")
                 .TextMatrix(I, 11) = Format((((Val(mCalculatedBaseRent)) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorY, "#0.0000")
                 .TextMatrix(I, 16) = Format((((Val(mCalculatedBaseRent)) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorY, "#0.0000")
              
              
              Else
'                 .TextMatrix(I, 7) = Format((((Val(.TextMatrix(I, 5)) / Val(.TextMatrix(I, 6))) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorN, "#0.0000")
'                 .TextMatrix(I, 11) = Format((((Val(.TextMatrix(I, 5)) / Val(.TextMatrix(I, 6))) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorN, "#0.0000")
'                 .TextMatrix(I, 16) = Format((((Val(.TextMatrix(I, 5)) / Val(.TextMatrix(I, 6))) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorN, "#0.0000")
              
              
                 .TextMatrix(I, 7) = Format((((Val(mCalculatedBaseRent)) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorN, "#0.0000")
                 .TextMatrix(I, 11) = Format((((Val(mCalculatedBaseRent)) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorN, "#0.0000")
                 .TextMatrix(I, 16) = Format((((Val(mCalculatedBaseRent)) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorN, "#0.0000")
              
              End If
           
           
           Else
'              .TextMatrix(I, 7) = Val(mCurrentBaseRentFactor) '"0.0000"
'              .TextMatrix(I, 11) = Val(mCurrentBaseRentFactor) '"0.0000"
'              .TextMatrix(I, 16) = Val(mCurrentBaseRentFactor) '"0.0000"
              mCalculatedBaseRent = mCurrentBaseRentFactor
              If LCase(mFumigationFlag) = "y" Then
                 .TextMatrix(I, 7) = Format((((Val(mCalculatedBaseRent)) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorY, "#0.0000")
                 .TextMatrix(I, 11) = Format((((Val(mCalculatedBaseRent)) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorY, "#0.0000")
                 .TextMatrix(I, 16) = Format((((Val(mCalculatedBaseRent)) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorY, "#0.0000")
              Else
                 .TextMatrix(I, 7) = Format((((Val(mCalculatedBaseRent)) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorN, "#0.0000")
                 .TextMatrix(I, 11) = Format((((Val(mCalculatedBaseRent)) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorN, "#0.0000")
                 .TextMatrix(I, 16) = Format((((Val(mCalculatedBaseRent)) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorN, "#0.0000")
              End If
           
           End If
           
           .TextMatrix(I, 9) = IIf(IsNull(TmpRs2!GodownID), 0, TmpRs2!GodownID)
           
           .TextMatrix(I, 13) = Val(.TextMatrix(I, 7))
           .TextMatrix(I, 14) = Val(.TextMatrix(I, 7))
           .TextMatrix(I, 15) = IIf(IsNull(TmpRs2!UnitTypeID), 0, TmpRs2!UnitTypeID)
           .TextMatrix(I, 17) = Format(IIf(Val(TmpRs2!ApprovalLimit) = 0, mCMPMtoAMRMApprovalLimit, TmpRs2!ApprovalLimit), "#0.00")
           
            mBillingOnePercentage = Format(IIf(Val(.TextMatrix(I, 7)) > 0, Val(.TextMatrix(I, 7)) / 100, 0), "#0.0000")
            mDepositOnePercentage = Format(IIf(Val(.TextMatrix(I, 11)) > 0, Val(.TextMatrix(I, 11)) / 100, 0), "#0.0000")
           
           .TextMatrix(I, 8) = Format(Val(.TextMatrix(I, 7)) + (mBillingOnePercentage * .TextMatrix(I, 17)), "#0.0000")
           .TextMatrix(I, 12) = Format(Val(.TextMatrix(I, 11)) + (mDepositOnePercentage * .TextMatrix(I, 17)), "#0.0000")
           
           TmpRs2.MoveNext
       Next
    Else
       MsgBox "No Godown Found !!!"
       Exit Sub
    End If
    '''-- Billing Cycle / Deposite Biling Cycle Wise Rate Calculation
    If CmbBillingCycle.Text <> "" Then
       For I = 1 To .Rows - 1
           If .TextMatrix(I, 1) <> "" Then
              .TextMatrix(I, 7) = Format(CalculateRate(.TextMatrix(I, 13), mBillingRatio, CmbBillingCycle), "#0.0000")
              mBillingOnePercentage = Format(IIf(Val(.TextMatrix(I, 7)) > 0, Val(.TextMatrix(I, 7)) / 100, 0), "#0.0000")
              .TextMatrix(I, 8) = Format(Val(.TextMatrix(I, 7)) + (mBillingOnePercentage * .TextMatrix(I, 17)), "#0.0000")
           End If
       Next
    End If
    If CmbDepositBillingCycle.Text <> "" Then
       For I = 1 To .Rows - 1
           .TextMatrix(I, 11) = Format(CalculateRate(.TextMatrix(I, 14), mDepositBillingRatio, CmbDepositBillingCycle), "#0.0000")
           mDepositOnePercentage = Format(IIf(Val(.TextMatrix(I, 11)) > 0, Val(.TextMatrix(I, 11)) / 100, 0), "#0.0000")
           .TextMatrix(I, 12) = Format(Val(.TextMatrix(I, 11)) + (mDepositOnePercentage * .TextMatrix(I, 17)), "#0.0000")
       Next
    End If
   End With
   
Else
   Call Grid_Head4
   
'''--05 Feb 2018 As per Sanjeev sir discuss and sugesstion on Tuesday, January 30, 2018
'''-- Last Rate need to capture Godown Wise

'''   tmp = " Select MC.CMPName,MG.GodownNo,MG.Address,MU.UnitType,sum(M1.RentAmount) 'Rent',MG.Area,MG.GodownID,cast('' As Varchar) DepConDID,MU.UnitTypeID,Isnull(MG.ApprovalLimit,0)ApprovalLimit"
'''   tmp = tmp & " From Mst_Godown MG"
'''   tmp = tmp & " Inner Join Mst_CMP MC On MG.CMPID=MC.CMPID"
'''   tmp = tmp & " Inner Join Mst_UnitType MU on MG.UnitTypeID = MU.UnitTypeID"
'''   tmp = tmp & " Inner Join Txn_GodownLessorAgreeDetail M1 On MG.GodownID=M1.GodownID"
'''   tmp = tmp & " inner Join Txn_LessorAgreement TLA on M1.LAID = TLA.LAID And TLA.Flag = 'A'"
'''   tmp = tmp & " Where MG.CloseDate is null  And MG.StartDate<='" & Format(TxtEndDate, "YYYY-MM-DD") & "' And ( '" & TxtStartDate.Value & "' <= TLA.ExpiryDate And '" & TxtEndDate.Value & "' >= TLA.StartDate )"
'''   tmp = tmp & " And MC.LocationID = " & mLocationID & " And  MU.UnitTypeID = 1"
'''   If mTmpDepConID <> "" And mValidGodownID <> "" Then
'''      tmp = tmp & " AND MG.GodownID NOT IN (" & mValidGodownID & ") "
'''   End If
'''   tmp = tmp & " Group by MC.CMPName,MG.GodownNo,MG.Address,MU.UnitType,MG.Area,MG.GodownID,MU.UnitTypeID,Isnull(MG.ApprovalLimit,0) "
'''   tmp = tmp & " Union All"
'''   tmp = tmp & " Select MC.CMPName,MG.GodownNo,MG.Address,MU.UnitType,0 'Rent',MG.Area,MG.GodownID,cast('' As Varchar) DepConDID,MU.UnitTypeID,Isnull(MG.ApprovalLimit,0)ApprovalLimit"
'''   tmp = tmp & " From Mst_Godown MG"
'''   tmp = tmp & " Inner Join Mst_CMP MC On MG.CMPID=MC.CMPID"
'''   tmp = tmp & " Inner Join Mst_UnitType MU on MG.UnitTypeID = MU.UnitTypeID"
'''   tmp = tmp & " Where MG.CloseDate is null  And MG.StartDate<='" & Format(TxtEndDate, "YYYY-MM-DD") & "' And '" & Format(TxtEndDate, "YYYY-MM-DD") & "' <= MG.AgreementEndDate "
'''   tmp = tmp & " And MC.LocationID = " & mLocationID & " And  MU.UnitTypeID <> 1"
'''   If mTmpDepConID <> "" And mValidGodownID <> "" Then
'''      tmp = tmp & " AND MG.GodownID NOT IN (" & mValidGodownID & ") "
'''   End If
'''   tmp = tmp & " Group by MC.CMPName,MG.GodownNo,MG.Address,MU.UnitType,MG.Area,MG.GodownID,MU.UnitTypeID,Isnull(MG.ApprovalLimit,0) "

   tmp = " Select MC.CMPName,MG.GodownNo,MG.Address,MU.UnitType,sum(M1.RentAmount) 'Rent',MG.Area,MG.GodownID,cast('' As Varchar) DepConDID,MU.UnitTypeID,Isnull(MG.ApprovalLimit,0)ApprovalLimit"
   tmp = tmp & " From Mst_Godown MG"
   tmp = tmp & " Inner Join Mst_CMP MC On MG.CMPID=MC.CMPID"
   tmp = tmp & " Inner Join Mst_UnitType MU on MG.UnitTypeID = MU.UnitTypeID"
   tmp = tmp & " Inner Join Txn_GodownLessorAgreeDetail M1 On MG.GodownID=M1.GodownID"
   tmp = tmp & " inner Join Txn_LessorAgreement TLA on M1.LAID = TLA.LAID And TLA.Flag = 'A'"
   tmp = tmp & " Where MG.CloseDate is null  And MG.StartDate<='" & Format(TxtEndDate, "YYYY-MM-DD") & "' And ( '" & TxtStartDate.Value & "' <= TLA.ExpiryDate And '" & TxtEndDate.Value & "' >= TLA.StartDate )"
   tmp = tmp & " And MC.LocationID = " & mLocationID & " And  MU.UnitTypeID = 1"
   If mTmpDepConID <> "" And mValidGodownID <> "" Then
      tmp = tmp & " AND MG.GodownID NOT IN (" & mValidGodownID & ") "
   End If
   
        tmp = tmp & " And Convert(Varchar,MG.GodownID)+''+Convert(Varchar,TLA.ExpiryDate,112) In "
        tmp = tmp & " ( "
        tmp = tmp & "   Select Convert(Varchar,AA.GodownID)+''+Convert(Varchar,Max(BB.ExpiryDate),112) "
        tmp = tmp & "   From Txn_GodownLessorAgreeDetail AA Inner Join Txn_LessorAgreement BB On AA.LAID = BB.LAID "
        tmp = tmp & "   Where Convert(Varchar,AA.LAID)+''+Convert(Varchar,AA.GodownID) "
        tmp = tmp & "   In ( Select Convert(Varchar,E.LAID)+''+Convert(Varchar,D.GodownID) LAIDGodownKey "
        tmp = tmp & "   From Mst_Godown A Inner Join Mst_CMP B On A.CMPID=B.CMPID "
        tmp = tmp & "   Inner Join Mst_UnitType C on A.UnitTypeID = C.UnitTypeID "
        tmp = tmp & "   Inner Join Txn_GodownLessorAgreeDetail D On A.GodownID=D.GodownID "
        tmp = tmp & "   Inner Join Txn_LessorAgreement E on D.LAID = E.LAID And E.Flag = 'A' "
        tmp = tmp & "   Where A.CloseDate is null And A.StartDate<='" & Format(TxtEndDate, "YYYY-MM-DD") & "' And ( '" & TxtStartDate.Value & "' <= E.ExpiryDate And '" & TxtEndDate.Value & "' >= E.StartDate ) "
        tmp = tmp & "   And B.LocationID = " & mLocationID & " And C.UnitTypeID = 1 "
            If mTmpDepConID <> "" And mValidGodownID <> "" Then
               tmp = tmp & " AND A.GodownID NOT IN (" & mValidGodownID & ") "
            End If
        tmp = tmp & "     ) Group By AA.GodownID "
        tmp = tmp & " ) "
   
   tmp = tmp & " Group by MC.CMPName,MG.GodownNo,MG.Address,MU.UnitType,MG.Area,MG.GodownID,MU.UnitTypeID,Isnull(MG.ApprovalLimit,0) "
   tmp = tmp & " Union All"
   tmp = tmp & " Select MC.CMPName,MG.GodownNo,MG.Address,MU.UnitType,0 'Rent',MG.Area,MG.GodownID,cast('' As Varchar) DepConDID,MU.UnitTypeID,Isnull(MG.ApprovalLimit,0)ApprovalLimit"
   tmp = tmp & " From Mst_Godown MG"
   tmp = tmp & " Inner Join Mst_CMP MC On MG.CMPID=MC.CMPID"
   tmp = tmp & " Inner Join Mst_UnitType MU on MG.UnitTypeID = MU.UnitTypeID"
   tmp = tmp & " Where MG.CloseDate is null  And MG.StartDate<='" & Format(TxtEndDate, "YYYY-MM-DD") & "' And '" & Format(TxtEndDate, "YYYY-MM-DD") & "' <= MG.AgreementEndDate "
   tmp = tmp & " And MC.LocationID = " & mLocationID & " And  MU.UnitTypeID <> 1"
   If mTmpDepConID <> "" And mValidGodownID <> "" Then
      tmp = tmp & " AND MG.GodownID NOT IN (" & mValidGodownID & ") "
   End If
   tmp = tmp & " Group by MC.CMPName,MG.GodownNo,MG.Address,MU.UnitType,MG.Area,MG.GodownID,MU.UnitTypeID,Isnull(MG.ApprovalLimit,0) "
   
'''------------------

   Call Tmp2Table

   With MSHFlexGrid4
    If TmpRs2.RecordCount > 0 Then
     .Rows = TmpRs2.RecordCount + 1
       For I = 1 To TmpRs2.RecordCount
           .row = I
           .Col = 0
           .CellFontName = "Wingdings"
           .CellFontSize = 12
           .CellAlignment = flexAlignCenterCenter
           .Text = strUnChecked
           .TextMatrix(I, 1) = IIf(IsNull(TmpRs2!CMPName), 0, TmpRs2!CMPName)
           .TextMatrix(I, 2) = IIf(IsNull(TmpRs2!GodownNo), 0, TmpRs2!GodownNo)
           .TextMatrix(I, 3) = IIf(IsNull(TmpRs2!Address), 0, TmpRs2!Address)
           .TextMatrix(I, 4) = IIf(IsNull(TmpRs2!UnitType), 0, TmpRs2!UnitType)
           .TextMatrix(I, 5) = Format(IIf(IsNull(TmpRs2!Rent), 0, TmpRs2!Rent), "#0.00")
           .TextMatrix(I, 6) = Format(IIf(IsNull(TmpRs2!Area), 0, TmpRs2!Area), "#0.000")
           .TextMatrix(I, 7) = Format(mMinBagSize, "#0.000")
           .TextMatrix(I, 8) = Format(mMaxBagSize, "#0.000")
           
           If TmpRs2!UnitTypeID = 7 And Val(.TextMatrix(I, 5)) = 0 Then
                .TextMatrix(I, 5) = Format(Val(.TextMatrix(I, 6)) * mOwnGodownSqrFeetRate, "#0.00")
           End If
           mCalculatedBaseRent = (Val(.TextMatrix(I, 5)) / Val(.TextMatrix(I, 6)))
           If TmpRs2!UnitTypeID = 1 Or TmpRs2!UnitTypeID = 7 Then
              If TmpRs2!UnitTypeID = 1 Then
                 If mCalculatedBaseRent < mCurrentBaseRentFactor Then
                    mCalculatedBaseRent = mCurrentBaseRentFactor
                 End If
              End If
              If LCase(mFumigationFlag) = "y" Then
'                 .TextMatrix(I, 9) = Format((((Val(.TextMatrix(I, 5)) / Val(.TextMatrix(I, 6))) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorY, "#0.0000")
'                 .TextMatrix(I, 13) = Format((((Val(.TextMatrix(I, 5)) / Val(.TextMatrix(I, 6))) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorY, "#0.0000")
'                 .TextMatrix(I, 18) = Format((((Val(.TextMatrix(I, 5)) / Val(.TextMatrix(I, 6))) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorY, "#0.0000")
              
                 .TextMatrix(I, 9) = Format((((Val(mCalculatedBaseRent)) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorY, "#0.0000")
                 .TextMatrix(I, 13) = Format((((Val(mCalculatedBaseRent)) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorY, "#0.0000")
                 .TextMatrix(I, 18) = Format((((Val(mCalculatedBaseRent)) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorY, "#0.0000")
              
              
              Else
'                 .TextMatrix(I, 9) = Format((((Val(.TextMatrix(I, 5)) / Val(.TextMatrix(I, 6))) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorN, "#0.0000")
'                 .TextMatrix(I, 13) = Format((((Val(.TextMatrix(I, 5)) / Val(.TextMatrix(I, 6))) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorN, "#0.0000")
'                 .TextMatrix(I, 18) = Format((((Val(.TextMatrix(I, 5)) / Val(.TextMatrix(I, 6))) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorN, "#0.0000")
              
                 .TextMatrix(I, 9) = Format((((Val(mCalculatedBaseRent)) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorN, "#0.0000")
                 .TextMatrix(I, 13) = Format((((Val(mCalculatedBaseRent)) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorN, "#0.0000")
                 .TextMatrix(I, 18) = Format((((Val(mCalculatedBaseRent)) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorN, "#0.0000")
              
              End If
           Else
'              .TextMatrix(I, 9) = Val(mCurrentBaseRentFactor) '"0.0000"
'              .TextMatrix(I, 13) = Val(mCurrentBaseRentFactor) '"0.0000"
'              .TextMatrix(I, 18) = Val(mCurrentBaseRentFactor) '"0.0000"
               mCalculatedBaseRent = mCurrentBaseRentFactor
               If LCase(mFumigationFlag) = "y" Then
                 .TextMatrix(I, 9) = Format((((Val(mCalculatedBaseRent)) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorY, "#0.0000")
                 .TextMatrix(I, 13) = Format((((Val(mCalculatedBaseRent)) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorY, "#0.0000")
                 .TextMatrix(I, 18) = Format((((Val(mCalculatedBaseRent)) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorY, "#0.0000")
               Else
                 .TextMatrix(I, 9) = Format((((Val(mCalculatedBaseRent)) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorN, "#0.0000")
                 .TextMatrix(I, 13) = Format((((Val(mCalculatedBaseRent)) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorN, "#0.0000")
                 .TextMatrix(I, 18) = Format((((Val(mCalculatedBaseRent)) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorN, "#0.0000")
               End If
           End If
           
           .TextMatrix(I, 11) = IIf(IsNull(TmpRs2!GodownID), 0, TmpRs2!GodownID)
           .TextMatrix(I, 12) = ""
           .TextMatrix(I, 15) = Format(Val(.TextMatrix(I, 9)), "#0.0000")
           .TextMatrix(I, 16) = Format(Val(.TextMatrix(I, 9)), "#0.0000")
           .TextMatrix(I, 17) = IIf(IsNull(TmpRs2!UnitTypeID), 0, TmpRs2!UnitTypeID)
           .TextMatrix(I, 19) = Format(IIf(Val(TmpRs2!ApprovalLimit) = 0, mCMPMtoAMRMApprovalLimit, TmpRs2!ApprovalLimit), "#0.00")
           .TextMatrix(I, 20) = "S" ''-- System Bags
            mBillingOnePercentage = Format(IIf(Val(.TextMatrix(I, 9)) > 0, Val(.TextMatrix(I, 9)) / 100, 0), "#0.0000")
            mDepositOnePercentage = Format(IIf(Val(.TextMatrix(I, 13)) > 0, Val(.TextMatrix(I, 13)) / 100, 0), "#0.0000")
           .TextMatrix(I, 10) = Format(Val(.TextMatrix(I, 9)) + (mBillingOnePercentage * .TextMatrix(I, 19)), "#0.0000")
           .TextMatrix(I, 14) = Format(Val(.TextMatrix(I, 13)) + (mDepositOnePercentage * .TextMatrix(I, 19)), "#0.0000")
           
           TmpRs2.MoveNext
       Next
    Else
       MsgBox "No Godown Found !!!"
       Exit Sub
    End If
    
    '''-- Billing Cycle / Deposite Biling Cycle Wise Rate Calculation
    If CmbBillingCycle.Text <> "" Then
       For I = 1 To .Rows - 1
               .TextMatrix(I, 9) = Format(CalculateRate(.TextMatrix(I, 15), mBillingRatio, CmbBillingCycle) * Val(.TextMatrix(I, 8)) / 1000, "#0.0000")
               mBillingOnePercentage = Format(Val(.TextMatrix(I, 9)) / 100, "#0.0000")
               .TextMatrix(I, 10) = Format(Val(.TextMatrix(I, 9)) + (mBillingOnePercentage * .TextMatrix(I, 19)), "#0.0000")
       Next
    End If
    If CmbDepositBillingCycle.Text <> "" Then
       For I = 1 To .Rows - 1
           .TextMatrix(I, 13) = Format(CalculateRate(.TextMatrix(I, 16), mDepositBillingRatio, CmbDepositBillingCycle) * Val(.TextMatrix(I, 8)) / 1000, "#0.0000")
           mDepositOnePercentage = Format(Val(.TextMatrix(I, 13)) / 100, "#0.0000")
           .TextMatrix(I, 14) = Format(Val(.TextMatrix(I, 13)) + (mDepositOnePercentage * .TextMatrix(I, 19)), "#0.0000")
       Next
    End If
   End With
End If
End Sub

Public Function CheckGodownSelectOrNot(msh As MSHFlexGrid) As Boolean
Dim Retval As Boolean
For I = 1 To msh.Rows - 1
    If msh.TextMatrix(I, 0) = strChecked Then
        Retval = True
        Exit For
    End If
Next
CheckGodownSelectOrNot = Retval
End Function

Public Function CheckForBagSize(msh As MSHFlexGrid) As Boolean
Dim Retval As Boolean
Dim mR As Double
For mR = 1 To msh.Rows - 1
    If msh.TextMatrix(mR, 0) = strChecked Then
        Select Case LCase(msh.Name)
            Case "mshflexgrid4"
             If Val(msh.TextMatrix(mR, 7)) > 0 And Val(msh.TextMatrix(mR, 8)) > 0 Then
                 Retval = True
             Else
                 Retval = False
                 Exit For
             End If
        End Select
    End If
Next mR
CheckForBagSize = Retval
End Function

Public Function CheckGodownForRate(msh As MSHFlexGrid) As Boolean
Dim Retval As Boolean
Dim mR, mC As Double
For I = 1 To msh.Rows - 1
  If msh.TextMatrix(I, 0) = strChecked Then
    Select Case LCase(msh.Name)
       Case "mshflexgrid1"
             If Val(msh.TextMatrix(I, 8)) > 0 And Val(msh.TextMatrix(I, 12)) > 0 Then
                 Retval = True
             Else
                 Retval = False
                 Exit For
             End If
       Case "mshflexgrid4"
            If Val(msh.TextMatrix(I, 10)) > 0 And Val(msh.TextMatrix(I, 14)) > 0 Then
                Retval = True
            Else
                Retval = False
                Exit For
            End If
    End Select
  End If
Next
CheckGodownForRate = Retval
End Function

Private Sub Grid_Head3()
With MSHFlexGrid3
    .Font.Size = 9
    .Clear
    .Rows = 2
    .FixedRows = 1
    .Cols = 5
    .TextMatrix(0, 0) = "Row ID":                                   .ColWidth(0) = 0
    .TextMatrix(0, 1) = "Old Employee to Whom was Forwarded":       .ColWidth(1) = 3800
    .TextMatrix(0, 2) = "New Employee to Whom is Forwarded":        .ColWidth(2) = 3800
    .TextMatrix(0, 3) = "Updated By Employee":                      .ColWidth(3) = 3800
    .TextMatrix(0, 4) = "Concurrency":                              .ColWidth(4) = 3000
    .RowHeight(0) = 600
End With
End Sub

Private Sub Grid_HeadAttachment()
With MSHFlexGrid5
    .Clear
    .Rows = 2
    .Cols = 5
    .Font.Size = 9
    .WordWrap = True
    .TextMatrix(0, 0) = "SR No":                                    .ColWidth(0) = 1000
    .TextMatrix(0, 1) = "Attachment File Name":                     .ColWidth(1) = 6000
    .TextMatrix(0, 2) = "Remarks":                                  .ColWidth(2) = 7000
    .TextMatrix(0, 3) = "File Path":                                .ColWidth(3) = 0
    .TextMatrix(0, 4) = "ID":                                       .ColWidth(4) = 0
End With
End Sub

Public Function BagSizeRangeCheck(mGodownID_ As Variant, mBagSize As Double, mRowSel_ As Integer) As Boolean
Dim mI As Double
Dim mResult As Boolean
mResult = True
With MSHFlexGrid4
    For mI = 1 To .Rows - 1
           If mI <> mRowSel_ Then
            If .TextMatrix(mI, 11) = mGodownID_ Then  ''And .TextMatrix(mI, 0) = strChecked
               If Val(mBagSize) >= Val(.TextMatrix(mI, 7)) And Val(mBagSize) <= Val(.TextMatrix(mI, 8)) Then
                    mResult = False
                    Exit For
               End If
            End If
           End If
    Next mI
End With
BagSizeRangeCheck = mResult
End Function


Private Sub Text1_Change()

End Sub

Private Sub TxtBagSize_LostFocus()
If TxtBagSize.Text = "" Then TxtBagSize.Visible = False: Exit Sub
If Val(TxtBagSize) < 0 Or Not IsNumeric(TxtBagSize) Then
    MsgBox "Invalid Numeric Format!!!"
    SSTab1.Tab = 0
    TxtBagSize.SelStart = 0
    TxtBagSize.SelLength = Len(TxtBagSize)
    TxtBagSize.SetFocus
    Exit Sub
End If
If Val(TxtBagSize.Text) = 0 Then
     MsgBox "Bag size could be zero."
     TxtBagSize.SetFocus
     Exit Sub
End If
If Check4.Value = 1 Then
    With MSHFlexGrid4
       If Val(TxtBagSize.Text) < Val(mMinBagSize) Or Val(TxtBagSize.Text) > Val(mMaxBagSize) Then
            MsgBox "Bag Size is not as per define Commodity Minimum-Maximum bag size range." & vbNewLine & "Minimum Bag size Limit: " & mMinBagSize & vbNewLine & "Maximum Bag size Limit: " & mMaxBagSize
            TxtBagSize.SetFocus
            Exit Sub
       End If
       If mSelCol = 7 Then  ''-- Min Bag Size
            If .TextMatrix(mSelRow, 8) <> "" Then
               If Val(.TextMatrix(mSelRow, 8)) < Val(TxtBagSize) Then
                    MsgBox "Minimum bag size should not more then Maximum bag size." & vbNewLine & "Minimum Bag size Limit: " & mMinBagSize & vbNewLine & "Maximum Bag size Limit: " & mMaxBagSize
                    TxtBagSize.SetFocus
                    Exit Sub
               End If
               If Val(.TextMatrix(mSelRow, 8)) = Val(TxtBagSize) Then
                    MsgBox "Minimum bag size and Maximim bag size should not be same." & vbNewLine & "Minimum Bag size Limit: " & mMinBagSize & vbNewLine & "Maximum Bag size Limit: " & mMaxBagSize
                    TxtBagSize.SetFocus
                    Exit Sub
               End If
            End If
       ElseIf mSelCol = 8 Then ''-- Max Bag Size
            If .TextMatrix(mSelRow, 7) <> "" Then
               If Val(.TextMatrix(mSelRow, 7)) > Val(TxtBagSize) Then
                    MsgBox "Maximum bag size should not less then Minimum bag size." & vbNewLine & "Minimum Bag size Limit: " & mMinBagSize & vbNewLine & "Maximum Bag size Limit: " & mMaxBagSize
                    TxtBagSize.SetFocus
                    Exit Sub
               End If
               If Val(.TextMatrix(mSelRow, 7)) = Val(TxtBagSize) Then
                    MsgBox "Minimum bag size and Maximim bag size should not be same." & vbNewLine & "Minimum Bag size Limit: " & mMinBagSize & vbNewLine & "Maximum Bag size Limit: " & mMaxBagSize
                    TxtBagSize.SetFocus
                    Exit Sub
               End If
            End If
       End If
       If BagSizeRangeCheck(.TextMatrix(mSelRow, 11), CDbl(TxtBagSize), CInt(mSelRow)) = False Then
            MsgBox "Bag size is already define in given Bag size range for same selected godown." & vbNewLine & "Minimum Bag size Limit: " & mMinBagSize & vbNewLine & "Maximum Bag size Limit: " & mMaxBagSize
            TxtBagSize.SetFocus
            Exit Sub
       End If
       .TextMatrix(mSelRow, mSelCol) = Format(TxtBagSize, "#0.000")
       
       If .TextMatrix(mSelRow, 17) = 1 Or .TextMatrix(mSelRow, 17) = 7 Then
            .TextMatrix(mSelRow, 9) = Format(CalculateRate(.TextMatrix(mSelRow, 15), mBillingRatio, CmbBillingCycle) * Val(.TextMatrix(mSelRow, 8)) / 1000, "#0.0000")
            mBillingOnePercentage = Format(Val(.TextMatrix(mSelRow, 9)) / 100, "#0.0000")
            .TextMatrix(mSelRow, 10) = Format(Val(.TextMatrix(mSelRow, 9)) + (mBillingOnePercentage * .TextMatrix(mSelRow, 19)), "#0.0000")

            .TextMatrix(mSelRow, 13) = Format(CalculateRate(.TextMatrix(mSelRow, 16), mDepositBillingRatio, CmbDepositBillingCycle) * Val(.TextMatrix(mSelRow, 8)) / 1000, "#0.0000")
            mDepositOnePercentage = Format(Val(.TextMatrix(mSelRow, 13)) / 100, "#0.0000")
            .TextMatrix(mSelRow, 14) = Format(Val(.TextMatrix(mSelRow, 13)) + (mDepositOnePercentage * .TextMatrix(mSelRow, 19)), "#0.0000")
       Else
            .TextMatrix(mSelRow, 10) = .TextMatrix(mSelRow, 10)
            .TextMatrix(mSelRow, 14) = .TextMatrix(mSelRow, 14)
       End If
       Call CheckBagRateByColor(CInt(mSelRow))
       
''        .TextMatrix(mSelRow, 9) = Format(CalculateRate(.TextMatrix(mSelRow, 15), mBillingRatio, CmbBillingCycle) * Val(.TextMatrix(mSelRow, 8)) / 1000, "#0.0000")
''        mBillingOnePercentage = Format(Val(.TextMatrix(mSelRow, 9)) / 100, "#0.0000")
''        .TextMatrix(mSelRow, 10) = Format(Val(.TextMatrix(mSelRow, 9)) + (mBillingOnePercentage * .TextMatrix(mSelRow, 19)), "#0.0000")
''
''        .TextMatrix(mSelRow, 13) = Format(CalculateRate(.TextMatrix(mSelRow, 16), mDepositBillingRatio, CmbDepositBillingCycle) * Val(.TextMatrix(mSelRow, 8)) / 1000, "#0.0000")
''        mDepositOnePercentage = Format(Val(.TextMatrix(mSelRow, 13)) / 100, "#0.0000")
''        .TextMatrix(mSelRow, 14) = Format(Val(.TextMatrix(mSelRow, 13)) + (mDepositOnePercentage * .TextMatrix(mSelRow, 19)), "#0.0000")
''        Call CheckBagRateByColor(CInt(mSelRow))
       
''       ''-- Lease / Own Godown Type
''       If .TextMatrix(mSelRow, 17) = 1 Or .TextMatrix(mSelRow, 17) = 7 Then
''            If CmbBillingCycle.Text <> "" Then
''                 .TextMatrix(mSelRow, 8) = Format(CalculateRate(.TextMatrix(mSelRow, 14), mBillingRatio, CmbBillingCycle) * Val(.TextMatrix(mSelRow, 7)) / 1000, "#0.00")
''                 .TextMatrix(mSelRow, 9) = Format(CalculateRate(.TextMatrix(mSelRow, 14), mBillingRatio, CmbBillingCycle) * Val(.TextMatrix(mSelRow, 7)) / 1000, "#0.00")
''            End If
''            If CmbDepositBillingCycle.Text <> "" Then
''                 .TextMatrix(mSelRow, 12) = Format(CalculateRate(.TextMatrix(mSelRow, 15), mDepositBillingRatio, CmbDepositBillingCycle) * Val(.TextMatrix(mSelRow, 7)) / 1000, "#0.00")
''                 .TextMatrix(mSelRow, 13) = Format(CalculateRate(.TextMatrix(mSelRow, 15), mDepositBillingRatio, CmbDepositBillingCycle) * Val(.TextMatrix(mSelRow, 7)) / 1000, "#0.00")
''            End If
''       ElseIf .TextMatrix(mSelRow, 16) = 2 Then
''            If CmbBillingCycle.Text <> "" Then
''                 .TextMatrix(mSelRow, 9) = Format(CalculateRate(.TextMatrix(mSelRow, 14), mBillingRatio, CmbBillingCycle) * Val(.TextMatrix(mSelRow, 7)) / 1000, "#0.00")
''            End If
''            If CmbDepositBillingCycle.Text <> "" Then
''                 .TextMatrix(mSelRow, 13) = Format(CalculateRate(.TextMatrix(mSelRow, 15), mDepositBillingRatio, CmbDepositBillingCycle) * Val(.TextMatrix(mSelRow, 7)) / 1000, "#0.00")
''            End If
''       End If
       
    End With
End If
TxtBagSize.Visible = False
MSHFlexGrid4.Col = MSHFlexGrid4.MouseCol
End Sub

Private Sub TxtEndDate_GotFocus()
tmp = TxtEndDate
End Sub

Private Sub TxtEndDate_LostFocus()
If Edit_Flg = "A" Then
    If CDate(TxtEndDate) < CDate(TxtStartDate) Then
       MsgBox "Deposit Contract 'End Date' should not less then 'Start Date'.", vbInformation + vbOKOnly
       TxtEndDate.SetFocus
       Exit Sub
    End If
    If LCase(Gen_UserRole) <> "power" And mApprovedRights = False Then
       If CDate(TxtEndDate.Value) - CDate(TxtStartDate.Value) > Val(mDepositContractDaysGapStartDateEndDate) Then
          MsgBox "Difference between Start date and End Date should not be more then " & mDepositContractDaysGapStartDateEndDate & " day(s).", vbInformation + vbOKOnly
          TxtEndDate.SetFocus
          Exit Sub
       End If
    End If
End If
If tmp <> TxtEndDate Then
   Call GetGodownDetail
End If
End Sub

Private Sub TxtFinalPrice_LostFocus()
If TxtFinalPrice.Text = "" Then TxtFinalPrice.Visible = False: Exit Sub
If Val(TxtFinalPrice) < 0 Or Not IsNumeric(TxtFinalPrice) Then
   MsgBox "Invalid Numeric Format!!!"
   SSTab1.Tab = 0
   TxtFinalPrice.SelStart = 0
   TxtFinalPrice.SelLength = Len(TxtFinalPrice)
   TxtFinalPrice.SetFocus
   Exit Sub
End If
If Val(TxtFinalPrice) = 0 Then
   MsgBox "Rate should be greater than 0!!!"
   SSTab1.Tab = 0
   TxtFinalPrice.SelStart = 0
   TxtFinalPrice.SelLength = Len(TxtFinalPrice)
   TxtFinalPrice.Visible = True: TxtFinalPrice.SetFocus
   Exit Sub
End If
If Check4.Value = 1 Then
    If MSHFlexGrid4.TextMatrix(mSelRow, 17) <> 1 And MSHFlexGrid4.TextMatrix(mSelRow, 17) <> 7 Then
      If mSelCol = 10 Then MSHFlexGrid4.TextMatrix(mSelRow, 15) = Val(TxtFinalPrice)
      If mSelCol = 14 Then MSHFlexGrid4.TextMatrix(mSelRow, 16) = Val(TxtFinalPrice)
    End If
    ''Call GridRateCheck(MSHFlexGrid4, mSelRow, mSelCol, TxtFinalPrice)
    MSHFlexGrid4.TextMatrix(mSelRow, mSelCol) = Format(TxtFinalPrice, "#0.0000")
    Call CheckBagRateByColor(CInt(mSelRow))
Else
   If MSHFlexGrid1.TextMatrix(mSelRow, 15) <> 1 And MSHFlexGrid1.TextMatrix(mSelRow, 15) <> 7 Then
      If mSelCol = 8 Then MSHFlexGrid1.TextMatrix(mSelRow, 13) = Val(TxtFinalPrice)
      If mSelCol = 12 Then MSHFlexGrid1.TextMatrix(mSelRow, 14) = Val(TxtFinalPrice)
   End If
   ''Call GridRateCheck(MSHFlexGrid1, mSelRow, mSelCol, TxtFinalPrice)
    MSHFlexGrid1.TextMatrix(mSelRow, mSelCol) = Format(TxtFinalPrice, "#0.0000")
   Call CheckRateByColor(CInt(mSelRow))
End If
TxtFinalPrice.Visible = False
If Check4.Value = 1 Then MSHFlexGrid4.Col = MSHFlexGrid4.MouseCol Else MSHFlexGrid1.Col = MSHFlexGrid1.MouseCol
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Txn_DepositContract.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid2.Rows - 1
    mRow = I + 1 ' + 6
    For C = 0 To MSHFlexGrid2.Cols - 1
        If InStr(MSHFlexGrid2.TextMatrix(I, C), "/") > 0 Then
           If IsDate(MSHFlexGrid2.TextMatrix(I, C)) Then
              oWS.Cells(mRow, C + 1) = Format(MSHFlexGrid2.TextMatrix(I, C), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, C + 1) = MSHFlexGrid2.TextMatrix(I, C)
           End If
        Else
           oWS.Cells(mRow, C + 1) = MSHFlexGrid2.TextMatrix(I, C)
        End If
    Next
Next
oWB.Save
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub

Public Sub Xls_GridDetail(Grid As MSHFlexGrid)
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Integer
Dim mCol As Integer
MsgBox ("Wait till next message")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "_Data_Txn_DepositContract.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1

With Grid
    If LCase(Grid.Name) = LCase("MSHFlexGrid1") Then
        For I = 0 To .Rows - 1
            mRow = I + 1
            mCol = 0
            For C = 1 To .Cols - 1
                ''-- Only Visible columns with Colour if any
                If C <> 9 And C <> 10 And C <> 13 And C <> 14 And C <> 15 And C <> 17 Then
                    mCol = mCol + 1
                    oWS.Cells(mRow, mCol) = .TextMatrix(I, C)
                    .row = I
                    .Col = C
                    If .CellBackColor = vbRed Then oWS.Range(oWS.Cells(mRow, mCol), oWS.Cells(mRow, mCol)).Interior.Color = vbRed
                    If .CellBackColor = vbGreen Then oWS.Range(oWS.Cells(mRow, mCol), oWS.Cells(mRow, mCol)).Interior.Color = vbGreen
                    If .CellBackColor = vbYellow Then oWS.Range(oWS.Cells(mRow, mCol), oWS.Cells(mRow, mCol)).Interior.Color = vbYellow
                End If
            Next
        Next
        
    ElseIf LCase(Grid.Name) = LCase("MSHFlexGrid4") Then
        mCol = 0
        For I = 0 To .Rows - 1
            mRow = I + 1
            mCol = 0
            For C = 1 To .Cols - 1
                ''-- Only Visible columns with Colour if any
                If C <> 11 And C <> 12 And C <> 15 And C <> 16 And C <> 17 And C <> 19 And C <> 20 Then
                    mCol = mCol + 1
                    oWS.Cells(mRow, mCol) = .TextMatrix(I, C)
                    .row = I
                    .Col = C
                    If .CellBackColor = vbRed Then oWS.Range(oWS.Cells(mRow, mCol), oWS.Cells(mRow, mCol)).Interior.Color = vbRed
                    If .CellBackColor = vbGreen Then oWS.Range(oWS.Cells(mRow, mCol), oWS.Cells(mRow, mCol)).Interior.Color = vbGreen
                    If .CellBackColor = vbYellow Then oWS.Range(oWS.Cells(mRow, mCol), oWS.Cells(mRow, mCol)).Interior.Color = vbYellow
                End If
            Next
        Next
    
    End If
End With

oWB.Save
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub



Public Function CalculateRate(mmRateOnQty, mmRatio, mmBillingCycle) As Double
Dim Retval As Variant

If LCase(mmBillingCycle) = "daily" Then
   Retval = "30"   'According to Mayur sir, it will be 30 days
ElseIf LCase(mmBillingCycle) = "weekly" Then
   Retval = "4"
ElseIf LCase(mmBillingCycle) = "fortnightly" Then
   Retval = "2"
ElseIf LCase(mmBillingCycle) = "monthly" Then
   Retval = "1"
End If

mmCalculateRate = Format((Val(mmRateOnQty) / Val(Retval)) * Val(mmRatio) / 100, "########0.0000")
CalculateRate = Format((Val(mmRateOnQty) / Val(Retval)) + Val(mmCalculateRate), "########0.0000")


End Function

Public Function CheckBaseRate() As Variant
Dim Retval As Variant
Dim mColor As ColorConstants
Retval = ""
With MSHFlexGrid1
For I = 1 To .Rows - 1
    If .TextMatrix(I, 1) <> "" And .TextMatrix(I, 0) = strChecked Then
       If Val(.TextMatrix(I, 8)) < Val(.TextMatrix(I, 7)) Then
          mColor = vbRed
          .Col = 8: .row = I: .CellBackColor = mColor
       End If
       If Val(.TextMatrix(I, 8)) > Val(.TextMatrix(I, 7)) Then
          mColor = vbGreen
          .Col = 8: .row = I: .CellBackColor = mColor
       End If

       If Val(.TextMatrix(I, 12)) < Val(.TextMatrix(I, 11)) Then
          mColor = vbRed
          .Col = 12: .row = I: .CellBackColor = mColor
       End If
       If Val(.TextMatrix(I, 12)) > Val(.TextMatrix(I, 11)) Then
          mColor = vbGreen
          .Col = 12: .row = I: .CellBackColor = mColor
       End If
       If mColor <> vbRed And mColor <> vbGreen Then
          mColor = vbWhite
          .Col = 8: .row = I: .CellBackColor = mColor
          .Col = 12: .row = I: .CellBackColor = mColor
       End If
       If mColor = vbRed Or mColor = vbGreen Then
          If Retval = "" Then
             Retval = I
          Else
             Retval = Retval & "," & I
          End If
       End If
    End If
Next
End With
If Retval <> "" Then
    Retval = "Please Check " & Retval & " Rows,final price or final deposit price may be less or greater than base storage rate."
End If
CheckBaseRate = Retval
End Function
Public Function CheckBaseRateBags() As Variant
Dim Retval As Variant
Dim mColor As ColorConstants
Retval = ""
With MSHFlexGrid4
For I = 1 To .Rows - 1
    If .TextMatrix(I, 1) <> "" And .TextMatrix(I, 0) = strChecked Then
       If Val(.TextMatrix(I, 9)) < Val(.TextMatrix(I, 8)) Then
          mColor = vbRed
          .Col = 9: .row = I: .CellBackColor = mColor
       End If
       If Val(.TextMatrix(I, 9)) > Val(.TextMatrix(I, 8)) Then
          mColor = vbGreen
          .Col = 9: .row = I: .CellBackColor = mColor
       End If

       If Val(.TextMatrix(I, 13)) < Val(.TextMatrix(I, 12)) Then
          mColor = vbRed
          .Col = 13: .row = I: .CellBackColor = mColor
       End If
       If Val(.TextMatrix(I, 13)) > Val(.TextMatrix(I, 12)) Then
          mColor = vbGreen
          .Col = 13: .row = I: .CellBackColor = mColor
       End If
       If mColor <> vbRed And mColor <> vbGreen Then
          mColor = vbWhite
          .Col = 9: .row = I: .CellBackColor = mColor
          .Col = 13: .row = I: .CellBackColor = mColor
       End If
       If mColor = vbRed Or mColor = vbGreen Then
          If Retval = "" Then
             Retval = I
          Else
             Retval = Retval & "," & I
          End If
       End If
    End If
Next
End With
If Retval <> "" Then
    Retval = "Please Check " & Retval & " Rows,final price or final deposit price may be less or greater than base storage rate."
End If
CheckBaseRateBags = Retval
End Function

Public Sub InDataLogHistory()
Call Grid_Head3

tmp = " SELECT LTDC.LogDepConID,TDC.DepConID,LTDC.EmployeeID,EMP.EmployeeName+'~'+EMP.EmployeeNo[OldEmployee],"
tmp = tmp & " TDC.EmployeeID,EMP1.EmployeeName+'~'+EMP1.EmployeeNo[NewEmployee],EMP2.EmployeeName+'~'+EMP2.EmployeeNo[UpdateBy],LTDC.UpdatedDate"
tmp = tmp & " FROM Log_Txn_DepositContract LTDC"
tmp = tmp & " Inner Join Txn_DepositContract TDC On LTDC.DepConID = TDC.DepConID"
tmp = tmp & " Left Join Mst_Employee EMP On LTDC.EmployeeId = EMP.EmployeeID"
tmp = tmp & " Left Join Mst_Employee EMP1 On TDC.EmployeeId = EMP1.EmployeeID"
tmp = tmp & " Left Join Mst_Employee EMP2 On LTDC.UpdatedBy = EMP2.EmployeeNo"
tmp = tmp & " Where TDC.DepConID = " & TxtDepositContractID & " "
tmp = tmp & " Order BY LTDC.LogDepConID DESC"
Call TmpTable
For iu = 1 To TmpRs.RecordCount
    With MSHFlexGrid3
    
    .TextMatrix(iu, 0) = IIf(IsNull(TmpRs!DepConID), "", TmpRs!DepConID)
    .TextMatrix(iu, 1) = IIf(IsNull(TmpRs!OldEmployee), "", TmpRs!OldEmployee)
    .TextMatrix(iu, 2) = IIf(IsNull(TmpRs!NewEmployee), "", TmpRs!NewEmployee)
    .TextMatrix(iu, 3) = IIf(IsNull(TmpRs!UpdateBy), "", TmpRs!UpdateBy)
    .TextMatrix(iu, 4) = IIf(IsNull(TmpRs!UpdatedDate), "", Format(TmpRs!UpdatedDate, "dd-mmm-yyyy hh:mm:ss"))
    .Rows = .Rows + 1
    
    End With
    TmpRs.MoveNext
Next iu

End Sub

Public Sub GetOldBillingCycleData()

tmp = " Select TDCG.DepConDID,TDCG.FinalPrice,TDCG.DepositFinalPrice"
tmp = tmp & " From Txn_DepositContractGodownDetail TDCG"
tmp = tmp & " Inner Join Mst_Godown MG on TDCG.GodownID=MG.GodownID"
tmp = tmp & " Inner Join Mst_CMP MC on MG.CMPID=MC.CMPID"
tmp = tmp & " Where TDCG.DepConID=" & RsTxn_DepositContract!DepConID & ""

Call Tmp2Table
If TmpRs2.RecordCount > 0 Then
   MSHFlexGrid1.Rows = TmpRs2.RecordCount + 2
   For I = 1 To TmpRs2.RecordCount
       For r = 1 To MSHFlexGrid1.Rows - 1
           If Val(TmpRs2!DepConDID) = Val(MSHFlexGrid1.TextMatrix(r, 9)) Then
              MSHFlexGrid1.TextMatrix(r, 7) = IIf(IsNull(TmpRs2!FinalPrice), "", TmpRs2!FinalPrice)
           End If
       Next
   TmpRs2.MoveNext
   Next
End If

End Sub
Public Sub GetOldDepositBillingCycleData()
tmp = " Select TDCG.DepConDID,TDCG.FinalPrice,TDCG.DepositFinalPrice"
tmp = tmp & " From Txn_DepositContractGodownDetail TDCG"
tmp = tmp & " Inner Join Mst_Godown MG on TDCG.GodownID=MG.GodownID"
tmp = tmp & " Inner Join Mst_CMP MC on MG.CMPID=MC.CMPID"
tmp = tmp & " Where TDCG.DepConID=" & RsTxn_DepositContract!DepConID & ""

Call Tmp2Table
If TmpRs2.RecordCount > 0 Then
   MSHFlexGrid1.Rows = TmpRs2.RecordCount + 2
   For I = 1 To TmpRs2.RecordCount
       For r = 1 To MSHFlexGrid1.Rows - 1
           If Val(TmpRs2!DepConDID) = Val(MSHFlexGrid1.TextMatrix(r, 9)) Then
              MSHFlexGrid1.TextMatrix(r, 10) = IIf(IsNull(TmpRs2!DepositFinalPrice), "", TmpRs2!DepositFinalPrice)
              Exit For
           End If
       Next
   TmpRs2.MoveNext
   Next
End If

End Sub


Public Sub CheckBoxChecked(msh As MSHFlexGrid)
If Check2.Value = 1 Then
   For I = 1 To msh.Rows - 1
       If msh.TextMatrix(I, 1) <> "" Then
          msh.row = I
          msh.Col = 0
          msh.CellFontName = "Wingdings"
          msh.CellFontSize = 12
          msh.CellAlignment = flexAlignCenterCenter
          msh.Text = strChecked
      End If
    Next
Else
   For I = 1 To msh.Rows - 1
       If msh.TextMatrix(I, 1) <> "" Then
          msh.row = I
          msh.Col = 0
          msh.CellFontName = "Wingdings"
          msh.CellFontSize = 12
          msh.CellAlignment = flexAlignCenterCenter
          msh.Text = strUnChecked
          If LCase(msh.Name) = LCase("MSHFlexGrid1") Then
            Call ResetCell(MSHFlexGrid1, 8, 12, I)
          ElseIf LCase(msh.Name) = LCase("MSHFlexGrid4") Then
            Call ResetCell(MSHFlexGrid4, 10, 14, I)
          End If
       End If
  Next
End If
End Sub

Public Sub GridRateCheck(msh As MSHFlexGrid, mmSelRow As Variant, mmSelCol As Variant, mValue As Double)
If LCase(msh.Name) = LCase("MSHFlexGrid1") Then
    Call CheckRateByColor(CInt(mmSelRow))
''   Call CheckBaseRate
''   If mmSelCol = 8 Then
''      If Val(msh.TextMatrix(mmSelRow, mmSelCol)) > Val(msh.TextMatrix(mmSelRow, 6)) Then
''         MsgBox "Final Price greater than base storage rate!!!"
''      ElseIf Val(msh.TextMatrix(mmSelRow, mmSelCol)) < Val(TxtMonthlyMTRate.Text) Then
''         MsgBox "Final Price less than base storage rate!!!"
''      End If
''   End If
''   If mmSelCol = 10 Then
''      If Val(msh.TextMatrix(mmSelRow, mmSelCol)) > Val(msh.TextMatrix(mmSelRow, 9)) Then
''         MsgBox "Final Price greater than base storage rate!!!"
''      ElseIf Val(msh.TextMatrix(mmSelRow, mmSelCol)) < Val(TxtDepMonthlyMTRate.Text) Then
''         MsgBox "Final Price less than base storage rate!!!"
''      End If
''   End If
Else
    Call CheckBagRateByColor(CInt(mmSelRow))
''   Call CheckBaseRateBags
''   If mmSelCol = 8 Then
''      If Val(msh.TextMatrix(mmSelRow, mmSelCol)) > Val(msh.TextMatrix(mmSelRow, 7)) Then
''         MsgBox "Final Price greater than base storage rate!!!"
''      ElseIf Val(msh.TextMatrix(mmSelRow, mmSelCol)) < Val(TxtMonthlyMTRate.Text) Then
''         MsgBox "Final Price less than base storage rate!!!"
''      End If
''   End If
''   If mmSelCol = 11 Then
''      If Val(msh.TextMatrix(mmSelRow, mmSelCol)) > Val(msh.TextMatrix(mmSelRow, 7)) Then
''         MsgBox "Final Price greater than base storage rate!!!"
''      ElseIf Val(msh.TextMatrix(mmSelRow, mmSelCol)) < Val(TxtDepMonthlyMTRate.Text) Then
''         MsgBox "Final Price less than base storage rate!!!"
''      End If
''   End If
End If
End Sub

Public Function CheckApprovalNeed(msh As MSHFlexGrid, mmUnitTypeID As Variant, mC1 As Variant, mC2 As Variant) As Boolean
Dim Retval As Boolean
Retval = True
For I = 1 To msh.Rows - 1
    If msh.TextMatrix(I, 0) = strChecked Then
        msh.Col = mC1
        msh.row = I
        If msh.CellBackColor = vbRed Then
           Retval = False
           Exit For
        End If
        msh.Col = mC2
        msh.row = I
        If msh.CellBackColor = vbRed Then
           Retval = False
           Exit For
        End If
        If Val(msh.TextMatrix(I, mmUnitTypeID)) <> 1 Then
''            If Val(msh.TextMatrix(I, mC1 - 1)) = Val(msh.TextMatrix(I, mC1)) And Val(msh.TextMatrix(I, mC2 - 1)) = Val(msh.TextMatrix(I, mC2)) Then
''                Retval = False
''                Exit For
''            ElseIf Val(msh.TextMatrix(I, mC1 - 1)) = Val(msh.TextMatrix(I, mC1)) And Val(msh.TextMatrix(I, mC2 - 1)) = Val(msh.TextMatrix(I, mC2)) Then
''
''            End If
            
''            If Val(msh.TextMatrix(I, mC1 - 1)) = 0 And Val(msh.TextMatrix(I, mC1)) > 0 Then
''                Retval = False
''                Exit For
''            End If
''            If Val(msh.TextMatrix(I, mC2 - 1)) = 0 And Val(msh.TextMatrix(I, mC2)) > 0 Then
''                Retval = False
''                Exit For
''            End If
        
        End If
    End If
Next
CheckApprovalNeed = Retval
End Function

Public Sub ClearData()
CmbCommodity.Text = ""
TxtBaseRate.Text = ""
Call Grid_Head
Call Grid_Head4
End Sub

Public Sub GetGodownDetail()
If CmbContractType.Text <> "" And CmbState.Text <> "" And CmbLocation.Text <> "" And CmbCommodity.Text <> "" And CmbClient.Text <> "" Then
   If (TxtStartDate.Value <> "" And TxtEndDate.Value <> "") Then
      Call GetGodownDetails
   End If
Else
   Call Grid_Head
   Call Grid_Head4
End If
End Sub

Private Sub TxtMaxNewBagSize_LostFocus()
If Check4.Value = 1 Then
    If Trim(TxtMaxNewBagSize) = "" Then Exit Sub
    If Val(TxtMaxNewBagSize) <= 0 Or Not IsNumeric(TxtMaxNewBagSize) Then
        MsgBox "Invalid Numeric Format!!!"
        TxtMaxNewBagSize.SelStart = 0
        TxtMaxNewBagSize.SelLength = Len(TxtMaxNewBagSize)
        TxtMaxNewBagSize.SetFocus
        Exit Sub
    End If
End If
End Sub

Private Sub TxtMinNewBagSize_LostFocus()
If Check4.Value = 1 Then
    If Trim(TxtMinNewBagSize) = "" Then Exit Sub
    If Val(TxtMinNewBagSize) <= 0 Or Not IsNumeric(TxtMinNewBagSize) Then
        MsgBox "Invalid Numeric Format!!!"
        TxtMinNewBagSize.SelStart = 0
        TxtMinNewBagSize.SelLength = Len(TxtMinNewBagSize)
        TxtMinNewBagSize.SetFocus
        Exit Sub
    End If
End If
End Sub


''-- 14 Feb 2018
Private Sub TxtSDepositContractID_LostFocus()
If TxtSDepositContractID.Text = "" Then Exit Sub
If IsNumeric(TxtSDepositContractID.Text) = False Or Val(TxtSDepositContractID.Text) < 0 Then
   MsgBox "Invalid Numeric Value !!"
   TxtSDepositContractID.SetFocus
   Exit Sub
End If
End Sub
''---------------

Private Sub TxtStartDate_GotFocus()
tmp = TxtStartDate
End Sub

Private Sub TxtStartDate_LostFocus()
If Edit_Flg = "A" Then
    If LCase(CmbContractType.Text) = "normal" Then  ''-- 14 Feb 2018
        If LCase(mDepositContractInFutureDateFlag) = "n" Then
            If CDate(TxtStartDate) > CDate(Date) Then
               MsgBox "New Deposit Contract can not create in Future Date.", vbInformation + vbOKOnly
               TxtStartDate.SetFocus
               Exit Sub
            End If
        End If
        ''-- Rights applyed on 05 Jan 2018
        'If LCase(Gen_UserRole) <> "superadmin" And LCase(Gen_UserRole) <> "head cmg" And LCase(Gen_UserRole) <> "power" Then
        If LCase(Gen_UserRole) <> "power" And mApprovedRights = False Then

            
            'If CDate(TxtEndDate.Value) - CDate(TxtStartDate.Value) > mDepositContractDaysGapStartDateEndDate Then
            '    MsgBox "Difference between Start date and End Date should not be more then " & mDepositContractDaysGapStartDateEndDate & " day(s).", vbInformation + vbOKOnly
            '    TxtStartDate.SetFocus
            '    Exit Sub
            'End If
            If CDate(TxtEndDate.Value) - CDate(TxtStartDate.Value) > mDepositContractDaysGapStartDateEndDate Then
               TxtEndDate.Value = TxtStartDate.Value + Val(mDepositContractDaysGapStartDateEndDate)
            End If
        
            If CDate(TxtStartDate) < CDate(Date) - mDepositContractAllowBackDateEntryDays Then
                MsgBox "Start date should not prior more then " & mDepositContractAllowBackDateEntryDays & " day.", vbInformation + vbOKOnly
                TxtStartDate.SetFocus
                Exit Sub
            End If
        
        End If
        
        
        
    ElseIf LCase(CmbContractType.Text) = "revision" Then  ''-- 14 Feb 2018
        If CDate(TxtStartDate) >= CDate(Date) Then
           MsgBox "Start Date should be past date for Revision Deposit Contract.", vbInformation + vbOKOnly
           TxtStartDate.SetFocus
           Exit Sub
        End If
    End If
End If

If tmp <> TxtStartDate Then
   Call GetGodownDetail
End If
End Sub

Private Sub TxtTentativeDuration_LostFocus()
If TxtTentativeDuration.Text = "" Then Exit Sub
If IsNumeric(TxtTentativeDuration.Text) = False Or Val(TxtTentativeDuration.Text) < 0 Then
   MsgBox "Invalid Numeric Value !!"
   TxtTentativeDuration.SetFocus
   Exit Sub
End If
End Sub

Private Sub TxtLockInPeriod_LostFocus()
If TxtLockInPeriod.Text = "" Then Exit Sub
If IsNumeric(TxtLockInPeriod.Text) = False Or Val(TxtLockInPeriod.Text) < 0 Then
   MsgBox "Invalid Numeric Value !!"
   TxtLockInPeriod.SetFocus
   Exit Sub
End If
End Sub

Private Sub TxtQuantum_LostFocus()
If TxtQuantum.Text = "" Then Exit Sub
If IsNumeric(TxtQuantum.Text) = False Or Val(TxtQuantum.Text) < 0 Then
   MsgBox "Invalid Numeric Value !!"
   TxtQuantum.SetFocus
   Exit Sub
End If
End Sub

Private Function GetMaxAttachID() As Double
Dim Retval As Double
tmp = "Select max(AttachmentID) from Txn_DepositContractAttachment"
Call TmpTableSMAttachment
Retval = IIf(IsNull(TmpRsSM_Attachment.Fields(0)), 0, TmpRsSM_Attachment.Fields(0))
Retval = Retval + 1
GetMaxAttachID = Retval
End Function

Private Sub Gen_Email()
Dim TOADD As Variant
Dim TOCC, mMsg, mGridMsg As Variant
Dim strSplitString() As String
Dim EmpNo_, EmpName_ As String
Dim mR As Variant
Dim mColor As Variant

mMsg = ""
TOADD = ""
TOCC = ""

mMsg = "<table border=1 width='100%' align='center'>"
mMsg = mMsg & "<tr><td width='40%'><b>Deposit Contract ID</b></td><td width='60%'>" & TxtDepositContractID.Text & "</td></tr>"
mMsg = mMsg & "<tr><td width='40%'><b>State</b></td><td width='60%'>" & CmbState.Text & "</td></tr>"
mMsg = mMsg & "<tr><td width='40%'><b>Location</b></td><td width='60%'>" & CmbLocation.Text & "</td></tr>"
mMsg = mMsg & "<tr><td width='40%'><b>Commodity</b></td><td width='60%'>" & CmbCommodity.Text & "</td></tr>"
mMsg = mMsg & "<tr><td width='40%'><b>Client Name</b></td><td width='60%'>" & CmbClient.Text & "</td></tr>"
mMsg = mMsg & "<tr><td width='40%'><b>Start Date</b></td><td width='60%'>" & Format(TxtStartDate, "DD-MMM-yyyy") & "</td></tr>"
mMsg = mMsg & "<tr><td width='40%'><b>End Date</b></td><td width='60%'>" & Format(TxtEndDate, "DD-MMM-yyyy") & "</td></tr>"
mMsg = mMsg & "<tr><td width='40%'><b>Insured By</b></td><td width='60%'>" & CmbInsurance.Text & "</td></tr>"
mMsg = mMsg & "<tr><td width='40%'><b>Billing Cycle</b></td><td width='60%'>" & CmbBillingCycle.Text & "</td></tr>"
mMsg = mMsg & "<tr><td width='40%'><b>Depository Billing Cycle</b></td><td width='60%'>" & CmbDepositBillingCycle.Text & "</td></tr>"
mMsg = mMsg & "<tr><td width='40%'><b>Status</b></td><td width='60%'>" & CmbStatus.Text & "</td></tr>"
mMsg = mMsg & "<tr><td width='40%'><b>Forwared TO</b></td><td width='60%'>" & CmbForwardedTo.Text & "</td></tr>"
mMsg = mMsg & "<tr><td width='40%'><b>Remarks</b></td><td width='60%'>" & TxtRemarks & "</td></tr>"
mMsg = mMsg & "</table>"
mMsg = mMsg & "<table border=1 width='100%' align='center'>"
mMsg = mMsg & "<tr><td><b>CMP Name</b></td>"
mMsg = mMsg & "<td><b>Godown No</b></td>"
mMsg = mMsg & "<td><b>Address</b></td>"
mMsg = mMsg & "<td><b>Godown Type</b></td>"
mMsg = mMsg & "<td><b>Rent</b> </td>"
mMsg = mMsg & "<td><b>Area in Sq. Ft.</b></td>"
If Check4.Value = vbChecked Then
   mMsg = mMsg & "<td><b>Bag Size (Minimum)</b></td>"
   mMsg = mMsg & "<td><b>Bag Size (Maximum)</b></td>"
End If
mMsg = mMsg & "<td><b>Billing Rate/MT (As Per System)</b></td>"
mMsg = mMsg & "<td><b>Billing Rate/MT (Proposed / Approved)</b></td>"
mMsg = mMsg & "<td><b>Deposit/ Withdrawal Rate/MT (As Per System)</b></td>"
mMsg = mMsg & "<td><b>Deposit/ Withdrawal Rate/MT (Proposed / Approved)</b></td>"
mMsg = mMsg & "<td><b>Billing Rate/MT (Monthly Cycle)</b></td>"
mMsg = mMsg & "</tr>"
mGridMsg = ""
If Check4.Value = vbUnchecked Then
    With MSHFlexGrid1
        mGridMsg = ""
        For mR = 1 To .Rows - 1
          If .TextMatrix(mR, 0) = strChecked Then
            mGridMsg = mGridMsg & "<tr>"
            mGridMsg = mGridMsg & "<td>" & .TextMatrix(mR, 1) & "</td>"
            mGridMsg = mGridMsg & "<td>" & .TextMatrix(mR, 2) & "</td>"
            mGridMsg = mGridMsg & "<td>" & .TextMatrix(mR, 3) & "</td>"
            mGridMsg = mGridMsg & "<td>" & .TextMatrix(mR, 4) & "</td>"
            mGridMsg = mGridMsg & "<td>" & .TextMatrix(mR, 5) & "</td>"
            mGridMsg = mGridMsg & "<td>" & .TextMatrix(mR, 6) & "</td>"
            mGridMsg = mGridMsg & "<td>" & .TextMatrix(mR, 7) & "</td>"
            .Col = 8: .row = mR: mColor = IIf(.CellBackColor = vbRed, "#ff0000", IIf(.CellBackColor = vbYellow, "#FFF700", IIf(.CellBackColor = vbGreen, "#01CA0A", "#FFFFFF")))
            mGridMsg = mGridMsg & "<td bgcolor=" & mColor & " >" & .TextMatrix(mR, 8) & "</td>"
            mGridMsg = mGridMsg & "<td>" & .TextMatrix(mR, 11) & "</td>"
            .Col = 12: .row = mR: mColor = IIf(.CellBackColor = vbRed, "#ff0000", IIf(.CellBackColor = vbYellow, "#FFF700", IIf(.CellBackColor = vbGreen, "#01CA0A", "#FFFFFF")))
            mGridMsg = mGridMsg & "<td bgcolor=" & mColor & ">" & .TextMatrix(mR, 12) & "</td>"
            mGridMsg = mGridMsg & "<td>" & .TextMatrix(mR, 16) & "</td>"
            mGridMsg = mGridMsg & "</tr>"
          End If
        Next mR
    End With
ElseIf Check4.Value = vbChecked Then
    With MSHFlexGrid4
        mGridMsg = ""
        For mR = 1 To .Rows - 1
          If .TextMatrix(mR, 0) = strChecked Then
            mGridMsg = mGridMsg & "<tr>"
            mGridMsg = mGridMsg & "<td>" & .TextMatrix(mR, 1) & "</td>"
            mGridMsg = mGridMsg & "<td>" & .TextMatrix(mR, 2) & "</td>"
            mGridMsg = mGridMsg & "<td>" & .TextMatrix(mR, 3) & "</td>"
            mGridMsg = mGridMsg & "<td>" & .TextMatrix(mR, 4) & "</td>"
            mGridMsg = mGridMsg & "<td>" & .TextMatrix(mR, 5) & "</td>"
            mGridMsg = mGridMsg & "<td>" & .TextMatrix(mR, 6) & "</td>"
            mGridMsg = mGridMsg & "<td>" & .TextMatrix(mR, 7) & "</td>"
            mGridMsg = mGridMsg & "<td>" & .TextMatrix(mR, 8) & "</td>"
            mGridMsg = mGridMsg & "<td>" & .TextMatrix(mR, 9) & "</td>"
            .Col = 10: .row = mR: mColor = IIf(.CellBackColor = vbRed, "#ff0000", IIf(.CellBackColor = vbYellow, "#FFF700", IIf(.CellBackColor = vbGreen, "#01CA0A", "#FFFFFF")))
            mGridMsg = mGridMsg & "<td bgcolor=" & mColor & ">" & .TextMatrix(mR, 10) & "</td>"
            mGridMsg = mGridMsg & "<td>" & .TextMatrix(mR, 13) & "</td>"
            .Col = 14: .row = mR: mColor = IIf(.CellBackColor = vbRed, "#ff0000", IIf(.CellBackColor = vbYellow, "#FFF700", IIf(.CellBackColor = vbGreen, "#01CA0A", "#FFFFFF")))
            mGridMsg = mGridMsg & "<td bgcolor=" & mColor & ">" & .TextMatrix(mR, 14) & "</td>"
            mGridMsg = mGridMsg & "<td>" & .TextMatrix(mR, 18) & "</td>"
            mGridMsg = mGridMsg & "</tr>"
          End If
        Next mR
    End With
End If
mMsg = mMsg & mGridMsg
mMsg = mMsg & "</table>"
''TOADD = "ullhas.p@ncml.com"

If mForwardID <> Gen_UserID Then
    TOADD = GetEmployeeEmailID(mForwardID)
Else
    tmp = " Select EmployeeID,EmployeeNo,EmployeeEmailID from Mst_Employee Where EmployeeNo = '" & IIf(IsNull(RsTxn_DepositContract!CreatedBy), 0, RsTxn_DepositContract!CreatedBy) & "' And Isnull(Flag,'Y') = 'Y'  "
    Call TmpTable
    If TmpRs.RecordCount > 0 Then
        TOADD = IIf(IsNull(TmpRs!EmployeeEmailID), "", TmpRs!EmployeeEmailID)
    End If
End If

TOCC = Gen_UserEmailID
''---- Location Wise Email and Reporty
tmp = " Select Distinct ME.EmployeeEmailID[First],MEE.EmployeeEmailID[Second],MEEE.EmployeeEmailID[Third],MEEEE.EmployeeEmailID[Fourth]"
tmp = tmp & " From Mst_EmpCMPMapping MEC"
tmp = tmp & " Inner Join Mst_Employee ME on MEC.EmployeeID = ME.EmployeeID And ISnull(ME.Flag,'Y') = 'Y' And ME.DesignationID <> 1"
tmp = tmp & " Left Join Mst_Employee MEE On ME.DirectReportingTo = MEE.EmployeeID And ISnull(MEE.Flag,'Y') = 'Y' And MEE.DesignationID <> 1"
tmp = tmp & " Left Join Mst_Employee MEEE On MEE.DirectReportingTo = MEEE.EmployeeID And ISnull(MEEE.Flag,'Y') = 'Y' And MEEE.EmployeeID <> 20"
tmp = tmp & " Left Join Mst_Employee MEEEE On MEEE.DirectReportingTo = MEEEE.EmployeeID And ISnull(MEEEE.Flag,'Y') = 'Y'  And MEEE.EmployeeID <> 20"
tmp = tmp & " Where MEC.LocationID = " & mLocationID & ""
Call TmpTable

For I = 1 To TmpRs.RecordCount
    If InStr(LCase(TOCC), LCase(IIf(IsNull(TmpRs!First), "", TmpRs!First))) = 0 Then TOCC = TOCC & "," & IIf(IsNull(TmpRs!First), "", TmpRs!First)
    If InStr(LCase(TOCC), LCase(IIf(IsNull(TmpRs!Second), "", TmpRs!Second))) = 0 Then TOCC = TOCC & "," & IIf(IsNull(TmpRs!Second), "", TmpRs!Second)
    If InStr(LCase(TOCC), LCase(IIf(IsNull(TmpRs!Third), "", TmpRs!Third))) = 0 Then TOCC = TOCC & "," & IIf(IsNull(TmpRs!Third), "", TmpRs!Third)
    If InStr(LCase(TOCC), LCase(IIf(IsNull(TmpRs!Fourth), "", TmpRs!Fourth))) = 0 Then TOCC = TOCC & "," & IIf(IsNull(TmpRs!Fourth), "", TmpRs!Fourth)
    TmpRs.MoveNext
Next I
''---------

''-- Ullhas Patil 03 July 2018
''----Add The CreatedBy By In TOCC Mail List for future reference Mail
tmp = " Select EmployeeID,EmployeeNo,EmployeeEmailID from Mst_Employee Where EmployeeNo = '" & IIf(IsNull(RsTxn_DepositContract!CreatedBy), 0, RsTxn_DepositContract!CreatedBy) & "' And Isnull(Flag,'Y') = 'Y' "
Call TmpTable
If TmpRs.RecordCount > 0 Then
    If InStr(LCase(TOCC), LCase(IIf(IsNull(TmpRs!EmployeeEmailID), "", TmpRs!EmployeeEmailID))) = 0 Then TOCC = TOCC & "," & IIf(IsNull(TmpRs!EmployeeEmailID), "", TmpRs!EmployeeEmailID)
End If
''---------

''tmp = "Select EmployeeID,EmployeeNo,EmployeeEmailID From Mst_Employee Where EmployeeID IN (" & mUpperHierarchyEmployeeIDList & ") And Flag='Y' "
''tmp = tmp & " Union "
''tmp = tmp & " Select ME.EmployeeID,ME.EmployeeNo,ME.EmployeeEmailID from Mst_Employee ME Inner Join Mst_Users MU on Me.EmployeeNo=MU.Employee_ID Where MU.UserGroupID in (" & mDepositContractCMGGroupToBeApproved & ") And ME.Flag='Y' "


''tmp = " Select ME.EmployeeID,ME.EmployeeNo,ME.EmployeeEmailID from Mst_Employee ME Inner Join Mst_Users MU on Me.EmployeeNo=MU.Employee_ID Where MU.UserGroupID in (" & mDepositContractCMGGroupToBeApproved & ") And ME.Flag='Y' "
''Call TmpTable
''For i = 1 To TmpRs.RecordCount
''    If InStr(LCase(TOCC), LCase(IIf(IsNull(TmpRs!EmployeeEmailID), "", TmpRs!EmployeeEmailID))) = 0 Then
''        TOCC = TOCC & "," & IIf(IsNull(TmpRs!EmployeeEmailID), "", TmpRs!EmployeeEmailID)
''    End If
''    TmpRs.MoveNext
''Next i
''-- 11 Jan 2018

If TOCC <> "" Then
    TOCC = TOCC & ",CMG@ncml.com"
End If

'TOCC = "parag.r@ncml.com"
'If TOCC <> "" Then
'   TOCC = TOCC & "," & Gen_UserEmailID
'Else
'   TOCC = Gen_UserEmailID
'End If

mMsgSubject = "Deposit Contract ID " & TxtDepositContractID
Call SendMail(TOADD, TOCC, mMsgSubject, mMsg, "")
End Sub

Private Function SendMail(ToADD_ As Variant, TOCC_ As Variant, MsgSubject_ As Variant, Message1_ As Variant, mMailTypeFlg As Variant)
Dim htmlStringhead As String
Dim htmlStringtail As String
Set oSMTP = New OSSMTP.SMTPSession
servertxt = "webmail.ncml.com"
htmlStringhead = "<html>"
htmlStringhead = htmlStringhead & " <head>"
htmlStringhead = htmlStringhead & "<style> body{ color: #000000;margin: 5px;font-family: Candara;} table, th, td { border: 1px solid black; border-collapse: collapse;} th { padding: 5px; text-align: center;} td { padding: 5px; text-align: left;} </style>"
htmlStringhead = htmlStringhead & "</head><body>"

htmlStringtail = "</body></html>"
   
''With oSMTP
''  'connection
''  .Server = servertxt
''  If LCase(mMailTypeFlg) = "n" Then
''     .MailFrom = "info@ncmsl.com"
''  Else
''     '.MailFrom = Gen_UserEmailID
''     .MailFrom = "ullhas.p@ncml.com"
''  End If
'''    .SendTo = "ullhas.p@ncml.com"
'''    .CC = Gen_UserEmailID
''  .SendTo = ToADD_
''  If ToADD_ <> TOCC_ Then
''     .CC = TOCC_
''  End If
''  If LCase(mMailTypeFlg) = "n" Then
''     .Attachments.Add (mPDFFileName)
''  End If
''  .MessageSubject = MsgSubject_ & " (from Store-Man)" '"Confidential Mail"
''  .MessageHTML = htmlStringhead & Message1_ & htmlStringtail
''  .TimeStamp = Now()
''  If LCase(Gen_EmailConfigFlag) = "y" Then
''     .SendEmail
''  End If
''End With

Dim objSMTP
Set objSMTP = CreateObject("MailBee.SMTP")
objSMTP.LicenseKey = "MBC900-5F470FAD67-D8E202A2E22927335F24F5A667BFD196"
objSMTP.ServerName = Gen_ReadMailServerIPINI("MailIP") ' = "mail.ncml.com" '"mail.ncml.com"
If LCase(mMailTypeFlg) = "n" Then
   objSMTP.FromAddr = "info@ncmsl.com"
Else
   objSMTP.FromAddr = Gen_UserEmailID
End If
If IsNull(Trim(ToADD_)) = True Or Trim(ToADD_) = "" Then
   ToADD_ = "manu.ag@ncml.com"
End If
objSMTP.ToAddr = ToADD_ '"parag.r@ncml.com" 'ToADD_
If ToADD_ <> TOCC_ Then
   objSMTP.CCAddr = TOCC_ '"ullhas.p@ncml.com" 'TOCC_
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

Public Function GetCommodityBagEquivalent(CID_ As Variant) As String
Dim Retval As Variant
Retval = 0

tmp = "Select CommodityID , MAX(Equivalent) 'Equivalent' from Mst_BagEquivalent"
tmp = tmp & " Where CommodityID = " & CID_ & ""
tmp = tmp & " Group By CommodityID"
Call TmpTable

If Not TmpRs.EOF Then
      Retval = IIf(IsNull(TmpRs!Equivalent), 1, TmpRs!Equivalent)
End If

GetCommodityBagEquivalent = Retval
End Function

Public Function GetBillingCycleID(CID_ As Variant)
Call TableMst_BillingCycle(" Where Flag<>'R' And BillingCycle = '" & CID_ & "'")
If Not RsMst_BillingCycle.EOF Then
       mBillingCycleID = IIf(IsNull(RsMst_BillingCycle!BillingCycleID), 0, RsMst_BillingCycle!BillingCycleID)
       mBillingRatio = IIf(IsNull(RsMst_BillingCycle!Ratio), 0, RsMst_BillingCycle!Ratio)
       mBillingCycleNoofDays = IIf(IsNull(RsMst_BillingCycle!NoOfDays), 0, RsMst_BillingCycle!NoOfDays)
End If
End Function

Public Function GetDepositeBillingCycleID(CID_ As Variant)
Call TableMst_BillingCycle(" Where Flag<>'R' And BillingCycle = '" & CID_ & "'")
If Not RsMst_BillingCycle.EOF Then
       mDepositBillingCycleID = IIf(IsNull(RsMst_BillingCycle!BillingCycleID), 0, RsMst_BillingCycle!BillingCycleID)
       mDepositBillingRatio = IIf(IsNull(RsMst_BillingCycle!Ratio), 0, RsMst_BillingCycle!Ratio)
       mDepositBillingCycleNoofDays = IIf(IsNull(RsMst_BillingCycle!NoOfDays), 0, RsMst_BillingCycle!NoOfDays)
End If
End Function

Private Sub ResetCell(Grid As MSHFlexGrid, mC1 As Variant, mC2 As Variant, mR As Variant)
With Grid
    Select Case LCase(Grid.Name)
        Case LCase("MSHFlexGrid1")    ''-- Without Bags Size
            
            If Val(.TextMatrix(mR, 15)) = 1 And Val(.TextMatrix(mR, 15)) = 7 Then
                .TextMatrix(mR, mC1) = Format(CalculateRate(.TextMatrix(mR, 16), mBillingRatio, CmbBillingCycle), "#0.0000")
                mBillingOnePercentage = Format(Val(.TextMatrix(mR, mC1)) / 100, "#0.0000")
                .TextMatrix(mR, mC1) = Format(Val(.TextMatrix(mR, mC1)) + (mBillingOnePercentage * .TextMatrix(mR, 17)), "#0.0000")
                .TextMatrix(mR, mC2) = Format(CalculateRate(.TextMatrix(mR, 16), mDepositBillingRatio, CmbDepositBillingCycle), "#0.0000")
                mDepositOnePercentage = Format(Val(.TextMatrix(mR, mC2)) / 100, "#0.0000")
                .TextMatrix(mR, mC2) = Format(Val(.TextMatrix(mR, mC2)) + (mDepositOnePercentage * .TextMatrix(mR, 17)), "#0.0000")
            Else
                .TextMatrix(mR, mC1) = Format(0, "#0.0000")
                .TextMatrix(mR, mC2) = Format(0, "#0.0000")
                '.TextMatrix(mR, 13) = Format(0, "#0.0000")
                '.TextMatrix(mR, 14) = Format(0, "#0.0000")
            End If
            .Col = mC1: .row = mR
            .CellBackColor = vbWhite
            .Col = mC2: .row = mR
            .CellBackColor = vbWhite
            
        Case LCase("MSHFlexGrid4")    ''-- With Bags Size
            If Val(.TextMatrix(mR, 17)) = 1 And Val(.TextMatrix(mR, 17)) = 7 Then
                 .TextMatrix(mR, mC1) = Format(CalculateRate(.TextMatrix(mR, 15), mBillingRatio, CmbBillingCycle) * Val(.TextMatrix(mR, 8)) / 1000, "#0.0000")
                 mBillingOnePercentage = Format(Val(.TextMatrix(mR, mC1)) / 100, "#0.0000")
                 .TextMatrix(mR, mC1) = Format(Val(.TextMatrix(mR, mC1)) + (mBillingOnePercentage * .TextMatrix(mR, 19)), "#0.0000")
                 .TextMatrix(mR, mC2) = Format(CalculateRate(.TextMatrix(mR, 16), mDepositBillingRatio, CmbDepositBillingCycle) * Val(.TextMatrix(mR, 8)) / 1000, "#0.0000")
                 mDepositOnePercentage = Format(Val(.TextMatrix(mR, mC2)) / 100, "#0.0000")
                 .TextMatrix(mR, mC2) = Format(Val(.TextMatrix(mR, mC2)) + (mDepositOnePercentage * .TextMatrix(mR, 19)), "#0.0000")
            Else
                .TextMatrix(mR, mC1) = Format(0, "#0.0000")
                .TextMatrix(mR, mC2) = Format(0, "#0.0000")
                '.TextMatrix(mR, 15) = Format(0, "#0.0000")
                '.TextMatrix(mR, 16) = Format(0, "#0.0000")
            End If
            
            .Col = mC1: .row = mR
            .CellBackColor = vbWhite
            .Col = mC2: .row = mR
            .CellBackColor = vbWhite
            
'            .TextMatrix(mR, mC1) = Format(Val(.TextMatrix(mR, 8)), "#0.00")
'            .TextMatrix(mR, mC1) = Format(CalculateRate(.TextMatrix(mR, 8), mBillingCycleIDRatio, CmbBillingCycle) * Val(.TextMatrix(mR, 7)) / 1000, "#0.00")
'            .Col = mC1: .row = mR
'            .CellBackColor = vbWhite
'
'            .TextMatrix(mR, mC2) = Format(Val(.TextMatrix(mR, 8)), "#0.00")
'            .TextMatrix(mR, mC2) = Format(CalculateRate(.TextMatrix(mR, 8), mDepositBillingRatio, CmbDepositBillingCycle) * Val(.TextMatrix(mR, 7)) / 1000, "#0.00")
'            .Col = mC2: .row = mR
'            .CellBackColor = vbWhite
            
'            If Val(.TextMatrix(mR, 17)) <> 1 And Val(.TextMatrix(mR, 17)) <> 7 Then
'                .TextMatrix(mR, 15) = Format(0, "#0.0000")
'                .TextMatrix(mR, 16) = Format(0, "#0.0000")
'            End If
            
    End Select
End With
End Sub

Public Sub UpperHierarchy(mEmployeeID_ As Variant)
tmp = "Select EmployeeID,EmployeeName + '~' + EmployeeNo As EmployeeName,DirectReportingTo From Mst_Employee Where EmployeeID=" & mEmployeeID_ & " And Flag='Y' And EmployeeNO <> '1'"
Call TmpTable
If TmpRs.RecordCount > 0 Then
   If Not IsNull(TmpRs!DirectReportingTo) Then
      CmbForwardedTo.AddItem TmpRs!EmployeeName
      Call UpperHierarchy(TmpRs!DirectReportingTo)
   End If
End If
End Sub

Private Sub GetUpperHierarchyEmployeeNoList(mEmployeeNo_ As Variant)
tmp = " Select E.EmployeeID,E.EmployeeName + '~' + E.EmployeeNo As EmployeeName,E.EmployeeNo,E.DirectReportingTo,EE.EmployeeNo As ReportingToEmployeeNo "
tmp = tmp & " From Mst_Employee E"
tmp = tmp & " Inner Join Mst_Employee EE On E.DirectReportingTo = EE.EmployeeID"
tmp = tmp & " Where E.EmployeeNo='" & mEmployeeNo_ & "' And EE.EmployeeNo <> '1'"
Call TmpTable
If TmpRs.RecordCount > 0 Then
   If Not IsNull(TmpRs!ReportingToEmployeeNo) Then
      If mUpperHierarchyEmployeeNoList = "" Then
         mUpperHierarchyEmployeeNoList = TmpRs!ReportingToEmployeeNo
      Else
         mUpperHierarchyEmployeeNoList = mUpperHierarchyEmployeeNoList & "," & TmpRs!ReportingToEmployeeNo
      End If
      Call GetUpperHierarchyEmployeeNoList(TmpRs!ReportingToEmployeeNo)
   End If
End If
End Sub

Private Function LoginEmployeeInHierarchy(mEmployeeNo_ As Variant) As Boolean
Dim mList As Variant
Dim mI, mR As Double
If mUpperHierarchyEmployeeNoList <> "" Then mList = Split(mUpperHierarchyEmployeeNoList, ",")
For mI = 0 To UBound(mList)
    If Trim(mEmployeeNo_) = Trim(mList(mI)) Then
        LoginEmployeeInHierarchy = True
        Exit Function
    End If
Next
LoginEmployeeInHierarchy = False
End Function

Private Function LoginEmployeeInAdditionalGroup(mEmployeeNo_ As Variant) As Boolean
Dim mList As Variant
Dim mI, mR As Double
If mDepositContractAdditionalGroupToBeApprovedList <> "" Then mList = Split(mDepositContractAdditionalGroupToBeApprovedList, ",")
For mI = 0 To UBound(mList)
    If Trim("'" & mEmployeeNo_ & "'") = Trim(mList(mI)) Then
        LoginEmployeeInAdditionalGroup = True
        Exit Function
    End If
Next
LoginEmployeeInAdditionalGroup = False
End Function


Public Sub UpperHierarchyEmployeeID(mEmployeeID_ As Variant)
tmp = "Select EmployeeID,EmployeeName + '~' + EmployeeNo As EmployeeName,DirectReportingTo From Mst_Employee Where EmployeeID=" & mEmployeeID_ & "  And EmployeeNO <> '1'"
Call TmpTable
If TmpRs.RecordCount > 0 Then
   If Not IsNull(TmpRs!DirectReportingTo) Then
      If Trim(mUpperHierarchyEmployeeIDList) = "" Then
        mUpperHierarchyEmployeeIDList = TmpRs!EmployeeID
      Else
        mUpperHierarchyEmployeeIDList = mUpperHierarchyEmployeeIDList & "," & TmpRs!EmployeeID
      End If
      Call UpperHierarchyEmployeeID(TmpRs!DirectReportingTo)
   End If
End If
End Sub

Public Function CheckforApproval(Grid_ As MSHFlexGrid) As String
'''-- If Proposed Rate is more than the Limit of RM Approval the direcly System Approved
'''-- If Proposed Rate is in the Limit of RM Approval the direcly Forwarded to RM for Approval
'''-- If Proposed Rate is less then system rate then Approved by CMG Team Only
Dim mR As Double
Dim mApprovalResult As String
Dim mBillingOnePercentage, mDepositOnePercentage As Variant
Dim mGetResult As Variant
mApprovalResult = ""
With Grid_
    Select Case LCase(Grid_.Name)
        Case LCase("MSHFlexGrid1")  ''-- Without Bag Size
             For mR = 1 To .Rows - 1
                If .TextMatrix(mR, 0) = strChecked Then
                   mBillingOnePercentage = Format(Val(.TextMatrix(mR, 7)) / 100, "#0.0000")
                   mDepositOnePercentage = Format(Val(.TextMatrix(mR, 11)) / 100, "#0.0000")
                
                   ''-- Lease / Own Godown
                   If .TextMatrix(mR, 15) = 1 Or .TextMatrix(mR, 15) = 7 Then
                   
                        ''-- Rate is less then System rate (CMG Approval)
                        If Val(.TextMatrix(mR, 8)) < Val(.TextMatrix(mR, 7)) Then
                           mApprovalResult = "C"
                           Exit For
                        ElseIf Val(.TextMatrix(mR, 12)) < Val(.TextMatrix(mR, 11)) Then
                           mApprovalResult = "C"
                           Exit For
                           
                        ''-- Rate is less then System rate (RM / Upper Approval)
                        ElseIf Val(.TextMatrix(mR, 8)) >= Val(.TextMatrix(mR, 7)) And Val(.TextMatrix(mR, 8)) <= Format(Val(.TextMatrix(mR, 7)) + (mBillingOnePercentage * Val(.TextMatrix(mR, 17))), "#0.0000") Then
                           mApprovalResult = "U"
                        ElseIf Val(.TextMatrix(mR, 12)) >= Val(.TextMatrix(mR, 11)) And Val(.TextMatrix(mR, 12)) < Format(Val(.TextMatrix(mR, 11)) + (mDepositOnePercentage * Val(.TextMatrix(mR, 17))), "#0.0000") Then
                           mApprovalResult = "U"
                           
                        ''-- Rate is more then System rate + Approval Limit (System Approval)
                        ElseIf Val(.TextMatrix(mR, 8)) > Format(Val(.TextMatrix(mR, 7)) - (mBillingOnePercentage * Val(.TextMatrix(mR, 17))), "#0.0000") Then
                           mApprovalResult = "S"
                        ElseIf Val(.TextMatrix(mR, 12)) > Format(Val(.TextMatrix(mR, 11)) - (mDepositOnePercentage * Val(.TextMatrix(mR, 17))), "#0.0000") Then
                           mApprovalResult = "S"
                        End If
                   Else ''--Un-Lease Godown
                      If .TextMatrix(mR, 15) = 2 Then ''-- Franchisee
                         mGetResult = GetFranchiseeRate(mLocationID, Val(.TextMatrix(mR, 9)), mCommodityID, CDbl(.TextMatrix(mR, 8)), CDbl(.TextMatrix(mR, 12)), CDbl(.TextMatrix(mR, 17)))
                         mApprovalResult = mGetResult
                     Else
                         mApprovalResult = "C"      ''-- for Any rate Approvel from CMG
                     End If
                   End If
                End If
             Next mR
        Case LCase("MSHFlexGrid4")  ''-- With Bag Size
             For mR = 1 To .Rows - 1
                If .TextMatrix(mR, 0) = strChecked Then
                    mBillingOnePercentage = Format(Val(.TextMatrix(mR, 9)) / 100, "#0.0000")
                    mDepositOnePercentage = Format(Val(.TextMatrix(mR, 13)) / 100, "#0.0000")
                    
                   ''-- Lease / Own Godown
                   If .TextMatrix(mR, 17) = 1 Or .TextMatrix(mR, 17) = 7 Then
                        ''-- Rate is less then System rate (CMG Approval)
                        If Val(.TextMatrix(mR, 10)) < Val(.TextMatrix(mR, 9)) Then
                           mApprovalResult = "C"
                           Exit For
                        ElseIf Val(.TextMatrix(mR, 14)) < Val(.TextMatrix(mR, 13)) Then
                           mApprovalResult = "C"
                           Exit For
                        ''-- Rate is less then System rate (RM / Upper Approval)
                        ElseIf Val(.TextMatrix(mR, 10)) >= Val(.TextMatrix(mR, 9)) And Val(.TextMatrix(mR, 10)) <= Format(Val(.TextMatrix(mR, 9)) + (mBillingOnePercentage * Val(.TextMatrix(mR, 19))), "#0.0000") Then
                           mApprovalResult = "U"
                        ElseIf Val(.TextMatrix(mR, 14)) >= Val(.TextMatrix(mR, 13)) And Val(.TextMatrix(mR, 14)) <= Format(Val(.TextMatrix(mR, 13)) + (mDepositOnePercentage * Val(.TextMatrix(mR, 19))), "#0.0000") Then
                           mApprovalResult = "U"
                        ''-- Rate is more then System rate + Approval Limit (System Approval)
                        ElseIf Val(.TextMatrix(mR, 10)) > Format(Val(.TextMatrix(mR, 9)) - (mBillingOnePercentage * Val(.TextMatrix(mR, 19))), "#0.0000") Then
                           mApprovalResult = "S"
                        ElseIf Val(.TextMatrix(mR, 14)) > Format(Val(.TextMatrix(mR, 13)) - (mDepositOnePercentage * Val(.TextMatrix(mR, 19))), "#0.0000") Then
                           mApprovalResult = "S"
                        End If
                    Else ''--Un-Lease Godown
                        If .TextMatrix(mR, 17) = 2 Then ''-- Franchisee
                            mGetResult = GetFranchiseeRate(mLocationID, CDbl(.TextMatrix(mR, 11)), mCommodityID, CDbl(.TextMatrix(mR, 10)), CDbl(.TextMatrix(mR, 14)), CDbl(.TextMatrix(mR, 19)))
                            mApprovalResult = mGetResult
                        Else
                            mApprovalResult = "C"      ''-- for Any rate Approvel from CMG
                        End If
                    End If
                End If
             Next mR
    End Select
End With
CheckforApproval = UCase(Trim(mApprovalResult))
End Function


Public Function GetOwnGodownPerSqtFeet(mLocation_ As Variant) As Double
tmp = " SELECT LNE.LocationID,LNE.NotionalExp,Sum(MG.Area)[TotalArea],LNE.NotionalExp/Sum(MG.Area)[SqrFeetRate] FROM Txn_LocationwiseNotionalExpense LNE"
tmp = tmp & " Inner Join Mst_CMP MC On LNE.LocationID = MC.LocationID"
tmp = tmp & " Inner Join Mst_Godown MG On MC.CMPID = MG.CMPID"
tmp = tmp & " Where LNE.LocationID = " & Val(mLocation_) & " Group By LNE.LocationID,LNE.NotionalExp"
Call Tmp1Table
If TmpRs1.EOF = False Then
    GetOwnGodownPerSqtFeet = Format(IIf(IsNull(TmpRs1!SqrFeetRate), "0.00", TmpRs1!SqrFeetRate), "#0.0000")
Else
    GetOwnGodownPerSqtFeet = 0
End If
End Function

Public Function GetFranchiseeRate(mLocation_ As Variant, mGodownID_ As Variant, mCommodityID_ As Variant, mmRate As Double, mmDRate As Double, mApprovalLimit_ As Double) As Variant

Dim mGFRCycleID, mGFRRate, mGFRUnit, mGFRRatio, mGFRCycle As Variant
Dim mGFRBRatePerMonth, mGFRDRatePerMonth, mIncomePart, mDepositIncomePart, mGFRBRatio, mGFRDRatio As Variant
Dim mReturn, mDReturn, mBKnownRentMinimumValue, mDKnownRentMinimumValue, mApprovalLimit As Variant
Dim mBOnePercentage, mDOnePercentage, mBRate, mDRate As Variant
Dim mGodownMinimumRent As Variant

''-- Get Godown wise minimum charges monthly otherwise from config
tmp = " Select RentMinimumValue From Mst_Godown Where CloseDate Is Null And GodownID = " & mGodownID_ & " "
Call Tmp1Table
If Not TmpRs1.EOF Then
    mGodownMinimumRent = IIf(IsNull(TmpRs1!RentMinimumValue), mKnownRentMinimumValue, RentMinimumValue)
Else
    mGodownMinimumRent = mKnownRentMinimumValue
End If


''-- Billing Cycle Wise Minimum Godown Rent Value
tmp = " Select * From Mst_BillingCycle Where Flag<>'R' And BillingCycleID = " & mBillingCycleID & " "
Call Tmp1Table
If Not TmpRs1.EOF Then
    mGFRBRatio = IIf(IsNull(TmpRs1!Ratio), 0, TmpRs1!Ratio)
    mBKnownRentMinimumValue = CalculateRate(mGodownMinimumRent, mGFRBRatio, CmbBillingCycle.Text)
Else
    mGFRBRatio = 0
    mBKnownRentMinimumValue = mGodownMinimumRent
End If
''-- Deposite Billing Cycle Wise Minimum Godown Rent Value
tmp = " Select * From Mst_BillingCycle Where Flag<>'R' And BillingCycleID = " & mDepositBillingCycleID & " "
Call Tmp1Table
If Not TmpRs1.EOF Then
    mGFRDRatio = IIf(IsNull(TmpRs1!Ratio), 0, TmpRs1!Ratio)
    mDKnownRentMinimumValue = CalculateRate(mGodownMinimumRent, mGFRDRatio, CmbDepositBillingCycle.Text)
Else
    mGFRDRatio = 0
    mDKnownRentMinimumValue = mGodownMinimumRent
End If

mBOnePercentage = mmRate / 100
mDOnePercentage = mmDRate / 100

tmp = " SELECT FFRCG.BillingCycleID,Max(FFRCG.RateOnQty) [Rate],FFRCG.BillingUnit,MBC.Ratio,MBC.BillingCycle"
tmp = tmp & " FROM Mst_FF_RateCard_Godown FFRCG"
tmp = tmp & " Inner Join Mst_FF_RateCard FFRC On FFRCG.FF_RateCardID = FFRC.FF_RateCardID"
tmp = tmp & " Inner Join Mst_CommodityGroup MCG On FFRC.CommodityGroupID = MCG.CommodityGroupID"
tmp = tmp & " Inner Join Mst_Commodity MC On MCG.CommodityGroupID = MC.CommodityGroupID"
tmp = tmp & " Inner Join Mst_BillingCycle MBC On FFRCG.BillingCycleID = MBC.BillingCycleID"
tmp = tmp & " Where FFRC.Flag = 'A' And FFRC.ExchangeTypeID = 1"
tmp = tmp & " And FFRC.LocationID = " & mLocation_ & " And FFRCG.GodownID = " & mGodownID_ & " And MC.CommodityID = " & mCommodityID_ & ""
tmp = tmp & " Group By FFRCG.GodownID,FFRCG.BillingCycleID,FFRCG.BillingUnit,MBC.Ratio,MBC.BillingCycle"
Call Tmp1Table

If TmpRs1.EOF = False Then
    
    mGFRCycleID = IIf(IsNull(TmpRs1!BillingCycleID), "", TmpRs1!BillingCycleID)
    mGFRRate = Format(IIf(IsNull(TmpRs1!Rate), 0#, TmpRs1!Rate), "#0.0000")
    mGFRUnit = IIf(IsNull(TmpRs1!BillingUnit), "", TmpRs1!BillingUnit)
    mGFRRatio = Format(IIf(IsNull(TmpRs1!Ratio), 0#, TmpRs1!Ratio), "#0.0000")
    mGFRCycle = IIf(IsNull(TmpRs1!BillingCycle), "", TmpRs1!BillingCycle)
    
    Select Case LCase(mGFRUnit)
        Case "p"
            mIncomePart = mmRate - ((mmRate * mGFRRate) / 100)
            mDepositIncomePart = mmDRate - ((mmDRate * mGFRRate) / 100)
            
            ''-- Rate is less then System rate (CMG Approval)
            If mIncomePart < mBKnownRentMinimumValue Then
                mReturn = "C"
            ''-- Rate is less then System rate (RM / Upper Approval)
            ElseIf (mIncomePart >= mBKnownRentMinimumValue) And (mIncomePart < mBKnownRentMinimumValue + (mBOnePercentage * mApprovalLimit_)) Then
                mReturn = "U"
            ''-- Rate is more then System rate + Approval Limit (System Approval)
            ElseIf mIncomePart > (mBKnownRentMinimumValue + (mBOnePercentage * mApprovalLimit_)) Then
               mReturn = "S"
            End If
               
            ''-- Rate is less then System rate (CMG Approval)
            If mDepositIncomePart < mDKnownRentMinimumValue Then
                mDReturn = "C"
            ''-- Rate is less then System rate (RM / Upper Approval)
            ElseIf (mDepositIncomePart >= mDKnownRentMinimumValue) And (mDepositIncomePart < mDKnownRentMinimumValue + (mDOnePercentage * mApprovalLimit_)) Then
                mDReturn = "U"
            ''-- Rate is more then System rate + Approval Limit (System Approval)
            ElseIf mDepositIncomePart > (mBKnownRentMinimumValue + (mDOnePercentage * mApprovalLimit_)) Then
               mDReturn = "S"
            End If
            
        Case "b"
            mReturn = "C"
            mDReturn = "C"
        Case "f"
            mReturn = "C"
            mDReturn = "C"
        Case "q"
            If mGFRCycleID = 4 Then
                mGFRBRatePerMonth = mGFRRate
            ElseIf mGFRCycleID = 3 Then
                mGFRBRatePerMonth = mGFRRate * 2
            ElseIf mGFRCycleID = 2 Then
                mGFRBRatePerMonth = mGFRRate * 4
            ElseIf mGFRCycleID = 1 Then
                mGFRBRatePerMonth = mGFRRate * 30
            End If
            mGFRDRatePerMonth = mGFRBRatePerMonth
            
            mBRate = CalculateRate(mGFRBRatePerMonth, mGFRBRatio, CmbBillingCycle.Text)
            mDRate = CalculateRate(mGFRDRatePerMonth, mGFRDRatio, CmbDepositBillingCycle.Text)
        
            ''-- Rate is less then System rate (CMG Approval)
            If mmRate < mBRate Then
                mReturn = "C"
            ''-- Rate is less then System rate (RM / Upper Approval)
            ElseIf (mmRate >= mBRate) And (mmRate < mBRate + (mBOnePercentage * mApprovalLimit_)) Then
                mReturn = "U"
            ''-- Rate is more then System rate + Approval Limit (System Approval)
            ElseIf mmRate > (mBRate + (mBOnePercentage * mApprovalLimit_)) Then
               mReturn = "S"
            End If
            
            ''-- Rate is less then System rate (CMG Approval)
            If mmDRate < mDRate Then
                mDReturn = "C"
            ''-- Rate is less then System rate (RM / Upper Approval)
            ElseIf (mmDRate >= mDRate) And (mmDRate < mDRate + (mBOnePercentage * mApprovalLimit_)) Then
                mDReturn = "U"
            ''-- Rate is more then System rate + Approval Limit (System Approval)
            ElseIf mmDRate > (mDRate + (mBOnePercentage * mApprovalLimit_)) Then
               mDReturn = "S"
            End If
            
            'mReturn = "C"
            'mDReturn = "C"
    End Select
Else    ''-- Rate for Found
    mReturn = "C"
    mDReturn = "C"
End If
GetFranchiseeRate = mReturn & "" & mDReturn

End Function


Sub CheckRateByColor(Index_ As Integer)
Dim mBillingOnePercentage, mDepositOnePercentage As Variant
Dim Retval, mGetResult, mBResult, mDResult As Variant
Dim mColor As ColorConstants
Retval = ""
With MSHFlexGrid1
    If .TextMatrix(Index_, 1) <> "" And .TextMatrix(Index_, 0) = strChecked Then
        mBillingOnePercentage = Format(Val(.TextMatrix(Index_, 7)) / 100, "#0.0000")
        mDepositOnePercentage = Format(Val(.TextMatrix(Index_, 11)) / 100, "#0.0000")
        ''-- Lease / Own Godown
        If .TextMatrix(Index_, 15) = 1 Or .TextMatrix(Index_, 15) = 7 Then
           
           If Val(.TextMatrix(Index_, 8)) < Val(.TextMatrix(Index_, 7)) Then
           ''-- Rate is less then System rate (CMG Approval)
              mColor = vbRed
              .Col = 8: .row = Index_: .CellBackColor = mColor
           ElseIf Val(.TextMatrix(Index_, 8)) >= Val(.TextMatrix(Index_, 7)) And Val(.TextMatrix(Index_, 8)) <= Format(Val(.TextMatrix(Index_, 7)) + (mBillingOnePercentage * Val(.TextMatrix(Index_, 17))), "#0.0000") Then
           ''-- Rate is less then System rate (RM / Upper Approval)
              mColor = vbYellow
              .Col = 8: .row = Index_: .CellBackColor = mColor
           ElseIf Val(.TextMatrix(Index_, 8)) > Format(Val(.TextMatrix(Index_, 7)) - (mBillingOnePercentage * Val(.TextMatrix(Index_, 17))), "#0.0000") Then
           ''-- Rate is more then System rate + Approval Limit (System Approval)
              mColor = vbGreen
              .Col = 8: .row = Index_: .CellBackColor = mColor
           Else
              mColor = vbWhite
              .Col = 8: .row = Index_: .CellBackColor = mColor
           End If
           
           If Val(.TextMatrix(Index_, 12)) < Val(.TextMatrix(Index_, 11)) Then
           ''-- Rate is less then System rate (CMG Approval)
              mColor = vbRed
              .Col = 12: .row = Index_: .CellBackColor = mColor
           ElseIf Val(.TextMatrix(Index_, 12)) >= Val(.TextMatrix(Index_, 11)) And Val(.TextMatrix(Index_, 12)) <= Format(Val(.TextMatrix(Index_, 11)) + (mDepositOnePercentage * Val(.TextMatrix(Index_, 17))), "#0.0000") Then
           ''-- Rate is less then System rate (RM / Upper Approval)
              mColor = vbYellow
              .Col = 12: .row = Index_: .CellBackColor = mColor
           ElseIf Val(.TextMatrix(Index_, 12)) > Format(Val(.TextMatrix(Index_, 11)) - (mDepositOnePercentage * Val(.TextMatrix(Index_, 17))), "#0.0000") Then
           ''-- Rate is more then System rate + Approval Limit (System Approval)
              mColor = vbGreen
              .Col = 12: .row = Index_: .CellBackColor = mColor
           Else
              mColor = vbWhite
              .Col = 12: .row = Index_: .CellBackColor = mColor
           End If
           
        Else ''--Un-Lease Godown
           If .TextMatrix(Index_, 15) = 2 Then ''-- Franchisee
               mGetResult = GetFranchiseeRate(mLocationID, Val(.TextMatrix(Index_, 9)), mCommodityID, Val(.TextMatrix(Index_, 8)), Val(.TextMatrix(Index_, 12)), CDbl(.TextMatrix(Index_, 17)))
               mBResult = Left(mGetResult, 1)
               mDResult = Right(mGetResult, 1)
               
               If Val(.TextMatrix(Index_, 8)) >= 0 Then
                    If LCase(mBResult) = LCase("C") Then
                       mColor = vbRed
                       .Col = 8: .row = Index_: .CellBackColor = mColor
                    ElseIf LCase(mBResult) = LCase("U") Then
                       mColor = vbYellow
                       .Col = 8: .row = Index_: .CellBackColor = mColor
                    ElseIf LCase(mBResult) = LCase("S") Then
                       mColor = vbGreen
                       .Col = 8: .row = Index_: .CellBackColor = mColor
                    Else
                        mColor = vbWhite
                        .Col = 8: .row = Index_: .CellBackColor = mColor
                    End If
               Else
                    mColor = vbWhite
                    .Col = 8: .row = Index_: .CellBackColor = mColor
               End If
               
               If Val(.TextMatrix(Index_, 12)) >= 0 Then
                    If LCase(mDResult) = LCase("C") Then
                       mColor = vbRed
                       .Col = 12: .row = Index_: .CellBackColor = mColor
                    ElseIf LCase(mDResult) = LCase("U") Then
                       mColor = vbYellow
                       .Col = 12: .row = Index_: .CellBackColor = mColor
                    ElseIf LCase(mDResult) = LCase("S") Then
                       mColor = vbGreen
                       .Col = 12: .row = Index_: .CellBackColor = mColor
                    Else
                        mColor = vbWhite
                        .Col = 12: .row = Index_: .CellBackColor = mColor
                    End If
               Else
                    mColor = vbWhite
                    .Col = 12: .row = Index_: .CellBackColor = mColor
               End If

           Else
               mGetResult = "CC"      ''-- for Any rate Approvel from CMG
               mBResult = Left(mGetResult, 1)
               mDResult = Right(mGetResult, 1)
               
               If Val(.TextMatrix(Index_, 8)) >= 0 Then
                    If LCase(mBResult) = LCase("C") Then
                       mColor = vbRed
                       .Col = 8: .row = Index_: .CellBackColor = mColor
                    ElseIf LCase(mBResult) = LCase("U") Then
                       mColor = vbYellow
                       .Col = 8: .row = Index_: .CellBackColor = mColor
                    ElseIf LCase(mBResult) = LCase("S") Then
                       mColor = vbGreen
                       .Col = 8: .row = Index_: .CellBackColor = mColor
                    Else
                        mColor = vbWhite
                        .Col = 8: .row = Index_: .CellBackColor = mColor
                    End If
               Else
                    mColor = vbWhite
                    .Col = 8: .row = Index_: .CellBackColor = mColor
               End If
               
               If Val(.TextMatrix(Index_, 12)) >= 0 Then
                    If LCase(mDResult) = LCase("C") Then
                       mColor = vbRed
                       .Col = 12: .row = Index_: .CellBackColor = mColor
                    ElseIf LCase(mDResult) = LCase("U") Then
                       mColor = vbYellow
                       .Col = 12: .row = Index_: .CellBackColor = mColor
                    ElseIf LCase(mDResult) = LCase("S") Then
                       mColor = vbGreen
                       .Col = 12: .row = Index_: .CellBackColor = mColor
                    Else
                        mColor = vbWhite
                        .Col = 12: .row = Index_: .CellBackColor = mColor
                    End If
               Else
                    mColor = vbWhite
                    .Col = 12: .row = Index_: .CellBackColor = mColor
               End If
               
           End If

        End If
    End If
End With
End Sub

Sub CheckBagRateByColor(Index_ As Integer)
Dim mBillingOnePercentage, mDepositOnePercentage As Variant
Dim Retval, mGetResult, mBResult, mDResult As Variant
Dim mColor As ColorConstants
Retval = ""
With MSHFlexGrid4
    If .TextMatrix(Index_, 1) <> "" And .TextMatrix(Index_, 0) = strChecked Then
        mBillingOnePercentage = Format(Val(.TextMatrix(Index_, 9)) / 100, "#0.0000")
        mDepositOnePercentage = Format(Val(.TextMatrix(Index_, 13)) / 100, "#0.0000")
        
        ''-- Lease / Own Godown
        If .TextMatrix(Index_, 17) = 1 Or .TextMatrix(Index_, 17) = 7 Then
        
           If Val(.TextMatrix(Index_, 10)) < Val(.TextMatrix(Index_, 9)) Then
           ''-- Rate is less then System rate (CMG Approval)
              mColor = vbRed
              .Col = 10: .row = Index_: .CellBackColor = mColor
           ElseIf Val(.TextMatrix(Index_, 10)) >= Val(.TextMatrix(Index_, 9)) And Val(.TextMatrix(Index_, 10)) <= Format(Val(.TextMatrix(Index_, 9)) + (mBillingOnePercentage * Val(.TextMatrix(Index_, 19))), "#0.0000") Then
           ''-- Rate is less then System rate (RM / Upper Approval)
              mColor = vbYellow
              .Col = 10: .row = Index_: .CellBackColor = mColor
           ElseIf Val(.TextMatrix(Index_, 10)) > Format(Val(.TextMatrix(Index_, 9)) - (mBillingOnePercentage * Val(.TextMatrix(Index_, 19))), "#0.0000") Then
           ''-- Rate is more then System rate + Approval Limit (System Approval)
              mColor = vbGreen
              .Col = 10: .row = Index_: .CellBackColor = mColor
           Else
              mColor = vbWhite
              .Col = 10: .row = Index_: .CellBackColor = mColor
           End If
           
           If Val(.TextMatrix(Index_, 14)) < Val(.TextMatrix(Index_, 13)) Then
           ''-- Rate is less then System rate (CMG Approval)
              mColor = vbRed
              .Col = 14: .row = Index_: .CellBackColor = mColor
           ElseIf Val(.TextMatrix(Index_, 14)) >= Val(.TextMatrix(Index_, 13)) And Val(.TextMatrix(Index_, 14)) <= Format(Val(.TextMatrix(Index_, 13)) + (mDepositOnePercentage * Val(.TextMatrix(Index_, 19))), "#0.0000") Then
           ''-- Rate is less then System rate (RM / Upper Approval)
              mColor = vbYellow
              .Col = 14: .row = Index_: .CellBackColor = mColor
           ElseIf Val(.TextMatrix(Index_, 14)) > Format(Val(.TextMatrix(Index_, 13)) - (mDepositOnePercentage * Val(.TextMatrix(Index_, 19))), "#0.0000") Then
           ''-- Rate is more then System rate + Approval Limit (System Approval)
              mColor = vbGreen
              .Col = 14: .row = Index_: .CellBackColor = mColor
           Else
              mColor = vbWhite
              .Col = 14: .row = Index_: .CellBackColor = mColor
           End If
           
        Else ''--Un-Lease Godown
           If .TextMatrix(Index_, 17) = 2 Then ''-- Franchisee
               mGetResult = GetFranchiseeRate(mLocationID, Val(.TextMatrix(Index_, 11)), mCommodityID, Val(.TextMatrix(Index_, 10)), Val(.TextMatrix(Index_, 14)), CDbl(.TextMatrix(Index_, 19)))
               mBResult = Left(mGetResult, 1)
               mDResult = Right(mGetResult, 1)
               
               If Val(.TextMatrix(Index_, 10)) >= 0 Then
                    If LCase(mBResult) = LCase("C") Then
                       mColor = vbRed
                       .Col = 10: .row = Index_: .CellBackColor = mColor
                    ElseIf LCase(mBResult) = LCase("U") Then
                       mColor = vbYellow
                       .Col = 10: .row = Index_: .CellBackColor = mColor
                    ElseIf LCase(mBResult) = LCase("S") Then
                       mColor = vbGreen
                       .Col = 10: .row = Index_: .CellBackColor = mColor
                    Else
                        mColor = vbWhite
                        .Col = 10: .row = Index_: .CellBackColor = mColor
                    End If
               Else
                    mColor = vbWhite
                    .Col = 10: .row = Index_: .CellBackColor = mColor
               End If
               
               If Val(.TextMatrix(Index_, 14)) >= 0 Then
                    If LCase(mDResult) = LCase("C") Then
                       mColor = vbRed
                       .Col = 14: .row = Index_: .CellBackColor = mColor
                    ElseIf LCase(mDResult) = LCase("U") Then
                       mColor = vbYellow
                       .Col = 14: .row = Index_: .CellBackColor = mColor
                    ElseIf LCase(mDResult) = LCase("S") Then
                       mColor = vbGreen
                       .Col = 14: .row = Index_: .CellBackColor = mColor
                    Else
                        mColor = vbWhite
                        .Col = 14: .row = Index_: .CellBackColor = mColor
                    End If
               Else
                    mColor = vbWhite
                    .Col = 14: .row = Index_: .CellBackColor = mColor
               End If

           Else
               mGetResult = "CC"      ''-- for Any rate Approvel from CMG
               mBResult = Left(mGetResult, 1)
               mDResult = Right(mGetResult, 1)
               
               If Val(.TextMatrix(Index_, 10)) >= 0 Then
                    If LCase(mBResult) = LCase("C") Then
                       mColor = vbRed
                       .Col = 10: .row = Index_: .CellBackColor = mColor
                    ElseIf LCase(mBResult) = LCase("U") Then
                       mColor = vbYellow
                       .Col = 10: .row = Index_: .CellBackColor = mColor
                    ElseIf LCase(mBResult) = LCase("S") Then
                       mColor = vbGreen
                       .Col = 10: .row = Index_: .CellBackColor = mColor
                    Else
                        mColor = vbWhite
                        .Col = 10: .row = Index_: .CellBackColor = mColor
                    End If
               Else
                    mColor = vbWhite
                    .Col = 10: .row = Index_: .CellBackColor = mColor
               End If
               
               If Val(.TextMatrix(Index_, 14)) >= 0 Then
                    If LCase(mDResult) = LCase("C") Then
                       mColor = vbRed
                       .Col = 14: .row = Index_: .CellBackColor = mColor
                    ElseIf LCase(mDResult) = LCase("U") Then
                       mColor = vbYellow
                       .Col = 14: .row = Index_: .CellBackColor = mColor
                    ElseIf LCase(mDResult) = LCase("S") Then
                       mColor = vbGreen
                       .Col = 14: .row = Index_: .CellBackColor = mColor
                    Else
                        mColor = vbWhite
                        .Col = 14: .row = Index_: .CellBackColor = mColor
                    End If
               Else
                    mColor = vbWhite
                    .Col = 14: .row = Index_: .CellBackColor = mColor
               End If
               
           End If

        End If
    End If
End With
End Sub


Sub ShowPopDetails(Grid As MSHFlexGrid, Index As Integer)
TxtGodownRent = ""
TxtGodownArea = ""
TxtGodownRentPSFeet = ""
TxtGodownRentPSFeet_Min = mCurrentBaseRentFactor
TxtConversionFactor = ""
TxtBagEquivalent = ""
TxtDivisionFactor = ""
TxtFumigationCharges = ""
TxtBillingRateMonthlyMTRate = ""
TxtBillingRateMonthlyMTRate_Min = ""
TxtGodownRMMargin = ""
TxtGodownLastRentPaid = ""
TxtGodownLastPNL = ""
TxtRentPaidDate = ""
TxtMonthYear = ""
TxtFranchiseRatio = ""
TxtUtilizationLevel = ""

Label46.Caption = ""
Label47.Caption = ""

If LCase(Grid.Name) = LCase("MSHFlexGrid1") Then
With Grid
    
    Label46.Caption = "CMP Name : " & .TextMatrix(Index, 1)
    Label47.Caption = "Godown No : " & .TextMatrix(Index, 2)
    
    TxtGodownRent = .TextMatrix(Index, 5)
    TxtGodownArea = .TextMatrix(Index, 6)
    TxtGodownRentPSFeet = Format(Val(.TextMatrix(Index, 5)) / Val(.TextMatrix(Index, 6)), "#0.00")
    TxtConversionFactor = mCurrentConversionFactor
    TxtBagEquivalent = mBagEquivalent
    TxtDivisionFactor = mDivisionFactor
    TxtFumigationCharges = IIf(LCase(mFumigationFlag) = "y", mFumigationFactorY, mFumigationFactorN)
    
    TxtBillingRateMonthlyMTRate = .TextMatrix(Index, 16)
    
    TxtBillingRateMonthlyMTRate_Min = (((Val(TxtGodownRentPSFeet) * mBagEquivalent * TxtConversionFactor)) / (mDivisionFactor / 100)) + Val(TxtFumigationCharges)
    TxtBillingRateMonthlyMTRate_Min = Format(TxtBillingRateMonthlyMTRate_Min, "########.####")
    
    TxtBillingRateMonthlyMTRate_Min.FontBold = False
    TxtBillingRateMonthlyMTRate.FontBold = True
        
    
    
    TxtGodownRMMargin = .TextMatrix(Index, 17)
    
    If Val(.TextMatrix(Index, 15)) = 1 Then ''-- Lease Godown
        tmp = " Select TGRD.ClaimID,TGR.ExpenseFrom,TGR.ExpenseTo,TGRD.Amount,TGRD.Arrears from Txn_GodownRent TGR "
        tmp = tmp & " Inner Join Txn_GodownRentDetail TGRD On TGR.ClaimID = TGRD.ClaimID And TGRD.GodownID = " & Val(.TextMatrix(Index, 9)) & ""
        tmp = tmp & " And TGR.ExpenseFrom = (Select Max(GR.ExpenseFrom) From Txn_GodownRent GR"
        tmp = tmp & " Inner Join Txn_GodownRentDetail GRD On GR.ClaimID = GRD.ClaimID"
        tmp = tmp & " Where GRD.GodownID = " & Val(.TextMatrix(Index, 9)) & ")"
        Call Tmp2Table
        If TmpRs2.RecordCount > 0 Then
            TxtGodownLastRentPaid = IIf(IsNull(TmpRs2!Amount), 0, TmpRs2!Amount) + IIf(IsNull(TmpRs2!Arrears), 0, TmpRs2!Arrears)
            TxtRentPaidDate = Format(TmpRs2!ExpenseFrom, "MMMM-YYYY")
        End If
    Else
        TxtGodownLastRentPaid = ""
        TxtRentPaidDate = ""
    End If
    
    tmp = " Select GodownID,((Isnull(NNNRIncomeAdhoc,0)+Isnull(NNNRIncome,0)+Isnull(ReservationIncomeAdhoc,0)+Isnull(ReservationIncome,0)+Isnull(SevenDaysIncome,0)+Isnull(DPIncomeAdHoc,0)+Isnull(FinalDPIncome,0)+Isnull(RejectedIncome,0)+Isnull(SpotIncomeAdHoc,0)+Isnull(OtherIncome,0)+Isnull(SCMIncome,0)+Isnull(CMUnbilledIncome,0)+Isnull(CMPendingIncome,0)+Isnull(CMApprovedIncome,0)+Isnull(LeftOverStock,0)+Isnull(PledgeIncome,0))-Isnull(PartiallyReversedIncome,0))[TotalIncome],"
    tmp = tmp & " (Isnull(LeaseRentPaid,0)+Isnull(LeaseRentFAG,0)+Isnull(LeaseRentDRC,0)+Isnull(EmpClaimFAG,0)+Isnull(EmpClaimAccrual,0)+Isnull(VendorExpenses,0)+Isnull(VendorExpensesAccrued,0)+Isnull(SecurityFeesPaid,0)+Isnull(SecurityFeesAccrual,0)+Isnull(FumigationFeesVendor,0)+Isnull(FumigationFeesPaid,0)+Isnull(FumigationFeesAccrual,0)+Isnull(FranchiseeFeesPaid,0)+Isnull(FranchiseeFeesCMG,0)+Isnull(FranchiseeFeesFAG,0)+Isnull(Salary,0)+Isnull(NCDEXCharges,0)+Isnull(LeaseRentAccrued,0))[TotalExpense],"
    tmp = tmp & " Isnull(CoExpense,0)CoExpense,Isnull(InsurancePremium,0)InsurancePremium,Isnull(InterestCharges,0)InterestCharges,MonthYear"
    tmp = tmp & " From Txn_MonGodownWiseProfitability "
    tmp = tmp & " Where GodownID = " & Val(.TextMatrix(Index, 9)) & " "
    tmp = tmp & " And MonthYear = (SELECT MAX(MonthYear) FROM Txn_MonGodownWiseProfitability Where GodownID = " & Val(.TextMatrix(Index, 9)) & ") "
    Call Tmp2Table
    If TmpRs2.RecordCount > 0 Then
        TxtGodownLastPNL = (TmpRs2!TotalIncome - TmpRs2!TotalExpense) - (TmpRs2!CoExpense - TmpRs2!InsurancePremium - TmpRs2!InterestCharges)
        TxtMonthYear = Format(TmpRs2!MonthYear, "MMMM-YYYY")
    End If
    
    tmp = " SELECT SUM(BalanceWeight)BalanceWeight FROM Mst_GSL WHERE Flag NOT IN ('Z','E') And BalanceWeight > 0 And GodownID = " & Val(.TextMatrix(Index, 9)) & " "
    Call Tmp2Table
    If TmpRs2.RecordCount > 0 Then
        TxtUtilizationLevel = Format(IIf(IsNull(TmpRs2!BalanceWeight) = True, 0, TmpRs2!BalanceWeight) / (TxtGodownArea.Text / 6) * 100, "0.00")
    Else
        TxtUtilizationLevel = ""
    End If
    
End With

ElseIf LCase(Grid.Name) = LCase("MSHFlexGrid4") Then

With Grid
    Label46.Caption = "CMP Name : " & .TextMatrix(Index, 1)
    Label47.Caption = "Godown No : " & .TextMatrix(Index, 2)
    
    TxtGodownRent = .TextMatrix(Index, 5)
    TxtGodownArea = .TextMatrix(Index, 6)
    TxtGodownRentPSFeet = Format(Val(.TextMatrix(Index, 5)) / Val(.TextMatrix(Index, 6)), "#0.00")
    TxtConversionFactor = mCurrentConversionFactor
    TxtBagEquivalent = mBagEquivalent
    TxtDivisionFactor = mDivisionFactor
    TxtFumigationCharges = IIf(LCase(mFumigationFlag) = "y", mFumigationFactorY, mFumigationFactorN)
    TxtBillingRateMonthlyMTRate = .TextMatrix(Index, 18)
    TxtGodownRMMargin = .TextMatrix(Index, 19)

    If Val(.TextMatrix(Index, 17)) = 1 Then ''-- Lease Godown
        tmp = " Select TGRD.ClaimID,TGR.ExpenseFrom,TGR.ExpenseTo,TGRD.Amount,TGRD.Arrears from Txn_GodownRent TGR "
        tmp = tmp & " Inner Join Txn_GodownRentDetail TGRD On TGR.ClaimID = TGRD.ClaimID And TGRD.GodownID = " & Val(.TextMatrix(Index, 11)) & ""
        tmp = tmp & " And TGR.ExpenseFrom = (Select Max(GR.ExpenseFrom) From Txn_GodownRent GR"
        tmp = tmp & " Inner Join Txn_GodownRentDetail GRD On GR.ClaimID = GRD.ClaimID"
        tmp = tmp & " Where GRD.GodownID = " & Val(.TextMatrix(Index, 11)) & ")"
        Call Tmp2Table
        If TmpRs2.RecordCount > 0 Then
            TxtGodownLastRentPaid = IIf(IsNull(TmpRs2!Amount), 0, TmpRs2!Amount) + IIf(IsNull(TmpRs2!Arrears), 0, TmpRs2!Arrears)
            TxtRentPaidDate = Format(TmpRs2!ExpenseFrom, "MMMM-YYYY")
        End If
    Else
        TxtGodownLastRentPaid = ""
        TxtRentPaidDate = ""
    End If
    
    tmp = " Select GodownID,((Isnull(NNNRIncomeAdhoc,0)+Isnull(NNNRIncome,0)+Isnull(ReservationIncomeAdhoc,0)+Isnull(ReservationIncome,0)+Isnull(SevenDaysIncome,0)+Isnull(DPIncomeAdHoc,0)+Isnull(FinalDPIncome,0)+Isnull(RejectedIncome,0)+Isnull(SpotIncomeAdHoc,0)+Isnull(OtherIncome,0)+Isnull(SCMIncome,0)+Isnull(CMUnbilledIncome,0)+Isnull(CMPendingIncome,0)+Isnull(CMApprovedIncome,0)+Isnull(LeftOverStock,0)+Isnull(PledgeIncome,0))-Isnull(PartiallyReversedIncome,0))[TotalIncome],"
    tmp = tmp & " (Isnull(LeaseRentPaid,0)+Isnull(LeaseRentFAG,0)+Isnull(LeaseRentDRC,0)+Isnull(EmpClaimFAG,0)+Isnull(EmpClaimAccrual,0)+Isnull(VendorExpenses,0)+Isnull(VendorExpensesAccrued,0)+Isnull(SecurityFeesPaid,0)+Isnull(SecurityFeesAccrual,0)+Isnull(FumigationFeesVendor,0)+Isnull(FumigationFeesPaid,0)+Isnull(FumigationFeesAccrual,0)+Isnull(FranchiseeFeesPaid,0)+Isnull(FranchiseeFeesCMG,0)+Isnull(FranchiseeFeesFAG,0)+Isnull(Salary,0)+Isnull(NCDEXCharges,0)+Isnull(LeaseRentAccrued,0))[TotalExpense],"
    tmp = tmp & " Isnull(CoExpense,0)CoExpense,Isnull(InsurancePremium,0)InsurancePremium,Isnull(InterestCharges,0)InterestCharges,MonthYear"
    tmp = tmp & " From Txn_MonGodownWiseProfitability "
    tmp = tmp & " Where GodownID = " & Val(.TextMatrix(Index, 11)) & " "
    tmp = tmp & " And MonthYear = (SELECT MAX(MonthYear) FROM Txn_MonGodownWiseProfitability Where GodownID = " & Val(.TextMatrix(Index, 11)) & ") "
    Call Tmp2Table
    If TmpRs2.RecordCount > 0 Then
        TxtGodownLastPNL = (TmpRs2!TotalIncome - TmpRs2!TotalExpense) - (TmpRs2!CoExpense - TmpRs2!InsurancePremium - TmpRs2!InterestCharges)
        TxtMonthYear = Format(TmpRs2!MonthYear, "MMMM-YYYY")
    End If
    
    tmp = " SELECT SUM(BalanceWeight)BalanceWeight FROM Mst_GSL WHERE Flag NOT IN ('Z','E') And BalanceWeight > 0 And GodownID = " & Val(.TextMatrix(Index, 11)) & " "
    Call Tmp2Table
    If TmpRs2.RecordCount > 0 Then
        TxtUtilizationLevel = Format(IIf(IsNull(TmpRs2!BalanceWeight) = True, 0, TmpRs2!BalanceWeight) / (TxtGodownArea.Text / 6) * 100, "0.00")
    Else
        TxtUtilizationLevel = ""
    End If

End With

End If
End Sub
