VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmTxn_AdjTransaction 
   Caption         =   "Adjustment Transaction"
   ClientHeight    =   8790
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14340
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8790
   ScaleWidth      =   14340
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   1245
      Left            =   13440
      TabIndex        =   55
      Top             =   120
      Visible         =   0   'False
      Width           =   2475
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   350
         Left            =   80
         TabIndex        =   56
         Top             =   150
         Width           =   1000
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1455
         Left            =   70
         TabIndex        =   57
         Top             =   600
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   2566
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1200
      Left            =   120
      TabIndex        =   40
      Top             =   7560
      Width           =   13080
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New"
         Height          =   400
         Left            =   105
         TabIndex        =   51
         Top             =   195
         Width           =   2220
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   400
         Left            =   2325
         TabIndex        =   50
         Top             =   195
         Width           =   2340
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   400
         Left            =   4665
         TabIndex        =   49
         Top             =   195
         Width           =   2220
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   400
         Left            =   6885
         TabIndex        =   48
         Top             =   195
         Width           =   2220
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "Search"
         Height          =   400
         Left            =   9105
         TabIndex        =   47
         Top             =   195
         Width           =   2100
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   400
         Left            =   105
         TabIndex        =   46
         Top             =   600
         Width           =   2220
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   400
         Left            =   2325
         TabIndex        =   45
         Top             =   600
         Width           =   2340
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   400
         Left            =   4665
         TabIndex        =   44
         Top             =   600
         Width           =   2220
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   400
         Left            =   6885
         TabIndex        =   43
         Top             =   600
         Width           =   2220
      End
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   400
         Left            =   9105
         TabIndex        =   42
         Top             =   600
         Width           =   2100
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Print"
         Height          =   810
         Left            =   11205
         TabIndex        =   41
         Top             =   195
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4575
      Left            =   120
      TabIndex        =   27
      Top             =   3000
      Width           =   13080
      Begin VB.CommandButton CmdGSL 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3600
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   120
         Width           =   375
      End
      Begin VB.CheckBox ChkSpillage 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Check1"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   10440
         TabIndex        =   52
         Top             =   240
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox TxtBalQty 
         Alignment       =   1  'Right Justify
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
         Left            =   5535
         Locked          =   -1  'True
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   540
         Width           =   1470
      End
      Begin VB.TextBox TxtBalBags 
         Alignment       =   1  'Right Justify
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
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   540
         Width           =   1560
      End
      Begin VB.CommandButton CmdDeleteList 
         Caption         =   "Delete From List"
         Height          =   345
         Left            =   11535
         TabIndex        =   28
         Top             =   555
         Width           =   1335
      End
      Begin VB.TextBox TxtNoofBags 
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
         Left            =   7020
         TabIndex        =   10
         Top             =   540
         Width           =   1380
      End
      Begin VB.TextBox TxtQuantity 
         Alignment       =   1  'Right Justify
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
         Height          =   375
         Left            =   8415
         TabIndex        =   11
         Top             =   540
         Width           =   1575
      End
      Begin VB.CommandButton CmdAdd2List 
         Caption         =   "Add to List"
         Height          =   345
         Left            =   10200
         TabIndex        =   12
         Top             =   555
         Width           =   1335
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
         ItemData        =   "FrmTxn_AdjTransaction.frx":0000
         Left            =   120
         List            =   "FrmTxn_AdjTransaction.frx":0002
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   540
         Width           =   1335
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
         ItemData        =   "FrmTxn_AdjTransaction.frx":0004
         Left            =   1440
         List            =   "FrmTxn_AdjTransaction.frx":0006
         Sorted          =   -1  'True
         TabIndex        =   8
         Top             =   540
         Width           =   1215
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
         ItemData        =   "FrmTxn_AdjTransaction.frx":0008
         Left            =   2640
         List            =   "FrmTxn_AdjTransaction.frx":000A
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   540
         Width           =   1215
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
         Height          =   3495
         Left            =   120
         TabIndex        =   29
         Top             =   960
         Width           =   12735
         _ExtentX        =   22463
         _ExtentY        =   6165
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label6 
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
         Height          =   255
         Left            =   5618
         TabIndex        =   39
         Top             =   240
         Width           =   1305
      End
      Begin VB.Label Label5 
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
         Height          =   255
         Left            =   4088
         TabIndex        =   37
         Top             =   240
         Width           =   1305
      End
      Begin VB.Label Label20 
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
         Left            =   120
         TabIndex        =   34
         Top             =   240
         Width           =   1335
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
         Left            =   1680
         TabIndex        =   33
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label22 
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
         Left            =   2880
         TabIndex        =   32
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label23 
         Caption         =   "No Of Bags"
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
         Left            =   7103
         TabIndex        =   31
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label24 
         Caption         =   "Quantity"
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
         Left            =   8775
         TabIndex        =   30
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2865
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   13065
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
         Left            =   1815
         Locked          =   -1  'True
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   2310
         Width           =   3735
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
         Left            =   8190
         Locked          =   -1  'True
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   2310
         Width           =   3930
      End
      Begin VB.TextBox TxtCurrentDate 
         Alignment       =   1  'Right Justify
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
         Left            =   10800
         Locked          =   -1  'True
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   1080
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.CommandButton CmdAgent 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   10200
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   1470
         Width           =   375
      End
      Begin VB.TextBox TxtTxnID 
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
         Left            =   1815
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   240
         Width           =   3675
      End
      Begin VB.ComboBox CmbCMPID 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1815
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   630
         Width           =   8295
      End
      Begin VB.ComboBox CmbAgent 
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
         Left            =   1815
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   1470
         Width           =   8295
      End
      Begin VB.TextBox TxtPreFix 
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
         Left            =   8160
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1035
         Width           =   1920
      End
      Begin VB.ComboBox CmbCommodityID 
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
         Left            =   1815
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   1050
         Width           =   3735
      End
      Begin VB.ComboBox CmbType 
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
         ItemData        =   "FrmTxn_AdjTransaction.frx":000C
         Left            =   1815
         List            =   "FrmTxn_AdjTransaction.frx":001F
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   1875
         Width           =   3735
      End
      Begin VB.TextBox TxtAmount 
         Alignment       =   1  'Right Justify
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
         Left            =   8175
         TabIndex        =   6
         Top             =   1875
         Width           =   1920
      End
      Begin MSComCtl2.DTPicker TxtTxnDate 
         Height          =   375
         Left            =   8160
         TabIndex        =   1
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         Format          =   58916865
         CurrentDate     =   39884
      End
      Begin VB.TextBox TxtCMPID 
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
         Left            =   1815
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   15
         Top             =   630
         Visible         =   0   'False
         Width           =   8265
      End
      Begin VB.TextBox TxtCommodityID 
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
         Left            =   1815
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1050
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.TextBox TxtAgent 
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
         Left            =   1815
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1470
         Visible         =   0   'False
         Width           =   8295
      End
      Begin VB.TextBox TxtType 
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
         Left            =   1815
         ScrollBars      =   2  'Vertical
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1875
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.Label Label10 
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
         TabIndex        =   62
         Top             =   2377
         Width           =   1695
      End
      Begin VB.Label Label9 
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
         Left            =   6870
         TabIndex        =   61
         Top             =   2370
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "BOD"
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
         Left            =   10200
         TabIndex        =   54
         Top             =   1125
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label8 
         Caption         =   "Txn Date"
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
         Left            =   6840
         TabIndex        =   26
         Top             =   270
         Width           =   1800
      End
      Begin VB.Label Label4 
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
         TabIndex        =   25
         Top             =   270
         Width           =   1800
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
         Left            =   120
         TabIndex        =   24
         Top             =   660
         Width           =   1800
      End
      Begin VB.Label Label26 
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
         Height          =   255
         Left            =   6840
         TabIndex        =   23
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label1 
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
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   1110
         Width           =   1215
      End
      Begin VB.Label Label29 
         Caption         =   "Agent"
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
         TabIndex        =   21
         Top             =   1530
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Type"
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
         TabIndex        =   20
         Top             =   1935
         Width           =   855
      End
      Begin VB.Label Label3 
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
         Left            =   6840
         TabIndex        =   19
         Top             =   1928
         Width           =   855
      End
   End
End
Attribute VB_Name = "FrmTxn_AdjTransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg, mCMPID, mGodownID, mAgentID, mCommodityID, mType, mGSLID As Variant

Private Sub AddCmd_Click()
Edit_Flg = "A"

Call GButtonLock(Me, False)

Call Grid_Head
TxtCMPID.Locked = False
CmbCMPID.Visible = True
TxtCMPID.Visible = False
TxtCommodityID.Locked = False
CmbCommodityID.Visible = True
TxtCommodityID.Visible = False

TxtAgent.Locked = False
CmbAgent.Visible = True
CmdAgent.Visible = True
TxtAgent.Visible = False
CmbAgent.Text = ""
TxtType.Visible = False
CmbType.Visible = True
CmbType.Text = ""
TxtAmount = ""
TxtAmount.Locked = False
TxtTxnID = ""
TxtTxnDate = Date
TxtTxnDate.Enabled = True
TxtPreFix = ""
TxtCurrentDate = ""
TxtCMPID = ""
TxtCommodityID = ""
TxtAgent = ""
CmbCMPID.Text = ""
CmbCommodityID.Text = ""

TxtCreated = ""
TxtUpdated = ""

'TxtTxnDate.SetFocus
End Sub
Private Sub CmbCMPID_GotFocus()
tmp = CmbCMPID.Text

Call TableMst_CMP(" where CloseDate is null ")
CmbCMPID.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_CMP.RecordCount
    CmbCMPID.AddItem RsMst_CMP!CMPName
    RsMst_CMP.MoveNext
Next
CmbCMPID.Text = tmp

End Sub
'----
Private Sub CmbCMPID_LostFocus()
If CmbCMPID.Text = "" Then
   Exit Sub
End If

RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPName = '" & CmbCMPID.Text & "'"
If RsMst_CMP.EOF Then
   MsgBox "Invalid selection "
   CmbCMPID.SetFocus
   Exit Sub
End If
If mCMPID <> RsMst_CMP!CMPID Then
   Call Grid_Head
End If

mCMPID = RsMst_CMP!CMPID
TxtPreFix = GetPreFix(RsMst_CMP!CMPID, "P")
TxtCurrentDate = IIf(IsNull(RsMst_CMP!CurrentDate), "", RsMst_CMP!CurrentDate) + 1
TxtTxnDate.Value = IIf(IsNull(RsMst_CMP!CurrentDate), Date, RsMst_CMP!CurrentDate) + 1
End Sub
'----
Private Sub CmbCommodityID_GotFocus()

tmp = CmbCommodityID.Text
Call TableMst_Commodity
CmbCommodityID.Clear
If RsMst_Commodity.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_Commodity.RecordCount
    CmbCommodityID.AddItem RsMst_Commodity!Commodity
    RsMst_Commodity.MoveNext
Next
CmbCommodityID.Text = tmp


End Sub
'----
Private Sub CmbCommodityID_LostFocus()
If CmbCommodityID.Text = "" Then
   Exit Sub
End If

RsMst_Commodity.MoveFirst
RsMst_Commodity.Find "Commodity = '" & CmbCommodityID.Text & "'"
If RsMst_Commodity.EOF Then
   MsgBox "Invalid selection "
   CmbCommodityID.SetFocus
   Exit Sub
End If

If mCommodityID <> RsMst_Commodity!CommodityID Then
   Call Grid_Head
End If
mCommodityID = RsMst_Commodity!CommodityID

'SSTab1.Tab = 0
End Sub
Private Sub CmbGodownNo_GotFocus()
If CmbCMPID.Text = "" Then
   MsgBox "Select CMP !!!! "
   CmbCMPID.SetFocus
   Exit Sub
End If
If CmbCommodityID.Text <> "" Then
    tmp = CmbGodownNo.Text
    Call TableMst_Godown(" Where CloseDate is null And CMPID = " & mCMPID)
    CmbGodownNo.Clear
    If RsMst_Godown.RecordCount = 0 Then
       MsgBox "No Record found !!!! "
       Exit Sub
    End If
    For i = 1 To RsMst_Godown.RecordCount
        CmbGodownNo.AddItem RsMst_Godown!GodownNo
        RsMst_Godown.MoveNext
    Next
    CmbGodownNo.Text = tmp
Else
   MsgBox "Please Select Commodity!!!"
   CmbCommodityID.SetFocus
   Exit Sub
End If
End Sub
'----
Private Sub CmbGodownNo_LostFocus()
If CmbGodownNo.Text = "" Then
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
   mCmbStack = ""
   mCmbLot = ""
End If
mGodownID = RsMst_Godown!GodownID
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
For i = 1 To RsMst_GSL.RecordCount
    CmbStack.AddItem RsMst_GSL!StackNo
    RsMst_GSL.MoveNext
Next

CmbStack.Text = tmp
End Sub
'----
Private Sub CmbStack_LostFocus()
If CmbStack.Text = "" Then
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
   mCmbLot = ""
End If
mCmbStack = CmbStack.Text
End Sub

Private Sub CmbLot_GotFocus()
If CmbStack.Text = "" Then
  MsgBox "Please Select Stack No !!!"
  CmbStack.SetFocus
  Exit Sub
End If
If CmbCommodityID.Text = "" Then
  MsgBox "Please Select Commodity !!!"
  CmbCommodityID.SetFocus
  Exit Sub
End If
tmp = CmbLot.Text
Call TableMst_GSL(" Where  " & Gen_ExpireGSLFlag & " And  CMPID = " & mCMPID & " And GodownID = " & mGodownID & " And StackNo = '" & CmbStack.Text & "'")
CmbLot.Clear
If RsMst_GSL.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If

For i = 1 To RsMst_GSL.RecordCount
    CmbLot.AddItem RsMst_GSL!LotNo
    RsMst_GSL.MoveNext
Next i
CmbLot.Text = tmp
End Sub
'----
Private Sub CmbLot_LostFocus()
If CmbLot.Text = "" Then
   Exit Sub
End If
RsMst_GSL.MoveFirst
RsMst_GSL.Find "LotNo = '" & CmbLot.Text & "'"
If RsMst_GSL.EOF Then
   MsgBox "Invalid selection "
   CmbLot.SetFocus
   Exit Sub
End If

If mCommodityID <> RsMst_GSL!CommodityID Then
   MsgBox "GSL Contain Stock of other Commodity!!! "
   CmbLot.SetFocus
   Exit Sub
End If

mGSLID = RsMst_GSL!GSLID
TxtBalBags = IIf(IsNull(RsMst_GSL!DepositeBags), 0, RsMst_GSL!DepositeBags) - IIf(IsNull(RsMst_GSL!CDTFBags), 0, RsMst_GSL!CDTFBags)
TxtBalQty = IIf(IsNull(RsMst_GSL!DepositeWeight), 0, RsMst_GSL!DepositeWeight) - IIf(IsNull(RsMst_GSL!CDTFWeight), 0, RsMst_GSL!CDTFWeight)
ChkSpillage.Value = IIf(IsNull(RsMst_GSL!SpillageFlag), 0, RsMst_GSL!SpillageFlag)
End Sub

Private Sub CmbType_LostFocus()
If CmbType.Text = "" Then
   mType = Null
   TxtAmount = 0
   TxtAmount.Enabled = False
   Exit Sub
End If

If CmbType.Text = "Sale" Or CmbType.Text = "Theft" Or CmbType.Text = "Fire" Or CmbType.Text = "Adjustment Received from client" Or CmbType.Text = "Adjustment Given to client" Then

Else
   MsgBox "Invalid selection!!!"
   CmbType.SetFocus
   Exit Sub
End If

If CmbType.Text = "Sale" Then
   TxtAmount.Enabled = True
   TxtAmount.SetFocus
Else
   TxtAmount.Enabled = Flase
   TxtAmount = ""
End If

If CmbType.Text = "Adjustment Received from client" Then
   mType = "AR"
ElseIf CmbType.Text = "Adjustment Given to client" Then
   mType = "AG"
Else
   mType = Left(CmbType.Text, 2)
End If
Call Grid_Head


End Sub

Private Sub CmdAdd2List_Click()
Dim r As Variant

If (CmbGodownNo.Text) = "" Then
   MsgBox "Blank Godown No Not Allowed !!! "
   CmbGodownNo.SetFocus
   Exit Sub
End If
If (CmbStack.Text) = "" Then
   MsgBox "Blank Stack  Not Allowed !!! "
   CmbStack.SetFocus
   Exit Sub
End If
If (CmbLot.Text) = "" Then
   MsgBox "Blank Lot Not Allowed !!! "
   CmbLot.SetFocus
   Exit Sub
End If
If CheckDuplicateGSL() = True Then
   MsgBox "Duplicate G.S.L.  Not Allowed !!! "
   CmbGodownNo.SetFocus
   Exit Sub
End If
If Val(TxtNoofBags.Text) <= 0 And Val(TxtQuantity) <= 0 Then
   MsgBox "Invalid No of Bags !!! "
   TxtNoofBags.SetFocus
   Exit Sub
End If
If Val(TxtNoofBags.Text) <= 0 And Val(TxtQuantity) <= 0 Then
   MsgBox "Invalid Weight !!! "
   TxtQuantity.SetFocus
   Exit Sub
End If
If mType = "AR" Or mType = "AD" Then
   If CmbLot.Text <> "" Then
      If ChkSpillage.Value = 0 Then
         MsgBox "Non Spillage Lot not allowed!!!"
         CmbLot.SetFocus
         Exit Sub
      End If
   End If
End If
If CmdAdd2List.Caption = "Update List" Then
   r = MSHFlexGrid2.RowSel
   MSHFlexGrid2.TextMatrix(0, 6) = Val(MSHFlexGrid2.TextMatrix(0, 6)) - Val(MSHFlexGrid2.TextMatrix(r, 6))
   MSHFlexGrid2.TextMatrix(0, 7) = Val(MSHFlexGrid2.TextMatrix(0, 7)) - Val(MSHFlexGrid2.TextMatrix(r, 7))
Else
   r = MSHFlexGrid2.Rows - 1
   MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
End If

MSHFlexGrid2.TextMatrix(r, 1) = CmbGodownNo.Text
MSHFlexGrid2.TextMatrix(r, 2) = CmbStack.Text
MSHFlexGrid2.TextMatrix(r, 3) = CmbLot.Text
MSHFlexGrid2.TextMatrix(r, 4) = TxtBalBags
MSHFlexGrid2.TextMatrix(r, 5) = TxtBalQty
MSHFlexGrid2.TextMatrix(r, 6) = Val(TxtNoofBags)
MSHFlexGrid2.TextMatrix(r, 7) = Val(TxtQuantity)
MSHFlexGrid2.TextMatrix(r, 8) = mGSLID
MSHFlexGrid2.TextMatrix(r, 9) = ChkSpillage.Value
 
MSHFlexGrid2.TextMatrix(0, 6) = Val(MSHFlexGrid2.TextMatrix(0, 6)) + Val(MSHFlexGrid2.TextMatrix(r, 6))
MSHFlexGrid2.TextMatrix(0, 7) = Val(MSHFlexGrid2.TextMatrix(0, 7)) + Val(MSHFlexGrid2.TextMatrix(r, 7))

Call ClearGSL
CmbGodownNo.SetFocus

End Sub
Private Sub ClearGSL()
CmbGodownNo.Text = ""
CmbStack.Text = ""
CmbLot.Text = ""
TxtBalBags = ""
TxtBalQty = ""
TxtNoofBags = ""
TxtQuantity = ""

CmdAdd2List.Caption = "Add to List"
End Sub

Private Sub CmdAgent_Click()
FrmMst_Agent.Show
End Sub
Private Sub CmdDeleteList_Click()
i = MSHFlexGrid2.RowSel
If i > 1 Then
   If MSHFlexGrid2.TextMatrix(i, 1) <> "" Then
      MSHFlexGrid2.TextMatrix(0, 6) = Val(MSHFlexGrid2.TextMatrix(0, 6)) - Val(MSHFlexGrid2.TextMatrix(i, 6))
      MSHFlexGrid2.TextMatrix(0, 7) = Val(MSHFlexGrid2.TextMatrix(0, 7)) - Val(MSHFlexGrid2.TextMatrix(i, 7))
      For RO = i To MSHFlexGrid2.Rows - 2
          For c = 0 To MSHFlexGrid2.Cols - 1
               MSHFlexGrid2.TextMatrix(RO, c) = MSHFlexGrid2.TextMatrix(RO + 1, c)
          Next
      Next
      MSHFlexGrid2.Rows = MSHFlexGrid2.Rows - 1
   End If
End If
Call ClearGSL
End Sub



Private Sub CmdExcel_Click()
'Call Gen_Create_Xls("Rpt_KR Pending.xls")

Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "AdjTransaction.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "AdjTransaction.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For i = 0 To MSHFlexGrid1.Rows - 1
    mRow = i + 1 ' + 6
    For c = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(i, c), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(i, c)) Then
              oWS.Cells(mRow, c + 1) = Format(MSHFlexGrid1.TextMatrix(i, c), "mm/dd/yyyy")
           Else
              oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(i, c)
           End If
        Else
           oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(i, c)
        End If
    Next
Next
oWB.Save
Set oWB = Nothing
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub

Private Sub CmdGSL_Click()
FrmMst_GSL.Show
End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub CmbAgent_GotFocus()
tmp = CmbAgent.Text
Call TableMst_Agent
CmbAgent.Clear
If RsMst_Agent.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_Agent.RecordCount
    CmbAgent.AddItem RsMst_Agent!AgentName
    RsMst_Agent.MoveNext
Next
CmbAgent.Text = tmp
End Sub
'----
Private Sub CmbAgent_LostFocus()
If CmbAgent.Text = "" Then
   Exit Sub
End If
RsMst_Agent.MoveFirst
RsMst_Agent.Find "agentName = '" & CmbAgent.Text & "'"
If RsMst_Agent.EOF Then
   MsgBox "Invalid selection "
   CmbAgent.SetFocus
   Exit Sub
End If
mAgentID = RsMst_Agent!AgentID
End Sub
Private Sub Grid_Head()
MSHFlexGrid2.Clear
MSHFlexGrid2.Rows = 3
MSHFlexGrid2.FixedRows = 2

MSHFlexGrid2.Cols = 10
MSHFlexGrid2.TextMatrix(1, 0) = "Row ID"
MSHFlexGrid2.TextMatrix(1, 1) = "Godown No"
MSHFlexGrid2.TextMatrix(1, 2) = "Stack No"
MSHFlexGrid2.TextMatrix(1, 3) = "Lot No"
MSHFlexGrid2.TextMatrix(1, 4) = "Balance Bags"
MSHFlexGrid2.TextMatrix(1, 5) = "Balance Qty"
MSHFlexGrid2.TextMatrix(1, 6) = "No of Bags"
MSHFlexGrid2.TextMatrix(1, 7) = "Quantity"
MSHFlexGrid2.TextMatrix(1, 8) = "GSLID"
MSHFlexGrid2.TextMatrix(1, 9) = ""

MSHFlexGrid2.ColWidth(0) = 0
MSHFlexGrid2.ColWidth(1) = 1800
MSHFlexGrid2.ColWidth(2) = 1800
MSHFlexGrid2.ColWidth(3) = 1800
MSHFlexGrid2.ColWidth(4) = 1800
MSHFlexGrid2.ColWidth(5) = 1800
MSHFlexGrid2.ColWidth(6) = 1800 '2000
MSHFlexGrid2.ColWidth(7) = 1800
MSHFlexGrid2.ColWidth(8) = 0
MSHFlexGrid2.ColWidth(9) = 0
MSHFlexGrid2.RowHeight(1) = 750

End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Call GButtonLockAD(Me)
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If
Call Grid_Head
Call TableTxn_AdjTransaction("")
If RsTxn_AdjTransaction.RecordCount = 0 Then
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
Call InData

End Sub



Private Sub MSHFlexGrid2_Click()
Call ClearGSL
r = MSHFlexGrid2.RowSel
If r <= 1 Then Exit Sub
If MSHFlexGrid2.TextMatrix(r, 1) = "" Then
   Call ClearGSL
   Exit Sub
End If

CmdAdd2List.Caption = "Update List"


CmbGodownNo.Text = MSHFlexGrid2.TextMatrix(r, 1)
CmbStack.Text = MSHFlexGrid2.TextMatrix(r, 2)
CmbLot.Text = MSHFlexGrid2.TextMatrix(r, 3)
TxtBalBags = MSHFlexGrid2.TextMatrix(r, 4)
TxtBalQty = MSHFlexGrid2.TextMatrix(r, 5)
TxtNoofBags = MSHFlexGrid2.TextMatrix(r, 6)
TxtQuantity = MSHFlexGrid2.TextMatrix(r, 7)
mGSLID = MSHFlexGrid2.TextMatrix(r, 8)

Call CmbGodownNo_GotFocus
Call CmbGodownNo_LostFocus
CmdDeleteList.Enabled = True

If MSHFlexGrid2.TextMatrix(r, 0) <> "" Then
   CmbGodownNo.Enabled = False
   CmbStack.Enabled = False
   CmbLot.Enabled = False
Else
   CmbGodownNo.Enabled = True
   CmbStack.Enabled = True
   CmbLot.Enabled = True
End If
End Sub

Private Sub SaveCmd_Click()

If TxtTxnDate.Value > Date Then
   MsgBox "Future Date not allowed!!!"
   TxtTxnDate.SetFocus
   Exit Sub
End If

If TxtTxnDate.Value > CDate(TxtCurrentDate) Then
   'MsgBox "Transaction Date must be less than BOD date of CMP!!!"
   MsgBox "Transaction date cannot be later than BOD+1 date i.e." & TxtCurrentDate & "!!!"
   TxtTxnDate.SetFocus
   Exit Sub
End If


If CmbCMPID.Text = "" Then
   MsgBox "Select CMP "
   CmbCMPID.SetFocus
   Exit Sub
End If
If CmbCommodityID.Text = "" Then
   MsgBox "Select Commodity "
   CmbCommodityID.SetFocus
   Exit Sub
End If
If CmbAgent.Text = "" Then
   MsgBox "Select Agent "
   CmbAgent.SetFocus
   Exit Sub
End If

If CmbType.Text = "" Then
   MsgBox "Select Type"
   CmbType.SetFocus
   Exit Sub
End If


If CmbType.Text = "Sale" Then
   If Val(TxtAmount) = 0 Then
      MsgBox "Amount can't be zero!!!"
      TxtAmount.SetFocus
      Exit Sub
   End If
End If

If MSHFlexGrid2.Rows <= 3 Then
   MsgBox "Missing GSL Detail !!!! "
   Exit Sub
End If

If Edit_Flg = "A" Then
   RsTxn_AdjTransaction.AddNew
   RsTxn_AdjTransaction!TxnID = GetTxnID
   RsTxn_AdjTransaction!G_UID = GetGUID
   TxtTxnID.Text = RsTxn_AdjTransaction!TxnID
Else
    RsTxn_AdjTransaction.MoveFirst
    RsTxn_AdjTransaction.Find " TxnID= " & TxtTxnID.Text
    'Call DeleteChildTableDetail
End If

RsTxn_AdjTransaction!TxnID = TxtTxnID.Text
RsTxn_AdjTransaction!TxnDate = TxtTxnDate.Value
RsTxn_AdjTransaction!CMPID = mCMPID
RsTxn_AdjTransaction!CommodityID = mCommodityID
RsTxn_AdjTransaction!TxnType = mType
RsTxn_AdjTransaction!AgentID = mAgentID
RsTxn_AdjTransaction!Amount = Val(TxtAmount)
RsTxn_AdjTransaction!SM_Prefix = TxtPreFix

If IsNull(RsTxn_AdjTransaction!CreatedBy) = True Or RsTxn_AdjTransaction!CreatedBy = "" Then
   RsTxn_AdjTransaction!CreatedBy = Gen_UserLoginID
   RsTxn_AdjTransaction!CreatedDate = Now
Else
   RsTxn_AdjTransaction!UpdatedBy = Gen_UserLoginID
   RsTxn_AdjTransaction!UpdatedDate = Now
End If


RsTxn_AdjTransaction.Update

SaveGSL
SaveCmd.Enabled = False
RsTxn_GSLAdjTransaction.Close

Call InData


End Sub
Private Sub SaveGSL()
Call TableTxn_GSLAdjTransaction(" Where TxnID = " & TxtTxnID & "")
Call TableMst_GSL("where SM_Prefix = '" & TxtPreFix & "'")
For i = 2 To MSHFlexGrid2.Rows - 1
    If MSHFlexGrid2.TextMatrix(i, 1) <> "" Then
       RsTxn_GSLAdjTransaction.AddNew
       RsTxn_GSLAdjTransaction!TxnGSLAdjID = GetGSLAdjID
       RsTxn_GSLAdjTransaction!TxnGSLAdjDat = TxtTxnDate.Value
       RsTxn_GSLAdjTransaction!G_UID = GetGUID
       RsMst_GSL.MoveFirst
       RsMst_GSL.Find "GSLID = " & Val(MSHFlexGrid2.TextMatrix(i, 8))
       RsTxn_GSLAdjTransaction!TxnID = TxtTxnID
       RsTxn_GSLAdjTransaction!NoofBags = Val(MSHFlexGrid2.TextMatrix(i, 6))
       RsTxn_GSLAdjTransaction!Quantity = Val(MSHFlexGrid2.TextMatrix(i, 7))
       RsTxn_GSLAdjTransaction!GSLID = MSHFlexGrid2.TextMatrix(i, 8)
       RsTxn_GSLAdjTransaction!SpillageFlag = MSHFlexGrid2.TextMatrix(i, 9)
'       RsTxn_GSLAdjTransaction!SM_Prefix = RsMst_GSL!SM_Prefix
       RsTxn_GSLAdjTransaction.Update
       RsTxn_GSLAdjTransaction.Requery
       If Not RsMst_GSL.EOF Then
          If mType = "AR" Then
             If IsNull(RsMst_GSL!CommodityID) Then
                RsMst_GSL!CommodityID = mCommodityID
             End If
             RsMst_GSL!DepositeBags = IIf(IsNull(RsMst_GSL!DepositeBags), 0, RsMst_GSL!DepositeBags) + Val(MSHFlexGrid2.TextMatrix(i, 6))
             RsMst_GSL!DepositeWeight = IIf(IsNull(RsMst_GSL!DepositeWeight), 0, RsMst_GSL!DepositeWeight) + Val(MSHFlexGrid2.TextMatrix(i, 7))
             RsMst_GSL!BalanceBags = IIf(IsNull(RsMst_GSL!BalanceBags), 0, RsMst_GSL!BalanceBags) + Val(MSHFlexGrid2.TextMatrix(i, 6))
             RsMst_GSL!BalanceWeight = IIf(IsNull(RsMst_GSL!BalanceWeight), 0, RsMst_GSL!BalanceWeight) + Val(MSHFlexGrid2.TextMatrix(i, 7))
          Else
             RsMst_GSL!CDTFBags = IIf(IsNull(RsMst_GSL!CDTFBags), 0, RsMst_GSL!CDTFBags) + Val(MSHFlexGrid2.TextMatrix(i, 6))
             RsMst_GSL!CDTFWeight = IIf(IsNull(RsMst_GSL!CDTFWeight), 0, RsMst_GSL!CDTFWeight) + Val(MSHFlexGrid2.TextMatrix(i, 7))
             RsMst_GSL!BalanceBags = IIf(IsNull(RsMst_GSL!BalanceBags), 0, RsMst_GSL!BalanceBags) - Val(MSHFlexGrid2.TextMatrix(i, 6))
             RsMst_GSL!BalanceWeight = IIf(IsNull(RsMst_GSL!BalanceWeight), 0, RsMst_GSL!BalanceWeight) - Val(MSHFlexGrid2.TextMatrix(i, 7))
          End If
          If Val(RsMst_GSL!BalanceBags) = 0 Then
             RsMst_GSL!Flag = "E"
          End If
          RsMst_GSL.Update
       End If
    End If
Next

End Sub

Private Function GetTxnID() As Double
Dim retVal As Double
tmp = "Select max(TxnID) from Txn_AdjTransaction" ' Where TxnType = 'D' And SM_Prefix = '" & TxtPreFix & "'"
Call TmpTable
retVal = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
retVal = retVal + 1
GetTxnID = retVal
End Function
Function GetGSLAdjID() As Double
Dim retVal As Double
retVal = 0
tmp = "select max(TxnGSLAdjID) from Txn_GSLAdjTransaction" ' where SM_Prefix = '" & TxtPreFix & "'"
Call TmpTable
If TmpRs.RecordCount > 0 Then
   retVal = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
End If
retVal = retVal + 1
TmpRs.Close
GetGSLAdjID = retVal

End Function

Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   If RsTxn_AdjTransaction.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   Call InData
   Exit Sub
End If

Frame0.Height = Me.Height - 350
Frame0.Left = Frame3.Left - 100
Frame0.Top = Frame3.Top - 100
Frame0.Width = Me.Width - 150
Frame0.Visible = True
MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 1000
RsTxn_AdjTransaction.MoveFirst
Call Grid_Head1
MSHFlexGrid1.Rows = RsTxn_AdjTransaction.RecordCount + 1
'MSHFlexGrid1.ColWidth(0) = 1000
For i = 1 To RsTxn_AdjTransaction.RecordCount
    MSHFlexGrid1.TextMatrix(i, 0) = RsTxn_AdjTransaction!TxnID
    MSHFlexGrid1.TextMatrix(i, 1) = RsTxn_AdjTransaction!TxnDate
    MSHFlexGrid1.TextMatrix(i, 2) = GetCMPName(RsTxn_AdjTransaction!CMPID)
    MSHFlexGrid1.TextMatrix(i, 3) = RsTxn_AdjTransaction!SM_Prefix
    MSHFlexGrid1.TextMatrix(i, 4) = GetCommodityName(RsTxn_AdjTransaction!CommodityID)
    MSHFlexGrid1.TextMatrix(i, 5) = GetAgentName(RsTxn_AdjTransaction!AgentID) 'IIf(IsNull(RsTxn_AdjTransaction!AgentID), "", RsTxn_AdjTransaction!Agent)
    MSHFlexGrid1.TextMatrix(i, 6) = IIf(IsNull(RsTxn_AdjTransaction!TxnType), "", RsTxn_AdjTransaction!TxnType)
    MSHFlexGrid1.TextMatrix(i, 7) = IIf(IsNull(RsTxn_AdjTransaction!Amount), "", RsTxn_AdjTransaction!Amount)
    MSHFlexGrid1.TextMatrix(i, 8) = IIf(IsNull(RsTxn_AdjTransaction!Flag), "", RsTxn_AdjTransaction!Flag)
    MSHFlexGrid1.TextMatrix(i, 9) = IIf(IsNull(RsTxn_AdjTransaction!G_UID), "", RsTxn_AdjTransaction!G_UID)
    RsTxn_AdjTransaction.MoveNext
Next

End Sub
Private Sub Grid_Head1()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 3


MSHFlexGrid1.Cols = 10
MSHFlexGrid1.TextMatrix(0, 0) = "TxnID"
MSHFlexGrid1.TextMatrix(0, 1) = "TxnDate"
MSHFlexGrid1.TextMatrix(0, 2) = "CMP"
MSHFlexGrid1.TextMatrix(0, 3) = "SM_Prefix"
MSHFlexGrid1.TextMatrix(0, 4) = "Commodity"
MSHFlexGrid1.TextMatrix(0, 5) = "Agent"
MSHFlexGrid1.TextMatrix(0, 6) = "Txn Type"
MSHFlexGrid1.TextMatrix(0, 7) = "Amount"
MSHFlexGrid1.TextMatrix(0, 8) = "Flag"
MSHFlexGrid1.TextMatrix(0, 9) = "G_UID"

MSHFlexGrid1.ColWidth(0) = 2000
MSHFlexGrid1.ColWidth(1) = 2000
MSHFlexGrid1.ColWidth(2) = 2000
MSHFlexGrid1.ColWidth(3) = 2000
MSHFlexGrid1.ColWidth(4) = 2000
MSHFlexGrid1.ColWidth(5) = 2000
MSHFlexGrid1.ColWidth(6) = 2000 '2000
MSHFlexGrid1.ColWidth(7) = 2000
MSHFlexGrid1.ColWidth(8) = 2000
MSHFlexGrid1.ColWidth(9) = 2000





End Sub
Private Sub TxtAmount_LostFocus()
If TxtAmount = "" Then Exit Sub
If IsNumeric(TxtAmount) = False Or Val(TxtAmount) < 0 Then
   MsgBox "Invalid numeric format!!!"
   TxtAmount.SetFocus
   Exit Sub
End If

End Sub

Private Sub TxtNoOfBags_LostFocus()
If TxtNoofBags = "" Then Exit Sub
If IsNumeric(TxtNoofBags) = False Or Val(TxtNoofBags) < 0 Then
   MsgBox "Invalid numeric format!!!"
   TxtNoofBags.SetFocus
   Exit Sub
End If


If mType <> "AR" Then
    If Val(TxtNoofBags) > Val(TxtBalBags) Then
       MsgBox "No of bags must be less than Balance bags!!!"
       TxtNoofBags.SetFocus
       Exit Sub
    End If
End If

TxtNoofBags = Round(Val(TxtNoofBags))

End Sub

Private Sub TxtQuantity_LostFocus()
If TxtQuantity = "" Then Exit Sub
If IsNumeric(TxtQuantity) = False Or Val(TxtQuantity) < 0 Then
   MsgBox "Invalid numeric format!!!"
   TxtQuantity.SetFocus
   Exit Sub
End If

If mType <> "AR" Then
   If Val(TxtQuantity) > Val(TxtBalQty) Then
      MsgBox "Quantity must be less than Balance Quantity!!!"
      TxtQuantity.SetFocus
      Exit Sub
   End If
End If
End Sub
Public Sub InData()

Edit_Flg = ""
Call Grid_Head
Call ClearGSL
TxtTxnDate.Enabled = False
TxtCMPID.Locked = True
CmbCMPID.Visible = False
TxtCMPID.Visible = True
TxtCommodityID.Locked = True
CmbCommodityID.Visible = False
TxtCommodityID.Visible = True
TxtAgent.Locked = True
CmbAgent.Visible = False
CmdAgent.Visible = False
TxtAgent.Visible = True
CmbType.Visible = False
TxtType.Visible = True
TxtAmount.Locked = True
TxtTxnID = IIf(IsNull(RsTxn_AdjTransaction!TxnID), "", RsTxn_AdjTransaction!TxnID)
TxtTxnDate.Value = IIf(IsNull(RsTxn_AdjTransaction!TxnDate), "", RsTxn_AdjTransaction!TxnDate)
TxtPreFix = IIf(IsNull(RsTxn_AdjTransaction!SM_Prefix), "", RsTxn_AdjTransaction!SM_Prefix)
TxtCMPID = GetCMPName(IIf(IsNull(RsTxn_AdjTransaction!CMPID), "", RsTxn_AdjTransaction!CMPID))
mCMPID = IIf(IsNull(RsTxn_AdjTransaction!CMPID), "", RsTxn_AdjTransaction!CMPID)
TxtCommodityID = GetCommodityName(IIf(IsNull(RsTxn_AdjTransaction!CommodityID), "", RsTxn_AdjTransaction!CommodityID))
mCommodityID = IIf(IsNull(RsTxn_AdjTransaction!CommodityID), "", RsTxn_AdjTransaction!CommodityID)
If Not IsNull(RsTxn_AdjTransaction!AgentID) Then
   mAgentID = IIf(IsNull(RsTxn_AdjTransaction!AgentID), "", RsTxn_AdjTransaction!AgentID)
   TxtAgent = GetAgentName(IIf(IsNull(RsTxn_AdjTransaction!AgentID), "", RsTxn_AdjTransaction!AgentID))
End If
If RsTxn_AdjTransaction!TxnType = "Sa" Then
   TxtType = "Sale"
ElseIf RsTxn_AdjTransaction!TxnType = "Th" Then
   TxtType = "Theft"
ElseIf RsTxn_AdjTransaction!TxnType = "Fi" Then
   TxtType = "Fire"
ElseIf RsTxn_AdjTransaction!TxnType = "AR" Then
   TxtType = "Adjustment Received from client"
ElseIf RsTxn_AdjTransaction!TxnType = "AG" Then
   TxtType = "Adjustment Given to client"
End If
TxtAmount = IIf(IsNull(RsTxn_AdjTransaction!Amount), 0, RsTxn_AdjTransaction!Amount)

TxtCreated = IIf(IsNull(RsTxn_AdjTransaction!CreatedBy), "", RsTxn_AdjTransaction!CreatedBy) & " :- " & IIf(IsNull(RsTxn_AdjTransaction!CreatedDate), "", RsTxn_AdjTransaction!CreatedDate)
TxtUpdated = IIf(IsNull(RsTxn_AdjTransaction!UpdatedBy), "", RsTxn_AdjTransaction!UpdatedBy) & " :- " & IIf(IsNull(RsTxn_AdjTransaction!UpdatedDate), "", RsTxn_AdjTransaction!UpdatedDate)


CmbCMPID.Text = TxtCMPID
Call CmbCMPID_GotFocus
Call CmbCMPID_LostFocus
CmbCommodityID.Text = TxtCommodityID
Call CmbCommodityID_GotFocus
Call CmbCommodityID_LostFocus
Call InDataGSL
If RsTxn_AdjTransaction.RecordCount = 1 Then
   Call GButtonLock_1(Me, True)
Else
   Call GButtonLock(Me, True)
End If
CmdPrint.Enabled = True

End Sub

Private Sub InDataGSL()
Call TableTxn_GSLAdjTransaction(" Where TxnID = " & TxtTxnID & "")
MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + RsTxn_GSLAdjTransaction.RecordCount
If RsTxn_GSLAdjTransaction.RecordCount = 0 Then Exit Sub

Dim Ans As Variant

For i = 2 To MSHFlexGrid2.Rows - 2
    
    Ans = GetGSLByGLSID(RsTxn_GSLAdjTransaction!GSLID, TxtPreFix)
    MSHFlexGrid2.TextMatrix(i, 0) = RsTxn_GSLAdjTransaction!TxnGSLAdjID
    MSHFlexGrid2.TextMatrix(i, 1) = Ans(3) ' TmpRs!GodownNo
    MSHFlexGrid2.TextMatrix(i, 2) = Ans(4) 'TmpRs!StackNo
    MSHFlexGrid2.TextMatrix(i, 3) = Ans(5) 'TmpRs!LotNo
    MSHFlexGrid2.TextMatrix(i, 4) = "0" 'RsTxn_GSLAdjTransaction!NoofBags
    MSHFlexGrid2.TextMatrix(i, 5) = "0" 'RsTxn_GSLAdjTransaction!GrossWeight
    MSHFlexGrid2.TextMatrix(i, 6) = RsTxn_GSLAdjTransaction!NoofBags
    MSHFlexGrid2.TextMatrix(i, 7) = RsTxn_GSLAdjTransaction!Quantity
    MSHFlexGrid2.TextMatrix(i, 8) = RsTxn_GSLAdjTransaction!GSLID
    MSHFlexGrid2.TextMatrix(0, 6) = Val(MSHFlexGrid2.TextMatrix(0, 6)) + Val(MSHFlexGrid2.TextMatrix(i, 6))
    MSHFlexGrid2.TextMatrix(0, 7) = Val(MSHFlexGrid2.TextMatrix(0, 7)) + Val(MSHFlexGrid2.TextMatrix(i, 7))
    
    RsTxn_GSLAdjTransaction.MoveNext
Next
End Sub
Function CheckDuplicateGSL() As Boolean
Dim retVal As Boolean
retVal = False
If MSHFlexGrid2.Rows > 3 Then
   For u = 1 To MSHFlexGrid2.Rows - 1
       If MSHFlexGrid2.TextMatrix(u, 1) = CmbGodownNo.Text And _
          MSHFlexGrid2.TextMatrix(u, 2) = CmbStack.Text And _
          MSHFlexGrid2.TextMatrix(u, 3) = CmbLot.Text Then
          If LCase(CmdAdd2List.Caption) = "add to list" Then
             retVal = True
             Exit For
          Else
             If u <> MSHFlexGrid2.RowSel Then
                retVal = True
             End If
          End If
       End If
   Next
End If

CheckDuplicateGSL = retVal
End Function
Private Sub MSHFlexGrid1_Click()
Dim i As Variant
i = MSHFlexGrid1.RowSel
If i <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(i, 0) = "" Then Exit Sub
RsTxn_AdjTransaction.MoveFirst
RsTxn_AdjTransaction.Find "G_UID = '" & MSHFlexGrid1.TextMatrix(i, 9) & "'"
If Not RsTxn_AdjTransaction.EOF Then
   Call InData
   Frame0.Visible = False
End If
End Sub



Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsTxn_AdjTransaction.MoveNext
If RsTxn_AdjTransaction.EOF Then
   RsTxn_AdjTransaction.MoveLast
   LF_Flag = "Y"
End If
Call InData
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
RsTxn_AdjTransaction.MovePrevious
If RsTxn_AdjTransaction.BOF Then
   RsTxn_AdjTransaction.MoveFirst
   LF_Flag = "Y"
End If
Call InData
If LF_Flag = "Y" Then
   PreviousCmd.Enabled = False
   FirstCmd.Enabled = False
End If
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub


Private Sub FirstCmd_Click()
RsTxn_AdjTransaction.MoveFirst
Call InData
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True

End Sub


Private Sub LastCmd_Click()
RsTxn_AdjTransaction.MoveLast
Call InData
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False
End Sub


