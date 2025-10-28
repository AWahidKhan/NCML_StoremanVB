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
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      TabIndex        =   78
      Top             =   10560
      Width           =   10935
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   9015
      Left            =   15000
      TabIndex        =   77
      Top             =   0
      Width           =   255
   End
   Begin VB.Frame Frame4 
      Height          =   10095
      Left            =   120
      TabIndex        =   15
      Top             =   0
      Width           =   14895
      Begin VB.Frame Frame0 
         Height          =   1485
         Left            =   2400
         TabIndex        =   16
         Top             =   6120
         Visible         =   0   'False
         Width           =   10080
         Begin VB.CommandButton CmdExcel 
            Caption         =   "Excel"
            Height          =   350
            Left            =   12150
            TabIndex        =   22
            Top             =   735
            Width           =   1000
         End
         Begin VB.ComboBox CmbSCMPID 
            Height          =   315
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   21
            Top             =   360
            Width           =   3720
         End
         Begin VB.ComboBox CmbSAgent 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   7125
            Sorted          =   -1  'True
            TabIndex        =   20
            Top             =   360
            Width           =   3615
         End
         Begin VB.ComboBox CmbSCommodityID 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3900
            Sorted          =   -1  'True
            TabIndex        =   19
            Top             =   360
            Width           =   3180
         End
         Begin VB.ComboBox CmbSType 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "FrmTxn_AdjTransaction.frx":0000
            Left            =   10785
            List            =   "FrmTxn_AdjTransaction.frx":0013
            Sorted          =   -1  'True
            TabIndex        =   18
            Top             =   360
            Width           =   3240
         End
         Begin VB.CommandButton CmdGo 
            Caption         =   "Go"
            Height          =   345
            Left            =   10995
            TabIndex        =   17
            Top             =   735
            Width           =   1125
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   1455
            Left            =   90
            TabIndex        =   23
            Top             =   1140
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   2566
            _Version        =   393216
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin VB.Label Label11 
            Caption         =   "CMP"
            Height          =   300
            Left            =   1763
            TabIndex        =   29
            Top             =   180
            Width           =   435
         End
         Begin VB.Label Label12 
            Caption         =   "Commodity"
            Height          =   240
            Left            =   5070
            TabIndex        =   28
            Top             =   180
            Width           =   840
         End
         Begin VB.Label Label13 
            Caption         =   "Agent"
            Height          =   255
            Left            =   8430
            TabIndex        =   27
            Top             =   180
            Width           =   525
         End
         Begin VB.Label Label14 
            Caption         =   "Type"
            Height          =   255
            Left            =   12360
            TabIndex        =   26
            Top             =   180
            Width           =   450
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "_"
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
            Left            =   6840
            TabIndex        =   25
            Top             =   780
            Width           =   135
         End
         Begin VB.Label Label16 
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
            Left            =   4320
            TabIndex        =   24
            Top             =   765
            Width           =   2415
         End
      End
      Begin VB.Frame Frame3 
         Height          =   4605
         Left            =   120
         TabIndex        =   54
         Top             =   120
         Width           =   13065
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
            Left            =   6960
            TabIndex        =   6
            Top             =   1875
            Width           =   1920
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
            ItemData        =   "FrmTxn_AdjTransaction.frx":0067
            Left            =   1815
            List            =   "FrmTxn_AdjTransaction.frx":007A
            Sorted          =   -1  'True
            TabIndex        =   5
            Top             =   1875
            Width           =   3735
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
            Left            =   6960
            Locked          =   -1  'True
            TabIndex        =   60
            TabStop         =   0   'False
            Top             =   1035
            Width           =   1920
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
            Width           =   10290
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
            Left            =   1800
            Sorted          =   -1  'True
            TabIndex        =   2
            Top             =   630
            Width           =   10320
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
            TabIndex        =   59
            TabStop         =   0   'False
            Top             =   240
            Width           =   3675
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
            Left            =   12240
            TabIndex        =   58
            TabStop         =   0   'False
            Top             =   1470
            Width           =   375
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
            TabIndex        =   57
            TabStop         =   0   'False
            Top             =   1035
            Visible         =   0   'False
            Width           =   1320
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
            TabIndex        =   56
            TabStop         =   0   'False
            Top             =   4110
            Width           =   3930
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
            Left            =   1815
            Locked          =   -1  'True
            TabIndex        =   55
            TabStop         =   0   'False
            Top             =   4110
            Width           =   4845
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
            Height          =   1770
            Left            =   1815
            MaxLength       =   1000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   7
            Top             =   2295
            Width           =   10305
         End
         Begin MSComCtl2.DTPicker TxtTxnDate 
            Height          =   375
            Left            =   6960
            TabIndex        =   1
            Top             =   240
            Width           =   1950
            _ExtentX        =   3440
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
            Format          =   107675649
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
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   61
            Top             =   630
            Visible         =   0   'False
            Width           =   10320
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
            TabIndex        =   62
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
            Locked          =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   63
            TabStop         =   0   'False
            Top             =   1470
            Visible         =   0   'False
            Width           =   10290
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
            Locked          =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   64
            TabStop         =   0   'False
            Top             =   1875
            Visible         =   0   'False
            Width           =   3735
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
            Left            =   5880
            TabIndex        =   76
            Top             =   1928
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
            TabIndex        =   75
            Top             =   1928
            Width           =   855
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
            TabIndex        =   74
            Top             =   1523
            Width           =   855
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
            Height          =   285
            Left            =   120
            TabIndex        =   73
            Top             =   1088
            Width           =   1215
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
            Left            =   5880
            TabIndex        =   72
            Top             =   1080
            Width           =   1095
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
            Height          =   255
            Left            =   120
            TabIndex        =   71
            Top             =   683
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
            TabIndex        =   70
            Top             =   270
            Width           =   1800
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
            Left            =   5850
            TabIndex        =   69
            Top             =   270
            Width           =   1800
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
            TabIndex        =   68
            Top             =   1088
            Visible         =   0   'False
            Width           =   495
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
            TabIndex        =   67
            Top             =   4170
            Width           =   975
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
            TabIndex        =   66
            Top             =   4170
            Width           =   1695
         End
         Begin VB.Label Label17 
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
            TabIndex        =   65
            Top             =   3060
            Width           =   1800
         End
      End
      Begin VB.Frame Frame1 
         Height          =   3615
         Left            =   120
         TabIndex        =   40
         Top             =   4680
         Width           =   13080
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
            ItemData        =   "FrmTxn_AdjTransaction.frx":00CE
            Left            =   2640
            List            =   "FrmTxn_AdjTransaction.frx":00D0
            Sorted          =   -1  'True
            TabIndex        =   10
            Top             =   540
            Width           =   1215
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
            ItemData        =   "FrmTxn_AdjTransaction.frx":00D2
            Left            =   1440
            List            =   "FrmTxn_AdjTransaction.frx":00D4
            Sorted          =   -1  'True
            TabIndex        =   9
            Top             =   540
            Width           =   1215
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
            ItemData        =   "FrmTxn_AdjTransaction.frx":00D6
            Left            =   120
            List            =   "FrmTxn_AdjTransaction.frx":00D8
            Sorted          =   -1  'True
            TabIndex        =   8
            Top             =   540
            Width           =   1335
         End
         Begin VB.CommandButton CmdAdd2List 
            Caption         =   "Add to List"
            Height          =   345
            Left            =   10200
            TabIndex        =   13
            Top             =   555
            Width           =   1335
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
            Left            =   8400
            TabIndex        =   12
            Top             =   540
            Width           =   1575
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
            TabIndex        =   11
            Top             =   540
            Width           =   1380
         End
         Begin VB.CommandButton CmdDeleteList 
            Caption         =   "Delete From List"
            Height          =   345
            Left            =   11535
            TabIndex        =   45
            Top             =   555
            Width           =   1335
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
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   540
            Width           =   1560
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
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   540
            Width           =   1470
         End
         Begin VB.CheckBox ChkSpillage 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Check1"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   10440
            TabIndex        =   42
            Top             =   240
            Visible         =   0   'False
            Width           =   195
         End
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
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   120
            Width           =   375
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
            Height          =   2535
            Left            =   120
            TabIndex        =   46
            Top             =   960
            Width           =   12735
            _ExtentX        =   22463
            _ExtentY        =   4471
            _Version        =   393216
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
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
            TabIndex        =   53
            Top             =   240
            Width           =   855
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
            TabIndex        =   52
            Top             =   240
            Width           =   1215
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
            TabIndex        =   51
            Top             =   240
            Width           =   735
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
            TabIndex        =   50
            Top             =   240
            Width           =   975
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
            TabIndex        =   49
            Top             =   240
            Width           =   1335
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
            TabIndex        =   48
            Top             =   240
            Width           =   1305
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
            TabIndex        =   47
            Top             =   240
            Width           =   1305
         End
      End
      Begin VB.Frame ButtonFrm 
         Height          =   1200
         Left            =   120
         TabIndex        =   30
         Top             =   8280
         Width           =   13080
         Begin VB.CommandButton CmdPrint 
            Caption         =   "&Print"
            Height          =   810
            Left            =   11205
            TabIndex        =   39
            Top             =   195
            Width           =   1695
         End
         Begin VB.CommandButton ExitCmd 
            Caption         =   "E&xit"
            Height          =   400
            Left            =   9105
            TabIndex        =   38
            Top             =   600
            Width           =   2100
         End
         Begin VB.CommandButton LastCmd 
            Caption         =   "&Last"
            Height          =   400
            Left            =   6885
            TabIndex        =   37
            Top             =   600
            Width           =   2220
         End
         Begin VB.CommandButton FirstCmd 
            Caption         =   "&First"
            Height          =   400
            Left            =   4665
            TabIndex        =   36
            Top             =   600
            Width           =   2220
         End
         Begin VB.CommandButton PreviousCmd 
            Caption         =   "&Previous"
            Height          =   400
            Left            =   2325
            TabIndex        =   35
            Top             =   600
            Width           =   2340
         End
         Begin VB.CommandButton NextCmd 
            Caption         =   "&Next"
            Height          =   400
            Left            =   105
            TabIndex        =   34
            Top             =   600
            Width           =   2220
         End
         Begin VB.CommandButton SearchCmd 
            Caption         =   "Search"
            Height          =   400
            Left            =   9105
            TabIndex        =   33
            Top             =   195
            Width           =   2100
         End
         Begin VB.CommandButton DeleteCmd 
            Caption         =   "&Delete"
            Height          =   400
            Left            =   6885
            TabIndex        =   32
            Top             =   195
            Width           =   2220
         End
         Begin VB.CommandButton EditCmd 
            Caption         =   "&Edit"
            Height          =   400
            Left            =   4665
            TabIndex        =   31
            Top             =   195
            Width           =   2220
         End
         Begin VB.CommandButton SaveCmd 
            Caption         =   "Save"
            Height          =   400
            Left            =   2325
            TabIndex        =   14
            Top             =   195
            Width           =   2340
         End
         Begin VB.CommandButton AddCmd 
            Caption         =   "&Add New"
            Height          =   400
            Left            =   105
            TabIndex        =   0
            Top             =   195
            Width           =   2220
         End
      End
   End
End
Attribute VB_Name = "FrmTxn_AdjTransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg, mCMPID, mGodownID, mAgentID, mCommodityID, mType, mGSLID, mSType As Variant

Private Sub AddCmd_Click()
Edit_Flg = "A"

TxtTxnID = ""
TxtTxnDate.Enabled = True
TxtTxnDate = Date

TxtCMPID.Locked = False
CmbCMPID.Visible = True
TxtCMPID.Visible = False
CmbCMPID.Text = ""
TxtPreFix = ""
TxtCurrentDate = ""



TxtCommodityID.Locked = False
CmbCommodityID.Visible = True
TxtCommodityID.Visible = False
CmbCommodityID.Text = ""


CmbAgent.Visible = True
CmdAgent.Visible = True
TxtAgent.Visible = False
CmbAgent.Text = ""

TxtType.Visible = False
CmbType.Visible = True
CmbType.Text = ""

TxtAmount = ""
TxtAmount.Locked = False
TxtRemark.Locked = False
CmdPrint.Enabled = False
TxtRemark = ""

TxtCreated = ""
TxtUpdated = ""
'TxtTxnDate.SetFocus


Call GButtonLock(Me, False)
Call Grid_Head

End Sub
Private Sub CmbCMPID_GotFocus()
tmp = CmbCMPID.Text

wc = ""
If Gen_WcCMP <> "" Then
   wc = " And " & Gen_WcCMP
End If

Call TableMst_CMP(" where CloseDate is null " & wc)


CmbCMPID.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_CMP.RecordCount
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
For I = 1 To RsMst_Commodity.RecordCount
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
    For I = 1 To RsMst_Godown.RecordCount
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

Private Sub CmbSAgent_GotFocus()
tmp = CmbSAgent.Text
Call TableMst_Agent
CmbSAgent.Clear
If RsMst_Agent.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Agent.RecordCount
    CmbSAgent.AddItem RsMst_Agent!AgentName
    RsMst_Agent.MoveNext
Next
CmbSAgent.Text = tmp
End Sub

Private Sub CmbSAgent_LostFocus()
If CmbSAgent.Text = "" Then
   Exit Sub
End If
RsMst_Agent.MoveFirst
RsMst_Agent.Find "agentName = '" & CmbSAgent.Text & "'"
If RsMst_Agent.EOF Then
   MsgBox "Invalid selection "
   CmbSAgent.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmbSCMPID_GotFocus()
tmp = CmbSCMPID.Text
If Gen_WcCMP = "" Then
   Call TableMst_CMP ' (" where CloseDate is null ")
Else
   Call TableMst_CMP("Where " & Gen_WcCMP)
End If
CmbSCMPID.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_CMP.RecordCount
    CmbSCMPID.AddItem RsMst_CMP!CMPName
    RsMst_CMP.MoveNext
Next
CmbSCMPID.Text = tmp
End Sub

Private Sub CmbSCMPID_LostFocus()
If CmbSCMPID.Text = "" Then
   Exit Sub
End If

RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPName = '" & CmbSCMPID.Text & "'"
If RsMst_CMP.EOF Then
   MsgBox "Invalid selection "
   CmbSCMPID.SetFocus
   Exit Sub
End If

End Sub

Private Sub CmbSCommodityID_GotFocus()
tmp = CmbSCommodityID.Text
Call TableMst_Commodity
CmbSCommodityID.Clear
If RsMst_Commodity.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Commodity.RecordCount
    CmbSCommodityID.AddItem RsMst_Commodity!Commodity
    RsMst_Commodity.MoveNext
Next
CmbSCommodityID.Text = tmp

End Sub

Private Sub CmbSCommodityID_LostFocus()
If CmbSCommodityID.Text = "" Then
   Exit Sub
End If

RsMst_Commodity.MoveFirst
RsMst_Commodity.Find "Commodity = '" & CmbSCommodityID.Text & "'"
If RsMst_Commodity.EOF Then
   MsgBox "Invalid selection "
   CmbSCommodityID.SetFocus
   Exit Sub
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
'Gen_ExpireGSLFlag = " (MST_GSL.Flag <> 'E' And MST_GSL.Flag <> 'Z') "
If CmbType.Text = "Adjustment Received from client" Or CmbType.Text = "Adjustment Given to client" Then
   Call TableMst_GSL(" Where  " & Gen_ExpireGSLFlag & " And Flag <> 'B' And  CMPID = " & mCMPID & " And GodownID = " & mGodownID & " And StackNo = '" & CmbStack.Text & "' And WHRBags = 0")
Else
   Call TableMst_GSL(" Where  " & Gen_ExpireGSLFlag & " And Flag <> 'B' And  CMPID = " & mCMPID & " And GodownID = " & mGodownID & " And StackNo = '" & CmbStack.Text & "' ")
End If

CmbLot.Clear
If RsMst_GSL.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If

For I = 1 To RsMst_GSL.RecordCount
    CmbLot.AddItem RsMst_GSL!LotNo
    RsMst_GSL.MoveNext
Next I
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

Private Sub CmbSType_LostFocus()
If CmbSType.Text = "" Then
   mSType = Null
   Exit Sub
End If

If CmbSType.Text = "Sale" Or CmbSType.Text = "Theft" Or CmbSType.Text = "Fire" Or CmbSType.Text = "Adjustment Received from client" Or CmbSType.Text = "Adjustment Given to client" Then

Else
   MsgBox "Invalid selection!!!"
   CmbSType.SetFocus
   Exit Sub
End If

If CmbSType.Text = "Adjustment Received from client" Then
   mSType = "AR"
ElseIf CmbSType.Text = "Adjustment Given to client" Then
   mSType = "AG"
Else
   mSType = Left(CmbSType.Text, 2)
End If


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

If GetCDTFPrintPending(mGSLID) = True Then
   MsgBox "Print pending for one or more CDTF having this GSL!!!"
   CmbLot.SetFocus
   Exit Sub
End If

'Commented on 22 March 2013
'If mType = "AR" Or mType = "AD" Then
'   If CmbLot.Text <> "" Then
''      If mType = "AG" Then
'         If ChkSpillage.Value = 0 Then
'            MsgBox "Non Spillage Lot not allowed!!!"
'            CmbLot.SetFocus
'            Exit Sub
'         End If
''      End If
'   End If
'End If

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
I = MSHFlexGrid2.RowSel
If I > 1 Then
   If MSHFlexGrid2.TextMatrix(I, 1) <> "" Then
      MSHFlexGrid2.TextMatrix(0, 6) = Val(MSHFlexGrid2.TextMatrix(0, 6)) - Val(MSHFlexGrid2.TextMatrix(I, 6))
      MSHFlexGrid2.TextMatrix(0, 7) = Val(MSHFlexGrid2.TextMatrix(0, 7)) - Val(MSHFlexGrid2.TextMatrix(I, 7))
      For RO = I To MSHFlexGrid2.Rows - 2
          For C = 0 To MSHFlexGrid2.Cols - 1
               MSHFlexGrid2.TextMatrix(RO, C) = MSHFlexGrid2.TextMatrix(RO + 1, C)
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
For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = I + 1 ' + 6
    For C = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(I, C), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(I, C)) Then
              oWS.Cells(mRow, C + 1) = Format(MSHFlexGrid1.TextMatrix(I, C), "mm/dd/yyyy")
           Else
              oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
           End If
        Else
           oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
        End If
    Next
Next
oWB.Save
Set oWB = Nothing
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub

Private Sub CmdGo_Click()
Dim wc As Variant

tmp = "Select TAT.TxnID,TAT.TxnDate,MC.CMPName,TAT.SM_Prefix,MComm.Commodity,MA.AgentName,TAT.TxnType,TAT.Amount,TAT.Remarks,TAT.Flag,TAT.G_UID  " & _
      " From Txn_AdjTransaction TAT " & _
      " Inner Join Mst_CMP MC On TAT.CMPID = MC.CMPID " & _
      " Inner Join Mst_Commodity MComm On TAT.CommodityID = MComm.CommodityID " & _
      " Inner Join Mst_Agent MA On TAT.AgentID = MA.AgentID "

wc = ""

If CmbSCMPID.Text <> "" Then
   wc = " Where MC.CMPName = '" & CmbSCMPID.Text & "'"
Else
   If Gen_WcCMP <> "" Then
      wc = "Where MC." & Gen_WcCMP
   End If
End If

If CmbSCommodityID.Text <> "" Then
'   If Wc = "" Then
'      Wc = " Where MComm.Commodity = '" & CmbSCommodityID.Text & "'"
'   Else
      wc = wc & " And  MComm.Commodity = '" & CmbSCommodityID.Text & "'"
'   End If
End If

If CmbSAgent.Text <> "" Then
'   If Wc = "" Then
'      Wc = " Where MA.AgentName = '" & CmbSAgent.Text & "'"
'   Else
      wc = wc & " And  MA.AgentName = '" & CmbSAgent.Text & "'"
'   End If
End If

If CmbSType.Text <> "" Then
'   If Wc = "" Then
'      Wc = " Where TAT.TxnType = '" & mSType & "'"
'   Else
      wc = wc & " And  TAT.TxnType = '" & mSType & "'"
'   End If
End If

tmp = tmp & wc
Call TmpTable

Call Grid_Head1
For iu = 1 To TmpRs.RecordCount
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
    For C = 0 To MSHFlexGrid1.Cols - 1
        MSHFlexGrid1.TextMatrix(iu, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
    Next
    MSHFlexGrid1.TextMatrix(iu, 1) = Format(MSHFlexGrid1.TextMatrix(iu, 1), "dd-mmm-yyyy")
    If LCase(MSHFlexGrid1.TextMatrix(iu, 6)) = "sa" Then
       MSHFlexGrid1.TextMatrix(iu, 6) = "Sale"
    ElseIf LCase(MSHFlexGrid1.TextMatrix(iu, 6)) = "th" Then
       MSHFlexGrid1.TextMatrix(iu, 6) = "Theft"
    ElseIf LCase(MSHFlexGrid1.TextMatrix(iu, 6)) = "fi" Then
       MSHFlexGrid1.TextMatrix(iu, 6) = "Fire"
    ElseIf LCase(MSHFlexGrid1.TextMatrix(iu, 6)) = "ar" Then
       MSHFlexGrid1.TextMatrix(iu, 6) = "Adjustment Received from client"
    ElseIf LCase(MSHFlexGrid1.TextMatrix(iu, 6)) = "ag" Then
       MSHFlexGrid1.TextMatrix(iu, 6) = "Adjustment Given to client"
    End If
    TmpRs.MoveNext
Next
Label15.Refresh
Label15.Caption = TmpRs.RecordCount

End Sub
Private Sub CmdGSL_Click()
FrmMst_GSL.Show
End Sub

Private Sub DeleteCmd_Click()

Dim ans As Variant
On Error GoTo msgError

If LCase(Gen_UserRole) <> "power" Then
   MsgBox "Only power user can delete records!!!"
   Exit Sub
End If

Call TableMst_GSL("Where SM_Prefix = '" & TxtPreFix & "'")
If mType = "AR" Then
   If ChkGSL = False Then
      MsgBox "One or more GSL is Withdrwan. Can't Delete this Transaction!!!"
      Exit Sub
   End If
End If

ans = MsgBox("Do you really want to DELETE this Record??", vbYesNo)
If ans = vbYes Then
   For I = 2 To MSHFlexGrid2.Rows - 1
       RsMst_GSL.MoveFirst
       RsMst_GSL.Find "GSLID = " & Val(MSHFlexGrid2.TextMatrix(I, 8))
       If Not RsMst_GSL.EOF Then
          If mType = "AR" Then
             RsMst_GSL!DepositeBags = IIf(IsNull(RsMst_GSL!DepositeBags), 0, RsMst_GSL!DepositeBags) - Val(MSHFlexGrid2.TextMatrix(I, 6))
             RsMst_GSL!DepositeWeight = IIf(IsNull(RsMst_GSL!DepositeWeight), 0, RsMst_GSL!DepositeWeight) - Val(MSHFlexGrid2.TextMatrix(I, 7))
             RsMst_GSL!BalanceBags = IIf(IsNull(RsMst_GSL!BalanceBags), 0, RsMst_GSL!BalanceBags) - Val(MSHFlexGrid2.TextMatrix(I, 6))
             RsMst_GSL!BalanceWeight = IIf(IsNull(RsMst_GSL!BalanceWeight), 0, RsMst_GSL!BalanceWeight) - Val(MSHFlexGrid2.TextMatrix(I, 7))
          Else
             RsMst_GSL!CDTFBags = IIf(IsNull(RsMst_GSL!CDTFBags), 0, RsMst_GSL!CDTFBags) - Val(MSHFlexGrid2.TextMatrix(I, 6))
             RsMst_GSL!CDTFWeight = IIf(IsNull(RsMst_GSL!CDTFWeight), 0, RsMst_GSL!CDTFWeight) - Val(MSHFlexGrid2.TextMatrix(I, 7))
             RsMst_GSL!BalanceBags = IIf(IsNull(RsMst_GSL!BalanceBags), 0, RsMst_GSL!BalanceBags) + Val(MSHFlexGrid2.TextMatrix(I, 6))
             RsMst_GSL!BalanceWeight = IIf(IsNull(RsMst_GSL!BalanceWeight), 0, RsMst_GSL!BalanceWeight) + Val(MSHFlexGrid2.TextMatrix(I, 7))
             If RsMst_GSL!ExchangeTypeID = 1 Then
                RsMst_GSL!WHRBags = IIf(IsNull(RsMst_GSL!WHRBags), 0, RsMst_GSL!WHRBags) - Val(MSHFlexGrid2.TextMatrix(I, 6))
                RsMst_GSL!WHRWeight = IIf(IsNull(RsMst_GSL!WHRWeight), 0, RsMst_GSL!WHRWeight) - Val(MSHFlexGrid2.TextMatrix(I, 7))
             End If
          End If
          RsMst_GSL!Flag = IIf(IsNull(RsMst_GSL!Flag1), "O", RsMst_GSL!Flag1)
          If Val(RsMst_GSL!DepositeBags) = 0 And Val(RsMst_GSL!DepositeWeight) = 0 Then
             RsMst_GSL!CommodityID = Null
             RsMst_GSL!Flag = "B"
          End If
          RsMst_GSL.Update
       End If
   Next
   
   tmp = "Delete From Txn_GSLAdjTransaction Where TxnID = " & TxtTxnID & " "
   Call TmpTable
 
   RsTxn_AdjTransaction.Delete
   RsTxn_AdjTransaction.Update
   RsTxn_AdjTransaction.Requery
   If RsTxn_AdjTransaction.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsTxn_AdjTransaction.MoveNext
   If RsTxn_AdjTransaction.EOF() Then
      RsTxn_AdjTransaction.MoveLast
   End If
   Call Indata
End If
Exit Sub
msgError:
MsgBox "Child record Present "
RsTxn_AdjTransaction.CancelUpdate

End Sub
Private Function ChkGSL() As Boolean
Dim Retval As Boolean
Retval = True
For I = 2 To MSHFlexGrid2.Rows - 1
    If MSHFlexGrid2.TextMatrix(I, 1) = "" Then Exit For
    RsMst_GSL.MoveFirst
    RsMst_GSL.Find "GSLID = " & Val(MSHFlexGrid2.TextMatrix(I, 8)) & ""
    If Not RsMst_GSL.EOF = True Then
       If Val(RsMst_GSL!CDTFBags) <> 0 And Val(RsMst_GSL!CDTFWeight) <> 0 Then
          Retval = False
          Exit For
       End If
    End If
Next

ChkGSL = Retval
End Function

Private Sub EditCmd_Click()
'Edit_Flg = "E"
'CmdPrint.Enabled = False
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
For I = 1 To RsMst_Agent.RecordCount
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
Call Grid_Head1

CmbSAgent.Text = ""
CmbSType.Text = ""
CmbSCMPID.Text = ""
CmbSCommodityID.Text = ""
Label15.Caption = ""

If Gen_WcCMP = "" Then
   Call TableTxn_AdjTransaction("")
Else
   Call TableTxn_AdjTransaction("Where " & Gen_WcCMP)
End If
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
Call Indata

End Sub

Private Sub Form_Resize()
If Me.WindowState = vbMaximized Then
   SetScrollBars
End If
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

If CmbCMPID.Text = "" Then
   MsgBox "Blank CMP not allowed!!!"
   CmbCMPID.SetFocus
   Exit Sub
End If
If CmbCommodityID.Text = "" Then
   MsgBox "Blank Commodity not allowed!!!"
   CmbCommodityID.SetFocus
   Exit Sub
End If
If CmbAgent.Text = "" Then
   MsgBox "Blank Agent not allowed!!!"
   CmbAgent.SetFocus
   Exit Sub
End If

If CmbType.Text = "" Then
   MsgBox "Blank Type not allowed!!!"
   CmbType.SetFocus
   Exit Sub
End If


If LCase(Gen_UserRole) = "power" Or LCase(Gen_UserRole) = "headcmg" Or LCase(Gen_UserRole) = "superadmin" Then
  
   ' Do not check any validation changed on 20 jan 2011
Else
   If CmbType.Text = "Adjustment Received from client" Or CmbType.Text = "Adjustment Given to client" Then
      MsgBox "This option is for CMG users only. Please get in touch with CMG for updating this details in Storeman"
      CmbType.SetFocus
      Exit Sub
  End If
End If

If CmbType.Text = "Sale" Then
   If Val(TxtAmount) = 0 Then
      MsgBox "Amount can't be zero!!!"
      TxtAmount.SetFocus
      Exit Sub
   End If
End If

If TxtTxnDate.Value > CDate(TxtCurrentDate) Then
   'MsgBox "Transaction Date must be less than BOD date of CMP!!!"
   MsgBox "Transaction date cannot be later than BOD+1 date i.e." & TxtCurrentDate & "!!!"
   TxtTxnDate.SetFocus
   Exit Sub
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
RsTxn_AdjTransaction!Remarks = TxtRemark

If IsNull(RsTxn_AdjTransaction!CreatedBy) = True Or RsTxn_AdjTransaction!CreatedBy = "" Then
   RsTxn_AdjTransaction!CreatedBy = Gen_UserLoginID
   RsTxn_AdjTransaction!CreatedDate = Now
Else
   RsTxn_AdjTransaction!UpdatedBy = Gen_UserLoginID
   RsTxn_AdjTransaction!UpdatedDate = Now
End If


RsTxn_AdjTransaction.Update

Call SaveGSL
'SaveCmd.Enabled = False
'RsTxn_GSLAdjTransaction.Close

Call Indata


End Sub
Private Sub SaveGSL()
Call TableTxn_GSLAdjTransaction(" Where TxnID = " & TxtTxnID & "")
Call TableMst_GSL("where SM_Prefix = '" & TxtPreFix & "'")
For I = 2 To MSHFlexGrid2.Rows - 1
    If MSHFlexGrid2.TextMatrix(I, 1) <> "" Then
       RsTxn_GSLAdjTransaction.AddNew
       RsTxn_GSLAdjTransaction!TxnGSLAdjID = GetGSLAdjID
       RsTxn_GSLAdjTransaction!TxnGSLAdjDat = TxtTxnDate.Value
       RsTxn_GSLAdjTransaction!G_UID = GetGUID
       RsMst_GSL.MoveFirst
       RsMst_GSL.Find "GSLID = " & Val(MSHFlexGrid2.TextMatrix(I, 8))
       RsTxn_GSLAdjTransaction!TxnID = TxtTxnID
       RsTxn_GSLAdjTransaction!NoofBags = Val(MSHFlexGrid2.TextMatrix(I, 6))
       RsTxn_GSLAdjTransaction!Quantity = Val(MSHFlexGrid2.TextMatrix(I, 7))
       RsTxn_GSLAdjTransaction!GSLID = MSHFlexGrid2.TextMatrix(I, 8)
       RsTxn_GSLAdjTransaction!SpillageFlag = MSHFlexGrid2.TextMatrix(I, 9)
'       RsTxn_GSLAdjTransaction!SM_Prefix = RsMst_GSL!SM_Prefix
       RsTxn_GSLAdjTransaction.Update
       RsTxn_GSLAdjTransaction.Requery
       If Not RsMst_GSL.EOF Then
          If mType = "AR" Then
             If IsNull(RsMst_GSL!CommodityID) Then
                RsMst_GSL!CommodityID = mCommodityID
                RsMst_GSL!Flag1 = RsMst_GSL!Flag
                RsMst_GSL!Flag = "O"
             End If
             RsMst_GSL!DepositeBags = IIf(IsNull(RsMst_GSL!DepositeBags), 0, RsMst_GSL!DepositeBags) + Val(MSHFlexGrid2.TextMatrix(I, 6))
             RsMst_GSL!DepositeWeight = IIf(IsNull(RsMst_GSL!DepositeWeight), 0, RsMst_GSL!DepositeWeight) + Val(MSHFlexGrid2.TextMatrix(I, 7))
             RsMst_GSL!BalanceBags = IIf(IsNull(RsMst_GSL!BalanceBags), 0, RsMst_GSL!BalanceBags) + Val(MSHFlexGrid2.TextMatrix(I, 6))
             RsMst_GSL!BalanceWeight = IIf(IsNull(RsMst_GSL!BalanceWeight), 0, RsMst_GSL!BalanceWeight) + Val(MSHFlexGrid2.TextMatrix(I, 7))
          Else
             RsMst_GSL!CDTFBags = IIf(IsNull(RsMst_GSL!CDTFBags), 0, RsMst_GSL!CDTFBags) + Val(MSHFlexGrid2.TextMatrix(I, 6))
             RsMst_GSL!CDTFWeight = IIf(IsNull(RsMst_GSL!CDTFWeight), 0, RsMst_GSL!CDTFWeight) + Val(MSHFlexGrid2.TextMatrix(I, 7))
             RsMst_GSL!BalanceBags = IIf(IsNull(RsMst_GSL!BalanceBags), 0, RsMst_GSL!BalanceBags) - Val(MSHFlexGrid2.TextMatrix(I, 6))
             RsMst_GSL!BalanceWeight = IIf(IsNull(RsMst_GSL!BalanceWeight), 0, RsMst_GSL!BalanceWeight) - Val(MSHFlexGrid2.TextMatrix(I, 7))
             'Added on 22 March 2013 after discussion on 21th March 2013 with Mayur Sir and Sankar Sir
             If RsMst_GSL!ExchangeTypeID = 1 Then
                RsMst_GSL!WHRBags = IIf(IsNull(RsMst_GSL!WHRBags), 0, RsMst_GSL!WHRBags) + Val(MSHFlexGrid2.TextMatrix(I, 6))
                RsMst_GSL!WHRWeight = IIf(IsNull(RsMst_GSL!WHRWeight), 0, RsMst_GSL!WHRWeight) + Val(MSHFlexGrid2.TextMatrix(I, 7))
             End If
          End If
          If Val(RsMst_GSL!BalanceBags) = 0 Then
             RsMst_GSL!Flag1 = RsMst_GSL!Flag
             RsMst_GSL!Flag = "E"
          End If
          RsMst_GSL.Update
       End If
    End If
Next

End Sub
Private Function GetTxnID() As Double
Dim Retval As Double
tmp = "Select max(TxnID) from Txn_AdjTransaction" ' Where TxnType = 'D' And SM_Prefix = '" & TxtPreFix & "'"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetTxnID = Retval
End Function
Function GetGSLAdjID() As Double
Dim Retval As Double
Retval = 0
tmp = "select max(TxnGSLAdjID) from Txn_GSLAdjTransaction" ' where SM_Prefix = '" & TxtPreFix & "'"
Call TmpTable
If TmpRs.RecordCount > 0 Then
   Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
End If
Retval = Retval + 1
TmpRs.Close
GetGSLAdjID = Retval

End Function

Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   If RsTxn_AdjTransaction.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   Call Indata
   Exit Sub
End If

'Frame0.Height = Me.Height - 500
'Frame0.Left = Frame3.Left - 100
'Frame0.Top = Frame3.Top - 100
'Frame0.Width = Me.Width - 150
'Frame0.Visible = True
'MSHFlexGrid1.Width = Frame0.Width - 200
'MSHFlexGrid1.Height = Frame0.Height - 1200

Frame0.Height = Frame4.Height - 500
Frame0.Left = Frame3.Left
Frame0.Top = Frame3.Top
Frame0.Width = Frame4.Width - 200
Frame0.Visible = True
'MSHFlexGrid1.Clear
MSHFlexGrid1.Width = Frame0.Width - 300
MSHFlexGrid1.Height = Frame0.Height - 1400
RsTxn_AdjTransaction.MoveFirst

''===========comment on 7 july 2012 for not clearing search grid=============
'Call Grid_Head1
'
'CmbSAgent.Text = ""
'CmbSType.Text = ""
'CmbSCMPID.Text = ""
'CmbSCommodityID.Text = ""
'Label15.Caption = ""
''===========comment on 7 july 2012 for not clearing search grid=============


End Sub
Private Sub Grid_Head1()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 11


MSHFlexGrid1.TextMatrix(0, 0) = "Txn ID"
MSHFlexGrid1.TextMatrix(0, 1) = "Txn Date"
MSHFlexGrid1.TextMatrix(0, 2) = "CMP"
MSHFlexGrid1.TextMatrix(0, 3) = "SM_Prefix"
MSHFlexGrid1.TextMatrix(0, 4) = "Commodity"
MSHFlexGrid1.TextMatrix(0, 5) = "Agent"
MSHFlexGrid1.TextMatrix(0, 6) = "Txn Type"
MSHFlexGrid1.TextMatrix(0, 7) = "Amount"
MSHFlexGrid1.TextMatrix(0, 8) = "Remarks"
MSHFlexGrid1.TextMatrix(0, 9) = "Flag"
MSHFlexGrid1.TextMatrix(0, 10) = "G_UID"

MSHFlexGrid1.ColWidth(0) = 600
MSHFlexGrid1.ColWidth(1) = 1000
MSHFlexGrid1.ColWidth(2) = 1400
MSHFlexGrid1.ColWidth(3) = 800
MSHFlexGrid1.ColWidth(4) = 1600
MSHFlexGrid1.ColWidth(5) = 2400
MSHFlexGrid1.ColWidth(6) = 2000 '2000
MSHFlexGrid1.ColWidth(7) = 1500
MSHFlexGrid1.ColWidth(8) = 2400
MSHFlexGrid1.ColWidth(9) = 800
MSHFlexGrid1.ColWidth(10) = 0


End Sub
Private Sub TxtAmount_LostFocus()
If SaveCmd.Enabled = False Then Exit Sub
If TxtAmount = "" Then Exit Sub
If IsNumeric(TxtAmount) = False Or Val(TxtAmount) < 0 Then
   MsgBox "Invalid numeric format!!!"
   TxtAmount.SetFocus
   Exit Sub
End If

TxtAmount = Round(Val(TxtAmount), 2)
TxtRemark.SetFocus
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

TxtNoofBags = Round(Val(TxtNoofBags), 0)

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
Public Sub Indata()

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
TxtRemark.Locked = True
TxtTxnID = IIf(IsNull(RsTxn_AdjTransaction!TxnID), "", RsTxn_AdjTransaction!TxnID)
TxtTxnDate.Value = IIf(IsNull(RsTxn_AdjTransaction!TxnDate), "", RsTxn_AdjTransaction!TxnDate)
TxtPreFix = IIf(IsNull(RsTxn_AdjTransaction!SM_Prefix), "", RsTxn_AdjTransaction!SM_Prefix)
TxtCMPID = GetCMPName(IIf(IsNull(RsTxn_AdjTransaction!CMPID), "", RsTxn_AdjTransaction!CMPID))
mCMPID = IIf(IsNull(RsTxn_AdjTransaction!CMPID), "", RsTxn_AdjTransaction!CMPID)

'TxtCurrentDate = a 'IIf(IsNull(RsMst_CMP!CurrentDate), "", RsMst_CMP!CurrentDate) + 1

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

If TxtType.Text = "Adjustment Received from client" Then
   mType = "AR"
ElseIf TxtType.Text = "Adjustment Given to client" Then
   mType = "AG"
Else
   mType = Left(TxtType.Text, 2)
End If

TxtAmount = IIf(IsNull(RsTxn_AdjTransaction!Amount), 0, RsTxn_AdjTransaction!Amount)

TxtRemark = IIf(IsNull(RsTxn_AdjTransaction!Remarks), "", RsTxn_AdjTransaction!Remarks)
TxtCreated = IIf(IsNull(RsTxn_AdjTransaction!CreatedBy), "", RsTxn_AdjTransaction!CreatedBy) & " :- " & IIf(IsNull(RsTxn_AdjTransaction!CreatedDate), "", RsTxn_AdjTransaction!CreatedDate)
TxtUpdated = IIf(IsNull(RsTxn_AdjTransaction!UpdatedBy), "", RsTxn_AdjTransaction!UpdatedBy) & " :- " & IIf(IsNull(RsTxn_AdjTransaction!UpdatedDate), "", RsTxn_AdjTransaction!UpdatedDate)

Call InDataGSL

If RsTxn_AdjTransaction.RecordCount = 1 Then
   Call GButtonLock_1(Me, True)
Else
   Call GButtonLock(Me, True)
End If

'CmdPrint.Enabled = True

CmdPrint.Enabled = False


End Sub

Private Sub InDataGSL()
Call TableTxn_GSLAdjTransaction(" Where TxnID = " & TxtTxnID & "")
MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + RsTxn_GSLAdjTransaction.RecordCount
If RsTxn_GSLAdjTransaction.RecordCount = 0 Then Exit Sub

Dim ans As Variant

For I = 2 To MSHFlexGrid2.Rows - 2
    
    ans = GetGSLByGLSID(RsTxn_GSLAdjTransaction!GSLID, TxtPreFix)
    MSHFlexGrid2.TextMatrix(I, 0) = RsTxn_GSLAdjTransaction!TxnGSLAdjID
    MSHFlexGrid2.TextMatrix(I, 1) = ans(3) ' TmpRs!GodownNo
    MSHFlexGrid2.TextMatrix(I, 2) = ans(4) 'TmpRs!StackNo
    MSHFlexGrid2.TextMatrix(I, 3) = ans(5) 'TmpRs!LotNo
    MSHFlexGrid2.TextMatrix(I, 4) = "0" 'RsTxn_GSLAdjTransaction!NoofBags
    MSHFlexGrid2.TextMatrix(I, 5) = "0" 'RsTxn_GSLAdjTransaction!GrossWeight
    MSHFlexGrid2.TextMatrix(I, 6) = RsTxn_GSLAdjTransaction!NoofBags
    MSHFlexGrid2.TextMatrix(I, 7) = RsTxn_GSLAdjTransaction!Quantity
    MSHFlexGrid2.TextMatrix(I, 8) = RsTxn_GSLAdjTransaction!GSLID
    MSHFlexGrid2.TextMatrix(0, 6) = Val(MSHFlexGrid2.TextMatrix(0, 6)) + Val(MSHFlexGrid2.TextMatrix(I, 6))
    MSHFlexGrid2.TextMatrix(0, 7) = Val(MSHFlexGrid2.TextMatrix(0, 7)) + Val(MSHFlexGrid2.TextMatrix(I, 7))
    
    RsTxn_GSLAdjTransaction.MoveNext
Next
End Sub
Function CheckDuplicateGSL() As Boolean
Dim Retval As Boolean
Retval = False
If MSHFlexGrid2.Rows > 3 Then
   For u = 1 To MSHFlexGrid2.Rows - 1
       If MSHFlexGrid2.TextMatrix(u, 1) = CmbGodownNo.Text And _
          MSHFlexGrid2.TextMatrix(u, 2) = CmbStack.Text And _
          MSHFlexGrid2.TextMatrix(u, 3) = CmbLot.Text Then
          If LCase(CmdAdd2List.Caption) = "add to list" Then
             Retval = True
             Exit For
          Else
             If u <> MSHFlexGrid2.RowSel Then
                Retval = True
             End If
          End If
       End If
   Next
End If

CheckDuplicateGSL = Retval
End Function
Private Sub MSHFlexGrid1_Click()
Dim I As Variant
I = MSHFlexGrid1.RowSel
If I <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(I, 0) = "" Then Exit Sub
RsTxn_AdjTransaction.MoveFirst
RsTxn_AdjTransaction.Find "G_UID = '" & MSHFlexGrid1.TextMatrix(I, 10) & "'"
If Not RsTxn_AdjTransaction.EOF Then
   Call Indata
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
RsTxn_AdjTransaction.MovePrevious
If RsTxn_AdjTransaction.BOF Then
   RsTxn_AdjTransaction.MoveFirst
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
RsTxn_AdjTransaction.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True

End Sub
Private Sub LastCmd_Click()
RsTxn_AdjTransaction.MoveLast
Call Indata
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False
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
HScroll1.Max = Me.ScaleWidth - Frame1.Width
HScroll1.LargeChange = Me.ScaleWidth
HScroll1.SmallChange = Me.ScaleWidth / 5
If Frame4.Height >= Me.Height Then VScroll1.Enabled = True
If Frame4.Height < Me.Height Then VScroll1.Enabled = False
If Frame1.Width >= Me.Width Then HScroll1.Enabled = True
If Frame1.Width < Me.Width Then HScroll1.Enabled = False
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

Private Function GetCDTFPrintPending(mGSLID_ As Variant) As Boolean

Dim Retval As Boolean

Retval = False

tmp = "Select Distinct a.CxTFNo from Txn_CxTF_Details a"
tmp = tmp & " Inner Join Txn_CxTF_GSL b On a.CxTFNo = b.CxTFNo and a.TxnType = b.TxnType And a.SM_Prefix = b.SM_prefix"
tmp = tmp & " Where a.TxnType = 'D' and a.SM_Prefix = '" & TxtPreFix & "' and a.CMPID = " & mCMPID & " and isnull(a.Flag,'') = ''  and a.CommodityID = " & mCommodityID & " and b.GSLID = " & mGSLID_ & ""

Call TmpTable

If TmpRs.RecordCount > 0 Then
   Retval = True
End If
GetCDTFPrintPending = Retval

End Function

