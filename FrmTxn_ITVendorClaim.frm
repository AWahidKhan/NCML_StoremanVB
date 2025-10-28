VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmTxn_ITVendorClaim 
   Caption         =   "IT Vendor Claim"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.VScrollBar VScroll1 
      Height          =   9015
      Left            =   15000
      TabIndex        =   94
      Top             =   0
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      TabIndex        =   93
      Top             =   10320
      Width           =   10935
   End
   Begin VB.Frame Frame4 
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   9780
      Left            =   0
      TabIndex        =   36
      Top             =   -120
      Width           =   15180
      Begin VB.Frame Frame0 
         Height          =   1695
         Left            =   720
         TabIndex        =   81
         Top             =   5880
         Visible         =   0   'False
         Width           =   13680
         Begin VB.ComboBox CmbSVendor 
            Height          =   315
            Left            =   2745
            Sorted          =   -1  'True
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   330
            Width           =   3990
         End
         Begin VB.CommandButton CmdGo 
            Caption         =   "Go"
            Height          =   350
            Left            =   12240
            TabIndex        =   34
            Top             =   750
            Width           =   1000
         End
         Begin VB.CommandButton CmdExcel 
            Caption         =   "Excel"
            Height          =   350
            Left            =   13245
            TabIndex        =   35
            Top             =   750
            Width           =   1000
         End
         Begin VB.ComboBox CmbSStatus 
            Height          =   315
            ItemData        =   "FrmTxn_ITVendorClaim.frx":0000
            Left            =   6720
            List            =   "FrmTxn_ITVendorClaim.frx":0010
            Sorted          =   -1  'True
            TabIndex        =   30
            Top             =   330
            Width           =   1845
         End
         Begin VB.TextBox TxtSClaimID 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   1440
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   675
            Width           =   1335
         End
         Begin VB.ComboBox CmbSForwardTo 
            Height          =   315
            Left            =   8580
            Sorted          =   -1  'True
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   330
            Width           =   3435
         End
         Begin VB.CheckBox ChkBatchNo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Check1"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   4800
            TabIndex        =   33
            Top             =   720
            Width           =   225
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   1455
            Left            =   75
            TabIndex        =   82
            Top             =   1170
            Width           =   14355
            _ExtentX        =   25321
            _ExtentY        =   2566
            _Version        =   393216
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSComCtl2.DTPicker STo 
            Height          =   300
            Left            =   1395
            TabIndex        =   28
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
            Format          =   107413505
            CurrentDate     =   36494
         End
         Begin MSComCtl2.DTPicker SFrom 
            Height          =   300
            Left            =   60
            TabIndex        =   27
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
            Format          =   107413505
            CurrentDate     =   36494
         End
         Begin VB.Label Label44 
            Caption         =   "From Date"
            Height          =   255
            Left            =   360
            TabIndex        =   91
            Top             =   120
            Width           =   855
         End
         Begin VB.Label Label37 
            Caption         =   "Vendor Name"
            Height          =   225
            Left            =   4208
            TabIndex        =   90
            Top             =   150
            Width           =   1275
         End
         Begin VB.Label Label45 
            Caption         =   "To Date"
            Height          =   255
            Left            =   1680
            TabIndex        =   89
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
            Left            =   8520
            TabIndex        =   88
            Top             =   840
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
            Left            =   6000
            TabIndex        =   87
            Top             =   840
            Width           =   2415
         End
         Begin VB.Label Label16 
            Caption         =   "Status"
            Height          =   300
            Left            =   7395
            TabIndex        =   86
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
            TabIndex        =   85
            Top             =   727
            Width           =   900
         End
         Begin VB.Label Label31 
            Caption         =   "Forwarded To"
            Height          =   225
            Left            =   9705
            TabIndex        =   84
            Top             =   120
            Width           =   1275
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
            Left            =   3000
            TabIndex        =   83
            Top             =   720
            Width           =   1710
         End
      End
      Begin VB.Frame Frame1 
         Height          =   8595
         Left            =   120
         TabIndex        =   37
         Top             =   120
         Width           =   14880
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
            TabIndex        =   67
            TabStop         =   0   'False
            Top             =   1065
            Width           =   3600
         End
         Begin VB.TextBox TxtClaimID 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   375
            Left            =   1245
            Locked          =   -1  'True
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   240
            Width           =   1455
         End
         Begin VB.ComboBox CmbVendorID 
            BeginProperty Font 
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
            TabIndex        =   3
            Top             =   675
            Width           =   10455
         End
         Begin VB.CommandButton CmdVendorID 
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
            Left            =   11760
            TabIndex        =   43
            Top             =   720
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
            Height          =   480
            Left            =   1260
            MaxLength       =   200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   6
            Top             =   1905
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
            Left            =   13365
            Locked          =   -1  'True
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   240
            Visible         =   0   'False
            Width           =   1440
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
            ItemData        =   "FrmTxn_ITVendorClaim.frx":0039
            Left            =   1245
            List            =   "FrmTxn_ITVendorClaim.frx":0049
            Sorted          =   -1  'True
            TabIndex        =   41
            Top             =   1065
            Width           =   3600
         End
         Begin VB.TextBox TxtClaimAmount 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   375
            Left            =   9990
            Locked          =   -1  'True
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   1065
            Width           =   1710
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
            Height          =   435
            Left            =   1260
            Locked          =   -1  'True
            MaxLength       =   200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   7
            Top             =   2430
            Width           =   13560
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
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   1485
            Width           =   1455
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
            Left            =   11760
            TabIndex        =   38
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
            Left            =   1245
            Sorted          =   -1  'True
            TabIndex        =   5
            Top             =   1500
            Width           =   10455
         End
         Begin TabDlg.SSTab SSTab1 
            Height          =   5535
            Left            =   120
            TabIndex        =   45
            Top             =   3000
            Width           =   14655
            _ExtentX        =   25850
            _ExtentY        =   9763
            _Version        =   393216
            Tabs            =   2
            TabHeight       =   520
            TabCaption(0)   =   "Expense Details"
            TabPicture(0)   =   "FrmTxn_ITVendorClaim.frx":0072
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "MSHFlexGrid2"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "Frame2"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).ControlCount=   2
            TabCaption(1)   =   "Log History"
            TabPicture(1)   =   "FrmTxn_ITVendorClaim.frx":008E
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "TxtUpdatedByFAG"
            Tab(1).Control(0).Enabled=   0   'False
            Tab(1).Control(1)=   "TxtCreated"
            Tab(1).Control(1).Enabled=   0   'False
            Tab(1).Control(2)=   "TxtUpdated"
            Tab(1).Control(2).Enabled=   0   'False
            Tab(1).Control(3)=   "Label24"
            Tab(1).Control(4)=   "Label21"
            Tab(1).Control(5)=   "Label20"
            Tab(1).ControlCount=   6
            Begin VB.Frame Frame2 
               Height          =   1635
               Left            =   75
               TabIndex        =   49
               Top             =   315
               Width           =   14520
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
                  Left            =   45
                  Sorted          =   -1  'True
                  TabIndex        =   8
                  Top             =   480
                  Width           =   2550
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
                  Left            =   5520
                  TabIndex        =   51
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
                  Left            =   11205
                  TabIndex        =   16
                  Top             =   1140
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
                  Left            =   12660
                  TabIndex        =   17
                  Top             =   1140
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
                  Left            =   75
                  MaxLength       =   12
                  TabIndex        =   11
                  Top             =   1200
                  Width           =   1890
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
                  Height          =   480
                  Left            =   7035
                  MaxLength       =   250
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   10
                  Top             =   480
                  Width           =   7425
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
                  Left            =   2640
                  Sorted          =   -1  'True
                  TabIndex        =   9
                  Top             =   480
                  Width           =   4365
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
                  Left            =   3420
                  Sorted          =   -1  'True
                  TabIndex        =   13
                  Top             =   1200
                  Width           =   3135
               End
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
                  Left            =   6630
                  TabIndex        =   50
                  Top             =   1200
                  Width           =   375
               End
               Begin VB.TextBox TxtInvoiceNo 
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
                  Left            =   7065
                  MaxLength       =   50
                  TabIndex        =   14
                  Top             =   1200
                  Width           =   1710
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
                  Left            =   8820
                  MaxLength       =   12
                  TabIndex        =   15
                  Top             =   1200
                  Width           =   1710
               End
               Begin MSComCtl2.DTPicker TxtInvoiceDate 
                  Height          =   375
                  Left            =   2010
                  TabIndex        =   12
                  Top             =   1200
                  Width           =   1395
                  _ExtentX        =   2461
                  _ExtentY        =   661
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
                  Format          =   107413505
                  CurrentDate     =   39884
               End
               Begin VB.Label Label1 
                  Caption         =   "Invoice Date"
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
                  Left            =   2070
                  TabIndex        =   59
                  Top             =   885
                  Width           =   1320
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
                  Left            =   840
                  TabIndex        =   58
                  Top             =   150
                  Width           =   960
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
                  Left            =   600
                  TabIndex        =   57
                  Top             =   885
                  Width           =   855
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
                  Left            =   10230
                  TabIndex        =   56
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
                  Left            =   3855
                  TabIndex        =   55
                  Top             =   150
                  Width           =   1530
               End
               Begin VB.Label Label6 
                  Caption         =   "Asset Type"
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
                  Left            =   4320
                  TabIndex        =   54
                  Top             =   885
                  Width           =   1335
               End
               Begin VB.Label Label32 
                  Caption         =   "Invoice No"
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
                  Left            =   7545
                  TabIndex        =   53
                  Top             =   945
                  Width           =   1005
               End
               Begin VB.Label Label7 
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
                  Left            =   9300
                  TabIndex        =   52
                  Top             =   945
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
               TabIndex        =   48
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
               TabIndex        =   47
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
               TabIndex        =   46
               TabStop         =   0   'False
               Top             =   2520
               Width           =   9855
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
               Height          =   3405
               Left            =   90
               TabIndex        =   60
               Top             =   2025
               Width           =   14415
               _ExtentX        =   25426
               _ExtentY        =   6006
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
               TabIndex        =   63
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
               TabIndex        =   62
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
               TabIndex        =   61
               Top             =   2580
               Width           =   2175
            End
         End
         Begin MSComCtl2.DTPicker TxtClaimDate 
            Height          =   375
            Left            =   4200
            TabIndex        =   1
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   107413505
            CurrentDate     =   39884
         End
         Begin MSComCtl2.DTPicker TxtExpenseFrom 
            Height          =   375
            Left            =   7440
            TabIndex        =   2
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Format          =   107413505
            CurrentDate     =   39884
         End
         Begin MSComCtl2.DTPicker TxtExpenseTo 
            Height          =   375
            Left            =   10335
            TabIndex        =   64
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   107413505
            CurrentDate     =   39884
         End
         Begin MSComCtl2.DTPicker TxtApprovalDate 
            Height          =   375
            Left            =   6840
            TabIndex        =   4
            Top             =   1065
            Visible         =   0   'False
            Width           =   1710
            _ExtentX        =   3016
            _ExtentY        =   661
            _Version        =   393216
            Format          =   107413505
            CurrentDate     =   39884
         End
         Begin VB.TextBox TxtVendorID 
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
            TabIndex        =   65
            TabStop         =   0   'False
            Top             =   675
            Width           =   10455
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
            TabIndex        =   66
            TabStop         =   0   'False
            Top             =   1500
            Width           =   10455
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
            TabIndex        =   80
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
            TabIndex        =   79
            Top             =   292
            Width           =   1095
         End
         Begin VB.Label Label4 
            Caption         =   "Vendor"
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
            TabIndex        =   78
            Top             =   727
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
            Left            =   135
            TabIndex        =   77
            Top             =   2025
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
            Left            =   12420
            TabIndex        =   76
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
            TabIndex        =   75
            Top             =   1140
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
            Height          =   270
            Left            =   8685
            TabIndex        =   74
            Top             =   1110
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
            TabIndex        =   73
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
            TabIndex        =   72
            Top             =   292
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
            Left            =   4920
            TabIndex        =   71
            Top             =   1125
            Visible         =   0   'False
            Width           =   1980
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
            Left            =   135
            TabIndex        =   70
            Top             =   2475
            Width           =   1110
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
            TabIndex        =   69
            Top             =   1515
            Width           =   1065
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
            Left            =   105
            TabIndex        =   68
            Top             =   1560
            Width           =   1230
         End
      End
      Begin VB.Frame ButtonFrm 
         Height          =   1080
         Left            =   120
         TabIndex        =   92
         Top             =   8625
         Width           =   14865
         Begin VB.CommandButton ExitCmd 
            Caption         =   "E&xit"
            Height          =   400
            Left            =   12105
            TabIndex        =   22
            Top             =   600
            Width           =   2700
         End
         Begin VB.CommandButton SearchCmd 
            Caption         =   "Search"
            Height          =   400
            Left            =   12105
            TabIndex        =   21
            Top             =   195
            Width           =   2700
         End
         Begin VB.CommandButton LastCmd 
            Caption         =   "&Last"
            Height          =   400
            Left            =   9165
            TabIndex        =   23
            Top             =   600
            Width           =   2940
         End
         Begin VB.CommandButton DeleteCmd 
            Caption         =   "&Delete"
            Height          =   400
            Left            =   9165
            TabIndex        =   20
            Top             =   195
            Width           =   2940
         End
         Begin VB.CommandButton FirstCmd 
            Caption         =   "&First"
            Height          =   400
            Left            =   6225
            TabIndex        =   24
            Top             =   600
            Width           =   2940
         End
         Begin VB.CommandButton EditCmd 
            Caption         =   "&Edit"
            Height          =   400
            Left            =   6225
            TabIndex        =   19
            Top             =   195
            Width           =   2940
         End
         Begin VB.CommandButton PreviousCmd 
            Caption         =   "&Previous"
            Height          =   400
            Left            =   3165
            TabIndex        =   25
            Top             =   600
            Width           =   3060
         End
         Begin VB.CommandButton SaveCmd 
            Caption         =   "Save"
            Height          =   400
            Left            =   3165
            TabIndex        =   18
            Top             =   195
            Width           =   3060
         End
         Begin VB.CommandButton NextCmd 
            Caption         =   "&Next"
            Height          =   400
            Left            =   105
            TabIndex        =   26
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
End
Attribute VB_Name = "FrmTxn_ITVendorClaim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg, mExpenseTypeID, mVendorID, mSVendorID, mStatus, mSStatus, mVendorEmailID As Variant
Dim mLocationID, mSLocationID As Double
Dim mBatchID As Double
Dim mForwardToID, mForwardToEmailID, mSForwardToID, mAssetTypeId As Variant
Dim mMonth, mYear, mTotDays As Variant
'Dim oSMTP As SMTPSession

Private Sub Grid_Head()

MSHFlexGrid2.Clear
MSHFlexGrid2.Rows = 3
MSHFlexGrid2.FixedRows = 2
MSHFlexGrid2.Cols = 12


MSHFlexGrid2.TextMatrix(1, 0) = "Row ID"
MSHFlexGrid2.TextMatrix(1, 1) = "Location"
MSHFlexGrid2.TextMatrix(1, 2) = "Expense Type"
MSHFlexGrid2.TextMatrix(1, 3) = "Particulars"
MSHFlexGrid2.TextMatrix(1, 4) = "Amount"
MSHFlexGrid2.TextMatrix(1, 5) = "ExpenseTypeID"
MSHFlexGrid2.TextMatrix(1, 6) = "LocationID"
MSHFlexGrid2.TextMatrix(1, 7) = "Invoice Date"
MSHFlexGrid2.TextMatrix(1, 8) = "Asset Type"
MSHFlexGrid2.TextMatrix(1, 9) = "AssetTypeID"
MSHFlexGrid2.TextMatrix(1, 10) = "Invoice No"
MSHFlexGrid2.TextMatrix(1, 11) = "Quantity"

MSHFlexGrid2.ColWidth(0) = 500
MSHFlexGrid2.ColWidth(1) = 2000
MSHFlexGrid2.ColWidth(2) = 3200
MSHFlexGrid2.ColWidth(3) = 5000
MSHFlexGrid2.ColWidth(4) = 1200
MSHFlexGrid2.ColWidth(5) = 0
MSHFlexGrid2.ColWidth(6) = 0
MSHFlexGrid2.ColWidth(7) = 1000
MSHFlexGrid2.ColWidth(8) = 1000
MSHFlexGrid2.ColWidth(9) = 0
MSHFlexGrid2.ColWidth(10) = 800
MSHFlexGrid2.ColWidth(11) = 800

MSHFlexGrid2.RowHeight(1) = 650

End Sub

Private Sub AddCmd_Click()

Edit_Flg = "A"
Call GButtonLock(Me, False)

Call Grid_Head

'TxtClaimDate.Enabled = True
TxtVendorID.Visible = False
CmbVendorID.Visible = True
CmbVendorID.Text = ""
CmdVendorID.Visible = True

TxtForwardTo.Visible = False
CmbForwardTo.Visible = True
CmbForwardTo.Text = ""
CmdForwardTo.Visible = True

TxtRemark.Locked = False
TxtClaimID = ""
TxtClaimDate.Value = Date


TxtExpenseFrom.Enabled = True
'TxtExpenseTo.Enabled = True

TxtStatus.Visible = True
CmbStatus.Visible = False
TxtStatus = "Pending"
CmbStatus = ""
mStatus = "P"


TxtClaimAmount = ""
TxtRemark = ""

mVendorID = Null
mForwardToEmailID = Null
TxtFWDDesignation = ""

tmp = "Select VendorID,VendorName,EmailID From Mst_Vendor Where VendorName = '" & Gen_UserLoginID & "'"
Call TmpTable

If TmpRs.RecordCount = 1 Then
   CmbVendorID.Text = TmpRs!VendorName
   mVendorID = TmpRs!VendorID
   mVendorEmailID = TmpRs!EmailID
Else
   CmbVendorID.Text = ""
   mVendorID = Null
End If

mBatchID = 0
TxtCreated = ""
TxtUpdated = ""
TxtUpdatedByFAG = ""

Call ClearFrame
Frame2.Enabled = True

If RsTxn_ITVendorClaim.RecordCount > 0 Then
   TxtExpenseFrom.SetFocus
End If

'TxtExpenseFrom.Value = Date
'TxtExpenseTo.Value = Date
mMonth = Month(Date)
mYear = Year(Date)
TxtExpenseFrom.Value = mYear & "-" & mMonth & "- 01"
mTotDays = Day(DateSerial(mYear, mMonth + 1, 0))
TxtExpenseTo.Value = mYear & "-" & mMonth & "-" & mTotDays
SSTab1.Tab = 0

End Sub

Private Sub Form_Resize()
If Me.WindowState = vbMaximized Then
   SetScrollBars
End If

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
Private Sub ClearFrame()
TxtInvoiceNo.Text = ""
TxtQty.Text = ""
CmbLocationID.Text = ""
CmbExpenseTypeID.Text = ""
TxtParticulars = ""
TxtAmount = ""
TxtInvoiceDate.Value = Date
CmbAssetType.Text = ""
CmdAdd2List.Caption = "Add to List"

End Sub

Private Sub CmbSVendor_GotFocus()
tmp = CmbSVendor.Text
CmbSVendor.Clear
Call GetSearchVendorDataonGotFocus
CmbSVendor.Text = tmp
End Sub
Private Sub GetSearchVendorDataonGotFocus()

Call TableMst_ITVendor '("Where Flag = 'Y'")
CmbVendorID.Clear
If RsMst_ITVendor.RecordCount = 0 Then
   MsgBox "No Vendor Found!!!"
   CmbSVendor.SetFocus
   'RsMst_ITVendor.MoveFirst
   Exit Sub
End If
For I = 1 To RsMst_ITVendor.RecordCount
    CmbSVendor.AddItem RsMst_ITVendor!VendorName
    RsMst_ITVendor.MoveNext
Next

End Sub

Private Sub CmbSVendor_LostFocus()

Call GetSearchVendorDataOnLostFocus

End Sub

Private Sub CmbSForwardTo_GotFocus()
tmp = CmbSForwardTo.Text
Call TableMst_Employee("Where Flag = 'Y'")
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


End Sub

Private Sub CmdAdd2List_Click()
Dim r As Variant

If CmbLocationID.Text = "" Then
   MsgBox "Blank Location not allowed!!!"
   CmbLocationID.SetFocus
   Exit Sub
End If

If CmbExpenseTypeID.Text = "" Then
   MsgBox "Blank Expense Type not allowed!!!"
   CmbExpenseTypeID.SetFocus
   Exit Sub
End If

If TxtParticulars = "" Then
   MsgBox "Blank particulas not allowed!!!"
   TxtParticulars.SetFocus
   Exit Sub
End If

If Val(TxtAmount) = 0 Then
   MsgBox "Amount should be more than '0'!!!"
   TxtAmount.SetFocus
   Exit Sub
End If

If TxtInvoiceDate.Value > Date Then
   MsgBox "Future Invoice Date not allowed!!!"
   TxtInvoiceDate.SetFocus
   Exit Sub
End If

If CmbAssetType.Text = "" Then
   MsgBox "Blank Asset Type not allowed!!!"
   CmbAssetType.SetFocus
   Exit Sub
End If

If LCase(CmbExpenseTypeID.Text) = "assets to be capitalised" Then
   If TxtQty.Text = "" Then
      MsgBox "Blank Quantity not allowed for 'Assets to be capitalised'."
      TxtQty.SetFocus
      Exit Sub
   End If
End If

If CmdAdd2List.Caption = "Update List" Then
   r = MSHFlexGrid2.RowSel
   MSHFlexGrid2.TextMatrix(0, 4) = Val(MSHFlexGrid2.TextMatrix(0, 4)) - Val(MSHFlexGrid2.TextMatrix(r, 4))
   TxtClaimAmount = Val(TxtClaimAmount) - Val(MSHFlexGrid2.TextMatrix(r, 4))
Else
   r = MSHFlexGrid2.Rows - 1
   MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
   MSHFlexGrid2.TextMatrix(r, 0) = ""
End If


MSHFlexGrid2.TextMatrix(r, 1) = CmbLocationID.Text
MSHFlexGrid2.TextMatrix(r, 2) = CmbExpenseTypeID.Text
MSHFlexGrid2.TextMatrix(r, 3) = TxtParticulars
MSHFlexGrid2.TextMatrix(r, 4) = Val(TxtAmount)
MSHFlexGrid2.TextMatrix(r, 5) = mExpenseTypeID
MSHFlexGrid2.TextMatrix(r, 6) = mLocationID
MSHFlexGrid2.TextMatrix(r, 7) = TxtInvoiceDate
MSHFlexGrid2.TextMatrix(r, 8) = CmbAssetType.Text
MSHFlexGrid2.TextMatrix(r, 9) = mAssetTypeId
MSHFlexGrid2.TextMatrix(r, 10) = TxtInvoiceNo.Text
MSHFlexGrid2.TextMatrix(r, 11) = TxtQty.Text
 
MSHFlexGrid2.TextMatrix(0, 4) = Val(MSHFlexGrid2.TextMatrix(0, 4)) + Val(MSHFlexGrid2.TextMatrix(r, 4))
TxtClaimAmount = Val(TxtClaimAmount) + Val(MSHFlexGrid2.TextMatrix(r, 4))
Call ClearFrame

CmbLocationID.SetFocus

End Sub

Private Sub CmdAssetType_Click()
FrmMst_AssetType.Show
End Sub

Private Sub CmdDeleteList_Click()
I = MSHFlexGrid2.RowSel
If I > 0 Then
   If MSHFlexGrid2.TextMatrix(I, 1) <> "" Then
      TxtClaimAmount = Val(TxtClaimAmount) - Val(MSHFlexGrid2.TextMatrix(I, 4))
      MSHFlexGrid2.TextMatrix(0, 4) = Val(MSHFlexGrid2.TextMatrix(0, 4)) - Val(MSHFlexGrid2.TextMatrix(I, 4))
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

Private Sub CmdVendorID_Click()
FrmMst_ITVendor.Show
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
RsTxn_ITVendorClaim.Filter = ""

tmp = "Select VC.ClaimID,VC.ClaimDate,VC.ExpenseFrom,VC.ExpenseTo, " & _
      " MV.VendorName,VC.Amount,VC.Remarks ,VC.Flag, VC.G_UID, " & _
      " VC.ApprovalRemarks,VC.ApprovalDate, " & _
      " EMP.EmployeeName +'~'+ EMP.EmployeeNo as ForwardedTo,VC.BatchID " & _
      " From Txn_ITVendorClaim VC " & _
      " Inner Join Mst_ITVendor MV On VC.VendorID = MV.VendorID " & _
      " Left Outer Join Mst_Employee EMP On VC.ForwardToID = EMP.EmployeeID "

'tmp = tmp & " Where CS.ClaimDate between '2009-08-20 00:00:00' and '2009-08-20 00:00:00'"

If Gen_DBType = "MDB" Then
   tmp = tmp & " Where ClaimDate Between #" & Format(SFrom, Gen_DatFormat) & "# And #" & Format(STo, Gen_DatFormat) & "#) "
   wc = " ClaimDate >= #" & Format(SFrom, Gen_DatFormat) & "# And ClaimDate <= #" & Format(STo, Gen_DatFormat) & "#) "
Else
   tmp = tmp & " Where ClaimDate Between '" & Format(SFrom, Gen_DatFormat) & "' And '" & Format(STo, Gen_DatFormat) & "' "
   wc = " ClaimDate >= '" & Format(SFrom, Gen_DatFormat) & "' And ClaimDate <= '" & Format(STo, Gen_DatFormat) & "' "
End If

If CmbSVendor.Text <> "" Then
   tmp = tmp & " AND MV.VendorName = '" & CmbSVendor.Text & "' "
   wc = wc & " And VendorID = " & mSVendorID & ""
End If

If CmbSStatus.Text <> "" Then
   tmp = tmp & " AND VC.Flag = '" & mSStatus & "' "
   wc = wc & " And Flag = '" & mSStatus & "'"
End If

If CmbSForwardTo.Text <> "" Then
   strSplitString = Split(CmbSForwardTo, "~")
   EmpName_ = Trim(strSplitString(0))
   EmpNo_ = Trim(strSplitString(1))
 
   tmp = tmp & " AND EMP.EmployeeNo = '" & EmpNo_ & "' "
   wc = wc & " And ForwardToID = " & mSForwardToID & ""
End If

If TxtSClaimID <> "" Then
   tmp = tmp & " AND VC.ClaimID Like " & TxtSClaimID & " "
   wc = wc & " And ClaimID = " & TxtSClaimID & ""
End If

If ChkBatchNo.Value = 0 Then
   tmp = tmp & " AND VC.BatchID = 0 "
   wc = wc & " And BatchID = 0 "
End If

Call TmpTable

Call Grid_Head1
RsTxn_ITVendorClaim.Filter = "" & wc & ""

For iu = 1 To TmpRs.RecordCount
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
    For C = 0 To MSHFlexGrid1.Cols - 1
        MSHFlexGrid1.TextMatrix(iu, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
    Next
    If LCase(MSHFlexGrid1.TextMatrix(iu, 7)) = "a" Then
       MSHFlexGrid1.TextMatrix(iu, 7) = "Approved"
    ElseIf LCase(MSHFlexGrid1.TextMatrix(iu, 7)) = "p" Then
       MSHFlexGrid1.TextMatrix(iu, 7) = "Pending"
    ElseIf LCase(MSHFlexGrid1.TextMatrix(iu, 7)) = "r" Then
       MSHFlexGrid1.TextMatrix(iu, 7) = "Rejected"
    ElseIf LCase(MSHFlexGrid1.TextMatrix(iu, 7)) = "f" Then
       MSHFlexGrid1.TextMatrix(iu, 7) = "Frozen"
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
TxtVendorID.Visible = False
CmbVendorID.Visible = True
CmbVendorID.Text = TxtVendorID
CmdVendorID.Visible = True

TxtForwardTo.Visible = False
CmbForwardTo.Visible = True
CmbForwardTo.Text = TxtForwardTo
CmdForwardTo.Visible = True


If LCase(Gen_UserRole) = "fag" Or LCase(Gen_UserRole) = "fagadmin" Or LCase(Gen_UserRole) = "power" Then
   TxtStatus.Visible = False
   CmbStatus.Visible = True
   CmbStatus.Text = TxtStatus
End If

TxtApprovalDate.Value = Date
TxtExpenseFrom.Enabled = True
'TxtExpenseTo.Enabled = True
TxtRemark.Locked = False
Call ClearFrame
Frame2.Enabled = True
CmbVendorID.SetFocus
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
Call Grid_Head1

SFrom.Value = CDate(From_GenDate) - 900
STo.Value = Date 'To_GenDate
CmbSVendor.Text = ""
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

Call TableTxn_ITVendorClaim(" Order by ClaimID")
If RsTxn_ITVendorClaim.RecordCount = 0 Then
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
End Sub
Private Sub GetVendorinList()

tmp = CmbVendorID.Text
Call TableMst_ITVendor '("Where Flag = 'Y'")
CmbVendorID.Clear
If RsMst_ITVendor.RecordCount = 0 Then
   MsgBox "No Vendor Found!!!"
   CmbVendorID.SetFocus
   'RsMst_ITVendor.MoveFirst
   Exit Sub
End If
For I = 1 To RsMst_ITVendor.RecordCount
    CmbVendorID.AddItem RsMst_ITVendor!VendorName
    RsMst_ITVendor.MoveNext
Next

CmbVendorID.Text = tmp


End Sub

Private Sub CmbVendorID_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "VendorID")
End If
End Sub
Private Sub GetVendorDataOnLostFocus()

If CmbVendorID.Text = "" Then
   mVendorID = Null
   Exit Sub
End If


RsMst_ITVendor.MoveFirst
RsMst_ITVendor.Find "VendorName = '" & CmbVendorID.Text & "' "

If RsMst_ITVendor.EOF Then
   MsgBox "Invalid Selection "
   CmbVendorID.SetFocus
   Exit Sub
End If
mVendorID = RsMst_ITVendor!VendorID
mVendorEmailID = RsMst_ITVendor!EmailID

End Sub
Private Sub CmbVendorID_LostFocus()
Call GetVendorDataOnLostFocus
End Sub
Private Sub CmbExpenseTypeID_GotFocus()
Dim wc As String
tmp = CmbExpenseTypeID.Text
Call TableMst_ExpenseType(" Where (AllocateTo = 'B' Or AllocateTo = 'V') And Flag = 'A' And OtherDeductionFlag = 'N' ")
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
Private Sub MSHFlexGrid1_Click()
Dim I As Variant
I = MSHFlexGrid1.RowSel
If I <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(I, 0) = "" Then Exit Sub
RsTxn_ITVendorClaim.MoveFirst
RsTxn_ITVendorClaim.Find "G_UID = '" & MSHFlexGrid1.TextMatrix(I, 8) & "'"
If Not RsTxn_ITVendorClaim.EOF Then
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

CmbLocationID.Text = MSHFlexGrid2.TextMatrix(r, 1)
mLocationID = MSHFlexGrid2.TextMatrix(r, 6)
CmbExpenseTypeID.Text = MSHFlexGrid2.TextMatrix(r, 2)
mExpenseTypeID = MSHFlexGrid2.TextMatrix(r, 5)
TxtParticulars = MSHFlexGrid2.TextMatrix(r, 3)
TxtAmount = Val(MSHFlexGrid2.TextMatrix(r, 4))
TxtInvoiceDate = IIf(MSHFlexGrid2.TextMatrix(r, 7) = "", Date, MSHFlexGrid2.TextMatrix(r, 7))
CmbAssetType.Text = MSHFlexGrid2.TextMatrix(r, 8)
mAssetTypeId = Val(MSHFlexGrid2.TextMatrix(r, 9))
TxtInvoiceNo = MSHFlexGrid2.TextMatrix(r, 10)
TxtQty = MSHFlexGrid2.TextMatrix(r, 11)
CmdAdd2List.Caption = "Update List"
CmbExpenseTypeID.SetFocus

End Sub
Private Sub SaveCmd_Click()

If TxtExpenseFrom.Value > TxtExpenseTo.Value Then
   MsgBox "'Expense From ' must be less than 'Expense To'!!!"
   TxtExpenseFrom.SetFocus
   Exit Sub
End If

If TxtExpenseFrom.Value > Date Then
   MsgBox "'Expense From ' must be less than 'Claim Date'!!!"
   TxtExpenseFrom.SetFocus
   Exit Sub
End If

If CmbVendorID.Text = "" Then
   MsgBox "Blank IT Vendor not allowed!!!"
   CmbVendorID.SetFocus
   Exit Sub
End If

If Edit_Flg = "E" Then
   If LCase(Gen_UserRole) = "fag" Or LCase(Gen_UserRole) = "fagadmin" Or LCase(Gen_UserRole) = "power" Then
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

   If ChkExpenseType = False Then
      MsgBox "Expense Type must present in all rows!!!"
      Exit Sub
   End If
End If

If CmbForwardTo.Text = "" Then
   MsgBox "Blank Forward to not allowed!!!"
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
   RsTxn_ITVendorClaim.AddNew
   RsTxn_ITVendorClaim!ClaimID = GetMaxClaimID
   RsTxn_ITVendorClaim!G_UID = GetGUID
   TxtClaimID = RsTxn_ITVendorClaim!ClaimID
   TxtG_UID = RsTxn_ITVendorClaim!G_UID
Else
   RsTxn_ITVendorClaim.MoveFirst
   RsTxn_ITVendorClaim.Find "G_UID= '" & TxtG_UID & "'"
End If

RsTxn_ITVendorClaim!ClaimDate = TxtClaimDate.Value
RsTxn_ITVendorClaim!VendorID = mVendorID

RsTxn_ITVendorClaim!Amount = Val(TxtClaimAmount)
RsTxn_ITVendorClaim!Flag = mStatus
'RsTxn_ITVendorClaim!Remarks = TxtRemark

TxtRemark = Replace(TxtRemark, ",", "-")
RsTxn_ITVendorClaim!Remarks = TxtRemark

If TxtApprovalRemarks.Locked = False Then
   TxtApprovalRemarks = Replace(TxtApprovalRemarks, ",", "-")
   RsTxn_ITVendorClaim!ApprovalRemarks = Trim(TxtApprovalRemarks)
End If
If TxtApprovalDate.Visible = True Then
   RsTxn_ITVendorClaim!ApprovalDate = TxtApprovalDate.Value
End If
RsTxn_ITVendorClaim!ForwardToID = mForwardToID
RsTxn_ITVendorClaim!G_UID = TxtG_UID
RsTxn_ITVendorClaim!ExpenseFrom = TxtExpenseFrom.Value
RsTxn_ITVendorClaim!ExpenseTo = TxtExpenseTo.Value
RsTxn_ITVendorClaim!BatchID = mBatchID


If IsNull(RsTxn_ITVendorClaim!CreatedBy) = True Or RsTxn_ITVendorClaim!CreatedBy = "" Then
   RsTxn_ITVendorClaim!CreatedBy = Gen_UserLoginID
   RsTxn_ITVendorClaim!CreatedDate = Now
Else
   RsTxn_ITVendorClaim!UpdatedBy = Gen_UserLoginID
   RsTxn_ITVendorClaim!UpdatedDate = Now
End If

If LCase(Gen_UserRole) = "fag" Or LCase(Gen_UserRole) = "fagadmin" Or LCase(Gen_UserRole) = "power" Then
   If LCase(mStatus) = "a" Then
      RsTxn_ITVendorClaim!UpdatedByFAG = Gen_UserLoginID
      RsTxn_ITVendorClaim!UpdatedDateFAG = Now
   End If
End If
'RsTxn_ITVendorClaim!Concurrency = Now
RsTxn_ITVendorClaim.Update

TxtCreated = IIf(IsNull(RsTxn_ITVendorClaim!CreatedBy), "", RsTxn_ITVendorClaim!CreatedBy) & " :- " & IIf(IsNull(RsTxn_ITVendorClaim!CreatedDate), "", RsTxn_ITVendorClaim!CreatedDate)
TxtUpdated = IIf(IsNull(RsTxn_ITVendorClaim!UpdatedBy), "", RsTxn_ITVendorClaim!UpdatedBy) & " :- " & IIf(IsNull(RsTxn_ITVendorClaim!UpdatedDate), "", RsTxn_ITVendorClaim!UpdatedDate)

Call SaveVendorClaimDetails

If LCase(mStatus) <> "f" Then
   Call Gen_Email
End If

RsTxn_ITVendorClaim.Filter = ""
RsTxn_ITVendorClaim.Requery
RsTxn_ITVendorClaim.MoveFirst
RsTxn_ITVendorClaim.Find "G_UID = '" & TxtG_UID & "'"

Call Indata

End Sub
Private Sub SaveVendorClaimDetails()
Call DeleteVendorClaimDetails
Call TableTxn_ITVendorClaimDetails("") '(" Where ClaimID = " & TxtClaimID & " AND SM_Prefix = '" & TxtPreFix & "'")
For I = 2 To MSHFlexGrid2.Rows - 1
    If MSHFlexGrid2.TextMatrix(I, 1) <> "" Then
       RsTxn_ITVendorClaimDetails.AddNew
       RsTxn_ITVendorClaimDetails!ClaimDetailID = GetMaxVendorClaimDetails
       RsTxn_ITVendorClaimDetails!ClaimID = TxtClaimID
       RsTxn_ITVendorClaimDetails!LocationID = MSHFlexGrid2.TextMatrix(I, 6)
       RsTxn_ITVendorClaimDetails!ExpenseTypeID = MSHFlexGrid2.TextMatrix(I, 5)
       RsTxn_ITVendorClaimDetails!Particulars = MSHFlexGrid2.TextMatrix(I, 3)
       RsTxn_ITVendorClaimDetails!Amount = Val(MSHFlexGrid2.TextMatrix(I, 4))
       If MSHFlexGrid2.TextMatrix(I, 7) <> "" Then
          RsTxn_ITVendorClaimDetails!InvoiceDate = MSHFlexGrid2.TextMatrix(I, 7)
       End If
       RsTxn_ITVendorClaimDetails!AssetsID = MSHFlexGrid2.TextMatrix(I, 9)
       If MSHFlexGrid2.TextMatrix(I, 10) <> "" Then
          RsTxn_ITVendorClaimDetails!InvoiceNo = MSHFlexGrid2.TextMatrix(I, 10)
       End If
       If MSHFlexGrid2.TextMatrix(I, 11) <> "" Then
          RsTxn_ITVendorClaimDetails!Quantity = Val(MSHFlexGrid2.TextMatrix(I, 11))
       End If
      ' RsTxn_ITVendorClaimDetails!Concurrency = Now
       RsTxn_ITVendorClaimDetails.Update
       RsTxn_ITVendorClaimDetails.Requery
    End If
Next
End Sub
Private Function ChkExpenseType() As Boolean
Dim Retval As Boolean
Retval = True

For I = 2 To MSHFlexGrid2.Rows - 1
    If MSHFlexGrid2.TextMatrix(I, 1) <> "" Then
       If MSHFlexGrid2.TextMatrix(I, 8) = "" Then
          Retval = False
          Exit For
       End If
    End If
Next

ChkExpenseType = Retval
End Function
Private Sub Gen_Email()
Dim TOADD As Variant
Dim TOCC, mMsg, mGridMsg As Variant
Dim strSplitString() As String
Dim EmpNo_, EmpName_ As String

mMsg = ""
mMsg = "<tr>IT Vendor Claim Submission (From Store-Man)</tr>"
mMsg = mMsg & "<tr><td> Claim ID </td><td> " & TxtClaimID & " </td></tr>"
mMsg = mMsg & "<tr><td> Claim Date </td><td> " & Format(TxtClaimDate.Value, "dd-mmm-yyyy") & "  </td></tr>"
If CmbStatus.Visible = True Then
   mMsg = mMsg & "<tr><td> Status </td><td> " & CmbStatus.Text & "  </td></tr>"
Else
   mMsg = mMsg & "<tr><td> Status </td><td> " & TxtStatus & "  </td></tr>"
End If

mMsg = mMsg & "<tr><td> Vendor Name </td><td> " & CmbVendorID.Text & "  </td></tr>"
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
       mGridMsg = mGridMsg & "<tr><td> " & MSHFlexGrid2.TextMatrix(I, 1) & " </td><td> " & MSHFlexGrid2.TextMatrix(I, 2) & " </td><td> " & MSHFlexGrid2.TextMatrix(I, 3) & " </td><td> " & MSHFlexGrid2.TextMatrix(I, 4) & " </td><td> " & Format(MSHFlexGrid2.TextMatrix(I, 7), "dd-mmm-yyyy") & " </td><td> " & MSHFlexGrid2.TextMatrix(I, 8) & " </td></tr>"
    End If
Next

mMsg = mMsg & "<tr><td> Details </td><td><table border = 1>" & mGridMsg & "</table></td>"


'TOADD = "shajil.p@ncml.com" 'mForwardToEmailID ' mVendorEmailID
'TOCC = "parag.r@ncml.com" 'Gen_UserEmailID '& "," & mForwardToEmailID

TOADD = mForwardToEmailID ' mVendorEmailID
TOCC = Gen_UserEmailID '& "," & mForwardToEmailID


If LCase(mStatus) = "p" Then
   mMsgSubject = "IT Vendor Claim Pending"
ElseIf LCase(mStatus) = "a" Then
   mMsgSubject = "IT Vendor Claim Approved"
ElseIf LCase(mStatus) = "r" Then
   mMsgSubject = "IT Vendor Claim Rejected"
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
'  .MailFrom = TOCC_ '"mayur.d@ncml.com"
'  If IsNull(Trim(ToADD_)) = True Or Trim(ToADD_) = "" Then
'     ToADD_ = "gaurav.k@ncml.com"
'  End If
'  .SendTo = ToADD_
'  TOCC_ = TOCC_ '& "," & mVendorEmailID
'  If ToADD_ <> TOCC_ Then
'     .CC = TOCC_
'  End If
'  '.BCC = "mayur.d@ncml.com"
'  .BCC = "parag.r@ncml.com"
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
objSMTP.FromAddr = TOCC_
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
   If RsTxn_ITVendorClaim.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   RsTxn_ITVendorClaim.MoveFirst
 
   RsTxn_ITVendorClaim.Find "G_UID = '" & TxtG_UID & "'"
   If RsTxn_ITVendorClaim.EOF Then
      RsTxn_ITVendorClaim.MoveLast
   End If
   Call Indata
   Exit Sub
End If

Frame0.Height = Frame4.Height - 200 'Me.Height - 500
Frame0.Left = Frame1.Left - 100
Frame0.Top = Frame1.Top
Frame0.Width = Frame4.Width - 200 'Me.Width - 200
Frame0.Visible = True

MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 1400

RsTxn_ITVendorClaim.MoveFirst

''===========comment on 7 july 2012 for not clearing search grid=============
'Call Grid_Head1
'
'SFrom.Value = CDate(From_GenDate) - 900
'STo.Value = Date 'To_GenDate
'CmbSVendor.Text = ""
'CmbSStatus.Text = "Pending"
'CmbSForwardTo.Text = ""
'ChkBatchNo.Value = 0
'TxtSClaimID = ""
'Label22.Caption = ""
'mSStatus = "P"
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

Private Sub TxtAmount_LostFocus()
If TxtAmount = "" Then Exit Sub
If IsNumeric(TxtAmount) = False Or Val(TxtAmount) < 0 Then
   MsgBox "Invalid Format!!!"
   TxtAmount.SetFocus
   Exit Sub
End If

TxtAmount = Format(TxtAmount, "########0.00")
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

Private Sub DeleteVendorClaimDetails()
tmp = "Delete From Txn_ITVendorClaimDetail Where ClaimID=" & TxtClaimID & " "
Call Tmp4Table
End Sub
Private Function GetMaxClaimID() As Double
Dim Retval As Double
tmp = "Select max(ClaimID) from Txn_ITVendorClaim "
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetMaxClaimID = Retval
End Function

Function GetMaxVendorClaimDetails() As Double
Dim Retval As Double
Retval = 0
tmp = "Select max(ClaimDetailID) From Txn_ITVendorClaimDetail " 'Where ClaimID = " & TxtClaimID & " And SM_prefix = '" & TxtPreFix & "'"
Call TmpTable
If TmpRs.RecordCount > 0 Then
   Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
End If
Retval = Retval + 1
TmpRs.Close
GetMaxVendorClaimDetails = Retval

End Function
Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsTxn_ITVendorClaim.MoveNext
If RsTxn_ITVendorClaim.EOF Then
   RsTxn_ITVendorClaim.MoveLast
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
RsTxn_ITVendorClaim.MovePrevious
If RsTxn_ITVendorClaim.BOF Then
   RsTxn_ITVendorClaim.MoveFirst
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
RsTxn_ITVendorClaim.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsTxn_ITVendorClaim.MoveLast
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
TxtVendorID.Visible = True
CmbVendorID.Visible = False
CmdVendorID.Visible = Fasle

TxtForwardTo.Visible = True
CmbForwardTo.Visible = False
CmdForwardTo.Visible = False

TxtStatus.Visible = True
CmbStatus.Visible = False

TxtExpenseFrom.Enabled = False
TxtExpenseTo.Enabled = False
TxtClaimDate.Enabled = False

TxtClaimID = RsTxn_ITVendorClaim!ClaimID
TxtClaimDate.Value = RsTxn_ITVendorClaim!ClaimDate
TxtExpenseFrom.Value = RsTxn_ITVendorClaim!ExpenseFrom
TxtExpenseTo.Value = RsTxn_ITVendorClaim!ExpenseTo
mVendorID = RsTxn_ITVendorClaim!VendorID
TxtVendorID = GetITVendorName(RsTxn_ITVendorClaim!VendorID, "N")
mVendorEmailID = GetITVendorName(mVendorID, "E")
TxtClaimAmount = RsTxn_ITVendorClaim!Amount 'a 'Val(MSHFlexGrid2.TextMatrix(0, 4))

mStatus = RsTxn_ITVendorClaim!Flag
mBatchID = IIf(IsNull(RsTxn_ITVendorClaim!BatchID), 0, RsTxn_ITVendorClaim!BatchID)
mForwardToID = IIf(IsNull(RsTxn_ITVendorClaim!ForwardToID), "", RsTxn_ITVendorClaim!ForwardToID)

If mForwardToID <> "" Then
   TxtForwardTo = GetEmployeeNonName(RsTxn_ITVendorClaim!ForwardToID)
   mForwardToEmailID = GetEmployeeEmailID(mForwardToID)
Else
   TxtForwardTo = "" 'GetVendorNonName(RsTxn_ITVendorClaim!ForwardToID)
   mForwardToEmailID = Null 'GetVendorEmailID(mForwardToID)
End If
TxtFWDDesignation = GetDesignationFromEmp(mForwardToID)

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
TxtRemark = IIf(IsNull(RsTxn_ITVendorClaim!Remarks), "", RsTxn_ITVendorClaim!Remarks)
TxtApprovalRemarks = IIf(IsNull(RsTxn_ITVendorClaim!ApprovalRemarks), "", RsTxn_ITVendorClaim!ApprovalRemarks)

If IsNull(RsTxn_ITVendorClaim!ApprovalDate) = False Then
   TxtApprovalDate.Visible = True
   Label27.Visible = True
   TxtApprovalDate.Value = RsTxn_ITVendorClaim!ApprovalDate
   TxtApprovalDate.Enabled = False
End If

TxtG_UID = IIf(IsNull(RsTxn_ITVendorClaim!G_UID), "", RsTxn_ITVendorClaim!G_UID)

TxtRemark.Locked = True
TxtApprovalRemarks.Locked = True

TxtCreated = IIf(IsNull(RsTxn_ITVendorClaim!CreatedBy), "", RsTxn_ITVendorClaim!CreatedBy) & " :- " & IIf(IsNull(RsTxn_ITVendorClaim!CreatedDate), "", RsTxn_ITVendorClaim!CreatedDate)
TxtUpdated = IIf(IsNull(RsTxn_ITVendorClaim!UpdatedBy), "", RsTxn_ITVendorClaim!UpdatedBy) & " :- " & IIf(IsNull(RsTxn_ITVendorClaim!UpdatedDate), "", RsTxn_ITVendorClaim!UpdatedDate)
TxtUpdatedByFAG = IIf(IsNull(RsTxn_ITVendorClaim!UpdatedByFAG), "", RsTxn_ITVendorClaim!UpdatedByFAG) & " :- " & IIf(IsNull(RsTxn_ITVendorClaim!UpdatedDateFAG), "", RsTxn_ITVendorClaim!UpdatedDateFAG)

Call InDataClaimDetails
If RsTxn_ITVendorClaim.RecordCount = 1 Then
   Call GButtonLock_1(Me, True)
Else
   Call GButtonLock(Me, True)
End If

SSTab1.Tab = 0

End Sub
Private Sub InDataClaimDetails()

Call Grid_Head
Call TableTxn_ITVendorClaimDetails(" Where ClaimID= " & TxtClaimID & " ") 'AND SM_Prefix = '" & TxtPreFix & "'")
If RsTxn_ITVendorClaimDetails.RecordCount > 0 Then
    MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + RsTxn_ITVendorClaimDetails.RecordCount
    For I = 2 To MSHFlexGrid2.Rows - 2
        MSHFlexGrid2.TextMatrix(I, 0) = RsTxn_ITVendorClaimDetails!ClaimDetailID
        MSHFlexGrid2.TextMatrix(I, 1) = GetLocationName(RsTxn_ITVendorClaimDetails!LocationID)
        MSHFlexGrid2.TextMatrix(I, 2) = GetExpenseType(RsTxn_ITVendorClaimDetails!ExpenseTypeID)
        MSHFlexGrid2.TextMatrix(I, 3) = RsTxn_ITVendorClaimDetails!Particulars
        MSHFlexGrid2.TextMatrix(I, 4) = Val(RsTxn_ITVendorClaimDetails!Amount)
        MSHFlexGrid2.TextMatrix(I, 5) = RsTxn_ITVendorClaimDetails!ExpenseTypeID
        MSHFlexGrid2.TextMatrix(I, 6) = RsTxn_ITVendorClaimDetails!LocationID
        MSHFlexGrid2.TextMatrix(0, 4) = Val(MSHFlexGrid2.TextMatrix(0, 4)) + Val(RsTxn_ITVendorClaimDetails!Amount)
        MSHFlexGrid2.TextMatrix(I, 7) = IIf(IsNull(RsTxn_ITVendorClaimDetails!InvoiceDate), "", RsTxn_ITVendorClaimDetails!InvoiceDate)
        If IsNull(RsTxn_ITVendorClaimDetails!AssetsID) = False Then
           MSHFlexGrid2.TextMatrix(I, 8) = GetAssetType(RsTxn_ITVendorClaimDetails!AssetsID)
           MSHFlexGrid2.TextMatrix(I, 9) = RsTxn_ITVendorClaimDetails!AssetsID
        Else
           MSHFlexGrid2.TextMatrix(I, 8) = ""
           MSHFlexGrid2.TextMatrix(I, 9) = "" 'RsTxn_ITVendorClaimDetails!AssetsID
        End If
        MSHFlexGrid2.TextMatrix(I, 10) = IIf(IsNull(RsTxn_ITVendorClaimDetails!InvoiceNo), "", RsTxn_ITVendorClaimDetails!InvoiceNo)
        MSHFlexGrid2.TextMatrix(I, 11) = IIf(IsNull(RsTxn_ITVendorClaimDetails!Quantity), "", RsTxn_ITVendorClaimDetails!Quantity)
        RsTxn_ITVendorClaimDetails.MoveNext
    Next
    'TxtClaimAmount = Val(MSHFlexGrid2.TextMatrix(0, 4))
End If
End Sub
Private Sub Grid_Head1()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.FixedRows = 1
MSHFlexGrid1.Cols = 13

MSHFlexGrid1.TextMatrix(0, 0) = "Claim ID"
MSHFlexGrid1.TextMatrix(0, 1) = "Claim Date"
MSHFlexGrid1.TextMatrix(0, 2) = "Expense From"
MSHFlexGrid1.TextMatrix(0, 3) = "Expense To"
MSHFlexGrid1.TextMatrix(0, 4) = "Vendor"
MSHFlexGrid1.TextMatrix(0, 5) = "Claim Amount"
MSHFlexGrid1.TextMatrix(0, 6) = "Remarks"
MSHFlexGrid1.TextMatrix(0, 7) = "Status"
MSHFlexGrid1.TextMatrix(0, 8) = "G_UID"
MSHFlexGrid1.TextMatrix(0, 9) = "FAG Remarks"
MSHFlexGrid1.TextMatrix(0, 10) = "Booked in NAV Date"
MSHFlexGrid1.TextMatrix(0, 11) = "Claim Forwarded To"
MSHFlexGrid1.TextMatrix(0, 12) = "Batch No"


MSHFlexGrid1.ColWidth(0) = 1000
MSHFlexGrid1.ColWidth(1) = 1000
MSHFlexGrid1.ColWidth(2) = 1000
MSHFlexGrid1.ColWidth(3) = 1100
MSHFlexGrid1.ColWidth(4) = 1100
MSHFlexGrid1.ColWidth(5) = 1200
MSHFlexGrid1.ColWidth(6) = 3200
MSHFlexGrid1.ColWidth(7) = 1400
MSHFlexGrid1.ColWidth(8) = 0
MSHFlexGrid1.ColWidth(9) = 2600
MSHFlexGrid1.ColWidth(10) = 1400
MSHFlexGrid1.ColWidth(11) = 1800
MSHFlexGrid1.ColWidth(12) = 1200
   
   
End Sub
Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Txn_ITVendorClaim.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Txn_ITVendorClaim.xls")
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
   MsgBox "Invalid Selection!!!"
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

Private Sub CmbVendorID_GotFocus()
Call GetVendorDataonGotFocus

End Sub
Private Sub GetVendorDataonGotFocus()
tmp = CmbVendorID.Text
Call TableMst_ITVendor '("Where Flag = 'Y'")
CmbVendorID.Clear
If RsMst_ITVendor.RecordCount = 0 Then
   MsgBox "No Vendor Found!!!"
   CmbVendorID.SetFocus
   'RsMst_ITVendor.MoveFirst
   Exit Sub
End If
For I = 1 To RsMst_ITVendor.RecordCount
    CmbVendorID.AddItem RsMst_ITVendor!VendorName ' & "~" & RsMst_ITVendor!VendorNo
    RsMst_ITVendor.MoveNext
Next

CmbVendorID.Text = tmp

End Sub

Private Sub GetSearchVendorDataOnLostFocus()

If CmbSVendor.Text = "" Then
   mSVendorID = Null
   Exit Sub
End If

RsMst_ITVendor.MoveFirst
RsMst_ITVendor.Find "VendorName = '" & CmbSVendor.Text & "'"

If RsMst_ITVendor.EOF Then
   MsgBox "Invalid Selection "
   CmbSVendor.SetFocus
   Exit Sub
End If
mSVendorID = RsMst_ITVendor!VendorID

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

If UBound(strSplitString) <> 1 Then
   MsgBox "Invalid Selection!!!"
   CmbSForwardTo.SetFocus
   Exit Sub
End If

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


Private Sub CmbAssetType_GotFocus()

tmp = CmbAssetType.Text
Call TableMst_AssetType("")
CmbAssetType.Clear
If RsMst_AssetType.RecordCount = 0 Then
   MsgBox "No Asset found"
   CmbAssetType.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_AssetType.RecordCount
    CmbAssetType.AddItem RsMst_AssetType!AssetType
    RsMst_AssetType.MoveNext
Next
CmbAssetType.Text = tmp

End Sub

Private Sub CmbAssetType_LostFocus()

If CmbAssetType.Text = "" Then
   mAssetTypeId = Null
   Exit Sub
End If

RsMst_AssetType.MoveFirst
RsMst_AssetType.Find "AssetType = '" & CmbAssetType.Text & "'"

If RsMst_AssetType.EOF Then
   MsgBox "Invalid Asset selection "
   CmbAssetType.SetFocus
   Exit Sub
End If
mAssetTypeId = RsMst_AssetType!AssetsID

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


