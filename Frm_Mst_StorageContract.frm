VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmMst_StorageContract 
   Caption         =   "Storage Reservation Master"
   ClientHeight    =   8100
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10830
   MDIChild        =   -1  'True
   ScaleHeight     =   8100
   ScaleWidth      =   10830
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   1725
      Left            =   2760
      TabIndex        =   0
      Top             =   1440
      Visible         =   0   'False
      Width           =   13800
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
         ItemData        =   "Frm_Mst_StorageContract.frx":0000
         Left            =   3600
         List            =   "Frm_Mst_StorageContract.frx":000A
         Sorted          =   -1  'True
         TabIndex        =   65
         Top             =   720
         Width           =   1665
      End
      Begin VB.TextBox TxtSContractID 
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
         Height          =   315
         Left            =   1590
         TabIndex        =   64
         Top             =   720
         Width           =   1395
      End
      Begin VB.ComboBox CmbSClient 
         Height          =   315
         Left            =   5400
         Sorted          =   -1  'True
         TabIndex        =   61
         Top             =   330
         Width           =   3525
      End
      Begin VB.ComboBox CmbSLocation 
         Height          =   315
         Left            =   3120
         Sorted          =   -1  'True
         TabIndex        =   60
         Top             =   330
         Width           =   2205
      End
      Begin VB.ComboBox CmbSCommodity 
         Height          =   315
         Left            =   9000
         Sorted          =   -1  'True
         TabIndex        =   63
         Top             =   330
         Width           =   2325
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "Clear"
         Height          =   350
         Left            =   13680
         TabIndex        =   56
         Top             =   720
         Width           =   1000
      End
      Begin VB.CommandButton CmdGo 
         Caption         =   "Go"
         Height          =   350
         Left            =   11520
         TabIndex        =   69
         Top             =   720
         Width           =   1000
      End
      Begin VB.TextBox TxtSearchContract 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6480
         TabIndex        =   67
         Top             =   720
         Width           =   4785
      End
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   350
         Left            =   12600
         TabIndex        =   26
         Top             =   720
         Width           =   1000
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1455
         Left            =   75
         TabIndex        =   18
         Top             =   1320
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   2566
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
         Height          =   300
         Left            =   1590
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   330
         Width           =   1410
         _ExtentX        =   2487
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
         Format          =   46530561
         CurrentDate     =   36494
      End
      Begin MSComCtl2.DTPicker SFrom 
         Height          =   300
         Left            =   45
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   330
         Width           =   1410
         _ExtentX        =   2487
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
         Format          =   46530561
         CurrentDate     =   36494
      End
      Begin VB.Label Label21 
         Caption         =   "Status"
         Height          =   165
         Left            =   3120
         TabIndex        =   117
         Top             =   765
         Width           =   690
      End
      Begin VB.Label Label20 
         Caption         =   "Reservation ID"
         Height          =   225
         Left            =   75
         TabIndex        =   116
         Top             =   765
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Client"
         Height          =   225
         Left            =   6922
         TabIndex        =   72
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label42 
         Caption         =   "Location"
         Height          =   225
         Left            =   3862
         TabIndex        =   71
         Top             =   120
         Width           =   720
      End
      Begin VB.Label Label45 
         Caption         =   "To Date"
         Height          =   255
         Left            =   1860
         TabIndex        =   70
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label44 
         Caption         =   "From Date"
         Height          =   255
         Left            =   300
         TabIndex        =   68
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label39 
         Caption         =   "Commodity"
         Height          =   210
         Left            =   9735
         TabIndex        =   66
         Top             =   120
         Width           =   960
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
         Left            =   11520
         TabIndex        =   62
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label15 
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
         Left            =   13995
         TabIndex        =   58
         Top             =   375
         Width           =   75
      End
      Begin VB.Label Label1 
         Caption         =   "Description"
         Height          =   195
         Left            =   5520
         TabIndex        =   52
         Top             =   810
         Width           =   840
      End
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   3540
      Left            =   105
      TabIndex        =   75
      Top             =   4170
      Width           =   13815
      _ExtentX        =   24368
      _ExtentY        =   6244
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
      TabCaption(0)   =   "Billing Cycle Detail"
      TabPicture(0)   =   "Frm_Mst_StorageContract.frx":0021
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "MSHFlexGrid2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Agreement Attachment Details"
      TabPicture(1)   =   "Frm_Mst_StorageContract.frx":003D
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "CmdGetFile"
      Tab(1).Control(1)=   "CmdAttachment"
      Tab(1).Control(2)=   "CmdRemoveAttachment"
      Tab(1).Control(3)=   "TxtAttachmentRemarks"
      Tab(1).Control(4)=   "MSHFlexGrid3"
      Tab(1).Control(5)=   "lblAttachment"
      Tab(1).Control(6)=   "Label16"
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "Log History"
      TabPicture(2)   =   "Frm_Mst_StorageContract.frx":0059
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label10"
      Tab(2).Control(1)=   "Label9"
      Tab(2).Control(2)=   "TxtCreated"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "TxtUpdated"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).ControlCount=   4
      Begin VB.CommandButton CmdGetFile 
         Caption         =   "Get File"
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
         Left            =   -67560
         Style           =   1  'Graphical
         TabIndex        =   96
         Top             =   1290
         Width           =   2070
      End
      Begin VB.CommandButton CmdAttachment 
         Caption         =   "Add Attachment"
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
         Left            =   -65490
         Style           =   1  'Graphical
         TabIndex        =   95
         Top             =   1290
         Width           =   2070
      End
      Begin VB.CommandButton CmdRemoveAttachment 
         Caption         =   "Remove Attachment"
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
         Left            =   -63420
         Style           =   1  'Graphical
         TabIndex        =   94
         Top             =   1290
         Width           =   2070
      End
      Begin VB.TextBox TxtAttachmentRemarks 
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
         Left            =   -68880
         MaxLength       =   5000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   93
         Top             =   645
         Width           =   7515
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
         Left            =   -73200
         Locked          =   -1  'True
         TabIndex        =   90
         TabStop         =   0   'False
         Top             =   1155
         Width           =   7455
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
         Left            =   -73200
         Locked          =   -1  'True
         TabIndex        =   89
         TabStop         =   0   'False
         Top             =   480
         Width           =   7455
      End
      Begin VB.Frame Frame2 
         Height          =   1245
         Left            =   120
         TabIndex        =   76
         Top             =   315
         Width           =   13620
         Begin VB.CommandButton CmdExcelBilling 
            Caption         =   "Excel "
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
            Left            =   12300
            TabIndex        =   115
            Top             =   810
            Width           =   1200
         End
         Begin VB.TextBox TxtResQtyforDailyMIS 
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
            Left            =   1875
            TabIndex        =   111
            Top             =   810
            Visible         =   0   'False
            Width           =   2220
         End
         Begin VB.TextBox TxtBookedQty 
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
            Left            =   9510
            TabIndex        =   14
            Top             =   390
            Width           =   1755
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
            Height          =   300
            Left            =   8925
            TabIndex        =   83
            Top             =   420
            Width           =   375
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
            ItemData        =   "Frm_Mst_StorageContract.frx":0075
            Left            =   3375
            List            =   "Frm_Mst_StorageContract.frx":0077
            Sorted          =   -1  'True
            TabIndex        =   13
            Top             =   390
            Width           =   5505
         End
         Begin VB.TextBox TxtBookedRate 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   11370
            TabIndex        =   15
            Top             =   390
            Width           =   2040
         End
         Begin VB.CommandButton CmdAddList 
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
            Left            =   10890
            TabIndex        =   16
            Top             =   810
            Width           =   1410
         End
         Begin VB.CommandButton CmdDeleteList 
            Caption         =   "Delete from List"
            Height          =   375
            Left            =   9480
            TabIndex        =   77
            Top             =   810
            Width           =   1410
         End
         Begin MSComCtl2.DTPicker TxtDFromDate 
            Height          =   360
            Left            =   240
            TabIndex        =   86
            Top             =   390
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   635
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
            Format          =   46530561
            CurrentDate     =   36494
         End
         Begin MSComCtl2.DTPicker TxtDToDate 
            Height          =   360
            Left            =   1875
            TabIndex        =   12
            Top             =   390
            Width           =   1470
            _ExtentX        =   2593
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
            Format          =   46530561
            CurrentDate     =   36494
         End
         Begin VB.Label LblResQtyforMIS 
            Caption         =   "Reserved Qty for Daily MIS"
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
            Left            =   240
            TabIndex        =   112
            Top             =   773
            Visible         =   0   'False
            Width           =   1590
         End
         Begin VB.Label Label6 
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
            Height          =   240
            Left            =   503
            TabIndex        =   88
            Top             =   120
            Width           =   1035
         End
         Begin VB.Label Label5 
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
            Height          =   270
            Left            =   2160
            TabIndex        =   87
            Top             =   120
            Width           =   825
         End
         Begin VB.Label LblBookedQty 
            Caption         =   "Reserved Qty"
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
            Left            =   9780
            TabIndex        =   85
            Top             =   120
            Width           =   1335
         End
         Begin VB.Label Label8 
            Caption         =   "Billing Cycle / Peak"
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
            Left            =   5550
            TabIndex        =   84
            Top             =   120
            Width           =   1995
         End
         Begin VB.Label LblBookedRate 
            Caption         =   "Billing Rate"
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
            Left            =   11850
            TabIndex        =   82
            Top             =   120
            Width           =   1080
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
         Height          =   1815
         Left            =   120
         TabIndex        =   78
         Top             =   1620
         Width           =   13605
         _ExtentX        =   23998
         _ExtentY        =   3201
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid4 
         Height          =   3135
         Left            =   -74880
         TabIndex        =   79
         Top             =   1350
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   5530
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
         Height          =   1710
         Left            =   -74880
         TabIndex        =   97
         Top             =   1710
         Width           =   13545
         _ExtentX        =   23892
         _ExtentY        =   3016
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
      Begin VB.Label lblAttachment 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Path of  file selected for attachment"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1140
         Left            =   -74880
         TabIndex        =   99
         Top             =   390
         Width           =   5880
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Attachment Remarks"
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
         Left            =   -67680
         TabIndex        =   98
         Top             =   360
         Width           =   2145
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
         Left            =   -74760
         TabIndex        =   92
         Top             =   1215
         Width           =   1695
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
         Height          =   255
         Left            =   -74760
         TabIndex        =   91
         Top             =   540
         Width           =   1695
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
         Left            =   -74880
         TabIndex        =   81
         Top             =   780
         Width           =   855
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
         Left            =   -74880
         TabIndex        =   80
         Top             =   1245
         Width           =   885
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4140
      Left            =   75
      TabIndex        =   27
      Top             =   0
      Width           =   13845
      Begin VB.CheckBox ChkStockManagement 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   13215
         TabIndex        =   120
         Top             =   3638
         Width           =   225
      End
      Begin VB.CheckBox ChkBlockDeposit 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   10095
         TabIndex        =   118
         Top             =   3638
         Width           =   225
      End
      Begin VB.TextBox TxtPreviousBkdQty 
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
         Left            =   6540
         Locked          =   -1  'True
         TabIndex        =   113
         TabStop         =   0   'False
         Top             =   3645
         Width           =   1335
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
         ItemData        =   "Frm_Mst_StorageContract.frx":0079
         Left            =   1575
         List            =   "Frm_Mst_StorageContract.frx":0083
         Sorted          =   -1  'True
         TabIndex        =   108
         Top             =   3645
         Width           =   2625
      End
      Begin VB.ComboBox CmbPeakFrom 
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
         ItemData        =   "Frm_Mst_StorageContract.frx":009A
         Left            =   11040
         List            =   "Frm_Mst_StorageContract.frx":00A4
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   960
         Width           =   2625
      End
      Begin VB.ComboBox CmbMSC 
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
         TabIndex        =   10
         Top             =   2280
         Width           =   12120
      End
      Begin VB.TextBox TxtInvGenUpTo 
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
         Left            =   11895
         Locked          =   -1  'True
         TabIndex        =   102
         TabStop         =   0   'False
         Top             =   165
         Width           =   1755
      End
      Begin VB.CheckBox ChkSCMFlag 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   8895
         TabIndex        =   100
         Top             =   158
         Width           =   225
      End
      Begin VB.TextBox TxtContractName 
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
         MaxLength       =   150
         TabIndex        =   2
         Top             =   570
         Width           =   12105
      End
      Begin VB.ComboBox CmbBookedAsPer 
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
         ItemData        =   "Frm_Mst_StorageContract.frx":00D2
         Left            =   8400
         List            =   "Frm_Mst_StorageContract.frx":00DF
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   1800
         Width           =   5250
      End
      Begin VB.CommandButton CmdClientID 
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
         Left            =   9120
         TabIndex        =   51
         Top             =   990
         Width           =   375
      End
      Begin VB.CommandButton CmdCommodityID 
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
         Left            =   13305
         TabIndex        =   50
         Top             =   1395
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
         Left            =   6570
         TabIndex        =   49
         Top             =   1410
         Width           =   375
      End
      Begin VB.ComboBox CmbClientID 
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
         TabIndex        =   3
         Top             =   975
         Width           =   7515
      End
      Begin VB.TextBox TxtContractID 
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
         Left            =   1575
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   165
         Width           =   1440
      End
      Begin MSComCtl2.DTPicker ContractDate 
         Height          =   360
         Left            =   4875
         TabIndex        =   1
         Top             =   165
         Width           =   1560
         _ExtentX        =   2752
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
         Format          =   46530561
         CurrentDate     =   36494
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
         Height          =   360
         Left            =   1575
         MaxLength       =   50
         TabIndex        =   32
         Top             =   975
         Width           =   7515
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
         Left            =   1575
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   1395
         Width           =   4860
      End
      Begin VB.ComboBox CmbCommodityID 
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
         Left            =   8400
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   1395
         Width           =   4875
      End
      Begin MSComCtl2.DTPicker BookedFrom 
         Height          =   360
         Left            =   1575
         TabIndex        =   7
         Top             =   1800
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
         Format          =   46530561
         CurrentDate     =   36494
      End
      Begin VB.TextBox TxtBookedAsPer 
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
         Left            =   8400
         TabIndex        =   40
         Top             =   1800
         Width           =   5235
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
         Height          =   870
         Left            =   1575
         MaxLength       =   2000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   2700
         Width           =   12105
      End
      Begin VB.TextBox TxtFlag 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   360
         Left            =   1800
         MaxLength       =   1
         TabIndex        =   45
         Top             =   7005
         Visible         =   0   'False
         Width           =   3960
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
         Left            =   1575
         TabIndex        =   34
         Top             =   1395
         Width           =   4800
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
         Left            =   8400
         TabIndex        =   36
         Top             =   1395
         Width           =   5280
      End
      Begin MSComCtl2.DTPicker TxtExtPeriod 
         Height          =   360
         Left            =   4965
         TabIndex        =   73
         Top             =   1800
         Width           =   1470
         _ExtentX        =   2593
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
         Format          =   46530561
         CurrentDate     =   36494
      End
      Begin VB.TextBox TxtPreFix 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   360
         Left            =   11640
         Locked          =   -1  'True
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   480
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox TxtStorageCharges 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   360
         Left            =   11280
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   2640
         Visible         =   0   'False
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker BookedTo 
         Height          =   360
         Left            =   11280
         TabIndex        =   8
         Top             =   2640
         Visible         =   0   'False
         Width           =   1050
         _ExtentX        =   1852
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
         Format          =   46530561
         CurrentDate     =   36494
      End
      Begin VB.TextBox TxtMSC 
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
         MaxLength       =   50
         TabIndex        =   105
         Top             =   2280
         Width           =   12120
      End
      Begin VB.TextBox TxtPeakfrom 
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
         Left            =   11040
         Locked          =   -1  'True
         MaxLength       =   150
         TabIndex        =   107
         Top             =   960
         Width           =   2625
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
         Left            =   1575
         Locked          =   -1  'True
         MaxLength       =   150
         TabIndex        =   109
         Top             =   3645
         Width           =   2625
      End
      Begin VB.Label Label24 
         Caption         =   "Stock Managment Contract"
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
         Left            =   10680
         TabIndex        =   121
         Top             =   3705
         Width           =   2490
      End
      Begin VB.Label Label22 
         Caption         =   "Block Deposit"
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
         Left            =   8640
         TabIndex        =   119
         Top             =   3705
         Width           =   1290
      End
      Begin VB.Label Label19 
         Caption         =   "Previous Booked Qty"
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
         Left            =   4560
         TabIndex        =   114
         Top             =   3720
         Width           =   1935
      End
      Begin VB.Label Label18 
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
         Height          =   285
         Left            =   120
         TabIndex        =   110
         Top             =   3683
         Width           =   1080
      End
      Begin VB.Label Label17 
         Caption         =   "Peak from"
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
         Left            =   9600
         TabIndex        =   106
         Top             =   1005
         Width           =   1080
      End
      Begin VB.Label Label14 
         Caption         =   "Master Contract"
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
         Left            =   75
         TabIndex        =   104
         Top             =   2220
         Width           =   1260
      End
      Begin VB.Label Label11 
         Caption         =   "Invoice Generated UpTo"
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
         Left            =   9600
         TabIndex        =   103
         Top             =   225
         Width           =   2295
      End
      Begin VB.Label Label52 
         Caption         =   "SCM Reservation"
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
         Left            =   6960
         TabIndex        =   101
         Top             =   225
         Width           =   1650
      End
      Begin VB.Label Label4 
         Caption         =   "Extended Upto"
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
         Left            =   3240
         TabIndex        =   74
         Top             =   1845
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Description"
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
         Left            =   75
         TabIndex        =   53
         Top             =   630
         Width           =   1305
      End
      Begin VB.Label LblContractID 
         Caption         =   "Reservation ID"
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
         Left            =   75
         TabIndex        =   28
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label LblContractDate 
         Caption         =   "Reservation Date"
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
         Left            =   3240
         TabIndex        =   30
         Top             =   210
         Width           =   1650
      End
      Begin VB.Label LblClientID 
         Caption         =   "Client"
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
         Left            =   75
         TabIndex        =   31
         Top             =   1035
         Width           =   1380
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
         Left            =   75
         TabIndex        =   33
         Top             =   1448
         Width           =   1290
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
         Height          =   285
         Left            =   7200
         TabIndex        =   35
         Top             =   1440
         Width           =   1080
      End
      Begin VB.Label LblBookedFrom 
         Caption         =   "Booked From"
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
         Left            =   75
         TabIndex        =   37
         Top             =   1860
         Width           =   1440
      End
      Begin VB.Label LblBookedTo 
         Caption         =   "Booked To"
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
         Left            =   12600
         TabIndex        =   38
         Top             =   2640
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Label LblBookedAsPer 
         Caption         =   "Booked As Per"
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
         Left            =   7200
         TabIndex        =   39
         Top             =   1740
         Width           =   1425
      End
      Begin VB.Label LblRemarks 
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
         Height          =   225
         Left            =   75
         TabIndex        =   43
         Top             =   2783
         Width           =   960
      End
      Begin VB.Label LblFlag 
         Caption         =   "Flag"
         Enabled         =   0   'False
         Height          =   225
         Left            =   240
         TabIndex        =   44
         Top             =   7080
         Visible         =   0   'False
         Width           =   1800
      End
      Begin VB.Label Label7 
         Caption         =   "SM-Pre Fix"
         Height          =   180
         Left            =   11400
         TabIndex        =   55
         Top             =   330
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label LblStorageCharges 
         Caption         =   "Storage Charges"
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
         Left            =   11520
         TabIndex        =   41
         Top             =   2520
         Visible         =   0   'False
         Width           =   1065
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1200
      Left            =   120
      TabIndex        =   46
      Top             =   7680
      Width           =   13785
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
         Left            =   10965
         TabIndex        =   47
         Top             =   645
         Width           =   2715
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
         Left            =   10965
         TabIndex        =   48
         Top             =   240
         Width           =   2715
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
         Left            =   8250
         TabIndex        =   19
         Top             =   645
         Width           =   2715
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
         Left            =   8250
         TabIndex        =   20
         Top             =   240
         Width           =   2715
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
         Left            =   5535
         TabIndex        =   21
         Top             =   645
         Width           =   2715
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
         Left            =   5535
         TabIndex        =   22
         Top             =   240
         Width           =   2715
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
         Left            =   2820
         TabIndex        =   23
         Top             =   645
         Width           =   2715
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
         Left            =   2820
         TabIndex        =   17
         Top             =   240
         Width           =   2715
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
         Left            =   100
         TabIndex        =   24
         Top             =   645
         Width           =   2715
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
         Left            =   100
         TabIndex        =   25
         Top             =   240
         Width           =   2715
      End
   End
   Begin MSComDlg.CommonDialog cdialog 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmMst_StorageContract"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg As Variant
Dim mLocationID, mCommodityID, mClientIDRow, mBillingCycleID, mMContractIDRow As Double
Dim mDToDate, mDFromDate  As Variant
Dim mClientID As String
Dim File_Path As String
Dim mPeakFrom As String
Private Sub BookedFrom_GotFocus()
tmp = BookedFrom.Value
End Sub

Private Sub BookedFrom_LostFocus()

TxtDFromDate.Value = BookedFrom.Value
TxtDToDate.Value = TxtExtPeriod.Value
If tmp <> BookedFrom.Value Then
   If MSHFlexGrid2.Rows > 2 Then
      If MsgBox("Change in Booked from date will clear Billing details, you wan't to continue?", vbYesNo) = vbYes Then
         Call Grid_Head1
         Exit Sub
      Else
         BookedFrom.Value = CDate(tmp)
      End If
   End If
End If

If Format(Gen_BlockingDate, "yyyy-mm") >= Format(BookedFrom.Value, "yyyy-mm") Then
   MsgBox "Booked From Date allowed for next month of " & Format(Gen_BlockingDate, "MMM-yyyy") & "!!!"
   BookedFrom.SetFocus
   Exit Sub
End If

End Sub

Private Sub ChkSCMFlag_Click()
CmbClientID.Text = ""
End Sub

Private Sub CmbBookedAsPer_GotFocus()
tmp = CmbBookedAsPer.Text
End Sub

Private Sub CmbBookedAsPer_LostFocus()
If CmbBookedAsPer.Text = "" Then Exit Sub
If CmbBookedAsPer.Text = "No Of Bags" Or CmbBookedAsPer.Text = "Weight In MT" Or CmbBookedAsPer.Text = "Fixed" Then

Else
    MsgBox "Invalid selection !!!!!"
    CmbBookedAsPer.SetFocus
End If

If tmp <> CmbBookedAsPer.Text Then
   If MSHFlexGrid2.Rows > 2 Then
      If MsgBox("Change in Booked from date will clear Billing details, you wan't to continue?", vbYesNo) = vbYes Then
         Call Grid_Head1
         Exit Sub
      Else
         CmbBookedAsPer.Text = tmp
      End If
   End If
End If

''If CmbBookedAsPer.Text = "Fixed" Then
''   TxtDToDate.Enabled = False
''Else
''   TxtDToDate.Enabled = True
''End If

End Sub

Private Sub CmbPeakFrom_LostFocus()
If CmbPeakFrom.Text = "" Then Exit Sub
If CmbPeakFrom.Text = "Closing Balance" Or CmbPeakFrom.Text = "Opening and Closing Bal" Then

Else
    MsgBox "Invalid selection !!!!!"
    CmbPeakFrom.SetFocus
End If
If CmbPeakFrom.Text = "Closing Balance" Then
   mPeakFrom = "C"
ElseIf CmbPeakFrom.Text = "Opening and Closing Bal" Then
   mPeakFrom = "OC"
End If

End Sub

Private Sub CmbSClient_GotFocus()

tmp = CmbSClient.Text
Call TableMst_Client("Where ExchangeTypeID = " & GetExchangeID("Non NCDEX"))
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
If RsMst_Client.EOF Then
   MsgBox "Invalid selection "
   CmbSClient.SetFocus
   Exit Sub
End If

End Sub

Private Sub CmbSCommodity_GotFocus()
tmp = CmbSCommodity.Text
Call TableMst_Commodity
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
If RsMst_Commodity.EOF Then
   MsgBox "Invalid selection "
   CmbSCommodity.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmbSLocation_GotFocus()
tmp = CmbSLocation.Text
Call TableMst_Location
CmbSLocation.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
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

End Sub

Private Sub CmbSStatus_LostFocus()

If CmbSStatus.Text = "" Then Exit Sub

If LCase(CmbSStatus.Text) <> "active" And LCase(CmbSStatus.Text) <> "de-active" Then
   MsgBox "Invalid Selection!!!"
   CmbSStatus.SetFocus
   Exit Sub
End If

End Sub

Private Sub CmbStatus_LostFocus()
If CmbStatus.Text = "" Then Exit Sub

If LCase(CmbStatus.Text) <> "active" And LCase(CmbStatus.Text) <> "de-active" Then
   MsgBox "Invalid Selection!!!"
   CmbStatus.SetFocus
   Exit Sub
End If


End Sub

Private Sub CmdAddList_Click()
Dim mInvDate As Variant
If TxtDToDate.Value < TxtDFromDate.Value Then
   MsgBox "To Date Should be Greater than or equal to From Date."
   TxtDToDate.SetFocus
   Exit Sub
End If
If TxtDToDate.Value > TxtExtPeriod.Value Then
   MsgBox "To Date Should be less than or equal to Above To Date."
   TxtDToDate.SetFocus
   Exit Sub
End If
If CmbBillingCycle.Text = "" Then
   MsgBox "Select Billing Cycle."
   CmbBillingCycle.SetFocus
   Exit Sub
End If

If TxtResQtyforDailyMIS.Visible = True Then
   If Val(TxtBookedQty.Text) = 0 Then
      If Val(TxtResQtyforDailyMIS) > 0 Then
         MsgBox "Reservation Qty for Daily MIS should less than or equals to Actual Res Qty!!!"
         TxtResQtyforDailyMIS.SetFocus
         Exit Sub
      End If
   End If
End If

If CmdAddList.Caption = "Update List" Then
   If Edit_Flg = "E" Then
      mInvDate = CheckReservationInvDate()
      If mInvDate <> "NA" Then
         If TxtDToDate.Value < CDate(mInvDate) Then
            MsgBox "Reservation Invoice is generate till " & Format(mInvDate, "dd-MMM-yyyy") & ", you can't update rates below than this date!!  "
            TxtDToDate.Value = CDate(mDToDate)
            TxtDToDate.SetFocus
            Exit Sub
         End If
      End If
   End If

   r = MSHFlexGrid2.RowSel
   If CDate(mDFromDate) <> TxtDFromDate.Value Or CDate(mDToDate) <> TxtDToDate.Value Then
      Call RemoveBillingDetailRows(r, "U")
   End If
   
Else
   If EndBillingDetails = True Then
      MsgBox "Billing Detail Entry Completed."
      Call ClearBillingDetailsFrame
      Exit Sub
   End If
End If

If CmdAddList.Caption = "Update List" Then
   r = MSHFlexGrid2.RowSel
Else
   r = MSHFlexGrid2.Rows - 1
   MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
End If
MSHFlexGrid2.TextMatrix(r, 0) = TxtDFromDate.Value
MSHFlexGrid2.TextMatrix(r, 1) = TxtDToDate.Value
MSHFlexGrid2.TextMatrix(r, 2) = CmbBillingCycle.Text
MSHFlexGrid2.TextMatrix(r, 3) = TxtBookedQty.Text
MSHFlexGrid2.TextMatrix(r, 4) = TxtBookedRate.Text
MSHFlexGrid2.TextMatrix(r, 5) = mBillingCycleID
If LCase(CmbBookedAsPer.Text) <> "no of bags" Then
   MSHFlexGrid2.TextMatrix(r, 6) = Val(TxtBookedQty)
Else
   MSHFlexGrid2.TextMatrix(r, 6) = Val(TxtResQtyforDailyMIS)
End If

Call ClearBillingDetailsFrame

End Sub


Private Sub CmdAttachment_Click()
Dim r As Variant

If lblAttachment.Caption = "File Attachment" Then
   MsgBox "Blank Attachment not allowed!!!"
   Exit Sub
End If

If TxtAttachmentRemarks = "" Then
   MsgBox "Blank Remarks not allowed!!!"
   TxtAttachmentRemarks.SetFocus
   Exit Sub
End If

If CmdAttachment.Caption = "Update Attachment" Then
   r = MSHFlexGrid3.RowSel
Else
   r = MSHFlexGrid3.Rows - 1
   MSHFlexGrid3.Rows = MSHFlexGrid3.Rows + 1
   MSHFlexGrid3.TextMatrix(r, 0) = ""
End If

MSHFlexGrid3.TextMatrix(r, 1) = GetFileFromPath(lblAttachment.Caption)
MSHFlexGrid3.TextMatrix(r, 2) = TxtAttachmentRemarks
MSHFlexGrid3.TextMatrix(r, 3) = lblAttachment.Caption
Call ClearFrame

End Sub

Private Sub CmdClear_Click()
SFrom.Value = Date - 365
STo.Value = Date
CmbSLocation.Text = ""
CmbSClient.Text = ""
CmbSCommodity.Text = ""
TxtSearchContract = ""
Label15.Caption = ""
CmbSStatus.Text = ""


End Sub

Private Sub CmdDeleteList_Click()
I = MSHFlexGrid2.RowSel
MSHFlexGrid2.Rows = I
MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
If I > 1 Then
   TxtDFromDate.Value = CDate(MSHFlexGrid2.TextMatrix(I - 1, 0))
   TxtDToDate.Value = CDate(MSHFlexGrid2.TextMatrix(I - 1, 1))
   Call ClearBillingDetailsFrame
Else
   Call ClearBillingDetailsFrame
   TxtDFromDate.Value = BookedFrom.Value
   TxtDToDate.Value = TxtExtPeriod.Value
End If

End Sub

Private Sub CmdExcelBilling_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "StorageContract_BillingDetails.xls")
Call Xls_Detail_Rpt1
End Sub
Public Sub Xls_Detail_Rpt1()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Integer
Dim mCol As Integer
MsgBox ("Wait till next message")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "StorageContract_BillingDetails.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid2.Rows - 1
    mRow = I + 1 ' + 6
    For C = 0 To MSHFlexGrid2.Cols - 3
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

Private Sub CmdGo_Click()
Dim wc As String
Call Grid_Head

If STo.Value < SFrom.Value Then
   MsgBox "To date must be always greater than from date!!!"
   STo.SetFocus
   Exit Sub
End If

tmp = "Select MSC.ContractID,MSC.ContractDate,MSC.ContractName,MC.ClientName,ML.LocationName,MComm.Commodity,MSC.BookedFrom,MSC.ExtendedPeriod,MSC.BookedAsPer, " & _
      " MSC.Remarks, MSC.Flag,MSC.BookedQty,MSC.SCMFLag,MSC.BlockDeposit,MSC.StockManagmentContract " & _
      " From Mst_StorageContract MSC " & _
      " Inner Join Mst_Location ML On MSC.LocationID = ML.LocationID " & _
      " Inner Join Mst_Client MC On MSC.ClientID = MC.ClientID " & _
      " Left Join Mst_Commodity MComm On MSC.CommodityID = MComm.CommodityID"

wc = ""
If Gen_DBType = "MDB" Then
   wc = " Where (MSC.ContractDate Between #" & Format(SFrom, Gen_DatFormat) & "# And #" & Format(STo, Gen_DatFormat) & "#) "
Else
   wc = " Where (MSC.ContractDate Between '" & Format(SFrom, Gen_DatFormat) & "' And '" & Format(STo, Gen_DatFormat) & "') "
End If

If CmbSLocation.Text <> "" Then
   wc = wc & " And ML.LocationName = '" & CmbSLocation.Text & "'"
End If

If CmbSClient.Text <> "" Then
   wc = wc & " And MC.ClientName = '" & CmbSClient.Text & "'"
End If

If CmbSCommodity.Text <> "" Then
   wc = wc & " And MComm.Commodity = '" & CmbSCommodity.Text & "'"
End If

If TxtSearchContract <> "" Then
   wc = wc & " And MSC.ContractName Like '%" & TxtSearchContract & "%'"
End If

If TxtSContractID <> "" Then
   wc = wc & " And MSC.ContractID = " & Val(TxtSContractID) & ""
End If

If CmbSStatus.Text <> "" Then
   wc = wc & " And MSC.Flag = '" & Left(CmbSStatus.Text, 1) & "'"
End If



tmp1 = " Order By MSC.ContractID"
tmp = tmp & wc & tmp1

Call TmpTable



MSHFlexGrid1.Rows = TmpRs.RecordCount + 2

For I = 1 To TmpRs.RecordCount
    For C = 0 To TmpRs.Fields.Count - 1
        MSHFlexGrid1.TextMatrix(I, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
    Next
    MSHFlexGrid1.TextMatrix(I, 1) = Format(MSHFlexGrid1.TextMatrix(I, 1), "dd-mmm-yyyy")
    MSHFlexGrid1.TextMatrix(I, 6) = Format(MSHFlexGrid1.TextMatrix(I, 6), "dd-mmm-yyyy")
    MSHFlexGrid1.TextMatrix(I, 7) = Format(MSHFlexGrid1.TextMatrix(I, 7), "dd-mmm-yyyy")
    If LCase(MSHFlexGrid1.TextMatrix(I, 10)) = "a" Then
       MSHFlexGrid1.TextMatrix(I, 10) = "Active"
    ElseIf LCase(MSHFlexGrid1.TextMatrix(I, 10)) = "d" Then
       MSHFlexGrid1.TextMatrix(I, 10) = "De-Active"
    End If
    TmpRs.MoveNext
Next

Label15.Caption = TmpRs.RecordCount
If MSHFlexGrid1.Rows > 2 Then
   MSHFlexGrid1.FixedRows = 1
End If

End Sub


Private Sub Form_Load()
'Area In SqFT
'No Of Bags

Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Call GButtonLockAD(Me)
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If
Call Grid_Head
Call CmdClear_Click

Call TableMst_StorageContract
If RsMst_StorageContract.RecordCount = 0 Then
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
Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Mst_StorageContract.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Mst_StorageContract.xls")
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



Private Sub MSHFlexGrid2_Click()
If SaveCmd.Enabled = False Then Exit Sub
Dim r As Variant
r = MSHFlexGrid2.RowSel
If r < 1 Or r = MSHFlexGrid2.Rows - 1 Then
   CmbBillingCycle.Text = ""
   mBillingCycleID = 0
   TxtBookedQty = ""
   TxtBookedRate = ""
   CmdAddList.Caption = "Add to List"
   Exit Sub
End If
If MSHFlexGrid2.TextMatrix(r, 1) = "" Then
   Exit Sub
End If
TxtDFromDate.Value = MSHFlexGrid2.TextMatrix(r, 0)
TxtDToDate.Value = MSHFlexGrid2.TextMatrix(r, 1)
mDFromDate = TxtDFromDate.Value
mDToDate = TxtDToDate.Value

CmbBillingCycle.Text = MSHFlexGrid2.TextMatrix(r, 2)
TxtBookedQty.Text = MSHFlexGrid2.TextMatrix(r, 3)
TxtBookedRate.Text = MSHFlexGrid2.TextMatrix(r, 4)
mBillingCycleID = Val(MSHFlexGrid2.TextMatrix(r, 5))
TxtResQtyforDailyMIS.Text = Val(MSHFlexGrid2.TextMatrix(r, 6))

If LCase(CmbBookedAsPer.Text) = "no of bags" Then
   LblResQtyforMIS.Visible = True
   TxtResQtyforDailyMIS.Visible = True
End If

CmbBillingCycle.Enabled = True
TxtBookedQty.Enabled = True
TxtBookedRate.Enabled = True
'TxtDFromDate.Enabled = True
TxtResQtyforDailyMIS.Enabled = True

If Edit_Flg = "E" Then
   mInvDate = CheckReservationInvDate()
   If mInvDate <> "NA" Then
      If CDate(mInvDate) >= TxtDFromDate.Value And CDate(mInvDate) <= TxtDToDate.Value Then
         CmbBillingCycle.Enabled = False
         TxtBookedQty.Enabled = False
         TxtBookedRate.Enabled = False
         'TxtDFromDate.Enabled = False
         TxtResQtyforDailyMIS.Enabled = False
      End If
   End If
End If

CmdAddList.Caption = "Update List"
CmdDeleteList.Enabled = True
TxtDToDate.SetFocus
End Sub

Private Sub MSHFlexGrid3_Click()
Dim ans, r As Variant
Dim ID As Integer
r = MSHFlexGrid3.RowSel
If r < 1 Then Exit Sub
If MSHFlexGrid3.TextMatrix(r, 1) = "" Then Exit Sub

If SaveCmd.Enabled = False Then

   ans = MsgBox("Do You want to view this document ", vbYesNo)
   If ans = vbYes Then
      If Right(App.Path, 1) = "\" Then
         Shell "Explorer.exe " & App.Path & "Attachment", vbMaximizedFocus
      Else
         Shell "Explorer.exe " & App.Path & "\Attachment", vbMaximizedFocus
      End If
   End If
   Exit Sub
End If


lblAttachment.Caption = MSHFlexGrid3.TextMatrix(r, 3)
TxtAttachmentRemarks = MSHFlexGrid3.TextMatrix(r, 2)

CmdRemoveAttachment.Enabled = True
If MSHFlexGrid3.TextMatrix(r, 4) <> "" Then
   CmdRemoveAttachment.Enabled = False
End If
CmdAttachment.Caption = "Update Attachment"

End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_StorageContract.MoveNext
If RsMst_StorageContract.EOF Then
   RsMst_StorageContract.MoveLast
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
RsMst_StorageContract.MovePrevious
If RsMst_StorageContract.BOF Then
   RsMst_StorageContract.MoveFirst
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
RsMst_StorageContract.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsMst_StorageContract.MoveLast
Call Indata
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False
End Sub
Private Sub ExitCmd_Click()
Unload Me
End Sub
Private Sub SearchCmd_Click()
Dim x As Integer
If SearchCmd.Caption = "Cancel" Then
   If RsMst_StorageContract.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   Call Indata
   Exit Sub
End If
Frame0.Height = Me.Height - 500
Frame0.Left = Frame1.Left - 100
Frame0.Top = Frame1.Top - 100
Frame0.Width = Me.Width - 150
Frame0.Visible = True
MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 1400
RsMst_StorageContract.MoveFirst

''===========comment on 7 july 2012 for not clearing search grid=============
'MSHFlexGrid1.Clear
'MSHFlexGrid1.Cols = 3
'Call Grid_Head
'Call CmdClear_Click
''===========comment on 7 july 2012 for not clearing search grid=============

End Sub
'Private Sub ShowGrid(i As Integer)
'    For c = 0 To MSHFlexGrid1.Cols - 1
'       MSHFlexGrid1.TextMatrix(i, c) = IIf(IsNull(RsMst_StorageContract.Fields(c)), "", RsMst_StorageContract.Fields(c))
'    Next
'    MSHFlexGrid1.TextMatrix(i, 2) = GetClientName(MSHFlexGrid1.TextMatrix(i, 2))
'    MSHFlexGrid1.TextMatrix(i, 3) = GetLocationName(MSHFlexGrid1.TextMatrix(i, 3))
'    MSHFlexGrid1.TextMatrix(i, 4) = GetCommodityName(MSHFlexGrid1.TextMatrix(i, 4))
'End Sub
Public Sub Indata()

Call Grid_Head1
SSTab.Tab = 0
Frame2.Enabled = False
Call ClearFrame
Call ClearBillingDetailsFrame

TxtContractID.Locked = True
ContractDate.Enabled = False
TxtClientID.Locked = True
CmbClientID.Visible = False
CmdClientID.Visible = False
TxtClientID.Visible = True
TxtLocationID.Locked = True
CmbLocationID.Visible = False
CmdLocationID.Visible = False
TxtLocationID.Visible = True
TxtCommodityID.Locked = True
CmbCommodityID.Visible = False
CmdCommodityID.Visible = False
TxtCommodityID.Visible = True
BookedFrom.Enabled = False
BookedTo.Enabled = False
TxtExtPeriod.Enabled = False
TxtBookedAsPer.Locked = True
TxtBookedAsPer.Visible = True
CmbBookedAsPer.Visible = False
TxtBookedRate.Locked = True
TxtStorageCharges.Locked = True
TxtBookedQty.Locked = True
TxtRemarks.Locked = True
TxtFlag.Locked = True
TxtContractName.Locked = True

TxtMSC.Locked = True
CmbMSC.Visible = False
TxtMSC.Visible = True

CmbPeakFrom.Visible = False
TxtPeakfrom.Visible = True
TxtPeakfrom.Text = ""

CmbStatus.Visible = False
TxtStatus.Visible = True
TxtStatus = ""

TxtContractID = IIf(IsNull(RsMst_StorageContract!ContractID), "", RsMst_StorageContract!ContractID)
TxtContractName = IIf(IsNull(RsMst_StorageContract!ContractName), "", RsMst_StorageContract!ContractName)

ContractDate = IIf(IsNull(RsMst_StorageContract!ContractDate), "", RsMst_StorageContract!ContractDate)
mClientIDRow = IIf(IsNull(RsMst_StorageContract!ClientID), "", RsMst_StorageContract!ClientID)
TxtClientID = GetClientName(RsMst_StorageContract!ClientID)  'IIf(IsNull(RsMst_StorageContract!ClientID), "", RsMst_StorageContract!ClientID)
mLocationID = IIf(IsNull(RsMst_StorageContract!LocationID), "", RsMst_StorageContract!LocationID)
TxtLocationID = GetLocationName(RsMst_StorageContract!LocationID) '(IsNull(RsMst_StorageContract!LocationName), "", RsMst_StorageContract!LocationName)
mCommodityID = IIf(IsNull(RsMst_StorageContract!CommodityID), 0, RsMst_StorageContract!CommodityID)
TxtCommodityID = GetCommodityName(IIf(IsNull(RsMst_StorageContract!CommodityID), 0, RsMst_StorageContract!CommodityID))

TxtPreFix = GetPreFixByLocation(RsMst_StorageContract!LocationID, "P")
mPeakFrom = IIf(IsNull(RsMst_StorageContract!PeakFrom), "", RsMst_StorageContract!PeakFrom)
If mPeakFrom = "C" Then
   TxtPeakfrom.Text = "Closing Balance"
ElseIf mPeakFrom = "OC" Then
   TxtPeakfrom.Text = "Opening and Closing Bal"
Else
   TxtPeakfrom.Text = ""
End If

mStatus = IIf(IsNull(RsMst_StorageContract!Flag), "", RsMst_StorageContract!Flag)

If LCase(mStatus) = "a" Then
   TxtStatus = "Active"
ElseIf LCase(mStatus) = "d" Then
   TxtStatus = "De-Active"
End If

BookedFrom = IIf(IsNull(RsMst_StorageContract!BookedFrom), "", RsMst_StorageContract!BookedFrom)
'BookedTo = IIf(IsNull(RsMst_StorageContract!BookedTo), "", RsMst_StorageContract!BookedTo)
TxtExtPeriod.Value = IIf(IsNull(RsMst_StorageContract!ExtendedPeriod), BookedTo.Value, RsMst_StorageContract!ExtendedPeriod)
TxtBookedAsPer = IIf(IsNull(RsMst_StorageContract!BookedAsPer), "", RsMst_StorageContract!BookedAsPer)
'TxtBookedRate = IIf(IsNull(RsMst_StorageContract!BookedRate), "", RsMst_StorageContract!BookedRate)
'TxtStorageCharges = IIf(IsNull(RsMst_StorageContract!StorageCharges), "", RsMst_StorageContract!StorageCharges)
'TxtBookedQty = IIf(IsNull(RsMst_StorageContract!BookedQty), "", RsMst_StorageContract!BookedQty)
TxtPreviousBkdQty = IIf(IsNull(RsMst_StorageContract!BookedQty), "", RsMst_StorageContract!BookedQty)
TxtRemarks = IIf(IsNull(RsMst_StorageContract!Remarks), "", RsMst_StorageContract!Remarks)
TxtFlag = IIf(IsNull(RsMst_StorageContract!Flag), "", RsMst_StorageContract!Flag)

If RsMst_StorageContract!SCMFlag = False Then
   ChkSCMFlag.Value = vbUnchecked 'IIf(IsNull(RsMst_StorageContract!SCMFlag), False, RsMst_StorageContract!SCMFlag)
Else
   ChkSCMFlag.Value = vbChecked
End If
ChkSCMFlag.Enabled = False

If RsMst_StorageContract!BlockDeposit = False Then
   ChkBlockDeposit.Value = vbUnchecked 'IIf(IsNull(RsMst_StorageContract!BlockDeposit), 0, RsMst_StorageContract!BlockDeposit)
Else
   ChkBlockDeposit.Value = vbChecked
End If
ChkBlockDeposit.Enabled = False

ChkStockManagement.Value = IIf(IsNull(RsMst_StorageContract!StockManagmentContract), 0, RsMst_StorageContract!StockManagmentContract)
ChkStockManagement.Enabled = False

TxtCreated = IIf(IsNull(RsMst_StorageContract!CreatedBy), "", RsMst_StorageContract!CreatedBy) & " :- " & IIf(IsNull(RsMst_StorageContract!CreatedDate), "", RsMst_StorageContract!CreatedDate)
TxtUpdated = IIf(IsNull(RsMst_StorageContract!UpdatedBy), "", RsMst_StorageContract!UpdatedBy) & " :- " & IIf(IsNull(RsMst_StorageContract!UpdatedDate), "", RsMst_StorageContract!UpdatedDate)



TxtInvGenUpTo.Text = CheckReservationInvDate
TxtInvGenUpTo.Text = IIf(TxtInvGenUpTo.Text = "NA", TxtInvGenUpTo.Text, Format(TxtInvGenUpTo.Text, "dd-MMM-yyyy"))
If IsNull(RsMst_StorageContract!MasterStorageContractID) = False Then
   TxtMSC.Text = RsMst_StorageContract!MasterStorageContractID & "~" & ShowContractName(RsMst_StorageContract!MasterStorageContractID)
Else
   TxtMSC.Text = ""
End If
Call IndataAttachmentDetails
Call InDataBillingDetails

If RsMst_StorageContract.RecordCount = 1 Then
   Call GButtonLock_1(Me, True)
Else
   Call GButtonLock(Me, True)
End If

End Sub

Private Sub TxtContractID_LostFocus()
If TxtContractID = "" Then Exit Sub
If IsNumeric(TxtContractID) = False Then
   MsgBox "Invalid numeric format !!!!"
   TxtContractID.SetFocus
   Exit Sub
End If
TxtContractID = Val(TxtContractID)
End Sub

Private Sub TxtBookedRate_LostFocus()
If TxtBookedRate = "" Then Exit Sub
If IsNumeric(TxtBookedRate) = False Then
   MsgBox "Invalid numeric format !!!!"
   TxtBookedRate.SetFocus
   Exit Sub
End If
TxtBookedRate = Format(Val(TxtBookedRate), "#####0.00")
End Sub



Private Sub TxtExtPeriod_GotFocus()
tmp = TxtExtPeriod.Value
End Sub

Private Sub TxtExtPeriod_LostFocus()
TxtDFromDate.Value = BookedFrom.Value
TxtDToDate.Value = TxtExtPeriod.Value
'If tmp <> TxtExtPeriod.Value Then
   If MSHFlexGrid2.Rows > 2 Then
      ''''added by rakesh on 4 mar 2013
      If TxtExtPeriod.Value < CDate(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 2, 1)) Then
         MsgBox "Extendend upto Date should not less than " & Format(CDate(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 2, 1)), "dd-MMM-yyyy")
         TxtExtPeriod.SetFocus
         Exit Sub
      Else
         TxtDFromDate.Value = CDate(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 2, 0))
         TxtDToDate.Value = CDate(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 2, 1))
         Call ClearBillingDetailsFrame
      End If
      'If MsgBox("Change in Booked from date will clear Billing details, you wan't to continue?", vbYesNo) = vbYes Then
         'Call Grid_Head1
         'Exit Sub
      'Else
      '   TxtExtPeriod.Value = CDate(tmp)
      'End If
   End If
'End If

End Sub

Private Sub TxtStorageCharges_LostFocus()
If TxtStorageCharges = "" Then Exit Sub
If IsNumeric(TxtStorageCharges) = False Then
   MsgBox "Invalid numeric format !!!!"
   TxtStorageCharges.SetFocus
   Exit Sub
End If
TxtStorageCharges = Val(TxtStorageCharges)
End Sub
Private Sub TxtBookedQty_LostFocus()
If TxtBookedQty = "" Then Exit Sub
If IsNumeric(TxtBookedQty) = False Then
   MsgBox "Invalid numeric format !!!!"
   TxtBookedQty.SetFocus
   Exit Sub
End If
TxtBookedQty = Format(Val(TxtBookedQty), "#####0.000")
TxtStorageCharges = Val(TxtBookedQty) * Val(TxtBookedRate)

End Sub
Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)
Frame2.Enabled = True

TxtContractID.Locked = True
ContractDate.Enabled = True
TxtClientID.Locked = False
CmbClientID.Visible = True
CmdClientID.Visible = True
TxtClientID.Visible = False
TxtLocationID.Locked = False
CmbLocationID.Visible = True
CmdLocationID.Visible = True
TxtLocationID.Visible = False
TxtCommodityID.Locked = False
CmbCommodityID.Visible = True
CmdCommodityID.Visible = True
TxtCommodityID.Visible = False
TxtContractName.Locked = False

CmbPeakFrom.Visible = True
TxtPeakfrom.Visible = False
CmbPeakFrom.Text = ""

CmbStatus.Visible = True
TxtStatus.Visible = False
CmbStatus.Text = ""

CmbMSC.Visible = True
TxtMSC.Visible = False
CmbMSC.Text = ""

ContractName = ""
TxtExtPeriod.Enabled = True
BookedFrom.Enabled = True
BookedTo.Enabled = True
TxtBookedAsPer.Visible = False
CmbBookedAsPer.Visible = True
TxtBookedRate.Locked = False
TxtStorageCharges.Locked = False
TxtBookedQty.Locked = False
TxtRemarks.Locked = False
TxtFlag.Locked = False
mCommodityID = 0
TxtContractID = ""
TxtContractName = ""
ContractDate = Date
TxtPreviousBkdQty = ""

TxtCreated = ""
TxtUpdated = ""

TxtClientID = ""
CmbClientID.Text = ""

TxtLocationID = ""
CmbLocationID.Text = ""

TxtInvGenUpTo = ""
TxtCommodityID = ""
TxtCommodityID.Text = ""

BookedFrom = Date
BookedTo = Date
TxtExtPeriod.Value = Date
TxtBookedAsPer = ""
CmbBookedAsPer.Text = ""
TxtBookedRate = ""
TxtStorageCharges = ""
TxtBookedQty = ""
TxtRemarks = ""
TxtFlag = ""
ChkSCMFlag.Enabled = True
ChkSCMFlag.Value = 0
ChkBlockDeposit.Value = 0
ChkStockManagement.Value = 0
ChkStockManagement.Enabled = True

Call Grid_Head1
Call ClearBillingDetailsFrame

lblAttachment.Enabled = True
TxtAttachmentRemarks.Locked = False
CmdAttachment.Enabled = True
CmdGetFile.Enabled = True
CmdRemoveAttachment.Enabled = True
Call Grid_Head3
SSTab.Tab = 0

TxtDFromDate.Value = Date
TxtDToDate.Value = Date
'ContractDate.SetFocus

End Sub
Private Sub EditCmd_Click()
Edit_Flg = "E"
Call GButtonLock(Me, False)
Frame2.Enabled = True

lblAttachment.Enabled = True
TxtAttachmentRemarks.Locked = False
CmdAttachment.Enabled = True
CmdGetFile.Enabled = True
CmdRemoveAttachment.Enabled = True
SSTab.Tab = 0

TxtContractID.Locked = True
ContractDate.Enabled = True
TxtContractName.Locked = False
TxtClientID.Locked = False
CmbClientID.Visible = True
CmdClientID.Visible = True
TxtClientID.Visible = False
TxtLocationID.Locked = False
CmbLocationID.Visible = True
CmdLocationID.Visible = True
TxtLocationID.Visible = False
TxtCommodityID.Locked = False
CmbCommodityID.Visible = True
CmdCommodityID.Visible = True
TxtCommodityID.Visible = False
ChkSCMFlag.Enabled = True
ChkBlockDeposit.Enabled = True
ChkStockManagement.Enabled = True

If TxtInvGenUpTo.Text = "NA" Then
   BookedFrom.Enabled = True
Else
   BookedFrom.Enabled = False
End If
'
BookedTo.Enabled = True
TxtExtPeriod.Enabled = True
TxtBookedAsPer.Visible = False
CmbBookedAsPer.Visible = True
TxtBookedRate.Locked = False
TxtStorageCharges.Locked = True
TxtBookedQty.Locked = False
TxtRemarks.Locked = False
'TxtFlag.Locked = False

CmbMSC.Visible = True
TxtMSC.Visible = False
CmbMSC.Text = TxtMSC

CmbPeakFrom.Visible = True
TxtPeakfrom.Visible = False
CmbPeakFrom.Text = TxtPeakfrom

CmbStatus.Visible = True
TxtStatus.Visible = False
CmbStatus.Text = TxtStatus


CmbClientID.Text = TxtClientID
CmbLocationID.Text = TxtLocationID
CmbCommodityID.Text = TxtCommodityID
CmbBookedAsPer.Text = TxtBookedAsPer
TxtDFromDate.Value = BookedFrom.Value
TxtDToDate.Value = TxtExtPeriod.Value

End Sub
Private Sub DeleteCmd_Click()
Dim ans As Variant
On Error GoTo msgError
ans = MsgBox("Do you really want to DELETE this record???", vbYesNo)
If ans = vbYes Then
   tmp = "Select * From Mst_StorageContract Where MasterStorageContractID=" & TxtContractID
   Call Tmp2Table
   If TmpRs2.RecordCount > 0 Then
      MsgBox " This contract is used as Master Contract, you can't delete!!"
      Exit Sub
   End If
   Call DeleteBillingDetails
   
   RsMst_StorageContract.Delete
   RsMst_StorageContract.Update
   If RsMst_StorageContract.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_StorageContract.MoveNext
   If RsMst_StorageContract.EOF() Then
      RsMst_StorageContract.MoveLast
   End If
   Call Indata
End If
Exit Sub
msgError:
MsgBox "Child record Present "
RsMst_StorageContract.CancelUpdate
End Sub
Private Sub SaveCmd_Click()

If CmbClientID.Text = "" Then
  MsgBox "Blank Client Not Allowed !!!"
  CmbClientID.SetFocus
  Exit Sub
End If

If CmbLocationID.Text = "" Then
  MsgBox "Blank Location Not Allowed !!!"
  CmbLocationID.SetFocus
  Exit Sub
End If

'If CmbCommodityID.Text = "" Then
'  MsgBox "Blank Commodity Not Allowed!!!"
'  CmbCommodityID.SetFocus
'  Exit Sub
'End If


If CmbBookedAsPer.Text = "" Then
  MsgBox "Blank Booked As Per Not Allowed !!!"
  CmbBookedAsPer.SetFocus
  Exit Sub
End If

If CmbPeakFrom.Text = "" Then
  MsgBox "Blank Peak From Not Allowed !!!"
  CmbPeakFrom.SetFocus
  Exit Sub
End If

If CmbStatus.Text = "" Then
  MsgBox "Blank Status Not Allowed !!!"
  CmbStatus.SetFocus
  Exit Sub
End If

''If TxtBookedRate = "" Then
''  MsgBox "Blank Booked Rate Not Allowed !!!"
''  TxtBookedRate.SetFocus
''  Exit Sub
''End If
''
''If TxtBookedQty = "" Then
''  MsgBox "Blank Booked Qty Not Allowed !!!"
''  TxtBookedQty.SetFocus
''  Exit Sub
''End If


''If BookedTo.Value < BookedFrom.Value Then
''      MsgBox "Booked To Date Must Be Greater Than Booked From Date"
''      BookedTo.SetFocus
''      Exit Sub
''End If

If TxtExtPeriod.Value < BookedFrom.Value Then
   MsgBox "'Extended Period' must be greater than or equal to 'Booked From Date'"
   TxtExtPeriod.SetFocus
   Exit Sub
End If

If MSHFlexGrid2.Rows < 3 Then
   MsgBox "Billing details missing "
   SSTab.Tab = 0
   TxtDToDate.SetFocus
   Exit Sub
End If
   
If StartBillingDetails = False Then
   MsgBox "Billing Details entry not Completed , check Booked From Date in grid!!!"
   SSTab.Tab = 0
   TxtDFromDate.SetFocus
   Exit Sub
End If
   
If EndBillingDetails = False Then
   MsgBox "Billing Details entry not Completed , check extended Date in grid!!!"
   SSTab.Tab = 0
   TxtDToDate.SetFocus
   Exit Sub
End If
   
   
If Edit_Flg = "A" Then
   If RsMst_StorageContract.RecordCount > 0 Then
      RsMst_StorageContract.MoveFirst
      RsMst_StorageContract.Find "ContractName= '" & TxtContractName.Text & "'"
      If Not RsMst_StorageContract.EOF Then
        MsgBox "Duplicate Contract Discription Not Allowed !!! "
         TxtContractName.SetFocus
         Exit Sub
      End If
   End If
      RsMst_StorageContract.AddNew
      RsMst_StorageContract!ContractID = GetContractID
      RsMst_StorageContract!G_UID = GetGUID
      TxtContractID.Text = RsMst_StorageContract!ContractID
Else
    RsMst_StorageContract.MoveFirst
    RsMst_StorageContract.Find "ContractName= '" & TxtContractName.Text & "'"
    If Not RsMst_StorageContract.EOF Then
       If RsMst_StorageContract!ContractID <> TxtContractID.Text Then
          MsgBox "Duplicate Contract Discription  Not Allowed !!! "
          TxtContractName.SetFocus
          Exit Sub
       End If
    End If
    RsMst_StorageContract.MoveFirst
    RsMst_StorageContract.Find " ContractID= " & TxtContractID.Text
End If

RsMst_StorageContract!ContractID = TxtContractID.Text
RsMst_StorageContract!ContractName = TxtContractName.Text
RsMst_StorageContract!ContractDate = ContractDate.Value
RsMst_StorageContract!ClientID = mClientIDRow

RsMst_StorageContract!LocationID = mLocationID
If mCommodityID <> 0 Then
   RsMst_StorageContract!CommodityID = mCommodityID
Else
   RsMst_StorageContract!CommodityID = Null
End If

RsMst_StorageContract!PeakFrom = mPeakFrom

RsMst_StorageContract!BookedFrom = BookedFrom.Value
'RsMst_StorageContract!BookedTo = BookedTo.Value
RsMst_StorageContract!ExtendedPeriod = TxtExtPeriod.Value
RsMst_StorageContract!BookedAsPer = CmbBookedAsPer.Text
'RsMst_StorageContract!BookedRate = Val(TxtBookedRate.Text)
'RsMst_StorageContract!StorageCharges = Val(TxtStorageCharges.Text)
'RsMst_StorageContract!BookedQty = Val(TxtBookedQty.Text)
RsMst_StorageContract!SCMFlag = ChkSCMFlag.Value
RsMst_StorageContract!BlockDeposit = ChkBlockDeposit.Value
RsMst_StorageContract!StockManagmentContract = ChkStockManagement.Value
RsMst_StorageContract!Remarks = TxtRemarks.Text
RsMst_StorageContract!Flag = Left(CmbStatus.Text, 1) '"0" 'TxtFlag.Text
If CmbMSC.Text <> "" Then
   RsMst_StorageContract!MasterStorageContractID = mMContractIDRow
Else
   RsMst_StorageContract!MasterStorageContractID = Null
End If
If IsNull(RsMst_StorageContract!CreatedBy) = True Or RsMst_StorageContract!CreatedBy = "" Then
   RsMst_StorageContract!CreatedBy = Gen_UserLoginID
   RsMst_StorageContract!CreatedDate = Now
Else
   RsMst_StorageContract!UpdatedBy = Gen_UserLoginID
   RsMst_StorageContract!UpdatedDate = Now
End If


RsMst_StorageContract.Update

Call SaveBillingDetails
Call SaveAttachmentDetails


'Code added on 18 March 2013 due to usage of only new reservation master but still process of reservation generation is old

If Edit_Flg = "A" Then
   tmp = " Update  Mst_StorageContract Set BookedTo=ExtendedPeriod Where BookedTo Is null"
   Call TmpTable
    
   tmp = " Update  A set a.BookedRate=(select max(BillingRate) from Mst_SContractBillingDetail Where Mst_SContractBillingDetail.ContractID=A.ContractID) from Mst_StorageContract A Where A.BookedRate is null"
   Call TmpTable
    
   tmp = " Update  A set a.BookedQty=(select max(BillingQty) from Mst_SContractBillingDetail Where Mst_SContractBillingDetail.ContractID=A.ContractID) from Mst_StorageContract A Where A.BookedQty is null"
   Call TmpTable
    
   tmp = " Update Mst_StorageContract Set StorageCharges=BookedQty*BookedRate Where StorageCharges is null"
   Call TmpTable
End If

Call Indata

End Sub
Private Function GetContractID() As Double
Dim Retval As Double
tmp = "Select max(ContractID) from Mst_StorageContract"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetContractID = Retval
End Function
Private Sub Grid_Head()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 15 '14
MSHFlexGrid1.TextMatrix(0, 0) = "Reservation ID"
MSHFlexGrid1.TextMatrix(0, 1) = "Reservation Date"
MSHFlexGrid1.TextMatrix(0, 2) = "Description"
MSHFlexGrid1.TextMatrix(0, 3) = "Client"
MSHFlexGrid1.TextMatrix(0, 4) = "Location"
MSHFlexGrid1.TextMatrix(0, 5) = "Commodity"
MSHFlexGrid1.TextMatrix(0, 6) = "Booked From"
MSHFlexGrid1.TextMatrix(0, 7) = "Extended Period"
MSHFlexGrid1.TextMatrix(0, 8) = "Booked As Per"
MSHFlexGrid1.TextMatrix(0, 9) = "Remarks"
MSHFlexGrid1.TextMatrix(0, 10) = "Flag"
MSHFlexGrid1.TextMatrix(0, 11) = "Previous Booked Qty"
MSHFlexGrid1.TextMatrix(0, 12) = "SCM Flag"
MSHFlexGrid1.TextMatrix(0, 13) = "Block Deposit"
MSHFlexGrid1.TextMatrix(0, 14) = "Stock Managment Contract"


MSHFlexGrid1.ColWidth(0) = 1200
MSHFlexGrid1.ColWidth(1) = 1600
MSHFlexGrid1.ColWidth(2) = 2000
MSHFlexGrid1.ColWidth(3) = 1000
MSHFlexGrid1.ColWidth(4) = 1200
MSHFlexGrid1.ColWidth(5) = 1200
MSHFlexGrid1.ColWidth(6) = 1200
MSHFlexGrid1.ColWidth(7) = 1200
MSHFlexGrid1.ColWidth(8) = 1200
MSHFlexGrid1.ColWidth(9) = 1000
MSHFlexGrid1.ColWidth(10) = 1000
MSHFlexGrid1.ColWidth(11) = 1600
MSHFlexGrid1.ColWidth(12) = 1600

End Sub
Private Sub Grid_Head1()

MSHFlexGrid2.Clear
MSHFlexGrid2.Rows = 2
MSHFlexGrid2.Cols = 7
MSHFlexGrid2.TextMatrix(0, 0) = "From Date"
MSHFlexGrid2.TextMatrix(0, 1) = "To Date"
MSHFlexGrid2.TextMatrix(0, 2) = "Billing Cycle"
MSHFlexGrid2.TextMatrix(0, 3) = "Reserved Qty"
MSHFlexGrid2.TextMatrix(0, 4) = "Billing Rate"
MSHFlexGrid2.TextMatrix(0, 5) = "Billing Cycle ID"
MSHFlexGrid2.TextMatrix(0, 6) = "Reserved Qty for Daily MIS"

MSHFlexGrid2.ColWidth(0) = 1200
MSHFlexGrid2.ColWidth(1) = 1200
MSHFlexGrid2.ColWidth(2) = 2000
MSHFlexGrid2.ColWidth(3) = 2000
MSHFlexGrid2.ColWidth(4) = 2000
MSHFlexGrid2.ColWidth(5) = 0
MSHFlexGrid2.ColWidth(6) = 0

End Sub

Private Sub MSHFlexGrid1_Click()
Dim I As Variant
I = MSHFlexGrid1.RowSel
If I <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(I, 0) = "" Then Exit Sub
RsMst_StorageContract.MoveFirst
RsMst_StorageContract.Find "ContractID = " & MSHFlexGrid1.TextMatrix(I, 0)
If Not RsMst_StorageContract.EOF Then
   Call Indata
   Frame0.Visible = False
End If
End Sub
Private Sub CmdClientID_Click()
    FrmMst_Client.Show
End Sub
'----
Private Sub CmbClientId_GotFocus()
tmp = CmbClientID.Text
Call TableMst_Client("Where SCMFlag = '" & ChkSCMFlag.Value & "' And ExchangeTypeID = " & GetExchangeID("Non NCDEX") & " And Converted = 0 And (Status = 'A' Or Status is null) ")
CmbClientID.Clear
If RsMst_Client.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Client.RecordCount
    CmbClientID.AddItem RsMst_Client!ClientName
    RsMst_Client.MoveNext
Next
CmbClientID.Text = tmp
End Sub
'----
Private Sub CmbClientId_LostFocus()
If CmbClientID.Text = "" Then
   Exit Sub
End If
RsMst_Client.MoveFirst
RsMst_Client.Find "ClientName = '" & CmbClientID.Text & "'"
If RsMst_Client.EOF Then
   MsgBox "Invalid selection "
   CmbClientID.SetFocus
   Exit Sub
End If
mClientIDRow = RsMst_Client!ClientID
End Sub
'----
Private Sub CmdLocationID_Click()
    FrmMst_Location.Show
End Sub
'----
Private Sub CmbLocationID_GotFocus()
tmp = CmbLocationID.Text
Call TableMst_Location
CmbLocationID.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbLocationID.AddItem RsMst_Location!LocationName
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
If mLocationID <> RsMst_Location!LocationID Then
   mMContractIDRow = 0
   CmbMSC.Text = ""
End If
mLocationID = RsMst_Location!LocationID
TxtPreFix = GetPreFixByLocation(RsMst_Location!LocationID, "P")
End Sub
'----
Private Sub CmdCommodityID_Click()
    FrmMst_Commodity.Show
End Sub
'----
Private Sub CmbCommodityID_GotFocus()
tmp = CmbCommodityID.Text
mCommodityID = 0
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
   mCommodityID = Null
   Exit Sub
End If
RsMst_Commodity.MoveFirst
RsMst_Commodity.Find "Commodity = '" & CmbCommodityID.Text & "'"
If RsMst_Commodity.EOF Then
   MsgBox "Invalid selection "
   CmbCommodityID.SetFocus
   Exit Sub
End If
mCommodityID = RsMst_Commodity!CommodityID
End Sub
'----
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
   mBillingCycleID = 0
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

Private Sub ClearBillingDetailsFrame()
If TxtDToDate.Value <> TxtExtPeriod.Value Then
   TxtDFromDate.Value = TxtDToDate.Value + 1
   TxtDToDate.Value = TxtExtPeriod.Value
Else
   TxtDFromDate.Value = TxtDToDate.Value
End If
CmbBillingCycle.Enabled = True
TxtBookedQty.Enabled = True
TxtBookedRate.Enabled = True
'TxtDFromDate.Enabled = True
TxtDToDate.Enabled = True

CmbBillingCycle.Text = ""
TxtBookedQty = ""
TxtBookedRate = ""
mDToDate = Null
mDFromDate = Null
mBillingCycleID = 0
LblResQtyforMIS.Visible = False
TxtResQtyforDailyMIS.Visible = False
CmdAddList.Caption = "Add to List"
CmdDeleteList.Enabled = False

End Sub

Private Function EndBillingDetails() As Boolean
Dim Retval As Boolean
Retval = False
If MSHFlexGrid2.Rows > 1 Then
   For u = 1 To MSHFlexGrid2.Rows - 1
       If MSHFlexGrid2.TextMatrix(u, 1) = "" Then Exit For
       If CDate(MSHFlexGrid2.TextMatrix(u, 1)) = TxtExtPeriod.Value Then
          Retval = True
      '''''Added by Rakesh on 4 Mar, 2013=====================
       'ElseIf CDate(MSHFlexGrid2.TextMatrix(u, 1)) > TxtExtPeriod.Value Then
       '   Retval = False
       '   EndBillingDetails = Retval
       '   Exit Function
       'ElseIf TxtDFromDate.Value >= CDate(MSHFlexGrid2.TextMatrix(u, 0)) And TxtDFromDate.Value <= CDate(MSHFlexGrid2.TextMatrix(u, 1)) Then
       '   Retval = False
       '   EndBillingDetails = Retval
       '   Exit Function
       'ElseIf TxtDToDate.Value >= CDate(MSHFlexGrid2.TextMatrix(u, 0)) And TxtDToDate.Value <= CDate(MSHFlexGrid2.TextMatrix(u, 1)) Then
       '   Retval = False
       '   EndBillingDetails = Retval
       '   Exit Function
      '''''Added by Rakesh on 4 Mar, 2013=====================
       End If
   Next
End If
EndBillingDetails = Retval
End Function

Private Sub RemoveBillingDetailRows(RowNo_ As Variant, Flag_ As String)
Dim mTotRows, mTotDeleteRows As Double
mTotRows = MSHFlexGrid2.Rows - 2
If Flag_ = "U" Then
   mTotDeleteRows = mTotRows - RowNo_
ElseIf Flag_ = "D" Then
   mTotDeleteRows = mTotRows - (RowNo_ - 1)
End If
If RowNo_ > 0 Then
   MSHFlexGrid2.Rows = MSHFlexGrid2.Rows - (mTotDeleteRows + 1)
   MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
   MSHFlexGrid2.FixedRows = 1
End If
End Sub
Private Sub Grid_Head3()
With MSHFlexGrid3
    .Clear
    .Rows = 2
    .Cols = 5
    .Font.Size = 10
    .WordWrap = True
    .TextMatrix(0, 0) = "SR No"
    .TextMatrix(0, 1) = "Attachment File Name"
    .TextMatrix(0, 2) = "Remarks"
    .TextMatrix(0, 3) = "File Path"
    .TextMatrix(0, 4) = "ID"
    .ColWidth(0) = 1000
    .ColWidth(1) = 5000
    .ColWidth(2) = 8200
    .ColWidth(3) = 0
    .ColWidth(4) = 0
End With
End Sub

Private Sub CmdGetFile_Click()

With cdialog
     .InitDir = App.Path
     .DialogTitle = "Select Image"
     .FilterIndex = 1
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

Private Sub CmdRemoveAttachment_Click()
I = MSHFlexGrid3.RowSel
If I > 0 Then
    If MSHFlexGrid3.TextMatrix(I, 4) = "" And MSHFlexGrid3.TextMatrix(I, 1) <> "" Then
        For RO = I To MSHFlexGrid3.Rows - 2
            For C = 0 To MSHFlexGrid3.Cols - 1
                MSHFlexGrid3.TextMatrix(RO, C) = MSHFlexGrid3.TextMatrix(RO + 1, C)
            Next
        Next
        MSHFlexGrid3.Rows = MSHFlexGrid3.Rows - 1
    End If
End If

Call ClearFrame
End Sub
Private Sub ClearFrame()

lblAttachment.Caption = "File Attachment"
CmdAttachment.Caption = "Add Attachment"
TxtAttachmentRemarks = ""

End Sub

Private Sub SaveAttachmentDetails()

If MSHFlexGrid3.Rows <= 1 Then Exit Sub

'tmp = "Select * from Txn_StorageContractAttachment Where SContractID = " & TxtContractID & " "

Call TableTxn_SContractAttachment(" Where SContractID = " & TxtContractID & " ")

For I = 1 To MSHFlexGrid3.Rows - 1
    If MSHFlexGrid3.TextMatrix(I, 1) = "" Then Exit For
    If MSHFlexGrid3.TextMatrix(I, 4) = "" Then
       RsTxn_SContractAttachment.AddNew
 
       RsTxn_SContractAttachment!TxnID = GetSContractAttachID
       RsTxn_SContractAttachment!SContractID = TxtContractID.Text
       mFlag = SaveFileToDB(MSHFlexGrid3.TextMatrix(I, 3), RsTxn_SContractAttachment, "File")
       mFlag = GetFileFromPath(MSHFlexGrid3.TextMatrix(I, 3))
       RsTxn_SContractAttachment!FileName = "SContract " & TxtContractID & "-" & mFlag
    Else
       RsTxn_SContractAttachment.MoveFirst
       RsTxn_SContractAttachment.Find "TxnID = " & MSHFlexGrid3.TextMatrix(I, 4) & ""
    End If
    RsTxn_SContractAttachment!Remarks = MSHFlexGrid3.TextMatrix(I, 2)
    RsTxn_SContractAttachment!UpdatedDate = Now
    RsTxn_SContractAttachment!UpdatedBy = Gen_UserLoginID
    RsTxn_SContractAttachment!G_UID = GetGUID
    If IsNull(RsTxn_SContractAttachment!CreatedBy) = True Or RsTxn_SContractAttachment!CreatedBy = "" Then
       RsTxn_SContractAttachment!CreatedBy = Gen_UserLoginID
       RsTxn_SContractAttachment!CreatedDate = Now
    Else
       RsTxn_SContractAttachment!UpdatedBy = Gen_UserLoginID
       RsTxn_SContractAttachment!UpdatedDate = Now
    End If
    RsTxn_SContractAttachment.Update
Next

End Sub

Private Sub IndataAttachmentDetails()
Call Grid_Head3
Call ClearFrame

Call TableTxn_SContractAttachment(" Where SContractID = " & TxtContractID.Text & " ")

MSHFlexGrid3.Rows = RsTxn_SContractAttachment.RecordCount + 2
For I = 1 To RsTxn_SContractAttachment.RecordCount

    MSHFlexGrid3.TextMatrix(I, 0) = I
    MSHFlexGrid3.TextMatrix(I, 1) = RsTxn_SContractAttachment!FileName
    MSHFlexGrid3.TextMatrix(I, 2) = RsTxn_SContractAttachment!Remarks
    MSHFlexGrid3.TextMatrix(I, 3) = App.Path & "\Attachment\" & MSHFlexGrid3.TextMatrix(I, 1)
    File_Path = MSHFlexGrid3.TextMatrix(I, 3) 'App.Path & "\Project Attachment\"
    Call LoadFileFromDB(File_Path, RsTxn_SContractAttachment, "File")
    MSHFlexGrid3.TextMatrix(I, 4) = RsTxn_SContractAttachment!TxnID
    RsTxn_SContractAttachment.MoveNext
Next
lblAttachment.Enabled = False
TxtAttachmentRemarks.Locked = True
CmdAttachment.Enabled = False
CmdGetFile.Enabled = False
CmdRemoveAttachment.Enabled = False
End Sub

Private Function GetSContractAttachID() As Double
Dim Retval As Double
tmp = "Select max(TxnID) from Txn_StorageContractAttachment"
Call TmpTableSMAttachment
Retval = IIf(IsNull(TmpRsSM_Attachment.Fields(0)), 0, TmpRsSM_Attachment.Fields(0))
Retval = Retval + 1
GetSContractAttachID = Retval
End Function

Private Sub DeleteBillingDetails()

tmp = " Delete from Mst_SContractBillingDetail Where ContractID=" & Val(TxtContractID.Text) & ""
Call TmpTable

End Sub


Private Sub SaveBillingDetails()

Call DeleteBillingDetails

Call TableMst_SContractBillingDetail(" Where ContractID=" & Val(TxtContractID.Text) & "")

For I = 1 To MSHFlexGrid2.Rows - 1
    If MSHFlexGrid2.TextMatrix(I, 1) <> "" Then
       RsMst_SContractBillingDetail.AddNew
       RsMst_SContractBillingDetail!SCBDID = GetMaxSCBDID
       RsMst_SContractBillingDetail!ContractID = Val(TxtContractID.Text)
       RsMst_SContractBillingDetail!FromDate = CDate(MSHFlexGrid2.TextMatrix(I, 0))
       RsMst_SContractBillingDetail!ToDate = CDate(MSHFlexGrid2.TextMatrix(I, 1))
       RsMst_SContractBillingDetail!BillingCycleID = Val(MSHFlexGrid2.TextMatrix(I, 5))
       RsMst_SContractBillingDetail!BillingQty = Val(MSHFlexGrid2.TextMatrix(I, 3))
       RsMst_SContractBillingDetail!BillingRate = Val(MSHFlexGrid2.TextMatrix(I, 4))
       RsMst_SContractBillingDetail!ResQtyforMIS = Val(MSHFlexGrid2.TextMatrix(I, 6))
       RsMst_SContractBillingDetail!G_UID = GetGUID
       RsMst_SContractBillingDetail.Update
       RsMst_SContractBillingDetail.Requery
    End If
Next

End Sub

Private Function GetMaxSCBDID() As Double
Dim Retval As Double
tmp = "Select max(SCBDID) from Mst_SContractBillingDetail"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetMaxSCBDID = Retval
End Function

Private Sub InDataBillingDetails()
Call TableMst_SContractBillingDetail(" Where ContractID=" & Val(TxtContractID.Text) & "")
MSHFlexGrid2.Rows = RsMst_SContractBillingDetail.RecordCount + 2
For I = 1 To MSHFlexGrid2.Rows - 2
    MSHFlexGrid2.TextMatrix(I, 0) = RsMst_SContractBillingDetail!FromDate
    MSHFlexGrid2.TextMatrix(I, 1) = RsMst_SContractBillingDetail!ToDate
    MSHFlexGrid2.TextMatrix(I, 3) = Format(RsMst_SContractBillingDetail!BillingQty, "#####0.000")
    MSHFlexGrid2.TextMatrix(I, 4) = Format(RsMst_SContractBillingDetail!BillingRate, "#####0.00")
    MSHFlexGrid2.TextMatrix(I, 5) = RsMst_SContractBillingDetail!BillingCycleID
    MSHFlexGrid2.TextMatrix(I, 2) = GetBillingCycle(RsMst_SContractBillingDetail!BillingCycleID, "N")
    MSHFlexGrid2.TextMatrix(I, 6) = Format(RsMst_SContractBillingDetail!ResQtyforMIS, "#####0.000")
    RsMst_SContractBillingDetail.MoveNext
Next
End Sub

''Private Sub SaveBillingDetails()
''Call TableMst_SContractBillingDetail(" Where ContractID=" & Val(TxtContractID.Text) & "")
''
''For I = 1 To MSHFlexGrid2.Rows - 1
''    If MSHFlexGrid2.TextMatrix(I, 1) <> "" Then
''
''       RsMst_SContractBillingDetail!FromDate = CDate(MSHFlexGrid2.TextMatrix(I, 0))
''       RsMst_SContractBillingDetail!ToDate = CDate(MSHFlexGrid2.TextMatrix(I, 1))
''       RsMst_SContractBillingDetail!BillingCycleID = Val(MSHFlexGrid2.TextMatrix(I, 5))
''       RsMst_SContractBillingDetail!BillingQty = Val(MSHFlexGrid2.TextMatrix(I, 3))
''       RsMst_SContractBillingDetail!BillingRate = Val(MSHFlexGrid2.TextMatrix(I, 4))
''       RsMst_SContractBillingDetail.Update
''       RsMst_SContractBillingDetail.Requery
''    End If
''Next
''End Sub

''Private Sub TxtDToDate_GotFocus()
''tmp = TxtDToDate.Value
''End Sub
''
''Private Sub TxtDToDate_LostFocus()
''If tmp <> TxtDToDate.Value Then
''  r = MSHFlexGrid2.RowSel
''  Call RemoveBillingDetailRows(r, "U")
''End If
''End Sub
''
''Private Sub TxtDFromDate_GotFocus()
''tmp = TxtDFromDate.Value
''End Sub
''
''Private Sub TxtDfromDate_LostFocus()
''If tmp <> TxtDFromDate.Value Then
''  r = MSHFlexGrid2.RowSel
''  Call RemoveBillingDetailRows(r, "U")
''End If
''End Sub

Private Function CheckReservationInvDate() As Variant

Dim mRetval As Variant

tmp = "Select * from Txn_ReservationInvoice Where ContractID=" & TxtContractID
tmp = tmp & " Order by YearMonth desc"

Call Tmp3Table

If TmpRs3.RecordCount > 0 Then
   mRetval = TmpRs3!YearMonth
   mYear = Year(mRetval)
   mTotalDays = Day(DateSerial(mYear, Month(mRetval) + 1, 0))

   mRetval = Format(Year(mRetval) & "-" & Month(mRetval) & "-" & mTotalDays, "yyyy-mm-dd")

Else
   mRetval = "NA"
End If

CheckReservationInvDate = mRetval
End Function

 
Private Sub CmbMSC_GotFocus()
If CmbLocationID.Text = "" Then
   MsgBox " Select Location!!"
   CmbLocationID.SetFocus
   Exit Sub
End If
tmp1 = CmbMSC.Text
tmp = "select  convert(varchar(30),ContractID)  + '~' +  ContractName  as ContractName,ContractID from Mst_StorageContract Where LocationID=" & mLocationID
If Val(TxtContractID) > 0 Then
   tmp = tmp & " And ContractID<>" & Val(TxtContractID) & ""
End If

Call Tmp3Table
CmbMSC.Clear
If TmpRs3.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To TmpRs3.RecordCount
    CmbMSC.AddItem TmpRs3!ContractName
    TmpRs3.MoveNext
Next
CmbMSC.Text = tmp1
End Sub
Private Sub CmbMSC_LostFocus()
If CmbMSC.Text = "" Then
   Exit Sub
End If
TmpRs3.MoveFirst
TmpRs3.Find "ContractName = '" & CmbMSC.Text & "'"
If TmpRs3.EOF Then
   MsgBox "Invalid selection "
   CmbMSC.SetFocus
   Exit Sub
End If
mMContractIDRow = TmpRs3!ContractID

End Sub


Public Function ShowContractName(CID_ As Double) As String
Dim Retval As String
Retval = ""
tmp = " Select * from Mst_StorageContract Where ContractID = " & CID_
Call Tmp3Table
If Not TmpRs3.EOF Then
  Retval = TmpRs3!ContractName
End If
ShowContractName = Retval
End Function

Private Function StartBillingDetails() As Boolean
Dim Retval As Boolean
Retval = False
If MSHFlexGrid2.Rows > 1 Then
   For u = 1 To MSHFlexGrid2.Rows - 1
       If MSHFlexGrid2.TextMatrix(u, 1) = "" Then Exit For
       If CDate(MSHFlexGrid2.TextMatrix(u, 0)) = BookedFrom.Value Then
          Retval = True
       End If
   Next
End If
StartBillingDetails = Retval
End Function
