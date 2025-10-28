VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmTxn_FumigationApproval 
   Caption         =   "Fumigation and Spraying Bill Approval"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   20160
   WindowState     =   2  'Maximized
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   9810
      Width           =   10935
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   9015
      Left            =   15120
      TabIndex        =   17
      Top             =   360
      Width           =   255
   End
   Begin VB.Frame Frame4 
      Height          =   9735
      Left            =   15
      TabIndex        =   18
      Top             =   0
      Width           =   15135
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
         Left            =   480
         TabIndex        =   19
         Top             =   5280
         Visible         =   0   'False
         Width           =   14220
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
            Left            =   3225
            Sorted          =   -1  'True
            TabIndex        =   78
            Top             =   555
            Width           =   5475
         End
         Begin VB.TextBox TxtSBillNo 
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
            Left            =   6240
            TabIndex        =   27
            Top             =   1305
            Width           =   1890
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
            Left            =   12210
            TabIndex        =   26
            Top             =   1740
            Width           =   1000
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
            Left            =   13230
            TabIndex        =   25
            Top             =   1740
            Width           =   1000
         End
         Begin VB.ComboBox CmbSVendorID 
            BeginProperty Font 
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
            Sorted          =   -1  'True
            TabIndex        =   24
            Top             =   1320
            Width           =   4485
         End
         Begin VB.TextBox TxtSFumgID 
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
            Left            =   4680
            MaxLength       =   12
            TabIndex        =   23
            Top             =   1305
            Width           =   1455
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
            ItemData        =   "FrmTxn_FumigationApproval.frx":0000
            Left            =   8760
            List            =   "FrmTxn_FumigationApproval.frx":0010
            Sorted          =   -1  'True
            TabIndex        =   22
            Top             =   555
            Width           =   1845
         End
         Begin VB.CheckBox ChkBatchNo 
            Appearance      =   0  'Flat
            Caption         =   "Check1"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   4380
            TabIndex        =   21
            Top             =   1830
            Width           =   225
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
            Left            =   8175
            Sorted          =   -1  'True
            TabIndex        =   20
            Top             =   1320
            Width           =   6450
         End
         Begin MSComCtl2.DTPicker STo 
            Height          =   375
            Left            =   1695
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   555
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
            Format          =   3997697
            CurrentDate     =   36494
         End
         Begin MSComCtl2.DTPicker SFrom 
            Height          =   375
            Left            =   135
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   555
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
            Format          =   3997697
            CurrentDate     =   36494
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
            Height          =   1455
            Left            =   120
            TabIndex        =   30
            Top             =   2160
            Width           =   13815
            _ExtentX        =   24368
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
            Height          =   300
            Left            =   5055
            TabIndex        =   79
            Top             =   180
            Width           =   1320
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Bill No"
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
            TabIndex        =   40
            Top             =   960
            Width           =   585
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
            Left            =   1995
            TabIndex        =   39
            Top             =   255
            Width           =   855
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
            Left            =   480
            TabIndex        =   38
            Top             =   255
            Width           =   855
         End
         Begin VB.Label Label28 
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
            Left            =   6360
            TabIndex        =   37
            Top             =   1800
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
            Left            =   8880
            TabIndex        =   36
            Top             =   1815
            Width           =   90
         End
         Begin VB.Label Label18 
            Caption         =   "Vendor Name"
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
            Left            =   1710
            TabIndex        =   35
            Top             =   990
            Width           =   1320
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Mapping ID"
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
            Left            =   4785
            TabIndex        =   34
            Top             =   960
            Width           =   1035
         End
         Begin VB.Label Label33 
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
            Left            =   9360
            TabIndex        =   33
            Top             =   255
            Width           =   555
         End
         Begin VB.Label Label34 
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
            Left            =   2640
            TabIndex        =   32
            Top             =   1830
            Width           =   1710
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            Caption         =   "ForwardTo"
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
            Left            =   10905
            TabIndex        =   31
            Top             =   960
            Width           =   990
         End
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   5055
         Left            =   120
         TabIndex        =   74
         Top             =   3495
         Width           =   14955
         _ExtentX        =   26379
         _ExtentY        =   8916
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
         TabCaption(0)   =   "Fumigation / Spraying Details"
         TabPicture(0)   =   "FrmTxn_FumigationApproval.frx":003A
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label14"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label15"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "MSHFlexGrid1"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Check1"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "TxtVendorData"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "CmdExcelDetail"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).ControlCount=   6
         TabCaption(1)   =   "Payment Details"
         TabPicture(1)   =   "FrmTxn_FumigationApproval.frx":0056
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame3"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Log Details"
         TabPicture(2)   =   "FrmTxn_FumigationApproval.frx":0072
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Label2"
         Tab(2).Control(1)=   "Label24"
         Tab(2).Control(2)=   "Label25"
         Tab(2).Control(3)=   "MSHFlexGrid3"
         Tab(2).Control(4)=   "TxtUpdatedByFAG"
         Tab(2).Control(4).Enabled=   0   'False
         Tab(2).Control(5)=   "TxtCreated"
         Tab(2).Control(5).Enabled=   0   'False
         Tab(2).Control(6)=   "TxtUpdated"
         Tab(2).Control(6).Enabled=   0   'False
         Tab(2).ControlCount=   7
         Begin VB.Frame Frame3 
            Height          =   4635
            Left            =   -74880
            TabIndex        =   98
            Top             =   360
            Width           =   14535
            Begin VB.Frame Frame5 
               Caption         =   "Vendor Bank A/C  Details"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1170
               Left            =   75
               TabIndex        =   130
               Top             =   105
               Width           =   14385
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
                  TabIndex        =   134
                  Top             =   705
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
                  TabIndex        =   133
                  Top             =   705
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
                  TabIndex        =   132
                  Top             =   300
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
                  TabIndex        =   131
                  Top             =   300
                  Width           =   5625
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
                  TabIndex        =   138
                  Top             =   705
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
                  TabIndex        =   137
                  Top             =   765
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
                  TabIndex        =   136
                  Top             =   360
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
                  TabIndex        =   135
                  Top             =   360
                  Width           =   1065
               End
            End
            Begin VB.Frame FrmFAG 
               Caption         =   "FAG Payment Details"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   3315
               Left            =   120
               TabIndex        =   99
               Top             =   1320
               Width           =   14355
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
                  Height          =   855
                  Left            =   8850
                  MaxLength       =   1000
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   114
                  Top             =   1155
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
                  Height          =   855
                  Left            =   2820
                  MaxLength       =   250
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   113
                  Top             =   1155
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
                  TabIndex        =   112
                  Top             =   270
                  Width           =   1785
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
                  MaxLength       =   18
                  TabIndex        =   109
                  TabStop         =   0   'False
                  Top             =   2070
                  Width           =   4800
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
                  MaxLength       =   18
                  TabIndex        =   108
                  TabStop         =   0   'False
                  Top             =   2070
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
                  MaxLength       =   18
                  TabIndex        =   107
                  TabStop         =   0   'False
                  Top             =   2070
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
                  TabIndex        =   106
                  TabStop         =   0   'False
                  Top             =   2475
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
                  TabIndex        =   105
                  TabStop         =   0   'False
                  Top             =   2895
                  Width           =   4800
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
                  TabIndex        =   104
                  TabStop         =   0   'False
                  Top             =   2895
                  Width           =   2055
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
                  TabIndex        =   103
                  TabStop         =   0   'False
                  Top             =   2475
                  Width           =   4800
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
                  Left            =   12090
                  TabIndex        =   102
                  TabStop         =   0   'False
                  Top             =   2895
                  Width           =   1935
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
                  TabIndex        =   101
                  Top             =   690
                  Width           =   1800
               End
               Begin MSComCtl2.DTPicker TxtFagDisbDate 
                  Height          =   375
                  Left            =   7125
                  TabIndex        =   115
                  Top             =   270
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
                  Format          =   3997697
                  CurrentDate     =   40255
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
                  ItemData        =   "FrmTxn_FumigationApproval.frx":008E
                  Left            =   2820
                  List            =   "FrmTxn_FumigationApproval.frx":009E
                  TabIndex        =   110
                  Top             =   690
                  Width           =   1815
               End
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
                  Left            =   7125
                  TabIndex        =   111
                  Top             =   690
                  Width           =   4575
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
                  TabIndex        =   100
                  Top             =   690
                  Width           =   4575
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
                  TabIndex        =   129
                  Top             =   1350
                  Width           =   1530
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
                  TabIndex        =   128
                  Top             =   1515
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
                  TabIndex        =   127
                  Top             =   330
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
                  TabIndex        =   126
                  Top             =   330
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
                  TabIndex        =   125
                  Top             =   750
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
                  TabIndex        =   124
                  Top             =   750
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
                  TabIndex        =   123
                  Top             =   2130
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
                  TabIndex        =   122
                  Top             =   2130
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
                  TabIndex        =   121
                  Top             =   2130
                  Width           =   1290
               End
               Begin VB.Label Label26 
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
                  TabIndex        =   120
                  Top             =   2565
                  Width           =   1140
               End
               Begin VB.Label Label19 
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
                  TabIndex        =   119
                  Top             =   2520
                  Width           =   1140
               End
               Begin VB.Label Label16 
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
                  TabIndex        =   118
                  Top             =   2940
                  Width           =   2445
               End
               Begin VB.Label Label21 
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
                  TabIndex        =   117
                  Top             =   2955
                  Width           =   900
               End
               Begin VB.Label Label20 
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
                  Left            =   11160
                  TabIndex        =   116
                  Top             =   2940
                  Width           =   900
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
            Height          =   360
            Left            =   -72495
            Locked          =   -1  'True
            TabIndex        =   93
            TabStop         =   0   'False
            Top             =   1185
            Width           =   5400
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
            Height          =   360
            Left            =   -72495
            Locked          =   -1  'True
            TabIndex        =   92
            TabStop         =   0   'False
            Top             =   720
            Width           =   5400
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
            Left            =   -72495
            Locked          =   -1  'True
            TabIndex        =   91
            TabStop         =   0   'False
            Top             =   1650
            Width           =   9855
         End
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
            Left            =   6240
            TabIndex        =   88
            Top             =   420
            Width           =   1000
         End
         Begin VB.TextBox TxtVendorData 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   1320
            MaxLength       =   12
            ScrollBars      =   2  'Vertical
            TabIndex        =   73
            Top             =   1020
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Select All"
            Height          =   255
            Left            =   165
            TabIndex        =   16
            Top             =   480
            Width           =   1215
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   3945
            Left            =   120
            TabIndex        =   75
            Top             =   1020
            Width           =   14715
            _ExtentX        =   25956
            _ExtentY        =   6959
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
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
            Height          =   2385
            Left            =   -74880
            TabIndex        =   94
            Top             =   2160
            Width           =   14715
            _ExtentX        =   25956
            _ExtentY        =   4207
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
         Begin VB.Label Label25 
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
            Left            =   -74760
            TabIndex        =   97
            Top             =   1245
            Width           =   885
         End
         Begin VB.Label Label24 
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
            Left            =   -74760
            TabIndex        =   96
            Top             =   780
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Updated By FAG"
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
            TabIndex        =   95
            Top             =   1710
            Width           =   2175
         End
         Begin VB.Label Label15 
            Caption         =   "Fumigation done with in days specified in commodity master"
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
            Left            =   2790
            TabIndex        =   90
            Top             =   690
            Width           =   5145
         End
         Begin VB.Label Label14 
            BackColor       =   &H0000FF00&
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
            Left            =   2205
            TabIndex        =   89
            Top             =   690
            Width           =   495
         End
      End
      Begin VB.Frame ButtonFrm 
         Height          =   1170
         Left            =   120
         TabIndex        =   62
         Top             =   8520
         Width           =   14970
         Begin VB.CommandButton CmdPrint 
            Caption         =   "Print"
            Height          =   825
            Left            =   13065
            TabIndex        =   84
            Top             =   195
            Width           =   1785
         End
         Begin VB.CommandButton ExitCmd 
            Caption         =   "E&xit"
            Height          =   400
            Left            =   10665
            TabIndex        =   72
            Top             =   600
            Width           =   2385
         End
         Begin VB.CommandButton SearchCmd 
            Caption         =   "Search"
            Height          =   400
            Left            =   10665
            TabIndex        =   71
            Top             =   195
            Width           =   2385
         End
         Begin VB.CommandButton LastCmd 
            Caption         =   "&Last"
            Height          =   400
            Left            =   8085
            TabIndex        =   70
            Top             =   600
            Width           =   2580
         End
         Begin VB.CommandButton DeleteCmd 
            Caption         =   "&Delete"
            Height          =   400
            Left            =   8085
            TabIndex        =   69
            Top             =   195
            Width           =   2580
         End
         Begin VB.CommandButton FirstCmd 
            Caption         =   "&First"
            Height          =   400
            Left            =   5505
            TabIndex        =   68
            Top             =   600
            Width           =   2580
         End
         Begin VB.CommandButton EditCmd 
            Caption         =   "&Edit"
            Height          =   400
            Left            =   5505
            TabIndex        =   67
            Top             =   195
            Width           =   2580
         End
         Begin VB.CommandButton PreviousCmd 
            Caption         =   "&Previous"
            Height          =   400
            Left            =   2805
            TabIndex        =   66
            Top             =   600
            Width           =   2700
         End
         Begin VB.CommandButton SaveCmd 
            Caption         =   "Save"
            Height          =   400
            Left            =   2805
            TabIndex        =   65
            Top             =   195
            Width           =   2700
         End
         Begin VB.CommandButton NextCmd 
            Caption         =   "&Next"
            Height          =   400
            Left            =   105
            TabIndex        =   64
            Top             =   600
            Width           =   2700
         End
         Begin VB.CommandButton AddCmd 
            Caption         =   "&Add New"
            Height          =   400
            Left            =   105
            TabIndex        =   63
            Top             =   195
            Width           =   2700
         End
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3345
         Left            =   120
         TabIndex        =   41
         Top             =   120
         Width           =   14970
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
            Left            =   12225
            Sorted          =   -1  'True
            TabIndex        =   12
            Top             =   2050
            Width           =   2655
         End
         Begin VB.TextBox TxtSTAmount 
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
            Left            =   4395
            MaxLength       =   12
            TabIndex        =   5
            Top             =   1095
            Width           =   1515
         End
         Begin VB.CommandButton CmdSTaxDesc 
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
            Left            =   15180
            TabIndex        =   81
            TabStop         =   0   'False
            Top             =   2670
            Width           =   375
         End
         Begin VB.ComboBox CmbSTaxDesc 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   12330
            Sorted          =   -1  'True
            TabIndex        =   15
            Top             =   2925
            Visible         =   0   'False
            Width           =   2550
         End
         Begin VB.TextBox TxtBatchID 
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
            Left            =   11040
            Locked          =   -1  'True
            MaxLength       =   12
            TabIndex        =   76
            TabStop         =   0   'False
            Top             =   225
            Width           =   1470
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
            Left            =   1575
            Sorted          =   -1  'True
            TabIndex        =   10
            Top             =   2050
            Width           =   4800
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
            ItemData        =   "FrmTxn_FumigationApproval.frx":00C5
            Left            =   7335
            List            =   "FrmTxn_FumigationApproval.frx":00D5
            Sorted          =   -1  'True
            TabIndex        =   11
            Top             =   2050
            Width           =   1830
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
            TabIndex        =   2
            Top             =   675
            Width           =   4755
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
            Left            =   8880
            Sorted          =   -1  'True
            TabIndex        =   3
            Top             =   675
            Width           =   5595
         End
         Begin VB.CommandButton CmdVendor 
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
            Left            =   14520
            TabIndex        =   45
            TabStop         =   0   'False
            Top             =   675
            Width           =   375
         End
         Begin VB.CommandButton CmdLocation 
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
            Left            =   6360
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   720
            Width           =   375
         End
         Begin VB.TextBox TxtBillNo 
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
            Left            =   10095
            MaxLength       =   50
            TabIndex        =   7
            Top             =   1095
            Width           =   1950
         End
         Begin VB.TextBox TxtPayableAmount 
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
            Left            =   7530
            Locked          =   -1  'True
            MaxLength       =   12
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   1095
            Width           =   1620
         End
         Begin VB.TextBox TxtActualPayable 
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
            MaxLength       =   12
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   1095
            Width           =   1515
         End
         Begin VB.TextBox TxtMappingID 
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
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   247
            Width           =   1515
         End
         Begin VB.TextBox TxtRemark 
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
            Height          =   480
            Left            =   1575
            MaxLength       =   250
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   9
            Top             =   1500
            Width           =   13290
         End
         Begin VB.TextBox TxtFAGRemarks 
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
            Height          =   840
            Left            =   1575
            MaxLength       =   250
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   14
            Top             =   2445
            Width           =   9300
         End
         Begin MSComCtl2.DTPicker TxtMappingDate 
            Height          =   375
            Left            =   4800
            TabIndex        =   43
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
            Format          =   3997697
            CurrentDate     =   39884
         End
         Begin MSComCtl2.DTPicker TxtFAGPostingDate 
            Height          =   375
            Left            =   13200
            TabIndex        =   13
            Top             =   2505
            Width           =   1650
            _ExtentX        =   2910
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
            Format          =   3997697
            CurrentDate     =   39884
         End
         Begin VB.TextBox TxtVendorID 
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
            Left            =   8880
            Locked          =   -1  'True
            TabIndex        =   49
            Top             =   675
            Visible         =   0   'False
            Width           =   5595
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
            Left            =   1575
            Locked          =   -1  'True
            TabIndex        =   48
            Top             =   675
            Visible         =   0   'False
            Width           =   4755
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
            Left            =   7335
            Locked          =   -1  'True
            TabIndex        =   46
            Top             =   2050
            Visible         =   0   'False
            Width           =   1770
         End
         Begin VB.TextBox TxtForwardTo 
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
            TabIndex        =   47
            Top             =   2050
            Visible         =   0   'False
            Width           =   4770
         End
         Begin MSComCtl2.DTPicker TxtBillDate 
            Height          =   375
            Left            =   13440
            TabIndex        =   8
            Top             =   1095
            Width           =   1410
            _ExtentX        =   2487
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
            Format          =   3997697
            CurrentDate     =   39884
         End
         Begin VB.TextBox TxtSTaxDesc 
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
            Left            =   12330
            Locked          =   -1  'True
            TabIndex        =   83
            TabStop         =   0   'False
            Top             =   2925
            Visible         =   0   'False
            Width           =   2550
         End
         Begin MSComCtl2.DTPicker TxtMonthYear 
            Height          =   375
            Left            =   8880
            TabIndex        =   1
            Top             =   255
            Width           =   1260
            _ExtentX        =   2223
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
            CustomFormat    =   "MMM-yyyy"
            Format          =   3997699
            CurrentDate     =   39850
            MinDate         =   38353
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
            Left            =   12225
            Locked          =   -1  'True
            TabIndex        =   139
            TabStop         =   0   'False
            Top             =   2050
            Width           =   2655
         End
         Begin VB.Label Label29 
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
            Height          =   255
            Left            =   9360
            TabIndex        =   140
            Top             =   2100
            Width           =   2685
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "For Month"
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
            Left            =   7515
            TabIndex        =   87
            Top             =   285
            Width           =   1035
         End
         Begin VB.Label Label12 
            Caption         =   "Note: Don't enter more than one bill in this screen"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   450
            Left            =   12600
            TabIndex        =   86
            Top             =   165
            Width           =   2295
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Tax Amount"
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
            Left            =   3195
            TabIndex        =   85
            Top             =   1140
            Width           =   1230
         End
         Begin VB.Label LblST 
            Caption         =   "Tax Desc"
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
            Left            =   11100
            TabIndex        =   82
            Top             =   3000
            Visible         =   0   'False
            Width           =   1230
         End
         Begin VB.Label Label10 
            Caption         =   "Bill Date"
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
            Left            =   12120
            TabIndex        =   80
            Top             =   1125
            Width           =   960
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Batch ID"
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
            Left            =   10200
            TabIndex        =   77
            Top             =   285
            Width           =   750
         End
         Begin VB.Label Label31 
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
            Left            =   6600
            TabIndex        =   61
            Top             =   2110
            Width           =   660
         End
         Begin VB.Label lblFAGPostingDate 
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
            Height          =   285
            Left            =   10980
            TabIndex        =   60
            Top             =   2550
            Width           =   1800
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            Caption         =   "ForwardTo"
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
            TabIndex        =   59
            Top             =   2110
            Width           =   1140
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
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
            Height          =   240
            Left            =   7530
            TabIndex        =   58
            Top             =   720
            Width           =   765
         End
         Begin VB.Label LblLocationID 
            AutoSize        =   -1  'True
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
            Height          =   240
            Left            =   120
            TabIndex        =   57
            Top             =   735
            Width           =   900
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "Bill No"
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
            Left            =   9360
            TabIndex        =   56
            Top             =   1155
            Width           =   705
         End
         Begin VB.Label Label7 
            Caption         =   "Total Payable"
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
            Left            =   5955
            TabIndex        =   55
            Top             =   1133
            Width           =   2205
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Actual Payable "
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
            TabIndex        =   54
            Top             =   1155
            Width           =   1410
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Mapping date"
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
            Left            =   3360
            TabIndex        =   53
            Top             =   300
            Width           =   1245
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Mapping ID"
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
            Top             =   300
            Width           =   1035
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
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
            Height          =   240
            Left            =   120
            TabIndex        =   51
            Top             =   1560
            Width           =   825
         End
         Begin VB.Label lblFAGRemarks 
            AutoSize        =   -1  'True
            Caption         =   "FAG Remarks"
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
            TabIndex        =   50
            Top             =   2745
            Width           =   1275
         End
      End
   End
End
Attribute VB_Name = "FrmTxn_FumigationApproval"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mVendorID, mLocationID, mForwardToID, mSForwardToID, mServiceTaxID, mForwardToEmailID, mBatchID As Variant, mPBatchID, mFAGBankID, mStatus As Variant
Const strChecked = "þ"
Const strUnChecked = "q"
Dim mSelRow As Variant
Dim mSelCol As Variant
Dim Edit_Flg As Variant
Dim FileName  As String
Dim tmp1 As String

Dim mBillingStateID As Double   ''-- Ullhas Patil 25 Oct 2017

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


Private Sub TxtFAGDisbAmt_LostFocus()

If TxtFAGDisbAmt = "" Then Exit Sub

If IsNumeric(TxtFAGDisbAmt) = False Or Val(TxtFAGDisbAmt) < 0 Then
   MsgBox "Invalid format!!!"
   TxtFAGDisbAmt.SetFocus
   Exit Sub
End If

If Val(TxtFAGDisbAmt) > Val(TxtPayableAmount) Then
   MsgBox "Disbusment amount should be less than or equal to Txn Amount!!!"
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
Public Sub Grid_Head()
With MSHFlexGrid1
    .Clear
    .Cols = 19
    .Rows = 3
    .FixedRows = 2
    .FixedCols = 0
    
    .TextMatrix(1, 0) = "Select"
    .TextMatrix(1, 1) = "Type"
    .TextMatrix(1, 2) = "Fumigation Detail ID/Spraying ID"
    .TextMatrix(1, 3) = "CMP"
    .TextMatrix(1, 4) = "Godown No"
    .TextMatrix(1, 5) = "Godown Address"
    .TextMatrix(1, 6) = "Rate per MT (NCMSL)"
    .TextMatrix(1, 7) = "Quantity (NCMSL)"
    .TextMatrix(1, 8) = "Amount"
    .TextMatrix(1, 9) = "Remark "
    .TextMatrix(1, 10) = "Rate per MT (Vendor)"
    .TextMatrix(1, 11) = "Quantity (Vendor)"
    .TextMatrix(1, 12) = "Amount (Vendor)"
    .TextMatrix(1, 13) = "Fumigation Date"
    .TextMatrix(1, 14) = "Stack No"
    .TextMatrix(1, 15) = "Commodity"
    .TextMatrix(1, 16) = "FumigationID"
    .TextMatrix(1, 17) = "Printed Date"
    .TextMatrix(1, 18) = "Last Fumigation Date"
    
    .WordWrap = True
    .ColWidth(0) = 700
    .ColWidth(1) = 800
    .ColWidth(2) = 1000
    .ColWidth(3) = 1600
    .ColWidth(4) = 900
    .ColWidth(5) = 1800
    .ColWidth(6) = 900
    .ColWidth(7) = 1000
    .ColWidth(8) = 1100
    .ColWidth(9) = 1500
    .ColWidth(10) = 900
    .ColWidth(11) = 1000
    .ColWidth(12) = 1100
    .ColWidth(13) = 1000
    
    .RowHeight(1) = 900
    
End With
End Sub

Private Sub AddCmd_Click()

Edit_Flg = "A"

Call GButtonLock(Me, False)
Check1.Enabled = True
Frame1.Enabled = True
TxtMappingID.Text = ""
CmbVendorID.Text = ""
CmbLocationID.Text = ""
TxtBatchID = ""
TxtVendorID.Text = ""
TxtLocationID.Text = ""

TxtVendorID.Visible = False
TxtLocationID.Visible = False

CmbVendorID.Visible = True
CmbLocationID.Visible = True
CmdVendor.Visible = True
CmdLocation.Visible = True

CmbStatus.Visible = True
TxtStatus.Visible = False

CmbForwardTo.Visible = True
TxtForwardTo.Visible = False
CmbForwardTo.Text = ""
CmbStatus.Text = "Pending"
lblFAGPostingDate.Visible = False
lblFAGRemarks.Visible = False
TxtFAGPostingDate.Visible = False
TxtFAGRemarks.Visible = False

CmbSTaxDesc.Visible = False
TxtSTaxDesc.Visible = False
LblST.Visible = False
TxtSTaxDesc.Text = ""
CmbSTaxDesc.Text = ""

'TxtRate.Text = ""
TxtMappingDate.Value = Date
TxtRemark.Text = ""
TxtPayableAmount = 0
TxtActualPayable = 0
TxtSTAmount = ""
TxtCreated = ""
TxtUpdated = ""

TxtBillNo = ""
TxtMappingDate.Value = Date
TxtBillDate.Value = Date
mLocationID = 0
mVendorID = 0
TxtMonthYear.Enabled = True
TxtMonthYear.Value = Date

''--- Ullhas Patil 25 Oct 2017
CmbBillingState.Visible = True: CmbBillingState = ""
TxtBillingState.Visible = False: TxtBillingState = ""
''-----------------------------

Call Grid_Head
Call Grid_HeadLog

End Sub

Private Sub Check1_Click()
If Format(TxtMappingDate.Value, "yyyy-mm-dd") > "2013-04-29" Then GoTo NewDetail
If Check1.Value = vbChecked Then
   For I = 2 To MSHFlexGrid1.Rows - 2
      MSHFlexGrid1.row = I
      MSHFlexGrid1.Col = 0
      MSHFlexGrid1.CellFontName = "Wingdings"
      MSHFlexGrid1.CellFontSize = 12
      MSHFlexGrid1.CellAlignment = flexAlignCenterCenter
      If MSHFlexGrid1.Text <> strChecked Then
         MSHFlexGrid1.Text = strChecked
         'MSHFlexGrid1.TextMatrix(0, 6) = Val(MSHFlexGrid1.TextMatrix(0, 6)) + Val(MSHFlexGrid1.TextMatrix(I, 6))
         MSHFlexGrid1.TextMatrix(0, 7) = Val(MSHFlexGrid1.TextMatrix(0, 7)) + Val(MSHFlexGrid1.TextMatrix(I, 7))
         MSHFlexGrid1.TextMatrix(0, 8) = Val(MSHFlexGrid1.TextMatrix(0, 8)) + Val(MSHFlexGrid1.TextMatrix(I, 8))
         MSHFlexGrid1.TextMatrix(I, 10) = MSHFlexGrid1.TextMatrix(I, 6)
         MSHFlexGrid1.TextMatrix(I, 11) = MSHFlexGrid1.TextMatrix(I, 7)
         MSHFlexGrid1.TextMatrix(I, 12) = Format(Val(MSHFlexGrid1.TextMatrix(I, 10)) * Val(MSHFlexGrid1.TextMatrix(I, 11)), "#####0.00")
         MSHFlexGrid1.TextMatrix(0, 11) = Val(MSHFlexGrid1.TextMatrix(0, 11)) + Val(MSHFlexGrid1.TextMatrix(I, 11))
         MSHFlexGrid1.TextMatrix(0, 12) = Val(MSHFlexGrid1.TextMatrix(0, 12)) + Val(MSHFlexGrid1.TextMatrix(I, 12))
      End If
   Next
Else
   For I = 2 To MSHFlexGrid1.Rows - 2
      MSHFlexGrid1.row = I
      MSHFlexGrid1.Col = 0
      MSHFlexGrid1.CellFontName = "Wingdings"
      MSHFlexGrid1.CellFontSize = 12
      MSHFlexGrid1.CellAlignment = flexAlignCenterCenter
      MSHFlexGrid1.Text = strUnChecked
      'MSHFlexGrid1.TextMatrix(0, 6) = Val(MSHFlexGrid1.TextMatrix(0, 6)) - Val(MSHFlexGrid1.TextMatrix(I, 6))
      MSHFlexGrid1.TextMatrix(0, 7) = ""
      MSHFlexGrid1.TextMatrix(0, 8) = ""
      
      MSHFlexGrid1.TextMatrix(0, 11) = ""
      MSHFlexGrid1.TextMatrix(0, 12) = ""
      
      MSHFlexGrid1.TextMatrix(I, 10) = ""
      MSHFlexGrid1.TextMatrix(I, 11) = ""
      MSHFlexGrid1.TextMatrix(I, 12) = ""
   Next
End If
'TxtPayableAmount.Text = MSHFlexGrid1.TextMatrix(0, 8)
TxtActualPayable = MSHFlexGrid1.TextMatrix(0, 12)

Exit Sub
NewDetail:
If Check1.Value = vbChecked Then
   For I = 2 To MSHFlexGrid1.Rows - 2
      MSHFlexGrid1.row = I
      MSHFlexGrid1.Col = 0
      MSHFlexGrid1.CellFontName = "Wingdings"
      MSHFlexGrid1.CellFontSize = 12
      MSHFlexGrid1.CellAlignment = flexAlignCenterCenter
      If MSHFlexGrid1.Text <> strChecked Then
         MSHFlexGrid1.Text = strChecked
         'MSHFlexGrid1.TextMatrix(0, 6) = Val(MSHFlexGrid1.TextMatrix(0, 6)) + Val(MSHFlexGrid1.TextMatrix(I, 6))
         MSHFlexGrid1.TextMatrix(0, 7) = Val(MSHFlexGrid1.TextMatrix(0, 7)) + Val(MSHFlexGrid1.TextMatrix(I, 7))
         MSHFlexGrid1.TextMatrix(0, 8) = Val(MSHFlexGrid1.TextMatrix(0, 8)) + Val(MSHFlexGrid1.TextMatrix(I, 8))
         MSHFlexGrid1.TextMatrix(I, 10) = MSHFlexGrid1.TextMatrix(I, 6)
         MSHFlexGrid1.TextMatrix(I, 11) = MSHFlexGrid1.TextMatrix(I, 7)
         MSHFlexGrid1.TextMatrix(I, 12) = Format(Val(MSHFlexGrid1.TextMatrix(I, 10)) * Val(MSHFlexGrid1.TextMatrix(I, 11)), "#####0.00")
         MSHFlexGrid1.TextMatrix(0, 11) = Val(MSHFlexGrid1.TextMatrix(0, 11)) + Val(MSHFlexGrid1.TextMatrix(I, 11))
         MSHFlexGrid1.TextMatrix(0, 12) = Val(MSHFlexGrid1.TextMatrix(0, 12)) + Val(MSHFlexGrid1.TextMatrix(I, 12))
      End If
   Next
Else
   For I = 2 To MSHFlexGrid1.Rows - 2
      MSHFlexGrid1.row = I
      MSHFlexGrid1.Col = 0
      MSHFlexGrid1.CellFontName = "Wingdings"
      MSHFlexGrid1.CellFontSize = 12
      MSHFlexGrid1.CellAlignment = flexAlignCenterCenter
      MSHFlexGrid1.Text = strUnChecked
      'MSHFlexGrid1.TextMatrix(0, 6) = Val(MSHFlexGrid1.TextMatrix(0, 6)) - Val(MSHFlexGrid1.TextMatrix(I, 6))
      MSHFlexGrid1.TextMatrix(0, 7) = ""
      MSHFlexGrid1.TextMatrix(0, 8) = ""
      
      MSHFlexGrid1.TextMatrix(0, 11) = ""
      MSHFlexGrid1.TextMatrix(0, 12) = ""
      
      MSHFlexGrid1.TextMatrix(I, 10) = ""
      MSHFlexGrid1.TextMatrix(I, 11) = ""
      MSHFlexGrid1.TextMatrix(I, 12) = ""
   Next
End If
'TxtPayableAmount.Text = MSHFlexGrid1.TextMatrix(0, 8)
TxtActualPayable = MSHFlexGrid1.TextMatrix(0, 12)
End Sub

Private Sub CmbStatus_LostFocus()
If SaveCmd.Enabled = False Then Exit Sub
If CmbStatus.Text = "" Then Exit Sub
If LCase(CmbStatus.Text) <> "pending" And LCase(CmbStatus.Text) <> "approved" And LCase(CmbStatus.Text) <> "rejected" And LCase(CmbStatus.Text) <> "settled" Then
    MsgBox "Invalid Selection!!!"
    CmbStatus.SetFocus
End If
LblST.Visible = False
CmbSTaxDesc.Visible = False
If LCase(CmbStatus.Text) = "pending" Then
    mStatus = "P"
    lblFAGPostingDate.Visible = False
    lblFAGRemarks.Visible = False
    TxtFAGPostingDate.Visible = False
    TxtFAGRemarks.Visible = False
    TxtFAGRemark = ""
ElseIf LCase(CmbStatus.Text) = "settled" Then
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
   TxtFAGDisbAmt.Text = Val(TxtPayableAmount)
   TxtDeductionRemark = "TDS@"
   TxtFAGRemark = Trim(TxtFAGRemarks)
   Call TableMst_FAGBank(" Where FAGBankID in (" & Gen_FAGBankIDForPayment & ") ")
   If RsMst_FAGBank.RecordCount > 0 Then
      CmbFAGBank.Text = RsMst_FAGBank!BankName & "~" & RsMst_FAGBank!AccountNo
      mFAGBankID = RsMst_FAGBank!FAGBankID
   End If
Else
    mStatus = "A"
    lblFAGPostingDate.Visible = True
    lblFAGRemarks.Visible = True
    TxtFAGPostingDate.Visible = True
    TxtFAGRemarks.Visible = True
    TxtFAGPostingDate.Value = Date
    TxtFAGRemarks = ""
    If LCase(CmbStatus.Text) = "approved" Then
       LblST.Visible = True
       CmbSTaxDesc.Visible = True
    End If
    TxtFAGRemark = ""
End If

If mStatus = "S" Then
    If mBatchID = 0 Then
        MsgBox "Txn not booked in NAVISION, please export batch from Store-Man and upload in NAVISION!!!"
           If FrmFAG.Enabled = True Then
              TxtFAGDisbAmt.Text = ""
              FrmFAG.Enabled = False
           End If
           CmbStatus.SetFocus
        Exit Sub
    End If
End If
 
If LCase(Gen_UserRole) <> "fag" And LCase(Gen_UserRole) <> "fagadmin" And LCase(Gen_UserRole) <> "power" Then
   If LCase(CmbStatus.Text) <> "pending" Then
      MsgBox "Please forward this claim to FAG user, Only FAG user have access to approve claim '!!!"
      CmbStatus.SetFocus
      Exit Sub
   End If
End If

End Sub

Private Sub CmdExcelDetail_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Txn_FumigationMapping.xls")
Call Xls_Detail_Rpt(MSHFlexGrid1)
End Sub


Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Txn_FumigationMapping.xls")
Call Xls_Detail_Rpt(MSHFlexGrid2)
End Sub

Private Sub Xls_Detail_Rpt(MSH_Grid As MSHFlexGrid)
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Integer
Dim mCol As Integer
MsgBox ("Wait till next message")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Txn_FumigationMapping.xls")
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


Private Sub CmdSearch_Click()
Dim Wc1 As Variant
Call Grid_Head1

If SFrom.Value > STo.Value Then
   MsgBox "From Date Must be less than To Date!!!"
   SFrom.SetFocus
   Exit Sub
End If

tmp = "Select TF.MappingID,TF.MappingDate,TF.MonthYear,MV.VendorName,"
tmp = tmp & " ML.LocationName ,TF.ActualQty,TF.ActualAmount, TF.PayableAmount,TF.BillNo, TF.Remark,ME.EmployeeName+'~'+ME.EmployeeNo 'Employee',TF.Flag,TF.ApprovalRemark,TF.ApprovedDate,TF.BatchID, "
tmp = tmp & " TF.BankName , TF.ChequeNo, TF.PaidAmount, TF.TDS, TF.PaidDate, TF.PaymentdetailsUpdatedBy, TF.PaymentdetailsUpdatedOn,MSII.StateIndiaName[BillingStateName] "
tmp = tmp & " ,MFB.BankName+'~'+MFB.AccountNo 'FAGBankName' ,TF.DisbuAmount,TF.DisbuDate,TF.ModeOfPay,TF.DeductionRemark From Txn_FumigationApproval TF"
tmp = tmp & " Inner Join Mst_Vendor MV on TF.VendorID=MV.VendorID"
tmp = tmp & " Inner Join Mst_Location ML on TF.LocationID=ML.LocationID"
tmp = tmp & " Inner Join Mst_Employee ME on TF.ForwardTo=ME.EmployeeID"
tmp = tmp & " Left join Mst_StateIndia MSII On TF.BillingStateIndiaID = MSII.StateIndiaID "     ''--- Ullhas Patil 25 Oct 2017
tmp = tmp & " Left Join Mst_FAGBank MFB On TF.FAGBankID = MFB.FAGBankID"
If Gen_DBType = "MDB" Then
   Wc1 = " Where TF.MappingDate Between #" & Format(SFrom.Value, Gen_DatFormat) & "# And #" & Format(STo.Value, Gen_DatFormat) & "# "
Else
   Wc1 = " Where TF.MappingDate Between '" & Format(SFrom.Value, Gen_DatFormat) & "' And '" & Format(STo.Value, Gen_DatFormat) & "' "
End If
If CmbSLocationID.Text <> "" Then
   Wc1 = Wc1 & " And ML.LocationName='" & CmbSLocationID.Text & "'"
Else
   If Gen_WcLocation <> "" Then
      Wc1 = Wc1 & " And ML." & Gen_WcLocation
   End If
End If

If CmbSVendorID.Text <> "" Then
   Wc1 = Wc1 & " And MV.VendorName='" & CmbSVendorID.Text & "'"
End If

If CmbSStatus.Text <> "" Then
   Wc1 = Wc1 & " And TF.Flag='" & Left(CmbSStatus.Text, 1) & "'"
End If


If TxtSFumgID.Text <> "" Then
   Wc1 = Wc1 & " And TF.MappingID=" & Val(TxtSFumgID.Text) & ""
End If

If TxtSBillNo.Text <> "" Then
   Wc1 = Wc1 & " And TF.BillNo Like '%" & Trim(TxtSBillNo.Text) & "%'"
End If

If ChkBatchNo.Value = 0 Then
   Wc1 = Wc1 & " And TF.BatchID = 0 "
End If

If CmbSForwardTo.Text <> "" Then
   Wc1 = Wc1 & " And TF.ForwardTo =" & Val(mSForwardToID) & ""
End If

tmp = tmp & Wc1 & " Order By MappingID"

Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   MSHFlexGrid2.Rows = TmpRs1.RecordCount + 3
   x = 2
   For I = 1 To TmpRs1.RecordCount
       For C = 0 To MSHFlexGrid2.Cols - 1
           MSHFlexGrid2.TextMatrix(x, C) = IIf(IsNull(TmpRs1.Fields(C)), "", TmpRs1.Fields(C))
       Next
       MSHFlexGrid2.TextMatrix(x, 1) = Format(MSHFlexGrid2.TextMatrix(x, 1), "dd-MMM-yyyy")
       MSHFlexGrid2.TextMatrix(x, 2) = Format(MSHFlexGrid2.TextMatrix(x, 2), "MMM-yyyy")
       MSHFlexGrid2.TextMatrix(x, 13) = Format(MSHFlexGrid2.TextMatrix(x, 13), "dd-MMM-yyyy")
       MSHFlexGrid2.TextMatrix(x, 21) = Format(MSHFlexGrid2.TextMatrix(x, 21), "dd-MMM-yyyy HH:mm:ss")
       If LCase(MSHFlexGrid2.TextMatrix(x, 11)) = "a" Then
          MSHFlexGrid2.TextMatrix(x, 11) = "Approved"
       ElseIf LCase(MSHFlexGrid2.TextMatrix(x, 11)) = "p" Then
          MSHFlexGrid2.TextMatrix(x, 11) = "Pending"
       ElseIf LCase(MSHFlexGrid2.TextMatrix(x, 11)) = "r" Then
          MSHFlexGrid2.TextMatrix(x, 11) = "Rejected"
       End If
       MSHFlexGrid2.TextMatrix(0, 7) = Val(MSHFlexGrid2.TextMatrix(0, 7)) + Val(MSHFlexGrid2.TextMatrix(x, 7))
       
       If MSHFlexGrid2.TextMatrix(x, 15) = "" Then
          MSHFlexGrid2.TextMatrix(x, 15) = IIf(IsNull(TmpRs1!FAGBankName), "", TmpRs1!FAGBankName)
       End If
 '
       If MSHFlexGrid2.TextMatrix(x, 16) = "" Then
          MSHFlexGrid2.TextMatrix(x, 16) = IIf(IsNull(TmpRs1!ModeOfPay), "", TmpRs1!ModeOfPay)
       End If
       
       If MSHFlexGrid2.TextMatrix(x, 17) = "" Then
          MSHFlexGrid2.TextMatrix(x, 17) = IIf(IsNull(TmpRs1!DisbuAmount), "", TmpRs1!DisbuAmount)
       End If
'DeductionRemark
       
       If MSHFlexGrid2.TextMatrix(x, 18) = "" Then
          MSHFlexGrid2.TextMatrix(x, 18) = IIf(IsNull(TmpRs1!DeductionRemark), "", TmpRs1!DeductionRemark)
       End If
       
       
       If MSHFlexGrid2.TextMatrix(x, 19) = "" Then
          MSHFlexGrid2.TextMatrix(x, 19) = IIf(IsNull(TmpRs1!DisbuDate), "", Format(TmpRs1!DisbuDate, "DD-MMM-YYYY"))
       End If
       
       TmpRs1.MoveNext
       x = x + 1
   Next
End If

Label22.Caption = TmpRs1.RecordCount
Label22.Refresh

End Sub

Private Sub CmdVendor_Click()
FrmMst_Vendor.Show
End Sub

Private Sub DeleteCmd_Click()
Dim ans As Variant
On Error GoTo msgError
If LCase(TxtStatus) = "approved" Or LCase(TxtStatus) = "rejected" Then
   MsgBox "This Claim is " & TxtStatus & ". You can't delete this claim!!!"
   Exit Sub
End If

ans = MsgBox("Do you really want to DELETE this record??? ", vbYesNo)
If ans = vbYes Then
   Call DeleteChildTableDetail
   RsTxn_FumigationApproval.Delete
   RsTxn_FumigationApproval.Update
   If RsTxn_FumigationApproval.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsTxn_FumigationApproval.MoveNext
   If RsTxn_FumigationApproval.EOF() Then
      RsTxn_FumigationApproval.MoveLast
   End If
   Call Indata
End If

Exit Sub
msgError:
MsgBox "Child record Present "
RsTxn_FumigationApproval.CancelUpdate

End Sub

Private Sub EditCmd_Click()

'If LCase(TxtStatus) = "approved" Or LCase(TxtStatus) = "rejected" Or LCase(TxtStatus) = "settled" Then
'   MsgBox "This Claim is " & TxtStatus & ". You can't edit this claim any more!!!"
'   Exit Sub
'End If

If LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "fag" And LCase(Gen_UserRole) <> "fagadmin" Then
   If LCase(TxtStatus) = "settled" Or LCase(TxtStatus) = "approved" Or LCase(TxtStatus) = "rejected" Or LCase(TxtStatus) = "frozen" Then
      MsgBox "This Claim is " & TxtStatus & ". You can't edit this Claim any more!!!"
      Exit Sub
   End If
Else
   If LCase(TxtStatus) = "rejected" Or LCase(TxtStatus) = "settled" Then
      MsgBox "This Claim is " & TxtStatus & ". You can't edit this Claim any more!!!"
      Exit Sub
   End If
End If


Check1.Enabled = True
Edit_Flg = "E"
Call GButtonLock(Me, False)
Frame1.Enabled = True

CmbVendorID.Text = TxtVendorID
CmbLocationID.Text = TxtLocationID

CmbStatus.Visible = True
TxtStatus.Visible = False

CmbForwardTo.Visible = True
TxtForwardTo.Visible = False
CmbForwardTo = TxtForwardTo
CmbStatus.Text = TxtStatus
TxtMonthYear.Enabled = False

''--- Ullhas Patil 25 Oct 2017
CmbBillingState.Visible = True
TxtBillingState.Visible = False
CmbBillingState = TxtBillingState
''-----------------------------

Call AddFumigationDetail
Call AddSprayingDetail

End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub
Public Sub Grid_HeadLog()
With MSHFlexGrid3
    .Clear
    .Cols = 7
    .Rows = 2
    .FixedRows = 1
    
    .TextMatrix(0, 0) = "Mapping ID"
    .TextMatrix(0, 1) = "Old Actual Amt"
    .TextMatrix(0, 2) = "New Actual Amt"
    .TextMatrix(0, 3) = "Old Forwarded To"
    .TextMatrix(0, 4) = "New Forwarded To"
    .TextMatrix(0, 5) = "Updated By"
    .TextMatrix(0, 6) = "Updated Date"
    
    .WordWrap = True
    .ColWidth(0) = 0
    .ColWidth(1) = 0
    .ColWidth(2) = 0
    .ColWidth(3) = 3000
    .ColWidth(4) = 3000
    .ColWidth(5) = 2000
    .ColWidth(6) = 2000
End With
End Sub

Private Sub CmbVendorID_GotFocus()
If CmbLocationID.Text = "" Then
   MsgBox "Please select Location!!"
   CmbLocationID.SetFocus
   Exit Sub
End If

tmp1 = CmbVendorID.Text

''Call TableMst_Vendor '("Where Flag = 'Y'")
tmp = " Select VendorID ,VendorName from Mst_Vendor Where VendorID in ("
tmp = tmp & " Select distinct VendorID  From Txn_Fumigation TF"
tmp = tmp & " Inner join Mst_Godown MG on MG.GodownID=TF.GodownID"
tmp = tmp & " Inner join Mst_CMP MC on MC.CMPID=MG.CMPID"
tmp = tmp & " Where (IsNull(TF.MappingID, 0) = 0 Or IsNull(TF.MappingID, 0)=" & Val(TxtMappingID) & ") And MC.LocationID =" & mLocationID & " )"
tmp = tmp & " OR VendorID in ("
tmp = tmp & " Select distinct VendorID From Txn_Spraying TF"
tmp = tmp & " Inner join Mst_Godown MG on MG.GodownID=TF.GodownID"
tmp = tmp & " Inner join Mst_CMP MC on MC.CMPID=MG.CMPID"
tmp = tmp & " Where (IsNull(TF.MappingID, 0) = 0 Or IsNull(TF.MappingID, 0)=" & Val(TxtMappingID) & ") And MC.LocationID =" & mLocationID & " )"

Call Tmp2Table

CmbVendorID.Clear
If TmpRs2.RecordCount = 0 Then
   MsgBox "No Vendor Found!!!"
   CmbVendorID.SetFocus
   'RsMst_Vendor.MoveFirst
   Exit Sub
End If
For I = 1 To TmpRs2.RecordCount
    CmbVendorID.AddItem TmpRs2!VendorName ' & "~" & RsMst_Vendor!VendorNo
    TmpRs2.MoveNext
Next

CmbVendorID.Text = tmp1

End Sub

Private Sub CmbVendorID_LostFocus()

If Trim(CmbVendorID.Text) = "" Then
   mVendorID = 0
   Exit Sub
End If

TmpRs2.MoveFirst
TmpRs2.Find "VendorName = '" & CmbVendorID.Text & "'"

If TmpRs2.EOF Then
   MsgBox "Invalid Selection "
   CmbVendorID.SetFocus
   Exit Sub
End If

If mVendorID <> TmpRs2!VendorID Then
   mVendorID = TmpRs2!VendorID
   Call Grid_Head
   If CmbLocationID.Text <> "" Then
      Call AddFumigationDetail
      Call AddSprayingDetail
   End If
End If

End Sub

Private Sub CmbLocationID_GotFocus()
tmp = CmbLocationID.Text
If Gen_WcLocation <> "" Then
   Call TableMst_Location(" Where " & Gen_WcLocation)
Else
   Call TableMst_Location '(" Where LocationID <> 0 ")
End If
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

Private Sub CmbLocationID_LostFocus()
If CmbLocationID.Text = "" Then
   mLocationID = 0
   mVendorID = 0
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
   
   CmbVendorID.Text = ""
   mLocationID = RsMst_Location!LocationID
   TxtPreFix = RsMst_Location!SM_Prefix

   Call Grid_Head
   If CmbVendorID.Text <> "" Then
      Call AddFumigationDetail
      Call AddSprayingDetail
   End If

End If


End Sub

Private Sub CmbForwardTo_GotFocus()

tmp1 = CmbForwardTo.Text
Call TableMst_Employee("Where  Flag = 'Y' ")
CmbForwardTo.Clear
''If RsMst_Employee.RecordCount = 0 Then
''   MsgBox "No Employee Found!!!"
''   CmbForwardTo.SetFocus
''   Exit Sub
''End If

'COmmeneted on 06 Jan 2014
'tmp = "Select DirectReportingTo from Mst_Employee Where EmployeeID=" & Gen_EmployeeID
'Call Tmp2Table
'If TmpRs2.RecordCount > 0 Then
'   Call UpperHierarchy(TmpRs2!DirectReportingTo)
'End If
'
'tmp = "Select EmployeeID,EmployeeName ,EmployeeNo  from Mst_Employee Where DesignationID=1 And Flag='Y'"
'Call Tmp2Table
'
'For i = 1 To TmpRs2.RecordCount
'    CmbForwardTo.AddItem TmpRs2!EmployeeName & "~" & TmpRs2!EmployeeNo
'    TmpRs2.MoveNext
'Next


If LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "superadmin" And LCase(Gen_UserRole) <> "headcmg" And LCase(Gen_UserRole) <> "admin" And LCase(Gen_UserRole) <> "fagadmin" And LCase(Gen_UserRole) <> "fag" Then
   'tmp = "Select EmployeeID,EmployeeName ,EmployeeNo  from Mst_Employee Where EmployeeNo in (Select Employee_ID from Mst_Users Where ) " 'DesignationID=1 And Flag='Y'"
   'Field user can't send bill directly top FAG person as discussed with Sankar Sir on 06 Jan 2014
   tmp = "Select EmployeeID,EmployeeName ,EmployeeNo from Mst_Employee Where Flag = 'Y' And EmployeeNo in (Select EmployeeNo From Mst_Users Where UserRoleID not in (7,14,11,4,13))"
   Call Tmp2Table

   For I = 1 To TmpRs2.RecordCount
       CmbForwardTo.AddItem TmpRs2!EmployeeName & "~" & TmpRs2!EmployeeNo
       TmpRs2.MoveNext
   Next
Else
   tmp = "Select EmployeeID,EmployeeName ,EmployeeNo from Mst_Employee Where Flag = 'Y' And EmployeeNo in (Select EmployeeNo From Mst_Users Where UserRoleID in (1,5,6,7,12,14))"
   Call Tmp2Table

   For I = 1 To TmpRs2.RecordCount
       CmbForwardTo.AddItem TmpRs2!EmployeeName & "~" & TmpRs2!EmployeeNo
       TmpRs2.MoveNext
   Next

End If


CmbForwardTo.Text = tmp1
End Sub

Private Sub CmbForwardTo_LostFocus()
Dim strSplitString() As String
Dim EmpNo_, EmpName_ As String

If CmbForwardTo.Text = "" Then Exit Sub

strSplitString = Split(CmbForwardTo, "~")
If UBound(strSplitString) < 1 Then
   MsgBox "Invalid Selection "
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
'TxtFWDDesignation = GetDesignaion(RsMst_Employee!DesignationID)
'mForwardToEmailID = RsMst_Employee!EmployeeEmailID
End Sub


Private Sub Form_Activate()
If Me.WindowState = vbMaximized Then
   SetScrollBars
End If
End Sub

Private Sub Form_Load()
Dim wc As Variant

Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Frame1.Enabled = False
   Call GButtonLockAD(Me)
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If
Call Grid_Head
Call Grid_HeadLog
Call Grid_Head1

SFrom.Value = Date - 365
STo.Value = Date
Label22.Caption = ""
CmbSLocationID.Text = ""
CmbSVendorID.Text = ""
TxtSFumgID.Text = ""
TxtSBillNo.Text = ""
CmbSStatus.Text = ""

tmp = "Select EmployeeID,EmployeeName +'~'+ EmployeeNo as Employee,DesignationID,EmployeeEmailID From Mst_Employee Where EmployeeNo = '" & Gen_UserLoginID & "'"
Call TmpTable

If TmpRs.RecordCount = 1 Then
   CmbSForwardTo.Text = TmpRs!Employee
   mSForwardToID = TmpRs!EmployeeID
Else
   CmbSForwardTo.Text = ""
   mSForwardToID = Null
End If

If Gen_WcLocation <> "" Then
   Call TableTxn_FumigationApproval(" Where " & Gen_WcLocation & " Order by MappingID")
Else
   Call TableTxn_FumigationApproval(" Order by MappingID")
End If
If RsTxn_FumigationApproval.RecordCount = 0 Then
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
If SaveCmd.Enabled = False Then Exit Sub

If Format(TxtMappingDate.Value, "yyyy-mm-dd") > "2013-04-29" Then GoTo NewDetail

TxtVendorData.Visible = False
With MSHFlexGrid1
    If .row < 2 Or .row = .Rows - 1 Then Exit Sub
    If .Col = 0 Then  '.Row = 1 And
      If .TextMatrix(.row, .Col) = "" Then Exit Sub
      .Col = 0
      '
      If LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "headcmg" And LCase(Gen_UserRole) <> "superadmin" And LCase(Gen_UserRole) <> "admin" Then
         If CDate(.TextMatrix(.row, 13)) < Date - Val(Gen_BlockingFumigationBilling) Then
            MsgBox "For marking billing beyond the past " & Val(Gen_BlockingFumigationBilling) & " day(s), you need to provide the details to CMG with reasons for such delay in marking the bills. CMG will mark the bills on your behalf, if they are satisfied with your explanation!!!"
            If .TextMatrix(.row, 0) = strChecked Then
               .TextMatrix(.row, 0) = strUnChecked
            End If
            Exit Sub
         End If
      End If
      
      If .TextMatrix(.row, .Col) = strUnChecked Then
         .TextMatrix(.row, .Col) = strChecked
         .TextMatrix(0, 7) = Val(.TextMatrix(0, 7)) + Val(.TextMatrix(.row, 7))
         .TextMatrix(0, 8) = Val(.TextMatrix(0, 8)) + Val(.TextMatrix(.row, 8))
         .TextMatrix(.row, 10) = .TextMatrix(.row, 6)
         .TextMatrix(.row, 11) = .TextMatrix(.row, 7)
         .TextMatrix(.row, 12) = Format(Val(.TextMatrix(.row, 10)) * Val(.TextMatrix(.row, 11)), "#####0.00")
         .TextMatrix(0, 11) = Val(.TextMatrix(0, 11)) + Val(.TextMatrix(.row, 11))
         .TextMatrix(0, 12) = Val(.TextMatrix(0, 12)) + Val(.TextMatrix(.row, 12))
      Else
         .TextMatrix(.row, .Col) = strUnChecked
         .TextMatrix(0, 7) = Val(.TextMatrix(0, 7)) - Val(.TextMatrix(.row, 7))
         .TextMatrix(0, 8) = Val(.TextMatrix(0, 8)) - Val(.TextMatrix(.row, 8))
         
         .TextMatrix(0, 11) = Val(.TextMatrix(0, 11)) - Val(.TextMatrix(.row, 11))
         .TextMatrix(0, 12) = Val(.TextMatrix(0, 12)) - Val(.TextMatrix(.row, 12))
         
         .TextMatrix(.row, 10) = ""
         .TextMatrix(.row, 11) = ""
         .TextMatrix(.row, 12) = ""
      
      End If
      'TxtPayableAmount.Text = .TextMatrix(0, 8)
      TxtActualPayable = .TextMatrix(0, 12)
   End If
   
   If .Col = 10 Or .Col = 11 Then
      If .TextMatrix(.row, 0) = strUnChecked Then Exit Sub
      mSelRow = .row
      mSelCol = .Col
      TxtVendorData.Visible = True
      TxtVendorData.Left = .Left + .CellLeft
      TxtVendorData.Top = .Top + .CellTop
      TxtVendorData.Width = .CellWidth - 10
      TxtVendorData.Height = .CellHeight - 10
      TxtVendorData.Text = .Text
      TxtVendorData.SelStart = 0
      TxtVendorData.SelLength = Len(.Text)
      TxtVendorData.ZOrder
      TxtVendorData.SetFocus
   End If
End With

Exit Sub
NewDetail:

TxtVendorData.Visible = False

With MSHFlexGrid1
    If .row < 2 Or .row = .Rows - 1 Then Exit Sub
    If .Col = 0 Then  '.Row = 1 And
      If .TextMatrix(.row, .Col) = "" Then Exit Sub
      .Col = 0
      'LCase(Gen_UserRole) <> "power" And
      If LCase(Gen_UserRole) <> "headcmg" And LCase(Gen_UserRole) <> "superadmin" And LCase(Gen_UserRole) <> "admin" Then
         If CDate(.TextMatrix(.row, 13)) < Date - Val(Gen_BlockingFumigationBilling) Then
            MsgBox "For marking billing beyond the past " & Val(Gen_BlockingFumigationBilling) & " day(s), you need to provide the details to CMG with reasons for such delay in marking the bills. CMG will mark the bills on your behalf, if they are satisfied with your explanation!!!"
            If .TextMatrix(.row, 0) = strChecked Then
               .TextMatrix(.row, 0) = strUnChecked
            End If
            Exit Sub
         End If
      End If
      
      
      If .TextMatrix(.row, .Col) = strUnChecked Then
         If LCase(.TextMatrix(.row, 1)) = "spraying" Then
            .TextMatrix(.row, .Col) = strChecked
            .TextMatrix(0, 7) = Val(.TextMatrix(0, 7)) + Val(.TextMatrix(.row, 7))
            .TextMatrix(0, 8) = Val(.TextMatrix(0, 8)) + Val(.TextMatrix(.row, 8))
            .TextMatrix(.row, 10) = .TextMatrix(.row, 6)
            .TextMatrix(.row, 11) = .TextMatrix(.row, 7)
            .TextMatrix(.row, 12) = Format(Val(.TextMatrix(.row, 10)) * Val(.TextMatrix(.row, 11)), "#####0.00")
            .TextMatrix(0, 11) = Val(.TextMatrix(0, 11)) + Val(.TextMatrix(.row, 11))
            .TextMatrix(0, 12) = Val(.TextMatrix(0, 12)) + Val(.TextMatrix(.row, 12))
         Else
            For x = 2 To .Rows - 2
                If .TextMatrix(.row, 16) = .TextMatrix(x, 16) Then
                   .TextMatrix(x, .Col) = strChecked
                   .TextMatrix(0, 7) = Val(.TextMatrix(0, 7)) + Val(.TextMatrix(x, 7))
                   .TextMatrix(0, 8) = Val(.TextMatrix(0, 8)) + Val(.TextMatrix(x, 8))
                   .TextMatrix(x, 10) = .TextMatrix(x, 6)
                   .TextMatrix(x, 11) = .TextMatrix(x, 7)
                   .TextMatrix(x, 12) = Format(Val(.TextMatrix(x, 10)) * Val(.TextMatrix(x, 11)), "#####0.00")
                   .TextMatrix(0, 11) = Val(.TextMatrix(0, 11)) + Val(.TextMatrix(x, 11))
                   .TextMatrix(0, 12) = Val(.TextMatrix(0, 12)) + Val(.TextMatrix(x, 12))
                End If
            Next
         End If
      Else
         If LCase(.TextMatrix(.row, 1)) = "spraying" Then
            .TextMatrix(.row, .Col) = strUnChecked
            .TextMatrix(0, 7) = Val(.TextMatrix(0, 7)) - Val(.TextMatrix(.row, 7))
            .TextMatrix(0, 8) = Val(.TextMatrix(0, 8)) - Val(.TextMatrix(.row, 8))
   
            .TextMatrix(0, 11) = Val(.TextMatrix(0, 11)) - Val(.TextMatrix(.row, 11))
            .TextMatrix(0, 12) = Val(.TextMatrix(0, 12)) - Val(.TextMatrix(.row, 12))
   
            .TextMatrix(.row, 10) = ""
            .TextMatrix(.row, 11) = ""
            .TextMatrix(.row, 12) = ""
         Else
            For x = 2 To .Rows - 2
                If .TextMatrix(.row, 16) = .TextMatrix(x, 16) Then
                   .TextMatrix(x, .Col) = strUnChecked
                   .TextMatrix(0, 7) = Val(.TextMatrix(0, 7)) - Val(.TextMatrix(x, 7))
                   .TextMatrix(0, 8) = Val(.TextMatrix(0, 8)) - Val(.TextMatrix(x, 8))
          
                   .TextMatrix(0, 11) = Val(.TextMatrix(0, 11)) - Val(.TextMatrix(x, 11))
                   .TextMatrix(0, 12) = Val(.TextMatrix(0, 12)) - Val(.TextMatrix(x, 12))
         
                   .TextMatrix(x, 10) = ""
                   .TextMatrix(x, 11) = ""
                   .TextMatrix(x, 12) = ""
                End If
            Next
         End If
      End If
      'TxtPayableAmount.Text = .TextMatrix(0, 8)
      TxtActualPayable = .TextMatrix(0, 12)
   End If
   
   If .Col = 10 Or .Col = 11 Then
      If .TextMatrix(.row, 0) = strUnChecked Then Exit Sub
      mSelRow = .row
      mSelCol = .Col
      TxtVendorData.Visible = True
      TxtVendorData.Left = .Left + .CellLeft
      TxtVendorData.Top = .Top + .CellTop
      TxtVendorData.Width = .CellWidth - 10
      TxtVendorData.Height = .CellHeight - 10
      TxtVendorData.Text = .Text
      TxtVendorData.SelStart = 0
      TxtVendorData.SelLength = Len(.Text)
      TxtVendorData.ZOrder
      TxtVendorData.SetFocus
   End If
End With
End Sub

Private Sub MSHFlexGrid2_Click()
Dim I As Variant
I = MSHFlexGrid2.RowSel
If I <= 0 Then Exit Sub
If MSHFlexGrid2.TextMatrix(I, 0) = "" Then Exit Sub
RsTxn_FumigationApproval.MoveFirst
RsTxn_FumigationApproval.Find "MappingID= " & Val(MSHFlexGrid2.TextMatrix(I, 0)) & ""
If Not RsTxn_FumigationApproval.EOF Then
   Call Indata
   Frame0.Visible = False
End If

End Sub

Private Sub SaveCmd_Click()

''If TxtFumDate.Value > Date Then
''   MsgBox "Future Date not allowed!!"
''   TxtFumDate.SetFocus
''   Exit Sub
''End If

If CmbVendorID.Text = "" Then
   MsgBox "Select Vendor Name!!"
   CmbVendorID.SetFocus
   Exit Sub
End If

If CmbLocationID.Text = "" Then
   MsgBox "Select Location Name!!"
   CmbLocationID.SetFocus
   Exit Sub
End If

If TxtSTAmount = "" Then
   MsgBox "Blank ST Amount not Allowed!!"
   TxtSTAmount.SetFocus
   Exit Sub
End If

If TxtBillNo = "" Then
   MsgBox "Blank Bill No not Allowed!!"
   TxtBillNo.SetFocus
   Exit Sub
End If

If TxtRemark = "" Then
   MsgBox "Blank Remark not Allowed!!"
   TxtRemark.SetFocus
   Exit Sub
End If
   
If CmbForwardTo.Text = "" Then
   MsgBox "Select Forwarded To!!"
   CmbForwardTo.SetFocus
   Exit Sub
End If

If CmbStatus.Text = "" Then
   MsgBox "Blank Status not Allowed!!"
   CmbStatus.SetFocus
   Exit Sub
End If

''-- Ullhas Patil 25 Oct 2017
If CmbBillingState.Text = "" Then
   MsgBox "Blank Billing State not allowed!!!"
   CmbBillingState.SetFocus
   Exit Sub
End If
''---------------------------

If LCase(CmbStatus.Text) <> "pending" Then
   If TxtFAGRemarks.Text = "" Then
      MsgBox "Blank FAG Remarks not Allowed!!"
      TxtFAGRemarks.SetFocus
      Exit Sub
   End If
End If

If LCase(CmbStatus.Text) = "approved" Then
   If CmbSTaxDesc.Text = "" Then
      MsgBox "Please select S.Tax!!"
      CmbSTaxDesc.SetFocus
      Exit Sub
   End If
End If

tmp = "Select * from Txn_FumigationApproval Where LocationID=" & mLocationID & " And VendorID=" & mVendorID & " And BillNo='" & TxtBillNo.Text & "' And MappingID <> " & Val(TxtMappingID) & ""
Call TmpTable
If TmpRs.RecordCount > 0 Then
   MsgBox " Duplicate Bill No. not allowed for selected location & vendor!!"
   TxtBillNo.SetFocus
   Exit Sub
End If
If MSHFlexGrid1.Rows < 4 Then
   MsgBox "Fumigation/Spraying Details Missing."
   Exit Sub
End If

If TxtBillDate.Value > Date Then
   MsgBox "Future Date not allowed!!"
   TxtBillDate.SetFocus
   Exit Sub
End If

If CDate(Format(TxtMonthYear, "yyyymm")) > Format(Date, "yyyymm") Then 'If TxtMonthYear.Value > Date Then
   MsgBox "Future Month not allowed!!"
   TxtMonthYear.SetFocus
   Exit Sub
End If

If Edit_Flg = "A" Then
   If Format(Gen_BlockingDate, "yyyy-mm") >= Format(TxtMonthYear.Value, "yyyy-mm") Then
      MsgBox "Select next month from " & Format(Gen_BlockingDate, "MMM-yyyy") & "!!! "
      TxtMonthYear.SetFocus
      Exit Sub
   End If
End If

If CheckGridSelect = False Then
   MsgBox "Select atleast one Fumigation/Spraying ID!!!"
   Exit Sub
End If

If LCase(CmbStatus.Text) <> "rejected" Then
   If CheckGridData = tr Then
      MsgBox "Enter Vendor rate & vendor Quantity for all selected rows!!!"
      Exit Sub
   End If
End If

TxtPayableAmount.Text = Format(Val(TxtActualPayable) + Val(TxtSTAmount), "#####0.00")

If LCase(mStatus) = "s" Then
  If Val(TxtFAGDisbAmt.Text) > 0 Then
      If TxtFagDisbDate.Value > Date Then
         MsgBox "Future date not allowed!!"
         TxtFagDisbDate.SetFocus
         SSTab1.Tab = 1
         Exit Sub
      End If
      If TxtFagDisbDate.Value < TxtMappingDate.Value Then
         MsgBox "Disbursement Date should not less than Mapping Date!!"
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
      If CmbModeOfPay.Text = "" Then
         MsgBox "Blank Mode of Payment not allowed!!"
         CmbModeOfPay.SetFocus
         Exit Sub
      End If
      If LCase(CmbModeOfPay.Text) = "bank transfer" Or LCase(CmbModeOfPay.Text) = "neft" Or LCase(CmbModeOfPay.Text) = "rtgs" Then
         If TxtACNo.Text = "" Then
            MsgBox "Blank Account Details not allowed for 'Bank Transfer', Contact FAG to update Vendor Account details!!"
            SSTab1.Tab = 1
            CmbModeOfPay.SetFocus
            Exit Sub
         End If
      End If
      If Val(TxtFAGDisbAmt) < Val(TxtPayableAmount) Then
         If Trim(TxtDeductionRemark) = "" Then
            MsgBox "As disbursment amount is less than txn amount please provide deduction remarks!!!"
            SSTab1.Tab = 1
            TxtDeductionRemark.SetFocus
            Exit Sub
         End If
      End If
    End If
End If

If Edit_Flg = "A" Then
   RsTxn_FumigationApproval.AddNew
'   RsTxn_FumigationApproval!MappingID = GetMaxMappingID
   RsTxn_FumigationApproval!G_UID = GetGUID
   mGUID = RsTxn_FumigationApproval!G_UID
   RsTxn_FumigationApproval!BatchID = 0
'   TxtMappingID.Text = RsTxn_FumigationApproval!MappingID
Else
   RsTxn_FumigationApproval.MoveFirst
   RsTxn_FumigationApproval.Find "MappingID= '" & Val(TxtMappingID.Text) & "'"
End If

RsTxn_FumigationApproval!MappingDate = Format(TxtMappingDate.Value, Gen_DatFormat)
RsTxn_FumigationApproval!MonthYear = Format(TxtMonthYear.Value, Gen_DatFormat)
RsTxn_FumigationApproval!VendorID = mVendorID
RsTxn_FumigationApproval!LocationID = mLocationID
RsTxn_FumigationApproval!Remark = TxtRemark.Text
RsTxn_FumigationApproval!ActualQty = Val(MSHFlexGrid1.TextMatrix(0, 7))
RsTxn_FumigationApproval!PayableAmount = Format(Val(TxtPayableAmount.Text), "#####0.00")
RsTxn_FumigationApproval!ActualAmount = Format(Val(TxtActualPayable.Text), "#####0.00")
RsTxn_FumigationApproval!BillNo = TxtBillNo.Text
RsTxn_FumigationApproval!BillDate = Format(TxtBillDate.Value, Gen_DatFormat)
RsTxn_FumigationApproval!STAmount = Val(TxtSTAmount)
RsTxn_FumigationApproval!ForwardTo = mForwardToID
RsTxn_FumigationApproval!ServiceTaxID = mServiceTaxID
RsTxn_FumigationApproval!Flag = Left$(CmbStatus.Text, 1)

If LCase(CmbStatus.Text) <> "pending" Then
    RsTxn_FumigationApproval!ApprovedDate = TxtFAGPostingDate.Value
    RsTxn_FumigationApproval!ApprovalRemark = TxtFAGRemarks.Text
End If

If LCase(CmbStatus.Text) = "settled" Then
   RsTxn_FumigationApproval!DisbuDate = TxtFagDisbDate.Value
   RsTxn_FumigationApproval!DisbuRemark = TxtFAGRemark.Text
   RsTxn_FumigationApproval!DisbuAmount = TxtFAGDisbAmt.Text
   RsTxn_FumigationApproval!ModeOfPay = CmbModeOfPay.Text
   RsTxn_FumigationApproval!FAGBankID = mFAGBankID
   RsTxn_FumigationApproval!DeductionRemark = Trim(TxtDeductionRemark)
Else
   RsTxn_FumigationApproval!DisbuDate = Null
   RsTxn_FumigationApproval!DisbuRemark = Null
   RsTxn_FumigationApproval!DisbuAmount = Null
   RsTxn_FumigationApproval!ModeOfPay = Null
   RsTxn_FumigationApproval!FAGBankID = Null
   RsTxn_FumigationApproval!DeductionRemark = Null
End If

RsTxn_FumigationApproval!BillingStateIndiaID = mBillingStateID      ''-- Ullhas Patil 25 Oct 2017

If IsNull(RsTxn_FumigationApproval!CreatedBy) = True Or RsTxn_FumigationApproval!CreatedBy = "" Then
   RsTxn_FumigationApproval!CreatedBy = Gen_UserLoginID
   RsTxn_FumigationApproval!CreatedDate = Now
Else
   RsTxn_FumigationApproval!UpdatedBy = Gen_UserLoginID
   RsTxn_FumigationApproval!UpdatedDate = Now
End If

RsTxn_FumigationApproval.Update


tmp = "Select MappingID From Txn_FumigationApproval Where G_UID='" & mGUID & "'"
Call TmpTable

If TmpRs.RecordCount > 0 Then
'   mCxTF_DetailsID = TmpRs!CxTF_DetailsID
   TxtMappingID.Text = TmpRs!MappingID
End If

Call SaveFumigationDetails

mForwardToEmailID = GetEmployeeEmailID(mForwardToID)

SaveCmd.Enabled = False

Call Indata
Call Gen_Email

End Sub

Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   If RsTxn_FumigationApproval.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   RsTxn_FumigationApproval.MoveFirst
   RsTxn_FumigationApproval.Find "MappingID = " & Val(TxtMappingID) & ""
   If RsTxn_FumigationApproval.EOF Then
      RsTxn_FumigationApproval.MoveLast
   End If
   Call Indata
   Exit Sub
End If
Frame0.Height = Me.Height - 800
Frame0.Left = Frame1.Left
Frame0.Top = Frame1.Top
Frame0.Width = Frame1.Width
Frame0.Visible = True

MSHFlexGrid2.Width = Frame0.Width - 300
MSHFlexGrid2.Height = Frame0.Height - 1000

RsTxn_FumigationApproval.MoveFirst
RsTxn_FumigationApproval.Filter = ""

''===========comment on 7 july 2012 for not clearing search grid=============
'Call Grid_Head1
'SFrom.Value = Date - 365
'STo.Value = Date
'Label22.Caption = ""
'CmbSLocationID.Text = ""
'CmbSVendorID.Text = ""
'TxtSFumgID.Text = ""
'TxtSBillNo.Text = ""
'CmbSStatus.Text = ""
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


Private Sub TxtBillDate_GotFocus()
tmp = TxtBillDate.Value
End Sub

Private Sub TxtBillDate_LostFocus()
If TxtBillDate.Value <> tmp Then
   If CmbVendorID.Text <> "" And CmbLocationID.Text <> "" Then
      Call Grid_Head
      Call AddFumigationDetail
      Call AddSprayingDetail
   End If
End If
End Sub

Private Sub TxtBillNo_LostFocus()
If TxtBillNo.Text = "" Then Exit Sub
If InStr(1, TxtBillNo, ",") > 0 Then
   MsgBox "Comma not allowed!!"
   TxtBillNo.SetFocus
   Exit Sub
End If

End Sub

Private Sub TxtFAGPostingDate_LostFocus()

If Format(Gen_BlockingDate, "yyyy-mm") >= Format(TxtFAGPostingDate.Value, "yyyy-mm") Then
   MsgBox "FAG Posting Date allowed for next month of " & Format(Gen_BlockingDate, "MMM-yyyy") & "!!!"
   TxtFAGPostingDate.SetFocus
   Exit Sub
End If

End Sub

Private Sub TxtMonthYear_LostFocus()
mStartDate = (Year(TxtMonthYear.Value) & "-" & Month(TxtMonthYear.Value) & "-01")
mTotalDays = Day(DateSerial(Year(TxtMonthYear.Value), Month(TxtMonthYear.Value) + 1, 0))
TxtMonthYear.Value = (Year(TxtMonthYear.Value) & "-" & Month(TxtMonthYear.Value) & "-" & mTotalDays)

End Sub

Private Sub TxtSTAmount_LostFocus()
If TxtSTAmount = "" Then Exit Sub
If Not IsNumeric(TxtSTAmount) Or Val(TxtSTAmount) < 0 Then
   MsgBox "Invalid numeric format !!!! "
   TxtSTAmount.SetFocus
End If
TxtPayableAmount.Text = Format(Val(TxtActualPayable) + Val(TxtSTAmount), "#####0.00")
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
VScroll1.Max = Me.ScaleHeight - Frame4.Height '4
VScroll1.LargeChange = Me.ScaleHeight
VScroll1.SmallChange = Me.ScaleHeight / 5
HScroll1.Max = Me.ScaleWidth - Frame4.Width '4
HScroll1.LargeChange = Me.ScaleWidth
HScroll1.SmallChange = Me.ScaleWidth / 5
If Frame4.Height >= Me.Height Then VScroll1.Enabled = True ''4
If Frame4.Height < Me.Height Then VScroll1.Enabled = False '4
If Frame4.Width >= Me.Width Then HScroll1.Enabled = True '4
If Frame4.Width < Me.Width Then HScroll1.Enabled = False '4

End Sub

Private Sub AddFumigationDetail()

If Format(TxtMappingDate.Value, "yyyy-mm-dd") > "2013-04-29" Then GoTo NewDetail

tmp = "Select TF.FumigationID,MC.CMPName,MG.GodownNo,MG.Address,TF.Rate,TF.BillableQty,TF.Amount,TF.Remark,TF.VendorRate,TF.VendorQuantity,TF.MappingID,TF.FumigationDate,TF.PrintedDate from Txn_Fumigation TF"
tmp = tmp & " Inner join Mst_Godown MG on MG.GodownID=TF.GodownID"
tmp = tmp & " Inner join Mst_CMP MC on MC.CMPID=MG.CMPID"
tmp = tmp & " Where IsNull(TF.MappingID, 0) = 0 And isnull(TF.BillFlag,'')='' And TF.VendorID = " & mVendorID & " And MC.LocationID = " & mLocationID & " And TF.FumigationDate<= '" & TxtBillDate.Value & "'"
tmp = tmp & " Order by TF.FumigationID"
Call TmpTable

If TmpRs.RecordCount > 0 Then
   For I = 1 To TmpRs.RecordCount
       rn = MSHFlexGrid1.Rows - 1
       MSHFlexGrid1.TextMatrix(rn, 1) = "Fumigation"
       MSHFlexGrid1.TextMatrix(rn, 2) = TmpRs!FumigationID
       MSHFlexGrid1.TextMatrix(rn, 3) = TmpRs!CMPName
       MSHFlexGrid1.TextMatrix(rn, 4) = TmpRs!GodownNo
       MSHFlexGrid1.TextMatrix(rn, 5) = TmpRs!Address
       MSHFlexGrid1.TextMatrix(rn, 6) = TmpRs!Rate
       MSHFlexGrid1.TextMatrix(rn, 7) = TmpRs!BillableQty
       MSHFlexGrid1.TextMatrix(rn, 8) = TmpRs!Amount
       MSHFlexGrid1.TextMatrix(rn, 9) = TmpRs!Remark
       MSHFlexGrid1.TextMatrix(rn, 13) = TmpRs!FumigationDate
       MSHFlexGrid1.TextMatrix(rn, 17) = IIf(IsNull(TmpRs!PrintedDate), TmpRs!FumigationDate, TmpRs!PrintedDate)

       MSHFlexGrid1.row = rn
       MSHFlexGrid1.Col = 0
       MSHFlexGrid1.CellFontName = "Wingdings"
       MSHFlexGrid1.CellFontSize = 12
       MSHFlexGrid1.CellAlignment = flexAlignCenterCenter
       If TmpRs!MappingID > 0 Then
          MSHFlexGrid1.Text = strChecked
          MSHFlexGrid1.TextMatrix(0, 7) = Val(MSHFlexGrid1.TextMatrix(0, 7)) + Val(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 7))
          MSHFlexGrid1.TextMatrix(0, 8) = Val(MSHFlexGrid1.TextMatrix(0, 8)) + Val(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 8))

          MSHFlexGrid1.TextMatrix(rn, 10) = IIf(IsNull(TmpRs!VendorRate), "", TmpRs!VendorRate)
          MSHFlexGrid1.TextMatrix(rn, 11) = IIf(IsNull(TmpRs!VendorQuantity), "", TmpRs!VendorQuantity)

          MSHFlexGrid1.TextMatrix(rn, 12) = Format(Val(MSHFlexGrid1.TextMatrix(rn, 10)) * Val(MSHFlexGrid1.TextMatrix(rn, 11)), "#####0.00")
          MSHFlexGrid1.TextMatrix(0, 11) = Val(MSHFlexGrid1.TextMatrix(0, 11)) + Val(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 11))
          MSHFlexGrid1.TextMatrix(0, 12) = Val(MSHFlexGrid1.TextMatrix(0, 12)) + Val(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 12))
       Else
          MSHFlexGrid1.Text = strUnChecked
       End If

       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
       TmpRs.MoveNext
   Next
End If

Exit Sub
NewDetail:

tmp = "Select TFD.FumigationID, TFD.FumgDetailID,MC.CMPName,MG.GodownNo,MG.Address,MFT.RateperMT,TFD.ActualQty,TFD.Amount,TFD.Remark, TFD.VendorRate,"
tmp = tmp & " TFD.VendorQuantity,TFD.MappingID,TF.FumigationDate, TFD.StackNo, MCo.Commodity,TF.PrintedDate,TFD.LastFumgDate,isnull(Mco.EffectiveDays,0) 'EffectiveDays' from Txn_Fumigation TF"
tmp = tmp & " Inner join Txn_FumigationDetail TFD on TFD.FumigationID = TF.FumigationID"
tmp = tmp & " Left Join Mst_FumigationTreatmentDetail MFT on MFT.DetailID =TFD.AgreementDetailID "
tmp = tmp & " Left join Mst_TreatmentType MTT on MFT.TreatmentTypeID =MTT.TreatmentTypeID"
tmp = tmp & " Left Join Mst_Commodity MCo on MCo.CommodityID = TFD.CommodityID "
tmp = tmp & " Inner join Mst_Godown MG on MG.GodownID=TF.GodownID"
tmp = tmp & " Inner join Mst_CMP MC on MC.CMPID=MG.CMPID"
tmp = tmp & " Where IsNull(TFD.MappingID, 0) = 0 And isnull(TF.BillFlag,'')='' And TF.VendorID = " & mVendorID & " And MC.LocationID = " & mLocationID & " And TF.FumigationDate<= '" & TxtBillDate.Value & "'"
tmp = tmp & " And TFD.BillableFlag = 'Y' And IsNull(TF.MappingID, 0) = 0"
tmp = tmp & " Order by TFD.FumigationID, TFD.FumgDetailID"
Call TmpTable

If TmpRs.RecordCount > 0 Then
   For I = 1 To TmpRs.RecordCount
       rn = MSHFlexGrid1.Rows - 1
       MSHFlexGrid1.TextMatrix(rn, 1) = "Fumigation"
       MSHFlexGrid1.TextMatrix(rn, 2) = IIf(IsNull(TmpRs!FumgDetailID), 0, TmpRs!FumgDetailID)
       MSHFlexGrid1.TextMatrix(rn, 3) = IIf(IsNull(TmpRs!CMPName), "", TmpRs!CMPName)
       MSHFlexGrid1.TextMatrix(rn, 4) = IIf(IsNull(TmpRs!GodownNo), "", TmpRs!GodownNo)
       MSHFlexGrid1.TextMatrix(rn, 5) = IIf(IsNull(TmpRs!Address), "", TmpRs!Address)
       MSHFlexGrid1.TextMatrix(rn, 6) = IIf(IsNull(TmpRs!RateperMT), "", TmpRs!RateperMT)
       MSHFlexGrid1.TextMatrix(rn, 7) = IIf(IsNull(TmpRs!ActualQty), "", TmpRs!ActualQty)
       MSHFlexGrid1.TextMatrix(rn, 8) = IIf(IsNull(TmpRs!Amount), "", TmpRs!Amount)
       MSHFlexGrid1.TextMatrix(rn, 9) = IIf(IsNull(TmpRs!Remark), "", TmpRs!Remark)
       MSHFlexGrid1.TextMatrix(rn, 13) = IIf(IsNull(TmpRs!FumigationDate), "", TmpRs!FumigationDate)
       MSHFlexGrid1.TextMatrix(rn, 14) = IIf(IsNull(TmpRs!StackNo), "", TmpRs!StackNo)
       MSHFlexGrid1.TextMatrix(rn, 15) = IIf(IsNull(TmpRs!Commodity), "", TmpRs!Commodity)
       MSHFlexGrid1.TextMatrix(rn, 16) = IIf(IsNull(TmpRs!FumigationID), "", TmpRs!FumigationID)
       MSHFlexGrid1.TextMatrix(rn, 17) = IIf(IsNull(TmpRs!PrintedDate), TmpRs!FumigationDate, TmpRs!PrintedDate)
       MSHFlexGrid1.TextMatrix(rn, 18) = IIf(IsNull(TmpRs!LastFumgDate), TmpRs!FumigationDate, TmpRs!LastFumgDate)
       
       If DateDiff("d", TmpRs!LastFumgDate, TmpRs!FumigationDate) <= TmpRs!EffectiveDays Then
          For y = 0 To MSHFlexGrid1.Cols - 1
              MSHFlexGrid1.Col = y
              MSHFlexGrid1.row = rn
              MSHFlexGrid1.CellBackColor = vbGreen
          Next
       End If
       
       MSHFlexGrid1.row = rn
       MSHFlexGrid1.Col = 0
       MSHFlexGrid1.CellFontName = "Wingdings"
       MSHFlexGrid1.CellFontSize = 12
       MSHFlexGrid1.CellAlignment = flexAlignCenterCenter
       If TmpRs!MappingID > 0 Then
          MSHFlexGrid1.Text = strChecked
          MSHFlexGrid1.TextMatrix(0, 7) = Val(MSHFlexGrid1.TextMatrix(0, 7)) + Val(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 7))
          MSHFlexGrid1.TextMatrix(0, 8) = Val(MSHFlexGrid1.TextMatrix(0, 8)) + Val(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 8))
          
          MSHFlexGrid1.TextMatrix(rn, 10) = IIf(IsNull(TmpRs!VendorRate), "", TmpRs!VendorRate)
          MSHFlexGrid1.TextMatrix(rn, 11) = IIf(IsNull(TmpRs!VendorQuantity), "", TmpRs!VendorQuantity)
       
          MSHFlexGrid1.TextMatrix(rn, 12) = Format(Val(MSHFlexGrid1.TextMatrix(rn, 10)) * Val(MSHFlexGrid1.TextMatrix(rn, 11)), "#####0.00")
          MSHFlexGrid1.TextMatrix(0, 11) = Val(MSHFlexGrid1.TextMatrix(0, 11)) + Val(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 11))
          MSHFlexGrid1.TextMatrix(0, 12) = Val(MSHFlexGrid1.TextMatrix(0, 12)) + Val(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 12))
       Else
          MSHFlexGrid1.Text = strUnChecked
       End If
       
       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
       TmpRs.MoveNext
   Next
End If
End Sub

Private Sub AddSprayingDetail()

If Format(TxtMappingDate.Value, "yyyy-mm-dd") > "2014-01-17" Then GoTo NewDetail

tmp = "Select TS.SprayingID,MC.CMPName,MG.GodownNo,MG.Address,TS.Rate,TS.ActualQty,TS.Amount,TS.Remark,TS.VendorRate,TS.VendorQuantity,TS.MappingID,TS.SprayingDate,TS.PrintedDate from Txn_Spraying TS"
tmp = tmp & " Inner join Mst_Godown MG on MG.GodownID=TS.GodownID"
tmp = tmp & " Inner join Mst_CMP MC on MC.CMPID=MG.CMPID"
tmp = tmp & " Where IsNull(TS.MappingID, 0) = 0 And isnull(TS.BillFlag,'')='' And TS.VendorID=" & mVendorID & " And MC.LocationID = " & mLocationID & " And TS.SprayingDate<= '" & TxtBillDate.Value & "' And TS.BillableFlag = 'Y' "
tmp = tmp & " Order by TS.MappingID Desc"

Call TmpTable

If TmpRs.RecordCount > 0 Then
   For I = 1 To TmpRs.RecordCount
       rn = MSHFlexGrid1.Rows - 1
       MSHFlexGrid1.TextMatrix(rn, 1) = "Spraying"
       MSHFlexGrid1.TextMatrix(rn, 2) = TmpRs!SprayingID
       MSHFlexGrid1.TextMatrix(rn, 3) = TmpRs!CMPName
       MSHFlexGrid1.TextMatrix(rn, 4) = TmpRs!GodownNo
       MSHFlexGrid1.TextMatrix(rn, 5) = TmpRs!Address
       MSHFlexGrid1.TextMatrix(rn, 6) = TmpRs!Rate
       MSHFlexGrid1.TextMatrix(rn, 7) = TmpRs!ActualQty
       MSHFlexGrid1.TextMatrix(rn, 8) = TmpRs!Amount
       MSHFlexGrid1.TextMatrix(rn, 9) = TmpRs!Remark
       MSHFlexGrid1.TextMatrix(rn, 13) = TmpRs!SprayingDate
       MSHFlexGrid1.TextMatrix(rn, 17) = IIf(IsNull(TmpRs!PrintedDate), TmpRs!SprayingDate, TmpRs!PrintedDate)
       
       MSHFlexGrid1.row = rn
       MSHFlexGrid1.Col = 0
       MSHFlexGrid1.CellFontName = "Wingdings"
       MSHFlexGrid1.CellFontSize = 12
       MSHFlexGrid1.CellAlignment = flexAlignCenterCenter
       If TmpRs!MappingID > 0 Then
          MSHFlexGrid1.Text = strChecked
          MSHFlexGrid1.TextMatrix(0, 7) = Val(MSHFlexGrid1.TextMatrix(0, 7)) + Val(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 7))
          MSHFlexGrid1.TextMatrix(0, 8) = Val(MSHFlexGrid1.TextMatrix(0, 8)) + Val(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 8))
          
          MSHFlexGrid1.TextMatrix(rn, 10) = IIf(IsNull(TmpRs!VendorRate), "", TmpRs!VendorRate)
          MSHFlexGrid1.TextMatrix(rn, 11) = IIf(IsNull(TmpRs!VendorQuantity), "", TmpRs!VendorQuantity)
       
          MSHFlexGrid1.TextMatrix(rn, 12) = Format(Val(MSHFlexGrid1.TextMatrix(rn, 10)) * Val(MSHFlexGrid1.TextMatrix(rn, 11)), "#####0.00")
          MSHFlexGrid1.TextMatrix(0, 11) = Val(MSHFlexGrid1.TextMatrix(0, 11)) + Val(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 11))
          MSHFlexGrid1.TextMatrix(0, 12) = Val(MSHFlexGrid1.TextMatrix(0, 12)) + Val(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 12))
          
       Else
          MSHFlexGrid1.Text = strUnChecked
       End If
              
       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
       TmpRs.MoveNext
   Next
End If

Exit Sub

NewDetail:


tmp = "Select TS.SprayingID,MC.CMPName,MG.GodownNo,MG.Address,TS.Rate,TS.ActualArea,TS.Amount,TS.Remark,TS.VendorRate,TS.VendorQuantity,TS.MappingID,TS.SprayingDate,TS.PrintedDate from Txn_Spraying TS"
tmp = tmp & " Inner join Mst_Godown MG on MG.GodownID=TS.GodownID"
tmp = tmp & " Inner join Mst_CMP MC on MC.CMPID=MG.CMPID"
tmp = tmp & " Where IsNull(TS.MappingID, 0) = 0 And isnull(TS.BillFlag,'')='' And TS.VendorID=" & mVendorID & " And MC.LocationID = " & mLocationID & " And TS.SprayingDate<= '" & TxtBillDate.Value & "' And TS.BillableFlag = 'Y' "
tmp = tmp & " Order by TS.MappingID Desc"

Call TmpTable

If TmpRs.RecordCount > 0 Then
   For I = 1 To TmpRs.RecordCount
       rn = MSHFlexGrid1.Rows - 1
       MSHFlexGrid1.TextMatrix(rn, 1) = "Spraying"
       MSHFlexGrid1.TextMatrix(rn, 2) = TmpRs!SprayingID
       MSHFlexGrid1.TextMatrix(rn, 3) = TmpRs!CMPName
       MSHFlexGrid1.TextMatrix(rn, 4) = TmpRs!GodownNo
       MSHFlexGrid1.TextMatrix(rn, 5) = TmpRs!Address
       MSHFlexGrid1.TextMatrix(rn, 6) = TmpRs!Rate
       MSHFlexGrid1.TextMatrix(rn, 7) = TmpRs!ActualArea 'TmpRs!ActualQty
       MSHFlexGrid1.TextMatrix(rn, 8) = TmpRs!Amount
       MSHFlexGrid1.TextMatrix(rn, 9) = TmpRs!Remark
       MSHFlexGrid1.TextMatrix(rn, 13) = TmpRs!SprayingDate
       MSHFlexGrid1.TextMatrix(rn, 17) = IIf(IsNull(TmpRs!PrintedDate), TmpRs!SprayingDate, TmpRs!PrintedDate)
       
       MSHFlexGrid1.row = rn
       MSHFlexGrid1.Col = 0
       MSHFlexGrid1.CellFontName = "Wingdings"
       MSHFlexGrid1.CellFontSize = 12
       MSHFlexGrid1.CellAlignment = flexAlignCenterCenter
       If TmpRs!MappingID > 0 Then
          MSHFlexGrid1.Text = strChecked
          MSHFlexGrid1.TextMatrix(0, 7) = Val(MSHFlexGrid1.TextMatrix(0, 7)) + Val(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 7))
          MSHFlexGrid1.TextMatrix(0, 8) = Val(MSHFlexGrid1.TextMatrix(0, 8)) + Val(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 8))
          
          MSHFlexGrid1.TextMatrix(rn, 10) = IIf(IsNull(TmpRs!VendorRate), "", TmpRs!VendorRate)
          MSHFlexGrid1.TextMatrix(rn, 11) = IIf(IsNull(TmpRs!VendorQuantity), "", TmpRs!VendorQuantity)
       
          MSHFlexGrid1.TextMatrix(rn, 12) = Format(Val(MSHFlexGrid1.TextMatrix(rn, 10)) * Val(MSHFlexGrid1.TextMatrix(rn, 11)), "#####0.00")
          MSHFlexGrid1.TextMatrix(0, 11) = Val(MSHFlexGrid1.TextMatrix(0, 11)) + Val(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 11))
          MSHFlexGrid1.TextMatrix(0, 12) = Val(MSHFlexGrid1.TextMatrix(0, 12)) + Val(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 12))
          
       Else
          MSHFlexGrid1.Text = strUnChecked
       End If
              
       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
       TmpRs.MoveNext
   Next
End If


End Sub

Private Sub TxtVendorData_LostFocus()
If TxtVendorData.Text = "" Then
   TxtVendorData.Visible = False
   Exit Sub
End If
If IsNumeric(TxtVendorData.Text) = False Then
   MsgBox "Invalid numeric format!!"
   TxtVendorData.SetFocus
   Exit Sub
End If

MSHFlexGrid1.TextMatrix(0, 11) = Val(MSHFlexGrid1.TextMatrix(0, 11)) - Val(MSHFlexGrid1.TextMatrix(mSelRow, 11))
MSHFlexGrid1.TextMatrix(0, 12) = Val(MSHFlexGrid1.TextMatrix(0, 12)) - Val(MSHFlexGrid1.TextMatrix(mSelRow, 12))

MSHFlexGrid1.TextMatrix(mSelRow, mSelCol) = TxtVendorData.Text

MSHFlexGrid1.TextMatrix(mSelRow, 12) = Format(Val(MSHFlexGrid1.TextMatrix(mSelRow, 10)) * Val(MSHFlexGrid1.TextMatrix(mSelRow, 11)), "#####0.00")
MSHFlexGrid1.TextMatrix(0, 11) = Val(MSHFlexGrid1.TextMatrix(0, 11)) + Val(MSHFlexGrid1.TextMatrix(mSelRow, 11))
MSHFlexGrid1.TextMatrix(0, 12) = Val(MSHFlexGrid1.TextMatrix(0, 12)) + Val(MSHFlexGrid1.TextMatrix(mSelRow, 12))
TxtActualPayable = MSHFlexGrid1.TextMatrix(0, 12)
TxtVendorData.Visible = False
End Sub

Private Function CheckGridSelect() As Boolean
Dim Retval As Boolean
Retval = False

For I = 2 To MSHFlexGrid1.Rows - 2
    If MSHFlexGrid1.TextMatrix(I, 0) = strChecked Then
       Retval = True
       Exit For
    End If
Next
CheckGridSelect = Retval
End Function

Private Function GetMaxMappingID() As Double
Dim Retval As Double
tmp = "Select max(MappingID) from Txn_FumigationApproval"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetMaxMappingID = Retval
End Function

Sub DeleteChildTableDetail()

tmp = "Update Txn_Spraying Set MappingID = null,VendorRate=Null,VendorQuantity=Null where MappingID = " & Val(TxtMappingID.Text) & ""
Call TmpTable

If Format(TxtMappingDate.Value, "yyyy-mm-dd") > "2013-04-29" Then GoTo NewDetail

tmp = "Update Txn_Fumigation Set MappingID = null,VendorRate=Null,VendorQuantity=Null where MappingID = " & Val(TxtMappingID.Text) & ""
Call TmpTable

Exit Sub
NewDetail:

tmp = "Update Txn_Fumigation Set MappingID = null,VendorRate=Null,VendorQuantity=Null where MappingID = " & Val(TxtMappingID.Text) & ""
Call TmpTable

tmp = "Update Txn_FumigationDetail Set MappingID = null,VendorRate=Null,VendorQuantity=Null where MappingID = " & Val(TxtMappingID.Text) & ""
Call TmpTable

End Sub

Private Sub SaveFumigationDetails()

Call DeleteChildTableDetail
Call TableTxn_Spraying(" Where MappingID is null And VendorID=" & mVendorID)
Call TableTxn_Fumigation(" Where MappingID is null And VendorID=" & mVendorID)

If Format(TxtMappingDate.Value, "yyyy-mm-dd") > "2013-04-29" Then GoTo NewDetail


For I = 2 To MSHFlexGrid1.Rows - 2
    If MSHFlexGrid1.TextMatrix(I, 0) = strChecked Then
       If MSHFlexGrid1.TextMatrix(I, 1) = "Fumigation" Then
          If RsTxn_Fumigation.RecordCount > 0 Then RsTxn_Fumigation.MoveFirst
          RsTxn_Fumigation.Find "FumigationID=" & Val(MSHFlexGrid1.TextMatrix(I, 2))
          If RsTxn_Fumigation.EOF = False Then
             RsTxn_Fumigation!MappingID = Val(TxtMappingID)
             RsTxn_Fumigation!VendorRate = Val(MSHFlexGrid1.TextMatrix(I, 10))
             RsTxn_Fumigation!VendorQuantity = Val(MSHFlexGrid1.TextMatrix(I, 11))
             RsTxn_Fumigation.Update
             RsTxn_Fumigation.Requery
          End If
       ElseIf MSHFlexGrid1.TextMatrix(I, 1) = "Spraying" Then
          If RsTxn_Spraying.RecordCount > 0 Then RsTxn_Spraying.MoveFirst
          RsTxn_Spraying.Find "SprayingID=" & Val(MSHFlexGrid1.TextMatrix(I, 2))
          If RsTxn_Spraying.EOF = False Then
             RsTxn_Spraying!MappingID = Val(TxtMappingID)
             RsTxn_Spraying!VendorRate = Val(MSHFlexGrid1.TextMatrix(I, 10))
             RsTxn_Spraying!VendorQuantity = Val(MSHFlexGrid1.TextMatrix(I, 11))
             RsTxn_Spraying.Update
             RsTxn_Spraying.Requery
          End If
       End If
    End If
Next

Exit Sub
NewDetail:
Call TableTxn_FumigationDetail
For I = 2 To MSHFlexGrid1.Rows - 2
    If MSHFlexGrid1.TextMatrix(I, 0) = strChecked Then
       If MSHFlexGrid1.TextMatrix(I, 1) = "Fumigation" Then
          If RsTxn_Fumigation.RecordCount > 0 Then RsTxn_Fumigation.MoveFirst
          RsTxn_Fumigation.Find "FumigationID=" & Val(MSHFlexGrid1.TextMatrix(I, 16))
          If RsTxn_Fumigation.EOF = False Then
             RsTxn_Fumigation!MappingID = Val(TxtMappingID)
             ''RsTxn_Fumigation!VendorRate = Val(MSHFlexGrid1.TextMatrix(0, 10))
             ''RsTxn_Fumigation!VendorQuantity = Val(MSHFlexGrid1.TextMatrix(0, 11))
             RsTxn_Fumigation.Update
             RsTxn_Fumigation.Requery
          End If
       
          If RsTxn_FumigationDetail.RecordCount > 0 Then RsTxn_FumigationDetail.MoveFirst
          RsTxn_FumigationDetail.Find "FumgDetailID=" & Val(MSHFlexGrid1.TextMatrix(I, 2))
          If RsTxn_FumigationDetail.EOF = False Then
             RsTxn_FumigationDetail!MappingID = Val(TxtMappingID)
             RsTxn_FumigationDetail!VendorRate = Val(MSHFlexGrid1.TextMatrix(I, 10))
             RsTxn_FumigationDetail!VendorQuantity = Val(MSHFlexGrid1.TextMatrix(I, 11))
             RsTxn_FumigationDetail.Update
          End If
       ElseIf MSHFlexGrid1.TextMatrix(I, 1) = "Spraying" Then
          If RsTxn_Spraying.RecordCount > 0 Then RsTxn_Spraying.MoveFirst
          RsTxn_Spraying.Find "SprayingID=" & Val(MSHFlexGrid1.TextMatrix(I, 2))
          If RsTxn_Spraying.EOF = False Then
             RsTxn_Spraying!MappingID = Val(TxtMappingID)
             RsTxn_Spraying!VendorRate = Val(MSHFlexGrid1.TextMatrix(I, 10))
             RsTxn_Spraying!VendorQuantity = Val(MSHFlexGrid1.TextMatrix(I, 11))
             RsTxn_Spraying.Update
             RsTxn_Spraying.Requery
          End If
       End If
    End If
Next
End Sub
Public Sub Indata()
 
Frame1.Enabled = False
Frame3.Enabled = False
Check1.Enabled = False

CmbStatus.Visible = False
TxtStatus.Visible = True

CmbForwardTo.Visible = False
TxtForwardTo.Visible = True

TxtMappingID.Enabled = True
CmbVendorID.Visible = False
CmbLocationID.Visible = False
CmdVendor.Visible = False
CmdLocation.Visible = False
TxtVendorID.Visible = True
TxtLocationID.Visible = True

''--- Ullhas Patil 25 Oct 2017
CmbBillingState.Visible = False: CmbBillingState = ""
TxtBillingState.Visible = True: TxtBillingState = ""
''-----------------------------

TxtMappingID = RsTxn_FumigationApproval!MappingID
TxtMappingDate.Value = RsTxn_FumigationApproval!MappingDate
TxtMonthYear.Value = IIf(IsNull(RsTxn_FumigationApproval!MonthYear), Date, RsTxn_FumigationApproval!MonthYear)
TxtVendorID.Text = GetVendorName(RsTxn_FumigationApproval!VendorID, "N")
mVendorID = RsTxn_FumigationApproval!VendorID
mLocationID = RsTxn_FumigationApproval!LocationID
TxtLocationID.Text = GetLocationName(RsTxn_FumigationApproval!LocationID)
TxtRemark.Text = RsTxn_FumigationApproval!Remark
TxtBillNo = IIf(IsNull(RsTxn_FumigationApproval!BillNo), "", RsTxn_FumigationApproval!BillNo)
TxtBillDate.Value = IIf(IsNull(RsTxn_FumigationApproval!BillDate), Date, RsTxn_FumigationApproval!BillDate)
TxtPayableAmount.Text = Format(RsTxn_FumigationApproval!PayableAmount, "#####0.00")
TxtActualPayable.Text = Format(RsTxn_FumigationApproval!ActualAmount, "#####0.00")
TxtSTAmount.Text = Format(RsTxn_FumigationApproval!STAmount, "#####0.00")
mForwardToID = IIf(IsNull(RsTxn_FumigationApproval!ForwardTo), 0, RsTxn_FumigationApproval!ForwardTo)
TxtForwardTo = GetEmployeeNonName(Val(mForwardToID))
CmbForwardTo = TxtForwardTo
Call ShowAccountDetails(mVendorID)
mBatchID = IIf(IsNull(RsTxn_FumigationApproval!BatchID), 0, RsTxn_FumigationApproval!BatchID)
mPBatchID = IIf(IsNull(RsTxn_FumigationApproval!PBatchID), 0, RsTxn_FumigationApproval!PBatchID)

CmbModeOfPay.Visible = False
TxtModeOfPayment.Visible = True

CmbFAGBank.Visible = False
TxtFAGBank.Visible = True

TxtFAGDisbAmt = IIf(IsNull(RsTxn_FumigationApproval!DisbuAmount), "", RsTxn_FumigationApproval!DisbuAmount)
TxtFAGRemark = IIf(IsNull(RsTxn_FumigationApproval!DisbuRemark), "", RsTxn_FumigationApproval!DisbuRemark)
TxtFagDisbDate.Value = IIf(IsNull(RsTxn_FumigationApproval!DisbuDate), Date, RsTxn_FumigationApproval!DisbuDate)
mFAGBankID = IIf(IsNull(RsTxn_FumigationApproval!FAGBankID), 0, RsTxn_FumigationApproval!FAGBankID)
TxtFAGBank.Text = GetFAGBankAcNo(Val(mFAGBankID))
TxtModeOfPayment.Text = IIf(IsNull(RsTxn_FumigationApproval!ModeOfPay), "", RsTxn_FumigationApproval!ModeOfPay)
TxtDeductionRemark = IIf(IsNull(RsTxn_FumigationApproval!DeductionRemark), "", RsTxn_FumigationApproval!DeductionRemark)

If RsTxn_FumigationApproval!Flag = "A" Then
    TxtStatus.Text = "Approved"
ElseIf RsTxn_FumigationApproval!Flag = "R" Then
    TxtStatus.Text = "Rejected"
ElseIf RsTxn_FumigationApproval!Flag = "S" Then
    TxtStatus.Text = "Settled"
Else
    TxtStatus.Text = "Pending"
End If
CmbStatus = TxtStatus
Call CmbStatus_LostFocus
CmbSTaxDesc.Visible = False
TxtSTaxDesc.Visible = False
TxtSTaxDesc.Text = ""
If LCase(CmbStatus.Text) = "approved" Then
   TxtSTaxDesc.Visible = True
   TxtSTaxDesc.Text = GetServiceTaxDescription(RsTxn_FumigationApproval!ServiceTaxID)
End If
If LCase(TxtStatus.Text) <> "pending" Then
   lblFAGPostingDate.Visible = True
   TxtFAGPostingDate.Visible = True
   lblFAGRemarks.Visible = True
   TxtFAGRemarks.Visible = True
   TxtFAGPostingDate.Value = IIf(IsNull(RsTxn_FumigationApproval!ApprovedDate), Date, RsTxn_FumigationApproval!ApprovedDate)
   TxtFAGRemarks = IIf(IsNull(RsTxn_FumigationApproval!ApprovalRemark), "", RsTxn_FumigationApproval!ApprovalRemark)
Else
   lblFAGPostingDate.Visible = False
   TxtFAGPostingDate.Visible = False
   lblFAGRemarks.Visible = False
   TxtFAGRemarks.Visible = False
End If
TxtBatchID = IIf(IsNull(RsTxn_FumigationApproval!BatchID), "", RsTxn_FumigationApproval!BatchID)

''--- Ullhas Patil 25 Oct 2017
mBillingStateID = IIf(IsNull(RsTxn_FumigationApproval!BillingStateIndiaID), 0, RsTxn_FumigationApproval!BillingStateIndiaID)
TxtBillingState = GetStateIndiaName(mBillingStateID)
TxtBillingState = IIf(Trim(TxtBillingState) = "0", "", TxtBillingState)
''-----------------------------

TxtCreated = IIf(IsNull(RsTxn_FumigationApproval!CreatedBy), "", RsTxn_FumigationApproval!CreatedBy) & " :- " & IIf(IsNull(RsTxn_FumigationApproval!CreatedDate), "", RsTxn_FumigationApproval!CreatedDate)
TxtUpdated = IIf(IsNull(RsTxn_FumigationApproval!UpdatedBy), "", RsTxn_FumigationApproval!UpdatedBy) & " :- " & IIf(IsNull(RsTxn_FumigationApproval!UpdatedDate), "", RsTxn_FumigationApproval!UpdatedDate)

''CmdPrintPreview.Enabled = True
CmdPrint.Enabled = True
Call Grid_Head
Call InDataFumigationDetail
Call InDataSprayingDetail
Call InDataLogHistory
Call GButtonLock(Me, True)

End Sub


Private Sub ShowAccountDetails(mVendorID_ As Variant)
tmp = "Select BankID ,BranchID,BankAcNo,IFSCCode from Mst_Vendor Where VendorID=" & mVendorID_
Call TmpTable
If TmpRs.RecordCount > 0 Then
   TxtBankName.Text = GetBankName(IIf(IsNull(TmpRs!BankID), 0, TmpRs!BankID), "N")
   TxtBranchName = GetBranchName(IIf(IsNull(TmpRs!BranchID), 0, TmpRs!BranchID), "N")  ' IIf(IsNull(TmpRs!BankBranch), "", TmpRs!BankBranch)
   TxtACNo = IIf(IsNull(TmpRs!BankAcNo), "", TmpRs!BankAcNo)
   TxtIFSCCode = IIf(IsNull(TmpRs!IFSCCode), "", TmpRs!IFSCCode)
End If
End Sub
Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsTxn_FumigationApproval.MoveNext
If RsTxn_FumigationApproval.EOF Then
   RsTxn_FumigationApproval.MoveLast
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
RsTxn_FumigationApproval.MovePrevious
If RsTxn_FumigationApproval.BOF Then
   RsTxn_FumigationApproval.MoveFirst
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

Private Sub LastCmd_Click()
RsTxn_FumigationApproval.MoveLast
Call Indata
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False
End Sub
Private Sub FirstCmd_Click()
RsTxn_FumigationApproval.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub


Private Sub CmbSLocationID_GotFocus()
tmp = CmbSLocationID.Text
If Gen_WcLocation <> "" Then
   Call TableMst_Location(" Where " & Gen_WcLocation)
Else
   Call TableMst_Location '(" Where LocationID <> 0 ")
End If
CmbSLocationID.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbSLocationID.AddItem RsMst_Location!LocationName
    RsMst_Location.MoveNext
Next
CmbSLocationID.Text = tmp
End Sub

Private Sub CmbSLocationID_LostFocus()
If CmbSLocationID.Text = "" Then
   mLocationID = 0
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbSLocationID.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection "
   CmbSLocationID.SetFocus
   Exit Sub
End If

mSLocationID = RsMst_Location!LocationID
End Sub

Private Sub Grid_Head1()

MSHFlexGrid2.Clear
MSHFlexGrid2.Rows = 3
MSHFlexGrid2.FixedRows = 2
MSHFlexGrid2.Cols = 23
MSHFlexGrid2.WordWrap = True

MSHFlexGrid2.WordWrap = True
MSHFlexGrid2.TextMatrix(1, 0) = "Mapping ID"
MSHFlexGrid2.TextMatrix(1, 1) = "Mapping Date"
MSHFlexGrid2.TextMatrix(1, 2) = "Month Year"
MSHFlexGrid2.TextMatrix(1, 3) = "Vendor Name"
MSHFlexGrid2.TextMatrix(1, 4) = "Location"
MSHFlexGrid2.TextMatrix(1, 5) = "Actual Qty"
MSHFlexGrid2.TextMatrix(1, 6) = "Actual Amount"
MSHFlexGrid2.TextMatrix(1, 7) = "Payable Amount"
MSHFlexGrid2.TextMatrix(1, 8) = "Bill No"
MSHFlexGrid2.TextMatrix(1, 9) = "Remark"
MSHFlexGrid2.TextMatrix(1, 10) = "Forwarded To"
MSHFlexGrid2.TextMatrix(1, 11) = "Status"
MSHFlexGrid2.TextMatrix(1, 12) = "FAG Remarks"
MSHFlexGrid2.TextMatrix(1, 13) = "FAG Posting Date"
MSHFlexGrid2.TextMatrix(1, 14) = "Batch ID"
MSHFlexGrid2.TextMatrix(1, 15) = "Bank Name"
MSHFlexGrid2.TextMatrix(1, 16) = "Cheque No"
MSHFlexGrid2.TextMatrix(1, 17) = "Paid Amount"
MSHFlexGrid2.TextMatrix(1, 18) = "TDS"
MSHFlexGrid2.TextMatrix(1, 19) = "Paid Date"
MSHFlexGrid2.TextMatrix(1, 20) = "Paymentdetails Updated By"
MSHFlexGrid2.TextMatrix(1, 21) = "Paymentdetails Updated On"
MSHFlexGrid2.TextMatrix(1, 22) = "Billing State India"

MSHFlexGrid2.ColWidth(0) = 1000
MSHFlexGrid2.ColWidth(1) = 1300
MSHFlexGrid2.ColWidth(2) = 1500
MSHFlexGrid2.ColWidth(3) = 2200
MSHFlexGrid2.ColWidth(4) = 2200
MSHFlexGrid2.ColWidth(5) = 1200
MSHFlexGrid2.ColWidth(6) = 1200
MSHFlexGrid2.ColWidth(7) = 1200
MSHFlexGrid2.ColWidth(8) = 1200
MSHFlexGrid2.ColWidth(9) = 1500
MSHFlexGrid2.ColWidth(10) = 1600
MSHFlexGrid2.ColWidth(11) = 1400
MSHFlexGrid2.ColWidth(12) = 1300
MSHFlexGrid2.ColWidth(13) = 1600
MSHFlexGrid2.ColWidth(14) = 800
MSHFlexGrid2.ColWidth(15) = 1600
MSHFlexGrid2.ColWidth(16) = 1600
MSHFlexGrid2.ColWidth(17) = 1600
MSHFlexGrid2.ColWidth(18) = 1600
MSHFlexGrid2.ColWidth(19) = 1600
MSHFlexGrid2.ColWidth(20) = 1600
MSHFlexGrid2.ColWidth(21) = 1600
MSHFlexGrid2.ColWidth(22) = 1600

MSHFlexGrid2.RowHeight(1) = 800

End Sub

Private Function CheckGridData() As Boolean
Dim Retval As Boolean
Retval = True

For I = 2 To MSHFlexGrid1.Rows - 2
    If MSHFlexGrid1.TextMatrix(I, 0) = strChecked Then
       If Val(MSHFlexGrid1.TextMatrix(I, 10)) <= 0 Or Val(MSHFlexGrid1.TextMatrix(I, 11)) <= 0 Then
          Retval = False
          Exit For
       End If
    End If
Next
CheckGridData = Retval
End Function


Private Sub InDataLogHistory()

Call Grid_HeadLog

tmp = "Select MappingID ,OldActualAmt,NewActualAmt, OldForwardToID, NewForwardToID, UpdatedBy, Concurrency"
tmp = tmp & " From Log_FumigationApproval Where MappingID=" & Val(TxtMappingID) & ""

Call TmpTable
MSHFlexGrid3.Rows = MSHFlexGrid3.Rows + TmpRs.RecordCount
For I = 1 To MSHFlexGrid3.Rows - 2
   mTmpEmpName = ""
   MSHFlexGrid3.TextMatrix(I, 0) = IIf(IsNull(TmpRs!MappingID), 0, TmpRs!MappingID)
   MSHFlexGrid3.TextMatrix(I, 1) = IIf(IsNull(TmpRs!OldActualAmt), 0, TmpRs!OldActualAmt)
   MSHFlexGrid3.TextMatrix(I, 2) = IIf(IsNull(TmpRs!NewActualAmt), 0, TmpRs!NewActualAmt)
   mTmpEmpName = GetEmployeeNonName(IIf(IsNull(TmpRs!OldForwardToID), "", TmpRs!OldForwardToID))
   MSHFlexGrid3.TextMatrix(I, 3) = mTmpEmpName
   mTmpEmpName = GetEmployeeNonName(IIf(IsNull(TmpRs!NewForwardToID), 0, TmpRs!NewForwardToID))
   MSHFlexGrid3.TextMatrix(I, 4) = mTmpEmpName

   MSHFlexGrid3.TextMatrix(I, 5) = IIf(IsNull(TmpRs!UpdatedBy), "", TmpRs!UpdatedBy)
   MSHFlexGrid3.TextMatrix(I, 6) = Format(IIf(IsNull(TmpRs!Concurrency), "", TmpRs!Concurrency), "dd-mmm-yyyy hh:mm:ss")
   TmpRs.MoveNext
Next
End Sub

Private Sub Gen_Email()
Dim TOADD As Variant
Dim TOCC, mMsg, mGridMsg As Variant
Dim strSplitString() As String
Dim EmpNo_, EmpName_ As String

mMsg = ""
mMsg = "<tr> Fumigation Expense (From Store-Man)</tr>"
mMsg = mMsg & "<tr><td> Mapping ID </td><td> " & TxtMappingID & " </td></tr>"
mMsg = mMsg & "<tr><td> Mapping Date </td><td> " & Format(TxtMappingDate.Value, "dd-mmm-yyyy") & "  </td></tr>"
mMsg = mMsg & "<tr><td> Status </td><td> " & TxtStatus.Text & "  </td></tr>"
mMsg = mMsg & "<tr><td> Location </td><td> " & TxtLocationID.Text & "  </td></tr>"
mMsg = mMsg & "<tr><td> Vendor</td><td> " & TxtVendorID.Text & "  </td></tr>"
mMsg = mMsg & "<tr><td> Entered By </td><td> " & TxtCreated & "  </td></tr>"
If Edit_Flg = "E" Then
   mMsg = mMsg & "<tr><td> Updated By </td><td> " & TxtUpdated & "  </td></tr>"
End If
mMsg = mMsg & "<tr><td> Actual Amount </td><td> " & TxtActualPayable & "  </td></tr>"
mMsg = mMsg & "<tr><td> ST Amount </td><td> " & TxtSTAmount & "  </td></tr>"
mMsg = mMsg & "<tr><td> Payable Amount </td><td> " & TxtPayableAmount & "  </td></tr>"

mMsg = mMsg & "<tr><td> Bill No </td><td> " & TxtBillNo & "  </td></tr>"
mMsg = mMsg & "<tr><td> Remark </td><td> " & TxtRemark & "   </td></tr>"

strSplitString = Split(TxtForwardTo.Text, "~")
EmpName_ = Trim(strSplitString(0))
EmpNo_ = Trim(strSplitString(1))
mMsg = mMsg & "<tr><td> Forwarded To </td><td> " & EmpName_ & " " & " ~ " & " " & EmpNo_ & " </td></tr>"

If LCase(TxtStatus.Text) <> "pending" Then
   mMsg = mMsg & "<tr><td> FAG Remarks </td><td> " & TxtFAGRemarks & " </td></tr>"
   mMsg = mMsg & "<tr><td> FAG Posting Date </td><td> " & Format(TxtFAGPostingDate.Value, "dd-mmm-yyyy") & " </td></tr>"
End If
For I = 1 To MSHFlexGrid1.Rows - 1
    If I = 1 Then
       mGridMsg = mGridMsg & "<tr><td> " & MSHFlexGrid1.TextMatrix(I, 1) & " </td><td> " & MSHFlexGrid1.TextMatrix(I, 2) & " </td><td> " & MSHFlexGrid1.TextMatrix(I, 3) & " </td><td> " & MSHFlexGrid1.TextMatrix(I, 4) & " </td><td> " & MSHFlexGrid1.TextMatrix(I, 5) & " </td><td> " & MSHFlexGrid1.TextMatrix(I, 6) & " </td><td> " & MSHFlexGrid1.TextMatrix(I, 7) & " </td><td> " & MSHFlexGrid1.TextMatrix(I, 8) & " </td><td> " & MSHFlexGrid1.TextMatrix(I, 9) & " </td><td> " & MSHFlexGrid1.TextMatrix(I, 10) & " </td><td> " & MSHFlexGrid1.TextMatrix(I, 11) & " </td>"
       If Format(TxtMappingDate.Value, "yyyy-mm-dd") > "2013-04-29" Then
            mGridMsg = mGridMsg & "<td> " & MSHFlexGrid1.TextMatrix(I, 12) & " </td><td> " & MSHFlexGrid1.TextMatrix(I, 14) & " </td><td> " & MSHFlexGrid1.TextMatrix(I, 15) & " </td><td> " & MSHFlexGrid1.TextMatrix(I, 16) & " </td>"
       End If
       mGridMsg = mGridMsg & "</tr>"
    Else
       If MSHFlexGrid1.TextMatrix(I, 0) = strChecked Then
          mGridMsg = mGridMsg & "<tr><td> " & MSHFlexGrid1.TextMatrix(I, 1) & " </td><td> " & MSHFlexGrid1.TextMatrix(I, 2) & " </td><td> " & MSHFlexGrid1.TextMatrix(I, 3) & " </td><td> " & MSHFlexGrid1.TextMatrix(I, 4) & " </td><td> " & MSHFlexGrid1.TextMatrix(I, 5) & " </td><td> " & MSHFlexGrid1.TextMatrix(I, 6) & " </td><td> " & MSHFlexGrid1.TextMatrix(I, 7) & " </td><td> " & MSHFlexGrid1.TextMatrix(I, 8) & " </td><td> " & MSHFlexGrid1.TextMatrix(I, 9) & " </td><td> " & MSHFlexGrid1.TextMatrix(I, 10) & " </td><td> " & MSHFlexGrid1.TextMatrix(I, 11) & " </td>"
          If Format(TxtMappingDate.Value, "yyyy-mm-dd") > "2013-04-29" Then
             mGridMsg = mGridMsg & "<td> " & MSHFlexGrid1.TextMatrix(I, 12) & " </td><td> " & MSHFlexGrid1.TextMatrix(I, 14) & " </td><td> " & MSHFlexGrid1.TextMatrix(I, 15) & " </td><td> " & MSHFlexGrid1.TextMatrix(I, 16) & " </td>"
          End If
          mGridMsg = mGridMsg & "</tr>"
       End If
    End If
Next
mMsg = mMsg & "<tr><td> Details </td><td><table border = 1>" & mGridMsg & "</table></td>"

TOCC = Gen_UserEmailID '& "," & mForwardToEmailID

If LCase(TxtStatus.Text) = "pending" Then
   TOADD = mForwardToEmailID
   mMsgSubject = "Fumigation Pending"
ElseIf LCase(TxtStatus.Text) = "approved" Then
   If LCase(TOCC) <> LCase(mForwardToEmailID) Then
      TOCC = TOCC & "," & mForwardToEmailID
   End If
   mMsgSubject = "Fumigation Approved"
ElseIf LCase(TxtStatus.Text) = "rejected" Then
   If LCase(TOCC) <> LCase(mForwardToEmailID) Then
      TOCC = TOCC & "," & mForwardToEmailID
   End If
   mMsgSubject = "Fumigation Rejected"
End If

Call SendMail(TOADD, TOCC, mMsgSubject, mMsg)

End Sub
Private Function SendMail(ToADD_ As Variant, TOCC_ As Variant, MsgSubject_ As Variant, Message1_ As Variant)
Dim htmlStringhead As String
Dim htmlStringtail As String
Set oSMTP = New OSSMTP.SMTPSession

servertxt = "10.0.0.13"

htmlStringhead = "<html><body><font size= 3 face= Zurich BT color = Blue><table border = 1>"
htmlStringtail = "</table></Font></body></html>"
   
With oSMTP
  'connection
  .Server = servertxt
  'message
  If Edit_Flg = "A" Then
     .MailFrom = mEmployeeEmailID
  Else
     .MailFrom = Gen_UserEmailID
  End If
  If IsNull(Trim(ToADD_)) = True Or Trim(ToADD_) = "" Then
     ToADD_ = "manu.ag@ncml.com"
  End If
  .SendTo = ToADD_
  If ToADD_ <> TOCC_ Then
     .CC = TOCC_
  End If
  '.BCC = "mayur.d@ncml.com"
  .MessageSubject = MsgSubject_ & " (from Store-Man)" '"Confidential Mail"
  .MessageHTML = htmlStringhead & Message1_ & htmlStringtail
  
  '.MailFrom = "rakesh.go@ncml.com"
  '.SendTo = "parag.r@ncml.com"
  '.CC = "rakesh.go@ncml.com"
  .TimeStamp = Now()
  If LCase(Gen_EmailConfigFlag) = "y" Then
     .SendEmail
  End If
End With

End Function

Private Sub CmbSTaxDesc_GotFocus()
tmp = CmbSTaxDesc.Text
Call TableMst_ServiceTax
CmbSTaxDesc.Clear
If RsMst_ServiceTax.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_ServiceTax.RecordCount
    CmbSTaxDesc.AddItem RsMst_ServiceTax!Description
    RsMst_ServiceTax.MoveNext
Next
CmbSTaxDesc.Text = tmp
End Sub

Private Sub CmbSTaxDesc_LostFocus()
If CmbSTaxDesc.Text = "" Then
   mServiceTaxID = Null
   TxtServiceTaxNAVCode = ""
   Exit Sub
End If
RsMst_ServiceTax.MoveFirst
RsMst_ServiceTax.Find "Description = '" & CmbSTaxDesc.Text & "'"
If RsMst_ServiceTax.EOF Then
   MsgBox "Invalid selection "
   CmbSTaxDesc.SetFocus
   Exit Sub
End If
mServiceTaxID = RsMst_ServiceTax!ServiceTaxID
End Sub

Private Sub CmdPrint_Click()

Gen_ExcelPassword = "1234"
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Fumigation" & Trim(TxtMappingID) & ".xls")
FileName = Gen_mTimeStamp & Gen_UserName & "Fumigation" & Trim(TxtMappingID) & ".xls"

Call Xls_Main_Rpt

End Sub

Public Sub Xls_Main_Rpt()

Dim ans As Variant
MsgBox ("Wait till Next Message")
Set oXL = New Excel.Application
'Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & FileName)
Set oWS = oWB.Worksheets("Sheet1")
oWS.Name = "Bill Voucher"
'For I = 2 To MSHFlexGrid2.Rows - 2

    'oWS.PageSetup.RightMargin = 0.05
'    oWS.Name = MSHFlexGrid2.TextMatrix(I, 0) & "~" & MSHFlexGrid2.TextMatrix(I, 2)
    oWS.ClearArrows
    mRow = 1
    
    'oWS.Range("a1:i1").Merge
    
    oWS.Range("a:a").ColumnWidth = 10.86 '.157 '10.86
    oWS.Range("b:b").ColumnWidth = 10.86 '17.71
    oWS.Range("c:c").ColumnWidth = 13.57 '17.14
    oWS.Range("d:d").ColumnWidth = 8.29 '12.86
    oWS.Range("e:e").ColumnWidth = 8.57 '33 '32.14
    oWS.Range("f:f").ColumnWidth = 8
    oWS.Range("g:g").ColumnWidth = 7.86
    oWS.Range("h:h").ColumnWidth = 8.57
    oWS.Range("i:i").ColumnWidth = 10.14
      
    'oWS.Range("A1:E1").Borders.Color = RGB(0, 0, 0)
    tmp = "a" & mRow & ":" & "i" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 1) = "NATIONAL COLLATERAL MANAGEMENT SERVICES LIMITED"
    oWS.Cells(mRow, C + 1).HorizontalAlignment = -4108
    oWS.Cells(mRow, C + 1).Font.Bold = True
    oWS.Cells(mRow, C + 1).Font.Size = 10
    'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
    
    'oWS.Cells(1, 1).Font.ColorIndex = 4
    mRow = mRow + 1
    tmp = "a" & mRow & ":" & "i" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 1) = "BILL VOUCHER"
    oWS.Cells(mRow, C + 1).Font.Bold = True
    oWS.Cells(mRow, C + 1).Font.Size = 10
    oWS.Cells(mRow, C + 1).HorizontalAlignment = -4108
    'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
    
    mRow = mRow + 1
    tmp = "a" & mRow & ":" & "b" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 1) = "Business Segment :"
    oWS.Cells(mRow, C + 1).Font.Bold = True
    oWS.Cells(mRow, C + 1).Font.Size = 10
    'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
    
    tmp = "c" & mRow & ":" & "d" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 3) = "CMP"
    oWS.Cells(mRow, C + 3).Font.Size = 10
    oWS.Cells(mRow, C + 3).HorizontalAlignment = -4108
    'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
    
    tmp = "e" & mRow & ":" & "f" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 5) = "Store-Man ID"
    oWS.Cells(mRow, C + 5).Font.Bold = True
    oWS.Cells(mRow, C + 5).Font.Size = 10
    'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
    
    tmp = "g" & mRow & ":" & "g" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 7) = TxtMappingID
    oWS.Cells(mRow, C + 7).Font.Bold = True
    oWS.Cells(mRow, C + 7).Font.Size = 10
    'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
       
    tmp = "h" & mRow & ":" & "h" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 8) = "Claim Date"
    oWS.Cells(mRow, C + 8).Font.Bold = True
    oWS.Cells(mRow, C + 8).Font.Size = 10
       
    tmp = "i" & mRow & ":" & "i" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 9) = Format(TxtMappingDate, "dd-mmm-yyyy")
    oWS.Cells(mRow, C + 9).Font.Bold = True
    oWS.Cells(mRow, C + 9).Font.Size = 10
       
       
    mRow = mRow + 1
    tmp = "a" & mRow & ":" & "b" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 1) = "Vendor Name :"
    oWS.Cells(mRow, C + 1).Font.Bold = True
    oWS.Cells(mRow, C + 1).Font.Size = 10
    'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
    
    tmp = "c" & mRow & ":" & "i" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 3) = TxtVendorID
    oWS.Cells(mRow, C + 3).Font.Bold = True
    oWS.Cells(mRow, C + 3).Font.Size = 10
    'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
       
    mRow = mRow + 1
    tmp = "a" & mRow & ":" & "b" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 1) = "Location:"
    oWS.Cells(mRow, C + 1).Font.Bold = True
    oWS.Cells(mRow, C + 1).Font.Size = 10
    'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
    
    tmp = "c" & mRow & ":" & "i" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 3) = TxtLocationID
    oWS.Cells(mRow, C + 3).Font.Bold = True
    oWS.Cells(mRow, C + 3).Font.Size = 10
    'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
    
    mRow = mRow + 1
    tmp = "a" & mRow & ":" & "b" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 1) = "Month/Year"
    oWS.Cells(mRow, C + 1).Font.Bold = True
    oWS.Cells(mRow, C + 1).Font.Size = 10
    'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
    
    tmp = "c" & mRow & ":" & "i" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 3) = "Expense For " & Format(TxtMonthYear, "MMM-yyyy") & "" 'To " & Format(TxtBillDate, "MMM-yyyy") & " "
    oWS.Cells(mRow, C + 3).Font.Bold = True
    oWS.Cells(mRow, C + 3).Font.Size = 10
    'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
    
    mRow = mRow + 1
    tmp = "a" & mRow & ":" & "b" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 1) = "Invoice No"
    oWS.Cells(mRow, C + 1).Font.Bold = True
    oWS.Cells(mRow, C + 1).Font.Size = 10
    'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
    
    tmp = "c" & mRow & ":" & "d" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 3) = TxtBillNo
    oWS.Cells(mRow, C + 3).Font.Size = 10
    'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
    
    tmp = "e" & mRow & ":" & "f" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 5) = "Invoice Date"
    oWS.Cells(mRow, C + 5).Font.Bold = True
    oWS.Cells(mRow, C + 5).Font.Size = 10
    'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
    
    tmp = "g" & mRow & ":" & "h" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 7) = "'" & Format(TxtBillDate, "dd-MMM-yyyy")
    oWS.Cells(mRow, C + 7).Font.Size = 10
    'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
    
    mRow = mRow + 1
    tmp = "a" & mRow & ":" & "i" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 1) = "Particulars of Expenditure : "
    oWS.Cells(mRow, C + 1).Font.Bold = True
    oWS.Cells(mRow, C + 1).Font.Size = 10
    'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
    
    mRow = mRow + 1
    tmp = "a" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 1) = "Fumigation / Spraying ID"
    oWS.Cells(mRow, C + 1).Font.Bold = True
    oWS.Cells(mRow, C + 1).Font.Size = 10
    oWS.Cells(mRow, C + 1).RowHeight = 27.75
    oWS.Cells(mRow, C + 1).WrapText = True
    'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
    
    tmp = "b" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 2) = "Fumigation / Spraying Date"
    oWS.Cells(mRow, C + 2).Font.Bold = True
    oWS.Cells(mRow, C + 2).Font.Size = 10
    oWS.Cells(mRow, C + 2).RowHeight = 27.75
    oWS.Cells(mRow, C + 2).WrapText = True
    'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
    
    tmp = "c" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 3) = "Expense Type"
    oWS.Cells(mRow, C + 3).Font.Bold = True
    oWS.Cells(mRow, C + 3).Font.Size = 10
    oWS.Cells(mRow, C + 3).RowHeight = 27.75
    'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
    
    tmp = "d" & mRow & ":" & "g" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 4) = "Particulars"
    oWS.Cells(mRow, C + 4).Font.Bold = True
    oWS.Cells(mRow, C + 4).Font.Size = 10
    oWS.Cells(mRow, C + 4).RowHeight = 27.75
    'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
    
    tmp = "h" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 8) = "Quantity"
    oWS.Cells(mRow, C + 8).Font.Bold = True
    oWS.Cells(mRow, C + 8).Font.Size = 10
    oWS.Cells(mRow, C + 8).RowHeight = 27.75
    'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
    
    tmp = "i" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 9) = "Amount (Without ST)"
    oWS.Cells(mRow, C + 9).Font.Bold = True
    oWS.Cells(mRow, C + 9).Font.Size = 10
    oWS.Cells(mRow, C + 9).RowHeight = 27.75
    oWS.Cells(mRow, C + 9).WrapText = True
    'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
    
''    tmp = "h" & mRow
''    oWS.Range(tmp).Merge
''    oWS.Cells(mRow, c + 8) = "Service Tax"
''    oWS.Cells(mRow, c + 8).Font.Bold = True
''    oWS.Cells(mRow, c + 8).Font.Size = 10
''    oWS.Cells(mRow, c + 8).RowHeight = 27.75
''    'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
    
''    tmp = "i" & mRow
''    oWS.Range(tmp).Merge
''    oWS.Cells(mRow, c + 9) = "Total Amount"
''    oWS.Cells(mRow, c + 9).Font.Bold = True
''    oWS.Cells(mRow, c + 9).Font.Size = 10
''    oWS.Cells(mRow, c + 9).WrapText = True
''    oWS.Cells(mRow, c + 9).RowHeight = 27.75
''    'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
    
    For I = 2 To MSHFlexGrid1.Rows - 2
        If MSHFlexGrid1.TextMatrix(I, 0) = strChecked Then
           mRow = mRow + 1 ' + 6
               C = 0
                    'MSHFlexGrid2.TextMatrix(1, 0) = "Row ID"
                    'MSHFlexGrid2.TextMatrix(1, 1) = "Location"
                    'MSHFlexGrid2.TextMatrix(1, 2) = "Expense Type"
                    'MSHFlexGrid2.TextMatrix(1, 3) = "Particulars"
                    'MSHFlexGrid2.TextMatrix(1, 4) = "Amount"
                    'MSHFlexGrid2.TextMatrix(1, 5) = "Service Tax"
                    'MSHFlexGrid2.TextMatrix(1, 6) = "Total Amount"
                    'MSHFlexGrid2.TextMatrix(1, 7) = "ExpenseTypeID"
                    'MSHFlexGrid2.TextMatrix(1, 8) = "LocationID"
                    'MSHFlexGrid2.TextMatrix(1, 9) = "Invoice Date"
                    'MSHFlexGrid2.TextMatrix(1, 10) = "Quantity"

               For j = 1 To 9
                   If j <> 5 Then
                       oWS.Cells(mRow, C + j).Font.Size = 10
                       oWS.Cells(mRow, C + j).WrapText = True
                   End If
               Next j
               
               
               oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, 2)
               If MSHFlexGrid1.TextMatrix(I, 13) <> "" Then
                    oWS.Cells(mRow, C + 2) = Format(MSHFlexGrid1.TextMatrix(I, 13), "dd-mmm-yyyy")
               End If
               
               oWS.Cells(mRow, C + 3) = MSHFlexGrid1.TextMatrix(I, 1) ' Format(MSHFlexGrid2.TextMatrix(I, c + 3), "dd-mmm-yyyy")
            
               l = Len(MSHFlexGrid1.TextMatrix(I, 9))
                m = 1
Remainder:
               If l > 30 Then
                   l = l - 30
                   m = m + 1
                   GoTo Remainder
               End If
               If l <> 0 Then
                   oWS.Cells(mRow, C + 4).RowHeight = 15 * m + 5
               Else
                   oWS.Cells(mRow, C + 4).RowHeight = 15 * m
               End If
               
               tmp = "d" & mRow & ":g" & mRow
               oWS.Range(tmp).Merge
               oWS.Cells(mRow, C + 4) = MSHFlexGrid1.TextMatrix(I, 9)
              
               oWS.Cells(mRow, C + 8) = MSHFlexGrid1.TextMatrix(I, 11)
               oWS.Cells(mRow, C + 9) = MSHFlexGrid1.TextMatrix(I, 12)
               'oWS.Cells(mRow, c + 8) = MSHFlexGrid1.TextMatrix(I, 5)
               'oWS.Cells(mRow, c + 9) = MSHFlexGrid1.TextMatrix(I, 12)
          End If
    Next I
    
    mRow = mRow + 1
    tmp = "a" & mRow & ":" & "h" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 1) = "Total Amount"
    oWS.Cells(mRow, C + 1).Font.Bold = True
    oWS.Cells(mRow, C + 1).Font.Size = 10
    oWS.Cells(mRow, C + 1).HorizontalAlignment = -4108
    'oWS.Cells(mRow, c + j).Font.Size = 10
    
    tmp = "i" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 9) = Format(MSHFlexGrid1.TextMatrix(0, 12), "#####0.00")
    oWS.Cells(mRow, C + 9).Font.Bold = True
    oWS.Cells(mRow, C + 9).Font.Size = 10
    'oWS.Cells(mRow, c + j).Font.Size = 10
        
    mRow = mRow + 1
    tmp = "a" & mRow & ":" & "h" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 1) = "ST Amount"
    oWS.Cells(mRow, C + 1).Font.Bold = True
    oWS.Cells(mRow, C + 1).Font.Size = 10
    oWS.Cells(mRow, C + 1).HorizontalAlignment = -4108
    'oWS.Cells(mRow, c + j).Font.Size = 10
    
    tmp = "i" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 9) = Format(TxtSTAmount.Text, "#####0.00")
    oWS.Cells(mRow, C + 9).Font.Bold = True
    oWS.Cells(mRow, C + 9).Font.Size = 10
    'oWS.Cells(mRow, c + j).Font.Size = 10
        
    mRow = mRow + 1
    tmp = "a" & mRow & ":" & "h" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 1) = "Total Payable"
    oWS.Cells(mRow, C + 1).Font.Bold = True
    oWS.Cells(mRow, C + 1).Font.Size = 10
    oWS.Cells(mRow, C + 1).HorizontalAlignment = -4108
    'oWS.Cells(mRow, c + j).Font.Size = 10
    
    tmp = "i" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 9) = Format(MSHFlexGrid1.TextMatrix(0, 12) + Val(TxtSTAmount), "#####0.00")
    oWS.Cells(mRow, C + 9).Font.Bold = True
    oWS.Cells(mRow, C + 9).Font.Size = 10
    'oWS.Cells(mRow, c + j).Font.Size = 10
        
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
    oWS.Cells(mRow, C + 7) = "" '"Mr Sanjeev Chhabra"
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
    oWS.Cells(mRow, C + 9) = "Head S&P"
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
    oWS.Cells(mRow, C + 1) = "BILL VOUCHER"
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
    oWS.Cells(mRow, C + 8) = TxtMappingID
    oWS.Cells(mRow, C + 8).Font.Bold = True
    oWS.Cells(mRow, C + 8).Font.Size = 10
    'oWS.Cells(mRow, c + j).Font.Size = 10
       
    mRow = mRow + 1
    tmp = "a" & mRow & ":" & "b" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 1) = "Vendor Name :"
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
    oWS.Cells(mRow, C + 3) = "Expense From " & Format(TxtBillDate, "MMM-yyyy") & " To " & Format(TxtBillDate, "MMM-yyyy") & " "
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


Private Sub InDataFumigationDetail()

If Format(TxtMappingDate.Value, "yyyy-mm-dd") > "2013-04-29" Then GoTo NewDetail

tmp = "Select TF.FumigationID,MC.CMPName,MG.GodownNo,MG.Address,TF.Rate,TF.BillableQty,TF.Amount,TF.Remark,TF.VendorRate,TF.VendorQuantity,TF.MappingID,TF.FumigationDate,TF.PrintedDate from Txn_Fumigation TF"
tmp = tmp & " Inner join Mst_Godown MG on MG.GodownID=TF.GodownID"
tmp = tmp & " Inner join Mst_CMP MC on MC.CMPID=MG.CMPID"
tmp = tmp & " Where TF.MappingID=" & Val(TxtMappingID)
tmp = tmp & " Order by TF.FumigationID"

Call TmpTable

If TmpRs.RecordCount > 0 Then
   For I = 1 To TmpRs.RecordCount
       rn = MSHFlexGrid1.Rows - 1
       MSHFlexGrid1.TextMatrix(rn, 1) = "Fumigation"
       MSHFlexGrid1.TextMatrix(rn, 2) = TmpRs!FumigationID
       MSHFlexGrid1.TextMatrix(rn, 3) = TmpRs!CMPName
       MSHFlexGrid1.TextMatrix(rn, 4) = TmpRs!GodownNo
       MSHFlexGrid1.TextMatrix(rn, 5) = TmpRs!Address
       MSHFlexGrid1.TextMatrix(rn, 6) = TmpRs!Rate
       MSHFlexGrid1.TextMatrix(rn, 7) = TmpRs!BillableQty
       MSHFlexGrid1.TextMatrix(rn, 8) = TmpRs!Amount
       MSHFlexGrid1.TextMatrix(rn, 9) = TmpRs!Remark
       MSHFlexGrid1.TextMatrix(rn, 13) = TmpRs!FumigationDate
       MSHFlexGrid1.TextMatrix(rn, 17) = IIf(IsNull(TmpRs!PrintedDate), TmpRs!FumigationDate, TmpRs!PrintedDate)

       MSHFlexGrid1.row = rn
       MSHFlexGrid1.Col = 0
       MSHFlexGrid1.CellFontName = "Wingdings"
       MSHFlexGrid1.CellFontSize = 12
       MSHFlexGrid1.CellAlignment = flexAlignCenterCenter
       MSHFlexGrid1.Text = strChecked
       MSHFlexGrid1.TextMatrix(0, 7) = Val(MSHFlexGrid1.TextMatrix(0, 7)) + Val(MSHFlexGrid1.TextMatrix(rn, 7))
       MSHFlexGrid1.TextMatrix(0, 8) = Val(MSHFlexGrid1.TextMatrix(0, 8)) + Val(MSHFlexGrid1.TextMatrix(rn, 8))

       MSHFlexGrid1.TextMatrix(rn, 10) = IIf(IsNull(TmpRs!VendorRate), "", TmpRs!VendorRate)
       MSHFlexGrid1.TextMatrix(rn, 11) = IIf(IsNull(TmpRs!VendorQuantity), "", TmpRs!VendorQuantity)
       MSHFlexGrid1.TextMatrix(rn, 12) = Format(Val(MSHFlexGrid1.TextMatrix(rn, 10)) * Val(MSHFlexGrid1.TextMatrix(rn, 11)), "########0.00")
       MSHFlexGrid1.TextMatrix(0, 11) = Val(MSHFlexGrid1.TextMatrix(0, 11)) + Val(MSHFlexGrid1.TextMatrix(rn, 11))
       MSHFlexGrid1.TextMatrix(0, 12) = Format(Val(MSHFlexGrid1.TextMatrix(0, 12)) + Val(MSHFlexGrid1.TextMatrix(rn, 12)), "#####0.00")

       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
       TmpRs.MoveNext
   Next
End If
Exit Sub
NewDetail:

tmp = "Select TFD.FumigationID, TFD.FumgDetailID,MC.CMPName,MG.GodownNo,MG.Address,MFT.RateperMT,TFD.ActualQty,TFD.Amount,TFD.Remark, TFD.VendorRate,"
tmp = tmp & " TFD.VendorQuantity,TFD.MappingID,TF.FumigationDate, TFD.StackNo, MCo.Commodity,TF.PrintedDate,TFD.LastFumgDate,isnull(Mco.EffectiveDays,0) 'EffectiveDays'  from Txn_Fumigation TF"
tmp = tmp & " Inner join Txn_FumigationDetail TFD on TFD.FumigationID = TF.FumigationID"
tmp = tmp & " Left Join Mst_FumigationTreatmentDetail MFT on MFT.DetailID =TFD.AgreementDetailID "
tmp = tmp & " Left join Mst_TreatmentType MTT on MFT.TreatmentTypeID =MTT.TreatmentTypeID"
tmp = tmp & " Left Join Mst_Commodity MCo on MCo.CommodityID = TFD.CommodityID "
tmp = tmp & " Inner join Mst_Godown MG on MG.GodownID=TF.GodownID"
tmp = tmp & " Inner join Mst_CMP MC on MC.CMPID=MG.CMPID"
tmp = tmp & " Where TFD.MappingID=" & Val(TxtMappingID)
tmp = tmp & " Order by TFD.FumigationID, TFD.FumgDetailID"
Call TmpTable

If TmpRs.RecordCount > 0 Then
   For I = 1 To TmpRs.RecordCount
       rn = MSHFlexGrid1.Rows - 1
       MSHFlexGrid1.TextMatrix(rn, 1) = "Fumigation"
       MSHFlexGrid1.TextMatrix(rn, 2) = IIf(IsNull(TmpRs!FumgDetailID), 0, TmpRs!FumgDetailID)
       MSHFlexGrid1.TextMatrix(rn, 3) = IIf(IsNull(TmpRs!CMPName), "", TmpRs!CMPName)
       MSHFlexGrid1.TextMatrix(rn, 4) = IIf(IsNull(TmpRs!GodownNo), "", TmpRs!GodownNo)
       MSHFlexGrid1.TextMatrix(rn, 5) = IIf(IsNull(TmpRs!Address), "", TmpRs!Address)
       MSHFlexGrid1.TextMatrix(rn, 6) = IIf(IsNull(TmpRs!RateperMT), "", TmpRs!RateperMT)
       MSHFlexGrid1.TextMatrix(rn, 7) = IIf(IsNull(TmpRs!ActualQty), "", TmpRs!ActualQty)
       MSHFlexGrid1.TextMatrix(rn, 8) = IIf(IsNull(TmpRs!Amount), "", TmpRs!Amount)
       MSHFlexGrid1.TextMatrix(rn, 9) = IIf(IsNull(TmpRs!Remark), "", TmpRs!Remark)
       MSHFlexGrid1.TextMatrix(rn, 13) = IIf(IsNull(TmpRs!FumigationDate), "", TmpRs!FumigationDate)
       MSHFlexGrid1.TextMatrix(rn, 14) = IIf(IsNull(TmpRs!StackNo), "", TmpRs!StackNo)
       MSHFlexGrid1.TextMatrix(rn, 15) = IIf(IsNull(TmpRs!Commodity), "", TmpRs!Commodity)
       MSHFlexGrid1.TextMatrix(rn, 16) = IIf(IsNull(TmpRs!FumigationID), "", TmpRs!FumigationID)
       MSHFlexGrid1.TextMatrix(rn, 17) = IIf(IsNull(TmpRs!PrintedDate), TmpRs!FumigationDate, TmpRs!PrintedDate)
       MSHFlexGrid1.TextMatrix(rn, 18) = IIf(IsNull(TmpRs!LastFumgDate), TmpRs!FumigationDate, TmpRs!LastFumgDate)

       If DateDiff("d", TmpRs!LastFumgDate, TmpRs!FumigationDate) <= TmpRs!EffectiveDays Then
          For y = 0 To MSHFlexGrid1.Cols - 1
              MSHFlexGrid1.Col = y
              MSHFlexGrid1.row = rn
              MSHFlexGrid1.CellBackColor = vbGreen
          Next
       End If

       MSHFlexGrid1.row = rn
       MSHFlexGrid1.Col = 0
       MSHFlexGrid1.CellFontName = "Wingdings"
       MSHFlexGrid1.CellFontSize = 12
       MSHFlexGrid1.CellAlignment = flexAlignCenterCenter
       MSHFlexGrid1.Text = strChecked
       MSHFlexGrid1.TextMatrix(0, 7) = Val(MSHFlexGrid1.TextMatrix(0, 7)) + Val(MSHFlexGrid1.TextMatrix(rn, 7))
       MSHFlexGrid1.TextMatrix(0, 8) = Val(MSHFlexGrid1.TextMatrix(0, 8)) + Val(MSHFlexGrid1.TextMatrix(rn, 8))

       MSHFlexGrid1.TextMatrix(rn, 10) = IIf(IsNull(TmpRs!VendorRate), "", TmpRs!VendorRate)
       MSHFlexGrid1.TextMatrix(rn, 11) = IIf(IsNull(TmpRs!VendorQuantity), "", TmpRs!VendorQuantity)
       MSHFlexGrid1.TextMatrix(rn, 12) = Format(Val(MSHFlexGrid1.TextMatrix(rn, 10)) * Val(MSHFlexGrid1.TextMatrix(rn, 11)), "########0.00")
       MSHFlexGrid1.TextMatrix(0, 11) = Val(MSHFlexGrid1.TextMatrix(0, 11)) + Val(MSHFlexGrid1.TextMatrix(rn, 11))
       MSHFlexGrid1.TextMatrix(0, 12) = Format(Val(MSHFlexGrid1.TextMatrix(0, 12)) + Val(MSHFlexGrid1.TextMatrix(rn, 12)), "#####0.00")

       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
       TmpRs.MoveNext
   Next
End If
End Sub

Private Sub InDataSprayingDetail()

If Format(TxtMappingDate.Value, "yyyy-mm-dd") > "2014-01-17" Then GoTo NewDetail

tmp = "Select TS.SprayingID,MC.CMPName,MG.GodownNo,MG.Address,TS.Rate,TS.ActualQty,TS.Amount,TS.Remark,TS.VendorRate,TS.VendorQuantity,TS.MappingID,TS.SprayingDate,TS.PrintedDate from Txn_Spraying TS"
tmp = tmp & " Inner join Mst_Godown MG on MG.GodownID=TS.GodownID"
tmp = tmp & " Inner join Mst_CMP MC on MC.CMPID=MG.CMPID"
tmp = tmp & " Where TS.MappingID=" & Val(TxtMappingID)
tmp = tmp & " Order by TS.SprayingID"

Call TmpTable

If TmpRs.RecordCount > 0 Then
   For I = 1 To TmpRs.RecordCount
       rn = MSHFlexGrid1.Rows - 1
       MSHFlexGrid1.TextMatrix(rn, 1) = "Spraying"
       MSHFlexGrid1.TextMatrix(rn, 2) = TmpRs!SprayingID
       MSHFlexGrid1.TextMatrix(rn, 3) = TmpRs!CMPName
       MSHFlexGrid1.TextMatrix(rn, 4) = TmpRs!GodownNo
       MSHFlexGrid1.TextMatrix(rn, 5) = TmpRs!Address
       MSHFlexGrid1.TextMatrix(rn, 6) = TmpRs!Rate
       MSHFlexGrid1.TextMatrix(rn, 7) = TmpRs!ActualQty
       MSHFlexGrid1.TextMatrix(rn, 8) = TmpRs!Amount
       MSHFlexGrid1.TextMatrix(rn, 9) = TmpRs!Remark
       MSHFlexGrid1.TextMatrix(rn, 13) = TmpRs!SprayingDate
       MSHFlexGrid1.TextMatrix(rn, 17) = IIf(IsNull(TmpRs!PrintedDate), TmpRs!SprayingDate, TmpRs!PrintedDate)
       
       MSHFlexGrid1.row = rn
       MSHFlexGrid1.Col = 0
       MSHFlexGrid1.CellFontName = "Wingdings"
       MSHFlexGrid1.CellFontSize = 12
       MSHFlexGrid1.CellAlignment = flexAlignCenterCenter
       MSHFlexGrid1.Text = strChecked
       MSHFlexGrid1.TextMatrix(0, 7) = Val(MSHFlexGrid1.TextMatrix(0, 7)) + Val(MSHFlexGrid1.TextMatrix(rn, 7))
       MSHFlexGrid1.TextMatrix(0, 8) = Val(MSHFlexGrid1.TextMatrix(0, 8)) + Val(MSHFlexGrid1.TextMatrix(rn, 8))
          
       MSHFlexGrid1.TextMatrix(rn, 10) = IIf(IsNull(TmpRs!VendorRate), "", TmpRs!VendorRate)
       MSHFlexGrid1.TextMatrix(rn, 11) = IIf(IsNull(TmpRs!VendorQuantity), "", TmpRs!VendorQuantity)
       
       MSHFlexGrid1.TextMatrix(rn, 12) = Format(Val(MSHFlexGrid1.TextMatrix(rn, 10)) * Val(MSHFlexGrid1.TextMatrix(rn, 11)), "########0.00")
       MSHFlexGrid1.TextMatrix(0, 11) = Val(MSHFlexGrid1.TextMatrix(0, 11)) + Val(MSHFlexGrid1.TextMatrix(rn, 11))
       MSHFlexGrid1.TextMatrix(0, 12) = Format(Val(MSHFlexGrid1.TextMatrix(0, 12)) + Val(MSHFlexGrid1.TextMatrix(rn, 12)), "#####0.00")
                     
       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
       TmpRs.MoveNext
   Next
End If

Exit Sub

NewDetail:

tmp = "Select TS.SprayingID,MC.CMPName,MG.GodownNo,MG.Address,TS.Rate,TS.ActualArea,TS.Amount,TS.Remark,TS.VendorRate,TS.VendorQuantity,TS.MappingID,TS.SprayingDate,TS.PrintedDate from Txn_Spraying TS"
tmp = tmp & " Inner join Mst_Godown MG on MG.GodownID=TS.GodownID"
tmp = tmp & " Inner join Mst_CMP MC on MC.CMPID=MG.CMPID"
tmp = tmp & " Where TS.MappingID=" & Val(TxtMappingID)
tmp = tmp & " Order by TS.SprayingID"

Call TmpTable

If TmpRs.RecordCount > 0 Then
   For I = 1 To TmpRs.RecordCount
       rn = MSHFlexGrid1.Rows - 1
       MSHFlexGrid1.TextMatrix(rn, 1) = "Spraying"
       MSHFlexGrid1.TextMatrix(rn, 2) = TmpRs!SprayingID
       MSHFlexGrid1.TextMatrix(rn, 3) = TmpRs!CMPName
       MSHFlexGrid1.TextMatrix(rn, 4) = TmpRs!GodownNo
       MSHFlexGrid1.TextMatrix(rn, 5) = TmpRs!Address
       MSHFlexGrid1.TextMatrix(rn, 6) = TmpRs!Rate
       MSHFlexGrid1.TextMatrix(rn, 7) = TmpRs!ActualArea 'TmpRs!ActualQty
       MSHFlexGrid1.TextMatrix(rn, 8) = TmpRs!Amount
       MSHFlexGrid1.TextMatrix(rn, 9) = TmpRs!Remark
       MSHFlexGrid1.TextMatrix(rn, 13) = TmpRs!SprayingDate
       MSHFlexGrid1.TextMatrix(rn, 17) = IIf(IsNull(TmpRs!PrintedDate), TmpRs!SprayingDate, TmpRs!PrintedDate)
       
       MSHFlexGrid1.row = rn
       MSHFlexGrid1.Col = 0
       MSHFlexGrid1.CellFontName = "Wingdings"
       MSHFlexGrid1.CellFontSize = 12
       MSHFlexGrid1.CellAlignment = flexAlignCenterCenter
       MSHFlexGrid1.Text = strChecked
       MSHFlexGrid1.TextMatrix(0, 7) = Val(MSHFlexGrid1.TextMatrix(0, 7)) + Val(MSHFlexGrid1.TextMatrix(rn, 7))
       MSHFlexGrid1.TextMatrix(0, 8) = Val(MSHFlexGrid1.TextMatrix(0, 8)) + Val(MSHFlexGrid1.TextMatrix(rn, 8))
          
       MSHFlexGrid1.TextMatrix(rn, 10) = IIf(IsNull(TmpRs!VendorRate), "", TmpRs!VendorRate)
       MSHFlexGrid1.TextMatrix(rn, 11) = IIf(IsNull(TmpRs!VendorQuantity), "", TmpRs!VendorQuantity)
       
       MSHFlexGrid1.TextMatrix(rn, 12) = Format(Val(MSHFlexGrid1.TextMatrix(rn, 10)) * Val(MSHFlexGrid1.TextMatrix(rn, 11)), "########0.00")
       MSHFlexGrid1.TextMatrix(0, 11) = Val(MSHFlexGrid1.TextMatrix(0, 11)) + Val(MSHFlexGrid1.TextMatrix(rn, 11))
       MSHFlexGrid1.TextMatrix(0, 12) = Format(Val(MSHFlexGrid1.TextMatrix(0, 12)) + Val(MSHFlexGrid1.TextMatrix(rn, 12)), "#####0.00")
                     
       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
       TmpRs.MoveNext
   Next
End If


End Sub

Private Sub UpperHierarchy(mEmployeeID_ As Variant)
tmp = "Select EmployeeID,EmployeeName + '~' + EmployeeNo As EmployeeName,DirectReportingTo From Mst_Employee Where EmployeeID=" & mEmployeeID_ & " And Flag='Y' And DesignationID<>1"
Call TmpTable
If TmpRs.RecordCount > 0 Then
   If Not IsNull(TmpRs!DirectReportingTo) Then
      CmbForwardTo.AddItem TmpRs!EmployeeName
      Call UpperHierarchy(TmpRs!DirectReportingTo)
   End If
End If
End Sub



Private Sub CmbSForwardTo_GotFocus()
tmp1 = CmbSForwardTo.Text
Call TableMst_Employee("Where  Flag = 'Y' ")
CmbSForwardTo.Clear
''If RsMst_Employee.RecordCount = 0 Then
''   MsgBox "No Employee Found!!!"
''   CmbForwardTo.SetFocus
''   Exit Sub
''End If

'tmp = "Select DirectReportingTo from Mst_Employee Where EmployeeID=" & Gen_EmployeeID
'Call Tmp2Table
'If TmpRs2.RecordCount > 0 Then
'   Call UpperHierarchy(TmpRs2!DirectReportingTo)
'End If

tmp = "Select EmployeeID,EmployeeName ,EmployeeNo  from Mst_Employee " 'Where DesignationID=1 And Flag='Y'"
Call Tmp2Table

For I = 1 To TmpRs2.RecordCount
    CmbSForwardTo.AddItem TmpRs2!EmployeeName & "~" & TmpRs2!EmployeeNo
    TmpRs2.MoveNext
Next

CmbForwardTo.Text = tmp1
End Sub

Private Sub CmbSForwardTo_LostFocus()
Dim strSplitString() As String
Dim EmpNo_, EmpName_ As String

If CmbSForwardTo.Text = "" Then Exit Sub

strSplitString = Split(CmbForwardTo, "~")
If UBound(strSplitString) < 1 Then
   MsgBox "Invalid Selection "
   Exit Sub
End If
EmpName_ = Trim(strSplitString(0))
EmpNo_ = Trim(strSplitString(1))

TmpRs2.MoveFirst
TmpRs2.Find "EmployeeNo = '" & EmpNo_ & "'"

If TmpRs2.EOF Then
   MsgBox "Invalid Selection "
   CmbSForwardTo.SetFocus
   Exit Sub
End If
mSForwardToID = TmpRs2!EmployeeID
'TxtFWDDesignation = GetDesignaion(RsMst_Employee!DesignationID)
'mForwardToEmailID = RsMst_Employee!EmployeeEmailID
End Sub



Private Sub CmbSVendorID_GotFocus()

tmp1 = CmbSVendorID.Text

'''Call TableMst_Vendor '("Where Flag = 'Y'")
'tmp = " Select VendorID ,VendorName from Mst_Vendor Where VendorID in ("
'tmp = tmp & " Select distinct VendorID  From Txn_Fumigation TF"
'tmp = tmp & " Inner join Mst_Godown MG on MG.GodownID=TF.GodownID"
'tmp = tmp & " Inner join Mst_CMP MC on MC.CMPID=MG.CMPID"
'tmp = tmp & " Where (IsNull(TF.MappingID, 0) = 0 Or IsNull(TF.MappingID, 0)=" & Val(TxtMappingID) & ") And MC.LocationID =" & mLocationID & " )"
'tmp = tmp & " OR VendorID in ("
'tmp = tmp & " Select distinct VendorID From Txn_Spraying TF"
'tmp = tmp & " Inner join Mst_Godown MG on MG.GodownID=TF.GodownID"
'tmp = tmp & " Inner join Mst_CMP MC on MC.CMPID=MG.CMPID"
'tmp = tmp & " Where (IsNull(TF.MappingID, 0) = 0 Or IsNull(TF.MappingID, 0)=" & Val(TxtMappingID) & ") And MC.LocationID =" & mLocationID & " )"

tmp = "Select * from Mst_Vendor "
Call Tmp2Table

CmbVendorID.Clear
If TmpRs2.RecordCount = 0 Then
   MsgBox "No Vendor Found!!!"
   CmbSVendorID.SetFocus
   'RsMst_Vendor.MoveFirst
   Exit Sub
End If
For I = 1 To TmpRs2.RecordCount
    CmbSVendorID.AddItem TmpRs2!VendorName ' & "~" & RsMst_Vendor!VendorNo
    TmpRs2.MoveNext
Next

CmbSVendorID.Text = tmp1

End Sub

Private Sub CmbSVendorID_LostFocus()

If Trim(CmbSVendorID.Text) = "" Then
   mSVendorID = 0
   Exit Sub
End If

TmpRs2.MoveFirst
TmpRs2.Find "VendorName = '" & CmbSVendorID.Text & "'"

If TmpRs2.EOF Then
   MsgBox "Invalid Selection "
   CmbSVendorID.SetFocus
   Exit Sub
End If

End Sub
