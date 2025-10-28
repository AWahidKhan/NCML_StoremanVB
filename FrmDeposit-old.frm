VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmDeposit 
   Caption         =   "Deposit"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   645
      Left            =   0
      TabIndex        =   102
      Top             =   0
      Visible         =   0   'False
      Width           =   1995
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   350
         Left            =   80
         TabIndex        =   103
         Top             =   150
         Width           =   1000
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1455
         Left            =   70
         TabIndex        =   104
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
   Begin VB.PictureBox Picture1 
      Height          =   8655
      Left            =   60
      ScaleHeight     =   8595
      ScaleWidth      =   15075
      TabIndex        =   38
      Top             =   120
      Width           =   15135
      Begin VB.PictureBox Picture2 
         Height          =   1935
         Left            =   120
         ScaleHeight     =   1875
         ScaleWidth      =   14835
         TabIndex        =   54
         Top             =   120
         Width           =   14895
         Begin VB.TextBox TxtCVariation 
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
            Left            =   13440
            Locked          =   -1  'True
            TabIndex        =   97
            TabStop         =   0   'False
            Top             =   1320
            Width           =   1215
         End
         Begin VB.TextBox TxtCLotSize 
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
            Left            =   13440
            Locked          =   -1  'True
            TabIndex        =   95
            TabStop         =   0   'False
            Top             =   900
            Width           =   1215
         End
         Begin VB.CommandButton CmdContractID 
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
            Left            =   11040
            TabIndex        =   92
            TabStop         =   0   'False
            Top             =   960
            Width           =   375
         End
         Begin VB.CommandButton CmdCMPID 
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
            Left            =   11040
            TabIndex        =   91
            TabStop         =   0   'False
            Top             =   105
            Width           =   375
         End
         Begin VB.CommandButton CmdCommodityID 
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
            Left            =   11040
            TabIndex        =   90
            TabStop         =   0   'False
            Top             =   510
            Width           =   375
         End
         Begin VB.TextBox TxtCStdDed 
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
            Left            =   13440
            Locked          =   -1  'True
            TabIndex        =   86
            TabStop         =   0   'False
            Top             =   510
            Width           =   1215
         End
         Begin VB.TextBox TxtCGunnyWt 
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
            Left            =   13440
            Locked          =   -1  'True
            TabIndex        =   84
            TabStop         =   0   'False
            Top             =   105
            Width           =   1215
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Check1"
            Height          =   375
            Left            =   4800
            TabIndex        =   7
            Top             =   960
            Width           =   255
         End
         Begin VB.TextBox TxtCxTFNo 
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
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   1
            Top             =   150
            Width           =   1815
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
            Left            =   1200
            ScrollBars      =   2  'Vertical
            TabIndex        =   4
            Top             =   1365
            Width           =   9735
         End
         Begin MSComCtl2.DTPicker TxtCxTFDate 
            Height          =   360
            Left            =   1200
            TabIndex        =   2
            Top             =   555
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   635
            _Version        =   393216
            Format          =   20709377
            CurrentDate     =   36494
         End
         Begin VB.ComboBox CmbContractID 
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
            Sorted          =   -1  'True
            TabIndex        =   8
            Top             =   960
            Width           =   5895
         End
         Begin VB.TextBox TxtContractID 
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
            Left            =   5040
            TabIndex        =   55
            TabStop         =   0   'False
            Top             =   960
            Width           =   5895
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
            ItemData        =   "FrmDeposit.frx":0000
            Left            =   1200
            List            =   "FrmDeposit.frx":0002
            Sorted          =   -1  'True
            TabIndex        =   3
            Top             =   960
            Width           =   1815
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
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   99
            Top             =   960
            Width           =   1815
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
            Left            =   4800
            Sorted          =   -1  'True
            TabIndex        =   6
            Top             =   510
            Width           =   6135
         End
         Begin VB.ComboBox CmbCMPID 
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
            Left            =   4800
            Sorted          =   -1  'True
            TabIndex        =   5
            Top             =   105
            Width           =   6135
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
            Height          =   375
            Left            =   4800
            Locked          =   -1  'True
            TabIndex        =   57
            TabStop         =   0   'False
            Top             =   510
            Width           =   6135
         End
         Begin VB.TextBox TxtCMPID 
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
            Left            =   4800
            Locked          =   -1  'True
            TabIndex        =   56
            TabStop         =   0   'False
            Top             =   105
            Width           =   6135
         End
         Begin VB.Label Label43 
            Caption         =   "% Variation "
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
            Left            =   11520
            TabIndex        =   98
            Top             =   1320
            Width           =   1695
         End
         Begin VB.Label Label41 
            Caption         =   "Lot Size"
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
            Left            =   11520
            TabIndex        =   96
            Top             =   960
            Width           =   1695
         End
         Begin VB.Label Label32 
            Caption         =   "Std. Deduction"
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
            Left            =   11520
            TabIndex        =   87
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label Label31 
            Caption         =   "Gunny Wt. per Bag"
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
            Left            =   11520
            TabIndex        =   85
            Top             =   180
            Width           =   1695
         End
         Begin VB.Label Label2 
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
            Height          =   375
            Left            =   120
            TabIndex        =   64
            Top             =   150
            Width           =   1215
         End
         Begin VB.Label Label3 
            Caption         =   "CMP"
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
            Left            =   3120
            TabIndex        =   63
            Top             =   180
            Width           =   1215
         End
         Begin VB.Label Label4 
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
            Height          =   375
            Left            =   3120
            TabIndex        =   62
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label Label5 
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
            Height          =   375
            Left            =   120
            TabIndex        =   61
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label Label7 
            Caption         =   "Mode"
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
            TabIndex        =   60
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label28 
            Caption         =   "Booked Space ID"
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
            Left            =   3120
            TabIndex        =   59
            Top             =   960
            Width           =   1695
         End
         Begin VB.Label Label29 
            Caption         =   "Agent"
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
            TabIndex        =   58
            Top             =   1440
            Width           =   855
         End
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   6375
         Left            =   120
         TabIndex        =   39
         Top             =   2160
         Width           =   14895
         _ExtentX        =   26273
         _ExtentY        =   11245
         _Version        =   393216
         Tab             =   2
         TabHeight       =   520
         TabCaption(0)   =   "Client Detail"
         TabPicture(0)   =   "FrmDeposit.frx":0004
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Pic_Client"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Truck Details"
         TabPicture(1)   =   "FrmDeposit.frx":0020
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Pic_TTD"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Deposit Lot (G-S-L) Details"
         TabPicture(2)   =   "FrmDeposit.frx":003C
         Tab(2).ControlEnabled=   -1  'True
         Tab(2).Control(0)=   "Pic_GCTD"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).ControlCount=   1
         Begin VB.PictureBox Pic_TTD 
            Height          =   5535
            Left            =   -74880
            ScaleHeight     =   5475
            ScaleWidth      =   14355
            TabIndex        =   74
            Top             =   780
            Width           =   14415
            Begin VB.TextBox txtTTDRemarks 
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
               Height          =   495
               Left            =   10215
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   17
               Top             =   420
               Width           =   3975
            End
            Begin VB.CommandButton CmdTTDAdd2List 
               Caption         =   "Add to List"
               Height          =   375
               Left            =   12600
               TabIndex        =   18
               Top             =   960
               Width           =   1575
            End
            Begin VB.TextBox TxtTTDTareWt 
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
               Left            =   6600
               TabIndex        =   15
               Top             =   420
               Width           =   1815
            End
            Begin VB.TextBox TxtTTDTotalWt 
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
               Left            =   5040
               TabIndex        =   14
               Top             =   420
               Width           =   1575
            End
            Begin VB.TextBox TxtWeighBridgNo 
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
               Left            =   2520
               TabIndex        =   13
               Top             =   420
               Width           =   2535
            End
            Begin VB.TextBox TxtTTDTruckNo 
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
               Left            =   120
               TabIndex        =   12
               Top             =   420
               Width           =   2415
            End
            Begin VB.TextBox TxtTTDGrossWt 
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
               Left            =   8400
               TabIndex        =   16
               Top             =   420
               Width           =   1815
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
               Height          =   3930
               Left            =   120
               TabIndex        =   75
               Top             =   1455
               Width           =   14055
               _ExtentX        =   24791
               _ExtentY        =   6932
               _Version        =   393216
               WordWrap        =   -1  'True
               AllowUserResizing=   3
               _NumberOfBands  =   1
               _Band(0).Cols   =   2
            End
            Begin VB.Label Label6 
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
               Height          =   255
               Left            =   10920
               TabIndex        =   81
               Top             =   120
               Width           =   855
            End
            Begin VB.Label Label17 
               Caption         =   "Tare Weight"
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
               TabIndex        =   80
               Top             =   120
               Width           =   1335
            End
            Begin VB.Label Label16 
               Caption         =   "Total Weight"
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
               Left            =   5160
               TabIndex        =   79
               Top             =   120
               Width           =   1335
            End
            Begin VB.Label Label15 
               Caption         =   "Weighbridge Slip No."
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
               Left            =   2880
               TabIndex        =   78
               Top             =   120
               Width           =   1935
            End
            Begin VB.Label Label30 
               Caption         =   "Truck No."
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
               Left            =   600
               TabIndex        =   77
               Top             =   120
               Width           =   975
            End
            Begin VB.Label Label18 
               Caption         =   "Gross Weight"
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
               Left            =   8640
               TabIndex        =   76
               Top             =   120
               Width           =   1335
            End
         End
         Begin VB.PictureBox Pic_GCTD 
            Height          =   5535
            Left            =   120
            ScaleHeight     =   5475
            ScaleWidth      =   14595
            TabIndex        =   65
            Top             =   660
            Width           =   14655
            Begin VB.TextBox TxtAdjGrossWt 
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
               Left            =   8160
               Locked          =   -1  'True
               TabIndex        =   108
               TabStop         =   0   'False
               Top             =   420
               Width           =   1575
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
               Left            =   720
               TabIndex        =   106
               TabStop         =   0   'False
               Top             =   840
               Width           =   375
            End
            Begin VB.TextBox TxtLotStock 
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
               Left            =   2640
               Locked          =   -1  'True
               TabIndex        =   101
               TabStop         =   0   'False
               Top             =   840
               Width           =   1215
            End
            Begin VB.TextBox TxtAdjWt 
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
               Left            =   6600
               TabIndex        =   24
               Top             =   420
               Width           =   1575
            End
            Begin VB.TextBox TxtGCTDDematWt 
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
               Left            =   12840
               Locked          =   -1  'True
               TabIndex        =   93
               TabStop         =   0   'False
               Top             =   420
               Width           =   1680
            End
            Begin VB.TextBox TxtGCTDStdDeduction 
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
               Left            =   11280
               Locked          =   -1  'True
               TabIndex        =   88
               TabStop         =   0   'False
               Top             =   420
               Width           =   1575
            End
            Begin VB.TextBox TxtGCTDGunnyWt 
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
               Left            =   9720
               Locked          =   -1  'True
               TabIndex        =   82
               TabStop         =   0   'False
               Top             =   420
               Width           =   1575
            End
            Begin VB.ComboBox CmbGCTDClient 
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
               ItemData        =   "FrmDeposit.frx":0058
               Left            =   5190
               List            =   "FrmDeposit.frx":005A
               TabIndex        =   25
               Top             =   900
               Width           =   3975
            End
            Begin VB.TextBox TxtGCTDNoofBags 
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
               Left            =   3840
               TabIndex        =   22
               Top             =   420
               Width           =   1215
            End
            Begin VB.TextBox TxtGCTDGrossWt 
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
               Left            =   5040
               TabIndex        =   23
               Top             =   420
               Width           =   1575
            End
            Begin VB.CommandButton CmdGCTDAdd2List 
               Caption         =   "Add to List"
               Height          =   350
               Left            =   13200
               TabIndex        =   27
               Top             =   1395
               Width           =   1335
            End
            Begin VB.TextBox TxtGCTDremarks 
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
               Height          =   495
               Left            =   10080
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   26
               Top             =   840
               Width           =   4455
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid4 
               Height          =   3615
               Left            =   120
               TabIndex        =   66
               Top             =   1800
               Width           =   14415
               _ExtentX        =   25426
               _ExtentY        =   6376
               _Version        =   393216
               WordWrap        =   -1  'True
               AllowUserResizing=   3
               _NumberOfBands  =   1
               _Band(0).Cols   =   2
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
               ItemData        =   "FrmDeposit.frx":005C
               Left            =   120
               List            =   "FrmDeposit.frx":005E
               Sorted          =   -1  'True
               TabIndex        =   19
               Top             =   420
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
               ItemData        =   "FrmDeposit.frx":0060
               Left            =   1440
               List            =   "FrmDeposit.frx":0062
               Sorted          =   -1  'True
               TabIndex        =   20
               Top             =   420
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
               ItemData        =   "FrmDeposit.frx":0064
               Left            =   2640
               List            =   "FrmDeposit.frx":0066
               Sorted          =   -1  'True
               TabIndex        =   21
               Top             =   420
               Width           =   1215
            End
            Begin VB.Label Label19 
               Caption         =   "Adj Gross Wt."
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
               Left            =   8400
               TabIndex        =   109
               Top             =   120
               Width           =   1335
            End
            Begin VB.Label Label14 
               Caption         =   "Adjusted Wt."
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
               TabIndex        =   107
               Top             =   120
               Width           =   1335
            End
            Begin VB.Label Label1 
               Caption         =   "Lot Balance"
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
               Left            =   1440
               TabIndex        =   100
               Top             =   840
               Width           =   1095
            End
            Begin VB.Label Label40 
               Caption         =   "Electronic Weight"
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
               Left            =   12870
               TabIndex        =   94
               Top             =   105
               Width           =   1575
            End
            Begin VB.Label Label34 
               Caption         =   "Std. Deduction"
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
               Left            =   11400
               TabIndex        =   89
               Top             =   120
               Width           =   1335
            End
            Begin VB.Label Label27 
               Caption         =   "Gunny Weight"
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
               Left            =   9960
               TabIndex        =   83
               Top             =   120
               Width           =   1335
            End
            Begin VB.Label Label20 
               Caption         =   "Godown No"
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
               TabIndex        =   73
               Top             =   120
               Width           =   1095
            End
            Begin VB.Label Label21 
               Caption         =   "Stack No"
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
               TabIndex        =   72
               Top             =   120
               Width           =   975
            End
            Begin VB.Label Label22 
               Caption         =   "Lot No"
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
               TabIndex        =   71
               Top             =   120
               Width           =   735
            End
            Begin VB.Label Label23 
               Caption         =   "No Of Bags"
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
               Left            =   3960
               TabIndex        =   70
               Top             =   120
               Width           =   1095
            End
            Begin VB.Label Label24 
               Caption         =   "Gross Weight"
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
               Left            =   5280
               TabIndex        =   69
               Top             =   120
               Width           =   1215
            End
            Begin VB.Label LblClient_WHR_RRN 
               Caption         =   "Client Name ~ DP Name"
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
               Left            =   3960
               TabIndex        =   68
               Top             =   840
               Width           =   1095
            End
            Begin VB.Label Label25 
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
               Height          =   255
               Left            =   9240
               TabIndex        =   67
               Top             =   960
               Width           =   975
            End
         End
         Begin VB.PictureBox Pic_Client 
            Height          =   5535
            Left            =   -74880
            ScaleHeight     =   5475
            ScaleWidth      =   18315
            TabIndex        =   40
            Top             =   660
            Width           =   18375
            Begin VB.ComboBox CmbNonNcdexClient 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   1680
               Sorted          =   -1  'True
               TabIndex        =   110
               Top             =   120
               Visible         =   0   'False
               Width           =   495
            End
            Begin VB.CommandButton CmdClient 
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
               Height          =   240
               Left            =   1200
               TabIndex        =   105
               TabStop         =   0   'False
               Top             =   120
               Width           =   375
            End
            Begin VB.ComboBox CmbClientDP 
               Height          =   315
               Left            =   5160
               Sorted          =   -1  'True
               TabIndex        =   10
               Top             =   480
               Visible         =   0   'False
               Width           =   5055
            End
            Begin VB.CommandButton CmdClientAdd2List 
               Caption         =   "Add to List"
               Height          =   375
               Left            =   13200
               TabIndex        =   11
               Top             =   480
               Width           =   1335
            End
            Begin VB.TextBox TxtClientWeight 
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
               Left            =   11640
               Locked          =   -1  'True
               TabIndex        =   45
               TabStop         =   0   'False
               Top             =   480
               Width           =   1455
            End
            Begin VB.TextBox TxtClientNoBags 
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
               Left            =   10200
               Locked          =   -1  'True
               TabIndex        =   44
               TabStop         =   0   'False
               Top             =   480
               Width           =   1455
            End
            Begin VB.TextBox TxtDPName 
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
               Left            =   6360
               Locked          =   -1  'True
               TabIndex        =   43
               TabStop         =   0   'False
               Top             =   480
               Width           =   3855
            End
            Begin VB.TextBox TxtDPID 
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
               Left            =   5160
               Locked          =   -1  'True
               TabIndex        =   42
               TabStop         =   0   'False
               Top             =   480
               Width           =   1215
            End
            Begin VB.TextBox TxtClientName 
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
               Left            =   1680
               Locked          =   -1  'True
               TabIndex        =   41
               TabStop         =   0   'False
               Top             =   480
               Width           =   3495
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
               Left            =   120
               TabIndex        =   9
               Top             =   480
               Width           =   1575
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
               Height          =   4455
               Left            =   120
               TabIndex        =   46
               Top             =   960
               Width           =   14415
               _ExtentX        =   25426
               _ExtentY        =   7858
               _Version        =   393216
               WordWrap        =   -1  'True
               AllowUserResizing=   3
               _NumberOfBands  =   1
               _Band(0).Cols   =   2
            End
            Begin VB.Label lblNameDid 
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
               Left            =   6240
               TabIndex        =   53
               Top             =   180
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.Label Label13 
               Caption         =   "Weight"
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
               Left            =   12000
               TabIndex        =   52
               Top             =   120
               Width           =   975
            End
            Begin VB.Label Label12 
               Caption         =   "No Bags"
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
               Left            =   10440
               TabIndex        =   51
               Top             =   120
               Width           =   855
            End
            Begin VB.Label Label11 
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
               Left            =   7800
               TabIndex        =   50
               Top             =   180
               Width           =   1095
            End
            Begin VB.Label Label10 
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
               Left            =   5400
               TabIndex        =   49
               Top             =   180
               Width           =   735
            End
            Begin VB.Label Label8 
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
               Left            =   2400
               TabIndex        =   48
               Top             =   180
               Width           =   1095
            End
            Begin VB.Label Label9 
               Caption         =   "Client ID"
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
               TabIndex        =   47
               Top             =   180
               Width           =   975
            End
         End
      End
   End
   Begin VB.PictureBox Picture3 
      Height          =   930
      Left            =   75
      ScaleHeight     =   870
      ScaleWidth      =   14955
      TabIndex        =   0
      Top             =   8880
      Width           =   15015
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   375
         Left            =   11805
         TabIndex        =   37
         Top             =   450
         Width           =   3135
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   375
         Left            =   8790
         TabIndex        =   36
         Top             =   450
         Width           =   3015
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   375
         Left            =   5775
         TabIndex        =   35
         Top             =   450
         Width           =   3015
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   375
         Left            =   2880
         TabIndex        =   34
         Top             =   450
         Width           =   2895
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   375
         Left            =   120
         TabIndex        =   33
         Top             =   450
         Width           =   2775
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "&Search"
         Height          =   375
         Left            =   11805
         TabIndex        =   32
         Top             =   75
         Width           =   3135
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   8790
         TabIndex        =   31
         Top             =   75
         Width           =   3015
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   5775
         TabIndex        =   30
         Top             =   75
         Width           =   3015
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   375
         Left            =   2880
         TabIndex        =   29
         Top             =   75
         Width           =   2895
      End
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New "
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   75
         Width           =   2775
      End
   End
End
Attribute VB_Name = "FrmDeposit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg As Variant
Dim mClientIDRow, mCMPID, mCommodityID, mContractID, mGodownID, mGSLID, mGCTDClientID As Variant
Dim mClientIDRowGsl As Variant
Dim mExchangeTypeID As Double
Dim mCmbStack, mCmbLot As Variant

Private Sub Check1_Click()
CmbContractID.Enabled = False
Check1.Enabled = False
If mExchangeTypeID <> 1 Then Exit Sub
Check1.Enabled = True
If Check1.Value = 1 Then
   CmbContractID.Enabled = True
End If

End Sub

Private Sub CmbClientDP_LostFocus()

Dim Str_() As String
Str_ = Split(CmbClientDP.Text, "~")
If UBound(Str_) <> 1 Then
   MsgBox "Invalid Selection !!!!! "
   CmbClientDP.SetFocus
   Exit Sub
End If
TmpRs1.MoveFirst


TmpRs1.MoveFirst
TmpRs1.Find "DepositoryID  = '" & Str_(0) & "'"
If TmpRs1.EOF Then
   MsgBox "Invalid Selection !!!!! "
   CmbClientDP.SetFocus
   Exit Sub
End If

TxtClientName = TmpRs1!ClientName
TxtDPID = TmpRs1!DepositoryID
TxtDPName = TmpRs1!DepositoryName
CmbClientDP.Visible = False
mClientIDRow = TmpRs1!ClientIDRow
CmdClientAdd2List.SetFocus
End Sub



Private Sub CmbGCTDClient_GotFocus()
tmp = CmbGCTDClient.Text
CmbGCTDClient.Clear
If MSHFlexGrid2.Rows < 4 Then Exit Sub

For i = 2 To MSHFlexGrid2.Rows - 2
   CmbGCTDClient.AddItem MSHFlexGrid2.TextMatrix(i, 2) & " ~ " & MSHFlexGrid2.TextMatrix(i, 4)
Next i
CmbGCTDClient.Text = tmp
End Sub


Private Sub CmbGCTDClient_LostFocus()

For i = 2 To MSHFlexGrid2.Rows - 2
   If CmbGCTDClient.Text = MSHFlexGrid2.TextMatrix(i, 2) & " ~ " & MSHFlexGrid2.TextMatrix(i, 4) Then
      mGCTDClientID = Val(MSHFlexGrid2.TextMatrix(i, 7))
      Exit Sub
   End If
Next i


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
   MsgBox "Invalid  selection "
   CmbGodownNo.SetFocus
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
Call TableMst_GSL(" Where CMPID = " & mCMPID & " And GodownID = " & mGodownID & " And StackNo = '" & CmbStack.Text & "'")
CmbLot.Clear
If RsMst_GSL.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If

For i = 1 To RsMst_GSL.RecordCount
    CmbLot.AddItem RsMst_GSL!LotNo
    RsMst_GSL.MoveNext
Next i



'For i = 1 To RsMst_GSL.RecordCount
'    CmbLot.AddItem RsMst_GSL!LotNo
'    RsMst_GSL.MoveNext
'Next
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
   MsgBox "Invalid  selection "
   CmbLot.SetFocus
   Exit Sub
End If
TxtLotStock = IIf(IsNull(RsMst_GSL!DepositeWeight), 0, RsMst_GSL!DepositeWeight)

If Val(TxtLotStock) = 0 Then
   mGSLID = RsMst_GSL!GslID
   mClientIDRowGsl = Null
   Exit Sub
End If
If RsMst_GSL!CommodityID <> mCommodityID Then
   MsgBox "GSL Contain Stock of other Commodity "
   CmbLot.SetFocus
   Exit Sub
End If
If RsMst_GSL!ExchangeTypeID <> mExchangeTypeID Then
   MsgBox "GSL Contain Stock of other Exchange Type "
   CmbLot.SetFocus
   Exit Sub
End If
If RsMst_GSL!ISINID <> "" Or RsMst_GSL!ISINID <> "0" Then
   MsgBox "ISIN / CDC Mapped against Selected Lot "
   CmbLot.SetFocus
   Exit Sub
End If

mGSLID = RsMst_GSL!GslID
mClientIDRowGsl = RsMst_GSL!ClientIDRow

End Sub

Private Sub CmbNonNcdexClient_GotFocus()
tmp = CmbNonNcdexClient.Text
Call TableMst_Client("where ExchangeTypeId = " & mExchangeTypeID)
CmbNonNcdexClient.Clear
If RsMst_Client.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_Client.RecordCount
    CmbNonNcdexClient.AddItem RsMst_Client!ClientName
    RsMst_Client.MoveNext
Next
CmbNonNcdexClient.Text = tmp
End Sub
Private Sub CmbNonNcdexClient_LostFocus()
If CmbNonNcdexClient.Text = "" Then
   Exit Sub
End If
RsMst_Client.MoveFirst
RsMst_Client.Find "ClientName = '" & CmbNonNcdexClient.Text & "'"
If RsMst_Client.EOF Then
   MsgBox "Invalid  selection "
   CmbNonNcdexClient.SetFocus
   Exit Sub
End If
mClientIDRow = RsMst_Client!ClientIDRow
Dim ans As Variant
ans = GetClientNameByRow(RsMst_Client!ClientIDRow)
TxtClientID = ans(0)
TxtClientName = ans(1)
TxtDPID = ans(2)
TxtDPName = ans(3)

End Sub
Private Sub CmbStack_GotFocus()
If CmbGodownNo.Text = "" Then
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
   MsgBox "Invalid  selection "
   CmbStack.SetFocus
   Exit Sub
End If
If mCmbStack <> CmbStack.Text Then
   CmbLot.Text = ""
   
   mCmbLot = ""
End If
mCmbStack = CmbStack.Text
End Sub

Private Sub CmdClient_Click()
FrmMst_Client.Show
End Sub
Function CheckDuplicateClient() As Boolean
Dim RetVal As Boolean
RetVal = False
If MSHFlexGrid2.Rows > 3 Then
   For u = 1 To MSHFlexGrid2.Rows - 1
       If MSHFlexGrid2.TextMatrix(u, 1) = TxtClientID And MSHFlexGrid2.TextMatrix(u, 3) = TxtDPID Then
          If LCase(CmdClientAdd2List.Caption) = "add to list" Then
             RetVal = True
             Exit For
          Else
             If u <> MSHFlexGrid2.RowSel Then
                RetVal = True
             End If
          End If
       End If
   Next
End If

CheckDuplicateClient = RetVal
End Function

Private Sub CmdClientAdd2List_Click()
Dim r As Variant

If CmbClientDP.Visible = True Then
   MsgBox "Invalid Selection !!! "
   Exit Sub
End If
If CmbNonNcdexClient.Visible = True Then
    If CmbNonNcdexClient.Text = "" Then
       MsgBox "Invalid Selection !!! "
       CmbNonNcdexClient.SetFocus
       Exit Sub
    End If
End If
If TxtClientID = "" Then
   MsgBox "Invalid Client ID !!! "
   Exit Sub
End If

If CheckDuplicateClient = True Then
  MsgBox "Duplicate Client ID  not allowd !!!"
  TxtClientID.SetFocus
  Exit Sub
End If

If CmdClientAdd2List.Caption = "Update List" Then
   r = MSHFlexGrid2.RowSel
Else
  r = MSHFlexGrid2.Rows - 1
  MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
End If
MSHFlexGrid2.TextMatrix(r, 1) = TxtClientID
MSHFlexGrid2.TextMatrix(r, 2) = TxtClientName
MSHFlexGrid2.TextMatrix(r, 3) = TxtDPID
MSHFlexGrid2.TextMatrix(r, 4) = TxtDPName
MSHFlexGrid2.TextMatrix(r, 5) = TxtClientNoBags
MSHFlexGrid2.TextMatrix(r, 6) = TxtClientWeight
MSHFlexGrid2.TextMatrix(r, 7) = mClientIDRow

MSHFlexGrid2.TextMatrix(0, 5) = Val(MSHFlexGrid2.TextMatrix(0, 5)) + Val(MSHFlexGrid2.TextMatrix(r, 5))
MSHFlexGrid2.TextMatrix(0, 6) = Val(MSHFlexGrid2.TextMatrix(0, 6)) + Val(MSHFlexGrid2.TextMatrix(r, 6))
Call ClearClientTabControl
TxtClientID.Locked = False
If mExchangeTypeID <> 1 Then
   TxtClientID.SetFocus
Else
   CmbNonNcdexClient.SetFocus
End If
End Sub
Sub ClearClientTabControl()

TxtClientID = ""
TxtClientName = ""
TxtDPID = ""
TxtDPName = ""
TxtClientNoBags = ""
TxtClientWeight = ""
CmbClientDP.Text = ""
CmbClientDP.Visible = False
CmdClientAdd2List.Caption = "Add to List"
End Sub
Function CheckDuplicateGSL() As Boolean
Dim RetVal As Boolean
RetVal = False
If MSHFlexGrid4.Rows > 3 Then
   For u = 1 To MSHFlexGrid4.Rows - 1
       If MSHFlexGrid4.TextMatrix(u, 1) = CmbGodownNo.Text And _
          MSHFlexGrid4.TextMatrix(u, 2) = CmbStack.Text And _
          MSHFlexGrid4.TextMatrix(u, 3) = CmbLot.Text Then
          If LCase(CmdGCTDAdd2List.Caption) = "add to list" Then
             RetVal = True
             Exit For
          Else
             If u <> MSHFlexGrid4.RowSel Then
                RetVal = True
             End If
          End If
       End If
   Next
End If

CheckDuplicateGSL = RetVal
End Function

Private Sub CmdGCTDAdd2List_Click()
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

If (CmbGCTDClient.Text) = "" Then
   MsgBox "Blank Client ID Not Allowed !!! "
   CmbGCTDClient.SetFocus
   Exit Sub
End If



If mClientIDRowGsl <> 0 Then
   If mClientIDRowGsl <> (mGCTDClientID) Then
      MsgBox "GSL Contain Stock of other Client  "
      CmbLot.SetFocus
      Exit Sub
   End If
End If

MSHFlexGrid4.TextMatrix(1, 6) = "Adjusted Wt."
MSHFlexGrid4.TextMatrix(1, 7) = "Adj.Gross Wt."
MSHFlexGrid4.TextMatrix(1, 8) = "Gunny Weight"
MSHFlexGrid4.TextMatrix(1, 9) = "Std. Deduction"
MSHFlexGrid4.TextMatrix(1, 10) = "Demat Weight"
MSHFlexGrid4.TextMatrix(1, 11) = "Client ID"
MSHFlexGrid4.TextMatrix(1, 12) = "Remark"
MSHFlexGrid4.TextMatrix(1, 13) = "GSL ID"
MSHFlexGrid4.TextMatrix(1, 14) = "Godown Close"


If CmdGCTDAdd2List.Caption = "Update List" Then
   r = MSHFlexGrid4.RowSel
   MSHFlexGrid4.TextMatrix(0, 4) = Val(MSHFlexGrid4.TextMatrix(0, 4)) - Val(MSHFlexGrid4.TextMatrix(r, 4))
   MSHFlexGrid4.TextMatrix(0, 5) = Val(MSHFlexGrid4.TextMatrix(0, 5)) - Val(MSHFlexGrid4.TextMatrix(r, 5))
   MSHFlexGrid4.TextMatrix(0, 6) = Val(MSHFlexGrid4.TextMatrix(0, 6)) - Val(MSHFlexGrid4.TextMatrix(r, 6))
   MSHFlexGrid4.TextMatrix(0, 7) = Val(MSHFlexGrid4.TextMatrix(0, 7)) - Val(MSHFlexGrid4.TextMatrix(r, 7))
   MSHFlexGrid4.TextMatrix(0, 8) = Val(MSHFlexGrid4.TextMatrix(0, 8)) - Val(MSHFlexGrid4.TextMatrix(r, 8))
   MSHFlexGrid4.TextMatrix(0, 9) = Val(MSHFlexGrid4.TextMatrix(0, 9)) - Val(MSHFlexGrid4.TextMatrix(r, 9))
   MSHFlexGrid4.TextMatrix(0, 10) = Val(MSHFlexGrid4.TextMatrix(0, 10)) - Val(MSHFlexGrid4.TextMatrix(r, 10))
Else
  r = MSHFlexGrid4.Rows - 1
  MSHFlexGrid4.Rows = MSHFlexGrid4.Rows + 1
End If

MSHFlexGrid4.TextMatrix(r, 1) = CmbGodownNo.Text
MSHFlexGrid4.TextMatrix(r, 2) = CmbStack.Text
MSHFlexGrid4.TextMatrix(r, 3) = CmbLot.Text
MSHFlexGrid4.TextMatrix(r, 4) = Val(TxtGCTDNoofBags)

MSHFlexGrid4.TextMatrix(r, 5) = Val(TxtGCTDGrossWt)
MSHFlexGrid4.TextMatrix(r, 6) = Val(TxtAdjWt)
MSHFlexGrid4.TextMatrix(r, 7) = Val(TxtAdjGrossWt)

MSHFlexGrid4.TextMatrix(r, 8) = Val(TxtGCTDGunnyWt)
MSHFlexGrid4.TextMatrix(r, 9) = Val(TxtGCTDStdDeduction)
MSHFlexGrid4.TextMatrix(r, 10) = Val(TxtGCTDDematWt)
MSHFlexGrid4.TextMatrix(r, 11) = CmbGCTDClient.Text
MSHFlexGrid4.TextMatrix(r, 12) = TxtGCTDremarks
MSHFlexGrid4.TextMatrix(r, 13) = mGSLID
MSHFlexGrid4.TextMatrix(r, 15) = mGodownID
MSHFlexGrid4.TextMatrix(r, 16) = mGCTDClientID
 
MSHFlexGrid4.TextMatrix(0, 4) = Val(MSHFlexGrid4.TextMatrix(0, 4)) + Val(MSHFlexGrid4.TextMatrix(r, 4))
MSHFlexGrid4.TextMatrix(0, 5) = Val(MSHFlexGrid4.TextMatrix(0, 5)) + Val(MSHFlexGrid4.TextMatrix(r, 5))
MSHFlexGrid4.TextMatrix(0, 6) = Val(MSHFlexGrid4.TextMatrix(0, 6)) + Val(MSHFlexGrid4.TextMatrix(r, 6))
MSHFlexGrid4.TextMatrix(0, 7) = Val(MSHFlexGrid4.TextMatrix(0, 7)) + Val(MSHFlexGrid4.TextMatrix(r, 7))
MSHFlexGrid4.TextMatrix(0, 8) = Val(MSHFlexGrid4.TextMatrix(0, 8)) + Val(MSHFlexGrid4.TextMatrix(r, 8))
MSHFlexGrid4.TextMatrix(0, 9) = Val(MSHFlexGrid4.TextMatrix(0, 9)) + Val(MSHFlexGrid4.TextMatrix(r, 9))
MSHFlexGrid4.TextMatrix(0, 10) = Val(MSHFlexGrid4.TextMatrix(0, 10)) + Val(MSHFlexGrid4.TextMatrix(r, 10))

Call ClearGCTDTabControl
CmbGodownNo.SetFocus
End Sub
Private Sub ClearGCTDTabControl()
CmbGodownNo.Text = ""
CmbStack.Text = ""
CmbLot.Text = ""
CmbGCTDClient.Text = ""
TxtGCTDNoofBags = ""
TxtGCTDGrossWt = ""
TxtGCTDGunnyWt = ""
TxtGCTDStdDeduction = ""
TxtGCTDDematWt = ""
TxtGCTDremarks = ""
TxtRecClient = ""
TxtRecNCMSL = ""
TxtAdjGrossWt = ""
TxtAdjWt = ""
mGSLID = Null
CmdGCTDAdd2List.Caption = "Add to List"

End Sub

Private Sub CmdGSL_Click()

Gen_GSLID(0) = False
Gen_GSLID(1) = mCMPID
Gen_GSLID(2) = mGodownID
Gen_GSLID(3) = CmbStack.Text
Gen_GSLID(4) = CmbLot.Text
Gen_GSLID(5) = CmbCMPID.Text
Gen_GSLID(6) = CmbGodownNo.Text


FrmMst_GSL.Show
End Sub

Private Sub CmdTTDAdd2List_Click()
Dim r As Variant
If TxtTTDTruckNo = "" Then
   MsgBox "Blank Truck No !!! "
   TxtTTDTruckNo.SetFocus
   Exit Sub
End If


If (Val(TxtTTDTotalWt) - Val(TxtTTDTareWt)) <> Val(TxtTTDGrossWt) Then
   MsgBox "Invalid Weight !!! "
   Exit Sub
End If
If CmdTTDAdd2List.Caption = "Update List" Then
   r = MSHFlexGrid3.RowSel
   MSHFlexGrid3.TextMatrix(0, 3) = Val(MSHFlexGrid3.TextMatrix(0, 3)) - Val(MSHFlexGrid3.TextMatrix(r, 3))
   MSHFlexGrid3.TextMatrix(0, 4) = Val(MSHFlexGrid3.TextMatrix(0, 4)) - Val(MSHFlexGrid3.TextMatrix(r, 4))
   MSHFlexGrid3.TextMatrix(0, 5) = Val(MSHFlexGrid3.TextMatrix(0, 5)) - Val(MSHFlexGrid3.TextMatrix(r, 5))
Else
  r = MSHFlexGrid3.Rows - 1
  MSHFlexGrid3.Rows = MSHFlexGrid3.Rows + 1
End If
MSHFlexGrid3.TextMatrix(r, 1) = TxtTTDTruckNo
MSHFlexGrid3.TextMatrix(r, 2) = TxtWeighBridgNo
MSHFlexGrid3.TextMatrix(r, 3) = Val(TxtTTDTotalWt)
MSHFlexGrid3.TextMatrix(r, 4) = Val(TxtTTDTareWt)
MSHFlexGrid3.TextMatrix(r, 5) = Val(TxtTTDGrossWt)
MSHFlexGrid3.TextMatrix(r, 6) = txtTTDRemarks
MSHFlexGrid3.TextMatrix(0, 3) = Val(MSHFlexGrid3.TextMatrix(0, 3)) + Val(MSHFlexGrid3.TextMatrix(r, 3))
MSHFlexGrid3.TextMatrix(0, 4) = Val(MSHFlexGrid3.TextMatrix(0, 4)) + Val(MSHFlexGrid3.TextMatrix(r, 4))
MSHFlexGrid3.TextMatrix(0, 5) = Val(MSHFlexGrid3.TextMatrix(0, 5)) + Val(MSHFlexGrid3.TextMatrix(r, 5))
Call ClearTTDTabControl
TxtTTDTruckNo.SetFocus
End Sub
Sub ClearTTDTabControl()
TxtTTDTruckNo = ""
TxtWeighBridgNo = ""
TxtTTDTotalWt = ""
TxtTTDTareWt = ""
TxtTTDGrossWt = ""
txtTTDRemarks = ""
CmdTTDAdd2List.Caption = "Add to List"
End Sub
Private Sub Form_Load()
Call TableTxn_CxTF_Details
If RsTxn_CxTF_Details.RecordCount = 0 Then
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
Private Sub CmdExcel_Click()
Call Gen_Create_Xls("Txn_CxTF_Details.xls")
Call Xls_Detail_Rpt
End Sub
Public Sub Xls_Detail_Rpt()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Integer
Dim mCol As Integer
MsgBox ("Wait till Next Message")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\Txn_CxTF_Details.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For i = 0 To MSHFlexGrid1.Rows - 1
    mRow = i + 1 ' + 6
    For C = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(i, C), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(i, C)) Then
              oWS.Cells(mRow, C + 1) = Format(MSHFlexGrid1.TextMatrix(i, C), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(i, C)
           End If
        Else
           oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(i, C)
        End If
    Next
Next
oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub

Private Sub Label44_Click()
End Sub

Private Sub MSHFlexGrid2_Click()
    CmdClientAdd2List.Caption = "Update List"
    TxtClientID.Locked = False
    r = MSHFlexGrid2.RowSel
    If r <= 0 Then Exit Sub
    
    TxtClientID = MSHFlexGrid2.TextMatrix(r, 1)
    TxtClientName = MSHFlexGrid2.TextMatrix(r, 2)
    TxtDPID = MSHFlexGrid2.TextMatrix(r, 3)
    TxtDPName = MSHFlexGrid2.TextMatrix(r, 4)
    TxtClientNoBags = MSHFlexGrid2.TextMatrix(r, 5)
    TxtClientWeight = MSHFlexGrid2.TextMatrix(r, 6)
    
    If CmbNonNcdexClient.Visible = True Then
       CmbNonNcdexClient.Text = MSHFlexGrid2.TextMatrix(r, 2)
    End If

'    MSHFlexGrid2.TextMatrix(0, 5) = Val(MSHFlexGrid2.TextMatrix(0, 5)) - Val(MSHFlexGrid2.TextMatrix(r, 5))
'    MSHFlexGrid2.TextMatrix(0, 6) = Val(MSHFlexGrid2.TextMatrix(0, 6)) - Val(MSHFlexGrid2.TextMatrix(r, 6))

    
End Sub

Private Sub MSHFlexGrid3_Click()
    CmdTTDAdd2List.Caption = "Update List"
    r = MSHFlexGrid3.RowSel
    If r <= 1 Then Exit Sub
    TxtTTDTruckNo = MSHFlexGrid3.TextMatrix(r, 1)
    TxtWeighBridgNo = MSHFlexGrid3.TextMatrix(r, 2)
    TxtTTDTotalWt = MSHFlexGrid3.TextMatrix(r, 3)
    TxtTTDTareWt = MSHFlexGrid3.TextMatrix(r, 4)
    TxtTTDGrossWt = MSHFlexGrid3.TextMatrix(r, 5)
    txtTTDRemarks = MSHFlexGrid3.TextMatrix(r, 6)
End Sub

Private Sub MSHFlexGrid4_Click()
     
     
     r = MSHFlexGrid4.RowSel
     If r <= 1 Then Exit Sub
     If MSHFlexGrid4.TextMatrix(r, 14) <> "" Then
        MsgBox "This godown is Closed you can't Edit !!!!!"
        Exit Sub
     End If
     
     If MSHFlexGrid4.TextMatrix(r, 1) = "" Then
        Call ClearGCTDTabControl
        Exit Sub
     End If
     
     CmdGCTDAdd2List.Caption = "Update List"
     
     
     CmbGodownNo.Text = MSHFlexGrid4.TextMatrix(r, 1)
     CmbStack.Text = MSHFlexGrid4.TextMatrix(r, 2)
     CmbLot.Text = MSHFlexGrid4.TextMatrix(r, 3)
     TxtGCTDNoofBags = MSHFlexGrid4.TextMatrix(r, 4)
     TxtGCTDGrossWt = MSHFlexGrid4.TextMatrix(r, 5)
     
     TxtAdjWt = MSHFlexGrid4.TextMatrix(r, 6)
     TxtAdjGrossWt = MSHFlexGrid4.TextMatrix(r, 7)
     
     TxtGCTDGunnyWt = MSHFlexGrid4.TextMatrix(r, 8)
     TxtGCTDStdDeduction = MSHFlexGrid4.TextMatrix(r, 9)
     TxtGCTDDematWt = MSHFlexGrid4.TextMatrix(r, 10)
     
     CmbGCTDClient.Text = MSHFlexGrid4.TextMatrix(r, 11)
     TxtGCTDremarks = MSHFlexGrid4.TextMatrix(r, 12)
     mGCTDClientID = MSHFlexGrid4.TextMatrix(r, 16)
     
     mGSLID = MSHFlexGrid4.TextMatrix(r, 13)
     mGodownID = Val(MSHFlexGrid4.TextMatrix(r, 15))

     Call CmbGodownNo_GotFocus
     Call CmbGodownNo_LostFocus
     
     
End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsTxn_CxTF_Details.MoveNext
If RsTxn_CxTF_Details.EOF Then
   RsTxn_CxTF_Details.MoveLast
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
RsTxn_CxTF_Details.MovePrevious
If RsTxn_CxTF_Details.BOF Then
   RsTxn_CxTF_Details.MoveFirst
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
RsTxn_CxTF_Details.MoveFirst
Call InData
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsTxn_CxTF_Details.MoveLast
Call InData
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False
End Sub
Private Sub ExitCmd_Click()
Unload Me
End Sub
Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   If RsTxn_CxTF_Details.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   Call FirstCmd_Click
   Exit Sub
End If

Frame0.Height = Me.Height - 350
Frame0.Left = Picture1.Left - 100
Frame0.Top = Picture1.Top - 100
Frame0.Width = Me.Width - 150
Frame0.Visible = True
MSHFlexGrid1.Clear
MSHFlexGrid1.Cols = 3
MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 1000
RsTxn_CxTF_Details.MoveFirst
Call Grid_Head
MSHFlexGrid1.Rows = RsTxn_CxTF_Details.RecordCount + 1
MSHFlexGrid1.ColWidth(0) = 1000
For i = 1 To RsTxn_CxTF_Details.RecordCount
    For C = 0 To MSHFlexGrid1.Cols - 1
       MSHFlexGrid1.TextMatrix(i, C) = IIf(IsNull(RsTxn_CxTF_Details.Fields(C)), "", RsTxn_CxTF_Details.Fields(C))
    Next
    MSHFlexGrid1.TextMatrix(i, 2) = GetCMPName(MSHFlexGrid1.TextMatrix(i, 2))
    MSHFlexGrid1.TextMatrix(i, 3) = GetCommodityName(MSHFlexGrid1.TextMatrix(i, 3))
    MSHFlexGrid1.TextMatrix(i, 4) = GetExchangeName(MSHFlexGrid1.TextMatrix(i, 4))
    MSHFlexGrid1.TextMatrix(i, 9) = "Deposit"
    RsTxn_CxTF_Details.MoveNext
Next
End Sub
Public Sub InData()

Pic_Client.Enabled = False
Pic_GCTD.Enabled = False
Pic_TTD.Enabled = False

Call Grid_Head_TTD
Call Grid_Head_Client
Call Grid_Head_GCTD

TxtCxTFNo.Locked = True
TxtCxTFDate.Enabled = False
TxtCMPID.Locked = True
CmbCMPID.Visible = False
TxtCMPID.Visible = True
TxtCommodityID.Locked = True
CmbCommodityID.Visible = False
TxtCommodityID.Visible = True
TxtExchangeTypeID.Locked = True
CmbExchangeTypeID.Visible = False
TxtExchangeTypeID.Visible = True
'TxtBookedSpaceFlag.Locked = True
'TxtFlag.Locked = True
TxtAgent.Locked = True
'TxtRemark.Locked = True
'TxtTxnType.Locked = True
TxtContractID.Locked = True
CmbContractID.Visible = False
TxtContractID.Visible = True

TxtCxTFNo = IIf(IsNull(RsTxn_CxTF_Details!CxTFNo), "", RsTxn_CxTF_Details!CxTFNo)
TxtCxTFDate = IIf(IsNull(RsTxn_CxTF_Details!CxTFDate), "", RsTxn_CxTF_Details!CxTFDate)
TxtCMPID = GetCMPName(IIf(IsNull(RsTxn_CxTF_Details!CMPID), "", RsTxn_CxTF_Details!CMPID))
mCMPID = IIf(IsNull(RsTxn_CxTF_Details!CMPID), "", RsTxn_CxTF_Details!CMPID)
TxtCommodityID = GetCommodityName(IIf(IsNull(RsTxn_CxTF_Details!CommodityID), "", RsTxn_CxTF_Details!CommodityID))
mCommodityID = IIf(IsNull(RsTxn_CxTF_Details!CommodityID), "", RsTxn_CxTF_Details!CommodityID)
TxtExchangeTypeID = GetExchangeName(IIf(IsNull(RsTxn_CxTF_Details!ExchangeTypeID), "", RsTxn_CxTF_Details!ExchangeTypeID))
mExchangeTypeID = IIf(IsNull(RsTxn_CxTF_Details!ExchangeTypeID), "", RsTxn_CxTF_Details!ExchangeTypeID)
TxtBookedSpaceFlag = IIf(IsNull(RsTxn_CxTF_Details!BookedSpaceFlag), "", RsTxn_CxTF_Details!BookedSpaceFlag)
TxtFlag = IIf(IsNull(RsTxn_CxTF_Details!Flag), "", RsTxn_CxTF_Details!Flag)
TxtAgent = IIf(IsNull(RsTxn_CxTF_Details!Agent), "", RsTxn_CxTF_Details!Agent)
TxtRemark = IIf(IsNull(RsTxn_CxTF_Details!Remark), "", RsTxn_CxTF_Details!Remark)
TxtTxnType = IIf(IsNull(RsTxn_CxTF_Details!TxnType), "", RsTxn_CxTF_Details!TxnType)
'TxtContractID = GetContract(IIf(IsNull(RsTxn_CxTF_Details!ContractID), "", RsTxn_CxTF_Details!ContractID))
mContractID = IIf(IsNull(RsTxn_CxTF_Details!ContractID), "", RsTxn_CxTF_Details!ContractID)
CmbCMPID.Text = TxtCMPID
CmbExchangeTypeID.Text = TxtExchangeTypeID



Call CmbExchangeTypeID_GotFocus
Call CmbExchangeTypeID_LostFocus
Call CmbCMPID_GotFocus
Call CmbCMPID_LostFocus
CmbCommodityID.Text = TxtCommodityID
Call CmbCommodityID_GotFocus
Call CmbCommodityID_LostFocus

Call InDataTruck
Call InDataClient
Call InDataGSL
If RsTxn_CxTF_Details.RecordCount = 1 Then
   Call GButtonLock_1(Me, True)
Else
   Call GButtonLock(Me, True)
End If
End Sub



Private Sub SSTab1_Click(PreviousTab As Integer)
If RsTxn_CxTF_Details.RecordCount = 0 Then Exit Sub

If SaveCmd.Enabled = False Then
   Exit Sub
End If
If SSTab1.Tab = 0 Then
   TxtClientID.SetFocus
End If
If SSTab1.Tab = 1 Then
   TxtTTDTruckNo.SetFocus
End If
If SSTab1.Tab = 2 Then
   CmbGodownNo.SetFocus
End If

End Sub

Private Sub TxtAdjWt_LostFocus()
If TxtAdjWt = "" Then
   TxtAdjWt = Val(TxtAdjWt)
Else
    If IsNumeric(TxtAdjWt) = False Then
       MsgBox "Invalid Number format !!!"
       TxtAdjWt.SetFocus
       Exit Sub
    End If
End If
Call CalDmatWt
End Sub

Private Sub TxtClientID_KeyPress(KeyAscii As Integer)
tmp = KeyAscii
If KeyAscii = 14 Then
   SSTab1.Tab = 1
   TxtTTDTruckNo.SetFocus
End If
End Sub
Private Sub TxtTTDTruckNo_KeyPress(KeyAscii As Integer)
tmp = KeyAscii
If KeyAscii = 14 Then
   SSTab1.Tab = 2
   CmbGodownNo.SetFocus
End If
End Sub
Private Sub CmbGodownNo_KeyPress(KeyAscii As Integer)
tmp = KeyAscii
If KeyAscii = 14 Then
   SSTab1.Tab = 0
   If mExchangeTypeID <> 1 Then
      TxtClientID.SetFocus
   Else
       CmbNonNcdexClient.SetFocus
   End If
End If
End Sub

Private Sub TxtClientID_LostFocus()
CmbClientDP.Visible = False
If TxtClientID = "" Then Exit Sub
If CmbExchangeTypeID.Text = "" Then Exit Sub

Dim ans As Variant
ans = GetClientNameDtl(TxtClientID, mExchangeTypeID)

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
tmp = "SELECT Mst_Depository.DepositoryID, Mst_Depository.DepositoryName, Mst_Client.ClientID, " & _
      "Mst_Client.ClientName,  Mst_Client.ClientIDRow " & _
      "FROM Mst_Depository INNER JOIN Mst_Client ON Mst_Depository.DepositoryID = Mst_Client.DepositoryID " & _
      "WHERE (((Mst_Client.ExchangeTypeID)=" & mExchangeTypeID & ") AND ((Mst_Client.ClientID)='" & TxtClientID & "'));"
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

Private Sub TxtCxTFDate_LostFocus()
If TxtCxTFDate.Value > Date Then
   MsgBox "Future Date is not Allowed !!!! "
   TxtCxTFDate.SetFocus
End If
End Sub

Private Sub TxtCxTFNo_LostFocus()
If TxtCxTFNo = "" Then Exit Sub
If IsNumeric(TxtCxTFNo) = False Then
   MsgBox "Invalid number format !!!!"
   TxtCxTFNo.SetFocus
   Exit Sub
End If
TxtCxTFNo = Val(TxtCxTFNo)
End Sub
Private Sub AddCmd_Click()
Edit_Flg = "A"
SSTab1.Tab = 0
Call GButtonLock(Me, False)
Call Grid_Head_TTD
Call Grid_Head_Client
Call Grid_Head_GCTD

SSTab1.Tab = 0
Pic_Client.Enabled = True
Pic_GCTD.Enabled = True
Pic_TTD.Enabled = True

TxtCxTFNo.Locked = True
TxtCxTFDate.Enabled = True

TxtCMPID.Locked = False
CmbCMPID.Visible = True
TxtCMPID.Visible = False
TxtCommodityID.Locked = False
CmbCommodityID.Visible = True
TxtCommodityID.Visible = False
TxtExchangeTypeID.Locked = False
CmbExchangeTypeID.Visible = True
TxtExchangeTypeID.Visible = False
'TxtBookedSpaceFlag.Locked = False
'TxtFlag.Locked = False
TxtAgent.Locked = False
'TxtRemark.Locked = False
'TxtTxnType.Locked = False
TxtContractID.Locked = False
CmbContractID.Visible = True
TxtContractID.Visible = False
TxtClientID.Locked = False
TxtCxTFNo = ""
TxtCxTFDate = Date
TxtCMPID = ""
TxtCommodityID = ""
TxtExchangeTypeID = ""
TxtBookedSpaceFlag = ""
TxtFlag = ""
TxtAgent = ""
TxtRemark = ""
TxtTxnType = ""
TxtContractID = ""
CmbCMPID.Text = ""
CmbExchangeTypeID.Text = ""
CmbCommodityID.Text = ""

ClearClientTabControl
ClearTTDTabControl
ClearGCTDTabControl

TxtCGunnyWt = ""
TxtCStdDed = ""
TxtCLotSize = ""
TxtCVariation = ""

CmbContractID.Enabled = False

End Sub
Private Sub EditCmd_Click()
Edit_Flg = "E"
Call GButtonLock(Me, False)


Pic_Client.Enabled = True
Pic_GCTD.Enabled = True
Pic_TTD.Enabled = True


TxtCxTFNo.Locked = True
TxtCxTFDate.Enabled = True
TxtCMPID.Locked = True
'CmbCMPID.Visible = True
'TxtCMPID.Visible = False
TxtCommodityID.Locked = True
'CmbCommodityID.Visible = True
'TxtCommodityID.Visible = False
TxtExchangeTypeID.Locked = True
'CmbExchangeTypeID.Visible = True
'TxtExchangeTypeID.Visible = False
'TxtBookedSpaceFlag.Locked = False
'TxtFlag.Locked = False
TxtAgent.Locked = False
'TxtRemark.Locked = False
'TxtTxnType.Locked = False
TxtContractID.Locked = False
CmbContractID.Visible = True
TxtContractID.Visible = False
TxtClientID.Locked = False
CmbCMPID.Text = TxtCMPID
CmbCommodityID.Text = TxtCommodityID
CmbExchangeTypeID.Text = TxtExchangeTypeID
CmbContractID.Text = TxtContractID

End Sub
Private Sub DeleteCmd_Click()
Dim ans As Variant
On Error GoTo msgError
ans = MsgBox("Are you Sure you want to Delete this Record", vbYesNo)
If ans = vbYes Then
   Call TableMst_GSL("")
   For i = 2 To MSHFlexGrid4.Rows - 1
       If MSHFlexGrid4.TextMatrix(i, 1) <> "" Then
          RsMst_GSL.MoveFirst
          RsMst_GSL.Find "GSLID = " & Val(MSHFlexGrid4.TextMatrix(i, 13))
          If Not RsMst_GSL.EOF Then
             RsMst_GSL!CommodityID = Null
             RsMst_GSL!ExchangeTypeID = Null
             RsMst_GSL!DepositeBags = IIf(IsNull(RsMst_GSL!DepositeBags), 0, RsMst_GSL!DepositeBags) - Val(MSHFlexGrid4.TextMatrix(i, 4))
             RsMst_GSL!DepositeWeight = IIf(IsNull(RsMst_GSL!DepositeWeight), 0, RsMst_GSL!DepositeWeight) - Val(MSHFlexGrid4.TextMatrix(i, 5))
             RsMst_GSL.Update
          End If
      End If
   Next
   tmp = "Delete from Txn_CxTF_ClientDetail where CxTFNo = " & TxtCxTFNo
   Call TmpTable
   tmp = "Delete from Txn_CxTF_GSL where CxTFNo = " & TxtCxTFNo
   Call TmpTable
   tmp = "Delete from Txn_CxTF_TruckDetail where CxTFNo = " & TxtCxTFNo
   Call TmpTable
   
   tmp = "select * from Txn_CxTF_TruckDetail where CxTFNo = " & TxtCxTFNo
   Call TmpTable
   TmpRs.Close
   
   RsTxn_CxTF_Details.Delete
   RsTxn_CxTF_Details.Update
   If RsTxn_CxTF_Details.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsTxn_CxTF_Details.MoveNext
   If RsTxn_CxTF_Details.EOF() Then
      RsTxn_CxTF_Details.MoveLast
   End If
   Call InData
End If
Exit Sub
msgError:
MsgBox "Child Record Present "
RsTxn_CxTF_Details.CancelUpdate
End Sub
Private Sub SaveCmd_Click()


If CmbExchangeTypeID.Text = "" Then
   MsgBox "Select mode "
   CmbExchangeTypeID.SetFocus
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


If MSHFlexGrid2.Rows <= 3 Then
   MsgBox "Missing Client Detail !!!! "
   Exit Sub
End If
If MSHFlexGrid3.Rows <= 3 Then
   MsgBox "Missing Truck wise Detail !!!! "
   Exit Sub
End If
If MSHFlexGrid4.Rows <= 3 Then
   MsgBox "Missing Godown Stack Lot wise Detail !!!! "
   Exit Sub
End If

If Val(MSHFlexGrid3.TextMatrix(0, 5)) <> Val(MSHFlexGrid4.TextMatrix(0, 5)) Then
   MsgBox "Mis-match in TTD and GCTD Gross Weight  !!!! "
   Exit Sub
End If

If Edit_Flg = "A" Then
   RsTxn_CxTF_Details.AddNew
   RsTxn_CxTF_Details!CxTFNo = GetCxTFNo
   TxtCxTFNo.Text = RsTxn_CxTF_Details!CxTFNo
Else
    RsTxn_CxTF_Details.MoveFirst
    RsTxn_CxTF_Details.Find " CxTFNo= " & TxtCxTFNo.Text
End If

RsTxn_CxTF_Details!CxTFNo = TxtCxTFNo.Text
RsTxn_CxTF_Details!CxTFDate = TxtCxTFDate.Value
RsTxn_CxTF_Details!CMPID = mCMPID
RsTxn_CxTF_Details!CommodityID = mCommodityID
RsTxn_CxTF_Details!ExchangeTypeID = mExchangeTypeID
RsTxn_CxTF_Details!BookedSpaceFlag = CmbContractID.Text
'RsTxn_CxTF_Details!Flag = TxtFlag.Text
RsTxn_CxTF_Details!Agent = TxtAgent.Text
'RsTxn_CxTF_Details!Remark = TxtRemark.Text
RsTxn_CxTF_Details!TxnType = "D"
If mContractID <> "" Then
   RsTxn_CxTF_Details!ContractID = mContractID
End If
RsTxn_CxTF_Details.Update

SaveClient
SaveTruck
SaveGSL
SaveCmd.Enabled = False
Call InData
End Sub
Function getCxtf_ClinetMax() As Double
Dim RetVal As Double
RetVal = 0
tmp = "select max(ClientDetailID) from Txn_CxTF_ClientDetail "
Call TmpTable
If TmpRs.RecordCount > 0 Then
   RetVal = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
End If
RetVal = RetVal + 1
TmpRs.Close
getCxtf_ClinetMax = RetVal

End Function
Private Sub SaveClient()
Call TableTxn_CxTF_ClientDetail(" Where CxTFNo = " & TxtCxTFNo)

For i = 2 To MSHFlexGrid2.Rows - 1
    If MSHFlexGrid2.TextMatrix(i, 1) <> "" Then
       If MSHFlexGrid2.TextMatrix(i, 0) <> "" Then
          RsTxn_CxTF_ClientDetail.MoveFirst
          RsTxn_CxTF_ClientDetail.Find "ClientDetailID = " & MSHFlexGrid2.TextMatrix(i, 0)
          If RsTxn_CxTF_ClientDetail.EOF Then
             MsgBox "Not found client !!!!"
             Exit Sub
          End If
       Else
          RsTxn_CxTF_ClientDetail.AddNew
          RsTxn_CxTF_ClientDetail!ClientDetailID = getCxtf_ClinetMax
       End If

       RsTxn_CxTF_ClientDetail!CxTFNo = TxtCxTFNo
       RsTxn_CxTF_ClientDetail!ClientIDRow = MSHFlexGrid2.TextMatrix(i, 7)
       RsTxn_CxTF_ClientDetail.Update
       RsTxn_CxTF_ClientDetail.Requery
    End If
Next
End Sub

Private Sub InDataClient()
Dim ans As Variant

Call TableTxn_CxTF_ClientDetail(" Where CxTFNo = " & TxtCxTFNo)
MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + RsTxn_CxTF_ClientDetail.RecordCount
For i = 2 To MSHFlexGrid2.Rows - 2
    MSHFlexGrid2.TextMatrix(i, 0) = RsTxn_CxTF_ClientDetail!ClientDetailID
    ans = GetClientNameByRow(RsTxn_CxTF_ClientDetail!ClientIDRow)
    MSHFlexGrid2.TextMatrix(i, 1) = ans(0)
    MSHFlexGrid2.TextMatrix(i, 2) = ans(1) 'TmpRs!ClientName
    MSHFlexGrid2.TextMatrix(i, 3) = ans(2) 'TmpRs!DepositoryID
    MSHFlexGrid2.TextMatrix(i, 4) = ans(3) 'TmpRs!DepositoryName
    MSHFlexGrid2.TextMatrix(i, 7) = RsTxn_CxTF_ClientDetail!ClientIDRow
    RsTxn_CxTF_ClientDetail.MoveNext
Next
'MsgBox TmpRs.State

End Sub
Function getCxtf_TruckMax() As Double
Dim RetVal As Double
RetVal = 0
tmp = "select max(TruckDetailID) from Txn_CxTF_TruckDetail "
Call TmpTable
If TmpRs.RecordCount > 0 Then
   RetVal = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
End If
RetVal = RetVal + 1
TmpRs.Close
getCxtf_TruckMax = RetVal

End Function
Private Sub SaveTruck()
Call TableTxn_CxTF_TruckDetail(" Where CxTFNo = " & TxtCxTFNo)
For i = 2 To MSHFlexGrid3.Rows - 1
    If MSHFlexGrid3.TextMatrix(i, 1) <> "" Then
       If MSHFlexGrid3.TextMatrix(i, 0) <> "" Then
          RsTxn_CxTF_TruckDetail.MoveFirst
          RsTxn_CxTF_TruckDetail.Find "TruckDetailID = " & MSHFlexGrid3.TextMatrix(i, 0)
          If RsTxn_CxTF_TruckDetail.EOF Then
             MsgBox "Not found Truck !!!!"
             Exit Sub
          End If
       Else
          RsTxn_CxTF_TruckDetail.AddNew
          RsTxn_CxTF_TruckDetail!TruckDetailID = getCxtf_TruckMax
       End If
       RsTxn_CxTF_TruckDetail!CxTFNo = TxtCxTFNo
       RsTxn_CxTF_TruckDetail!TruckNo = MSHFlexGrid3.TextMatrix(i, 1)
       
       RsTxn_CxTF_TruckDetail!WeighBridgeSlipno = MSHFlexGrid3.TextMatrix(i, 2)
       RsTxn_CxTF_TruckDetail!TotalWeight = MSHFlexGrid3.TextMatrix(i, 3)
       RsTxn_CxTF_TruckDetail!TareWeight = MSHFlexGrid3.TextMatrix(i, 4)
       RsTxn_CxTF_TruckDetail!GrossWeight = MSHFlexGrid3.TextMatrix(i, 5)
       RsTxn_CxTF_TruckDetail!Remarks = MSHFlexGrid3.TextMatrix(i, 6)
       RsTxn_CxTF_TruckDetail.Update
       RsTxn_CxTF_TruckDetail.Requery
    End If
Next
End Sub
Function getCxtf_GSLMax() As Double
Dim RetVal As Double
RetVal = 0
tmp = "select max(GSLDetailID) from Txn_CxTF_GSL "
Call TmpTable
If TmpRs.RecordCount > 0 Then
   RetVal = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
End If
RetVal = RetVal + 1
TmpRs.Close
getCxtf_GSLMax = RetVal

End Function
Private Sub InDataGSL()
Call TableTxn_CxTF_GSL(" Where CxTFNo = " & TxtCxTFNo)
MSHFlexGrid4.Rows = MSHFlexGrid4.Rows + RsTxn_CxTF_GSL.RecordCount
If RsTxn_CxTF_GSL.RecordCount = 0 Then Exit Sub

Dim ans As Variant

For i = 2 To MSHFlexGrid4.Rows - 2
    
    ans = GetGSLByGLSID(RsTxn_CxTF_GSL!GslID)
    MSHFlexGrid4.TextMatrix(i, 0) = RsTxn_CxTF_GSL!GSLDetailID
    
    MSHFlexGrid4.TextMatrix(i, 1) = ans(3) ' TmpRs!GodownNo
    MSHFlexGrid4.TextMatrix(i, 2) = ans(4) 'TmpRs!StackNo
    MSHFlexGrid4.TextMatrix(i, 3) = ans(5) 'TmpRs!LotNo
    MSHFlexGrid4.TextMatrix(i, 14) = IIf(IsNull(ans(6)), "", ans(6))
    MSHFlexGrid4.TextMatrix(i, 15) = IIf(IsNull(ans(2)), "", ans(2))
    
    MSHFlexGrid4.TextMatrix(i, 4) = RsTxn_CxTF_GSL!NoOfBags
    MSHFlexGrid4.TextMatrix(i, 5) = RsTxn_CxTF_GSL!GrossWeight
    MSHFlexGrid4.TextMatrix(i, 6) = RsTxn_CxTF_GSL!AdjNClientWeight
    MSHFlexGrid4.TextMatrix(i, 7) = RsTxn_CxTF_GSL!AdjNCMSLWeight
    MSHFlexGrid4.TextMatrix(i, 8) = RsTxn_CxTF_GSL!GunnyWeight
    MSHFlexGrid4.TextMatrix(i, 9) = RsTxn_CxTF_GSL!StdDeduction
    MSHFlexGrid4.TextMatrix(i, 10) = RsTxn_CxTF_GSL!DematWeight
    
    ans = GetClientNameByRow(RsTxn_CxTF_GSL!ClientIDRow)
    MSHFlexGrid4.TextMatrix(i, 11) = ans(1) & " ~ " & ans(3)
    MSHFlexGrid4.TextMatrix(i, 12) = RsTxn_CxTF_GSL!Remarks
    
    MSHFlexGrid4.TextMatrix(i, 13) = RsTxn_CxTF_GSL!GslID
   
    'MSHFlexGrid4.TextMatrix(i, 14) = 'look at top Used Up after lot no
    'MSHFlexGrid4.TextMatrix(i, 15) = 'look at top Used Up after lot no
    
    MSHFlexGrid4.TextMatrix(i, 16) = RsTxn_CxTF_GSL!ClientIDRow
   
    MSHFlexGrid4.TextMatrix(0, 4) = Val(MSHFlexGrid4.TextMatrix(0, 4)) + Val(MSHFlexGrid4.TextMatrix(i, 4))
    MSHFlexGrid4.TextMatrix(0, 5) = Val(MSHFlexGrid4.TextMatrix(0, 5)) + Val(MSHFlexGrid4.TextMatrix(i, 5))
    MSHFlexGrid4.TextMatrix(0, 6) = Val(MSHFlexGrid4.TextMatrix(0, 6)) + Val(MSHFlexGrid4.TextMatrix(i, 6))
    MSHFlexGrid4.TextMatrix(0, 7) = Val(MSHFlexGrid4.TextMatrix(0, 7)) + Val(MSHFlexGrid4.TextMatrix(i, 7))
    MSHFlexGrid4.TextMatrix(0, 8) = Val(MSHFlexGrid4.TextMatrix(0, 8)) + Val(MSHFlexGrid4.TextMatrix(i, 8))
    MSHFlexGrid4.TextMatrix(0, 9) = Val(MSHFlexGrid4.TextMatrix(0, 9)) + Val(MSHFlexGrid4.TextMatrix(i, 9))
    MSHFlexGrid4.TextMatrix(0, 10) = Val(MSHFlexGrid4.TextMatrix(0, 10)) + Val(MSHFlexGrid4.TextMatrix(i, 10))
    RsTxn_CxTF_GSL.MoveNext
Next
End Sub

Private Sub SaveGSL()
Call TableTxn_CxTF_GSL(" Where CxTFNo = " & TxtCxTFNo)
Call TableMst_GSL("")
For i = 2 To MSHFlexGrid4.Rows - 1
    If MSHFlexGrid4.TextMatrix(i, 1) <> "" Then
       If MSHFlexGrid4.TextMatrix(i, 0) <> "" Then
          RsTxn_CxTF_GSL.MoveFirst
          RsTxn_CxTF_GSL.Find "GSLDetailID = " & MSHFlexGrid4.TextMatrix(i, 0)
          If RsTxn_CxTF_GSL.EOF Then
             MsgBox "Not found Truck !!!!"
             Exit Sub
          End If
       Else
          RsTxn_CxTF_GSL.AddNew
          RsTxn_CxTF_GSL!GSLDetailID = getCxtf_GSLMax
       End If
       RsMst_GSL.MoveFirst
       RsMst_GSL.Find "GSLID = " & Val(MSHFlexGrid4.TextMatrix(i, 13))
       If Not RsMst_GSL.EOF Then
          RsMst_GSL!CommodityID = mCommodityID
          RsMst_GSL!ExchangeTypeID = mExchangeTypeID
          RsMst_GSL!DepositeBags = IIf(IsNull(RsMst_GSL!DepositeBags), 0, RsMst_GSL!DepositeBags) - IIf(IsNull(RsTxn_CxTF_GSL!NoOfBags), 0, RsTxn_CxTF_GSL!NoOfBags)
          RsMst_GSL!DepositeWeight = IIf(IsNull(RsMst_GSL!DepositeWeight), 0, RsMst_GSL!DepositeWeight) - IIf(IsNull(RsTxn_CxTF_GSL!GrossWeight), 0, RsTxn_CxTF_GSL!GrossWeight)
          RsMst_GSL.Update
       End If
        
       RsTxn_CxTF_GSL!CxTFNo = TxtCxTFNo
       RsTxn_CxTF_GSL!NoOfBags = MSHFlexGrid4.TextMatrix(i, 4)
       RsTxn_CxTF_GSL!GrossWeight = MSHFlexGrid4.TextMatrix(i, 5)
       RsTxn_CxTF_GSL!AdjNClientWeight = Val(MSHFlexGrid4.TextMatrix(i, 6))
       RsTxn_CxTF_GSL!AdjNCMSLWeight = Val(MSHFlexGrid4.TextMatrix(i, 7))
       RsTxn_CxTF_GSL!GunnyWeight = Val(MSHFlexGrid4.TextMatrix(i, 8))
       RsTxn_CxTF_GSL!StdDeduction = Val(MSHFlexGrid4.TextMatrix(i, 9))
       RsTxn_CxTF_GSL!DematWeight = Val(MSHFlexGrid4.TextMatrix(i, 10))
       RsTxn_CxTF_GSL!ClientIDRow = Val(MSHFlexGrid4.TextMatrix(i, 16))
       RsTxn_CxTF_GSL!Remarks = MSHFlexGrid4.TextMatrix(i, 12)
       RsTxn_CxTF_GSL!GslID = MSHFlexGrid4.TextMatrix(i, 13)
       RsTxn_CxTF_GSL.Update
       RsTxn_CxTF_GSL.Requery
       If Not RsMst_GSL.EOF Then
          RsMst_GSL!CommodityID = mCommodityID
          RsMst_GSL!ExchangeTypeID = mExchangeTypeID
          RsMst_GSL!DepositeBags = IIf(IsNull(RsMst_GSL!DepositeBags), 0, RsMst_GSL!DepositeBags) + Val(MSHFlexGrid4.TextMatrix(i, 4))
          RsMst_GSL!DepositeWeight = IIf(IsNull(RsMst_GSL!DepositeWeight), 0, RsMst_GSL!DepositeWeight) + Val(MSHFlexGrid4.TextMatrix(i, 5))
          RsMst_GSL!BalanceBags = IIf(IsNull(RsMst_GSL!BalanceBags), 0, RsMst_GSL!BalanceBags) + Val(MSHFlexGrid4.TextMatrix(i, 4))
          RsMst_GSL!BalanceWeight = IIf(IsNull(RsMst_GSL!BalanceWeight), 0, RsMst_GSL!BalanceWeight) + Val(MSHFlexGrid4.TextMatrix(i, 5))
          RsMst_GSL!Flag = "O"
          RsMst_GSL!ClientIDRow = Val(MSHFlexGrid4.TextMatrix(i, 16))
          RsMst_GSL.Update
       End If
       
    End If
Next
End Sub



Private Sub InDataTruck()
Call TableTxn_CxTF_TruckDetail(" Where CxTFNo = " & TxtCxTFNo)
MSHFlexGrid3.Rows = MSHFlexGrid3.Rows + RsTxn_CxTF_TruckDetail.RecordCount
For i = 2 To MSHFlexGrid3.Rows - 2
   MSHFlexGrid3.TextMatrix(i, 0) = RsTxn_CxTF_TruckDetail!TruckDetailID
   MSHFlexGrid3.TextMatrix(i, 1) = RsTxn_CxTF_TruckDetail!TruckNo
   MSHFlexGrid3.TextMatrix(i, 2) = RsTxn_CxTF_TruckDetail!WeighBridgeSlipno
   MSHFlexGrid3.TextMatrix(i, 3) = RsTxn_CxTF_TruckDetail!TotalWeight
   MSHFlexGrid3.TextMatrix(i, 4) = RsTxn_CxTF_TruckDetail!TareWeight
   MSHFlexGrid3.TextMatrix(i, 5) = RsTxn_CxTF_TruckDetail!GrossWeight
   MSHFlexGrid3.TextMatrix(i, 6) = RsTxn_CxTF_TruckDetail!Remarks
   
   MSHFlexGrid3.TextMatrix(0, 3) = Val(MSHFlexGrid3.TextMatrix(0, 3)) + RsTxn_CxTF_TruckDetail!TotalWeight
   MSHFlexGrid3.TextMatrix(0, 4) = Val(MSHFlexGrid3.TextMatrix(0, 4)) + RsTxn_CxTF_TruckDetail!TareWeight
   MSHFlexGrid3.TextMatrix(0, 5) = Val(MSHFlexGrid3.TextMatrix(0, 5)) + RsTxn_CxTF_TruckDetail!GrossWeight
   
   RsTxn_CxTF_TruckDetail.MoveNext
Next
End Sub

Private Function GetCxTFNo() As Double
Dim RetVal As Double
tmp = "Select max(CxTFNo) from Txn_CxTF_Details "
Call TmpTable
RetVal = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
RetVal = RetVal + 1
GetCxTFNo = RetVal
End Function
Private Sub Grid_Head()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 11
MSHFlexGrid1.TextMatrix(0, 0) = "CDTF No"
MSHFlexGrid1.TextMatrix(0, 1) = "CDTF Date"
MSHFlexGrid1.TextMatrix(0, 2) = "CMP"
MSHFlexGrid1.TextMatrix(0, 3) = "Commodity"
MSHFlexGrid1.TextMatrix(0, 4) = "Exchange Type"
MSHFlexGrid1.TextMatrix(0, 5) = "a"
MSHFlexGrid1.TextMatrix(0, 6) = "Flag"
MSHFlexGrid1.TextMatrix(0, 7) = "Agent"
MSHFlexGrid1.TextMatrix(0, 8) = "Remark"
MSHFlexGrid1.TextMatrix(0, 9) = "Txn Type"
MSHFlexGrid1.TextMatrix(0, 10) = "Booked Space ID"
End Sub
Private Sub Grid_Head_Client()
MSHFlexGrid2.Clear
MSHFlexGrid2.Rows = 3
MSHFlexGrid2.FixedRows = 2

MSHFlexGrid2.Cols = 8
MSHFlexGrid2.TextMatrix(1, 0) = "Row ID"
MSHFlexGrid2.TextMatrix(1, 1) = "Client ID"
MSHFlexGrid2.TextMatrix(1, 2) = "Client Name"
MSHFlexGrid2.TextMatrix(1, 3) = "DP ID"
MSHFlexGrid2.TextMatrix(1, 4) = "DP Name "
MSHFlexGrid2.TextMatrix(1, 5) = "No of Bags"
MSHFlexGrid2.TextMatrix(1, 6) = "Weight"
MSHFlexGrid2.TextMatrix(1, 7) = "ClientRowID"


MSHFlexGrid2.ColWidth(0) = 500
MSHFlexGrid2.ColWidth(1) = 1800
MSHFlexGrid2.ColWidth(2) = 3500
MSHFlexGrid2.ColWidth(3) = 1800
MSHFlexGrid2.ColWidth(4) = 3500
MSHFlexGrid2.ColWidth(5) = 1500
MSHFlexGrid2.ColWidth(6) = 1500
MSHFlexGrid2.ColWidth(7) = 0

End Sub
Private Sub Grid_Head_TTD()
MSHFlexGrid3.Clear
MSHFlexGrid3.Rows = 3
MSHFlexGrid3.FixedRows = 2

MSHFlexGrid3.Cols = 7
MSHFlexGrid3.TextMatrix(1, 0) = "Row ID"
MSHFlexGrid3.TextMatrix(1, 1) = "Truck No"
MSHFlexGrid3.TextMatrix(1, 2) = "Weighbridge Slip No"
MSHFlexGrid3.TextMatrix(1, 3) = "Total Weight"
MSHFlexGrid3.TextMatrix(1, 4) = "Tare Weight"
MSHFlexGrid3.TextMatrix(1, 5) = "Gross Weight"
MSHFlexGrid3.TextMatrix(1, 6) = "Remark"

MSHFlexGrid3.ColWidth(0) = 0
MSHFlexGrid3.ColWidth(1) = 2000
MSHFlexGrid3.ColWidth(2) = 2500
MSHFlexGrid3.ColWidth(3) = 1500
MSHFlexGrid3.ColWidth(4) = 1500
MSHFlexGrid3.ColWidth(5) = 1500
MSHFlexGrid3.ColWidth(6) = 4500


End Sub
Private Sub Grid_Head_GCTD()
MSHFlexGrid4.Clear
MSHFlexGrid4.Rows = 3
MSHFlexGrid4.FixedRows = 2

MSHFlexGrid4.Cols = 17
MSHFlexGrid4.TextMatrix(1, 0) = "Row ID"
MSHFlexGrid4.TextMatrix(1, 1) = "Godown No"
MSHFlexGrid4.TextMatrix(1, 2) = "Stack No"
MSHFlexGrid4.TextMatrix(1, 3) = "Lot No"
MSHFlexGrid4.TextMatrix(1, 4) = "No of Bags"
MSHFlexGrid4.TextMatrix(1, 5) = "Gross Weight"

MSHFlexGrid4.TextMatrix(1, 6) = "Adjusted Wt."
MSHFlexGrid4.TextMatrix(1, 7) = "Adj.Gross Wt."
MSHFlexGrid4.TextMatrix(1, 8) = "Gunny Weight"
MSHFlexGrid4.TextMatrix(1, 9) = "Std. Deduction"
MSHFlexGrid4.TextMatrix(1, 10) = "Demat Weight"
MSHFlexGrid4.TextMatrix(1, 11) = "Client ID"
MSHFlexGrid4.TextMatrix(1, 12) = "Remark"
MSHFlexGrid4.TextMatrix(1, 13) = "GSL ID"
MSHFlexGrid4.TextMatrix(1, 14) = "Godown Close"
MSHFlexGrid4.TextMatrix(1, 15) = "Godown ID"
MSHFlexGrid4.TextMatrix(1, 16) = "ClientRowID"

MSHFlexGrid4.ColWidth(0) = 0
MSHFlexGrid4.ColWidth(1) = 1200
MSHFlexGrid4.ColWidth(2) = 1200
MSHFlexGrid4.ColWidth(3) = 1200
MSHFlexGrid4.ColWidth(4) = 1000
MSHFlexGrid4.ColWidth(11) = 1500
MSHFlexGrid4.ColWidth(12) = 2000
MSHFlexGrid4.ColWidth(13) = 0
MSHFlexGrid4.ColWidth(14) = 0
MSHFlexGrid4.ColWidth(15) = 0
MSHFlexGrid4.ColWidth(16) = 0
MSHFlexGrid4.RowHeight(1) = 750

End Sub


Private Sub MSHFlexGrid1_Click()
Dim i As Variant
i = MSHFlexGrid1.RowSel
If i <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(i, 0) = "" Then Exit Sub
RsTxn_CxTF_Details.MoveFirst
RsTxn_CxTF_Details.Find "CxTFNo = " & MSHFlexGrid1.TextMatrix(i, 0)
If Not RsTxn_CxTF_Details.EOF Then
   Call InData
   Frame0.Visible = False
End If
End Sub
Private Sub CmdCMPID_Click()
    FrmMst_CMP.Show
End Sub
'----
Private Sub CmbCMPID_GotFocus()
    tmp = CmbCMPID.Text
    Call TableMst_CMP
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
   MsgBox "Invalid  selection "
   CmbCMPID.SetFocus
   Exit Sub
End If
If mCMPID <> RsMst_CMP!CMPID Then
   Call Grid_Head_GCTD
End If

mCMPID = RsMst_CMP!CMPID
End Sub
'----
Private Sub CmdCommodityID_Click()
    FrmMst_Commodity.Show
End Sub
'----
Private Sub CmbCommodityID_GotFocus()
 If CmbExchangeTypeID.Text <> "" Then
'    If CmbCMPID.Text <> "" Then
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
    
'    Else
'       MsgBox "Please Select CMP!!!!!"
'       CmbCMPID.SetFocus
'       Exit Sub
'    End If
 Else
    MsgBox "Please Select Exchange Mode !!!!"
    CmbExchangeTypeID.SetFocus
    Exit Sub
 End If
End Sub
'----
Private Sub CmbCommodityID_LostFocus()
If CmbCommodityID.Text = "" Then
   Exit Sub
End If

RsMst_Commodity.MoveFirst
RsMst_Commodity.Find "Commodity = '" & CmbCommodityID.Text & "'"
If RsMst_Commodity.EOF Then
   MsgBox "Invalid  selection "
   CmbCommodityID.SetFocus
   Exit Sub
End If

If mCommodityID <> RsMst_Commodity!CommodityID Then
   Call Grid_Head_GCTD
End If
mCommodityID = RsMst_Commodity!CommodityID

Call TableMst_CommodityDetail("Where CommodityId = " & mCommodityID & " and ExchangeTypeId = " & mExchangeTypeID)
 If RsMst_CommodityDetail.RecordCount > 0 Then
    TxtCGunnyWt = RsMst_CommodityDetail!GunnyWt
    TxtCStdDed = IIf(IsNull(RsMst_CommodityDetail!StdDeduction), "", RsMst_CommodityDetail!StdDeduction)
    TxtCLotSize = IIf(IsNull(RsMst_CommodityDetail!LotSize), "", RsMst_CommodityDetail!LotSize)
    TxtCVariation = IIf(IsNull(RsMst_CommodityDetail!Variation), "", RsMst_CommodityDetail!Variation)
 Else
    MsgBox "No Record Found For Exchange" & CmbExchangeTypeID.Text & " - " & CmbCommodityID.Text
    CmbCommodityID.SetFocus
    Exit Sub
 End If
'SSTab1.Tab = 0
End Sub
'----
Private Sub CmdExchangeTypeID_Click()
    Frm_Mst_ExchangeType.Show
End Sub
'----
Private Sub CmbExchangeTypeID_GotFocus()
tmp = CmbExchangeTypeID.Text
Call TableMst_ExchangeType
CmbExchangeTypeID.Clear
If RsMst_ExchangeType.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_ExchangeType.RecordCount
    CmbExchangeTypeID.AddItem RsMst_ExchangeType!ExchangeType
    RsMst_ExchangeType.MoveNext
Next
CmbExchangeTypeID.Text = tmp
End Sub
'----
Private Sub CmbExchangeTypeID_LostFocus()
If CmbExchangeTypeID.Text = "" Then
   Exit Sub
End If
RsMst_ExchangeType.MoveFirst
RsMst_ExchangeType.Find "ExchangeType = '" & CmbExchangeTypeID.Text & "'"
If RsMst_ExchangeType.EOF Then
   MsgBox "Invalid  selection "
   CmbExchangeTypeID.SetFocus
   Exit Sub
End If

If mExchangeTypeID <> RsMst_ExchangeType!ExchangeTypeID Then
   Call Grid_Head_Client
   Call Grid_Head_GCTD
   Call ClearClientTabControl
End If
mExchangeTypeID = RsMst_ExchangeType!ExchangeTypeID


CmbCommodityID.Text = ""
TxtCGunnyWt = ""
TxtCStdDed = ""
TxtCLotSize = ""
TxtCVariation = ""
If mExchangeTypeID = 1 Then
   CmbNonNcdexClient.Width = 10095
   CmbNonNcdexClient.Top = TxtClientID.Top
   CmbNonNcdexClient.Left = TxtClientID.Left
   Label9.Visible = False
   Label10.Visible = False
   Label11.Visible = False
   
   CmbNonNcdexClient.Visible = True
Else
   Label9.Visible = True
   Label10.Visible = True
   Label11.Visible = True
   CmbNonNcdexClient.Visible = False
End If
Call Check1_Click
End Sub
'----
Private Sub CmdContractID_Click()
    FrmMst_StorageContract.Show
End Sub
'----
Private Sub CmbContractID_GotFocus()
tmp = CmbContractID.Text
Call TableMst_StorageContract
CmbContractID.Clear
If RsMst_StorageContract.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_StorageContract.RecordCount
    CmbContractID.AddItem RsMst_StorageContract!ContractID
    RsMst_StorageContract.MoveNext
Next
CmbContractID.Text = tmp
End Sub
'----
Private Sub CmbContractID_LostFocus()
If CmbContractID.Text = "" Then
   Exit Sub
End If
RsMst_StorageContract.MoveFirst
RsMst_StorageContract.Find "ContractID = '" & CmbContractID.Text & "'"
If RsMst_StorageContract.EOF Then
   MsgBox "Invalid  selection "
   CmbContractID.SetFocus
   Exit Sub
End If
mContractID = RsMst_StorageContract!ContractID
End Sub
'----

Private Sub TxtGCTDGrossWt_LostFocus()
If TxtGCTDGrossWt = "" Then Exit Sub

If IsNumeric(TxtGCTDGrossWt) = False Then
   MsgBox "Invalid Number format !!!"
   TxtGCTDGrossWt.SetFocus
   Exit Sub
End If
Call CalDmatWt

End Sub

Private Sub TxtGCTDNoofBags_LostFocus()
If TxtGCTDNoofBags = "" Then Exit Sub

If IsNumeric(TxtGCTDNoofBags) = False Then
   MsgBox "Invalid Number format !!!"
   TxtGCTDNoofBags.SetFocus
   Exit Sub
End If
Call CalDmatWt
End Sub

Private Sub CalDmatWt()

TxtGCTDGunnyWt = Format((Val(TxtGCTDNoofBags) * Val(TxtCGunnyWt)) / 1000, "#######0.000")
TxtAdjGrossWt = Val(TxtGCTDGrossWt) + Val(TxtAdjWt)
tmp = Val(TxtAdjGrossWt) - Val(TxtGCTDGunnyWt)
TxtGCTDStdDeduction = Format(((tmp) * Val(TxtCStdDed)) / 100, "#######0.000")
TxtGCTDDematWt = Format(Val(TxtAdjGrossWt) - (Val(TxtGCTDStdDeduction) + Val(TxtGCTDGunnyWt)), "#######0.000")


End Sub
Private Sub TxtTTDTareWt_LostFocus()
If TxtTTDTareWt = "" Then Exit Sub
If Not IsNumeric(TxtTTDTareWt) Then
   MsgBox "Invalid Number Format"
   TxtTTDTareWt.SetFocus
End If
End Sub

Private Sub TxtTTDGrossWt_LostFocus()
If TxtTTDGrossWt = "" Then Exit Sub
If Not IsNumeric(TxtTTDGrossWt) Then
   MsgBox "Invalid Number Format"
   TxtTTDGrossWt.SetFocus
End If

End Sub

Private Sub TxtTTDTotalWt_LostFocus()
If TxtTTDTotalWt = "" Then Exit Sub
If Not IsNumeric(TxtTTDTotalWt) Then
   MsgBox "Invalid Number Format"
   TxtTTDTotalWt.SetFocus
End If
End Sub
