VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmTxn_Spraying 
   Caption         =   "Spraying"
   ClientHeight    =   8250
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8880
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8250
   ScaleWidth      =   8880
   WindowState     =   2  'Maximized
   Begin VB.VScrollBar VScroll1 
      Height          =   9015
      Left            =   14880
      TabIndex        =   72
      Top             =   0
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   2880
      TabIndex        =   71
      Top             =   9960
      Width           =   10935
   End
   Begin VB.Frame Frame4 
      Height          =   9975
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   15015
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
         Height          =   2325
         Left            =   0
         TabIndex        =   59
         Top             =   6840
         Visible         =   0   'False
         Width           =   14595
         Begin VB.ComboBox CmbSBillableFlag 
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
            ItemData        =   "FrmTxn_Spraying.frx":0000
            Left            =   9315
            List            =   "FrmTxn_Spraying.frx":000A
            Sorted          =   -1  'True
            TabIndex        =   101
            Top             =   1320
            Width           =   1365
         End
         Begin VB.TextBox TxtSBillQty 
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
            Left            =   7920
            MaxLength       =   12
            TabIndex        =   98
            Top             =   1320
            Width           =   1335
         End
         Begin VB.ComboBox CmbSBillGenerated 
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
            ItemData        =   "FrmTxn_Spraying.frx":0017
            Left            =   6480
            List            =   "FrmTxn_Spraying.frx":0021
            Sorted          =   -1  'True
            TabIndex        =   97
            Top             =   1305
            Width           =   1365
         End
         Begin VB.ComboBox CmbSTreatmentType 
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
            Left            =   11400
            Sorted          =   -1  'True
            TabIndex        =   38
            Top             =   540
            Width           =   3045
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
            Left            =   12405
            TabIndex        =   41
            Top             =   1665
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
            Left            =   13425
            TabIndex        =   42
            Top             =   1665
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
            ItemData        =   "FrmTxn_Spraying.frx":002E
            Left            =   3210
            List            =   "FrmTxn_Spraying.frx":0030
            Sorted          =   -1  'True
            TabIndex        =   35
            Top             =   555
            Width           =   2280
         End
         Begin VB.ComboBox CmbSGodownNo 
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
            Left            =   9000
            Sorted          =   -1  'True
            TabIndex        =   37
            Top             =   555
            Width           =   2325
         End
         Begin VB.ComboBox CmbSCMPName 
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
            Left            =   5520
            Sorted          =   -1  'True
            TabIndex        =   36
            Top             =   555
            Width           =   3405
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
            Left            =   60
            Sorted          =   -1  'True
            TabIndex        =   39
            Top             =   1320
            Width           =   4485
         End
         Begin VB.TextBox TxtSSprayingID 
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
            Left            =   4560
            MaxLength       =   12
            TabIndex        =   40
            Top             =   1305
            Width           =   1770
         End
         Begin MSComCtl2.DTPicker STo 
            Height          =   375
            Left            =   1695
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   548
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
            Format          =   107413505
            CurrentDate     =   36494
         End
         Begin MSComCtl2.DTPicker SFrom 
            Height          =   375
            Left            =   60
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   555
            Width           =   1545
            _ExtentX        =   2725
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
            Format          =   107413505
            CurrentDate     =   36494
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
            Height          =   2775
            Left            =   120
            TabIndex        =   60
            Top             =   2160
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   4895
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
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "Billable Flag"
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
            Left            =   9435
            TabIndex        =   102
            Top             =   990
            Width           =   1125
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            Caption         =   "Actual Qty"
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
            Left            =   8130
            TabIndex        =   100
            Top             =   990
            Width           =   900
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            Caption         =   "Bill Generated"
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
            Left            =   6525
            TabIndex        =   99
            Top             =   990
            Width           =   1275
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Treatment Type"
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
            Left            =   12202
            TabIndex        =   89
            Top             =   240
            Width           =   1440
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
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
            Height          =   240
            Left            =   1995
            TabIndex        =   69
            Top             =   255
            Width           =   735
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
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
            Left            =   315
            TabIndex        =   68
            Top             =   255
            Width           =   945
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
            Left            =   3900
            TabIndex        =   67
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
            Left            =   6375
            TabIndex        =   66
            Top             =   1815
            Width           =   90
         End
         Begin VB.Label Label23 
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
            Left            =   4080
            TabIndex        =   65
            Top             =   255
            Width           =   765
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "Godown"
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
            Left            =   9780
            TabIndex        =   64
            Top             =   255
            Width           =   750
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
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
            Height          =   240
            Left            =   7005
            TabIndex        =   63
            Top             =   255
            Width           =   435
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
            TabIndex        =   62
            Top             =   990
            Width           =   1320
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Spraying ID"
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
            Left            =   4920
            TabIndex        =   61
            Top             =   990
            Width           =   1050
         End
      End
      Begin VB.Frame ButtonFrm 
         Height          =   1155
         Left            =   120
         TabIndex        =   70
         Top             =   5160
         Width           =   14520
         Begin VB.CommandButton CmdPrint 
            Caption         =   "Print FCR"
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
            Left            =   11595
            TabIndex        =   91
            Top             =   615
            Width           =   2295
         End
         Begin VB.CommandButton CmdPrintPreview 
            Caption         =   "Preview FCR"
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
            Left            =   11595
            TabIndex        =   90
            Top             =   240
            Width           =   2295
         End
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
            TabIndex        =   17
            Top             =   240
            Width           =   2295
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
            Left            =   2415
            TabIndex        =   13
            Top             =   240
            Width           =   2295
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
            Left            =   4710
            TabIndex        =   18
            Top             =   240
            Width           =   2295
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
            Left            =   7005
            TabIndex        =   19
            Top             =   240
            Width           =   2295
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
            Left            =   9300
            TabIndex        =   15
            Top             =   240
            Width           =   2295
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
            TabIndex        =   20
            Top             =   615
            Width           =   2295
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
            Left            =   2415
            TabIndex        =   21
            Top             =   615
            Width           =   2295
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
            Left            =   4710
            TabIndex        =   22
            Top             =   615
            Width           =   2295
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
            Left            =   7005
            TabIndex        =   23
            Top             =   615
            Width           =   2295
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
            Left            =   9300
            TabIndex        =   14
            Top             =   615
            Width           =   2295
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
         Height          =   5055
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   14535
         Begin VB.ComboBox CmbBillableFlag 
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
            ItemData        =   "FrmTxn_Spraying.frx":0032
            Left            =   1455
            List            =   "FrmTxn_Spraying.frx":003C
            Sorted          =   -1  'True
            TabIndex        =   94
            Top             =   2790
            Width           =   1920
         End
         Begin VB.TextBox TxtAddress 
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
            Height          =   735
            Left            =   2145
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   92
            TabStop         =   0   'False
            Top             =   5160
            Width           =   5040
         End
         Begin VB.TextBox TxtArea 
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
            Left            =   12120
            Locked          =   -1  'True
            MaxLength       =   12
            TabIndex        =   86
            TabStop         =   0   'False
            Top             =   1845
            Width           =   1800
         End
         Begin VB.TextBox TxtVendorQty 
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
            Left            =   4920
            Locked          =   -1  'True
            TabIndex        =   74
            TabStop         =   0   'False
            Top             =   4005
            Width           =   1800
         End
         Begin VB.TextBox TxtMappID 
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
            Left            =   12135
            Locked          =   -1  'True
            TabIndex        =   73
            TabStop         =   0   'False
            Top             =   4005
            Width           =   1800
         End
         Begin VB.TextBox TxtActualQty 
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
            Left            =   4935
            MaxLength       =   12
            TabIndex        =   7
            Top             =   1845
            Width           =   1800
         End
         Begin MSComCtl2.DTPicker TxtSprayingDate 
            Height          =   360
            Left            =   4920
            TabIndex        =   1
            Top             =   315
            Width           =   1800
            _ExtentX        =   3175
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
            Format          =   107413505
            CurrentDate     =   39884
         End
         Begin VB.CommandButton CmdCMP 
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
            Left            =   14010
            TabIndex        =   85
            TabStop         =   0   'False
            Top             =   795
            Width           =   375
         End
         Begin VB.TextBox TxtBalQty 
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
            Left            =   12120
            Locked          =   -1  'True
            MaxLength       =   12
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   1320
            Width           =   1800
         End
         Begin VB.CommandButton CmdTreatmentType 
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
            Left            =   6765
            TabIndex        =   81
            TabStop         =   0   'False
            Top             =   2355
            Width           =   375
         End
         Begin VB.TextBox TxtTotAmount 
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
            Left            =   12135
            Locked          =   -1  'True
            MaxLength       =   12
            TabIndex        =   11
            Top             =   2355
            Width           =   1800
         End
         Begin VB.TextBox TxtActualArea 
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
            Left            =   8640
            MaxLength       =   12
            TabIndex        =   8
            Top             =   1845
            Width           =   1920
         End
         Begin VB.TextBox TxtNoofBags 
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
            Left            =   8640
            Locked          =   -1  'True
            MaxLength       =   12
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   1320
            Width           =   1920
         End
         Begin VB.TextBox TxtRate 
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
            Left            =   8640
            Locked          =   -1  'True
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   2355
            Width           =   1920
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
            Left            =   8640
            Sorted          =   -1  'True
            TabIndex        =   4
            Top             =   795
            Width           =   5295
         End
         Begin VB.TextBox TxtBillNo 
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
            Left            =   8640
            Locked          =   -1  'True
            TabIndex        =   79
            TabStop         =   0   'False
            Top             =   4005
            Width           =   1920
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
            Left            =   8640
            Locked          =   -1  'True
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   4575
            Width           =   5295
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
            Left            =   8640
            Sorted          =   -1  'True
            TabIndex        =   2
            Top             =   315
            Width           =   5295
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
            Left            =   1440
            Sorted          =   -1  'True
            TabIndex        =   3
            Top             =   795
            Width           =   5295
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
            Height          =   360
            Left            =   6810
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   795
            Width           =   375
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
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   30
            Top             =   795
            Visible         =   0   'False
            Width           =   5295
         End
         Begin VB.ComboBox CmbTreatmentType 
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
            Left            =   1440
            Sorted          =   -1  'True
            TabIndex        =   9
            Top             =   2355
            Width           =   5295
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
            Left            =   1440
            Sorted          =   -1  'True
            TabIndex        =   5
            Top             =   1320
            Width           =   5295
         End
         Begin VB.TextBox TxtSprayingID 
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
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   315
            Width           =   1800
         End
         Begin VB.TextBox TxtActualBags 
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
            Left            =   1440
            MaxLength       =   12
            TabIndex        =   6
            Top             =   1845
            Width           =   1920
         End
         Begin VB.TextBox TxtVendorRate 
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
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   75
            TabStop         =   0   'False
            Top             =   4005
            Width           =   1800
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
            Height          =   660
            Left            =   1440
            MaxLength       =   250
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   12
            Top             =   3240
            Width           =   12495
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
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   4575
            Width           =   5295
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
            Height          =   360
            Left            =   14010
            TabIndex        =   84
            TabStop         =   0   'False
            Top             =   315
            Width           =   375
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
            Left            =   8640
            Locked          =   -1  'True
            TabIndex        =   31
            Top             =   315
            Visible         =   0   'False
            Width           =   5295
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
            Left            =   8640
            Locked          =   -1  'True
            TabIndex        =   32
            Top             =   795
            Visible         =   0   'False
            Width           =   5295
         End
         Begin VB.TextBox TxtTreatmentType 
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
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   82
            Top             =   2355
            Visible         =   0   'False
            Width           =   5295
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
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   43
            Top             =   1320
            Visible         =   0   'False
            Width           =   5295
         End
         Begin VB.TextBox TxtBillableFlag 
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
            Left            =   1455
            MaxLength       =   12
            TabIndex        =   96
            Top             =   2790
            Width           =   1920
         End
         Begin MSComCtl2.DTPicker TxtLastSprayingDate 
            Height          =   360
            Left            =   8640
            TabIndex        =   103
            Top             =   2805
            Width           =   1950
            _ExtentX        =   3440
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
            Format          =   107413505
            CurrentDate     =   39884
         End
         Begin VB.Label Label30 
            Caption         =   "Last Spraying Date"
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
            Left            =   7320
            TabIndex        =   104
            Top             =   2753
            Width           =   1605
         End
         Begin VB.Label Label16 
            Caption         =   "Billable Flag"
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
            TabIndex        =   95
            Top             =   2835
            Width           =   1380
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
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
            Height          =   240
            Left            =   120
            TabIndex        =   93
            Top             =   5205
            Width           =   1560
         End
         Begin VB.Label Label14 
            Caption         =   "Actual Area (in Sq.Ft)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   7320
            TabIndex        =   88
            Top             =   1755
            Width           =   1305
         End
         Begin VB.Label Label13 
            Caption         =   "Area (in Sq.Ft)"
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
            Left            =   10800
            TabIndex        =   87
            Top             =   1785
            Width           =   1065
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Bill No."
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
            Left            =   7320
            TabIndex        =   80
            Top             =   4065
            Width           =   630
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            Caption         =   "Vendor Rate"
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
            TabIndex        =   78
            Top             =   4065
            Width           =   1140
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            Caption         =   "Vendor Quantity"
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
            TabIndex        =   77
            Top             =   4065
            Width           =   1425
         End
         Begin VB.Label Label38 
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
            Left            =   10800
            TabIndex        =   76
            Top             =   4065
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
            TabIndex        =   50
            Top             =   3450
            Width           =   825
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
            Left            =   7320
            TabIndex        =   49
            Top             =   4635
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
            Left            =   120
            TabIndex        =   48
            Top             =   4635
            Width           =   855
         End
         Begin VB.Label Label10 
            Caption         =   "Actual No of Bags"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   120
            TabIndex        =   45
            Top             =   1770
            Width           =   1200
         End
         Begin VB.Label Label11 
            Caption         =   "Actual Bal Qty"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   3480
            TabIndex        =   44
            Top             =   1755
            Width           =   1380
         End
         Begin VB.Label Label8 
            Caption         =   "Spraying Date"
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
            Left            =   3360
            TabIndex        =   53
            Top             =   345
            Width           =   1335
         End
         Begin VB.Label Label4 
            Caption         =   "Spraying ID"
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
            TabIndex        =   52
            Top             =   345
            Width           =   1095
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
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
            Height          =   240
            Left            =   7320
            TabIndex        =   55
            Top             =   855
            Width           =   495
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "No of Bags"
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
            Left            =   7320
            TabIndex        =   47
            Top             =   1380
            Width           =   1005
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
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
            Height          =   240
            Left            =   10800
            TabIndex        =   46
            Top             =   1380
            Width           =   1095
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Godown"
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
            TabIndex        =   54
            Top             =   1380
            Width           =   855
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Amount"
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
            Left            =   10800
            TabIndex        =   58
            Top             =   2415
            Width           =   675
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Rate"
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
            Left            =   7320
            TabIndex        =   57
            Top             =   2415
            Width           =   435
         End
         Begin VB.Label Label12 
            Caption         =   "Treatment Type"
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
            Left            =   120
            TabIndex        =   83
            Top             =   2295
            Width           =   1065
         End
         Begin VB.Label Label1 
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
            TabIndex        =   56
            Top             =   840
            Width           =   840
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
            Left            =   7320
            TabIndex        =   51
            Top             =   375
            Width           =   900
         End
      End
   End
End
Attribute VB_Name = "FrmTxn_Spraying"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mVendorID, mGodownID, mLocationID, mCommodityID, mSVendorID, mSGodownID, mSLocationID, mSCommodityID As Variant
Dim mCMPID As Variant
Dim Edit_Flg, mSCMPID As Variant
Dim mTreatmentTypeID, mAgreementDetailID As Variant

Private Sub AddCmd_Click()

Edit_Flg = "A"

Frame1.Enabled = True
TxtSprayingID.Text = ""
CmbVendorID.Text = ""
CmbLocationID.Text = ""
CmbTreatmentType.Text = ""
CmbCMPName.Text = ""
CmbGodownNo.Text = ""
CmbBillableFlag.Text = ""

TxtVendorID.Text = ""
TxtLocationID.Text = ""
TxtTreatmentType.Text = ""
TxtCMPName.Text = ""
TxtGodownNo.Text = ""

CmdVendor.Visible = True
TxtVendorID.Visible = False
TxtLocationID.Visible = False
TxtTreatmentType.Visible = False
TxtCMPName.Visible = False
TxtGodownNo.Visible = False
TxtBillableFlag.Visible = False

CmbVendorID.Visible = True
CmbLocationID.Visible = True
CmbTreatmentType.Visible = True
CmbCMPName.Visible = True
CmbGodownNo.Visible = True
CmbBillableFlag.Visible = True

CmdVendor.Visible = True
CmdLocation.Visible = True
CmdCMP.Visible = True
CmdTreatmentType.Visible = True

CmbVendorID.Visible = True
CmbLocationID.Visible = True
CmbTreatmentType.Visible = True

TxtRate.Text = ""
TxtNoofBags.Text = ""
TxtBalQty.Text = ""
TxtArea.Text = ""
TxtActualBags.Text = ""
TxtActualQty.Text = ""
TxtActualArea.Text = ""

TxtRemark.Text = ""
TxtTotAmount.Text = ""
TxtVendorQty = ""
TxtVendorRate = ""
TxtBillNo = ""
TxtMappID = ""

TxtCreated = ""
TxtUpdated = ""

TxtSprayingDate.Value = Date


Call GButtonLock(Me, False)
If RsTxn_Spraying.RecordCount > 0 Then
   TxtSprayingDate.SetFocus
End If


End Sub

Private Sub CmbBillableFlag_LostFocus()

If CmbBillableFlag.Text = "" Then Exit Sub

If LCase(CmbBillableFlag.Text) <> "no" And LCase(CmbBillableFlag.Text) <> "yes" Then
   MsgBox "Invalid Selection!!!"
   CmbBillableFlag.SetFocus
   Exit Sub
End If

End Sub

Private Sub CmbSBillableFlag_LostFocus()

If CmbSBillableFlag.Text = "" Then Exit Sub

If LCase(CmbSBillableFlag.Text) <> "no" And LCase(CmbSBillableFlag.Text) <> "yes" Then
   MsgBox "Invalid Selection!!!"
   CmbSBillableFlag.SetFocus
   Exit Sub
End If

End Sub

Private Sub CmbSBillGenerated_LostFocus()
If CmbSBillGenerated.Text = "" Then Exit Sub

If LCase(CmbSBillGenerated.Text) <> "no" And LCase(CmbSBillGenerated.Text) <> "yes" Then
   MsgBox "Invalid Selection!!!"
   CmbSBillGenerated.SetFocus
   Exit Sub
End If


End Sub

Private Sub CmbSTreatmentType_GotFocus()
tmp = CmbSTreatmentType.Text

Call TableMst_TreatmentType

CmbSTreatmentType.Clear
If RsMst_TreatmentType.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_TreatmentType.RecordCount
   CmbSTreatmentType.AddItem RsMst_TreatmentType!TreatmentType
   RsMst_TreatmentType.MoveNext
Next
CmbSTreatmentType.Text = tmp
End Sub

Private Sub CmbSTreatmentType_LostFocus()
If CmbSTreatmentType.Text = "" Then
   Exit Sub
End If
RsMst_TreatmentType.MoveFirst
RsMst_TreatmentType.Find "TreatmentType= '" & CmbSTreatmentType.Text & "'"
If RsMst_TreatmentType.EOF Then
   MsgBox "Invalid selection!!!"
   CmbSTreatmentType.SetFocus
   Exit Sub
End If

'mTreatmentTypeID = RsMst_TreatmentType!TreatmentTypeID
End Sub

Private Sub CmbTreatmentType_GotFocus()
If CmbVendorID.Text = "" Then
   MsgBox "Select Vendor!!!"
   CmbVendorID.SetFocus
   Exit Sub
ElseIf CmbLocationID.Text = "" Then
   MsgBox "Select Location!!!"
   CmbLocationID.SetFocus
   Exit Sub
End If

tmp1 = CmbTreatmentType.Text
'Call TableMst_TreatmentType
tmp = " Select MFT.DetailID, MTT.TreatmentType, MFT.RateperMT from Mst_FumigationTreatmentDetail MFT "
tmp = tmp & " Inner Join MSt_FumigationAgreement MFA on MFA.AgreementID = MFT.AgreementID"
tmp = tmp & " Inner Join MSt_TreatmentType MTT on MTT.TreatmentTypeID = MFT.TreatmentTypeID"
tmp = tmp & " Where MFA.Status = 'A' And MFA.Type = 'S' And MFA.LocationID = " & mLocationID & " And MFA.VendorID = " & mVendorID
tmp = tmp & " And '" & Format(TxtSprayingDate.Value, "yyyy-mm-dd") & "' between MFA.StartDate and MFA.ExpiryDate"
Call Tmp7Table

CmbTreatmentType.Clear
If TmpRs7.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To TmpRs7.RecordCount
    CmbTreatmentType.AddItem TmpRs7!TreatmentType
    TmpRs7.MoveNext
Next
CmbTreatmentType.Text = tmp1
End Sub

Private Sub CmbTreatmentType_LostFocus()
If CmbTreatmentType.Text = "" Then
   Exit Sub
End If
TmpRs7.MoveFirst
TmpRs7.Find "TreatmentType= '" & CmbTreatmentType.Text & "'"
If TmpRs7.EOF Then
   MsgBox "Invalid selection!!!"
   CmbTreatmentType.SetFocus
   Exit Sub
End If

mAgreementDetailID = TmpRs7!DetailID
TxtRate.Text = IIf(IsNull(TmpRs7!RateperMT), 0, TmpRs7!RateperMT)
Call TxtRate_Lostfocus
End Sub

Private Sub CmbVendorID_GotFocus()
If CmbLocationID.Text = "" Then
   MsgBox "Select Location!!!"
   CmbLocationID.SetFocus
   Exit Sub
End If
tmp = CmbVendorID.Text
If Edit_Flg = "A" Then
   Call TableMst_Vendor("Where VendorID in (Select VendorID from Mst_FumigationAgreement Where Type = 'S' And Status = 'A' And LocationID = " & Val(mLocationID) & ")") '("Where Flag = 'Y'")
Else
   Call TableMst_Vendor("Where VendorID = " & Val(mVendorID) & " Or  VendorID in (Select VendorID from Mst_FumigationAgreement Where Type = 'S' And Status = 'A' And LocationID = " & Val(mLocationID) & ")") '("Where Flag = 'Y'")
End If
CmbVendorID.Clear
If RsMst_Vendor.RecordCount = 0 Then
   MsgBox "No Vendor Found!!!"
   CmbVendorID.SetFocus
   'RsMst_Vendor.MoveFirst
   Exit Sub
End If
For I = 1 To RsMst_Vendor.RecordCount
    CmbVendorID.AddItem RsMst_Vendor!VendorName ' & "~" & RsMst_Vendor!VendorNo
    RsMst_Vendor.MoveNext
Next

CmbVendorID.Text = tmp

End Sub

Private Sub CmbVendorID_LostFocus()

If CmbVendorID.Text = "" Then
   mVendorID = Null
   Exit Sub
End If

RsMst_Vendor.MoveFirst
RsMst_Vendor.Find "VendorName = '" & CmbVendorID.Text & "'"

If RsMst_Vendor.EOF Then
   MsgBox "Invalid Selection "
   CmbVendorID.SetFocus
   Exit Sub
End If
mVendorID = RsMst_Vendor!VendorID

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
   CmbCMPName.Text = ""
   CmbGodownNo.Text = ""
   mLocationID = 0
   CmbVendorID.Text = ""
   TxtLastSprayingDate.Value = Date
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
   CmbVendorID.Text = ""
   TxtLastSprayingDate.Value = Date
End If

mLocationID = RsMst_Location!LocationID
TxtPreFix = RsMst_Location!SM_Prefix
End Sub

Private Sub CmbCMPName_GotFocus()
If CmbLocationID.Text = "" Then
   MsgBox "Select Location!!!"
   CmbLocationID.SetFocus
   Exit Sub
End If
tmp = CmbCMPName.Text
If Gen_WcCMP <> "" Then
   Call TableMst_CMP(" Where CloseDate is null And LocationID = " & mLocationID & " And " & Gen_WcCMP)
Else
   Call TableMst_CMP(" Where CloseDate is null And LocationID = " & mLocationID)
End If
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
   mCMPID = 0
   Exit Sub
End If

RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPName = '" & CmbCMPName.Text & "'"
If RsMst_CMP.EOF Then
   MsgBox "Invalid selection "
   CmbCMPName.SetFocus
   Exit Sub
End If
If mCMPID <> RsMst_CMP!CMPID Then
   CmbGodownNo.Text = ""
End If
mCMPID = RsMst_CMP!CMPID
'TxtPreFix = GetPreFix(RsMst_CMP!CMPID, "P")
End Sub

Private Sub CmbGodownNo_GotFocus()
If CmbCMPName.Text = "" Then
   MsgBox "Select CMP !!!! "
   CmbCMPName.SetFocus
   Exit Sub
End If
tmp = CmbGodownNo.Text
Call TableMst_Godown("Where CloseDate is null And CMPID = " & mCMPID)
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
mGodownID = RsMst_Godown!GodownID
Call GetBalanceStock(mGodownID)
Call GetLastSprayingDate(mGodownID)
End Sub
Private Sub GetLastSprayingDate(mGodownID_ As Variant)

tmp = "Select Max(SprayingDate) as SprayingDate  From Txn_Spraying "
tmp = tmp & " Where GodownID=" & mGodownID_ & " "

Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   'TxtLastSprayingDate.Value = Format(TmpRs1!SprayingDate, "dd-MMM-yyyy")
   If IsNull(TmpRs1!SprayingDate) = False Then
      TxtLastSprayingDate.Value = Format(TmpRs1!SprayingDate, "dd-MMM-yyyy")
   Else
      TxtLastSprayingDate.Value = TxtSprayingDate.Value
   End If
Else
   TxtLastSprayingDate.Value = TxtSprayingDate.Value
End If

End Sub


Private Sub GetBalanceStock(mGodownID_ As Variant)

tmp = " Select Area from Mst_Godown Where GodownID = " & mGodownID_
Call Tmp7Table
If TmpRs7.RecordCount > 0 Then
   TxtArea.Text = IIf(IsNull(TmpRs7!Area), 0, TmpRs7!Area)
Else
   TxtArea.Text = ""
End If

tmp = "Select GodownID,sum(BalanceBags)BalanceBags,sum(BalanceWeight)BalanceWeight from Mst_GSL"
tmp = tmp & " Where Flag not in('E','Z') And GodownID=" & mGodownID_
tmp = tmp & " Group By GodownID"

Call TmpTable
If TmpRs.RecordCount > 0 Then
   TxtNoofBags.Text = IIf(IsNull(TmpRs!BalanceBags), 0, TmpRs!BalanceBags)
   TxtBalQty.Text = IIf(IsNull(TmpRs!BalanceWeight), 0, TmpRs!BalanceWeight)
Else
   TxtNoofBags.Text = ""
   TxtBalQty.Text = ""
   MsgBox " Stock not found."
   CmbGodownNo.SetFocus
   Exit Sub
End If

End Sub

Private Sub CmdCMP_Click()
FrmMst_CMP.Show
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Txn_Spraying.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Txn_Spraying.xls")
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

Private Sub CmdLocation_Click()
FrmMst_Location.Show
End Sub

Private Sub CmdPrint_Click()

RsTxn_Spraying!Flag = "P"
RsTxn_Spraying!PrintedDate = Date
RsTxn_Spraying.Update

Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "FCR_" & TxtSprayingID.Text & ".xls")
Call Xls_Detail_Rpt1

End Sub

Private Sub CmdPrintPreview_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "FCR_" & TxtSprayingID.Text & ".xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "FCR_" & TxtSprayingID.Text & ".xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows

oWS.Shapes.AddPicture(App.Path & "\logo.gif", msoTrue, msoTrue, 10, 10, 40, 40).Select

tmp = "B" & 1 & ":i" & 1
oWS.Range(tmp).Merge
oWS.Cells(1, 2) = "National Collateral Management Services Limited"
oWS.Cells(1, 2).HorizontalAlignment = -4108
oWS.Range(tmp).Font.Bold = True

tmp = "B" & 2 & ":i" & 2
oWS.Range(tmp).Merge
oWS.Cells(2, 2) = "Lodha Supremus, Unit No 505 to 509 , 5th Floor,"
oWS.Cells(2, 2).HorizontalAlignment = -4108

tmp = "B" & 3 & ":i" & 3
oWS.Range(tmp).Merge
oWS.Cells(3, 2) = "Kanjurmarg (East), Mumbai-400042"
oWS.Cells(3, 2).HorizontalAlignment = -4108

tmp = "B" & 4 & ":i" & 4
oWS.Range(tmp).Merge
oWS.Cells(4, 2) = "Tel.No-022 40419191, Fax. No-022 40419193"
oWS.Cells(4, 2).HorizontalAlignment = -4108

'tmp = "a" & 7 & ":" & 7
'oWS.Range(tmp).Merge
mRow = 7
mtmpRow = mRow
tmp = "a" & mRow & ":b" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, 1) = "Spraying ID"
tmp = "c" & mRow & ":i" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, 3) = "'" & TxtSprayingID.Text

oWS.Cells(mRow, 1).Font.Bold = True

mRow = mRow + 1
tmp = "a" & mRow & ":b" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, 1) = "Spraying Date"
tmp = "c" & mRow & ":i" & mRow
oWS.Range(tmp).Merge

oWS.Cells(mRow, 3) = "'" & Format(TxtSprayingDate.Value, "dd-MMM-yyyy")
oWS.Cells(mRow, 1).Font.Bold = True

'mRow = mRow + 1
'tmp = "a" & mRow & ":b" & mRow
'oWS.Range(tmp).Merge
'oWS.Cells(mRow, 1) = "Next Fumigation Date"
'tmp = "c" & mRow & ":i" & mRow
'oWS.Range(tmp).Merge
'oWS.Cells(mRow, 3) = "'" & Format(TxtNextFumDate.Value, "dd-MMM-yyyy")
'oWS.Cells(mRow, 1).Font.Bold = True

mRow = mRow + 1
tmp = "a" & mRow & ":b" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, 1) = "Last Spraying Date"
tmp = "c" & mRow & ":i" & mRow
oWS.Range(tmp).Merge

oWS.Cells(mRow, 3) = "'" & Format(TxtLastSprayingDate.Value, "dd-MMM-yyyy")
oWS.Cells(mRow, 1).Font.Bold = True


mRow = mRow + 1
tmp = "a" & mRow & ":b" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, 1) = "Vendor Name"
tmp = "c" & mRow & ":i" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, 3) = TxtVendorID.Text
oWS.Cells(mRow, 1).Font.Bold = True

mRow = mRow + 1

tmp = "a" & mRow & ":b" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, 1) = "Location"
tmp = "c" & mRow & ":i" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, 3) = TxtLocationID.Text
oWS.Cells(mRow, 1).Font.Bold = True

mRow = mRow + 1

tmp = "a" & mRow & ":b" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, 1) = "CMP Name"
tmp = "c" & mRow & ":i" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, 3) = TxtCMPName.Text
oWS.Cells(mRow, 1).Font.Bold = True

mRow = mRow + 1

tmp = "a" & mRow & ":b" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, 1) = "Godown No"
tmp = "c" & mRow & ":i" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, 3) = "'" & TxtGodownNo.Text
oWS.Cells(mRow, 1).Font.Bold = True

mRow = mRow + 1

tmp = "a" & mRow & ":b" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, 1) = "Godown Address"
oWS.Cells(mRow, 1).VerticalAlignment = xlTop
tmp = "c" & mRow & ":i" & mRow
oWS.Range(tmp).Merge
oWS.Range(tmp).WrapText = True
oWS.Cells(mRow, 3) = TxtAddress.Text
oWS.Cells(mRow, 3).VerticalAlignment = xlTop
oWS.Cells(mRow, 3).RowHeight = 45
oWS.Cells(mRow, 1).Font.Bold = True

mRow = mRow + 1

tmp = "a" & mRow & ":b" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, 1) = "Rate per MT"
tmp = "c" & mRow & ":i" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, 3) = TxtRate.Text
oWS.Cells(mRow, 3).HorizontalAlignment = 2
oWS.Cells(mRow, 1).Font.Bold = True

mRow = mRow + 1

tmp = "a" & mRow & ":b" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, 1) = "Actual Qty"
tmp = "c" & mRow & ":i" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, 3) = TxtActualQty.Text
oWS.Cells(mRow, 3).HorizontalAlignment = 2
oWS.Cells(mRow, 1).Font.Bold = True

mRow = mRow + 1

tmp = "a" & mRow & ":b" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, 1) = "Actual Area (in Sq.Ft)"
tmp = "c" & mRow & ":i" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, 3) = TxtActualArea
oWS.Cells(mRow, 3).HorizontalAlignment = 2
oWS.Cells(mRow, 1).Font.Bold = True

mRow = mRow + 1


tmp = "a" & mRow & ":b" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, 1) = "Amount"
oWS.Cells(mRow, 3) = TxtTotAmount.Text
oWS.Cells(mRow, 3).HorizontalAlignment = 2
tmp = "c" & mRow & ":i" & mRow
oWS.Range(tmp).Merge

oWS.Cells(mRow, 1).Font.Bold = True

mRow = mRow + 1
tmp = "a" & mRow & ":b" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, 1) = "Remark"
oWS.Cells(mRow, 1).VerticalAlignment = xlTop

tmp = "c" & mRow & ":i" & mRow

oWS.Range(tmp).Merge
oWS.Range(tmp).WrapText = True
oWS.Cells(mRow, 3) = TxtRemark.Text
oWS.Cells(mRow, 3).VerticalAlignment = xlTop
oWS.Cells(mRow, 3).RowHeight = 45
oWS.Cells(mRow, 1).Font.Bold = True


'mRow = mRow + 2
'tmp = "a" & mRow & ":k" & mRow
'oWS.Range(tmp).Font.Bold = True
'oWS.Range(tmp).WrapText = True
'oWS.Cells(mRow, 1).RowHeight = 30
'mtmpRow = mRow
'
'oWS.Range("a:a").ColumnWidth = 8
'oWS.Range("b:b").ColumnWidth = 11
'oWS.Range("c:c").ColumnWidth = 7
'oWS.Range("d:d").ColumnWidth = 7
'oWS.Range("e:e").ColumnWidth = 9.5 '32.14
'oWS.Range("f:f").ColumnWidth = 8
'oWS.Range("g:g").ColumnWidth = 14
'oWS.Range("h:h").ColumnWidth = 7.75
'oWS.Range("i:i").ColumnWidth = 9.5
'oWS.Range("j:j").ColumnWidth = 8
'oWS.Range("k:k").ColumnWidth = 10
'
'For i = 1 To MSHFlexGrid1.Rows - 2
'    'c = 1
'    For c = 0 To MSHFlexGrid1.Cols - 2
'        If InStr(MSHFlexGrid1.TextMatrix(i, c), "/") > 0 Then
'           If IsDate(MSHFlexGrid1.TextMatrix(i, c)) Then
'              oWS.Cells(mRow, c + 1) = Format(MSHFlexGrid1.TextMatrix(i, c), "dd-MMM-yyyy")
'           Else
'              oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(i, c)
'           End If
'        Else
'           oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(i, c)
'        End If
'         'c = c + 1
'    Next
'    'TmpRs.MoveNext
'    mRow = mRow + 1
'Next

tmp = "a" & mtmpRow & ":i" & mRow
oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)


tmp = "a" & mRow + 3 & ":g" & mRow + 3
oWS.Range(tmp).Merge
oWS.Cells(mRow + 3, 1) = "For National Collateral Management Services Ltd."

tmp = "a" & mRow + 5 & ":g" & mRow + 5
oWS.Range(tmp).Merge
oWS.Cells(mRow + 5, 1) = "(Authorised Signatory)"

If Gen_ExcelPassword <> "" Then
   oWS.Protect Gen_ExcelPassword, , , , , , True, True
End If

'oWS.Columns.AutoFit
oWS.PageSetup.Orientation = xlPortrait
oWB.Save
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub

Private Sub CmdSearch_Click()
Dim Wc1 As String
Call Grid_Head1

If SFrom.Value > STo.Value Then
   MsgBox "'From Date' Must be less than 'To Date'!!!"
   SFrom.SetFocus
   Exit Sub
End If

tmp = "Select TF.SprayingID, TF.SprayingDate, MV.VendorName, ML.LocationName , MC.CMPName, MG.GodownNo, "
tmp = tmp & " TF.Rate, TF.NoofBags, TF.BalanceQty, TF.ActualNoofBags, TF.ActualQty , TF.Amount, "
tmp = tmp & " TF.Remark,TF.VendorRate,TF.VendorQuantity,TF.MappingID, TF.Area, TF.ActualArea, MTT.TreatmentType,TF.BillableFlag, TF.LastSprayingDate "
tmp = tmp & " From Txn_Spraying TF"
tmp = tmp & " Inner Join Mst_Vendor MV on TF.VendorID=MV.VendorID"
tmp = tmp & " Inner Join Mst_Godown MG on TF.GodownID=MG.GodownID"
tmp = tmp & " Inner Join Mst_CMP MC on MG.CMPID=MC.CMPID"
tmp = tmp & " Inner Join Mst_Location ML on MC.LocationID=ML.LocationID"
tmp = tmp & " Left Join Mst_FumigationTreatmentDetail MFT on MFT.DetailID =TF.AgreementDetailID "
tmp = tmp & " Left join Mst_TreatmentType MTT on MFT.TreatmentTypeID =MTT.TreatmentTypeID"

If Gen_DBType = "MDB" Then
   Wc1 = " Where TF.SprayingDate Between #" & Format(SFrom.Value, Gen_DatFormat) & "# And #" & Format(STo.Value, Gen_DatFormat) & "# "
Else
   Wc1 = " Where TF.SprayingDate Between '" & Format(SFrom.Value, Gen_DatFormat) & "' And '" & Format(STo.Value, Gen_DatFormat) & "' "
End If
If CmbSLocationID.Text <> "" Then
   Wc1 = Wc1 & " And ML.LocationName='" & CmbSLocationID.Text & "'"
Else
   If Gen_WcLocation <> "" Then
      Wc1 = Wc1 & " And ML." & Gen_WcLocation
   End If
End If
If CmbSCMPName.Text <> "" Then
   Wc1 = Wc1 & " And MC.CMPName='" & CmbSCMPName.Text & "'"
Else
   If Gen_WcCMP <> "" Then
      Wc1 = Wc1 & " And MC." & Gen_WcCMP
   End If
End If

If CmbSGodownNo.Text <> "" Then
   Wc1 = Wc1 & " And MG.GodownNo='" & CmbSGodownNo.Text & "'"
End If

If CmbSTreatmentType.Text <> "" Then
   Wc1 = Wc1 & " And MTT.TreatmentType ='" & CmbSTreatmentType.Text & "'"
End If

If CmbSVendorID.Text <> "" Then
   Wc1 = Wc1 & " And MV.VendorName='" & CmbSVendorID.Text & "'"
End If

If TxtSSprayingID.Text <> "" Then
   Wc1 = Wc1 & " And TF.SprayingID=" & Val(TxtSSprayingID.Text) & ""
End If

If LCase(CmbSBillGenerated.Text) = "no" Then
   Wc1 = Wc1 & " And (isnull(TF.MappingID,0) = 0 ) "
ElseIf LCase(CmbSBillGenerated.Text) = "yes" Then
   Wc1 = Wc1 & " And isnull(TF.MappingID,0) <> 0 "
End If

If TxtSBillQty <> "" Then
    If Not IsOprater(TxtSBillQty) Then
       Wc1 = GenWc(Wc1, TxtSBillQty, "TF.ActualQty", "N", "N")
    Else
       Wc1 = GenWc(Wc1, TxtSBillQty, "TF.ActualQty", "N", "I")
    End If
End If


If LCase(CmbSBillableFlag.Text) = "no" Then
   Wc1 = Wc1 & " And (isnull(TF.BillableFlag,'') = 'N' ) "
ElseIf LCase(CmbSBillableFlag.Text) = "yes" Then
   Wc1 = Wc1 & " And (isnull(TF.BillableFlag,'') = 'Y' ) "
End If

tmp = tmp & Wc1 & " Order By TF.SprayingID"

Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   MSHFlexGrid2.Rows = TmpRs1.RecordCount + 2
   For I = 1 To TmpRs1.RecordCount
       For C = 0 To MSHFlexGrid2.Cols - 1
           MSHFlexGrid2.TextMatrix(I, C) = IIf(IsNull(TmpRs1.Fields(C)), "", TmpRs1.Fields(C))
       Next
       MSHFlexGrid2.TextMatrix(I, 1) = Format(MSHFlexGrid2.TextMatrix(I, 1), "dd-MMM-yyyy")
       If LCase(MSHFlexGrid2.TextMatrix(I, 19)) = "n" Then
          MSHFlexGrid2.TextMatrix(I, 19) = "No"
       ElseIf LCase(MSHFlexGrid2.TextMatrix(I, 19)) = "y" Then
          MSHFlexGrid2.TextMatrix(I, 19) = "Yes"
       End If
       TmpRs1.MoveNext
   Next
   
End If

Label22.Caption = TmpRs1.RecordCount
Label22.Refresh


End Sub

Private Sub CmdTreatmentType_Click()
FrmMst_TreatmentType.Show
End Sub

Private Sub CmdVendor_Click()
FrmMst_Vendor.Show
End Sub

Private Sub Command2_Click()

End Sub

Private Sub DeleteCmd_Click()


Dim ans As Variant
On Error GoTo msgError

tmp = "Select max(SprayingDate)SprayingDate from Txn_Spraying Where GodownID=" & mGodownID & " And SprayingID <>" & Val(TxtSprayingID)
Call TmpTable
If TmpRs.RecordCount > 0 Then
   If IsNull(SprayingDate) = False Then
      If TmpRs!SprayingDate > TxtSprayingDate.Value Then
         MsgBox "You can't edit this record."
         Exit Sub
      End If
   End If
End If

If Val(TxtMappID) > 0 Then
   MsgBox "Mapping created for this fumigation, You can't edit this record."
   Exit Sub
End If

ans = MsgBox("Do you really want to DELETE this record??? ", vbYesNo)
If ans = vbYes Then
   
   RsTxn_Spraying.Delete
   RsTxn_Spraying.Update
   If RsTxn_Spraying.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsTxn_Spraying.MoveNext
   If RsTxn_Spraying.EOF() Then
      RsTxn_Spraying.MoveLast
   End If
   Call Indata
End If

Exit Sub
msgError:
MsgBox "Child record Present "
RsTxn_Spraying.CancelUpdate


End Sub

Private Sub EditCmd_Click()

tmp = "Select max(SprayingDate)SprayingDate from Txn_Spraying Where GodownID=" & mGodownID & " And SprayingID <>" & Val(TxtSprayingID)
Call TmpTable
If TmpRs.RecordCount > 0 Then
   If IsNull(SprayingDate) = False Then
      If TmpRs!SprayingDate > TxtSprayingDate.Value Then
         MsgBox "You can't edit this record."
         Exit Sub
      End If
   End If
End If

If Val(TxtMappID) > 0 Then
   MsgBox "Mapping created for this fumigation, You can't edit this record."
   Exit Sub
End If

If LCase(RsTxn_Spraying!Flag) = "p" Then
   MsgBox "FCR is printed no more edition is allowed in this record!!!"
   Exit Sub
End If

Edit_Flg = "E"

Call GButtonLock(Me, False)

Frame1.Enabled = True
TxtSprayingDate.Enabled = False
'TxtRate.Locked = True

'TxtVendorID.Visible = False
'TxtLocationID.Visible = False
'TxtCMPName.Visible = False
'TxtGodownNo.Visible = False

'CmbVendorID.Visible = True
'CmbLocationID.Visible = True
'CmbCMPName.Visible = True
'CmbGodownNo.Visible = True
'
'CmbVendorID.Enabled = False
'CmbLocationID.Enabled = False
'CmbCMPName.Enabled = False
'CmbGodownNo.Enabled = False


CmbVendorID.Text = TxtVendorID.Text
CmbLocationID.Text = TxtLocationID.Text
CmbTreatmentType.Text = TxtTreatmentType.Text
CmbCMPName.Text = TxtCMPName.Text
CmbGodownNo.Text = TxtGodownNo.Text

CmbTreatmentType.Visible = True
TxtTreatmentType.Visible = False

CmbBillableFlag.Visible = True
TxtBillableFlag.Visible = False
CmbBillableFlag.Text = TxtBillableFlag

End Sub

Private Sub ExitCmd_Click()
Unload Me
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
Frame1.Enabled = True
Call Grid_Head1
SFrom.Value = Date - 365
STo.Value = Date
Label22.Caption = ""
CmbSLocationID.Text = ""
CmbSTreatmentType.Text = ""
CmbSCMPName.Text = ""
CmbSGodownNo.Text = ""
CmbSVendorID.Text = ""
TxtSSprayingID.Text = ""

Call TableTxn_Spraying(" Order by SprayingID")
If RsTxn_Spraying.RecordCount = 0 Then
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
Dim I As Variant
I = MSHFlexGrid2.RowSel
If I <= 0 Then Exit Sub
If MSHFlexGrid2.TextMatrix(I, 0) = "" Then Exit Sub
RsTxn_Spraying.MoveFirst
RsTxn_Spraying.Find "SprayingID= " & Val(MSHFlexGrid2.TextMatrix(I, 0)) & ""
If Not RsTxn_Spraying.EOF Then
   Call Indata
   Frame0.Visible = False
End If

End Sub

Private Sub SaveCmd_Click()

If TxtSprayingDate.Value > Date Then
   MsgBox "Future Date not allowed!!"
   TxtSprayingDate.SetFocus
   Exit Sub
End If

If TxtSprayingDate.Value < TxtLastSprayingDate.Value Then
   MsgBox "Spraying Date must be same or greater than Last Spraying Date always !!!"
   TxtSprayingDate.SetFocus
   Exit Sub
End If

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

If CmbCMPName.Text = "" Then
   MsgBox "Select CMP Name!!"
   CmbCMPName.SetFocus
   Exit Sub
End If

If CmbGodownNo.Text = "" Then
   MsgBox "Select Godown!!"
   CmbGodownNo.SetFocus
   Exit Sub
End If

If CmbBillableFlag.Text = "" Then
   MsgBox "Select Billable Flag!!!"
   CmbBillableFlag.SetFocus
   Exit Sub
End If


If TxtActualBags.Text = "" Then
   MsgBox "Blank Actual Bags not Allowed!!"
   TxtActualBags.SetFocus
   Exit Sub
End If

If TxtActualQty.Text = "" Then
   MsgBox "Blank Actual Quantity not Allowed!!"
   TxtActualQty.SetFocus
   Exit Sub
End If
If Edit_Flg = "A" Then
   If TxtActualArea.Text = "" Then
      MsgBox "Blank Actual Area not Allowed!!"
      TxtActualArea.SetFocus
      Exit Sub
   End If
   
   If CmbTreatmentType.Text = "" Then
      MsgBox "Blank Treatment Type not Allowed!!"
      CmbTreatmentType.SetFocus
      Exit Sub
   End If
End If
If TxtRate.Text = "" Then
   MsgBox "Blank Rate not Allowed!!"
   TxtRate.SetFocus
   Exit Sub
End If

If TxtRemark.Text = "" Then
   MsgBox "Blank Remark not Allowed!!"
   TxtRemark.SetFocus
   Exit Sub
End If

If CheckSprayingDate = False Then
   MsgBox "Spraying Date should be Greater than Last Spraying Date!!"
   TxtSprayingDate.SetFocus
   Exit Sub
End If

If LCase(Gen_UserRole) = "power" Or LCase(Gen_UserRole) = "admin" Or LCase(Gen_UserRole) = "superadmin" Or LCase(Gen_UserRole) = "headcmg" Then
   
Else
   If TxtSprayingDate.Value < Date - 7 Then
      MsgBox " You can't enter spraying details less than 7 days, Contact CMG!!"
      Exit Sub
   End If
End If


If Edit_Flg = "A" Then
   RsTxn_Spraying.AddNew
   RsTxn_Spraying!SprayingID = getMaxSprayingID
   RsTxn_Spraying!G_UID = GetGUID
   TxtSprayingID.Text = RsTxn_Spraying!SprayingID
Else
   RsTxn_Spraying.MoveFirst
   RsTxn_Spraying.Find " SprayingID= '" & Val(TxtSprayingID.Text) & "'"
End If

RsTxn_Spraying!SprayingID = Val(TxtSprayingID.Text)
RsTxn_Spraying!SprayingDate = Format(TxtSprayingDate.Value, Gen_DatFormat)
RsTxn_Spraying!VendorID = mVendorID
RsTxn_Spraying!GodownID = mGodownID
If CmbTreatmentType.Text <> "" Then
   RsTxn_Spraying!AgreementDetailID = mAgreementDetailID
End If
RsTxn_Spraying!Remark = TxtRemark.Text
RsTxn_Spraying!Rate = Format(Val(TxtRate.Text), "#####0.00")

RsTxn_Spraying!NoofBags = Val(TxtNoofBags.Text)
RsTxn_Spraying!BalanceQty = Val(TxtBalQty.Text)
RsTxn_Spraying!Area = Val(TxtArea.Text)
RsTxn_Spraying!ActualNoOfBags = Val(TxtActualBags.Text)
RsTxn_Spraying!ActualQty = Val(TxtActualQty.Text)
RsTxn_Spraying!ActualArea = Val(TxtActualArea.Text)
RsTxn_Spraying!BillableFlag = Left(CmbBillableFlag.Text, 1)
RsTxn_Spraying!Amount = Format(Val(TxtTotAmount.Text), "#####0.00")
RsTxn_Spraying!LastSprayingDate = Format(TxtLastSprayingDate.Value, Gen_DatFormat)

If IsNull(RsTxn_Spraying!CreatedBy) = True Or RsTxn_Spraying!CreatedBy = "" Then
   RsTxn_Spraying!CreatedBy = Gen_UserLoginID
   RsTxn_Spraying!CreatedDate = Now
Else
   RsTxn_Spraying!UpdatedBy = Gen_UserLoginID
   RsTxn_Spraying!UpdatedDate = Now
End If

RsTxn_Spraying.Update

SaveCmd.Enabled = False
Call Indata

End Sub
Private Function getMaxSprayingID() As Double
Dim Retval As Double
tmp = "Select max(SprayingID) from Txn_Spraying"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
getMaxSprayingID = Retval
End Function
Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsTxn_Spraying.MoveNext
If RsTxn_Spraying.EOF Then
   RsTxn_Spraying.MoveLast
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
RsTxn_Spraying.MovePrevious
If RsTxn_Spraying.BOF Then
   RsTxn_Spraying.MoveFirst
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
RsTxn_Spraying.MoveLast
Call Indata
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False
End Sub
Private Sub FirstCmd_Click()
RsTxn_Spraying.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub

Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   If RsTxn_Spraying.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   RsTxn_Spraying.MoveFirst
   RsTxn_Spraying.Find "SprayingID = " & Val(TxtSprayingID) & ""
   If RsTxn_Spraying.EOF Then
      RsTxn_Spraying.MoveLast
   End If
   Call Indata
   Exit Sub
End If
Frame0.Height = Me.Height - 500
Frame0.Left = Frame1.Left - 100
Frame0.Top = Frame1.Top
Frame0.Width = Me.Width - 200
Frame0.Visible = True

MSHFlexGrid2.Width = Frame0.Width - 400
MSHFlexGrid2.Height = Frame0.Height - 2500

RsTxn_Spraying.MoveFirst
RsTxn_Spraying.Filter = ""

''===========comment on 7 july 2012 for not clearing search grid=============
'Call Grid_Head1
'SFrom.Value = Date - 365
'STo.Value = Date
'Label22.Caption = ""
'CmbSLocationID.Text = ""
'CmbSCMPName.Text = ""
'CmbSGodownNo.Text = ""
'CmbSVendorID.Text = ""
'TxtSSprayingID.Text = ""
''===========comment on 7 july 2012 for not clearing search grid=============

End Sub

Private Sub TxtActualBags_LostFocus()
If TxtActualBags = "" Then Exit Sub
If Not IsNumeric(TxtActualBags) Or Val(TxtActualBags) < 0 Then
   MsgBox "Invalid numeric format !!!! "
   TxtActualBags.SetFocus
End If

End Sub

Private Sub TxtActualQty_LostFocus()
If TxtActualQty = "" Then Exit Sub
If Not IsNumeric(TxtActualQty) Or Val(TxtActualQty) < 0 Then
   MsgBox "Invalid numeric format !!!! "
   TxtActualQty.SetFocus
End If
'TxtTotAmount.Text = Format(Val(TxtRate.Text) * Val(TxtActualArea.Text), "#####0.00")
End Sub

Private Sub TxtActualArea_LostFocus()
If TxtActualArea = "" Then Exit Sub
If Not IsNumeric(TxtActualArea) Or Val(TxtActualArea) < 0 Then
   MsgBox "Invalid numeric format !!!! "
   TxtActualArea.SetFocus
End If
TxtTotAmount.Text = Format(Val(TxtRate.Text) * Val(TxtActualArea.Text), "#####0.00")
End Sub

Private Sub TxtRate_Lostfocus()
If TxtRate = "" Then Exit Sub
If Not IsNumeric(TxtRate) Or Val(TxtRate) < 0 Then
   MsgBox "Invalid numeric format !!!! "
   TxtRate.SetFocus
End If
TxtTotAmount.Text = Format(Val(TxtRate.Text) * Val(TxtActualArea), "#####0.00")
End Sub
Public Sub Indata()

Frame1.Enabled = False

TxtSprayingDate.Enabled = True
CmbVendorID.Visible = False
CmbLocationID.Visible = False
CmbTreatmentType.Visible = False
CmbCMPName.Visible = False
CmbGodownNo.Visible = False

TxtVendorID.Visible = True
TxtLocationID.Visible = True
TxtTreatmentType.Visible = True
TxtCMPName.Visible = True
TxtGodownNo.Visible = True

CmdVendor.Visible = False
CmdLocation.Visible = False
CmdCMP.Visible = False
CmdTreatmentType.Visible = False

TxtSprayingID = RsTxn_Spraying!SprayingID
TxtSprayingDate.Value = RsTxn_Spraying!SprayingDate
TxtVendorID.Text = GetVendorName(RsTxn_Spraying!VendorID, "N")
mVendorID = RsTxn_Spraying!VendorID
TxtGodownNo.Text = GetGodownName(RsTxn_Spraying!GodownID, "N")
TxtAddress.Text = GetGodownName(RsTxn_Spraying!GodownID, "")
mGodownID = RsTxn_Spraying!GodownID
mCMPID = GetCMPIDFromGodown(RsTxn_Spraying!GodownID)
mLocationID = GetLocationIDFromCMP(Val(mCMPID))
mAgreementDetailID = IIf(IsNull(RsTxn_Spraying!AgreementDetailID), 0, RsTxn_Spraying!AgreementDetailID)
TxtCMPName.Text = GetCMPFromGodown(RsTxn_Spraying!GodownID)
TxtLocationID.Text = GetLocationFromCMPName(TxtCMPName.Text)
If mAgreementDetailID <> 0 Then
   tmp = " Select MTT.TreatmentType From Mst_FumigationTreatmentDetail MFT"
   'tmp = tmp & " Inner Join Mst_FumigationTreatmentDetail MFT on MFT.DetailID =TF.AgreementDetailID "
   tmp = tmp & " Inner join Mst_TreatmentType MTT on MFT.TreatmentTypeID =MTT.TreatmentTypeID"
   tmp = tmp & " Where MFT.DetailID = " & mAgreementDetailID
   Call Tmp7Table
   If TmpRs7.RecordCount > 0 Then
      TxtTreatmentType.Text = IIf(IsNull(TmpRs7!TreatmentType), "", TmpRs7!TreatmentType)
   Else
      TxtTreatmentType.Text = ""
   End If
Else
   TxtTreatmentType.Text = ""
End If
   
TxtRemark.Text = RsTxn_Spraying!Remark
TxtBillableFlag = IIf(IsNull(RsTxn_Spraying!BillableFlag), "", RsTxn_Spraying!BillableFlag)
TxtLastSprayingDate.Value = RsTxn_Spraying!LastSprayingDate

If LCase(TxtBillableFlag) = "y" Then
   TxtBillableFlag = "Yes"
ElseIf LCase(TxtBillableFlag) = "n" Then
   TxtBillableFlag = "No"
End If

CmbBillableFlag.Visible = False
TxtBillableFlag.Visible = True
CmbBillableFlag.Text = TxtBillableFlag

TxtNoofBags.Text = RsTxn_Spraying!NoofBags
TxtBalQty.Text = RsTxn_Spraying!BalanceQty
TxtArea.Text = IIf(IsNull(RsTxn_Spraying!Area), 0, RsTxn_Spraying!Area)
TxtActualBags.Text = RsTxn_Spraying!ActualNoOfBags
TxtActualQty.Text = RsTxn_Spraying!ActualQty
TxtActualArea.Text = IIf(IsNull(RsTxn_Spraying!ActualArea), 0, RsTxn_Spraying!ActualArea)

TxtRate.Text = Format(RsTxn_Spraying!Rate, "#####0.00")
TxtTotAmount.Text = Format(RsTxn_Spraying!Amount, "#####0.00")

TxtCreated = IIf(IsNull(RsTxn_Spraying!CreatedBy), "", RsTxn_Spraying!CreatedBy) & " :- " & IIf(IsNull(RsTxn_Spraying!CreatedDate), "", RsTxn_Spraying!CreatedDate)
TxtUpdated = IIf(IsNull(RsTxn_Spraying!UpdatedBy), "", RsTxn_Spraying!UpdatedBy) & " :- " & IIf(IsNull(RsTxn_Spraying!UpdatedDate), "", RsTxn_Spraying!UpdatedDate)


TxtVendorQty = IIf(IsNull(RsTxn_Spraying!VendorQuantity), "", RsTxn_Spraying!VendorQuantity)
TxtVendorRate = IIf(IsNull(RsTxn_Spraying!VendorRate), "", RsTxn_Spraying!VendorRate)

TxtMappID = IIf(IsNull(RsTxn_Spraying!MappingID), 0, RsTxn_Spraying!MappingID)
TxtBillNo.Text = ""
If Val(TxtMappID) > 0 Then
   tmp = "Select BillNo from Txn_FumigationApproval Where MappingID=" & RsTxn_Spraying!MappingID
   Call TmpTable
   If TmpRs.RecordCount > 0 Then
      TxtBillNo.Text = IIf(IsNull(TmpRs!BillNo), "", TmpRs!BillNo)
   End If
End If

If RsTxn_Spraying.RecordCount = 1 Then
   Call GButtonLock_1(Me, True)
Else
   Call GButtonLock(Me, True)
End If
End Sub

Private Sub Grid_Head1()

MSHFlexGrid2.Clear
MSHFlexGrid2.Rows = 2
MSHFlexGrid2.FixedRows = 1
MSHFlexGrid2.Cols = 21
MSHFlexGrid2.WordWrap = True
MSHFlexGrid2.TextMatrix(0, 0) = "Spraying ID"
MSHFlexGrid2.TextMatrix(0, 1) = "Spraying Date"
MSHFlexGrid2.TextMatrix(0, 2) = "Vendor Name"
MSHFlexGrid2.TextMatrix(0, 3) = "Location"
MSHFlexGrid2.TextMatrix(0, 4) = "CMP Name"
MSHFlexGrid2.TextMatrix(0, 5) = "Godown No"
MSHFlexGrid2.TextMatrix(0, 6) = "Rate"
MSHFlexGrid2.TextMatrix(0, 7) = "No of Bags"
MSHFlexGrid2.TextMatrix(0, 8) = "Balance Qty"
MSHFlexGrid2.TextMatrix(0, 9) = "Actual No of Bags"
MSHFlexGrid2.TextMatrix(0, 10) = "Actual Qty"
MSHFlexGrid2.TextMatrix(0, 11) = "Amount"
MSHFlexGrid2.TextMatrix(0, 12) = "Remark"
MSHFlexGrid2.TextMatrix(0, 13) = "Vendor Rate"
MSHFlexGrid2.TextMatrix(0, 14) = "Vendor Quantity"
MSHFlexGrid2.TextMatrix(0, 15) = "Mapping ID"
MSHFlexGrid2.TextMatrix(0, 16) = "Area (in Sq.Ft)"
MSHFlexGrid2.TextMatrix(0, 17) = "Actual Area (in Sq.Ft)"
MSHFlexGrid2.TextMatrix(0, 18) = "Treatment Type"
MSHFlexGrid2.TextMatrix(0, 19) = "Billable Flag"
MSHFlexGrid2.TextMatrix(0, 20) = "Last Spraying Date"

MSHFlexGrid2.ColWidth(0) = 1000
MSHFlexGrid2.ColWidth(1) = 1200
MSHFlexGrid2.ColWidth(2) = 2200
MSHFlexGrid2.ColWidth(3) = 2000
MSHFlexGrid2.ColWidth(4) = 1800
MSHFlexGrid2.ColWidth(5) = 1200
MSHFlexGrid2.ColWidth(6) = 1100
MSHFlexGrid2.ColWidth(7) = 1100
MSHFlexGrid2.ColWidth(8) = 1100
MSHFlexGrid2.ColWidth(9) = 1100
MSHFlexGrid2.ColWidth(10) = 1100
MSHFlexGrid2.ColWidth(11) = 1100
MSHFlexGrid2.ColWidth(12) = 2200
MSHFlexGrid2.RowHeight(0) = 800

End Sub

Private Sub CmbSVendorID_GotFocus()
tmp = CmbSVendorID.Text
Call TableMst_Vendor '("Where Flag = 'Y'")
CmbSVendorID.Clear
If RsMst_Vendor.RecordCount = 0 Then
   MsgBox "No Vendor Found!!!"
   CmbSVendorID.SetFocus
   'RsMst_Vendor.MoveFirst
   Exit Sub
End If
For I = 1 To RsMst_Vendor.RecordCount
    CmbSVendorID.AddItem RsMst_Vendor!VendorName ' & "~" & RsMst_Vendor!VendorNo
    RsMst_Vendor.MoveNext
Next

CmbSVendorID.Text = tmp

End Sub

Private Sub CmbSVendorID_LostFocus()

If CmbSVendorID.Text = "" Then
   mSVendorID = Null
   Exit Sub
End If

RsMst_Vendor.MoveFirst
RsMst_Vendor.Find "VendorName = '" & CmbSVendorID.Text & "'"

If RsMst_Vendor.EOF Then
   MsgBox "Invalid Selection "
   CmbSVendorID.SetFocus
   Exit Sub
End If
mSVendorID = RsMst_Vendor!VendorID

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
   CmbSCMPName.Text = ""
   CmbSGodownNo.Text = ""
   mSLocationID = 0
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbSLocationID.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection "
   CmbSLocationID.SetFocus
   Exit Sub
End If

If mSLocationID <> RsMst_Location!LocationID Then
   CmbSCMPName.Text = ""
   CmbSGodownNo.Text = ""
End If

mSLocationID = RsMst_Location!LocationID

End Sub

Private Sub CmbSCMPName_GotFocus()
If CmbSLocationID.Text = "" Then
   MsgBox "Select Location!!!"
   CmbSLocationID.SetFocus
   Exit Sub
End If
tmp = CmbSCMPName.Text
If Gen_WcCMP <> "" Then
   Call TableMst_CMP(" Where CloseDate is null And LocationID = " & mSLocationID & " And " & Gen_WcCMP)
Else
   Call TableMst_CMP(" Where CloseDate is null And LocationID = " & mSLocationID)
End If
CmbSCMPName.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_CMP.RecordCount
    CmbSCMPName.AddItem RsMst_CMP!CMPName
    RsMst_CMP.MoveNext
Next
CmbSCMPName.Text = tmp

End Sub

Private Sub CmbSCMPName_LostFocus()
If CmbSCMPName.Text = "" Then
   CmbSGodownNo.Text = ""
   mSCMPID = 0
   Exit Sub
End If

RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPName = '" & CmbSCMPName.Text & "'"
If RsMst_CMP.EOF Then
   MsgBox "Invalid selection "
   CmbSCMPName.SetFocus
   Exit Sub
End If
If mSCMPID <> RsMst_CMP!CMPID Then
   CmbSGodownNo.Text = ""
End If
mSCMPID = RsMst_CMP!CMPID
'TxtPreFix = GetPreFix(RsMst_CMP!CMPID, "P")
End Sub

Private Sub CmbSGodownNo_GotFocus()
If CmbSCMPName.Text = "" Then
   MsgBox "Select CMP !!!! "
   CmbSCMPName.SetFocus
   Exit Sub
End If
tmp = CmbSGodownNo.Text
Call TableMst_Godown("Where CloseDate is null And CMPID = " & mSCMPID)
CmbSGodownNo.Clear
If RsMst_Godown.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Godown.RecordCount
    CmbSGodownNo.AddItem RsMst_Godown!GodownNo
    RsMst_Godown.MoveNext
Next
CmbSGodownNo.Text = tmp
End Sub

Private Sub CmbSGodownNo_LostFocus()
If CmbSGodownNo.Text = "" Then
   mSGodownID = Null
   Exit Sub
End If
RsMst_Godown.MoveFirst
RsMst_Godown.Find "GodownNo = '" & CmbSGodownNo.Text & "'"
If RsMst_Godown.EOF Then
   MsgBox "Invalid selection "
   If CmbSGodownNo.Enabled = True Then
      CmbSGodownNo.SetFocus
   End If
   Exit Sub
End If
mSGodownID = RsMst_Godown!GodownID
End Sub

Private Function CheckSprayingDate() As Boolean
Dim Retval As Boolean
Retval = True
tmp = "Select max(SprayingDate)SprayingDate from Txn_Spraying Where GodownID=" & mGodownID & " And SprayingID <> " & Val(TxtSprayingID)
Call TmpTable
If TmpRs.RecordCount > 0 Then
   If IsNull(SprayingDate) = False Then
      If TmpRs!SprayingDate >= TxtSprayingDate.Value Then
          Retval = False
      End If
   End If
End If
CheckSprayingDate = Retval
End Function

Private Sub Form_Activate()
If Me.WindowState = vbMaximized Then
   SetScrollBars
End If
End Sub

Private Sub TxtSprayingDate_LostFocus()

If Format(Gen_BlockingDate, "yyyy-mm") >= Format(TxtSprayingDate.Value, "yyyy-mm") Then
   MsgBox "Spraying Date allowed for next month of " & Format(Gen_BlockingDate, "MMM-yyyy") & "!!!"
   TxtSprayingDate.SetFocus
   Exit Sub
End If

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
