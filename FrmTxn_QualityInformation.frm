VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmTxn_QualityInformation 
   Caption         =   " Quality Information"
   ClientHeight    =   6225
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8760
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   6225
   ScaleWidth      =   8760
   WindowState     =   2  'Maximized
   Begin VB.VScrollBar VScroll1 
      Height          =   9015
      Left            =   15000
      TabIndex        =   93
      Top             =   600
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      TabIndex        =   92
      Top             =   9600
      Width           =   10935
   End
   Begin VB.Frame Frame4 
      Height          =   9450
      Left            =   120
      TabIndex        =   0
      Top             =   30
      Width           =   14940
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   75
         TabIndex        =   96
         Top             =   3630
         Width           =   5895
         Begin VB.OptionButton OptLotWise 
            Caption         =   "Lot Wise"
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
            Left            =   2130
            TabIndex        =   98
            Top             =   135
            Value           =   -1  'True
            Width           =   1935
         End
         Begin VB.OptionButton OptStackwise 
            Caption         =   "Stack Wise"
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
            TabIndex        =   97
            Top             =   135
            Width           =   1935
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
         Height          =   1005
         Left            =   8400
         TabIndex        =   11
         Top             =   120
         Visible         =   0   'False
         Width           =   5235
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
            Left            =   9645
            TabIndex        =   20
            Top             =   1425
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
            Left            =   10665
            TabIndex        =   19
            Top             =   1425
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
            ItemData        =   "FrmTxn_QualityInformation.frx":0000
            Left            =   6780
            List            =   "FrmTxn_QualityInformation.frx":000A
            Sorted          =   -1  'True
            TabIndex        =   18
            Top             =   555
            Width           =   3000
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
            Left            =   12825
            Sorted          =   -1  'True
            TabIndex        =   17
            Top             =   555
            Width           =   2085
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
            Left            =   9780
            Sorted          =   -1  'True
            TabIndex        =   16
            Top             =   555
            Width           =   3045
         End
         Begin VB.ComboBox CmbSTestingAgency 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   3045
            Sorted          =   -1  'True
            TabIndex        =   15
            Top             =   1440
            Width           =   4485
         End
         Begin VB.ComboBox CmbSEmployeeID 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   3045
            Sorted          =   -1  'True
            TabIndex        =   14
            Top             =   555
            Width           =   3735
         End
         Begin VB.ComboBox CmbSCommodityID 
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
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   13
            Top             =   1440
            Width           =   2910
         End
         Begin VB.TextBox TxtSQualityInfoID 
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
            Left            =   7560
            MaxLength       =   12
            TabIndex        =   12
            Top             =   1425
            Width           =   2010
         End
         Begin MSComCtl2.DTPicker STo 
            Height          =   375
            Left            =   1590
            TabIndex        =   21
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
            Format          =   109379585
            CurrentDate     =   36494
         End
         Begin MSComCtl2.DTPicker SFrom 
            Height          =   375
            Left            =   135
            TabIndex        =   22
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
            Format          =   109379585
            CurrentDate     =   36494
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   1455
            Left            =   120
            TabIndex        =   23
            Top             =   1920
            Width           =   14655
            _ExtentX        =   25850
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
            TabIndex        =   34
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
            Left            =   375
            TabIndex        =   33
            Top             =   255
            Width           =   975
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
            Left            =   11820
            TabIndex        =   32
            Top             =   1440
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
            Left            =   14295
            TabIndex        =   31
            Top             =   1455
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
            Left            =   7898
            TabIndex        =   30
            Top             =   255
            Width           =   765
         End
         Begin VB.Label Label29 
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
            Height          =   300
            Left            =   13447
            TabIndex        =   29
            Top             =   255
            Width           =   840
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
            Left            =   11085
            TabIndex        =   28
            Top             =   255
            Width           =   435
         End
         Begin VB.Label Label18 
            Caption         =   "Testing Agency Name"
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
            Left            =   4215
            TabIndex        =   27
            Top             =   1110
            Width           =   2040
         End
         Begin VB.Label Label13 
            Caption         =   "Employee"
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
            Left            =   4380
            TabIndex        =   26
            Top             =   255
            Width           =   1065
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
            Height          =   255
            Left            =   968
            TabIndex        =   25
            Top             =   1110
            Width           =   1215
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Quality Information ID"
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
            Left            =   7628
            TabIndex        =   24
            Top             =   1080
            Width           =   1875
         End
      End
      Begin VB.Frame ButtonFrm 
         Height          =   1125
         Left            =   75
         TabIndex        =   82
         Top             =   8265
         Width           =   14805
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
            Left            =   11640
            TabIndex        =   10
            Top             =   600
            Width           =   2700
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
            Left            =   11640
            TabIndex        =   90
            Top             =   195
            Width           =   2700
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
            Left            =   8820
            TabIndex        =   89
            Top             =   600
            Width           =   2820
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
            Left            =   8820
            TabIndex        =   88
            Top             =   195
            Width           =   2820
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
            Left            =   6000
            TabIndex        =   87
            Top             =   600
            Width           =   2820
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
            Left            =   6000
            TabIndex        =   86
            Top             =   195
            Width           =   2820
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
            Left            =   3060
            TabIndex        =   85
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
            Left            =   3060
            TabIndex        =   9
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
            Left            =   120
            TabIndex        =   84
            Top             =   600
            Width           =   2940
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
            Left            =   120
            TabIndex        =   83
            Top             =   195
            Width           =   2940
         End
      End
      Begin VB.Frame FrameDetail 
         Caption         =   "Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1080
         Left            =   75
         TabIndex        =   64
         Top             =   4095
         Width           =   14775
         Begin VB.ComboBox CmbQualityParameter 
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
            ItemData        =   "FrmTxn_QualityInformation.frx":0024
            Left            =   5400
            List            =   "FrmTxn_QualityInformation.frx":0026
            Sorted          =   -1  'True
            TabIndex        =   73
            Top             =   630
            Width           =   4155
         End
         Begin VB.TextBox TxtTestValue 
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
            Left            =   11730
            MaxLength       =   100
            TabIndex        =   72
            Top             =   630
            Width           =   1620
         End
         Begin VB.CommandButton CmdUpdateList 
            Caption         =   "Update List"
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
            Left            =   13380
            TabIndex        =   71
            Top             =   630
            Width           =   1335
         End
         Begin VB.TextBox TxtExchangeType 
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
            Left            =   2040
            Locked          =   -1  'True
            MaxLength       =   12
            TabIndex        =   70
            TabStop         =   0   'False
            Top             =   630
            Width           =   2130
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
            Left            =   4185
            Locked          =   -1  'True
            MaxLength       =   12
            TabIndex        =   69
            TabStop         =   0   'False
            Top             =   630
            Width           =   1170
         End
         Begin VB.TextBox TxtLotNo 
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
            Left            =   1095
            Locked          =   -1  'True
            MaxLength       =   12
            TabIndex        =   68
            TabStop         =   0   'False
            Top             =   630
            Width           =   930
         End
         Begin VB.TextBox TxtStackNo 
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
            Left            =   150
            Locked          =   -1  'True
            MaxLength       =   12
            TabIndex        =   67
            TabStop         =   0   'False
            Top             =   630
            Width           =   930
         End
         Begin VB.TextBox TxtMaxValue 
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
            Left            =   10665
            Locked          =   -1  'True
            MaxLength       =   12
            TabIndex        =   66
            TabStop         =   0   'False
            Top             =   630
            Width           =   1050
         End
         Begin VB.TextBox TxtMinValue 
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
            Left            =   9600
            Locked          =   -1  'True
            MaxLength       =   12
            TabIndex        =   65
            TabStop         =   0   'False
            Top             =   630
            Width           =   1050
         End
         Begin VB.Label Label16 
            Caption         =   "Test Value"
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
            Left            =   12060
            TabIndex        =   81
            Top             =   285
            Width           =   1005
         End
         Begin VB.Label Label15 
            Caption         =   "Quality Parameter"
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
            Left            =   6570
            TabIndex        =   80
            Top             =   285
            Width           =   1695
         End
         Begin VB.Label LblExchange 
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
            Height          =   255
            Left            =   2385
            TabIndex        =   79
            Top             =   285
            Width           =   1455
         End
         Begin VB.Label Label11 
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
            Height          =   255
            Left            =   4245
            TabIndex        =   78
            Top             =   285
            Width           =   1050
         End
         Begin VB.Label LbLLotNo 
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
            Left            =   1215
            TabIndex        =   77
            Top             =   285
            Width           =   690
         End
         Begin VB.Label Label6 
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
            Left            =   180
            TabIndex        =   76
            Top             =   285
            Width           =   885
         End
         Begin VB.Label Label7 
            Caption         =   "Max Value"
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
            Left            =   10695
            TabIndex        =   75
            Top             =   285
            Width           =   1005
         End
         Begin VB.Label Label5 
            Caption         =   "Min Value"
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
            Left            =   9630
            TabIndex        =   74
            Top             =   285
            Width           =   885
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
         Height          =   3525
         Left            =   75
         TabIndex        =   35
         Top             =   90
         Width           =   14775
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
            Height          =   570
            Left            =   1800
            Locked          =   -1  'True
            MaxLength       =   2000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   2475
            Width           =   12480
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
            Left            =   14280
            TabIndex        =   45
            TabStop         =   0   'False
            Top             =   735
            Width           =   375
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
            Height          =   810
            Left            =   9240
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   1590
            Width           =   5040
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
            Left            =   1800
            Sorted          =   -1  'True
            TabIndex        =   4
            Top             =   1149
            Width           =   5400
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
            Left            =   9240
            Sorted          =   -1  'True
            TabIndex        =   5
            Top             =   1155
            Width           =   5055
         End
         Begin VB.TextBox TxtQualityInformationID 
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
            Left            =   2295
            Locked          =   -1  'True
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   240
            Width           =   1635
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
            Left            =   9240
            Sorted          =   -1  'True
            TabIndex        =   3
            Top             =   720
            Width           =   5040
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
            Left            =   9240
            Locked          =   -1  'True
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   3105
            Width           =   5040
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
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   3105
            Width           =   5400
         End
         Begin VB.ComboBox CmbEmployeeID 
            BeginProperty Font 
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
            Top             =   720
            Width           =   5400
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
            Left            =   1800
            Sorted          =   -1  'True
            TabIndex        =   6
            Top             =   1605
            Width           =   5400
         End
         Begin VB.CommandButton CmdTestingAgency 
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
            Left            =   7290
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   2055
            Width           =   375
         End
         Begin VB.ComboBox CmbTestingAgency 
            BeginProperty Font 
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
            TabIndex        =   7
            Top             =   2040
            Width           =   5400
         End
         Begin VB.CommandButton CmdEmployeeID 
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
            Left            =   7290
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   750
            Width           =   375
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
            Height          =   300
            Left            =   7290
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   1179
            Width           =   375
         End
         Begin VB.CommandButton CmdCommodity 
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
            Left            =   7290
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   1635
            Width           =   375
         End
         Begin VB.TextBox TxtDesignationID 
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
            Left            =   9240
            Locked          =   -1  'True
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   240
            Width           =   5040
         End
         Begin MSComCtl2.DTPicker TxtDate 
            Height          =   375
            Left            =   5640
            TabIndex        =   1
            Top             =   240
            Width           =   1560
            _ExtentX        =   2752
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
            Format          =   109379585
            CurrentDate     =   39884
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
            Left            =   9240
            Locked          =   -1  'True
            TabIndex        =   49
            Top             =   720
            Visible         =   0   'False
            Width           =   5040
         End
         Begin VB.TextBox txtGodownNo 
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
            Left            =   9240
            Locked          =   -1  'True
            TabIndex        =   48
            Top             =   1155
            Visible         =   0   'False
            Width           =   5040
         End
         Begin VB.TextBox TxtEmployeeID 
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
            TabIndex        =   46
            Top             =   720
            Visible         =   0   'False
            Width           =   5400
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
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   47
            Top             =   1155
            Visible         =   0   'False
            Width           =   5400
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
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   1605
            Visible         =   0   'False
            Width           =   5400
         End
         Begin VB.TextBox TxtTestingAgency 
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
            TabIndex        =   51
            Top             =   2040
            Visible         =   0   'False
            Width           =   5400
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
            Height          =   240
            Left            =   120
            TabIndex        =   94
            Top             =   2640
            Width           =   1335
         End
         Begin VB.Label Label19 
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
            Height          =   495
            Left            =   8040
            TabIndex        =   63
            Top             =   1748
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Employee"
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
            TabIndex        =   62
            Top             =   765
            Width           =   1185
         End
         Begin VB.Label Label26 
            Caption         =   "CMP Name"
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
            TabIndex        =   61
            Top             =   1230
            Width           =   1200
         End
         Begin VB.Label Label21 
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
            Height          =   255
            Left            =   8040
            TabIndex        =   60
            Top             =   1230
            Width           =   900
         End
         Begin VB.Label Label8 
            Caption         =   "Date"
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
            Left            =   4920
            TabIndex        =   59
            Top             =   270
            Width           =   720
         End
         Begin VB.Label Label4 
            Caption         =   "Quality Information ID"
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
            TabIndex        =   58
            Top             =   270
            Width           =   2160
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
            Left            =   8040
            TabIndex        =   57
            Top             =   765
            Width           =   1080
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
            Left            =   8040
            TabIndex        =   56
            Top             =   3165
            Width           =   975
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
            Left            =   120
            TabIndex        =   55
            Top             =   3165
            Width           =   1695
         End
         Begin VB.Label Label2 
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
            Height          =   255
            Left            =   120
            TabIndex        =   54
            Top             =   1665
            Width           =   1215
         End
         Begin VB.Label Label3 
            Caption         =   "Testing Agency"
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
            TabIndex        =   53
            Top             =   2070
            Width           =   1680
         End
         Begin VB.Label Label9 
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
            Height          =   255
            Left            =   8040
            TabIndex        =   52
            Top             =   293
            Width           =   1695
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
         Height          =   2985
         Left            =   75
         TabIndex        =   91
         Top             =   5265
         Width           =   14790
         _ExtentX        =   26088
         _ExtentY        =   5265
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
         Height          =   2985
         Left            =   75
         TabIndex        =   95
         Top             =   5265
         Width           =   14790
         _ExtentX        =   26088
         _ExtentY        =   5265
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
End
Attribute VB_Name = "FrmTxn_QualityInformation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg As Variant
Dim mEmployeeID, mLocationID, mCMPID, mGodownID, mAgencyID As Variant
Dim mDesignationID, mCommodityID, mSM_Prefix As Variant
Dim mSEmployeeID, mSLocationID, mSCMPID, mSGodownID, mSCommodityID, mSAgencyID As Variant

Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)
TxtQualityInformationID.Text = ""
TxtDate.Enabled = True
TxtDate.Value = Date


CmbEmployeeID.Visible = True
CmbLocationID.Visible = True
CmbCMPName.Visible = True
CmbGodownNo.Visible = True
CmbCommodityID.Visible = True
CmbTestingAgency.Visible = True


CmdEmployeeID.Visible = True
CmdLocation.Visible = True
CmdCMP.Visible = True
CmdCommodity.Visible = True
CmdTestingAgency.Visible = True


TxtEmployeeID.Visible = False
TxtLocationID.Visible = False
TxtCMPName.Visible = False
TxtGodownNo.Visible = False
TxtCommodityID.Visible = False
TxtTestingAgency.Visible = False
TxtRemarks.Locked = False


CmbEmployeeID.Text = ""
CmbLocationID.Text = ""
CmbCMPName.Text = ""
CmbGodownNo.Text = ""
CmbCommodityID.Text = ""
CmbTestingAgency.Text = ""

TxtAddress.Text = ""
TxtRemarks = ""
TxtCreated = ""
TxtUpdated = ""

Call CmbEmployeeID_GotFocus
CmbEmployeeID.Text = GetEmployeeNonName(Gen_EmployeeID)
Call CmbEmployeeID_LostFocus

Frame2.Enabled = True
OptLotWise.Value = True
OptLotWise_Click


Call ClearFrame
Call Grid_HeadDetail
Call Grid_HeadDetailStack

End Sub

Private Sub CmbEmployeeID_GotFocus()
tmp = CmbEmployeeID.Text
Call TableMst_Employee("Where Flag = 'Y'")
CmbEmployeeID.Clear
If RsMst_Employee.RecordCount = 0 Then
   MsgBox "No Employee Found!!!"
   CmbEmployeeID.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Employee.RecordCount
    CmbEmployeeID.AddItem RsMst_Employee!EmployeeName & "~" & RsMst_Employee!EmployeeNo
    RsMst_Employee.MoveNext
Next
CmbEmployeeID.Text = tmp
End Sub

Private Sub CmbEmployeeID_LostFocus()
Dim strSplitString() As String
Dim EmpNo_, EmpName_ As String

If CmbEmployeeID.Text = "" Then
   mEmployeeID = Null
   Exit Sub
End If

strSplitString = Split(CmbEmployeeID, "~")
If UBound(strSplitString) <> 1 Then
   MsgBox "Invalid Selection."
   CmbEmployeeID.SetFocus
   Exit Sub
End If

EmpName_ = Trim(strSplitString(0))
EmpNo_ = Trim(strSplitString(1))

RsMst_Employee.MoveFirst
RsMst_Employee.Find "EmployeeNo = '" & EmpNo_ & "'"

If RsMst_Employee.EOF Then
   MsgBox "Invalid Selection "
   CmbEmployeeID.SetFocus
   Exit Sub
End If
mEmployeeID = RsMst_Employee!EmployeeID
TxtDesignationID.Text = GetDesignaion(RsMst_Employee!DesignationID)
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
   TxtAddress.Text = ""
   mLocationID = 0
   Call Grid_HeadDetail
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
   CmbCommodityID.Text = ""
   Call Grid_HeadDetail
   Call ClearFrame
End If

mLocationID = RsMst_Location!LocationID
mSM_Prefix = RsMst_Location!SM_Prefix
Call FillQualityParameter
Call FillStackLotInfo
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
   TxtAddress.Text = ""
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
   CmbCommodityID.Text = ""
   Call Grid_HeadDetail
End If
mCMPID = RsMst_CMP!CMPID
Call FillQualityParameter
Call FillStackLotInfo
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
   CmbCommodityID.Text = ""
   Call Grid_HeadDetail
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
   CmbCommodityID.Text = ""
   Call Grid_HeadDetail
End If

mGodownID = RsMst_Godown!GodownID
TxtAddress.Text = RsMst_Godown!Address

Call FillQualityParameter
Call FillStackLotInfo

Call FillQualityParameterStack
Call FillStackInfo

End Sub

Private Sub CmbCommodityID_GotFocus()
If CmbLocationID.Text = "" Then
   MsgBox "Select Location."
   CmbLocationID.SetFocus
   Exit Sub
End If
If CmbCMPName.Text = "" Then
   MsgBox "Select CMP."
   CmbCMPName.SetFocus
   Exit Sub
End If
If CmbGodownNo.Text = "" Then
   MsgBox "Select Godown."
   CmbGodownNo.SetFocus
   Exit Sub
End If

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
   Call Grid_HeadDetail
End If
mCommodityID = RsMst_Commodity!CommodityID
Call FillQualityParameter
Call FillStackLotInfo

Call FillQualityParameterStack
Call FillStackInfo

End Sub
Private Sub FillStackLotInfo()
If Edit_Flg = "A" Then
   If CmbLocationID.Text = "" Or CmbCMPName.Text = "" Or CmbGodownNo.Text = "" Or CmbCommodityID.Text = "" Then Exit Sub
End If
If MSHFlexGrid2.Cols <= 5 Then
   MsgBox "No Quality Parameter added in 'Quality Parameter Master'!!!"
   Exit Sub
End If
'Call Grid_HeadDetail
Call TableMst_GSL(" Where Flag not in ('E','Z') And BalanceBags > 0 And SM_PreFix='" & mSM_Prefix & "' And CMPID=" & mCMPID & " And CommodityID=" & mCommodityID & " And GodownID=" & mGodownID & " Order By StackNo,LotNo")
For I = 2 To RsMst_GSL.RecordCount + 1
     MSHFlexGrid2.TextMatrix(I, 0) = RsMst_GSL!StackNo
     MSHFlexGrid2.TextMatrix(I, 1) = RsMst_GSL!LotNo
     MSHFlexGrid2.TextMatrix(I, 2) = GetExchangeName(IIf(IsNull(RsMst_GSL!ExchangeTypeID), 0, RsMst_GSL!ExchangeTypeID))
     MSHFlexGrid2.TextMatrix(I, 3) = RsMst_GSL!BalanceBags 'RsMst_GSL!DepositeBags
     MSHFlexGrid2.TextMatrix(I, 4) = RsMst_GSL!GSLID
     MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
     RsMst_GSL.MoveNext
Next
End Sub

Private Sub FillStackInfo()
If Edit_Flg = "A" Then
   If CmbLocationID.Text = "" Or CmbCMPName.Text = "" Or CmbGodownNo.Text = "" Or CmbCommodityID.Text = "" Then Exit Sub
End If
If MSHFlexGrid3.Cols <= 5 Then
   MsgBox "No Quality Parameter added in 'Quality Parameter Master'!!!"
   Exit Sub
End If
'Call Grid_HeadDetail
'Call TableMst_GSL(" Where Flag not in ('E','Z') And BalanceBags > 0 And SM_PreFix='" & mSM_Prefix & "' And CMPID=" & mCMPID & " And CommodityID=" & mCommodityID & " And GodownID=" & mGodownID & " Order By StackNo,LotNo")
tmp = "Select StackNo,'','',sum(BalanceBags) 'BalanceBags','' from Mst_GSL  Where Flag not in ('E','Z') And BalanceBags > 0 And SM_PreFix='" & mSM_Prefix & "' And CMPID=" & mCMPID & " And CommodityID=" & mCommodityID & " And GodownID=" & mGodownID & " Group By StackNo Order By StackNo "
Call TmpTable
For I = 2 To TmpRs.RecordCount + 1
     MSHFlexGrid3.TextMatrix(I, 0) = TmpRs!StackNo
     MSHFlexGrid3.TextMatrix(I, 1) = "" 'RsMst_GSL!LotNo
     MSHFlexGrid3.TextMatrix(I, 2) = "" 'GetExchangeName(IIf(IsNull(RsMst_GSL!ExchangeTypeID), 0, RsMst_GSL!ExchangeTypeID))
     MSHFlexGrid3.TextMatrix(I, 3) = TmpRs!BalanceBags 'RsMst_GSL!DepositeBags
     MSHFlexGrid3.TextMatrix(I, 4) = "" 'RsMst_GSL!gslid
     MSHFlexGrid3.Rows = MSHFlexGrid3.Rows + 1
     TmpRs.MoveNext
Next
End Sub

Private Sub FillQualityParameter()

If Edit_Flg = "A" Then
   If CmbCommodityID.Text = "" Then Exit Sub
End If

tmp = "Select B.QualityParameterID,B.QualityParameter From Mst_CommodityQualityParameter A"
tmp = tmp & " Inner Join Mst_QualityParameter B On A.QualityParameterId=B.QualityParameterID "
tmp = tmp & " Where A.CommodityID=" & mCommodityID & ""
Call TmpTable

Call Grid_HeadDetail

CmbQualityParameter.Clear
If TmpRs.RecordCount > 0 Then
   MSHFlexGrid2.Cols = MSHFlexGrid2.Cols + TmpRs.RecordCount
   For I = 2 To TmpRs.RecordCount + 1
       MSHFlexGrid2.TextMatrix(0, 3 + I) = TmpRs!QualityParameterID
       MSHFlexGrid2.TextMatrix(1, 3 + I) = TmpRs!QualityParameter
       CmbQualityParameter.AddItem TmpRs!QualityParameter, I - 2
       MSHFlexGrid2.ColWidth(3 + I) = Len(TmpRs!QualityParameter) * 120
       TmpRs.MoveNext
   Next
Else

   Exit Sub
End If

End Sub
Private Sub FillQualityParameterStack()

If Edit_Flg = "A" Then
   If CmbCommodityID.Text = "" Then Exit Sub
End If

tmp = "Select B.QualityParameterID,B.QualityParameter From Mst_CommodityQualityParameter A"
tmp = tmp & " Inner Join Mst_QualityParameter B On A.QualityParameterId=B.QualityParameterID "
tmp = tmp & " Where A.CommodityID=" & mCommodityID & ""
Call TmpTable

Call Grid_HeadDetailStack

CmbQualityParameter.Clear
If TmpRs.RecordCount > 0 Then
   MSHFlexGrid3.Cols = MSHFlexGrid3.Cols + TmpRs.RecordCount
   For I = 2 To TmpRs.RecordCount + 1
       MSHFlexGrid3.TextMatrix(0, 3 + I) = TmpRs!QualityParameterID
       MSHFlexGrid3.TextMatrix(1, 3 + I) = TmpRs!QualityParameter
       CmbQualityParameter.AddItem TmpRs!QualityParameter, I - 2
       MSHFlexGrid3.ColWidth(3 + I) = Len(TmpRs!QualityParameter) * 120
       TmpRs.MoveNext
   Next
Else
   Exit Sub
End If

End Sub

Private Sub Grid_HeadDetail()
With MSHFlexGrid2
     .Clear
     .Rows = 3
     .Cols = 5
     .FixedRows = 2
     .AllowUserResizing = flexResizeBoth
     .TextMatrix(1, 0) = "Stack No"
     .TextMatrix(1, 1) = "Lot No"
     .TextMatrix(1, 2) = "Exchange Type"
     .TextMatrix(1, 3) = "No of Bags"
     .TextMatrix(1, 4) = "GSL ID"
     .ColWidth(0) = 1000
     .ColWidth(1) = 1000
     .ColWidth(2) = 1600
     .ColWidth(3) = 1500
     .ColWidth(4) = 0
     .RowHeight(0) = 0
     .RowHeight(1) = 300
End With
End Sub

Private Sub Grid_HeadDetailStack()
With MSHFlexGrid3
     .Clear
     .Rows = 3
     .Cols = 5
     .FixedRows = 2
     .AllowUserResizing = flexResizeBoth
     .TextMatrix(1, 0) = "Stack No"
     .TextMatrix(1, 1) = "Lot No"
     .TextMatrix(1, 2) = "Exchange Type"
     .TextMatrix(1, 3) = "No of Bags"
     .TextMatrix(1, 4) = "GSL ID"
     .ColWidth(0) = 1000
     .ColWidth(1) = 0
     .ColWidth(2) = 0
     .ColWidth(3) = 1500
     .ColWidth(4) = 0
     .RowHeight(0) = 0
     .RowHeight(1) = 300
End With
End Sub

Private Sub CmbQualityParameter_LostFocus()
If CmbQualityParameter.Text = "" Then Exit Sub
Call TableMst_QualityParameter(" Where QualityParameter='" & CmbQualityParameter.Text & "'")
If RsMst_QualityParameter.RecordCount = 0 Then
   MsgBox "Invalid Selection."
   CmbQualityParameter.SetFocus
   Exit Sub
End If
Call TableMst_CommodityQualityParameter(" Where CommodityID=" & mCommodityID & " And QualityParameterID=" & RsMst_QualityParameter!QualityParameterID & "")
TxtMinValue.Text = RsMst_CommodityQualityParameter!MinValue
TxtMaxValue.Text = RsMst_CommodityQualityParameter!MaxValue
End Sub

Private Sub CmbTestingAgency_GotFocus()
tmp = CmbTestingAgency.Text
CmbTestingAgency.Clear
Call TableMst_TestingAgency(" Where Status='A'")
If RsMst_TestingAgency.RecordCount = 0 Then
   MsgBox "Testing Agency Not Found"
   Exit Sub
End If
For I = 1 To RsMst_TestingAgency.RecordCount
    CmbTestingAgency.AddItem RsMst_TestingAgency!AgencyName
    RsMst_TestingAgency.MoveNext
Next
CmbTestingAgency.Text = tmp
End Sub

Private Sub CmbTestingAgency_LostFocus()
If CmbTestingAgency.Text = "" Then Exit Sub
RsMst_TestingAgency.MoveFirst
RsMst_TestingAgency.Find "AgencyName='" & CmbTestingAgency.Text & "'"
If RsMst_TestingAgency.EOF Then
   MsgBox "Invalid Selection."
   Exit Sub
End If
mAgencyID = RsMst_TestingAgency!TestingAgencyID
End Sub

Private Sub CmdCMP_Click()
FrmMst_CMP.Show
End Sub

Private Sub CmdCommodity_Click()
FrmMst_Commodity.Show
End Sub


Private Sub CmdEmployeeID_Click()
FrmMst_Employee.Show
End Sub

Private Sub CmdLocation_Click()
FrmMst_Location.Show
End Sub

Private Sub CmdSearch_Click()
Dim wc As Variant
Call Grid_Head

tmp = "Select TQ.QualityInformationID,TQ.[Date],MD.Designation,ME.EmployeeName + '~' +ME.EmployeeNo As EmployeeName,"
tmp = tmp & " ML.LocationName , MC.CMPName, MG.GodownNo, MG.Address, MCom.Commodity, TA.AgencyName,TQ.Remarks"
tmp = tmp & " From Txn_QualityInformation TQ"
tmp = tmp & " Inner Join Mst_Employee ME On ME.EmployeeID=TQ.EmployeeID"
tmp = tmp & " Inner Join Mst_Designation MD On MD.DesignationID=ME.DesignationID"
tmp = tmp & " Inner Join Mst_Godown MG On MG.GodownID=TQ.GodownID"
tmp = tmp & " Inner Join Mst_CMP MC On MC.CMPID=MG.CMPID"
tmp = tmp & " Inner Join Mst_Location ML On ML.LocationID=MC.LocationID"
tmp = tmp & " Inner Join Mst_TestingAgency TA On TA.TestingAgencyID=TQ.TestingAgencyID"
tmp = tmp & " Inner Join Mst_Commodity MCom On MCom.CommodityID=TQ.CommodityID"
tmp = tmp & " Where "
If Gen_DBType = "MDB" Then
   tmp = tmp & " TQ.[Date] >= #" & SFrom.Value & "# And TQ.[Date]<=#" & STo.Value & "#"
Else
   tmp = tmp & " TQ.[Date] >= '" & SFrom.Value & "' And TQ.[Date]<='" & STo.Value & "'"
End If

If CmbSEmployeeID.Text <> "" Then
   tmp = tmp & "  And TQ.EmployeeID=" & mSEmployeeID & ""
End If
If CmbSLocationID.Text <> "" Then
   tmp = tmp & " And ML.LocationID=" & mSLocationID & ""
End If
If CmbSCMPName.Text <> "" Then
   tmp = tmp & " And MC.CMPID=" & mSCMPID & ""
End If
If CmbSGodownNo.Text <> "" Then
   tmp = tmp & " And TQ.GodownID=" & mSGodownID & ""
End If
If CmbSCommodityID.Text <> "" Then
   tmp = tmp & " And TQ.CommodityID=" & mSCommodityID & ""
End If
If CmbSTestingAgency.Text <> "" Then
   tmp = tmp & " And TQ.TestingAgencyID=" & mSAgencyID & ""
End If
If TxtSQualityInfoID.Text <> "" Then
   tmp = tmp & " And TQ.QualityInformationID=" & Val(TxtSQualityInfoID.Text) & " "
End If

tmp = tmp & " Order by TQ.QualityInformationID"

Call TmpTable
For I = 1 To TmpRs.RecordCount
    For C = 0 To TmpRs.Fields.Count - 1
        MSHFlexGrid1.TextMatrix(I, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
    Next
    MSHFlexGrid1.TextMatrix(I, 1) = Format(MSHFlexGrid1.TextMatrix(I, 1), "dd-mmm-yyyy")
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
    TmpRs.MoveNext
Next
Label22.Caption = TmpRs.RecordCount
End Sub

Private Sub Grid_Head()
With MSHFlexGrid1
     .Clear
     .Rows = 2
     .Cols = 11
     .WordWrap = True
     .AllowUserResizing = flexResizeBoth
     .TextMatrix(0, 0) = "Quality Information ID"
     .TextMatrix(0, 1) = "Date"
     .TextMatrix(0, 2) = "Designation"
     .TextMatrix(0, 3) = "Employee"
     .TextMatrix(0, 4) = "Location"
     .TextMatrix(0, 5) = "CMP"
     .TextMatrix(0, 6) = "Godown No"
     .TextMatrix(0, 7) = "Address"
     .TextMatrix(0, 8) = "Commodity"
     .TextMatrix(0, 9) = "Testing Agency"
     .TextMatrix(0, 10) = "Remarks"
     
     .ColWidth(0) = 1100
     .ColWidth(1) = 1400
     .ColWidth(2) = 1400
     .ColWidth(3) = 2000
     .ColWidth(4) = 1500
     .ColWidth(5) = 1500
     .ColWidth(6) = 1200
     .ColWidth(7) = 2200
     .ColWidth(8) = 1800
     .ColWidth(9) = 2000
     .ColWidth(10) = 2000
     .RowHeight(0) = 700
End With
End Sub

Private Sub CmdTestingAgency_Click()
FrmMst_TestingAgency.Show
End Sub

Private Sub CmdUpdateList_Click()
Dim x, SelIndex As Variant
If OptLotWise.Value = True Then
   x = MSHFlexGrid2.RowSel
   SelIndex = CmbQualityParameter.ListIndex
   MSHFlexGrid2.TextMatrix(x, SelIndex + 5) = TxtTestValue.Text
   CmbQualityParameter.Text = ""
   TxtMinValue.Text = ""
   TxtMaxValue.Text = ""
   TxtTestValue.Text = ""

   'Call ClearFrame
   'FrameDetail.Enabled = False
ElseIf OptStackwise.Value = True Then
   x = MSHFlexGrid3.RowSel
   SelIndex = CmbQualityParameter.ListIndex
   MSHFlexGrid3.TextMatrix(x, SelIndex + 5) = TxtTestValue.Text
   Call UpdateLotwiseData(MSHFlexGrid3.TextMatrix(x, 0), SelIndex)
   CmbQualityParameter.Text = ""
   TxtMinValue.Text = ""
   TxtMaxValue.Text = ""
   TxtTestValue.Text = ""
   
End If

End Sub
Private Sub UpdateLotwiseData(StackNo_ As Variant, SelIndx_ As Variant)

For y = 2 To MSHFlexGrid2.Rows - 1
    If MSHFlexGrid2.TextMatrix(y, 0) = StackNo_ Then
       MSHFlexGrid2.TextMatrix(y, SelIndx_ + 5) = TxtTestValue.Text
    End If
Next

End Sub

Private Sub ClearFrame()
TxtStackNo.Text = ""
TxtLotNo.Text = ""
TxtNoofBags.Text = ""
CmbQualityParameter.Text = ""
TxtMinValue.Text = ""
TxtMaxValue.Text = ""
TxtTestValue.Text = ""
TxtExchangeType.Text = ""
End Sub

Private Sub DeleteCmd_Click()
Dim ans As Variant
On Error GoTo msgError
ans = MsgBox("Are you Sure you want to Delete this Record ", vbYesNo)
If ans = vbYes Then
   tmp = "Delete Txn_QualityInformationDetail Where QualityInformationID=" & Val(TxtQualityInformationID.Text) & ""
   Call TmpTable
   RsTxn_QualityInformation.Delete
   RsTxn_QualityInformation.Update
   If RsTxn_QualityInformation.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsTxn_QualityInformation.MoveNext
   If RsTxn_QualityInformation.EOF() Then
      RsTxn_QualityInformation.MoveLast
   End If
   Call Indata
End If

Exit Sub
msgError:
MsgBox "Child Record Present "
RsTxn_QualityInformation.CancelUpdate

End Sub

Private Sub EditCmd_Click()
Edit_Flg = "E"
Call GButtonLock(Me, False)
TxtDate.Enabled = True


CmbEmployeeID.Visible = True
CmbLocationID.Visible = True
CmbCMPName.Visible = True
CmbGodownNo.Visible = True
CmbCommodityID.Visible = True
CmbTestingAgency.Visible = True


CmdEmployeeID.Visible = True
CmdLocation.Visible = True
CmdCMP.Visible = True
CmdCommodity.Visible = True
CmdTestingAgency.Visible = True


TxtEmployeeID.Visible = False
TxtLocationID.Visible = False
TxtCMPName.Visible = False
TxtGodownNo.Visible = False
TxtCommodityID.Visible = False
TxtTestingAgency.Visible = False
TxtRemarks.Locked = False

CmbEmployeeID.Text = TxtEmployeeID.Text
CmbLocationID.Text = TxtLocationID.Text
CmbCMPName.Text = TxtCMPName.Text
CmbGodownNo.Text = TxtGodownNo.Text
CmbCommodityID.Text = TxtCommodityID.Text
CmbTestingAgency.Text = TxtTestingAgency.Text

If LCase(RsTxn_QualityInformation!DetailsEnteredFlg) = "l" Then
   OptLotWise.Value = True
   Call OptLotWise_Click
Else
   OptStackwise.Value = True
   Call OptStackwise_Click
End If

Frame2.Enabled = False

Call FillEditGSLInfo

End Sub

Private Sub FillEditGSLInfo()

tmp = "Select M.StackNo,M.LotNo,ME.ExchangeType,M.BalanceBags,M.GSLID"
tmp = tmp & " From Mst_GSL M"
tmp = tmp & " Inner Join Mst_ExchangeType ME On ME.ExchangeTypeID = M.ExchangeTypeID"
tmp = tmp & " Where M.GSLID Not In"
tmp = tmp & " ("
tmp = tmp & " Select A.GSLID"
tmp = tmp & " From Txn_QualityInformationDetail A"
tmp = tmp & " Inner Join Txn_QualityInformation B On A.QualityInformationID =B.QualityInformationID"
tmp = tmp & " Inner Join Mst_Godown MG On MG.GodownID=B.GodownID"
tmp = tmp & " Inner Join Mst_CMP MC On MC.CMPID=MG.CMPID"
tmp = tmp & " Inner Join Mst_Location ML On ML.LocationID = MC.LocationID"
tmp = tmp & " Where B.QualityInformationID=" & Val(TxtQualityInformationID.Text) & " And MC.CMPID=" & mCMPID & " And  MG.GodownID=" & mGodownID & " And ML.SM_Prefix='" & mSM_Prefix & "'"
tmp = tmp & " )"
tmp = tmp & " And M.CMPID=" & mCMPID & " And M.SM_PreFix='" & mSM_Prefix & "' And M.CommodityID=" & mCommodityID & "  And M.GodownID=" & mGodownID & ""

Call TmpTable
mRow = MSHFlexGrid2.Rows - 1
For I = mRow To TmpRs.RecordCount + mRow - 1
    For C = 0 To TmpRs.Fields.Count - 1
        MSHFlexGrid2.TextMatrix(I, C) = TmpRs.Fields(C)
    Next
    TmpRs.MoveNext
    MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
Next


tmp = "Select M.StackNo,'','',Sum(M.BalanceBags) 'BalanceBags',''"
tmp = tmp & " From Mst_GSL M"
tmp = tmp & " Where M.Flag not in ('E','Z') And M.BalanceBags > 0 And M.GSLID Not In"
tmp = tmp & " ("
tmp = tmp & " Select A.GSLID"
tmp = tmp & " From Txn_QualityInformationDetail A"
tmp = tmp & " Inner Join Txn_QualityInformation B On A.QualityInformationID =B.QualityInformationID"
tmp = tmp & " Inner Join Mst_Godown MG On MG.GodownID=B.GodownID"
tmp = tmp & " Inner Join Mst_CMP MC On MC.CMPID=MG.CMPID"
tmp = tmp & " Inner Join Mst_Location ML On ML.LocationID = MC.LocationID"
tmp = tmp & " Where B.QualityInformationID=" & Val(TxtQualityInformationID.Text) & " And MC.CMPID=" & mCMPID & " And  MG.GodownID=" & mGodownID & " And ML.SM_Prefix='" & mSM_Prefix & "'"
tmp = tmp & " )"
tmp = tmp & " And M.CMPID=" & mCMPID & " And M.SM_PreFix='" & mSM_Prefix & "' And M.CommodityID=" & mCommodityID & "  And M.GodownID=" & mGodownID & ""
tmp = tmp & " Group By M.StackNo"
tmp = tmp & " Order By M.StackNo"

Call TmpTable
mRow = MSHFlexGrid3.Rows - 1
For I = mRow To TmpRs.RecordCount + mRow - 1
    For C = 0 To TmpRs.Fields.Count - 1
        MSHFlexGrid3.TextMatrix(I, C) = TmpRs.Fields(C)
    Next
    TmpRs.MoveNext
    MSHFlexGrid3.Rows = MSHFlexGrid3.Rows + 1
Next



End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrameDetail.Enabled = False
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Call GButtonLockAD(Me)
   Frame1.Enabled = False
   MsgBox "Access Denied !!!!!!!!!"
   Exit Sub
End If

Frame1.Enabled = True
Call Grid_Head
SFrom.Value = Date - 365
STo.Value = Date
Label22.Caption = ""
CmbSEmployeeID.Text = ""
CmbSLocationID.Text = ""
CmbSCMPName.Text = ""
CmbSGodownNo.Text = ""
CmbSTestingAgency.Text = ""
TxtSQualityInfoID.Text = ""

Call Grid_HeadDetail
Call TableTxn_QualityInformation(" Order by QualityInformationID ")

If RsTxn_QualityInformation.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
Call Indata
Call GButtonLock(Me, True)
End Sub

Private Sub Indata()
Dim mTemp As Double
Dim mTemp1 As Double
TxtDate.Enabled = False

TxtQualityInformationID.Text = ""
'TxtDate.Value = Date

CmbEmployeeID.Visible = False
CmbLocationID.Visible = False
CmbCMPName.Visible = False
CmbGodownNo.Visible = False
CmbCommodityID.Visible = False
CmbTestingAgency.Visible = False


CmdEmployeeID.Visible = False
CmdLocation.Visible = False
CmdCMP.Visible = False
CmdCommodity.Visible = False
CmdTestingAgency.Visible = False


TxtEmployeeID.Visible = True
TxtLocationID.Visible = True
TxtCMPName.Visible = True
TxtGodownNo.Visible = True
TxtCommodityID.Visible = True
TxtTestingAgency.Visible = True
TxtRemarks.Locked = True

TxtDesignationID.Text = ""
TxtEmployeeID.Text = ""
TxtLocationID.Text = ""
TxtCMPName.Text = ""
TxtGodownNo.Text = ""
TxtCommodityID.Text = ""
TxtTestingAgency.Text = ""
Call ClearFrame
FrameDetail.Enabled = False

Call Grid_HeadDetail

TxtQualityInformationID.Text = RsTxn_QualityInformation!QualityInformationID
TxtDate.Value = RsTxn_QualityInformation!Date
mEmployeeID = RsTxn_QualityInformation!EmployeeID
TxtEmployeeID = GetEmployeeNonName(RsTxn_QualityInformation!EmployeeID)
TxtDesignationID = GetDesignationFromEmp(RsTxn_QualityInformation!EmployeeID)
mCMPID = GetCMPIDFromGodown(RsTxn_QualityInformation!GodownID)
mTemp = mCMPID
mLocationID = GetLocationIDFromCMP(mTemp)
Call TableMst_Location(" Where LocationID=" & mLocationID & "")
TxtLocationID = RsMst_Location!LocationName
mSM_Prefix = RsMst_Location!SM_Prefix
TxtCMPName = GetCMPName(mTemp)
mGodownID = RsTxn_QualityInformation!GodownID
TxtGodownNo = GetGodownName(RsTxn_QualityInformation!GodownID, "N")
TxtAddress = GetGodownName(RsTxn_QualityInformation!GodownID, "A")
mCommodityID = RsTxn_QualityInformation!CommodityID
TxtCommodityID = GetCommodityName(RsTxn_QualityInformation!CommodityID)
mAgencyID = RsTxn_QualityInformation!TestingAgencyID
TxtTestingAgency = GetTestingAgencyName(RsTxn_QualityInformation!TestingAgencyID)
TxtRemarks = IIf(IsNull(RsTxn_QualityInformation!Remarks), "", RsTxn_QualityInformation!Remarks)
TxtCreated = IIf(IsNull(RsTxn_QualityInformation!CreatedBy), "", RsTxn_QualityInformation!CreatedBy) & " :- " & IIf(IsNull(RsTxn_QualityInformation!CreatedDate), "", RsTxn_QualityInformation!CreatedDate)
TxtUpdated = IIf(IsNull(RsTxn_QualityInformation!UpdatedBy), "", RsTxn_QualityInformation!UpdatedBy) & " :- " & IIf(IsNull(RsTxn_QualityInformation!UpdatedDate), "", RsTxn_QualityInformation!UpdatedDate)

If LCase(RsTxn_QualityInformation!DetailsEnteredFlg) = "l" Then
   OptLotWise.Value = True
   Call OptLotWise_Click
Else
   OptStackwise.Value = True
   Call OptStackwise_Click
End If

Frame2.Enabled = True

Call InDataDetail
Call GButtonLock(Me, True)

End Sub

Private Sub InDataDetail()

Call FillQualityParameter
Call FillQualityParameterStack


tmp = "Select G.StackNo, G.LotNo, E.ExchangeType,A.BalanceBags, G.GSLID,A.QualityParameterID,A.TestValue "
tmp = tmp & " From Txn_QualityInformationDetail A"
tmp = tmp & " Inner join Mst_GSL G On G.GSLID=A.GSLID And G.CMPID=" & mCMPID & " And G.SM_Prefix='" & mSM_Prefix & "'"
tmp = tmp & " Inner Join Mst_ExchangeType E On G.ExchangeTypeID=E.ExchangeTypeID"
tmp = tmp & " Where A.QualityInformationID=" & TxtQualityInformationID.Text & ""
tmp = tmp & " Order By G.StackNo, G.LotNo, G.GSLID"

Call TmpTable

tID = 0
mRow = 2

For I = 2 To TmpRs.RecordCount + 1
    If tID <> TmpRs!GSLID Then
       tID = TmpRs!GSLID
       For C = 0 To TmpRs.Fields.Count - 3
           MSHFlexGrid2.TextMatrix(mRow, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
       Next
       
       For C = 5 To MSHFlexGrid2.Cols - 1
           If MSHFlexGrid2.TextMatrix(0, C) = TmpRs!QualityParameterID Then
              MSHFlexGrid2.TextMatrix(mRow, C) = TmpRs!TestValue
              Exit For
           End If
       Next
       MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
       mRow = mRow + 1
    Else
       For C = 5 To MSHFlexGrid2.Cols - 1
           If MSHFlexGrid2.TextMatrix(0, C) = TmpRs!QualityParameterID Then
              MSHFlexGrid2.TextMatrix(mRow - 1, C) = TmpRs!TestValue
              Exit For
           End If
       Next
    End If
    
    TmpRs.MoveNext
Next


tmp = "Select G.StackNo, '','',sum(isnull(A.BalanceBags,0)) 'BalanceBags', '',A.QualityParameterID,A.TestValue "
tmp = tmp & " From Txn_QualityInformationDetail A"
tmp = tmp & " Inner join Mst_GSL G On G.GSLID=A.GSLID And G.CMPID=" & mCMPID & " And G.SM_Prefix='" & mSM_Prefix & "'"
tmp = tmp & " Inner Join Mst_ExchangeType E On G.ExchangeTypeID=E.ExchangeTypeID"
tmp = tmp & " Where A.QualityInformationID=" & TxtQualityInformationID.Text & ""
tmp = tmp & " Group By G.StackNo,A.QualityParameterID,A.TestValue"
tmp = tmp & " Order By G.StackNo"

Call TmpTable

tID = ""
mRow = 2

For I = 2 To TmpRs.RecordCount + 1
    If tID <> TmpRs!StackNo Then
       tID = TmpRs!StackNo
       For C = 0 To TmpRs.Fields.Count - 3
           MSHFlexGrid3.TextMatrix(mRow, C) = TmpRs.Fields(C)
       Next
       For C = 5 To MSHFlexGrid3.Cols - 1
           If MSHFlexGrid3.TextMatrix(0, C) = TmpRs!QualityParameterID Then
              MSHFlexGrid3.TextMatrix(mRow, C) = TmpRs!TestValue
              Exit For
           End If
       Next
       MSHFlexGrid3.Rows = MSHFlexGrid3.Rows + 1
       mRow = mRow + 1
    Else
       For C = 5 To MSHFlexGrid3.Cols - 1
           If MSHFlexGrid3.TextMatrix(0, C) = TmpRs!QualityParameterID Then
              MSHFlexGrid3.TextMatrix(mRow - 1, C) = TmpRs!TestValue
              Exit For
           End If
       Next
    End If
    
    TmpRs.MoveNext
Next


End Sub

Private Sub Form_Resize()
If Me.WindowState = vbMaximized Then
   SetScrollBars
End If
End Sub

Private Sub MSHFlexGrid1_Click()
Dim x As Variant
Dim mCaId As Variant

x = MSHFlexGrid1.RowSel
mCaId = MSHFlexGrid1.TextMatrix(x, 0)
If mCaId = "" Then Exit Sub
RsTxn_QualityInformation.MoveFirst

If RsTxn_QualityInformation.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
RsTxn_QualityInformation.MoveFirst
RsTxn_QualityInformation.Find "QualityInformationID = " & Val(mCaId) & ""
If RsTxn_QualityInformation.EOF Then
   RsTxn_QualityInformation.MoveLast
End If
Call Indata
Frame0.Visible = False
End Sub

Private Sub MSHFlexGrid2_Click()
Dim x As Variant
x = MSHFlexGrid2.RowSel
If SaveCmd.Enabled = False Then Exit Sub
If MSHFlexGrid2.TextMatrix(x, 0) = "" Then Exit Sub
If OptStackwise.Value = True Then Exit Sub
FrameDetail.Enabled = True

TxtStackNo.Text = MSHFlexGrid2.TextMatrix(x, 0)
TxtLotNo.Text = MSHFlexGrid2.TextMatrix(x, 1)
TxtExchangeType.Text = MSHFlexGrid2.TextMatrix(x, 2)
TxtNoofBags.Text = MSHFlexGrid2.TextMatrix(x, 3)
'CmbQualityParameter.Text = CmbQualityParameter.List(MSHFlexGrid2.ColSel - 5)  ' = MSHFlexGrid2.TextMatrix(1, MSHFlexGrid2.ColSel)
'Call CmbQualityParameter_LostFocus
'TxtTestValue.Text = MSHFlexGrid2.TextMatrix(x, MSHFlexGrid2.ColSel)
'CmbQualityParameter.Text = ""
TxtMinValue.Text = ""
TxtMaxValue.Text = ""
TxtTestValue.Text = ""
End Sub

Private Sub MSHFlexGrid3_Click()

Dim x As Variant
x = MSHFlexGrid3.RowSel
If SaveCmd.Enabled = False Then Exit Sub
If MSHFlexGrid2.TextMatrix(x, 0) = "" Then Exit Sub
If OptLotWise.Value = True Then Exit Sub
FrameDetail.Enabled = True

TxtStackNo.Text = MSHFlexGrid3.TextMatrix(x, 0)
TxtLotNo.Text = "" 'MSHFlexGrid2.TextMatrix(x, 1)
TxtExchangeType.Text = "" 'MSHFlexGrid2.TextMatrix(x, 2)
TxtNoofBags.Text = MSHFlexGrid3.TextMatrix(x, 3)
'CmbQualityParameter.Text = CmbQualityParameter.List(MSHFlexGrid2.ColSel - 5)  ' = MSHFlexGrid2.TextMatrix(1, MSHFlexGrid2.ColSel)
'Call CmbQualityParameter_LostFocus
'TxtTestValue.Text = MSHFlexGrid2.TextMatrix(x, MSHFlexGrid2.ColSel)
'CmbQualityParameter.Text = ""
TxtMinValue.Text = ""
TxtMaxValue.Text = ""
TxtTestValue.Text = ""

End Sub

Private Sub OptLotWise_Click()
MSHFlexGrid2.Visible = True
MSHFlexGrid3.Visible = False
LbLLotNo.Visible = True
TxtLotNo.Visible = True
TxtExchangeType.Visible = True
LblExchange.Visible = True
TxtStackNo.Width = 930


If SaveCmd.Enabled = True And Edit_Flg = "A" Then
   Call FillQualityParameter
   Call FillStackLotInfo
   Call FillQualityParameterStack
   Call FillStackInfo
ElseIf SaveCmd.Enabled = False Then
   Call InDataDetail
End If



End Sub

Private Sub OptStackwise_Click()

MSHFlexGrid2.Visible = False
MSHFlexGrid3.Visible = True
LbLLotNo.Visible = False
TxtLotNo.Visible = False
TxtExchangeType.Visible = False
LblExchange.Visible = False
TxtStackNo.Width = 4020

If SaveCmd.Enabled = True And Edit_Flg = "A" Then
   Call FillQualityParameter
   Call FillStackLotInfo
   Call FillQualityParameterStack
   Call FillStackInfo
ElseIf SaveCmd.Enabled = False Then
   Call InDataDetail
End If

End Sub

Private Sub SaveCmd_Click()


If CmbEmployeeID.Text = "" Then
   MsgBox "Blank Employee Name Not Allowed !!"
   CmbEmployeeID.SetFocus
   Exit Sub
End If

If CmbLocationID.Text = "" Then
   MsgBox "Blank Location Not Allowed !!"
   CmbLocationID.SetFocus
   Exit Sub
End If

If CmbCMPName.Text = "" Then
   MsgBox " Blank CMP Name Not Allowed !!"
   CmbCMPName.SetFocus
   Exit Sub
End If

If CmbGodownNo.Text = "" Then
   MsgBox "Blank Godown No Not Allowed !!"
   CmbGodownNo.SetFocus
   Exit Sub
End If

If CmbCommodityID.Text = "" Then
   MsgBox "Blank Commodity Not Allowed !!"
   CmbCommodityID.SetFocus
   Exit Sub
End If

If CmbTestingAgency.Text = "" Then
   MsgBox "Blank Testing Agency Not Allowed !!"
   CmbTestingAgency.SetFocus
   Exit Sub
End If

If CheckGrid = False Then
   MsgBox "Please Enter Values For Quality Parameters !!"
   Exit Sub
End If

If TxtDate.Value > Date Then
   MsgBox "Future Date not allowed!!"
   TxtDate.SetFocus
   Exit Sub
End If

If Format(Gen_BlockingDate, "yyyy-mm") >= Format(TxtDate.Value, "yyyy-mm") Then
   MsgBox "Date allowed for next month of " & Format(Gen_BlockingDate, "MMM-yyyy") & "!!!"
   TxtDate.SetFocus
   Exit Sub
End If

If Edit_Flg = "A" Then
'   RsTxn_QualityInformation.Filter = "CommodityID=" & mCommodityID & " And GodownID =" & mGodownID & ""
'   If RsTxn_QualityInformation.RecordCount > 0 Then
'      MsgBox "Duplicate Quality Information Not Allowed !!"
'      RsTxn_QualityInformation.Filter = ""
'      Exit Sub
'   End If
   RsTxn_QualityInformation.AddNew
   RsTxn_QualityInformation!QualityInformationID = GetMaxID("QualityInformationID", "Txn_QualityInformation")
   TxtQualityInformationID.Text = RsTxn_QualityInformation!QualityInformationID
   RsTxn_QualityInformation!G_UID = GetGUID
Else
'   RsTxn_QualityInformation.Filter = "CommodityID=" & mCommodityID & " And GodownID =" & mGodownID & ""
'   If RsTxn_QualityInformation.RecordCount > 0 Then
'      If RsTxn_QualityInformation!QualityInformationID <> TxtQualityInformationID Then
'         MsgBox "Duplicate Quality Information Record Not Allowed !!"
'         Exit Sub
'      End If
'   End If
   RsTxn_QualityInformation.MoveFirst
   RsTxn_QualityInformation.Find "QualityInformationID = " & TxtQualityInformationID & ""
End If

RsTxn_QualityInformation!Date = TxtDate.Value
RsTxn_QualityInformation!EmployeeID = mEmployeeID
RsTxn_QualityInformation!GodownID = mGodownID
RsTxn_QualityInformation!CommodityID = mCommodityID
RsTxn_QualityInformation!TestingAgencyID = mAgencyID
RsTxn_QualityInformation!Remarks = TxtRemarks

If OptLotWise.Value = True Then
   RsTxn_QualityInformation!DetailsEnteredFlg = "L"
Else
   RsTxn_QualityInformation!DetailsEnteredFlg = "S"
End If

If IsNull(RsTxn_QualityInformation!CreatedBy) = True Or RsTxn_QualityInformation!CreatedBy = "" Then
   RsTxn_QualityInformation!CreatedBy = Gen_UserLoginID
   RsTxn_QualityInformation!CreatedDate = Now
Else
   RsTxn_QualityInformation!UpdatedBy = Gen_UserLoginID
   RsTxn_QualityInformation!UpdatedDate = Now
End If

RsTxn_QualityInformation.Update

Call SaveDetails

Call Indata

End Sub

Private Sub SaveDetails()
Call TableTxn_QualityInformationDetail
For I = 2 To MSHFlexGrid2.Rows - 1
    For C = 5 To MSHFlexGrid2.Cols - 1
        If MSHFlexGrid2.TextMatrix(I, C) <> "" Then
           RsTxn_QualityInformationDetail.Filter = " QualityInformationID=" & TxtQualityInformationID.Text & " And QualityParameterID=" & MSHFlexGrid2.TextMatrix(0, C) & " And GSLID=" & MSHFlexGrid2.TextMatrix(I, 4) & ""
           If RsTxn_QualityInformationDetail.RecordCount = 0 Then
              RsTxn_QualityInformationDetail.Filter = ""
              RsTxn_QualityInformationDetail.AddNew
              RsTxn_QualityInformationDetail!QualityInformationDetailID = GetMaxID("QualityInformationDetailID", "Txn_QualityInformationDetail")
              RsTxn_QualityInformationDetail!G_UID = GetGUID
           End If
           RsTxn_QualityInformationDetail!QualityInformationID = Val(TxtQualityInformationID.Text)
           RsTxn_QualityInformationDetail!QualityParameterID = Val(MSHFlexGrid2.TextMatrix(0, C))
           RsTxn_QualityInformationDetail!GSLID = Val(MSHFlexGrid2.TextMatrix(I, 4))
           RsTxn_QualityInformationDetail!BalanceBags = Val(MSHFlexGrid2.TextMatrix(I, 3))
           RsTxn_QualityInformationDetail!SM_Prefix = mSM_Prefix
           RsTxn_QualityInformationDetail!TestValue = Val(MSHFlexGrid2.TextMatrix(I, C))
           If IsNull(RsTxn_QualityInformationDetail!CreatedBy) = True Or RsTxn_QualityInformationDetail!CreatedBy = "" Then
              RsTxn_QualityInformationDetail!CreatedBy = Gen_UserLoginID
              RsTxn_QualityInformationDetail!CreatedDate = Now
           Else
              RsTxn_QualityInformationDetail!UpdatedBy = Gen_UserLoginID
              RsTxn_QualityInformationDetail!UpdatedDate = Now
           End If
           RsTxn_QualityInformationDetail.Update
           RsTxn_QualityInformationDetail.Filter = ""
        End If
    Next
Next
End Sub

Private Function GetMaxID(mCol_ As Variant, mTbl_ As Variant)
tmp = "Select Max(" & mCol_ & ") as ID From " & mTbl_ & ""
Call TmpTable
GetMaxID = IIf(IsNull(TmpRs!ID), 1, TmpRs!ID + 1)
End Function

Private Function CheckGrid() As Boolean
Dim Retval As Boolean

Retval = False
If OptLotWise.Value = True Then
   For I = 2 To MSHFlexGrid2.Rows - 1
       For C = 5 To MSHFlexGrid2.Cols - 1
           If MSHFlexGrid2.TextMatrix(I, C) <> "" Then
              Retval = True
              Exit For
           End If
      Next
      If Retval = True Then Exit For
  Next
Else
   For I = 2 To MSHFlexGrid3.Rows - 1
       For C = 5 To MSHFlexGrid3.Cols - 1
           If MSHFlexGrid3.TextMatrix(I, C) <> "" Then
              Retval = True
              Exit For
           End If
      Next
      If Retval = True Then Exit For
  Next
End If

CheckGrid = Retval
End Function
Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsTxn_QualityInformation.MoveNext
If RsTxn_QualityInformation.EOF Then
   RsTxn_QualityInformation.MoveLast
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
RsTxn_QualityInformation.MovePrevious
If RsTxn_QualityInformation.BOF Then
   RsTxn_QualityInformation.MoveFirst
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
RsTxn_QualityInformation.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsTxn_QualityInformation.MoveLast
Call Indata
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False
End Sub

Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   
   If RsTxn_QualityInformation.RecordCount = 0 Then Exit Sub
   Edit_Flg = ""
   Call GButtonLock(Me, True)
   If RsTxn_QualityInformation.EOF Then
      RsTxn_QualityInformation.MoveFirst
      RsTxn_QualityInformation.Find "QualityInformationID = " & TxtQualityInformationID & ""
   End If
   Call Indata
   Exit Sub
End If
Frame0.Height = Frame4.Height - 200
Frame0.Left = Frame1.Left - 100
Frame0.Top = Frame1.Top
Frame0.Width = Frame4.Width - 50
Frame0.Visible = True

MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 2600

RsTxn_QualityInformation.MoveFirst

''===========comment on 7 july 2012 for not clearing search grid=============
'Call Grid_Head
'SFrom.Value = Date - 365
'STo.Value = Date
'Label22.Caption = ""
'CmbSEmployeeID.Text = ""
'CmbSLocationID.Text = ""
'CmbSCMPName.Text = ""
'CmbSGodownNo.Text = ""
'CmbSTestingAgency.Text = ""
'TxtSQualityInfoID.Text = ""
''===========comment on 7 july 2012 for not clearing search grid=============

End Sub

Private Sub TxtTestValue_LostFocus()
If TxtTestValue.Text = "" Then Exit Sub
If IsNumeric(TxtTestValue.Text) = False Then
   MsgBox "Invalid Numeric Format."
   TxtTestValue.SetFocus
   Exit Sub
End If
End Sub
Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Txn_QualityInformation.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Txn_QualityInformation.xls")
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
Private Sub CmbSEmployeeID_GotFocus()

tmp = CmbSEmployeeID.Text
Call TableMst_Employee("Where Flag = 'Y' ")
CmbSEmployeeID.Clear
If RsMst_Employee.RecordCount = 0 Then
   MsgBox "No Employee Found!!!"
   CmbSEmployeeID.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Employee.RecordCount
    CmbSEmployeeID.AddItem RsMst_Employee!EmployeeName & "~" & RsMst_Employee!EmployeeNo
    RsMst_Employee.MoveNext
Next
CmbSEmployeeID.Text = tmp
End Sub

Private Sub CmbSEmployeeID_LostFocus()
Dim strSplitString() As String
Dim EmpNo_, EmpName_ As String

If CmbSEmployeeID.Text = "" Then
   mSEmployeeID = Null
   Exit Sub
End If

strSplitString = Split(CmbSEmployeeID, "~")
If UBound(strSplitString) <> 1 Then
   MsgBox "Invalid Selection."
   CmbSEmployeeID.SetFocus
   Exit Sub
End If

EmpName_ = Trim(strSplitString(0))
EmpNo_ = Trim(strSplitString(1))

RsMst_Employee.MoveFirst
RsMst_Employee.Find "EmployeeNo = '" & EmpNo_ & "'"

If RsMst_Employee.EOF Then
   MsgBox "Invalid Selection "
   CmbSEmployeeID.SetFocus
   Exit Sub
End If
mSEmployeeID = RsMst_Employee!EmployeeID
End Sub

Private Sub CmbSTestingAgency_GotFocus()
tmp = CmbSTestingAgency.Text
CmbSTestingAgency.Clear
Call TableMst_TestingAgency
If RsMst_TestingAgency.RecordCount = 0 Then
   MsgBox "Testing Agency Not Found"
   Exit Sub
End If
For I = 1 To RsMst_TestingAgency.RecordCount
    CmbSTestingAgency.AddItem RsMst_TestingAgency!AgencyName
    RsMst_TestingAgency.MoveNext
Next
CmbSTestingAgency.Text = tmp
End Sub

Private Sub CmbSTestingAgency_LostFocus()
If CmbSTestingAgency.Text = "" Then Exit Sub
RsMst_TestingAgency.MoveFirst
RsMst_TestingAgency.Find "AgencyName='" & CmbSTestingAgency.Text & "'"
If RsMst_TestingAgency.EOF Then
   MsgBox "Invalid Selection."
   Exit Sub
End If
mSAgencyID = RsMst_TestingAgency!TestingAgencyID
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
'----
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
mSCommodityID = RsMst_Commodity!CommodityID
End Sub

Private Sub Form_Activate()
If Me.WindowState = vbMaximized Then
   SetScrollBars
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

