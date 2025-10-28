VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmTxn_LockSealHologram 
   Caption         =   "Lock-Seal Hologram"
   ClientHeight    =   9285
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10650
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9285
   ScaleWidth      =   10650
   WindowState     =   2  'Maximized
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      TabIndex        =   79
      Top             =   10560
      Width           =   10935
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   9015
      Left            =   15000
      TabIndex        =   78
      Top             =   0
      Width           =   255
   End
   Begin VB.Frame Frame4 
      Height          =   10215
      Left            =   120
      TabIndex        =   16
      Top             =   0
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
         Height          =   2205
         Left            =   720
         TabIndex        =   61
         Top             =   5520
         Visible         =   0   'False
         Width           =   14835
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
            ItemData        =   "FrmTxn_LockSealHologram.frx":0000
            Left            =   9570
            List            =   "FrmTxn_LockSealHologram.frx":0002
            Sorted          =   -1  'True
            TabIndex        =   69
            Top             =   555
            Width           =   4455
         End
         Begin VB.TextBox TxtLockSealHologramNo 
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
            Left            =   1680
            MaxLength       =   12
            TabIndex        =   71
            Top             =   1020
            Width           =   1755
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
            Left            =   9585
            TabIndex        =   73
            Top             =   975
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
            Left            =   10605
            TabIndex        =   75
            Top             =   975
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
            ItemData        =   "FrmTxn_LockSealHologram.frx":0004
            Left            =   6330
            List            =   "FrmTxn_LockSealHologram.frx":0006
            Sorted          =   -1  'True
            TabIndex        =   68
            Top             =   555
            Width           =   3240
         End
         Begin VB.ComboBox CmbSStateID 
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
            ItemData        =   "FrmTxn_LockSealHologram.frx":0008
            Left            =   3210
            List            =   "FrmTxn_LockSealHologram.frx":000A
            Sorted          =   -1  'True
            TabIndex        =   66
            Top             =   555
            Width           =   3120
         End
         Begin MSComCtl2.DTPicker STo 
            Height          =   375
            Left            =   1695
            TabIndex        =   64
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
            Format          =   112263169
            CurrentDate     =   36494
         End
         Begin MSComCtl2.DTPicker SFrom 
            Height          =   375
            Left            =   135
            TabIndex        =   62
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
            Format          =   112263169
            CurrentDate     =   36494
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   1455
            Left            =   120
            TabIndex        =   63
            Top             =   1680
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
         Begin VB.Label Label12 
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
            Height          =   255
            Left            =   11445
            TabIndex        =   80
            Top             =   210
            Width           =   930
         End
         Begin VB.Label Label1 
            Caption         =   "Hologram No."
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
            Left            =   135
            TabIndex        =   77
            Top             =   1065
            Width           =   1320
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
            TabIndex        =   76
            Top             =   210
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
            Left            =   360
            TabIndex        =   74
            Top             =   210
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
            Left            =   6330
            TabIndex        =   72
            Top             =   1035
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
            Left            =   8805
            TabIndex        =   70
            Top             =   1050
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
            Left            =   7440
            TabIndex        =   67
            Top             =   210
            Width           =   765
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "State"
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
            Left            =   4440
            TabIndex        =   65
            Top             =   210
            Width           =   465
         End
      End
      Begin VB.Frame ButtonFrm 
         Height          =   1155
         Left            =   120
         TabIndex        =   50
         Top             =   8520
         Width           =   14760
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
            TabIndex        =   60
            Top             =   240
            Width           =   2895
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
            Left            =   3015
            TabIndex        =   59
            Top             =   240
            Width           =   2895
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
            Left            =   5910
            TabIndex        =   58
            Top             =   240
            Width           =   2895
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
            Left            =   8805
            TabIndex        =   57
            Top             =   240
            Width           =   2895
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
            Left            =   11700
            TabIndex        =   56
            Top             =   240
            Width           =   2895
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
            TabIndex        =   55
            Top             =   615
            Width           =   2895
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
            Left            =   3015
            TabIndex        =   54
            Top             =   615
            Width           =   2895
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
            Left            =   5910
            TabIndex        =   53
            Top             =   615
            Width           =   2895
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
            Left            =   8805
            TabIndex        =   52
            Top             =   615
            Width           =   2895
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
            Left            =   11700
            TabIndex        =   51
            Top             =   615
            Width           =   2895
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
         Height          =   2970
         Left            =   120
         TabIndex        =   17
         Top             =   120
         Width           =   14655
         Begin VB.TextBox TxtTotalQty 
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
            Left            =   9825
            Locked          =   -1  'True
            MaxLength       =   12
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   1824
            Width           =   1755
         End
         Begin VB.TextBox TxtLockSealHologramID 
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
            Left            =   2280
            Locked          =   -1  'True
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   285
            Width           =   1755
         End
         Begin VB.ComboBox CmbStateID 
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
            ItemData        =   "FrmTxn_LockSealHologram.frx":000C
            Left            =   9840
            List            =   "FrmTxn_LockSealHologram.frx":000E
            Sorted          =   -1  'True
            TabIndex        =   1
            Top             =   285
            Width           =   4695
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
            Height          =   540
            Left            =   2280
            MaxLength       =   1000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   7
            Top             =   2280
            Width           =   12180
         End
         Begin VB.TextBox TxtFromNo 
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
            Left            =   2280
            MaxLength       =   12
            TabIndex        =   5
            Top             =   1824
            Width           =   1755
         End
         Begin VB.TextBox TxtToNo 
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
            Left            =   6300
            MaxLength       =   12
            TabIndex        =   6
            Top             =   1824
            Width           =   1515
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
            Left            =   9825
            Sorted          =   -1  'True
            TabIndex        =   3
            Top             =   738
            Width           =   4695
         End
         Begin VB.TextBox TxtAddress 
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
            Height          =   540
            Left            =   2295
            MaxLength       =   2000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   4
            Top             =   1191
            Width           =   12210
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
            ItemData        =   "FrmTxn_LockSealHologram.frx":0010
            Left            =   2280
            List            =   "FrmTxn_LockSealHologram.frx":0012
            Sorted          =   -1  'True
            TabIndex        =   2
            Top             =   750
            Width           =   5535
         End
         Begin MSComCtl2.DTPicker TxtLockSealHologramDate 
            Height          =   360
            Left            =   6300
            TabIndex        =   0
            Top             =   285
            Width           =   1515
            _ExtentX        =   2672
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
            Format          =   112263169
            CurrentDate     =   39884
         End
         Begin VB.TextBox TxtStateID 
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
            Left            =   9840
            Locked          =   -1  'True
            TabIndex        =   20
            Top             =   285
            Visible         =   0   'False
            Width           =   4695
         End
         Begin VB.TextBox TxtEmployeeID 
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
            Left            =   9825
            Locked          =   -1  'True
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   738
            Width           =   4695
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
            Left            =   2280
            Locked          =   -1  'True
            TabIndex        =   22
            Top             =   750
            Visible         =   0   'False
            Width           =   5535
         End
         Begin VB.Label Label9 
            Caption         =   "Total"
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
            Left            =   8280
            TabIndex        =   32
            Top             =   1877
            Width           =   855
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
            Left            =   4680
            TabIndex        =   31
            Top             =   315
            Width           =   735
         End
         Begin VB.Label Label4 
            Caption         =   "Lock Seal Hologram ID"
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
            TabIndex        =   30
            Top             =   195
            Width           =   1815
         End
         Begin VB.Label LblLocationID 
            Caption         =   "State"
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
            Left            =   8265
            TabIndex        =   29
            Top             =   338
            Width           =   930
         End
         Begin VB.Label Label5 
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
            Height          =   255
            Left            =   240
            TabIndex        =   28
            Top             =   2488
            Width           =   960
         End
         Begin VB.Label Label10 
            Caption         =   "From Lock Seal Hologram No."
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
            Left            =   240
            TabIndex        =   27
            Top             =   1742
            Width           =   1800
         End
         Begin VB.Label Label11 
            Caption         =   "To Lock Seal Hologram No."
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
            Left            =   4680
            TabIndex        =   26
            Top             =   1734
            Width           =   1860
         End
         Begin VB.Label Label27 
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
            Left            =   8280
            TabIndex        =   25
            Top             =   783
            Width           =   1065
         End
         Begin VB.Label Label29 
            Caption         =   "Address"
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
            TabIndex        =   24
            Top             =   1334
            Width           =   960
         End
         Begin VB.Label Label31 
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
            Left            =   240
            TabIndex        =   23
            Top             =   810
            Width           =   930
         End
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   5340
         Left            =   120
         TabIndex        =   33
         Top             =   3120
         Width           =   14745
         _ExtentX        =   26009
         _ExtentY        =   9419
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
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
         TabCaption(0)   =   "Issue Details"
         TabPicture(0)   =   "FrmTxn_LockSealHologram.frx":0014
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "MSHFlexGrid2"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Frame2"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Log Details"
         TabPicture(1)   =   "FrmTxn_LockSealHologram.frx":0030
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "TxtCreated"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "TxtUpdated"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "MSHFlexGrid4"
         Tab(1).Control(3)=   "Label24"
         Tab(1).Control(4)=   "Label25"
         Tab(1).ControlCount=   5
         Begin VB.Frame Frame2 
            Height          =   1500
            Left            =   120
            TabIndex        =   36
            Top             =   360
            Width           =   14490
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
               Left            =   12690
               TabIndex        =   14
               Top             =   465
               Width           =   1695
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
               Left            =   12690
               TabIndex        =   15
               Top             =   855
               Width           =   1710
            End
            Begin VB.TextBox TxtITotalQty 
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
               Left            =   7995
               Locked          =   -1  'True
               MaxLength       =   12
               TabIndex        =   37
               TabStop         =   0   'False
               Top             =   525
               Width           =   1230
            End
            Begin VB.TextBox TxtIAddress 
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
               Height          =   420
               Left            =   1500
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   13
               Top             =   960
               Width           =   11145
            End
            Begin VB.ComboBox CmbILocationID 
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
               ItemData        =   "FrmTxn_LockSealHologram.frx":004C
               Left            =   1485
               List            =   "FrmTxn_LockSealHologram.frx":004E
               Sorted          =   -1  'True
               TabIndex        =   9
               Top             =   525
               Width           =   2895
            End
            Begin VB.TextBox TxtIFromNo 
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
               Left            =   4410
               MaxLength       =   12
               TabIndex        =   10
               Top             =   525
               Width           =   1755
            End
            Begin VB.TextBox TxtIToNo 
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
               Left            =   6195
               MaxLength       =   12
               TabIndex        =   11
               Top             =   525
               Width           =   1770
            End
            Begin VB.ComboBox CmbIEmployeeID 
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
               ItemData        =   "FrmTxn_LockSealHologram.frx":0050
               Left            =   9285
               List            =   "FrmTxn_LockSealHologram.frx":0052
               Sorted          =   -1  'True
               TabIndex        =   12
               Top             =   525
               Width           =   3375
            End
            Begin MSComCtl2.DTPicker TxtIDate 
               Height          =   360
               Left            =   60
               TabIndex        =   8
               Top             =   525
               Width           =   1395
               _ExtentX        =   2461
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
               Format          =   112263169
               CurrentDate     =   39884
            End
            Begin VB.Label LblTot 
               Caption         =   "Total"
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
               Left            =   8310
               TabIndex        =   44
               Top             =   195
               Width           =   495
            End
            Begin VB.Label Label6 
               Caption         =   "Address"
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
               Left            =   195
               TabIndex        =   43
               Top             =   1005
               Width           =   1200
            End
            Begin VB.Label Label2 
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
               Height          =   255
               Left            =   2520
               TabIndex        =   42
               Top             =   195
               Width           =   930
            End
            Begin VB.Label Label7 
               Caption         =   "From Hologram No"
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
               Left            =   4410
               TabIndex        =   41
               Top             =   195
               Width           =   1815
            End
            Begin VB.Label Label19 
               Caption         =   "To Hologram No"
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
               Left            =   6300
               TabIndex        =   40
               Top             =   195
               Width           =   1620
            End
            Begin VB.Label Label26 
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
               Left            =   360
               TabIndex        =   39
               Top             =   195
               Width           =   735
            End
            Begin VB.Label Label30 
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
               Height          =   255
               Left            =   10530
               TabIndex        =   38
               Top             =   210
               Width           =   930
            End
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
            Left            =   -72840
            Locked          =   -1  'True
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   780
            Width           =   8655
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
            Left            =   -72840
            Locked          =   -1  'True
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   1275
            Width           =   8655
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
            Height          =   3255
            Left            =   120
            TabIndex        =   45
            Top             =   1920
            Width           =   14460
            _ExtentX        =   25506
            _ExtentY        =   5741
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
            Height          =   3375
            Left            =   -74880
            TabIndex        =   81
            Top             =   1800
            Width           =   14460
            _ExtentX        =   25506
            _ExtentY        =   5953
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
         Begin VB.Label Label21 
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
            TabIndex        =   49
            Top             =   540
            Width           =   855
         End
         Begin VB.Label Label20 
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
            TabIndex        =   48
            Top             =   945
            Width           =   885
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
            Left            =   -74880
            TabIndex        =   47
            Top             =   840
            Width           =   1695
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
            Left            =   -74880
            TabIndex        =   46
            Top             =   1335
            Width           =   975
         End
      End
   End
End
Attribute VB_Name = "FrmTxn_LockSealHologram"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mVendorID, mLocationID, mCommodityID, mSVendorID, mSLocationID, mSCommodityID As Variant
Dim mCMPID As Variant
Dim Edit_Flg As Variant
Dim mStateID, mSStateID, mEmployeeID, mILocationID, mIEmployeeID, mSEmployeeID As Variant
Private Sub AddCmd_Click()
Edit_Flg = "A"

Frame1.Enabled = True
TxtLockSealHologramID.Text = ""

CmbLocationID.Text = ""
TxtLocationID.Text = ""

TxtLocationID.Visible = False
CmbLocationID.Visible = True

CmbStateID.Text = ""
TxtStateID.Text = ""

TxtStateID.Visible = False
CmbStateID.Visible = True

CmbEmployeeID.Text = ""
TxtEmployeeID.Text = ""

TxtEmployeeID.Visible = False
CmbEmployeeID.Visible = True

TxtFromNo.Text = ""
TxtToNo.Text = ""
TxtTotalQty.Text = ""
TxtAddress = ""
TxtRemark.Text = ""
TxtCreated = ""
TxtUpdated = ""

TxtLockSealHologramDate.Value = Date
Call Grid_Head1
Call Grid_Head3
Call ClearDetail
Frame2.Enabled = True
Call GButtonLock(Me, False)
If RsTxn_LockSealHologram.RecordCount > 0 Then
   TxtLockSealHologramDate.SetFocus
End If
End Sub

Private Sub CmbEmployeeID_GotFocus()
tmp = CmbEmployeeID.Text
Call TableMst_Employee("Where Flag = 'Y'")
CmbEmployeeID.Clear
If RsMst_Employee.RecordCount = 0 Then
   MsgBox "No Employee Found!!!"
   CmbEmployeeID.SetFocus
   'RsMst_Employee.MoveFirst
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
   MsgBox "Invalid Selection "
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
End Sub

Private Sub CmbLocationID_GotFocus()
If CmbStateID.Text = "" Then
   MsgBox " Please Select State!!"
   CmbStateID.SetFocus
   Exit Sub
End If
tmp = CmbLocationID.Text

If Gen_WcLocation <> "" Then
   Call TableMst_Location(" Where StateID=" & mStateID & " And " & Gen_WcLocation & " ")
Else
   Call TableMst_Location(" Where StateID=" & mStateID)
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
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocationID.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection "
   CmbLocationID.SetFocus
   Exit Sub
End If

mLocationID = RsMst_Location!LocationID
End Sub

Private Sub CmdAdd2List_Click()

If TxtIDate.Value > Date Then
   MsgBox "Future Date not allowed!!!"
   TxtIDate.SetFocus
   Exit Sub
End If

If TxtIDate.Value < TxtLockSealHologramDate Then
   MsgBox "Date should not less than above Date!!!"
   TxtIDate.SetFocus
   Exit Sub
End If

If CmbILocationID.Text = "" Then
   MsgBox "Blank Location not Allowed!!!"
   CmbILocationID.SetFocus
   Exit Sub
End If

If TxtIFromNo.Text = "" Then
   MsgBox "Blank From Hologram No not Allowed!!!"
   TxtIFromNo.SetFocus
   Exit Sub
End If

If TxtIToNo.Text = "" Then
   MsgBox "Blank To Hologram No not Allowed!!!"
   TxtIToNo.SetFocus
   Exit Sub
End If

If CmbIEmployeeID.Text = "" Then
   MsgBox "Blank Employee not Allowed!!!"
   CmbIEmployeeID.SetFocus
   Exit Sub
End If

If TxtIAddress.Text = "" Then
   MsgBox "Blank Address not Allowed!!!"
   TxtIAddress.SetFocus
   Exit Sub
End If

If Val(TxtIFromNo.Text) < Val(TxtFromNo) Or Val(TxtIFromNo.Text) > Val(TxtToNo) Then
   MsgBox "'From Hologram No' Should be in above range!!!"
   TxtIFromNo.SetFocus
   Exit Sub
End If

If Val(TxtIToNo.Text) < Val(TxtFromNo) Or Val(TxtIToNo.Text) > Val(TxtToNo) Then
   MsgBox "'To Hologram No' Should be in above range!!!"
   TxtIToNo.SetFocus
   Exit Sub
End If

If Val(TxtIToNo.Text) < Val(TxtIFromNo.Text) Then
   MsgBox "'To Hologram No' must be more than 'From Hologram No'"
   TxtIToNo.SetFocus
   Exit Sub
End If

If CheckDuplicateHologramNo = True Then
   MsgBox "'From Hologram No' or 'To Hologram No' in already used in range. Please change range!!! "
   TxtIFromNo.SetFocus
   Exit Sub
End If

If CmdAdd2List.Caption = "Update List" Then
   r = MSHFlexGrid2.RowSel
Else
   r = MSHFlexGrid2.Rows - 1
   MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
End If

MSHFlexGrid2.TextMatrix(r, 0) = Format(TxtIDate.Value, "dd-MMM-yyyy")
MSHFlexGrid2.TextMatrix(r, 1) = CmbILocationID.Text
MSHFlexGrid2.TextMatrix(r, 2) = mILocationID
MSHFlexGrid2.TextMatrix(r, 3) = TxtIFromNo.Text
MSHFlexGrid2.TextMatrix(r, 4) = TxtIToNo.Text
MSHFlexGrid2.TextMatrix(r, 5) = TxtITotalQty.Text
MSHFlexGrid2.TextMatrix(r, 6) = CmbIEmployeeID.Text
MSHFlexGrid2.TextMatrix(r, 7) = mIEmployeeID
MSHFlexGrid2.TextMatrix(r, 8) = TxtIAddress

Call ClearDetail
TxtIDate.SetFocus

End Sub


Private Sub CmdDeleteList_Click()
I = MSHFlexGrid2.RowSel
If I > 0 Then
   If MSHFlexGrid2.TextMatrix(I, 0) <> "" Then
      For RO = I To MSHFlexGrid2.Rows - 2
          For C = 0 To MSHFlexGrid2.Cols - 1
              MSHFlexGrid2.TextMatrix(RO, C) = MSHFlexGrid2.TextMatrix(RO + 1, C)
          Next
      Next
      MSHFlexGrid2.Rows = MSHFlexGrid2.Rows - 1
   End If
End If
Call ClearDetail

End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Txn_LockSealHologram.xls")
Call Xls_Detail_Rpt
End Sub
Public Sub Xls_Detail_Rpt()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Double
Dim mCol As Double
MsgBox ("Wait till next message")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Txn_LockSealHologram.xls")
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



Private Sub CmdSearch_Click()
Dim Wc1 As Variant
Call Grid_Head

If SFrom.Value > STo.Value Then
   MsgBox "'From Date' Must be less than 'To Date'!!!"
   SFrom.SetFocus
   Exit Sub
End If

tmp = "Select LockSealHologramID,LockSealHologramDate,StateName,"
tmp = tmp & " LocationName ,ME.EmployeeName + '~' +ME.EmployeeNo as EmployeeName,TF.Address, FromLockSealHologramNo, ToLockSealHologramNo, TotalQty, Remark"
tmp = tmp & " From Txn_LockSealHologram TF"
tmp = tmp & " Inner Join Mst_Location ML on TF.LocationID=ML.LocationID"
tmp = tmp & " Inner Join Mst_Employee ME on TF.EmployeeID=ME.EmployeeID"
tmp = tmp & " Inner Join Mst_State MS on MS.StateID=ML.StateID"

If Gen_DBType = "MDB" Then
   Wc1 = " Where LockSealHologramDate Between #" & Format(SFrom.Value, Gen_DatFormat) & "# And #" & Format(STo.Value, Gen_DatFormat) & "# "
Else
   Wc1 = " Where LockSealHologramDate Between '" & Format(SFrom.Value, Gen_DatFormat) & "' And '" & Format(STo.Value, Gen_DatFormat) & "' "
End If

If CmbSStateID.Text <> "" Then
   Wc1 = Wc1 & " And StateName='" & CmbSStateID.Text & "'"
Else
   If Gen_WcState <> "" Then
      Wc1 = Wc1 & " And MS." & Gen_WcState
   End If
End If

If CmbSLocationID.Text <> "" Then
   Wc1 = Wc1 & " And LocationName='" & CmbSLocationID.Text & "'"
Else
   If Gen_WcLocation <> "" Then
      Wc1 = Wc1 & " And ML." & Gen_WcLocation
   End If
End If

If CmbSEmployeeID.Text <> "" Then
   Wc1 = Wc1 & " And TF.EmployeeID=" & mSEmployeeID
End If

If TxtLockSealHologramNo.Text <> "" Then
   Wc1 = Wc1 & " And FromLockSealHologramNo<=" & Val(TxtLockSealHologramNo.Text) & " And ToLockSealHologramNo>=" & Val(TxtLockSealHologramNo.Text) & ""
End If

tmp = tmp & Wc1 & " Order By FromLockSealHologramNo"

Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   MSHFlexGrid1.Rows = TmpRs1.RecordCount + 2
   For I = 1 To TmpRs1.RecordCount
       For C = 0 To MSHFlexGrid1.Cols - 1
           MSHFlexGrid1.TextMatrix(I, C) = IIf(IsNull(TmpRs1.Fields(C)), "", TmpRs1.Fields(C))
       Next
       MSHFlexGrid1.TextMatrix(I, 1) = Format(MSHFlexGrid1.TextMatrix(I, 1), "dd-MMM-yyyy")
       TmpRs1.MoveNext
   Next
   
End If

Label22.Caption = TmpRs1.RecordCount
Label22.Refresh


End Sub


Private Sub DeleteCmd_Click()

Dim ans As Variant
On Error GoTo msgError

ans = MsgBox("Do you really want to DELETE this record??? ", vbYesNo)
If ans = vbYes Then
   Call DeleteHologramDetails
   RsTxn_LockSealHologram.Delete
   RsTxn_LockSealHologram.Update
   If RsTxn_LockSealHologram.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsTxn_LockSealHologram.MoveNext
   If RsTxn_LockSealHologram.EOF() Then
      RsTxn_LockSealHologram.MoveLast
   End If
   Call Indata
End If

Exit Sub
msgError:
MsgBox "Child record Present "
RsTxn_LockSealHologram.CancelUpdate


End Sub

Private Sub EditCmd_Click()

Edit_Flg = "E"

Call GButtonLock(Me, False)

Frame1.Enabled = True
Frame2.Enabled = True

CmbStateID.Text = TxtStateID.Text
CmbStateID.Visible = True
TxtStateID.Visible = False

CmbLocationID.Text = TxtLocationID.Text
CmbLocationID.Visible = True
TxtLocationID.Visible = False

CmbEmployeeID.Text = TxtEmployeeID.Text
CmbEmployeeID.Visible = True
TxtEmployeeID.Visible = False

If MSHFlexGrid2.Rows > 2 Then
   Frame1.Enabled = False
End If

End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim wc As Variant

Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
Gen_FormAccess = FrmLoadAccess(0)
If FrmLoadAccess(0) = False Then
   Frame1.Enabled = False
   Call GButtonLockAD(Me)
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If
Call Grid_Head
SFrom.Value = Date - 365
STo.Value = Date
Label22.Caption = ""
CmbSStateID.Text = ""
CmbSEmployeeID.Text = ""
CmbSLocationID.Text = ""
TxtLockSealHologramNo.Text = ""

Frame1.Enabled = True
If Gen_WcLocation <> "" Then
   Call TableTxn_LockSealHologram(" Where " & Gen_WcLocation & " Order by LockSealHologramID")
Else
   Call TableTxn_LockSealHologram(" Order by LockSealHologramID")
End If
If RsTxn_LockSealHologram.RecordCount = 0 Then
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

Private Sub MSHFlexGrid1_Click()
Dim I As Variant
I = MSHFlexGrid1.RowSel
If I <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(I, 0) = "" Then Exit Sub
RsTxn_LockSealHologram.MoveFirst
RsTxn_LockSealHologram.Find "LockSealHologramID= " & Val(MSHFlexGrid1.TextMatrix(I, 0)) & ""
If Not RsTxn_LockSealHologram.EOF Then
   Call Indata
   Frame0.Visible = False
End If

End Sub

Private Sub MSHFlexGrid2_Click()
If SaveCmd.Enabled = False Then Exit Sub
Dim r As Variant
r = MSHFlexGrid2.RowSel
If r < 1 Or r = MSHFlexGrid2.Rows - 1 Then
   Call ClearDetail
   Exit Sub
End If
If MSHFlexGrid2.TextMatrix(r, 0) = "" Then
   Exit Sub
End If
TxtIDate.Value = CDate(MSHFlexGrid2.TextMatrix(r, 0))
CmbILocationID.Text = MSHFlexGrid2.TextMatrix(r, 1)
mILocationID = Val(MSHFlexGrid2.TextMatrix(r, 2))
TxtIFromNo = Val(MSHFlexGrid2.TextMatrix(r, 3))
TxtIToNo = Val(MSHFlexGrid2.TextMatrix(r, 4))
TxtITotalQty = Val(MSHFlexGrid2.TextMatrix(r, 5))
CmbIEmployeeID.Text = MSHFlexGrid2.TextMatrix(r, 6)
mIEmployeeID = Val(MSHFlexGrid2.TextMatrix(r, 7))
TxtIAddress.Text = MSHFlexGrid2.TextMatrix(r, 8)
CmdAdd2List.Caption = "Update List"
TxtIDate.SetFocus
End Sub

Private Sub SaveCmd_Click()

If TxtLockSealHologramDate.Value > Date Then
   MsgBox "Future Date not allowed!!"
   TxtLockSealHologramDate.SetFocus
   Exit Sub
End If

If CmbStateID.Text = "" Then
   MsgBox "Select State Name!!"
   CmbStateID.SetFocus
   Exit Sub
End If

If CmbLocationID.Text = "" Then
   MsgBox "Select Location Name!!"
   CmbLocationID.SetFocus
   Exit Sub
End If

If CmbEmployeeID.Text = "" Then
   MsgBox "Select Employee!!"
   CmbEmployeeID.SetFocus
   Exit Sub
End If

If TxtAddress.Text = "" Then
   MsgBox "Blank Address not Allowed!!"
   TxtAddress.SetFocus
   Exit Sub
End If

If TxtFromNo.Text = "" Then
   MsgBox "Blank From Lock Seal Hologram No. not Allowed!!"
   TxtFromNo.SetFocus
   Exit Sub
End If

If TxtToNo.Text = "" Then
   MsgBox "Blank To Lock Seal Hologram No. not Allowed!!"
   TxtToNo.SetFocus
   Exit Sub
End If

TxtTotalQty.Text = Val(TxtToNo.Text) - Val(TxtFromNo.Text) + 1
If Val(TxtTotalQty.Text) <= 0 Then
   MsgBox "'To LockSeal Hologram No' must be greater than 'From LockSeal Hologram No'!!!"
   TxtToNo.SetFocus
   Exit Sub
End If

If TxtRemark.Text = "" Then
   MsgBox "Blank Remark not Allowed!!"
   TxtRemark.SetFocus
   Exit Sub
End If

'Call CheckHologramRange

If Edit_Flg = "A" Then
   RsTxn_LockSealHologram.AddNew
   RsTxn_LockSealHologram!LockSealHologramID = getMaxLockSealHologramID
   RsTxn_LockSealHologram!G_UID = GetGUID
   TxtLockSealHologramID.Text = RsTxn_LockSealHologram!LockSealHologramID
Else
   RsTxn_LockSealHologram.MoveFirst
   RsTxn_LockSealHologram.Find " LockSealHologramID= '" & Val(TxtLockSealHologramID.Text) & "'"
End If

RsTxn_LockSealHologram!LockSealHologramID = Val(TxtLockSealHologramID.Text)
RsTxn_LockSealHologram!LockSealHologramdate = Format(TxtLockSealHologramDate.Value, Gen_DatFormat)
RsTxn_LockSealHologram!LocationID = mLocationID
RsTxn_LockSealHologram!EmployeeID = mEmployeeID
RsTxn_LockSealHologram!Address = TxtAddress.Text
RsTxn_LockSealHologram!Remark = TxtRemark.Text
RsTxn_LockSealHologram!FromLockSealHologramNo = Val(TxtFromNo.Text)
RsTxn_LockSealHologram!ToLockSealHologramNo = Val(TxtToNo.Text)
RsTxn_LockSealHologram!TotalQty = Val(TxtTotalQty.Text)

If IsNull(RsTxn_LockSealHologram!CreatedBy) = True Or RsTxn_LockSealHologram!CreatedBy = "" Then
   RsTxn_LockSealHologram!CreatedBy = Gen_UserLoginID
   RsTxn_LockSealHologram!CreatedDate = Now
Else
   RsTxn_LockSealHologram!UpdatedBy = Gen_UserLoginID
   RsTxn_LockSealHologram!UpdatedDate = Now
End If

RsTxn_LockSealHologram.Update
Call SaveHologramDetails

SaveCmd.Enabled = False

Call Indata

End Sub
Private Function getMaxLockSealHologramID() As Double
Dim Retval As Double
tmp = "Select max(LockSealHologramID) from Txn_LockSealHologram"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
getMaxLockSealHologramID = Retval
End Function
Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsTxn_LockSealHologram.MoveNext
If RsTxn_LockSealHologram.EOF Then
   RsTxn_LockSealHologram.MoveLast
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
RsTxn_LockSealHologram.MovePrevious
If RsTxn_LockSealHologram.BOF Then
   RsTxn_LockSealHologram.MoveFirst
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
RsTxn_LockSealHologram.MoveLast
Call Indata
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False
End Sub
Private Sub FirstCmd_Click()
RsTxn_LockSealHologram.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub

Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   If RsTxn_LockSealHologram.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   RsTxn_LockSealHologram.MoveFirst
   RsTxn_LockSealHologram.Find "LockSealHologramID = " & Val(TxtLockSealHologramID) & ""
   If RsTxn_LockSealHologram.EOF Then
      RsTxn_LockSealHologram.MoveLast
   End If
   Call Indata
   Exit Sub
End If
Frame0.Height = Frame4.Height - 200
Frame0.Left = Frame4.Left - 100
Frame0.Top = Frame4.Top
Frame0.Width = Frame4.Width - 100
Frame0.Visible = True

MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 2500

RsTxn_LockSealHologram.MoveFirst
RsTxn_LockSealHologram.Filter = ""

''===========comment on 7 july 2012 for not clearing search grid=============
'Call Grid_Head
'SFrom.Value = Date - 365
'STo.Value = Date
'Label22.Caption = ""
'CmbSStateID.Text = ""
'CmbSEmployeeID.Text = ""
'CmbSLocationID.Text = ""
'TxtLockSealHologramNo.Text = ""
''===========comment on 7 july 2012 for not clearing search grid=============

End Sub

Private Sub TxtFromNo_LostFocus()
If TxtFromNo = "" Then Exit Sub
If Not IsNumeric(TxtFromNo) Or Val(TxtFromNo) < 0 Then
   MsgBox "Invalid numeric format !!!! "
   TxtFromNo.SetFocus
End If
If TxtToNo = "" Then Exit Sub
TxtTotalQty.Text = Val(TxtToNo.Text) - Val(TxtFromNo.Text) + 1
Call CheckHologramRange
Call Grid_Head1
End Sub


Private Sub TxtToNo_LostFocus()
If TxtToNo = "" Then Exit Sub
If Not IsNumeric(TxtToNo) Or Val(TxtToNo) < 0 Then
   MsgBox "Invalid numeric format !!!! "
   TxtToNo.SetFocus
End If
TxtTotalQty.Text = Val(TxtToNo.Text) - Val(TxtFromNo.Text) + 1
If Val(TxtTotalQty.Text) <= 0 Then
   MsgBox "'To LockSeal Hologram No' must be greater than 'From LockSeal Hologram No'!!!"
   TxtToNo.SetFocus
   Exit Sub
End If
Call CheckHologramRange
Call Grid_Head1
End Sub

Public Sub Indata()

Frame1.Enabled = False
Frame2.Enabled = False
Call ClearDetail
TxtLockSealHologramDate.Enabled = True
CmbLocationID.Visible = False
TxtLocationID.Visible = True

CmbStateID.Visible = False
TxtStateID.Visible = True

CmbEmployeeID.Visible = False
TxtEmployeeID.Visible = True

TxtLockSealHologramID = RsTxn_LockSealHologram!LockSealHologramID
TxtLockSealHologramDate.Value = RsTxn_LockSealHologram!LockSealHologramdate
TxtStateID = GetStateFromLocation(RsTxn_LockSealHologram!LocationID)
mStateID = GetStateID(TxtStateID)
mLocationID = RsTxn_LockSealHologram!LocationID
TxtLocationID.Text = GetLocationName(RsTxn_LockSealHologram!LocationID)
mEmployeeID = RsTxn_LockSealHologram!EmployeeID
TxtEmployeeID.Text = GetEmployeeNonName(RsTxn_LockSealHologram!EmployeeID)
TxtAddress.Text = RsTxn_LockSealHologram!Address
TxtRemark.Text = RsTxn_LockSealHologram!Remark
TxtFromNo.Text = RsTxn_LockSealHologram!FromLockSealHologramNo
TxtToNo.Text = RsTxn_LockSealHologram!ToLockSealHologramNo
TxtTotalQty.Text = RsTxn_LockSealHologram!TotalQty
TxtCreated = IIf(IsNull(RsTxn_LockSealHologram!CreatedBy), "", RsTxn_LockSealHologram!CreatedBy) & " :- " & IIf(IsNull(RsTxn_LockSealHologram!CreatedDate), "", RsTxn_LockSealHologram!CreatedDate)
TxtUpdated = IIf(IsNull(RsTxn_LockSealHologram!UpdatedBy), "", RsTxn_LockSealHologram!UpdatedBy) & " :- " & IIf(IsNull(RsTxn_LockSealHologram!UpdatedDate), "", RsTxn_LockSealHologram!UpdatedDate)

Call InDataHologramDetails
Call InDataLogHistory
If RsTxn_LockSealHologram.RecordCount = 1 Then
   Call GButtonLock_1(Me, True)
Else
   Call GButtonLock(Me, True)
End If
End Sub

Private Sub Grid_Head()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.FixedRows = 1
MSHFlexGrid1.Cols = 10
MSHFlexGrid1.WordWrap = True
MSHFlexGrid1.TextMatrix(0, 0) = "LockSeal Hologram ID"

MSHFlexGrid1.TextMatrix(0, 1) = "LockSeal Hologram Date"
MSHFlexGrid1.TextMatrix(0, 2) = "State"
MSHFlexGrid1.TextMatrix(0, 3) = "Location"
MSHFlexGrid1.TextMatrix(0, 4) = "Employee"
MSHFlexGrid1.TextMatrix(0, 5) = "Address"
MSHFlexGrid1.TextMatrix(0, 6) = "From LockSeal Hologram No."
MSHFlexGrid1.TextMatrix(0, 7) = "To LockSeal Hologram No."
MSHFlexGrid1.TextMatrix(0, 8) = "Total Qty"
MSHFlexGrid1.TextMatrix(0, 9) = "Remark"

MSHFlexGrid1.ColWidth(0) = 1000
MSHFlexGrid1.ColWidth(1) = 1400
MSHFlexGrid1.ColWidth(2) = 2400
MSHFlexGrid1.ColWidth(3) = 2400
MSHFlexGrid1.ColWidth(4) = 2400
MSHFlexGrid1.ColWidth(5) = 2000
MSHFlexGrid1.ColWidth(6) = 2000
MSHFlexGrid1.ColWidth(7) = 2000
MSHFlexGrid1.ColWidth(8) = 1400
MSHFlexGrid1.ColWidth(9) = 2800

MSHFlexGrid1.RowHeight(0) = 800

End Sub


Private Sub CmbSLocationID_GotFocus()
If CmbSStateID.Text = "" Then
   MsgBox " Please Select State!!!"
   CmbSStateID.SetFocus
   Exit Sub
End If
tmp = CmbSLocationID.Text

If Gen_WcLocation <> "" Then
   Call TableMst_Location(" Where StateID=" & mSStateID & " And " & Gen_WcLocation)
Else
   Call TableMst_Location(" Where StateID=" & mSStateID & " ") ' & Gen_WcLocation)
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


mSLocationID = RsMst_Location!LocationID

End Sub

Private Sub CmbStateID_GotFocus()
tmp = CmbStateID.Text
If Gen_WcState <> "" Then
   Call TableMst_State(" Where " & Gen_WcState)
Else
   Call TableMst_State '(" Where LocationID <> 0 ")
End If
CmbStateID.Clear
If RsMst_State.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_State.RecordCount
    CmbStateID.AddItem RsMst_State!StateName
    RsMst_State.MoveNext
Next
CmbStateID.Text = tmp
End Sub

Private Sub CmbStateID_LostFocus()
If CmbStateID.Text = "" Then
   mStateID = 0
   Exit Sub
End If
RsMst_State.MoveFirst
RsMst_State.Find "StateName = '" & CmbStateID.Text & "'"
If RsMst_State.EOF Then
   MsgBox "Invalid selection "
   CmbStateID.SetFocus
   Exit Sub
End If
If mStateID <> RsMst_State!StateID Then
   CmbLocationID.Clear
   mLocationID = 0
End If
mStateID = RsMst_State!StateID
End Sub

Private Sub CmbSStateID_GotFocus()
tmp = CmbSStateID.Text
If Gen_WcState <> "" Then
   Call TableMst_State(" Where " & Gen_WcState)
Else
   Call TableMst_State '(" Where LocationID <> 0 ")
End If
CmbSStateID.Clear
If RsMst_State.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_State.RecordCount
    CmbSStateID.AddItem RsMst_State!StateName
    RsMst_State.MoveNext
Next
CmbSStateID.Text = tmp
End Sub

Private Sub CmbSStateID_LostFocus()
If CmbSStateID.Text = "" Then
   mSStateID = 0
   Exit Sub
End If
RsMst_State.MoveFirst
RsMst_State.Find "StateName = '" & CmbSStateID.Text & "'"
If RsMst_State.EOF Then
   MsgBox "Invalid selection "
   CmbSStateID.SetFocus
   Exit Sub
End If
If mSStateID <> RsMst_State!StateID Then
   CmbSLocationID.Clear
   mSLocationID = 0
End If
mSStateID = RsMst_State!StateID
End Sub

Private Sub Grid_Head1()

MSHFlexGrid2.Clear
MSHFlexGrid2.Rows = 2
MSHFlexGrid2.FixedRows = 1
MSHFlexGrid2.Cols = 9
MSHFlexGrid2.WordWrap = True
MSHFlexGrid2.TextMatrix(0, 0) = "Date"
MSHFlexGrid2.TextMatrix(0, 1) = "Location"
MSHFlexGrid2.TextMatrix(0, 2) = "LocationID"
MSHFlexGrid2.TextMatrix(0, 3) = "From Lock Seal Hologram No"
MSHFlexGrid2.TextMatrix(0, 4) = "To Lock Seal Hologram No"
MSHFlexGrid2.TextMatrix(0, 5) = "Total"
MSHFlexGrid2.TextMatrix(0, 6) = "Employee"
MSHFlexGrid2.TextMatrix(0, 7) = "EmployeeID"
MSHFlexGrid2.TextMatrix(0, 8) = "Address"

MSHFlexGrid2.ColWidth(0) = 1500
MSHFlexGrid2.ColWidth(1) = 2400
MSHFlexGrid2.ColWidth(2) = 0
MSHFlexGrid2.ColWidth(3) = 1700
MSHFlexGrid2.ColWidth(4) = 1700
MSHFlexGrid2.ColWidth(5) = 1300
MSHFlexGrid2.ColWidth(6) = 2800
MSHFlexGrid2.ColWidth(7) = 0
MSHFlexGrid2.ColWidth(8) = 2500

MSHFlexGrid2.RowHeight(0) = 800

End Sub

Private Sub CmbILocationID_GotFocus()

tmp = CmbILocationID.Text

Call TableMst_Location("") ' Where StateID=" & mStateID)

CmbILocationID.Clear

If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbILocationID.AddItem RsMst_Location!LocationName
    RsMst_Location.MoveNext
Next

CmbILocationID.Text = tmp

End Sub

Private Sub CmbILocationID_LostFocus()

If CmbILocationID.Text = "" Then
   mILocationID = 0
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbILocationID.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection "
   CmbILocationID.SetFocus
   Exit Sub
End If

mILocationID = RsMst_Location!LocationID

End Sub

Private Sub CmbIEmployeeID_GotFocus()
tmp = CmbIEmployeeID.Text
Call TableMst_Employee("Where Flag = 'Y'")
CmbIEmployeeID.Clear
If RsMst_Employee.RecordCount = 0 Then
   MsgBox "No Employee Found!!!"
   CmbIEmployeeID.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Employee.RecordCount
    CmbIEmployeeID.AddItem RsMst_Employee!EmployeeName & "~" & RsMst_Employee!EmployeeNo
    RsMst_Employee.MoveNext
Next
CmbIEmployeeID.Text = tmp
End Sub

Private Sub CmbIEmployeeID_LostFocus()
Dim strSplitString() As String
Dim EmpNo_, EmpName_ As String

If CmbIEmployeeID.Text = "" Then
   mIEmployeeID = Null
   Exit Sub
End If

strSplitString = Split(CmbIEmployeeID, "~")

If UBound(strSplitString) <> 1 Then
   MsgBox "Invalid Selection "
   CmbIEmployeeID.SetFocus
   Exit Sub
End If

EmpName_ = Trim(strSplitString(0))
EmpNo_ = Trim(strSplitString(1))


RsMst_Employee.MoveFirst
RsMst_Employee.Find "EmployeeNo = '" & EmpNo_ & "'"

If RsMst_Employee.EOF Then
   MsgBox "Invalid Selection "
   CmbIEmployeeID.SetFocus
   Exit Sub
End If
mIEmployeeID = RsMst_Employee!EmployeeID
End Sub

Private Sub CheckHologramRange()
tmp = "Select LockSealHologramID,LockSealHologramdate,LocationID,FromLockSealHologramNo,ToLockSealHologramNo From Txn_LockSealHologram"
''tmp = tmp & " Where LocationID = " & mLocationID
Call Tmp1Table
If TmpRs1.RecordCount > 0 Then
   For I = 0 To TmpRs1.RecordCount - 1
       If Edit_Flg = "E" Then
          If Val(TxtLockSealHologramID.Text) <> Val(TmpRs1!LockSealHologramID) Then
             If Val(TxtFromNo.Text) >= Val(TmpRs1!FromLockSealHologramNo) And Val(TxtFromNo.Text) <= Val(TmpRs1!ToLockSealHologramNo) Then
                MsgBox "From LockSeal Hologram No. Already in " & GetLocationName(TmpRs1!LocationID) & ", Range " & TmpRs1!FromLockSealHologramNo & " To " & TmpRs1!ToLockSealHologramNo
                TxtFromNo.SetFocus
                Exit Sub
             End If
             If Val(TxtFromNo.Text) <= Val(TmpRs1!FromLockSealHologramNo) And Val(TxtToNo.Text) >= Val(TmpRs1!ToLockSealHologramNo) Then
                MsgBox "LockSeal Hologram No. already in " & GetLocationName(TmpRs1!LocationID) & ", Range " & TmpRs1!FromLockSealHologramNo & " To " & TmpRs1!ToLockSealHologramNo
                TxtFromNo.SetFocus
                Exit Sub
             End If
             If Val(TxtToNo.Text) >= Val(TmpRs1!FromLockSealHologramNo) And Val(TxtToNo.Text) <= Val(TmpRs1!ToLockSealHologramNo) Then
                MsgBox "To LockSeal Hologram No. already in " & GetLocationName(TmpRs1!LocationID) & ", Range " & TmpRs1!FromLockSealHologramNo & " To " & TmpRs1!ToLockSealHologramNo
                TxtToNo.SetFocus
                Exit Sub
             End If
          End If
       Else
          If Val(TxtFromNo.Text) >= Val(TmpRs1!FromLockSealHologramNo) And Val(TxtFromNo.Text) <= Val(TmpRs1!ToLockSealHologramNo) Then
             MsgBox "From LockSeal Hologram No. already in " & GetLocationName(TmpRs1!LocationID) & ", Range " & TmpRs1!FromLockSealHologramNo & " To " & TmpRs1!ToLockSealHologramNo
             TxtFromNo.SetFocus
             Exit Sub
          End If
          If Val(TxtFromNo.Text) <= Val(TmpRs1!FromLockSealHologramNo) And Val(TxtToNo.Text) >= Val(TmpRs1!ToLockSealHologramNo) Then
                MsgBox "LockSeal Hologram No. already  in " & GetLocationName(TmpRs1!LocationID) & ", Range " & TmpRs1!FromLockSealHologramNo & " To " & TmpRs1!ToLockSealHologramNo
                TxtFromNo.SetFocus
                Exit Sub
          End If
          If Val(TxtToNo.Text) >= Val(TmpRs1!FromLockSealHologramNo) And Val(TxtToNo.Text) <= Val(TmpRs1!ToLockSealHologramNo) Then
             MsgBox "To LockSeal Hologram No. already in " & GetLocationName(TmpRs1!LocationID) & ", Range " & TmpRs1!FromLockSealHologramNo & " To " & TmpRs1!ToLockSealHologramNo
             TxtToNo.SetFocus
             Exit Sub
          End If
 
       End If
       TmpRs1.MoveNext
   Next
End If
End Sub


Function CheckDuplicateHologramNo() As Boolean
Dim Retval As Boolean
Retval = False
If MSHFlexGrid2.Rows > 1 Then
   For u = 1 To MSHFlexGrid2.Rows - 1
       If (Val(TxtIFromNo) >= Val(MSHFlexGrid2.TextMatrix(u, 3)) And Val(TxtIFromNo) <= Val(MSHFlexGrid2.TextMatrix(u, 4))) Or Val(TxtIToNo) >= Val(MSHFlexGrid2.TextMatrix(u, 3)) And Val(TxtIToNo) <= Val(MSHFlexGrid2.TextMatrix(u, 4)) Or (Val(TxtIFromNo) <= Val(MSHFlexGrid2.TextMatrix(u, 3)) And Val(TxtIToNo) >= Val(MSHFlexGrid2.TextMatrix(u, 4))) Then
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

CheckDuplicateHologramNo = Retval

End Function

Private Sub ClearDetail()
TxtIDate.Value = Date
CmbILocationID.Text = ""
TxtIFromNo = ""
TxtIToNo = ""
TxtITotalQty = ""
CmbIEmployeeID.Text = ""
TxtIAddress = ""
mILocationID = 0
mIEmployeeID = 0
CmdAdd2List.Caption = "Add to List"
End Sub

Private Sub TxtIToNo_LostFocus()
If TxtIToNo = "" Then Exit Sub
If Not IsNumeric(TxtIToNo) Or Val(TxtIToNo) < 0 Then
   MsgBox "Invalid numeric format !!!! "
   TxtIToNo.SetFocus
End If
TxtITotalQty.Text = Val(TxtIToNo.Text) - Val(TxtIFromNo.Text) + 1
End Sub

Private Sub TxtIFromNo_LostFocus()
If TxtIFromNo = "" Then Exit Sub
If Not IsNumeric(TxtIFromNo) Or Val(TxtIFromNo) < 0 Then
   MsgBox "Invalid numeric format !!!! "
   TxtIFromNo.SetFocus
End If
If TxtIToNo = "" Then Exit Sub
TxtITotalQty.Text = Val(TxtIToNo.Text) - Val(TxtIFromNo.Text) + 1

End Sub

Private Sub SaveHologramDetails()

Call DeleteHologramDetails

Call TableTxn_LockSealHologramDetail("")
For I = 1 To MSHFlexGrid2.Rows - 1
    If MSHFlexGrid2.TextMatrix(I, 0) <> "" Then
       RsTxn_LockSealHologramDetail.AddNew
       RsTxn_LockSealHologramDetail!LSHDID = GetMaxDetailID
       RsTxn_LockSealHologramDetail!IssueDate = CDate(MSHFlexGrid2.TextMatrix(I, 0))
       RsTxn_LockSealHologramDetail!LockSealHologramID = TxtLockSealHologramID
       RsTxn_LockSealHologramDetail!LocationID = MSHFlexGrid2.TextMatrix(I, 2)
       RsTxn_LockSealHologramDetail!FromLSHologramNo = Val(MSHFlexGrid2.TextMatrix(I, 3))
       RsTxn_LockSealHologramDetail!ToLSHologramNo = Val(MSHFlexGrid2.TextMatrix(I, 4))
       RsTxn_LockSealHologramDetail!EmployeeID = Val(MSHFlexGrid2.TextMatrix(I, 7))
       RsTxn_LockSealHologramDetail!Address = MSHFlexGrid2.TextMatrix(I, 8)
       RsTxn_LockSealHologramDetail!G_UID = GetGUID
       RsTxn_LockSealHologramDetail.Update
       RsTxn_LockSealHologramDetail.Requery
    End If
Next
End Sub

Private Sub DeleteHologramDetails()
tmp = "Delete From Txn_LockSealHologramDetail Where LockSealHologramID =" & TxtLockSealHologramID & " "
Call Tmp4Table
End Sub
Private Function GetMaxDetailID() As Double
Dim Retval As Double
tmp = "Select max(LSHDID) from Txn_LockSealHologramDetail "
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetMaxDetailID = Retval
End Function

Private Sub InDataHologramDetails()

Call Grid_Head1
Call TableTxn_LockSealHologramDetail(" Where LockSealHologramID= " & TxtLockSealHologramID & " ")
If RsTxn_LockSealHologramDetail.RecordCount > 0 Then
    MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + RsTxn_LockSealHologramDetail.RecordCount
    For I = 1 To MSHFlexGrid2.Rows - 2
        MSHFlexGrid2.TextMatrix(I, 0) = Format(RsTxn_LockSealHologramDetail!IssueDate, "dd-MMM-yyyy")
        MSHFlexGrid2.TextMatrix(I, 1) = GetLocationName(RsTxn_LockSealHologramDetail!LocationID)
        MSHFlexGrid2.TextMatrix(I, 2) = RsTxn_LockSealHologramDetail!LocationID
        MSHFlexGrid2.TextMatrix(I, 3) = RsTxn_LockSealHologramDetail!FromLSHologramNo
        MSHFlexGrid2.TextMatrix(I, 4) = RsTxn_LockSealHologramDetail!ToLSHologramNo
        MSHFlexGrid2.TextMatrix(I, 5) = RsTxn_LockSealHologramDetail!ToLSHologramNo - RsTxn_LockSealHologramDetail!FromLSHologramNo
        MSHFlexGrid2.TextMatrix(I, 6) = GetEmployeeNonName(RsTxn_LockSealHologramDetail!EmployeeID)
        MSHFlexGrid2.TextMatrix(I, 7) = RsTxn_LockSealHologramDetail!EmployeeID
        MSHFlexGrid2.TextMatrix(I, 8) = RsTxn_LockSealHologramDetail!Address
        RsTxn_LockSealHologramDetail.MoveNext
    Next
End If
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

Private Sub CmbSEmployeeID_GotFocus()
tmp = CmbSEmployeeID.Text
Call TableMst_Employee("Where Flag = 'Y'")
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
   mSEmployeeID = 0
   Exit Sub
End If

strSplitString = Split(CmbSEmployeeID, "~")

If UBound(strSplitString) <> 1 Then
   MsgBox "Invalid Selection "
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

Private Sub Grid_Head3()

MSHFlexGrid4.Clear
MSHFlexGrid4.Rows = 2
MSHFlexGrid4.FixedRows = 1
MSHFlexGrid4.Cols = 3

MSHFlexGrid4.TextMatrix(0, 0) = "Lock Seal Hologram ID"
MSHFlexGrid4.TextMatrix(0, 1) = "Updated By"
MSHFlexGrid4.TextMatrix(0, 2) = "Updated On"

MSHFlexGrid4.ColWidth(0) = 2200
MSHFlexGrid4.ColWidth(1) = 2000
MSHFlexGrid4.ColWidth(2) = 3000

End Sub

Private Sub InDataLogHistory()

Call Grid_Head3
tmp = "Select LockSealHologramID, UpdatedBy, UpdatedOn from Log_LockSealHologramHistory Where LockSealHologramID=" & Val(TxtLockSealHologramID) & ""
Call TmpTable
MSHFlexGrid4.Rows = MSHFlexGrid4.Rows + TmpRs.RecordCount
For I = 1 To MSHFlexGrid4.Rows - 2
   MSHFlexGrid4.TextMatrix(I, 0) = TmpRs!LockSealHologramID
   MSHFlexGrid4.TextMatrix(I, 1) = IIf(IsNull(TmpRs!UpdatedBy), "", TmpRs!UpdatedBy)
   MSHFlexGrid4.TextMatrix(I, 2) = TmpRs!UpdatedOn
   TmpRs.MoveNext
Next
End Sub
