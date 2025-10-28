VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmTxn_Fumigation 
   Caption         =   "Fumigation"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      TabIndex        =   97
      Top             =   9690
      Width           =   10935
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   9015
      Left            =   15240
      TabIndex        =   96
      Top             =   240
      Width           =   255
   End
   Begin VB.Frame Frame4 
      Height          =   9735
      Left            =   15
      TabIndex        =   0
      Top             =   -120
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
         Height          =   2445
         Left            =   4440
         TabIndex        =   22
         Top             =   5880
         Visible         =   0   'False
         Width           =   14940
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
            Left            =   9240
            MaxLength       =   12
            TabIndex        =   126
            Top             =   1425
            Width           =   1455
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
            ItemData        =   "FrmTxn_Fumigation.frx":0000
            Left            =   7680
            List            =   "FrmTxn_Fumigation.frx":000A
            Sorted          =   -1  'True
            TabIndex        =   124
            Top             =   1425
            Width           =   1485
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
            Left            =   14280
            Sorted          =   -1  'True
            TabIndex        =   108
            Top             =   960
            Visible         =   0   'False
            Width           =   1410
         End
         Begin VB.CheckBox ChkBatchNo 
            Appearance      =   0  'Flat
            Caption         =   "Check1"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   15450
            TabIndex        =   106
            Top             =   930
            Visible         =   0   'False
            Width           =   225
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
            ItemData        =   "FrmTxn_Fumigation.frx":0017
            Left            =   12480
            List            =   "FrmTxn_Fumigation.frx":0024
            Sorted          =   -1  'True
            TabIndex        =   104
            Top             =   555
            Visible         =   0   'False
            Width           =   1845
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
            Left            =   6120
            MaxLength       =   12
            TabIndex        =   30
            Top             =   1425
            Width           =   1455
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
            TabIndex        =   29
            Top             =   1440
            Width           =   5925
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
            Left            =   6240
            Sorted          =   -1  'True
            TabIndex        =   28
            Top             =   570
            Width           =   3645
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
            Left            =   9960
            Sorted          =   -1  'True
            TabIndex        =   27
            Top             =   555
            Width           =   2445
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
            ItemData        =   "FrmTxn_Fumigation.frx":0045
            Left            =   3210
            List            =   "FrmTxn_Fumigation.frx":0047
            Sorted          =   -1  'True
            TabIndex        =   26
            Top             =   570
            Width           =   3000
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
            Left            =   13845
            TabIndex        =   25
            Top             =   1905
            Width           =   1000
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
            Left            =   12825
            TabIndex        =   24
            Top             =   1905
            Width           =   1000
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
            Left            =   13320
            TabIndex        =   23
            Top             =   585
            Visible         =   0   'False
            Width           =   1890
         End
         Begin MSComCtl2.DTPicker STo 
            Height          =   375
            Left            =   1695
            TabIndex        =   31
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
            Format          =   120127489
            CurrentDate     =   36494
         End
         Begin MSComCtl2.DTPicker SFrom 
            Height          =   375
            Left            =   135
            TabIndex        =   32
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
            Format          =   120127489
            CurrentDate     =   36494
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
            Height          =   1455
            Left            =   120
            TabIndex        =   33
            Top             =   2280
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
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            Caption         =   "Billable Qty"
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
            Left            =   9457
            TabIndex        =   127
            Top             =   1080
            Width           =   1020
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
            Left            =   7785
            TabIndex        =   125
            Top             =   1080
            Width           =   1275
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
            Left            =   13425
            TabIndex        =   109
            Top             =   1080
            Visible         =   0   'False
            Width           =   990
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
            Left            =   13710
            TabIndex        =   107
            Top             =   930
            Visible         =   0   'False
            Width           =   1710
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
            Left            =   13020
            TabIndex        =   105
            Top             =   255
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Fumigation ID"
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
            Left            =   6225
            TabIndex        =   43
            Top             =   1080
            Width           =   1230
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
            Left            =   2550
            TabIndex        =   42
            Top             =   1110
            Width           =   1320
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
            Left            =   7755
            TabIndex        =   41
            Top             =   262
            Width           =   435
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
            Left            =   10807
            TabIndex        =   40
            Top             =   262
            Width           =   750
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
            Left            =   4320
            TabIndex        =   39
            Top             =   262
            Width           =   765
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
            Left            =   5760
            TabIndex        =   38
            Top             =   1890
            Width           =   90
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
            Left            =   3240
            TabIndex        =   37
            Top             =   1890
            Width           =   2415
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
            TabIndex        =   36
            Top             =   255
            Width           =   855
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
            TabIndex        =   35
            Top             =   255
            Width           =   855
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
            Left            =   14040
            TabIndex        =   34
            Top             =   240
            Visible         =   0   'False
            Width           =   585
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
         Height          =   3930
         Left            =   120
         TabIndex        =   67
         Top             =   120
         Width           =   14970
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
            Left            =   6990
            TabIndex        =   123
            TabStop         =   0   'False
            Top             =   1110
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
            Left            =   6990
            TabIndex        =   122
            TabStop         =   0   'False
            Top             =   705
            Width           =   375
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
            Left            =   9495
            Locked          =   -1  'True
            TabIndex        =   114
            TabStop         =   0   'False
            Top             =   2925
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
            Left            =   5160
            Locked          =   -1  'True
            TabIndex        =   112
            TabStop         =   0   'False
            Top             =   2925
            Width           =   1800
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
            Left            =   1575
            Locked          =   -1  'True
            TabIndex        =   110
            TabStop         =   0   'False
            Top             =   2925
            Width           =   1680
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
            Height          =   780
            Left            =   9000
            MaxLength       =   250
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   12
            Top             =   4350
            Visible         =   0   'False
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
            Left            =   1590
            Locked          =   -1  'True
            TabIndex        =   74
            TabStop         =   0   'False
            Top             =   3375
            Width           =   5400
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
            Left            =   9495
            Locked          =   -1  'True
            TabIndex        =   73
            TabStop         =   0   'False
            Top             =   3375
            Width           =   5040
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
            Left            =   1575
            MaxLength       =   250
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   8
            Top             =   2310
            Width           =   12960
         End
         Begin VB.TextBox TxtFumigationID 
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
            TabIndex        =   72
            TabStop         =   0   'False
            Top             =   247
            Width           =   1635
         End
         Begin VB.TextBox TxtRate 
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
            Left            =   1575
            MaxLength       =   12
            TabIndex        =   6
            Top             =   1485
            Visible         =   0   'False
            Width           =   915
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
            Left            =   5160
            Locked          =   -1  'True
            MaxLength       =   12
            TabIndex        =   71
            TabStop         =   0   'False
            Top             =   1485
            Width           =   1800
         End
         Begin VB.TextBox TxtBillableQty 
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
            TabIndex        =   70
            TabStop         =   0   'False
            Top             =   1485
            Width           =   1635
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
            Left            =   9480
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   69
            TabStop         =   0   'False
            Top             =   1485
            Width           =   5040
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
            Left            =   1575
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   1890
            Width           =   5355
         End
         Begin MSComCtl2.DTPicker TxtFumDate 
            Height          =   375
            Left            =   5400
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
            Format          =   120127489
            CurrentDate     =   39884
         End
         Begin MSComCtl2.DTPicker TxtNextFumDate 
            Height          =   375
            Left            =   9480
            TabIndex        =   78
            TabStop         =   0   'False
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
            Format          =   120127489
            CurrentDate     =   39884
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
            Left            =   1575
            Sorted          =   -1  'True
            TabIndex        =   4
            Top             =   1080
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
            Left            =   9480
            Sorted          =   -1  'True
            TabIndex        =   5
            Top             =   1080
            Width           =   5040
         End
         Begin MSComCtl2.DTPicker TxtFAGPostingDate 
            Height          =   375
            Left            =   4920
            TabIndex        =   11
            Top             =   4755
            Visible         =   0   'False
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
            Format          =   120127489
            CurrentDate     =   39884
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
            ItemData        =   "FrmTxn_Fumigation.frx":0049
            Left            =   1095
            List            =   "FrmTxn_Fumigation.frx":0056
            Sorted          =   -1  'True
            TabIndex        =   10
            Top             =   4755
            Visible         =   0   'False
            Width           =   1560
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
            Left            =   1110
            Locked          =   -1  'True
            TabIndex        =   98
            Top             =   4755
            Visible         =   0   'False
            Width           =   1530
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
            Left            =   1095
            Sorted          =   -1  'True
            TabIndex        =   9
            Top             =   4350
            Visible         =   0   'False
            Width           =   5400
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
            Left            =   1095
            Locked          =   -1  'True
            TabIndex        =   102
            Top             =   4350
            Visible         =   0   'False
            Width           =   5370
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
            Left            =   1575
            Locked          =   -1  'True
            TabIndex        =   76
            Top             =   1080
            Visible         =   0   'False
            Width           =   5370
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
            Left            =   9480
            Locked          =   -1  'True
            TabIndex        =   77
            Top             =   1080
            Visible         =   0   'False
            Width           =   5040
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
            Width           =   5400
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
            TabIndex        =   75
            Top             =   675
            Visible         =   0   'False
            Width           =   5400
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
            Left            =   14550
            TabIndex        =   68
            TabStop         =   0   'False
            Top             =   705
            Width           =   375
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
            Left            =   9480
            Sorted          =   -1  'True
            TabIndex        =   3
            Top             =   675
            Width           =   5040
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
            Left            =   9480
            Locked          =   -1  'True
            TabIndex        =   79
            Top             =   675
            Visible         =   0   'False
            Width           =   5040
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
            Left            =   7455
            TabIndex        =   115
            Top             =   3015
            Width           =   1035
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
            Left            =   3600
            TabIndex        =   113
            Top             =   3000
            Width           =   1425
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
            Left            =   240
            TabIndex        =   111
            Top             =   2985
            Width           =   1140
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
            Left            =   6960
            TabIndex        =   101
            Top             =   4620
            Visible         =   0   'False
            Width           =   1275
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
            Left            =   255
            TabIndex        =   94
            Top             =   3435
            Width           =   855
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
            Left            =   7455
            TabIndex        =   93
            Top             =   3435
            Width           =   885
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
            Left            =   240
            TabIndex        =   92
            Top             =   2460
            Width           =   825
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Fumigation ID"
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
            Left            =   240
            TabIndex        =   90
            Top             =   307
            Width           =   1230
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Fumigation Start Date"
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
            TabIndex        =   89
            Top             =   307
            Width           =   1920
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Next Fumigation Date"
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
            Left            =   7455
            TabIndex        =   85
            Top             =   307
            Width           =   1920
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Rate per MT"
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
            Left            =   240
            TabIndex        =   84
            Top             =   1545
            Visible         =   0   'False
            Width           =   1305
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
            Left            =   4320
            TabIndex        =   83
            Top             =   1545
            Width           =   675
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Billable Qty"
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
            Left            =   240
            TabIndex        =   82
            Top             =   1545
            Width           =   1020
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
            Left            =   7455
            TabIndex        =   81
            Top             =   1770
            Width           =   1560
         End
         Begin VB.Label Label27 
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
            Left            =   240
            TabIndex        =   80
            Top             =   1950
            Width           =   585
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
            Left            =   7455
            TabIndex        =   88
            Top             =   1140
            Width           =   855
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
            Left            =   240
            TabIndex        =   87
            Top             =   1140
            Width           =   495
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
            Left            =   -240
            TabIndex        =   103
            Top             =   4410
            Visible         =   0   'False
            Width           =   1140
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
            Height          =   300
            Left            =   2760
            TabIndex        =   100
            Top             =   4785
            Visible         =   0   'False
            Width           =   2040
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
            Left            =   -240
            TabIndex        =   99
            Top             =   4815
            Visible         =   0   'False
            Width           =   660
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
            Left            =   240
            TabIndex        =   91
            Top             =   735
            Width           =   900
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
            Left            =   7455
            TabIndex        =   86
            Top             =   735
            Width           =   765
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Stack Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1590
         Left            =   120
         TabIndex        =   54
         Top             =   4125
         Width           =   14970
         Begin VB.TextBox TxtAmountStack 
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
            Left            =   11535
            Locked          =   -1  'True
            MaxLength       =   12
            TabIndex        =   120
            TabStop         =   0   'False
            Top             =   1080
            Width           =   1755
         End
         Begin VB.TextBox TxtRateStack 
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
            Left            =   8895
            Locked          =   -1  'True
            MaxLength       =   12
            TabIndex        =   118
            TabStop         =   0   'False
            Top             =   1080
            Width           =   1635
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
            Left            =   14520
            TabIndex        =   116
            TabStop         =   0   'False
            Top             =   630
            Width           =   375
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
            MaxLength       =   12
            TabIndex        =   58
            TabStop         =   0   'False
            Top             =   630
            Width           =   930
         End
         Begin VB.TextBox TxtCommodity 
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
            MaxLength       =   12
            TabIndex        =   57
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
            Left            =   3240
            MaxLength       =   12
            TabIndex        =   56
            TabStop         =   0   'False
            Top             =   630
            Width           =   1170
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
            Left            =   4425
            MaxLength       =   12
            TabIndex        =   55
            TabStop         =   0   'False
            Top             =   630
            Width           =   1170
         End
         Begin VB.TextBox TxtActualBags 
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
            Left            =   5625
            MaxLength       =   12
            TabIndex        =   13
            Top             =   630
            Width           =   1260
         End
         Begin VB.TextBox TxtActualQty 
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
            Left            =   6900
            MaxLength       =   12
            TabIndex        =   14
            Top             =   630
            Width           =   1290
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
            Left            =   13440
            TabIndex        =   18
            Top             =   1080
            Width           =   1335
         End
         Begin VB.TextBox TxtRemark1 
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
            Left            =   1080
            MaxLength       =   100
            TabIndex        =   17
            Top             =   1080
            Width           =   6375
         End
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
            ItemData        =   "FrmTxn_Fumigation.frx":0077
            Left            =   8220
            List            =   "FrmTxn_Fumigation.frx":0081
            Sorted          =   -1  'True
            TabIndex        =   15
            Top             =   630
            Width           =   1260
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
            Left            =   9600
            Sorted          =   -1  'True
            TabIndex        =   16
            Top             =   630
            Width           =   4815
         End
         Begin VB.Label Label41 
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
            Left            =   10680
            TabIndex        =   121
            Top             =   1140
            Width           =   675
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            Caption         =   "Rate per MT"
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
            Left            =   7560
            TabIndex        =   119
            Top             =   1140
            Width           =   1125
         End
         Begin VB.Label Label39 
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
            Left            =   11287
            TabIndex        =   117
            Top             =   300
            Width           =   1440
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
            Left            =   240
            TabIndex        =   66
            Top             =   285
            Width           =   1005
         End
         Begin VB.Label Label10 
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
            Left            =   1560
            TabIndex        =   65
            Top             =   285
            Width           =   1410
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
            Left            =   3345
            TabIndex        =   64
            Top             =   315
            Width           =   1050
         End
         Begin VB.Label Label12 
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
            Left            =   4470
            TabIndex        =   63
            Top             =   315
            Width           =   1215
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            Caption         =   "Actual No of Bags"
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
            Left            =   5700
            TabIndex        =   62
            Top             =   150
            Width           =   1155
         End
         Begin VB.Label Label14 
            Caption         =   "Actual  Qty"
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
            Left            =   7110
            TabIndex        =   61
            Top             =   300
            Width           =   1170
         End
         Begin VB.Label Label15 
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
            Height          =   255
            Left            =   8265
            TabIndex        =   60
            Top             =   300
            Width           =   1215
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Remark"
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
            Left            =   240
            TabIndex        =   59
            Top             =   1140
            Width           =   720
         End
      End
      Begin VB.Frame ButtonFrm 
         Height          =   1170
         Left            =   120
         TabIndex        =   44
         Top             =   8400
         Width           =   14970
         Begin VB.CommandButton AddCmd 
            Caption         =   "&Add New"
            Height          =   400
            Left            =   105
            TabIndex        =   53
            Top             =   195
            Width           =   2580
         End
         Begin VB.CommandButton NextCmd 
            Caption         =   "&Next"
            Height          =   400
            Left            =   105
            TabIndex        =   52
            Top             =   600
            Width           =   2580
         End
         Begin VB.CommandButton SaveCmd 
            Caption         =   "Save"
            Height          =   400
            Left            =   2685
            TabIndex        =   19
            Top             =   195
            Width           =   2580
         End
         Begin VB.CommandButton PreviousCmd 
            Caption         =   "&Previous"
            Height          =   400
            Left            =   2685
            TabIndex        =   51
            Top             =   600
            Width           =   2580
         End
         Begin VB.CommandButton EditCmd 
            Caption         =   "&Edit"
            Height          =   400
            Left            =   5265
            TabIndex        =   50
            Top             =   195
            Width           =   2460
         End
         Begin VB.CommandButton FirstCmd 
            Caption         =   "&First"
            Height          =   400
            Left            =   5265
            TabIndex        =   49
            Top             =   600
            Width           =   2460
         End
         Begin VB.CommandButton DeleteCmd 
            Caption         =   "&Delete"
            Height          =   400
            Left            =   7725
            TabIndex        =   48
            Top             =   195
            Width           =   2460
         End
         Begin VB.CommandButton LastCmd 
            Caption         =   "&Last"
            Height          =   400
            Left            =   7725
            TabIndex        =   47
            Top             =   600
            Width           =   2460
         End
         Begin VB.CommandButton SearchCmd 
            Caption         =   "Search"
            Height          =   400
            Left            =   10185
            TabIndex        =   21
            Top             =   195
            Width           =   2340
         End
         Begin VB.CommandButton ExitCmd 
            Caption         =   "E&xit"
            Height          =   400
            Left            =   10185
            TabIndex        =   20
            Top             =   600
            Width           =   2340
         End
         Begin VB.CommandButton CmdPrintPreview 
            Caption         =   "Preview FCR"
            Height          =   400
            Left            =   12525
            TabIndex        =   46
            Top             =   195
            Width           =   2340
         End
         Begin VB.CommandButton CmdPrint 
            Caption         =   "Print FCR"
            Height          =   400
            Left            =   12525
            TabIndex        =   45
            Top             =   600
            Width           =   2340
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   2625
         Left            =   120
         TabIndex        =   95
         Top             =   5775
         Width           =   14955
         _ExtentX        =   26379
         _ExtentY        =   4630
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
Attribute VB_Name = "FrmTxn_Fumigation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mVendorID, mGodownID, mLocationID, mCommodityID, mSVendorID, mSGodownID, mSLocationID, mSCommodityID As Variant
Dim mCMPID As Variant
Dim Edit_Flg, mSCMPID, mForwardToID, mForwardToEmailID, mSForwardToID As Variant
Dim mTreatmentTypeID, mAgreementDetailID As Variant

Public Sub Grid_Head()
With MSHFlexGrid1
    .Clear
    .Cols = 16
    .Rows = 3
    .FixedRows = 2
    .TextMatrix(1, 0) = "Stack No"
    .TextMatrix(1, 1) = "Commodity"
    .TextMatrix(1, 2) = "No. of Bags"
    .TextMatrix(1, 3) = "Balance Qty"
    .TextMatrix(1, 4) = "Actual No. of Bags"
    .TextMatrix(1, 5) = "Actual Qty"
    .TextMatrix(1, 6) = "Last Fumigation Date"
    .TextMatrix(1, 7) = "Effective Days"
    .TextMatrix(1, 8) = "Days Difference"
    .TextMatrix(1, 9) = "Billable Flag"
    .TextMatrix(1, 10) = "Remark"
    .TextMatrix(1, 11) = "CommodityID"
    .TextMatrix(1, 12) = "Treatment Type"
    .TextMatrix(1, 13) = "TreatmentID"
    .TextMatrix(1, 14) = "Rate per MT"
    .TextMatrix(1, 15) = "Amount"
    
    
    .WordWrap = True
    .ColWidth(0) = 1000
    .ColWidth(1) = 1800
    .ColWidth(2) = 1100
    .ColWidth(3) = 1100
    .ColWidth(4) = 1100
    .ColWidth(5) = 1100
    .ColWidth(6) = 1500
    .ColWidth(7) = 1100
    .ColWidth(8) = 1100
    .ColWidth(9) = 1000
    .ColWidth(10) = 1500
    .ColWidth(11) = 0
    .ColWidth(12) = 1500
    .ColWidth(13) = 0
    .RowHeight(1) = 600
    
End With
End Sub

Private Sub AddCmd_Click()

Edit_Flg = "A"
Call GButtonLock(Me, False)
Frame2.Enabled = True
Frame1.Enabled = True
TxtFumigationID.Text = ""
CmbVendorID.Text = ""
CmbLocationID.Text = ""
CmbCMPName.Text = ""
CmbGodownNo.Text = ""

TxtVendorID.Text = ""
TxtLocationID.Text = ""
TxtCMPName.Text = ""
TxtGodownNo.Text = ""


TxtVendorID.Visible = False
TxtLocationID.Visible = False
TxtCMPName.Visible = False
TxtGodownNo.Visible = False

CmbVendorID.Visible = True
CmbLocationID.Visible = True
CmbCMPName.Visible = True
CmbGodownNo.Visible = True
CmdVendor.Visible = True
CmdLocation.Visible = True
CmdCMP.Visible = True

''CmbStatus.Visible = True
''TxtStatus.Visible = False
''
''CmbForwardTo.Visible = True
''TxtForwardTo.Visible = False
''CmbForwardTo.Text = ""
''CmbStatus.Text = "Pending"
''lblFAGPostingDate.Visible = False
''lblFAGRemarks.Visible = False
''TxtFAGPostingDate.Visible = False
''TxtFAGRemarks.Visible = False
TxtMappID = ""
TxtRate.Text = ""
TxtFumDate.Value = Date
TxtNextFumDate = Date + 45
TxtRemark.Text = ""
TxtTotAmount.Text = ""
TxtBillableQty.Text = ""
TxtAddress.Text = ""
TxtVendorQty = ""
TxtVendorRate = ""
TxtCreated = ""
TxtUpdated = ""

CmdPrintPreview.Enabled = False
CmdPrint.Enabled = False

TxtBillNo = ""

If RsTxn_Fumigation.RecordCount > 0 Then
   TxtFumDate.SetFocus
End If

Call ClearStackFrame

Call Grid_Head

End Sub
Private Sub ClearStackFrame()
TxtStackNo.Text = ""
TxtCommodity.Text = ""
TxtNoofBags.Text = ""
TxtBalQty.Text = ""
TxtActualBags.Text = ""
TxtActualQty.Text = ""
CmbBillableFlag = ""
TxtRemark1.Text = ""
CmbTreatmentType.Text = ""
TxtRateStack.Text = ""
TxtAmountStack.Text = ""

End Sub

Private Sub CmbBillableFlag_LostFocus()

If CmbBillableFlag.Text = "" Then Exit Sub

If LCase(CmbBillableFlag.Text) <> "no" And LCase(CmbBillableFlag.Text) <> "yes" Then
   MsgBox "Invalid Selection!!!"
   CmbBillableFlag.SetFocus
   Exit Sub
End If

End Sub

Private Sub CmbForwardTo_GotFocus()
tmp = CmbForwardTo.Text
Call TableMst_Employee("Where  Flag = 'Y' ")
CmbForwardTo.Clear
If RsMst_Employee.RecordCount = 0 Then
   MsgBox "No Employee Found!!!"
   CmbForwardTo.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Employee.RecordCount
    CmbForwardTo.AddItem RsMst_Employee!EmployeeName & "~" & RsMst_Employee!EmployeeNo
    RsMst_Employee.MoveNext
Next
CmbForwardTo.Text = tmp
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

Private Sub CmbSBillGenerated_LostFocus()

If CmbSBillGenerated.Text = "" Then Exit Sub

If LCase(CmbSBillGenerated.Text) <> "no" And LCase(CmbSBillGenerated.Text) <> "yes" Then
   MsgBox "Invalid Selection!!!"
   CmbSBillGenerated.SetFocus
   Exit Sub
End If


End Sub

Private Sub CmbSForwardTo_GotFocus()

tmp = CmbSForwardTo.Text
Call TableMst_Employee("Where  Flag = 'Y' ")
CmbSForwardTo.Clear
If RsMst_Employee.RecordCount = 0 Then
   MsgBox "No Employee Found!!!"
   CmbSForwardTo.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Employee.RecordCount
    CmbSForwardTo.AddItem RsMst_Employee!EmployeeName & "~" & RsMst_Employee!EmployeeNo
    RsMst_Employee.MoveNext
Next
CmbSForwardTo.Text = tmp

End Sub

Private Sub CmbSForwardTo_LostFocus()

Dim strSplitString() As String
Dim EmpNo_, EmpName_ As String

If CmbSForwardTo.Text = "" Then Exit Sub


strSplitString = Split(CmbSForwardTo, "~")
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
   CmbSForwardTo.SetFocus
   Exit Sub
End If
mSForwardToID = RsMst_Employee!EmployeeID
'TxtFWDDesignation = GetDesignaion(RsMst_Employee!DesignationID)
'mForwardToEmailID = RsMst_Employee!EmployeeEmailID

End Sub

Private Sub CmbSStatus_LostFocus()
If CmbSStatus.Text = "" Then Exit Sub
If LCase(CmbSStatus.Text) <> "pending" And LCase(CmbSStatus.Text) <> "approved" And LCase(CmbSStatus.Text) <> "rejected" Then
    MsgBox "Invalid Selection!!!"
    CmbSStatus.SetFocus
End If
End Sub

Private Sub CmbStatus_LostFocus()
If CmbStatus.Text = "" Then Exit Sub
If LCase(CmbStatus.Text) <> "pending" And LCase(CmbStatus.Text) <> "approved" And LCase(CmbStatus.Text) <> "rejected" Then
    MsgBox "Invalid Selection!!!"
    CmbStatus.SetFocus
End If
If LCase(CmbStatus.Text) = "pending" Then
    lblFAGPostingDate.Visible = False
    lblFAGRemarks.Visible = False
    TxtFAGPostingDate.Visible = False
    TxtFAGRemarks.Visible = False
Else
    lblFAGPostingDate.Visible = True
    lblFAGRemarks.Visible = True
    TxtFAGPostingDate.Visible = True
    TxtFAGRemarks.Visible = True
    TxtFAGPostingDate.Value = Date
    TxtFAGRemarks = ""
End If

If LCase(Gen_UserRole) <> "fag" And LCase(Gen_UserRole) <> "fagadmin" And LCase(Gen_UserRole) <> "power" Then
   If LCase(CmbStatus.Text) <> "pending" Then
      MsgBox "Please forward this claim to FAG user, Only FAG user have access to approve claim '!!!"
      CmbStatus.SetFocus
      Exit Sub
   End If
End If

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
tmp = tmp & " Where MFA.Status = 'A' And MFA.Type = 'F' And MFA.LocationID = " & mLocationID & " And MFA.VendorID = " & mVendorID
tmp = tmp & " And '" & Format(TxtFumDate.Value, "yyyy-mm-dd") & "' between MFA.StartDate and MFA.ExpiryDate"
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
TxtRateStack.Text = IIf(IsNull(TmpRs7!RateperMT), 0, TmpRs7!RateperMT)
TxtAmountStack.Text = Format(Val(TxtRateStack.Text) * Val(TxtActualQty.Text), "#0")
End Sub

Private Sub CmbVendorID_GotFocus()
If CmbLocationID.Text = "" Then
   MsgBox "Select Location!!!"
   CmbLocationID.SetFocus
   Exit Sub
End If
tmp = CmbVendorID.Text
If Edit_Flg = "A" Then
   Call TableMst_Vendor("Where VendorID in (Select VendorID from Mst_FumigationAgreement Where Type = 'F' And Status = 'A' And LocationID = " & Val(mLocationID) & ")") '("Where Flag = 'Y'")
Else
   Call TableMst_Vendor("Where VendorID = " & Val(mVendorID) & " Or  VendorID in (Select VendorID from Mst_FumigationAgreement Where Type = 'F' And Status = 'A' And LocationID = " & Val(mLocationID) & ")") '("Where Flag = 'Y'")
End If
'Call TableMst_Vendor '("Where Flag = 'Y'")
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


Private Sub CmdCMP_Click()
FrmMst_CMP.Show
End Sub


Private Sub CmdLocation_Click()
FrmMst_Location.Show
End Sub

Private Sub CmdPrint_Click()

RsTxn_Fumigation!Flag = "P"
RsTxn_Fumigation!PrintedDate = Date
RsTxn_Fumigation.Update

Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "FCR_" & TxtFumigationID.Text & ".xls")
Call Xls_Detail_Rpt1
End Sub

Private Sub CmdPrintPreview_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "FCR_" & TxtFumigationID.Text & ".xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "FCR_" & TxtFumigationID.Text & ".xls")
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
tmp = "a" & mRow & ":b" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, 1) = "Fumigation ID"
tmp = "c" & mRow & ":i" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, 3) = TxtFumigationID.Text

oWS.Cells(mRow, 1).Font.Bold = True

mRow = mRow + 1
tmp = "a" & mRow & ":b" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, 1) = "Fumigation Start Date"
tmp = "c" & mRow & ":i" & mRow
oWS.Range(tmp).Merge

oWS.Cells(mRow, 3) = "'" & Format(TxtFumDate.Value, "dd-MMM-yyyy")
oWS.Cells(mRow, 1).Font.Bold = True

mRow = mRow + 1

tmp = "a" & mRow & ":b" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, 1) = "Next Fumigation Date"
tmp = "c" & mRow & ":i" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, 3) = "'" & Format(TxtNextFumDate.Value, "dd-MMM-yyyy")
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
oWS.Cells(mRow, 1) = "Billable Qty"
tmp = "c" & mRow & ":i" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, 3) = TxtBillableQty.Text
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


mRow = mRow + 2
tmp = "a" & mRow & ":k" & mRow
oWS.Range(tmp).Font.Bold = True
oWS.Range(tmp).WrapText = True
oWS.Cells(mRow, 1).RowHeight = 30
mtmpRow = mRow

oWS.Range("a:a").ColumnWidth = 8
oWS.Range("b:b").ColumnWidth = 11
oWS.Range("c:c").ColumnWidth = 7
oWS.Range("d:d").ColumnWidth = 7
oWS.Range("e:e").ColumnWidth = 9.5 '32.14
oWS.Range("f:f").ColumnWidth = 8
oWS.Range("g:g").ColumnWidth = 14
oWS.Range("h:h").ColumnWidth = 7.75
oWS.Range("i:i").ColumnWidth = 9.5
oWS.Range("j:j").ColumnWidth = 8
oWS.Range("k:k").ColumnWidth = 10

For I = 1 To MSHFlexGrid1.Rows - 2
    'c = 1
    For C = 0 To MSHFlexGrid1.Cols - 2
        If InStr(MSHFlexGrid1.TextMatrix(I, C), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(I, C)) Then
              oWS.Cells(mRow, C + 1) = Format(MSHFlexGrid1.TextMatrix(I, C), "dd-MMM-yyyy")
           Else
              oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
           End If
        Else
           oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
        End If
         'c = c + 1
    Next
    'TmpRs.MoveNext
    mRow = mRow + 1
Next

tmp = "a" & mtmpRow & ":k" & mRow - 1
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
   MsgBox "From Date Must be less than To Date!!!"
   SFrom.SetFocus
   Exit Sub
End If

tmp = "Select TF.FumigationID,TF.FumigationDate,TF.NextFumigationDate,MV.VendorName,"
tmp = tmp & " ML.LocationName , MC.CMPName, MG.GodownNo, TF.Rate, TF.TotalQty, TF.BillableQty, TF.Amount,TF.Remark,TF.VendorRate,TF.VendorQuantity,TF.MappingID"
tmp = tmp & " From Txn_Fumigation TF"
tmp = tmp & " Inner Join Mst_Vendor MV on TF.VendorID=MV.VendorID"
tmp = tmp & " Inner Join Mst_Godown MG on TF.GodownID=MG.GodownID"
tmp = tmp & " Inner Join Mst_CMP MC on MG.CMPID=MC.CMPID"
tmp = tmp & " Inner Join Mst_Location ML on MC.LocationID=ML.LocationID"

If Gen_DBType = "MDB" Then
   Wc1 = " Where FumigationDate Between #" & Format(SFrom.Value, Gen_DatFormat) & "# And #" & Format(STo.Value, Gen_DatFormat) & "# "
Else
   Wc1 = " Where FumigationDate Between '" & Format(SFrom.Value, Gen_DatFormat) & "' And '" & Format(STo.Value, Gen_DatFormat) & "' "
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

If CmbSVendorID.Text <> "" Then
   Wc1 = Wc1 & " And MV.VendorName='" & CmbSVendorID.Text & "'"
End If
''
''If CmbSStatus.Text <> "" Then
''   wc1 = wc1 & " And TF.Status='" & Left(CmbSStatus.Text, 1) & "'"
''End If


If TxtSFumgID.Text <> "" Then
   Wc1 = Wc1 & " And TF.FumigationID=" & Val(TxtSFumgID.Text) & ""
End If

If LCase(CmbSBillGenerated.Text) = "no" Then
   Wc1 = Wc1 & " And (isnull(TF.MappingID,0) = 0 ) "
ElseIf LCase(CmbSBillGenerated.Text) = "yes" Then
   Wc1 = Wc1 & " And isnull(TF.MappingID,0) <> 0 "
End If

''If TxtSBillNo.Text <> "" Then
''   wc1 = wc1 & " And TF.BillNo Like '%" & Trim(TxtSBillNo.Text) & "%'"
''End If
''
''If ChkBatchNo.Value = 0 Then
''   wc1 = wc1 & " And isnull(TF.BatchID,0) = 0 "
''End If
''
''If CmbSForwardTo.Text <> "" Then
''   wc1 = wc1 & " And TF.ForwardTo =" & Val(mSForwardToID) & ""
''End If

If TxtSBillQty <> "" Then
    If Not IsOprater(TxtSBillQty) Then
       Wc1 = GenWc(Wc1, TxtSBillQty, "TF.BillableQty", "N", "N")
    Else
       Wc1 = GenWc(Wc1, TxtSBillQty, "TF.BillableQty", "N", "I")
    End If
End If

tmp = tmp & Wc1 & " Order By FumigationID"

Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   MSHFlexGrid2.Rows = TmpRs1.RecordCount + 3
   x = 2
   For I = 1 To TmpRs1.RecordCount
       For C = 0 To MSHFlexGrid2.Cols - 1
           MSHFlexGrid2.TextMatrix(x, C) = IIf(IsNull(TmpRs1.Fields(C)), "", TmpRs1.Fields(C))
       Next
       MSHFlexGrid2.TextMatrix(x, 1) = Format(MSHFlexGrid2.TextMatrix(x, 1), "dd-MMM-yyyy")
       MSHFlexGrid2.TextMatrix(x, 2) = Format(MSHFlexGrid2.TextMatrix(x, 2), "dd-MMM-yyyy")
       TmpRs1.MoveNext
       x = x + 1
   Next
End If

Label22.Caption = TmpRs1.RecordCount
Label22.Refresh

End Sub

Private Sub CmdTreatmentType_Click()
FrmMst_TreatmentType.Show
End Sub

Private Sub CmdUpdateList_Click()

r = MSHFlexGrid1.RowSel
If r < 2 Or r = MSHFlexGrid1.Rows - 1 Then
   Call ClearStackFrame
   Exit Sub
End If
If TxtActualBags.Text = "" Then
   MsgBox "Blank Actual No of bags not Allowed!"
   TxtActualBags.SetFocus
   Exit Sub
End If
If TxtActualQty.Text = "" Then
   MsgBox "Blank Actual Qty not Allowed!"
   TxtActualQty.SetFocus
   Exit Sub
End If
If CmbTreatmentType.Text = "" Then
   MsgBox "Blank Treatment Type not Allowed!"
   CmbTreatmentType.SetFocus
   Exit Sub
End If
If TxtRemark1.Text = "" Then
   MsgBox "Blank Remark not Allowed!"
   TxtRemark1.SetFocus
   Exit Sub
End If

MSHFlexGrid1.TextMatrix(0, 4) = Val(MSHFlexGrid1.TextMatrix(0, 4)) - Val(MSHFlexGrid1.TextMatrix(r, 4))
MSHFlexGrid1.TextMatrix(0, 5) = Val(MSHFlexGrid1.TextMatrix(0, 5)) - Val(MSHFlexGrid1.TextMatrix(r, 5))
MSHFlexGrid1.TextMatrix(0, 15) = Val(MSHFlexGrid1.TextMatrix(0, 15)) - Val(MSHFlexGrid1.TextMatrix(r, 15))
MSHFlexGrid1.TextMatrix(r, 9) = CmbBillableFlag.Text

MSHFlexGrid1.TextMatrix(r, 4) = Val(TxtActualBags.Text)
MSHFlexGrid1.TextMatrix(r, 5) = Val(TxtActualQty.Text)
MSHFlexGrid1.TextMatrix(r, 10) = TxtRemark1.Text
MSHFlexGrid1.TextMatrix(r, 12) = CmbTreatmentType.Text
MSHFlexGrid1.TextMatrix(r, 13) = mAgreementDetailID
MSHFlexGrid1.TextMatrix(r, 14) = TxtRateStack.Text
MSHFlexGrid1.TextMatrix(r, 15) = TxtAmountStack.Text

MSHFlexGrid1.TextMatrix(0, 4) = Val(MSHFlexGrid1.TextMatrix(0, 4)) + Val(MSHFlexGrid1.TextMatrix(r, 4))
MSHFlexGrid1.TextMatrix(0, 5) = Val(MSHFlexGrid1.TextMatrix(0, 5)) + Val(MSHFlexGrid1.TextMatrix(r, 5))
MSHFlexGrid1.TextMatrix(0, 15) = Val(MSHFlexGrid1.TextMatrix(0, 15)) + Val(MSHFlexGrid1.TextMatrix(r, 15))
TxtBillableQty.Text = "0"
For I = 2 To MSHFlexGrid1.Rows - 1
    If LCase(MSHFlexGrid1.TextMatrix(I, 9)) = "yes" Then
       TxtBillableQty.Text = Val(TxtBillableQty.Text) + Val(MSHFlexGrid1.TextMatrix(I, 5))
    End If
Next

'TxtTotAmount.Text = Format(Val(TxtRate.Text) * Val(TxtBillableQty), "#####0.00")
TxtTotAmount.Text = Format(Val(MSHFlexGrid1.TextMatrix(0, 15)), "#####0.00")
Call ClearStackFrame
End Sub


Private Sub CmdVendor_Click()
FrmMst_Vendor.Show
End Sub

Private Sub DeleteCmd_Click()

Dim ans As Variant
On Error GoTo msgError

If Val(TxtMappID) > 0 Then
   MsgBox "Mapping created for this fumigation, You can't edit this record."
   Exit Sub
End If

''If RsTxn_Fumigation!BillNo <> "" Then
''   MsgBox "You can't delete this record."
''   Exit Sub
''End If
''
''If LCase(TxtStatus) = "approved" Or LCase(TxtStatus) = "rejected" Then
''   MsgBox "This Claim is " & TxtStatus & ". You can't delete this claim!!!"
''   Exit Sub
''End If



ans = MsgBox("Do you really want to DELETE this record??? ", vbYesNo)
If ans = vbYes Then
   tmp = "Delete From Txn_FumigationDetail Where FumigationID = " & TxtFumigationID & ""
   Call TmpTable
   
   'RsTxn_Fumigation.MoveFirst
   'RsTxn_Fumigation.Find "FumigationID = " & TxtFumigationID & ""
   
   RsTxn_Fumigation.Delete
   RsTxn_Fumigation.Update
   If RsTxn_Fumigation.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsTxn_Fumigation.MoveNext
   If RsTxn_Fumigation.EOF() Then
      RsTxn_Fumigation.MoveLast
   End If
   Call Indata
End If

Exit Sub
msgError:
MsgBox "Child record Present "
RsTxn_Fumigation.CancelUpdate


End Sub

Private Sub EditCmd_Click()

If Val(TxtMappID) > 0 Then
   MsgBox "Mapping created for this fumigation, You can't edit this record."
   Exit Sub
End If
''If RsTxn_Fumigation!BillNo <> "" Then
''   MsgBox "You can't edit this record."
''   Exit Sub
''End If

''If LCase(TxtStatus) = "approved" Or LCase(TxtStatus) = "rejected" Then
''   MsgBox "This Claim is " & TxtStatus & ". You can't edit this claim any more!!!"
''   Exit Sub
''End If


Edit_Flg = "E"

Call GButtonLock(Me, False)


Frame1.Enabled = True
TxtFumDate.Enabled = False
'TxtRate.Locked = False

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

''CmbStatus.Visible = True
''TxtStatus.Visible = False
''CmbForwardTo.Visible = True
''TxtForwardTo.Visible = False


If RsTxn_Fumigation!Flag = "P" Then
'   TxtBillNo.Locked = False
   TxtRate.Locked = True
   TxtRemark.Locked = True
   Frame2.Enabled = False
Else
   TxtRate.Locked = False
   TxtRemark.Locked = False
   Frame2.Enabled = True
End If
CmbVendorID.Text = TxtVendorID.Text
CmbLocationID.Text = TxtLocationID.Text
CmbCMPName.Text = TxtCMPName.Text
CmbGodownNo.Text = TxtGodownNo.Text

CmdPrintPreview.Enabled = False
CmdPrint.Enabled = False

End Sub

Private Sub ExitCmd_Click()
Unload Me
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
   Frame2.Enabled = False
   Frame1.Enabled = False
   CmdPrintPreview.Enabled = False
   CmdPrint.Enabled = False
   Call GButtonLockAD(Me)
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If
Call Grid_Head
Call Grid_Head1
SFrom.Value = Date - 365
STo.Value = Date
Label22.Caption = ""
CmbSLocationID.Text = ""
CmbSCMPName.Text = ""
CmbSGodownNo.Text = ""
CmbSVendorID.Text = ""
TxtSFumgID.Text = ""

Call TableTxn_Fumigation(" Order by FumigationID")
If RsTxn_Fumigation.RecordCount = 0 Then
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
    CmbLocationID.AddItem RsMst_Location!Locationname
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
   CmbVendorID.Text = ""
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
   Call Grid_Head
   Call ClearStackFrame
   CmbVendorID.Text = ""
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
   Call TableMst_CMP(" Where CloseDate is null And LocationID = " & mLocationID & " And  " & Gen_WcCMP)
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
   Call Grid_Head
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
TxtAddress.Text = RsMst_Godown!Address
Call FillStackDetails(mGodownID)
End Sub
Private Sub FillStackDetails(mGodownID_ As Variant)

Call Grid_Head

tmp = "Select GSL.StackNo,MC.Commodity,sum(GSL.BalanceBags)BalanceBags,sum(GSL.BalanceWeight)BalanceWeight,EffectiveDays,MC.CommodityID"
tmp = tmp & " From Mst_GSL  GSL"
tmp = tmp & " Inner join Mst_Commodity MC on GSL.CommodityID=MC.CommodityID"
tmp = tmp & " Where  GSL.GodownID=" & mGodownID_ & " And GSL.Status not in ('E','Z') And MC.FumigationFlag='Y' And GSL.MadeUPFlag = '0' "
tmp = tmp & " Group By GSL.StackNo,MC.Commodity,EffectiveDays,MC.CommodityID"
'tmp = tmp & " Order By GSL.StackNo,MC.Commodity"

Call TmpTable

If TmpRs.RecordCount > 0 Then
   MSHFlexGrid1.Rows = TmpRs.RecordCount + 3
   For I = 2 To TmpRs.RecordCount + 1
       MSHFlexGrid1.TextMatrix(I, 0) = TmpRs!StackNo
       MSHFlexGrid1.TextMatrix(I, 1) = TmpRs!Commodity
       MSHFlexGrid1.TextMatrix(I, 2) = TmpRs!BalanceBags
       MSHFlexGrid1.TextMatrix(I, 3) = TmpRs!BalanceWeight
       MSHFlexGrid1.TextMatrix(I, 6) = GetLastFumigationDate(mGodownID_, TmpRs!StackNo, TmpRs!CommodityID)
       MSHFlexGrid1.TextMatrix(I, 7) = IIf(IsNull(TmpRs!EffectiveDays), 0, TmpRs!EffectiveDays)
       MSHFlexGrid1.TextMatrix(I, 11) = TmpRs!CommodityID
       If MSHFlexGrid1.TextMatrix(I, 6) <> "" Then
          MSHFlexGrid1.TextMatrix(I, 8) = DateDiff("d", CDate(MSHFlexGrid1.TextMatrix(I, 6)), TxtFumDate.Value)
       Else
          MSHFlexGrid1.TextMatrix(I, 8) = 0 'TmpRs!EffectiveDays
       End If
       TmpRs.MoveNext
   Next
End If
TxtBillableQty.Text = ""
TxtTotAmount = ""
End Sub
Private Function GetLastFumigationDate(mGodownID_ As Variant, mStack_ As Variant, mCommodityID_ As Variant) As Variant
Dim mRetval As Variant
mRetval = ""
tmp = "Select Max(FumigationDate) as FumigationDate  From Txn_Fumigation "
tmp = tmp & " inner Join txn_FumigationDetail On Txn_Fumigation.FumigationID=txn_FumigationDetail.FumigationID"
tmp = tmp & " Where GodownID=" & mGodownID_ & "  And StackNo='" & mStack_ & "' And CommodityID=" & mCommodityID_ & ""

Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   mRetval = Format(TmpRs1!FumigationDate, "dd-MMM-yyyy")
End If
GetLastFumigationDate = mRetval

End Function

Private Sub Form_Resize()
If Me.WindowState = vbMaximized Then
   SetScrollBars
End If
End Sub

Private Sub MSHFlexGrid1_Click()
If SaveCmd.Enabled = False Then Exit Sub
If Frame2.Enabled = False Then Exit Sub
Dim r As Variant

r = MSHFlexGrid1.RowSel
If r < 2 Or r = MSHFlexGrid1.Rows - 1 Then
   Call ClearStackFrame
   Exit Sub
End If
If MSHFlexGrid1.TextMatrix(r, 1) = "" Then Exit Sub

mCommodityID = MSHFlexGrid1.TextMatrix(r, 11)
TxtStackNo.Text = MSHFlexGrid1.TextMatrix(r, 0)
TxtCommodity.Text = MSHFlexGrid1.TextMatrix(r, 1)
TxtNoofBags.Text = MSHFlexGrid1.TextMatrix(r, 2)
TxtBalQty.Text = MSHFlexGrid1.TextMatrix(r, 3)
If MSHFlexGrid1.TextMatrix(r, 4) = "" Then
   TxtActualBags.Text = TxtNoofBags.Text
Else
   TxtActualBags.Text = MSHFlexGrid1.TextMatrix(r, 4)
End If

If MSHFlexGrid1.TextMatrix(r, 5) = "" Then
   TxtActualQty.Text = TxtBalQty.Text
Else
   TxtActualQty.Text = MSHFlexGrid1.TextMatrix(r, 5)
End If

If MSHFlexGrid1.TextMatrix(r, 9) = "" Then
   CmbBillableFlag.Text = "Yes"
Else
   CmbBillableFlag.Text = MSHFlexGrid1.TextMatrix(r, 9)
End If
TxtRemark1.Text = MSHFlexGrid1.TextMatrix(r, 10)
CmbTreatmentType.Text = MSHFlexGrid1.TextMatrix(r, 12)
mAgreementDetailID = MSHFlexGrid1.TextMatrix(r, 13)
TxtRateStack.Text = MSHFlexGrid1.TextMatrix(r, 14)
TxtAmountStack.Text = MSHFlexGrid1.TextMatrix(r, 15)
TxtActualBags.SetFocus

End Sub

Private Sub MSHFlexGrid2_Click()
Dim I As Variant
I = MSHFlexGrid2.RowSel
If I <= 0 Then Exit Sub
If MSHFlexGrid2.TextMatrix(I, 0) = "" Then Exit Sub
RsTxn_Fumigation.MoveFirst
RsTxn_Fumigation.Find "FumigationID= " & Val(MSHFlexGrid2.TextMatrix(I, 0)) & ""
If Not RsTxn_Fumigation.EOF Then
   Call Indata
   Frame0.Visible = False
End If

End Sub

Private Sub SaveCmd_Click()

If TxtFumDate.Value > Date Then
   MsgBox "Future Date not allowed!!"
   TxtFumDate.SetFocus
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
''comment when added fumigation agreement master
'If TxtRate.Text = "" Then
'   MsgBox "Blank Rate not Allowed!!"
'   TxtRate.SetFocus
'   Exit Sub
'End If

If TxtRemark.Text = "" Then
   MsgBox "Blank Remark not Allowed!!"
   TxtRemark.SetFocus
   Exit Sub
End If

''If CmbForwardTo.Text = "" Then
''   MsgBox "Blank Forward To not Allowed!!"
''   CmbForwardTo.SetFocus
''   Exit Sub
''End If
''
''If CmbStatus.Text = "" Then
''   MsgBox "Blank Status not Allowed!!"
''   CmbStatus.SetFocus
''   Exit Sub
''End If

''If LCase(CmbStatus.Text) <> "pending" Then
''   If TxtFAGRemarks.Text = "" Then
''      MsgBox "Blank FAG Remarks not Allowed!!"
''      TxtFAGRemarks.SetFocus
''      Exit Sub
''   End If
''   If TxtBillNo = "" Then
''      MsgBox "Blank Bill No not Allowed!!"
''      TxtBillNo.SetFocus
''      Exit Sub
''   End If
''End If

If LCase(Gen_UserRole) = "power" Or LCase(Gen_UserRole) = "admin" Or LCase(Gen_UserRole) = "superadmin" Or LCase(Gen_UserRole) = "headcmg" Then
   
Else
   If TxtFumDate.Value < Date - 7 Then
      MsgBox " You can't enter fumigation details less than 7 days, Contact CMG!!"
      Exit Sub
   End If
End If

If MSHFlexGrid1.Rows < 4 Then
   MsgBox "Stack Details Missing."
   Exit Sub
End If

If CheckActualQtyEntered = False Then
   MsgBox "Blank Actual Quantity for all Stacks not allowed!!"
   Exit Sub
End If

If CheckFumigationDate = False Then
   MsgBox "Fumigation Date should be Greater than Last Fumigation Date!!"
   TxtFumDate.SetFocus
   Exit Sub
End If

''comment when added fumigation agreement master
'TxtTotAmount.Text = Format(Val(TxtRate.Text) * Val(TxtBillableQty), "#####0.00")
If Val(TxtBillableQty) > 0 Then
   TxtRate.Text = Format(Val(TxtTotAmount.Text) / Val(TxtBillableQty), "#####0.00")
End If
   

If Edit_Flg = "A" Then
   RsTxn_Fumigation.AddNew
   RsTxn_Fumigation!FumigationID = getMaxFumigationID
   RsTxn_Fumigation!G_UID = GetGUID
   TxtFumigationID.Text = RsTxn_Fumigation!FumigationID
Else
   RsTxn_Fumigation.MoveFirst
   RsTxn_Fumigation.Find " FumigationID= '" & Val(TxtFumigationID.Text) & "'"
   Call DeleteChildTableDetail
End If

RsTxn_Fumigation!FumigationID = Val(TxtFumigationID.Text)
RsTxn_Fumigation!FumigationDate = Format(TxtFumDate.Value, Gen_DatFormat)
RsTxn_Fumigation!NextFumigationDate = Format(TxtNextFumDate.Value, Gen_DatFormat)
RsTxn_Fumigation!VendorID = mVendorID
RsTxn_Fumigation!GodownID = mGodownID
RsTxn_Fumigation!Remark = TxtRemark.Text
RsTxn_Fumigation!Rate = Format(Val(TxtRate.Text), "#####0.00")
RsTxn_Fumigation!TotalQty = Val(MSHFlexGrid1.TextMatrix(0, 5))
RsTxn_Fumigation!BillableQty = Val(TxtBillableQty.Text)
RsTxn_Fumigation!Amount = Format(Val(TxtTotAmount.Text), "#####0.00")

''RsTxn_Fumigation!BatchID = 0
''RsTxn_Fumigation!ForwardTo = mForwardToID
''RsTxn_Fumigation!Status = Left$(CmbStatus.Text, 1)
''If LCase(CmbStatus.Text) <> "pending" Then
''    RsTxn_Fumigation!ApprovalDate = TxtFAGPostingDate.Value
''    RsTxn_Fumigation!ApprovalRemark = TxtFAGRemarks.Text
''End If
''
''If TxtBillNo.Text <> "" Then
''   RsTxn_Fumigation!BillNo = TxtBillNo.Text
''Else
''   RsTxn_Fumigation!BillNo = ""
''End If

If IsNull(RsTxn_Fumigation!CreatedBy) = True Or RsTxn_Fumigation!CreatedBy = "" Then
   RsTxn_Fumigation!CreatedBy = Gen_UserLoginID
   RsTxn_Fumigation!CreatedDate = Now
Else
   RsTxn_Fumigation!UpdatedBy = Gen_UserLoginID
   RsTxn_Fumigation!UpdatedDate = Now
End If

RsTxn_Fumigation.Update

Call SaveFumigationDetails
''mForwardToEmailID = GetEmployeeEmailID(mForwardToID)

'If LCase(CmbStatus.Text) <> "pending" Then
   'Call Gen_Email
'End If

SaveCmd.Enabled = False
Call Indata
''Call Gen_Email
End Sub

Private Sub Gen_Email()
Dim TOADD As Variant
Dim TOCC, mMsg, mGridMsg As Variant
Dim strSplitString() As String
Dim EmpNo_, EmpName_ As String

mMsg = ""
mMsg = "<tr> Fumigation Expense (From Store-Man)</tr>"
mMsg = mMsg & "<tr><td> Fumigation ID </td><td> " & TxtFumigationID & " </td></tr>"  'mMsg = mMsg & "<tr><td> Fumigation ID </td><td> " & TxtPreFix & " - " & TxtClaimID & " </td></tr>"
mMsg = mMsg & "<tr><td> Fumigation Start Date </td><td> " & Format(TxtFumDate.Value, "dd-mmm-yyyy") & "  </td></tr>"
'mMsg = mMsg & "<tr><td> Expense From </td><td> " & Format(TxtExpenseFrom.Value, "dd-mmm-yyyy") & "  </td></tr>"
'mMsg = mMsg & "<tr><td> Expense To </td><td> " & Format(TxtExpenseTo.Value, "dd-mmm-yyyy") & "  </td></tr>"
'If CmbStatus.Visible = True Then
'   mMsg = mMsg & "<tr><td> Status </td><td> " & CmbStatus.Text & "  </td></tr>"
'Else
'   mMsg = mMsg & "<tr><td> Status </td><td> " & TxtStatus & "  </td></tr>"
'End If
mMsg = mMsg & "<tr><td> Status </td><td> " & TxtStatus.Text & "  </td></tr>"
'strSplitString = Split(CmbEmployeeID.Text, "~")
'EmpName_ = Trim(strSplitString(0))
'EmpNo_ = Trim(strSplitString(1))
 
'mMsg = mMsg & "<tr><td> Employee Name </td><td> " & EmpName_ & "  </td></tr>"
'mMsg = mMsg & "<tr><td> Employee No </td><td> " & EmpNo_ & "  </td></tr>"

mMsg = mMsg & "<tr><td> Location </td><td> " & TxtLocationID.Text & "  </td></tr>"
mMsg = mMsg & "<tr><td> Entered By </td><td> " & TxtCreated & "  </td></tr>"
If Edit_Flg = "E" Then
   mMsg = mMsg & "<tr><td> Updated By </td><td> " & TxtUpdated & "  </td></tr>"
End If
mMsg = mMsg & "<tr><td> Amount </td><td> " & TxtTotAmount & "  </td></tr>"
mMsg = mMsg & "<tr><td> Remark </td><td> " & TxtRemark & "   </td></tr>"

strSplitString = Split(TxtForwardTo.Text, "~")
EmpName_ = Trim(strSplitString(0))
EmpNo_ = Trim(strSplitString(1))
mMsg = mMsg & "<tr><td> Forwarded To </td><td> " & EmpName_ & " " & " ~ " & " " & EmpNo_ & " </td></tr>"


If LCase(TxtStatus.Text) <> "pending" Then
   mMsg = mMsg & "<tr><td> FAG Remarks </td><td> " & TxtFAGRemarks & " </td></tr>"
   mMsg = mMsg & "<tr><td> FAG Posting Date </td><td> " & Format(TxtFAGPostingDate.Value, "dd-mmm-yyyy") & " </td></tr>"
End If
'mGridMsg = "<tr><td>Expense Type</td><td>Claim Amount</td><td>Particulars</td></tr>"
For I = 1 To MSHFlexGrid1.Rows - 1
    If MSHFlexGrid1.TextMatrix(I, 1) <> "" Then
       mGridMsg = mGridMsg & "<tr><td> " & MSHFlexGrid1.TextMatrix(I, 0) & " </td><td> " & MSHFlexGrid1.TextMatrix(I, 1) & " </td><td> " & MSHFlexGrid1.TextMatrix(I, 2) & " </td><td> " & MSHFlexGrid1.TextMatrix(I, 3) & " </td><td> " & MSHFlexGrid1.TextMatrix(I, 4) & " </td><td> " & MSHFlexGrid1.TextMatrix(I, 5) & " </td><td> " & MSHFlexGrid1.TextMatrix(I, 6) & " </td><td> " & MSHFlexGrid1.TextMatrix(I, 7) & " </td><td> " & MSHFlexGrid1.TextMatrix(I, 8) & " </td><td> " & MSHFlexGrid1.TextMatrix(I, 9) & " </td></tr>"
    End If
Next

mMsg = mMsg & "<tr><td> Details </td><td><table border = 1>" & mGridMsg & "</table></td>"


'TOADD = mForwardToEmailID ' mEmployeeEmailID
TOCC = Gen_UserEmailID '& "," & mForwardToEmailID

If LCase(TxtStatus.Text) = "pending" Then
   TOADD = mForwardToEmailID
   'If LCase(TOCC) <> LCase(mEmployeeEmailID) Then
   '   TOCC = TOCC & "," & mEmployeeEmailID
   'End If
   mMsgSubject = "Fumigation Pending"
ElseIf LCase(TxtStatus.Text) = "approved" Then
   'TOADD = mEmployeeEmailID
   If LCase(TOCC) <> LCase(mForwardToEmailID) Then
      TOCC = TOCC & "," & mForwardToEmailID
   End If
   mMsgSubject = "Fumigation Approved"
ElseIf LCase(TxtStatus.Text) = "rejected" Then
   'TOADD = mEmployeeEmailID
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

servertxt = "mail.ncml.com"

htmlStringhead = "<html><body><font size= 3 face= Zurich BT color = Blue><table border = 1>"
htmlStringtail = "</table></Font></body></html>"
   
'With oSMTP
'  'connection
'  .Server = servertxt
'  'message
'  If Edit_Flg = "A" Then
'     .MailFrom = mEmployeeEmailID
'  Else
'     .MailFrom = Gen_UserEmailID
'  End If
'  If IsNull(Trim(ToADD_)) = True Or Trim(ToADD_) = "" Then
'     ToADD_ = "priti.g@ncml.com"
'  End If
'  .SendTo = ToADD_
'  If ToADD_ <> TOCC_ Then
'     .CC = TOCC_
'  End If
'  '.BCC = "mayur.d@ncml.com"
'  .MessageSubject = MsgSubject_ & " (from Store-Man)" '"Confidential Mail"
'  .MessageHTML = htmlStringhead & Message1_ & htmlStringtail
'
'  '.MailFrom = "rakesh.go@ncml.com"
'  '.SendTo = "parag.r@ncml.com"
'  '.CC = "rakesh.go@ncml.com"
'  .TimeStamp = Now()
'  If LCase(Gen_EmailConfigFlag) = "y" Then
'     .SendEmail
'  End If
'End With

Dim objSMTP
Set objSMTP = CreateObject("MailBee.SMTP")
objSMTP.LicenseKey = "MBC900-5F470FAD67-D8E202A2E22927335F24F5A667BFD196"
objSMTP.ServerName = Gen_ReadMailServerIPINI("MailIP") ' = "mail.ncml.com" '"mail.ncml.com"
'objSMTP.FromAddr = Gen_UserEmailID
If Edit_Flg = "A" Then
   objSMTP.FromAddr = mEmployeeEmailID
Else
   objSMTP.FromAddr = Gen_UserEmailID
End If
If IsNull(Trim(ToADD_)) = True Or Trim(ToADD_) = "" Then
   ToADD_ = "manu.ag@ncml.com"
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


Private Function CheckActualQtyEntered() As Boolean
Dim Retval As Boolean
Retval = False
For I = 2 To MSHFlexGrid1.Rows - 1
    If MSHFlexGrid1.TextMatrix(I, 5) <> "" Then
       Retval = True
       Exit For
    End If
Next
CheckActualQtyEntered = Retval
End Function
Private Function CheckFumigationDate() As Boolean
Dim Retval As Boolean
Retval = True
For I = 2 To MSHFlexGrid1.Rows - 1
    If MSHFlexGrid1.TextMatrix(I, 4) <> "" Then
       If MSHFlexGrid1.TextMatrix(I, 6) <> "" Then
          If CDate(MSHFlexGrid1.TextMatrix(I, 6)) >= TxtFumDate Then
             Retval = False
             Exit For
          End If
       End If
    End If
Next
CheckFumigationDate = Retval
End Function


Sub DeleteChildTableDetail()
tmp = "Delete from Txn_FumigationDetail where FumigationID = " & Val(TxtFumigationID.Text) & ""
Call TmpTable
End Sub
Private Sub SaveFumigationDetails()
Call TableTxn_FumigationDetail(" Where FumigationID=" & Val(TxtFumigationID) & "")

For I = 2 To MSHFlexGrid1.Rows - 1
    If MSHFlexGrid1.TextMatrix(I, 4) <> "" Then
       RsTxn_FumigationDetail.AddNew
       RsTxn_FumigationDetail!G_UID = GetGUID
       RsTxn_FumigationDetail!FumgDetailID = getMaxFumigationDetailID
       RsTxn_FumigationDetail!FumigationID = Val(TxtFumigationID)
       RsTxn_FumigationDetail!StackNo = MSHFlexGrid1.TextMatrix(I, 0)
       RsTxn_FumigationDetail!CommodityID = Val(MSHFlexGrid1.TextMatrix(I, 11))
       RsTxn_FumigationDetail!NoofBags = Val(MSHFlexGrid1.TextMatrix(I, 2))
       RsTxn_FumigationDetail!BalanceQty = Val(MSHFlexGrid1.TextMatrix(I, 3))
       RsTxn_FumigationDetail!ActualNoOfBags = Val(MSHFlexGrid1.TextMatrix(I, 4))
       RsTxn_FumigationDetail!ActualQty = Val(MSHFlexGrid1.TextMatrix(I, 5))
       If MSHFlexGrid1.TextMatrix(I, 6) <> "" Then
         RsTxn_FumigationDetail!LastFumgDate = Format(MSHFlexGrid1.TextMatrix(I, 6), Gen_DatFormat)
       End If
       RsTxn_FumigationDetail!BillableFlag = IIf(MSHFlexGrid1.TextMatrix(I, 9) = "Yes", "Y", "N")
       RsTxn_FumigationDetail!Remark = MSHFlexGrid1.TextMatrix(I, 10)
       If MSHFlexGrid1.TextMatrix(I, 13) <> "" Then
          RsTxn_FumigationDetail!AgreementDetailID = Val(MSHFlexGrid1.TextMatrix(I, 13))
       'Else
       '   RsTxn_FumigationDetail!AgreementDetailID = Null
       End If
       'RsTxn_FumigationDetail!RateperMT = Val(MSHFlexGrid1.TextMatrix(I, 14))
       RsTxn_FumigationDetail!Amount = Val(MSHFlexGrid1.TextMatrix(I, 15))
       RsTxn_FumigationDetail.Update
       RsTxn_FumigationDetail.Requery
    End If
Next
End Sub

Private Function getMaxFumigationID() As Double
Dim Retval As Double
tmp = "Select max(FumigationID) from Txn_Fumigation"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
getMaxFumigationID = Retval
End Function
Private Function getMaxFumigationDetailID() As Double
Dim Retval As Double
tmp = "Select max(FumgDetailID) from Txn_FumigationDetail"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
getMaxFumigationDetailID = Retval
End Function

Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   If RsTxn_Fumigation.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   RsTxn_Fumigation.MoveFirst
   Call ClearStackFrame
   RsTxn_Fumigation.Find "FumigationID = " & Val(TxtFumigationID) & ""
   If RsTxn_Fumigation.EOF Then
      RsTxn_Fumigation.MoveLast
   End If
   Call Indata
   Exit Sub
End If
Frame0.Height = Frame4.Height - 200
Frame0.Left = Frame1.Left - 100
Frame0.Top = Frame1.Top
Frame0.Width = Frame4.Width - 100
Frame0.Visible = True

MSHFlexGrid2.Width = Frame0.Width - 500
MSHFlexGrid2.Height = Frame0.Height - 2500

RsTxn_Fumigation.MoveFirst
RsTxn_Fumigation.Filter = ""

''===========comment on 7 july 2012 for not clearing search grid=============
'Call Grid_Head1
'SFrom.Value = Date - 365
'STo.Value = Date
'Label22.Caption = ""
'CmbSLocationID.Text = ""
'CmbSCMPName.Text = ""
'CmbSGodownNo.Text = ""
'CmbSVendorID.Text = ""
'TxtSFumgID.Text = ""
''===========comment on 7 july 2012 for not clearing search grid=============

'TxtSBillNo.Text = ""
'CmbSStatus.Text = ""


'tmp = "Select EmployeeID,EmployeeName +'~'+ EmployeeNo as Employee,DesignationID,EmployeeEmailID From Mst_Employee Where EmployeeNo = '" & Gen_UserLoginID & "'"
'Call TmpTable

'If TmpRs.RecordCount = 1 Then
'   CmbSForwardTo.Text = TmpRs!Employee
'   mSForwardToID = TmpRs!EmployeeID
'Else
'   CmbSForwardTo.Text = ""
'   mSForwardToID = Null
'End If

End Sub

Private Sub TxtActualBags_LostFocus()
If TxtActualBags = "" Then Exit Sub
If Not IsNumeric(TxtActualBags) Or Val(TxtActualBags) <= 0 Then
   MsgBox "Invalid numeric format !!!! "
   TxtActualBags.SetFocus
End If

End Sub

Private Sub TxtActualQty_LostFocus()
If TxtActualQty = "" Then Exit Sub
If Not IsNumeric(TxtActualQty) Or Val(TxtActualQty) <= 0 Then
   MsgBox "Invalid numeric format !!!! "
   TxtActualQty.SetFocus
End If
TxtAmountStack.Text = Format(Val(TxtRateStack.Text) * Val(TxtActualQty.Text), "#0")
End Sub

Private Sub TxtFAGPostingDate_LostFocus()
If TxtFAGPostingDate.Value > Date Then
    MsgBox "Future FAG Posting Date not Allowed!!!"
    TxtFAGPostingDate.SetFocus
    Exit Sub
End If
End Sub

Private Sub TxtFumDate_LostFocus()

If Format(Gen_BlockingDate, "yyyy-mm") >= Format(TxtFumDate.Value, "yyyy-mm") Then
   MsgBox "Fumigation Start Date allowed for next month of " & Format(Gen_BlockingDate, "MMM-yyyy") & "!!!"
   TxtFumDate.SetFocus
   Exit Sub
End If

TxtNextFumDate = TxtFumDate.Value + 45


End Sub

Private Sub TxtRate_Lostfocus()
If TxtRate = "" Then Exit Sub
If Not IsNumeric(TxtRate) Or Val(TxtRate) <= 0 Then
   MsgBox "Invalid numeric format !!!! "
   TxtRate.SetFocus
End If


End Sub
Public Sub Indata()
 
Call ClearStackFrame
Frame2.Enabled = False
Frame1.Enabled = False
Call Grid_Head

''CmbStatus.Visible = False
''TxtStatus.Visible = True
''CmbForwardTo.Visible = False
''TxtForwardTo.Visible = True

TxtFumDate.Enabled = True
CmbVendorID.Visible = False
CmbLocationID.Visible = False
CmbCMPName.Visible = False
CmbGodownNo.Visible = False
CmdVendor.Visible = False

CmdLocation.Visible = False
CmdCMP.Visible = False

TxtVendorID.Visible = True
TxtLocationID.Visible = True
TxtCMPName.Visible = True
TxtGodownNo.Visible = True

TxtFumigationID = RsTxn_Fumigation!FumigationID
TxtFumDate.Value = RsTxn_Fumigation!FumigationDate
TxtNextFumDate.Value = RsTxn_Fumigation!NextFumigationDate
TxtVendorID.Text = GetVendorName(RsTxn_Fumigation!VendorID, "N")
mVendorID = RsTxn_Fumigation!VendorID
TxtGodownNo.Text = GetGodownName(RsTxn_Fumigation!GodownID, "N")
TxtAddress.Text = GetGodownName(RsTxn_Fumigation!GodownID, "")
mGodownID = RsTxn_Fumigation!GodownID
mCMPID = GetCMPIDFromGodown(RsTxn_Fumigation!GodownID)
mLocationID = GetLocationIDFromCMP(Val(mCMPID))
TxtCMPName.Text = GetCMPFromGodown(RsTxn_Fumigation!GodownID)
TxtLocationID.Text = GetLocationFromCMPName(TxtCMPName.Text)
TxtRemark.Text = RsTxn_Fumigation!Remark
TxtBillableQty.Text = RsTxn_Fumigation!BillableQty
'TxtBillNo = IIf(IsNull(RsTxn_Fumigation!BillNo), "", RsTxn_Fumigation!BillNo)
TxtRate.Text = Format(RsTxn_Fumigation!Rate, "#####0.00")
TxtTotAmount.Text = Format(RsTxn_Fumigation!Amount, "#####0.00")

'mForwardToID = IIf(IsNull(RsTxn_Fumigation!ForwardTo), 0, RsTxn_Fumigation!ForwardTo)
'TxtForwardTo = GetEmployeeNonName(Val(mForwardToID))
'CmbForwardTo = TxtForwardTo
'If RsTxn_Fumigation!Status = "A" Then
'    TxtStatus.Text = "Approved"
'ElseIf RsTxn_Fumigation!Status = "R" Then
'    TxtStatus.Text = "Rejected"
'Else
'    TxtStatus.Text = "Pending"
'End If
'CmbStatus = TxtStatus
'Call CmbStatus_LostFocus
'If LCase(TxtStatus.Text) <> "pending" Then
'    TxtFAGPostingDate.Value = IIf(IsNull(RsTxn_Fumigation!ApprovalDate), Date, RsTxn_Fumigation!ApprovalDate)
'    TxtFAGRemarks = IIf(IsNull(RsTxn_Fumigation!ApprovalRemark), "", RsTxn_Fumigation!ApprovalRemark)
'End If

TxtVendorQty = IIf(IsNull(RsTxn_Fumigation!VendorQuantity), "", RsTxn_Fumigation!VendorQuantity)
TxtVendorRate = IIf(IsNull(RsTxn_Fumigation!VendorRate), "", RsTxn_Fumigation!VendorRate)

TxtCreated = IIf(IsNull(RsTxn_Fumigation!CreatedBy), "", RsTxn_Fumigation!CreatedBy) & " :- " & IIf(IsNull(RsTxn_Fumigation!CreatedDate), "", RsTxn_Fumigation!CreatedDate)
TxtUpdated = IIf(IsNull(RsTxn_Fumigation!UpdatedBy), "", RsTxn_Fumigation!UpdatedBy) & " :- " & IIf(IsNull(RsTxn_Fumigation!UpdatedDate), "", RsTxn_Fumigation!UpdatedDate)

CmdPrintPreview.Enabled = True
CmdPrint.Enabled = True

TxtMappID = IIf(IsNull(RsTxn_Fumigation!MappingID), 0, RsTxn_Fumigation!MappingID)
TxtBillNo.Text = ""
If Val(TxtMappID) > 0 Then
   tmp = "Select BillNo from Txn_FumigationApproval Where MappingID=" & RsTxn_Fumigation!MappingID
   Call TmpTable
   If TmpRs.RecordCount > 0 Then
      TxtBillNo.Text = IIf(IsNull(TmpRs!BillNo), "", TmpRs!BillNo)
   End If
End If
Call InDataStack

If RsTxn_Fumigation.RecordCount = 1 Then
   Call GButtonLock_1(Me, True)
Else
   Call GButtonLock(Me, True)
End If
End Sub

Private Sub InDataStack()

'Call TableTxn_FumigationDetail(" Where FumigationID=" & Val(TxtFumigationID))

tmp = "Select TF.StackNo, TF.CommodityID, MC.Commodity, TF.NoOfBags, TF.BalanceQty, TF.ActualNoOfBags, TF.ActualQty, TF.BillableFlag, MC.EffectiveDays, TF.LastFumgDate, TF.Remark, "
tmp = tmp & " MTT.TreatmentType, TF.AgreementDetailID, MFT.RateperMT, TF.Amount"
tmp = tmp & " From Txn_FumigationDetail TF"
tmp = tmp & " Left join Mst_Commodity MC on TF.CommodityID=MC.CommodityID"
tmp = tmp & " Left Join Mst_FumigationTreatmentDetail MFT on MFT.DetailID =TF.AgreementDetailID "
tmp = tmp & " Left join Mst_TreatmentType MTT on MFT.TreatmentTypeID =MTT.TreatmentTypeID"
'tmp = tmp & " And MFT.AgreementID in (Select MFA.AgreementID From MSt_FumigationAgreement MFA Where MFA.LocationID = " & mLocationID & " And MFA.VendorID = " & mVendorID & " ) "
tmp = tmp & " Where FumigationID=" & Val(TxtFumigationID)

Call TmpTable

MSHFlexGrid1.Rows = TmpRs.RecordCount + 3
For I = 2 To MSHFlexGrid1.Rows - 2
   MSHFlexGrid1.TextMatrix(I, 0) = TmpRs!StackNo
   MSHFlexGrid1.TextMatrix(I, 1) = TmpRs!Commodity
   MSHFlexGrid1.TextMatrix(I, 2) = TmpRs!NoofBags
   MSHFlexGrid1.TextMatrix(I, 3) = TmpRs!BalanceQty
   MSHFlexGrid1.TextMatrix(I, 4) = TmpRs!ActualNoOfBags
   MSHFlexGrid1.TextMatrix(I, 5) = TmpRs!ActualQty
   If IsNull(TmpRs!LastFumgDate) = False Then
      MSHFlexGrid1.TextMatrix(I, 6) = Format(TmpRs!LastFumgDate, "dd-MMM-yyyy")
      MSHFlexGrid1.TextMatrix(I, 8) = DateDiff("d", CDate(MSHFlexGrid1.TextMatrix(I, 6)), TxtFumDate.Value)
   Else
      MSHFlexGrid1.TextMatrix(I, 8) = 0
   End If
   MSHFlexGrid1.TextMatrix(I, 7) = IIf(IsNull(TmpRs!EffectiveDays), "0", TmpRs!EffectiveDays)
   MSHFlexGrid1.TextMatrix(I, 9) = IIf(TmpRs!BillableFlag = "Y", "Yes", "No")
   MSHFlexGrid1.TextMatrix(I, 10) = TmpRs!Remark
   MSHFlexGrid1.TextMatrix(I, 11) = TmpRs!CommodityID
   MSHFlexGrid1.TextMatrix(I, 12) = IIf(IsNull(TmpRs!TreatmentType), "", TmpRs!TreatmentType)
   MSHFlexGrid1.TextMatrix(I, 13) = IIf(IsNull(TmpRs!AgreementDetailID), "", TmpRs!AgreementDetailID)
   MSHFlexGrid1.TextMatrix(I, 14) = IIf(IsNull(TmpRs!RateperMT), "", TmpRs!RateperMT)
   MSHFlexGrid1.TextMatrix(I, 15) = IIf(IsNull(TmpRs!Amount), "", TmpRs!Amount)
  
   MSHFlexGrid1.TextMatrix(0, 4) = Val(MSHFlexGrid1.TextMatrix(0, 4)) + Val(MSHFlexGrid1.TextMatrix(I, 4))
   MSHFlexGrid1.TextMatrix(0, 5) = Val(MSHFlexGrid1.TextMatrix(0, 5)) + Val(MSHFlexGrid1.TextMatrix(I, 5))
   MSHFlexGrid1.TextMatrix(0, 15) = Val(MSHFlexGrid1.TextMatrix(0, 15)) + Val(MSHFlexGrid1.TextMatrix(I, 15))
   TmpRs.MoveNext
Next
End Sub
Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsTxn_Fumigation.MoveNext
If RsTxn_Fumigation.EOF Then
   RsTxn_Fumigation.MoveLast
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
RsTxn_Fumigation.MovePrevious
If RsTxn_Fumigation.BOF Then
   RsTxn_Fumigation.MoveFirst
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
RsTxn_Fumigation.MoveLast
Call Indata
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False
End Sub
Private Sub FirstCmd_Click()
RsTxn_Fumigation.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
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
    CmbSLocationID.AddItem RsMst_Location!Locationname
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
Private Sub Grid_Head1()

MSHFlexGrid2.Clear
MSHFlexGrid2.Rows = 3
MSHFlexGrid2.FixedRows = 2
MSHFlexGrid2.Cols = 15
MSHFlexGrid2.WordWrap = True

MSHFlexGrid2.WordWrap = True
MSHFlexGrid2.TextMatrix(1, 0) = "Fumigation ID"
MSHFlexGrid2.TextMatrix(1, 1) = "Fumigation Date"
MSHFlexGrid2.TextMatrix(1, 2) = "Next Fumigation Date"
MSHFlexGrid2.TextMatrix(1, 3) = "Vendor Name"
MSHFlexGrid2.TextMatrix(1, 4) = "Location"
MSHFlexGrid2.TextMatrix(1, 5) = "CMP Name"
MSHFlexGrid2.TextMatrix(1, 6) = "Godown No"
MSHFlexGrid2.TextMatrix(1, 7) = "Rate"
MSHFlexGrid2.TextMatrix(1, 8) = "Total Qty"
MSHFlexGrid2.TextMatrix(1, 9) = "Billable Qty"
MSHFlexGrid2.TextMatrix(1, 10) = "Amount"
MSHFlexGrid2.TextMatrix(1, 11) = "Remark"
MSHFlexGrid2.TextMatrix(1, 12) = "Vendor Rate"
MSHFlexGrid2.TextMatrix(1, 13) = "Vendor Quantity"
MSHFlexGrid2.TextMatrix(1, 14) = "Mapping ID"

MSHFlexGrid2.ColWidth(0) = 1000
MSHFlexGrid2.ColWidth(1) = 1200
MSHFlexGrid2.ColWidth(2) = 1200
MSHFlexGrid2.ColWidth(3) = 2200
MSHFlexGrid2.ColWidth(4) = 2000
MSHFlexGrid2.ColWidth(5) = 1800
MSHFlexGrid2.ColWidth(6) = 1200
MSHFlexGrid2.ColWidth(7) = 1100
MSHFlexGrid2.ColWidth(8) = 1100
MSHFlexGrid2.ColWidth(9) = 1100
MSHFlexGrid2.ColWidth(10) = 1100
MSHFlexGrid2.ColWidth(11) = 2200
MSHFlexGrid2.ColWidth(12) = 1200
MSHFlexGrid2.ColWidth(13) = 1200
MSHFlexGrid2.ColWidth(14) = 1400

MSHFlexGrid2.RowHeight(1) = 800

End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Txn_Fumigation.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Txn_Fumigation.xls")
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
