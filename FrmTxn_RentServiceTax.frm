VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmTxn_RentServiceTax 
   Caption         =   "Godown Rent Service Tax"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   1365
      Left            =   120
      TabIndex        =   72
      Top             =   9000
      Visible         =   0   'False
      Width           =   15360
      Begin VB.ComboBox CmbSLessorName 
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
         ItemData        =   "FrmTxn_RentServiceTax.frx":0000
         Left            =   6555
         List            =   "FrmTxn_RentServiceTax.frx":0002
         Sorted          =   -1  'True
         TabIndex        =   80
         Top             =   375
         Width           =   3585
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
         ItemData        =   "FrmTxn_RentServiceTax.frx":0004
         Left            =   10200
         List            =   "FrmTxn_RentServiceTax.frx":0011
         TabIndex        =   79
         Top             =   375
         Width           =   1725
      End
      Begin VB.ComboBox CmbSLocation 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3480
         Sorted          =   -1  'True
         TabIndex        =   78
         Top             =   390
         Width           =   3000
      End
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   350
         Left            =   13905
         TabIndex        =   77
         Top             =   780
         Width           =   1000
      End
      Begin VB.CommandButton CmdGo 
         Caption         =   "Go"
         Height          =   345
         Left            =   12750
         TabIndex        =   76
         Top             =   780
         Width           =   1125
      End
      Begin VB.CheckBox ChkBatchNo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Check1"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   7470
         TabIndex        =   75
         Top             =   840
         Width           =   225
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Invoice Date"
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
         Left            =   1320
         TabIndex        =   74
         Top             =   240
         Width           =   1440
      End
      Begin VB.OptionButton Option1 
         Caption         =   "PI Date"
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
         TabIndex        =   73
         Top             =   240
         Value           =   -1  'True
         Width           =   1140
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
         Height          =   855
         Left            =   90
         TabIndex        =   81
         Top             =   1200
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   1508
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
      Begin MSComCtl2.DTPicker SFrom 
         Height          =   345
         Left            =   120
         TabIndex        =   82
         Top             =   735
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   609
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
         Format          =   113508353
         CurrentDate     =   39884
      End
      Begin MSComCtl2.DTPicker STo 
         Height          =   360
         Left            =   1575
         TabIndex        =   83
         Top             =   735
         Width           =   1335
         _ExtentX        =   2355
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
         Format          =   113508353
         CurrentDate     =   39884
      End
      Begin VB.Label Label31 
         Caption         =   "Lessor Name"
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
         Left            =   7800
         TabIndex        =   91
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label28 
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
         Left            =   10755
         TabIndex        =   90
         Top             =   120
         Width           =   555
      End
      Begin VB.Label Label23 
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
         Height          =   225
         Left            =   2107
         TabIndex        =   89
         Top             =   480
         Width           =   270
      End
      Begin VB.Label Label34 
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
         Height          =   225
         Left            =   4680
         TabIndex        =   88
         Top             =   120
         Width           =   930
      End
      Begin VB.Label Label22 
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
         Height          =   225
         Left            =   532
         TabIndex        =   87
         Top             =   480
         Width           =   510
      End
      Begin VB.Label Label18 
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
         Left            =   10920
         TabIndex        =   86
         Top             =   840
         Width           =   135
      End
      Begin VB.Label Label17 
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
         Left            =   8400
         TabIndex        =   85
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label Label35 
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
         Left            =   5640
         TabIndex        =   84
         Top             =   840
         Width           =   1710
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
      Height          =   4575
      Left            =   120
      TabIndex        =   43
      Top             =   0
      Width           =   13470
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
         ItemData        =   "FrmTxn_RentServiceTax.frx":0032
         Left            =   9915
         List            =   "FrmTxn_RentServiceTax.frx":003F
         Sorted          =   -1  'True
         TabIndex        =   67
         Top             =   1680
         Width           =   3360
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
         Height          =   615
         Left            =   1680
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   2640
         Width           =   11535
      End
      Begin VB.TextBox TxtBillingNo 
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
         Left            =   7095
         TabIndex        =   5
         Top             =   1695
         Width           =   1545
      End
      Begin VB.TextBox TxtSTRate 
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
         Left            =   4440
         TabIndex        =   7
         Top             =   2160
         Width           =   1425
      End
      Begin VB.TextBox TxtGodownRentSTID 
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
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   240
         Width           =   1260
      End
      Begin VB.ComboBox CmbLocation 
         BeginProperty Font 
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
         TabIndex        =   1
         Top             =   720
         Width           =   11595
      End
      Begin VB.ComboBox CmbLessorName 
         BeginProperty Font 
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
         TabIndex        =   2
         Top             =   1200
         Width           =   11595
      End
      Begin VB.TextBox TxtSTAmount 
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
         Left            =   7110
         Locked          =   -1  'True
         TabIndex        =   47
         Top             =   2160
         Width           =   1545
      End
      Begin VB.TextBox TxtBillingAmount 
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
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   2175
         Width           =   1425
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
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   4080
         Width           =   4935
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
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   4095
         Width           =   4935
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
         Height          =   615
         Left            =   1680
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   3360
         Width           =   11535
      End
      Begin VB.TextBox TxtBatchID 
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
         Left            =   12300
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   2175
         Width           =   915
      End
      Begin MSComCtl2.DTPicker TxtMonYear 
         Height          =   360
         Left            =   1680
         TabIndex        =   3
         Top             =   1695
         Width           =   1455
         _ExtentX        =   2566
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
         CustomFormat    =   "MMM-yyyy"
         Format          =   113508355
         CurrentDate     =   39884
      End
      Begin MSComCtl2.DTPicker TxtBillingDate 
         Height          =   360
         Left            =   4440
         TabIndex        =   4
         Top             =   1680
         Width           =   1440
         _ExtentX        =   2540
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
         Format          =   113508353
         CurrentDate     =   39884
      End
      Begin VB.TextBox TxtLocation 
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
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   75
         TabIndex        =   50
         Top             =   720
         Width           =   11535
      End
      Begin VB.TextBox TxtLessorName 
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
         Locked          =   -1  'True
         TabIndex        =   48
         Top             =   1200
         Visible         =   0   'False
         Width           =   11535
      End
      Begin MSComCtl2.DTPicker TxtFAGPostingDate 
         Height          =   360
         Left            =   9915
         TabIndex        =   68
         Top             =   2160
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
         Format          =   113508353
         CurrentDate     =   39884
      End
      Begin MSComCtl2.DTPicker TxtEntryDate 
         Height          =   360
         Left            =   5850
         TabIndex        =   69
         Top             =   240
         Width           =   1440
         _ExtentX        =   2540
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
         Format          =   113508353
         CurrentDate     =   39884
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
         Left            =   9915
         TabIndex        =   71
         Top             =   1680
         Width           =   3345
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Entry Date"
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
         TabIndex        =   70
         Top             =   300
         Width           =   930
      End
      Begin VB.Label Label6 
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
         Height          =   225
         Left            =   120
         TabIndex        =   66
         Top             =   2835
         Width           =   1215
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Tax Rate"
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
         TabIndex        =   65
         Top             =   2220
         Width           =   960
      End
      Begin VB.Label Label2 
         Caption         =   "Tax Amount"
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
         Left            =   5910
         TabIndex        =   64
         Top             =   2205
         Width           =   1350
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Rent ST ID"
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
         TabIndex        =   63
         Top             =   300
         Width           =   975
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
         Left            =   120
         TabIndex        =   62
         Top             =   773
         Width           =   930
      End
      Begin VB.Label Label19 
         Caption         =   "Lessor Name"
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
         Top             =   1253
         Width           =   1200
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Month Year"
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
         TabIndex        =   60
         Top             =   1755
         Width           =   1185
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Billing No"
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
         Left            =   5910
         TabIndex        =   59
         Top             =   1755
         Width           =   1020
      End
      Begin VB.Label Label4 
         Caption         =   "Billing Amount"
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
         TabIndex        =   58
         Top             =   2220
         Width           =   1590
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Billing Date"
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
         Left            =   3150
         TabIndex        =   57
         Top             =   1755
         Width           =   1215
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
         TabIndex        =   56
         Top             =   4140
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
         Left            =   6840
         TabIndex        =   55
         Top             =   4155
         Width           =   975
      End
      Begin VB.Label Label10 
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
         Height          =   225
         Left            =   8760
         TabIndex        =   54
         Top             =   1755
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "FAG Remark"
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
         Left            =   120
         TabIndex        =   53
         Top             =   3555
         Width           =   1215
      End
      Begin VB.Label Label12 
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
         Height          =   480
         Left            =   8760
         TabIndex        =   52
         Top             =   2160
         Width           =   1200
      End
      Begin VB.Label Label13 
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
         Left            =   11370
         TabIndex        =   51
         Top             =   2235
         Width           =   750
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1155
      Left            =   120
      TabIndex        =   14
      Top             =   7680
      Width           =   13470
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Print"
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
         Left            =   11160
         TabIndex        =   92
         Top             =   240
         Width           =   2055
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
         Left            =   255
         TabIndex        =   19
         Top             =   225
         Width           =   2175
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
         Left            =   2430
         TabIndex        =   0
         Top             =   225
         Width           =   2175
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
         Left            =   4605
         TabIndex        =   10
         Top             =   225
         Width           =   2175
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
         Left            =   6780
         TabIndex        =   11
         Top             =   225
         Width           =   2175
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
         Left            =   8955
         TabIndex        =   18
         Top             =   225
         Width           =   2175
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
         Left            =   255
         TabIndex        =   13
         Top             =   600
         Width           =   2175
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
         Left            =   2430
         TabIndex        =   17
         Top             =   600
         Width           =   2175
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
         Left            =   4605
         TabIndex        =   16
         Top             =   600
         Width           =   2175
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
         Left            =   6780
         TabIndex        =   15
         Top             =   600
         Width           =   2175
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
         Left            =   8955
         TabIndex        =   12
         Top             =   600
         Width           =   2175
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2895
      Left            =   120
      TabIndex        =   20
      Top             =   4680
      Width           =   13470
      _ExtentX        =   23760
      _ExtentY        =   5106
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Claim Details"
      TabPicture(0)   =   "FrmTxn_RentServiceTax.frx":0060
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Agreement Details"
      TabPicture(1)   =   "FrmTxn_RentServiceTax.frx":007C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame3 
         Height          =   2415
         Left            =   120
         TabIndex        =   41
         Top             =   360
         Width           =   13230
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
            Height          =   2175
            Left            =   150
            TabIndex        =   42
            Top             =   150
            Width           =   12975
            _ExtentX        =   22886
            _ExtentY        =   3836
            _Version        =   393216
            FixedCols       =   0
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
      End
      Begin VB.Frame Frame2 
         Height          =   2460
         Left            =   -74850
         TabIndex        =   39
         Top             =   345
         Width           =   13095
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   2220
            Left            =   90
            TabIndex        =   40
            Top             =   165
            Width           =   12855
            _ExtentX        =   22675
            _ExtentY        =   3916
            _Version        =   393216
            FixedCols       =   0
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
      End
      Begin VB.TextBox TxtBalAmt 
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
         IMEMode         =   3  'DISABLE
         Left            =   -64365
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   750
         Width           =   1815
      End
      Begin VB.TextBox TxtTIAmount 
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
         IMEMode         =   3  'DISABLE
         Left            =   -69810
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   750
         Width           =   1815
      End
      Begin VB.TextBox TxtTIDate 
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
         IMEMode         =   3  'DISABLE
         Left            =   -71625
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   750
         Width           =   1815
      End
      Begin VB.TextBox TxtTINo 
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
         IMEMode         =   3  'DISABLE
         Left            =   -73515
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   750
         Width           =   1890
      End
      Begin VB.TextBox TxtTISTAmt 
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
         IMEMode         =   3  'DISABLE
         Left            =   -67995
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   750
         Width           =   1815
      End
      Begin VB.TextBox TxtRecAmt 
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
         IMEMode         =   3  'DISABLE
         Left            =   -66180
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   750
         Width           =   1815
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
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   -62550
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   750
         Width           =   1935
      End
      Begin VB.TextBox Text7 
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
         Left            =   -73515
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   1965
         Width           =   10950
      End
      Begin VB.TextBox Text6 
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
         Left            =   -73515
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1440
         Width           =   10965
      End
      Begin VB.Label Label16 
         Caption         =   "Balance Amount"
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
         Left            =   -64215
         TabIndex        =   38
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label27 
         Caption         =   "TI Date"
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
         Left            =   -71055
         TabIndex        =   37
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label15 
         Caption         =   "TI Amount"
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
         Left            =   -69375
         TabIndex        =   36
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label21 
         Caption         =   "TI No"
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
         Left            =   -72765
         TabIndex        =   35
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label20 
         Caption         =   "TI ST Amount"
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
         Left            =   -67695
         TabIndex        =   34
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label14 
         Caption         =   "Received Amount"
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
         Left            =   -66075
         TabIndex        =   33
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label26 
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
         Height          =   255
         Left            =   -62055
         TabIndex        =   32
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label47 
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
         TabIndex        =   31
         Top             =   2025
         Width           =   975
      End
      Begin VB.Label Label46 
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
         TabIndex        =   30
         Top             =   1500
         Width           =   1695
      End
   End
End
Attribute VB_Name = "FrmTxn_RentServiceTax"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mLocationID, mSM_Prefix, mCommodityID, mClientIDRow, mSLocationID, msSM_prefix, mSClientIDRow, mSCommodityID, mLessorID, tmp1 As Variant
Dim Edit_Flg As Variant
Dim mStatusFlag As Variant
Dim FileName As String
Const strChecked = "þ"
Const strUnChecked = "q"
Dim mClaimID As Variant


Private Sub AddCmd_Click()
Edit_Flg = "A"

Frame1.Enabled = True
Frame2.Enabled = False
CmdPrint.Enabled = False

TxtEntryDate.Value = Date

CmbLocation.Visible = True
TxtLocation.Visible = False

CmbLessorName.Visible = True
TxtLessorName.Visible = False

TxtGodownRentSTID.Text = ""
CmbLocation.Text = ""
CmbLessorName.Text = ""
TxtRemark.Locked = False
TxtCreated = ""
TxtUpdated = ""

TxtMonYear.Value = Date
TxtMonYear.Day = 1
TxtBillingDate.Value = Date
TxtBillingAmount.Text = ""
TxtBillingNo.Text = ""
TxtSTRate.Text = ""
TxtSTAmount.Text = ""
TxtRemark.Text = "BEING GST PAYMENT FOR THE MONTH"

Call Grid_Head2
Call Grid_Head1

CmbStatus.Visible = False
TxtStatus.Visible = True
CmbStatus.Text = "Pending"
TxtStatus.Text = "Pending"
mStatusFlag = "P"
TxtBatchID.Visible = False
Label13.Visible = False
TxtFAGRemarks.Locked = True
TxtFAGRemarks = ""
TxtFAGPostingDate.Visible = False
Label12.Visible = False

Call GButtonLock(Me, False)

End Sub

Private Sub Grid_Head()
With MSHFlexGrid2
     .Clear
     .Rows = 2
     .Cols = 15
     .AllowUserResizing = flexResizeBoth
     .Font.Size = 10
     .WordWrap = True
     '.RowHeight(1) = 700
     .TextMatrix(0, 0) = "GodownRentST ID"
     .TextMatrix(0, 1) = "Entry Date"
     .TextMatrix(0, 2) = "Location Name"
     .TextMatrix(0, 3) = "Lessor Name"
     .TextMatrix(0, 4) = "Month Year"
     .TextMatrix(0, 5) = "Billing Date"
     .TextMatrix(0, 6) = "Billing No"
     .TextMatrix(0, 7) = "Billing Amount"
     .TextMatrix(0, 8) = "Tax Rate"
     .TextMatrix(0, 9) = "Tax Amount"
     .TextMatrix(0, 10) = "Status"
     .TextMatrix(0, 11) = "Batch ID"
     .TextMatrix(0, 12) = "Remarks"
     .TextMatrix(0, 13) = "FAG Remarks"
     .TextMatrix(0, 14) = "FAGPosting Date"

     .ColWidth(0) = 800
     .ColWidth(1) = 1500
     .ColWidth(2) = 3000
     .ColWidth(3) = 1500
     .ColWidth(4) = 1200
     .ColWidth(5) = 1200
     .ColWidth(6) = 1500
     .ColWidth(7) = 1200

End With
End Sub
Private Sub Grid_Head1()
With MSHFlexGrid3
     .Clear
     .Rows = 3
     .FixedRows = 2
     .Cols = 5
     .AllowUserResizing = flexResizeBoth
     .Font.Size = 10
     .WordWrap = True
     .RowHeight(1) = 500
     
     .TextMatrix(1, 0) = "Select"
     .TextMatrix(1, 1) = "Agreement ID"
     .TextMatrix(1, 2) = "Agreement Date"
     .TextMatrix(1, 3) = "Lease Amount (No Tax)"
     .TextMatrix(1, 4) = "Lease Amount (With Tax)"
     
     .ColWidth(0) = 1000
     .ColWidth(1) = 2000
     .ColWidth(2) = 2000
     .ColWidth(3) = 3000
     .ColWidth(4) = 3000
End With
End Sub
Private Sub Grid_Head2()
With MSHFlexGrid1
     .Clear
     .Rows = 2
     .Cols = 4
     .AllowUserResizing = flexResizeBoth
     .Font.Size = 10
     .WordWrap = True
     .RowHeight(0) = 700
     .TextMatrix(0, 0) = "Claim ID"
     .TextMatrix(0, 1) = "Claim Date"
     .TextMatrix(0, 2) = "Amount"
     .TextMatrix(0, 3) = "Status"
     
     .ColWidth(0) = 2000
     .ColWidth(1) = 2500
     .ColWidth(2) = 2500
     .ColWidth(3) = 2500
End With
End Sub
Private Sub CmbLocation_GotFocus()
tmp = CmbLocation.Text

If Gen_WcLocation = "" Then
   Call TableMst_Location("")
Else
   Call TableMst_Location("where " & Gen_WcLocation)
End If

CmbLocation.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No record found!!! "
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbLocation.AddItem RsMst_Location!LocationName
    RsMst_Location.MoveNext
Next
CmbLocation.Text = tmp
End Sub


Private Sub CmbLocation_LostFocus()
If CmbLocation.Text = "" Then
   mLocationID = 0
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocation.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection!!! "
   CmbLocation.SetFocus
   Exit Sub
End If
If mLocationID <> RsMst_Location!LocationID Then
   'CmbNonNcdexClient.Text = ""
End If
mLocationID = RsMst_Location!LocationID
mSM_Prefix = RsMst_Location!SM_Prefix

If CmbLessorName.Text <> "" Then
   Call GetAgreementDetails
End If
End Sub


Private Sub CmbSLocation_GotFocus()
tmp = CmbSLocation.Text

If Gen_WcLocation = "" Then
   Call TableMst_Location("")
Else
   Call TableMst_Location("where " & Gen_WcLocation)
End If

CmbSLocation.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No record found!!! "
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
   mSLocationID = 0
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbSLocation.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection!!! "
   CmbSLocation.SetFocus
   Exit Sub
End If
mSLocationID = RsMst_Location!LocationID
End Sub

Private Sub CmdSearch_Click()
Dim wc As String
Dim mFromDt As Date
Dim mToDt As Date
wc = ""
Call Grid_Head1
TxtSStartDate.Value = DateSerial(Year(TxtSStartDate.Value), Month(TxtSStartDate.Value), 1)
TxtSEndDate.Value = DateSerial(Year(TxtSEndDate.Value), Month(TxtSEndDate.Value) + 1, 0)

tmp = "Select Txn_IEM.GodownRentSTID, Txn_IEM.LocationID, ML.SM_Prefix, ML.LocationName, Txn_IEM.CommodityID, MCO.Commodity, Txn_IEM.ClientID, MC.ClientName, Txn_IEM.StartDate, Txn_IEM.EndDate, Txn_IEM.Concurrency, Txn_IEM.CreatedBy "
tmp = tmp & " From  Txn_GodownRentST as Txn_IEM"
tmp = tmp & " Inner Join  Mst_Location as ML on ML.LocationID = Txn_IEM.LocationID"
tmp = tmp & " Left Join Mst_Client as MC on MC.ClientID = Txn_IEM.ClientID"
tmp = tmp & " Left Join Mst_Commodity as MCO on MCO.CommodityID = Txn_IEM.CommodityID"
''tmp = tmp & " Where Txn_IEM.StartDate > = '" & Format(TxtSStartDate.Value, Gen_DatFormat) & "' And Txn_IEM.EndDate <='" & Format(TxtSEndDate.Value, Gen_DatFormat) & "'"
If OptStartDt.Value = True Then
   'mFromDt
   wc = " Where Txn_IEM.StartDate Between '" & Format(TxtSStartDate.Value, Gen_DatFormat) & "' And '" & Format(TxtSEndDate, Gen_DatFormat) & "'"
  ' Wc = " Where Txn_IEM.StartDate Between " & TxtSStartDate.Value & " And " & TxtSEndDate & " "
ElseIf OptEndDt.Value = True Then
   wc = " Where Txn_IEM.EndDate Between '" & Format(TxtSStartDate.Value, Gen_DatFormat) & "' And '" & Format(TxtSEndDate, Gen_DatFormat) & "'"
End If
If CmbSLocation.Text <> "" Then
   wc = wc & " And Txn_IEM.LocationID =" & mSLocationID
End If
If CmbSClient.Text <> "" Then
   wc = wc & " And Txn_IEM.ClientID =" & mSClientIDRow
End If
If CmbSCommodity.Text <> "" Then
   wc = wc & " And Txn_IEM.CommodityID =" & mSCommodityID
End If
If TxtSGodownRentSTID.Text <> "" Then
   wc = wc & " And Txn_IEM.GodownRentSTID =" & Val(TxtSGodownRentSTID)
End If

tmp = tmp & wc
tmp = tmp & " Order by  Txn_IEM.GodownRentSTID"
Call Tmp2Table

If Not TmpRs2.EOF Then
   Label22.Caption = TmpRs2.RecordCount
   With MSHFlexGrid2
        For I = 1 To TmpRs2.RecordCount
            .TextMatrix(I, 0) = TmpRs2!GodownRentSTID
            .TextMatrix(I, 1) = TmpRs2!LocationName
            .TextMatrix(I, 2) = IIf(IsNull(TmpRs2!ClientName), "", TmpRs2!ClientName)
            .TextMatrix(I, 3) = IIf(IsNull(TmpRs2!Commodity), "", TmpRs2!Commodity)
            .TextMatrix(I, 4) = Format(TmpRs2!StartDate, "dd-MMM-yyyy")
            .TextMatrix(I, 5) = Format(TmpRs2!EndDate, "dd-MMM-yyyy")
            TmpRs2.MoveNext
            .Rows = .Rows + 1
        Next
   End With
End If
End Sub



Private Sub CmdGo_Click()
Dim wc As Variant
wc = ""
If SFrom.Value > STo.Value Then
   MsgBox "'From Date' should be less than 'To Date' !!"
   SFrom.SetFocus
   Exit Sub
End If

Call Grid_Head

mTotalDays = Day(DateSerial(Year(STo.Value), Month(STo.Value) + 1, 0))

If Option2.Value = True Then
   mSFromDate = Format(Year(SFrom.Value) & "-" & Month(SFrom.Value) & "-" & "01", "yyyy-mm-dd")
   mSToDate = Format(Year(STo.Value) & "-" & Month(STo.Value) & "-" & mTotalDays, "yyyy-mm-dd")
Else
   mSFromDate = SFrom.Value  'Format(Year(SFrom.Value) & "-" & Month(SFrom.Value) & "-" & "01", "yyyy-mm-dd")
   mSToDate = STo.Value  'Format(Year(STo.Value) & "-" & Month(STo.Value) & "-" & "01", "yyyy-mm-dd")
End If


tmp = "Select TGR.GodownRentSTID,TGR.GodownRentSTDate,Ml.LocationName,MS.LessorName,TGR.MonthYear,TGR.BillDate,TGR.BillNo,TGR.BillAmount,TGR.STRate,TGR.STAmount,TGR.Flag,TGR.BatchID,TGR.Remarks,TGR.FAGRemarks,TGR.FAGPostingDate"
tmp = tmp & " from Txn_GodownRentST TGR"
tmp = tmp & " Inner join Mst_location ML on ML.LocationID=TGR.LocationID"
tmp = tmp & " Inner join Mst_Lessor MS on MS.LessorID=TGR.LessorID"

'If Gen_DBType = "MDB" Then
'   wc = " Where TTI.PIDate >= #" & SFrom.Value & "# And TTI.PIDate <= #" & STo.Value & "# "
'Else
'   wc = " Where TTI.PIDate >= '" & SFrom.Value & "' And TTI.PIDate <= '" & STo.Value & "' "
'End If
            

If Gen_DBType = "MDB" Then
   If Option1.Value = True Then
      wc = " Where TGR.GodownRentSTDate >= #" & SFrom.Value & "# And TGR.GodownRentSTDate <= #" & STo.Value & "# "
   Else
      wc = " Where TGR.MonthYear >= #" & mSFromDate & "# And TGR.MonthYear <= #" & mSToDate & "# "
   End If
Else
'   wc = " Where TTI.PIDate >= '" & SFrom.Value & "' And TTI.PIDate <= '" & STo.Value & "' "
   If Option1.Value = True Then
      wc = " Where TGR.GodownRentSTDate >= '" & SFrom.Value & "' And TGR.GodownRentSTDate <= '" & STo.Value & "' "
   Else
      wc = " Where TGR.MonthYear >= '" & mSFromDate & "' And TGR.MonthYear  <= '" & mSToDate & "' "
   End If
End If
            
If CmbSLocation.Text <> "" Then
   wc = wc & " And  ML.LocationName = '" & CmbSLocation.Text & "'"
Else
   If Gen_WcLocation <> "" Then
      wc = wc & " And ML." & Gen_WcLocation
   End If
End If
If CmbSLessorName.Text <> "" Then
   wc = wc & " And MS.LessorName = '" & CmbSLessorName.Text & "'"
End If

If CmbSStatus.Text <> "" Then
   wc = wc & " And  TGR.Flag = '" & Left(CmbSStatus.Text, 1) & "'"
End If
           
If ChkBatchNo.Value = 0 Then
   wc = wc & " And isnull(TGR.BatchID,0) = 0 " '" & Val(TxtSAINo)
End If
                      
tmp1 = " Order By TGR.GodownRentSTID"
tmp = tmp & " " & wc & " " & tmp1

Call TmpTable

If TmpRs.RecordCount > 0 Then
   For I = 1 To TmpRs.RecordCount
       MSHFlexGrid2.Rows = TmpRs.RecordCount + 2
       For C = 0 To TmpRs.Fields.Count - 1
           MSHFlexGrid2.TextMatrix(I, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
       Next
       
       MSHFlexGrid2.TextMatrix(I, 1) = Format(MSHFlexGrid2.TextMatrix(I, 1), "dd-mmm-yyyy")
       MSHFlexGrid2.TextMatrix(I, 4) = Format(MSHFlexGrid2.TextMatrix(I, 4), "mmm-yyyy")
       MSHFlexGrid2.TextMatrix(I, 5) = Format(MSHFlexGrid2.TextMatrix(I, 5), "dd-mmm-yyyy")
       If MSHFlexGrid2.TextMatrix(I, 10) = "P" Then
          MSHFlexGrid2.TextMatrix(I, 10) = "Pending"
       ElseIf MSHFlexGrid2.TextMatrix(I, 10) = "A" Then
          MSHFlexGrid2.TextMatrix(I, 10) = "Approved"
       ElseIf MSHFlexGrid2.TextMatrix(I, 10) = "R" Then
          MSHFlexGrid2.TextMatrix(I, 10) = "Rejected"
       End If
       If MSHFlexGrid2.TextMatrix(I, 13) <> "" Then
          If InStr("" & Trim(MSHFlexGrid2.TextMatrix(I, 13)) & "", "Claim ID :- " & Val(MSHFlexGrid2.TextMatrix(I, 0)) & "") = 0 Then
             MSHFlexGrid2.TextMatrix(I, 13) = Trim(MSHFlexGrid2.TextMatrix(I, 13)) & " Claim ID :- " & Val(MSHFlexGrid2.TextMatrix(I, 0)) & " "
          End If
       End If
       TmpRs.MoveNext
   Next
End If
Label18.Caption = TmpRs.RecordCount

End Sub

Private Sub CmdPrint_Click()

Gen_ExcelPassword = "1234"
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "RentServiceTax" & Trim(TxtGodownRentSTID) & ".xls")
FileName = Gen_mTimeStamp & Gen_UserName & "RentServiceTax" & Trim(TxtGodownRentSTID) & ".xls"
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
    
    oWS.Range("a:a").ColumnWidth = 11.43 '.157 '10.86
    oWS.Range("b:b").ColumnWidth = 8.29 '17.71
    oWS.Range("c:c").ColumnWidth = 11 '17.14
    oWS.Range("d:d").ColumnWidth = 7 '12.86
    oWS.Range("e:e").ColumnWidth = 8.57 '33 '32.14
    oWS.Range("f:f").ColumnWidth = 8
    oWS.Range("g:g").ColumnWidth = 7.86
    oWS.Range("h:h").ColumnWidth = 7
    oWS.Range("i:i").ColumnWidth = 10.57
      
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
    oWS.Cells(mRow, C + 7) = TxtGodownRentSTID
    oWS.Cells(mRow, C + 7).Font.Bold = True
    oWS.Cells(mRow, C + 7).Font.Size = 10
    'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
       
    tmp = "h" & mRow & ":" & "h" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 8) = "Date"
    oWS.Cells(mRow, C + 8).Font.Bold = True
    oWS.Cells(mRow, C + 8).Font.Size = 10
    
    
    tmp = "i" & mRow & ":" & "i" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 9) = "'" & Format(TxtEntryDate.Value, "dd-mmm-yyyy")
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
    oWS.Cells(mRow, C + 3) = TxtLessorName.Text
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
    oWS.Cells(mRow, C + 3) = "'" & Format(TxtMonYear.Value, "MMM-yyyy")
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
    oWS.Cells(mRow, C + 3) = "'" & TxtBillingNo.Text
    oWS.Cells(mRow, C + 3).Font.Size = 10
    'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
    
    tmp = "e" & mRow & ":" & "f" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 5) = "Invoice Date"
    oWS.Cells(mRow, C + 5).Font.Bold = True
    oWS.Cells(mRow, C + 5).Font.Size = 10
    'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
    
    tmp = "g" & mRow & ":" & "i" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 7) = "'" & Format(TxtBillingDate.Value, "dd-MMM-yyyy")
    oWS.Cells(mRow, C + 7).Font.Size = 10
    'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
    
    mRow = mRow + 1
    tmp = "a" & mRow & ":" & "a" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 1) = "Remark : "
    oWS.Cells(mRow, C + 1).Font.Bold = True
    oWS.Cells(mRow, C + 1).Font.Size = 10
    
    tmp = "b" & mRow & ":" & "i" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 2) = TxtRemark.Text
    oWS.Cells(mRow, C + 2).Font.Bold = False
    oWS.Cells(mRow, C + 2).Font.Size = 10
    
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
    oWS.Cells(mRow, C + 1) = "Claim ID"
    oWS.Cells(mRow, C + 1).Font.Bold = True
    oWS.Cells(mRow, C + 1).Font.Size = 10
    oWS.Cells(mRow, C + 1).RowHeight = 27.75
    oWS.Cells(mRow, C + 1).WrapText = True
    'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
    
    tmp = "b" & mRow & ":" & "c" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 2) = "Claim Date"
    oWS.Cells(mRow, C + 2).Font.Bold = True
    oWS.Cells(mRow, C + 2).Font.Size = 10
    oWS.Cells(mRow, C + 2).RowHeight = 27.75
    oWS.Cells(mRow, C + 2).WrapText = True
    'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
    
    tmp = "d" & mRow & ":" & "f" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 4) = "Claim Status"
    oWS.Cells(mRow, C + 4).Font.Bold = True
    oWS.Cells(mRow, C + 4).Font.Size = 10
    oWS.Cells(mRow, C + 4).RowHeight = 27.75
    'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
    
    tmp = "g" & mRow & ":" & "i" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 7) = "Claim Amount"
    oWS.Cells(mRow, C + 7).Font.Bold = True
    oWS.Cells(mRow, C + 7).Font.Size = 10
    oWS.Cells(mRow, C + 7).RowHeight = 27.75
    'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
       
    
    For I = 1 To MSHFlexGrid1.Rows - 2
        mRow = mRow + 1 ' + 6
        C = 0
        
        oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, 0)
        oWS.Cells(mRow, C + 1).Font.Size = 10
        tmp = "b" & mRow & ":" & "c" & mRow
        oWS.Range(tmp).Merge
    
        oWS.Cells(mRow, C + 2) = Format(MSHFlexGrid1.TextMatrix(I, 1), "dd-mmm-yyyy")
        oWS.Cells(mRow, C + 2).Font.Size = 10
        tmp = "d" & mRow & ":" & "f" & mRow
        oWS.Range(tmp).Merge
        oWS.Cells(mRow, C + 4) = MSHFlexGrid1.TextMatrix(I, 3) ' Format(MSHFlexGrid2.TextMatrix(I, c + 3), "dd-mmm-yyyy")
        oWS.Cells(mRow, C + 4).Font.Size = 10
        tmp = "g" & mRow & ":" & "i" & mRow
        oWS.Range(tmp).Merge
        oWS.Cells(mRow, C + 7) = MSHFlexGrid1.TextMatrix(I, 2)
        oWS.Cells(mRow, C + 7).Font.Size = 10
    Next I
    
    mRow = mRow + 1
    tmp = "a" & mRow & ":" & "f" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 1) = "Bill Amount"
    oWS.Cells(mRow, C + 1).Font.Bold = True
    oWS.Cells(mRow, C + 1).Font.Size = 10
    oWS.Cells(mRow, C + 1).HorizontalAlignment = -4108
    'oWS.Cells(mRow, c + j).Font.Size = 10
    
    tmp = "g" & mRow & ":" & "i" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 7) = TxtBillingAmount.Text
    oWS.Cells(mRow, C + 7).Font.Bold = True
    oWS.Cells(mRow, C + 7).Font.Size = 10
    'oWS.Cells(mRow, c + j).Font.Size = 10
    
    mRow = mRow + 1
    tmp = "a" & mRow & ":" & "f" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 1) = "Servce Tax Rate"
    oWS.Cells(mRow, C + 1).Font.Bold = True
    oWS.Cells(mRow, C + 1).Font.Size = 10
    oWS.Cells(mRow, C + 1).HorizontalAlignment = -4108
    'oWS.Cells(mRow, c + j).Font.Size = 10
    
    tmp = "g" & mRow & ":" & "i" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 7) = TxtSTRate.Text
    oWS.Cells(mRow, C + 7).Font.Bold = True
    oWS.Cells(mRow, C + 7).Font.Size = 10
    'oWS.Cells(mRow, c + j).Font.Size = 10
    
    mRow = mRow + 1
    tmp = "a" & mRow & ":" & "f" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 1) = "Total Payable by NCMSL"
    oWS.Cells(mRow, C + 1).Font.Bold = True
    oWS.Cells(mRow, C + 1).Font.Size = 10
    oWS.Cells(mRow, C + 1).HorizontalAlignment = -4108
    'oWS.Cells(mRow, c + j).Font.Size = 10
    
    tmp = "g" & mRow & ":" & "i" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 7) = TxtSTAmount.Text
    oWS.Cells(mRow, C + 7).Font.Bold = True
    oWS.Cells(mRow, C + 7).Font.Size = 10
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
    oWS.Cells(mRow, C + 8) = TxtClaimID
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
    oWS.Cells(mRow, C + 3) = "Expense From " & Format(TxtExpenseFrom, "MMM-yyyy") & " To " & Format(TxtExpenseFrom, "MMM-yyyy") & " "
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
Private Sub DeleteCmd_Click()
Dim ans As Variant
On Error GoTo msgError

If LCase(TxtStatus) = "approved" Or LCase(TxtStatus) = "rejected" Then
   MsgBox "Status is " & TxtStatus & ". Can't edit!!!"
   Exit Sub
End If


ans = MsgBox("Do you really want to DELETE this record???", vbYesNo)
If ans = vbYes Then
   
   tmp = "Delete From Txn_GodownRentSTAgreeDetail Where GodownRentSTID = " & Val(TxtGodownRentSTID) & " "
   Call TmpTable
   
   tmp = "Delete From Txn_GodownRentSTClaimDetail Where GodownRentSTID = " & Val(TxtGodownRentSTID) & " "
   Call TmpTable


   RsTxn_GodownRentST.Delete
   RsTxn_GodownRentST.Update
   If RsTxn_GodownRentST.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsTxn_GodownRentST.MoveNext
   If RsTxn_GodownRentST.EOF() Then
      RsTxn_GodownRentST.MoveLast
   End If
   Call Indata
End If
Exit Sub
msgError:
MsgBox "Child record Present "
RsMst_State.CancelUpdate
End Sub


Private Sub EditCmd_Click()
If LCase(TxtStatus) = "approved" Or LCase(TxtStatus) = "rejected" Then
   MsgBox "Status is " & TxtStatus & ". Can't edit!!!"
   Exit Sub
End If

Edit_Flg = "E"

Call GButtonLock(Me, False)

Frame1.Enabled = True
Frame2.Enabled = True
CmdPrint.Enabled = False
CmbLocation.Text = TxtLocation.Text
'CmbLocation.Visible = True
'TxtLocation.Visible = False

CmbLessorName.Text = TxtLessorName.Text
'CmbLessorName.Visible = True
'TxtLessorName.Visible = False
TxtMonYear.Enabled = False
CmbStatus.Visible = True
TxtStatus.Visible = False
CmbStatus.Text = TxtStatus
TxtFAGRemarks.Locked = True
TxtRemark.Locked = False


End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub FirstCmd_Click()
RsTxn_GodownRentST.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub

Private Sub Form_Load()

Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Call GButtonLockAD(Me)
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If
Call Grid_Head1
Call Grid_Head2
SFrom.Value = Date - 30
STo.Value = Date
Call TableTxn_GodownRentST("  Order by GodownRentSTID ")
If RsTxn_GodownRentST.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
Call Indata
Call GButtonLock(Me, True)
End Sub
Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Txn_GodownRentST.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Txn_GodownRentST.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
'ProgressBar1.Value = 0
'ProgressBar1.Max = MSHFlexGrid2.Rows

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
        MSHFlexGrid2.row = I
        MSHFlexGrid2.Col = C
                
        If MSHFlexGrid2.CellBackColor = vbRed Then
           oWS.Range(oWS.Cells(mRow, C), oWS.Cells(mRow, C)).Interior.Color = vbRed
        ElseIf MSHFlexGrid2.CellBackColor = vbGreen Then
           oWS.Range(oWS.Cells(mRow, C), oWS.Cells(mRow, C)).Interior.Color = vbGreen
        ElseIf MSHFlexGrid2.CellBackColor = vbYellow Then
           oWS.Range(oWS.Cells(mRow, C), oWS.Cells(mRow, C)).Interior.Color = vbYellow
        End If
              
    Next
'    ProgressBar1.Value = ProgressBar1.Value + 1
Next
'ProgressBar1.Value = ProgressBar1.Max
oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub

Private Sub CmbLessorName_GotFocus()
If CmbLocation.Text = "" Then
   MsgBox "Please select Location!!"
   CmbLocation.SetFocus
   Exit Sub
End If
tmp1 = CmbLessorName.Text

mLessorID = ""
'Call TableMst_Lessor("")

tmp = "Select Distinct LS.LessorID,LS.LessorName from Txn_GodownLessorAgreeDetail TLG"
tmp = tmp & " inner join mst_Lessor LS on LS.LessorID=TLG.LessorID"
tmp = tmp & " Inner join Mst_Godown MG on MG.GodownID=TLG.GodownID"
tmp = tmp & " Inner join Mst_CMP MC on MC.CMPID=MG.CMPID"
tmp = tmp & " Where MC.LocationID=" & mLocationID

Call Tmp5Table

CmbLessorName.Clear
If TmpRs5.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To TmpRs5.RecordCount
    CmbLessorName.AddItem TmpRs5!LessorName
    TmpRs5.MoveNext
Next
CmbLessorName.Text = tmp1
End Sub

Private Sub CmbLessorName_LostFocus()
If CmbLessorName.Text = "" Then
   Exit Sub
End If
TmpRs5.MoveFirst
TmpRs5.Find "LessorName = '" & CmbLessorName.Text & "'"
If TmpRs5.EOF Then
   MsgBox "Invalid selection "
   CmbLessorName.SetFocus
   Exit Sub
End If
mLessorID = TmpRs5!LessorID
If CmbLocation.Text <> "" Then
   Call GetAgreementDetails
End If

End Sub


Private Sub LastCmd_Click()
RsTxn_GodownRentST.MoveLast
Call Indata
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False
End Sub

Private Sub MSHFlexGrid1_Click()
Dim I As Variant
I = MSHFlexGrid1.RowSel
If I <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(I, 0) = "" Then Exit Sub
RsTxn_GodownRentST.MoveFirst
RsTxn_GodownRentST.Find "GodownRentSTID= " & Val(MSHFlexGrid1.TextMatrix(I, 0)) & ""
If Not RsTxn_GodownRentST.EOF Then
   Call Indata
   Frame0.Visible = False
End If
End Sub

Private Sub MSHFlexGrid2_Click()
Dim I As Variant
I = MSHFlexGrid2.RowSel
If I <= 0 Then Exit Sub
If MSHFlexGrid2.TextMatrix(I, 0) = "" Then Exit Sub
RsTxn_GodownRentST.MoveFirst
RsTxn_GodownRentST.Find "GodownRentSTID= " & Val(MSHFlexGrid2.TextMatrix(I, 0)) & ""
If Not RsTxn_GodownRentST.EOF Then
   Call Indata
   Frame0.Visible = False
End If
End Sub

Private Sub MSHFlexGrid3_Click()
If SaveCmd.Enabled = False Then Exit Sub

Dim r As Variant
r = MSHFlexGrid3.RowSel
If r < 2 Then Exit Sub

mSelRow = MSHFlexGrid3.row
mSelCol = MSHFlexGrid3.Col

With MSHFlexGrid3
     If mSelCol = 0 Then
        If .TextMatrix(mSelRow, mSelCol) = "" Then Exit Sub
        If .TextMatrix(mSelRow, 4) = "" Then Exit Sub
        
        mSelCol = 0
        If .TextMatrix(mSelRow, mSelCol) = strUnChecked Then
           .TextMatrix(mSelRow, mSelCol) = strChecked
            TxtBillingAmount.Text = Val(TxtBillingAmount.Text) + Val(.TextMatrix(mSelRow, 4))
            TxtSTAmount.Text = Format(TxtBillingAmount * (Val(TxtSTRate) / 100), "#####0.00")
        Else
           .TextMatrix(mSelRow, mSelCol) = strUnChecked
           TxtBillingAmount.Text = Val(TxtBillingAmount.Text) - Val(.TextMatrix(mSelRow, 4))
           TxtSTAmount.Text = Format(TxtBillingAmount * (Val(TxtSTRate) / 100), "#####0.00")
           If Val(TxtBillingAmount.Text) = 0 Then TxtBillingAmount.Text = ""
        End If
     End If
End With

End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsTxn_GodownRentST.MoveNext
If RsTxn_GodownRentST.EOF Then
   RsTxn_GodownRentST.MoveLast
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
RsTxn_GodownRentST.MovePrevious
If RsTxn_GodownRentST.BOF Then
   RsTxn_GodownRentST.MoveFirst
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


Private Sub SaveCmd_Click()
If CmbLocation.Text = "" Then
   MsgBox "Select Location Name!!"
   CmbLocation.SetFocus
   Exit Sub
End If

If CmbLessorName.Text = "" Then
   MsgBox "Select Lessor Name not allowed!!!"
   CmbLessorName.SetFocus
   Exit Sub
End If

If TxtBillingNo.Text = "" Then
   MsgBox "Blank Bill No not allowed!!!"
   TxtBillingNo.SetFocus
   Exit Sub
End If

If TxtBillingAmount.Text = "" Then
   MsgBox "Blank Billing Amount not allowed!!!"
   TxtBillingAmount.SetFocus
   Exit Sub
End If

mSumAmt = 0
For I = 2 To MSHFlexGrid3.Rows - 1
    If MSHFlexGrid3.TextMatrix(I, 0) = strChecked Then
       mSumAmt = Val(mSumAmt) + Val(MSHFlexGrid3.TextMatrix(I, 4))
    End If
Next

If Val(mSumAmt) <> Val(TxtBillingAmount) Then
   MsgBox "Please Check Billing Amount"
   TxtBillingAmount.SetFocus
   Exit Sub
End If

If TxtSTRate.Text = "" Then
   MsgBox "Blank ST Rate not allowed!!!"
   TxtSTRate.SetFocus
   Exit Sub
End If

If LCase(CmbStatus.Text) = "pending" Then
   If Format(Gen_BlockingDate, "yyyy-mm") >= Format(CDate(TxtMonYear.Value), "yyyy-mm") Then
      MsgBox "Select next month from " & Format(Gen_BlockingDate, "MMM-yyyy") & "!!! "
      TxtMonYear.SetFocus
      Exit Sub
   End If
''   If Format(Gen_BlockingDate, "yyyy-mm") >= Format(CDate(TxtBillingDate.Value), "yyyy-mm") Then
''      MsgBox "Select next month from " & Format(Gen_BlockingDate, "MMM-yyyy") & "!!! "
''      TxtBillingDate.SetFocus
''      Exit Sub
''   End If
End If
If MSHFlexGrid1.Rows < 3 Then
   MsgBox "Claims details not found for selected month"
   Exit Sub
End If

'If Val(TxtBillingAmount) <> Val(MSHFlexGrid3.TextMatrix(0, 3)) Then
'   MsgBox "Billing Amount should be eqaual to total 'Lease Amount (With Service Tax)' "
'   TxtBillingAmount.SetFocus
'   Exit Sub
'End If

If LCase(CmbStatus.Text) <> "pending" Then
   If TxtFAGRemarks = "" Then
      MsgBox " Blank FAG Remarks not allowed!!! "
      TxtFAGRemarks.SetFocus
      Exit Sub
   End If
     
   If TxtFAGPostingDate.Value > Date Then
      MsgBox "Future date not allowed!!"
      TxtFAGPostingDate.SetFocus
      Exit Sub
   End If
   
   If Format(Gen_BlockingDate, "yyyy-mm") >= Format(CDate(TxtFAGPostingDate.Value), "yyyy-mm") Then
      MsgBox "Select next month from " & Format(Gen_BlockingDate, "MMM-yyyy") & "!!! "
      TxtFAGPostingDate.SetFocus
      Exit Sub
   End If
End If

If Edit_Flg = "A" Then
'   If CheckDuplicate = True Then
'      MsgBox "Duplicate Location,Lesor & Month-Year not allowed!!"
'      TxtMonYear.SetFocus
'      Exit Sub
'   End If
   RsTxn_GodownRentST.AddNew
   'RsTxn_GodownRentST!GodownRentSTID = getGodownRentSTID
   RsTxn_GodownRentST!G_UID = GetGUID
   'TxtGodownRentSTID = RsTxn_GodownRentST!GodownRentSTID
   mG_UIDNew = RsTxn_GodownRentST!G_UID
Else
   RsTxn_GodownRentST.MoveFirst
   RsTxn_GodownRentST.Find "GodownRentSTID = " & TxtGodownRentSTID
End If
RsTxn_GodownRentST!LocationID = mLocationID
RsTxn_GodownRentST!LessorID = mLessorID
RsTxn_GodownRentST!GodownRentSTDate = Format(TxtEntryDate.Value, Gen_DatFormat)
RsTxn_GodownRentST!BillNo = TxtBillingNo.Text
RsTxn_GodownRentST!BillDate = Format(TxtBillingDate.Value, Gen_DatFormat)
RsTxn_GodownRentST!BillAmount = Val(TxtBillingAmount.Text)
RsTxn_GodownRentST!STRate = Val(TxtSTRate.Text)
RsTxn_GodownRentST!STAmount = Val(TxtSTAmount.Text)
RsTxn_GodownRentST!MonthYear = Format(TxtMonYear.Value, Gen_DatFormat)
RsTxn_GodownRentST!Flag = mStatusFlag
RsTxn_GodownRentST!Remarks = TxtRemark.Text
RsTxn_GodownRentST!BatchID = 0
If LCase(CmbStatus) <> "pending" Then
   RsTxn_GodownRentST!FAGRemarks = Trim(TxtFAGRemarks.Text) & " Claim ID :- " & Val(TxtGodownRentSTID) & " "
   RsTxn_GodownRentST!FAGPostingDate = TxtFAGPostingDate.Value
Else
   RsTxn_GodownRentST!FAGRemarks = Null
   RsTxn_GodownRentST!FAGPostingDate = Null
End If

If IsNull(RsTxn_GodownRentST!CreatedBy) = True Or RsTxn_GodownRentST!CreatedBy = "" Then
   RsTxn_GodownRentST!CreatedBy = Gen_UserLoginID
   RsTxn_GodownRentST!CreatedDate = Now
Else
   RsTxn_GodownRentST!UpdatedBy = Gen_UserLoginID
   RsTxn_GodownRentST!UpdatedDate = Now
End If

RsTxn_GodownRentST.Update

tmp = "Select GodownRentSTID From Txn_GodownRentST Where G_UID='" & mG_UIDNew & "'"
Call TmpTable

If TmpRs.RecordCount > 0 Then
'   mCxTF_DetailsID = TmpRs!CxTF_DetailsID
   TxtGodownRentSTID.Text = TmpRs!GodownRentSTID
End If


Call SaveRentSTClaimDetails
Call SaveRentSTAgreeDetail

RsTxn_GodownRentST.Requery
RsTxn_GodownRentST.MoveFirst
RsTxn_GodownRentST.Find "GodownRentSTID = " & TxtGodownRentSTID
If RsTxn_GodownRentST.EOF Then
   RsTxn_GodownRentST.MoveFirst
End If
Call GButtonLock(Me, True)
Call Indata
Exit Sub
End Sub

Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   If RsTxn_GodownRentST.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   RsTxn_GodownRentST.MoveFirst
   RsTxn_GodownRentST.Find "GodownRentSTID  = " & Val(TxtGodownRentSTID) & ""
   If RsTxn_GodownRentST.EOF Then
      RsTxn_GodownRentST.MoveLast
   End If
   Call Indata
   Exit Sub
End If
Frame0.Height = Me.Height - 500
Frame0.Left = Frame1.Left - 100
Frame0.Top = Frame1.Top
Frame0.Width = Me.Width - 200
Frame0.Visible = True

MSHFlexGrid2.Width = Frame0.Width - 200
MSHFlexGrid2.Height = Frame0.Height - 2500

RsTxn_GodownRentST.MoveFirst
'RsTxn_GodownRentST.Filter = ""
End Sub


Public Sub Indata()

Frame1.Enabled = False
Frame2.Enabled = True

CmdPrint.Enabled = True

CmbLocation.Visible = False
TxtLocation.Visible = True

CmbLessorName.Visible = False
TxtLessorName.Visible = True

CmbStatus.Visible = False
TxtStatus.Visible = True

TxtMonYear.Enabled = True

TxtGodownRentSTID = RsTxn_GodownRentST!GodownRentSTID

mLocationID = RsTxn_GodownRentST!LocationID
TxtLocation.Text = GetLocationName(RsTxn_GodownRentST!LocationID)

mLessorID = IIf(IsNull(RsTxn_GodownRentST!LessorID), "", RsTxn_GodownRentST!LessorID)
TxtLessorName.Text = GetLessorName(RsTxn_GodownRentST!LessorID)

TxtEntryDate.Value = RsTxn_GodownRentST!GodownRentSTDate
TxtMonYear.Value = RsTxn_GodownRentST!MonthYear
TxtBillingDate.Value = RsTxn_GodownRentST!BillDate
TxtBillingNo.Text = RsTxn_GodownRentST!BillNo
TxtBillingAmount.Text = RsTxn_GodownRentST!BillAmount
TxtSTRate.Text = RsTxn_GodownRentST!STRate
TxtSTAmount.Text = RsTxn_GodownRentST!STAmount

mStatusFlag = RsTxn_GodownRentST!Flag

If LCase(mStatusFlag) = "p" Then
   TxtStatus.Text = "Pending"
ElseIf LCase(mStatusFlag) = "a" Then
   TxtStatus.Text = "Approved"
ElseIf LCase(mStatusFlag) = "r" Then
   TxtStatus.Text = "Rejected"
End If
TxtRemark.Locked = True
TxtFAGRemarks.Locked = True

TxtRemark = RsTxn_GodownRentST!Remarks
TxtFAGRemarks = IIf(IsNull(RsTxn_GodownRentST!FAGRemarks), "", RsTxn_GodownRentST!FAGRemarks)
If IsNull(RsTxn_GodownRentST!FAGPostingDate) = False Then
   TxtFAGPostingDate.Value = RsTxn_GodownRentST!FAGPostingDate
   TxtFAGPostingDate.Enabled = False
Else
   TxtFAGPostingDate.Visible = False
   Label12.Visible = False
End If

CmbLocation.Text = ""
CmbLessorName.Text = ""

TxtCreated = IIf(IsNull(RsTxn_GodownRentST!CreatedBy), "", RsTxn_GodownRentST!CreatedBy) & " :- " & IIf(IsNull(RsTxn_GodownRentST!CreatedDate), "", RsTxn_GodownRentST!CreatedDate)
TxtUpdated = IIf(IsNull(RsTxn_GodownRentST!UpdatedBy), "", RsTxn_GodownRentST!UpdatedBy) & " :- " & IIf(IsNull(RsTxn_GodownRentST!UpdatedDate), "", RsTxn_GodownRentST!UpdatedDate)
Call InDataAgreementDetails
Call InDataClaimDetails

If RsTxn_GodownRentST.RecordCount = 1 Then
   Call GButtonLock_1(Me, True)
Else
   Call GButtonLock(Me, True)
End If
End Sub
Private Function getGodownRentSTID() As Double
Dim Retval As Double
tmp = "Select max(GodownRentSTID) from Txn_GodownRentST"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
getGodownRentSTID = Retval
End Function
Private Sub GetClaimDetails()

tmp = "Select Distinct TGR.ClaimID,TGR.ClaimDate,TGR.Flag,Sum(TG.Amount) as Amount"
tmp = tmp & " From Txn_GodownRent TGR"
tmp = tmp & " Inner join Txn_GodownRentDetail TG on TG.ClaimID=TGR.ClaimID"
tmp = tmp & " Inner join Mst_Godown MG on TG.GodownID=MG.GodownID"
tmp = tmp & " Inner Join Mst_Lessor ML On TGR.LessorID = ML.LessorID"
tmp = tmp & " Inner join Mst_CMP MC on MG.CMPID=MC.CMPID"
tmp = tmp & " Where TGR.Flag in ('S','A') And TGR.LessorID=" & mLessorID
tmp = tmp & " And MC.LocationID=" & mLocationID & " and TGR.ExpenseFrom='" & Format(TxtMonYear.Value, Gen_DatFormat) & "'"
tmp = tmp & " And TG.LAID Not In (Select TGRSD.AgreementID From Txn_GodownRentST TGRS Inner Join Txn_GodownRentSTAgreeDetail TGRSD On TGRS.GodownRentSTID=TGRSD.GodownRentSTID Where TGRS.Flag not in ('R') And TGRS.MonthYear='" & Format(TxtMonYear.Value, Gen_DatFormat) & "' And TGRS.LessorID=" & mLessorID & ")"
tmp = tmp & " Group by TGR.ClaimID,TGR.ClaimDate,TGR.Flag"
Call Tmp2Table
Call Grid_Head2
If TmpRs2.RecordCount > 0 Then
   MSHFlexGrid1.Rows = TmpRs2.RecordCount + 2
   For iu = 1 To TmpRs2.RecordCount
       MSHFlexGrid1.TextMatrix(iu, 0) = IIf(IsNull(TmpRs2!ClaimID), "", TmpRs2!ClaimID)
       MSHFlexGrid1.TextMatrix(iu, 1) = IIf(IsNull(TmpRs2!ClaimDate), "", Format(TmpRs2!ClaimDate, "dd-MMM-yyyy"))
       MSHFlexGrid1.TextMatrix(iu, 2) = IIf(IsNull(TmpRs2!Amount), "", TmpRs2!Amount)
       MSHFlexGrid1.TextMatrix(iu, 3) = IIf(IsNull(TmpRs2!Flag), "", TmpRs2!Flag)
       If MSHFlexGrid1.TextMatrix(iu, 3) = "A" Then
          MSHFlexGrid1.TextMatrix(iu, 3) = "Approved"
       ElseIf MSHFlexGrid1.TextMatrix(iu, 3) = "P" Then
          MSHFlexGrid1.TextMatrix(iu, 3) = "Pending"
       ElseIf MSHFlexGrid1.TextMatrix(iu, 3) = "S" Then
          MSHFlexGrid1.TextMatrix(iu, 3) = "Settled"
       End If
       TmpRs2.MoveNext
   Next
Else
    MsgBox "Claims details not found for selected month"
End If
End Sub


Private Sub TxtBillingAmount_LostFocus()
If TxtBillingAmount = "" Then Exit Sub
If IsNumeric(TxtBillingAmount) = False Or Val(TxtBillingAmount) < 0 Then
   MsgBox "Invalid numeric format !!!!"
   TxtBillingAmount.SetFocus
   Exit Sub
End If
TxtSTAmount.Text = Format(TxtBillingAmount * (Val(TxtSTRate) / 100), "#####0.00")
End Sub

Private Sub TxtMonYear_LostFocus()
TxtMonYear.Day = 1
If CmbLocation.Text <> "" And CmbLessorName.Text <> "" Then
   Call GetAgreementDetails
   TxtBillingAmount.Text = ""
End If

TxtRemark.Text = "BEING GST PAYMENT FOR THE MONTH " & Format(TxtMonYear, "MMM-YYYY")

End Sub
Private Sub GetAgreementDetails()

tmp = "Select TL.LAID,TL.EntryDate,sum(TG.Amount)Amount,TL.ServiceTaxID"
tmp = tmp & " From Txn_GodownRent TGR"
tmp = tmp & " Inner join Txn_GodownRentDetail TG on TG.ClaimID=TGR.ClaimID"
tmp = tmp & " Inner join Txn_LessorAgreement TL on TG.LAID=TL.LAID"
tmp = tmp & " Inner join Mst_Godown MG on TG.GodownID=MG.GodownID"
tmp = tmp & " Inner Join Mst_Lessor ML On TGR.LessorID = ML.LessorID"
tmp = tmp & " Inner join Mst_CMP MC on MG.CMPID=MC.CMPID"
tmp = tmp & " Where TGR.Flag in ('S','A') And TGR.LessorID=" & mLessorID
tmp = tmp & " And MC.LocationID=" & mLocationID & " and TGR.ExpenseFrom='" & Format(TxtMonYear.Value, Gen_DatFormat) & "'"
tmp = tmp & " And TL.LAID Not In (Select TGRSD.AgreementID From Txn_GodownRentST TGRS Inner Join Txn_GodownRentSTAgreeDetail TGRSD On TGRS.GodownRentSTID=TGRSD.GodownRentSTID Where TGRS.Flag not in ('R') And TGRS.MonthYear='" & Format(TxtMonYear.Value, Gen_DatFormat) & "' And TGRS.LessorID=" & mLessorID & " )"
tmp = tmp & " Group by  TL.LAID,TL.ServiceTaxID,TL.EntryDate"

Call Tmp2Table
Call Grid_Head1
If TmpRs2.RecordCount > 0 Then
   MSHFlexGrid3.Rows = TmpRs2.RecordCount + 3
   For iu = 1 To TmpRs2.RecordCount
       
       MSHFlexGrid3.TextMatrix(iu + 1, 1) = IIf(IsNull(TmpRs2!LAID), "", TmpRs2!LAID)
       MSHFlexGrid3.TextMatrix(iu + 1, 2) = IIf(IsNull(TmpRs2!EntryDate), "", Format(TmpRs2!EntryDate, "dd-MMM-yyyy"))
       If TmpRs2!ServiceTaxID = 3 Or TmpRs2!ServiceTaxID = 5 Then
          MSHFlexGrid3.TextMatrix(iu + 1, 4) = IIf(IsNull(TmpRs2!Amount), "", TmpRs2!Amount)
       Else
          MSHFlexGrid3.TextMatrix(iu + 1, 3) = IIf(IsNull(TmpRs2!Amount), "", TmpRs2!Amount)
       End If
       MSHFlexGrid3.TextMatrix(0, 4) = Val(MSHFlexGrid3.TextMatrix(0, 4)) + Val(MSHFlexGrid3.TextMatrix(iu + 1, 4))
       MSHFlexGrid3.TextMatrix(0, 3) = Val(MSHFlexGrid3.TextMatrix(0, 3)) + Val(MSHFlexGrid3.TextMatrix(iu + 1, 3))
       
       MSHFlexGrid3.row = iu + 1
       MSHFlexGrid3.Col = 0
       MSHFlexGrid3.CellFontName = "Wingdings"
       MSHFlexGrid3.CellFontSize = 12
       MSHFlexGrid3.CellAlignment = flexAlignCenterCenter
       MSHFlexGrid3.Text = strUnChecked

       TmpRs2.MoveNext
   Next
      Call GetClaimDetails
Else
    MsgBox "Claims details not found for selected month"
    Call Grid_Head1
    Call Grid_Head2

End If

End Sub

Private Sub TxtSTRate_GotFocus()
'If TxtBillingAmount.Text = "" Then
'   MsgBox "Enter Billing Amount!!"
'   TxtBillingAmount.SetFocus
'   Exit Sub
'End If

End Sub

Private Sub TxtSTRate_LostFocus()

If TxtSTRate = "" Then Exit Sub
If IsNumeric(TxtSTRate) = False Or Val(TxtSTRate) < 0 Then
   MsgBox "Invalid numeric format !!!!"
   TxtSTRate.SetFocus
   Exit Sub
End If
TxtSTAmount.Text = Format(Val(TxtBillingAmount) * (Val(TxtSTRate) / 100), "#####0.00")

End Sub
Private Sub CmbStatus_LostFocus()
If CmbStatus.Text = "" Then
   TxtFAGRemarks.Locked = True
   TxtFAGPostingDate.Visible = False
   Label12.Visible = False
   Exit Sub
End If

If LCase(CmbStatus.Text) = "approved" Then
   mStatusFlag = "A"
   TxtFAGRemarks.Locked = False
   TxtFAGPostingDate.Visible = True
   Label12.Visible = True
   TxtFAGPostingDate.Enabled = True
   mMonth = Month(TxtBillingDate.Value)
   If mMonth < 10 Then
      mMonth = mMonth
   End If
   mTotalDays = Day(DateSerial(Year(TxtBillingDate.Value), mMonth + 1, 0))
   mFAGToDate = CDate(Year(TxtBillingDate.Value) & "/" & mMonth & "/" & mTotalDays)
    '"" 'Date
   If CDate(mFAGToDate) > Date Then
      TxtFAGPostingDate.Value = Date
   Else
      TxtFAGPostingDate.Value = CDate(mFAGToDate)
   End If
   

ElseIf LCase(CmbStatus.Text) = "pending" Then
   mStatusFlag = "P"
   TxtFAGRemarks.Locked = True
   TxtFAGPostingDate.Visible = False
   Label12.Visible = False
ElseIf LCase(CmbStatus.Text) = "rejected" Then
   mStatusFlag = "R"
   TxtFAGRemarks.Locked = False
   TxtFAGPostingDate.Visible = True
   Label12.Visible = True
   TxtFAGPostingDate.Enabled = True
   TxtFAGPostingDate.Value = Date
Else
   MsgBox "Invalid selection!!!"
   CmbStatus.SetFocus
   TxtFAGRemarks.Locked = True
   TxtFAGPostingDate.Visible = False
   Label12.Visible = False
   Exit Sub
End If

If LCase(Gen_UserRole) <> "fag" And LCase(Gen_UserRole) <> "fagadmin" And LCase(Gen_UserRole) <> "power" Then
   If mStatusFlag = "A" Then
      MsgBox "Only FAG user have access to approve Ad-Hoc Invoice '!!!"
      CmbStatus.SetFocus
      Exit Sub
   End If
End If

End Sub

Private Sub SaveRentSTClaimDetails()
Call DeleteRentClaimDetails
Call TableTxn_GodownRentSTClaimDetail(" Where GodownRentSTID=" & Val(TxtGodownRentSTID) & "")  '(" Where ClaimID = " & TxtClaimID & " AND SM_Prefix = '" & TxtPreFix & "'")
For I = 1 To MSHFlexGrid1.Rows - 1
    If MSHFlexGrid1.TextMatrix(I, 1) <> "" Then
       RsTxn_GodownRentSTClaimDetail.AddNew
 '      RsTxn_GodownRentSTClaimDetail!GSTCID = getMaxGSTCID
       RsTxn_GodownRentSTClaimDetail!GodownRentSTID = TxtGodownRentSTID
       RsTxn_GodownRentSTClaimDetail!ClaimID = MSHFlexGrid1.TextMatrix(I, 0)
       RsTxn_GodownRentSTClaimDetail.Update
       RsTxn_GodownRentSTClaimDetail.Requery
    End If
Next
End Sub

Private Sub DeleteRentClaimDetails()
tmp = "Delete From Txn_GodownRentSTClaimDetail Where GodownRentSTID=" & Val(TxtGodownRentSTID) & " "
Call Tmp4Table
End Sub

Private Function getMaxGSTCID() As Double
Dim Retval As Double
tmp = "Select max(GSTCID) from Txn_GodownRentSTClaimDetail"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
getMaxGSTCID = Retval
End Function
Private Sub DeleteRentAgreeDetails()
tmp = "Delete From Txn_GodownRentSTAgreeDetail Where GodownRentSTID=" & Val(TxtGodownRentSTID) & " "
Call Tmp4Table
End Sub

Private Sub SaveRentSTAgreeDetail()
Call DeleteRentAgreeDetails
Call TableTxn_GodownRentSTAgreeDetail(" Where GodownRentSTID=" & Val(TxtGodownRentSTID) & "")  '(" Where ClaimID = " & TxtClaimID & " AND SM_Prefix = '" & TxtPreFix & "'")
For I = 2 To MSHFlexGrid3.Rows - 1
    If MSHFlexGrid3.TextMatrix(I, 1) <> "" Then
       If MSHFlexGrid3.TextMatrix(I, 0) = strChecked Then
          RsTxn_GodownRentSTAgreeDetail.AddNew
'          RsTxn_GodownRentSTAgreeDetail!GSTAID = getMaxGSTAID
          RsTxn_GodownRentSTAgreeDetail!GodownRentSTID = TxtGodownRentSTID
          RsTxn_GodownRentSTAgreeDetail!AgreementID = MSHFlexGrid3.TextMatrix(I, 1)
          RsTxn_GodownRentSTAgreeDetail!AmountWithoutST = Val(MSHFlexGrid3.TextMatrix(I, 3))
          RsTxn_GodownRentSTAgreeDetail!AmountWithST = Val(MSHFlexGrid3.TextMatrix(I, 4))
          RsTxn_GodownRentSTAgreeDetail.Update
          RsTxn_GodownRentSTAgreeDetail.Requery
       End If
    End If
Next
End Sub

Private Function getMaxGSTAID() As Double
Dim Retval As Double
tmp = "Select max(GSTAID) from Txn_GodownRentSTAgreeDetail"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
getMaxGSTAID = Retval
End Function

Private Sub InDataClaimDetails()

Call Grid_Head

'tmp = "Select Distinct TGR.ClaimID,TGR.ClaimDate,TGR.Flag,Sum(TGRD.Amount) as Amount"
'tmp = tmp & " From Txn_GodownRentSTClaimDetail TGC"
'tmp = tmp & " Inner join Txn_GodownRent TGR on TGR.ClaimID=TGC.ClaimID"
'tmp = tmp & " Inner Join Txn_GodownRentDetail TGRD on TGR.ClaimID=TGRD.ClaimID"
'tmp = tmp & " Inner Join Txn_GodownRentSTAgreeDetail TGRAD on TGRD.LAID=TGRAD.AgreementID"
'tmp = tmp & " Where TGC.GodownRentSTID= " & Val(TxtGodownRentSTID) & " "
'tmp = tmp & " Group By TGR.ClaimID,TGR.ClaimDate,TGR.Flag"

tmp = "Select Distinct TGR.ClaimID,TGR.ClaimDate,TGR.Flag,Sum(TG.Amount) as Amount"
tmp = tmp & " From Txn_GodownRent TGR"
tmp = tmp & " Inner join Txn_GodownRentDetail TG on TG.ClaimID=TGR.ClaimID"
tmp = tmp & " Inner join Mst_Godown MG on TG.GodownID=MG.GodownID"
tmp = tmp & " Inner Join Mst_Lessor ML On TGR.LessorID = ML.LessorID"
tmp = tmp & " Inner join Mst_CMP MC on MG.CMPID=MC.CMPID"
tmp = tmp & " Where TGR.Flag in ('S','A') And TGR.LessorID=" & mLessorID
tmp = tmp & " And MC.LocationID=" & mLocationID & " and TGR.ExpenseFrom='" & Format(TxtMonYear.Value, Gen_DatFormat) & "'"
tmp = tmp & " And TG.LAID In (Select TGRSD.AgreementID From Txn_GodownRentST TGRS Inner Join Txn_GodownRentSTAgreeDetail TGRSD On TGRS.GodownRentSTID=TGRSD.GodownRentSTID Where TGRSD.GodownRentSTID=" & TxtGodownRentSTID.Text & ")"
tmp = tmp & " Group by TGR.ClaimID,TGR.ClaimDate,TGR.Flag"

Call Tmp2Table

Call Grid_Head2

If TmpRs2.RecordCount > 0 Then
   MSHFlexGrid1.Rows = TmpRs2.RecordCount + 2
   For iu = 1 To TmpRs2.RecordCount
       MSHFlexGrid1.TextMatrix(iu, 0) = IIf(IsNull(TmpRs2!ClaimID), "", TmpRs2!ClaimID)
       MSHFlexGrid1.TextMatrix(iu, 1) = IIf(IsNull(TmpRs2!ClaimDate), "", Format(TmpRs2!ClaimDate, "dd-MMM-yyyy"))
       MSHFlexGrid1.TextMatrix(iu, 2) = IIf(IsNull(TmpRs2!Amount), "", TmpRs2!Amount)
       MSHFlexGrid1.TextMatrix(iu, 3) = IIf(IsNull(TmpRs2!Flag), "", TmpRs2!Flag)
       If MSHFlexGrid1.TextMatrix(iu, 3) = "A" Then
          MSHFlexGrid1.TextMatrix(iu, 3) = "Approved"
       ElseIf MSHFlexGrid1.TextMatrix(iu, 3) = "P" Then
          MSHFlexGrid1.TextMatrix(iu, 3) = "Pending"
       ElseIf MSHFlexGrid1.TextMatrix(iu, 3) = "S" Then
          MSHFlexGrid1.TextMatrix(iu, 3) = "Settled"
       End If
       TmpRs2.MoveNext
   Next
End If
End Sub

Private Sub InDataAgreementDetails()

tmp = "Select TL.LAID,TL.EntryDate,TGA.AmountWithoutST,TGA.AmountWithST from"
tmp = tmp & " Txn_GodownRentSTAgreeDetail TGA"
tmp = tmp & " Inner join Txn_LessorAgreement TL on TGA.AgreementID=TL.LAID"
tmp = tmp & " Where TGA.GodownRentSTID= " & Val(TxtGodownRentSTID) & " "

Call Tmp2Table
Call Grid_Head1
If TmpRs2.RecordCount > 0 Then
   MSHFlexGrid3.Rows = TmpRs2.RecordCount + 3
   For iu = 1 To TmpRs2.RecordCount
       MSHFlexGrid3.row = iu + 1
       MSHFlexGrid3.Col = 0
       MSHFlexGrid3.CellFontName = "Wingdings"
       MSHFlexGrid3.CellFontSize = 12
       MSHFlexGrid3.CellAlignment = flexAlignCenterCenter
       MSHFlexGrid3.Text = strChecked
       
       MSHFlexGrid3.TextMatrix(iu + 1, 1) = IIf(IsNull(TmpRs2!LAID), "", TmpRs2!LAID)
       MSHFlexGrid3.TextMatrix(iu + 1, 2) = IIf(IsNull(TmpRs2!EntryDate), "", Format(TmpRs2!EntryDate, "dd-MMM-yyyy"))
       MSHFlexGrid3.TextMatrix(iu + 1, 3) = IIf(IsNull(TmpRs2!AmountWithoutST), "", TmpRs2!AmountWithoutST)
       MSHFlexGrid3.TextMatrix(iu + 1, 4) = IIf(IsNull(TmpRs2!AmountWithST), "", TmpRs2!AmountWithST)
       
       MSHFlexGrid3.TextMatrix(0, 4) = Val(MSHFlexGrid3.TextMatrix(0, 4)) + Val(MSHFlexGrid3.TextMatrix(iu + 1, 4))
       MSHFlexGrid3.TextMatrix(0, 3) = Val(MSHFlexGrid3.TextMatrix(0, 3)) + Val(MSHFlexGrid3.TextMatrix(iu + 1, 3))
       TmpRs2.MoveNext
   Next
End If

End Sub

Private Sub CmbSLessorName_GotFocus()
tmp = CmbSLessorName.Text
Call TableMst_Lessor("")
CmbSLessorName.Clear
If RsMst_Lessor.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Lessor.RecordCount
    CmbSLessorName.AddItem RsMst_Lessor!LessorName
    RsMst_Lessor.MoveNext
Next
CmbSLessorName.Text = tmp
End Sub

Private Sub CmbSLessorName_LostFocus()
If CmbSLessorName.Text = "" Then
   Exit Sub
End If
RsMst_Lessor.MoveFirst
RsMst_Lessor.Find "LessorName = '" & CmbSLessorName.Text & "'"
If RsMst_Lessor.EOF Then
   MsgBox "Invalid selection "
   CmbSLessorName.SetFocus
   Exit Sub
End If
mSLessorID = RsMst_Lessor!LessorID

End Sub

Private Sub CmbSStatus_LostFocus()
If CmbSStatus.Text = "" Then
   Exit Sub
End If

If LCase(CmbSStatus.Text) <> "approved" And LCase(CmbSStatus.Text) <> "pending" And LCase(CmbSStatus.Text) <> "rejected" Then
   MsgBox "Invalid selection!!!"
   CmbSStatus.SetFocus
   Exit Sub
End If

End Sub

Private Function CheckDuplicate() As Boolean
Dim Retval As Boolean
Retval = False
tmp = "Select * from Txn_GodownRentST Where LocationID=" & mLocationID & " And LessorID=" & mLessorID & " and MonthYear='" & TxtMonYear.Value & "' And Flag<>'R'"

Call Tmp3Table

If TmpRs3.RecordCount > 0 Then
   Retval = True
End If
CheckDuplicate = Retval
End Function

