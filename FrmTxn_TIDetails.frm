VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmTxn_TIDetails 
   Caption         =   "Invoice Details"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   12915
   ScaleWidth      =   21360
   WindowState     =   2  'Maximized
   Begin VB.VScrollBar VScroll1 
      Height          =   10095
      Left            =   14565
      TabIndex        =   17
      Top             =   0
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   9600
      Width           =   10935
   End
   Begin VB.Frame Frame4 
      Height          =   9735
      Left            =   120
      TabIndex        =   19
      Top             =   0
      Width           =   14415
      Begin VB.Frame Frame0 
         Height          =   1965
         Left            =   0
         TabIndex        =   20
         Top             =   4080
         Visible         =   0   'False
         Width           =   14385
         Begin VB.TextBox TxtSNAVCode 
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
            Left            =   1440
            TabIndex        =   111
            Top             =   1440
            Width           =   3840
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
            Left            =   3945
            Sorted          =   -1  'True
            TabIndex        =   109
            Top             =   1020
            Visible         =   0   'False
            Width           =   4095
         End
         Begin VB.ComboBox CmbSReverseFlag 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            ItemData        =   "FrmTxn_TIDetails.frx":0000
            Left            =   2520
            List            =   "FrmTxn_TIDetails.frx":000A
            TabIndex        =   104
            Top             =   1020
            Width           =   1335
         End
         Begin VB.ComboBox CmbSEntryType 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            ItemData        =   "FrmTxn_TIDetails.frx":0017
            Left            =   105
            List            =   "FrmTxn_TIDetails.frx":0021
            Sorted          =   -1  'True
            TabIndex        =   93
            Top             =   1020
            Width           =   2280
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
            ItemData        =   "FrmTxn_TIDetails.frx":0035
            Left            =   10800
            List            =   "FrmTxn_TIDetails.frx":0037
            Sorted          =   -1  'True
            TabIndex        =   79
            Top             =   360
            Width           =   3240
         End
         Begin VB.TextBox TxtSEntryNo 
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
            Left            =   8085
            TabIndex        =   77
            Top             =   360
            Width           =   2655
         End
         Begin VB.TextBox TxtSPostInvoiceNo 
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
            Left            =   5505
            TabIndex        =   75
            Top             =   360
            Width           =   2535
         End
         Begin VB.CommandButton CmdExcel 
            Caption         =   "Excel"
            Height          =   350
            Left            =   12945
            TabIndex        =   23
            Top             =   1025
            Width           =   1000
         End
         Begin VB.CommandButton CmdGo 
            Caption         =   "Go"
            Height          =   345
            Left            =   11790
            TabIndex        =   22
            Top             =   1028
            Width           =   1125
         End
         Begin VB.TextBox TxtSTINo 
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
            Left            =   2925
            TabIndex        =   21
            Top             =   360
            Width           =   2535
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   855
            Left            =   90
            TabIndex        =   24
            Top             =   1920
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
            TabIndex        =   25
            Top             =   375
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
            Format          =   107937793
            CurrentDate     =   39884
         End
         Begin MSComCtl2.DTPicker STo 
            Height          =   360
            Left            =   1575
            TabIndex        =   26
            Top             =   375
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
            Format          =   107937793
            CurrentDate     =   39884
         End
         Begin VB.Label Label40 
            Caption         =   "NAV Code"
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
            TabIndex        =   112
            Top             =   1500
            Width           =   1140
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            Caption         =   "Last UpdatedBy"
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
            Left            =   5265
            TabIndex        =   110
            Top             =   750
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "Entry Type"
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
            Left            =   705
            TabIndex        =   108
            Top             =   750
            Width           =   975
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            Caption         =   "Reverse Flag"
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
            Left            =   2565
            TabIndex        =   107
            Top             =   750
            Width           =   1230
         End
         Begin VB.Label Label19 
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
            Left            =   11970
            TabIndex        =   80
            Top             =   120
            Width           =   765
         End
         Begin VB.Label Label18 
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
            Left            =   9015
            TabIndex        =   78
            Top             =   120
            Width           =   765
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Post Invoice No"
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
            Left            =   6060
            TabIndex        =   76
            Top             =   120
            Width           =   1410
         End
         Begin VB.Label Label14 
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
            Left            =   2070
            TabIndex        =   31
            Top             =   120
            Width           =   735
         End
         Begin VB.Label Label12 
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
            Left            =   495
            TabIndex        =   30
            Top             =   120
            Width           =   945
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
            Left            =   10800
            TabIndex        =   29
            Top             =   960
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
            Left            =   8160
            TabIndex        =   28
            Top             =   960
            Width           =   2415
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            Caption         =   "TI No (Voucher No)"
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
            Left            =   3300
            TabIndex        =   27
            Top             =   120
            Width           =   1725
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1575
         Left            =   120
         TabIndex        =   90
         Top             =   6960
         Width           =   13335
         Begin VB.TextBox TxtAdjAmount 
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
            Left            =   1800
            TabIndex        =   12
            Top             =   240
            Width           =   3840
         End
         Begin VB.TextBox TxtRemarks 
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
            Left            =   1800
            MaxLength       =   2000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   13
            Top             =   660
            Width           =   11385
         End
         Begin VB.Label Label22 
            Caption         =   "Adjusted Amount"
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
            Left            =   120
            TabIndex        =   92
            Top             =   285
            Width           =   1695
         End
         Begin VB.Label Label23 
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
            Height          =   270
            Left            =   120
            TabIndex        =   91
            Top             =   945
            Width           =   1695
         End
      End
      Begin VB.Frame Frame1 
         Height          =   6825
         Left            =   120
         TabIndex        =   39
         Top             =   120
         Width           =   13305
         Begin VB.TextBox TxtReverseFlag 
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
            Left            =   7800
            Locked          =   -1  'True
            TabIndex        =   105
            TabStop         =   0   'False
            Top             =   225
            Width           =   1890
         End
         Begin VB.TextBox TxtCUpdatedBy 
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
            Left            =   3240
            Locked          =   -1  'True
            TabIndex        =   101
            TabStop         =   0   'False
            Top             =   6360
            Width           =   2370
         End
         Begin VB.TextBox TxtCUpdatedDate 
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
            Left            =   9120
            Locked          =   -1  'True
            TabIndex        =   100
            TabStop         =   0   'False
            Top             =   6360
            Width           =   3450
         End
         Begin VB.TextBox TxtCRemark 
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
            Height          =   615
            Left            =   1800
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   98
            TabStop         =   0   'False
            Top             =   5640
            Width           =   10770
         End
         Begin VB.TextBox TxtSMID 
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
            Left            =   10920
            Locked          =   -1  'True
            TabIndex        =   96
            TabStop         =   0   'False
            Top             =   2400
            Width           =   1650
         End
         Begin VB.TextBox TxtInvType 
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
            Left            =   7800
            Locked          =   -1  'True
            TabIndex        =   94
            TabStop         =   0   'False
            Top             =   2400
            Width           =   1890
         End
         Begin VB.TextBox TxtNAVClientName 
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
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   88
            TabStop         =   0   'False
            Top             =   5190
            Width           =   10770
         End
         Begin VB.ComboBox CmbNAVCode 
            BeginProperty Font 
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
            TabIndex        =   84
            Top             =   2400
            Width           =   3720
         End
         Begin VB.ComboBox CmbServiceTaxID 
            BeginProperty Font 
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
            TabIndex        =   11
            Top             =   3780
            Width           =   3870
         End
         Begin VB.CommandButton CmdServiceTaxID 
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
            Height          =   330
            Left            =   5715
            TabIndex        =   81
            Top             =   3795
            Width           =   375
         End
         Begin VB.ComboBox CmbDPName 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2400
            Sorted          =   -1  'True
            TabIndex        =   74
            Top             =   1920
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.ComboBox CmbCNonNcdexClient 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            ItemData        =   "FrmTxn_TIDetails.frx":0039
            Left            =   1815
            List            =   "FrmTxn_TIDetails.frx":003B
            Sorted          =   -1  'True
            TabIndex        =   41
            Top             =   1920
            Visible         =   0   'False
            Width           =   495
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
            ItemData        =   "FrmTxn_TIDetails.frx":003D
            Left            =   7800
            List            =   "FrmTxn_TIDetails.frx":003F
            Sorted          =   -1  'True
            TabIndex        =   5
            Top             =   1185
            Width           =   3840
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Show DP "
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
            TabIndex        =   6
            Top             =   1605
            Width           =   1140
         End
         Begin VB.TextBox TxtRecOtherDeduction 
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
            Left            =   10575
            Locked          =   -1  'True
            TabIndex        =   72
            TabStop         =   0   'False
            Top             =   4260
            Width           =   2610
         End
         Begin VB.TextBox TxtRecTDS 
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
            Left            =   6375
            Locked          =   -1  'True
            TabIndex        =   70
            TabStop         =   0   'False
            Top             =   4260
            Width           =   2370
         End
         Begin VB.TextBox TxtRecAmount 
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
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   68
            TabStop         =   0   'False
            Top             =   4260
            Width           =   2730
         End
         Begin VB.TextBox TxtEntryNo 
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
            Left            =   1800
            TabIndex        =   4
            Top             =   1185
            Width           =   3870
         End
         Begin VB.TextBox TxtPostInvoiceNo 
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
            Left            =   7800
            MaxLength       =   25
            TabIndex        =   3
            Top             =   720
            Width           =   3840
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
            Left            =   3360
            Locked          =   -1  'True
            TabIndex        =   49
            TabStop         =   0   'False
            Top             =   1920
            Width           =   3495
         End
         Begin VB.TextBox TxtExchangeTypeID 
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
            Left            =   7800
            Locked          =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   1185
            Width           =   3810
         End
         Begin VB.TextBox TxtServiceTaxAmt 
            Alignment       =   1  'Right Justify
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
            Left            =   7800
            TabIndex        =   10
            Top             =   3315
            Width           =   3840
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
            Left            =   1815
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   1920
            Width           =   1575
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
            Height          =   360
            Left            =   12720
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   1920
            Width           =   375
         End
         Begin VB.CommandButton CmdLocationID 
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
            Left            =   5715
            TabIndex        =   45
            TabStop         =   0   'False
            Top             =   2865
            Width           =   375
         End
         Begin VB.TextBox TxtTINo 
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
            MaxLength       =   50
            TabIndex        =   2
            Top             =   720
            Width           =   3840
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
            ItemData        =   "FrmTxn_TIDetails.frx":0041
            Left            =   1800
            List            =   "FrmTxn_TIDetails.frx":0043
            Sorted          =   -1  'True
            TabIndex        =   8
            Top             =   2865
            Width           =   3855
         End
         Begin VB.TextBox TxtTIID 
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
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   240
            Width           =   1515
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
            Left            =   7680
            Locked          =   -1  'True
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   4725
            Width           =   4890
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
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   4725
            Width           =   3810
         End
         Begin VB.ComboBox CmbClientDP 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   6840
            Sorted          =   -1  'True
            TabIndex        =   40
            Top             =   1920
            Visible         =   0   'False
            Width           =   5745
         End
         Begin MSComCtl2.DTPicker TxtTIDate 
            Height          =   375
            Left            =   4320
            TabIndex        =   1
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
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
            Format          =   107937793
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
            Left            =   1815
            Locked          =   -1  'True
            TabIndex        =   51
            Top             =   2865
            Width           =   3840
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
            Left            =   6840
            Locked          =   -1  'True
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   1920
            Width           =   1215
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
            Left            =   8040
            Locked          =   -1  'True
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   1920
            Width           =   4530
         End
         Begin VB.TextBox TxtTIAmount 
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
            Left            =   1800
            MaxLength       =   12
            TabIndex        =   9
            Top             =   3315
            Width           =   3840
         End
         Begin VB.TextBox TxtServiceTaxID 
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
            Left            =   1815
            Locked          =   -1  'True
            MaxLength       =   25
            TabIndex        =   82
            Top             =   3780
            Visible         =   0   'False
            Width           =   3840
         End
         Begin VB.TextBox TxtNAVCode 
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
            Left            =   1800
            TabIndex        =   86
            Top             =   2400
            Width           =   3840
         End
         Begin VB.Label Label38 
            Caption         =   "Reverse Flag"
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
            Left            =   5865
            TabIndex        =   106
            Top             =   285
            Width           =   1410
         End
         Begin VB.Label Label36 
            Caption         =   "Collection Remark Updated By"
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
            TabIndex        =   103
            Top             =   6420
            Width           =   3015
         End
         Begin VB.Label Label34 
            Caption         =   "Collection Remark Updated Date"
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
            Left            =   5820
            TabIndex        =   102
            Top             =   6420
            Width           =   3015
         End
         Begin VB.Label Label32 
            Caption         =   "Collection Remark"
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
            Left            =   75
            TabIndex        =   99
            Top             =   5827
            Width           =   1695
         End
         Begin VB.Label Label31 
            Caption         =   "SM ID"
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
            Left            =   9840
            TabIndex        =   97
            Top             =   2460
            Width           =   1410
         End
         Begin VB.Label Label30 
            Caption         =   "Invoice Type"
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
            Left            =   5880
            TabIndex        =   95
            Top             =   2460
            Width           =   1410
         End
         Begin VB.Label Label21 
            Caption         =   "NAVISION Client Name"
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
            Left            =   120
            TabIndex        =   89
            Top             =   5175
            Width           =   1695
         End
         Begin VB.Label Label35 
            Caption         =   "NAV Code"
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
            TabIndex        =   85
            Top             =   2460
            Width           =   1500
         End
         Begin VB.Label Label20 
            Caption         =   "Service Tax Code"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   120
            TabIndex        =   83
            Top             =   3750
            Width           =   1695
         End
         Begin VB.Label Label17 
            Caption         =   "Received Other Deduction"
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
            Left            =   9000
            TabIndex        =   73
            Top             =   4200
            Width           =   1515
         End
         Begin VB.Label Label11 
            Caption         =   "Received TDS"
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
            TabIndex        =   71
            Top             =   4320
            Width           =   1410
         End
         Begin VB.Label Label7 
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
            Height          =   240
            Left            =   120
            TabIndex        =   69
            Top             =   4320
            Width           =   1695
         End
         Begin VB.Label Label6 
            Caption         =   "Entry No"
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
            Left            =   135
            TabIndex        =   67
            Top             =   1238
            Width           =   1815
         End
         Begin VB.Label Label1 
            Caption         =   "Post Invoice No"
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
            TabIndex        =   66
            Top             =   773
            Width           =   1815
         End
         Begin VB.Label Label28 
            Caption         =   "Service Tax Amount"
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
            Left            =   5880
            TabIndex        =   65
            Top             =   3225
            Width           =   1845
         End
         Begin VB.Label Label27 
            Caption         =   "Client ID"
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
            TabIndex        =   64
            Top             =   1980
            Width           =   975
         End
         Begin VB.Label Label26 
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
            Left            =   4560
            TabIndex        =   63
            Top             =   1680
            Width           =   1095
         End
         Begin VB.Label Label25 
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
            Left            =   7132
            TabIndex        =   62
            Top             =   1680
            Width           =   630
         End
         Begin VB.Label Label24 
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
            Left            =   9480
            TabIndex        =   61
            Top             =   1680
            Width           =   1095
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
            Left            =   7920
            TabIndex        =   60
            Top             =   1680
            Visible         =   0   'False
            Width           =   1110
         End
         Begin VB.Label Label5 
            Caption         =   "TI No (Voucher No)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   135
            TabIndex        =   59
            Top             =   668
            Width           =   1635
         End
         Begin VB.Label Label3 
            Caption         =   "TI Amount"
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
            Top             =   3345
            Width           =   1695
         End
         Begin VB.Label Label2 
            Caption         =   "Exchange Type"
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
            TabIndex        =   57
            Top             =   1238
            Width           =   1815
         End
         Begin VB.Label LblCMPName 
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
            Height          =   210
            Left            =   120
            TabIndex        =   56
            Top             =   2940
            Width           =   900
         End
         Begin VB.Label Label4 
            Caption         =   "TI ID"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   55
            Top             =   315
            Width           =   1800
         End
         Begin VB.Label Label8 
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
            Height          =   300
            Left            =   3525
            TabIndex        =   54
            Top             =   285
            Width           =   765
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
            Left            =   6180
            TabIndex        =   53
            Top             =   4785
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
            TabIndex        =   52
            Top             =   4830
            Width           =   1695
         End
      End
      Begin VB.Frame ButtonFrm 
         Height          =   1140
         Left            =   120
         TabIndex        =   32
         Top             =   8520
         Width           =   13320
         Begin VB.CommandButton CmdViewInvoice 
            Caption         =   "View Invoice"
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
            Height          =   810
            Left            =   10605
            TabIndex        =   87
            Top             =   195
            Width           =   2580
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
            Height          =   400
            Left            =   8505
            TabIndex        =   15
            Top             =   600
            Width           =   2100
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
            Left            =   6405
            TabIndex        =   38
            Top             =   600
            Width           =   2100
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
            Left            =   4305
            TabIndex        =   37
            Top             =   600
            Width           =   2100
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
            Left            =   2205
            TabIndex        =   36
            Top             =   600
            Width           =   2100
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
            Left            =   105
            TabIndex        =   35
            Top             =   600
            Width           =   2100
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
            Left            =   8505
            TabIndex        =   16
            Top             =   195
            Width           =   2100
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
            Left            =   6405
            TabIndex        =   34
            Top             =   195
            Width           =   2100
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
            Left            =   4305
            TabIndex        =   33
            Top             =   195
            Width           =   2100
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
            Left            =   2205
            TabIndex        =   14
            Top             =   195
            Width           =   2100
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
            Left            =   105
            TabIndex        =   0
            Top             =   195
            Width           =   2100
         End
      End
   End
End
Attribute VB_Name = "FrmTxn_TIDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg, mEntryFlag As Variant
Dim mExchangeTypeID As Double
Dim mClientIDRow As Double
Dim ans As Variant
Dim mLocationID As Double
Dim mDPID As String
Dim mClientNAVCode As Variant
Dim mServiceTaxID, mServiceTaxNAVCode As Variant
Dim mEmployeeID As Variant

Private Sub AddCmd_Click()

Call GButtonLock(Me, False)
Edit_Flg = "A"
Frame1.Enabled = True
Frame2.Enabled = True
TxtRemarks = ""

TxtTIID = ""
TxtClientID = ""
TxtClientName = ""
TxtDPID = ""
TxtDPName = ""
CmbClientDP.Text = ""
TxtNAVClientName = ""
TxtCRemark = ""
TxtCUpdatedBy = ""
TxtCUpdatedDate = ""
CmbClientDP.Visible = False

mExchangeTypeID = 0
CmbExchangeTypeID.Visible = True
CmbExchangeTypeID.Text = ""
TxtExchangeTypeID.Visible = False


mClientIDRow = 0
mDPID = ""
mServiceTaxID = ""

CmbCNonNcdexClient.Visible = False
CmbCNonNcdexClient.Text = ""
Check3.Value = 0
CmbDPName.Text = ""
TxtTINo = ""
TxtTIDate = Date
TxtPostInvoiceNo = ""
TxtEntryNo = ""
TxtTIAmount = ""
TxtServiceTaxAmt = ""
TxtRecAmount = ""
TxtRecTDS = ""
TxtRecOtherDeduction = ""
mEntryFlag = "Manual"
TxtAdjAmount = ""
TxtInvType = ""
TxtSMID = ""
TxtAdjAmount.Locked = False
TxtRemarks.Locked = False

TxtTIAmount.Locked = False
TxtServiceTaxAmt.Locked = False

CmbLocationID.Visible = True
CmbLocationID.Text = ""
TxtLocationID.Visible = False
CmdLocationID.Visible = True

TxtServiceTaxID.Visible = False
CmbServiceTaxID.Visible = True
CmdServiceTaxID.Visible = True
CmbServiceTaxID.Text = ""

TxtReverseFlag = ""
TxtNAVCode.Visible = False
CmbNAVCode.Visible = True
CmbNAVCode.Text = ""

mServiceTaxNAVCode = ""

CmdViewInvoice.Enabled = False



End Sub

Private Sub CmbExchangeTypeID_GotFocus()
tmp = CmbExchangeTypeID.Text
Call TableMst_ExchangeType
CmbExchangeTypeID.Clear
If RsMst_ExchangeType.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_ExchangeType.RecordCount
    CmbExchangeTypeID.AddItem RsMst_ExchangeType!ExchangeType
    RsMst_ExchangeType.MoveNext
Next
CmbExchangeTypeID.Text = tmp
End Sub
'----
Private Sub CmbExchangeTypeID_LostFocus()
If CmbExchangeTypeID.Text = "" Then
   mExchangeTypeID = 0
   CmbCNonNcdexClient.Clear
   CmbClientDP.Clear
   Exit Sub
End If
RsMst_ExchangeType.MoveFirst
RsMst_ExchangeType.Find "ExchangeType = '" & CmbExchangeTypeID.Text & "'"
If RsMst_ExchangeType.EOF Then
   MsgBox "Invalid selection "
   CmbExchangeTypeID.SetFocus
   Exit Sub
End If

If mExchangeTypeID <> RsMst_ExchangeType!ExchangeTypeID Then
   CmbCNonNcdexClient.Clear
   CmbClientDP.Clear
   CmbDPName.Clear
End If
mExchangeTypeID = RsMst_ExchangeType!ExchangeTypeID

If mExchangeTypeID = 1 Then
   CmbCNonNcdexClient.Width = 10800
   CmbCNonNcdexClient.Top = TxtClientID.Top
   CmbCNonNcdexClient.Left = TxtClientID.Left
   CmbCNonNcdexClient.Visible = True
   Label25.Visible = False
   Label27.Visible = False
   Label24.Visible = False
   Check3.Value = 0
   Check3.Enabled = False
Else
   Label25.Visible = True
   Label27.Visible = True
   Label24.Visible = True
   CmbCNonNcdexClient.Visible = False
   Check3.Enabled = True
   Call Check3_Click
End If

End Sub

Private Sub CmbSEntryType_LostFocus()

If CmbSEntryType.Text = "" Then Exit Sub

If LCase(CmbSEntryType.Text) <> "system" And LCase(CmbSEntryType.Text) <> "manual" Then
   MsgBox "Invalid Selection!!!"
   CmbSEntryType.SetFocus
   Exit Sub
End If

End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Txn_TIDetails.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Txn_TIDetails.xls")
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

Private Sub CmdGo_Click()

Dim wc As Variant

wc = ""

Call Grid_Head

tmp = " Select TI.TIID,TI.TIDate,TI.TINo,TI.PostInvoiceNo,TI.EntryNo,TIAmount,TI.TISTAmount,"
''tmp = tmp & " (select Sum(RecAmount) From Txn_NewCashReceiptTIDetail Where TIID=TI.TIID ) as RecAmount,"
''tmp = tmp & " (select Sum(RecTDS) From Txn_NewCashReceiptTIDetail Where TIID=TI.TIID ) as RecTDSAmount ,"
''tmp = tmp & " (select Sum(RecOtherDed) From Txn_NewCashReceiptTIDetail Where TIID=TI.TIID ) as RecOtherDed"
tmp = tmp & " (Select isnull(Sum(TCRT.RecAmount),0) as RecAmount From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.NewCRID= TCRT.NewCRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID) as RecAmount,"
tmp = tmp & " (Select isnull(Sum(TCRT.RecTDS),0) as RecTDS From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.NewCRID= TCRT.NewCRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID) + (Select isnull(Sum(TCRT.RecTDS),0) as RecTDS From Txn_CRTDSAdjustmentDetail TCRT Inner Join Txn_CRTDSAdjustment TCR on TCR.CRTDSAdjustmentID= TCRT.CRTDSAdjustmentID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.Flag<>'R'And TIID=TI.TIID )RecTDSAmount,"
tmp = tmp & " (Select isnull(Sum(TCRT.RecOtherDed),0) as RecOtherDed From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.NewCRID= TCRT.NewCRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID)+ (Select isnull(Sum(TCRT.RecOtherDed),0) as RecOtherDed From Txn_CRTDSAdjustmentDetail TCRT Inner Join Txn_CRTDSAdjustment TCR on TCR.CRTDSAdjustmentID= TCRT.CRTDSAdjustmentID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.Flag<>'R'And TIID=TI.TIID )RecOtherDed"
tmp = tmp & " ,TI.NAVCode,ML.LocationName,TI.EntryFlag,TI.AdjAmountFAG,TI.Remarks,TI.PIFlag,TI.PIID,TI.ReverseFlag, TI.UpdatedBy, TI.UpdatedDate"
tmp = tmp & " From  Txn_TIDetails TI"
tmp = tmp & " Left Join Mst_Location ML on TI.LocationID=ML.LocationID"
tmp = tmp & " Where  TINo<> '' "

If Gen_DBType = "MDB" Then
   wc = " And TI.TIDate Between #" & Format(SFrom.Value, Gen_DatFormat) & "# And #" & Format(STo.Value, Gen_DatFormat) & "# "
Else
   wc = " And TI.TIDate Between '" & Format(SFrom.Value, Gen_DatFormat) & "' And '" & Format(STo.Value, Gen_DatFormat) & "' "
End If


If TxtSTINo <> "" Then
   wc = wc & " And TI.TINO like '%" & TxtSTINo & "%'"
End If

If TxtSPostInvoiceNo <> "" Then
   wc = wc & " And TI.PostInvoiceNo like '%" & TxtSPostInvoiceNo & "%'"
End If

If TxtSEntryNo <> "" Then
   wc = wc & " And TI.EntryNo like '%" & TxtSEntryNo & "%'"
End If

If CmbSLocation.Text <> "" Then
   wc = wc & " And ML.LocationName = '" & CmbSLocation.Text & "'"
End If

If CmbSEntryType.Text <> "" Then
   If LCase(CmbSEntryType.Text) = "system" Then
      wc = wc & " And (TI.EntryFlag = '" & CmbSEntryType.Text & "' Or TI.EntryFlag is Null) "
   Else
      wc = wc & " And TI.EntryFlag = '" & CmbSEntryType.Text & "'"
   End If
End If

If LCase(CmbSReverseFlag.Text) = "yes" Then
   wc = wc & " And isnull(TI.ReverseFlag,'N') = 'Y'"
ElseIf LCase(CmbSReverseFlag.Text) = "no" Then
   wc = wc & " And isnull(TI.ReverseFlag,'N') = 'N'"
End If
If CmbEmployeeID.Text <> "" Then
   strSplitString = Split(CmbEmployeeID, "~")
   wc = wc & " And TI.UpdatedBy = '" & Trim(strSplitString(1)) & "'"
End If

If Trim(TxtSNAVCode) <> "" Then
   wc = wc & " And TI.NAVCode like '%" & TxtSNAVCode & "%'"
End If

tmp = tmp & wc

Call TmpTable


For iu = 1 To TmpRs.RecordCount
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
    For C = 0 To MSHFlexGrid1.Cols - 2
       MSHFlexGrid1.TextMatrix(iu, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
    Next
    MSHFlexGrid1.TextMatrix(iu, 1) = Format(MSHFlexGrid1.TextMatrix(iu, 1), "dd-mmm-yyyy")
    If MSHFlexGrid1.TextMatrix(iu, 12) = "" Then
       MSHFlexGrid1.TextMatrix(iu, 12) = "System"
    End If
    
    If MSHFlexGrid1.TextMatrix(iu, 17) = "" Or MSHFlexGrid1.TextMatrix(iu, 17) = "N" Then
       MSHFlexGrid1.TextMatrix(iu, 17) = "No"
    Else
       MSHFlexGrid1.TextMatrix(iu, 17) = "Yes"
    End If
    MSHFlexGrid1.TextMatrix(iu, 18) = IIf(IsNull(RsTxn_TIDetails!UpdatedBy), "", RsTxn_TIDetails!UpdatedBy & " :- " & RsTxn_TIDetails!UpdatedDate)
    
    TmpRs.MoveNext
Next

Label15.Caption = TmpRs.RecordCount

End Sub

Private Sub CmdServiceTaxID_Click()
FrmMst_ServiceTax.Show
End Sub

Private Sub CmdViewInvoice_Click()
Dim fso As New FileSystemObject

On Error GoTo Error:
'Shell (App.Path & "\Invoices\" & TxtEntryNo & " ")
If TxtEntryNo <> "" Then
   wsPathAdobe = Gen_ReadPathConfigINI("PDFPATH") '"\\10.0.0.8\Reader\AcroRD32.exe" '"C:\Program Files\Adobe\Acrobat 6.0\Reader\AcroRd32.exe" '"\\10.0.0.8\Reader\AcroRD32.exe"
   mFileName = App.Path & "\NavisionInvoice\" & TxtEntryNo & ".pdf"
   'mFileName = App.Path & "\Invoices\" & Replace(TxtPostInvoiceNo, "/", "") & ".pdf"
'   If wsPathAdobe <> "" Then
'      iRet = Shell(wsPathAdobe & " " & mFileName, vbMaximizedFocus)
'   End If
   If fso.FileExists(mFileName) = True Then
      If wsPathAdobe <> "" Then
         iRet = Shell(wsPathAdobe & " " & mFileName, vbMaximizedFocus)
         Call Log_Print_History("", "TaxInvoice", TxtTIID, TxtEntryNo)
      End If
   Else
      MsgBox "Invoice not found. Please contact CMG Department for same!!!"
   End If
End If
Exit Sub

Error:
      MsgBox "Please contact Technology department!!!"

End Sub

Private Sub EditCmd_Click()


'If Val(TxtTIAmount) = Val(TxtRecAmount) + Val(TxtRecTDS) + Val(TxtRecOtherDeduction) + Val(TxtAdjAmount) Then
If Format(Val(TxtTIAmount), "#######0.00") = Format(Val(TxtRecAmount) + Val(TxtRecTDS) + Val(TxtRecOtherDeduction) + Val(TxtAdjAmount), "#######0.00") Then
   MsgBox "Invoice Amount is equal to Received Amount. Can't Edit!!!"
   Exit Sub
End If

Edit_Flg = "E"
Call GButtonLock(Me, False)

Frame2.Enabled = True

If LCase(mEntryFlag) <> "manual" Then
'   MsgBox "TI No is system entry can't edit!!!"
'   Exit Sub
   TxtAdjAmount.Locked = False
   TxtRemarks.Locked = False
   Exit Sub
End If

If Val(TxtRecAmount) > 0 Then
'   MsgBox "Some amount is received against this TI No. Can't edit!!!"
'   Exit Sub
   TxtAdjAmount.Locked = False
   TxtRemarks.Locked = False
   Exit Sub
End If


If LCase(mEntryFlag) = "manual" And Val(TxtRecAmount) = 0 Then

   CmbExchangeTypeID.Visible = True
   CmbExchangeTypeID.Text = TxtExchangeTypeID
   TxtExchangeTypeID.Visible = False
   
   CmbCNonNcdexClient.Visible = False
   CmbCNonNcdexClient.Text = ""
    
   CmbLocationID.Visible = True
   CmbLocationID.Text = TxtLocationID
   TxtLocationID.Visible = False
   CmdLocationID.Visible = True
   
   TxtServiceTaxID.Visible = False
   CmbServiceTaxID.Visible = True
   CmdServiceTaxID.Visible = True
   CmbServiceTaxID.Text = TxtServiceTaxID
   
   
   TxtNAVCode.Visible = False
   CmbNAVCode.Visible = True
   CmbNAVCode.Text = TxtNAVCode
   CmdViewInvoice.Enabled = False

   TxtTIAmount.Locked = False
   TxtServiceTaxAmt.Locked = False
   
   TxtAdjAmount.Locked = False
   TxtRemarks.Locked = False

   Frame1.Enabled = True


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
Private Sub Form_Load()

Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Call GButtonLockAD(Me)
   MsgBox "Access denied !!!!!!!!!"
   CmdViewInvoice.Enabled = False
   Exit Sub
End If

wc = ""
If Gen_WcLocation <> "" Then
   wc = "Where " & Gen_WcLocation
End If
Call Grid_Head
SFrom.Value = Date - 30
STo.Value = Date
TxtSTINo = ""
TxtSPostInvoiceNo = ""
TxtSEntryNo = ""
CmbSLocation.Text = ""
CmbSEntryType.Text = ""
Label15.Caption = "-"

Call TableTxn_TIDetails(" " & wc & "  ") 'Order by SM_Prefix,CRID")
If RsTxn_TIDetails.RecordCount = 0 Then
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
Private Sub MSHFlexGrid1_Click()
Dim I As Variant
I = MSHFlexGrid1.RowSel
If I <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(I, 0) = "" Then Exit Sub
RsTxn_TIDetails.MoveFirst
'RsTxn_NewCashReceipt.Find "CRID = " & Val(MSHFlexGrid1.TextMatrix(I, 0))
RsTxn_TIDetails.Find "TIID= " & MSHFlexGrid1.TextMatrix(I, 0) & ""

If Not RsTxn_TIDetails.EOF Then
   Call Indata
   Frame0.Visible = False
End If
End Sub

Private Sub SaveCmd_Click()

If Frame1.Enabled = True Then
   If TxtTIDate.Value > Date Then
      MsgBox "Future Date not allowed!!!"
      TxtTIDate.SetFocus
      Exit Sub
   End If
   
   If TxtTINo = "" Then
      MsgBox "Blank TI No not allowed!!!"
      TxtTINo.SetFocus
      Exit Sub
   End If
   
   If TxtPostInvoiceNo = "" Then
      MsgBox "Blank Post Invoice No not allowed!!!"
      TxtPostInvoiceNo.SetFocus
      Exit Sub
   End If
   
   If TxtEntryNo = "" Then
      MsgBox "Blank Entry No not allowed!!!"
      TxtEntryNo.SetFocus
      Exit Sub
   End If
   
   If CmbExchangeTypeID.Text = "" Then
      MsgBox "Blank Exchange Type  not allowed!!!"
      CmbExchangeTypeID.SetFocus
      Exit Sub
   Else
      If Check3.Value = 0 Then
         If CmbClientDP.Visible = True Then
            MsgBox "Invalid selection !!! "
            Exit Sub
         End If
         If CmbCNonNcdexClient.Visible = True Then
            If CmbCNonNcdexClient.Text = "" Then
               MsgBox "Invalid selection !!! "
               CmbCNonNcdexClient.SetFocus
               Exit Sub
            End If
         End If
         If TxtClientID = "" Then
            MsgBox "Invalid Client ID !!! "
            Exit Sub
         End If
      Else
         If CmbDPName.Visible = True Then
            If CmbDPName.Text = "" Then
               MsgBox "Blank DP Name Not allowed !!! "
               CmbDPName.SetFocus
               Exit Sub
            End If
         End If
      End If
   End If
   
   If CmbNAVCode.Text = "" Then
      MsgBox "Blank NAV Code not allowed!!!"
      CmbNAVCode.SetFocus
      Exit Sub
   End If
   
   If TxtTIAmount = "" Then
      MsgBox "Blank TI Amount not allowed!!!"
      TxtTIAmount.SetFocus
      Exit Sub
   End If
   
   If TxtServiceTaxAmt = "" Then
      MsgBox "Blank Service Tax Amount not allowed!!!"
      TxtServiceTaxAmt.SetFocus
      Exit Sub
   End If
   
   If CmbServiceTaxID.Text = "" Then
      MsgBox "Blank Service Tax Code not allowed!!!"
      CmbServiceTaxID.SetFocus
      Exit Sub
   End If
End If

If Val(TxtAdjAmount) <> 0 Then
   If TxtRemarks = "" Then
      MsgBox "Blank Remarks not allowed!!!"
      TxtRemarks.SetFocus
      Exit Sub
   End If
   'If Val(TxtTIAmount) < Val(TxtRecAmount) + Val(TxtAdjAmount) Then
   'If Val(TxtTIAmount) = Val(TxtRecAmount) + Val(TxtRecTDS) + Val(TxtRecOtherDeduction) + Val(TxtAdjAmount) Then
   
'   If Val(TxtRecAmount) = "0.00" And Val(TxtRecTDS) = "0.00" And Val(TxtRecOtherDeduction) = "0.00" And Val(TxtAdjAmount) > 0 Then
'
'      If Val(TxtTIAmount) < Val(TxtRecAmount) + Val(TxtRecTDS) + Val(TxtRecOtherDeduction) + Val(TxtAdjAmount) Then
'         MsgBox "Sum of all Received Amount and Adjustment Amount must be less than or equal to Invoice Amount!!!"
'         TxtRemarks.SetFocus
'         Exit Sub
'      End If
'
'   Else
'      If Format(Val(TxtTIAmount), "#######0.00") < Format(Val(TxtRecAmount) + Val(TxtRecTDS) + Val(TxtRecOtherDeduction) + Val(TxtAdjAmount), "#######0.00") Then
'         MsgBox "Sum of all Received Amount and Adjustment Amount must be less than or equal to Invoice Amount!!!"
'         TxtRemarks.SetFocus
'         Exit Sub
'      End If
'   End If
   
   mRecA = Val(TxtRecAmount) + Val(TxtRecTDS) + Val(TxtRecOtherDeduction) + Val(TxtAdjAmount)
   mTA = Val(TxtTIAmount)
   
   If Val(mTA) < Val(mRecA) Then
      MsgBox "Sum of all Received Amount and Adjustment Amount must be less than or equal to Invoice Amount!!!"
      TxtRemarks.SetFocus
      Exit Sub
   End If
   
   
End If

If Edit_Flg = "A" Then
'   If RsTxn_TIDetails.RecordCount > 0 Then
'      RsTxn_TIDetails.MoveFirst
'      RsTxn_TIDetails.Find "EntryNo = '" & TxtAccountNo & "'"
'      If Not RsTxn_TIDetails.EOF Then
'         MsgBox "Duplicate Account No. not allowed. Already exist for " & RsTxn_TIDetails!BankName & " !!!"
'         TxtAccountNo.SetFocus
'         Exit Sub
'      End If
'      RsTxn_TIDetails.MoveLast
'   End If
   RsTxn_TIDetails.AddNew
   RsTxn_TIDetails!TIID = GetMaxTIID
   RsTxn_TIDetails!G_UID = GetGUID
   RsTxn_TIDetails!EntryFlag = mEntryFlag
   TxtTIID = RsTxn_TIDetails!TIID
   RsTxn_TIDetails!TIDate = TxtTIDate
    
Else
'   RsTxn_TIDetails.MoveFirst
'   RsTxn_TIDetails.Find "AccountNo = '" & TxtAccountNo & "'"
'   If Not RsTxn_TIDetails.EOF Then
'       If RsTxn_TIDetails!FAGBankId <> TxtBankID Then
'          MsgBox "Duplicate Account No. not allowed. Already exist for " & RsTxn_TIDetails!BankName & " !!!"
'          TxtAccountNo.SetFocus
'          Exit Sub
'       End If
'   End If
   RsTxn_TIDetails.MoveFirst
   RsTxn_TIDetails.Find "TIID = " & TxtTIID & ""
End If

RsTxn_TIDetails!ExchangeTypeID = mExchangeTypeID
RsTxn_TIDetails!TINo = TxtTINo
RsTxn_TIDetails!PostInvoiceNo = TxtPostInvoiceNo
RsTxn_TIDetails!EntryNo = TxtEntryNo

If CmbLocationID.Text <> "" Then
   RsTxn_TIDetails!LocationID = mLocationID
Else
   RsTxn_TIDetails!LocationID = Null
End If

If Check3.Value = 0 Then
   RsTxn_TIDetails!ClientID = mClientIDRow
Else
   RsTxn_TIDetails!DepositoryID = mDPID
End If

RsTxn_TIDetails!NavCode = CmbNAVCode.Text  'mClientNAVCode
RsTxn_TIDetails!TIAmount = Val(TxtTIAmount)
RsTxn_TIDetails!TISTAmount = Val(TxtServiceTaxAmt)
If mServiceTaxID <> "" Then
   RsTxn_TIDetails!ServiceTaxID = mServiceTaxID
End If
RsTxn_TIDetails!AdjAmountFAG = Val(TxtAdjAmount)
RsTxn_TIDetails!Remarks = Trim(TxtRemarks)

If IsNull(RsTxn_TIDetails!CreatedBy) = True Or RsTxn_TIDetails!CreatedBy = "" Then
   RsTxn_TIDetails!CreatedBy = Gen_UserLoginID
   RsTxn_TIDetails!CreatedDate = Now
Else
   RsTxn_TIDetails!UpdatedBy = Gen_UserLoginID
   RsTxn_TIDetails!UpdatedDate = Now
End If

RsTxn_TIDetails.Update
RsTxn_TIDetails.Requery
RsTxn_TIDetails.MoveFirst
RsTxn_TIDetails.Find "TIID = " & TxtTIID & ""

If RsTxn_TIDetails.EOF Then
   RsTxn_TIDetails.MoveFirst
End If

Call GButtonLock(Me, True)
Call Indata

End Sub
Private Function GetMaxTIID() As Double
Dim Retval As Double
tmp = "Select max(TIID) from Txn_TIDetails"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetMaxTIID = Retval
End Function

Private Sub SearchCmd_Click()

If SearchCmd.Caption = "Cancel" Then
   If RsTxn_TIDetails.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   RsTxn_TIDetails.MoveFirst
   RsTxn_TIDetails.Find "G_UID= '" & TxtG_UID & "'"
   If RsTxn_TIDetails.EOF Then
      RsTxn_TIDetails.MoveLast
   End If
   Call Indata
   Exit Sub
End If
Frame0.Height = Frame4.Height - 200
Frame0.Left = Frame1.Left - 100
Frame0.Top = Frame1.Top - 100
Frame0.Width = Frame4.Width - 150
Frame0.Visible = True
MSHFlexGrid1.Width = Frame0.Width - 400
MSHFlexGrid1.Height = Frame0.Height - 2800

''===========comment on 7 july 2012 for not clearing search grid=============
'MSHFlexGrid1.Clear
'MSHFlexGrid1.Cols = 11
'Call Grid_Head
'SFrom.Value = Date - 30
'STo.Value = Date
'TxtSTINo = ""
'TxtSPostInvoiceNo = ""
'TxtSEntryNo = ""
'CmbSLocation.Text = ""
'CmbSEntryType.Text = ""
'Label15.Caption = "-"
''===========comment on 7 july 2012 for not clearing search grid=============

End Sub

Private Sub TxtAdjAmount_LostFocus()
If TxtAdjAmount = "" Then Exit Sub

If IsNumeric(TxtAdjAmount) = False Or Val(TxtAdjAmount) < 0 Then
   MsgBox "Invalid format!!!"
   TxtAdjAmount.SetFocus
   Exit Sub
End If

If InStr(1, TxtAdjAmount, ",") > 0 Then
   TxtAdjAmount = Replace(TxtAdjAmount, ",", "")
End If

End Sub

Private Sub TxtEntryNo_LostFocus()
If TxtEntryNo = "" Then Exit Sub

If IsNumeric(TxtEntryNo) = False Or Val(TxtEntryNo) < 0 Then
   MsgBox ("Invalid Entry No.!!!")
   TxtEntryNo.SetFocus
   Exit Sub
End If

RsTxn_TIDetails.MoveFirst
RsTxn_TIDetails.Find "EntryNo = '" & TxtEntryNo & "'"
If Not RsTxn_TIDetails.EOF Then
   If RsTxn_TIDetails!TIID <> TxtTIID Then
      MsgBox ("Entry No. already exist for TIID - " & RsTxn_TIDetails!TIID & " !!!")
      TxtEntryNo.SetFocus
      Exit Sub
   End If
End If

End Sub

Private Sub TxtPostInvoiceNo_LostFocus()

If TxtPostInvoiceNo = "" Then Exit Sub

TxtPostInvoiceNo = Replace(TxtPostInvoiceNo, Chr(32), "")

RsTxn_TIDetails.MoveFirst
RsTxn_TIDetails.Find "PostInvoiceNo = '" & TxtPostInvoiceNo & "'"
If Not RsTxn_TIDetails.EOF Then
   If RsTxn_TIDetails!TIID <> TxtTIID Then
      ans = MsgBox("Post Invoice No. already exist for TIID - " & RsTxn_TIDetails!TIID & " !!!", vbYesNo)
      If ans = vbNo Then
         TxtPostInvoiceNo.SetFocus
         Exit Sub
      End If
   End If
End If

End Sub

Private Sub TxtTIAmount_LostFocus()
If TxtTIAmount = "" Then Exit Sub

If IsNumeric(TxtTIAmount) = False Or Val(TxtTIAmount) < 0 Then
   MsgBox "Invalid format!!!"
   TxtTIAmount.SetFocus
   Exit Sub
End If


End Sub

Private Sub TxtServiceTaxAmt_LostFocus()

If TxtServiceTaxAmt = "" Then Exit Sub

If IsNumeric(TxtServiceTaxAmt) = False Or Val(TxtServiceTaxAmt) < 0 Then
   MsgBox "Invalid format!!!"
   TxtServiceTaxAmt.SetFocus
   Exit Sub
End If


End Sub

Private Sub TxtTIDate_LostFocus()

'Commented on 25 March 2013 after mail from Mayur Sir as mail by Basant Acharya to him on 25th March 2013
'If Format(Gen_BlockingDate, "yyyy-mm") >= Format(TxtTIDate.Value, "yyyy-mm") Then
'   MsgBox "TI Date allowed for next month of " & Format(Gen_BlockingDate, "MMM-yyyy") & "!!!"
'   TxtTIDate.SetFocus
'   Exit Sub
'End If

End Sub

Private Sub TxtTINo_LostFocus()

If TxtTINo = "" Then Exit Sub

TxtTINo = Replace(TxtTINo, Chr(32), "")

RsTxn_TIDetails.MoveFirst
RsTxn_TIDetails.Find "TINo = '" & TxtTINo & "'"
If Not RsTxn_TIDetails.EOF Then
   If RsTxn_TIDetails!TIID <> TxtTIID Then
      ans = MsgBox("TI No. already exist for TIID - " & RsTxn_TIDetails!TIID & " !!!", vbYesNo)
      If ans = vbNo Then
         TxtTINo.SetFocus
         Exit Sub
      End If
   End If
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

Private Sub CmbDPName_GotFocus()
If CmbExchangeTypeID.Text = "" Then
   MsgBox "Please select Exchange Type!!!"
   CmbDPName.Clear
   CmbExchangeTypeID.SetFocus
   Exit Sub
End If
tmp = CmbDPName.Text

Call TableMst_Depository("Where ExchangeTypeId = " & mExchangeTypeID & " ")

CmbDPName.Clear
If RsMst_Depository.RecordCount = 0 Then
   MsgBox "No record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Depository.RecordCount
    CmbDPName.AddItem RsMst_Depository!DepositoryName & "~" & RsMst_Depository!DepositoryID
    RsMst_Depository.MoveNext
Next
CmbDPName.Text = tmp

End Sub

Private Sub CmbDPName_LostFocus()

Dim strSplitString() As String
Dim DPID_, DPName_ As String

If CmbDPName.Text = "" Then
   TxtCDPID = ""
   TxtCDPName = ""
   Exit Sub
End If
strSplitString = Split(CmbDPName, "~")
DPName_ = Trim(strSplitString(0))
DPID_ = Trim(strSplitString(1))

RsMst_Depository.MoveFirst
RsMst_Depository.Find "DepositoryID = '" & DPID_ & "'"
If RsMst_Depository.EOF Then
   MsgBox "Invalid selection "
   CmbDPName.SetFocus
   Exit Sub
End If

mDPID = RsMst_Depository!DepositoryID
mDPNAVCode = IIf(IsNull(RsMst_Depository!NavCode), "", RsMst_Depository!NavCode)
TxtCDPID = mDPID
TxtCDPName = DPName_ 'CmbDPName.Text
TxtNAVClientName = IIf(IsNull(RsMst_Depository!NAVDepositoryName), "", RsMst_Depository!NAVDepositoryName)
CmbNAVCode.Clear
CmbNAVCode.AddItem mDPNAVCode
CmbNAVCode.Text = mDPNAVCode

End Sub

Private Sub Check3_Click()

If Check3.Value = 1 Then
   Label27.Visible = False
   Label25.Visible = False
   Label24.Visible = False
   CmbDPName.Visible = True
   CmbDPName.Width = 10800 '9400 'CmbNonNcdexClient.Width
   CmbDPName.Top = TxtClientID.Top 'CmbNonNcdexClient.Top
   CmbDPName.Left = TxtClientID.Left 'CmbNonNcdexClient.Left
   CmbCNonNcdexClient.Visible = False
Else
   Label27.Visible = True
   Label25.Visible = True
   Label24.Visible = True
'   CmbNonNcdexClient.Visible = True
'   CmbNonNcdexClient.Width = CmbDPName.Width
'   CmbNonNcdexClient.Top = CmbDPName.Top
'   CmbNonNcdexClient.Left = CmbDPName.Left
   CmbDPName.Visible = False
End If

End Sub

Private Sub CmbClientDP_LostFocus()
Dim Str_() As String
Str_ = Split(CmbClientDP.Text, "~")
If UBound(Str_) <> 1 Then
   MsgBox "Invalid selection !!!!! "
   If CmbClientDP.Visible = True Then
      CmbClientDP.SetFocus
   Else
      CmbClientDP.SetFocus
   End If
   Exit Sub
End If
TmpRs1.MoveFirst


TmpRs1.MoveFirst
TmpRs1.Find "DepositoryID  = '" & Str_(0) & "'"
If TmpRs1.EOF Then
   MsgBox "Invalid selection !!!!! "
   CmbClientDP.SetFocus
   Exit Sub
End If

TxtCClientName = TmpRs1!ClientName
TxtCDPID = TmpRs1!DepositoryID
TxtCDPName = TmpRs1!DepositoryName
TxtNAVClientName = IIf(IsNull(TmpRs1!NAVClientName), "", TmpRs1!NAVClientName)
CmbClientDP.Visible = False
mClientIDRow = TmpRs1!ClientID
mClientNAVCode = IIf(IsNull(TmpRs1!NavCode), "", TmpRs1!NavCode)
mClientRNAVCode = IIf(IsNull(TmpRs1!RNAVCode), "", TmpRs1!RNAVCode)

CmbNAVCode.Clear
If mClientRNAVCode <> "" And mClientNAVCode <> "" Then
   CmbNAVCode.AddItem mClientNAVCode
   CmbNAVCode.AddItem mClientRNAVCode
ElseIf mClientNAVCode <> "" Then
   CmbNAVCode.AddItem mClientNAVCode
   CmbNAVCode.Text = mClientNAVCode
ElseIf mClientRNAVCode <> "" Then
   CmbNAVCode.AddItem mClientRNAVCode
   CmbNAVCode.Text = mClientRNAVCode
End If


'CmdCAddToList.SetFocus
End Sub

Private Sub CmbCNonNcdexClient_GotFocus()

tmp = CmbCNonNcdexClient.Text
Call TableMst_Client("where ExchangeTypeId = " & mExchangeTypeID)
CmbCNonNcdexClient.Clear
If RsMst_Client.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Client.RecordCount
    CmbCNonNcdexClient.AddItem RsMst_Client!ClientName
    RsMst_Client.MoveNext
Next
CmbCNonNcdexClient.Text = tmp
End Sub

Private Sub CmbCNonNcdexClient_LostFocus()
If CmbCNonNcdexClient.Text = "" Then
   Exit Sub
End If
RsMst_Client.MoveFirst
RsMst_Client.Find "ClientName = '" & CmbCNonNcdexClient.Text & "'"
If RsMst_Client.EOF Then
   MsgBox "Invalid selection "
   CmbCNonNcdexClient.SetFocus
   Exit Sub
End If
mClientIDRow = RsMst_Client!ClientID
mClientNAVCode = IIf(IsNull(RsMst_Client!NavCode), "", RsMst_Client!NavCode)
mClientRNAVCode = IIf(IsNull(RsMst_Client!RNAVCode), "", RsMst_Client!RNAVCode)

CmbNAVCode.Clear
If mClientRNAVCode <> "" And mClientNAVCode <> "" Then
   CmbNAVCode.AddItem mClientNAVCode
   CmbNAVCode.AddItem mClientRNAVCode
ElseIf mClientNAVCode <> "" Then
   CmbNAVCode.AddItem mClientNAVCode
   CmbNAVCode.Text = mClientNAVCode
ElseIf mClientRNAVCode <> "" Then
   CmbNAVCode.AddItem mClientRNAVCode
   CmbNAVCode.Text = mClientRNAVCode
End If


Dim ans As Variant
ans = GetClientNameByRow(RsMst_Client!ClientID)
TxtClientID = ans(0)
TxtCClientName = ans(1)
TxtCDPID = ans(2)
TxtCDPName = ans(3)
TxtNAVClientName = ans(5)
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
tmp = " SELECT Mst_Depository.DepositoryID, Mst_Depository.DepositoryName, Mst_Client.ClientID, " & _
      " Mst_Client.ClientName,Mst_Client.ClientIDRow ,Mst_Client.NAVCode,Mst_Client.RNAVCode,Mst_Client.NAVClientName " & _
      " FROM Mst_Depository INNER JOIN Mst_Client ON Mst_Depository.DepositoryID = Mst_Client.DepositoryID " & _
      " WHERE Mst_Depository.ExchangeTypeID = " & mExchangeTypeID & _
      "  And  (((Mst_Client.ExchangeTypeID)=" & mExchangeTypeID & ") AND ((Mst_Client.ClientID)='" & TxtClientID & "'));"

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

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsTxn_TIDetails.MoveNext
If RsTxn_TIDetails.EOF Then
   RsTxn_TIDetails.MoveLast
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
RsTxn_TIDetails.MovePrevious
If RsTxn_TIDetails.BOF Then
   RsTxn_TIDetails.MoveFirst
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
RsTxn_TIDetails.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsTxn_TIDetails.MoveLast
Call Indata
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False
End Sub
Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub Grid_Head()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 19

MSHFlexGrid1.TextMatrix(0, 0) = "TI ID"
MSHFlexGrid1.TextMatrix(0, 1) = "TI Date"
MSHFlexGrid1.TextMatrix(0, 2) = "TI No"
MSHFlexGrid1.TextMatrix(0, 3) = "Post Invoice No"
MSHFlexGrid1.TextMatrix(0, 4) = "Entry No"
MSHFlexGrid1.TextMatrix(0, 5) = "TI Amount"
MSHFlexGrid1.TextMatrix(0, 6) = "TI Service Tax Amount"
MSHFlexGrid1.TextMatrix(0, 7) = "Received Amount"
MSHFlexGrid1.TextMatrix(0, 8) = "Received TDS"
MSHFlexGrid1.TextMatrix(0, 9) = "Received Other Ded."
MSHFlexGrid1.TextMatrix(0, 10) = "Client NAV Code"
MSHFlexGrid1.TextMatrix(0, 11) = "Location Name"
MSHFlexGrid1.TextMatrix(0, 12) = "Entry Flag"
MSHFlexGrid1.TextMatrix(0, 13) = "Adjustment Entry by FAG"
MSHFlexGrid1.TextMatrix(0, 14) = "Remarks"
MSHFlexGrid1.TextMatrix(0, 15) = "Invoice Type"
MSHFlexGrid1.TextMatrix(0, 16) = "SM ID"
MSHFlexGrid1.TextMatrix(0, 17) = "Reverse Flag"
MSHFlexGrid1.TextMatrix(0, 18) = "Updated By"
MSHFlexGrid1.ColWidth(0) = 1200
MSHFlexGrid1.ColWidth(1) = 1000
MSHFlexGrid1.ColWidth(2) = 1200
MSHFlexGrid1.ColWidth(3) = 1400
MSHFlexGrid1.ColWidth(4) = 1800
MSHFlexGrid1.ColWidth(5) = 1800
MSHFlexGrid1.ColWidth(6) = 1600
MSHFlexGrid1.ColWidth(7) = 1400
MSHFlexGrid1.ColWidth(8) = 1800
MSHFlexGrid1.ColWidth(9) = 1800
MSHFlexGrid1.ColWidth(10) = 3400
MSHFlexGrid1.ColWidth(11) = 1400
MSHFlexGrid1.ColWidth(12) = 1400
MSHFlexGrid1.ColWidth(13) = 1400
MSHFlexGrid1.ColWidth(14) = 1400

MSHFlexGrid1.WordWrap = True
MSHFlexGrid1.RowHeight(0) = 600

End Sub

Private Sub CmbSLocation_GotFocus()
tmp = CmbSLocation.Text
'Call TableMst_Location(" Where LocationId > 0 ")

If Gen_WcLocation = "" Then
   Call TableMst_Location("")
Else
   Call TableMst_Location("where " & Gen_WcLocation)
End If

CmbSLocation.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No record found !!!! "
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


Private Sub Indata()

Call GButtonLock(Me, True)

Frame1.Enabled = False
Frame2.Enabled = False

TxtTIID = RsTxn_TIDetails!TIID

CmbExchangeTypeID.Visible = False
TxtExchangeTypeID.Visible = True


mExchangeTypeID = IIf(IsNull(RsTxn_TIDetails!ExchangeTypeID), 0, RsTxn_TIDetails!ExchangeTypeID)
If mExchangeTypeID <> 0 Then
   TxtExchangeTypeID.Text = GetExchangeName(mExchangeTypeID)
Else
   TxtExchangeTypeID.Text = ""
End If
CmbExchangeTypeID.Text = TxtExchangeTypeID


CmbDPName.Visible = False
CmbCNonNcdexClient.Visible = False
mClientNAVCode = RsTxn_TIDetails!NavCode

If IsNull(RsTxn_TIDetails!ClientID) = True And IsNull(RsTxn_TIDetails!DepositoryID) = True Then
   mClientIDRow = 0
   mDPID = ""
   tmp = "Select ClientID From Mst_Client Where (NAVCode = '" & mClientNAVCode & "' Or RNAVCode = '" & mClientNAVCode & "')"
   Call TmpTable
   If TmpRs.EOF = False Then
      mClientIDRow = TmpRs!ClientID
   Else
      tmp = "Select DepositoryID From Mst_Depository Where NavCode = '" & mClientNAVCode & "'"
      Call TmpTable
      If TmpRs.EOF = False Then
         mDPID = TmpRs!DepositoryID
      End If
   End If
Else
   mClientIDRow = IIf(IsNull(RsTxn_TIDetails!ClientID), 0, RsTxn_TIDetails!ClientID)
   mDPID = IIf(IsNull(RsTxn_TIDetails!DepositoryID), "", RsTxn_TIDetails!DepositoryID)
End If

If mClientIDRow <> 0 Then 'IsNull(RsTxn_TIDetails!ClientIDRow) = False Then
   'mClientIDRow = RsTxn_TIDetails!ClientIDRow
   Check3.Value = 0
   ans = GetClientNameByRow(CDbl(mClientID))
   TxtClientID = ans(0)
   TxtClientName = ans(1)  'TmpRs!ClientName
   TxtDPID = ans(2)  'TmpRs!DepositoryID
   TxtDPName = ans(3)  'TmpRs!DepositoryName
   TxtNAVClientName = ans(5)
      
   If mExchangeTypeID = 0 Then
      tmp = "Select ExchangeTypeID From Mst_Client Where ClientID = " & mClientIDRow & ""
      Call TmpTable
      mExchangeTypeID = TmpRs!ExchangeTypeID
      TxtExchangeTypeID.Text = GetExchangeName(mExchangeTypeID)
      CmbExchangeTypeID = TxtExchangeTypeID
   End If
   If mExchangeTypeID = 1 Then
      CmbCNonNcdexClient.Visible = True
      CmbCNonNcdexClient.Text = ans(1)
      CmbCNonNcdexClient.Width = 10800
      CmbCNonNcdexClient.Top = TxtClientID.Top
      CmbCNonNcdexClient.Left = TxtClientID.Left
      CmbCNonNcdexClient.Visible = True
      Label25.Visible = False
      Label27.Visible = False
      Label24.Visible = False
   Else
      Label25.Visible = True
      Label27.Visible = True
      Label24.Visible = True
      Check3.Enabled = True
   End If
ElseIf mDPID <> "" Then
   If mExchangeTypeID = 0 Then
      mExchangeTypeID = 2
      TxtExchangeTypeID.Text = GetExchangeName(mExchangeTypeID)
   End If
   'mDPID = RsTxn_TIDetails!DepositoryID
   Check3.Value = 1
   CmbDPName.Visible = True
   CmbDPName.Text = GetDepositorName(mDPID, "N") & "~" & mDPID
   TxtNAVClientName = GetDepositorName(mDPID, "NN")
End If
'mDPID = Null

TxtTINo = RsTxn_TIDetails!TINo
TxtTIDate = RsTxn_TIDetails!TIDate
TxtPostInvoiceNo = RsTxn_TIDetails!PostInvoiceNo
TxtEntryNo = RsTxn_TIDetails!EntryNo

TxtTIAmount = IIf(IsNull(RsTxn_TIDetails!TIAmount), "0.00", Format(RsTxn_TIDetails!TIAmount, "#######0.00"))
TxtServiceTaxAmt = IIf(IsNull(RsTxn_TIDetails!TISTAmount), 0, RsTxn_TIDetails!TISTAmount)

TxtInvType = IIf(IsNull(RsTxn_TIDetails!PIFlag), "", RsTxn_TIDetails!PIFlag)
TxtSMID = IIf(IsNull(RsTxn_TIDetails!PIID), "", RsTxn_TIDetails!PIID)
If TxtInvType = "" Then
   TxtInvType = IIf(IsNull(RsTxn_TIDetails!EntryFlag), "", RsTxn_TIDetails!EntryFlag)
End If

TxtReverseFlag = IIf(IsNull(RsTxn_TIDetails!ReverseFlag), "No", IIf(RsTxn_TIDetails!EntryFlag = "N", "No", "Yes"))

TxtTIAmount.Locked = True
TxtServiceTaxAmt.Locked = True

CmbLocationID.Visible = False
TxtLocationID.Visible = True
CmdLocationID.Visible = False

TxtServiceTaxID.Visible = True
CmbServiceTaxID.Visible = False
CmdServiceTaxID.Visible = False

TxtNAVCode.Visible = True
CmbNAVCode.Visible = False
TxtNAVCode = IIf(IsNull(RsTxn_TIDetails!NavCode), "", RsTxn_TIDetails!NavCode)
CmbNAVCode = TxtNAVCode

mServiceTaxID = IIf(IsNull(RsTxn_TIDetails!ServiceTaxID), "", RsTxn_TIDetails!ServiceTaxID)

TxtCRemark = IIf(IsNull(RsTxn_TIDetails!CollectionRemark), "", RsTxn_TIDetails!CollectionRemark)
TxtCUpdatedBy = IIf(IsNull(RsTxn_TIDetails!CRemarkUpdatedBy), "", RsTxn_TIDetails!CRemarkUpdatedBy)
TxtCUpdatedDate = IIf(IsNull(RsTxn_TIDetails!CRemarkUpdatedDate), "", RsTxn_TIDetails!CRemarkUpdatedDate)

If mServiceTaxID <> "" Then
   TxtServiceTaxID = GetServiceTaxDescription(mServiceTaxID)
   mServiceTaxNAVCode = GetServiceTaxNAVCode(mServiceTaxID)
Else
   TxtServiceTaxID = "" 'GetServiceTaxDescription(mServiceTaxID)
   mServiceTaxNAVCode = "" 'GetServiceTaxNAVCode(mServiceTaxID)"
End If

CmbServiceTaxID.Text = TxtServiceTaxID

mLocationID = IIf(IsNull(RsTxn_TIDetails!LocationID), 0, RsTxn_TIDetails!LocationID)
If mLocationID <> 0 Then
   TxtLocationID = GetLocationName(mLocationID)
   CmbLocationID.Text = TxtLocationID
Else
   TxtLocationID = ""
   CmbLocationID.Text = ""
End If

TxtAdjAmount.Locked = True
TxtAdjAmount = IIf(IsNull(RsTxn_TIDetails!AdjAmountFAG), 0, RsTxn_TIDetails!AdjAmountFAG)

TxtRemarks.Locked = True
TxtRemarks = IIf(IsNull(RsTxn_TIDetails!Remarks), "", RsTxn_TIDetails!Remarks)

'tmp = "Select Sum(RecAmount)as RecAmount, Sum(RecTDS) as RecTDSAmount , Sum(RecOtherDed) as RecOtherDed "
'tmp = tmp & " From Txn_NewCashReceiptTIDetail Where TIID= " & TxtTIID & " "

tmp = "Select TI.TIID,TI.TINo,TI.TIDate,TI.TIAmount,TI.PostInvoiceNo,TI.EntryNo,TI.NAVCode,TI.TISTAmount ,"
tmp = tmp & " (Select isnull(Sum(TCRT.RecAmount),0) as RecAmount From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.NewCRID= TCRT.NewCRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID) RecAmount,"
tmp = tmp & " (Select isnull(Sum(TCRT.RecTDS),0) as RecTDS From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.NewCRID= TCRT.NewCRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID ) + (Select isnull(Sum(TCRT.RecTDS),0) as RecTDS From Txn_CRTDSAdjustmentDetail TCRT Inner Join Txn_CRTDSAdjustment TCR on TCR.CRTDSAdjustmentID= TCRT.CRTDSAdjustmentID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.Flag<>'R'And TIID=TI.TIID )RecTDSAmount,"
tmp = tmp & " (Select isnull(Sum(TCRT.RecOtherDed),0) as RecOtherDed From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.NewCRID= TCRT.NewCRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID )+ (Select isnull(Sum(TCRT.RecOtherDed),0) as RecOtherDed From Txn_CRTDSAdjustmentDetail TCRT Inner Join Txn_CRTDSAdjustment TCR on TCR.CRTDSAdjustmentID= TCRT.CRTDSAdjustmentID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.Flag<>'R'And TIID=TI.TIID  )RecOtherDed"
tmp = tmp & " From  Txn_TIDetails TI"
tmp = tmp & " Where TI.TIID= " & TxtTIID & " "

Call TmpTable

If TmpRs.RecordCount > 0 Then

   TxtRecAmount = IIf(IsNull(TmpRs!RecAmount), 0, Format(TmpRs!RecAmount, "#######0.00"))
   TxtRecTDS = IIf(IsNull(TmpRs!RecTDSAmount), 0, Format(TmpRs!RecTDSAmount, "#######0.00"))
   TxtRecOtherDeduction = IIf(IsNull(TmpRs!RecOtherDed), 0, Format(TmpRs!RecOtherDed, "#######0.00"))
Else

   TxtRecAmount = Format(0, "#######0.00")  'IIf(IsNull(TmpRs!RecAmount), 0, TmpRs!RecAmount)
   TxtRecTDS = Format(0, "#######0.00") 'IIf(IsNull(TmpRs!RecTDSAmount), 0, TmpRs!RecTDSAmount)
   TxtRecOtherDeduction = Format(0, "#######0.00") 'IIf(IsNull(TmpRs!RecOtherDed), 0, TmpRs!RecOtherDed)

End If

mEntryFlag = IIf(IsNull(RsTxn_TIDetails!EntryFlag), "System", RsTxn_TIDetails!EntryFlag)

TxtCreated = IIf(IsNull(RsTxn_TIDetails!CreatedBy), "", RsTxn_TIDetails!CreatedBy) & " :- " & IIf(IsNull(RsTxn_TIDetails!CreatedDate), "", RsTxn_TIDetails!CreatedDate)
TxtUpdated = IIf(IsNull(RsTxn_TIDetails!UpdatedBy), "", RsTxn_TIDetails!UpdatedBy) & " :- " & IIf(IsNull(RsTxn_TIDetails!UpdatedDate), "", RsTxn_TIDetails!UpdatedDate)


CmdViewInvoice.Enabled = True

End Sub

Private Sub CmbLocationID_GotFocus()

tmp = CmbLocationID.Text

If Gen_WcLocation = "" Then
   Call TableMst_Location("")
Else
   Call TableMst_Location("where " & Gen_WcLocation)
End If

CmbLocationID.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No record found !!!! "
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

mLocationID = RsMst_Location!LocationID

End Sub

Private Sub CmbServiceTaxID_GotFocus()
tmp = CmbServiceTaxID.Text
Call TableMst_ServiceTax
CmbServiceTaxID.Clear
If RsMst_ServiceTax.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_ServiceTax.RecordCount
    CmbServiceTaxID.AddItem RsMst_ServiceTax!Description
    RsMst_ServiceTax.MoveNext
Next
CmbServiceTaxID.Text = tmp
End Sub
Private Sub CmbServiceTaxID_LostFocus()
If CmbServiceTaxID.Text = "" Then
   mServiceTaxID = Null
   Exit Sub
End If
RsMst_ServiceTax.MoveFirst
RsMst_ServiceTax.Find "Description = '" & CmbServiceTaxID.Text & "'"
If RsMst_ServiceTax.EOF Then
   MsgBox "Invalid selection "
   CmbServiceTaxID.SetFocus
   Exit Sub
End If
mServiceTaxID = RsMst_ServiceTax!ServiceTaxID
mServiceTaxNAVCode = IIf(IsNull(RsMst_ServiceTax!NavCode), "", RsMst_ServiceTax!NavCode)

End Sub
Private Sub CmbSReverseFlag_LostFocus()
If CmbSReverseFlag.Text = "" Then Exit Sub
If LCase(CmbSReverseFlag.Text) <> "yes" And LCase(CmbSReverseFlag.Text) <> "no" Then
   MsgBox "Invalid Selection!!!"
   CmbSReverseFlag.SetFocus
   Exit Sub
End If
End Sub

