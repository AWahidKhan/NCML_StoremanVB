VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmTxn_SecurityBilling 
   Caption         =   "Security Invoice Updation and Payment"
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
      TabIndex        =   42
      Top             =   0
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      TabIndex        =   41
      Top             =   10440
      Width           =   10935
   End
   Begin VB.Frame Frame4 
      BorderStyle     =   0  'None
      Caption         =   "s"
      Height          =   10620
      Left            =   0
      TabIndex        =   29
      Top             =   -120
      Width           =   15180
      Begin VB.Frame Frame0 
         Height          =   1695
         Left            =   120
         TabIndex        =   36
         Top             =   9660
         Visible         =   0   'False
         Width           =   14850
         Begin VB.ComboBox CmbSStatus 
            Height          =   315
            ItemData        =   "FrmTxn_SecurityBilling.frx":0000
            Left            =   12480
            List            =   "FrmTxn_SecurityBilling.frx":0010
            Sorted          =   -1  'True
            TabIndex        =   82
            TabStop         =   0   'False
            Top             =   330
            Width           =   1755
         End
         Begin VB.ComboBox CmbSForwardTo 
            Height          =   315
            Left            =   8925
            Sorted          =   -1  'True
            TabIndex        =   75
            TabStop         =   0   'False
            Top             =   345
            Width           =   3435
         End
         Begin VB.ComboBox CmbSSecurityAgency 
            Height          =   315
            Left            =   4860
            Sorted          =   -1  'True
            TabIndex        =   74
            TabStop         =   0   'False
            Top             =   330
            Width           =   3990
         End
         Begin VB.OptionButton OptTransactionDate 
            Caption         =   "Transaction Date"
            Height          =   255
            Left            =   120
            TabIndex        =   73
            Top             =   353
            Width           =   1800
         End
         Begin VB.OptionButton OptMonthYear 
            Caption         =   "Month and Year"
            Height          =   255
            Left            =   120
            TabIndex        =   72
            Top             =   120
            Width           =   1680
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
            Left            =   1245
            MaxLength       =   100
            TabIndex        =   70
            Top             =   840
            Width           =   3495
         End
         Begin VB.CheckBox ChkBatchNo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Check1"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   7320
            TabIndex        =   68
            Top             =   885
            Width           =   225
         End
         Begin VB.CommandButton CmdGo 
            Caption         =   "Go"
            Height          =   350
            Left            =   12165
            TabIndex        =   27
            Top             =   690
            Width           =   1000
         End
         Begin VB.CommandButton CmdExcel 
            Caption         =   "Excel"
            Height          =   350
            Left            =   13170
            TabIndex        =   28
            Top             =   690
            Width           =   1000
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   1455
            Left            =   75
            TabIndex        =   37
            Top             =   1320
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
            Left            =   3465
            TabIndex        =   76
            TabStop         =   0   'False
            Top             =   345
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
            Format          =   50987009
            CurrentDate     =   36494
         End
         Begin MSComCtl2.DTPicker SFrom 
            Height          =   300
            Left            =   2055
            TabIndex        =   77
            TabStop         =   0   'False
            Top             =   345
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
            Format          =   50987009
            CurrentDate     =   36494
         End
         Begin VB.Label Label19 
            Caption         =   "Status"
            Height          =   225
            Left            =   12930
            TabIndex        =   83
            Top             =   120
            Width           =   855
         End
         Begin VB.Label Label31 
            Caption         =   "Forwarded"
            Height          =   225
            Left            =   10215
            TabIndex        =   81
            Top             =   135
            Width           =   855
         End
         Begin VB.Label Label45 
            Caption         =   "To Date"
            Height          =   255
            Left            =   3705
            TabIndex        =   80
            Top             =   120
            Width           =   855
         End
         Begin VB.Label Label37 
            Caption         =   "Security Agency"
            Height          =   225
            Left            =   6248
            TabIndex        =   79
            Top             =   135
            Width           =   1215
         End
         Begin VB.Label Label44 
            Caption         =   "From Date"
            Height          =   255
            Left            =   2295
            TabIndex        =   78
            Top             =   120
            Width           =   855
         End
         Begin VB.Label LblBillNo 
            Caption         =   "Bill No"
            Height          =   255
            Left            =   240
            TabIndex        =   71
            Top             =   885
            Width           =   885
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
            Left            =   5520
            TabIndex        =   69
            Top             =   885
            Width           =   1710
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
            Left            =   11085
            TabIndex        =   39
            Top             =   750
            Width           =   90
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
            Left            =   8565
            TabIndex        =   38
            Top             =   750
            Width           =   2415
         End
      End
      Begin VB.Frame Frame1 
         Height          =   3480
         Left            =   120
         TabIndex        =   30
         Top             =   150
         Width           =   14880
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
            Left            =   5760
            Sorted          =   -1  'True
            TabIndex        =   15
            Top             =   2535
            Width           =   3135
         End
         Begin VB.ComboBox CmbStateIndiaName 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   11385
            Sorted          =   -1  'True
            TabIndex        =   3
            Top             =   217
            Width           =   2295
         End
         Begin VB.TextBox TxtFAGRemarks 
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
            Left            =   1245
            Locked          =   -1  'True
            MaxLength       =   1000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   17
            Top             =   2970
            Width           =   12555
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
            Height          =   435
            Left            =   1245
            MaxLength       =   1000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   13
            Top             =   2040
            Width           =   12555
         End
         Begin VB.ComboBox CmbSecurityAgency 
            BeginProperty Font 
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
            TabIndex        =   4
            Top             =   675
            Width           =   8355
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
            ItemData        =   "FrmTxn_SecurityBilling.frx":003A
            Left            =   1245
            List            =   "FrmTxn_SecurityBilling.frx":004A
            Sorted          =   -1  'True
            TabIndex        =   14
            Top             =   2535
            Width           =   1650
         End
         Begin VB.TextBox TxtBatchNo 
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
            Left            =   11385
            Locked          =   -1  'True
            TabIndex        =   57
            TabStop         =   0   'False
            Top             =   668
            Width           =   2415
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
            Left            =   11370
            Sorted          =   -1  'True
            TabIndex        =   12
            Top             =   1590
            Width           =   2430
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
            Left            =   11400
            Locked          =   -1  'True
            TabIndex        =   56
            TabStop         =   0   'False
            Top             =   1590
            Visible         =   0   'False
            Width           =   2400
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
            Left            =   13860
            TabIndex        =   55
            TabStop         =   0   'False
            Top             =   1620
            Width           =   375
         End
         Begin VB.CommandButton CmdSecurityAgency 
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
            Left            =   9580
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   690
            Width           =   375
         End
         Begin VB.TextBox TxtTotBillAmt 
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
            Left            =   11385
            Locked          =   -1  'True
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   1155
            Width           =   2415
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
            TabIndex        =   11
            Top             =   1586
            Width           =   8325
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
            Left            =   9580
            TabIndex        =   52
            TabStop         =   0   'False
            Top             =   1620
            Width           =   375
         End
         Begin VB.TextBox TxtSTax 
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
            Left            =   8655
            TabIndex        =   9
            Top             =   1155
            Width           =   1335
         End
         Begin VB.TextBox TxtBillAmt 
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
            Left            =   6435
            TabIndex        =   8
            Top             =   1140
            Width           =   1455
         End
         Begin VB.TextBox TxtBillNo 
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
            Left            =   1245
            MaxLength       =   100
            TabIndex        =   6
            Top             =   1125
            Width           =   1455
         End
         Begin VB.TextBox TxtSecurityAgency 
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
            TabIndex        =   45
            TabStop         =   0   'False
            Top             =   675
            Width           =   8715
         End
         Begin VB.ComboBox CmbYear 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   8655
            Sorted          =   -1  'True
            TabIndex        =   2
            Top             =   217
            Width           =   1380
         End
         Begin VB.ComboBox CmbMonth 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   6435
            TabIndex        =   1
            Top             =   217
            Width           =   1455
         End
         Begin VB.TextBox TxtTxnID 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   375
            Left            =   1245
            Locked          =   -1  'True
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   210
            Width           =   1455
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
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   240
            Visible         =   0   'False
            Width           =   1440
         End
         Begin MSComCtl2.DTPicker TxtTxnDate 
            Height          =   375
            Left            =   3900
            TabIndex        =   33
            Top             =   210
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   50987009
            CurrentDate     =   39884
         End
         Begin VB.TextBox TxtYear 
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
            Left            =   8655
            Locked          =   -1  'True
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   210
            Width           =   1380
         End
         Begin VB.TextBox TxtMonth 
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
            Left            =   6435
            Locked          =   -1  'True
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   217
            Width           =   1455
         End
         Begin MSComCtl2.DTPicker TxtBillDate 
            Height          =   375
            Left            =   3900
            TabIndex        =   7
            Top             =   1139
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Format          =   50987009
            CurrentDate     =   39884
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
            Left            =   1245
            Locked          =   -1  'True
            TabIndex        =   53
            TabStop         =   0   'False
            Top             =   1590
            Visible         =   0   'False
            Width           =   8715
         End
         Begin VB.TextBox TXtStatus 
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
            Left            =   1245
            Locked          =   -1  'True
            TabIndex        =   59
            TabStop         =   0   'False
            Top             =   2535
            Visible         =   0   'False
            Width           =   1650
         End
         Begin MSComCtl2.DTPicker TxtApprovalDate 
            Height          =   375
            Left            =   11370
            TabIndex        =   16
            Top             =   2535
            Visible         =   0   'False
            Width           =   2430
            _ExtentX        =   4286
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
            Format          =   50987009
            CurrentDate     =   39884
         End
         Begin VB.TextBox TxtStateIndiaName 
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
            Left            =   11385
            Locked          =   -1  'True
            TabIndex        =   135
            TabStop         =   0   'False
            Top             =   217
            Width           =   2415
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
            Left            =   5760
            Locked          =   -1  'True
            TabIndex        =   137
            TabStop         =   0   'False
            Top             =   2535
            Width           =   3135
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
            Left            =   3000
            TabIndex        =   136
            Top             =   2595
            Width           =   2685
         End
         Begin VB.Label Label28 
            Caption         =   "State India"
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
            Left            =   10140
            TabIndex        =   134
            Top             =   270
            Width           =   1245
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
            Left            =   9120
            TabIndex        =   67
            Top             =   2580
            Visible         =   0   'False
            Width           =   1980
         End
         Begin VB.Label Label16 
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
            Height          =   255
            Left            =   120
            TabIndex        =   66
            Top             =   2130
            Width           =   990
         End
         Begin VB.Label Label4 
            Caption         =   "Security Agency"
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
            TabIndex        =   65
            Top             =   622
            Width           =   1140
         End
         Begin VB.Label LbClaimID 
            Caption         =   "Txn ID"
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
            TabIndex        =   64
            Top             =   247
            Width           =   900
         End
         Begin VB.Label Label6 
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
            Height          =   255
            Left            =   120
            TabIndex        =   63
            Top             =   1199
            Width           =   885
         End
         Begin VB.Label Label7 
            Caption         =   "Forwarded"
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
            TabIndex        =   62
            Top             =   1665
            Width           =   1125
         End
         Begin VB.Label Label8 
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
            Height          =   465
            Left            =   120
            TabIndex        =   61
            Top             =   2955
            Width           =   990
         End
         Begin VB.Label Label11 
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
            Left            =   120
            TabIndex        =   60
            Top             =   2603
            Width           =   810
         End
         Begin VB.Label Label10 
            Caption         =   "Batch No"
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
            Left            =   10140
            TabIndex        =   58
            Top             =   735
            Width           =   855
         End
         Begin VB.Label Label9 
            Caption         =   "S.Tax Desc"
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
            Left            =   10080
            TabIndex        =   54
            Top             =   1665
            Width           =   1230
         End
         Begin VB.Label Label5 
            Caption         =   "Total Bill Amount"
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
            Left            =   10140
            TabIndex        =   51
            Top             =   1110
            Width           =   810
         End
         Begin VB.Label Label3 
            Caption         =   "S.Tax"
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
            Left            =   7995
            TabIndex        =   50
            Top             =   1230
            Width           =   705
         End
         Begin VB.Label Label2 
            Caption         =   "Bill Amount"
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
            Left            =   5535
            TabIndex        =   49
            Top             =   1095
            Width           =   810
         End
         Begin VB.Label Label1 
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
            Height          =   270
            Left            =   2955
            TabIndex        =   48
            Top             =   1191
            Width           =   900
         End
         Begin VB.Label Label14 
            Caption         =   "Month"
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
            Left            =   5535
            TabIndex        =   44
            Top             =   255
            Width           =   645
         End
         Begin VB.Label Label13 
            Caption         =   "Year"
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
            Left            =   7995
            TabIndex        =   43
            Top             =   255
            Width           =   555
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
            Left            =   13500
            TabIndex        =   35
            Top             =   45
            Visible         =   0   'False
            Width           =   630
         End
         Begin VB.Label Label18 
            Caption         =   "Txn Date"
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
            Left            =   2985
            TabIndex        =   34
            Top             =   262
            Width           =   900
         End
      End
      Begin VB.Frame ButtonFrm 
         Height          =   1080
         Left            =   120
         TabIndex        =   40
         Top             =   8745
         Width           =   14865
         Begin VB.CommandButton CmdPrint 
            Caption         =   "Print"
            Height          =   810
            Left            =   13005
            TabIndex        =   84
            Top             =   195
            Width           =   1740
         End
         Begin VB.CommandButton ExitCmd 
            Caption         =   "E&xit"
            Height          =   400
            Left            =   10425
            TabIndex        =   22
            Top             =   600
            Width           =   2580
         End
         Begin VB.CommandButton SearchCmd 
            Caption         =   "Search"
            Height          =   400
            Left            =   10425
            TabIndex        =   21
            Top             =   195
            Width           =   2580
         End
         Begin VB.CommandButton LastCmd 
            Caption         =   "&Last"
            Height          =   400
            Left            =   7845
            TabIndex        =   23
            Top             =   600
            Width           =   2580
         End
         Begin VB.CommandButton DeleteCmd 
            Caption         =   "&Delete"
            Height          =   400
            Left            =   7845
            TabIndex        =   20
            Top             =   195
            Width           =   2580
         End
         Begin VB.CommandButton FirstCmd 
            Caption         =   "&First"
            Height          =   400
            Left            =   5265
            TabIndex        =   24
            Top             =   600
            Width           =   2580
         End
         Begin VB.CommandButton EditCmd 
            Caption         =   "&Edit"
            Height          =   400
            Left            =   5265
            TabIndex        =   19
            Top             =   195
            Width           =   2580
         End
         Begin VB.CommandButton PreviousCmd 
            Caption         =   "&Previous"
            Height          =   400
            Left            =   2685
            TabIndex        =   25
            Top             =   600
            Width           =   2580
         End
         Begin VB.CommandButton SaveCmd 
            Caption         =   "Save"
            Height          =   400
            Left            =   2685
            TabIndex        =   18
            Top             =   195
            Width           =   2580
         End
         Begin VB.CommandButton NextCmd 
            Caption         =   "&Next"
            Height          =   400
            Left            =   105
            TabIndex        =   26
            Top             =   600
            Width           =   2580
         End
         Begin VB.CommandButton AddCmd 
            Caption         =   "&Add New"
            Height          =   400
            Left            =   120
            TabIndex        =   0
            Top             =   195
            Width           =   2580
         End
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   5055
         Left            =   120
         TabIndex        =   85
         Top             =   3720
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   8916
         _Version        =   393216
         TabHeight       =   520
         TabCaption(0)   =   "Security Billing"
         TabPicture(0)   =   "FrmTxn_SecurityBilling.frx":0074
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "MSHFlexGrid2"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "ChkCheckAll"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "TxtColAmt"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "Payment Details"
         TabPicture(1)   =   "FrmTxn_SecurityBilling.frx":0090
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame3"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Log Details"
         TabPicture(2)   =   "FrmTxn_SecurityBilling.frx":00AC
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "TxtUpdated"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "TxtCreated"
         Tab(2).Control(1).Enabled=   0   'False
         Tab(2).Control(2)=   "Label17"
         Tab(2).Control(3)=   "Label15"
         Tab(2).ControlCount=   4
         Begin VB.Frame Frame3 
            Height          =   4635
            Left            =   -74925
            TabIndex        =   90
            Top             =   375
            Width           =   14535
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
               TabIndex        =   100
               Top             =   1320
               Width           =   14355
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
                  TabIndex        =   113
                  TabStop         =   0   'False
                  Top             =   2895
                  Width           =   1935
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
                  TabIndex        =   112
                  TabStop         =   0   'False
                  Top             =   2475
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
                  TabIndex        =   111
                  TabStop         =   0   'False
                  Top             =   2895
                  Width           =   2055
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
                  TabIndex        =   110
                  TabStop         =   0   'False
                  Top             =   2895
                  Width           =   4800
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
                  TabIndex        =   109
                  TabStop         =   0   'False
                  Top             =   2475
                  Width           =   5175
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
                  TabIndex        =   108
                  TabStop         =   0   'False
                  Top             =   2070
                  Width           =   1785
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
                  TabIndex        =   107
                  TabStop         =   0   'False
                  Top             =   2070
                  Width           =   1860
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
                  TabIndex        =   106
                  TabStop         =   0   'False
                  Top             =   2070
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
                  TabIndex        =   103
                  Top             =   270
                  Width           =   1785
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
                  TabIndex        =   102
                  Top             =   1155
                  Width           =   4800
               End
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
                  TabIndex        =   101
                  Top             =   1155
                  Width           =   5160
               End
               Begin MSComCtl2.DTPicker TxtFagDisbDate 
                  Height          =   375
                  Left            =   7125
                  TabIndex        =   114
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
                  Format          =   50987009
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
                  ItemData        =   "FrmTxn_SecurityBilling.frx":00C8
                  Left            =   2820
                  List            =   "FrmTxn_SecurityBilling.frx":00D8
                  TabIndex        =   105
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
                  TabIndex        =   104
                  Top             =   690
                  Width           =   6855
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
                  TabIndex        =   116
                  Top             =   690
                  Width           =   6855
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
                  TabIndex        =   115
                  Top             =   690
                  Width           =   1800
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
                  TabIndex        =   130
                  Top             =   2940
                  Width           =   900
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
                  TabIndex        =   129
                  Top             =   2955
                  Width           =   900
               End
               Begin VB.Label Label24 
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
                  TabIndex        =   128
                  Top             =   2940
                  Width           =   2445
               End
               Begin VB.Label Label25 
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
                  TabIndex        =   127
                  Top             =   2520
                  Width           =   1140
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
                  TabIndex        =   126
                  Top             =   2565
                  Width           =   1140
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
                  TabIndex        =   125
                  Top             =   2130
                  Width           =   1290
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
                  TabIndex        =   124
                  Top             =   2130
                  Width           =   765
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
                  TabIndex        =   122
                  Top             =   750
                  Width           =   1575
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
                  TabIndex        =   121
                  Top             =   750
                  Width           =   1305
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
                  TabIndex        =   120
                  Top             =   330
                  Width           =   1725
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
                  TabIndex        =   119
                  Top             =   330
                  Width           =   2370
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
                  TabIndex        =   118
                  Top             =   1515
                  Width           =   2010
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
                  TabIndex        =   117
                  Top             =   1350
                  Width           =   1530
               End
            End
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
               TabIndex        =   91
               Top             =   105
               Width           =   14385
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
                  TabIndex        =   95
                  Top             =   300
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
                  TabIndex        =   94
                  Top             =   300
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
                  TabIndex        =   93
                  Top             =   705
                  Width           =   5625
               End
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
                  TabIndex        =   92
                  Top             =   705
                  Width           =   5145
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
                  TabIndex        =   99
                  Top             =   360
                  Width           =   1065
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
                  TabIndex        =   98
                  Top             =   360
                  Width           =   1230
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
                  TabIndex        =   97
                  Top             =   765
                  Width           =   645
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
                  TabIndex        =   96
                  Top             =   705
                  Width           =   975
               End
            End
         End
         Begin VB.TextBox TxtColAmt 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
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
            ForeColor       =   &H000000FF&
            Height          =   360
            Left            =   1440
            ScrollBars      =   2  'Vertical
            TabIndex        =   89
            Top             =   360
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.CheckBox ChkCheckAll 
            Appearance      =   0  'Flat
            Caption         =   "Select All"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   88
            Top             =   360
            Width           =   1335
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
            Left            =   -73725
            Locked          =   -1  'True
            TabIndex        =   87
            TabStop         =   0   'False
            Top             =   615
            Width           =   4605
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
            Left            =   -73725
            Locked          =   -1  'True
            TabIndex        =   86
            TabStop         =   0   'False
            Top             =   1200
            Width           =   4605
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
            Height          =   4110
            Left            =   120
            TabIndex        =   131
            Top             =   840
            Width           =   14490
            _ExtentX        =   25559
            _ExtentY        =   7250
            _Version        =   393216
            FixedCols       =   0
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
         Begin VB.Label Label15 
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
            TabIndex        =   133
            Top             =   1260
            Width           =   915
         End
         Begin VB.Label Label17 
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
            TabIndex        =   132
            Top             =   660
            Width           =   945
         End
      End
   End
End
Attribute VB_Name = "FrmTxn_SecurityBilling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg, mForwardToID, mMonth, intMonthDays, mMonthStartDt, mMonthEndDate, mSecurityAgencyID, mServiceTaxID   As Variant
Dim mSelRow, mSelCol                    As Variant
Dim mSForwardToID, mSSecurityAgencyID    As Variant
Const strChecked = "þ"
Const strUnChecked = "q"
Dim mTotalDays As Variant
Dim mFromDate, mToDate, mStatus As Variant
Dim FileName As String
Dim mFAGBankID, mBatchID As Double, mPBatchID As Double

''-- Ullhas Patil 04 Oct 2017
Dim mStateIndiaID As Double
Dim mBillingStateID As Double


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

''-- Ullhas Patil 04 Oct 2017
Private Sub CmbStateIndiaName_GotFocus()
tmp = CmbStateIndiaName.Text
Call TableMst_StateIndia
CmbStateIndiaName.Clear
If RsMst_StateIndia.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_StateIndia.RecordCount
    CmbStateIndiaName.AddItem RsMst_StateIndia!StateIndiaName
    RsMst_StateIndia.MoveNext
Next
CmbStateIndiaName.Text = tmp
End Sub

Private Sub CmbStateIndiaName_LostFocus()
If CmbStateIndiaName.Text = "" Then
   Exit Sub
End If
RsMst_StateIndia.MoveFirst
RsMst_StateIndia.Find "StateIndiaName = '" & CmbStateIndiaName.Text & "'"
If RsMst_StateIndia.EOF Then
   MsgBox "Invalid selection "
   CmbStateIndiaName.SetFocus
   Exit Sub
End If
mStateIndiaID = RsMst_StateIndia!StateIndiaID
End Sub
''-------------------------------------

Private Sub TxtFAGDisbAmt_LostFocus()

If TxtFAGDisbAmt = "" Then Exit Sub

If IsNumeric(TxtFAGDisbAmt) = False Or Val(TxtFAGDisbAmt) < 0 Then
   MsgBox "Invalid format!!!"
   TxtFAGDisbAmt.SetFocus
   Exit Sub
End If

If Val(TxtFAGDisbAmt) > Val(TxtTotBillAmt) Then
   MsgBox "Disbusment amount should be less than or equal to total bill amount!!!"
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







Private Sub ClearCtrl()

TxtBillNo.Text = ""
TxtBillDate.Value = Date
TxtBillAmt.Text = ""
TxtSTax.Text = ""
TxtTotBillAmt.Text = ""

CmbForwardTo.Text = ""
CmbForwardTo.Visible = True
CmdForwardTo.Visible = True
TxtForwardTo.Visible = False

CmbSTaxDesc.Text = ""
CmbSTaxDesc.Visible = True
CmdSTaxDesc.Visible = True
TxtSTaxDesc.Visible = False

TxtFAGRemarks.Text = ""

CmbStatus.Text = "Pending"
CmbStatus.Visible = True
TxtStatus.Visible = False

TxtCreated.Text = ""
TxtUpdated.Text = ""

Call Grid_Head
End Sub

Private Sub AddCmd_Click()
Edit_Flg = "A"
Frame1.Enabled = True
Call GButtonLock(Me, False)

TxtTxnID.Text = ""
TxtTxnDate.Value = Date
TxtMonth.Visible = False
CmbMonth.Visible = True
CmbMonth.Text = ""

TxtYear.Visible = False
CmbYear.Visible = True
CmbYear.Text = ""

TxtBatchNo.Text = ""

CmbSecurityAgency.Text = ""
CmbSecurityAgency.Visible = True
CmdSecurityAgency.Visible = True
TxtSecurityAgency.Visible = False

TxtBillNo.Locked = False
TxtBillDate.Enabled = True
TxtBillAmt.Locked = False
CmbStatus.Enabled = False
TxtSTax.Locked = False
TxtRemarks.Locked = False
TxtRemarks = ""

''--- Ullhas Patil 04 Oct 2017
CmbStateIndiaName.Visible = True: CmbStateIndiaName = ""
TxtStateIndiaName.Visible = False: TxtStateIndiaName = ""
''--- Ullhas Patil 25 Oct 2017
CmbBillingState.Visible = True: CmbBillingState = ""
TxtBillingState.Visible = False: TxtBillingState = ""
''-----------------------------

Call ClearCtrl
If RsTxn_SecurityBilling.RecordCount > 0 Then
   CmbMonth.SetFocus
End If
CmdPrint.Enabled = False
End Sub


Private Sub CmbSForwardTo_GotFocus()
tmp = CmbSForwardTo.Text
Call TableMst_Employee '("Where Flag = 'Y'")
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

Private Sub CmbStatus_LostFocus()

If CmbStatus.Text = "" Then
   TxtFAGRemarks.Text = ""
   TxtFAGRemarks.Locked = True
   TxtApprovalDate.Visible = False
   Label27.Visible = False
   Exit Sub
End If


If LCase(CmbStatus.Text) <> "pending" And LCase(CmbStatus.Text) <> "approved" And LCase(CmbStatus.Text) <> "rejected" And LCase(CmbStatus.Text) <> "settled" Then
   MsgBox "Invalid Selection!!!"
   CmbStatus.SetFocus
   Exit Sub
End If

If LCase(CmbStatus.Text) = "approved" Then
   mStatus = "A"
   TxtFAGRemarks.Locked = False
   TxtApprovalDate.Visible = True
   Label27.Visible = True
   TxtApprovalDate.Enabled = True
   TxtApprovalDate.Value = Date
   TxtFAGRemark = ""
ElseIf LCase(CmbStatus.Text) = "pending" Then
   mStatus = "P"
   TxtFAGRemarks.Locked = True
   TxtApprovalDate.Visible = False
   Label27.Visible = False
   TxtFAGRemark = ""
ElseIf LCase(CmbStatus.Text) = "rejected" Then
   mStatus = "R"
   TxtFAGRemarks.Locked = False
   TxtApprovalDate.Visible = True
   Label27.Visible = True
   TxtApprovalDate.Enabled = True
   TxtApprovalDate.Value = Date
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
   TxtFAGDisbAmt.Text = Val(TxtTotBillAmt)
   TxtDeductionRemark = "TDS@"
   TxtFAGRemark = Trim(TxtFAGRemarks)
   Call TableMst_FAGBank(" Where FAGBankID in (" & Gen_FAGBankIDForPayment & ") ")
   If RsMst_FAGBank.RecordCount > 0 Then
      CmbFAGBank.Text = RsMst_FAGBank!BankName & "~" & RsMst_FAGBank!AccountNo
      mFAGBankID = RsMst_FAGBank!FAGBankID
   End If

Else
   MsgBox "Invalid selection!!!"
   CmbStatus.SetFocus
   TxtFAGRemarks.Locked = True
   TxtApprovalDate.Visible = False
   Label27.Visible = False
   Exit Sub
End If

If LCase(CmbStatus.Text) = "approved" Or LCase(CmbStatus.Text) = "rejected" Or LCase(CmbStatus.Text) = "settled" Then
    If LCase(Gen_UserRole) <> "fag" And LCase(Gen_UserRole) <> "fagadmin" And LCase(Gen_UserRole) <> "power" Then
       MsgBox "Only FAG can approve or reject or settled!!!"
       CmbStatus.Text = "Pending"
       TxtFAGRemarks.Locked = True
       TxtApprovalDate.Visible = False
       Label27.Visible = False
       Exit Sub
    Else
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
    End If
End If
End Sub

Private Sub CmdForwardTo_Click()
FrmMst_Employee.Show
End Sub

Private Sub CmdGo_Click()
Dim wc As Variant
Dim strSplitString() As String
Dim mMonth, mYear As Variant
Dim mToMonth, mToYear As Variant
Dim mTotalDays As Variant


Call Grid_Head1

If SFrom.Value > STo.Value Then
   MsgBox "'From Date' must be less than 'To date'!!!"
   SFrom.SetFocus
   Exit Sub
End If

mMonth = Month(SFrom.Value)
mYear = Year(SFrom.Value)
mToMonth = Month(STo.Value)
mToYear = Year(STo.Value)
mTotalDays = Day(DateSerial(mToYear, mToMonth + 1, 0))


mFromDate = Format(mYear & "-" & mMonth & "-" & "01", "yyyy/mm/dd")
mToDate = Format(mToYear & "-" & mToMonth & "-" & mTotalDays, "yyyy/mm/dd")

wc = ""
' Commented On 10 Aug 2011 12.00 PM
'tmp = "Select TSB.BillTxnID, TSB.BillTxnDate, Month(MonthYear),Year(MonthYear), MSA.AgencyName,"
'tmp = tmp & " BillNo, ME.EmployeeName +'~'+ ME.EmployeeNo as ForwardTo, Status"
'tmp = tmp & " From Txn_SecurityBilling TSB "
'tmp = tmp & " Inner Join Mst_SecurityAgency MSA On MSA.AgencyID = TSB.SecurityAgencyID "
'tmp = tmp & " Inner Join Mst_Employee ME On ME.EmployeeID = TSB.ForwardTo"

tmp = " Select TSB.BillTxnID, TSB.BillTxnDate, datename(Month,MonthYear)'Month',Year(MonthYear)'Year', MSA.AgencyName,"
tmp = tmp & " BillNo,BillDate,BillAmt,BillSTax,TotalAmt ,ME.EmployeeName +'~'+ ME.EmployeeNo as ForwardTo,"
tmp = tmp & " TSB.Remarks, Status, ApprovedDate, FAGRemarks"
tmp = tmp & " ,BatchID,Mst.Description,TSB.BankName,TSB.ChequeNo,TSB.PaidAmount,TSB.TDS,TSB.PaidDate,TSB.PaymentdetailsUpdatedBy,"
tmp = tmp & " TSB.PaymentdetailsUpdatedOn,MSI.StateIndiaName,MSI.StateIndiaID,MSII.StateIndiaName[BillingStateName],TSB.BillingStateIndiaID"
tmp = tmp & " ,MFB.BankName+'~'+MFB.AccountNo 'FAGBankName' ,TSB.DisbuAmount,TSB.DisbuDate,TSB.ModeOfPay,TSB.DeductionRemark From Txn_SecurityBilling TSB"
tmp = tmp & " Inner Join Mst_SecurityAgency MSA On MSA.SecurityAgencyID = TSB.SecurityAgencyID"
tmp = tmp & " Inner Join Mst_Employee ME On ME.EmployeeID = TSB.ForwardTo"
tmp = tmp & " inner join MST_ServiceTax MST On MST.ServiceTAxID=TSB.ServiceTaxID"
tmp = tmp & " Left join Mst_StateIndia MSI On TSB.StateIndiaID = MSI.StateIndiaID "         ''--- Ullhas Patil 04 Oct 2017
tmp = tmp & " Left join Mst_StateIndia MSII On TSB.BillingStateIndiaID = MSII.StateIndiaID "     ''--- Ullhas Patil 25 Oct 2017
tmp = tmp & " Left Join Mst_FAGBank MFB On TSB.FAGBankID = MFB.FAGBankID "
If OptTransactionDate.Value = True Then
   If Gen_DBType = "MDB" Then
      tmp = tmp & " Where BillTxnDate Between #" & Format(SFrom, Gen_DatFormat) & "# And #" & Format(STo, Gen_DatFormat) & "#) "
      wc = " BillTxnDate >= #" & Format(SFrom, Gen_DatFormat) & "# And BillTxnDate <= #" & Format(STo, Gen_DatFormat) & "#) "
   Else
      tmp = tmp & " Where BillTxnDate Between '" & Format(SFrom, Gen_DatFormat) & "' And '" & Format(STo, Gen_DatFormat) & "' "
      wc = " BillTxnDate >= '" & Format(SFrom, Gen_DatFormat) & "' And BillTxnDate <= '" & Format(STo, Gen_DatFormat) & "' "
   End If
 
ElseIf OptMonthYear.Value = True Then
   If Gen_DBType = "MDB" Then
      tmp = tmp & " Where MonthYear Between #" & Format(mFromDate, Gen_DatFormat) & "# And #" & Format(mToDate, Gen_DatFormat) & "#) "
      wc = " MonthYear >= #" & Format(mFromDate, Gen_DatFormat) & "# And MonthYear <= #" & Format(mToDate, Gen_DatFormat) & "#) "
   Else
      tmp = tmp & " Where MonthYear Between '" & Format(mFromDate, Gen_DatFormat) & "' And '" & Format(mToDate, Gen_DatFormat) & "' "
      wc = " MonthYear >= '" & Format(mFromDate, Gen_DatFormat) & "' And MonthYear <= '" & Format(mToDate, Gen_DatFormat) & "' "
   End If
End If

' Commentd On 10082011 12 pm
'If Gen_DBType = "MDB" Then
'
'   tmp = tmp & " Where BillTxnDate Between #" & Format(SFrom, Gen_DatFormat) & "# And #" & Format(STo, Gen_DatFormat) & "#) "
'   wc = " BillTxnDate >= #" & Format(SFrom, Gen_DatFormat) & "# And BillTxnDate <= #" & Format(STo, Gen_DatFormat) & "#) "
'Else
'   tmp = tmp & " Where BillTxnDate Between '" & Format(SFrom, Gen_DatFormat) & "' And '" & Format(STo, Gen_DatFormat) & "' "
'   wc = " BillTxnDate >= '" & Format(SFrom, Gen_DatFormat) & "' And BillTxnDate <= '" & Format(STo, Gen_DatFormat) & "' "
'End If

If CmbSSecurityAgency.Text <> "" Then
   tmp = tmp & " AND MSA.AgencyName = '" & CmbSSecurityAgency.Text & "' "
   wc = wc & " And SecurityAgencyID = " & mSSecurityAgencyID & ""
End If

If TxtSBillNo.Text <> "" Then
   tmp = tmp & " AND TSB.BillNo like  '%" & TxtSBillNo.Text & "%'  "
   'wc = wc & " And SecurityAgencyID = " & mSSecurityAgencyID & ""
End If

If CmbSStatus.Text <> "" Then
   tmp = tmp & " AND TSB.Status = '" & Left(CmbSStatus.Text, 1) & "' "
   'wc = wc & " And SecurityAgencyID = " & mSSecurityAgencyID & ""
End If

If ChkBatchNo.Value = 0 Then
   tmp = tmp & " AND TSB.BatchID = 0 "
   
  ' wc = wc & " And BatchID = 0 "
End If

strSplitString = Split(CmbSForwardTo, "~")
If CmbSForwardTo.Text <> "" Then
   tmp = tmp & " AND ME.EmployeeName = '" & strSplitString(0) & "' "
   'wc = wc & " And ForwardTo = " & mSForwardToID
End If

Call TmpTable
Call Grid_Head1

If TmpRs.RecordCount > 0 Then
   MSHFlexGrid1.Rows = TmpRs.RecordCount + 2
   For I = 1 To TmpRs.RecordCount
       MSHFlexGrid1.TextMatrix(I, 0) = TmpRs!BillTxnID
       MSHFlexGrid1.TextMatrix(I, 1) = TmpRs!BillTxnDate
       MSHFlexGrid1.TextMatrix(I, 1) = Format(TmpRs!BillTxnDate, "dd-mm-yyyy")
       'mLocationId = TmpRs!LocationID
       
       MSHFlexGrid1.TextMatrix(I, 2) = TmpRs!Month & "~" & TmpRs!Year
       MSHFlexGrid1.TextMatrix(I, 3) = IIf(IsNull(TmpRs!AgencyName), "", TmpRs!AgencyName)
       MSHFlexGrid1.TextMatrix(I, 4) = TmpRs!BillNo
       'MSHFlexGrid1.TextMatrix(I, 4) = Format(TmpRs!YearMonth, "dd-mm-yyyy")
       MSHFlexGrid1.TextMatrix(I, 5) = TmpRs!BillDate
       MSHFlexGrid1.TextMatrix(I, 5) = Format(TmpRs!BillDate, "dd-mm-yyyy")
       MSHFlexGrid1.TextMatrix(I, 6) = TmpRs!BillAmt
        MSHFlexGrid1.TextMatrix(I, 7) = IIf(IsNull(TmpRs!BillSTax), "", TmpRs!BillSTax)
       
       MSHFlexGrid1.TextMatrix(I, 8) = IIf(IsNull(TmpRs!TotalAmt), "", TmpRs!TotalAmt) 'TmpRs!TINO
       MSHFlexGrid1.TextMatrix(I, 9) = IIf(IsNull(TmpRs!ForwardTo), "", TmpRs!ForwardTo) '
'       MSHFlexGrid1.TextMatrix(I, 10) = Format(TmpRs!TIDate, "dd-mm-yyyy")
       MSHFlexGrid1.TextMatrix(I, 15) = IIf(IsNull(TmpRs!Description), "", TmpRs!Description)
       MSHFlexGrid1.TextMatrix(I, 10) = IIf(IsNull(TmpRs!Remarks), 0, TmpRs!Remarks) '
       MSHFlexGrid1.TextMatrix(I, 11) = TmpRs!Status
       If LCase(MSHFlexGrid1.TextMatrix(I, 11)) = "p" Then
          MSHFlexGrid1.TextMatrix(I, 11) = "Pending"
       ElseIf LCase(MSHFlexGrid1.TextMatrix(I, 11)) = "r" Then
          MSHFlexGrid1.TextMatrix(I, 11) = "Rejected"
       ElseIf LCase(MSHFlexGrid1.TextMatrix(I, 11)) = "a" Then
          MSHFlexGrid1.TextMatrix(I, 11) = "Approved"
       End If
       
       MSHFlexGrid1.TextMatrix(I, 12) = IIf(IsNull(TmpRs!ApprovedDate), "", TmpRs!ApprovedDate)
       If MSHFlexGrid1.TextMatrix(I, 12) <> "" Then
          MSHFlexGrid1.TextMatrix(I, 12) = Format(MSHFlexGrid1.TextMatrix(I, 12), "dd-mm-yyyy")
       End If
       MSHFlexGrid1.TextMatrix(I, 13) = IIf(IsNull(TmpRs!FAGRemarks), "", TmpRs!FAGRemarks) '
       
       MSHFlexGrid1.TextMatrix(I, 14) = IIf(IsNull(TmpRs!BatchID), 0, TmpRs!BatchID)
       '
       'MSHFlexGrid1.TextMatrix(I, 16) = IIf(IsNull(TmpRs!ApprovalDate), "", TmpRs!ApprovalDate)
       'MSHFlexGrid1.TextMatrix(I, 17) = IIf(IsNull(TmpRs!EmployeeName), "", TmpRs!EmployeeName)
       
'TSB.BankName,TSB.ChequeNo,TSB.PaidAmount,TSB.TDS,TSB.PaidDate,TSB.PaymentdetailsUpdatedBy,"
       
       MSHFlexGrid1.TextMatrix(I, 16) = IIf(IsNull(TmpRs!BankName), "", TmpRs!BankName)
       MSHFlexGrid1.TextMatrix(I, 17) = IIf(IsNull(TmpRs!ChequeNo), "", TmpRs!ChequeNo)
       MSHFlexGrid1.TextMatrix(I, 18) = IIf(IsNull(TmpRs!PaidAmount), "", TmpRs!PaidAmount)
       MSHFlexGrid1.TextMatrix(I, 19) = IIf(IsNull(TmpRs!TDS), "", TmpRs!TDS)
       MSHFlexGrid1.TextMatrix(I, 20) = IIf(IsNull(TmpRs!PaidDate), "", TmpRs!PaidDate)
       MSHFlexGrid1.TextMatrix(I, 21) = IIf(IsNull(TmpRs!PaymentdetailsUpdatedBy), "", TmpRs!PaymentdetailsUpdatedBy)
       MSHFlexGrid1.TextMatrix(I, 22) = IIf(IsNull(TmpRs!PaymentdetailsUpdatedOn), "", TmpRs!PaymentdetailsUpdatedOn)
       
       MSHFlexGrid1.TextMatrix(I, 23) = IIf(IsNull(TmpRs!StateIndiaName), "", TmpRs!StateIndiaName)     ''-- Ullhas Patil 04 Oct 2017
       MSHFlexGrid1.TextMatrix(I, 24) = IIf(IsNull(TmpRs!BillingStateName), "", TmpRs!BillingStateName) ''-- Ullhas Patil 25 Oct 2017
',MFB.BankName+'~'+MFB.AccountNo 'FAGBankName' ,TSB.DisbuAmount,TSB.DisbuDate,TSB.ModeOfPay,TSB.DeductionRemark
       If MSHFlexGrid1.TextMatrix(I, 16) = "" Then
          MSHFlexGrid1.TextMatrix(I, 16) = IIf(IsNull(TmpRs!FAGBankName), "", TmpRs!FAGBankName)
       End If
       
       If MSHFlexGrid1.TextMatrix(I, 17) = "" Then
          MSHFlexGrid1.TextMatrix(I, 17) = IIf(IsNull(TmpRs!ModeOfPay), "", TmpRs!ModeOfPay)
       End If
       
       If MSHFlexGrid1.TextMatrix(I, 18) = "" Then
          MSHFlexGrid1.TextMatrix(I, 18) = IIf(IsNull(TmpRs!DisbuAmount), "", TmpRs!DisbuAmount)
       End If
       
       If MSHFlexGrid1.TextMatrix(I, 19) = "" Then
          MSHFlexGrid1.TextMatrix(I, 19) = IIf(IsNull(TmpRs!DeductionRemark), "", TmpRs!DeductionRemark)
       End If
       
       If MSHFlexGrid1.TextMatrix(I, 20) = "" Then
          MSHFlexGrid1.TextMatrix(I, 20) = IIf(IsNull(TmpRs!DisbuDate), "", Format(TmpRs!DisbuDate, "DD-MMM-YYYY"))
       End If
       
       'MSHFlexGrid1.TextMatrix(I, 3) = Format(MSHFlexGrid1.TextMatrix(I, 3), "dd-mmm-yyyy")
       'MSHFlexGrid1.TextMatrix(I, 4) = Format(MSHFlexGrid1.TextMatrix(I, 4), "dd-mmm-yyyy")
       
       TmpRs.MoveNext
   Next
    
End If

' Commented On 10082011 on 12.25 PM
'For iu = 1 To TmpRs.RecordCount
'    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
'    For c = 0 To MSHFlexGrid1.Cols - 1
'        MSHFlexGrid1.TextMatrix(iu, c) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
''        If c = 2 Then
''           MSHFlexGrid1.TextMatrix(iu, c) = MonthName(MSHFlexGrid1.TextMatrix(iu, c))
''        End If
'        If c = 13 Then
'           If LCase(Trim(MSHFlexGrid1.TextMatrix(iu, c))) = "p" Then
'              MSHFlexGrid1.TextMatrix(iu, c) = "Pending"
'           ElseIf LCase(Trim(MSHFlexGrid1.TextMatrix(iu, c))) = "a" Then
'              MSHFlexGrid1.TextMatrix(iu, c) = "Approved"
'           ElseIf LCase(Trim(MSHFlexGrid1.TextMatrix(iu, c))) = "r" Then
'              MSHFlexGrid1.TextMatrix(iu, c) = "Rejected"
'           End If
'        End If
'    Next
'    TmpRs.MoveNext
'Next

Label22.Refresh
Label22.Caption = TmpRs.RecordCount
End Sub

Private Sub CmdSecurityAgency_Click()
FrmMst_SecurityAgency.Show
End Sub

Private Sub CmdSTaxDesc_Click()
FrmMst_ServiceTax.Show
End Sub

Private Sub DeleteCmd_Click()
Dim ans As Variant
On Error GoTo msgError

If TxtTxnID.Text = "" Then Exit Sub


If LCase(TxtStatus) = "approved" Or LCase(TxtStatus) = "rejected" Then
   MsgBox "This Billing is " & TxtStatus & ". You can't delete this Bill any more!!!"
   Exit Sub
End If

ans = MsgBox("Do you really want to DELETE this record???", vbYesNo)
If ans = vbYes Then
   
   Call UpdateMonthWiseDetails(mMonWiseID, "D", I, BillingID)
   
   tmp = "Delete From Txn_SecurityBillingDetail where BillTxnID = " & Val(TxtTxnID)
   Call TmpTable
   
   RsTxn_SecurityBilling.Delete
   RsTxn_SecurityBilling.Update
   
   If RsTxn_SecurityBilling.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   
   RsTxn_SecurityBilling.MoveNext
   If RsTxn_SecurityBilling.EOF() Then
      RsTxn_SecurityBilling.MoveLast
   End If
   Call Indata
End If


Exit Sub
msgError:
MsgBox "Child record Present "
RsTxn_SecurityBilling.CancelUpdate
End Sub

Private Sub EditCmd_Click()

'If LCase(TxtStatus) = "approved" Or LCase(TxtStatus) = "rejected" Then
'   MsgBox "This Billing is " & TxtStatus & ". You can't edit this Bill any more!!!"
'   Exit Sub
'End If

If LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "fag" And LCase(Gen_UserRole) <> "fagadmin" Then
   If LCase(TxtStatus) = "settled" Or LCase(TxtStatus) = "approved" Or LCase(TxtStatus) = "rejected" Or LCase(TxtStatus) = "frozen" Then
      MsgBox "This Txn is " & TxtStatus & ". You can't edit this txn any more!!!"
      Exit Sub
   End If
Else
   If LCase(TxtStatus) = "rejected" Or LCase(TxtStatus) = "settled" Then
      MsgBox "This Txn is " & TxtStatus & ". You can't edit this txn any more!!!"
      Exit Sub
   End If
End If

Edit_Flg = "E"
Call GButtonLock(Me, False)
CmdPrint.Enabled = False
'TxtMonth.Visible = False
'CmbMonth.Visible = True
CmbMonth.Text = TxtMonth

'TxtYear.Visible = False
'CmbYear.Visible = True
CmbYear.Text = TxtYear

mMonth = Month(CmbYear & "-" & CmbMonth.Text & "-" & 1)
intMonthDays = Day(DateSerial(CmbYear.Text, mMonth + 1, 0))
mMonthStartDt = mMonth & "/01" & "/" & CmbYear
mMonthEndDate = mMonth & "/" & intMonthDays & "/" & CmbYear

'TxtSecurityAgency.Visible = False
'CmbSecurityAgency.Visible = True
CmbSecurityAgency.Text = TxtSecurityAgency
'CmdSecurityAgency.Visible = True
CmbStatus.Enabled = True
CmbSTaxDesc.Text = TxtSTaxDesc
'CmdSTaxDesc.Visible = True
If LCase(Gen_UserRole) = "power" Then
   TxtSTaxDesc.Visible = False
   CmbSTaxDesc.Visible = True
End If

TxtForwardTo.Visible = False
CmbForwardTo.Visible = True
CmbForwardTo.Text = TxtForwardTo
CmdForwardTo.Visible = True

 
'TxtBillNo.Locked = True
'TxtBillDate.Enabled = False
'TxtBillAmt.Locked = True
'TxtSTax.Locked = True
TxtFAGRemarks.Locked = True
TxtRemarks.Locked = False

CmbStatus.Visible = True
CmbStatus.Text = TxtStatus

''--- Ullhas Patil 04 Oct 2017
CmbStateIndiaName.Visible = False
TxtStateIndiaName.Visible = True
CmbStateIndiaName = TxtStateIndiaName
''--- Ullhas Patil 25 Oct 2017
CmbBillingState.Visible = True
TxtBillingState.Visible = False
CmbBillingState = TxtBillingState
''-----------------------------

Frame1.Enabled = True
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
RsTxn_SecurityBilling.MoveFirst
RsTxn_SecurityBilling.Find "BillTxnID = '" & MSHFlexGrid1.TextMatrix(I, 0) & "'"
If Not RsTxn_SecurityBilling.EOF Then
   Call Indata
   Frame0.Visible = False
End If
End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsTxn_SecurityBilling.MoveNext
If RsTxn_SecurityBilling.EOF Then
   RsTxn_SecurityBilling.MoveLast
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
RsTxn_SecurityBilling.MovePrevious
If RsTxn_SecurityBilling.BOF Then
   RsTxn_SecurityBilling.MoveFirst
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
RsTxn_SecurityBilling.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsTxn_SecurityBilling.MoveLast
Call Indata
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False
End Sub

Private Sub ChkCheckAll_Click()
If MSHFlexGrid2.Rows = 2 Then
   ChkCheckAll.Value = 0
   Exit Sub
End If

If ChkCheckAll.Value = vbChecked Then
   For I = 2 To MSHFlexGrid2.Rows - 2
       MSHFlexGrid2.row = I
       MSHFlexGrid2.Col = 0
       MSHFlexGrid2.TextMatrix(I, 0) = strChecked
       MSHFlexGrid2.TextMatrix(I, 18) = Val(MSHFlexGrid2.TextMatrix(I, 12))
       MSHFlexGrid2.TextMatrix(I, 16) = Val(MSHFlexGrid2.TextMatrix(I, 13))
       MSHFlexGrid2.TextMatrix(I, 17) = Val(MSHFlexGrid2.TextMatrix(I, 14))
       MSHFlexGrid2.TextMatrix(0, 16) = Val(MSHFlexGrid2.TextMatrix(0, 16)) + Val(MSHFlexGrid2.TextMatrix(I, 16))
       MSHFlexGrid2.TextMatrix(0, 17) = Val(MSHFlexGrid2.TextMatrix(0, 17)) + Val(MSHFlexGrid2.TextMatrix(I, 17))
   Next
Else
   For I = 2 To MSHFlexGrid2.Rows - 2
       MSHFlexGrid2.row = I
       MSHFlexGrid2.TextMatrix(I, 0) = strUnChecked
       MSHFlexGrid2.TextMatrix(I, 18) = 0 'Val(MSHFlexGrid2.TextMatrix(I, 12))
       MSHFlexGrid2.TextMatrix(I, 16) = 0 'Val(MSHFlexGrid2.TextMatrix(I, 13))
       MSHFlexGrid2.TextMatrix(I, 17) = 0 'Val(MSHFlexGrid2.TextMatrix(I, 14))
       MSHFlexGrid2.TextMatrix(0, 16) = 0 'Val(MSHFlexGrid2.TextMatrix(0, 16)) - Val(MSHFlexGrid2.TextMatrix(I, 16))
       MSHFlexGrid2.TextMatrix(0, 17) = 0 'Val(MSHFlexGrid2.TextMatrix(0, 17)) - Val(MSHFlexGrid2.TextMatrix(I, 17))
   Next
End If
End Sub

Private Sub CmbForwardTo_GotFocus()

tmp1 = CmbForwardTo.Text
Call TableMst_Employee("Where Flag = 'Y'")

 
CmbForwardTo.Clear
'If RsMst_Employee.RecordCount = 0 Then
'   MsgBox "No Employee Found!!!"
'   CmbForwardTo.SetFocus
'   'RsMst_Employee.MoveFirst
'   Exit Sub
'End If
'For i = 1 To RsMst_Employee.RecordCount
'    CmbForwardTo.AddItem RsMst_Employee!EmployeeName & "~" & RsMst_Employee!EmployeeNo
'    RsMst_Employee.MoveNext
'Next


If LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "superadmin" And LCase(Gen_UserRole) <> "headcmg" And LCase(Gen_UserRole) <> "admin" And LCase(Gen_UserRole) <> "fag" And LCase(Gen_UserRole) <> "fagadmin" Then
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

'RsMst_Employee.MoveFirst
'RsMst_Employee.Find "EmployeeNo = '" & EmpNo_ & "'"

'If RsMst_Employee.EOF Then
'   MsgBox "Invalid Selection "
'   CmbForwardTo.SetFocus
'   Exit Sub
'End If
'mForwardToID = RsMst_Employee!EmployeeID

TmpRs2.MoveFirst
TmpRs2.Find "EmployeeNo = '" & EmpNo_ & "'"

If TmpRs2.EOF Then
   MsgBox "Invalid Selection "
   CmbForwardTo.SetFocus
   Exit Sub
End If
mForwardToID = TmpRs2!EmployeeID


End Sub

Private Sub CmbMonth_GotFocus()
tmp = CmbMonth.Text
CmbMonth.Clear
For I = 1 To 12
    CmbMonth.AddItem MonthName(I)
Next
CmbMonth.Text = tmp
End Sub

Private Sub CmbMonth_LostFocus()
If CmbMonth.Text = "" Then Exit Sub
mStatus = False
For I = 0 To 11
    If Trim(CmbMonth.Text) = CmbMonth.List(I) Then
       CmbMonth.Text = CmbMonth.List(I)
       mStatus = True
       Exit For
    End If
Next

If mStatus = False Then
   MsgBox "Invalid Selection"
   CmbMonth.SetFocus
   Exit Sub
End If

If CmbYear.Text <> "" Then
    mMonth = Month(CmbYear & "-" & CmbMonth.Text & "-" & 1)
    intMonthDays = Day(DateSerial(CmbYear.Text, mMonth + 1, 0))
    mMonthStartDt = mMonth & "/01" & "/" & CmbYear
    mMonthEndDate = mMonth & "/" & intMonthDays & "/" & CmbYear
End If

mTotalDays = Day(DateSerial(mYear, mMonth + 1, 0))

If CmbSecurityAgency.Text <> "" And LCase(tmp) <> LCase(CmbMonth.Text) And CmbYear <> "" Then
   Call GetMonthwiseSecLoc
End If
End Sub

Private Sub CmbSecurityAgency_GotFocus()
If CmbStateIndiaName.Text = "" Then
   MsgBox "Select State India."
   CmbStateIndiaName.SetFocus
   Exit Sub
End If

tmp = CmbSecurityAgency.Text

Call TableMst_SecurityAgency("Where Flag = 'A'")

CmbSecurityAgency.Clear
If RsMst_SecurityAgency.RecordCount = 0 Then
   MsgBox "No Security Agency Found!!!"
   CmbSecurityAgency.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_SecurityAgency.RecordCount
    CmbSecurityAgency.AddItem RsMst_SecurityAgency!AgencyName
    RsMst_SecurityAgency.MoveNext
Next
CmbSecurityAgency.Text = tmp
End Sub

Private Sub CmbSecurityAgency_LostFocus()
If CmbSecurityAgency.Text = "" Then
   mSecurityAgencyID = Null
   Exit Sub
End If

RsMst_SecurityAgency.MoveFirst
RsMst_SecurityAgency.Find "AgencyName = '" & CmbSecurityAgency.Text & "'"

If RsMst_SecurityAgency.EOF Then
   MsgBox "Invalid Selection "
   CmbSecurityAgency.SetFocus
   Exit Sub
End If
mSecurityAgencyID = RsMst_SecurityAgency!SecurityAgencyID

If CmbMonth.Text <> "" And CmbYear.Text <> "" And LCase(tmp) <> LCase(RsMst_SecurityAgency!AgencyName) Then
    Call ClearCtrl
    Call GetMonthwiseSecLoc
End If


End Sub

Private Sub CmbSSecurityAgency_GotFocus()
tmp = CmbSSecurityAgency.Text
Call TableMst_SecurityAgency("Where Flag = 'Y'")
CmbSSecurityAgency.Clear
If RsMst_SecurityAgency.RecordCount = 0 Then
   MsgBox "No Security Agency Found!!!"
   CmbSSecurityAgency.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_SecurityAgency.RecordCount
    CmbSSecurityAgency.AddItem RsMst_SecurityAgency!AgencyName
    RsMst_SecurityAgency.MoveNext
Next
CmbSSecurityAgency.Text = tmp
End Sub

Private Sub CmbSSecurityAgency_LostFocus()
If CmbSSecurityAgency.Text = "" Then
   mSSecurityAgencyID = Null
   Exit Sub
End If

RsMst_SecurityAgency.MoveFirst
RsMst_SecurityAgency.Find "AgencyName = '" & CmbSSecurityAgency.Text & "'"

If RsMst_SecurityAgency.EOF Then
   MsgBox "Invalid Selection "
   CmbSSecurityAgency.SetFocus
   Exit Sub
End If
mSSecurityAgencyID = RsMst_SecurityAgency!SecurityAgencyID
End Sub

Private Sub CmbSTaxDesc_GotFocus()
tmp = CmbSTaxDesc.Text
Call TableMst_ServiceTax(" Where Flag = 'A' ")
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

Private Sub CmbYear_GotFocus()
tmp = CmbYear.Text
CmbYear.Clear
For I = 2011 To Year(Now)
    CmbYear.AddItem I
Next
If Month(Now) = 12 Then
   CmbYear.AddItem Year(Now) + 1
End If
CmbYear.Text = tmp
End Sub

Private Sub CmbYear_LostFocus()
If CmbYear.Text = "" Then Exit Sub
Dim flg As Variant
Dim myr As Variant

flg = "N"

If Month(Now) = 12 Then
   myr = Year(Now) + 1
Else
   myr = Year(Now)
End If

For I = 2010 To myr 'Year(Now)
    If CmbYear.Text = I Then
       flg = "Y"
       Exit For
    End If
Next
If flg = "N" Then
   MsgBox "Invalid Selection!!!"
   CmbYear.SetFocus
   Exit Sub
End If
If CmbMonth.Text <> "" Then
    mMonth = Month(CmbYear & "-" & CmbMonth.Text & "-" & 1)
    intMonthDays = Day(DateSerial(CmbYear.Text, mMonth + 1, 0))
    mMonthStartDt = mMonth & "/01" & "/" & CmbYear
    mMonthEndDate = mMonth & "/" & intMonthDays & "/" & CmbYear
End If

mTotalDays = Day(DateSerial(CmbYear, mMonth + 1, 0))
If CmbSecurityAgency.Text <> "" And LCase(tmp) <> LCase(CmbYear.Text) Then
    'Call EditMode
End If
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
   Call GButtonLockAD(Me)
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If
Call Grid_Head
Call Grid_Head1
tmp = "Select EmployeeID,EmployeeName +'~'+ EmployeeNo as Employee,DesignationID,EmployeeEmailID From Mst_Employee Where EmployeeNo = '" & Gen_UserLoginID & "'"
Call TmpTable

If TmpRs.RecordCount = 1 Then
   CmbSForwardTo.Text = TmpRs!Employee
   mSForwardToID = TmpRs!EmployeeID
Else
   CmbSForwardTo.Text = ""
   mSForwardToID = Null
End If

SFrom.Value = CDate(From_GenDate) - 900
STo.Value = Date 'To_GenDate
CmbSSecurityAgency.Text = ""
'CmbSForwardTo.Text = ""
Label22.Caption = ""
OptMonthYear.Value = True

Call TableTxn_SecurityBilling
If RsTxn_SecurityBilling.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If

Call GButtonLock(Me, True)
Call Indata
End Sub

Private Sub Grid_Head()

MSHFlexGrid2.Clear
MSHFlexGrid2.Rows = 3
MSHFlexGrid2.FixedRows = 2
MSHFlexGrid2.Cols = 21

MSHFlexGrid2.TextMatrix(1, 0) = "Select"
MSHFlexGrid2.TextMatrix(1, 1) = "Location"
MSHFlexGrid2.TextMatrix(1, 2) = "LocationID"
MSHFlexGrid2.TextMatrix(1, 3) = "Security Type"
MSHFlexGrid2.TextMatrix(1, 4) = "Security TypeID"
MSHFlexGrid2.TextMatrix(1, 5) = "Security Designation"
MSHFlexGrid2.TextMatrix(1, 6) = "Security DesignationID"
MSHFlexGrid2.TextMatrix(1, 7) = "No of Godown"
MSHFlexGrid2.TextMatrix(1, 8) = "Rate per month"
MSHFlexGrid2.TextMatrix(1, 9) = "CMPM Name"
MSHFlexGrid2.TextMatrix(1, 10) = "CMPID"
MSHFlexGrid2.TextMatrix(1, 11) = "Attendance Sheet ID"
MSHFlexGrid2.TextMatrix(1, 12) = "No of Days"
MSHFlexGrid2.TextMatrix(1, 13) = "Esitmated Amount"
MSHFlexGrid2.TextMatrix(1, 14) = "Esitmated S.Tax"
MSHFlexGrid2.TextMatrix(1, 15) = "Total Esitmated Amount"
MSHFlexGrid2.TextMatrix(1, 16) = "Billing Amount"
MSHFlexGrid2.TextMatrix(1, 17) = "S. Tax"
MSHFlexGrid2.TextMatrix(1, 18) = "Billing No Of Days"
MSHFlexGrid2.TextMatrix(1, 19) = "MonthwiseTxnDetailID"
MSHFlexGrid2.TextMatrix(1, 20) = "Processed Date"

MSHFlexGrid2.ColWidth(0) = 550
MSHFlexGrid2.ColWidth(1) = 1600
MSHFlexGrid2.ColWidth(2) = 0
MSHFlexGrid2.ColWidth(3) = 1000
MSHFlexGrid2.ColWidth(4) = 0
MSHFlexGrid2.ColWidth(5) = 1200
MSHFlexGrid2.ColWidth(6) = 0
MSHFlexGrid2.ColWidth(7) = 700
MSHFlexGrid2.ColWidth(8) = 700
MSHFlexGrid2.ColWidth(9) = 1800
MSHFlexGrid2.ColWidth(10) = 0
MSHFlexGrid2.ColWidth(11) = 800
MSHFlexGrid2.ColWidth(12) = 600
MSHFlexGrid2.ColWidth(13) = 1000
MSHFlexGrid2.ColWidth(14) = 950
MSHFlexGrid2.ColWidth(15) = 1000
MSHFlexGrid2.ColWidth(16) = 1000
MSHFlexGrid2.ColWidth(17) = 800
MSHFlexGrid2.ColWidth(18) = 1000
MSHFlexGrid2.ColWidth(19) = 0
MSHFlexGrid2.ColWidth(20) = 1000

MSHFlexGrid2.RowHeight(1) = 650
MSHFlexGrid2.WordWrap = True

End Sub

Private Sub MSHFlexGrid2_Click()
If SaveCmd.Enabled = False Then Exit Sub


With MSHFlexGrid2
    
    If .TextMatrix(.row, 2) = "" Then Exit Sub
    '
    If LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "headcmg" And LCase(Gen_UserRole) <> "superadmin" And LCase(Gen_UserRole) <> "admin" Then
       If CDate(.TextMatrix(.row, 20)) < Date - Val(Gen_BlockingSecurityBilling) Then
          MsgBox "For marking billing beyond the past " & Val(Gen_BlockingSecurityBilling) & " day(s), you need to provide the details to CMG with reasons for such delay in marking the bills. CMG will mark the bills on your behalf, if they are satisfied with your explanation!!!"
          If .TextMatrix(.row, 0) = strChecked Then
             .TextMatrix(.row, 0) = strUnChecked
          End If
          Exit Sub
       End If
    End If
    
    If .Col = 0 Then
        If .TextMatrix(.row, 0) = strUnChecked Then
           .TextMatrix(.row, 0) = strChecked
           .TextMatrix(.row, 18) = Val(.TextMatrix(.row, 12))
           .TextMatrix(.row, 16) = Val(.TextMatrix(.row, 13))
           .TextMatrix(.row, 17) = Val(.TextMatrix(.row, 14))
           .TextMatrix(0, 16) = Val(.TextMatrix(0, 16)) + Val(.TextMatrix(.row, 16))
           .TextMatrix(0, 17) = Val(.TextMatrix(0, 17)) + Val(.TextMatrix(.row, 17))
        Else
           .TextMatrix(.row, 0) = strUnChecked
           .TextMatrix(.row, 16) = 0
           .TextMatrix(.row, 17) = 0
           .TextMatrix(.row, 18) = 0
           .TextMatrix(0, 16) = Val(.TextMatrix(0, 16)) - Val(.TextMatrix(.row, 16))
           .TextMatrix(0, 17) = Val(.TextMatrix(0, 17)) - Val(.TextMatrix(.row, 17))
        End If
    End If
End With

If MSHFlexGrid2.Col = 18 Or MSHFlexGrid2.Col = 16 Or MSHFlexGrid2.Col = 17 Then
    If MSHFlexGrid2.TextMatrix(MSHFlexGrid2.row, 0) = strUnChecked Then Exit Sub
    mSelRow = MSHFlexGrid2.row
    mSelCol = MSHFlexGrid2.Col
    TxtColAmt.Visible = True
    TxtColAmt.Text = MSHFlexGrid2.TextMatrix(MSHFlexGrid2.row, mSelCol)
    With TxtColAmt
        .Left = MSHFlexGrid2.Left + MSHFlexGrid2.CellLeft
        .Top = MSHFlexGrid2.Top + MSHFlexGrid2.CellTop
        .Width = MSHFlexGrid2.CellWidth - 10
        .Height = MSHFlexGrid2.CellHeight - 10
        .ZOrder
        .SetFocus
    End With
End If

Call CalcInGrid
End Sub

Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   If RsTxn_SecurityBilling.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   RsTxn_SecurityBilling.MoveFirst

   RsTxn_SecurityBilling.Find "G_UID = '" & TxtG_UID & "'"
   If RsTxn_SecurityBilling.EOF Then
      RsTxn_SecurityBilling.MoveLast
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

RsTxn_SecurityBilling.MoveFirst

''===========comment on 7 july 2012 for not clearing search grid=============
'Call Grid_Head1
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
'
'
'
'SFrom.Value = CDate(From_GenDate) - 900
'STo.Value = Date 'To_GenDate
'CmbSSecurityAgency.Text = ""
''CmbSForwardTo.Text = ""
'Label22.Caption = ""
'OptMonthYear.Value = True
''===========comment on 7 july 2012 for not clearing search grid=============

End Sub

Private Sub Grid_Head1()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.FixedRows = 1
MSHFlexGrid1.Cols = 25
MSHFlexGrid1.WordWrap = True
MSHFlexGrid1.RowHeight(0) = 700

MSHFlexGrid1.TextMatrix(0, 0) = "Security Txn ID"
MSHFlexGrid1.TextMatrix(0, 1) = "Security Txn Date"
MSHFlexGrid1.TextMatrix(0, 2) = "Month & Year"
MSHFlexGrid1.TextMatrix(0, 3) = "Security Agency"
MSHFlexGrid1.TextMatrix(0, 4) = "Bill No"
MSHFlexGrid1.TextMatrix(0, 5) = "Bill Date"
MSHFlexGrid1.TextMatrix(0, 6) = "Bill Amount"
MSHFlexGrid1.TextMatrix(0, 7) = "S.Tax"
MSHFlexGrid1.TextMatrix(0, 8) = "Total Bill Amt"
MSHFlexGrid1.TextMatrix(0, 9) = "Forwarded To"
MSHFlexGrid1.TextMatrix(0, 10) = "Remarks"
MSHFlexGrid1.TextMatrix(0, 11) = "Status"
MSHFlexGrid1.TextMatrix(0, 12) = "Booked In NAV Date"
MSHFlexGrid1.TextMatrix(0, 13) = "FAG Remarks"
MSHFlexGrid1.TextMatrix(0, 14) = "Batch No"
MSHFlexGrid1.TextMatrix(0, 15) = "S.Tax Desc"
MSHFlexGrid1.TextMatrix(0, 16) = "Bank Name"
MSHFlexGrid1.TextMatrix(0, 17) = "Cheque No"
MSHFlexGrid1.TextMatrix(0, 18) = "Paid Amount"
MSHFlexGrid1.TextMatrix(0, 19) = "TDS"
MSHFlexGrid1.TextMatrix(0, 20) = "Paid Date"
MSHFlexGrid1.TextMatrix(0, 21) = "Paymentdetails Updated By"
MSHFlexGrid1.TextMatrix(0, 22) = "Paymentdetails Updated On"

MSHFlexGrid1.TextMatrix(0, 23) = "State India Name"     ''-- Ullhas Patil 04 oct 2017
MSHFlexGrid1.TextMatrix(0, 24) = "Billing State India"   ''-- Ullhas Patil 25 oct 2017

MSHFlexGrid1.ColWidth(0) = 1000
MSHFlexGrid1.ColWidth(1) = 1200
MSHFlexGrid1.ColWidth(2) = 1200
MSHFlexGrid1.ColWidth(3) = 4200
MSHFlexGrid1.ColWidth(4) = 1000
MSHFlexGrid1.ColWidth(5) = 1000
MSHFlexGrid1.ColWidth(6) = 1400
MSHFlexGrid1.ColWidth(7) = 1000
MSHFlexGrid1.ColWidth(8) = 1000
MSHFlexGrid1.ColWidth(9) = 1200
MSHFlexGrid1.ColWidth(10) = 2500
MSHFlexGrid1.ColWidth(11) = 1000
MSHFlexGrid1.ColWidth(12) = 3000
MSHFlexGrid1.ColWidth(13) = 1000
MSHFlexGrid1.ColWidth(14) = 1550
MSHFlexGrid1.ColWidth(15) = 2500

End Sub

Public Sub Indata()

'Call ClearFrame
Frame1.Enabled = False
CmdPrint.Enabled = True
TxtBillDate.Enabled = True

TxtMonth.Visible = True
CmbMonth.Visible = False

TxtYear.Visible = True
CmbYear.Visible = False

TxtSecurityAgency.Visible = True
CmbSecurityAgency.Visible = False
CmdSecurityAgency.Visible = False

TxtForwardTo.Visible = True
CmbForwardTo.Visible = False
CmdForwardTo.Visible = False

TxtSTaxDesc.Visible = True
CmbSTaxDesc.Visible = False
CmdSTaxDesc.Visible = False

TxtStatus.Visible = True
CmbStatus.Visible = False

''--- Ullhas Patil 04 Oct 2017
CmbStateIndiaName.Visible = False: CmbStateIndiaName = ""
TxtStateIndiaName.Visible = True: TxtStateIndiaName = ""
''--- Ullhas Patil 25 Oct 2017
CmbBillingState.Visible = False: CmbBillingState = ""
TxtBillingState.Visible = True: TxtBillingState = ""
''-----------------------------

'CmdAdd2List.Enabled = False
TxtTxnID = RsTxn_SecurityBilling!BillTxnID
TxtTxnDate.Value = RsTxn_SecurityBilling!BillTxnDate
TxtMonth = MonthName(Month(RsTxn_SecurityBilling!MonthYear))
CmbMonth.Text = TxtMonth
TxtYear = Year(RsTxn_SecurityBilling!MonthYear)
CmbYear.Text = TxtYear
mMonth = Month(CmbYear.Text & "-" & CmbMonth.Text & "-" & 1)
mTotalDays = Day(DateSerial(CmbYear.Text, mMonth + 1, 0))
mSecurityAgencyID = RsTxn_SecurityBilling!SecurityAgencyID
Call ShowAccountDetails(mSecurityAgencyID)
TxtSecurityAgency = GetSecurityAgency(RsTxn_SecurityBilling!SecurityAgencyID, "N")
TxtBillNo = RsTxn_SecurityBilling!BillNo
TxtBillDate.Value = RsTxn_SecurityBilling!BillDate
TxtBillAmt = RsTxn_SecurityBilling!BillAmt
TxtSTax = RsTxn_SecurityBilling!BillSTax
TxtTotBillAmt = RsTxn_SecurityBilling!TotalAmt
mServiceTaxID = RsTxn_SecurityBilling!ServiceTaxID
TxtSTaxDesc = GetServiceTaxDescription(RsTxn_SecurityBilling!ServiceTaxID)
mForwardToID = RsTxn_SecurityBilling!ForwardTo
TxtForwardTo = GetEmployeeNonName(RsTxn_SecurityBilling!ForwardTo)

If Trim(RsTxn_SecurityBilling!Status) = "P" Then
   TxtStatus.Text = "Pending"
ElseIf Trim(RsTxn_SecurityBilling!Status) = "A" Then
   TxtStatus.Text = "Approved"
ElseIf Trim(RsTxn_SecurityBilling!Status) = "R" Then
   TxtStatus.Text = "Rejected"
ElseIf Trim(RsTxn_SecurityBilling!Status) = "S" Then
   TxtStatus.Text = "Settled"
End If

TxtBatchNo = IIf(IsNull(RsTxn_SecurityBilling!BatchID), 0, RsTxn_SecurityBilling!BatchID)

TxtRemarks = IIf(IsNull(RsTxn_SecurityBilling!Remarks), "", RsTxn_SecurityBilling!Remarks)
TxtFAGRemarks = IIf(IsNull(RsTxn_SecurityBilling!FAGRemarks), "", RsTxn_SecurityBilling!FAGRemarks)

If IsNull(RsTxn_SecurityBilling!ApprovedDate) = False Then
   TxtApprovalDate.Visible = True
   Label27.Visible = True
   TxtApprovalDate.Value = RsTxn_SecurityBilling!ApprovedDate
   TxtApprovalDate.Enabled = False
Else
   TxtApprovalDate.Visible = False
   Label27.Visible = False
End If

CmbModeOfPay.Visible = False
TxtModeOfPayment.Visible = True

CmbFAGBank.Visible = False
TxtFAGBank.Visible = True

TxtFAGDisbAmt = IIf(IsNull(RsTxn_SecurityBilling!DisbuAmount), "", RsTxn_SecurityBilling!DisbuAmount)
TxtFAGRemark = IIf(IsNull(RsTxn_SecurityBilling!DisbuRemark), "", RsTxn_SecurityBilling!DisbuRemark)
TxtFagDisbDate.Value = IIf(IsNull(RsTxn_SecurityBilling!DisbuDate), Date, RsTxn_SecurityBilling!DisbuDate)
mFAGBankID = IIf(IsNull(RsTxn_SecurityBilling!FAGBankID), 0, RsTxn_SecurityBilling!FAGBankID)
TxtFAGBank.Text = GetFAGBankAcNo(Val(mFAGBankID))
TxtModeOfPayment.Text = IIf(IsNull(RsTxn_SecurityBilling!ModeOfPay), "", RsTxn_SecurityBilling!ModeOfPay)
TxtDeductionRemark = IIf(IsNull(RsTxn_SecurityBilling!DeductionRemark), "", RsTxn_SecurityBilling!DeductionRemark)

TxtDocNo = IIf(IsNull(RsTxn_SecurityBilling!DocNo), "", RsTxn_SecurityBilling!DocNo)
TxtEntryNo = IIf(IsNull(RsTxn_SecurityBilling!EntryNo), "", RsTxn_SecurityBilling!EntryNo)
TxtExportedDate = IIf(IsNull(RsTxn_SecurityBilling!FAGPayExportedDate), "", Format(RsTxn_SecurityBilling!FAGPayExportedDate, "dd-MMM-yyyy"))

TxtBankID.Text = IIf(IsNull(RsTxn_SecurityBilling!BankName), "", RsTxn_SecurityBilling!BankName)
TxtChequeNo.Text = IIf(IsNull(RsTxn_SecurityBilling!ChequeNo), "", RsTxn_SecurityBilling!ChequeNo)
TxtPaidAmt.Text = IIf(IsNull(RsTxn_SecurityBilling!PaidAmount), "", Format(RsTxn_SecurityBilling!PaidAmount, "#####0.00"))
TxtTDS.Text = IIf(IsNull(RsTxn_SecurityBilling!TDS), "", Format(RsTxn_SecurityBilling!TDS, "#####0.00"))
TxtPaidDate.Text = IIf(IsNull(RsTxn_SecurityBilling!PaidDate), "", Format(RsTxn_SecurityBilling!PaidDate, "dd-MMM-yyyy"))
mBatchID = IIf(IsNull(RsTxn_SecurityBilling!BatchID), 0, RsTxn_SecurityBilling!BatchID)
mPBatchID = IIf(IsNull(RsTxn_SecurityBilling!PBatchID), 0, RsTxn_SecurityBilling!PBatchID)

Frame3.Enabled = False

TxtFAGRemarks.Locked = True
TxtG_UID = IIf(IsNull(RsTxn_SecurityBilling!G_UID), "", RsTxn_SecurityBilling!G_UID)
TxtCreated = IIf(IsNull(RsTxn_SecurityBilling!CreatedBy), "", RsTxn_SecurityBilling!CreatedBy) & " :- " & IIf(IsNull(RsTxn_SecurityBilling!CreatedDate), "", RsTxn_SecurityBilling!CreatedDate)
TxtUpdated = IIf(IsNull(RsTxn_SecurityBilling!UpdatedBy), "", RsTxn_SecurityBilling!UpdatedBy) & " :- " & IIf(IsNull(RsTxn_SecurityBilling!UpdatedDate), "", RsTxn_SecurityBilling!UpdatedDate)

''--- Ullhas Patil 04 Oct 2017
mStateIndiaID = IIf(IsNull(RsTxn_SecurityBilling!StateIndiaID), 0, RsTxn_SecurityBilling!StateIndiaID)
TxtStateIndiaName = GetStateIndiaName(mStateIndiaID)
TxtStateIndiaName = IIf(Trim(TxtStateIndiaName) = "0", "", TxtStateIndiaName)
''--- Ullhas Patil 25 Oct 2017
mBillingStateID = IIf(IsNull(RsTxn_SecurityBilling!BillingStateIndiaID), 0, RsTxn_SecurityBilling!BillingStateIndiaID)
TxtBillingState = GetStateIndiaName(mBillingStateID)
TxtBillingState = IIf(Trim(TxtBillingState) = "0", "", TxtBillingState)
''-----------------------------

Call InDataDetails
If RsTxn_SecurityBilling.RecordCount = 1 Then
   Call GButtonLock_1(Me, True)
Else
   Call GButtonLock(Me, True)
End If
blnSave = False

SSTab1.Tab = 0
End Sub

Private Sub InDataDetails()
Dim strSplitString()        As String
Dim strEmployeeName         As String

Call Grid_Head

tmp = " SELECT TSBD.LocationID, LocationName, TSBD.SecurityTypeID, SecurityShifts, TSBD.SecurityDesignationID,"
tmp = tmp & " SecurityDesignation, TSBD.NoofGodown, TSBD.RatePerMonth, ME.EmployeeName +'~'+ ME.EmployeeNo as ForwardTo,"
tmp = tmp & " Me.EmployeeID, TSBD.AttendanceSheetID, TSBD.NoofDays, TSBD.EstimatedAmt, TSBD.EstimatedSTax, TSBD.TotEstimatedAmt,"
tmp = tmp & " TSBD.BillDetailAmt , TSBD.BillDetailSTax, TSBD.BillNoOfDays, TSBD.MonthwiseTxnDetailID, TMS.MonthwiseTxnDate"
tmp = tmp & " FROM Txn_SecurityBillingDetail TSBD"
tmp = tmp & " Inner Join Mst_SecurityShifts MSS On MSS.SecurityShiftsID = TSBD.SecurityTypeID"
tmp = tmp & " Inner Join Mst_SecurityDesignation MSD On MSD.SecurityDesignationID = TSBD.SecurityDesignationID"
tmp = tmp & " Inner Join Mst_Location ML On ML.LocationID = TSBD.LocationID"
''-- Ullhas Patil 04 Oct 2017
tmp = tmp & " Inner Join Mst_State MS On ML.StateID = MS.StateID"
''---------------------------
tmp = tmp & " Left Join Mst_Employee ME On ME.EmployeeID = TSBD.CMPMID"
tmp = tmp & " Inner Join Txn_MonthwiseSecLocDetail TMSD On TSBD.MonthwiseTxnDetailID = TMSD.MonthwiseTxnDetailID"
tmp = tmp & " Inner Join Txn_MonthwiseSecLoc TMS On TMSD.MonthwiseTxnID = TMS.MonthwiseTxnID"
tmp = tmp & " Where TSBD.BillTxnID =" & TxtTxnID.Text & ""
''-- Ullhas Patil 04 Oct 2017
If Val(mStateIndiaID) > 0 Then
tmp = tmp & " And MS.StateIndiaID = " & mStateIndiaID & " "
End If
''---------------------------
Call TmpTable

If TmpRs.RecordCount > 0 Then
    MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + TmpRs.RecordCount
    For I = 2 To MSHFlexGrid2.Rows - 2
        MSHFlexGrid2.row = I
        MSHFlexGrid2.Col = 0
        MSHFlexGrid2.CellFontName = "Wingdings"
        MSHFlexGrid2.CellFontSize = 12
        MSHFlexGrid2.CellAlignment = flexAlignCenterCenter
        MSHFlexGrid2.Text = strChecked
        MSHFlexGrid2.TextMatrix(I, 1) = TmpRs!Locationname
        MSHFlexGrid2.TextMatrix(I, 2) = TmpRs!LocationID
        MSHFlexGrid2.TextMatrix(I, 3) = TmpRs!SecurityShifts
        MSHFlexGrid2.TextMatrix(I, 4) = TmpRs!SecurityTypeID
        MSHFlexGrid2.TextMatrix(I, 5) = TmpRs!SecurityDesignation
        MSHFlexGrid2.TextMatrix(I, 6) = TmpRs!SecurityDesignationID
        MSHFlexGrid2.TextMatrix(I, 7) = TmpRs!NoofGodown
        MSHFlexGrid2.TextMatrix(0, 7) = Val(MSHFlexGrid2.TextMatrix(0, 7)) + Val(TmpRs!NoofGodown)
        MSHFlexGrid2.TextMatrix(I, 8) = TmpRs!RatePerMonth
        MSHFlexGrid2.TextMatrix(I, 9) = IIf(IsNull(TmpRs!ForwardTo), "", TmpRs!ForwardTo)
        MSHFlexGrid2.TextMatrix(I, 10) = IIf(IsNull(TmpRs!EmployeeID), "", TmpRs!EmployeeID)
        MSHFlexGrid2.TextMatrix(I, 11) = IIf(IsNull(TmpRs!AttendanceSheetID), "", TmpRs!AttendanceSheetID)
        MSHFlexGrid2.TextMatrix(I, 12) = TmpRs!NoOfDays
        MSHFlexGrid2.TextMatrix(I, 13) = TmpRs!EstimatedAmt
        MSHFlexGrid2.TextMatrix(I, 14) = IIf(IsNull(TmpRs!EstimatedSTax), "", TmpRs!EstimatedSTax)
        MSHFlexGrid2.TextMatrix(I, 15) = IIf(IsNull(TmpRs!TotEstimatedAmt), "", TmpRs!TotEstimatedAmt)
        MSHFlexGrid2.TextMatrix(I, 16) = TmpRs!BillDetailAmt
        MSHFlexGrid2.TextMatrix(I, 17) = TmpRs!BillDetailSTax
        MSHFlexGrid2.TextMatrix(I, 18) = IIf(IsNull(TmpRs!BillNoOfDays), 0, TmpRs!BillNoOfDays)
        MSHFlexGrid2.TextMatrix(I, 19) = TmpRs!MonthwiseTxnDetailID
        MSHFlexGrid2.TextMatrix(I, 20) = TmpRs!MonthwiseTxnDate
        TmpRs.MoveNext
    Next
End If
Call CalcInGrid
End Sub

Private Sub TxtBillAmt_LostFocus()
If Trim(TxtBillAmt.Text) = "" Then Exit Sub
If IsNumeric(TxtBillAmt) = False Then
   MsgBox "Invalid Format!!!"
   TxtBillAmt.SetFocus
   Exit Sub
End If
TxtTotBillAmt.Text = Val(TxtBillAmt.Text) + Val(TxtSTax.Text)
End Sub

Private Sub TxtBillNo_LostFocus()
'If Trim(TxtBillNo.Text) = "" Then Exit Sub
'If IsNumeric(TxtBillNo) = False Then
'   MsgBox "Invalid Format!!!"
'   TxtBillNo.SetFocus
'   Exit Sub
'End If
End Sub

Private Sub TxtColAmt_LostFocus()
If TxtColAmt = "" Then Exit Sub
If IsNumeric(TxtColAmt.Text) = False Or Val(TxtColAmt.Text) < 0 Then
   MsgBox "Invalid Format!!!"
   TxtColAmt.SetFocus
   Exit Sub
End If

If mSelCol = 18 Then
   If Val(mTotalDays) < TxtColAmt Then
      MsgBox "Billing No Of Days must be less than total no of days of selected month!!!"
      TxtColAmt.SetFocus
      Exit Sub
   End If
   MSHFlexGrid2.TextMatrix(mSelRow, mSelCol) = Round(TxtColAmt.Text, 0)
Else
   MSHFlexGrid2.TextMatrix(mSelRow, mSelCol) = Round(TxtColAmt.Text, 2)

End If

TxtColAmt.Visible = False


'If mSelCol = 18 Then
''   Call CalcEstimateAmt
'End If

Call CalcInGrid

End Sub

Private Sub TxtColAmt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 27 Then 'Enter/ESC
   'MSHFlexGrid2.TextMatrix(1, 3) = Trim(TxtColAmt.Text)
   TxtColAmt.Visible = False
End If
End Sub

Private Sub TxtSTax_LostFocus()
If Trim(TxtSTax.Text) = "" Then Exit Sub
If IsNumeric(TxtSTax) = False Then
   MsgBox "Invalid Format!!!"
   TxtSTax.SetFocus
   Exit Sub
End If
TxtTotBillAmt.Text = Val(TxtBillAmt.Text) + Val(TxtSTax.Text)
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Txn_SecurityBilling.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Txn_SecurityBilling.xls")
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

Private Sub EditMode()

'tmp = " SELECT BillTxnID, BillTxnDate, BillNo, BillDate, BillAmt, BillSTax, TotalAmt, Description, "
'tmp = tmp & " EmployeeName +'~'+ EmployeeNo as ForwardTo, TSB.G_UID, TSB.CreatedBy, TSB.CreatedDate, TSB.UpdatedDate, TSB.UpdatedBy"
'tmp = tmp & " FROM Txn_SecurityBilling TSB"
'tmp = tmp & " Inner Join Mst_Employee ME On ME.EmployeeID = TSB.ForwardTo"
'tmp = tmp & " Inner Join Mst_ServiceTax MS On MS.ServiceTaxID = TSB.STaxID"
'tmp = tmp & " Where Month(MonthYear)=" & mMonth & " AND Year(MonthYear)=" & CmbYear.Text & " AND SecurityAgencyID=" & mSecurityAgencyID & ""
'Call TmpTable
'
'If Not TmpRs.EOF Then
'    Edit_Flg = "E"
'    TxtTxnID = TmpRs!BillTxnID
'    TxtTxnDate.Value = TmpRs!BillTxnDate
'    TxtBillNo.Text = TmpRs!BillNo
'    TxtBillDate.Value = TmpRs!BillDate
'    TxtBillAmount.Text = TmpRs!BillAmt
'    TxtSTax.Text = TmpRs!BillSTax
'    TxtTotBillAmt.Text = TmpRs!TotalAmt
'    CmbSTaxDesc.Text = TmpRs!Description
'    CmbForwardTo.Text = TmpRs!ForwardTo
'    TxtFAGRemarks.Text = TmpRs!FAGRemarks
'    TxtCreated = IIf(IsNull(TmpRs!CreatedBy), "", TmpRs!CreatedBy) & " :- " & IIf(IsNull(TmpRs!CreatedDate), "", TmpRs!CreatedDate)
'    TxtUpdated = IIf(IsNull(TmpRs!UpdatedBy), "", TmpRs!UpdatedBy) & " :- " & IIf(IsNull(TmpRs!UpdatedDate), "", TmpRs!UpdatedDate)
'    'Call EditCmd_Click
'    CmbMonth.Visible = False
'    TxtMonth.Visible = True
'    CmbYear.Visible = False
'    TxtYear.Visible = True
'    CmbSTaxDesc.Visible = False
'    CmdSTaxDesc.Visible = False
'    TxtSTaxDesc.Visible = True
'    CmbSecurityAgency.Visible = False
'    CmdSecurityAgency.Visible = False
'    TxtSecurityAgency.Visible = True
'    CmbForwardTo.Visible = False
'    CmdForwardTo.Visible = False
'    TxtForwardTo.Visible = True
'Else
'    Edit_Flg = "A"
'    Call GetMonthwiseSecLoc
'    Exit Sub
'End If

'tmp = " SELECT TSBD.BillTxnDetailID, LocationID,LocationName, SecurityShifts, SecurityTypeID, SecurityDesignation, "
'tmp = tmp & " SecurityDesignationID, NoofGodown, RatePerMonth"
'tmp = tmp & " ME.EmployeeName + '~'+ ME.EmployeeNo As ForwardTo, CMPMID, AttendanceSheetID,"
'tmp = tmp & " NoofDays, EstimatedAmt, BillDetailAmt,BillDetailSTax, SecLocMapDetailID"
'tmp = tmp & " FROM Txn_SecurityBillingDetail TSBD"
'tmp = tmp & " Inner Join Mst_SecurityShifts MSS On MSS.SecurityShiftsID = TSBD.SecurityTypeID"
'tmp = tmp & " Inner Join Mst_SecurityDesignation MSD On MSD.SecurityDesignationID = TSBD.SecurityDesignationID"
'tmp = tmp & " Inner Join Mst_Employee ME On ME.EmployeeID = TSBD.CMPMID"
'tmp = tmp & " Inner Join Mst_Location ML On ML.LocationID = TSBD.LocationID"
'tmp = tmp & " Where MonthwiseTxnID = " & TxtTxnID & " ORDER BY MonthwiseTxnDetailID"
'Call TmpTable
'
'If TmpRs.RecordCount > 0 Then
'    For I = 2 To MSHFlexGrid2.Rows - 2
'        TmpRs.MoveFirst
'        TmpRs.Find "SecLocMapDetailID = " & MSHFlexGrid2.TextMatrix(I, 17)
'        If Not TmpRs.EOF Then
'           MSHFlexGrid2.TextMatrix(I, 1) = TmpRs!LocationName
'           MSHFlexGrid2.TextMatrix(I, 2) = TmpRs!LocationID
'           MSHFlexGrid2.TextMatrix(I, 3) = TmpRs!SecurityShifts
'           MSHFlexGrid2.TextMatrix(I, 4) = TmpRs!SecurityTypeID
'           MSHFlexGrid2.TextMatrix(I, 5) = TmpRs!SecurityDesignation
'           MSHFlexGrid2.TextMatrix(I, 6) = TmpRs!SecurityDesignationID
'           MSHFlexGrid2.TextMatrix(I, 7) = TmpRs!NoofGodown
'           MSHFlexGrid2.TextMatrix(0, 7) = Val(MSHFlexGrid2.TextMatrix(0, 7)) + Val(TmpRs!NoofGodown)
'           MSHFlexGrid2.TextMatrix(I, 8) = TmpRs!RatePerMonth
'           MSHFlexGrid2.TextMatrix(I, 9) = TmpRs!ForwardTo
'           MSHFlexGrid2.TextMatrix(I, 10) = TmpRs!EmployeeID
'           MSHFlexGrid2.TextMatrix(1, 11) = TmpRs!AttendanceSheetID
'           MSHFlexGrid2.TextMatrix(1, 12) = TmpRs!NoofDays
'           MSHFlexGrid2.TextMatrix(1, 13) = TmpRs!EstimatedAmt
'           MSHFlexGrid2.TextMatrix(1, 16) = TmpRs!BillDetailAmt
'           MSHFlexGrid2.TextMatrix(1, 17) = TmpRs!BillDetailSTax
'           MSHFlexGrid2.TextMatrix(I, 18) = TmpRs!BillingAmount
'        End If
'    Next
'End If
End Sub

Private Sub GetMonthwiseSecLoc()

Dim Retval As Boolean
Retval = True

If CmbSecurityAgency.Text = "" Then
   MsgBox "Blank Security Agency not allowed!!!"
   CmbSecurityAgency.SetFocus
   Exit Sub
End If

tmp = " SELECT TMSL.MonthwiseTxnID, TMSL.MonthwiseTxnDate, MonthYear, AgencyName"
tmp = tmp & " FROM Txn_MonthwiseSecLoc TMSL"
tmp = tmp & " Inner Join Mst_SecurityAgency MSA On MSA.SecurityAgencyID = TMSL.SecurityAgencyID "
tmp = tmp & " Where Month(MonthYear)=" & mMonth & " AND Year(MonthYear)=" & CmbYear.Text & " AND TMSL.SecurityAgencyID=" & mSecurityAgencyID & ""

Call TmpTable

If TmpRs.RecordCount > 0 Then
   Call InDataMonthwiseSecLocDetails
Else
   TxtTxnID.Text = ""
   TxtTxnDate.Value = Date
   TxtCreated = ""
   TxtUpdated = ""
   Call Grid_Head
End If
End Sub

Private Sub InDataMonthwiseSecLocDetails()
Dim strMonthTxnDetailID     As String

tmp = " SELECT MonthwiseTxnDetailID FROM Txn_SecurityBillingDetail TSBD"
tmp = tmp & " Inner Join Txn_SecurityBilling TSB On TSB.BillTxnID = TSBD.BillTxnID"
tmp = tmp & " Where Month(MonthYear)=" & mMonth & " AND Year(MonthYear)=" & CmbYear.Text & " AND TSB.SecurityAgencyID =" & mSecurityAgencyID & ""
Call Tmp1Table
strMonthTxnDetailID = ""
If TmpRs1.RecordCount > 0 Then
   For I = 1 To TmpRs1.RecordCount
       If strMonthTxnDetailID <> "" Then
          strMonthTxnDetailID = strMonthTxnDetailID & "," & TmpRs1.Fields(0)
       Else
          strMonthTxnDetailID = TmpRs1.Fields(0)
       End If
       TmpRs1.MoveNext
   Next
End If

Call Grid_Head

tmp = " SELECT ML.LocationID, LocationName, TMSLD.SecurityTypeID, SecurityShifts, TMSLD.SecurityDesignationID, "
tmp = tmp & " SecurityDesignation, NoofGodown, RatePerMonth, ME.EmployeeName +'~'+ ME.EmployeeNo as ForwardTo, "
tmp = tmp & " Me.EmployeeID, AttendanceSheetID, NoofDays, EstimatedAmt, EstimatedSTax, TotEstimatedAmt, MonthwiseTxnDetailID,TMSL.MonthwiseTxnDate"
tmp = tmp & " FROM Txn_MonthwiseSecLocDetail TMSLD"
tmp = tmp & " Inner Join Txn_MonthwiseSecLoc TMSL On TMSL.MonthwiseTxnID = TMSLD.MonthwiseTxnID"
tmp = tmp & " Inner Join Mst_SecurityShifts MSS On MSS.SecurityShiftsID = TMSLD.SecurityTypeID"
tmp = tmp & " Inner Join Mst_SecurityDesignation MSD On MSD.SecurityDesignationID = TMSLD.SecurityDesignationID"
tmp = tmp & " Left Join Mst_Employee ME On ME.EmployeeID = TMSLD.CMPMID"
tmp = tmp & " Inner Join Mst_Location ML On ML.LocationID = TMSL.LocationID"
''--- Ullhas Patil 04 Oct 2017
tmp = tmp & " Inner Join Mst_State MS On ML.StateID = MS.StateID"
''-------------------------------
tmp = tmp & " Where TMSL.SecurityAgencyID =" & mSecurityAgencyID & " And Month(TMSL.MonthYear)=" & mMonth & " AND Year(TMSL.MonthYear)=" & CmbYear.Text & ""

If strMonthTxnDetailID <> "" Then
   tmp = tmp & " AND MonthwiseTxnDetailID NOT IN (" & strMonthTxnDetailID & ")"
End If

If Gen_WcLocation <> "" Then
    tmp = tmp & " AND TMSL." & Gen_WcLocation
End If

''--- Ullhas Patil 04 Oct 2017
If CmbStateIndiaName.Text <> "" Then
    tmp = tmp & " AND MS.StateIndiaID = " & mStateIndiaID & " "
End If
''--------

Call TmpTable

If TmpRs.RecordCount > 0 Then
    MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + TmpRs.RecordCount
    For I = 2 To MSHFlexGrid2.Rows - 2
        MSHFlexGrid2.row = I
        MSHFlexGrid2.Col = 0
        MSHFlexGrid2.CellFontName = "Wingdings"
        MSHFlexGrid2.CellFontSize = 12
        MSHFlexGrid2.CellAlignment = flexAlignCenterCenter
        MSHFlexGrid2.Text = strUnChecked
        MSHFlexGrid2.TextMatrix(I, 1) = TmpRs!Locationname
        MSHFlexGrid2.TextMatrix(I, 2) = TmpRs!LocationID
        MSHFlexGrid2.TextMatrix(I, 3) = TmpRs!SecurityShifts
        MSHFlexGrid2.TextMatrix(I, 4) = TmpRs!SecurityTypeID
        MSHFlexGrid2.TextMatrix(I, 5) = TmpRs!SecurityDesignation
        MSHFlexGrid2.TextMatrix(I, 6) = TmpRs!SecurityDesignationID
        MSHFlexGrid2.TextMatrix(I, 7) = TmpRs!NoofGodown
        MSHFlexGrid2.TextMatrix(0, 7) = Val(MSHFlexGrid2.TextMatrix(0, 7)) + Val(TmpRs!NoofGodown)
        MSHFlexGrid2.TextMatrix(I, 8) = TmpRs!RatePerMonth
        MSHFlexGrid2.TextMatrix(I, 9) = IIf(IsNull(TmpRs!ForwardTo), "", TmpRs!ForwardTo)
        MSHFlexGrid2.TextMatrix(I, 10) = IIf(IsNull(TmpRs!EmployeeID), "", TmpRs!EmployeeID)
        MSHFlexGrid2.TextMatrix(I, 11) = IIf(IsNull(TmpRs!AttendanceSheetID), "", TmpRs!AttendanceSheetID)
        MSHFlexGrid2.TextMatrix(I, 12) = TmpRs!NoOfDays
        MSHFlexGrid2.TextMatrix(I, 13) = TmpRs!EstimatedAmt
        MSHFlexGrid2.TextMatrix(I, 14) = IIf(IsNull(TmpRs!EstimatedSTax), "", TmpRs!EstimatedSTax)
        MSHFlexGrid2.TextMatrix(I, 15) = IIf(IsNull(TmpRs!TotEstimatedAmt), "", TmpRs!TotEstimatedAmt)
        MSHFlexGrid2.TextMatrix(I, 16) = ""
        MSHFlexGrid2.TextMatrix(I, 17) = ""
        MSHFlexGrid2.TextMatrix(I, 18) = ""
        MSHFlexGrid2.TextMatrix(I, 19) = TmpRs!MonthwiseTxnDetailID
        MSHFlexGrid2.TextMatrix(I, 20) = TmpRs!MonthwiseTxnDate
        TmpRs.MoveNext
    Next
End If
End Sub

Private Function GetMaxTxnID() As Double
Dim Retval As Double
tmp = "Select max(BillTxnID) from Txn_SecurityBilling "
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetMaxTxnID = Retval
End Function

Function GetMaxTxnDetails() As Double
Dim Retval As Double
Retval = 0
tmp = "Select max(BillTxnDetailID) From Txn_SecurityBillingDetail "
Call TmpTable
If TmpRs.RecordCount > 0 Then
   Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
End If
Retval = Retval + 1
TmpRs.Close
GetMaxTxnDetails = Retval

End Function

Private Sub SaveCmd_Click()
Dim blnSaveDetails  As Boolean
Dim dtMonthYear     As Date

If MSHFlexGrid2.Rows <= 3 Then
   MsgBox "No details found in grid!!!"
   Exit Sub
End If

blnSaveDetails = False
For I = 2 To MSHFlexGrid2.Rows - 1
    If Trim(MSHFlexGrid2.TextMatrix(I, 0)) = strChecked Then
       blnSaveDetails = True
       Exit For
    End If
Next

If blnSaveDetails = False Then
    MsgBox "Atleast update Bill amount for one row in grid!!!"
    Exit Sub
End If

''-- Ullhas Patil 04 Oct 2017
If CmbStateIndiaName.Text = "" Then
   MsgBox "Blank State India not allowed!!!"
   CmbStateIndiaName.SetFocus
   Exit Sub
End If
''-- Ullhas Patil 25 Oct 2017
If CmbBillingState.Text = "" Then
   MsgBox "Blank Billing State not allowed!!!"
   CmbBillingState.SetFocus
   Exit Sub
End If
''---------------------------

If Trim(TxtBillNo.Text) = "" Then
   MsgBox "Blank Bill No not allowed!!!"
   TxtBillNo.SetFocus
   Exit Sub
End If

If Trim(TxtBillAmt.Text) = "" Then
   MsgBox "Blank Bill Amount not allowed!!!"
   TxtBillAmt.SetFocus
   Exit Sub
End If

If Trim(TxtSTax.Text) = "" Then
   MsgBox "Blank S.Tax not allowed!!!"
   TxtSTax.SetFocus
   Exit Sub
End If

If Trim(TxtBillNo.Text) = "" Then
   MsgBox "Blank bill No not allowed!!!"
   TxtBillNo.SetFocus
   Exit Sub
End If

If Trim(CmbSTaxDesc.Text) = "" Then
   MsgBox "Blank Service Tax Deacription not allowed!!!"
   CmbSTaxDesc.SetFocus
   Exit Sub
End If

If Trim(CmbForwardTo.Text) = "" Then
   MsgBox "Blank Forward To not allowed!!!"
   CmbForwardTo.SetFocus
   Exit Sub
End If

If TxtBillDate > Date Then
   MsgBox "Bill Date not be future Date !!! "
   TxtBillDate.SetFocus
   Exit Sub
End If

If CmbMonth.Text = "" Then
   MsgBox "Blank Month not allowed!!!"
   CmbMonth.SetFocus
   Exit Sub
End If
If CmbYear.Text = "" Then
   MsgBox "Blank Year not allowed!!!"
   CmbYear.SetFocus
   Exit Sub
End If

dtMonthYear = Left(CmbMonth, 3) & "-01-" & CmbYear
 
If dtMonthYear > Date Then
   MsgBox "Future Month not allowed!!"
   CmbMonth.SetFocus
   Exit Sub
End If

If LCase(CmbStatus.Text) = "settled" Then
   If Len(TxtFAGRemark) > 250 Then
      MsgBox "Length of disbursment remark is more than 250 characters , please check!!!"
      TxtFAGRemark.SetFocus
      Exit Sub
   End If
End If


If Edit_Flg = "A" Then
   If Format(Gen_BlockingDate, "yyyy-mm") >= Format(CDate(dtMonthYear), "yyyy-mm") Then
      MsgBox "Select next month from " & Format(Gen_BlockingDate, "MMM-yyyy") & "!!! "
      Exit Sub
   End If
End If
If CheckDuplicateBill() = True Then
   MsgBox " Duplicate Bill No & Bill Date not allowed for selected Security Agency!!"
   TxtBillNo.SetFocus
   Exit Sub
End If

If Edit_Flg = "E" Then
   If LCase(mStatus) = "a" Then
      If LCase(Gen_UserRole) = "fag" Or LCase(Gen_UserRole) = "power" Or LCase(Gen_UserRole) = "fagadmin" Then
         If CmbStatus.Text = "" Then
            MsgBox "Blank Status not allowed!!!"
            CmbStatus.SetFocus
            Exit Sub
         End If
         If LCase(CmbStatus.Text) = "approved" Or LCase(CmbStatus.Text) = "rejected" Then
            If TxtFAGRemarks = "" Then
               MsgBox "Blank FAG Remarks not allowed!!!"
               TxtFAGRemarks.SetFocus
               Exit Sub
            End If
            If TxtApprovalDate.Value > Date Then
               MsgBox "Future NAV Booked date not allowed!!!"
               TxtApprovalDate.SetFocus
               Exit Sub
            End If
         End If
         If Format(Gen_BlockingDate, "yyyy-mm") >= Format(TxtApprovalDate.Value, "yyyy-mm") Then
            MsgBox "NAV Booked date allowed for next month of " & Format(Gen_BlockingDate, "MMM-yyyy") & "!!! "
            TxtApprovalDate.SetFocus
            Exit Sub
         End If
      End If
   ElseIf LCase(mStatus) = "s" Then
      If Val(TxtFAGDisbAmt.Text) > 0 Then
         If TxtFagDisbDate.Value > Date Then
            MsgBox "Future date not allowed!!"
            TxtFagDisbDate.SetFocus
            SSTab1.Tab = 1
            Exit Sub
         End If
         If TxtFagDisbDate.Value < TxtTxnDate.Value Then
            MsgBox "Disbursement Date should not less than Txn Date!!"
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
            SSTab1.Tab = 1
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
         If Val(TxtFAGDisbAmt) < Val(TxtTotBillAmt) Then
            If Trim(TxtDeductionRemark) = "" Then
               MsgBox "As disbursment amount is less than total bill amount please provide deduction remarks!!!"
               SSTab1.Tab = 1
               TxtDeductionRemark.SetFocus
               Exit Sub
            End If
         End If
     End If
   End If
End If

If TxtRemarks = "" Then
   MsgBox "Blank remarks not allowed!!!"
   TxtRemarks.SetFocus
   Exit Sub
End If

If Val(TxtBillAmt.Text) <> Val(MSHFlexGrid2.TextMatrix(0, 16)) Then
   MsgBox "Mismatch Bill Amount!!"
   MSHFlexGrid2.Col = 16
   MSHFlexGrid2.row = 2
   MSHFlexGrid2.SetFocus
   Exit Sub
End If

If Val(TxtSTax.Text) <> Val(MSHFlexGrid2.TextMatrix(0, 17)) Then
   MsgBox "Mismatch Service Tax!!"
   MSHFlexGrid2.Col = 17
   MSHFlexGrid2.row = 2
   MSHFlexGrid2.SetFocus
   Exit Sub
End If



If Edit_Flg = "A" Then
   RsTxn_SecurityBilling.AddNew
'   RsTxn_SecurityBilling!BillTxnID = GetMaxTxnID
   RsTxn_SecurityBilling!G_UID = GetGUID
'   TxtTxnID = RsTxn_SecurityBilling!BillTxnID
   TxtG_UID = RsTxn_SecurityBilling!G_UID
   RsTxn_SecurityBilling!BatchID = 0
End If

RsTxn_SecurityBilling!BillTxnDate = TxtTxnDate.Value
RsTxn_SecurityBilling!MonthYear = dtMonthYear
RsTxn_SecurityBilling!SecurityAgencyID = mSecurityAgencyID
RsTxn_SecurityBilling!G_UID = TxtG_UID
RsTxn_SecurityBilling!ServiceTaxID = mServiceTaxID
RsTxn_SecurityBilling!BillNo = TxtBillNo
RsTxn_SecurityBilling!BillDate = TxtBillDate
RsTxn_SecurityBilling!BillAmt = TxtBillAmt
RsTxn_SecurityBilling!BillSTax = TxtSTax
RsTxn_SecurityBilling!TotalAmt = Val(TxtTotBillAmt)
RsTxn_SecurityBilling!ForwardTo = mForwardToID
'RsTxn_SecurityBilling!FAGRemarks = Trim(TxtFAGRemarks)
RsTxn_SecurityBilling!Status = Left(Trim(CmbStatus), 1)

RsTxn_SecurityBilling!Remarks = TxtRemarks

'If TxtApprovalRemarks.Locked = False Then
'   RsTxnClaimSubmission!ApprovalRemarks = TxtApprovalRemarks
'End If
'If TxtApprovalDate.Visible = True Then
'   RsTxnClaimSubmission!ApprovalDate = TxtApprovalDate.Value
'End If

If Edit_Flg = "E" Then
   If LCase(CmbStatus.Text) = "approved" Or LCase(CmbStatus.Text) = "rejected" Then
      TxtFAGRemarks = Replace(TxtFAGRemarks, ",", "-")
      RsTxn_SecurityBilling!FAGRemarks = Trim(TxtFAGRemarks)
      RsTxn_SecurityBilling!ApprovedDate = TxtApprovalDate.Value
   End If
End If

If LCase(CmbStatus.Text) = "settled" Then
   RsTxn_SecurityBilling!DisbuDate = TxtFagDisbDate.Value
   RsTxn_SecurityBilling!DisbuRemark = TxtFAGRemark.Text
   RsTxn_SecurityBilling!DisbuAmount = TxtFAGDisbAmt.Text
   RsTxn_SecurityBilling!ModeOfPay = CmbModeOfPay.Text
   RsTxn_SecurityBilling!FAGBankID = mFAGBankID
   RsTxn_SecurityBilling!DeductionRemark = Trim(TxtDeductionRemark)
Else
   RsTxn_SecurityBilling!DisbuDate = Null
   RsTxn_SecurityBilling!DisbuRemark = Null
   RsTxn_SecurityBilling!DisbuAmount = Null
   RsTxn_SecurityBilling!ModeOfPay = Null
   RsTxn_SecurityBilling!FAGBankID = Null
   RsTxn_SecurityBilling!DeductionRemark = Null
End If

RsTxn_SecurityBilling!StateIndiaID = mStateIndiaID          ''-- Ullhas Patil 04 Oct 2017
RsTxn_SecurityBilling!BillingStateIndiaID = mBillingStateID      ''-- Ullhas Patil 25 Oct 2017

If IsNull(RsTxn_SecurityBilling!CreatedBy) = True Or RsTxn_SecurityBilling!CreatedBy = "" Then
   RsTxn_SecurityBilling!CreatedBy = Gen_UserLoginID
   RsTxn_SecurityBilling!CreatedDate = Now
Else
   RsTxn_SecurityBilling!UpdatedBy = Gen_UserLoginID
   RsTxn_SecurityBilling!UpdatedDate = Now
End If
RsTxn_SecurityBilling.Update
TxtCreated = IIf(IsNull(RsTxn_SecurityBilling!CreatedBy), "", RsTxn_SecurityBilling!CreatedBy) & " :- " & IIf(IsNull(RsTxn_SecurityBilling!CreatedDate), "", RsTxn_SecurityBilling!CreatedDate)
TxtUpdated = IIf(IsNull(RsTxn_SecurityBilling!UpdatedBy), "", RsTxn_SecurityBilling!UpdatedBy) & " :- " & IIf(IsNull(RsTxn_SecurityBilling!UpdatedDate), "", RsTxn_SecurityBilling!UpdatedDate)
RsTxn_SecurityBilling.Requery
RsTxn_SecurityBilling.MoveFirst
RsTxn_SecurityBilling.Find "G_UID = '" & TxtG_UID & "'"

tmp = "Select BillTxnID From Txn_SecurityBilling Where G_UID = '" & TxtG_UID & "'"
Call TmpTable
If TmpRs.RecordCount > 0 Then
'   mCxTF_DetailsID = TmpRs!CxTF_DetailsID
   TxtTxnID = TmpRs!BillTxnID
End If


Call SaveSecurityBillingDetails
Call Indata
End Sub

Private Sub ShowAccountDetails(mAgencyID_ As Variant)
tmp = "Select BankID ,BranchID,BankAcNo,IFSCCode from Mst_SecurityAgency Where SecurityAgencyID=" & mAgencyID_
Call TmpTable
If TmpRs.RecordCount > 0 Then
   TxtBankName.Text = GetBankName(IIf(IsNull(TmpRs!BankID), 0, TmpRs!BankID), "N")
   TxtBranchName = GetBranchName(IIf(IsNull(TmpRs!BranchID), 0, TmpRs!BranchID), "N")  ' IIf(IsNull(TmpRs!BankBranch), "", TmpRs!BankBranch)
   TxtACNo = IIf(IsNull(TmpRs!BankAcNo), "", TmpRs!BankAcNo)
   TxtIFSCCode = IIf(IsNull(TmpRs!IFSCCode), "", TmpRs!IFSCCode)
End If
End Sub
Private Sub DeleteSecurityBillingDetail()
tmp = "Delete From Txn_SecurityBillingDetail Where BillTxnID=" & TxtTxnID & " "
Call Tmp4Table
End Sub

Private Sub SaveSecurityBillingDetails()

Call DeleteSecurityBillingDetail

Call TableTxn_SecurityBillingDetail '("Where BillTxnID=" & TxtTxnID & "")

For I = 2 To MSHFlexGrid2.Rows - 1
    If MSHFlexGrid2.TextMatrix(I, 0) = strChecked Then
       'RsTxn_SecurityBillingDetail.Find "BillTxnDetailID = " & Val(MSHFlexGrid2.TextMatrix(I, 17))
       'If RsTxn_SecurityBillingDetail.EOF Then
          RsTxn_SecurityBillingDetail.AddNew
'          RsTxn_SecurityBillingDetail!BillTxnDetailID = GetMaxTxnDetails
       'Else
       '   RsTxn_SecurityBillingDetail!BillTxnDetailID = MSHFlexGrid2.TextMatrix(I, 16)
       'End If
       RsTxn_SecurityBillingDetail!BillTxnID = TxtTxnID
       RsTxn_SecurityBillingDetail!LocationID = MSHFlexGrid2.TextMatrix(I, 2)
       RsTxn_SecurityBillingDetail!SecurityTypeID = MSHFlexGrid2.TextMatrix(I, 4)
       RsTxn_SecurityBillingDetail!SecurityDesignationID = MSHFlexGrid2.TextMatrix(I, 6)
       RsTxn_SecurityBillingDetail!NoofGodown = Val(MSHFlexGrid2.TextMatrix(I, 7))
       RsTxn_SecurityBillingDetail!RatePerMonth = Val(MSHFlexGrid2.TextMatrix(I, 8))
       If MSHFlexGrid2.TextMatrix(I, 10) <> "" Then
          RsTxn_SecurityBillingDetail!CMPMID = Val(MSHFlexGrid2.TextMatrix(I, 10))
       End If
       RsTxn_SecurityBillingDetail!AttendanceSheetID = Trim(MSHFlexGrid2.TextMatrix(I, 11))
       RsTxn_SecurityBillingDetail!NoOfDays = Val(MSHFlexGrid2.TextMatrix(I, 12))
       RsTxn_SecurityBillingDetail!EstimatedAmt = Val(MSHFlexGrid2.TextMatrix(I, 13))
       RsTxn_SecurityBillingDetail!EstimatedSTax = Val(MSHFlexGrid2.TextMatrix(I, 14))
       RsTxn_SecurityBillingDetail!TotEstimatedAmt = Val(MSHFlexGrid2.TextMatrix(I, 15))
       RsTxn_SecurityBillingDetail!BillDetailAmt = Val(MSHFlexGrid2.TextMatrix(I, 16))
       RsTxn_SecurityBillingDetail!BillDetailSTax = Val(MSHFlexGrid2.TextMatrix(I, 17))
       RsTxn_SecurityBillingDetail!BillNoOfDays = Val(MSHFlexGrid2.TextMatrix(I, 18))
       RsTxn_SecurityBillingDetail!MonthwiseTxnDetailID = Val(MSHFlexGrid2.TextMatrix(I, 19))
       RsTxn_SecurityBillingDetail!CreatedBy = Gen_UserLoginID
       RsTxn_SecurityBillingDetail!CreatedDate = Now
       RsTxn_SecurityBillingDetail.Update
       
       tmp = "Select BillTxnDetailID From Txn_SecurityBillingDetail Where BillTxnID =  " & TxtTxnID & " And MonthwiseTxnDetailID = " & Val(MSHFlexGrid2.TextMatrix(I, 19)) & " "
       Call TmpTable
        
       If TmpRs.RecordCount > 0 Then
          mBillTxnDetailID = TmpRs!BillTxnDetailID
       End If
       
       Call UpdateMonthWiseDetails(Val(MSHFlexGrid2.TextMatrix(I, 19)), "S", I, mBillTxnDetailID)
       
       RsTxn_SecurityBillingDetail.Requery
       
    End If
Next
End Sub
Private Sub UpdateMonthWiseDetails(MonthwiseTxnDetailID_ As Variant, Mod_ As Variant, x As Variant, BillingID_ As Variant)

If LCase(Mod_) = "s" Then
 
   tmp = "UPDATE Txn_MonthwiseSecLocDetail SET BillingNoOfDays = " & Val(MSHFlexGrid2.TextMatrix(x, 18)) & ","
   tmp = tmp & " BillingAmount=" & Val(MSHFlexGrid2.TextMatrix(x, 16)) & ", BillingSTax=" & Val(MSHFlexGrid2.TextMatrix(x, 17)) & ","
   tmp = tmp & " BillingTotalAmount= " & Val(MSHFlexGrid2.TextMatrix(x, 16)) & " + " & Val(MSHFlexGrid2.TextMatrix(x, 17)) & ","
   tmp = tmp & " BillingID = " & BillingID_ & " Where MonthwiseTxnDetailID=" & Val(MonthwiseTxnDetailID_) & " "
   
   Call Tmp4Table
ElseIf LCase(Mod_) = "d" Then
   mMonWiseID = ""
   For x = 2 To MSHFlexGrid2.Rows - 1
       If x = 2 Then
          mMonWiseID = "(" & Val(MSHFlexGrid2.TextMatrix(x, 19))
       Else
          mMonWiseID = mMonWiseID & "," & Val(MSHFlexGrid2.TextMatrix(x, 19))
       End If
       If x = MSHFlexGrid2.Rows - 1 Then
          mMonWiseID = mMonWiseID & ")"
       End If
   Next

   tmp = "UPDATE Txn_MonthwiseSecLocDetail SET BillingNoOfDays = 0,"
   tmp = tmp & " BillingAmount=0, BillingSTax=0, "
   tmp = tmp & " BillingTotalAmount=0,BillingID = Null "
   tmp = tmp & " Where MonthwiseTxnDetailID in " & mMonWiseID '= " & Val(MSHFlexGrid2.TextMatrix(I, 18)) & " "
   
   Call Tmp4Table

End If

End Sub

Private Sub CalcInGrid()
MSHFlexGrid2.TextMatrix(0, 16) = 0
MSHFlexGrid2.TextMatrix(0, 17) = 0
For I = 2 To MSHFlexGrid2.Rows - 1
    MSHFlexGrid2.TextMatrix(0, 16) = Val(MSHFlexGrid2.TextMatrix(0, 16)) + Val(MSHFlexGrid2.TextMatrix(I, 16))
    MSHFlexGrid2.TextMatrix(0, 17) = Val(MSHFlexGrid2.TextMatrix(0, 17)) + Val(MSHFlexGrid2.TextMatrix(I, 17))
Next
End Sub

Private Sub CalcEstimateAmt()
Dim dblEstimateAmt      As Double
intMonthDays = 0
If CmbYear.Text <> "" Then
    intMonthDays = Day(DateSerial(CmbYear.Text, mMonth + 1, 0))
    dblEstimateAmt = Val(MSHFlexGrid2.TextMatrix(mSelRow, 8)) * Val(MSHFlexGrid2.TextMatrix(mSelRow, 18))
    dblEstimateAmt = Format(dblEstimateAmt / intMonthDays, "########0.00")
    MSHFlexGrid2.TextMatrix(mSelRow, 16) = Format(dblEstimateAmt, "#.00")
End If
'MSHFlexGrid2.TextMatrix(mSelRow, 15) = Val(MSHFlexGrid2.TextMatrix(mSelRow, 13)) + Val(MSHFlexGrid2.TextMatrix(mSelRow, 14))
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


Private Sub CmdPrint_Click()
Gen_ExcelPassword = "1234"
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "SecurityClaim" & Trim(TxtTxnID) & ".xls")
FileName = Gen_mTimeStamp & Gen_UserName & "SecurityClaim" & Trim(TxtTxnID) & ".xls"
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
    oWS.Range("c:c").ColumnWidth = 13.57 '17.14
    oWS.Range("d:d").ColumnWidth = 8.29 '12.86
    oWS.Range("e:e").ColumnWidth = 8.57 '33 '32.14
    oWS.Range("f:f").ColumnWidth = 8
    oWS.Range("g:g").ColumnWidth = 7.86
    oWS.Range("h:h").ColumnWidth = 7
    oWS.Range("i:i").ColumnWidth = 7.71
      
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
    oWS.Cells(mRow, C + 1) = "SECURITY BILL VOUCHER"
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
    oWS.Cells(mRow, C + 7) = TxtTxnID
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
    oWS.Cells(mRow, C + 9) = Format(TxtTxnDate.Value, "dd-mmm-yyyy")
    oWS.Cells(mRow, C + 9).Font.Bold = True
    oWS.Cells(mRow, C + 9).Font.Size = 10
       
    mRow = mRow + 1
    tmp = "a" & mRow & ":" & "b" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 1) = "Security Agency Name :"
    oWS.Cells(mRow, C + 1).Font.Bold = True
    oWS.Cells(mRow, C + 1).Font.Size = 10
    'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
    
    tmp = "c" & mRow & ":" & "i" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 3) = TxtSecurityAgency
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
    oWS.Cells(mRow, C + 3) = "'" & CmbMonth.Text & "-" & CmbYear.Text
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
    
    tmp = "c" & mRow & ":" & "i" & mRow
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
    oWS.Cells(mRow, C + 7) = Format(TxtBillDate.Value, "dd-MMM-yyyy")
    oWS.Cells(mRow, C + 7).Font.Bold = True
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
    oWS.Cells(mRow, C + 1) = "Location Name"
    oWS.Cells(mRow, C + 1).Font.Bold = True
    oWS.Cells(mRow, C + 1).Font.Size = 10
    oWS.Cells(mRow, C + 1).RowHeight = 27.75
    oWS.Cells(mRow, C + 1).WrapText = True
    'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
    
    tmp = "b" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 2) = "Security Attendance ID"
    oWS.Cells(mRow, C + 2).Font.Bold = True
    oWS.Cells(mRow, C + 2).Font.Size = 10
    oWS.Cells(mRow, C + 2).RowHeight = 27.75
    oWS.Cells(mRow, C + 2).WrapText = True
    'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
    
    tmp = "c" & mRow & ":" & "c" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 3) = "Attendance Date"
    oWS.Cells(mRow, C + 3).Font.Bold = True
    oWS.Cells(mRow, C + 3).Font.Size = 10
    oWS.Cells(mRow, C + 3).RowHeight = 27.75
    'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
    
    tmp = "d" & mRow & ":" & "e" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 4) = "Expense Type"
    oWS.Cells(mRow, C + 4).Font.Bold = True
    oWS.Cells(mRow, C + 4).Font.Size = 10
    oWS.Cells(mRow, C + 4).RowHeight = 27.75
    'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
    
    tmp = "f" & mRow & ":" & "g" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 6) = "Particulars"
    oWS.Cells(mRow, C + 6).Font.Bold = True
    oWS.Cells(mRow, C + 6).Font.Size = 10
    oWS.Cells(mRow, C + 6).RowHeight = 27.75
    'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
    
    tmp = "h" & mRow & ":" & "h" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 8) = "Billing Days"
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
    
      
tmp = "Select TM.MonthwiseTxnID,TM.MonthYear,TMD.Remarks ,TMD.MonthwiseTxnDetailID"
tmp = tmp & " from Txn_MonthwiseSecLoc Tm"
tmp = tmp & " Inner join Txn_MonthwiseSecLocDetail TMD on TM.MonthwiseTxnID=TMD.MonthwiseTxnID"
tmp = tmp & " Where MonthwiseTxnDetailID in ( select MonthwiseTxnDetailID from Txn_SecurityBillingDetail Where BillTxnID=" & Val(TxtTxnID) & ")"

Call Tmp2Table
    
    For I = 2 To MSHFlexGrid2.Rows - 2
        mRow = mRow + 1 ' + 6
               C = 0
               
               oWS.Cells(mRow, C + 1) = MSHFlexGrid2.TextMatrix(I, 1)
               
               tmp = "c" & mRow & ":" & "c" & mRow
               oWS.Range(tmp).Merge
               
               tmp = "d" & mRow & ":" & "e" & mRow
               oWS.Range(tmp).Merge
               
               tmp = "f" & mRow & ":" & "g" & mRow
               oWS.Range(tmp).Merge
               
               oWS.Cells(mRow, C + 4) = "Security Expense"
               
               oWS.Cells(mRow, C + 8) = MSHFlexGrid2.TextMatrix(I, 18)
               
               oWS.Cells(mRow, C + 9) = MSHFlexGrid2.TextMatrix(I, 16)
               
               If TmpRs2.RecordCount > 0 Then TmpRs2.MoveFirst
               TmpRs2.Find "MonthwiseTxnDetailID=" & Val(MSHFlexGrid2.TextMatrix(I, 19))
               If TmpRs2.EOF = False Then
                  oWS.Cells(mRow, C + 2) = TmpRs2!MonthWiseTxnID
                  oWS.Cells(mRow, C + 3) = "'" & Format(TmpRs2!MonthYear, "MMM-yyyy")
                  oWS.Cells(mRow, C + 6) = TmpRs2!Remarks
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
    oWS.Cells(mRow, C + 9) = Format(MSHFlexGrid2.TextMatrix(0, 16), "#####0.00")
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
    oWS.Cells(mRow, C + 9) = Format(TxtSTax, "#####0.00")
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
    oWS.Cells(mRow, C + 9) = Format(MSHFlexGrid2.TextMatrix(0, 16) + Val(TxtSTax), "#####0.00")
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
    oWS.Cells(mRow, C + 1) = "SECURITY BILL VOUCHER"
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
    oWS.Cells(mRow, C + 8) = TxtTxnID
    oWS.Cells(mRow, C + 8).Font.Bold = True
    oWS.Cells(mRow, C + 8).Font.Size = 10
    'oWS.Cells(mRow, c + j).Font.Size = 10
       
    mRow = mRow + 1
    tmp = "a" & mRow & ":" & "b" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 1) = "Security Agency Name :"
    oWS.Cells(mRow, C + 1).Font.Bold = True
    oWS.Cells(mRow, C + 1).Font.Size = 10
    'oWS.Cells(mRow, c + j).Font.Size = 10
    
    tmp = "c" & mRow & ":" & "i" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 3) = TxtSecurityAgency
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
    oWS.Cells(mRow, C + 3) = "'" & CmbMonth.Text & "-" & CmbYear.Text
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

Private Function CheckDuplicateBill() As Boolean
Dim Retval As Boolean
Retval = False
tmp = " Select * from Txn_SecurityBilling Where BillNo='" & TxtBillNo.Text & "' and BillDate='" & TxtBillDate.Value & "' And Status Not in ('R') And SecurityAgencyID=" & mSecurityAgencyID
If Edit_Flg = "E" Then
   tmp = tmp & " And BillTxnID<>" & Val(TxtTxnID) & ""
End If

Call TmpTable
If TmpRs.RecordCount > 0 Then
   Retval = True
End If
CheckDuplicateBill = Retval
End Function
