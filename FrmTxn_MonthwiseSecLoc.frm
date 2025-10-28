VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmTxn_MonthwiseSecLoc 
   Caption         =   "Monthly Security Expense Accrual"
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
      TabIndex        =   51
      Top             =   0
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      TabIndex        =   50
      Top             =   10320
      Width           =   10935
   End
   Begin VB.Frame Frame4 
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   9780
      Left            =   0
      TabIndex        =   28
      Top             =   -120
      Width           =   15180
      Begin VB.Frame Frame0 
         Height          =   1695
         Left            =   120
         TabIndex        =   41
         Top             =   4560
         Visible         =   0   'False
         Width           =   14370
         Begin VB.OptionButton OptMonthYear 
            Caption         =   "Month and Year"
            Height          =   255
            Left            =   75
            TabIndex        =   86
            Top             =   195
            Width           =   1680
         End
         Begin VB.OptionButton OptTransactionDate 
            Caption         =   "Transaction Date"
            Height          =   255
            Left            =   75
            TabIndex        =   85
            Top             =   450
            Width           =   1650
         End
         Begin VB.ComboBox CmbSSecurityAgency 
            Height          =   315
            Left            =   4545
            Sorted          =   -1  'True
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   330
            Width           =   3990
         End
         Begin VB.CommandButton CmdGo 
            Caption         =   "Go"
            Height          =   350
            Left            =   12045
            TabIndex        =   26
            Top             =   327
            Width           =   1000
         End
         Begin VB.CommandButton CmdExcel 
            Caption         =   "Excel"
            Height          =   350
            Left            =   13050
            TabIndex        =   27
            Top             =   327
            Width           =   1000
         End
         Begin VB.ComboBox CmbSLocation 
            Height          =   315
            Left            =   8565
            Sorted          =   -1  'True
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   330
            Width           =   3435
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   1455
            Left            =   75
            TabIndex        =   42
            Top             =   1170
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
            Left            =   3195
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   330
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
            Format          =   33947649
            CurrentDate     =   36494
         End
         Begin MSComCtl2.DTPicker SFrom 
            Height          =   300
            Left            =   1860
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   330
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
            Format          =   33947649
            CurrentDate     =   36494
         End
         Begin VB.Label Label44 
            Caption         =   "From Date"
            Height          =   255
            Left            =   2160
            TabIndex        =   48
            Top             =   120
            Width           =   855
         End
         Begin VB.Label Label37 
            Caption         =   "Security Agency"
            Height          =   225
            Left            =   5940
            TabIndex        =   47
            Top             =   135
            Width           =   1215
         End
         Begin VB.Label Label45 
            Caption         =   "To Date"
            Height          =   255
            Left            =   3480
            TabIndex        =   46
            Top             =   120
            Width           =   855
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
            Left            =   12765
            TabIndex        =   45
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
            Left            =   10245
            TabIndex        =   44
            Top             =   750
            Width           =   2415
         End
         Begin VB.Label Label31 
            Caption         =   "Location"
            Height          =   225
            Left            =   9945
            TabIndex        =   43
            Top             =   135
            Width           =   675
         End
      End
      Begin VB.Frame Frame2 
         Height          =   2535
         Left            =   120
         TabIndex        =   52
         Top             =   2175
         Width           =   14880
         Begin VB.TextBox TxtEstimatedSTax 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   10755
            TabIndex        =   8
            Top             =   1515
            Width           =   1200
         End
         Begin VB.TextBox TxtSalaryPaidUpto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   375
            Left            =   12000
            TabIndex        =   82
            TabStop         =   0   'False
            Top             =   630
            Width           =   1575
         End
         Begin VB.CommandButton CmdCMPM 
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
            Left            =   5640
            TabIndex        =   78
            TabStop         =   0   'False
            Top             =   1530
            Width           =   375
         End
         Begin VB.ComboBox CmbCMPM 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1350
            Sorted          =   -1  'True
            TabIndex        =   5
            Top             =   1515
            Width           =   4260
         End
         Begin VB.TextBox TxtEndDate 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   375
            Left            =   10755
            Locked          =   -1  'True
            TabIndex        =   69
            TabStop         =   0   'False
            Top             =   630
            Width           =   1200
         End
         Begin VB.TextBox TxtStartDate 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   375
            Left            =   9495
            Locked          =   -1  'True
            TabIndex        =   68
            TabStop         =   0   'False
            Top             =   630
            Width           =   1215
         End
         Begin VB.TextBox TxtSecurityDesignation 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   375
            Left            =   4800
            Locked          =   -1  'True
            TabIndex        =   67
            TabStop         =   0   'False
            Top             =   630
            Width           =   4650
         End
         Begin VB.TextBox TxtSecurityType 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   375
            Left            =   90
            Locked          =   -1  'True
            TabIndex        =   66
            TabStop         =   0   'False
            Top             =   630
            Width           =   4665
         End
         Begin VB.TextBox TxtRatePerMonth 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   375
            Left            =   13620
            Locked          =   -1  'True
            TabIndex        =   65
            TabStop         =   0   'False
            Top             =   630
            Width           =   1215
         End
         Begin VB.TextBox TxtNoofGodown 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   375
            Left            =   90
            TabIndex        =   64
            TabStop         =   0   'False
            Top             =   1508
            Width           =   1215
         End
         Begin VB.TextBox TxtTotEstimatedAmt 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   375
            Left            =   12000
            Locked          =   -1  'True
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   1515
            Width           =   1665
         End
         Begin VB.TextBox TxtEstimatedAmt 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   375
            Left            =   9240
            Locked          =   -1  'True
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   1515
            Width           =   1470
         End
         Begin VB.TextBox TxtNoofDays 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   7650
            TabIndex        =   7
            Top             =   1515
            Width           =   1530
         End
         Begin VB.TextBox TxtAttenSheetID 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   6060
            TabIndex        =   6
            Top             =   1515
            Width           =   1530
         End
         Begin VB.CommandButton CmdAdd2List 
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
            Left            =   9270
            TabIndex        =   12
            Top             =   2010
            Width           =   1395
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
            Height          =   465
            Left            =   1350
            MaxLength       =   1000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   11
            Top             =   1965
            Width           =   7860
         End
         Begin VB.Label Label20 
            Alignment       =   2  'Center
            Caption         =   "Estimated S.Tax"
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
            Left            =   10935
            TabIndex        =   84
            Top             =   1028
            Width           =   900
         End
         Begin VB.Label Label19 
            Alignment       =   2  'Center
            Caption         =   "Salary Payable Upto"
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
            Left            =   12120
            TabIndex        =   83
            Top             =   135
            Width           =   1335
         End
         Begin VB.Label Label16 
            Caption         =   "Name of employee approving the attendance"
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
            Left            =   1380
            TabIndex        =   77
            Top             =   1185
            Width           =   4080
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            Caption         =   "Total Estimated Amount"
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
            Left            =   12120
            TabIndex        =   63
            Top             =   1020
            Width           =   1425
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            Caption         =   "Estimated Amount"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   9570
            TabIndex        =   62
            Top             =   1035
            Width           =   885
         End
         Begin VB.Label Label9 
            Caption         =   "No of Days"
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
            Left            =   7800
            TabIndex        =   61
            Top             =   1110
            Width           =   1215
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            Caption         =   "Attendance sheet ID"
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
            Left            =   6210
            TabIndex        =   60
            Top             =   1035
            Width           =   1230
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Caption         =   "No of Godown"
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
            Left            =   240
            TabIndex        =   59
            Top             =   1035
            Width           =   735
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "Rate per month"
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
            Left            =   13830
            TabIndex        =   58
            Top             =   135
            Width           =   780
         End
         Begin VB.Label Label1 
            Caption         =   "End Date"
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
            Left            =   10860
            TabIndex        =   57
            Top             =   255
            Width           =   990
         End
         Begin VB.Label Label8 
            Caption         =   "Start Date"
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
            Left            =   9570
            TabIndex        =   56
            Top             =   255
            Width           =   1065
         End
         Begin VB.Label Lbl 
            Caption         =   "Security Type"
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
            Left            =   1785
            TabIndex        =   55
            Top             =   240
            Width           =   1275
         End
         Begin VB.Label Label5 
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
            Left            =   105
            TabIndex        =   54
            Top             =   2070
            Width           =   1140
         End
         Begin VB.Label Label3 
            Caption         =   "Security Designation"
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
            Left            =   6015
            TabIndex        =   53
            Top             =   240
            Width           =   1890
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2040
         Left            =   120
         TabIndex        =   29
         Top             =   75
         Width           =   14880
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
            TabIndex        =   3
            Top             =   660
            Width           =   10455
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
            TabIndex        =   76
            TabStop         =   0   'False
            Top             =   660
            Width           =   10455
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
            Left            =   1245
            Locked          =   -1  'True
            TabIndex        =   73
            TabStop         =   0   'False
            Top             =   1560
            Width           =   4590
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
            Left            =   7095
            Locked          =   -1  'True
            TabIndex        =   72
            TabStop         =   0   'False
            Top             =   1545
            Width           =   4590
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
            Left            =   8325
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
            Left            =   6075
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
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   210
            Width           =   1455
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
            Left            =   11745
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   705
            Width           =   375
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
            Height          =   300
            Left            =   11760
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   1125
            Width           =   375
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
            Left            =   1245
            Sorted          =   -1  'True
            TabIndex        =   4
            Top             =   1095
            Width           =   10455
         End
         Begin MSComCtl2.DTPicker TxtTxnDate 
            Height          =   375
            Left            =   3900
            TabIndex        =   34
            Top             =   210
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   33947649
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
            Left            =   1245
            Locked          =   -1  'True
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   1095
            Width           =   10455
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
            Left            =   8325
            Locked          =   -1  'True
            TabIndex        =   81
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
            Left            =   6075
            Locked          =   -1  'True
            TabIndex        =   80
            TabStop         =   0   'False
            Top             =   217
            Width           =   1455
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
            Left            =   105
            TabIndex        =   75
            Top             =   1620
            Width           =   945
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
            Left            =   6135
            TabIndex        =   74
            Top             =   1620
            Width           =   915
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
            Left            =   5385
            TabIndex        =   71
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
            Left            =   7740
            TabIndex        =   70
            Top             =   255
            Width           =   555
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
            TabIndex        =   40
            Top             =   240
            Width           =   900
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
            TabIndex        =   39
            Top             =   600
            Width           =   1140
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
            Left            =   12420
            TabIndex        =   38
            Top             =   285
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
            TabIndex        =   37
            Top             =   255
            Width           =   900
         End
         Begin VB.Label Label28 
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
            Height          =   270
            Left            =   105
            TabIndex        =   36
            Top             =   1140
            Width           =   1230
         End
      End
      Begin VB.Frame ButtonFrm 
         Height          =   1080
         Left            =   120
         TabIndex        =   49
         Top             =   8625
         Width           =   14865
         Begin VB.CommandButton ExitCmd 
            Caption         =   "E&xit"
            Height          =   400
            Left            =   12105
            TabIndex        =   17
            Top             =   600
            Width           =   2700
         End
         Begin VB.CommandButton SearchCmd 
            Caption         =   "Search"
            Height          =   400
            Left            =   12105
            TabIndex        =   16
            Top             =   195
            Width           =   2700
         End
         Begin VB.CommandButton LastCmd 
            Caption         =   "&Last"
            Height          =   400
            Left            =   9165
            TabIndex        =   18
            Top             =   600
            Width           =   2940
         End
         Begin VB.CommandButton DeleteCmd 
            Caption         =   "&Delete"
            Height          =   400
            Left            =   9165
            TabIndex        =   15
            Top             =   195
            Width           =   2940
         End
         Begin VB.CommandButton FirstCmd 
            Caption         =   "&First"
            Height          =   400
            Left            =   6225
            TabIndex        =   19
            Top             =   600
            Width           =   2940
         End
         Begin VB.CommandButton EditCmd 
            Caption         =   "&Edit"
            Height          =   400
            Left            =   6225
            TabIndex        =   14
            Top             =   195
            Width           =   2940
         End
         Begin VB.CommandButton PreviousCmd 
            Caption         =   "&Previous"
            Height          =   400
            Left            =   3165
            TabIndex        =   20
            Top             =   600
            Width           =   3060
         End
         Begin VB.CommandButton SaveCmd 
            Caption         =   "Save"
            Height          =   400
            Left            =   3165
            TabIndex        =   13
            Top             =   195
            Width           =   3060
         End
         Begin VB.CommandButton NextCmd 
            Caption         =   "&Next"
            Height          =   400
            Left            =   105
            TabIndex        =   21
            Top             =   600
            Width           =   3060
         End
         Begin VB.CommandButton AddCmd 
            Caption         =   "&Add New"
            Height          =   400
            Left            =   105
            TabIndex        =   0
            Top             =   195
            Width           =   3060
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
         Height          =   3810
         Left            =   120
         TabIndex        =   79
         Top             =   4800
         Width           =   14850
         _ExtentX        =   26194
         _ExtentY        =   6720
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
   End
End
Attribute VB_Name = "FrmTxn_MonthwiseSecLoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg, mTempTxnID, mSecurityAgencyID, mSSecurityAgencyID, mSecurityTypeID, mSecurityDesignationID As Variant
Dim mMasterTxnID, mLocationID, mSLocationID, mCMPMID, mSCMPMID                  As Double
Dim mTotalDays, mMonth, mStatus, intMonthDays, mMonthStartDt, mMonthEndDate     As Variant
Dim blnSave As Boolean
Dim tmp1, mYear As Variant
Dim blnsearch As Boolean



Private Sub CmbCMPM_GotFocus()
tmp1 = CmbCMPM.Text

If CmbLocationID.Text = "" Then
   mLocationID = 0
   Exit Sub
End If

tmp = "Select DISTINCT(ME.EmployeeID), EmployeeName,EmployeeNo "
tmp = tmp & " From Mst_Employee ME"
'tmp = tmp & " Left Join Mst_EmpCMPMapping MECM On MECM.EmployeeID = ME.EmployeeID"
tmp = tmp & " Left Join Mst_EmpCMPMapping MECM On MECM.EmployeeID = ME.EmployeeID"
tmp = tmp & " Inner Join Mst_Location ML On ML.LocationID = MECM.LocationID"
'If Gen_WcLocation <> "" Then
'   tmp = tmp & " WHERE ML." & Gen_WcLocation
'End If
If mLocationID <> "" Then
   tmp = tmp & " WHERE ML.LocationID=" & mLocationID
End If
Call Tmp1Table

CmbCMPM.Clear
If TmpRs1.RecordCount = 0 Then
   MsgBox "No CMPM Found!!!"
   CmbCMPM.SetFocus
   Exit Sub
End If
For I = 1 To TmpRs1.RecordCount
    CmbCMPM.AddItem TmpRs1!EmployeeName & "~" & TmpRs1!EmployeeNo
    TmpRs1.MoveNext
Next
'If TmpRs1.RecordCount = 1 Then
'   CmbCMPM.ListIndex = 0
'   Call CmbCMPM_LostFocus
'   TxtAttenSheetID.SetFocus
'Else
'   CmbCMPM.SetFocus
'   CmbCMPM.Text = tmp1
'End If
End Sub

Private Sub CmbCMPM_LostFocus()
Dim strSplitString()        As String
Dim strEmployeeName         As String

If CmbCMPM.Text = "" Then
   mCMPMID = Null
   Exit Sub
End If

TmpRs1.MoveFirst

strEmployeeName = Trim(CmbCMPM.Text)
strSplitString = Split(strEmployeeName, "~")
TmpRs1.Find "EmployeeName = '" & strSplitString(0) & "'"

If TmpRs1.EOF Then
   MsgBox "Invalid Selection "
   CmbCMPM.SetFocus
   Exit Sub
End If
mCMPMID = TmpRs1!EmployeeID
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
If CmbMonth.Text = "" Or CmbYear.Text = "" Then Exit Sub
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

If CmbSecurityAgency.Text <> "" And CmbLocationID.Text <> "" And LCase(tmp) <> LCase(CmbMonth.Text) Then
   Call EditMode
End If
CmdAdd2List.Enabled = False
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

Private Sub Grid_Head()

MSHFlexGrid2.Clear
MSHFlexGrid2.Rows = 3
MSHFlexGrid2.FixedRows = 2
MSHFlexGrid2.Cols = 24

MSHFlexGrid2.TextMatrix(1, 0) = "Security Type"
MSHFlexGrid2.TextMatrix(1, 1) = "Security TypeID"
MSHFlexGrid2.TextMatrix(1, 2) = "Security Designation"
MSHFlexGrid2.TextMatrix(1, 3) = "Security DesignationID"
MSHFlexGrid2.TextMatrix(1, 4) = "Start Date"
MSHFlexGrid2.TextMatrix(1, 5) = "End Date"
MSHFlexGrid2.TextMatrix(1, 6) = "No of Godown"
MSHFlexGrid2.TextMatrix(1, 7) = "Rate per month"
MSHFlexGrid2.TextMatrix(1, 8) = "CMPM Name"
MSHFlexGrid2.TextMatrix(1, 9) = "CMPID"
MSHFlexGrid2.TextMatrix(1, 10) = "Attendance Sheet ID"
MSHFlexGrid2.TextMatrix(1, 11) = "No of Days"
MSHFlexGrid2.TextMatrix(1, 12) = "Esitmated Amount"
MSHFlexGrid2.TextMatrix(1, 13) = "Esitmated S.Tax"
MSHFlexGrid2.TextMatrix(1, 14) = "Total Est. Amount"
MSHFlexGrid2.TextMatrix(1, 15) = "Remarks"
MSHFlexGrid2.TextMatrix(1, 16) = "Monthwise DetailID"
MSHFlexGrid2.TextMatrix(1, 17) = "SecLocMap DetailID"
MSHFlexGrid2.TextMatrix(1, 18) = "Salary Payable Upto"
MSHFlexGrid2.TextMatrix(1, 19) = "Billing No Of Days"
MSHFlexGrid2.TextMatrix(1, 20) = "Billing Amount"
MSHFlexGrid2.TextMatrix(1, 21) = "Billing Service Tax"
MSHFlexGrid2.TextMatrix(1, 22) = "Total Billing Amount"
MSHFlexGrid2.TextMatrix(1, 23) = "BillingID"

MSHFlexGrid2.ColWidth(0) = 800
MSHFlexGrid2.ColWidth(1) = 0
MSHFlexGrid2.ColWidth(2) = 1200
MSHFlexGrid2.ColWidth(3) = 0
MSHFlexGrid2.ColWidth(4) = 900
MSHFlexGrid2.ColWidth(5) = 900
MSHFlexGrid2.ColWidth(6) = 700
MSHFlexGrid2.ColWidth(7) = 700
MSHFlexGrid2.ColWidth(8) = 1200
MSHFlexGrid2.ColWidth(9) = 0
MSHFlexGrid2.ColWidth(10) = 900
MSHFlexGrid2.ColWidth(11) = 600
MSHFlexGrid2.ColWidth(12) = 800
MSHFlexGrid2.ColWidth(13) = 600
MSHFlexGrid2.ColWidth(14) = 600
MSHFlexGrid2.ColWidth(15) = 1500
MSHFlexGrid2.ColWidth(16) = 0
MSHFlexGrid2.ColWidth(17) = 0
MSHFlexGrid2.ColWidth(18) = 600
MSHFlexGrid2.ColWidth(19) = 600
MSHFlexGrid2.ColWidth(20) = 600
MSHFlexGrid2.ColWidth(21) = 800
MSHFlexGrid2.ColWidth(22) = 700
MSHFlexGrid2.ColWidth(23) = 0

MSHFlexGrid2.RowHeight(1) = 850
MSHFlexGrid2.WordWrap = True
End Sub

Private Sub AddCmd_Click()

Edit_Flg = "A"
Frame1.Enabled = True
Frame2.Enabled = True
Call GButtonLock(Me, False)

Call Grid_Head

TxtMonth.Visible = False
CmbMonth.Visible = True
CmbMonth.Text = "" 'MonthName(Month(Date))

TxtYear.Visible = False
CmbYear.Visible = True
CmbYear.Text = "" 'Year(Date)

TxtTxnID.Text = ""
TxtTxnDate.Value = Date

TxtSecurityAgency.Visible = False
CmbSecurityAgency.Visible = True
CmbSecurityAgency.Text = ""
CmdSecurityAgency.Visible = True

TxtLocation.Visible = False
CmbLocationID.Visible = True
CmbLocationID.Text = ""
CmdLocationID.Visible = True

TxtCreated = ""
TxtUpdated = ""

Call ClearFrame

CmdAdd2List.Enabled = False

If RsTxn_MonthwiseSecLoc.RecordCount > 0 Then
   CmbMonth.SetFocus
End If

End Sub

Private Sub ClearFrame()
TxtSecurityType.Text = ""
TxtSecurityDesignation.Text = ""
TxtStartDate.Text = ""
TxtEndDate.Text = ""
TxtSalaryPaidUpto.Text = ""
TxtNoofGodown = ""
TxtRatePerMonth = ""
CmbCMPM.Text = ""
TxtAttenSheetID = ""
TxtNoofDays = ""
TxtEstimatedAmt = ""
TxtEstimatedSTax = ""
TxtTotEstimatedAmt = ""
TxtRemarks.Text = ""
'CmdAdd2List.Caption = "Add to List"
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
    mMonth = Month(CmbYear.Text & "-" & CmbMonth.Text & "-" & 1)
    intMonthDays = Day(DateSerial(CmbYear.Text, mMonth + 1, 0))
    mMonthStartDt = mMonth & "/01" & "/" & CmbYear
    mMonthEndDate = mMonth & "/" & intMonthDays & "/" & CmbYear
End If
mYear = CmbYear.Text
mTotalDays = Day(DateSerial(CmbYear.Text, mMonth + 1, 0))

If CmbSecurityAgency.Text <> "" And CmbLocationID.Text <> "" And LCase(tmp) <> LCase(CmbYear.Text) Then
    Call EditMode
End If
CmdAdd2List.Enabled = False
End Sub

Private Sub CmdAdd2List_Click()
Dim r As Variant

If TxtSecurityType.Text = "" Then
   MsgBox "Select record to update!!!"
   MSHFlexGrid2.SetFocus
   Exit Sub
End If
If CmbCMPM.Text = "" Then
   'MsgBox "Blank CMPM not allowed!!!"
   'CmbCMPM.SetFocus
   'Exit Sub
End If

If CheckDuplicateAttenSheetID = True Then
'   MsgBox "Duplicate Attendance Sheet ID Not Allowed."
'   TxtAttenSheetID.SetFocus
'   Exit Sub
End If

If Trim(TxtAttenSheetID.Text) = "" Then
   MsgBox "Blank Attendance Sheet ID not allowed. If not available give as 'NA'!!!"
   TxtAttenSheetID.SetFocus
   Exit Sub
End If

'If Val(TxtNoofDays) = Then
'   MsgBox "No of Days should be more than '0'!!!"
'   TxtNoofDays.SetFocus
'   Exit Sub
'End If

If Val(TxtEstimatedAmt) = 0 Then
'   MsgBox "Estimated Amount should be more than '0'!!!"
'   TxtEstimatedAmt.SetFocus
'   Exit Sub
End If

If TxtRemarks = "" Then
'Commented by mayur Sir
'   MsgBox "Blank Remarks not allowed!!!"
'   TxtRemarks.SetFocus
'   Exit Sub
End If

If CmdAdd2List.Caption = "Update List" Then
   r = MSHFlexGrid2.RowSel
'   MSHFlexGrid2.TextMatrix(0, 5) = Val(MSHFlexGrid2.TextMatrix(0, 5)) - Val(MSHFlexGrid2.TextMatrix(r, 5))
   'TxtNoofGodown = Val(TxtNoofGodown) - Val(MSHFlexGrid2.TextMatrix(r, 5))
Else
'   r = MSHFlexGrid2.Rows - 1
'   MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
'   MSHFlexGrid2.TextMatrix(r, 0) = ""
End If

MSHFlexGrid2.TextMatrix(r, 7) = TxtRatePerMonth.Text
MSHFlexGrid2.TextMatrix(r, 8) = CmbCMPM.Text
MSHFlexGrid2.TextMatrix(r, 9) = IIf(IsNull(mCMPMID), "", mCMPMID)
MSHFlexGrid2.TextMatrix(r, 10) = IIf(TxtNoofDays = "", "-", TxtAttenSheetID)
MSHFlexGrid2.TextMatrix(r, 11) = Val(TxtNoofDays)
MSHFlexGrid2.TextMatrix(r, 12) = IIf(TxtNoofDays = "", "", Val(TxtEstimatedAmt))
MSHFlexGrid2.TextMatrix(r, 13) = IIf(TxtNoofDays = "", "", Val(TxtEstimatedSTax))
MSHFlexGrid2.TextMatrix(r, 14) = IIf(TxtNoofDays = "", "", Val(TxtTotEstimatedAmt))
MSHFlexGrid2.TextMatrix(r, 15) = TxtRemarks
MSHFlexGrid2.TextMatrix(r, 16) = mTempTxnID

'MSHFlexGrid2.TextMatrix(0, 5) = Val(MSHFlexGrid2.TextMatrix(0, 5)) + Val(MSHFlexGrid2.TextMatrix(r, 5))
'TxtNoofGodown = Val(TxtNoofGodown) + Val(MSHFlexGrid2.TextMatrix(r, 5))
Call ClearFrame
'blnSave = True
CmbCMPM.SetFocus
End Sub

Private Sub CmdCMPM_Click()
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
mTotalDays = Day(DateSerial(mYear, mToMonth + 1, 0))


mFromDate = Format(mYear & "-" & mMonth & "-" & "01", "yyyy/mm/dd")
mToDate = Format(mToYear & "-" & mToMonth & "-" & mTotalDays, "yyyy/mm/dd")


wc = ""
'RsTxn_MonthwiseSecLoc.Filter = ""

tmp = "Select TMSL.MonthwiseTxnID, TMSL.MonthwiseTxnDate, Month(MonthYear),Year(MonthYear), MSA.AgencyName, ML.LocationName "
tmp = tmp & " From Txn_MonthwiseSecLoc TMSL "
tmp = tmp & " Inner Join Mst_SecurityAgency MSA On MSA.SecurityAgencyID = TMSL.SecurityAgencyID "
tmp = tmp & " Inner Join Mst_Location ML On ML.LocationID = TMSL.LocationID"

If OptMonthYear.Value = True Then
   If Gen_DBType = "MDB" Then
      tmp = tmp & " Where MonthYear Between #" & Format(mFromDate, Gen_DatFormat) & "# And #" & Format(mToDate, Gen_DatFormat) & "#) "
      wc = " MonthYear >= #" & Format(mFromDate, Gen_DatFormat) & "# And MonthYear <= #" & Format(mToDate, Gen_DatFormat) & "#) "
   Else
      tmp = tmp & " Where MonthYear Between '" & Format(mFromDate, Gen_DatFormat) & "' And '" & Format(mToDate, Gen_DatFormat) & "' "
      wc = " MonthYear >= '" & Format(mFromDate, Gen_DatFormat) & "' And MonthYear <= '" & Format(mToDate, Gen_DatFormat) & "' "
   End If
ElseIf OptTransactionDate.Value = True Then
   If Gen_DBType = "MDB" Then
      tmp = tmp & " Where MonthwiseTxnDate Between #" & Format(SFrom, Gen_DatFormat) & "# And #" & Format(STo, Gen_DatFormat) & "#) "
      wc = " MonthwiseTxnDate >= #" & Format(SFrom, Gen_DatFormat) & "# And MonthwiseTxnDate <= #" & Format(STo, Gen_DatFormat) & "#) "
   Else
      tmp = tmp & " Where MonthwiseTxnDate Between '" & Format(SFrom, Gen_DatFormat) & "' And '" & Format(STo, Gen_DatFormat) & "' "
      wc = " MonthwiseTxnDate >= '" & Format(SFrom, Gen_DatFormat) & "' And MonthwiseTxnDate <= '" & Format(STo, Gen_DatFormat) & "' "
   End If
End If

If CmbSSecurityAgency.Text <> "" Then
   tmp = tmp & " AND MSA.AgencyName = '" & CmbSSecurityAgency.Text & "' "
   wc = wc & " And SecurityAgencyID = " & mSSecurityAgencyID & ""
End If

If CmbSLocation.Text <> "" Then
   tmp = tmp & " AND ML.LocationName = '" & CmbSLocation.Text & "' "
   wc = wc & " And LocationID = " & mSLocationID
End If

Call TmpTable

'RsTxn_MonthwiseSecLoc.Filter = "" & wc & ""

For iu = 1 To TmpRs.RecordCount
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
    For C = 0 To MSHFlexGrid1.Cols - 1
        MSHFlexGrid1.TextMatrix(iu, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
        If C = 2 Then
           MSHFlexGrid1.TextMatrix(iu, C) = MonthName(MSHFlexGrid1.TextMatrix(iu, C))
        End If
    Next
    MSHFlexGrid1.TextMatrix(iu, 1) = Format(MSHFlexGrid1.TextMatrix(iu, 1), "dd-mmm-yyyy")
    TmpRs.MoveNext
Next

Label22.Refresh
Label22.Caption = TmpRs.RecordCount

End Sub

Private Sub DeleteCmd_Click()
Dim ans As Variant
On Error GoTo msgError

If TxtTxnID.Text = "" Then Exit Sub

'tmp = "SELECT BillTxnID FROM Txn_SecurityBilling WHERE SecurityAgencyID = " & mSecurityAgencyID & ""
'tmp = tmp & " AND Month(MonthYear)=" & mMonth & " AND Year(MonthYear)=" & CmbYear & ""
'Call TmpTable
'
'If TmpRs.RecordCount > 0 Then
'    MsgBox "Month wise details for Security Agency '" & TxtSecurityAgency.Text & "' and Location '" & TxtLocation.Text & "' is already exists."
'    Exit Sub
'End If


tmp = " Select * from Txn_SecurityBilling TSB"
tmp = tmp & " Inner Join Txn_SecurityBillingDetail TSBD On TSB.BillTxnID = TSBD.BillTxnID"
tmp = tmp & " Where TSB.SecurityAgencyID = " & mSecurityAgencyID & " And TSBD.LocationID = " & mLocationID & " And Month(MonthYear)=" & mMonth & " AND Year(MonthYear)=" & CmbYear & ""

Call TmpTable

If TmpRs.RecordCount > 0 Then
   MsgBox "Security Billing already exists . Can't Delete!!!"
   Exit Sub
End If

ans = MsgBox("Do you really want to DELETE this record???", vbYesNo)
If ans = vbYes Then
   tmp = "Delete From Txn_MonthwiseSecLocDetail where MonthwiseTxnID = " & Val(TxtTxnID)
   Call TmpTable
   RsTxn_MonthwiseSecLoc.Delete
   RsTxn_MonthwiseSecLoc.Update
   If RsTxn_MonthwiseSecLoc.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsTxn_MonthwiseSecLoc.MoveNext
   If RsTxn_MonthwiseSecLoc.EOF() Then
      RsTxn_MonthwiseSecLoc.MoveLast
   End If
   Call Indata
End If
CmdAdd2List.Enabled = False

Exit Sub
msgError:
MsgBox "Child record Present "
RsSecLocMapping.CancelUpdate
End Sub

Private Sub EditCmd_Click()

'tmp = "SELECT BillTxnID FROM Txn_SecurityBilling WHERE SecurityAgencyID = " & mSecurityAgencyID & ""
'tmp = tmp & " AND Month(MonthYear)=" & mMonth & " AND Year(MonthYear)=" & CmbYear & ""
'Call TmpTable
'
'If TmpRs.RecordCount > 0 Then
'    MsgBox "Month wise details for Security Agency '" & TxtSecurityAgency.Text & "' and Location '" & TxtLocation.Text & "' is already exists."
'    Exit Sub
'End If

Edit_Flg = "E"
Call GButtonLock(Me, False)

'TxtMonth.Visible = False
'CmbMonth.Visible = True
'CmbMonth.Text = TxtMonth

'TxtYear.Visible = False
'CmbYear.Visible = True
'CmbYear.Text = TxtYear
mMonth = Month(CmbYear & "-" & CmbMonth.Text & "-" & 1)
intMonthDays = Day(DateSerial(CmbYear.Text, mMonth + 1, 0))
mMonthStartDt = mMonth & "/01" & "/" & CmbYear
mMonthEndDate = mMonth & "/" & intMonthDays & "/" & CmbYear

'TxtSecurityAgency.Visible = False
'CmbSecurityAgency.Visible = True
CmbSecurityAgency.Text = TxtSecurityAgency
'CmdSecurityAgency.Visible = True

'TxtLocation.Visible = False
'CmbLocationID.Visible = True
CmbLocationID.Text = TxtLocation
'CmdLocationID.Visible = True

Call ClearFrame
'Frame1.Enabled = True
Frame2.Enabled = True
CmdAdd2List.Enabled = False
'CmbSecurityAgency.SetFocus
blnSave = True

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

blnsearch = False
'Call TableMst_SecLocMapping(" Order by TxnID")

If Gen_WcLocation <> "" Then
   Call TableTxn_MonthwiseSecLoc(" Where " & Gen_WcLocation & " Order by MonthwiseTxnID")
Else
   Call TableTxn_MonthwiseSecLoc(" Order by MonthwiseTxnID")
End If


Call Grid_Head1
SFrom.Value = CDate(From_GenDate) - 900
STo.Value = Date 'To_GenDate
CmbSSecurityAgency.Text = ""
CmbSLocation.Text = ""
Label22.Caption = ""
OptMonthYear.Value = True


If RsTxn_MonthwiseSecLoc.RecordCount = 0 Then
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

Private Sub CmbSecurityAgency_GotFocus()
tmp = CmbSecurityAgency.Text

Call TableMst_SecurityAgency("Where Flag = 'Y'")

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
   CmbLocationID.Text = ""
   mLocationID = Null
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

CmdAdd2List.Enabled = False

If LCase(tmp) <> LCase(RsMst_SecurityAgency!AgencyName) Then
   CmbLocationID.Text = ""
   mLocationID = Null
End If

'If CmbLocationID.Text <> "" And LCase(tmp) <> LCase(RsMst_SecurityAgency!AgencyName) Then
'   Call Grid_Head
'   Call EditMode
'End If
End Sub
Private Sub CmbLocationID_GotFocus()

If CmbSecurityAgency.Text = "" Then
   CmbLocationID.Clear
   MsgBox "Please select Security Agency!!!"
   CmbSecurityAgency.SetFocus
   Exit Sub
End If

tmp1 = CmbLocationID.Text

tmp = " Select Distinct ML.LocationName,ML.LocationID "
tmp = tmp & " From Mst_SecLocMapping MSM"
tmp = tmp & " Inner Join Mst_Location ML On MSM.LocationID = ML.LocationID"
tmp = tmp & " Where SecurityAgencyID = " & mSecurityAgencyID & ""


If Gen_WcLocation <> "" Then
'   Call TableMst_Location
   tmp = tmp & " And MSM." & Gen_WcLocation
End If

Call TmpTable

CmbLocationID.Clear
If TmpRs.RecordCount = 0 Then
   MsgBox "No Location found"
   CmbLocationID.SetFocus
   Exit Sub
End If
For I = 1 To TmpRs.RecordCount
    CmbLocationID.AddItem TmpRs!LocationName
    TmpRs.MoveNext
Next
CmbLocationID.Text = tmp1

End Sub

Private Sub CmbLocationID_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "CmbLocationID")
End If
End Sub

Private Sub CmbLocationID_LostFocus()

If CmbLocationID.Text = "" Then
   mLocationID = 0
   Exit Sub
End If

TmpRs.MoveFirst
TmpRs.Find "LocationName = '" & CmbLocationID.Text & "'"

If TmpRs.EOF Then
   MsgBox "Invalid Location  selection "
   CmbLocationID.SetFocus
   Exit Sub
End If
mLocationID = TmpRs!LocationID
If LCase(tmp1) <> LCase(TmpRs!LocationName) Then
   Call EditMode
End If
End Sub
Private Sub CmdLocationID_Click()
FrmMst_Location.Show
End Sub
Private Sub CmbSSecurityAgency_GotFocus()
tmp = CmbSSecurityAgency.Text
Call TableMst_SecurityAgency("Where Flag = 'A'")
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
Private Sub CmbSLocation_GotFocus()
tmp = CmbSLocation.Text
If Gen_WcLocation = "" Then
   Call TableMst_Location
   
Else
    Call TableMst_Location(" WHERE " & Gen_WcLocation)
End If
CmbSLocation.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Location found"
   CmbSLocationID.SetFocus
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
   MsgBox "Invalid Location  selection "
   CmbSLocation.SetFocus
   Exit Sub
End If
mSLocationID = RsMst_Location!LocationID
End Sub

Private Sub CmdSecurityAgency_Click()
FrmMst_SecurityAgency.Show
End Sub

Private Sub Form_Resize()
If Me.WindowState = vbMaximized Then
   SetScrollBars
End If
End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsTxn_MonthwiseSecLoc.MoveNext
If RsTxn_MonthwiseSecLoc.EOF Then
   RsTxn_MonthwiseSecLoc.MoveLast
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
RsTxn_MonthwiseSecLoc.MovePrevious
If RsTxn_MonthwiseSecLoc.BOF Then
   RsTxn_MonthwiseSecLoc.MoveFirst
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
RsTxn_MonthwiseSecLoc.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsTxn_MonthwiseSecLoc.MoveLast
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
RsTxn_MonthwiseSecLoc.MoveFirst
RsTxn_MonthwiseSecLoc.Find "MonthwiseTxnID = '" & MSHFlexGrid1.TextMatrix(I, 0) & "'"
If Not RsTxn_MonthwiseSecLoc.EOF Then
   Call Indata
   Frame0.Visible = False
End If
End Sub

Private Sub MSHFlexGrid2_Click()
If SaveCmd.Enabled = False Then Exit Sub
Dim r As Variant
r = MSHFlexGrid2.RowSel

If r < 2 Then Exit Sub
If MSHFlexGrid2.TextMatrix(r, 1) = "" Then Call ClearFrame: Exit Sub

'tmp = "SELECT MSBD.MonthwiseTxnDetailID,MSBD.BillTxnID FROM Txn_SecurityBillingDetail MSBD"
'tmp = tmp & " Inner Join Txn_SecurityBilling MSB On MSB.BillTxnID = MSBD.BillTxnID"
'tmp = tmp & " Where MSBD.MonthwiseTxnDetailID = " & Val(MSHFlexGrid2.TextMatrix(r, 16)) & ""
'tmp = tmp & " And SecurityAgencyID = " & mSecurityAgencyID & " And Year(MonthYear) = " & CmbYear.Text & ""
'tmp = tmp & " And Month(MonthYear) = " & mMonth & ""
'Call TmpTable

If MSHFlexGrid2.TextMatrix(r, 23) <> 0 Then
   'MsgBox "Security Billing for Security Agency '" & TxtSecurityAgency.Text & "' And Month-Year '" & TxtMonth.Text & "-" & TxtYear.Text & "' already exists."
   MsgBox "Security Billing already genereted ID : " & MSHFlexGrid2.TextMatrix(r, 23) 'for Security Agency '" & TxtSecurityAgency.Text & "' And Month-Year '" & TxtMonth.Text & "-" & TxtYear.Text & "' already exists."
   Exit Sub
End If

'Commneted for temporary on  18/Mar/2011
If MSHFlexGrid2.row = r And MSHFlexGrid2.CellBackColor = vbRed Then Exit Sub

'MSHFlexGrid2.TextMatrix(0, 18) = "Salary Paid Upto"
If Year(MSHFlexGrid2.TextMatrix(r, 18)) < Year(mMonthEndDate) Then
    MsgBox "Salary Payable Upto Date is expired."
Else
    If Month(MSHFlexGrid2.TextMatrix(r, 18)) < Month(mMonthEndDate) And Year(MSHFlexGrid2.TextMatrix(r, 18)) = Year(mMonthEndDate) Then
        MsgBox "Salary Payable Upto Date is expired."
    End If
End If

CmdAdd2List.Enabled = True

TxtSecurityType.Text = MSHFlexGrid2.TextMatrix(r, 0)
mSecurityTypeID = MSHFlexGrid2.TextMatrix(r, 1)
TxtSecurityDesignation.Text = MSHFlexGrid2.TextMatrix(r, 2)
mSecurityDesignationID = MSHFlexGrid2.TextMatrix(r, 3)
TxtStartDate.Text = IIf(MSHFlexGrid2.TextMatrix(r, 4) = "", Date, MSHFlexGrid2.TextMatrix(r, 4))
TxtEndDate.Text = IIf(MSHFlexGrid2.TextMatrix(r, 5) = "", Date, MSHFlexGrid2.TextMatrix(r, 5))
TxtRatePerMonth = Trim(MSHFlexGrid2.TextMatrix(r, 7))
TxtNoofGodown = Trim(MSHFlexGrid2.TextMatrix(r, 6))
CmbCMPM.Text = Trim(MSHFlexGrid2.TextMatrix(r, 8))
mCMPMID = Val(MSHFlexGrid2.TextMatrix(r, 9))
TxtAttenSheetID = Trim(MSHFlexGrid2.TextMatrix(r, 10))
TxtEstimatedSTax.Text = Trim(MSHFlexGrid2.TextMatrix(r, 13))
If Trim(MSHFlexGrid2.TextMatrix(r, 11)) = "" Then
   Call GetNoofDays
   Call CalcEstimateAmt
Else
   TxtNoofDays = Trim(MSHFlexGrid2.TextMatrix(r, 11))
   TxtEstimatedAmt = Trim(MSHFlexGrid2.TextMatrix(r, 12))
   TxtTotEstimatedAmt = Val(MSHFlexGrid2.TextMatrix(r, 14))
End If
TxtEstimatedSTax = Val(MSHFlexGrid2.TextMatrix(r, 13))
'TxtTotEstimatedAmt = Val(MSHFlexGrid2.TextMatrix(r, 14))
TxtRemarks = Trim(MSHFlexGrid2.TextMatrix(r, 15))
mTempTxnID = Val(MSHFlexGrid2.TextMatrix(r, 16))
TxtSalaryPaidUpto = Trim(MSHFlexGrid2.TextMatrix(r, 18))
'CmdAdd2List.Caption = "Update List"

If CmbCMPM.ListCount = 1 Then
   CmbCMPM.ListIndex = 0
   Call CmbCMPM_LostFocus
   TxtAttenSheetID.SetFocus
Else
   CmbCMPM.SetFocus
End If
End Sub
Private Sub SaveCmd_Click()

Dim dtMonthYear     As Date

If MSHFlexGrid2.Rows <= 3 Then
   MsgBox "No details found in grid!!!"
   CmbCMPM.SetFocus
   Exit Sub
End If

If ChkGridDetails = False Then
   MsgBox "Please update atlease one row from grid!!!"
   CmbCMPM.SetFocus
   Exit Sub
End If

'tmp = "SELECT AttendanceSheetID,MonthwiseTxnID FROM Txn_MonthwiseSecLocDetail"
'Call TmpTable
'For I = 2 To MSHFlexGrid2.Rows - 1
'    If TmpRs.RecordCount > 0 Then TmpRs.MoveFirst Else Exit For
'    TmpRs.Find "AttendanceSheetID ='" & Trim(MSHFlexGrid2.TextMatrix(I, 10)) & "'"
'    If Not TmpRs.EOF Then
'       If TmpRs!MonthwiseTxnID <> Val(TxtTxnID.Text) Then
'          MsgBox "Attendance Sheet ID '" & Trim(MSHFlexGrid2.TextMatrix(I, 10)) & "' can't be duplicate!!!"
'          MSHFlexGrid2.Row = I
'          MSHFlexGrid2.Col = 10
'          MSHFlexGrid2.SetFocus
'          Exit Sub
'       End If
'    End If
'Next

If Edit_Flg = "A" Then
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
   If CmbSecurityAgency.Text = "" Then
      MsgBox "Blank Security Agency not allowed!!!"
      CmbSecurityAgency.SetFocus
      Exit Sub
   End If
   If CmbLocationID.Text = "" Then
      MsgBox "Blank Location not allowed!!!"
      CmbLocationID.SetFocus
      Exit Sub
   End If
   
   'tmp = "SELECT MonthwiseTxnID FROM Txn_MonthwiseSecLoc WHERE MonthYear='" & dtMonthYear & "' AND SecurityAgencyID = " & mSecurityAgencyID & " And LocationID = " & mLocationID
   'Call TmpTable
   
   'If TmpRs.RecordCount > 0 Then
   '   MsgBox "Duplicate Record Not Allowed !!"
   '   Exit Sub
   'End If
   RsTxn_MonthwiseSecLoc.AddNew
   RsTxn_MonthwiseSecLoc!MonthWiseTxnID = GetMaxTxnID
   RsTxn_MonthwiseSecLoc!G_UID = GetGUID
   TxtTxnID = RsTxn_MonthwiseSecLoc!MonthWiseTxnID
   TxtG_UID = RsTxn_MonthwiseSecLoc!G_UID
   RsTxn_MonthwiseSecLoc!MonthwiseTxnDate = TxtTxnDate.Value
   RsTxn_MonthwiseSecLoc!MonthYear = dtMonthYear
   RsTxn_MonthwiseSecLoc!SecurityAgencyID = mSecurityAgencyID
   RsTxn_MonthwiseSecLoc!LocationID = mLocationID
   If IsNull(RsTxn_MonthwiseSecLoc!CreatedBy) = True Or RsTxn_MonthwiseSecLoc!CreatedBy = "" Then
      RsTxn_MonthwiseSecLoc!CreatedBy = Gen_UserLoginID
      RsTxn_MonthwiseSecLoc!CreatedDate = Now
   Else
      RsTxn_MonthwiseSecLoc!UpdatedBy = Gen_UserLoginID
      RsTxn_MonthwiseSecLoc!UpdatedDate = Now
   End If
   RsTxn_MonthwiseSecLoc.Update
   TxtCreated = IIf(IsNull(RsTxn_MonthwiseSecLoc!CreatedBy), "", RsTxn_MonthwiseSecLoc!CreatedBy) & " :- " & IIf(IsNull(RsTxn_MonthwiseSecLoc!CreatedDate), "", RsTxn_MonthwiseSecLoc!CreatedDate)
   TxtUpdated = IIf(IsNull(RsTxn_MonthwiseSecLoc!UpdatedBy), "", RsTxn_MonthwiseSecLoc!UpdatedBy) & " :- " & IIf(IsNull(RsTxn_MonthwiseSecLoc!UpdatedDate), "", RsTxn_MonthwiseSecLoc!UpdatedDate)
   RsTxn_MonthwiseSecLoc.Requery
'   RsTxn_MonthwiseSecLoc.MoveFirst
'   RsTxn_MonthwiseSecLoc.Find "G_UID = '" & TxtG_UID & "'"
End If

'RsTxn_MonthwiseSecLoc.Filter = ""
Call SaveMonthwiseSecLocDetails

RsTxn_MonthwiseSecLoc.MoveFirst
RsTxn_MonthwiseSecLoc.Find "MonthwiseTxnID = " & TxtTxnID & ""

Call Indata
End Sub
Private Function ChkGridDetails() As Boolean
Dim Retval As Boolean

Retval = False

For I = 2 To MSHFlexGrid2.Rows - 1
    If MSHFlexGrid2.TextMatrix(I, 2) = "" Then Exit For
    If MSHFlexGrid2.TextMatrix(I, 10) <> "" Then
       Retval = True
       Exit For
    End If
Next
ChkGridDetails = Retval
End Function
Private Sub SaveMonthwiseSecLocDetails()

'Call DeleteMonthwiseSecLocDetail
Call TableTxn_MonthwiseSecLocDetails("Where MonthwiseTxnID=" & TxtTxnID & "") '(" Where ClaimID = " & TxtClaimID & " AND SM_Prefix = '" & TxtPreFix & "'")
'If RsTxn_MonthwiseSecLocDetail.RecordCount < 1 Then
'   RsTxn_MonthwiseSecLocDetail.AddNew
'End If
For I = 2 To MSHFlexGrid2.Rows - 1
    If MSHFlexGrid2.TextMatrix(I, 10) <> "" Then
       RsTxn_MonthwiseSecLocDetail.Find "SecLocMapDetailID = " & Val(MSHFlexGrid2.TextMatrix(I, 17))
       If RsTxn_MonthwiseSecLocDetail.EOF Then
          RsTxn_MonthwiseSecLocDetail.AddNew
          RsTxn_MonthwiseSecLocDetail!MonthwiseTxnDetailID = GetMaxTxnDetails
       Else
          RsTxn_MonthwiseSecLocDetail!MonthwiseTxnDetailID = MSHFlexGrid2.TextMatrix(I, 16)
       End If
       
       RsTxn_MonthwiseSecLocDetail!MonthWiseTxnID = TxtTxnID
       RsTxn_MonthwiseSecLocDetail!SecurityTypeID = MSHFlexGrid2.TextMatrix(I, 1)
       RsTxn_MonthwiseSecLocDetail!SecurityDesignationID = MSHFlexGrid2.TextMatrix(I, 3)
       RsTxn_MonthwiseSecLocDetail!StartDate = MSHFlexGrid2.TextMatrix(I, 4)
       RsTxn_MonthwiseSecLocDetail!EndDate = MSHFlexGrid2.TextMatrix(I, 5)
       RsTxn_MonthwiseSecLocDetail!NoofGodown = Val(MSHFlexGrid2.TextMatrix(I, 6))
       RsTxn_MonthwiseSecLocDetail!RatePerMonth = Val(MSHFlexGrid2.TextMatrix(I, 7))
       If MSHFlexGrid2.TextMatrix(I, 9) <> "" Then
          RsTxn_MonthwiseSecLocDetail!CMPMID = Val(MSHFlexGrid2.TextMatrix(I, 9))
       Else
          RsTxn_MonthwiseSecLocDetail!CMPMID = Null
       End If
       RsTxn_MonthwiseSecLocDetail!AttendanceSheetID = Trim(MSHFlexGrid2.TextMatrix(I, 10))
       RsTxn_MonthwiseSecLocDetail!NoOfDays = Val(MSHFlexGrid2.TextMatrix(I, 11))
       RsTxn_MonthwiseSecLocDetail!EstimatedAmt = Val(MSHFlexGrid2.TextMatrix(I, 12))
       RsTxn_MonthwiseSecLocDetail!EstimatedSTax = Val(MSHFlexGrid2.TextMatrix(I, 13))
       RsTxn_MonthwiseSecLocDetail!TotEstimatedAmt = Val(MSHFlexGrid2.TextMatrix(I, 14))
       RsTxn_MonthwiseSecLocDetail!Remarks = Trim(MSHFlexGrid2.TextMatrix(I, 15))
       RsTxn_MonthwiseSecLocDetail!SecLocMapDetailID = Val(MSHFlexGrid2.TextMatrix(I, 17))
       RsTxn_MonthwiseSecLocDetail!SalaryPaidUpto = MSHFlexGrid2.TextMatrix(I, 18)
       RsTxn_MonthwiseSecLocDetail!CreatedBy = Gen_UserLoginID
       RsTxn_MonthwiseSecLocDetail!CreatedDate = Now
       RsTxn_MonthwiseSecLocDetail.Update
       RsTxn_MonthwiseSecLocDetail.Requery
    End If
Next
End Sub

Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   If RsTxn_MonthwiseSecLoc.RecordCount = 0 Then Exit Sub
   
   Call GButtonLock(Me, True)
   RsTxn_MonthwiseSecLoc.MoveFirst
   RsTxn_MonthwiseSecLoc.Find "G_UID = '" & TxtG_UID & "'"
   If RsTxn_MonthwiseSecLoc.EOF Then
      RsTxn_MonthwiseSecLoc.MoveLast
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

RsTxn_MonthwiseSecLoc.MoveFirst

''===========comment on 7 july 2012 for not clearing search grid=============
'If blnsearch = False Then
'   Call Grid_Head1
'   SFrom.Value = CDate(From_GenDate) - 900
'   STo.Value = Date 'To_GenDate
'   CmbSSecurityAgency.Text = ""
'   CmbSLocation.Text = ""
'   Label22.Caption = ""
'   OptMonthYear.Value = True
'End If
'
'blnsearch = True
''===========comment on 7 july 2012 for not clearing search grid=============

End Sub

Private Sub TxtAttenSheetID_LostFocus()
If TxtAttenSheetID = "" Then Exit Sub
'If IsNumeric(TxtAttenSheetID) = False Or Val(TxtAttenSheetID) < 0 Then
'   MsgBox "Invalid Format!!!"
'   TxtAttenSheetID.SetFocus
'   Exit Sub
'End If
End Sub

Private Sub TxtEstimatedAmt_LostFocus()
If TxtEstimatedAmt = "" Then Exit Sub
If IsNumeric(TxtEstimatedAmt) = False Or Val(TxtEstimatedAmt) < 0 Then
   MsgBox "Invalid Format!!!"
   TxtEstimatedAmt.SetFocus
   Exit Sub
End If
End Sub

Private Sub GetNoofDays()
If Val(TxtEndDate.Text) = 0 Then Exit Sub

If Month(TxtEndDate.Text) = mMonth And Year(TxtEndDate.Text) <= CmbYear.Text Then
   TxtNoofDays.Text = Day(TxtEndDate.Text) + 1
ElseIf Month(TxtStartDate.Text) = mMonth And Year(TxtStartDate.Text) <= CmbYear.Text Then
   intMonthDays = Day(DateSerial(CmbYear.Text, mMonth + 1, 0))
   TxtNoofDays.Text = (intMonthDays - Day(TxtStartDate.Text)) + 1
Else
   intMonthDays = Day(DateSerial(CmbYear.Text, mMonth + 1, 0))
   TxtNoofDays.Text = intMonthDays
End If
End Sub

Private Sub TxtEstimatedSTax_LostFocus()
TxtTotEstimatedAmt.Text = Val(TxtEstimatedAmt.Text) + Val(TxtEstimatedSTax.Text)
End Sub

Private Sub TxtNoofDays_LostFocus()
If TxtNoofDays = "" Then Exit Sub
If IsNumeric(TxtNoofDays) = False Or Val(TxtNoofDays) < 0 Then
   MsgBox "Invalid Format!!!"
   TxtNoofDays.SetFocus
   Exit Sub
End If
If Val(TxtEndDate.Text) = 0 Then Exit Sub

'If Month(TxtEndDate.Text) = mMonth And Year(TxtEndDate.Text) <= CmbYear.Text Then
If Month(TxtEndDate.Text) = mMonth And Year(TxtSalaryPaidUpto.Text) <= CmbYear.Text Then
   If Val(TxtNoofDays.Text) > Day(TxtSalaryPaidUpto.Text) + 1 Then
      MsgBox "'No. of Days' can't be greater than days of 'Salary Payable Upto'!!!"
      TxtNoofDays.SetFocus
      Exit Sub
   End If
ElseIf Month(TxtStartDate.Text) = mMonth And Year(TxtStartDate.Text) = CmbYear.Text Then
   intMonthDays = Day(DateSerial(CmbYear.Text, mMonth + 1, 0))
   If Val(TxtNoofDays.Text) > (intMonthDays - Day(TxtStartDate.Text) + 1) Then
      MsgBox "'No. of Days' can't be greater than days of 'Start date'!!!"
      TxtNoofDays.Text = ""
      TxtNoofDays.SetFocus
      Exit Sub
   End If
ElseIf Val(TxtNoofDays.Text) > DateDiff("d", TxtStartDate, TxtSalaryPaidUpto) Then
       MsgBox "'No. of Days' can't be greater than days between 'Start date' and 'Salary Payable Upto'!!!"
       TxtNoofDays.SetFocus
       Exit Sub
ElseIf Val(mTotalDays) < Val(TxtNoofDays) Then
       MsgBox "'No. of Days' can't be greater than total days for selected month!!!"
       TxtNoofDays.SetFocus
       Exit Sub
End If
Call CalcEstimateAmt
End Sub

Private Sub TxtNoofGodown_LostFocus()
If TxtNoofGodown = "" Then Exit Sub
If IsNumeric(TxtNoofGodown) = False Or Val(TxtNoofGodown) < 0 Then
   MsgBox "Invalid Format!!!"
   TxtNoofGodown.SetFocus
   Exit Sub
End If
End Sub

Private Sub TxtRatePerMonth_LostFocus()
If TxtRatePerMonth = "" Then Exit Sub
If IsNumeric(TxtRatePerMonth) = False Or Val(TxtRatePerMonth) < 0 Then
   MsgBox "Invalid Format!!!"
   TxtRatePerMonth.SetFocus
   Exit Sub
End If

TxtRatePerMonth = Format(TxtRatePerMonth, "########0.00")
End Sub

Function ChkDuplicateExpenseType() As Boolean
'Dim retval As Boolean
'retval = True
'If MSHFlexGrid2.Rows > 2 Then
'   For u = 2 To MSHFlexGrid2.Rows - 1
'       If MSHFlexGrid2.TextMatrix(u, 1) = CmbExpenseTypeID.Text Then
'          If LCase(CmdAdd2List.Caption) = "add to list" Then
'             retval = False
'             Exit For
'          Else
'             If u <> MSHFlexGrid2.RowSel Then
'                retval = False
'             End If
'          End If
'       End If
'   Next
'End If
'
'ChkDuplicateExpenseType = retval
End Function

Private Sub DeleteMonthwiseSecLocDetail()
'tmp = "Delete From Txn_MonthwiseSecLocDetail Where MonthwiseTxnID=" & TxtTxnID & " "
'Call Tmp4Table
End Sub

Private Function GetMaxTxnID() As Double
Dim Retval As Double
tmp = "Select max(MonthwiseTxnID) from Txn_MonthwiseSecLoc "
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetMaxTxnID = Retval
End Function

Function GetMaxTxnDetails() As Double
Dim Retval As Double
Retval = 0
tmp = "Select max(MonthwiseTxnDetailID) From Txn_MonthwiseSecLocDetail " 'Where ClaimID = " & TxtClaimID & " And SM_prefix = '" & TxtPreFix & "'"
Call TmpTable
If TmpRs.RecordCount > 0 Then
   Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
End If
Retval = Retval + 1
TmpRs.Close
GetMaxTxnDetails = Retval

End Function

Public Sub Indata()

Call ClearFrame
Frame1.Enabled = False
Frame2.Enabled = False

TxtMonth.Visible = True
CmbMonth.Visible = False

TxtYear.Visible = True
CmbYear.Visible = False

TxtSecurityAgency.Visible = True
CmbSecurityAgency.Visible = False
CmdSecurityAgency.Visible = Fasle

TxtLocation.Visible = True
CmbLocationID.Visible = False
CmbLocationID = TxtLocation
CmdLocationID.Visible = False
CmdAdd2List.Enabled = False

TxtTxnID = RsTxn_MonthwiseSecLoc!MonthWiseTxnID
TxtTxnDate.Value = RsTxn_MonthwiseSecLoc!MonthwiseTxnDate
TxtMonth = MonthName(Month(RsTxn_MonthwiseSecLoc!MonthYear))
CmbMonth.Text = TxtMonth
TxtYear = Year(RsTxn_MonthwiseSecLoc!MonthYear)
CmbYear.Text = TxtYear
mMonth = Month(CmbYear & "-" & CmbMonth.Text & "-" & 1)
mSecurityAgencyID = RsTxn_MonthwiseSecLoc!SecurityAgencyID
TxtSecurityAgency = GetSecurityAgency(RsTxn_MonthwiseSecLoc!SecurityAgencyID, "N")
mLocationID = RsTxn_MonthwiseSecLoc!LocationID
TxtLocation = GetLocationName(RsTxn_MonthwiseSecLoc!LocationID)
TxtG_UID = IIf(IsNull(RsTxn_MonthwiseSecLoc!G_UID), "", RsTxn_MonthwiseSecLoc!G_UID)
TxtCreated = IIf(IsNull(RsTxn_MonthwiseSecLoc!CreatedBy), "", RsTxn_MonthwiseSecLoc!CreatedBy) & " :- " & IIf(IsNull(RsTxn_MonthwiseSecLoc!CreatedDate), "", RsTxn_MonthwiseSecLoc!CreatedDate)
TxtUpdated = IIf(IsNull(RsTxn_MonthwiseSecLoc!UpdatedBy), "", RsTxn_MonthwiseSecLoc!UpdatedBy) & " :- " & IIf(IsNull(RsTxn_MonthwiseSecLoc!UpdatedDate), "", RsTxn_MonthwiseSecLoc!UpdatedDate)

mTotalDays = Day(DateSerial(CmbYear.Text, mMonth + 1, 0))


Call InDataMonthwiseSecLocDetails
'If RsTxn_MonthwiseSecLoc.RecordCount = 1 Then
'   Call GButtonLock_1(Me, True)
'Else
   Call GButtonLock(Me, True)
'End If
blnSave = False
End Sub

Private Sub InDataMonthwiseSecLocDetails()
Dim strSplitString()        As String
Dim strEmployeeName         As String

Call Grid_Head

'Call TableTableTxn_MonthwiseSecLocDetails(" Where MonthwiseTxnID= " & TxtTxnID & " ") 'AND SM_Prefix = '" & TxtPreFix & "'")
'Tem Query is required.....-Amit
tmp = " SELECT TMSL.MonthwiseTxnDetailID, TMSL.SecurityTypeID, SecurityShifts, TMSL.SecurityDesignationID, SecurityDesignation, StartDate,"
tmp = tmp & " EndDate, NoofGodown, RatePerMonth, ME.EmployeeName + '~'+ ME.EmployeeNo As EmpName, CMPMID, AttendanceSheetID,"
tmp = tmp & " NoofDays, EstimatedAmt, EstimatedSTax,TotEstimatedAmt, TMSL.Remarks, SecLocMapDetailID, SalarypaidUpto,BillingNoOfDays,BillingAmount,BillingSTax,BillingTotalAmount,BillingID"
tmp = tmp & " FROM Txn_MonthwiseSecLocDetail TMSL"
tmp = tmp & " Inner Join Mst_SecurityShifts MSS On MSS.SecurityShiftsID = TMSL.SecurityTypeID"
tmp = tmp & " Inner Join Mst_SecurityDesignation MSD On MSD.SecurityDesignationID = TMSL.SecurityDesignationID"
tmp = tmp & " Left Join Mst_Employee ME On ME.EmployeeID = TMSL.CMPMID"
tmp = tmp & " Where MonthwiseTxnID = " & TxtTxnID & " ORDER BY MonthwiseTxnDetailID"
Call TmpTable

If TmpRs.RecordCount > 0 Then
    MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + TmpRs.RecordCount
    For I = 2 To MSHFlexGrid2.Rows - 2
        MSHFlexGrid2.TextMatrix(I, 0) = TmpRs!SecurityShifts
        MSHFlexGrid2.TextMatrix(I, 1) = TmpRs!SecurityTypeID
        MSHFlexGrid2.TextMatrix(I, 2) = TmpRs!SecurityDesignation
        MSHFlexGrid2.TextMatrix(I, 3) = TmpRs!SecurityDesignationID
        MSHFlexGrid2.TextMatrix(I, 4) = TmpRs!StartDate
        MSHFlexGrid2.TextMatrix(I, 5) = TmpRs!EndDate
        MSHFlexGrid2.TextMatrix(I, 6) = TmpRs!NoofGodown
        MSHFlexGrid2.TextMatrix(0, 6) = Val(MSHFlexGrid2.TextMatrix(0, 6)) + Val(TmpRs!NoofGodown)
        MSHFlexGrid2.TextMatrix(I, 7) = TmpRs!RatePerMonth
        'strEmployeeName = GetEmployeeNonName(RsTxn_MonthwiseSecLocDetail!CMPMID)
        'strSplitString = Split(strEmployeeName, "~")
        MSHFlexGrid2.TextMatrix(I, 8) = IIf(IsNull(TmpRs!EmpName), "", TmpRs!EmpName)
        MSHFlexGrid2.TextMatrix(I, 9) = IIf(IsNull(TmpRs!CMPMID), "", TmpRs!CMPMID)
        MSHFlexGrid2.TextMatrix(I, 10) = IIf(IsNull(TmpRs!AttendanceSheetID), "", TmpRs!AttendanceSheetID)
        MSHFlexGrid2.TextMatrix(I, 11) = TmpRs!NoOfDays
        MSHFlexGrid2.TextMatrix(I, 12) = TmpRs!EstimatedAmt
        MSHFlexGrid2.TextMatrix(I, 13) = IIf(IsNull(TmpRs!EstimatedSTax), "", TmpRs!EstimatedSTax)
        MSHFlexGrid2.TextMatrix(I, 14) = IIf(IsNull(TmpRs!TotEstimatedAmt), "", TmpRs!TotEstimatedAmt)
        MSHFlexGrid2.TextMatrix(I, 15) = IIf(IsNull(TmpRs!Remarks), "", TmpRs!Remarks)
        MSHFlexGrid2.TextMatrix(I, 16) = TmpRs!MonthwiseTxnDetailID
        MSHFlexGrid2.TextMatrix(I, 17) = IIf(IsNull(TmpRs!SecLocMapDetailID), 0, TmpRs!SecLocMapDetailID)
        MSHFlexGrid2.TextMatrix(I, 18) = IIf(IsNull(TmpRs!SalaryPaidUpto), 0, TmpRs!SalaryPaidUpto)
        MSHFlexGrid2.TextMatrix(I, 19) = IIf(IsNull(TmpRs!BillingNoOfDays), 0, TmpRs!BillingNoOfDays)
        MSHFlexGrid2.TextMatrix(I, 20) = IIf(IsNull(TmpRs!BillingAmount), 0, TmpRs!BillingAmount)
        MSHFlexGrid2.TextMatrix(I, 21) = IIf(IsNull(TmpRs!BillingSTax), 0, TmpRs!BillingSTax)
        MSHFlexGrid2.TextMatrix(I, 22) = IIf(IsNull(TmpRs!BillingTotalAmount), 0, TmpRs!BillingTotalAmount)
        MSHFlexGrid2.TextMatrix(I, 23) = IIf(IsNull(TmpRs!BillingID), 0, TmpRs!BillingID)
        
        MSHFlexGrid2.row = I
        For x = 19 To 23
            MSHFlexGrid2.Col = x
            MSHFlexGrid2.CellBackColor = vbYellow
        Next
        TmpRs.MoveNext
    Next
End If
End Sub
Private Sub Grid_Head1()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.FixedRows = 1
MSHFlexGrid1.Cols = 6

MSHFlexGrid1.TextMatrix(0, 0) = "Monthwise Txn ID"
MSHFlexGrid1.TextMatrix(0, 1) = "Monthwise Txn Date"
MSHFlexGrid1.TextMatrix(0, 2) = "Month"
MSHFlexGrid1.TextMatrix(0, 3) = "Year"
MSHFlexGrid1.TextMatrix(0, 4) = "Security Agency"
MSHFlexGrid1.TextMatrix(0, 5) = "Location Name"

MSHFlexGrid1.ColWidth(0) = 1000
MSHFlexGrid1.ColWidth(1) = 1200
MSHFlexGrid1.ColWidth(2) = 1000
MSHFlexGrid1.ColWidth(3) = 1000
MSHFlexGrid1.ColWidth(4) = 3000
MSHFlexGrid1.ColWidth(5) = 3500
End Sub
Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Txn_MonthwiseSecLoc.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Txn_MonthwiseSecLoc.xls")
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

Private Sub GetSecLocMapping()
Dim Retval As Boolean
Retval = True

If CmbSecurityAgency.Text = "" Or CmbLocationID.Text = "" Then
   MsgBox "Either Security agency or Location Name is blank."
   CmbSecurityAgency.SetFocus
   Exit Sub
End If

tmp = " SELECT MSLM.TxnID, MSLM.TxnDate, AgencyName, LocationName "
tmp = tmp & " FROM Mst_SecLocMapping MSLM"
tmp = tmp & " Inner Join Mst_SecurityAgency MSA On MSA.SecurityAgencyID = MSLM.SecurityAgencyID "
tmp = tmp & " Inner Join Mst_Location ML On ML.LocationID = MSLM.LocationID"
tmp = tmp & " Where MSLM.SecurityAgencyID = " & mSecurityAgencyID & " And ML.LocationID = " & mLocationID & ""

Call TmpTable

If TmpRs.RecordCount > 0 Then
   mMasterTxnID = TmpRs!TxnID
   'mSecurityAgencyID = TmpRs!SecurityAgencyID
   TxtSecurityAgency = TmpRs!AgencyName
   'mLocationID = RsSecLocMapping!LocationID
   TxtLocation = TmpRs!LocationName
   Call InDataSecLocMappingDetails
Else
   TxtTxnID.Text = ""
   TxtTxnDate.Value = Date
   TxtCreated = ""
   TxtUpdated = ""
   Call Grid_Head
End If
Call ClearFrame
'RsSecLocMapping.Filter = ""
End Sub

Private Sub InDataSecLocMappingDetails()

Call Grid_Head
Call TableMst_SecLocMappingDetails(" Where TxnID= " & mMasterTxnID & " ") 'AND SM_Prefix = '" & TxtPreFix & "'")
If RsMst_SecLocMappingDetails.RecordCount > 0 Then
    MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + RsMst_SecLocMappingDetails.RecordCount
    For I = 2 To MSHFlexGrid2.Rows - 2
        MSHFlexGrid2.TextMatrix(I, 0) = GetSecurityType(RsMst_SecLocMappingDetails!SecurityTypeID)
        MSHFlexGrid2.TextMatrix(I, 1) = RsMst_SecLocMappingDetails!SecurityTypeID
        MSHFlexGrid2.TextMatrix(I, 2) = GetSecurityDesignation(RsMst_SecLocMappingDetails!SecurityDesignationID)
        MSHFlexGrid2.TextMatrix(I, 3) = RsMst_SecLocMappingDetails!SecurityDesignationID
        MSHFlexGrid2.TextMatrix(I, 4) = RsMst_SecLocMappingDetails!StartDate
        If Format(RsMst_SecLocMappingDetails!StartDate, Gen_DatFormat) > Format(mMonthEndDate, Gen_DatFormat) Then
           For C = 0 To MSHFlexGrid2.Cols - 1
                MSHFlexGrid2.Col = C:      MSHFlexGrid2.row = I:                MSHFlexGrid2.CellBackColor = vbRed
           Next
        End If
        MSHFlexGrid2.TextMatrix(I, 5) = RsMst_SecLocMappingDetails!EndDate
        'If Format(RsMst_SecLocMappingDetails!EndDate, Gen_DatFormat) < Format(mMonthStartDt, Gen_DatFormat) Then
        'Later it will be on end date only.....!
        If Format(RsMst_SecLocMappingDetails!SalaryPaidUpto, Gen_DatFormat) < Format(mMonthStartDt, Gen_DatFormat) Then
           For C = 0 To MSHFlexGrid2.Cols - 1
                MSHFlexGrid2.Col = C:      MSHFlexGrid2.row = I:                MSHFlexGrid2.CellBackColor = vbRed
           Next
        End If
        MSHFlexGrid2.TextMatrix(I, 6) = RsMst_SecLocMappingDetails!NoofGodown
        MSHFlexGrid2.TextMatrix(0, 6) = Val(MSHFlexGrid2.TextMatrix(0, 6)) + Val(RsMst_SecLocMappingDetails!NoofGodown)
        MSHFlexGrid2.TextMatrix(I, 7) = RsMst_SecLocMappingDetails!RatePerMonth
        MSHFlexGrid2.TextMatrix(I, 17) = RsMst_SecLocMappingDetails!TxnDetailID
        MSHFlexGrid2.TextMatrix(I, 18) = IIf(IsNull(RsMst_SecLocMappingDetails!SalaryPaidUpto), "", RsMst_SecLocMappingDetails!SalaryPaidUpto)
        MSHFlexGrid2.TextMatrix(I, 19) = 0 'IIf(IsNull(TmpRs!BillingNoOfDays), 0, TmpRs!BillingNoOfDays)
        MSHFlexGrid2.TextMatrix(I, 20) = 0 'IIf(IsNull(TmpRs!BillingAmount), 0, TmpRs!BillingAmount)
        MSHFlexGrid2.TextMatrix(I, 21) = 0 'IIf(IsNull(TmpRs!BillingSTax), 0, TmpRs!BillingSTax)
        MSHFlexGrid2.TextMatrix(I, 22) = 0 'IIf(IsNull(TmpRs!BillingTotalAmount), 0, TmpRs!BillingTotalAmount)
        MSHFlexGrid2.TextMatrix(I, 23) = 0
         
        MSHFlexGrid2.row = I
        For x = 19 To 23
            MSHFlexGrid2.Col = x
            MSHFlexGrid2.CellBackColor = vbYellow
        Next

        RsMst_SecLocMappingDetails.MoveNext
    Next
End If
End Sub

Private Sub CalcEstimateAmt()
Dim dblEstimateAmt      As Double
intMonthDays = 0
If CmbYear.Text <> "" Then
    intMonthDays = Day(DateSerial(CmbYear.Text, mMonth + 1, 0))
    dblEstimateAmt = Val(TxtRatePerMonth.Text) * Val(TxtNoofDays)
    dblEstimateAmt = Format(dblEstimateAmt / intMonthDays, "########0.00")
    TxtEstimatedAmt.Text = Format(dblEstimateAmt, "#.00")
End If
TxtTotEstimatedAmt.Text = Val(TxtEstimatedAmt.Text) + Val(TxtEstimatedSTax.Text)
End Sub
Private Sub EditMode()

tmp = " SELECT MonthwiseTxnID, MonthwiseTxnDate, G_UID, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy FROM Txn_MonthwiseSecLoc"
tmp = tmp & " Where Month(MonthYear)=" & mMonth & " And Year(MonthYear) = " & mYear & " AND SecurityAgencyID=" & mSecurityAgencyID & " AND LocationID= " & mLocationID & ""
Call TmpTable
If Not TmpRs.EOF Then
    Edit_Flg = "E"
    TxtTxnID = TmpRs!MonthWiseTxnID
    TxtTxnDate.Value = TmpRs!MonthwiseTxnDate
    TxtCreated = IIf(IsNull(TmpRs!CreatedBy), "", TmpRs!CreatedBy) & " :- " & IIf(IsNull(TmpRs!CreatedDate), "", TmpRs!CreatedDate)
    TxtUpdated = IIf(IsNull(TmpRs!UpdatedBy), "", TmpRs!UpdatedBy) & " :- " & IIf(IsNull(TmpRs!UpdatedDate), "", TmpRs!UpdatedDate)
    Call EditCmd_Click
    CmbMonth.Visible = False
    TxtMonth.Visible = True
    CmbYear.Visible = False
    TxtYear.Visible = True
    CmbSecurityAgency.Visible = False
    CmdSecurityAgency.Visible = False
    TxtSecurityAgency.Visible = True
    CmbLocationID.Visible = False
    CmdLocationID.Visible = False
    TxtLocation.Visible = True
    'Call GetSecLocMapping
Else
'    TxtTxnID.Text = ""
'    TxtTxnDate.Value = Date
'    TxtCreated = ""
'    TxtUpdated = ""
'    CmbMonth.Visible = True
'    TxtMonth.Visible = False
'    CmbYear.Visible = False
'    TxtYear.Visible = True
'    CmbSecurityAgency.Visible = False
'    CmdSecurityAgency.Visible = False
'    TxtSecurityAgency.Visible = True
'    CmbLocationID.Visible = False
'    CmdLocationID.Visible = False
'    TxtLocation.Visible = True
    Edit_Flg = "A"
    Call GetSecLocMapping
    Exit Sub
End If

Call GetSecLocMapping

tmp = " SELECT TMSL.MonthwiseTxnDetailID, "
tmp = tmp & " ME.EmployeeName + '~'+ ME.EmployeeNo As EmpName, CMPMID, AttendanceSheetID,"
tmp = tmp & " NoofDays, EstimatedAmt, EstimatedSTax,TotEstimatedAmt, Remarks, SecLocMapDetailID,BillingNoOfDays,BillingAmount,BillingSTax,BillingTotalAmount"
tmp = tmp & " FROM Txn_MonthwiseSecLocDetail TMSL"
tmp = tmp & " Left Join Mst_Employee ME On ME.EmployeeID = TMSL.CMPMID"
tmp = tmp & " Where MonthwiseTxnID = " & TxtTxnID & " ORDER BY MonthwiseTxnDetailID"

Call TmpTable

If TmpRs.RecordCount > 0 Then
    For I = 2 To MSHFlexGrid2.Rows - 2
        TmpRs.MoveFirst
        TmpRs.Find "SecLocMapDetailID = " & MSHFlexGrid2.TextMatrix(I, 17)
        If Not TmpRs.EOF Then
           MSHFlexGrid2.TextMatrix(I, 8) = IIf(IsNull(TmpRs!EmpName), "", TmpRs!EmpName)
           MSHFlexGrid2.TextMatrix(I, 9) = IIf(IsNull(TmpRs!CMPMID), "", TmpRs!CMPMID)
           MSHFlexGrid2.TextMatrix(I, 10) = TmpRs!AttendanceSheetID
           MSHFlexGrid2.TextMatrix(I, 11) = TmpRs!NoOfDays
           MSHFlexGrid2.TextMatrix(I, 12) = TmpRs!EstimatedAmt
           MSHFlexGrid2.TextMatrix(I, 13) = TmpRs!EstimatedSTax
           MSHFlexGrid2.TextMatrix(I, 14) = TmpRs!TotEstimatedAmt
           MSHFlexGrid2.TextMatrix(I, 15) = TmpRs!Remarks
           MSHFlexGrid2.TextMatrix(I, 16) = TmpRs!MonthwiseTxnDetailID
           MSHFlexGrid2.TextMatrix(I, 19) = IIf(IsNull(TmpRs!BillingNoOfDays), 0, TmpRs!BillingNoOfDays)
           MSHFlexGrid2.TextMatrix(I, 20) = IIf(IsNull(TmpRs!BillingAmount), 0, TmpRs!BillingAmount)
           MSHFlexGrid2.TextMatrix(I, 21) = IIf(IsNull(TmpRs!BillingSTax), 0, TmpRs!BillingSTax)
           MSHFlexGrid2.TextMatrix(I, 22) = IIf(IsNull(TmpRs!BillingTotalAmount), 0, TmpRs!BillingTotalAmount)
         End If
         MSHFlexGrid2.row = I
         For x = 19 To 22
             MSHFlexGrid2.Col = x
             MSHFlexGrid2.CellBackColor = vbYellow
         Next
    Next
End If
End Sub
Function CheckDuplicateAttenSheetID() As Boolean
Dim Retval As Boolean
Retval = False
If MSHFlexGrid2.Rows > 1 Then
   For u = 2 To MSHFlexGrid2.Rows - 1
       If MSHFlexGrid2.TextMatrix(u, 10) = TxtAttenSheetID.Text Then
          If LCase(CmdAdd2List.Caption) = "update list" Then
             If u <> MSHFlexGrid2.RowSel Then
                Retval = True
                Exit For
             End If
          End If
       End If
   Next
End If

CheckDuplicateAttenSheetID = Retval
End Function

