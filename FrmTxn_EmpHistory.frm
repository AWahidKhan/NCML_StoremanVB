VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmTxn_EmpHistory 
   Caption         =   "Monthwise Employee History"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.VScrollBar VScroll1 
      Height          =   9015
      Left            =   14880
      TabIndex        =   75
      Top             =   0
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      TabIndex        =   74
      Top             =   10320
      Width           =   10935
   End
   Begin VB.Frame Frame4 
      Height          =   9975
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   15135
      Begin VB.Frame Frame0 
         Height          =   2400
         Left            =   120
         TabIndex        =   61
         Top             =   6960
         Visible         =   0   'False
         Width           =   14625
         Begin VB.ComboBox CmbSearchStatus 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            ItemData        =   "FrmTxn_EmpHistory.frx":0000
            Left            =   11220
            List            =   "FrmTxn_EmpHistory.frx":000A
            Sorted          =   -1  'True
            TabIndex        =   69
            Top             =   420
            Width           =   1470
         End
         Begin VB.TextBox TxtSEmpNo 
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
            Left            =   4050
            TabIndex        =   66
            Top             =   1140
            Width           =   1350
         End
         Begin VB.ComboBox CmbSDesignation 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   7695
            Sorted          =   -1  'True
            TabIndex        =   67
            Top             =   435
            Width           =   3435
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
            Left            =   225
            Sorted          =   -1  'True
            TabIndex        =   68
            Top             =   1125
            Width           =   3735
         End
         Begin VB.CommandButton CmdExcel 
            Caption         =   "Excel"
            Height          =   350
            Left            =   11415
            TabIndex        =   64
            Top             =   960
            Width           =   1245
         End
         Begin VB.TextBox TxtEmployeeName 
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
            Left            =   3360
            TabIndex        =   65
            Top             =   435
            Width           =   4140
         End
         Begin VB.CommandButton CmdGo 
            Caption         =   "Go"
            Height          =   345
            Left            =   10140
            TabIndex        =   24
            Top             =   960
            Width           =   1245
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   2220
            Left            =   240
            TabIndex        =   70
            Top             =   1560
            Width           =   6135
            _ExtentX        =   10821
            _ExtentY        =   3916
            _Version        =   393216
            Rows            =   3
            AllowUserResizing=   1
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSComCtl2.DTPicker TxtSTo 
            Height          =   375
            Left            =   1815
            TabIndex        =   63
            TabStop         =   0   'False
            Top             =   435
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
            Format          =   61341697
            CurrentDate     =   36494
         End
         Begin MSComCtl2.DTPicker TxtSFrom 
            Height          =   375
            Left            =   300
            TabIndex        =   62
            TabStop         =   0   'False
            Top             =   435
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
            Format          =   61341697
            CurrentDate     =   36494
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
            Left            =   525
            TabIndex        =   88
            Top             =   150
            Width           =   990
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
            Left            =   2040
            TabIndex        =   87
            Top             =   150
            Width           =   855
         End
         Begin VB.Label Label24 
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
            Left            =   11715
            TabIndex        =   86
            Top             =   165
            Width           =   555
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "Employee No "
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
            TabIndex        =   85
            Top             =   855
            Width           =   1290
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
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
            Height          =   240
            Left            =   8865
            TabIndex        =   84
            Top             =   165
            Width           =   1080
         End
         Begin VB.Label Label17 
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
            Left            =   1710
            TabIndex        =   83
            Top             =   855
            Width           =   765
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Employee Name "
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
            Left            =   4845
            TabIndex        =   73
            Top             =   150
            Width           =   1575
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
            TabIndex        =   72
            Top             =   1095
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
            Left            =   6405
            TabIndex        =   71
            Top             =   1095
            Width           =   2415
         End
      End
      Begin VB.Frame ButtonFrm 
         Height          =   1020
         Left            =   120
         TabIndex        =   53
         Top             =   5880
         Width           =   12720
         Begin VB.CommandButton ExitCmd 
            Caption         =   "E&xit"
            Height          =   375
            Left            =   9945
            TabIndex        =   21
            Top             =   540
            Width           =   2415
         End
         Begin VB.CommandButton LastCmd 
            Caption         =   "&Last"
            Height          =   375
            Left            =   7530
            TabIndex        =   60
            Top             =   540
            Width           =   2415
         End
         Begin VB.CommandButton FirstCmd 
            Caption         =   "&First"
            Height          =   375
            Left            =   4995
            TabIndex        =   59
            Top             =   540
            Width           =   2535
         End
         Begin VB.CommandButton PreviousCmd 
            Caption         =   "&Previous"
            Height          =   375
            Left            =   2580
            TabIndex        =   58
            Top             =   540
            Width           =   2415
         End
         Begin VB.CommandButton NextCmd 
            Caption         =   "&Next"
            Height          =   375
            Left            =   165
            TabIndex        =   57
            Top             =   540
            Width           =   2415
         End
         Begin VB.CommandButton SearchCmd 
            Caption         =   "&Search"
            Height          =   375
            Left            =   9945
            TabIndex        =   22
            Top             =   165
            Width           =   2415
         End
         Begin VB.CommandButton DeleteCmd 
            Caption         =   "&Delete"
            Height          =   375
            Left            =   7530
            TabIndex        =   56
            Top             =   165
            Width           =   2415
         End
         Begin VB.CommandButton EditCmd 
            Caption         =   "&Edit"
            Height          =   375
            Left            =   4980
            TabIndex        =   55
            Top             =   165
            Width           =   2535
         End
         Begin VB.CommandButton SaveCmd 
            Caption         =   "Save"
            Height          =   375
            Left            =   2580
            TabIndex        =   20
            Top             =   165
            Width           =   2415
         End
         Begin VB.CommandButton AddCmd 
            Caption         =   "&Add New "
            Height          =   375
            Left            =   165
            TabIndex        =   54
            Top             =   165
            Width           =   2415
         End
      End
      Begin VB.Frame Frame1 
         Height          =   5700
         Left            =   120
         TabIndex        =   23
         Top             =   120
         Width           =   12720
         Begin VB.TextBox TxtAddress 
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
            Left            =   2280
            MaxLength       =   250
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   12
            Top             =   3240
            Width           =   10215
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
            Left            =   7080
            Sorted          =   -1  'True
            TabIndex        =   4
            Top             =   1095
            Width           =   5385
         End
         Begin VB.TextBox TxtCTC 
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
            Left            =   2280
            TabIndex        =   16
            Top             =   4290
            Visible         =   0   'False
            Width           =   2940
         End
         Begin VB.TextBox TxtGrossPay 
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
            Left            =   6570
            TabIndex        =   17
            Top             =   4275
            Visible         =   0   'False
            Width           =   2805
         End
         Begin VB.ComboBox CmbAdditionalReporting 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   11160
            Sorted          =   -1  'True
            TabIndex        =   26
            Top             =   540
            Visible         =   0   'False
            Width           =   705
         End
         Begin VB.TextBox TxtMonthEmpID 
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
            Left            =   2280
            Locked          =   -1  'True
            TabIndex        =   77
            TabStop         =   0   'False
            Top             =   165
            Width           =   2925
         End
         Begin VB.TextBox TxtEmpEmailID 
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
            Left            =   2280
            TabIndex        =   10
            Top             =   2820
            Width           =   6285
         End
         Begin VB.ComboBox CmbConnectedRAG 
            BeginProperty Font 
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
            Sorted          =   -1  'True
            TabIndex        =   9
            Top             =   2407
            Width           =   10020
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
            Left            =   2280
            Locked          =   -1  'True
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   5070
            Width           =   4095
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
            Left            =   8400
            Locked          =   -1  'True
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   5070
            Width           =   4095
         End
         Begin VB.TextBox TxtEmpID 
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
            Left            =   2280
            Locked          =   -1  'True
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   615
            Width           =   2925
         End
         Begin VB.TextBox TxtEmpNo 
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
            Left            =   2280
            TabIndex        =   3
            Top             =   1095
            Width           =   2940
         End
         Begin VB.ComboBox CmbDesignationID 
            BeginProperty Font 
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
            TabIndex        =   6
            Top             =   1515
            Width           =   2940
         End
         Begin VB.TextBox TxtEmpName 
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
            Left            =   7080
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   1095
            Width           =   5385
         End
         Begin VB.TextBox TxtAdditionalReporting 
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
            Left            =   11070
            Locked          =   -1  'True
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   525
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.CommandButton CmdDesignationID 
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
            Left            =   5280
            TabIndex        =   27
            Top             =   1530
            Width           =   375
         End
         Begin VB.ComboBox CmbDirectReporting 
            BeginProperty Font 
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
            Sorted          =   -1  'True
            TabIndex        =   8
            Top             =   1995
            Width           =   10230
         End
         Begin VB.ComboBox CmbFlag 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            ItemData        =   "FrmTxn_EmpHistory.frx":0021
            Left            =   7095
            List            =   "FrmTxn_EmpHistory.frx":002B
            Sorted          =   -1  'True
            TabIndex        =   7
            Top             =   1515
            Width           =   2295
         End
         Begin VB.TextBox TxtMobileNo 
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
            Left            =   9840
            TabIndex        =   11
            Top             =   2820
            Width           =   2415
         End
         Begin VB.CheckBox ChkSelfReporting 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Check1"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   11760
            TabIndex        =   25
            Top             =   1575
            Visible         =   0   'False
            Width           =   225
         End
         Begin VB.CheckBox ChkAllocateExpense 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2265
            TabIndex        =   18
            Top             =   4785
            Width           =   225
         End
         Begin VB.CheckBox ChkCMPAllocation 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   4905
            TabIndex        =   19
            Top             =   4785
            Width           =   225
         End
         Begin VB.ComboBox CmbBaseLocationID 
            BeginProperty Font 
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
            Sorted          =   -1  'True
            TabIndex        =   13
            Top             =   3855
            Width           =   7110
         End
         Begin VB.CommandButton CmdBaseLocationID 
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
            Left            =   9450
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   3855
            Width           =   375
         End
         Begin VB.TextBox TxtDesignationID 
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
            Left            =   2280
            Locked          =   -1  'True
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   1508
            Width           =   2940
         End
         Begin VB.TextBox TxtFlag 
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
            Left            =   7095
            Locked          =   -1  'True
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   1515
            Width           =   2280
         End
         Begin VB.TextBox TxtDirectReporting 
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
            Left            =   2280
            Locked          =   -1  'True
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   1988
            Width           =   10215
         End
         Begin VB.TextBox TxtConnectedRAG 
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
            Left            =   2280
            Locked          =   -1  'True
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   2400
            Width           =   10215
         End
         Begin VB.TextBox TxtBaseLocationID 
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
            TabIndex        =   76
            TabStop         =   0   'False
            Top             =   3855
            Visible         =   0   'False
            Width           =   7110
         End
         Begin MSComCtl2.DTPicker TxtDOJ 
            Height          =   390
            Left            =   7080
            TabIndex        =   2
            Top             =   615
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   688
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
            Format          =   61341697
            CurrentDate     =   39884
         End
         Begin MSComCtl2.DTPicker TxtTxnDate 
            Height          =   390
            Left            =   7080
            TabIndex        =   1
            Top             =   165
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   688
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
            Format          =   61341697
            CurrentDate     =   39884
         End
         Begin MSComCtl2.DTPicker TxtDOL 
            Height          =   390
            Left            =   10815
            TabIndex        =   15
            Top             =   3840
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   688
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
            Format          =   61341697
            CurrentDate     =   39884
         End
         Begin VB.Label LblDOL 
            Caption         =   "D.O.L."
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
            Left            =   10080
            TabIndex        =   89
            Top             =   3900
            Width           =   660
         End
         Begin VB.Label Label16 
            Caption         =   "CTC"
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
            Left            =   105
            TabIndex        =   82
            Top             =   4350
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.Label Label15 
            Caption         =   "Gross Pay"
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
            Left            =   5385
            TabIndex        =   81
            Top             =   4350
            Visible         =   0   'False
            Width           =   1125
         End
         Begin VB.Label lblDOJ 
            Alignment       =   1  'Right Justify
            Caption         =   "Joining Date"
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
            Left            =   5820
            TabIndex        =   80
            Top             =   660
            Width           =   1215
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
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
            Left            =   5745
            TabIndex        =   79
            Top             =   210
            Width           =   1290
         End
         Begin VB.Label Label7 
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
            Height          =   285
            Left            =   120
            TabIndex        =   78
            Top             =   210
            Width           =   1695
         End
         Begin VB.Label Label13 
            Caption         =   "Connected RAG"
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
            TabIndex        =   52
            Top             =   2445
            Width           =   2055
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
            Left            =   135
            TabIndex        =   51
            Top             =   5130
            Width           =   1695
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
            Left            =   7365
            TabIndex        =   50
            Top             =   5130
            Width           =   975
         End
         Begin VB.Label Label3 
            Caption         =   "Employee Name"
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
            Left            =   5280
            TabIndex        =   49
            Top             =   1155
            Width           =   1815
         End
         Begin VB.Label Label8 
            Caption         =   "Designation"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   48
            Top             =   1515
            Width           =   1695
         End
         Begin VB.Label Label2 
            Caption         =   "Employee No"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   47
            Top             =   1095
            Width           =   1695
         End
         Begin VB.Label Label1 
            Caption         =   "Employee ID"
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
            TabIndex        =   46
            Top             =   615
            Width           =   1695
         End
         Begin VB.Label Label4 
            Caption         =   "Direct Reportee"
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
            TabIndex        =   45
            Top             =   2025
            Width           =   2055
         End
         Begin VB.Label Label5 
            Caption         =   "Connected DRC"
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
            Left            =   10170
            TabIndex        =   44
            Top             =   225
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.Label Label6 
            Caption         =   "Employee EmailID"
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
            TabIndex        =   43
            Top             =   2850
            Width           =   1695
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
            Height          =   285
            Left            =   6360
            TabIndex        =   42
            Top             =   1530
            Width           =   735
         End
         Begin VB.Label Label20 
            Caption         =   "Mobile No"
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
            Left            =   8640
            TabIndex        =   41
            Top             =   2857
            Width           =   975
         End
         Begin VB.Label Label21 
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
            Height          =   300
            Left            =   120
            TabIndex        =   40
            Top             =   3375
            Width           =   1695
         End
         Begin VB.Label Label25 
            Caption         =   "CMPM-cum-DRC"
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
            Left            =   10080
            TabIndex        =   39
            Top             =   1575
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.Label Label26 
            Caption         =   "Allocate Expense"
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
            TabIndex        =   38
            Top             =   4785
            Width           =   1695
         End
         Begin VB.Label Label30 
            Caption         =   "CMP Allocation"
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
            Left            =   3150
            TabIndex        =   37
            Top             =   4785
            Width           =   1695
         End
         Begin VB.Label LblLocationID 
            Caption         =   "Base Location"
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
            Left            =   105
            TabIndex        =   36
            Top             =   3915
            Width           =   2010
         End
      End
   End
End
Attribute VB_Name = "FrmTxn_EmpHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg, mDesignationID, mFlag, mDirectReportingID, mOldDesignationID, mAdditionalReportingID, mConnectedRAGID As Variant
Dim mAssetTypeId, mAssetStatusID, mEmployeeID, mSM_Prefix, mReportingDesignationID As Variant
Dim mLocationID, mSLocationId, mSDesignationID, mCMPID, mBaseLocationID As Double
Dim mTotalDays  As Variant

Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)

mTotalDays = Day(DateSerial(Year(Date), Month(Date) + 1, 0))
TxtTxnDate.Value = (Year(Date) & "-" & Month(Date) & "-" & mTotalDays)

Frame1.Enabled = True
TxtDOJ.Value = Date
CmbEmployeeID.Visible = True

CmbDesignationID.Visible = True
CmbDesignationID.Enabled = True
CmbDirectReporting.Visible = True
'CmbAdditionalReporting.Visible = True
CmbConnectedRAG.Visible = True
'CmbAdditionalReporting.Enabled = False
CmbConnectedRAG.Enabled = False
CmbFlag.Visible = True
'CmbCMPID.Enabled = True
'CmbLocationID.Enabled = True
''Frame2.Enabled = False
''Frame3.Enabled = True

Label25.Visible = False
ChkSelfReporting.Visible = False
CmbBaseLocationID.Visible = True

TxtEmpName.Visible = False
TxtDesignationID.Visible = False
TxtDirectReporting.Visible = False
TxtAdditionalReporting.Visible = False
TxtConnectedRAG.Visible = False
TxtFlag.Visible = False
TxtBaseLocationID.Visible = False

CmdBaseLocationID.Visible = True
CmdDesignationID.Visible = True
''CmdLocationID.Visible = True
''CmdCMPID.Visible = True
''cmdAssetStatus.Visible = True
''CmdAssetType.Visible = True
ChkAllocateExpense.Enabled = True
ChkAllocateExpense.Value = 0
ChkCMPAllocation.Enabled = False
ChkCMPAllocation.Value = 1
TxtEmpNo.Locked = False
'TxtEmpName.Locked = False
TxtEmpEmailID.Locked = False
TxtMobileNo.Locked = False
TxtAddress.Locked = False
TxtCreated = ""
TxtUpdated = ""
TxtEmpID = ""
TxtEmpNo = ""
TxtEmpName = ""
TxtEmpEmailID = ""
TxtMobileNo = ""
TxtAddress = ""
TxtCTC.Text = ""
TxtGrossPay.Text = ""
'Frame2.Enabled = True
'Frame3.Enabled = True
mDirectReportingID = Null
mAdditionalReportingID = Null
mFlag = Null
CmbDesignationID.Text = ""
CmbDirectReporting.Text = ""
'CmbAdditionalReporting.Text = ""
CmbConnectedRAG.Text = ""
CmbBaseLocationID.Text = ""
CmbFlag.Text = ""
LblDOL.Visible = False
TxtDOL.Visible = False
''CmbCMPID.Text = ""
''CmbAssetType.Text = ""
''CmbAssetStatus.Text = ""
ChkSelfReporting.Enabled = True
''CmdAddToGrid.Caption = "Add To Grid"
''CmdAdd2GridAsset.Caption = "Add To Grid"
''AssetVerifiedon.Value = Date
''TxtAMCExpDate.Value = Date
TxtAssetDescription = ""
TxtAssetQty = ""
TxtRemarks = ""
TxtAMCAmount = ""

''SSTab1.Tab = 0
'Call Grid_Head2

End Sub



Private Sub AssetVerifiedon_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "AssetVerifiedon")
End If
End Sub

Private Sub ChkSelfReporting_Click()
If SaveCmd.Enabled = False Then Exit Sub
'If ChkSelfReporting.Value = 0 Then
'   CmbAdditionalReporting.Enabled = True
'   CmbAdditionalReporting.SetFocus
'Else
'   CmbAdditionalReporting.Enabled = False
'
'End If
'CmbAdditionalReporting.Text = ""
mAdditionalReportingID = Null
End Sub

Private Sub CmbAdditionalReporting_GotFocus()
tmp = CmbAdditionalReporting.Text
'Call Mst_EmployeeTable("Where Flag = 'Y'")
RsTxn_MonthEmp_History.Filter = "DesignationID =  " & 5
CmbAdditionalReporting.Clear
If RsTxn_MonthEmp_History.RecordCount = 0 Then
   MsgBox "No Employee Found!!!"
   CmbAdditionalReporting.SetFocus
   RsTxn_MonthEmp_History.Filter = "DesignationId <> Null"
   'RsTxn_MonthEmp_History.MoveFirst
   Exit Sub
End If
For I = 1 To RsTxn_MonthEmp_History.RecordCount
    CmbAdditionalReporting.AddItem RsTxn_MonthEmp_History!EmployeeName & "~" & RsTxn_MonthEmp_History!EmployeeNo
    RsTxn_MonthEmp_History.MoveNext
Next
CmbAdditionalReporting.Text = tmp
End Sub

Private Sub CmbAdditionalReporting_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "CmbAdditionalReporting")
End If
End Sub

Private Sub CmbAdditionalReporting_LostFocus()
Dim strSplitString() As String
Dim EmpNo_, EmpName_ As String

If CmbAdditionalReporting.Text = "" Then
   mAdditionalReportingID = Null
   RsTxn_MonthEmp_History.Filter = "DesignationId <> Null"
   Exit Sub
End If

strSplitString = Split(CmbAdditionalReporting, "~")
EmpName_ = Trim(strSplitString(0))
EmpNo_ = Trim(strSplitString(1))

RsTxn_MonthEmp_History.MoveFirst

RsTxn_MonthEmp_History.Find "EmployeeNo = '" & EmpNo_ & "'"
If RsTxn_MonthEmp_History.EOF Then
   MsgBox "Invalid Selection "
   CmbAdditionalReporting.SetFocus
   Exit Sub
End If
mAdditionalReportingID = RsTxn_MonthEmp_History!EmployeeID
RsTxn_MonthEmp_History.Filter = "DesignationId <> Null"
RsTxn_MonthEmp_History.MoveFirst
End Sub

Private Sub CmbAssetStatus_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "CmbAssetStatus")
End If
End Sub

Private Sub CmbAssetType_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "CmbAssetType")
End If
End Sub



Private Sub CmbBaseLocationID_GotFocus()
tmp = CmbBaseLocationID.Text
Call TableMst_Location("")
CmbBaseLocationID.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Location found"
   CmbBaseLocationID.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbBaseLocationID.AddItem RsMst_Location!LocationName
    RsMst_Location.MoveNext
Next
CmbBaseLocationID.Text = tmp

End Sub

Private Sub CmbBaseLocationID_LostFocus()
If CmbBaseLocationID.Text = "" Then
   mBaseLocationID = 0
   Exit Sub
End If

RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbBaseLocationID.Text & "'"

If RsMst_Location.EOF Then
   MsgBox "Invalid Location  selection "
   CmbBaseLocationID.SetFocus
   Exit Sub
End If
mBaseLocationID = RsMst_Location!LocationID

End Sub

Private Sub CmbConnectedRAG_GotFocus()

temp = CmbConnectedRAG.Text
'Call Mst_EmployeeTable("Where Flag = 'Y'")
RsTxn_MonthEmp_History.Filter = "DesignationId = " & 7
CmbConnectedRAG.Clear
If RsTxn_MonthEmp_History.RecordCount = 0 Then
   MsgBox "No Employee Found!!!"
   CmbConnectedRAG.SetFocus
   RsTxn_MonthEmp_History.Filter = "DesignationId <> Null"
   'RsTxn_MonthEmp_History.MoveFirst
   Exit Sub
End If
For I = 1 To RsTxn_MonthEmp_History.RecordCount
    CmbConnectedRAG.AddItem RsTxn_MonthEmp_History!EmployeeName & "~" & RsTxn_MonthEmp_History!EmployeeNo
    RsTxn_MonthEmp_History.MoveNext
Next
CmbConnectedRAG.Text = temp

End Sub

Private Sub CmbConnectedRAG_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "CmbConnectedRAG")
End If
End Sub

Private Sub CmbConnectedRAG_LostFocus()
Dim strSplitString() As String
Dim EmpNo_, EmpName_ As String

If CmbConnectedRAG.Text = "" Then
   mConnectedRAGID = Null
   RsTxn_MonthEmp_History.Filter = "DesignationId <> Null"
   Exit Sub
End If

strSplitString = Split(CmbConnectedRAG, "~")
EmpName_ = Trim(strSplitString(0))
EmpNo_ = Trim(strSplitString(1))


RsTxn_MonthEmp_History.MoveFirst
RsTxn_MonthEmp_History.Find "EmployeeNo = '" & EmpNo_ & "'"

If RsTxn_MonthEmp_History.EOF Then
   MsgBox "Invalid Selection "
   CmbConnectedRAG.SetFocus
   Exit Sub
End If
mConnectedRAGID = RsTxn_MonthEmp_History!EmployeeID
RsTxn_MonthEmp_History.Filter = "DesignationId <> Null"
RsTxn_MonthEmp_History.MoveFirst
End Sub

Private Sub CmbDesignationID_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "CmbDesignationID")
End If
End Sub

Private Sub CmbDirectReporting_GotFocus()
tmp = CmbDirectReporting.Text
'Call Mst_EmployeeTable("Where Flag = 'Y'")

'Commented on 07 Aug 2010 Parag
'If mDesignationID = 1 Then
'   RsTxn_MonthEmp_History.Filter = "DesignationId = " & 1
'ElseIf mDesignationID = 2 Then
'   RsTxn_MonthEmp_History.Filter = "DesignationId = " & 1
'ElseIf mDesignationID = 3 Then
'   RsTxn_MonthEmp_History.Filter = "DesignationId = " & 2
'ElseIf mDesignationID = 4 Then
'   RsTxn_MonthEmp_History.Filter = "DesignationId = " & 3
'ElseIf mDesignationID = 5 Then
'   RsTxn_MonthEmp_History.Filter = "DesignationId = " & 6
'ElseIf mDesignationID = 6 Then
'   'RsTxn_MonthEmp_History.Filter = "DesignationId =  4 Or DesignationID = 3 "
'   RsTxn_MonthEmp_History.Filter = " DesignationID = 3 "
'ElseIf mDesignationID = 7 Then
'   RsTxn_MonthEmp_History.Filter = "DesignationId = " & 1
'ElseIf mDesignationID = 8 Then                   'QQS
'   RsTxn_MonthEmp_History.Filter = "DesignationId =  " & 6
'ElseIf mDesignationID = 9 Then                   'BDO
'   'RsTxn_MonthEmp_History.Filter = "DesignationId = 2 Or DesignationID = 3 "
'   RsTxn_MonthEmp_History.Filter = "DesignationID = 3 "
'ElseIf mDesignationID = 10 Then                   'BDO
'   'RsTxn_MonthEmp_History.Filter = "DesignationId = 2 Or DesignationID = 3 "
'   RsTxn_MonthEmp_History.Filter = "DesignationID = 3 "
'End If

If CmbDesignationID.Text = "" Then
   MsgBox "Select Designation!!!"
   CmbDesignationID.SetFocus
   Exit Sub
End If


If IsNull(mReportingDesignationID) = False Then
   RsMst_Employee.Filter = "DesignationID = " & mReportingDesignationID & " "
Else
   RsMst_Employee.Filter = "DesignationID = " & mDesignationID & " "
End If

CmbDirectReporting.Clear
If RsMst_Employee.RecordCount = 0 Then
   MsgBox "No Employee Found!!!"
   CmbDirectReporting.SetFocus
   RsTxn_MonthEmp_History.Filter = "DesignationId <> Null"
   Exit Sub
End If
For I = 1 To RsMst_Employee.RecordCount
    CmbDirectReporting.AddItem RsMst_Employee!EmployeeName & "~" & RsMst_Employee!EmployeeNo
    RsMst_Employee.MoveNext
Next
CmbDirectReporting.Text = tmp
End Sub

Private Sub CmbDirectReporting_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "CmbDirectReporting")
End If
End Sub

Private Sub CmbDirectReporting_LostFocus()
Dim strSplitString() As String
Dim EmpNo_, EmpName_ As String

If CmbDirectReporting.Text = "" Then
   mDirectReportingID = Null
   RsMst_Employee.Filter = "DesignationId <> Null"
   Exit Sub
End If

strSplitString = Split(CmbDirectReporting, "~")
If UBound(strSplitString) <> 1 Then
   MsgBox "Invalid Selection "
   CmbDirectReporting.SetFocus
   Exit Sub
End If

EmpName_ = Trim(strSplitString(0))
EmpNo_ = Trim(strSplitString(1))


RsMst_Employee.MoveFirst

RsMst_Employee.Find "EmployeeNo = '" & EmpNo_ & "'"
If RsMst_Employee.EOF Then
   MsgBox "Invalid Selection "
   CmbDirectReporting.SetFocus
   Exit Sub
End If
mDirectReportingID = RsMst_Employee!EmployeeID
RsMst_Employee.Filter = "DesignationId <> Null"
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
Call FillDetails(EmpNo_)
End Sub

Private Sub CmbFlag_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "CmbFlag")
End If
End Sub

Private Sub CmbFlag_LostFocus()
If CmbFlag.Text = "" Then Exit Sub
If CmbFlag.Text = "Active" Then
   mFlag = "Y"
   LblDOL.Visible = False
   TxtDOL.Visible = False
ElseIf CmbFlag.Text = "De-Active" Then
   mFlag = "N"
   LblDOL.Visible = True
   TxtDOL.Visible = True
   TxtDOL.Value = Date
Else
   MsgBox "Invalid Selection!!!"
   CmbFlag.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmbCMPID_GotFocus()
If CmbLocationID.Text = "" Then Exit Sub
temp = CmbCMPID.Text
'tmp = " Select mg.GodownID,mg.GodownNO,mg.Address,mc.CMPID " & _
'      " From Mst_Godown mg " & _
'      " Inner Join Mst_CMP mc on mg.CMPID = mc.CMPID " & _
'      " Where mc.locationID = " & mLocationId & " and mg.CloseDate is null "

tmp = "Select * from Mst_CMP Where LocationID = " & mLocationID & " And CloseDate is null"
Call Tmp4Table

CmbCMPID.Clear

If TmpRs4.RecordCount = 0 Then
   MsgBox "No CMP Found in Location selected!!!"
   CmbCMPID.SetFocus
   Exit Sub
End If

For I = 1 To TmpRs4.RecordCount
    CmbCMPID.AddItem TmpRs4!CMPName
    TmpRs4.MoveNext
Next

CmbCMPID.Text = temp

End Sub

Private Sub CmbCMPID_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "CmbCMPID")
End If
End Sub

Private Sub CmbCMPID_LostFocus()
If CmbCMPID.Text = "" Then
   Exit Sub
End If

TmpRs4.MoveFirst
TmpRs4.Find "CMPName = '" & CmbCMPID.Text & "'"

If TmpRs4.EOF Then
   MsgBox "Invalid CMP Selection "
   CmbCMPID.SetFocus
   Exit Sub
End If
mCMPID = TmpRs4!CMPID

End Sub

Private Sub CmbLocationID_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "CmbLocationID")
End If
End Sub

Private Sub CmbSDesignation_GotFocus()
tmp = CmbSDesignation.Text
Call TableMst_Designation("")
CmbSDesignation.Clear
If RsMst_Designation.RecordCount = 0 Then
   MsgBox "No Designation found"
   CmbSDesignation.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Designation.RecordCount
    CmbSDesignation.AddItem RsMst_Designation!Designation
    RsMst_Designation.MoveNext
Next
CmbSDesignation.Text = tmp
End Sub

Private Sub CmbSDesignation_LostFocus()
If CmbSDesignation.Text = "" Then
   mSDesignationID = Null
   Exit Sub
End If

RsMst_Designation.MoveFirst
RsMst_Designation.Find "Designation = '" & CmbSDesignation.Text & "'"

If RsMst_Designation.EOF Then
   MsgBox "Invalid Designation Selection "
   CmbSDesignation.SetFocus
   Exit Sub
End If
mSDesignationID = RsMst_Designation!DesignationID
End Sub

Private Sub CmbSearchStatus_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "CmbSearchStatus")
End If
End Sub

Private Sub CmbSearchStatus_LostFocus()
If CmbSearchStatus.Text = "" Then Exit Sub

If CmbSearchStatus.Text = "Active" Or CmbSearchStatus.Text = "De-Active" Then
   
Else
   MsgBox "Invalid selection!!!"
   CmbSearchStatus.SetFocus
   Exit Sub

End If

'Dim x As Integer
'If CmbSearchStatus.Text = "" Then
'   Call Grid_Head1
'   Exit Sub
'End If
'
'If CmbSearchStatus.Text = "Active" Then
'   mFlag = "Y"
'
'ElseIf CmbSearchStatus.Text = "De-Active" Then
'   mFlag = "N"
'Else
'   MsgBox "Invalid Selection!!!"
'   CmbSearchStatus.SetFocus
'   Exit Sub
'End If
'tmp = "Select * from Mst_Employee where Flag = '" & mFlag & "'"
'Call Tmp1Table
'If TmpRs1.RecordCount = 0 Then
'      MsgBox "No such Record Found!!!"
'      Call Grid_Head1
'      Exit Sub
'End If
'MSHFlexGrid1.Rows = TmpRs1.RecordCount + 1
'For x = 1 To TmpRs1.RecordCount
'    Call ShowGrid(x)
'    TmpRs1.MoveNext
'Next
End Sub

Private Sub CmdAdd2GridAsset_Click()
If CmbAssetType.Text = "" Then
   MsgBox "Blank Asset Type not allowed!!!"
   CmbAssetType.SetFocus
   Exit Sub
End If
If CmbAssetStatus.Text = "" Then
   MsgBox "Blank Asset Status not allowed!!!"
   CmbAssetStatus.SetFocus
   Exit Sub
End If

If AssetVerifiedon.Value > Date Then
   MsgBox "Future Verified date not allowed!!!"
   AssetVerifiedon.SetFocus
   Exit Sub
End If

If TxtAMCExpDate.Value < Date Then
   MsgBox "Back AMC Expiry date not allowed!!!"
   TxtAMCExpDate.SetFocus
   Exit Sub
End If

'If TxtAssetDescription = "" Then
'   MsgBox "Blank Description Not Allowed!!!"
'   TxtAssetDescription.SetFocus
'   Exit Sub
'End If

Dim I, ans As Variant

If CmdAdd2GridAsset.Caption = "Add To Grid" Then
   I = MSHFlexGrid3.Rows - 1
   MSHFlexGrid3.Rows = MSHFlexGrid3.Rows + 1
Else
   I = MSHFlexGrid3.RowSel
End If

MSHFlexGrid3.TextMatrix(I, 0) = CmbAssetType.Text
MSHFlexGrid3.TextMatrix(I, 1) = CmbAssetStatus.Text
MSHFlexGrid3.TextMatrix(I, 2) = Val(TxtAssetQty)
MSHFlexGrid3.TextMatrix(I, 3) = TxtAssetDescription
MSHFlexGrid3.TextMatrix(I, 4) = AssetVerifiedon.Value
MSHFlexGrid3.TextMatrix(I, 6) = TxtRemarks
MSHFlexGrid3.TextMatrix(I, 7) = TxtAMCExpDate.Value
MSHFlexGrid3.TextMatrix(I, 8) = Val(TxtAMCAmount)

CmdAdd2GridAsset.Caption = "Add To Grid"
CmbAssetType.Text = ""
CmbAssetStatus.Text = ""
TxtAssetQty = ""
TxtAssetDescription = ""
AssetVerifiedon.Value = Date
TxtRemarks = ""
TxtAMCExpDate.Value = Date
TxtAMCAmount = ""

End Sub

Private Sub CmdAddAll2grid_Click()
Dim I, ans As Variant
If CmbLocationID.Text = "" Then
   MsgBox "Blank Location not allowed!!!"
   CmbLocationID.SetFocus
   Exit Sub
End If
Call CmbCMPID_GotFocus
If TmpRs4.RecordCount > 0 Then
   TmpRs4.MoveFirst
   tmp = ""
   For iu = 0 To TmpRs4.RecordCount - 1
       mCMPID = TmpRs4!CMPID
       ans = Validate_DuplicateGridForAll()
       If ans = True Then
          Call FillGrid
       End If
       TmpRs4.MoveNext
   Next
End If
End Sub
Private Sub FillGrid()
Dim I
I = MSHFlexGrid2.Rows - 1
MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
MSHFlexGrid2.TextMatrix(I, 0) = CmbLocationID.Text
MSHFlexGrid2.TextMatrix(I, 1) = mCMPID
MSHFlexGrid2.TextMatrix(I, 2) = TmpRs4!CMPName
MSHFlexGrid2.TextMatrix(I, 3) = TmpRs4!Address
MSHFlexGrid2.TextMatrix(I, 4) = mLocationID
MSHFlexGrid2.TextMatrix(I, 6) = mSM_Prefix

End Sub
Private Sub CmdAddToGrid_Click()

If CmbLocationID.Text = "" Then
   MsgBox "Blank Location Not Allowed!!!"
   CmbLocationID.SetFocus
   Exit Sub
End If

If CmbCMPID.Text = "" Then
   MsgBox "Blank CMP Not Allowed!!!"
   CmbCMPID.SetFocus
   Exit Sub
End If

Dim I, ans As Variant


'tmp = "Select CMPID,DesignationID from Mst_EMPCMPMapping "
'Call Tmp3Table
'If TmpRs3.RecordCount > 0 Then
'   If CheckCMPMapping = False Then
'      MsgBox "CMP Already mapped !!!! "
'      Exit Sub
'   End If
'End If


ans = Validate_DuplicateGrid()
If ans = False Then
   MsgBox "Duplicate Entry !!!! "
   Exit Sub
End If

If CmdAddToGrid.Caption = "Add To Grid" Then
   I = MSHFlexGrid2.Rows - 1
   MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
Else
   I = MSHFlexGrid2.RowSel
End If

MSHFlexGrid2.TextMatrix(I, 0) = CmbLocationID.Text
MSHFlexGrid2.TextMatrix(I, 1) = mCMPID
MSHFlexGrid2.TextMatrix(I, 2) = TmpRs4!CMPName
MSHFlexGrid2.TextMatrix(I, 3) = TmpRs4!Address
MSHFlexGrid2.TextMatrix(I, 4) = mLocationID
MSHFlexGrid2.TextMatrix(I, 6) = mSM_Prefix

CmdAddToGrid.Caption = "Add To Grid"
CmbCMPID.Text = ""
CmbCMPID.SetFocus
End Sub

Private Sub cmdAssetStatus_Click()
FrmMst_AssetStatus.Show
End Sub

Private Sub CmdAssetType_Click()
FrmMst_AssetType.Show
End Sub

Private Sub CmbSLocation_GotFocus()
tmp = CmbSLocation.Text
Call TableMst_Location("")
CmbSLocation.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Location found"
   CmbSLocation.SetFocus
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
   mSLocationId = 0
   Exit Sub
End If

RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbSLocation.Text & "'"

If RsMst_Location.EOF Then
   MsgBox "Invalid Location  selection "
   CmbSLocation.SetFocus
   Exit Sub
End If
mSLocationId = RsMst_Location!LocationID
End Sub

Private Sub CmdBaseLocationID_Click()
FrmMst_Location.Show
End Sub

Private Sub CmdDesignationID_Click()
FrmMst_Designation.Show
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Txn_EmpHistory.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Txn_EmpHistory.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = I + 1 ' + 6
    For c = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(I, c), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(I, c)) Then
              oWS.Cells(mRow, c + 1) = Format(MSHFlexGrid1.TextMatrix(I, c), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(I, c)
           End If
        Else
           oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(I, c)
        End If
    Next
Next
oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub


Private Sub CmdGo_Click()
Dim tmp1 As Variant
'tmp = "Select Emp.MonthEmpID,EMP.TxnDate,EMP.EmployeeID,EMP.EmployeeNo,EMP.EmployeeName,D.Designation,EMP1.EmployeeName + '~' + EMP1.EmployeeNo as DirectReporte, " & _
'      "EMP2.EmployeeName +'~'+ EMP2.EmployeeNo as ConnectedDRC, " & _
'      "EMP3.EmployeeName +'~'+ EMP3.EmployeeNo as ConnectedRAG,EMP.EmployeeEmailID,EMP.MobileNo,EMP.Address, ML.LocationName," & _
'      "EMP.Flag,EMP.SelfReporting,EMP.AllocateExpense,EMP.CTC,EMP.Grosspay  From Txn_MonthEmp_History EMP " & _
'      "Left Join Mst_Employee EMP1 on EMP.DirectReportingTo = EMP1.EmployeeID " & _
'      "Left Join Mst_Employee EMP2 on EMP.ReportingToDRC = EMP2.EmployeeID " & _
'      "Left Join Mst_Employee EMP3 on EMP.ConnectedToRAG = EMP3.EmployeeID " & _
'      "Inner Join Mst_Designation D on EMP.DesignationID = D.DesignationID " & _
'      "Left Join Mst_Location ML on EMP.BaseLocationID = ML.LocationID "
'

tmp = "Select Emp.MonthEmpID,EMP.TxnDate,EMP.EmployeeID,EMP.EmployeeNo,EMP.EmployeeName,D.Designation,EMP1.EmployeeName + '~' + EMP1.EmployeeNo as DirectReporte, " & _
      "EMP2.EmployeeName +'~'+ EMP2.EmployeeNo as ConnectedDRC, " & _
      "EMP3.EmployeeName +'~'+ EMP3.EmployeeNo as ConnectedRAG,EMP.EmployeeEmailID,EMP.MobileNo,EMP.Address, ML.LocationName," & _
      "EMP.Flag,EMP.SelfReporting,EMP.AllocateExpense, EMP.CMPAllocation  From Txn_MonthEmp_History EMP " & _
      "Left Join Mst_Employee EMP1 on EMP.DirectReportingTo = EMP1.EmployeeID " & _
      "Left Join Mst_Employee EMP2 on EMP.ReportingToDRC = EMP2.EmployeeID " & _
      "Left Join Mst_Employee EMP3 on EMP.ConnectedToRAG = EMP3.EmployeeID " & _
      "Inner Join Mst_Designation D on EMP.DesignationID = D.DesignationID " & _
      "Left Join Mst_Location ML on EMP.BaseLocationID = ML.LocationID "
      
      
'      "--Where EMP.DesignationID = 6 or EMP.DesignationID = 7 " & _
'      "Group By EMP.EmployeeID,EMP.EmployeeNo,EMP.EmployeeName,D.Designation,EMP.Flag,EMP1.EmployeeNo,EMP1.EmployeeName, " & _
'      "EMP2.EmployeeNo , EMP2.EmployeeName, EMP3.EmployeeNo, EMP3.EmployeeName, EMP.EmployeeEmailID, EMP.MobileNo, EMP.Address "

Dim Wc As Variant
Wc = ""
If Gen_DBType = "MDB" Then
   Wc = " Where Emp.TxnDate Between #" & Format(TxtSFrom.Value, Gen_DatFormat) & "# And #" & Format(TxtSTo.Value, Gen_DatFormat) & "# "
Else
   Wc = " Where Emp.TxnDate Between '" & Format(TxtSFrom.Value, Gen_DatFormat) & "' And '" & Format(TxtSTo.Value, Gen_DatFormat) & "' "
End If

If TxtEmployeeName.Text <> "" Then
   Wc = Wc & " And EMP.EmployeeName Like '%" & TxtEmployeeName.Text & "%' "
End If
If TxtSEmpNo.Text <> "" Then
   Wc = Wc & " And EMP.EmployeeNo  = '" & TxtSEmpNo.Text & "'"
End If
If CmbSDesignation.Text <> "" Then
   Wc = Wc & " AND D.Designation  ='" & CmbSDesignation.Text & "'"
End If

If CmbSearchStatus.Text <> "" Then
   If CmbSearchStatus.Text = "Active" Then
      Wc = Wc & " And EMP.Flag = 'Y'"
   ElseIf CmbSearchStatus.Text = "De-Active" Then
      Wc = Wc & " And EMP.Flag = 'N'"
   End If
End If

If CmbSLocation.Text <> "" Then
   Wc = Wc & " And EMP.BaseLocationID = '" & mSLocationId & "'"
End If
tmp = tmp & " " & Wc

'tmp1 = "Group By EMP.MonthEmpID,EMP.TxnDate,EMP.EmployeeID,EMP.EmployeeNo,EMP.EmployeeName,D.Designation,EMP.Flag,EMP1.EmployeeNo,EMP1.EmployeeName, " & _
'"EMP2.EmployeeNo , EMP2.EmployeeName, EMP3.EmployeeNo, EMP3.EmployeeName, EMP.EmployeeEmailID, EMP.MobileNo, EMP.Address,ML.LocationName,EMP.SelfReporting,EMP.AllocateExpense "
'

'tmp = tmp & "" & tmp1

Call TmpTable

Call Grid_Head1
For iu = 1 To TmpRs.RecordCount
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
    For c = 0 To MSHFlexGrid1.Cols - 1
       MSHFlexGrid1.TextMatrix(iu, c) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
    Next
    If MSHFlexGrid1.TextMatrix(iu, 13) = "Y" Then
       MSHFlexGrid1.TextMatrix(iu, 13) = "Active"
    Else
       MSHFlexGrid1.TextMatrix(iu, 13) = "De-Active"
    End If
    If MSHFlexGrid1.TextMatrix(iu, 14) = "1" Then
       MSHFlexGrid1.TextMatrix(iu, 14) = "Yes"
    Else
       MSHFlexGrid1.TextMatrix(iu, 14) = "No"
    End If
    TmpRs.MoveNext
Next
Label22.Refresh
Label22.Caption = TmpRs.RecordCount
End Sub

Private Sub CmdCMPID_Click()
FrmMst_CMP.Show
End Sub

Private Sub CmdLocationID_Click()
FrmMst_Location.Show
End Sub
Private Sub CmdRemoveFromGrid_Click()
If MSHFlexGrid2.Rows < 3 Then Exit Sub
Dim ans As Variant
If CmbCMPID.Text = "" Then
   MsgBox "Please Select atleast one row from grid!!!"
   Exit Sub
End If
ans = MsgBox("Are you Sure you want to Delete this Record ", vbYesNo)
If ans = vbYes Then
   I = MSHFlexGrid2.RowSel
   If I > 0 Then
      If MSHFlexGrid2.TextMatrix(I, 0) <> "" Then
         For RO = I To MSHFlexGrid2.Rows - 2
             For c = 0 To MSHFlexGrid2.Cols - 1
                 MSHFlexGrid2.TextMatrix(RO, c) = MSHFlexGrid2.TextMatrix(RO + 1, c)
             Next
         Next
         MSHFlexGrid2.Rows = MSHFlexGrid2.Rows - 1
      End If
   End If
'CmbLocationID.Text = ""
   CmbCMPID.Text = ""
   CmbLocationID.SetFocus
   CmdAddToGrid.Caption = "Add To Grid"
Else
   CmbCMPID.SetFocus
End If

End Sub
Private Sub CmdDeleteFromGridAsset_Click()
If MSHFlexGrid3.Rows < 3 Then Exit Sub
Dim ans As Variant
If CmbAssetType.Text = "" Then
   MsgBox "Please Select atleast one row from grid!!!"
   Exit Sub
End If

ans = MsgBox("Are you Sure you want to Delete this Record ", vbYesNo)
If ans = vbYes Then
   I = MSHFlexGrid3.RowSel
   If I > 0 Then
      If MSHFlexGrid3.TextMatrix(I, 0) <> "" Then
         For RO = I To MSHFlexGrid3.Rows - 2
             For c = 0 To MSHFlexGrid3.Cols - 1
                 MSHFlexGrid3.TextMatrix(RO, c) = MSHFlexGrid3.TextMatrix(RO + 1, c)
             Next
         Next
         MSHFlexGrid3.Rows = MSHFlexGrid3.Rows - 1
      End If
   End If
   CmbAssetType.Text = ""
   CmbAssetStatus.Text = ""
   CmdAddToGrid.Caption = "Add To Grid"
Else
   CmbAssetStatus.SetFocus
End If
End Sub

Private Sub DeleteCmd_Click()
Dim ans As Variant
On Error GoTo msgError
ans = MsgBox("Are you Sure you want to Delete this Record ", vbYesNo)
If ans = vbYes Then
   tmp = "Select * from Txn_MonthEmp_History where (DirectReportingTo = " & TxtEmpID & " Or ReportingToDRC = " & TxtEmpID & " Or ConnectedToRAG = " & TxtEmpID & ") And TxnDate='" & Format(TxtTxnDate.Value, Gen_DatFormat) & "' "
   Call TmpTable
   If TmpRs.RecordCount > 0 Then
      MsgBox "Child Record Present!!!"
      Exit Sub
   End If
   TmpRs.Close
   
   tmp = "Select * from Txn_MonthEmp_HistoryDetail where MonthEmpID=" & Val(TxtMonthEmpID.Text)
   Call TmpTable
   If TmpRs.RecordCount > 0 Then
      MsgBox "Child Record Present!!!"
      Exit Sub
   End If
   TmpRs.Close
   
   RsTxn_MonthEmp_History.Requery
   RsTxn_MonthEmp_History.Find "MonthEmpID = " & Val(TxtMonthEmpID)
   If RsTxn_MonthEmp_History.EOF Then Exit Sub
   RsTxn_MonthEmp_History.Delete
   RsTxn_MonthEmp_History.Update
   If RsTxn_MonthEmp_History.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsTxn_MonthEmp_History.MoveNext
   If RsTxn_MonthEmp_History.EOF() Then
      RsTxn_MonthEmp_History.MoveLast
   End If
   Call Indata
End If

Exit Sub
msgError:
MsgBox "Child Record Present "
RsTxn_MonthEmp_History.CancelUpdate
End Sub

Private Sub EditCmd_Click()
Edit_Flg = "E"
Frame1.Enabled = True
'Frame2.Enabled = True
'CmbEmployeeID.Visible = True
CmbDesignationID.Visible = True
CmbDirectReporting.Visible = True
'CmbAdditionalReporting.Visible = True
CmbConnectedRAG.Visible = True
CmbFlag.Visible = True
''CmbCMPID.Enabled = True
CmbBaseLocationID.Visible = True
CmbDesignationID.Text = TxtDesignationID


CmbDirectReporting.Text = TxtDirectReporting
'CmbAdditionalReporting.Text = TxtAdditionalReporting
CmbConnectedRAG.Text = TxtConnectedRAG
CmbFlag.Text = TxtFlag
CmbBaseLocationID.Text = TxtBaseLocationID

If CmbFlag.Text = "De-Active" Then
   If TxtDOL.Visible = False Then
      LblDOL.Visible = True
      TxtDOL.Visible = True
      TxtDOL.Value = Date
    End If
End If

If TxtDOJ.Visible = False Then
   lblDOJ.Visible = True
   TxtDOJ.Visible = True
   TxtDOJ.Value = Date
End If

TxtDesignationID.Visible = False
TxtDirectReporting.Visible = False
TxtAdditionalReporting.Visible = False
TxtConnectedRAG.Visible = False
TxtFlag.Visible = False
TxtBaseLocationID.Visible = False

CmdDesignationID.Visible = True
''CmdLocationID.Visible = True
''CmdCMPID.Visible = True
''CmdAssetType.Visible = True
''cmdAssetStatus.Visible = True

'If LCase(CmbDesignationID.Text) = "managment" Then
'   ChkAllocateExpense.Enabled = True
'Else
'   ChkAllocateExpense.Enabled = False
'End If

If LCase(CmbDesignationID.Text) = "managment" Then
   ChkCMPAllocation.Enabled = True
Else
   ChkCMPAllocation.Enabled = False
End If

'TxtEmpNo.Locked = True
'TxtEmpName.Locked = True
'TxtEmpEmailID.Locked = False
'TxtMobileNo.Locked = False
'TxtAddress.Locked = False

If CmbDesignationID.Text = "CMPM" Then
   ''Frame2.Enabled = True
   CmbConnectedRAG.Enabled = True
Else
''   Frame2.Enabled = False
   CmbConnectedRAG.Enabled = False
End If
ChkSelfReporting.Enabled = True
''Frame3.Enabled = True

If LCase(Gen_UserRole) <> "power" Then
   CmbDesignationID.Enabled = False
End If

Call GButtonLock(Me, False)
End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub
Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Call GButtonLockAD(Me)
   MsgBox "Access Denied !!!!!!!!!"
   Exit Sub
End If

Call TableTxn_MonthEmp_History("Order by MonthEmpID ")
Call TableMst_Employee("Order by EmployeeID ")

If RsTxn_MonthEmp_History.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
Call Indata
Call GButtonLock(Me, True)


End Sub
Private Sub Grid_Head1()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 17

MSHFlexGrid1.TextMatrix(0, 0) = "Txn ID"
MSHFlexGrid1.TextMatrix(0, 1) = "Txn Date"
MSHFlexGrid1.TextMatrix(0, 2) = "SM's Employee Id"
MSHFlexGrid1.TextMatrix(0, 3) = "HR Employee Id"
MSHFlexGrid1.TextMatrix(0, 4) = "Employee Name"
MSHFlexGrid1.TextMatrix(0, 5) = "Designation"
MSHFlexGrid1.TextMatrix(0, 6) = "Direct Repotee"
MSHFlexGrid1.TextMatrix(0, 7) = "Connected DRC"
MSHFlexGrid1.TextMatrix(0, 8) = "Connected RAG"
MSHFlexGrid1.TextMatrix(0, 9) = "Email ID"
MSHFlexGrid1.TextMatrix(0, 10) = "Mobile No"
MSHFlexGrid1.TextMatrix(0, 11) = "Address"
MSHFlexGrid1.TextMatrix(0, 12) = "Base Location"
MSHFlexGrid1.TextMatrix(0, 13) = "Status"
MSHFlexGrid1.TextMatrix(0, 14) = "Self Reporting"
MSHFlexGrid1.TextMatrix(0, 15) = "Allocation Expense"
MSHFlexGrid1.TextMatrix(0, 16) = "CMP Allocation"
'MSHFlexGrid1.TextMatrix(0, 16) = "CTC"
'MSHFlexGrid1.TextMatrix(0, 17) = "Gross Pay"

MSHFlexGrid1.ColWidth(0) = 1000
MSHFlexGrid1.ColWidth(1) = 1600
MSHFlexGrid1.ColWidth(2) = 1200
MSHFlexGrid1.ColWidth(3) = 1400
MSHFlexGrid1.ColWidth(4) = 3000
MSHFlexGrid1.ColWidth(5) = 1200
MSHFlexGrid1.ColWidth(6) = 1800
MSHFlexGrid1.ColWidth(7) = 1800
MSHFlexGrid1.ColWidth(8) = 1800
MSHFlexGrid1.ColWidth(9) = 1400
MSHFlexGrid1.ColWidth(0) = 1000
MSHFlexGrid1.ColWidth(11) = 2000
MSHFlexGrid1.ColWidth(12) = 1000
MSHFlexGrid1.ColWidth(13) = 1000
MSHFlexGrid1.ColWidth(14) = 1000
MSHFlexGrid1.ColWidth(15) = 1000
MSHFlexGrid1.ColWidth(16) = 1000

End Sub
'Private Sub Grid_Head()
'
'MSHFlexGrid2.Clear
'MSHFlexGrid2.Rows = 2
'MSHFlexGrid2.Cols = 7
'
'MSHFlexGrid2.TextMatrix(0, 0) = "Location"
'MSHFlexGrid2.TextMatrix(0, 1) = "CMP Id"
'MSHFlexGrid2.TextMatrix(0, 2) = "CMP Name "
'MSHFlexGrid2.TextMatrix(0, 3) = "CMP Address"
'MSHFlexGrid2.TextMatrix(0, 4) = "Location ID"
'MSHFlexGrid2.TextMatrix(0, 5) = "EmpCMPMap ID"
'MSHFlexGrid2.TextMatrix(0, 6) = "SM_Prefix"
'
'MSHFlexGrid2.ColWidth(0) = 1400
'MSHFlexGrid2.ColWidth(1) = 1200
'MSHFlexGrid2.ColWidth(2) = 2600
'MSHFlexGrid2.ColWidth(3) = 6600
'MSHFlexGrid2.ColWidth(4) = 0
'MSHFlexGrid2.ColWidth(5) = 0
'MSHFlexGrid2.ColWidth(6) = 0
'
'End Sub
'Private Sub Grid_Head2()
'
'''MSHFlexGrid3.Clear
'''MSHFlexGrid3.Rows = 2
'''MSHFlexGrid3.Cols = 9
'''MSHFlexGrid3.TextMatrix(0, 0) = "Asset Type"
'''MSHFlexGrid3.TextMatrix(0, 1) = "Asset Status"
'''MSHFlexGrid3.TextMatrix(0, 2) = "Asset Qty"
'''MSHFlexGrid3.TextMatrix(0, 3) = "Description"
'''MSHFlexGrid3.TextMatrix(0, 4) = "Verified On"
'''MSHFlexGrid3.TextMatrix(0, 5) = "EmpAssetMapID"
'''MSHFlexGrid3.TextMatrix(0, 6) = "Remarks"
'''MSHFlexGrid3.TextMatrix(0, 7) = "AMC Exp Date"
'''MSHFlexGrid3.TextMatrix(0, 8) = "AMC Amount"
'''
'''MSHFlexGrid3.ColWidth(0) = 1800
'''MSHFlexGrid3.ColWidth(1) = 1600
'''MSHFlexGrid3.ColWidth(2) = 1000
'''MSHFlexGrid3.ColWidth(3) = 5400
'''MSHFlexGrid3.ColWidth(4) = 1000
'''MSHFlexGrid3.ColWidth(5) = 0
'''MSHFlexGrid3.ColWidth(6) = 2000
'''MSHFlexGrid3.ColWidth(7) = 1000
'''MSHFlexGrid3.ColWidth(8) = 1000
'
'End Sub
Private Sub CmbAssetType_GotFocus()

tmp = CmbAssetType.Text
Call TableMst_AssetType("")
CmbAssetType.Clear
If RsMst_AssetType.RecordCount = 0 Then
   MsgBox "No Asset found"
   CmbAssetType.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_AssetType.RecordCount
    CmbAssetType.AddItem RsMst_AssetType!AssetType
    RsMst_AssetType.MoveNext
Next
CmbAssetType.Text = tmp

End Sub
Private Sub CmbLocationID_GotFocus()

tmp = CmbLocationID.Text
Call TableMst_Location("")
CmbLocationID.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Location found"
   CmbLocationID.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbLocationID.AddItem RsMst_Location!LocationName
    RsMst_Location.MoveNext
Next
CmbLocationID.Text = tmp

End Sub
Private Sub CmbAssetStatus_GotFocus()

tmp = CmbAssetStatus.Text
Call TableMst_AssetStatus("")
CmbAssetStatus.Clear
If RsMst_AssetStatus.RecordCount = 0 Then
   MsgBox "No Asset Status found"
   CmbAssetStatus.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_AssetStatus.RecordCount
    CmbAssetStatus.AddItem RsMst_AssetStatus!AssetStatus
    RsMst_AssetStatus.MoveNext
Next
CmbAssetStatus.Text = tmp

End Sub
Private Sub CmbAssetType_LostFocus()

If CmbAssetType.Text = "" Then
   mAssetTypeId = Null
   Exit Sub
End If

RsMst_AssetType.MoveFirst
RsMst_AssetType.Find "AssetType = '" & CmbAssetType.Text & "'"

If RsMst_AssetType.EOF Then
   MsgBox "Invalid Asset  selection "
   CmbAssetType.SetFocus
   Exit Sub
End If
mAssetTypeId = RsMst_AssetType.Fields(0)

End Sub
Private Sub CmbAssetStatus_LostFocus()
If CmbAssetStatus.Text = "" Then
   mAssetStatusID = Null
   Exit Sub
End If

RsMst_AssetStatus.MoveFirst
RsMst_AssetStatus.Find "AssetStatus = '" & CmbAssetStatus.Text & "'"

If RsMst_AssetStatus.EOF Then
   MsgBox "Invalid Status selection "
   CmbAssetStatus.SetFocus
   Exit Sub
End If
mAssetStatusID = RsMst_AssetStatus.Fields(0)

End Sub
Private Sub CmbLocationID_LostFocus()

If CmbLocationID.Text = "" Then
   mLocationID = 0
   Exit Sub
End If
' Code for clearing Godown combox in case of change
If CmbLocationID.Text <> tmp Then
   CmbCMPID.Clear
   CmbCMPID.Text = ""
   mCMPID = 0
   
End If

RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocationID.Text & "'"

If RsMst_Location.EOF Then
   MsgBox "Invalid Location  selection "
   CmbLocationID.SetFocus
   Exit Sub
End If
mLocationID = RsMst_Location!LocationID
mSM_Prefix = RsMst_Location!SM_Prefix

End Sub

Private Sub CmbDesignationID_GotFocus()
tmp = CmbDesignationID.Text
Call TableMst_Designation("")
CmbDesignationID.Clear
If RsMst_Designation.RecordCount = 0 Then
   MsgBox "No Designation found"
   CmbDesignationID.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Designation.RecordCount
    CmbDesignationID.AddItem RsMst_Designation!Designation
    RsMst_Designation.MoveNext
Next
CmbDesignationID.Text = tmp
End Sub
Private Sub CmbDesignationID_LostFocus()

If CmbDesignationID.Text = "" Then
   mDesignationID = Null
   Exit Sub
End If

If CmbDesignationID.Text <> tmp Then
   Call ClearTabs
   CmbAdditionalReporting.Text = ""
   CmbConnectedRAG.Text = ""
   mConnectedRAGID = Null
   mAdditionalReportingID = Null
End If


RsMst_Designation.MoveFirst
RsMst_Designation.Find "Designation = '" & CmbDesignationID.Text & "'"

If RsMst_Designation.EOF Then
   MsgBox "Invalid Designation Selection "
   CmbDesignationID.SetFocus
   Exit Sub
End If
mDesignationID = RsMst_Designation!DesignationID
mReportingDesignationID = RsMst_Designation!ReportingDesignationID

If tmp <> CmbDesignationID.Text Then
   CmbDirectReporting.Text = ""
End If

If mDesignationID = 6 Then
   ''Frame2.Enabled = True
   'Frame3.Enabled = True
''   AssetVerifiedon.Value = Date
   Label25.Visible = True
   ChkSelfReporting.Visible = True
   ChkSelfReporting.Value = 0
   CmbAdditionalReporting.Enabled = True
   CmbConnectedRAG.Enabled = True
Else
   Label25.Visible = False
   ChkSelfReporting.Visible = False
''   Frame2.Enabled = False
   'Frame3.Enabled = False
   ChkSelfReporting.Value = 0
''   AssetVerifiedon.Value = Date
   CmbAdditionalReporting.Enabled = False
   CmbConnectedRAG.Enabled = False
End If


If mDesignationID = 1 Then
   ChkCMPAllocation.Enabled = True
Else
   ChkCMPAllocation.Enabled = False
   ChkCMPAllocation.Value = 1
End If


End Sub
Private Sub ClearTabs()
''CmbLocationID.Text = ""
''CmbCMPID.Text = ""
''Call Grid_Head
''CmbAssetType.Text = ""
''CmbAssetStatus.Text = ""
''TxtAssetDescription = ""
''Call Grid_Head2
End Sub
Private Function Validate_DuplicateGrid() As Boolean
Dim Retval As Boolean
Retval = True
If MSHFlexGrid2.Rows = 2 Then
   Retval = True
Else
    For I = 1 To MSHFlexGrid2.Rows - 1
        If MSHFlexGrid2.TextMatrix(I, 1) <> "" Then
           If CmdAddToGrid.Caption = "Update Grid" Then
              If I <> MSHFlexGrid2.RowSel Then
                 If Val(MSHFlexGrid2.TextMatrix(I, 1)) = mCMPID Then
                    Retval = False
                 End If
              End If
           Else
              If Val(MSHFlexGrid2.TextMatrix(I, 1)) = mCMPID Then
                 Retval = False
              End If
           End If
        End If
    Next
End If
Validate_DuplicateGrid = Retval
End Function
Private Function Validate_DuplicateGridForAll() As Boolean
Dim Retval As Boolean
Retval = True
If MSHFlexGrid2.Rows = 2 Then
   Retval = True
Else
    For I = 1 To MSHFlexGrid2.Rows - 1
        If MSHFlexGrid2.TextMatrix(I, 1) <> "" Then
           If Val(MSHFlexGrid2.TextMatrix(I, 1)) = mCMPID Then
              Retval = False
           End If
        End If
    Next
End If
Validate_DuplicateGridForAll = Retval
End Function

Private Function CheckCMPMapping() As Boolean
Dim Retval As Boolean
Retval = True
TmpRs3.Find "CMPID = " & mCMPID & ""
If Not TmpRs3.EOF = True Then
   If TmpRs3!DesignationID = 6 Then
      Retval = False
   End If
   If TmpRs3!DesignationID = 7 Then
      If CmdAddToGrid.Caption = "Add To Grid" Then
         Retval = False
      End If
   End If
End If
CheckCMPMapping = Retval
End Function



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

RsTxn_MonthEmp_History.MoveFirst

If RsTxn_MonthEmp_History.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
RsTxn_MonthEmp_History.MoveFirst
RsTxn_MonthEmp_History.Find "MonthEmpID = " & Val(mCaId) & ""
If RsTxn_MonthEmp_History.EOF Then
   RsTxn_MonthEmp_History.MoveLast
End If
Frame0.Visible = False
Call Indata
End Sub

Private Sub MSHFlexGrid2_Click()
Dim I As Variant
I = MSHFlexGrid2.RowSel
If MSHFlexGrid2.TextMatrix(I, 0) = "" Then Exit Sub
CmbLocationID.Text = MSHFlexGrid2.TextMatrix(I, 0)
mCMPID = GetCMPID(MSHFlexGrid2.TextMatrix(I, 2))
CmbCMPID.Text = MSHFlexGrid2.TextMatrix(I, 2)
CmbLocationID.SetFocus
CmdAddToGrid.Caption = "Update Grid"
CmbCMPID.SetFocus
End Sub
Private Sub MSHFlexGrid3_Click()
Dim I As Variant
I = MSHFlexGrid3.RowSel
If MSHFlexGrid3.TextMatrix(I, 0) = "" Then Exit Sub
CmbAssetType.Text = MSHFlexGrid3.TextMatrix(I, 0)
CmbAssetStatus.Text = MSHFlexGrid3.TextMatrix(I, 1)
TxtAssetQty = Val(MSHFlexGrid3.TextMatrix(I, 2))
TxtAssetDescription = MSHFlexGrid3.TextMatrix(I, 3)
AssetVerifiedon.Value = MSHFlexGrid3.TextMatrix(I, 4)
TxtRemarks = MSHFlexGrid3.TextMatrix(I, 6)
TxtAMCExpDate.Value = IIf(MSHFlexGrid3.TextMatrix(I, 7) = "", Date, MSHFlexGrid3.TextMatrix(I, 7))
TxtAMCAmount = MSHFlexGrid3.TextMatrix(I, 8)
CmbAssetType.SetFocus
CmdAdd2GridAsset.Caption = "Update Grid"
End Sub

Private Sub Option2_Click()
'Call Option1_Click
End Sub

Private Sub Option3_Click()
'Call Option1_Click
End Sub

Private Sub Option5_Click()
'Call Option1_Click
End Sub

Private Sub SaveCmd_Click()
If SaveCmd.Enabled = False Then Exit Sub
Dim strSplitString() As String
Dim EmpNo_, EmpName_ As String

If TxtEmpNo = "" Then
   MsgBox "Blank Employee No Not Allowed!!!"
   TxtEmpNo.SetFocus
   Exit Sub
End If

If TxtEmpName = "" Then
   MsgBox "Blank Employee Name Not Allowed!!!"
   TxtEmpName.SetFocus
   Exit Sub
End If

'If CmbDirectReporting.Text <> "" Or CmbAdditionalReporting.Text <> "" Then
'   If CmbDirectReporting.Text = CmbAdditionalReporting.Text Then
'      MsgBox "Direct Reportee And Connected To DRC Should Be Different!!!"
'      Exit Sub
'   End If
'End If

If CmbDirectReporting.Text <> "" Or CmbConnectedRAG.Text <> "" Then
   If CmbDirectReporting.Text = CmbConnectedRAG.Text Then
      MsgBox "Direct Reportee  And Connected To RAG Should Be Different!!!"
      Exit Sub
   End If
End If

If CmbDirectReporting.Text <> "" Then
   strSplitString = Split(CmbDirectReporting.Text, "~")
   EmpName_ = Trim(strSplitString(0))
   EmpNo_ = Trim(strSplitString(1))
   If TxtEmpNo = EmpNo_ Then
      MsgBox "Self Reporting Not Allowed!!!"
      CmbDirectReporting.SetFocus
   End If
Else
   MsgBox "Blank Direct Reportee not allowed!!!"
   CmbDirectReporting.SetFocus
   Exit Sub
End If

If CmbDesignationID.Text = "" Then
   MsgBox "Blank Designation Not Allowed!!!"
   CmbDesignationID.SetFocus
   Exit Sub
End If

'If LCase(Gen_UserRole) <> "power" Then
'   If Edit_Flg = "E" Then
'      If mDesignationID <> mOldDesignationID Then
'         MsgBox "Change in Designation not allowed!!!"
'         CmbDesignationID.SetFocus
'         Exit Sub
'      End If
'   End If
'End If

If CmbFlag.Text = "" Then
   MsgBox "Blank Flag Not Allowed!!!"
   CmbFlag.SetFocus
   Exit Sub
End If

If CmbBaseLocationID.Text = "" Then
   MsgBox "Blank Base Location Not Allowed!!!"
   CmbBaseLocationID.SetFocus
   Exit Sub
End If


If Edit_Flg = "A" Then
   If RsTxn_MonthEmp_History.RecordCount = 0 Then
      tmp = 1
   Else
        RsTxn_MonthEmp_History.MoveFirst
        
        RsTxn_MonthEmp_History.Filter = "EmployeeID = " & mEmployeeID & " And TxnDate='" & TxtTxnDate.Value & "'"
        If Not RsTxn_MonthEmp_History.EOF Then
           MsgBox "Duplicate TxnDate,Employee not Allowed"
           TxtEmpNo.SetFocus
           Exit Sub
        End If
        RsTxn_MonthEmp_History.Filter = ""
        tmp = "select max(MonthEmpID) from Txn_MonthEmp_History "
        Call TmpTable
        tmp = IIf(IsNull(TmpRs.Fields(0)), 1, TmpRs.Fields(0) + 1)
    End If
    RsTxn_MonthEmp_History.AddNew
    RsTxn_MonthEmp_History!MonthEmpID = tmp
    TxtMonthEmpID = tmp
Else
'    RsTxn_MonthEmp_History.MoveFirst
'    RsTxn_MonthEmp_History.Find "MonthEmpID = '" & TxtMonthEmpID & "'"
'    If Not RsTxn_MonthEmp_History.EOF Then
'       If RsTxn_MonthEmp_History!MonthEmpID <> TxtMonthEmpID Then
'          MsgBox "Duplicate Txn ID not Allowed"
'          TxtMonthEmpID.SetFocus
'          Exit Sub
'       End If
'    End If
   RsTxn_MonthEmp_History.MoveFirst
   RsTxn_MonthEmp_History.Find "MonthEmpID = " & TxtMonthEmpID & ""
End If
RsTxn_MonthEmp_History!MonthEmpID = TxtMonthEmpID
RsTxn_MonthEmp_History!TxnDate = TxtTxnDate.Value
RsTxn_MonthEmp_History!DOJ = TxtDOJ.Value
RsTxn_MonthEmp_History!EmployeeID = TxtEmpID
RsTxn_MonthEmp_History!EmployeeNo = TxtEmpNo
RsTxn_MonthEmp_History!EmployeeName = TxtEmpName
RsTxn_MonthEmp_History!DesignationID = mDesignationID
RsTxn_MonthEmp_History!SelfReporting = ChkSelfReporting.Value
RsTxn_MonthEmp_History!DirectReportingTo = mDirectReportingID
RsTxn_MonthEmp_History!ReportingToDRC = mAdditionalReportingID
RsTxn_MonthEmp_History!ConnectedToRAG = mConnectedRAGID
RsTxn_MonthEmp_History!EmployeeEmailID = TxtEmpEmailID
RsTxn_MonthEmp_History!MobileNo = TxtMobileNo
RsTxn_MonthEmp_History!Address = TxtAddress
RsTxn_MonthEmp_History!BaseLocationID = mBaseLocationID
RsTxn_MonthEmp_History!AllocateExpense = ChkAllocateExpense.Value
RsTxn_MonthEmp_History!CMPAllocation = ChkCMPAllocation.Value
'RsTxn_MonthEmp_History!LocationID = mLocationID
RsTxn_MonthEmp_History!CTC = Val(TxtCTC.Text)
RsTxn_MonthEmp_History!GrossPay = Val(TxtGrossPay.Text)
If TxtDOL.Visible = True Then
   RsTxn_MonthEmp_History!DOL = TxtDOL.Value
End If
RsTxn_MonthEmp_History!Flag = mFlag

If IsNull(RsTxn_MonthEmp_History!CreatedBy) = True Or RsTxn_MonthEmp_History!CreatedBy = "" Then
   RsTxn_MonthEmp_History!CreatedBy = Gen_UserLoginID
   RsTxn_MonthEmp_History!CreatedDate = Now
Else
   RsTxn_MonthEmp_History!UpdatedBy = Gen_UserLoginID
   RsTxn_MonthEmp_History!UpdatedDate = Now
End If
RsTxn_MonthEmp_History.Update

'If MSHFlexGrid2.Rows > 2 Then
'   tmp = "Delete from Mst_EmpCMPMapping where EmployeeID = " & RsTxn_MonthEmp_History!EmployeeID & ""
'   Call TmpTable
'
'   Call TableMst_EmpCMPMap(" Where EmployeeID = " & RsTxn_MonthEmp_History!EmployeeID & "")
'
'   For I = 1 To MSHFlexGrid2.Rows - 2
'       If MSHFlexGrid2.TextMatrix(I, 5) <> "" Then
'          RsMst_EmpCMPMap.Find "EmpCMPMapID = " & MSHFlexGrid2.TextMatrix(I, 5)
'          If RsMst_EmpCMPMap.EOF Then
'             RsMst_EmpCMPMap.AddNew
'             RsMst_EmpCMPMap!EmpCMPMapID = GetEmpCMPMapID()
'          End If
'       Else
'          RsMst_EmpCMPMap.AddNew
'          RsMst_EmpCMPMap!EmpCMPMapID = GetEmpCMPMapID()
'       End If
'       RsMst_EmpCMPMap!LocationID = MSHFlexGrid2.TextMatrix(I, 4)
'       RsMst_EmpCMPMap!CMPID = MSHFlexGrid2.TextMatrix(I, 1)
'       RsMst_EmpCMPMap!EmployeeID = TxtEmpID
'       RsMst_EmpCMPMap!DesignationID = mDesignationID
'       RsMst_EmpCMPMap!SM_Prefix = MSHFlexGrid2.TextMatrix(I, 6)
'       If IsNull(RsMst_EmpCMPMap!CreatedBy) = True Or RsMst_EmpCMPMap!CreatedBy = "" Then
'          RsMst_EmpCMPMap!CreatedBy = Gen_UserLoginID
'          RsMst_EmpCMPMap!CreatedDate = Now
'       Else
'          RsMst_EmpCMPMap!UpdatedBy = Gen_UserLoginID
'          RsMst_EmpCMPMap!UpdatedDate = Now
'       End If
'       RsMst_EmpCMPMap.Update
'   Next
'Else
'   tmp = "Delete from Mst_EmpCMPMapping where EmployeeID = " & RsTxn_MonthEmp_History!EmployeeID & ""
'   Call TmpTable
'End If
'
'If MSHFlexGrid3.Rows > 2 Then
'   tmp = "Delete from Mst_EmpAssetMap where EmployeeID = " & RsTxn_MonthEmp_History!EmployeeID & ""
'   Call TmpTable
'
'   Call TableMst_EmpAssetMap(" Where EmployeeID = " & RsTxn_MonthEmp_History!EmployeeID & "")
'
'   For I = 1 To MSHFlexGrid3.Rows - 2
'       If MSHFlexGrid3.TextMatrix(I, 5) <> "" Then
'
'          RsMst_EmpAssetMap.Find "EmpAssetMapID = " & MSHFlexGrid3.TextMatrix(I, 5)
'          If RsMst_EmpAssetMap.EOF Then
'             RsMst_EmpAssetMap.AddNew
'             RsMst_EmpAssetMap!EmpAssetMapID = GetEmpAssetMapID()
'          End If
'       Else
'          RsMst_EmpAssetMap.AddNew
'          RsMst_EmpAssetMap!EmpAssetMapID = GetEmpAssetMapID()
'       End If
'       RsMst_EmpAssetMap!EmployeeID = TxtEmpID
'       RsMst_EmpAssetMap!AssetID = GetAssetTypeID(MSHFlexGrid3.TextMatrix(I, 0))
'       RsMst_EmpAssetMap!AssetStatusID = GetAssetStatusID(MSHFlexGrid3.TextMatrix(I, 1))
'       RsMst_EmpAssetMap!AssetQty = Val(MSHFlexGrid3.TextMatrix(I, 2))
'       RsMst_EmpAssetMap!Description = MSHFlexGrid3.TextMatrix(I, 3)
'       RsMst_EmpAssetMap!VerifiedDate = MSHFlexGrid3.TextMatrix(I, 4)
'       RsMst_EmpAssetMap!Remarks = MSHFlexGrid3.TextMatrix(I, 6)
'       If MSHFlexGrid3.TextMatrix(I, 7) <> "" Then
'          RsMst_EmpAssetMap!AMCExpDate = MSHFlexGrid3.TextMatrix(I, 7)
'       End If
'       RsMst_EmpAssetMap!AMCAmount = Val(MSHFlexGrid3.TextMatrix(I, 8))
'       If IsNull(RsMst_EmpAssetMap!CreatedBy) = True Or RsMst_EmpAssetMap!CreatedBy = "" Then
'          RsMst_EmpAssetMap!CreatedBy = Gen_UserLoginID
'          RsMst_EmpAssetMap!CreatedDate = Now
'       Else
'          RsMst_EmpAssetMap!UpdatedBy = Gen_UserLoginID
'          RsMst_EmpAssetMap!UpdatedDate = Now
'       End If
'       RsMst_EmpAssetMap.Update
'   Next
'Else
'   tmp = "Delete from Mst_EmpAssetMap where EmployeeID = " & RsTxn_MonthEmp_History!EmployeeID & ""
'   Call TmpTable
'End If
Call GButtonLock(Me, True)
RsTxn_MonthEmp_History.Requery
RsTxn_MonthEmp_History.MoveFirst
RsTxn_MonthEmp_History.Find "MonthEmpID = " & TxtMonthEmpID & ""
If RsTxn_MonthEmp_History.EOF Then
   RsTxn_MonthEmp_History.MoveLast
End If
Call Indata

End Sub
Private Function GetEmpCMPMapID() As Double
Dim Retval As Double
tmp = "select max(EmpCMPMapID) from Mst_EmpCMPMapping"
Call TmpTable
tmp = IIf(IsNull(TmpRs.Fields(0)), 1, TmpRs.Fields(0) + 1)
GetEmpCMPMapID = tmp
End Function
Private Function GetEmpAssetMapID() As Double
Dim Retval As Double
tmp = "select max(EmpAssetMapID) from Mst_EmpAssetMap"
Call TmpTable
tmp = IIf(IsNull(TmpRs.Fields(0)), 1, TmpRs.Fields(0) + 1)
GetEmpAssetMapID = tmp
End Function
Public Sub Indata()
''Frame2.Enabled = False
''Frame3.Enabled = False
Frame1.Enabled = False

CmbEmployeeID.Visible = False

CmbDesignationID.Visible = False
CmbDirectReporting.Visible = False
CmbAdditionalReporting.Visible = False
CmbConnectedRAG.Visible = False
CmbFlag.Visible = False
CmbBaseLocationID.Visible = False


TxtEmpName.Visible = True
TxtDesignationID.Visible = True
TxtDirectReporting.Visible = True
'TxtAdditionalReporting.Visible = True
TxtConnectedRAG.Visible = True
TxtFlag.Visible = True
TxtBaseLocationID.Visible = True

CmdDesignationID.Visible = False
''CmdLocationID.Visible = False
''CmdCMPID.Visible = False
''CmdAssetType.Visible = False
''cmdAssetStatus.Visible = False
CmdBaseLocationID.Visible = False
ChkSelfReporting.Visible = False

TxtEmpNo.Locked = True
TxtEmpName.Locked = True
TxtEmpEmailID.Locked = True
TxtMobileNo.Locked = True
TxtAddress.Locked = True
ChkAllocateExpense.Enabled = False
ChkCMPAllocation.Enabled = False

TxtEmpID = ""
TxtEmpNo = ""
TxtEmpName = ""
TxtEmpEmailID = ""
TxtMobileNo = ""
TxtAddress = ""
TxtDesignationID = ""
TxtDirectReporting = ""
TxtCTC.Text = ""
TxtGrossPay.Text = ""
'TxtAdditionalReporting = ""
TxtConnectedRAG = ""
TxtFlag = ""
TxtBaseLocationID = ""

TxtMonthEmpID = RsTxn_MonthEmp_History!MonthEmpID
TxtTxnDate.Value = RsTxn_MonthEmp_History!TxnDate
If IsNull(RsTxn_MonthEmp_History!DOJ) = False Then
   TxtDOJ.Visible = True
   lblDOJ.Visible = True
   TxtDOJ.Value = RsTxn_MonthEmp_History!DOJ
Else
   TxtDOJ.Visible = False
   lblDOJ.Visible = False
End If
mEmployeeID = RsTxn_MonthEmp_History!EmployeeID
TxtEmpID = RsTxn_MonthEmp_History!EmployeeID
TxtEmpNo = RsTxn_MonthEmp_History!EmployeeNo
TxtEmpName = RsTxn_MonthEmp_History!EmployeeName
TxtEmpEmailID = IIf(IsNull(RsTxn_MonthEmp_History!EmployeeEmailID), "", RsTxn_MonthEmp_History!EmployeeEmailID)
TxtMobileNo = IIf(IsNull(RsTxn_MonthEmp_History!MobileNo), "", RsTxn_MonthEmp_History!MobileNo)
TxtAddress = IIf(IsNull(RsTxn_MonthEmp_History!Address), "", RsTxn_MonthEmp_History!Address)
ChkAllocateExpense.Value = IIf(IsNull(RsTxn_MonthEmp_History!AllocateExpense), 0, RsTxn_MonthEmp_History!AllocateExpense)
ChkCMPAllocation.Value = IIf(IsNull(RsTxn_MonthEmp_History!CMPAllocation), 0, RsTxn_MonthEmp_History!CMPAllocation)
mOldDesignationID = RsTxn_MonthEmp_History!DesignationID
mDesignationID = RsTxn_MonthEmp_History!DesignationID
TxtDesignationID = GetDesignaion(mDesignationID)
mReportingDesignationID = GetReportingDesignation(mDesignationID)

mDirectReportingID = IIf(IsNull(RsTxn_MonthEmp_History!DirectReportingTo), "", RsTxn_MonthEmp_History!DirectReportingTo)
If mDirectReportingID <> "" Then
   TxtDirectReporting = GetEmployeeNonName(mDirectReportingID)
Else
   mDirectReportingID = Null
End If

'mAdditionalReportingID = IIf(IsNull(RsTxn_MonthEmp_History!ReportingToDRC), "", RsTxn_MonthEmp_History!ReportingToDRC)
'If mAdditionalReportingID <> "" Then
'   TxtAdditionalReporting = GetEmployeeNonName(mAdditionalReportingID)
'Else
'   mAdditionalReportingID = Null
'End If

mConnectedRAGID = RsTxn_MonthEmp_History!ConnectedToRAG
If mConnectedRAGID <> 0 Then
   TxtConnectedRAG = GetEmployeeNonName(mConnectedRAGID)
Else
   mConnectedRAGID = Null
End If

mFlag = RsTxn_MonthEmp_History!Flag

If mFlag = "Y" Then
   TxtFlag = "Active"
   LblDOL.Visible = False
   TxtDOL.Visible = False
Else
   TxtFlag = "De-Active"
   If IsNull(RsTxn_MonthEmp_History!DOL) = False Then
      LblDOL.Visible = True
      TxtDOL.Visible = True
      TxtDOL.Value = RsTxn_MonthEmp_History!DOL
   End If
End If

mBaseLocationID = IIf(IsNull(RsTxn_MonthEmp_History!BaseLocationID), 0, RsTxn_MonthEmp_History!BaseLocationID)

If mBaseLocationID <> 0 Then
   TxtBaseLocationID = GetLocationName(mBaseLocationID)
Else
   mBaseLocationID = 0
End If
TxtCTC.Text = RsTxn_MonthEmp_History!CTC
TxtGrossPay.Text = RsTxn_MonthEmp_History!GrossPay

If RsTxn_MonthEmp_History!SelfReporting = 1 Then
   Label25.Visible = True
   ChkSelfReporting.Visible = True
Else
   Label25.Visible = False
   ChkSelfReporting.Visible = False
End If
ChkSelfReporting.Value = IIf(IsNull(RsTxn_MonthEmp_History!SelfReporting), 0, RsTxn_MonthEmp_History!SelfReporting)
ChkSelfReporting.Enabled = False

Call ClearTabs
TxtCreated = IIf(IsNull(RsTxn_MonthEmp_History!CreatedBy), "", RsTxn_MonthEmp_History!CreatedBy) & " :- " & IIf(IsNull(RsTxn_MonthEmp_History!CreatedDate), "", RsTxn_MonthEmp_History!CreatedDate)
TxtUpdated = IIf(IsNull(RsTxn_MonthEmp_History!UpdatedBy), "", RsTxn_MonthEmp_History!UpdatedBy) & " :- " & IIf(IsNull(RsTxn_MonthEmp_History!UpdatedDate), "", RsTxn_MonthEmp_History!UpdatedDate)

''Call Grid_Head

'Call TableMst_EmpCMPMap(" Where EmployeeID = " & RsTxn_MonthEmp_History!EmployeeID & "")

'tmp = "Select  ML.LocationName,MC.CMPId,MC.CMPName,MC.Address,ML.LocationID,ECM.EmpCMPMapID,ML.SM_Prefix " & _
'      " From Mst_EmpCMPMapping ECM  " & _
'      " Inner Join Mst_Location ML On ECM.LocationID = ML.LocationID  " & _
'      " Inner Join Mst_CMP MC On ECM.CMPID = MC.CMPID " & _
'      " Where MC.CloseDate Is Null ANd ECM.EmployeeID = " & RsTxn_MonthEmp_History!EmployeeID & "  "
'
'Call Tmp1Table
'
'If TmpRs1.RecordCount > 0 Then
'   For I = 1 To TmpRs1.RecordCount
'       MSHFlexGrid2.TextMatrix(I, 0) = TmpRs1!LocationName 'GetLocationName(TmpRs1!LocationID)
'       MSHFlexGrid2.TextMatrix(I, 1) = TmpRs1!CMPID
'       MSHFlexGrid2.TextMatrix(I, 2) = TmpRs1!CMPName 'GetCMPName(TmpRs1!CMPID)
'       MSHFlexGrid2.TextMatrix(I, 3) = IIf(IsNull(TmpRs1!Address), "", TmpRs1!Address) 'GetCMPAddress(TmpRs1!CMPID, "A")
'       MSHFlexGrid2.TextMatrix(I, 4) = TmpRs1!LocationID
'       MSHFlexGrid2.TextMatrix(I, 5) = TmpRs1!EmpCMPMapID
'       MSHFlexGrid2.TextMatrix(I, 6) = TmpRs1!SM_Prefix
'       TmpRs1.MoveNext
'       MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
'   Next
'End If

'Call Grid_Head2
''Call TableMst_EmpAssetMap(" Where EmployeeID = " & RsTxn_MonthEmp_History!EmployeeID & "")
''
''If RsMst_EmpAssetMap.RecordCount > 0 Then
''   For I = 1 To RsMst_EmpAssetMap.RecordCount
''       MSHFlexGrid3.TextMatrix(I, 0) = GetAssetType(RsMst_EmpAssetMap!AssetID)
''       MSHFlexGrid3.TextMatrix(I, 1) = GetAssetStatus(RsMst_EmpAssetMap!AssetStatusID)
''       MSHFlexGrid3.TextMatrix(I, 2) = RsMst_EmpAssetMap!AssetQty
''       MSHFlexGrid3.TextMatrix(I, 3) = RsMst_EmpAssetMap!Description
''       MSHFlexGrid3.TextMatrix(I, 4) = RsMst_EmpAssetMap!VerifiedDate
''       MSHFlexGrid3.TextMatrix(I, 5) = RsMst_EmpAssetMap!EmpAssetMapID
''       MSHFlexGrid3.TextMatrix(I, 6) = IIf(IsNull(RsMst_EmpAssetMap!Remarks), "", RsMst_EmpAssetMap!Remarks)
''       MSHFlexGrid3.TextMatrix(I, 7) = IIf(IsNull(RsMst_EmpAssetMap!AMCExpDate), "", RsMst_EmpAssetMap!AMCExpDate)
''       MSHFlexGrid3.TextMatrix(I, 8) = IIf(IsNull(RsMst_EmpAssetMap!AMCAmount), 0, RsMst_EmpAssetMap!AMCAmount)
''       RsMst_EmpAssetMap.MoveNext
''       MSHFlexGrid3.Rows = MSHFlexGrid3.Rows + 1
''   Next
''End If
''SSTab1.Tab = 0

Call GButtonLock(Me, True)
End Sub
Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsTxn_MonthEmp_History.MoveNext
If RsTxn_MonthEmp_History.EOF Then
   RsTxn_MonthEmp_History.MoveLast
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
RsTxn_MonthEmp_History.MovePrevious
If RsTxn_MonthEmp_History.BOF Then
   RsTxn_MonthEmp_History.MoveFirst
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
RsTxn_MonthEmp_History.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsTxn_MonthEmp_History.MoveLast
Call Indata
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = Fasle
End Sub

Private Sub SearchCmd_Click()
Dim x As Integer
If SearchCmd.Caption = "Cancel" Then
   If RsTxn_MonthEmp_History.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   If RsTxn_MonthEmp_History.EOF Then
      RsTxn_MonthEmp_History.MoveFirst
   End If
   Call Indata
   Exit Sub
End If

Frame0.Height = Frame4.Height - 500
Frame0.Left = Frame1.Left
Frame0.Top = Frame1.Top
Frame0.Width = Frame4.Width - 400
Frame0.Visible = True
MSHFlexGrid1.Clear
MSHFlexGrid1.Width = Frame0.Width - 300
MSHFlexGrid1.Height = Frame0.Height - 1600

Call Grid_Head1
TxtEmployeeName.Text = ""
CmbSearchStatus.Text = ""
CmbSLocation.Text = ""
CmbSDesignation.Text = ""
TxtSEmpNo.Text = ""
mTotalDays = Day(DateSerial(Year(Date), Month(Date) + 1, 0))
TxtSFrom.Value = (Year(Date) & "-" & Month(Date) & "-01")
TxtSTo.Value = Year(Date) & "-" & Month(Date) & "-" & mTotalDays
Label22.Caption = ""

End Sub

Private Sub TxtAddress_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "TxtAddress")
End If
End Sub

Private Sub TxtAMCAmount_LostFocus()
If TxtAMCAmount = "" Then Exit Sub
If IsNumeric(TxtAMCAmount) = False Or Val(TxtAMCAmount) < 0 Then
   MsgBox "Invalid number format!!!"
   TxtAMCAmount.SetFocus
   Exit Sub
End If
TxtAMCAmount = Format(TxtAMCAmount, "#########0.00")

End Sub

Private Sub TxtAssetDescription_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "TxtAssetDescription")
End If
End Sub

Private Sub TxtAssetQty_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "TxtAssetQty")
End If
End Sub

Private Sub TxtAssetQty_LostFocus()
If TxtAssetQty = "" Then Exit Sub
If IsNumeric(TxtAssetQty) = False Or Val(TxtAssetQty) < 0 Then
   MsgBox "Invalid number format!!!"
   TxtAssetQty.SetFocus
   Exit Sub
End If
End Sub
Private Sub Option1_Click()
'Label19.Visible = True
'TxtSearchEmployee.Visible = True
'TxtSearchEmployee = ""
''CmbSearchStatus.Visible = False
'CmbSearchStatus.Text = ""
'CmbSLocation.Text = ""
'CmdGo.Visible = True
'If Option1.Value = True Then
'   Label19.Caption = "Search By EHOP Employee ID"
'ElseIf Option2.Value = True Then
'   Label19.Caption = "Search By Designation"
'ElseIf Option5.Value = True Then
'   Label19.Caption = "Search By Employee Name"
'End If
End Sub

Private Sub TxtEmpEmailID_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "TxtEmpEmailID")
End If
End Sub

Private Sub TxtEmpId_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "TxtEmpID")
End If
End Sub

Private Sub TxtEmpName_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "TxtEmpName")
End If
End Sub

Private Sub TxtEmpNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "TxtEmpNo")
End If
End Sub

Private Sub TxtEmpNo_LostFocus()
    If TxtEmpNo.Text <> "" Then 'And Edit_Flg = "A"
       Call FillDetails(TxtEmpNo.Text)
       'TxtCTC.SetFocus
    End If
End Sub

Private Sub TxtMobileNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "TxtMobileNo")
End If
End Sub

Private Sub TxtSearchEmployee_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "TxtSearchEmployee")
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


Private Sub TxtSFrom_LostFocus()
'mTotalDays = Day(DateSerial(Year(TxtSFrom.Value), Month(TxtSFrom.Value) + 1, 0))
TxtSFrom.Value = (Year(TxtSFrom.Value) & "-" & Month(TxtSFrom.Value) & "-01") ''& mTotalDays)
End Sub
Private Sub TxtSTo_LostFocus()
mTotalDays = Day(DateSerial(Year(TxtSTo.Value), Month(TxtSTo.Value) + 1, 0))
TxtSTo.Value = (Year(TxtSTo.Value) & "-" & Month(TxtSTo.Value) & "-" & mTotalDays)
End Sub
Private Sub TxtTxnDate_LostFocus()
mTotalDays = Day(DateSerial(Year(TxtTxnDate.Value), Month(TxtTxnDate.Value) + 1, 0))
TxtTxnDate.Value = (Year(TxtTxnDate.Value) & "-" & Month(TxtTxnDate.Value) & "-" & mTotalDays)
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

Private Sub FillDetails(strEmpNo_ As Variant)
RsMst_Employee.MoveFirst
If strEmpNo_ = "" Then
   RsMst_Employee.Find "EmployeeID = " & mEmployeeID & ""
Else
   RsMst_Employee.Find "EmployeeNo= '" & strEmpNo_ & "'"
End If
If Not RsMst_Employee.EOF Then
   mEmployeeID = RsMst_Employee!EmployeeID
   CmbEmployeeID.Text = RsMst_Employee!EmployeeName & "~" & RsMst_Employee!EmployeeNo
   TxtEmpID = RsMst_Employee!EmployeeID
   TxtEmpNo = RsMst_Employee!EmployeeNo
   TxtEmpName = RsMst_Employee!EmployeeName
   TxtEmpEmailID = IIf(IsNull(RsMst_Employee!EmployeeEmailID), "", RsMst_Employee!EmployeeEmailID)
   TxtMobileNo = IIf(IsNull(RsMst_Employee!MobileNo), "", RsMst_Employee!MobileNo)
   TxtAddress = IIf(IsNull(RsMst_Employee!Address), "", RsMst_Employee!Address)
   ChkAllocateExpense.Value = IIf(IsNull(RsMst_Employee!AllocateExpense), 0, RsMst_Employee!AllocateExpense)
   ChkCMPAllocation.Value = IIf(IsNull(RsMst_Employee!CMPAllocation), 0, RsMst_Employee!CMPAllocation)
   mOldDesignationID = RsMst_Employee!DesignationID
   mDesignationID = RsMst_Employee!DesignationID
   TxtDesignationID = GetDesignaion(mDesignationID)
   CmbDesignationID.Text = TxtDesignationID.Text
   mReportingDesignationID = GetReportingDesignation(mDesignationID)
   
   mDirectReportingID = IIf(IsNull(RsMst_Employee!DirectReportingTo), "", RsMst_Employee!DirectReportingTo)
   If mDirectReportingID <> "" Then
      TxtDirectReporting = GetEmployeeNonName(mDirectReportingID)
   Else
       mDirectReportingID = Null
   End If
   CmbDirectReporting.Text = TxtDirectReporting.Text
   mConnectedRAGID = RsMst_Employee!ConnectedToRAG
   If mConnectedRAGID <> 0 Then
      TxtConnectedRAG = GetEmployeeNonName(mConnectedRAGID)
   Else
      mConnectedRAGID = Null
   End If
   CmbConnectedRAG.Text = TxtConnectedRAG.Text
   mFlag = RsMst_Employee!Flag
   If mFlag = "Y" Then
      TxtFlag = "Active"
   Else
      TxtFlag = "De-Active"
   End If
   CmbFlag.Text = TxtFlag.Text
   mBaseLocationID = IIf(IsNull(RsMst_Employee!BaseLocationID), 0, RsMst_Employee!BaseLocationID)
   If mBaseLocationID <> 0 Then
      TxtBaseLocationID = GetLocationName(mBaseLocationID)
   Else
      mBaseLocationID = 0
   End If
   CmbBaseLocationID.Text = TxtBaseLocationID.Text
    
    
   If RsMst_Employee!SelfReporting = 1 Then
      Label25.Visible = True
      ChkSelfReporting.Visible = True
   Else
      Label25.Visible = False
      ChkSelfReporting.Visible = False
   End If
   ChkSelfReporting.Value = IIf(IsNull(RsMst_Employee!SelfReporting), 0, RsMst_Employee!SelfReporting)
   ChkSelfReporting.Enabled = False
Else
   MsgBox "Employee No does not exist."
End If
End Sub

