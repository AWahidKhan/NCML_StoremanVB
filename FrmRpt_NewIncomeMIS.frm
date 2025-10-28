VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmRpt_NewIncomeMIS 
   Caption         =   "New Income MIS"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   780
      Left            =   120
      TabIndex        =   15
      Top             =   0
      Width           =   15015
      Begin VB.Frame Frame12 
         Height          =   780
         Left            =   11280
         TabIndex        =   38
         Top             =   0
         Width           =   3735
         Begin VB.CheckBox Check1 
            Caption         =   "Gross pay"
            Height          =   255
            Left            =   2625
            TabIndex        =   42
            Top             =   450
            Width           =   1080
         End
         Begin VB.OptionButton OptIncome 
            Caption         =   "Income"
            Height          =   375
            Left            =   120
            TabIndex        =   41
            Top             =   390
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton OptBoth 
            Caption         =   "Both"
            Height          =   375
            Left            =   1905
            TabIndex        =   40
            Top             =   390
            Width           =   675
         End
         Begin VB.OptionButton OptCap 
            Caption         =   "Capacity"
            Height          =   375
            Left            =   960
            TabIndex        =   39
            Top             =   390
            Width           =   975
         End
         Begin VB.Label Label7 
            Caption         =   "Salary Allocation"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1080
            TabIndex        =   43
            Top             =   120
            Width           =   1695
         End
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
         Left            =   3135
         Sorted          =   -1  'True
         TabIndex        =   30
         Top             =   360
         Width           =   4455
      End
      Begin VB.CommandButton CmdGetReport 
         Caption         =   "Get Report"
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
         Left            =   7635
         TabIndex        =   18
         Top             =   345
         Width           =   1215
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
         Height          =   375
         Left            =   8850
         TabIndex        =   17
         Top             =   345
         Width           =   1215
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "Clear"
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
         Left            =   10065
         TabIndex        =   16
         Top             =   345
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker TxtFromDate 
         Height          =   330
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   582
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
         CurrentDate     =   39850
         MinDate         =   38353
      End
      Begin MSComCtl2.DTPicker TxtToDate 
         Height          =   330
         Left            =   1665
         TabIndex        =   20
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
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
         CurrentDate     =   39850
         MinDate         =   38353
      End
      Begin VB.Label lblRM 
         Caption         =   "RM Name"
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
         Left            =   4875
         TabIndex        =   31
         Top             =   105
         Width           =   1065
      End
      Begin VB.Label Label5 
         Caption         =   "#"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9480
         TabIndex        =   28
         Top             =   105
         Width           =   210
      End
      Begin VB.Label Label4 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7620
         TabIndex        =   27
         Top             =   105
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "- As per CMG"
         Height          =   255
         Left            =   9705
         TabIndex        =   26
         Top             =   105
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "- As per Store-Man"
         Height          =   255
         Left            =   7860
         TabIndex        =   25
         Top             =   105
         Width           =   1575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "To  Date"
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
         Left            =   1920
         TabIndex        =   22
         Top             =   120
         Width           =   780
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "From  Date"
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
         Left            =   360
         TabIndex        =   21
         Top             =   120
         Width           =   990
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8895
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   15690
      _Version        =   393216
      Tabs            =   11
      Tab             =   7
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "AM Wise Details *"
      TabPicture(0)   =   "FrmRpt_NewIncomeMIS.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Vendor Expense *"
      TabPicture(1)   =   "FrmRpt_NewIncomeMIS.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Godown Rent *"
      TabPicture(2)   =   "FrmRpt_NewIncomeMIS.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame4"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "NN-NR Income *"
      TabPicture(3)   =   "FrmRpt_NewIncomeMIS.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame5"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Employee Expense *"
      TabPicture(4)   =   "FrmRpt_NewIncomeMIS.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame6"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Employee Salary #"
      TabPicture(5)   =   "FrmRpt_NewIncomeMIS.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame7"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "Godowns without Agreement*"
      TabPicture(6)   =   "FrmRpt_NewIncomeMIS.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Frame8"
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "DP Income *"
      TabPicture(7)   =   "FrmRpt_NewIncomeMIS.frx":00C4
      Tab(7).ControlEnabled=   -1  'True
      Tab(7).Control(0)=   "Frame9"
      Tab(7).Control(0).Enabled=   0   'False
      Tab(7).ControlCount=   1
      TabCaption(8)   =   "Other Income *"
      TabPicture(8)   =   "FrmRpt_NewIncomeMIS.frx":00E0
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Frame10"
      Tab(8).ControlCount=   1
      TabCaption(9)   =   "CMG Expense"
      TabPicture(9)   =   "FrmRpt_NewIncomeMIS.frx":00FC
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "Frame11"
      Tab(9).ControlCount=   1
      TabCaption(10)  =   "Security Expense *"
      TabPicture(10)  =   "FrmRpt_NewIncomeMIS.frx":0118
      Tab(10).ControlEnabled=   0   'False
      Tab(10).Control(0)=   "Frame13"
      Tab(10).ControlCount=   1
      Begin VB.Frame Frame13 
         Height          =   7935
         Left            =   -74880
         TabIndex        =   44
         Top             =   720
         Width           =   14670
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid12 
            Height          =   7575
            Left            =   120
            TabIndex        =   45
            Top             =   240
            Width           =   14415
            _ExtentX        =   25426
            _ExtentY        =   13361
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
      End
      Begin VB.Frame Frame11 
         Height          =   8055
         Left            =   -74880
         TabIndex        =   34
         Top             =   1020
         Width           =   14670
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid10 
            Height          =   3975
            Left            =   120
            TabIndex        =   35
            Top             =   240
            Width           =   14415
            _ExtentX        =   25426
            _ExtentY        =   7011
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
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid11 
            Height          =   3375
            Left            =   120
            TabIndex        =   36
            Top             =   4560
            Width           =   14415
            _ExtentX        =   25426
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
         Begin VB.Label Label8 
            Caption         =   "CMG Salary"
            Height          =   255
            Left            =   135
            TabIndex        =   37
            Top             =   4290
            Width           =   1575
         End
      End
      Begin VB.Frame Frame7 
         Height          =   7935
         Left            =   -74880
         TabIndex        =   32
         Top             =   1080
         Width           =   14550
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid6 
            Height          =   7575
            Left            =   120
            TabIndex        =   33
            Top             =   240
            Width           =   14295
            _ExtentX        =   25215
            _ExtentY        =   13361
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
      End
      Begin VB.Frame Frame10 
         Height          =   7695
         Left            =   -74880
         TabIndex        =   23
         Top             =   1140
         Width           =   14655
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid9 
            Height          =   7335
            Left            =   120
            TabIndex        =   24
            Top             =   240
            Width           =   14415
            _ExtentX        =   25426
            _ExtentY        =   12938
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
      End
      Begin VB.Frame Frame2 
         Height          =   7815
         Left            =   -74880
         TabIndex        =   13
         Top             =   1080
         Width           =   14670
         Begin VB.TextBox Text1 
            Height          =   2175
            Left            =   6120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   29
            Text            =   "FrmRpt_NewIncomeMIS.frx":0134
            Top             =   1320
            Visible         =   0   'False
            Width           =   4215
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   7455
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   14415
            _ExtentX        =   25426
            _ExtentY        =   13150
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
      End
      Begin VB.Frame Frame3 
         Height          =   7935
         Left            =   -74880
         TabIndex        =   11
         Top             =   1080
         Width           =   14670
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
            Height          =   7695
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   14415
            _ExtentX        =   25426
            _ExtentY        =   13573
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
      End
      Begin VB.Frame Frame4 
         Height          =   7935
         Left            =   -74880
         TabIndex        =   9
         Top             =   1080
         Width           =   14670
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
            Height          =   7695
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   14415
            _ExtentX        =   25426
            _ExtentY        =   13573
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
      End
      Begin VB.Frame Frame5 
         Height          =   7935
         Left            =   -74880
         TabIndex        =   7
         Top             =   1080
         Width           =   14670
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid4 
            Height          =   7575
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   14415
            _ExtentX        =   25426
            _ExtentY        =   13361
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
      End
      Begin VB.Frame Frame6 
         Height          =   8055
         Left            =   -74880
         TabIndex        =   5
         Top             =   1080
         Width           =   14670
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid5 
            Height          =   7695
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   14415
            _ExtentX        =   25426
            _ExtentY        =   13573
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
      End
      Begin VB.Frame Frame8 
         Height          =   7935
         Left            =   -74880
         TabIndex        =   3
         Top             =   1080
         Width           =   14655
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid7 
            Height          =   7575
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   14415
            _ExtentX        =   25426
            _ExtentY        =   13361
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
      End
      Begin VB.Frame Frame9 
         Height          =   7935
         Left            =   120
         TabIndex        =   1
         Top             =   1080
         Width           =   14655
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid8 
            Height          =   7575
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   14415
            _ExtentX        =   25426
            _ExtentY        =   13361
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
      End
   End
End
Attribute VB_Name = "FrmRpt_NewIncomeMIS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mTotDays As Variant
Dim mDate As Variant
Dim mStartDate As Variant
Dim mEndDate As Variant
Dim mYear As Variant
Dim mMonthNo As Variant
Dim mTotOtherEmpSal As Double
Dim mTotOtherEmpID As String
Dim mDoublePaidAmt As Double
Dim mInvStartDate As Variant
Dim mInvEndDate As Variant
Dim mRow As Double
Dim mWCLocationIDs As String
Dim mEmployeeID As Variant
Dim mTotMISGroups, mTotIncomes, mTotExpenses As Variant
Dim mEmpExpCol, mVendExpCol, mRentCol, mEmpSalCol, mCOExpCol As Variant
Dim mTotCapacity, mTotIncome As Variant
Dim mAllTotCapacity, mAllTotIncome As Variant
Dim mCOExpense As Double
Dim mCMGExpense As Double
Dim mCMGSal As Double
Dim mRMFlag As Boolean
Dim mLocationIDs As String
Dim mNewDate As Variant
Dim mRowSecurity As Double
Private Sub Grid_Head()

Dim mCol As Variant

tmp = "Select Distinct MISExpenseGroup from Mst_ExpenseType"
Call TmpTable

tmp = "Select IEType from Mst_IncomeExpenseType Where IEFlag='I' And Flag='A'"
Call Tmp1Table

tmp = "Select IEType from Mst_IncomeExpenseType Where IEFlag='E' And Flag='A'"
Call Tmp2Table

mTotMISGroups = TmpRs.RecordCount
mTotIncomes = TmpRs1.RecordCount
mTotExpenses = TmpRs2.RecordCount

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 3
MSHFlexGrid1.Cols = 13 + mTotMISGroups + mTotIncomes + mTotExpenses + 2
MSHFlexGrid1.FixedCols = 4
MSHFlexGrid1.FixedRows = 2

MSHFlexGrid1.TextMatrix(1, 0) = "LocationID"
MSHFlexGrid1.TextMatrix(1, 1) = "RM"
MSHFlexGrid1.TextMatrix(1, 2) = "AM"
MSHFlexGrid1.TextMatrix(1, 3) = "Name of Location"
MSHFlexGrid1.TextMatrix(1, 4) = "Lease Capacity + Franchisee Stock (As on " & Format(TxtToDate.Value, "dd-MMM-yyyy") & ")"
MSHFlexGrid1.TextMatrix(1, 5) = "Total Income"
 
MSHFlexGrid1.row = 1
MSHFlexGrid1.Col = 5
MSHFlexGrid1.CellBackColor = vbYellow

MSHFlexGrid1.TextMatrix(1, 6) = "Total Expense"

MSHFlexGrid1.row = 1
MSHFlexGrid1.Col = 6
MSHFlexGrid1.CellBackColor = vbRed
MSHFlexGrid1.TextMatrix(1, 7) = "Gross Margin"

MSHFlexGrid1.row = 1
MSHFlexGrid1.Col = 7
MSHFlexGrid1.CellBackColor = vbBlue

MSHFlexGrid1.TextMatrix(1, 8) = "DP Income"
MSHFlexGrid1.row = 1
MSHFlexGrid1.Col = 8
MSHFlexGrid1.CellBackColor = vbYellow

MSHFlexGrid1.TextMatrix(1, 9) = "Non NCDEX Non Reservation Income"
MSHFlexGrid1.row = 1
MSHFlexGrid1.Col = 9
MSHFlexGrid1.CellBackColor = vbYellow

mCol = 9
If TmpRs1.RecordCount > 0 Then
   For r = 1 To TmpRs1.RecordCount
       mCol = mCol + 1
       MSHFlexGrid1.TextMatrix(1, mCol) = IIf(IsNull(TmpRs1!IEType), "", TmpRs1!IEType)
       MSHFlexGrid1.row = 1
       MSHFlexGrid1.Col = mCol
       MSHFlexGrid1.CellBackColor = vbYellow

       TmpRs1.MoveNext
   Next
End If
MSHFlexGrid1.TextMatrix(1, mCol + 1) = "Employee Claims (TA bills)"
MSHFlexGrid1.row = 1
MSHFlexGrid1.Col = mCol + 1
MSHFlexGrid1.CellBackColor = vbRed

mCol = mCol + 1
If TmpRs.RecordCount > 0 Then
   For r = 1 To TmpRs.RecordCount
       mCol = mCol + 1
       MSHFlexGrid1.TextMatrix(1, mCol) = "Vendor Expenses-" & IIf(IsNull(TmpRs!MISExpenseGroup), "", TmpRs!MISExpenseGroup)
       MSHFlexGrid1.row = 1
       MSHFlexGrid1.Col = mCol
       MSHFlexGrid1.CellBackColor = vbRed
       TmpRs.MoveNext
   Next
End If
MSHFlexGrid1.TextMatrix(1, mCol + 1) = "Lease Rent"
MSHFlexGrid1.row = 1
MSHFlexGrid1.Col = mCol + 1
MSHFlexGrid1.CellBackColor = vbRed
       
MSHFlexGrid1.TextMatrix(1, mCol + 2) = "Employee Salaries"
MSHFlexGrid1.row = 1
MSHFlexGrid1.Col = mCol + 2
MSHFlexGrid1.CellBackColor = vbRed
mCol = mCol + 2
If TmpRs2.RecordCount > 0 Then
   For r = 1 To TmpRs2.RecordCount
       mCol = mCol + 1
       MSHFlexGrid1.TextMatrix(1, mCol) = IIf(IsNull(TmpRs2!IEType), "", TmpRs2!IEType)
       TmpRs2.MoveNext
       MSHFlexGrid1.row = 1
       MSHFlexGrid1.Col = mCol
       MSHFlexGrid1.CellBackColor = vbRed
   Next
End If

MSHFlexGrid1.TextMatrix(1, mCol + 1) = "Corporate Office Expense"
MSHFlexGrid1.row = 1
MSHFlexGrid1.Col = mCol + 1
MSHFlexGrid1.CellBackColor = vbGreen

MSHFlexGrid1.TextMatrix(1, mCol + 2) = "Net Margin"
MSHFlexGrid1.row = 1
MSHFlexGrid1.Col = mCol + 2
MSHFlexGrid1.CellBackColor = vbBlue

mEmpExpCol = 9 + mTotIncomes + 1
mVendExpCol = 9 + mTotIncomes + 2
mRentCol = 9 + mTotIncomes + 1 + mTotMISGroups + 1
mEmpSalCol = 9 + mTotIncomes + 1 + mTotMISGroups + 2
mCOExpCol = MSHFlexGrid1.Cols - 2

MSHFlexGrid1.ColWidth(0) = 0
MSHFlexGrid1.ColWidth(1) = 2000
MSHFlexGrid1.ColWidth(2) = 1400
MSHFlexGrid1.ColWidth(3) = 1400
MSHFlexGrid1.ColWidth(4) = 1400
MSHFlexGrid1.ColWidth(5) = 1400
MSHFlexGrid1.ColWidth(6) = 1400
MSHFlexGrid1.ColWidth(7) = 1400
MSHFlexGrid1.ColWidth(8) = 1200
MSHFlexGrid1.ColWidth(9) = 1200
MSHFlexGrid1.ColWidth(10) = 1200
MSHFlexGrid1.ColWidth(11) = 1200
MSHFlexGrid1.ColWidth(12) = 1200
MSHFlexGrid1.ColWidth(13) = 1200

MSHFlexGrid1.RowHeight(1) = 900

End Sub

Private Sub CmdClear_Click()
Call Grid_Head
Call Grid_Head1
Call Grid_Head2
Call Grid_Head3
Call Grid_Head4
Call Grid_Head5
Call Grid_Head6
Call Grid_Head7
Call Grid_Head8
Call Grid_Head9
Call Grid_Head10
Call Grid_Head11

TxtFromDate = Date - 30
TxtToDate = Date
CmdExcel.Enabled = False
SSTab1.Tab = 0
End Sub

Private Sub CmdGetReport_Click()
Dim mCCExp, mBDoExp As Double

Dim aftTime, bfrTime, curTime As Variant


aftTime = "19.00"
bfrTime = "09.30"
curTime = Format(Now, "hh.mm")
Dim mRow1 As Double


'If CmbEmployeeID.Visible = True And CmbEmployeeID.Text = "" Then
'   'If LCase(Gen_UserRole) <> "power" Then
'   If curTime < aftTime And curTime > bfrTime Then
'      MsgBox "Please run this Report before 9.30 am or After 7.00 pm."
'      Exit Sub
'   End If
'   'End If
'End If

If MsgBox("This report will take 20 to 30 minutes to run, do you wan't to continue?", vbYesNo) = vbNo Then
   Exit Sub
End If

If TxtFromDate > TxtToDate Then
   MsgBox "FromDate Should be Less Than Todate."
   TxtFromDate.SetFocus
   Exit Sub
End If


If Month(TxtToDate.Value) > Month(Date) And Year(TxtToDate.Value) = Year(Date) Then
   MsgBox "Future  Date not allowed."
   TxtToDate.SetFocus
   Exit Sub
End If

mStartDate = TxtFromDate.Value
mEndDate = TxtToDate.Value

tmp = " Select ExpenseAmt from Mst_CoExpense Where FromDate<='" & mStartDate & "' And ToDate>='" & mEndDate & "'"
Call TmpTable

If TmpRs.RecordCount = 0 Then
   MsgBox "Date range should be in same Financial Year!!"
   Exit Sub
Else
   mCOExpense = TmpRs.Fields(0)
End If

Text1 = ""
Text1 = "Start Time : " & Now() & Chr(13)


mInvStartDate = Year(mStartDate) & "-" & Month(mStartDate) & "-" & 1
mTotDays = Day(DateSerial(Year(mEndDate), Month(mEndDate) + 1, 0))
mInvEndDate = Year(mEndDate) & "-" & Month(mEndDate) & "-" & mTotDays

mCMGExpense = GetCMGExpense
mCMGSal = GetCMGSalary

mAllTotCapacity = 0
mAllTotIncome = 0
 
 
'all Other income
tmp = "Select sum(Amount)As IncomeAmt"
tmp = tmp & " from Txn_OtherIncome TI inner Join Mst_location ML on TI.LocationID=ML.LocationID inner Join Mst_IncomeExpenseType MIE on MIE.IETypeID=TI.IETypeID Where TI.TxnDate>='" & mStartDate & "' And TI.TxnDate<='" & mEndDate & "' And MIE.Flag='A'"
Call TmpTable
mAllTotIncome = mAllTotIncome + IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))

'all reservation
tmp = " Select sum(ActualAmount)As 'IncomeAmt' "
tmp = tmp & " from Txn_ReservationInvoice TI"
tmp = tmp & " inner Join Mst_location ML on TI.LocationID=ML.LocationID"
tmp = tmp & " inner Join Mst_StorageContract MSC On TI.ContractID = MSC.ContractID"
tmp = tmp & " inner Join Mst_Client MC On MSC.ClientID = MC.ClientID"
tmp = tmp & " Where TI.YearMonth>='" & mStartDate & "' And TI.YearMonth<='" & mEndDate & "'"

Call TmpTable
mAllTotIncome = mAllTotIncome + IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))

'all Ad-Hoc

tmp = "Select sum(Amount)As IncomeAmt "
tmp = tmp & " from Txn_TempInvoice TI"
tmp = tmp & " inner Join Mst_location ML on TI.LocationID=ML.LocationID"
tmp = tmp & " inner Join Mst_IncomeExpenseType MIE on MIE.IETypeID=TI.InvoiceTypeID"
tmp = tmp & " inner Join Mst_Client MC On TI.ClientID = MC.ClientID"
tmp = tmp & " Where TI.InvoiceDate>='" & mStartDate & "' And TI.InvoiceDate<='" & mEndDate & "' And MIE.Flag='A' And TI.Flag = 'A' "

Call TmpTable

mAllTotIncome = mAllTotIncome + IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))


'all NNNR income
tmp = "Select Sum(TIG.InvoiceAmount) IncomeAmt"
tmp = tmp & " from  Txn_InvMonGSLDetail TIG Inner Join Mst_CMP MC on TIG.CMPID=MC.CMPID Inner Join MSt_Location ML on MC.LocationID=ML.LocationID Where TIG.ExchangeTypeID = 1 And (TIG.ContractID = 0 Or TIG.ContractID Is Null) And TIG.InvoiceDate >='" & mStartDate & "' And TIG.InvoiceDate <='" & mEndDate & "' And PINo is not Null "
Call TmpTable
mAllTotIncome = mAllTotIncome + IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))

'all DP income
tmp = "Select sum(TotalAmt)As IncomeAmt"
tmp = tmp & " from Mst_DPInvoiceUpload MDP inner Join Mst_location ML on MDp.LocationID=ML.LocationID Where MDP.MonthYear>='" & mStartDate & "' And MDP.MonthYear<='" & mEndDate & "'"
Call TmpTable
mAllTotIncome = mAllTotIncome + IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))

'All Godown Cap
tmp = "Select Sum(MG.GodownCapacity) as TotalCapacity From Mst_Godown MG "
tmp = tmp & " Where MG.UnitTypeID =1 And (MG.CloseDate is Null or MG.CloseDate >'" & mEndDate & "') And MG.StartDate <='" & mEndDate & "'"

Call TmpTable
mAllTotCapacity = mAllTotCapacity + IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))

'All Franchisee Stock
tmp = "Select Sum(GSL.BalanceWeight) as Stock from Mst_GSL GSL"
tmp = tmp & " Inner Join Mst_Godown MG on GSL.GodownID=MG.GodownID"
tmp = tmp & " Where (GSL.Flag <> 'E') AND (GSL.Flag <> 'Z') And (MG.CloseDate is Null or MG.CloseDate >'" & mEndDate & "') And MG.StartDate <='" & mEndDate & "' And MG.UnitTypeID=2"
Call TmpTable
mAllTotCapacity = mAllTotCapacity + IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))

Call Grid_Head

'--For Location Wise AM

tmp = "Select ML.LocationID,ML.LocationName,ML.StateID,EMP1.EmployeeID AS 'CCAMID',EMP1.EmployeeName +'~'+ EMP1.EmployeeNo as 'CCAM',MD.Designation ,  "
tmp = tmp & " EMP2.EmployeeID as 'AMRMID',EMP2.EmployeeName+'~'+EMP2.EmployeeNo  as 'AMRM'"
tmp = tmp & " ,EMP3.EmployeeID as 'RMManagmentID',EMP3.EmployeeName +'~'+EMP3.EmployeeNo +'~'+ MD2.Designation as 'RM/Managment'"
tmp = tmp & " ,EMP4.EmployeeID as 'ManagmentID',EMP4.EmployeeName +'~'+EMP4.EmployeeNo as Managment"
tmp = tmp & " From Mst_EMPCMPMapping ECM"
tmp = tmp & " Right Outer Join Mst_Employee EMP On ECM.EmployeeID = EMP.EMployeeID"
tmp = tmp & " Left Outer Join Mst_Employee EMP1 On EMP.DirectReportingTo = EMP1.EmployeeID"
tmp = tmp & " Left Outer Join Mst_Employee EMP2 On EMP1.DirectReportingTo = EMP2.EmployeeID"
tmp = tmp & " Left Outer Join Mst_Employee EMP3 On EMP2.DirectReportingTo = EMP3.EmployeeID"
tmp = tmp & " Left Outer Join Mst_Employee EMP4 On EMP3.DirectReportingTo = EMP4.EmployeeID"
tmp = tmp & " Inner Join Mst_Designation MD On EMP1.DesignationID = MD.DesignationID"
tmp = tmp & " Inner Join Mst_Designation MD1 On EMP2.DesignationID = MD1.DesignationID"
tmp = tmp & " Inner Join Mst_Designation MD2 On EMP3.DesignationID = MD2.DesignationID"
tmp = tmp & " Right Join Mst_Location ML On ECM.LocationID = ML.LocationID"
If Gen_WcLocation <> "" Then
   tmp = tmp & " Where ML." & Trim(Gen_WcLocation)
End If
tmp = tmp & " Group By ML.LocationID,ML.LocationName,ML.StateID,EMP1.EmployeeNo,EMP1.EmployeeName,EMP2.EmployeeNo,EMP2.EmployeeName,"
tmp = tmp & " EMP3.EmployeeNo , EMP3.EmployeeName, EMP4.EmployeeName, EMP4.EmployeeNo, MD.Designation, MD1.Designation, MD2.Designation"
tmp = tmp & " ,EMP1.EmployeeID,EMP2.EmployeeID,EMP3.EmployeeID,EMP4.EmployeeID"
tmp = tmp & " Order By EMP1.EmployeeID,EMP2.EmployeeID,EMP3.EmployeeID,EMP4.EmployeeID,ML.LocationName"

Call TmpTable

Text1 = Text1 & "Level 1 : " & Now() & Chr(13)

tmp = "Select ML.LocationID,ML.LocationName,ML.Am_EmployeeID,ML.StateID,EMP1.EmployeeID AS 'CCAMID',EMP1.EmployeeName +'~'+ EMP1.EmployeeNo as 'CCAM',MD.Designation ,  "
tmp = tmp & " EMP2.EmployeeID as 'AMRMID',EMP2.EmployeeName+'~'+EMP2.EmployeeNo  as 'AMRM'"
tmp = tmp & " ,EMP3.EmployeeID as 'RMManagmentID',EMP3.EmployeeName +'~'+EMP3.EmployeeNo +'~'+ MD2.Designation as 'RM/Managment'"
tmp = tmp & " ,EMP4.EmployeeID as 'ManagmentID',EMP4.EmployeeName +'~'+EMP4.EmployeeNo as Managment"
tmp = tmp & " From Mst_EMPCMPMapping ECM"
tmp = tmp & " Right Outer Join Mst_Employee EMP On ECM.EmployeeID = EMP.EMployeeID"
tmp = tmp & " Left Outer Join Mst_Employee EMP1 On EMP.DirectReportingTo = EMP1.EmployeeID"
tmp = tmp & " Left Outer Join Mst_Employee EMP2 On EMP1.DirectReportingTo = EMP2.EmployeeID"
tmp = tmp & " Left Outer Join Mst_Employee EMP3 On EMP2.DirectReportingTo = EMP3.EmployeeID"
tmp = tmp & " Left Outer Join Mst_Employee EMP4 On EMP3.DirectReportingTo = EMP4.EmployeeID"
tmp = tmp & " Inner Join Mst_Designation MD On EMP1.DesignationID = MD.DesignationID"
tmp = tmp & " Inner Join Mst_Designation MD1 On EMP2.DesignationID = MD1.DesignationID"
tmp = tmp & " Inner Join Mst_Designation MD2 On EMP3.DesignationID = MD2.DesignationID"
tmp = tmp & " Right Join Mst_Location ML On ECM.LocationID = ML.LocationID"
tmp = tmp & " Where EMP1.EmployeeID Is Null"
'If Gen_WcLocation <> "" Then
'   tmp = tmp & " And  ML." & Trim(Gen_WcLocation)
'End If
tmp = tmp & " Group By ML.LocationID,ML.LocationName,ML.Am_EmployeeID,ML.StateID,EMP1.EmployeeNo,EMP1.EmployeeName,EMP2.EmployeeNo,EMP2.EmployeeName,"
tmp = tmp & " EMP3.EmployeeNo , EMP3.EmployeeName, EMP4.EmployeeName, EMP4.EmployeeNo, MD.Designation, MD1.Designation, MD2.Designation"
tmp = tmp & " ,EMP1.EmployeeID,EMP2.EmployeeID,EMP3.EmployeeID,EMP4.EmployeeID"
tmp = tmp & " Order By EMP1.EmployeeID,EMP2.EmployeeID,EMP3.EmployeeID,EMP4.EmployeeID,ML.LocationName"

Call Tmp4Table

Text1 = Text1 & "Level 2 : " & Now() & Chr(13)

'--For All AMs & SH


If CmbEmployeeID.Text <> "" Then
   'If Gen_DesignationID = 13 Then
   '  tmp = "Select * from Mst_Employee Where DesignationID = 2 or EmployeeID=" & mEmployeeID
   'Else
   tmp = "Select * from Mst_Employee Where EmployeeID=" & mEmployeeID
   
   Call TableMst_Employee(" Where (DesignationID= 3 or DesignationID= 13 ) And DirectReportingTo= " & mEmployeeID)
   'End If
Else
   tmp = "Select * from Mst_Employee Where (DesignationID= 2 or DesignationID= 13 ) And Flag='Y'  Order By DesignationID,EmployeeName"
   
   Call TableMst_Employee(" Where DesignationID= 3 or DesignationID= 13")
   
End If

Call Tmp1Table

Text1 = Text1 & "Level 3 : " & Now() & Chr(13)

mWCLocationIDs = "LocationID in ("

For I = 1 To TmpRs1.RecordCount
    
    TmpRs.Filter = "AMRMID=" & TmpRs1!EmployeeID
    If RsMst_Employee.RecordCount > 0 Then
       RsMst_Employee.MoveFirst
       mRMFlag = False
       For m = 1 To RsMst_Employee.RecordCount
           Call AddLocations(RsMst_Employee!EmployeeID, TmpRs1!EmployeeID)
           If RsMst_Employee!DesignationID = 13 And RsMst_Employee!DirectReportingTo = TmpRs1!EmployeeID Then
              tmp = "Select * from Mst_Employee Where DirectReportingTo=" & RsMst_Employee!EmployeeID & " And DesignationID=3"
              Call Tmp3Table
              For RM = 1 To TmpRs3.RecordCount
                  Call AddLocations(TmpRs3!EmployeeID, RsMst_Employee!EmployeeID)
                  TmpRs3.MoveNext
              Next
           End If
           
           If TmpRs1!DesignationID = 13 Then
             ' tmp = "Select * from Mst_Employee Where DirectReportingTo=" & TmpRs1!EmployeeID & " And DesignationID=3"
             ' Call Tmp3Table
             ' For RM = 1 To TmpRs3.RecordCount
               Call AddLocations(TmpRs1!EmployeeID, TmpRs1!DirectReportingTo)
             ' Next
           End If
           
           RsMst_Employee.MoveNext
       Next
         
       'MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
       If mRMFlag = True And Gen_DesignationID <> 3 And TmpRs1!DesignationID <> 13 Then
          'MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
          mRow1 = MSHFlexGrid1.Rows - 1
        
          MSHFlexGrid1.TextMatrix(mRow1, 1) = TmpRs1!EmployeeName & "~" & TmpRs1!EmployeeNo
          MSHFlexGrid1.TextMatrix(mRow1, 2) = ""
         
          MSHFlexGrid1.TextMatrix(mRow1, 3) = "Total"
         
          For C = 0 To MSHFlexGrid1.Cols - 1
              MSHFlexGrid1.row = mRow1
              MSHFlexGrid1.Col = C
              MSHFlexGrid1.CellBackColor = vbRed
          Next
          MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
       End If
    End If
    TmpRs.Filter = ""
    TmpRs1.MoveNext
Next

Text1 = Text1 & "Level 4 : " & Now() & Chr(13)

mWCLocationIDs = mWCLocationIDs & ")"

Text1 = Text1 & "Level 5 : " & Now() & Chr(13)

Call dispdata

Text1 = Text1 & "Level 6 : " & Now() & Chr(13)

Call FillLocWiseEmpExpense

Text1 = Text1 & "Level 7 : " & Now() & Chr(13)

Call FillEmployeeExpense

Text1 = Text1 & "Level 8 : " & Now() & Chr(13)

Call FillVendorExpense

Text1 = Text1 & "Level 10 : " & Now() & Chr(13)

'
Call FillSecurityExpense

Text1 = Text1 & "Level 11 : " & Now() & Chr(13)


Call FillGodownRent

Text1 = Text1 & "Level 12 : " & Now() & Chr(13)

Call FillNNNCIncome

Text1 = Text1 & "Level 13 : " & Now() & Chr(13)

Call FillDPIncome

Text1 = Text1 & "Level 14 : " & Now() & Chr(13)

Call FillOtherIncome

Text1 = Text1 & "Level 15 : " & Now() & Chr(13)

Call FillNotAgreeGodowns

Text1 = Text1 & "Level 16 : " & Now() & Chr(13)

For I = 2 To MSHFlexGrid1.Rows - 1
    For C = 8 To mEmpExpCol - 1
        MSHFlexGrid1.TextMatrix(I, 5) = Val(MSHFlexGrid1.TextMatrix(I, 5)) + Val(MSHFlexGrid1.TextMatrix(I, C))
    Next
Next

Call FillTotalSalaryLocWise

Text1 = Text1 & "Level 17 : " & Now() & Chr(13)

Call FillCOExpense

For I = 2 To MSHFlexGrid1.Rows - 1
    MSHFlexGrid1.TextMatrix(I, 5) = ""
Next
Text1 = Text1 & "Level 18 : " & Now() & Chr(13)

For I = 2 To MSHFlexGrid1.Rows - 1
    For C = 4 To MSHFlexGrid1.Cols - 1
        If C <> 4 Then
           MSHFlexGrid1.TextMatrix(I, C) = Round(Val(MSHFlexGrid1.TextMatrix(I, C)) / 1000)
        End If
        If MSHFlexGrid1.TextMatrix(I, 3) = "Total" And MSHFlexGrid1.TextMatrix(I, 2) <> "" Then
           MSHFlexGrid1.TextMatrix(I, C) = MSHFlexGrid1.TextMatrix(0, C)
           MSHFlexGrid1.TextMatrix(0, C) = ""
        ElseIf MSHFlexGrid1.TextMatrix(I, 3) <> "Total" And MSHFlexGrid1.TextMatrix(I, 2) <> "" Then
           MSHFlexGrid1.TextMatrix(0, C) = Val(MSHFlexGrid1.TextMatrix(0, C)) + Val(MSHFlexGrid1.TextMatrix(I, C))
        End If
    Next
Next

Text1 = Text1 & "Level 19 : " & Now() & Chr(13)

For I = 2 To MSHFlexGrid1.Rows - 1
    For C = 4 To MSHFlexGrid1.Cols - 1
        If MSHFlexGrid1.TextMatrix(I, 3) = "Total" And MSHFlexGrid1.TextMatrix(I, 2) = "" Then
           MSHFlexGrid1.TextMatrix(I, C) = MSHFlexGrid1.TextMatrix(0, C)
           MSHFlexGrid1.TextMatrix(0, C) = ""
        ElseIf MSHFlexGrid1.TextMatrix(I, 3) = "Total" And MSHFlexGrid1.TextMatrix(I, 2) <> "" Then
           MSHFlexGrid1.TextMatrix(0, C) = Val(MSHFlexGrid1.TextMatrix(0, C)) + Val(MSHFlexGrid1.TextMatrix(I, C))
        End If
    Next
Next
For C = 4 To MSHFlexGrid1.Cols - 1
    MSHFlexGrid1.TextMatrix(0, C) = ""
Next

For I = 2 To MSHFlexGrid1.Rows - 1
    For C = 8 To mEmpExpCol - 1
        MSHFlexGrid1.TextMatrix(I, 5) = Val(MSHFlexGrid1.TextMatrix(I, 5)) + Val(MSHFlexGrid1.TextMatrix(I, C))
    Next
    
    For C = mEmpExpCol To MSHFlexGrid1.Cols - 3
        MSHFlexGrid1.TextMatrix(I, 6) = Val(MSHFlexGrid1.TextMatrix(I, 6)) + Val(MSHFlexGrid1.TextMatrix(I, C))
    Next
    MSHFlexGrid1.TextMatrix(I, 7) = Val(MSHFlexGrid1.TextMatrix(I, 5)) - Val(MSHFlexGrid1.TextMatrix(I, 6))
    MSHFlexGrid1.TextMatrix(I, MSHFlexGrid1.Cols - 1) = Val(MSHFlexGrid1.TextMatrix(I, 7)) - Val(MSHFlexGrid1.TextMatrix(I, mCOExpCol))
    If MSHFlexGrid1.TextMatrix(I, 3) <> "Total" And MSHFlexGrid1.TextMatrix(I, 2) <> "" Then
       For C = 4 To MSHFlexGrid1.Cols - 1
           MSHFlexGrid1.TextMatrix(0, C) = Val(MSHFlexGrid1.TextMatrix(0, C)) + Val(MSHFlexGrid1.TextMatrix(I, C))
       Next
    End If
Next
CmdExcel.Enabled = True
End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Frame1.Enabled = False
   Frame2.Enabled = False
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If

Frame1.Enabled = True
Frame2.Enabled = True
mNewDate = CDate("2011-04-01")
If Gen_DesignationID = 3 Then
   CmbEmployeeID.Text = ""
   lblRM.Visible = False
   CmbEmployeeID.Visible = False
Else
   lblRM.Visible = True
   CmbEmployeeID.Visible = True
   If Gen_DesignationID = 2 Or Gen_DesignationID = 13 Then
      CmbEmployeeID.Text = GetEmployeeNonName(Gen_EmployeeID)
      CmbEmployeeID.Enabled = False
      mEmployeeID = Gen_EmployeeID
   Else
      CmbEmployeeID.Text = ""
      CmbEmployeeID.Enabled = True
   End If
End If
If LCase(UserType) = "power" Or LCase(UserType) = "superadmin" Then
   Frame12.Visible = True
Else
   Frame12.Visible = False
End If
Call CmdClear_Click
CmdExcel.Enabled = True
End Sub
Private Sub Grid_Head1()

MSHFlexGrid2.Clear
MSHFlexGrid2.Rows = 3
MSHFlexGrid2.Cols = 11
MSHFlexGrid2.FixedCols = 2
MSHFlexGrid2.FixedRows = 2
MSHFlexGrid2.TextMatrix(1, 0) = "Claim ID"
MSHFlexGrid2.TextMatrix(1, 1) = "Location"
MSHFlexGrid2.TextMatrix(1, 2) = "Vendor Name"
MSHFlexGrid2.TextMatrix(1, 3) = "Expense Type"
MSHFlexGrid2.TextMatrix(1, 4) = "MIS ExpenseGroup"
MSHFlexGrid2.TextMatrix(1, 5) = "Expense Claim From the Month"
MSHFlexGrid2.TextMatrix(1, 6) = "Expense Claim To the Month"
MSHFlexGrid2.TextMatrix(1, 7) = "Claimed Amount"
MSHFlexGrid2.TextMatrix(1, 8) = "Expense for the month"
MSHFlexGrid2.TextMatrix(1, 9) = "Claim Amount for the Month"
MSHFlexGrid2.TextMatrix(1, 10) = "Status"
MSHFlexGrid2.ColWidth(0) = 800
MSHFlexGrid2.ColWidth(1) = 2000
MSHFlexGrid2.ColWidth(2) = 2000
MSHFlexGrid2.ColWidth(3) = 1400
MSHFlexGrid2.ColWidth(4) = 1400
MSHFlexGrid2.ColWidth(5) = 0
MSHFlexGrid2.ColWidth(6) = 0
MSHFlexGrid2.ColWidth(7) = 0
MSHFlexGrid2.ColWidth(8) = 1400
MSHFlexGrid2.ColWidth(9) = 1200
MSHFlexGrid2.ColWidth(10) = 1200

MSHFlexGrid2.RowHeight(1) = 900

End Sub
Private Sub Grid_Head11()

MSHFlexGrid12.Clear
MSHFlexGrid12.Rows = 3
MSHFlexGrid12.Cols = 7
MSHFlexGrid12.FixedCols = 2
MSHFlexGrid12.FixedRows = 2

MSHFlexGrid12.TextMatrix(1, 0) = "Bill ID"
MSHFlexGrid12.TextMatrix(1, 1) = "Location"
MSHFlexGrid12.TextMatrix(1, 2) = "Security Agency"
MSHFlexGrid12.TextMatrix(1, 3) = "Expense Type"
MSHFlexGrid12.TextMatrix(1, 4) = "Expense for the month"
MSHFlexGrid12.TextMatrix(1, 5) = "Claim Amount for the Month"
MSHFlexGrid12.TextMatrix(1, 6) = "Status"

MSHFlexGrid12.ColWidth(0) = 800
MSHFlexGrid12.ColWidth(1) = 2000
MSHFlexGrid12.ColWidth(2) = 2000
MSHFlexGrid12.ColWidth(3) = 1400
MSHFlexGrid12.ColWidth(4) = 1400
MSHFlexGrid12.ColWidth(5) = 1200
MSHFlexGrid12.ColWidth(6) = 1200

MSHFlexGrid12.RowHeight(1) = 900

End Sub



Private Sub Grid_Head2()

MSHFlexGrid3.Clear
MSHFlexGrid3.Rows = 3
MSHFlexGrid3.Cols = 12
MSHFlexGrid3.FixedCols = 2
MSHFlexGrid3.FixedRows = 2
MSHFlexGrid3.TextMatrix(1, 0) = "LocationID"
MSHFlexGrid3.TextMatrix(1, 1) = "Location"
MSHFlexGrid3.TextMatrix(1, 2) = "CMP Name"
MSHFlexGrid3.TextMatrix(1, 3) = "Godown No"
MSHFlexGrid3.TextMatrix(1, 4) = "Lessor Name"
MSHFlexGrid3.TextMatrix(1, 5) = "Lessor NavCode"
'MSHFlexGrid3.TextMatrix(1, 6) = "Rent Start Date"
'MSHFlexGrid3.TextMatrix(1, 7) = "Rent End Date"
'MSHFlexGrid3.TextMatrix(1, 8) = "Rent Amount"
'MSHFlexGrid3.TextMatrix(1, 9) = "Rent From Date"
MSHFlexGrid3.TextMatrix(1, 6) = "Rent for the Month"
MSHFlexGrid3.TextMatrix(1, 7) = "Rent Days"
MSHFlexGrid3.TextMatrix(1, 8) = "Rent Amount for the Month"
MSHFlexGrid3.TextMatrix(1, 9) = "Paid Rent Amount"
MSHFlexGrid3.TextMatrix(1, 10) = "Duplicate Rent Amount"
MSHFlexGrid3.TextMatrix(1, 11) = "Status"

MSHFlexGrid3.ColWidth(0) = 0
MSHFlexGrid3.ColWidth(1) = 2000
MSHFlexGrid3.ColWidth(2) = 1400
MSHFlexGrid3.ColWidth(3) = 1400
MSHFlexGrid3.ColWidth(4) = 1400
MSHFlexGrid3.ColWidth(5) = 1400
MSHFlexGrid3.ColWidth(6) = 1400
MSHFlexGrid3.ColWidth(7) = 1400
MSHFlexGrid3.ColWidth(8) = 1200
MSHFlexGrid3.RowHeight(1) = 900

End Sub

Private Sub Grid_Head3()

MSHFlexGrid4.Clear
MSHFlexGrid4.Rows = 3
MSHFlexGrid4.Cols = 5
MSHFlexGrid4.FixedCols = 2
MSHFlexGrid4.FixedRows = 2
MSHFlexGrid4.TextMatrix(1, 0) = "LocationID"
MSHFlexGrid4.TextMatrix(1, 1) = "Location"
MSHFlexGrid4.TextMatrix(1, 2) = "Unit Type"
MSHFlexGrid4.TextMatrix(1, 3) = "For the Month"
MSHFlexGrid4.TextMatrix(1, 4) = "Amount"

MSHFlexGrid4.ColWidth(0) = 0
MSHFlexGrid4.ColWidth(1) = 2000
MSHFlexGrid4.ColWidth(2) = 1400
MSHFlexGrid4.ColWidth(3) = 1400
MSHFlexGrid4.ColWidth(4) = 1400
MSHFlexGrid4.RowHeight(1) = 900

End Sub
Private Sub Grid_Head7()

MSHFlexGrid8.Clear
MSHFlexGrid8.Rows = 3
MSHFlexGrid8.Cols = 4
MSHFlexGrid8.FixedCols = 2
MSHFlexGrid8.FixedRows = 2
MSHFlexGrid8.TextMatrix(1, 0) = "LocationID"
MSHFlexGrid8.TextMatrix(1, 1) = "Location"
MSHFlexGrid8.TextMatrix(1, 2) = "For the Month"
MSHFlexGrid8.TextMatrix(1, 3) = "Amount"
MSHFlexGrid8.ColWidth(0) = 0
MSHFlexGrid8.ColWidth(1) = 2000
MSHFlexGrid8.ColWidth(2) = 1400
MSHFlexGrid8.ColWidth(3) = 1400
MSHFlexGrid8.RowHeight(1) = 900

End Sub

Private Sub Grid_Head4()

MSHFlexGrid5.Clear
MSHFlexGrid5.Rows = 3
MSHFlexGrid5.Cols = 11
MSHFlexGrid5.FixedCols = 2
MSHFlexGrid5.FixedRows = 2
MSHFlexGrid5.TextMatrix(1, 0) = "ClaimID"
MSHFlexGrid5.TextMatrix(1, 1) = "Location"
MSHFlexGrid5.TextMatrix(1, 2) = "EmployeeName"
MSHFlexGrid5.TextMatrix(1, 3) = "Current Designation"
MSHFlexGrid5.TextMatrix(1, 4) = "Expesne Claim From the Month"
MSHFlexGrid5.TextMatrix(1, 5) = "Expense Claim To the Month"
MSHFlexGrid5.TextMatrix(1, 6) = "Expense Type"
MSHFlexGrid5.TextMatrix(1, 7) = "Claimed Amount"
MSHFlexGrid5.TextMatrix(1, 8) = "Expense for the Month"
MSHFlexGrid5.TextMatrix(1, 9) = "Claim Amount for the Month"
MSHFlexGrid5.TextMatrix(1, 10) = "Status"

MSHFlexGrid5.ColWidth(0) = 800
MSHFlexGrid5.ColWidth(1) = 2000
MSHFlexGrid5.ColWidth(2) = 2000
MSHFlexGrid5.ColWidth(3) = 1400
MSHFlexGrid5.ColWidth(4) = 0
MSHFlexGrid5.ColWidth(5) = 0
MSHFlexGrid5.ColWidth(6) = 1400
MSHFlexGrid5.ColWidth(7) = 0
MSHFlexGrid5.ColWidth(8) = 1400
MSHFlexGrid5.ColWidth(9) = 1200
MSHFlexGrid5.ColWidth(10) = 1200

MSHFlexGrid5.RowHeight(1) = 900

End Sub

Private Sub Grid_Head5()

MSHFlexGrid6.Clear
MSHFlexGrid6.Rows = 3
MSHFlexGrid6.Cols = 5
MSHFlexGrid6.FixedCols = 2
MSHFlexGrid6.FixedRows = 2
MSHFlexGrid6.TextMatrix(1, 0) = "Month"
MSHFlexGrid6.TextMatrix(1, 1) = "Employee Name"
MSHFlexGrid6.TextMatrix(1, 2) = "Designation"
MSHFlexGrid6.TextMatrix(1, 3) = "Location"
MSHFlexGrid6.TextMatrix(1, 4) = "Salary"


MSHFlexGrid6.ColWidth(0) = 1400
MSHFlexGrid6.ColWidth(1) = 2200
MSHFlexGrid6.ColWidth(2) = 1600
MSHFlexGrid6.ColWidth(3) = 1800
MSHFlexGrid6.ColWidth(4) = 1400
MSHFlexGrid6.RowHeight(1) = 900

End Sub

Private Sub Grid_Head6()

MSHFlexGrid7.Clear
MSHFlexGrid7.Rows = 3
MSHFlexGrid7.Cols = 6
MSHFlexGrid7.FixedCols = 1
MSHFlexGrid7.FixedRows = 2
MSHFlexGrid7.TextMatrix(1, 0) = "Location Name"
MSHFlexGrid7.TextMatrix(1, 1) = "GodownID"
MSHFlexGrid7.TextMatrix(1, 2) = "Godown No"
MSHFlexGrid7.TextMatrix(1, 3) = "Address"
MSHFlexGrid7.TextMatrix(1, 4) = "Start Date"
MSHFlexGrid7.TextMatrix(1, 5) = "Close Date"

MSHFlexGrid7.ColWidth(0) = 2200
MSHFlexGrid7.ColWidth(1) = 1000
MSHFlexGrid7.ColWidth(2) = 1400
MSHFlexGrid7.ColWidth(3) = 3000
MSHFlexGrid7.ColWidth(4) = 1500
MSHFlexGrid7.ColWidth(5) = 1500


MSHFlexGrid7.RowHeight(1) = 900

End Sub
Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_New Income MIS.xls")
Call Xls_Detail_Rpt
End Sub
Public Sub Xls_Detail_Rpt()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim oWS As Excel.Worksheet
Dim mRow As Integer
Dim mCol As Integer
MsgBox ("Wait till Next Message")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_New Income MIS.xls")

mRow = oWB.Worksheets.Count
If 15 - oWB.Worksheets.Count > 0 Then
   oWB.Worksheets.Add , oWB.Worksheets(mRow), 15 - oWB.Worksheets.Count
End If
Set oWS = oWB.Worksheets("Sheet1")
oWS.Name = "Report"
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = I + 1 ' + 6
    For C = 0 To MSHFlexGrid1.Cols - 1
           If I = 1 Then
               oWS.Cells(mRow, C + 1).WrapText = True
           End If
           MSHFlexGrid1.row = I
           MSHFlexGrid1.Col = C
           If MSHFlexGrid1.CellBackColor = vbGreen Then
              oWS.Cells(mRow, C + 1).Interior.Color = vbGreen
           ElseIf MSHFlexGrid1.CellBackColor = vbRed Then
              oWS.Cells(mRow, C + 1).Interior.Color = vbRed
           ElseIf MSHFlexGrid1.CellBackColor = vbYellow Then
              oWS.Cells(mRow, C + 1).Interior.Color = vbYellow
              ElseIf MSHFlexGrid1.CellBackColor = vbBlue Then
              oWS.Cells(mRow, C + 1).Interior.Color = vbBlue
           End If
        If InStr(MSHFlexGrid1.TextMatrix(I, C), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(I, C)) Then
              oWS.Cells(mRow, C + 1) = Format(MSHFlexGrid1.TextMatrix(I, C), "MMM/dd/yyyy")
           Else
              oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
           End If
        Else
           oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
        End If
        
    Next
Next
Set oWS = oWB.Worksheets("Sheet2")
oWS.Name = "Vendor Expense Details"
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid2.Rows - 1
    mRow = I + 1 ' + 6
    oWS.Range("f:f").ColumnWidth = 0
    oWS.Range("g:g").ColumnWidth = 0
    oWS.Range("h:h").ColumnWidth = 0
    For C = 0 To MSHFlexGrid2.Cols - 1
        
        If InStr(MSHFlexGrid2.TextMatrix(I, C), "/") > 0 Then
           If IsDate(MSHFlexGrid2.TextMatrix(I, C)) Then
              oWS.Cells(mRow, C + 1) = Format(MSHFlexGrid2.TextMatrix(I, C), "MMM/dd/yyyy")
           Else
              oWS.Cells(mRow, C + 1) = MSHFlexGrid2.TextMatrix(I, C)
           End If
        Else
           oWS.Cells(mRow, C + 1) = MSHFlexGrid2.TextMatrix(I, C)
        End If

    Next
Next


Set oWS = oWB.Worksheets("Sheet3")
oWS.Name = "Godown Rent Details"
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid3.Rows - 1
    mRow = I + 1 ' + 6
    For C = 0 To MSHFlexGrid3.Cols - 1
        If InStr(MSHFlexGrid3.TextMatrix(I, C), "/") > 0 Then
           If IsDate(MSHFlexGrid3.TextMatrix(I, C)) Then
              oWS.Cells(mRow, C + 1) = Format(MSHFlexGrid3.TextMatrix(I, C), "MMM/dd/yyyy")
           Else
              oWS.Cells(mRow, C + 1) = MSHFlexGrid3.TextMatrix(I, C)
           End If
        Else
           oWS.Cells(mRow, C + 1) = MSHFlexGrid3.TextMatrix(I, C)
        End If


    Next
Next

Set oWS = oWB.Worksheets("Sheet4")
oWS.Name = "Employee Expense"
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid5.Rows - 1
    mRow = I + 1 ' + 6
    oWS.Range("f:f").ColumnWidth = 0
    oWS.Range("e:e").ColumnWidth = 0
    oWS.Range("h:h").ColumnWidth = 0

    For C = 0 To MSHFlexGrid5.Cols - 1
        If InStr(MSHFlexGrid5.TextMatrix(I, C), "/") > 0 Then
           If IsDate(MSHFlexGrid5.TextMatrix(I, C)) Then
              oWS.Cells(mRow, C + 1) = Format(MSHFlexGrid5.TextMatrix(I, C), "MMM/dd/yyyy")
           Else
              oWS.Cells(mRow, C + 1) = MSHFlexGrid5.TextMatrix(I, C)
           End If
        Else
           oWS.Cells(mRow, C + 1) = MSHFlexGrid5.TextMatrix(I, C)
        End If
    Next
Next

Set oWS = oWB.Worksheets("Sheet5")
oWS.Name = "Salary Details"
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid6.Rows - 1
    mRow = I + 1 ' + 6
    For C = 0 To MSHFlexGrid6.Cols - 1
        If InStr(MSHFlexGrid6.TextMatrix(I, C), "/") > 0 Then
           If IsDate(MSHFlexGrid6.TextMatrix(I, C)) Then
              oWS.Cells(mRow, C + 1) = Format(MSHFlexGrid6.TextMatrix(I, C), "MMM/dd/yyyy")
           Else
              oWS.Cells(mRow, C + 1) = MSHFlexGrid6.TextMatrix(I, C)
           End If
        Else
           oWS.Cells(mRow, C + 1) = MSHFlexGrid6.TextMatrix(I, C)
        End If
    Next
Next

Set oWS = oWB.Worksheets("Sheet6")
oWS.Name = "Godowns without Agreement"
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid7.Rows - 1
    mRow = I + 1 ' + 6
    For C = 0 To MSHFlexGrid7.Cols - 1
        If InStr(MSHFlexGrid7.TextMatrix(I, C), "/") > 0 Then
           If IsDate(MSHFlexGrid7.TextMatrix(I, C)) Then
              oWS.Cells(mRow, C + 1) = Format(MSHFlexGrid7.TextMatrix(I, C), "MMM/dd/yyyy")
           Else
              oWS.Cells(mRow, C + 1) = MSHFlexGrid7.TextMatrix(I, C)
           End If
        Else
           oWS.Cells(mRow, C + 1) = MSHFlexGrid7.TextMatrix(I, C)
        End If
    Next
Next

Set oWS = oWB.Worksheets("Sheet7")
oWS.Name = "DP Income"
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid8.Rows - 1
    mRow = I + 1 ' + 6
    For C = 0 To MSHFlexGrid8.Cols - 1
        If InStr(MSHFlexGrid8.TextMatrix(I, C), "/") > 0 Then
           If IsDate(MSHFlexGrid8.TextMatrix(I, C)) Then
              oWS.Cells(mRow, C + 1) = Format(MSHFlexGrid8.TextMatrix(I, C), "MMM/dd/yyyy")
           Else
              oWS.Cells(mRow, C + 1) = MSHFlexGrid8.TextMatrix(I, C)
           End If
        Else
           oWS.Cells(mRow, C + 1) = MSHFlexGrid8.TextMatrix(I, C)
        End If
    Next
Next

Set oWS = oWB.Worksheets("Sheet8")
oWS.Name = "NN-NR Income"
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid4.Rows - 1
    mRow = I + 1 ' + 6
    For C = 0 To MSHFlexGrid4.Cols - 1
        If InStr(MSHFlexGrid4.TextMatrix(I, C), "/") > 0 Then
           If IsDate(MSHFlexGrid4.TextMatrix(I, C)) Then
              oWS.Cells(mRow, C + 1) = Format(MSHFlexGrid4.TextMatrix(I, C), "MMM/dd/yyyy")
           Else
              oWS.Cells(mRow, C + 1) = MSHFlexGrid4.TextMatrix(I, C)
           End If
        Else
           oWS.Cells(mRow, C + 1) = MSHFlexGrid4.TextMatrix(I, C)
        End If
    Next
Next

Set oWS = oWB.Worksheets("Sheet9")
oWS.Name = "Other Income"
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid9.Rows - 1
    mRow = I + 1 ' + 6
    For C = 0 To MSHFlexGrid9.Cols - 1
        If InStr(MSHFlexGrid9.TextMatrix(I, C), "/") > 0 Then
           If IsDate(MSHFlexGrid9.TextMatrix(I, C)) Then
              oWS.Cells(mRow, C + 1) = Format(MSHFlexGrid9.TextMatrix(I, C), "MMM/dd/yyyy")
           Else
              oWS.Cells(mRow, C + 1) = MSHFlexGrid9.TextMatrix(I, C)
           End If
        Else
           oWS.Cells(mRow, C + 1) = MSHFlexGrid9.TextMatrix(I, C)
        End If
    Next
Next

Set oWS = oWB.Worksheets("Sheet10")
oWS.Name = "CMG Expense"
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid10.Rows - 1
    mRow = I + 1 ' + 6
    For C = 0 To MSHFlexGrid10.Cols - 1
        If InStr(MSHFlexGrid10.TextMatrix(I, C), "/") > 0 Then
           If IsDate(MSHFlexGrid10.TextMatrix(I, C)) Then
              oWS.Cells(mRow, C + 1) = Format(MSHFlexGrid10.TextMatrix(I, C), "MMM/dd/yyyy")
           Else
              oWS.Cells(mRow, C + 1) = MSHFlexGrid10.TextMatrix(I, C)
           End If
        Else
           oWS.Cells(mRow, C + 1) = MSHFlexGrid10.TextMatrix(I, C)
        End If
    Next
Next


Set oWS = oWB.Worksheets("Sheet11")
oWS.Name = "CMG Salary"
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid11.Rows - 1
    mRow = I + 1 ' + 6
    For C = 0 To MSHFlexGrid11.Cols - 1
        If InStr(MSHFlexGrid11.TextMatrix(I, C), "/") > 0 Then
           If IsDate(MSHFlexGrid11.TextMatrix(I, C)) Then
              oWS.Cells(mRow, C + 1) = Format(MSHFlexGrid11.TextMatrix(I, C), "MMM/dd/yyyy")
           Else
              oWS.Cells(mRow, C + 1) = MSHFlexGrid11.TextMatrix(I, C)
           End If
        Else
           oWS.Cells(mRow, C + 1) = MSHFlexGrid11.TextMatrix(I, C)
        End If
    Next
Next


Set oWS = oWB.Worksheets("Sheet12")
oWS.Name = "Security Expense"
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid12.Rows - 1
    mRow = I + 1 ' + 6
    For C = 0 To MSHFlexGrid12.Cols - 1
        If InStr(MSHFlexGrid12.TextMatrix(I, C), "/") > 0 Then
           If IsDate(MSHFlexGrid12.TextMatrix(I, C)) Then
              oWS.Cells(mRow, C + 1) = Format(MSHFlexGrid12.TextMatrix(I, C), "MMM/dd/yyyy")
           Else
              oWS.Cells(mRow, C + 1) = MSHFlexGrid12.TextMatrix(I, C)
           End If
        Else
           oWS.Cells(mRow, C + 1) = MSHFlexGrid12.TextMatrix(I, C)
        End If
    Next
Next


oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True

End Sub

Private Function CheckLocation(mLocation_ As Variant) As Boolean
Dim Retval As Boolean
Retval = False
For r = 1 To MSHFlexGrid1.Rows - 1
    If LCase(MSHFlexGrid1.TextMatrix(r, 3)) = LCase(mLocation_) Then
       Retval = True
       Exit For
    End If
Next
CheckLocation = Retval
End Function

Private Sub dispdata()

'--For Location Wise DP Income

tmp = "Select MDP.LocationID,ML.LocationName,sum(TotalAmt)As IncomeAmt from Mst_DPInvoiceUpload MDP"
tmp = tmp & " inner Join Mst_location ML on MDp.LocationID=ML.LocationID"
tmp = tmp & " Where MDP.MonthYear>='" & mStartDate & "' And MDP.MonthYear<='" & mEndDate & "'"
If mWCLocationIDs <> "" Then
   tmp = tmp & " And MDP." & Trim(mWCLocationIDs)
End If
tmp = tmp & " Group By MDP.LocationID,ML.LocationName"

Call Tmp3Table


'--For Location wise NN-NC Income

tmp = "Select ML.LocationID,ML.LocationName,Sum(TIG.InvoiceAmount) IncomeAmt from  Txn_InvMonGSLDetail TIG"
tmp = tmp & " Inner Join Mst_CMP MC on TIG.CMPID=MC.CMPID"
tmp = tmp & " Inner Join MSt_Location ML on MC.LocationID=ML.LocationID"
tmp = tmp & " Where TIG.ExchangeTypeID = 1 And (TIG.ContractID = 0 Or TIG.ContractID Is Null) And PINo is not Null "
tmp = tmp & " And TIG.InvoiceDate >='" & mStartDate & "' And TIG.InvoiceDate <='" & mEndDate & "'"
If mWCLocationIDs <> "" Then
   tmp = tmp & " And ML." & Trim(mWCLocationIDs)
End If
tmp = tmp & " Group By ML.LocationID,ML.LocationName"

Call Tmp6Table

'--For Location wise Lease Capacity

'tmp = "Select ML.LocationName As Location,Sum(MG.GodownCapacity) as TotalCapacity From Mst_Location ML"
'tmp = tmp & " Left Outer Join Mst_CMP MC On ML.LocationID = MC.LocationID"
'tmp = tmp & " Left Outer Join Mst_Godown MG On MC.CMPID = MG.CMPID"
'tmp = tmp & " Where MG.UnitTypeID in (1,2,3) And (MG.CloseDate is Null or MG.CloseDate <='" & mEndDate & "') And MG.StartDate <='" & mEndDate & "'"
'tmp = tmp & " Group By ML.LocationName"

'-- After discus with Sankar Sir, Added on 19-07-2010

tmp = "Select ML.LocationName As Location,Sum(MG.GodownCapacity) as TotalCapacity From Mst_Location ML"
tmp = tmp & " Left Outer Join Mst_CMP MC On ML.LocationID = MC.LocationID"
tmp = tmp & " Left Outer Join Mst_Godown MG On MC.CMPID = MG.CMPID"
tmp = tmp & " Where MG.UnitTypeID =1 And (MG.CloseDate is Null or MG.CloseDate >'" & mEndDate & "') And MG.StartDate <='" & mEndDate & "'"
If mWCLocationIDs <> "" Then
   tmp = tmp & " And ML." & Trim(mWCLocationIDs)
End If

tmp = tmp & " Group By ML.LocationName"

Call Tmp1Table

'--For Location wise Franchisee Stock
tmp = "Select ML.LocationName,Sum(GSL.BalanceWeight) as Stock from Mst_GSL GSL"
tmp = tmp & " Inner Join Mst_Godown MG on GSL.GodownID=MG.GodownID"
tmp = tmp & " Inner join Mst_CMP  MC on MG.CMPID=MC.CMPID"
tmp = tmp & " Inner join Mst_Location  ML on MC.LocationID=ML.LocationID"
tmp = tmp & " Where (GSL.Flag <> 'E') AND (GSL.Flag <> 'Z') And (MG.CloseDate is Null or MG.CloseDate >'" & mEndDate & "') And MG.StartDate <='" & mEndDate & "' And MG.UnitTypeID=2"
If mWCLocationIDs <> "" Then
   tmp = tmp & " And ML." & Trim(mWCLocationIDs)
End If
tmp = tmp & " Group by ML.LocationName"

Call Tmp7Table

'''---For 7 Days income
''tmp = "Select TI.LocationID,ML.LocationName,TI.Flag,sum(Amount)As IncomeAmt from Txn_OtherIncome TI"
''tmp = tmp & " inner Join Mst_location ML on TI.LocationID=ML.LocationID"
''tmp = tmp & " Where TI.TxnDate>='" & mStartDate & "' And TI.TxnDate<='" & mEndDate & "'"
''If mWCLocationIDs <> "" Then
''   tmp = tmp & " And TI." & Trim(mWCLocationIDs)
''End If
''tmp = tmp & " Group By TI.LocationID,ML.LocationName,TI.Flag"
''
''Call Tmp2Table

'---For Other income
tmp = "Select TI.LocationID,ML.LocationName,MIE.IEType,sum(Amount)As IncomeAmt from Txn_OtherIncome TI"
tmp = tmp & " inner Join Mst_location ML on TI.LocationID=ML.LocationID"
tmp = tmp & " inner Join Mst_IncomeExpenseType MIE on MIE.IETypeID=TI.IETypeID"
tmp = tmp & " Where TI.TxnDate>='" & mStartDate & "' And TI.TxnDate<='" & mEndDate & "' And MIE.Flag='A'"
If mWCLocationIDs <> "" Then
   tmp = tmp & " And TI." & Trim(mWCLocationIDs)
End If
tmp = tmp & " Group By TI.LocationID,ML.LocationName,MIE.IEType"

Call Tmp2Table

'Reservation
tmp = " Select TI.LocationID,ML.LocationName,'Reservation Income' as 'IEType',sum(ActualAmount)As IncomeAmt"
tmp = tmp & " from Txn_ReservationInvoice TI"
tmp = tmp & " inner Join Mst_location ML on TI.LocationID=ML.LocationID"
tmp = tmp & " Where TI.YearMonth>='" & mStartDate & "' And TI.YearMonth<='" & mEndDate & "' "
If mWCLocationIDs <> "" Then
   tmp = tmp & " And TI." & Trim(mWCLocationIDs)
End If
tmp = tmp & " Group By TI.LocationID,ML.LocationName"
Call Tmp4Table

'ADHoc
tmp = " Select TI.LocationID,ML.LocationName,MIE.IEType,sum(Amount)As IncomeAmt"
tmp = tmp & " from Txn_TempInvoice TI"
tmp = tmp & " inner Join Mst_location ML on TI.LocationID=ML.LocationID"
tmp = tmp & " inner Join Mst_IncomeExpenseType MIE on MIE.IETypeID=TI.InvoiceTypeID"
tmp = tmp & " Where TI.InvoiceDate>='" & mStartDate & "' And TI.InvoiceDate<='" & mEndDate & "' And MIE.Flag='A' And TI.Flag = 'A' "
If mWCLocationIDs <> "" Then
   tmp = tmp & " And TI." & Trim(mWCLocationIDs)
End If
tmp = tmp & " Group By TI.LocationID,ML.LocationName,MIE.IEType"

Call Tmp5Table


For r = 2 To MSHFlexGrid1.Rows - 1
    If MSHFlexGrid1.TextMatrix(r, 3) <> "" Then
    '--For Lease Capacity
       If TmpRs1.RecordCount > 0 Then
          TmpRs1.MoveFirst
       End If
       TmpRs1.Find "Location='" & MSHFlexGrid1.TextMatrix(r, 3) & "'"
       If Not TmpRs1.EOF Then
          MSHFlexGrid1.TextMatrix(r, 4) = IIf(IsNull(TmpRs1!TotalCapacity), 0, Format(TmpRs1!TotalCapacity, "######0"))
       End If
       
    '--For Franchisee Stock
       If TmpRs7.RecordCount > 0 Then
          TmpRs7.MoveFirst
       End If
       TmpRs7.Find "LocationName='" & MSHFlexGrid1.TextMatrix(r, 3) & "'"
       If Not TmpRs7.EOF Then
          MSHFlexGrid1.TextMatrix(r, 4) = Val(MSHFlexGrid1.TextMatrix(r, 4)) + IIf(IsNull(TmpRs7!Stock), 0, Format(TmpRs7!Stock, "#####0"))
       End If

       
     '--For DP Income
       If TmpRs3.RecordCount > 0 Then
          TmpRs3.MoveFirst
       End If
       TmpRs3.Find "LocationName='" & MSHFlexGrid1.TextMatrix(r, 3) & "'"
       If Not TmpRs3.EOF Then
          MSHFlexGrid1.TextMatrix(r, 8) = Format(IIf(IsNull(TmpRs3!IncomeAmt), 0, TmpRs3!IncomeAmt), "######0.00")
       End If
       
     '--For NN-NC Income
       If TmpRs6.RecordCount > 0 Then
          TmpRs6.MoveFirst
       End If
       TmpRs6.Find "LocationName='" & MSHFlexGrid1.TextMatrix(r, 3) & "'"
       If Not TmpRs6.EOF Then
          MSHFlexGrid1.TextMatrix(r, 9) = Format(IIf(IsNull(TmpRs6!IncomeAmt), 0, TmpRs6!IncomeAmt), "######0.00")
       End If
       
     '--Reservation
       If TmpRs4.RecordCount > 0 Then
          TmpRs4.MoveFirst
       End If
       TmpRs4.Find "LocationName='" & MSHFlexGrid1.TextMatrix(r, 3) & "'"
       If Not TmpRs4.EOF Then
           For C = 10 To mTotIncomes + 10
               If LCase(MSHFlexGrid1.TextMatrix(1, C)) = "reservation income" Then 'LCase(MSHFlexGrid1.TextMatrix(1, c)) Then
                  MSHFlexGrid1.TextMatrix(r, C) = Format(IIf(IsNull(TmpRs4!IncomeAmt), 0, TmpRs4!IncomeAmt), "######0.00")
               End If
           Next
       End If
       
       'Adhoc
       If TmpRs5.RecordCount > 0 Then
          TmpRs5.MoveFirst
       End If
       TmpRs5.Filter = "LocationName='" & MSHFlexGrid1.TextMatrix(r, 3) & "'"
       For I = 1 To TmpRs5.RecordCount
           For C = 10 To mTotIncomes + 10
               If LCase(TmpRs5!IEType) = LCase(MSHFlexGrid1.TextMatrix(1, C)) Then
                  MSHFlexGrid1.TextMatrix(r, C) = Val(MSHFlexGrid1.TextMatrix(r, C)) + Format(IIf(IsNull(TmpRs5!IncomeAmt), 0, TmpRs5!IncomeAmt), "######0.00")
               End If
           Next
           TmpRs5.MoveNext
       Next
       TmpRs5.Filter = ""
       
       
       If TmpRs2.RecordCount > 0 Then
          TmpRs2.MoveFirst
       End If
       TmpRs2.Filter = "LocationName='" & MSHFlexGrid1.TextMatrix(r, 3) & "'"
       For I = 1 To TmpRs2.RecordCount
           For C = 10 To mTotIncomes + 10
               If LCase(TmpRs2!IEType) = LCase(MSHFlexGrid1.TextMatrix(1, C)) Then
                  MSHFlexGrid1.TextMatrix(r, C) = Val(MSHFlexGrid1.TextMatrix(r, C)) + Format(IIf(IsNull(TmpRs2!IncomeAmt), 0, TmpRs2!IncomeAmt), "######0.00")
               End If
           Next
           TmpRs2.MoveNext
       Next
       TmpRs2.Filter = ""
    End If
Next
End Sub

Private Sub FillLocWiseEmpExpense()
Dim mFlag As Boolean
Dim mMonths, mExpAmt, mTotMonths As Variant

If mEndDate < mNewDate Then
    tmp = "Select TCS.LocationID,ML.LocationName,Sum(TCS.Amount) As TotalExpenseAmount,TCS.ExpenseFrom,TCS.ExpenseTo"
    tmp = tmp & " From Txn_ClaimSubmission TCS"
    tmp = tmp & " Inner Join Mst_Employee ME On TCS.EmployeeID = ME.EmployeeID"
    tmp = tmp & " Inner Join Mst_Designation MD On ME.DesignationID = MD.DesignationID"
    tmp = tmp & " Inner Join Mst_Location ML on TCS.LocationID=ML.LocationID"
    tmp = tmp & " Where TCS.ExpenseFrom <= '" & mEndDate & "' And TCS.ExpenseTo >= '" & mStartDate & "' And TCS.Flag <> 'R' And MD.MISFlag='Y'" 'ME.DesignationID in (5,6,8,4,9,3,2)"
    If mWCLocationIDs <> "" Then
       tmp = tmp & " And TCS." & Trim(mWCLocationIDs)
    End If
    tmp = tmp & " Group By TCS.LocationID,ML.LocationName,TCS.ExpenseFrom,TCS.ExpenseTo"
    
    Call Tmp2Table
    
    mFlag = False
    If TmpRs2.RecordCount > 0 Then
       For I = 1 To TmpRs2.RecordCount
           mMonths = DateDiff("m", mStartDate, mEndDate) + 1
           mTotMonths = DateDiff("m", TmpRs2!ExpenseFrom, TmpRs2!ExpenseTo) + 1
           mExpAmt = IIf(IsNull(TmpRs2!TotalExpenseAmount), 0, TmpRs2!TotalExpenseAmount) / mTotMonths
           
           If TmpRs2!ExpenseFrom < mStartDate Then
              mTotMonths = DateDiff("m", mStartDate, TmpRs2!ExpenseTo) + 1
           ElseIf TmpRs2!ExpenseTo > mEndDate Then
               mTotMonths = DateDiff("m", TmpRs2!ExpenseFrom, mEndDate) + 1
           Else
               mTotMonths = DateDiff("m", TmpRs2!ExpenseFrom, TmpRs2!ExpenseTo) + 1
           End If
           
           If mTotMonths > mMonths Then
              mTotExpAmt = mExpAmt * mMonths
           Else
              mTotExpAmt = mExpAmt * mTotMonths
           End If
           For r = 1 To MSHFlexGrid1.Rows - 1
               If MSHFlexGrid1.TextMatrix(r, 3) = TmpRs2!Locationname Then
                  MSHFlexGrid1.TextMatrix(r, mEmpExpCol) = Val(MSHFlexGrid1.TextMatrix(r, mEmpExpCol)) + Format(mTotExpAmt, "######0.00")
                  Exit For
               End If
           Next r
           TmpRs2.MoveNext
       Next
    End If
Else

    tmp = "Select TCS.LocationID,Sum(TCS.AllocatedAmt) As TotalExpenseAmount,TC.ExpenseFrom,TC.ExpenseTo"
    tmp = tmp & " From Txn_ClaimAllocationDetail TCS"
    tmp = tmp & " inner Join Txn_ClaimSubmission TC on TCS.ClaimID=TC.ClaimID"
    tmp = tmp & " Inner Join Mst_Employee ME On TC.EmployeeID = ME.EmployeeID"
    tmp = tmp & " Inner Join Mst_Designation MD On ME.DesignationID = MD.DesignationID"
    tmp = tmp & " Where TC.ExpenseFrom <= '" & mEndDate & "' And TC.ExpenseTo >= '" & mStartDate & "' And TC.Flag <> 'R' And TC.DepartmentID<>38 And TCS.SegmentID=1"
    If mWCLocationIDs <> "" Then
       tmp = tmp & " And TCS." & Trim(mWCLocationIDs)
    End If
    tmp = tmp & " Group By TCS.LocationID,TCS.LocationID,TC.ExpenseFrom,TC.ExpenseTo"
    
    Call TmpTableHRPro
    mFlag = False
    If TmpRsHRPro.RecordCount > 0 Then
       For I = 1 To TmpRsHRPro.RecordCount
           mMonths = DateDiff("m", mStartDate, mEndDate) + 1
           mTotMonths = DateDiff("m", TmpRsHRPro!ExpenseFrom, TmpRsHRPro!ExpenseTo) + 1
           mExpAmt = IIf(IsNull(TmpRsHRPro!TotalExpenseAmount), 0, TmpRsHRPro!TotalExpenseAmount) / mTotMonths
           
           If TmpRsHRPro!ExpenseFrom < mStartDate Then
              mTotMonths = DateDiff("m", mStartDate, TmpRsHRPro!ExpenseTo) + 1
           ElseIf TmpRsHRPro!ExpenseTo > mEndDate Then
               mTotMonths = DateDiff("m", TmpRsHRPro!ExpenseFrom, mEndDate) + 1
           Else
               mTotMonths = DateDiff("m", TmpRsHRPro!ExpenseFrom, TmpRsHRPro!ExpenseTo) + 1
           End If
           
           If mTotMonths > mMonths Then
              mTotExpAmt = mExpAmt * mMonths
           Else
              mTotExpAmt = mExpAmt * mTotMonths
           End If
           For r = 2 To MSHFlexGrid1.Rows - 1
               If MSHFlexGrid1.TextMatrix(r, 3) = GetLocationName(TmpRsHRPro!LocationID) Then
                  MSHFlexGrid1.TextMatrix(r, mEmpExpCol) = Val(MSHFlexGrid1.TextMatrix(r, mEmpExpCol)) + Format(mTotExpAmt, "######0.00")
                  Exit For
               End If
           Next r
           TmpRsHRPro.MoveNext
       Next
    End If
End If
End Sub

Private Sub FillEmployeeExpense()
Dim mRentDays, mTotExpAmt, mExpAmt As Double
Dim mRentStartDate, mRentEndDate As Variant
Dim mSStartDate, mSEndDate As Variant
Dim mRentStart, mRentEnd As Variant
Dim mPaidRent  As Double
Dim mRow As Integer
Dim mTotMonths As Double

If mEndDate < mNewDate Then
    tmp = "Select TC.ClaimRowID,ML.LocationName,ME.EmployeeID, ME.EmployeeName + '~' + ME.EmployeeNo  as Employee,MD.Designation,TC.ExpenseFrom,TC.ExpenseTo,MEx.ExpenseType, Sum(TCS.Amount) As TotalExpenseAmount,TC.Remarks,Me.AllocateExpense,TC.Flag"
    tmp = tmp & " From Txn_ClaimSubmissionDetail TCS"
    tmp = tmp & " inner Join Txn_ClaimSubmission TC on TCS.ClaimRowID=TC.ClaimRowID"
    tmp = tmp & " inner Join Mst_ExpenseType MEx on TCS.ExpenseTypeID=MEx.ExpenseTypeID"
    tmp = tmp & " Inner Join Mst_Employee ME On TC.EmployeeID = ME.EmployeeID"
    tmp = tmp & " Inner Join Mst_Designation MD On ME.DesignationID = MD.DesignationID"
    tmp = tmp & " Inner Join Mst_Location ML On TC.LocationID = ML.LocationID"
    tmp = tmp & " Where TC.ExpenseFrom <= '" & mEndDate & "' And TC.ExpenseTo >= '" & mStartDate & "' And TC.Flag <> 'R' And MD.MISFlag='Y'" 'ME.DesignationID in (5,6,8,4,9,3,2)"
    If mWCLocationIDs <> "" Then
       tmp = tmp & " And TC." & Trim(mWCLocationIDs)
    End If
    tmp = tmp & " Group By TC.ClaimRowID,ML.LocationName,ME.EmployeeID,ME.EmployeeNo,ME.EmployeeName,ME.AllocateExpense,MD.Designation,TC.ExpenseFrom,TC.ExpenseTo,MEx.ExpenseType ,TC.Remarks,TC.Flag "
    
    Call Tmp5Table
    
    Call Grid_Head4
    If TmpRs5.RecordCount > 0 Then
       
        mRow = 2
        mTotMonths = DateDiff("m", mStartDate, mEndDate)
        For I = 0 To mTotMonths
            mCurrMonth = DateAdd("m", I, mStartDate)
            mYear = Year(mCurrMonth)
            mMonthNo = Month(mCurrMonth)
            mTotDays = Day(DateSerial(mYear, mMonthNo + 1, 0))
            mSStartDate = mYear & "-" & mMonthNo & "-" & "01"
            mSEndDate = mYear & "-" & mMonthNo & "-" & mTotDays
            
            TmpRs5.Filter = "ExpenseFrom <= '" & mSEndDate & "' And ExpenseTo >= '" & mSStartDate & "'"
            
            For r = 1 To TmpRs5.RecordCount
                mTotMonths = DateDiff("m", TmpRs5!ExpenseFrom, TmpRs5!ExpenseTo) + 1
                mExpAmt = TmpRs5!TotalExpenseAmount / mTotMonths
                MSHFlexGrid5.TextMatrix(mRow, 0) = TmpRs5!ClaimRowID
                MSHFlexGrid5.TextMatrix(mRow, 1) = TmpRs5!Locationname
                MSHFlexGrid5.TextMatrix(mRow, 2) = TmpRs5!Employee
                MSHFlexGrid5.TextMatrix(mRow, 3) = TmpRs5!Designation
                MSHFlexGrid5.TextMatrix(mRow, 4) = Format(TmpRs5!ExpenseFrom, "dd-MMM-yyyy")
                MSHFlexGrid5.TextMatrix(mRow, 5) = Format(TmpRs5!ExpenseTo, "dd-MMM-yyyy")
                MSHFlexGrid5.TextMatrix(mRow, 6) = TmpRs5!ExpenseType
                MSHFlexGrid5.TextMatrix(mRow, 7) = TmpRs5!TotalExpenseAmount
                MSHFlexGrid5.TextMatrix(mRow, 8) = Format(mSStartDate, "MMM-yyyy")
                MSHFlexGrid5.TextMatrix(mRow, 9) = Format(mExpAmt, "#####0.00")
                MSHFlexGrid5.TextMatrix(0, 9) = Val(MSHFlexGrid5.TextMatrix(0, 9)) + Val(MSHFlexGrid5.TextMatrix(mRow, 9))
                MSHFlexGrid5.TextMatrix(mRow, 10) = IIf(TmpRs5!Flag = "A", "Approved", "Pending")
                mRow = mRow + 1
                MSHFlexGrid5.Rows = MSHFlexGrid5.Rows + 1
                TmpRs5.MoveNext
            Next r
            TmpRs5.Filter = ""
        Next
    
    End If

Else
    tmp = "Select TC.ClaimID,TCA.LocationID,ME.EmployeeID, ME.FullName + '~' + ME.EmployeeNo  as Employee,MD.Designation,TC.ExpenseFrom,TC.ExpenseTo,MEx.ExpenseType, Sum(TCA.AllocatedAmt) As TotalExpenseAmount"
    tmp = tmp & " ,TC.Remarks,'',TC.Flag"
    tmp = tmp & " From Txn_ClaimAllocationDetail TCA"
    tmp = tmp & " inner Join Txn_ClaimSubmission TC on TCA.ClaimID=TC.ClaimID"
    tmp = tmp & " inner Join Mst_ExpenseType MEx on TCA.ExpenseTypeID=MEx.ExpenseTypeID"
    tmp = tmp & " Inner Join Mst_Employee ME On TC.EmployeeID = ME.EmployeeID"
    tmp = tmp & " Inner Join Mst_Designation MD On ME.DesignationID = MD.DesignationID"
    tmp = tmp & " Where TC.ExpenseFrom <= '" & mEndDate & "' And TC.ExpenseTo >= '" & mStartDate & "' And TC.Flag <> 'R' And TC.DepartmentID<>38 And TCA.SegmentID=1"
    If mWCLocationIDs <> "" Then
       tmp = tmp & " And TCA." & Trim(mWCLocationIDs)
    End If
    tmp = tmp & " Group By  TC.ClaimID,TCA.LocationID,ME.EmployeeID, ME.FullName, ME.EmployeeNo,MD.Designation,TC.ExpenseFrom,TC.ExpenseTo,MEx.ExpenseType,TC.Remarks,TC.Flag"
    
    
    Call TmpTableHRPro
    
    Call Grid_Head4
    If TmpRsHRPro.RecordCount > 0 Then
       
        mRow = 2
        mTotMonths = DateDiff("m", mStartDate, mEndDate)
        For I = 0 To mTotMonths
            mCurrMonth = DateAdd("m", I, mStartDate)
            mYear = Year(mCurrMonth)
            mMonthNo = Month(mCurrMonth)
            mTotDays = Day(DateSerial(mYear, mMonthNo + 1, 0))
            mSStartDate = mYear & "-" & mMonthNo & "-" & "01"
            mSEndDate = mYear & "-" & mMonthNo & "-" & mTotDays
            
            TmpRsHRPro.Filter = "ExpenseFrom <= '" & mSEndDate & "' And ExpenseTo >= '" & mSStartDate & "'"
            
            For r = 1 To TmpRsHRPro.RecordCount
                mTotMonths = DateDiff("m", TmpRsHRPro!ExpenseFrom, TmpRsHRPro!ExpenseTo) + 1
                mExpAmt = TmpRsHRPro!TotalExpenseAmount / mTotMonths
                MSHFlexGrid5.TextMatrix(mRow, 0) = TmpRsHRPro!ClaimID
                MSHFlexGrid5.TextMatrix(mRow, 1) = GetLocationName(TmpRsHRPro!LocationID)
                MSHFlexGrid5.TextMatrix(mRow, 2) = TmpRsHRPro!Employee
                MSHFlexGrid5.TextMatrix(mRow, 3) = TmpRsHRPro!Designation
                MSHFlexGrid5.TextMatrix(mRow, 4) = Format(TmpRsHRPro!ExpenseFrom, "dd-MMM-yyyy")
                MSHFlexGrid5.TextMatrix(mRow, 5) = Format(TmpRsHRPro!ExpenseTo, "dd-MMM-yyyy")
                MSHFlexGrid5.TextMatrix(mRow, 6) = TmpRsHRPro!ExpenseType
                MSHFlexGrid5.TextMatrix(mRow, 7) = TmpRsHRPro!TotalExpenseAmount
                MSHFlexGrid5.TextMatrix(mRow, 8) = Format(mSStartDate, "MMM-yyyy")
                MSHFlexGrid5.TextMatrix(mRow, 9) = Format(mExpAmt, "#####0.00")
                MSHFlexGrid5.TextMatrix(0, 9) = Val(MSHFlexGrid5.TextMatrix(0, 9)) + Val(MSHFlexGrid5.TextMatrix(mRow, 9))
                MSHFlexGrid5.TextMatrix(mRow, 10) = IIf(TmpRsHRPro!Flag = "A", "Approved", "Pending")
                mRow = mRow + 1
                MSHFlexGrid5.Rows = MSHFlexGrid5.Rows + 1
                TmpRsHRPro.MoveNext
            Next r
            TmpRsHRPro.Filter = ""
        Next
    
    End If
End If


End Sub


Private Sub FillGodownRent()
Dim mRentDays, mTotRentAmt, mRentAmt, mTotRent, mTotMonths As Double
Dim mRentStartDate, mRentEndDate As Variant
Dim mSStartDate, mSEndDate As Variant
Dim mRentStart, mRentEnd As Variant
Dim mPaidRent  As Double
Dim mRow As Integer

tmp = "Select ML.LocationID,ML.LocationName,MC.CMPName,TGL.GodownID,MG.GodownNo,TGL.LAID,TGL.LessorID,LA.LessorName,LA.NAVCode,TGL.RentAmount,TLA.StartDate,TLA.ExpiryDate,MG.CloseDate,MG.StartDate as GodownStartDate,TLA.Flag"
tmp = tmp & " From Txn_GodownLessorAgreeDetail TGL"
tmp = tmp & " inner join Txn_LessorAgreement TLA on TGL.LAID=TLA.LAID"
tmp = tmp & " Inner Join Mst_Godown MG on TGL.GodownID=MG.GodownID"
tmp = tmp & " Inner join Mst_CMP  MC on MG.CMPID=MC.CMPID"
tmp = tmp & " Inner join Mst_Location  ML on MC.LocationID=ML.LocationID"
tmp = tmp & " Inner join Mst_Lessor  LA on TGL.LessorID=LA.LessorID"
tmp = tmp & " Where TLA.StartDate <= '" & mEndDate & "' And TLA.ExpiryDate>='" & mStartDate & "' And (MG.CloseDate >= '" & mStartDate & "' or MG.CloseDate is null) And TLA.Flag<>'R'"
If mWCLocationIDs <> "" Then
   tmp = tmp & " And ML." & Trim(mWCLocationIDs)
End If

tmp = tmp & " Order By ML.LocationID"

Call Tmp5Table

Call Grid_Head2
If TmpRs5.RecordCount > 0 Then
   mTotRentAmt = 0
   mRow = 2
   TmpRs5.MoveFirst
   For r = 1 To TmpRs5.RecordCount
        mRentAmt = 0
        mTotMonths = DateDiff("m", mStartDate, mEndDate)
        For I = 0 To mTotMonths
            mCurrMonth = DateAdd("m", I, mStartDate)
            mYear = Year(mCurrMonth)
            mMonthNo = Month(mCurrMonth)
            mTotDays = Day(DateSerial(mYear, mMonthNo + 1, 0))
            mSStartDate = mYear & "-" & mMonthNo & "-" & "01"
            mSEndDate = mYear & "-" & mMonthNo & "-" & mTotDays
            mRentStartDate = IIf(TmpRs5!StartDate < TmpRs5!GodownStartDate, TmpRs5!GodownStartDate, TmpRs5!StartDate)
            
            If IsNull(TmpRs5!CloseDate) = False Then
               mRentEndDate = IIf(TmpRs5!ExpiryDate > TmpRs5!CloseDate, TmpRs5!CloseDate, TmpRs5!ExpiryDate)
            Else
               mRentEndDate = TmpRs5!ExpiryDate
            End If
            If mTotDays <> 0 Then
               
               mTotRent = TmpRs5!RentAmount / mTotDays
               
               If mRentStartDate < CDate(mSStartDate) And mRentEndDate > CDate(mSEndDate) Then
                  mRentDays = DateDiff("d", CDate(mSStartDate), CDate(mSEndDate))
                  mRentStart = mSStartDate
                  mRentEnd = mSEndDate
               
               ElseIf mRentStartDate < CDate(mSStartDate) And mRentEndDate <= CDate(mSEndDate) And mRentEndDate >= CDate(mSStartDate) Then
                  mRentDays = DateDiff("d", CDate(mSStartDate), mRentEndDate)
                  mRentStart = mSStartDate
                  mRentEnd = mRentEndDate
               ElseIf mRentStartDate < CDate(mSStartDate) And mRentEndDate <= CDate(mSEndDate) Then
                  mRentDays = 0
                  mRentStart = mSStartDate
                  mRentEnd = mSEndDate
                  
               ElseIf mRentStartDate >= CDate(mSEndDate) And mRentEndDate >= CDate(mSStartDate) Then  ' And mRentEndDate > CDate(mSEndDate) Then
                  mRentDays = 0
                  mRentStart = mSStartDate
                  mRentEnd = mSEndDate
               ElseIf mRentStartDate <= CDate(mSEndDate) And mRentEndDate > CDate(mSEndDate) Then
                  mRentDays = DateDiff("d", mRentStartDate, CDate(mSEndDate))
                  mRentStart = mRentStartDate
                  mRentEnd = mSEndDate
               Else
                  mRentDays = mTotDays 'DateDiff("d", mRentStartDate, mRentEndDate)
               End If
               If mRentDays <> 0 Then
                  mRentDays = mRentDays + 1
                  mRentAmt = mRentAmt + mTotRent * mRentDays
               End If
            End If
            mDoublePaidAmt = 0
            mPaidRent = GetPaidRent(TmpRs5!GodownID, TmpRs5!LessorID, mYear, mMonthNo, TmpRs5!LAID)
            Call FixGodownRent(TmpRs5!Locationname, Format(mTotRent * mRentDays, "#####0.00"))
            
            MSHFlexGrid3.TextMatrix(mRow, 0) = TmpRs5!LocationID
            MSHFlexGrid3.TextMatrix(mRow, 1) = TmpRs5!Locationname
            MSHFlexGrid3.TextMatrix(mRow, 2) = TmpRs5!CMPName
            MSHFlexGrid3.TextMatrix(mRow, 3) = TmpRs5!GodownNo
            MSHFlexGrid3.TextMatrix(mRow, 4) = TmpRs5!LessorName
            MSHFlexGrid3.TextMatrix(mRow, 5) = IIf(IsNull(TmpRs5!NavCode), "", TmpRs5!NavCode)
'            MSHFlexGrid3.TextMatrix(mRow, 6) = Format(mRentStartDate, "dd-MMM-yyyy")
'            MSHFlexGrid3.TextMatrix(mRow, 7) = Format(mRentEndDate, "dd-MMM-yyyy")
'            MSHFlexGrid3.TextMatrix(mRow, 8) = TmpRs5!RentAmount
'            MSHFlexGrid3.TextMatrix(mRow, 9) = Format(mRentStart, "dd-MMM-yyyy")
            MSHFlexGrid3.TextMatrix(mRow, 6) = Format(mRentEnd, "MMM-yyyy")
            MSHFlexGrid3.TextMatrix(mRow, 7) = mRentDays
            MSHFlexGrid3.TextMatrix(mRow, 8) = Format(mTotRent * mRentDays, "#####0.00")
            MSHFlexGrid3.TextMatrix(mRow, 9) = Format(mPaidRent, "#####0.00")
            MSHFlexGrid3.TextMatrix(mRow, 10) = Format(mDoublePaidAmt, "#####0.00")
            
            MSHFlexGrid3.TextMatrix(0, 8) = Val(MSHFlexGrid3.TextMatrix(0, 8)) + Val(MSHFlexGrid3.TextMatrix(mRow, 8))
            MSHFlexGrid3.TextMatrix(0, 9) = Val(MSHFlexGrid3.TextMatrix(0, 9)) + Val(MSHFlexGrid3.TextMatrix(mRow, 9))
            MSHFlexGrid3.TextMatrix(0, 10) = Val(MSHFlexGrid3.TextMatrix(0, 10)) + Val(MSHFlexGrid3.TextMatrix(mRow, 10))
            MSHFlexGrid3.TextMatrix(mRow, 11) = IIf(TmpRs5!Flag = "A", "Approved", "Pending")
            
            mRow = mRow + 1
            MSHFlexGrid3.Rows = MSHFlexGrid3.Rows + 1
            
        Next
        TmpRs5.MoveNext
    Next
End If
End Sub
Private Function GetPaidRent(mGodownID_ As Double, mLessorID_ As Double, myear_ As Variant, mMonth_ As Variant, mLAID_ As Variant) As Double
Dim Retval As Variant
Retval = 0

tmp = " Select TG.GodownID,TGr.LessorID,TLA.StartDate,TGr.ExpenseFrom,TGr.ExpenseTo,TG.Amount as RentAmount from  Txn_GodownRentDetail TG"
tmp = tmp & " Inner join Txn_GodownRent TGr on TG.ClaimID=TGr.ClaimID"
tmp = tmp & " Inner Join Mst_Godown MG On MG.GodownID=TG.GodownID"
tmp = tmp & " inner join Txn_LessorAgreement TLA on TG.LAID=TLA.LAID"
tmp = tmp & " Where  TGr.Flag <> 'R' And TG.GodownID=" & mGodownID_ & " And TGr.LessorID=" & mLessorID_ & " And Month(ExpenseFrom)=" & mMonth_ & " And year(ExpenseFrom)=" & myear_ & " And TLA.LAID=" & mLAID_

Call Tmp2Table
For I = 1 To TmpRs2.RecordCount
    If I = 1 Then
       Retval = TmpRs2!RentAmount
    Else
       mDoublePaidAmt = mDoublePaidAmt + TmpRs2!RentAmount
    End If
Next
GetPaidRent = Retval
End Function

Private Sub FixGodownRent(mLoc_ As Variant, mexp_ As Variant)
For ix = 1 To MSHFlexGrid1.Rows - 1
    If LCase(MSHFlexGrid1.TextMatrix(ix, 3)) = LCase(mLoc_) Then
       MSHFlexGrid1.TextMatrix(ix, mRentCol) = Val(MSHFlexGrid1.TextMatrix(ix, mRentCol)) + Val(mexp_)
       Exit For
    End If
Next
End Sub

Private Sub FillVendorExpense()
Dim mTotExpAmt, mExpAmt, mTotExpense As Double
Dim mSEndDate As Date
Dim mSStartDate As Date
Dim mRow As Double
''tmp = "Select TVCD.LocationID,ML.LocationName,MV.VendorName,ME.ExpenseType,TVC.ExpenseFrom,TVC.ExpenseTo,Sum(TVCD.Amount) as TotalExpenseAmount"
''tmp = tmp & " From Txn_VendorClaim TVC"
''tmp = tmp & " inner join Txn_VendorClaimDetail TVCD on TVC.ClaimID=TVCD.ClaimID"
''tmp = tmp & " inner Join Mst_ExpenseType ME on TVCD.ExpenseTypeID=ME.ExpenseTypeID"
''tmp = tmp & " Inner Join Mst_Vendor MV on TVC.VendorID=MV.VendorID"
''tmp = tmp & " inner join Mst_Location ML on TVCD.LocationID=ML.LocationID"
''tmp = tmp & " Where  TVC.ExpenseFrom <= '" & mEndDate & "' And TVC.ExpenseTo >= '" & mStartDate & "' And TVC.Flag <> 'R' And TVCD.ExpenseTypeID Not in (8,69)"
''tmp = tmp & " Group by TVCD.LocationID,ML.LocationName,MV.VendorName,ME.ExpenseType,TVC.ExpenseFrom,TVC.ExpenseTo"
''tmp = tmp & " Order By ML.LocationName"

tmp = "Select TVC.ClaimID,TVCD.LocationID,ML.LocationName,MV.VendorName,ME.ExpenseType,ME.MISExpenseGroup,TVC.ExpenseFrom,TVC.ExpenseTo,Sum(TVCD.Amount) as TotalExpenseAmount,TVC.Flag"
tmp = tmp & " From Txn_VendorClaim TVC"
tmp = tmp & " inner join Txn_VendorClaimDetail TVCD on TVC.ClaimID=TVCD.ClaimID"
tmp = tmp & " inner Join Mst_ExpenseType ME on TVCD.ExpenseTypeID=ME.ExpenseTypeID"
tmp = tmp & " Inner Join Mst_Vendor MV on TVC.VendorID=MV.VendorID"
tmp = tmp & " inner join Mst_Location ML on TVCD.LocationID=ML.LocationID"
tmp = tmp & " Where  TVC.ExpenseFrom <= '" & mEndDate & "' And TVC.ExpenseTo >= '" & mStartDate & "' And TVC.Flag <> 'R' And ME.MISFlag='Y'"
If mWCLocationIDs <> "" Then
   tmp = tmp & " And TVCD." & Trim(mWCLocationIDs)
End If
tmp = tmp & " Group by TVC.ClaimID,TVCD.LocationID,ML.LocationName,MV.VendorName,ME.ExpenseType,ME.MISExpenseGroup,TVC.ExpenseFrom,TVC.ExpenseTo,TVC.Flag"
tmp = tmp & " Order By ML.LocationName"

Call Grid_Head1
Call Tmp2Table
If TmpRs2.RecordCount > 0 Then
    mRow = 2
    mTotMonths = DateDiff("m", mStartDate, mEndDate)
    For I = 0 To mTotMonths
        mCurrMonth = DateAdd("m", I, mStartDate)
        mYear = Year(mCurrMonth)
        mMonthNo = Month(mCurrMonth)
        mTotDays = Day(DateSerial(mYear, mMonthNo + 1, 0))
        mSStartDate = mYear & "-" & mMonthNo & "-" & "01"
        mSEndDate = mYear & "-" & mMonthNo & "-" & mTotDays
        
        TmpRs2.Filter = "ExpenseFrom <= '" & mSEndDate & "' And ExpenseTo >= '" & mSStartDate & "'"
        
        For r = 1 To TmpRs2.RecordCount
            mExpAmt = 0
            mTotDays = DateDiff("d", TmpRs2!ExpenseFrom, TmpRs2!ExpenseTo)
            mTotDays = mTotDays + 1
            If mTotDays <> 0 Then
               mTotExpense = TmpRs2!TotalExpenseAmount / mTotDays
               If TmpRs2!ExpenseFrom < mSStartDate And TmpRs2!ExpenseTo > mSEndDate Then
                  mExpDays = DateDiff("d", mSStartDate, mSEndDate)
               ElseIf TmpRs2!ExpenseFrom < mSStartDate And TmpRs2!ExpenseTo <= mSEndDate Then
                  mExpDays = DateDiff("d", mSStartDate, TmpRs2!ExpenseTo)
               ElseIf TmpRs2!ExpenseFrom >= mSStartDate And TmpRs2!ExpenseTo > mSEndDate Then
                  mExpDays = DateDiff("d", TmpRs2!ExpenseFrom, mSEndDate)
               Else
                  mExpDays = DateDiff("d", TmpRs2!ExpenseFrom, TmpRs2!ExpenseTo)
               End If
               mExpDays = mExpDays + 1
               mExpAmt = mTotExpense * mExpDays
            End If
            
            Call FixVendorExpense(TmpRs2!Locationname, Format(mExpAmt, "#####0.00"), "Vendor Expenses-" & TmpRs2!MISExpenseGroup)
            
            MSHFlexGrid2.TextMatrix(mRow, 0) = TmpRs2!ClaimID
            MSHFlexGrid2.TextMatrix(mRow, 1) = TmpRs2!Locationname
            MSHFlexGrid2.TextMatrix(mRow, 2) = TmpRs2!VendorName
            MSHFlexGrid2.TextMatrix(mRow, 3) = TmpRs2!ExpenseType
            MSHFlexGrid2.TextMatrix(mRow, 4) = TmpRs2!MISExpenseGroup
            MSHFlexGrid2.TextMatrix(mRow, 5) = Format(TmpRs2!ExpenseFrom, "dd-MMM-yyyy")
            MSHFlexGrid2.TextMatrix(mRow, 6) = Format(TmpRs2!ExpenseTo, "dd-MMM-yyyy")
            MSHFlexGrid2.TextMatrix(mRow, 7) = TmpRs2!TotalExpenseAmount
            MSHFlexGrid2.TextMatrix(mRow, 8) = Format(mSEndDate, "MMM-yyyy")
            MSHFlexGrid2.TextMatrix(mRow, 9) = Format(mExpAmt, "#####0.00")
            MSHFlexGrid2.TextMatrix(mRow, 10) = IIf(TmpRs2!Flag = "A", "Approved", "Pending")
            MSHFlexGrid2.TextMatrix(0, 9) = Val(MSHFlexGrid2.TextMatrix(0, 9)) + Val(MSHFlexGrid2.TextMatrix(mRow, 9))
            MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
            mRow = mRow + 1
            TmpRs2.MoveNext
        Next r
        TmpRs2.Filter = ""
    Next
End If
mRowSecurity = mRow + 1
End Sub
Private Sub FillSecurityExpense()
Dim mTotExpAmt, mExpAmt, mTotExpense As Double
Dim mSEndDate As Date
Dim mSStartDate As Date
Dim mRow As Double

'tmp = "Select TVC.ClaimID,TVCD.LocationID,ML.LocationName,MV.VendorName,ME.ExpenseType,ME.MISExpenseGroup,TVC.ExpenseFrom,TVC.ExpenseTo,Sum(TVCD.Amount) as TotalExpenseAmount,TVC.Flag"
'tmp = tmp & " From Txn_VendorClaim TVC"
'tmp = tmp & " inner join Txn_VendorClaimDetail TVCD on TVC.ClaimID=TVCD.ClaimID"
'tmp = tmp & " inner Join Mst_ExpenseType ME on TVCD.ExpenseTypeID=ME.ExpenseTypeID"
'tmp = tmp & " Inner Join Mst_Vendor MV on TVC.VendorID=MV.VendorID"
'tmp = tmp & " inner join Mst_Location ML on TVCD.LocationID=ML.LocationID"
'tmp = tmp & " Where  TVC.ExpenseFrom <= '" & mEndDate & "' And TVC.ExpenseTo >= '" & mStartDate & "' And TVC.Flag <> 'R' And ME.MISFlag='Y'"
'If mWCLocationIDs <> "" Then
'   tmp = tmp & " And TVCD." & Trim(mWCLocationIDs)
'End If
'
'tmp = tmp & " Group by TVC.ClaimID,TVCD.LocationID,ML.LocationName,MV.VendorName,ME.ExpenseType,ME.MISExpenseGroup,TVC.ExpenseFrom,TVC.ExpenseTo,TVC.Flag"
'tmp = tmp & " Order By ML.LocationName"


mTotMonths = DateDiff("m", mStartDate, mEndDate)
mCurrMonth = DateAdd("m", I, mStartDate)
mYear = Year(mCurrMonth)
mMonthNo = Month(mCurrMonth)
mTotDays = Day(DateSerial(mYear, mMonthNo + 1, 0))
mSStartDate = mYear & "-" & mMonthNo & "-" & "01"

mCurrMonth = DateAdd("m", I, mEndDate)
mYear = Year(mCurrMonth)
mMonthNo = Month(mCurrMonth)
mTotDays = Day(DateSerial(mYear, mMonthNo + 1, 0))
mSEndDate = mYear & "-" & mMonthNo & "-" & mTotDays


tmp = " Select TSB.BillTxnID,TSBD.LocationID,ML.LocationName,MSA.AgencyName,TSB.MonthYear,Sum(BillDetailAmt) + Sum(BillDetailSTax) "
tmp = tmp & " as TotalBillAmount,TSB.Status"
tmp = tmp & " From Txn_SecurityBilling TSB"
tmp = tmp & " inner join Txn_SecurityBillingDetail TSBD on TSB.BillTxnID = TSBD.BillTxnID"
tmp = tmp & " Inner Join Mst_SecurityAgency MSA on TSB.SecurityAgencyID = MSA.SecurityAgencyID"
tmp = tmp & " inner join Mst_Location ML on TSBD.LocationID=ML.LocationID"
tmp = tmp & " Where TSB.Status <> 'R' And  TSB.MonthYear between '" & mSStartDate & "' And '" & mSEndDate & "' "
tmp = tmp & " Group by TSB.BillTxnID,TSBD.LocationID,ML.LocationName,MSA.AgencyName,TSB.MonthYear,TSB.Status"
tmp = tmp & " Order By ML.LocationName,TSB.MonthYear"

Call Grid_Head11

Call Tmp2Table

If TmpRs2.RecordCount > 0 Then
   mRow = 2
        
'        TmpRs2.Filter = "ExpenseFrom <= '" & mSEndDate & "' And ExpenseTo >= '" & mSStartDate & "'"
'        TmpRs2.Find = "MonthYear = '" & mSStartDate & "' " ' And ExpenseTo >= '" & mSStartDate & "'"
   For r = 1 To TmpRs2.RecordCount
       
       Call FixVendorExpense(TmpRs2!Locationname, Format(TmpRs2!TotalBillAmount, "#####0.00"), "Vendor Expenses-Security")
       
           
       MSHFlexGrid12.TextMatrix(mRow, 0) = TmpRs2!BillTxnID
       MSHFlexGrid12.TextMatrix(mRow, 1) = TmpRs2!Locationname
       MSHFlexGrid12.TextMatrix(mRow, 2) = TmpRs2!AgencyName
       MSHFlexGrid12.TextMatrix(mRow, 3) = "Security Expense"
       MSHFlexGrid12.TextMatrix(mRow, 4) = Format(TmpRs2!MonthYear, "MMM-yyyy")
       MSHFlexGrid12.TextMatrix(mRow, 5) = TmpRs2!TotalBillAmount
       MSHFlexGrid12.TextMatrix(mRow, 6) = IIf(TmpRs2!Status = "A", "Approved", "Pending")
       MSHFlexGrid12.TextMatrix(0, 5) = Val(MSHFlexGrid12.TextMatrix(0, 5)) + Val(MSHFlexGrid12.TextMatrix(mRow, 5))
       MSHFlexGrid12.Rows = MSHFlexGrid12.Rows + 1
       mRow = mRow + 1
       TmpRs2.MoveNext
   Next r
'        TmpRs2.Filter = ""
'    Next
End If
End Sub

Private Sub FixVendorExpense(mLoc_ As Variant, mexp_ As Variant, mMISExpenseGroup_ As Variant)
For ix = 2 To MSHFlexGrid1.Rows - 1
    If LCase(MSHFlexGrid1.TextMatrix(ix, 3)) = LCase(mLoc_) Then
       For C = mVendExpCol To mVendExpCol + mTotMISGroups
           If LCase(MSHFlexGrid1.TextMatrix(1, C)) = LCase(mMISExpenseGroup_) Then
              MSHFlexGrid1.TextMatrix(ix, C) = Val(MSHFlexGrid1.TextMatrix(ix, C)) + Format(Val(mexp_), "#####0.00")
              Exit Sub
           End If
       Next
    End If
Next
End Sub

Private Sub FillNNNCIncome()

tmp = "Select ML.LocationID,ML.LocationName,MU.UnitType,TIG.InvoiceDate,Sum(TIG.InvoiceAmount)InvoiceAmount"
tmp = tmp & " from  Txn_InvMonGSLDetail TIG"
tmp = tmp & " Inner Join Mst_Godown MG on TIG.GodownID=MG.GodownID"
tmp = tmp & " Inner Join Mst_UnitType MU on MG.UnitTypeID=MU.UnitTypeID"
tmp = tmp & " Inner Join Mst_CMP MC on TIG.CMPID=MC.CMPID"
tmp = tmp & " Inner Join MSt_Location ML on MC.LocationID=ML.LocationID"
tmp = tmp & " Where TIG.ExchangeTypeID=1 And (TIG.ContractID=0 Or TIG.ContractID is Null)  And PINo is not Null " '' And MG.UnitTypeID in (1,2)"
tmp = tmp & " And TIG.InvoiceDate >='" & mStartDate & "' And TIG.InvoiceDate <='" & mEndDate & "'"
If mWCLocationIDs <> "" Then
   tmp = tmp & " And ML." & Trim(mWCLocationIDs)
End If
tmp = tmp & " Group By ML.LocationID,ML.LocationName,MU.UnitType,TIG.InvoiceDate"
tmp = tmp & " Order By ML.LocationName,MU.UnitType,TIG.InvoiceDate"

Call Grid_Head3
Call Tmp2Table
If TmpRs2.RecordCount > 0 Then
   'MSHFlexGrid4.Rows = TmpRs2.RecordCount + 2
   For I = 1 To TmpRs2.RecordCount
       r = MSHFlexGrid4.Rows - 1
       MSHFlexGrid4.TextMatrix(r, 0) = TmpRs2!LocationID
       MSHFlexGrid4.TextMatrix(r, 1) = TmpRs2!Locationname
       MSHFlexGrid4.TextMatrix(r, 2) = TmpRs2!UnitType
       MSHFlexGrid4.TextMatrix(r, 3) = Format(TmpRs2!InvoiceDate, "MMM-yyyy")
       MSHFlexGrid4.TextMatrix(r, 4) = Format(TmpRs2!InvoiceAmount, "#####0.00")
       MSHFlexGrid4.TextMatrix(0, 4) = Val(MSHFlexGrid4.TextMatrix(0, 4)) + Val(MSHFlexGrid4.TextMatrix(r, 4))
       MSHFlexGrid4.Rows = MSHFlexGrid4.Rows + 1
       TmpRs2.MoveNext
   Next
End If
End Sub

Private Sub FillDPIncome()

tmp = "Select MDP.LocationID,ML.LocationName,MDP.MonthYear,sum(TotalAmt)As IncomeAmt from Mst_DPInvoiceUpload MDP"
tmp = tmp & " inner Join Mst_location ML on MDp.LocationID=ML.LocationID"
tmp = tmp & " Where MDP.MonthYear>='" & mStartDate & "' And MDP.MonthYear<='" & mEndDate & "'"
If mWCLocationIDs <> "" Then
   tmp = tmp & " And MDP." & Trim(mWCLocationIDs)
End If
tmp = tmp & " Group By MDP.LocationID,ML.LocationName,MDP.MonthYear"
tmp = tmp & " Order By MDP.MonthYear,ML.LocationName"

Call Grid_Head7
Call Tmp2Table
If TmpRs2.RecordCount > 0 Then
   For I = 1 To TmpRs2.RecordCount
       r = MSHFlexGrid8.Rows - 1
       MSHFlexGrid8.TextMatrix(r, 0) = TmpRs2!LocationID
       MSHFlexGrid8.TextMatrix(r, 1) = TmpRs2!Locationname
       MSHFlexGrid8.TextMatrix(r, 2) = Format(TmpRs2!MonthYear, "MMM-yyyy")
       MSHFlexGrid8.TextMatrix(r, 3) = Format(TmpRs2!IncomeAmt, "#####0.00")
       MSHFlexGrid8.TextMatrix(0, 3) = Val(MSHFlexGrid8.TextMatrix(0, 3)) + Val(MSHFlexGrid8.TextMatrix(r, 3))
       MSHFlexGrid8.Rows = MSHFlexGrid8.Rows + 1
       TmpRs2.MoveNext
   Next
End If
End Sub

Private Sub FillNotAgreeGodowns()

'Commented On 04 Jan 2010
'tmp = "Select ML.LocationName,MG.GodownID,MG.GodownNo,MG.Address From Mst_Godown MG"
'tmp = tmp & " inner join Mst_CMP MC on MC.CMPID=MG.CMPID"
'tmp = tmp & " inner join Mst_Location ML on MC.LocationID=ML.LocationID"
'tmp = tmp & " Where MG.UnitTypeID=1 And (MG.CloseDate is Null or MG.CloseDate <='" & mEndDate & "') And MG.StartDate <='" & mEndDate & "' And MG.GodownID not in (Select distinct TGL.GodownID"
'tmp = tmp & " From Txn_GodownLessorAgreeDetail TGL"
'tmp = tmp & " inner join Txn_LessorAgreement TLA on TGL.LAID=TLA.LAID"
'tmp = tmp & " Inner Join Mst_Godown MG on TGL.GodownID=MG.GodownID"
'tmp = tmp & " Where TLA.StartDate <= '" & mEndDate & "' And TLA.ExpiryDate>='2010-07-01' And (MG.CloseDate >= '" & mStartDate & "' or MG.CloseDate is null) And TLA.Flag<>'R')"


tmp = "Select ML.LocationName,MG.GodownID,MG.GodownNo,MG.Address,MG.StartDate,MG.CloseDate From Mst_Godown MG"
tmp = tmp & " inner join Mst_CMP MC on MC.CMPID=MG.CMPID"
tmp = tmp & " inner join Mst_Location ML on MC.LocationID=ML.LocationID"
tmp = tmp & " Where MG.UnitTypeID=1 And (MG.CloseDate is Null or (MG.CloseDate <='" & mEndDate & "' And MG.CloseDate >='" & mStartDate & "')) And MG.StartDate <='" & mEndDate & "' And MG.GodownID not in (Select distinct TGL.GodownID"
tmp = tmp & " From Txn_GodownLessorAgreeDetail TGL"
tmp = tmp & " inner join Txn_LessorAgreement TLA on TGL.LAID=TLA.LAID"
tmp = tmp & " Inner Join Mst_Godown MG on TGL.GodownID=MG.GodownID"
tmp = tmp & " Where TLA.StartDate <= '" & mEndDate & "' And TLA.ExpiryDate>='" & mStartDate & "' And (MG.CloseDate >= '" & mStartDate & "' or MG.CloseDate is null) And TLA.Flag<>'R')"
If mWCLocationIDs <> "" Then
   tmp = tmp & " And ML." & Trim(mWCLocationIDs)
End If

Call Tmp2Table
Call Grid_Head6
If TmpRs2.RecordCount > 0 Then
   For I = 1 To TmpRs2.RecordCount
       r = MSHFlexGrid7.Rows - 1
       MSHFlexGrid7.TextMatrix(r, 0) = TmpRs2!Locationname
       MSHFlexGrid7.TextMatrix(r, 1) = TmpRs2!GodownID
       MSHFlexGrid7.TextMatrix(r, 2) = TmpRs2!GodownNo
       MSHFlexGrid7.TextMatrix(r, 3) = TmpRs2!Address
       MSHFlexGrid7.TextMatrix(r, 4) = IIf(IsNull(TmpRs2!StartDate), "", TmpRs2!StartDate)
       MSHFlexGrid7.TextMatrix(r, 5) = IIf(IsNull(TmpRs2!CloseDate), "", TmpRs2!CloseDate)
       MSHFlexGrid7.TextMatrix(r, 4) = Format(MSHFlexGrid7.TextMatrix(r, 4), "dd-mmm-yyyy")
       MSHFlexGrid7.TextMatrix(r, 5) = Format(MSHFlexGrid7.TextMatrix(r, 5), "dd-mmm-yyyy")
       MSHFlexGrid7.Rows = MSHFlexGrid7.Rows + 1
       TmpRs2.MoveNext
   Next
End If
End Sub

Private Sub Grid_Head8()

MSHFlexGrid9.Clear
MSHFlexGrid9.Rows = 3
MSHFlexGrid9.Cols = 6
MSHFlexGrid9.FixedCols = 2
MSHFlexGrid9.FixedRows = 2
MSHFlexGrid9.TextMatrix(1, 0) = "LocationID"
MSHFlexGrid9.TextMatrix(1, 1) = "Location"
MSHFlexGrid9.TextMatrix(1, 2) = "Client Name"
MSHFlexGrid9.TextMatrix(1, 3) = "Income Type"
MSHFlexGrid9.TextMatrix(1, 4) = "For the Month"
MSHFlexGrid9.TextMatrix(1, 5) = "Amount"
MSHFlexGrid9.ColWidth(0) = 0
MSHFlexGrid9.ColWidth(1) = 2000
MSHFlexGrid9.ColWidth(2) = 2400
MSHFlexGrid9.ColWidth(3) = 2200
MSHFlexGrid9.ColWidth(4) = 1400
MSHFlexGrid9.ColWidth(5) = 1400
MSHFlexGrid9.RowHeight(1) = 900

End Sub


Private Sub FillOtherIncome()
'---For 7 Days income

tmp = "Select TI.LocationID,ML.LocationName,ClientName,MIE.IEType,MonthYear,sum(Amount)As IncomeAmt from Txn_OtherIncome TI"
tmp = tmp & " inner Join Mst_location ML on TI.LocationID=ML.LocationID"
tmp = tmp & " inner Join Mst_IncomeExpenseType MIE on MIE.IETypeID=TI.IETypeID"
tmp = tmp & " Where TI.TxnDate>='" & mStartDate & "' And TI.TxnDate<='" & mEndDate & "' And MIE.Flag='A'"
If mWCLocationIDs <> "" Then
   tmp = tmp & " And TI." & Trim(mWCLocationIDs)
End If
tmp = tmp & " Group By TI.LocationID,ML.LocationName,ClientName,MIE.IEType,MonthYear"

Call Tmp2Table

Call Grid_Head8

If TmpRs2.RecordCount > 0 Then
   
   For I = 1 To TmpRs2.RecordCount
       r = MSHFlexGrid9.Rows - 1
       MSHFlexGrid9.TextMatrix(r, 0) = TmpRs2!LocationID
       MSHFlexGrid9.TextMatrix(r, 1) = TmpRs2!Locationname
       MSHFlexGrid9.TextMatrix(r, 2) = TmpRs2!ClientName
       MSHFlexGrid9.TextMatrix(r, 3) = TmpRs2!IEType
       MSHFlexGrid9.TextMatrix(r, 4) = TmpRs2!MonthYear
       MSHFlexGrid9.TextMatrix(r, 5) = Format(TmpRs2!IncomeAmt, "#####0.00")
       MSHFlexGrid9.TextMatrix(0, 5) = Val(MSHFlexGrid9.TextMatrix(0, 5)) + Val(MSHFlexGrid9.TextMatrix(r, 5))
       MSHFlexGrid9.Rows = MSHFlexGrid9.Rows + 1
       TmpRs2.MoveNext
   Next
End If



'Reservation


tmp = " Select TI.LocationID,ML.LocationName,MC.ClientName,'Reservation Income' as 'IEType',TI.YearMonth,sum(ActualAmount)As IncomeAmt"
tmp = tmp & " from Txn_ReservationInvoice TI"
tmp = tmp & " inner Join Mst_location ML on TI.LocationID=ML.LocationID"
tmp = tmp & " inner Join Mst_StorageContract MSC On TI.ContractID = MSC.ContractID"
tmp = tmp & " inner Join Mst_Client MC On MSC.ClientID = MC.ClientID"
tmp = tmp & " Where TI.YearMonth>='" & mStartDate & "' And TI.YearMonth<='" & mEndDate & "' "
If mWCLocationIDs <> "" Then
   tmp = tmp & " And TI." & Trim(mWCLocationIDs)
End If
tmp = tmp & " Group By TI.LocationID,ML.LocationName,MC.ClientName,TI.YearMonth"


Call Tmp2Table

If TmpRs2.RecordCount > 0 Then
   
   For I = 1 To TmpRs2.RecordCount
       r = MSHFlexGrid9.Rows - 1
       MSHFlexGrid9.TextMatrix(r, 0) = TmpRs2!LocationID
       MSHFlexGrid9.TextMatrix(r, 1) = TmpRs2!Locationname
       MSHFlexGrid9.TextMatrix(r, 2) = TmpRs2!ClientName
       MSHFlexGrid9.TextMatrix(r, 3) = TmpRs2!IEType
       MSHFlexGrid9.TextMatrix(r, 4) = MonthName(Month(TmpRs2!YearMonth)) & "-" & Year(TmpRs2!YearMonth)
       MSHFlexGrid9.TextMatrix(r, 5) = Format(TmpRs2!IncomeAmt, "#####0.00")
       MSHFlexGrid9.TextMatrix(0, 5) = Val(MSHFlexGrid9.TextMatrix(0, 5)) + Val(MSHFlexGrid9.TextMatrix(r, 5))
       MSHFlexGrid9.Rows = MSHFlexGrid9.Rows + 1
       TmpRs2.MoveNext
   Next
End If

'ADHoc

tmp = " Select TI.LocationID,ML.LocationName,MC.ClientName,MIE.IEType,InvoiceDate,sum(Amount)As IncomeAmt"
tmp = tmp & " from Txn_TempInvoice TI"
tmp = tmp & " inner Join Mst_location ML on TI.LocationID=ML.LocationID"
tmp = tmp & " inner Join Mst_IncomeExpenseType MIE on MIE.IETypeID=TI.InvoiceTypeID"
tmp = tmp & " inner Join Mst_Client MC On TI.ClientID = MC.ClientID"
tmp = tmp & " Where TI.InvoiceDate>='" & mStartDate & "' And TI.InvoiceDate<='" & mEndDate & "' And MIE.Flag='A' And TI.Flag = 'A' "
If mWCLocationIDs <> "" Then
   tmp = tmp & " And TI." & Trim(mWCLocationIDs)
End If
tmp = tmp & " Group By TI.LocationID,ML.LocationName,MC.ClientName,MIE.IEType,InvoiceDate"
tmp = tmp & " Order By InvoiceDate,MIE.IEType"

Call Tmp2Table

If TmpRs2.RecordCount > 0 Then
   For I = 1 To TmpRs2.RecordCount
       r = MSHFlexGrid9.Rows - 1
       MSHFlexGrid9.TextMatrix(r, 0) = TmpRs2!LocationID
       MSHFlexGrid9.TextMatrix(r, 1) = TmpRs2!Locationname
       MSHFlexGrid9.TextMatrix(r, 2) = TmpRs2!ClientName
       MSHFlexGrid9.TextMatrix(r, 3) = TmpRs2!IEType
       MSHFlexGrid9.TextMatrix(r, 4) = Format(TmpRs2!InvoiceDate, "mmm-yyyy")
       
       MSHFlexGrid9.TextMatrix(r, 5) = Format(TmpRs2!IncomeAmt, "#####0.00")
       MSHFlexGrid9.TextMatrix(0, 5) = Val(MSHFlexGrid9.TextMatrix(0, 5)) + Val(MSHFlexGrid9.TextMatrix(r, 5))
       MSHFlexGrid9.Rows = MSHFlexGrid9.Rows + 1
       TmpRs2.MoveNext
   Next
End If

End Sub

Private Sub TxtFromDate_LostFocus()
TxtFromDate.Value = (Year(TxtFromDate.Value) & "-" & Month(TxtFromDate.Value) & "-01")
End Sub

Private Sub TxtToDate_LostFocus()
mTotalDays = Day(DateSerial(Year(TxtToDate.Value), Month(TxtToDate.Value) + 1, 0))
TxtToDate.Value = (Year(TxtToDate.Value) & "-" & Month(TxtToDate.Value) & "-" & mTotalDays)
End Sub
Private Sub CmbEmployeeID_GotFocus()
Call GetEmployeeDataonGotFocus

End Sub
Private Sub CmbEmployeeID_LostFocus()
Call GetEmployeeDataOnLostFocus
End Sub
Private Sub GetEmployeeDataonGotFocus()
tmp = CmbEmployeeID.Text
Call TableMst_Employee("Where Flag = 'Y' And DesignationID= 2")
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
Private Sub GetEmployeeDataOnLostFocus()
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
End Sub

Private Sub FillTotalSalaryLocWise()

tmp = "Select EMP.MonthEmpID,EMP.EmployeeNo,EMP.EmployeeName,MD.DesignationID,MD.Designation,EMP.Flag,Emp.EmployeeID,EMP.CTC,EMP.GrossPay,EMP.TxnDate"
tmp = tmp & "  From Txn_MonthEmp_History EMP "
tmp = tmp & "  Inner Join Mst_Designation MD On EMP.DesignationID = MD.DesignationID"
tmp = tmp & "  where EMP.TxnDate>='" & Format(mStartDate, Gen_DatFormat) & "' And EMP.TxnDate<='" & Format(mEndDate, Gen_DatFormat) & "'  And EMP.DesignationID not in (1,7)" 'And EMP.Flag='Y'"
tmp = tmp & "  order by EMP.EmployeeName"

Call TmpTable

Call Grid_Head5
mRow = 2
If TmpRs.RecordCount > 0 Then
   For ir = 1 To TmpRs.RecordCount
       If Check1.Value = 0 Then
          If Val(IIf(IsNull(TmpRs!CTC), 0, TmpRs!CTC)) > 0 Then
             Call GetEmpLocations
          End If
       Else
          If Val(IIf(IsNull(TmpRs!GrossPay), 0, TmpRs!GrossPay)) > 0 Then
             Call GetEmpLocations
          End If
       End If
       TmpRs.MoveNext
   Next
End If

End Sub
Private Sub GetEmpLocations()

mTotCapacity = 0
mTotIncome = 0

If TmpRs!DesignationID = 2 Then
   Call GetRMCapIncome
   For ix = 2 To MSHFlexGrid1.Rows - 1
        If LCase(MSHFlexGrid1.TextMatrix(ix, 1)) = LCase(TmpRs!EmployeeName & "~" & TmpRs!EmployeeNo) And LCase(MSHFlexGrid1.TextMatrix(ix, 3)) <> "total" Then
           MSHFlexGrid6.TextMatrix(mRow, 0) = Format(TmpRs!TxnDate, "MMM-yyyy")
           MSHFlexGrid6.TextMatrix(mRow, 1) = TmpRs!EmployeeName & "~" & TmpRs!EmployeeNo
           MSHFlexGrid6.TextMatrix(mRow, 2) = TmpRs!Designation
           MSHFlexGrid6.TextMatrix(mRow, 3) = MSHFlexGrid1.TextMatrix(ix, 3)
           If OptIncome.Value = True Then
              If Check1.Value = 0 Then
                 MSHFlexGrid1.TextMatrix(ix, mEmpSalCol) = Val(MSHFlexGrid1.TextMatrix(ix, mEmpSalCol)) + Format(((Val(IIf(IsNull(TmpRs!CTC), 0, TmpRs!CTC)) * Val(MSHFlexGrid1.TextMatrix(ix, 5))) / IIf(Val(mTotIncome) = 0, 1, Val(mTotIncome))), "#####0.00")
                 'MSHFlexGrid6.TextMatrix(mRow, 4) = Format(((Val(IIf(IsNull(TmpRs!CTC), 0, TmpRs!CTC)) * Val(MSHFlexGrid1.TextMatrix(ix, 5))) / IIf(Val(mTotIncome) = 0, 1, Val(mTotIncome))), "#####0.00")
              Else
                 MSHFlexGrid1.TextMatrix(ix, mEmpSalCol) = Val(MSHFlexGrid1.TextMatrix(ix, mEmpSalCol)) + Format(((Val(IIf(IsNull(TmpRs!GrossPay), 0, TmpRs!GrossPay)) * Val(MSHFlexGrid1.TextMatrix(ix, 5))) / IIf(Val(mTotIncome) = 0, 1, Val(mTotIncome))), "#####0.00")
                 'MSHFlexGrid6.TextMatrix(mRow, 4) = Format(((Val(IIf(IsNull(TmpRs!GrossPay), 0, TmpRs!GrossPay)) * Val(MSHFlexGrid1.TextMatrix(ix, 5))) / IIf(Val(mTotIncome) = 0, 1, Val(mTotIncome))), "#####0.00")
              End If
           ElseIf OptCap.Value = True Then
              If Check1.Value = 0 Then
                 MSHFlexGrid1.TextMatrix(ix, mEmpSalCol) = Val(MSHFlexGrid1.TextMatrix(ix, mEmpSalCol)) + Format(((Val(IIf(IsNull(TmpRs!CTC), 0, TmpRs!CTC)) * Val(MSHFlexGrid1.TextMatrix(ix, 4))) / IIf(Val(mTotCapacity) = 0, 1, Val(mTotCapacity))), "#####0.00")
                 'MSHFlexGrid6.TextMatrix(mRow, 4) = Format(((Val(IIf(IsNull(TmpRs!CTC), 0, TmpRs!CTC)) * Val(MSHFlexGrid1.TextMatrix(ix, 4))) / IIf(Val(mTotCapacity) = 0, 1, Val(mTotCapacity))), "#####0.00")
              Else
                 MSHFlexGrid1.TextMatrix(ix, mEmpSalCol) = Val(MSHFlexGrid1.TextMatrix(ix, mEmpSalCol)) + Format(((Val(IIf(IsNull(TmpRs!GrossPay), 0, TmpRs!GrossPay)) * Val(MSHFlexGrid1.TextMatrix(ix, 4))) / IIf(Val(mTotCapacity) = 0, 1, Val(mTotCapacity))), "#####0.00")
                 'MSHFlexGrid6.TextMatrix(mRow, 4) = Format(((Val(IIf(IsNull(TmpRs!GrossPay), 0, TmpRs!GrossPay)) * Val(MSHFlexGrid1.TextMatrix(ix, 4))) / IIf(Val(mTotCapacity) = 0, 1, Val(mTotCapacity))), "#####0.00")
              End If
           ElseIf OptBoth.Value = True Then
              If Check1.Value = 0 Then
                 MSHFlexGrid1.TextMatrix(ix, mEmpSalCol) = Val(MSHFlexGrid1.TextMatrix(ix, mEmpSalCol)) + Format(((Val(IIf(IsNull(TmpRs!CTC), 0, TmpRs!CTC) / 2) * Val(MSHFlexGrid1.TextMatrix(ix, 4))) / IIf(Val(mTotCapacity) = 0, 1, Val(mTotCapacity))), "#####0.00")
                 MSHFlexGrid1.TextMatrix(ix, mEmpSalCol) = Val(MSHFlexGrid1.TextMatrix(ix, mEmpSalCol)) + Format(((Val(IIf(IsNull(TmpRs!CTC), 0, TmpRs!CTC) / 2) * Val(MSHFlexGrid1.TextMatrix(ix, 5))) / IIf(Val(mTotIncome) = 0, 1, Val(mTotIncome))), "#####0.00")
                 'MSHFlexGrid6.TextMatrix(mRow, 4) = Format(((Val(IIf(IsNull(TmpRs!CTC), 0, TmpRs!CTC) / 2) * Val(MSHFlexGrid1.TextMatrix(ix, 4))) / IIf(Val(mTotCapacity) = 0, 1, Val(mTotCapacity))) + ((Val(IIf(IsNull(TmpRs!CTC), 0, TmpRs!CTC) / 2) * Val(MSHFlexGrid1.TextMatrix(ix, 5))) / IIf(Val(mTotIncome) = 0, 1, Val(mTotIncome))), "#####0.00")
              Else
                 MSHFlexGrid1.TextMatrix(ix, mEmpSalCol) = Val(MSHFlexGrid1.TextMatrix(ix, mEmpSalCol)) + Format(((Val(IIf(IsNull(TmpRs!GrossPay), 0, TmpRs!GrossPay) / 2) * Val(MSHFlexGrid1.TextMatrix(ix, 4))) / IIf(Val(mTotCapacity) = 0, 1, Val(mTotCapacity))), "#####0.00")
                 MSHFlexGrid1.TextMatrix(ix, mEmpSalCol) = Val(MSHFlexGrid1.TextMatrix(ix, mEmpSalCol)) + Format(((Val(IIf(IsNull(TmpRs!GrossPay), 0, TmpRs!GrossPay) / 2) * Val(MSHFlexGrid1.TextMatrix(ix, 5))) / IIf(Val(mTotIncome) = 0, 1, Val(mTotIncome))), "#####0.00")
                 'MSHFlexGrid6.TextMatrix(mRow, 4) = Format(((Val(IIf(IsNull(TmpRs!GrossPay), 0, TmpRs!GrossPay) / 2) * Val(MSHFlexGrid1.TextMatrix(ix, 4))) / IIf(Val(mTotCapacity) = 0, 1, Val(mTotCapacity))) + ((Val(IIf(IsNull(TmpRs!GrossPay), 0, TmpRs!GrossPay) / 2) * Val(MSHFlexGrid1.TextMatrix(ix, 5))) / IIf(Val(mTotIncome) = 0, 1, Val(mTotIncome))), "#####0.00")
              End If
           End If
           'MSHFlexGrid6.TextMatrix(0, 4) = Val(MSHFlexGrid6.TextMatrix(0, 4)) + Val(MSHFlexGrid6.TextMatrix(mRow, 4))
           MSHFlexGrid6.Rows = MSHFlexGrid6.Rows + 1
           mRow = mRow + 1
        ElseIf LCase(MSHFlexGrid1.TextMatrix(ix, 1)) = LCase(TmpRs!EmployeeName & "~" & TmpRs!EmployeeNo) And LCase(MSHFlexGrid1.TextMatrix(ix, 3)) = "total" Then
           Exit For
        End If
    Next
ElseIf TmpRs!DesignationID = 3 Then
   Call GetAMCapIncome
   For ix = 2 To MSHFlexGrid1.Rows - 1
        If LCase(MSHFlexGrid1.TextMatrix(ix, 2)) = LCase(TmpRs!EmployeeName & "~" & TmpRs!EmployeeNo) And LCase(MSHFlexGrid1.TextMatrix(ix, 3)) <> "total" Then
           MSHFlexGrid6.TextMatrix(mRow, 0) = Format(TmpRs!TxnDate, "MMM-yyyy")
           MSHFlexGrid6.TextMatrix(mRow, 1) = TmpRs!EmployeeName & "~" & TmpRs!EmployeeNo
           MSHFlexGrid6.TextMatrix(mRow, 2) = TmpRs!Designation
           MSHFlexGrid6.TextMatrix(mRow, 3) = MSHFlexGrid1.TextMatrix(ix, 3)
           If OptIncome.Value = True Then
              If Check1.Value = 0 Then
                 MSHFlexGrid1.TextMatrix(ix, mEmpSalCol) = Val(MSHFlexGrid1.TextMatrix(ix, mEmpSalCol)) + Format(((Val(IIf(IsNull(TmpRs!CTC), 0, TmpRs!CTC)) * Val(MSHFlexGrid1.TextMatrix(ix, 5))) / IIf(Val(mTotIncome) = 0, 1, Val(mTotIncome))), "#####0.00")
                 'MSHFlexGrid6.TextMatrix(mRow, 4) = Format(((Val(IIf(IsNull(TmpRs!CTC), 0, TmpRs!CTC)) * Val(MSHFlexGrid1.TextMatrix(ix, 5))) / IIf(Val(mTotIncome) = 0, 1, Val(mTotIncome))), "#####0.00")
              Else
                 MSHFlexGrid1.TextMatrix(ix, mEmpSalCol) = Val(MSHFlexGrid1.TextMatrix(ix, mEmpSalCol)) + Format(((Val(IIf(IsNull(TmpRs!GrossPay), 0, TmpRs!GrossPay)) * Val(MSHFlexGrid1.TextMatrix(ix, 5))) / IIf(Val(mTotIncome) = 0, 1, Val(mTotIncome))), "#####0.00")
                 'MSHFlexGrid6.TextMatrix(mRow, 4) = Format(((Val(IIf(IsNull(TmpRs!GrossPay), 0, TmpRs!GrossPay)) * Val(MSHFlexGrid1.TextMatrix(ix, 5))) / IIf(Val(mTotIncome) = 0, 1, Val(mTotIncome))), "#####0.00")
              End If
           ElseIf OptCap.Value = True Then
              If Check1.Value = 0 Then
                 MSHFlexGrid1.TextMatrix(ix, mEmpSalCol) = Val(MSHFlexGrid1.TextMatrix(ix, mEmpSalCol)) + Format(((Val(IIf(IsNull(TmpRs!CTC), 0, TmpRs!CTC)) * Val(MSHFlexGrid1.TextMatrix(ix, 4))) / IIf(Val(mTotCapacity) = 0, 1, Val(mTotCapacity))), "#####0.00")
                 'MSHFlexGrid6.TextMatrix(mRow, 4) = Format(((Val(IIf(IsNull(TmpRs!CTC), 0, TmpRs!CTC)) * Val(MSHFlexGrid1.TextMatrix(ix, 4))) / IIf(Val(mTotCapacity) = 0, 1, Val(mTotCapacity))), "#####0.00")
              Else
                 MSHFlexGrid1.TextMatrix(ix, mEmpSalCol) = Val(MSHFlexGrid1.TextMatrix(ix, mEmpSalCol)) + Format(((Val(IIf(IsNull(TmpRs!GrossPay), 0, TmpRs!GrossPay)) * Val(MSHFlexGrid1.TextMatrix(ix, 4))) / IIf(Val(mTotCapacity) = 0, 1, Val(mTotCapacity))), "#####0.00")
                ' MSHFlexGrid6.TextMatrix(mRow, 4) = Format(((Val(IIf(IsNull(TmpRs!GrossPay), 0, TmpRs!GrossPay)) * Val(MSHFlexGrid1.TextMatrix(ix, 4))) / IIf(Val(mTotCapacity) = 0, 1, Val(mTotCapacity))), "#####0.00")
              End If
           ElseIf OptBoth.Value = True Then
              If Check1.Value = 0 Then
                 MSHFlexGrid1.TextMatrix(ix, mEmpSalCol) = Val(MSHFlexGrid1.TextMatrix(ix, mEmpSalCol)) + Format(((Val(IIf(IsNull(TmpRs!CTC), 0, TmpRs!CTC) / 2) * Val(MSHFlexGrid1.TextMatrix(ix, 4))) / IIf(Val(mTotCapacity) = 0, 1, Val(mTotCapacity))), "#####0.00")
                 MSHFlexGrid1.TextMatrix(ix, mEmpSalCol) = Val(MSHFlexGrid1.TextMatrix(ix, mEmpSalCol)) + Format(((Val(IIf(IsNull(TmpRs!CTC), 0, TmpRs!CTC) / 2) * Val(MSHFlexGrid1.TextMatrix(ix, 5))) / IIf(Val(mTotIncome) = 0, 1, Val(mTotIncome))), "#####0.00")
                 'MSHFlexGrid6.TextMatrix(mRow, 4) = Format(((Val(IIf(IsNull(TmpRs!CTC), 0, TmpRs!CTC) / 2) * Val(MSHFlexGrid1.TextMatrix(ix, 4))) / IIf(Val(mTotCapacity) = 0, 1, Val(mTotCapacity))) + ((Val(IIf(IsNull(TmpRs!CTC), 0, TmpRs!CTC) / 2) * Val(MSHFlexGrid1.TextMatrix(ix, 5))) / IIf(Val(mTotIncome) = 0, 1, Val(mTotIncome))), "#####0.00")
              Else
                 MSHFlexGrid1.TextMatrix(ix, mEmpSalCol) = Val(MSHFlexGrid1.TextMatrix(ix, mEmpSalCol)) + Format(((Val(IIf(IsNull(TmpRs!GrossPay), 0, TmpRs!GrossPay) / 2) * Val(MSHFlexGrid1.TextMatrix(ix, 4))) / IIf(Val(mTotCapacity) = 0, 1, Val(mTotCapacity))), "#####0.00")
                 MSHFlexGrid1.TextMatrix(ix, mEmpSalCol) = Val(MSHFlexGrid1.TextMatrix(ix, mEmpSalCol)) + Format(((Val(IIf(IsNull(TmpRs!GrossPay), 0, TmpRs!GrossPay) / 2) * Val(MSHFlexGrid1.TextMatrix(ix, 5))) / IIf(Val(mTotIncome) = 0, 1, Val(mTotIncome))), "#####0.00")
                 'MSHFlexGrid6.TextMatrix(mRow, 4) = Format(((Val(IIf(IsNull(TmpRs!GrossPay), 0, TmpRs!GrossPay) / 2) * Val(MSHFlexGrid1.TextMatrix(ix, 4))) / IIf(Val(mTotCapacity) = 0, 1, Val(mTotCapacity))) + ((Val(IIf(IsNull(TmpRs!GrossPay), 0, TmpRs!GrossPay) / 2) * Val(MSHFlexGrid1.TextMatrix(ix, 5))) / IIf(Val(mTotIncome) = 0, 1, Val(mTotIncome))), "#####0.00")
              End If
           End If
           'MSHFlexGrid6.TextMatrix(0, 4) = Val(MSHFlexGrid6.TextMatrix(0, 4)) + Val(MSHFlexGrid6.TextMatrix(mRow, 4))
           MSHFlexGrid6.Rows = MSHFlexGrid6.Rows + 1
           mRow = mRow + 1
        ElseIf LCase(MSHFlexGrid1.TextMatrix(ix, 2)) = LCase(TmpRs!EmployeeName & "~" & TmpRs!EmployeeNo) And LCase(MSHFlexGrid1.TextMatrix(ix, 3)) = "total" Then
           Exit For
        End If
    Next
Else
     tmp = "Select distinct EGM.BaselocationID,ML.LocationName"
     tmp = tmp & " From Txn_MonthEmp_History EGM"
     tmp = tmp & " inner Join Mst_Location ML On EGM.BaseLocationID = ML.LocationID"
     tmp = tmp & " left Outer Join Txn_MonthEmp_History EMP On EGM.EmployeeID = EMP.EmployeeID And  EGM.TxnDate=EMP.TxnDate"
     tmp = tmp & " Left Outer Join Txn_MonthEmp_History EMP1 On EMP.DirectReportingTo = EMP1.EmployeeID And  EMP.TxnDate=EMP1.TxnDate"
     tmp = tmp & " Left Outer Join Txn_MonthEmp_History EMP2 On EMP1.DirectReportingTo = EMP2.EmployeeID and EMP1.TxnDate=EMP2.TxnDate"
     tmp = tmp & " Left Outer Join Txn_MonthEmp_History EMP3 On EMP2.DirectReportingTo = EMP3.EmployeeID and EMP2.TxnDate=EMP3.TxnDate"
     tmp = tmp & " Left Outer Join Txn_MonthEmp_History EMP4 On EMP3.DirectReportingTo = EMP4.EmployeeID And EMP3.TxnDate=EMP4.TxnDate"
     tmp = tmp & " Left Outer Join Txn_MonthEmp_History EMP5 On EMP.ConnectedToRAG = EMP5.EmployeeID And EMP.TxnDate=EMP5.TxnDate"
     tmp = tmp & " Where (EGM.EmployeeID =" & TmpRs!EmployeeID & " or EMP.EmployeeID=" & TmpRs!EmployeeID & " or EMP1.EmployeeID=" & TmpRs!EmployeeID & " or EMP2.EmployeeID=" & TmpRs!EmployeeID & ""
     tmp = tmp & " or EMP3.EmployeeID=" & TmpRs!EmployeeID & " or EMP4.EmployeeID=" & TmpRs!EmployeeID & " or EMP5.EmployeeID=" & TmpRs!EmployeeID & " ) and EGM.BaselocationID is not null And EGM.TxnDate='" & Format(TmpRs!TxnDate, Gen_DatFormat) & "'"
     tmp = tmp & " And (EGM.DOL >= '" & Format(TmpRs!TxnDate, Gen_DatFormat) & "' Or EGM.DOL is Null) "
     Call Tmp2Table
    
     If TmpRs2.RecordCount > 0 Then
       For I = 1 To TmpRs2.RecordCount
           Call GetCapIncome
           TmpRs2.MoveNext
       Next
    
       TmpRs2.MoveFirst
       For I = 1 To TmpRs2.RecordCount
           For ix = 2 To MSHFlexGrid1.Rows - 1
               If LCase(MSHFlexGrid1.TextMatrix(ix, 3)) = LCase(TmpRs2!Locationname) Then
                  MSHFlexGrid6.TextMatrix(mRow, 0) = Format(TmpRs!TxnDate, "MMM-yyyy")
                  MSHFlexGrid6.TextMatrix(mRow, 1) = TmpRs!EmployeeName & "~" & TmpRs!EmployeeNo
                  MSHFlexGrid6.TextMatrix(mRow, 2) = TmpRs!Designation
                  MSHFlexGrid6.TextMatrix(mRow, 3) = TmpRs2!Locationname
             
                  If OptIncome.Value = True Then
                     If Check1.Value = 0 Then
                        MSHFlexGrid1.TextMatrix(ix, mEmpSalCol) = Val(MSHFlexGrid1.TextMatrix(ix, mEmpSalCol)) + Format(((Val(IIf(IsNull(TmpRs!CTC), 0, TmpRs!CTC)) * Val(MSHFlexGrid1.TextMatrix(ix, 5))) / IIf(Val(mTotIncome) = 0, 1, Val(mTotIncome))), "#####0.00")
                        'MSHFlexGrid6.TextMatrix(mRow, 4) = Format(((Val(IIf(IsNull(TmpRs!CTC), 0, TmpRs!CTC)) * Val(MSHFlexGrid1.TextMatrix(ix, 5))) / IIf(Val(mTotIncome) = 0, 1, Val(mTotIncome))), "#####0.00")
                     Else
                        MSHFlexGrid1.TextMatrix(ix, mEmpSalCol) = Val(MSHFlexGrid1.TextMatrix(ix, mEmpSalCol)) + Format(((Val(IIf(IsNull(TmpRs!GrossPay), 0, TmpRs!GrossPay)) * Val(MSHFlexGrid1.TextMatrix(ix, 5))) / IIf(Val(mTotIncome) = 0, 1, Val(mTotIncome))), "#####0.00")
                        'MSHFlexGrid6.TextMatrix(mRow, 4) = Format(((Val(IIf(IsNull(TmpRs!GrossPay), 0, TmpRs!GrossPay)) * Val(MSHFlexGrid1.TextMatrix(ix, 5))) / IIf(Val(mTotIncome) = 0, 1, Val(mTotIncome))), "#####0.00")
                     End If
                  ElseIf OptCap.Value = True Then
                     If Check1.Value = 0 Then
                        MSHFlexGrid1.TextMatrix(ix, mEmpSalCol) = Val(MSHFlexGrid1.TextMatrix(ix, mEmpSalCol)) + Format(((Val(IIf(IsNull(TmpRs!CTC), 0, TmpRs!CTC)) * Val(MSHFlexGrid1.TextMatrix(ix, 4))) / IIf(Val(mTotCapacity) = 0, 1, Val(mTotCapacity))), "#####0.00")
                        'MSHFlexGrid6.TextMatrix(mRow, 4) = Format(((Val(IIf(IsNull(TmpRs!CTC), 0, TmpRs!CTC)) * Val(MSHFlexGrid1.TextMatrix(ix, 4))) / IIf(Val(mTotCapacity) = 0, 1, Val(mTotCapacity))), "#####0.00")
                     Else
                        MSHFlexGrid1.TextMatrix(ix, mEmpSalCol) = Val(MSHFlexGrid1.TextMatrix(ix, mEmpSalCol)) + Format(((Val(IIf(IsNull(TmpRs!GrossPay), 0, TmpRs!GrossPay)) * Val(MSHFlexGrid1.TextMatrix(ix, 4))) / IIf(Val(mTotCapacity) = 0, 1, Val(mTotCapacity))), "#####0.00")
                        'MSHFlexGrid6.TextMatrix(mRow, 4) = Format(((Val(IIf(IsNull(TmpRs!GrossPay), 0, TmpRs!GrossPay)) * Val(MSHFlexGrid1.TextMatrix(ix, 4))) / IIf(Val(mTotCapacity) = 0, 1, Val(mTotCapacity))), "#####0.00")
                     End If
                  ElseIf OptBoth.Value = True Then
                     If Check1.Value = 0 Then
                        MSHFlexGrid1.TextMatrix(ix, mEmpSalCol) = Val(MSHFlexGrid1.TextMatrix(ix, mEmpSalCol)) + Format(((Val(IIf(IsNull(TmpRs!CTC), 0, TmpRs!CTC) / 2) * Val(MSHFlexGrid1.TextMatrix(ix, 4))) / IIf(Val(mTotCapacity) = 0, 1, Val(mTotCapacity))), "#####0.00")
                        MSHFlexGrid1.TextMatrix(ix, mEmpSalCol) = Val(MSHFlexGrid1.TextMatrix(ix, mEmpSalCol)) + Format(((Val(IIf(IsNull(TmpRs!CTC), 0, TmpRs!CTC) / 2) * Val(MSHFlexGrid1.TextMatrix(ix, 5))) / IIf(Val(mTotIncome) = 0, 1, Val(mTotIncome))), "#####0.00")
                        'MSHFlexGrid6.TextMatrix(mRow, 4) = Format(((Val(IIf(IsNull(TmpRs!CTC), 0, TmpRs!CTC) / 2) * Val(MSHFlexGrid1.TextMatrix(ix, 4))) / IIf(Val(mTotCapacity) = 0, 1, Val(mTotCapacity))) + ((Val(IIf(IsNull(TmpRs!CTC), 0, TmpRs!CTC) / 2) * Val(MSHFlexGrid1.TextMatrix(ix, 5))) / IIf(Val(mTotIncome) = 0, 1, Val(mTotIncome))), "#####0.00")
                     Else
                        MSHFlexGrid1.TextMatrix(ix, mEmpSalCol) = Val(MSHFlexGrid1.TextMatrix(ix, mEmpSalCol)) + Format(((Val(IIf(IsNull(TmpRs!GrossPay), 0, TmpRs!GrossPay) / 2) * Val(MSHFlexGrid1.TextMatrix(ix, 4))) / IIf(Val(mTotCapacity) = 0, 1, Val(mTotCapacity))), "#####0.00")
                        MSHFlexGrid1.TextMatrix(ix, mEmpSalCol) = Val(MSHFlexGrid1.TextMatrix(ix, mEmpSalCol)) + Format(((Val(IIf(IsNull(TmpRs!GrossPay), 0, TmpRs!GrossPay) / 2) * Val(MSHFlexGrid1.TextMatrix(ix, 5))) / IIf(Val(mTotIncome) = 0, 1, Val(mTotIncome))), "#####0.00")
                        'MSHFlexGrid6.TextMatrix(mRow, 4) = Format(((Val(IIf(IsNull(TmpRs!GrossPay), 0, TmpRs!GrossPay) / 2) * Val(MSHFlexGrid1.TextMatrix(ix, 4))) / IIf(Val(mTotCapacity) = 0, 1, Val(mTotCapacity))) + ((Val(IIf(IsNull(TmpRs!GrossPay), 0, TmpRs!GrossPay) / 2) * Val(MSHFlexGrid1.TextMatrix(ix, 5))) / IIf(Val(mTotIncome) = 0, 1, Val(mTotIncome))), "#####0.00")
                     End If
                  End If
                  'MSHFlexGrid6.TextMatrix(0, 4) = Val(MSHFlexGrid6.TextMatrix(0, 4)) + Val(MSHFlexGrid6.TextMatrix(mRow, 4))
                  MSHFlexGrid6.Rows = MSHFlexGrid6.Rows + 1
                  mRow = mRow + 1
                  Exit For
               End If
           Next
           TmpRs2.MoveNext
       Next
    End If
End If
End Sub
Private Sub GetCapIncome()
For ix = 2 To MSHFlexGrid1.Rows - 1
    If LCase(MSHFlexGrid1.TextMatrix(ix, 3)) = LCase(TmpRs2!Locationname) Then
       mTotCapacity = mTotCapacity + Val(MSHFlexGrid1.TextMatrix(ix, 4))
       mTotIncome = mTotIncome + Val(MSHFlexGrid1.TextMatrix(ix, 5))
       Exit For
    End If
Next
End Sub
Private Sub GetRMCapIncome()
For ix = 2 To MSHFlexGrid1.Rows - 1
    If LCase(MSHFlexGrid1.TextMatrix(ix, 1)) = LCase(TmpRs!EmployeeName & "~" & TmpRs!EmployeeNo) Then
       If MSHFlexGrid1.TextMatrix(ix, 3) <> "Total" And MSHFlexGrid1.TextMatrix(ix, 2) <> "" Then
          mTotCapacity = mTotCapacity + Val(MSHFlexGrid1.TextMatrix(ix, 4))
          mTotIncome = mTotIncome + Val(MSHFlexGrid1.TextMatrix(ix, 5))
       ElseIf MSHFlexGrid1.TextMatrix(ix, 3) = "Total" And MSHFlexGrid1.TextMatrix(ix, 2) <> "" Then
          Exit For
       End If
    End If
Next
End Sub
Private Sub GetAMCapIncome()
For ix = 2 To MSHFlexGrid1.Rows - 1
    If LCase(MSHFlexGrid1.TextMatrix(ix, 2)) = LCase(TmpRs!EmployeeName & "~" & TmpRs!EmployeeNo) Then
       If MSHFlexGrid1.TextMatrix(ix, 3) <> "Total" Then
          mTotCapacity = mTotCapacity + Val(MSHFlexGrid1.TextMatrix(ix, 4))
          mTotIncome = mTotIncome + Val(MSHFlexGrid1.TextMatrix(ix, 5))
       ElseIf MSHFlexGrid1.TextMatrix(ix, 3) = "Total" Then
          Exit For
       End If
    End If
Next
End Sub

'Private Sub CalDirectIncomeLocWise(mrow_ As Variant, mEmpID_ As Variant, mEMPDesignation_ As Variant, mExpenseAmount_ As Variant)
'Dim mAmtAfterDeduction, mAmtAfterDeductionI, mCorporateExpense, mCorporateExpenseAfterDed, mCorporateExpenseAfterDedI As Variant
'Dim mEMPIDDummy_ As Double
'Dim mTotCap As Double
'Dim mTotIncome As Double
'
'
'mEMPIDDummy_ = mEmpID_
'
'If LCase(mEMPDesignation_) = "managment" Or LCase(mEMPDesignation_) = "rag" Then
'   Exit Sub
'Else
'   If LCase(mEMPDesignation_) = "qqs" Or LCase(mEMPDesignation_) = "drc" Or LCase(mEMPDesignation_) = "bdo" Then
'      tmp = "Select DirectReportingTo From Mst_Employee Where EmployeeID = " & mEMPIDDummy_ & ""
'      Call Tmp2Table
'      If TmpRs2.RecordCount <> 0 Then
'         mEMPIDDummy_ = IIf(IsNull(TmpRs2!DirectReportingTo), 0, TmpRs2!DirectReportingTo)
'      End If
'   End If
'   tmp = "Select distinct ECM.LocationID,ML.LocationName " & _
'         " From Mst_EMPCMPMapping ECM " & _
'         " Right Outer Join Mst_Employee EMP On ECM.EmployeeID = EMP.EMployeeID " & _
'         " Left Outer Join Mst_Employee EMP1 On EMP.DirectReportingTo = EMP1.EmployeeID " & _
'         " Left Outer Join Mst_Employee EMP2 On EMP1.DirectReportingTo = EMP2.EmployeeID " & _
'         " Left Outer Join Mst_Employee EMP3 On EMP2.DirectReportingTo = EMP3.EmployeeID " & _
'         " Left Outer Join Mst_Employee EMP4 On EMP3.DirectReportingTo = EMP4.EmployeeID " & _
'         " Left Outer Join Mst_Employee EMP6 On EMP.ConnectedToRAG = EMP6.EmployeeID " & _
'         " Left Outer Join Mst_Location ML On ECM.LocationID = ML.LocationID " & _
'         " Left Outer Join Mst_CMP MC On ML.LocationID = MC.LocationID " & _
'         " Inner Join Mst_Godown MG On MC.CMPID = MG.CMPID " & _
'         " Where (EMP.EmployeeID = " & mEMPIDDummy_ & " Or EMP1.EmployeeID = " & mEMPIDDummy_ & " Or EMP2.EmployeeID = " & mEMPIDDummy_ & " Or EMP3.EmployeeID = " & mEMPIDDummy_ & " Or EMP4.EmployeeID = " & mEMPIDDummy_ & " Or EMP6.EmployeeID = " & mEMPIDDummy_ & ") And " & _
'         " MG.CloseDate is Null " & _
'         " Group By ECM.LocationID,ML.LocationName " & _
'         " Order By ECM.LocationID "
'End If
'
'Call Tmp2Table
'
'mTotExpense = mExpenseAmount_ / 12
'''SSTab1.Tab = 0
'mExpDays = DateDiff("m", mStartDate, mEndDate)
'mExpDays = mExpDays + 1
'mExpenseAmount_ = Format(mTotExpense * mExpDays, "#####0.00")
'
'If TmpRs2.RecordCount > 0 Then
'   For I = 1 To TmpRs2.RecordCount
'       For ix = 1 To MSHFlexGrid1.Rows - 1
'           If MSHFlexGrid1.TextMatrix(ix, 1) = TmpRs2!LocationName Then
'               If TmpRs.RecordCount > 0 Then
'                  TmpRs.MoveFirst
'               End If
'               TmpRs.Filter = ""
'               TmpRs.Filter = "LocationName='" & TmpRs2!LocationName & "'"
'               If TmpRs.RecordCount > 0 Then
'                  If FixAMSalary("AM(Current)-" & TmpRs!CCAM, mExpenseAmount_) = False Then
'                     mTotOtherEmpSal = mTotOtherEmpSal + mExpenseAmount_
'                  End If
'                  For r = 0 To mExpDays - 1
'                      mCurrMonth = DateAdd("m", r, mStartDate)
'                      mYear = Year(mCurrMonth)
'                      mMonthNo = Month(mCurrMonth)
'                      mTotDays = Day(DateSerial(mYear, mMonthNo + 1, 0))
'                      mSStartDate = mYear & "-" & mMonthNo & "-" & "01"
'                      mRow = MSHFlexGrid6.Rows - 1
'                      MSHFlexGrid6.TextMatrix(mRow, 0) = Format(mSStartDate, "dd-MMM-yyyy")
'                      MSHFlexGrid6.TextMatrix(mRow, 1) = TmpRs1!EmployeeName
'                      MSHFlexGrid6.TextMatrix(mRow, 2) = TmpRs1!CTC_YEARLY
'                      MSHFlexGrid6.TextMatrix(mRow, 3) = mTotExpense
'                      MSHFlexGrid6.TextMatrix(mRow, 4) = TmpRs!CCAM
'
'                      MSHFlexGrid6.Rows = MSHFlexGrid6.Rows + 1
'                  Next r
'                  Exit Sub
'               Else
'                  mTotOtherEmpSal = mTotOtherEmpSal + mExpenseAmount_
'               End If
'           End If
'       Next
'       TmpRs2.MoveNext
'   Next
'
'Else
'  tmp = "Select LocationName From Mst_Employee Inner Join Mst_Location on Mst_Employee.BaseLocationID=Mst_Location.LocationID Where EmployeeID=" & mEmpID_
'
'  Call Tmp2Table
'
'  If TmpRs2.RecordCount > 0 Then
'     For I = 1 To TmpRs2.RecordCount
'         For ix = 1 To MSHFlexGrid1.Rows - 1
'             If MSHFlexGrid1.TextMatrix(ix, 1) = TmpRs2!LocationName Then
'                If TmpRs.RecordCount > 0 Then
'                   TmpRs.MoveFirst
'                End If
'                TmpRs.Filter = ""
'                TmpRs.Filter = "LocationName='" & TmpRs2!LocationName & "'"
'                If TmpRs.RecordCount > 0 Then
'                   If FixAMSalary("AM(Current)-" & TmpRs!CCAM, mExpenseAmount_) = False Then
'                      mTotOtherEmpSal = mTotOtherEmpSal + mExpenseAmount_
'                   End If
'
'                   For r = 0 To mExpDays - 1
'                       mCurrMonth = DateAdd("m", r, mStartDate)
'                       mYear = Year(mCurrMonth)
'                       mMonthNo = Month(mCurrMonth)
'                       mTotDays = Day(DateSerial(mYear, mMonthNo + 1, 0))
'                       mSStartDate = mYear & "-" & mMonthNo & "-" & "01"
'                       mRow = MSHFlexGrid6.Rows - 1
'                       MSHFlexGrid6.TextMatrix(mRow, 0) = Format(mSStartDate, "dd-MMM-yyyy")
'                       MSHFlexGrid6.TextMatrix(mRow, 1) = TmpRs1!EmployeeName
'                       MSHFlexGrid6.TextMatrix(mRow, 2) = TmpRs1!CTC_YEARLY
'                       MSHFlexGrid6.TextMatrix(mRow, 3) = mTotExpense
'                       MSHFlexGrid6.TextMatrix(mRow, 4) = IIf(IsNull(TmpRs!CCAM), "", TmpRs!CCAM)
'
'                       MSHFlexGrid6.Rows = MSHFlexGrid6.Rows + 1
'                   Next r
'                   Exit Sub
'                Else
'                   mTotOtherEmpSal = mTotOtherEmpSal + mExpenseAmount_
'                End If
'             End If
'         Next
'         TmpRs2.MoveNext
'     Next
'  Else
'    For r = 0 To mExpDays - 1
'        mCurrMonth = DateAdd("m", r, mStartDate)
'        mYear = Year(mCurrMonth)
'        mMonthNo = Month(mCurrMonth)
'        mTotDays = Day(DateSerial(mYear, mMonthNo + 1, 0))
'        mSStartDate = mYear & "-" & mMonthNo & "-" & "01"
'
'        mRow = MSHFlexGrid6.Rows - 1
'        MSHFlexGrid6.TextMatrix(mRow, 0) = Format(mSStartDate, "dd-MMM-yyyy")
'        MSHFlexGrid6.TextMatrix(mRow, 1) = TmpRs1!EmployeeName
'        MSHFlexGrid6.TextMatrix(mRow, 2) = TmpRs1!CTC_YEARLY
'        MSHFlexGrid6.TextMatrix(mRow, 3) = mTotExpense
'        MSHFlexGrid6.TextMatrix(mRow, 4) = "" 'IIf(IsNull(TmpRs!CCAM), "", TmpRs!CCAM)
'
'        MSHFlexGrid6.Rows = MSHFlexGrid6.Rows + 1
'    Next
'    mTotOtherEmpSal = mTotOtherEmpSal + mExpenseAmount_
'  End If
'
'  If mTotOtherEmpID = "" Then
'     mTotOtherEmpID = mEmpID_
'  Else
'     mTotOtherEmpID = mTotOtherEmpID & "," & mEmpID_
'  End If
'End If
'
'End Sub

Private Sub FillCOExpense()
Dim mTotMonths, mMonthsCOExpense, mTotCOExpense As Double
mTotMonths = DateDiff("m", TxtFromDate.Value, TxtToDate.Value) + 1
mMonthsCOExpense = (mTotMonths * mCOExpense) / 12
mTotCOExpense = mMonthsCOExpense + mCMGExpense + mCMGSal
For ix = 2 To MSHFlexGrid1.Rows - 1
    If LCase(MSHFlexGrid1.TextMatrix(ix, 3)) <> "total" Then
       If OptIncome.Value = True Then
          MSHFlexGrid1.TextMatrix(ix, mCOExpCol) = Format(((Val(mTotCOExpense) * Val(MSHFlexGrid1.TextMatrix(ix, 5))) / IIf(Val(mAllTotIncome) = 0, 1, Val(mAllTotIncome))), "#####0.00")
       ElseIf OptCap.Value = True Then
          MSHFlexGrid1.TextMatrix(ix, mCOExpCol) = Format(((Val(mTotCOExpense) * Val(MSHFlexGrid1.TextMatrix(ix, 4))) / IIf(Val(mAllTotCapacity) = 0, 1, Val(mAllTotCapacity))), "#####0.00")
       ElseIf OptBoth.Value = True Then
          MSHFlexGrid1.TextMatrix(ix, mCOExpCol) = Format((((Val(mTotCOExpense) / 2) * Val(MSHFlexGrid1.TextMatrix(ix, 4))) / IIf(Val(mAllTotCapacity) = 0, 1, Val(mAllTotCapacity))), "#####0.00")
          MSHFlexGrid1.TextMatrix(ix, mCOExpCol) = Format((((Val(mTotCOExpense) / 2) * Val(MSHFlexGrid1.TextMatrix(ix, 5))) / IIf(Val(mAllTotIncome) = 0, 1, Val(mAllTotIncome))), "#####0.00")
       End If
    End If
Next
End Sub

Private Function GetCMGExpense() As Double
Dim mFlag As Boolean
Dim mMonths, mExpAmt, mTotMonths As Variant
Dim mExpenseAmt As Double

If mEndDate < mNewDate Then
    tmp = "Select TC.ClaimRowID,ML.LocationName,ME.EmployeeID, ME.EmployeeName + '~' + ME.EmployeeNo  as Employee,MD.Designation,TC.ExpenseFrom,TC.ExpenseTo,MEx.ExpenseType, Sum(TCS.Amount) As TotalExpenseAmount,TC.Remarks,Me.AllocateExpense,TC.Flag"
    tmp = tmp & " From Txn_ClaimSubmissionDetail TCS"
    tmp = tmp & " inner Join Txn_ClaimSubmission TC on TCS.ClaimRowID=TC.ClaimRowID"
    tmp = tmp & " inner Join Mst_ExpenseType MEx on TCS.ExpenseTypeID=MEx.ExpenseTypeID"
    tmp = tmp & " Inner Join Mst_Employee ME On TC.EmployeeID = ME.EmployeeID"
    tmp = tmp & " Inner Join Mst_Designation MD On ME.DesignationID = MD.DesignationID"
    tmp = tmp & " Inner Join Mst_Location ML On TC.LocationID = ML.LocationID"
    tmp = tmp & " Where TC.ExpenseFrom <= '" & mEndDate & "' And TC.ExpenseTo >= '" & mStartDate & "' And TC.Flag <> 'R' And ME.DesignationID in (1)And ME.CMPAllocation=1"
    'If mWCLocationIDs <> "" Then
    '   tmp = tmp & " And TC." & Trim(mWCLocationIDs)
    'End If
    tmp = tmp & " Group By TC.ClaimRowID,ML.LocationName,ME.EmployeeID,ME.EmployeeNo,ME.EmployeeName,ME.AllocateExpense,MD.Designation,TC.ExpenseFrom,TC.ExpenseTo,MEx.ExpenseType ,TC.Remarks,TC.Flag "
    
    Call Tmp5Table
    
    Call Grid_Head9
    mExpenseAmt = 0
    If TmpRs5.RecordCount > 0 Then
       
        mRow = 2
        mTotMonths = DateDiff("m", mStartDate, mEndDate)
        For I = 0 To mTotMonths
            mCurrMonth = DateAdd("m", I, mStartDate)
            mYear = Year(mCurrMonth)
            mMonthNo = Month(mCurrMonth)
            mTotDays = Day(DateSerial(mYear, mMonthNo + 1, 0))
            mSStartDate = mYear & "-" & mMonthNo & "-" & "01"
            mSEndDate = mYear & "-" & mMonthNo & "-" & mTotDays
            
            TmpRs5.Filter = "ExpenseFrom <= '" & mSEndDate & "' And ExpenseTo >= '" & mSStartDate & "'"
            
            For r = 1 To TmpRs5.RecordCount
                mTotMonths = DateDiff("m", TmpRs5!ExpenseFrom, TmpRs5!ExpenseTo) + 1
                mExpAmt = TmpRs5!TotalExpenseAmount / mTotMonths
                MSHFlexGrid10.TextMatrix(mRow, 0) = TmpRs5!ClaimRowID
                MSHFlexGrid10.TextMatrix(mRow, 1) = TmpRs5!Locationname
                MSHFlexGrid10.TextMatrix(mRow, 2) = TmpRs5!Employee
                MSHFlexGrid10.TextMatrix(mRow, 3) = TmpRs5!Designation
                MSHFlexGrid10.TextMatrix(mRow, 4) = Format(TmpRs5!ExpenseFrom, "dd-MMM-yyyy")
                MSHFlexGrid10.TextMatrix(mRow, 5) = Format(TmpRs5!ExpenseTo, "dd-MMM-yyyy")
                MSHFlexGrid10.TextMatrix(mRow, 6) = TmpRs5!ExpenseType
                MSHFlexGrid10.TextMatrix(mRow, 7) = TmpRs5!TotalExpenseAmount
                MSHFlexGrid10.TextMatrix(mRow, 8) = Format(mSStartDate, "MMM-yyyy")
                MSHFlexGrid10.TextMatrix(mRow, 9) = Format(mExpAmt, "#####0.00")
                MSHFlexGrid10.TextMatrix(0, 9) = Val(MSHFlexGrid10.TextMatrix(0, 9)) + Val(MSHFlexGrid10.TextMatrix(mRow, 9))
                MSHFlexGrid10.TextMatrix(mRow, 10) = IIf(TmpRs5!Flag = "A", "Approved", "Pending")
                mExpenseAmt = mExpenseAmt + Val(Format(mExpAmt, "#####0.00"))
                mRow = mRow + 1
                MSHFlexGrid10.Rows = MSHFlexGrid10.Rows + 1
                TmpRs5.MoveNext
            Next r
            TmpRs5.Filter = ""
        Next
    
    End If

Else
    tmp = "Select TC.ClaimID,TCA.LocationID,ME.EmployeeID, ME.FullName + '~' + ME.EmployeeNo  as Employee,MD.Designation,TC.ExpenseFrom,TC.ExpenseTo,MEx.ExpenseType, Sum(TCA.AllocatedAmt) As TotalExpenseAmount"
    tmp = tmp & " ,TC.Remarks,TC.Flag"
    tmp = tmp & " From Txn_ClaimAllocationDetail TCA"
    tmp = tmp & " inner Join Txn_ClaimSubmission TC on TCA.ClaimID=TC.ClaimID"
    tmp = tmp & " inner Join Mst_ExpenseType MEx on TCA.ExpenseTypeID=MEx.ExpenseTypeID"
    tmp = tmp & " Inner Join Mst_Employee ME On TC.EmployeeID = ME.EmployeeID"
    tmp = tmp & " Inner Join Mst_Designation MD On ME.DesignationID = MD.DesignationID"
    tmp = tmp & " Where TC.ExpenseFrom <= '" & mEndDate & "' And TC.ExpenseTo >= '" & mStartDate & "' And TC.Flag <> 'R' And TC.DepartmentID=38 And TCA.SegmentID=1"
    tmp = tmp & " Group By  TC.ClaimID,TCA.LocationID,ME.EmployeeID, ME.FullName, ME.EmployeeNo,MD.Designation,TC.ExpenseFrom,TC.ExpenseTo,MEx.ExpenseType,TC.Remarks,TC.Flag"
    
    Call TmpTableHRPro
    
    Call Grid_Head9
    mExpenseAmt = 0
    If TmpRsHRPro.RecordCount > 0 Then
       
        mRow = 2
        mTotMonths = DateDiff("m", mStartDate, mEndDate)
        For I = 0 To mTotMonths
            mCurrMonth = DateAdd("m", I, mStartDate)
            mYear = Year(mCurrMonth)
            mMonthNo = Month(mCurrMonth)
            mTotDays = Day(DateSerial(mYear, mMonthNo + 1, 0))
            mSStartDate = mYear & "-" & mMonthNo & "-" & "01"
            mSEndDate = mYear & "-" & mMonthNo & "-" & mTotDays
            
            TmpRsHRPro.Filter = "ExpenseFrom <= '" & mSEndDate & "' And ExpenseTo >= '" & mSStartDate & "'"
            
            For r = 1 To TmpRsHRPro.RecordCount
                mTotMonths = DateDiff("m", TmpRsHRPro!ExpenseFrom, TmpRsHRPro!ExpenseTo) + 1
                mExpAmt = TmpRsHRPro!TotalExpenseAmount / mTotMonths
                MSHFlexGrid10.TextMatrix(mRow, 0) = TmpRsHRPro!ClaimID
                MSHFlexGrid10.TextMatrix(mRow, 1) = GetLocationName(TmpRsHRPro!LocationID)
                MSHFlexGrid10.TextMatrix(mRow, 2) = TmpRsHRPro!Employee
                MSHFlexGrid10.TextMatrix(mRow, 3) = TmpRsHRPro!Designation
                MSHFlexGrid10.TextMatrix(mRow, 4) = Format(TmpRsHRPro!ExpenseFrom, "dd-MMM-yyyy")
                MSHFlexGrid10.TextMatrix(mRow, 5) = Format(TmpRsHRPro!ExpenseTo, "dd-MMM-yyyy")
                MSHFlexGrid10.TextMatrix(mRow, 6) = TmpRsHRPro!ExpenseType
                MSHFlexGrid10.TextMatrix(mRow, 7) = TmpRsHRPro!TotalExpenseAmount
                MSHFlexGrid10.TextMatrix(mRow, 8) = Format(mSStartDate, "MMM-yyyy")
                MSHFlexGrid10.TextMatrix(mRow, 9) = Format(mExpAmt, "#####0.00")
                MSHFlexGrid10.TextMatrix(0, 9) = Val(MSHFlexGrid10.TextMatrix(0, 9)) + Val(MSHFlexGrid10.TextMatrix(mRow, 9))
                MSHFlexGrid10.TextMatrix(mRow, 10) = IIf(TmpRsHRPro!Flag = "A", "Approved", "Pending")
                mExpenseAmt = mExpenseAmt + Val(Format(mExpAmt, "#####0.00"))
                mRow = mRow + 1
                MSHFlexGrid10.Rows = MSHFlexGrid10.Rows + 1
                TmpRsHRPro.MoveNext
            Next r
            TmpRsHRPro.Filter = ""
        Next
    End If
End If

GetCMGExpense = mExpenseAmt
End Function

Private Function GetCMGSalary()
Dim mSalAmt As Double
'tmp = "Select sum(EMP.CTC)CTC,sum(EMP.GrossPay)GrossPay"
'tmp = tmp & "  From Txn_MonthEmp_History EMP "
'tmp = tmp & "  where EMP.TxnDate>='" & Format(mStartDate, Gen_DatFormat) & "' And EMP.TxnDate<='" & Format(mEndDate, Gen_DatFormat) & "'  And EMP.DesignationID in (1) And EMP.CMPAllocation=1"
'
'Call TmpTable

tmp = "Select EMP.MonthEmpID,EMP.EmployeeNo,EMP.EmployeeName,MD.DesignationID,MD.Designation,EMP.Flag,Emp.EmployeeID,EMP.CTC,EMP.GrossPay,EMP.TxnDate"
tmp = tmp & "  From Txn_MonthEmp_History EMP "
tmp = tmp & "  Inner Join Mst_Designation MD On EMP.DesignationID = MD.DesignationID"
tmp = tmp & "  where EMP.TxnDate>='" & Format(mStartDate, Gen_DatFormat) & "' And EMP.TxnDate<='" & Format(mEndDate, Gen_DatFormat) & "'  And EMP.DesignationID in (1) And EMP.CMPAllocation=1" 'And EMP.Flag='Y'"
tmp = tmp & "  order by EMP.EmployeeName"

Call TmpTable

Call Grid_Head10
mRow = 2
mSalAmt = 0
If TmpRs.RecordCount > 0 Then
   For I = 1 To TmpRs.RecordCount
       
      
       If Check1.Value = 0 Then
          mSalAmt = mSalAmt + IIf(IsNull(TmpRs!CTC), 0, TmpRs!CTC)
          If Val(IIf(IsNull(TmpRs!CTC), 0, TmpRs!CTC)) > 0 Then
            MSHFlexGrid11.TextMatrix(mRow, 0) = Format(TmpRs!TxnDate, "MMM-yyyy")
            MSHFlexGrid11.TextMatrix(mRow, 1) = TmpRs!EmployeeName & "~" & TmpRs!EmployeeNo
            MSHFlexGrid11.TextMatrix(mRow, 2) = TmpRs!Designation
            'MSHFlexGrid11.TextMatrix(mRow, 3) = IIf(IsNull(TmpRs!CTC), 0, TmpRs!CTC)
            'MSHFlexGrid11.TextMatrix(0, 3) = Val(MSHFlexGrid11.TextMatrix(0, 3)) + Val(MSHFlexGrid11.TextMatrix(mRow, 3))
            MSHFlexGrid11.Rows = MSHFlexGrid11.Rows + 1
            mRow = mRow + 1
         End If
       Else
          mSalAmt = mSalAmt + IIf(IsNull(TmpRs!GrossPay), 0, TmpRs!GrossPay)
          If Val(IIf(IsNull(TmpRs!GrossPay), 0, TmpRs!GrossPay)) > 0 Then
             MSHFlexGrid11.TextMatrix(mRow, 0) = Format(TmpRs!TxnDate, "MMM-yyyy")
             MSHFlexGrid11.TextMatrix(mRow, 1) = TmpRs!EmployeeName & "~" & TmpRs!EmployeeNo
             MSHFlexGrid11.TextMatrix(mRow, 2) = TmpRs!Designation
             'MSHFlexGrid11.TextMatrix(mRow, 3) = IIf(IsNull(TmpRs!GrossPay), 0, TmpRs!GrossPay)
             'MSHFlexGrid11.TextMatrix(0, 3) = Val(MSHFlexGrid11.TextMatrix(0, 3)) + Val(MSHFlexGrid11.TextMatrix(mRow, 3))
             MSHFlexGrid11.Rows = MSHFlexGrid11.Rows + 1
             mRow = mRow + 1
          End If
       End If
       
       TmpRs.MoveNext
   Next
End If
GetCMGSalary = mSalAmt
End Function

Private Sub Grid_Head9()

MSHFlexGrid10.Clear
MSHFlexGrid10.Rows = 3
MSHFlexGrid10.Cols = 11
MSHFlexGrid10.FixedCols = 2
MSHFlexGrid10.FixedRows = 2
MSHFlexGrid10.TextMatrix(1, 0) = "ClaimID"
MSHFlexGrid10.TextMatrix(1, 1) = "Location"
MSHFlexGrid10.TextMatrix(1, 2) = "EmployeeName"
MSHFlexGrid10.TextMatrix(1, 3) = "Current Designation"
MSHFlexGrid10.TextMatrix(1, 4) = "Expesne Claim From the Month"
MSHFlexGrid10.TextMatrix(1, 5) = "Expense Claim To the Month"
MSHFlexGrid10.TextMatrix(1, 6) = "Expense Type"
MSHFlexGrid10.TextMatrix(1, 7) = "Claimed Amount"
MSHFlexGrid10.TextMatrix(1, 8) = "Expense for the Month"
MSHFlexGrid10.TextMatrix(1, 9) = "Claim Amount for the Month"
MSHFlexGrid10.TextMatrix(1, 10) = "Status"

MSHFlexGrid10.ColWidth(0) = 800
MSHFlexGrid10.ColWidth(1) = 2000
MSHFlexGrid10.ColWidth(2) = 2000
MSHFlexGrid10.ColWidth(3) = 1400
MSHFlexGrid10.ColWidth(4) = 0
MSHFlexGrid10.ColWidth(5) = 0
MSHFlexGrid10.ColWidth(6) = 1400
MSHFlexGrid10.ColWidth(7) = 0
MSHFlexGrid10.ColWidth(8) = 1400
MSHFlexGrid10.ColWidth(9) = 1200
MSHFlexGrid10.ColWidth(10) = 1200

MSHFlexGrid10.RowHeight(1) = 900

End Sub

Private Sub Grid_Head10()

MSHFlexGrid11.Clear
MSHFlexGrid11.Rows = 3
MSHFlexGrid11.Cols = 4
MSHFlexGrid11.FixedCols = 2
MSHFlexGrid11.FixedRows = 2
MSHFlexGrid11.TextMatrix(1, 0) = "Month"
MSHFlexGrid11.TextMatrix(1, 1) = "Employee Name"
MSHFlexGrid11.TextMatrix(1, 2) = "Designation"
MSHFlexGrid11.TextMatrix(1, 3) = "Salary"


MSHFlexGrid11.ColWidth(0) = 1400
MSHFlexGrid11.ColWidth(1) = 2200
MSHFlexGrid11.ColWidth(2) = 1600
MSHFlexGrid11.ColWidth(3) = 1800
MSHFlexGrid11.RowHeight(1) = 900

End Sub

Private Sub AddLocations(mAMIDs_ As Variant, mRMID_ As Variant)
mLocationIDs = ""

TmpRs.Filter = "CCAMID=" & mAMIDs_ & " And AMRMID=" & mRMID_
mCCExp = 0
mBDoExp = 0
If TmpRs.RecordCount > 0 Then
   mRMFlag = True
   For j = 1 To TmpRs.RecordCount
       If CheckLocation(TmpRs!Locationname) = False Then
          mStateID = TmpRs!StateID
          mRow = MSHFlexGrid1.Rows - 1
          MSHFlexGrid1.TextMatrix(mRow, 0) = TmpRs!LocationID
          MSHFlexGrid1.TextMatrix(mRow, 3) = TmpRs!Locationname
          MSHFlexGrid1.TextMatrix(mRow, 1) = GetEmployeeNonName(mRMID_) '   TmpRs1!EmployeeName & "~" & TmpRs1!EmployeeNo
          MSHFlexGrid1.TextMatrix(mRow, 2) = GetEmployeeNonName(mAMIDs_) 'RsMst_Employee!EmployeeName & "~" & RsMst_Employee!EmployeeNo
          If mLocationIDs = "" Then
             mLocationIDs = TmpRs!LocationID
          Else
             mLocationIDs = mLocationIDs & "," & TmpRs!LocationID
          End If
        
          MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
          
      End If
      TmpRs.MoveNext
  Next

    TmpRs4.Filter = "AM_EmployeeID=" & mAMIDs_
    For k = 1 To TmpRs4.RecordCount
        If CheckLocation(TmpRs4!Locationname) = False Then
           ''mStateID = TmpRs!StateID
           mRow = MSHFlexGrid1.Rows - 1
           MSHFlexGrid1.TextMatrix(mRow, 0) = TmpRs4!LocationID
           MSHFlexGrid1.TextMatrix(mRow, 3) = TmpRs4!Locationname
           MSHFlexGrid1.TextMatrix(mRow, 1) = GetEmployeeNonName(mRMID_) 'TmpRs1!EmployeeName & "~" & TmpRs1!EmployeeNo
           MSHFlexGrid1.TextMatrix(mRow, 2) = GetEmployeeNonName(mAMIDs_) 'RsMst_Employee!EmployeeName & "~" & RsMst_Employee!EmployeeNo
           If mLocationIDs = "" Then
              mLocationIDs = TmpRs4!LocationID
           Else
              mLocationIDs = mLocationIDs & "," & TmpRs4!LocationID
           End If
           MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
        End If
    
        TmpRs4.MoveNext
    Next
    TmpRs4.Filter = ""


  mRow1 = MSHFlexGrid1.Rows - 1
'-------for BDO & CC of Locations
  If mLocationIDs <> "" Then
     If mWCLocationIDs = "LocationID in (" Then
        mWCLocationIDs = mWCLocationIDs & mLocationIDs
     Else
        mWCLocationIDs = mWCLocationIDs & "," & mLocationIDs
     End If
     
     'MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
  
     mRow1 = MSHFlexGrid1.Rows - 1
   
     MSHFlexGrid1.TextMatrix(mRow1, 1) = GetEmployeeNonName(mRMID_) 'TmpRs1!EmployeeName & "~" & TmpRs1!EmployeeNo
     MSHFlexGrid1.TextMatrix(mRow1, 2) = GetEmployeeNonName(mAMIDs_) 'RsMst_Employee!EmployeeName & "~" & RsMst_Employee!EmployeeNo
     MSHFlexGrid1.TextMatrix(mRow1, 3) = "Total"
     For C = 0 To MSHFlexGrid1.Cols - 1
         MSHFlexGrid1.row = mRow1
         MSHFlexGrid1.Col = C
         MSHFlexGrid1.CellBackColor = vbGreen
     Next
     MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
     '-------------
  End If
  'MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1

End If
TmpRs.Filter = ""

End Sub
