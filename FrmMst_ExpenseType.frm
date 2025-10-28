VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmMst_ExpenseType 
   Caption         =   "Expense Type"
   ClientHeight    =   6285
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9810
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   6285
   ScaleWidth      =   9810
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   3375
      Left            =   120
      TabIndex        =   41
      Top             =   6600
      Visible         =   0   'False
      Width           =   14415
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   350
         Left            =   9765
         TabIndex        =   45
         Top             =   195
         Width           =   1000
      End
      Begin VB.TextBox TxtSExpenseType 
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
         Left            =   1680
         TabIndex        =   44
         Top             =   240
         Width           =   3375
      End
      Begin VB.CommandButton CmdGo 
         Caption         =   "Go"
         Height          =   350
         Left            =   8760
         TabIndex        =   43
         Top             =   195
         Width           =   1000
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
         ItemData        =   "FrmMst_ExpenseType.frx":0000
         Left            =   6225
         List            =   "FrmMst_ExpenseType.frx":000A
         TabIndex        =   42
         Top             =   240
         Width           =   1575
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1740
         Left            =   120
         TabIndex        =   46
         Top             =   720
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   3069
         _Version        =   393216
         Rows            =   3
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label11 
         Caption         =   "Expense Type"
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
         Left            =   240
         TabIndex        =   50
         Top             =   270
         Width           =   1695
      End
      Begin VB.Label Label12 
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
         Height          =   255
         Left            =   5400
         TabIndex        =   49
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label13 
         Caption         =   "Total Record Found :"
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
         Left            =   10920
         TabIndex        =   48
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label LblTotalRec 
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
         Left            =   13200
         TabIndex        =   47
         Top             =   240
         Width           =   90
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5310
      Left            =   120
      TabIndex        =   24
      Top             =   120
      Width           =   8010
      Begin VB.TextBox TxtHSNNo 
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
         Left            =   5895
         MaxLength       =   25
         TabIndex        =   6
         Top             =   2400
         Width           =   1560
      End
      Begin VB.ComboBox CmbReverseFlag 
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
         ItemData        =   "FrmMst_ExpenseType.frx":0021
         Left            =   5640
         List            =   "FrmMst_ExpenseType.frx":002B
         TabIndex        =   11
         Top             =   3825
         Width           =   1845
      End
      Begin VB.ComboBox CmbOhterDedFlag 
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
         ItemData        =   "FrmMst_ExpenseType.frx":0038
         Left            =   2175
         List            =   "FrmMst_ExpenseType.frx":0042
         TabIndex        =   10
         Top             =   3840
         Width           =   1605
      End
      Begin VB.TextBox TxtMISExpGroup 
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
         Left            =   2175
         TabIndex        =   9
         Top             =   3330
         Width           =   5280
      End
      Begin VB.ComboBox CmbMISStatus 
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
         ItemData        =   "FrmMst_ExpenseType.frx":004F
         Left            =   5640
         List            =   "FrmMst_ExpenseType.frx":0059
         TabIndex        =   8
         Top             =   2835
         Width           =   1815
      End
      Begin VB.TextBox TxtMISStatus 
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
         Left            =   5640
         MaxLength       =   25
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   2835
         Width           =   1815
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
         ItemData        =   "FrmMst_ExpenseType.frx":0066
         Left            =   2175
         List            =   "FrmMst_ExpenseType.frx":0070
         TabIndex        =   7
         Top             =   2850
         Width           =   1605
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
         Left            =   2175
         Locked          =   -1  'True
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   4740
         Width           =   5295
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
         Left            =   2175
         Locked          =   -1  'True
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   4275
         Width           =   5295
      End
      Begin VB.ComboBox CmbAllocateTo 
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
         ItemData        =   "FrmMst_ExpenseType.frx":0087
         Left            =   2160
         List            =   "FrmMst_ExpenseType.frx":0094
         TabIndex        =   2
         Top             =   1095
         Width           =   2085
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
         Left            =   7560
         TabIndex        =   32
         Top             =   1950
         Width           =   375
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
         Left            =   2160
         TabIndex        =   4
         Top             =   1950
         Width           =   5295
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
         Left            =   2160
         MaxLength       =   25
         TabIndex        =   30
         Top             =   1950
         Width           =   5295
      End
      Begin VB.CommandButton CmdExpenseGroupID 
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
         Left            =   7560
         TabIndex        =   15
         Top             =   1515
         Width           =   375
      End
      Begin VB.TextBox TxtNAVCode 
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
         Left            =   2175
         MaxLength       =   25
         TabIndex        =   5
         Top             =   2385
         Width           =   2520
      End
      Begin VB.TextBox TxtExpenseTypeID 
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
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox TxtExpenseType 
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
         Left            =   2160
         TabIndex        =   1
         Top             =   660
         Width           =   5295
      End
      Begin VB.ComboBox CmbExpenseGroupID 
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
         Left            =   2160
         TabIndex        =   3
         Top             =   1515
         Width           =   5295
      End
      Begin VB.TextBox TxtExpenseGroupID 
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
         Left            =   2160
         TabIndex        =   14
         Top             =   1515
         Width           =   5295
      End
      Begin VB.TextBox TxtAllocateTo 
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
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   1095
         Width           =   2055
      End
      Begin VB.TextBox TxtStatus 
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
         Left            =   2175
         MaxLength       =   25
         TabIndex        =   39
         Top             =   2850
         Width           =   1605
      End
      Begin VB.TextBox TxtOhterDedFlag 
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
         Left            =   2175
         MaxLength       =   25
         TabIndex        =   54
         Top             =   3840
         Width           =   1605
      End
      Begin VB.TextBox TxtReverseFlag 
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
         Left            =   5640
         MaxLength       =   25
         TabIndex        =   56
         Top             =   3825
         Width           =   1845
      End
      Begin VB.Label Label17 
         Caption         =   "HSN No"
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
         Left            =   5040
         TabIndex        =   58
         Top             =   2460
         Width           =   855
      End
      Begin VB.Label Label16 
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
         Height          =   270
         Left            =   3960
         TabIndex        =   57
         Top             =   3877
         Width           =   1695
      End
      Begin VB.Label Label15 
         Caption         =   "Other Deduction Flag"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   225
         TabIndex        =   55
         Top             =   3765
         Width           =   1695
      End
      Begin VB.Label Label14 
         Caption         =   "MIS Expense Group"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   240
         TabIndex        =   53
         Top             =   3300
         Width           =   1995
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label10 
         Caption         =   "MIS Status"
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
         Left            =   3960
         TabIndex        =   51
         Top             =   2910
         Width           =   1140
      End
      Begin VB.Label Label7 
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
         Height          =   255
         Left            =   240
         TabIndex        =   40
         Top             =   2910
         Width           =   1695
      End
      Begin VB.Label Label8 
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
         Height          =   270
         Left            =   240
         TabIndex        =   38
         Top             =   4792
         Width           =   1695
      End
      Begin VB.Label Label9 
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
         Left            =   240
         TabIndex        =   37
         Top             =   4342
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   "Allocate To"
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
         Left            =   240
         TabIndex        =   33
         Top             =   1163
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "Tax Type"
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
         TabIndex        =   31
         Top             =   2003
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Expense Group"
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
         Left            =   240
         TabIndex        =   29
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label3 
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
         Height          =   240
         Left            =   240
         TabIndex        =   28
         Top             =   2452
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Expense Type ID"
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
         Left            =   240
         TabIndex        =   27
         Top             =   277
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Expense Type"
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
         Left            =   240
         TabIndex        =   26
         Top             =   720
         Width           =   1695
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1095
      Left            =   120
      TabIndex        =   17
      Top             =   5370
      Width           =   8010
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New "
         Height          =   375
         Left            =   240
         TabIndex        =   0
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   375
         Left            =   1695
         TabIndex        =   12
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   3150
         TabIndex        =   23
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   4725
         TabIndex        =   22
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "&Search"
         Height          =   375
         Left            =   6180
         TabIndex        =   16
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   615
         Width           =   1455
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   375
         Left            =   1695
         TabIndex        =   20
         Top             =   615
         Width           =   1455
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   375
         Left            =   3150
         TabIndex        =   19
         Top             =   615
         Width           =   1575
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   375
         Left            =   4725
         TabIndex        =   18
         Top             =   615
         Width           =   1455
      End
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   375
         Left            =   6180
         TabIndex        =   13
         Top             =   615
         Width           =   1455
      End
   End
End
Attribute VB_Name = "FrmMst_ExpenseType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg, mServiceTaxID, mExpenseGroupID, mAllocateTo As Variant

Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)
TxtExpenseType.Locked = False
TxtNAVCode.Locked = False
TxtHSNNo.Locked = False
TxtHSNNo = ""


TxtAllocateTo.Visible = False
CmbAllocateTo.Visible = True
CmbAllocateTo.Text = ""

TxtStatus.Visible = False
CmbStatus.Visible = True
CmbStatus.Text = ""

TxtExpenseGroupID.Locked = False
CmbExpenseGroupID.Visible = True
CmdExpenseGroupID.Visible = True

TxtServiceTaxID.Locked = False
CmbServiceTaxID.Visible = True
CmdServiceTaxID.Visible = True

TxtExpenseType = ""
TxtExpenseGroupID = ""
TxtExpenseTypeID = ""
TxtNAVCode = ""

TxtMISStatus.Visible = False
CmbMISStatus.Visible = True
CmbMISStatus.Text = ""
TxtMISExpGroup.Text = "NA"

TxtOhterDedFlag.Visible = False
CmbOhterDedFlag.Visible = True
CmbOhterDedFlag.Text = ""

TxtReverseFlag.Visible = False
CmbReverseFlag.Visible = True
CmbReverseFlag.Text = ""

TxtCreated = ""
TxtUpdated = ""
TxtServiceTaxID.Text = ""
CmbServiceTaxID.Text = ""
CmbExpenseGroupID.Text = ""

If RsMst_ExpenseType.RecordCount > 0 Then
   TxtExpenseType.SetFocus
Else
   SearchCmd.Enabled = False
End If
End Sub

Private Sub CmbAllocateTo_LostFocus()
If CmbAllocateTo.Text = "" Then
   mAllocateTo = Null
   Exit Sub
End If

If LCase(CmbAllocateTo.Text) = "employee" Then
   mAllocateTo = "E"
ElseIf LCase(CmbAllocateTo.Text) = "vendor" Then
   mAllocateTo = "V"
ElseIf LCase(CmbAllocateTo.Text) = "both" Then
   mAllocateTo = "B"
Else
   MsgBox "Invalid Selection!!!"
   CmbAllocateTo.SetFocus
   Exit Sub
End If



End Sub

Private Sub CmbMISStatus_LostFocus()
If CmbMISStatus.Text = "" Then Exit Sub

If LCase(CmbMISStatus.Text) <> "yes" And LCase(CmbMISStatus.Text) <> "no" Then
   MsgBox "Invalid Selection!!!"
   CmbMISStatus.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmbOhterDedFlag_LostFocus()

If CmbOhterDedFlag.Text = "" Then Exit Sub

If LCase(CmbOhterDedFlag.Text) <> "yes" And LCase(CmbOhterDedFlag.Text) <> "no" Then
   MsgBox "Invalid Selection!!!"
   CmbOhterDedFlag.SetFocus
   Exit Sub
End If

End Sub

Private Sub CmbServiceTaxID_GotFocus()
tmp = CmbServiceTaxID.Text
Call TableMst_ServiceTax(" Where Flag = 'A' ")
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
End Sub

Private Sub CmbExpenseGroupID_GotFocus()
tmp = CmbExpenseGroupID.Text
Call TableMst_ExpenseGroup
CmbExpenseGroupID.Clear
If RsMst_ExpenseGroup.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_ExpenseGroup.RecordCount
    CmbExpenseGroupID.AddItem RsMst_ExpenseGroup!ExpenseGroup
    RsMst_ExpenseGroup.MoveNext
Next
CmbExpenseGroupID.Text = tmp
End Sub

Private Sub CmbExpenseGroupID_LostFocus()
If CmbExpenseGroupID.Text = "" Then
   mExpenseGroupID = ""
   Exit Sub
End If
RsMst_ExpenseGroup.MoveFirst
RsMst_ExpenseGroup.Find "ExpenseGroup = '" & CmbExpenseGroupID.Text & "'"
If RsMst_ExpenseGroup.EOF Then
   MsgBox "Invalid selection "
   CmbExpenseGroupID.SetFocus
   Exit Sub
End If
mExpenseGroupID = RsMst_ExpenseGroup!ExpenseGroupID
End Sub

Private Sub CmbSStatus_LostFocus()
If CmbSStatus.Text = "" Then Exit Sub

If LCase(CmbSStatus.Text) <> "active" And LCase(CmbSStatus.Text) <> "de-active" Then
   MsgBox "Invalid Selection!!!"
   CmbSStatus.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmbStatus_LostFocus()
If CmbStatus.Text = "" Then Exit Sub

If LCase(CmbStatus.Text) <> "active" And LCase(CmbStatus.Text) <> "de-active" Then
   MsgBox "Invalid Selection!!!"
   CmbStatus.SetFocus
   Exit Sub
End If

End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Mst_ExpenseType.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Mst_ExpenseType.xls")
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
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub

Private Sub CmdExpenseGroupID_Click()
FrmMst_ExpenseGroup.Show
End Sub
Private Sub Grid_Head()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 9
MSHFlexGrid1.RowHeight(0) = 500
MSHFlexGrid1.AllowUserResizing = flexResizeBoth
MSHFlexGrid1.TextMatrix(0, 0) = "Expense Type Id"
MSHFlexGrid1.TextMatrix(0, 1) = "Expense Type"
MSHFlexGrid1.TextMatrix(0, 2) = "Allocate To"
MSHFlexGrid1.TextMatrix(0, 3) = "NAV Code"
MSHFlexGrid1.TextMatrix(0, 4) = "Expense Group"
MSHFlexGrid1.TextMatrix(0, 5) = "Service Tax Description"
MSHFlexGrid1.TextMatrix(0, 6) = "Status"
MSHFlexGrid1.TextMatrix(0, 7) = "HSN No"
MSHFlexGrid1.TextMatrix(0, 8) = "Reverse Flag"

MSHFlexGrid1.ColWidth(0) = 1400
MSHFlexGrid1.ColWidth(1) = 3800
MSHFlexGrid1.ColWidth(2) = 1400
MSHFlexGrid1.ColWidth(3) = 2200
MSHFlexGrid1.ColWidth(4) = 2000
MSHFlexGrid1.ColWidth(5) = 1800
MSHFlexGrid1.ColWidth(6) = 1200
MSHFlexGrid1.ColWidth(7) = 1200

MSHFlexGrid1.WordWrap = True

End Sub

Private Sub CmdGo_Click()
Dim wc As Variant
wc = ""
Call Grid_Head
tmp = "Select * from Mst_ExpenseType"
If TxtSExpenseType.Text <> "" Then
   wc = " Where ExpenseType Like '%" & TxtSExpenseType.Text & "%'"
End If
If CmbSStatus.Text <> "" Then
   If wc <> "" Then
      wc = wc & " And Flag='" & Left(CmbSStatus.Text, 1) & "'"
   Else
      wc = " Where Flag='" & Left(CmbSStatus.Text, 1) & "'"
   End If
End If
tmp = tmp & wc
Call TmpTable

If TmpRs.RecordCount > 0 Then
   
   MSHFlexGrid1.Rows = TmpRs.RecordCount + 2
   For I = 1 To TmpRs.RecordCount
       MSHFlexGrid1.TextMatrix(I, 0) = TmpRs!ExpenseTypeID
       MSHFlexGrid1.TextMatrix(I, 1) = TmpRs!ExpenseType
       If LCase(TmpRs!AllocateTo) = "e" Then
          MSHFlexGrid1.TextMatrix(I, 2) = "Employee"
       ElseIf LCase(TmpRs!AllocateTo) = "v" Then
          MSHFlexGrid1.TextMatrix(I, 2) = "Vendor"
       ElseIf LCase(TmpRs!AllocateTo) = "b" Then
          MSHFlexGrid1.TextMatrix(I, 2) = "Both"
       End If
       MSHFlexGrid1.TextMatrix(I, 3) = IIf(IsNull(TmpRs!AllocateTo), "", TmpRs.Fields(2))
       MSHFlexGrid1.TextMatrix(I, 7) = IIf(IsNull(TmpRs!HSNNo), "", TmpRs!HSNNo)
       MSHFlexGrid1.TextMatrix(I, 8) = IIf(IsNull(TmpRs!ReverseFlag), "", TmpRs!ReverseFlag)
       If IsNull(TmpRs!ExpenseGroupID) = False Then
          MSHFlexGrid1.TextMatrix(I, 4) = GetExpenseGroupName(TmpRs!ExpenseGroupID)
       End If
       If IsNull(TmpRs!ServiceTaxID) = False Then
          MSHFlexGrid1.TextMatrix(I, 5) = GetServiceTaxDescription(TmpRs!ServiceTaxID)
       End If
       
       If LCase(TmpRs!Flag) = "a" Then
          MSHFlexGrid1.TextMatrix(I, 6) = "Active"
       ElseIf LCase(TmpRs!Flag) = "d" Then
          MSHFlexGrid1.TextMatrix(I, 6) = "De-Active"
       End If
              
       TmpRs.MoveNext
   Next
End If
LblTotalRec.Refresh
LblTotalRec.Caption = TmpRs.RecordCount


End Sub

Private Sub CmdServiceTaxID_Click()
FrmMst_ServiceTax.Show
End Sub

Private Sub DeleteCmd_Click()

Dim ans As Variant
On Error GoTo msgError
ans = MsgBox("Are you Sure you want to Delete this Record ", vbYesNo)
If ans = vbYes Then
   'tmp = "Select * from Mst_Employee Where (ExpenseTypeID = " & TxtExpenseTypeID & ")"
   'Call Tmp1Table
   'If TmpRs1.RecordCount > 0 Then
   '   MsgBox "Child Record Present!!!"
   '   Exit Sub
   'End If
   RsMst_ExpenseType.Delete
   RsMst_ExpenseType.Update
   'MsgBox "Record Successfully Removed"
   If RsMst_ExpenseType.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_ExpenseType.MoveNext
   If RsMst_ExpenseType.EOF() Then
      RsMst_ExpenseType.MoveLast
   End If
   Call Indata
   'TmpRs1.Close
End If

Exit Sub
msgError:
MsgBox "Child Record Present "
RsMst_ExpenseType.CancelUpdate


End Sub

Private Sub EditCmd_Click()
Edit_Flg = "E"
Call GButtonLock(Me, False)
TxtExpenseType.Locked = False
TxtNAVCode.Locked = False
TxtAllocateTo.Visible = False
CmbAllocateTo.Visible = True
CmbAllocateTo.Text = TxtAllocateTo.Text
TxtHSNNo.Locked = False
TxtExpenseGroupID.Locked = False
CmbExpenseGroupID.Visible = True
CmdExpenseGroupID.Visible = True
TxtServiceTaxID.Locked = False
CmbServiceTaxID.Visible = True
CmdServiceTaxID.Visible = True
TxtServiceTaxID.Visible = False
TxtExpenseGroupID.Visible = False
CmbExpenseGroupID.Text = TxtExpenseGroupID.Text
CmbServiceTaxID.Text = TxtServiceTaxID.Text

TxtOhterDedFlag.Visible = False
CmbOhterDedFlag.Visible = True
CmbOhterDedFlag.Text = TxtOhterDedFlag

TxtReverseFlag.Visible = False
CmbReverseFlag.Visible = True
CmbReverseFlag.Text = TxtReverseFlag
TxtStatus.Visible = False
CmbStatus.Visible = True
CmbStatus.Text = TxtStatus.Text
CmbMISStatus.Visible = True
CmbMISStatus.Text = TxtMISStatus.Text
End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub FirstCmd_Click()
RsMst_ExpenseType.MoveFirst
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
   MsgBox "Access Denied !!!!!!!!!"
   Exit Sub
End If
Call Grid_Head
Call TableMst_ExpenseType(" Order by ExpenseType ")

If RsMst_ExpenseType.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
Call Indata
Call GButtonLock(Me, True)


End Sub

Public Sub Indata()
'mServiceTaxID, mExpenseGroupID, mAllocateTo
TxtExpenseType.Locked = True
TxtNAVCode.Locked = True
TxtHSNNo.Locked = True

TxtAllocateTo.Visible = True
CmbAllocateTo.Visible = False

TxtStatus.Locked = True
TxtStatus.Visible = True
CmbStatus.Visible = False



TxtExpenseGroupID.Locked = True
CmdExpenseGroupID.Visible = False
CmbExpenseGroupID.Visible = False
TxtExpenseGroupID.Visible = True
TxtServiceTaxID.Locked = True
CmdServiceTaxID.Visible = False
CmbServiceTaxID.Visible = False
TxtServiceTaxID.Visible = True



TxtExpenseTypeID = RsMst_ExpenseType!ExpenseTypeID
TxtExpenseType = RsMst_ExpenseType!ExpenseType
mAllocateTo = RsMst_ExpenseType!AllocateTo
If LCase(RsMst_ExpenseType!AllocateTo) = "e" Then
   TxtAllocateTo = "Employee"
ElseIf LCase(RsMst_ExpenseType!AllocateTo) = "v" Then
   TxtAllocateTo = "Vendor"
ElseIf LCase(RsMst_ExpenseType!AllocateTo) = "b" Then
   TxtAllocateTo = "Both"
End If

mServiceTaxID = IIf(IsNull(RsMst_ExpenseType!ServiceTaxID), 0, RsMst_ExpenseType!ServiceTaxID)
If mServiceTaxID <> 0 Then
   TxtServiceTaxID.Text = GetServiceTaxDescription(RsMst_ExpenseType!ServiceTaxID)
Else
   TxtServiceTaxID.Text = ""
End If
mExpenseGroupID = IIf(IsNull(RsMst_ExpenseType!ExpenseGroupID), "", RsMst_ExpenseType!ExpenseGroupID) '
If IsNull(RsMst_ExpenseType!ExpenseGroupID) = False Then
   TxtExpenseGroupID = GetExpenseGroupName(RsMst_ExpenseType!ExpenseGroupID)
Else
   TxtExpenseGroupID.Text = ""
End If

TxtHSNNo = IIf(IsNull(RsMst_ExpenseType!HSNNo), "", RsMst_ExpenseType!HSNNo)

If LCase(RsMst_ExpenseType!Flag) = "a" Then
   TxtStatus.Text = "Active"
ElseIf LCase(RsMst_ExpenseType!Flag) = "d" Then
   TxtStatus.Text = "De-Active"
Else
   TxtStatus.Text = ""
End If


TxtNAVCode = IIf(IsNull(RsMst_ExpenseType!NAVCode), "", RsMst_ExpenseType!NAVCode)

TxtMISStatus.Locked = True
TxtMISStatus.Visible = True
CmbMISStatus.Visible = False

If LCase(RsMst_ExpenseType!MISFlag) = "y" Then
   TxtMISStatus.Text = "Yes"
ElseIf LCase(RsMst_ExpenseType!MISFlag) = "n" Then
   TxtMISStatus.Text = "No"
Else
   TxtMISStatus.Text = ""
End If
TxtMISExpGroup.Text = IIf(IsNull(RsMst_ExpenseType!MISExpenseGroup), "NA", RsMst_ExpenseType!MISExpenseGroup)

TxtOhterDedFlag.Visible = True
CmbOhterDedFlag.Visible = False
mOtherDedFlag = IIf(IsNull(RsMst_ExpenseType!OtherDeductionFlag), "N", RsMst_ExpenseType!OtherDeductionFlag)

If LCase(mOtherDedFlag) = "y" Then
   TxtOhterDedFlag.Text = "Yes"
ElseIf LCase(mOtherDedFlag) = "n" Then
   TxtOhterDedFlag.Text = "No"
End If

TxtReverseFlag.Visible = True
CmbReverseFlag.Visible = False
mReverseFlag = IIf(IsNull(RsMst_ExpenseType!ReverseFlag), "", RsMst_ExpenseType!ReverseFlag)

If LCase(mReverseFlag) = "y" Then
   TxtReverseFlag.Text = "Yes"
ElseIf LCase(mReverseFlag) = "n" Then
   TxtReverseFlag.Text = "No"
Else
   TxtReverseFlag.Text = ""
End If

'TxtOhterDedFlag = IIf(IsNull(RsMst_ExpenseType!NAVCode), "", RsMst_ExpenseType!NAVCode)

TxtCreated = IIf(IsNull(RsMst_ExpenseType!CreatedBy), "", RsMst_ExpenseType!CreatedBy) & " :- " & IIf(IsNull(RsMst_ExpenseType!CreatedDate), "", RsMst_ExpenseType!CreatedDate)
TxtUpdated = IIf(IsNull(RsMst_ExpenseType!UpdatedBy), "", RsMst_ExpenseType!UpdatedBy) & " :- " & IIf(IsNull(RsMst_ExpenseType!UpdatedDate), "", RsMst_ExpenseType!UpdatedDate)
Call GButtonLock(Me, True)
End Sub

Private Sub LastCmd_Click()
RsMst_ExpenseType.MoveLast
Call Indata
End Sub

Private Sub MSHFlexGrid1_Click()
Dim x As Variant
Dim mCaId As Variant

x = MSHFlexGrid1.RowSel
mCaId = MSHFlexGrid1.TextMatrix(x, 0)
RsMst_ExpenseType.MoveFirst

If RsMst_ExpenseType.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
RsMst_ExpenseType.MoveFirst
RsMst_ExpenseType.Find "ExpenseTypeID = " & Val(mCaId) & ""
If RsMst_ExpenseType.EOF Then
   RsMst_ExpenseType.MoveLast
End If
Frame0.Visible = False
Call Indata

End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_ExpenseType.MoveNext
If RsMst_ExpenseType.EOF Then
   RsMst_ExpenseType.MoveLast
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
RsMst_ExpenseType.MovePrevious
If RsMst_ExpenseType.BOF Then
   RsMst_ExpenseType.MoveFirst
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
If SaveCmd.Enabled = False Then
   Exit Sub
End If

If TxtExpenseType = "" Then
   MsgBox "Blank Expense Type  not allowed !!!"
   TxtExpenseType.SetFocus
   Exit Sub
End If

If CmbAllocateTo.Text = "" Then
   MsgBox "Blank Allocate to  not allowed !!!"
   CmbAllocateTo.SetFocus
   Exit Sub
End If

If CmbStatus.Text = "" Then
   MsgBox "Blank Status not allowed !!!"
   CmbStatus.SetFocus
   Exit Sub
End If

'If CmbExpenseGroupID.Text = "" Then
'   MsgBox "Blank Expense Group  not allowed !!!"
'   CmbExpenseGroupID.SetFocus
'   Exit Sub
'End If

If CmbServiceTaxID.Text = "" Then
   MsgBox "Blank Tax not allowed !!!"
   CmbServiceTaxID.SetFocus
   Exit Sub
End If

If CmbMISStatus.Text = "" Then
   MsgBox "Blank MIS Status not allowed !!!"
   CmbMISStatus.SetFocus
   Exit Sub
End If

If CmbOhterDedFlag.Text = "" Then
   MsgBox "Blank Other Deduction Flag not allowed !!!"
   CmbOhterDedFlag.SetFocus
   Exit Sub
End If
If LCase(CmbOhterDedFlag.Text) = "yes" Then
   If CmbReverseFlag.Text = "" Then
      MsgBox "Blank Reverse Flag not allowed !!!"
      CmbReverseFlag.SetFocus
   Exit Sub
   End If
End If

If Edit_Flg = "A" Then
   If RsMst_ExpenseType.RecordCount = 0 Then
      tmp = 1
   Else
        RsMst_ExpenseType.MoveFirst
        RsMst_ExpenseType.Find "ExpenseType = '" & TxtExpenseType & "'"
        If Not RsMst_ExpenseType.EOF Then
           MsgBox "Duplicate ExpenseType not Allowed"
           Exit Sub
        End If
        RsMst_ExpenseType.MoveLast
        
        tmp = "select max(ExpenseTypeID) from Mst_ExpenseType"
        Call TmpTable
        tmp = IIf(IsNull(TmpRs.Fields(0)), 1, TmpRs.Fields(0) + 1)
        
    End If
    RsMst_ExpenseType.AddNew
    RsMst_ExpenseType!ExpenseTypeID = tmp
    TxtExpenseTypeID = tmp
Else
    RsMst_ExpenseType.MoveFirst
    RsMst_ExpenseType.Find "ExpenseType = '" & TxtExpenseType & "'"
    If Not RsMst_ExpenseType.EOF Then
       If RsMst_ExpenseType!ExpenseTypeID <> TxtExpenseTypeID Then
          MsgBox "Duplicate ExpenseType not Allowed"
          Exit Sub
       End If
    End If
   RsMst_ExpenseType.MoveFirst
   RsMst_ExpenseType.Find "ExpenseTypeID = " & TxtExpenseTypeID & ""
End If
RsMst_ExpenseType!ExpenseType = TxtExpenseType
If mExpenseGroupID <> "" Then
   RsMst_ExpenseType!ExpenseGroupID = mExpenseGroupID
Else
   RsMst_ExpenseType!ExpenseGroupID = Null
End If
If mServiceTaxID <> 0 Then
   RsMst_ExpenseType!ServiceTaxID = mServiceTaxID
End If
RsMst_ExpenseType!NAVCode = TxtNAVCode
RsMst_ExpenseType!HSNNo = Trim(TxtHSNNo)
RsMst_ExpenseType!AllocateTo = mAllocateTo
 'mAllocateTo

If LCase(CmbStatus.Text) = "active" Then
   RsMst_ExpenseType!Flag = "A"
ElseIf LCase(CmbStatus.Text) = "de-active" Then
   RsMst_ExpenseType!Flag = "D"
End If

If LCase(CmbMISStatus.Text) = "yes" Then
   RsMst_ExpenseType!MISFlag = "Y"
ElseIf LCase(CmbMISStatus.Text) = "no" Then
   RsMst_ExpenseType!MISFlag = "N"
End If

RsMst_ExpenseType!MISExpenseGroup = Trim(TxtMISExpGroup.Text)

If LCase(CmbOhterDedFlag.Text) = "yes" Then
   RsMst_ExpenseType!OtherDeductionFlag = "Y"
   RsMst_ExpenseType!ReverseFlag = Left(CmbReverseFlag.Text, 1)
ElseIf LCase(CmbOhterDedFlag.Text) = "no" Then
   RsMst_ExpenseType!OtherDeductionFlag = "N"
   RsMst_ExpenseType!ReverseFlag = Null
End If

'RsMst_ExpenseType!OtherDeductionFlag = Left(CmbOhterDedFlag.Text, 1)

If IsNull(RsMst_ExpenseType!CreatedBy) = True Or RsMst_ExpenseType!CreatedBy = "" Then
   RsMst_ExpenseType!CreatedBy = Gen_UserLoginID
   RsMst_ExpenseType!CreatedDate = Now
Else
   RsMst_ExpenseType!UpdatedBy = Gen_UserLoginID
   RsMst_ExpenseType!UpdatedDate = Now 'Format(Date, "yyyy-mm-dd hh:mm:ss")
End If
RsMst_ExpenseType.Update
RsMst_ExpenseType.Requery
RsMst_ExpenseType.MoveFirst
RsMst_ExpenseType.Find "ExpenseTypeID = " & TxtExpenseTypeID & ""

Call GButtonLock(Me, True)
TxtExpenseType.Locked = True
TxtNAVCode.Locked = True
TxtExpenseGroupID.Locked = True
Call Indata
End Sub

Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   If RsMst_ExpenseType.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   If RsMst_ExpenseType.EOF Then
      RsMst_ExpenseType.MoveFirst
   End If
   Call Indata
   Exit Sub
End If


Frame0.Height = Me.Height - 700
Frame0.Width = Me.Width - 500
Frame0.Left = 120
Frame0.Top = 120

Frame0.Visible = True

MSHFlexGrid1.Width = Frame0.Width - 500
MSHFlexGrid1.Height = Frame0.Height - 700

RsMst_ExpenseType.MoveFirst

''===========comment on 7 july 2012 for not clearing search grid=============
'Call Grid_Head
''===========comment on 7 july 2012 for not clearing search grid=============

End Sub

Private Sub TxtExpenseType_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "TxtExpenseType")
End If
End Sub

Private Sub TxtExpenseType_LostFocus()
Dim ans As Boolean
ans = IsClean(TxtExpenseType)
If ans = False Then
   MsgBox "Special character not Allowed  !!!!!!!! "
   TxtExpenseType.SetFocus
End If
End Sub


Private Sub TxtExpenseTypeID_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "TxtExpenseTypeID")
End If
End Sub
Private Sub TxtNAVCode_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "TxtNAVCode")
End If
End Sub

Private Sub CmbReverseFlag_LostFocus()

If CmbReverseFlag.Text = "" Then Exit Sub

If LCase(CmbReverseFlag.Text) <> "yes" And LCase(CmbReverseFlag.Text) <> "no" Then
   MsgBox "Invalid Selection!!!"
   CmbReverseFlag.SetFocus
   Exit Sub
End If

End Sub

