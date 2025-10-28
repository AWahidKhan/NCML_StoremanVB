VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmTxn_FranchiseeRent 
   Caption         =   "Franchisee Fee Process"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      Height          =   1635
      Left            =   120
      TabIndex        =   29
      Top             =   1560
      Width           =   15015
      Begin VB.TextBox TxtGodownNo 
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
         Left            =   5445
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   465
         Width           =   1725
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
         Height          =   360
         Left            =   7230
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   465
         Width           =   4590
      End
      Begin VB.TextBox TxtPresentRate 
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
         Left            =   5430
         Locked          =   -1  'True
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   1185
         Width           =   1575
      End
      Begin VB.TextBox TxtCMPName 
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
         Left            =   2835
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   465
         Width           =   2565
      End
      Begin VB.TextBox TxtLocationName 
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
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   465
         Width           =   2670
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
         Left            =   11850
         Locked          =   -1  'True
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   465
         Width           =   3090
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
         Left            =   8880
         TabIndex        =   33
         Top             =   1155
         Width           =   1545
      End
      Begin VB.TextBox TxtNewRate 
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
         Left            =   7050
         TabIndex        =   32
         Top             =   1185
         Width           =   1575
      End
      Begin VB.ComboBox CmbBillingUnit 
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
         ItemData        =   "FrmTxn_FranchiseeRent.frx":0000
         Left            =   2835
         List            =   "FrmTxn_FranchiseeRent.frx":0010
         Sorted          =   -1  'True
         TabIndex        =   31
         Top             =   1185
         Width           =   2565
      End
      Begin VB.ComboBox CmbBillingCycle 
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
         TabIndex        =   30
         Top             =   1185
         Width           =   2670
      End
      Begin VB.Label Label10 
         Caption         =   "Godown No"
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
         Left            =   5760
         TabIndex        =   50
         Top             =   180
         Width           =   1080
      End
      Begin VB.Label Label20 
         Caption         =   "Present Rate"
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
         Left            =   5617
         TabIndex        =   46
         Top             =   900
         Width           =   1200
      End
      Begin VB.Label Label18 
         Caption         =   "CMP Name"
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
         Left            =   3577
         TabIndex        =   45
         Top             =   180
         Width           =   1080
      End
      Begin VB.Label Label17 
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
         Height          =   240
         Left            =   12915
         TabIndex        =   44
         Top             =   180
         Width           =   1080
      End
      Begin VB.Label Label15 
         Caption         =   "New Rate"
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
         Left            =   7387
         TabIndex        =   43
         Top             =   885
         Width           =   900
      End
      Begin VB.Label Label9 
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
         Height          =   240
         Left            =   8805
         TabIndex        =   42
         Top             =   180
         Width           =   1560
      End
      Begin VB.Label Label6 
         Caption         =   "Location Name"
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
         Left            =   735
         TabIndex        =   41
         Top             =   180
         Width           =   1440
      End
      Begin VB.Label Label11 
         Caption         =   "Billing Unit"
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
         Left            =   3615
         TabIndex        =   40
         Top             =   900
         Width           =   1005
      End
      Begin VB.Label Label12 
         Caption         =   "Billing Cycle"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   855
         TabIndex        =   39
         Top             =   900
         Width           =   1200
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   15015
      Begin VB.TextBox Text3 
         BackColor       =   &H0000C000&
         Enabled         =   0   'False
         Height          =   285
         Left            =   9840
         TabIndex        =   48
         Top             =   1200
         Width           =   495
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
         Height          =   450
         Left            =   13935
         TabIndex        =   27
         Top             =   690
         Width           =   975
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
         Left            =   6600
         Sorted          =   -1  'True
         TabIndex        =   22
         Top             =   360
         Width           =   3855
      End
      Begin VB.ComboBox CmbSGodown 
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
         Left            =   10500
         Sorted          =   -1  'True
         TabIndex        =   21
         Top             =   360
         Width           =   1365
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
         Left            =   3480
         Sorted          =   -1  'True
         TabIndex        =   19
         Top             =   360
         Width           =   3060
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H0000FFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   17
         Top             =   1200
         Width           =   495
      End
      Begin VB.ComboBox CmbVendor 
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
         Left            =   1995
         Sorted          =   -1  'True
         TabIndex        =   12
         Top             =   780
         Width           =   6465
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
         Left            =   60
         TabIndex        =   11
         Top             =   360
         Width           =   1890
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
         ItemData        =   "FrmTxn_FranchiseeRent.frx":0037
         Left            =   1995
         List            =   "FrmTxn_FranchiseeRent.frx":0039
         TabIndex        =   10
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton CmdShow 
         Caption         =   "Show"
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
         Left            =   12960
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H000000FF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2760
         TabIndex        =   8
         Top             =   1200
         Width           =   495
      End
      Begin VB.CommandButton CmdShowAll 
         Caption         =   "Show All"
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
         Left            =   11880
         TabIndex        =   7
         Top             =   720
         Visible         =   0   'False
         Width           =   975
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
         Height          =   450
         Left            =   12960
         TabIndex        =   6
         Top             =   690
         Width           =   975
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
         Height          =   450
         Left            =   13935
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label13 
         Caption         =   "Rate Updated by Employee on this Screen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   10485
         TabIndex        =   47
         Top             =   1200
         Width           =   4455
      End
      Begin VB.Label LblTotalRecord 
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
         Left            =   11280
         TabIndex        =   26
         Top             =   840
         Width           =   90
      End
      Begin VB.Label Label7 
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
         Left            =   8880
         TabIndex        =   25
         Top             =   840
         Width           =   2295
      End
      Begin VB.Label lblCMP 
         AutoSize        =   -1  'True
         Caption         =   "CMP Name"
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
         Left            =   8070
         TabIndex        =   24
         Top             =   120
         Width           =   1035
      End
      Begin VB.Label Label5 
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
         Left            =   10755
         TabIndex        =   23
         Top             =   120
         Width           =   840
      End
      Begin VB.Label LblLocationID 
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
         Left            =   4530
         TabIndex        =   20
         Top             =   120
         Width           =   840
      End
      Begin VB.Label Label4 
         Caption         =   "Godown Closed"
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
         Left            =   720
         TabIndex        =   18
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label8 
         Caption         =   "Franchisee Rate not availabe in Franchisee rate card master"
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
         Left            =   3360
         TabIndex        =   16
         Top             =   1200
         Width           =   6375
      End
      Begin VB.Label Label1 
         Caption         =   "Franchisee Name"
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
         TabIndex        =   15
         Top             =   825
         Width           =   1785
      End
      Begin VB.Label Label2 
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
         Height          =   255
         Left            =   668
         TabIndex        =   14
         Top             =   120
         Width           =   675
      End
      Begin VB.Label Label3 
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
         Height          =   255
         Left            =   2475
         TabIndex        =   13
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      Height          =   6735
      Left            =   120
      TabIndex        =   1
      Top             =   3165
      Width           =   15015
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   345
         Left            =   120
         TabIndex        =   28
         Top             =   5940
         Width           =   14820
         _ExtentX        =   26141
         _ExtentY        =   609
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
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
         TabIndex        =   2
         Top             =   255
         Width           =   1245
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   5295
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   9340
         _Version        =   393216
         FixedCols       =   0
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
      Begin MSComctlLib.ProgressBar ProgressBar2 
         Height          =   345
         Left            =   120
         TabIndex        =   51
         Top             =   6360
         Width           =   14820
         _ExtentX        =   26141
         _ExtentY        =   609
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   8280
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1296
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "FrmTxn_FranchiseeRent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mVendorID, mEmployeeName, mDirectReportingTo As Variant
Const strChecked = "þ"
Const strUnChecked = "q"
Dim mUserTypeFlag As Variant
Dim mSelRow As Variant
Dim mSelCol As Variant
Dim mFlagChk, mClaimID As Variant
Dim mMonth, mExpenseFrom, mExpenseTo, mTotalDays, mRow As Variant
Dim mMsg, mGridMsg As String
Dim oSMTP As SMTPSession
Dim mYear As Variant
Dim mGodownID, mCMPID, mLocationID, mBillingCycleID, mFlag As Variant
Dim wc As Variant
Dim mTFlagOnUpdate, mTxnNo As Variant
Dim mSM_Prefix As Variant
Dim mSelectedGSL As Variant

Private Sub ChkCheckAll_Click()

If MSHFlexGrid1.Rows = 2 Then
   ChkCheckAll.Value = 0
   Exit Sub
End If

If ChkCheckAll.Value = vbChecked Then
   For I = 1 To MSHFlexGrid1.Rows - 2
       MSHFlexGrid1.row = I
       MSHFlexGrid1.Col = 0
       If MSHFlexGrid1.CellBackColor <> vbRed And MSHFlexGrid1.CellBackColor <> vbYellow Then
          MSHFlexGrid1.row = I
          MSHFlexGrid1.TextMatrix(I, 0) = strChecked
       End If
   Next
Else
   For I = 1 To MSHFlexGrid1.Rows - 2
       MSHFlexGrid1.row = I
       MSHFlexGrid1.TextMatrix(I, 0) = strUnChecked
   Next
End If
End Sub


Private Sub CmbLocationID_GotFocus()
tmp = CmbLocationID.Text

If Gen_WcLocation = "" Then
   Call TableMst_Location("")
Else
   Call TableMst_Location("Where " & Gen_WcLocation)
End If

CmbLocationID.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No record found!!! "
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbLocationID.AddItem RsMst_Location!Locationname
    RsMst_Location.MoveNext
Next
CmbLocationID.Text = tmp
End Sub
'----
Private Sub CmbLocationID_LostFocus()
If CmbLocationID.Text = "" Then
   mLocationID = Null
   Call Grid_Head
   Call Grid_Head1
'   Frame4.Enabled = False
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocationID.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection!!! "
   CmbLocationID.SetFocus
   Exit Sub
End If
If mLocationID <> RsMst_Location!LocationID Then
   Call Grid_Head
End If
mLocationID = RsMst_Location!LocationID
mSM_Prefix = RsMst_Location!SM_Prefix

End Sub


Private Sub CmbVendor_GotFocus()
tmp1 = CmbVendor.Text
If CmbLocationID.Text <> "" Then
   tmp = "select distinct MV.VendorID,MV.VendorName from"
   tmp = tmp & " Mst_FF_RateCard MF inner join Mst_Vendor MV on MV.VendorID=MF.VendorID"
   tmp = tmp & " Where MF.LocationID = " & mLocationID
Else
   tmp = "Select  MV.VendorID,MV.VendorName from Mst_Vendor MV"
End If

CmbVendor.Clear
Call TmpTable
If TmpRs.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To TmpRs.RecordCount
    CmbVendor.AddItem TmpRs!VendorName
    TmpRs.MoveNext
Next
CmbVendor.Text = tmp1

End Sub
'----
Private Sub CmbVendor_LostFocus()
Call Grid_Head
If CmbVendor.Text = "" Then
   Exit Sub
End If
TmpRs.MoveFirst
TmpRs.Find "VendorName = '" & CmbVendor.Text & "'"
If TmpRs.EOF Then
   MsgBox "Invalid selection "
   CmbVendor.SetFocus
   Exit Sub
End If
mVendorID = TmpRs!VendorID
End Sub

Private Sub CmbMonth_LostFocus()
Call Grid_Head
Call ClearFrame
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
mMonth = CmbMonth.ListIndex + 1
If CmbYear.Text <> "" Then
   mTotalDays = Day(DateSerial(mYear, mMonth + 1, 0))
End If
Call Grid_Head

End Sub

Private Sub CmbYear_GotFocus()
If CmbMonth.Text = "" Then
   MsgBox "Select Month!!!"
   CmbMonth.SetFocus
   Exit Sub
End If
tmp = CmbYear.Text
CmbYear.Clear
For I = 2010 To Year(Now)
    CmbYear.AddItem I
Next

If Month(Now) = 12 Then
   CmbYear.AddItem Year(Now) + 1
End If

CmbYear.Text = tmp
End Sub

Private Sub CmbYear_LostFocus()
Call Grid_Head
Call ClearFrame
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

mYear = CmbYear.Text
mTotalDays = Day(DateSerial(mYear, mMonth + 1, 0))

End Sub

Private Sub CmdClear_Click()
CmbMonth.Text = ""
CmbYear.Text = ""
CmbLocationID.Text = ""
CmbCMPName.Text = ""
CmbSGodown.Text = ""
CmbVendor.Text = ""
Call Grid_Head
Call Grid_Head1
Call ClearFrame
LblTotalRecord.Caption = ""
ChkCheckAll.Value = 0
SaveCmd.Enabled = True
ProgressBar1.Value = 0
ProgressBar2.Value = 0
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_GodownRent.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_GodownRent.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
ProgressBar1.Value = 0
ProgressBar2.Value = 0
ProgressBar1.Max = MSHFlexGrid1.Rows + 1
For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = I + 1 ' + 6
    For C = 1 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(I, C), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(I, C)) Then
              oWS.Cells(mRow, C) = Format(MSHFlexGrid1.TextMatrix(I, C), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, C) = MSHFlexGrid1.TextMatrix(I, C)
           End If
        Else
           oWS.Cells(mRow, C) = MSHFlexGrid1.TextMatrix(I, C)
        End If
        MSHFlexGrid1.row = I
        MSHFlexGrid1.Col = C
                
        If MSHFlexGrid1.CellBackColor = vbRed Then
           oWS.Range(oWS.Cells(mRow, C), oWS.Cells(mRow, C)).Interior.Color = vbRed
''        ElseIf MSHFlexGrid1.CellBackColor = vbGreen Then
''           oWS.Range(oWS.Cells(mRow, c), oWS.Cells(mRow, c)).Interior.Color = vbGreen
        ElseIf MSHFlexGrid1.CellBackColor = vbYellow Then
           oWS.Range(oWS.Cells(mRow, C), oWS.Cells(mRow, C)).Interior.Color = vbYellow
        End If
              
    Next
    ProgressBar1.Value = ProgressBar1.Value + 1
Next
ProgressBar1.Value = ProgressBar1.Max
oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub


Private Sub CmdShow_Click()
Dim vMonth As Variant

wc = ""
Call Grid_Head
Call ClearFrame
ChkCheckAll.Value = 0
ProgressBar1.Value = 0
ProgressBar2.Value = 0

'If CmbVendor.Text = "" Then
'   MsgBox "Please Select Vendor Name From List !!!"
'   CmbVendor.SetFocus
'   Exit Sub
'End If

If CmbMonth.Text = "" Then
   MsgBox " Blank Month not allowed  !!!"
   CmbMonth.SetFocus
   Exit Sub
End If

If CmbYear.Text = "" Then
   MsgBox " Blank Year not allowed  !!!"
   CmbYear.SetFocus
   Exit Sub
End If

If CmbLocationID.Text = "" Then
   MsgBox " Blank Location not allowed  !!!"
   CmbLocationID.SetFocus
   Exit Sub
End If

''If CheckGSLExist = True Then
''   If CmbVendor.Text = "" Then
''      MsgBox " Franchisee name is compulsory,as this Location has no franchisee GSL!!"
''      CmbVendor.SetFocus
''      Exit Sub
''   End If
''End If

If CmbLocationID.Text <> "" Then
   wc = " And ML.LocationName = '" & CmbLocationID.Text & "'"
Else
   If Gen_WcLocation <> "" Then
      wc = " And ML." & Gen_WcLocation
   End If
End If

If CmbCMPName.Text <> "" Then
   wc = wc & " And MCMP.CMPName = '" & CmbCMPName.Text & "'"
Else
   If Gen_WcCMP <> "" Then
      wc = wc & " And MCMP." & Gen_WcCMP
   End If
End If

If CmbSGodown.Text <> "" Then
   wc = wc & " And MG.GodownNo = '" & CmbSGodown.Text & "'"
End If

If CmbVendor.Text <> "" Then
   wc = wc & " And MV.VendorName = '" & CmbVendor.Text & "'"
End If

'If mMonth <= 3 And Val(CmbYear.Text) <= 2012 Then
'   MsgBox " Month earlier to April 2012 is not allowed  !!!"
'   CmbMonth.SetFocus
'   Exit Sub
'End If

Dim mSelDate, mCurrDate, mFutDate As Variant

mSelDate = CmbYear.Text & "-" & mMonth & "-01"
mCurrDate = Year(Now) & "-" & Month(Now) & "-01"
mFutDate = DateAdd("m", 1, mCurrDate)
mFutDate = Format(mFutDate, "yyyy-MM-dd")
mSelDate = Format(mSelDate, "yyyy-MM-dd")
If mSelDate > mFutDate Then
   MsgBox "Future Month not Allowed !!!"
   CmbMonth.SetFocus
   Exit Sub
End If

mExpenseFrom = Format(CmbYear.Text & "-" & mMonth & "-" & "01", "yyyy-mm-dd")
mExpenseTo = Format(CmbYear.Text & "-" & mMonth & "-" & mTotalDays, "yyyy-mm-dd")

If Format(Gen_BlockingDate, "yyyy-mm") >= Format(CDate(mExpenseFrom), "yyyy-mm") Then
   MsgBox "Select next month from " & Format(Gen_BlockingDate, "MMM-yyyy") & "!!! "
   CmbMonth.SetFocus
   Exit Sub
End If

MsgBox "Please Wait till next message !!"

''If CmbVendor.Text <> "" Then
''   If CheckGSLExist = True Then
''      tmp = "Insert into Txn_InvMonGSLDetail (GSLID,SM_Prefix,CMPID,CommodityID,GodownID,StackNo,LotNo,FranchiseeRate,FranchBillingCycleID,"
''      tmp = tmp & " FranchBillingUnit,FranchNoofDays,FranchInvoiceAmount,FF_RateCardID,FranchProcessFlag,YearMonth,TFlag,ISINID)"
''      tmp = tmp & " Select 0,ML.SM_Prefix,CMPID,0 CommodityID,MFG.GodownID,'Dummy-F' as StackNo,'Dummy-F' as LotNo ,MFG.RateOnQty as FranchiseeRate,"
''      tmp = tmp & " MFg.BillingCycleID,MFG.BillingUnit,1 as FranchNoofDays,MFG.RateOnQty as FranchInvoiceAmount,MF.FF_RateCardID,'Y','" & mExpenseFrom & "','','" & "V" & mVendorID & "'"
''      tmp = tmp & " From Mst_FF_RateCard_Godown MFG"
''      tmp = tmp & " inner join Mst_FF_RateCard MF on MF.FF_RateCardID=MFG.FF_RateCardID"
''      'tmp = tmp & " Inner join Mst_Commodity MC on MC.CommodityGroupID=MF.CommodityGroupID"
''      tmp = tmp & " Inner join Mst_Location ML on ML.LocationID=MF.LocationID"
''      tmp = tmp & " Inner join Mst_Vendor MD on MD.VendorID=MF.VendorID"
''      tmp = tmp & " Where MFG.BillingUnit='F' And  MF.LocationID=" & mLocationID & " And MF.VendorID=" & mVendorID
''
''      Call TmpTable
''   End If
''End If
    

tmp = "Select ML.LocationName, TInv.SM_PreFix,MCMP.CMPName,MG.GodownNo+ '~'+ MGSL.StackNo + '~' + MGSL.LotNo 'GSL',MV.VendorName,MCG.CommodityGroup, Me.ExchangeType ,"
tmp = tmp & " '',Sum (TInv.BalanceBags) 'BalanceBags',Sum (TINV.BalanceWeight) 'BalanceWeight', TINV.Bagsize,TINV.YearMonth,"
tmp = tmp & " TINV.CxTFDate ,TINV.TFlag, MCMP.LocationID , TINV.CMPID, MF.VendorID, TINV.CommodityID, MC.CommodityGroupID,"
tmp = tmp & " TINV.ExchangeTypeID,TINV.Mst_GSL_ID,TINV.CDTFDate,MC.Commodity,MST.NAVCode,MBC.BillingCycle,TINV.FranchNoofDays,TINV.FranchiseeRate,"
If CDate(mExpenseFrom) < "2022-02-01" Then
   tmp = tmp & " TInv.Bagsize , Sum(TInv.FranchInvoiceAmount) 'InvoiceAmount',TINV.FranchBillingCycleID,TINV.FranchBillingUnit,TINV.CxTF_DetailsID,MG.CloseDate,TINV.GodownID,Sum(TINV.InvoiceAmount) 'InvoiceAmt',Sum(TINV.InsuranceCostDeduction) 'InsuranceAmt'"
Else
   tmp = tmp & " TInv.Bagsize , Sum(TInv.FranchInvoiceAmount) 'InvoiceAmount',TINV.FranchBillingCycleID,TINV.FranchBillingUnit,TINV.CxTF_DetailsID,MG.CloseDate,TINV.GodownID,Sum(TINV.InvoiceAmtForFranch) 'InvoiceAmt',Sum(TINV.InsuranceCostDeduction) 'InsuranceAmt'"
End If
tmp = tmp & " From Txn_InvMonGSLDetail TINV"
tmp = tmp & " Left Join Mst_Commodity MC On TINV.CommodityID = MC.CommodityID"
tmp = tmp & " Left Join Mst_CommodityGroup MCG On MC.CommodityGroupID = MCG.CommodityGroupID"
tmp = tmp & " Inner Join Mst_CMP MCMP On TINV.CMPID = MCMP.CMPID"
tmp = tmp & " Inner Join Mst_Godown MG On TINV.GodownID = MG.GodownID"
tmp = tmp & " Inner Join Mst_Location ML On MCMP.LocationID = ML.LocationID"
tmp = tmp & " Left Join Mst_GSL MGSL On TINV.Mst_GSL_ID = MGSL.Mst_GSL_ID And TINV.SM_Prefix = MGSL.SM_Prefix"
tmp = tmp & " left Join Mst_FF_RateCard MF on MF.FF_RateCardID=TINV.FF_RateCardID"
tmp = tmp & " left Join Mst_Vendor MV on MV.VendorID=MF.VendorID"
tmp = tmp & " left Join Mst_ExchangeType ME On TINV.ExchangeTypeID = ME.ExchangeTypeID"
tmp = tmp & " Left Join Mst_BillingCycle MBC On TINV.FranchBillingCycleID = MBC.BillingCycleID"
tmp = tmp & " Left Join Mst_ServiceTax MST On MC.ServiceTaxID = MST.ServiceTaxID"
'Parag
'If CDate(mExpenseFrom) < "2021-08-01" Then
'   tmp = tmp & " Where MG.UnitTypeID in (2,6) And TINV.YearMonth = '" & Format(mExpenseFrom, Gen_DatFormat) & "' And TINV.FranchClaimID is null"
'Else
 If CDate(mExpenseFrom) < "2022-02-01" Then
   tmp = tmp & " Where MG.UnitTypeID in (2,6) And TINV.YearMonth = '" & Format(mExpenseFrom, Gen_DatFormat) & "' And Isnull(InvoiceAmount,0) > 0 And TINV.MonG_UID not in (Select Distinct MonG_UID From Txn_MonthlyFranchiseeRentShare Where YearMonth = '" & CDate(mExpenseFrom) & "') "
 Else
 'InvoiceAmount
   tmp = tmp & " Where MG.UnitTypeID in (2,6) And TINV.YearMonth = '" & Format(mExpenseFrom, Gen_DatFormat) & "' And Isnull(InvoiceAmtForFranch,0) > 0 And TINV.MonG_UID not in (Select Distinct MonG_UID From Txn_MonthlyFranchiseeRentShare Where YearMonth = '" & CDate(mExpenseFrom) & "') "
 End If
'End If
''If Gen_WcLocation <> "" Then
''   tmp = tmp & " And ML." & Gen_WcLocation
''End If
tmp = tmp & " And isnull(TINV.FranchProcessFlag,'N') = 'Y' "
tmp = tmp & wc
tmp = tmp & " Group by ML.LocationName , TInv.SM_PreFix, MCMP.CMPName,MG.GodownNo,MGSL.StackNo,MGSL.LotNo, MF.VendorID, MCG.CommodityGroup, Me.ExchangeType,"
tmp = tmp & " TINV.Bagsize,TINV.YearMonth,"
tmp = tmp & " TINV.CxTFDate ,TINV.TFlag, MCMP.LocationID , TINV.CMPID, MV.VendorName, TINV.CommodityID, MC.CommodityGroupID,"
tmp = tmp & " TINV.ExchangeTypeID,TINV.Mst_GSL_ID,TINV.CDTFDate,MC.Commodity,MST.NAVCode,MBC.BillingCycle,TINV.FranchNoofDays,TINV.FranchiseeRate,"
tmp = tmp & " TInv.Bagsize , TInv.FranchBillingCycleID, TInv.FranchBillingUnit, TInv.CxTF_DetailsID,MG.CloseDate,TINV.GodownID"
'--,TINV.InvoiceAmount"

Call TmpTable

ProgressBar1.Max = TmpRs.RecordCount + 2
mRow = 1
If TmpRs.RecordCount > 0 Then
   For m = 1 To TmpRs.RecordCount
       With MSHFlexGrid1
         mR = MSHFlexGrid1.Rows - 1
        .TextMatrix(mR, 1) = TmpRs!Locationname
        .TextMatrix(mR, 2) = TmpRs!CMPName
        .TextMatrix(mR, 3) = IIf(IsNull(TmpRs!GSL), "", TmpRs!GSL) 'TmpRs!GodownNo
        .TextMatrix(mR, 4) = IIf(IsNull(TmpRs!VendorName), "", TmpRs!VendorName)
        .TextMatrix(mR, 5) = IIf(IsNull(TmpRs!COmmodityGroup), "", TmpRs!COmmodityGroup)
        .TextMatrix(mR, 6) = IIf(IsNull(TmpRs!ExchangeType), "", TmpRs!ExchangeType)
        .TextMatrix(mR, 7) = IIf(IsNull(TmpRs!tFlag), "", TmpRs!tFlag)
        .TextMatrix(mR, 8) = Format(IIf(IsNull(TmpRs!YearMonth), "", TmpRs!YearMonth), "dd-MMM-yyyy")
        .TextMatrix(mR, 9) = Format(IIf(IsNull(TmpRs!CxTFDate), "", TmpRs!CxTFDate), "dd-MMM-yyyy")
        .TextMatrix(mR, 10) = IIf(IsNull(TmpRs!CDTFDate), "", TmpRs!CDTFDate)
        .TextMatrix(mR, 10) = Format(.TextMatrix(mR, 10), "dd-mmm-yyyy")
        
        .TextMatrix(mR, 11) = IIf(IsNull(TmpRs!BalanceBags), 0, TmpRs!BalanceBags)
        .TextMatrix(mR, 12) = IIf(IsNull(TmpRs!BalanceWeight), "", TmpRs!BalanceWeight)
    
        .TextMatrix(mR, 13) = IIf(IsNull(TmpRs!BillingCycle), "", TmpRs!BillingCycle) '"Billing Cycle"
        .TextMatrix(mR, 14) = IIf(IsNull(TmpRs!FranchNoofDays), "", TmpRs!FranchNoofDays) '"No of Cycles"
        .TextMatrix(mR, 15) = IIf(IsNull(TmpRs!FranchiseeRate), "", TmpRs!FranchiseeRate) '"Rate"
    
        
        .TextMatrix(mR, 16) = Format(IIf(IsNull(TmpRs!BagSize), 0, TmpRs!BagSize), "########0.00")
        .TextMatrix(mR, 17) = Format(IIf(IsNull(TmpRs!InvoiceAmount), 0, TmpRs!InvoiceAmount), "########0.0000") '"Total Amount"
        If .TextMatrix(mR, 7) = "" Then
           .TextMatrix(mR, 18) = TmpRs!GodownID
        Else
           .TextMatrix(mR, 18) = TmpRs!Mst_GSL_ID
        End If
        .TextMatrix(mR, 19) = IIf(IsNull(TmpRs!NavCode), "", TmpRs!NavCode)
        .TextMatrix(mR, 20) = IIf(IsNull(TmpRs!VendorID), "", TmpRs!VendorID)
        .TextMatrix(mR, 21) = IIf(IsNull(TmpRs!CommodityID), "", TmpRs!CommodityID)
        .TextMatrix(mR, 22) = IIf(IsNull(TmpRs!ExchangeTypeID), "", TmpRs!ExchangeTypeID)
        .TextMatrix(mR, 23) = IIf(IsNull(TmpRs!LocationID), "", TmpRs!LocationID)
        .TextMatrix(mR, 24) = IIf(IsNull(TmpRs!SM_Prefix), "", TmpRs!SM_Prefix)
        .TextMatrix(mR, 25) = IIf(IsNull(TmpRs!FranchBillingCycleID), "", TmpRs!FranchBillingCycleID) '"BillingCycleID"
        .TextMatrix(mR, 26) = IIf(IsNull(TmpRs!FranchBillingUnit), "", TmpRs!FranchBillingUnit) '"BillingUnit"
        .TextMatrix(mR, 27) = IIf(IsNull(TmpRs!CxTF_DetailsID), "", TmpRs!CxTF_DetailsID) '"BillingUnit"
        .TextMatrix(mR, 28) = IIf(IsNull(TmpRs!InvoiceAmt), 0, TmpRs!InvoiceAmt) '"BillingUnit"
        .TextMatrix(mR, 29) = IIf(IsNull(TmpRs!InsuranceAmt), 0, TmpRs!InsuranceAmt) '"BillingUnit"
        If LCase(.TextMatrix(mR, 26)) = "b" Then
           .TextMatrix(mR, 26) = "Bags"
        ElseIf LCase(.TextMatrix(mR, 26)) = "q" Then
           .TextMatrix(mR, 26) = "Quantity"
        ElseIf LCase(.TextMatrix(mR, 26)) = "p" Then
           .TextMatrix(mR, 26) = "Percentage"
        ElseIf LCase(.TextMatrix(mR, 26)) = "f" Then
           .TextMatrix(mR, 26) = "Fixed"
        End If
        
    'if Rate = 0 then mark as red
      
      If (Val(.TextMatrix(mR, 15)) = 0) Then
         For I = 0 To MSHFlexGrid1.Cols - 1
             .Col = I
             .row = mR
             .CellBackColor = vbRed
         Next
      End If
      
    'if D and W in same month then mark as green
''      If .TextMatrix(mr, 7) = "W" Then
''         If Month(.TextMatrix(mr, 10)) & "~" & Year(.TextMatrix(mr, 10)) = Month(.TextMatrix(mr, 9)) & "~" & Year(.TextMatrix(mr, 9)) Then
''            For I = 0 To MSHFlexGrid1.Cols - 1
''                .Col = I
''                .Row = mr
''                .CellBackColor = vbGreen
''            Next
''         End If
''      End If
      If IsNull(TmpRs!CloseDate) = False Then
        If TmpRs!CloseDate < CDate(mExpenseFrom) Then
           For I = 0 To MSHFlexGrid1.Cols - 1
               .Col = I
               .row = mR
               .CellBackColor = vbYellow
           Next
        End If
      End If
      .row = mR
      .Col = 0
      .CellFontName = "Wingdings"
      .CellFontSize = 12
      .CellAlignment = flexAlignCenterCenter
      .Text = strUnChecked
      TmpRs.MoveNext
      ProgressBar1.Value = ProgressBar1.Value + 1
      MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
      End With
   Next
End If
ProgressBar1.Value = ProgressBar1.Max

LblTotalRecord.Caption = MSHFlexGrid1.Rows - 2
SaveCmd.Enabled = True
MsgBox "Done !!"


End Sub

Private Sub CmdShowAll_Click()

SaveCmd.Enabled = False
Dim vMonth As Variant

Call Grid_Head
Call ClearFrame
ChkCheckAll.Value = 0
CmbVendor.Text = ""
LblTotalRecord.Caption = "-"

ProgressBar1.Value = 0
ProgressBar2.Value = 0

If CmbMonth.Text = "" Then
   MsgBox " Blank Month not allowed  !!!"
   CmbMonth.SetFocus
   Exit Sub
End If

If CmbYear.Text = "" Then
   MsgBox " Blank Year not allowed  !!!"
   CmbYear.SetFocus
   Exit Sub
End If

If CmbLocationID.Text <> "" Then
   wc = " And ML.LocationName = '" & CmbLocationID.Text & "'"
Else
   If Gen_WcLocation <> "" Then
      wc = " And ML." & Gen_WcLocation
   End If
End If

If CmbCMPName.Text <> "" Then
   wc = wc & " And MCMP.CMPName = '" & CmbCMPName.Text & "'"
Else
   If Gen_WcCMP <> "" Then
      wc = wc & " And MCMP." & Gen_WcCMP
   End If
End If

If CmbSGodown.Text <> "" Then
   wc = wc & " And MG.GodownNo = '" & CmbSGodown.Text & "'"
End If

If CmbVendor.Text <> "" Then
   wc = wc & " And MV.VendorName = '" & CmbVendor.Text & "'"
End If


'Commented on 27 Dec 2010 'Parag
'If mMonth <= 3 Then
'   MsgBox " Month earlier to April 2010 is not allowed  !!!"
'   CmbMonth.SetFocus
'   Exit Sub
'End If


If mMonth <= 3 And Val(CmbYear.Text) <= 2010 Then
   MsgBox " Month earlier to April 2010 is not allowed  !!!"
   CmbMonth.SetFocus
   Exit Sub
End If

Dim mSelDate, mCurrDate, mFutDate As Variant

mSelDate = CmbYear.Text & "-" & mMonth & "-01"
mCurrDate = Year(Now) & "-" & Month(Now) & "-01"
mFutDate = DateAdd("m", 1, mCurrDate)
mFutDate = Format(mFutDate, "yyyy-MM-dd")
mSelDate = Format(mSelDate, "yyyy-MM-dd")
If mSelDate > mFutDate Then
   MsgBox "Future Month not Allowed !!!"
   CmbMonth.SetFocus
   Exit Sub
End If


'Commented on 27 Dec 2010 'Parag
'vMonth = Month(Now) + 1
'
'If mMonth > vMonth Then
'   MsgBox "Future Month not Allowed !!!"
'   CmbMonth.SetFocus
'   Exit Sub
'End If
'
'If CmbMonth.Text = "" Then
'   MsgBox "Please Select Month"
'   CmbMonth.SetFocus
'   Exit Sub
'End If
'
'If CmbYear.Text = "" Then
'   MsgBox "Please Select Year "
'   CmbYear.SetFocus
'   Exit Sub
'End If

''If mUserTypeFlag = "cmpm" Then
''   mVendorID = Gen_EmployeeID
''End If


mExpenseFrom = Format(CmbYear.Text & "-" & mMonth & "-" & "01", "yyyy-mm-dd")
mExpenseTo = Format(CmbYear.Text & "-" & mMonth & "-" & mTotalDays, "yyyy-mm-dd")

MsgBox "Please Wait till next message !!"

tmp = "Select ML.LocationName, TInv.SM_PreFix,MCMP.CMPName,MG.GodownNo+'~'+MGSL.StackNo+'~'+MGSL.LotNo 'GSL',MV.VendorName,MCG.CommodityGroup, Me.ExchangeType ,"
tmp = tmp & " '',Sum (TInv.BalanceBags) 'BalanceBags',Sum (TINV.BalanceWeight) 'BalanceWeight', TINV.Bagsize,TINV.YearMonth,"
tmp = tmp & " TINV.CxTFDate ,TINV.TFlag, MCMP.LocationID , TINV.CMPID, MF.VendorID, TINV.CommodityID, MC.CommodityGroupID,"
tmp = tmp & " TINV.ExchangeTypeID,TINV.Mst_GSL_ID,TINV.CDTFDate,MC.Commodity,MST.NAVCode,MBC.BillingCycle,TINV.FranchNoofDays,TINV.FranchiseeRate,"
tmp = tmp & " TInv.Bagsize , Sum(TInv.FranchInvoiceAmount) 'InvoiceAmount',TINV.FranchBillingCycleID,TINV.FranchBillingUnit,TINV.CxTF_DetailsID,MG.CloseDate, Sum(TINV.InvoiceAmtForFranch) 'InvoiceAmt'"
tmp = tmp & " From Txn_InvMonGSLDetail TINV"
tmp = tmp & " Inner Join Mst_Commodity MC On TINV.CommodityID = MC.CommodityID"
tmp = tmp & " Inner Join Mst_CommodityGroup MCG On MC.CommodityGroupID = MCG.CommodityGroupID"
tmp = tmp & " Inner Join Mst_CMP MCMP On TINV.CMPID = MCMP.CMPID"
tmp = tmp & " Inner Join Mst_Godown MG On TINV.GodownID = MG.GodownID"
tmp = tmp & " Inner Join Mst_Location ML On MCMP.LocationID = ML.LocationID"
tmp = tmp & " left Join Mst_FF_RateCard MF on MF.FF_RateCardID=TINV.FF_RateCardID"
tmp = tmp & " left Join Mst_Vendor MV on MV.VendorID=MF.VendorID"
tmp = tmp & " Inner Join Mst_ExchangeType ME On TINV.ExchangeTypeID = ME.ExchangeTypeID"
tmp = tmp & " Left Join Mst_GSL MGSL On TINV.Mst_GSL_ID = MGSL.Mst_GSL_ID And TINV.SM_PRefix = MGSL.SM_Prefix "
tmp = tmp & " Left Join Mst_BillingCycle MBC On TINV.FranchBillingCycleID = MBC.BillingCycleID"
tmp = tmp & " Left Join Mst_ServiceTax MST On MC.ServiceTaxID = MST.ServiceTaxID"
'tmp = tmp & " Where MG.UnitTypeID in (2,6) And TINV.YearMonth = '" & Format(mExpenseFrom, Gen_DatFormat) & "' And TINV.FranchClaimID is null And isnull(TINV.FranchProcessFlag,'N')='Y' "
'Parag
If CDate(mExpenseFrom) < "2021-08-01" Then
   tmp = tmp & " Where MG.UnitTypeID in (2,6) And TINV.YearMonth = '" & Format(mExpenseFrom, Gen_DatFormat) & "' And TINV.FranchClaimID is null"
Else
   tmp = tmp & " Where MG.UnitTypeID in (2,6) And TINV.YearMonth = '" & Format(mExpenseFrom, Gen_DatFormat) & "' And Isnull(InvoiceAmount,0) > 0 And TINV.MonG_UID not in (Select Distinct MonG_UID From Txn_MonthlyFranchiseeRentShare Where YearMonth = '" & mExpenseFrom & "') "
End If

tmp = tmp & wc
tmp = tmp & " Group by ML.LocationName , TInv.SM_PreFix, MCMP.CMPName,MG.GodownNo, MF.VendorID, MCG.CommodityGroup, Me.ExchangeType,"
tmp = tmp & " TINV.Bagsize,TINV.YearMonth,"
tmp = tmp & " TINV.CxTFDate ,TINV.TFlag, MCMP.LocationID , TINV.CMPID, MV.VendorName, TINV.CommodityID, MC.CommodityGroupID,"
tmp = tmp & " TINV.ExchangeTypeID,TINV.Mst_GSL_ID,TINV.CDTFDate,MC.Commodity,MST.NAVCode,MBC.BillingCycle,TINV.FranchNoofDays,TINV.FranchiseeRate,"
tmp = tmp & " TInv.Bagsize , TInv.FranchBillingCycleID, TInv.FranchBillingUnit, TInv.CxTF_DetailsID,MG.CloseDate,MGSL.StackNo,MGSL.LotNo"
tmp = tmp & " Order By ML.LocationName"

Call TmpTable

mRow = 1
ProgressBar1.Max = TmpRs.RecordCount + 2

If TmpRs.RecordCount > 0 Then
   For m = 1 To TmpRs.RecordCount
       With MSHFlexGrid1
        mR = MSHFlexGrid1.Rows - 1
        
        .TextMatrix(mR, 1) = TmpRs!Locationname
        .TextMatrix(mR, 2) = TmpRs!CMPName
        .TextMatrix(mR, 3) = IIf(IsNull(TmpRs!GSL), "", TmpRs!GSL) 'TmpRs!GodownNo
        .TextMatrix(mR, 4) = IIf(IsNull(TmpRs!VendorName), "", TmpRs!VendorName)
        .TextMatrix(mR, 5) = TmpRs!COmmodityGroup
        .TextMatrix(mR, 6) = IIf(IsNull(TmpRs!ExchangeType), "", TmpRs!ExchangeType)
        .TextMatrix(mR, 7) = TmpRs!tFlag
        .TextMatrix(mR, 8) = Format(IIf(IsNull(TmpRs!YearMonth), "", TmpRs!YearMonth), "dd-MMM-yyyy")
        .TextMatrix(mR, 9) = Format(IIf(IsNull(TmpRs!CxTFDate), "", TmpRs!CxTFDate), "dd-MMM-yyyy")
        .TextMatrix(mR, 10) = IIf(IsNull(TmpRs!CDTFDate), "", TmpRs!CDTFDate)
        .TextMatrix(mR, 10) = Format(.TextMatrix(mR, 10), "dd-mmm-yyyy")
        .TextMatrix(mR, 11) = IIf(IsNull(TmpRs!BalanceBags), 0, TmpRs!BalanceBags)
        .TextMatrix(mR, 12) = IIf(IsNull(TmpRs!BalanceWeight), "", TmpRs!BalanceWeight)
        .TextMatrix(mR, 13) = IIf(IsNull(TmpRs!BillingCycle), "", TmpRs!BillingCycle) '"Billing Cycle"
        .TextMatrix(mR, 14) = IIf(IsNull(TmpRs!FranchNoofDays), "", TmpRs!FranchNoofDays) '"No of Cycles"
        .TextMatrix(mR, 15) = IIf(IsNull(TmpRs!FranchiseeRate), "", TmpRs!FranchiseeRate) '"Rate"
        .TextMatrix(mR, 16) = Format(IIf(IsNull(TmpRs!BagSize), 0, TmpRs!BagSize), "########0.00")
        .TextMatrix(mR, 17) = Format(IIf(IsNull(TmpRs!InvoiceAmount), 0, TmpRs!InvoiceAmount), "########0.0000") '"Total Amount"
        .TextMatrix(mR, 18) = TmpRs!Mst_GSL_ID
        .TextMatrix(mR, 19) = IIf(IsNull(TmpRs!NavCode), "", TmpRs!NavCode)
        .TextMatrix(mR, 20) = IIf(IsNull(TmpRs!VendorID), "", TmpRs!VendorID)
        .TextMatrix(mR, 21) = IIf(IsNull(TmpRs!CommodityID), "", TmpRs!CommodityID)
        .TextMatrix(mR, 22) = IIf(IsNull(TmpRs!ExchangeTypeID), "", TmpRs!ExchangeTypeID)
        .TextMatrix(mR, 23) = IIf(IsNull(TmpRs!LocationID), "", TmpRs!LocationID)
        .TextMatrix(mR, 24) = IIf(IsNull(TmpRs!SM_Prefix), "", TmpRs!SM_Prefix)
        .TextMatrix(mR, 25) = IIf(IsNull(TmpRs!FranchBillingCycleID), "", TmpRs!FranchBillingCycleID) '"BillingCycleID"
        .TextMatrix(mR, 26) = IIf(IsNull(TmpRs!FranchBillingUnit), "", TmpRs!FranchBillingUnit) '"BillingUnit"
        .TextMatrix(mR, 27) = IIf(IsNull(TmpRs!CxTF_DetailsID), "", TmpRs!CxTF_DetailsID) '"BillingUnit"
        .TextMatrix(mR, 28) = IIf(IsNull(TmpRs!InvoiceAmt), 0, TmpRs!InvoiceAmt) '"Billing Amount"
        
        If LCase(.TextMatrix(mR, 26)) = "b" Then
           .TextMatrix(mR, 26) = "Bags"
        ElseIf LCase(.TextMatrix(mR, 26)) = "q" Then
           .TextMatrix(mR, 26) = "Quantity"
        ElseIf LCase(.TextMatrix(mR, 26)) = "p" Then
           .TextMatrix(mR, 26) = "Percentage"
        ElseIf LCase(.TextMatrix(mR, 26)) = "f" Then
           .TextMatrix(mR, 26) = "Fixed"
        End If
        
    'if Rate = 0 then mark as red
      If (Val(.TextMatrix(mR, 15)) = 0) Then
         For I = 0 To MSHFlexGrid1.Cols - 1
             .Col = I
             .row = mR
             .CellBackColor = vbRed
         Next
      End If
      
    'if D and W in same month then mark as green
''      If .TextMatrix(mr, 7) = "W" Then
''         If Month(.TextMatrix(mr, 10)) & "~" & Year(.TextMatrix(mr, 10)) = Month(.TextMatrix(mr, 9)) & "~" & Year(.TextMatrix(mr, 9)) Then
''            For I = 0 To MSHFlexGrid1.Cols - 1
''                .Col = I
''                .Row = mr
''                .CellBackColor = vbGreen
''            Next
''         End If
''      End If
      If IsNull(TmpRs!CloseDate) = False Then
        If TmpRs!CloseDate < CDate(mExpenseFrom) Then
           For I = 0 To MSHFlexGrid1.Cols - 1
               .Col = I
               .row = mR
               .CellBackColor = vbYellow
           Next
        End If
      End If
      .row = mR
      .Col = 0
      .CellFontName = "Wingdings"
      .CellFontSize = 12
      .CellAlignment = flexAlignCenterCenter
      .Text = strUnChecked
      
         
       TmpRs.MoveNext
       ProgressBar1.Value = ProgressBar1.Value + 1
       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
       End With
       
   Next
End If
ProgressBar1.Value = ProgressBar1.Max
LblTotalRecord.Caption = MSHFlexGrid1.Rows - 2
MsgBox "Done !!"

End Sub

Private Sub CmdUpdateList_Click()
Dim x As Variant

If CmbBillingCycle.Text = "" Then
   MsgBox "Blank Billing Cycle not allowed!!!"
   CmbBillingCycle.SetFocus
   Exit Sub
End If

If CmbBillingUnit.Text = "" Then
   MsgBox "Blank Billing Unit not allowed!!!"
   CmbBillingUnit.SetFocus
   Exit Sub
End If

If TxtNewRate.Text = "" Then
   MsgBox "Please Enter New Rate !!"
   TxtNewRate.SetFocus
   Exit Sub
End If

x = MSHFlexGrid1.RowSel

MSHFlexGrid1.TextMatrix(x, 15) = Val(TxtNewRate.Text)
MSHFlexGrid1.TextMatrix(x, 13) = CmbBillingCycle.Text
MSHFlexGrid1.TextMatrix(x, 26) = CmbBillingUnit.Text

'Update Details in InvMonGSL table
Call UpdateTxn_InvMonGSL(x)

MSHFlexGrid1.row = x

For C = 0 To MSHFlexGrid1.Cols - 1
    MSHFlexGrid1.Col = C
    MSHFlexGrid1.CellBackColor = vbGreen
Next

Call ClearFrame
'CmdGo.SetFocus
Frame4.Enabled = False

End Sub

Private Sub Form_Load()

Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   'Call GButtonLockAD(Me)
   Frame1.Enabled = False
   Frame2.Enabled = False
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If

Frame1.Enabled = True
Frame2.Enabled = True
'SaveCmd.Enabled = False

Call CheckAccess
Call Fill_MonthYear
Call CmdClear_Click
'Call Grid_Head
'Call Grid_Head1
'Call Fill_MstLessor
End Sub
Private Sub Fill_MonthYear()
For I = 1 To 12
   CmbMonth.AddItem MonthName(I), I - 1
Next

End Sub
Private Sub Grid_Head()
With MSHFlexGrid1
     .Clear
     .Rows = 2
     .Cols = 30 '29
     .AllowUserResizing = flexResizeBoth
     .Font.Size = 10
     .WordWrap = True
     .RowHeight(0) = 700

     .TextMatrix(0, 0) = "Select"
     .TextMatrix(0, 1) = "Location Name"
     .TextMatrix(0, 2) = "CMP Name"
     .TextMatrix(0, 3) = "Godown No"
     .TextMatrix(0, 4) = "Franchisee Name"
     .TextMatrix(0, 5) = "Commodity"
     .TextMatrix(0, 6) = "Exchange Type"
     .TextMatrix(0, 7) = "Txn Type"
     .TextMatrix(0, 8) = "Billing Month" '"Year Month"
     .TextMatrix(0, 9) = "Txn Date" '"CxTF Date"
     .TextMatrix(0, 10) = "CDTF Date" '.TextMatrix(0,20) = "CDTF Date"
     .TextMatrix(0, 11) = "No Of Bags" '"Balance Bags"
     .TextMatrix(0, 12) = "Weight in MT" '"Balance Weight"
     .TextMatrix(0, 13) = "Billing Cycle"
     .TextMatrix(0, 14) = "No of Cycles"
     .TextMatrix(0, 15) = "Franchisee Revenue Share"
     .TextMatrix(0, 16) = "Bagsize"
     .TextMatrix(0, 17) = "Franchisee Share Amount"
     .TextMatrix(0, 18) = "GSLID"
     .TextMatrix(0, 19) = "ST NAV Code"
     .TextMatrix(0, 20) = "VendorID"
     .TextMatrix(0, 21) = "CommodityID"
     .TextMatrix(0, 22) = "ExchangeTypeID"
     .TextMatrix(0, 23) = "LocationID"
     .TextMatrix(0, 24) = "SM_Prefix"
     .TextMatrix(0, 25) = "BillingCycleID"
     .TextMatrix(0, 26) = "Billing Unit"
     .TextMatrix(0, 27) = "CxTFNo"
     .TextMatrix(0, 28) = "Total Amount"
     .TextMatrix(0, 29) = "Total Insurance Cost"
     
     .ColWidth(0) = 800
     .ColWidth(1) = 2100
     .ColWidth(2) = 2100
     .ColWidth(3) = 1000
     .ColWidth(4) = 3500
     .ColWidth(5) = 1800
     .ColWidth(6) = 1600
     .ColWidth(7) = 800 '1200
     .ColWidth(8) = 1500
     .ColWidth(9) = 1500 '800
     .ColWidth(10) = 1300
     .ColWidth(11) = 1300
     .ColWidth(12) = 1200
     .ColWidth(13) = 800
     .ColWidth(14) = 800
     .ColWidth(15) = 1200
     .ColWidth(16) = 800 '1200
     .ColWidth(17) = 800
     .ColWidth(18) = 0
     .ColWidth(19) = 0
     .ColWidth(20) = 0
     .ColWidth(21) = 0
     .ColWidth(22) = 0
     .ColWidth(23) = 0
     .ColWidth(24) = 0
     .ColWidth(25) = 0
     .ColWidth(28) = 1200
     .ColWidth(29) = 1200
     
End With

End Sub
Private Sub Grid_Head1()
With MSHFlexGrid2
    .Rows = 2
    .Cols = 3
    .AllowUserResizing = flexResizeBoth
    .TextMatrix(0, 0) = "Vendor ID"
    .TextMatrix(0, 1) = "Vendor Name"
    .ColWidth(0) = 1200
    .ColWidth(1) = 2500
    
End With
End Sub
Private Sub Fill_MstVendor()
MSHFlexGrid2.Rows = 1
MSHFlexGrid2.Clear
Dim flg As Variant
ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid1.Rows + 1
For y = 1 To MSHFlexGrid1.Rows - 1
    If MSHFlexGrid1.TextMatrix(y, 0) = strChecked Then
       flg = "Y"
       For o = 1 To MSHFlexGrid2.Rows - 1
          If MSHFlexGrid2.TextMatrix(o, 0) = MSHFlexGrid1.TextMatrix(y, 20) Then
             flg = "N"
             Exit For
          End If
       Next
       If flg = "Y" Then
          I = MSHFlexGrid2.Rows
          MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
          MSHFlexGrid2.TextMatrix(I, 0) = MSHFlexGrid1.TextMatrix(y, 20)
          MSHFlexGrid2.TextMatrix(I, 1) = MSHFlexGrid1.TextMatrix(y, 4)
       End If
  End If
  ProgressBar1.Value = ProgressBar1.Value + 1
Next
ProgressBar1.Value = ProgressBar1.Max
End Sub
Private Sub CheckAccess()
CmdShowAll.Visible = False


'If LCase(Gen_UserRole) = "superadmin" Then
'   CmdShowAll.Visible = True
'End If
'If LCase(Gen_UserRole) = "admin" Then
'   CmdShowAll.Visible = True
'End If

If LCase(Gen_UserRole) = "power" Or LCase(Gen_UserRole) = "superadmin" Or LCase(Gen_UserRole) = "admin" Or LCase(Gen_UserRole) = "headcmg" Then
'   TxtCMPMName.Visible = False
   CmbVendor.Visible = True
   mUserTypeFlag = "power"
   CmdShowAll.Visible = True

ElseIf Gen_DesignationID = 6 Then
'   TxtCMPMName.Visible = True
'   TxtCMPMName.Text = Gen_UserName
'   CmbVendor.Visible = False
   mUserTypeFlag = "cmpm"
Else
   Frame1.Enabled = False
   Frame2.Enabled = False
End If
End Sub

Private Sub MSHFlexGrid1_Click()
With MSHFlexGrid1
    If .Col = 0 Then  '.Row = 1 And
       If .TextMatrix(.row, .Col) = "" Then Exit Sub
       'If (.TextMatrix(.Row, 13)) = "" Then
       .Col = 0
       If (.CellBackColor) <> vbRed And (.CellBackColor) <> vbYellow Then
          If .TextMatrix(.row, .Col) = strUnChecked Then
             .TextMatrix(.row, .Col) = strChecked
          Else
             .TextMatrix(.row, .Col) = strUnChecked
          End If
       Else
          If .CellBackColor = vbRed Then
             MsgBox "Franchisee rate card master not updated!!!"
          ElseIf .CellBackColor = vbYellow Then
             MsgBox "Godown Closed!!!"
          End If
       End If
    Else
        'If LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "admin" And LCase(Gen_UserRole) <> "superadmin" And LCase(Gen_UserRole) <> "head cmg" Then
        '   Exit Sub
        'End If
        If .TextMatrix(.RowSel, 1) = "" Then Exit Sub
        
        If .TextMatrix(.RowSel, 4) = "" Then Exit Sub
           
        If .TextMatrix(.RowSel, 26) = "F" And Val(.TextMatrix(.RowSel, 15)) > 0 And Val(.TextMatrix(.RowSel, 17)) <= 0 Then Exit Sub ' for fixed rate to any one GSL of godow
        If .CellBackColor = vbGreen Then Exit Sub
        Frame4.Enabled = True
        
        TxtLocationName.Text = .TextMatrix(.RowSel, 1)
        TxtCMPName.Text = .TextMatrix(.RowSel, 2)
        TxtGodownNo.Text = .TextMatrix(.RowSel, 3)
        TxtClientName.Text = .TextMatrix(.RowSel, 4)
        TxtCommodity.Text = .TextMatrix(.RowSel, 5)
        CmbBillingCycle.Text = .TextMatrix(.RowSel, 13)
        CmbBillingUnit.Text = .TextMatrix(.RowSel, 26)
        TxtPresentRate.Text = .TextMatrix(.RowSel, 15)
        mTFlagOnUpdate = .TextMatrix(.RowSel, 7)
        mTxnNo = .TextMatrix(.RowSel, 27)
        
        Call CmbBillingCycle_GotFocus
        Call CmbBillingCycle_LostFocus
        Call CmbBillingUnit_LostFocus
        
        If CmbBillingCycle.Text = "" Then
           CmbBillingCycle.SetFocus
        Else
           TxtNewRate.SetFocus
        End If
    End If
End With

End Sub
Private Sub SaveCmd_Click()

If MSHFlexGrid1.Rows = 2 Then
   MsgBox "No Record Present To Save !!!"
   Exit Sub
End If

SaveCmd.Enabled = False
ProgressBar1.Value = 0

tmp = "Select ME.EmployeeID,ME.EmployeeName from Mst_Employee ME"
tmp = tmp & " Inner Join Mst_Employee EMP on EMP.DirectReportingTo=ME.EmployeeID"
tmp = tmp & " Where EMP.EmployeeID = " & Gen_EmployeeID & ""
Call TmpTable

If TmpRs.RecordCount > 0 Then
   mDirectReportingTo = TmpRs!EmployeeID
Else
   MsgBox "Direct Reporting Not Available, You Can't Create Claim !!!"
   Exit Sub
End If
''

'If CDate(mExpenseFrom) < "2021-08-01" Then
'    Call Fill_MstVendor 'you will get Distinct Vendor from selected List
'
'    mGridMsg = "<tr  bgcolor=" & "#E9CFEC" & "  style=" & "font:bold" & " >"  '
'    For mMail = 0 To MSHFlexGrid1.Cols - 1
'        If mMail = 0 Then
'           mGridMsg = mGridMsg & " <td> Claim ID </td>"
'        Else
'           mGridMsg = mGridMsg & " <td> " & MSHFlexGrid1.TextMatrix(0, mMail) & "</td>"
'        End If
'    Next
'    mGridMsg = mGridMsg & "</tr>"
'
'
'    Call TableTxn_FranchiseeRent
'
'    ProgressBar1.Value = 0
'
'    Dim mVendorID_ As Double
'    ProgressBar1.Max = MSHFlexGrid2.Rows + 1
'    For I = 1 To MSHFlexGrid2.Rows - 1
'        mFlagChk = 0
'        If MSHFlexGrid2.TextMatrix(I, 0) = "" Then Exit For
'        mVendorID_ = MSHFlexGrid2.TextMatrix(I, 0)
'        Call GetClaimDetails(mVendorID_)
'        ProgressBar1.Value = ProgressBar1.Value + 1
'    Next
'    ProgressBar1.Value = ProgressBar1.Max
'    If MSHFlexGrid2.Rows > 1 Then
'       Call Gen_Email(mGridMsg)
'       MsgBox "Done !"
'       Call Grid_Head
'    Else
'       MsgBox "No record to save !!!"
'    End If
'Else
    mSelectedGSL = "(0"
    For j = 1 To MSHFlexGrid1.Rows - 1
        If MSHFlexGrid1.TextMatrix(j, 0) = strChecked Then
           mSelectedGSL = mSelectedGSL & "," & Val(MSHFlexGrid1.TextMatrix(j, 18))
        End If
    Next
    If mSelectedGSL <> "" Then
       mSelectedGSL = mSelectedGSL & ")"
       tmp = "Select Distinct FF_RateCardID,SM_Prefix from Txn_InvMonGSLDetail Where YearMonth = '" & CDate(mExpenseFrom) & "' ANd FF_RateCardID is not null And Mst_GSL_ID in " & mSelectedGSL & " "
       Call Tmp0Table
       
       If TmpRs0.RecordCount > 0 Then
          For p = 1 To TmpRs0.RecordCount
              mSM_Prefix = TmpRs0!SM_Prefix
              tmp = "Select Distinct GodownID from Txn_InvMonGSLDetail Where YearMonth = '" & CDate(mExpenseFrom) & "' ANd SM_Prefix = '" & mSM_Prefix & "' And FF_RateCardID = " & TmpRs0!FF_RateCardID & " And Mst_GSL_ID in " & mSelectedGSL & ""
              Call Tmp1Table
              For a = 1 To TmpRs1.RecordCount
                  tmp = "Select Distinct VendorID,VendorShare from Mst_FF_RateCard_GodownVendorMapping Where FF_RateCardID  = " & TmpRs0!FF_RateCardID & " And GodownID = " & TmpRs1!GodownID & " "
                  Call Tmp2Table
                  For r = 1 To TmpRs2.RecordCount
                      If CDate(mExpenseFrom) < "2022-02-01" Then
                         tmp = "Select YearMonth,SM_Prefix,Mst_GSL_ID,InvoiceAmount 'InvoiceAmount',FranchiseeRate,FranchInvoiceAmount,MonG_UID from Txn_InvMonGSLDetail Where YearMonth = '" & CDate(mExpenseFrom) & "' ANd SM_Prefix = '" & mSM_Prefix & "' And FF_RateCardID = " & TmpRs0!FF_RateCardID & " And GodownID = " & TmpRs1!GodownID & " And InvoiceAmount  is not null And FranchInvoiceAmount <> 0 And Mst_GSL_ID in " & mSelectedGSL & " "
                      Else
                         tmp = "Select YearMonth,SM_Prefix,Mst_GSL_ID,InvoiceAmtForFranch 'InvoiceAmount',FranchiseeRate,FranchInvoiceAmount,MonG_UID from Txn_InvMonGSLDetail Where YearMonth = '" & CDate(mExpenseFrom) & "' ANd SM_Prefix = '" & mSM_Prefix & "' And FF_RateCardID = " & TmpRs0!FF_RateCardID & " And GodownID = " & TmpRs1!GodownID & " And InvoiceAmtForFranch  is not null And FranchInvoiceAmount <> 0 And Mst_GSL_ID in " & mSelectedGSL & " "
                      End If
                      Call Tmp3Table
                      For ag = 1 To TmpRs3.RecordCount
                          
                          tmp = "Select * from Txn_MonthlyFranchiseeRentShare Where [ID] = 0"
                          Call Tmp4Table
                          
                          TmpRs4.AddNew
                          TmpRs4!YearMonth = TmpRs3!YearMonth
                          TmpRs4!SM_Prefix = TmpRs3!SM_Prefix
                          TmpRs4!Mst_GSL_ID = TmpRs3!Mst_GSL_ID
                          TmpRs4!InvoiceAmount = TmpRs3!InvoiceAmount
                          TmpRs4!FranchiseeRate = TmpRs3!FranchiseeRate
                          TmpRs4!FranchInvoiceAmount = TmpRs3!FranchInvoiceAmount
                          TmpRs4!VendorID = TmpRs2!VendorID
                          TmpRs4!VendorShare = TmpRs2!VendorShare
                          TmpRs4!ShareAmount = Format((TmpRs4!FranchInvoiceAmount * TmpRs4!VendorShare / 100), "#######0.00")
                          TmpRs4!MonG_UID = TmpRs3!MonG_UID
                          TmpRs4!G_UID = GetGUID
                          TmpRs4!CreatedBy = Gen_UserLoginID
                          TmpRs4!CreatedDate = Now
                          TmpRs4.Update
                          TmpRs3.MoveNext
                      Next ag
                      TmpRs2.MoveNext
                  Next r
                  TmpRs1.MoveNext
              Next a
              TmpRs0.MoveNext
          Next p
       End If
    
       tmp = "Select  YearMonth,SM_Prefix,VendorID,Sum(ShareAmount) 'ShareAmount' From Txn_MonthlyFranchiseeRentShare Where YearMonth = '" & CDate(mExpenseFrom) & "' And Mst_GSL_ID in " & mSelectedGSL & " And ClaimID is null Group by YearMonth,SM_Prefix,VendorID"
       Call Tmp1Table
       
       For p = 1 To TmpRs1.RecordCount
           
           Call GetClaimDetailsNewProcess(TmpRs1!VendorID, TmpRs1!ShareAmount, TmpRs1!SM_Prefix)
           TmpRs1.MoveNext
       Next
    
       MsgBox "Done!!!"
       Call Grid_Head
    
    End If
    


'End If




ChkCheckAll.Value = 0

End Sub

Private Sub GetClaimDetails(mVendorID1_ As Double)

mExpenseTo = Format(CmbYear.Text & "-" & mMonth & "-" & mTotalDays, "yyyy-mm-dd")

RsTxn_FranchiseeRent.AddNew
'mClaimID = GetMaxClaimID
'RsTxn_FranchiseeRent!ClaimID = mClaimID
RsTxn_FranchiseeRent!ClaimDate = Date
RsTxn_FranchiseeRent!VendorID = mVendorID1_ ' Val(MSHFlexGrid1.TextMatrix(j, 11))
RsTxn_FranchiseeRent!Amount = 0 'Val(MSHFlexGrid1.TextMatrix(j, 8))
RsTxn_FranchiseeRent!ExpenseFrom = mExpenseFrom
RsTxn_FranchiseeRent!ExpenseTo = mExpenseTo
RsTxn_FranchiseeRent!LocationID = mLocationID
RsTxn_FranchiseeRent!ForwardToID = mDirectReportingTo ' mVendorID
RsTxn_FranchiseeRent!Flag = "P"
RsTxn_FranchiseeRent!BatchID = 0
RsTxn_FranchiseeRent!G_UID = GetGUID
mG_UID = RsTxn_FranchiseeRent!G_UID
RsTxn_FranchiseeRent!CreatedBy = Gen_UserLoginID
RsTxn_FranchiseeRent!CreatedDate = Now
RsTxn_FranchiseeRent!Remarks = "Being Franchisee Rent for the month of " & CmbMonth.Text & " - " & CmbYear.Text
RsTxn_FranchiseeRent.Update

ProgressBar2.Value = 0
ProgressBar2.Max = MSHFlexGrid1.Rows + 1


tmp = "Select ClaimID From Txn_FranchiseeRent Where G_UID= '" & mG_UID & "' "
Call TmpTable

If TmpRs.RecordCount > 0 Then
   mClaimID = TmpRs!ClaimID
End If

For j = 1 To MSHFlexGrid1.Rows - 1
    
    If MSHFlexGrid1.TextMatrix(j, 0) = strChecked And Val(MSHFlexGrid1.TextMatrix(j, 20)) = mVendorID1_ Then
       mFlagChk = 1
    Else
       mFlagChk = 0
    End If
    If mFlagChk = 1 Then
       
       RsTxn_FranchiseeRent!Amount = RsTxn_FranchiseeRent!Amount + Val(MSHFlexGrid1.TextMatrix(j, 17))
       RsTxn_FranchiseeRent.Update
       
       mGridMsg = mGridMsg & "<tr>"
       For mMail = 0 To MSHFlexGrid1.Cols - 1
           If mMail = 0 Then
              mGridMsg = mGridMsg & " <td> " & mClaimID & "</td>"
           Else
              mGridMsg = mGridMsg & " <td> " & IIf(MSHFlexGrid1.TextMatrix(j, mMail) = "", "&nbsp", MSHFlexGrid1.TextMatrix(j, mMail)) & "</td>"
           End If
       Next
       mGridMsg = mGridMsg & "</tr>"
       
       tmp = "SET IMPLICIT_TRANSACTIONS ON"
       Call Tmp1Table
       
       
       If MSHFlexGrid1.TextMatrix(j, 7) = "" Then
       
          tmp = " Set DeadLock_Priority Low Update Txn_InvMonGSLDetail Set FranchClaimID = " & mClaimID & ", " & _
          "InvoiceDate = '" & mExpenseTo & "'  Where GodownID = " & MSHFlexGrid1.TextMatrix(j, 18) & " And SM_Prefix = '" & MSHFlexGrid1.TextMatrix(j, 24) & "' And YearMonth = '" & MSHFlexGrid1.TextMatrix(j, 8) & "' And TFlag = '" & MSHFlexGrid1.TextMatrix(j, 7) & "' And ISINID='" & "V" & mVendorID1_ & "'"
          
       Else
       
          tmp = " Set DeadLock_Priority Low  Update Txn_InvMonGSLDetail Set FranchClaimID = " & mClaimID & ", " & _
          "InvoiceDate = '" & mExpenseTo & "'  Where Mst_GSL_ID = " & MSHFlexGrid1.TextMatrix(j, 18) & " And SM_Prefix = '" & MSHFlexGrid1.TextMatrix(j, 24) & "' And YearMonth = '" & MSHFlexGrid1.TextMatrix(j, 8) & "' And TFlag = '" & MSHFlexGrid1.TextMatrix(j, 7) & "'" ' And ISINID='" & "V" & mVendorID1_ & "'"
       End If
       
       Call Tmp1Table
          
    End If
    ProgressBar2.Value = ProgressBar2.Value + 1
Next
ProgressBar2.Value = ProgressBar2.Max

End Sub
Private Sub GetClaimDetailsNewProcess(mVendorID1_ As Double, mClaimAmount_ As Double, mSM_Prefix_ As Variant)

mExpenseTo = Format(CmbYear.Text & "-" & mMonth & "-" & mTotalDays, "yyyy-mm-dd")
Call TableTxn_FranchiseeRent("Where ClaimID = 0")
RsTxn_FranchiseeRent.AddNew
'mClaimID = GetMaxClaimID
'RsTxn_FranchiseeRent!ClaimID = mClaimID
RsTxn_FranchiseeRent!ClaimDate = Date
RsTxn_FranchiseeRent!VendorID = mVendorID1_ ' Val(MSHFlexGrid1.TextMatrix(j, 11))
RsTxn_FranchiseeRent!Amount = Val(mClaimAmount_) 'Val(MSHFlexGrid1.TextMatrix(j, 8))
RsTxn_FranchiseeRent!ExpenseFrom = mExpenseFrom
RsTxn_FranchiseeRent!ExpenseTo = mExpenseTo
RsTxn_FranchiseeRent!LocationID = mLocationID
RsTxn_FranchiseeRent!ForwardToID = mDirectReportingTo ' mVendorID
RsTxn_FranchiseeRent!Flag = "P"
RsTxn_FranchiseeRent!BatchID = 0
RsTxn_FranchiseeRent!G_UID = GetGUID
mG_UID = RsTxn_FranchiseeRent!G_UID
RsTxn_FranchiseeRent!CreatedBy = Gen_UserLoginID
RsTxn_FranchiseeRent!CreatedDate = Now
RsTxn_FranchiseeRent!Remarks = "Being Franchisee Rent for the month of " & CmbMonth.Text & " - " & CmbYear.Text
RsTxn_FranchiseeRent.Update

ProgressBar2.Value = 0
ProgressBar2.Max = MSHFlexGrid1.Rows + 1


tmp = "Select ClaimID From Txn_FranchiseeRent Where G_UID= '" & mG_UID & "' "
Call TmpTable

If TmpRs.RecordCount > 0 Then
   mClaimID = TmpRs!ClaimID
   tmp = "Update Txn_MonthlyFranchiseeRentShare Set ClaimID = " & mClaimID & " Where YearMonth = '" & CDate(mExpenseFrom) & "' And VendorID = " & mVendorID1_ & " And SM_Prefix = '" & mSM_Prefix_ & "' "
   Call Tmp2Table
End If
ProgressBar2.Value = ProgressBar2.Max

End Sub


Private Function GetMaxClaimID()
Dim Retval As Double
tmp = "Select max(ClaimID) from Txn_FranchiseeRent"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetMaxClaimID = Retval
End Function

Private Function Gen_Email(mMailData_ As String)
Dim TOADD, TOCC As Variant
Dim mCMPMName, mEmpName_ As Variant

mCMPMName = GetEmployeeNonName(Gen_EmployeeID)

mMsg = ""
mMsg = mMsg & "<tr style=" & "font:bold" & " > <td Colspan=2><C><b>Godown Rent Claim (From Store-Man)</b></C></td><tr>"
mMsg = mMsg & " <tr><td><b>CMPM Name</b></td><td>" & mCMPMName & "</td></tr>"
mMsg = mMsg & " <tr><td><b>Rent From</b> </td><td>" & Format(mExpenseFrom, "dd-MMM-yyyy") & "</td></tr>"
mMsg = mMsg & " <tr><td><b>Rent To</b></td><td>" & Format(mExpenseTo, "dd-MMM-yyyy") & "</td></tr>"
mEmpName_ = GetEmployeeNonName(mDirectReportingTo)
mMsg = mMsg & " <tr><td><b>Forwarded To</b></td><td>" & mEmpName_ & "</td></tr>"
mMsg = mMsg & "<tr> <td Colspan=2><b>Details</b></td><tr>"

mMsg = mMsg & "<tr> <td Colspan=2><table border=1>" & mGridMsg & "</table></td> </tr> "
mMsgSubject = "Franchisee Rent Claim"

mForwardToEmailID = GetEmployeeEmailID(mDirectReportingTo)


'tmp = "Select ME.EmployeeID,ME.EmployeeName,ME.EmployeeEmailID from Mst_Employee ME"
'tmp = tmp & " Inner Join Mst_Employee EMP on EMP.DirectReportingTo=ME.EmployeeID"
'tmp = tmp & " Where EMP.EmployeeID = " & mDirectReportingTo & ""
'Call TmpTable
'If TmpRs.RecordCount > 0 Then
'   mForwardToEmailID = mForwardToEmailID & "," & TmpRs!EmployeeEmailID
'End If

TOADD = mForwardToEmailID
TOCC = Gen_UserEmailID
Call SendMail(TOADD, TOCC, mMsgSubject, mMsg)

End Function
Private Function SendMail(ToADD_ As Variant, TOCC_ As Variant, MsgSubject_ As Variant, Message1_ As Variant)
Dim htmlStringhead As String
Dim htmlStringtail As String
Set oSMTP = New OSSMTP.SMTPSession

servertxt = "mail.ncml.com"

htmlStringhead = "<html><body><font size= 3 face=Zurich BT color = Blue><table border = 1 cellpadding=1 cellspacing =1>"
htmlStringtail = "</table></Font></body></html>"
   
'With oSMTP
'  'connection
'  .Server = servertxt
'  'message
'  .MailFrom = TOCC_ '"mayur.d@ncml.com"
'  If IsNull(Trim(ToADD_)) = True Or Trim(ToADD_) = "" Then
'     ToADD_ = "bhagwat.k@ncml.com"
'  End If
'  .SendTo = ToADD_
'  'TOCC_ = TOCC_ & ",cmg@ncml.com"
'  If ToADD_ <> TOCC_ Then
'     .CC = TOCC_
'  End If
''  .BCC = "mayur.d@ncml.com"
'  .MessageSubject = MsgSubject_ & " (from Store-Man)" '"Confidential Mail"
'  .MessageHTML = htmlStringhead & Message1_ & htmlStringtail
'  .TimeStamp = Now()
'  If LCase(Gen_EmailConfigFlag) = "y" Then
'     .SendEmail
'  End If
'End With

Dim objSMTP
Set objSMTP = CreateObject("MailBee.SMTP")
objSMTP.LicenseKey = "MBC900-5F470FAD67-D8E202A2E22927335F24F5A667BFD196"
objSMTP.ServerName = Gen_ReadMailServerIPINI("MailIP") ' = "mail.ncml.com" '"mail.ncml.com"
objSMTP.FromAddr = TOCC_
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

Private Sub CmbCMPName_GotFocus()
If CmbLocationID.Text = "" Then
   MsgBox "Please Select Location !!"
   CmbCMPName.SetFocus
   Exit Sub
End If

tmp = CmbCMPName.Text
Call TableMst_CMP(" where LocationID=" & mLocationID & " ")
CmbCMPName.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "First Select Location!!!! "
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
   Exit Sub
End If
RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPName = '" & CmbCMPName.Text & "'"
If RsMst_CMP.EOF Then
   MsgBox "Invalid selection "
   CmbCMPName.SetFocus
   Exit Sub
End If
mCMPID = RsMst_CMP!CMPID
End Sub


Private Sub CmbSGodown_GotFocus()
If CmbCMPName.Text = "" Then
   MsgBox "Please Select CMP  !!! "
   CmbSGodown.SetFocus
   Exit Sub
End If
tmp = CmbSGodown.Text
Call TableMst_Godown("Where CMPID = " & mCMPID)
CmbSGodown.Clear
If RsMst_Godown.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Godown.RecordCount
    CmbSGodown.AddItem RsMst_Godown!GodownNo
    RsMst_Godown.MoveNext
Next
CmbSGodown.Text = tmp
End Sub

Private Sub CmbSGodown_LostFocus()
If CmbSGodown.Text = "" Then
   Exit Sub
End If
RsMst_Godown.MoveFirst
RsMst_Godown.Find "GodownNo = '" & CmbSGodown.Text & "'"
If RsMst_Godown.EOF Then
   MsgBox "Invalid selection "
   CmbSGodown.SetFocus
   Exit Sub
End If
mGodownID = RsMst_Godown!GodownID
End Sub


Private Sub CmbBillingCycle_GotFocus()
tmp = CmbBillingCycle.Text
Call TableMst_BillingCycle
CmbBillingCycle.Clear
If RsMst_BillingCycle.RecordCount = 0 Then
  MsgBox "No Record found !!!! "
  Exit Sub
End If
For I = 1 To RsMst_BillingCycle.RecordCount
   CmbBillingCycle.AddItem RsMst_BillingCycle!BillingCycle
   RsMst_BillingCycle.MoveNext
Next
CmbBillingCycle.Text = tmp
End Sub

Private Sub CmbBillingCycle_LostFocus()
If CmbBillingCycle.Text = "" Then
   mBillingCycleID = Null
   Exit Sub
End If
RsMst_BillingCycle.MoveFirst
RsMst_BillingCycle.Find "BillingCycle = '" & CmbBillingCycle.Text & "'"
If RsMst_BillingCycle.EOF Then
   MsgBox "Invalid selection "
   CmbBillingCycle.SetFocus
   Exit Sub
End If
mBillingCycleID = RsMst_BillingCycle!BillingCycleID
mTxtTotDays = RsMst_BillingCycle!NoOfDays

'Call CalNoOfCylce(
End Sub

Private Sub CmbBillingUnit_LostFocus()
If CmbBillingUnit.Text = "" Then
   Exit Sub
End If

If LCase(CmbBillingUnit.Text) <> "bags" And LCase(CmbBillingUnit.Text) <> "quantity" And LCase(CmbBillingUnit.Text) <> "fixed" And LCase(CmbBillingUnit.Text) <> "percentage" Then
   MsgBox "Invalid Selection!!!"
   CmbBillingUnit.SetFocus
   Exit Sub
End If

If LCase(CmbBillingUnit.Text) = "bags" Then
   mFlag = "B"
ElseIf LCase(CmbBillingUnit.Text) = "quantity" Then
   mFlag = "Q"
ElseIf LCase(CmbBillingUnit.Text) = "percentage" Then
   mFlag = "P"
ElseIf LCase(CmbBillingUnit.Text) = "fixed" Then
   mFlag = "F"
Else
   mFlag = Null
End If
End Sub


Private Sub TxtNewRate_LostFocus()
If TxtNewRate.Text = "" Then Exit Sub

If IsNumeric(TxtNewRate.Text) = False Then
   MsgBox "Invalid Numeric Format."
   TxtNewRate.SetFocus
   Exit Sub
End If

If LCase(Gen_UserRole) = "account" Then
   If Val(TxtNewRate) < Val(TxtPresentRate) Then
      MsgBox "You can not give rate less then present rate!!!"
      TxtNewRate.SetFocus
      Exit Sub
   End If
End If
End Sub
Private Sub UpdateTxn_InvMonGSL(mROw_ As Variant)
Dim mChoice As Variant
Dim mNoofDays, mNoOfCycle As Variant
mChoice = ""
If MSHFlexGrid1.TextMatrix(mROw_, 7) <> "" Then
   mChoice = MSHFlexGrid1.TextMatrix(mROw_, 7) & Left(CmbBillingCycle.Text, 1)
End If
mSM_Prefix_ = MSHFlexGrid1.TextMatrix(mROw_, 24)
''Find No Of Days
mNoofDays = GetBillingCycleDays(mBillingCycleID)

'' Year Month
'mFromDate = "01-" & CmbMonth.Text & "-" & CmbYear.Text
mFromDate = mExpenseFrom
mToDate = mExpenseTo

'Transaction Date To Calculate No OF Cycle
mDate_ = Format(MSHFlexGrid1.TextMatrix(mROw_, 9), "yyyy-mm-dd")

tmp = "SET IMPLICIT_TRANSACTIONS ON"
Call TmpTable


tmp1 = " Set DeadLock_Priority Low  Update Txn_InvMonGslDetail Set OldFranchRate=FranchiseeRate,OldFranchBillingCycleID=FranchBillingCycleID,OldFranchBillingUnit=FranchBillingUnit,"
tmp1 = tmp1 & " FranchiseeRate=" & Val(MSHFlexGrid1.TextMatrix(mROw_, 15)) & ",FranchBillingCycleID=" & mBillingCycleID & " ,FranchBillingUnit='" & mFlag & "',UpdatedDate = Getdate(), UpdatedBy = " & Gen_UserLoginID & ","
tmp = ""
Select Case (mChoice)

Case "OM"
   mNoOfCycle = 1
   tmp = "  FranchNoofDays = 1 Where YearMonth = '" & mFromDate & "' And TFlag = 'O'"
   tmp = tmp & " And SM_Prefix = '" & mSM_Prefix_ & "' "
Case "OF"
   mNoOfCycle = 2
   tmp = " FranchNoofDays = 2 Where YearMonth = '" & mFromDate & "' And TFlag = 'O'"
   tmp = tmp & " And  And SM_Prefix = '" & mSM_Prefix_ & "' "
Case "OW"
   mNoOfCycle = 4
   tmp = " FranchNoofDays = 4 Where YearMonth = '" & mFromDate & "' And TFlag = 'O'"
   tmp = tmp & " And SM_Prefix = '" & mSM_Prefix_ & "' "
Case "OD"
   mNoOfCycle = mTotalDays
   tmp = " FranchNoofDays = " & mTotalDays & " Where YearMonth = '" & mFromDate & "' And TFlag = 'O'"
   tmp = tmp & " And SM_Prefix = '" & mSM_Prefix_ & "' "
Case "DM" 'DM & WM
   mNoOfCycle = 1
   tmp = " FranchNoofDays = 1 Where YearMonth = '" & mFromDate & "' And TFlag = 'D'"
   tmp = tmp & " And SM_Prefix = '" & mSM_Prefix_ & "' "
Case "WM"
   mNoOfCycle = 1
   tmp = " FranchNoofDays = 1 Where YearMonth = '" & mFromDate & "'  And TFlag = 'W'"
   tmp = tmp & " And SM_Prefix = '" & mSM_Prefix_ & "' "
Case "DD" 'DD & WD
    mNoOfCycle = mTotalDays - Format(mDate_, "dd") + 1 'Day(" & mDate_ & ") + 1
    tmp = " FranchNoofDays = " & mTotalDays & " - Day('" & mDate_ & "') + 1 Where YearMonth = '" & mFromDate & "' And TFlag = 'D'  And SM_Prefix = '" & mSM_Prefix_ & "' "
Case "WD"
   mNoOfCycle = Format(mDate_, "dd")
   tmp = " FranchNoofDays = Day('" & mDate_ & "') Where YearMonth = '" & mFromDate & "'  And TFlag = 'W' And SM_Prefix = '" & mSM_Prefix_ & "' "
Case "DW"
   tmp = "Select Case Floor((4 - (Convert(Numeric,Day('" & mDate_ & "'))/" & mNoofDays & ") + 1))  When 0 then 1 else Floor((4 - (Convert(Numeric,Day('" & mDate_ & "'))/" & mNoofDays & ") + 1)) End As NoOfCycle"
   Call TmpTable
   mNoOfCycle = TmpRs!NoofCycle
   tmp = " FranchNoofDays =  Case Floor((4 - (Convert(Numeric,Day('" & mDate_ & "'))/" & mNoofDays & ") + 1))  When 0 then 1 else Floor((4 - (Convert(Numeric,Day('" & mDate_ & "'))/" & mNoofDays & ") + 1)) End "
   tmp = tmp & " Where Txn_InvMonGslDetail.YearMonth = '" & mFromDate & "' "
   tmp = tmp & " And Txn_InvMonGslDetail.TFlag = 'D' And SM_Prefix = '" & mSM_Prefix_ & "' "
Case "DF"
   tmp = "Select Case Floor((4 - Round(Convert(Decimal,Day('" & mDate_ & "'))/" & mNoofDays & ",1)) - 1) When 0 then 1 else Floor((4 - Round(Convert(Decimal,Day('" & mDate_ & "'))/" & mNoofDays & ",1)) - 1) End As NoOfCycle"
   Call TmpTable
   mNoOfCycle = TmpRs!NoofCycle
   tmp = " FranchNoofDays = Case Floor((4 - Round(Convert(Decimal,Day('" & mDate_ & "'))/" & mNoofDays & ",1)) - 1) When 0 then 1 else Floor((4 - Round(Convert(Decimal,Day('" & mDate_ & "'))/" & mNoofDays & ",1)) - 1) End "
   tmp = tmp & " Where Txn_InvMonGslDetail.YearMonth = '" & mFromDate & "' "
   tmp = tmp & " And Txn_InvMonGslDetail.TFlag = 'D' And SM_Prefix = '" & mSM_Prefix_ & "' "
Case "WW"
   tmp = "Select Case CEILING(((Round(Convert(Numeric,Day('" & mDate_ & "'))/" & mNoofDays & ",1)))) When 5 Then 4 Else CEILING(((Round(Convert(Numeric,Day('" & mDate_ & "'))/" & mNoofDays & ",1)))) End  As NoOfCycle"
   Call TmpTable
   mNoOfCycle = TmpRs!NoofCycle
   tmp = " FranchNoofDays = Case CEILING(((Round(Convert(Numeric,Day('" & mDate_ & "'))/" & mNoofDays & ",1)))) When 5 Then 4 Else CEILING(((Round(Convert(Numeric,Day('" & mDate_ & "'))/" & mNoofDays & ",1)))) End "
   tmp = tmp & " Where Txn_InvMonGslDetail.YearMonth = '" & mFromDate & "' "
   tmp = tmp & " And Txn_InvMonGslDetail.TFlag = 'W' And SM_Prefix = '" & mSM_Prefix_ & "' "
Case "WF"
   tmp = "Select Case CEILING(((Round(Convert(Numeric,Day('" & mDate_ & "'))/" & mNoofDays & ",1)))) When 3 Then 2 Else CEILING(((Round(Convert(Numeric,Day('" & mDate_ & "'))/" & mNoofDays & ",1)))) End  As NoOfCycle"
   Call TmpTable
   mNoOfCycle = TmpRs!NoofCycle
   
   tmp = " FranchNoofDays = Case CEILING(((Round(Convert(Numeric,Day('" & mDate_ & "'))/" & mNoofDays & ",1)))) When 3 Then 2 Else CEILING(((Round(Convert(Numeric,Day('" & mDate_ & "'))/" & mNoofDays & ",1)))) End "
   tmp = tmp & " Where Txn_InvMonGslDetail.YearMonth = '" & mFromDate & "' "
   tmp = tmp & " And Txn_InvMonGslDetail.TFlag = 'W' And SM_Prefix = '" & mSM_Prefix_ & "' "
Case ""
   'Exit Sub
End Select

If tmp <> "" Then
   tmp = tmp & " And Mst_GSL_ID=" & Val(MSHFlexGrid1.TextMatrix(mROw_, 18)) & ""
   If LCase(mTFlagOnUpdate) = "w" Or LCase(mTFlagOnUpdate) = "d" Then
      tmp = tmp & " And CxTF_DetailsID =" & Val(MSHFlexGrid1.TextMatrix(mROw_, 27)) & ""
   End If
   tmp = tmp1 & tmp
   Call TmpTable

Else
   tmp = "SET IMPLICIT_TRANSACTIONS ON"
   Call TmpTable
  
  
  tmp1 = " Set DeadLock_Priority Low  Update Txn_InvMonGslDetail Set OldFranchRate=FranchiseeRate,OldFranchBillingCycleID=FranchBillingCycleID,OldFranchBillingUnit=FranchBillingUnit,"
  tmp1 = tmp1 & " FranchiseeRate=" & Val(MSHFlexGrid1.TextMatrix(mROw_, 15)) & ",FranchBillingCycleID=" & mBillingCycleID & " ,FranchBillingUnit='" & mFlag & "',UpdatedDate = Getdate(), UpdatedBy = '" & Gen_UserLoginID & "' "
  
  tmp = tmp & " Where YearMonth = '" & mFromDate & "' And GodownID=" & Val(MSHFlexGrid1.TextMatrix(mROw_, 18)) & " And Mst_GSL_ID = 0 "
  tmp = tmp1 & tmp
  
  Call TmpTable
End If

'Op Bal and Monthly Cycle** OM
'Op Bal and Forthnightly Cycle ** OF
'Op Bal and Weekly Cycle ** OW
'Op Bal and Daily Cycle ** OD
'Depo and Withd and Monthly Cycle ** DWM
'Depo and Withd and Daily Cycle ** DWD
'Deposit & Weekly  ** DW
'Deposit & FortNightly  ** DF
'Withdrawal & weekly ** WW
'Withdrawal & Fortnightly ** WF
'Updating Invoice Amount And Invoice Date

tmp = "SET IMPLICIT_TRANSACTIONS ON"
Call TmpTable


If LCase(MSHFlexGrid1.TextMatrix(mROw_, 26)) = "bags" Then

   mNoofBagsNQuantity = MSHFlexGrid1.TextMatrix(mROw_, 11)
   
   tmp = " Set DeadLock_Priority Low  Update Txn_InvMonGslDetail Set FranchInvoiceAmount = BalanceBags * FranchNoOfDays * FranchiseeRate "
   tmp = tmp & " Where YearMonth = '" & mFromDate & "'"
   tmp = tmp & " And FranchBillingUnit = 'B' And SM_Prefix = '" & mSM_Prefix_ & "' "
   tmp = tmp & " And Mst_GSL_ID=" & Val(MSHFlexGrid1.TextMatrix(mROw_, 18)) & ""
   
   Call TmpTable
   
ElseIf LCase(MSHFlexGrid1.TextMatrix(mROw_, 26)) = "quantity" Then

   mNoofBagsNQuantity = MSHFlexGrid1.TextMatrix(mROw_, 12)
   
   tmp = " Set DeadLock_Priority Low  Update Txn_InvMonGslDetail Set FranchInvoiceAmount = BalanceWeight * FranchNoOfDays * FranchiseeRate "
   tmp = tmp & " Where YearMonth = '" & mFromDate & "'"
   tmp = tmp & " And FranchBillingUnit = 'Q' And SM_Prefix = '" & mSM_Prefix_ & "' "
   tmp = tmp & " And Mst_GSL_ID=" & Val(MSHFlexGrid1.TextMatrix(mROw_, 18)) & ""
   
   Call TmpTable
   
ElseIf LCase(MSHFlexGrid1.TextMatrix(mROw_, 26)) = "fixed" Then
    mNoofBagsNQuantity = 1
    
   tmp = " Set DeadLock_Priority Low Update Txn_InvMonGSLDetail Set FranchInvoiceAmount = FranchiseeRate ,InvoiceDate =  '" & mToDate & "'"
   tmp = tmp & " Where YearMonth = '" & mFromDate & "' And FranchBillingUnit = 'F'  And GodownID = " & Val(MSHFlexGrid1.TextMatrix(mROw_, 18)) & " And Mst_GSL_ID = 0 And SM_Prefix = '" & mSM_Prefix_ & "'"
        
   Call TmpTable
    
ElseIf LCase(MSHFlexGrid1.TextMatrix(mROw_, 26)) = "percentage" Then
        
    tmp = " Set DeadLock_Priority Low  Update Txn_InvMonGSLDetail Set FranchInvoiceAmount =  (InvoiceAmount)* FranchiseeRate /100 ,InvoiceDate =  '" & mToDate & "' "
    tmp = tmp & " Where YearMonth = '" & mFromDate & "'"
    tmp = tmp & " And FranchBillingUnit = 'P' And SM_Prefix = '" & mSM_Prefix_ & "'"
    tmp = tmp & " And Mst_GSL_ID=" & Val(MSHFlexGrid1.TextMatrix(mROw_, 18)) & ""
   
    Call TmpTable
   
End If

'Invoice Amount
If LCase(MSHFlexGrid1.TextMatrix(mROw_, 26)) <> "fixed" And LCase(MSHFlexGrid1.TextMatrix(mROw_, 26)) <> "percentage" Then
   MSHFlexGrid1.TextMatrix(mROw_, 17) = mNoofBagsNQuantity * Val(MSHFlexGrid1.TextMatrix(mROw_, 15)) * mNoOfCycle
ElseIf LCase(MSHFlexGrid1.TextMatrix(mROw_, 26)) = "fixed" Then
   MSHFlexGrid1.TextMatrix(mROw_, 17) = Val(MSHFlexGrid1.TextMatrix(mROw_, 15))
ElseIf LCase(MSHFlexGrid1.TextMatrix(mROw_, 26)) = "percentage" Then
   MSHFlexGrid1.TextMatrix(mROw_, 17) = Val(MSHFlexGrid1.TextMatrix(mROw_, 28)) * Val(MSHFlexGrid1.TextMatrix(mROw_, 15) / 100)
End If
'Condition for below qry updated by Parag Sir
'tmp = " Update Txn_InvMonGslDetail Set InvoiceAmount = 0,InvoiceDate = Null Where YearMonth = '" & mFromDate & "' And Txn_InvMonGslDetail.ExchangeTypeID = 1 And TFlag = 'W' And (ContractID = 0 Or ContractID is Null) And Month('" & mDate_ & "') = Month(CDTFDate) And SM_Prefix = '" & mSM_Prefix_ & "' "
'Call TmpTable

MSHFlexGrid1.TextMatrix(mROw_, 14) = mNoOfCycle

End Sub

Private Sub ClearFrame()
TxtLocationName.Text = ""
TxtCMPName.Text = ""
TxtClientName.Text = ""
TxtCommodity.Text = ""
CmbBillingCycle.Text = ""
CmbBillingUnit.Text = ""
TxtPresentRate.Text = ""
TxtNewRate.Text = ""
TxtGodownNo.Text = ""
End Sub

''Private Function CheckGSLExist() As Boolean
''Dim Retval As Boolean
''Retval = True
''tmp = "Select GSLID from Txn_InvMonGSLDetail TINV"
''tmp = tmp & " Inner Join Mst_Godown MG On TINV.GodownID = MG.GodownID"
''tmp = tmp & " Where TINV.YearMonth = '" & mExpenseFrom & "' And SM_Prefix='" & mSM_Prefix & "' And MG.UnitTypeID in (2,6)"
''tmp = tmp & " And TINV.FF_RateCardID is not null  And (ISINID='V' + '" & mVendorID & "' Or TINV.TFLag <> '') "
''Call TmpTable
''If TmpRs.RecordCount > 0 Then
''   Retval = False
''End If
''CheckGSLExist = Retval
''End Function


