VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmMst_InspectionExpenseType 
   Caption         =   "Inspection Expense Type"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   2625
      Left            =   240
      TabIndex        =   41
      Top             =   6600
      Visible         =   0   'False
      Width           =   14895
      Begin VB.ComboBox CmbSType 
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
         ItemData        =   "FrmMst_InspectionExpenseType.frx":0000
         Left            =   10440
         List            =   "FrmMst_InspectionExpenseType.frx":000A
         Sorted          =   -1  'True
         TabIndex        =   25
         Top             =   450
         Width           =   2220
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   2400
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
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
         Left            =   13770
         TabIndex        =   27
         Top             =   460
         Width           =   1020
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
         Left            =   12735
         TabIndex        =   26
         Top             =   460
         Width           =   1020
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
         Height          =   360
         Left            =   75
         MaxLength       =   50
         TabIndex        =   22
         Top             =   450
         Width           =   5175
      End
      Begin VB.ComboBox CmbSCalculationOn 
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
         ItemData        =   "FrmMst_InspectionExpenseType.frx":0020
         Left            =   5400
         List            =   "FrmMst_InspectionExpenseType.frx":0036
         Sorted          =   -1  'True
         TabIndex        =   23
         Top             =   450
         Width           =   2565
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
         ItemData        =   "FrmMst_InspectionExpenseType.frx":0093
         Left            =   8055
         List            =   "FrmMst_InspectionExpenseType.frx":009D
         Sorted          =   -1  'True
         TabIndex        =   24
         Top             =   450
         Width           =   2220
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1455
         Left            =   75
         TabIndex        =   42
         Top             =   885
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
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Type"
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
         Left            =   11280
         TabIndex        =   51
         Top             =   165
         Width           =   480
      End
      Begin VB.Label Label4 
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
         Left            =   8895
         TabIndex        =   45
         Top             =   165
         Width           =   555
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Calculation on"
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
         Left            =   6052
         TabIndex        =   44
         Top             =   165
         Width           =   1260
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   2002
         TabIndex        =   43
         Top             =   165
         Width           =   1320
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1095
      Left            =   45
      TabIndex        =   38
      Top             =   5430
      Width           =   9360
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   375
         Left            =   7380
         TabIndex        =   21
         Top             =   615
         Width           =   1815
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   375
         Left            =   5565
         TabIndex        =   20
         Top             =   615
         Width           =   1815
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   375
         Left            =   3750
         TabIndex        =   19
         Top             =   615
         Width           =   1815
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   375
         Left            =   1935
         TabIndex        =   18
         Top             =   615
         Width           =   1815
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   615
         Width           =   1815
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "&Search"
         Height          =   375
         Left            =   7380
         TabIndex        =   16
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   5565
         TabIndex        =   15
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   3750
         TabIndex        =   14
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   375
         Left            =   1935
         TabIndex        =   13
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New "
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5415
      Left            =   45
      TabIndex        =   0
      Top             =   30
      Width           =   9360
      Begin VB.TextBox TxtOrderNo 
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
         Left            =   7530
         MaxLength       =   100
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   240
         Width           =   1515
      End
      Begin VB.CheckBox ChkClearFlag 
         Caption         =   "Clear Flag"
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
         Left            =   3960
         TabIndex        =   52
         Top             =   293
         Width           =   1335
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
         Height          =   360
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   240
         Width           =   1275
      End
      Begin VB.TextBox TxtCalculationOn 
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
         Left            =   3960
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   1680
         Width           =   2475
      End
      Begin VB.ComboBox CmbCalculationOn 
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
         ItemData        =   "FrmMst_InspectionExpenseType.frx":00B4
         Left            =   3960
         List            =   "FrmMst_InspectionExpenseType.frx":00CA
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   1680
         Width           =   2475
      End
      Begin VB.TextBox TxtAmount 
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
         Left            =   3960
         MaxLength       =   100
         TabIndex        =   8
         Top             =   2160
         Width           =   2475
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
         ItemData        =   "FrmMst_InspectionExpenseType.frx":0126
         Left            =   7530
         List            =   "FrmMst_InspectionExpenseType.frx":0130
         TabIndex        =   9
         Top             =   2160
         Width           =   1515
      End
      Begin VB.ComboBox CmbType 
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
         ItemData        =   "FrmMst_InspectionExpenseType.frx":0147
         Left            =   7530
         List            =   "FrmMst_InspectionExpenseType.frx":0151
         TabIndex        =   6
         Top             =   1680
         Width           =   1515
      End
      Begin VB.TextBox TxtFormula 
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
         Height          =   795
         Left            =   2520
         MaxLength       =   4000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   2640
         Width           =   6525
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
         Height          =   795
         Left            =   2520
         MaxLength       =   4000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   3555
         Width           =   6525
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
         Height          =   360
         Left            =   2520
         MaxLength       =   200
         TabIndex        =   3
         Top             =   1200
         Width           =   6525
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
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   4920
         Width           =   6525
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
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   4440
         Width           =   6525
      End
      Begin VB.CheckBox ChkUserInputAmount 
         Caption         =   "User Input"
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
         Left            =   2520
         TabIndex        =   7
         Top             =   2160
         Width           =   1335
      End
      Begin VB.CheckBox ChkUserInputCalc 
         Caption         =   "User Input"
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
         Left            =   2520
         TabIndex        =   4
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox TxtType 
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
         Left            =   7530
         MaxLength       =   25
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1515
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
         Height          =   360
         Left            =   7530
         MaxLength       =   25
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   2160
         Width           =   1515
      End
      Begin VB.ComboBox CmbGodownType 
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
         Left            =   2520
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   720
         Width           =   6525
      End
      Begin VB.TextBox TxtGodownType 
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
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   720
         Width           =   6525
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Godown Type"
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
         TabIndex        =   55
         Top             =   780
         Width           =   1275
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Order No"
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
         Left            =   6600
         TabIndex        =   53
         Top             =   300
         Width           =   825
      End
      Begin VB.Label LblVendorID 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   240
         TabIndex        =   33
         Top             =   300
         Width           =   1560
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Formula"
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
         TabIndex        =   48
         Top             =   2910
         Width           =   735
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
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
         Left            =   240
         TabIndex        =   47
         Top             =   3825
         Width           =   825
      End
      Begin VB.Label Label9 
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
         Left            =   240
         TabIndex        =   37
         Top             =   4500
         Width           =   855
      End
      Begin VB.Label Label8 
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
         Left            =   240
         TabIndex        =   36
         Top             =   4980
         Width           =   885
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Amount per Unit"
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
         TabIndex        =   35
         Top             =   2220
         Width           =   1410
      End
      Begin VB.Label LblVendorName 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   240
         TabIndex        =   32
         Top             =   1260
         Width           =   1515
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Calculation on (Units)"
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
         TabIndex        =   34
         Top             =   1740
         Width           =   2220
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Type"
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
         Left            =   6600
         TabIndex        =   50
         Top             =   1740
         Width           =   555
      End
      Begin VB.Label Label7 
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
         Left            =   6600
         TabIndex        =   40
         Top             =   2220
         Width           =   660
      End
   End
End
Attribute VB_Name = "FrmMst_InspectionExpenseType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg As String
Dim mInsExpTypeID As Double, mGodownTypeID As Double
Dim mOrderNo As Double

Private Sub CmbCalculationOn_LostFocus()
If CmbCalculationOn.Text = "" Then Exit Sub

If LCase(CmbCalculationOn.Text) <> "qty(per mt)" And LCase(CmbCalculationOn.Text) <> "area(per sq.ft)" And LCase(CmbCalculationOn.Text) <> "number(per unit)" And LCase(CmbCalculationOn.Text) <> "fixed" And LCase(CmbCalculationOn.Text) <> "per aum" And LCase(CmbCalculationOn.Text) <> "fumigation qty(per mt)" Then
   MsgBox "Invalid selection!!! "
   CmbCalculationOn.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmbGodownType_GotFocus()
tmp = CmbGodownType.Text
Call TableMst_UnitType(" Where Status  = 'A'")
CmbGodownType.Clear
If RsMst_UnitType.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_UnitType.RecordCount
    CmbGodownType.AddItem RsMst_UnitType!UnitType
    RsMst_UnitType.MoveNext
Next
CmbGodownType.Text = tmp
End Sub

Private Sub CmbGodownType_LostFocus()
If CmbGodownType.Text = "" Then
   Exit Sub
End If

If tmp <> CmbGodownType.Text Then
   ''write if clear date
End If

RsMst_UnitType.MoveFirst
RsMst_UnitType.Find "UnitType = '" & CmbGodownType.Text & "'"
If RsMst_UnitType.EOF Then
   MsgBox "Invalid selection!!!"
   CmbGodownType.SetFocus
   Exit Sub
End If

mGodownTypeID = RsMst_UnitType!UnitTypeID
End Sub

Private Sub CmbSCalculationOn_LostFocus()
If CmbSCalculationOn.Text = "" Then Exit Sub

If LCase(CmbSCalculationOn.Text) <> "qty(per mt)" And LCase(CmbSCalculationOn.Text) <> "area(per sq.ft)" And LCase(CmbSCalculationOn.Text) <> "number(per piece)" And LCase(CmbSCalculationOn.Text) <> "fixed" And LCase(CmbSCalculationOn.Text) <> "per aum" And LCase(CmbSCalculationOn.Text) <> "fumigation qty(per mt)" Then
   MsgBox "Invalid selection!!! "
   CmbSCalculationOn.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmbSStatus_LostFocus()
If CmbSStatus.Text = "" Then Exit Sub

If LCase(CmbSStatus.Text) <> "active" And LCase(CmbSStatus.Text) <> "de-active" Then
   MsgBox "Invalid selection!!! "
   CmbSStatus.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmbSType_LostFocus()
If CmbSType.Text = "" Then Exit Sub

If LCase(CmbSType.Text) <> "direct" And LCase(CmbSType.Text) <> "indirect" Then
   MsgBox "Invalid selection!!! "
   CmbSType.SetFocus
   Exit Sub
End If
End Sub


Private Sub CmbStatus_LostFocus()
If CmbStatus.Text = "" Then Exit Sub

If LCase(CmbStatus.Text) <> "active" And LCase(CmbStatus.Text) <> "de-active" Then
   MsgBox "Invalid selection!!! "
   CmbStatus.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmbType_LostFocus()
If CmbType.Text = "" Then Exit Sub

If LCase(CmbType.Text) <> "direct" And LCase(CmbType.Text) <> "indirect" Then
   MsgBox "Invalid selection!!! "
   CmbType.SetFocus
   Exit Sub
End If
End Sub


Private Sub CmdSearch_Click()
Dim x As Long
Dim wc As String

wc = ""
RsMst_InspectionExpenseType.MoveFirst
Call Grid_Head

If Trim(TxtSExpenseType) <> "" Then
   wc = " where ExpenseType like '%" & Trim(TxtSExpenseType) & "%' "
End If

If CmbSCalculationOn.Text <> "" Then
   If wc = "" Then
      wc = " where CalculationOn ='" & CmbSCalculationOn.Text & "'"
   Else
      wc = wc & " and CalculationOn ='" & CmbSCalculationOn.Text & "'"
   End If
End If

If CmbSStatus.Text <> "" Then
   If wc = "" Then
      wc = " where Status = '" & Left(CmbSStatus.Text, 1) & "'"
   Else
      wc = wc & " and Status ='" & Left(CmbSStatus.Text, 1) & "'"
   End If
End If

If CmbSType.Text <> "" Then
   If wc = "" Then
      wc = " where Type = '" & Left(CmbSType.Text, 1) & "'"
   Else
      wc = wc & " and Type ='" & Left(CmbSType.Text, 1) & "'"
   End If
End If

tmp = "Select MIE.InsExpTypeID, MIE.ExpenseType, MIE.CalculationOn, MIE.Amount, MIE.Formula, MIE.Remarks, MIE.Status, MIE.Type, MIE.UserInputAmount, MIE.UserInputCalc, MIE.ClearFlag, MIE.OrderNo, MIE.CreatedDate, MIE.CreatedBy, MIE.UpdatedDate, MIE.UpdatedBy, MUT.UnitType  FROM Mst_InspectionExpenseType MIE"
tmp = tmp & " Left JOin Mst_UnitType MUT on MUT.UnitTypeID = MIE.UnitTypeID"
tmp = tmp & wc
tmp = tmp & " order by MIE.OrderNo"
Call Tmp4Table

If TmpRs4.RecordCount > 0 Then
   CmdExcel.Enabled = True
   MSHFlexGrid1.Rows = TmpRs4.RecordCount + MSHFlexGrid1.Rows
   ProgressBar1.Max = MSHFlexGrid1.Rows
   For I = 1 To TmpRs4.RecordCount
      MSHFlexGrid1.TextMatrix(I, 0) = IIf(IsNull(TmpRs4!InsExpTypeID), "", TmpRs4!InsExpTypeID)
      MSHFlexGrid1.TextMatrix(I, 1) = IIf(IsNull(TmpRs4!UnitType), "", TmpRs4!UnitType)
      MSHFlexGrid1.TextMatrix(I, 2) = IIf(IsNull(TmpRs4!ExpenseType), "", TmpRs4!ExpenseType)
      MSHFlexGrid1.TextMatrix(I, 3) = IIf(IsNull(TmpRs4!CalculationOn), "", TmpRs4!CalculationOn)
      MSHFlexGrid1.TextMatrix(I, 4) = IIf(IsNull(TmpRs4!Amount), "", TmpRs4!Amount)
      MSHFlexGrid1.TextMatrix(I, 5) = IIf(IsNull(TmpRs4!Formula), "", TmpRs4!Formula)
      MSHFlexGrid1.TextMatrix(I, 6) = IIf(IsNull(TmpRs4!Remarks), "", TmpRs4!Remarks)
      MSHFlexGrid1.TextMatrix(I, 7) = IIf(IsNull(TmpRs4!Status), "", TmpRs4!Status)
      MSHFlexGrid1.TextMatrix(I, 8) = IIf(IsNull(TmpRs4!Type), "", TmpRs4!Type)
      MSHFlexGrid1.TextMatrix(I, 9) = IIf(IsNull(TmpRs4!UserInputAmount), "", TmpRs4!UserInputAmount)
      MSHFlexGrid1.TextMatrix(I, 10) = IIf(IsNull(TmpRs4!UserInputCalc), "", TmpRs4!UserInputCalc)
      MSHFlexGrid1.TextMatrix(I, 11) = IIf(IsNull(TmpRs4!ClearFlag), "", TmpRs4!ClearFlag)
      MSHFlexGrid1.TextMatrix(I, 12) = IIf(IsNull(TmpRs4!OrderNo), "", TmpRs4!OrderNo)
      MSHFlexGrid1.TextMatrix(I, 13) = IIf(IsNull(TmpRs4!CreatedBy), "", TmpRs4!CreatedBy & " : " & TmpRs4!CreatedDate)
      MSHFlexGrid1.TextMatrix(I, 14) = IIf(IsNull(TmpRs4!UpdatedBy), "", TmpRs4!UpdatedBy & " : " & TmpRs4!UpdatedDate)
      If LCase(MSHFlexGrid1.TextMatrix(I, 7)) = "a" Then
         MSHFlexGrid1.TextMatrix(I, 7) = "Active"
      ElseIf LCase(MSHFlexGrid1.TextMatrix(I, 7)) = "d" Then
         MSHFlexGrid1.TextMatrix(I, 7) = "De-Active"
      End If
      If LCase(MSHFlexGrid1.TextMatrix(I, 8)) = "d" Then
         MSHFlexGrid1.TextMatrix(I, 8) = "Direct"
      ElseIf LCase(MSHFlexGrid1.TextMatrix(I, 8)) = "i" Then
         MSHFlexGrid1.TextMatrix(I, 8) = "Indirect"
      End If
      If Val(MSHFlexGrid1.TextMatrix(I, 9)) = "1" Then
         MSHFlexGrid1.TextMatrix(I, 9) = "Yes"
      Else
         MSHFlexGrid1.TextMatrix(I, 9) = "No"
      End If
      If Val(MSHFlexGrid1.TextMatrix(I, 10)) = "1" Then
         MSHFlexGrid1.TextMatrix(I, 10) = "Yes"
      Else
         MSHFlexGrid1.TextMatrix(I, 10) = "No"
      End If
      If Val(MSHFlexGrid1.TextMatrix(I, 11)) = "1" Then
         MSHFlexGrid1.TextMatrix(I, 11) = "Yes"
      Else
         MSHFlexGrid1.TextMatrix(I, 11) = "No"
      End If
      TmpRs4.MoveNext
      ProgressBar1.Value = ProgressBar1.Value + 1
   Next
   ProgressBar1.Value = ProgressBar1.Max
Else
   MsgBox "Record not found!!!"
End If
End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Call GButtonLockAD(Me)
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If

Call Grid_Head

TxtSExpenseType.Text = ""
CmbSCalculationOn.Text = ""
CmbSStatus.Text = ""
CmbSType.Text = ""

Call TableMst_InspectionExpenseType(" Order by OrderNo")
If RsMst_InspectionExpenseType.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
Call GButtonLock(Me, True)
Call Indata
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Mst_InspectionExpenseType.xls")
Call Xls_Detail_Rpt
End Sub

Public Sub Xls_Detail_Rpt()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Long
Dim mCol As Long
MsgBox ("Wait till next message!!!")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Mst_InspectionExpenseType.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid1.Rows - 1
   mRow = I + 1
   For c = 0 To MSHFlexGrid1.Cols - 1
      oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(I, c)
   Next
Next
oWB.Save
MsgBox ("Excel file created!!!")
oXL.Visible = True
End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_InspectionExpenseType.MoveNext
If RsMst_InspectionExpenseType.EOF Then
   RsMst_InspectionExpenseType.MoveLast
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
RsMst_InspectionExpenseType.MovePrevious
If RsMst_InspectionExpenseType.BOF Then
   RsMst_InspectionExpenseType.MoveFirst
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
RsMst_InspectionExpenseType.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub

Private Sub LastCmd_Click()
RsMst_InspectionExpenseType.MoveLast
Call Indata
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False
End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   If RsMst_InspectionExpenseType.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   Call Indata
   Exit Sub
End If

Frame0.Height = Me.Height - 600
Frame0.Left = Frame1.Left
Frame0.Top = Frame1.Top
Frame0.Width = Me.Width - 150
Frame0.Visible = True

MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 1200

ProgressBar1.Width = MSHFlexGrid1.Width
ProgressBar1.Left = MSHFlexGrid1.Left
ProgressBar1.Top = MSHFlexGrid1.Top + MSHFlexGrid1.Height
End Sub

Public Sub Indata()
Frame1.Enabled = False

CmbCalculationOn.Visible = False
TxtCalculationOn.Visible = True
CmbStatus.Visible = False
TxtStatus.Visible = True
CmbType.Visible = False
TxtType.Visible = True
CmbGodownType.Visible = False
TxtGodownType.Visible = True

TxtExpenseTypeID.Text = IIf(IsNull(RsMst_InspectionExpenseType!InsExpTypeID), "", RsMst_InspectionExpenseType!InsExpTypeID)
TxtExpenseType.Text = IIf(IsNull(RsMst_InspectionExpenseType!ExpenseType), "", RsMst_InspectionExpenseType!ExpenseType)
TxtCalculationOn.Text = IIf(IsNull(RsMst_InspectionExpenseType!CalculationOn), "", RsMst_InspectionExpenseType!CalculationOn)
TxtAmount.Text = IIf(IsNull(RsMst_InspectionExpenseType!Amount), "", RsMst_InspectionExpenseType!Amount)
ChkUserInputAmount.Value = IIf(IsNull(RsMst_InspectionExpenseType!UserInputAmount), 0, RsMst_InspectionExpenseType!UserInputAmount)
ChkUserInputCalc.Value = IIf(IsNull(RsMst_InspectionExpenseType!UserInputCalc), 0, RsMst_InspectionExpenseType!UserInputCalc)
ChkClearFlag.Value = IIf(IsNull(RsMst_InspectionExpenseType!ClearFlag), 0, RsMst_InspectionExpenseType!ClearFlag)
TxtGodownType.Text = GetUnitName(IIf(IsNull(RsMst_InspectionExpenseType!UnitTypeID), 0, RsMst_InspectionExpenseType!UnitTypeID))
TxtStatus.Text = IIf(IsNull(RsMst_InspectionExpenseType!Status), "", RsMst_InspectionExpenseType!Status)
TxtType.Text = IIf(IsNull(RsMst_InspectionExpenseType!Type), "", RsMst_InspectionExpenseType!Type)
TxtFormula.Text = IIf(IsNull(RsMst_InspectionExpenseType!Formula), "", RsMst_InspectionExpenseType!Formula)
TxtRemarks.Text = IIf(IsNull(RsMst_InspectionExpenseType!Remarks), "", RsMst_InspectionExpenseType!Remarks)
TxtOrderNo.Text = IIf(IsNull(RsMst_InspectionExpenseType!OrderNo), "", RsMst_InspectionExpenseType!OrderNo)
TxtCreated.Text = IIf(IsNull(RsMst_InspectionExpenseType!CreatedBy), "", RsMst_InspectionExpenseType!CreatedBy) & " :- " & IIf(IsNull(RsMst_InspectionExpenseType!CreatedDate), "", RsMst_InspectionExpenseType!CreatedDate)
TxtUpdated.Text = IIf(IsNull(RsMst_InspectionExpenseType!UpdatedBy), "", RsMst_InspectionExpenseType!UpdatedBy) & " :- " & IIf(IsNull(RsMst_InspectionExpenseType!UpdatedDate), "", RsMst_InspectionExpenseType!UpdatedDate)
If LCase(TxtStatus.Text) = "a" Then
   TxtStatus.Text = "Active"
ElseIf LCase(TxtStatus.Text) = "d" Then
   TxtStatus.Text = "De-Active"
End If
If LCase(TxtType.Text) = "d" Then
   TxtType.Text = "Direct"
ElseIf LCase(TxtType.Text) = "i" Then
   TxtType.Text = "Indirect"
End If
CmbGodownType.Text = TxtGodownType.Text
CmbStatus.Text = TxtStatus.Text
CmbType.Text = TxtType.Text
CmbCalculationOn.Text = TxtCalculationOn.Text

Call GButtonLock(Me, True)
End Sub

Private Sub AddCmd_Click()

Edit_Flg = "A"
Call GButtonLock(Me, False)
Frame1.Enabled = True

CmbGodownType.Visible = True
TxtGodownType.Visible = False
CmbCalculationOn.Visible = True
TxtCalculationOn.Visible = False
CmbStatus.Visible = True
TxtStatus.Visible = False
CmbType.Visible = True
TxtType.Visible = False

TxtExpenseTypeID.Text = ""
TxtExpenseType.Text = ""
TxtCalculationOn.Text = ""
TxtAmount.Text = ""
ChkUserInputAmount.Value = vbUnchecked
ChkUserInputCalc.Value = vbUnchecked
ChkClearFlag.Value = vbUnchecked
TxtGodownType.Text = ""
TxtStatus.Text = ""
TxtFormula.Text = ""
TxtRemarks.Text = ""
TxtCreated.Text = ""
TxtUpdated.Text = ""
CmbCalculationOn.Text = ""
CmbGodownType.Text = ""
CmbStatus.Text = ""
CmbType.Text = ""
TxtOrderNo.Text = GetOrderNo
If RsMst_InspectionExpenseType.RecordCount > 0 Then
   TxtExpenseType.SetFocus
End If

End Sub

Private Sub EditCmd_Click()
Edit_Flg = "E"
Call GButtonLock(Me, False)
Frame1.Enabled = True

CmbGodownType.Visible = True
TxtGodownType.Visible = False
CmbCalculationOn.Visible = True
TxtCalculationOn.Visible = False
CmbStatus.Visible = True
TxtStatus.Visible = False
CmbType.Visible = True
TxtType.Visible = False

End Sub

Private Sub DeleteCmd_Click()
Dim ans As Variant
On Error GoTo msgError
ans = MsgBox("Do you really want to DELETE this record???", vbYesNo)
If ans = vbYes Then
   RsMst_InspectionExpenseType.Delete
   RsMst_InspectionExpenseType.Update
   If RsMst_InspectionExpenseType.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_InspectionExpenseType.MoveNext
   If RsMst_InspectionExpenseType.EOF() Then
      RsMst_InspectionExpenseType.MoveLast
   End If
   Call Indata
End If
Exit Sub
msgError:
   MsgBox "Child record Present "
   RsMst_InspectionExpenseType.CancelUpdate
End Sub

Private Sub SaveCmd_Click()
If TxtExpenseType.Text = "" Then
   MsgBox "Blank Expense Type Not Allowed!!!"
   TxtExpenseType.SetFocus
   Exit Sub
End If
If CmbCalculationOn.Text = "" Then
   MsgBox "Blank Calculation On Not Allowed!!!"
   CmbCalculationOn.SetFocus
   Exit Sub
End If
'If TxtAmount.Text = "" Then
'   MsgBox "Blank Amount Not Allowed!!!"
'   TxtAmount.SetFocus
'   Exit Sub
'End If
If CmbStatus.Text = "" Then
   MsgBox "Blank Status Not Allowed!!!"
   CmbStatus.SetFocus
   Exit Sub
End If
If CmbType.Text = "" Then
   MsgBox "Blank Type Not Allowed!!!"
   CmbType.SetFocus
   Exit Sub
End If
'If TxtFormula.Text = "" Then
'   MsgBox "Blank Formula Not Allowed!!!"
'   TxtFormula.SetFocus
'   Exit Sub
'End If
If TxtRemarks.Text = "" Then
   MsgBox "Blank Remarks Not Allowed!!!"
   TxtRemarks.SetFocus
   Exit Sub
End If

If Edit_Flg = "A" Then
   If RsMst_InspectionExpenseType.RecordCount > 0 Then
      RsMst_InspectionExpenseType.MoveFirst
      RsMst_InspectionExpenseType.Find "ExpenseType= '" & Trim(TxtExpenseType.Text) & "'"
      If Not RsMst_InspectionExpenseType.EOF Then
         MsgBox "Duplicate Expense Type Not Allowed!!!"
         TxtExpenseType.SetFocus
         Exit Sub
      End If
   End If
   RsMst_InspectionExpenseType.AddNew
   RsMst_InspectionExpenseType!InsExpTypeID = GetInsExpTypeID
   RsMst_InspectionExpenseType!G_UID = GetGUID
   TxtExpenseTypeID.Text = RsMst_InspectionExpenseType!InsExpTypeID
   mOrderNo = 0
Else
   RsMst_InspectionExpenseType.MoveFirst
   RsMst_InspectionExpenseType.Find "ExpenseType= '" & Trim(TxtExpenseType.Text) & "'"
   If Not RsMst_InspectionExpenseType.EOF Then
      If RsMst_InspectionExpenseType!InsExpTypeID <> TxtExpenseTypeID.Text Then
         MsgBox "Duplicate Expense Type Not Allowed!!!"
         TxtExpenseType.SetFocus
         Exit Sub
      End If
   End If
   RsMst_InspectionExpenseType.MoveFirst
   RsMst_InspectionExpenseType.Find " InsExpTypeID= " & TxtExpenseTypeID.Text
   mOrderNo = IIf(IsNull(RsMst_InspectionExpenseType!OrderNo), 0, RsMst_InspectionExpenseType!OrderNo)
End If
If Trim(CmbGodownType.Text) <> "" And mGodownTypeID > 0 Then
   RsMst_InspectionExpenseType!UnitTypeID = mGodownTypeID
Else
   RsMst_InspectionExpenseType!UnitTypeID = Null
End If
RsMst_InspectionExpenseType!ExpenseType = TxtExpenseType.Text
RsMst_InspectionExpenseType!CalculationOn = CmbCalculationOn.Text
RsMst_InspectionExpenseType!Amount = Val(TxtAmount.Text)
RsMst_InspectionExpenseType!UserInputAmount = ChkUserInputAmount.Value
RsMst_InspectionExpenseType!UserInputCalc = ChkUserInputCalc.Value
RsMst_InspectionExpenseType!ClearFlag = ChkClearFlag.Value
RsMst_InspectionExpenseType!Status = Left$(CmbStatus.Text, 1)
RsMst_InspectionExpenseType!Type = Left$(CmbType.Text, 1)
RsMst_InspectionExpenseType!Formula = TxtFormula.Text
RsMst_InspectionExpenseType!Remarks = TxtRemarks.Text
If Val(TxtOrderNo.Text) = 0 Then
   RsMst_InspectionExpenseType!OrderNo = GetOrderNo
Else
   RsMst_InspectionExpenseType!OrderNo = Val(TxtOrderNo.Text)
End If
If IsNull(RsMst_InspectionExpenseType!CreatedBy) = True Or RsMst_InspectionExpenseType!CreatedBy = "" Then
   RsMst_InspectionExpenseType!CreatedBy = Gen_UserLoginID
   RsMst_InspectionExpenseType!CreatedDate = Now
Else
   RsMst_InspectionExpenseType!UpdatedBy = Gen_UserLoginID
   RsMst_InspectionExpenseType!UpdatedDate = Now
End If
RsMst_InspectionExpenseType.Update

If Val(TxtOrderNo.Text) <> 0 Then
   tmp = "Select InsExpTypeID From Mst_InspectionExpenseType WHere OrderNo = " & Val(TxtOrderNo.Text) & " And InsExpTypeID <> " & Val(TxtExpenseTypeID.Text)
   Call TmpTable
   If TmpRs.RecordCount > 0 Then
      If mOrderNo = 0 Then
         tmp = " Update Mst_InspectionExpenseType set OrderNo = OrderNo + 1 WHere OrderNo >= " & Val(TxtOrderNo.Text) & " And InsExpTypeID <> " & Val(TxtExpenseTypeID.Text)
         Call Tmp7Table
      Else
         If Val(TxtOrderNo.Text) > mOrderNo Then
            tmp = " Update Mst_InspectionExpenseType set OrderNo = OrderNo - 1 WHere InsExpTypeID <> " & Val(TxtExpenseTypeID.Text)
            tmp = tmp & " And OrderNo <= " & Val(TxtOrderNo.Text) & " And OrderNo > " & mOrderNo
            Call Tmp7Table
         Else
            tmp = " Update Mst_InspectionExpenseType set OrderNo = OrderNo + 1 WHere InsExpTypeID <> " & Val(TxtExpenseTypeID.Text)
            tmp = tmp & " And OrderNo >= " & Val(TxtOrderNo.Text) & " And OrderNo < " & mOrderNo
            Call Tmp7Table
         End If
      End If
   End If
End If

RsMst_InspectionExpenseType.Requery
RsMst_InspectionExpenseType.MoveFirst
RsMst_InspectionExpenseType.Find " InsExpTypeID= " & TxtExpenseTypeID.Text
Call Indata
End Sub

Private Function GetInsExpTypeID() As Double
Dim Retval As Double
tmp = "Select max(InsExpTypeID) from Mst_InspectionExpenseType"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetInsExpTypeID = Retval
End Function

Private Function GetOrderNo() As Double
Dim Retval As Double
tmp = "Select max(OrderNo) from Mst_InspectionExpenseType"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetOrderNo = Retval
End Function

Private Sub Grid_Head()
ProgressBar1.Value = 0
CmdExcel.Enabled = False

With MSHFlexGrid1
   .Clear
   .Rows = 2
   .Cols = 15
   
   .TextMatrix(0, 0) = "Expense TypeID"
   .TextMatrix(0, 1) = "Godown Type"
   .TextMatrix(0, 2) = "Expense Type"
   .TextMatrix(0, 3) = "Calculation On"
   .TextMatrix(0, 4) = "Amount per Unit"
   .TextMatrix(0, 5) = "Formula"
   .TextMatrix(0, 6) = "Remarks"
   .TextMatrix(0, 7) = "Status"
   .TextMatrix(0, 8) = "Type"
   .TextMatrix(0, 9) = "User Input Amount"
   .TextMatrix(0, 10) = "User Input Calculation"
   .TextMatrix(0, 11) = "Clear Flag"
   .TextMatrix(0, 12) = "Order No"
   .TextMatrix(0, 13) = "CreatedBy"
   .TextMatrix(0, 14) = "UpdatedBy"
   
   .ColWidth(0) = 1000
   .ColWidth(1) = 1500
   .ColWidth(2) = 3000
   .ColWidth(3) = 2500
   .ColWidth(4) = 1100
   .ColWidth(5) = 1500
   .ColWidth(6) = 2200
   .ColWidth(7) = 1100
   
   .RowHeight(0) = 600
   
End With
End Sub

Private Sub MSHFlexGrid1_Click()
Dim I As Variant
I = MSHFlexGrid1.RowSel
If I <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(I, 0) = "" Then Exit Sub
RsMst_InspectionExpenseType.MoveFirst
RsMst_InspectionExpenseType.Find "InsExpTypeID = " & MSHFlexGrid1.TextMatrix(I, 0)
If Not RsMst_InspectionExpenseType.EOF Then
   Call Indata
   Frame0.Visible = False
End If
End Sub

Private Sub TxtAmount_LostFocus()
If Trim(TxtAmount.Text) = "" Then Exit Sub
If IsNumeric(TxtAmount.Text) = False Or Val(TxtAmount.Text) < 0 Then
   MsgBox "Invalid numeric format"
   TxtAmount.SetFocus
   Exit Sub
End If

TxtAmount.Text = Format(TxtAmount.Text, "#0.00")
End Sub

'Private Sub TxtFormula_LostFocus()
'On Error GoTo Error
'If SaveCmd.Enabled = False Then Exit Sub
'If TxtFormula = "" Then
'   TxtFormulaView.Text = ""
'   Exit Sub
'End If
'Dim mFormulaStr As String
'Dim mFormula As String
'Dim mRes As Double
'mFormulaStr = ""
'Dim Str_() As String
'TxtFormula.Text = Replace(TxtFormula.Text, Chr(32), "")
'TxtFormula.Text = Replace(TxtFormula.Text, vbCrLf, "")
'mFormula = ""
'Str_ = Split(TxtFormula.Text, "~")
'For I = 0 To UBound(Str_)
'    If Len(Str_(I)) = 0 Then
'       'mFormulaStr = mFormulaStr & Str_(i)
'    ElseIf (Asc(Str_(I)) >= 65 And Asc(Str_(I)) <= 90) Or (Asc(Str_(I)) >= 97 And Asc(Str_(I)) <= 122) Then
'       mFormulaStr = mFormulaStr & Str_(I)
'       mFormula = mFormula & 2
'    ElseIf (Asc(Str_(I)) >= 40 And Asc(Str_(I)) <= 43) Or Asc(Str_(I)) = 45 Or Asc(Str_(I)) = 47 Or Asc(Str_(I)) = 94 Then
'       mFormulaStr = mFormulaStr & Str_(I)
'       mFormula = mFormula & Str_(I)
'    Else
'       mFormulaStr = mFormulaStr & Val(Str_(I))
'       mFormula = mFormula & Val(Str_(I))
'    End If
'Next
'mRes = SC.Eval(mFormula)
''TxtFormulaView.Text = mFormulaStr
'Exit Sub
'Error:
'MsgBox "Invalid Formula!!!"
'TxtFormula.SetFocus
'End Sub
Private Sub TxtOrderNo_Change()

End Sub

Private Sub TxtOrderNo_LostFocus()
If Trim(TxtOrderNo.Text) = "" Then Exit Sub
If IsNumeric(TxtOrderNo.Text) = False Or Val(TxtOrderNo.Text) < 0 Then
   MsgBox "Invalid numeric format"
   TxtOrderNo.SetFocus
   Exit Sub
End If

TxtOrderNo.Text = Format(TxtOrderNo.Text, "#0")
End Sub
