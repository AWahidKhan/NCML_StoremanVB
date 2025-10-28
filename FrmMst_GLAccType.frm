VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmMst_GLAccType 
   Caption         =   "GL A/C Type Master"
   ClientHeight    =   6405
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9435
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   6405
   ScaleWidth      =   9435
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   1365
      Left            =   120
      TabIndex        =   13
      Top             =   9840
      Visible         =   0   'False
      Width           =   12720
      Begin VB.ComboBox CmbSGroupType 
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
         ItemData        =   "FrmMst_GLAccType.frx":0000
         Left            =   5880
         List            =   "FrmMst_GLAccType.frx":000A
         Sorted          =   -1  'True
         TabIndex        =   23
         Top             =   187
         Width           =   1560
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
         Left            =   8655
         TabIndex        =   25
         Top             =   195
         Width           =   1000
      End
      Begin VB.TextBox TxtSearchGLAccType 
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
         Left            =   1800
         TabIndex        =   22
         Top             =   165
         Width           =   2685
      End
      Begin VB.CommandButton CmdGo 
         Caption         =   "Go"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   7530
         TabIndex        =   24
         Top             =   195
         Width           =   1005
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   615
         Left            =   75
         TabIndex        =   27
         Top             =   600
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   1085
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
      Begin VB.Label Label7 
         Caption         =   "Group Type"
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
         Left            =   4680
         TabIndex        =   52
         Top             =   210
         Width           =   1095
      End
      Begin VB.Label Label19 
         Caption         =   "GL Account Type"
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
         Left            =   135
         TabIndex        =   29
         Top             =   210
         Width           =   1575
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
         Left            =   9765
         TabIndex        =   28
         Top             =   210
         Width           =   2355
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
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
         Left            =   12165
         TabIndex        =   26
         Top             =   210
         Width           =   75
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
      Height          =   4095
      Left            =   120
      TabIndex        =   10
      Top             =   4560
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   7223
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
   Begin VB.Frame Frame2 
      Caption         =   "Type Mapping"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   45
      Top             =   3240
      Width           =   8175
      Begin VB.ComboBox CmbExpenseTypeID 
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
         Left            =   2040
         Sorted          =   -1  'True
         TabIndex        =   8
         Top             =   240
         Width           =   5895
      End
      Begin VB.TextBox TxtExchangeType 
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
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   720
         Width           =   2040
      End
      Begin VB.CommandButton CmdAdd2List 
         Caption         =   "Add To List"
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
         Left            =   4200
         TabIndex        =   9
         Top             =   705
         Width           =   1635
      End
      Begin VB.CommandButton CmdRemoveFromList 
         Caption         =   "Remove From List"
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
         Left            =   5880
         TabIndex        =   11
         Top             =   705
         Width           =   1995
      End
      Begin VB.ComboBox CmbIncomeExpenseTypeID 
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
         Left            =   2040
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   240
         Width           =   5895
      End
      Begin VB.Label Label6 
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
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   255
         Width           =   1545
      End
      Begin VB.Label Label4 
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
         Left            =   120
         TabIndex        =   48
         Top             =   780
         Width           =   1665
      End
      Begin VB.Label Label12 
         Caption         =   "Income Type"
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
         TabIndex        =   46
         Top             =   255
         Width           =   1545
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1155
      Left            =   120
      TabIndex        =   39
      Top             =   8640
      Width           =   8160
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
         Left            =   100
         TabIndex        =   0
         Top             =   240
         Width           =   1560
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
         Left            =   100
         TabIndex        =   17
         Top             =   645
         Width           =   1560
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
         Left            =   1665
         TabIndex        =   12
         Top             =   240
         Width           =   1560
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
         Left            =   1665
         TabIndex        =   18
         Top             =   645
         Width           =   1560
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
         Left            =   3225
         TabIndex        =   14
         Top             =   240
         Width           =   1560
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
         Height          =   405
         Left            =   3225
         TabIndex        =   19
         Top             =   645
         Width           =   1560
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
         Left            =   4785
         TabIndex        =   15
         Top             =   240
         Width           =   1560
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
         Left            =   4785
         TabIndex        =   20
         Top             =   645
         Width           =   1560
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
         Left            =   6345
         TabIndex        =   16
         Top             =   240
         Width           =   1560
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
         Left            =   6345
         TabIndex        =   21
         Top             =   645
         Width           =   1560
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3255
      Left            =   120
      TabIndex        =   30
      Top             =   0
      Width           =   8160
      Begin VB.ComboBox CmbGroupType 
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
         ItemData        =   "FrmMst_GLAccType.frx":001F
         Left            =   5550
         List            =   "FrmMst_GLAccType.frx":0029
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   1900
         Width           =   2400
      End
      Begin VB.TextBox TxtGroupType 
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
         Left            =   5550
         Locked          =   -1  'True
         TabIndex        =   50
         Top             =   1900
         Width           =   2400
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
         ItemData        =   "FrmMst_GLAccType.frx":003E
         Left            =   5550
         List            =   "FrmMst_GLAccType.frx":0048
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   1450
         Width           =   2400
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
         Left            =   5550
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   1455
         Width           =   2400
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
         ItemData        =   "FrmMst_GLAccType.frx":0059
         Left            =   2070
         List            =   "FrmMst_GLAccType.frx":0063
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   1900
         Width           =   1920
      End
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
         Height          =   360
         Left            =   2070
         MaxLength       =   15
         TabIndex        =   3
         Top             =   1450
         Width           =   1920
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
         Height          =   360
         Left            =   2070
         TabIndex        =   40
         Top             =   1900
         Width           =   1920
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
         Height          =   360
         Left            =   2070
         MaxLength       =   25
         TabIndex        =   2
         Top             =   1050
         Width           =   5880
      End
      Begin VB.TextBox TxtGLAccType 
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
         Left            =   2070
         MaxLength       =   100
         TabIndex        =   1
         Top             =   650
         Width           =   5880
      End
      Begin VB.TextBox TxtGLAccTypeID 
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
         Left            =   2085
         Locked          =   -1  'True
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   240
         Width           =   1920
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
         Left            =   2070
         Locked          =   -1  'True
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   2750
         Width           =   5880
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
         Left            =   2070
         Locked          =   -1  'True
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   2310
         Width           =   5880
      End
      Begin VB.Label Label5 
         Caption         =   "Group Type"
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
         Left            =   4080
         TabIndex        =   49
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label3 
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
         Height          =   285
         Left            =   4080
         TabIndex        =   44
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "HSN No"
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
         Left            =   120
         TabIndex        =   42
         Top             =   1488
         Width           =   1800
      End
      Begin VB.Label Label1 
         Caption         =   "Flag"
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
         Left            =   120
         TabIndex        =   41
         Top             =   1938
         Width           =   1800
      End
      Begin VB.Label LblNAVCode 
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
         TabIndex        =   38
         Top             =   1103
         Width           =   1800
      End
      Begin VB.Label LblGLAccType 
         Caption         =   "GL Account Type"
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
         TabIndex        =   37
         Top             =   703
         Width           =   1920
      End
      Begin VB.Label LblGLACID 
         Caption         =   "GL A/C Type ID"
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
         TabIndex        =   36
         Top             =   315
         Width           =   1800
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
         Left            =   120
         TabIndex        =   35
         Top             =   2803
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
         TabIndex        =   34
         Top             =   2370
         Width           =   1695
      End
   End
End
Attribute VB_Name = "FrmMst_GLAccType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg As Variant
Dim mIncomeExpenseTypeID As Double
Dim mExchangeTypeID As Double

Dim mGroupType As Variant
Dim mExpenseTypeID As Double


''-- Ullhas Patil 28 Sep 2017
Private Sub CmbExpenseTypeID_GotFocus()
tmp = CmbExpenseTypeID.Text
Call TableMst_ExpenseType(" Where Flag = 'A' ")
CmbExpenseTypeID.Clear
If RsMst_ExpenseType.RecordCount = 0 Then
   MsgBox "No Expense Type found"
   CmbExpenseTypeID.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_ExpenseType.RecordCount
    CmbExpenseTypeID.AddItem RsMst_ExpenseType!ExpenseType
    RsMst_ExpenseType.MoveNext
Next
CmbExpenseTypeID.Text = tmp
End Sub
Private Sub CmbExpenseTypeID_LostFocus()
If CmbExpenseTypeID.Text = "" Then
   TxtExchangeType = ""
   Exit Sub
End If
RsMst_ExpenseType.MoveFirst
RsMst_ExpenseType.Find "ExpenseType = '" & CmbExpenseTypeID.Text & "'"
If RsMst_ExpenseType.EOF Then
   MsgBox "Invalid selection "
   CmbExpenseTypeID.SetFocus
   Exit Sub
End If
mExpenseTypeID = RsMst_ExpenseType!ExpenseTypeID
End Sub
'''-------------------

Private Sub CmbFlag_LostFocus()
If LCase(CmbFlag.Text) = "" Then Exit Sub

If LCase(CmbFlag.Text) <> "active" And LCase(CmbFlag.Text) <> "de-active" Then
   MsgBox "Invalid Selection"
   CmbFlag.SetFocus
   Exit Sub
End If

End Sub

''-- Ullhas Patil 28 Sep 2017
Private Sub CmbGroupType_LostFocus()
If LCase(CmbGroupType.Text) = "" Then Exit Sub
If LCase(CmbGroupType.Text) <> "income" And LCase(CmbGroupType.Text) <> "expense" Then
   MsgBox "Invalid Selection"
   CmbGroupType.SetFocus
   Exit Sub
End If
mGroupType = Left(CmbGroupType.Text, 1)
If LCase(mGroupType) = LCase("I") Then
    Label6.Visible = False
    CmbExpenseTypeID.Visible = False
    
    Label12.Visible = True
    CmbIncomeExpenseTypeID.Visible = True
    Label4.Visible = True
    TxtExchangeType.Visible = True
ElseIf LCase(mGroupType) = LCase("E") Then
    Label6.Visible = True
    CmbExpenseTypeID.Visible = True
    
    Label12.Visible = False
    CmbIncomeExpenseTypeID.Visible = False
    Label4.Visible = False
    TxtExchangeType.Visible = False
End If
Call Grid_Head2

End Sub
''--------------

''-- 12th Aug 2017 Ullhas Patil
Private Sub CmbIncomeExpenseTypeID_GotFocus()
tmp = CmbIncomeExpenseTypeID.Text
Call TableMst_IncomeExpenseType(" Where IEFlag = 'I' And Flag = 'A' ")
CmbIncomeExpenseTypeID.Clear
If RsMst_IncomeExpenseType.RecordCount = 0 Then
   MsgBox "No Income Expense Type found"
   CmbIncomeExpenseTypeID.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_IncomeExpenseType.RecordCount
    CmbIncomeExpenseTypeID.AddItem RsMst_IncomeExpenseType!IEType
    RsMst_IncomeExpenseType.MoveNext
Next
CmbIncomeExpenseTypeID.Text = tmp
End Sub
Private Sub CmbIncomeExpenseTypeID_LostFocus()
If CmbIncomeExpenseTypeID.Text = "" Then
   TxtExchangeType = ""
   Exit Sub
End If
RsMst_IncomeExpenseType.MoveFirst
RsMst_IncomeExpenseType.Find "IEType = '" & CmbIncomeExpenseTypeID.Text & "'"
If RsMst_IncomeExpenseType.EOF Then
   MsgBox "Invalid selection "
   CmbIncomeExpenseTypeID.SetFocus
   Exit Sub
End If
mIncomeExpenseTypeID = RsMst_IncomeExpenseType!IETypeID
mExchangeTypeID = RsMst_IncomeExpenseType!ExchangeTypeID
TxtExchangeType = GetExchangeName(RsMst_IncomeExpenseType!ExchangeTypeID)
End Sub
'''-----------------------------------------

Private Sub CmbSGroupType_LostFocus()
    If LCase(CmbSGroupType.Text) = "" Then Exit Sub
    If LCase(CmbSGroupType.Text) <> "income" And LCase(CmbSGroupType.Text) <> "expense" Then
       MsgBox "Invalid Selection"
       CmbSGroupType.SetFocus
       Exit Sub
    End If
End Sub

Private Sub CmbType_LostFocus()
If LCase(CmbType.Text) = "" Then Exit Sub
If LCase(CmbType.Text) <> "agri" And LCase(CmbType.Text) <> "nagri" Then
   MsgBox "Invalid Selection"
   CmbType.SetFocus
   Exit Sub
End If
End Sub

''-- 12th AUg 2017 Ullhas Patil
Function CheckIncomeTypeCount(mGLAccountTypeID_ As Double, mIncomeTypeID_ As Double, mAccountType_ As Variant) As Boolean
Dim mResult As Boolean
mResult = True
tmp = " Select MGAIM.GLAccountIncomeTypeMappingID,MGAIM.GLAccTypeID,MGLT.GLAccType,MGAIM.IETypeID,MIET.IEType,MGLT.AccountType"
tmp = tmp & " FROM Mst_GLAccountIncomeTypeMapping MGAIM"
tmp = tmp & " Inner Join Mst_GLAccType MGLT On MGAIM.GLAccTypeID = MGLT.GLAccTypeID"
tmp = tmp & " Inner Join Mst_IncomeExpenseType MIET On MGAIM.IETypeID = MIET.IETypeID"
tmp = tmp & " Where MGLT.GLAccTypeID <> " & Val(mGLAccountTypeID_) & " And MGLT.AccountType = '" & mAccountType_ & "' And MIET.IETypeID = " & Val(mIncomeTypeID_) & ""
Call TmpTable
If TmpRs.EOF = False Then
    If TmpRs.RecordCount >= 1 Then
       mResult = False
    End If
End If
CheckIncomeTypeCount = mResult
End Function
''-- Ullhas Patil 29 Sep 2017
Function CheckExpenseTypeCount(mGLAccountTypeID_ As Double, mExpenseTypeID_ As Double, mAccountType_ As Variant) As Boolean
Dim mResult As Boolean
mResult = True
tmp = " Select MGAIM.GLAccountIncomeTypeMappingID,MGAIM.GLAccTypeID,MGLT.GLAccType,MET.ExpenseTypeID,MGLT.AccountType"
tmp = tmp & " FROM Mst_GLAccountIncomeTypeMapping MGAIM"
tmp = tmp & " Inner Join Mst_GLAccType MGLT On MGAIM.GLAccTypeID = MGLT.GLAccTypeID"
tmp = tmp & " Inner Join Mst_ExpenseType MET On MGAIM.ExpenseTypeID = MET.ExpenseTypeID"
tmp = tmp & " Where MGLT.GLAccTypeID <> " & Val(mGLAccountTypeID_) & " And"
tmp = tmp & " MGLT.AccountType = '" & mAccountType_ & "' And MET.ExpenseTypeID = " & Val(mExpenseTypeID_) & ""
Call TmpTable
If TmpRs.EOF = False Then
    If TmpRs.RecordCount >= 1 Then
       mResult = False
    End If
End If
CheckExpenseTypeCount = mResult
End Function
''-------------------------------

''-- 12th Aug 2017 Ullhas Patil
Private Sub CmdAdd2List_Click()
Dim mR, mRowsel As Double

If LCase(mGroupType) = LCase("I") Then
    If CmbIncomeExpenseTypeID.Text = "" Then
        MsgBox "Blank Income Type not Allowed!!!"
        CmbIncomeExpenseTypeID.SetFocus
        Exit Sub
    End If
    
    If LCase(CmdAdd2List.Caption) = LCase("Add to List") Then
        For mR = 1 To MSHFlexGrid2.Rows - 1
            If CmbIncomeExpenseTypeID.Text = MSHFlexGrid2.TextMatrix(mR, 1) Then
                MsgBox "Income Type Already Exist!!!"
                Exit Sub
            End If
        Next
        If CheckIncomeTypeCount(CDbl(Val(TxtGLAccTypeID)), mIncomeExpenseTypeID, CmbType.Text) = False Then
            MsgBox "Income Type - " & CmbIncomeExpenseTypeID.Text & " already mapped with other G/L Account Type."
            Exit Sub
        End If
        mR = MSHFlexGrid2.Rows - 1
        MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
    Else
        mRowsel = MSHFlexGrid2.RowSel
        For mR = 1 To MSHFlexGrid2.Rows - 1
            If CmbIncomeExpenseTypeID.Text = MSHFlexGrid2.TextMatrix(mR, 1) Then
               If mR <> mRowsel Then
                    MsgBox "Income Type Already Exist!!!"
                    Exit Sub
               End If
            End If
        Next
        If CheckIncomeTypeCount(CDbl(Val(TxtGLAccTypeID)), mIncomeExpenseTypeID, CmbType.Text) = False Then
            MsgBox "Income Type - " & CmbIncomeExpenseTypeID.Text & " already mapped with other G/L Account Type."
            Exit Sub
        End If
        mR = mRowsel
    End If
    
    MSHFlexGrid2.TextMatrix(mR, 0) = ""
    MSHFlexGrid2.TextMatrix(mR, 1) = CmbIncomeExpenseTypeID.Text
    MSHFlexGrid2.TextMatrix(mR, 2) = mIncomeExpenseTypeID
    MSHFlexGrid2.TextMatrix(mR, 3) = mExchangeTypeID
    
    CmbIncomeExpenseTypeID.Text = ""
    TxtExchangeType = ""
    CmdAdd2List.Caption = "Add To List"
    CmdRemoveFromList.Caption = "Remove From List"
Else
    ''-- Expense Type
    If CmbExpenseTypeID.Text = "" Then
        MsgBox "Blank Expense Type not Allowed!!!"
        CmbExpenseTypeID.SetFocus
        Exit Sub
    End If

    If LCase(CmdAdd2List.Caption) = LCase("Add to List") Then
        For mR = 1 To MSHFlexGrid2.Rows - 1
            If CmbExpenseTypeID.Text = MSHFlexGrid2.TextMatrix(mR, 1) Then
                MsgBox "Expense Type Already Exist!!!"
                Exit Sub
            End If
        Next
        If CheckExpenseTypeCount(CDbl(Val(TxtGLAccTypeID)), mExpenseTypeID, CmbType.Text) = False Then
            MsgBox "Expense Type - " & CmbExpenseTypeID.Text & " already mapped with other G/L Account Type."
            Exit Sub
        End If
        mR = MSHFlexGrid2.Rows - 1
        MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
    Else
        mRowsel = MSHFlexGrid2.RowSel
        For mR = 1 To MSHFlexGrid2.Rows - 1
            If CmbExpenseTypeID.Text = MSHFlexGrid2.TextMatrix(mR, 1) Then
               If mR <> mRowsel Then
                    MsgBox "Expense Type Already Exist!!!"
                    Exit Sub
               End If
            End If
        Next
        If CheckExpenseTypeCount(CDbl(Val(TxtGLAccTypeID)), mExpenseTypeID, CmbType.Text) = False Then
            MsgBox "Expense Type - " & CmbExpenseTypeID.Text & " already mapped with other G/L Account Type."
            Exit Sub
        End If
        mR = mRowsel
    End If

    MSHFlexGrid2.TextMatrix(mR, 0) = ""
    MSHFlexGrid2.TextMatrix(mR, 1) = CmbExpenseTypeID.Text
    MSHFlexGrid2.TextMatrix(mR, 2) = mExpenseTypeID
    MSHFlexGrid2.TextMatrix(mR, 3) = ""
    
    CmbExpenseTypeID.Text = ""
    TxtExchangeType = ""
    CmdAdd2List.Caption = "Add To List"
    CmdRemoveFromList.Caption = "Remove From List"

End If


End Sub

''-- 12th AUg 2017 Ullhas Patil
Private Sub CmdRemoveFromList_Click()
Dim mR, mRow, mCol As Double
mR = MSHFlexGrid2.RowSel
If MSHFlexGrid2.TextMatrix(mR, 1) <> "" Then
   For mRow = mR To MSHFlexGrid2.Rows - 2
       For mCol = 0 To MSHFlexGrid2.Cols - 1
            MSHFlexGrid2.TextMatrix(mRow, mCol) = MSHFlexGrid2.TextMatrix(mRow + 1, mCol)
       Next
   Next
   MSHFlexGrid2.Rows = MSHFlexGrid2.Rows - 1
End If

CmdAdd2List.Caption = "Add To List"
CmdRemoveFromList.Caption = "Remove From List"

End Sub



Private Sub CmdGo_Click()
tmp = "Select GLAccTypeID,GLAccType,NAVCode,HSNNo,AccountType,GroupType,Flag from Mst_GLAccType "

Dim wc As Variant
wc = ""
If TxtSearchGLAccType <> "" Then
   wc = " Where GLAccType Like '%" & TxtSearchGLAccType & "%' "
End If

If CmbSGroupType.Text <> "" Then
    If wc = "" Then
        wc = " Where GroupType = '" & Left(CmbSGroupType.Text, 1) & "' "
    Else
        wc = wc & " And GroupType = '" & Left(CmbSGroupType.Text, 1) & "' "
    End If
End If
tmp = tmp & wc & " Order By GLAccTypeID "

Call TmpTable
Call Grid_Head
For iu = 1 To TmpRs.RecordCount
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
    For C = 0 To MSHFlexGrid1.Cols - 1
        MSHFlexGrid1.TextMatrix(iu, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
    Next
    MSHFlexGrid1.TextMatrix(iu, 5) = IIf(LCase(Trim(MSHFlexGrid1.TextMatrix(iu, 5))) = "i", "Income", "Expense")
    MSHFlexGrid1.TextMatrix(iu, 6) = IIf(LCase(Trim(MSHFlexGrid1.TextMatrix(iu, 6))) = "a", "Active", "De-Active")
    
    TmpRs.MoveNext
Next
Label22.Refresh
Label22.Caption = TmpRs.RecordCount

End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Call GButtonLockAD(Me)
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If
Call Grid_Head2

Call Grid_Head
TxtSearchGLAccType = ""
Label22.Caption = ""

Call TableMst_GLAccType
If RsMst_GLAccType.RecordCount = 0 Then
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
Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Mst_GLAccType.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Mst_GLAccType.xls")
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
              If C = 2 Then
                 oWS.Cells(mRow, C + 1) = "'" & MSHFlexGrid1.TextMatrix(I, C)
              Else
                 oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
              End If
           End If
        Else
           If C = 2 Then
              oWS.Cells(mRow, C + 1) = "'" & MSHFlexGrid1.TextMatrix(I, C)
           Else
              oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
           End If
        End If
    Next
Next
oWB.Save
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub

Private Sub MSHFlexGrid2_Click()
Dim mR As Double
With MSHFlexGrid2

    If SaveCmd.Enabled = True Then
       mR = .RowSel
       If mR = 0 Then
        If LCase(mGroupType) = LCase("I") Then
            CmbIncomeExpenseTypeID = ""
            TxtExchangeType = ""
        Else
            CmbExpenseTypeID = ""
        End If
        CmdAdd2List.Caption = "Add To List"
       Else
            If .TextMatrix(mR, 1) <> "" Then
                If LCase(mGroupType) = LCase("I") Then
                    CmbIncomeExpenseTypeID.Text = IIf(Trim(.TextMatrix(mR, 1)) = "", "", .TextMatrix(mR, 1))
                    mIncomeExpenseTypeID = Val(.TextMatrix(mR, 2))
                    mExchangeTypeID = Val(.TextMatrix(mR, 3))
                    TxtExchangeType = GetExchangeName(mExchangeTypeID)
                Else
                    CmbExpenseTypeID.Text = IIf(Trim(.TextMatrix(mR, 1)) = "", "", .TextMatrix(mR, 1))
                    mExpenseTypeID = Val(.TextMatrix(mR, 2))
                End If
                CmdAdd2List.Caption = "Update to List"
            Else
                If LCase(mGroupType) = LCase("I") Then
                    CmbIncomeExpenseTypeID = ""
                    TxtExchangeType = ""
                Else
                    CmbExpenseTypeID = ""
                End If
                CmdAdd2List.Caption = "Add To List"
            End If
       End If
    End If

End With
End Sub

Private Sub NextCmd_Click()
Dim LF_NAVCode As Variant
LF_NAVCode = "N"
RsMst_GLAccType.MoveNext
If RsMst_GLAccType.EOF Then
   RsMst_GLAccType.MoveLast
   LF_NAVCode = "Y"
End If
Call Indata
If LF_NAVCode = "Y" Then
    NextCmd.Enabled = False
    LastCmd.Enabled = False
End If
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
End Sub
Private Sub PreviousCmd_Click()
Dim LF_NAVCode As Variant
LF_NAVCode = "N"
RsMst_GLAccType.MovePrevious
If RsMst_GLAccType.BOF Then
   RsMst_GLAccType.MoveFirst
   LF_NAVCode = "Y"
End If
Call Indata
If LF_NAVCode = "Y" Then
    PreviousCmd.Enabled = False
    FirstCmd.Enabled = False
End If
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub FirstCmd_Click()
RsMst_GLAccType.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsMst_GLAccType.MoveLast
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
   If RsMst_GLAccType.RecordCount = 0 Then
    Exit Sub
   Else
    Call Indata
   End If
   Call GButtonLock(Me, True)
   'Call FirstCmd_Click
   Exit Sub
End If

Frame0.Height = Me.Height - 250
Frame0.Left = Frame1.Left ' - 100
Frame0.Top = Frame1.Top ' - 100
Frame0.Width = Me.Width - 600
Frame0.Visible = True

MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 1000
RsMst_GLAccType.MoveFirst

End Sub
Public Sub Indata()

TxtGLAccTypeID.Locked = True
TxtGLAccType.Locked = True
TxtNAVCode.Locked = True
TxtFlag.Locked = True
TxtGLAccTypeID = IIf(IsNull(RsMst_GLAccType!GLAccTypeID), "", RsMst_GLAccType!GLAccTypeID)
TxtGLAccType = IIf(IsNull(RsMst_GLAccType!GLAccType), "", RsMst_GLAccType!GLAccType)
TxtNAVCode = IIf(IsNull(RsMst_GLAccType!NavCode), "", RsMst_GLAccType!NavCode)

TxtCreated = IIf(IsNull(RsMst_GLAccType!CreatedBy), "", RsMst_GLAccType!CreatedBy) & " :- " & IIf(IsNull(RsMst_GLAccType!CreatedDate), "", RsMst_GLAccType!CreatedDate)
TxtUpdated = IIf(IsNull(RsMst_GLAccType!UpdatedBy), "", RsMst_GLAccType!UpdatedBy) & " :- " & IIf(IsNull(RsMst_GLAccType!UpdatedDate), "", RsMst_GLAccType!UpdatedDate)

TxtHSNNo.Locked = True
TxtHSNNo = IIf(IsNull(RsMst_GLAccType!HSNNo), "", RsMst_GLAccType!HSNNo)
CmbType.Visible = False
TxtType.Visible = True
TxtType = IIf(IsNull(RsMst_GLAccType!AccountType), "", RsMst_GLAccType!AccountType)
CmbType.Text = TxtType

CmbFlag.Visible = False
TxtFlag.Visible = True
TxtFlag = IIf(IsNull(RsMst_GLAccType!Flag), "", RsMst_GLAccType!Flag)
If LCase(TxtFlag) = "a" Then
   TxtFlag = "Active"
ElseIf LCase(TxtFlag) = "d" Then
   TxtFlag = "De-Active"
End If
CmbFlag.Text = TxtFlag

''-- Ullhas Patil 28 Sep 2017
mGroupType = IIf(IsNull(RsMst_GLAccType!GroupType), "I", RsMst_GLAccType!GroupType)
CmbGroupType.Text = ""
TxtGroupType = IIf(LCase(mGroupType) = LCase("I"), "Income", "Expense")
TxtGroupType.Visible = True
CmbGroupType.Visible = False

If LCase(mGroupType) = LCase("I") Then
    ''--- 12th Aug 2017 Ullhas Patil
    Label6.Visible = False
    CmbExpenseTypeID.Visible = False
    Label12.Visible = True
    CmbIncomeExpenseTypeID.Visible = True
    Label4.Visible = True
    TxtExchangeType.Visible = True
    Call IndataGLAccountIncomeTypeMapping
Else
    Label6.Visible = True
    CmbExpenseTypeID.Visible = True
    Label12.Visible = False
    CmbIncomeExpenseTypeID.Visible = False
    Label4.Visible = False
    TxtExchangeType.Visible = False
    Call IndataGLAccountExpenseTypeMapping
End If
''-----------------
Frame2.Enabled = False

If RsMst_GLAccType.RecordCount = 1 Then
   Call GButtonLock_1(Me, True)
Else
   Call GButtonLock(Me, True)
End If
End Sub

Private Sub IndataGLAccountExpenseTypeMapping()
Dim mR As Double
Call Grid_Head2
CmdAdd2List.Caption = "Add To List"
CmdRemoveFromList.Caption = "Remove From List"
CmbExpenseTypeID.Text = ""
TxtExchangeType = ""

With MSHFlexGrid2
tmp = "SELECT MGAIM.GLAccountIncomeTypeMappingID,MGAIM.ExpenseTypeID,MET.ExpenseType"
tmp = tmp & " FROM Mst_GLAccountIncomeTypeMapping MGAIM"
tmp = tmp & " Inner Join Mst_ExpenseType MET On MGAIM.ExpenseTypeID = MET.ExpenseTypeID"
tmp = tmp & " Where MGAIM.GLAccTypeID = " & Val(TxtGLAccTypeID.Text) & ""
Call Tmp1Table

If Not TmpRs1.EOF Then
    For mR = 1 To TmpRs1.RecordCount
        .TextMatrix(mR, 0) = mR
        .TextMatrix(mR, 1) = IIf(IsNull(TmpRs1!ExpenseType), "", TmpRs1!ExpenseType)
        .TextMatrix(mR, 2) = IIf(IsNull(TmpRs1!ExpenseTypeID), "", TmpRs1!ExpenseTypeID)
        .TextMatrix(mR, 3) = ""
        .TextMatrix(mR, 4) = TmpRs1!GLAccountIncomeTypeMappingID
        TmpRs1.MoveNext
        .Rows = .Rows + 1
    Next mR
End If
End With

End Sub

Private Sub IndataGLAccountIncomeTypeMapping()
Dim mR As Double
Call Grid_Head2
CmdAdd2List.Caption = "Add To List"
CmdRemoveFromList.Caption = "Remove From List"
CmbIncomeExpenseTypeID.Text = ""
TxtExchangeType = ""

With MSHFlexGrid2
tmp = "SELECT MGAIM.GLAccountIncomeTypeMappingID,MGAIM.GLAccTypeID,MGAIM.IETypeID,MIET.IEType,MIET.IEFlag,MIET.ExchangeTypeID"
tmp = tmp & " FROM Mst_GLAccountIncomeTypeMapping MGAIM"
tmp = tmp & " Inner Join Mst_IncomeExpenseType MIET On MGAIM.IETypeID = MIET.IETypeID"
tmp = tmp & " Where MIET.IEFlag = 'I' And MGAIM.GLAccTypeID = " & Val(TxtGLAccTypeID.Text) & ""
Call Tmp1Table
If Not TmpRs1.EOF Then
    For mR = 1 To TmpRs1.RecordCount
        .TextMatrix(mR, 0) = mR
        .TextMatrix(mR, 1) = IIf(IsNull(TmpRs1!IEType), "", TmpRs1!IEType)
        .TextMatrix(mR, 2) = IIf(IsNull(TmpRs1!IETypeID), "", TmpRs1!IETypeID)
        .TextMatrix(mR, 3) = IIf(IsNull(TmpRs1!ExchangeTypeID), "", TmpRs1!ExchangeTypeID)
        .TextMatrix(mR, 4) = TmpRs1!GLAccountIncomeTypeMappingID
        TmpRs1.MoveNext
        .Rows = .Rows + 1
    Next mR
End If
End With
End Sub

Private Sub TxtNAVCode_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then
   Call AssignHelpData(Me, "TxtNAVCode")
End If
End Sub

Private Sub TxtGLAccTypeID_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then
   Call AssignHelpData(Me, "TxtGLAccTypeID")
End If

End Sub

Private Sub TxtGLAccTypeID_LostFocus()
If TxtGLAccTypeID = "" Then Exit Sub
If IsNumeric(TxtGLAccTypeID) = False Then
   MsgBox "Invalid numeric format !!!!"
   TxtGLAccTypeID.SetFocus
   Exit Sub
End If
TxtGLAccTypeID = Val(TxtGLAccTypeID)
End Sub
Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)

TxtGLAccTypeID.Locked = True
TxtGLAccType.Locked = False
TxtNAVCode.Locked = False
TxtGLAccTypeID = ""
TxtGLAccType = ""
TxtHSNNo.Locked = False
TxtHSNNo = ""
CmbType.Visible = True
TxtType.Visible = False
CmbType.Text = ""
CmbFlag.Visible = True
TxtFlag.Visible = False
CmbFlag.Text = ""

TxtCreated = ""
TxtUpdated = ""
TxtNAVCode = ""

''-- 12th Aug 2017 Ullhas Patil
CmbGroupType.Text = ""
CmbGroupType.Visible = True
TxtGroupType = ""
TxtGroupType.Visible = False
mGroupType = "I"

If LCase(mGroupType) = LCase("I") Then
    Label12.Visible = True
    CmbIncomeExpenseTypeID.Text = ""
    CmbIncomeExpenseTypeID.Visible = True
    Label4.Visible = True
    TxtExchangeType.Visible = True
Else
    Label6.Visible = False
    CmbExpenseTypeID.Text = ""
    CmbExpenseTypeID.Visible = False
    Label4.Visible = False
    TxtExchangeType.Visible = False
End If

Call Grid_Head2
Frame2.Enabled = True

If RsMst_GLAccType.RecordCount > 0 Then
   TxtGLAccType.SetFocus
End If
End Sub
Private Sub EditCmd_Click()
Edit_Flg = "E"
Call GButtonLock(Me, False)

TxtGLAccTypeID.Locked = True
TxtGLAccType.Locked = False
TxtNAVCode.Locked = False
TxtHSNNo.Locked = False
CmbType.Visible = True
TxtType.Visible = False
CmbType.Text = TxtType
CmbFlag.Visible = True
TxtFlag.Visible = False
CmbFlag.Text = TxtFlag
Frame2.Enabled = True

''-- Ullhas Patil 29 Sep 2017
CmbGroupType = TxtGroupType

End Sub

Private Sub DeleteCmd_Click()
Dim ans As Variant
On Error GoTo msgError
ans = MsgBox("Do you really want to DELETE this record???", vbYesNo)
If ans = vbYes Then

   ''-- Delete Mapped Income Type First
   tmp = " DELETE FROM Mst_GLAccountIncomeTypeMapping WHERE GLAccTypeID = " & Val(TxtGLAccTypeID) & ""
   Call TmpTable

   RsMst_GLAccType.Delete
   RsMst_GLAccType.Update
   If RsMst_GLAccType.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_GLAccType.MoveNext
   If RsMst_GLAccType.EOF() Then
      RsMst_GLAccType.MoveLast
   End If
   Call Indata
End If
Exit Sub
msgError:
MsgBox "Child record Present "
RsMst_GLAccType.CancelUpdate
End Sub
Private Sub SaveCmd_Click()
Dim mR As Double

If TxtGLAccType.Text = "" Then
   MsgBox "Blank GLAccType Not Allowed !!! "
   TxtGLAccType.SetFocus
   Exit Sub
End If
If TxtNAVCode.Text = "" Then
   MsgBox "Blank NAV Code Not Allowed !!! "
   TxtNAVCode.SetFocus
   Exit Sub
End If
If TxtHSNNo.Text = "" Then
   MsgBox "Blank HSN No Not Allowed !!! "
   TxtHSNNo.SetFocus
   Exit Sub
End If
If CmbType.Text = "" Then
   MsgBox "Blank Type Not Allowed !!! "
   CmbType.SetFocus
   Exit Sub
End If
''-- Ullhas Patil 29 Sep 2017
If CmbGroupType.Text = "" Then
   MsgBox "Blank Group Type Not Allowed !!! "
   CmbGroupType.SetFocus
   Exit Sub
End If
''----
If CmbFlag.Text = "" Then
   MsgBox "Blank Flag Not Allowed !!! "
   CmbFlag.SetFocus
   Exit Sub
End If

''-- 12th Aug 2017 Ullhas Patil
If MSHFlexGrid2.Rows <= 2 Then
   ''MsgBox "Blank Flag Not Allowed !!! "
   ''CmbFlag.SetFocus
   ''Exit Sub
Else
    If LCase(CmbGroupType.Text) = "income" Then
       For mR = 1 To MSHFlexGrid2.Rows - 2
           If CheckIncomeTypeCount(CDbl(Val(TxtGLAccTypeID)), CDbl(MSHFlexGrid2.TextMatrix(mR, 2)), CmbType.Text) = False Then
               MsgBox "Income Type - " & MSHFlexGrid2.TextMatrix(mR, 1) & " already mapped with other G/L Account Type."
               Exit Sub
           End If
       Next mR
    Else
       For mR = 1 To MSHFlexGrid2.Rows - 2
           If CheckExpenseTypeCount(CDbl(Val(TxtGLAccTypeID)), CDbl(MSHFlexGrid2.TextMatrix(mR, 2)), CmbType.Text) = False Then
              MsgBox "Expense Type - " & MSHFlexGrid2.TextMatrix(mR, 1) & " already mapped with other G/L Account Type."
              Exit Sub
           End If
    Next mR
    
    End If
End If
''-- Ullhas Patil 29 Sep 2017
If MSHFlexGrid2.Rows > 2 Then
End If
''-----------------------------

If Edit_Flg = "A" Then
   If RsMst_GLAccType.RecordCount > 0 Then
      RsMst_GLAccType.MoveFirst
      RsMst_GLAccType.Find "GLAccType= '" & TxtGLAccType.Text & "'"
      If Not RsMst_GLAccType.EOF Then
        MsgBox "Duplicate GL Acc Type Not Allowed !!! "
         TxtGLAccType.SetFocus
         Exit Sub
      End If
   End If
      RsMst_GLAccType.AddNew
      RsMst_GLAccType!GLAccTypeID = GetGLAccTypeID
      RsMst_GLAccType!G_UID = GetGUID
      TxtGLAccTypeID.Text = RsMst_GLAccType!GLAccTypeID
      RsMst_GLAccType!GroupType = mGroupType
Else
    RsMst_GLAccType.MoveFirst
    RsMst_GLAccType.Find "GLAccType= '" & TxtGLAccType.Text & "'"
    If Not RsMst_GLAccType.EOF Then
       If RsMst_GLAccType!GLAccTypeID <> TxtGLAccTypeID.Text Then
          MsgBox "Duplicate GL Acc Type Not Allowed !!! "
          TxtGLAccType.SetFocus
          Exit Sub
       End If
    End If
    RsMst_GLAccType.MoveFirst
    RsMst_GLAccType.Find " GLAccTypeID= " & TxtGLAccTypeID.Text
End If

RsMst_GLAccType!GLAccTypeID = TxtGLAccTypeID.Text
RsMst_GLAccType!GLAccType = TxtGLAccType.Text
RsMst_GLAccType!NavCode = TxtNAVCode.Text
RsMst_GLAccType!HSNNo = Trim(TxtHSNNo.Text)
RsMst_GLAccType!AccountType = Trim(CmbType.Text)
RsMst_GLAccType!Flag = Left(CmbFlag.Text, 1)

If IsNull(RsMst_GLAccType!CreatedBy) = True Or RsMst_GLAccType!CreatedBy = "" Then
   RsMst_GLAccType!CreatedBy = Gen_UserLoginID
   RsMst_GLAccType!CreatedDate = Now
Else
   RsMst_GLAccType!UpdatedBy = Gen_UserLoginID
   RsMst_GLAccType!UpdatedDate = Now
End If
RsMst_GLAccType.Update

Call SaveInComeType

RsMst_GLAccType.Requery
RsMst_GLAccType.Find "GLAccTypeID = " & Val(TxtGLAccTypeID)
Call Indata
End Sub

''-- 12 AUg 2017 Ullhas Patil
Sub SaveInComeType()
Dim mMappingID As Variant
Dim I As Double
With MSHFlexGrid2
    If .Rows > 2 Then
    
       Call TableMst_GLAccountIncomeTypeMapping(" Where GLAccTypeID = " & Val(TxtGLAccTypeID.Text) & " Order By GLAccountIncomeTypeMappingID ")
       For I = 1 To .Rows - 2
           If RsMst_GLAccountIncomeTypeMapping.RecordCount > 0 Then
            RsMst_GLAccountIncomeTypeMapping.MoveFirst
           End If
           RsMst_GLAccountIncomeTypeMapping.Find ("GLAccountIncomeTypeMappingID =" & Val(.TextMatrix(I, 4)))
           If RsMst_GLAccountIncomeTypeMapping.EOF = True Then
                RsMst_GLAccountIncomeTypeMapping.AddNew
                RsMst_GLAccountIncomeTypeMapping!GLAccountIncomeTypeMappingID = GetMaxID("GLAccountIncomeTypeMappingID", "Mst_GLAccountIncomeTypeMapping")
                RsMst_GLAccountIncomeTypeMapping!G_UID = GetGUID
                .TextMatrix(I, 4) = RsMst_GLAccountIncomeTypeMapping!GLAccountIncomeTypeMappingID
                RsMst_GLAccountIncomeTypeMapping!GLAccTypeID = Val(TxtGLAccTypeID)
           End If
           
           ''-- Ullhas Patil 29 Sep 2017
           If LCase(mGroupType) = LCase("I") Then
                RsMst_GLAccountIncomeTypeMapping!IETypeID = Val(.TextMatrix(I, 2))
           Else
                RsMst_GLAccountIncomeTypeMapping!ExpenseTypeID = Val(.TextMatrix(I, 2))
           End If
           ''--------
           
           If IsNull(RsMst_GLAccountIncomeTypeMapping!CreatedBy) = True Or RsMst_GLAccountIncomeTypeMapping!CreatedBy = "" Then
               RsMst_GLAccountIncomeTypeMapping!CreatedBy = Gen_UserLoginID
               RsMst_GLAccountIncomeTypeMapping!CreatedDate = Now
           Else
               RsMst_GLAccountIncomeTypeMapping!UpdatedBy = Gen_UserLoginID
               RsMst_GLAccountIncomeTypeMapping!UpdatedDate = Now
           End If
           RsMst_GLAccountIncomeTypeMapping.Update
           RsMst_GLAccountIncomeTypeMapping.Requery
           
       Next I
       mMappingID = ""
       ''---- Delete Details
       For I = 1 To .Rows - 2
           mMappingID = mMappingID & Val(.TextMatrix(I, 4)) & ","
       Next I
       If mMappingID <> "" Then
          mMappingID = Left(mMappingID, Len(mMappingID) - 1)
       End If
       
       tmp = " DELETE FROM Mst_GLAccountIncomeTypeMapping WHERE GLAccTypeID = " & Val(TxtGLAccTypeID) & " AND GLAccountIncomeTypeMappingID NOT IN ( " & mMappingID & " )"
       Call TmpTable
       ''------------------------
       
    End If
End With
End Sub
''-----------------------------

Private Function GetGLAccTypeID() As Double
Dim Retval As Double
tmp = "Select max(GLAccTypeID) from Mst_GLAccType"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetGLAccTypeID = Retval
End Function
Private Sub Grid_Head()
With MSHFlexGrid1

    .Clear
    .Rows = 2
    .Cols = 7
    .TextMatrix(0, 0) = "GLAccTypeID":      .ColWidth(0) = 1500
    .TextMatrix(0, 1) = "GL Acc Type":      .ColWidth(1) = 3500
    .TextMatrix(0, 2) = "NAV Code":         .ColWidth(2) = 2500
    .TextMatrix(0, 3) = "HSN No":           .ColWidth(3) = 1000
    .TextMatrix(0, 4) = "Account Type":     .ColWidth(4) = 1800
    .TextMatrix(0, 5) = "Group Type":       .ColWidth(5) = 1800
    .TextMatrix(0, 6) = "Flag":             .ColWidth(6) = 1800

End With
End Sub

Private Sub Grid_Head2()
With MSHFlexGrid2
.Clear
.Rows = 2
.Cols = 5
.FixedCols = 1
.FixedRows = 1
.TextMatrix(0, 0) = "Sr No":                .ColWidth(0) = 800
If LCase(mGroupType) = LCase("I") Then
    .TextMatrix(0, 1) = "Income Type":      .ColWidth(1) = 5000
    .TextMatrix(0, 2) = "IncomeTypeID":     .ColWidth(2) = 0
    .TextMatrix(0, 3) = "ExchangeTypeID":   .ColWidth(3) = 0
    .TextMatrix(0, 4) = "Mapping ID":       .ColWidth(4) = 0
Else
    .TextMatrix(0, 1) = "Expense Type":     .ColWidth(1) = 5000
    .TextMatrix(0, 2) = "ExpenseTypeID":    .ColWidth(2) = 0
    .TextMatrix(0, 3) = "":                 .ColWidth(3) = 0
    .TextMatrix(0, 4) = "Mapping ID":       .ColWidth(4) = 0
End If

End With
End Sub

Private Sub MSHFlexGrid1_Click()
Dim I As Variant
I = MSHFlexGrid1.RowSel
If I <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(I, 0) = "" Then Exit Sub
RsMst_GLAccType.MoveFirst
RsMst_GLAccType.Find "GLAccTypeID = " & MSHFlexGrid1.TextMatrix(I, 0)
If Not RsMst_GLAccType.EOF Then
   Call Indata
   Frame0.Visible = False
End If
End Sub

Private Sub TxtGLAccType_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "TxtGLAccType")
End If
End Sub

