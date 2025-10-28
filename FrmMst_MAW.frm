VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmMst_MAW 
   Caption         =   "MAW Master"
   ClientHeight    =   7515
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8145
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7515
   ScaleWidth      =   8145
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   2295
      Left            =   120
      TabIndex        =   45
      Top             =   9600
      Visible         =   0   'False
      Width           =   14895
      Begin VB.ComboBox CmbSStatusFlag 
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
         ItemData        =   "FrmMst_MAW.frx":0000
         Left            =   5880
         List            =   "FrmMst_MAW.frx":000A
         Sorted          =   -1  'True
         TabIndex        =   27
         Top             =   240
         Width           =   1695
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
         Left            =   9645
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   270
         Width           =   1000
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
         Left            =   8640
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   270
         Width           =   1005
      End
      Begin VB.ComboBox CmbSCommodityID 
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
         ItemData        =   "FrmMst_MAW.frx":0020
         Left            =   1200
         List            =   "FrmMst_MAW.frx":0022
         Sorted          =   -1  'True
         TabIndex        =   26
         Top             =   240
         Width           =   4095
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1455
         Left            =   120
         TabIndex        =   30
         Top             =   720
         Width           =   14655
         _ExtentX        =   25850
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
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Flag"
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
         Left            =   5400
         TabIndex        =   49
         Top             =   300
         Width           =   405
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
         Left            =   13275
         TabIndex        =   48
         Top             =   300
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
         Left            =   10845
         TabIndex        =   47
         Top             =   300
         Width           =   2355
      End
      Begin VB.Label Label13 
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
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   300
         Width           =   1215
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1095
      Left            =   120
      TabIndex        =   44
      Top             =   7800
      Width           =   7215
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
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   240
         Width           =   1350
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
         Left            =   1575
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   240
         Width           =   1335
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
         Left            =   2910
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   240
         Width           =   1335
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
         Left            =   4245
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   240
         Width           =   1335
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
         Left            =   5580
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   240
         Width           =   1335
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
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   615
         Width           =   1350
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
         Left            =   1575
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   615
         Width           =   1335
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
         Left            =   2910
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   615
         Width           =   1335
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
         Left            =   4245
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   615
         Width           =   1335
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
         Left            =   5580
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   615
         Width           =   1335
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHGrid_MAW 
      Height          =   3420
      Left            =   240
      TabIndex        =   15
      Top             =   4320
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   6033
      _Version        =   393216
      Appearance      =   0
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
      Height          =   4695
      Left            =   120
      TabIndex        =   40
      Top             =   3120
      Width           =   7215
      Begin VB.TextBox TxtMinimum 
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
         Left            =   1035
         MaxLength       =   10
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox TxtMaximum 
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
         Left            =   3150
         MaxLength       =   10
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox TxtMAWD 
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
         Left            =   6120
         MaxLength       =   10
         TabIndex        =   11
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton CmdAddtoList 
         Caption         =   "Add to List"
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
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   720
         Width           =   1575
      End
      Begin VB.CommandButton CmdDeletefromList 
         Caption         =   "Delete"
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
         Left            =   4575
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   720
         Width           =   1335
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
         Left            =   5910
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Minimum"
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
         TabIndex        =   43
         Top             =   300
         Width           =   915
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Maximum"
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
         Left            =   2160
         TabIndex        =   42
         Top             =   300
         Width           =   975
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Final MAWD in %"
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
         Left            =   4320
         TabIndex        =   41
         Top             =   300
         Width           =   1785
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3135
      Left            =   120
      TabIndex        =   31
      Top             =   0
      Width           =   7215
      Begin VB.TextBox TxtStatusFlag 
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
         Left            =   2280
         TabIndex        =   5
         Top             =   1680
         Width           =   1935
      End
      Begin VB.ComboBox CmbStatusFlag 
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
         ItemData        =   "FrmMst_MAW.frx":0024
         Left            =   2280
         List            =   "FrmMst_MAW.frx":002E
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   1680
         Width           =   1935
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
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   2625
         Width           =   4815
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
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   2160
         Width           =   4815
      End
      Begin VB.ComboBox CmbCommodityID 
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
         ItemData        =   "FrmMst_MAW.frx":0044
         Left            =   2280
         List            =   "FrmMst_MAW.frx":0046
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   720
         Width           =   4815
      End
      Begin VB.TextBox TxtMAWID 
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
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox TxtMAWMinimum 
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
         Left            =   3240
         MaxLength       =   10
         TabIndex        =   3
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox TxtMAWMaximum 
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
         Left            =   6120
         MaxLength       =   10
         TabIndex        =   4
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox TxtCommodityID 
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
         TabIndex        =   2
         Top             =   720
         Width           =   4815
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
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
         TabIndex        =   39
         Top             =   2685
         Width           =   1695
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
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
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   2220
         Width           =   1695
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status Flag"
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
         TabIndex        =   37
         Top             =   1740
         Width           =   1185
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "MAW ID"
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
         TabIndex        =   36
         Top             =   300
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Commodity"
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
         TabIndex        =   35
         Top             =   780
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Ready Reckoner Moisture Range"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   34
         Top             =   1140
         Width           =   1935
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Minimum"
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
         Left            =   2280
         TabIndex        =   33
         Top             =   1260
         Width           =   915
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Maximum"
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
         Left            =   5040
         TabIndex        =   32
         Top             =   1260
         Width           =   975
      End
   End
End
Attribute VB_Name = "FrmMst_MAW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mCommodityID As Double, mSCommodityID As Double
Dim Edit_Flg As String

Private Sub AddCmd_Click()
Me.MousePointer = vbHourglass

    Edit_Flg = "A"
    Call GButtonLock(Me, False)

    TxtMAWID.Text = ""
    CmbCommodityID.Text = "": CmbCommodityID.Visible = True
    TxtCommodityID.Text = "": TxtCommodityID.Locked = True: TxtCommodityID.Visible = False
    TxtMAWMinimum.Text = "":  TxtMAWMinimum.Locked = False
    TxtMAWMaximum.Text = "":  TxtMAWMaximum.Locked = False
    CmbStatusFlag.Text = "":  CmbStatusFlag.Visible = True
    TxtStatusFlag.Text = "":  TxtStatusFlag.Locked = True: TxtStatusFlag.Visible = False
    TxtCreated.Text = "":     TxtCreated.Locked = True
    TxtUpdated.Text = "":     TxtUpdated.Locked = True
    
    Call Grid_Header
    Call ClearMAWDetails
    Frame1.Enabled = True
    Frame2.Enabled = True
    
Me.MousePointer = vbDefault
End Sub

Private Sub ClearMAWDetails()
Me.MousePointer = vbHourglass

    TxtMinimum.Text = "":    TxtMinimum.Locked = False
    TxtMaximum.Text = "":    TxtMaximum.Locked = False
    TxtMAWD.Text = "":       TxtMAWD.Locked = False
    CmdAddtoList.Caption = "Add to List"

Me.MousePointer = vbDefault
End Sub

Private Sub CmbCommodityID_GotFocus()
Me.MousePointer = vbHourglass
On Error GoTo Error

tmp = CmbCommodityID.Text
Call TableMst_Commodity(" Where MAWFlag = '1' Order By Commodity ")
CmbCommodityID.Clear
If RsMst_Commodity.RecordCount = 0 Then
   Me.MousePointer = vbDefault
   MsgBox " No Records Found !!"
   Exit Sub
End If
For i = 1 To RsMst_Commodity.RecordCount
    CmbCommodityID.AddItem RsMst_Commodity!Commodity
    RsMst_Commodity.MoveNext
Next
CmbCommodityID.Text = tmp

Error:
    If Err.Number <> 0 Then
        Me.MousePointer = vbDefault
        MsgBox "Error Number : " & Err.Number & vbNewLine & "Error Description : " & Err.Description, vbOKOnly + vbInformation, App.Title
    End If
Me.MousePointer = vbDefault
End Sub

Private Sub CmbCommodityID_LostFocus()
Me.MousePointer = vbHourglass
On Error GoTo Error

If CmbCommodityID.Text = "" Then Exit Sub
RsMst_Commodity.MoveFirst
RsMst_Commodity.Find "Commodity='" & CmbCommodityID.Text & "'"
If RsMst_Commodity.EOF Then
   Me.MousePointer = vbDefault
   MsgBox "Invalid Selection."
   CmbCommodityID.SetFocus
   Exit Sub
Else
    mCommodityID = RsMst_Commodity!CommodityID
End If

Error:
    If Err.Number <> 0 Then
        Me.MousePointer = vbDefault
        MsgBox "Error Number : " & Err.Number & vbNewLine & "Error Description : " & Err.Description, vbOKOnly + vbInformation, App.Title
    End If
Me.MousePointer = vbDefault
End Sub

Private Sub CmbSCommodityID_GotFocus()
Me.MousePointer = vbHourglass
On Error GoTo Error

tmp = CmbSCommodityID.Text
Call TableMst_Commodity(" Where MAWFlag = '1' Order By Commodity ")
CmbSCommodityID.Clear
If RsMst_Commodity.RecordCount = 0 Then
   Me.MousePointer = vbDefault
   MsgBox " No Records Found !!"
   Exit Sub
End If
For i = 1 To RsMst_Commodity.RecordCount
    CmbSCommodityID.AddItem RsMst_Commodity!Commodity
    RsMst_Commodity.MoveNext
Next
CmbSCommodityID.Text = tmp

Error:
    If Err.Number <> 0 Then
        Me.MousePointer = vbDefault
        MsgBox "Error Number : " & Err.Number & vbNewLine & "Error Description : " & Err.Description, vbOKOnly + vbInformation, App.Title
    End If
Me.MousePointer = vbDefault
End Sub

Private Sub CmbSCommodityID_LostFocus()
Me.MousePointer = vbHourglass
On Error GoTo Error

If CmbSCommodityID.Text = "" Then: Me.MousePointer = vbDefault: Exit Sub
RsMst_Commodity.MoveFirst
RsMst_Commodity.Find "Commodity='" & CmbSCommodityID.Text & "'"
If RsMst_Commodity.EOF Then
   Me.MousePointer = vbDefault
   MsgBox "Invalid Selection."
   CmbSCommodityID.SetFocus
   Exit Sub
Else
    mSCommodityID = RsMst_Commodity!CommodityID
End If

Error:
    If Err.Number <> 0 Then
        Me.MousePointer = vbDefault
        MsgBox "Error Number : " & Err.Number & vbNewLine & "Error Description : " & Err.Description, vbOKOnly + vbInformation, App.Title
    End If
Me.MousePointer = vbDefault
End Sub

Private Sub CmbSStatusFlag_LostFocus()
If CmbSStatusFlag.Text = "" Then Exit Sub

If LCase(CmbSStatusFlag.Text) = "active" Or LCase(CmbSStatusFlag.Text) = "deactive" Then
Else
   MsgBox "Invalid selection!!!"
   CmbSStatusFlag.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmbStatusFlag_LostFocus()
If CmbStatusFlag.Text = "" Then Exit Sub

If LCase(CmbStatusFlag.Text) = "active" Or LCase(CmbStatusFlag.Text) = "deactive" Then
Else
   MsgBox "Invalid selection!!!"
   CmbStatusFlag.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmdAddtoList_Click()
Me.MousePointer = vbHourglass

    If SaveCmd.Enabled = False Then
       Me.MousePointer = vbDefault
       Exit Sub
    End If

    If CmbCommodityID.Text = "" Then
        Me.MousePointer = vbDefault
        MsgBox "Blank Commodity Not Allowed.", vbOKOnly + vbInformation, App.Title
        CmbCommodityID.SetFocus
        Exit Sub
    End If
    
    If TxtMAWMinimum.Text = "" Then
        Me.MousePointer = vbDefault
        MsgBox "Blank Ready Reckoner Moisture Minimum Range Not Allowed.", vbOKOnly + vbInformation, App.Title
        TxtMAWMinimum.SetFocus
        Exit Sub
    End If
    If TxtMAWMaximum.Text = "" Then
        Me.MousePointer = vbDefault
        MsgBox "Blank Ready Reckoner Moisture Maximum Range Not Allowed.", vbOKOnly + vbInformation, App.Title
        TxtMAWMaximum.SetFocus
        Exit Sub
    End If

    If TxtMinimum.Text = "" Then
        Me.MousePointer = vbDefault
        MsgBox "Blank Minimum Not Allowed.", vbOKOnly + vbInformation, App.Title
        TxtMinimum.SetFocus
        Exit Sub
    End If
    If TxtMaximum.Text = "" Then
        Me.MousePointer = vbDefault
        MsgBox "Blank Maximum Not Allowed.", vbOKOnly + vbInformation, App.Title
        TxtMaximum.SetFocus
        Exit Sub
    End If
    If TxtMAWD.Text = "" Then
        Me.MousePointer = vbDefault
        MsgBox "Blank MAWD Not Allowed.", vbOKOnly + vbInformation, App.Title
        TxtMAWD.SetFocus
        Exit Sub
    End If
    
    If Val(TxtMinimum) < Val(TxtMAWMinimum) Then
        Me.MousePointer = vbDefault
        MsgBox "Minimum Value can not less then Ready Reckoner Moisture Minimum Range.", vbOKOnly + vbInformation, App.Title
        TxtMinimum.SetFocus
        Exit Sub
    End If
    
    If Val(TxtMinimum) > Val(TxtMAWMaximum) Then
        Me.MousePointer = vbDefault
        MsgBox "Minimum Value can not more then Ready Reckoner Moisture Maximum Range.", vbOKOnly + vbInformation, App.Title
        TxtMinimum.SetFocus
        Exit Sub
    End If
    If Val(TxtMaximum) < Val(TxtMAWMinimum) Then
        Me.MousePointer = vbDefault
        MsgBox "Maximum Value can not less then Ready Reckoner Moisture Minimum Range.", vbOKOnly + vbInformation, App.Title
        TxtMaximum.SetFocus
        Exit Sub
    End If
    If Val(TxtMaximum) > Val(TxtMAWMaximum) Then
        Me.MousePointer = vbDefault
        MsgBox "Maximum Value can not more then Ready Reckoner Moisture Maximum Range.", vbOKOnly + vbInformation, App.Title
        TxtMaximum.SetFocus
        Exit Sub
    End If
    
    If Val(TxtMinimum) > Val(TxtMaximum) Then
        Me.MousePointer = vbDefault
        MsgBox "Minimum Value can not more then Maximum Value.", vbOKOnly + vbInformation, App.Title
        TxtMinimum.SetFocus
        Exit Sub
    End If
    
    If CheckRangeValue(TxtMinimum, TxtMaximum, 1) = True Then
            Me.MousePointer = vbDefault
            MsgBox "Duplicate Range Not Allowed.", vbOKOnly + vbInformation, App.Title
            TxtMinimum.SetFocus
            Exit Sub
    End If
    
    If LCase(CmdAddtoList.Caption) = LCase("Add to List") Then
       i = MSHGrid_MAW.Rows - 1
       MSHGrid_MAW.Rows = MSHGrid_MAW.Rows + 1
    Else
       i = MSHGrid_MAW.RowSel
    End If
    
    With MSHGrid_MAW
        .TextMatrix(i, 0) = i
        .TextMatrix(i, 1) = Format(TxtMinimum.Text, "#0.000")
        .TextMatrix(i, 2) = Format(TxtMaximum.Text, "#0.000")
        .TextMatrix(i, 3) = Format(TxtMAWD.Text, "#0.00")
        .TextMatrix(i, 4) = IIf(Trim(.TextMatrix(i, 4)) <> "0", .TextMatrix(i, 4), "0")
    End With
    Call ClearMAWDetails

Me.MousePointer = vbDefault
End Sub

Private Sub CmdDeleteFromList_Click()
Dim i As Double, RO As Double, C As Double
    With MSHGrid_MAW
     If TxtMinimum.Text <> "" And TxtMaximum <> "" Then
        i = .RowSel
        If i > 0 Then
           If .TextMatrix(i, 0) = "" Then Exit Sub
              If .TextMatrix(i, 1) <> "" Then
                 For RO = i To .Rows - 2
                     .row = i
                     For C = 0 To .Cols - 1
                         .TextMatrix(RO, C) = .TextMatrix(RO + 1, C)
                     Next
                 Next
                 .Rows = .Rows - 1
              End If
        End If
        .RowSel = 0
        Call ClearMAWDetails
     Else
        MsgBox "Select the Particular row to 'Delete'.", vbOKOnly + vbInformation, App.Title
     End If
    End With
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Mst_MAW.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Mst_MAW.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For i = 0 To MSHFlexGrid1.Rows - 1
    mRow = i + 1 ' + 6
    For C = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(i, C), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(i, C)) Then
              oWS.Cells(mRow, C + 1) = Format(MSHFlexGrid1.TextMatrix(i, C), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(i, C)
           End If
        Else
           oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(i, C)
        End If
    Next
Next
oWB.Save
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub


Private Sub CmdGo_Click()
Dim wc As Variant
Dim i As Double

wc = ""
Call Grid_Head

tmp = " SELECT MM.MAWID,MM.CommodityID,MC.Commodity,MM.MinimumRange,MM.MaximumRange,MM.Flag from Mst_MAW AS MM"
tmp = tmp & " Inner Join Mst_Commodity AS MC On MM.CommodityID = MC.CommodityID"

If CmbSCommodityID.Text <> "" Then
   If wc = "" Then
      wc = " Where MM.CommodityID = " & mSCommodityID
   End If
End If

If CmbSStatusFlag.Text <> "" Then
   If wc = "" Then
    wc = " Where MM.Flag ='" & Left$(CmbSStatusFlag.Text, 1) & "'"
   Else
    wc = wc & " And MM.Flag ='" & Left$(CmbSStatusFlag.Text, 1) & "'"
   End If
End If

tmp = tmp & wc & " Order By MM.MAWID"
Call TmpTable

For i = 1 To TmpRs.RecordCount
    With MSHFlexGrid1
        .TextMatrix(i, 0) = IIf(IsNull(TmpRs!MAWID), "", TmpRs!MAWID)
        .TextMatrix(i, 1) = IIf(IsNull(TmpRs!Commodity), "", TmpRs!Commodity)
        .TextMatrix(i, 2) = IIf(IsNull(TmpRs!MinimumRange), "", Format(TmpRs!MinimumRange, "#0.000"))
        .TextMatrix(i, 3) = IIf(IsNull(TmpRs!MaximumRange), "", Format(TmpRs!MaximumRange, "#0.000"))
        .TextMatrix(i, 4) = IIf(TmpRs!Flag = "A", "Active", "DeActive")
        .Rows = .Rows + 1
        TmpRs.MoveNext
    End With
Next
    Label22.Caption = TmpRs.RecordCount
End Sub

Private Sub CmdClear_Click()
    Call ClearMAWDetails
End Sub

Private Sub DeleteCmd_Click()
Dim ans As Variant
On Error GoTo msgError
ans = MsgBox("Do you really want to DELETE this record???", vbYesNo)
If ans = vbYes Then

    ''---- Delete MAWD Details
    tmp = " DELETE FROM Mst_MAWDetails WHERE MAWID = " & TxtMAWID & ""
    Call TmpTable
    ''------------------------

   RsMst_MAW.Delete
   RsMst_MAW.Update
   If RsMst_MAW.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_MAW.MoveNext
   If RsMst_MAW.EOF() Then
      RsMst_MAW.MoveLast
   End If
   Call Indata
End If
Exit Sub
msgError:
MsgBox "Child record Present "
RsMst_MAW.CancelUpdate
End Sub

Private Sub EditCmd_Click()
    Edit_Flg = "E"
    Call GButtonLock(Me, False)
    
'    TxtCommodityID.Locked = True
'    CmbCommodityID.Visible = True
'    TxtCommodityID.Visible = False
    CmbCommodityID.Text = TxtCommodityID.Text
    
    TxtMAWMinimum.Locked = False
    TxtMAWMaximum.Locked = False
    
    TxtStatusFlag.Locked = True
    TxtStatusFlag.Visible = False
    CmbStatusFlag.Visible = True
    CmbStatusFlag.Text = TxtStatusFlag
    
    Frame2.Enabled = True
    TxtMinimum.Locked = False
    TxtMaximum.Locked = False
    TxtMAWD.Locked = False
    
End Sub

Private Sub Form_Load()

Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Call GButtonLockAD(Me)
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If

Frame1.Enabled = True
Frame2.Enabled = True
Call Grid_Header
Call Grid_Head
Label22.Caption = "-"

Call TableMst_MAW(" Order By MAWID")
If RsMst_MAW.RecordCount = 0 Then
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

Private Sub Indata()
If RsMst_MAW.RecordCount > 0 Then
            
    TxtMAWID.Locked = True
    CmbCommodityID.Visible = False
    TxtCommodityID.Visible = True
    TxtMAWMinimum.Locked = True
    TxtMAWMaximum.Locked = True
    CmbStatusFlag.Visible = False
    TxtStatusFlag.Locked = True: TxtStatusFlag.Visible = True
            
    TxtMAWID = IIf(IsNull(RsMst_MAW!MAWID), "", RsMst_MAW!MAWID)
    mCommodityID = IIf(IsNull(RsMst_MAW!CommodityID), 0, RsMst_MAW!CommodityID)
    TxtCommodityID.Text = GetCommodityName(CDbl(mCommodityID))
    TxtMAWMinimum.Text = IIf(IsNull(RsMst_MAW!MinimumRange), "", Format(RsMst_MAW!MinimumRange, "#0.000"))
    TxtMAWMaximum.Text = IIf(IsNull(RsMst_MAW!MaximumRange), "", Format(RsMst_MAW!MaximumRange, "#0.000"))
    TxtStatusFlag.Text = IIf(IsNull(RsMst_MAW!Flag), "", IIf(LCase(RsMst_MAW!Flag) = LCase("A"), "Active", "Deactive"))
    TxtCreated.Text = IIf(IsNull(RsMst_MAW!CreatedBy), "", RsMst_MAW!CreatedBy & " :- " & Format(RsMst_MAW!CreatedDate, "DD-MMM-YYYY hh:mm:SS"))
    TxtUpdated.Text = IIf(IsNull(RsMst_MAW!UpdatedBy), "", RsMst_MAW!UpdatedBy & " :- " & Format(RsMst_MAW!UpdatedDate, "DD-MMM-YYYY hh:mm:SS"))
    
    Frame2.Enabled = False
    Call IndataMAWDetails
            
End If

If RsMst_MAW.RecordCount = 1 Then
   Call GButtonLock_1(Me, True)
Else
   Call GButtonLock(Me, True)
End If

End Sub

Private Sub IndataMAWDetails()
Dim i As Double
If RsMst_MAW.RecordCount > 0 Then
    
    Call Grid_Header
    Call TableMst_MAWDetails(" Where MAWID = " & TxtMAWID & " Order By MAWDetailID ")
    If RsMst_MAWDetails.RecordCount > 0 Then
       For i = 1 To RsMst_MAWDetails.RecordCount
            With MSHGrid_MAW
               .TextMatrix(i, 0) = i
               .TextMatrix(i, 1) = IIf(IsNull(RsMst_MAWDetails!MinimumMoisture), "", Format(RsMst_MAWDetails!MinimumMoisture, "#0.000"))
               .TextMatrix(i, 2) = IIf(IsNull(RsMst_MAWDetails!MaximumMoisture), "", Format(RsMst_MAWDetails!MaximumMoisture, "#0.000"))
               .TextMatrix(i, 3) = IIf(IsNull(RsMst_MAWDetails!MAWD), "", Format(RsMst_MAWDetails!MAWD, "#0.00"))
               .TextMatrix(i, 4) = IIf(IsNull(RsMst_MAWDetails!MAWDetailID), "0", RsMst_MAWDetails!MAWDetailID)
               .Rows = .Rows + 1
            End With
            RsMst_MAWDetails.MoveNext
       Next i
    End If
    Call ClearMAWDetails
    
End If
End Sub

Private Sub Grid_Header()
    With MSHGrid_MAW
        .Clear
        .Rows = 2
        .Cols = 5
        .AllowUserResizing = flexResizeBoth
        .WordWrap = True
    
        .TextMatrix(0, 0) = "Srno":                             .ColWidth(0) = 600
        .TextMatrix(0, 1) = "Minimum Moisture Range":           .ColWidth(1) = 2000
        .TextMatrix(0, 2) = "Maximum Moisture Range":           .ColWidth(2) = 2000
        .TextMatrix(0, 3) = "Final MAWD in %":                  .ColWidth(3) = 1800
        .TextMatrix(0, 4) = "MAWDetailsID":                     .ColWidth(4) = 0
                
        .RowHeight(0) = 600
    End With
End Sub

Private Sub Grid_Head()
    With MSHFlexGrid1
        .Clear
        .Rows = 2
        .Cols = 5
        .AllowUserResizing = flexResizeBoth
        .WordWrap = True
        
        .TextMatrix(0, 0) = "MAWID":                                        .ColWidth(0) = 800
        .TextMatrix(0, 1) = "Commodity":                                    .ColWidth(1) = 5500
        .TextMatrix(0, 2) = "Ready Reckoner Moisture Minimum Range":        .ColWidth(2) = 3000
        .TextMatrix(0, 3) = "Ready Reckoner Moisture Maximum Range":        .ColWidth(3) = 3000
        .TextMatrix(0, 4) = "Status":                                       .ColWidth(4) = 1500
        .RowHeight(0) = 600
    End With
End Sub

Private Sub MSHFlexGrid1_Click()
    Dim i As Variant
    i = MSHFlexGrid1.RowSel
    If i <= 0 Then Exit Sub
    If MSHFlexGrid1.TextMatrix(i, 0) = "" Then Exit Sub
    RsMst_MAW.MoveFirst
    RsMst_MAW.Find "MAWID = " & MSHFlexGrid1.TextMatrix(i, 0)
    If Not RsMst_MAW.EOF Then
       Call Indata
       Frame0.Visible = False
    End If
End Sub

Private Sub MSHGrid_MAW_Click()
    If SaveCmd.Enabled = True Then
       i = MSHGrid_MAW.RowSel
       If i = 0 Then
        Call ClearMAWDetails
       Else
        With MSHGrid_MAW
            If .TextMatrix(i, 0) <> "" Then
             TxtMinimum.Text = IIf(Trim(.TextMatrix(i, 1)) = "", "", Format(.TextMatrix(i, 1), "#0.000"))
             TxtMaximum.Text = IIf(Trim(.TextMatrix(i, 2)) = "", "", Format(.TextMatrix(i, 2), "#0.000"))
             TxtMAWD.Text = IIf(Trim(.TextMatrix(i, 3)) = "", "", Format(.TextMatrix(i, 3), "#0.00"))
             CmdAddtoList.Caption = "Update to List"
            Else
             Call ClearMAWDetails
            End If
        End With
       End If
    End If
End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_MAW.MoveNext
If RsMst_MAW.EOF Then
   RsMst_MAW.MoveLast
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
RsMst_MAW.MovePrevious
If RsMst_MAW.BOF Then
   RsMst_MAW.MoveFirst
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
RsMst_MAW.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsMst_MAW.MoveLast
Call Indata
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False
End Sub
Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub SaveCmd_Click()
    If CmbCommodityID.Text = "" Then
        MsgBox "Blank Commodity Not Allowed.", vbOKOnly + vbInformation, App.Title
        CmbCommodityID.SetFocus
        Exit Sub
    End If
    If TxtMAWMinimum.Text = "" Then
        MsgBox "Blank Ready Reckoner Moisture Minimum Range Not Allowed.", vbOKOnly + vbInformation, App.Title
        TxtMAWMinimum.SetFocus
        Exit Sub
    End If
    If TxtMAWMaximum.Text = "" Then
        MsgBox "Blank Ready Reckoner Moisture Maximum Range Not Allowed.", vbOKOnly + vbInformation, App.Title
        TxtMAWMaximum.SetFocus
        Exit Sub
    End If
    
    If CmbStatusFlag.Text = "" Then
        MsgBox "Blank Status Not Allowed.", vbOKOnly + vbInformation, App.Title
        CmbStatusFlag.SetFocus
        Exit Sub
    End If
    
    If Val(TxtMAWMinimum) > Val(TxtMAWMaximum) Then
        MsgBox "Minimum Value can not be more then Maximum Value.", vbOKOnly + vbInformation, App.Title
        TxtMAWMaximum.SetFocus
        Exit Sub
    End If
    
    If CheckRangeValue(TxtMAWMinimum, TxtMAWMaximum, 2) = True Then
        MsgBox "Ready Reckoner Moisture Minimum / Maximum Range is not valid with given MAWD Range.", vbOKOnly + vbInformation, App.Title
        TxtMAWMinimum.SetFocus
        Exit Sub
    End If
    
    If MSHGrid_MAW.Rows < 3 Then
        MsgBox "Ready Reckoner Moisture Range Details is not set.", vbOKOnly + vbInformation, App.Title
        TxtMinimum.SetFocus
        Exit Sub
    End If

If Edit_Flg = "A" Then
   If RsMst_MAW.RecordCount > 0 Then
      RsMst_MAW.MoveFirst
      RsMst_MAW.Find "CommodityID= '" & mCommodityID & "'"
      If Not RsMst_MAW.EOF Then
         MsgBox "Duplicate Commodity Not Allowed !!! "
         CmbCommodityID.SetFocus
         Exit Sub
      End If
   End If
      RsMst_MAW.AddNew
      RsMst_MAW!MAWID = GetMaxID("MAWID", "Mst_MAW")
      RsMst_MAW!G_UID = GetGUID
      TxtMAWID.Text = RsMst_MAW!MAWID
ElseIf Edit_Flg = "E" Then
    RsMst_MAW.MoveFirst
    RsMst_MAW.Find "CommodityID= '" & mCommodityID & "'"
    If Not RsMst_MAW.EOF Then
       If RsMst_MAW!MAWID <> TxtMAWID.Text Then
          MsgBox "Duplicate Commodity Not Allowed !!! "
          CmbCommodityID.SetFocus
          Exit Sub
       End If
    End If
    RsMst_MAW.MoveFirst
    RsMst_MAW.Find " MAWID= " & TxtMAWID.Text
End If

RsMst_MAW!CommodityID = mCommodityID
RsMst_MAW!Flag = Left(CmbStatusFlag.Text, 1)
RsMst_MAW!MinimumRange = Format(TxtMAWMinimum, "#0.000")
RsMst_MAW!MaximumRange = Format(TxtMAWMaximum, "#0.000")

If IsNull(RsMst_MAW!CreatedBy) = True Or RsMst_MAW!CreatedBy = "" Then
   RsMst_MAW!CreatedBy = Gen_UserLoginID
   RsMst_MAW!CreatedDate = Now
Else
   RsMst_MAW!UpdatedBy = Gen_UserLoginID
   RsMst_MAW!UpdatedDate = Now
End If

RsMst_MAW.Update

Call SaveMAWDetail

RsMst_MAW.Requery
RsMst_MAW.MoveFirst
RsMst_MAW.Find "MAWID = " & TxtMAWID.Text
Call Indata

End Sub

Private Sub SaveMAWDetail()
Dim mTmpMAWID As Variant
Dim i As Double
With MSHGrid_MAW
    If .Rows > 2 Then
       Call TableMst_MAWDetails(" Where MAWID = " & TxtMAWID & " Order By MAWDetailID ")
       For i = 1 To .Rows - 2
           If RsMst_MAWDetails.RecordCount > 0 Then
            RsMst_MAWDetails.MoveFirst
           End If
           
           RsMst_MAWDetails.Find ("MAWDetailID =" & Val(.TextMatrix(i, 4)))
           If RsMst_MAWDetails.EOF = True Then
                RsMst_MAWDetails.AddNew
                RsMst_MAWDetails!MAWDetailID = GetMaxID("MAWDetailID", "Mst_MAWDetails")
                RsMst_MAWDetails!G_UID = GetGUID
                .TextMatrix(i, 4) = RsMst_MAWDetails!MAWDetailID
                RsMst_MAWDetails!MAWID = TxtMAWID.Text
           End If
            
            RsMst_MAWDetails!MinimumMoisture = Format(.TextMatrix(i, 1), "#0.000")
            RsMst_MAWDetails!MaximumMoisture = Format(.TextMatrix(i, 2), "#0.000")
            RsMst_MAWDetails!MAWD = Format(.TextMatrix(i, 3), "#0.00")
            If IsNull(RsMst_MAWDetails!CreatedBy) = True Or RsMst_MAWDetails!CreatedBy = "" Then
               RsMst_MAWDetails!CreatedBy = Gen_UserLoginID
               RsMst_MAWDetails!CreatedDate = Now
            Else
               RsMst_MAWDetails!UpdatedBy = Gen_UserLoginID
               RsMst_MAWDetails!UpdatedDate = Now
            End If
            
            RsMst_MAWDetails.Update
            RsMst_MAWDetails.Requery
           
       Next i
       
       ''---- Delete MAWD Details
       For i = 1 To .Rows - 2
           mTmpMAWID = mTmpMAWID & Val(.TextMatrix(i, 4)) & ","
       Next i
       If mTmpMAWID <> "" Then
          mTmpMAWID = Left(mTmpMAWID, Len(mTmpMAWID) - 1)
       End If
       tmp = " DELETE FROM Mst_MAWDetails WHERE MAWID = " & TxtMAWID & " AND MAWDetailID NOT IN ( " & mTmpMAWID & " )"
       Call TmpTable
       ''------------------------
       
    End If
End With
End Sub

Private Sub SearchCmd_Click()

If SearchCmd.Caption = "Cancel" Then
   If RsMst_MAW.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   Call Indata
   Exit Sub
End If
Frame0.Height = Me.Height - 400
Frame0.Left = Frame1.Left
Frame0.Top = Frame1.Top
Frame0.Width = Me.Width - 300
Frame0.Visible = True
MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 1200
RsMst_MAW.MoveFirst

End Sub

Private Sub TxtMAWD_LostFocus()
    Call NumericValidation(TxtMAWD, 2)
End Sub
Private Sub TxtMAWMaximum_LostFocus()
    Call NumericValidation(TxtMAWMaximum, 3)
End Sub
Private Sub TxtMAWMinimum_LostFocus()
    Call NumericValidation(TxtMAWMinimum, 3)
End Sub
Private Sub TxtMaximum_LostFocus()
    Call NumericValidation(TxtMaximum, 3)
End Sub
Private Sub TxtMinimum_LostFocus()
    Call NumericValidation(TxtMinimum, 3)
End Sub

''-- 1 for Int, 2 for Decimal, 3 for Qty
Private Sub NumericValidation(Objname As Object, InputType As Integer)
    If Objname.Text = "" Then Exit Sub
    If IsNumeric(Objname) = False Then
       MsgBox "Invalid Format.", vbOKOnly + vbInformation, App.Title
       Objname.SetFocus
       Exit Sub
    ElseIf Val(Objname) < 0 Then
       MsgBox "Negative Value Not allow.", vbOKOnly + vbInformation, App.Title
       Objname.SetFocus
       Exit Sub
    End If
    
    If InputType = 1 Then
        Objname = Format(Objname, "#######0")
    ElseIf InputType = 2 Then
        Objname = Format(Objname, "########0.00")
    ElseIf InputType = 3 Then
        Objname = Format(Objname, "########0.000")
    End If
End Sub

Private Function CheckRangeValue(Minimum As Variant, Maximum As Variant, Index As Double) As Boolean
Dim Retval As Boolean
Retval = False
    With MSHGrid_MAW
    
    Select Case Index
    Case 1
        If .Rows > 1 Then
           For u = 1 To .Rows - 2
               If (Val(Minimum) >= Val(.TextMatrix(u, 1)) And Val(Minimum) <= Val(.TextMatrix(u, 2))) Or Val(Maximum) >= Val(.TextMatrix(u, 1)) And Val(Maximum) <= Val(.TextMatrix(u, 2)) Or (Val(Maximum) <= Val(.TextMatrix(u, 1)) And Val(Maximum) >= Val(.TextMatrix(u, 2))) Then
                  If LCase(CmdAddtoList.Caption) = "add to list" Then
                     Retval = True
                     Exit For
                  Else
                     If u <> .RowSel Then
                        Retval = True
                     End If
                  End If
               End If
           Next
        End If
    Case 2
    If .Rows > 1 Then
       For u = 1 To .Rows - 2
           ''If (Val(Minimum) < Val(.TextMatrix(u, 1)) And Val(Minimum) > Val(.TextMatrix(u, 2)))
           ''Or (Val(Maximum) < Val(.TextMatrix(u, 1)) And Val(Maximum) < Val(.TextMatrix(u, 2))) Then
           ''Val(Maximum) >= Val(.TextMatrix(u, 1)) And Val(Maximum) <= Val(.TextMatrix(u, 2)) Or
           ''If (Val(Minimum) < Val(.TextMatrix(u, 1)) And Val(Minimum) > Val(.TextMatrix(u, 2))) Or (Val(Maximum) < Val(.TextMatrix(u, 1)) And Val(Maximum) < Val(.TextMatrix(u, 2))) Then
                         
           If (Val(.TextMatrix(u, 1)) < Val(Minimum) Or Val(.TextMatrix(u, 2)) < Val(Minimum)) Then
              If LCase(CmdAddtoList.Caption) = "add to list" Then
                 Retval = True
                 Exit For
              Else
                 If u <> .RowSel Then
                    Retval = True
                 End If
              End If
           ElseIf (Val(.TextMatrix(u, 1)) > Val(Maximum) Or Val(.TextMatrix(u, 2)) > Val(Maximum)) Then
              If LCase(CmdAddtoList.Caption) = "add to list" Then
                 Retval = True
                 Exit For
              Else
                 If u <> .RowSel Then
                    Retval = True
                 End If
              End If
           End If
       Next
    End If
    End Select
    
    End With
    CheckRangeValue = Retval
End Function

''Private Function CheckDuplicateMAWD() As Boolean
''Dim Retval As Boolean
''Retval = False
''    With MSHGrid_MAW
''        If .Rows > 1 Then
''            For u = 1 To .Rows - 1
''                If TxtMAWD = .TextMatrix(u, 3) Then
''                    If LCase(CmdAddtoList.Caption) = "add to list" Then
''                        Retval = True
''                        Exit For
''                    Else
''                        If u <> .RowSel Then
''                            Retval = True
''                        End If
''                    End If
''                End If
''            Next u
''        End If
''    End With
''    CheckDuplicateMAWD = Retval
''End Function
