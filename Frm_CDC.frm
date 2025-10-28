VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Frm_CDC 
   Caption         =   "CDC"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   11490
   ScaleWidth      =   19080
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbCommodity 
      Height          =   315
      Left            =   2040
      TabIndex        =   25
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1200
      Left            =   120
      TabIndex        =   14
      Top             =   2280
      Width           =   10000
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   400
         Left            =   4920
         TabIndex        =   24
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "Search"
         Height          =   400
         Left            =   4920
         TabIndex        =   23
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   400
         Left            =   3720
         TabIndex        =   22
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   400
         Left            =   3720
         TabIndex        =   21
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   400
         Left            =   2520
         TabIndex        =   20
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   400
         Left            =   2520
         TabIndex        =   19
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   400
         Left            =   1320
         TabIndex        =   18
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   400
         Left            =   1320
         TabIndex        =   17
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   400
         Left            =   100
         TabIndex        =   16
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New"
         Height          =   400
         Left            =   100
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2265
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   10000
      Begin VB.TextBox TxtClientID 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   360
         Left            =   1920
         TabIndex        =   8
         Top             =   240
         Width           =   1200
      End
      Begin VB.TextBox TxtClientName 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1920
         MaxLength       =   50
         TabIndex        =   7
         Top             =   615
         Width           =   3600
      End
      Begin VB.TextBox TxtDPName 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   325
         Left            =   1920
         TabIndex        =   6
         Top             =   1395
         Width           =   3600
      End
      Begin VB.TextBox TxtDPID 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   325
         Left            =   1920
         TabIndex        =   5
         Top             =   1035
         Width           =   1560
      End
      Begin VB.Label LblClientID 
         Caption         =   "Client ID"
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
         TabIndex        =   13
         Top             =   240
         Width           =   1500
      End
      Begin VB.Label LblClientName 
         Caption         =   "Client Name"
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
         TabIndex        =   12
         Top             =   615
         Width           =   1500
      End
      Begin VB.Label LblCommodity 
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
         Height          =   300
         Left            =   240
         TabIndex        =   11
         Top             =   1755
         Width           =   1500
      End
      Begin VB.Label LblDPName 
         Caption         =   "DP Name"
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
         TabIndex        =   10
         Top             =   1395
         Width           =   1500
      End
      Begin VB.Label LblDPID 
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
         Height          =   300
         Left            =   240
         TabIndex        =   9
         Top             =   1035
         Width           =   1500
      End
   End
   Begin VB.Frame Frame0 
      Height          =   2085
      Left            =   120
      TabIndex        =   0
      Top             =   3480
      Width           =   10000
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   350
         Left            =   80
         TabIndex        =   1
         Top             =   150
         Width           =   1000
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1455
         Left            =   75
         TabIndex        =   2
         Top             =   600
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   2566
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   11115
      Width           =   19080
      _ExtentX        =   33655
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Frm_CDC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
