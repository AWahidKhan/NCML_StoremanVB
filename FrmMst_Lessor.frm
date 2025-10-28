VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmMst_Lessor 
   Caption         =   "Lessor Master"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   2295
      Left            =   120
      TabIndex        =   40
      Top             =   9000
      Visible         =   0   'False
      Width           =   14655
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
         ItemData        =   "FrmMst_Lessor.frx":0000
         Left            =   7320
         List            =   "FrmMst_Lessor.frx":000A
         Sorted          =   -1  'True
         TabIndex        =   43
         Top             =   172
         Width           =   1695
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
         Left            =   9075
         TabIndex        =   44
         Top             =   180
         Width           =   1005
      End
      Begin VB.TextBox TxtSearchLessor 
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
         Left            =   1545
         TabIndex        =   42
         Top             =   165
         Width           =   4845
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
         Left            =   10080
         TabIndex        =   41
         Top             =   180
         Width           =   1000
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   2220
         Left            =   120
         TabIndex        =   45
         Top             =   600
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   3916
         _Version        =   393216
         Rows            =   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label2 
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
         Left            =   6480
         TabIndex        =   63
         Top             =   232
         Width           =   555
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "_"
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
         Left            =   13590
         TabIndex        =   48
         Top             =   195
         Width           =   135
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
         Left            =   11190
         TabIndex        =   47
         Top             =   195
         Width           =   2370
      End
      Begin VB.Label Label19 
         Caption         =   "Lessor Name"
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
         Left            =   120
         TabIndex        =   46
         Top             =   195
         Width           =   1335
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1095
      Left            =   120
      TabIndex        =   39
      Top             =   7320
      Width           =   9510
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
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   1815
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
         Left            =   1920
         TabIndex        =   18
         Top             =   240
         Width           =   1815
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
         Left            =   3735
         TabIndex        =   19
         Top             =   240
         Width           =   1935
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
         Left            =   5670
         TabIndex        =   20
         Top             =   240
         Width           =   1815
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
         Left            =   7485
         TabIndex        =   21
         Top             =   240
         Width           =   1815
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
         Left            =   120
         TabIndex        =   22
         Top             =   615
         Width           =   1815
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
         Left            =   1920
         TabIndex        =   23
         Top             =   615
         Width           =   1815
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
         Left            =   3720
         TabIndex        =   24
         Top             =   615
         Width           =   1935
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
         Left            =   5670
         TabIndex        =   25
         Top             =   615
         Width           =   1815
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
         Left            =   7485
         TabIndex        =   26
         Top             =   615
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Height          =   7215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9495
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
         ItemData        =   "FrmMst_Lessor.frx":0021
         Left            =   6240
         List            =   "FrmMst_Lessor.frx":002B
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   2887
         Width           =   3015
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
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   2887
         Width           =   3015
      End
      Begin VB.TextBox TxtTANNO 
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
         Left            =   6240
         MaxLength       =   10
         TabIndex        =   9
         Top             =   3360
         Width           =   3015
      End
      Begin VB.ComboBox CmbTDSID 
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
         Left            =   6240
         Sorted          =   -1  'True
         TabIndex        =   16
         Top             =   5760
         Width           =   3015
      End
      Begin VB.TextBox TxtTDSCode 
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
         Left            =   6240
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   59
         Top             =   5760
         Width           =   3015
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
         Left            =   1800
         MaxLength       =   100
         TabIndex        =   15
         Top             =   5760
         Width           =   2895
      End
      Begin VB.TextBox TxtRTGSNo 
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
         Left            =   6240
         MaxLength       =   100
         TabIndex        =   14
         Top             =   5280
         Width           =   3015
      End
      Begin VB.TextBox TxtAccNo 
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
         MaxLength       =   100
         TabIndex        =   13
         Top             =   5280
         Width           =   2895
      End
      Begin VB.ComboBox CmbBranch 
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
         Left            =   1800
         Sorted          =   -1  'True
         TabIndex        =   12
         Top             =   4800
         Width           =   7005
      End
      Begin VB.TextBox TxtBranchName 
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
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   4800
         Width           =   6990
      End
      Begin VB.CommandButton CmbBranchName 
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
         Left            =   8865
         TabIndex        =   53
         Top             =   4800
         Width           =   375
      End
      Begin VB.ComboBox CmbBankName 
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
         Left            =   1800
         Sorted          =   -1  'True
         TabIndex        =   11
         Top             =   4320
         Width           =   7005
      End
      Begin VB.TextBox TxtBankName 
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
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   4320
         Width           =   6990
      End
      Begin VB.CommandButton CmdBank 
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
         Left            =   8865
         TabIndex        =   50
         Top             =   4320
         Width           =   375
      End
      Begin VB.TextBox TxtPAN 
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
         MaxLength       =   10
         TabIndex        =   8
         Top             =   3360
         Width           =   2895
      End
      Begin VB.TextBox TxtLessorID 
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox TxtLessorName 
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
         MaxLength       =   200
         TabIndex        =   1
         Top             =   720
         Width           =   7455
      End
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
         Height          =   1095
         Left            =   1815
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   1185
         Width           =   7455
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
         Left            =   6240
         MaxLength       =   15
         TabIndex        =   4
         Top             =   2400
         Width           =   3015
      End
      Begin VB.TextBox TxtEmailId 
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
         MaxLength       =   100
         TabIndex        =   10
         Top             =   3840
         Width           =   7455
      End
      Begin VB.TextBox TxtFax 
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
         MaxLength       =   15
         TabIndex        =   5
         Top             =   2880
         Width           =   2895
      End
      Begin VB.TextBox TxtContactNo 
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
         MaxLength       =   15
         TabIndex        =   3
         Top             =   2400
         Width           =   2895
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   6240
         Width           =   7455
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   6720
         Width           =   7455
      End
      Begin VB.Label Label38 
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
         Left            =   4920
         TabIndex        =   62
         Top             =   2940
         Width           =   660
      End
      Begin VB.Label Label1 
         Caption         =   "TAN No"
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
         Left            =   4920
         TabIndex        =   61
         Top             =   3420
         Width           =   1695
      End
      Begin VB.Label LblTDSCode 
         Caption         =   "TDS Code"
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
         Left            =   4920
         TabIndex        =   60
         Top             =   5820
         Width           =   1695
      End
      Begin VB.Label LblNAVCode 
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
         Height          =   255
         Left            =   240
         TabIndex        =   58
         Top             =   5820
         Width           =   1695
      End
      Begin VB.Label LblRTGSNo 
         Caption         =   "RTGS No"
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
         Left            =   4950
         TabIndex        =   57
         Top             =   5340
         Width           =   1695
      End
      Begin VB.Label LblAccNo 
         Caption         =   "Account No"
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
         Left            =   240
         TabIndex        =   56
         Top             =   5340
         Width           =   1695
      End
      Begin VB.Label LblBranchName 
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
         Height          =   255
         Left            =   240
         TabIndex        =   55
         Top             =   4860
         Width           =   1695
      End
      Begin VB.Label Label3 
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
         Height          =   255
         Left            =   240
         TabIndex        =   52
         Top             =   4380
         Width           =   1695
      End
      Begin VB.Label LblPAN 
         Caption         =   "PAN No"
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
         TabIndex        =   49
         Top             =   3420
         Width           =   1695
      End
      Begin VB.Label LblAgentName 
         Caption         =   "Lessor Name"
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
         TabIndex        =   38
         Top             =   780
         Width           =   1695
      End
      Begin VB.Label LblLessorID 
         Caption         =   "Lessor ID"
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
         Left            =   240
         TabIndex        =   37
         Top             =   300
         Width           =   1695
      End
      Begin VB.Label LblAdess 
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
         Height          =   255
         Left            =   255
         TabIndex        =   36
         Top             =   1665
         Width           =   1695
      End
      Begin VB.Label LblMobileNo 
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
         Height          =   255
         Left            =   4920
         TabIndex        =   35
         Top             =   2460
         Width           =   1695
      End
      Begin VB.Label LblEmailId 
         Caption         =   "EmailId"
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
         Left            =   240
         TabIndex        =   34
         Top             =   3900
         Width           =   1695
      End
      Begin VB.Label LblFax 
         Caption         =   "Fax"
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
         Left            =   240
         TabIndex        =   33
         Top             =   2940
         Width           =   1695
      End
      Begin VB.Label LblContactNo 
         Caption         =   "Contact No"
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
         Left            =   240
         TabIndex        =   32
         Top             =   2460
         Width           =   1695
      End
      Begin VB.Label Label10 
         Caption         =   "Created By"
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
         Left            =   240
         TabIndex        =   31
         Top             =   6300
         Width           =   1695
      End
      Begin VB.Label Label9 
         Caption         =   "Updated By"
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
         Left            =   240
         TabIndex        =   30
         Top             =   6780
         Width           =   1695
      End
   End
End
Attribute VB_Name = "FrmMst_Lessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mBankID, mBranchID, mTDSID, Edit_Flg As Variant

Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)
CmbBankName.Visible = True
TxtBankName.Visible = False
CmbBranch.Visible = True
TxtBranchName.Visible = False

TxtCreated = ""
TxtUpdated = ""

TxtLessorID = ""
TxtLessorName = ""
TxtAddress = ""
TxtContactNo = ""
TxtMobileNo = ""
TxtEmailId = ""
TxtFax = ""
TxtPAN = ""
TxtAccNo = ""
TxtRTGSNo = ""

mBankID = ""
mBranchID = ""
mTDSID = ""
If LCase(Gen_UserRole) = "fag" Or LCase(Gen_UserRole) = "fagadmin" Then 'LCase(Gen_UserRole) = "power" Or
   TxtNAVCode.Enabled = True
   CmbTDSID.Visible = True
   CmbTDSID.Enabled = True
   CmbTDSID.Text = ""
   TxtTDSCode.Text = ""
   TxtTDSCode.Visible = False
Else
   TxtNAVCode.Enabled = False
   CmbTDSID.Visible = True
   CmbTDSID.Enabled = False
   CmbTDSID.Text = ""
   TxtTDSCode.Visible = False
End If

''-- 09 Aug 2017 Ullhas Patil
TxtTANNO = ""
TxtTANNO.Locked = False
''------------------
''-- Ullhas Patil 20 Oct 2017
TxtStatus.Text = ""
CmbStatus.Text = ""
CmbStatus.Visible = True
''------------------

Frame1.Enabled = True
If RsMst_Lessor.RecordCount > 0 Then
   TxtLessorName.SetFocus
Else
   SearchCmd.Enabled = False
End If
End Sub

Private Sub CmbBankName_GotFocus()
tmp = CmbBankName.Text
Call TableMst_Bank("")
CmbBranch.Clear
CmbBankName.Clear
If RsMst_Bank.RecordCount = 0 Then
   MsgBox "No Bank found"
   CmbBankName.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Bank.RecordCount
    CmbBankName.AddItem RsMst_Bank!BankName
    RsMst_Bank.MoveNext
Next
CmbBankName.Text = tmp
End Sub

Private Sub CmbBankName_LostFocus()
If CmbBankName.Text = "" Then
   mBankID = ""
   Exit Sub
End If
RsMst_Bank.MoveFirst
RsMst_Bank.Find "BankName = '" & CmbBankName.Text & "'"

If RsMst_Bank.EOF Then
   MsgBox "Invalid Bank selection "
   CmbBankName.SetFocus
   Exit Sub
End If
mBankID = RsMst_Bank!BankID

End Sub
Private Sub CmbBranch_GotFocus()
tmp = CmbBranch.Text
If CmbBankName = "" Then
'   MsgBox "Select Bank Name First."
'   CmbBankName.SetFocus
   Exit Sub
End If
Call TableMst_Branch(" Where BankID=" & mBankID & "")

CmbBranch.Clear
If RsMst_Branch.RecordCount = 0 Then
   MsgBox "No Branch found"
   CmbBranch.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Branch.RecordCount
    CmbBranch.AddItem RsMst_Branch!BranchName
    RsMst_Branch.MoveNext
Next
CmbBranch.Text = tmp
End Sub

Private Sub CmbBranch_LostFocus()
If CmbBranch.Text = "" Then
   mBranchID = ""
   Exit Sub
End If
RsMst_Branch.MoveFirst
RsMst_Branch.Find "BranchName = '" & CmbBranch.Text & "'"

If RsMst_Branch.EOF Then
   MsgBox "Invalid Branch selection "
   CmbBranch.SetFocus
   Exit Sub
End If
mBranchID = RsMst_Branch!BranchID
End Sub

Private Sub CmbBranchName_Click()
FrmMst_Branch.Show
End Sub

Private Sub CmbSStatus_LostFocus()
If CmbSStatus.Text = "" Then Exit Sub
If LCase(CmbSStatus.Text) <> "active" And LCase(CmbSStatus.Text) <> "de-active" Then
   MsgBox "Invalid selection!!! "
   CmbSStatus.SetFocus
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

Private Sub CmbTDSID_GotFocus()
tmp = CmbTDSID.Text
Call TableMst_TDS
CmbTDSID.Clear
If RsMst_TDS.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_TDS.RecordCount
    CmbTDSID.AddItem RsMst_TDS!Description
    RsMst_TDS.MoveNext
Next
CmbTDSID.Text = tmp
End Sub

Private Sub CmbTDSID_LostFocus()
If CmbTDSID.Text = "" Then
   mTDSID = ""
   Exit Sub
End If
RsMst_TDS.MoveFirst
RsMst_TDS.Find "Description = '" & CmbTDSID.Text & "'"
If RsMst_TDS.EOF Then
   MsgBox "Invalid selection "
   CmbTDSID.SetFocus
   Exit Sub
End If
mTDSID = RsMst_TDS!TDSID

End Sub

Private Sub CmdBank_Click()
FrmMst_Bank.Show
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "RsMst_Lessor.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "RsMst_Lessor.xls")
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
           If C = 11 Then
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

Private Sub CmdGo_Click()
If LCase(Gen_UserRole) = "power" Or LCase(Gen_UserRole) = "fag" Or LCase(Gen_UserRole) = "fagadmin" Then

   tmp = "Select ML.LessorID,ML.LessorName,ML.Address,ML.ContactNo,ML.MobileNo,ML.EmailID,ML.Fax,ML.PANNo,ML.TANNO,MB.BankName,MBr.BranchName,ML.AccountNo,ML.RTGSNo,ML.Status,ML.NAVCode " & _
         " From Mst_Lessor ML  left Join Mst_Bank MB on ML.BankID=MB.BankID left Join Mst_Branch MBr on ML.BranchID=MBr.BranchID"
Else

   tmp = "Select ML.LessorID,ML.LessorName,ML.Address,ML.ContactNo,ML.MobileNo,ML.EmailID,ML.Fax,ML.PANNo,ML.TANNO,MB.BankName,MBr.BranchName,ML.AccountNo,ML.Status,ML.RTGSNo " & _
         " From Mst_Lessor ML  left Join Mst_Bank MB on ML.BankID=MB.BankID left Join Mst_Branch MBr on ML.BranchID=MBr.BranchID"

End If
Dim wc As Variant
wc = ""
If TxtSearchLessor <> "" Then
   wc = " Where ML.LessorName Like '%" & Trim(TxtSearchLessor) & "%' "
End If

''-- Ullhas Patil 20 Oct 2017
If CmbSStatus.Text <> "" Then
    If wc = "" Then
        wc = " Where ML.Status = '" & Left(CmbSStatus.Text, 1) & "' "
    Else
        wc = wc & " Where ML.Status = '" & Left(CmbSStatus.Text, 1) & "' "
    End If
End If
''--------------------

tmp = tmp & wc & " Order by ML.LessorName"

Call TmpTable
Call GridHead
For iu = 1 To TmpRs.RecordCount
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
    For C = 0 To MSHFlexGrid1.Cols - 1
        If C = 10 Then
           MSHFlexGrid1.TextMatrix(iu, C) = IIf(IsNull(TmpRs.Fields(C)), "", "'" & TmpRs.Fields(C))
        Else
           MSHFlexGrid1.TextMatrix(iu, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
        End If
    Next
    
    If LCase(MSHFlexGrid1.TextMatrix(iu, 13)) = "a" Then
        MSHFlexGrid1.TextMatrix(iu, 13) = "Active"
    ElseIf LCase(MSHFlexGrid1.TextMatrix(iu, 13)) = "d" Then
        MSHFlexGrid1.TextMatrix(iu, 13) = "De-Active"
    Else
        MSHFlexGrid1.TextMatrix(iu, 13) = ""
    End If
    
    TmpRs.MoveNext
Next
Label22.Refresh
Label22.Caption = TmpRs.RecordCount

End Sub

Private Sub DeleteCmd_Click()
Dim ans As Variant
On Error GoTo msgError
ans = MsgBox("Do you really want to DELETE this record??? ", vbYesNo)
If ans = vbYes Then
   
   RsMst_Lessor.Delete
   RsMst_Lessor.Update
   'MsgBox "Record Successfully Removed"
   If RsMst_Lessor.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_Lessor.MoveNext
   If RsMst_Lessor.EOF() Then
      RsMst_Lessor.MoveLast
   End If
   Call Indata
End If

Exit Sub
msgError:
MsgBox "Child record Present "
RsMst_Lessor.CancelUpdate

End Sub

Private Sub EditCmd_Click()
Edit_Flg = "E"
Frame1.Enabled = True
Call GButtonLock(Me, False)
TxtLessorID.Locked = False
CmbBankName.Visible = True
TxtBankName.Visible = False
CmbBranch.Visible = True
TxtBranchName.Visible = False
CmbBankName.Text = TxtBankName.Text
CmbBranch.Text = TxtBranchName.Text

If LCase(Gen_UserRole) = "power" Or LCase(Gen_UserRole) = "fag" Or LCase(Gen_UserRole) = "fagadmin" Then
   TxtNAVCode.Enabled = True
   CmbTDSID.Visible = True
   CmbTDSID.Enabled = True
   CmbTDSID.Text = TxtTDSCode.Text
   TxtTDSCode.Visible = False
Else
   TxtNAVCode.Enabled = False
   CmbTDSID.Visible = False
   TxtTDSCode.Visible = True
End If


TxtTANNO.Locked = False             ''-- 09 Aug 2017 Ullhas Patil
CmbStatus.Text = TxtStatus.Text     ''-- Ullhas Patil 20 Oct 2017
CmbStatus.Visible = True            ''-- Ullhas Patil 20 Oct 2017

End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub FirstCmd_Click()
RsMst_Lessor.MoveFirst
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
   MsgBox "Access denied !!!!!!!!!"
   Frame1.Enabled = False
   Exit Sub
End If
Frame1.Enabled = True

Call GridHead
TxtSearchLessor = ""
Label22.Caption = ""

Call TableMst_Lessor("  Order by LessorID ")
If RsMst_Lessor.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
Call Indata
Call GButtonLock(Me, True)
End Sub
Private Function GetLessorID() As Double
Dim Retval As Double
tmp = "Select max(LessorID) from Mst_Lessor"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetLessorID = Retval
End Function
Private Sub GridHead()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
If LCase(Gen_UserRole) = "power" Or LCase(Gen_UserRole) = "fag" Or LCase(Gen_UserRole) = "fagadmin" Then
   MSHFlexGrid1.Cols = 15
Else
   MSHFlexGrid1.Cols = 14
End If
MSHFlexGrid1.TextMatrix(0, 0) = "Lessor Id":            MSHFlexGrid1.ColWidth(0) = 1000
MSHFlexGrid1.TextMatrix(0, 1) = "Lessor Name":          MSHFlexGrid1.ColWidth(1) = 2500
MSHFlexGrid1.TextMatrix(0, 2) = "Address":              MSHFlexGrid1.ColWidth(2) = 2500
MSHFlexGrid1.TextMatrix(0, 3) = "Contact No":           MSHFlexGrid1.ColWidth(3) = 1600
MSHFlexGrid1.TextMatrix(0, 4) = "Mobile No":            MSHFlexGrid1.ColWidth(4) = 1400
MSHFlexGrid1.TextMatrix(0, 5) = "Email ID":             MSHFlexGrid1.ColWidth(5) = 1800
MSHFlexGrid1.TextMatrix(0, 6) = "Fax":                  MSHFlexGrid1.ColWidth(6) = 1400
MSHFlexGrid1.TextMatrix(0, 7) = "PAN No":               MSHFlexGrid1.ColWidth(7) = 1500
MSHFlexGrid1.TextMatrix(0, 8) = "TAN No":               MSHFlexGrid1.ColWidth(8) = 1500
MSHFlexGrid1.TextMatrix(0, 9) = "Bank Name":            MSHFlexGrid1.ColWidth(9) = 2000
MSHFlexGrid1.TextMatrix(0, 10) = "Branch Name":         MSHFlexGrid1.ColWidth(10) = 1400
MSHFlexGrid1.TextMatrix(0, 11) = "Account No":          MSHFlexGrid1.ColWidth(11) = 1400
MSHFlexGrid1.TextMatrix(0, 12) = "RTGS No":             MSHFlexGrid1.ColWidth(12) = 1400
MSHFlexGrid1.TextMatrix(0, 13) = "Status":              MSHFlexGrid1.ColWidth(13) = 1000

If LCase(Gen_UserRole) = "power" Or LCase(Gen_UserRole) = "fag" Or LCase(Gen_UserRole) = "fagadmin" Then
   MSHFlexGrid1.TextMatrix(0, 14) = "NAV Code"
End If

End Sub
Public Sub Indata()
CmbBankName.Visible = False
TxtBankName.Visible = True
CmbBranch.Visible = False
CmbTDSID.Visible = False
TxtTDSCode.Visible = True
TxtBranchName.Visible = True

TxtLessorID = RsMst_Lessor!LessorID
TxtLessorName = RsMst_Lessor!LessorName

TxtAddress = IIf(IsNull(RsMst_Lessor!Address), "", RsMst_Lessor!Address)
TxtContactNo = IIf(IsNull(RsMst_Lessor!ContactNo), "", RsMst_Lessor!ContactNo)
TxtMobileNo = IIf(IsNull(RsMst_Lessor!MobileNo), "", RsMst_Lessor!MobileNo)
TxtEmailId = IIf(IsNull(RsMst_Lessor!EmailID), "", RsMst_Lessor!EmailID)
TxtBankName = GetBankName(IIf(IsNull(RsMst_Lessor!BankID), "0", RsMst_Lessor!BankID), "N")
mBankID = RsMst_Lessor!BankID
mBranchID = RsMst_Lessor!BranchID
TxtBranchName = GetBranchName(IIf(IsNull(RsMst_Lessor!BranchID), "0", RsMst_Lessor!BranchID), "N")
TxtAccNo = IIf(IsNull(RsMst_Lessor!AccountNo), "", RsMst_Lessor!AccountNo)
TxtRTGSNo = IIf(IsNull(RsMst_Lessor!RTGSNo), "", RsMst_Lessor!RTGSNo)
TxtPAN = IIf(IsNull(RsMst_Lessor!PANNO), "", RsMst_Lessor!PANNO)
TxtNAVCode = IIf(IsNull(RsMst_Lessor!NavCode), "", RsMst_Lessor!NavCode)
TxtTDSCode = GetTDSDescription(IIf(IsNull(RsMst_Lessor!TDSCode), 0, RsMst_Lessor!TDSCode))
mTDSID = IIf(IsNull(RsMst_Lessor!TDSCode), 0, RsMst_Lessor!TDSCode)
TxtFax = IIf(IsNull(RsMst_Lessor!Fax), "", RsMst_Lessor!Fax)
''-- 09 Aug 2017 Ullhas Patil
TxtTANNO.Locked = True
TxtTANNO = IIf(IsNull(RsMst_Lessor!TANNO), "", RsMst_Lessor!TANNO)
''-- Ullhas Patil 20 Oct 2017
CmbStatus.Visible = False
TxtStatus.Text = IIf(IsNull(RsMst_Lessor!Status), "Active", IIf(LCase(RsMst_Lessor!Status) = "a", "Active", "De-Active"))
''-----------------------
TxtCreated = IIf(IsNull(RsMst_Lessor!CreatedBy), "", RsMst_Lessor!CreatedBy) & " :- " & IIf(IsNull(RsMst_Lessor!CreatedDate), "", RsMst_Lessor!CreatedDate)
TxtUpdated = IIf(IsNull(RsMst_Lessor!UpdatedBy), "", RsMst_Lessor!UpdatedBy) & " :- " & IIf(IsNull(RsMst_Lessor!UpdatedDate), "", RsMst_Lessor!UpdatedDate)

Frame1.Enabled = False
Call GButtonLock(Me, True)
End Sub
Private Sub LastCmd_Click()
RsMst_Lessor.MoveLast
Call Indata
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False
End Sub

Private Sub MSHFlexGrid1_Click()
Dim x As Variant
Dim mCaId As Variant

x = MSHFlexGrid1.RowSel

mCaId = MSHFlexGrid1.TextMatrix(x, 0)

RsMst_Lessor.MoveFirst

If RsMst_Lessor.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
RsMst_Lessor.MoveFirst
RsMst_Lessor.Find "LessorID = " & Val(mCaId) & ""
If RsMst_Lessor.EOF Then
   RsMst_Lessor.MoveLast
End If
Frame0.Visible = False
Call Indata
End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_Lessor.MoveNext
If RsMst_Lessor.EOF Then
   RsMst_Lessor.MoveLast
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
RsMst_Lessor.MovePrevious
If RsMst_Lessor.BOF Then
   RsMst_Lessor.MoveFirst
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
If TxtLessorName = "" Then
   MsgBox "Blank Lessor Name not Allowed"
   TxtLessorName.SetFocus
   Exit Sub
End If
''---------- 20 Oct 2017 Ullhas Patil
If CmbStatus.Text = "" Then
      MsgBox "Blank Status Not Allow."
      CmbStatus.SetFocus
      Exit Sub
End If
''----------
If TxtPAN = "" Then
   MsgBox "Blank PAN No not Allowed"
   TxtPAN.SetFocus
   Exit Sub
End If
''--- 09 Aug 2017 Ullhas Patil
''If TxtTANNO.Text = "" Then
''   MsgBox "Blank TAN No Allowed."
''   TxtTANNO.SetFocus
''   Exit Sub
''Else
If Trim(TxtTANNO.Text) <> "" Then
   If ValidateTANNo(UCase(TxtTANNO.Text)) = False Then
      MsgBox "Invalid TAN No."
      TxtTANNO.SetFocus
      Exit Sub
   End If
End If
''-----------------------------

If Edit_Flg = "A" Then
    If RsMst_Lessor.RecordCount > 0 Then
    RsMst_Lessor.MoveFirst
    RsMst_Lessor.Find "LessorName = '" & TxtLessorName & "'"
    If Not RsMst_Lessor.EOF Then
       MsgBox "Duplicate Lessor Name not Allowed"
       Exit Sub
    End If
    RsMst_Lessor.MoveLast
    ''-- Alter for duplicate TAN No 09 Aug 2017 Ullhas Patil
    RsMst_Lessor.Filter = ""
    RsMst_Lessor.MoveFirst
    If Trim(TxtTANNO.Text) <> "" Then
    RsMst_Lessor.Find ("TANNO ='" & Trim(TxtTANNO.Text) & "'")
    If Not RsMst_Lessor.EOF Then
      Select Case MsgBox("Duplicate TAN No Found with another Lessor, Do you want to save the Lessor with same TAN No?", vbYesNo)
          Case vbNo
              TxtTANNO.SetFocus
              Exit Sub
      End Select
    End If
    End If
    ''-----------------------------
    End If
    RsMst_Lessor.AddNew
    RsMst_Lessor!LessorID = GetLessorID
    RsMst_Lessor!G_UID = GetGUID
    TxtLessorID = RsMst_Lessor!LessorID
Else
    RsMst_Lessor.MoveFirst
    RsMst_Lessor.Find "LessorName = '" & TxtLessorName & "'"
    If Not RsMst_Lessor.EOF Then
       If RsMst_Lessor!LessorID <> TxtLessorID Then
          MsgBox "Duplicate Lessor Name not Allowed"
          Exit Sub
       End If
    End If
    ''-- Alter for duplicate TAN No 09 Aug 2017 Ullhas Patil
    RsMst_Lessor.Filter = ""
    RsMst_Lessor.MoveFirst
    If Trim(TxtTANNO.Text) <> "" Then
    RsMst_Lessor.Find ("TANNO ='" & Trim(TxtTANNO.Text) & "'")
    If Not RsMst_Lessor.EOF Then
       If RsMst_Lessor!LessorID <> TxtLessorID.Text Then
          Select Case MsgBox("Duplicate TAN No Found with another Lessor, Do you want to save the Lessor with same TAN No?", vbYesNo)
             Case vbNo
                 TxtTANNO.SetFocus
                 Exit Sub
          End Select
       End If
    End If
    End If
    ''-----------------------------
    RsMst_Lessor.MoveFirst
    RsMst_Lessor.Find "LessorID = " & TxtLessorID & ""
End If
RsMst_Lessor!LessorName = TxtLessorName
RsMst_Lessor!Address = TxtAddress
RsMst_Lessor!ContactNo = TxtContactNo
RsMst_Lessor!MobileNo = TxtMobileNo
RsMst_Lessor!EmailID = TxtEmailId
RsMst_Lessor!Fax = TxtFax
RsMst_Lessor!PANNO = TxtPAN
RsMst_Lessor!AccountNo = TxtAccNo
RsMst_Lessor!RTGSNo = TxtRTGSNo

If LCase(Gen_UserRole) = "power" Or LCase(Gen_UserRole) = "fag" Or LCase(Gen_UserRole) = "fagadmin" Then '
   RsMst_Lessor!TDSCode = Val(mTDSID)
   RsMst_Lessor!NavCode = Trim(TxtNAVCode.Text)
End If
If CmbBankName.Text <> "" Then
   
   If CmbBranch.Text <> "" Then
      RsMst_Lessor!BranchID = mBranchID
      RsMst_Lessor!BankID = mBankID
   End If
   
Else
   RsMst_Lessor!BankID = Null
   RsMst_Lessor!BranchID = Null
End If

RsMst_Lessor!TANNO = Trim(TxtTANNO.Text)        ''-- 09 Aug 2017 Ullhas Patil
RsMst_Lessor!Status = Left(CmbStatus.Text, 1)   ''-- Ullhas Patil 20 Oct 2017

If IsNull(RsMst_Lessor!CreatedBy) = True Or RsMst_Lessor!CreatedBy = "" Then
   RsMst_Lessor!CreatedBy = Gen_UserLoginID
   RsMst_Lessor!CreatedDate = Now
Else
   RsMst_Lessor!UpdatedBy = Gen_UserLoginID
   RsMst_Lessor!UpdatedDate = Now
End If

RsMst_Lessor.Update
RsMst_Lessor.Requery
RsMst_Lessor.MoveFirst
RsMst_Lessor.Find "LessorID = " & TxtLessorID & ""
If RsMst_Lessor.EOF Then
   RsMst_Lessor.MoveFirst
End If
Call GButtonLock(Me, True)
Call Indata

End Sub

Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   Call GButtonLock(Me, True)
   If RsMst_Lessor.EOF Then
      RsMst_Lessor.MoveFirst
   End If
   Call Indata

   Exit Sub
End If

Frame0.Left = Frame1.Left
Frame0.Top = Frame1.Top
Frame0.Height = Me.Height - 800
Frame0.Width = Me.Width - 200
Frame0.Visible = True

MSHFlexGrid1.Width = Frame0.Width - 400
MSHFlexGrid1.Height = Frame0.Height - 700

RsMst_Lessor.MoveFirst

''===========comment on 7 july 2012 for not clearing search grid=============
'MSHFlexGrid1.Clear
'Call GridHead
'TxtSearchLessor = ""
'Label22.Caption = ""
''===========comment on 7 july 2012 for not clearing search grid=============

End Sub

Private Sub TxtLessorName_LostFocus()
Dim ans As Boolean
ans = IsClean(TxtLessorName)
If ans = False Then
   MsgBox "Special character not Allowed  !!!!!!!! "
   TxtLessorName.SetFocus
End If
End Sub

Private Sub TxtTANNo_LostFocus()
If TxtTANNO.Locked = False Then
    If Trim(TxtTANNO.Text) <> "" Then
       TxtTANNO = UCase(TxtTANNO)
       If ValidateTANNo(TxtTANNO) = False Then
        MsgBox "Invalid TAN No"
        TxtTANNO.SetFocus
        Exit Sub
       End If
    End If
End If
End Sub
