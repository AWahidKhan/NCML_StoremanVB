VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Frm_NewISINRevalidation 
   Caption         =   "New ISIN Revalidation"
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
      Height          =   2655
      Left            =   480
      TabIndex        =   40
      Top             =   4800
      Visible         =   0   'False
      Width           =   14295
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
         Height          =   1215
         Left            =   120
         TabIndex        =   53
         Top             =   1560
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   2143
         _Version        =   393216
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
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
         Height          =   465
         Left            =   12360
         TabIndex        =   47
         Top             =   945
         Width           =   1125
      End
      Begin VB.TextBox TxtSOldISIN 
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
         Left            =   6397
         TabIndex        =   46
         Top             =   480
         Width           =   1905
      End
      Begin VB.ComboBox CmbPreFix 
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
         TabIndex        =   45
         Top             =   480
         Width           =   2685
      End
      Begin VB.ComboBox CmbSCmp 
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
         Left            =   2921
         Sorted          =   -1  'True
         TabIndex        =   44
         Top             =   480
         Width           =   3360
      End
      Begin VB.TextBox TxtSClientName 
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
         Left            =   10320
         TabIndex        =   43
         Top             =   480
         Width           =   4305
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
         Height          =   465
         Left            =   13500
         TabIndex        =   42
         Top             =   945
         Width           =   1125
      End
      Begin VB.TextBox TxtSARNo 
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
         Left            =   8418
         TabIndex        =   41
         Top             =   480
         Width           =   1785
      End
      Begin VB.Label Label52 
         Caption         =   "Old ISIN"
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
         Left            =   6840
         TabIndex        =   52
         Top             =   150
         Width           =   855
      End
      Begin VB.Label Label50 
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
         Height          =   300
         Left            =   960
         TabIndex        =   51
         Top             =   150
         Width           =   1320
      End
      Begin VB.Label Label49 
         Caption         =   "CMP"
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
         Left            =   4395
         TabIndex        =   50
         Top             =   150
         Width           =   1320
      End
      Begin VB.Label Label48 
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
         Height          =   255
         Left            =   11760
         TabIndex        =   49
         Top             =   150
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "AR No"
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
         Left            =   9000
         TabIndex        =   48
         Top             =   150
         Width           =   855
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1200
      Left            =   120
      TabIndex        =   29
      Top             =   8400
      Width           =   14925
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New"
         Height          =   400
         Left            =   105
         TabIndex        =   39
         Top             =   195
         Width           =   3060
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   400
         Left            =   105
         TabIndex        =   38
         Top             =   600
         Width           =   3060
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   400
         Left            =   3165
         TabIndex        =   37
         Top             =   195
         Width           =   2940
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   400
         Left            =   3165
         TabIndex        =   36
         Top             =   600
         Width           =   2940
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   400
         Left            =   6105
         TabIndex        =   35
         Top             =   195
         Width           =   2940
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   400
         Left            =   6105
         TabIndex        =   34
         Top             =   600
         Width           =   2940
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   400
         Left            =   9045
         TabIndex        =   33
         Top             =   195
         Width           =   2940
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   400
         Left            =   9045
         TabIndex        =   32
         Top             =   600
         Width           =   2940
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "Search"
         Height          =   400
         Left            =   11985
         TabIndex        =   31
         Top             =   195
         Width           =   2820
      End
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   400
         Left            =   11985
         TabIndex        =   30
         Top             =   600
         Width           =   2820
      End
   End
   Begin VB.Frame Frame4 
      Height          =   3540
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   15015
      Begin VB.ComboBox CmbClientDP 
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
         TabIndex        =   5
         Top             =   3000
         Visible         =   0   'False
         Width           =   12180
      End
      Begin VB.TextBox TxtDPName 
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
         Left            =   7680
         Locked          =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   3000
         Width           =   6525
      End
      Begin VB.TextBox TxtDPID 
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
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   3000
         Width           =   2520
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
         Left            =   7680
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   2520
         Width           =   6525
      End
      Begin VB.TextBox TxtClientID 
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
         TabIndex        =   4
         Top             =   2520
         Width           =   2520
      End
      Begin VB.TextBox TxtOldCommodity 
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
         Left            =   7680
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   2040
         Width           =   6525
      End
      Begin VB.TextBox TxtOldLocationID 
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
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   2055
         Width           =   2520
      End
      Begin VB.TextBox txtOldCMP 
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
         Left            =   7680
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1545
         Width           =   6525
      End
      Begin VB.TextBox TxtOldISINID 
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
         TabIndex        =   3
         Top             =   1560
         Width           =   2520
      End
      Begin VB.TextBox TxtGrade 
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
         Left            =   7680
         TabIndex        =   6
         Top             =   720
         Width           =   6525
      End
      Begin VB.TextBox TxtRevalidationID 
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
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   240
         Width           =   2520
      End
      Begin VB.TextBox TxtARNo 
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
         MaxLength       =   50
         TabIndex        =   1
         Top             =   720
         Width           =   2535
      End
      Begin MSComCtl2.DTPicker DtpRevalidationDate 
         Height          =   360
         Left            =   7680
         TabIndex        =   8
         Top             =   225
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   61669377
         CurrentDate     =   40271
      End
      Begin VB.TextBox TxtPreFix 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   325
         Left            =   7965
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   1560
         Visible         =   0   'False
         Width           =   2000
      End
      Begin MSComCtl2.DTPicker DtpARDate 
         Height          =   360
         Left            =   2040
         TabIndex        =   2
         Top             =   1155
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   635
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
         Format          =   61669377
         CurrentDate     =   40271
      End
      Begin VB.Label Label5 
         Caption         =   "AR Date"
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
         TabIndex        =   60
         Top             =   1200
         Width           =   1740
      End
      Begin VB.Label Label1 
         Caption         =   "Location ID"
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
         Left            =   7680
         TabIndex        =   28
         Top             =   1605
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.Label LblOldDPName 
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
         Left            =   5160
         TabIndex        =   26
         Top             =   3030
         Width           =   1200
      End
      Begin VB.Label LblOldDPID 
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
         Height          =   240
         Left            =   240
         TabIndex        =   24
         Top             =   3060
         Width           =   1200
      End
      Begin VB.Label LblOldClientName 
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
         Left            =   5160
         TabIndex        =   22
         Top             =   2520
         Width           =   1200
      End
      Begin VB.Label LblOldClientID 
         Caption         =   "Client ID"
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
         Left            =   240
         TabIndex        =   20
         Top             =   2565
         Width           =   1200
      End
      Begin VB.Label LblOldCommodity 
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
         Left            =   5160
         TabIndex        =   19
         Top             =   2070
         Width           =   1500
      End
      Begin VB.Label LbOldlLocationID 
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
         Height          =   300
         Left            =   240
         TabIndex        =   17
         Top             =   2070
         Width           =   1200
      End
      Begin VB.Label LblOldCMPID 
         Caption         =   "CMP"
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
         Left            =   5160
         TabIndex        =   15
         Top             =   1575
         Width           =   1080
      End
      Begin VB.Label LblOldISINID 
         Caption         =   "Old ISIN ID"
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
         TabIndex        =   13
         Top             =   1620
         Width           =   1500
      End
      Begin VB.Label LblGrade 
         Caption         =   "Grade"
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
         Left            =   5160
         TabIndex        =   12
         Top             =   705
         Width           =   1305
      End
      Begin VB.Label Label2 
         Caption         =   "Revalidation ID"
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
         Left            =   225
         TabIndex        =   11
         Top             =   293
         Width           =   1620
      End
      Begin VB.Label LblRevalidationDate 
         Caption         =   "Date"
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
         Left            =   5160
         TabIndex        =   10
         Top             =   255
         Width           =   1740
      End
      Begin VB.Label LblARNo 
         Caption         =   "AR No"
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
         TabIndex        =   9
         Top             =   780
         Width           =   1305
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4740
      Left            =   120
      TabIndex        =   54
      Top             =   3600
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   8361
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "GSL Details"
      TabPicture(0)   =   "Frm_NewISINRevalidation.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "LblRecCount"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "MSHFlexGrid1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "ChkCheckAll"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Log History"
      TabPicture(1)   =   "Frm_NewISINRevalidation.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label20"
      Tab(1).Control(1)=   "Label21"
      Tab(1).Control(2)=   "TxtUpdated"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "TxtCreated"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
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
         TabIndex        =   61
         Top             =   480
         Width           =   1455
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
         Left            =   -73110
         Locked          =   -1  'True
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   600
         Width           =   7410
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
         Left            =   -73110
         Locked          =   -1  'True
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   1005
         Width           =   7410
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   3780
         Left            =   120
         TabIndex        =   57
         Top             =   840
         Width           =   14700
         _ExtentX        =   25929
         _ExtentY        =   6668
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
      Begin VB.Label Label3 
         Caption         =   "Total GSLs  : "
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
         Left            =   11430
         TabIndex        =   63
         Top             =   405
         Width           =   1380
      End
      Begin VB.Label LblRecCount 
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
         Height          =   345
         Left            =   12870
         TabIndex        =   62
         Top             =   405
         Width           =   1305
      End
      Begin VB.Label Label21 
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
         Left            =   -74880
         TabIndex        =   59
         Top             =   660
         Width           =   855
      End
      Begin VB.Label Label20 
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
         Left            =   -74880
         TabIndex        =   58
         Top             =   1065
         Width           =   885
      End
   End
End
Attribute VB_Name = "Frm_NewISINRevalidation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mExchangeTypeID As Double
Dim mDPID As Double
Dim mCommodityID As Double
Dim mCMPID As Double
Dim Edit_Flg As Variant
Const strChecked = "þ"
Const strUnChecked = "q"
Dim mClientIDRow As Double

Dim mSCMPName, mSLocationId As Double
Dim msSM_prefix As Variant
Private Sub AddCmd_Click()
Edit_Flg = "A"
'TxtARNo.Locked = False
'TxtOldISINID.Locked = False
'TxtClientID.Locked = False
'TxtGrade.Locked = False

Frame4.Enabled = True
SSTab1.Tab = 0
TxtOldISINID.Locked = False
DtpRevalidationDate.Value = Date
DtpARDate.Value = Date

ChkCheckAll.Enabled = True
MSHFlexGrid1.Enabled = True

TxtRevalidationID.Text = ""
TxtARNo.Text = ""
    
TxtOldISINID.Text = ""
txtOldCMP.Text = ""
TxtOldLocationID.Text = ""
TxtOldCommodity.Text = ""
TxtPreFix = ""
TxtGrade.Text = ""
TxtClientID = ""
TxtDPID = ""
TxtDPName = ""
TxtClientName = ""
mCMPID = 0
mCommodityID = 0
ChkCheckAll.Value = 0
LblRecCount.Caption = 0
Call Grid_Head
Call GButtonLock(Me, False)
End Sub

Private Sub ChkCheckAll_Click()

If SaveCmd.Enabled = False Then Exit Sub

If MSHFlexGrid1.Rows = 2 Then
   ChkCheckAll.Value = 0
   Exit Sub
End If

If ChkCheckAll.Value = vbChecked Then
   For I = 1 To MSHFlexGrid1.Rows - 2
       MSHFlexGrid1.Row = I
       MSHFlexGrid1.TextMatrix(I, 1) = strChecked
   Next
Else
   For I = 1 To MSHFlexGrid1.Rows - 2
       MSHFlexGrid1.Row = I
       MSHFlexGrid1.TextMatrix(I, 1) = strUnChecked
   Next
End If
End Sub

Private Sub CmbClientDP_LostFocus()

Dim Str_() As String
Str_ = Split(CmbClientDP.Text, "~")
If UBound(Str_) <> 1 Then
   MsgBox "Invalid selection !!!!! "
   If CmbClientDP.Visible = True Then
      CmbClientDP.SetFocus
   Else
      TxtClientID.SetFocus
   End If
   Exit Sub
End If
TmpRs1.MoveFirst
TmpRs1.Find "DepositoryID  = '" & Str_(0) & "'"
If TmpRs1.EOF Then
   MsgBox "Invalid selection !!!!! "
   CmbClientDP.SetFocus
   Exit Sub
End If

TxtClientName = TmpRs1!ClientName
TxtDPID = TmpRs1!DepositoryID
TxtDPName = TmpRs1!DepositoryName
CmbClientDP.Visible = False
mClientIDRow = TmpRs1!ClientIDRow

End Sub

Private Sub DeleteCmd_Click()
Dim ans As Variant
On Error GoTo msgError

If CheckNewISIN = True Then
   MsgBox "You can't Delete this Record."
   Exit Sub
End If

ans = MsgBox("Do you really want to DELETE this record???", vbYesNo)
If ans = vbYes Then
   Call DeleteChildTableDetail
   RsTxn_NewISINRevalidation.Delete
   RsTxn_NewISINRevalidation.Update
   If RsTxn_NewISINRevalidation.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsTxn_NewISINRevalidation.MoveNext
   If RsTxn_NewISINRevalidation.EOF() Then
      RsTxn_NewISINRevalidation.MoveLast
   End If
   Call Indata
End If
Exit Sub
msgError:
MsgBox "Child record Present "
RsTxn_NewISINRevalidation.CancelUpdate

End Sub

Private Sub EditCmd_Click()
If CheckNewISIN = True Then
   MsgBox "You can't Edit this Record."
   Exit Sub
End If
Edit_Flg = "E"
SSTab1.Tab = 0
Frame4.Enabled = True
ChkCheckAll.Enabled = True
TxtOldISINID.Locked = True
Call GButtonLock(Me, False)
End Sub
Private Function CheckNewISIN() As Boolean
Dim RetVal As Boolean
RetVal = False
For I = 1 To MSHFlexGrid1.Rows - 1
    If MSHFlexGrid1.TextMatrix(I, 8) <> "" Then
       RetVal = True
       Exit For
    End If
Next
CheckNewISIN = RetVal
End Function
Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Call GButtonLockAD(Me)
   Frame4.Enabled = False
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If
Frame4.Enabled = True
Call Grid_Head
Call TableTxn_NewISINRevalidation
mExchangeTypeID = GetExchangeID("NCDEX")
If RsTxn_NewISINRevalidation.RecordCount = 0 Then
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

Dim ans As Variant

Call Grid_Head
SSTab1.Tab = 0
TxtRevalidationID.Text = RsTxn_NewISINRevalidation!NIRID
TxtPreFix.Text = RsTxn_NewISINRevalidation!SM_Prefix
TxtARNo.Text = RsTxn_NewISINRevalidation!ARNo
DtpARDate.Value = RsTxn_NewISINRevalidation!ARDate
DtpRevalidationDate.Value = RsTxn_NewISINRevalidation!RevalDate
TxtGrade.Text = RsTxn_NewISINRevalidation!Grade
TxtOldISINID.Text = RsTxn_NewISINRevalidation!OldISIN

TxtCreated = IIf(IsNull(RsTxn_NewISINRevalidation!CreatedBy), "", RsTxn_NewISINRevalidation!CreatedBy) & " :- " & IIf(IsNull(RsTxn_NewISINRevalidation!CreatedDate), "", RsTxn_NewISINRevalidation!CreatedDate)
TxtUpdated = IIf(IsNull(RsTxn_NewISINRevalidation!UpdatedBy), "", RsTxn_NewISINRevalidation!UpdatedBy) & " :- " & IIf(IsNull(RsTxn_NewISINRevalidation!UpdatedDate), "", RsTxn_NewISINRevalidation!UpdatedDate)

mClientIDRow = RsTxn_NewISINRevalidation!ClientIDRow
ans = GetClientNameByRow(mClientIDRow)

TxtClientID.Text = ans(0)
TxtClientName.Text = ans(1)   'TmpRs!ClientName
TxtDPID.Text = ans(2)         'TmpRs!DepositoryID
TxtDPName.Text = ans(3)       'TmpRs!DepositoryName

Call GButtonLock(Me, True)
Call TxtOldISINID_LostFocus
Call GridInData

'TxtARNo.Locked = True
'TxtClientID.Locked = True
'TxtGrade.Locked = True
'TxtOldISINID.Locked = True

Frame4.Enabled = False
ChkCheckAll.Enabled = False
End Sub
Private Sub GridInData()
Call Grid_Head

tmp = "Select TNI.GSLID,MC.CMPName,MGd.GodownNo,MG.StackNo,MG.LotNo,TNI.BalanceBags,TNI.BalanceWeight,TNI.NewISIN,TNI.DRNo,TNI.DRDate,TNI.OldQty,TNI.NewQty"
tmp = tmp & " From Txn_NewISINGSLRevalidation TNI"
tmp = tmp & " left join Mst_GSL MG on TNI.GSLID=MG.GSLID And TNI.SM_Prefix=MG.SM_Prefix"
tmp = tmp & " left join Mst_Godown MGd on MG.GodownID=MGd.GodownID"
tmp = tmp & " left join Mst_CMP MC on MGd.CMPID=MC.CMPID"
tmp = tmp & " Where TNI.NIRID=" & Val(TxtRevalidationID) & ""

Call TmpTable

If TmpRs.RecordCount > 0 Then
    For I = 1 To TmpRs.RecordCount
        MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
        MSHFlexGrid1.TextMatrix(I, 0) = TmpRs!GSLID
        MSHFlexGrid1.TextMatrix(I, 2) = TmpRs!CMPName
        MSHFlexGrid1.TextMatrix(I, 3) = TmpRs!GodownNo
        MSHFlexGrid1.TextMatrix(I, 4) = TmpRs!StackNo
        MSHFlexGrid1.TextMatrix(I, 5) = TmpRs!LotNo
        MSHFlexGrid1.TextMatrix(I, 6) = TmpRs!BalanceBags
        MSHFlexGrid1.TextMatrix(I, 7) = TmpRs!BalanceWeight
        MSHFlexGrid1.TextMatrix(I, 8) = IIf(IsNull(TmpRs!NewISIN), "", TmpRs!NewISIN)
        MSHFlexGrid1.TextMatrix(I, 9) = IIf(IsNull(TmpRs!DRNo), "", TmpRs!DRNo)
        MSHFlexGrid1.TextMatrix(I, 10) = IIf(IsNull(TmpRs!DRDate), "", TmpRs!DRDate)
        MSHFlexGrid1.TextMatrix(I, 11) = IIf(IsNull(TmpRs!OldQty), "", TmpRs!OldQty)
        MSHFlexGrid1.TextMatrix(I, 12) = IIf(IsNull(TmpRs!NewQty), "", TmpRs!NewQty)
        MSHFlexGrid1.Row = I
        MSHFlexGrid1.Col = 1
        MSHFlexGrid1.CellFontName = "Wingdings"
        MSHFlexGrid1.CellFontSize = 12
        MSHFlexGrid1.CellAlignment = flexAlignCenterCenter
        MSHFlexGrid1.Text = strChecked
        TmpRs.MoveNext
    Next
End If
ChkCheckAll.Value = 1
LblRecCount.Caption = TmpRs.RecordCount

End Sub

Public Sub Grid_Head()
MSHFlexGrid1.Clear
MSHFlexGrid1.Cols = 13
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.FixedRows = 1
MSHFlexGrid1.TextMatrix(0, 0) = "GSLID"
MSHFlexGrid1.TextMatrix(0, 1) = "Select"
MSHFlexGrid1.TextMatrix(0, 2) = "CMP Name"
MSHFlexGrid1.TextMatrix(0, 3) = "Godown "
MSHFlexGrid1.TextMatrix(0, 4) = "Stack No"
MSHFlexGrid1.TextMatrix(0, 5) = "Lot No"
MSHFlexGrid1.TextMatrix(0, 6) = "No of Bags"
MSHFlexGrid1.TextMatrix(0, 7) = "Balance Qty"
MSHFlexGrid1.TextMatrix(0, 8) = "New ISIN ID"
MSHFlexGrid1.TextMatrix(0, 9) = "DR No"
'MSHFlexGrid1.TextMatrix(0, 10) = "New CMP Name  "
MSHFlexGrid1.TextMatrix(0, 10) = "DR Date"
MSHFlexGrid1.TextMatrix(0, 11) = "Old Qty"
MSHFlexGrid1.TextMatrix(0, 12) = "New Qty"

MSHFlexGrid1.ColWidth(0) = 0
MSHFlexGrid1.ColWidth(1) = 1000
MSHFlexGrid1.ColWidth(2) = 2400
MSHFlexGrid1.ColWidth(3) = 1400
MSHFlexGrid1.ColWidth(4) = 1400
MSHFlexGrid1.ColWidth(5) = 1400
MSHFlexGrid1.ColWidth(6) = 1400
MSHFlexGrid1.ColWidth(7) = 1400
MSHFlexGrid1.ColWidth(8) = 1500
MSHFlexGrid1.ColWidth(9) = 1400
'MSHFlexGrid1.ColWidth(10) = 2400
MSHFlexGrid1.ColWidth(10) = 1400
MSHFlexGrid1.ColWidth(11) = 1400
MSHFlexGrid1.ColWidth(12) = 1400

End Sub

Private Sub MSHFlexGrid1_Click()
If SaveCmd.Enabled = False Then Exit Sub
With MSHFlexGrid1
    If .TextMatrix(.Row, 1) = "" Then Exit Sub
    If .Col = 1 Then
        ChkCheckAll.Value = 0
        If .TextMatrix(.Row, 1) = strUnChecked Then
           .TextMatrix(.Row, 1) = strChecked
        Else
           .TextMatrix(.Row, 1) = strUnChecked
        End If
    End If
End With

End Sub

Private Sub LastCmd_Click()
RsTxn_NewISINRevalidation.MoveLast
Call Indata
'TmpRs.MoveLast
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False
End Sub

Private Sub MSHFlexGrid2_Click()
Dim I As Variant
I = MSHFlexGrid2.RowSel
If I <= 0 Then Exit Sub
If MSHFlexGrid2.TextMatrix(I, 0) = "" Then Exit Sub
RsTxn_NewISINRevalidation.MoveFirst
RsTxn_NewISINRevalidation.Find "NIRID = " & MSHFlexGrid2.TextMatrix(I, 0) & ""
If Not RsTxn_NewISINRevalidation.EOF Then
   Call Indata
   Frame0.Visible = False
End If
End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsTxn_NewISINRevalidation.MoveNext
If RsTxn_NewISINRevalidation.EOF Then
   RsTxn_NewISINRevalidation.MoveLast
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
RsTxn_NewISINRevalidation.MovePrevious

If RsTxn_NewISINRevalidation.BOF Then
   RsTxn_NewISINRevalidation.MoveFirst
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
RsTxn_NewISINRevalidation.MoveFirst
Call Indata
'TmpRs.MoveFirst
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub SaveCmd_Click()


If TxtARNo.Text = "" Then
    MsgBox "Blank AR No. Not Allowed !!! "
    TxtARNo.SetFocus
    Exit Sub
End If

If DtpARDate.Value > Date Then
    MsgBox "Future AR Date not Allowed!!! "
    DtpARDate.SetFocus
    Exit Sub
End If


If TxtOldISINID.Text = "" Then
    MsgBox "Blank Old ISIN ID Not Allowed !!! "
    TxtOldISINID.SetFocus
    Exit Sub
End If

If TxtClientID.Text = "" Then
    MsgBox "Blank Client ID Not Allowed !!! "
    TxtClientID.SetFocus
    Exit Sub
End If

If MSHFlexGrid1.Rows <= 2 Then
   MsgBox "Blank GSL Details Not Allowed !!! "
   TxtOldISINID.SetFocus
   Exit Sub
End If

If CheckGSLSelect = False Then
   MsgBox "Please Select Atleast one record from GSL Details"
   Exit Sub
End If


If TxtGrade.Text = "" Then
    MsgBox "Blank Grade Not Allowed !!! "
    TxtGrade.SetFocus
    Exit Sub
End If


If Edit_Flg = "A" Then
   RsTxn_NewISINRevalidation.AddNew
   RsTxn_NewISINRevalidation!NIRID = getMaxNIRID
   RsTxn_NewISINRevalidation!G_UID = GetGUID
   TxtRevalidationID.Text = RsTxn_NewISINRevalidation!NIRID
Else
   Call DeleteChildTableDetail
   RsTxn_NewISINRevalidation.MoveFirst
   RsTxn_NewISINRevalidation.Find " NIRID= " & Val(TxtRevalidationID.Text) & ""
End If

RsTxn_NewISINRevalidation!NIRID = Val(TxtRevalidationID.Text)
RsTxn_NewISINRevalidation!SM_Prefix = TxtPreFix.Text
RsTxn_NewISINRevalidation!ARNo = TxtARNo.Text
RsTxn_NewISINRevalidation!ARDate = DtpARDate.Value
RsTxn_NewISINRevalidation!RevalDate = DtpRevalidationDate.Value
RsTxn_NewISINRevalidation!Grade = TxtGrade.Text
RsTxn_NewISINRevalidation!OldISIN = TxtOldISINID.Text
RsTxn_NewISINRevalidation!ClientIDRow = Val(mClientIDRow)

If IsNull(RsTxn_NewISINRevalidation!CreatedBy) = True Or RsTxn_NewISINRevalidation!CreatedBy = "" Then
   RsTxn_NewISINRevalidation!CreatedBy = Gen_UserLoginID
   RsTxn_NewISINRevalidation!CreatedDate = Now
Else
   RsTxn_NewISINRevalidation!UpdatedBy = Gen_UserLoginID
   RsTxn_NewISINRevalidation!UpdatedDate = Now
End If

RsTxn_NewISINRevalidation.Update

Call SaveISINGSLDetails
SaveCmd.Enabled = False
Call Indata

End Sub
Private Sub SaveISINGSLDetails()
Call TableTxn_NewISINGSLRevalidation(" Where NIRID=" & Val(TxtRevalidationID) & "")
Call TableMst_GSL(" Where sm_prefix = '" & TxtPreFix & "'")
For I = 1 To MSHFlexGrid1.Rows - 1
    If MSHFlexGrid1.TextMatrix(I, 1) = strChecked Then
       RsTxn_NewISINGSLRevalidation.AddNew
       RsTxn_NewISINGSLRevalidation!G_UID = GetGUID
       RsTxn_NewISINGSLRevalidation!NIGRID = getMaxNIGRID
       RsTxn_NewISINGSLRevalidation!NIRID = Val(TxtRevalidationID)
       RsTxn_NewISINGSLRevalidation!GSLID = Val(MSHFlexGrid1.TextMatrix(I, 0))
       RsTxn_NewISINGSLRevalidation!BalanceBags = Val(MSHFlexGrid1.TextMatrix(I, 6))
       RsTxn_NewISINGSLRevalidation!BalanceWeight = Val(MSHFlexGrid1.TextMatrix(I, 7))
       RsTxn_NewISINGSLRevalidation!SM_Prefix = TxtPreFix.Text
       RsMst_GSL.Find ("GSLID = " & Val(MSHFlexGrid1.TextMatrix(I, 0)))
       If Not RsMst_GSL.EOF Then
          RsMst_GSL!ISINID = Trim(MSHFlexGrid1.TextMatrix(I, 8))
          RsMst_GSL!FirstISINID = Trim(TxtOldISINID.Text)
          RsMst_GSL!Flag1 = RsMst_GSL!Flag
          RsMst_GSL!Flag = "R"
          RsMst_GSL.Update
       End If
       RsTxn_NewISINGSLRevalidation.Update
    End If
Next
End Sub
Private Function getMaxNIRID() As Double
Dim RetVal As Double
tmp = "Select max(NIRID) from Txn_NewISINRevalidation"
Call TmpTable
RetVal = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
RetVal = RetVal + 1
getMaxNIRID = RetVal
End Function
Private Function getMaxNIGRID() As Double
Dim RetVal As Double
tmp = "Select max(NIGRID) from Txn_NewISINGSLRevalidation"
Call TmpTable
RetVal = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
RetVal = RetVal + 1
getMaxNIGRID = RetVal
End Function
Sub DeleteChildTableDetail()
Call TableMst_GSL(" Where sm_prefix = '" & TxtPreFix & "'")
For I = 1 To MSHFlexGrid1.Rows - 1
    If MSHFlexGrid1.TextMatrix(I, 0) <> "" Then
       RsMst_GSL.Find ("GSLID = " & Val(MSHFlexGrid1.TextMatrix(I, 0)))
       If Not RsMst_GSL.EOF Then
          RsMst_GSL!ISINID = RsMst_GSL!FirstISINID ' Trim(MSHFlexGrid1.TextMatrix(i, 8))
          RsMst_GSL!FirstISINID = "" 'Trim(TxtOldISINID.Text)
          RsMst_GSL!Flag = RsMst_GSL!Flag1
          RsMst_GSL.Update
       End If
    End If
Next

tmp = "Delete from Txn_NewISINGSLRevalidation where NIRID= " & Val(TxtRevalidationID.Text) & " "
Call TmpTable
End Sub
Private Function CheckGSLSelect() As Boolean
Dim RetVal As Boolean
RetVal = False

For I = 0 To MSHFlexGrid1.Rows - 2
    If MSHFlexGrid1.TextMatrix(I, 1) = strChecked Then
       RetVal = True
       Exit For
    End If
Next
CheckGSLSelect = RetVal
End Function

Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
  If RsTxn_NewISINRevalidation.RecordCount = 0 Then Exit Sub
  Call GButtonLock(Me, True)
  Call FirstCmd_Click
  Exit Sub
End If
Frame0.Height = Me.Height - 700
Frame0.Left = Frame4.Left - 50
Frame0.Top = Frame4.Top - 40
Frame0.Width = Me.Width - 300
Frame0.Visible = True
MSHFlexGrid2.Width = Frame0.Width - 200
MSHFlexGrid2.Height = Frame0.Height - 1800
Call Grid_Head1
CmbPreFix.Text = ""
CmbSCmp.Text = ""
TxtSClientName.Text = ""
TxtSOldISIN.Text = ""
TxtSARNo.Text = ""

End Sub
Public Sub Grid_Head1()
MSHFlexGrid2.Clear
MSHFlexGrid2.Cols = 10
MSHFlexGrid2.Rows = 2
MSHFlexGrid2.FixedRows = 1
MSHFlexGrid2.TextMatrix(0, 0) = "Revalidation ID"
MSHFlexGrid2.TextMatrix(0, 1) = "Revalidation Date"
MSHFlexGrid2.TextMatrix(0, 2) = "AR No"
MSHFlexGrid2.TextMatrix(0, 3) = "Grade"
MSHFlexGrid2.TextMatrix(0, 4) = "Old ISIN"
MSHFlexGrid2.TextMatrix(0, 5) = "Location"
MSHFlexGrid2.TextMatrix(0, 6) = "CMP Name"
MSHFlexGrid2.TextMatrix(0, 7) = "Client ID"
MSHFlexGrid2.TextMatrix(0, 8) = "Client Name"
MSHFlexGrid2.TextMatrix(0, 9) = "DP Name"


MSHFlexGrid2.ColWidth(0) = 1000
MSHFlexGrid2.ColWidth(1) = 1000
MSHFlexGrid2.ColWidth(2) = 1200
MSHFlexGrid2.ColWidth(3) = 1200
MSHFlexGrid2.ColWidth(4) = 1400
MSHFlexGrid2.ColWidth(5) = 1800
MSHFlexGrid2.ColWidth(6) = 1800
MSHFlexGrid2.ColWidth(7) = 1400
MSHFlexGrid2.ColWidth(8) = 1800
MSHFlexGrid2.ColWidth(9) = 1800

MSHFlexGrid2.RowHeight(0) = 600
MSHFlexGrid2.WordWrap = True

End Sub

Private Sub CmbPreFix_GotFocus()
tmp = CmbPreFix.Text

If Gen_WcLocation = "" Then
   Call TableMst_StoreManLocations("")
Else
   Call TableMst_StoreManLocations("where " & Gen_WcSMLocation)
End If


CmbPreFix.Clear
If RsMst_StoreManLocations.RecordCount = 0 Then
   MsgBox "No Location found"
   CmbPreFix.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_StoreManLocations.RecordCount
    CmbPreFix.AddItem RsMst_StoreManLocations!StoreManLocations

    RsMst_StoreManLocations.MoveNext
Next
CmbPreFix.Text = tmp

End Sub

Private Sub CmbPreFix_LostFocus()
If CmbPreFix.Text = "" Then
   Exit Sub
End If
RsMst_StoreManLocations.MoveFirst
RsMst_StoreManLocations.Find "StoreManLocations = '" & CmbPreFix.Text & "'"
If RsMst_StoreManLocations.EOF Then
   MsgBox "Invalid Location Selection !!!!"
   CmbPreFix.SetFocus
   Exit Sub
End If
msSM_prefix = RsMst_StoreManLocations!SM_Prefix
mSLocationId = RsMst_StoreManLocations!StoreManLocationsID
End Sub

Private Sub CmbSCmp_GotFocus()
If CmbPreFix.Text = "" Then
   MsgBox "Plase Select Location !!!"
   CmbPreFix.SetFocus
   Exit Sub
End If

tmp1 = ""
If Gen_WcCMP <> "" Then
   tmp1 = "And " & Gen_WcCMP
End If

tmp = CmbSCmp.Text
Call TableMst_CMP(" Where LocationID = " & mSLocationId & "" & tmp1)
CmbSCmp.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_CMP.RecordCount
    CmbSCmp.AddItem RsMst_CMP!CMPName
    RsMst_CMP.MoveNext
Next
CmbSCmp.Text = tmp

End Sub

Private Sub CmbSCmp_LostFocus()
If CmbSCmp.Text = "" Then
   Exit Sub
End If
RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPName = '" & CmbSCmp.Text & "'"
If RsMst_CMP.EOF Then
   MsgBox "Invalid selection "
   CmbSCmp.SetFocus
   Exit Sub
End If
mSCMPName = RsMst_CMP!CMPID
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Txn_NewISINRevalidation.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Txn_NewISINRevalidation.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid2.Rows - 1
    mRow = I + 1 ' + 6
    For c = 0 To MSHFlexGrid2.Cols - 1
        If InStr(MSHFlexGrid2.TextMatrix(I, c), "/") > 0 Then
           If IsDate(MSHFlexGrid2.TextMatrix(I, c)) Then
              oWS.Cells(mRow, c + 1) = Format(MSHFlexGrid2.TextMatrix(I, c), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, c + 1) = MSHFlexGrid2.TextMatrix(I, c)
           End If
        Else
           oWS.Cells(mRow, c + 1) = MSHFlexGrid2.TextMatrix(I, c)
        End If
    Next
Next
oWB.Save
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub
Private Sub CmdSearch_Click()
Dim wc As String
Dim tmp1 As String
Call Grid_Head1

tmp = "Select TIR.NIRID,TIR.RevalDate,TIR.ARNo,TIR.Grade,TIR.OldISIN,ML.LocationName"
tmp = tmp & " ,MC.CMPName,MCl.ClientID,MCl.ClientName,MD.DepositoryName"
tmp = tmp & " From Txn_NewISINRevalidation TIR"
tmp = tmp & " Inner join Mst_ISIN MI on TIR.OldISIN=MI.ISINID"
tmp = tmp & " Inner join Mst_Location ML on MI.LocationID=ML.LocationID"
tmp = tmp & " Inner join Mst_CMP MC on MI.CMPID=MC.CMPID"
tmp = tmp & " Inner join Mst_Client MCl on TIR.ClientIDRow=MCl.ClientIDRow"
tmp = tmp & " Inner join Mst_Depository MD on MCl.DepositoryID=MD.DepositoryID "

If CmbPreFix.Text <> "" Then
   wc = GenWc(wc, CmbPreFix.Text, "ML.LocationName", "S", "N")
Else
   If Gen_WcLocation <> "" Then
      tmp1 = Replace(Gen_WcLocation, "LocationName", "ML.LocationName")
      wc = GenWc(wc, "", tmp1, "S", "N")
      'Wc1 = Wc1 & " And  ML." & Gen_WcLocation
   End If
End If

If CmbSCmp.Text <> "" Then
   wc = GenWc(wc, CmbSCmp.Text, "MC.CMPName", "S", "N")
Else
   If Gen_WcCMP <> "" Then
      tmp1 = Replace(Gen_WcLocation, "CMPName", "MC.CMPName")
      wc = GenWc(wc, "", tmp1, "S", "N")
     'Wc1 = Wc1 & " And  MC." & Gen_WcCMP
   End If
End If

If TxtSClientName.Text <> "" Then
   wc = GenWc(wc, TxtSClientName.Text, "MCl.ClientName", "S", "L")
End If

If TxtSOldISIN.Text <> "" Then
   wc = GenWc(wc, TxtSOldISIN.Text, "TIR.OldISIN", "S", "L")
End If

If TxtSARNo.Text <> "" Then
   wc = GenWc(wc, TxtSARNo.Text, "TIR.ARNo", "S", "L")
End If
tmp = tmp & wc

tmp = tmp & " order By TIR.NIRID"

Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   MSHFlexGrid2.Rows = TmpRs1.RecordCount + 2
   For I = 1 To TmpRs1.RecordCount
       For c = 0 To MSHFlexGrid2.Cols - 1
           MSHFlexGrid2.TextMatrix(I, c) = IIf(IsNull(TmpRs1.Fields(c)), "", TmpRs1.Fields(c))
       Next
       MSHFlexGrid2.TextMatrix(I, 1) = Format(MSHFlexGrid2.TextMatrix(I, 1), "dd-mmm-yyyy")
       TmpRs1.MoveNext
   Next
End If

End Sub

Private Sub TxtClientID_LostFocus()


CmbClientDP.Visible = False
If TxtClientID = "" Then
   TxtClientName = ""
   TxtDPID = ""
   TxtDPName = ""
   CmbClientDP.Text = ""
   CmbClientDP.Visible = False
   Exit Sub
End If
Dim ans As Variant
ans = GetClientNameDtl(TxtClientID, mExchangeTypeID)

If ans(0, 0) = False Then
   MsgBox "Not Found !!!"
   TxtClientID.SetFocus
   Exit Sub
End If
If ans(3, 0) > 1 Then
   CmbClientDP.Visible = True
   CmbClientDP.SetFocus
End If
CmbClientDP.Text = ""
CmbClientDP.Clear
tmp = "SELECT Mst_Depository.DepositoryID, Mst_Depository.DepositoryName, Mst_Client.ClientID, " & _
      "Mst_Client.ClientName,  Mst_Client.ClientIDRow " & _
      "FROM Mst_Depository INNER JOIN Mst_Client ON Mst_Depository.DepositoryID = Mst_Client.DepositoryID " & _
      "WHERE Mst_Depository.ExchangeTypeID = " & mExchangeTypeID & _
      "  And  (((Mst_Client.ExchangeTypeID)=" & mExchangeTypeID & ") AND ((Mst_Client.ClientID)='" & TxtClientID & "'));"

Call Tmp1Table

For op = 1 To TmpRs1.RecordCount
    If CmbClientDP.Text = "" Then
       CmbClientDP.Text = TmpRs1!DepositoryID & " ~ " & TmpRs1!DepositoryName
    End If
    CmbClientDP.AddItem TmpRs1!DepositoryID & " ~ " & TmpRs1!DepositoryName
    TmpRs1.MoveNext
Next
If ans(3, 0) <= 1 Then
   CmbClientDP_LostFocus
End If


End Sub

Private Sub TxtOldISINID_LostFocus()

If TxtOldISINID.Text = "" Then
    mCommodityID = 0
    mCMPID = 0
    txtOldCMP = ""
    TxtOldLocationID = ""
    TxtOldCommodity = ""
    Exit Sub
End If

'TxtOldISINID.Text = "INC000" & Trim(TxtOldISINID.Text)

Call TableMst_ISIN("Where ISINID = '" & Trim(TxtOldISINID.Text) & "' ")

If RsMst_ISIN.RecordCount = 0 Then
   MsgBox "No record found !!!"
   TxtOldISINID.SetFocus
   Exit Sub
End If
mCMPID = RsMst_ISIN!CMPID
txtOldCMP = GetCMPName(RsMst_ISIN!CMPID)
mLocationID = IIf(IsNull(RsMst_ISIN!LocationID), "", RsMst_ISIN!LocationID)
TxtOldLocationID = GetLocationName(RsMst_ISIN!LocationID)
TxtPreFix = GetPreFixByLocation(RsMst_ISIN!LocationID, "P")
mCommodityID = IIf(IsNull(RsMst_ISIN!CommodityID), "", RsMst_ISIN!CommodityID)
TxtOldCommodity = GetCommodityName(RsMst_ISIN!CommodityID)
'TxtGrade = IIf(IsNull(RsMst_ISIN!Grade), "", RsMst_ISIN!Grade)

If SaveCmd.Enabled = False Then Exit Sub
If Edit_Flg = "E" Then Exit Sub


Call GetGridData

End Sub
Private Sub GetGridData()
Call Grid_Head
If SaveCmd.Enabled = False Then Exit Sub

tmp = "SELECT Mst_GSL.GSLID,MC.CMPName,Mst_Godown.GodownNo, Mst_GSL.StackNo, Mst_GSL.LotNo, Mst_GSL.BalanceBags, Mst_GSL.BalanceWeight"
tmp = tmp & " From Mst_Godown"
tmp = tmp & " INNER JOIN Mst_GSL ON Mst_Godown.GodownID = Mst_GSL.GodownID"
tmp = tmp & " Inner Join Mst_CMP MC on Mst_Godown.CMPID=MC.CMPID"
tmp = tmp & " where Mst_GSL.ExchangeTypeID =" & mExchangeTypeID & " And (Mst_GSL.Flag='O' or Mst_GSL.Flag='R') And Mst_GSL.ISINID ='" & Trim(TxtOldISINID.Text) & "'"
tmp = tmp & " Order By MC.CMPName"

Call TmpTable

If TmpRs.RecordCount <= 0 Then
  MsgBox "No Record Found For Grid!!!"
  Exit Sub
End If
k = 1
For I = 0 To TmpRs.RecordCount - 1
    With MSHFlexGrid1
        .TextMatrix(k, 0) = TmpRs!GSLID
        .TextMatrix(k, 2) = TmpRs!CMPName
        .TextMatrix(k, 3) = TmpRs!GodownNo
        .TextMatrix(k, 4) = TmpRs!StackNo
        .TextMatrix(k, 5) = TmpRs!LotNo
        .TextMatrix(k, 6) = TmpRs!BalanceBags
        .TextMatrix(k, 7) = TmpRs!BalanceWeight
        .Row = k
        .Col = 1
        .CellFontName = "Wingdings"
        .CellFontSize = 12
        .CellAlignment = flexAlignCenterCenter
        .Text = strUnChecked
        .Rows = .Rows + 1
        k = k + 1
    End With
   TmpRs.MoveNext
Next
LblRecCount.Caption = TmpRs.RecordCount
LblRecCount.Refresh
End Sub

'
'Private Sub TxtSOldISIN_LostFocus()
'TxtSOldISINID.Text = "INC000" & Trim(TxtSOldISINID.Text)
'End Sub
