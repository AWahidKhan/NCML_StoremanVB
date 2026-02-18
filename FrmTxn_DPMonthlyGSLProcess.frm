VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmTxn_DPMonthlyGSLProcess 
   Caption         =   "Franchisee Fees/ DP Op Stock Process"
   ClientHeight    =   9120
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12690
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9120
   ScaleWidth      =   12690
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame6 
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   720
      TabIndex        =   27
      Top             =   3720
      Visible         =   0   'False
      Width           =   5655
      Begin VB.CommandButton CmdCancel 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cancel"
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
         Left            =   3120
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   840
         Width           =   2055
      End
      Begin VB.CommandButton CmdAddDummyLots 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Add Dummy Lots"
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
         Left            =   1080
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   840
         Width           =   2055
      End
      Begin VB.ComboBox CmbLocation 
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
         Left            =   1200
         Sorted          =   -1  'True
         TabIndex        =   32
         Top             =   360
         Width           =   4335
      End
      Begin VB.ComboBox CmbStateName 
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
         ItemData        =   "FrmTxn_DPMonthlyGSLProcess.frx":0000
         Left            =   1200
         List            =   "FrmTxn_DPMonthlyGSLProcess.frx":0010
         Sorted          =   -1  'True
         TabIndex        =   28
         Top             =   480
         Visible         =   0   'False
         Width           =   4335
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Location"
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
         TabIndex        =   33
         Top             =   420
         Width           =   900
      End
      Begin VB.Label Label55 
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "State "
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
         TabIndex        =   31
         Top             =   660
         Visible         =   0   'False
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Height          =   8820
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   15090
      Begin VB.Frame Frame5 
         Height          =   2775
         Left            =   8280
         TabIndex        =   42
         Top             =   3360
         Width           =   6375
         Begin VB.CommandButton CmdUpdateGodownCode 
            Caption         =   "Update Godown No"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1680
            TabIndex        =   51
            Top             =   2160
            Width           =   2040
         End
         Begin VB.ComboBox CmbGodownNo 
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
            Left            =   1680
            Sorted          =   -1  'True
            TabIndex        =   49
            Top             =   1200
            Width           =   3240
         End
         Begin VB.ComboBox CmbOldLocation 
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
            Left            =   1680
            Sorted          =   -1  'True
            TabIndex        =   45
            Top             =   240
            Width           =   2805
         End
         Begin VB.ComboBox CmbCMPName 
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
            Left            =   1680
            Sorted          =   -1  'True
            TabIndex        =   44
            Top             =   720
            Width           =   4440
         End
         Begin VB.TextBox TxtNewGodowNo 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   1680
            MaxLength       =   15
            TabIndex        =   43
            Top             =   1680
            Width           =   1875
         End
         Begin VB.Label lblGodownNo 
            Caption         =   "Old Godown No"
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
            Left            =   195
            TabIndex        =   50
            Top             =   1253
            Width           =   1440
         End
         Begin VB.Label Label20 
            Caption         =   "Location"
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
            Left            =   195
            TabIndex        =   48
            Top             =   270
            Width           =   1005
         End
         Begin VB.Label Label17 
            Caption         =   "CMP"
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
            Left            =   195
            TabIndex        =   47
            Top             =   780
            Width           =   1695
         End
         Begin VB.Label Label13 
            Caption         =   "New Godown No"
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
            Left            =   195
            TabIndex        =   46
            Top             =   1800
            Width           =   1335
         End
      End
      Begin VB.Frame Frame4 
         Height          =   1095
         Left            =   8280
         TabIndex        =   37
         Top             =   1320
         Visible         =   0   'False
         Width           =   6375
         Begin VB.CommandButton CmdDeleteFranchProcess 
            Caption         =   "Franchaisee Fees Process Deletion"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   4080
            TabIndex        =   40
            Top             =   240
            Width           =   2040
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
            Left            =   360
            Sorted          =   -1  'True
            TabIndex        =   38
            Top             =   405
            Width           =   3300
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
            Left            =   1530
            TabIndex        =   39
            Top             =   120
            Width           =   840
         End
      End
      Begin VB.CommandButton CmdReserDistri 
         Caption         =   "Reservation Income Distribute"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   12225
         TabIndex        =   35
         Top             =   360
         Visible         =   0   'False
         Width           =   2040
      End
      Begin VB.CommandButton CmdCommtrack 
         Caption         =   "Commtrack Income Distribute"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   10080
         TabIndex        =   34
         Top             =   360
         Visible         =   0   'False
         Width           =   2040
      End
      Begin VB.CommandButton CmdDPDistribute 
         Caption         =   "DP Income Distribute"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   8280
         TabIndex        =   26
         Top             =   360
         Width           =   1680
      End
      Begin VB.CommandButton CmdAddDummyLot 
         Caption         =   "Add dummy lot for Fixed franchisee godowns"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   105
         TabIndex        =   25
         Top             =   8100
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.CommandButton CmdUpdatePendingRateCard 
         Caption         =   "Update Pending Rate Card"
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
         Left            =   2535
         TabIndex        =   24
         Top             =   8340
         Visible         =   0   'False
         Width           =   2745
      End
      Begin VB.ComboBox CmbProcessType 
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
         ItemData        =   "FrmTxn_DPMonthlyGSLProcess.frx":002C
         Left            =   120
         List            =   "FrmTxn_DPMonthlyGSLProcess.frx":002E
         Sorted          =   -1  'True
         TabIndex        =   22
         Top             =   375
         Width           =   2055
      End
      Begin VB.Frame Frame3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1650
         Left            =   8280
         TabIndex        =   19
         Top             =   8520
         Visible         =   0   'False
         Width           =   8055
         Begin VB.TextBox TxtConfirmation 
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
            Height          =   765
            Left            =   120
            MaxLength       =   1000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   20
            Top             =   765
            Width           =   7845
         End
         Begin VB.Label Label8 
            Caption         =   "Enter the words 'All Checkpoint items are completed' in the box below and then press the 'process' key to proceed further"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Width           =   7800
         End
      End
      Begin VB.TextBox TxtRateCard 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   585
         Left            =   10080
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Top             =   9120
         Visible         =   0   'False
         Width           =   4845
      End
      Begin VB.CommandButton CmdShowRateCard 
         Caption         =   "Show Rate Card Report"
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
         Height          =   585
         Left            =   8400
         TabIndex        =   16
         Top             =   9120
         Visible         =   0   'False
         Width           =   1560
      End
      Begin VB.CommandButton CmdProcess 
         Caption         =   "Process"
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
         Left            =   5295
         TabIndex        =   9
         Top             =   8340
         Width           =   1440
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "Exit"
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
         Left            =   6735
         TabIndex        =   8
         Top             =   8340
         Width           =   1440
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
         Left            =   4635
         TabIndex        =   7
         Top             =   375
         Width           =   1620
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
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   8760
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.CommandButton CmdShowLocation 
         Caption         =   "Show Locations"
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
         Left            =   6360
         TabIndex        =   5
         Top             =   360
         Width           =   1785
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
         Left            =   2160
         TabIndex        =   4
         Top             =   360
         Width           =   2415
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
         TabIndex        =   3
         Top             =   795
         Width           =   1335
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   6945
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   12250
         _Version        =   393216
         FixedCols       =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label5 
         Caption         =   $"FrmTxn_DPMonthlyGSLProcess.frx":0030
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   8430
         TabIndex        =   36
         Top             =   9960
         Visible         =   0   'False
         Width           =   6465
      End
      Begin VB.Label Label1 
         Caption         =   "Process Type"
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
         Left            =   360
         TabIndex        =   23
         Top             =   120
         Width           =   1680
      End
      Begin VB.Label Label10 
         Caption         =   "Enter ""Rate Card Report Completed"" in below box and press tab."
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
         Left            =   8400
         TabIndex        =   18
         Top             =   9600
         Visible         =   0   'False
         Width           =   6705
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
         Left            =   5130
         TabIndex        =   15
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label36 
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
         Left            =   120
         TabIndex        =   14
         Top             =   8400
         Visible         =   0   'False
         Width           =   885
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
         Left            =   3060
         TabIndex        =   13
         Top             =   135
         Width           =   600
      End
      Begin VB.Label Label7 
         Caption         =   "Total Location Found :"
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
         Left            =   4335
         TabIndex        =   12
         Top             =   780
         Width           =   2355
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
         Left            =   6735
         TabIndex        =   11
         Top             =   780
         Width           =   90
      End
   End
   Begin VB.Frame Frame2 
      Height          =   660
      Left            =   120
      TabIndex        =   0
      Top             =   8805
      Width           =   15090
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   75
         TabIndex        =   1
         Top             =   210
         Width           =   14940
         _ExtentX        =   26353
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
      Height          =   4815
      Left            =   15240
      TabIndex        =   41
      Top             =   120
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   8493
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
Attribute VB_Name = "FrmTxn_DPMonthlyGSLProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mMonth, mFromDate, mMonthListindex, mSM_Prefix, mTotalDays, mToDate, mNoofDays As Variant
Const strChecked = "þ"
Const strUnChecked = "q"
Dim mGenInvoiceLocation As Variant
Dim mCountLocation As Variant
'Dim mShowDTD As Boolean
Dim mKRNo, mDRDate, mDRISINID As Variant
Dim mLocationID As Variant
Dim mCMPID, mGodownID As Variant

Private Sub ChkCheckAll_Click()
If MSHFlexGrid1.Rows = 2 Then
   ChkCheckAll.Value = 0
   Exit Sub
End If

If ChkCheckAll.Value = vbChecked Then
   For I = 1 To MSHFlexGrid1.Rows - 2
       MSHFlexGrid1.row = I
       MSHFlexGrid1.Col = 0
       MSHFlexGrid1.TextMatrix(I, 0) = strChecked
   Next
Else
   For I = 1 To MSHFlexGrid1.Rows - 2
       MSHFlexGrid1.row = I
       MSHFlexGrid1.TextMatrix(I, 0) = strUnChecked
   Next
End If
End Sub

Private Sub CmbLocation_GotFocus()
tmp1 = CmbLocation.Text
Call TableMst_Location("")


tmp = " Select Distinct ML.LocationID,ML.LocationName,ML.SM_Prefix from Mst_FF_RateCard_Godown MFG"
tmp = tmp & " inner join Mst_FF_RateCard MF on MF.FF_RateCardID=MFG.FF_RateCardID"
tmp = tmp & " inner join mst_location ml on ml.locationid = mf.locationid"
tmp = tmp & " Where MFG.BillingUnit='F' " 'And ML.SM_Prefix not in (Select Distinct SM_Prefix From Txn_InvMonGSLDetail Where YearMonth = '" & mFromDate & "' )"
tmp = tmp & " Order By ML.LocationName"

Call TmpTable

CmbLocation.Clear
If TmpRs.RecordCount = 0 Then
   MsgBox "No Location found!!!"
   CmbLocation.SetFocus
   Exit Sub
End If
For I = 1 To TmpRs.RecordCount
   CmbLocation.AddItem TmpRs!Locationname
   TmpRs.MoveNext
Next
CmbLocation.Text = tmp1
End Sub

Private Sub CmbLocation_LostFocus()
If Trim(CmbLocation.Text) = "" Then Exit Sub

TmpRs.MoveFirst
TmpRs.Find "LocationName = '" & CmbLocation.Text & "'"

If TmpRs.EOF Then
   MsgBox "Invalid Location selection!!!"
   CmbLocation.SetFocus
   Exit Sub
End If

End Sub

Private Sub CmbMonth_GotFocus()

tmp = CmbMonth.Text
CmbMonth.Clear
For I = 1 To 12
    CmbMonth.AddItem MonthName(I), I - 1
Next

CmbMonth.Text = tmp

End Sub

Private Sub CmbMonth_LostFocus()
If CmbMonth.Text = "" Then Exit Sub
mStatus = False
For I = 0 To 11
    If Trim(CmbMonth.Text) = CmbMonth.List(I) Then
       CmbMonth.Text = CmbMonth.List(I)
       mMonthListindex = I '+ 1
       mStatus = True
       Exit For
    End If
Next

If mStatus = False Then
   MsgBox "Invalid Selection"
   CmbMonth.SetFocus
   Exit Sub
End If

End Sub


Private Sub CmbProcessType_GotFocus()
tmp = CmbProcessType.Text
CmbProcessType.Clear
CmbProcessType.AddItem "Franchisee"
If LCase(Gen_UserRole) = "power" Then
   CmbProcessType.AddItem "DP"
End If
CmbProcessType.Text = tmp
End Sub

Private Sub CmbProcessType_LostFocus()
If CmbProcessType.Text = "" Then Exit Sub
If LCase(Gen_UserRole) = "power" Then
   If LCase(CmbProcessType) <> "dp" And LCase(CmbProcessType) <> "franchisee" Then
      MsgBox "Invalid selection!!"
      CmbProcessType.SetFocus
      Exit Sub
   End If
Else
   If LCase(CmbProcessType) <> "franchisee" Then
      MsgBox "Invalid selection!!"
      CmbProcessType.SetFocus
      Exit Sub
   End If
End If
If tmp <> CmbProcessType.Text Then
   Call Grid_Head
End If

If LCase(CmbProcessType) = "dp" Then
   CmdUpdatePendingRateCard.Enabled = False
   CmdAddDummyLot.Enabled = False
ElseIf LCase(CmbProcessType) = "franchisee" Then
   CmdUpdatePendingRateCard.Enabled = True
   CmdAddDummyLot.Enabled = True
End If

End Sub

Private Sub CmbStateName_GotFocus()
tmp = CmbStateName.Text

Call TableMst_State

CmbStateName.Clear
If RsMst_State.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_State.RecordCount
    CmbStateName.AddItem RsMst_State!StateName
    RsMst_State.MoveNext
Next
CmbStateName.Text = tmp
End Sub

Private Sub CmbStateName_LostFocus()
If CmbStateName.Text = "" Then Exit Sub

RsMst_State.MoveFirst
RsMst_State.Find "StateName = '" & CmbStateName.Text & "'"
If RsMst_State.EOF Then
   MsgBox "Invalid selection!!!"
   CmbStateName.SetFocus
   Exit Sub
End If

End Sub

Private Sub CmbYear_GotFocus()
tmp = CmbYear.Text
CmbYear.Clear
For I = 2010 To Year(Now)
    CmbYear.AddItem I
Next
CmbYear.Text = tmp
End Sub

Private Sub CmbYear_LostFocus()

If CmbYear.Text = "" Then Exit Sub
Dim flg As Variant
flg = "N"
For I = 2010 To Year(Now)
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

End Sub

Private Sub CmdAddDummyLot_Click()
If CmbProcessType.Text = "" Then
   MsgBox " Blank Process Type not allowed  !!!"
   CmbProcessType.SetFocus
   Exit Sub
End If

If LCase(CmbProcessType) <> "franchisee" Then
   MsgBox "Process type must be 'Franchisee'!!!"
   CmbProcessType.SetFocus
   Exit Sub
End If

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



mMonth = mMonthListindex + 1

If mMonth < 10 Then
   mMonth = "0" & mMonth
End If

mFromDate = Format(CmbYear.Text & "-" & mMonth & "-" & "01", "yyyy-mm-dd")
mYear = Val(CmbYear.Text)
mTotalDays = Day(DateSerial(mYear, mMonth + 1, 0))

mToDate = Format(CmbYear.Text & "-" & mMonth & "-" & mTotalDays, "yyyy-mm-dd")
mNoofDays = ""

If CDate(Format(mFromDate, "yyyymm")) >= Format(Date, "yyyymm") Then
   MsgBox "Present and Future Month not Allowed !!!"
   CmbMonth.SetFocus
   Exit Sub
End If

If Format(Gen_BlockingDate, "yyyy-mm") >= Format(mFromDate, "yyyy-mm") Then
   MsgBox "Select next month from " & Format(Gen_BlockingDate, "MMM-yyyy") & "!!! "
   CmbMonth.SetFocus
   Exit Sub
End If


ProgressBar1.Value = 0

''comment on 1st march 2013=========================
'CmdAddDummyLot.Enabled = False
'If LCase(CmbProcessType) = "franchisee" Then
'   Call AddDummyLot
'End If


'Call AddDummyLot
'MsgBox "Done!!"
'''''''''''''''''''''''''''''''''''''''''''''
CmbLocation.Text = ""
Frame6.Visible = True
Frame1.Enabled = False

End Sub

Private Sub CmdAddDummyLots_Click()
'If Trim(CmbStateName.Text) = "" Then
'   MsgBox "Blank State name not allowed!!!"
'   CmbStateName.SetFocus
'   Exit Sub
'End If
If Trim(CmbLocation.Text) = "" Then
   MsgBox "Blank Location name not allowed!!!"
   CmbLocation.SetFocus
   Exit Sub
End If
Call AddDummyLot
MsgBox "Done!!"
Frame6.Visible = False
Frame1.Enabled = True

End Sub

Private Sub CmdCancel_Click()
Frame6.Visible = False
Frame1.Enabled = True
End Sub

Private Sub CmdCommtrack_Click()
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

mMonth = mMonthListindex + 1

If mMonth < 10 Then
   mMonth = "0" & mMonth
End If

mFromDate = Format(CmbYear.Text & "-" & mMonth & "-" & "01", "yyyy-mm-dd")
mYear = Val(CmbYear.Text)
mTotalDays = Day(DateSerial(mYear, mMonth + 1, 0))

mToDate = Format(CmbYear.Text & "-" & mMonth & "-" & mTotalDays, "yyyy-mm-dd")

If CDate(Format(mFromDate, "yyyymm")) >= Format(Date, "yyyymm") Then
   MsgBox "Present and Future Month not Allowed!!!"
   CmbMonth.SetFocus
   Exit Sub
End If

Dim ans As Variant


'''''commtrack amount distribution in txn_invmongsldetail TABLE
    
tmp = "Select Distinct Mst_GSL_ID From Txn_InvMonGSLDetail Where ExchangeTypeID = 2 and YearMonth = '" & mFromDate & "' and ISINID is null "
Call Tmp2Table

For k = 1 To TmpRs2.RecordCount
    tmp = "Select Mst_GSL_ID , ISINID from Mst_GSL Where Mst_GSL_ID = " & TmpRs2!Mst_GSL_ID & " and ISINID is not null"
    Call Tmp3Table
    If TmpRs3.RecordCount > 0 Then
       mISINID = TmpRs3!ISINID
       mTxn_GSL_ID = TmpRs3!Mst_GSL_ID
    
       tmp = "Update Txn_InvMonGSLDetail Set ISINID = '" & mISINID & "' Where ExchangeTypeID = 2 and YearMonth = '" & mFromDate & "' and ISINID is null And Mst_GSL_ID = " & mTxn_GSL_ID & ""
       Call Tmp3Table
    End If
    TmpRs2.MoveNext
Next

    
    
tmp = " Select TMCF.LotID, sum(TMCF.Amount) Amount from Txn_MonthlyCommtrackFileDetail TMCF"
tmp = tmp & " where TMCF.Amount > 0 And TMCF.monthyear between '" & mFromDate & "' and '" & mToDate & "'"
tmp = tmp & " group by TMCF.LotID"
Call Tmp2Table

tmp = " Update tim set CommtrackIncome = null, isinid = mgsl.isinid from Txn_InvMonGSLDetail tim"
tmp = tmp & " inner join mst_gsl mgsl on mgsl.Mst_GSL_ID = tim.Mst_GSL_ID and tim.sm_prefix = mgsl.sm_prefix"
tmp = tmp & " where tim.exchangetypeid = 2 and yearmonth between '" & mFromDate & "' and '" & mToDate & "'"
tmp = tmp & " and tim.commtrackflag = 1 and isnull(mgsl.isinid, '') <> ''"
Call Tmp3Table
    
tmp = " select * from Txn_InvMonGSLDetail "
tmp = tmp & " where exchangetypeid = 2 and yearmonth between '" & mFromDate & "' and '" & mToDate & "'"
tmp = tmp & " and commtrackflag = 1 and isnull(isinid,'') <> '' "

Call Tmp3Table

If TmpRs2.RecordCount > 0 And TmpRs3.RecordCount > 0 Then
    ProgressBar1.Value = 0
    ProgressBar1.Max = TmpRs2.RecordCount + 2
    For k = 1 To TmpRs2.RecordCount
    
        TmpRs3.Find "ISINID = '" & TmpRs2!LotId & "'"   'TmpRs3.Filter = "ISINID = '" & TmpRs2!LotId & "'"
        If TmpRs3.EOF = False Then
           TmpRs3!CommtrackIncome = TmpRs2!Amount
           TmpRs3.Update
        'Else
        '   MsgBox "stop"
        End If
        'TmpRs3.Filter = ""
        TmpRs3.MoveFirst
        TmpRs2.MoveNext
        ProgressBar1.Value = ProgressBar1.Value + 1
    Next
    ProgressBar1.Value = ProgressBar1.Max
End If

MsgBox "Done!!!"

End Sub

Private Sub CmdDPDistribute_Click()
If LCase(CmbProcessType.Text) <> "dp" Then
   MsgBox " Select Process Type 'DP'!!!"
   CmbProcessType.SetFocus
   Exit Sub
End If

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

mMonth = mMonthListindex + 1

If mMonth < 10 Then
   mMonth = "0" & mMonth
End If

mFromDate = Format(CmbYear.Text & "-" & mMonth & "-" & "01", "yyyy-mm-dd")
mYear = Val(CmbYear.Text)
mTotalDays = Day(DateSerial(mYear, mMonth + 1, 0))

mToDate = Format(CmbYear.Text & "-" & mMonth & "-" & mTotalDays, "yyyy-mm-dd")

If CDate(Format(mFromDate, "yyyymm")) >= Format(Date, "yyyymm") Then
   MsgBox "Present and Future Month not Allowed!!!"
   CmbMonth.SetFocus
   Exit Sub
End If

Dim ans As Variant

'If LCase(CmbProcessType.Text) = "dp" Then
'    tmp = " update tig set dpincome = isnull(tig.InvoiceAmount,0)* (select isnull(sum(MDP.Amt),0) from Mst_DPInvoiceUpload MDP  inner join mst_location ml on ml.locationid = MDP.locationid"
'    tmp = tmp & " where ml.sm_prefix = tig.sm_prefix and monthyear between '" & mFromDate & "' and '" & mToDate & "' and isnull(MDP.ReverseFlag,'')<>'Y')/(select isnull(sum(InvoiceAmount),0) from txn_invmongsldetail where exchangetypeid = 2 and yearmonth between '" & mFromDate & "' and '" & mToDate & "' and sm_prefix = tig.sm_prefix)"
'    tmp = tmp & " from txn_invmongsldetail  tig where yearmonth between '" & mFromDate & "' and '" & mToDate & "' and isnull(tig.InvoiceAmount,0) > 0"
'    tmp = tmp & " and exchangetypeid = 2"
'    Call Tmp2Table
'End If


If LCase(CmbProcessType.Text) = "dp" Then
    tmp = " select ml.sm_prefix, mdp.locationid, isnull(sum(MDP.Amt),0) dpamount from Mst_DPInvoiceUpload MDP"
    tmp = tmp & " inner join mst_location ml on ml.locationid = MDP.locationid"
    tmp = tmp & " where monthyear between '" & mFromDate & "' and '" & mToDate & "' and isnull(MDP.ReverseFlag,'')<>'Y'"
    tmp = tmp & " group by ml.sm_prefix,  mdp.locationid"
    Call Tmp2Table
   
    tmp = " select sm_prefix, isnull(sum(InvoiceAmount),0) gslamount from txn_invmongsldetail"
    tmp = tmp & " where exchangetypeid = 2 and isnull(InvoiceAmount,0) > 0 and yearmonth between '" & mFromDate & "' and '" & mToDate & "'"
    tmp = tmp & " and isnull(commtrackflag,0) = 0"
    tmp = tmp & " group by sm_prefix"
    Call Tmp3Table
    
    If TmpRs2.RecordCount > 0 And TmpRs3.RecordCount > 0 Then
        For k = 1 To TmpRs2.RecordCount
            TmpRs3.MoveFirst
            TmpRs3.Find "SM_Prefix = '" & TmpRs2!SM_Prefix & "'"
            If Not TmpRs3.EOF Then
                tmp = " update txn_invmongsldetail set dpincome = isnull(InvoiceAmount,0) * " & TmpRs2!dpAmount & "/" & TmpRs3!GSLAmount
                tmp = tmp & " where exchangetypeid = 2 and isnull(InvoiceAmount,0) > 0 and yearmonth between '" & mFromDate & "' and '" & mToDate & "'"
                tmp = tmp & " and sm_prefix = '" & TmpRs2!SM_Prefix & "'"
                tmp = tmp & " and isnull(commtrackflag,0) = 0"
                Call Tmp4Table
            End If
            TmpRs2.MoveNext
        Next
    End If
    
    '''''commtrack amount distribution in txn_invmongsldetail TABLE
'
'    tmp = " select TMCF.LotID, sum(TMCF.Amount) Amount from Txn_MonthlyCommtrackFileDetail TMCF"
'    tmp = tmp & " where TMCF.Amount > 0 And TMCF.monthyear between '" & mFromDate & "' and '" & mToDate & "'"
'    tmp = tmp & " group by TMCF.LotID"
'    Call Tmp2Table
'
'    tmp = " update tim set CommtrackIncome = null, isinid = mgsl.isinid from Txn_InvMonGSLDetail tim"
'    tmp = tmp & " inner join mst_gsl mgsl on mgsl.gslid = tim.gslid and tim.sm_prefix = mgsl.sm_prefix"
'    tmp = tmp & " where tim.exchangetypeid = 2 and yearmonth between '" & mFromDate & "' and '" & mToDate & "'"
'    tmp = tmp & " and tim.commtrackflag = 1 and isnull(mgsl.isinid, '') <> ''"
'    Call Tmp3Table
'
'    tmp = " select * from Txn_InvMonGSLDetail "
'    tmp = tmp & " where exchangetypeid = 2 and yearmonth between '" & mFromDate & "' and '" & mToDate & "'"
'    tmp = tmp & " and commtrackflag = 1 and isnull(isinid,'') <> '' "
'    Call Tmp3Table
'
'    If TmpRs2.RecordCount > 0 And TmpRs3.RecordCount > 0 Then
'        For k = 1 To TmpRs2.RecordCount
'            TmpRs3.Filter = "ISINID = '" & TmpRs2!LotId & "'"
'            If TmpRs3.RecordCount > 0 Then
'               TmpRs3!CommtrackIncome = TmpRs2!Amount
'               TmpRs3.Update
'               TmpRs3.Filter = ""
''                tmp = " update txn_invmongsldetail1 set commtrackincome = isnull(InvoiceAmount,0) * " & TmpRs2!Amount & "/" & TmpRs3!gslamount
''                tmp = tmp & " where exchangetypeid = 2 and yearmonth between '" & mFromDate & "' and '" & mToDate & "'"
''                tmp = tmp & " and commtrackflag = 1 and isnull(isinid,'') <> ''"
''                tmp = tmp & " and isinid = '" & TmpRs3!ISINID & "'"
''                Call Tmp4Table
'            End If
'
'            TmpRs2.MoveNext
'        Next
'    End If
End If

MsgBox "Done!!!"
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

'Private Sub CmdProcess_Click()
'
''On Error GoTo msgError
'
'ProgressBar1.Value = 0
'
'If CmbProcessType.Text = "" Then
'   MsgBox " Blank Process Type not allowed  !!!"
'   CmbProcessType.SetFocus
'   Exit Sub
'End If
'
'If CmbMonth.Text = "" Then
'   MsgBox " Blank Month not allowed  !!!"
'   CmbMonth.SetFocus
'   Exit Sub
'End If
'
'If CmbYear.Text = "" Then
'   MsgBox " Blank Year not allowed  !!!"
'   CmbYear.SetFocus
'   Exit Sub
'End If
'
'If MSHFlexGrid1.Rows < 3 Then
'   MsgBox " No Location populated in Grid  !!!"
'   Exit Sub
'End If
'
'If ChkSelection = False Then
'   MsgBox " No Location selected from Grid  !!!"
'   Exit Sub
'End If
'
'mMonth = mMonthListindex + 1
'
'If mMonth < 10 Then
'   mMonth = "0" & mMonth
'End If
'
'
'mFromDate = Format(CmbYear.Text & "-" & mMonth & "-" & "01", "yyyy-mm-dd")
'mYear = Val(CmbYear.Text)
'mTotalDays = Day(DateSerial(mYear, mMonth + 1, 0))
'
'mToDate = Format(CmbYear.Text & "-" & mMonth & "-" & mTotalDays, "yyyy-mm-dd")
'mNoofdays = ""
'
'If CDate(Format(mFromDate, "yyyymm")) >= Format(Date, "yyyymm") Then
'   MsgBox "Present and Future Month not Allowed !!!"
'   CmbMonth.SetFocus
'   Exit Sub
'End If
'
'Dim ans As Variant
'
'If LCase(CmbProcessType.Text) <> "dp" Then
'   ans = MsgBox("All rates updated in franchisee rate master???", vbYesNo)
'Else
'   ans = MsgBox("All rates updated in NCDEX rate master???", vbYesNo)
'End If
'If ans = vbNo Then
'   Exit Sub
'End If
'
'CmdProcess.Enabled = False
'
'mGenInvoiceLocation = ""
'
'If MSHFlexGrid1.Rows > 2 Then
'      mGenInvoiceLocation = "SM_Prefix in ("
'      For I = 1 To MSHFlexGrid1.Rows - 1
'          If MSHFlexGrid1.TextMatrix(I, 1) = "" Then Exit For
'          If MSHFlexGrid1.TextMatrix(I, 0) = strChecked Then
'             If mGenInvoiceLocation = "SM_Prefix in (" Then
'                mGenInvoiceLocation = mGenInvoiceLocation & "'" & MSHFlexGrid1.TextMatrix(I, 2) & "'"
'             Else
'                mGenInvoiceLocation = mGenInvoiceLocation & ",'" & MSHFlexGrid1.TextMatrix(I, 2) & "' "
'             End If
'          End If
'      Next
'      mGenInvoiceLocation = mGenInvoiceLocation & " )"
'End If
'
'
'tmp = "Select Distinct SM_Prefix From Mst_Location Where " & mGenInvoiceLocation
'Call TmpTable
'
'ProgressBar1.Value = 0
'ProgressBar1.Max = TmpRs.RecordCount + 5
'
'For I = 1 To TmpRs.RecordCount
'
'    If LCase(CmbProcessType.Text) = "dp" Then
'    '---For Update KRNo
'
'        tmp = "Update Txn_InvMonGSLDetail Set KRNo=(Select KRNo from Mst_GSL a"
'        tmp = tmp & " where a.gslid = Txn_InvMonGSLDetail.gslid and a.sm_prefix = Txn_InvMonGSLDetail.sm_prefix)"
'        tmp = tmp & " Where ExchangeTypeID = 2  and yearmonth = '" & mFromDate & "' And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
'
'        Call Tmp2Table
'    ''
'''        tmp = " Update Txn_InvMonGSLDetail Set DRDate=(Select DRDate from Mst_DR a where DRNo=Txn_InvMonGSLDetail.KRNo),"
'''        tmp = tmp & " CRDate=(Select DRDate from Mst_DR a where DRNo=Txn_InvMonGSLDetail.KRNo),"
'''        tmp = tmp & " ISINID=(Select NewISINID from Mst_DR a where DRNo=Txn_InvMonGSLDetail.KRNo)"
'''        tmp = tmp & " Where ExchangeTypeID = 2  and yearmonth = '" & mFromDate & "' And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
'''
'        tmp = " Update Txn_InvMonGSLDetail Set DRDate=(Select DRDate from Mst_DR a where DRNo=Txn_InvMonGSLDetail.KRNo)"
'        tmp = tmp & " Where ExchangeTypeID = 2  and yearmonth = '" & mFromDate & "' And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
'
'        Call Tmp2Table
'
'        tmp = " Update Txn_InvMonGSLDetail Set ISINID=(Select NewISINID from Mst_DR a where DRNo=Txn_InvMonGSLDetail.KRNo),CRDate=DRDate"
'        tmp = tmp & " Where ExchangeTypeID = 2  and yearmonth = '" & mFromDate & "' And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
'
'        Call Tmp2Table
'
'        '--for Update rate
'
'        tmp = "Update Txn_InvMonGSLDetail Set Txn_InvMonGSLDetail.Rate ="
'        tmp = tmp & " (Select MNRC.Rate from Mst_NCDEXStorageRateCard MNR"
'        tmp = tmp & " Left Join Mst_NCDEXStorageRateCardDetail MNRC On MNR.NSRCID = MNRC.NSRCID And G.GodownID = MNRC.GodownID"
'        tmp = tmp & " Where MNR.LocationID=L.LocationID  And MNR.CommodityGroupID=C.CommodityGroupID)"
'        tmp = tmp & " from  Txn_InvMonGSLDetail  b"
'        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
'        tmp = tmp & " inner join mst_Godown g on b.godownId = g.godownid"
'        tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
'        tmp = tmp & " Where b.ExchangeTypeid = 2 And b.YearMonth = '" & mFromDate & "' And b.SM_Prefix = '" & TmpRs!SM_Prefix & "'"
'
'        Call Tmp2Table
'
'        tmp = "Update Txn_InvMonGSLDetail Set NoofDays = 30 Where  ExchangeTypeID = 2  And  YearMonth = '" & mFromDate & "' And DRDate < '" & mFromDate & "' And TFLag in ('O') And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
'        Call Tmp2Table
'
'        tmp = "Update Txn_InvMonGSLDetail Set NoofDays = (datediff(day,DRDate,'" & mToDate & "') + 1) Where  ExchangeTypeID = 2  And  YearMonth = '" & mFromDate & "' And DRDate >= '" & mFromDate & "' And DRDate <= '" & mToDate & "' And TFLag in ('O') And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
'        Call Tmp2Table
'
'        tmp = "Update Txn_InvMonGSLDetail Set NoofDays = (datediff(day,DRDate,'" & mToDate & "') + 1) Where  ExchangeTypeID = 2  And  YearMonth = '" & mFromDate & "' And DRDate >= '" & mFromDate & "' And DRDate <= '" & mToDate & "' And TFLag in ('D') And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
'        Call Tmp2Table
'
'        tmp = "Update Txn_InvMonGSLDetail Set NoofDays = (datediff(day,'" & mFromDate & "',DRDate) + 1) Where  ExchangeTypeID = 2  And  YearMonth = '" & mFromDate & "' And DRDate >= '" & mFromDate & "' And DRDate <= '" & mToDate & "' And TFLag in ('W') And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
'        Call Tmp2Table
'
'        tmp = "Update Txn_InvMonGSLDetail Set NoofDays = (datediff(day,'" & mFromDate & "',CxTFDate) + 1) Where  ExchangeTypeID = 2  And  YearMonth = '" & mFromDate & "' And DRDate < '" & mFromDate & "' And TFLag in ('W') And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
'        Call Tmp2Table
'
'        tmp = "Update Txn_InvMonGSLDetail  Set InvoiceAmount = Rate*NoOfDays*BalanceWeight Where  ExchangeTypeID = 2  And  YearMonth = '" & mFromDate & "' And Rate is not null And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
'        Call Tmp2Table
'
'    Else
'
'        tmp = "Update Txn_InvMonGSLDetail Set Txn_InvMonGSLDetail.FF_RateCardID="
'        tmp = tmp & " (Select distinct MNR.FF_RateCardID"
'        tmp = tmp & " from  Mst_FF_RateCard_Godown  b"
'        tmp = tmp & " inner Join Mst_FF_RateCard MNR On b.FF_RateCardID = MNR.FF_RateCardID"
'        tmp = tmp & " Where (MNR.CommodityGroupID = c.CommodityGroupID or b.BillingUnit='F') And (MNR.exchangetypeid = Txn_InvMonGSLDetail.exchangetypeid Or b.BillingUnit = 'F')"
'        tmp = tmp & " And b.GodownID=Txn_InvMonGSLDetail.GodownID)"
'        tmp = tmp & " From Txn_InvMonGSLDetail"
'        tmp = tmp & " inner join mst_commodity c on Txn_InvMonGSLDetail.commodityid = c.commodityid"
'        tmp = tmp & " Where Txn_InvMonGSLDetail.YearMonth = '" & mFromDate & "' And Txn_InvMonGSLDetail.SM_Prefix = '" & TmpRs!SM_Prefix & "'"
'
'        Call Tmp2Table
'
'        tmp = "Update Txn_InvMonGSLDetail Set Txn_InvMonGSLDetail.FranchBillingCycleID ="
'        tmp = tmp & " (Select b.BillingCycleID"
'        tmp = tmp & " from  Mst_FF_RateCard_Godown  b"
'        tmp = tmp & " inner Join Mst_FF_RateCard MNR On b.FF_RateCardID = MNR.FF_RateCardID"
'        tmp = tmp & " Where (MNR.CommodityGroupID = c.CommodityGroupID or b.BillingUnit='F' )And (MNR.exchangetypeid = Txn_InvMonGSLDetail.exchangetypeid Or b.BillingUnit = 'F')"
'        tmp = tmp & " And b.GodownID=Txn_InvMonGSLDetail.GodownID)"
'        tmp = tmp & " From Txn_InvMonGSLDetail"
'        tmp = tmp & " inner join mst_commodity c on Txn_InvMonGSLDetail.commodityid = c.commodityid"
'        tmp = tmp & " Where Txn_InvMonGSLDetail.YearMonth = '" & mFromDate & "' And Txn_InvMonGSLDetail.SM_Prefix = '" & TmpRs!SM_Prefix & "'"
'
'        Call Tmp2Table
'
'        tmp = " Update Txn_InvMonGSLDetail Set Txn_InvMonGSLDetail.FranchBillingUnit ="
'        tmp = tmp & " (Select b.BillingUnit"
'        tmp = tmp & " from  Mst_FF_RateCard_Godown  b"
'        tmp = tmp & " inner Join Mst_FF_RateCard MNR On b.FF_RateCardID = MNR.FF_RateCardID"
'        tmp = tmp & " Where (MNR.CommodityGroupID = c.CommodityGroupID or b.BillingUnit='F' )And (MNR.exchangetypeid = Txn_InvMonGSLDetail.exchangetypeid Or b.BillingUnit = 'F')"
'        tmp = tmp & " And b.GodownID=Txn_InvMonGSLDetail.GodownID)"
'        tmp = tmp & " From Txn_InvMonGSLDetail"
'        tmp = tmp & " inner join mst_commodity c on Txn_InvMonGSLDetail.commodityid = c.commodityid"
'        tmp = tmp & " Where Txn_InvMonGSLDetail.YearMonth = '" & mFromDate & "' And Txn_InvMonGSLDetail.SM_Prefix = '" & TmpRs!SM_Prefix & "'"
'
'        Call Tmp2Table
'
'        tmp = " Update Txn_InvMonGSLDetail Set Txn_InvMonGSLDetail.FranchiseeRate ="
'        tmp = tmp & " (Select b.RateOnQty"
'        tmp = tmp & " from  Mst_FF_RateCard_Godown  b"
'        tmp = tmp & " inner Join Mst_FF_RateCard MNR On b.FF_RateCardID = MNR.FF_RateCardID"
'        tmp = tmp & " Where (MNR.CommodityGroupID = c.CommodityGroupID or b.BillingUnit='F') And (MNR.exchangetypeid = Txn_InvMonGSLDetail.exchangetypeid Or b.BillingUnit = 'F')"
'        tmp = tmp & " And b.GodownID=Txn_InvMonGSLDetail.GodownID)"
'        tmp = tmp & " From Txn_InvMonGSLDetail"
'        tmp = tmp & " inner join mst_commodity c on Txn_InvMonGSLDetail.commodityid = c.commodityid"
'        tmp = tmp & " Where Txn_InvMonGSLDetail.YearMonth = '" & mFromDate & "' And Txn_InvMonGSLDetail.SM_Prefix = '" & TmpRs!SM_Prefix & "'  And Txn_InvMonGSLDetail.FranchBillingUnit = 'P'"
'
'        Call Tmp2Table
'
'        '------for Fix rate in date range
'
'        tmp = "Update Txn_InvMonGSLDetail Set FranchiseeRate = 0"
'        tmp = tmp & " Where Txn_InvMonGSLDetail.YearMonth = '" & mFromDate & "' And Txn_InvMonGSLDetail.SM_Prefix = '" & TmpRs!SM_Prefix & "' And Txn_InvMonGSLDetail.FranchBillingUnit = 'F'"
'        tmp = tmp & " And GodownID in (Select distinct b.GodownID"
'        tmp = tmp & " from  Mst_FF_RateCard_Godown_Fix  MNRCB"
'        tmp = tmp & " inner Join Mst_FF_RateCard_Godown b On b.FF_RateCard_GodownID = MNRCB.FF_RateCard_GodownID"
'        tmp = tmp & " inner Join Mst_FF_RateCard MNR On b.FF_RateCardID = MNR.FF_RateCardID"
'        tmp = tmp & " inner Join Mst_Location Ml On Ml.LocationID= MNR.LocationID"
'        tmp = tmp & " Where b.BillingUnit='F'"
'        tmp = tmp & " And MNRCB.StartDate<='" & mToDate & "'  And MNRCB.EndDate>='" & mFromDate & "' And ML.SM_Prefix = '" & TmpRs!SM_Prefix & "')"
'
'        Call Tmp2Table
'
'        tmp = "Select MNRCB.FixAmount,MNR.exchangetypeid,b.GodownID,MNRCB.StartDate,MNRCB.EndDate"
'        tmp = tmp & " from  Mst_FF_RateCard_Godown_Fix  MNRCB"
'        tmp = tmp & " inner Join Mst_FF_RateCard_Godown b On b.FF_RateCard_GodownID = MNRCB.FF_RateCard_GodownID"
'        tmp = tmp & " inner Join Mst_FF_RateCard MNR On b.FF_RateCardID = MNR.FF_RateCardID"
'        tmp = tmp & " inner Join Mst_Location Ml On Ml.LocationID= MNR.LocationID"
'        tmp = tmp & " Where b.BillingUnit='F'"
'        tmp = tmp & " And MNRCB.StartDate<='" & mToDate & "'  And MNRCB.EndDate>='" & mFromDate & "' And ML.SM_Prefix = '" & TmpRs!SM_Prefix & "' "
'
'        Call Tmp2Table
'
'        For rn = 1 To TmpRs2.RecordCount
'
'            If TmpRs2!StartDate <= CDate(mFromDate) And TmpRs2!EndDate >= CDate(mToDate) Then
'               mMonFixRate = TmpRs2!FixAmount
'            ElseIf TmpRs2!StartDate <= CDate(mFromDate) And TmpRs2!EndDate < CDate(mToDate) Then
'               mTotDays = DateDiff("d", CDate(mFromDate), TmpRs2!EndDate) + 1
'               mMonFixRate = (TmpRs2!FixAmount / 30) * mTotDays
'            ElseIf TmpRs2!StartDate > CDate(mFromDate) And TmpRs2!EndDate >= CDate(mToDate) Then
'               mTotDays = DateDiff("d", TmpRs2!StartDate, CDate(mToDate)) + 1
'               mMonFixRate = (TmpRs2!FixAmount / 30) * mTotDays
'            ElseIf TmpRs2!StartDate > CDate(mFromDate) And TmpRs2!EndDate < CDate(mToDate) Then
'               mTotDays = DateDiff("d", TmpRs2!StartDate, TmpRs2!EndDate) + 1
'               mMonFixRate = (TmpRs2!FixAmount / 30) * mTotDays
'            End If
'
'            tmp = "Update Txn_InvMonGSLDetail Set FranchiseeRate = isnull(FranchiseeRate,0) + " & mMonFixRate
'            tmp = tmp & " Where Txn_InvMonGSLDetail.YearMonth = '" & mFromDate & "' And Txn_InvMonGSLDetail.SM_Prefix = '" & TmpRs!SM_Prefix & "' And Txn_InvMonGSLDetail.FranchBillingUnit = 'F'"
'            tmp = tmp & " And GodownID=" & TmpRs2!GodownID & ""
'
'            Call Tmp3Table
'
'            TmpRs2.MoveNext
'        Next
'
'        '------------
'
'        tmp = " Update Txn_InvMonGSLDetail Set Txn_InvMonGSLDetail.FranchiseeRate ="
'        tmp = tmp & " (Select MNRCB.RatePerBag"
'        tmp = tmp & " from  Mst_FF_RateCard_Godown_Bag  MNRCB"
'        tmp = tmp & " inner Join Mst_FF_RateCard_Godown b On b.FF_RateCard_GodownID = MNRCB.FF_RateCard_GodownID"
'        tmp = tmp & " inner Join Mst_FF_RateCard MNR On b.FF_RateCardID = MNR.FF_RateCardID"
'        tmp = tmp & " Where (MNR.CommodityGroupID = c.CommodityGroupID or b.BillingUnit='F' )And (MNR.exchangetypeid = Txn_InvMonGSLDetail.exchangetypeid Or b.BillingUnit = 'F')"
'        tmp = tmp & " And b.GodownID=Txn_InvMonGSLDetail.GodownID"
'        tmp = tmp & " And Txn_InvMonGSLDetail.BagSize >= MNRCB.BagSizeFrom And Txn_InvMonGSLDetail.BagSize <= MNRCB.BagSizeTo  )"
'        tmp = tmp & " From Txn_InvMonGSLDetail"
'        tmp = tmp & " inner join mst_commodity c on Txn_InvMonGSLDetail.commodityid = c.commodityid"
'        tmp = tmp & " Where Txn_InvMonGSLDetail.YearMonth = '" & mFromDate & "' And Txn_InvMonGSLDetail.SM_Prefix = '" & TmpRs!SM_Prefix & "' And Txn_InvMonGSLDetail.FranchBillingUnit = 'B'"
'
'        Call Tmp2Table
'
'        'Op Bal and Monthly Cycle
'        tmp = " Update Txn_InvMonGSLDetail Set FranchNoofDays = 1 Where YearMonth = '" & mFromDate & "' And TFlag = 'O'"
'        tmp = tmp & " And FranchBillingCycleID = 4  And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
'
'        Call Tmp2Table
'
'        'Op Bal and Forthnightly Cycle
'        tmp = " Update Txn_InvMonGSLDetail Set FranchNoofDays = 2 Where YearMonth = '" & mFromDate & "' And TFlag = 'O'"
'        tmp = tmp & " And FranchBillingCycleID = 3 And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
'
'        Call Tmp2Table
'
'        'Op Bal and Weekly Cycle
'        tmp = " Update Txn_InvMonGSLDetail Set FranchNoofDays = 4 Where YearMonth = '" & mFromDate & "' And TFlag = 'O'"
'        tmp = tmp & " And FranchBillingCycleID = 2 And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
'
'        Call Tmp2Table
'
'        'ProgressBar1.Value = ProgressBar1.Value + 1
'
'        'Op Bal and Daily Cycle
'        tmp = " Update Txn_InvMonGSLDetail Set FranchNoofDays = " & mTotalDays & " Where YearMonth = '" & mFromDate & "' And TFlag = 'O'"
'        tmp = tmp & " And FranchBillingCycleID = 1 And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
'
'        Call Tmp2Table
'
'        'ProgressBar1.Value = ProgressBar1.Value + 1
'
'        'Depo and Withd and Monthly Cycle
'        tmp = " Update Txn_InvMonGSLDetail Set FranchNoofDays = 1 Where YearMonth = '" & mFromDate & "' And TFlag <> 'O'"
'        tmp = tmp & " And FranchBillingCycleID = 4 And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
'
'        Call Tmp2Table
'
'        'ProgressBar1.Value = ProgressBar1.Value + 1
'
'        'Depo and Withd and Daily Cycle
'        tmp = "Update Txn_InvMonGSLDetail Set FranchNoofDays = " & mTotalDays & " - Day(CxTFDate) + 1 Where YearMonth = '" & mFromDate & "' And TFlag = 'D' And FranchBillingCycleID = 1 And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
'
'        Call Tmp2Table
'
'        'ProgressBar1.Value = ProgressBar1.Value + 1
'
'        'Withd and Daily Cycle
'        tmp = "Update Txn_InvMonGSLDetail Set FranchNoOfDays = Day(CxTFDate) Where YearMonth = '" & mFromDate & "' And TFlag = 'W' And FranchBillingCycleID = 1 And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
'
'        Call Tmp2Table
'
'        'ProgressBar1.Value = ProgressBar1.Value + 1
'
'        'Deposit & Weekly
'
'        mNoofdays = GetBillingCycleDays(2)
'
'        tmp = "Update Txn_InvMonGSLDetail Set FranchNoOfDays =  Case Floor((4 - (Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ") + 1))  When 0 then 1 else Floor((4 - (Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ") + 1)) End "
'        tmp = tmp & " Where Txn_InvMonGSLDetail.YearMonth = '" & mFromDate & "'"
'        tmp = tmp & " And Txn_InvMonGSLDetail.TFlag = 'D' And Txn_InvMonGSLDetail.FranchBillingCycleID = 2 And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
'
'        Call Tmp2Table
'
'        'ProgressBar1.Value = ProgressBar1.Value + 1
'
'        'Deposit & FortNightly
'
'        mNoofdays = GetBillingCycleDays(3)
'
'        'Case 1 Depo Date  <= 15 Then No Of Cycle = 2
'
'        tmp = " Update Txn_InvMonGSLDetail Set FranchNoOfDays = 2 "
'        tmp = tmp & " Where Txn_InvMonGSLDetail.YearMonth = '" & mFromDate & "' "
'        tmp = tmp & " And Txn_InvMonGSLDetail.TFlag = 'D' And Txn_InvMonGSLDetail.FranchBillingCycleID = 3 And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
'        tmp = tmp & " And Day(CxTFDate) <= 15 "
'
'        Call Tmp2Table
'
'        'ProgressBar1.Value = ProgressBar1.Value + 1
'
'        'Case 1 Depo Date  > 15 Then No Of Cycle = 1
'        tmp = " Update Txn_InvMonGSLDetail Set FranchNoOfDays = 1 "
'        tmp = tmp & " Where YearMonth = '" & mFromDate & "' "
'        tmp = tmp & " And TFlag = 'D' And FranchBillingCycleID = 3 "
'        tmp = tmp & " And Day(CxTFDate) > 15 And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
'
'        Call Tmp2Table
'
'        'ProgressBar1.Value = ProgressBar1.Value + 1
'
'        'Withdrawal & weekly
'        mNoofdays = GetBillingCycleDays(2)
'
'        tmp = "Update Txn_InvMonGSLDetail Set FranchNoOfDays = Case CEILING(((Round(Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ",1)))) When 5 Then 4 Else CEILING(((Round(Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ",1)))) End "
'        tmp = tmp & " Where YearMonth = '" & mFromDate & "' "
'        tmp = tmp & " And TFlag = 'W' And FranchBillingCycleID = 2 And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
'
'        Call Tmp2Table
'
'        'ProgressBar1.Value = ProgressBar1.Value + 1
'
'        'Withdrawal & Fortnightly
'        mNoofdays = GetBillingCycleDays(3)
'
'        tmp = "Update Txn_InvMonGSLDetail Set FranchNoOfDays = Case CEILING(((Round(Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ",1)))) When 3 Then 2 Else CEILING(((Round(Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ",1)))) End "
'        tmp = tmp & " Where YearMonth = '" & mFromDate & "' "
'        tmp = tmp & " And TFlag = 'W' And FranchBillingCycleID = 3 And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
'
'        Call Tmp2Table
'
'        'ProgressBar1.Value = ProgressBar1.Value + 1
'
'        'Updating Invoice Amount And Invoice Date
'        tmp = " Update Txn_InvMonGSLDetail Set FranchInvoiceAmount = BalanceBags * FranchNoOfDays * FranchiseeRate, InvoiceDate = '" & mToDate & "' "
'        tmp = tmp & " Where YearMonth = '" & mFromDate & "'"
'        tmp = tmp & " And FranchBillingUnit = 'B' And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
'
'        Call Tmp2Table
'
'        'ProgressBar1.Value = ProgressBar1.Value + 1
'
'        tmp = " Update Txn_InvMonGSLDetail Set FranchInvoiceAmount = BalanceWeight * FranchNoOfDays * FranchiseeRate,InvoiceDate =  '" & mToDate & "' "
'        tmp = tmp & " Where YearMonth = '" & mFromDate & "'"
'        tmp = tmp & " And FranchBillingUnit = 'Q' And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
'
'        Call Tmp2Table
'
'    '--For Fixed
'
'        tmp = "Select GodownID,FranchiseeRate,max(GSLID)GSLID,FF_RateCardID from Txn_InvMonGSLDetail"
'        tmp = tmp & " Where YearMonth = '" & mFromDate & "' And FranchBillingUnit = 'F' And SM_Prefix = '" & TmpRs!SM_Prefix & "' And GSLID>0  And FranchiseeRate>0"
'        tmp = tmp & " Group by GodownID,FranchiseeRate,FF_RateCardID"
'
'        Call Tmp2Table
'
'        For rn = 1 To TmpRs2.RecordCount
'
'            tmp = "Update Txn_InvMonGSLDetail Set FranchInvoiceAmount = " & TmpRs2!FranchiseeRate & ",InvoiceDate =  '" & mToDate & "'"
'            tmp = tmp & " Where YearMonth = '" & mFromDate & "' And GSLID=" & TmpRs2!GSLID & " And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
'
'            Call Tmp4Table
'
'            tmp = "Select GSLID,FranchInvoiceAmount,BalanceBags,MonG_UID from  Txn_InvMonGSLDetail "
'            tmp = tmp & " Where YearMonth = '" & mFromDate & "' And GSLID=" & TmpRs2!GSLID & " And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
'
'            Call Tmp3Table
'
'            If TmpRs3.RecordCount > 1 Then
'               For rn1 = 1 To TmpRs3.RecordCount - 1
'
'                   TmpRs3!FranchInvoiceAmount = 0
'                   TmpRs3.Update
'
'                   TmpRs3.MoveNext
'               Next
'            End If
'
'            tmp = "Update Txn_InvMonGSLDetail Set FF_RateCardID=" & TmpRs2!FF_RateCardID & ",FranchBillingUnit='F',FranchNoofDays=1,FranchiseeRate=" & TmpRs2!FranchiseeRate & ",FranchInvoiceAmount=0,FranchBillingCycleID=4"
'            tmp = tmp & " Where YearMonth = '" & mFromDate & "' And SM_Prefix = '" & TmpRs!SM_Prefix & "' And GodownID=" & TmpRs2!GodownID & " And GSLID<>" & TmpRs2!GSLID & ""
'
'            Call Tmp4Table
'
'            TmpRs2.MoveNext
'        Next
'
'        '-----For Percent
'
'        tmp = " Update Txn_InvMonGSLDetail Set FranchInvoiceAmount =  (InvoiceAmount)* FranchiseeRate /100 ,InvoiceDate =  '" & mToDate & "' "
'        tmp = tmp & " Where YearMonth = '" & mFromDate & "'"
'        tmp = tmp & " And FranchBillingUnit = 'P' And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
'
'        Call Tmp2Table
'
'    End If
'
'    TmpRs.MoveNext
'    ProgressBar1.Value = ProgressBar1.Value + 1
'
'Next
'
'If LCase(CmbProcessType.Text) <> "dp" Then
'
'   tmp = "Update Txn_InvMonGSLDetail set FranchProcessFlag='Y',FranchProcessdBy='" & Gen_UserLoginID & "',FranchProcessedDate=getdate() Where YearMonth = '" & mFromDate & "' And " & mGenInvoiceLocation
'   Call TmpTable
'
'End If
'
'ProgressBar1.Value = ProgressBar1.Max
'MsgBox "Done!!!"
'
'Call Grid_Head
'
'Exit Sub
'
'msgError:
'MsgBox "Some error occurred during process!!!"
'
'End Sub

Private Sub CmdProcess_Click()

'On Error GoTo msgError

ProgressBar1.Value = 0

If CmbProcessType.Text = "" Then
   MsgBox " Blank Process Type not allowed  !!!"
   CmbProcessType.SetFocus
   Exit Sub
End If

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

If MSHFlexGrid1.Rows < 3 Then
   MsgBox " No Location populated in Grid  !!!"
   Exit Sub
End If

If ChkSelection = False Then
   MsgBox " No Location selected from Grid  !!!"
   Exit Sub
End If

mMonth = mMonthListindex + 1

If mMonth < 10 Then
   mMonth = "0" & mMonth
End If

mFromDate = Format(CmbYear.Text & "-" & mMonth & "-" & "01", "yyyy-mm-dd")
mYear = Val(CmbYear.Text)
mTotalDays = Day(DateSerial(mYear, mMonth + 1, 0))

mToDate = Format(CmbYear.Text & "-" & mMonth & "-" & mTotalDays, "yyyy-mm-dd")
mNoofDays = ""

If CDate(Format(mFromDate, "yyyymm")) >= Format(Date, "yyyymm") Then
   MsgBox "Present and Future Month not Allowed !!!"
   CmbMonth.SetFocus
   Exit Sub
End If


'If Format(Gen_BlockingDate, "yyyy-mm") >= Format(mFromDate, "yyyy-mm") Then
'   MsgBox "Select next month from " & Format(Gen_BlockingDate, "MMM-yyyy") & "!!! "
'   CmbMonth.SetFocus
'   Exit Sub
'End If


Dim ans As Variant

If LCase(CmbProcessType.Text) <> "dp" Then
   ans = MsgBox("All rates updated in franchisee rate master???", vbYesNo)
Else
   ans = MsgBox("All rates updated in NCDEX rate master???", vbYesNo)
End If
If ans = vbNo Then
   Exit Sub
End If

CmdProcess.Enabled = False

mGenInvoiceLocation = ""

If MSHFlexGrid1.Rows > 2 Then
      mGenInvoiceLocation = "SM_Prefix in ("
      For I = 1 To MSHFlexGrid1.Rows - 1
          If MSHFlexGrid1.TextMatrix(I, 1) = "" Then Exit For
          If MSHFlexGrid1.TextMatrix(I, 0) = strChecked Then
             If mGenInvoiceLocation = "SM_Prefix in (" Then
                mGenInvoiceLocation = mGenInvoiceLocation & "'" & MSHFlexGrid1.TextMatrix(I, 2) & "'"
             Else
                mGenInvoiceLocation = mGenInvoiceLocation & ",'" & MSHFlexGrid1.TextMatrix(I, 2) & "' "
             End If
          End If
      Next
      mGenInvoiceLocation = mGenInvoiceLocation & " )"
End If


tmp = "Select Distinct SM_Prefix, LocationID From Mst_Location Where " & mGenInvoiceLocation
Call TmpTable

ProgressBar1.Value = 0
ProgressBar1.Max = TmpRs.RecordCount + 5

For I = 1 To TmpRs.RecordCount

        
    If LCase(CmbProcessType.Text) = "dp" Then
    '---For Update KRNo
    
        tmp = "SET IMPLICIT_TRANSACTIONS ON"
        Call Tmp2Table
    
    
        tmp = "Update Txn_InvMonGSLDetail Set KRNo=(Select KRNo from Mst_GSL a"
        tmp = tmp & " where a.Mst_GSL_ID = Txn_InvMonGSLDetail.Mst_GSL_ID and a.sm_prefix = Txn_InvMonGSLDetail.sm_prefix)"
        tmp = tmp & " Where ExchangeTypeID = 2  and yearmonth = '" & mFromDate & "' And SM_Prefix = '" & TmpRs!SM_Prefix & "' And KRNo is null "
    
        Call Tmp2Table
    ''
''        tmp = " Update Txn_InvMonGSLDetail Set DRDate=(Select DRDate from Mst_DR a where DRNo=Txn_InvMonGSLDetail.KRNo),"
''        tmp = tmp & " CRDate=(Select DRDate from Mst_DR a where DRNo=Txn_InvMonGSLDetail.KRNo),"
''        tmp = tmp & " ISINID=(Select NewISINID from Mst_DR a where DRNo=Txn_InvMonGSLDetail.KRNo)"
''        tmp = tmp & " Where ExchangeTypeID = 2  and yearmonth = '" & mFromDate & "' And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
''
        tmp = "SET IMPLICIT_TRANSACTIONS ON"
        Call Tmp2Table
       
       tmp = " Update Txn_InvMonGSLDetail Set ISINID=(Select ISINID from Mst_GSL a"
       tmp = tmp & " where a.Mst_GSL_ID = Txn_InvMonGSLDetail.Mst_GSL_ID and a.sm_prefix = Txn_InvMonGSLDetail.sm_prefix)"
       tmp = tmp & " Where ExchangeTypeID = 2  and Txn_InvMonGSLDetail.CommtrackFlag = '1' and yearmonth = '" & mFromDate & "' And SM_Prefix = '" & TmpRs!SM_Prefix & "' And ISINID Is null"
       
       Call Tmp2Table

        tmp = "SET IMPLICIT_TRANSACTIONS ON"
        Call Tmp2Table

       '''Introduced on 10 Feb 2014 for Implementation of new calculation process required by CMG
       tmp = "Update Txn_InvMonGSLDetail Set CommtrackFlag =(Select Distinct CommtrackFlag from Mst_DR a where NewISINID=Txn_InvMonGSLDetail.ISINID ) Where ExchangeTypeID = 2  and yearmonth = '" & mFromDate & "' And SM_Prefix = '" & TmpRs!SM_Prefix & "' And CommtrackFlag is null "
       Call Tmp2Table
       
       
        tmp = "SET IMPLICIT_TRANSACTIONS ON"
        Call Tmp2Table
       
       tmp = " Update Txn_InvMonGSLDetail Set CommtrackFlag = 0  Where ISINID like '%INC%'  And YearMonth = '" & mFromDate & "' And ExchangeTypeID = 2 And CommtrackFlag is null And SM_Prefix = '" & TmpRs!SM_Prefix & "' "
       Call Tmp2Table
       
        tmp = "SET IMPLICIT_TRANSACTIONS ON"
        Call Tmp2Table
       
       tmp = " Update Txn_InvMonGSLDetail Set DRDate=(Select Max(DRDate) 'DRDate' from Mst_DR a where NewISINID=Txn_InvMonGSLDetail.ISINID)"
       tmp = tmp & " Where ExchangeTypeID = 2  and yearmonth = '" & mFromDate & "' And SM_Prefix = '" & TmpRs!SM_Prefix & "' And Txn_InvMonGSLDetail.CommtrackFlag = '1' And DRDate is null"

       Call Tmp2Table

        tmp = "SET IMPLICIT_TRANSACTIONS ON"
        Call Tmp2Table
        
       tmp = " Update Txn_InvMonGSLDetail Set RRNDate=(Select b.RRNDate from Mst_RRN_LotDetail a inner join Mst_RRN b on a.RRNID = b.RRNID Where a.ISINID=Txn_InvMonGSLDetail.ISINID)"
       tmp = tmp & " Where ExchangeTypeID = 2  and yearmonth = '" & mFromDate & "' And SM_Prefix = '" & TmpRs!SM_Prefix & "' And Txn_InvMonGSLDetail.CommtrackFlag = '1' And RRNDate is null "

       Call Tmp2Table

       'Commented on 10 Feb 2014 for Implementation of new calculation process required by CMG
       'tmp = " Update Txn_InvMonGSLDetail Set DRDate= KRDate "
       'tmp = tmp & " Where ExchangeTypeID = 2  And Txn_InvMonGSLDetail.CommtrackFlag = '1' and yearmonth = '" & mFromDate & "' And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
       
      ' Call Tmp2Table
        tmp = "SET IMPLICIT_TRANSACTIONS ON"
        Call Tmp2Table
        
       tmp = " Update Txn_InvMonGSLDetail Set DRDate=(Select Max(DRDate) from Mst_DR a where DRNo=Txn_InvMonGSLDetail.KRNo)"
       tmp = tmp & " Where ExchangeTypeID = 2  and yearmonth = '" & mFromDate & "' And SM_Prefix = '" & TmpRs!SM_Prefix & "' And Txn_InvMonGSLDetail.CommtrackFlag = '0' And DRDate is null "

       Call Tmp2Table
    
        tmp = "SET IMPLICIT_TRANSACTIONS ON"
        Call Tmp2Table
    
       tmp = " Update Txn_InvMonGSLDetail Set ISINID=(Select NewISINID from Mst_DR a where DRNo=Txn_InvMonGSLDetail.KRNo),CRDate=DRDate"
       tmp = tmp & " Where ExchangeTypeID = 2  and yearmonth = '" & mFromDate & "' And SM_Prefix = '" & TmpRs!SM_Prefix & "' And Txn_InvMonGSLDetail.CommtrackFlag = '0' And ISINID is null "

       Call Tmp2Table
        
        '--for Update rate
        
        tmp = "SET IMPLICIT_TRANSACTIONS ON"
        Call Tmp2Table
        
       tmp = "Update Txn_InvMonGSLDetail Set Txn_InvMonGSLDetail.Rate ="
       tmp = tmp & " (Select MNRC.Rate from Mst_NCDEXStorageRateCard MNR"
       tmp = tmp & " Left Join Mst_NCDEXStorageRateCardDetail MNRC On MNR.NSRCID = MNRC.NSRCID And G.GodownID = MNRC.GodownID"
       tmp = tmp & " Where MNR.LocationID=L.LocationID  And MNR.CommodityGroupID=C.CommodityGroupID)"
       tmp = tmp & " from  Txn_InvMonGSLDetail  b"
       tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
       tmp = tmp & " inner join mst_Godown g on b.godownId = g.godownid"
       tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
       tmp = tmp & " Where b.ExchangeTypeid = 2 And b.YearMonth = '" & mFromDate & "' And b.SM_Prefix = '" & TmpRs!SM_Prefix & "' And b.Rate is null "

       Call Tmp2Table
    
        tmp = "SET IMPLICIT_TRANSACTIONS ON"
        Call Tmp2Table
    
       tmp = "Update Txn_InvMonGSLDetail Set NoofDays = 30 Where  ExchangeTypeID = 2  And  YearMonth = '" & mFromDate & "' And DRDate < '" & mFromDate & "' And TFLag in ('O') And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
       Call Tmp2Table
    
        tmp = "SET IMPLICIT_TRANSACTIONS ON"
        Call Tmp2Table
    
       tmp = "Update Txn_InvMonGSLDetail Set NoofDays = (datediff(day,DRDate,'" & mToDate & "') + 1) Where  ExchangeTypeID = 2  And  YearMonth = '" & mFromDate & "' And DRDate >= '" & mFromDate & "' And DRDate <= '" & mToDate & "' And TFLag in ('O') And SM_Prefix = '" & TmpRs!SM_Prefix & "' "
       Call Tmp2Table
    
        tmp = "SET IMPLICIT_TRANSACTIONS ON"
        Call Tmp2Table
    
       tmp = "Update Txn_InvMonGSLDetail Set NoofDays = (datediff(day,DRDate,'" & mToDate & "') + 1) Where  ExchangeTypeID = 2  And  YearMonth = '" & mFromDate & "' And DRDate >= '" & mFromDate & "' And DRDate <= '" & mToDate & "' And TFLag in ('D') And SM_Prefix = '" & TmpRs!SM_Prefix & "' "
       Call Tmp2Table
            
        tmp = "SET IMPLICIT_TRANSACTIONS ON"
        Call Tmp2Table
            
       tmp = "Update Txn_InvMonGSLDetail Set NoofDays = (datediff(day,'" & mFromDate & "',DRDate) + 1) Where  ExchangeTypeID = 2  And  YearMonth = '" & mFromDate & "' And DRDate >= '" & mFromDate & "' And DRDate <= '" & mToDate & "' And TFLag in ('W') And SM_Prefix = '" & TmpRs!SM_Prefix & "' And CommtrackFlag = '0'"
       Call Tmp2Table
       
        tmp = "SET IMPLICIT_TRANSACTIONS ON"
        Call Tmp2Table
       
       tmp = "Update Txn_InvMonGSLDetail Set NoofDays = (datediff(day,'" & mFromDate & "',CxTFDate) + 1) Where  ExchangeTypeID = 2  And  YearMonth = '" & mFromDate & "' And DRDate < '" & mFromDate & "' And TFLag in ('W') And SM_Prefix = '" & TmpRs!SM_Prefix & "' And CommtrackFlag = '0'"
       Call Tmp2Table
    
        tmp = "SET IMPLICIT_TRANSACTIONS ON"
        Call Tmp2Table
       '''Introduced on 10 Feb 2014 for Implementation of new calculation process required by CMG
    
       tmp = "Update Txn_InvMonGSLDetail Set NoofDays = (datediff(day,DRDate,RRNDate) + 1) Where  ExchangeTypeID = 2  And  YearMonth = '" & mFromDate & "' And DRDate >= '" & mFromDate & "' And DRDate <= '" & mToDate & "' And TFLag in ('W') And SM_Prefix = '" & TmpRs!SM_Prefix & "' And CommtrackFlag = '1'  And NoofDays is null "
       Call Tmp2Table
       
        tmp = "SET IMPLICIT_TRANSACTIONS ON"
        Call Tmp2Table
       
       tmp = "Update Txn_InvMonGSLDetail Set NoofDays = (datediff(day,'" & mFromDate & "',RRNDate) + 1) Where  ExchangeTypeID = 2  And  YearMonth = '" & mFromDate & "' And DRDate < '" & mFromDate & "' And TFLag in ('W') And SM_Prefix = '" & TmpRs!SM_Prefix & "' And CommtrackFlag = '1' And RRNDate > '" & mFromDate & "' And NoofDays is null "
       Call Tmp2Table
    
        tmp = "SET IMPLICIT_TRANSACTIONS ON"
        Call Tmp2Table
    
       tmp = "Update Txn_InvMonGSLDetail  Set InvoiceAmount = Rate*NoOfDays*BalanceWeight Where  ExchangeTypeID = 2  And  YearMonth = '" & mFromDate & "' And Rate is not null And SM_Prefix = '" & TmpRs!SM_Prefix & "' And isnull(InvoiceAmount,0) = 0  "
       Call Tmp2Table
    
        tmp = "SET IMPLICIT_TRANSACTIONS ON"
        Call Tmp2Table
    
       tmp = "Update Txn_InvMonGSLDetail  Set InvoiceAmount = 0 Where  ExchangeTypeID = 2  And  YearMonth = '" & mFromDate & "' And Rate is not null And SM_Prefix = '" & TmpRs!SM_Prefix & "' And (Flag = 'J' Or Flag1 = 'J')"
       Call Tmp2Table
       
        tmp = "SET IMPLICIT_TRANSACTIONS ON"
        Call Tmp2Table
       
       tmp = "Update Txn_InvMonGSLDetail  Set InvoiceAmount = 0 Where  ExchangeTypeID = 2  And  YearMonth = '" & mFromDate & "' And Rate is not null And SM_Prefix = '" & TmpRs!SM_Prefix & "' And DRDate > '" & mToDate & "' "
       Call Tmp2Table
       
        tmp = "SET IMPLICIT_TRANSACTIONS ON"
        Call Tmp2Table
    
       tmp = " Update TIG set dpincome = isnull(tig.InvoiceAmount,0)* (select isnull(sum(MDP.Amt),0) from Mst_DPInvoiceUpload MDP where MDP.LocationID  = " & TmpRs!LocationID & " and "
       tmp = tmp & " monthyear between '" & mFromDate & "' and '" & mToDate & "' and isnull(MDP.ReverseFlag,'')<>'Y')/(select isnull(sum(InvoiceAmount),0) from txn_invmongsldetail where exchangetypeid = 2 and yearmonth between '" & mFromDate & "' and '" & mToDate & "' and sm_prefix = '" & TmpRs!SM_Prefix & "' and isnull(commtrackflag,0) = 0)"
       tmp = tmp & " from txn_invmongsldetail  tig where yearmonth between '" & mFromDate & "' and '" & mToDate & "' and isnull(tig.InvoiceAmount,0) > 0"
       tmp = tmp & " and exchangetypeid = 2 and sm_prefix = '" & TmpRs!SM_Prefix & "'"
       tmp = tmp & " and isnull(commtrackflag,0) = 0"
       Call Tmp2Table
    Else
     
     '''''For DP Process, Added on 15/03/2013 by Rakesh=========================================
        '---For Update KRNo
   
     '''''For DP Process, Added on 15/03/2013 by Rakesh=========================================
        '---For Update KRNo
   
        tmp = "SET IMPLICIT_TRANSACTIONS ON"
        Call Tmp2Table
   
       tmp = "UPDATE Txn_InvMonGSLDetail SET FranchBillingCycleID = Null,FranchBillingUnit = Null,FranchNoofDays = Null,FranchInvoiceAmount = Null,FF_RateCardID = Null,FranchClaimID = Null,FranchProcessFlag = 'N',FranchProcessdBy = Null,FranchProcessedDate = Null,InvoiceAmtForFranch = 0,InsuranceCostDeduction=0 WHere YearMonth = '" & mFromDate & "' And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
       Call Tmp2Table
       
       
       Call GetNNNRIncome(TmpRs!SM_Prefix)
       Call GetReservationIncome(TmpRs!SM_Prefix)
       Call GetNCDEXIncome(TmpRs!SM_Prefix)
   
'       tmp = "Update Txn_InvMonGSLDetail Set KRNo=(Select KRNo from Mst_GSL a"
'       tmp = tmp & " where a.Mst_GSL_ID = Txn_InvMonGSLDetail.Mst_GSL_ID and a.sm_prefix = Txn_InvMonGSLDetail.sm_prefix)"
'       tmp = tmp & " Where ExchangeTypeID = 2  and yearmonth = '" & mFromDate & "' And SM_Prefix = '" & TmpRs!SM_Prefix & "' And KRNo is null"
'
'       Call Tmp2Table
'
'        tmp = "SET IMPLICIT_TRANSACTIONS ON"
'        Call Tmp2Table
'
'       tmp = " Update Txn_InvMonGSLDetail Set DRDate=(Select DRDate from Mst_DR a where DRNo=Txn_InvMonGSLDetail.KRNo)"
'       tmp = tmp & " Where ExchangeTypeID = 2  and yearmonth = '" & mFromDate & "' And SM_Prefix = '" & TmpRs!SM_Prefix & "' And Txn_InvMonGSLDetail.CommtrackFlag = '0' And DRDate  is null "
'
'       Call Tmp2Table
'
'       tmp = "SET IMPLICIT_TRANSACTIONS ON"
'       Call Tmp2Table
'
'       tmp = " Update Txn_InvMonGSLDetail Set ISINID=(Select NewISINID from Mst_DR a where DRNo=Txn_InvMonGSLDetail.KRNo),CRDate=DRDate"
'       tmp = tmp & " Where ExchangeTypeID = 2  and yearmonth = '" & mFromDate & "' And SM_Prefix = '" & TmpRs!SM_Prefix & "' And ISINID is null "
'
'       Call Tmp2Table
'
'
'        tmp = "SET IMPLICIT_TRANSACTIONS ON"
'        Call Tmp2Table
'
'       tmp = " Update Txn_InvMonGSLDetail Set ISINID=(Select ISINID from Mst_GSL a"
'       tmp = tmp & " where a.Mst_GSL_ID = Txn_InvMonGSLDetail.Mst_GSL_ID and a.sm_prefix = Txn_InvMonGSLDetail.sm_prefix)"
'       tmp = tmp & " Where ExchangeTypeID = 2  and Txn_InvMonGSLDetail.CommtrackFlag = '1' and yearmonth = '" & mFromDate & "' And SM_Prefix = '" & TmpRs!SM_Prefix & "' And ISINID is null"
'
'       Call Tmp2Table
'
''       tmp = " Update Txn_InvMonGSLDetail Set DRDate=(Select DRDate from Mst_DR a where Convert(Varchar,DRNo)=Txn_InvMonGSLDetail.ISINID)"
''       tmp = tmp & " Where ExchangeTypeID = 2  And Txn_InvMonGSLDetail.CommtrackFlag = '1' and yearmonth = '" & mFromDate & "' And SM_Prefix = '" & TmpRs!sm_prefix & "'"
'
'
'       '''Introduced on 10 Feb 2014 for Implementation of new calculation process required by CMG
'
'        tmp = "SET IMPLICIT_TRANSACTIONS ON"
'        Call Tmp2Table
'
'       tmp = "Update Txn_InvMonGSLDetail Set CommtrackFlag =(Select Distinct CommtrackFlag from Mst_DR a where NewISINID=Txn_InvMonGSLDetail.ISINID ) Where ExchangeTypeID = 2  and yearmonth = '" & mFromDate & "' And SM_Prefix = '" & TmpRs!SM_Prefix & "' And CommtrackFlag is null "
'       Call Tmp2Table
'
'        tmp = "SET IMPLICIT_TRANSACTIONS ON"
'        Call Tmp2Table
'
'       tmp = " Update Txn_InvMonGSLDetail Set CommtrackFlag = 0  Where ISINID like '%INC%'  And YearMonth = '" & mFromDate & "' And ExchangeTypeID = 2 And CommtrackFlag is null And SM_Prefix = '" & TmpRs!SM_Prefix & "' "
'       Call Tmp2Table
'
'        tmp = "SET IMPLICIT_TRANSACTIONS ON"
'        Call Tmp2Table
'
'       tmp = " Update Txn_InvMonGSLDetail Set DRDate=(Select Max(DRDate) from Mst_DR a where NewISINID=Txn_InvMonGSLDetail.ISINID)"
'       tmp = tmp & " Where ExchangeTypeID = 2  and yearmonth = '" & mFromDate & "' And SM_Prefix = '" & TmpRs!SM_Prefix & "' And Txn_InvMonGSLDetail.CommtrackFlag = '1' And DRDate is null"
'
'       Call Tmp2Table
'
'        tmp = "SET IMPLICIT_TRANSACTIONS ON"
'        Call Tmp2Table
'
'       tmp = " Update Txn_InvMonGSLDetail Set RRNDate=(Select b.RRNDate from Mst_RRN_LotDetail a inner join Mst_RRN b on a.RRNID = b.RRNID Where a.ISINID=Txn_InvMonGSLDetail.ISINID)"
'       tmp = tmp & " Where ExchangeTypeID = 2  and yearmonth = '" & mFromDate & "' And SM_Prefix = '" & TmpRs!SM_Prefix & "' And Txn_InvMonGSLDetail.CommtrackFlag = '1' And RRNDate is null"
'
'       Call Tmp2Table

       'Commented on 10 Feb 2014 for Implementation of new calculation process required by CMG
       
       'tmp = " Update Txn_InvMonGSLDetail Set DRDate= KRDate "
       'tmp = tmp & " Where ExchangeTypeID = 2  And Txn_InvMonGSLDetail.CommtrackFlag = '1' and yearmonth = '" & mFromDate & "' And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
       
       'Call Tmp2Table
        
        
        '--for Update rate
        
        tmp = "SET IMPLICIT_TRANSACTIONS ON"
        Call Tmp2Table
        
        tmp = "Update Txn_InvMonGSLDetail Set Txn_InvMonGSLDetail.Rate ="
        tmp = tmp & " (Select MNRC.Rate from Mst_NCDEXStorageRateCard MNR"
        tmp = tmp & " Left Join Mst_NCDEXStorageRateCardDetail MNRC On MNR.NSRCID = MNRC.NSRCID And G.GodownID = MNRC.GodownID"
        tmp = tmp & " Where MNR.LocationID=L.LocationID  And MNR.CommodityGroupID=C.CommodityGroupID)"
        tmp = tmp & " from  Txn_InvMonGSLDetail  b"
        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
        tmp = tmp & " inner join mst_Godown g on b.godownId = g.godownid"
        tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
        tmp = tmp & " Where b.ExchangeTypeid = 2 And b.YearMonth = '" & mFromDate & "' And b.SM_Prefix = '" & TmpRs!SM_Prefix & "' And b.Rate is null"
   
        Call Tmp2Table
   
        tmp = "SET IMPLICIT_TRANSACTIONS ON"
        Call Tmp2Table
   
        tmp = "Update Txn_InvMonGSLDetail Set NoofDays = 30 Where  ExchangeTypeID = 2  And  YearMonth = '" & mFromDate & "' And DRDate < '" & mFromDate & "' And TFLag in ('O') And SM_Prefix = '" & TmpRs!SM_Prefix & "' And NoofDays is null "
        Call Tmp2Table
        
        tmp = "SET IMPLICIT_TRANSACTIONS ON"
        Call Tmp2Table
   
        tmp = "Update Txn_InvMonGSLDetail Set NoofDays = (datediff(day,DRDate,'" & mToDate & "') + 1) Where  ExchangeTypeID = 2  And  YearMonth = '" & mFromDate & "' And DRDate >= '" & mFromDate & "' And DRDate <= '" & mToDate & "' And TFLag in ('O') And SM_Prefix = '" & TmpRs!SM_Prefix & "' And NoofDays is null "
        Call Tmp2Table
   
        tmp = "SET IMPLICIT_TRANSACTIONS ON"
        Call Tmp2Table
   
        tmp = "Update Txn_InvMonGSLDetail Set NoofDays = (datediff(day,DRDate,'" & mToDate & "') + 1) Where  ExchangeTypeID = 2  And  YearMonth = '" & mFromDate & "' And DRDate >= '" & mFromDate & "' And DRDate <= '" & mToDate & "' And TFLag in ('D') And SM_Prefix = '" & TmpRs!SM_Prefix & "' And NoofDays is null "
        Call Tmp2Table
   
        tmp = "SET IMPLICIT_TRANSACTIONS ON"
        Call Tmp2Table
   
        tmp = "Update Txn_InvMonGSLDetail Set NoofDays = (datediff(day,'" & mFromDate & "',DRDate) + 1) Where  ExchangeTypeID = 2  And  YearMonth = '" & mFromDate & "' And DRDate >= '" & mFromDate & "' And DRDate <= '" & mToDate & "' And TFLag in ('W') And SM_Prefix = '" & TmpRs!SM_Prefix & "' And CommtrackFlag = '0' And NoofDays is null "
        Call Tmp2Table
   
        tmp = "SET IMPLICIT_TRANSACTIONS ON"
        Call Tmp2Table
   
        tmp = "Update Txn_InvMonGSLDetail Set NoofDays = (datediff(day,'" & mFromDate & "',CxTFDate) + 1) Where  ExchangeTypeID = 2  And  YearMonth = '" & mFromDate & "' And DRDate < '" & mFromDate & "' And TFLag in ('W') And SM_Prefix = '" & TmpRs!SM_Prefix & "' And CommtrackFlag = '0' And NoofDays is null "
        Call Tmp2Table
   
   
        tmp = "SET IMPLICIT_TRANSACTIONS ON"
        Call Tmp2Table
        
'------Introduced on 10 Feb 2014 for Implementation of new calculation process required by CMG
       tmp = "Update Txn_InvMonGSLDetail Set NoofDays = (datediff(day,DRDate,RRNDate) + 1) Where  ExchangeTypeID = 2  And  YearMonth = '" & mFromDate & "' And DRDate >= '" & mFromDate & "' And DRDate <= '" & mToDate & "' And TFLag in ('W') And SM_Prefix = '" & TmpRs!SM_Prefix & "' And CommtrackFlag = '1' And NoofDays is null  "
       Call Tmp2Table
       
        tmp = "SET IMPLICIT_TRANSACTIONS ON"
        Call Tmp2Table
       
       tmp = "Update Txn_InvMonGSLDetail Set NoofDays = (datediff(day,'" & mFromDate & "',RRNDate) + 1) Where  ExchangeTypeID = 2  And  YearMonth = '" & mFromDate & "' And DRDate < '" & mFromDate & "' And TFLag in ('W') And SM_Prefix = '" & TmpRs!SM_Prefix & "' And CommtrackFlag = '1' And RRNDate > '" & mFromDate & "' And NoofDays is null  "
       Call Tmp2Table
   
        tmp = "SET IMPLICIT_TRANSACTIONS ON"
        Call Tmp2Table
   
        tmp = "Update Txn_InvMonGSLDetail  Set InvoiceAmount = Rate*NoOfDays*BalanceWeight Where  ExchangeTypeID = 2  And  YearMonth = '" & mFromDate & "' And Rate is not null And SM_Prefix = '" & TmpRs!SM_Prefix & "' And isnull(InvoiceAmount,0) = 0 "
        Call Tmp2Table
   
        tmp = "SET IMPLICIT_TRANSACTIONS ON"
        Call Tmp2Table
   
       tmp = "Update Txn_InvMonGSLDetail  Set InvoiceAmount = 0 Where  ExchangeTypeID = 2  And  YearMonth = '" & mFromDate & "' And Rate is not null And SM_Prefix = '" & TmpRs!SM_Prefix & "' And (Flag = 'J' Or Flag1 = 'J')"
       Call Tmp2Table
   
       tmp = "SET IMPLICIT_TRANSACTIONS ON"
        Call Tmp2Table
        
       tmp = "Update Txn_InvMonGSLDetail  Set InvoiceAmount = 0 Where  ExchangeTypeID = 2  And  YearMonth = '" & mFromDate & "' And Rate is not null And SM_Prefix = '" & TmpRs!SM_Prefix & "' And DRDate > '" & mToDate & "' "
       Call Tmp2Table
   
        tmp = "SET IMPLICIT_TRANSACTIONS ON"
        Call Tmp2Table
        
        tmp = " Update TIG set dpincome = isnull(tig.InvoiceAmount,0)* (select isnull(sum(MDP.Amt),0) from Mst_DPInvoiceUpload MDP where MDP.LocationID  = " & TmpRs!LocationID & " and "
        tmp = tmp & " monthyear between '" & mFromDate & "' and '" & mToDate & "' and isnull(MDP.ReverseFlag,'')<>'Y')/(select isnull(sum(InvoiceAmount),0) from txn_invmongsldetail where exchangetypeid = 2 and yearmonth between '" & mFromDate & "' and '" & mToDate & "' and sm_prefix = '" & TmpRs!SM_Prefix & "' and isnull(commtrackflag,0) = 0)"
        tmp = tmp & " from txn_invmongsldetail  tig where yearmonth between '" & mFromDate & "' and '" & mToDate & "' and isnull(tig.InvoiceAmount,0) > 0"
        tmp = tmp & " and exchangetypeid = 2 and sm_prefix = '" & TmpRs!SM_Prefix & "'"
        tmp = tmp & " and isnull(commtrackflag,0) = 0"
        Call Tmp2Table

        '==========================================================================================
        tmp = "SET IMPLICIT_TRANSACTIONS ON"
        Call Tmp2Table
     
     
        tmp = "Update Txn_InvMonGSLDetail Set Txn_InvMonGSLDetail.FF_RateCardID="
        tmp = tmp & " (Select distinct MNR.FF_RateCardID"
        tmp = tmp & " from  Mst_FF_RateCard_Godown  b"
        tmp = tmp & " inner Join Mst_FF_RateCard MNR On b.FF_RateCardID = MNR.FF_RateCardID"
        tmp = tmp & " inner Join Mst_Godown MG On b.GodownID = MG.GodownID"
        tmp = tmp & " Where (MNR.CommodityGroupID = c.CommodityGroupID or b.BillingUnit='F') And (MNR.exchangetypeid = Txn_InvMonGSLDetail.exchangetypeid Or b.BillingUnit = 'F')"
        tmp = tmp & " And b.GodownID=Txn_InvMonGSLDetail.GodownID And MNR.Flag='A' And MG.AgreementEndDate >= '" & mFromDate & "' And  (('" & mFromDate & "' >= MNR.StartDate And '" & mFromDate & "' <= MNR.EndDate) Or '" & mToDate & "' >= MNR.StartDate And '" & mToDate & "' <= MNR.EndDate Or ('" & mFromDate & "' <= MNR.StartDate And '" & mToDate & "' >= MNR.EndDate)) )"
        tmp = tmp & " From Txn_InvMonGSLDetail"
        tmp = tmp & " inner join mst_commodity c on Txn_InvMonGSLDetail.commodityid = c.commodityid"
        tmp = tmp & " Where Txn_InvMonGSLDetail.YearMonth = '" & mFromDate & "' And Txn_InvMonGSLDetail.SM_Prefix = '" & TmpRs!SM_Prefix & "'"
        
        Call Tmp2Table
        
        
        
        tmp = "SET IMPLICIT_TRANSACTIONS ON"
        Call Tmp2Table
        
        tmp = "Update Txn_InvMonGSLDetail Set Txn_InvMonGSLDetail.FranchBillingCycleID ="
        tmp = tmp & " (Select b.BillingCycleID"
        tmp = tmp & " from  Mst_FF_RateCard_Godown  b"
        tmp = tmp & " inner Join Mst_FF_RateCard MNR On b.FF_RateCardID = MNR.FF_RateCardID"
        tmp = tmp & " inner Join Mst_Godown MG On b.GodownID = MG.GodownID"
        tmp = tmp & " Where (MNR.CommodityGroupID = c.CommodityGroupID or b.BillingUnit='F' )And (MNR.exchangetypeid = Txn_InvMonGSLDetail.exchangetypeid Or b.BillingUnit = 'F')"
        tmp = tmp & " And b.GodownID=Txn_InvMonGSLDetail.GodownID And MNR.Flag='A' And MG.AgreementEndDate >= '" & mFromDate & "' And  (('" & mFromDate & "' >= MNR.StartDate And '" & mFromDate & "' <= MNR.EndDate) Or '" & mToDate & "' >= MNR.StartDate And '" & mToDate & "' <= MNR.EndDate Or ('" & mFromDate & "' <= MNR.StartDate And '" & mToDate & "' >= MNR.EndDate)) )"
        tmp = tmp & " From Txn_InvMonGSLDetail"
        tmp = tmp & " inner join mst_commodity c on Txn_InvMonGSLDetail.commodityid = c.commodityid"
        tmp = tmp & " Where Txn_InvMonGSLDetail.YearMonth = '" & mFromDate & "' And Txn_InvMonGSLDetail.SM_Prefix = '" & TmpRs!SM_Prefix & "'"
        
        Call Tmp2Table
        
        tmp = "SET IMPLICIT_TRANSACTIONS ON"
        Call Tmp2Table
        
        tmp = " Update Txn_InvMonGSLDetail Set Txn_InvMonGSLDetail.FranchBillingUnit ="
        tmp = tmp & " (Select b.BillingUnit"
        tmp = tmp & " from  Mst_FF_RateCard_Godown  b"
        tmp = tmp & " inner Join Mst_FF_RateCard MNR On b.FF_RateCardID = MNR.FF_RateCardID"
        tmp = tmp & " inner Join Mst_Godown MG On b.GodownID = MG.GodownID"
        tmp = tmp & " Where (MNR.CommodityGroupID = c.CommodityGroupID or b.BillingUnit='F' )And (MNR.exchangetypeid = Txn_InvMonGSLDetail.exchangetypeid Or b.BillingUnit = 'F')"
        tmp = tmp & " And b.GodownID=Txn_InvMonGSLDetail.GodownID And MNR.Flag='A' And MG.AgreementEndDate >= '" & mFromDate & "' And  (('" & mFromDate & "' >= MNR.StartDate And '" & mFromDate & "' <= MNR.EndDate) Or '" & mToDate & "' >= MNR.StartDate And '" & mToDate & "' <= MNR.EndDate Or ('" & mFromDate & "' <= MNR.StartDate And '" & mToDate & "' >= MNR.EndDate)))"
        tmp = tmp & " From Txn_InvMonGSLDetail"
        tmp = tmp & " inner join mst_commodity c on Txn_InvMonGSLDetail.commodityid = c.commodityid"
        tmp = tmp & " Where Txn_InvMonGSLDetail.YearMonth = '" & mFromDate & "' And Txn_InvMonGSLDetail.SM_Prefix = '" & TmpRs!SM_Prefix & "'"
        
        Call Tmp2Table
        
        tmp = "SET IMPLICIT_TRANSACTIONS ON"
        Call Tmp2Table
        
        tmp = " Update Txn_InvMonGSLDetail Set Txn_InvMonGSLDetail.FranchiseeRate ="
        tmp = tmp & " (Select b.RateOnQty"
        tmp = tmp & " from  Mst_FF_RateCard_Godown  b"
        tmp = tmp & " inner Join Mst_FF_RateCard MNR On b.FF_RateCardID = MNR.FF_RateCardID"
        tmp = tmp & " inner Join Mst_Godown MG On b.GodownID = MG.GodownID"
        tmp = tmp & " Where (MNR.CommodityGroupID = c.CommodityGroupID or b.BillingUnit='F') And (MNR.exchangetypeid = Txn_InvMonGSLDetail.exchangetypeid Or b.BillingUnit = 'F')"
        tmp = tmp & " And b.GodownID=Txn_InvMonGSLDetail.GodownID And MNR.Flag='A' And MG.AgreementEndDate >= '" & mFromDate & "' And  (('" & mFromDate & "' >= MNR.StartDate And '" & mFromDate & "' <= MNR.EndDate) Or '" & mToDate & "' >= MNR.StartDate And '" & mToDate & "' <= MNR.EndDate Or ('" & mFromDate & "' <= MNR.StartDate And '" & mToDate & "' >= MNR.EndDate)))"
        tmp = tmp & " From Txn_InvMonGSLDetail"
        tmp = tmp & " inner join mst_commodity c on Txn_InvMonGSLDetail.commodityid = c.commodityid"
        tmp = tmp & " Where Txn_InvMonGSLDetail.YearMonth = '" & mFromDate & "' And Txn_InvMonGSLDetail.SM_Prefix = '" & TmpRs!SM_Prefix & "'  And (Txn_InvMonGSLDetail.FranchBillingUnit = 'P' Or Txn_InvMonGSLDetail.FranchBillingUnit = 'Q')"
        
        Call Tmp2Table
        
        
        tmp = "SET IMPLICIT_TRANSACTIONS ON"
        Call Tmp2Table
        '------for Fix rate in date range
        
        tmp = "Update Txn_InvMonGSLDetail Set FranchiseeRate = 0"
        tmp = tmp & " Where Txn_InvMonGSLDetail.YearMonth = '" & mFromDate & "' And Txn_InvMonGSLDetail.SM_Prefix = '" & TmpRs!SM_Prefix & "' And Txn_InvMonGSLDetail.FranchBillingUnit = 'F'"
        tmp = tmp & " And GodownID in (Select distinct b.GodownID"
        tmp = tmp & " from  Mst_FF_RateCard_Godown_Fix  MNRCB"
        tmp = tmp & " inner Join Mst_FF_RateCard_Godown b On b.FF_RateCard_GodownID = MNRCB.FF_RateCard_GodownID"
        tmp = tmp & " inner Join Mst_FF_RateCard MNR On b.FF_RateCardID = MNR.FF_RateCardID"
        tmp = tmp & " inner Join Mst_Godown MG On b.GodownID = MG.GodownID"
        tmp = tmp & " inner Join Mst_Location Ml On Ml.LocationID= MNR.LocationID"
        tmp = tmp & " Where b.BillingUnit='F' And MG.AgreementEndDate >= '" & mFromDate & "' "
        tmp = tmp & " And MNRCB.StartDate<='" & mToDate & "'  And MNRCB.EndDate>='" & mFromDate & "' And ML.SM_Prefix = '" & TmpRs!SM_Prefix & "' And MNR.Flag='A' And  (('" & mFromDate & "' >= MNR.StartDate And '" & mFromDate & "' <= MNR.EndDate) Or '" & mToDate & "' >= MNR.StartDate And '" & mToDate & "' <= MNR.EndDate Or ('" & mFromDate & "' <= MNR.StartDate And '" & mToDate & "' >= MNR.EndDate)))"
        
        Call Tmp2Table
        
        tmp = "SET IMPLICIT_TRANSACTIONS ON"
        Call Tmp2Table
        
        tmp = "Select MNRCB.FixAmount,MNR.exchangetypeid,b.GodownID,MNRCB.StartDate,MNRCB.EndDate"
        tmp = tmp & " from  Mst_FF_RateCard_Godown_Fix  MNRCB"
        tmp = tmp & " inner Join Mst_FF_RateCard_Godown b On b.FF_RateCard_GodownID = MNRCB.FF_RateCard_GodownID"
        tmp = tmp & " inner Join Mst_FF_RateCard MNR On b.FF_RateCardID = MNR.FF_RateCardID"
        tmp = tmp & " inner Join Mst_Godown MG On b.GodownID = MG.GodownID"
        tmp = tmp & " inner Join Mst_Location Ml On Ml.LocationID= MNR.LocationID"
        tmp = tmp & " Where b.BillingUnit='F'"
        tmp = tmp & " And MNRCB.StartDate<='" & mToDate & "'  And MNRCB.EndDate>='" & mFromDate & "' And ML.SM_Prefix = '" & TmpRs!SM_Prefix & "'  And MNR.Flag='A' And MG.AgreementEndDate >= '" & mFromDate & "' And  (('" & mFromDate & "' >= MNR.StartDate And '" & mFromDate & "' <= MNR.EndDate) Or '" & mToDate & "' >= MNR.StartDate And '" & mToDate & "' <= MNR.EndDate Or ('" & mFromDate & "' <= MNR.StartDate And '" & mToDate & "' >= MNR.EndDate))"
        
        Call Tmp2Table
            
        For rn = 1 To TmpRs2.RecordCount
        
            If TmpRs2!StartDate <= CDate(mFromDate) And TmpRs2!EndDate >= CDate(mToDate) Then
               mMonFixRate = TmpRs2!FixAmount
            ElseIf TmpRs2!StartDate <= CDate(mFromDate) And TmpRs2!EndDate < CDate(mToDate) Then
               mTotDays = DateDiff("d", CDate(mFromDate), TmpRs2!EndDate) + 1
               mMonFixRate = (TmpRs2!FixAmount / 30) * mTotDays
            ElseIf TmpRs2!StartDate > CDate(mFromDate) And TmpRs2!EndDate >= CDate(mToDate) Then
               mTotDays = DateDiff("d", TmpRs2!StartDate, CDate(mToDate)) + 1
               mMonFixRate = (TmpRs2!FixAmount / 30) * mTotDays
            ElseIf TmpRs2!StartDate > CDate(mFromDate) And TmpRs2!EndDate < CDate(mToDate) Then
               mTotDays = DateDiff("d", TmpRs2!StartDate, TmpRs2!EndDate) + 1
               mMonFixRate = (TmpRs2!FixAmount / 30) * mTotDays
            End If
            
            tmp = "SET IMPLICIT_TRANSACTIONS ON"
            Call Tmp3Table
            
            tmp = "Update Txn_InvMonGSLDetail Set FranchiseeRate = isnull(FranchiseeRate,0) + " & mMonFixRate
            tmp = tmp & " Where Txn_InvMonGSLDetail.YearMonth = '" & mFromDate & "' And Txn_InvMonGSLDetail.SM_Prefix = '" & TmpRs!SM_Prefix & "' And Txn_InvMonGSLDetail.FranchBillingUnit = 'F'"
            tmp = tmp & " And GodownID=" & TmpRs2!GodownID & ""
               
            Call Tmp3Table
               
            TmpRs2.MoveNext
        Next
        
        '------------
        
        tmp = "SET IMPLICIT_TRANSACTIONS ON"
        Call Tmp2Table
        
        tmp = " Update Txn_InvMonGSLDetail Set Txn_InvMonGSLDetail.FranchiseeRate ="
        tmp = tmp & " (Select MNRCB.RatePerBag"
        tmp = tmp & " from  Mst_FF_RateCard_Godown_Bag  MNRCB"
        tmp = tmp & " inner Join Mst_FF_RateCard_Godown b On b.FF_RateCard_GodownID = MNRCB.FF_RateCard_GodownID"
        tmp = tmp & " inner Join Mst_FF_RateCard MNR On b.FF_RateCardID = MNR.FF_RateCardID"
        tmp = tmp & " inner Join Mst_Godown MG On b.GodownID = MG.GodownID"
        tmp = tmp & " Where (MNR.CommodityGroupID = c.CommodityGroupID or b.BillingUnit='F' )And (MNR.exchangetypeid = Txn_InvMonGSLDetail.exchangetypeid Or b.BillingUnit = 'F')"
        tmp = tmp & " And b.GodownID=Txn_InvMonGSLDetail.GodownID"
        tmp = tmp & " And Txn_InvMonGSLDetail.BagSize >= MNRCB.BagSizeFrom And Txn_InvMonGSLDetail.BagSize <= MNRCB.BagSizeTo  And MNR.Flag='A' And MG.AgreementEndDate >= '" & mFromDate & "' And  (('" & mFromDate & "' >= MNR.StartDate And '" & mFromDate & "' <= MNR.EndDate) Or '" & mToDate & "' >= MNR.StartDate And '" & mToDate & "' <= MNR.EndDate Or ('" & mFromDate & "' <= MNR.StartDate And '" & mToDate & "' >= MNR.EndDate)) )"
        tmp = tmp & " From Txn_InvMonGSLDetail"
        tmp = tmp & " inner join mst_commodity c on Txn_InvMonGSLDetail.commodityid = c.commodityid"
        tmp = tmp & " Where Txn_InvMonGSLDetail.YearMonth = '" & mFromDate & "' And Txn_InvMonGSLDetail.SM_Prefix = '" & TmpRs!SM_Prefix & "' And Txn_InvMonGSLDetail.FranchBillingUnit = 'B'"
        
        Call Tmp2Table
        
        tmp = "SET IMPLICIT_TRANSACTIONS ON"
        Call Tmp2Table
        
        
        'Op Bal and Monthly Cycle
        tmp = " Update Txn_InvMonGSLDetail Set FranchNoofDays = 1 Where YearMonth = '" & mFromDate & "' And TFlag = 'O'"
        tmp = tmp & " And FranchBillingCycleID = 4  And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
    
        Call Tmp2Table
           
        tmp = "SET IMPLICIT_TRANSACTIONS ON"
        Call Tmp2Table
           
        'Op Bal and Forthnightly Cycle
        tmp = " Update Txn_InvMonGSLDetail Set FranchNoofDays = 2 Where YearMonth = '" & mFromDate & "' And TFlag = 'O'"
        tmp = tmp & " And FranchBillingCycleID = 3 And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
        
        Call Tmp2Table
        
        tmp = "SET IMPLICIT_TRANSACTIONS ON"
        Call Tmp2Table
        
        'Op Bal and Weekly Cycle
        tmp = " Update Txn_InvMonGSLDetail Set FranchNoofDays = 4 Where YearMonth = '" & mFromDate & "' And TFlag = 'O'"
        tmp = tmp & " And FranchBillingCycleID = 2 And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
        
        Call Tmp2Table
        
        'ProgressBar1.Value = ProgressBar1.Value + 1
        tmp = "SET IMPLICIT_TRANSACTIONS ON"
        Call Tmp2Table
        
        'Op Bal and Daily Cycle
        tmp = " Update Txn_InvMonGSLDetail Set FranchNoofDays = " & mTotalDays & " Where YearMonth = '" & mFromDate & "' And TFlag = 'O'"
        tmp = tmp & " And FranchBillingCycleID = 1 And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
        
        Call Tmp2Table
         
        'ProgressBar1.Value = ProgressBar1.Value + 1
        tmp = "SET IMPLICIT_TRANSACTIONS ON"
        Call Tmp2Table
         
        'Depo and Withd and Monthly Cycle
        tmp = " Update Txn_InvMonGSLDetail Set FranchNoofDays = 1 Where YearMonth = '" & mFromDate & "' And TFlag <> 'O'"
        tmp = tmp & " And FranchBillingCycleID = 4 And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
           
        Call Tmp2Table
            
        'ProgressBar1.Value = ProgressBar1.Value + 1
        tmp = "SET IMPLICIT_TRANSACTIONS ON"
        Call Tmp2Table
        
        'Depo and Withd and Daily Cycle
        tmp = "Update Txn_InvMonGSLDetail Set FranchNoofDays = " & mTotalDays & " - Day(CxTFDate) + 1 Where YearMonth = '" & mFromDate & "' And TFlag = 'D' And FranchBillingCycleID = 1 And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
            
        Call Tmp2Table
            
        tmp = "SET IMPLICIT_TRANSACTIONS ON"
        Call Tmp2Table
        'ProgressBar1.Value = ProgressBar1.Value + 1
    
        'Withd and Daily Cycle
        tmp = "Update Txn_InvMonGSLDetail Set FranchNoOfDays = Day(CxTFDate) Where YearMonth = '" & mFromDate & "' And TFlag = 'W' And FranchBillingCycleID = 1 And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
          
        Call Tmp2Table
    
        'ProgressBar1.Value = ProgressBar1.Value + 1
        tmp = "SET IMPLICIT_TRANSACTIONS ON"
        Call Tmp2Table
        
        'Deposit & Weekly
        
        mNoofDays = GetBillingCycleDays(2)
        
        tmp = "SET IMPLICIT_TRANSACTIONS ON"
        Call Tmp2Table
        
        tmp = "Update Txn_InvMonGSLDetail Set FranchNoOfDays =  Case Floor((4 - (Convert(Numeric,Day(CxTFDate))/" & mNoofDays & ") + 1))  When 0 then 1 else Floor((4 - (Convert(Numeric,Day(CxTFDate))/" & mNoofDays & ") + 1)) End "
        tmp = tmp & " Where Txn_InvMonGSLDetail.YearMonth = '" & mFromDate & "'"
        tmp = tmp & " And Txn_InvMonGSLDetail.TFlag = 'D' And Txn_InvMonGSLDetail.FranchBillingCycleID = 2 And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
    
        Call Tmp2Table
        
        'ProgressBar1.Value = ProgressBar1.Value + 1
            
        'Deposit & FortNightly
        
        mNoofDays = GetBillingCycleDays(3)
        
        'Case 1 Depo Date  <= 15 Then No Of Cycle = 2
        
        tmp = "SET IMPLICIT_TRANSACTIONS ON"
        Call Tmp2Table
        
        
        tmp = " Update Txn_InvMonGSLDetail Set FranchNoOfDays = 2 "
        tmp = tmp & " Where Txn_InvMonGSLDetail.YearMonth = '" & mFromDate & "' "
        tmp = tmp & " And Txn_InvMonGSLDetail.TFlag = 'D' And Txn_InvMonGSLDetail.FranchBillingCycleID = 3 And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
        tmp = tmp & " And Day(CxTFDate) <= 15 "
        
        Call Tmp2Table
              
        'ProgressBar1.Value = ProgressBar1.Value + 1
        
        tmp = "SET IMPLICIT_TRANSACTIONS ON"
        Call Tmp2Table
        
        'Case 1 Depo Date  > 15 Then No Of Cycle = 1
        tmp = " Update Txn_InvMonGSLDetail Set FranchNoOfDays = 1 "
        tmp = tmp & " Where YearMonth = '" & mFromDate & "' "
        tmp = tmp & " And TFlag = 'D' And FranchBillingCycleID = 3 "
        tmp = tmp & " And Day(CxTFDate) > 15 And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
        
        Call Tmp2Table
              
        'ProgressBar1.Value = ProgressBar1.Value + 1
        
        'Withdrawal & weekly
        mNoofDays = GetBillingCycleDays(2)
                   
        tmp = "SET IMPLICIT_TRANSACTIONS ON"
        Call Tmp2Table
                   
                   
        tmp = "Update Txn_InvMonGSLDetail Set FranchNoOfDays = Case CEILING(((Round(Convert(Numeric,Day(CxTFDate))/" & mNoofDays & ",1)))) When 5 Then 4 Else CEILING(((Round(Convert(Numeric,Day(CxTFDate))/" & mNoofDays & ",1)))) End "
        tmp = tmp & " Where YearMonth = '" & mFromDate & "' "
        tmp = tmp & " And TFlag = 'W' And FranchBillingCycleID = 2 And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
           
        Call Tmp2Table
        
        'ProgressBar1.Value = ProgressBar1.Value + 1
           
        'Withdrawal & Fortnightly
        mNoofDays = GetBillingCycleDays(3)
        
        tmp = "SET IMPLICIT_TRANSACTIONS ON"
        Call Tmp2Table
        
        
        tmp = "Update Txn_InvMonGSLDetail Set FranchNoOfDays = Case CEILING(((Round(Convert(Numeric,Day(CxTFDate))/" & mNoofDays & ",1)))) When 3 Then 2 Else CEILING(((Round(Convert(Numeric,Day(CxTFDate))/" & mNoofDays & ",1)))) End "
        tmp = tmp & " Where YearMonth = '" & mFromDate & "' "
        tmp = tmp & " And TFlag = 'W' And FranchBillingCycleID = 3 And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
        
        Call Tmp2Table
                
        'ProgressBar1.Value = ProgressBar1.Value + 1
        tmp = "SET IMPLICIT_TRANSACTIONS ON"
        Call Tmp2Table
        
        
        'Updating Invoice Amount And Invoice Date
        tmp = " Update Txn_InvMonGSLDetail Set FranchInvoiceAmount = BalanceBags * FranchNoOfDays * FranchiseeRate, InvoiceDate = '" & mToDate & "' "
        tmp = tmp & " Where YearMonth = '" & mFromDate & "'"
        tmp = tmp & " And FranchBillingUnit = 'B' And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
               
        Call Tmp2Table
    
        'ProgressBar1.Value = ProgressBar1.Value + 1
        tmp = "SET IMPLICIT_TRANSACTIONS ON"
        Call Tmp2Table
    
    
        tmp = " Update Txn_InvMonGSLDetail Set FranchInvoiceAmount = BalanceWeight * FranchNoOfDays * FranchiseeRate,InvoiceDate =  '" & mToDate & "' "
        tmp = tmp & " Where YearMonth = '" & mFromDate & "'"
        tmp = tmp & " And FranchBillingUnit = 'Q' And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
            
        Call Tmp2Table
    
        tmp = "SET IMPLICIT_TRANSACTIONS ON"
        Call Tmp2Table
    
    '--For Fixed
    
        tmp = "Select GodownID,FranchiseeRate,max(Mst_GSL_ID)GSLID,FF_RateCardID from Txn_InvMonGSLDetail"
        tmp = tmp & " Where YearMonth = '" & mFromDate & "' And FranchBillingUnit = 'F' And SM_Prefix = '" & TmpRs!SM_Prefix & "' And Mst_GSL_ID>0  And FranchiseeRate>0"
        tmp = tmp & " Group by GodownID,FranchiseeRate,FF_RateCardID"
    
        Call Tmp2Table
    
        For rn = 1 To TmpRs2.RecordCount
        
            tmp = "Select * from Txn_InvMonGSLDetail Where YearMOnth = '" & mFromDate & "' And StackNo = 'Dummy-F' And Mst_GSL_ID = 0 And GodownID = " & TmpRs2!GodownID & ""
            Call Tmp4Table
            
            If TmpRs4.RecordCount = 0 Then
               
               tmp = "SET IMPLICIT_TRANSACTIONS ON"
               Call Tmp4Table
               
               tmp = "Update Txn_InvMonGSLDetail Set FranchInvoiceAmount = " & TmpRs2!FranchiseeRate & ",InvoiceDate =  '" & mToDate & "'"
               tmp = tmp & " Where YearMonth = '" & mFromDate & "' And Mst_GSL_ID=" & TmpRs2!GSLID & " And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
            
               Call Tmp4Table
                
               tmp = "SET IMPLICIT_TRANSACTIONS ON"
               Call Tmp4Table
                
               tmp = "Select Mst_GSL_ID,FranchInvoiceAmount,BalanceBags,MonG_UID from  Txn_InvMonGSLDetail "
               tmp = tmp & " Where YearMonth = '" & mFromDate & "' And Mst_GSL_ID=" & TmpRs2!GSLID & " And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
               
               Call Tmp3Table
               
               If TmpRs3.RecordCount > 1 Then
                  For rn1 = 1 To TmpRs3.RecordCount - 1
                      TmpRs3!FranchInvoiceAmount = 0
                      TmpRs3.Update
                      TmpRs3.MoveNext
                  Next
               End If
                
               tmp = "SET IMPLICIT_TRANSACTIONS ON"
               Call Tmp4Table
                
               tmp = "Update Txn_InvMonGSLDetail Set FF_RateCardID=" & TmpRs2!FF_RateCardID & ",FranchBillingUnit='F',FranchNoofDays=1,FranchiseeRate=" & TmpRs2!FranchiseeRate & ",FranchInvoiceAmount=0,FranchBillingCycleID=4"
               tmp = tmp & " Where YearMonth = '" & mFromDate & "' And SM_Prefix = '" & TmpRs!SM_Prefix & "' And GodownID=" & TmpRs2!GodownID & " And Mst_GSL_ID<>" & TmpRs2!GSLID & ""
                
               Call Tmp4Table
            Else
               tmp = "SET IMPLICIT_TRANSACTIONS ON"
               Call Tmp4Table
               
               
               tmp = "Update Txn_InvMonGSLDetail Set FF_RateCardID=" & TmpRs2!FF_RateCardID & ",FranchBillingUnit='F',FranchNoofDays=1,FranchiseeRate=" & TmpRs2!FranchiseeRate & ",FranchInvoiceAmount=0,FranchBillingCycleID=4"
               tmp = tmp & " Where YearMonth = '" & mFromDate & "' And SM_Prefix = '" & TmpRs!SM_Prefix & "' And GodownID=" & TmpRs2!GodownID & " And Mst_GSL_ID > 0  "
                
               Call Tmp4Table
            
            End If
            TmpRs2.MoveNext
        Next
        
        '-----For Percent
        
        tmp = "SET IMPLICIT_TRANSACTIONS ON"
        Call Tmp2Table
        
''Commented by Parag on 29 June 2022
''        tmp = "Select Mst_GSL_ID,GodownID,0'CommodityID',Sum(InvoiceAmtForFranch)'InvoiceAmtForFranch',Sum(AUM)*(0.02/100) 'AUM' From ("
''        tmp = tmp & " Select GSL.Mst_GSL_ID,GSL.GodownID,GSL.CommodityID,Sum(InvoiceAmtForFranch) 'InvoiceAmtForFranch',Sum(GSL.BalanceWeight)*Isnull(CRate.Rate,0)[AUM]"
''        tmp = tmp & " From Txn_InvmonGSLDetail GSL"
''        tmp = tmp & " Inner Join Mst_Location ML On GSL.SM_Prefix = ML.SM_Prefix"
''        tmp = tmp & " Inner Join Mst_State MS On ML.StateID = MS.StateID"
''        tmp = tmp & " inner Join Mst_Godown G on GSL.GOdownid = G.GodownID"
''        tmp = tmp & " Inner Join Mst_Commodity MC On GSL.CommodityID = MC.CommodityID"
''        tmp = tmp & " Left Join"
''        tmp = tmp & " (   Select MCP.StateID,MCP.CommodityID,MCP.YearMonth,"
''        tmp = tmp & "     Case When Convert(date,MCP.YearMonth+'-01') = Convert(date,Convert(Varchar,DATEPART(Year,GETDATE()))+'-'+Convert(Varchar,DATEPART(Month,GETDATE()))+'-01')"
''        tmp = tmp & "     Then"
''        tmp = tmp & "       COALESCE(MCP.Day31,MCP.Day30,MCP.Day29,MCP.Day28,MCP.Day27,MCP.Day26,MCP.Day25,MCP.Day24,MCP.Day23,MCP.Day22,MCP.Day21,MCP.Day20,MCP.Day19,MCP.Day18,MCP.Day17,MCP.Day16,MCP.Day15,MCP.Day14,MCP.Day13,MCP.Day12,MCP.Day11,MCP.Day10,MCP.Day9,MCP.Day8,MCP.Day7,MCP.Day6,MCP.Day5,MCP.Day4,MCP.Day3,MCP.Day2,MCP.Day1)"
''        tmp = tmp & "     Else"
''        tmp = tmp & "       COALESCE(MCP.Day31,MCP.Day30,MCP.Day29,MCP.Day28,MCP.Day27,MCP.Day26,MCP.Day25,MCP.Day24,MCP.Day23,MCP.Day22,MCP.Day21,MCP.Day20,MCP.Day19,MCP.Day18,MCP.Day17,MCP.Day16,MCP.Day15,MCP.Day14,MCP.Day13,MCP.Day12,MCP.Day11,MCP.Day10,MCP.Day9,MCP.Day8,MCP.Day7,MCP.Day6,MCP.Day5,MCP.Day4,MCP.Day3,MCP.Day2,MCP.Day1)"
''        tmp = tmp & "     End [Rate]"
''        tmp = tmp & "     from Mst_CommodityPrice MCP"
''        tmp = tmp & "     Where Convert(Varchar,StateID)+'~'+Convert(Varchar,CommodityID)+'~'+Convert(Varchar,YearMonth) In  (Select Convert(Varchar,StateID)+'~'+Convert(Varchar,CommodityID)+'~'+Convert(Varchar,MAX(YearMonth)) From Mst_CommodityPrice Where YearMonth <= '2022-05' Group By StateID,CommodityID )"
''        tmp = tmp & " ) As CRate On MS.StateID = CRate.StateID And GSL.CommodityID = CRate.CommodityID"
''        tmp = tmp & " Where GSL.TFLAG  In ('O') And GSL.BalanceWeight > 0 ANd GSL.YearMOnth = '" & CDate(mFromDate) & "' And g.UnitTypeID in (2,6) And GSL.SM_Prefix = '" & TmpRs!SM_Prefix & "' "
''        tmp = tmp & " Group By GSL.Mst_GSL_ID,GSL.CommodityID,GSL.GodownID,CRate.Rate)"
''        tmp = tmp & " A"
''        tmp = tmp & " Group BY Mst_GSL_ID,GOdownID"
''
''        Call Tmp1Table
''
''        For ab = 1 To TmpRs1.RecordCount
''
''            mICost = IIf(IsNull(TmpRs1!AUM), 0, TmpRs1!AUM)
''
''            tmp = " Update Txn_InvMonGSLDetail Set FranchInvoiceAmount =  (InvoiceAmtForFranch - " & mICost & ")* FranchiseeRate /100 ,InvoiceDate =  '" & mToDate & "' "
''            tmp = tmp & " Where YearMonth = '" & mFromDate & "' And InvoiceAmtForFranch > 0 "
''            tmp = tmp & " And FranchBillingUnit = 'P' And SM_Prefix = '" & TmpRs!SM_Prefix & "' And Mst_GSL_ID = " & TmpRs1!Mst_GSL_ID & " "
''
''            Call Tmp2Table
''
''            TmpRs1.MoveNext
''        Next ab
''
''        tmp = "Select Mst_GSL_ID,GodownID,0'CommodityID',Sum(InvoiceAmtForFranch)'InvoiceAmtForFranch',Sum(AUM)*(0.02/100) 'AUM' From ("
''        tmp = tmp & " Select GSL.Mst_GSL_ID,GSL.GodownID,GSL.CommodityID,Sum(InvoiceAmtForFranch) 'InvoiceAmtForFranch',(Max(GSL.BalanceWeight)* Isnull(CRate.Rate,0)* (((" & mTotalDays & "-Convert(decimal,Day(GSL.CxTFDate))+1)/" & mTotalDays & ")))[AUM]"
''        tmp = tmp & " From Txn_InvmonGSLDetail GSL"
''        tmp = tmp & " Inner Join Mst_Location ML On GSL.SM_Prefix = ML.SM_Prefix"
''        tmp = tmp & " Inner Join Mst_State MS On ML.StateID = MS.StateID"
''        tmp = tmp & " inner Join Mst_Godown G on GSL.GOdownid = G.GodownID"
''        tmp = tmp & " Inner Join Mst_Commodity MC On GSL.CommodityID = MC.CommodityID"
''        tmp = tmp & " Left Join"
''        tmp = tmp & " (   Select MCP.StateID,MCP.CommodityID,MCP.YearMonth,"
''        tmp = tmp & "     Case When Convert(date,MCP.YearMonth+'-01') = Convert(date,Convert(Varchar,DATEPART(Year,GETDATE()))+'-'+Convert(Varchar,DATEPART(Month,GETDATE()))+'-01')"
''        tmp = tmp & "     Then"
''        tmp = tmp & "       COALESCE(MCP.Day31,MCP.Day30,MCP.Day29,MCP.Day28,MCP.Day27,MCP.Day26,MCP.Day25,MCP.Day24,MCP.Day23,MCP.Day22,MCP.Day21,MCP.Day20,MCP.Day19,MCP.Day18,MCP.Day17,MCP.Day16,MCP.Day15,MCP.Day14,MCP.Day13,MCP.Day12,MCP.Day11,MCP.Day10,MCP.Day9,MCP.Day8,MCP.Day7,MCP.Day6,MCP.Day5,MCP.Day4,MCP.Day3,MCP.Day2,MCP.Day1)"
''        tmp = tmp & "     Else"
''        tmp = tmp & "       COALESCE(MCP.Day31,MCP.Day30,MCP.Day29,MCP.Day28,MCP.Day27,MCP.Day26,MCP.Day25,MCP.Day24,MCP.Day23,MCP.Day22,MCP.Day21,MCP.Day20,MCP.Day19,MCP.Day18,MCP.Day17,MCP.Day16,MCP.Day15,MCP.Day14,MCP.Day13,MCP.Day12,MCP.Day11,MCP.Day10,MCP.Day9,MCP.Day8,MCP.Day7,MCP.Day6,MCP.Day5,MCP.Day4,MCP.Day3,MCP.Day2,MCP.Day1)"
''        tmp = tmp & "     End [Rate]"
''        tmp = tmp & "     from Mst_CommodityPrice MCP"
''        tmp = tmp & "     Where Convert(Varchar,StateID)+'~'+Convert(Varchar,CommodityID)+'~'+Convert(Varchar,YearMonth) In  (Select Convert(Varchar,StateID)+'~'+Convert(Varchar,CommodityID)+'~'+Convert(Varchar,MAX(YearMonth)) From Mst_CommodityPrice Where YearMonth <= '2022-05' Group By StateID,CommodityID )"
''        tmp = tmp & " ) As CRate On MS.StateID = CRate.StateID And GSL.CommodityID = CRate.CommodityID"
''        tmp = tmp & " Where GSL.TFLAG  In ('D') And GSL.BalanceWeight > 0 ANd GSL.YearMOnth = '" & CDate(mFromDate) & "' And g.UnitTypeID in (2,6) And GSL.SM_Prefix = '" & TmpRs!SM_Prefix & "' "
''        tmp = tmp & " Group By GSL.Mst_GSL_ID,GSL.CommodityID,GSL.GodownID,CRate.Rate,GSL.CxTFDate)"
''        tmp = tmp & " A"
''        tmp = tmp & " Group BY Mst_GSL_ID,GOdownID"
''
''        Call Tmp1Table
''
''        For ab = 1 To TmpRs1.RecordCount
''
''            mICost = IIf(IsNull(TmpRs1!AUM), 0, TmpRs1!AUM)
''
''            tmp = " Update Txn_InvMonGSLDetail Set FranchInvoiceAmount =  (InvoiceAmtForFranch - " & mICost & ")* FranchiseeRate /100 ,InvoiceDate =  '" & mToDate & "' "
''            tmp = tmp & " Where YearMonth = '" & mFromDate & "' And InvoiceAmtForFranch > 0 "
''            tmp = tmp & " And FranchBillingUnit = 'P' And SM_Prefix = '" & TmpRs!SM_Prefix & "' And Mst_GSL_ID = " & TmpRs1!Mst_GSL_ID & " "
''
''            Call Tmp2Table
''
''            TmpRs1.MoveNext
''        Next ab
    
       
       tmp = "Select Distinct GodownID from Txn_InvMonGSLDetail Where YearMonth = '" & mFromDate & "' And SM_Prefix = '" & TmpRs!SM_Prefix & "' And InvoiceAmtForFranch > 0 And GodownID in (Select GodownID from Mst_GOdown Where UnitTypeID in (2,6))"
       Call Tmp4Table
       
       For pq = 1 To TmpRs4.RecordCount
           
           If TmpRs4!GodownID = "5771" Then
              a = a
           End If
           
           tmp = "Select Distinct GodownID,ClientID,CommodityID from Txn_InvMonGSLDetail Where YearMonth = '" & mFromDate & "' And SM_Prefix = '" & TmpRs!SM_Prefix & "' And InvoiceAmtForFranch > 0 And GodownID = " & TmpRs4!GodownID & ""
           Call Tmp6Table
                        
           For wxy = 1 To TmpRs6.RecordCount
           
               If TmpRs6!CommodityID = "14" Then
                  a = a
               End If
               
               
                tmp = "select isnull(sum(AUM),0) AUM from ("
                tmp = tmp & " Select Max(isnull(a.AUM,0)) 'AUM' from Txn_ArchiveInsurancePolicyData a"
                tmp = tmp & " Inner Join Mst_Godown g on a.GOdownID = g.GodownID"
                tmp = tmp & " Inner Join Mst_CMP c on g.CMPID = c.CMPID"
                tmp = tmp & " Inner Join Mst_Location l on c.LOcationID = l.LocationID"
                tmp = tmp & " Where a.YearMonth = '" & Year(mFromDate) & "-0" & Month(mFromDate) & "' And a.GodownID = " & TmpRs6!GodownID & " And g.UnitTypeID in (2,6) and l.SM_Prefix = '" & TmpRs!SM_Prefix & "' And a.ClientNCMLStockFlag = 'N'"
                tmp = tmp & " Group By a.YearMonth,a.GodownID,ClientNCMLStockFlag,a.ClientID,a.CommodityID)T"
                
                Call Tmp1Table
                Dim mGodownWiseAUM As Variant
                mGodownWiseAUM = 0
                
                If TmpRs1.RecordCount > 0 Then
                For aa = 1 To TmpRs1.RecordCount
                    mGodownWiseAUM = TmpRs1!AUM
                    TmpRs1.MoveNext
                Next aa
                End If
                
               
               tmp = "Select a.YearMonth,a.GodownID,Isnull(a.ClientID,'') 'ClientID',a.CommodityID,ClientNCMLStockFlag,Max(a.AUM) 'AUM',"
               'tmp = " Max(a.AUM)*0.02/100 'InsuranceAmt'  from Txn_ArchiveInsurancePolicyData a"
               tmp = tmp & " case when (" & mGodownWiseAUM & " >= 50000000) then Max(a.AUM)*0.03/100 else  Max(a.AUM)*0.02/100 end 'InsuranceAmt'  from Txn_ArchiveInsurancePolicyData a"
               tmp = tmp & " Inner Join Mst_Godown g on a.GOdownID = g.GodownID"
               tmp = tmp & " Inner Join Mst_CMP c on g.CMPID = c.CMPID"
               tmp = tmp & " Inner Join Mst_Location l on c.LOcationID = l.LocationID"
               'tmp = tmp & " Where a.YearMonth = '2022-02' And a.GodownID = " & TmpRs6!GodownID & " And a.ClientID = " & TmpRs6!ClientID & " And a.CommodityID = " & TmpRs6!CommodityID & " "
               If Month(mFromDate) < 10 Then
                  tmp = tmp & " Where a.YearMonth = '" & Year(mFromDate) & "-0" & Month(mFromDate) & "' And a.GodownID = " & TmpRs6!GodownID & " And a.ClientID = " & TmpRs6!ClientID & " And a.CommodityID = " & TmpRs6!CommodityID & " "
               Else
                  tmp = tmp & " Where a.YearMonth = '" & Year(mFromDate) & "-" & Month(mFromDate) & "' And a.GodownID = " & TmpRs6!GodownID & " And a.ClientID = " & TmpRs6!ClientID & " And a.CommodityID = " & TmpRs6!CommodityID & " "
               End If
               'tmp = tmp & " And Convert(Varchar,Isnull(a.ClientID,''))+'-'+Convert(Varchar,a.CommodityID)+'-'+Convert(Varchar,a.GodownID) not in  (Select Convert(Varchar,Isnull(ClientID,''))+'-'+Convert(Varchar,CommodityID)+'-'+Convert(Varchar,GodownID) from Txn_ArchiveInsurancePolicyData Where YearMonth = '2022-05' And ClientNCMLStockFlag = 'C' )"
               tmp = tmp & " And g.UnitTypeID in (2,6) and l.SM_Prefix = '" & TmpRs!SM_Prefix & "'"
               tmp = tmp & " Group By a.YearMonth,a.GodownID,a.ClientID,a.CommodityID,ClientNCMLStockFlag order by a.YearMonth"
               
               Call Tmp1Table
                
               If TmpRs1.RecordCount > 0 Then
                  
                  For abc = 1 To TmpRs1.RecordCount
                      
                       If TmpRs1!ClientNCMLStockFlag = "N" Then
                              
                          mTotInvoiceAmt = ""
                          mTotFranchBagAmt = ""
                          mTotFranchQtyAmt = ""
                          
                          tmp = "Select Sum(InvoiceAmtForFranch) 'TotInvoiceAmt' from Txn_InvMonGSLDetail Where YearMonth = '" & mFromDate & "' And SM_Prefix = '" & TmpRs!SM_Prefix & "' And GodownID = " & TmpRs1!GodownID & " And CommodityID = " & TmpRs1!CommodityID & " And  ClientID = " & TmpRs1!ClientID & " And InvoiceAmtForFranch > 0 And FranchBillingUnit = 'P' "
                          Call Tmp2Table
                          If TmpRs2.RecordCount > 0 Then
                             mTotInvoiceAmt = IIf(IsNull(TmpRs2!TotInvoiceAmt), 0, TmpRs2!TotInvoiceAmt)
                             'tmp = "Select * from Txn_InvMonGSLDetail Where YearMonth = '" & mFromDate & "' And SM_Prefix = '" & TmpRs!SM_Prefix & "' And GodownID = " & TmpRs1!GodownID & " And CommodityID = " & TmpRs1!CommodityID & " And  ClientID = " & TmpRs1!ClientID & " And InvoiceAmtForFranch > 0 And FranchBillingUnit = 'P' "
                             'Call Tmp2Table
                          End If
                          
                          tmp = "Select Sum(FranchInvoiceAmount) 'TotInvoiceAmt' from Txn_InvMonGSLDetail Where YearMonth = '" & mFromDate & "' And SM_Prefix = '" & TmpRs!SM_Prefix & "' And GodownID = " & TmpRs1!GodownID & " And CommodityID = " & TmpRs1!CommodityID & " And  ClientID = " & TmpRs1!ClientID & " And InvoiceAmtForFranch > 0 And FranchBillingUnit = 'B' "
                          Call Tmp2Table
                          
                          If TmpRs2.RecordCount > 0 Then
                             mTotFranchBagAmt = IIf(IsNull(TmpRs2!TotInvoiceAmt), 0, TmpRs2!TotInvoiceAmt)
                             'tmp = "Select * from Txn_InvMonGSLDetail Where YearMonth = '" & mFromDate & "' And SM_Prefix = '" & TmpRs!SM_Prefix & "' And GodownID = " & TmpRs1!GodownID & " And CommodityID = " & TmpRs1!CommodityID & " And  ClientID = " & TmpRs1!ClientID & " And InvoiceAmtForFranch > 0 And FranchBillingUnit = 'B' "
                             'Call Tmp2Table
                          End If
                          
                          tmp = "Select Sum(FranchInvoiceAmount) 'TotInvoiceAmt' from Txn_InvMonGSLDetail Where YearMonth = '" & mFromDate & "' And SM_Prefix = '" & TmpRs!SM_Prefix & "' And GodownID = " & TmpRs1!GodownID & " And CommodityID = " & TmpRs1!CommodityID & " And  ClientID = " & TmpRs1!ClientID & " And InvoiceAmtForFranch > 0 And FranchBillingUnit = 'Q' "
                          Call Tmp2Table
                      
                          If TmpRs2.RecordCount > 0 Then
                             mTotFranchQtyAmt = IIf(IsNull(TmpRs2!TotInvoiceAmt), 0, TmpRs2!TotInvoiceAmt)
                             'tmp = "Select * from Txn_InvMonGSLDetail Where YearMonth = '" & mFromDate & "' And SM_Prefix = '" & TmpRs!SM_Prefix & "' And GodownID = " & TmpRs1!GodownID & " And CommodityID = " & TmpRs1!CommodityID & " And  ClientID = " & TmpRs1!ClientID & " And InvoiceAmtForFranch > 0 And FranchBillingUnit = 'Q' "
                             'Call Tmp2Table
                          
                          End If
                      
                          tmp = "Select * from Txn_InvMonGSLDetail Where YearMonth = '" & mFromDate & "' And SM_Prefix = '" & TmpRs!SM_Prefix & "' And GodownID = " & TmpRs1!GodownID & " And CommodityID = " & TmpRs1!CommodityID & " And  ClientID = " & TmpRs1!ClientID & " And InvoiceAmtForFranch > 0 "
                          Call Tmp2Table
                     
                          If TmpRs2.RecordCount > 0 Then
                             For pqr = 1 To TmpRs2.RecordCount
                                 If IsNull(TmpRs2!FF_RateCardID) = True Then
                                    mRCommodityName = GetCommodityName(TmpRs2!CommodityID)
                                    mRGOdownno = GetGodownName(TmpRs2!GodownID, "N")
                                    mRCMPID = GetCMPName(TmpRs2!CMPID)
                                    MsgBox "No rate card found for CMP-Godown-Commodity " & mRCMPID & "-" & mRGOdownno & "-" & mRCommodityName & " !!!"
                                    
                                    tmp = "UPDATE Txn_InvMonGSLDetail SET FranchBillingCycleID = Null,FranchBillingUnit = Null,FranchNoofDays = Null,FranchInvoiceAmount = Null,FF_RateCardID = Null,FranchClaimID = Null,FranchProcessFlag = 'N',FranchProcessdBy = Null,FranchProcessedDate = Null,InvoiceAmtForFranch = 0,InsuranceCostDeduction=0 WHere YearMonth = '" & mFromDate & "' And SM_Prefix = '" & mSM_Prefix & "'"
                                    Call Tmp7Table
                                    
                                    Exit Sub
                                 End If
                                 tmp = "Select Isnull(InsuranceDedException,0) 'InsuranceDedException' from Mst_FF_RateCard_Godown Where FF_RateCardID = " & TmpRs2!FF_RateCardID & " And GodownID =  " & TmpRs2!GodownID & " "
                                 Call Tmp7Table
                                  
                                 If TmpRs7.RecordCount > 0 Then
                                    If TmpRs7!InsuranceDedException = False Then
                                       If Val(mTotInvoiceAmt) <> 0 Then
                                          iCost = Format(TmpRs1!InsuranceAmt * TmpRs2!InvoiceAmtForFranch / Val(mTotInvoiceAmt), "#############.00")
                                       ElseIf Val(mTotFranchBagAmt) <> 0 Then
                                          iCost = Format(TmpRs1!InsuranceAmt * TmpRs2!FranchInvoiceAmount / Val(mTotFranchBagAmt), "#############.00")
                                       ElseIf Val(mTotFranchQtyAmt) <> 0 Then
                                          iCost = Format(TmpRs1!InsuranceAmt * TmpRs2!FranchInvoiceAmount / Val(mTotFranchQtyAmt), "#############.00")
                                       Else
                                          iCost = 0
                                       End If
                                    Else
                                       iCost = 0
                                    End If
                                 End If
                                  
                                 'iCost = Format(TmpRs1!InsuranceAmt * TmpRs2!InvoiceAmtForFranch / Val(mTotInvoiceAmt), "#############.00")
                                
                                 
'                                 tmp = " Update Txn_InvMonGSLDetail Set FranchInvoiceAmount =  (InvoiceAmtForFranch - " & iCost & ")* FranchiseeRate /100 ,InvoiceDate =  '" & mToDate & "',InsuranceCostDeduction = InsuranceCostDeduction + " & iCost & " "
'                                 tmp = tmp & " Where YearMonth = '" & mFromDate & "' And InvoiceAmtForFranch > 0 "
'                                 tmp = tmp & " And FranchBillingUnit = 'P' And SM_Prefix = '" & TmpRs!SM_Prefix & "' And ID = " & TmpRs2!ID & " "
                               If TmpRs6!CommodityID = "14" Then
                                   a = a
                                End If
                               
                                 If mTotInvoiceAmt <> "0" Then
                                    tmp = " Update Txn_InvMonGSLDetail Set FranchInvoiceAmount =  (InvoiceAmtForFranch - " & iCost & ")* FranchiseeRate /100 ,InvoiceDate =  '" & mToDate & "',InsuranceCostDeduction = InsuranceCostDeduction + " & iCost & " "
                                    tmp = tmp & " Where YearMonth = '" & mFromDate & "' And InvoiceAmtForFranch > 0 "
                                    tmp = tmp & " And FranchBillingUnit = 'P' And SM_Prefix = '" & TmpRs!SM_Prefix & "' And ID = " & TmpRs2!ID & " "
                                 ElseIf mTotFranchBagAmt <> "0" Then
'                                    iCost = Format(TmpRs1!InsuranceAmt * TmpRs2!FranchInvoiceAmount / Val(mTotFranchBagAmt), "#############.00")
                                    tmp = " Update Txn_InvMonGSLDetail Set FranchInvoiceAmount =  (FranchInvoiceAmount - " & iCost & "),InvoiceDate =  '" & mToDate & "',InsuranceCostDeduction = InsuranceCostDeduction + " & iCost & " "
                                    tmp = tmp & " Where YearMonth = '" & mFromDate & "' And InvoiceAmtForFranch > 0 "
                                    tmp = tmp & " And FranchBillingUnit = 'B' And SM_Prefix = '" & TmpRs!SM_Prefix & "' And ID = " & TmpRs2!ID & " And FranchInvoiceAmount <> 0 "
                                 ElseIf mTotFranchQtyAmt <> "0" Then
'                                    iCost = Format(TmpRs1!InsuranceAmt * TmpRs2!FranchInvoiceAmount / Val(mTotFranchQtyAmt), "#############.00")
                                    tmp = " Update Txn_InvMonGSLDetail Set FranchInvoiceAmount =  (FranchInvoiceAmount - " & iCost & "),InvoiceDate =  '" & mToDate & "',InsuranceCostDeduction = InsuranceCostDeduction + " & iCost & " "
                                    tmp = tmp & " Where YearMonth = '" & mFromDate & "' And InvoiceAmtForFranch > 0 "
                                    tmp = tmp & " And FranchBillingUnit = 'Q' And SM_Prefix = '" & TmpRs!SM_Prefix & "' And ID = " & TmpRs2!ID & " And FranchInvoiceAmount <> 0 "
                                 Else
                                    iCost = 0
                                 End If
                               
                                 Call Tmp3Table
                              
                                 TmpRs2.MoveNext
                             Next
                          End If
                       Else
                          tmp = "Select Sum(InvoiceAmtForFranch) 'TotInvoiceAmt' from Txn_InvMonGSLDetail Where YearMonth = '" & mFromDate & "' And SM_Prefix = '" & TmpRs!SM_Prefix & "' And GodownID = " & TmpRs1!GodownID & " And CommodityID = " & TmpRs1!CommodityID & " And ClientID = " & TmpRs1!ClientID & " And InvoiceAmtForFranch > 0 "
                          Call Tmp5Table
                         
                          mTotInvoiceAmt = IIf(IsNull(TmpRs5!TotInvoiceAmt), 0, TmpRs5!TotInvoiceAmt)
                        
                          tmp = "Select * from Txn_InvMonGSLDetail Where YearMonth = '" & mFromDate & "' And SM_Prefix = '" & TmpRs!SM_Prefix & "' And GodownID = " & TmpRs1!GodownID & " And CommodityID = " & TmpRs1!CommodityID & " And ClientID = " & TmpRs1!ClientID & "  And InvoiceAmtForFranch > 0 "
                          Call Tmp5Table
                        
                          If TmpRs5.RecordCount > 0 Then
                             For pqr = 1 To TmpRs5.RecordCount
                                 
                                 iCost = 0 'Format(TmpRs1!InsuranceAmt * TmpRs5!InvoiceAmtForFranch / Val(mTotInvoiceAmt), "#############.00")
                                 
                                 tmp = " Update Txn_InvMonGSLDetail Set FranchInvoiceAmount =  (InvoiceAmtForFranch - " & iCost & ")* FranchiseeRate /100 ,InvoiceDate =  '" & mToDate & "',InsuranceCostDeduction = InsuranceCostDeduction + " & iCost & " "
                                 tmp = tmp & " Where YearMonth = '" & mFromDate & "' And InvoiceAmtForFranch > 0 "
                                 tmp = tmp & " And FranchBillingUnit = 'P' And SM_Prefix = '" & TmpRs!SM_Prefix & "' And ID = " & TmpRs5!ID & " "
                                  
                                 Call Tmp3Table
                                 
                                 TmpRs5.MoveNext
                             Next
                         End If
                       End If
                       TmpRs1.MoveNext
                   Next
                Else
                   
                   iCost = 0
                  
                   tmp = " Update Txn_InvMonGSLDetail Set FranchInvoiceAmount =  (InvoiceAmtForFranch)* FranchiseeRate /100 ,InvoiceDate =  '" & mToDate & "',InsuranceCostDeduction = 0 "
                   tmp = tmp & " Where YearMonth = '" & mFromDate & "' And InvoiceAmtForFranch > 0 "
                   tmp = tmp & " And FranchBillingUnit = 'P' And SM_Prefix = '" & TmpRs!SM_Prefix & "' And GodownID = " & TmpRs4!GodownID & " And ClientID = " & TmpRs6!ClientID & " And CommodityID = " & TmpRs6!CommodityID & " "
                  
                   Call Tmp3Table
                End If
                TmpRs6.MoveNext
            Next wxy
            TmpRs4.MoveNext
       Next pq
    End If
    
    TmpRs.MoveNext
    ProgressBar1.Value = ProgressBar1.Value + 1
    
Next

If LCase(CmbProcessType.Text) <> "dp" Then

   tmp = "SET IMPLICIT_TRANSACTIONS ON"
   Call TmpTable


   tmp = "Update Txn_InvMonGSLDetail set FranchProcessFlag='Y',FranchProcessdBy='" & Gen_UserLoginID & "',FranchProcessedDate=getdate() Where YearMonth = '" & mFromDate & "' And " & mGenInvoiceLocation
   Call TmpTable
   
End If

ProgressBar1.Value = ProgressBar1.Max
MsgBox "Done!!!"

Call Grid_Head

Exit Sub

msgError:
MsgBox Err.Description, vbExclamation, "Error" '"Some error occurred during process!!!"

End Sub

Private Sub CmdReserDistri_Click()
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

mMonth = mMonthListindex + 1

If mMonth < 10 Then
   mMonth = "0" & mMonth
End If

mFromDate = Format(CmbYear.Text & "-" & mMonth & "-" & "01", "yyyy-mm-dd")
mYear = Val(CmbYear.Text)
mTotalDays = Day(DateSerial(mYear, mMonth + 1, 0))

mToDate = Format(CmbYear.Text & "-" & mMonth & "-" & mTotalDays, "yyyy-mm-dd")

If CDate(Format(mFromDate, "yyyymm")) >= Format(Date, "yyyymm") Then
   MsgBox "Present and Future Month not Allowed!!!"
   CmbMonth.SetFocus
   Exit Sub
End If

tmp = "Select InvoiceNo,ContractID,BillingDate,ActualAmount from Txn_ReservationInvoice Where BillingDate='" & mToDate & "' And Flag in ('A','P')"

Call Tmp1Table

tmp = "Select Mst_GSL_ID,PINO,InvoiceDate,ContractID,isnull(BalanceWeight,0)BalanceWeight,InvoiceAmount,MonG_UID from Txn_InvMonGSLDetail Where YearMonth ='" & mFromDate & "' and isnull(ContractID,0)>0 And PINO is not null"
Call Tmp2Table

If TmpRs1.RecordCount > 0 Then
   ProgressBar1.Max = TmpRs1.RecordCount
   ProgressBar1.Value = 0
   For I = 1 To TmpRs1.RecordCount
       mTotQty = 0
       TmpRs2.Filter = "PINO=" & TmpRs1!InvoiceNo & " And ContractID=" & TmpRs1!ContractID & " And InvoiceDate='" & TmpRs1!BillingDate & "'"
       If TmpRs2.RecordCount > 0 Then
          For Ln = 1 To TmpRs2.RecordCount
              If TmpRs2!BalanceWeight > 0 Then
                 mTotQty = mTotQty + TmpRs2!BalanceWeight
              End If
              TmpRs2.MoveNext
          Next Ln
''          TmpRs2.MoveFirst
''          For Ln = 1 To TmpRs2.RecordCount
''              If mTotQty > 0 Then
''                 If TmpRs2!BalanceWeight > 0 Then
''                    TmpRs2!InvoiceAmount = Format((TmpRs1!ActualAmount * TmpRs2!BalanceWeight) / mTotQty, "#####0.00")
''                    TmpRs2.Update
''                 End If
''              Else
''                 TmpRs2!InvoiceAmount = TmpRs1!ActualAmount
''                 TmpRs2.Update
''                 Exit For
''              End If
''
''              TmpRs2.MoveNext
''          Next Ln
           If mTotQty > 0 Then
              tmp = " Update Txn_InvMonGSLDetail Set InvoiceAmount=" & TmpRs1!ActualAmount & " * BalanceWeight/" & mTotQty & "  Where PINO=" & TmpRs1!InvoiceNo & " And ContractID=" & TmpRs1!ContractID & " And InvoiceDate='" & Format(TmpRs1!BillingDate, Gen_DatFormat) & "' And YearMonth ='" & Format(mFromDate, Gen_DatFormat) & "' And BalanceWeight>0"
              Call TmpTable
           End If
       End If
       TmpRs2.Filter = ""
       TmpRs1.MoveNext
       DoEvents
       ProgressBar1.Value = ProgressBar1.Value + 1
   Next
   
End If

tmp = " Update Txn_InvMonGSLDetail Set FranchInvoiceAmount =  (InvoiceAmount)* FranchiseeRate /100 ,InvoiceDate =  '" & mToDate & "' "
tmp = tmp & " Where YearMonth = '" & mFromDate & "'"
tmp = tmp & " And FranchBillingUnit = 'P' and isnull(ContractID,0)>0 And PINO is not null"

Call TmpTable
MsgBox "Done!!"

End Sub

'Private Sub CmdProcess_Click()
'
'On Error GoTo msgError
'
'ProgressBar1.Value = 0
'
'If CmbMonth.Text = "" Then
'   MsgBox " Blank Month not allowed  !!!"
'   CmbMonth.SetFocus
'   Exit Sub
'End If
'
'If CmbYear.Text = "" Then
'   MsgBox " Blank Year not allowed  !!!"
'   CmbYear.SetFocus
'   Exit Sub
'End If
'
'If MSHFlexGrid1.Rows < 3 Then
'   MsgBox " No Location populated in Grid  !!!"
'   Exit Sub
'End If
'
'If ChkSelection = False Then
'   MsgBox " No Location selected from Grid  !!!"
'   Exit Sub
'End If
'
'mMonth = mMonthListindex + 1
'
'If mMonth < 10 Then
'   mMonth = "0" & mMonth
'End If
'
'
'mFromDate = Format(CmbYear.Text & "-" & mMonth & "-" & "01", "yyyy-mm-dd")
'
'mYear = Val(CmbYear.Text)
'mTotalDays = Day(DateSerial(mYear, mMonth + 1, 0))
'
'mToDate = Format(CmbYear.Text & "-" & mMonth & "-" & mTotalDays, "yyyy-mm-dd")
'
'mNoofdays = ""
'
'If CDate(Format(mFromDate, "yyyymm")) >= Format(Date, "yyyymm") Then
'   MsgBox "Present and Future Month not Allowed !!!"
'   CmbMonth.SetFocus
'   Exit Sub
'End If
'
'CmdProcess.Enabled = False
'
'mGenInvoiceLocation = ""
'
'If MSHFlexGrid1.Rows > 2 Then
'      mGenInvoiceLocation = "SM_Prefix in ("
'      For I = 1 To MSHFlexGrid1.Rows - 1
'          If MSHFlexGrid1.TextMatrix(I, 1) = "" Then Exit For
'          If MSHFlexGrid1.TextMatrix(I, 0) = strChecked Then
'             If mGenInvoiceLocation = "SM_Prefix in (" Then
'                mGenInvoiceLocation = mGenInvoiceLocation & "'" & MSHFlexGrid1.TextMatrix(I, 2) & "'"
'             Else
'                mGenInvoiceLocation = mGenInvoiceLocation & ",'" & MSHFlexGrid1.TextMatrix(I, 2) & "' "
'             End If
'          End If
'      Next
'      mGenInvoiceLocation = mGenInvoiceLocation & " )"
'End If
'
'
'
''Commented temprorly for Test server
'
'If ChkBOD = False Then
'   MsgBox " BOD pending for Location selected from Grid  !!!"
'   CmdProcess.Enabled = True
'   Exit Sub
'End If
'
'
'
''If CmbPreFix.Text <> "" Then
''   tmp = "Select Count(*) as 'TotalRows' from Txn_InvMonGSLDetail Where SM_Prefix = '" & mSM_Prefix & "' And YearMonth = '" & mFromDate & "' And TFlag = 'O' "
'''Else
'''   tmp = "Select Count(*) as 'TotalRows' from Txn_InvMonGSLDetail Where YearMonth = '" & mFromDate & "' And TFlag = 'O' "
''
''  Call TmpTable
''
''  If TmpRs!TotalRows > 0 Then
''     If CmbPreFix.Text <> "" Then
''        MsgBox "Opening balance process for " & CmbPreFix.Text & " " & CmbMonth.Text & "-" & CmbYear.Text & "  month is already executed!!!"
'''     Else
'''        MsgBox "Opening balance process for this month is already executed!!!"
''     End If
''     Exit Sub
''  End If
''
''
''End If
'
'If LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "admin" And LCase(Gen_UserRole) <> "superadmin" And LCase(Gen_UserRole) <> "head cmg" Then
'   If mShowDTD = False Then
'      MsgBox "Please Complete all check points before Processing Opening Balance!!!"
''      CmdProcess.Enabled = True
'      Exit Sub
'   End If
'   If TxtConfirmation.Text = "" Then
'      MsgBox "Blank Confirmation message not allowed!!!"
'      TxtConfirmation.SetFocus
''      CmdProcess.Enabled = True
'      Exit Sub
'   End If
'End If
'
'Gen_mTimeStamp = GetTimeStamp
'
''-- Create Temp Table with record of GSL
'MsgBox "Wait till next message!!!"
'
''Commented on 02 July 2011
'
''tmp = " if exists (select * from dbo.sysobjects where id = object_id(N'dbo.TempGSL') and OBJECTPROPERTY(id, N'IsUserTable') = 1)"
''tmp = tmp & " drop table dbo.TempGSL"
''Call TmpTable
'
''Commented On 12 Nov 2010
''If CmbPreFix.Text <> "" Then
''Call TableTxn_CreateTempGSL(mFromDate, mSM_Prefix, mGenInvoiceLocation)
''Else
''   Call TableTxn_CreateTempGSL(mFromDate, "", mGenInvoiceLocation)
''End If
'
'
'Call TableTxn_CreateTempGSL(mFromDate, mSM_Prefix, mGenInvoiceLocation)
'
''-- Find Distinct location for those location which are having Withdral txn.
'
'tmp = " Select Distinct TCG.SM_Prefix"
'tmp = tmp & " From Txn_CxTF_GSL  TCG"
'tmp = tmp & " Inner Join Txn_CxTF_Details TCD On TCG.CxTFNo = TCD.CxTFNo And TCG.SM_Prefix = TCD.SM_Prefix And TCG.TxnType = TCD.TxnType"
''tmp = tmp & " Inner Join Mst_GSL MGSL On TCG.GSLID = MGSL.GSLID And TCG.SM_Prefix = MGSL.SM_Prefix"
'
''Commented on 12 Nov 2010 Parag
''If CmbPreFix.Text <> "" Then
''   tmp = tmp & " Where TCG.SM_Prefix = '" & mSM_Prefix & "' And TCG.TxnType = 'W' And  TCD.CxTFDate >= '" & mFromDate & "' "
''Else
''   tmp = tmp & " Where TCG.TxnType = 'W' And  TCD.CxTFDate >= '" & mFromDate & "' " 'And  '" & Format(Date, Gen_DatFormat) & "'"
''End If
'
'tmp = tmp & " Where TCG.TxnType = 'W' And  TCD.CxTFDate >= '" & mFromDate & "' " 'And " & mGenInvoiceLocation
'tmp = tmp & " And TCG." & mGenInvoiceLocation
'tmp = tmp & " Order By TCG.SM_Prefix"
'
'Call TmpTable
'
''If TmpRs.RecordCount > 0 Then
'   ProgressBar1.Value = 0
'   ProgressBar1.Max = TmpRs.RecordCount + 1
'   For I = 1 To TmpRs.RecordCount
'
'       '---fetch Withdrawal record for adding back to opening stock. which is > [To Date]
''
''       tmp = "Select TCG.GSLID,TCG.SM_Prefix,TCG.CxTFNo,TCD.CxTFDate,TCG.NoOfBags,TCG.DematWeight,TCG.TxnType"
''       tmp = tmp & " From Txn_CxTF_GSL  TCG"
''       tmp = tmp & " Inner Join Txn_CxTF_Details TCD On TCG.CxTFNo = TCD.CxTFNo And TCG.SM_Prefix = TCD.SM_Prefix And TCG.TxnType = TCD.TxnType"
''       tmp = tmp & " Inner Join Mst_GSL MGSL On TCG.GSLID = MGSL.GSLID And TCG.SM_Prefix = MGSL.SM_Prefix"
''       'tmp = tmp & " Where TCG.TxnType = 'W' And  TCD.CxTFDate Between '2010-04-01' And  '" & Format(Date, Gen_DatFormat) & "' "
''       tmp = tmp & " Where TCG.TxnType = 'W' And  TCD.CxTFDate Between '" & mFromDate & "' And  '" & Format(Date, Gen_DatFormat) & "' "
''       tmp = tmp & " And TCG.SM_Prefix = '" & TmpRs!SM_Prefix & "'"
''       tmp = tmp & " Order By TCG.SM_Prefix,TCG.GSLID,TCD.CxTFDate"
''
'
'       tmp = "Select TCG.GSLID,TCG.SM_Prefix,sum(TCG.NoOfBags)'NoOfBags' ,sum(TCG.DematWeight)'DematWeight' "
'       tmp = tmp & " From Txn_CxTF_GSL  TCG"
'       tmp = tmp & " Inner Join Txn_CxTF_Details TCD On TCG.CxTFNo = TCD.CxTFNo And TCG.SM_Prefix = TCD.SM_Prefix And TCG.TxnType = TCD.TxnType"
'       tmp = tmp & " Inner Join Mst_GSL MGSL On TCG.GSLID = MGSL.GSLID And TCG.SM_Prefix = MGSL.SM_Prefix"
'       'tmp = tmp & " Where TCG.TxnType = 'W' And  TCD.CxTFDate Between '2010-04-01' And  '" & Format(Date, Gen_DatFormat) & "' "
'       tmp = tmp & " Where TCG.TxnType = 'W' And  TCD.CxTFDate > '" & mToDate & "' "
'       tmp = tmp & " And TCG.SM_Prefix = '" & TmpRs!SM_Prefix & "'" 'And Month(MGSL.CDTFDate) <> Month(TCD.CxTFDate) "
'       tmp = tmp & " Group By TCG.SM_Prefix,TCG.GSLID "
'       tmp = tmp & " Order By TCG.SM_Prefix,TCG.GSLID "
'
'
'       Call Tmp1Table
'
'       '--Update temp table with Withdrawal transaction record to get opening stock
'
'
'       Call TableTxn_TempGSL(" Where SM_Prefix = '" & TmpRs!SM_Prefix & "'")
'
'
'       If RsTxn_TempGSL.RecordCount > 0 Then
'          For j = 1 To TmpRs1.RecordCount
'              If RsTxn_TempGSL.RecordCount > 0 Then
'                 RsTxn_TempGSL.MoveFirst
'              End If
'
'              RsTxn_TempGSL.Find "GSLID=" & TmpRs1!GSLID
'              If Not RsTxn_TempGSL.EOF Then
'                 RsTxn_TempGSL!BalanceBags = RsTxn_TempGSL!BalanceBags + TmpRs1!NoofBags
'                 RsTxn_TempGSL!BalanceWeight = RsTxn_TempGSL!BalanceWeight + TmpRs1!DematWeight
'                 RsTxn_TempGSL!CDTFBags = RsTxn_TempGSL!CDTFBags - TmpRs1!NoofBags
'                 RsTxn_TempGSL!CDTFWeight = RsTxn_TempGSL!CDTFWeight - TmpRs1!DematWeight
'                 RsTxn_TempGSL!Flag = "O"
'                 RsTxn_TempGSL.Update
'              End If
'           TmpRs1.MoveNext
'          Next
'       End If
'       TmpRs.MoveNext
'       ProgressBar1.Value = ProgressBar1.Value + 1
'   Next
'
'
'   ProgressBar1.Value = 0
'
'   If TmpRs.RecordCount < 50 Then
'      ProgressBar1.Max = 1000
'   End If
'
'   '-- Delete All data from temp table having Flag='E'
'
'   tmp = "Delete from TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Where Flag = 'E' "
'   Call TmpTable
'
'   ProgressBar1.Value = ProgressBar1.Value + 1
'
'   tmp = "Alter Table TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Add YearMonth DateTime,BillingCycleID decimal(18,0),BillingUnit Varchar(1),Drno decimal(18,0),Rate decimal(18,4), "
'   tmp = tmp & " NoofDays decimal(18,0) ,InvoiceAmount decimal(18,2),TFLag Varchar(1),CxTFNo Decimal(18,0),CxTFDate DateTime,InvoiceDate DateTime,DateOfWithdrawal Datetime  "
'
'   Call TmpTable
'
'
'
'   '-- Insert all Temp data into GSLInvoice table for billing
'
''   tmp = " Insert into Txn_InvMonGSLDetail "
''   tmp = tmp & " (GSLID,SM_Prefix,CMPID,CommodityID,GodownID,StackNo,LotNo,BalanceBags,"
''   tmp = tmp & " BalanceWeight,ExchangeTypeID,ClientIDRow,Flag,Flag1,Tflag,GrossWt,KRDate,"
''   tmp = tmp & " ContractID,BagSize,YearMonth,G_UID,CreatedBy,CreatedDate,"
''   tmp = tmp & " UpdatedDate,UpdatedBy,CDTFDate,CDTFNo)"
''   tmp = tmp & " Select GSLID,SM_Prefix,CMPID,CommodityID,GodownID,StackNo,LotNo,BalanceBags,"
''   tmp = tmp & " BalanceWeight,ExchangeTypeID,ClientIDRow,Flag,Flag1,'O',GrossWt,KRDate,"
''   tmp = tmp & " NContractID,BagSize,'" & mFromDate & "',G_UID,CreatedBy,CreatedDate,"
''   tmp = tmp & " UpdatedDate , UpdatedBy,CDTFDate,CDTFNo"
''   tmp = tmp & " From TempGSL"
''
''   Call TmpTable
'
'   If Gen_DBType = "MDB" Then
'      tmp = "Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TFlag = 'O',YearMonth = #" & mFromDate & "#"
'   Else
'      tmp = "Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TFlag = 'O',YearMonth = '" & mFromDate & "'"
'   End If
'   Call TmpTable
'   ProgressBar1.Value = ProgressBar1.Value + 1
'
''   tmp = "Drop Table TempGSL"
''   Call TmpTable
'
'
'   'YearMonth,BillingCycleID,BillingUnit,Drno,Rate,NoofDays,InvoiceAmount
'
'   'ProgressBar1.Value = ProgressBar1.Value + 1
'   ' Copy GSL from transaction table for selected month for generation of billing
'
''   tmp = " Insert into Txn_InvMonGSLDetail (GSLID,SM_Prefix,TFlag,CxTFNo,CxTFDate,CMPID,"
'
'   tmp = " Insert into TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " (GSLID,SM_Prefix,TFlag,CxTFNo,CxTFDate,CMPID,"
'   tmp = tmp & " GodownID,StackNo,LotNo,CommodityID,ExchangeTypeID,"
'   tmp = tmp & " ClientIDRow,BalanceBags,BalanceWeight,Flag,Flag1,GrossWt,"
'   tmp = tmp & " BagSize,CDTFNo,CDTFDate,KRDate,NContractID,YearMonth,G_UID,CreatedBy,CreatedDate,"
'   tmp = tmp & " UpdatedDate , UpdatedBy,KRNo)"
'   tmp = tmp & " Select TCG.GSLID,TCG.SM_Prefix,TCG.TxnType,TCG.CxTFNo,TCD.CxTFDate,TCD.CMPID,MGSL.GodownID,MGSL.StackNo,"
'   tmp = tmp & " MGSL.LotNo,MGSL.CommodityID,MGSL.ExchangeTypeID,"
'   tmp = tmp & " MGSL.ClientIDRow,TCG.NoOfBags,TCG.DematWeight,MGSL.Flag,MGSL.Flag1,MGSL.GrossWt,"
'   tmp = tmp & " MGSL.BagSize,MGSL.CDTFNo,MGSL.CDTFDate,MGSL.KRDate,MGSL.NContractID,'" & mFromDate & "',MGSL.G_UID,MGSL.CreatedBy,MGSL.CreatedDate,"
'   tmp = tmp & " MGSL.UpdatedDate , MGSL.UpdatedBy,KRNo"
'   tmp = tmp & " From Txn_CxTF_GSL  TCG"
'   tmp = tmp & " Inner Join Txn_CxTF_Details TCD On TCG.CxTFNo = TCD.CxTFNo And TCG.SM_Prefix = TCD.SM_Prefix"
'   tmp = tmp & " And TCG.TxnType = TCD.TxnType"
'   tmp = tmp & " Inner Join Mst_GSL MGSL On TCG.GSLID = MGSL.GSLID And TCG.SM_Prefix = MGSL.SM_Prefix"
'   tmp = tmp & " Inner Join Mst_CMP MCMP On TCD.CMPID = MCMP.CMPID"
'   tmp = tmp & " Where MGSL.SpillageFlag = 0 And (TCG.BillFlag is Null or TCG.BillFlag = '')"
'   tmp = tmp & " and TCD.CxTFDate Between '" & mFromDate & "' And '" & mToDate & "' "
'
''Commented on 12 Nov 2010 Parag
''   If CmbPreFix.Text <> "" Then
''      tmp = tmp & " and TCG.SM_Prefix = '" & mSM_Prefix & "'"
''   End If
'
'   tmp = tmp & " And TCG." & mGenInvoiceLocation
'
'   Call TmpTable
'
''   Commented on 22 March 2011 order given by Mayur Sir
''   tmp = "Select SM_Prefix,CxTFNo,TFlag,CDTFNo,CDTFDate,CxTFDate From TempGSL "
''   tmp = tmp & " Where YearMonth = '" & mFromDate & "' And TFlag = 'W' And ExchangeTypeID = 1 "
''   tmp = tmp & " And (ContractID = 0 Or ContractID is Null) And DateDiff(Month,CDTFDate,CxTFDate) = 0 And "
''   tmp = tmp & " (SpillageFlag = 0 Or SpillageFlag is Null) And " & mGenInvoiceLocation & " "
''   tmp = tmp & " Group by SM_Prefix,CxTFNo,TFlag,CDTFNo,CDTFDate,CxTFDate"
''
''   Call TmpTable
''
''
''   For I = 1 To TmpRs.RecordCount
''
''       tmp = " Update TempGSL Set DateofWithdrawal = '" & TmpRs!CxtfDate & "'"
''       tmp = tmp & " Where CxTFNo = " & TmpRs!CdtfNo & "  And SM_Prefix = '" & TmpRs!SM_Prefix & "' And "
''       tmp = tmp & " YearMonth = '" & mFromDate & "' And TFlag = 'D' And ExchangeTypeID = 1 "
''       tmp = tmp & " And (ContractID = 0 Or ContractID is Null) And (SpillageFlag = 0 Or SpillageFlag is Null) "
''
''       Call Tmp1Table
''
''       TmpRs.MoveNext
''   Next
'
'   'Implemented on 03 Nov 2010
'   'Due to decimal place defination difference in GSL Master table and  GSL Invoice Detail Table
'
'   tmp = "Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set BagSize = Round(Bagsize,0) "
'   Call TmpTable
'
'   ProgressBar1.Value = ProgressBar1.Value + 1
'
'   ' Flag off all GSL which are used for selected month in transaction table so to avoid duplicate billing
'
'   tmp = " Update Txn_CxTF_GSL Set Txn_CxTF_GSL.BillFlag = 'I'"
'   tmp = tmp & " Where  GSLDetailID in (Select TCG.GSLDetailID"
'   tmp = tmp & " From Txn_CxTF_GSL  TCG"
'   tmp = tmp & " Inner Join Txn_CxTF_Details TCD On TCG.CxTFNo = TCD.CxTFNo And TCG.SM_Prefix = TCD.SM_Prefix"
'   tmp = tmp & " And TCG.TxnType = TCD.TxnType"
'   tmp = tmp & " where  (TCG.BillFlag is Null or TCG.BillFlag = '')  And"
'   tmp = tmp & " TCD.CxTFDate Between '" & mFromDate & "' And '" & mToDate & "' And Txn_CxTF_GSL.SM_Prefix = TCG.SM_Prefix) "
'
'   'tmp = tmp & " TCD.CxTFDate <= '" & mToDate & "' And Txn_CxTF_GSL.SM_Prefix = TCG.SM_Prefix)"
'
''Commented on 12 Nov 2010 Parag
''   If CmbPreFix.Text <> "" Then
''      tmp = tmp & " and Txn_CxTF_GSL.SM_Prefix = '" & mSM_Prefix & "'"
''   End If
'
'   tmp = tmp & " And Txn_CxTF_GSL." & mGenInvoiceLocation
'
'   Call TmpTable
'
'   ProgressBar1.Value = ProgressBar1.Value + 1
'
'   '--------------------------------------------------
'   '| A. Apply Special Rate to GSL Table.                |
'   '--------------------------------------------------
'
''    tmp = " Update Txn_InvMonGslDetail Set Txn_InvMonGslDetail.BillingCycleID = "
''    tmp = tmp & " (Select a.BillingCycleID"
''    tmp = tmp & " from  Txn_InvMonGslDetail b"
''    tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
''    tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
''    tmp = tmp & " left join Mst_SpecialRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid "
''    tmp = tmp & " And a.ClientIDRow = b.ClientIDRow And a.commoditygroupid = c.commoditygroupid"
''    tmp = tmp & " where (b.ContractID = 0 Or b.ContractID is Null) And b.YearMonth = '" & mFromDate & "'  And a.SPRCID is not null And b.GSLID = Txn_InvMonGslDetail.GSLID"
''    tmp = tmp & " And b.SM_Prefix = Txn_InvMonGslDetail.SM_Prefix"
''    tmp = tmp & " Group by b.GSLID,b.SM_Prefix,"
''    tmp = tmp & " a.SPRCID, a.BillingCycleID,a.BillingUnit,"
''    tmp = tmp & " a.RateOnQty,a.BaseRate,a.AbsoluteDiscount,a.Amount) "
''    tmp = tmp & " Where Txn_InvMonGslDetail.YearMonth = '" & mFromDate & "' And ExchangeTypeID = 1"
''
''    If CmbPreFix.Text <> "" Then
''       tmp = tmp & " and Txn_InvMonGslDetail.SM_Prefix = '" & mSM_Prefix & "'"
''    End If
'
'    ' For Op and With
'
'    tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingCycleID = "
'    tmp = tmp & " (Select a.BillingCycleID"
'    tmp = tmp & " from  TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " b"
'    tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
'    tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
'    tmp = tmp & " left join Mst_SpecialRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid "
'    tmp = tmp & " And a.ClientIDRow = b.ClientIDRow And a.commoditygroupid = c.commoditygroupid"
'    tmp = tmp & " where (b.ContractID = 0 Or b.ContractID is Null) And b.YearMonth = '" & mFromDate & "'  And a.SPRCID is not null And a.FLag = 'A' And b.GSLID = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".GSLID"
'    tmp = tmp & " And b.SM_Prefix = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix"
'    tmp = tmp & " Group by b.GSLID,b.SM_Prefix,"
'    tmp = tmp & " a.SPRCID, a.BillingCycleID,a.BillingUnit,"
'    tmp = tmp & " a.RateOnQty,a.BaseRate,a.AbsoluteDiscount,a.Amount) "
'    tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1 And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag in ('O','W')"
'
'    If CmbPreFix.Text <> "" Then
'       tmp = tmp & " and TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix = '" & mSM_Prefix & "'"
'    End If
'
'    Call TmpTable
'
'    ProgressBar1.Value = ProgressBar1.Value + 1
'
'    'For Depo
'    tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingCycleID = "
'    tmp = tmp & " (Select a.DepoBillingCycleID"
'    tmp = tmp & " from  TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " b"
'    tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
'    tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
'    tmp = tmp & " left join Mst_SpecialRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid "
'    tmp = tmp & " And a.ClientIDRow = b.ClientIDRow And a.commoditygroupid = c.commoditygroupid"
'    tmp = tmp & " where (b.ContractID = 0 Or b.ContractID is Null) And b.YearMonth = '" & mFromDate & "'  And a.SPRCID is not null And a.FLag = 'A' And b.GSLID = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".GSLID"
'    tmp = tmp & " And b.SM_Prefix = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix"
'    tmp = tmp & " Group by b.GSLID,b.SM_Prefix,"
'    tmp = tmp & " a.SPRCID, a.DepoBillingCycleID,a.BillingUnit,"
'    tmp = tmp & " a.RateOnQty,a.BaseRate,a.AbsoluteDiscount,a.Amount) "
'    tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1  And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag in ('D')"
'
'    If CmbPreFix.Text <> "" Then
'       tmp = tmp & " and TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix = '" & mSM_Prefix & "'"
'    End If
'
'
'    Call TmpTable
'
'
'
'    ProgressBar1.Value = ProgressBar1.Value + 1
'
''    tmp = " Update Txn_InvMonGslDetail Set Txn_InvMonGslDetail.BillingUnit = "
''    tmp = tmp & " (Select a.BillingUnit"
''    tmp = tmp & " from  Txn_InvMonGslDetail b"
''    tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
''    tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
''    tmp = tmp & " left join Mst_SpecialRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid "
''    tmp = tmp & " And a.ClientIDRow = b.ClientIDRow And a.commoditygroupid = c.commoditygroupid"
''    tmp = tmp & " where (b.ContractID = 0 Or b.ContractID is Null) And b.YearMonth = '" & mFromDate & "'  And a.SPRCID is not null And b.GSLID = Txn_InvMonGslDetail.GSLID"
''    tmp = tmp & " And b.SM_Prefix = Txn_InvMonGslDetail.SM_Prefix"
''    tmp = tmp & " Group by b.GSLID,b.SM_Prefix,"
''    tmp = tmp & " a.SPRCID, a.BillingCycleID,a.BillingUnit,"
''    tmp = tmp & " a.RateOnQty,a.BaseRate,a.AbsoluteDiscount,a.Amount)"
''    tmp = tmp & " Where Txn_InvMonGslDetail.YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1"
''
''    If CmbPreFix.Text <> "" Then
''       tmp = tmp & " and Txn_InvMonGslDetail.SM_Prefix = '" & mSM_Prefix & "'"
''    End If
''
''    Call TmpTable
'
'
'    tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingUnit = "
'    tmp = tmp & " (Select a.BillingUnit"
'    tmp = tmp & " from  TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " b"
'    tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
'    tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
'    tmp = tmp & " left join Mst_SpecialRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid "
'    tmp = tmp & " And a.ClientIDRow = b.ClientIDRow And a.commoditygroupid = c.commoditygroupid"
'    tmp = tmp & " where (b.ContractID = 0 Or b.ContractID is Null) And b.YearMonth = '" & mFromDate & "'  And a.SPRCID is not null And a.FLag = 'A' And b.GSLID = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".GSLID"
'    tmp = tmp & " And b.SM_Prefix = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix"
'    tmp = tmp & " Group by b.GSLID,b.SM_Prefix,"
'    tmp = tmp & " a.SPRCID, a.BillingCycleID,a.BillingUnit,"
'    tmp = tmp & " a.RateOnQty,a.BaseRate,a.AbsoluteDiscount,a.Amount)"
'    tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "'  And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1"
'
'    If CmbPreFix.Text <> "" Then
'       tmp = tmp & " and TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix = '" & mSM_Prefix & "'"
'    End If
'
'    Call TmpTable
'
'
'    ProgressBar1.Value = ProgressBar1.Value + 1
'
''    tmp = " Update Txn_InvMonGslDetail Set Txn_InvMonGslDetail.Drno = "
''    tmp = tmp & " (Select a.SPRCID"
''    tmp = tmp & " from  Txn_InvMonGslDetail b"
''    tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
''    tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
''    tmp = tmp & " left join Mst_SpecialRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid "
''    tmp = tmp & " And a.ClientIDRow = b.ClientIDRow And a.commoditygroupid = c.commoditygroupid"
''    tmp = tmp & " where (b.ContractID = 0 Or b.ContractID is Null) And b.YearMonth = '" & mFromDate & "'  And a.SPRCID is not null And b.GSLID = Txn_InvMonGslDetail.GSLID"
''    tmp = tmp & " And b.SM_Prefix = Txn_InvMonGslDetail.SM_Prefix"
''    tmp = tmp & " Group by b.GSLID,b.SM_Prefix,"
''    tmp = tmp & " a.SPRCID, a.BillingCycleID,a.BillingUnit,"
''    tmp = tmp & " a.RateOnQty,a.BaseRate,a.AbsoluteDiscount,a.Amount)"
''    tmp = tmp & " Where Txn_InvMonGslDetail.YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1"
''
''    If CmbPreFix.Text <> "" Then
''       tmp = tmp & " and Txn_InvMonGslDetail.SM_Prefix = '" & mSM_Prefix & "'"
''    End If
''
''    Call TmpTable
''
''
''    ProgressBar1.Value = ProgressBar1.Value + 1
''
''    tmp = " Update Txn_InvMonGslDetail Set Txn_InvMonGslDetail.Rate = "
''    tmp = tmp & " (Select a.Amount"
''    tmp = tmp & " from  Txn_InvMonGslDetail b"
''    tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
''    tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
''    tmp = tmp & " left join Mst_SpecialRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid "
''    tmp = tmp & " And a.ClientIDRow = b.ClientIDRow And a.commoditygroupid = c.commoditygroupid"
''    tmp = tmp & " where (b.ContractID = 0 Or b.ContractID is Null) And b.YearMonth = '" & mFromDate & "'  And a.SPRCID is not null And b.GSLID = Txn_InvMonGslDetail.GSLID"
''    tmp = tmp & " And b.SM_Prefix = Txn_InvMonGslDetail.SM_Prefix"
''    tmp = tmp & " Group by b.GSLID,b.SM_Prefix,"
''    tmp = tmp & " a.SPRCID, a.BillingCycleID,a.BillingUnit,"
''    tmp = tmp & " a.RateOnQty,a.BaseRate,a.AbsoluteDiscount,a.Amount)"
''    tmp = tmp & " Where Txn_InvMonGslDetail.YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1"
''
''    If CmbPreFix.Text <> "" Then
''       tmp = tmp & " and Txn_InvMonGslDetail.SM_Prefix = '" & mSM_Prefix & "'"
''    End If
''
''    Call TmpTable
''
''
''    ProgressBar1.Value = ProgressBar1.Value + 1
''
''    tmp = " Update Txn_InvMonGslDetail Set Flag1 = Null Where YearMonth = '" & mFromDate & "'"
''
''    If CmbPreFix.Text <> "" Then
''       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
''    End If
''
''    Call TmpTable
''
''
''    ProgressBar1.Value = ProgressBar1.Value + 1
''
''    tmp = " Update Txn_InvMonGslDetail Set Flag1 = 'S' Where DRNo is not null And YearMonth = '" & mFromDate & "' "
''
''    If CmbPreFix.Text <> "" Then
''       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
''    End If
''
''    Call TmpTable
''
''
''    ProgressBar1.Value = ProgressBar1.Value + 1
''
''
''    tmp = " Update Txn_InvMonGslDetail Set Txn_InvMonGslDetail.Rate ="
''    tmp = tmp & " (Select  MSRCD.RatePerBag"
''    tmp = tmp & " From Txn_InvMonGSLDetail TIMGD"
''    tmp = tmp & " Inner Join Mst_SpecialRateCardDetail MSRCD On TIMGD.DRNo = MSRCD.SPRCID"
''    tmp = tmp & " Where TIMGD.YearMonth = '" & mFromDate & "' And TIMGD.DRNo is Not null And TIMGD.BagSize >= MSRCD.FromBagSize And"
''    tmp = tmp & " TIMGD.BagSize <= MSRCD.ToBagSize And TIMGD.GSLID = Txn_InvMonGslDetail.GSLID"
''    tmp = tmp & " And TIMGD.SM_Prefix = Txn_InvMonGslDetail.SM_Prefix And TIMGD.BillingUnit = 'B'"
''    tmp = tmp & " Group By TIMGD.GSLID,TIMGD.SM_Prefix,TIMGD.BagSize,MSRCD.FromBagSize,MSRCD.ToBagSize,MSRCD.RatePerBag)"
''    tmp = tmp & " Where Txn_InvMonGslDetail.BillingUnit = 'B' And Txn_InvMonGslDetail.Flag1 = 'S' "
''    tmp = tmp & " And Txn_InvMonGslDetail.YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1"
''
''    If CmbPreFix.Text <> "" Then
''       tmp = tmp & " and Txn_InvMonGslDetail.SM_Prefix = '" & mSM_Prefix & "'"
''    End If
''
''
''    Call TmpTable
''
''
''    ProgressBar1.Value = ProgressBar1.Value + 1
''   '--------------------------------------------------
''   '| B. Apply  Base Rate to GSL Table.                |
''   '--------------------------------------------------
''
''    tmp = " Update Txn_InvMonGslDetail Set Txn_InvMonGslDetail.BillingCycleID ="
''    tmp = tmp & " (Select a.BillingCycleID"
''    tmp = tmp & " from  Txn_InvMonGslDetail b"
''    tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
''    tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
''    tmp = tmp & " left join Mst_StorageRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid And a.commoditygroupid = c.commoditygroupid"
''    tmp = tmp & " Where a.SRCID is not null And (b.ContractID = 0 Or b.ContractID is Null) And b.YearMonth = '" & mFromDate & "' And a.ExchangeTypeID = 1 And a.DefaultRate = 1 And b.Flag1 is null "
''    tmp = tmp & " And b.GSLID = Txn_InvMonGslDetail.GSLID And b.SM_Prefix = Txn_InvMonGslDetail.SM_Prefix"
''    tmp = tmp & " Group by b.SM_Prefix,b.GSLID,b.ExchangeTypeID,"
''    tmp = tmp & " a.SRCID, a.BillingCycleID,a.Flag,a.RateOnQty)"
''    tmp = tmp & " Where (Txn_InvMonGslDetail.Flag1 is Null)  "
''    tmp = tmp & " And Txn_InvMonGslDetail.YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 "
''
''    If CmbPreFix.Text <> "" Then
''       tmp = tmp & " and Txn_InvMonGslDetail.SM_Prefix = '" & mSM_Prefix & "'"
''    End If
''
''    Call TmpTable
''
''
''    ProgressBar1.Value = ProgressBar1.Value + 1
''
''    tmp = " Update Txn_InvMonGslDetail Set Txn_InvMonGslDetail.BillingUnit ="
''    tmp = tmp & " (Select a.Flag"
''    tmp = tmp & " from  Txn_InvMonGslDetail b"
''    tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
''    tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
''    tmp = tmp & " left join Mst_StorageRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid And a.commoditygroupid = c.commoditygroupid"
''    tmp = tmp & " Where a.SRCID is not null And (b.ContractID = 0 Or b.ContractID is Null) And b.YearMonth = '" & mFromDate & "' And a.ExchangeTypeID = 1 And a.DefaultRate = 1 And b.Flag1 is null"
''    tmp = tmp & " And b.GSLID = Txn_InvMonGslDetail.GSLID And b.SM_Prefix = Txn_InvMonGslDetail.SM_Prefix"
''    tmp = tmp & " Group by b.SM_Prefix,b.GSLID,b.ExchangeTypeID,"
''    tmp = tmp & " a.SRCID, a.BillingCycleID,a.Flag,a.RateOnQty)"
''    tmp = tmp & " Where (Txn_InvMonGslDetail.Flag1 is Null) "
''    tmp = tmp & " And Txn_InvMonGslDetail.YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 "
''
''    If CmbPreFix.Text <> "" Then
''       tmp = tmp & " and Txn_InvMonGslDetail.SM_Prefix = '" & mSM_Prefix & "'"
''    End If
''
''    Call TmpTable
''
''
''    ProgressBar1.Value = ProgressBar1.Value + 1
''
''    tmp = " Update Txn_InvMonGslDetail Set Txn_InvMonGslDetail.DRNo ="
''    tmp = tmp & " (Select a.SRCID"
''    tmp = tmp & " from  Txn_InvMonGslDetail b"
''    tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
''    tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
''    tmp = tmp & " left join Mst_StorageRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid And a.commoditygroupid = c.commoditygroupid"
''    tmp = tmp & " Where a.SRCID is not null And (b.ContractID = 0 Or b.ContractID is Null) And b.YearMonth = '" & mFromDate & "' And a.ExchangeTypeID = 1 And a.DefaultRate = 1 And b.Flag1 is null"
''    tmp = tmp & " And b.GSLID = Txn_InvMonGslDetail.GSLID And b.SM_Prefix = Txn_InvMonGslDetail.SM_Prefix"
''    tmp = tmp & " Group by b.SM_Prefix,b.GSLID,b.ExchangeTypeID,"
''    tmp = tmp & " a.SRCID, a.BillingCycleID,a.Flag,a.RateOnQty)"
''    tmp = tmp & " Where (Txn_InvMonGslDetail.Flag1 is Null) "
''    tmp = tmp & " And Txn_InvMonGslDetail.YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 "
''
''
''    If CmbPreFix.Text <> "" Then
''       tmp = tmp & " and Txn_InvMonGslDetail.SM_Prefix = '" & mSM_Prefix & "'"
''    End If
''
''
''    Call TmpTable
''
''
''    ProgressBar1.Value = ProgressBar1.Value + 1
''
''    tmp = " Update Txn_InvMonGslDetail Set Txn_InvMonGslDetail.Rate ="
''    tmp = tmp & " (Select a.RateOnQty"
''    tmp = tmp & " from  Txn_InvMonGslDetail b"
''    tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
''    tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
''    tmp = tmp & " left join Mst_StorageRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid And a.commoditygroupid = c.commoditygroupid"
''    tmp = tmp & " Where a.SRCID is not null And (b.ContractID = 0 Or b.ContractID is Null) And b.YearMonth = '" & mFromDate & "' And a.ExchangeTypeID = 1 And a.DefaultRate = 1 And b.Flag1 is null"
''    tmp = tmp & " And b.GSLID = Txn_InvMonGslDetail.GSLID And b.SM_Prefix = Txn_InvMonGslDetail.SM_Prefix"
''    tmp = tmp & " Group by b.SM_Prefix,b.GSLID,b.ExchangeTypeID,"
''    tmp = tmp & " a.SRCID, a.BillingCycleID,a.Flag,a.RateOnQty)"
''    tmp = tmp & " Where (Txn_InvMonGslDetail.Flag1 is Null) "
''    tmp = tmp & " And Txn_InvMonGslDetail.YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 "
''
''    If CmbPreFix.Text <> "" Then
''       tmp = tmp & " and Txn_InvMonGslDetail.SM_Prefix = '" & mSM_Prefix & "'"
''    End If
''
''    Call TmpTable
''
''
''
''    ProgressBar1.Value = ProgressBar1.Value + 1
''
''    tmp = " Update Txn_InvMonGslDetail Set Flag1 = 'N' Where DRNo is not null And YearMonth = '" & mFromDate & "' And Flag1 is null And BillingCycleID is not Null And ExchangeTypeID = 1 "
''
''    If CmbPreFix.Text <> "" Then
''       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
''    End If
''
''    Call TmpTable
''
''
''    ProgressBar1.Value = ProgressBar1.Value + 1
''
''    tmp = " Update Txn_InvMonGslDetail Set Txn_InvMonGslDetail.Rate ="
''    tmp = tmp & " (Select MSRCD.RatePerBag"
''    tmp = tmp & " From Txn_InvMonGSLDetail TIMGD"
''    tmp = tmp & " Inner Join Mst_StorageRateCardDetail MSRCD On TIMGD.DRNo = MSRCD.SRCID"
''    tmp = tmp & " Inner Join Mst_StorageRateCard MSRC On MSRCD.SRCID = MSRC.SRCID"
''    tmp = tmp & " Where TIMGD.YearMonth = '" & mFromDate & "' And TIMGD.DRNo is Not null And TIMGD.BagSize >= MSRCD.FromBagSize And"
''    tmp = tmp & " TIMGD.BagSize <= MSRCD.ToBagSize  And TIMGD.BillingUnit = 'B' And TIMGD.Flag1 = 'N' And  MSRC.DefaultRate = 1"
''    tmp = tmp & " And TIMGD.GSLID = Txn_InvMonGslDetail.GSLID And TIMGD.SM_Prefix = Txn_InvMonGslDetail.SM_Prefix"
''    tmp = tmp & " And TIMGD.DRNo = Txn_InvMonGslDetail.DRNo"
''    tmp = tmp & " Group By TIMGD.GSLID,TIMGD.SM_Prefix,TIMGD.BagSize,MSRCD.FromBagSize,MSRCD.ToBagSize,MSRCD.RatePerBag)"
''    tmp = tmp & " Where Txn_InvMonGslDetail.BillingUnit = 'B' And Txn_InvMonGslDetail.Flag1 = 'N'"
''    tmp = tmp & " And Txn_InvMonGslDetail.YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 "
''
''    If CmbPreFix.Text <> "" Then
''       tmp = tmp & " and Txn_InvMonGslDetail.SM_Prefix = '" & mSM_Prefix & "'"
''    End If
''
''    Call TmpTable
''
''
''
''    ProgressBar1.Value = ProgressBar1.Value + 1
''
''   '--------------------------------------------------
''   '| C. Apply  No Of Cycles to GSL Table as per Billing Unit and Transaction Flag.                |
''   '--------------------------------------------------
''
''    'Op Bal and Monthly Cycle
''    tmp = " Update Txn_InvMonGslDetail Set NoofDays = 1 Where YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TFlag = 'O'"
''    tmp = tmp & " And BillingCycleID = 4  And (ContractID = 0 Or ContractID is Null)"
''
''    If CmbPreFix.Text <> "" Then
''       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
''    End If
''
''    Call TmpTable
''
''
''
''    ProgressBar1.Value = ProgressBar1.Value + 1
''
''    'Op Bal and Forthnightly Cycle
''    tmp = " Update Txn_InvMonGslDetail Set NoofDays = 2 Where YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TFlag = 'O'"
''    tmp = tmp & " And BillingCycleID = 3 And (ContractID = 0 Or ContractID is Null)"
''
''    If CmbPreFix.Text <> "" Then
''       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
''    End If
''
''
''    Call TmpTable
''
''
''
''    ProgressBar1.Value = ProgressBar1.Value + 1
''
''    'Op Bal and Weekly Cycle
''    tmp = " Update Txn_InvMonGslDetail Set NoofDays = 4 Where YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TFlag = 'O'"
''    tmp = tmp & " And BillingCycleID = 2 And (ContractID = 0 Or ContractID is Null)"
''
''    If CmbPreFix.Text <> "" Then
''       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
''    End If
''
''
''    Call TmpTable
''
''
''
''    ProgressBar1.Value = ProgressBar1.Value + 1
''    'Op Bal and Daily Cycle
''    tmp = " Update Txn_InvMonGslDetail Set NoofDays = " & mTotalDays & " Where YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TFlag = 'O'"
''    tmp = tmp & " And BillingCycleID = 1 And (ContractID = 0 Or ContractID is Null)"
''
''    If CmbPreFix.Text <> "" Then
''       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
''    End If
''
''
''    Call TmpTable
''
''    ProgressBar1.Value = ProgressBar1.Value + 1
''
''    'Depo and Withd and Monthly Cycle
''    tmp = " Update Txn_InvMonGslDetail Set NoofDays = 1 Where YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TFlag <> 'O'"
''    tmp = tmp & " And BillingCycleID = 4 And (ContractID = 0 Or ContractID is Null)"
''
''    If CmbPreFix.Text <> "" Then
''       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
''    End If
''
''
''    Call TmpTable
''
''    ProgressBar1.Value = ProgressBar1.Value + 1
''
''    'Depo and Withd and Daily Cycle
''    tmp = "Update Txn_InvMonGslDetail Set NoOfDays = 30-Day(CxTFDate) + 1 Where YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TFlag = 'D' And BillingCycleID = 1 And (ContractID = 0 Or ContractID is Null)"
''
''    If CmbPreFix.Text <> "" Then
''       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
''    End If
''
''    Call TmpTable
''
''
''
''    ProgressBar1.Value = ProgressBar1.Value + 1
''
''    tmp = "Update Txn_InvMonGslDetail Set NoOfDays = Day(CxTFDate) Where YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TFlag = 'W' And BillingCycleID = 1 And (ContractID = 0 Or ContractID is Null)"
''
''    If CmbPreFix.Text <> "" Then
''       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
''    End If
''
''    Call TmpTable
''
''
''
''    ProgressBar1.Value = ProgressBar1.Value + 1
''
''    'Deposit & Weekly
''    mNoofdays = GetBillingCycleDays(2)
''
'''    tmp = "Update Txn_InvMonGslDetail Set NoOfDays = Floor((4 - Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ") + 1)"
'''    tmp = tmp & " Where Txn_InvMonGslDetail.YearMonth = '" & mFromDate & "' And Txn_InvMonGslDetail.ExchangeTypeID = 1"
'''    tmp = tmp & " And Txn_InvMonGslDetail.TFlag = 'D' And Txn_InvMonGslDetail.BillingCycleID = 2 And (ContractID = 0 Or ContractID is Null)"
'''
'''    Call TmpTable
''
''
'''    tmp = "Update Txn_InvMonGslDetail Set NoOfDays = 1 Where YearMonth = '" & mFromDate & "' And ExchangeTypeID = 1 And TFlag = 'D' And BillingCycleID in (2) And"
'''    tmp = tmp & " Floor ((4 - (Convert(Numeric, Day(CxTFDate)) / 7) + 1)) < 1"
'''
'''    Call TmpTable
''
''
''    tmp = "Update Txn_InvMonGslDetail Set NoOfDays =  Case Floor((4 - (Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ") + 1))  When 0 then 1 else Floor((4 - (Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ") + 1)) End "
''    tmp = tmp & " Where Txn_InvMonGslDetail.YearMonth = '" & mFromDate & "' And Txn_InvMonGslDetail.ExchangeTypeID = 1"
''    tmp = tmp & " And Txn_InvMonGslDetail.TFlag = 'D' And Txn_InvMonGslDetail.BillingCycleID = 2 And (ContractID = 0 Or ContractID is Null)"
''
''    If CmbPreFix.Text <> "" Then
''       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
''    End If
''
''
''    Call TmpTable
''
''
''    ProgressBar1.Value = ProgressBar1.Value + 1
''
''    'Deposit & FortNightly
''    mNoofdays = GetBillingCycleDays(3)
''
''
'''    tmp = "Update Txn_InvMonGslDetail Set NoOfDays = Floor(((4 - Round(Convert(Decimal,Day(CxTFDate))/" & mNoofdays & ",1)) - 1))"
'''    tmp = tmp & " Where Txn_InvMonGslDetail.YearMonth = '" & mFromDate & "' And Txn_InvMonGslDetail.ExchangeTypeID = 1"
'''    tmp = tmp & " And Txn_InvMonGslDetail.TFlag = 'D' And Txn_InvMonGslDetail.BillingCycleID = 3 And (ContractID = 0 Or ContractID is Null)"
'''
'''    Call TmpTable
''
''
'''    tmp = "Update Txn_InvMonGslDetail Set NoOfDays = Case Floor((4 - (Convert(Numeric,Day(CxTFDate))/7) + 1))  When 0 then 1 else Floor((4 - (Convert(Numeric,Day(CxTFDate))/7) + 1)) End "
'''    tmp = tmp & " Where Txn_InvMonGslDetail.YearMonth = '" & mFromDate & "' And Txn_InvMonGslDetail.ExchangeTypeID = 1"
'''    tmp = tmp & " And Txn_InvMonGslDetail.TFlag = 'D' And Txn_InvMonGslDetail.BillingCycleID = 3 And (ContractID = 0 Or ContractID is Null)"
''
''
''    tmp = " Update Txn_InvMonGslDetail Set NoOfDays = Case Floor((4 - Round(Convert(Decimal,Day(CxTFDate))/" & mNoofdays & ",1)) - 1) When 0 then 1 else Floor((4 - Round(Convert(Decimal,Day(CxTFDate))/" & mNoofdays & ",1)) - 1) End "
''    tmp = tmp & " Where Txn_InvMonGslDetail.YearMonth = '" & mFromDate & "' And Txn_InvMonGslDetail.ExchangeTypeID = 1 "
''    tmp = tmp & " And Txn_InvMonGslDetail.TFlag = 'D' And Txn_InvMonGslDetail.BillingCycleID = 3 And (ContractID = 0 Or ContractID is Null) "
''
''    If CmbPreFix.Text <> "" Then
''       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
''    End If
''
''
''    Call TmpTable
''
''
''    ProgressBar1.Value = ProgressBar1.Value + 1
''    'Withdrawal & weekly
''    mNoofdays = GetBillingCycleDays(2)
''
''    tmp = "Update Txn_InvMonGslDetail Set NoOfDays = Case CEILING(((Round(Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ",1)))) When 5 Then 4 Else CEILING(((Round(Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ",1)))) End "
''    tmp = tmp & " Where Txn_InvMonGslDetail.YearMonth = '" & mFromDate & "' And Txn_InvMonGslDetail.ExchangeTypeID = 1 "
''    tmp = tmp & " And Txn_InvMonGslDetail.TFlag = 'W' And Txn_InvMonGslDetail.BillingCycleID = 2 And (ContractID = 0 Or ContractID is Null) "
''
''    If CmbPreFix.Text <> "" Then
''       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
''    End If
''
''
''    Call TmpTable
''
''    ProgressBar1.Value = ProgressBar1.Value + 1
''
''    'Withdrawal & Fortnightly
''    mNoofdays = GetBillingCycleDays(3)
''
''
''    tmp = "Update Txn_InvMonGslDetail Set NoOfDays = Case CEILING(((Round(Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ",1)))) When 3 Then 2 Else CEILING(((Round(Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ",1)))) End "
''    tmp = tmp & " Where Txn_InvMonGslDetail.YearMonth = '" & mFromDate & "' And Txn_InvMonGslDetail.ExchangeTypeID = 1"
''    tmp = tmp & " And Txn_InvMonGslDetail.TFlag = 'W' And Txn_InvMonGslDetail.BillingCycleID = 3 And (ContractID = 0 Or ContractID is Null)"
''
''    If CmbPreFix.Text <> "" Then
''       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
''    End If
''
''    Call TmpTable
''
''
''    ProgressBar1.Value = ProgressBar1.Value + 1
''
''    'Updating Invoice Amount And Invoice Date
''    tmp = " Update Txn_InvMonGslDetail Set InvoiceAmount = BalanceBags * NoOfDays * Rate, InvoiceDate = '" & mToDate & "' "
''    tmp = tmp & " Where YearMonth = '" & mFromDate & "'"
''    tmp = tmp & " And Txn_InvMonGslDetail.ExchangeTypeID = 1 And BillingUnit = 'B' And (ContractID = 0 Or ContractID is Null)"
''
''    If CmbPreFix.Text <> "" Then
''       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
''    End If
''
''    Call TmpTable
''
''
''    ProgressBar1.Value = ProgressBar1.Value + 1
''
''    tmp = " Update Txn_InvMonGslDetail Set InvoiceAmount = BalanceWeight * NoOfDays * Rate,InvoiceDate =  '" & mToDate & "' "
''    tmp = tmp & " Where YearMonth = '" & mFromDate & "'"
''    tmp = tmp & " And Txn_InvMonGslDetail.ExchangeTypeID = 1 And BillingUnit = 'Q' And (ContractID = 0 Or ContractID is Null) "
''
''    If CmbPreFix.Text <> "" Then
''       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
''    End If
''
''
''    Call TmpTable
''
''
''    ProgressBar1.Value = ProgressBar1.Value + 1
''
'''    tmp = " Update Txn_InvMonGslDetail Set InvoiceAmount = 0,InvoiceDate = Null Where YearMonth = '" & mFromDate & "' And Txn_InvMonGslDetail.ExchangeTypeID = 1 And TFlag = 'W' And (ContractID = 0 Or ContractID is Null) And Month(CxTFDate) = Month(CDTFDate) "
''
''    tmp = " Update Txn_InvMonGslDetail Set InvoiceAmount = 0,InvoiceDate = Null Where YearMonth = '" & mFromDate & "' And Txn_InvMonGslDetail.ExchangeTypeID = 1 And TFlag = 'W' And (ContractID = 0 Or ContractID is Null) And DateDiff(MOnth,CDTFDate,CxTFDate) = 0 "
''
''    If CmbPreFix.Text <> "" Then
''       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
''    End If
''
''
''    Call TmpTable
'
'
'    '==============================================
'
'
'    tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".Drno = "
'    tmp = tmp & " (Select a.SPRCID"
'    tmp = tmp & " from  TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " b"
'    tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
'    tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
'    tmp = tmp & " left join Mst_SpecialRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid "
'    tmp = tmp & " And a.ClientIDRow = b.ClientIDRow And a.commoditygroupid = c.commoditygroupid"
'    tmp = tmp & " where (b.ContractID = 0 Or b.ContractID is Null) And b.YearMonth = '" & mFromDate & "'  And a.SPRCID is not null And a.FLag = 'A' And b.GSLID = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".GSLID"
'    tmp = tmp & " And b.SM_Prefix = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix"
'    tmp = tmp & " Group by b.GSLID,b.SM_Prefix,"
'    tmp = tmp & " a.SPRCID, a.BillingCycleID,a.BillingUnit,"
'    tmp = tmp & " a.RateOnQty,a.BaseRate,a.AbsoluteDiscount,a.Amount)"
'    tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1"
'
'    If CmbPreFix.Text <> "" Then
'       tmp = tmp & " and TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix = '" & mSM_Prefix & "'"
'    End If
'
'    Call TmpTable
'
'
'    ProgressBar1.Value = ProgressBar1.Value + 1
'
'    tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".Rate = "
'    tmp = tmp & " (Select a.Amount"
'    tmp = tmp & " from  TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " b"
'    tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
'    tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
'    tmp = tmp & " left join Mst_SpecialRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid "
'    tmp = tmp & " And a.ClientIDRow = b.ClientIDRow And a.commoditygroupid = c.commoditygroupid"
'    tmp = tmp & " where (b.ContractID = 0 Or b.ContractID is Null) And b.YearMonth = '" & mFromDate & "'  And a.SPRCID is not null And a.FLag = 'A' And b.GSLID = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".GSLID"
'    tmp = tmp & " And b.SM_Prefix = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix"
'    tmp = tmp & " Group by b.GSLID,b.SM_Prefix,"
'    tmp = tmp & " a.SPRCID, a.BillingCycleID,a.BillingUnit,"
'    tmp = tmp & " a.RateOnQty,a.BaseRate,a.AbsoluteDiscount,a.Amount)"
'    tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1"
'
'    If CmbPreFix.Text <> "" Then
'       tmp = tmp & " and TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix = '" & mSM_Prefix & "'"
'    End If
'
'    Call TmpTable
'
'
'    ProgressBar1.Value = ProgressBar1.Value + 1
'
'    tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set Flag1 = Null Where YearMonth = '" & mFromDate & "'"
'
'    If CmbPreFix.Text <> "" Then
'       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
'    End If
'
'    Call TmpTable
'
'
'    ProgressBar1.Value = ProgressBar1.Value + 1
'
'    tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set Flag1 = 'S' Where DRNo is not null And YearMonth = '" & mFromDate & "' "
'
'    If CmbPreFix.Text <> "" Then
'       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
'    End If
'
'    Call TmpTable
'
'
'    ProgressBar1.Value = ProgressBar1.Value + 1
'
'
'    tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".Rate ="
'    tmp = tmp & " (Select  MSRCD.RatePerBag"
'    tmp = tmp & " From TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " TIMGD"
'    tmp = tmp & " Inner Join Mst_SpecialRateCardDetail MSRCD On TIMGD.DRNo = MSRCD.SPRCID"
'    tmp = tmp & " Where TIMGD.YearMonth = '" & mFromDate & "' And TIMGD.DRNo is Not null And TIMGD.BagSize >= MSRCD.FromBagSize And"
'    tmp = tmp & " TIMGD.BagSize <= MSRCD.ToBagSize And TIMGD.GSLID = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".GSLID"
'    tmp = tmp & " And TIMGD.SM_Prefix = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix And TIMGD.BillingUnit = 'B'"
'    tmp = tmp & " Group By TIMGD.GSLID,TIMGD.SM_Prefix,TIMGD.BagSize,MSRCD.FromBagSize,MSRCD.ToBagSize,MSRCD.RatePerBag)"
'    tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingUnit = 'B' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".Flag1 = 'S' "
'    tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag in ('O','W')"
'
'    If CmbPreFix.Text <> "" Then
'       tmp = tmp & " and TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix = '" & mSM_Prefix & "'"
'    End If
'
'
'    Call TmpTable
'
'
'    ProgressBar1.Value = ProgressBar1.Value + 1
'
'
'    tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".Rate ="
'    tmp = tmp & " (Select  MSRCD.RatePerBag"
'    tmp = tmp & " From TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " TIMGD"
'    tmp = tmp & " Inner Join Mst_SpecialRateCardDepoDetail MSRCD On TIMGD.DRNo = MSRCD.SPRCID"
'    tmp = tmp & " Where TIMGD.YearMonth = '" & mFromDate & "' And TIMGD.DRNo is Not null And TIMGD.BagSize >= MSRCD.FromBagSize And"
'    tmp = tmp & " TIMGD.BagSize <= MSRCD.ToBagSize And TIMGD.GSLID = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".GSLID"
'    tmp = tmp & " And TIMGD.SM_Prefix = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix And TIMGD.BillingUnit = 'B'"
'    tmp = tmp & " Group By TIMGD.GSLID,TIMGD.SM_Prefix,TIMGD.BagSize,MSRCD.FromBagSize,MSRCD.ToBagSize,MSRCD.RatePerBag)"
'    tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingUnit = 'B' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".Flag1 = 'S' "
'    tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag in ('D')"
'
'    If CmbPreFix.Text <> "" Then
'       tmp = tmp & " and TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix = '" & mSM_Prefix & "'"
'    End If
'
'
'    Call TmpTable
'
'
'    ProgressBar1.Value = ProgressBar1.Value + 1
'
'
'   '--------------------------------------------------
'   '| B. Apply  Base Rate to GSL Table.                |
'   '--------------------------------------------------
'
'
'    tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingCycleID ="
'    tmp = tmp & " (Select a.BillingCycleID"
'    tmp = tmp & " from  TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " b"
'    tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
'    tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
'    tmp = tmp & " left join Mst_StorageRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid And a.commoditygroupid = c.commoditygroupid"
'    tmp = tmp & " Where a.SRCID is not null And (b.ContractID = 0 Or b.ContractID is Null) And b.YearMonth = '" & mFromDate & "' And a.ExchangeTypeID = 1 And a.DefaultRate = 1 And b.Flag1 is null "
'    tmp = tmp & " And b.GSLID = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".GSLID And b.SM_Prefix = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix"
'    tmp = tmp & " Group by b.SM_Prefix,b.GSLID,b.ExchangeTypeID,"
'    tmp = tmp & " a.SRCID, a.BillingCycleID,a.Flag,a.RateOnQty)"
'    tmp = tmp & " Where (TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".Flag1 is Null)  And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag in ('O','W') "
'    tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 "
'
'    If CmbPreFix.Text <> "" Then
'       tmp = tmp & " and TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix = '" & mSM_Prefix & "'"
'    End If
'
'    Call TmpTable
'
'
'    ProgressBar1.Value = ProgressBar1.Value + 1
'
'    tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingCycleID ="
'    tmp = tmp & " (Select a.DepoBillingCycleID"
'    tmp = tmp & " from  TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " b"
'    tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
'    tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
'    tmp = tmp & " left join Mst_StorageRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid And a.commoditygroupid = c.commoditygroupid"
'    tmp = tmp & " Where a.SRCID is not null And (b.ContractID = 0 Or b.ContractID is Null) And b.YearMonth = '" & mFromDate & "' And a.ExchangeTypeID = 1 And a.DefaultRate = 1 And b.Flag1 is null "
'    tmp = tmp & " And b.GSLID = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".GSLID And b.SM_Prefix = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix"
'    tmp = tmp & " Group by b.SM_Prefix,b.GSLID,b.ExchangeTypeID,"
'    tmp = tmp & " a.SRCID, a.DepoBillingCycleID,a.Flag,a.RateOnQty)"
'    tmp = tmp & " Where (TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".Flag1 is Null)  And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag in ('D') "
'    tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 "
'
'    If CmbPreFix.Text <> "" Then
'       tmp = tmp & " and TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix = '" & mSM_Prefix & "'"
'    End If
'
'    Call TmpTable
'
'
'    ProgressBar1.Value = ProgressBar1.Value + 1
'
'    tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingUnit ="
'    tmp = tmp & " (Select a.Flag"
'    tmp = tmp & " from  TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " b"
'    tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
'    tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
'    tmp = tmp & " left join Mst_StorageRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid And a.commoditygroupid = c.commoditygroupid"
'    tmp = tmp & " Where a.SRCID is not null And (b.ContractID = 0 Or b.ContractID is Null) And b.YearMonth = '" & mFromDate & "' And a.ExchangeTypeID = 1 And a.DefaultRate = 1 And b.Flag1 is null"
'    tmp = tmp & " And b.GSLID = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".GSLID And b.SM_Prefix = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix"
'    tmp = tmp & " Group by b.SM_Prefix,b.GSLID,b.ExchangeTypeID,"
'    tmp = tmp & " a.SRCID, a.BillingCycleID,a.Flag,a.RateOnQty)"
'    tmp = tmp & " Where (TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".Flag1 is Null) "
'    tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 "
'
'    If CmbPreFix.Text <> "" Then
'       tmp = tmp & " and TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix = '" & mSM_Prefix & "'"
'    End If
'
'    Call TmpTable
'
'    ProgressBar1.Value = ProgressBar1.Value + 1
'
'    tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".DRNo ="
'    tmp = tmp & " (Select a.SRCID"
'    tmp = tmp & " from  TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " b"
'    tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
'    tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
'    tmp = tmp & " left join Mst_StorageRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid And a.commoditygroupid = c.commoditygroupid"
'    tmp = tmp & " Where a.SRCID is not null And (b.ContractID = 0 Or b.ContractID is Null) And b.YearMonth = '" & mFromDate & "' And a.ExchangeTypeID = 1 And a.DefaultRate = 1 And b.Flag1 is null"
'    tmp = tmp & " And b.GSLID = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".GSLID And b.SM_Prefix = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix"
'    tmp = tmp & " Group by b.SM_Prefix,b.GSLID,b.ExchangeTypeID,"
'    tmp = tmp & " a.SRCID, a.BillingCycleID,a.Flag,a.RateOnQty)"
'    tmp = tmp & " Where (TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".Flag1 is Null) "
'    tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "'  And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1 "
'
'
'    If CmbPreFix.Text <> "" Then
'       tmp = tmp & " and TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix = '" & mSM_Prefix & "'"
'    End If
'
'
'    Call TmpTable
'
'
'    ProgressBar1.Value = ProgressBar1.Value + 1
'
'    tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".Rate = "
'    tmp = tmp & " (Select a.RateOnQty"
'    tmp = tmp & " from  TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " b"
'    tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
'    tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
'    tmp = tmp & " left join Mst_StorageRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid And a.commoditygroupid = c.commoditygroupid"
'    tmp = tmp & " Where a.SRCID is not null And (b.ContractID = 0 Or b.ContractID is Null) And b.YearMonth = '" & mFromDate & "' And a.ExchangeTypeID = 1 And a.DefaultRate = 1 And b.Flag1 is null"
'    tmp = tmp & " And b.GSLID = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".GSLID And b.SM_Prefix = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix"
'    tmp = tmp & " Group by b.SM_Prefix,b.GSLID,b.ExchangeTypeID,"
'    tmp = tmp & " a.SRCID, a.BillingCycleID,a.Flag,a.RateOnQty)"
'    tmp = tmp & " Where (TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".Flag1 is Null) "
'    tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 "
'
'    If CmbPreFix.Text <> "" Then
'       tmp = tmp & " and TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix = '" & mSM_Prefix & "'"
'    End If
'
'    Call TmpTable
'
'    ProgressBar1.Value = ProgressBar1.Value + 1
'
'    tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set Flag1 = 'N' Where DRNo is not null And YearMonth = '" & mFromDate & "' And Flag1 is null And BillingCycleID is not Null And ExchangeTypeID = 1 "
'
'    If CmbPreFix.Text <> "" Then
'       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
'    End If
'
'    Call TmpTable
'
'
'    ProgressBar1.Value = ProgressBar1.Value + 1
'
'    tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".Rate ="
'    tmp = tmp & " (Select MSRCD.RatePerBag"
'    tmp = tmp & " From TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " TIMGD"
'    tmp = tmp & " Inner Join Mst_StorageRateCardDetail MSRCD On TIMGD.DRNo = MSRCD.SRCID"
'    tmp = tmp & " Inner Join Mst_StorageRateCard MSRC On MSRCD.SRCID = MSRC.SRCID"
'    tmp = tmp & " Where TIMGD.YearMonth = '" & mFromDate & "' And TIMGD.DRNo is Not null And TIMGD.BagSize >= MSRCD.FromBagSize And"
'    tmp = tmp & " TIMGD.BagSize <= MSRCD.ToBagSize  And TIMGD.BillingUnit = 'B' And TIMGD.Flag1 = 'N' And  MSRC.DefaultRate = 1"
'    tmp = tmp & " And TIMGD.GSLID = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".GSLID And TIMGD.SM_Prefix = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix"
'    tmp = tmp & " And TIMGD.DRNo = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".DRNo"
'    tmp = tmp & " Group By TIMGD.GSLID,TIMGD.SM_Prefix,TIMGD.BagSize,MSRCD.FromBagSize,MSRCD.ToBagSize,MSRCD.RatePerBag)"
'    tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingUnit = 'B' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".Flag1 = 'N' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag in ('O','W') "
'    tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "'  And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1 "
'
'    If CmbPreFix.Text <> "" Then
'       tmp = tmp & " and TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix = '" & mSM_Prefix & "'"
'    End If
'
'    Call TmpTable
'
'    ProgressBar1.Value = ProgressBar1.Value + 1
'
'    tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".Rate ="
'    tmp = tmp & " (Select MSRCD.RatePerBag"
'    tmp = tmp & " From TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " TIMGD"
'    tmp = tmp & " Inner Join Mst_StorageRateCardDepoDetail MSRCD On TIMGD.DRNo = MSRCD.SRCID"
'    tmp = tmp & " Inner Join Mst_StorageRateCard MSRC On MSRCD.SRCID = MSRC.SRCID"
'    tmp = tmp & " Where TIMGD.YearMonth = '" & mFromDate & "' And TIMGD.DRNo is Not null And TIMGD.BagSize >= MSRCD.FromBagSize And"
'    tmp = tmp & " TIMGD.BagSize <= MSRCD.ToBagSize  And TIMGD.BillingUnit = 'B' And TIMGD.Flag1 = 'N' And  MSRC.DefaultRate = 1"
'    tmp = tmp & " And TIMGD.GSLID = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".GSLID And TIMGD.SM_Prefix = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix"
'    tmp = tmp & " And TIMGD.DRNo = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".DRNo"
'    tmp = tmp & " Group By TIMGD.GSLID,TIMGD.SM_Prefix,TIMGD.BagSize,MSRCD.FromBagSize,MSRCD.ToBagSize,MSRCD.RatePerBag)"
'    tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingUnit = 'B' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".Flag1 = 'N' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag in ('D') "
'    tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "'  And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1 "
'
'
'    If CmbPreFix.Text <> "" Then
'       tmp = tmp & " and TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix = '" & mSM_Prefix & "'"
'    End If
'
'    Call TmpTable
'
'    ProgressBar1.Value = ProgressBar1.Value + 1
'
'   '--------------------------------------------------
'   '| C. Apply  No Of Cycles to GSL Table as per Billing Unit and Transaction Flag.                |
'   '--------------------------------------------------
'
'    'Op Bal and Monthly Cycle
'    tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoofDays = 1 Where YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TFlag = 'O'"
'    tmp = tmp & " And BillingCycleID = 4  And (ContractID = 0 Or ContractID is Null)"
'
'    If CmbPreFix.Text <> "" Then
'       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
'    End If
'
'    Call TmpTable
'
'
'
'    ProgressBar1.Value = ProgressBar1.Value + 1
'
'    'Op Bal and Forthnightly Cycle
'    tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoofDays = 2 Where YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TFlag = 'O'"
'    tmp = tmp & " And BillingCycleID = 3 And (ContractID = 0 Or ContractID is Null)"
'
'    If CmbPreFix.Text <> "" Then
'       tmp = tmp & " and TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix = '" & mSM_Prefix & "'"
'    End If
'
'
'    Call TmpTable
'
'
'
'    ProgressBar1.Value = ProgressBar1.Value + 1
'
'    'Op Bal and Weekly Cycle
'    tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoofDays = 4 Where YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TFlag = 'O'"
'    tmp = tmp & " And BillingCycleID = 2 And (ContractID = 0 Or ContractID is Null)"
'
'    If CmbPreFix.Text <> "" Then
'       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
'    End If
'
'
'    Call TmpTable
'
'
'
'    ProgressBar1.Value = ProgressBar1.Value + 1
'    'Op Bal and Daily Cycle
'    tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoofDays = " & mTotalDays & " Where YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TFlag = 'O'"
'    tmp = tmp & " And BillingCycleID = 1 And (ContractID = 0 Or ContractID is Null)"
'
'    If CmbPreFix.Text <> "" Then
'       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
'    End If
'
'
'    Call TmpTable
'
'    ProgressBar1.Value = ProgressBar1.Value + 1
'
'    'Depo and Withd and Monthly Cycle
'
'    tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoofDays = 1 Where YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TFlag <> 'O'"
'    tmp = tmp & " And BillingCycleID = 4 And (ContractID = 0 Or ContractID is Null)"
'
'    If CmbPreFix.Text <> "" Then
'       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
'    End If
'
'
'    Call TmpTable
'
'    ProgressBar1.Value = ProgressBar1.Value + 1
'
'    'Depo and Withd and Daily Cycle
'
'
''    Commented on 16 March 2011 Parag
'
''    tmp = "Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoOfDays = 30-Day(CxTFDate) + 1 Where YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TFlag = 'D' And BillingCycleID = 1 And (ContractID = 0 Or ContractID is Null) "
''
''    If CmbPreFix.Text <> "" Then
''       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
''    End If
''
''   Call TmpTable
'
'
'
'''    tmp = "Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoOfDays = 30-Day(CxTFDate) + 1 Where YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TFlag = 'D' And BillingCycleID = 1 And (ContractID = 0 Or ContractID is Null)  And DateDiff(Month,CxTFDate,DateofWithdrawal) = 0 "
''
''    'tmp = "Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoOfDays = Day(DateOfWithdrawal)-Day(CxTFDate) + 1 Where YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TFlag = 'D' And BillingCycleID = 1 And (ContractID = 0 Or ContractID is Null)  And DateDiff(Month,CxTFDate,DateofWithdrawal) = 0 "
''
''    If CmbPreFix.Text <> "" Then
''       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
''    End If
''
''    Call TmpTable
'
'    'tmp = "Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoOfDays = 30-Day(CxTFDate) + 1 Where YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TFlag = 'D' And BillingCycleID = 1 And (ContractID = 0 Or ContractID is Null)  " 'And DateofWithdrawal is Null "
'
'    tmp = "Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoOfDays = " & mTotalDays & " - Day(CxTFDate) + 1 Where YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TFlag = 'D' And BillingCycleID = 1 And (ContractID = 0 Or ContractID is Null)  " 'And DateofWithdrawal is Null "
'
'
'    If CmbPreFix.Text <> "" Then
'       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
'    End If
'
'    Call TmpTable
'
'
'
'    ProgressBar1.Value = ProgressBar1.Value + 1
'
'    'Withd and Daily Cycle
'
'    tmp = "Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoOfDays = Day(CxTFDate) Where YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TFlag = 'W' And BillingCycleID = 1 And (ContractID = 0 Or ContractID is Null)"
'
'    If CmbPreFix.Text <> "" Then
'       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
'    End If
'
'    Call TmpTable
'
'
'
'    ProgressBar1.Value = ProgressBar1.Value + 1
'
'    'Deposit & Weekly
'    mNoofdays = GetBillingCycleDays(2)
'
''    tmp = "Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoOfDays = Floor((4 - Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ") + 1)"
''    tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1"
''    tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag = 'D' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingCycleID = 2 And (ContractID = 0 Or ContractID is Null)"
''
''    Call TmpTable
'
'
''    tmp = "Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoOfDays = 1 Where YearMonth = '" & mFromDate & "' And ExchangeTypeID = 1 And TFlag = 'D' And BillingCycleID in (2) And"
''    tmp = tmp & " Floor ((4 - (Convert(Numeric, Day(CxTFDate)) / 7) + 1)) < 1"
''
''    Call TmpTable
'
'
'' Commented on 16th March 2011 722 PM Parag
'
''    tmp = "Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoOfDays =  Case Floor((4 - (Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ") + 1))  When 0 then 1 else Floor((4 - (Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ") + 1)) End "
''    tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1"
''    tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag = 'D' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingCycleID = 2 And (ContractID = 0 Or ContractID is Null)"
''
''    If CmbPreFix.Text <> "" Then
''       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
''    End If
''
''
''    Call TmpTable
'
'
'    tmp = "Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoOfDays =  Case Floor((4 - (Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ") + 1))  When 0 then 1 else Floor((4 - (Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ") + 1)) End "
'    tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1"
'    tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag = 'D' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingCycleID = 2 And (ContractID = 0 Or ContractID is Null) " 'And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".DateOfWithdrawal is null "
'
'    If CmbPreFix.Text <> "" Then
'       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
'    End If
'
'    Call TmpTable
'
'
'
''    tmp = "Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoOfDays =  Case Floor(((Case CEILING(((Round(Convert(Numeric,Day(DateOfWithdrawal))/" & mNoofdays & ",1)))) When 5 Then 4 Else CEILING(((Round(Convert(Numeric,Day(DateOfWithdrawal))/" & mNoofdays & ",1)))) End) - (Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ") + 1))  When 0 then 1 else Floor(((Case CEILING(((Round(Convert(Numeric,Day(DateOfWithdrawal))/7,1)))) When 5 Then 4 Else CEILING(((Round(Convert(Numeric,Day(DateOfWithdrawal))/7,1)))) End) - (Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ") + 1)) End "
''    tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1"
''    tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag = 'D' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingCycleID = 2 And (ContractID = 0 Or ContractID is Null) And DateDiff(Month,CDTFDate,DateOfWithdrawal) = 0"
''
''    If CmbPreFix.Text <> "" Then
''       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
''    End If
''
''    Call TmpTable
'
'    ProgressBar1.Value = ProgressBar1.Value + 1
'
'    'Deposit & FortNightly
'    mNoofdays = GetBillingCycleDays(3)
'
'
''    tmp = "Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoOfDays = Floor(((4 - Round(Convert(Decimal,Day(CxTFDate))/" & mNoofdays & ",1)) - 1))"
''    tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1"
''    tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag = 'D' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingCycleID = 3 And (ContractID = 0 Or ContractID is Null)"
''
''    Call TmpTable
'
'
''    tmp = "Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoOfDays = Case Floor((4 - (Convert(Numeric,Day(CxTFDate))/7) + 1))  When 0 then 1 else Floor((4 - (Convert(Numeric,Day(CxTFDate))/7) + 1)) End "
''    tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1"
''    tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag = 'D' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingCycleID = 3 And (ContractID = 0 Or ContractID is Null)"
'
'
'' Commented on 16th March 2011 722 PM Parag
'
''    tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoOfDays = Case Floor((4 - Round(Convert(Decimal,Day(CxTFDate))/" & mNoofdays & ",1)) - 1) When 0 then 1 else Floor((4 - Round(Convert(Decimal,Day(CxTFDate))/" & mNoofdays & ",1)) - 1) End "
''    tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1 "
''    tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag = 'D' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingCycleID = 3 And (ContractID = 0 Or ContractID is Null) "
'
''    If CmbPreFix.Text <> "" Then
''       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
''    End If
'
'
''    Call TmpTable
'
'
'    'Commented on 21 March 2011
''    tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoOfDays = Case Floor(((Case CEILING(((Round(Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ",1)))) When 3 Then 2 Else CEILING(((Round(Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ",1)))) End) - Round(Convert(Decimal,Day(CxTFDate))/" & mNoofdays & ",1)) - 1) When 0 then 1 else Floor(((Case CEILING(((Round(Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ",1)))) When 3 Then 2 Else CEILING(((Round(Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ",1)))) End) - Round(Convert(Decimal,Day(CxTFDate))/" & mNoofdays & ",1)) - 1) End "
''    tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1 "
''    tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag = 'D' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingCycleID = 3 And (ContractID = 0 Or ContractID is Null) And DateDiff(Month,CDTFDate,DateOfWithdrawal) = 0 "
''
''    If CmbPreFix.Text <> "" Then
''       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
''    End If
''
''
''    Call TmpTable
'
'
''    tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoOfDays = 1 "
''    tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1  And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag = 'D' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingCycleID = 3 And (ContractID = 0 Or ContractID is Null) And DateDiff(Month,CDTFDate,DateOfWithdrawal) = 0 "
''    tmp = tmp & " And Day(DateOfWithdrawal) <= 15 And Day(CxTFDate) <= 15 "
''
''    If CmbPreFix.Text <> "" Then
''       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
''    End If
''
''    Call TmpTable
''
''    tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoOfDays = 2 "
''    tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1  And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag = 'D' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingCycleID = 3 And (ContractID = 0 Or ContractID is Null) And DateDiff(Month,CDTFDate,DateOfWithdrawal) = 0 "
''    tmp = tmp & " And Day(DateOfWithdrawal) >15 And Day(CxTFDate) <= 15"
''
''    If CmbPreFix.Text <> "" Then
''       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
''    End If
''
''    Call TmpTable
''
''    tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoOfDays = 1 "
''    tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1  And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag = 'D' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingCycleID = 3 And (ContractID = 0 Or ContractID is Null) And DateDiff(Month,CDTFDate,DateOfWithdrawal) = 0 "
''    tmp = tmp & " And Day(DateOfWithdrawal) <= 31 And Day(CxTFDate) > 15 "
''
''    If CmbPreFix.Text <> "" Then
''       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
''    End If
''
''    Call TmpTable
'
'
''    tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoOfDays = Case Floor((4 - Round(Convert(Decimal,Day(CxTFDate))/" & mNoofdays & ",1)) - 1) When 0 then 1 else Floor((4 - Round(Convert(Decimal,Day(CxTFDate))/" & mNoofdays & ",1)) - 1) End "
''    tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1 "
''    tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag = 'D' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingCycleID = 3 And (ContractID = 0 Or ContractID is Null) And DateofWithdrawal is null "
'
''    If CmbPreFix.Text <> "" Then
''       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
''    End If
'
'
'    'Case 1 Depo Date  <= 15 Then No Of Cycle = 2
'    tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoOfDays = 2 "
'    tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1 "
'    tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag = 'D' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingCycleID = 3 And (ContractID = 0 Or ContractID is Null) " 'And DateofWithdrawal is null "
'    tmp = tmp & " And Day(CxTFDate) <= 15 "
'
'    If CmbPreFix.Text <> "" Then
'       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
'    End If
'
'    Call TmpTable
'
'    ProgressBar1.Value = ProgressBar1.Value + 1
'
'    'Case 1 Depo Date  > 15 Then No Of Cycle = 1
'    tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoOfDays = 1 "
'    tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1 "
'    tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag = 'D' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingCycleID = 3 And (ContractID = 0 Or ContractID is Null) " 'And DateofWithdrawal is null "
'    tmp = tmp & " And Day(CxTFDate) > 15 "
'
'    If CmbPreFix.Text <> "" Then
'       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
'    End If
'
'    Call TmpTable
'
'    ProgressBar1.Value = ProgressBar1.Value + 1
'
'    'Withdrawal & weekly
'    mNoofdays = GetBillingCycleDays(2)
'
'    tmp = "Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoOfDays = Case CEILING(((Round(Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ",1)))) When 5 Then 4 Else CEILING(((Round(Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ",1)))) End "
'    tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1 "
'    tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag = 'W' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingCycleID = 2 And (ContractID = 0 Or ContractID is Null) "
'
'    If CmbPreFix.Text <> "" Then
'       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
'    End If
'
'
'    Call TmpTable
'
'    ProgressBar1.Value = ProgressBar1.Value + 1
'
'    'Withdrawal & Fortnightly
'    mNoofdays = GetBillingCycleDays(3)
'
'
'    tmp = "Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoOfDays = Case CEILING(((Round(Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ",1)))) When 3 Then 2 Else CEILING(((Round(Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ",1)))) End "
'    tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1"
'    tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag = 'W' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingCycleID = 3 And (ContractID = 0 Or ContractID is Null)"
'
'    If CmbPreFix.Text <> "" Then
'       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
'    End If
'
'    Call TmpTable
'
'
'    ProgressBar1.Value = ProgressBar1.Value + 1
'
'    'Updating Invoice Amount And Invoice Date
'    tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set InvoiceAmount = BalanceBags * NoOfDays * Rate, InvoiceDate = '" & mToDate & "' "
'    tmp = tmp & " Where YearMonth = '" & mFromDate & "'"
'    tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1 And BillingUnit = 'B' And (ContractID = 0 Or ContractID is Null)"
'
'    If CmbPreFix.Text <> "" Then
'       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
'    End If
'
'    Call TmpTable
'
'
'    ProgressBar1.Value = ProgressBar1.Value + 1
'
'    tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set InvoiceAmount = BalanceWeight * NoOfDays * Rate,InvoiceDate =  '" & mToDate & "' "
'    tmp = tmp & " Where YearMonth = '" & mFromDate & "'"
'    tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1 And BillingUnit = 'Q' And (ContractID = 0 Or ContractID is Null) "
'
'    If CmbPreFix.Text <> "" Then
'       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
'    End If
'
'
'    Call TmpTable
'
'
'    ProgressBar1.Value = ProgressBar1.Value + 1
'
''    tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set InvoiceAmount = 0,InvoiceDate = Null Where YearMonth = '" & mFromDate & "' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1 And TFlag = 'W' And (ContractID = 0 Or ContractID is Null) And Month(CxTFDate) = Month(CDTFDate) "
'
'    tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set InvoiceAmount = 0,InvoiceDate = Null Where YearMonth = '" & mFromDate & "' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1 And TFlag = 'W' And (ContractID = 0 Or ContractID is Null) And DateDiff(Month,CDTFDate,CxTFDate) = 0 "
'
'    If CmbPreFix.Text <> "" Then
'       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
'    End If
'
'
'    Call TmpTable
'
'
'    ProgressBar1.Value = ProgressBar1.Value + 1
'
'
'    tmp = " Insert into Txn_InvMonGSLDetail "
'    tmp = tmp & " (GSLID,SM_Prefix,CMPID,CommodityID,GodownID,StackNo,LotNo,BalanceBags,"
'    tmp = tmp & " BalanceWeight,ExchangeTypeID,ClientIDRow,Flag,Flag1,Tflag,GrossWt,KRDate,"
'    tmp = tmp & " ContractID,BagSize,YearMonth,G_UID,CreatedBy,CreatedDate,"
'    tmp = tmp & " UpdatedDate,UpdatedBy,CDTFDate,CDTFNo,BillingCycleID ,BillingUnit,Drno,Rate,NoofDays,InvoiceAmount,CxTFNo,CxTFDate,InvoiceDate)"
'    tmp = tmp & " Select GSLID,SM_Prefix,CMPID,CommodityID,GodownID,StackNo,LotNo,BalanceBags,"
'    'tmp = tmp & " BalanceWeight,ExchangeTypeID,ClientIDRow,Flag,Flag1,'O',GrossWt,KRDate,"
'
'    tmp = tmp & " BalanceWeight,ExchangeTypeID,ClientIDRow,Flag,Flag1,TFlag,GrossWt,KRDate,"
'    tmp = tmp & " NContractID,BagSize,'" & mFromDate & "',G_UID,CreatedBy,CreatedDate,"
'    tmp = tmp & " UpdatedDate , UpdatedBy,CDTFDate,CDTFNo,BillingCycleID ,BillingUnit,Drno,Rate,NoofDays,InvoiceAmount,CxTFNo,CxTFDate,InvoiceDate,KRNo "
'    tmp = tmp & " From TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ""
'
'    Call TmpTable
'
'    tmp = "Drop Table TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ""
'    Call TmpTable
'
''End If
'
'
'ProgressBar1.Value = ProgressBar1.Max
'
'ChkCheckAll.Value = 0
'Call Grid_Head
'MsgBox "Done !!"
'
''CmdProcess.Enabled = True
'
'TxtConfirmation = ""
'TxtRateCard.Text = ""
'TxtDTD.Text = ""
'TxtGSLMapping.Text = ""
'TxtConfirmation.Locked = True
'
'Exit Sub
'
'msgError:
'
'    tmp = " if exists (select * from dbo.sysobjects where id = object_id(N'dbo.TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & "') and OBJECTPROPERTY(id, N'IsUserTable') = 1)"
'    tmp = tmp & " drop table dbo.TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ""
'    Call TmpTable
'
'    MsgBox "Some problem occured during process!!!"
'
'    TxtConfirmation = ""
'    TxtRateCard.Text = ""
'    TxtDTD.Text = ""
'    TxtGSLMapping.Text = ""
'    TxtConfirmation.Locked = True
'
'End Sub

Private Sub CmdShowLocation_Click()

Call Grid_Head
ChkCheckAll.Value = 0
CmdProcess.Enabled = True
ProgressBar1.Value = 0

If CmbProcessType.Text = "" Then
   MsgBox " Blank Process Type not allowed  !!!"
   CmbProcessType.SetFocus
   Exit Sub
End If

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

mMonth = mMonthListindex + 1

If mMonth < 10 Then
   mMonth = "0" & mMonth
End If

mFromDate = Format(CmbYear.Text & "-" & mMonth & "-" & "01", "yyyy-mm-dd")

mYear = Val(CmbYear.Text)
mTotalDays = Day(DateSerial(mYear, mMonth + 1, 0))

mToDate = Format(CmbYear.Text & "-" & mMonth & "-" & mTotalDays, "yyyy-mm-dd")


'-------Commented By Sagar on 26-08-211

If CDate(Format(mFromDate, "yyyymm")) >= Format(Date, "yyyymm") Then
   MsgBox "Present and Future Month not Allowed !!!"
   CmbMonth.SetFocus
   Exit Sub
End If

Dim mChooseLocation As Variant

mChooseLocation = ""


If LCase(CmbProcessType.Text) = "franchisee" Then
   If Gen_DBType = "MDB" Then
      tmp = " Select Distinct TIG.SM_Prefix From Txn_InvMonGSLDetail TIG"
      tmp = tmp & " Inner join Mst_Godown MG on TIG.GodownID=MG.GodownID"
      tmp = tmp & " Where (MG.CloseDate is null Or MG.CloseDate >= '" & mFromDate & "') And MG.StartDate <= '" & mToDate & "' And MG.UnitTypeID in (2,6) And TIG.YearMonth = #" & mFromDate & "# And isnull(TIG.FranchProcessFlag,'N')='N'"
      
   Else
      tmp = " Select Distinct TIG.SM_Prefix From Txn_InvMonGSLDetail TIG"
      tmp = tmp & " Inner join Mst_Godown MG on TIG.GodownID=MG.GodownID"
      tmp = tmp & " Where (MG.CloseDate is null Or MG.CloseDate >= '" & mFromDate & "') And MG.StartDate <= '" & mToDate & "' And MG.UnitTypeID in (2,6) And TIG.YearMonth = '" & mFromDate & "' And isnull(TIG.FranchProcessFlag,'N')='N'"
   End If
Else
   If Gen_DBType = "MDB" Then
      tmp = " Select Distinct TIG.SM_Prefix From Txn_InvMonGSLDetail TIG"
      tmp = tmp & " Where TIG.ExchangeTypeID=2 And TIG.YearMonth = #" & mFromDate & "# " 'And isnull(TIG.FranchProcessFlag,'N')='N'"
      ''''added by rakesh on 15 march 2013 (as in combo box select only franchisee)
      If LCase(Gen_UserRole) <> "power" Then
         tmp = tmp & " And TIG.SM_Prefix not in ("
         tmp = tmp & " Select Distinct TIG.SM_Prefix From Txn_InvMonGSLDetail TIG"
         tmp = tmp & " Inner join Mst_Godown MG on TIG.GodownID=MG.GodownID"
         tmp = tmp & " Where (MG.CloseDate is null Or MG.CloseDate >= '" & mFromDate & "') And MG.StartDate <= '" & mToDate & "' And MG.UnitTypeID in (2,6) And TIG.YearMonth = #" & mFromDate & "# And isnull(TIG.FranchProcessFlag,'N')='Y')"
      End If
      
   Else
      tmp = " Select Distinct TIG.SM_Prefix From Txn_InvMonGSLDetail TIG"
      tmp = tmp & " Where TIG.ExchangeTypeID=2 And TIG.YearMonth = '" & mFromDate & "'" ' And isnull(TIG.FranchProcessFlag,'N')='N'"
      ''''added by rakesh on 15 march 2013 (as in combo box select only franchisee)
      If LCase(Gen_UserRole) <> "power" Then
         tmp = tmp & " And TIG.SM_Prefix not in ("
         tmp = tmp & " Select Distinct TIG.SM_Prefix From Txn_InvMonGSLDetail TIG"
         tmp = tmp & " Inner join Mst_Godown MG on TIG.GodownID=MG.GodownID"
         tmp = tmp & " Where MG.UnitTypeID in (2,6) And TIG.YearMonth = '" & mFromDate & "' And isnull(TIG.FranchProcessFlag,'N')='Y')"
      End If
   End If
End If

Call TmpTable

If TmpRs.RecordCount = 0 Then Exit Sub

For I = 1 To TmpRs.RecordCount
    If mChooseLocation = "" Then
       mChooseLocation = "(" & "'" & TmpRs!SM_Prefix & "'"
    Else
       mChooseLocation = mChooseLocation & "," & "'" & TmpRs!SM_Prefix & "'"
    End If
    
    If I = TmpRs.RecordCount Then
       mChooseLocation = mChooseLocation & ")"
    End If
    TmpRs.MoveNext
Next

tmp = "Select Distinct ML.LocationName,ML.SM_Prefix  "
tmp = tmp & " From Mst_Location ML "
tmp = tmp & " Left Join Mst_CMP  MC On ML.LocationID = MC.LocationID "
tmp = tmp & " Left Join Mst_GSL MG On MC.CMPID = MG.CMPID "

If mChooseLocation <> "" Then
   tmp = tmp & " Where ML.SM_Prefix in " & mChooseLocation & ""
End If

If Gen_WcLocation <> "" Then
   If mChooseLocation = "" Then
      tmp = tmp & " Where ML." & Gen_WcLocation
   Else
      tmp = tmp & " And ML." & Gen_WcLocation
   End If
End If

tmp = tmp & " Order By ML.LocationName "

Call TmpTable

ProgressBar1.Value = 0

ProgressBar1.Max = TmpRs.RecordCount + 1
MSHFlexGrid1.Rows = TmpRs.RecordCount + 2

With MSHFlexGrid1
    For mR = 1 To TmpRs.RecordCount
        .TextMatrix(mR, 1) = TmpRs!Locationname
        .TextMatrix(mR, 2) = TmpRs!SM_Prefix
        
        .row = mR
        .Col = 0
        .CellFontName = "Wingdings"
        .CellFontSize = 12
        .CellAlignment = flexAlignCenterCenter
        .Text = strUnChecked
         
        TmpRs.MoveNext
        ProgressBar1.Value = ProgressBar1.Value + 1
    Next
End With

ProgressBar1.Value = ProgressBar1.Max
LblTotalRecord.Refresh
LblTotalRecord.Caption = TmpRs.RecordCount
mCountLocation = 0

End Sub

Private Sub CmdUpdateGodownCode_Click()


If CmbOldLocation.Text = "" Then
   MsgBox "Blank Old Location Not Allowed!!!"
   CmbOldLocation.SetFocus
   Exit Sub
End If

If CmbCMPName.Text = "" Then
   MsgBox "Blank CMP Name Not Allowed!!!"
   CmbCMPName.SetFocus
   Exit Sub
End If
If CmbGodownNo.Text = "" Then
   MsgBox "Blank Old Godown Not Allowed!!!"
   CmbGodownNo.SetFocus
   Exit Sub
End If

If Trim(TxtNewGodowNo.Text) = "" Then
   MsgBox "Blank New GodownNo Not Allowed!!!"
   TxtNewGodowNo.SetFocus
   Exit Sub
End If

If Trim(TxtNewGodowNo.Text) <> "" Then
   Call TableMst_Godown(" Where CMPID = " & mCMPID & " and GodownNo= '" & Trim(TxtNewGodowNo.Text) & "'")
   If RsMst_Godown.RecordCount > 0 Then
      MsgBox "Already same Godown No exist in selected CMP!!!"
      TxtNewGodowNo.SetFocus
      Exit Sub
   End If
   
   tmp = "Update Mst_Godown Set GodownNo = '" & Trim(TxtNewGodowNo.Text) & "',UpdatedBy = '" & Gen_UserLoginID & "',UpdatedDate = '" & Now() & "'  Where CMPID = " & mCMPID & " And GodownID = " & mGodownID & " "
   Call Tmp1Table
   
   tmp = "Update Mst_GodownInspection Set GodownNo = '" & Trim(TxtNewGodowNo.Text) & "' Where CMPID = " & mCMPID & " And GodownNo = '" & CmbGodownNo.Text & "' "
   Call Tmp1Table
   
   MsgBox "Godown No update , please inform details to IT software team to update in mobile app system!!!"
   
   TxtNewGodowNo.Text = ""
   CmbGodownNo.Text = ""
   CmbCMPName.Text = ""
   CmbOldLocation.Text = ""

End If


End Sub

'Private Sub CmdUpdatePendingRateCard_Click()
'
'If CmbProcessType.Text = "" Then
'   MsgBox " Blank Process Type not allowed  !!!"
'   CmbProcessType.SetFocus
'   Exit Sub
'End If
'
'If CmbMonth.Text = "" Then
'   MsgBox " Blank Month not allowed  !!!"
'   CmbMonth.SetFocus
'   Exit Sub
'End If
'
'If CmbYear.Text = "" Then
'   MsgBox " Blank Year not allowed  !!!"
'   CmbYear.SetFocus
'   Exit Sub
'End If
'
'mMonth = mMonthListindex + 1
'
'If mMonth < 10 Then
'   mMonth = "0" & mMonth
'End If
'
'mFromDate = Format(CmbYear.Text & "-" & mMonth & "-" & "01", "yyyy-mm-dd")
'mYear = Val(CmbYear.Text)
'mTotalDays = Day(DateSerial(mYear, mMonth + 1, 0))
'
'mToDate = Format(CmbYear.Text & "-" & mMonth & "-" & mTotalDays, "yyyy-mm-dd")
'mNoofdays = ""
'
'If CDate(Format(mFromDate, "yyyymm")) >= Format(Date, "yyyymm") Then
'   MsgBox "Present and Future Month not Allowed !!!"
'   CmbMonth.SetFocus
'   Exit Sub
'End If
'
'ProgressBar1.Value = 0
'
'CmdUpdatePendingRateCard.Enabled = False
'
'If LCase(CmbProcessType.Text) = "franchisee" Then
'
'   If Gen_DBType = "MDB" Then
'      tmp = " Select Distinct TIG.SM_Prefix From Txn_InvMonGSLDetail TIG"
'      tmp = tmp & " Inner join Mst_Godown MG on TIG.GodownID=MG.GodownID"
'      tmp = tmp & " Where MG.UnitTypeID in (2,6) And TIG.YearMonth = #" & mFromDate & "# And FranchProcessFlag ='Y' "
'   Else
'      tmp = " Select Distinct TIG.SM_Prefix From Txn_InvMonGSLDetail TIG"
'      tmp = tmp & " Inner join Mst_Godown MG on TIG.GodownID=MG.GodownID"
'      tmp = tmp & " Where MG.UnitTypeID in (2,6) And TIG.YearMonth = '" & mFromDate & "' And FranchProcessFlag ='Y'"
'   End If
'
'   Call TmpTable
'
'   For I = 1 To TmpRs.RecordCount
'
'       tmp = "Update Txn_InvMonGSLDetail Set Txn_InvMonGSLDetail.FF_RateCardID="
'       tmp = tmp & " (Select distinct MNR.FF_RateCardID"
'       tmp = tmp & " from  Mst_FF_RateCard_Godown  b"
'       tmp = tmp & " inner Join Mst_FF_RateCard MNR On b.FF_RateCardID = MNR.FF_RateCardID"
'       tmp = tmp & " Where (MNR.CommodityGroupID = c.CommodityGroupID Or b.BillingUnit = 'F') And (MNR.exchangetypeid = Txn_InvMonGSLDetail.exchangetypeid Or b.BillingUnit = 'F')"
'       tmp = tmp & " And b.GodownID=Txn_InvMonGSLDetail.GodownID)"
'       tmp = tmp & " From Txn_InvMonGSLDetail"
'       tmp = tmp & " inner join mst_commodity c on Txn_InvMonGSLDetail.commodityid = c.commodityid"
'       tmp = tmp & " Where Txn_InvMonGSLDetail.YearMonth = '" & mFromDate & "' And Txn_InvMonGSLDetail.SM_Prefix = '" & TmpRs!SM_Prefix & "' "
'       tmp = tmp & " And Txn_InvMonGSLDetail.FranchProcessFlag='Y' And Txn_InvMonGSLDetail.FF_RateCardID is Null"
'
'       Call Tmp2Table
'
'       tmp = "Update Txn_InvMonGSLDetail Set Txn_InvMonGSLDetail.FranchBillingCycleID ="
'       tmp = tmp & " (Select top 1 b.BillingCycleID"
'       tmp = tmp & " from  Mst_FF_RateCard_Godown  b"
'       tmp = tmp & " inner Join Mst_FF_RateCard MNR On b.FF_RateCardID = MNR.FF_RateCardID"
'       tmp = tmp & " Where (MNR.CommodityGroupID = c.CommodityGroupID Or b.BillingUnit = 'F') And (MNR.exchangetypeid = Txn_InvMonGSLDetail.exchangetypeid Or b.BillingUnit = 'F')"
'       tmp = tmp & " And b.GodownID=Txn_InvMonGSLDetail.GodownID)"
'       tmp = tmp & " From Txn_InvMonGSLDetail"
'       tmp = tmp & " inner join mst_commodity c on Txn_InvMonGSLDetail.commodityid = c.commodityid"
'       tmp = tmp & " Where Txn_InvMonGSLDetail.YearMonth = '" & mFromDate & "' And Txn_InvMonGSLDetail.SM_Prefix = '" & TmpRs!SM_Prefix & "'"
'       tmp = tmp & " And Txn_InvMonGSLDetail.FranchProcessFlag='Y' And Txn_InvMonGSLDetail.BillingCycleID is Null"
'
'       Call Tmp2Table
'
'       tmp = " Update Txn_InvMonGSLDetail Set Txn_InvMonGSLDetail.FranchBillingUnit ="
'       tmp = tmp & " (Select b.BillingUnit"
'       tmp = tmp & " from  Mst_FF_RateCard_Godown  b"
'       tmp = tmp & " inner Join Mst_FF_RateCard MNR On b.FF_RateCardID = MNR.FF_RateCardID"
'       tmp = tmp & " Where (MNR.CommodityGroupID = c.CommodityGroupID Or b.BillingUnit = 'F') And (MNR.exchangetypeid = Txn_InvMonGSLDetail.exchangetypeid Or b.BillingUnit = 'F')"
'       tmp = tmp & " And b.GodownID=Txn_InvMonGSLDetail.GodownID)"
'       tmp = tmp & " From Txn_InvMonGSLDetail"
'       tmp = tmp & " inner join mst_commodity c on Txn_InvMonGSLDetail.commodityid = c.commodityid"
'       tmp = tmp & " Where Txn_InvMonGSLDetail.YearMonth = '" & mFromDate & "' And Txn_InvMonGSLDetail.SM_Prefix = '" & TmpRs!SM_Prefix & "'"
'       tmp = tmp & " And Txn_InvMonGSLDetail.FranchProcessFlag='Y' And Txn_InvMonGSLDetail.FranchBillingUnit is Null"
'
'       Call Tmp2Table
'
'       tmp = " Update Txn_InvMonGSLDetail Set Txn_InvMonGSLDetail.FranchiseeRate ="
'       tmp = tmp & " (Select top 1 b.RateOnQty"
'       tmp = tmp & " from  Mst_FF_RateCard_Godown  b"
'       tmp = tmp & " inner Join Mst_FF_RateCard MNR On b.FF_RateCardID = MNR.FF_RateCardID"
'       tmp = tmp & " Where (MNR.CommodityGroupID = c.CommodityGroupID Or b.BillingUnit = 'F') And (MNR.exchangetypeid = Txn_InvMonGSLDetail.exchangetypeid Or b.BillingUnit = 'F')"
'       tmp = tmp & " And b.GodownID=Txn_InvMonGSLDetail.GodownID)"
'       tmp = tmp & " From Txn_InvMonGSLDetail"
'       tmp = tmp & " inner join mst_commodity c on Txn_InvMonGSLDetail.commodityid = c.commodityid"
'       tmp = tmp & " Where Txn_InvMonGSLDetail.YearMonth = '" & mFromDate & "' And Txn_InvMonGSLDetail.SM_Prefix = '" & TmpRs!SM_Prefix & "'"
'       tmp = tmp & " And Txn_InvMonGSLDetail.FranchProcessFlag='Y' And (Txn_InvMonGSLDetail.FranchiseeRate is Null Or Txn_InvMonGSLDetail.FranchiseeRate = 0 )"
'
'       Call Tmp2Table
'
'        tmp = "Update Txn_InvMonGSLDetail Set FranchiseeRate = 0"
'        tmp = tmp & " Where Txn_InvMonGSLDetail.YearMonth = '" & mFromDate & "' And Txn_InvMonGSLDetail.SM_Prefix = '" & TmpRs!SM_Prefix & "' And Txn_InvMonGSLDetail.FranchBillingUnit = 'F'"
'        tmp = tmp & " And GodownID in (Select distinct b.GodownID"
'        tmp = tmp & " from  Mst_FF_RateCard_Godown_Fix  MNRCB"
'        tmp = tmp & " inner Join Mst_FF_RateCard_Godown b On b.FF_RateCard_GodownID = MNRCB.FF_RateCard_GodownID"
'        tmp = tmp & " inner Join Mst_FF_RateCard MNR On b.FF_RateCardID = MNR.FF_RateCardID"
'        tmp = tmp & " inner Join Mst_Location Ml On Ml.LocationID= MNR.LocationID"
'        tmp = tmp & " Where b.BillingUnit='F'"
'        tmp = tmp & " And MNRCB.StartDate<='" & mToDate & "'  And MNRCB.EndDate>='" & mFromDate & "' And ML.SM_Prefix = '" & TmpRs!SM_Prefix & "')"
'        tmp = tmp & " And Txn_InvMonGSLDetail.FranchProcessFlag='Y' And (Txn_InvMonGSLDetail.FranchiseeRate is Null Or Txn_InvMonGSLDetail.FranchiseeRate = 0) "
'
'        Call Tmp2Table
'
'
'        tmp = "Select MNRCB.FixAmount,MNR.exchangetypeid,b.GodownID,MNRCB.StartDate,MNRCB.EndDate"
'        tmp = tmp & " from  Mst_FF_RateCard_Godown_Fix  MNRCB"
'        tmp = tmp & " inner Join Mst_FF_RateCard_Godown b On b.FF_RateCard_GodownID = MNRCB.FF_RateCard_GodownID"
'        tmp = tmp & " inner Join Mst_FF_RateCard MNR On b.FF_RateCardID = MNR.FF_RateCardID"
'        tmp = tmp & " inner Join Mst_Location Ml On Ml.LocationID= MNR.LocationID"
'        tmp = tmp & " Where b.BillingUnit='F'"
'        tmp = tmp & " And MNRCB.StartDate<='" & mToDate & "'  And MNRCB.EndDate>='" & mFromDate & "' And ML.SM_Prefix = '" & TmpRs!SM_Prefix & "' "
'
'        Call Tmp2Table
'
'        For rn = 1 To TmpRs2.RecordCount
'
'            If TmpRs2!StartDate <= CDate(mFromDate) And TmpRs2!EndDate >= CDate(mToDate) Then
'               mMonFixRate = TmpRs2!FixAmount
'            ElseIf TmpRs2!StartDate <= CDate(mFromDate) And TmpRs2!EndDate < CDate(mToDate) Then
'               mTotDays = DateDiff("d", CDate(mFromDate), TmpRs2!EndDate) + 1
'               mMonFixRate = (TmpRs2!FixAmount / 30) * mTotDays
'            ElseIf TmpRs2!StartDate > CDate(mFromDate) And TmpRs2!EndDate >= CDate(mToDate) Then
'               mTotDays = DateDiff("d", TmpRs2!StartDate, CDate(mToDate)) + 1
'               mMonFixRate = (TmpRs2!FixAmount / 30) * mTotDays
'            ElseIf TmpRs2!StartDate > CDate(mFromDate) And TmpRs2!EndDate < CDate(mToDate) Then
'               mTotDays = DateDiff("d", TmpRs2!StartDate, TmpRs2!EndDate) + 1
'               mMonFixRate = (TmpRs2!FixAmount / 30) * mTotDays
'            End If
'
'            tmp = "Update Txn_InvMonGSLDetail Set FranchiseeRate = isnull(FranchiseeRate,0) + " & Format(mMonFixRate, "#####0.00")
'            tmp = tmp & " Where Txn_InvMonGSLDetail.YearMonth = '" & mFromDate & "' And Txn_InvMonGSLDetail.SM_Prefix = '" & TmpRs!SM_Prefix & "' And Txn_InvMonGSLDetail.FranchBillingUnit = 'F'"
'            tmp = tmp & " And GodownID=" & TmpRs2!GodownID & ""
'            tmp = tmp & " And Txn_InvMonGSLDetail.FranchProcessFlag='Y' " 'And (Txn_InvMonGSLDetail.FranchiseeRate is Null Or Txn_InvMonGSLDetail.FranchiseeRate = 0) "
'
'            Call Tmp3Table
'
'            TmpRs2.MoveNext
'        Next
'
'
'       tmp = " Update Txn_InvMonGSLDetail Set Txn_InvMonGSLDetail.FranchiseeRate ="
'       tmp = tmp & " (Select MNRCB.RatePerBag"
'       tmp = tmp & " from  Mst_FF_RateCard_Godown_Bag  MNRCB"
'       tmp = tmp & " inner Join Mst_FF_RateCard_Godown b On b.FF_RateCard_GodownID = MNRCB.FF_RateCard_GodownID"
'       tmp = tmp & " inner Join Mst_FF_RateCard MNR On b.FF_RateCardID = MNR.FF_RateCardID"
'       tmp = tmp & " Where (MNR.CommodityGroupID = c.CommodityGroupID Or b.BillingUnit = 'F') And (MNR.exchangetypeid = Txn_InvMonGSLDetail.exchangetypeid Or b.BillingUnit = 'F')"
'       tmp = tmp & " And b.GodownID=Txn_InvMonGSLDetail.GodownID"
'       tmp = tmp & " And Txn_InvMonGSLDetail.BagSize >= MNRCB.BagSizeFrom And Txn_InvMonGSLDetail.BagSize <= MNRCB.BagSizeTo  )"
'       tmp = tmp & " From Txn_InvMonGSLDetail"
'       tmp = tmp & " inner join mst_commodity c on Txn_InvMonGSLDetail.commodityid = c.commodityid"
'       tmp = tmp & " Where Txn_InvMonGSLDetail.YearMonth = '" & mFromDate & "' And Txn_InvMonGSLDetail.SM_Prefix = '" & TmpRs!SM_Prefix & "' And Txn_InvMonGSLDetail.FranchBillingUnit = 'B'"
'       tmp = tmp & " And Txn_InvMonGSLDetail.FranchProcessFlag='Y' And (Txn_InvMonGSLDetail.FranchiseeRate is Null Or Txn_InvMonGSLDetail.FranchiseeRate = 0) "
'
'       Call Tmp2Table
'
'       'Op Bal and Monthly Cycle
'       tmp = " Update Txn_InvMonGSLDetail Set FranchNoofDays = 1 Where YearMonth = '" & mFromDate & "' And TFlag = 'O'"
'       tmp = tmp & " And FranchBillingCycleID = 4  And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
'       tmp = tmp & " And Txn_InvMonGSLDetail.FranchProcessFlag='Y' And Txn_InvMonGSLDetail.FranchNoofDays is Null"
'
'       Call Tmp2Table
'
'        'Op Bal and Forthnightly Cycle
'       tmp = " Update Txn_InvMonGSLDetail Set FranchNoofDays = 2 Where YearMonth = '" & mFromDate & "' And TFlag = 'O'"
'       tmp = tmp & " And FranchBillingCycleID = 3 And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
'       tmp = tmp & " And Txn_InvMonGSLDetail.FranchProcessFlag='Y' And Txn_InvMonGSLDetail.FranchNoofDays is Null"
'
'       Call Tmp2Table
'
'        'Op Bal and Weekly Cycle
'       tmp = " Update Txn_InvMonGSLDetail Set FranchNoofDays = 4 Where YearMonth = '" & mFromDate & "' And TFlag = 'O'"
'       tmp = tmp & " And FranchBillingCycleID = 2 And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
'       tmp = tmp & " And Txn_InvMonGSLDetail.FranchProcessFlag='Y' And Txn_InvMonGSLDetail.FranchNoofDays is Null"
'       Call Tmp2Table
'
'        'ProgressBar1.Value = ProgressBar1.Value + 1
'
'        'Op Bal and Daily Cycle
'       tmp = " Update Txn_InvMonGSLDetail Set FranchNoofDays = " & mTotalDays & " Where YearMonth = '" & mFromDate & "' And TFlag = 'O'"
'       tmp = tmp & " And FranchBillingCycleID = 1 And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
'       tmp = tmp & " And Txn_InvMonGSLDetail.FranchProcessFlag='Y' And Txn_InvMonGSLDetail.FranchNoofDays is Null "
'
'       Call Tmp2Table
'
'        'ProgressBar1.Value = ProgressBar1.Value + 1
'
'        'Depo and Withd and Monthly Cycle
'       tmp = " Update Txn_InvMonGSLDetail Set FranchNoofDays = 1 Where YearMonth = '" & mFromDate & "' And TFlag <> 'O'"
'       tmp = tmp & " And FranchBillingCycleID = 4 And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
'       tmp = tmp & " And Txn_InvMonGSLDetail.FranchProcessFlag='Y' And Txn_InvMonGSLDetail.FranchNoofDays is Null "
'
'       Call Tmp2Table
'
'       'ProgressBar1.Value = ProgressBar1.Value + 1
'
'       'Depo and Withd and Daily Cycle
'       tmp = "Update Txn_InvMonGSLDetail Set FranchNoofDays = " & mTotalDays & " - Day(CxTFDate) + 1 Where YearMonth = '" & mFromDate & "' And TFlag = 'D' And FranchBillingCycleID = 1 And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
'       tmp = tmp & " And Txn_InvMonGSLDetail.FranchProcessFlag='Y' And Txn_InvMonGSLDetail.FranchNoofDays is Null "
'
'       Call Tmp2Table
'
'        'ProgressBar1.Value = ProgressBar1.Value + 1
'
'        'Withd and Daily Cycle
'       tmp = "Update Txn_InvMonGSLDetail Set FranchNoOfDays = Day(CxTFDate) Where YearMonth = '" & mFromDate & "' And TFlag = 'W' And FranchBillingCycleID = 1 And SM_Prefix = '" & TmpRs!SM_Prefix & "' "
'       tmp = tmp & " And Txn_InvMonGSLDetail.FranchProcessFlag='Y' And Txn_InvMonGSLDetail.FranchNoofDays is Null"
'
'       Call Tmp2Table
'
'        'ProgressBar1.Value = ProgressBar1.Value + 1
'
'        'Deposit & Weekly
'
'       mNoofdays = GetBillingCycleDays(2)
'
'       tmp = "Update Txn_InvMonGSLDetail Set FranchNoOfDays =  Case Floor((4 - (Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ") + 1))  When 0 then 1 else Floor((4 - (Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ") + 1)) End "
'       tmp = tmp & " Where Txn_InvMonGSLDetail.YearMonth = '" & mFromDate & "'"
'       tmp = tmp & " And Txn_InvMonGSLDetail.TFlag = 'D' And Txn_InvMonGSLDetail.FranchBillingCycleID = 2 And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
'       tmp = tmp & " And Txn_InvMonGSLDetail.FranchProcessFlag='Y' And Txn_InvMonGSLDetail.FranchNoofDays is Null"
'       Call Tmp2Table
'
'       'ProgressBar1.Value = ProgressBar1.Value + 1
'
'       'Deposit & FortNightly
'
'       mNoofdays = GetBillingCycleDays(3)
'
'       'Case 1 Depo Date  <= 15 Then No Of Cycle = 2
'
'       tmp = " Update Txn_InvMonGSLDetail Set FranchNoOfDays = 2 "
'       tmp = tmp & " Where Txn_InvMonGSLDetail.YearMonth = '" & mFromDate & "' "
'       tmp = tmp & " And Txn_InvMonGSLDetail.TFlag = 'D' And Txn_InvMonGSLDetail.FranchBillingCycleID = 3 And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
'       tmp = tmp & " And Day(CxTFDate) <= 15 "
'       tmp = tmp & " And Txn_InvMonGSLDetail.FranchProcessFlag='Y' And Txn_InvMonGSLDetail.FranchNoofDays is Null"
'       Call Tmp2Table
'
'       'ProgressBar1.Value = ProgressBar1.Value + 1
'
'       'Case 1 Depo Date  > 15 Then No Of Cycle = 1
'       tmp = " Update Txn_InvMonGSLDetail Set FranchNoOfDays = 1 "
'       tmp = tmp & " Where YearMonth = '" & mFromDate & "' "
'       tmp = tmp & " And TFlag = 'D' And FranchBillingCycleID = 3 "
'       tmp = tmp & " And Day(CxTFDate) > 15 And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
'       tmp = tmp & " And Txn_InvMonGSLDetail.FranchProcessFlag='Y' And Txn_InvMonGSLDetail.FranchNoofDays is Null "
'       Call Tmp2Table
'
'       'ProgressBar1.Value = ProgressBar1.Value + 1
'
'       'Withdrawal & weekly
'       mNoofdays = GetBillingCycleDays(2)
'
'       tmp = "Update Txn_InvMonGSLDetail Set FranchNoOfDays = Case CEILING(((Round(Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ",1)))) When 5 Then 4 Else CEILING(((Round(Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ",1)))) End "
'       tmp = tmp & " Where YearMonth = '" & mFromDate & "' "
'       tmp = tmp & " And TFlag = 'W' And FranchBillingCycleID = 2 And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
'       tmp = tmp & " And Txn_InvMonGSLDetail.FranchProcessFlag='Y' And Txn_InvMonGSLDetail.FranchNoofDays is Null"
'       Call Tmp2Table
'
'       'ProgressBar1.Value = ProgressBar1.Value + 1
'
'       'Withdrawal & Fortnightly
'       mNoofdays = GetBillingCycleDays(3)
'
'       tmp = "Update Txn_InvMonGSLDetail Set FranchNoOfDays = Case CEILING(((Round(Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ",1)))) When 3 Then 2 Else CEILING(((Round(Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ",1)))) End "
'       tmp = tmp & " Where YearMonth = '" & mFromDate & "' "
'       tmp = tmp & " And TFlag = 'W' And FranchBillingCycleID = 3 And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
'       tmp = tmp & " And Txn_InvMonGSLDetail.FranchProcessFlag='Y' And Txn_InvMonGSLDetail.FranchNoofDays is Null"
'       Call Tmp2Table
'
'       'ProgressBar1.Value = ProgressBar1.Value + 1
'
'       'Updating Invoice Amount And Invoice Date
'       tmp = " Update Txn_InvMonGSLDetail Set FranchInvoiceAmount = BalanceBags * FranchNoOfDays * FranchiseeRate, InvoiceDate = '" & mToDate & "' "
'       tmp = tmp & " Where YearMonth = '" & mFromDate & "'"
'       tmp = tmp & " And FranchBillingUnit = 'B' And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
'       tmp = tmp & " And Txn_InvMonGSLDetail.FranchProcessFlag='Y' And Txn_InvMonGSLDetail.FranchInvoiceAmount is Null"
'       Call Tmp2Table
'
'       'ProgressBar1.Value = ProgressBar1.Value + 1
'
'       tmp = " Update Txn_InvMonGSLDetail Set FranchInvoiceAmount = BalanceWeight * FranchNoOfDays * FranchiseeRate,InvoiceDate =  '" & mToDate & "' "
'       tmp = tmp & " Where YearMonth = '" & mFromDate & "'"
'       tmp = tmp & " And FranchBillingUnit = 'Q' And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
'       tmp = tmp & " And Txn_InvMonGSLDetail.FranchProcessFlag='Y' And Txn_InvMonGSLDetail.FranchInvoiceAmount is Null"
'       Call Tmp2Table
'        '--For Fixed
'
'       tmp = "Select GodownID,FranchiseeRate,max(GSLID)GSLID,FF_RateCardID from Txn_InvMonGSLDetail"
'       tmp = tmp & " Where YearMonth = '" & mFromDate & "' And FranchBillingUnit = 'F' And SM_Prefix = '" & TmpRs!SM_Prefix & "' And GSLID>0 "
'       tmp = tmp & " Group by GodownID,FranchiseeRate,FF_RateCardID"
'       Call Tmp2Table
'
'       For rn = 1 To TmpRs2.RecordCount
'
'           tmp = "Update Txn_InvMonGSLDetail Set FranchInvoiceAmount = " & IIf(IsNull(TmpRs2!FranchiseeRate), 0, TmpRs2!FranchiseeRate) & ",InvoiceDate =  '" & mToDate & "'"
'           tmp = tmp & " Where YearMonth = '" & mFromDate & "' And GSLID=" & TmpRs2!GSLID & " And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
'           tmp = tmp & " And Txn_InvMonGSLDetail.FranchProcessFlag='Y' And Txn_InvMonGSLDetail.FranchInvoiceAmount is Null "
'           Call Tmp4Table
'
'           tmp = "Select GSLID,FranchInvoiceAmount,BalanceBags,MonG_UID from  Txn_InvMonGSLDetail "
'           tmp = tmp & " Where YearMonth = '" & mFromDate & "' And GSLID=" & TmpRs2!GSLID & " And SM_Prefix = '" & TmpRs!SM_Prefix & "' And Txn_InvMonGSLDetail.FranchInvoiceAmount is Null and Txn_InvMonGSLDetail.FranchProcessFlag='Y'"
'
'           Call Tmp3Table
'
'           If TmpRs3.RecordCount > 1 Then
'              For rn1 = 1 To TmpRs3.RecordCount - 1
'                  TmpRs3!FranchInvoiceAmount = 0
'                  TmpRs3.Update
'
'                  TmpRs3.MoveNext
'              Next
'           End If
'
'           tmp = "Update Txn_InvMonGSLDetail Set FF_RateCardID=" & TmpRs2!FF_RateCardID & ",FranchBillingUnit='F',FranchNoofDays=1,FranchiseeRate=" & IIf(IsNull(TmpRs2!FranchiseeRate), 0, TmpRs2!FranchiseeRate) & ",FranchInvoiceAmount=0,FranchBillingCycleID=4"
'           tmp = tmp & " Where YearMonth = '" & mFromDate & "' And SM_Prefix = '" & TmpRs!SM_Prefix & "' And GodownID=" & TmpRs2!GodownID & " And GSLID<>" & TmpRs2!GSLID & " And Txn_InvMonGSLDetail.FranchInvoiceAmount is Null and Txn_InvMonGSLDetail.FranchProcessFlag='Y'"
'
'           Call Tmp4Table
'
'           TmpRs2.MoveNext
'       Next
'
'        '--For Percent
'
'       tmp = " Update Txn_InvMonGSLDetail Set FranchInvoiceAmount =  (InvoiceAmount)* FranchiseeRate /100 ,InvoiceDate =  '" & mToDate & "' "
'       tmp = tmp & " Where YearMonth = '" & mFromDate & "'"
'       tmp = tmp & " And FranchBillingUnit = 'P' And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
'       tmp = tmp & " And Txn_InvMonGSLDetail.FranchProcessFlag='Y' And Txn_InvMonGSLDetail.FranchInvoiceAmount is Null"
'       Call Tmp2Table
'
'       TmpRs.MoveNext
'   Next
'
'End If
'MsgBox "Done!!!"
'Call Grid_Head
'End Sub

Private Sub CmdUpdatePendingRateCard_Click()

If CmbProcessType.Text = "" Then
   MsgBox " Blank Process Type not allowed  !!!"
   CmbProcessType.SetFocus
   Exit Sub
End If

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

mMonth = mMonthListindex + 1

If mMonth < 10 Then
   mMonth = "0" & mMonth
End If

mFromDate = Format(CmbYear.Text & "-" & mMonth & "-" & "01", "yyyy-mm-dd")
mYear = Val(CmbYear.Text)
mTotalDays = Day(DateSerial(mYear, mMonth + 1, 0))

mToDate = Format(CmbYear.Text & "-" & mMonth & "-" & mTotalDays, "yyyy-mm-dd")
mNoofDays = ""

If CDate(Format(mFromDate, "yyyymm")) >= Format(Date, "yyyymm") Then
   MsgBox "Present and Future Month not Allowed !!!"
   CmbMonth.SetFocus
   Exit Sub
End If

If Format(Gen_BlockingDate, "yyyy-mm") >= Format(mFromDate, "yyyy-mm") Then
   MsgBox "Select next month from " & Format(Gen_BlockingDate, "MMM-yyyy") & "!!! "
   CmbMonth.SetFocus
   Exit Sub
End If

ProgressBar1.Value = 0

CmdUpdatePendingRateCard.Enabled = False

If LCase(CmbProcessType.Text) = "franchisee" Then

   If Gen_DBType = "MDB" Then
      tmp = " Select Distinct TIG.SM_Prefix From Txn_InvMonGSLDetail TIG"
      tmp = tmp & " Inner join Mst_Godown MG on TIG.GodownID=MG.GodownID"
      tmp = tmp & " Where MG.UnitTypeID in (2,6) And TIG.YearMonth = #" & mFromDate & "# And FranchProcessFlag ='Y' "
   Else
      tmp = " Select Distinct TIG.SM_Prefix From Txn_InvMonGSLDetail TIG"
      tmp = tmp & " Inner join Mst_Godown MG on TIG.GodownID=MG.GodownID"
      tmp = tmp & " Where MG.UnitTypeID in (2,6) And TIG.YearMonth = '" & mFromDate & "' And FranchProcessFlag ='Y'"
   End If
    
   Call TmpTable
      
   For I = 1 To TmpRs.RecordCount
   
        tmp = "SET IMPLICIT_TRANSACTIONS ON"
        Call Tmp2Table
   
       tmp = "Update Txn_InvMonGSLDetail Set Txn_InvMonGSLDetail.FF_RateCardID="
       tmp = tmp & " (Select distinct MNR.FF_RateCardID"
       tmp = tmp & " from  Mst_FF_RateCard_Godown  b"
       tmp = tmp & " inner Join Mst_FF_RateCard MNR On b.FF_RateCardID = MNR.FF_RateCardID"
       tmp = tmp & " Where (MNR.CommodityGroupID = c.CommodityGroupID Or b.BillingUnit = 'F') And (MNR.exchangetypeid = Txn_InvMonGSLDetail.exchangetypeid Or b.BillingUnit = 'F')"
       tmp = tmp & " And b.GodownID=Txn_InvMonGSLDetail.GodownID  And MNR.Flag='A')"
       tmp = tmp & " From Txn_InvMonGSLDetail"
       tmp = tmp & " inner join mst_commodity c on Txn_InvMonGSLDetail.commodityid = c.commodityid"
       tmp = tmp & " Where Txn_InvMonGSLDetail.YearMonth = '" & mFromDate & "' And Txn_InvMonGSLDetail.SM_Prefix = '" & TmpRs!SM_Prefix & "' "
       tmp = tmp & " And Txn_InvMonGSLDetail.FranchProcessFlag='Y' And Txn_InvMonGSLDetail.FF_RateCardID is Null"
        
       Call Tmp2Table
        
        tmp = "SET IMPLICIT_TRANSACTIONS ON"
        Call Tmp2Table
        
        
       tmp = "Update Txn_InvMonGSLDetail Set Txn_InvMonGSLDetail.FranchBillingCycleID ="
       tmp = tmp & " (Select top 1 b.BillingCycleID"
       tmp = tmp & " from  Mst_FF_RateCard_Godown  b"
       tmp = tmp & " inner Join Mst_FF_RateCard MNR On b.FF_RateCardID = MNR.FF_RateCardID"
       tmp = tmp & " Where (MNR.CommodityGroupID = c.CommodityGroupID Or b.BillingUnit = 'F') And (MNR.exchangetypeid = Txn_InvMonGSLDetail.exchangetypeid Or b.BillingUnit = 'F')"
       tmp = tmp & " And b.GodownID=Txn_InvMonGSLDetail.GodownID  And MNR.Flag='A')"
       tmp = tmp & " From Txn_InvMonGSLDetail"
       tmp = tmp & " inner join mst_commodity c on Txn_InvMonGSLDetail.commodityid = c.commodityid"
       tmp = tmp & " Where Txn_InvMonGSLDetail.YearMonth = '" & mFromDate & "' And Txn_InvMonGSLDetail.SM_Prefix = '" & TmpRs!SM_Prefix & "'"
       tmp = tmp & " And Txn_InvMonGSLDetail.FranchProcessFlag='Y' And Txn_InvMonGSLDetail.FranchBillingCycleID is Null"
       
       Call Tmp2Table
        
        tmp = "SET IMPLICIT_TRANSACTIONS ON"
        Call Tmp2Table
        
        
       tmp = " Update Txn_InvMonGSLDetail Set Txn_InvMonGSLDetail.FranchBillingUnit ="
       tmp = tmp & " (Select b.BillingUnit"
       tmp = tmp & " from  Mst_FF_RateCard_Godown  b"
       tmp = tmp & " inner Join Mst_FF_RateCard MNR On b.FF_RateCardID = MNR.FF_RateCardID"
       tmp = tmp & " Where (MNR.CommodityGroupID = c.CommodityGroupID Or b.BillingUnit = 'F') And (MNR.exchangetypeid = Txn_InvMonGSLDetail.exchangetypeid Or b.BillingUnit = 'F')"
       tmp = tmp & " And b.GodownID=Txn_InvMonGSLDetail.GodownID  And MNR.Flag='A')"
       tmp = tmp & " From Txn_InvMonGSLDetail"
       tmp = tmp & " inner join mst_commodity c on Txn_InvMonGSLDetail.commodityid = c.commodityid"
       tmp = tmp & " Where Txn_InvMonGSLDetail.YearMonth = '" & mFromDate & "' And Txn_InvMonGSLDetail.SM_Prefix = '" & TmpRs!SM_Prefix & "'"
       tmp = tmp & " And Txn_InvMonGSLDetail.FranchProcessFlag='Y' And Txn_InvMonGSLDetail.FranchBillingUnit is Null"
       
       Call Tmp2Table
        
        tmp = "SET IMPLICIT_TRANSACTIONS ON"
        Call Tmp2Table
        
        
       tmp = " Update Txn_InvMonGSLDetail Set Txn_InvMonGSLDetail.FranchiseeRate ="
       tmp = tmp & " (Select top 1 b.RateOnQty"
       tmp = tmp & " from  Mst_FF_RateCard_Godown  b"
       tmp = tmp & " inner Join Mst_FF_RateCard MNR On b.FF_RateCardID = MNR.FF_RateCardID"
       tmp = tmp & " Where (MNR.CommodityGroupID = c.CommodityGroupID Or b.BillingUnit = 'F') And (MNR.exchangetypeid = Txn_InvMonGSLDetail.exchangetypeid Or b.BillingUnit = 'F')"
       tmp = tmp & " And b.GodownID=Txn_InvMonGSLDetail.GodownID  And MNR.Flag='A')"
       tmp = tmp & " From Txn_InvMonGSLDetail"
       tmp = tmp & " inner join mst_commodity c on Txn_InvMonGSLDetail.commodityid = c.commodityid"
       tmp = tmp & " Where Txn_InvMonGSLDetail.YearMonth = '" & mFromDate & "' And Txn_InvMonGSLDetail.SM_Prefix = '" & TmpRs!SM_Prefix & "'"
       tmp = tmp & " And Txn_InvMonGSLDetail.FranchProcessFlag='Y' And (Txn_InvMonGSLDetail.FranchiseeRate is Null Or Txn_InvMonGSLDetail.FranchiseeRate = 0 )"
       
       Call Tmp2Table
       
        tmp = "SET IMPLICIT_TRANSACTIONS ON"
        Call Tmp2Table
       
       
        tmp = "Update Txn_InvMonGSLDetail Set FranchiseeRate = 0"
        tmp = tmp & " Where Txn_InvMonGSLDetail.YearMonth = '" & mFromDate & "' And Txn_InvMonGSLDetail.SM_Prefix = '" & TmpRs!SM_Prefix & "' And Txn_InvMonGSLDetail.FranchBillingUnit = 'F'"
        tmp = tmp & " And GodownID in (Select distinct b.GodownID"
        tmp = tmp & " from  Mst_FF_RateCard_Godown_Fix  MNRCB"
        tmp = tmp & " inner Join Mst_FF_RateCard_Godown b On b.FF_RateCard_GodownID = MNRCB.FF_RateCard_GodownID"
        tmp = tmp & " inner Join Mst_FF_RateCard MNR On b.FF_RateCardID = MNR.FF_RateCardID"
        tmp = tmp & " inner Join Mst_Location Ml On Ml.LocationID= MNR.LocationID"
        tmp = tmp & " Where b.BillingUnit='F'"
        tmp = tmp & " And MNRCB.StartDate<='" & mToDate & "'  And MNRCB.EndDate>='" & mFromDate & "' And ML.SM_Prefix = '" & TmpRs!SM_Prefix & "'  And MNR.Flag='A')"
        tmp = tmp & " And Txn_InvMonGSLDetail.FranchProcessFlag='Y' And (Txn_InvMonGSLDetail.FranchiseeRate is Null Or Txn_InvMonGSLDetail.FranchiseeRate = 0) "
        
        Call Tmp2Table
        
        tmp = "SET IMPLICIT_TRANSACTIONS ON"
        Call Tmp2Table
        
       
        tmp = "Select MNRCB.FixAmount,MNR.exchangetypeid,b.GodownID,MNRCB.StartDate,MNRCB.EndDate"
        tmp = tmp & " from  Mst_FF_RateCard_Godown_Fix  MNRCB"
        tmp = tmp & " inner Join Mst_FF_RateCard_Godown b On b.FF_RateCard_GodownID = MNRCB.FF_RateCard_GodownID"
        tmp = tmp & " inner Join Mst_FF_RateCard MNR On b.FF_RateCardID = MNR.FF_RateCardID"
        tmp = tmp & " inner Join Mst_Location Ml On Ml.LocationID= MNR.LocationID"
        tmp = tmp & " Where b.BillingUnit='F'"
        tmp = tmp & " And MNRCB.StartDate<='" & mToDate & "'  And MNRCB.EndDate>='" & mFromDate & "' And ML.SM_Prefix = '" & TmpRs!SM_Prefix & "'  And MNR.Flag='A'"
        
        Call Tmp2Table
            
        For rn = 1 To TmpRs2.RecordCount
        
            If TmpRs2!StartDate <= CDate(mFromDate) And TmpRs2!EndDate >= CDate(mToDate) Then
               mMonFixRate = TmpRs2!FixAmount
            ElseIf TmpRs2!StartDate <= CDate(mFromDate) And TmpRs2!EndDate < CDate(mToDate) Then
               mTotDays = DateDiff("d", CDate(mFromDate), TmpRs2!EndDate) + 1
               mMonFixRate = (TmpRs2!FixAmount / 30) * mTotDays
            ElseIf TmpRs2!StartDate > CDate(mFromDate) And TmpRs2!EndDate >= CDate(mToDate) Then
               mTotDays = DateDiff("d", TmpRs2!StartDate, CDate(mToDate)) + 1
               mMonFixRate = (TmpRs2!FixAmount / 30) * mTotDays
            ElseIf TmpRs2!StartDate > CDate(mFromDate) And TmpRs2!EndDate < CDate(mToDate) Then
               mTotDays = DateDiff("d", TmpRs2!StartDate, TmpRs2!EndDate) + 1
               mMonFixRate = (TmpRs2!FixAmount / 30) * mTotDays
            End If
            
            tmp = "SET IMPLICIT_TRANSACTIONS ON"
            Call Tmp3Table
            
            
            tmp = "Update Txn_InvMonGSLDetail Set FranchiseeRate = isnull(FranchiseeRate,0) + " & Format(mMonFixRate, "#####0.00")
            tmp = tmp & " Where Txn_InvMonGSLDetail.YearMonth = '" & mFromDate & "' And Txn_InvMonGSLDetail.SM_Prefix = '" & TmpRs!SM_Prefix & "' And Txn_InvMonGSLDetail.FranchBillingUnit = 'F'"
            tmp = tmp & " And GodownID=" & TmpRs2!GodownID & ""
            tmp = tmp & " And Txn_InvMonGSLDetail.FranchProcessFlag='Y' " 'And (Txn_InvMonGSLDetail.FranchiseeRate is Null Or Txn_InvMonGSLDetail.FranchiseeRate = 0) "
            
            Call Tmp3Table
               
            TmpRs2.MoveNext
        Next
        
        tmp = "SET IMPLICIT_TRANSACTIONS ON"
        Call Tmp2Table
        

       tmp = " Update Txn_InvMonGSLDetail Set Txn_InvMonGSLDetail.FranchiseeRate ="
       tmp = tmp & " (Select MNRCB.RatePerBag"
       tmp = tmp & " from  Mst_FF_RateCard_Godown_Bag  MNRCB"
       tmp = tmp & " inner Join Mst_FF_RateCard_Godown b On b.FF_RateCard_GodownID = MNRCB.FF_RateCard_GodownID"
       tmp = tmp & " inner Join Mst_FF_RateCard MNR On b.FF_RateCardID = MNR.FF_RateCardID"
       tmp = tmp & " Where (MNR.CommodityGroupID = c.CommodityGroupID Or b.BillingUnit = 'F') And (MNR.exchangetypeid = Txn_InvMonGSLDetail.exchangetypeid Or b.BillingUnit = 'F')"
       tmp = tmp & " And b.GodownID=Txn_InvMonGSLDetail.GodownID"
       tmp = tmp & " And Txn_InvMonGSLDetail.BagSize >= MNRCB.BagSizeFrom And Txn_InvMonGSLDetail.BagSize <= MNRCB.BagSizeTo  And MNR.Flag='A')"
       tmp = tmp & " From Txn_InvMonGSLDetail"
       tmp = tmp & " inner join mst_commodity c on Txn_InvMonGSLDetail.commodityid = c.commodityid"
       tmp = tmp & " Where Txn_InvMonGSLDetail.YearMonth = '" & mFromDate & "' And Txn_InvMonGSLDetail.SM_Prefix = '" & TmpRs!SM_Prefix & "' And Txn_InvMonGSLDetail.FranchBillingUnit = 'B'"
       tmp = tmp & " And Txn_InvMonGSLDetail.FranchProcessFlag='Y' And (Txn_InvMonGSLDetail.FranchiseeRate is Null Or Txn_InvMonGSLDetail.FranchiseeRate = 0) "
       
       Call Tmp2Table
       
       tmp = "SET IMPLICIT_TRANSACTIONS ON"
       Call Tmp2Table
        
       'Op Bal and Monthly Cycle
       tmp = " Update Txn_InvMonGSLDetail Set FranchNoofDays = 1 Where YearMonth = '" & mFromDate & "' And TFlag = 'O'"
       tmp = tmp & " And FranchBillingCycleID = 4  And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
       tmp = tmp & " And Txn_InvMonGSLDetail.FranchProcessFlag='Y' And Txn_InvMonGSLDetail.FranchNoofDays is Null"
       
       Call Tmp2Table
           
       tmp = "SET IMPLICIT_TRANSACTIONS ON"
       Call Tmp2Table
        'Op Bal and Forthnightly Cycle
       tmp = " Update Txn_InvMonGSLDetail Set FranchNoofDays = 2 Where YearMonth = '" & mFromDate & "' And TFlag = 'O'"
       tmp = tmp & " And FranchBillingCycleID = 3 And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
       tmp = tmp & " And Txn_InvMonGSLDetail.FranchProcessFlag='Y' And Txn_InvMonGSLDetail.FranchNoofDays is Null"
       
       Call Tmp2Table
        
       tmp = "SET IMPLICIT_TRANSACTIONS ON"
       Call Tmp2Table
        'Op Bal and Weekly Cycle
       tmp = " Update Txn_InvMonGSLDetail Set FranchNoofDays = 4 Where YearMonth = '" & mFromDate & "' And TFlag = 'O'"
       tmp = tmp & " And FranchBillingCycleID = 2 And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
       tmp = tmp & " And Txn_InvMonGSLDetail.FranchProcessFlag='Y' And Txn_InvMonGSLDetail.FranchNoofDays is Null"
       Call Tmp2Table
        
        'ProgressBar1.Value = ProgressBar1.Value + 1
        
        'Op Bal and Daily Cycle
       tmp = "SET IMPLICIT_TRANSACTIONS ON"
       Call Tmp2Table
       
       
       tmp = " Update Txn_InvMonGSLDetail Set FranchNoofDays = " & mTotalDays & " Where YearMonth = '" & mFromDate & "' And TFlag = 'O'"
       tmp = tmp & " And FranchBillingCycleID = 1 And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
       tmp = tmp & " And Txn_InvMonGSLDetail.FranchProcessFlag='Y' And Txn_InvMonGSLDetail.FranchNoofDays is Null "
       
       Call Tmp2Table
         
        'ProgressBar1.Value = ProgressBar1.Value + 1
       tmp = "SET IMPLICIT_TRANSACTIONS ON"
       Call Tmp2Table
         
        'Depo and Withd and Monthly Cycle
       tmp = " Update Txn_InvMonGSLDetail Set FranchNoofDays = 1 Where YearMonth = '" & mFromDate & "' And TFlag <> 'O'"
       tmp = tmp & " And FranchBillingCycleID = 4 And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
       tmp = tmp & " And Txn_InvMonGSLDetail.FranchProcessFlag='Y' And Txn_InvMonGSLDetail.FranchNoofDays is Null "
          
       Call Tmp2Table
            
       'ProgressBar1.Value = ProgressBar1.Value + 1
       tmp = "SET IMPLICIT_TRANSACTIONS ON"
       Call Tmp2Table
        
       'Depo and Withd and Daily Cycle
       tmp = "Update Txn_InvMonGSLDetail Set FranchNoofDays = " & mTotalDays & " - Day(CxTFDate) + 1 Where YearMonth = '" & mFromDate & "' And TFlag = 'D' And FranchBillingCycleID = 1 And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
       tmp = tmp & " And Txn_InvMonGSLDetail.FranchProcessFlag='Y' And Txn_InvMonGSLDetail.FranchNoofDays is Null "
       
       Call Tmp2Table
            
        'ProgressBar1.Value = ProgressBar1.Value + 1
       tmp = "SET IMPLICIT_TRANSACTIONS ON"
       Call Tmp2Table
    
        'Withd and Daily Cycle
       tmp = "Update Txn_InvMonGSLDetail Set FranchNoOfDays = Day(CxTFDate) Where YearMonth = '" & mFromDate & "' And TFlag = 'W' And FranchBillingCycleID = 1 And SM_Prefix = '" & TmpRs!SM_Prefix & "' "
       tmp = tmp & " And Txn_InvMonGSLDetail.FranchProcessFlag='Y' And Txn_InvMonGSLDetail.FranchNoofDays is Null"
          
       Call Tmp2Table
    
        'ProgressBar1.Value = ProgressBar1.Value + 1
        
        'Deposit & Weekly
        
       mNoofDays = GetBillingCycleDays(2)
        
       tmp = "SET IMPLICIT_TRANSACTIONS ON"
       Call Tmp2Table
        
        
       tmp = "Update Txn_InvMonGSLDetail Set FranchNoOfDays =  Case Floor((4 - (Convert(Numeric,Day(CxTFDate))/" & mNoofDays & ") + 1))  When 0 then 1 else Floor((4 - (Convert(Numeric,Day(CxTFDate))/" & mNoofDays & ") + 1)) End "
       tmp = tmp & " Where Txn_InvMonGSLDetail.YearMonth = '" & mFromDate & "'"
       tmp = tmp & " And Txn_InvMonGSLDetail.TFlag = 'D' And Txn_InvMonGSLDetail.FranchBillingCycleID = 2 And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
       tmp = tmp & " And Txn_InvMonGSLDetail.FranchProcessFlag='Y' And Txn_InvMonGSLDetail.FranchNoofDays is Null"
       Call Tmp2Table
        
       'ProgressBar1.Value = ProgressBar1.Value + 1
            
       'Deposit & FortNightly
        
       mNoofDays = GetBillingCycleDays(3)
        
       'Case 1 Depo Date  <= 15 Then No Of Cycle = 2
       tmp = "SET IMPLICIT_TRANSACTIONS ON"
       Call Tmp2Table
        
        
       tmp = " Update Txn_InvMonGSLDetail Set FranchNoOfDays = 2 "
       tmp = tmp & " Where Txn_InvMonGSLDetail.YearMonth = '" & mFromDate & "' "
       tmp = tmp & " And Txn_InvMonGSLDetail.TFlag = 'D' And Txn_InvMonGSLDetail.FranchBillingCycleID = 3 And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
       tmp = tmp & " And Day(CxTFDate) <= 15 "
       tmp = tmp & " And Txn_InvMonGSLDetail.FranchProcessFlag='Y' And Txn_InvMonGSLDetail.FranchNoofDays is Null"
       Call Tmp2Table
              
       'ProgressBar1.Value = ProgressBar1.Value + 1
       tmp = "SET IMPLICIT_TRANSACTIONS ON"
       Call Tmp2Table
        
       'Case 1 Depo Date  > 15 Then No Of Cycle = 1
       tmp = " Update Txn_InvMonGSLDetail Set FranchNoOfDays = 1 "
       tmp = tmp & " Where YearMonth = '" & mFromDate & "' "
       tmp = tmp & " And TFlag = 'D' And FranchBillingCycleID = 3 "
       tmp = tmp & " And Day(CxTFDate) > 15 And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
       tmp = tmp & " And Txn_InvMonGSLDetail.FranchProcessFlag='Y' And Txn_InvMonGSLDetail.FranchNoofDays is Null "
       Call Tmp2Table
              
       'ProgressBar1.Value = ProgressBar1.Value + 1
        
       'Withdrawal & weekly
       mNoofDays = GetBillingCycleDays(2)
       
       tmp = "SET IMPLICIT_TRANSACTIONS ON"
       Call Tmp2Table
                   
       tmp = "Update Txn_InvMonGSLDetail Set FranchNoOfDays = Case CEILING(((Round(Convert(Numeric,Day(CxTFDate))/" & mNoofDays & ",1)))) When 5 Then 4 Else CEILING(((Round(Convert(Numeric,Day(CxTFDate))/" & mNoofDays & ",1)))) End "
       tmp = tmp & " Where YearMonth = '" & mFromDate & "' "
       tmp = tmp & " And TFlag = 'W' And FranchBillingCycleID = 2 And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
       tmp = tmp & " And Txn_InvMonGSLDetail.FranchProcessFlag='Y' And Txn_InvMonGSLDetail.FranchNoofDays is Null"
       Call Tmp2Table
        
       'ProgressBar1.Value = ProgressBar1.Value + 1
           
       'Withdrawal & Fortnightly
       mNoofDays = GetBillingCycleDays(3)
        
       tmp = "SET IMPLICIT_TRANSACTIONS ON"
       Call Tmp2Table
        
        
       tmp = "Update Txn_InvMonGSLDetail Set FranchNoOfDays = Case CEILING(((Round(Convert(Numeric,Day(CxTFDate))/" & mNoofDays & ",1)))) When 3 Then 2 Else CEILING(((Round(Convert(Numeric,Day(CxTFDate))/" & mNoofDays & ",1)))) End "
       tmp = tmp & " Where YearMonth = '" & mFromDate & "' "
       tmp = tmp & " And TFlag = 'W' And FranchBillingCycleID = 3 And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
       tmp = tmp & " And Txn_InvMonGSLDetail.FranchProcessFlag='Y' And Txn_InvMonGSLDetail.FranchNoofDays is Null"
       Call Tmp2Table
                
       'ProgressBar1.Value = ProgressBar1.Value + 1
        
       tmp = "SET IMPLICIT_TRANSACTIONS ON"
       Call Tmp2Table
        
        
       'Updating Invoice Amount And Invoice Date
       tmp = " Update Txn_InvMonGSLDetail Set FranchInvoiceAmount = BalanceBags * FranchNoOfDays * FranchiseeRate, InvoiceDate = '" & mToDate & "' "
       tmp = tmp & " Where YearMonth = '" & mFromDate & "'"
       tmp = tmp & " And FranchBillingUnit = 'B' And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
       tmp = tmp & " And Txn_InvMonGSLDetail.FranchProcessFlag='Y' And Txn_InvMonGSLDetail.FranchInvoiceAmount is Null"
       Call Tmp2Table
    
       'ProgressBar1.Value = ProgressBar1.Value + 1
    
       tmp = "SET IMPLICIT_TRANSACTIONS ON"
       Call Tmp2Table
    
    
       tmp = " Update Txn_InvMonGSLDetail Set FranchInvoiceAmount = BalanceWeight * FranchNoOfDays * FranchiseeRate,InvoiceDate =  '" & mToDate & "' "
       tmp = tmp & " Where YearMonth = '" & mFromDate & "'"
       tmp = tmp & " And FranchBillingUnit = 'Q' And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
       tmp = tmp & " And Txn_InvMonGSLDetail.FranchProcessFlag='Y' And Txn_InvMonGSLDetail.FranchInvoiceAmount is Null"
       Call Tmp2Table
        '--For Fixed
    
       tmp = "SET IMPLICIT_TRANSACTIONS ON"
       Call Tmp2Table
    
    
       tmp = "Select GodownID,FranchiseeRate,max(Mst_GSL_ID)GSLID,FF_RateCardID from Txn_InvMonGSLDetail"
       tmp = tmp & " Where YearMonth = '" & mFromDate & "' And FranchBillingUnit = 'F' And SM_Prefix = '" & TmpRs!SM_Prefix & "' And Mst_GSL_ID>0 "
       tmp = tmp & " Group by GodownID,FranchiseeRate,FF_RateCardID"
       Call Tmp2Table
            
       For rn = 1 To TmpRs2.RecordCount
       
           tmp = "SET IMPLICIT_TRANSACTIONS ON"
           Call Tmp4Table
       
       
           tmp = "Select * from Txn_InvMonGSLDetail Where YearMOnth = '" & mFromDate & "' And StackNo = 'Dummy-F' And Mst_GSL_ID = 0 And GodownID = " & TmpRs2!GodownID & ""
           Call Tmp4Table
            
           If TmpRs4.RecordCount = 0 Then
       
              tmp = "SET IMPLICIT_TRANSACTIONS ON"
              Call Tmp4Table
       
       
              tmp = "Update Txn_InvMonGSLDetail Set FranchInvoiceAmount = " & IIf(IsNull(TmpRs2!FranchiseeRate), 0, TmpRs2!FranchiseeRate) & ",InvoiceDate =  '" & mToDate & "'"
              tmp = tmp & " Where YearMonth = '" & mFromDate & "' And Mst_GSL_ID=" & TmpRs2!GSLID & " And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
              tmp = tmp & " And Txn_InvMonGSLDetail.FranchProcessFlag='Y' And Txn_InvMonGSLDetail.FranchInvoiceAmount is Null "
              Call Tmp4Table
                
              tmp = "SET IMPLICIT_TRANSACTIONS ON"
              Call Tmp4Table
                
                
              tmp = "Select Mst_GSL_ID,FranchInvoiceAmount,BalanceBags,MonG_UID from  Txn_InvMonGSLDetail "
              tmp = tmp & " Where YearMonth = '" & mFromDate & "' And Mst_GSL_ID=" & TmpRs2!GSLID & " And SM_Prefix = '" & TmpRs!SM_Prefix & "' And Txn_InvMonGSLDetail.FranchInvoiceAmount is Null and Txn_InvMonGSLDetail.FranchProcessFlag='Y'"
                 
              Call Tmp3Table
                 
              If TmpRs3.RecordCount > 1 Then
                 For rn1 = 1 To TmpRs3.RecordCount - 1
                     TmpRs3!FranchInvoiceAmount = 0
                     TmpRs3.Update
                     TmpRs3.MoveNext
                 Next
              End If
                
              tmp = "SET IMPLICIT_TRANSACTIONS ON"
              Call Tmp4Table
               
              tmp = "Update Txn_InvMonGSLDetail Set FF_RateCardID=" & TmpRs2!FF_RateCardID & ",FranchBillingUnit='F',FranchNoofDays=1,FranchiseeRate=" & IIf(IsNull(TmpRs2!FranchiseeRate), 0, TmpRs2!FranchiseeRate) & ",FranchInvoiceAmount=0,FranchBillingCycleID=4"
              tmp = tmp & " Where YearMonth = '" & mFromDate & "' And SM_Prefix = '" & TmpRs!SM_Prefix & "' And GodownID=" & TmpRs2!GodownID & " And Mst_GSL_ID<>" & TmpRs2!GSLID & " And Txn_InvMonGSLDetail.FranchInvoiceAmount is Null and Txn_InvMonGSLDetail.FranchProcessFlag='Y'"
                 
              Call Tmp4Table
           
           Else
             
              tmp = "SET IMPLICIT_TRANSACTIONS ON"
              Call Tmp4Table
             
             
              tmp = "Update Txn_InvMonGSLDetail Set FF_RateCardID=" & TmpRs2!FF_RateCardID & ",FranchBillingUnit='F',FranchNoofDays=1,FranchiseeRate=" & IIf(IsNull(TmpRs2!FranchiseeRate), 0, TmpRs2!FranchiseeRate) & ",FranchInvoiceAmount=0,FranchBillingCycleID=4"
              tmp = tmp & " Where YearMonth = '" & mFromDate & "' And SM_Prefix = '" & TmpRs!SM_Prefix & "' And GodownID=" & TmpRs2!GodownID & " And Txn_InvMonGSLDetail.FranchInvoiceAmount is Null and Txn_InvMonGSLDetail.Mst_GSL_ID > 0 and Txn_InvMonGSLDetail.FranchProcessFlag='Y'"
                 
              Call Tmp4Table
           
           End If
           
           TmpRs2.MoveNext
       Next
        
        
        '--For Percent
        
       tmp = "SET IMPLICIT_TRANSACTIONS ON"
       Call Tmp2Table
        
        
       tmp = " Update Txn_InvMonGSLDetail Set FranchInvoiceAmount =  (InvoiceAmount)* FranchiseeRate /100 ,InvoiceDate =  '" & mToDate & "' "
       tmp = tmp & " Where YearMonth = '" & mFromDate & "'"
       tmp = tmp & " And FranchBillingUnit = 'P' And SM_Prefix = '" & TmpRs!SM_Prefix & "'"
       tmp = tmp & " And Txn_InvMonGSLDetail.FranchProcessFlag='Y' And Txn_InvMonGSLDetail.FranchInvoiceAmount is Null"
       Call Tmp2Table
    
       TmpRs.MoveNext
   Next

End If
MsgBox "Done!!!"
Call Grid_Head
End Sub

Private Sub CmdDeleteFranchProcess_Click()

ProgressBar1.Value = 0

If CmbProcessType.Text = "" Then
   MsgBox " Blank Process Type not allowed  !!!"
   CmbProcessType.SetFocus
   Exit Sub
End If

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

mMonth = mMonthListindex + 1

If mMonth < 10 Then
   mMonth = "0" & mMonth
End If

mFromDate = Format(CmbYear.Text & "-" & mMonth & "-" & "01", "yyyy-mm-dd")
mYear = Val(CmbYear.Text)
mTotalDays = Day(DateSerial(mYear, mMonth + 1, 0))

mToDate = Format(CmbYear.Text & "-" & mMonth & "-" & mTotalDays, "yyyy-mm-dd")
mNoofDays = ""

If CDate(Format(mFromDate, "yyyymm")) >= Format(Date, "yyyymm") Then
   MsgBox "Present and Future Month not Allowed !!!"
   CmbMonth.SetFocus
   Exit Sub
End If

tmp = "Select * from Txn_FranchiseeRent Where ExpenseFrom = '" & mFromDate & "' And LocationID = " & mLocationID & " and Flag <> 'R'"
Call Tmp1Table
If TmpRs1.RecordCount > 0 Then
   MsgBox "Franchisee Claim ID present for rejection , please check!!!"
   CmbLocationID.SetFocus
   Exit Sub
End If

tmp = "UPDATE Txn_InvMonGSLDetail SET FranchBillingCycleID = Null,FranchBillingUnit = Null,FranchNoofDays = Null,FranchInvoiceAmount = Null,FF_RateCardID = Null,FranchClaimID = Null,FranchProcessFlag = 'N',FranchProcessdBy = Null,FranchProcessedDate = Null,InvoiceAmtForFranch = 0,InsuranceCostDeduction=0 WHere YearMonth = '" & mFromDate & "' And SM_Prefix = '" & mSM_Prefix & "'"
Call Tmp1Table

MsgBox "Franchisee Process Deleted!!!"
CmbLocationID.Text = ""

End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()

Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   'Call GButtonLockAD(Me)
   Frame1.Enabled = False
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If

Call Grid_Head
Call Grid_HeadGodownList
If LCase(Gen_UserRole) = "power" Or LCase(Gen_UserRole) = "headcmg" Or LCase(Gen_UserRole) = "superadmin" Then
   Frame4.Visible = True
   MSHFlexGrid2.Visible = True
   Frame5.Visible = True
Else
   Frame4.Visible = False
   MSHFlexGrid2.Visible = False
   Frame5.Visible = False
End If

If LCase(Gen_UserRole) = "power" Or LCase(Gen_UserRole) = "headcmg" Or LCase(Gen_UserRole) = "superadmin" Then
   CmdDPDistribute.Visible = True
   CmdCommtrack.Visible = True
'   CmdReserDistri.Visible = True
Else
   CmdDPDistribute.Visible = False
   CmdCommtrack.Visible = False
'   CmdReserDistri.Visible = False
End If
Frame1.Enabled = True
'CmdDelete.Visible = False
ChkCheckAll.Visible = False
CmdUpdatePendingRateCard.Visible = False
If LCase(Gen_UserRole) = "power" Then
'   CmdDelete.Visible = True
   ChkCheckAll.Visible = True
End If
mCountLocation = 0
'CmdProcess.Enabled = False
ProgressBar1.Value = 0

If LCase(Gen_UserRole) = "power" Or LCase(Gen_UserRole) = "admin" Or LCase(Gen_UserRole) = "superadmin" Or LCase(Gen_UserRole) = "headcmg" Then
   CmdUpdatePendingRateCard.Visible = True
   CmdUpdatePendingRateCard.Enabled = True
End If

End Sub


Private Sub CmbPreFix_GotFocus()
tmp = CmbPreFix.Text

If Gen_WcLocation = "" Then
   Call TableMst_Location("")
Else
   Call TableMst_Location("where " & Gen_WcLocation)
End If
CmbPreFix.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbPreFix.AddItem RsMst_Location!Locationname
    RsMst_Location.MoveNext
Next
CmbPreFix.Text = tmp
End Sub

Private Sub CmbPreFix_LostFocus()
If CmbPreFix.Text = "" Then
   mLocationID = 0
   mSM_Prefix = Null
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbPreFix.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection "
   CmbPreFix.SetFocus
   Exit Sub
End If

mLocationID = RsMst_Location!LocationID
mLocationName = RsMst_Location!Locationname
mSM_Prefix = RsMst_Location!SM_Prefix
End Sub

Private Sub Grid_Head()
With MSHFlexGrid1
     .Clear
     .Rows = 2
     .Cols = 3 '26 '27 '23 '25
     .AllowUserResizing = flexResizeBoth
     .Font.Size = 10
     .WordWrap = True
     .RowHeight(0) = 500

     .TextMatrix(0, 0) = "Select"
     .TextMatrix(0, 1) = "Location Name"
     .TextMatrix(0, 2) = "SM_Prefix"
          
     .ColWidth(0) = 1000
     .ColWidth(1) = 4600
     .ColWidth(2) = 1000
          
End With
End Sub

Private Function ChkSelection() As Boolean
Dim Retval As Boolean
Retval = False

For k = 1 To MSHFlexGrid1.Rows
    If MSHFlexGrid1.TextMatrix(k, 1) = "" Then Exit For
    If MSHFlexGrid1.TextMatrix(k, 0) = strChecked Then
       Retval = True
       Exit For
    End If
Next
ChkSelection = Retval
End Function

Private Function ChkBOD() As Boolean
Dim Retval As Boolean
Retval = True

tmp = "Select ML.LocationName"
tmp = tmp & " From Mst_CMP MC Inner Join Mst_Location ML On MC.LocationID = ML.LocationID"
tmp = tmp & " Inner Join Mst_State MS On MS.StateID = ML.StateID"
'tmp = tmp & " Where MC.CloseDate Is Null "

If Gen_DBType = "MDB" Then
   tmp = tmp & " Where MC.CurrentDate <= #" & Format(mToDate, Gen_DatFormat) & "# "
Else
   tmp = tmp & " Where MC.CurrentDate <= '" & Format(mToDate, Gen_DatFormat) & "' "
End If

tmp = tmp & " And MC.Closedate is null "

tmp = tmp & " And ML." & mGenInvoiceLocation

tmp = tmp & " Order by ML.LocationName,MC.CMPName"
Call TmpTable

If TmpRs.RecordCount > 0 Then
   Retval = False
End If

ChkBOD = Retval

End Function

Private Sub MSHFlexGrid1_Click()
'
'Head CMG


If LCase(Gen_UserRole) <> "power" Then
   If LCase(Gen_UserRole) = "admin" Or LCase(Gen_UserRole) = "superadmin" Or LCase(Gen_UserRole) = "headcmg" Then
      If mCountLocation = 5 Then
         If MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, MSHFlexGrid1.Col) = strUnChecked Then
            MsgBox "You can't select more than 5 location. Only 'Power' user can select all location!!!"
            Exit Sub
         End If
      End If
   Else
      If mCountLocation = 1 Then
         If MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, MSHFlexGrid1.Col) = strUnChecked Then
            MsgBox "You can't select more than 1 location. Only 'Power' user can select all location!!!"
            Exit Sub
         End If
      End If
   End If
End If

With MSHFlexGrid1
     If .Col = 0 Then
        If .TextMatrix(.row, .Col) = "" Then Exit Sub
        .Col = 0
        If .TextMatrix(.row, .Col) = strUnChecked Then
           .TextMatrix(.row, .Col) = strChecked
           mCountLocation = mCountLocation + 1
        Else
           .TextMatrix(.row, .Col) = strUnChecked
           mCountLocation = mCountLocation - 1
        End If
     End If
End With

End Sub

'Private Sub AddDummyLot()
'
'mFromDate = Format(CmbYear.Text & "-" & mMonth & "-" & "01", "yyyy-mm-dd")
'mYear = Val(CmbYear.Text)
'mTotalDays = Day(DateSerial(mYear, mMonth + 1, 0))
'
'mToDate = Format(CmbYear.Text & "-" & mMonth & "-" & mTotalDays, "yyyy-mm-dd")
'mNoofdays = ""
'
'
'tmp = "Select MF.LocationID,MF.VendorID,MFG.GodownID from Mst_FF_RateCard_Godown MFG"
'tmp = tmp & " inner join Mst_FF_RateCard MF on MF.FF_RateCardID=MFG.FF_RateCardID"
'tmp = tmp & " Where MFG.BillingUnit='F'"
'
'Call TmpTable
'
'For I = 1 To TmpRs.RecordCount
'
'    tmp = "Select MG.GodownID from Txn_InvMonGSLDetail TINV"
'    tmp = tmp & " Inner Join Mst_Godown MG On TINV.GodownID = MG.GodownID"
'    tmp = tmp & " Where TINV.YearMonth = '" & mFromDate & "' And MG.GodownID=" & TmpRs!GodownID
'
'    Call Tmp1Table
'
'    If TmpRs1.RecordCount = 0 Then
'
'       tmp = "Insert into Txn_InvMonGSLDetail (GSLID,SM_Prefix,CMPID,CommodityID,GodownID,StackNo,LotNo,FranchiseeRate,FranchBillingCycleID,"
'       tmp = tmp & " FranchBillingUnit,FranchNoofDays,FranchInvoiceAmount,FF_RateCardID,FranchProcessFlag,YearMonth,TFlag,ISINID,MonG_UID,FranchProcessdBy,FranchProcessedDate)"
'       tmp = tmp & " Select 0,ML.SM_Prefix,CMPID,0 CommodityID,MFG.GodownID,'Dummy-F' as StackNo,'Dummy-F' as LotNo ,0 as FranchiseeRate,"
'       tmp = tmp & " MFg.BillingCycleID,MFG.BillingUnit,1 as FranchNoofDays,0 as FranchInvoiceAmount,MF.FF_RateCardID,'Y','" & mFromDate & "','','" & "V" & TmpRs!VendorID & "',NewID(),'" & Gen_UserLoginID & "',getdate()"
'       tmp = tmp & " From Mst_FF_RateCard_Godown MFG"
'       tmp = tmp & " inner join Mst_FF_RateCard MF on MF.FF_RateCardID=MFG.FF_RateCardID"
'       tmp = tmp & " Inner join Mst_Location ML on ML.LocationID=MF.LocationID"
'       tmp = tmp & " Inner join Mst_Vendor MD on MD.VendorID=MF.VendorID"
'       tmp = tmp & " Where MFG.BillingUnit='F' And  MF.LocationID=" & TmpRs!LocationID & " And MF.VendorID=" & TmpRs!VendorID
'
'       Call Tmp2Table
'
'       tmp = "Select MNRCB.FixAmount,MNR.exchangetypeid,b.GodownID,MNRCB.StartDate,MNRCB.EndDate,ML.SM_Prefix"
'       tmp = tmp & " from  Mst_FF_RateCard_Godown_Fix  MNRCB"
'       tmp = tmp & " inner Join Mst_FF_RateCard_Godown b On b.FF_RateCard_GodownID = MNRCB.FF_RateCard_GodownID"
'       tmp = tmp & " inner Join Mst_FF_RateCard MNR On b.FF_RateCardID = MNR.FF_RateCardID"
'       tmp = tmp & " inner Join Mst_Location Ml On Ml.LocationID= MNR.LocationID"
'       tmp = tmp & " Where b.BillingUnit='F'"
'       tmp = tmp & " And MNRCB.StartDate<='" & mToDate & "'  And MNRCB.EndDate>='" & mFromDate & "' And b.godownid = '" & TmpRs!GodownID & "' "
'
'       Call Tmp2Table
'
'       For rn = 1 To TmpRs2.RecordCount
'
'           If TmpRs2!StartDate <= CDate(mFromDate) And TmpRs2!EndDate >= CDate(mToDate) Then
'              mMonFixRate = TmpRs2!FixAmount
'           ElseIf TmpRs2!StartDate <= CDate(mFromDate) And TmpRs2!EndDate < CDate(mToDate) Then
'              mTotDays = DateDiff("d", CDate(mFromDate), TmpRs2!EndDate) + 1
'              mMonFixRate = (TmpRs2!FixAmount / 30) * mTotDays
'           ElseIf TmpRs2!StartDate > CDate(mFromDate) And TmpRs2!EndDate >= CDate(mToDate) Then
'              mTotDays = DateDiff("d", TmpRs2!StartDate, CDate(mToDate)) + 1
'              mMonFixRate = (TmpRs2!FixAmount / 30) * mTotDays
'           ElseIf TmpRs2!StartDate > CDate(mFromDate) And TmpRs2!EndDate < CDate(mToDate) Then
'              mTotDays = DateDiff("d", TmpRs2!StartDate, TmpRs2!EndDate) + 1
'              mMonFixRate = (TmpRs2!FixAmount / 30) * mTotDays
'           End If
'
'           tmp = "Update Txn_InvMonGSLDetail Set FranchiseeRate = isnull(FranchiseeRate,0) + " & mMonFixRate & " ,FranchInvoiceAmount=isnull(FranchInvoiceAmount,0) + " & mMonFixRate & ""
'           tmp = tmp & " Where Txn_InvMonGSLDetail.YearMonth = '" & mFromDate & "' And Txn_InvMonGSLDetail.SM_Prefix = '" & TmpRs2!SM_Prefix & "' And Txn_InvMonGSLDetail.FranchBillingUnit = 'F'"
'           tmp = tmp & " And GodownID=" & TmpRs2!GodownID & ""
'
'           Call Tmp3Table
'
'           TmpRs2.MoveNext
'       Next
'
'    End If
'
'    TmpRs.MoveNext
'Next
'
'End Sub


Private Sub AddDummyLot()

mFromDate = Format(CmbYear.Text & "-" & mMonth & "-" & "01", "yyyy-mm-dd")
mYear = Val(CmbYear.Text)
mTotalDays = Day(DateSerial(mYear, mMonth + 1, 0))

mToDate = Format(CmbYear.Text & "-" & mMonth & "-" & mTotalDays, "yyyy-mm-dd")
mNoofDays = ""

''''comment on 1st march 2013 by rakesh
'tmp = "Select MF.LocationID,MF.VendorID,MFG.GodownID from Mst_FF_RateCard_Godown MFG"
'tmp = tmp & " inner join Mst_FF_RateCard MF on MF.FF_RateCardID=MFG.FF_RateCardID"
'tmp = tmp & " Where MFG.BillingUnit='F'"

''''comment on 4 march 2013 by rakesh
'tmp = "Select MF.LocationID,MF.VendorID,MFG.GodownID from Mst_FF_RateCard_Godown MFG"
'tmp = tmp & " inner join Mst_FF_RateCard MF on MF.FF_RateCardID=MFG.FF_RateCardID"
'tmp = tmp & " inner join mst_location ml on ml.locationid = mf.locationid"
'tmp = tmp & " inner join mst_state ms on ms.stateid = ml.stateid"
'tmp = tmp & " Where MFG.BillingUnit='F' and ms.statename = '" & CmbStateName.Text & "'"

tmp = "Select MF.LocationID,MF.VendorID,MFG.GodownID from Mst_FF_RateCard_Godown MFG"
tmp = tmp & " inner join Mst_FF_RateCard MF on MF.FF_RateCardID=MFG.FF_RateCardID"
tmp = tmp & " inner join mst_location ml on ml.locationid = mf.locationid"
tmp = tmp & " Where MFG.BillingUnit='F' and ml.locationname = '" & CmbLocation.Text & "'"

'---------------

Call TmpTable

For I = 1 To TmpRs.RecordCount

    tmp = "Select MG.GodownID from Txn_InvMonGSLDetail TINV"
    tmp = tmp & " Inner Join Mst_Godown MG On TINV.GodownID = MG.GodownID"
    tmp = tmp & " Where TINV.YearMonth = '" & mFromDate & "' And MG.GodownID=" & TmpRs!GodownID
    
    Call Tmp1Table
    
    If TmpRs1.RecordCount = 0 Then
    
       tmp = "Insert into Txn_InvMonGSLDetail (GSLID,SM_Prefix,CMPID,CommodityID,GodownID,StackNo,LotNo,FranchiseeRate,FranchBillingCycleID,"
       tmp = tmp & " FranchBillingUnit,FranchNoofDays,FranchInvoiceAmount,FF_RateCardID,FranchProcessFlag,YearMonth,TFlag,ISINID,MonG_UID,FranchProcessdBy,FranchProcessedDate)"
       tmp = tmp & " Select 0,ML.SM_Prefix,CMPID,0 CommodityID,MFG.GodownID,'Dummy-F' as StackNo,'Dummy-F' as LotNo ,0 as FranchiseeRate,"
       tmp = tmp & " MFg.BillingCycleID,MFG.BillingUnit,1 as FranchNoofDays,0 as FranchInvoiceAmount,MF.FF_RateCardID,'Y','" & mFromDate & "','','" & "V" & TmpRs!VendorID & "',NewID(),'" & Gen_UserLoginID & "',getdate()"
       tmp = tmp & " From Mst_FF_RateCard_Godown MFG"
       tmp = tmp & " inner join Mst_FF_RateCard MF on MF.FF_RateCardID=MFG.FF_RateCardID"
       tmp = tmp & " Inner join Mst_Location ML on ML.LocationID=MF.LocationID"
       tmp = tmp & " Inner join Mst_Vendor MD on MD.VendorID=MF.VendorID"
       tmp = tmp & " Where MFG.BillingUnit='F' And  MF.LocationID=" & TmpRs!LocationID & " And MF.VendorID=" & TmpRs!VendorID & " And MFG.GodownID = " & TmpRs!GodownID & " "
 
       Call Tmp2Table
       
       tmp = "Select MNRCB.FixAmount,MNR.exchangetypeid,b.GodownID,MNRCB.StartDate,MNRCB.EndDate,ML.SM_Prefix"
       tmp = tmp & " from  Mst_FF_RateCard_Godown_Fix  MNRCB"
       tmp = tmp & " inner Join Mst_FF_RateCard_Godown b On b.FF_RateCard_GodownID = MNRCB.FF_RateCard_GodownID"
       tmp = tmp & " inner Join Mst_FF_RateCard MNR On b.FF_RateCardID = MNR.FF_RateCardID"
       tmp = tmp & " inner Join Mst_Location Ml On Ml.LocationID= MNR.LocationID"
       tmp = tmp & " Where b.BillingUnit='F'"
       tmp = tmp & " And MNRCB.StartDate<='" & mToDate & "'  And MNRCB.EndDate>='" & mFromDate & "' And b.GodownID = " & TmpRs!GodownID & " "
        
       Call Tmp2Table
            
       For rn = 1 To TmpRs2.RecordCount
        
           If TmpRs2!StartDate <= CDate(mFromDate) And TmpRs2!EndDate >= CDate(mToDate) Then
              mMonFixRate = TmpRs2!FixAmount
           ElseIf TmpRs2!StartDate <= CDate(mFromDate) And TmpRs2!EndDate < CDate(mToDate) Then
              mTotDays = DateDiff("d", CDate(mFromDate), TmpRs2!EndDate) + 1
              mMonFixRate = (TmpRs2!FixAmount / 30) * mTotDays
           ElseIf TmpRs2!StartDate > CDate(mFromDate) And TmpRs2!EndDate >= CDate(mToDate) Then
              mTotDays = DateDiff("d", TmpRs2!StartDate, CDate(mToDate)) + 1
              mMonFixRate = (TmpRs2!FixAmount / 30) * mTotDays
           ElseIf TmpRs2!StartDate > CDate(mFromDate) And TmpRs2!EndDate < CDate(mToDate) Then
              mTotDays = DateDiff("d", TmpRs2!StartDate, TmpRs2!EndDate) + 1
              mMonFixRate = (TmpRs2!FixAmount / 30) * mTotDays
           End If
            
           tmp = "Update Txn_InvMonGSLDetail Set FranchiseeRate = isnull(FranchiseeRate,0) + " & mMonFixRate & " ,FranchInvoiceAmount=isnull(FranchInvoiceAmount,0) + " & mMonFixRate & ""
           tmp = tmp & " Where Txn_InvMonGSLDetail.YearMonth = '" & mFromDate & "' And Txn_InvMonGSLDetail.SM_Prefix = '" & TmpRs2!SM_Prefix & "' And Txn_InvMonGSLDetail.FranchBillingUnit = 'F'"
           tmp = tmp & " And GodownID=" & TmpRs2!GodownID & ""
              
           Call Tmp3Table
               
           TmpRs2.MoveNext
       Next
       
    End If
    
    TmpRs.MoveNext
Next

End Sub


Private Sub CmbLocationID_GotFocus()
tmp = CmbLocationID.Text

If Gen_WcLocation = "" Then
   Call TableMst_Location("")
Else
   Call TableMst_Location("where " & Gen_WcLocation)
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
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocationID.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection!!! "
   CmbLocationID.SetFocus
   Exit Sub
End If
mLocationID = RsMst_Location!LocationID
mSM_Prefix = RsMst_Location!SM_Prefix

End Sub

Private Sub GetNNNRIncome(mFSM_Prefix As Variant)

tmp = "Update Txn_InvMonGSLDetail set InvoiceAmtForFranch = 0,InsuranceCostDeduction=0 Where yearmonth = '" & mFromDate & "' And Isnull(InvoiceAmtForFranch,0) = 0 And SM_Prefix = '" & mFSM_Prefix & "' "
Call Tmp1Table

tmp = "Update Txn_InvMonGSLDetail set InvoiceAmtForFranch = InvoiceAmount where yearmonth = '" & mFromDate & "' And InvoiceAmount is not null and ExchangeTypeID = 1 And (ContractID = 0 Or ContractID Is Null) And SM_Prefix = '" & mFSM_Prefix & "' "
Call Tmp1Table
 

''========================Discount And Premium

tmp = " Select TIG.PINo, TIG.Mst_GSL_ID,MonG_UID, TIG.InvoiceAmount 'IncomeAmt'"
tmp = tmp & " from Txn_InvMonGSLDetail TIG"
tmp = tmp & " Inner Join Mst_Godown MG on MG.GodownID = TIG.GodownID"
tmp = tmp & " Inner Join Txn_ProInvoice TPI on TPI.InvoiceNo = TIG.PINo And isnull(TPI.ReverseFlag, '') <> 'Y'"
tmp = tmp & " Where TIG.ExchangeTypeID = 1 And (TIG.ContractID = 0 Or TIG.ContractID Is Null)"
tmp = tmp & " And PINo is not Null And TIG.BalanceWeight >= 0"
tmp = tmp & " And TIG.YearMonth Between '" & mFromDate & "' and '" & mToDate & "'"
'tmp = tmp & " And (MG.CloseDate is null Or MG.CloseDate >= '" & mFromDate & "') And MG.StartDate <= '" & mToDate & "'"
tmp = tmp & " And TIG.SM_Prefix = '" & mFSM_Prefix & "'"


Call Tmp1Table

tmp = " Select Amount, InvoiceNo, Discount, Premium from txn_proinvoice Where (Discount > 0 Or Premium > 0) "
tmp = tmp & " and InvoiceNo in ("
tmp = tmp & " Select TIG.PINo from Txn_InvMonGSLDetail TIG"
tmp = tmp & " Inner Join Mst_Godown MG on MG.GodownID = TIG.GodownID"
tmp = tmp & " Inner Join Txn_ProInvoice TPI on TPI.InvoiceNo = TIG.PINo And isnull(TPI.ReverseFlag, '') <> 'Y'"
tmp = tmp & " Where TIG.ExchangeTypeID = 1 And (TIG.ContractID = 0 Or TIG.ContractID Is Null)"
tmp = tmp & " And PINo is not Null And TIG.BalanceWeight >= 0"
tmp = tmp & " And TIG.YearMonth Between '" & mFromDate & "' and '" & mToDate & "'"
'tmp = tmp & " And (MG.CloseDate is null Or MG.CloseDate >= '" & mFromDate & "') And MG.StartDate <= '" & mToDate & "'"
tmp = tmp & " And TIG.SM_Prefix = '" & mFSM_Prefix & "' "
tmp = tmp & " )"

Call Tmp3Table

tmp = "Select Mst_GSL_ID,MonG_UID,InvoiceAmtForFranch From Txn_InvMonGSLDetail TIG "
tmp = tmp & " Where TIG.ExchangeTypeID = 1 And (TIG.ContractID = 0 Or TIG.ContractID Is Null)"
tmp = tmp & " And PINo is not Null And TIG.BalanceWeight >= 0"
tmp = tmp & " And TIG.YearMonth Between '" & mFromDate & "' and '" & mToDate & "'"
tmp = tmp & " And TIG.SM_Prefix = '" & mFSM_Prefix & "' "

Call Tmp4Table

If TmpRs1.RecordCount > 0 And TmpRs3.RecordCount > 0 Then
   For pqr = 1 To TmpRs1.RecordCount
      TmpRs3.MoveFirst
      TmpRs3.Find " InvoiceNo = " & TmpRs1!PINo
      If Not TmpRs3.EOF Then
         mDiscount = (TmpRs3!Discount * IIf(IsNull(TmpRs1!IncomeAmt), 0, TmpRs1!IncomeAmt)) / TmpRs3!Amount
         mPremium = (TmpRs3!Premium * IIf(IsNull(TmpRs1!IncomeAmt), 0, TmpRs1!IncomeAmt)) / TmpRs3!Amount
         TmpRs4.MoveFirst
         TmpRs4.Find "MonG_UID = '" & TmpRs1!MonG_UID & " '"
         If TmpRs4.EOF = False Then
            TmpRs4!InvoiceAmtForFranch = Val(TmpRs4!InvoiceAmtForFranch) + Format(Val(mPremium) - Val(mDiscount), "##########.00")
            TmpRs4.Update
         End If
      End If
      TmpRs1.MoveNext
   Next
End If

tmp = " Select TBGSL.Mst_GSL_ID, TBGSL.Amount 'Amount' from Txn_CDTFBOE_GSL TBGSL"
tmp = tmp & " Inner Join Txn_CDTFBOE TB On TBGSL.BOEID = TB.BOEID"
tmp = tmp & " Where TB.BillingMonth Between '" & mFromDate & "' And '" & mToDate & "'"
'tmp = tmp & " And (MG.CloseDate is null Or MG.CloseDate >= '" & mFromDate & "' ) And MG.StartDate <= '" & mToDate & "'"
tmp = tmp & " And TB.SM_Prefix = '" & mFSM_Prefix & "'"
tmp = tmp & " And TB.TxnType = 'Y' And isnuLL(TB.ReverseFlag,'') <> 'Y' "

Call Tmp1Table


tmp = "Select Mst_GSL_ID,MonG_UID,InvoiceAmtForFranch From Txn_InvMonGSLDetail TIG "
tmp = tmp & " Where TIG.YearMonth Between '" & mFromDate & "' and '" & mToDate & "'"
tmp = tmp & " And TIG.SM_Prefix = '" & mFSM_Prefix & "' And TIG.Mst_GSL_ID in "
tmp = tmp & " (Select Distinct TBGSL.Mst_GSL_ID from Txn_CDTFBOE TB"
tmp = tmp & " Inner Join Txn_CDTFBOE_GSL TBGSL on TBGSL.BOEID = TB.BOEID"
tmp = tmp & " Where TB.BillingMonth Between '" & mFromDate & "' And '" & mToDate & "' "
tmp = tmp & " And TB.TxnType = 'Y' And isnuLL(TB.ReverseFlag,'') <> 'Y') "
'tmp = tmp & " And (MG.CloseDate is null Or MG.CloseDate >= '" & mFromDate & "' ) And MG.StartDate <= '" & mToDate & "'"
tmp = tmp & " And TIG.SM_Prefix = '" & mFSM_Prefix & "'"



Call Tmp4Table

If TmpRs1.RecordCount > 0 Then
   For pqr = 1 To TmpRs4.RecordCount
      TmpRs1.MoveFirst
      TmpRs1.Find "Mst_GSL_ID = " & Val(TmpRs4!Mst_GSL_ID)
      If Not TmpRs1.EOF Then
         TmpRs4!InvoiceAmtForFranch = Val(TmpRs4!InvoiceAmtForFranch) + IIf(IsNull(TmpRs1!Amount), 0, TmpRs1!Amount)
         TmpRs4.Update
      End If
      TmpRs4.MoveNext
   Next
End If


''========================Discount And Premium

tmp = " Select TB.BOEID,MGSL.Mst_GSL_ID,MGSL.MonG_UID , TBGSL.Amount IncomeAmt from Txn_CDTFBOE TB"
tmp = tmp & " Inner Join Txn_CDTFBOE_GSL TBGSL on TBGSL.BOEID = TB.BOEID"
tmp = tmp & " Inner Join Txn_InvMonGSLDetail MGSL on MGSL.SM_Prefix = TBGSL.SM_Prefix And MGSL.Mst_GSL_ID = TBGSL.Mst_GSL_ID"
tmp = tmp & " Where TB.BillingMonth Between '" & mFromDate & "' And '" & mToDate & "' And isnull(TB.ReverseFlag, '') <> 'Y'"
tmp = tmp & " And MGSL.SM_Prefix = '" & mFSM_Prefix & "' And MGSL.YearMonth Between '" & mFromDate & "' and '" & mToDate & "'"
tmp = tmp & " And TB.TxnType = 'Y'"
Call Tmp1Table

tmp = " Select Amount, BOEID, Discount, Premium from Txn_CDTFBOE Where (Discount > 0 Or Premium > 0) "
tmp = tmp & " and BOEID in ("
tmp = tmp & " Select TB.BOEID from Txn_CDTFBOE TB"
tmp = tmp & " Inner Join Txn_CDTFBOE_GSL TBGSL on TBGSL.BOEID = TB.BOEID"
tmp = tmp & " Inner Join Txn_InvMonGSLDetail MGSL on MGSL.SM_Prefix = TBGSL.SM_Prefix And MGSL.Mst_GSL_ID = TBGSL.Mst_GSL_ID"
tmp = tmp & " Where TB.BillingMonth Between '" & mFromDate & "' And '" & mToDate & "' And isnull(TB.ReverseFlag, '') <> 'Y'"
tmp = tmp & " And MGSL.SM_Prefix = '" & mFSM_Prefix & "'"
tmp = tmp & " And TB.TxnType = 'Y'"
tmp = tmp & " And MGSL.SM_Prefix = '" & mFSM_Prefix & "' And MGSL.YearMonth Between '" & mFromDate & "' and '" & mToDate & "')"

Call Tmp3Table

tmp = "Select Mst_GSL_ID,MonG_UID,InvoiceAmtForFranch From Txn_InvMonGSLDetail TIG "
tmp = tmp & " Where TIG.ExchangeTypeID = 1 And (TIG.ContractID = 0 Or TIG.ContractID Is Null)"
tmp = tmp & " And PINo is not Null And TIG.BalanceWeight >= 0"
tmp = tmp & " And TIG.YearMonth Between '" & mFromDate & "' and '" & mToDate & "'"
tmp = tmp & " And TIG.SM_Prefix = '" & mFSM_Prefix & "' "

Call Tmp4Table

If TmpRs1.RecordCount > 0 And TmpRs3.RecordCount > 0 Then
   For pqr = 1 To TmpRs1.RecordCount
       TmpRs3.MoveFirst
       TmpRs3.Find " BOEID = " & TmpRs1!BOEID
       If Not TmpRs3.EOF Then
          mDiscount = (TmpRs3!Discount * IIf(IsNull(TmpRs1!IncomeAmt), 0, TmpRs1!IncomeAmt)) / TmpRs3!Amount
          mPremium = (TmpRs3!Premium * IIf(IsNull(TmpRs1!IncomeAmt), 0, TmpRs1!IncomeAmt)) / TmpRs3!Amount
          TmpRs4.MoveFirst
          TmpRs4.Find "MonG_UID = '" & TmpRs1!MonG_UID & " '"
          If TmpRs4.EOF = False Then
             TmpRs4!InvoiceAmtForFranch = Val(TmpRs4!InvoiceAmtForFranch) + Format(Val(mPremium) - Val(mDiscount), "##########.00")
             TmpRs4.Update
          End If
       End If
       TmpRs1.MoveNext
    Next
End If

tmp = " Select TISGD.GodownID, sum(TISGD.InvoiceAmt) IncomeAmt"
tmp = tmp & " from Txn_TempInvoice TI"
tmp = tmp & " Inner Join Txn_TempInvGodownDetails TISGD on TI.PINo = TISGD.PINo"
tmp = tmp & " Where TI.Flag in ('A', 'P','F') And TI.IETypeID = 7 And isnuLL(TI.ReverseFlag,'') <> 'Y'"
tmp = tmp & " And TI.InvoiceDate Between   '" & mFromDate & "' And '" & mToDate & "' and Isnull(TI.ToBeExcludeForFranchaiseeCal,0) = 0"
tmp = tmp & " Group by TISGD.GodownID"

Call Tmp1Table


If TmpRs1.RecordCount > 0 Then
   For pqr = 1 To TmpRs1.RecordCount
       tmp = "Select GodownID,Mst_GSL_ID,MonG_UID,InvoiceAmtForFranch From Txn_InvMonGSLDetail TIG "
       tmp = tmp & " Where TIG.ExchangeTypeID = 1 And (TIG.ContractID = 0 Or TIG.ContractID Is Null)"
       tmp = tmp & " And TIG.BalanceWeight >= 0"
       tmp = tmp & " And TIG.YearMonth Between '" & mFromDate & "' and '" & mToDate & "'"
       tmp = tmp & " And TIG.SM_Prefix = '" & mFSM_Prefix & "' And TIG.GodownID = " & TmpRs1!GodownID & ""
        
       Call Tmp4Table
       
       For I = 1 To TmpRs4.RecordCount
           TmpRs4!InvoiceAmtForFranch = TmpRs4!InvoiceAmtForFranch + Val((TmpRs1!IncomeAmt / TmpRs4.RecordCount))  'IIf(IsNull(TmpRs1!IncomeAmt), 0, Format(TmpRs1!IncomeAmt * Val(MSHFlexGrid1.TextMatrix(i, E.GodownCapacity)) / TmpRs2!Capacity, "#############.00"))
           TmpRs4.Update
           TmpRs4.MoveNext
       Next
       TmpRs1.MoveNext
   Next
End If


End Sub

Private Sub GetNCDEXIncome(mFSM_Prefix As Variant)

tmp = " Select TISGD.GodownID, sum(TISGD.InvoiceAmt) IncomeAmt"
tmp = tmp & " from Txn_TempInvoice TI"
tmp = tmp & " Inner Join Txn_TempInvGodownDetails TISGD on TI.PINo = TISGD.PINo"
tmp = tmp & " Where TI.Flag in ('A', 'P','F') And TI.IETypeID in (1,2,3,13) And isnuLL(TI.ReverseFlag,'') <> 'Y' and Isnull(TI.ToBeExcludeForFranchaiseeCal,0) = 0"
tmp = tmp & " And TI.InvoiceDate Between   '" & mFromDate & "' And '" & mToDate & "'"
tmp = tmp & " Group by TISGD.GodownID"

Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   For pqr = 1 To TmpRs1.RecordCount
       tmp = "Select GodownID,Mst_GSL_ID,MonG_UID,InvoiceAmtForFranch From Txn_InvMonGSLDetail TIG "
       tmp = tmp & " Where TIG.ExchangeTypeID = 2 "
       tmp = tmp & " And TIG.BalanceWeight >= 0"
       tmp = tmp & " And TIG.YearMonth Between '" & mFromDate & "' and '" & mToDate & "'"
       tmp = tmp & " And TIG.SM_Prefix = '" & mFSM_Prefix & "' And TIG.GodownID = " & TmpRs1!GodownID & ""
        
       Call Tmp4Table
       
       For Ilk = 1 To TmpRs4.RecordCount
           TmpRs4!InvoiceAmtForFranch = TmpRs4!InvoiceAmtForFranch + Val((TmpRs1!IncomeAmt / TmpRs4.RecordCount))  'IIf(IsNull(TmpRs1!IncomeAmt), 0, Format(TmpRs1!IncomeAmt * Val(MSHFlexGrid1.TextMatrix(i, E.GodownCapacity)) / TmpRs2!Capacity, "#############.00"))
           TmpRs4.Update
           TmpRs4.MoveNext
       Next
       TmpRs1.MoveNext
   Next
End If


tmp = " Select TBGSL.Mst_GSL_ID, Sum(TBGSL.Amount) 'Amount' from Txn_CDTFBOE_GSL TBGSL"
tmp = tmp & " Inner Join Txn_CDTFBOE TB On TBGSL.BOEID = TB.BOEID"
tmp = tmp & " Where TB.BillingMonth Between '" & mFromDate & "' And '" & mToDate & "'"
tmp = tmp & " And TB.SM_Prefix = '" & mFSM_Prefix & "'"
tmp = tmp & " And TB.TxnType in ('D') And isnuLL(TB.ReverseFlag,'') <> 'Y'  And ((TB.Amount+TB.Premium)-TB.Discount) > 0 Group by TBGSL.Mst_GSL_ID"

Call Tmp1Table


tmp = "Select Mst_GSL_ID,MonG_UID,InvoiceAmtForFranch From Txn_InvMonGSLDetail TIG "
tmp = tmp & " Where TIG.YearMonth Between '" & mFromDate & "' and '" & mToDate & "'"
tmp = tmp & " And TIG.SM_Prefix = '" & mFSM_Prefix & "' And TIG.Mst_GSL_ID in "
tmp = tmp & " (Select Distinct TBGSL.Mst_GSL_ID from Txn_CDTFBOE TB"
tmp = tmp & " Inner Join Txn_CDTFBOE_GSL TBGSL on TBGSL.BOEID = TB.BOEID"
tmp = tmp & " Where TB.BillingMonth Between '" & mFromDate & "' And '" & mToDate & "' "
tmp = tmp & " And TB.TxnType in ('D') And isnuLL(TB.ReverseFlag,'') <> 'Y' And ((TB.Amount+TB.Premium)-TB.Discount) > 0) "
tmp = tmp & " And TIG.SM_Prefix = '" & mFSM_Prefix & "' And TIG.TFLag in ('D','O')"

Call Tmp4Table

'If TmpRs1.RecordCount > 0 Then
'   For pqr = 1 To TmpRs4.RecordCount
'      TmpRs1.MoveFirst
'      TmpRs1.Filter = "Mst_GSL_ID = " & Val(TmpRs4!Mst_GSL_ID)
''      If TmpRs1.RecordCount > 0 Then
'''         TmpRs4!InvoiceAmtForFranch = Val(TmpRs4!InvoiceAmtForFranch) + (IIf(IsNull(TmpRs1!Amount), 0, TmpRs1!Amount) / TmpRs1.RecordCount)
'''         TmpRs4.Update
''      End If
'
'      For Ilk = 1 To TmpRs1.RecordCount
'          TmpRs1!InvoiceAmtForFranch = Val(TmpRs4!InvoiceAmtForFranch) + (IIf(IsNull(TmpRs1!Amount), 0, TmpRs1!Amount) / TmpRs1.RecordCount)  'IIf(IsNull(TmpRs1!IncomeAmt), 0, Format(TmpRs1!IncomeAmt * Val(MSHFlexGrid1.TextMatrix(i, E.GodownCapacity)) / TmpRs2!Capacity, "#############.00"))
'          TmpRs1.Update
'          TmpRs1.MoveNext
'      Next
'
'      TmpRs1.Filter = ""
'      TmpRs4.MoveNext
'   Next
'End If


If TmpRs1.RecordCount > 0 Then
   For pqr = 1 To TmpRs1.RecordCount
    If TmpRs4.RecordCount > 0 Then
          TmpRs4.MoveFirst
          TmpRs4.Filter = "Mst_GSL_ID = " & Val(TmpRs1!Mst_GSL_ID)
    '      If TmpRs1.RecordCount > 0 Then
    ''         TmpRs4!InvoiceAmtForFranch = Val(TmpRs4!InvoiceAmtForFranch) + (IIf(IsNull(TmpRs1!Amount), 0, TmpRs1!Amount) / TmpRs1.RecordCount)
    ''         TmpRs4.Update
    '      End If
          
          For Ilk = 1 To TmpRs4.RecordCount
              TmpRs4!InvoiceAmtForFranch = Val(TmpRs4!InvoiceAmtForFranch) + (IIf(IsNull(TmpRs1!Amount), 0, TmpRs1!Amount) / TmpRs4.RecordCount)  'IIf(IsNull(TmpRs1!IncomeAmt), 0, Format(TmpRs1!IncomeAmt * Val(MSHFlexGrid1.TextMatrix(i, E.GodownCapacity)) / TmpRs2!Capacity, "#############.00"))
              TmpRs4.Update
              TmpRs4.MoveNext
          Next
          
          TmpRs4.Filter = ""
      End If
    TmpRs1.MoveNext
   Next
End If


''========================Discount And Premium

tmp = " Select TB.BOEID,MGSL.Mst_GSL_ID,MGSL.MonG_UID , TBGSL.Amount IncomeAmt from Txn_CDTFBOE TB"
tmp = tmp & " Inner Join Txn_CDTFBOE_GSL TBGSL on TBGSL.BOEID = TB.BOEID"
tmp = tmp & " Inner Join Txn_InvMonGSLDetail MGSL on MGSL.SM_Prefix = TBGSL.SM_Prefix And MGSL.Mst_GSL_ID = TBGSL.Mst_GSL_ID"
tmp = tmp & " Where TB.BillingMonth Between '" & mFromDate & "' And '" & mToDate & "' And isnull(TB.ReverseFlag, '') <> 'Y'"
tmp = tmp & " And MGSL.SM_Prefix = '" & mFSM_Prefix & "' And MGSL.YearMonth Between '" & mFromDate & "' and '" & mToDate & "'"
tmp = tmp & " And TB.TxnType in ('D') And ((TB.Amount+TB.Premium)-TB.Discount) > 0 And MGSL.TFLag = 'D'"
Call Tmp1Table

tmp = " Select Amount, BOEID, Discount, Premium from Txn_CDTFBOE Where (Discount > 0 Or Premium > 0) "
tmp = tmp & " and BOEID in ("
tmp = tmp & " Select TB.BOEID from Txn_CDTFBOE TB"
tmp = tmp & " Inner Join Txn_CDTFBOE_GSL TBGSL on TBGSL.BOEID = TB.BOEID"
tmp = tmp & " Inner Join Txn_InvMonGSLDetail MGSL on MGSL.SM_Prefix = TBGSL.SM_Prefix And MGSL.Mst_GSL_ID = TBGSL.Mst_GSL_ID"
tmp = tmp & " Where TB.BillingMonth Between '" & mFromDate & "' And '" & mToDate & "' And isnull(TB.ReverseFlag, '') <> 'Y'"
tmp = tmp & " And MGSL.SM_Prefix = '" & mFSM_Prefix & "'"
tmp = tmp & " And TB.TxnType in ('D') And ((TB.Amount+TB.Premium)-TB.Discount) > 0 "
tmp = tmp & " And MGSL.SM_Prefix = '" & mFSM_Prefix & "' And MGSL.YearMonth Between '" & mFromDate & "' and '" & mToDate & "' And MGSL.TFLag = 'D')"

Call Tmp3Table

tmp = "Select Mst_GSL_ID,MonG_UID,InvoiceAmtForFranch From Txn_InvMonGSLDetail TIG "
tmp = tmp & " Where TIG.ExchangeTypeID = 2 "
tmp = tmp & " And TIG.BalanceWeight >= 0"
tmp = tmp & " And TIG.YearMonth Between '" & mFromDate & "' and '" & mToDate & "'"
tmp = tmp & " And TIG.SM_Prefix = '" & mFSM_Prefix & "' And TIG.TFLag = 'D' "

Call Tmp4Table

If TmpRs1.RecordCount > 0 And TmpRs3.RecordCount > 0 Then
   For pqr = 1 To TmpRs1.RecordCount
       TmpRs3.MoveFirst
       TmpRs3.Find " BOEID = " & TmpRs1!BOEID
       If Not TmpRs3.EOF Then
          mDiscount = (TmpRs3!Discount * IIf(IsNull(TmpRs1!IncomeAmt), 0, TmpRs1!IncomeAmt)) / TmpRs3!Amount
          mPremium = (TmpRs3!Premium * IIf(IsNull(TmpRs1!IncomeAmt), 0, TmpRs1!IncomeAmt)) / TmpRs3!Amount
          TmpRs4.MoveFirst
          TmpRs4.Find "MonG_UID = '" & TmpRs1!MonG_UID & " '"
          If TmpRs4.EOF = False Then
             TmpRs4!InvoiceAmtForFranch = Val(TmpRs4!InvoiceAmtForFranch) + Format(Val(mPremium) - Val(mDiscount), "##########.00")
             TmpRs4.Update
          End If
       End If
       TmpRs1.MoveNext
    Next
End If


'---------------------

tmp = " Select TBGSL.Mst_GSL_ID, Sum(TBGSL.Amount) 'Amount' from Txn_CDTFBOE_GSL TBGSL"
tmp = tmp & " Inner Join Txn_CDTFBOE TB On TBGSL.BOEID = TB.BOEID"
tmp = tmp & " Where TB.BillingMonth Between '" & mFromDate & "' And '" & mToDate & "'"
tmp = tmp & " And TB.SM_Prefix = '" & mFSM_Prefix & "'"
tmp = tmp & " And TB.TxnType in ('W','R','P') And isnuLL(TB.ReverseFlag,'') <> 'Y'  And ((TB.Amount+TB.Premium)-TB.Discount) > 0 Group by TBGSL.Mst_GSL_ID "

Call Tmp1Table


tmp = "Select Mst_GSL_ID,MonG_UID,InvoiceAmtForFranch From Txn_InvMonGSLDetail TIG "
tmp = tmp & " Where TIG.YearMonth Between '" & mFromDate & "' and '" & mToDate & "'"
tmp = tmp & " And TIG.SM_Prefix = '" & mFSM_Prefix & "' And TIG.Mst_GSL_ID in "
tmp = tmp & " (Select Distinct TBGSL.Mst_GSL_ID from Txn_CDTFBOE TB"
tmp = tmp & " Inner Join Txn_CDTFBOE_GSL TBGSL on TBGSL.BOEID = TB.BOEID"
tmp = tmp & " Where TB.BillingMonth Between '" & mFromDate & "' And '" & mToDate & "' "
tmp = tmp & " And TB.TxnType in ('W','R','P') And isnuLL(TB.ReverseFlag,'') <> 'Y' And ((TB.Amount+TB.Premium)-TB.Discount) > 0) "
tmp = tmp & " And TIG.SM_Prefix = '" & mFSM_Prefix & "' And TIG.TFLag = 'W'"

Call Tmp4Table

'If TmpRs1.RecordCount > 0 Then
'   For pqr = 1 To TmpRs4.RecordCount
'      TmpRs1.MoveFirst
'      TmpRs1.Filter = "Mst_GSL_ID = " & Val(TmpRs4!Mst_GSL_ID)
''      If TmpRs1.RecordCount > 0 Then
''         TmpRs4!InvoiceAmtForFranch = Val(TmpRs4!InvoiceAmtForFranch) + (IIf(IsNull(TmpRs1!Amount), 0, TmpRs1!Amount) / TmpRs1.RecordCount)
''         TmpRs4.Update
''      End If
'
'      For Ilk = 1 To TmpRs1.RecordCount
'          TmpRs1!InvoiceAmtForFranch = Val(TmpRs4!InvoiceAmtForFranch) + (IIf(IsNull(TmpRs1!Amount), 0, TmpRs1!Amount) / TmpRs1.RecordCount)  'IIf(IsNull(TmpRs1!IncomeAmt), 0, Format(TmpRs1!IncomeAmt * Val(MSHFlexGrid1.TextMatrix(i, E.GodownCapacity)) / TmpRs2!Capacity, "#############.00"))
'          TmpRs1.Update
'          TmpRs1.MoveNext
'      Next
'
'      TmpRs1.Filter = ""
'      TmpRs4.MoveNext
'   Next
'End If


If TmpRs1.RecordCount > 0 Then
   For pqr = 1 To TmpRs1.RecordCount
      TmpRs4.MoveFirst
      TmpRs4.Filter = "Mst_GSL_ID = " & Val(TmpRs1!Mst_GSL_ID)
'      If TmpRs1.RecordCount > 0 Then
''         TmpRs4!InvoiceAmtForFranch = Val(TmpRs4!InvoiceAmtForFranch) + (IIf(IsNull(TmpRs1!Amount), 0, TmpRs1!Amount) / TmpRs1.RecordCount)
''         TmpRs4.Update
'      End If
      
      For Ilk = 1 To TmpRs4.RecordCount
          TmpRs4!InvoiceAmtForFranch = Val(TmpRs4!InvoiceAmtForFranch) + (IIf(IsNull(TmpRs1!Amount), 0, TmpRs1!Amount) / TmpRs4.RecordCount)  'IIf(IsNull(TmpRs1!IncomeAmt), 0, Format(TmpRs1!IncomeAmt * Val(MSHFlexGrid1.TextMatrix(i, E.GodownCapacity)) / TmpRs2!Capacity, "#############.00"))
          TmpRs4.Update
          TmpRs4.MoveNext
      Next
      
      TmpRs4.Filter = ""
      TmpRs1.MoveNext
   Next
End If

''========================Discount And Premium

tmp = " Select TB.BOEID,MGSL.Mst_GSL_ID,MGSL.MonG_UID , TBGSL.Amount IncomeAmt from Txn_CDTFBOE TB"
tmp = tmp & " Inner Join Txn_CDTFBOE_GSL TBGSL on TBGSL.BOEID = TB.BOEID"
tmp = tmp & " Inner Join Txn_InvMonGSLDetail MGSL on MGSL.SM_Prefix = TBGSL.SM_Prefix And MGSL.Mst_GSL_ID = TBGSL.Mst_GSL_ID"
tmp = tmp & " Where TB.BillingMonth Between '" & mFromDate & "' And '" & mToDate & "' And isnull(TB.ReverseFlag, '') <> 'Y'"
tmp = tmp & " And MGSL.SM_Prefix = '" & mFSM_Prefix & "' And MGSL.YearMonth Between '" & mFromDate & "' and '" & mToDate & "'"
tmp = tmp & " And TB.TxnType in ('W','R','P') And ((TB.Amount+TB.Premium)-TB.Discount) > 0 And MGSL.TFLag = 'W'"
Call Tmp1Table

tmp = " Select Amount, BOEID, Discount, Premium from Txn_CDTFBOE Where (Discount > 0 Or Premium > 0) "
tmp = tmp & " and BOEID in ("
tmp = tmp & " Select TB.BOEID from Txn_CDTFBOE TB"
tmp = tmp & " Inner Join Txn_CDTFBOE_GSL TBGSL on TBGSL.BOEID = TB.BOEID"
tmp = tmp & " Inner Join Txn_InvMonGSLDetail MGSL on MGSL.SM_Prefix = TBGSL.SM_Prefix And MGSL.Mst_GSL_ID = TBGSL.Mst_GSL_ID"
tmp = tmp & " Where TB.BillingMonth Between '" & mFromDate & "' And '" & mToDate & "' And isnull(TB.ReverseFlag, '') <> 'Y'"
tmp = tmp & " And MGSL.SM_Prefix = '" & mFSM_Prefix & "'"
tmp = tmp & " And TB.TxnType in ('W','R','P') And ((TB.Amount+TB.Premium)-TB.Discount) > 0 "
tmp = tmp & " And MGSL.SM_Prefix = '" & mFSM_Prefix & "' And MGSL.YearMonth Between '" & mFromDate & "' and '" & mToDate & "' And MGSL.TFLag = 'W')"

Call Tmp3Table

tmp = "Select Mst_GSL_ID,MonG_UID,InvoiceAmtForFranch From Txn_InvMonGSLDetail TIG "
tmp = tmp & " Where TIG.ExchangeTypeID = 2 "
tmp = tmp & " And TIG.BalanceWeight >= 0"
tmp = tmp & " And TIG.YearMonth Between '" & mFromDate & "' and '" & mToDate & "'"
tmp = tmp & " And TIG.SM_Prefix = '" & mFSM_Prefix & "' And TIG.TFLag = 'W'"

Call Tmp4Table

If TmpRs1.RecordCount > 0 And TmpRs3.RecordCount > 0 Then
   For pqr = 1 To TmpRs1.RecordCount
       TmpRs3.MoveFirst
       TmpRs3.Find " BOEID = " & TmpRs1!BOEID
       If Not TmpRs3.EOF Then
          mDiscount = (TmpRs3!Discount * IIf(IsNull(TmpRs1!IncomeAmt), 0, TmpRs1!IncomeAmt)) / TmpRs3!Amount
          mPremium = (TmpRs3!Premium * IIf(IsNull(TmpRs1!IncomeAmt), 0, TmpRs1!IncomeAmt)) / TmpRs3!Amount
          TmpRs4.MoveFirst
          TmpRs4.Find "MonG_UID = '" & TmpRs1!MonG_UID & " '"
          If TmpRs4.EOF = False Then
             TmpRs4!InvoiceAmtForFranch = Val(TmpRs4!InvoiceAmtForFranch) + Format(Val(mPremium) - Val(mDiscount), "##########.00")
             TmpRs4.Update
          End If
       End If
       TmpRs1.MoveNext
    Next
End If



tmp = "Select Distinct Mst_GSL_ID From Txn_InvMonGSLDetail Where ExchangeTypeID = 2 and YearMonth = '" & mFromDate & "' and ISINID is null and SM_PRefix = '" & mFSM_Prefix & "' "
Call Tmp2Table

For pqr = 1 To TmpRs2.RecordCount
    tmp = "Select Mst_GSL_ID , ISINID from Mst_GSL Where Mst_GSL_ID = " & TmpRs2!Mst_GSL_ID & " and ISINID is not null And SM_PRefix = '" & mFSM_Prefix & "' "
    Call Tmp3Table
    If TmpRs3.RecordCount > 0 Then
       mISINID = TmpRs3!ISINID
       mTxn_GSL_ID = TmpRs3!Mst_GSL_ID
    
       tmp = "Update Txn_InvMonGSLDetail Set ISINID = '" & mISINID & "' Where ExchangeTypeID = 2 and YearMonth = '" & mFromDate & "' and ISINID is null And Mst_GSL_ID = " & mTxn_GSL_ID & ""
       Call Tmp3Table
    End If
    TmpRs2.MoveNext
Next

    
    
tmp = " Select TMCF.LotID, sum(TMCF.Amount) Amount from Txn_MonthlyCommtrackFileDetail TMCF"
tmp = tmp & " where TMCF.Amount > 0 And TMCF.monthyear between '" & mFromDate & "' and '" & mToDate & "'"
tmp = tmp & " group by TMCF.LotID"
Call Tmp2Table
   
tmp = " select * from Txn_InvMonGSLDetail "
tmp = tmp & " where exchangetypeid = 2 and yearmonth between '" & mFromDate & "' and '" & mToDate & "'"
tmp = tmp & " and commtrackflag = 1 and isnull(isinid,'') <> '' and SM_PRefix = '" & mFSM_Prefix & "' "

Call Tmp3Table

If TmpRs2.RecordCount > 0 And TmpRs3.RecordCount > 0 Then
    For pqr = 1 To TmpRs2.RecordCount
        TmpRs3.Find "ISINID = '" & TmpRs2!LotId & "'"   'TmpRs3.Filter = "ISINID = '" & TmpRs2!LotId & "'"
        If TmpRs3.EOF = False Then
           TmpRs3!InvoiceAmtForFranch = TmpRs3!InvoiceAmtForFranch + TmpRs2!Amount
           TmpRs3.Update
        End If
        TmpRs3.MoveFirst
        TmpRs2.MoveNext
    Next
End If


''Code for Partial reversal

tmp = " Select TCA.LocationID,sum(RecOtherDed) RecOtherDed"
tmp = tmp & " From Txn_CRTDSAdjustmentDetail TCAD"
tmp = tmp & " Inner Join Txn_CRTDSAdjustment TCA on TCA.CRTDSAdjustmentID = TCAD.CRTDSAdjustmentID And TCA.SM_Prefix = TCAD.SM_Prefix"
tmp = tmp & " Where TCA.Flag in ('A', 'P') And AdjDate Between '" & mFromDate & "' And '" & mToDate & "'"
tmp = tmp & " And TCAD.ExpenseAccID  in (90,91,114,88,112,92,113,89) And TCA.LocationID in (Select LocationID from Mst_Location Where SM_Prefix = '" & mFSM_Prefix & "')"  '91 means 7 Days Reversal"
tmp = tmp & " group by TCA.LocationID"

Call Tmp5Table

    If TmpRs5.RecordCount > 0 Then
       
       tmp = " Select Sum(InvoiceAmtForFranch) 'TotInvoiceAmtForFranch' from Txn_InvMonGSLDetail "
       tmp = tmp & " where yearmonth between '" & mFromDate & "' and '" & mToDate & "' and SM_PRefix = '" & mFSM_Prefix & "' ANd InvoiceAmtForFranch <> 0 "

       Call Tmp3Table
       
       mTotInvoiceAmtForFranch = IIf(IsNull(TmpRs3!TotInvoiceAmtForFranch), 0, TmpRs3!TotInvoiceAmtForFranch)
       
       tmp = " Select * from Txn_InvMonGSLDetail "
       tmp = tmp & " where yearmonth between '" & mFromDate & "' and '" & mToDate & "' and SM_PRefix = '" & mFSM_Prefix & "' ANd InvoiceAmtForFranch <> 0 "

       Call Tmp3Table
       
       For pqr = 1 To TmpRs3.RecordCount
           mRecOtherDed = Format(Val(IIf(IsNull(TmpRs5!RecOtherDed), "", TmpRs5!RecOtherDed)) * (Val(TmpRs3!InvoiceAmtForFranch) / Val(mTotInvoiceAmtForFranch)), "########.00")
          'mRecOtherDed = Format(Val(Val(IIf(IsNull(TmpRs5!RecOtherDed), "", TmpRs5!RecOtherDed)) * Val(MSHFlexGrid1.TextMatrix(i, E.GodownCapacity))) / Val(mTotLocCap), "########.00")
           TmpRs3!InvoiceAmtForFranch = TmpRs3!InvoiceAmtForFranch - Val(mRecOtherDed)
           TmpRs3.Update
           TmpRs3.MoveNext
       Next pqr
    End If


End Sub


Private Sub Grid_HeadGodownList()

With MSHFlexGrid2
   .Clear
   .Rows = 3
   .FixedRows = 2
   .Cols = 11
   
   .TextMatrix(1, 0) = "Month Year"
   .TextMatrix(1, 1) = "State Name"
   .TextMatrix(1, 2) = "Location Name"
   .TextMatrix(1, 3) = "CMP Name"
   .TextMatrix(1, 4) = "Godown No"
   .TextMatrix(1, 5) = "Godown Capacity"
   .TextMatrix(1, 6) = "Unit Type"
   .TextMatrix(1, 7) = "StateID"
   .TextMatrix(1, 8) = "LocationID"
   .TextMatrix(1, 9) = "GodownID"
   .TextMatrix(1, 10) = "Amount"
   
   .ColWidth(0) = 1500
   .ColWidth(1) = 1800
   .ColWidth(2) = 1800
   .ColWidth(3) = 2800
   .ColWidth(4) = 2800
   .ColWidth(5) = 1800
   .ColWidth(6) = 1000
   .ColWidth(7) = 1000
   .ColWidth(8) = 1000
   .ColWidth(9) = 1000
   .ColWidth(10) = 1000
   .ColWidth(11) = 1000

   .WordWrap = True
   .RowHeight(1) = 700
End With
End Sub


Private Sub GetReservationIncome(mFSM_Prefix As Variant)


    Call Grid_HeadGodownList
    
    tmp = "Select Distinct GodownID"
    tmp = tmp & " From ("
    tmp = tmp & " Select MS.StateName, ML.LocationID, ML.LocationName, MCMP.CMPName, MG.GodownID, MG.GodownNo, MG.Address, MG.GodownCapacity, MUT.UnitType, MG.StartDate, MG.CloseDate,  MS.StateID"
    tmp = tmp & " From Mst_Godown MG"
    tmp = tmp & " Inner Join Mst_UnitType MUT on MUT.UnitTypeID = MG.UNitTypeID"
    tmp = tmp & " Inner Join Mst_CMP MCMP on MCMP.CMPID = MG.CMPID"
    tmp = tmp & " Inner Join Mst_Location ML on ML.LocationID = MCMP.LocationID"
    tmp = tmp & " Inner Join Mst_State MS on MS.StateID = ML.StateID"
    tmp = tmp & " Where (MG.CloseDate is null Or MG.CloseDate >= '" & mFromDate & "') And MG.StartDate <= '" & mToDate & "' And ML.SM_Prefix = '" & mFSM_Prefix & "' "
    tmp = tmp & " Union"
    tmp = tmp & " Select MS.StateName, ML.LocationID, ML.LocationName, MCMP.CMPName, MG.GodownID, MG.GodownNo, MG.Address, MG.GodownCapacity, MUT.UnitType, MG.StartDate, MG.CloseDate,  MS.StateID"
    tmp = tmp & " From Mst_Godown MG"
    tmp = tmp & " Inner join Txn_InvMonGSLDetail TIG on TIG.GodownID=MG.GodownID"
    tmp = tmp & " Inner Join Mst_UnitType MUT on MUT.UnitTypeID = MG.UNitTypeID"
    tmp = tmp & " Inner Join Mst_CMP MCMP on MCMP.CMPID = MG.CMPID"
    tmp = tmp & " Inner Join Mst_Location ML on ML.LocationID = MCMP.LocationID"
    tmp = tmp & " Inner Join Mst_State MS on MS.StateID = ML.StateID"
    tmp = tmp & " Where TIG.YearMonth between '" & mFromDate & "' And '" & mToDate & "' And ML.SM_Prefix = '" & mFSM_Prefix & "' And TIG.BalanceBags > 0 And (MG.CloseDate is null Or MG.CloseDate >= '" & mFromDate & "') And MG.StartDate <= '" & mToDate & "') A"
    tmp = tmp & " Order by GodownID"
    Call Tmp1Table
    
    mGodownIDs = ""
    For pqr = 1 To TmpRs1.RecordCount
        If mGodownIDs = "" Then
           mGodownIDs = "(" & TmpRs1!GodownID
        Else
           mGodownIDs = mGodownIDs & "," & TmpRs1!GodownID
        End If
        TmpRs1.MoveNext
    Next
    
    If mGodownIDs <> "" Then
       mGodownIDs = mGodownIDs & ")"
    End If
    
    tmp = " Select MS.StateName, ML.LocationID, ML.LocationName, MCMP.CMPName, MG.GodownID, MG.GodownNo, MG.Address, MG.GodownCapacity, MUT.UnitType, MG.StartDate, MG.CloseDate,  MS.StateID "
    tmp = tmp & " From Mst_Godown MG"
    tmp = tmp & " Inner Join Mst_UnitType MUT on MUT.UnitTypeID = MG.UNitTypeID"
    tmp = tmp & " Inner Join Mst_CMP MCMP on MCMP.CMPID = MG.CMPID"
    tmp = tmp & " Inner Join Mst_Location ML on ML.LocationID = MCMP.LocationID"
    tmp = tmp & " Inner Join Mst_State MS on MS.StateID = ML.StateID"
    tmp = tmp & " Where MG.GodownID in " & mGodownIDs & ""
    tmp = tmp & " Order by MS.StateName, ML.LocationName, MCMP.CMPName, MG.GodownNo"
    
    Call Tmp1Table
    
    If TmpRs1.RecordCount > 0 Then
       MSHFlexGrid2.Rows = TmpRs1.RecordCount + 2

       For pqr = 2 To TmpRs1.RecordCount + 1
          With MSHFlexGrid2
             .TextMatrix(pqr, 0) = Format(mFromDate, "MMM-YYYY") 'mYearMonth
             .TextMatrix(pqr, 1) = IIf(IsNull(TmpRs1!StateName), "", TmpRs1!StateName)
             .TextMatrix(pqr, 2) = IIf(IsNull(TmpRs1!Locationname), "", TmpRs1!Locationname)
             .TextMatrix(pqr, 3) = IIf(IsNull(TmpRs1!CMPName), "", TmpRs1!CMPName)
             .TextMatrix(pqr, 4) = IIf(IsNull(TmpRs1!GodownNo), "", TmpRs1!GodownNo)
             .TextMatrix(pqr, 5) = IIf(IsNull(TmpRs1!GodownCapacity), "", TmpRs1!GodownCapacity)
             .TextMatrix(0, 5) = Val(.TextMatrix(0, 5)) + Val(.TextMatrix(pqr, 5))
             .TextMatrix(pqr, 6) = IIf(IsNull(TmpRs1!UnitType), "", TmpRs1!UnitType)
             .TextMatrix(pqr, 7) = IIf(IsNull(TmpRs1!StateID), "", TmpRs1!StateID)
             .TextMatrix(pqr, 8) = IIf(IsNull(TmpRs1!LocationID), "", TmpRs1!LocationID)
             .TextMatrix(pqr, 9) = IIf(IsNull(TmpRs1!GodownID), "", TmpRs1!GodownID)
             .TextMatrix(pqr, 10) = 0

          End With
          TmpRs1.MoveNext
       Next pqr
    End If


tmp = " Select Distinct TI.ContractID,ML.LocationID,Isnull(TI.ActualAmount,0) 'Amount',Sum(MG.GodownCapacity) 'Capacity'"
tmp = tmp & " from Txn_ReservationInvoice TI"
tmp = tmp & " Inner Join Mst_StorageContract MSC On TI.ContractID = MSC.ContractID"
tmp = tmp & " Inner Join Txn_SRGodownMapping TSRG On MSC.ContractID = TSRG.ContractID"
tmp = tmp & " Inner Join Txn_SRGodownMapDetails TSRGD On TSRG.SRGodownMappingID = TSRGD.SRGodownMappingID"
tmp = tmp & " Inner Join Mst_Godown MG On TSRGD.GodownID = MG.GodownID"
tmp = tmp & " Inner Join Mst_CMP MC On MG.CMPID = MC.CMPID"
tmp = tmp & " Inner Join Mst_Location ML on ML.LocationID = MC.LocationID"
tmp = tmp & " Inner Join Mst_State MS on MS.StateID = ML.StateID "
tmp = tmp & " Where  isnull(TI.ReverseFlag, '') <> 'Y' And MG.GodownID in " & mGodownIDs & " "
tmp = tmp & " And TI.Flag in ('A', 'P','F') And TI.YearMonth Between '" & mFromDate & "' and '" & mToDate & "' and Isnull(TI.ToBeExcludeForFranchaiseeCal,0) = 0"
tmp = tmp & " Group by  ML.LocationID ,TI.ContractID,TI.ActualAmount"
tmp = tmp & " Order by  ML.LocationID"

Call Tmp1Table

For pqr = 1 To TmpRs1.RecordCount
    mInvoiceAmount = IIf(IsNull(TmpRs1!Amount), 0, TmpRs1!Amount)
    
    mTotCapacity = 0
    
    tmp = " Select Distinct TI.ContractID,Sum(MG.GodownCapacity) 'Capacity'"
    tmp = tmp & " from Txn_ReservationInvoice TI"
    tmp = tmp & " Inner Join Mst_StorageContract MSC On TI.ContractID = MSC.ContractID"
    tmp = tmp & " Inner Join Txn_SRGodownMapping TSRG On MSC.ContractID = TSRG.ContractID"
    tmp = tmp & " Inner Join Txn_SRGodownMapDetails TSRGD On TSRG.SRGodownMappingID = TSRGD.SRGodownMappingID"
    tmp = tmp & " Inner Join Mst_Godown MG On TSRGD.GodownID = MG.GodownID"
    tmp = tmp & " Inner Join Mst_CMP MC On MG.CMPID = MC.CMPID"
    tmp = tmp & " Inner Join Mst_Location ML on ML.LocationID = MC.LocationID"
    tmp = tmp & " Inner Join Mst_State MS on MS.StateID = ML.StateID Where  isnull(TI.ReverseFlag, '') <> 'Y'"
    tmp = tmp & " And MG.GodownID in " & mGodownIDs & " And TI.ContractID = " & TmpRs1!ContractID & " "
    tmp = tmp & " And TI.Flag in ('A', 'P','F') And TI.YearMonth Between '" & mFromDate & "' and '" & mToDate & "' and Isnull(TI.ToBeExcludeForFranchaiseeCal,0) = 0"
    tmp = tmp & " Group by  TI.ContractID"
    tmp = tmp & " Order by  TI.ContractID"

    Call Tmp4Table
    
    If TmpRs4.RecordCount > 0 Then
       mTotCapacity = IIf(IsNull(TmpRs4!Capacity), 0, TmpRs4!Capacity)
    End If
    
    tmp = " Select Distinct TI.ContractID,ML.LocationID,MG.GodownID,MG.GodownCapacity"
    tmp = tmp & " from Txn_ReservationInvoice TI"
    tmp = tmp & " Inner Join Mst_StorageContract MSC On TI.ContractID = MSC.ContractID"
    tmp = tmp & " Inner Join Txn_SRGodownMapping TSRG On MSC.ContractID = TSRG.ContractID"
    tmp = tmp & " Inner Join Txn_SRGodownMapDetails TSRGD On TSRG.SRGodownMappingID = TSRGD.SRGodownMappingID"
    tmp = tmp & " Inner Join Mst_Godown MG On TSRGD.GodownID = MG.GodownID"
    tmp = tmp & " Inner Join Mst_CMP MC On MG.CMPID = MC.CMPID"
    tmp = tmp & " Inner Join Mst_Location ML on ML.LocationID = MC.LocationID"
    tmp = tmp & " Inner Join Mst_State MS on MS.StateID = ML.StateID Where  isnull(TI.ReverseFlag, '') <> 'Y'"
    tmp = tmp & " And MG.GodownID in " & mGodownIDs & " And TI.ContractID = " & TmpRs1!ContractID & " "
    tmp = tmp & " And TI.Flag in ('A', 'P','F') And TI.YearMonth Between '" & mFromDate & "' and '" & mToDate & "' and Isnull(TI.ToBeExcludeForFranchaiseeCal,0) = 0"
    tmp = tmp & " Order by  ML.LocationID"

    Call Tmp4Table
    
    If TmpRs4.RecordCount > 0 Then
        For y = 1 To TmpRs4.RecordCount
            If Val(mInvoiceAmount) > 0 Then
               For j = 2 To MSHFlexGrid2.Rows - 1
                   If MSHFlexGrid2.TextMatrix(j, 9) = TmpRs4!GodownID Then
                      If mTotCapacity > 0 Then
                         mInvAmt = Val(mInvoiceAmount) * Val(MSHFlexGrid2.TextMatrix(j, 5)) / Val(mTotCapacity)
                         mInvAmt = Format(mInvAmt, "#######0.00")
                         MSHFlexGrid2.TextMatrix(j, 10) = Val(MSHFlexGrid2.TextMatrix(j, 10)) + IIf(IsNull(mInvAmt), 0, mInvAmt) 'IIf(IsNull(TmpRs1!IncomeAmt), 0, TmpRs1!IncomeAmt)
                      End If
                   End If
               Next j
            End If
            TmpRs4.MoveNext
        Next y
    Else
        tmp = " Select Distinct TI.ContractID,ML.LocationID,Sum(MG.GodownCapacity) 'Capacity' "
        tmp = tmp & " from Txn_ReservationInvoice TI"
        tmp = tmp & " Inner Join Mst_StorageContract MSC On TI.ContractID = MSC.ContractID"
        tmp = tmp & " Inner Join Txn_SRGodownMapping TSRG On MSC.ContractID = TSRG.ContractID"
        tmp = tmp & " Inner Join Txn_SRGodownMapDetails TSRGD On TSRG.SRGodownMappingID = TSRGD.SRGodownMappingID"
        tmp = tmp & " Inner Join Mst_Godown MG On TSRGD.GodownID = MG.GodownID"
        tmp = tmp & " Inner Join Mst_CMP MC On MG.CMPID = MC.CMPID"
        tmp = tmp & " Inner Join Mst_Location ML on ML.LocationID = MC.LocationID"
        tmp = tmp & " Inner Join Mst_State MS on MS.StateID = ML.StateID Where  isnull(TI.ReverseFlag, '') <> 'Y'"
        tmp = tmp & " And TI.ContractID = " & TmpRs1!ContractID & " "
        tmp = tmp & " And TI.Flag in ('A', 'P','F') And TI.YearMonth Between '" & mFromDate & "' and '" & mToDate & "' and Isnull(TI.ToBeExcludeForFranchaiseeCal,0) = 0"
        tmp = tmp & " Group by TI.ContractID,ML.LocationID"
        tmp = tmp & " Order by  ML.LocationID"
    
        Call Tmp5Table
    
        If TmpRs5.RecordCount > 0 Then
           For y = 1 To TmpRs5.RecordCount
               If Val(mInvoiceAmount) > 0 Then
                  mTotCapacity = 0
                  For j = 2 To MSHFlexGrid2.Rows - 1
                      If TmpRs5!LocationID = (MSHFlexGrid2.TextMatrix(j, 7)) Then
                         mTotCapacity = mTotCapacity + Val(MSHFlexGrid2.TextMatrix(j, 5))
                      End If
                  Next
                  For j = 2 To MSHFlexGrid2.Rows - 1
                      If MSHFlexGrid2.TextMatrix(j, 7) = TmpRs5!LocationID Then
                         If mTotCapacity > 0 Then
                            mInvAmt = Val(mInvoiceAmount) * Val(MSHFlexGrid2.TextMatrix(j, 5)) / Val(mTotCapacity)
                            mInvAmt = Format(mInvAmt, "#######0.00")
                            MSHFlexGrid2.TextMatrix(j, 10) = Val(MSHFlexGrid2.TextMatrix(j, 10)) + IIf(IsNull(mInvAmt), 0, mInvAmt) 'IIf(IsNull(TmpRs1!IncomeAmt), 0, TmpRs1!IncomeAmt)
                         End If
                      End If
                  Next j
               End If
               TmpRs5.MoveNext
           Next y
       End If
    
    End If
    TmpRs1.MoveNext
Next pqr


tmp = " Select Distinct ML.LocationID,Sum(Isnull(TI.ActualAmount,0)) 'Amount',Sum(MG.GodownCapacity) 'Capacity'"
tmp = tmp & " from Txn_ReservationInvoice TI"
tmp = tmp & " Inner Join Mst_StorageContract MSC On TI.ContractID = MSC.ContractID"
tmp = tmp & " Left Join Txn_SRGodownMapping TSRG On MSC.ContractID = TSRG.ContractID"
tmp = tmp & " Left Join Txn_SRGodownMapDetails TSRGD On TSRG.SRGodownMappingID = TSRGD.SRGodownMappingID"
tmp = tmp & " Left Join Mst_Godown MG On TSRGD.GodownID = MG.GodownID"
tmp = tmp & " Left Join Mst_CMP MC On MG.CMPID = MC.CMPID"
tmp = tmp & " Left Join Mst_Location ML on ML.LocationID = MSC.LocationID"
tmp = tmp & " Left Join Mst_State MS on MS.StateID = ML.StateID Where  isnull(TI.ReverseFlag, '') <> 'Y'"
tmp = tmp & " And TI.Flag in ('A', 'P','F') And TI.YearMonth Between '" & mFromDate & "' and '" & mToDate & "' and Isnull(TI.ToBeExcludeForFranchaiseeCal,0) = 0"
tmp = tmp & " And MG.GodownCapacity is null"
tmp = tmp & " Group by  ML.LocationID"
tmp = tmp & " Order by  ML.LocationID"


Call Tmp5Table

If TmpRs5.RecordCount > 0 Then
   For y = 1 To TmpRs5.RecordCount
       mInvoiceAmount = IIf(IsNull(TmpRs5!Amount), 0, TmpRs5!Amount)
       If Val(mInvoiceAmount) > 0 Then
          mTotCapacity = 0
          For j = 2 To MSHFlexGrid2.Rows - 1
              If TmpRs5!LocationID = (MSHFlexGrid2.TextMatrix(j, 7)) Then
                 mTotCapacity = mTotCapacity + Val(MSHFlexGrid2.TextMatrix(j, 5))
              End If
          Next
          For j = 2 To MSHFlexGrid2.Rows - 1
              If MSHFlexGrid2.TextMatrix(j, 7) = TmpRs5!LocationID Then
                 If mTotCapacity > 0 Then
                    mInvAmt = Val(mInvoiceAmount) * Val(MSHFlexGrid2.TextMatrix(j, 5)) / Val(mTotCapacity)
                    mInvAmt = Format(mInvAmt, "#######0.00")
                    MSHFlexGrid2.TextMatrix(j, 10) = Val(MSHFlexGrid2.TextMatrix(j, 10)) + IIf(IsNull(mInvAmt), 0, mInvAmt) 'IIf(IsNull(TmpRs1!IncomeAmt), 0, TmpRs1!IncomeAmt)
                 End If
              End If
          Next j
       End If
       TmpRs5.MoveNext
   Next y
End If

'
'        tmp = " Select ML.LocationID, sum(MG.GodownCapacity) Capacity"
'        tmp = tmp & " From Mst_Godown MG"
'        tmp = tmp & " Inner Join Mst_CMP MCMP on MCMP.CMPID = MG.CMPID"
'        tmp = tmp & " Inner Join Mst_Location ML on ML.LocationID = MCMP.LocationID"
'        tmp = tmp & " Inner Join Mst_State MS on MS.StateID = ML.StateID"
'        tmp = tmp & " Where MG.GodownID in " & mGodownIDs & ""
'        tmp = tmp & " group by ML.LocationID"
'        Call Tmp2Table
'
'


tmp = " Select TISGD.GodownID, sum(TISGD.InvoiceAmt) IncomeAmt"
tmp = tmp & " from Txn_TempInvoice TI"
tmp = tmp & " Inner Join Txn_TempInvSRGodownDetails TISGD on TI.PINo = TISGD.PINo"
tmp = tmp & " Where TI.Flag in ('A', 'P','F') And TI.IETypeID = 4 And isnuLL(TI.ReverseFlag,'') <> 'Y' and Isnull(TI.ToBeExcludeForFranchaiseeCal,0) = 0 "
tmp = tmp & " And TI.InvoiceDate Between  '" & mFromDate & "' And '" & mToDate & "' And TISGD.GodownID in " & mGodownIDs & ""
tmp = tmp & " Group by TISGD.GodownID"

Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   For x = 1 To TmpRs1.RecordCount
       mStatus = 0
       For xy = 2 To MSHFlexGrid2.Rows - 1
           If MSHFlexGrid2.TextMatrix(xy, 9) = TmpRs1!GodownID Then
              mStatus = 1
              MSHFlexGrid2.TextMatrix(xy, 10) = Val(MSHFlexGrid2.TextMatrix(xy, 10)) + TmpRs1!IncomeAmt 'IIf(IsNull(TmpRs1!IncomeAmt), 0, Format(TmpRs1!IncomeAmt * Val(MSHFlexGrid2.TextMatrix(pqr, E.GodownCapacity)) / TmpRs2!Capacity, "#############.00"))
           End If
       Next
       TmpRs1.MoveNext
   Next
End If

For xy = 2 To MSHFlexGrid2.Rows - 1
    tmp = "Select * from Txn_InvMonGSLDetail Where SM_Prefix = '" & mFSM_Prefix & "' And  Yearmonth = '" & mFromDate & "' And GodownID = " & Val(MSHFlexGrid2.TextMatrix(xy, 9)) & "   And PINo is not null   and COntractID is not null "
    Call Tmp1Table
    If TmpRs1.RecordCount > 0 Then
       For pqr = 1 To TmpRs1.RecordCount
           TmpRs1!InvoiceAmtForFranch = TmpRs1!InvoiceAmtForFranch + (Val(MSHFlexGrid2.TextMatrix(xy, 10)) / TmpRs1.RecordCount)
           TmpRs1.Update
           TmpRs1.MoveNext
       Next pqr
    
    Else
        If LCase(MSHFlexGrid2.TextMatrix(xy, 6)) = LCase("Franchisee") Then
            tmp = "Select * from Txn_InvMonGSLDetail Where SM_Prefix = '" & mFSM_Prefix & "' And  Yearmonth = '" & mFromDate & "' And GodownID in " & mGodownIDs & "   And PINo is not null   and COntractID is not null And SM_Prefix = '" & mFSM_Prefix & "' And GodownID in (Select GodownID from Mst_Godown Where UnitTypeID = 2) "
            Call Tmp1Table
            If TmpRs1.RecordCount > 0 Then
               For pqr = 1 To TmpRs1.RecordCount
                   TmpRs1!InvoiceAmtForFranch = TmpRs1!InvoiceAmtForFranch + (Val(MSHFlexGrid2.TextMatrix(xy, 10)) / TmpRs1.RecordCount)
                   TmpRs1.Update
                   TmpRs1.MoveNext
               Next pqr
            End If
        ElseIf LCase(MSHFlexGrid2.TextMatrix(xy, 6)) = LCase("Leased") Then
        
            tmp = "Select * from Txn_InvMonGSLDetail Where SM_Prefix = '" & mFSM_Prefix & "' And  Yearmonth = '" & mFromDate & "' And GodownID in " & mGodownIDs & "  And PINo is not null   and COntractID is not null And SM_Prefix = '" & mFSM_Prefix & "' And GodownID in (Select GodownID from Mst_Godown Where UnitTypeID = 1) "
            Call Tmp1Table
            If TmpRs1.RecordCount > 0 Then
               For pqr = 1 To TmpRs1.RecordCount
                   TmpRs1!InvoiceAmtForFranch = TmpRs1!InvoiceAmtForFranch + (Val(MSHFlexGrid2.TextMatrix(xy, 10)) / TmpRs1.RecordCount)
                   TmpRs1.Update
                   TmpRs1.MoveNext
               Next pqr
            End If
        End If
    End If
Next xy

'Gen_mTimeStamp = GetTimeStamp
'Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Txn_ResIncome.xls")
'Call Xls_Detail_Rpt

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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Txn_ResIncome.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For pqr = 0 To MSHFlexGrid2.Rows - 1
    mRow = pqr + 1 ' + 6
    For C = 1 To MSHFlexGrid2.Cols - 1
        If InStr(MSHFlexGrid2.TextMatrix(pqr, C), "/") > 0 Then
           If IsDate(MSHFlexGrid2.TextMatrix(pqr, C)) Then
              oWS.Cells(mRow, C) = Format(MSHFlexGrid2.TextMatrix(pqr, C), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, C) = MSHFlexGrid2.TextMatrix(pqr, C)
           End If
        Else
           oWS.Cells(mRow, C) = MSHFlexGrid2.TextMatrix(pqr, C)
        End If
    Next
Next

oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub



Private Sub CmbOldLocation_GotFocus()

tmp = CmbOldLocation.Text
If Gen_WcLocation <> "" Then
   Call TableMst_Location(" Where " & Gen_WcLocation)
Else
   Call TableMst_Location("")
End If
CmbOldLocation.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Location found"
   CmbOldLocation.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbOldLocation.AddItem RsMst_Location!Locationname
    RsMst_Location.MoveNext
Next
CmbOldLocation.Text = tmp

End Sub

Private Sub CmbOldLocation_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "CmbOldLocation")
End If
End Sub

Private Sub CmbOldLocation_LostFocus()

If CmbOldLocation.Text = "" Then
   mLocationID = 0
   TxtPreFix = ""
   mSM_Prefix = ""
   CmbCMPName.Text = ""
   CmbGodownNo.Text = ""
   TxtNewGodowNo.Text = ""
   Exit Sub
End If
' Code for clearing Godown combox in case of change
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbOldLocation.Text & "'"

If RsMst_Location.EOF Then
   MsgBox "Invalid Location  selection "
   CmbOldLocation.SetFocus
   Exit Sub
End If
If mLocationID <> RsMst_Location!LocationID Then
   CmbCMPName.Text = ""
   CmbGodownNo.Text = ""
   TxtNewGodowNo.Text = ""
End If
mLocationID = RsMst_Location!LocationID
End Sub


Private Sub CmbCMPName_GotFocus()
If CmbOldLocation.Text = "" Then
   MsgBox "Select Location!!!"
   If RsMst_Location.RecordCount > 0 Then
      CmbOldLocation.SetFocus
   End If
   Exit Sub
End If
tmp = CmbCMPName.Text
If Gen_WcCMP <> "" Then
   Call TableMst_CMP(" Where CloseDate is null And LocationID = " & mLocationID & " And " & Gen_WcCMP)
Else
   Call TableMst_CMP(" Where CloseDate is null And LocationID = " & mLocationID)
End If
CmbCMPName.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
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
   CmbGodownNo.Text = ""
   mCMPID = 0
   Exit Sub
End If

RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPName = '" & CmbCMPName.Text & "'"
If RsMst_CMP.EOF Then
   MsgBox "Invalid selection "
   CmbCMPName.SetFocus
   Exit Sub
End If
If mCMPID <> RsMst_CMP!CMPID Then
   CmbGodownNo.Text = ""
End If
mCMPID = RsMst_CMP!CMPID

End Sub

Private Sub CmbGodownNo_GotFocus()
If CmbCMPName.Text = "" Then
   MsgBox "Select CMP !!!! "
   CmbCMPName.SetFocus
   Exit Sub
End If
tmp = CmbGodownNo.Text
Call TableMst_Godown("Where CMPID = " & mCMPID)
CmbGodownNo.Clear
If RsMst_Godown.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Godown.RecordCount
    CmbGodownNo.AddItem RsMst_Godown!GodownNo
    RsMst_Godown.MoveNext
Next
CmbGodownNo.Text = tmp
End Sub

Private Sub CmbGodownNo_LostFocus()
If CmbGodownNo.Text = "" Then
   mGodownID = 0
   Exit Sub
End If
RsMst_Godown.MoveFirst
RsMst_Godown.Find "GodownNo = '" & CmbGodownNo.Text & "'"
If RsMst_Godown.EOF Then
   MsgBox "Invalid selection "
   If CmbGodownNo.Enabled = True Then
      CmbGodownNo.SetFocus
   End If
   Exit Sub
End If
mGodownID = RsMst_Godown!GodownID

End Sub

