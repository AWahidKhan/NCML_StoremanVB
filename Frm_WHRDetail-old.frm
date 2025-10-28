VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Frm_WHRDetail 
   Caption         =   "WHR Detail"
   ClientHeight    =   9750
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15075
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9750
   ScaleWidth      =   15075
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   2775
      Left            =   120
      TabIndex        =   97
      Top             =   2160
      Visible         =   0   'False
      Width           =   14535
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1695
         Left            =   120
         TabIndex        =   115
         Top             =   1200
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   2990
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.CommandButton CmdGo 
         Caption         =   "Go"
         Height          =   345
         Left            =   10680
         TabIndex        =   109
         Top             =   720
         Width           =   1245
      End
      Begin VB.ComboBox CmbSearchStatus 
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
         ItemData        =   "Frm_WHRDetail.frx":0000
         Left            =   8760
         List            =   "Frm_WHRDetail.frx":000A
         Sorted          =   -1  'True
         TabIndex        =   108
         Top             =   720
         Width           =   1500
      End
      Begin VB.OptionButton OptAll 
         Caption         =   "All"
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
         TabIndex        =   107
         Top             =   240
         Value           =   -1  'True
         Width           =   585
      End
      Begin VB.OptionButton OptLocationName 
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3480
         TabIndex        =   106
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton OptGodown 
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   7080
         TabIndex        =   105
         Top             =   240
         Width           =   1185
      End
      Begin VB.TextBox TxtSearchWHR 
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
         Left            =   3120
         TabIndex        =   104
         Top             =   720
         Width           =   4815
      End
      Begin VB.OptionButton OptCMPName 
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5430
         TabIndex        =   103
         Top             =   240
         Width           =   1470
      End
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   350
         Left            =   13440
         TabIndex        =   102
         Top             =   720
         Width           =   1245
      End
      Begin VB.OptionButton OptCommodity 
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   9840
         TabIndex        =   101
         Top             =   240
         Width           =   1515
      End
      Begin VB.OptionButton OptWHRDate 
         Caption         =   "WHR Date"
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
         Left            =   840
         TabIndex        =   100
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton OptWHRNo 
         Caption         =   "WHR No"
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
         Left            =   2280
         TabIndex        =   99
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton OptClientName 
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   8280
         TabIndex        =   98
         Top             =   240
         Width           =   1470
      End
      Begin VB.Label Label24 
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
         Height          =   375
         Left            =   8160
         TabIndex        =   114
         Top             =   720
         Width           =   1335
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
         Left            =   11640
         TabIndex        =   113
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label25 
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
         Left            =   14160
         TabIndex        =   112
         Top             =   240
         Width           =   75
      End
      Begin VB.Label LblSearchWHR 
         Caption         =   "Search By Employee Name "
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
         Left            =   150
         TabIndex        =   111
         Top             =   705
         Width           =   2895
      End
      Begin VB.Label LblRecordCount 
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
         Left            =   14280
         TabIndex        =   110
         Top             =   240
         Width           =   75
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H008080FF&
      Caption         =   "Hologram No "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   5760
      TabIndex        =   55
      Top             =   3960
      Visible         =   0   'False
      Width           =   4695
      Begin VB.CommandButton CmdCancel 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cancel"
         Height          =   400
         Left            =   2880
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton CmdSaveHologram 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Save"
         Height          =   400
         Left            =   1440
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   840
         Width           =   1440
      End
      Begin VB.TextBox TxtHologramNo 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1440
         MaxLength       =   25
         TabIndex        =   56
         Top             =   360
         Width           =   2880
      End
      Begin VB.Label Label33 
         BackColor       =   &H008080FF&
         Caption         =   "Hologram No"
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
         TabIndex        =   59
         Top             =   390
         Width           =   1260
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1200
      Left            =   120
      TabIndex        =   31
      Top             =   8520
      Width           =   14895
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Print"
         Height          =   810
         Left            =   12390
         TabIndex        =   49
         Top             =   195
         Width           =   2295
      End
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   400
         Left            =   10065
         TabIndex        =   24
         Top             =   600
         Width           =   2340
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "Search"
         Height          =   400
         Left            =   10065
         TabIndex        =   23
         Top             =   195
         Width           =   2340
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   400
         Left            =   7605
         TabIndex        =   37
         Top             =   600
         Width           =   2460
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   400
         Left            =   7605
         TabIndex        =   36
         Top             =   195
         Width           =   2460
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   400
         Left            =   5145
         TabIndex        =   35
         Top             =   600
         Width           =   2460
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   400
         Left            =   5145
         TabIndex        =   34
         Top             =   195
         Width           =   2460
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   400
         Left            =   2565
         TabIndex        =   33
         Top             =   600
         Width           =   2580
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   400
         Left            =   2565
         TabIndex        =   22
         Top             =   195
         Width           =   2580
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   400
         Left            =   105
         TabIndex        =   32
         Top             =   600
         Width           =   2460
      End
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New"
         Height          =   400
         Left            =   105
         TabIndex        =   0
         Top             =   195
         Width           =   2460
      End
   End
   Begin VB.Frame Frame1 
      Height          =   8475
      Left            =   120
      TabIndex        =   25
      Top             =   0
      Width           =   14925
      Begin VB.CheckBox ChkEstimatedWgt 
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   14160
         TabIndex        =   92
         Top             =   2685
         Width           =   210
      End
      Begin VB.CheckBox ChkSaidToContain 
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   11640
         TabIndex        =   91
         Top             =   2685
         Width           =   225
      End
      Begin VB.TextBox TxtRisks 
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
         Height          =   555
         Left            =   3600
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   87
         TabStop         =   0   'False
         Top             =   3045
         Width           =   10815
      End
      Begin VB.TextBox TxtAmount 
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
         Left            =   8160
         Locked          =   -1  'True
         TabIndex        =   80
         TabStop         =   0   'False
         Top             =   2640
         Width           =   1680
      End
      Begin VB.CommandButton CmdInsurance 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   14460
         TabIndex        =   77
         Top             =   2220
         Width           =   360
      End
      Begin VB.ComboBox CmbPolicyNo 
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
         Left            =   3600
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   2220
         Width           =   10845
      End
      Begin VB.CommandButton CmdInsuranceCo 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   14445
         TabIndex        =   74
         Top             =   1800
         Width           =   375
      End
      Begin VB.ComboBox CmbInsuranceCo 
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
         Left            =   3600
         Sorted          =   -1  'True
         TabIndex        =   8
         Top             =   1800
         Width           =   10845
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   3855
         Left            =   240
         TabIndex        =   13
         Top             =   4440
         Width           =   14445
         _ExtentX        =   25479
         _ExtentY        =   6800
         _Version        =   393216
         Tabs            =   2
         TabHeight       =   520
         TabCaption(0)   =   "GSL Detail"
         TabPicture(0)   =   "Frm_WHRDetail.frx":0026
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "MSHFlexGrid2"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Frame2"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Other Details"
         TabPicture(1)   =   "Frm_WHRDetail.frx":0042
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame4"
         Tab(1).ControlCount=   1
         Begin VB.Frame Frame4 
            Height          =   3615
            Left            =   -74880
            TabIndex        =   18
            Top             =   360
            Width           =   14175
            Begin VB.TextBox TxtWeigher 
               Appearance      =   0  'Flat
               Height          =   360
               Left            =   3480
               MaxLength       =   50
               TabIndex        =   19
               Top             =   360
               Width           =   10440
            End
            Begin VB.TextBox TxtAssayer 
               Appearance      =   0  'Flat
               Height          =   360
               Left            =   3480
               MaxLength       =   50
               TabIndex        =   20
               Top             =   1020
               Width           =   10440
            End
            Begin VB.TextBox TxtAssayerCerti 
               Appearance      =   0  'Flat
               Height          =   360
               Left            =   3480
               MaxLength       =   50
               TabIndex        =   21
               Top             =   1665
               Width           =   10440
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               Caption         =   "Name and License No. of Assayer"
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
               TabIndex        =   90
               Top             =   1050
               Width           =   3180
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               Caption         =   "Name and License No. of Weigher"
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
               TabIndex        =   89
               Top             =   390
               Width           =   3300
            End
            Begin VB.Label Label5 
               Caption         =   "Grade as per Assayer’s Certificate"
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
               Left            =   360
               TabIndex        =   88
               Top             =   1695
               Width           =   3060
            End
         End
         Begin VB.Frame Frame2 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   870
            Left            =   120
            TabIndex        =   60
            Top             =   360
            Width           =   14250
            Begin VB.TextBox TxtCMPID 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000A&
               Height          =   360
               Left            =   60
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   95
               TabStop         =   0   'False
               Top             =   405
               Width           =   5130
            End
            Begin VB.TextBox TxtNoOfBags 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H8000000A&
               Height          =   360
               Left            =   1905
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   65
               TabStop         =   0   'False
               Top             =   885
               Visible         =   0   'False
               Width           =   1905
            End
            Begin VB.TextBox TxtLicenseNo 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000A&
               Height          =   360
               Left            =   300
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   64
               TabStop         =   0   'False
               Top             =   885
               Visible         =   0   'False
               Width           =   1575
            End
            Begin VB.TextBox TxtLicenseDescription 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000A&
               Height          =   360
               Left            =   7365
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   63
               TabStop         =   0   'False
               Top             =   885
               Visible         =   0   'False
               Width           =   3630
            End
            Begin VB.TextBox TxtEleWeight 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H8000000A&
               Height          =   360
               Left            =   3825
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   62
               TabStop         =   0   'False
               Top             =   885
               Visible         =   0   'False
               Width           =   1680
            End
            Begin VB.TextBox TxtWHRQty 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   360
               Left            =   10605
               MaxLength       =   50
               TabIndex        =   16
               Top             =   405
               Width           =   1740
            End
            Begin VB.CommandButton CmdUpdate 
               Caption         =   "Update List"
               Height          =   360
               Left            =   12435
               TabIndex        =   17
               Top             =   405
               Width           =   1680
            End
            Begin VB.TextBox TxtWHRNoOfBags 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   360
               Left            =   8880
               MaxLength       =   50
               TabIndex        =   15
               Top             =   405
               Width           =   1710
            End
            Begin VB.TextBox TxtGodownNo 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000A&
               Height          =   360
               Left            =   5220
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   61
               TabStop         =   0   'False
               Top             =   405
               Width           =   3600
            End
            Begin VB.Label Label17 
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
               Height          =   300
               Left            =   2055
               TabIndex        =   96
               Top             =   150
               Width           =   1140
            End
            Begin VB.Label Label15 
               Caption         =   "Available No Of Bags"
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
               Left            =   1860
               TabIndex        =   72
               Top             =   465
               Visible         =   0   'False
               Width           =   1965
            End
            Begin VB.Label Label11 
               Caption         =   "License No"
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
               Left            =   540
               TabIndex        =   71
               Top             =   510
               Visible         =   0   'False
               Width           =   1020
            End
            Begin VB.Label Label12 
               Caption         =   "License Description"
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
               Left            =   8235
               TabIndex        =   70
               Top             =   495
               Visible         =   0   'False
               Width           =   1860
            End
            Begin VB.Label Label13 
               Caption         =   "Electronic Weight"
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
               Left            =   3915
               TabIndex        =   69
               Top             =   465
               Visible         =   0   'False
               Width           =   1620
            End
            Begin VB.Label Label14 
               Caption         =   "WHR Qty"
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
               Left            =   11055
               TabIndex        =   68
               Top             =   120
               Width           =   900
            End
            Begin VB.Label Label16 
               Caption         =   "WHR No Of Bags"
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
               Left            =   9000
               TabIndex        =   67
               Top             =   120
               Width           =   1620
            End
            Begin VB.Label Label18 
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
               Height          =   300
               Left            =   6450
               TabIndex        =   66
               Top             =   120
               Width           =   1140
            End
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
            Height          =   2385
            Left            =   120
            TabIndex        =   14
            Top             =   1320
            Width           =   14220
            _ExtentX        =   25083
            _ExtentY        =   4207
            _Version        =   393216
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin VB.TextBox TxtWHRNoNew 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   375
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox TxtAgreeDate 
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
         Left            =   13680
         Locked          =   -1  'True
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   4080
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.TextBox TxtRate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   3600
         MaxLength       =   50
         TabIndex        =   11
         Top             =   4020
         Width           =   1800
      End
      Begin VB.TextBox TxtValueofCommodity 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Height          =   360
         Left            =   8160
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   4020
         Width           =   1680
      End
      Begin VB.TextBox TxtStorageAmount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Height          =   360
         Left            =   8160
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   3630
         Width           =   1680
      End
      Begin VB.TextBox TxtStorageCharges 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   3600
         MaxLength       =   50
         TabIndex        =   10
         Top             =   3630
         Width           =   1800
      End
      Begin VB.TextBox TxtGrade 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   11280
         MaxLength       =   50
         TabIndex        =   5
         Top             =   600
         Width           =   3120
      End
      Begin VB.ComboBox CmbCommodityCondition 
         Height          =   315
         ItemData        =   "Frm_WHRDetail.frx":005E
         Left            =   7800
         List            =   "Frm_WHRDetail.frx":006B
         TabIndex        =   4
         Top             =   600
         Width           =   3375
      End
      Begin VB.ComboBox CmbClientId 
         Height          =   315
         Left            =   3600
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   1080
         Width           =   10815
      End
      Begin VB.TextBox TxtTotalWHRQty 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   375
         Left            =   11640
         Locked          =   -1  'True
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   4080
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.ComboBox CmbWHRCommodity 
         Height          =   315
         Left            =   3600
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   1440
         Width           =   10815
      End
      Begin VB.TextBox TxtWHRNo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   375
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   1080
         Visible         =   0   'False
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker WHRDate 
         Height          =   330
         Left            =   1920
         TabIndex        =   1
         Top             =   600
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   582
         _Version        =   393216
         Format          =   49545217
         CurrentDate     =   39884
      End
      Begin MSComCtl2.DTPicker StorageFrom 
         Height          =   315
         Left            =   3600
         TabIndex        =   2
         Top             =   600
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         Format          =   49545217
         CurrentDate     =   39884
      End
      Begin MSComCtl2.DTPicker StorageTo 
         Height          =   315
         Left            =   5760
         TabIndex        =   3
         Top             =   600
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         Format          =   49545217
         CurrentDate     =   39884
      End
      Begin VB.TextBox TxtClientID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   330
         Left            =   3600
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   47
         Top             =   1080
         Width           =   10800
      End
      Begin VB.TextBox TxtWHRCommodityId 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   330
         Left            =   3600
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   48
         Top             =   1440
         Width           =   10800
      End
      Begin MSComCtl2.DTPicker GradeValidUpto 
         Height          =   375
         Left            =   11640
         TabIndex        =   12
         Top             =   3630
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         Format          =   49545217
         CurrentDate     =   39884
      End
      Begin MSComCtl2.DTPicker ValidFrom 
         Height          =   375
         Left            =   3600
         TabIndex        =   81
         TabStop         =   0   'False
         Top             =   2640
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   49545217
         CurrentDate     =   39884
      End
      Begin MSComCtl2.DTPicker ValidTo 
         Height          =   375
         Left            =   6120
         TabIndex        =   82
         TabStop         =   0   'False
         Top             =   2640
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   49545217
         CurrentDate     =   39884
      End
      Begin VB.TextBox TxtPolicyNo 
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
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   78
         Top             =   2220
         Width           =   10815
      End
      Begin VB.TextBox TxtInsuranceCo 
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
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   75
         Top             =   1800
         Width           =   10815
      End
      Begin VB.Label Label22 
         Caption         =   "Estimated Weight"
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
         Left            =   12360
         TabIndex        =   94
         Top             =   2685
         Width           =   1695
      End
      Begin VB.Label Label21 
         Caption         =   "Said To Contain"
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
         Left            =   10080
         TabIndex        =   93
         Top             =   2685
         Width           =   1455
      End
      Begin VB.Label Label43 
         Alignment       =   1  'Right Justify
         Caption         =   "Risks"
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
         TabIndex        =   86
         Top             =   3195
         Width           =   3375
      End
      Begin VB.Label Label42 
         Caption         =   "Amount"
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
         Left            =   7440
         TabIndex        =   85
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Label Label41 
         Alignment       =   1  'Right Justify
         Caption         =   "Valid From"
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
         TabIndex        =   84
         Top             =   2640
         Width           =   3375
      End
      Begin VB.Label Label40 
         Caption         =   "Valid To"
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
         Left            =   5160
         TabIndex        =   83
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label Label39 
         Alignment       =   1  'Right Justify
         Caption         =   "Policy No /  Cover note."
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
         TabIndex        =   79
         Top             =   2280
         Width           =   3375
      End
      Begin VB.Label Label38 
         Alignment       =   1  'Right Justify
         Caption         =   "Insurance Co."
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
         TabIndex        =   76
         Top             =   1845
         Width           =   3255
      End
      Begin VB.Label Label6 
         Caption         =   "Grade Valid Till"
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
         Left            =   9960
         TabIndex        =   73
         Top             =   3667
         Width           =   1470
      End
      Begin VB.Label Label20 
         Caption         =   "Value of Commodity"
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
         Left            =   5760
         TabIndex        =   52
         Top             =   4050
         Width           =   2160
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Caption         =   "Commodity Price (Rate Per MT)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   51
         Top             =   4050
         Width           =   3300
      End
      Begin VB.Label LblWHRCommodity 
         Alignment       =   1  'Right Justify
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
         Height          =   300
         Left            =   120
         TabIndex        =   46
         Top             =   1440
         Width           =   3300
      End
      Begin VB.Label LblClientName 
         Alignment       =   1  'Right Justify
         Caption         =   "Client Name"
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
         Left            =   120
         TabIndex        =   45
         Top             =   1080
         Width           =   3300
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Storage Rate Per Month Per Bag"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   44
         Top             =   3660
         Width           =   3300
      End
      Begin VB.Label Label9 
         Caption         =   "Date of Deposit "
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
         Left            =   3720
         TabIndex        =   43
         Top             =   240
         Width           =   1680
      End
      Begin VB.Label Label10 
         Caption         =   "Storage upto Date"
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
         Left            =   5775
         TabIndex        =   42
         Top             =   240
         Width           =   2040
      End
      Begin VB.Label Label8 
         Caption         =   "Storage Charge Amount"
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
         Left            =   5760
         TabIndex        =   41
         Top             =   3660
         Width           =   2160
      End
      Begin VB.Label Label3 
         Caption         =   "Private Marks on Packaging"
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
         Left            =   11280
         TabIndex        =   39
         Top             =   240
         Width           =   3180
      End
      Begin VB.Label Label1 
         Caption         =   "Condition of Commodities"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   8430
         TabIndex        =   38
         Top             =   255
         Width           =   2700
      End
      Begin VB.Label LblWHRDate 
         Caption         =   "WHR Date"
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
         Left            =   2160
         TabIndex        =   30
         Top             =   240
         Width           =   1020
      End
      Begin VB.Label LblTotalWHRQty 
         Caption         =   "Total WHR Qty"
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
         Left            =   10110
         TabIndex        =   28
         Top             =   4080
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Label LblWHRNo 
         Caption         =   "WHR No"
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
         Left            =   600
         TabIndex        =   27
         Top             =   240
         Width           =   900
      End
   End
End
Attribute VB_Name = "Frm_WHRDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mExchangeTypeID, mCommodityCondition, mWHRNoOfBags, mWHRQty As Double
Dim pq, temp, GSLQty, ChkSelection As Double
Dim mClientIDRow As Double
Dim mCommodityId As Double
Dim mGodownID As Double
Dim mCMPID As Double
Dim mWHRNo As Double
Dim PolicyNo_, CNo As String
Dim Edit_Flg, mLicenseNo, mInsuranceCoID As Variant
Private Sub AddCmd_Click()
TxtAgreeDate = ""
  Edit_Flg = "A"
  WHRDate.Enabled = True
  CmbCommodityCondition.Locked = False
  StorageFrom.Enabled = True
  StorageTo.Enabled = True
  TxtGrade.Locked = False
  TxtInsuranceCo.Visible = False
  TxtPolicyNo.Visible = False
  CmbInsuranceCo.Visible = True
  CmbPolicyNo.Visible = True
  CmbInsuranceCo.Text = ""
  CmbPolicyNo.Text = ""
  CmdInsuranceCo.Visible = True
  CmdInsurance.Visible = True
  TxtClientID.Visible = False
  CmbClientId.Visible = True
  CmbClientId.Locked = False
  TxtWHRCommodityId.Visible = False
  CmbWHRCommodity.Visible = True
  CmbWHRCommodity.Locked = False
  TxtWeigher.Locked = False
  TxtAssayer.Locked = False
  TxtAssayerCerti.Locked = False
  GradeValidUpto.Enabled = True
  TxtStorageCharges.Locked = False
  TxtRate.Locked = False
  TxtWHRNoOfBags.Locked = True
  TxtWHRQty.Locked = True
  TxtWHRNo = ""
  TxtPreFix = ""
  TxtWHRNoNew = ""
  CmbCommodityCondition.Text = ""
  mCommodityId = 0
  mClientIDRow = 0
  WHRDate.Value = Date
  StorageFrom.Value = Date
  StorageTo.Value = Date
  TxtGrade = ""
  CmbClientId.Text = ""
  CmbWHRCommodity.Text = ""
  TxtWeigher = ""
  TxtAssayer = ""
  TxtAssayerCerti = ""
  GradeValidUpto = Date
  TxtStorageCharges = ""
  TxtRate = ""
  TxtValueofCommodity = ""
  TxtAmount = ""
  TxtRisks = ""
  
  
  TxtStorageAmount = ""
  TxtTotalWHRQty = ""
  ChkSaidToContain.Enabled = True
  ChkEstimatedWgt.Enabled = True
  ChkSaidToContain.Value = 0
  ChkEstimatedWgt.Value = 0
  ChkSelection = 0
  
  Call Grid_Head
  CmdUpdate.Enabled = False
  Call GButtonLock(Me, False)
  CmdPrint.Enabled = False
  SSTab1.Tab = 0
End Sub

Private Sub CmbClientId_GotFocus()
tmp = CmbClientId.Text
mExchangeTypeID = GetExchangeID("NON NCDEX")
Call TableMst_Client("where ExchangeTypeId = " & mExchangeTypeID)
CmbClientId.Clear
If RsMst_Client.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_Client.RecordCount
    CmbClientId.AddItem RsMst_Client!ClientName
    RsMst_Client.MoveNext
Next
CmbClientId.Text = tmp
End Sub
Private Sub CmbClientId_LostFocus()
If CmbClientId.Text = "" Then
   Exit Sub
End If
RsMst_Client.MoveFirst
RsMst_Client.Find "ClientName = '" & CmbClientId.Text & "'"
If RsMst_Client.EOF Then
   MsgBox "Invalid  selection "
   CmbClientId.SetFocus
   Exit Sub
End If
mClientIDRow = RsMst_Client!ClientIDRow

Call GetGridData(mExchangeTypeID, mCommodityId, mClientIDRow)


End Sub

Private Sub CmbCommodityCondition_LostFocus()
mCommodityCondition = CmbCommodityCondition.ListIndex

If CmbCommodityCondition.Text = "Average" Or CmbCommodityCondition.Text = "Fair" Or _
   CmbCommodityCondition.Text = "Good" Then
Else
   MsgBox "Invalid selection !!!! "
   CmbCommodityCondition.SetFocus
   Exit Sub
End If
     
End Sub

Private Sub CmbWHRCommodity_GotFocus()
tmp = CmbWHRCommodity.Text
Call TableMst_Commodity
CmbWHRCommodity.Clear
If RsMst_Commodity.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_Commodity.RecordCount
    CmbWHRCommodity.AddItem RsMst_Commodity!Commodity
    RsMst_Commodity.MoveNext
Next
CmbWHRCommodity.Text = tmp
End Sub
'----
Private Sub CmbWHRCommodity_LostFocus()
If CmbWHRCommodity.Text = "" Then
   Exit Sub
End If
RsMst_Commodity.MoveFirst
RsMst_Commodity.Find "Commodity = '" & CmbWHRCommodity.Text & "'"
If RsMst_Commodity.EOF Then
   MsgBox "Invalid  selection "
   CmbWHRCommodity.SetFocus
   Exit Sub
End If
mCommodityId = RsMst_Commodity!CommodityID

Call GetGridData(mExchangeTypeID, mCommodityId, mClientIDRow)

End Sub
Private Sub GetGridData(mExchangeTypeID_, mCommodityID_, mClientIDRow_)
Dim mStorageFrom As Date
Dim mWHRDate As Date
Call Grid_Head
If mExchangeTypeID_ = "" Then Exit Sub
If mCommodityID_ = "" Or mCommodityID_ = 0 Then Exit Sub
If mClientIDRow_ = "" Then Exit Sub


mStorageFrom = StorageFrom.Value
mWHRDate = WHRDate.Value
ChkSelection = 0


TxtCMPID = ""
TxtGodownNo = ""
TxtLicenseNo = ""
TxtLicenseDescription = ""
TxtNoOfBags = ""
TxtEleWeight = ""
TxtWHRNoOfBags = ""
TxtWHRQty = ""

If Gen_DBType = "MDB" Then
    tmp = "SELECT Count(Txn_CxTF_Details.CxTFNo) AS CountOfCxTFNo"
    tmp = tmp & " From Txn_CxTF_Details"
    tmp = tmp & " WHERE (Txn_CxTF_Details.CxTFDate Between #" & _
    Format(mStorageFrom, Gen_DatFormat) & "# And #" & Format(mWHRDate, Gen_DatFormat) & "#) AND (Txn_CxTF_Details.Flag Is Null OR Txn_CxTF_Details.Flag='') "
    tmp = tmp & " And (Txn_CxTF_Details.ExchangeTypeID = " & mExchangeTypeID_ & " And Txn_CxTF_Details.CommodityID = " & mCommodityID_ & ")"
    tmp = tmp & " And (Txn_CxTF_Details.TxnType = 'D')"
Else
    tmp = "SELECT Count(Txn_CxTF_Details.CxTFNo) AS CountOfCxTFNo"
    tmp = tmp & " From Txn_CxTF_Details"
    tmp = tmp & " WHERE (Txn_CxTF_Details.CxTFDate Between '" & _
    Format(mStorageFrom, Gen_DatFormat) & "' And '" & Format(mWHRDate, Gen_DatFormat) & "') AND (Txn_CxTF_Details.Flag Is Null OR Txn_CxTF_Details.Flag='') "
    tmp = tmp & " And (Txn_CxTF_Details.ExchangeTypeID = " & mExchangeTypeID_ & " And Txn_CxTF_Details.CommodityID = " & mCommodityID_ & ")"
    tmp = tmp & " And (Txn_CxTF_Details.TxnType = 'D')"
End If
Call TmpTable

If Not IsNull(TmpRs.Fields(0)) Then
   If TmpRs.Fields(0) > 0 Then
      MsgBox " WHR can't be generated since all related CDTFs are not printed !!!!! "
      Exit Sub
   End If
   
End If


If Gen_DBType = "MDB" Then
'    "((Mst_GSL.CDTFDate) Between #" & Format(mStorageFrom, Gen_DatFormat) & "# And #" & Format(mWHRDate, Gen_DatFormat) & "#)) " & _

    tmp = "SELECT Sum((Mst_GSL.DepositeBags)) AS BalanceBags, Sum((Mst_GSL.DepositeWeight)) AS BalanceWeight, " & _
    "Mst_License.LicenseNo, Mst_License.LicenseDesc, Mst_CMP.CMPName, " & _
    "Mst_GSL.CommodityID,  Mst_GSL.CMPID, Mst_GSL.ClientIDRow,  Mst_License.LicenseID,Mst_Godown.GodownNo, " & _
    "Mst_Godown.GodownID,Mst_Godown.AgreementEndDate " & _
    "FROM Mst_CMP INNER JOIN (Mst_License INNER JOIN (Mst_Godown INNER JOIN Mst_GSL ON Mst_Godown.GodownID = Mst_GSL.GodownID) ON Mst_License.LicenseID = Mst_Godown.LicenseID) ON Mst_CMP.CMPID = Mst_GSL.CMPID " & _
    "WHERE (((Mst_GSL.ClientIDRow)=" & mClientIDRow & ") AND ((Mst_GSL.ExchangeTypeID)=" & mExchangeTypeID_ & ") AND ((Mst_GSL.CommodityID)=" & mCommodityID_ & ") AND " & _
    "((Mst_GSL.CDTFDate) <=  #" & Format(mWHRDate, Gen_DatFormat) & "#)) " & _
    "GROUP BY Mst_License.LicenseNo, Mst_License.LicenseDesc, Mst_CMP.CMPName , " & _
    "Mst_GSL.CommodityID,  Mst_GSL.CMPID, Mst_GSL.ClientIDRow,  Mst_License.LicenseID,Mst_Godown.GodownNo,Mst_Godown.GodownID,Mst_Godown.AgreementEndDate "
Else
   ' "((Mst_GSL.CDTFDate) Between '" & Format(mStorageFrom, Gen_DatFormat) & "' And '" & Format(mWHRDate, Gen_DatFormat) & "')) " & _

    tmp = "SELECT Sum((Mst_GSL.DepositeBags)) AS BalanceBags, Sum((Mst_GSL.DepositeWeight)) AS BalanceWeight, " & _
    "Mst_License.LicenseNo, Mst_License.LicenseDesc, Mst_CMP.CMPName, " & _
    "Mst_GSL.CommodityID,  Mst_GSL.CMPID, Mst_GSL.ClientIDRow,  Mst_License.LicenseID,Mst_Godown.GodownNo, " & _
    "Mst_Godown.GodownID,Mst_Godown.AgreementEndDate " & _
    "FROM Mst_CMP INNER JOIN (Mst_License INNER JOIN (Mst_Godown INNER JOIN Mst_GSL ON Mst_Godown.GodownID = Mst_GSL.GodownID) ON Mst_License.LicenseID = Mst_Godown.LicenseID) ON Mst_CMP.CMPID = Mst_GSL.CMPID " & _
    "WHERE (((Mst_GSL.ClientIDRow)=" & mClientIDRow & ") AND ((Mst_GSL.ExchangeTypeID)=" & mExchangeTypeID_ & ") AND ((Mst_GSL.CommodityID)=" & mCommodityID_ & ") AND " & _
    "((Mst_GSL.CDTFDate) <= '" & Format(mWHRDate, Gen_DatFormat) & "')) " & _
    "GROUP BY Mst_License.LicenseNo, Mst_License.LicenseDesc, Mst_CMP.CMPName , " & _
    "Mst_GSL.CommodityID,  Mst_GSL.CMPID, Mst_GSL.ClientIDRow,  Mst_License.LicenseID,Mst_Godown.GodownNo,Mst_Godown.GodownID,Mst_Godown.AgreementEndDate "
End If
Call TmpTable


If TmpRs.RecordCount = 0 Then
  MsgBox "No deposits available for generating WHR !!!!"
  Exit Sub
End If

k = 1
Dim BalBag, BalWt As Variant
For i = 0 To TmpRs.RecordCount - 1
    
    tmp = " SELECT  Sum(Txn_WHR_Detail.WHRNoOfBags) AS WHRNoOfBags,Sum(Txn_WHR_Detail.WHRQuantity) As WHRQuantity "
    tmp = tmp & " From Txn_WHR_Detail "
    tmp = tmp & " where Txn_WHR_Detail.CommodityID = " & TmpRs!CommodityID
    tmp = tmp & " And Txn_WHR_Detail.CMPID =" & TmpRs!CMPID
    tmp = tmp & " And Txn_WHR_Detail.ClientIDRow =" & TmpRs!ClientIDRow
    tmp = tmp & " And Txn_WHR_Detail.LicenseID =" & TmpRs!LicenseID
    tmp = tmp & " And Txn_WHR_Detail.GodownID =" & TmpRs!GodownID
    tmp = tmp & " And Txn_WHR_Detail.WHRDate <= '" & Format(mWHRDate, Gen_DatFormat) & "'"
    Call Tmp1Table
    BalBag = TmpRs!BalanceBags - IIf(IsNull(TmpRs1!WHRNoOfBags), 0, TmpRs1!WHRNoOfBags)
    BalWt = TmpRs!BalanceWeight - IIf(IsNull(TmpRs1!WHRQuantity), 0, TmpRs1!WHRQuantity)
    If BalBag <> 0 And BalWt <> 0 Then
       k = MSHFlexGrid2.Rows - 1
       MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
       MSHFlexGrid2.TextMatrix(k, 1) = TmpRs!CMPName
       MSHFlexGrid2.TextMatrix(k, 2) = TmpRs!GodownNo
       MSHFlexGrid2.TextMatrix(k, 3) = TmpRs!LicenseNo
       MSHFlexGrid2.TextMatrix(k, 4) = TmpRs!LicenseDesc
       MSHFlexGrid2.TextMatrix(k, 5) = BalBag
       MSHFlexGrid2.TextMatrix(k, 6) = BalWt
       MSHFlexGrid2.TextMatrix(k, 7) = ""
       MSHFlexGrid2.TextMatrix(k, 11) = IIf(IsNull(TmpRs!AgreementEndDate), "", TmpRs!AgreementEndDate)
       MSHFlexGrid2.TextMatrix(k, 12) = GetPreFix(IIf(IsNull(TmpRs!CMPID), "", TmpRs!CMPID), "P")

  End If
  TmpRs.MoveNext
Next

End Sub

Private Sub CmdCancel_Click()
TxtHologramNo = ""
Frame3.Visible = False
End Sub

Private Sub CmdInsurance_Click()
FrmMstInsurance.Show
End Sub

Private Sub CmdInsuranceCo_Click()
FrmMstInsuranceCo.Show
End Sub

Private Sub CmdPrint_Click()
If RsTxn_WHRDetail!HologramNo = "" Or IsNull(RsTxn_WHRDetail!HologramNo) Then
   Frame3.Visible = True
   TxtHologramNo.SetFocus
   Call TableMst_Hologram
   Exit Sub
Else
   TxtHologramNo = ""
   Frame3.Visible = False
End If

RsTxn_WHRDetail!Flag = "P"
RsTxn_WHRDetail.Update


Gen_mTimeStamp = GetTimeStamp
Call Create_Xls(Gen_mTimeStamp & Gen_UserName & "_WHR" & Trim(TxtWHRNoNew) & ".xls")


Call Xls_Print_Rpt_New

End Sub
Public Sub Create_Xls(FileName As String)
'Dim oExcel As Object
'Dim oBook As Object
'Dim oSheet As Object
'Dim Pat As String
Pat = App.Path
Set oExcel = CreateObject("Excel.Application")
Set oBook = oExcel.Workbooks.Add
Gen_FLName = FileName
oBook.SaveAs (Pat & "\excel\" & FileName)

Set oSheet = oBook.Worksheets(1)
With oExcel
'.Application.Sheets.Delete
.SheetsInNewWorkbook = 15 'MSHFlexGrid2.Rows - 2
.Workbooks.Add
End With
'oBook.SaveAs (Pat & "\" & FileName)

oExcel.Quit
'oExcel.Close
End Sub


'-----WHR Function-----

Public Sub Xls_Print_Rpt()
'Dim oXL As Excel.Application
'Dim oWB As Excel.Workbook
'Dim oWS As Excel.Worksheet
'Dim mRow As Integer
'Dim mCol As Integer
Dim ans As Variant
MsgBox ("Wait till Next Message")
Set oXL = New Excel.Application
'Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_UserName & "_WHR" & Trim(TxtWHRNoNew) & ".xls")
 
i = 2
Set oWS = oWB.Worksheets("Sheet" & 1)
oWS.PageSetup.RightMargin = 0.05
oWS.Name = TxtWHRNoNew  'MSHFlexGrid2.TextMatrix(i, 23)
oWS.ClearArrows
mRow = 4

oWS.Range("a1:f1").Merge
oWS.Range("a:a").ColumnWidth = 10.86
oWS.Range("b:b").ColumnWidth = 17.71
oWS.Range("c:c").ColumnWidth = 17.14
oWS.Range("d:d").ColumnWidth = 12.86
oWS.Range("e:e").ColumnWidth = 12.86 '32.14
oWS.Range("f:f").ColumnWidth = 17.14

tmp = "a" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = GetCMPName(RsTxn_WHRDetail!CMPID) & "-" & GetCMPAddress(RsTxn_WHRDetail!CMPID, "A")
oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).Font.Size = 7
'oWS.Cells(mRow, c + 1).RowHeight = 24
mRow = mRow + 1

tmp = "a" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = "Warehouse Receipt (WHR)"
oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).Font.Size = 10
oWS.Cells(mRow, c + 1).RowHeight = 24

mRow = mRow + 1

oWS.Cells(mRow, c + 1) = "Warehouse License No"
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).WrapText = True
oWS.Cells(mRow, c + 1).Font.Size = 7
oWS.Cells(mRow, c + 2) = MSHFlexGrid2.TextMatrix(i, 3)
oWS.Cells(mRow, c + 2).HorizontalAlignment = 2
oWS.Cells(mRow, c + 2).Font.Size = 7
oWS.Cells(mRow, c + 3) = "License Expiry Date"
oWS.Cells(mRow, c + 3).Font.Bold = True
oWS.Cells(mRow, c + 3).WrapText = True
oWS.Cells(mRow, c + 3).Font.Size = 7
oWS.Cells(mRow, c + 4) = Format(RsMst_License!LicenseExpiryDate, "dd-mmm-yyyy")
oWS.Cells(mRow, c + 4).HorizontalAlignment = 2
oWS.Cells(mRow, c + 4).Font.Size = 7
'j = 23
'xlsdata (mRow)
'oWS.Range("d5:e5").Merge
oWS.Cells(mRow, c + 5) = "Receipt No"
oWS.Cells(mRow, c + 5).Font.Size = 7
oWS.Cells(mRow, c + 5).Font.Bold = True
'oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 6) = TxtWHRNoNew
oWS.Cells(mRow, c + 6).HorizontalAlignment = 2
oWS.Cells(mRow, c + 6).Font.Size = 7
'xlsdata (mRow)

mRow = mRow + 1

oWS.Cells(mRow, c + 1) = "Location & Address of Warehouse" '& " " & Trim(TxtBankName) & " - " & TxtBranchName
'oWS.Range("b8:e8").Merge
oWS.Cells(mRow, c + 1).WrapText = True
oWS.Cells(mRow, c + 1).Font.Bold = True
tmp = "b" & mRow & ":" & "d" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1).Font.Size = 7
oWS.Cells(mRow, c + 2) = GetLocationFromCMP(RsMst_Godown!CMPID) & " - " & GetGodownName(RsTxn_WHRDetail!GodownID, "A")
oWS.Cells(mRow, c + 2).Font.Size = 7
'oWS.Cells(mRow, c + 2).HorizontalAlignment = -4108
oWS.Cells(mRow, c + 5) = "Godown No"
oWS.Cells(mRow, c + 5).Font.Size = 7
oWS.Cells(mRow, c + 5).Font.Bold = True
'oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 6) = MSHFlexGrid2.TextMatrix(i, 2)
oWS.Cells(mRow, c + 6).HorizontalAlignment = 2
oWS.Cells(mRow, c + 6).Font.Size = 7


mRow = mRow + 1
' tmp = "a" & mRow & ":" & "E" & mRow
' oWS.Range(tmp).Merge
' oWS.Cells(mRow, c + 1).RowHeight = 8
' mRow = mRow + 1

'oWS.Range("a9:e9").Merge
oWS.Cells(mRow, c + 1) = "Name and Address of the Depositor"
'oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).Font.Size = 7
oWS.Cells(mRow, c + 1).WrapText = True
tmp = "b" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
ans = GetClientNameByRow(mClientIDRow)

If ans(4) <> "" Then
   oWS.Cells(mRow, c + 2) = TxtClientID & "-" & Trim(ans(4))
Else
   oWS.Cells(mRow, c + 2) = TxtClientID
End If
oWS.Cells(mRow, c + 2).Font.Size = 7

mRow = mRow + 1

oWS.Cells(mRow, c + 1) = "Commodity"
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).Font.Size = 7
oWS.Cells(mRow, c + 2) = "No of Bags"
oWS.Cells(mRow, c + 2).Font.Bold = True
oWS.Cells(mRow, c + 2).Font.Size = 7
oWS.Cells(mRow, c + 3) = "Net Quantity(MT)"
oWS.Cells(mRow, c + 3).Font.Bold = True
oWS.Cells(mRow, c + 3).Font.Size = 7
oWS.Cells(mRow, c + 4) = "Name and License No of Weigher"
oWS.Cells(mRow, c + 4).Font.Bold = True
oWS.Cells(mRow, c + 4).WrapText = True
oWS.Cells(mRow, c + 4).Font.Size = 7
oWS.Cells(mRow, c + 5) = "Private Marks on Packaging"
oWS.Cells(mRow, c + 5).WrapText = True
oWS.Cells(mRow, c + 5).Font.Bold = True
oWS.Cells(mRow, c + 5).Font.Size = 7
oWS.Cells(mRow, c + 6) = "Condition of Commodity"
oWS.Cells(mRow, c + 6).Font.Bold = True
oWS.Cells(mRow, c + 6).WrapText = True
oWS.Cells(mRow, c + 6).Font.Size = 7
 
mRow = mRow + 1
oWS.Cells(mRow, c + 1) = TxtWHRCommodityId
oWS.Cells(mRow, c + 1).Font.Size = 7
oWS.Cells(mRow, c + 2) = Format(MSHFlexGrid2.TextMatrix(i, 7), "#######0.000") 'MSHFlexGrid2.TextMatrix(i, 7)
oWS.Cells(mRow, c + 2).Font.Size = 7
oWS.Cells(mRow, c + 3) = Format(MSHFlexGrid2.TextMatrix(i, 8), "#######0.000") 'MSHFlexGrid2.TextMatrix(i, 8)
oWS.Cells(mRow, c + 3).Font.Size = 7
oWS.Cells(mRow, c + 4) = TxtWeigher
oWS.Cells(mRow, c + 4).Font.Size = 7
oWS.Cells(mRow, c + 5) = TxtGrade
oWS.Cells(mRow, c + 5).Font.Size = 7
oWS.Cells(mRow, c + 6) = CmbCommodityCondition
oWS.Cells(mRow, c + 6).Font.Size = 7
mRow = mRow + 1

tmp = "a" & mRow & ":" & "d" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = "Name & License No of Assayer"
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).WrapText = True
oWS.Cells(mRow, c + 1).Font.Size = 7
oWS.Cells(mRow, c + 2) = ""
oWS.Cells(mRow, c + 5) = "Quality As per Grader"
oWS.Cells(mRow, c + 5).WrapText = True
oWS.Cells(mRow, c + 5).Font.Size = 7
oWS.Cells(mRow, c + 5).Font.Bold = True
oWS.Cells(mRow, c + 6) = "Grade Valid Till"
oWS.Cells(mRow, c + 6).Font.Size = 7
oWS.Cells(mRow, c + 6).Font.Bold = True
mRow = mRow + 1

tmp = "a" & mRow & ":" & "d" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = TxtAssayer
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).WrapText = True
oWS.Cells(mRow, c + 1).Font.Size = 7
oWS.Cells(mRow, c + 5) = TxtAssayerCerti
oWS.Cells(mRow, c + 5).WrapText = True
oWS.Cells(mRow, c + 5).Font.Size = 7
If TxtAssayer <> "" Then
   oWS.Cells(mRow, c + 6) = Format(GradeValidUpto, "dd-mmm-yyyy")
   oWS.Cells(mRow, c + 6).Font.Size = 7
End If
mRow = mRow + 1

tmp = "a" & mRow & ":" & "b" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = "Storage Charge(Rate)-excluding Taxes(Rupees per bag/month)"
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).WrapText = True
oWS.Cells(mRow, c + 1).Font.Size = 7
oWS.Cells(mRow, c + 1).RowHeight = 24

oWS.Cells(mRow, c + 3) = Format(TxtStorageCharges, "#######0.000") 'TxtStorageCharges
oWS.Cells(mRow, c + 3).Font.Size = 7
oWS.Cells(mRow, c + 3).RowHeight = 24

tmp = "d" & mRow & ":" & "e" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 4) = "Storage Charge - excluding Taxes (Total Amount) per month"
oWS.Cells(mRow, c + 4).Font.Bold = True
oWS.Cells(mRow, c + 4).WrapText = True
oWS.Cells(mRow, c + 4).Font.Size = 7
oWS.Cells(mRow, c + 4).RowHeight = 24

oWS.Cells(mRow, c + 6) = Format(TxtStorageAmount, "#######0.000")  'MSHFlexGrid2.TextMatrix(i, 7) * TxtStorageCharges
oWS.Cells(mRow, c + 6).Font.Size = 7
oWS.Cells(mRow, c + 6).RowHeight = 24
mRow = mRow + 1


tmp = "a" & mRow & ":" & "b" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = "Risk Coverd"
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).WrapText = True
oWS.Cells(mRow, c + 1).Font.Size = 7
'oWS.Cells(mRow, c + 1).RowHeight = 24
tmp = "c" & mRow & ":" & "d" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 3) = "Insurance Co."
oWS.Cells(mRow, c + 3).Font.Bold = True
oWS.Cells(mRow, c + 3).WrapText = True
oWS.Cells(mRow, c + 3).Font.Size = 7
'oWS.Cells(mRow, c + 1).RowHeight = 8
tmp = "e" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 5) = "Amount Insured"
oWS.Cells(mRow, c + 5).WrapText = True
oWS.Cells(mRow, c + 5).Font.Size = 7
oWS.Cells(mRow, c + 5).Font.Bold = True

mRow = mRow + 1

tmp = "a" & mRow & ":" & "b" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = "Fire,Burglary,Spontaneous Combustion,EarthQuakes,Terrorism and allied causes"
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).WrapText = True
oWS.Cells(mRow, c + 1).Font.Size = 7
'oWS.Cells(mRow, c + 1).RowHeight = 24
tmp = "c" & mRow & ":" & "d" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 3) = ""
oWS.Cells(mRow, c + 3).WrapText = True
oWS.Cells(mRow, c + 3).Font.Size = 7
'oWS.Cells(mRow, c + 1).RowHeight = 8
tmp = "e" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 5) = ""
oWS.Cells(mRow, c + 5).WrapText = True
oWS.Cells(mRow, c + 5).Font.Size = 7
oWS.Cells(mRow, c + 5).RowHeight = 30

mRow = mRow + 1


tmp = "a" & mRow & ":" & "b" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = "Goods are Accepted for storage From"
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).WrapText = True
oWS.Cells(mRow, c + 1).Font.Size = 7

oWS.Cells(mRow, c + 3) = Format(StorageFrom, "dd-mmm-yyyy")
oWS.Cells(mRow, c + 3).Font.Size = 7

tmp = "d" & mRow & ":" & "e" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 4) = "To"
oWS.Cells(mRow, c + 4).Font.Bold = True
oWS.Cells(mRow, c + 4).WrapText = True
oWS.Cells(mRow, c + 4).Font.Size = 7

oWS.Cells(mRow, c + 6) = Format(StorageTo, "dd-mmm-yyyy")
oWS.Cells(mRow, c + 6).Font.Size = 7

mRow = mRow + 1

tmp = "a" & mRow & ":" & "b" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = "Market Rate at the time of deposit(Rs/MT)"
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).WrapText = True
oWS.Cells(mRow, c + 1).Font.Size = 7
oWS.Cells(mRow, c + 1).RowHeight = 10

oWS.Cells(mRow, c + 3) = Format(TxtRate, "#######0.000") 'TxtRate
oWS.Cells(mRow, c + 3).Font.Size = 7

tmp = "d" & mRow & ":" & "e" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 4) = "Value of Commodities at time of deposit"
oWS.Cells(mRow, c + 4).Font.Bold = True
oWS.Cells(mRow, c + 4).WrapText = True
oWS.Cells(mRow, c + 4).Font.Size = 7
oWS.Cells(mRow, c + 4).RowHeight = 20

oWS.Cells(mRow, c + 6) = Format(TxtValueofCommodity, "#######0.000") 'TxtValueofCommodity
oWS.Cells(mRow, c + 6).Font.Size = 7

mRow = mRow + 1

oWS.Cells(mRow, c + 1) = "Date"
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).WrapText = True
oWS.Cells(mRow, c + 1).Font.Size = 7

oWS.Cells(mRow, c + 2) = Format(WHRDate, "dd-mmm-yyyy")
oWS.Cells(mRow, c + 2).Font.Size = 7

tmp = "c" & mRow & ":" & "d" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 3) = "Signature & Stamp of Warehouse Manager"
oWS.Cells(mRow, c + 3).Font.Bold = True
oWS.Cells(mRow, c + 3).WrapText = True
oWS.Cells(mRow, c + 3).Font.Size = 7
oWS.Cells(mRow, c + 4).RowHeight = 30

tmp = "e" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 5) = ""
oWS.Cells(mRow, c + 5).Font.Size = 7

mRow = mRow + 1

tmp = "a" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1).RowHeight = 8

mRow = mRow + 1

oWS.Cells(mRow, c + 1) = "Date"
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).WrapText = True
oWS.Cells(mRow, c + 1).Font.Size = 7

tmp = "b" & mRow & ":" & "c" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 2) = "Name and Address Of Lender"
oWS.Cells(mRow, c + 2).Font.Bold = True
oWS.Cells(mRow, c + 2).WrapText = True
oWS.Cells(mRow, c + 2).Font.Size = 7

oWS.Cells(mRow, c + 4) = "Pledge Number"
oWS.Cells(mRow, c + 4).Font.Bold = True
oWS.Cells(mRow, c + 4).WrapText = True
oWS.Cells(mRow, c + 4).Font.Size = 7

oWS.Cells(mRow, c + 5) = "Pledge Released on"
oWS.Cells(mRow, c + 5).Font.Bold = True
oWS.Cells(mRow, c + 5).WrapText = True
oWS.Cells(mRow, c + 5).Font.Size = 7

oWS.Cells(mRow, c + 6) = "Sign of Lender/Warehouse Manager"
oWS.Cells(mRow, c + 6).Font.Bold = True
oWS.Cells(mRow, c + 6).WrapText = True
oWS.Cells(mRow, c + 6).Font.Size = 7

mRow = mRow + 1
mRow = mRow + 1
mRow = mRow + 1
mRow = mRow + 1
mRow = mRow + 1

oWS.Cells(mRow, c + 1) = "Withdrawal History"
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).WrapText = True
oWS.Cells(mRow, c + 1).Font.Size = 7
oWS.Cells(mRow, c + 1).RowHeight = 24

tmp = "b" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 2) = "The good mentioned below are hereby released for delivery from the" & _
                         " Warehouse. The unreleased balance is subject to lien for" & _
                         " unpaid charges and borrowings from lending institutions."
oWS.Cells(mRow, c + 2).Font.Bold = True
oWS.Cells(mRow, c + 2).WrapText = True
oWS.Cells(mRow, c + 2).Font.Size = 6

mRow = mRow + 1

tmp = "a" & mRow & ":" & "a" & (mRow + 1)
oWS.Range(tmp).Merge


oWS.Cells(mRow, c + 1) = "Date"
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).WrapText = True
oWS.Cells(mRow, c + 1).Font.Size = 7

tmp = "b" & mRow & ":" & "c" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 2) = "Quantity Released"
oWS.Cells(mRow, c + 2).Font.Bold = True
oWS.Cells(mRow, c + 2).WrapText = True
oWS.Cells(mRow, c + 2).Font.Size = 7

tmp = "d" & mRow & ":" & "e" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 4) = "Quantity Balance"
oWS.Cells(mRow, c + 4).Font.Bold = True
oWS.Cells(mRow, c + 4).WrapText = True
oWS.Cells(mRow, c + 4).Font.Size = 7

tmp = "f" & mRow & ":" & "f" & (mRow + 1)
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 6) = "Sign & Stamp of the Warehouse Manager"
oWS.Cells(mRow, c + 6).Font.Bold = True
oWS.Cells(mRow, c + 6).WrapText = True
oWS.Cells(mRow, c + 6).Font.Size = 7


mRow = mRow + 1

oWS.Cells(mRow, c + 2) = "No. of Bags"
oWS.Cells(mRow, c + 2).Font.Size = 7
oWS.Cells(mRow, c + 2).Font.Bold = True

oWS.Cells(mRow, c + 3) = "Weight in MT"
oWS.Cells(mRow, c + 3).Font.Size = 7
oWS.Cells(mRow, c + 3).Font.Bold = True

oWS.Cells(mRow, c + 4) = "No. Of Bags"
oWS.Cells(mRow, c + 4).Font.Size = 7
oWS.Cells(mRow, c + 4).Font.Bold = True

oWS.Cells(mRow, c + 5) = "Weight in  MT"
oWS.Cells(mRow, c + 5).Font.Size = 7
oWS.Cells(mRow, c + 5).Font.Bold = True

mRow = mRow + 1
mRow = mRow + 1
mRow = mRow + 1
mRow = mRow + 1
mRow = mRow + 1

mRow = mRow + 1
tmp = "a" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1).RowHeight = 10

mRow = mRow + 1
tmp = "d" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
oWS.Cells(mRow, c + 4) = "HOLOGRAM NO. :- " & "" & RsTxn_WHRDetail!HologramNo
oWS.Cells(mRow, c + 4).Font.Bold = True
oWS.Cells(mRow, c + 4).Font.Size = 7
oWS.Cells(mRow, c + 4).RowHeight = 30



mRow = mRow + 1
tmp = "a" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1).RowHeight = 10

mRow = mRow + 1

tmp = "a" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = "The Warehouse Receipt is valid for transfer by endorsement, and/or for " & _
                         " marking of plegde, subject to confirmation by 'Warehouse Manager' "
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).WrapText = True
oWS.Cells(mRow, c + 1).Font.Size = 6
oWS.Cells(mRow, c + 1).RowHeight = 18

k = mRow

For Cell = 4 To k
    tmp = "a" & Cell & ":" & "f" & Cell
    oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
    oWS.Range(tmp).Font.Name = "Zurich BT"
    'oWS.Range(tmp).Font.Size = 7
    oWS.Range(tmp).WrapText = True
Next Cell
    
oWB.Save
'oWB.Close
'oXL.Quit
MsgBox ("Excel File Created !!! ")
oXL.Visible = True

End Sub
Public Sub Xls_Print_Rpt_New()

Dim ans As Variant
MsgBox ("Wait till Next Message")
Set oXL = New Excel.Application
'Dim Pat As String
Pat = App.Path
'Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_UserName & "_WHR" & Trim(TxtWHRNoNew) & ".xls")
 

Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "_WHR" & Trim(TxtWHRNoNew) & ".xls")
 

i = 2
Set oWS = oWB.Worksheets("Sheet" & 1)
oWS.PageSetup.RightMargin = 0.05
oWS.Name = TxtWHRNoNew  'MSHFlexGrid2.TextMatrix(i, 23)
oWS.ClearArrows
mRow = 4

oWS.Range("a1:f1").Merge
oWS.Range("a:a").ColumnWidth = 10.86
oWS.Range("b:b").ColumnWidth = 17.71
oWS.Range("c:c").ColumnWidth = 17.14
oWS.Range("d:d").ColumnWidth = 12.86
oWS.Range("e:e").ColumnWidth = 12.86 '32.14
oWS.Range("f:f").ColumnWidth = 17.14

tmp = "a" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = GetCMPAddress(RsTxn_WHRDetail!CMPID, "D") '"NCMSL" & "-" & GetCMPName(RsTxn_WHRDetail!CMPID)
oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).Font.Size = 10
'oWS.Cells(mRow, c + 1).RowHeight = 24
mRow = mRow + 1

tmp = "a" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = "Warehouse Receipt"
oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).Font.Size = 10
oWS.Cells(mRow, c + 1).RowHeight = 24

mRow = mRow + 1


tmp = "a" & mRow & ":" & "b" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = "Receipt No"
oWS.Cells(mRow, c + 1).Font.Size = 7
oWS.Cells(mRow, c + 1).Font.Bold = True
'oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 3) = TxtWHRNoNew
oWS.Cells(mRow, c + 3).HorizontalAlignment = 2
oWS.Cells(mRow, c + 3).Font.Size = 7

tmp = "d" & mRow & ":" & "e" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 4) = "Receipt Date"
oWS.Cells(mRow, c + 4).Font.Size = 7
oWS.Cells(mRow, c + 4).Font.Bold = True

oWS.Cells(mRow, c + 6).HorizontalAlignment = 2
oWS.Cells(mRow, c + 6) = Format(WHRDate, "dd-mmm-yyyy")
oWS.Cells(mRow, c + 6).Font.Size = 7

mRow = mRow + 1

tmp = "a" & mRow & ":" & "b" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = "Warehouse License No"
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).WrapText = True
oWS.Cells(mRow, c + 1).Font.Size = 7
oWS.Cells(mRow, c + 3) = MSHFlexGrid2.TextMatrix(i, 3)
oWS.Cells(mRow, c + 3).HorizontalAlignment = 2
oWS.Cells(mRow, c + 3).Font.Size = 7

tmp = "d" & mRow & ":" & "e" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 4) = "License Expiry Date"
oWS.Cells(mRow, c + 4).Font.Bold = True
oWS.Cells(mRow, c + 4).WrapText = True
oWS.Cells(mRow, c + 4).Font.Size = 7

oWS.Cells(mRow, c + 6) = Format(RsMst_License!LicenseExpiryDate, "dd-mmm-yyyy")
oWS.Cells(mRow, c + 6).HorizontalAlignment = 2
oWS.Cells(mRow, c + 6).Font.Size = 7

mRow = mRow + 1

tmp = "a" & mRow & ":" & "b" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = "Location & Address of Warehouse" '& " " & Trim(TxtBankName) & " - " & TxtBranchName
oWS.Cells(mRow, c + 1).WrapText = True
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).Font.Size = 7
tmp = "c" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge

oWS.Cells(mRow, c + 3) = GetLocationFromCMP(RsMst_Godown!CMPID) '& " - " & GetGodownName(RsTxn_WHRDetail!GodownID, "A")
oWS.Cells(mRow, c + 3).Font.Size = 7
'oWS.Cells(mRow, c + 2).HorizontalAlignment = -4108

mRow = mRow + 1

tmp = "a" & mRow & ":" & "b" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = "Godown No" '& " " & Trim(TxtBankName) & " - " & TxtBranchName
'oWS.Range("b8:e8").Merge
oWS.Cells(mRow, c + 1).WrapText = True
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).Font.Size = 7
tmp = "c" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 3) = GetGodownName(RsTxn_WHRDetail!GodownID, "N")
oWS.Cells(mRow, c + 3).Font.Size = 7
oWS.Cells(mRow, c + 3).HorizontalAlignment = 2

mRow = mRow + 1

tmp = "a" & mRow & ":" & "b" & mRow
oWS.Range(tmp).Merge

oWS.Cells(mRow, c + 1) = "Godown Address"
oWS.Cells(mRow, c + 1).Font.Size = 7
oWS.Cells(mRow, c + 1).Font.Bold = True
'oWS.Range(tmp).Merge
tmp = "c" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge

oWS.Cells(mRow, c + 3) = GetGodownName(RsTxn_WHRDetail!GodownID, "A")  'MSHFlexGrid2.TextMatrix(i, 2)
oWS.Cells(mRow, c + 3).HorizontalAlignment = 2
oWS.Cells(mRow, c + 3).Font.Size = 7


mRow = mRow + 1

tmp = "a" & mRow & ":" & "b" & mRow
oWS.Range(tmp).Merge

oWS.Cells(mRow, c + 1) = "Name and Address of the Depositor"
'oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).Font.Size = 7
oWS.Cells(mRow, c + 1).WrapText = True
tmp = "c" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
ans = GetClientNameByRow(mClientIDRow)

If ans(4) <> "" Then
   oWS.Cells(mRow, c + 3) = TxtClientID & "-" & Trim(ans(4))
Else
   oWS.Cells(mRow, c + 3) = TxtClientID
End If
oWS.Cells(mRow, c + 3).Font.Size = 7

mRow = mRow + 1

tmp = "a" & mRow & ":" & "b" & mRow
oWS.Range(tmp).Merge

oWS.Cells(mRow, c + 1) = "Commodity"
If ChkSaidToContain.Value = 1 Then
   oWS.Cells(mRow, c + 1) = "Commodity (Said to contain)"
End If
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).Font.Size = 7

oWS.Cells(mRow, c + 3) = TxtWHRCommodityId
oWS.Cells(mRow, c + 3).Font.Size = 7

tmp = "d" & mRow & ":" & "e" & mRow
oWS.Range(tmp).Merge

oWS.Cells(mRow, c + 4) = "Condition of Commodity"
oWS.Cells(mRow, c + 4).Font.Bold = True
oWS.Cells(mRow, c + 4).WrapText = True
oWS.Cells(mRow, c + 4).Font.Size = 7
oWS.Cells(mRow, c + 6) = CmbCommodityCondition
oWS.Cells(mRow, c + 6).Font.Size = 7

mRow = mRow + 1

tmp = "a" & mRow & ":" & "b" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = "No of Bags"
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).Font.Size = 7
oWS.Cells(mRow, c + 3) = Format(MSHFlexGrid2.TextMatrix(i, 7), "#######0.000") 'MSHFlexGrid2.TextMatrix(i, 7)
oWS.Cells(mRow, c + 3).Font.Size = 7

tmp = "d" & mRow & ":" & "e" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 4) = "Net Quantity(MT)"
If ChkEstimatedWgt.Value = 1 Then
   oWS.Cells(mRow, c + 4) = "Net Quantity(MT) (Estimated Weight)"
End If
oWS.Cells(mRow, c + 4).Font.Bold = True
oWS.Cells(mRow, c + 4).Font.Size = 7
oWS.Cells(mRow, c + 4).RowHeight = 24

oWS.Cells(mRow, c + 6) = Format(MSHFlexGrid2.TextMatrix(i, 8), "#######0.000") 'MSHFlexGrid2.TextMatrix(i, 8)
oWS.Cells(mRow, c + 6).Font.Size = 7


mRow = mRow + 1

tmp = "a" & mRow & ":" & "b" & mRow
oWS.Range(tmp).Merge

oWS.Cells(mRow, c + 1) = "Name and License No of Weigher"
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).WrapText = True
oWS.Cells(mRow, c + 1).Font.Size = 7
oWS.Cells(mRow, c + 3) = TxtWeigher
oWS.Cells(mRow, c + 3).Font.Size = 7

tmp = "d" & mRow & ":" & "e" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 4) = "Private Marks on Packaging"
oWS.Cells(mRow, c + 4).WrapText = True
oWS.Cells(mRow, c + 4).Font.Bold = True
oWS.Cells(mRow, c + 4).Font.Size = 7
oWS.Cells(mRow, c + 6) = TxtGrade
oWS.Cells(mRow, c + 6).Font.Size = 7

mRow = mRow + 1

tmp = "a" & mRow & ":" & "b" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = "Name & License No of Assayer"
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).WrapText = True
oWS.Cells(mRow, c + 1).Font.Size = 7
tmp = "c" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 3) = TxtAssayer
oWS.Cells(mRow, c + 3).Font.Size = 7

mRow = mRow + 1

tmp = "a" & mRow & ":" & "b" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = "Quality As per Grader"
oWS.Cells(mRow, c + 1).WrapText = True
oWS.Cells(mRow, c + 1).Font.Size = 7
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 3) = TxtAssayerCerti
oWS.Cells(mRow, c + 3).WrapText = True
oWS.Cells(mRow, c + 3).Font.Size = 7

tmp = "d" & mRow & ":" & "e" & mRow
oWS.Range(tmp).Merge

oWS.Cells(mRow, c + 4) = "Grade Valid Till"
oWS.Cells(mRow, c + 4).Font.Size = 7
oWS.Cells(mRow, c + 4).Font.Bold = True

If TxtAssayer <> "" Then
   oWS.Cells(mRow, c + 6) = Format(GradeValidUpto, "dd-mmm-yyyy")
   oWS.Cells(mRow, c + 6).Font.Size = 7
End If


mRow = mRow + 1

tmp = "a" & mRow & ":" & "b" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = "Storage Charge(Rate)-excluding Taxes(Rupees per bag/month)"
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).WrapText = True
oWS.Cells(mRow, c + 1).Font.Size = 7
oWS.Cells(mRow, c + 1).RowHeight = 24

oWS.Cells(mRow, c + 3) = Format(TxtStorageCharges, "#######0.000") 'TxtStorageCharges
oWS.Cells(mRow, c + 3).Font.Size = 7
oWS.Cells(mRow, c + 3).RowHeight = 24

tmp = "d" & mRow & ":" & "e" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 4) = "Storage Charge - excluding Taxes (Total Amount) per month"
oWS.Cells(mRow, c + 4).Font.Bold = True
oWS.Cells(mRow, c + 4).WrapText = True
oWS.Cells(mRow, c + 4).Font.Size = 7
oWS.Cells(mRow, c + 4).RowHeight = 24

oWS.Cells(mRow, c + 6) = Format(TxtStorageAmount, "#######0.000")  'MSHFlexGrid2.TextMatrix(i, 7) * TxtStorageCharges
oWS.Cells(mRow, c + 6).Font.Size = 7
oWS.Cells(mRow, c + 6).RowHeight = 24
mRow = mRow + 1


tmp = "a" & mRow & ":" & "b" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = "Goods are Accepted for storage From"
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).WrapText = True
oWS.Cells(mRow, c + 1).Font.Size = 7

oWS.Cells(mRow, c + 3) = Format(StorageFrom, "dd-mmm-yyyy")
oWS.Cells(mRow, c + 3).Font.Size = 7

tmp = "d" & mRow & ":" & "e" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 4) = "To"
oWS.Cells(mRow, c + 4).Font.Bold = True
oWS.Cells(mRow, c + 4).WrapText = True
oWS.Cells(mRow, c + 4).Font.Size = 7

oWS.Cells(mRow, c + 6) = Format(StorageTo, "dd-mmm-yyyy")
oWS.Cells(mRow, c + 6).Font.Size = 7

mRow = mRow + 1

tmp = "a" & mRow & ":" & "b" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = "Market Rate at the time of deposit(Rs/MT)"
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).WrapText = True
oWS.Cells(mRow, c + 1).Font.Size = 7
oWS.Cells(mRow, c + 1).RowHeight = 20

oWS.Cells(mRow, c + 3) = Format(TxtRate, "#######0.000") 'TxtRate
oWS.Cells(mRow, c + 3).Font.Size = 7

tmp = "d" & mRow & ":" & "e" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 4) = "Value of Commodities at time of deposit (Amount in Rs)"
oWS.Cells(mRow, c + 4).Font.Bold = True
oWS.Cells(mRow, c + 4).WrapText = True
oWS.Cells(mRow, c + 4).Font.Size = 7
oWS.Cells(mRow, c + 4).RowHeight = 20

oWS.Cells(mRow, c + 6) = Format(TxtValueofCommodity, "#######0.000") 'TxtValueofCommodity
oWS.Cells(mRow, c + 6).Font.Size = 7

mRow = mRow + 1


tmp = "a" & mRow & ":" & "b" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = "Signature & Stamp of Warehouse Manager"
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).WrapText = True
oWS.Cells(mRow, c + 1).Font.Size = 7
oWS.Cells(mRow, c + 1).RowHeight = 50


tmp = "d" & mRow & ":" & "e" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 4) = "Hologram No:- " & RsTxn_WHRDetail!HologramNo
oWS.Cells(mRow, c + 4).Font.Size = 7
oWS.Cells(mRow, c + 4).Font.Bold = True
oWS.Cells(mRow, c + 6) = "Affix Hologram Here"
oWS.Cells(mRow, c + 6).Font.Bold = True
oWS.Cells(mRow, c + 6).Font.Size = 7
oWS.Cells(mRow, c + 6).RowHeight = 50


mRow = mRow + 1
tmp = "a" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1).RowHeight = 8

mRow = mRow + 1


tmp = "a" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = "Insurance Details"
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).Font.Size = 7
mRow = mRow + 1

oWS.Cells(mRow, c + 1) = "Valid From - Valid to"
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).Font.Size = 7
oWS.Cells(mRow, c + 1).RowHeight = 20
oWS.Cells(mRow, c + 2) = "Policy & Cover Note No"
oWS.Cells(mRow, c + 2).Font.Bold = True
oWS.Cells(mRow, c + 2).Font.Size = 7
oWS.Cells(mRow, c + 3) = "Insurance Co."
oWS.Cells(mRow, c + 3).Font.Bold = True
oWS.Cells(mRow, c + 3).Font.Size = 7
oWS.Cells(mRow, c + 4) = "Policy Amount (in Rs.)"
oWS.Cells(mRow, c + 4).Font.Bold = True
oWS.Cells(mRow, c + 4).Font.Size = 7
tmp = "e" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 5) = "Risk Covered"
oWS.Cells(mRow, c + 5).Font.Bold = True
oWS.Cells(mRow, c + 5).Font.Size = 7

mRow = mRow + 1
If ValidFrom.Value <> "1/1/1900" Or ValidTo <> "1/1/1900" Then
   oWS.Cells(mRow, c + 1) = Format(ValidFrom, "DD-MMM-YYYY") & " To " & Format(ValidTo, "DD-MMM-YYYY")   '"Valid From - Valid To"
   
Else
   oWS.Cells(mRow, c + 1) = ""
End If
oWS.Cells(mRow, c + 1).RowHeight = 20
oWS.Cells(mRow, c + 1).WrapText = True
oWS.Cells(mRow, c + 1).Font.Size = 7
oWS.Cells(mRow, c + 2) = TxtPolicyNo
oWS.Cells(mRow, c + 2).WrapText = True
oWS.Cells(mRow, c + 2).Font.Size = 7
oWS.Cells(mRow, c + 3) = TxtInsuranceCo
oWS.Cells(mRow, c + 3).WrapText = True
oWS.Cells(mRow, c + 3).Font.Size = 7
If TxtAmount <> "0" Then
   oWS.Cells(mRow, c + 4) = TxtAmount
Else
   oWS.Cells(mRow, c + 4) = ""
End If
oWS.Cells(mRow, c + 4).WrapText = True
oWS.Cells(mRow, c + 4).Font.Size = 7
tmp = "e" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 5) = Trim(TxtRisks)
oWS.Cells(mRow, c + 5).WrapText = True
oWS.Cells(mRow, c + 5).Font.Size = 7

mRow = mRow + 1

tmp = "a" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = "Pledge Details"
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).Font.Size = 7
mRow = mRow + 1

oWS.Cells(mRow, c + 1) = "Date"
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).WrapText = True
oWS.Cells(mRow, c + 1).Font.Size = 7

tmp = "b" & mRow & ":" & "c" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 2) = "Name and Address Of Lender"
oWS.Cells(mRow, c + 2).Font.Bold = True
oWS.Cells(mRow, c + 2).WrapText = True
oWS.Cells(mRow, c + 2).Font.Size = 7

oWS.Cells(mRow, c + 4) = "Pledge Number"
oWS.Cells(mRow, c + 4).Font.Bold = True
oWS.Cells(mRow, c + 4).WrapText = True
oWS.Cells(mRow, c + 4).Font.Size = 7

oWS.Cells(mRow, c + 5) = "Pledge Released on"
oWS.Cells(mRow, c + 5).Font.Bold = True
oWS.Cells(mRow, c + 5).WrapText = True
oWS.Cells(mRow, c + 5).Font.Size = 7

oWS.Cells(mRow, c + 6) = "Sign of Lender/Warehouse Manager"
oWS.Cells(mRow, c + 6).Font.Bold = True
oWS.Cells(mRow, c + 6).WrapText = True
oWS.Cells(mRow, c + 6).Font.Size = 7

mRow = mRow + 1
mRow = mRow + 1
mRow = mRow + 1
mRow = mRow + 1
mRow = mRow + 1

oWS.Cells(mRow, c + 1) = "Withdrawal History"
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).WrapText = True
oWS.Cells(mRow, c + 1).Font.Size = 7
oWS.Cells(mRow, c + 1).RowHeight = 24

tmp = "b" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 2) = "The good mentioned below are hereby released for delivery from the" & _
                         " Warehouse. The unreleased balance is subject to lien for" & _
                         " unpaid charges and borrowings from lending institutions."
oWS.Cells(mRow, c + 2).Font.Bold = True
oWS.Cells(mRow, c + 2).WrapText = True
oWS.Cells(mRow, c + 2).Font.Size = 6

mRow = mRow + 1

tmp = "a" & mRow & ":" & "a" & (mRow + 1)
oWS.Range(tmp).Merge


oWS.Cells(mRow, c + 1) = "Date"
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).WrapText = True
oWS.Cells(mRow, c + 1).Font.Size = 7

tmp = "b" & mRow & ":" & "c" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 2) = "Quantity Released"
oWS.Cells(mRow, c + 2).Font.Bold = True
oWS.Cells(mRow, c + 2).WrapText = True
oWS.Cells(mRow, c + 2).Font.Size = 7

tmp = "d" & mRow & ":" & "e" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 4) = "Quantity Balance"
oWS.Cells(mRow, c + 4).Font.Bold = True
oWS.Cells(mRow, c + 4).WrapText = True
oWS.Cells(mRow, c + 4).Font.Size = 7

tmp = "f" & mRow & ":" & "f" & (mRow + 1)
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 6) = "Sign & Stamp of the Warehouse Manager"
oWS.Cells(mRow, c + 6).Font.Bold = True
oWS.Cells(mRow, c + 6).WrapText = True
oWS.Cells(mRow, c + 6).Font.Size = 7


mRow = mRow + 1

oWS.Cells(mRow, c + 2) = "No. of Bags"
oWS.Cells(mRow, c + 2).Font.Size = 7
oWS.Cells(mRow, c + 2).Font.Bold = True

oWS.Cells(mRow, c + 3) = "Weight in MT"
oWS.Cells(mRow, c + 3).Font.Size = 7
oWS.Cells(mRow, c + 3).Font.Bold = True

oWS.Cells(mRow, c + 4) = "No. Of Bags"
oWS.Cells(mRow, c + 4).Font.Size = 7
oWS.Cells(mRow, c + 4).Font.Bold = True

oWS.Cells(mRow, c + 5) = "Weight in  MT"
oWS.Cells(mRow, c + 5).Font.Size = 7
oWS.Cells(mRow, c + 5).Font.Bold = True

mRow = mRow + 1
mRow = mRow + 1
'mRow = mRow + 1
'mRow = mRow + 1
'mRow = mRow + 1

mRow = mRow + 1

tmp = "a" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1).RowHeight = 8

mRow = mRow + 1

tmp = "a" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = "The Warehouse Receipt is valid for transfer by endorsement, and/or for " & _
                         " marking of plegde, subject to confirmation by Warehouse Manager "
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).WrapText = True
oWS.Cells(mRow, c + 1).Font.Size = 6
oWS.Cells(mRow, c + 1).RowHeight = 18

k = mRow

For Cell = 4 To k
    tmp = "a" & Cell & ":" & "f" & Cell
    oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
    oWS.Range(tmp).Font.Name = "Zurich BT"
    'oWS.Range(tmp).Font.Size = 7
    oWS.Range(tmp).WrapText = True
Next Cell
    
    
If Gen_ExcelPassword <> "" Then
   oWS.Protect Gen_ExcelPassword, , , , , , True, True
End If
oWB.Save
'oWB.Close
'oXL.Quit

mPrefix = GetPreFix(RsTxn_WHRDetail!CMPID, "P")
Call Log_Print_History(mPrefix, "WHR", TxtWHRNo, Gen_FLName)
    
    
MsgBox ("Excel File Created !!! ")
oXL.Visible = True






End Sub


Private Sub CmdUpdate_Click()
If TxtStorageCharges = "" Then
  MsgBox "Enter Storage Charges Rate !!!"
  TxtStorageCharges.SetFocus
  Exit Sub
End If
If TxtRate = "" Then
  MsgBox "Enter Commodity Price !!!"
  TxtRate.SetFocus
  Exit Sub
End If
If TxtWHRNoOfBags = "" Then
  MsgBox "Please Provide WHR No Of Bags"
  TxtWHRNoOfBags.SetFocus
  Exit Sub
End If
If TxtWHRQty = "" Then
  MsgBox "Please Provide WHR Quantity!!!"
  TxtWHRQty.SetFocus
  Exit Sub
End If
If Val(TxtWHRQty) <> 0 And Val(TxtWHRQty) <> TxtWHRNoOfBags Then
   If Val(TxtNoOfBags) <> Val(TxtWHRNoOfBags) Or Val(TxtEleWeight) <> Val(TxtWHRQty) Then
      tmp = Validate_ComBags(mCommodityId, Val(TxtWHRNoOfBags), Val(TxtWHRQty))
      If Len(tmp) > 1 Then
         MsgBox tmp
         TxtWHRQty.SetFocus
         Exit Sub
      End If
   End If
End If
If Val(TxtWHRNoOfBags.Text) = 0 And Val(TxtWHRQty) = 0 Then
    ChkSelection = 0
    MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 0) = ""
Else
    MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 0) = "*"
    ChkSelection = 1
End If

MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 1) = TxtCMPID
mCMPID = GetCMPID(TxtCMPID)
Dim m As Double
m = mCMPID
MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 2) = TxtGodownNo
mGodownID = GetGodownID(TxtGodownNo, m)

MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 3) = TxtLicenseNo
mLicenseNo = TxtLicenseNo
MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 4) = TxtLicenseDescription
MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 5) = Val(TxtNoOfBags) - Val(TxtWHRNoOfBags)
MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 6) = Val(TxtEleWeight) - Val(TxtWHRQty)
MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 7) = Val(TxtWHRNoOfBags)
MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 8) = Val(TxtWHRQty)

MSHFlexGrid2.TextMatrix(0, 7) = Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 7))
MSHFlexGrid2.TextMatrix(0, 8) = Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 8))

TxtStorageAmount = Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 7)) * Val(TxtStorageCharges)
TxtValueofCommodity = Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 8)) * Val(TxtRate)
TxtAgreeDate = MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 11)
TxtCMPID = ""
TxtGodownNo = ""
TxtLicenseNo = ""
TxtLicenseDescription = ""
TxtNoOfBags = ""
TxtEleWeight = ""
TxtWHRNoOfBags = ""
TxtWHRQty = ""
Frame2.Enabled = False
CmdUpdate.Enabled = False
SSTab1.Tab = 1
End Sub

Private Sub DeleteCmd_Click()

If RsTxn_WHRDetail!Flag = "P" Then
   MsgBox "WHR Print is generated You Can't Delete this WHR !!!! "
   Exit Sub
End If
If Val(MSHFlexGrid2.TextMatrix(2, 10)) > 0 Then
   MsgBox "You Can't Delete this WHR Since, CITF is generatd against this WHR !!!! "
   Exit Sub
End If
Dim ans As Variant
'On Error GoTo msgError
mLicenseNo = GetLicenseNo(RsTxn_WHRDetail!LicenseID)
ans = MsgBox("Are you Sure you want to Delete this Record", vbYesNo)
If ans = vbYes Then
    RsTxn_WHRDetail.Delete
    RsTxn_WHRDetail.Update
    If RsTxn_WHRDetail.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
    End If
    
    RsTxn_WHRDetail.MoveNext
    If RsTxn_WHRDetail.EOF() Then
      RsTxn_WHRDetail.MoveLast
    End If
    Call InData
End If
Exit Sub
msgError:
MsgBox "Child Record Present "

End Sub

Private Sub EditCmd_Click()
'  Edit_Flg = "E"
'  CmbCommodityCondition.Locked = False
'  TxtGrade.Locked = False
'  TxtClientID.Visible = True
'  CmbClientID.Visible = False
'  CmbClientID.Locked = True
'  TxtWHRCommodityId.Visible = True
'  CmbWHRCommodity.Visible = False
'  CmbWHRCommodity.Locked = True
'  TxtWeigher.Locked = False
'  TxtAssayer.Locked = False
'  TxtAssayerCerti.Locked = False
'  GradeValidUpto.Enabled = False
'  TxtStorageCharges.Locked = False
'  TxtWHRNoOfBags.Locked = True
'  TxtWHRQty.Locked = True
'  TxtInsuranceCo.Visible = False
'  TxtPolicyNo.Visible = False
'  CmbInsuranceCo.Visible = True
'  CmbPolicyNo.Visible = True
'  CmbInsuranceCo.Text = TxtInsuranceCo
'  CmbPolicyNo.Text = TxtPolicyNo
'  ChkSaidToContain.Enabled = True
'  ChkEstimatedWgt.Enabled = True
'  ChkSelection = 0
'  Call GButtonLock(Me, False)
End Sub

Private Sub ExitCmd_Click()
Unload Me
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
Call TableTxn_WHRDetail("")
If RsTxn_WHRDetail.RecordCount = 0 Then
  If UserType = "Auditor" Then
     Call GButtonLock(Me, False)
     CmdPrint.Enabled = True
     SaveCmd.Enabled = False
     SearchCmd.Enabled = False
     Exit Sub
  End If
  Call AddCmd_Click
  Exit Sub
End If

Call InData
End Sub
Private Sub InData()
  WHRDate.Enabled = False
  CmbCommodityCondition.Locked = True
  StorageFrom.Enabled = False
  StorageTo.Enabled = False
  TxtGrade.Locked = True
  TxtClientID.Visible = True
  CmbClientId.Visible = False
  CmbClientId.Locked = True
  TxtWHRCommodityId.Visible = True
  CmbWHRCommodity.Visible = False
  CmbWHRCommodity.Locked = True
  TxtWeigher.Locked = True
  TxtAssayer.Locked = True
  TxtAssayerCerti.Locked = True
  GradeValidUpto.Enabled = False
  TxtStorageCharges.Locked = True
  TxtRate.Locked = True
  TxtWHRNoOfBags.Locked = True
  TxtWHRQty.Locked = True
  TxtInsuranceCo.Visible = True
  TxtPolicyNo.Visible = True
  CmbInsuranceCo.Visible = False
  CmbPolicyNo.Visible = False
  CmdInsuranceCo.Visible = False
  CmdInsurance.Visible = False
  ChkSaidToContain.Enabled = False
  ChkEstimatedWgt.Enabled = False
  
  CmdUpdate.Enabled = False
  Call GButtonLock(Me, True)
  CmdPrint.Enabled = True
  mExchangeTypeID = GetExchangeID("NON NCDEX")
  Call Grid_Head
  TxtWHRNo = IIf(IsNull(RsTxn_WHRDetail!WHRNo), "", RsTxn_WHRDetail!WHRNo)
  TxtWHRNoNew = IIf(IsNull(RsTxn_WHRDetail!CMPID), 0, RsTxn_WHRDetail!CMPID) & "-" & Val(TxtWHRNo) - (IIf(IsNull(RsTxn_WHRDetail!CMPID), 0, RsTxn_WHRDetail!CMPID) * 1000000)
  
  WHRDate.Value = IIf(IsNull(RsTxn_WHRDetail!WHRDate), "", RsTxn_WHRDetail!WHRDate)
  mCommodityCondition = IIf(IsNull(RsTxn_WHRDetail!ConditionID), "", RsTxn_WHRDetail!ConditionID)
  If mCommodityCondition = 0 Then
     CmbCommodityCondition.Text = "Average"
  ElseIf mCommodityCondition = 1 Then
     CmbCommodityCondition.Text = "Fair"
  ElseIf mCommodityCondition = 2 Then
     CmbCommodityCondition.Text = "Good"
  End If
  StorageFrom.Value = IIf(IsNull(RsTxn_WHRDetail!StorageFrom), "", RsTxn_WHRDetail!StorageFrom)
  StorageTo.Value = IIf(IsNull(RsTxn_WHRDetail!StorageTo), "", RsTxn_WHRDetail!StorageTo)
  TxtGrade = IIf(IsNull(RsTxn_WHRDetail!Grade), "", RsTxn_WHRDetail!Grade)
  mClientIDRow = IIf(IsNull(RsTxn_WHRDetail!ClientIDRow), "", RsTxn_WHRDetail!ClientIDRow)
  ans = GetClientNameByRow(mClientIDRow)
  
  If ans(0) = "" Then
     MsgBox "Not Found!!!"
     Exit Sub
  End If
  TxtClientID.Text = ans(1)
  mCommodityId = IIf(IsNull(RsTxn_WHRDetail!CommodityID), "", RsTxn_WHRDetail!CommodityID)
  TxtWHRCommodityId.Text = GetCommodityName(mCommodityId)
  TxtWeigher = IIf(IsNull(RsTxn_WHRDetail!Weigher), "", RsTxn_WHRDetail!Weigher)
  TxtAssayer = IIf(IsNull(RsTxn_WHRDetail!Assayer), "", RsTxn_WHRDetail!Assayer)
  TxtAssayerCerti = IIf(IsNull(RsTxn_WHRDetail!GradeByAssayer), "", RsTxn_WHRDetail!GradeByAssayer)
  GradeValidUpto.Value = IIf(IsNull(RsTxn_WHRDetail!GradeValidTill), "", RsTxn_WHRDetail!GradeValidTill)
  TxtStorageCharges = IIf(IsNull(RsTxn_WHRDetail!StorageCharges), "", RsTxn_WHRDetail!StorageCharges)
  TxtStorageAmount = IIf(IsNull(RsTxn_WHRDetail!StorageAmount), "", RsTxn_WHRDetail!StorageAmount)
  TxtRate = IIf(IsNull(RsTxn_WHRDetail!Rate), "", RsTxn_WHRDetail!Rate)
  TxtPolicyNo = RsTxn_WHRDetail!PolicyNo & " ~ " & RsTxn_WHRDetail!CoverNoteNo
  ChkSaidToContain.Value = IIf(IsNull(RsTxn_WHRDetail!SaidToContain), 0, RsTxn_WHRDetail!SaidToContain)
  ChkEstimatedWgt.Value = IIf(IsNull(RsTxn_WHRDetail!EstimatedWeight), 0, RsTxn_WHRDetail!EstimatedWeight)
  Call TableMst_PolicyDetail("where PolicyNo = '" & RsTxn_WHRDetail!PolicyNo & "' And CoverNoteNo = '" & RsTxn_WHRDetail!CoverNoteNo & "'")
  If RsMst_PolicyDetail.RecordCount > 0 Then
     TxtInsuranceCo = GetInsuranceCoName(RsMst_PolicyDetail!InsuranceCoID)
     ValidFrom.Value = RsMst_PolicyDetail!ValidFrom
     ValidTo.Value = RsMst_PolicyDetail!ValidTo
     TxtAmount = RsMst_PolicyDetail!Amount
     TxtRisks = RsMst_PolicyDetail!Risks
     CmdInsuranceCo.Visible = False
     CmdInsurance.Visible = False
     RsMst_PolicyDetail.Close
  End If
  

  
  Call GetGridInData
  
  TxtStorageAmount = Val(MSHFlexGrid2.TextMatrix(2, 7)) * Val(TxtStorageCharges)
  TxtValueofCommodity = Val(MSHFlexGrid2.TextMatrix(2, 8)) * Val(TxtRate)
  
  EditCmd.Enabled = False
  Frame3.Visible = False
  SSTab1.Tab = 0
  'Call GridInData
End Sub
Private Sub GetGridInData()
mLicenseID = IIf(IsNull(RsTxn_WHRDetail!LicenseID), "", RsTxn_WHRDetail!LicenseID)
k = 1
k = k + 1
Call TableMst_License(" Where LicenseID = " & mLicenseID)
MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
MSHFlexGrid2.TextMatrix(k, 1) = GetCMPName(RsTxn_WHRDetail!CMPID)
MSHFlexGrid2.TextMatrix(k, 2) = GetGodownName(RsTxn_WHRDetail!GodownID, "N")
MSHFlexGrid2.TextMatrix(k, 3) = RsMst_License!LicenseNo
MSHFlexGrid2.TextMatrix(k, 4) = RsMst_License!LicenseDesc
MSHFlexGrid2.TextMatrix(k, 5) = 0
MSHFlexGrid2.TextMatrix(k, 6) = 0
MSHFlexGrid2.TextMatrix(k, 7) = RsTxn_WHRDetail!WHRNoOfBags
MSHFlexGrid2.TextMatrix(k, 8) = RsTxn_WHRDetail!WHRQuantity
MSHFlexGrid2.TextMatrix(k, 10) = IIf(IsNull(RsTxn_WHRDetail!DeliveredQuantity), 0, RsTxn_WHRDetail!DeliveredQuantity)
'MSHFlexGrid2.TextMatrix(k, 11) = RsTxn_WHRDetail!DeliveredQuantity
MSHFlexGrid2.TextMatrix(k, 12) = RsTxn_WHRDetail!sm_prefix

End Sub
Public Sub Grid_Head()
MSHFlexGrid2.Clear
MSHFlexGrid2.Cols = 13
MSHFlexGrid2.Rows = 3
MSHFlexGrid2.FixedRows = 2

MSHFlexGrid2.TextMatrix(1, 0) = "   "
MSHFlexGrid2.TextMatrix(1, 1) = "CMP Name "
MSHFlexGrid2.TextMatrix(1, 2) = "Godown No"
MSHFlexGrid2.TextMatrix(1, 3) = "License No"
MSHFlexGrid2.TextMatrix(1, 4) = "License Description "
MSHFlexGrid2.TextMatrix(1, 5) = "Available No Of Bags"
MSHFlexGrid2.TextMatrix(1, 6) = "Available Electronic Weight"
MSHFlexGrid2.TextMatrix(1, 7) = "WHR No Of Bags"
MSHFlexGrid2.TextMatrix(1, 8) = "WHR Quantity"
MSHFlexGrid2.TextMatrix(1, 9) = "WHR LicenseDetailID"
MSHFlexGrid2.TextMatrix(1, 10) = "Delivered Qty"
MSHFlexGrid2.TextMatrix(1, 11) = "Agreement End Date"
MSHFlexGrid2.TextMatrix(1, 12) = "Prefix"

MSHFlexGrid2.ColWidth(0) = 400
MSHFlexGrid2.ColWidth(1) = 1400
MSHFlexGrid2.ColWidth(2) = 1400
MSHFlexGrid2.ColWidth(3) = 1200
MSHFlexGrid2.ColWidth(4) = 0
MSHFlexGrid2.ColWidth(5) = 1600
MSHFlexGrid2.ColWidth(6) = 1800
MSHFlexGrid2.ColWidth(7) = 1800
MSHFlexGrid2.ColWidth(8) = 1600
MSHFlexGrid2.ColWidth(9) = 0
MSHFlexGrid2.ColWidth(12) = 0

MSHFlexGrid2.RowHeight(1) = 500
End Sub
Public Sub Grid_Head1()
MSHFlexGrid1.Clear
MSHFlexGrid1.Cols = 12
MSHFlexGrid1.Rows = 3
MSHFlexGrid1.FixedRows = 1

MSHFlexGrid1.TextMatrix(0, 0) = "WHR No"
MSHFlexGrid1.TextMatrix(0, 1) = "Location ID"
MSHFlexGrid1.TextMatrix(0, 2) = "Condition Of Commodity"
MSHFlexGrid1.TextMatrix(0, 3) = "Deposit Date"
MSHFlexGrid1.TextMatrix(0, 4) = "Storage upto Date"
MSHFlexGrid1.TextMatrix(0, 5) = "Client Name"
MSHFlexGrid1.TextMatrix(0, 6) = "Commodity"
MSHFlexGrid1.TextMatrix(0, 7) = "Private Marks on Packaging"
MSHFlexGrid1.TextMatrix(0, 8) = "GUID"
MSHFlexGrid1.TextMatrix(0, 9) = "Bags"
MSHFlexGrid1.TextMatrix(0, 10) = "Quantity"
MSHFlexGrid1.TextMatrix(0, 11) = "Flag"

MSHFlexGrid1.ColWidth(0) = 1000
MSHFlexGrid1.ColWidth(1) = 1000
MSHFlexGrid1.ColWidth(2) = 1800
MSHFlexGrid1.ColWidth(3) = 1400
MSHFlexGrid1.ColWidth(4) = 1400
MSHFlexGrid1.ColWidth(5) = 3000
MSHFlexGrid1.ColWidth(6) = 1600
MSHFlexGrid1.ColWidth(7) = 1200
MSHFlexGrid1.ColWidth(8) = 0
MSHFlexGrid1.ColWidth(9) = 800
MSHFlexGrid1.ColWidth(10) = 800
MSHFlexGrid1.ColWidth(11) = 600


MSHFlexGrid2.RowHeight(0) = 500
End Sub

Private Sub MSHFlexGrid1_Click()
Dim i As Variant
i = MSHFlexGrid1.RowSel
If i <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(i, 0) = "" Then Exit Sub
RsTxn_WHRDetail.MoveFirst
RsTxn_WHRDetail.Find "G_Uid = '" & MSHFlexGrid1.TextMatrix(i, 8) & "'"
If Not RsTxn_WHRDetail.EOF Then
   Call InData
   Frame0.Visible = False
End If
End Sub

Private Sub MSHFlexGrid2_Click()

If SaveCmd.Enabled = False Then Exit Sub

If MSHFlexGrid2.Rows > 3 Then
    If MSHFlexGrid2.RowSel <> "0" Then
       If MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 1) = "" Then
          TxtCMPID = ""
          TxtGodownNo = ""
          TxtLicenseNo = ""
          TxtLicenseDescription = ""
          TxtNoOfBags = ""
          TxtEleWeight = ""
          TxtWHRNoOfBags = ""
          TxtWHRQty = ""
          CmdUpdate.Enabled = False
          Frame2.Enabled = False
          Exit Sub
       End If
       Frame2.Enabled = True
       If ChkSelection = 1 Then
          If MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 0) <> "*" Then
             MsgBox "You can't select more than 1 row. To discard the existing selection, click on commodity name !!!!"
             Exit Sub
          End If
       End If
      TxtWHRNoOfBags.Locked = False
      TxtWHRQty.Locked = False
      TxtCMPID = MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 1)
      TxtGodownNo = MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 2)
      
      TxtLicenseNo = MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 3)
      TxtLicenseDescription = MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 4)
      TxtNoOfBags = Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 5)) + Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 7))
      TxtEleWeight = Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 6)) + Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 8))
      TxtWHRNoOfBags = Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 5)) + Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 7))
      TxtWHRQty = Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 6)) + Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 8))
      'TxtWHRNoOfBags.SetFocus
      CmdUpdate.Enabled = True
      CmdUpdate.SetFocus
    End If
End If
End Sub

Private Sub SaveCmd_Click()


If CmbCommodityCondition.Text = "" Then
   MsgBox "Blank Commodity Condition Not Allowed!!!"
   CmbCommodityCondition.SetFocus
   Exit Sub
End If



If CmbCommodityCondition.Text = "" Then
   MsgBox "Blank Commodity Condition Not Allowed!!!"
   CmbCommodityCondition.SetFocus
   Exit Sub
End If

If WHRDate.Value < StorageFrom.Value Then
   MsgBox "WHR Date must Be Greater Than Storage From Date!!!"
   WHRDate.SetFocus
   Exit Sub
End If
If TxtAgreeDate <> "" Then
   If StorageTo.Value > CDate(TxtAgreeDate) Then
      MsgBox "Storage To Date must Be less or equal to Agreement End Date !!!"
      StorageTo.SetFocus
      Exit Sub
   End If
End If
If WHRDate.Value > StorageTo.Value Then
   MsgBox "WHR Date must Be Less Than Storage To Date!!!"
   WHRDate.SetFocus
   Exit Sub
End If
If StorageFrom.Value > StorageTo.Value Then
   MsgBox "Storage To Date must Be Greater Than Storage From Date!!!"
   StorageTo.SetFocus
   Exit Sub
End If

If CmbInsuranceCo.Text = "" Then
   MsgBox "Blank InsuranceCompany Not Allowed!!!"
   CmbInsuranceCo.SetFocus
   Exit Sub
End If
      
If CmbPolicyNo.Text = "" Then
   MsgBox "'Insurance policy' can't be empty !!!"
   CmbPolicyNo.SetFocus
   Exit Sub
End If

If TxtGrade = "" Then
   MsgBox "'Private Marks on Packaging' can't be empty !!!"
   TxtGrade.SetFocus
   Exit Sub
End If

If Edit_Flg = "A" Then
    If CmbClientId.Text = "" Then
       MsgBox "'Client Name' can't be empty  !!!"
       CmbClientId.SetFocus
       Exit Sub
    End If
    
    If CmbWHRCommodity.Text = "" Then
       MsgBox "'Commodity' can't be empty !!!"
       CmbWHRCommodity.SetFocus
       Exit Sub
    End If
End If
If TxtStorageCharges.Text = "" Then
   MsgBox "Storage Charges can't be empty !!!"
   TxtStorageCharges.SetFocus
   Exit Sub
End If
If TxtRate.Text = "" Then
   MsgBox "Commodity price can't be empty !!!"
   TxtRate.SetFocus
   Exit Sub
End If
If GradeValidUpto.Value < StorageTo.Value Then
   MsgBox "'Grade Valid till' Date should not be less than 'Storage To' date"
   GradeValidUpto.SetFocus
   Exit Sub
End If

If MSHFlexGrid2.Rows < 4 Then
   MsgBox "No deposits available for generating WHR !!!!"
   Exit Sub
End If
If ChkSelection = 0 Then
   MsgBox "Select atleast one godown from 'GSL Detail' tab for which WHR is to be Generate  !!!!!"
   
   Exit Sub
End If

If Edit_Flg = "A" Then
   RsTxn_WHRDetail.AddNew
   RsTxn_WHRDetail!WHRNo = GetWHRNo
   RsTxn_WHRDetail!G_UID = GetGUID
   TxtWHRNo.Text = RsTxn_WHRDetail!WHRNo
Else
  RsTxn_WHRDetail.MoveFirst
  RsTxn_WHRDetail.Find "WHRNo= '" & TxtWHRNo.Text & "'"
  If Not RsTxn_WHRDetail.EOF Then
    If RsTxn_WHRDetail!WHRNo <> TxtWHRNo.Text Then
      MsgBox "Duplicate WHR No Not Allowed !!! "
     'TxtISINRevalReqNo.SetFocus
      Exit Sub
    End If
  End If
  RsTxn_WHRDetail.MoveFirst
  RsTxn_WHRDetail.Find " WHRNo= " & TxtWHRNo.Text
End If
RsTxn_WHRDetail!WHRDate = WHRDate.Value
RsTxn_WHRDetail!ConditionID = mCommodityCondition
RsTxn_WHRDetail!StorageFrom = StorageFrom.Value
RsTxn_WHRDetail!StorageTo = StorageTo.Value
RsTxn_WHRDetail!Grade = TxtGrade
RsTxn_WHRDetail!Weigher = TxtWeigher
RsTxn_WHRDetail!Assayer = TxtAssayer
RsTxn_WHRDetail!GradeByAssayer = TxtAssayerCerti
RsTxn_WHRDetail!GradeValidTill = GradeValidUpto.Value
RsTxn_WHRDetail!StorageCharges = TxtStorageCharges
RsTxn_WHRDetail!PolicyNo = PolicyNo_
RsTxn_WHRDetail!CoverNoteNo = CNo
RsTxn_WHRDetail!Rate = TxtRate
RsTxn_WHRDetail!SaidToContain = ChkSaidToContain.Value
RsTxn_WHRDetail!EstimatedWeight = ChkEstimatedWgt.Value
RsTxn_WHRDetail!StorageAmount = TxtStorageAmount
RsTxn_WHRDetail!ClientIDRow = mClientIDRow
RsTxn_WHRDetail!CommodityID = mCommodityId
'RsTxn_WHRDetail.Update
'Call InData
    
For i = 2 To MSHFlexGrid2.Rows - 1
  If MSHFlexGrid2.TextMatrix(i, 0) = "*" Then
    RsTxn_WHRDetail!CMPID = GetCMPID(MSHFlexGrid2.TextMatrix(i, 1))
    RsTxn_WHRDetail!GodownID = GetGodownID(MSHFlexGrid2.TextMatrix(i, 2), RsTxn_WHRDetail!CMPID)
    RsTxn_WHRDetail!LicenseID = GetLicenseID(MSHFlexGrid2.TextMatrix(i, 3))
    If Edit_Flg = "A" Then
       RsTxn_WHRDetail!oldLicenseID = RsTxn_WHRDetail!LicenseID 'GetLicenseID(MSHFlexGrid2.TextMatrix(i, 2))
    End If
    RsTxn_WHRDetail!WHRNoOfBags = MSHFlexGrid2.TextMatrix(i, 7)
    RsTxn_WHRDetail!WHRQuantity = MSHFlexGrid2.TextMatrix(i, 8)
    RsTxn_WHRDetail!sm_prefix = MSHFlexGrid2.TextMatrix(i, 12)
    
    RsTxn_WHRDetail.Update
    'Call PostGSLData(RsTxn_WHRDetail!LicenseID, RsTxn_WHRDetail!WHRNoOfBags, RsTxn_WHRDetail!WHRQuantity)
    Exit For
  End If
Next
Call InData
End Sub

Function PostGSLData(LicenseID_ As Double, WHRNoBags_ As Double, WHRQty_ As Double)
Dim TBags, TQty, PBags, PQty As Variant

Call TableMst_GSL
tmp = "SELECT Mst_GSL.GSLID,(Mst_GSL.DepositeBags)-(Mst_GSL.WHRBags) as BalanceBags,(Mst_GSL.DepositeWeight)-(Mst_GSL.WHRWeight) as BalanceWeight,Mst_License.LicenseNo , " & _
      "  Mst_License.LicenseDesc, Mst_Godown.CMPID,Mst_GSL.WHRBags,Mst_GSL.WHRWeight From Mst_License " & _
      " INNER JOIN (Mst_Godown INNER JOIN Mst_GSL ON Mst_Godown.GodownID = Mst_GSL.GodownID) " & _
      " ON Mst_License.LicenseID = Mst_Godown.LicenseID Where ((Mst_GSL.ClientIDRow = " & mClientIDRow & ") And " & _
      " (Mst_GSL.ExchangeTypeID = " & mExchangeTypeID & ") And " & _
      " (Mst_GSL.CommodityID = " & mCommodityId & ") " & _
      " AND Mst_License.LicenseNo = '" & mLicenseNo & "') And (((Mst_GSL.DepositeBags)-(Mst_GSL.WHRBags)) > 0 OR ((Mst_GSL.DepositeWeight)-(Mst_GSL.WHRWeight)) > 0)  "
Call Tmp1Table

TBags = WHRNoBags_
TQty = WHRQty_
For i = 1 To TmpRs1.RecordCount
    If TBags >= TmpRs1!BalanceBags Then
       PBags = TmpRs1!BalanceBags
       TBags = TBags - PBags
    Else
       PBags = TBags
       TBags = 0
    End If
    TmpRs1!WHRBags = IIf(IsNull(TmpRs1!WHRBags), 0, TmpRs1!WHRBags) + PBags
    'TmpRs1!BalanceBags = TmpRs1!BalanceBags - PBags
    TmpRs1.Update
    If TBags = 0 Then Exit For
    TmpRs1.MoveNext
Next
TmpRs1.MoveFirst
TQty = WHRQty_
For i = 1 To TmpRs1.RecordCount
    If TQty >= TmpRs1!BalanceWeight Then
       PQty = TmpRs1!BalanceWeight
       TQty = TQty - PQty
    Else
       PQty = TQty
       TQty = 0
    End If
    TmpRs1!WHRWeight = IIf(IsNull(TmpRs1!WHRWeight), 0, TmpRs1!WHRWeight) + PQty
    'TmpRs1!BalanceWeight = TmpRs1!BalanceWeight - PQty
    TmpRs1.Update
    If TQty = 0 Then Exit For
    TmpRs1.MoveNext
Next
TmpRs1.Close

End Function
Private Function GetWHRNo() As Double
Dim RetVal As Double
tmp = "Select WHRNo from Txn_WHR_Detail Where CMPID = " & mCMPID & " Order by WHRNo"
Call TmpTable

If Not TmpRs.EOF Then
    TmpRs.MoveLast
    mWHRNo = TmpRs.Fields(0)
    mWHRNo = mWHRNo + 1
Else
    mWHRNo = mCMPID * 1000000
    mWHRNo = mWHRNo + 1
End If

'TmpRs.MoveLast
RetVal = mWHRNo 'IIf(IsNull(TmpRs!WHRNo), 0, TmpRs!WHRNo)
GetWHRNo = RetVal
End Function
Private Function GetWHRLicenseDetailID() As Double
Dim RetVal As Double
tmp = "Select max(WHRLicenseDetailID) from Txn_WHRLicense_Detail"
Call TmpTable
RetVal = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
RetVal = RetVal + 1
GetWHRLicenseDetailID = RetVal
End Function

Private Sub StorageFrom_LostFocus()
Call GetGridData(mExchangeTypeID, mCommodityId, mClientIDRow)


End Sub
Private Sub StorageTo_LostFocus()
GradeValidUpto.Value = StorageTo.Value
End Sub

Private Sub TxtRate_LostFocus()
If TxtRate = "" Then Exit Sub
If IsNumeric(TxtRate) = False Then
  MsgBox "Invalid Number Format!!!"
  TxtRate.SetFocus
  Exit Sub
End If
End Sub

Private Sub TxtStorageCharges_LostFocus()
If TxtStorageCharges = "" Then Exit Sub
If IsNumeric(TxtStorageCharges) = False Then
  MsgBox "Invalid Number Format!!!"
  TxtStorageCharges.SetFocus
  Exit Sub
End If
End Sub

Private Sub TxtWHRNoOfBags_LostFocus()
If TxtWHRNoOfBags = "" Then Exit Sub
If IsNumeric(TxtWHRNoOfBags) = False Then
  MsgBox "Invalid Number Format!!!"
  TxtWHRNoOfBags.SetFocus
  Exit Sub
End If
If Val(TxtWHRNoOfBags) < 0 Then
  MsgBox "WHR No OF Bags Can't Be Negetive"
  TxtWHRNoOfBags.SetFocus
  Exit Sub
End If
If Val(TxtWHRNoOfBags) > Val(TxtNoOfBags) Then
   MsgBox "WHR No Of Bags Should Be Less Than Electronic No Of Bags!!!"
   TxtWHRNoOfBags.SetFocus
   Exit Sub
End If
TxtWHRQty.SetFocus
End Sub

Private Sub TxtWHRQty_LostFocus()
If TxtWHRQty = "" Then Exit Sub
If IsNumeric(TxtWHRQty) = False Then
  MsgBox "Invalid Number Format!!!"
  TxtWHRQty.SetFocus
  Exit Sub
End If
If Val(TxtWHRQty) < 0 Then
  MsgBox "WHR Quantity Can't Be Negetive"
  TxtWHRQty.SetFocus
  Exit Sub
End If
If Val(TxtWHRQty) > Val(TxtEleWeight) Then
   MsgBox "WHR Weight Should Be Less Than Electronic Weight"
   TxtWHRQty.SetFocus
   Exit Sub
End If
CmdUpdate.Enabled = True
'CmdUpdate.SetFocus
End Sub
Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsTxn_WHRDetail.MoveNext
If RsTxn_WHRDetail.EOF Then
   RsTxn_WHRDetail.MoveLast
   LF_Flag = "Y"
End If
Call InData
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
RsTxn_WHRDetail.MovePrevious
If RsTxn_WHRDetail.BOF Then
   RsTxn_WHRDetail.MoveFirst
   LF_Flag = "Y"
End If
Call InData
If LF_Flag = "Y" Then
    PreviousCmd.Enabled = False
    FirstCmd.Enabled = False
End If

NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub FirstCmd_Click()
RsTxn_WHRDetail.MoveFirst
Call InData
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsTxn_WHRDetail.MoveLast
Call InData
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False
End Sub

Private Sub SearchCmd_Click()
Dim x As Integer
If SearchCmd.Caption = "Cancel" Then
   If RsEmployee.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   If RsEmployee.EOF Then
      RsEmployee.MoveFirst
   End If
   Call InData
   Exit Sub
End If

Frame0.Height = Me.Height - 500
Frame0.Left = Frame1.Left
Frame0.Top = Frame1.Top
Frame0.Width = Me.Width - 400
Frame0.Visible = True
MSHFlexGrid1.Clear
MSHFlexGrid1.Width = Frame0.Width - 300
MSHFlexGrid1.Height = Frame0.Height - 1200

Call Grid_Head1
TxtSearchWHR = ""
CmbSearchStatus.Text = ""
End Sub

Private Sub OptClientName_Click()
Call OptWHRDate_Click
End Sub

Private Sub OptCMPName_Click()
Call OptWHRDate_Click
End Sub

Private Sub OptCommodity_Click()
Call OptWHRDate_Click
End Sub

Private Sub OptGodown_Click()
Call OptWHRDate_Click
End Sub

Private Sub OptLocationName_Click()
Call OptWHRDate_Click
End Sub

Private Sub OptWHRDate_Click()
LblSearchWHR.Visible = True
TxtSearchWHR.Visible = True
TxtSearchWHR = ""
'CmbSearchStatus.Visible = False
CmbSearchStatus.Text = ""
CmdGo.Visible = True

If OptWHRDate.Value = True Then
   LblSearchWHR.Caption = "Search By WHR Date"
ElseIf OptWHRNo.Value = True Then
   LblSearchWHR.Caption = "Search By WHR No"
ElseIf OptLocationName.Value = True Then
   LblSearchWHR.Caption = "Search By Location Name"
ElseIf OptCMPName.Value = True Then
   LblSearchWHR.Caption = "Search By CMP Name"
ElseIf OptGodown.Value = True Then
   LblSearchWHR.Caption = "Search By Godown"
ElseIf OptClientName.Value = True Then
   LblSearchWHR.Caption = "Search By Client Name"
ElseIf OptCommodity.Value = True Then
   LblSearchWHR.Caption = "Search By Commodity Name"
End If

End Sub

Private Sub OptWHRNo_Click()
Call OptWHRDate_Click
End Sub

Private Sub CmdGo_Click()
'Dim tmp As Variant
Dim Wc As Variant
Wc = ""
tmp = ""
tmp = "SELECT  TWD.WHRNo,TWD.SM_Prefix,TWD.WHRDate,TWD.ConditionID,TWD.StorageFrom,TWD.StorageTo,TWD.Grade,TWD.Weigher,TWD.Assayer,TWD.GradeByAssayer, " & _
    "TWD.GradeValidTill,TWD.StorageCharges,TWD.StorageAmount,TWD.ClientIDRow,TWD.CommodityID,TWD.CMPID,TWD.LicenseID,TWD.WHRNoofBags,TWD.WHRQuantity, " & _
    "TWD.PrintFlag,TWD.Remark,TWD.Flag,TWD.PledgeID,TWD.OldLicenseID,TWD.GodownID,TWD.DeliveredQuantity,TWD.DeliveredNoOfBags,TWD.Rate,TWD.G_UID, " & _
    "TWD.HologramNo , TWD.PolicyNo, TWD.CoverNoteNo, TWD.SaidToContain, TWD.EstimatedWeight " & _
    "FROM    Txn_WHR_Detail TWD Inner Join Mst_Location LM ON TWD.SM_Prefix=LM.SM_Prefix Inner Join Mst_CMP CMP ON TWD.CMPID=CMP.CMPID " & _
    "Inner Join Mst_Godown GM ON TWD.GodownID=GM.GodownID Inner Join Mst_Client CM ON TWD.ClientIDRow=CM.ClientIDRow " & _
    "Inner Join Mst_Commodity ComM ON TWD.CommodityID=ComM.CommodityID"

If OptWHRDate.Value = True Then
   Wc = " Where TWD.WHRDate= '" & TxtSearchWHR & "' "
End If

If OptWHRNo.Value = True Then
   Wc = " Where TWD.WHRNo  Like '%" & TxtSearchWHR & "%'"
End If

If OptLocationName.Value = True Then
    Wc = " Where LM.LocationName  Like '%" & TxtSearchWHR & "%'"
End If

If OptCMPName.Value = True Then
    Wc = " Where CMP.CMPName Like '%" & TxtSearchWHR & "%' "
End If

If OptGodown.Value = True Then
    Wc = " Where GM.Address  Like '%" & TxtSearchWHR & "%'"
End If

If OptClientName.Value = True Then
    Wc = " Where CM.ClientName  Like '%" & TxtSearchWHR & "%'"
End If

If OptCommodity.Value = True Then
    Wc = " Where ComM.Commodity  Like '%" & TxtSearchWHR & "%'"
End If

If CmbSearchStatus.Text <> "" Then
   If Wc = "" Then
      
      If CmbSearchStatus.Text = "Printed" Then
         Wc = " Where TWD.Flag = 'P'"
      Else
         Wc = " Where TWD.flag = ''"
      End If
   Else
      If CmbSearchStatus.Text = "Printed" Then
         Wc = Wc & " And TWD.Flag = 'P'"
      Else
         Wc = Wc & " And TWD.Flag = ''"
   End If
End If
End If

tmp = tmp & Wc
Call Tmp3Table
Call Grid_Head1

MSHFlexGrid1.Rows = TmpRs3.RecordCount + 1
For i = 1 To TmpRs3.RecordCount
    For c = 0 To MSHFlexGrid1.Cols - 1
       MSHFlexGrid1.TextMatrix(i, c) = IIf(IsNull(TmpRs3.Fields(c)), "", TmpRs3.Fields(c))
    Next
    If TmpRs3!ConditionID = 0 Then
       MSHFlexGrid1.TextMatrix(i, 2) = "Average"
    ElseIf TmpRs3!ConditionID = 1 Then
       MSHFlexGrid1.TextMatrix(i, 2) = "Fair"
    ElseIf TmpRs3!ConditionID = 2 Then
       MSHFlexGrid1.TextMatrix(i, 2) = "Good"
    End If
    MSHFlexGrid1.TextMatrix(i, 0) = IIf(IsNull(TmpRs3!CMPID), 0, TmpRs3!CMPID) & "-" & TmpRs3!WHRNo - (IIf(IsNull(TmpRs3!CMPID), 0, TmpRs3!CMPID) * 1000000)
    MSHFlexGrid1.TextMatrix(i, 3) = IIf(IsNull(TmpRs3!StorageFrom), "", TmpRs3!StorageFrom)
    MSHFlexGrid1.TextMatrix(i, 4) = IIf(IsNull(TmpRs3!StorageTo), "", TmpRs3!StorageTo)
    MSHFlexGrid1.TextMatrix(i, 5) = GetClientName(TmpRs3!ClientIDRow)
    MSHFlexGrid1.TextMatrix(i, 6) = GetCommodityName(TmpRs3!CommodityID)
    MSHFlexGrid1.TextMatrix(i, 7) = IIf(IsNull(TmpRs3!Grade), "", TmpRs3!Grade)
    MSHFlexGrid1.TextMatrix(i, 8) = IIf(IsNull(TmpRs3!G_UID), "", TmpRs3!G_UID)
    MSHFlexGrid1.TextMatrix(i, 9) = IIf(IsNull(TmpRs3!WHRNoOfBags), "", TmpRs3!WHRNoOfBags) 'WHRNoOfBags
    MSHFlexGrid1.TextMatrix(i, 10) = IIf(IsNull(TmpRs3!WHRQuantity), "", TmpRs3!WHRQuantity)
    MSHFlexGrid1.TextMatrix(i, 11) = IIf(IsNull(TmpRs3!Flag), "", TmpRs3!Flag)
  TmpRs3.MoveNext
Next
LblRecordCount.Refresh
LblRecordCount.Caption = TmpRs3.RecordCount
End Sub



Private Sub SearchCmd_Click0ld()
If SearchCmd.Caption = "Cancel" Then
   If RsTxn_WHRDetail.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   Call FirstCmd_Click
   Exit Sub
End If
Frame0.Height = Frame1.Height
Frame0.Left = Frame1.Left
Frame0.Top = Frame1.Top
Frame0.Width = Frame1.Width
Frame0.Visible = True
MSHFlexGrid1.Clear
MSHFlexGrid1.Cols = 3
MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 1000
RsTxn_WHRDetail.MoveFirst
Call Grid_Head1
MSHFlexGrid1.Rows = RsTxn_WHRDetail.RecordCount + 1
For i = 1 To RsTxn_WHRDetail.RecordCount
    For c = 0 To MSHFlexGrid1.Cols - 1
       MSHFlexGrid1.TextMatrix(i, c) = IIf(IsNull(RsTxn_WHRDetail.Fields(c)), "", RsTxn_WHRDetail.Fields(c))
    Next
    If RsTxn_WHRDetail!ConditionID = 0 Then
       MSHFlexGrid1.TextMatrix(i, 2) = "Average"
    ElseIf RsTxn_WHRDetail!ConditionID = 1 Then
       MSHFlexGrid1.TextMatrix(i, 2) = "Fair"
    ElseIf RsTxn_WHRDetail!ConditionID = 2 Then
       MSHFlexGrid1.TextMatrix(i, 2) = "Good"
    End If
    
    MSHFlexGrid1.TextMatrix(i, 0) = IIf(IsNull(RsTxn_WHRDetail!CMPID), 0, RsTxn_WHRDetail!CMPID) & "-" & RsTxn_WHRDetail!WHRNo - (IIf(IsNull(RsTxn_WHRDetail!CMPID), 0, RsTxn_WHRDetail!CMPID) * 1000000)
    MSHFlexGrid1.TextMatrix(i, 3) = IIf(IsNull(RsTxn_WHRDetail!StorageFrom), "", RsTxn_WHRDetail!StorageFrom)
    MSHFlexGrid1.TextMatrix(i, 5) = GetClientName(RsTxn_WHRDetail!ClientIDRow)
    MSHFlexGrid1.TextMatrix(i, 6) = GetCommodityName(RsTxn_WHRDetail!CommodityID)
    MSHFlexGrid1.TextMatrix(i, 7) = IIf(IsNull(RsTxn_WHRDetail!Grade), "", RsTxn_WHRDetail!Grade)
    MSHFlexGrid1.TextMatrix(i, 8) = IIf(IsNull(RsTxn_WHRDetail!G_UID), "", RsTxn_WHRDetail!G_UID)
    MSHFlexGrid1.TextMatrix(i, 9) = IIf(IsNull(RsTxn_WHRDetail!WHRNoOfBags), "", RsTxn_WHRDetail!WHRNoOfBags) 'WHRNoOfBags
    MSHFlexGrid1.TextMatrix(i, 10) = IIf(IsNull(RsTxn_WHRDetail!WHRQuantity), "", RsTxn_WHRDetail!WHRQuantity)
    MSHFlexGrid1.TextMatrix(i, 11) = IIf(IsNull(RsTxn_WHRDetail!Flag), "", RsTxn_WHRDetail!Flag)

    
  RsTxn_WHRDetail.MoveNext
Next


End Sub

Private Sub CmdExcel_Click()
Call Gen_Create_Xls("Txn_WHRDetail.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\Txn_WHRDetail.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For i = 0 To MSHFlexGrid1.Rows - 1
    mRow = i + 1 ' + 6
    For c = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(i, c), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(i, c)) Then
              oWS.Cells(mRow, c + 1) = Format(MSHFlexGrid1.TextMatrix(i, c), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(i, c)
           End If
        Else
           oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(i, c)
        End If
    Next
Next
oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub
Private Sub WHRDate_LostFocus()
Call GetGridData(mExchangeTypeID, mCommodityId, mClientIDRow)
End Sub
Private Sub TxtHologramNo_LostFocus()
If TxtHologramNo = "" Then Exit Sub
If RsMst_Hologram.RecordCount > 0 Then
   RsMst_Hologram.MoveFirst
   RsMst_Hologram.Find "HologramNo = '" & TxtHologramNo & "' "
End If
If Not RsMst_Hologram.EOF Then
   If RsMst_Hologram!DocType = "WHR" Then
      tmp = "Select CMPID from Txn_WHR_Detail Where WHRNo = " & RsMst_Hologram!DocNo & ""
      Call Tmp4Table
      If TmpRs4.RecordCount > 0 Then
         MsgBox "Hologram Already Used In " & RsMst_Hologram!DocType & " " & TmpRs4!CMPID & "-" & (Val(RsMst_Hologram!DocNo) - TmpRs4!CMPID * 1000000)
         TmpRs4.Close
      End If
   Else
      MsgBox "Hologram Already Used In " & RsMst_Hologram!DocType & "-" & RsMst_Hologram!DocNo
   End If
   TxtHologramNo.SetFocus
   CmdSaveHologram.Enabled = False
   Exit Sub
End If
CmdSaveHologram.Enabled = True
CmdSaveHologram.SetFocus
End Sub
Private Sub CmdSaveHologram_Click()
Dim ans As Variant
If TxtHologramNo = "" Then
   MsgBox "Blank Hologram No not allowed!!!"
   TxtHologramNo.SetFocus
   Exit Sub
End If
ans = MsgBox("Are you Sure you want to save this HologramNo", vbYesNo)
If ans = vbYes Then
   RsMst_Hologram.AddNew
   RsMst_Hologram!HologramNo = TxtHologramNo
   RsMst_Hologram!DocType = "WHR"
   RsMst_Hologram!DocNo = TxtWHRNo
   RsMst_Hologram.Update
   RsTxn_WHRDetail!HologramNo = TxtHologramNo
   RsTxn_WHRDetail.Update
   TxtHologramNo = ""
   Frame3.Visible = False
   RsTxn_WHRDetail!Flag = "P"
   RsTxn_WHRDetail.Update
   Gen_mTimeStamp = GetTimeStamp
   
   Call Create_Xls(Gen_mTimeStamp & Gen_UserName & "_WHR" & Trim(TxtWHRNoNew) & ".xls")
   'Call Xls_Print_Rpt
   Call Xls_Print_Rpt_New
Else
   TxtHologramNo.SetFocus
End If

End Sub
Private Sub CmbInsuranceCo_GotFocus()
tmp = CmbInsuranceCo.Text
Call TableMst_InsuranceCo("")

CmbInsuranceCo.Clear
If RsMst_InsuranceCo.RecordCount = 0 Then
   MsgBox "No InsuranceCo found"
   'CmdInsuranceCo.SetFocus
   Exit Sub
End If
For i = 1 To RsMst_InsuranceCo.RecordCount
    CmbInsuranceCo.AddItem RsMst_InsuranceCo!InsuranceCoName
    RsMst_InsuranceCo.MoveNext
Next
CmbInsuranceCo.Text = tmp
   
End Sub
Private Sub CmbInsuranceCo_LostFocus()

If CmbInsuranceCo.Text = "" Then
   Exit Sub
End If
RsMst_InsuranceCo.MoveFirst
RsMst_InsuranceCo.Find "InsuranceCoName = '" & CmbInsuranceCo.Text & "'"

If RsMst_InsuranceCo.EOF Then
   MsgBox "Invalid InsuranceCo selection "
   CmbInsuranceCo.SetFocus
   Exit Sub
End If
mInsuranceCoID = RsMst_InsuranceCo.Fields(0)
If CmbInsuranceCo.Text <> tmp Then
   CmbPolicyNo.Text = ""
   TxtRisks = ""
   TxtAmount = ""
   ValidFrom.Value = Date
   ValidTo.Value = Date
End If

End Sub

Private Sub CmbPolicyNo_GotFocus()
If CmbInsuranceCo.Text = "" Or mInsuranceCoID = "" Then Exit Sub

tmp = CmbPolicyNo.Text
Call TableMst_PolicyDetail(" Where InsuranceCoID =  " & mInsuranceCoID & " Order by  PolicyNo,CoverNoteNo")

CmbPolicyNo.Clear
If RsMst_PolicyDetail.RecordCount = 0 Then
   MsgBox "No Insurance Detail found"
   CmbPolicyNo.SetFocus
   Exit Sub
End If
For i = 1 To RsMst_PolicyDetail.RecordCount
    CmbPolicyNo.AddItem RsMst_PolicyDetail!PolicyNo & " ~ " & RsMst_PolicyDetail!CoverNoteNo
    RsMst_PolicyDetail.MoveNext
Next
CmbPolicyNo.Text = tmp
End Sub

Private Sub CmbPolicyNo_LostFocus()
Dim strSplitString() As String
If CmbPolicyNo.Text = "" Then
   Exit Sub
End If
strSplitString = Split(CmbPolicyNo.Text, "~")
'On Error GoTo M_lbl:
tmp = UBound(strSplitString)
If tmp <> 1 Then
   MsgBox "Invalid Selection"
   CmbPolicyNo.SetFocus
   Exit Sub
End If
PolicyNo_ = Trim(strSplitString(0))
CNo = Trim(strSplitString(1))
RsMst_PolicyDetail.MoveFirst
For i = 1 To RsMst_PolicyDetail.RecordCount
  If PolicyNo_ = Trim(RsMst_PolicyDetail!PolicyNo) And CNo = Trim(RsMst_PolicyDetail!CoverNoteNo) Then
     TxtRisks = RsMst_PolicyDetail!Risks
     TxtAmount = RsMst_PolicyDetail!Amount
     ValidFrom.Value = RsMst_PolicyDetail!ValidFrom
     ValidTo.Value = RsMst_PolicyDetail!ValidTo
     Exit Sub
  End If
  RsMst_PolicyDetail.MoveNext
Next
If RsMst_PolicyDetail.EOF Then
  MsgBox "Invalid InsuranceCo selection "
  CmbPolicyNo.SetFocus
  Exit Sub
End If
mInsuranceCoID = RsMst_PolicyDetail.Fields(0)

End Sub
