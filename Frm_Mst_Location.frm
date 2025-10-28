VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmMst_Location 
   Caption         =   "Location Master"
   ClientHeight    =   4965
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10575
   MDIChild        =   -1  'True
   ScaleHeight     =   12915
   ScaleWidth      =   21360
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   3045
      Left            =   240
      TabIndex        =   13
      Top             =   6480
      Visible         =   0   'False
      Width           =   15000
      Begin VB.ComboBox CmbSRiskType 
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
         ItemData        =   "Frm_Mst_Location.frx":0000
         Left            =   10080
         List            =   "Frm_Mst_Location.frx":000D
         Sorted          =   -1  'True
         TabIndex        =   76
         Top             =   450
         Width           =   1590
      End
      Begin VB.ComboBox CmbSHRLocMapped 
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
         ItemData        =   "Frm_Mst_Location.frx":0024
         Left            =   8220
         List            =   "Frm_Mst_Location.frx":002E
         Sorted          =   -1  'True
         TabIndex        =   72
         Top             =   465
         Width           =   1590
      End
      Begin VB.ComboBox CmbSEQZoneID 
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
         Left            =   5880
         Sorted          =   -1  'True
         TabIndex        =   60
         Top             =   480
         Width           =   2310
      End
      Begin VB.TextBox TxtSearchLocation 
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
         Left            =   120
         TabIndex        =   44
         Top             =   480
         Width           =   2925
      End
      Begin VB.CommandButton CmdGo 
         Caption         =   "Go"
         Height          =   345
         Left            =   11760
         TabIndex        =   43
         Top             =   450
         Width           =   1005
      End
      Begin VB.ComboBox CmbSState 
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
         Left            =   3120
         Sorted          =   -1  'True
         TabIndex        =   41
         Top             =   480
         Width           =   2670
      End
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   350
         Left            =   12765
         TabIndex        =   27
         Top             =   450
         Width           =   1000
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1335
         Left            =   120
         TabIndex        =   26
         Top             =   900
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   2355
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label17 
         Caption         =   "Risk Type"
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
         Left            =   10380
         TabIndex        =   77
         Top             =   165
         Width           =   1095
      End
      Begin VB.Label Label14 
         Caption         =   "HR Location Mapped"
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
         Left            =   8055
         TabIndex        =   73
         Top             =   135
         Width           =   2055
      End
      Begin VB.Label Label11 
         Caption         =   "Earthquake Zone"
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
         Left            =   6240
         TabIndex        =   61
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label Label19 
         Caption         =   "Search By Location"
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
         Left            =   600
         TabIndex        =   47
         Top             =   120
         Width           =   2055
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
         Left            =   11910
         TabIndex        =   46
         Top             =   165
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
         Left            =   14310
         TabIndex        =   45
         Top             =   165
         Width           =   75
      End
      Begin VB.Label Label8 
         Caption         =   "State"
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
         Left            =   4320
         TabIndex        =   42
         Top             =   120
         Width           =   510
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5145
      Left            =   120
      TabIndex        =   28
      Top             =   120
      Width           =   7245
      Begin VB.CheckBox ChkSiloLocation 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5040
         TabIndex        =   4
         Top             =   1613
         Width           =   195
      End
      Begin VB.ComboBox CmbRiskType 
         Height          =   315
         ItemData        =   "Frm_Mst_Location.frx":003B
         Left            =   4995
         List            =   "Frm_Mst_Location.frx":0048
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   2445
         Width           =   1665
      End
      Begin VB.ComboBox CmbHRLocationID 
         Height          =   315
         Left            =   1800
         Sorted          =   -1  'True
         TabIndex        =   12
         Top             =   3720
         Width           =   4875
      End
      Begin VB.CommandButton CmdEQZoneID 
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
         Height          =   375
         Left            =   6720
         TabIndex        =   57
         Top             =   3270
         Width           =   375
      End
      Begin VB.TextBox TxtFeesForWithoutCM 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   4995
         MaxLength       =   10
         TabIndex        =   9
         Top             =   2887
         Width           =   1680
      End
      Begin VB.TextBox TxtLocationName 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1815
         MaxLength       =   100
         TabIndex        =   0
         Top             =   705
         Width           =   4860
      End
      Begin VB.ComboBox CmbStateID 
         Height          =   315
         Left            =   1800
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   1110
         Width           =   4875
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
         Left            =   1815
         Locked          =   -1  'True
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   4170
         Width           =   4860
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
         Left            =   1815
         Locked          =   -1  'True
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   4575
         Width           =   4860
      End
      Begin VB.TextBox TxtFeesForCM 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   8
         Top             =   2887
         Width           =   1800
      End
      Begin VB.CheckBox Chk7Days 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1800
         TabIndex        =   6
         Top             =   2520
         Width           =   195
      End
      Begin VB.CommandButton CmdEmployee 
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
         Height          =   375
         Left            =   6720
         TabIndex        =   50
         Top             =   1920
         Width           =   375
      End
      Begin VB.TextBox TxtPreFix 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   360
         Left            =   1800
         MaxLength       =   5
         TabIndex        =   3
         Top             =   1530
         Width           =   1800
      End
      Begin VB.ComboBox CmbStoreManLocation 
         Height          =   315
         Left            =   1920
         Sorted          =   -1  'True
         TabIndex        =   14
         Top             =   6120
         Visible         =   0   'False
         Width           =   4755
      End
      Begin VB.CommandButton CmdStoreManLocation 
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
         Height          =   375
         Left            =   6720
         TabIndex        =   37
         Top             =   5640
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton CmdStateID 
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
         Height          =   375
         Left            =   6720
         TabIndex        =   36
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox TxtLocationID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   360
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   240
         Width           =   1560
      End
      Begin VB.TextBox TxtStateID 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1815
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   1110
         Width           =   4860
      End
      Begin VB.TextBox TxtFlag 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         MaxLength       =   1
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   5805
         Visible         =   0   'False
         Width           =   3600
      End
      Begin VB.TextBox TxtStoreManLocation 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1800
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   5790
         Visible         =   0   'False
         Width           =   4860
      End
      Begin VB.ComboBox CmbCMPMName 
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
         TabIndex        =   5
         Top             =   1950
         Width           =   4875
      End
      Begin VB.TextBox TxtCMPMName 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1815
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1950
         Width           =   4860
      End
      Begin VB.ComboBox CmbEQZoneID 
         Height          =   315
         Left            =   1800
         Sorted          =   -1  'True
         TabIndex        =   10
         Top             =   3285
         Width           =   4875
      End
      Begin VB.TextBox TxtEQZoneID 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1815
         Locked          =   -1  'True
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   3270
         Width           =   4860
      End
      Begin VB.TextBox TxtHRLocationID 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   3720
         Width           =   4860
      End
      Begin VB.TextBox TxtRiskType 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4995
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   75
         Top             =   2445
         Width           =   1665
      End
      Begin VB.Label Label18 
         Caption         =   "Silo Location"
         Height          =   255
         Left            =   5520
         TabIndex        =   78
         Top             =   1590
         Width           =   1215
      End
      Begin VB.Label Label16 
         Caption         =   "Risk Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3720
         TabIndex        =   74
         Top             =   2475
         Width           =   1095
      End
      Begin VB.Label Label13 
         Caption         =   "HR Location"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   71
         Top             =   3765
         Width           =   1560
      End
      Begin VB.Label Label7 
         Caption         =   "Earthquake Zone"
         Height          =   180
         Left            =   240
         TabIndex        =   59
         Top             =   3330
         Width           =   1800
      End
      Begin VB.Label Label6 
         Caption         =   "Fees For Without CM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3720
         TabIndex        =   54
         Top             =   2850
         Width           =   1080
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Fees For CM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   53
         Top             =   2970
         Width           =   1080
      End
      Begin VB.Label Label4 
         Caption         =   "7 Days Collection"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   240
         TabIndex        =   52
         Top             =   2430
         Width           =   945
      End
      Begin VB.Label Label3 
         Caption         =   "AM Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   240
         TabIndex        =   51
         Top             =   1950
         Width           =   1800
      End
      Begin VB.Label Label9 
         Caption         =   "Updated :"
         Height          =   255
         Left            =   240
         TabIndex        =   49
         Top             =   4635
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "Created  :"
         Height          =   240
         Left            =   240
         TabIndex        =   48
         Top             =   4230
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Location ID"
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
         Left            =   240
         TabIndex        =   40
         Top             =   1560
         Width           =   1800
      End
      Begin VB.Label Label1 
         Caption         =   "Store Man Location "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   240
         TabIndex        =   39
         Top             =   5670
         Visible         =   0   'False
         Width           =   1560
      End
      Begin VB.Label LblLocationID 
         Caption         =   "Location Code"
         Height          =   300
         Left            =   240
         TabIndex        =   29
         Top             =   240
         Width           =   1800
      End
      Begin VB.Label LblLocationName 
         Caption         =   "Location Name"
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
         Left            =   240
         TabIndex        =   31
         Top             =   735
         Width           =   1800
      End
      Begin VB.Label LblStateID 
         Caption         =   "State Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   240
         TabIndex        =   32
         Top             =   1110
         Width           =   1800
      End
      Begin VB.Label LblFlag 
         Caption         =   "Flag"
         Enabled         =   0   'False
         Height          =   300
         Left            =   240
         TabIndex        =   34
         Top             =   5445
         Visible         =   0   'False
         Width           =   1800
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1155
      Left            =   135
      TabIndex        =   35
      Top             =   5250
      Width           =   7245
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   400
         Left            =   100
         TabIndex        =   22
         Top             =   660
         Width           =   1400
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   400
         Left            =   1500
         TabIndex        =   23
         Top             =   660
         Width           =   1400
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   400
         Left            =   2895
         TabIndex        =   24
         Top             =   660
         Width           =   1400
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   400
         Left            =   4290
         TabIndex        =   25
         Top             =   660
         Width           =   1400
      End
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   400
         Left            =   5685
         TabIndex        =   18
         Top             =   660
         Width           =   1400
      End
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New"
         Height          =   400
         Left            =   100
         TabIndex        =   19
         Top             =   240
         Width           =   1400
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   400
         Left            =   1500
         TabIndex        =   16
         Top             =   240
         Width           =   1400
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   400
         Left            =   2895
         TabIndex        =   20
         Top             =   240
         Width           =   1400
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   400
         Left            =   4290
         TabIndex        =   21
         Top             =   240
         Width           =   1400
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "Search"
         Height          =   400
         Left            =   5685
         TabIndex        =   17
         Top             =   240
         Width           =   1400
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   7440
      TabIndex        =   62
      Top             =   120
      Width           =   7695
      Begin VB.CheckBox Check1 
         Caption         =   "Select All"
         Height          =   255
         Left            =   120
         TabIndex        =   69
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton CmdAdd2List 
         Caption         =   "Add To List"
         Height          =   400
         Left            =   6120
         TabIndex        =   67
         Top             =   180
         Visible         =   0   'False
         Width           =   1400
      End
      Begin VB.CommandButton CmdRemoveFromList 
         Caption         =   "Remove From List"
         Height          =   400
         Left            =   6120
         TabIndex        =   66
         Top             =   580
         Visible         =   0   'False
         Width           =   1400
      End
      Begin VB.ComboBox CmbCommodityID 
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
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   64
         Top             =   600
         Visible         =   0   'False
         Width           =   5415
      End
      Begin VB.CommandButton CmdCommodityID 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5640
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   600
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Commodity Mapping"
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
         TabIndex        =   70
         Top             =   120
         Width           =   2130
      End
      Begin VB.Label Label12 
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
         Left            =   2220
         TabIndex        =   65
         Top             =   255
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
      Height          =   8415
      Left            =   7440
      TabIndex        =   68
      Top             =   1200
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   14843
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
Attribute VB_Name = "FrmMst_Location"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg As Variant
Dim mStateID, mStoreManLocationsID, mEmployeeID, mEQZoneID, mCommodityID As Double
Const strChecked = "þ"
Const strUnChecked = "q"
Dim mHRLocationID As Double
Dim tmp1 As String
Dim mStateIndiaName As Variant

Private Sub Check1_Click()
If Check1.Value = vbChecked Then
   For I = 1 To MSHFlexGrid2.Rows - 2
      MSHFlexGrid2.row = I
      MSHFlexGrid2.Col = 0
      MSHFlexGrid2.CellFontName = "Wingdings"
      MSHFlexGrid2.CellFontSize = 12
      MSHFlexGrid2.CellAlignment = flexAlignCenterCenter
      MSHFlexGrid2.Text = strChecked
   Next
Else
   For I = 1 To MSHFlexGrid2.Rows - 2
      MSHFlexGrid2.row = I
      MSHFlexGrid2.Col = 0
      MSHFlexGrid2.CellFontName = "Wingdings"
      MSHFlexGrid2.CellFontSize = 12
      MSHFlexGrid2.CellAlignment = flexAlignCenterCenter
      MSHFlexGrid2.Text = strUnChecked
   Next
End If
End Sub


Private Sub CmbCommodityID_GotFocus()
tmp = CmbCommodityID.Text
Call TableMst_Commodity
CmbCommodityID.Clear
If RsMst_Commodity.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Commodity.RecordCount
    CmbCommodityID.AddItem RsMst_Commodity!Commodity
    RsMst_Commodity.MoveNext
Next
CmbCommodityID.Text = tmp
End Sub

Private Sub CmbCommodityID_LostFocus()
If CmbCommodityID.Text = "" Then Exit Sub

RsMst_Commodity.MoveFirst
RsMst_Commodity.Find "Commodity = '" & CmbCommodityID.Text & "'"
If RsMst_Commodity.EOF Then
   MsgBox "Invalid selection "
   CmbCommodityID.SetFocus
   Exit Sub
End If

mCommodityID = RsMst_Commodity!CommodityID
End Sub


Private Sub CmbRiskType_LostFocus()
If CmbRiskType.Text = "" Then Exit Sub

If LCase(CmbRiskType.Text) <> "low" And LCase(CmbRiskType.Text) <> "high" And LCase(CmbRiskType.Text) <> "medium" Then
   MsgBox "Invalid selection!!! "
   CmbRiskType.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmbSRiskType_LostFocus()
If CmbSRiskType.Text = "" Then Exit Sub

If LCase(CmbSRiskType.Text) <> "low" And LCase(CmbSRiskType.Text) <> "high" And LCase(CmbSRiskType.Text) <> "medium" Then
   MsgBox "Invalid selection!!! "
   CmbSRiskType.SetFocus
   Exit Sub
End If
End Sub


Private Sub CmbStoreManLocation_GotFocus()
tmp = CmbStoreManLocation.Text
Call TableMst_StoreManLocations
CmbStoreManLocation.Clear
If RsMst_StoreManLocations.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_StoreManLocations.RecordCount
    CmbStoreManLocation.AddItem RsMst_StoreManLocations!StoreManLocations
    RsMst_StoreManLocations.MoveNext
Next
CmbStoreManLocation.Text = tmp
End Sub
'----
Private Sub CmbStoreManLocation_LostFocus()
If CmbStoreManLocation.Text = "" Then
   Exit Sub
End If
RsMst_StoreManLocations.MoveFirst
RsMst_StoreManLocations.Find "StoreManLocations = '" & CmbStoreManLocation.Text & "'"
If RsMst_StoreManLocations.EOF Then
   MsgBox "Invalid selection "
   CmbStoreManLocation.SetFocus
   Exit Sub
End If
mStoreManLocationsID = RsMst_StoreManLocations!StoreManLocationsID
TxtPreFix = RsMst_StoreManLocations!SM_Prefix
End Sub

Private Sub CmdAdd2List_Click()
If CmbCommodityID.Text = "" Then
    MsgBox "Blank Commodity not Allowed!!!"
    CmbCommodityID.SetFocus
    Exit Sub
End If
For I = 1 To MSHFlexGrid2.Rows - 1
    If CmbCommodityID.Text = MSHFlexGrid2.TextMatrix(I, 1) Then
        MsgBox "Commodity Already Exist!!!"
        Exit Sub
    End If
Next
r = MSHFlexGrid2.Rows - 1
MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1

MSHFlexGrid2.TextMatrix(r, 0) = mCommodityID
MSHFlexGrid2.TextMatrix(r, 1) = CmbCommodityID.Text

Call ClearCommodityControl
End Sub

Private Sub ClearCommodityControl()
CmdAdd2List.Enabled = True
CmdRemoveFromList.Enabled = False
CmbCommodityID.Enabled = True
CmdCommodityID.Enabled = True
CmbCommodityID.Text = ""
End Sub

Private Sub CmdCommodityID_Click()
FrmMst_Commodity.Show
End Sub

Private Sub CmdEmployee_Click()
FrmMst_Employee.Show
End Sub

Private Sub CmdEQZoneID_Click()
FrmMst_EarthquakeZone.Show
End Sub

Private Sub CmdGo_Click()

Call Grid_Head

tmp = "Select ML.LocationID,ML.LocationName,MS.StateName,ML.Flag, ME.EmployeeName,ML.Collect7Days, ML.FeesForCM, ML.FeesForWithoutCM,MEQ.EQZone,ML.HRLocationID, ML.RiskType,ML.NAVCode,ML.FAGUpdated,ML.FAGUpdatedDate,ML.SiloLocation " & _
      " From Mst_Location ML " & _
      " Inner Join Mst_State MS On ML.StateID = MS.StateID " & _
      " Left Join Mst_Employee ME On ML.AM_EmployeeID = ME.EmployeeID " & _
      " Left Join Mst_EarthquakeZone MEQ On MEQ.EQZoneID = ML.EQZoneID "

Dim wc As Variant
wc = ""
If TxtSearchLocation <> "" Then
   wc = " Where ML.LocationName Like '%" & TxtSearchLocation & "%' "
End If

If CmbSState.Text <> "" Then
   If wc = "" Then
      wc = " Where MS.StateName = '" & CmbSState.Text & "'"
   Else
      wc = wc & " And MS.StateName = '" & CmbSState.Text & "'"
   End If
End If
If CmbSEQZoneID.Text <> "" Then
   If wc = "" Then
      wc = " Where MEQ.EQZone = '" & CmbSEQZoneID.Text & "'"
   Else
      wc = wc & " And MEQ.EQZone = '" & CmbSEQZoneID.Text & "'"
   End If
End If

If CmbSRiskType.Text <> "" Then
   If wc = "" Then
      wc = " Where ML.RiskType = '" & CmbSRiskType.Text & "'"
   Else
      wc = wc & " And ML.RiskType = '" & CmbSRiskType.Text & "'"
   End If
End If

If LCase(CmbSHRLocMapped.Text) = "yes" Then
   If wc = "" Then
      wc = " Where ML.HRLocationID >0 "
   Else
      wc = wc & " And ML.HRLocationID>0"
   End If
ElseIf LCase(CmbSHRLocMapped.Text) = "no" Then
   If wc = "" Then
      wc = " Where isnull(ML.HRLocationID,0)=0 "
   Else
      wc = wc & " And isnull(ML.HRLocationID,0)=0 "
   End If
End If

tmp = tmp & wc & " Order By MS.StateName,ML.LocationName "

Call TmpTable

For iu = 1 To TmpRs.RecordCount
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
    For C = 0 To MSHFlexGrid1.Cols - 1
        MSHFlexGrid1.TextMatrix(iu, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
    Next
    If MSHFlexGrid1.TextMatrix(iu, 5) = "True" Then
       MSHFlexGrid1.TextMatrix(iu, 5) = "Yes"
    Else
       MSHFlexGrid1.TextMatrix(iu, 5) = "No"
    End If
    MSHFlexGrid1.TextMatrix(iu, 9) = GetHRLocation(Val(MSHFlexGrid1.TextMatrix(iu, 9)))
    MSHFlexGrid1.TextMatrix(iu, 14) = IIf(MSHFlexGrid1.TextMatrix(iu, 14) = "", "No", IIf(MSHFlexGrid1.TextMatrix(iu, 14) = "False", "No", "Yes"))
'    If MSHFlexGrid1.TextMatrix(iu, 6) = "" Then
'       MSHFlexGrid1.TextMatrix(iu, 6) = "500"
'    End If
'    If MSHFlexGrid1.TextMatrix(iu, 7) = "" Then
'       MSHFlexGrid1.TextMatrix(iu, 7) = "1000"
'    End If
    TmpRs.MoveNext
Next
Label22.Refresh
Label22.Caption = TmpRs.RecordCount
End Sub

Private Sub CmdRemoveFromList_Click()
I = MSHFlexGrid2.RowSel
If MSHFlexGrid2.TextMatrix(I, 1) <> "" Then
   For RO = I To MSHFlexGrid2.Rows - 2
       For C = 0 To MSHFlexGrid2.Cols - 1
            MSHFlexGrid2.TextMatrix(RO, C) = MSHFlexGrid2.TextMatrix(RO + 1, C)
       Next
   Next
   MSHFlexGrid2.Rows = MSHFlexGrid2.Rows - 1
End If
Call ClearCommodityControl
End Sub

Private Sub CmdStoreManLocation_Click()
FrmMst_StoreManLocations.Show
End Sub

Private Sub CmbCMPMName_GotFocus()
tmp = CmbCMPMName.Text
Call TableMst_Employee(" Where DesignationID=3 And Flag = 'Y' ")
CmbCMPMName.Clear
If RsMst_Employee.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Employee.RecordCount
    CmbCMPMName.AddItem RsMst_Employee!EmployeeName
    RsMst_Employee.MoveNext
Next
CmbCMPMName.Text = tmp

End Sub
'----
Private Sub CmbCMPMName_LostFocus()
Call Grid_Head
If CmbCMPMName.Text = "" Then
   Exit Sub
End If
RsMst_Employee.MoveFirst
RsMst_Employee.Find "EmployeeName = '" & CmbCMPMName.Text & "'"
If RsMst_Employee.EOF Then
   MsgBox "Invalid selection "
   CmbCMPMName.SetFocus
   Exit Sub
End If
mEmployeeID = RsMst_Employee!EmployeeID
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
TxtSearchLocation = ""
CmbSState.Text = ""
Label22.Caption = ""
CmbSEQZoneID.Text = ""
CmbSHRLocMapped = ""


Call TableMst_Location(" Order By StateID,LocationName")
'Call TableMst_StoreManLocations
If RsMst_Location.RecordCount = 0 Then
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
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Mst_Location.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Mst_Location.xls")
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
           oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
        End If
    Next
Next
oWB.Save
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub

Private Sub MSHFlexGrid2_Click()
If SaveCmd.Enabled = False Then Exit Sub
'If Edit_Flg = "E" And LCase(Gen_UserRole) <> "power" Then Exit Sub
'Call ClearCommodityControl
'If MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 0) = "" Then Exit Sub
'CmbCommodityID.Text = MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 1)
'CmbCommodityID.Enabled = False
'CmdCommodityID.Enabled = False
'CmdAdd2List.Enabled = False
'CmdRemoveFromList.Enabled = True
With MSHFlexGrid2
    If .Col = 0 Then  '.Row = 1 And
      If .TextMatrix(.row, .Col) = "" Then Exit Sub
      .Col = 0
         If .TextMatrix(.row, .Col) = strUnChecked Then
            .TextMatrix(.row, .Col) = strChecked
         Else
         .TextMatrix(.row, .Col) = strUnChecked
      End If
   End If
End With

End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_Location.MoveNext
If RsMst_Location.EOF Then
   RsMst_Location.MoveLast
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
RsMst_Location.MovePrevious
If RsMst_Location.BOF Then
   RsMst_Location.MoveFirst
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
RsMst_Location.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsMst_Location.MoveLast
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
   If RsMst_Location.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   'Call FirstCmd_Click
   Call Indata
   Exit Sub
End If
Frame0.Height = Me.Height - 800
Frame0.Left = Frame1.Left
Frame0.Top = Frame1.Top
Frame0.Width = Me.Width - 500
Frame0.Visible = True
MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 1000
RsMst_Location.MoveFirst

''===========comment on 7 july 2012 for not clearing search grid=============
'MSHFlexGrid1.Clear
'MSHFlexGrid1.Cols = 3
'Call Grid_Head
'TxtSearchLocation = ""
'CmbSState.Text = ""
'Label22.Caption = ""
'CmbSEQZoneID.Text = ""
'CmbSHRLocMapped = ""
'''===========comment on 7 july 2012 for not clearing search grid=============

End Sub
Public Sub Indata()
Frame2.Enabled = False
Call ClearCommodityControl
Call Grid_Head1

CmbEQZoneID.Visible = False
TxtEQZoneID.Visible = True
CmdEQZoneID.Visible = False

TxtLocationID.Locked = True
TxtLocationName.Locked = True
TxtStateID.Locked = True
TxtFeesForCM.Locked = True
TxtFeesForWithoutCM.Locked = True

CmbStateID.Visible = False
CmdStateID.Visible = False
TxtStateID.Visible = True

CmbHRLocationID.Visible = False
TxtHRLocationID.Visible = True

CmbCMPMName.Visible = False
CmdEmployee.Visible = False
TxtCMPMName.Visible = True

CmbRiskType.Visible = False
TxtRiskType.Visible = True
TxtRiskType.Locked = True

'TxtStoreManLocation.Locked = True
'CmbStoreManLocation.Visible = False
'CmdStoreManLocation.Visible = False
'TxtStoreManLocation.Visible = True

TxtFlag.Locked = True

TxtLocationID = IIf(IsNull(RsMst_Location!LocationID), "", RsMst_Location!LocationID)
TxtLocationName = IIf(IsNull(RsMst_Location!LocationName), "", RsMst_Location!LocationName)
'TxtStateID = IIf(IsNull(RsMst_Location!StateID), "", RsMst_Location!StateID)
mStateID = IIf(IsNull(RsMst_Location!StateID), "", RsMst_Location!StateID)
mStoreManLocationsID = IIf(IsNull(RsMst_Location!StoreManLocationsID), "", RsMst_Location!StoreManLocationsID)

TxtStateID = GetStateName(RsMst_Location!StateID)
'TxtStoreManLocation = GetStoreManLocationsName(RsMst_Location!StoreManLocationsID, "N")
TxtPreFix = IIf(IsNull(RsMst_Location!SM_Prefix), "", RsMst_Location!SM_Prefix) '= TxtPreFix.Text GetStoreManLocationsName(RsMst_Location!StoreManLocationsID, "N")
mStateIndiaName = GetStateIndiaFromStateName(TxtStateID)

If RsMst_Location!Collect7Days = False Then
   Chk7Days.Value = vbUnchecked 'IIf(IsNull(RsMst_Location!Collect7Days), False, RsMst_Location!Collect7Days)
Else
   Chk7Days.Value = vbChecked
End If
TxtFeesForCM = IIf(IsNull(RsMst_Location!FeesForCM), 500, RsMst_Location!FeesForCM)
TxtFeesForWithoutCM = IIf(IsNull(RsMst_Location!FeesForWithoutCM), 1000, RsMst_Location!FeesForWithoutCM)

mEQZoneID = IIf(IsNull(RsMst_Location!EQZoneID), 0, RsMst_Location!EQZoneID)
TxtEQZoneID = GetEarthequakeZone(Val(mEQZoneID))

TxtRiskType = IIf(IsNull(RsMst_Location!RiskType), "", RsMst_Location!RiskType)

TxtCreated = IIf(IsNull(RsMst_Location!CreatedBy), "", RsMst_Location!CreatedBy) & " :- " & IIf(IsNull(RsMst_Location!CreatedDate), "", RsMst_Location!CreatedDate)
TxtUpdated = IIf(IsNull(RsMst_Location!UpdatedBy), "", RsMst_Location!UpdatedBy) & " :- " & IIf(IsNull(RsMst_Location!UpdatedDate), "", RsMst_Location!UpdatedDate)

mEmployeeID = IIf(IsNull(RsMst_Location!AM_EmployeeID), 0, RsMst_Location!AM_EmployeeID)
TxtCMPMName = GetEmployeeName(mEmployeeID)
TxtFlag = IIf(IsNull(RsMst_Location!Flag), "", RsMst_Location!Flag)

mHRLocationID = IIf(IsNull(RsMst_Location!HRLocationID), 0, RsMst_Location!HRLocationID)
TxtHRLocationID = GetHRLocation(mHRLocationID)

Chk7Days.Enabled = False
''-- Ullhas Patil 21 April 2018
ChkSiloLocation.Enabled = False
ChkSiloLocation.Value = IIf(IsNull(RsMst_Location!SiloLocation), vbUnchecked, IIf(RsMst_Location!SiloLocation = False, vbUnchecked, vbChecked))
''-----------------------------
Call IndataCommodityMapping

If RsMst_Location.RecordCount = 1 Then
   Call GButtonLock_1(Me, True)
Else
   Call GButtonLock(Me, True)
End If

End Sub

Private Sub IndataCommodityMapping()
Call GetAllCommodity
'tmp = " Select MLCM.LocationCommodityMappingID, MCom.CommodityID , MCom.Commodity"
'tmp = tmp & " From Mst_LocationCOmmodityMapping MLCM"
'tmp = tmp & " Inner Join Mst_Commodity MCom on MLCM.CommodityID = MCom.CommodityID"
'tmp = tmp & " Where MLCM.LocationID = " & Val(TxtLocationID)
'Call TmpTable
'
'If TmpRs.RecordCount > 0 Then
'    MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + TmpRs.RecordCount
'    For I = 1 To TmpRs.RecordCount
'        MSHFlexGrid2.TextMatrix(I, 0) = IIf(IsNull(TmpRs!CommodityID), "", TmpRs!CommodityID)
'        MSHFlexGrid2.TextMatrix(I, 1) = IIf(IsNull(TmpRs!Commodity), "", TmpRs!Commodity)
'        MSHFlexGrid2.TextMatrix(I, 2) = IIf(IsNull(TmpRs!LocationCommodityMappingID), "", TmpRs!LocationCommodityMappingID)
'        TmpRs.MoveNext
'    Next
'End If

End Sub

Private Sub TxtFeesForCM_LostFocus()
If TxtFeesForCM = "" Then Exit Sub
If IsNumeric(TxtFeesForCM) = False Or Val(TxtFeesForCM) < 0 Then
   MsgBox "Invalid numeric format !!!!"
   TxtFeesForCM.SetFocus
   Exit Sub
End If
TxtFeesForCM = Val(TxtFeesForCM)
End Sub

Private Sub TxtFeesForWithoutCM_LostFocus()
If TxtFeesForWithoutCM = "" Then Exit Sub
If IsNumeric(TxtFeesForWithoutCM) = False Or Val(TxtFeesForWithoutCM) < 0 Then
   MsgBox "Invalid numeric format !!!!"
   TxtFeesForWithoutCM.SetFocus
   Exit Sub
End If
TxtFeesForWithoutCM = Val(TxtFeesForWithoutCM)
End Sub

Private Sub TxtLocationID_LostFocus()
If TxtLocationID = "" Then Exit Sub
If IsNumeric(TxtLocationID) = False Then
   MsgBox "Invalid numeric format !!!!"
   TxtLocationID.SetFocus
   Exit Sub
End If
TxtLocationID = Val(TxtLocationID)
End Sub
Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)

TxtLocationID.Locked = True
TxtLocationName.Locked = False

TxtStateID.Locked = False
TxtFeesForCM.Locked = False
TxtFeesForWithoutCM.Locked = False

TxtFeesForCM.Locked = False
TxtFeesForWithoutCM.Locked = False

CmbStateID.Visible = True
CmdStateID.Visible = True
TxtStateID.Visible = False

CmbEQZoneID.Visible = True
TxtEQZoneID.Visible = False
CmdEQZoneID.Visible = True

TxtHRLocationID.Visible = False
CmbHRLocationID.Visible = True
CmbHRLocationID.Text = ""

CmbCMPMName.Visible = True
CmdEmployee.Visible = True
TxtCMPMName.Visible = False
Chk7Days.Enabled = True

CmbRiskType.Visible = True
TxtRiskType.Visible = False
CmbRiskType.Text = ""

'TxtStoreManLocation.Locked = False
'CmbStoreManLocation.Visible = True
'CmdStoreManLocation.Visible = True
TxtStoreManLocation.Visible = False

TxtFlag.Locked = False
TxtPreFix.Locked = False
TxtPreFix.Enabled = True

TxtLocationID = ""
TxtLocationName = ""
TxtFeesForCM = ""
TxtFeesForWithoutCM = ""
TxtStateID = ""
TxtFlag = ""
TxtPreFix = ""
CmbStateID.Text = ""
CmbEQZoneID.Text = ""
CmbStoreManLocation.Text = ""
CmbRiskType.Text = ""
Chk7Days.Value = 0

''-- Ullhas Patil 21 April 2018
ChkSiloLocation.Value = False
ChkSiloLocation.Enabled = True
''--------------------------------

TxtCreated = ""
TxtUpdated = ""

Call ClearCommodityControl
Call Grid_Head1
'''''''''''''''''''''''''''''
Call GetAllCommodity
'''''''''''''''''''''''''''''
Frame2.Enabled = True

End Sub

Private Sub GetAllCommodity()
tmp = " Select MCom.CommodityID , MCom.Commodity "
tmp = tmp & " From Mst_Commodity MCom"
'tmp = tmp & " Left Join Mst_LocationCOmmodityMapping MLCM on MLCM.LocationID = ML.LocationID"
tmp = tmp & " Order by Commodity"
Call TmpTable

tmp = " Select MCom.CommodityID , MCom.Commodity, MLCM.LocationID, MLCM.LocationCommodityMappingID "
tmp = tmp & " From Mst_Commodity MCom"
tmp = tmp & " Inner Join Mst_LocationCOmmodityMapping MLCM on MLCM.CommodityID = MCom.CommodityID"
'tmp = tmp & " Where ML.LocationID not in (Select LocationID from Mst_LocationCOmmodityMapping MLCM Where CommodityID =  )"
tmp = tmp & " Where MLCM.LocationID = " & Val(TxtLocationID)
tmp = tmp & " Order by Commodity"
Call Tmp1Table


If TmpRs.RecordCount > 0 Then
    MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + TmpRs.RecordCount
    For I = 1 To TmpRs.RecordCount
        MSHFlexGrid2.TextMatrix(I, 1) = IIf(IsNull(TmpRs!Commodity), "", TmpRs!Commodity)
        MSHFlexGrid2.TextMatrix(I, 2) = IIf(IsNull(TmpRs!CommodityID), "", TmpRs!CommodityID)
        TmpRs1.Filter = " CommodityID = " & TmpRs!CommodityID & " And LocationID = " & Val(TxtLocationID) & " "
        If TmpRs1.EOF = True Then
            MSHFlexGrid2.row = I
            MSHFlexGrid2.Col = 0
            MSHFlexGrid2.CellFontName = "Wingdings"
            MSHFlexGrid2.CellFontSize = 12
            MSHFlexGrid2.CellAlignment = flexAlignCenterCenter
            MSHFlexGrid2.Text = strUnChecked
        Else
            MSHFlexGrid2.TextMatrix(I, 3) = TmpRs1!LocationCommodityMappingID
            MSHFlexGrid2.row = I
            MSHFlexGrid2.Col = 0
            MSHFlexGrid2.CellFontName = "Wingdings"
            MSHFlexGrid2.CellFontSize = 12
            MSHFlexGrid2.CellAlignment = flexAlignCenterCenter
            MSHFlexGrid2.Text = strChecked
        End If
            TmpRs.MoveNext
    Next
End If
End Sub

Private Sub EditCmd_Click()
Edit_Flg = "E"
Frame2.Enabled = True

If Val(TxtLocationID) = 0 Then
   MsgBox " you Can't Edit this Record"
   Exit Sub
End If
Call GButtonLock(Me, False)

CmbCMPMName.Text = TxtCMPMName

'TxtStoreManLocation.Locked = False
'CmbStoreManLocation.Visible = True
'CmdStoreManLocation.Visible = True
'TxtStoreManLocation.Visible = False
CmbStateID.Text = TxtStateID
CmbEQZoneID.Text = TxtEQZoneID
CmbStoreManLocation.Text = TxtStoreManLocation

CmbCMPMName.Visible = True
CmdEmployee.Visible = True
TxtCMPMName.Visible = False

CmbRiskType.Visible = True
TxtRiskType.Visible = False
CmbRiskType.Text = TxtRiskType

TxtHRLocationID.Visible = False
CmbHRLocationID.Visible = True
CmbHRLocationID.Text = TxtHRLocationID

If LCase(Gen_UserRole) = "power" Then
   TxtLocationID.Locked = True
   TxtLocationName.Locked = True
   TxtStateID.Locked = False
   TxtFeesForCM.Locked = False
   TxtFeesForWithoutCM.Locked = False
   
   CmbStateID.Visible = True
   CmdStateID.Visible = True
   TxtStateID.Visible = False
   
   CmbEQZoneID.Visible = True
   TxtEQZoneID.Visible = False
   CmdEQZoneID.Visible = True
   
   TxtFlag.Locked = False
   Chk7Days.Enabled = True
End If

''-- Ullhas Patil 21 April 2018
If ChkSiloLocation.Value = True Then
    ChkSiloLocation.Enabled = True
Else
    ChkSiloLocation.Enabled = False
    If LCase(Gen_UserRole) = "power" Then
        ChkSiloLocation.Enabled = True
    End If
End If
''-----------------------------
End Sub
Private Function CheckForUseage() As Boolean
Dim Retval As Boolean

Retval = True
tmp = "select 'Mst_License' As  'Type', Mst_License.LocationID  from Mst_License where Mst_License.locationID = " & Val(TxtLocationID)
If Gen_DBType = "MDB" Then
   Call TmpTable
   If TmpRs.RecordCount > 0 Then
      Retval = False
   End If
Else
   Call Tmp4Table
   If TmpRs4.RecordCount > 0 Then
      Retval = False
   End If
End If
CheckForUseage = Retval
End Function

Private Sub DeleteCmd_Click()
Dim ans As Variant
On Error GoTo msgError

If CheckForUseage() = False Then
   MsgBox "You can't delete this record since it is used in Transaction !!!  "
   Exit Sub
End If


If Val(TxtLocationID) = 0 Then
   MsgBox " you Can't Delete this Record"
   Exit Sub
End If
ans = MsgBox("Do you really want to DELETE this record???", vbYesNo)
If ans = vbYes Then

   tmp = " Delete From Mst_LocationCommodityMapping Where LocationID = " & Val(TxtLocationID.Text)
   Call TmpTable

   RsMst_Location.Delete
   RsMst_Location.Update
   If RsMst_Location.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_Location.MoveNext
   If RsMst_Location.EOF() Then
      RsMst_Location.MoveLast
   End If
   Call Indata
End If
Exit Sub
msgError:
MsgBox "Child record Present "
RsMst_Location.CancelUpdate
End Sub
Private Sub SaveCmd_Click()
If TxtLocationName.Text = "" Then
   MsgBox "Blank LocationName Not Allowed !!! "
   TxtLocationName.SetFocus
   Exit Sub
End If

If CmbStateID.Text = "" Then
   MsgBox "Blank StateID Not Allowed !!! "
   CmbStateID.SetFocus
   Exit Sub
End If
If CmbCMPMName.Text = "" Then
   MsgBox "Blank AM Name Not Allowed !!! "
   CmbCMPMName.SetFocus
   Exit Sub
End If

If TxtPreFix.Text = "" Then
   MsgBox "Blank Location ID Not Allowed !!! "
   CmbStateID.SetFocus
   Exit Sub
End If

If TxtFeesForCM.Text = "" Then
   MsgBox "Blank Fees For CM Not Allowed !!! "
   TxtFeesForCM.SetFocus
   Exit Sub
End If

If TxtFeesForWithoutCM.Text = "" Then
   MsgBox "Blank Fees For Without CM Not Allowed !!! "
   TxtFeesForWithoutCM.SetFocus
   Exit Sub
End If

If CmbHRLocationID.Text = "" Then
   MsgBox "Blank HR Location not Allowed"
   CmbHRLocationID.SetFocus
   Exit Sub
End If

If CmbRiskType.Text = "" Then
   MsgBox "Blank Risk Type not Allowed"
   CmbRiskType.SetFocus
   Exit Sub
End If


'If CmbEQZoneID.Text = "" Then
'   MsgBox "Blank Earthquake Zone Not Allowed !!! "
'   CmbEQZoneID.SetFocus
'   Exit Sub
'End If

If Edit_Flg = "A" Then
   If RsMst_Location.RecordCount > 0 Then
      RsMst_Location.MoveFirst
      RsMst_Location.Find "LocationName= '" & TxtLocationName.Text & "'"
      If Not RsMst_Location.EOF Then
        MsgBox "Duplicate Location Name Not Allowed !!! "
         TxtLocationName.SetFocus
         Exit Sub
      End If
      RsMst_Location.MoveFirst
      RsMst_Location.Find "SM_Prefix= '" & TxtPreFix.Text & "'"
      If Not RsMst_Location.EOF Then
        MsgBox "Duplicate Location ID Not Allowed !!! "
         TxtPreFix.SetFocus
         Exit Sub
      End If
   End If
'   RsMst_StoreManLocations.AddNew
   RsMst_Location.AddNew
'   RsMst_Location!LocationID = GetLocationID
   RsMst_Location!G_UID = GetGUID
   mG_UID = RsMst_Location!G_UID
'   RsMst_StoreManLocations!G_UID = GetGUID
'   TxtLocationID.Text = RsMst_Location!LocationID
Else
    RsMst_Location.MoveFirst
    RsMst_Location.Find "LocationName= '" & TxtLocationName.Text & "'"
    If Not RsMst_Location.EOF Then
       If RsMst_Location!LocationID <> TxtLocationID.Text Then
          MsgBox "Duplicate Location Name Not Allowed !!! "
          TxtLocationName.SetFocus
          Exit Sub
       End If
    End If
    RsMst_Location.MoveFirst
    RsMst_Location.Find "SM_Prefix= '" & TxtPreFix.Text & "'"
    If Not RsMst_Location.EOF Then
       If RsMst_Location!LocationID <> TxtLocationID.Text Then
          MsgBox "Duplicate Location Name Not Allowed !!! "
          TxtLocationName.SetFocus
          Exit Sub
       End If
    End If
    RsMst_Location.MoveFirst
    RsMst_Location.Find " LocationID= " & TxtLocationID.Text
End If

'RsMst_StoreManLocations!StoreManLocationsID = TxtLocationID.Text
'RsMst_StoreManLocations!StoreManLocations = TxtLocationName.Text
'RsMst_StoreManLocations!Address = ""
'RsMst_StoreManLocations!SM_Prefix = TxtPreFix.Text
'RsMst_StoreManLocations!Flag = "C"
'RsMst_StoreManLocations.Update

'RsMst_Location!LocationID = TxtLocationID.Text
RsMst_Location!LocationName = TxtLocationName.Text
RsMst_Location!StateID = mStateID
RsMst_Location!StoreManLocationsID = Val(TxtLocationID)
RsMst_Location!SM_Prefix = TxtPreFix.Text
RsMst_Location!AM_EmployeeID = mEmployeeID
RsMst_Location!Collect7Days = Chk7Days.Value
RsMst_Location!FeesForCM = TxtFeesForCM.Text
RsMst_Location!FeesForWithoutCM = TxtFeesForWithoutCM.Text
RsMst_Location!HRLocationID = mHRLocationID
RsMst_Location!RiskType = Trim(CmbRiskType.Text)
If CmbEQZoneID.Text <> "" Then
   RsMst_Location!EQZoneID = mEQZoneID
Else
   RsMst_Location!EQZoneID = Null
End If
RsMst_Location!Flag = "A" 'TxtFlag.Text

''-- Ullhas Patil 21 APril 2018
RsMst_Location!SiloLocation = IIf(ChkSiloLocation.Value = vbChecked, True, False)
''-----------------------------
If IsNull(RsMst_Location!CreatedBy) = True Or RsMst_Location!CreatedBy = "" Then
   RsMst_Location!CreatedBy = Gen_UserLoginID
   RsMst_Location!CreatedDate = Now
Else
   RsMst_Location!UpdatedBy = Gen_UserLoginID
   RsMst_Location!UpdatedDate = Now
End If

RsMst_Location.Update
tmp = "Select * from Mst_Location Where G_UID = '" & mG_UID & "'"
Call TmpTable
If Edit_Flg = "A" Then
   TxtLocationID = IIf(IsNull(TmpRs!LocationID), 0, TmpRs!LocationID)
End If
Call SaveCommodityMapping
Call Indata
End Sub

Private Sub SaveCommodityMapping()

tmp = " Delete From Mst_LocationCommodityMapping Where LocationID = " & Val(TxtLocationID.Text)
Call TmpTable

tmp = "Select * From Mst_LocationCommodityMapping Where LocationID = " & Val(TxtLocationID.Text)
Call Tmp1Table

For I = 1 To MSHFlexGrid2.Rows - 2
   If MSHFlexGrid2.TextMatrix(I, 0) = strChecked Then
      TmpRs1.AddNew
'      TmpRs1!LocationCommodityMappingID = GetCommodityMappingID
      TmpRs1!G_UID = GetGUID
      TmpRs1!CommodityID = Val(MSHFlexGrid2.TextMatrix(I, 2))
      TmpRs1!LocationID = Val(TxtLocationID.Text)
      If IsNull(TmpRs1!CreatedBy) = True Or TmpRs1!CreatedBy = "" Then
         TmpRs1!CreatedBy = Gen_UserLoginID
         TmpRs1!CreatedDate = Now
      Else
         TmpRs1!UpdatedBy = Gen_UserLoginID
         TmpRs1!UpdatedDate = Now
      End If
      TmpRs1.Update
      TmpRs1.Requery
   End If
Next


'If MSHFlexGrid2.Rows = 2 Then Exit Sub
'
'tmp = "Select * From Mst_LocationCommodityMapping Where LocationID = " & Val(TxtLocationID.Text)
'Call Tmp1Table
'tmp1 = " (0, "
'For I = 1 To MSHFlexGrid2.Rows - 2
'   If MSHFlexGrid2.TextMatrix(I, 2) = "" Then
'      TmpRs1.AddNew
'      TmpRs1!LocationCommodityMappingID = GetCommodityMappingID
'      tmp1 = tmp1 & TmpRs1!LocationCommodityMappingID & " ,"
'      TmpRs1!G_UID = GetGUID
'   Else
'      TmpRs1.MoveFirst
'      TmpRs1.Find " LocationCommodityMappingID= " & Val(MSHFlexGrid2.TextMatrix(I, 2))
'      tmp1 = tmp1 & Val(MSHFlexGrid2.TextMatrix(I, 2)) & " ,"
'   End If
'   TmpRs1!CommodityID = Val(MSHFlexGrid2.TextMatrix(I, 0))
'   TmpRs1!LocationID = Val(TxtLocationID.Text)
'   If IsNull(TmpRs1!CreatedBy) = True Or TmpRs1!CreatedBy = "" Then
'      TmpRs1!CreatedBy = Gen_UserLoginID
'      TmpRs1!CreatedDate = Now
'   Else
'      TmpRs1!UpdatedBy = Gen_UserLoginID
'      TmpRs1!UpdatedDate = Now
'   End If
'   TmpRs1.Update
'   TmpRs1.Requery
'Next
'
'tmp1 = Mid(tmp1, 1, Len(tmp1) - 1) & " )"
'tmp = " Delete From Mst_LocationCommodityMapping Where LocationID = " & Val(TxtLocationID.Text) & " And LocationCommodityMappingID not in " & tmp1 & ""
'Call TmpTable

End Sub

Private Function GetCommodityMappingID() As Double
Dim Retval As Double
tmp = "Select max(LocationCommodityMappingID) from Mst_LocationCOmmodityMapping"
Call TmpTable

Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetCommodityMappingID = Retval
End Function


Private Function GetLocationID() As Double
Dim Retval As Double
tmp = "Select max(LocationID) from Mst_Location"
Call TmpTable

Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))

If Retval = 0 Then
   RsMst_Location!G_UID = GetGUID
''   RsMst_Location!LocationId = 0
''   RsMst_Location!Locationname = "ALL"
''   RsMst_Location!StateID = 1
''   RsMst_Location!StoreManLocationsID = 1
''   RsMst_Location!Flag = "0" 'TxtFlag.Text
''   RsMst_Location.Update
''   RsMst_Location.AddNew
   
End If
Retval = Retval + 1
GetLocationID = Retval

End Function

Private Sub Grid_Head1()
'With MSHFlexGrid2
'    .Clear
'    .Rows = 2
'    .Cols = 3
'
'    .TextMatrix(0, 0) = "Commodity ID"
'    .TextMatrix(0, 1) = "Commodity Name"
'    .TextMatrix(0, 2) = "Location Commodity Mapping ID"
'    .ColWidth(0) = 1200
'    .ColWidth(1) = 6200
'    .ColWidth(2) = 0
'End With
Check1.Value = vbUnchecked
With MSHFlexGrid2
    .Clear
    .Rows = 2
    .Cols = 4
    .FixedCols = 0

    .TextMatrix(0, 0) = "Select"
    .TextMatrix(0, 1) = "Commodity"
    .TextMatrix(0, 2) = "Commodity ID"
    .TextMatrix(0, 3) = "Location Commodity Mapping ID"
    
    .ColWidth(0) = 700
    .ColWidth(1) = 4500
    .ColWidth(2) = 1200
    .ColWidth(3) = 0
End With

End Sub

Private Sub Grid_Head()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 15 ''14
MSHFlexGrid1.RowHeight(0) = 500
MSHFlexGrid1.WordWrap = True
MSHFlexGrid1.TextMatrix(0, 0) = "Location ID"
MSHFlexGrid1.TextMatrix(0, 1) = "Location Name"
MSHFlexGrid1.TextMatrix(0, 2) = "State Name"
MSHFlexGrid1.TextMatrix(0, 3) = "Flag"
MSHFlexGrid1.TextMatrix(0, 4) = "AM Name "
MSHFlexGrid1.TextMatrix(0, 5) = "Collect 7 Days"
MSHFlexGrid1.TextMatrix(0, 6) = "Fees For CM"
MSHFlexGrid1.TextMatrix(0, 7) = "Fees For Wihtout CM"
MSHFlexGrid1.TextMatrix(0, 8) = "Earthquake Zone"
MSHFlexGrid1.TextMatrix(0, 9) = "HR Location"
MSHFlexGrid1.TextMatrix(0, 10) = "Risk Type"
MSHFlexGrid1.TextMatrix(0, 11) = "NAV Code"
MSHFlexGrid1.TextMatrix(0, 12) = "NAV Code updated by"
MSHFlexGrid1.TextMatrix(0, 13) = "NAV Code updated on"
MSHFlexGrid1.TextMatrix(0, 14) = "Silo Location"    ''-- Ullhas Patil 21 April 2018

MSHFlexGrid1.ColWidth(0) = 800
MSHFlexGrid1.ColWidth(1) = 2500
MSHFlexGrid1.ColWidth(2) = 2500
MSHFlexGrid1.ColWidth(3) = 500
MSHFlexGrid1.ColWidth(4) = 3000
MSHFlexGrid1.ColWidth(5) = 800
MSHFlexGrid1.ColWidth(6) = 1100
MSHFlexGrid1.ColWidth(7) = 1100
MSHFlexGrid1.ColWidth(8) = 1000
MSHFlexGrid1.ColWidth(9) = 2000
MSHFlexGrid1.ColWidth(10) = 1000
MSHFlexGrid1.ColWidth(11) = 900
MSHFlexGrid1.ColWidth(12) = 1000
MSHFlexGrid1.ColWidth(13) = 1500
MSHFlexGrid1.ColWidth(14) = 800

End Sub
Private Sub MSHFlexGrid1_Click()
Dim I As Variant
I = MSHFlexGrid1.RowSel
If I <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(I, 0) = "" Then Exit Sub
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationID = " & MSHFlexGrid1.TextMatrix(I, 0)
If Not RsMst_Location.EOF Then
   Call Indata
   Frame0.Visible = False
End If
End Sub
Private Sub CmdStateID_Click()
    FrmMst_State.Show
End Sub
'----
Private Sub CmbStateID_GotFocus()
tmp = CmbStateID.Text
Call TableMst_State
CmbStateID.Clear
If RsMst_State.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_State.RecordCount
    CmbStateID.AddItem RsMst_State!StateName
    RsMst_State.MoveNext
Next
CmbStateID.Text = tmp
End Sub
'----
Private Sub CmbStateID_LostFocus()
If CmbStateID.Text = "" Then
   CmbHRLocationID.Text = ""
   mStateIndiaName = ""
   CmbHRLocationID.Clear
   Exit Sub
End If
RsMst_State.MoveFirst
RsMst_State.Find "StateName = '" & CmbStateID.Text & "'"
If RsMst_State.EOF Then
   MsgBox "Invalid selection "
   CmbStateID.SetFocus
   Exit Sub
End If

If tmp <> RsMst_State!StateName Then
   CmbHRLocationID.Text = ""
   mStateIndiaName = ""
   CmbHRLocationID.Clear
End If


mStateID = RsMst_State!StateID

mStateIndiaName = GetStateIndiaFromStateName(CmbStateID.Text)

End Sub

Private Sub TxtLocationName_LostFocus()
TxtLocationName = Trim(TxtLocationName)
End Sub
Private Sub CmbSState_GotFocus()
tmp = CmbSState.Text
Call TableMst_State
CmbSState.Clear
If RsMst_State.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_State.RecordCount
    CmbSState.AddItem RsMst_State!StateName
    RsMst_State.MoveNext
Next
CmbSState.Text = tmp
End Sub
'----
Private Sub CmbSState_LostFocus()
If CmbSState.Text = "" Then
   Exit Sub
End If
RsMst_State.MoveFirst
RsMst_State.Find "StateName = '" & CmbSState.Text & "'"
If RsMst_State.EOF Then
   MsgBox "Invalid selection "
   CmbSState.SetFocus
   Exit Sub
End If

End Sub
Private Sub CmbEQZoneID_GotFocus()
tmp = CmbEQZoneID.Text
Call TableMst_EarthquakeZone(" Where Flag='A'")
CmbEQZoneID.Clear
If RsMst_EarthquakeZone.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_EarthquakeZone.RecordCount
    CmbEQZoneID.AddItem RsMst_EarthquakeZone!EQZone
    RsMst_EarthquakeZone.MoveNext
Next
CmbEQZoneID.Text = tmp
End Sub
'----
Private Sub CmbEQZoneID_LostFocus()
If CmbEQZoneID.Text = "" Then
   Exit Sub
End If
RsMst_EarthquakeZone.MoveFirst
RsMst_EarthquakeZone.Find "EQZone = '" & CmbEQZoneID.Text & "'"
If RsMst_EarthquakeZone.EOF Then
   MsgBox "Invalid selection "
   CmbEQZoneID.SetFocus
   Exit Sub
End If
mEQZoneID = RsMst_EarthquakeZone!EQZoneID
End Sub
Private Sub CmbSEQZoneID_GotFocus()
tmp = CmbSEQZoneID.Text
Call TableMst_EarthquakeZone(" Where Flag='A'")
CmbSEQZoneID.Clear
If RsMst_EarthquakeZone.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_EarthquakeZone.RecordCount
    CmbSEQZoneID.AddItem RsMst_EarthquakeZone!EQZone
    RsMst_EarthquakeZone.MoveNext
Next
CmbSEQZoneID.Text = tmp
End Sub
'----
Private Sub CmbSEQZoneID_LostFocus()
If CmbSEQZoneID.Text = "" Then
   Exit Sub
End If
RsMst_EarthquakeZone.MoveFirst
RsMst_EarthquakeZone.Find "EQZone = '" & CmbSEQZoneID.Text & "'"
If RsMst_EarthquakeZone.EOF Then
   MsgBox "Invalid selection "
   CmbSEQZoneID.SetFocus
   Exit Sub
End If
mSEQZoneID = RsMst_EarthquakeZone!EQZoneID
End Sub

Private Sub CmbHRLocationID_GotFocus()
If CmbStateID.Text = "" Then Exit Sub

tmp1 = CmbHRLocationID.Text

tmp = "select LocationID,LocationName from Mst_Location ML"
tmp = tmp & " inner join Mst_State MS on MS.StateID=ML.StateID"
tmp = tmp & " inner join Mst_StateIndia MSI on MS.StateIndiaID=MSI.StateIndiaID "
tmp = tmp & " Where MSI.StateIndiaName = '" & mStateIndiaName & "' And ML.StatusFlag='A'"

Call TmpTableHRPro

CmbHRLocationID.Clear
If TmpRsHRPro.RecordCount = 0 Then
   MsgBox "No HR Location found for this State"
   CmbHRLocationID.SetFocus
   Exit Sub
End If
For I = 1 To TmpRsHRPro.RecordCount
    CmbHRLocationID.AddItem TmpRsHRPro!LocationName
    TmpRsHRPro.MoveNext
Next
CmbHRLocationID.Text = tmp1
End Sub
Private Sub CmbHRLocationID_LostFocus()
If CmbHRLocationID.Text = "" Then
   Exit Sub
End If
TmpRsHRPro.MoveFirst
TmpRsHRPro.Find "LocationName = '" & CmbHRLocationID.Text & "'"

If TmpRsHRPro.EOF Then
   MsgBox "Invalid HR Location selection "
   CmbHRLocationID.SetFocus
   Exit Sub
End If
mHRLocationID = TmpRsHRPro!LocationID



End Sub

Private Sub CmbSHRLocMapped_LostFocus()
If CmbSHRLocMapped.Text = "" Then Exit Sub

If LCase(CmbSHRLocMapped.Text) = "yes" Or LCase(CmbSHRLocMapped.Text) = "no" Then
Else
   MsgBox "Invalid selection!!!"
   CmbSHRLocMapped.SetFocus
   Exit Sub
End If

End Sub

