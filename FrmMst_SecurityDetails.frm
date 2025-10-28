VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmMst_SecurityDetails 
   Caption         =   "Security Details"
   ClientHeight    =   9645
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12885
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9645
   ScaleWidth      =   12885
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   1575
      Left            =   5160
      TabIndex        =   45
      Top             =   6120
      Visible         =   0   'False
      Width           =   3615
      Begin VB.OptionButton Option3 
         Caption         =   "Designation"
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
         Left            =   6120
         TabIndex        =   77
         Top             =   240
         Width           =   1470
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
         Height          =   350
         Left            =   8400
         TabIndex        =   72
         Top             =   720
         Width           =   1125
      End
      Begin VB.OptionButton Option4 
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
         TabIndex        =   68
         Top             =   240
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Security Name"
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
         Left            =   1080
         TabIndex        =   67
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Shift Type"
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
         Left            =   4680
         TabIndex        =   66
         Top             =   240
         Width           =   1230
      End
      Begin VB.TextBox TxtSearchSecurity 
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
         Left            =   2430
         TabIndex        =   65
         Top             =   720
         Width           =   5415
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
         Left            =   9525
         TabIndex        =   64
         Top             =   720
         Width           =   1125
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Agency"
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
         Left            =   3030
         TabIndex        =   63
         Top             =   240
         Width           =   1470
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1380
         Left            =   120
         TabIndex        =   46
         Top             =   1200
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   2434
         _Version        =   393216
         Rows            =   3
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label19 
         Caption         =   "Search By"
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
         TabIndex        =   71
         Top             =   720
         Width           =   3015
      End
      Begin VB.Label Label20 
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
         Left            =   8310
         TabIndex        =   70
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label21 
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
         Left            =   10830
         TabIndex        =   69
         Top             =   240
         Width           =   75
      End
   End
   Begin VB.Frame Frame1 
      Height          =   8475
      Left            =   75
      TabIndex        =   27
      Top             =   0
      Width           =   12705
      Begin VB.ComboBox CmbDesignation 
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
         ItemData        =   "FrmMst_SecurityDetails.frx":0000
         Left            =   1920
         List            =   "FrmMst_SecurityDetails.frx":000A
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   1080
         Width           =   7875
      End
      Begin VB.ComboBox CmbSuperwisername 
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
         Left            =   1920
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   1530
         Width           =   7875
      End
      Begin VB.CommandButton CmdSecurityType 
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
         Left            =   9960
         TabIndex        =   44
         Top             =   2325
         Width           =   375
      End
      Begin VB.TextBox TxtShiftTimings 
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
         Left            =   6840
         TabIndex        =   7
         Top             =   2760
         Width           =   2955
      End
      Begin VB.ComboBox CmbSecurityType 
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
         Left            =   1905
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   2325
         Width           =   7890
      End
      Begin VB.ComboBox CmbAgencyID 
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
         Left            =   1920
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   1920
         Width           =   4515
      End
      Begin VB.CommandButton CmdAgencyID 
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
         Left            =   6480
         TabIndex        =   40
         Top             =   1905
         Width           =   375
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
         Left            =   1905
         Locked          =   -1  'True
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   3600
         Width           =   4980
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
         Left            =   1905
         Locked          =   -1  'True
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   3180
         Width           =   4980
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
         Height          =   375
         Left            =   1905
         TabIndex        =   6
         Top             =   2760
         Width           =   2985
      End
      Begin VB.TextBox TxtSecurityType 
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
         Left            =   1905
         Locked          =   -1  'True
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   2340
         Width           =   7890
      End
      Begin VB.TextBox TxtAgencyID 
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
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   1920
         Width           =   4485
      End
      Begin VB.TextBox TxtSecurityDetailsID 
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
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   240
         Width           =   2895
      End
      Begin VB.TextBox TxtSecurityName 
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
         Left            =   1920
         TabIndex        =   1
         Top             =   675
         Width           =   7845
      End
      Begin VB.TextBox TxtFlag 
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
         Left            =   8520
         Locked          =   -1  'True
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   1920
         Width           =   1245
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   4335
         Left            =   90
         TabIndex        =   8
         Top             =   4065
         Width           =   12420
         _ExtentX        =   21908
         _ExtentY        =   7646
         _Version        =   393216
         Tabs            =   2
         TabHeight       =   520
         TabCaption(0)   =   "CMP Details"
         TabPicture(0)   =   "FrmMst_SecurityDetails.frx":0027
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame2"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Asset Details"
         TabPicture(1)   =   "FrmMst_SecurityDetails.frx":0043
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame3"
         Tab(1).ControlCount=   1
         Begin VB.Frame Frame2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3855
            Left            =   120
            TabIndex        =   54
            Top             =   315
            Width           =   12225
            Begin VB.CommandButton CmdAddAll2grid 
               Caption         =   "Add All CMP to grid"
               Height          =   375
               Left            =   7320
               TabIndex        =   61
               Top             =   645
               Width           =   1815
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
               Left            =   960
               Sorted          =   -1  'True
               TabIndex        =   9
               Top             =   240
               Width           =   3015
            End
            Begin VB.CommandButton CmdRemoveFromGrid 
               Caption         =   "Remove From Grid"
               Height          =   375
               Left            =   10590
               TabIndex        =   57
               Top             =   645
               Width           =   1575
            End
            Begin VB.ComboBox CmbCMPID 
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
               Left            =   5535
               Sorted          =   -1  'True
               TabIndex        =   10
               Top             =   270
               Width           =   6135
            End
            Begin VB.CommandButton CmdAddToGrid 
               Caption         =   "Add To Grid"
               Height          =   375
               Left            =   9135
               TabIndex        =   11
               Top             =   645
               Width           =   1455
            End
            Begin VB.CommandButton CmdGodownID 
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
               Height          =   360
               Left            =   11760
               TabIndex        =   56
               Top             =   240
               Width           =   375
            End
            Begin VB.CommandButton CmdLocationID 
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
               Height          =   360
               Left            =   4080
               TabIndex        =   55
               Top             =   240
               Width           =   375
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
               Height          =   2700
               Left            =   120
               TabIndex        =   58
               Top             =   1035
               Width           =   12030
               _ExtentX        =   21220
               _ExtentY        =   4763
               _Version        =   393216
               WordWrap        =   -1  'True
               AllowUserResizing=   3
               _NumberOfBands  =   1
               _Band(0).Cols   =   2
            End
            Begin VB.Label Label12 
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
               Height          =   375
               Left            =   4695
               TabIndex        =   60
               Top             =   270
               Width           =   975
            End
            Begin VB.Label Label8 
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
               Height          =   255
               Left            =   165
               TabIndex        =   59
               Top             =   300
               Width           =   975
            End
         End
         Begin VB.Frame Frame3 
            Height          =   3735
            Left            =   -74880
            TabIndex        =   47
            Top             =   360
            Width           =   12210
            Begin VB.CommandButton CmdAssetType 
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
               Height          =   360
               Left            =   3960
               TabIndex        =   62
               Top             =   240
               Width           =   375
            End
            Begin VB.CommandButton CmdAdd2GridAsset 
               Caption         =   "Add To Grid"
               Height          =   375
               Left            =   8955
               TabIndex        =   16
               Top             =   960
               Width           =   1455
            End
            Begin VB.ComboBox CmbAssetType 
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
               Left            =   1335
               TabIndex        =   12
               Top             =   270
               Width           =   2535
            End
            Begin VB.CommandButton CmdDeleteFromGridAsset 
               Caption         =   "Remove From Grid"
               Height          =   375
               Left            =   10425
               TabIndex        =   48
               Top             =   960
               Width           =   1575
            End
            Begin VB.TextBox TxtAssetDescription 
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
               Height          =   615
               Left            =   1320
               MaxLength       =   500
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   15
               Top             =   720
               Width           =   7215
            End
            Begin VB.ComboBox CmbAssetStatus 
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
               ItemData        =   "FrmMst_SecurityDetails.frx":005F
               Left            =   5760
               List            =   "FrmMst_SecurityDetails.frx":006C
               Sorted          =   -1  'True
               TabIndex        =   13
               Top             =   270
               Width           =   2295
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
               Height          =   2220
               Left            =   120
               TabIndex        =   49
               Top             =   1395
               Width           =   11910
               _ExtentX        =   21008
               _ExtentY        =   3916
               _Version        =   393216
               WordWrap        =   -1  'True
               AllowUserResizing=   3
               _NumberOfBands  =   1
               _Band(0).Cols   =   2
            End
            Begin MSComCtl2.DTPicker AssetVerifiedon 
               Height          =   330
               Left            =   9720
               TabIndex        =   14
               Top             =   240
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   582
               _Version        =   393216
               Format          =   109969409
               CurrentDate     =   39850
            End
            Begin VB.Label Label14 
               Caption         =   "Asset Type"
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
               Left            =   135
               TabIndex        =   53
               Top             =   263
               Width           =   1095
            End
            Begin VB.Label Label15 
               Caption         =   "Description"
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
               Left            =   135
               TabIndex        =   52
               Top             =   877
               Width           =   1095
            End
            Begin VB.Label Label16 
               Caption         =   "Asset Status"
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
               Left            =   4440
               TabIndex        =   51
               Top             =   270
               Width           =   1215
            End
            Begin VB.Label Label17 
               Caption         =   "Verified On"
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
               Left            =   8640
               TabIndex        =   50
               Top             =   270
               Width           =   1095
            End
         End
      End
      Begin VB.TextBox TxtSuperwiserName 
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
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   1530
         Width           =   7845
      End
      Begin VB.TextBox TxtDesignation 
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
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   75
         TabStop         =   0   'False
         Top             =   1080
         Width           =   7845
      End
      Begin VB.Label Label13 
         Caption         =   "Designation"
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
         TabIndex        =   76
         Top             =   1110
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   "Superwiser Name"
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
         Left            =   120
         TabIndex        =   74
         Top             =   1530
         Width           =   1695
      End
      Begin VB.Label Label7 
         Caption         =   "Shift Timings"
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
         Left            =   5295
         TabIndex        =   42
         Top             =   2820
         Width           =   1380
      End
      Begin VB.Label Label11 
         Caption         =   "Agency Status"
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
         Left            =   6960
         TabIndex        =   41
         Top             =   1965
         Width           =   1530
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
         Height          =   240
         Left            =   120
         TabIndex        =   39
         Top             =   3667
         Width           =   1695
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
         TabIndex        =   38
         Top             =   3247
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "Amount"
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
         Left            =   150
         TabIndex        =   35
         Top             =   2820
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Security Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   34
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Agency Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   32
         Top             =   1995
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Security Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   30
         Top             =   690
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Security ID"
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
         TabIndex        =   29
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1095
      Left            =   75
      TabIndex        =   0
      Top             =   8475
      Width           =   12705
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New "
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   2415
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   375
         Left            =   2535
         TabIndex        =   17
         Top             =   240
         Width           =   2415
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   4950
         TabIndex        =   25
         Top             =   240
         Width           =   2535
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   7485
         TabIndex        =   24
         Top             =   240
         Width           =   2415
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "&Search"
         Height          =   375
         Left            =   9900
         TabIndex        =   19
         Top             =   240
         Width           =   2415
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   615
         Width           =   2415
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   375
         Left            =   2535
         TabIndex        =   22
         Top             =   615
         Width           =   2415
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   375
         Left            =   4950
         TabIndex        =   21
         Top             =   615
         Width           =   2535
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   375
         Left            =   7485
         TabIndex        =   20
         Top             =   615
         Width           =   2415
      End
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   375
         Left            =   9900
         TabIndex        =   18
         Top             =   615
         Width           =   2415
      End
   End
End
Attribute VB_Name = "FrmMst_SecurityDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg, mAgencyID, mSecurityTypeID, mCMPID, mFlag, mSecurityassetID, mSuperwiserID As Variant
Dim mSuperwiseragencyID, mSM_Prefix As Variant
Dim mLocationID As Double
Private Sub AddCmd_Click()
Edit_Flg = "A"
Call Grid_Head
Call GButtonLock(Me, False)
CmbAgencyID.Visible = True
CmbSecurityType.Visible = True
CmbAgencyID.Text = ""
CmbSecurityType.Text = ""
CmbDesignation.Text = ""
CmbDesignation.Visible = True
CmbSuperwisername.Text = ""
CmbSuperwisername.Visible = True
CmdAgencyID.Visible = True
CmdSecurityType.Visible = True
CmdAssetType.Visible = True
 
CmbSuperwisername.Enabled = False
TxtSecurityName.Locked = False
TxtAmount.Locked = False
TxtShiftTimings.Locked = False
TxtAgencyID.Visible = False
TxtSecurityType.Visible = False
TxtDesignation.Visible = False
TxtSuperwiserName.Visible = False
TxtSecurityDetailsID = ""
TxtSecurityName = ""
TxtAmount = ""
TxtShiftTimings = ""
TxtFlag = ""
TxtCreated = ""
TxtUpdated = ""
'Frame2.Enabled = True
'Frame3.Enabled = True
Call ClearTabs
CmdAddToGrid.Caption = "Add To Grid"
'Call FillAssetGrid
SSTab1.Tab = 0
mSuperwiseragencyID = Null
End Sub
Private Sub FillAssetGrid()
Call Grid_Head2
Call TableMst_SecurityAssets(" Order By SecurityAsset")

MSHFlexGrid3.Rows = RsMst_SecurityAssets.RecordCount + 2
For I = 1 To RsMst_SecurityAssets.RecordCount
    MSHFlexGrid3.TextMatrix(I, 0) = RsMst_SecurityAssets!SecurityAsset
    MSHFlexGrid3.TextMatrix(I, 1) = "Not Available"
    MSHFlexGrid3.TextMatrix(I, 2) = ""
    MSHFlexGrid3.TextMatrix(I, 3) = Date
    MSHFlexGrid3.TextMatrix(I, 4) = ""
    MSHFlexGrid3.TextMatrix(I, 5) = RsMst_SecurityAssets!SecurityAssetID
    RsMst_SecurityAssets.MoveNext
Next

End Sub



Private Sub AssetVerifiedon_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "AssetVerifiedon")
End If
End Sub

Private Sub CmbAgencyID_GotFocus()
tmp = CmbAgencyID.Text
Call TableMst_SecurityAgency("")
CmbAgencyID.Clear
If RsMst_SecurityAgency.RecordCount = 0 Then
   MsgBox "No Agency found"
   CmbAgencyID.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_SecurityAgency.RecordCount
    CmbAgencyID.AddItem RsMst_SecurityAgency!AgencyName
    RsMst_SecurityAgency.MoveNext
Next
CmbAgencyID.Text = tmp
End Sub

Private Sub CmbAgencyID_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "CmbAgencyID")
End If
End Sub

Private Sub CmbAgencyID_LostFocus()
If CmbAgencyID.Text = "" Then
   mAgencyID = Null
   TxtFlag = ""
   Exit Sub
End If

RsMst_SecurityAgency.MoveFirst
RsMst_SecurityAgency.Find "AgencyName = '" & CmbAgencyID.Text & "'"

If RsMst_SecurityAgency.EOF Then
   MsgBox "Invalid Agency Selection "
   CmbAgencyID.SetFocus
   Exit Sub
End If
mAgencyID = RsMst_SecurityAgency!SecurityAgencyID
mFlag = RsMst_SecurityAgency!Flag
If mFlag = "Y" Then
   TxtFlag = "Active"
Else
   TxtFlag = "De-Active"
End If
End Sub

Private Sub CmbAssetStatus_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "CmbAssetStatus")
End If
End Sub

Private Sub CmbAssetStatus_LostFocus()
If CmbAssetStatus.Text = "" Then
   Exit Sub
End If
If CmbAssetStatus.Text = "Available" Or CmbAssetStatus.Text = "Not Available" Or CmbAssetStatus.Text = "Not Applicable" Then

Else
   MsgBox "Invalid Status selected!!!"
   CmbAssetStatus.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmbAssetType_GotFocus()
tmp = CmbAssetType.Text
Call TableMst_SecurityAssets("")
CmbAssetType.Clear
If RsMst_SecurityAssets.RecordCount = 0 Then
   MsgBox "No Asset found!!!"
   CmbAssetType.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_SecurityAssets.RecordCount
    CmbAssetType.AddItem RsMst_SecurityAssets!SecurityAsset
    RsMst_SecurityAssets.MoveNext
Next
CmbAssetType.Text = tmp
End Sub

Private Sub CmbAssetType_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "CmbAssetType")
End If
End Sub

Private Sub CmbAssetType_LostFocus()
If CmbAssetType.Text = "" Then
   mSecurityassetID = Null
   Exit Sub
End If
RsMst_SecurityAssets.MoveFirst
RsMst_SecurityAssets.Find "SecurityAsset = '" & CmbAssetType.Text & "'"

If RsMst_SecurityAssets.EOF Then
   MsgBox "Invalid Asset Selection "
   CmbAssetType.SetFocus
   Exit Sub
End If
mSecurityassetID = RsMst_SecurityAssets!SecurityAssetID

End Sub

Private Sub CmbDesignation_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "CmbDesignation")
End If
End Sub

Private Sub CmbDesignation_LostFocus()
If CmbDesignation.Text = "" Then Exit Sub
If CmbDesignation.Text = "Securityman" Or CmbDesignation.Text = "Superwiser" Then
   
Else
   MsgBox "Invalid Degisnation !!!"
   CmbDesignation.SetFocus
   Exit Sub
End If
If CmbDesignation.Text = "Securityman" Then
   CmbSuperwisername.Enabled = True
   Call FillAssetGrid
   Frame2.Enabled = True
   Frame3.Enabled = True
   CmbSuperwisername.SetFocus
   CmbSecurityType.Enabled = True
   TxtShiftTimings.Enabled = True
   CmdSecurityType.Enabled = True
Else
   CmbSuperwisername.Enabled = False
   Call Grid_Head2
   Frame2.Enabled = False
   Frame3.Enabled = False
   CmbSuperwisername.Text = ""
   CmbSecurityType.Enabled = False
   TxtShiftTimings.Enabled = False
   CmdSecurityType.Enabled = False
   CmbSecurityType.Text = ""
End If

End Sub

Private Sub CmbCMPID_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "CmbCMPID")
End If
End Sub

Private Sub CmbLocationID_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "CmbLocationID")
End If
End Sub

Private Sub CmbSecurityType_GotFocus()
tmp = CmbSecurityType.Text
Call TableMst_SecurityShifts("")
CmbSecurityType.Clear
If RsMst_SecurityShifts.RecordCount = 0 Then
   MsgBox "No Security Type found!!!"
   CmbSecurityType.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_SecurityShifts.RecordCount
    CmbSecurityType.AddItem RsMst_SecurityShifts!SecurityShifts
    RsMst_SecurityShifts.MoveNext
Next
CmbSecurityType.Text = tmp
End Sub

Private Sub CmbSecurityType_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "CmbSecurityType")
End If
End Sub

Private Sub CmbSecurityType_LostFocus()
If CmbSecurityType.Text = "" Then
   mSecurityTypeID = Null
   Exit Sub
End If

RsMst_SecurityShifts.MoveFirst
RsMst_SecurityShifts.Find "SecurityShifts = '" & CmbSecurityType.Text & "'"

If RsMst_SecurityShifts.EOF Then
   MsgBox "Invalid Shift Selection "
   CmbSecurityType.SetFocus
   Exit Sub
End If
mSecurityTypeID = RsMst_SecurityShifts!SecurityShiftsID

End Sub
Private Sub CmbSuperwisername_GotFocus()
tmp = CmbSuperwisername.Text
'Call Mst_SecurityDetailTable("")
RsMst_SecurityDetails.Filter = "Designation = 'Superwiser'"
CmbSuperwisername.Clear
If RsMst_SecurityDetails.RecordCount = 0 Then
   MsgBox "No Superwiser found"
   If CmbSuperwisername.Enabled = True Then
      CmbSuperwisername.SetFocus
   End If
   Exit Sub
End If
For I = 1 To RsMst_SecurityDetails.RecordCount
    CmbSuperwisername.AddItem RsMst_SecurityDetails!SecurityName
    RsMst_SecurityDetails.MoveNext
Next
CmbSuperwisername.Text = tmp
End Sub

Private Sub CmbSuperwisername_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "CmbSuperwisername")
End If
End Sub

Private Sub CmbSuperwisername_LostFocus()
If CmbSuperwisername.Text = "" Then
   mSecurityDetailID = Null
   RsMst_SecurityDetails.Filter = "Designation <> ''"
   mSuperwiserID = Null
   mSuperwiseragencyID = Null
   'RsMst_SecurityDetails.MoveFirst
   Exit Sub
End If

RsMst_SecurityDetails.MoveFirst
RsMst_SecurityDetails.Find "SecurityName = '" & CmbSuperwisername.Text & "'"

If RsMst_SecurityDetails.EOF Then
   MsgBox "Invalid Superwiser Selection "
   CmbSuperwisername.SetFocus
   Exit Sub
End If
mSuperwiserID = RsMst_SecurityDetails!SecurityDetailID
mSuperwiseragencyID = RsMst_SecurityDetails!SecurityAgencyID
RsMst_SecurityDetails.Filter = "Designation <> ''"
RsMst_SecurityDetails.MoveFirst
End Sub

Private Sub CmdAddAll2grid_Click()
Dim I, ans As Variant
If CmbLocationID.Text = "" Then
   MsgBox "Blank Location not allowed!!!"
   CmbLocationID.SetFocus
   Exit Sub
End If
Call CmbCMPID_GotFocus
If TmpRs4.RecordCount > 0 Then
   TmpRs4.MoveFirst
   tmp = ""
   For iu = 0 To TmpRs4.RecordCount - 1
       mCMPID = TmpRs4!CMPID
       ans = Validate_DuplicateGrid()
       If ans = True Then
          Call FillGrid
       End If
       TmpRs4.MoveNext
   Next
End If
End Sub
Private Sub FillGrid()
Dim I
I = MSHFlexGrid2.Rows - 1
MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
MSHFlexGrid2.TextMatrix(I, 0) = CmbLocationID.Text
MSHFlexGrid2.TextMatrix(I, 1) = mCMPID
MSHFlexGrid2.TextMatrix(I, 2) = TmpRs4!CMPName
MSHFlexGrid2.TextMatrix(I, 3) = TmpRs4!Address
MSHFlexGrid2.TextMatrix(I, 4) = mLocationID
MSHFlexGrid2.TextMatrix(I, 6) = mSM_Prefix

End Sub

Private Sub CmdAgencyID_Click()
FrmMstSecurityAgency.Show
End Sub

Private Sub CmdAssetType_Click()
FrmMstSecurityAsset.Show
End Sub

Private Sub CmdExcel_Click()
Call Gen_Create_Xls("Mst_SecurityDetails.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\Mst_SecurityDetails.xls")
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
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub

Private Sub CmdGo_Click()

Dim tmp1 As Variant

tmp = "Select SD.SecurityDetailId,SD.SecurityName,SD.Designation,SD1.SecurityName as 'SuperWiser',MSA.AgencyName, " & _
      " Mst.SecurityShifts , Sd.Amount, Sd.ShiftTimings " & _
      " From Mst_SecurityDetails SD " & _
      " Inner Join Mst_SecurityAgency MSA On SD.SecurityAgencyID = MSA.SecurityAgencyID " & _
      " Left Outer Join Mst_SecurityShifts MST On SD.SecurityTypeID = MST.SecurityShiftsID " & _
      " Left Outer Join Mst_SecurityDetails SD1 On SD.SuperwiserID = SD1.SecurityDetailId "


Dim wc As Variant
wc = ""
If Option5.Value = True Then
   wc = " Where MSA.AgencyName Like '%" & TxtSearchSecurity & "%' "
End If
If Option1.Value = True Then
   wc = " Where SD.SecurityName  Like '%" & TxtSearchSecurity & "%'"
End If
If Option2.Value = True Then
   wc = " Where MST.SecurityShifts  Like '%" & TxtSearchSecurity & "%'"
End If
If Option3.Value = True Then
   wc = " Where SD.Designation  Like '%" & TxtSearchSecurity & "%'"
End If


tmp = tmp & " " & wc

Call TmpTable

Call Grid_Head1
For iu = 1 To TmpRs.RecordCount
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
    For C = 0 To MSHFlexGrid1.Cols - 1
       MSHFlexGrid1.TextMatrix(iu, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
    Next
    TmpRs.MoveNext
Next
Label21.Refresh
Label21.Caption = TmpRs.RecordCount


End Sub

Private Sub CmdLocationID_Click()
FrmLocationMaster.Show
End Sub

Private Sub CmdGodownID_Click()
FrmMstGodown.Show
End Sub

Private Sub CmdSecurityType_Click()
FrmMst_SecurityShifts.Show

End Sub


Private Sub DeleteCmd_Click()
Dim ans As Variant
On Error GoTo msgError
ans = MsgBox("Are you Sure you want to Delete this Record ", vbYesNo)
If ans = vbYes Then

   tmp = "Select * from Mst_SecurityDetails Where Superwiserid = " & TxtSecurityDetailsID & ""
   Call TmpTable
   
   If TmpRs.RecordCount > 0 Then
      MsgBox "Child record present!!!"
      Exit Sub
   End If
   
   tmp = "Delete from Mst_SecurityCMPMap where SecurityDetailID = " & TxtSecurityDetailsID
   Call TmpTable
   
   tmp = "Delete from Mst_SecurityAssetMap where SecurityDetailID = " & TxtSecurityDetailsID
   Call TmpTable
   
   tmp = "SELECT * from Mst_SecurityCMPMap where SecurityDetailID = " & TxtSecurityDetailsID
   Call TmpTable
   TmpRs.Close
   
   RsMst_SecurityDetails.Requery
   RsMst_SecurityDetails.Find "SecurityDetailId = " & TxtSecurityDetailsID
   If RsMst_SecurityDetails.EOF Then Exit Sub
   RsMst_SecurityDetails.Delete
   RsMst_SecurityDetails.Update
   If RsMst_SecurityDetails.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_SecurityDetails.MoveNext
   If RsMst_SecurityDetails.EOF() Then
      RsMst_SecurityDetails.MoveLast
   End If
   Call Indata
End If

Exit Sub
msgError:
MsgBox "Child Record Present "
RsMst_SecurityDetails.CancelUpdate
End Sub

Private Sub EditCmd_Click()
Edit_Flg = "E"
Call GButtonLock(Me, False)
CmbAgencyID.Visible = True
CmbSecurityType.Visible = True
CmbDesignation.Visible = True
CmbSuperwisername.Visible = True
CmbAgencyID.Text = TxtAgencyID.Text
CmbSecurityType.Text = TxtSecurityType.Text
CmbDesignation.Text = TxtDesignation.Text
CmbSuperwisername.Text = TxtSuperwiserName.Text
CmdAgencyID.Visible = True
CmdSecurityType.Visible = True
CmdAssetType.Visible = True
 
 
TxtSecurityName.Locked = False
TxtAmount.Locked = False
TxtShiftTimings.Locked = False
TxtAgencyID.Visible = False
TxtSecurityType.Visible = False
TxtDesignation.Visible = False
TxtSuperwiserName.Visible = False

If CmbDesignation.Text = "Securityman" Then
   Frame2.Enabled = True
   Frame3.Enabled = True
Else
   Frame2.Enabled = False
   Frame3.Enabled = False
End If
CmdAddToGrid.Caption = "Add To Grid"

End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub
Private Sub Form_Load()
Call Grid_Head
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Call GButtonLockAD(Me)
   MsgBox "Access Denied !!!!!!!!!"
   Exit Sub
End If
SSTab1.Tab = 0
Call TableMst_SecurityDetail("Order by SecurityDetailID ")
If RsMst_SecurityDetails.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
Call Indata
Call GButtonLock(Me, True)
End Sub
Private Sub Grid_Head()

MSHFlexGrid2.Clear
MSHFlexGrid2.Rows = 2
MSHFlexGrid2.Cols = 7

MSHFlexGrid2.TextMatrix(0, 0) = "Location"
MSHFlexGrid2.TextMatrix(0, 1) = "CMP Id"
MSHFlexGrid2.TextMatrix(0, 2) = "CMP Name "
MSHFlexGrid2.TextMatrix(0, 3) = "CMP Address"
MSHFlexGrid2.TextMatrix(0, 4) = "LocationID"
MSHFlexGrid2.TextMatrix(0, 5) = "SecurityCMPMapID"
MSHFlexGrid2.TextMatrix(0, 6) = "SM_Prefix"

MSHFlexGrid2.ColWidth(0) = 1400
MSHFlexGrid2.ColWidth(1) = 1600
MSHFlexGrid2.ColWidth(2) = 2200
MSHFlexGrid2.ColWidth(3) = 6600
MSHFlexGrid2.ColWidth(4) = 0
MSHFlexGrid2.ColWidth(5) = 0
MSHFlexGrid2.ColWidth(6) = 0

End Sub
Private Sub Grid_Head1()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 8
MSHFlexGrid1.TextMatrix(0, 0) = "Security Detail Id"
MSHFlexGrid1.TextMatrix(0, 1) = "Security Name"
MSHFlexGrid1.TextMatrix(0, 2) = "Designation"
MSHFlexGrid1.TextMatrix(0, 3) = "Superwiser Name"
MSHFlexGrid1.TextMatrix(0, 4) = "Agency Name"
MSHFlexGrid1.TextMatrix(0, 5) = "Security Type"
MSHFlexGrid1.TextMatrix(0, 6) = "Amount"
MSHFlexGrid1.TextMatrix(0, 7) = "Shift Timing"

MSHFlexGrid1.ColWidth(0) = 1400
MSHFlexGrid1.ColWidth(1) = 2800
MSHFlexGrid1.ColWidth(2) = 1400
MSHFlexGrid1.ColWidth(3) = 2800
MSHFlexGrid1.ColWidth(4) = 3400
MSHFlexGrid1.ColWidth(5) = 1000
MSHFlexGrid1.ColWidth(6) = 800
MSHFlexGrid1.ColWidth(7) = 1000

End Sub

Private Sub MSHFlexGrid3_Click()
Dim I As Variant
I = MSHFlexGrid3.RowSel
If MSHFlexGrid3.TextMatrix(I, 0) = "" Then Exit Sub
CmbAssetType.Text = MSHFlexGrid3.TextMatrix(I, 0)
CmbAssetStatus.Text = MSHFlexGrid3.TextMatrix(I, 1)
TxtAssetDescription = MSHFlexGrid3.TextMatrix(I, 2)
AssetVerifiedon.Value = MSHFlexGrid3.TextMatrix(I, 3)
CmbAssetType.SetFocus
CmdAdd2GridAsset.Caption = "Update Grid"
End Sub

Private Sub SaveCmd_Click()
If SaveCmd.Enabled = False Then Exit Sub

If TxtSecurityName = "" Then
   MsgBox "Blank Security Name not allowed!!!"
   TxtSecurityName.SetFocus
   Exit Sub
End If

If CmbDesignation.Text = "" Then
   MsgBox "Blank designation not allowed!!!"
   CmbDesignation.SetFocus
   Exit Sub
End If



If CmbAgencyID.Text = "" Then
   MsgBox "Blank Agency not allowed!!!"
   CmbAgencyID.SetFocus
   Exit Sub
End If

If LCase(CmbDesignation.Text) = "securityman" Then
   If CmbSecurityType.Text = "" Then
      MsgBox "Blank Security Type not allowed!!!"
      CmbSecurityType.SetFocus
      Exit Sub
   End If
   If TxtShiftTimings = "" Then
      MsgBox "Blank Shift Timings not allowed!!!"
      TxtShiftTimings.SetFocus
      Exit Sub
   End If
   
   If mSuperwiserID = TxtSecurityDetailsID.Text Then
      MsgBox "Self reporting not allowed!!!"
      CmbSuperwisername.SetFocus
      Exit Sub
   End If

End If

If mSuperwiseragencyID <> mAgencyID Then
   MsgBox "Superwiser should be of same agency!!!"
   CmbAgencyID.SetFocus
   Exit Sub
End If

If TxtAmount = "" Then
   MsgBox "Blank Amount not allowed!!!"
   TxtAmount.SetFocus
   Exit Sub
End If



If Edit_Flg = "A" Then
   If RsMst_SecurityDetails.RecordCount = 0 Then
      tmp = 1
   Else
      RsMst_SecurityDetails.Find "SecurityName = '" & TxtSecurityName & "'"
      If Not RsMst_SecurityDetails.EOF Then
         MsgBox "Duplicate Security Name not Allowed"
         TxtSecurityName.SetFocus
         Exit Sub
      End If
      RsMst_SecurityDetails.MoveFirst
      tmp = "select max(SecurityDetailID) from Mst_SecurityDetails "
      Call TmpTable
      tmp = IIf(IsNull(TmpRs.Fields(0)), 1, TmpRs.Fields(0) + 1)
   End If
   RsMst_SecurityDetails.AddNew
   RsMst_SecurityDetails!SecurityDetailID = tmp
   TxtSecurityDetailsID = tmp
Else
   RsMst_SecurityDetails.MoveFirst
   RsMst_SecurityDetails.Find "SecurityName = '" & TxtSecurityName & "'"
   If Not RsMst_SecurityDetails.EOF Then
      If RsMst_SecurityDetails!SecurityDetailID <> TxtSecurityDetailsID Then
         MsgBox "Duplicate Security Name Not Allowed"
         TxtSecurityName.SetFocus
         Exit Sub
      End If
   End If
   RsMst_SecurityDetails.MoveFirst
   RsMst_SecurityDetails.Find "SecurityDetailID = " & TxtSecurityDetailsID & ""
End If
RsMst_SecurityDetails!SecurityName = TxtSecurityName
RsMst_SecurityDetails!Designation = CmbDesignation.Text
RsMst_SecurityDetails!SuperwiserID = IIf(mSuperwiserID = 0, Null, mSuperwiserID)
RsMst_SecurityDetails!SecurityAgencyID = mAgencyID
RsMst_SecurityDetails!SecurityTypeID = IIf(mSecurityTypeID = 0, Null, mSecurityTypeID)
RsMst_SecurityDetails!Amount = Val(TxtAmount)
RsMst_SecurityDetails!ShiftTimings = TxtShiftTimings
RsMst_SecurityDetails!Flag = ""

If IsNull(RsMst_SecurityDetails!CreatedBy) = True Or RsMst_SecurityDetails!CreatedBy = "" Then
   RsMst_SecurityDetails!CreatedBy = Gen_UserLoginID
   RsMst_SecurityDetails!CreatedDate = Now
Else
   RsMst_SecurityDetails!UpdatedBy = Gen_UserLoginID
   RsMst_SecurityDetails!UpdatedDate = Now
End If
RsMst_SecurityDetails.Update

tmp = "Delete from Mst_SecurityCMPMap where SecurityDetailID = " & RsMst_SecurityDetails!SecurityDetailID & ""
Call TmpTable
If MSHFlexGrid2.Rows > 2 Then
   
   
   Call TableMst_SecurityCMPMap(" Where SecurityDetailID = " & RsMst_SecurityDetails!SecurityDetailID & "")
   For I = 1 To MSHFlexGrid2.Rows - 2
       If MSHFlexGrid2.TextMatrix(I, 5) <> "" Then
          'RsMst_SecurityCMPMap.MoveFirst
          RsMst_SecurityCMPMap.Find "SecurityCMPMapID = " & MSHFlexGrid2.TextMatrix(I, 5)
          If RsMst_SecurityCMPMap.EOF Then
             RsMst_SecurityCMPMap.AddNew
             RsMst_SecurityCMPMap!SecurityCMPMapID = GetSecurityCMPMapID()
          End If
       Else
          RsMst_SecurityCMPMap.AddNew
          RsMst_SecurityCMPMap!SecurityCMPMapID = GetSecurityCMPMapID()
       End If
       RsMst_SecurityCMPMap!LocationID = MSHFlexGrid2.TextMatrix(I, 4)
       RsMst_SecurityCMPMap!CMPID = MSHFlexGrid2.TextMatrix(I, 1)
       RsMst_SecurityCMPMap!SecurityDetailID = TxtSecurityDetailsID
       RsMst_SecurityCMPMap!SM_Prefix = MSHFlexGrid2.TextMatrix(I, 6)
       If IsNull(RsMst_SecurityCMPMap!CreatedBy) = True Or RsMst_SecurityCMPMap!CreatedBy = "" Then
          RsMst_SecurityCMPMap!CreatedBy = Gen_UserLoginID
          RsMst_SecurityCMPMap!CreatedDate = Now
       Else
          RsMst_SecurityCMPMap!UpdatedBy = Gen_UserLoginID
          RsMst_SecurityCMPMap!UpdatedDate = Now
       End If
       RsMst_SecurityCMPMap.Update
   Next
End If

tmp = "Delete from Mst_SecurityAssetMap where SecurityDetailID = " & RsMst_SecurityDetails!SecurityDetailID & ""
Call TmpTable
If MSHFlexGrid3.Rows > 2 Then

   
   Call TableMst_SecurityAssetMap(" Where SecurityDetailID = " & RsMst_SecurityDetails!SecurityDetailID & "")
   For I = 1 To MSHFlexGrid3.Rows - 2
       If MSHFlexGrid3.TextMatrix(I, 4) <> "" Then
          'RsMst_SecurityAssetMap.MoveFirst
          RsMst_SecurityAssetMap.Find "SecurityAssetMapID = " & MSHFlexGrid3.TextMatrix(I, 4)
          If RsMst_SecurityAssetMap.EOF Then
             RsMst_SecurityAssetMap.AddNew
             RsMst_SecurityAssetMap!SecurityAssetMapID = GetSecurityAssetMapID()
          End If
       Else
          RsMst_SecurityAssetMap.AddNew
          RsMst_SecurityAssetMap!SecurityAssetMapID = GetSecurityAssetMapID()
       End If
       RsMst_SecurityAssetMap!SecurityDetailID = TxtSecurityDetailsID
       RsMst_SecurityAssetMap!SecurityAssetID = MSHFlexGrid3.TextMatrix(I, 5)
       RsMst_SecurityAssetMap!AssetStatus = MSHFlexGrid3.TextMatrix(I, 1)
       RsMst_SecurityAssetMap!Description = MSHFlexGrid3.TextMatrix(I, 2)
       RsMst_SecurityAssetMap!VerifiedOn = MSHFlexGrid3.TextMatrix(I, 3)
       
       If IsNull(RsMst_SecurityAssetMap!CreatedBy) = True Or RsMst_SecurityAssetMap!CreatedBy = "" Then
          RsMst_SecurityAssetMap!CreatedBy = Gen_UserLoginID
          RsMst_SecurityAssetMap!CreatedDate = Now
       Else
          RsMst_SecurityAssetMap!UpdatedBy = Gen_UserLoginID
          RsMst_SecurityAssetMap!UpdatedDate = Now
       End If
       RsMst_SecurityAssetMap.Update
   Next
End If



Call GButtonLock(Me, True)
RsMst_SecurityDetails.Requery
RsMst_SecurityDetails.MoveFirst
RsMst_SecurityDetails.Find "SecurityDetailID = " & TxtSecurityDetailsID & ""
If RsMst_SecurityDetails.EOF Then
   RsMst_SecurityDetails.MoveLast
End If
Call Indata


End Sub
Private Function GetSecurityCMPMapID() As Double
Dim Retval As Double
tmp = "select max(SecurityCMPMapID) from Mst_SecurityCMPMap"
Call TmpTable
tmp = IIf(IsNull(TmpRs.Fields(0)), 1, TmpRs.Fields(0) + 1)
GetSecurityCMPMapID = tmp
End Function
Private Function GetSecurityAssetMapID() As Double
Dim Retval As Double
tmp = "select max(SecurityAssetMapID) from Mst_SecurityAssetMap"
Call TmpTable
tmp = IIf(IsNull(TmpRs.Fields(0)), 1, TmpRs.Fields(0) + 1)
GetSecurityAssetMapID = tmp
End Function

Private Sub TxtAmount_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "TxtAmount")
End If
End Sub

Private Sub TxtAmount_LostFocus()

If TxtAmount = "" Then Exit Sub
If IsNumeric(TxtAmount) = False Or Val(TxtAmount) < 0 Then
   MsgBox "Invalid Number Format!!!"
   TxtAmount.SetFocus
   Exit Sub
End If

End Sub
Private Sub CmbLocationID_GotFocus()
tmp = CmbLocationID.Text
Call TableMst_Location("")
CmbLocationID.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Location found"
   CmbLocationID.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbLocationID.AddItem RsMst_Location!LocationName
    RsMst_Location.MoveNext
Next
CmbLocationID.Text = tmp
End Sub

Private Sub CmbLocationID_LostFocus()
If CmbLocationID.Text = "" Then
   mLocationID = 0
   Exit Sub
End If
' Code for clearing Godown combox in case of change
If CmbLocationID.Text <> tmp Then
   CmbCMPID.Clear
   CmbCMPID.Text = ""
   mCMPID = Null
End If

RsMst_Location.MoveFirst
RsMst_Location.Find "Locationname = '" & CmbLocationID.Text & "'"

If RsMst_Location.EOF Then
   MsgBox "Invalid Location  selection "
   CmbLocationID.SetFocus
   Exit Sub
End If
mLocationID = RsMst_Location.Fields(0)
mSM_Prefix = RsMst_Location!SM_Prefix
End Sub
Private Sub CmbCMPID_GotFocus()
If CmbLocationID.Text = "" Then Exit Sub
temp = CmbCMPID.Text
'tmp = " Select mg.GodownID,mg.GodownNo,mg.Address,mc.CMPID " & _
'      " From Mst_Godown mg " & _
'      " Inner Join Mst_CMP mc on mg.CMPID = mc.CMPID " & _
'      " Where mc.locationID = " & mLocationId & " and mg.closedate is null "

tmp = "Select * from Mst_CMP Where LocationID = " & mLocationID & " And CloseDate is null"
Call Tmp4Table

Call Tmp4Table

CmbCMPID.Clear

If TmpRs4.RecordCount = 0 Then
   MsgBox "No CMP Found In Above Location found"
   CmbCMPID.SetFocus
   Exit Sub
End If
For I = 1 To TmpRs4.RecordCount
    CmbCMPID.AddItem TmpRs4!CMPName
    TmpRs4.MoveNext
Next
CmbCMPID.Text = temp

End Sub
Private Sub CmbCMPID_LostFocus()
If CmbCMPID.Text = "" Then
   Exit Sub
End If

TmpRs4.MoveFirst
TmpRs4.Find "CMPName = '" & CmbCMPID.Text & "'"

If TmpRs4.EOF Then
   MsgBox "Invalid CMP Selection "
   CmbCMPID.SetFocus
   Exit Sub
End If
mCMPID = TmpRs4!CMPID
End Sub
Private Sub CmdAddToGrid_Click()

If CmbLocationID.Text = "" Then
   MsgBox "Blank Location Not Allowed!!!"
   CmbLocationID.SetFocus
   Exit Sub
End If

If CmbCMPID.Text = "" Then
   MsgBox "Blank CMP Not Allowed!!!"
   CmbCMPID.SetFocus
   Exit Sub
End If

Dim I, ans As Variant


'tmp = "Select CMPID,DesignationID from Mst_EMPCMPMapping "
'Call Tmp3Table
'If TmpRs3.RecordCount > 0 Then
'   If CheckCMPMapping = False Then
'      MsgBox "CMP Already mapped !!!! "
'      Exit Sub
'   End If
'End If


ans = Validate_DuplicateGrid()
If ans = False Then
   MsgBox "Duplicate Entry !!!! "
   Exit Sub
End If

If CmdAddToGrid.Caption = "Add To Grid" Then
   I = MSHFlexGrid2.Rows - 1
   MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
Else
   I = MSHFlexGrid2.RowSel
End If

MSHFlexGrid2.TextMatrix(I, 0) = CmbLocationID.Text
MSHFlexGrid2.TextMatrix(I, 1) = mCMPID
MSHFlexGrid2.TextMatrix(I, 2) = TmpRs4!CMPName
MSHFlexGrid2.TextMatrix(I, 3) = TmpRs4!Address
MSHFlexGrid2.TextMatrix(I, 4) = mLocationID
MSHFlexGrid2.TextMatrix(I, 6) = mSM_Prefix

CmdAddToGrid.Caption = "Add To Grid"
CmbCMPID.Text = ""
CmbCMPID.SetFocus

End Sub
Private Function Validate_DuplicateGrid() As Boolean
Dim Retval As Boolean
Retval = True
If MSHFlexGrid2.Rows = 2 Then
   Retval = True
Else
    For I = 1 To MSHFlexGrid2.Rows - 1
        If CmdAddToGrid.Caption = "Update Grid" Then
           If I <> MSHFlexGrid2.RowSel Then
              If MSHFlexGrid2.TextMatrix(I, 1) = mCMPID Then
                 Retval = False
              End If
           End If
        Else
           If MSHFlexGrid2.TextMatrix(I, 1) = mCMPID Then
              Retval = False
           End If
        End If
    Next
End If
Validate_DuplicateGrid = Retval
End Function

Private Sub MSHFlexGrid1_Click()
Dim x As Variant
Dim mCaId As Variant

x = MSHFlexGrid1.RowSel
mCaId = MSHFlexGrid1.TextMatrix(x, 0)
RsMst_SecurityDetails.MoveFirst

If RsMst_SecurityDetails.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
RsMst_SecurityDetails.MoveFirst
RsMst_SecurityDetails.Find "SecurityDetailID = " & Val(mCaId) & ""
If RsMst_SecurityDetails.EOF Then
   RsMst_SecurityDetails.MoveLast
End If
Frame0.Visible = False
Call Indata
End Sub

Private Sub MSHFlexGrid2_Click()
Dim I As Variant
I = MSHFlexGrid2.RowSel
If MSHFlexGrid2.TextMatrix(I, 0) = "" Then Exit Sub
CmbLocationID.Text = MSHFlexGrid2.TextMatrix(I, 0)
CmbCMPID.Text = MSHFlexGrid2.TextMatrix(I, 2)
CmbLocationID.SetFocus
CmdAddToGrid.Caption = "Update Grid"
End Sub
Private Sub CmdRemoveFromGrid_Click()
If MSHFlexGrid2.Rows < 3 Then Exit Sub
Dim ans As Variant
ans = MsgBox("Are you Sure you want to Delete this Record ", vbYesNo)
If ans = vbYes Then
   I = MSHFlexGrid2.RowSel
   If I > 0 Then
      If MSHFlexGrid2.TextMatrix(I, 0) <> "" Then
         For RO = I To MSHFlexGrid2.Rows - 2
             For C = 0 To MSHFlexGrid2.Cols - 1
                 MSHFlexGrid2.TextMatrix(RO, C) = MSHFlexGrid2.TextMatrix(RO + 1, C)
             Next
         Next
         MSHFlexGrid2.Rows = MSHFlexGrid2.Rows - 1
      End If
   End If
   CmdAddToGrid.Caption = "Add To Grid"
   'CmbLocationID.Text = ""
   CmbCMPID.Text = ""
Else
   CmbCMPID.SetFocus
End If
End Sub
Public Sub Indata()
Call GButtonLock(Me, True)
Frame2.Enabled = False
Frame3.Enabled = False

CmbAgencyID.Visible = False
CmbSecurityType.Visible = False
CmbDesignation.Visible = False
CmbSuperwisername.Visible = False

CmdAgencyID.Visible = False
CmdSecurityType.Visible = False
CmdLocationID.Visible = False
CmdAssetType.Visible = False

TxtAgencyID.Visible = True
TxtSecurityType.Visible = True
TxtDesignation.Visible = True
TxtSuperwiserName.Visible = True

TxtSecurityDetailsID = RsMst_SecurityDetails!SecurityDetailID
TxtSecurityName = RsMst_SecurityDetails!SecurityName
TxtDesignation = RsMst_SecurityDetails!Designation
mSuperwiserID = RsMst_SecurityDetails!SuperwiserID
If IsNull(mSuperwiserID) = False Then
   TxtSuperwiserName = GetSuperwiserName(RsMst_SecurityDetails!SuperwiserID, "N")
   mSuperwiseragencyID = GetSuperwiserName(RsMst_SecurityDetails!SuperwiserID, "A")
Else
   TxtSuperwiserName = ""
   mSuperwiseragencyID = Null
End If

mAgencyID = RsMst_SecurityDetails!SecurityAgencyID
TxtAgencyID = GetSecurityAgency(mAgencyID, "N")
mSecurityTypeID = RsMst_SecurityDetails!SecurityTypeID
If IsNull(mSecurityTypeID) = False Then
   TxtSecurityType = GetSecurityType(mSecurityTypeID)
Else
   TxtSecurityType = ""
End If
TxtAmount = RsMst_SecurityDetails!Amount
TxtShiftTimings = RsMst_SecurityDetails!ShiftTimings
mFlag = GetSecurityAgency(mAgencyID, "S")
If mFlag = "Y" Then
   TxtFlag = "Active"
Else
   TxtFlag = "De-Active"
End If

TxtSecurityName.Locked = True
TxtAmount.Locked = True
TxtShiftTimings.Locked = True

Frame2.Enabled = False
Frame3.Enabled = False

TxtCreated = IIf(IsNull(RsMst_SecurityDetails!CreatedBy), "", RsMst_SecurityDetails!CreatedBy) & " :- " & IIf(IsNull(RsMst_SecurityDetails!CreatedDate), "", RsMst_SecurityDetails!CreatedDate)
TxtUpdated = IIf(IsNull(RsMst_SecurityDetails!UpdatedBy), "", RsMst_SecurityDetails!UpdatedBy) & " :- " & IIf(IsNull(RsMst_SecurityDetails!UpdatedDate), "", RsMst_SecurityDetails!UpdatedDate)
Call ClearTabs
Call Grid_Head
Call TableMst_SecurityCMPMap(" Where SecurityDetailID = " & RsMst_SecurityDetails!SecurityDetailID & "")
If RsMst_SecurityCMPMap.RecordCount > 0 Then
   For I = 1 To RsMst_SecurityCMPMap.RecordCount
       MSHFlexGrid2.TextMatrix(I, 0) = GetLocationName(RsMst_SecurityCMPMap!LocationID)
       MSHFlexGrid2.TextMatrix(I, 1) = RsMst_SecurityCMPMap!CMPID
       MSHFlexGrid2.TextMatrix(I, 2) = GetCMPName(RsMst_SecurityCMPMap!CMPID)
       MSHFlexGrid2.TextMatrix(I, 3) = GetCMPAddress(RsMst_SecurityCMPMap!CMPID, "A")
       MSHFlexGrid2.TextMatrix(I, 4) = RsMst_SecurityCMPMap!LocationID
       MSHFlexGrid2.TextMatrix(I, 5) = RsMst_SecurityCMPMap!SecurityCMPMapID
       MSHFlexGrid2.TextMatrix(I, 6) = IIf(IsNull(RsMst_SecurityCMPMap!SM_Prefix), "", RsMst_SecurityCMPMap!SM_Prefix)
       RsMst_SecurityCMPMap.MoveNext
       MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
   Next
End If

Call Grid_Head2
Call TableMst_SecurityAssetMap(" Where SecurityDetailID = " & RsMst_SecurityDetails!SecurityDetailID & "")

If RsMst_SecurityAssetMap.RecordCount > 0 Then
   For I = 1 To RsMst_SecurityAssetMap.RecordCount
       MSHFlexGrid3.TextMatrix(I, 0) = GetSecurityAsset(RsMst_SecurityAssetMap!SecurityAssetID)
       MSHFlexGrid3.TextMatrix(I, 1) = RsMst_SecurityAssetMap!AssetStatus
       MSHFlexGrid3.TextMatrix(I, 2) = RsMst_SecurityAssetMap!Description
       MSHFlexGrid3.TextMatrix(I, 3) = RsMst_SecurityAssetMap!VerifiedOn
       MSHFlexGrid3.TextMatrix(I, 4) = RsMst_SecurityAssetMap!SecurityAssetMapID
       MSHFlexGrid3.TextMatrix(I, 5) = RsMst_SecurityAssetMap!SecurityAssetID
       RsMst_SecurityAssetMap.MoveNext
       MSHFlexGrid3.Rows = MSHFlexGrid3.Rows + 1
   Next
End If
SSTab1.Tab = 0

Call GButtonLock(Me, True)

End Sub
Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_SecurityDetails.MoveNext
If RsMst_SecurityDetails.EOF Then
   RsMst_SecurityDetails.MoveLast
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
RsMst_SecurityDetails.MovePrevious
If RsMst_SecurityDetails.BOF Then
   RsMst_SecurityDetails.MoveFirst
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
RsMst_SecurityDetails.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsMst_SecurityDetails.MoveLast
Call Indata
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = Fasle
End Sub

Private Sub SearchCmd_Click()
Dim x As Integer
If SearchCmd.Caption = "Cancel" Then
   If RsMst_SecurityDetails.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   If RsMst_SecurityDetails.EOF Then
      RsMst_SecurityDetails.MoveFirst
   End If
   Call Indata
   Exit Sub
End If

Frame0.Height = Me.Height - 700
Frame0.Left = Frame1.Left
Frame0.Top = Frame1.Top
Frame0.Width = Me.Width - 400
Frame0.Visible = True
MSHFlexGrid1.Clear

MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 1400


'Option4.Value = True
RsMst_SecurityDetails.MoveFirst
Call Grid_Head1

Label21.Refresh
Label21.Caption = "0" 'TmpRs.RecordCount

End Sub

Private Sub ClearTabs()
CmbLocationID.Text = ""
CmbCMPID.Text = ""
Call Grid_Head
CmbAssetType.Text = ""
CmbAssetStatus.Text = ""
TxtAssetDescription = ""
AssetVerifiedon.Value = Date
Call Grid_Head2
End Sub
Private Sub Grid_Head2()

MSHFlexGrid3.Clear
MSHFlexGrid3.Rows = 2
MSHFlexGrid3.Cols = 6
MSHFlexGrid3.TextMatrix(0, 0) = "Asset Type"
MSHFlexGrid3.TextMatrix(0, 1) = "Asset Status"
MSHFlexGrid3.TextMatrix(0, 2) = "Description"
MSHFlexGrid3.TextMatrix(0, 3) = "Verified On"
MSHFlexGrid3.TextMatrix(0, 4) = "SecurityAssetMapID"
MSHFlexGrid3.TextMatrix(0, 5) = "SecurityAssetID"

MSHFlexGrid3.ColWidth(0) = 1800
MSHFlexGrid3.ColWidth(1) = 1800
MSHFlexGrid3.ColWidth(2) = 6800
MSHFlexGrid3.ColWidth(3) = 1400
MSHFlexGrid3.ColWidth(4) = 0
MSHFlexGrid3.ColWidth(5) = 0

End Sub
Private Sub CmdAdd2GridAsset_Click()
If CmbAssetType.Text = "" Then
   MsgBox "Blank Asset Type not allowed!!!"
   CmbAssetType.SetFocus
   Exit Sub
End If
If CmbAssetStatus.Text = "" Then
   MsgBox "Blank Asset Status not allowed!!!"
   CmbAssetStatus.SetFocus
   Exit Sub
End If

If AssetVerifiedon.Value > Date Then
   MsgBox "Future Verified date not allowed!!!"
   AssetVerifiedon.SetFocus
   Exit Sub
End If



Dim I, ans As Variant

If CmdAdd2GridAsset.Caption = "Add To Grid" Then
   I = MSHFlexGrid3.Rows - 1
   MSHFlexGrid3.Rows = MSHFlexGrid3.Rows + 1
Else
   I = MSHFlexGrid3.RowSel
End If

MSHFlexGrid3.TextMatrix(I, 0) = CmbAssetType.Text
MSHFlexGrid3.TextMatrix(I, 1) = CmbAssetStatus.Text
MSHFlexGrid3.TextMatrix(I, 2) = TxtAssetDescription
MSHFlexGrid3.TextMatrix(I, 3) = AssetVerifiedon.Value


CmdAdd2GridAsset.Caption = "Add To Grid"
CmbAssetType.Text = ""
CmbAssetStatus.Text = ""
TxtAssetDescription = ""
AssetVerifiedon.Value = Date


End Sub
Private Sub CmdDeleteFromGridAsset_Click()
If MSHFlexGrid3.Rows < 3 Then Exit Sub
Dim ans As Variant
ans = MsgBox("Are you Sure you want to Delete this Record ", vbYesNo)
If ans = vbYes Then
   I = MSHFlexGrid3.RowSel
   If I > 0 Then
      If MSHFlexGrid3.TextMatrix(I, 0) <> "" Then
         For RO = I To MSHFlexGrid3.Rows - 2
             For C = 0 To MSHFlexGrid3.Cols - 1
                 MSHFlexGrid3.TextMatrix(RO, C) = MSHFlexGrid3.TextMatrix(RO + 1, C)
             Next
         Next
         MSHFlexGrid3.Rows = MSHFlexGrid3.Rows - 1
      End If
   CmdAdd2GridAsset.Caption = "Add To Grid"
   End If
End If
End Sub

Private Sub TxtAssetDescription_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "TxtAssetDescription")
End If
End Sub

Private Sub TxtFlag_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "TxtFlag")
End If
End Sub

Private Sub TxtSearchSecurity_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "TxtSearchSecurity")
End If
End Sub


Private Sub Option1_Click()
Label19.Visible = True
TxtSearchSecurity.Visible = True
TxtSearchSecurity = ""
CmdGo.Visible = True
If Option1.Value = True Then
   Label19.Caption = "Search By Security Name"
ElseIf Option2.Value = True Then
   Label19.Caption = "Search By Shift Type"
ElseIf Option5.Value = True Then
   Label19.Caption = "Search By Agency"
End If
TxtSearchSecurity.SetFocus
End Sub

Private Sub Option2_Click()
Call Option1_Click
End Sub

Private Sub Option5_Click()
Call Option1_Click
End Sub

Private Sub TxtSecurityDetailsID_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "TxtSecurityDetailsID")
End If
End Sub

Private Sub TxtSecurityName_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "TxtSecurityName")
End If
End Sub

Private Sub TxtShiftTimings_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "TxtShiftTimings")
End If
End Sub
