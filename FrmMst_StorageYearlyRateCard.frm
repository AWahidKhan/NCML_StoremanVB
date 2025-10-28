VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmMst_StorageYearlyRateCard 
   Caption         =   "Storage Yearly Rate Card"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   2295
      Left            =   720
      TabIndex        =   8
      Top             =   3720
      Visible         =   0   'False
      Width           =   11280
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
         Left            =   10725
         TabIndex        =   12
         Top             =   240
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
         Height          =   350
         Left            =   9720
         TabIndex        =   11
         Top             =   240
         Width           =   1000
      End
      Begin VB.ComboBox CmbSCommodityID 
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
         Left            =   6630
         Sorted          =   -1  'True
         TabIndex        =   10
         Top             =   240
         Width           =   2880
      End
      Begin VB.ComboBox CmbSLocationID 
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
         Left            =   1080
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   240
         Width           =   2865
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1095
         Left            =   120
         TabIndex        =   13
         Top             =   1005
         Width           =   14655
         _ExtentX        =   25850
         _ExtentY        =   1931
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
      Begin VB.Label LblTotalRecords 
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
         Left            =   13680
         TabIndex        =   17
         Top             =   690
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
         Left            =   11280
         TabIndex        =   16
         Top             =   690
         Width           =   2340
      End
      Begin VB.Label Label2 
         Caption         =   "Commodity ID"
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
         Left            =   4440
         TabIndex        =   15
         Top             =   300
         Width           =   1935
      End
      Begin VB.Label Label3 
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
         Height          =   210
         Left            =   120
         TabIndex        =   14
         Top             =   315
         Width           =   960
      End
   End
   Begin VB.Frame Frame1 
      Height          =   8355
      Left            =   120
      TabIndex        =   25
      Top             =   0
      Width           =   15030
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
         ItemData        =   "FrmMst_StorageYearlyRateCard.frx":0000
         Left            =   8400
         List            =   "FrmMst_StorageYearlyRateCard.frx":000A
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   2955
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
         Left            =   1080
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   720
         Width           =   4455
      End
      Begin VB.CommandButton CmdLocationID 
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
         Height          =   300
         Left            =   5625
         TabIndex        =   40
         Top             =   750
         Width           =   375
      End
      Begin VB.CommandButton CmdCommodityID 
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
         Height          =   315
         Left            =   14205
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   735
         Width           =   375
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
         Left            =   8400
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   720
         Width           =   5760
      End
      Begin VB.TextBox TxtYSRCID 
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
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   240
         Width           =   3240
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   6930
         Left            =   120
         TabIndex        =   26
         Top             =   1320
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   12224
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Godown Detail"
         TabPicture(0)   =   "FrmMst_StorageYearlyRateCard.frx":0020
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame2"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Log History"
         TabPicture(1)   =   "FrmMst_StorageYearlyRateCard.frx":003C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "TxtUpdated"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "TxtCreated"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "Label12"
         Tab(1).Control(3)=   "Label13"
         Tab(1).ControlCount=   4
         Begin VB.Frame Frame2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   6480
            Left            =   105
            TabIndex        =   29
            Top             =   330
            Width           =   14565
            Begin VB.CommandButton CmdUpdate 
               Caption         =   "Update List"
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
               Left            =   10275
               TabIndex        =   32
               Top             =   262
               Width           =   1680
            End
            Begin VB.CheckBox Check1 
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
               Left            =   165
               TabIndex        =   31
               Top             =   345
               Width           =   1200
            End
            Begin VB.TextBox TxtRate 
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
               Left            =   8280
               TabIndex        =   30
               Top             =   255
               Width           =   1815
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
               Height          =   5670
               Left            =   120
               TabIndex        =   33
               Top             =   720
               Width           =   14340
               _ExtentX        =   25294
               _ExtentY        =   10001
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
            Begin VB.Label LblRatePerBag 
               Caption         =   "Rate"
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
               Left            =   7650
               TabIndex        =   36
               Top             =   315
               Width           =   525
            End
            Begin VB.Label Label4 
               Caption         =   "Total Record Found :-"
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
               Left            =   3360
               TabIndex        =   35
               Top             =   345
               Width           =   2415
            End
            Begin VB.Label LblGodownCount 
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
               Left            =   5880
               TabIndex        =   34
               Top             =   345
               Width           =   90
            End
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
            Left            =   -73695
            Locked          =   -1  'True
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   1560
            Width           =   11175
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
            Left            =   -73695
            Locked          =   -1  'True
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   2025
            Width           =   11175
         End
         Begin VB.Label Label13 
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
            Left            =   -74760
            TabIndex        =   38
            Top             =   1620
            Width           =   855
         End
         Begin VB.Label Label12 
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
            Left            =   -74760
            TabIndex        =   37
            Top             =   2085
            Width           =   885
         End
      End
      Begin VB.TextBox TxtLocationID 
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
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   720
         Width           =   4440
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
         Height          =   360
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   720
         Width           =   5760
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
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   240
         Visible         =   0   'False
         Width           =   2940
      End
      Begin VB.Label Label5 
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
         Left            =   6390
         TabIndex        =   47
         Top             =   300
         Width           =   660
      End
      Begin VB.Label LblLocationID 
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
         Left            =   120
         TabIndex        =   45
         Top             =   750
         Width           =   1440
      End
      Begin VB.Label Label1 
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
         Left            =   6390
         TabIndex        =   44
         Top             =   750
         Width           =   1920
      End
      Begin VB.Label LblWHRNo 
         Caption         =   "YSCR ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   43
         Top             =   315
         Width           =   900
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1200
      Left            =   135
      TabIndex        =   18
      Top             =   8385
      Width           =   15045
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
         Left            =   11745
         TabIndex        =   6
         Top             =   645
         Width           =   2880
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
         Left            =   11745
         TabIndex        =   5
         Top             =   240
         Width           =   2880
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
         Left            =   8865
         TabIndex        =   24
         Top             =   645
         Width           =   2880
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
         Left            =   8865
         TabIndex        =   23
         Top             =   240
         Width           =   2880
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
         Height          =   400
         Left            =   5985
         TabIndex        =   22
         Top             =   645
         Width           =   2880
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
         Left            =   5985
         TabIndex        =   21
         Top             =   240
         Width           =   2880
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
         Left            =   3105
         TabIndex        =   20
         Top             =   645
         Width           =   2880
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
         Left            =   3105
         TabIndex        =   4
         Top             =   240
         Width           =   2880
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
         Left            =   225
         TabIndex        =   19
         Top             =   645
         Width           =   2880
      End
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
         Left            =   225
         TabIndex        =   0
         Top             =   240
         Width           =   2880
      End
   End
End
Attribute VB_Name = "FrmMst_StorageYearlyRateCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mLocationID, mCommodityId, Edit_Flg, mSLocationID As Variant
Const strChecked = "þ"
Const strUnChecked = "q"

Private Sub Grid_Head1()
LblGodownCount.Caption = "-"
MSHFlexGrid2.Clear
MSHFlexGrid2.Rows = 2
MSHFlexGrid2.Cols = 10

MSHFlexGrid2.FixedRows = 1
MSHFlexGrid2.FixedCols = 0

MSHFlexGrid2.TextMatrix(0, 0) = "Select"
MSHFlexGrid2.TextMatrix(0, 1) = "CMP Name"
MSHFlexGrid2.TextMatrix(0, 2) = "Godown No"
MSHFlexGrid2.TextMatrix(0, 3) = "Godown Address"
MSHFlexGrid2.TextMatrix(0, 4) = "Unit Type"
MSHFlexGrid2.TextMatrix(0, 5) = "Start Date"
MSHFlexGrid2.TextMatrix(0, 6) = "Close Date"
MSHFlexGrid2.TextMatrix(0, 7) = "Rate"
MSHFlexGrid2.TextMatrix(0, 8) = "GodownID"
MSHFlexGrid2.TextMatrix(0, 9) = "NSRCDID"

'MSHFlexGrid2.ColWidth(0) = 0
MSHFlexGrid2.ColWidth(0) = 800
MSHFlexGrid2.ColWidth(1) = 2600
MSHFlexGrid2.ColWidth(2) = 1200
MSHFlexGrid2.ColWidth(3) = 4200
MSHFlexGrid2.ColWidth(4) = 1400
MSHFlexGrid2.ColWidth(5) = 1200
MSHFlexGrid2.ColWidth(6) = 1200
MSHFlexGrid2.ColWidth(7) = 1200
MSHFlexGrid2.ColWidth(8) = 0
MSHFlexGrid2.ColWidth(9) = 0

MSHFlexGrid2.RowHeight(0) = 600
MSHFlexGrid2.WordWrap = True

End Sub
Private Sub Grid_Head()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 4 '3

MSHFlexGrid1.FixedRows = 1
MSHFlexGrid1.FixedCols = 1

MSHFlexGrid1.TextMatrix(0, 0) = "YSRC ID"
MSHFlexGrid1.TextMatrix(0, 1) = "Location Name"
MSHFlexGrid1.TextMatrix(0, 2) = "Commodity Group "
MSHFlexGrid1.TextMatrix(0, 3) = "Status "

'MSHFlexGrid2.ColWidth(0) = 0
MSHFlexGrid1.ColWidth(0) = 1800
MSHFlexGrid1.ColWidth(1) = 3000
MSHFlexGrid1.ColWidth(2) = 3000

MSHFlexGrid1.RowHeight(0) = 600
MSHFlexGrid1.WordWrap = True

End Sub


Private Sub AddCmd_Click()
Edit_Flg = "A"


TxtYSRCID = ""
CmbLocationID.Visible = True
CmdLocationID.Visible = True
TxtLocationID.Visible = False
CmbLocationID.Text = ""

CmbCommodityID.Visible = True
CmdCommodityID.Visible = True
TxtCommodityID.Visible = False
CmbCommodityID.Text = ""

CmbStatus.Visible = True
TxtStatus.Visible = False
CmbStatus.Text = "Active"


Call Grid_Head1
Call ClearData
Check1.Value = 0
mLocationID = Null
mCommodityId = Null
TxtRate = ""
TxtRate.Locked = False
CmdUpdate.Enabled = True
'CmbLocationID.SetFocus

Call GButtonLock(Me, False)

End Sub
Private Function GetMaxYSRCID()
Dim Retval As Double
tmp = "Select max(YSRCID) from Mst_YearlyStorageRateCard "
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetMaxYSRCID = Retval
End Function
Private Function GetMaxYSRCDID()
Dim Retval As Double
tmp = "Select max(YSRCDID) from Mst_YearlyStorageRateCardDetail "
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetMaxYSRCDID = Retval
End Function

Private Sub CmdISINID_Click()
FrmMst_ISIN.Show
End Sub

Private Sub Check1_Click()
If SaveCmd.Enabled = False Then Exit Sub
If Check1.Value = vbChecked Then
   For i = 1 To MSHFlexGrid2.Rows - 2
       MSHFlexGrid2.Row = i
       MSHFlexGrid2.Col = 0
       MSHFlexGrid2.TextMatrix(i, 0) = strChecked
   Next
Else
   For i = 1 To MSHFlexGrid2.Rows - 2
       MSHFlexGrid2.Row = i
       MSHFlexGrid2.TextMatrix(i, 0) = strUnChecked
   Next
End If
End Sub

Private Sub CmbStatus_LostFocus()

If CmbStatus.Text = "" Then Exit Sub
If LCase(CmbStatus.Text) <> "active" And LCase(CmbStatus.Text) <> "deactive" Then
   MsgBox "Invalid Selection!!!"
   CmbStatus.SetFocus
   Exit Sub
End If

End Sub

Private Sub CmdCommodityID_Click()
FrmMst_Commodity.Show
End Sub

Private Sub CmdLocationID_Click()
FrmMst_Location.Show
End Sub

Private Sub CmdUpdate_Click()

If ChkSelection = False Then
   MsgBox "Please select atleast one godown from grid for rate mapping!!!"
   Exit Sub
End If

For i = 1 To MSHFlexGrid2.Rows
    If MSHFlexGrid2.TextMatrix(i, 1) = "" Then Exit For
    If MSHFlexGrid2.TextMatrix(i, 0) = strChecked Then
       MSHFlexGrid2.TextMatrix(i, 7) = Val(TxtRate)
       MSHFlexGrid2.TextMatrix(i, 0) = strUnChecked
    End If
Next

Call ClearData

End Sub
Private Function ChkSelection() As Boolean
Dim Retval As Boolean

Retval = False

For i = 1 To MSHFlexGrid2.Rows - 2
    If MSHFlexGrid2.TextMatrix(i, 0) = strChecked Then
       Retval = True
       Exit For
    End If
    
Next

ChkSelection = Retval



End Function

Private Sub DeleteCmd_Click()

Dim Ans As Variant
On Error GoTo msgError

Ans = MsgBox("Do you really want to DELETE this record??? ", vbYesNo)
If Ans = vbYes Then
   Call DeleteChildTableDetail
   RsMst_YearlyStorageRateCard.Delete
   RsMst_YearlyStorageRateCard.Update
   RsMst_YearlyStorageRateCard.Requery
   If RsMst_YearlyStorageRateCard.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_YearlyStorageRateCard.MoveNext
   If RsMst_YearlyStorageRateCard.EOF() Then
      RsMst_YearlyStorageRateCard.MoveLast
   End If
   Call InData
End If

Exit Sub
msgError:
MsgBox "Child record Present "
RsMst_YearlyStorageRateCard.CancelUpdate


End Sub
Private Sub DeleteChildTableDetail()

tmp = "Delete From Mst_YearlyStorageRateCardDetail Where YSRCID = " & TxtYSRCID & ""
Call TmpTable


End Sub
Private Sub EditCmd_Click()
Edit_Flg = "E"


Call GButtonLock(Me, False)

CmbLocationID.Text = TxtLocationID.Text
CmbCommodityID.Text = TxtCommodityID
CmbStatus.Visible = True
TxtStatus.Visible = False
CmbStatus.Text = TxtStatus

TxtRate.Locked = False
CmdUpdate.Enabled = True


End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub Form_Load()
'Dim FrmLoadAccess() As Boolean
'FrmLoadAccess = GetFrmLoadAccess(Me.Name)
'If FrmLoadAccess(0) = False Then
'   Call GButtonLockAD(Me)
'   Frame1.Enabled = False
'   MsgBox "Access denied !!!!!!!!!"
'   Exit Sub
'End If
'
'Call Grid_Head
'CmbSCommodityID.Text = ""
'CmbSLocationID.Text = ""
'LblTotalRecords.Caption = ""
'
'Frame1.Enabled = True
'Call Grid_Head1
'If Gen_WcLocation = "" Then
'   Call TableMst_YearlyStorageRateCard
'Else
'   Call TableMst_YearlyStorageRateCard(" Where " & Gen_WcLocation)
'End If
'If RsMst_YearlyStorageRateCard.RecordCount = 0 Then
'   Call AddCmd_Click
'   Exit Sub
'End If
'
'
'Call Indata

Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Call GButtonLockAD(Me)
   Frame1.Enabled = False
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If

Call Grid_Head
CmbSCommodityID.Text = ""
CmbSLocationID.Text = ""
LblTotalRecords.Caption = ""

Frame1.Enabled = True
Call Grid_Head1

If Gen_WcLocation = "" Then
   Call TableMst_YearlyStorageRateCard
Else
   Call TableMst_YearlyStorageRateCard(" Where " & Gen_WcLocation)
End If

If RsMst_YearlyStorageRateCard.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If

'If Val(Gen_GodownInspection(0)) > 0 Then
'   RsMst_YearlyStorageRateCard.Filter = "LocationID = " & Val(Gen_GodownInspection(0)) & " And CommodityID = " & Val(Gen_GodownInspection(1))
'   If RsMst_YearlyStorageRateCard.RecordCount > 0 Then
'      RsMst_YearlyStorageRateCard.MoveLast
'      tmp = RsMst_YearlyStorageRateCard!YSRCID
'      RsMst_YearlyStorageRateCard.Filter = ""
'      RsMst_YearlyStorageRateCard.MoveFirst
'      RsMst_YearlyStorageRateCard.Find "YSRCID = " & tmp
'   Else
'      RsMst_YearlyStorageRateCard.Filter = ""
'   End If
'End If

Call InData



End Sub
Private Sub InData()

Call GButtonLock(Me, True)
Check1.Value = 0
Call Grid_Head1

CmbLocationID.Visible = False
CmbCommodityID.Visible = False
CmbStatus.Visible = False

TxtLocationID.Visible = True
TxtCommodityID.Visible = True
CmdLocationID.Visible = False
CmdCommodityID.Visible = False
TxtStatus.Visible = True

mLocationID = RsMst_YearlyStorageRateCard!LocationID
mCommodityId = RsMst_YearlyStorageRateCard!CommodityID

TxtYSRCID.Text = RsMst_YearlyStorageRateCard!YSRCID
TxtLocationID.Text = GetLocationName(RsMst_YearlyStorageRateCard!LocationID)
TxtCommodityID.Text = GetCommodityName(RsMst_YearlyStorageRateCard!CommodityID)

If LCase(RsMst_YearlyStorageRateCard!Flag) = "a" Then
   TxtStatus.Text = "Active"
Else
   TxtStatus.Text = "Deactive"
End If

CmbLocationID.Text = TxtLocationID.Text
CmbCommodityID.Text = TxtCommodityID.Text
CmbStatus.Text = TxtStatus

TxtCreated = IIf(IsNull(RsMst_YearlyStorageRateCard!CreatedBy), "", RsMst_YearlyStorageRateCard!CreatedBy) & " :- " & IIf(IsNull(RsMst_YearlyStorageRateCard!CreatedDate), "", RsMst_YearlyStorageRateCard!CreatedDate)
TxtUpdated = IIf(IsNull(RsMst_YearlyStorageRateCard!UpdatedBy), "", RsMst_YearlyStorageRateCard!UpdatedBy) & " :- " & IIf(IsNull(RsMst_YearlyStorageRateCard!UpdatedDate), "", RsMst_YearlyStorageRateCard!UpdatedDate)

Call InDataGodownDetails

TxtRate = ""
TxtRate.Locked = True
CmdUpdate.Enabled = False
SSTab1.Tab = 0
Check1.Value = 0

End Sub
Private Sub InDataGodownDetails()

'Call TableMst_NCDEXStorageCardRateDetail(" Where YSRCID=" & Val(TxtYSRCID.Text) & "")
''
''tmp = "Select MC.CMPName,MG.GodownNo,MG.Address,MU.UnitType,MG.StartDate,MG.CloseDate,MNSRCD.Rate,MG.GodownID,MNSRCD.NSRCDID "
''tmp = tmp & " From Mst_YearlyStorageRateCardDetail MNSRCD "
''tmp = tmp & " Inner Join Mst_Godown MG On MNSRCD.GodownID = MG.GodownID"
''tmp = tmp & " Inner Join Mst_CMP MC On MG.CMPID = MC.CMPID"
''tmp = tmp & " Inner Join Mst_Location ML On MC.LocationID = ML.LocationID"
''tmp = tmp & " Inner Join Mst_UnitType MU on MG.UnitTypeID = MU.UnitTypeID"
''tmp = tmp & " Where MNSRCD.YSRCID = " & TxtYSRCID & ""
''tmp = tmp & " Order by MC.CMPName,MG.GodownNo"

tmp = "Select MC.CMPName,MG.GodownNo,MG.Address,MU.UnitType,MG.StartDate,MG.CloseDate,MNSRCD.Rate,MG.GodownID,MNSRCD.YSRCDID"
tmp = tmp & " From  Mst_Godown MG"
tmp = tmp & " Left Join Mst_YearlyStorageRateCardDetail MNSRCD  On MNSRCD.GodownID = MG.GodownID And MNSRCD.YSRCID = " & Val(TxtYSRCID.Text) & ""
tmp = tmp & " Inner Join Mst_CMP MC On MG.CMPID = MC.CMPID"
tmp = tmp & " inner Join Mst_Location ML On MC.LocationID = ML.LocationID"
tmp = tmp & " Inner Join Mst_UnitType MU on MG.UnitTypeID = MU.UnitTypeID"
tmp = tmp & " Where ML.LocationName = '" & TxtLocationID.Text & "' And YearlyBillingFlg = '1'"
tmp = tmp & " Order By MNSRCD.Rate Desc,MC.CMPName,MG.GodownNo"

Call TmpTable

If TmpRs.RecordCount > 0 Then
   MSHFlexGrid2.Rows = TmpRs.RecordCount + 2
   For i = 1 To MSHFlexGrid2.Rows - 2
       MSHFlexGrid2.TextMatrix(i, 1) = TmpRs!CMPName
       MSHFlexGrid2.TextMatrix(i, 2) = TmpRs!GodownNo
       MSHFlexGrid2.TextMatrix(i, 3) = TmpRs!Address
       MSHFlexGrid2.TextMatrix(i, 4) = TmpRs!UnitType
       MSHFlexGrid2.TextMatrix(i, 5) = IIf(IsNull(TmpRs!StartDate), "", Format(TmpRs!StartDate, "dd-mmm-yyyy"))
       MSHFlexGrid2.TextMatrix(i, 6) = IIf(IsNull(TmpRs!CloseDate), "", Format(TmpRs!CloseDate, "dd-mmm-yyyy"))
       MSHFlexGrid2.TextMatrix(i, 7) = Format(TmpRs!Rate, "########0.00")
       MSHFlexGrid2.TextMatrix(i, 8) = TmpRs!GodownID
       MSHFlexGrid2.TextMatrix(i, 9) = IIf(IsNull(TmpRs!YSRCDID), "", TmpRs!YSRCDID)
       
       MSHFlexGrid2.Row = i
       MSHFlexGrid2.Col = 0
       MSHFlexGrid2.CellFontName = "Wingdings"
       MSHFlexGrid2.CellFontSize = 12
       MSHFlexGrid2.CellAlignment = flexAlignCenterCenter
       MSHFlexGrid2.Text = strUnChecked
       
       TmpRs.MoveNext
   Next
End If
LblGodownCount.Caption = TmpRs.RecordCount
End Sub

Private Sub CmbLocationID_GotFocus()
tmp = CmbLocationID.Text

If Gen_WcLocation = "" Then
   Call TableMst_Location("")
Else
   Call TableMst_Location("Where " & Gen_WcLocation)
End If

CmbLocationID.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No record found!!! "
   Exit Sub
End If
For i = 1 To RsMst_Location.RecordCount
    CmbLocationID.AddItem RsMst_Location!LocationName
    RsMst_Location.MoveNext
Next
CmbLocationID.Text = tmp
End Sub
'----
Private Sub CmbLocationID_LostFocus()
If CmbLocationID.Text = "" Then
   mLocationID = Null
   Call Grid_Head1
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocationID.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection!!! "
   CmbLocationID.SetFocus
   Exit Sub
End If
If mLocationID <> RsMst_Location!LocationID Then
   Call ClearData
   Call Grid_Head1
   Check1.Value = 0
End If
mLocationID = RsMst_Location!LocationID
Call GetGridData
End Sub

Private Sub ClearData()

TxtCMPID = ""
TxtGodownNo = ""
TxtRate = ""


End Sub

Private Sub GetGridData()

If CmbLocationID.Text = "" Then Exit Sub

Call Grid_Head1

tmp = " Select MC.CMPName,MG.GodownNo,MG.Address,MU.UnitType,MG.StartDate,MG.CloseDate,MG.GodownID from Mst_Godown MG"
tmp = tmp & " Inner Join Mst_CMP MC On MG.CMPID = MC.CMPID"
tmp = tmp & " Inner Join Mst_Location ML On MC.LocationID = ML.LocationID"
tmp = tmp & " Inner Join Mst_UnitType MU on MG.UnitTypeID = MU.UnitTypeID"
tmp = tmp & " Where ML.LocationName = '" & CmbLocationID.Text & "' And MG.YearlyBillingFlg = '1' "
tmp = tmp & " Order by MC.CMPName,MG.GodownNo "

Call TmpTable

x = MSHFlexGrid2.Rows - 1
If TmpRs.RecordCount > 0 Then

   For i = 1 To TmpRs.RecordCount
       
       MSHFlexGrid2.TextMatrix(x, 1) = IIf(IsNull(TmpRs!CMPName), "", TmpRs!CMPName)
       MSHFlexGrid2.TextMatrix(x, 2) = IIf(IsNull(TmpRs!GodownNo), "", TmpRs!GodownNo)
       MSHFlexGrid2.TextMatrix(x, 3) = IIf(IsNull(TmpRs!Address), "", TmpRs!Address)
       MSHFlexGrid2.TextMatrix(x, 4) = IIf(IsNull(TmpRs!UnitType), "", TmpRs!UnitType)
       MSHFlexGrid2.TextMatrix(x, 5) = IIf(IsNull(TmpRs!StartDate), "", Format(TmpRs!StartDate, "dd-mmm-yyyy"))
       MSHFlexGrid2.TextMatrix(x, 6) = IIf(IsNull(TmpRs!CloseDate), "", Format(TmpRs!CloseDate, "dd-mmm-yyyy"))
       MSHFlexGrid2.TextMatrix(x, 8) = IIf(IsNull(TmpRs!GodownID), "", TmpRs!GodownID) 'Format(MSHFlexGrid2.TextMatrix(x, 5), "dd-mmm-yyyy")

       MSHFlexGrid2.Row = x
       MSHFlexGrid2.Col = 0
       MSHFlexGrid2.CellFontName = "Wingdings"
       MSHFlexGrid2.CellFontSize = 12
       MSHFlexGrid2.CellAlignment = flexAlignCenterCenter
       MSHFlexGrid2.Text = strUnChecked

       x = x + 1
       MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
       TmpRs.MoveNext
   Next

End If
LblGodownCount.Caption = TmpRs.RecordCount
End Sub

Private Sub MSHFlexGrid2_Click()

If SaveCmd.Enabled = False Then Exit Sub
With MSHFlexGrid2
    If .TextMatrix(.Row, 1) = "" Then Exit Sub
    If .Col = 0 Then
        If .TextMatrix(.Row, 0) = strUnChecked Then
           .TextMatrix(.Row, 0) = strChecked
        Else
           .TextMatrix(.Row, 0) = strUnChecked
        End If
    End If
End With

End Sub

Private Sub SaveCmd_Click()

If CmbLocationID.Text = "" Then
   MsgBox "Blank Location Name not allowed !!!"
   CmbLocationID.SetFocus
   Exit Sub
End If

If CmbCommodityID.Text = "" Then
   MsgBox "Blank Commodity Name not allowed !!!"
   CmbCommodityID.SetFocus
   Exit Sub
End If

If CmbStatus.Text = "" Then
   MsgBox "Blank Status not allowed !!!"
   CmbStatus.SetFocus
   Exit Sub
End If

If Edit_Flg = "A" Then
   If RsMst_YearlyStorageRateCard.RecordCount > 0 Then
      RsMst_YearlyStorageRateCard.MoveFirst
   End If
   'RsMst_YearlyStorageRateCard.Filter = "LocationID = " & mLocationID & " And ExchangeTypeID =  " & mExchangeTypeID & " And CommodityID = " & mCommodityId & " And FromDate = '" & Format(TxtFromDate.Value, Gen_DatFormat) & "' And ToDate = '" & Format(TxtToDate.Value, Gen_DatFormat) & "' And Flag = '" & mFlag & "' And BillingCycleID = " & mBillingCycleID & " And DefaultRate = " & ChkDefault.Value & " "
   RsMst_YearlyStorageRateCard.Filter = "LocationID = " & mLocationID & " And CommodityID = " & mCommodityId & " "
   If RsMst_YearlyStorageRateCard.RecordCount > 0 Then
      MsgBox "Dulicate Storage Rate Card Not allowed !!!"
      RsMst_YearlyStorageRateCard.Filter = ""
      Exit Sub
   End If
   RsMst_YearlyStorageRateCard.Filter = ""
   RsMst_YearlyStorageRateCard.AddNew
   RsMst_YearlyStorageRateCard!YSRCID = GetMaxYSRCID
   RsMst_YearlyStorageRateCard!G_UID = GetGUID
   TxtYSRCID.Text = RsMst_YearlyStorageRateCard!YSRCID
Else
   RsMst_YearlyStorageRateCard.MoveFirst
   RsMst_YearlyStorageRateCard.Filter = "LocationID = " & mLocationID & " And CommodityID = " & mCommodityId & " "
   If RsMst_YearlyStorageRateCard.RecordCount > 0 Then
      If RsMst_YearlyStorageRateCard!YSRCID <> Val(TxtYSRCID) Then
         MsgBox "Dulicate Storage Rate Card Not allowed !!!"
         RsMst_YearlyStorageRateCard.Filter = ""
         Exit Sub
      End If
   End If
   RsMst_YearlyStorageRateCard.Filter = ""
   RsMst_YearlyStorageRateCard.MoveFirst
   RsMst_YearlyStorageRateCard.Find " YSRCID= " & TxtYSRCID.Text
End If

RsMst_YearlyStorageRateCard!YSRCID = TxtYSRCID.Text
RsMst_YearlyStorageRateCard!LocationID = mLocationID
RsMst_YearlyStorageRateCard!CommodityID = mCommodityId
RsMst_YearlyStorageRateCard!Flag = Left(CmbStatus.Text, 1)

If IsNull(RsMst_YearlyStorageRateCard!CreatedBy) = True Or RsMst_YearlyStorageRateCard!CreatedBy = "" Then
   RsMst_YearlyStorageRateCard!CreatedBy = Gen_UserLoginID
   RsMst_YearlyStorageRateCard!CreatedDate = Now
Else
   RsMst_YearlyStorageRateCard!UpdatedBy = Gen_UserLoginID
   RsMst_YearlyStorageRateCard!UpdatedDate = Now
End If

RsMst_YearlyStorageRateCard.Update

Call SaveYearlyStorageRateCardDetails

Call GButtonLock(Me, True)

RsMst_YearlyStorageRateCard.MoveFirst
RsMst_YearlyStorageRateCard.Find "YSRCID = " & TxtYSRCID & ""

Call InData

End Sub
Private Sub SaveYearlyStorageRateCardDetails()

Call TableMst_YearlyStorageRateCardDetail(" Where YSRCID = " & TxtYSRCID & "")

For i = 1 To MSHFlexGrid2.Rows - 1
    If MSHFlexGrid2.TextMatrix(i, 1) <> "" Then
       If MSHFlexGrid2.TextMatrix(i, 9) <> "" Then
          If RsMst_YearlyStorageRateCardDetail.RecordCount > 0 Then
             RsMst_YearlyStorageRateCardDetail.MoveFirst
          End If
          RsMst_YearlyStorageRateCardDetail.Find "YSRCDID = " & MSHFlexGrid2.TextMatrix(i, 9) & ""
       Else
          RsMst_YearlyStorageRateCardDetail.AddNew
          RsMst_YearlyStorageRateCardDetail!YSRCDID = GetMaxYSRCDID
       End If
       RsMst_YearlyStorageRateCardDetail!YSRCID = Val(TxtYSRCID)
       RsMst_YearlyStorageRateCardDetail!GodownID = MSHFlexGrid2.TextMatrix(i, 8)
       RsMst_YearlyStorageRateCardDetail!Rate = Val(MSHFlexGrid2.TextMatrix(i, 7))
       RsMst_YearlyStorageRateCardDetail.Update
       RsMst_YearlyStorageRateCardDetail.Requery
    End If
Next

End Sub

Private Sub SearchCmd_Click()

If SearchCmd.Caption = "Cancel" Then
   If RsMst_YearlyStorageRateCard.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   If RsMst_YearlyStorageRateCard.EOF Then
      RsMst_YearlyStorageRateCard.MoveFirst
   End If
   Call InData
   Exit Sub
End If

Frame0.Height = Me.Height - 500
Frame0.Left = Frame1.Left
Frame0.Top = Frame1.Top
Frame0.Width = Me.Width - 300
Frame0.Visible = True


MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 1400

''===========comment on 7 july 2012 for not clearing search grid=============
'Call Grid_Head
'CmbSCommodityID.Text = ""
'CmbSLocationID.Text = ""
'LblTotalRecords.Caption = ""
''===========comment on 7 july 2012 for not clearing search grid=============

End Sub

Private Sub TxtRate_Lostfocus()

If TxtRate = "" Then Exit Sub
If IsNumeric(TxtRate) = False Or Val(TxtRate) <= 0 Then
   MsgBox "Invalid numeric format!!!"
   TxtRate.SetFocus
   Exit Sub
End If

TxtRate = Round(TxtRate, 4)

End Sub

Private Sub CmbCommodityID_GotFocus()

tmp = CmbCommodityID.Text
Call TableMst_Commodity(" Where YearlyBillingFlg = '1' ")
CmbCommodityID.Clear
If RsMst_Commodity.RecordCount = 0 Then
  MsgBox "No Record found !!!! "
  Exit Sub
End If
For i = 1 To RsMst_Commodity.RecordCount
   CmbCommodityID.AddItem RsMst_Commodity!Commodity
   RsMst_Commodity.MoveNext
Next
CmbCommodityID.Text = tmp
    
End Sub

Private Sub CmbCommodityID_LostFocus()
If CmbCommodityID.Text = "" Then
   mCommodityId = Null
   Exit Sub
End If

RsMst_Commodity.MoveFirst
RsMst_Commodity.Find "Commodity = '" & CmbCommodityID.Text & "'"
If RsMst_Commodity.EOF Then
   MsgBox "Invalid selection "
   CmbCommodityID.SetFocus
   Exit Sub
End If
mCommodityId = RsMst_Commodity!CommodityID

End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_YearlyStorageRateCard.MoveNext
If RsMst_YearlyStorageRateCard.EOF Then
   RsMst_YearlyStorageRateCard.MoveLast
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
RsMst_YearlyStorageRateCard.MovePrevious
If RsMst_YearlyStorageRateCard.BOF Then
   RsMst_YearlyStorageRateCard.MoveFirst
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
RsMst_YearlyStorageRateCard.MoveFirst
Call InData
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsMst_YearlyStorageRateCard.MoveLast
Call InData
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False
End Sub

Private Sub CmbSLocationID_GotFocus()
tmp = CmbSLocationID.Text
If Gen_WcLocation = "" Then
   Call TableMst_Location
Else
   Call TableMst_Location(" Where " & Gen_WcLocation)
End If

CmbSLocationID.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_Location.RecordCount
    CmbSLocationID.AddItem RsMst_Location!LocationName
    RsMst_Location.MoveNext
Next
CmbSLocationID.Text = tmp
    
End Sub

'----
Private Sub CmbSLocationID_LostFocus()
If CmbSLocationID.Text = "" Then
   Exit Sub
End If

RsMst_Location.MoveFirst
RsMst_Location.Find "Locationname = '" & CmbSLocationID.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection "
   CmbSLocationID.SetFocus
   Exit Sub
End If
mSLocationID = RsMst_Location!LocationID
End Sub

Private Sub CmbSCommodityID_GotFocus()

tmp = CmbSCommodityID.Text
Call TableMst_Commodity(" Where YearlyBillingFlg = '1' ")
CmbSCommodityID.Clear
If RsMst_Commodity.RecordCount = 0 Then
  MsgBox "No Record found !!!! "
  Exit Sub
End If
For i = 1 To RsMst_Commodity.RecordCount
    CmbSCommodityID.AddItem RsMst_Commodity!Commodity
    RsMst_Commodity.MoveNext
Next
CmbSCommodityID.Text = tmp
    
End Sub

'----
Private Sub CmbSCommodityID_LostFocus()
If CmbSCommodityID.Text = "" Then
   Exit Sub
End If

RsMst_Commodity.MoveFirst
RsMst_Commodity.Find "Commodity = '" & CmbSCommodityID.Text & "'"
If RsMst_Commodity.EOF Then
   MsgBox "Invalid selection "
   CmbSCommodityID.SetFocus
   Exit Sub
End If

End Sub

Private Sub CmdGo_Click()
Dim wc As String
Call Grid_Head
wc = ""
If CmbSLocationID.Text <> "" Then
   wc = GenWc(wc, CmbSLocationID.Text, "ML.LocationName", "S", "N")
Else
   If Gen_WcLocation <> "" Then
      tmp = Replace(Gen_WcLocation, "LocationID", "ML.LocationID")
      wc = GenWc(wc, tmp, "", "S", "I")
   End If
End If
If CmbSCommodityID.Text <> "" Then
   wc = GenWc(wc, CmbSCommodityID.Text, "Commodity", "S", "N")
End If

tmp = "Select MNSRC.YSRCID,ML.LocationName,MCG.Commodity,MNSRC.Flag From Mst_YearlyStorageRateCard MNSRC"
tmp = tmp & " Inner Join Mst_Location ML On MNSRC.LocationID = ML.LocationID "
tmp = tmp & " Inner Join Mst_Commodity MCG On MNSRC.CommodityID = MCG.CommodityID "
tmp = tmp & wc

Call TmpTable

If TmpRs.RecordCount > 0 Then
   MSHFlexGrid1.Rows = TmpRs.RecordCount + 2
   For i = 1 To TmpRs.RecordCount
       For c = 0 To TmpRs.Fields.Count - 1
           MSHFlexGrid1.TextMatrix(i, c) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
       Next
       If LCase(MSHFlexGrid1.TextMatrix(i, 3)) = "a" Then
          MSHFlexGrid1.TextMatrix(i, 3) = "Active"
       Else
          MSHFlexGrid1.TextMatrix(i, 3) = "Deactive"
       End If
       
       TmpRs.MoveNext
       
   Next
Else
   MsgBox "No Record Found"
End If
LblTotalRecords.Caption = TmpRs.RecordCount
End Sub


Private Sub MSHFlexGrid1_Click()
Dim i As Variant
i = MSHFlexGrid1.RowSel
If i <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(i, 0) = "" Then Exit Sub
RsMst_YearlyStorageRateCard.MoveFirst
RsMst_YearlyStorageRateCard.Find "YSRCID = " & MSHFlexGrid1.TextMatrix(i, 0) & ""
If Not RsMst_YearlyStorageRateCard.EOF Then
   Call InData
   Frame0.Visible = False
End If
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Mst_YearlyStorageRateCard.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Mst_YearlyStorageRateCard.xls")
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
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub


