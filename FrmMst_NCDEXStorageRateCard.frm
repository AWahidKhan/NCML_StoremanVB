VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmMst_NCDEXStorageRateCard 
   Caption         =   "NCDEX Storage Rate Card"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   2295
      Left            =   600
      TabIndex        =   37
      Top             =   3120
      Visible         =   0   'False
      Width           =   11280
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
         TabIndex        =   14
         Top             =   240
         Width           =   2865
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
         TabIndex        =   15
         Top             =   240
         Width           =   2880
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
         TabIndex        =   16
         Top             =   240
         Width           =   1000
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
         Left            =   10725
         TabIndex        =   17
         Top             =   240
         Width           =   1000
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1095
         Left            =   120
         TabIndex        =   38
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
         TabIndex        =   42
         Top             =   315
         Width           =   960
      End
      Begin VB.Label Label2 
         Caption         =   "Commodity Group ID"
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
         TabIndex        =   41
         Top             =   300
         Width           =   1935
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
         TabIndex        =   40
         Top             =   690
         Width           =   2340
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
         TabIndex        =   39
         Top             =   690
         Width           =   90
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1200
      Left            =   120
      TabIndex        =   23
      Top             =   8400
      Width           =   15045
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
         TabIndex        =   13
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
         Left            =   3120
         TabIndex        =   5
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
         TabIndex        =   12
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
         TabIndex        =   6
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
         TabIndex        =   11
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
         TabIndex        =   7
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
         TabIndex        =   10
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
         TabIndex        =   8
         Top             =   240
         Width           =   2880
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
         Left            =   11745
         TabIndex        =   9
         Top             =   645
         Width           =   2880
      End
   End
   Begin VB.Frame Frame1 
      Height          =   8355
      Left            =   105
      TabIndex        =   18
      Top             =   15
      Width           =   15030
      Begin TabDlg.SSTab SSTab1 
         Height          =   6930
         Left            =   120
         TabIndex        =   29
         Top             =   1320
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   12224
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Godown Detail"
         TabPicture(0)   =   "FrmMst_NCDEXStorageRateCard.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame2"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Log History"
         TabPicture(1)   =   "FrmMst_NCDEXStorageRateCard.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "TxtUpdated"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "TxtCreated"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "Label12"
         Tab(1).Control(3)=   "Label13"
         Tab(1).ControlCount=   4
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
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   2025
            Width           =   11175
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
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   1560
            Width           =   11175
         End
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
            TabIndex        =   30
            Top             =   330
            Width           =   14565
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
               TabIndex        =   3
               Top             =   255
               Width           =   1815
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
               TabIndex        =   4
               Top             =   262
               Width           =   1680
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
               Height          =   5670
               Left            =   120
               TabIndex        =   19
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
               TabIndex        =   44
               Top             =   345
               Width           =   90
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
               TabIndex        =   43
               Top             =   345
               Width           =   2415
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
               TabIndex        =   32
               Top             =   315
               Width           =   525
            End
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
            TabIndex        =   36
            Top             =   2085
            Width           =   885
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
            TabIndex        =   35
            Top             =   1620
            Width           =   855
         End
      End
      Begin VB.TextBox TxtNSRCID 
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
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   240
         Width           =   3240
      End
      Begin VB.ComboBox CmbCommodityGroupID 
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
         TabIndex        =   2
         Top             =   720
         Width           =   5760
      End
      Begin VB.CommandButton CmdCommodityGroupID 
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
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   735
         Width           =   375
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
         TabIndex        =   24
         Top             =   750
         Width           =   375
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
         TabIndex        =   1
         Top             =   720
         Width           =   4455
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
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   720
         Width           =   4440
      End
      Begin VB.TextBox TxtCommodityGroupID 
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
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   720
         Width           =   5760
      End
      Begin VB.Label LblWHRNo 
         Caption         =   "NSCR ID"
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
         TabIndex        =   28
         Top             =   315
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "Commodity Group"
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
         TabIndex        =   21
         Top             =   750
         Width           =   1920
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
         TabIndex        =   20
         Top             =   750
         Width           =   1440
      End
   End
End
Attribute VB_Name = "FrmMst_NCDEXStorageRateCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mLocationID, mCommodityGroupId, Edit_Flg, mSLocationID As Variant
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
MSHFlexGrid1.Cols = 3

MSHFlexGrid1.FixedRows = 1
MSHFlexGrid1.FixedCols = 1

MSHFlexGrid1.TextMatrix(0, 0) = "NSRC ID"
MSHFlexGrid1.TextMatrix(0, 1) = "Location Name"
MSHFlexGrid1.TextMatrix(0, 2) = "Commodity Group "

'MSHFlexGrid2.ColWidth(0) = 0
MSHFlexGrid1.ColWidth(0) = 1800
MSHFlexGrid1.ColWidth(1) = 3000
MSHFlexGrid1.ColWidth(2) = 3000

MSHFlexGrid1.RowHeight(0) = 600
MSHFlexGrid1.WordWrap = True

End Sub


Private Sub AddCmd_Click()
Edit_Flg = "A"


TxtNSRCID = ""
CmbLocationID.Visible = True
CmdLocationID.Visible = True
TxtLocationID.Visible = False
CmbLocationID.Text = ""

CmbCommodityGroupID.Visible = True
CmdCommodityGroupID.Visible = True
TxtCommodityGroupID.Visible = False
CmbCommodityGroupID.Text = ""

Call Grid_Head1
Call ClearData
Check1.Value = 0
mLocationID = Null
mCommodityGroupId = Null
TxtRate = ""
TxtRate.Locked = False
CmdUpdate.Enabled = True
'CmbLocationID.SetFocus

Call GButtonLock(Me, False)

End Sub
Private Function GetMaxNSRCID()
Dim Retval As Double
tmp = "Select max(NSRCID) from Mst_NCDEXStorageRateCard "
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetMaxNSRCID = Retval
End Function
Private Function GetMaxNSRCDID()
Dim Retval As Double
tmp = "Select max(NSRCDID) from Mst_NCDEXStorageRateCardDetail "
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetMaxNSRCDID = Retval
End Function

Private Sub CmdISINID_Click()
FrmMst_ISIN.Show
End Sub

Private Sub Check1_Click()
If SaveCmd.Enabled = False Then Exit Sub
If Check1.Value = vbChecked Then
   For I = 1 To MSHFlexGrid2.Rows - 2
       MSHFlexGrid2.Row = I
       MSHFlexGrid2.Col = 0
       MSHFlexGrid2.TextMatrix(I, 0) = strChecked
   Next
Else
   For I = 1 To MSHFlexGrid2.Rows - 2
       MSHFlexGrid2.Row = I
       MSHFlexGrid2.TextMatrix(I, 0) = strUnChecked
   Next
End If
End Sub

Private Sub CmdCommodityGroupID_Click()
FrmMst_CommodityGroup.Show
End Sub

Private Sub CmdLocationID_Click()
FrmMst_Location.Show
End Sub

Private Sub CmdUpdate_Click()

If ChkSelection = False Then
   MsgBox "Please select atleast one godown from grid for rate mapping!!!"
   Exit Sub
End If

For I = 1 To MSHFlexGrid2.Rows
    If MSHFlexGrid2.TextMatrix(I, 1) = "" Then Exit For
    If MSHFlexGrid2.TextMatrix(I, 0) = strChecked Then
       MSHFlexGrid2.TextMatrix(I, 7) = Val(TxtRate)
       MSHFlexGrid2.TextMatrix(I, 0) = strUnChecked
    End If
Next

Call ClearData

End Sub
Private Function ChkSelection() As Boolean
Dim Retval As Boolean

Retval = False

For I = 1 To MSHFlexGrid2.Rows - 2
    If MSHFlexGrid2.TextMatrix(I, 0) = strChecked Then
       Retval = True
       Exit For
    End If
    
Next

ChkSelection = Retval



End Function

Private Sub DeleteCmd_Click()

Dim ans As Variant
On Error GoTo msgError

ans = MsgBox("Do you really want to DELETE this record??? ", vbYesNo)
If ans = vbYes Then
   Call DeleteChildTableDetail
   RsMst_NCDEXStorageRateCard.Delete
   RsMst_NCDEXStorageRateCard.Update
   RsMst_NCDEXStorageRateCard.Requery
   If RsMst_NCDEXStorageRateCard.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_NCDEXStorageRateCard.MoveNext
   If RsMst_NCDEXStorageRateCard.EOF() Then
      RsMst_NCDEXStorageRateCard.MoveLast
   End If
   Call Indata
End If

Exit Sub
msgError:
MsgBox "Child record Present "
RsMst_NCDEXStorageRateCard.CancelUpdate


End Sub
Private Sub DeleteChildTableDetail()

tmp = "Delete From Mst_NCDEXStorageRateCardDetail Where NSRCID = " & TxtNSRCID & ""
Call TmpTable


End Sub
Private Sub EditCmd_Click()
Edit_Flg = "E"


Call GButtonLock(Me, False)

CmbLocationID.Text = TxtLocationID.Text
CmbCommodityGroupID.Text = TxtCommodityGroupID

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
'   Call TableMst_NCDEXStorageRateCard
'Else
'   Call TableMst_NCDEXStorageRateCard(" Where " & Gen_WcLocation)
'End If
'If RsMst_NCDEXStorageRateCard.RecordCount = 0 Then
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
   Call TableMst_NCDEXStorageRateCard
Else
   Call TableMst_NCDEXStorageRateCard(" Where " & Gen_WcLocation)
End If
If RsMst_NCDEXStorageRateCard.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If

If Val(Gen_GodownInspection(0)) > 0 Then
   RsMst_NCDEXStorageRateCard.Filter = "LocationID = " & Val(Gen_GodownInspection(0)) & " And CommodityGroupID = " & Val(Gen_GodownInspection(1))
   If RsMst_NCDEXStorageRateCard.RecordCount > 0 Then
      RsMst_NCDEXStorageRateCard.MoveLast
      tmp = RsMst_NCDEXStorageRateCard!NSRCID
      RsMst_NCDEXStorageRateCard.Filter = ""
      RsMst_NCDEXStorageRateCard.MoveFirst
      RsMst_NCDEXStorageRateCard.Find "NSRCID = " & tmp
   Else
      RsMst_NCDEXStorageRateCard.Filter = ""
   End If
End If

Call Indata



End Sub
Private Sub Indata()

Call GButtonLock(Me, True)

Call Grid_Head1

CmbLocationID.Visible = False
CmbCommodityGroupID.Visible = False

TxtLocationID.Visible = True
TxtCommodityGroupID.Visible = True
CmdLocationID.Visible = False
CmdCommodityGroupID.Visible = False

mLocationID = RsMst_NCDEXStorageRateCard!LocationID
mCommodityGroupId = RsMst_NCDEXStorageRateCard!CommodityGroupID

TxtNSRCID.Text = RsMst_NCDEXStorageRateCard!NSRCID
TxtLocationID.Text = GetLocationName(RsMst_NCDEXStorageRateCard!LocationID)
TxtCommodityGroupID.Text = GetCommodityGroupName(RsMst_NCDEXStorageRateCard!CommodityGroupID)

CmbLocationID.Text = TxtLocationID.Text
CmbCommodityGroupID.Text = TxtCommodityGroupID.Text

TxtCreated = IIf(IsNull(RsMst_NCDEXStorageRateCard!CreatedBy), "", RsMst_NCDEXStorageRateCard!CreatedBy) & " :- " & IIf(IsNull(RsMst_NCDEXStorageRateCard!CreatedDate), "", RsMst_NCDEXStorageRateCard!CreatedDate)
TxtUpdated = IIf(IsNull(RsMst_NCDEXStorageRateCard!UpdatedBy), "", RsMst_NCDEXStorageRateCard!UpdatedBy) & " :- " & IIf(IsNull(RsMst_NCDEXStorageRateCard!UpdatedDate), "", RsMst_NCDEXStorageRateCard!UpdatedDate)

Call InDataGodownDetails
TxtRate = ""
TxtRate.Locked = True
CmdUpdate.Enabled = False
SSTab1.Tab = 0
Check1.Value = 0

End Sub
Private Sub InDataGodownDetails()

'Call TableMst_NCDEXStorageCardRateDetail(" Where NSRCID=" & Val(TxtNSRCID.Text) & "")
''
''tmp = "Select MC.CMPName,MG.GodownNo,MG.Address,MU.UnitType,MG.StartDate,MG.CloseDate,MNSRCD.Rate,MG.GodownID,MNSRCD.NSRCDID "
''tmp = tmp & " From Mst_NCDEXStorageRateCardDetail MNSRCD "
''tmp = tmp & " Inner Join Mst_Godown MG On MNSRCD.GodownID = MG.GodownID"
''tmp = tmp & " Inner Join Mst_CMP MC On MG.CMPID = MC.CMPID"
''tmp = tmp & " Inner Join Mst_Location ML On MC.LocationID = ML.LocationID"
''tmp = tmp & " Inner Join Mst_UnitType MU on MG.UnitTypeID = MU.UnitTypeID"
''tmp = tmp & " Where MNSRCD.NSRCID = " & TxtNSRCID & ""
''tmp = tmp & " Order by MC.CMPName,MG.GodownNo"

tmp = "Select MC.CMPName,MG.GodownNo,MG.Address,MU.UnitType,MG.StartDate,MG.CloseDate,MNSRCD.Rate,MG.GodownID,MNSRCD.NSRCDID"
tmp = tmp & " From  Mst_Godown MG"
tmp = tmp & " Left Join Mst_NCDEXStorageRateCardDetail MNSRCD  On MNSRCD.GodownID = MG.GodownID And MNSRCD.NSRCID = " & Val(TxtNSRCID.Text) & ""
tmp = tmp & " Inner Join Mst_CMP MC On MG.CMPID = MC.CMPID"
tmp = tmp & " inner Join Mst_Location ML On MC.LocationID = ML.LocationID"
tmp = tmp & " Inner Join Mst_UnitType MU on MG.UnitTypeID = MU.UnitTypeID"
tmp = tmp & " Where ML.LocationName = '" & TxtLocationID.Text & "'"
tmp = tmp & " Order By MNSRCD.Rate Desc,MC.CMPName,MG.GodownNo"

Call TmpTable

If TmpRs.RecordCount > 0 Then
   MSHFlexGrid2.Rows = TmpRs.RecordCount + 2
   For I = 1 To MSHFlexGrid2.Rows - 2
       MSHFlexGrid2.TextMatrix(I, 1) = TmpRs!CMPName
       MSHFlexGrid2.TextMatrix(I, 2) = TmpRs!GodownNo
       MSHFlexGrid2.TextMatrix(I, 3) = TmpRs!Address
       MSHFlexGrid2.TextMatrix(I, 4) = TmpRs!UnitType
       MSHFlexGrid2.TextMatrix(I, 5) = IIf(IsNull(TmpRs!StartDate), "", Format(TmpRs!StartDate, "dd-mmm-yyyy"))
       MSHFlexGrid2.TextMatrix(I, 6) = IIf(IsNull(TmpRs!CloseDate), "", Format(TmpRs!CloseDate, "dd-mmm-yyyy"))
       MSHFlexGrid2.TextMatrix(I, 7) = Format(TmpRs!Rate, "########0.00")
       MSHFlexGrid2.TextMatrix(I, 8) = TmpRs!godownid
       MSHFlexGrid2.TextMatrix(I, 9) = IIf(IsNull(TmpRs!NSRCDID), "", TmpRs!NSRCDID)
       
       MSHFlexGrid2.Row = I
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
For I = 1 To RsMst_Location.RecordCount
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
tmp = tmp & " Where ML.LocationName = '" & CmbLocationID.Text & "'"
tmp = tmp & " Order by MC.CMPName,MG.GodownNo "

Call TmpTable

x = MSHFlexGrid2.Rows - 1
If TmpRs.RecordCount > 0 Then

   For I = 1 To TmpRs.RecordCount
       
       MSHFlexGrid2.TextMatrix(x, 1) = IIf(IsNull(TmpRs!CMPName), "", TmpRs!CMPName)
       MSHFlexGrid2.TextMatrix(x, 2) = IIf(IsNull(TmpRs!GodownNo), "", TmpRs!GodownNo)
       MSHFlexGrid2.TextMatrix(x, 3) = IIf(IsNull(TmpRs!Address), "", TmpRs!Address)
       MSHFlexGrid2.TextMatrix(x, 4) = IIf(IsNull(TmpRs!UnitType), "", TmpRs!UnitType)
       MSHFlexGrid2.TextMatrix(x, 5) = IIf(IsNull(TmpRs!StartDate), "", Format(TmpRs!StartDate, "dd-mmm-yyyy"))
       MSHFlexGrid2.TextMatrix(x, 6) = IIf(IsNull(TmpRs!CloseDate), "", Format(TmpRs!CloseDate, "dd-mmm-yyyy"))
       MSHFlexGrid2.TextMatrix(x, 8) = IIf(IsNull(TmpRs!godownid), "", TmpRs!godownid) 'Format(MSHFlexGrid2.TextMatrix(x, 5), "dd-mmm-yyyy")

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

If CmbCommodityGroupID.Text = "" Then
   MsgBox "Blank Commodity Name not allowed !!!"
   CmbCommodityGroupID.SetFocus
   Exit Sub
End If

If Edit_Flg = "A" Then
   If RsMst_NCDEXStorageRateCard.RecordCount > 0 Then
      RsMst_NCDEXStorageRateCard.MoveFirst
   End If
   'RsMst_NCDEXStorageRateCard.Filter = "LocationID = " & mLocationID & " And ExchangeTypeID =  " & mExchangeTypeID & " And CommodityGroupID = " & mCommodityGroupId & " And FromDate = '" & Format(TxtFromDate.Value, Gen_DatFormat) & "' And ToDate = '" & Format(TxtToDate.Value, Gen_DatFormat) & "' And Flag = '" & mFlag & "' And BillingCycleID = " & mBillingCycleID & " And DefaultRate = " & ChkDefault.Value & " "
   RsMst_NCDEXStorageRateCard.Filter = "LocationID = " & mLocationID & " And CommodityGroupID = " & mCommodityGroupId & " "
   If RsMst_NCDEXStorageRateCard.RecordCount > 0 Then
      MsgBox "Dulicate Storage Rate Card Not allowed !!!"
      RsMst_NCDEXStorageRateCard.Filter = ""
      Exit Sub
   End If
   RsMst_NCDEXStorageRateCard.Filter = ""
   RsMst_NCDEXStorageRateCard.AddNew
   RsMst_NCDEXStorageRateCard!NSRCID = GetMaxNSRCID
   RsMst_NCDEXStorageRateCard!G_UID = GetGUID
   TxtNSRCID.Text = RsMst_NCDEXStorageRateCard!NSRCID
Else
   RsMst_NCDEXStorageRateCard.MoveFirst
   RsMst_NCDEXStorageRateCard.Filter = "LocationID = " & mLocationID & " And CommodityGroupID = " & mCommodityGroupId & " "
   If RsMst_NCDEXStorageRateCard.RecordCount > 0 Then
      If RsMst_NCDEXStorageRateCard!NSRCID <> TxtNSRCID Then
         MsgBox "Dulicate Storage Rate Card Not allowed !!!"
         RsMst_NCDEXStorageRateCard.Filter = ""
         Exit Sub
      End If
   End If
   RsMst_NCDEXStorageRateCard.Filter = ""
   RsMst_NCDEXStorageRateCard.MoveFirst
   RsMst_NCDEXStorageRateCard.Find " NSRCID= " & TxtNSRCID.Text
End If

RsMst_NCDEXStorageRateCard!NSRCID = TxtNSRCID.Text
RsMst_NCDEXStorageRateCard!LocationID = mLocationID
RsMst_NCDEXStorageRateCard!CommodityGroupID = mCommodityGroupId

If IsNull(RsMst_NCDEXStorageRateCard!CreatedBy) = True Or RsMst_NCDEXStorageRateCard!CreatedBy = "" Then
   RsMst_NCDEXStorageRateCard!CreatedBy = Gen_UserLoginID
   RsMst_NCDEXStorageRateCard!CreatedDate = Now
Else
   RsMst_NCDEXStorageRateCard!UpdatedBy = Gen_UserLoginID
   RsMst_NCDEXStorageRateCard!UpdatedDate = Now
End If

RsMst_NCDEXStorageRateCard.Update

Call SaveNCDEXStorageRateCardDetails

Call GButtonLock(Me, True)

RsMst_NCDEXStorageRateCard.MoveFirst
RsMst_NCDEXStorageRateCard.Find "NSRCID = " & TxtNSRCID & ""

Call Indata

End Sub
Private Sub SaveNCDEXStorageRateCardDetails()

Call TableMst_NCDEXStorageRateCardDetail(" Where NSRCID = " & TxtNSRCID & "")

For I = 1 To MSHFlexGrid2.Rows - 1
    If MSHFlexGrid2.TextMatrix(I, 1) <> "" Then
       If MSHFlexGrid2.TextMatrix(I, 9) <> "" Then
          If RsMst_NCDEXStorageRateCardDetail.RecordCount > 0 Then
             RsMst_NCDEXStorageRateCardDetail.MoveFirst
          End If
          RsMst_NCDEXStorageRateCardDetail.Find "NSRCDID = " & MSHFlexGrid2.TextMatrix(I, 9) & ""
       Else
          RsMst_NCDEXStorageRateCardDetail.AddNew
          RsMst_NCDEXStorageRateCardDetail!NSRCDID = GetMaxNSRCDID
       End If
       RsMst_NCDEXStorageRateCardDetail!NSRCID = Val(TxtNSRCID)
       RsMst_NCDEXStorageRateCardDetail!godownid = MSHFlexGrid2.TextMatrix(I, 8)
       RsMst_NCDEXStorageRateCardDetail!Rate = Val(MSHFlexGrid2.TextMatrix(I, 7))
       RsMst_NCDEXStorageRateCardDetail.Update
       RsMst_NCDEXStorageRateCardDetail.Requery
    End If
Next

End Sub

Private Sub SearchCmd_Click()

If SearchCmd.Caption = "Cancel" Then
   If RsMst_NCDEXStorageRateCard.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   If RsMst_NCDEXStorageRateCard.EOF Then
      RsMst_NCDEXStorageRateCard.MoveFirst
   End If
   Call Indata
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

Private Sub TxtRate_LostFocus()

If TxtRate = "" Then Exit Sub
If IsNumeric(TxtRate) = False Or Val(TxtRate) <= 0 Then
   MsgBox "Invalid numeric format!!!"
   TxtRate.SetFocus
   Exit Sub
End If

TxtRate = Round(TxtRate, 4)

End Sub

Private Sub CmbCommodityGroupID_GotFocus()

tmp = CmbCommodityGroupID.Text
Call TableMst_CommodityGroup
CmbCommodityGroupID.Clear
If RsMst_CommodityGroup.RecordCount = 0 Then
  MsgBox "No Record found !!!! "
  Exit Sub
End If
For I = 1 To RsMst_CommodityGroup.RecordCount
   CmbCommodityGroupID.AddItem RsMst_CommodityGroup!CommodityGroup
   RsMst_CommodityGroup.MoveNext
Next
CmbCommodityGroupID.Text = tmp
    
End Sub

Private Sub CmbCommodityGroupID_LostFocus()
If CmbCommodityGroupID.Text = "" Then
   mCommodityGroupId = Null
   Exit Sub
End If

RsMst_CommodityGroup.MoveFirst
RsMst_CommodityGroup.Find "CommodityGroup = '" & CmbCommodityGroupID.Text & "'"
If RsMst_CommodityGroup.EOF Then
   MsgBox "Invalid selection "
   CmbCommodityGroupID.SetFocus
   Exit Sub
End If
mCommodityGroupId = RsMst_CommodityGroup!CommodityGroupID

End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_NCDEXStorageRateCard.MoveNext
If RsMst_NCDEXStorageRateCard.EOF Then
   RsMst_NCDEXStorageRateCard.MoveLast
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
RsMst_NCDEXStorageRateCard.MovePrevious
If RsMst_NCDEXStorageRateCard.BOF Then
   RsMst_NCDEXStorageRateCard.MoveFirst
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
RsMst_NCDEXStorageRateCard.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsMst_NCDEXStorageRateCard.MoveLast
Call Indata
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
For I = 1 To RsMst_Location.RecordCount
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
Call TableMst_CommodityGroup
CmbSCommodityID.Clear
If RsMst_CommodityGroup.RecordCount = 0 Then
  MsgBox "No Record found !!!! "
  Exit Sub
End If
For I = 1 To RsMst_CommodityGroup.RecordCount
    CmbSCommodityID.AddItem RsMst_CommodityGroup!CommodityGroup
    RsMst_CommodityGroup.MoveNext
Next
CmbSCommodityID.Text = tmp
    
End Sub

'----
Private Sub CmbSCommodityID_LostFocus()
If CmbSCommodityID.Text = "" Then
   Exit Sub
End If

RsMst_CommodityGroup.MoveFirst
RsMst_CommodityGroup.Find "CommodityGroup = '" & CmbSCommodityID.Text & "'"
If RsMst_CommodityGroup.EOF Then
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
   wc = GenWc(wc, CmbSCommodityID.Text, "CommodityGroup", "S", "N")
End If

tmp = "Select MNSRC.NSRCID,ML.LocationName,MCG.CommodityGroup From Mst_NCDEXStorageRateCard MNSRC"
tmp = tmp & " Inner Join Mst_Location ML On MNSRC.LocationID = ML.LocationID "
tmp = tmp & " Inner Join Mst_CommodityGroup MCG On MNSRC.CommodityGroupID = MCG.CommodityGroupID "
tmp = tmp & wc

Call TmpTable

If TmpRs.RecordCount > 0 Then
   MSHFlexGrid1.Rows = TmpRs.RecordCount + 2
   For I = 1 To TmpRs.RecordCount
       For c = 0 To TmpRs.Fields.Count - 1
           MSHFlexGrid1.TextMatrix(I, c) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
       Next
       TmpRs.MoveNext
       
   Next
Else
   MsgBox "No Record Found"
End If
LblTotalRecords.Caption = TmpRs.RecordCount
End Sub


Private Sub MSHFlexGrid1_Click()
Dim I As Variant
I = MSHFlexGrid1.RowSel
If I <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(I, 0) = "" Then Exit Sub
RsMst_NCDEXStorageRateCard.MoveFirst
RsMst_NCDEXStorageRateCard.Find "NSRCID = " & MSHFlexGrid1.TextMatrix(I, 0) & ""
If Not RsMst_NCDEXStorageRateCard.EOF Then
   Call Indata
   Frame0.Visible = False
End If
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Mst_NCDEXStorageRateCard.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Mst_NCDEXStorageRateCard.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = I + 1 ' + 6
    For c = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(I, c), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(I, c)) Then
              oWS.Cells(mRow, c + 1) = Format(MSHFlexGrid1.TextMatrix(I, c), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(I, c)
           End If
        Else
           oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(I, c)
        End If
    Next
Next
oWB.Save
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub

