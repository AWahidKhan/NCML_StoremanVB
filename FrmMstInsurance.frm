VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmMstInsurance 
   Caption         =   "Insurance Policy Master"
   ClientHeight    =   9060
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15975
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   12915
   ScaleWidth      =   21360
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   46
      Top             =   9000
      Visible         =   0   'False
      Width           =   16260
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   495
         Left            =   120
         TabIndex        =   88
         Top             =   1680
         Width           =   16095
         _ExtentX        =   28390
         _ExtentY        =   873
         _Version        =   393216
         SelectionMode   =   1
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.ComboBox CmbSStatus 
         Appearance      =   0  'Flat
         Height          =   360
         ItemData        =   "FrmMstInsurance.frx":0000
         Left            =   9100
         List            =   "FrmMstInsurance.frx":000A
         Sorted          =   -1  'True
         TabIndex        =   85
         Top             =   1200
         Width           =   2415
      End
      Begin VB.ComboBox CmbSInsuranceByID 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   4560
         Sorted          =   -1  'True
         TabIndex        =   84
         Top             =   1200
         Width           =   4440
      End
      Begin VB.ComboBox CmbSInsuranceRiskTypeID 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   165
         Sorted          =   -1  'True
         TabIndex        =   83
         Top             =   1200
         Width           =   4320
      End
      Begin VB.TextBox TxtSPolicyNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   165
         TabIndex        =   80
         Top             =   500
         Width           =   4320
      End
      Begin VB.ComboBox CmbSInsuranceCo 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   9100
         Sorted          =   -1  'True
         TabIndex        =   82
         Top             =   500
         Width           =   7000
      End
      Begin VB.TextBox TxtSCoverNoteNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   4560
         TabIndex        =   81
         Top             =   500
         Width           =   4440
      End
      Begin VB.CommandButton CmdGo 
         Caption         =   "Go"
         Height          =   375
         Left            =   12930
         TabIndex        =   86
         Top             =   1205
         Width           =   1620
      End
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   375
         Left            =   14550
         TabIndex        =   87
         Top             =   1200
         Width           =   1620
      End
      Begin VB.Label Label34 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Status"
         Height          =   240
         Left            =   10005
         TabIndex        =   92
         Top             =   960
         Width           =   585
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Insurance By"
         Height          =   240
         Left            =   6210
         TabIndex        =   91
         Top             =   960
         Width           =   1185
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Policy No."
         Height          =   240
         Left            =   1875
         TabIndex        =   89
         Top             =   220
         Width           =   915
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Insurance Risk Type"
         Height          =   240
         Left            =   1320
         TabIndex        =   72
         Top             =   960
         Width           =   1845
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Insurance Co."
         Height          =   240
         Left            =   12015
         TabIndex        =   71
         Top             =   220
         Width           =   1215
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Cover Note No."
         Height          =   240
         Left            =   6082
         TabIndex        =   70
         Top             =   220
         Width           =   1380
      End
      Begin VB.Label Label12 
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
         Left            =   14835
         TabIndex        =   68
         Top             =   945
         Width           =   90
      End
      Begin VB.Label Label11 
         Caption         =   "Total Record Found :"
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
         Left            =   12525
         TabIndex        =   67
         Top             =   930
         Width           =   2295
      End
   End
   Begin VB.Frame ButtonFrm 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   64
      Top             =   7920
      Width           =   16260
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New "
         Height          =   375
         Left            =   75
         TabIndex        =   49
         Top             =   195
         Width           =   3270
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   375
         Left            =   3345
         TabIndex        =   50
         Top             =   195
         Width           =   3270
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   6615
         TabIndex        =   51
         Top             =   195
         Width           =   3270
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   9885
         TabIndex        =   52
         Top             =   195
         Width           =   3030
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "&Search"
         Height          =   375
         Left            =   12915
         TabIndex        =   53
         Top             =   195
         Width           =   3150
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   375
         Left            =   75
         TabIndex        =   54
         Top             =   580
         Width           =   3270
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   375
         Left            =   3345
         TabIndex        =   55
         Top             =   580
         Width           =   3270
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   375
         Left            =   6615
         TabIndex        =   57
         Top             =   580
         Width           =   3270
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   375
         Left            =   9885
         TabIndex        =   58
         Top             =   580
         Width           =   3030
      End
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   375
         Left            =   12915
         TabIndex        =   61
         Top             =   580
         Width           =   3150
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7800
      Left            =   120
      TabIndex        =   47
      Top             =   120
      Width           =   16260
      Begin VB.TextBox TxtInsurancePolicyID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   360
         Left            =   120
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   450
         Width           =   1920
      End
      Begin VB.TextBox TxtGUID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   360
         Left            =   14520
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1800
         Visible         =   0   'False
         Width           =   1560
      End
      Begin VB.TextBox TxtStructureWiseLimit 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   3720
         TabIndex        =   13
         Top             =   1850
         Width           =   3480
      End
      Begin VB.ComboBox CmbInsuranceByID 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   12480
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   450
         Width           =   3720
      End
      Begin VB.ComboBox CmbInsuranceRiskTypeID 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   7320
         Sorted          =   -1  'True
         TabIndex        =   8
         Top             =   1185
         Width           =   5040
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   5415
         Left            =   120
         TabIndex        =   17
         Top             =   2280
         Width           =   15975
         _ExtentX        =   28178
         _ExtentY        =   9551
         _Version        =   393216
         Tabs            =   5
         TabsPerRow      =   5
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Godown Type And Locations"
         TabPicture(0)   =   "FrmMstInsurance.frx":0020
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label36"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "MSHFlexGrid7"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Frame6"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Frame7"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).ControlCount=   4
         TabCaption(1)   =   "Godown Detail"
         TabPicture(1)   =   "FrmMstInsurance.frx":003C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "ProgressBar1"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Frame2"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "MSHFlexGrid2"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "Label37"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).ControlCount=   4
         TabCaption(2)   =   "Commodity Detail"
         TabPicture(2)   =   "FrmMstInsurance.frx":0058
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Frame3"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "Frame4"
         Tab(2).Control(1).Enabled=   0   'False
         Tab(2).Control(2)=   "MSHFlexGrid3"
         Tab(2).Control(2).Enabled=   0   'False
         Tab(2).Control(3)=   "MSHFlexGrid4"
         Tab(2).Control(3).Enabled=   0   'False
         Tab(2).Control(4)=   "Label21"
         Tab(2).Control(4).Enabled=   0   'False
         Tab(2).Control(5)=   "Line1"
         Tab(2).Control(5).Enabled=   0   'False
         Tab(2).ControlCount=   6
         TabCaption(3)   =   "Client Detail"
         TabPicture(3)   =   "FrmMstInsurance.frx":0074
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Frame5"
         Tab(3).Control(0).Enabled=   0   'False
         Tab(3).Control(1)=   "MSHFlexGrid5"
         Tab(3).Control(1).Enabled=   0   'False
         Tab(3).Control(2)=   "Label28"
         Tab(3).Control(2).Enabled=   0   'False
         Tab(3).ControlCount=   3
         TabCaption(4)   =   "Policy Remark && Log History"
         TabPicture(4)   =   "FrmMstInsurance.frx":0090
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "TxtRemark"
         Tab(4).Control(0).Enabled=   0   'False
         Tab(4).Control(1)=   "TxtCreated"
         Tab(4).Control(1).Enabled=   0   'False
         Tab(4).Control(2)=   "TxtUpdated"
         Tab(4).Control(2).Enabled=   0   'False
         Tab(4).Control(3)=   "MSHFlexGrid8"
         Tab(4).Control(3).Enabled=   0   'False
         Tab(4).Control(4)=   "Label4"
         Tab(4).Control(4).Enabled=   0   'False
         Tab(4).Control(5)=   "Label10"
         Tab(4).Control(5).Enabled=   0   'False
         Tab(4).Control(6)=   "Label9"
         Tab(4).Control(6).Enabled=   0   'False
         Tab(4).ControlCount=   7
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   75
            Left            =   -74880
            TabIndex        =   120
            Top             =   5040
            Width           =   15735
            _ExtentX        =   27755
            _ExtentY        =   132
            _Version        =   393216
            Appearance      =   0
            Scrolling       =   1
         End
         Begin VB.Frame Frame7 
            Caption         =   "Location Details"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Left            =   6600
            TabIndex        =   109
            Top             =   360
            Width           =   9255
            Begin VB.CheckBox ChkAddAllRegion 
               Caption         =   "Add All Region"
               Height          =   255
               Left            =   120
               TabIndex        =   18
               Top             =   720
               Width           =   1695
            End
            Begin VB.ComboBox CmbValidType 
               Appearance      =   0  'Flat
               Height          =   360
               ItemData        =   "FrmMstInsurance.frx":00AC
               Left            =   1680
               List            =   "FrmMstInsurance.frx":00B9
               Sorted          =   -1  'True
               TabIndex        =   20
               Top             =   225
               Width           =   1935
            End
            Begin VB.CommandButton CmdAddtoListFrame7 
               Caption         =   "Add To List"
               Height          =   375
               Left            =   5520
               TabIndex        =   22
               Top             =   650
               Width           =   1815
            End
            Begin VB.CommandButton CmdDeleteFrame7 
               Caption         =   "Delete From List"
               Height          =   375
               Left            =   7320
               TabIndex        =   23
               Top             =   650
               Width           =   1815
            End
            Begin VB.ComboBox CmbValidTypeID 
               Appearance      =   0  'Flat
               Height          =   360
               ItemData        =   "FrmMstInsurance.frx":00D6
               Left            =   4920
               List            =   "FrmMstInsurance.frx":00D8
               Sorted          =   -1  'True
               TabIndex        =   21
               Top             =   225
               Width           =   4215
            End
            Begin VB.Label Label19 
               Caption         =   "Name"
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
               Left            =   3720
               TabIndex        =   111
               Top             =   285
               Width           =   1335
            End
            Begin VB.Label Label18 
               Caption         =   "Location Type"
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
               TabIndex        =   110
               Top             =   285
               Width           =   1575
            End
         End
         Begin VB.Frame Frame2 
            Height          =   1050
            Left            =   -74880
            TabIndex        =   106
            Top             =   360
            Width           =   15735
            Begin VB.CommandButton CmdGodownList 
               Caption         =   "Get Godown List"
               Height          =   375
               Left            =   4440
               TabIndex        =   117
               Top             =   600
               Width           =   1815
            End
            Begin VB.CheckBox ChkSelectAll 
               Caption         =   "Select All"
               Height          =   255
               Left            =   120
               TabIndex        =   29
               Top             =   660
               Width           =   1215
            End
            Begin VB.CommandButton CmdGetFile 
               Caption         =   "Get File"
               Height          =   375
               Left            =   6240
               Style           =   1  'Graphical
               TabIndex        =   25
               Top             =   600
               Width           =   1245
            End
            Begin VB.CommandButton CmdImport 
               Caption         =   "Import"
               Height          =   375
               Left            =   7485
               TabIndex        =   26
               Top             =   600
               Width           =   1245
            End
            Begin VB.CommandButton CmdValidate 
               Caption         =   "Validate"
               Height          =   375
               Left            =   8760
               TabIndex        =   27
               Top             =   600
               Width           =   1245
            End
            Begin VB.TextBox TxtAttachment 
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               Height          =   360
               Left            =   1440
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   107
               TabStop         =   0   'False
               Top             =   180
               Width           =   14160
            End
            Begin VB.CommandButton CmdExportGodownList 
               Caption         =   "Export"
               Height          =   375
               Left            =   10005
               TabIndex        =   28
               Top             =   600
               Width           =   1245
            End
            Begin MSComDlg.CommonDialog CDialog 
               Left            =   0
               Top             =   0
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
            End
            Begin VB.Label Label39 
               AutoSize        =   -1  'True
               Caption         =   "-"
               Height          =   240
               Left            =   14040
               TabIndex        =   119
               Top             =   660
               Width           =   75
            End
            Begin VB.Label Label38 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               Caption         =   "No of Connected Godowns :"
               Height          =   240
               Left            =   11460
               TabIndex        =   118
               Top             =   660
               Width           =   2505
            End
            Begin VB.Label Label35 
               AutoSize        =   -1  'True
               Caption         =   "-"
               Height          =   240
               Left            =   3150
               TabIndex        =   113
               Top             =   667
               Width           =   75
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               Caption         =   "Total No of Rows :"
               Height          =   240
               Left            =   1470
               TabIndex        =   112
               Top             =   667
               Width           =   1665
            End
            Begin VB.Label Label72 
               Caption         =   "File Name"
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
               TabIndex        =   108
               Top             =   240
               Width           =   1095
            End
         End
         Begin VB.Frame Frame3 
            Height          =   1095
            Left            =   -74880
            TabIndex        =   102
            Top             =   360
            Width           =   7815
            Begin VB.ComboBox CmbIncludeCommodityID 
               Appearance      =   0  'Flat
               Height          =   360
               Left            =   1440
               Sorted          =   -1  'True
               TabIndex        =   31
               Top             =   240
               Width           =   6255
            End
            Begin VB.CommandButton CmdAddFrame3 
               Caption         =   "Add To List"
               Height          =   375
               Left            =   4080
               TabIndex        =   32
               Top             =   650
               Width           =   1815
            End
            Begin VB.CommandButton CmdDeleteFrame3 
               Caption         =   "Delete From List"
               Height          =   375
               Left            =   5880
               TabIndex        =   33
               Top             =   650
               Width           =   1815
            End
            Begin VB.Label Label25 
               Caption         =   "Commodity Covered Under Policy"
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
               TabIndex        =   104
               Top             =   720
               Width           =   3615
            End
            Begin VB.Label Label22 
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
               TabIndex        =   103
               Top             =   293
               Width           =   1335
            End
         End
         Begin VB.Frame Frame4 
            Height          =   1095
            Left            =   -66900
            TabIndex        =   99
            Top             =   360
            Width           =   7695
            Begin VB.CommandButton CmdDeleteFrame4 
               Caption         =   "Delete From List"
               Height          =   375
               Left            =   5880
               TabIndex        =   37
               Top             =   650
               Width           =   1695
            End
            Begin VB.CommandButton CmdAddFrame4 
               Caption         =   "Add To List"
               Height          =   375
               Left            =   4320
               TabIndex        =   36
               Top             =   650
               Width           =   1575
            End
            Begin VB.ComboBox CmbExcludeCommodityID 
               Appearance      =   0  'Flat
               Height          =   360
               Left            =   1440
               Sorted          =   -1  'True
               TabIndex        =   35
               Top             =   240
               Width           =   6135
            End
            Begin VB.Label Label27 
               Caption         =   "Commodity Not Covered Under Policy"
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
               TabIndex        =   101
               Top             =   720
               Width           =   4215
            End
            Begin VB.Label Label26 
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
               TabIndex        =   100
               Top             =   293
               Width           =   1335
            End
         End
         Begin VB.Frame Frame5 
            Height          =   735
            Left            =   -74880
            TabIndex        =   95
            Top             =   360
            Width           =   15735
            Begin VB.ComboBox CmbExchange 
               Height          =   360
               ItemData        =   "FrmMstInsurance.frx":00DA
               Left            =   1200
               List            =   "FrmMstInsurance.frx":00DC
               Sorted          =   -1  'True
               TabIndex        =   38
               Top             =   600
               Visible         =   0   'False
               Width           =   1530
            End
            Begin VB.TextBox TxtClientID 
               Appearance      =   0  'Flat
               Height          =   360
               Left            =   1200
               TabIndex        =   40
               Top             =   240
               Width           =   2040
            End
            Begin VB.ComboBox CmbClientRowID 
               Appearance      =   0  'Flat
               Height          =   360
               Left            =   4080
               Sorted          =   -1  'True
               TabIndex        =   41
               Top             =   240
               Width           =   8415
            End
            Begin VB.CommandButton CmdAddFrame5 
               Caption         =   "Add To List"
               Height          =   375
               Left            =   12600
               TabIndex        =   42
               Top             =   233
               Width           =   1455
            End
            Begin VB.CommandButton CmdDeleteFrame5 
               Caption         =   "Delete From List"
               Height          =   375
               Left            =   14040
               TabIndex        =   43
               Top             =   233
               Width           =   1575
            End
            Begin VB.Label Label40 
               Caption         =   "Exchange"
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
               TabIndex        =   121
               Top             =   540
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.Label Label31 
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
               Height          =   255
               Left            =   120
               TabIndex        =   97
               Top             =   300
               Width           =   975
            End
            Begin VB.Label Label30 
               Caption         =   "Name "
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
               TabIndex        =   96
               Top             =   300
               Width           =   1335
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "Godown Type"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4695
            Left            =   120
            TabIndex        =   94
            Top             =   360
            Width           =   6375
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid6 
               Height          =   4335
               Left            =   120
               TabIndex        =   19
               Top             =   240
               Width           =   6135
               _ExtentX        =   10821
               _ExtentY        =   7646
               _Version        =   393216
               Cols            =   3
               AllowUserResizing=   1
               _NumberOfBands  =   1
               _Band(0).Cols   =   3
            End
         End
         Begin VB.TextBox TxtRemark 
            Appearance      =   0  'Flat
            Height          =   4335
            Left            =   -73680
            MaxLength       =   4000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   45
            Top             =   960
            Width           =   14565
         End
         Begin VB.TextBox TxtCreated 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   360
            Left            =   -73680
            Locked          =   -1  'True
            TabIndex        =   74
            TabStop         =   0   'False
            Top             =   510
            Width           =   6360
         End
         Begin VB.TextBox TxtUpdated 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   360
            Left            =   -65640
            Locked          =   -1  'True
            TabIndex        =   73
            TabStop         =   0   'False
            Top             =   510
            Width           =   6480
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid5 
            Height          =   3855
            Left            =   -74880
            TabIndex        =   44
            Top             =   1200
            Width           =   15735
            _ExtentX        =   27755
            _ExtentY        =   6800
            _Version        =   393216
            AllowUserResizing=   1
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
            Height          =   3495
            Left            =   -74880
            TabIndex        =   34
            Top             =   1440
            Width           =   7815
            _ExtentX        =   13785
            _ExtentY        =   6165
            _Version        =   393216
            AllowUserResizing=   1
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid4 
            Height          =   3495
            Left            =   -66900
            TabIndex        =   39
            Top             =   1440
            Width           =   7695
            _ExtentX        =   13573
            _ExtentY        =   6165
            _Version        =   393216
            AllowUserResizing=   1
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
            Height          =   3645
            Left            =   -74880
            TabIndex        =   30
            Top             =   1395
            Width           =   15735
            _ExtentX        =   27755
            _ExtentY        =   6429
            _Version        =   393216
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid7 
            Height          =   3600
            Left            =   6600
            TabIndex        =   24
            Top             =   1455
            Width           =   9255
            _ExtentX        =   16325
            _ExtentY        =   6350
            _Version        =   393216
            Cols            =   3
            AllowUserResizing=   1
            _NumberOfBands  =   1
            _Band(0).Cols   =   3
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid8 
            Height          =   3855
            Left            =   -73680
            TabIndex        =   114
            Top             =   1440
            Visible         =   0   'False
            Width           =   15735
            _ExtentX        =   27755
            _ExtentY        =   6800
            _Version        =   393216
            AllowBigSelection=   0   'False
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin VB.Label Label37 
            Alignment       =   2  'Center
            Caption         =   $"FrmMstInsurance.frx":00DE
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   -74880
            TabIndex        =   116
            Top             =   5100
            Width           =   15600
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label36 
            Alignment       =   2  'Center
            Caption         =   $"FrmMstInsurance.frx":0179
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   120
            TabIndex        =   115
            Top             =   5100
            Width           =   15600
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label21 
            Caption         =   "By default policy will be cover all Commodity stored as per Store Man."
            Height          =   240
            Left            =   -74880
            TabIndex        =   105
            Top             =   4995
            Width           =   7320
            WordWrap        =   -1  'True
         End
         Begin VB.Line Line1 
            BorderWidth     =   2
            X1              =   -66975
            X2              =   -66975
            Y1              =   405
            Y2              =   5280
         End
         Begin VB.Label Label28 
            Caption         =   "This tab will be enable and medatory in case of Insurance by Client. "
            Height          =   240
            Left            =   -74880
            TabIndex        =   98
            Top             =   5080
            Width           =   15600
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label4 
            Caption         =   "Remark"
            Height          =   255
            Left            =   -74760
            TabIndex        =   77
            Top             =   3000
            Width           =   975
         End
         Begin VB.Label Label10 
            Caption         =   "Created  :"
            Height          =   270
            Left            =   -74760
            TabIndex        =   76
            Top             =   555
            Width           =   1695
         End
         Begin VB.Label Label9 
            Caption         =   "Updated :"
            Height          =   270
            Left            =   -66960
            TabIndex        =   75
            Top             =   555
            Width           =   1695
         End
      End
      Begin VB.TextBox TxtCoverNoteNo 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   7320
         MaxLength       =   255
         TabIndex        =   2
         Top             =   450
         Width           =   5040
      End
      Begin VB.ComboBox CmbInsuranceCo 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   1170
         Width           =   6765
      End
      Begin VB.ComboBox CmbStatus 
         Appearance      =   0  'Flat
         Height          =   360
         ItemData        =   "FrmMstInsurance.frx":0214
         Left            =   7320
         List            =   "FrmMstInsurance.frx":021E
         Sorted          =   -1  'True
         TabIndex        =   14
         Top             =   1845
         Width           =   2415
      End
      Begin VB.TextBox TxtG_UID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   375
         Left            =   15585
         Locked          =   -1  'True
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   2280
         Visible         =   0   'False
         Width           =   375
      End
      Begin MSComCtl2.DTPicker ValidFrom 
         Height          =   375
         Left            =   12480
         TabIndex        =   10
         Top             =   1200
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
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
         Format          =   113836033
         CurrentDate     =   39888
      End
      Begin VB.CommandButton CmdInsuranceCo 
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
         Left            =   6840
         TabIndex        =   6
         Top             =   1170
         Width           =   375
      End
      Begin VB.TextBox TxtAmount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   120
         TabIndex        =   12
         Top             =   1850
         Width           =   3495
      End
      Begin VB.TextBox TxtPolicyNo 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   1
         Top             =   450
         Width           =   5040
      End
      Begin VB.TextBox TxtInsuranceCo 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1170
         Width           =   7080
      End
      Begin MSComCtl2.DTPicker ValidTo 
         Height          =   375
         Left            =   14400
         TabIndex        =   11
         Top             =   1200
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
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
         Format          =   113836033
         CurrentDate     =   39888
      End
      Begin VB.TextBox TxtStatus 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   7320
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   1850
         Width           =   2400
      End
      Begin VB.TextBox TxtInsuranceBy 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   12480
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   450
         Width           =   3600
      End
      Begin VB.TextBox TxtInsuranceRiskType 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   7320
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1185
         Width           =   5040
      End
      Begin VB.Label Label29 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Insurance Policy ID"
         Height          =   240
         Left            =   98
         TabIndex        =   93
         Top             =   150
         Width           =   1965
      End
      Begin VB.Label Label32 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Insurance By"
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
         Left            =   13658
         TabIndex        =   90
         Top             =   165
         Width           =   1365
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Structure Wise Limit"
         Height          =   240
         Left            =   4583
         TabIndex        =   79
         Top             =   1560
         Width           =   1755
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Insurance Risk Type"
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
         Left            =   8753
         TabIndex        =   78
         Top             =   885
         Width           =   2175
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
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
         Left            =   8190
         TabIndex        =   69
         Top             =   1560
         Width           =   675
      End
      Begin VB.Label Label8 
         Caption         =   "G_UID"
         Height          =   255
         Left            =   15030
         TabIndex        =   66
         Top             =   2295
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Valid To"
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
         Left            =   14850
         TabIndex        =   63
         Top             =   885
         Width           =   915
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Valid From"
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
         Left            =   12810
         TabIndex        =   62
         Top             =   885
         Width           =   1155
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   1470
         TabIndex        =   60
         Top             =   1560
         Width           =   795
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Insurance Company"
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
         Left            =   2445
         TabIndex        =   59
         Top             =   840
         Width           =   2085
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Cover Note No"
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
         Left            =   9053
         TabIndex        =   56
         Top             =   150
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Policy No"
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
         Left            =   3983
         TabIndex        =   48
         Top             =   150
         Width           =   1395
      End
   End
End
Attribute VB_Name = "FrmMstInsurance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg As Variant
Dim mInsuranceCoID As Double, mSInsuranceCoID As Double
Dim mInsuranceRiskTypeID As Double, mSInsuranceRiskTypeID As Double
Dim mInsuranceByID As Double, mSInsuranceByID As Double
Dim mValidType As Variant, mValidTypeID As Variant
Dim mExcludeValidType As Variant, mExcludeValidTypeID As Variant

Dim mIncludeCommodityID As Double, mExcludeCommodityID As Double
Dim mIncludeCommodityIDList As Variant, mExcludeCommodityIDList As Variant
Dim mClientRowID As Double
Dim mUnitTypeList As Variant, mListOfSelectedUnitType As Variant
Dim mCMPID As Variant, mExcludeCMPID As Variant
Dim mLocationID As Variant, mExcludeLocationID As Variant
Dim mSplit As Variant, mTmpName As Variant

Const strChecked = "þ"
Const strUnChecked = "q"

Dim strFilename As Variant
Dim oFs, oFs1 As New FileSystemObject
Dim oFolder As Folder
Dim oFile As File

Dim mInsuranceByChangeFlag As Boolean, mInsuranceCoChangeFlag As Boolean, mInsuranceRiskTypeChangeFlag As Boolean
Dim mGodownChangeFlag As Boolean, mCommodityChangeFlag As Boolean, mClientChangeFlag As Boolean

Dim mExchangeTypeID As Double


Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)

TxtInsurancePolicyID = ""
TxtPolicyNo = "":               TxtPolicyNo.Locked = False
TxtCoverNoteNo = "":            TxtCoverNoteNo.Locked = False
TxtInsuranceCo = "":            TxtInsuranceCo.Visible = False
CmbInsuranceCo = "":            CmbInsuranceCo.Visible = True: CmdInsuranceCo.Visible = True
TxtInsuranceRiskType = "":      TxtInsuranceRiskType.Visible = False
CmbInsuranceRiskTypeID = "":    CmbInsuranceRiskTypeID.Visible = True
TxtInsuranceBy = "":            TxtInsuranceBy.Visible = False
CmbInsuranceByID = "":          CmbInsuranceByID.Visible = True
TxtAmount = "":                 TxtAmount.Locked = False
TxtStructureWiseLimit = "":     TxtStructureWiseLimit.Locked = False
ValidFrom.Enabled = True:       ValidFrom.Value = Date
ValidTo.Enabled = True:         ValidTo.Value = DateAdd("d", 365, Date)
TxtStatus.Text = "":            TxtStatus.Visible = False
CmbStatus.Text = "Active":      CmbStatus.Visible = True
TxtCreated = ""
TxtUpdated = ""
TxtGUID = ""

''--- Frame 6 Godown Type details
Call Grid_Head6
Call FillUnitType
Dim mR As Double
mListOfSelectedUnitType = 0
With MSHFlexGrid6
    If .Rows > 2 Then
        For mR = 1 To .Rows - 2
           If .TextMatrix(mR, 0) = strChecked Then
            mListOfSelectedUnitType = mListOfSelectedUnitType & "," & Val(.TextMatrix(mR, 2))
           End If
        Next mR
    End If
End With
''--- Frame 7 Location details
Call ClearFrame7
Call Grid_Head7
ChkAddAllRegion.Value = vbUnchecked
Frame7.Enabled = True
''--- Frame 2 Godown Location details
Call ClearFrame2
Call Grid_Head2
Label35.Caption = CountSelectedGodown & " / " & MSHFlexGrid2.Rows - 2
Label39.Caption = CountSelectedConnectedGodown & " / " & MSHFlexGrid2.Rows - 2
ChkSelectAll.Value = vbUnchecked
Frame2.Enabled = True
''--- Frame 3 Include Commodity details
Call ClearFrame3
Call Grid_Head3
Frame3.Enabled = True
''--- Frame 4 Exclude Commodity details
Call ClearFrame4
Call Grid_Head4
Frame4.Enabled = True
''--- Frame 5 Client details
''CmbExchange = ""
Call ClearFrame5
Call Grid_Head5
Frame5.Enabled = True
SSTab1.TabEnabled(3) = False
''----------------------------------
TxtRemark = "":                 TxtRemark.Locked = False
SSTab1.Tab = 0

If RsMst_PolicyDetail.RecordCount > 0 Then
   TxtPolicyNo.SetFocus
Else
   SearchCmd.Enabled = False
End If
End Sub

Private Sub ChkAddAllRegion_Click()
Dim mR As Double, mRow As Double, mCol As Double, mRLastRow As Double
Dim mEntry As Boolean

If ChkAddAllRegion.Tag = "NoClick" Then Exit Sub

If ChkAddAllRegion.Value = vbChecked Then
    tmp = " Select MR.RegionID,MR.Region from Mst_Godown MG"
    tmp = tmp & " Inner Join Mst_CMP CMP On MG.CMPID = CMP.CMPID"
    tmp = tmp & " Inner Join Mst_Location ML On CMP.LocationID = ML.LocationID"
    tmp = tmp & " Inner Join Mst_State MS On ML.StateID = MS.StateID"
    tmp = tmp & " Inner Join Mst_Region MR On MS.RegionID = MR.RegionID"
    tmp = tmp & " Where MG.CloseDate Is Null And MG.UnitTypeID In (" & mListOfSelectedUnitType & ")"
    tmp = tmp & " Group By MR.RegionID,MR.Region"
    Call TmpTable
    If TmpRs.RecordCount > 0 Then
        With MSHFlexGrid7
            ''-- Directly add if the Grid if Blank
            If .Rows = 2 Then
                For mR = 1 To TmpRs.RecordCount
                    .TextMatrix(mR, 1) = "Region"
                    .TextMatrix(mR, 2) = IIf(IsNull(TmpRs!Region), "", TmpRs!Region)
                    .TextMatrix(mR, 3) = IIf(IsNull(TmpRs!RegionID), "", TmpRs!RegionID)
                    .Rows = .Rows + 1
                    TmpRs.MoveNext
                Next mR
            Else
            ''-- Check if grid is not blank then add only that region which is not available
               For mR = 1 To TmpRs.RecordCount
                   mRLastRow = .Rows - 2
                   mEntry = False
                   For mRow = 1 To .Rows - 1
                        If LCase(Left(.TextMatrix(mRow, 1), 1)) = "r" And Val(.TextMatrix(mRow, 3)) = IIf(IsNull(TmpRs!RegionID), "", TmpRs!RegionID) Then ''-- Region
                            mEntry = True
                        End If
                   Next mRow
                   If mEntry = False Then
                   
                    .TextMatrix(mRLastRow + 1, 1) = "Region"
                    .TextMatrix(mRLastRow + 1, 2) = IIf(IsNull(TmpRs!Region), "", TmpRs!Region)
                    .TextMatrix(mRLastRow + 1, 3) = IIf(IsNull(TmpRs!RegionID), "", TmpRs!RegionID)
                    .Rows = .Rows + 1
                   
                   End If
                   TmpRs.MoveNext
               Next mR
            End If
        
        End With
    End If
Else
    With MSHFlexGrid7
        If .Rows > 2 Then
            mRLastRow = .Rows - 2
            For mR = mRLastRow To 1 Step -1
                If LCase(Left(.TextMatrix(mR, 1), 1)) = "r" Then
                    
                    ''-- If Locations are Already mapped and save with insurance then not allow to Edit/Remove
                    If Val(.TextMatrix(mR, 4)) > 0 Then Exit Sub
                    
                    For mRow = mR To .Rows - 2
                        For mCol = 0 To .Cols - 1
                            .TextMatrix(mRow, mCol) = .TextMatrix(mRow + 1, mCol)
                        Next
                    Next
                    .Rows = .Rows - 1
                    
                End If
            Next mR
        
        End If
    End With
End If

End Sub

Private Sub ChkSelectAll_Click()
Dim mR As Double

If SaveCmd.Enabled = False Then Exit Sub
If ChkSelectAll.Tag = "NoClick" Then
    Exit Sub
End If
With MSHFlexGrid2
    If .Rows > 2 Then
       If ChkSelectAll.Value = vbChecked Then
            For mR = 1 To .Rows - 2
                
             ''-- if Cell color is red the not allow to select that godown
'             .row = mR
'             .Col = 0
'             If .CellBackColor <> vbRed Then
'                 .TextMatrix(mR, 0) = strChecked
'             End If
            .row = mR
            .Col = 0
            If .CellBackColor = vbYellow And Val(.TextMatrix(mR, 12)) > 0 Then
                Call GRowColor(MSHFlexGrid2, mR, vbGreen)
                .TextMatrix(mR, 0) = strChecked
            Else
                If .CellBackColor <> vbRed Then
                    .TextMatrix(mR, 0) = strChecked
                    .TextMatrix(mR, 9) = strChecked
                End If
            End If
            
             ''.TextMatrix(mR, 0) = strChecked
            Next mR
       ElseIf ChkSelectAll.Value = vbUnchecked Then
            For mR = 1 To .Rows - 2
            
                .row = mR
                .Col = 0
                If .CellBackColor = vbGreen And Val(.TextMatrix(mR, 12)) > 0 Then
                    'Call GRowColor(MSHFlexGrid2, mR, vbYellow)
                    '.TextMatrix(mR, 0) = strUnChecked
                Else
                    If .CellBackColor <> vbRed Then
                        .TextMatrix(mR, 0) = strUnChecked
                    End If
                End If
            
            Next mR
       End If
       Label35.Caption = CountSelectedGodown & " / " & MSHFlexGrid2.Rows - 2
       Label39.Caption = CountSelectedConnectedGodown & " / " & MSHFlexGrid2.Rows - 2
    End If
End With

End Sub

Private Sub CmbClientRowID_GotFocus()
''If CmbExchange.Text = "" Then
''    MsgBox "Select Exchange."
''    CmbExchange.SetFocus
''    Exit Sub
''End If

tmp = CmbClientRowID.Text

If TxtClientID <> "" Then
    Call TableMst_Client(" Where ExchangeTypeID = " & mExchangeTypeID & " And Status = 'A' And ClientID = '" & TxtClientID & "' ")
Else
    Call TableMst_Client(" Where ExchangeTypeID = " & mExchangeTypeID & " And Status = 'A' ")
End If

CmbClientRowID.Clear
If RsMst_Client.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Client.RecordCount
    CmbClientRowID.AddItem RsMst_Client!ClientName & "#" & RsMst_Client!ClientID & "#" & RsMst_Client!ClientIDRow
    RsMst_Client.MoveNext
Next
CmbClientRowID.Text = tmp
End Sub
Private Sub CmbClientRowID_LostFocus()
If CmbClientRowID.Text = "" Then Exit Sub

mSplit = Split(CmbClientRowID, "#")
If UBound(mSplit) <> 2 Then
   MsgBox "Invalid selection "
   CmbClientRowID.SetFocus
   Exit Sub
End If
RsMst_Client.MoveFirst
RsMst_Client.Find "ClientIDRow = '" & mSplit(2) & "'"
If RsMst_Client.EOF Then
   MsgBox "Invalid selection "
   CmbClientRowID.SetFocus
   Exit Sub
End If
mClientRowID = RsMst_Client!ClientIDRow
TxtClientID = RsMst_Client!ClientID
mTmpName = RsMst_Client!ClientName
End Sub

Private Sub CmbCMPID_GotFocus()
If Trim(CmbValidType.Text) = "" Then
   MsgBox "Select Type!!!"
   CmbValidType.SetFocus
   Exit Sub
End If
If Trim(CmbLocationID.Text) = "" Then
   MsgBox "Select Location!!!"
   CmbLocationID.SetFocus
   Exit Sub
End If

tmp = CmbCMPID.Text
tmp1 = CmbCMPID.Text
Call TableMst_CMP(" Where CMPID In ( Select CMP.CMPID from Mst_Godown MG Inner Join Mst_CMP CMP On MG.CMPID = CMP.CMPID Where MG.CloseDate Is Null And CMP.CloseDate Is Null And CMP.LocationID = " & mLocationID & " And MG.UnitTypeID In (" & mListOfSelectedUnitType & ") Group By CMP.CMPID ) ")
CmbCMPID.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_CMP.RecordCount
    CmbCMPID.AddItem RsMst_CMP!CMPName
    RsMst_CMP.MoveNext
Next
CmbCMPID.Text = tmp
End Sub
Private Sub CmbCMPID_LostFocus()
If CmbCMPID.Text = "" Then CmbValidTypeID.Text = "":  Exit Sub
RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPName = '" & CmbCMPID.Text & "'"
If RsMst_CMP.EOF Then
   MsgBox "Invalid selection "
   CmbCMPID.SetFocus
   Exit Sub
End If
mCMPID = RsMst_CMP!CMPID
If tmp1 <> CmbCMPID.Text Then
    CmbValidTypeID.Text = ""
End If
End Sub

Private Sub CmbExcludeCMPID_GotFocus()
If Trim(CmbExcludeValidType.Text) = "" Then
   MsgBox "Select Type!!!"
   CmbExcludeValidType.SetFocus
   Exit Sub
End If
If Trim(CmbExcludeLocationID.Text) = "" Then
   MsgBox "Select Location!!!"
   CmbExcludeLocationID.SetFocus
   Exit Sub
End If

tmp = CmbExcludeCMPID.Text
tmp1 = CmbExcludeCMPID.Text
Call TableMst_CMP(" Where CMPID In ( Select CMP.CMPID from Mst_Godown MG Inner Join Mst_CMP CMP On MG.CMPID = CMP.CMPID Where MG.CloseDate Is Null And CMP.CloseDate Is Null And CMP.LocationID = " & mExcludeLocationID & " And MG.UnitTypeID In (" & mListOfSelectedUnitType & ") Group By CMP.CMPID ) ")
CmbExcludeCMPID.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_CMP.RecordCount
    CmbExcludeCMPID.AddItem RsMst_CMP!CMPName
    RsMst_CMP.MoveNext
Next
CmbExcludeCMPID.Text = tmp
End Sub

Private Sub CmbExcludeCMPID_LostFocus()
If CmbExcludeCMPID.Text = "" Then CmbExcludeValidTypeID.Text = "": Exit Sub
RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPName = '" & CmbExcludeCMPID.Text & "'"
If RsMst_CMP.EOF Then
   MsgBox "Invalid selection "
   CmbExcludeCMPID.SetFocus
   Exit Sub
End If
mExcludeCMPID = RsMst_CMP!CMPID
If tmp1 <> CmbExcludeCMPID.Text Then
    CmbExcludeValidTypeID.Text = ""
End If
End Sub

Private Sub CmbExchange_GotFocus()
tmp = CmbExchange.Text
Call TableMst_ExchangeType
CmbExchange.Clear
If RsMst_ExchangeType.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_ExchangeType.RecordCount
    CmbExchange.AddItem RsMst_ExchangeType!ExchangeType
    RsMst_ExchangeType.MoveNext
Next
CmbExchange.Text = tmp
End Sub

Private Sub CmbExchange_LostFocus()
If CmbExchange.Text = "" Then
   TxtClientID = ""
   CmbClientRowID = ""
   Exit Sub
End If

RsMst_ExchangeType.MoveFirst
RsMst_ExchangeType.Find "ExchangeType = '" & CmbExchange.Text & "'"
If RsMst_ExchangeType.EOF Then
   MsgBox "Invalid selection "
   CmbExchange.SetFocus
   Exit Sub
End If
mExchangeTypeID = RsMst_ExchangeType!ExchangeTypeID
TxtClientID = ""
CmbClientRowID = ""

End Sub


Private Sub CmbExcludeCommodityID_GotFocus()
If MSHFlexGrid3.Rows > 2 Then
   MsgBox "By default all commodities which are not entered in 'Covered Commodities' List on the left hand side of this screen are excluded from the Insurance Policy. You need not select any commodity name for this purpose."
   CmbIncludeCommodityID.SetFocus
   Exit Sub
End If

tmp = CmbExcludeCommodityID.Text
Call ListIncludeCommodity
Call TableMst_Commodity(" Where CommodityID Not In (" & mIncludeCommodityIDList & ") ")
CmbExcludeCommodityID.Clear
If RsMst_Commodity.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Commodity.RecordCount
 CmbExcludeCommodityID.AddItem RsMst_Commodity!Commodity
 RsMst_Commodity.MoveNext
Next
CmbExcludeCommodityID.Text = tmp
End Sub
Private Sub CmbExcludeCommodityID_LostFocus()
If CmbExcludeCommodityID.Text = "" Then Exit Sub
RsMst_Commodity.MoveFirst
RsMst_Commodity.Find "Commodity = '" & CmbExcludeCommodityID.Text & "'"
If RsMst_Commodity.EOF Then
   MsgBox "Invalid selection "
   CmbExcludeCommodityID.SetFocus
   Exit Sub
End If
mExcludeCommodityID = RsMst_Commodity!CommodityID
CmbIncludeCommodityID.Text = ""
End Sub

Private Sub CmbExcludeLocationID_GotFocus()
If Trim(CmbExcludeValidType.Text) = "" Then
   MsgBox "Select Type!!!"
   CmbExcludeValidType.SetFocus
   Exit Sub
End If
tmp = CmbExcludeLocationID.Text
Call TableMst_Location(" Where LocationID In ( Select ML.LocationID from Mst_Godown MG Inner Join Mst_CMP CMP On MG.CMPID = CMP.CMPID Inner Join Mst_Location ML On CMP.LocationID = ML.LocationID Where MG.CloseDate Is Null And MG.UnitTypeID In (" & mListOfSelectedUnitType & ") Group By ML.LocationID,ML.LocationName ) ")
CmbExcludeLocationID.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Record found!!!"
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
   CmbExcludeLocationID.AddItem RsMst_Location!LocationName
   RsMst_Location.MoveNext
Next
CmbExcludeLocationID.Text = tmp
End Sub

Private Sub CmbExcludeLocationID_LostFocus()
If CmbExcludeLocationID.Text = "" Then
   CmbExcludeValidTypeID.Text = ""
   Exit Sub
End If

RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbExcludeLocationID.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid  selection!!!"
   CmbExcludeLocationID.SetFocus
   Exit Sub
End If

mExcludeLocationID = RsMst_Location!LocationID

If tmp <> CmbExcludeLocationID.Text Then
   CmbExcludeValidTypeID.Text = ""
End If

End Sub

Private Sub CmbExcludeValidType_GotFocus()
If CheckSelectedUnitType = False Then
   MsgBox "Select Godown Type."
   SSTab1.Tab = 0
   Exit Sub
End If
End Sub

Private Sub CmbExcludeValidType_LostFocus()
If CmbExcludeValidType.Text = "" Then Exit Sub

If LCase(CmbExcludeValidType.Text) <> LCase("Region") And LCase(CmbExcludeValidType.Text) <> LCase("State") And LCase(CmbExcludeValidType.Text) <> LCase("Location") And LCase(CmbExcludeValidType.Text) <> LCase("CMP") And LCase(CmbExcludeValidType.Text) <> LCase("Godown") Then
    MsgBox "Invalid Selection!!!"
    CmbExcludeValidType.Text = ""
    CmbExcludeValidTypeID.Text = ""
    CmbExcludeLocationID = "": CmbExcludeLocationID.Visible = False
    CmbExcludeCMPID = "": CmbExcludeCMPID.Visible = False
    lblExcludeLocation.Visible = True
    lblExcludeCMP.Visible = False
    CmbExcludeValidType.SetFocus
    Exit Sub
End If

If LCase(mExcludeValidType) <> LCase(Left$(CmbExcludeValidType, 1)) Then
    CmbExcludeValidTypeID.Text = ""
End If
mExcludeValidType = Left$(CmbExcludeValidType, 1)
Label37.Caption = CmbExcludeValidType.Text

If LCase(mExcludeValidType) = LCase("C") Then
    lblExcludeLocation.Visible = True
    CmbExcludeLocationID = "": CmbExcludeLocationID.Visible = True
    lblExcludeCMP.Visible = False
    CmbExcludeCMPID = "": CmbExcludeCMPID.Visible = False
    CmbExcludeLocationID.SetFocus
ElseIf LCase(mExcludeValidType) = LCase("G") Then
    lblExcludeLocation.Visible = True
    CmbExcludeLocationID = "": CmbExcludeLocationID.Visible = True
    lblExcludeCMP.Visible = True
    CmbExcludeCMPID = "": CmbExcludeCMPID.Visible = True
    CmbExcludeLocationID.SetFocus
Else
    lblExcludeLocation.Visible = False
    CmbExcludeLocationID = "": CmbExcludeLocationID.Visible = False
    lblExcludeCMP.Visible = False
    CmbExcludeCMPID = "": CmbExcludeCMPID.Visible = False
End If

End Sub

Private Sub CmbExcludeValidTypeID_GotFocus()

If MSHFlexGrid2.Rows > 2 Then
   MsgBox "By default all Locations which are not entered in 'Covered Locations' List on the left hand side of this screen are excluded from the Insurance Policy. You need not select any Location name for this purpose."
   CmbExcludeValidTypeID.Clear
   CmbExcludeValidTypeID.SetFocus
   Exit Sub
End If

If CmbExcludeValidType.Text = "" Then
    MsgBox "Select Type First!!!"
    CmbExcludeValidType.SetFocus
    Exit Sub
End If

tmp1 = CmbExcludeValidTypeID.Text

If LCase(CmbExcludeValidType.Text) = LCase("Region") Then ''-- Region
    ''tmp = " Select * from Mst_Region"
    tmp = " Select MR.RegionID,MR.Region from Mst_Godown MG"
    tmp = tmp & " Inner Join Mst_CMP CMP On MG.CMPID = CMP.CMPID"
    tmp = tmp & " Inner Join Mst_Location ML On CMP.LocationID = ML.LocationID"
    tmp = tmp & " Inner Join Mst_State MS On ML.StateID = MS.StateID"
    tmp = tmp & " Inner Join Mst_Region MR On MS.RegionID = MR.RegionID"
    tmp = tmp & " Where MG.CloseDate Is Null And MG.UnitTypeID In (" & mListOfSelectedUnitType & ")"
    tmp = tmp & " Group By MR.RegionID,MR.Region "
    Call TmpTable
    CmbExcludeValidTypeID.Clear
    If TmpRs.RecordCount = 0 Then
       MsgBox "No Record found !!!! "
       Exit Sub
    End If
    For I = 1 To TmpRs.RecordCount
        CmbExcludeValidTypeID.AddItem TmpRs!Region
        TmpRs.MoveNext
    Next
ElseIf LCase(CmbExcludeValidType.Text) = LCase("State") Then  ''-- State
    ''tmp = " Select * from Mst_State"
    tmp = " Select MS.StateID,MS.StateName from Mst_Godown MG"
    tmp = tmp & " Inner Join Mst_CMP CMP On MG.CMPID = CMP.CMPID"
    tmp = tmp & " Inner Join Mst_Location ML On CMP.LocationID = ML.LocationID"
    tmp = tmp & " Inner Join Mst_State MS On ML.StateID = MS.StateID"
    tmp = tmp & " Inner Join Mst_Region MR On MS.RegionID = MR.RegionID"
    tmp = tmp & " Where MG.CloseDate Is Null And MG.UnitTypeID In (" & mListOfSelectedUnitType & ")"
    tmp = tmp & " Group By MS.StateID,MS.StateName"
    Call TmpTable
    CmbExcludeValidTypeID.Clear
    If TmpRs.RecordCount = 0 Then
       MsgBox "No Record found !!!! "
       Exit Sub
    End If
    For I = 1 To TmpRs.RecordCount
        CmbExcludeValidTypeID.AddItem TmpRs!StateName
        TmpRs.MoveNext
    Next
ElseIf LCase(CmbExcludeValidType.Text) = LCase("Location") Then   ''-- Location
    ''tmp = " Select * from Mst_Location "
    tmp = " Select ML.LocationID,ML.LocationName from Mst_Godown MG"
    tmp = tmp & " Inner Join Mst_CMP CMP On MG.CMPID = CMP.CMPID"
    tmp = tmp & " Inner Join Mst_Location ML On CMP.LocationID = ML.LocationID"
    tmp = tmp & " Inner Join Mst_State MS On ML.StateID = MS.StateID"
    tmp = tmp & " Inner Join Mst_Region MR On MS.RegionID = MR.RegionID"
    tmp = tmp & " Where MG.CloseDate Is Null And MG.UnitTypeID In (" & mListOfSelectedUnitType & ")"
    tmp = tmp & " Group By ML.LocationID,ML.LocationName"
    Call TmpTable
    CmbExcludeValidTypeID.Clear
    If TmpRs.RecordCount = 0 Then
       MsgBox "No Record found !!!! "
       Exit Sub
    End If
    For I = 1 To TmpRs.RecordCount
        CmbExcludeValidTypeID.AddItem TmpRs!LocationName
        TmpRs.MoveNext
    Next
ElseIf LCase(CmbExcludeValidType.Text) = LCase("CMP") Then    ''-- CMP
    If CmbExcludeLocationID.Text = "" Then
        MsgBox "Select Location"
        CmbExcludeLocationID.SetFocus
        Exit Sub
    End If
    
    ''tmp = " Select CMPID,CMPName,CMPName+'~'+convert(Varchar,CMPID)[CMP],Address from Mst_CMP Where LocationID = " & mLocationID & " And CloseDate Is Null "
    tmp = " Select CMP.CMPID,CMP.CMPName,CMP.CMPName+'~'+convert(Varchar,CMP.CMPID)[CMP] from Mst_Godown MG"
    tmp = tmp & " Inner Join Mst_CMP CMP On MG.CMPID = CMP.CMPID"
    tmp = tmp & " Inner Join Mst_Location ML On CMP.LocationID = ML.LocationID"
    tmp = tmp & " Inner Join Mst_State MS On ML.StateID = MS.StateID"
    tmp = tmp & " Inner Join Mst_Region MR On MS.RegionID = MR.RegionID"
    tmp = tmp & " Where MG.CloseDate Is Null And CMP.CloseDate Is Null And CMP.LocationID = " & mExcludeLocationID & " And MG.UnitTypeID In (" & mListOfSelectedUnitType & ")"
    tmp = tmp & " Group By CMP.CMPID,CMP.CMPName"
    Call TmpTable
    CmbExcludeValidTypeID.Clear
    If TmpRs.RecordCount = 0 Then
       MsgBox "No Record found !!!! "
       Exit Sub
    End If
    For I = 1 To TmpRs.RecordCount
        CmbExcludeValidTypeID.AddItem TmpRs!CMP
        TmpRs.MoveNext
    Next
ElseIf LCase(CmbExcludeValidType.Text) = LCase("Godown") Then    ''-- Godown
    If CmbExcludeCMPID.Text = "" Then
        MsgBox "Select CMP"
        CmbExcludeCMPID.SetFocus
        Exit Sub
    End If
    
    ''tmp = " Select CMPID,GodownID,GodownNo,GodownNo+'~'+convert(Varchar,GodownID)[Godown],Address from Mst_Godown Where CloseDate Is Null And CMPID = " & mCMPID & ""
    tmp = " Select MG.GodownID,MG.GodownNo,MG.GodownNo+'~'+convert(Varchar,MG.GodownID)[Godown]"
    tmp = tmp & " from Mst_Godown MG"
    tmp = tmp & " Inner Join Mst_CMP CMP On MG.CMPID = CMP.CMPID"
    tmp = tmp & " Inner Join Mst_Location ML On CMP.LocationID = ML.LocationID"
    tmp = tmp & " Inner Join Mst_State MS On ML.StateID = MS.StateID"
    tmp = tmp & " Inner Join Mst_Region MR On MS.RegionID = MR.RegionID"
    tmp = tmp & " Where MG.CloseDate Is Null And MG.CMPID = " & mExcludeCMPID & " And MG.UnitTypeID In (" & mListOfSelectedUnitType & ")"
    tmp = tmp & " Group By MG.GodownID,MG.GodownNo"
    Call TmpTable
    CmbExcludeValidTypeID.Clear
    If TmpRs.RecordCount = 0 Then
       MsgBox "No Record found !!!! "
       Exit Sub
    End If
    For I = 1 To TmpRs.RecordCount
        CmbExcludeValidTypeID.AddItem TmpRs!Godown
        TmpRs.MoveNext
    Next
End If
CmbExcludeValidTypeID.Text = tmp1

End Sub

Private Sub CmbExcludeValidTypeID_LostFocus()
If CmbExcludeValidTypeID.Text = "" Then Exit Sub

If LCase(CmbExcludeValidType.Text) = "region" Then
   If TmpRs.RecordCount > 0 Then TmpRs.MoveFirst
   TmpRs.Find "Region= '" & CmbExcludeValidTypeID.Text & "'"
   If TmpRs.EOF Then
      MsgBox "Invalid selection "
      CmbExcludeValidTypeID.SetFocus
      Exit Sub
   End If
   mExcludeValidTypeID = TmpRs!RegionID
ElseIf LCase(CmbExcludeValidType.Text) = "state" Then
   If TmpRs.RecordCount > 0 Then TmpRs.MoveFirst
   TmpRs.Find "StateName = '" & CmbExcludeValidTypeID.Text & "'"
   If TmpRs.EOF Then
      MsgBox "Invalid selection "
      CmbExcludeValidTypeID.SetFocus
      Exit Sub
   End If
   mExcludeValidTypeID = TmpRs!StateID
ElseIf LCase(CmbExcludeValidType.Text) = "location" Then
   If TmpRs.RecordCount > 0 Then TmpRs.MoveFirst
   TmpRs.Find "LocationName = '" & CmbExcludeValidTypeID.Text & "'"
   If TmpRs.EOF Then
      MsgBox "Invalid selection "
      CmbExcludeValidTypeID.SetFocus
      Exit Sub
   End If
   mExcludeValidTypeID = TmpRs!LocationID
ElseIf LCase(CmbValidType.Text) = "cmp" Then
   If TmpRs.RecordCount > 0 Then TmpRs.MoveFirst
   TmpRs.Find "CMP = '" & CmbExcludeValidTypeID.Text & "'"
   If TmpRs.EOF Then
      MsgBox "Invalid selection "
      CmbExcludeValidTypeID.SetFocus
      Exit Sub
   End If
   mExcludeValidTypeID = TmpRs!CMPID
   'TxtAddress.Visible = True: TxtAddress = IIf(IsNull(TmpRs!Address), "", TmpRs!Address)
ElseIf LCase(CmbExcludeValidType.Text) = "godown" Then
   If TmpRs.RecordCount > 0 Then TmpRs.MoveFirst
   TmpRs.Find "Godown = '" & CmbExcludeValidTypeID.Text & "'"
   If TmpRs.EOF Then
      MsgBox "Invalid selection "
      CmbExcludeValidTypeID.SetFocus
      Exit Sub
   End If
   mExcludeValidTypeID = TmpRs!GodownID
   'TxtAddress.Visible = True: TxtAddress = IIf(IsNull(TmpRs!Address), "", TmpRs!Address)
End If

End Sub

Private Sub CmbIncludeCommodityID_GotFocus()
If MSHFlexGrid4.Rows > 2 Then
   MsgBox "By default all commodities which are not entered in 'Not Covered Commodities' list on the right hand side of this screen are included in the Insurance Policy. You need not select any commodity name for this purpose."
   CmbExcludeCommodityID.SetFocus
   Exit Sub
End If

tmp = CmbIncludeCommodityID.Text
Call ListExcludeCommodity
Call TableMst_Commodity(" Where CommodityID Not In (" & mExcludeCommodityIDList & ") ")
CmbIncludeCommodityID.Clear
If RsMst_Commodity.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Commodity.RecordCount
 CmbIncludeCommodityID.AddItem RsMst_Commodity!Commodity
 RsMst_Commodity.MoveNext
Next
CmbIncludeCommodityID.Text = tmp
End Sub
Private Sub CmbIncludeCommodityID_LostFocus()
If CmbIncludeCommodityID.Text = "" Then Exit Sub
RsMst_Commodity.MoveFirst
RsMst_Commodity.Find "Commodity = '" & CmbIncludeCommodityID.Text & "'"
If RsMst_Commodity.EOF Then
   MsgBox "Invalid selection "
   CmbIncludeCommodityID.SetFocus
   Exit Sub
End If
mIncludeCommodityID = RsMst_Commodity!CommodityID
CmbExcludeCommodityID.Text = ""
End Sub

Private Function ListIncludeCommodity()
Dim mR As Double
mIncludeCommodityIDList = 0
With MSHFlexGrid3
    If .Rows > 2 Then
        For mR = 1 To .Rows - 2
           mIncludeCommodityIDList = mIncludeCommodityIDList & "," & Val(.TextMatrix(mR, 2))
        Next mR
    End If
End With
End Function
Private Function ListExcludeCommodity()
Dim mR As Double
mExcludeCommodityIDList = 0
With MSHFlexGrid4
    If .Rows > 2 Then
        For mR = 1 To .Rows - 2
           mExcludeCommodityIDList = mExcludeCommodityIDList & "," & Val(.TextMatrix(mR, 2))
        Next mR
    End If
End With
End Function

Private Function ListUnitType()
Dim mR As Double
mUnitTypeList = 0
With MSHFlexGrid6
    If .Rows > 2 Then
        For mR = 1 To .Rows - 2
           mUnitTypeList = mUnitTypeList & "," & Val(.TextMatrix(mR, 2))
        Next mR
    End If
End With
End Function

Private Function CheckSelectedUnitType() As Boolean
Dim mR As Double
Dim mResult As Boolean
mResult = False
    
With MSHFlexGrid6
    If .Rows > 2 Then
        For mR = 1 To .Rows - 2
           If .TextMatrix(mR, 0) = strChecked Then
            mResult = True
            Exit For
           End If
        Next mR
    End If
End With

CheckSelectedUnitType = mResult
End Function

Private Function CheckSelectedGodown() As Boolean
Dim mR As Double
Dim mResult As Boolean
mResult = False
    
With MSHFlexGrid2
    If .Rows > 2 Then
        For mR = 1 To .Rows - 2
           If .TextMatrix(mR, 0) = strChecked Then
            mResult = True
            Exit For
           End If
        Next mR
    End If
End With

CheckSelectedGodown = mResult
End Function

Private Function CheckSelectedConnectedGodown() As Boolean
Dim mR As Double
Dim mResult As Boolean
mResult = False
    
With MSHFlexGrid2
    If .Rows > 2 Then
        For mR = 1 To .Rows - 2
           If .TextMatrix(mR, 9) = strChecked Then
            mResult = True
            Exit For
           End If
        Next mR
    End If
End With

CheckSelectedConnectedGodown = mResult
End Function



Private Sub CmbInsuranceByID_GotFocus()
tmp = CmbInsuranceByID.Text
Call TableMst_InsuranceBy(" Where Flag = 'A' ")
CmbInsuranceByID.Clear
If RsMst_InsuranceBy.RecordCount = 0 Then
   MsgBox "No Insurance By found"
   CmbInsuranceByID.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_InsuranceBy.RecordCount
    CmbInsuranceByID.AddItem RsMst_InsuranceBy!InsuranceByName
    RsMst_InsuranceBy.MoveNext
Next
CmbInsuranceByID.Text = tmp
End Sub
Private Sub CmbInsuranceByID_LostFocus()
If CmbInsuranceByID.Text = "" Then Exit Sub

RsMst_InsuranceBy.MoveFirst
RsMst_InsuranceBy.Find "InsuranceByName = '" & CmbInsuranceByID.Text & "'"

If RsMst_InsuranceBy.EOF Then
   MsgBox "Invalid Insurance By selection "
   CmbInsuranceByID.SetFocus
   Exit Sub
End If
mInsuranceByID = RsMst_InsuranceBy!InsuranceByID
If mInsuranceByID = 1 Then
    SSTab1.TabEnabled(3) = False
    Call Grid_Head5
    If SSTab1.Tab = 3 Then SSTab1.Tab = 2
ElseIf mInsuranceByID = 2 Then
    SSTab1.TabEnabled(3) = True
End If

End Sub

Private Sub CmbInsuranceCo_GotFocus()
tmp = CmbInsuranceCo.Text
Call TableMst_InsuranceCo(" Where Flag = 'A' ")
CmbInsuranceCo.Clear
If RsMst_InsuranceCo.RecordCount = 0 Then
   MsgBox "No InsuranceCo found"
   CmdInsuranceCo.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_InsuranceCo.RecordCount
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
End Sub

Private Sub CmbSearchInssurance_GotFocus()
tmp = CmbSearchInssurance.Text
Call TableMst_InsuranceCo("")

CmbSearchInssurance.Clear
If RsMst_InsuranceCo.RecordCount = 0 Then
   MsgBox "No InsuranceCo found"
   CmbSearchInssurance.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_InsuranceCo.RecordCount
    CmbSearchInssurance.AddItem RsMst_InsuranceCo!InsuranceCoName
    RsMst_InsuranceCo.MoveNext
Next
CmbSearchInssurance.Text = tmp
End Sub

Private Sub CmbSearchInssurance_LostFocus()
If CmbSearchInssurance.Text = "" Then
   Exit Sub
End If
RsMst_InsuranceCo.MoveFirst
RsMst_InsuranceCo.Find "InsuranceCoName = '" & CmbSearchInssurance.Text & "'"

If RsMst_InsuranceCo.EOF Then
   MsgBox "Invalid InsuranceCo selection "
   CmbSearchInssurance.SetFocus
   Exit Sub
End If
mSearchInsuranceCoID = RsMst_InsuranceCo.Fields(0)
End Sub


Private Sub CmbInsuranceRiskTypeID_GotFocus()
tmp = CmbInsuranceRiskTypeID.Text
Call TableMst_InsuranceRiskCovered(" Where Status = 'A' ")
CmbInsuranceRiskTypeID.Clear
If RsMst_InsuranceRiskCovered.RecordCount = 0 Then
   MsgBox "No Insurance Risk Type found"
   CmbInsuranceRiskTypeID.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_InsuranceRiskCovered.RecordCount
    CmbInsuranceRiskTypeID.AddItem RsMst_InsuranceRiskCovered!InsuranceRiskCovered
    RsMst_InsuranceRiskCovered.MoveNext
Next
CmbInsuranceRiskTypeID.Text = tmp
End Sub
Private Sub CmbInsuranceRiskTypeID_LostFocus()
If CmbInsuranceRiskTypeID.Text = "" Then Exit Sub

RsMst_InsuranceRiskCovered.MoveFirst
RsMst_InsuranceRiskCovered.Find "InsuranceRiskCovered = '" & CmbInsuranceRiskTypeID.Text & "'"

If RsMst_InsuranceRiskCovered.EOF Then
   MsgBox "Invalid Insurance Risk Type selection "
   CmbInsuranceRiskTypeID.SetFocus
   Exit Sub
End If
mInsuranceRiskTypeID = RsMst_InsuranceRiskCovered!InsuranceRiskCoveredID
End Sub

Private Sub CmbLocationID_GotFocus()
If Trim(CmbValidType.Text) = "" Then
   MsgBox "Select Type!!!"
   CmbValidType.SetFocus
   Exit Sub
End If
tmp = CmbLocationID.Text
Call TableMst_Location(" Where LocationID In ( Select ML.LocationID from Mst_Godown MG Inner Join Mst_CMP CMP On MG.CMPID = CMP.CMPID Inner Join Mst_Location ML On CMP.LocationID = ML.LocationID Where MG.CloseDate Is Null And MG.UnitTypeID In (" & mListOfSelectedUnitType & ") Group By ML.LocationID,ML.LocationName ) ")
CmbLocationID.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Record found!!!"
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
   CmbValidTypeID.Text = ""
   Exit Sub
End If

RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocationID.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid  selection!!!"
   CmbLocationID.SetFocus
   Exit Sub
End If

mLocationID = RsMst_Location!LocationID

If tmp <> CmbLocationID.Text Then
   CmbValidTypeID.Text = ""
End If
End Sub

Private Sub CmbSInsuranceByID_GotFocus()
tmp = CmbSInsuranceByID.Text
Call TableMst_InsuranceBy(" Where Flag = 'A' ")
CmbSInsuranceByID.Clear
If RsMst_InsuranceBy.RecordCount = 0 Then
   MsgBox "No Insurance By found"
   CmbSInsuranceByID.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_InsuranceBy.RecordCount
    CmbSInsuranceByID.AddItem RsMst_InsuranceBy!InsuranceByName
    RsMst_InsuranceBy.MoveNext
Next
CmbSInsuranceByID.Text = tmp
End Sub
Private Sub CmbSInsuranceByID_LostFocus()
If CmbSInsuranceByID.Text = "" Then Exit Sub

RsMst_InsuranceBy.MoveFirst
RsMst_InsuranceBy.Find "InsuranceByName = '" & CmbSInsuranceByID.Text & "'"

If RsMst_InsuranceBy.EOF Then
   MsgBox "Invalid Insurance By selection "
   CmbSInsuranceByID.SetFocus
   Exit Sub
End If
mSInsuranceByID = RsMst_InsuranceBy!InsuranceByID
End Sub

Private Sub CmbSInsuranceCo_GotFocus()
tmp = CmbSInsuranceCo.Text
Call TableMst_InsuranceCo(" Where Flag = 'A' ")
CmbSInsuranceCo.Clear
If RsMst_InsuranceCo.RecordCount = 0 Then
   MsgBox "No InsuranceCo found"
   CmbSInsuranceCo.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_InsuranceCo.RecordCount
    CmbSInsuranceCo.AddItem RsMst_InsuranceCo!InsuranceCoName
    RsMst_InsuranceCo.MoveNext
Next
CmbSInsuranceCo.Text = tmp
End Sub
Private Sub CmbSInsuranceCo_LostFocus()
If CmbSInsuranceCo.Text = "" Then
   Exit Sub
End If
RsMst_InsuranceCo.MoveFirst
RsMst_InsuranceCo.Find "InsuranceCoName = '" & CmbSInsuranceCo.Text & "'"

If RsMst_InsuranceCo.EOF Then
   MsgBox "Invalid InsuranceCo selection "
   CmbSInsuranceCo.SetFocus
   Exit Sub
End If
mSInsuranceCoID = RsMst_InsuranceCo.Fields(0)
End Sub

Private Sub CmbSInsuranceRiskTypeID_GotFocus()
tmp = CmbSInsuranceRiskTypeID.Text
Call TableMst_InsuranceRiskCovered(" Where Status = 'A' ")
CmbSInsuranceRiskTypeID.Clear
If RsMst_InsuranceRiskCovered.RecordCount = 0 Then
   MsgBox "No Insurance Risk Type found"
   CmbSInsuranceRiskTypeID.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_InsuranceRiskCovered.RecordCount
    CmbSInsuranceRiskTypeID.AddItem RsMst_InsuranceRiskCovered!InsuranceRiskCovered
    RsMst_InsuranceRiskCovered.MoveNext
Next
CmbSInsuranceRiskTypeID.Text = tmp
End Sub
Private Sub CmbSInsuranceRiskTypeID_LostFocus()
If CmbSInsuranceRiskTypeID.Text = "" Then Exit Sub

RsMst_InsuranceRiskCovered.MoveFirst
RsMst_InsuranceRiskCovered.Find "InsuranceRiskCovered = '" & CmbSInsuranceRiskTypeID.Text & "'"

If RsMst_InsuranceRiskCovered.EOF Then
   MsgBox "Invalid Insurance Risk Type selection "
   CmbSInsuranceRiskTypeID.SetFocus
   Exit Sub
End If
mSInsuranceRiskTypeID = RsMst_InsuranceRiskCovered!InsuranceRiskCoveredID
End Sub

Private Sub CmbStatus_LostFocus()
If CmbStatus.Text = "" Then Exit Sub
If LCase(CmbStatus.Text) <> LCase("Active") And LCase(CmbStatus.Text) <> LCase("Deactive") Then
   MsgBox "Invalid Selection"
   CmbStatus.SetFocus
   Exit Sub
End If
End Sub
Private Sub CmbSStatus_LostFocus()
If CmbSStatus.Text = "" Then Exit Sub
If LCase(CmbSStatus.Text) <> LCase("Active") And LCase(CmbSStatus.Text) <> LCase("Deactive") Then
   MsgBox "Invalid Selection"
   CmbSStatus.SetFocus
   Exit Sub
End If
End Sub


Private Sub CmbValidType_GotFocus()
If CheckSelectedUnitType = False Then
   MsgBox "Select Godown Type."
   MSHFlexGrid6.SetFocus
   SSTab1.Tab = 0
   Exit Sub
End If
End Sub

Private Sub CmbValidType_LostFocus()
If CmbValidType.Text = "" Then Exit Sub

If LCase(CmbValidType.Text) <> LCase("Region") And LCase(CmbValidType.Text) <> LCase("State") And LCase(CmbValidType.Text) <> LCase("Location") Then
''-- And LCase(CmbValidType.Text) <> LCase("CMP") And LCase(CmbValidType.Text) <> LCase("Godown")
    MsgBox "Invalid Selection!!!"
    CmbValidTypeID.Text = ""
    CmbValidType.SetFocus
    Exit Sub
End If

If LCase(mValidType) <> LCase(Left$(CmbValidType, 1)) Then
    CmbValidTypeID.Text = ""
End If
mValidType = Left$(CmbValidType, 1)
Label19.Caption = CmbValidType.Text

''If LCase(mValidType) = LCase("C") Then
''    lblLocation.Visible = True
''    CmbLocationID = "": CmbLocationID.Visible = True
''    lblCMP.Visible = False
''    CmbCMPID = "": CmbCMPID.Visible = False
''    CmbLocationID.SetFocus
''ElseIf LCase(mValidType) = LCase("G") Then
''    lblLocation.Visible = True
''    CmbLocationID = "": CmbLocationID.Visible = True
''    lblCMP.Visible = True
''    CmbCMPID = "": CmbCMPID.Visible = True
''    CmbLocationID.SetFocus
''Else
''    lblLocation.Visible = False
''    CmbLocationID = "": CmbLocationID.Visible = False
''    lblCMP.Visible = False
''    CmbCMPID = "": CmbCMPID.Visible = False
''End If
End Sub

Private Sub CmbValidTypeID_GotFocus()
'If MSHFlexGrid7.Rows > 2 Then
'   MsgBox "By default all Locations which are not entered in 'Not Covered Locations' list on the right hand side of this screen are included in the Insurance Policy. You need not select any Location name for this purpose."
'   CmbValidTypeID.Clear
'   CmbValidTypeID.SetFocus
'   Exit Sub
'End If

If CmbValidType.Text = "" Then
    MsgBox "Select Type First!!!"
    CmbValidType.SetFocus
    Exit Sub
End If

tmp1 = CmbValidTypeID.Text

If LCase(CmbValidType.Text) = LCase("Region") Then ''-- Region
    ''tmp = " Select * from Mst_Region"
    tmp = " Select MR.RegionID,MR.Region from Mst_Godown MG"
    tmp = tmp & " Inner Join Mst_CMP CMP On MG.CMPID = CMP.CMPID"
    tmp = tmp & " Inner Join Mst_Location ML On CMP.LocationID = ML.LocationID"
    tmp = tmp & " Inner Join Mst_State MS On ML.StateID = MS.StateID"
    tmp = tmp & " Inner Join Mst_Region MR On MS.RegionID = MR.RegionID"
    tmp = tmp & " Where MG.CloseDate Is Null And MG.UnitTypeID In (" & mListOfSelectedUnitType & ")"
    tmp = tmp & " Group By MR.RegionID,MR.Region "
    Call TmpTable
    CmbValidTypeID.Clear
    If TmpRs.RecordCount = 0 Then
       MsgBox "No Record found !!!! "
       Exit Sub
    End If
    For I = 1 To TmpRs.RecordCount
        CmbValidTypeID.AddItem TmpRs!Region
        TmpRs.MoveNext
    Next
ElseIf LCase(CmbValidType.Text) = LCase("State") Then  ''-- State
    ''tmp = " Select * from Mst_State"
    tmp = " Select MS.StateID,MS.StateName from Mst_Godown MG"
    tmp = tmp & " Inner Join Mst_CMP CMP On MG.CMPID = CMP.CMPID"
    tmp = tmp & " Inner Join Mst_Location ML On CMP.LocationID = ML.LocationID"
    tmp = tmp & " Inner Join Mst_State MS On ML.StateID = MS.StateID"
    tmp = tmp & " Inner Join Mst_Region MR On MS.RegionID = MR.RegionID"
    tmp = tmp & " Where MG.CloseDate Is Null And MG.UnitTypeID In (" & mListOfSelectedUnitType & ")"
    tmp = tmp & " Group By MS.StateID,MS.StateName"
    Call TmpTable
    CmbValidTypeID.Clear
    If TmpRs.RecordCount = 0 Then
       MsgBox "No Record found !!!! "
       Exit Sub
    End If
    For I = 1 To TmpRs.RecordCount
        CmbValidTypeID.AddItem TmpRs!StateName
        TmpRs.MoveNext
    Next
ElseIf LCase(CmbValidType.Text) = LCase("Location") Then   ''-- Location
    ''tmp = " Select * from Mst_Location "
    tmp = " Select ML.LocationID,ML.LocationName from Mst_Godown MG"
    tmp = tmp & " Inner Join Mst_CMP CMP On MG.CMPID = CMP.CMPID"
    tmp = tmp & " Inner Join Mst_Location ML On CMP.LocationID = ML.LocationID"
    tmp = tmp & " Inner Join Mst_State MS On ML.StateID = MS.StateID"
    tmp = tmp & " Inner Join Mst_Region MR On MS.RegionID = MR.RegionID"
    tmp = tmp & " Where MG.CloseDate Is Null And MG.UnitTypeID In (" & mListOfSelectedUnitType & ")"
    tmp = tmp & " Group By ML.LocationID,ML.LocationName"
    Call TmpTable
    CmbValidTypeID.Clear
    If TmpRs.RecordCount = 0 Then
       MsgBox "No Record found !!!! "
       Exit Sub
    End If
    For I = 1 To TmpRs.RecordCount
        CmbValidTypeID.AddItem TmpRs!LocationName
        TmpRs.MoveNext
    Next
''ElseIf LCase(CmbValidType.Text) = LCase("CMP") Then    ''-- CMP
''    If CmbLocationID.Text = "" Then
''        MsgBox "Select Location"
''        CmbLocationID.SetFocus
''        Exit Sub
''    End If
''
''    ''tmp = " Select CMPID,CMPName,CMPName+'~'+convert(Varchar,CMPID)[CMP],Address from Mst_CMP Where LocationID = " & mLocationID & " And CloseDate Is Null "
''    tmp = " Select CMP.CMPID,CMP.CMPName,CMP.CMPName+'~'+convert(Varchar,CMP.CMPID)[CMP] from Mst_Godown MG"
''    tmp = tmp & " Inner Join Mst_CMP CMP On MG.CMPID = CMP.CMPID"
''    tmp = tmp & " Inner Join Mst_Location ML On CMP.LocationID = ML.LocationID"
''    tmp = tmp & " Inner Join Mst_State MS On ML.StateID = MS.StateID"
''    tmp = tmp & " Inner Join Mst_Region MR On MS.RegionID = MR.RegionID"
''    tmp = tmp & " Where MG.CloseDate Is Null And CMP.CloseDate Is Null And CMP.LocationID = " & mLocationID & " And MG.UnitTypeID In (" & mListOfSelectedUnitType & ")"
''    tmp = tmp & " Group By CMP.CMPID,CMP.CMPName"
''    Call TmpTable
''    CmbValidTypeID.Clear
''    If TmpRs.RecordCount = 0 Then
''       MsgBox "No Record found !!!! "
''       Exit Sub
''    End If
''    For i = 1 To TmpRs.RecordCount
''        CmbValidTypeID.AddItem TmpRs!CMP
''        TmpRs.MoveNext
''    Next
''ElseIf LCase(CmbValidType.Text) = LCase("Godown") Then    ''-- Godown
''    If CmbCMPID.Text = "" Then
''        MsgBox "Select CMP"
''        CmbCMPID.SetFocus
''        Exit Sub
''    End If
''
''    ''tmp = " Select CMPID,GodownID,GodownNo,GodownNo+'~'+convert(Varchar,GodownID)[Godown],Address from Mst_Godown Where CloseDate Is Null And CMPID = " & mCMPID & ""
''    tmp = " Select MG.GodownID,MG.GodownNo,MG.GodownNo+'~'+convert(Varchar,MG.GodownID)[Godown]"
''    tmp = tmp & " from Mst_Godown MG"
''    tmp = tmp & " Inner Join Mst_CMP CMP On MG.CMPID = CMP.CMPID"
''    tmp = tmp & " Inner Join Mst_Location ML On CMP.LocationID = ML.LocationID"
''    tmp = tmp & " Inner Join Mst_State MS On ML.StateID = MS.StateID"
''    tmp = tmp & " Inner Join Mst_Region MR On MS.RegionID = MR.RegionID"
''    tmp = tmp & " Where MG.CloseDate Is Null And MG.CMPID = " & mCMPID & " And MG.UnitTypeID In (" & mListOfSelectedUnitType & ")"
''    tmp = tmp & " Group By MG.GodownID,MG.GodownNo"
''    Call TmpTable
''    CmbValidTypeID.Clear
''    If TmpRs.RecordCount = 0 Then
''       MsgBox "No Record found !!!! "
''       Exit Sub
''    End If
''    For i = 1 To TmpRs.RecordCount
''        CmbValidTypeID.AddItem TmpRs!Godown
''        TmpRs.MoveNext
''    Next
End If
CmbValidTypeID.Text = tmp1
End Sub

Private Sub CmbValidTypeID_LostFocus()
If CmbValidTypeID.Text = "" Then Exit Sub

If LCase(CmbValidType.Text) = "region" Then
   If TmpRs.RecordCount > 0 Then TmpRs.MoveFirst
   TmpRs.Find "Region= '" & CmbValidTypeID.Text & "'"
   If TmpRs.EOF Then
      MsgBox "Invalid selection "
      CmbValidTypeID.SetFocus
      Exit Sub
   End If
   mValidTypeID = TmpRs!RegionID
ElseIf LCase(CmbValidType.Text) = "state" Then
   If TmpRs.RecordCount > 0 Then TmpRs.MoveFirst
   TmpRs.Find "StateName = '" & CmbValidTypeID.Text & "'"
   If TmpRs.EOF Then
      MsgBox "Invalid selection "
      CmbValidTypeID.SetFocus
      Exit Sub
   End If
   mValidTypeID = TmpRs!StateID
ElseIf LCase(CmbValidType.Text) = "location" Then
   If TmpRs.RecordCount > 0 Then TmpRs.MoveFirst
   TmpRs.Find "LocationName = '" & CmbValidTypeID.Text & "'"
   If TmpRs.EOF Then
      MsgBox "Invalid selection "
      CmbValidTypeID.SetFocus
      Exit Sub
   End If
   mValidTypeID = TmpRs!LocationID
ElseIf LCase(CmbValidType.Text) = "cmp" Then
   If TmpRs.RecordCount > 0 Then TmpRs.MoveFirst
   TmpRs.Find "CMP = '" & CmbValidTypeID.Text & "'"
   If TmpRs.EOF Then
      MsgBox "Invalid selection "
      CmbValidTypeID.SetFocus
      Exit Sub
   End If
   mValidTypeID = TmpRs!CMPID
   'TxtAddress.Visible = True: TxtAddress = IIf(IsNull(TmpRs!Address), "", TmpRs!Address)
ElseIf LCase(CmbValidType.Text) = "godown" Then
   If TmpRs.RecordCount > 0 Then TmpRs.MoveFirst
   TmpRs.Find "Godown = '" & CmbValidTypeID.Text & "'"
   If TmpRs.EOF Then
      MsgBox "Invalid selection "
      CmbValidTypeID.SetFocus
      Exit Sub
   End If
   mValidTypeID = TmpRs!GodownID
   'TxtAddress.Visible = True: TxtAddress = IIf(IsNull(TmpRs!Address), "", TmpRs!Address)
End If
End Sub

Private Sub CmdAddFrame3_Click()
Dim mR As Double

If CmbIncludeCommodityID = "" Then
   MsgBox "Blank Commodity Not Allowed."
   CmbIncludeCommodityID.SetFocus
   Exit Sub
End If

If MSHFlexGrid4.Rows > 2 Then
   MsgBox "By default all commodities which are not entered in 'Not Covered Commodities' list on the right hand side of this screen are included in the Insurance Policy. You need not select any commodity name for this purpose."
   CmbIncludeCommodityID.SetFocus
   Exit Sub
End If

If CheckDuplicateCommodity(MSHFlexGrid3) = True Then
   Exit Sub
End If
With MSHFlexGrid3
    If CmdAddFrame3.Caption <> "Add To List" Then
       mR = .RowSel
    Else
       mR = .Rows - 1
       .Rows = .Rows + 1
    End If
    If Val(.TextMatrix(mR, 3)) = 0 Then
        .TextMatrix(mR, 1) = CmbIncludeCommodityID.Text
        .TextMatrix(mR, 2) = mIncludeCommodityID
    End If
End With
Call ClearFrame3
End Sub

Private Sub CmdAddFrame4_Click()
Dim mR As Double

If CmbExcludeCommodityID = "" Then
   MsgBox "Blank Commodity Not Allowed."
   CmbExcludeCommodityID.SetFocus
   Exit Sub
End If

If MSHFlexGrid3.Rows > 2 Then
   MsgBox "By default all commodities which are not entered in 'Covered Commodities' List on the left hand side of this screen are excluded from the Insurance Policy. You need not select any commodity name for this purpose."
   CmbExcludeCommodityID.SetFocus
   Exit Sub
End If

If CheckDuplicateCommodity(MSHFlexGrid4) = True Then
   Exit Sub
End If
With MSHFlexGrid4
    If CmdAddFrame4.Caption <> "Add To List" Then
       mR = .RowSel
    Else
       mR = .Rows - 1
       .Rows = .Rows + 1
    End If
    If Val(.TextMatrix(mR, 3)) = 0 Then
        .TextMatrix(mR, 1) = CmbExcludeCommodityID
        .TextMatrix(mR, 2) = mExcludeCommodityID
    End If
End With
Call ClearFrame4

End Sub

Private Sub CmdAddFrame5_Click()
Dim mR As Double

'If CmbExchange.Text = "" Then
'    MsgBox "Select Exchange."
'    CmbExchange.SetFocus
'    Exit Sub
'End If
If CmbClientRowID = "" Then
   MsgBox "Blank Client Not Allowed."
   CmbClientRowID.SetFocus
   Exit Sub
End If
If CheckDuplicateClient = True Then Exit Sub
With MSHFlexGrid5
    If CmdAddFrame5.Caption <> "Add To List" Then
       mR = .RowSel
    Else
       mR = .Rows - 1
       .Rows = .Rows + 1
    End If
    If Val(.TextMatrix(mR, 4)) = 0 Then
        .TextMatrix(mR, 1) = TxtClientID.Text
        .TextMatrix(mR, 2) = mTmpName
        .TextMatrix(mR, 3) = mClientRowID
        '.TextMatrix(mR, 5) = "Non NCDEX" ''CmbExchange.Text
        '.TextMatrix(mR, 6) = mExchangeTypeID
    End If
End With
Call ClearFrame5
End Sub

Private Sub CmdAddtoListFrame2_Click()
Dim mR As Double
Dim mSplit As Variant
Dim mDeliminate As Boolean

If CmbValidType = "" Then
   MsgBox "Blank Type Not Allowed."
   CmbValidType.SetFocus
   Exit Sub
End If
If CmbValidTypeID = "" Then
   MsgBox "Blank " & CmbValidType & "  Not Allowed."
   CmbValidTypeID.SetFocus
   Exit Sub
End If
If InStr(CmbValidTypeID.Text, "~") > 0 Then
    mSplit = Split(CmbValidTypeID, "~")
    mDeliminate = True
Else
    mDeliminate = False
End If
''-- for Godown and CMP Delimited value
'If mDeliminate = True Then
'    If CheckDuplicateValidFor(True, mSplit(1)) = True Then
'       CmbValidTypeID.SetFocus
'       Exit Sub
'    End If
'Else
'    If CheckDuplicateValidFor(False, "") = True Then
'       CmbValidTypeID.SetFocus
'       Exit Sub
'    End If
'End If

If CheckDuplicateValidFor(MSHFlexGrid2) = True Then
   CmbValidTypeID.SetFocus
   Exit Sub
End If

With MSHFlexGrid2
    If CmdAddtoListFrame2.Caption <> "Add To List" Then
       mR = .RowSel
    Else
       mR = .Rows - 1
       .Rows = .Rows + 1
    End If
    .TextMatrix(mR, 1) = CmbValidType
    If mDeliminate = True Then
        .TextMatrix(mR, 2) = mSplit(0)
    Else
        .TextMatrix(mR, 2) = CmbValidTypeID
    End If
    .TextMatrix(mR, 3) = mValidTypeID
End With
Call ClearFrame2
End Sub

Private Sub ClearFrame2()
    TxtAttachment = ""
    CmdGetFile.Enabled = True
    CmdImport.Enabled = False
    CmdValidate.Enabled = False
    CmdExportGodownList.Enabled = True
End Sub
Private Sub ClearFrame7()
    CmbValidType = ""
    CmbValidTypeID = ""
    CmdAddtoListFrame7.Caption = "Add To List"
    CmdDeleteFrame7.Caption = "Delete From List"
    
''    CmbExcludeValidType = ""
''    CmbExcludeValidTypeID = ""
''    lblExcludeLocation.Visible = False
''    CmbExcludeLocationID = "": CmbExcludeLocationID.Visible = False
''    lblExcludeCMP.Visible = False
''    CmbExcludeCMPID = "": CmbExcludeCMPID.Visible = False
''    CmdAddtoListFrame2.Caption = "Add To List"
''    CmdDeleteFrame2.Caption = "Delete From List"
End Sub

Private Sub ClearFrame3()
    CmbIncludeCommodityID = ""
    CmdAddFrame3.Caption = "Add To List"
    CmdDeleteFrame3.Caption = "Delete From List"
End Sub
Private Sub ClearFrame4()
    CmbExcludeCommodityID = ""
    CmdAddFrame4.Caption = "Add To List"
    CmdDeleteFrame4.Caption = "Delete From List"
End Sub
Private Sub ClearFrame5()
    TxtClientID = ""
    CmbClientRowID = ""
    CmdAddFrame5.Caption = "Add To List"
    CmdDeleteFrame5.Caption = "Delete From List"
End Sub

Function CheckDuplicateClient() As Boolean
Dim Retval As Boolean
Retval = False
    With MSHFlexGrid5
        If .Rows > 2 Then
           For mR = 1 To .Rows - 2
            If .TextMatrix(mR, 3) = mClientRowID Then
               If CmdAddFrame5.Caption = "Add To List" Then
                  MsgBox "Duplicate Client not Allowed."
                  Retval = True
                  CheckDuplicateClient = Retval
                  Exit Function
               Else
                  If mR <> .RowSel Then
                     MsgBox "Duplicate Client not Allowed."
                     Retval = True
                     CheckDuplicateClient = Retval
                     Exit Function
                  End If
               End If
            End If
           Next
        End If
    End With
End Function

Function CheckDuplicateCommodity(Grid As MSHFlexGrid) As Boolean
Dim Retval As Boolean
Retval = False
    With Grid
        Select Case LCase(Grid.Name)
            Case LCase("MSHFlexGrid3")
            
                If .Rows > 2 Then
                   For mR = 1 To .Rows - 2
                    If .TextMatrix(mR, 2) = mIncludeCommodityID Then
                       If CmdAddFrame3.Caption = "Add To List" Then
                          MsgBox "Duplicate Commodity not Allowed."
                          Retval = True
                          CheckDuplicateCommodity = Retval
                          Exit Function
                       Else
                          If mR <> .RowSel Then
                             MsgBox "Duplicate Commodity not Allowed."
                             Retval = True
                             CheckDuplicateCommodity = Retval
                             Exit Function
                          End If
                       End If
                    End If
                   Next
                End If
            
            Case LCase("MSHFlexGrid4")
            
                If .Rows > 2 Then
                   For mR = 1 To .Rows - 2
                    If .TextMatrix(mR, 2) = mExcludeCommodityID Then
                       If CmdAddFrame3.Caption = "Add To List" Then
                          MsgBox "Duplicate Commodity not Allowed."
                          Retval = True
                          CheckDuplicateCommodity = Retval
                          Exit Function
                       Else
                          If mR <> .RowSel Then
                             MsgBox "Duplicate Commodity not Allowed."
                             Retval = True
                             CheckDuplicateCommodity = Retval
                             Exit Function
                          End If
                       End If
                    End If
                   Next
                End If
            
        End Select
        
    End With
End Function

Function FillUnitType()
Dim mR As Double, mRow As Double
With MSHFlexGrid6
    If Edit_Flg = "A" Then
        Call TableMst_UnitType(" Where Status = 'A' ")
        If RsMst_UnitType.RecordCount > 0 Then
            For mR = 1 To RsMst_UnitType.RecordCount
                .row = mR
                .Col = 0
                .CellFontName = "Wingdings"
                .CellFontSize = 12
                .CellAlignment = flexAlignCenterCenter
                .Text = strChecked
                .TextMatrix(mR, 1) = IIf(IsNull(RsMst_UnitType!UnitType), "", RsMst_UnitType!UnitType)
                .TextMatrix(mR, 2) = IIf(IsNull(RsMst_UnitType!UnitTypeID), "", RsMst_UnitType!UnitTypeID)
                .Rows = .Rows + 1
                RsMst_UnitType.MoveNext
            Next
        End If
    ElseIf Edit_Flg = "E" Then
        Call ListUnitType
        Call TableMst_UnitType(" Where Status = 'A' And UnitTypeID Not In (" & mUnitTypeList & ") ")
        mRow = MSHFlexGrid6.Rows - 2
        If RsMst_UnitType.RecordCount > 0 Then
            For mR = 1 To RsMst_UnitType.RecordCount
                .row = mR + mRow
                .Col = 0
                .CellFontName = "Wingdings"
                .CellFontSize = 12
                .CellAlignment = flexAlignCenterCenter
                .Text = strUnChecked
                .TextMatrix(mR + mRow, 1) = IIf(IsNull(RsMst_UnitType!UnitType), "", RsMst_UnitType!UnitType)
                .TextMatrix(mR + mRow, 2) = IIf(IsNull(RsMst_UnitType!UnitTypeID), "", RsMst_UnitType!UnitTypeID)
                .Rows = .Rows + 1
                RsMst_UnitType.MoveNext
            Next
        End If
    End If
End With
End Function

Function CheckDuplicateValidFor(Grid_ As MSHFlexGrid) As Boolean
Dim mR As Double, mRow As Double
Dim Retval As Boolean
Retval = False

If LCase(Grid_.Name) = LCase("MSHFlexGrid2") Then
    With Grid_
    If .Rows > 2 Then
       For mR = 1 To .Rows - 2
            If LCase(.TextMatrix(mR, 1)) = LCase(CmbValidType.Text) And Val(.TextMatrix(mR, 3)) = mValidTypeID Then
               If CmdAddtoListFrame2.Caption = "Add To List" Then
                  MsgBox "Duplicate " & CmbValidType.Text & "  not Allowed."
                  Retval = True
                  CheckDuplicateValidFor = Retval
                  Exit Function
               Else
                  If mR <> .RowSel Then
                     MsgBox "Duplicate " & CmbValidType.Text & " not Allowed."
                     Retval = True
                     CheckDuplicateValidFor = Retval
                     Exit Function
                  End If
               End If
            End If
       Next
    End If
    End With
ElseIf LCase(Grid_.Name) = LCase("MSHFlexGrid7") Then
    With Grid_
    If .Rows > 2 Then
       For mR = 1 To .Rows - 2
            If LCase(.TextMatrix(mR, 1)) = LCase(CmbValidType.Text) And Val(.TextMatrix(mR, 3)) = mValidTypeID Then
               If CmdAddtoListFrame7.Caption = "Add To List" Then
                  MsgBox "Duplicate " & CmbValidType.Text & "  not Allowed."
                  Retval = True
                  CheckDuplicateValidFor = Retval
                  Exit Function
               Else
                  If mR <> .RowSel Then
                     MsgBox "Duplicate " & CmbValidType.Text & " not Allowed."
                     Retval = True
                     CheckDuplicateValidFor = Retval
                     Exit Function
                  End If
               End If
            End If
       Next
    End If
    End With
End If

CheckDuplicateValidFor = Retval
End Function


Private Sub CmdAddtoListFrame7_Click()
Dim mR As Double
Dim mSplit As Variant
Dim mDeliminate As Boolean

If CmbValidType = "" Then
   MsgBox "Blank Location Type Not Allowed."
   CmbValidType.SetFocus
   Exit Sub
End If
If CmbValidTypeID = "" Then
   MsgBox "Blank " & CmbValidType & "  Not Allowed."
   CmbValidTypeID.SetFocus
   Exit Sub
End If

''If InStr(CmbExcludeValidTypeID.Text, "~") > 0 Then
''    mSplit = Split(CmbExcludeValidTypeID, "~")
''    mDeliminate = True
''Else
''    mDeliminate = False
''End If

''-- for Godown and CMP Delimited value
'If mDeliminate = True Then
'    If CheckDuplicateValidFor(True, mSplit(1)) = True Then
'       CmbValidTypeID.SetFocus
'       Exit Sub
'    End If
'Else
'    If CheckDuplicateValidFor(False, "") = True Then
'       CmbValidTypeID.SetFocus
'       Exit Sub
'    End If
'End If

If CheckDuplicateValidFor(MSHFlexGrid7) = True Then
   CmbValidTypeID.SetFocus
   Exit Sub
End If

With MSHFlexGrid7
    If CmdAddtoListFrame7.Caption <> "Add To List" Then
       mR = .RowSel
    Else
       mR = .Rows - 1
       .Rows = .Rows + 1
    End If
    ''-- Not allow to update data once it saved
    If Val(.TextMatrix(mR, 4)) = 0 Then
        .TextMatrix(mR, 1) = CmbValidType
        .TextMatrix(mR, 2) = CmbValidTypeID
        .TextMatrix(mR, 3) = mValidTypeID
    End If
End With

CmbValidTypeID = ""
CmdAddtoListFrame7.Caption = "Add To List"
CmdDeleteFrame7.Caption = "Delete From List"
CmbValidType.SetFocus
''Call ClearFrame7
End Sub

Function CheckAllGodownSelect() As Boolean
Dim mResult As Boolean
Dim mR As Double
mResult = True
If MSHFlexGrid2.Rows > 2 Then
    For mR = 1 To MSHFlexGrid2.Rows - 2
        If MSHFlexGrid2.TextMatrix(mR, 0) = strUnChecked Then
           mResult = False
           Exit For
        End If
    Next mR
Else
    mResult = False
End If
CheckAllGodownSelect = mResult
End Function

Function CountSelectedGodown() As Double
Dim mCount As Double
Dim mR As Double
mCount = 0
If MSHFlexGrid2.Rows > 2 Then
    For mR = 1 To MSHFlexGrid2.Rows - 2
        If MSHFlexGrid2.TextMatrix(mR, 0) = strChecked Then
           mCount = mCount + 1
        End If
    Next mR
End If
CountSelectedGodown = mCount
End Function
Function CountSelectedConnectedGodown() As Double
Dim mCount As Double
Dim mR As Double
mCount = 0
If MSHFlexGrid2.Rows > 2 Then
    For mR = 1 To MSHFlexGrid2.Rows - 2
        If MSHFlexGrid2.TextMatrix(mR, 9) = strChecked Then
           mCount = mCount + 1
        End If
    Next mR
End If
CountSelectedConnectedGodown = mCount
End Function


Function ListOfGodown()
Dim mRegionIDList As Variant, mStateIDList As Variant, mLocationIDList As Variant
Dim mR As Double, mC As Double
Dim mLastRow As Double

mRegionIDList = "0"
mStateIDList = "0"
mLocationIDList = "0"

Grid_Head2

With MSHFlexGrid7
    For mR = 1 To .Rows - 2
        If .TextMatrix(mR, 1) <> "" Then
            If LCase(.TextMatrix(mR, 1)) = LCase("Region") Then
                mRegionIDList = mRegionIDList & "," & Val(.TextMatrix(mR, 3))
            ElseIf LCase(.TextMatrix(mR, 1)) = LCase("State") Then
                mStateIDList = mStateIDList & "," & Val(.TextMatrix(mR, 3))
            ElseIf LCase(.TextMatrix(mR, 1)) = LCase("Location") Then
                mLocationIDList = mLocationIDList & "," & Val(.TextMatrix(mR, 3))
            End If
        End If
    Next mR
End With
ProgressBar1.Value = 0

With MSHFlexGrid2
    If mRegionIDList <> "" Or mStateIDList <> "" Or mLocationIDList <> "" Then
        ''-- Exitsing Godown List with Green Color
        tmp = " Select MG.GodownID,MG.GodownNo,MUT.UnitType,MG.UnitTypeID,MG.CMPID,CMP.CMPName,ML.LocationID,ML.LocationName,MS.StateID,MS.StateName,MR.RegionID,MR.Region,MG.Address,MG.GodownCapacity,Isnull(MIGM.Connected,'N')Connected,'G'TColor,MIGM.InsuranceGodownMappingID"
        tmp = tmp & " from Mst_Godown MG"
        tmp = tmp & " Inner Join Mst_UnitType MUT On MG.UnitTypeID = MUT.UnitTypeID"
        tmp = tmp & " Inner Join Mst_CMP CMP On MG.CMPID = CMP.CMPID"
        tmp = tmp & " Inner Join Mst_Location ML On CMP.LocationID = ML.LocationID"
        tmp = tmp & " Inner Join Mst_State MS On ML.StateID = MS.StateID"
        tmp = tmp & " Inner Join Mst_Region MR On MS.RegionID = MR.RegionID"
        tmp = tmp & " Inner Join Mst_InsuranceGodownMapping MIGM On MG.GodownID = MIGM.GodownID"
        tmp = tmp & " Where MIGM.InsurancePolicyID = " & Val(TxtInsurancePolicyID) & ""
        tmp = tmp & " Union"
        ''-- Remain Godown List with White Color
        tmp = tmp & " Select MG.GodownID,MG.GodownNo,MUT.UnitType,MG.UnitTypeID,MG.CMPID,CMP.CMPName,ML.LocationID,ML.LocationName,MS.StateID,MS.StateName,MR.RegionID,MR.Region,MG.Address,MG.GodownCapacity,'Y' AS Connected,'W'TColor,0 as InsuranceGodownMappingID "
        tmp = tmp & " from Mst_Godown MG"
        tmp = tmp & " Inner Join Mst_UnitType MUT On MG.UnitTypeID = MUT.UnitTypeID"
        tmp = tmp & " Inner Join Mst_CMP CMP On MG.CMPID = CMP.CMPID"
        tmp = tmp & " Inner Join Mst_Location ML On CMP.LocationID = ML.LocationID"
        tmp = tmp & " Inner Join Mst_State MS On ML.StateID = MS.StateID"
        tmp = tmp & " Inner Join Mst_Region MR On MS.RegionID = MR.RegionID"
        tmp = tmp & " Where MG.CloseDate Is Null And MG.UnitTypeID In (" & mListOfSelectedUnitType & ") And MR.RegionID In (" & mRegionIDList & ") And MG.GodownID NOT IN ( Select GodownID  from Mst_InsuranceGodownMapping Where InsurancePolicyID = " & Val(TxtInsurancePolicyID) & " ) "
        tmp = tmp & " Union"
        tmp = tmp & " Select MG.GodownID,MG.GodownNo,MUT.UnitType,MG.UnitTypeID,MG.CMPID,CMP.CMPName,ML.LocationID,ML.LocationName,MS.StateID,MS.StateName,MR.RegionID,MR.Region,MG.Address,MG.GodownCapacity,'Y' AS Connected,'W'TColor,0 as InsuranceGodownMappingID"
        tmp = tmp & " from Mst_Godown MG"
        tmp = tmp & " Inner Join Mst_UnitType MUT On MG.UnitTypeID = MUT.UnitTypeID"
        tmp = tmp & " Inner Join Mst_CMP CMP On MG.CMPID = CMP.CMPID"
        tmp = tmp & " Inner Join Mst_Location ML On CMP.LocationID = ML.LocationID"
        tmp = tmp & " Inner Join Mst_State MS On ML.StateID = MS.StateID"
        tmp = tmp & " Inner Join Mst_Region MR On MS.RegionID = MR.RegionID"
        tmp = tmp & " Where MG.CloseDate Is Null And MG.UnitTypeID In (" & mListOfSelectedUnitType & ") And MS.StateID In (" & mStateIDList & ") And MG.GodownID NOT IN ( Select GodownID  from Mst_InsuranceGodownMapping Where InsurancePolicyID = " & Val(TxtInsurancePolicyID) & " ) "
        tmp = tmp & " Union"
        tmp = tmp & " Select MG.GodownID,MG.GodownNo,MUT.UnitType,MG.UnitTypeID,MG.CMPID,CMP.CMPName,ML.LocationID,ML.LocationName,MS.StateID,MS.StateName,MR.RegionID,MR.Region,MG.Address,MG.GodownCapacity,'Y' AS Connected,'W'TColor,0 as InsuranceGodownMappingID"
        tmp = tmp & " from Mst_Godown MG"
        tmp = tmp & " Inner Join Mst_UnitType MUT On MG.UnitTypeID = MUT.UnitTypeID"
        tmp = tmp & " Inner Join Mst_CMP CMP On MG.CMPID = CMP.CMPID"
        tmp = tmp & " Inner Join Mst_Location ML On CMP.LocationID = ML.LocationID"
        tmp = tmp & " Inner Join Mst_State MS On ML.StateID = MS.StateID"
        tmp = tmp & " Inner Join Mst_Region MR On MS.RegionID = MR.RegionID"
        tmp = tmp & " Where MG.CloseDate Is Null And MG.UnitTypeID In (" & mListOfSelectedUnitType & ") And ML.LocationID In (" & mLocationIDList & ") And MG.GodownID NOT IN ( Select GodownID  from Mst_InsuranceGodownMapping Where InsurancePolicyID = " & Val(TxtInsurancePolicyID) & " ) "
        tmp = tmp & " Order By TColor,MIGM.InsuranceGodownMappingID,ML.LocationName,CMP.CMPName,MG.GodownNo "
        Call TmpTable
        If TmpRs.RecordCount > 0 Then
            ProgressBar1.Max = TmpRs.RecordCount
            mLastRow = .Rows - 1
            For mR = 1 To TmpRs.RecordCount
                
                    .row = mLastRow
                    .Col = 0
                    .CellFontName = "Wingdings"
                    .CellFontSize = 12
                    .CellAlignment = flexAlignCenterCenter
                    .Text = strChecked
                    .TextMatrix(mLastRow, 1) = IIf(IsNull(TmpRs!GodownNo), "", TmpRs!GodownNo)              ''"Godown No"
                    .TextMatrix(mLastRow, 2) = IIf(IsNull(TmpRs!UnitType), "", TmpRs!UnitType)              ''"Unit Type"
                    .TextMatrix(mLastRow, 3) = IIf(IsNull(TmpRs!CMPName), "", TmpRs!CMPName)                ''"CMP Name"
                    .TextMatrix(mLastRow, 4) = IIf(IsNull(TmpRs!LocationName), "", TmpRs!LocationName)      ''"Location"
                    .TextMatrix(mLastRow, 5) = IIf(IsNull(TmpRs!StateName), "", TmpRs!StateName)            ''"State"
                    .TextMatrix(mLastRow, 6) = IIf(IsNull(TmpRs!Region), "", TmpRs!Region)                  ''"Region"
                    .TextMatrix(mLastRow, 7) = IIf(IsNull(TmpRs!Address), "", TmpRs!Address)                ''"Address"
                    .TextMatrix(mLastRow, 8) = IIf(IsNull(TmpRs!GodownCapacity), "", TmpRs!GodownCapacity)  ''"Capacity"
                    
                    ''---Connected Default all new Godown is Connected ''"Connected/Not Connected"
                    .row = mLastRow
                    .Col = 9
                    .CellFontName = "Wingdings"
                    .CellFontSize = 12
                    .CellAlignment = flexAlignCenterCenter
                    If LCase(TmpRs!Connected) = "y" Then
                        .Text = strChecked
                    Else
                        .Text = strUnChecked
                    End If
                    ''---
                    
                    .TextMatrix(mLastRow, 10) = IIf(IsNull(TmpRs!GodownID), "", TmpRs!GodownID)             ''"GodownID"
                    .TextMatrix(mLastRow, 11) = IIf(IsNull(TmpRs!UnitTypeID), "", TmpRs!UnitTypeID)         ''"UnitID"
                    .TextMatrix(mLastRow, 12) = IIf(IsNull(TmpRs!InsuranceGodownMappingID), "", IIf(TmpRs!InsuranceGodownMappingID = "0", "", TmpRs!InsuranceGodownMappingID))     ''"RowID"
                    
                    If TmpRs!InsuranceGodownMappingID > 0 Then
                      If LCase(TmpRs!Connected) = "y" Then Call GRowColor(MSHFlexGrid2, mR, vbGreen)
                      If LCase(TmpRs!Connected) = "n" Then Call GRowColor(MSHFlexGrid2, mR, vbYellow)
                    Else
                        'Call GRowColor(MSHFlexGrid2, mR, vbWhite)
                    End If
                    
                    mLastRow = mLastRow + 1
                    .Rows = .Rows + 1
            
                TmpRs.MoveNext
                ProgressBar1.Value = mR
            Next mR
        End If
        
    End If
End With

''    If CheckGodownAvailable(MSHFlexGrid2, 9, IIf(IsNull(TmpRs!GodownID), "", TmpRs!GodownID)) = True Then
''    Else
''        For mR = 1 To .Rows - 2
''            .row = mR
''            .Col = 0
''            If Val(.TextMatrix(mR, 11)) > 0 Then
''                Call GRowColor(MSHFlexGrid2, mR, vbYellow)
''                .TextMatrix(mR, 0) = strUnChecked
''            Else
''                .TextMatrix(mR, 0) = strUnChecked
''            End If
''        Next
''    End If

If CheckAllGodownSelect = True Then
    ChkSelectAll.Tag = "NoClick"
    ChkSelectAll = 1
    ChkSelectAll.Tag = ""
Else
    ChkSelectAll.Tag = "NoClick"
    ChkSelectAll = 0
    ChkSelectAll.Tag = ""
End If

Label35.Caption = CountSelectedGodown & " / " & MSHFlexGrid2.Rows - 2
Label39.Caption = CountSelectedConnectedGodown & " / " & MSHFlexGrid2.Rows - 2

ProgressBar1.Value = ProgressBar1.Max
MsgBox "Done."

End Function

Public Function CheckGodownAvailable(Grid As MSHFlexGrid, Col As Double, GodownID As Integer) As Boolean
Dim mAllow As Boolean
Dim mR As Double
With Grid
    mAllow = True
    If .Rows > 2 Then
        For mR = 1 To .Rows - 2
            If Val(.TextMatrix(mR, Col)) = Val(GodownID) Then
            
                '.row = mR
                '.Col = 0
                'If .CellBackColor = vbYellow And Val(.TextMatrix(mR, 11)) > 0 Then
                '    Call GRowColor(MSHFlexGrid2, mR, vbGreen)
                '    .TextMatrix(mR, 0) = strChecked
                'Else
                '    If .CellBackColor <> vbRed Then
                '        .TextMatrix(mR, 0) = strChecked
                '    End If
                'End If
            
                mAllow = False
                Exit For
            End If
        Next mR
    End If
End With
CheckGodownAvailable = mAllow
End Function

Private Sub CmdDeleteFrame2_Click()
Dim mR As Double, mRow As Double, mCol As Double
With MSHFlexGrid2
    If CmbValidType = "" And CmbValidTypeID = "" Then Exit Sub
    mR = .RowSel
    If mR > 0 Then
       If .TextMatrix(mR, 1) <> "" Then
          For mRow = mR To .Rows - 2
              For mCol = 0 To .Cols - 1
                  .TextMatrix(mRow, mCol) = .TextMatrix(mRow + 1, mCol)
              Next
          Next
          .Rows = .Rows - 1
       End If
    End If
    .RowSel = 0
    Call ClearFrame2
End With
End Sub

Private Sub CmdDeleteFrame3_Click()
Dim mR As Double, mRow As Double, mCol As Double
With MSHFlexGrid3
    mR = .RowSel
    If CmbIncludeCommodityID = "" Then Exit Sub
    If Val(.TextMatrix(mR, 3)) > 0 Then Exit Sub
    If mR > 0 Then
       If .TextMatrix(mR, 1) <> "" Then
          For mRow = mR To .Rows - 2
              For mCol = 0 To .Cols - 1
                  .TextMatrix(mRow, mCol) = .TextMatrix(mRow + 1, mCol)
              Next
          Next
          .Rows = .Rows - 1
       End If
    End If
    .RowSel = 0
    Call ClearFrame3
End With
End Sub

Private Sub CmdDeleteFrame4_Click()
Dim mR As Double, mRow As Double, mCol As Double
With MSHFlexGrid4
    mR = .RowSel
    If CmbExcludeCommodityID = "" Then Exit Sub
    If Val(.TextMatrix(mR, 3)) > 0 Then Exit Sub
    If mR > 0 Then
       If .TextMatrix(mR, 1) <> "" Then
          For mRow = mR To .Rows - 2
              For mCol = 0 To .Cols - 1
                  .TextMatrix(mRow, mCol) = .TextMatrix(mRow + 1, mCol)
              Next
          Next
          .Rows = .Rows - 1
       End If
    End If
    .RowSel = 0
    Call ClearFrame4
End With
End Sub

Private Sub CmdDeleteFrame5_Click()
Dim mR As Double, mRow As Double, mCol As Double
With MSHFlexGrid5
    mR = .RowSel
    If CmbClientRowID = "" Then Exit Sub
    If Val(.TextMatrix(mR, 4)) > 0 Then Exit Sub
    If mR > 0 Then
       If .TextMatrix(mR, 1) <> "" Then
          For mRow = mR To .Rows - 2
              For mCol = 0 To .Cols - 1
                  .TextMatrix(mRow, mCol) = .TextMatrix(mRow + 1, mCol)
              Next
          Next
          .Rows = .Rows - 1
       End If
    End If
    .RowSel = 0
    Call ClearFrame5
End With
End Sub

Private Sub CmdDeleteFrame6_Click()

End Sub

Private Sub CmdDeleteFrame7_Click()
Dim mR As Double, mRow As Double, mCol As Double
With MSHFlexGrid7
    If CmbValidType = "" And CmbValidTypeID = "" Then Exit Sub
    mR = .RowSel
    ''-- If Locations are Already mapped and save with insurance then not allow to Edit/Remove
    If Val(.TextMatrix(mR, 4)) > 0 Then Exit Sub
        If mR > 0 Then
           If .TextMatrix(mR, 1) <> "" Then
              For mRow = mR To .Rows - 2
                  For mCol = 0 To .Cols - 1
                      .TextMatrix(mRow, mCol) = .TextMatrix(mRow + 1, mCol)
                  Next
              Next
              .Rows = .Rows - 1
           End If
        End If
    '.RowSel = 0
    .Refresh
    Call ClearFrame7
End With

If CheckAllRegionSelect = True Then
    ChkAddAllRegion.Tag = "NoClick"
    ChkAddAllRegion = 1
    ChkAddAllRegion.Tag = ""
Else
    ChkAddAllRegion.Tag = "NoClick"
    ChkAddAllRegion = 0
    ChkAddAllRegion.Tag = ""
End If

End Sub

Public Function CheckAllRegionSelect() As Boolean
Dim mR As Double, mNoofRegion As Double, mNoofRegionSelect As Double
Dim mEntry As Boolean

tmp = " Select MR.RegionID,MR.Region from Mst_Godown MG"
tmp = tmp & " Inner Join Mst_CMP CMP On MG.CMPID = CMP.CMPID"
tmp = tmp & " Inner Join Mst_Location ML On CMP.LocationID = ML.LocationID"
tmp = tmp & " Inner Join Mst_State MS On ML.StateID = MS.StateID"
tmp = tmp & " Inner Join Mst_Region MR On MS.RegionID = MR.RegionID"
tmp = tmp & " Where MG.CloseDate Is Null And MG.UnitTypeID In (" & mListOfSelectedUnitType & ")"
tmp = tmp & " Group By MR.RegionID,MR.Region"
Call TmpTable
If TmpRs.RecordCount > 0 Then mNoofRegion = Val(TmpRs.RecordCount) Else mNoofRegion = 0
mNoofRegionSelect = 0
With MSHFlexGrid7
    For mR = 1 To .Rows - 2
        If LCase(.TextMatrix(mR, 1)) = "region" Then
            mNoofRegionSelect = mNoofRegionSelect + 1
        End If
    Next mR
End With

If mNoofRegionSelect > 0 And mNoofRegion = mNoofRegionSelect Then
    mEntry = True
Else
    mEntry = False
End If
End Function


Private Sub CmdExportGodownList_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Mst_InsuranceGodownMappingList.xls")
Call Xls_GodownDetail_Rpt
End Sub

Public Sub Xls_GodownDetail_Rpt()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Double, mCol As Double
Dim C As Double, I As Double
Dim mAllList As Boolean

Select Case MsgBox("To export only selected godown list click 'Yes', otherwise click 'No' to export all godown list.", vbYesNo)
    Case vbYes
        mAllList = True
    Case vbNo
        mAllList = False
End Select

MsgBox ("Wait till next message")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Mst_InsuranceGodownMappingList.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
With MSHFlexGrid2
    If mAllList = False Then
        For I = 0 To MSHFlexGrid2.Rows - 2
        
            mRow = I + 1
            For C = 1 To .Cols - 2
                If InStr(.TextMatrix(I, C), "/") > 0 Then
                   If IsDate(.TextMatrix(I, C)) Then
                      oWS.Cells(mRow, C) = Format(.TextMatrix(I, C), "MM/dd/yyyy")
                   Else
                      oWS.Cells(mRow, C) = .TextMatrix(I, C)
                   End If
                Else
                    If C = 1 Then
                       oWS.Cells(mRow, C) = "'" & .TextMatrix(I, C)
                    ElseIf C = 9 Then
                       If .TextMatrix(I, C) = strChecked Then
                           oWS.Cells(mRow, C) = "Yes"
                       Else
                           oWS.Cells(mRow, C) = "No"
                       End If
                    Else
                       oWS.Cells(mRow, C) = .TextMatrix(I, C)
                    End If
                End If
            Next
            
            If I >= 1 Then
                .row = I
                .Col = 1
                If .CellBackColor = vbRed Then oWS.Range(oWS.Cells(mRow, 1), oWS.Cells(mRow, .Cols - 2)).Interior.Color = vbRed
                If .CellBackColor = vbGreen Then oWS.Range(oWS.Cells(mRow, 1), oWS.Cells(mRow, .Cols - 2)).Interior.Color = vbGreen
                If .CellBackColor = vbYellow Then oWS.Range(oWS.Cells(mRow, 1), oWS.Cells(mRow, .Cols - 2)).Interior.Color = vbYellow
            End If
            
        Next
    Else
        mRow = 0
        For I = 0 To MSHFlexGrid2.Rows - 2
           If I = 0 Then ''-- Header
           
                mRow = mRow + 1
                For C = 1 To .Cols - 2
                    If InStr(.TextMatrix(I, C), "/") > 0 Then
                       If IsDate(.TextMatrix(I, C)) Then
                          oWS.Cells(mRow, C) = Format(.TextMatrix(I, C), "MM/dd/yyyy")
                       Else
                          oWS.Cells(mRow, C) = .TextMatrix(I, C)
                       End If
                    Else
                        If C = 1 Then
                           oWS.Cells(mRow, C) = "'" & .TextMatrix(I, C)
                        Else
                           oWS.Cells(mRow, C) = .TextMatrix(I, C)
                        End If
                    End If
                Next
           
           Else ''-- Data Cells
           
                If .TextMatrix(I, 0) = strChecked Then
                     mRow = mRow + 1
                     For C = 1 To .Cols - 2
                         If InStr(.TextMatrix(I, C), "/") > 0 Then
                            If IsDate(.TextMatrix(I, C)) Then
                               oWS.Cells(mRow, C) = Format(.TextMatrix(I, C), "MM/dd/yyyy")
                            Else
                               oWS.Cells(mRow, C) = .TextMatrix(I, C)
                            End If
                         Else
                             If C = 1 Then
                                oWS.Cells(mRow, C) = "'" & .TextMatrix(I, C)
                             ElseIf C = 9 Then
                                If .TextMatrix(I, C) = strChecked Then
                                    oWS.Cells(mRow, C) = "Yes"
                                Else
                                    oWS.Cells(mRow, C) = "No"
                                End If
                             Else
                                oWS.Cells(mRow, C) = .TextMatrix(I, C)
                             End If
                         End If
                     Next
                     
                     If I >= 1 Then
                         .row = I
                         .Col = 1
                         If .CellBackColor = vbRed Then oWS.Range(oWS.Cells(mRow, 1), oWS.Cells(mRow, .Cols - 2)).Interior.Color = vbRed
                         If .CellBackColor = vbGreen Then oWS.Range(oWS.Cells(mRow, 1), oWS.Cells(mRow, .Cols - 2)).Interior.Color = vbGreen
                         If .CellBackColor = vbYellow Then oWS.Range(oWS.Cells(mRow, 1), oWS.Cells(mRow, .Cols - 2)).Interior.Color = vbYellow
                     End If
                End If
           
           End If
        Next
    End If

End With
oWB.Save
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub

Private Sub CmdGetFile_Click()
With cdialog
     .InitDir = App.Path & "\excel\"
     .DialogTitle = "Select File"
     .FilterIndex = 1
     .Filter = "Excel files(*.xls;*.xlsx)|*.xls;*.xlsx"
     .ShowOpen
End With
If cdialog.CancelError Then
   TxtAttachment.Text = ""
Else
   strFilename = cdialog.FileName
   If strFilename = "" Then Exit Sub
   TxtAttachment.Refresh
   TxtAttachment.Text = strFilename 'GetFileFromPath(strFilename)
   CmdImport.Enabled = True
   MsgBox "Select Import"
End If
End Sub

Private Sub CmdGo_Click()
Dim wc As Variant

Call Grid_Head
wc = ""
tmp = ""

If Trim(TxtSPolicyNo.Text) <> "" Then
   wc = " Where MID.PolicyNo Like '%" & Trim(TxtSPolicyNo) & "%'"
End If
If Trim(TxtSCoverNoteNo.Text) <> "" Then
   If wc = "" Then
      wc = " Where MID.CoverNoteNo Like '%" & Trim(TxtSCoverNoteNo) & "%'"
   Else
      wc = wc & " AND MID.CoverNoteNo Like '%" & Trim(TxtSCoverNoteNo) & "%'"
   End If
End If
If Trim(CmbSInsuranceCo.Text) <> "" Then
   If wc = "" Then
      wc = " Where MID.InsuranceCoID = " & mSInsuranceCoID & ""
   Else
      wc = wc & " AND MID.InsuranceCoID = " & mSInsuranceCoID & ""
   End If
End If
If Trim(CmbSInsuranceRiskTypeID.Text) <> "" Then
   If wc = "" Then
      wc = " Where MID.InsuranceRiskCoveredID = " & mSInsuranceRiskTypeID & ""
   Else
      wc = wc & " AND MID.InsuranceRiskCoveredID = " & mSInsuranceRiskTypeID & ""
   End If
End If
If Trim(CmbSInsuranceByID.Text) <> "" Then
   If wc = "" Then
      wc = " Where MID.InsuranceByID = " & mSInsuranceByID & ""
   Else
      wc = wc & " AND MID.InsuranceByID = " & mSInsuranceByID & ""
   End If
End If
If Trim(CmbSStatus.Text) <> "" Then
   If wc = "" Then
      wc = " Where MID.Flag = '" & Left(CmbSStatus.Text, 1) & "'"
   Else
      wc = wc & " AND MID.Flag = '" & Left(CmbSStatus.Text, 1) & "'"
   End If
End If

''-- Ullhas Patil 04 April 2018 change for Mapped Details
''tmp = " Select A.InsurancePolicyID,A.PolicyNo,A.CoverNoteNo,A.InsuranceCoID,B.InsuranceCoName,A.InsuranceRiskCoveredID,C.InsuranceRiskCovered,A.InsuranceByID,D.InsuranceByName,A.StructureWiseLimit,A.Flag,A.Amount,A.ValidFrom,A.ValidTo,A.Remark,A.G_UID"
''tmp = tmp & " from Mst_InsuranceDetails As A"
''tmp = tmp & " Left Join Mst_InsuranceCoMaster B On A.InsuranceCoID = B.InsuranceCoID"
''tmp = tmp & " Left Join Mst_InsuranceRiskCovered C On A.InsuranceRiskCoveredID = C.InsuranceRiskCoveredID"
''tmp = tmp & " Left Join Mst_InsuranceBy D On A.InsuranceByID = D.InsuranceByID"
''tmp = tmp & wc & " Order By A.InsurancePolicyID "

tmp = " Select MID.InsurancePolicyID,MID.PolicyNo,MID.CoverNoteNo,MID.InsuranceCoID,MICM.InsuranceCoName,MID.InsuranceRiskCoveredID, "
tmp = tmp & " MIRC.InsuranceRiskCovered,MID.InsuranceByID,MIB.InsuranceByName,MID.ValidFrom,MID.ValidTo,MID.Flag,MID.Amount,MID.StructureWiseLimit,MID.Remark, "
tmp = tmp & " Case When A.NoofGodownID > 0 then 'Yes' Else 'No' End[ConnectedGodown], "
tmp = tmp & " Case When B.NoofIncludeCommodity > 0 then 'Yes' Else 'No' End[CommodityMapped], "
tmp = tmp & " Case When C.NoofClient > 0 then 'Yes' Else 'No' End[ClientMapped] "
tmp = tmp & " from Mst_InsuranceDetails MID "
tmp = tmp & " Left Join Mst_InsuranceBy MIB On MID.InsuranceByID = MIB.InsuranceByID "
tmp = tmp & " Left Join Mst_InsuranceRiskCovered MIRC On MID.InsuranceRiskCoveredID = MIRC.InsuranceRiskCoveredID "
tmp = tmp & " Left Join Mst_InsuranceCoMaster MICM On MID.InsuranceCoID = MICM.InsuranceCoID "
tmp = tmp & " Left Join (SELECT MIGM.InsurancePolicyID,Count(MIGM.GodownID)[NoofGodownID] From  Mst_InsuranceGodownMapping MIGM Where Connected = 'Y' Group BY MIGM.InsurancePolicyID) A On MID.InsurancePolicyID = A.InsurancePolicyID "
tmp = tmp & " Left Join (SELECT MICDI.InsurancePolicyID,Count(MICDI.CommodityID)[NoofIncludeCommodity] From  Mst_InsuranceCommodityDetail MICDI Group BY MICDI.InsurancePolicyID) B On MID.InsurancePolicyID = B.InsurancePolicyID "
tmp = tmp & " Left Join (SELECT MICDD.InsurancePolicyID,Count(MICDD.ClientIDRow)[NoofClient] From  Mst_InsuranceClientDetail MICDD Group BY MICDD.InsurancePolicyID) C On MID.InsurancePolicyID = C.InsurancePolicyID "
tmp = tmp & " " & wc & " Order By MID.InsurancePolicyID "
Call Tmp3Table

With MSHFlexGrid1
    For I = 1 To TmpRs3.RecordCount
        .TextMatrix(I, 0) = IIf(IsNull(TmpRs3!InsurancePolicyID), "", TmpRs3!InsurancePolicyID)
        .TextMatrix(I, 1) = IIf(IsNull(TmpRs3!PolicyNo), "", TmpRs3!PolicyNo)
        .TextMatrix(I, 2) = IIf(IsNull(TmpRs3!CoverNoteNo), "", TmpRs3!CoverNoteNo)
        .TextMatrix(I, 3) = IIf(IsNull(TmpRs3!InsuranceCoName), "", TmpRs3!InsuranceCoName)
        .TextMatrix(I, 4) = IIf(IsNull(TmpRs3!InsuranceRiskCovered), "", TmpRs3!InsuranceRiskCovered)
        .TextMatrix(I, 5) = IIf(IsNull(TmpRs3!InsuranceByName), "", TmpRs3!InsuranceByName)
        .TextMatrix(I, 6) = IIf(IIf(IsNull(TmpRs3!Flag), "A", TmpRs3!Flag) = "A", "Active", "Deactive")
        .TextMatrix(I, 7) = IIf(IsNull(TmpRs3!Amount), "", Format(TmpRs3!Amount, "#0.00"))
        .TextMatrix(I, 8) = IIf(IsNull(TmpRs3!StructureWiseLimit), "", Format(TmpRs3!StructureWiseLimit, "#0.00"))
        .TextMatrix(I, 9) = Format(TmpRs3!ValidFrom, "dd-mmm-yyyy")
        .TextMatrix(I, 10) = Format(TmpRs3!ValidTo, "dd-mmm-yyyy")
        
        .TextMatrix(I, 11) = IIf(IsNull(TmpRs3!ConnectedGodown), "", TmpRs3!ConnectedGodown)
        .TextMatrix(I, 12) = IIf(IsNull(TmpRs3!CommodityMapped), "", TmpRs3!CommodityMapped)
        .TextMatrix(I, 13) = IIf(IsNull(TmpRs3!ClientMapped), "", TmpRs3!ClientMapped)
        .TextMatrix(I, 14) = IIf(IsNull(TmpRs3!Remark), "", TmpRs3!Remark)
        
        .Rows = .Rows + 1
        TmpRs3.MoveNext
    Next
End With

Label12.Caption = TmpRs3.RecordCount

End Sub

Private Sub CmdImport_Click()
If TxtAttachment <> "" Then
   If FillData("Sheet1", MSHFlexGrid8) = True Then
      CmdGetFile.Enabled = True
      CmdImport.Enabled = False
      CmdValidate.Enabled = True
      MsgBox "Select Validate"
   Else
      CmdGetFile.Enabled = True
      CmdImport.Enabled = False
      CmdValidate.Enabled = False
   End If
End If
Label35.Caption = CountSelectedGodown & " / " & MSHFlexGrid2.Rows - 2
Label39.Caption = CountSelectedConnectedGodown & " / " & MSHFlexGrid2.Rows - 2
End Sub

Private Function FillData(SheetNumber As Variant, Grid As MSHFlexGrid) As Boolean

On Error GoTo ExcelError
Dim intNumberRows As Double
Dim WorkSheetName As Variant
    With Grid
        Set oXL = New Excel.Application
        Set oWB = oXL.Workbooks.Open(strFilename)
        Set oWS = oWB.Worksheets(SheetNumber)
        oWS.ClearArrows
        
        Call Grid_Head8
        ''-- Sample Data Header Checking
        For I = 1 To .Cols - 2
            If Trim(LCase(oWS.Cells(1, I))) <> Trim(LCase(.TextMatrix(0, I))) Then
                MsgBox "Excel Caption mismatch!" & vbNewLine & "Invalid Caption Name for column - " & .TextMatrix(0, I) & vbNewLine & "Column No - " & I, vbOKOnly + vbExclamation, App.Title
                oXL.Quit
                Me.MousePointer = vbDefault
                FillData = False
                Exit Function
            End If
        Next I

        mRow = 1
        RowCount = oWB.Worksheets(1).UsedRange.Rows.Count
        If RowCount < 2 Then
            MsgBox "No Record found to Import."
            oXL.Quit
            FillData = False
            Exit Function
        End If
        'Call Grid_Head2
                
        Do
           If oWS.Cells(mRow + 1, 1) = "" Then
                If mRow = 1 Then
                    oXL.Quit
                    FillData = False
                    Exit Function
                End If
              Exit Do
           Else
           
'            .row = mRow
'            .Col = 0
'            .CellFontName = "Wingdings"
'            .CellFontSize = 12
'            .CellAlignment = flexAlignCenterCenter
'            .Text = strChecked
            For C = 1 To .Cols - 2
                .TextMatrix(mRow, C) = Trim(oWS.Cells(mRow + 1, C).Value)
            Next
                
           End If

           .Rows = .Rows + 1
           .Refresh
           mRow = mRow + 1
        Loop Until False
        
        oXL.Quit
        ''------------
        Dim mGodown As Variant
        Dim mLastRow As Double
        
        For mR = 1 To MSHFlexGrid2.Rows - 2
            MSHFlexGrid2.row = mR
            MSHFlexGrid2.Col = 0
            If Val(MSHFlexGrid2.TextMatrix(mR, 12)) > 0 Then
                Call GRowColor(MSHFlexGrid2, CDbl(mR), vbYellow)
                'MSHFlexGrid2.TextMatrix(mR, 0) = strUnChecked
                MSHFlexGrid2.TextMatrix(mR, 9) = strUnChecked
            Else
                MSHFlexGrid2.TextMatrix(mR, 0) = strUnChecked
                MSHFlexGrid2.TextMatrix(mR, 9) = strUnChecked
            End If
        Next mR
        
ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid8.Rows - 2

        For mR = 1 To MSHFlexGrid8.Rows - 2
        
            mGodown = MSHFlexGrid8.TextMatrix(mR, 10)

                If ChkGodown(mGodown, MSHFlexGrid8.TextMatrix(mR, 9)) = True Then
                   ''-- Mark checked in chkGodown Function
                   ''-- And Marked green if the Detail ID is available
                   
                Else
                   ''-- Add new godown if not available in List
                    mLastRow = (MSHFlexGrid2.Rows - 2)
                    MSHFlexGrid2.row = mLastRow + 1
                    MSHFlexGrid2.Col = 0
                    MSHFlexGrid2.CellFontName = "Wingdings"
                    MSHFlexGrid2.CellFontSize = 12
                    MSHFlexGrid2.CellAlignment = flexAlignCenterCenter
                    MSHFlexGrid2.Text = strChecked
                    
                    MSHFlexGrid2.TextMatrix(mLastRow + 1, 1) = .TextMatrix(mR, 1)                            ''"Godown No"
                    MSHFlexGrid2.TextMatrix(mLastRow + 1, 2) = .TextMatrix(mR, 2)                            ''"Unit Type"
                    MSHFlexGrid2.TextMatrix(mLastRow + 1, 3) = .TextMatrix(mR, 3)                            ''"CMP Name"
                    MSHFlexGrid2.TextMatrix(mLastRow + 1, 4) = .TextMatrix(mR, 4)                            ''"Location"
                    MSHFlexGrid2.TextMatrix(mLastRow + 1, 5) = .TextMatrix(mR, 5)                            ''"State"
                    MSHFlexGrid2.TextMatrix(mLastRow + 1, 6) = .TextMatrix(mR, 6)                            ''"Region"
                    MSHFlexGrid2.TextMatrix(mLastRow + 1, 7) = .TextMatrix(mR, 7)                            ''"Address"
                    MSHFlexGrid2.TextMatrix(mLastRow + 1, 8) = .TextMatrix(mR, 8)                            ''"Capacity"
                    
                    'MSHFlexGrid2.TextMatrix(mLastRow + 1, 9) = .TextMatrix(mR, 9)                            ''"Connected"
                    If LCase(.TextMatrix(mR, 9)) = "yes" Then
                        MSHFlexGrid2.row = mLastRow + 1
                        MSHFlexGrid2.Col = 0
                        MSHFlexGrid2.CellFontName = "Wingdings"
                        MSHFlexGrid2.CellFontSize = 12
                        MSHFlexGrid2.CellAlignment = flexAlignCenterCenter
                        MSHFlexGrid2.Text = strChecked
                    Else
                        MSHFlexGrid2.row = mLastRow + 1
                        MSHFlexGrid2.Col = 0
                        MSHFlexGrid2.CellFontName = "Wingdings"
                        MSHFlexGrid2.CellFontSize = 12
                        MSHFlexGrid2.CellAlignment = flexAlignCenterCenter
                        MSHFlexGrid2.Text = strUnChecked
                    End If
                    
                    MSHFlexGrid2.TextMatrix(mLastRow + 1, 10) = .TextMatrix(mR, 10)                            ''"GodownID"
                    MSHFlexGrid2.TextMatrix(mLastRow + 1, 11) = .TextMatrix(mR, 11)                           ''"UnitID"
                    MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
                   ''--------------------------------------------
                End If
                
                ProgressBar1.Value = mR
        Next mR
        ''-------------
        If CheckAllGodownSelect = True Then
            ChkSelectAll.Tag = "NoClick"
            ChkSelectAll = 1
            ChkSelectAll.Tag = ""
        Else
            ChkSelectAll.Tag = "NoClick"
            ChkSelectAll = 0
            ChkSelectAll.Tag = ""
        End If
        Label35.Caption = CountSelectedGodown & " / " & MSHFlexGrid2.Rows - 2
        Label39.Caption = CountSelectedConnectedGodown & " / " & MSHFlexGrid2.Rows - 2
        
        ProgressBar1.Value = ProgressBar1.Max
        
        MsgBox "Data Imported Successfully ", vbInformation, App.Title
        
        FillData = True
        
    End With
ExcelError:
    If Err.Number > 0 Then
        MsgBox "Error Number : " & Err.Number & vbNewLine & "Error Description : " & Err.Description, vbOKOnly + vbCritical, App.Title
        oXL.Quit
        FillData = False
        Me.MousePointer = Default
    End If
End Function

Private Sub CmdInsuranceCo_Click()
FrmMstInsuranceCo.Show
End Sub

Public Function ChkGodown(Id_ As Variant, Value_ As Variant) As Boolean
Dim mR As Double
Dim mResult As Boolean
mResult = False
With MSHFlexGrid2
   For mR = 1 To .Rows - 2
    
    If .TextMatrix(mR, 10) = Id_ Then
        mResult = True
                         
''        If Val(.TextMatrix(mR, 12)) > 0 Then
''            Call GRowColor(MSHFlexGrid2, mR, vbGreen)
''        Else
''            Call GRowColor(MSHFlexGrid2, mR, vbWhite)
''        End If
''        .TextMatrix(mR, 0) = strChecked
         If Val(.TextMatrix(mR, 12)) > 0 Then
            If LCase(Value_) <> IIf(.TextMatrix(mR, 9) = strChecked, "yes", "no") Then
                .TextMatrix(mR, 9) = IIf(LCase(Value_) = "yes", strChecked, strUnChecked)
                If .TextMatrix(mR, 9) = strChecked Then
                    Call GRowColor(MSHFlexGrid2, mR, vbGreen)
                Else
                    Call GRowColor(MSHFlexGrid2, mR, vbYellow)
                End If
            End If
         Else
            If LCase(Value_) <> IIf(.TextMatrix(mR, 9) = strChecked, "yes", "no") Then
                .TextMatrix(mR, 9) = IIf(LCase(Value_) = "yes", strChecked, strUnChecked)
                If .TextMatrix(mR, 9) = strChecked Then
                    .TextMatrix(mR, 0) = strChecked
                    Call GRowColor(MSHFlexGrid2, mR, vbWhite)
                Else
                    .TextMatrix(mR, 0) = strUnChecked
                    Call GRowColor(MSHFlexGrid2, mR, vbWhite)
                End If
            End If
         End If
        Exit For
    End If
   
   Next mR
End With
ChkGodown = mResult
End Function

Private Sub CmdGodownList_Click()
    Call ListOfGodown
    SSTab1.Tab = 1
End Sub

Private Sub CmdValidate_Click()
Dim mR As Double, mC As Double
Dim mListofUnitTypeID As Variant

mListofUnitTypeID = Split(mListOfSelectedUnitType, ",")
ProgressBar1.Value = 0
With MSHFlexGrid2
    If .Rows > 2 Then
    tmp = " Select MG.GodownID,MG.GodownNo,MUT.UnitType,MG.UnitTypeID,MG.CMPID,CMP.CMPName,ML.LocationID,ML.LocationName,MS.StateID,MS.StateName,MR.RegionID,MR.Region,MG.Address from Mst_Godown MG Inner Join Mst_UnitType MUT On MG.UnitTypeID = MUT.UnitTypeID Inner Join Mst_CMP CMP On MG.CMPID = CMP.CMPID Inner Join Mst_Location ML On CMP.LocationID = ML.LocationID Inner Join Mst_State MS On ML.StateID = MS.StateID Inner Join Mst_Region MR On MS.RegionID = MR.RegionID Where MG.CloseDate Is Null And MG.UnitTypeID In (" & mListOfSelectedUnitType & ")  "
    Call Tmp1Table
    If TmpRs1.RecordCount > 0 Then
    
        ProgressBar1.Max = .Rows
        
        For mR = 1 To .Rows - 2
            If .TextMatrix(mR, 0) = strChecked Then
                If TmpRs1.RecordCount > 0 Then TmpRs1.MoveFirst
                TmpRs1.Find (" GodownID = " & Val(.TextMatrix(mR, 10)) & " ")
                If TmpRs1.EOF = False Then
                    ''-- Godown No
                    If LCase(Trim(.TextMatrix(mR, 1))) <> LCase(IIf(IsNull(TmpRs1!GodownNo), "", TmpRs1!GodownNo)) Then
                        Call GRowColor(MSHFlexGrid2, mR, vbRed)
                        .TextMatrix(mR, 0) = strUnChecked
                        GoTo NextRow
                    End If
                    ''-- Unit Type
                    If LCase(Trim(.TextMatrix(mR, 2))) <> LCase(IIf(IsNull(TmpRs1!UnitType), "", TmpRs1!UnitType)) Then
                        Call GRowColor(MSHFlexGrid2, mR, vbRed)
                        .TextMatrix(mR, 0) = strUnChecked
                        GoTo NextRow
                    End If
                    ''-- CMP Name
                    If LCase(Trim(.TextMatrix(mR, 3))) <> LCase(IIf(IsNull(TmpRs1!CMPName), "", TmpRs1!CMPName)) Then
                        Call GRowColor(MSHFlexGrid2, mR, vbRed)
                        .TextMatrix(mR, 0) = strUnChecked
                        GoTo NextRow
                    End If
                    ''-- UnitTypeID
                    If Val(.TextMatrix(mR, 11)) <> Val(IIf(IsNull(TmpRs1!UnitTypeID), "", TmpRs1!UnitTypeID)) Then
                        Call GRowColor(MSHFlexGrid2, mR, vbRed)
                        .TextMatrix(mR, 0) = strUnChecked
                        GoTo NextRow
                    End If
                    
                Else
                    Call GRowColor(MSHFlexGrid2, mR, vbRed)
                    .TextMatrix(mR, 0) = strUnChecked
                End If
            End If
NextRow:
        ProgressBar1.Value = mR
        Next mR
    Else
        ''-- If Godown List is empty and grid with godown list then unmark all list
        For mR = 1 To .Rows - 2
            Call GRowColor(MSHFlexGrid2, mR, vbRed)
            .TextMatrix(mR, 0) = strUnChecked
        Next mR
    End If

    End If
End With

If CheckAllGodownSelect = True Then
    ChkSelectAll.Tag = "NoClick"
    ChkSelectAll = 1
    ChkSelectAll.Tag = ""
Else
    ChkSelectAll.Tag = "NoClick"
    ChkSelectAll = 0
    ChkSelectAll.Tag = ""
End If
Label35.Caption = CountSelectedGodown & " / " & MSHFlexGrid2.Rows - 2
Label39.Caption = CountSelectedConnectedGodown & " / " & MSHFlexGrid2.Rows - 2

ProgressBar1.Value = ProgressBar1.Max

MsgBox "Data Validate Successfully." & vbNewLine & "Red Color marked Godown ID's are invalid as per imported data in Excel.", vbInformation, App.Title

CmdGetFile.Enabled = True
CmdImport.Enabled = False
CmdValidate.Enabled = False

End Sub

Private Sub DeleteCmd_Click()
Dim ans As Variant
On Error GoTo msgError

tmp = "Select * from Txn_WHR_Detail Where PolicyNo = '" & TxtPolicyNo & "' And CoverNoteNo = '" & TxtCoverNoteNo & "' "
Call TmpTable
If TmpRs.RecordCount > 0 Then
   MsgBox "Policy No and CoverNote No used in WHR Details. Can't delete!!!"
   Exit Sub
End If

tmp = "Select * from Txn_CDC_Insurance_Detail Where PolicyNo = '" & TxtPolicyNo & "' And CoverNoteNo = '" & TxtCoverNoteNo & "' "
Call TmpTable
If TmpRs.RecordCount > 0 Then
   MsgBox "Policy No and CoverNote No used in CDC Transactions. Can't delete!!!"
   Exit Sub
End If


ans = MsgBox("Do you really want to DELETE this record??? ", vbYesNo)
If ans = vbYes Then

    RsMst_PolicyDetail!DeleteDate = Now
    RsMst_PolicyDetail!DeleteBy = Gen_UserLoginID
    RsMst_PolicyDetail.Update

    RsMst_PolicyDetail.Filter = ""
    RsMst_PolicyDetail.MoveFirst
    RsMst_PolicyDetail.Find "InsurancePolicyID = " & Val(TxtInsurancePolicyID) & ""

    ''-- Insurance Client Details
    tmp = " Delete From Mst_InsuranceClientDetail Where InsurancePolicyID = " & Val(TxtInsurancePolicyID) & ""
    Call TmpTable
    ''-- Insurance Commodity
    tmp = " Delete From Mst_InsuranceCommodityDetail Where InsurancePolicyID = " & Val(TxtInsurancePolicyID) & ""
    Call TmpTable
    ''-- Insurance Godown Mapping
    tmp = " Delete From Mst_InsuranceGodownMapping Where InsurancePolicyID = " & Val(TxtInsurancePolicyID) & " "
    Call TmpTable
    ''-- Insurance Valid For (Location Details)
    tmp = " Delete From Mst_InsuranceValidFor Where InsurancePolicyID = " & Val(TxtInsurancePolicyID) & " "
    Call TmpTable
    ''-- Insurance Unit Type
    tmp = " Delete From Mst_InsuranceForUnitTypeDetail Where InsurancePolicyID = " & Val(TxtInsurancePolicyID) & " "
    Call TmpTable

   RsMst_PolicyDetail.Delete
   RsMst_PolicyDetail.Update
   'MsgBox "Record Successfully Removed"
   If RsMst_PolicyDetail.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_PolicyDetail.MoveNext
   If RsMst_PolicyDetail.EOF() Then
      RsMst_PolicyDetail.MoveLast
   End If
   Call Indata
End If

Exit Sub
msgError:
MsgBox "Child record Present "
RsMst_PolicyDetail.CancelUpdate


End Sub

Private Sub EditCmd_Click()
Edit_Flg = "E"

TxtPolicyNo.Locked = True
TxtCoverNoteNo.Locked = True

''''----------------------------------
mGodownChangeFlag = True
mCommodityChangeFlag = True
mClientChangeFlag = True

'tmp = "Select * from Txn_WHR_Detail Where PolicyNo = '" & TxtPolicyNo & "' And CoverNoteNo = '" & TxtCoverNoteNo & "' "
'Call TmpTable
'If TmpRs.RecordCount > 0 Then
'    mGodownChangeFlag = True
'    mCommodityChangeFlag = False
'    mClientChangeFlag = False
'End If

If Val(mInsuranceCoID) > 0 Then
    TxtInsuranceCo.Visible = True: CmbInsuranceCo.Visible = False:  CmdInsuranceCo.Visible = False
    CmbInsuranceCo.Text = TxtInsuranceCo
Else
    TxtInsuranceCo.Visible = False: CmbInsuranceCo.Visible = True:  CmdInsuranceCo.Visible = True
    CmbInsuranceCo.Text = TxtInsuranceCo
End If

If Val(mInsuranceRiskTypeID) > 0 Then
    TxtInsuranceRiskType.Visible = True: CmbInsuranceRiskTypeID.Visible = False
    CmbInsuranceRiskTypeID.Text = TxtInsuranceRiskType
Else
    TxtInsuranceRiskType.Visible = False: CmbInsuranceRiskTypeID.Visible = True
    CmbInsuranceRiskTypeID.Text = TxtInsuranceRiskType
End If

If Val(mInsuranceByID) > 0 Then
    TxtInsuranceBy.Visible = True: CmbInsuranceByID.Visible = False
    CmbInsuranceByID.Text = TxtInsuranceBy
Else
    TxtInsuranceBy.Visible = False: CmbInsuranceByID.Visible = True
    CmbInsuranceByID.Text = TxtInsuranceBy
End If
''''----------------------------------

''TxtInsuranceCo.Visible = False: CmbInsuranceCo.Visible = True:  CmdInsuranceCo.Visible = True
''CmbInsuranceCo.Text = TxtInsuranceCo
''TxtInsuranceRiskType.Visible = False: CmbInsuranceRiskTypeID.Visible = True
''CmbInsuranceRiskTypeID.Text = TxtInsuranceRiskType
''TxtInsuranceBy.Visible = False: CmbInsuranceByID.Visible = True
''CmbInsuranceByID.Text = TxtInsuranceBy

TxtAmount.Locked = False
TxtStructureWiseLimit.Locked = False
ValidFrom.Enabled = True
ValidTo.Enabled = True
TxtRemark.Locked = False
TxtStatus.Visible = False: CmbStatus.Visible = True
CmbStatus.Text = TxtStatus.Text

Call FillUnitType       ''-- Unit Type
Frame6.Enabled = True   ''-- Unit Type
Frame7.Enabled = True   ''-- Location Details
Frame2.Enabled = True   ''-- Godown Mapping

If mCommodityChangeFlag = True Then
    Frame3.Enabled = True   ''-- Include Commodity
    Frame4.Enabled = True   ''-- Exclude Commodity
Else
    Frame3.Enabled = False   ''-- Include Commodity
    Frame4.Enabled = False   ''-- Exclude Commodity
End If

If mClientChangeFlag = True Then
    Frame5.Enabled = True   ''-- Client Details
Else
    Frame5.Enabled = False   ''-- Client Details
End If

Dim mR As Double
mListOfSelectedUnitType = 0
With MSHFlexGrid6
    If .Rows > 2 Then
        For mR = 1 To .Rows - 2
           If .TextMatrix(mR, 0) = strChecked Then
            mListOfSelectedUnitType = mListOfSelectedUnitType & "," & Val(.TextMatrix(mR, 2))
           End If
        Next mR
    End If
End With
SSTab1.Tab = 0
Call GButtonLock(Me, False)
End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub FirstCmd_Click()
RsMst_PolicyDetail.MoveFirst
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
   Exit Sub
End If
mExchangeTypeID = "1"
Call Grid_Head
Label12.Caption = "-"
TxtSPolicyNo = ""
TxtSCoverNoteNo = ""
CmbSInsuranceCo = ""
CmbSInsuranceRiskTypeID = ""
CmbSInsuranceByID = ""
CmbSStatus = ""

Call TableMst_PolicyDetail(" Order By InsurancePolicyID ")
If RsMst_PolicyDetail.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
If RsMst_PolicyDetail.RecordCount > 0 Then RsMst_PolicyDetail.MoveLast
Call Indata
Call GButtonLock(Me, True)

End Sub

Public Sub Indata()
TxtInsurancePolicyID = ""
TxtPolicyNo = "":               TxtPolicyNo.Locked = True
TxtCoverNoteNo = "":            TxtCoverNoteNo.Locked = True
TxtInsuranceCo = "":            TxtInsuranceCo.Visible = True
CmbInsuranceCo = "":            CmbInsuranceCo.Visible = False: CmdInsuranceCo.Visible = False
TxtInsuranceRiskType = "":      TxtInsuranceRiskType.Visible = True
CmbInsuranceRiskTypeID = "":    CmbInsuranceRiskTypeID.Visible = False
TxtInsuranceBy = "":            TxtInsuranceBy.Visible = True
CmbInsuranceByID = "":          CmbInsuranceByID.Visible = False
TxtAmount = "":                 TxtAmount.Locked = True
TxtStructureWiseLimit = "":     TxtStructureWiseLimit.Locked = True
ValidFrom.Enabled = False:      ValidFrom.Value = Date
ValidTo.Enabled = False:        ValidTo.Value = Date
TxtStatus.Text = "":            TxtStatus.Visible = True
CmbStatus.Text = "":            CmbStatus.Visible = False
TxtCreated = ""
TxtUpdated = ""
TxtGUID = ""

''--- Frame 6 Godown Type details
Call Grid_Head6
''--- Frame 7 Exclude Location details
Call ClearFrame7
Call Grid_Head7
Frame7.Enabled = False
''--- Frame 2 Godown Mapping details
Call ClearFrame2
Call Grid_Head2
Frame2.Enabled = False
''--- Frame 3 Include Commodity details
Call ClearFrame3
Call Grid_Head3
Frame3.Enabled = False
''--- Frame 4 Exclude Commodity details
Call ClearFrame4
Call Grid_Head4
Frame4.Enabled = False
''--- Frame 5 Client details
CmbExchange = ""
Call ClearFrame5
Call Grid_Head5
Frame5.Enabled = False
''----------------------------------
TxtRemark = "":                 TxtRemark.Locked = True
SSTab1.Tab = 0

TxtInsurancePolicyID = IIf(IsNull(RsMst_PolicyDetail!InsurancePolicyID), "", RsMst_PolicyDetail!InsurancePolicyID)
TxtPolicyNo = IIf(IsNull(RsMst_PolicyDetail!PolicyNo), "", RsMst_PolicyDetail!PolicyNo)
TxtCoverNoteNo = IIf(IsNull(RsMst_PolicyDetail!CoverNoteNo), "", RsMst_PolicyDetail!CoverNoteNo)
TxtGUID = IIf(IsNull(RsMst_PolicyDetail!G_UID), "", RsMst_PolicyDetail!G_UID)
mInsuranceCoID = IIf(IsNull(RsMst_PolicyDetail!InsuranceCoID), 0, RsMst_PolicyDetail!InsuranceCoID)
TxtInsuranceCo = GetInsuranceCoName(mInsuranceCoID)
mInsuranceRiskTypeID = IIf(IsNull(RsMst_PolicyDetail!InsuranceRiskCoveredID), 0, RsMst_PolicyDetail!InsuranceRiskCoveredID)
TxtInsuranceRiskType = GetInsuranceRiskCovered(mInsuranceRiskTypeID)
mInsuranceByID = IIf(IsNull(RsMst_PolicyDetail!InsuranceByID), 0, RsMst_PolicyDetail!InsuranceByID)
TxtInsuranceBy = GetInsuranceBy(mInsuranceByID)
TxtAmount = IIf(IsNull(RsMst_PolicyDetail!Amount), "", Format(RsMst_PolicyDetail!Amount, "#0.00"))
TxtStructureWiseLimit = IIf(IsNull(RsMst_PolicyDetail!StructureWiseLimit), "", Format(RsMst_PolicyDetail!StructureWiseLimit, "#0.00"))
ValidFrom.Value = IIf(IsNull(RsMst_PolicyDetail!ValidFrom), Date, RsMst_PolicyDetail!ValidFrom)
ValidTo.Value = IIf(IsNull(RsMst_PolicyDetail!ValidTo), Date, RsMst_PolicyDetail!ValidTo)
TxtStatus.Text = IIf(IIf(IsNull(RsMst_PolicyDetail!Flag), "A", RsMst_PolicyDetail!Flag) = "A", "Active", "Deactive")
TxtRemark = IIf(IsNull(RsMst_PolicyDetail!Remark), "", RsMst_PolicyDetail!Remark)
TxtCreated = IIf(IsNull(RsMst_PolicyDetail!CreatedBy), "", RsMst_PolicyDetail!CreatedBy) & " :- " & IIf(IsNull(RsMst_PolicyDetail!CreatedDate), "", RsMst_PolicyDetail!CreatedDate)
TxtUpdated = IIf(IsNull(RsMst_PolicyDetail!UpdatedBy), "", RsMst_PolicyDetail!UpdatedBy) & " :- " & IIf(IsNull(RsMst_PolicyDetail!UpdatedDate), "", RsMst_PolicyDetail!UpdatedDate)

If mInsuranceByID = 2 Then SSTab1.TabEnabled(3) = True Else SSTab1.TabEnabled(3) = False

Call IndataInsuranceUnitType
Call IndataInsuranceValidFor
Call IndataInsuranceGodownMapping
Call IndataInsuranceCommodity
Call IndataInsuranceClientDetail

Call GButtonLock(Me, True)
End Sub

Private Sub IndataInsuranceValidFor()
Dim mR As Double, mRow As Double
With MSHFlexGrid7

    tmp = "  Select A.InsuranceValidForID,A.InsurancePolicyID,A.ValidForType,A.ValidForTypeID,"
    tmp = tmp & "  Case a.ValidForType"
    tmp = tmp & "  When 'R' Then ( Select R.Region from Mst_Region R Where R.RegionID = A.ValidForTypeID )"
    tmp = tmp & "  When 'S' Then ( Select S.StateName from Mst_State S Where S.StateID = A.ValidForTypeID )"
    tmp = tmp & "  When 'L' Then ( Select L.LocationName from Mst_Location L Where L.LocationID = A.ValidForTypeID )"
''    tmp = tmp & "  When 'C' Then ( Select C.CMPName from Mst_CMP C Where C.CMPID  = A.ValidForTypeID )"
''    tmp = tmp & "  When 'G' Then ( Select G.GodownNo from Mst_Godown G Where G.GodownID = A.ValidForTypeID )"
    tmp = tmp & "  Else '' End [Name]"
    tmp = tmp & "  from Mst_InsuranceValidFor As A"
    tmp = tmp & "  Where A.InsurancePolicyID  = " & Val(TxtInsurancePolicyID) & ""
    Call TmpTable
    If TmpRs.RecordCount > 0 Then
       For mR = 1 To TmpRs.RecordCount
       
         .TextMatrix(mR, 0) = mR
         .TextMatrix(mR, 1) = IIf(IsNull(TmpRs!ValidForType), "", TmpRs!ValidForType)
            If LCase(.TextMatrix(mR, 1)) = "r" Then .TextMatrix(mR, 1) = "Region"
            If LCase(.TextMatrix(mR, 1)) = "s" Then .TextMatrix(mR, 1) = "State"
            If LCase(.TextMatrix(mR, 1)) = "l" Then .TextMatrix(mR, 1) = "Location"
            ''If LCase(.TextMatrix(mR, 1)) = "c" Then .TextMatrix(mR, 1) = "CMP"
            ''If LCase(.TextMatrix(mR, 1)) = "g" Then .TextMatrix(mR, 1) = "Godown"
         .TextMatrix(mR, 2) = IIf(IsNull(TmpRs!Name), "", TmpRs!Name)
         .TextMatrix(mR, 3) = IIf(IsNull(TmpRs!ValidForTypeID), "", TmpRs!ValidForTypeID)
         .TextMatrix(mR, 4) = IIf(IsNull(TmpRs!InsuranceValidForID), "", TmpRs!InsuranceValidForID)
         
         .Rows = .Rows + 1
         TmpRs.MoveNext
       Next mR
    End If
    
End With

End Sub

Private Sub IndataInsuranceGodownMapping()
Dim mR As Double, mRow As Double
With MSHFlexGrid2

    tmp = " Select A.InsuranceGodownMappingID,A.InsurancePolicyID,A.GodownID,MG.GodownNo,MUT.UnitType,MG.UnitTypeID,MG.CMPID,CMP.CMPName,ML.LocationID,ML.LocationName,MS.StateID,MS.StateName,MR.RegionID,MR.Region,MG.Address,MG.GodownCapacity,A.Connected"
    tmp = tmp & " from Mst_InsuranceGodownMapping A"
    tmp = tmp & " Inner Join Mst_Godown MG On A.GodownID = MG.GodownID"
    tmp = tmp & " Inner Join Mst_UnitType MUT On MG.UnitTypeID = MUT.UnitTypeID"
    tmp = tmp & " Inner Join Mst_CMP CMP On MG.CMPID = CMP.CMPID"
    tmp = tmp & " Inner Join Mst_Location ML On CMP.LocationID = ML.LocationID"
    tmp = tmp & " Inner Join Mst_State MS On ML.StateID = MS.StateID"
    tmp = tmp & " Inner Join Mst_Region MR On MS.RegionID = MR.RegionID"
    tmp = tmp & " Where a.InsurancePolicyID = " & Val(TxtInsurancePolicyID)
    tmp = tmp & " Order By ML.LocationName,CMP.CMPName,MG.GodownNo "
    Call TmpTable
    
    If TmpRs.RecordCount > 0 Then
       For mR = 1 To TmpRs.RecordCount
       
         .row = mR
         .Col = 0
         .CellFontName = "Wingdings"
         .CellFontSize = 12
         .CellAlignment = flexAlignCenterCenter
         .Text = strChecked
         
         .TextMatrix(mR, 1) = IIf(IsNull(TmpRs!GodownNo), "", TmpRs!GodownNo)                                    ''"Godown No"
         .TextMatrix(mR, 2) = IIf(IsNull(TmpRs!UnitType), "", TmpRs!UnitType)                                    ''"Unit Type"
         .TextMatrix(mR, 3) = IIf(IsNull(TmpRs!CMPName), "", TmpRs!CMPName)                                      ''"CMP Name"
         .TextMatrix(mR, 4) = IIf(IsNull(TmpRs!LocationName), "", TmpRs!LocationName)                            ''"Location"
         .TextMatrix(mR, 5) = IIf(IsNull(TmpRs!StateName), "", TmpRs!StateName)                                  ''"State"
         .TextMatrix(mR, 6) = IIf(IsNull(TmpRs!Region), "", TmpRs!Region)                                        ''"Region"
         .TextMatrix(mR, 7) = IIf(IsNull(TmpRs!Address), "", TmpRs!Address)                                      ''"Address"
         .TextMatrix(mR, 8) = Format(IIf(IsNull(TmpRs!GodownCapacity), "", TmpRs!GodownCapacity), "0.000")        ''"Capacity"
         .TextMatrix(mR, 10) = IIf(IsNull(TmpRs!GodownID), "", TmpRs!GodownID)                                    ''"GodownID"
         .TextMatrix(mR, 11) = IIf(IsNull(TmpRs!UnitTypeID), "", TmpRs!UnitTypeID)                               ''"UnitID"
         .TextMatrix(mR, 12) = IIf(IsNull(TmpRs!InsuranceGodownMappingID), "", TmpRs!InsuranceGodownMappingID)   ''"DetailID"
         'Call GRowColor(MSHFlexGrid2, mR, vbGreen)
         
         ''--- Connected (Yes / No)
         .row = mR
         .Col = 9
         .CellFontName = "Wingdings"
         .CellFontSize = 12
         .CellAlignment = flexAlignCenterCenter
         If LCase(TmpRs!Connected) = "y" Then
            .Text = strChecked
            Call GRowColor(MSHFlexGrid2, mR, vbGreen)
         Else
            .Text = strUnChecked
            Call GRowColor(MSHFlexGrid2, mR, vbYellow)
         End If
         ''---
                  
         .Rows = .Rows + 1
         TmpRs.MoveNext
       Next mR
    End If
ChkSelectAll.Tag = "NoClick"
ChkSelectAll.Value = vbChecked
ChkSelectAll.Tag = ""

Label35.Caption = CountSelectedGodown & " / " & MSHFlexGrid2.Rows - 2
Label39.Caption = CountSelectedConnectedGodown & " / " & MSHFlexGrid2.Rows - 2
    
End With
End Sub


Private Sub IndataInsuranceCommodity()
Dim mR As Double, mRow As Double
With MSHFlexGrid3

    tmp = "  Select A.InsuranceCommodityDetailID,A.InsurancePolicyID,A.CommodityID,A.EntryType,B.Commodity"
    tmp = tmp & " From Mst_InsuranceCommodityDetail As A"
    tmp = tmp & " Inner Join Mst_Commodity AS B On A.CommodityID = B.CommodityID"
    tmp = tmp & " Where A.InsurancePolicyID  = " & Val(TxtInsurancePolicyID) & " And A.EntryType = 'I' "
    Call TmpTable
    
    If TmpRs.RecordCount > 0 Then
       For mR = 1 To TmpRs.RecordCount
       
         .TextMatrix(mR, 0) = mR
         .TextMatrix(mR, 1) = IIf(IsNull(TmpRs!Commodity), "", TmpRs!Commodity)
         .TextMatrix(mR, 2) = IIf(IsNull(TmpRs!CommodityID), "", TmpRs!CommodityID)
         .TextMatrix(mR, 3) = IIf(IsNull(TmpRs!InsuranceCommodityDetailID), "", TmpRs!InsuranceCommodityDetailID)
         
         .Rows = .Rows + 1
         TmpRs.MoveNext
       Next mR
    End If
    
End With

With MSHFlexGrid4

    tmp = "  Select A.InsuranceCommodityDetailID,A.InsurancePolicyID,A.CommodityID,A.EntryType,B.Commodity"
    tmp = tmp & " From Mst_InsuranceCommodityDetail As A"
    tmp = tmp & " Inner Join Mst_Commodity AS B On A.CommodityID = B.CommodityID"
    tmp = tmp & " Where A.InsurancePolicyID  = " & Val(TxtInsurancePolicyID) & " And A.EntryType = 'E' "
    Call TmpTable
    
    If TmpRs.RecordCount > 0 Then
       For mR = 1 To TmpRs.RecordCount
       
         .TextMatrix(mR, 0) = mR
         .TextMatrix(mR, 1) = IIf(IsNull(TmpRs!Commodity), "", TmpRs!Commodity)
         .TextMatrix(mR, 2) = IIf(IsNull(TmpRs!CommodityID), "", TmpRs!CommodityID)
         .TextMatrix(mR, 3) = IIf(IsNull(TmpRs!InsuranceCommodityDetailID), "", TmpRs!InsuranceCommodityDetailID)
         
         .Rows = .Rows + 1
         TmpRs.MoveNext
       Next mR
    End If
    
End With

End Sub


Private Sub IndataInsuranceClientDetail()
Dim mR As Double, mRow As Double
With MSHFlexGrid5
    '',C.ExchangeType
    tmp = "  Select A.InsuranceClientDetailID,A.InsurancePolicyID,A.ClientIDRow,B.ClientID,B.ClientName "
    tmp = tmp & "  from Mst_InsuranceClientDetail As A "
    tmp = tmp & "  Inner Join Mst_Client as B On A.ClientIDRow = B.ClientIDRow "
    ''tmp = tmp & "  Inner Join Mst_ExchangeType as C On B.ExchangeTypeID = C.ExchangeTypeID "
    tmp = tmp & "  Where A.InsurancePolicyID = " & Val(TxtInsurancePolicyID)
    Call TmpTable
    If TmpRs.RecordCount > 0 Then
       For mR = 1 To TmpRs.RecordCount
       
         .TextMatrix(mR, 0) = mR
         .TextMatrix(mR, 1) = IIf(IsNull(TmpRs!ClientID), "", TmpRs!ClientID)
         .TextMatrix(mR, 2) = IIf(IsNull(TmpRs!ClientName), "", TmpRs!ClientName)
         .TextMatrix(mR, 3) = IIf(IsNull(TmpRs!ClientIDRow), "", TmpRs!ClientIDRow)
         .TextMatrix(mR, 4) = IIf(IsNull(TmpRs!InsuranceClientDetailID), "", TmpRs!InsuranceClientDetailID)
         ''.TextMatrix(mR, 5) = IIf(IsNull(TmpRs!ExchangeType), "", TmpRs!ExchangeType)
         ''.TextMatrix(mR, 6) = IIf(IsNull(TmpRs!ExchangeTypeID), 0, TmpRs!ExchangeTypeID)
         
         .Rows = .Rows + 1
         TmpRs.MoveNext
       Next mR
    End If
    
End With
End Sub


Private Sub IndataInsuranceUnitType()
Dim mR As Double, mRow As Double
With MSHFlexGrid6

    tmp = "  Select A.InsuranceForUnitTypeDetailID,A.InsurancePolicyID,A.UnitTypeID,B.UnitType"
    tmp = tmp & "  from Mst_InsuranceForUnitTypeDetail As A"
    tmp = tmp & "  Inner Join Mst_UnitType B On A.UnitTypeID = B.UnitTypeID"
    tmp = tmp & "  Where A.InsurancePolicyID = " & Val(TxtInsurancePolicyID)
    Call TmpTable
    If TmpRs.RecordCount > 0 Then
       For mR = 1 To TmpRs.RecordCount
       
         .row = mR
         .Col = 0
         .CellFontName = "Wingdings"
         .CellFontSize = 12
         .CellAlignment = flexAlignCenterCenter
         .Text = strChecked
         .TextMatrix(mR, 1) = IIf(IsNull(TmpRs!UnitType), "", TmpRs!UnitType)
         .TextMatrix(mR, 2) = IIf(IsNull(TmpRs!UnitTypeID), "", TmpRs!UnitTypeID)
         .TextMatrix(mR, 3) = IIf(IsNull(TmpRs!InsuranceForUnitTypeDetailID), "", TmpRs!InsuranceForUnitTypeDetailID)
         
         .Rows = .Rows + 1
         TmpRs.MoveNext
       Next mR
    End If
    
End With
End Sub



Private Sub LastCmd_Click()
RsMst_PolicyDetail.MoveLast
Call Indata
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False

End Sub

Private Sub MSHFlexGrid1_Click()
With MSHFlexGrid1
    RsMst_PolicyDetail.MoveFirst
    If RsMst_PolicyDetail.RecordCount = 0 Then
       Call AddCmd_Click
       Exit Sub
    End If
    RsMst_PolicyDetail.MoveFirst
    RsMst_PolicyDetail.Find "InsurancePolicyID = " & Val(.TextMatrix(.RowSel, 0))
    If RsMst_PolicyDetail.EOF Then
       RsMst_PolicyDetail.MoveLast
    End If
    Frame0.Visible = False
    Call Indata
End With
End Sub

Private Sub MSHFlexGrid2_Click()
Dim mR As Double

If SaveCmd.Enabled = False Then Exit Sub
With MSHFlexGrid2
    If .Rows > 2 Then
        mR = .RowSel
        If .TextMatrix(mR, 1) = "" Then Exit Sub
        If .ColSel = 0 Then
           mR = .RowSel
            If .TextMatrix(mR, 0) = strChecked Then
''                .row = mR
''                .Col = 0
''                If .CellBackColor = vbGreen And Val(.TextMatrix(mR, 11)) > 0 Then
''                    Call GRowColor(MSHFlexGrid2, mR, vbYellow)
''                    .TextMatrix(mR, 0) = strUnChecked
''                Else
''                    .TextMatrix(mR, 0) = strUnChecked
''                End If
                If Val(.TextMatrix(mR, 12)) = 0 Then
                    .TextMatrix(mR, 0) = strUnChecked
                End If
            Else
''                ''-- if Cell color is red the not allow to select that godown
''                .row = mR
''                .Col = 0
''                If .CellBackColor = vbYellow And Val(.TextMatrix(mR, 11)) > 0 Then
''                    Call GRowColor(MSHFlexGrid2, mR, vbGreen)
''                    .TextMatrix(mR, 0) = strChecked
''                Else
''                    If .CellBackColor <> vbRed Then
''                        .TextMatrix(mR, 0) = strChecked
''                    End If
''                End If
                .row = mR
                .Col = 0
                If Val(.TextMatrix(mR, 12)) = 0 Then
                    If .CellBackColor <> vbRed Then
                        .TextMatrix(mR, 9) = strChecked
                        .TextMatrix(mR, 0) = strChecked
                    End If
                End If
                
           End If
        ElseIf .ColSel = 9 Then
            mR = .RowSel
            If .TextMatrix(mR, 9) = strChecked Then
                .row = mR
                .Col = 9
                If .CellBackColor = vbGreen And Val(.TextMatrix(mR, 12)) > 0 Then
                    Call GRowColor(MSHFlexGrid2, mR, vbYellow)
                    .TextMatrix(mR, 9) = strUnChecked
                Else
                    .TextMatrix(mR, 9) = strUnChecked
                    .TextMatrix(mR, 0) = strUnChecked
                End If
            Else
                ''-- if Cell color is red the not allow to select that godown
                .row = mR
                .Col = 9
                If .CellBackColor = vbYellow And Val(.TextMatrix(mR, 12)) > 0 Then
                    Call GRowColor(MSHFlexGrid2, mR, vbGreen)
                    .TextMatrix(mR, 9) = strChecked
                Else
                    If .CellBackColor <> vbRed Then
                        .TextMatrix(mR, 9) = strChecked
                        .TextMatrix(mR, 0) = strChecked
                    End If
                End If
           End If
        End If
    End If
End With

If CheckAllGodownSelect = True Then
    ChkSelectAll.Tag = "NoClick"
    ChkSelectAll = 1
    ChkSelectAll.Tag = ""
Else
    ChkSelectAll.Tag = "NoClick"
    ChkSelectAll = 0
    ChkSelectAll.Tag = ""
End If

Label35.Caption = CountSelectedGodown & " / " & MSHFlexGrid2.Rows - 2
Label39.Caption = CountSelectedConnectedGodown & " / " & MSHFlexGrid2.Rows - 2

''With MSHFlexGrid2
''If SaveCmd.Enabled = False Then Exit Sub
''mR = .RowSel
''If mR < 1 Or mR = .Rows - 1 Then
''   Call ClearFrame2
''   Exit Sub
''End If
''If .TextMatrix(mR, 1) = "" Then Exit Sub
''
''CmbValidType = .TextMatrix(mR, 1)
''CmbValidTypeID = .TextMatrix(mR, 2)
''mValidType = Left$(.TextMatrix(mR, 1), 1)
''mValidTypeID = .TextMatrix(mR, 3)
''Label19.Caption = CmbValidType.Text
''CmdAddtoListFrame2.Caption = "Update List"
''CmdDeleteFrame2.Enabled = True
''
''CmbLocationID.Text = ""
''CmbCMPID.Text = ""
''
''If LCase(CmbValidType.Text) = LCase("CMP") Then    ''-- CMP
''    tmp = " Select C.CMPID,C.CMPName,C.CMPName+'~'+convert(Varchar,C.CMPID)[CMP],C.Address,L.LocationName,L.LocationID from Mst_CMP C Inner Join Mst_Location L On C.LocationID = L.LocationID Where C.CMPID = " & mValidTypeID & " "
''    Call TmpTable
''    If TmpRs.RecordCount > 0 Then
''        lblLocation.Visible = True
''        CmbLocationID = "": CmbLocationID.Visible = True
''        lblCMP.Visible = False
''        CmbCMPID = "": CmbCMPID.Visible = False
''
''        CmbLocationID.Text = IIf(IsNull(TmpRs!LocationName), "", TmpRs!LocationName)
''        mLocationID = IIf(IsNull(TmpRs!LocationID), 0, TmpRs!LocationID)
''
''        CmbValidTypeID.Text = IIf(IsNull(TmpRs!CMP), "", TmpRs!CMP)
''    End If
''ElseIf LCase(CmbValidType.Text) = LCase("Godown") Then    ''-- Godown
''    tmp = " Select CMP.CMPName,MG.CMPID,MG.GodownID,MG.GodownNo,MG.GodownNo+'~'+convert(Varchar,MG.GodownID)[Godown],MG.Address,L.LocationName,L.LocationID from Mst_Godown MG Inner Join Mst_CMP CMP On MG.CMPID = CMP.CMPID Inner Join Mst_Location L On CMP.LocationID = L.LocationID Where MG.GodownID = " & mValidTypeID & ""
''    Call TmpTable
''    If TmpRs.RecordCount > 0 Then
''
''        lblLocation.Visible = True
''        CmbLocationID = "": CmbLocationID.Visible = True
''        lblCMP.Visible = True
''        CmbCMPID = "": CmbCMPID.Visible = True
''
''        CmbLocationID.Text = IIf(IsNull(TmpRs!LocationName), "", TmpRs!LocationName)
''        mLocationID = IIf(IsNull(TmpRs!LocationID), 0, TmpRs!LocationID)
''        CmbCMPID.Text = IIf(IsNull(TmpRs!CMPName), "", TmpRs!CMPName)
''        mCMPID = IIf(IsNull(TmpRs!CMPID), 0, TmpRs!CMPID)
''
''        CmbValidTypeID.Text = IIf(IsNull(TmpRs!Godown), "", TmpRs!Godown)
''    End If
''Else
''    lblLocation.Visible = False
''    CmbLocationID = "": CmbLocationID.Visible = False
''    lblCMP.Visible = False
''    CmbCMPID = "": CmbCMPID.Visible = False
''End If
''
''End With
End Sub

Private Sub MSHFlexGrid3_Click()
Dim mR As Variant
With MSHFlexGrid3
    If SaveCmd.Enabled = False And Frame3.Enabled = False Then Exit Sub
    mR = .RowSel
    If mR < 1 Or mR = .Rows - 1 Then
       Call ClearFrame3
       Exit Sub
    End If
    If .TextMatrix(mR, 1) = "" Then Exit Sub
    CmbIncludeCommodityID = .TextMatrix(mR, 1)
    mIncludeCommodityID = .TextMatrix(mR, 2)
    CmdAddFrame3.Caption = "Update List"
    CmdDeleteFrame3.Enabled = True
    Call ClearFrame4
End With
End Sub

Private Sub MSHFlexGrid4_Click()
Dim mR As Variant
With MSHFlexGrid4
    If SaveCmd.Enabled = False And Frame4.Enabled = False Then Exit Sub
    mR = .RowSel
    If mR < 1 Or mR = .Rows - 1 Then
       Call ClearFrame3
       Exit Sub
    End If
    If .TextMatrix(mR, 1) = "" Then Exit Sub
    CmbExcludeCommodityID = .TextMatrix(mR, 1)
    mExcludeCommodityID = .TextMatrix(mR, 2)
    CmdAddFrame4.Caption = "Update List"
    CmdDeleteFrame4.Enabled = True
    Call ClearFrame3
End With
End Sub

Private Sub MSHFlexGrid5_Click()
Dim mR As Variant
With MSHFlexGrid5
    If SaveCmd.Enabled = False And Frame5.Enabled = False Then Exit Sub
    mR = .RowSel
    If mR < 1 Or mR = .Rows - 1 Then
       Call ClearFrame5
       Exit Sub
    End If
    If .TextMatrix(mR, 1) = "" Then Exit Sub
    'mExchangeTypeID = .TextMatrix(mR, 6)
    'CmbExchange = .TextMatrix(mR, 5)
    mClientRowID = .TextMatrix(mR, 3)
    TxtClientID = .TextMatrix(mR, 1)
    CmbClientRowID = .TextMatrix(mR, 2) & "#" & .TextMatrix(mR, 1) & "#" & mClientRowID
    CmdAddFrame5.Caption = "Update List"
    CmdDeleteFrame5.Enabled = True
End With
End Sub

Private Sub MSHFlexGrid6_Click()
Dim mR As Double
mListOfSelectedUnitType = 0

If SaveCmd.Enabled = False Then Exit Sub
With MSHFlexGrid6
    If .Rows > 2 And .ColSel = 0 Then
       mR = .RowSel
       If .TextMatrix(mR, 1) = "" Then Exit Sub
        If .TextMatrix(mR, 0) = strChecked Then
            ''-- If Unit Type Already mapped and savewith insurance then not allow to Edit/Remove
            If Val(.TextMatrix(mR, 3)) = 0 Then
                .TextMatrix(mR, 0) = strUnChecked
            End If
        Else
            .TextMatrix(mR, 0) = strChecked
       End If
       
        For mR = 1 To .Rows - 2
           If .TextMatrix(mR, 0) = strChecked Then
            mListOfSelectedUnitType = mListOfSelectedUnitType & "," & Val(.TextMatrix(mR, 2))
           End If
        Next mR
       
       'Call LocationCheckOnDeselectUnitType
        
    End If
End With

End Sub

Private Sub LocationCheckOnDeselectUnitType()
''-- Check the select Region/State/Location List i.e. is in selected Unit Type
''-- if Not then remove from  list

Dim mR As Double, mRLastRow As Double
Dim mRow As Double, mCol As Double
Dim mValidT As Variant, temp As Variant

With MSHFlexGrid7
    If .Rows > 2 Then
    
            tmp = " Select 'Region'VType,MR.RegionID,MR.Region from Mst_Godown MG Inner Join Mst_CMP CMP On MG.CMPID = CMP.CMPID Inner Join Mst_Location ML On CMP.LocationID = ML.LocationID Inner Join Mst_State MS On ML.StateID = MS.StateID Inner Join Mst_Region MR On MS.RegionID = MR.RegionID Where MG.CloseDate Is Null And MG.UnitTypeID In (" & mListOfSelectedUnitType & ") Group By MR.RegionID,MR.Region"
            tmp = tmp & " Union"
            tmp = tmp & " Select 'State'VType,MS.StateID,MS.StateName from Mst_Godown MG Inner Join Mst_CMP CMP On MG.CMPID = CMP.CMPID Inner Join Mst_Location ML On CMP.LocationID = ML.LocationID Inner Join Mst_State MS On ML.StateID = MS.StateID Inner Join Mst_Region MR On MS.RegionID = MR.RegionID Where MG.CloseDate Is Null And MG.UnitTypeID In (" & mListOfSelectedUnitType & ") Group By MS.StateID,MS.StateName"
            tmp = tmp & " Union"
            tmp = tmp & " Select 'Location'VType,ML.LocationID,ML.LocationName from Mst_Godown MG Inner Join Mst_CMP CMP On MG.CMPID = CMP.CMPID Inner Join Mst_Location ML On CMP.LocationID = ML.LocationID Inner Join Mst_State MS On ML.StateID = MS.StateID Inner Join Mst_Region MR On MS.RegionID = MR.RegionID Where MG.CloseDate Is Null And MG.UnitTypeID In (" & mListOfSelectedUnitType & ") Group By ML.LocationID,ML.LocationName"
            Call Tmp1Table
            
            If TmpRs1.RecordCount > 0 Then
               mRLastRow = .Rows - 2
               For mR = mRLastRow To 1 Step -1
               
                If Trim(.TextMatrix(mR, 1)) <> "" Then
                    temp = ""
                    If LCase(.TextMatrix(mR, 1)) = "region" Then temp = " VType = 'Region' AND RegionID = " & Val(.TextMatrix(mR, 3)) & " AND Region = '" & .TextMatrix(mR, 2) & "'"
                    If LCase(.TextMatrix(mR, 1)) = "state" Then temp = " VType = 'State' AND RegionID = " & Val(.TextMatrix(mR, 3)) & " AND Region = '" & .TextMatrix(mR, 2) & "'"
                    If LCase(.TextMatrix(mR, 1)) = "location" Then temp = " VType = 'Location' AND RegionID = " & Val(.TextMatrix(mR, 3)) & " AND Region = '" & .TextMatrix(mR, 2) & "'"
                    TmpRs1.Filter = ""
                    If temp <> "" Then
                        TmpRs1.MoveFirst
                        TmpRs1.Filter = temp
                        If TmpRs1.EOF = True Then
                            ''-->> remove the Row , if not found in Record set
                            For mRow = mR To .Rows - 2
                                For mCol = 0 To .Cols - 1
                                    .TextMatrix(mRow, mCol) = .TextMatrix(mRow + 1, mCol)
                                Next
                            Next
                            .Rows = .Rows - 1
                            ''-->>
                        End If
                    End If
        
                End If
               Next mR
                
            Else
            
                ''----------->>>>>>>>>>>>>>
                '' Remove all Region/State/Location if no record not with Selected Unit Type
                mRLastRow = .Rows - 2
                For mR = mRLastRow To 1 Step -1
                        For mRow = mR To .Rows - 2
                            For mCol = 0 To .Cols - 1
                                .TextMatrix(mRow, mCol) = .TextMatrix(mRow + 1, mCol)
                            Next
                        Next
                        .Rows = .Rows - 1
                Next mR
                ''----------->>>>>>>>>>>>>>
            End If
    
    
    
    
    

    End If
End With
End Sub


Private Sub MSHFlexGrid7_Click()
Dim mR As Variant

With MSHFlexGrid7
If SaveCmd.Enabled = False Then Exit Sub
mR = .RowSel
If mR < 1 Or mR = .Rows - 1 Then
   Call ClearFrame7
   Exit Sub
End If
If .TextMatrix(mR, 1) = "" Then Exit Sub

CmbValidType = .TextMatrix(mR, 1)
CmbValidTypeID = .TextMatrix(mR, 2)
mValidType = Left$(.TextMatrix(mR, 1), 1)
mValidTypeID = .TextMatrix(mR, 3)
Label19.Caption = CmbValidType.Text
CmdAddtoListFrame7.Caption = "Update List"
CmdDeleteFrame7.Enabled = True

''If LCase(CmbExcludeValidType.Text) = LCase("CMP") Then    ''-- CMP
''    tmp = " Select C.CMPID,C.CMPName,C.CMPName+'~'+convert(Varchar,C.CMPID)[CMP],C.Address,L.LocationName,L.LocationID from Mst_CMP C Inner Join Mst_Location L On C.LocationID = L.LocationID Where C.CMPID = " & mExcludeValidTypeID & " "
''    Call TmpTable
''    If TmpRs.RecordCount > 0 Then
''        lblExcludeLocation.Visible = True
''        CmbExcludeLocationID = "": CmbExcludeLocationID.Visible = True
''        lblExcludeCMP.Visible = False
''        CmbExcludeCMPID = "": CmbExcludeCMPID.Visible = False
''
''        CmbExcludeLocationID.Text = IIf(IsNull(TmpRs!LocationName), "", TmpRs!LocationName)
''        mExcludeLocationID = IIf(IsNull(TmpRs!LocationID), 0, TmpRs!LocationID)
''
''        CmbExcludeValidTypeID.Text = IIf(IsNull(TmpRs!CMP), "", TmpRs!CMP)
''    End If
''ElseIf LCase(CmbExcludeValidType.Text) = LCase("Godown") Then    ''-- Godown
''    tmp = " Select CMP.CMPName,MG.CMPID,MG.GodownID,MG.GodownNo,MG.GodownNo+'~'+convert(Varchar,MG.GodownID)[Godown],MG.Address,L.LocationName,L.LocationID from Mst_Godown MG Inner Join Mst_CMP CMP On MG.CMPID = CMP.CMPID Inner Join Mst_Location L On CMP.LocationID = L.LocationID Where MG.GodownID = " & mExcludeValidTypeID & ""
''    Call TmpTable
''    If TmpRs.RecordCount > 0 Then
''
''        lblExcludeLocation.Visible = True
''        CmbExcludeLocationID = "": CmbExcludeLocationID.Visible = True
''        lblExcludeCMP.Visible = True
''        CmbExcludeCMPID = "": CmbExcludeCMPID.Visible = True
''
''        CmbExcludeLocationID.Text = IIf(IsNull(TmpRs!LocationName), "", TmpRs!LocationName)
''        mExcludeLocationID = IIf(IsNull(TmpRs!LocationID), 0, TmpRs!LocationID)
''        CmbExcludeCMPID.Text = IIf(IsNull(TmpRs!CMPName), "", TmpRs!CMPName)
''        mExcludeCMPID = IIf(IsNull(TmpRs!CMPID), 0, TmpRs!CMPID)
''
''        CmbExcludeValidTypeID.Text = IIf(IsNull(TmpRs!Godown), "", TmpRs!Godown)
''    End If
''Else
''    lblExcludeLocation.Visible = False
''    CmbExcludeLocationID = "": CmbExcludeLocationID.Visible = False
''    lblExcludeCMP.Visible = False
''    CmbExcludeCMPID = "": CmbExcludeCMPID.Visible = False
''End If

End With

End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_PolicyDetail.MoveNext
If RsMst_PolicyDetail.EOF Then
   RsMst_PolicyDetail.MoveLast
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

RsMst_PolicyDetail.MovePrevious
If RsMst_PolicyDetail.BOF Then
   RsMst_PolicyDetail.MoveFirst
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
Dim temp As Variant
If SaveCmd.Enabled = False Then Exit Sub

If TxtPolicyNo = "" Then
   MsgBox "Blank Policy No not Allowed"
   TxtPolicyNo.SetFocus
   Exit Sub
End If
If TxtCoverNoteNo = "" Then
   MsgBox "Blank Cover Note No not Allowed"
   TxtCoverNoteNo.SetFocus
   Exit Sub
End If
If CmbInsuranceCo.Text = "" Then
   MsgBox "Blank Insurance Co not Allowed"
   CmbInsuranceCo.SetFocus
   Exit Sub
End If
If CmbInsuranceRiskTypeID.Text = "" Then
   MsgBox "Blank Insurance Risk Covered not Allowed"
   CmbInsuranceRiskTypeID.SetFocus
   Exit Sub
End If
If CmbInsuranceByID.Text = "" Then
   MsgBox "Blank Insurance By not Allowed"
   CmbInsuranceByID.SetFocus
   Exit Sub
End If
If TxtAmount = "" Then
   MsgBox "Blank Amount not Allowed"
   TxtAmount.SetFocus
   Exit Sub
End If
If Val(TxtAmount) = 0 Then
   MsgBox "Blank or Zero Amount not Allowed"
   TxtAmount.SetFocus
   Exit Sub
End If
'If TxtStructureWiseLimit = "" Then
'   MsgBox "Blank Structure Wise Limit not Allowed"
'   TxtStructureWiseLimit.SetFocus
'   Exit Sub
'End If
If ValidFrom.Value > ValidTo.Value Then
   MsgBox "'Valid From Date' must be less than 'Valid To Date'!!!"
   ValidFrom.SetFocus
   Exit Sub
End If
If CmbStatus.Text = "" Then
   MsgBox "Select Status."
   CmbStatus.SetFocus
   Exit Sub
End If
If mInsuranceByID = 2 Then
   If MSHFlexGrid5.Rows <= 2 Then
    MsgBox "Blank Client Detail not allowed."
    SSTab1.Tab = 3
    Exit Sub
   End If
End If

If CheckSelectedUnitType = False Then
   MsgBox "Select Godown Type."
   SSTab1.Tab = 0
   Exit Sub
End If
If MSHFlexGrid7.Rows <= 2 Then
   MsgBox "Blank Location Details not allowed."
   SSTab1.Tab = 0
   Exit Sub
End If

If CmdValidate.Enabled = True Then
   MsgBox "Validate the Import Godown List before save."
   SSTab1.Tab = 1
   Exit Sub
End If

If CheckSelectedGodown = False Then
   MsgBox "Select Godowns or Click 'Get Godown List' button to list the godowns."
   SSTab1.Tab = 1
   Exit Sub
End If

If CheckSelectedConnectedGodown = False Then
   MsgBox "No godown is connected to Policy."
   SSTab1.Tab = 1
   Exit Sub
End If

MsgBox ("Wait till next message")

If Edit_Flg = "A" Then
   temp = " PolicyNo = '" & TxtPolicyNo & "' AND CoverNoteNo = '" & TxtCoverNoteNo & "'"
   RsMst_PolicyDetail.Filter = temp
   If Not RsMst_PolicyDetail.EOF Then
      MsgBox "Duplicate Insurance Policy No and Cover Note No not Allowed"
      TxtPolicyNo.SetFocus
      Exit Sub
   End If
   RsMst_PolicyDetail.AddNew
   RsMst_PolicyDetail!InsurancePolicyID = GetMaxID("InsurancePolicyID", "Mst_InsuranceDetails")
   TxtInsurancePolicyID = RsMst_PolicyDetail!InsurancePolicyID
   RsMst_PolicyDetail!G_UID = GetGUID
   TxtGUID = RsMst_PolicyDetail!G_UID
   RsMst_PolicyDetail!PolicyNo = TxtPolicyNo
   RsMst_PolicyDetail!CoverNoteNo = TxtCoverNoteNo
Else
   RsMst_PolicyDetail.MoveFirst
   temp = "PolicyNo = '" & TxtPolicyNo & "' AND CoverNoteNo = '" & TxtCoverNoteNo & "'"
   RsMst_PolicyDetail.Filter = temp
   If RsMst_PolicyDetail.RecordCount > 0 Then
       If TxtGUID <> RsMst_PolicyDetail!G_UID Then
          MsgBox "Duplicate Insurance Policy No and Cover Note No not Allowed"
          TxtPolicyNo.SetFocus
          RsMst_PolicyDetail.Filter = ""
          Exit Sub
       End If
   End If
End If

RsMst_PolicyDetail!InsuranceCoID = mInsuranceCoID
RsMst_PolicyDetail!InsuranceRiskCoveredID = mInsuranceRiskTypeID
RsMst_PolicyDetail!InsuranceByID = mInsuranceByID
RsMst_PolicyDetail!Risks = CmbInsuranceRiskTypeID.Text
RsMst_PolicyDetail!Amount = Format(Val(TxtAmount), 0#)
RsMst_PolicyDetail!StructureWiseLimit = IIf(TxtStructureWiseLimit.Text = "", Null, Format(Val(TxtStructureWiseLimit.Text), 0#))
RsMst_PolicyDetail!ValidFrom = ValidFrom.Value
RsMst_PolicyDetail!ValidTo = ValidTo.Value
RsMst_PolicyDetail!Flag = UCase(Left(CmbStatus, 1))
RsMst_PolicyDetail!Remark = TxtRemark

If IsNull(RsMst_PolicyDetail!CreatedBy) = True Or RsMst_PolicyDetail!CreatedBy = "" Then
   RsMst_PolicyDetail!CreatedBy = Gen_UserLoginID
   RsMst_PolicyDetail!CreatedDate = Now
Else
   RsMst_PolicyDetail!UpdatedBy = Gen_UserLoginID
   RsMst_PolicyDetail!UpdatedDate = Now
End If

RsMst_PolicyDetail.Update

Call SaveInsuranceUnitType      ''-- Godown Unit Type
Call SaveInsuranceValidFor      ''-- Location Details
Call SaveInsuranceGodownMapping ''-- Godown Mapping
Call SaveInsuranceCommodity     ''-- Commodity Details
Call SaveInsuranceClientDetail  ''-- Client Details

RsMst_PolicyDetail.Filter = ""
RsMst_PolicyDetail.MoveFirst
RsMst_PolicyDetail.Find "G_UID = '" & TxtGUID & "'"
If RsMst_PolicyDetail.EOF Then
   RsMst_PolicyDetail.MoveFirst
End If
Call Indata

MsgBox "Data saved Successfully."

End Sub

Private Sub SaveInsuranceCommodity()
Dim mR As Double, mRow As Double
Dim mCommodityList As Variant
mCommodityList = "0"

''--- Include
With MSHFlexGrid3
    If .Rows > 2 Then
        
        Call TableMst_InsuranceCommodityDetail(" Where InsurancePolicyID = " & Val(TxtInsurancePolicyID) & " And EntryType = 'I' ")
        For mR = 1 To .Rows - 2
        
            If RsMst_InsuranceCommodityDetail.RecordCount > 0 Then RsMst_InsuranceCommodityDetail.MoveFirst
            RsMst_InsuranceCommodityDetail.Find " InsuranceCommodityDetailID = " & Val(.TextMatrix(mR, 3))
            If RsMst_InsuranceCommodityDetail.EOF = True Then
                RsMst_InsuranceCommodityDetail.AddNew
                RsMst_InsuranceCommodityDetail!InsuranceCommodityDetailID = GetMaxID("InsuranceCommodityDetailID", "Mst_InsuranceCommodityDetail")
                .TextMatrix(mR, 3) = RsMst_InsuranceCommodityDetail!InsuranceCommodityDetailID
                RsMst_InsuranceCommodityDetail!G_UID = GetGUID
                RsMst_InsuranceCommodityDetail!InsurancePolicyID = TxtInsurancePolicyID
                RsMst_InsuranceCommodityDetail!CommodityID = .TextMatrix(mR, 2)
                RsMst_InsuranceCommodityDetail!EntryType = "I"
            End If
            mCommodityList = mCommodityList & "," & .TextMatrix(mR, 3)
            If IsNull(RsMst_InsuranceCommodityDetail!CreatedBy) = True Or RsMst_InsuranceCommodityDetail!CreatedBy = "" Then
               RsMst_InsuranceCommodityDetail!CreatedBy = Gen_UserLoginID
               RsMst_InsuranceCommodityDetail!CreatedDate = Now
            Else
               RsMst_InsuranceCommodityDetail!UpdatedBy = Gen_UserLoginID
               RsMst_InsuranceCommodityDetail!UpdatedDate = Now
            End If
            RsMst_InsuranceCommodityDetail.Update
            RsMst_InsuranceCommodityDetail.Requery
        
        Next mR
    End If
    tmp = " Delete From Mst_InsuranceCommodityDetail Where InsurancePolicyID = " & Val(TxtInsurancePolicyID) & " And InsuranceCommodityDetailID Not In (" & mCommodityList & ") And EntryType = 'I' "
    Call TmpTable
End With

mCommodityList = 0
''--- Exclude
With MSHFlexGrid4
    If .Rows > 2 Then
        
        Call TableMst_InsuranceCommodityDetail(" Where InsurancePolicyID = " & Val(TxtInsurancePolicyID) & " And EntryType = 'E' ")
        For mR = 1 To .Rows - 2
        
            If RsMst_InsuranceCommodityDetail.RecordCount > 0 Then RsMst_InsuranceCommodityDetail.MoveFirst
            RsMst_InsuranceCommodityDetail.Find " InsuranceCommodityDetailID = " & Val(.TextMatrix(mR, 3))
            If RsMst_InsuranceCommodityDetail.EOF = True Then
                RsMst_InsuranceCommodityDetail.AddNew
                RsMst_InsuranceCommodityDetail!InsuranceCommodityDetailID = GetMaxID("InsuranceCommodityDetailID", "Mst_InsuranceCommodityDetail")
                .TextMatrix(mR, 3) = RsMst_InsuranceCommodityDetail!InsuranceCommodityDetailID
                RsMst_InsuranceCommodityDetail!G_UID = GetGUID
                RsMst_InsuranceCommodityDetail!InsurancePolicyID = TxtInsurancePolicyID
                RsMst_InsuranceCommodityDetail!CommodityID = .TextMatrix(mR, 2)
                RsMst_InsuranceCommodityDetail!EntryType = "E"
            End If
            mCommodityList = mCommodityList & "," & .TextMatrix(mR, 3)
            If IsNull(RsMst_InsuranceCommodityDetail!CreatedBy) = True Or RsMst_InsuranceCommodityDetail!CreatedBy = "" Then
               RsMst_InsuranceCommodityDetail!CreatedBy = Gen_UserLoginID
               RsMst_InsuranceCommodityDetail!CreatedDate = Now
            Else
               RsMst_InsuranceCommodityDetail!UpdatedBy = Gen_UserLoginID
               RsMst_InsuranceCommodityDetail!UpdatedDate = Now
            End If
            RsMst_InsuranceCommodityDetail.Update
            RsMst_InsuranceCommodityDetail.Requery
        
        Next mR
    End If
    tmp = " Delete From Mst_InsuranceCommodityDetail Where InsurancePolicyID = " & Val(TxtInsurancePolicyID) & " And InsuranceCommodityDetailID Not In (" & mCommodityList & ") And EntryType = 'E'  "
    Call TmpTable
End With

End Sub

Private Sub SaveInsuranceClientDetail()
Dim mR As Double, mRow As Double
Dim mClientList As Variant
mClientList = "0"
With MSHFlexGrid5

    If mInsuranceByID = 2 Then
        If .Rows > 2 Then
            Call TableMst_InsuranceClientDetail(" Where InsurancePolicyID = " & Val(TxtInsurancePolicyID) & " ")
            For mR = 1 To .Rows - 2
                 
                If RsMst_InsuranceClientDetail.RecordCount > 0 Then RsMst_InsuranceClientDetail.MoveFirst
                RsMst_InsuranceClientDetail.Find " InsuranceClientDetailID = " & Val(.TextMatrix(mR, 4))
                If RsMst_InsuranceClientDetail.EOF = True Then
                    RsMst_InsuranceClientDetail.AddNew
                    RsMst_InsuranceClientDetail!InsuranceClientDetailID = GetMaxID("InsuranceClientDetailID", "Mst_InsuranceClientDetail")
                    .TextMatrix(mR, 4) = RsMst_InsuranceClientDetail!InsuranceClientDetailID
                    RsMst_InsuranceClientDetail!G_UID = GetGUID
                    RsMst_InsuranceClientDetail!InsurancePolicyID = TxtInsurancePolicyID
                End If
                mClientList = mClientList & "," & .TextMatrix(mR, 4)
                RsMst_InsuranceClientDetail!ClientIDRow = .TextMatrix(mR, 3)
                If IsNull(RsMst_InsuranceClientDetail!CreatedBy) = True Or RsMst_InsuranceClientDetail!CreatedBy = "" Then
                   RsMst_InsuranceClientDetail!CreatedBy = Gen_UserLoginID
                   RsMst_InsuranceClientDetail!CreatedDate = Now
                Else
                   RsMst_InsuranceClientDetail!UpdatedBy = Gen_UserLoginID
                   RsMst_InsuranceClientDetail!UpdatedDate = Now
                End If
                RsMst_InsuranceClientDetail.Update
                RsMst_InsuranceClientDetail.Requery
                
            Next mR
        End If
    End If
    
    tmp = " Delete From Mst_InsuranceClientDetail Where InsurancePolicyID = " & Val(TxtInsurancePolicyID) & " And InsuranceClientDetailID Not In (" & mClientList & ") "
    Call TmpTable
End With
End Sub

Private Sub SaveInsuranceValidFor()
Dim mR As Double, mRow As Double
Dim mValidList As Variant
mValidList = "0"

With MSHFlexGrid7
    If .Rows > 2 Then
        Call TableMst_InsuranceValidFor(" Where InsurancePolicyID = " & Val(TxtInsurancePolicyID) & " ")
        For mR = 1 To .Rows - 2
             
            If RsMst_InsuranceValidFor.RecordCount > 0 Then RsMst_InsuranceValidFor.MoveFirst
            RsMst_InsuranceValidFor.Find " InsuranceValidForID = " & Val(.TextMatrix(mR, 4))
            If RsMst_InsuranceValidFor.EOF = True Then
                RsMst_InsuranceValidFor.AddNew
                RsMst_InsuranceValidFor!InsuranceValidForID = GetMaxID("InsuranceValidForID", "Mst_InsuranceValidFor")
                .TextMatrix(mR, 4) = RsMst_InsuranceValidFor!InsuranceValidForID
                RsMst_InsuranceValidFor!G_UID = GetGUID
                RsMst_InsuranceValidFor!InsurancePolicyID = TxtInsurancePolicyID
            End If
            mValidList = mValidList & "," & .TextMatrix(mR, 4)
            RsMst_InsuranceValidFor!ValidForType = UCase(Left(.TextMatrix(mR, 1), 1))
            RsMst_InsuranceValidFor!ValidForTypeID = .TextMatrix(mR, 3)
            If IsNull(RsMst_InsuranceValidFor!CreatedBy) = True Or RsMst_InsuranceValidFor!CreatedBy = "" Then
               RsMst_InsuranceValidFor!CreatedBy = Gen_UserLoginID
               RsMst_InsuranceValidFor!CreatedDate = Now
            Else
               RsMst_InsuranceValidFor!UpdatedBy = Gen_UserLoginID
               RsMst_InsuranceValidFor!UpdatedDate = Now
            End If
            RsMst_InsuranceValidFor.Update
            RsMst_InsuranceValidFor.Requery
            
        Next mR
    End If
    
    tmp = " Delete From Mst_InsuranceValidFor Where InsurancePolicyID = " & Val(TxtInsurancePolicyID) & " And InsuranceValidForID Not In (" & mValidList & ") "
    Call TmpTable
End With

End Sub

Private Sub SaveInsuranceUnitType()
Dim mR As Double, mRow As Double
Dim mUnitList As Variant
mUnitList = "0"
With MSHFlexGrid6
    If .Rows > 2 Then
        Call TableMst_InsuranceForUnitTypeDetail(" Where InsurancePolicyID = " & Val(TxtInsurancePolicyID) & " ")
        For mR = 1 To .Rows - 2
                
            If .TextMatrix(mR, 0) = strChecked Then
                
                If RsMst_InsuranceForUnitTypeDetail.RecordCount > 0 Then RsMst_InsuranceForUnitTypeDetail.MoveFirst
                RsMst_InsuranceForUnitTypeDetail.Find " InsuranceForUnitTypeDetailID = " & Val(.TextMatrix(mR, 3))
                If RsMst_InsuranceForUnitTypeDetail.EOF = True Then
                    RsMst_InsuranceForUnitTypeDetail.AddNew
                    RsMst_InsuranceForUnitTypeDetail!InsuranceForUnitTypeDetailID = GetMaxID("InsuranceForUnitTypeDetailID", "Mst_InsuranceForUnitTypeDetail")
                    .TextMatrix(mR, 3) = RsMst_InsuranceForUnitTypeDetail!InsuranceForUnitTypeDetailID
                    RsMst_InsuranceForUnitTypeDetail!G_UID = GetGUID
                    RsMst_InsuranceForUnitTypeDetail!InsurancePolicyID = TxtInsurancePolicyID
                    RsMst_InsuranceForUnitTypeDetail!UnitTypeID = .TextMatrix(mR, 2)
                End If
                mUnitList = mUnitList & "," & .TextMatrix(mR, 3)
                If IsNull(RsMst_InsuranceForUnitTypeDetail!CreatedBy) = True Or RsMst_InsuranceForUnitTypeDetail!CreatedBy = "" Then
                   RsMst_InsuranceForUnitTypeDetail!CreatedBy = Gen_UserLoginID
                   RsMst_InsuranceForUnitTypeDetail!CreatedDate = Now
                Else
                   RsMst_InsuranceForUnitTypeDetail!UpdatedBy = Gen_UserLoginID
                   RsMst_InsuranceForUnitTypeDetail!UpdatedDate = Now
                End If
                RsMst_InsuranceForUnitTypeDetail.Update
                RsMst_InsuranceForUnitTypeDetail.Requery
            
            End If
            
        Next mR
    End If
    
    tmp = " Delete From Mst_InsuranceForUnitTypeDetail Where InsurancePolicyID = " & Val(TxtInsurancePolicyID) & " And InsuranceForUnitTypeDetailID Not In (" & mUnitList & ") "
    Call TmpTable
    
End With
End Sub

Private Sub SaveInsuranceGodownMapping()
Dim mR As Double, mRow As Double
Dim mGList As Variant, mGListDetail As Variant

''mGList = "0"
''With MSHFlexGrid2
''    For mR = 1 To .Rows - 2
''        If .TextMatrix(mR, 0) = strChecked Then
''            mGList = mGList & "," & Val(.TextMatrix(mR, 9))
''        End If
''    Next mR
''End With
''
''tmp = " Delete From Mst_InsuranceGodownMapping Where InsurancePolicyID = " & Val(TxtInsurancePolicyID) & " And GodownID Not In (" & mGList & ") "
''Call TmpTable

mGListDetail = "0"
With MSHFlexGrid2
    If .Rows > 2 Then
        Call TableMst_InsuranceGodownMapping(" Where InsurancePolicyID = " & Val(TxtInsurancePolicyID) & " ")
        For mR = 1 To .Rows - 2
                
            If .TextMatrix(mR, 0) = strChecked Then
            
                If RsMst_InsuranceGodownMapping.RecordCount > 0 Then RsMst_InsuranceGodownMapping.MoveFirst
                RsMst_InsuranceGodownMapping.Find " InsuranceGodownMappingID = " & Val(.TextMatrix(mR, 12))
                If RsMst_InsuranceGodownMapping.EOF = True Then
                    RsMst_InsuranceGodownMapping.AddNew
                    RsMst_InsuranceGodownMapping!InsuranceGodownMappingID = GetMaxID("InsuranceGodownMappingID", "Mst_InsuranceGodownMapping")
                    .TextMatrix(mR, 12) = RsMst_InsuranceGodownMapping!InsuranceGodownMappingID
                    RsMst_InsuranceGodownMapping!G_UID = GetGUID
                    RsMst_InsuranceGodownMapping!InsurancePolicyID = TxtInsurancePolicyID
                    RsMst_InsuranceGodownMapping!GodownID = .TextMatrix(mR, 10)
                End If
                mGListDetail = mGListDetail & "," & .TextMatrix(mR, 12)
                RsMst_InsuranceGodownMapping!Connected = IIf(.TextMatrix(mR, 9) = strChecked, "Y", "N")
                
                If IsNull(RsMst_InsuranceGodownMapping!CreatedBy) = True Or RsMst_InsuranceGodownMapping!CreatedBy = "" Then
                   RsMst_InsuranceGodownMapping!CreatedBy = Gen_UserLoginID
                   RsMst_InsuranceGodownMapping!CreatedDate = Now
                Else
                   RsMst_InsuranceGodownMapping!UpdatedBy = Gen_UserLoginID
                   RsMst_InsuranceGodownMapping!UpdatedDate = Now
                End If
                RsMst_InsuranceGodownMapping.Update
                RsMst_InsuranceGodownMapping.Requery
            
            End If
                
        Next mR
    End If
    
    'tmp = " Delete From Mst_InsuranceGodownMapping Where InsurancePolicyID = " & Val(TxtInsurancePolicyID) & " And InsuranceGodownMappingID Not In (" & mGListDetail & ") "
    'Call TmpTable
    
End With
End Sub


Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   Call GButtonLock(Me, True)
   If RsMst_PolicyDetail.EOF Then
      RsMst_PolicyDetail.MoveFirst
   End If
   Call Indata
   Exit Sub
End If

Frame0.Left = Frame1.Left
Frame0.Top = Frame1.Top
Frame0.Height = Frame1.Height + ButtonFrm.Height + 100
Frame0.Width = Frame1.Width

MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 1800
Frame0.Visible = True
RsMst_PolicyDetail.MoveFirst

End Sub

Private Sub TxtAmount_LostFocus()
If TxtAmount <> "" Then
    TxtAmount = Replace(TxtAmount, ",", "")
    If Not IsNumeric(TxtAmount) Then
       MsgBox " Amount should be Numeric !!!", vbExclamation, "Message"
       TxtAmount.SetFocus
       Exit Sub
    End If
    If Val(TxtAmount) < 0 Then
       MsgBox " Amount should not be -Ve !!!", vbExclamation, "Message"
       TxtAmount.SetFocus
       Exit Sub
    End If
End If
End Sub

Private Sub TxtClientID_LostFocus()
''If CmbExchange.Text = "" Then
''    MsgBox "Select Exchange."
''    CmbExchange.SetFocus
''    Exit Sub
''End If

If TxtClientID <> "" Then
''If IsNumeric(TxtClientID) = False Then
''    MsgBox "Invalid input."
''    Exit Sub
''End If
    Call TableMst_Client(" Where ExchangeTypeID = " & mExchangeTypeID & " And Status = 'A' And ClientID = '" & TxtClientID & "' ")
    CmbClientRowID.Clear
    If RsMst_Client.RecordCount = 0 Then
       MsgBox "No Record found !!!! "
       Exit Sub
    End If
    If RsMst_Client.RecordCount > 0 Then
        For I = 1 To RsMst_Client.RecordCount
            CmbClientRowID.AddItem RsMst_Client!ClientName & "#" & RsMst_Client!ClientID & "#" & RsMst_Client!ClientIDRow
            RsMst_Client.MoveNext
        Next
        RsMst_Client.MoveFirst
        If RsMst_Client.EOF = False Then
            CmbClientRowID.Text = RsMst_Client!ClientName & "#" & RsMst_Client!ClientID & "#" & RsMst_Client!ClientIDRow
            CmbClientRowID.SetFocus
        End If
    End If
End If
End Sub

Private Sub TxtCoverNoteNo_LostFocus()
Dim ans As Boolean
ans = IsClean(TxtCoverNoteNo)
If ans = False Then
   MsgBox "Special character not Allowed  !!!!!!!! "
   TxtCoverNoteNo.SetFocus
End If
End Sub

Private Sub TxtPolicyNo_LostFocus()
Dim ans As Boolean
ans = IsClean(TxtPolicyNo)
If ans = False Then
   MsgBox "Special character not Allowed  !!!!!!!! "
   TxtPolicyNo.SetFocus
End If
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Mst_Insurance.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Mst_Insurance.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid1.Rows - 2
    mRow = I + 1 ' + 6
    For C = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(I, C), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(I, C)) Then
              oWS.Cells(mRow, C + 1) = Format(MSHFlexGrid1.TextMatrix(I, C), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
           End If
        Else
           If I > 0 Then
                If C = 1 Or C = 2 Then
                    oWS.Cells(mRow, C + 1) = "`" & MSHFlexGrid1.TextMatrix(I, C)
                Else
                    oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
                End If
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

Private Sub Grid_Head()
With MSHFlexGrid1
    .Clear
    .Rows = 2
    .Cols = 15
    .FixedCols = 0
    .Font.Size = 10
    .WordWrap = True
    .TextMatrix(0, 0) = "Insurance Policy ID":          .ColWidth(0) = 1000
    .TextMatrix(0, 1) = "Policy No":                    .ColWidth(1) = 3000
    .TextMatrix(0, 2) = "Cover Note No":                .ColWidth(2) = 3000
    .TextMatrix(0, 3) = "Insurance Company":            .ColWidth(3) = 2500
    .TextMatrix(0, 4) = "Insurance Risk Covered ":      .ColWidth(4) = 2500
    .TextMatrix(0, 5) = "Insurance By":                 .ColWidth(5) = 1500
    .TextMatrix(0, 6) = "Status":                       .ColWidth(6) = 1000
    .TextMatrix(0, 7) = "Amount":                       .ColWidth(7) = 1500
    .TextMatrix(0, 8) = "Structure Wise Limit":         .ColWidth(8) = 1500
    .TextMatrix(0, 9) = "Valid From":                   .ColWidth(9) = 1500
    .TextMatrix(0, 10) = "Valid To":                    .ColWidth(10) = 1500
    ''-- Ullhas Patil 04 April 2018
    .TextMatrix(0, 11) = "Connected Godown":            .ColWidth(11) = 1500
    .TextMatrix(0, 12) = "Commodity Mapped":            .ColWidth(12) = 1500
    .TextMatrix(0, 13) = "Client Mapped":               .ColWidth(13) = 1500
    .TextMatrix(0, 14) = "Remark":                      .ColWidth(14) = 1500
    .RowHeight(0) = 800
End With
End Sub

Private Sub Grid_Head8()
With MSHFlexGrid8
    .Clear
    .Rows = 2
    .Cols = 12
    .FixedCols = 0
    .Font.Size = 10
    .TextMatrix(0, 0) = "Select":           .ColWidth(0) = 700
    .TextMatrix(0, 1) = "Godown No":        .ColWidth(1) = 1200
    .TextMatrix(0, 2) = "Unit Type":        .ColWidth(2) = 1800
    .TextMatrix(0, 3) = "CMP Name":         .ColWidth(3) = 2000
    .TextMatrix(0, 4) = "Location":         .ColWidth(4) = 2000
    .TextMatrix(0, 5) = "State":            .ColWidth(5) = 2000
    .TextMatrix(0, 6) = "Region":           .ColWidth(6) = 2000
    .TextMatrix(0, 7) = "Address":          .ColWidth(7) = 3500
    .TextMatrix(0, 8) = "Capacity":         .ColWidth(8) = 1000
    .TextMatrix(0, 9) = "Connected (Yes / No)": .ColWidth(9) = 1200
    .TextMatrix(0, 10) = "GodownID":        .ColWidth(10) = 0
    .TextMatrix(0, 11) = "UnitID":          .ColWidth(11) = 0
End With
End Sub

Private Sub Grid_Head2()    ''-- Frame2
With MSHFlexGrid2
    .Clear
    .Rows = 2
    .Cols = 13
    .FixedCols = 0
    .Font.Size = 10
    .WordWrap = True
    .RowHeight(0) = 600
    
    .TextMatrix(0, 0) = "Select":                           .ColWidth(0) = 700
    .TextMatrix(0, 1) = "Godown No":                        .ColWidth(1) = 1200
    .TextMatrix(0, 2) = "Unit Type":                        .ColWidth(2) = 1500
    .TextMatrix(0, 3) = "CMP Name":                         .ColWidth(3) = 1800
    .TextMatrix(0, 4) = "Location":                         .ColWidth(4) = 1800
    .TextMatrix(0, 5) = "State":                            .ColWidth(5) = 2000
    .TextMatrix(0, 6) = "Region":                           .ColWidth(6) = 2000
    .TextMatrix(0, 7) = "Address":                          .ColWidth(7) = 2000
    .TextMatrix(0, 8) = "Capacity":                         .ColWidth(8) = 1000
    .TextMatrix(0, 9) = "Connected (Yes / No)":             .ColWidth(9) = 1200
    .TextMatrix(0, 10) = "GodownID":                        .ColWidth(10) = 0
    .TextMatrix(0, 11) = "UnitID":                          .ColWidth(11) = 0
    .TextMatrix(0, 12) = "RowID":                           .ColWidth(12) = 0

End With
End Sub

Private Sub Grid_Head7()    ''-- Frame7
With MSHFlexGrid7
    .Clear
    .Rows = 2
    .Cols = 5
    .Font.Size = 10
    .TextMatrix(0, 0) = "Sr No":        .ColWidth(0) = 800
    .TextMatrix(0, 1) = "Type":         .ColWidth(1) = 2500
    .TextMatrix(0, 2) = "Name":         .ColWidth(2) = 3500
    .TextMatrix(0, 3) = "ID":           .ColWidth(3) = 0
    .TextMatrix(0, 4) = "DetailID":     .ColWidth(4) = 0
End With
End Sub


Private Sub Grid_Head3()    ''-- Frame3
With MSHFlexGrid3
    .Clear
    .Rows = 2
    .Cols = 4
    .Font.Size = 10
    .TextMatrix(0, 0) = "Sr No":        .ColWidth(0) = 800
    .TextMatrix(0, 1) = "Commodity":    .ColWidth(1) = 6000
    .TextMatrix(0, 2) = "CommodityID":  .ColWidth(2) = 0
    .TextMatrix(0, 3) = "DetailID":     .ColWidth(3) = 0
End With
End Sub
Private Sub Grid_Head4()    ''-- Frame4
With MSHFlexGrid4
    .Clear
    .Rows = 2
    .Cols = 4
    .Font.Size = 10
    .TextMatrix(0, 0) = "Sr No":        .ColWidth(0) = 800
    .TextMatrix(0, 1) = "Commodity":    .ColWidth(1) = 6000
    .TextMatrix(0, 2) = "CommodityID":  .ColWidth(2) = 0
    .TextMatrix(0, 3) = "DetailID":     .ColWidth(3) = 0
End With
End Sub
Private Sub Grid_Head5()    ''-- Frame5
With MSHFlexGrid5
    .Clear
    .Rows = 2
    .Cols = 5 '7
    .Font.Size = 10
    .TextMatrix(0, 0) = "Sr No":            .ColWidth(0) = 800
    .TextMatrix(0, 1) = "Client ID":        .ColWidth(1) = 3000
    .TextMatrix(0, 2) = "Client Name":      .ColWidth(2) = 8000
    .TextMatrix(0, 3) = "ClientRowID":      .ColWidth(3) = 0
    .TextMatrix(0, 4) = "DetailID":         .ColWidth(4) = 0
    '.TextMatrix(0, 5) = "Exchange":         .ColWidth(5) = 1500
    '.TextMatrix(0, 6) = "ExchangeTypeID":   .ColWidth(6) = 0
End With
End Sub
Private Sub Grid_Head6()    ''-- Frame6
With MSHFlexGrid6
    .Clear
    .Rows = 2
    .Cols = 4
    .FixedCols = 0
    .Font.Size = 10
    .TextMatrix(0, 0) = "Select":       .ColWidth(0) = 800
    .TextMatrix(0, 1) = "Godown Type":  .ColWidth(1) = 4000
    .TextMatrix(0, 2) = "UnitTypeID":   .ColWidth(2) = 0
    .TextMatrix(0, 3) = "DetailID":     .ColWidth(3) = 0
End With
End Sub

Private Sub TxtStructureWiseLimit_LostFocus()
If TxtStructureWiseLimit <> "" Then
    TxtStructureWiseLimit = Replace(TxtStructureWiseLimit, ",", "")
    If Not IsNumeric(TxtStructureWiseLimit) Then
       MsgBox " Structure Wise Limit should be Numeric !!!", vbExclamation, "Message"
       TxtStructureWiseLimit.SetFocus
       Exit Sub
    End If
    If Val(TxtStructureWiseLimit) < 0 Then
       MsgBox " Structure Wise Limit should not be -Ve !!!", vbExclamation, "Message"
       TxtStructureWiseLimit.SetFocus
       Exit Sub
    End If
End If
End Sub

Public Function GRowColor(Grid As MSHFlexGrid, RowID As Double, RowColor As Variant)
Dim mC As Double
With Grid
    For mC = 0 To .Cols - 1
        .row = RowID
        .Col = mC
        .CellBackColor = RowColor
    Next mC
End With
End Function


