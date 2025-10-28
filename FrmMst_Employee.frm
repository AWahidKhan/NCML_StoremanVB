VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmMst_Employee 
   Caption         =   "Employee Master"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   8880
   WindowState     =   2  'Maximized
   Begin VB.VScrollBar VScroll1 
      Height          =   9015
      Left            =   13440
      TabIndex        =   99
      Top             =   0
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      TabIndex        =   98
      Top             =   10320
      Width           =   10935
   End
   Begin VB.Frame Frame4 
      Height          =   9975
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   13815
      Begin VB.Frame Frame0 
         Height          =   1575
         Left            =   3720
         TabIndex        =   91
         Top             =   9480
         Visible         =   0   'False
         Width           =   15480
         Begin VB.CommandButton CmdExcel 
            Caption         =   "Excel"
            Height          =   350
            Left            =   11640
            TabIndex        =   92
            Top             =   720
            Width           =   1245
         End
         Begin VB.OptionButton Option5 
            Caption         =   "Employee Name"
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
            Left            =   3630
            TabIndex        =   34
            Top             =   240
            Width           =   1830
         End
         Begin VB.TextBox TxtSearchEmployee 
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
            TabIndex        =   36
            Top             =   720
            Width           =   4815
         End
         Begin VB.OptionButton Option2 
            Caption         =   "StoreMan Group"
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
            Left            =   5760
            TabIndex        =   35
            Top             =   240
            Width           =   1905
         End
         Begin VB.OptionButton Option1 
            Caption         =   "HR Employee ID"
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
            TabIndex        =   33
            Top             =   240
            Width           =   2190
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
            TabIndex        =   32
            Top             =   240
            Value           =   -1  'True
            Width           =   585
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
            ItemData        =   "FrmMst_Employee.frx":0000
            Left            =   8760
            List            =   "FrmMst_Employee.frx":000A
            Sorted          =   -1  'True
            TabIndex        =   37
            Top             =   720
            Width           =   1500
         End
         Begin VB.CommandButton CmdGo 
            Caption         =   "Go"
            Height          =   345
            Left            =   10320
            TabIndex        =   38
            Top             =   720
            Width           =   1245
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   2220
            Left            =   120
            TabIndex        =   93
            Top             =   1140
            Width           =   6135
            _ExtentX        =   10821
            _ExtentY        =   3916
            _Version        =   393216
            Rows            =   3
            AllowUserResizing=   1
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin VB.Label Label19 
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
            TabIndex        =   97
            Top             =   705
            Width           =   2895
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
            Left            =   11640
            TabIndex        =   96
            Top             =   240
            Width           =   75
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
            Left            =   9120
            TabIndex        =   95
            Top             =   240
            Width           =   2415
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
            TabIndex        =   94
            Top             =   720
            Width           =   1335
         End
      End
      Begin VB.Frame ButtonFrm 
         Height          =   1020
         Left            =   120
         TabIndex        =   83
         Top             =   8685
         Width           =   12720
         Begin VB.CommandButton ExitCmd 
            Caption         =   "E&xit"
            Height          =   375
            Left            =   9945
            TabIndex        =   29
            Top             =   540
            Width           =   2415
         End
         Begin VB.CommandButton LastCmd 
            Caption         =   "&Last"
            Height          =   375
            Left            =   7530
            TabIndex        =   90
            Top             =   540
            Width           =   2415
         End
         Begin VB.CommandButton FirstCmd 
            Caption         =   "&First"
            Height          =   375
            Left            =   4995
            TabIndex        =   89
            Top             =   540
            Width           =   2535
         End
         Begin VB.CommandButton PreviousCmd 
            Caption         =   "&Previous"
            Height          =   375
            Left            =   2580
            TabIndex        =   88
            Top             =   540
            Width           =   2415
         End
         Begin VB.CommandButton NextCmd 
            Caption         =   "&Next"
            Height          =   375
            Left            =   165
            TabIndex        =   87
            Top             =   540
            Width           =   2415
         End
         Begin VB.CommandButton SearchCmd 
            Caption         =   "&Search"
            Height          =   375
            Left            =   9945
            TabIndex        =   30
            Top             =   165
            Width           =   2415
         End
         Begin VB.CommandButton DeleteCmd 
            Caption         =   "&Delete"
            Height          =   375
            Left            =   7530
            TabIndex        =   86
            Top             =   165
            Width           =   2415
         End
         Begin VB.CommandButton EditCmd 
            Caption         =   "&Edit"
            Height          =   375
            Left            =   4995
            TabIndex        =   85
            Top             =   165
            Width           =   2535
         End
         Begin VB.CommandButton SaveCmd 
            Caption         =   "Save"
            Height          =   375
            Left            =   2580
            TabIndex        =   28
            Top             =   165
            Width           =   2415
         End
         Begin VB.CommandButton AddCmd 
            Caption         =   "&Add New "
            Height          =   375
            Left            =   165
            TabIndex        =   84
            Top             =   165
            Width           =   2415
         End
      End
      Begin VB.Frame Frame1 
         Height          =   8580
         Left            =   120
         TabIndex        =   31
         Top             =   120
         Width           =   12720
         Begin VB.TextBox TxtHRDirectReportee 
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
            Left            =   2280
            Locked          =   -1  'True
            TabIndex        =   115
            TabStop         =   0   'False
            Top             =   2415
            Width           =   10230
         End
         Begin VB.TextBox TxtHRDesignation 
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
            Left            =   2265
            Locked          =   -1  'True
            TabIndex        =   101
            TabStop         =   0   'False
            Top             =   1530
            Width           =   10230
         End
         Begin VB.TextBox TxtEmpEmailID 
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
            Left            =   2280
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   2865
            Width           =   4485
         End
         Begin VB.ComboBox CmbConnectedRAG 
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
            Left            =   12960
            Sorted          =   -1  'True
            TabIndex        =   6
            Top             =   1905
            Visible         =   0   'False
            Width           =   2820
         End
         Begin VB.TextBox TxtEmpID 
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
            Left            =   2280
            Locked          =   -1  'True
            TabIndex        =   65
            TabStop         =   0   'False
            Top             =   180
            Width           =   2925
         End
         Begin VB.TextBox TxtEmpNo 
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
            Left            =   2280
            Locked          =   -1  'True
            TabIndex        =   1
            Top             =   627
            Width           =   2940
         End
         Begin VB.ComboBox CmbDesignationID 
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
            Left            =   2280
            TabIndex        =   3
            Top             =   1074
            Width           =   2940
         End
         Begin VB.TextBox TxtEmpName 
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
            Left            =   7080
            Locked          =   -1  'True
            TabIndex        =   2
            Top             =   627
            Width           =   5385
         End
         Begin VB.TextBox TxtAdditionalReporting 
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
            Left            =   10680
            Locked          =   -1  'True
            TabIndex        =   62
            TabStop         =   0   'False
            Top             =   120
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.CommandButton CmdDesignationID 
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
            Left            =   5280
            TabIndex        =   61
            Top             =   1095
            Width           =   375
         End
         Begin VB.ComboBox CmbDirectReporting 
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
            Left            =   2280
            Sorted          =   -1  'True
            TabIndex        =   5
            Top             =   1980
            Width           =   10230
         End
         Begin VB.ComboBox CmbAdditionalReporting 
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
            Left            =   10680
            Sorted          =   -1  'True
            TabIndex        =   60
            Top             =   135
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.ComboBox CmbFlag 
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
            ItemData        =   "FrmMst_Employee.frx":0021
            Left            =   7065
            List            =   "FrmMst_Employee.frx":002B
            Sorted          =   -1  'True
            TabIndex        =   4
            Top             =   1074
            Width           =   2625
         End
         Begin VB.TextBox TxtMobileNo 
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
            Left            =   9480
            Locked          =   -1  'True
            TabIndex        =   8
            Top             =   2865
            Width           =   3015
         End
         Begin VB.TextBox TxtAddress 
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
            Height          =   570
            Left            =   2280
            Locked          =   -1  'True
            MaxLength       =   250
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   9
            Top             =   3300
            Width           =   10215
         End
         Begin VB.CheckBox ChkSelfReporting 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Check1"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   12210
            TabIndex        =   40
            Top             =   1140
            Visible         =   0   'False
            Width           =   225
         End
         Begin VB.CheckBox ChkCMPAllocation 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   11985
            TabIndex        =   12
            Top             =   4035
            Visible         =   0   'False
            Width           =   225
         End
         Begin VB.ComboBox CmbBaseLocationID 
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
            Left            =   2280
            Sorted          =   -1  'True
            TabIndex        =   10
            Top             =   3960
            Width           =   7380
         End
         Begin VB.CommandButton CmdBaseLocationID 
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
            Left            =   9720
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   3960
            Width           =   375
         End
         Begin TabDlg.SSTab SSTab1 
            Height          =   3885
            Left            =   75
            TabIndex        =   41
            Top             =   4500
            Width           =   12420
            _ExtentX        =   21908
            _ExtentY        =   6853
            _Version        =   393216
            TabHeight       =   520
            TabCaption(0)   =   "CMP Details"
            TabPicture(0)   =   "FrmMst_Employee.frx":0042
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Frame2"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Allocation Details"
            TabPicture(1)   =   "FrmMst_Employee.frx":005E
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Frame5"
            Tab(1).Control(1)=   "MSHAllocation"
            Tab(1).Control(2)=   "Frame3"
            Tab(1).ControlCount=   3
            TabCaption(2)   =   "Log History"
            TabPicture(2)   =   "FrmMst_Employee.frx":007A
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "TxtCreated"
            Tab(2).Control(0).Enabled=   0   'False
            Tab(2).Control(1)=   "TxtUpdated"
            Tab(2).Control(1).Enabled=   0   'False
            Tab(2).Control(2)=   "Label10"
            Tab(2).Control(3)=   "Label9"
            Tab(2).ControlCount=   4
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
               Left            =   -73335
               Locked          =   -1  'True
               TabIndex        =   113
               TabStop         =   0   'False
               Top             =   480
               Width           =   4095
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
               Left            =   -73335
               Locked          =   -1  'True
               TabIndex        =   111
               TabStop         =   0   'False
               Top             =   960
               Width           =   4095
            End
            Begin VB.Frame Frame5 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   975
               Left            =   -74880
               TabIndex        =   103
               Top             =   360
               Width           =   12225
               Begin VB.ComboBox CmbAllocationType 
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
                  ItemData        =   "FrmMst_Employee.frx":0096
                  Left            =   120
                  List            =   "FrmMst_Employee.frx":00A3
                  Sorted          =   -1  'True
                  TabIndex        =   107
                  Top             =   495
                  Width           =   3240
               End
               Begin VB.ComboBox CmbAllocationTypeID 
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
                  ItemData        =   "FrmMst_Employee.frx":00C0
                  Left            =   3480
                  List            =   "FrmMst_Employee.frx":00C2
                  Sorted          =   -1  'True
                  TabIndex        =   106
                  Top             =   495
                  Width           =   5160
               End
               Begin VB.CommandButton CmdAdd2GridAllocation 
                  Caption         =   "Add To Grid"
                  Height          =   375
                  Left            =   9000
                  TabIndex        =   105
                  Top             =   495
                  Width           =   1455
               End
               Begin VB.CommandButton CmdDeleteFromGridAllocation 
                  Caption         =   "Remove From Grid"
                  Height          =   375
                  Left            =   10470
                  TabIndex        =   104
                  Top             =   495
                  Width           =   1575
               End
               Begin VB.Label Label43 
                  AutoSize        =   -1  'True
                  Caption         =   "Allocation Type"
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
                  Left            =   1035
                  TabIndex        =   109
                  Top             =   210
                  Width           =   1410
               End
               Begin VB.Label lblstateregionlocation 
                  AutoSize        =   -1  'True
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
                  Left            =   5685
                  TabIndex        =   108
                  Top             =   210
                  Width           =   765
               End
            End
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
               Height          =   3495
               Left            =   120
               TabIndex        =   54
               Top             =   315
               Width           =   12225
               Begin VB.CommandButton CmdRemoveFromGrid 
                  Caption         =   "Remove From Grid"
                  Height          =   375
                  Left            =   10470
                  TabIndex        =   17
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
                  Left            =   5520
                  Sorted          =   -1  'True
                  TabIndex        =   14
                  Top             =   270
                  Width           =   6135
               End
               Begin VB.CommandButton CmdAddToGrid 
                  Caption         =   "Add To Grid"
                  Height          =   375
                  Left            =   9000
                  TabIndex        =   16
                  Top             =   645
                  Width           =   1455
               End
               Begin VB.CommandButton CmdCMPID 
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
                  TabIndex        =   13
                  Top             =   240
                  Width           =   3015
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
               Begin VB.CommandButton CmdAddAll2grid 
                  Caption         =   "Add All CMP to grid"
                  Height          =   375
                  Left            =   7200
                  TabIndex        =   15
                  Top             =   645
                  Width           =   1815
               End
               Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
                  Height          =   2340
                  Left            =   120
                  TabIndex        =   57
                  Top             =   1035
                  Width           =   12030
                  _ExtentX        =   21220
                  _ExtentY        =   4128
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
                  TabIndex        =   59
                  Top             =   270
                  Width           =   975
               End
               Begin VB.Label Label7 
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
                  Left            =   120
                  TabIndex        =   58
                  Top             =   300
                  Width           =   855
               End
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHAllocation 
               Height          =   2340
               Left            =   -74880
               TabIndex        =   110
               Top             =   1380
               Width           =   12225
               _ExtentX        =   21564
               _ExtentY        =   4128
               _Version        =   393216
               WordWrap        =   -1  'True
               _NumberOfBands  =   1
               _Band(0).Cols   =   2
            End
            Begin VB.Frame Frame3 
               Height          =   1560
               Left            =   -63120
               TabIndex        =   42
               Top             =   1800
               Visible         =   0   'False
               Width           =   7890
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
                  TabIndex        =   44
                  TabStop         =   0   'False
                  Top             =   270
                  Width           =   375
               End
               Begin VB.CommandButton CmdAdd2GridAsset 
                  Caption         =   "Add To Grid"
                  Height          =   375
                  Left            =   9105
                  TabIndex        =   26
                  Top             =   1200
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
                  Sorted          =   -1  'True
                  TabIndex        =   18
                  Top             =   270
                  Width           =   2535
               End
               Begin VB.CommandButton CmdDeleteFromGridAsset 
                  Caption         =   "Remove From Grid"
                  Height          =   375
                  Left            =   10575
                  TabIndex        =   27
                  Top             =   1200
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
                  Height          =   375
                  Left            =   3600
                  MaxLength       =   500
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   22
                  Top             =   720
                  Width           =   5415
               End
               Begin VB.CommandButton cmdAssetStatus 
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
                  Left            =   9120
                  TabIndex        =   43
                  TabStop         =   0   'False
                  Top             =   270
                  Width           =   375
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
                  Left            =   5700
                  Sorted          =   -1  'True
                  TabIndex        =   19
                  Top             =   270
                  Width           =   3345
               End
               Begin VB.TextBox TxtAssetQty 
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
                  Left            =   1320
                  TabIndex        =   21
                  Top             =   720
                  Width           =   975
               End
               Begin VB.TextBox TxtRemarks 
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
                  Left            =   3585
                  MaxLength       =   1000
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   25
                  Top             =   1200
                  Width           =   5400
               End
               Begin VB.TextBox TxtAMCAmount 
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
                  Left            =   1320
                  TabIndex        =   24
                  Top             =   1200
                  Width           =   975
               End
               Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
                  Height          =   1500
                  Left            =   150
                  TabIndex        =   45
                  Top             =   1590
                  Width           =   12030
                  _ExtentX        =   21220
                  _ExtentY        =   2646
                  _Version        =   393216
                  WordWrap        =   -1  'True
                  AllowUserResizing=   3
                  _NumberOfBands  =   1
                  _Band(0).Cols   =   2
               End
               Begin MSComCtl2.DTPicker AssetVerifiedon 
                  Height          =   330
                  Left            =   10680
                  TabIndex        =   20
                  Top             =   240
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   582
                  _Version        =   393216
                  Format          =   46006273
                  CurrentDate     =   39850
               End
               Begin MSComCtl2.DTPicker TxtAMCExpDate 
                  Height          =   330
                  Left            =   10680
                  TabIndex        =   23
                  Top             =   720
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   582
                  _Version        =   393216
                  Format          =   46006273
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
                  Left            =   2415
                  TabIndex        =   52
                  Top             =   757
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
                  Left            =   4560
                  TabIndex        =   51
                  Top             =   263
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
                  Left            =   9645
                  TabIndex        =   50
                  Top             =   240
                  Width           =   1035
               End
               Begin VB.Label Label18 
                  Caption         =   "Asset Qty"
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
                  Left            =   120
                  TabIndex        =   49
                  Top             =   772
                  Width           =   1095
               End
               Begin VB.Label Label27 
                  Caption         =   "Remarks"
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
                  Left            =   2400
                  TabIndex        =   48
                  Top             =   1230
                  Width           =   1095
               End
               Begin VB.Label Label28 
                  Caption         =   "AMC Exp. Date"
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
                  Left            =   9240
                  TabIndex        =   47
                  Top             =   750
                  Width           =   1395
               End
               Begin VB.Label Label29 
                  Caption         =   "AMC Amount"
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
                  Left            =   120
                  TabIndex        =   46
                  Top             =   1245
                  Width           =   1200
               End
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
               Left            =   -74880
               TabIndex        =   114
               Top             =   547
               Width           =   1695
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
               Height          =   255
               Left            =   -74880
               TabIndex        =   112
               Top             =   1020
               Width           =   975
            End
         End
         Begin VB.TextBox TxtDesignationID 
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
            Left            =   2280
            Locked          =   -1  'True
            TabIndex        =   64
            TabStop         =   0   'False
            Top             =   1074
            Width           =   2940
         End
         Begin VB.TextBox TxtFlag 
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
            Left            =   7065
            Locked          =   -1  'True
            TabIndex        =   66
            TabStop         =   0   'False
            Top             =   1074
            Width           =   2625
         End
         Begin VB.TextBox TxtDirectReporting 
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
            Left            =   2280
            Locked          =   -1  'True
            TabIndex        =   63
            TabStop         =   0   'False
            Top             =   1980
            Width           =   10215
         End
         Begin VB.TextBox TxtConnectedRAG 
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
            Left            =   12840
            Locked          =   -1  'True
            TabIndex        =   67
            TabStop         =   0   'False
            Top             =   1920
            Visible         =   0   'False
            Width           =   2775
         End
         Begin VB.TextBox TxtBaseLocationID 
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
            Height          =   360
            Left            =   2280
            Locked          =   -1  'True
            TabIndex        =   100
            TabStop         =   0   'False
            Top             =   3960
            Visible         =   0   'False
            Width           =   7380
         End
         Begin VB.CheckBox ChkAllocateExpense 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   9825
            TabIndex        =   11
            Top             =   4035
            Visible         =   0   'False
            Width           =   225
         End
         Begin VB.Label Label32 
            Caption         =   "HR Direct Reportee"
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
            TabIndex        =   116
            Top             =   2460
            Width           =   1935
         End
         Begin VB.Label Label31 
            Caption         =   "HR Designation"
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
            TabIndex        =   102
            Top             =   1560
            Width           =   1695
         End
         Begin VB.Label Label13 
            Caption         =   "Connected RAG"
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
            Left            =   12720
            TabIndex        =   82
            Top             =   1920
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label Label3 
            Caption         =   "Employee Name"
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
            Left            =   5280
            TabIndex        =   81
            Top             =   687
            Width           =   1815
         End
         Begin VB.Label Label8 
            Caption         =   "StoreMan Role"
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
            TabIndex        =   80
            Top             =   1112
            Width           =   1695
         End
         Begin VB.Label Label2 
            Caption         =   "Employee No"
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
            TabIndex        =   79
            Top             =   672
            Width           =   1695
         End
         Begin VB.Label Label1 
            Caption         =   "Employee ID"
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
            Left            =   150
            TabIndex        =   78
            Top             =   217
            Width           =   1695
         End
         Begin VB.Label Label4 
            Caption         =   "Direct Reportee"
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
            Left            =   120
            TabIndex        =   77
            Top             =   2010
            Width           =   2055
         End
         Begin VB.Label Label5 
            Caption         =   "Connected DRC"
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
            Left            =   9120
            TabIndex        =   76
            Top             =   165
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.Label Label6 
            Caption         =   "Employee EmailID"
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
            TabIndex        =   75
            Top             =   2895
            Width           =   1695
         End
         Begin VB.Label Label11 
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
            Height          =   285
            Left            =   6255
            TabIndex        =   74
            Top             =   1112
            Width           =   735
         End
         Begin VB.Label Label20 
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
            Height          =   300
            Left            =   7920
            TabIndex        =   73
            Top             =   2895
            Width           =   975
         End
         Begin VB.Label Label21 
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
            Height          =   300
            Left            =   120
            TabIndex        =   72
            Top             =   3450
            Width           =   1695
         End
         Begin VB.Label Label25 
            Caption         =   "CMPM-cum-DRC"
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
            Left            =   10530
            TabIndex        =   71
            Top             =   1140
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.Label Label26 
            Caption         =   "Allocate Expense"
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
            Left            =   7680
            TabIndex        =   70
            Top             =   4035
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.Label Label30 
            Caption         =   "CMG Staff"
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
            Left            =   10680
            TabIndex        =   69
            Top             =   4035
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label LblLocationID 
            Caption         =   "Base Location"
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
            Left            =   105
            TabIndex        =   68
            Top             =   4020
            Width           =   2010
         End
      End
   End
End
Attribute VB_Name = "FrmMst_Employee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg, mDesignationID, mFlag, mDirectReportingID, mOldDesignationID, mAdditionalReportingID, mConnectedRAGID As Variant
Dim mAssetTypeId, mAssetStatusID, mEmployeeID, mSM_Prefix, mReportingDesignationID, mSecReportingDesignationID As Variant
Dim mLocationID, mCMPID, mBaseLocationID As Double
Dim mAllocationType, mAllocationID As Variant
Dim mHRLocationID  As Variant
Dim mPreDirectReportingID As Variant
Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)
CmbDesignationID.Visible = True
CmbDesignationID.Enabled = True
CmbDirectReporting.Visible = True
'CmbAdditionalReporting.Visible = True
CmbConnectedRAG.Visible = True
'CmbAdditionalReporting.Enabled = False
CmbConnectedRAG.Enabled = False
CmbFlag.Visible = True
'CmbCMPID.Enabled = True
'CmbLocationID.Enabled = True
Frame2.Enabled = False
Frame3.Enabled = True
Frame5.Enabled = True

Label25.Visible = False
ChkSelfReporting.Visible = False
CmbBaseLocationID.Visible = True

TxtDesignationID.Visible = False
TxtDirectReporting.Visible = False
TxtAdditionalReporting.Visible = False
TxtConnectedRAG.Visible = False
TxtFlag.Visible = False
TxtBaseLocationID.Visible = False

CmdBaseLocationID.Visible = True
CmdDesignationID.Visible = True
CmdLocationID.Visible = True
CmdCMPID.Visible = True
cmdAssetStatus.Visible = True
CmdAssetType.Visible = True
ChkAllocateExpense.Enabled = True
ChkAllocateExpense.Value = 0
ChkCMPAllocation.Enabled = False
ChkCMPAllocation.Value = 1
TxtEmpNo.Locked = False
TxtEmpName.Locked = False
TxtEmpEmailID.Locked = False
TxtMobileNo.Locked = False
TxtAddress.Locked = False
TxtCreated = ""
TxtUpdated = ""
TxtEmpID = ""
TxtHrDesignation = ""
TxtHRDirectReportee = ""
TxtEmpNo = ""
TxtEmpName = ""
TxtEmpEmailID = ""
TxtMobileNo = ""
TxtAddress = ""
'Frame2.Enabled = True
'Frame3.Enabled = True
mDirectReportingID = Null
mAdditionalReportingID = Null
mFlag = Null
CmbDesignationID.Text = ""
CmbDirectReporting.Text = ""
'CmbAdditionalReporting.Text = ""
CmbConnectedRAG.Text = ""
CmbLocationID.Text = ""
CmbFlag.Text = ""
CmbCMPID.Text = ""
CmbAssetType.Text = ""
CmbAssetStatus.Text = ""
ChkSelfReporting.Enabled = True
CmdAddToGrid.Caption = "Add To Grid"
CmdAdd2GridAsset.Caption = "Add To Grid"
AssetVerifiedon.Value = Date
TxtAMCExpDate.Value = Date
TxtAssetDescription = ""
TxtAssetQty = ""
TxtRemarks = ""
TxtAMCAmount = ""

SSTab1.Tab = 0
Call Grid_Head
Call Grid_Head2
Call Grid_HeadAllocation
End Sub

Private Sub Grid_HeadAllocation()
Call ClearAllocationFrame
With MSHAllocation
    .Clear
    .Rows = 2
    .Cols = 4

    .TextMatrix(0, 0) = "AllocationDetailID"
    .TextMatrix(0, 1) = "Allocation Type"
    .TextMatrix(0, 2) = "Region\State\Location"
    .TextMatrix(0, 3) = "AllocationID"
    
    .ColWidth(0) = 0
    .ColWidth(1) = 1800
    .ColWidth(2) = 2500
    .ColWidth(3) = 0
End With
End Sub

Private Sub AssetVerifiedon_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "AssetVerifiedon")
End If
End Sub

Private Sub ChkSelfReporting_Click()
If SaveCmd.Enabled = False Then Exit Sub
'If ChkSelfReporting.Value = 0 Then
'   CmbAdditionalReporting.Enabled = True
'   CmbAdditionalReporting.SetFocus
'Else
'   CmbAdditionalReporting.Enabled = False
'
'End If
'CmbAdditionalReporting.Text = ""
mAdditionalReportingID = Null
End Sub

Private Sub CmbAdditionalReporting_GotFocus()
tmp = CmbAdditionalReporting.Text
'Call Mst_EmployeeTable("Where Flag = 'Y'")
RsMst_Employee.Filter = "DesignationID =  " & 5
CmbAdditionalReporting.Clear
If RsMst_Employee.RecordCount = 0 Then
   MsgBox "No Employee Found!!!"
   CmbAdditionalReporting.SetFocus
   RsMst_Employee.Filter = "DesignationId <> Null"
   'RsMst_Employee.MoveFirst
   Exit Sub
End If
For I = 1 To RsMst_Employee.RecordCount
    CmbAdditionalReporting.AddItem RsMst_Employee!EmployeeName & "~" & RsMst_Employee!EmployeeNo
    RsMst_Employee.MoveNext
Next
CmbAdditionalReporting.Text = tmp
End Sub

Private Sub CmbAdditionalReporting_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "CmbAdditionalReporting")
End If
End Sub

Private Sub CmbAdditionalReporting_LostFocus()
Dim strSplitString() As String
Dim EmpNo_, EmpName_ As String

If CmbAdditionalReporting.Text = "" Then
   mAdditionalReportingID = Null
   RsMst_Employee.Filter = "DesignationId <> Null"
   Exit Sub
End If

strSplitString = Split(CmbAdditionalReporting, "~")
EmpName_ = Trim(strSplitString(0))
EmpNo_ = Trim(strSplitString(1))

RsMst_Employee.MoveFirst

RsMst_Employee.Find "EmployeeNo = '" & EmpNo_ & "'"
If RsMst_Employee.EOF Then
   MsgBox "Invalid Selection "
   CmbAdditionalReporting.SetFocus
   Exit Sub
End If
mAdditionalReportingID = RsMst_Employee!EmployeeID
RsMst_Employee.Filter = "DesignationId <> Null"
RsMst_Employee.MoveFirst
End Sub

Private Sub CmbAllocationType_LostFocus()
If CmbAllocationType.Text = "" Then
    CmbAllocationTypeID.Text = ""
    Exit Sub
End If

If LCase(CmbAllocationType.Text) <> "location" And LCase(CmbAllocationType.Text) <> "state" And LCase(CmbAllocationType.Text) <> "region" Then 'And LCase(CmbAllocationType.Text) <> "cluster" Then
    MsgBox "Invalid Selection!!!"
    CmbAllocationTypeID.Text = ""
    CmbAllocationType.SetFocus
    Exit Sub
End If

If mAllocationType <> Left$(CmbAllocationType, 1) Then
    CmbAllocationTypeID.Text = ""
End If
mAllocationType = Left$(CmbAllocationType, 1)
lblstateregionlocation.Caption = CmbAllocationType.Text
End Sub

Private Sub CmbAllocationTypeID_GotFocus()
If CmbAllocationType.Text = "" Then
    MsgBox "Select Allocation Type First!!!"
    CmbAllocationType.SetFocus
    Exit Sub
End If

tmp1 = CmbAllocationTypeID.Text

If LCase(CmbAllocationType.Text) = "location" Then
    tmp = " Select * from Mst_Location "
    Call TmpTable
    CmbAllocationTypeID.Clear
    If TmpRs.RecordCount = 0 Then
       MsgBox "No Record found !!!! "
       Exit Sub
    End If
    For I = 1 To TmpRs.RecordCount
        CmbAllocationTypeID.AddItem TmpRs!LocationName
        TmpRs.MoveNext
    Next
ElseIf LCase(CmbAllocationType.Text) = "state" Then
    tmp = " Select * from Mst_State"
    Call TmpTable
    CmbAllocationTypeID.Clear
    If TmpRs.RecordCount = 0 Then
       MsgBox "No Record found !!!! "
       Exit Sub
    End If
    For I = 1 To TmpRs.RecordCount
        CmbAllocationTypeID.AddItem TmpRs!StateName
        TmpRs.MoveNext
    Next
ElseIf LCase(CmbAllocationType.Text) = "region" Then
    tmp = " Select * from Mst_Region"
    Call TmpTable
    CmbAllocationTypeID.Clear
    If TmpRs.RecordCount = 0 Then
       MsgBox "No Record found !!!! "
       Exit Sub
    End If
    For I = 1 To TmpRs.RecordCount
        CmbAllocationTypeID.AddItem TmpRs!Region
        TmpRs.MoveNext
    Next
'ElseIf LCase(CmbAllocationType.Text) = "cluster" Then
'    tmp = " Select * from Mst_Cluster"
'    Call TmpTable
'    CmbAllocationTypeID.Clear
'    If TmpRs.RecordCount = 0 Then
'       MsgBox "No Record found !!!! "
'       Exit Sub
'    End If
'    For I = 1 To TmpRs.RecordCount
'        CmbAllocationTypeID.AddItem TmpRs!ClusterName
'        TmpRs.MoveNext
'    Next
End If
CmbAllocationTypeID.Text = tmp1
End Sub

Private Sub CmbAllocationTypeID_LostFocus()
If CmbAllocationTypeID.Text = "" Then Exit Sub

If LCase(CmbAllocationType.Text) = "location" Then
   If TmpRs.RecordCount > 0 Then TmpRs.MoveFirst
   TmpRs.Find "LocationName = '" & CmbAllocationTypeID.Text & "'"
   If TmpRs.EOF Then
      MsgBox "Invalid selection "
      CmbAllocationTypeID.SetFocus
      Exit Sub
   End If
   mAllocationID = TmpRs!LocationID
ElseIf LCase(CmbAllocationType.Text) = "state" Then
   If TmpRs.RecordCount > 0 Then TmpRs.MoveFirst
   TmpRs.Find "StateName = '" & CmbAllocationTypeID.Text & "'"
   If TmpRs.EOF Then
      MsgBox "Invalid selection "
      CmbAllocationTypeID.SetFocus
      Exit Sub
   End If
   mAllocationID = TmpRs!StateID
ElseIf LCase(CmbAllocationType.Text) = "region" Then
   If TmpRs.RecordCount > 0 Then TmpRs.MoveFirst
   TmpRs.Find "Region= '" & CmbAllocationTypeID.Text & "'" '"Region= '" & CmbAllocationTypeID.Text & "'"
   If TmpRs.EOF Then
      MsgBox "Invalid selection "
      CmbAllocationTypeID.SetFocus
      Exit Sub
   End If
   mAllocationID = TmpRs!RegionID
'ElseIf LCase(CmbAllocationType.Text) = "cluster" Then
'   If TmpRs.RecordCount > 0 Then TmpRs.MoveFirst
'   TmpRs.Find "ClusterName = '" & CmbAllocationTypeID.Text & "'"
'   If TmpRs.EOF Then
'      MsgBox "Invalid selection "
'      CmbAllocationTypeID.SetFocus
'      Exit Sub
'   End If
'   mAllocationID = TmpRs!ClusterID
End If
End Sub

Private Sub CmbAssetStatus_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "CmbAssetStatus")
End If
End Sub

Private Sub CmbAssetType_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "CmbAssetType")
End If
End Sub



Private Sub CmbBaseLocationID_GotFocus()
tmp = CmbBaseLocationID.Text
Call TableMst_Location(" Where HRLocationID=" & Val(mHRLocationID))
CmbBaseLocationID.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Location found"
   CmbBaseLocationID.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbBaseLocationID.AddItem RsMst_Location!LocationName
    RsMst_Location.MoveNext
Next
CmbBaseLocationID.Text = tmp

End Sub

Private Sub CmbBaseLocationID_LostFocus()
If CmbBaseLocationID.Text = "" Then
   Call Grid_HeadAllocation
   mBaseLocationID = 0
   Exit Sub
End If

RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbBaseLocationID.Text & "'"

If RsMst_Location.EOF Then
   MsgBox "Invalid Location  selection "
   CmbBaseLocationID.SetFocus
   Exit Sub
End If
If mBaseLocationID <> RsMst_Location!LocationID Then
    Call Grid_HeadAllocation

    MSHAllocation.Rows = MSHAllocation.Rows + 1
    MSHAllocation.TextMatrix(1, 1) = "Location" 'CmbAllocationType
    MSHAllocation.TextMatrix(1, 2) = CmbBaseLocationID 'CmbAllocationTypeID
    MSHAllocation.TextMatrix(1, 3) = RsMst_Location!LocationID 'mAllocationID
End If
mBaseLocationID = RsMst_Location!LocationID

End Sub

Private Sub CmbConnectedRAG_GotFocus()

temp = CmbConnectedRAG.Text
'Call Mst_EmployeeTable("Where Flag = 'Y'")
RsMst_Employee.Filter = "DesignationId = " & 7
CmbConnectedRAG.Clear
If RsMst_Employee.RecordCount = 0 Then
   MsgBox "No Employee Found!!!"
   CmbConnectedRAG.SetFocus
   RsMst_Employee.Filter = "DesignationId <> Null"
   'RsMst_Employee.MoveFirst
   Exit Sub
End If
For I = 1 To RsMst_Employee.RecordCount
    CmbConnectedRAG.AddItem RsMst_Employee!EmployeeName & "~" & RsMst_Employee!EmployeeNo
    RsMst_Employee.MoveNext
Next
CmbConnectedRAG.Text = temp

End Sub

Private Sub CmbConnectedRAG_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "CmbConnectedRAG")
End If
End Sub

Private Sub CmbConnectedRAG_LostFocus()
Dim strSplitString() As String
Dim EmpNo_, EmpName_ As String

If CmbConnectedRAG.Text = "" Then
   mConnectedRAGID = Null
   RsMst_Employee.Filter = "DesignationId <> Null"
   Exit Sub
End If

strSplitString = Split(CmbConnectedRAG, "~")
EmpName_ = Trim(strSplitString(0))
EmpNo_ = Trim(strSplitString(1))


RsMst_Employee.MoveFirst
RsMst_Employee.Find "EmployeeNo = '" & EmpNo_ & "'"

If RsMst_Employee.EOF Then
   MsgBox "Invalid Selection "
   CmbConnectedRAG.SetFocus
   Exit Sub
End If
mConnectedRAGID = RsMst_Employee!EmployeeID
RsMst_Employee.Filter = "DesignationId <> Null"
RsMst_Employee.MoveFirst
End Sub

Private Sub CmbDesignationID_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "CmbDesignationID")
End If
End Sub

Private Sub CmbDirectReporting_GotFocus()

tmp1 = CmbDirectReporting.Text
'Call Mst_EmployeeTable("Where Flag = 'Y'")

'Commented on 07 Aug 2010 Parag
'If mDesignationID = 1 Then
'   RsMst_Employee.Filter = "DesignationId = " & 1
'ElseIf mDesignationID = 2 Then
'   RsMst_Employee.Filter = "DesignationId = " & 1
'ElseIf mDesignationID = 3 Then
'   RsMst_Employee.Filter = "DesignationId = " & 2
'ElseIf mDesignationID = 4 Then
'   RsMst_Employee.Filter = "DesignationId = " & 3
'ElseIf mDesignationID = 5 Then
'   RsMst_Employee.Filter = "DesignationId = " & 6
'ElseIf mDesignationID = 6 Then
'   'RsMst_Employee.Filter = "DesignationId =  4 Or DesignationID = 3 "
'   RsMst_Employee.Filter = " DesignationID = 3 "
'ElseIf mDesignationID = 7 Then
'   RsMst_Employee.Filter = "DesignationId = " & 1
'ElseIf mDesignationID = 8 Then                   'QQS
'   RsMst_Employee.Filter = "DesignationId =  " & 6
'ElseIf mDesignationID = 9 Then                   'BDO
'   'RsMst_Employee.Filter = "DesignationId = 2 Or DesignationID = 3 "
'   RsMst_Employee.Filter = "DesignationID = 3 "
'ElseIf mDesignationID = 10 Then                   'BDO
'   'RsMst_Employee.Filter = "DesignationId = 2 Or DesignationID = 3 "
'   RsMst_Employee.Filter = "DesignationID = 3 "
'End If

If CmbDesignationID.Text = "" Then
   MsgBox "Select StoreMan Group!!!"
   CmbDesignationID.SetFocus
   Exit Sub
End If
wc = ""

tmp = "Select * from Mst_Employee  "

If IsNull(mReportingDesignationID) = False Then
   wc = " Where DesignationID = " & mReportingDesignationID & " "
Else
   wc = " Where DesignationID = " & mDesignationID & " "
End If

If IsNull(mSecReportingDesignationID) = False Or mSecReportingDesignationID <> 0 Then
   wc = Replace(wc, "Where", " where (")
   wc = wc & " Or DesignationID = " & mSecReportingDesignationID & ")"
End If

If Edit_Flg = "E" And mDirectReportingID > 0 Then
   wc = wc & " And (Flag='Y' or EmployeeID=" & mDirectReportingID & ")"
Else
   wc = wc & " And Flag='Y'"
End If

tmp = tmp & wc

Call TmpTable

CmbDirectReporting.Clear


If TmpRs.RecordCount = 0 Then
   MsgBox "No Employee Found!!!"
   CmbDirectReporting.SetFocus
   Exit Sub
End If
For I = 1 To TmpRs.RecordCount
    CmbDirectReporting.AddItem TmpRs!EmployeeName & "~" & TmpRs!EmployeeNo
    TmpRs.MoveNext
Next
CmbDirectReporting.Text = tmp1
End Sub

Private Sub CmbDirectReporting_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "CmbDirectReporting")
End If
End Sub

Private Sub CmbDirectReporting_LostFocus()
Dim strSplitString() As String
Dim EmpNo_, EmpName_ As String

If CmbDirectReporting.Text = "" Then
  '' mDirectReportingID = Null
   RsMst_Employee.Filter = "DesignationId <> Null"
   Exit Sub
End If

strSplitString = Split(CmbDirectReporting, "~")
If UBound(strSplitString) <> 1 Then
   MsgBox "Invalid Selection "
   CmbDirectReporting.SetFocus
   Exit Sub
End If

EmpName_ = Trim(strSplitString(0))
EmpNo_ = Trim(strSplitString(1))


TmpRs.MoveFirst

TmpRs.Find "EmployeeNo = '" & EmpNo_ & "'"
If TmpRs.EOF Then
   MsgBox "Invalid Selection "
   CmbDirectReporting.SetFocus
   Exit Sub
End If
mDirectReportingID = TmpRs!EmployeeID

End Sub

Private Sub CmbFlag_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "CmbFlag")
End If
End Sub

Private Sub CmbFlag_LostFocus()
If CmbFlag.Text = "" Then Exit Sub
If CmbFlag.Text = "Active" Then
   mFlag = "Y"
ElseIf CmbFlag.Text = "De-Active" Then
   mFlag = "N"
Else
   MsgBox "Invalid Selection!!!"
   CmbFlag.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmbCMPID_GotFocus()
If CmbLocationID.Text = "" Then Exit Sub
temp = CmbCMPID.Text
'tmp = " Select mg.GodownID,mg.GodownNO,mg.Address,mc.CMPID " & _
'      " From Mst_Godown mg " & _
'      " Inner Join Mst_CMP mc on mg.CMPID = mc.CMPID " & _
'      " Where mc.locationID = " & mLocationId & " and mg.CloseDate is null "

tmp = "Select * from Mst_CMP Where LocationID = " & mLocationID & " And CloseDate is null"
Call Tmp4Table

CmbCMPID.Clear

If TmpRs4.RecordCount = 0 Then
   MsgBox "No CMP Found in Location selected!!!"
   CmbCMPID.SetFocus
   Exit Sub
End If

For I = 1 To TmpRs4.RecordCount
    CmbCMPID.AddItem TmpRs4!CMPName
    TmpRs4.MoveNext
Next

CmbCMPID.Text = temp

End Sub

Private Sub CmbCMPID_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "CmbCMPID")
End If
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

Private Sub CmbLocationID_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "CmbLocationID")
End If
End Sub

Private Sub CmbSearchStatus_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "CmbSearchStatus")
End If
End Sub

Private Sub CmbSearchStatus_LostFocus()
If CmbSearchStatus.Text = "" Then Exit Sub

If CmbSearchStatus.Text = "Active" Or CmbSearchStatus.Text = "De-Active" Then
   
Else
   MsgBox "Invalid selection!!!"
   CmbSearchStatus.SetFocus
   Exit Sub

End If

'Dim x As Integer
'If CmbSearchStatus.Text = "" Then
'   Call Grid_Head1
'   Exit Sub
'End If
'
'If CmbSearchStatus.Text = "Active" Then
'   mFlag = "Y"
'
'ElseIf CmbSearchStatus.Text = "De-Active" Then
'   mFlag = "N"
'Else
'   MsgBox "Invalid Selection!!!"
'   CmbSearchStatus.SetFocus
'   Exit Sub
'End If
'tmp = "Select * from Mst_Employee where Flag = '" & mFlag & "'"
'Call Tmp1Table
'If TmpRs1.RecordCount = 0 Then
'      MsgBox "No such Record Found!!!"
'      Call Grid_Head1
'      Exit Sub
'End If
'MSHFlexGrid1.Rows = TmpRs1.RecordCount + 1
'For x = 1 To TmpRs1.RecordCount
'    Call ShowGrid(x)
'    TmpRs1.MoveNext
'Next
End Sub

Private Sub CmdAdd2GridAllocation_Click()
If CmbAllocationType = "" Then
   MsgBox "Blank Allocation Type Not Allowed."
   CmbAllocationType.SetFocus
   Exit Sub
End If
If CmbAllocationTypeID = "" Then
   MsgBox "Blank " & CmbAllocationType & "  Not Allowed."
   CmbAllocationTypeID.SetFocus
   Exit Sub
End If
If CheckDuplicateAllocation = True Then
   'MsgBox "Duplicate Allocation not Allowed."
   CmbAllocationType.SetFocus
   Exit Sub
End If

If CmdAdd2GridAllocation.Caption <> "Add To Grid" Then
   r = MSHAllocation.RowSel
Else
   r = MSHAllocation.Rows - 1
   MSHAllocation.Rows = MSHAllocation.Rows + 1
End If
MSHAllocation.TextMatrix(r, 1) = CmbAllocationType
MSHAllocation.TextMatrix(r, 2) = CmbAllocationTypeID
MSHAllocation.TextMatrix(r, 3) = mAllocationID

Call ClearAllocationFrame

End Sub

Private Sub ClearAllocationFrame()
CmbAllocationType = ""
CmbAllocationTypeID = ""
CmdAdd2GridAllocation.Caption = "Add To Grid"
CmbAllocationType.Enabled = True
CmdDeleteFromGridAllocation.Enabled = False
End Sub

Function CheckDuplicateAllocation() As Boolean
Dim Retval As Boolean
Retval = False
If MSHAllocation.Rows > 2 Then
   For u = 1 To MSHAllocation.Rows - 2
       If LCase(MSHAllocation.TextMatrix(u, 1)) = LCase(CmbAllocationType.Text) And LCase(MSHAllocation.TextMatrix(u, 2)) = LCase(CmbAllocationTypeID.Text) Then
          If CmdAdd2GridAllocation.Caption = "Add To Grid" Then
             MsgBox "Duplicate Allocation not Allowed."
             Retval = True
             CheckDuplicateAllocation = Retval
             Exit Function
             'Exit For
          Else
             If u <> MSHAllocation.RowSel Then
                MsgBox "Duplicate Allocation not Allowed."
                Retval = True
                CheckDuplicateAllocation = Retval
                Exit Function
             End If
          End If
       End If
   Next
End If

For I = 1 To MSHAllocation.Rows - 2
    If CmdAdd2GridAllocation.Caption <> "Add To Grid" Then
        If I = MSHAllocation.RowSel Then
           GoTo Nextrow2
        End If
    End If
        
    If LCase(MSHAllocation.TextMatrix(I, 1)) = "location" Then
       If LCase(CmbAllocationType.Text) = "state" Then
          tmp = "Select * From Mst_Location Where StateID = " & mAllocationID
          tmp = tmp & " And LocationID = " & Val(MSHAllocation.TextMatrix(I, 3))
          Call Tmp3Table
          If TmpRs3.RecordCount > 0 Then
             MsgBox "Location " & MSHAllocation.TextMatrix(I, 2) & " already mapped under " & CmbAllocationTypeID.Text & " State!!!"
             Retval = True
             Exit For
          End If
       ElseIf LCase(CmbAllocationType.Text) = "region" Then
          tmp = "Select ML.LocationID From Mst_Location ML "
          tmp = tmp & " Inner Join Mst_state MS on MS.StateID = ML.StateID"
          tmp = tmp & " Where MS.RegionID = " & mAllocationID
          tmp = tmp & " And LocationID = " & Val(MSHAllocation.TextMatrix(I, 3))
          Call Tmp3Table
          If TmpRs3.RecordCount > 0 Then
             MsgBox "Location " & MSHAllocation.TextMatrix(I, 2) & " already mapped under " & CmbAllocationTypeID.Text & "!!!"
             Retval = True
             Exit For
          End If
       'ElseIf LCase(CmbAllocationType.Text) = "cluster" Then
       '   tmp = "Select ML.LocationID From Mst_Location ML "
       '   tmp = tmp & " Inner Join Mst_District MD on MD.DistrictID = ML.DistrictID"
       '   tmp = tmp & " Where MD.ClusterID = " & mAllocationID
       '   tmp = tmp & " And LocationID = " & Val(MSHAllocation.TextMatrix(I, 3))
       '   Call Tmp3Table
       '   If TmpRs3.RecordCount > 0 Then
       '      MsgBox "Location " & MSHAllocation.TextMatrix(I, 2) & " already mapped under " & CmbAllocationTypeID.Text & " Cluster!!!"
       '      Retval = True
       '      Exit For
       '   End If
       End If
    ElseIf LCase(MSHAllocation.TextMatrix(I, 1)) = "state" Then
       If LCase(CmbAllocationType.Text) = "location" Then
          tmp = "Select * From Mst_Location Where locationID = " & mAllocationID
          tmp = tmp & " And StateID = " & Val(MSHAllocation.TextMatrix(I, 3))
          Call Tmp3Table
          If TmpRs3.RecordCount > 0 Then
             MsgBox "Location " & CmbAllocationTypeID.Text & " already mapped under " & MSHAllocation.TextMatrix(I, 2) & " State!!!"
             Retval = True
             Exit For
          End If
       ElseIf LCase(CmbAllocationType.Text) = "region" Then
          tmp = "Select MS.StateID From Mst_State MS "
          tmp = tmp & " Where MS.RegionID = " & mAllocationID
          tmp = tmp & " And StateID = " & Val(MSHAllocation.TextMatrix(I, 3))
          Call Tmp3Table
          If TmpRs3.RecordCount > 0 Then
             MsgBox "Location " & MSHAllocation.TextMatrix(I, 2) & " already mapped under " & CmbAllocationTypeID.Text & "!!!"
             Retval = True
             Exit For
          End If
'       ElseIf LCase(CmbAllocationType.Text) = "cluster" Then
'          tmp = "Select ML.LocationID From Mst_Location ML "
'          tmp = tmp & " Inner Join Mst_District MD on MD.DistrictID = ML.DistrictID"
'          tmp = tmp & " Where MD.ClusterID = " & mAllocationID
'          tmp = tmp & " And ML.StateID = " & Val(MSHAllocation.TextMatrix(I, 3))
'          Call Tmp3Table
'          If TmpRs3.RecordCount > 0 Then
'             MsgBox "State " & MSHAllocation.TextMatrix(I, 2) & "  and Cluster " & CmbAllocationTypeID.Text & "  has one or more same location!!!"
'             Retval = True
'             Exit For
'          End If
       End If
    ElseIf LCase(MSHAllocation.TextMatrix(I, 1)) = "region" Then
       If LCase(CmbAllocationType.Text) = "location" Then
          tmp = "Select ML.LocationID From Mst_Location ML "
          tmp = tmp & " Inner Join Mst_State MS on MS.StateID = ML.StateID "
          tmp = tmp & " Where MS.RegionID = " & Val(MSHAllocation.TextMatrix(I, 3))
          tmp = tmp & " And ML.LocationID = " & mAllocationID
          Call Tmp3Table
          If TmpRs3.RecordCount > 0 Then
             MsgBox "Location " & CmbAllocationTypeID.Text & " already mapped under " & MSHAllocation.TextMatrix(I, 2) & "!!!"
             Retval = True
             Exit For
          End If
       ElseIf LCase(CmbAllocationType.Text) = "state" Then
          tmp = "Select MS.StateID From Mst_State MS "
          tmp = tmp & " Where MS.StateID = " & mAllocationID
          tmp = tmp & " And MS.RegionID = " & Val(MSHAllocation.TextMatrix(I, 3))
          Call Tmp3Table
          If TmpRs3.RecordCount > 0 Then
             MsgBox "State " & CmbAllocationTypeID.Text & " already mapped under " & MSHAllocation.TextMatrix(I, 2) & "!!!"
             Retval = True
             Exit For
          End If
'       ElseIf LCase(CmbAllocationType.Text) = "cluster" Then
'          tmp = "Select ML.LocationID From Mst_Location ML "
'          tmp = tmp & " Inner Join Mst_State MS on MS.StateID = ML.StateID"
'          tmp = tmp & " Inner Join Mst_District MD on MD.DistrictID = ML.DistrictID"
'          tmp = tmp & " Where MD.ClusterID = " & mAllocationID
'          tmp = tmp & " And MS.RegionID = " & Val(MSHAllocation.TextMatrix(I, 3))
'          Call Tmp3Table
'          If TmpRs3.RecordCount > 0 Then
'             MsgBox "Cluster " & CmbAllocationTypeID.Text & " already mapped under " & MSHAllocation.TextMatrix(I, 2) & "!!!"
'             Retval = True
'             Exit For
'          End If
       End If
'    ElseIf LCase(MSHAllocation.TextMatrix(I, 1)) = "cluster" Then
'       If LCase(CmbAllocationType.Text) = "location" Then
'          tmp = "Select ML.LocationID From Mst_Location ML "
'          tmp = tmp & " Inner Join Mst_District MD on MD.DistrictID = ML.DistrictID"
'          tmp = tmp & " And MD.ClusterID = " & Val(MSHAllocation.TextMatrix(I, 3))
'          tmp = tmp & " Where ML.LocationID = " & mAllocationID
'          Call Tmp3Table
'          If TmpRs3.RecordCount > 0 Then
'             MsgBox "Location " & CmbAllocationTypeID.Text & " already mapped under " & MSHAllocation.TextMatrix(I, 2) & " Cluster!!!"
'             Retval = True
'             Exit For
'          End If
'       ElseIf LCase(CmbAllocationType.Text) = "state" Then
'          tmp = "Select ML.LocationID From Mst_Location ML "
'          tmp = tmp & " Inner Join Mst_State MS on MS.StateID = ML.StateID"
'          tmp = tmp & " Inner Join Mst_District MD on MD.DistrictID = ML.DistrictID"
'          tmp = tmp & " Where MS.StateID = " & mAllocationID
'          tmp = tmp & " And MD.ClusterID = " & Val(MSHAllocation.TextMatrix(I, 3))
'          Call Tmp3Table
'          If TmpRs3.RecordCount > 0 Then
'             MsgBox "Cluster " & MSHAllocation.TextMatrix(I, 2) & " already mapped under " & CmbAllocationTypeID.Text & " State!!!"
'             Retval = True
'             Exit For
'          End If
'       ElseIf LCase(CmbAllocationType.Text) = "region" Then
'          tmp = "Select ML.LocationID From Mst_Location ML "
'          tmp = tmp & " Inner Join Mst_State MS on MS.StateID = ML.StateID"
'          tmp = tmp & " Inner Join Mst_District MD on MD.DistrictID = ML.DistrictID"
'          tmp = tmp & " Where MS.RegionID = " & mAllocationID
'          tmp = tmp & " And MD.ClusterID = " & Val(MSHAllocation.TextMatrix(I, 3))
'          Call Tmp3Table
'          If TmpRs3.RecordCount > 0 Then
'             MsgBox "Cluster " & MSHAllocation.TextMatrix(I, 2) & " already mapped under " & CmbAllocationTypeID.Text & "!!!"
'             Retval = True
'             Exit For
'          End If
'       End If
    End If
Nextrow2:
Next
CheckDuplicateAllocation = Retval
End Function


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

If TxtAMCExpDate.Value < Date Then
   MsgBox "Back AMC Expiry date not allowed!!!"
   TxtAMCExpDate.SetFocus
   Exit Sub
End If

'If TxtAssetDescription = "" Then
'   MsgBox "Blank Description Not Allowed!!!"
'   TxtAssetDescription.SetFocus
'   Exit Sub
'End If

Dim I, ans As Variant

If CmdAdd2GridAsset.Caption = "Add To Grid" Then
   I = MSHFlexGrid3.Rows - 1
   MSHFlexGrid3.Rows = MSHFlexGrid3.Rows + 1
Else
   I = MSHFlexGrid3.RowSel
End If

MSHFlexGrid3.TextMatrix(I, 0) = CmbAssetType.Text
MSHFlexGrid3.TextMatrix(I, 1) = CmbAssetStatus.Text
MSHFlexGrid3.TextMatrix(I, 2) = Val(TxtAssetQty)
MSHFlexGrid3.TextMatrix(I, 3) = TxtAssetDescription
MSHFlexGrid3.TextMatrix(I, 4) = AssetVerifiedon.Value
MSHFlexGrid3.TextMatrix(I, 6) = TxtRemarks
MSHFlexGrid3.TextMatrix(I, 7) = TxtAMCExpDate.Value
MSHFlexGrid3.TextMatrix(I, 8) = Val(TxtAMCAmount)

CmdAdd2GridAsset.Caption = "Add To Grid"
CmbAssetType.Text = ""
CmbAssetStatus.Text = ""
TxtAssetQty = ""
TxtAssetDescription = ""
AssetVerifiedon.Value = Date
TxtRemarks = ""
TxtAMCExpDate.Value = Date
TxtAMCAmount = ""

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
       ans = Validate_DuplicateGridForAll()
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

Private Sub cmdAssetStatus_Click()
FrmMst_AssetStatus.Show
End Sub

Private Sub CmdAssetType_Click()
FrmMst_AssetType.Show
End Sub

Private Sub CmdBaseLocationID_Click()
FrmMst_Location.Show
End Sub

Private Sub CmdDeleteFromGridAllocation_Click()
I = MSHAllocation.RowSel
If I > 0 Then
   If MSHAllocation.TextMatrix(I, 1) <> "" Then
      For RO = I To MSHAllocation.Rows - 2
          For C = 0 To MSHAllocation.Cols - 1
              MSHAllocation.TextMatrix(RO, C) = MSHAllocation.TextMatrix(RO + 1, C)
          Next
      Next
      MSHAllocation.Rows = MSHAllocation.Rows - 1
   End If
End If
MSHAllocation.RowSel = 0
Call ClearAllocationFrame
End Sub

Private Sub CmdDesignationID_Click()
FrmMst_Designation.Show
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Mst_Employee.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Mst_Employee.xls")
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

tmp = "Select EMP.EmployeeID,EMP.EmployeeNo,EMP.EmployeeName,D.Designation,EMP1.EmployeeName + '~' + EMP1.EmployeeNo as DirectReporte, " & _
      "EMP2.EmployeeName +'~'+ EMP2.EmployeeNo as ConnectedDRC, " & _
      "EMP3.EmployeeName +'~'+ EMP3.EmployeeNo as ConnectedRAG,EMP.EmployeeEmailID,EMP.MobileNo,EMP.Address, ML.LocationName," & _
      "EMP.Flag,EMP.SelfReporting,EMP.AllocateExpense  From Mst_Employee EMP " & _
      "Left Join Mst_Employee EMP1 on EMP.DirectReportingTo = EMP1.EmployeeID " & _
      "Left Join Mst_Employee EMP2 on EMP.ReportingToDRC = EMP2.EmployeeID " & _
      "Left Join Mst_Employee EMP3 on EMP.ConnectedToRAG = EMP3.EmployeeID " & _
      "Inner Join Mst_Designation D on EMP.DesignationID = D.DesignationID " & _
      "Left Join Mst_Location ML on EMP.BaseLocationID = ML.LocationID "
'      "--Where EMP.DesignationID = 6 or EMP.DesignationID = 7 " & _
'      "Group By EMP.EmployeeID,EMP.EmployeeNo,EMP.EmployeeName,D.Designation,EMP.Flag,EMP1.EmployeeNo,EMP1.EmployeeName, " & _
'      "EMP2.EmployeeNo , EMP2.EmployeeName, EMP3.EmployeeNo, EMP3.EmployeeName, EMP.EmployeeEmailID, EMP.MobileNo, EMP.Address "

Dim wc As Variant
wc = ""
If Option5.Value = True Then
   wc = " Where EMP.EmployeeName Like '%" & TxtSearchEmployee & "%' "
End If
If Option1.Value = True Then
   wc = " Where EMP.EmployeeNo  Like '%" & TxtSearchEmployee & "%'"
End If
If Option2.Value = True Then
   wc = " Where D.Designation  Like '%" & TxtSearchEmployee & "%'"
End If

If CmbSearchStatus.Text <> "" Then
   If wc = "" Then
      If CmbSearchStatus.Text = "Active" Then
         wc = " Where EMP.Flag = 'Y'"
      ElseIf CmbSearchStatus.Text = "De-Active" Then
         wc = " Where EMP.Flag = 'N'"
      End If
   Else
      If CmbSearchStatus.Text = "Active" Then
         wc = wc & " And EMP.Flag = 'Y'"
      ElseIf CmbSearchStatus.Text = "De-Active" Then
         wc = wc & " And EMP.Flag = 'N'"
      End If
   
      'Wc = Wc & " And EMP.Flag = '" & CmbSearchStatus.Text & "'"
   End If
End If
tmp = tmp & " " & wc


tmp1 = "Group By EMP.EmployeeID,EMP.EmployeeNo,EMP.EmployeeName,D.Designation,EMP.Flag,EMP1.EmployeeNo,EMP1.EmployeeName, " & _
"EMP2.EmployeeNo , EMP2.EmployeeName, EMP3.EmployeeNo, EMP3.EmployeeName, EMP.EmployeeEmailID, EMP.MobileNo, EMP.Address,ML.LocationName,EMP.SelfReporting,EMP.AllocateExpense "


tmp = tmp & "" & tmp1

Call TmpTable

Call Grid_Head1
For iu = 1 To TmpRs.RecordCount
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
    For C = 0 To MSHFlexGrid1.Cols - 2
       MSHFlexGrid1.TextMatrix(iu, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
    Next
    If MSHFlexGrid1.TextMatrix(iu, 11) = "1" Then
       MSHFlexGrid1.TextMatrix(iu, 11) = "Yes"
    Else
       MSHFlexGrid1.TextMatrix(iu, 11) = "No"
    End If
    If MSHFlexGrid1.TextMatrix(iu, 12) = "1" Then
       MSHFlexGrid1.TextMatrix(iu, 12) = "Yes"
    Else
       MSHFlexGrid1.TextMatrix(iu, 12) = "No"
    End If
    MSHFlexGrid1.TextMatrix(iu, 14) = GetHRDesignationFromHRPro(MSHFlexGrid1.TextMatrix(iu, 1))
    TmpRs.MoveNext
Next
Label22.Refresh
Label22.Caption = TmpRs.RecordCount
End Sub

Private Sub CmdCMPID_Click()
FrmMst_CMP.Show
End Sub

Private Sub CmdLocationID_Click()
FrmMst_Location.Show
End Sub
Private Sub CmdRemoveFromGrid_Click()
If MSHFlexGrid2.Rows < 3 Then Exit Sub
Dim ans As Variant
If CmbCMPID.Text = "" Then
   MsgBox "Please Select atleast one row from grid!!!"
   Exit Sub
End If
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
'CmbLocationID.Text = ""
   CmbCMPID.Text = ""
   CmbLocationID.SetFocus
   CmdAddToGrid.Caption = "Add To Grid"
Else
   CmbCMPID.SetFocus
End If

End Sub
Private Sub CmdDeleteFromGridAsset_Click()
If MSHFlexGrid3.Rows < 3 Then Exit Sub
Dim ans As Variant
If CmbAssetType.Text = "" Then
   MsgBox "Please Select atleast one row from grid!!!"
   Exit Sub
End If

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
   End If
   CmbAssetType.Text = ""
   CmbAssetStatus.Text = ""
   CmdAddToGrid.Caption = "Add To Grid"
Else
   CmbAssetStatus.SetFocus
End If
End Sub

Private Sub DeleteCmd_Click()
Dim ans As Variant
On Error GoTo msgError
ans = MsgBox("Are you Sure you want to Delete this Record ", vbYesNo)
If ans = vbYes Then
   tmp = "Select * from Mst_Employee where DirectReportingTo = " & TxtEmpID & " Or ReportingToDRC = " & TxtEmpID & " Or ConnectedToRAG = " & TxtEmpID & " "
   Call TmpTable
   If TmpRs.RecordCount > 0 Then
      MsgBox "Child Record Present!!!"
      Exit Sub
   End If
   TmpRs.Close
   tmp = "Delete from Mst_EmpCMPMapping where EmployeeID = " & TxtEmpID
   Call TmpTable

   tmp = "Delete from Mst_EmpAssetMap where EmployeeID = " & TxtEmpID
   Call TmpTable

   tmp = "SELECT * from Mst_EmpCMPMapping where EmployeeID = " & TxtEmpID
   Call TmpTable
   RsMst_Employee.Requery
   RsMst_Employee.Find "EmployeeId = " & TxtEmpID
   If RsMst_Employee.EOF Then Exit Sub
   RsMst_Employee.Delete
   RsMst_Employee.Update
   If RsMst_Employee.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_Employee.MoveNext
   If RsMst_Employee.EOF() Then
      RsMst_Employee.MoveLast
   End If
   Call Indata
   'TmpRs.Close
End If

Exit Sub
msgError:
MsgBox "Child Record Present "
RsMst_Employee.CancelUpdate
End Sub

Private Sub EditCmd_Click()
Edit_Flg = "E"
'Frame2.Enabled = True
CmbDesignationID.Visible = True
CmbDirectReporting.Visible = True
'CmbAdditionalReporting.Visible = True
''CmbConnectedRAG.Visible = True
CmbFlag.Visible = True
CmbCMPID.Enabled = True
CmbBaseLocationID.Visible = True
CmdBaseLocationID.Visible = True
CmbDesignationID.Text = TxtDesignationID

CmbDirectReporting.Text = TxtDirectReporting
'CmbAdditionalReporting.Text = TxtAdditionalReporting
CmbConnectedRAG.Text = TxtConnectedRAG
CmbFlag.Text = TxtFlag
CmbBaseLocationID.Text = TxtBaseLocationID

TxtDesignationID.Visible = False
TxtDirectReporting.Visible = False
TxtAdditionalReporting.Visible = False
''TxtConnectedRAG.Visible = False
TxtFlag.Visible = False
TxtBaseLocationID.Visible = False

CmdDesignationID.Visible = True
CmdLocationID.Visible = True
CmdCMPID.Visible = True
CmdAssetType.Visible = True
cmdAssetStatus.Visible = True

'If LCase(CmbDesignationID.Text) = "managment" Then
'   ChkAllocateExpense.Enabled = True
'Else
'   ChkAllocateExpense.Enabled = False
'End If

If LCase(CmbDesignationID.Text) = "managment" Then
   ChkCMPAllocation.Enabled = True
Else
   ChkCMPAllocation.Enabled = False
End If

''If LCase(Gen_UserRole) = "power" Then
''   TxtEmpNo.Locked = False
''   TxtEmpName.Locked = False
''Else
''   TxtEmpNo.Locked = True
''   TxtEmpName.Locked = True
''End If

''TxtEmpEmailID.Locked = False
''TxtMobileNo.Locked = False
''TxtAddress.Locked = False

If CmbDesignationID.Text = "CMPM" Then
   Frame2.Enabled = True
   ''CmbConnectedRAG.Enabled = True
Else
   Frame2.Enabled = False
   ''CmbConnectedRAG.Enabled = False
End If
ChkSelfReporting.Enabled = True
Frame3.Enabled = True
Frame5.Enabled = True

'If LCase(Gen_UserRole) <> "power" Then
'   CmbDesignationID.Enabled = False
'End If

If LCase(Gen_UserRole) = "power" Or LCase(Gen_UserRole) = "superadmin" Or LCase(Gen_UserRole) = "headcmg" Then
   CmbDesignationID.Enabled = True
Else
   CmbDesignationID.Enabled = False
End If


Call GButtonLock(Me, False)

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

Call TableMst_Employee("Order by EmployeeID ")
Call Grid_Head1
TxtSearchEmployee = ""
CmbSearchStatus.Text = ""
Option4.Value = True
Label22.Caption = ""

SSTab1.Tab = 0
If RsMst_Employee.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
Call Grid_Head
Call Indata
Call GButtonLock(Me, True)


End Sub
Private Sub Grid_Head1()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 15

MSHFlexGrid1.TextMatrix(0, 0) = "SM's Employee Id"
MSHFlexGrid1.TextMatrix(0, 1) = "HR Employee Id"
MSHFlexGrid1.TextMatrix(0, 2) = "Employee Name"
MSHFlexGrid1.TextMatrix(0, 3) = "StoreMan Group"
MSHFlexGrid1.TextMatrix(0, 4) = "Direct Repotee"
MSHFlexGrid1.TextMatrix(0, 5) = "Connected DRC"
MSHFlexGrid1.TextMatrix(0, 6) = "Connected RAG"
MSHFlexGrid1.TextMatrix(0, 7) = "Email ID"
MSHFlexGrid1.TextMatrix(0, 8) = "Mobile No"
MSHFlexGrid1.TextMatrix(0, 9) = "Address"
MSHFlexGrid1.TextMatrix(0, 10) = "Base Location"
MSHFlexGrid1.TextMatrix(0, 11) = "Status"
MSHFlexGrid1.TextMatrix(0, 12) = "Self Reporting"
MSHFlexGrid1.TextMatrix(0, 13) = "Allocate Expense"
MSHFlexGrid1.TextMatrix(0, 14) = "HR Designation"

MSHFlexGrid1.ColWidth(0) = 1600
MSHFlexGrid1.ColWidth(1) = 1400
MSHFlexGrid1.ColWidth(2) = 3000
MSHFlexGrid1.ColWidth(3) = 1200
MSHFlexGrid1.ColWidth(4) = 1800
MSHFlexGrid1.ColWidth(5) = 1800
MSHFlexGrid1.ColWidth(6) = 1800
MSHFlexGrid1.ColWidth(7) = 1400
MSHFlexGrid1.ColWidth(8) = 1000
MSHFlexGrid1.ColWidth(9) = 2000
MSHFlexGrid1.ColWidth(10) = 1000
MSHFlexGrid1.ColWidth(11) = 1000
MSHFlexGrid1.ColWidth(12) = 1000
MSHFlexGrid1.ColWidth(13) = 1000
MSHFlexGrid1.ColWidth(14) = 1000

End Sub
Private Sub Grid_Head()

MSHFlexGrid2.Clear
MSHFlexGrid2.Rows = 2
MSHFlexGrid2.Cols = 7

MSHFlexGrid2.TextMatrix(0, 0) = "Location"
MSHFlexGrid2.TextMatrix(0, 1) = "CMP Id"
MSHFlexGrid2.TextMatrix(0, 2) = "CMP Name "
MSHFlexGrid2.TextMatrix(0, 3) = "CMP Address"
MSHFlexGrid2.TextMatrix(0, 4) = "Location ID"
MSHFlexGrid2.TextMatrix(0, 5) = "EmpCMPMap ID"
MSHFlexGrid2.TextMatrix(0, 6) = "SM_Prefix"

MSHFlexGrid2.ColWidth(0) = 1400
MSHFlexGrid2.ColWidth(1) = 1200
MSHFlexGrid2.ColWidth(2) = 2600
MSHFlexGrid2.ColWidth(3) = 6600
MSHFlexGrid2.ColWidth(4) = 0
MSHFlexGrid2.ColWidth(5) = 0
MSHFlexGrid2.ColWidth(6) = 0

End Sub
Private Sub Grid_Head2()

MSHFlexGrid3.Clear
MSHFlexGrid3.Rows = 2
MSHFlexGrid3.Cols = 9
MSHFlexGrid3.TextMatrix(0, 0) = "Asset Type"
MSHFlexGrid3.TextMatrix(0, 1) = "Asset Status"
MSHFlexGrid3.TextMatrix(0, 2) = "Asset Qty"
MSHFlexGrid3.TextMatrix(0, 3) = "Description"
MSHFlexGrid3.TextMatrix(0, 4) = "Verified On"
MSHFlexGrid3.TextMatrix(0, 5) = "EmpAssetMapID"
MSHFlexGrid3.TextMatrix(0, 6) = "Remarks"
MSHFlexGrid3.TextMatrix(0, 7) = "AMC Exp Date"
MSHFlexGrid3.TextMatrix(0, 8) = "AMC Amount"

MSHFlexGrid3.ColWidth(0) = 1800
MSHFlexGrid3.ColWidth(1) = 1600
MSHFlexGrid3.ColWidth(2) = 1000
MSHFlexGrid3.ColWidth(3) = 5400
MSHFlexGrid3.ColWidth(4) = 1000
MSHFlexGrid3.ColWidth(5) = 0
MSHFlexGrid3.ColWidth(6) = 2000
MSHFlexGrid3.ColWidth(7) = 1000
MSHFlexGrid3.ColWidth(8) = 1000

End Sub
Private Sub CmbAssetType_GotFocus()

tmp = CmbAssetType.Text
Call TableMst_AssetType("")
CmbAssetType.Clear
If RsMst_AssetType.RecordCount = 0 Then
   MsgBox "No Asset found"
   CmbAssetType.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_AssetType.RecordCount
    CmbAssetType.AddItem RsMst_AssetType!AssetType
    RsMst_AssetType.MoveNext
Next
CmbAssetType.Text = tmp

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
Private Sub CmbAssetStatus_GotFocus()

tmp = CmbAssetStatus.Text
Call TableMst_AssetStatus("")
CmbAssetStatus.Clear
If RsMst_AssetStatus.RecordCount = 0 Then
   MsgBox "No Asset Status found"
   CmbAssetStatus.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_AssetStatus.RecordCount
    CmbAssetStatus.AddItem RsMst_AssetStatus!AssetStatus
    RsMst_AssetStatus.MoveNext
Next
CmbAssetStatus.Text = tmp

End Sub
Private Sub CmbAssetType_LostFocus()

If CmbAssetType.Text = "" Then
   mAssetTypeId = Null
   Exit Sub
End If

RsMst_AssetType.MoveFirst
RsMst_AssetType.Find "AssetType = '" & CmbAssetType.Text & "'"

If RsMst_AssetType.EOF Then
   MsgBox "Invalid Asset  selection "
   CmbAssetType.SetFocus
   Exit Sub
End If
mAssetTypeId = RsMst_AssetType.Fields(0)

End Sub
Private Sub CmbAssetStatus_LostFocus()
If CmbAssetStatus.Text = "" Then
   mAssetStatusID = Null
   Exit Sub
End If

RsMst_AssetStatus.MoveFirst
RsMst_AssetStatus.Find "AssetStatus = '" & CmbAssetStatus.Text & "'"

If RsMst_AssetStatus.EOF Then
   MsgBox "Invalid Status selection "
   CmbAssetStatus.SetFocus
   Exit Sub
End If
mAssetStatusID = RsMst_AssetStatus.Fields(0)

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
   mCMPID = 0
   
End If

RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocationID.Text & "'"

If RsMst_Location.EOF Then
   MsgBox "Invalid Location  selection "
   CmbLocationID.SetFocus
   Exit Sub
End If
mLocationID = RsMst_Location!LocationID
mSM_Prefix = RsMst_Location!SM_Prefix

End Sub

Private Sub CmbDesignationID_GotFocus()
tmp = CmbDesignationID.Text
Call TableMst_Designation("")
CmbDesignationID.Clear
If RsMst_Designation.RecordCount = 0 Then
   MsgBox "No Designation found"
   CmbDesignationID.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Designation.RecordCount
    CmbDesignationID.AddItem RsMst_Designation!Designation
    RsMst_Designation.MoveNext
Next
CmbDesignationID.Text = tmp
End Sub
Private Sub CmbDesignationID_LostFocus()

If CmbDesignationID.Text = "" Then
   mDesignationID = Null
   Exit Sub
End If

If CmbDesignationID.Text <> tmp Then
   Call ClearTabs
   CmbAdditionalReporting.Text = ""
   CmbConnectedRAG.Text = ""
   mConnectedRAGID = Null
   mAdditionalReportingID = Null
End If


RsMst_Designation.MoveFirst
RsMst_Designation.Find "Designation = '" & CmbDesignationID.Text & "'"

If RsMst_Designation.EOF Then
   MsgBox "Invalid Designation Selection "
   CmbDesignationID.SetFocus
   Exit Sub
End If
mDesignationID = RsMst_Designation!DesignationID
mReportingDesignationID = RsMst_Designation!ReportingDesignationID
mSecReportingDesignationID = RsMst_Designation!SecReportingDesignationID

If tmp <> CmbDesignationID.Text Then
   CmbDirectReporting.Text = ""
End If

If mDesignationID = 6 Then
   Frame2.Enabled = True
   'Frame3.Enabled = True
   AssetVerifiedon.Value = Date
   Label25.Visible = True
   ChkSelfReporting.Visible = True
   ChkSelfReporting.Value = 0
   CmbAdditionalReporting.Enabled = True
   CmbConnectedRAG.Enabled = True
Else
   Label25.Visible = False
   ChkSelfReporting.Visible = False
   Frame2.Enabled = False
   'Frame3.Enabled = False
   ChkSelfReporting.Value = 0
   AssetVerifiedon.Value = Date
   CmbAdditionalReporting.Enabled = False
   CmbConnectedRAG.Enabled = False
End If


If mDesignationID = 1 Then
   ChkCMPAllocation.Enabled = True
Else
   ChkCMPAllocation.Enabled = False
   ChkCMPAllocation.Value = 1
End If


End Sub
Private Sub ClearTabs()
CmbLocationID.Text = ""
CmbCMPID.Text = ""
Call Grid_Head
CmbAssetType.Text = ""
CmbAssetStatus.Text = ""
TxtAssetDescription = ""
Call Grid_Head2
End Sub
Private Function Validate_DuplicateGrid() As Boolean
Dim Retval As Boolean
Retval = True
If MSHFlexGrid2.Rows = 2 Then
   Retval = True
Else
    For I = 1 To MSHFlexGrid2.Rows - 1
        If MSHFlexGrid2.TextMatrix(I, 1) <> "" Then
           If CmdAddToGrid.Caption = "Update Grid" Then
              If I <> MSHFlexGrid2.RowSel Then
                 If Val(MSHFlexGrid2.TextMatrix(I, 1)) = mCMPID Then
                    Retval = False
                 End If
              End If
           Else
              If Val(MSHFlexGrid2.TextMatrix(I, 1)) = mCMPID Then
                 Retval = False
              End If
           End If
        End If
    Next
End If
Validate_DuplicateGrid = Retval
End Function
Private Function Validate_DuplicateGridForAll() As Boolean
Dim Retval As Boolean
Retval = True
If MSHFlexGrid2.Rows = 2 Then
   Retval = True
Else
    For I = 1 To MSHFlexGrid2.Rows - 1
        If MSHFlexGrid2.TextMatrix(I, 1) <> "" Then
           If Val(MSHFlexGrid2.TextMatrix(I, 1)) = mCMPID Then
              Retval = False
           End If
        End If
    Next
End If
Validate_DuplicateGridForAll = Retval
End Function

Private Function CheckCMPMapping() As Boolean
Dim Retval As Boolean
Retval = True
TmpRs3.Find "CMPID = " & mCMPID & ""
If Not TmpRs3.EOF = True Then
   If TmpRs3!DesignationID = 6 Then
      Retval = False
   End If
   If TmpRs3!DesignationID = 7 Then
      If CmdAddToGrid.Caption = "Add To Grid" Then
         Retval = False
      End If
   End If
End If
CheckCMPMapping = Retval
End Function



Private Sub Form_Resize()
If Me.WindowState = vbMaximized Then
   SetScrollBars
End If
End Sub

Private Sub MSHAllocation_Click()
If SaveCmd.Enabled = False Then Exit Sub
Dim r As Variant
r = MSHAllocation.RowSel
If r < 1 Or r = MSHAllocation.Rows - 1 Then
   Call ClearAllocationFrame
   Exit Sub
End If
If MSHAllocation.TextMatrix(r, 1) = "" Then
   Exit Sub
End If
CmbAllocationType = MSHAllocation.TextMatrix(r, 1)
CmbAllocationTypeID = MSHAllocation.TextMatrix(r, 2)
mAllocationID = MSHAllocation.TextMatrix(r, 3)
mAllocationType = Left$(MSHAllocation.TextMatrix(r, 1), 1)
CmdAdd2GridAllocation.Caption = "Update Grid"
CmbAllocationType.Enabled = False
CmdDeleteFromGridAllocation.Enabled = True

End Sub

Private Sub MSHFlexGrid1_Click()
Dim x As Variant
Dim mCaId As Variant

x = MSHFlexGrid1.RowSel

mCaId = MSHFlexGrid1.TextMatrix(x, 0)

RsMst_Employee.MoveFirst

If RsMst_Employee.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
RsMst_Employee.MoveFirst
RsMst_Employee.Find "EmployeeID = " & Val(mCaId) & ""
If RsMst_Employee.EOF Then
   RsMst_Employee.MoveLast
End If
Frame0.Visible = False
Call Indata
End Sub

Private Sub MSHFlexGrid2_Click()
Dim I As Variant
I = MSHFlexGrid2.RowSel
If MSHFlexGrid2.TextMatrix(I, 0) = "" Then Exit Sub
CmbLocationID.Text = MSHFlexGrid2.TextMatrix(I, 0)
mCMPID = GetCMPID(MSHFlexGrid2.TextMatrix(I, 2))
CmbCMPID.Text = MSHFlexGrid2.TextMatrix(I, 2)
CmbLocationID.SetFocus
CmdAddToGrid.Caption = "Update Grid"
CmbCMPID.SetFocus
End Sub
Private Sub MSHFlexGrid3_Click()
Dim I As Variant
I = MSHFlexGrid3.RowSel
If MSHFlexGrid3.TextMatrix(I, 0) = "" Then Exit Sub
CmbAssetType.Text = MSHFlexGrid3.TextMatrix(I, 0)
CmbAssetStatus.Text = MSHFlexGrid3.TextMatrix(I, 1)
TxtAssetQty = Val(MSHFlexGrid3.TextMatrix(I, 2))
TxtAssetDescription = MSHFlexGrid3.TextMatrix(I, 3)
AssetVerifiedon.Value = MSHFlexGrid3.TextMatrix(I, 4)
TxtRemarks = MSHFlexGrid3.TextMatrix(I, 6)
TxtAMCExpDate.Value = IIf(MSHFlexGrid3.TextMatrix(I, 7) = "", Date, MSHFlexGrid3.TextMatrix(I, 7))
TxtAMCAmount = MSHFlexGrid3.TextMatrix(I, 8)
CmbAssetType.SetFocus
CmdAdd2GridAsset.Caption = "Update Grid"
End Sub

Private Sub Option2_Click()
Call Option1_Click
End Sub

Private Sub Option3_Click()
Call Option1_Click
End Sub

Private Sub Option5_Click()
Call Option1_Click
End Sub

Private Sub SaveCmd_Click()
If SaveCmd.Enabled = False Then Exit Sub
Dim strSplitString() As String
Dim EmpNo_, EmpName_ As String

If TxtEmpNo = "" Then
   MsgBox "Blank Employee No Not Allowed!!!"
   TxtEmpNo.SetFocus
   Exit Sub
End If

If TxtEmpName = "" Then
   MsgBox "Blank Employee Name Not Allowed!!!"
   TxtEmpName.SetFocus
   Exit Sub
End If

'If CmbDirectReporting.Text <> "" Or CmbAdditionalReporting.Text <> "" Then
'   If CmbDirectReporting.Text = CmbAdditionalReporting.Text Then
'      MsgBox "Direct Reportee And Connected To DRC Should Be Different!!!"
'      Exit Sub
'   End If
'End If

''If CmbDirectReporting.Text <> "" Or CmbConnectedRAG.Text <> "" Then
''   If CmbDirectReporting.Text = CmbConnectedRAG.Text Then
''      MsgBox "Direct Reportee  And Connected To RAG Should Be Different!!!"
''      Exit Sub
''   End If
''End If



If CmbDirectReporting.Text <> "" Then
   strSplitString = Split(CmbDirectReporting.Text, "~")
   EmpName_ = Trim(strSplitString(0))
   EmpNo_ = Trim(strSplitString(1))
   If TxtEmpNo = EmpNo_ Then
      MsgBox "Self Reporting Not Allowed!!!"
      CmbDirectReporting.SetFocus
   End If
Else
   MsgBox "Blank Direct Reportee not allowed!!!"
   CmbDirectReporting.SetFocus
   Exit Sub
End If

If CmbDesignationID.Text = "" Then
   MsgBox "Blank StoreMan Role Not Allowed!!!"
   CmbDesignationID.SetFocus
   Exit Sub
End If

'If LCase(Gen_UserRole) <> "power" Then
'   If Edit_Flg = "E" Then
'      If mDesignationID <> mOldDesignationID Then
'         MsgBox "Change in Designation not allowed!!!"
'         CmbDesignationID.SetFocus
'         Exit Sub
'      End If
'   End If
'End If

If CmbFlag.Text = "" Then
   MsgBox "Blank Status Not Allowed!!!"
   CmbFlag.SetFocus
   Exit Sub
End If

If CmbBaseLocationID.Text = "" Then
   MsgBox "Blank Base Location Not Allowed!!!"
   CmbBaseLocationID.SetFocus
   Exit Sub
End If

If MSHAllocation.Rows < 3 Then
   MsgBox "Allocation Details Missing."
   SSTab1.Tab = 1
   Exit Sub
End If

If LCase(CmbFlag.Text) = "de-active" Then
   If CheckHRProStatus("D") = False Then
      MsgBox "HR-Pro status for this employee is active, you can't de-activate!!"
      CmbFlag.SetFocus
      Exit Sub
   End If
End If
   
If LCase(CmbFlag.Text) = "active" Then
   If CheckHRProStatus("A") = True Then
      MsgBox "HR-Pro status for this employee is de-active or segment mapping changed, you can't activate!!"
      CmbFlag.SetFocus
      Exit Sub
   End If
End If
   
If Edit_Flg = "A" Then
   If RsMst_Employee.RecordCount = 0 Then
      tmp = 1
   Else
        RsMst_Employee.MoveFirst
        RsMst_Employee.Find "EmployeeNo = '" & TxtEmpNo & "'"
        If Not RsMst_Employee.EOF Then
           MsgBox "Duplicate Employee No not Allowed"
           TxtEmpNo.SetFocus
           Exit Sub
        End If
        tmp = "select max(EmployeeID) from Mst_Employee "
        Call TmpTable
        tmp = IIf(IsNull(TmpRs.Fields(0)), 1, TmpRs.Fields(0) + 1)
    End If
    RsMst_Employee.AddNew
    RsMst_Employee!EmployeeID = tmp
    TxtEmpID = tmp
Else
    RsMst_Employee.MoveFirst
    RsMst_Employee.Find "EmployeeNo = '" & TxtEmpNo & "'"
    If Not RsMst_Employee.EOF Then
       If RsMst_Employee!EmployeeID <> TxtEmpID Then
          MsgBox "Duplicate Employee No not Allowed"
          TxtEmpNo.SetFocus
          Exit Sub
       End If
    End If
   RsMst_Employee.MoveFirst
   RsMst_Employee.Find "EmployeeID = " & TxtEmpID & ""
End If
RsMst_Employee!EmployeeNo = TxtEmpNo
RsMst_Employee!EmployeeName = TxtEmpName
RsMst_Employee!DesignationID = mDesignationID
RsMst_Employee!SelfReporting = ChkSelfReporting.Value
RsMst_Employee!DirectReportingTo = mDirectReportingID
RsMst_Employee!ReportingToDRC = mAdditionalReportingID
RsMst_Employee!ConnectedToRAG = mConnectedRAGID
RsMst_Employee!EmployeeEmailID = TxtEmpEmailID
RsMst_Employee!MobileNo = TxtMobileNo
RsMst_Employee!Address = TxtAddress
RsMst_Employee!BaseLocationID = mBaseLocationID
RsMst_Employee!AllocateExpense = ChkAllocateExpense.Value
RsMst_Employee!CMPAllocation = ChkCMPAllocation.Value
'RsMst_Employee!LocationID = mLocationID
RsMst_Employee!Flag = mFlag

If IsNull(RsMst_Employee!CreatedBy) = True Or RsMst_Employee!CreatedBy = "" Then
   RsMst_Employee!CreatedBy = Gen_UserLoginID
   RsMst_Employee!CreatedDate = Now
Else
   RsMst_Employee!UpdatedBy = Gen_UserLoginID
   RsMst_Employee!UpdatedDate = Now
End If
RsMst_Employee.Update

If mDesignationID = 8 Then
   Call CheckQQSMapping
End If
Call TableMst_EmpCMPMap(" Where EmployeeID = " & RsMst_Employee!EmployeeID & "")

If MSHFlexGrid2.Rows > 2 Then
   'tmp = "Delete from Mst_EmpCMPMapping where EmployeeID = " & RsMst_Employee!EmployeeID & ""
   'Call TmpTable
   
   ''''''''''''''''''''Code added for Inserting Data into Log_EMPCMPMapping''''''''''''''''''''''''''''''''
   'tmp = "Delete from Mst_EmpCMPMapping where EmployeeID = " & RsMst_Employee!EmployeeID & ""
   'Call TmpTable
   
'   Call TableMst_EmpCMPMap(" Where EmployeeID = " & RsMst_Employee!EmployeeID & "")
   
   For j = 1 To RsMst_EmpCMPMap.RecordCount
       tmp = "Delete from Mst_EmpCMPMapping where EmpCMPMapID = " & RsMst_EmpCMPMap!EmpCMPMapID & ""
       Call TmpTable
       RsMst_EmpCMPMap.MoveNext
   Next
   RsMst_EmpCMPMap.Requery

   ''''''''''''''''''''Code added for Inserting Data into Log_EMPCMPMapping''''''''''''''''''''''''''''''''

   If RsMst_EmpCMPMap.RecordCount > 0 Then
      RsMst_EmpCMPMap.MoveFirst
   End If
 '  Call TableMst_EmpCMPMap(" Where EmployeeID = " & RsMst_Employee!EmployeeID & "")

   For I = 1 To MSHFlexGrid2.Rows - 2
       If MSHFlexGrid2.TextMatrix(I, 5) <> "" Then
          RsMst_EmpCMPMap.Find "EmpCMPMapID = " & MSHFlexGrid2.TextMatrix(I, 5)
          If RsMst_EmpCMPMap.EOF Then
             RsMst_EmpCMPMap.AddNew
             RsMst_EmpCMPMap!EmpCMPMapID = GetEmpCMPMapID()
          End If
       Else
          RsMst_EmpCMPMap.AddNew
          RsMst_EmpCMPMap!EmpCMPMapID = GetEmpCMPMapID()
       End If
       RsMst_EmpCMPMap!LocationID = MSHFlexGrid2.TextMatrix(I, 4)
       RsMst_EmpCMPMap!CMPID = MSHFlexGrid2.TextMatrix(I, 1)
       RsMst_EmpCMPMap!EmployeeID = TxtEmpID
       RsMst_EmpCMPMap!DesignationID = mDesignationID
       RsMst_EmpCMPMap!SM_Prefix = MSHFlexGrid2.TextMatrix(I, 6)
       If IsNull(RsMst_EmpCMPMap!CreatedBy) = True Or RsMst_EmpCMPMap!CreatedBy = "" Then
          RsMst_EmpCMPMap!CreatedBy = Gen_UserLoginID
          RsMst_EmpCMPMap!CreatedDate = Now
       Else
          RsMst_EmpCMPMap!UpdatedBy = Gen_UserLoginID
          RsMst_EmpCMPMap!UpdatedDate = Now
       End If
       RsMst_EmpCMPMap.Update
   Next
Else
   'tmp = "Delete from Mst_EmpCMPMapping where EmployeeID = " & RsMst_Employee!EmployeeID & ""
   'Call TmpTable
   ''''''''''''''''''''Code added for Inserting Data into Log_EMPCMPMapping''''''''''''''''''''''''''''''''
   'tmp = "Delete from Mst_EmpCMPMapping where EmployeeID = " & RsMst_Employee!EmployeeID & ""
   'Call TmpTable
   'Call TableMst_EmpCMPMap(" Where EmployeeID = " & RsMst_Employee!EmployeeID & "")
   
   For j = 1 To RsMst_EmpCMPMap.RecordCount
       tmp = "Delete from Mst_EmpCMPMapping where EmpCMPMapID = " & RsMst_EmpCMPMap!EmpCMPMapID & ""
       Call TmpTable
       RsMst_EmpCMPMap.MoveNext
   Next
   
   RsMst_EmpCMPMap.Requery
 
   ''''''''''''''''''''Code added for Inserting Data into Log_EMPCMPMapping''''''''''''''''''''''''''''''''

End If

If MSHFlexGrid3.Rows > 2 Then
   tmp = "Delete from Mst_EmpAssetMap where EmployeeID = " & RsMst_Employee!EmployeeID & ""
   Call TmpTable
   
   Call TableMst_EmpAssetMap(" Where EmployeeID = " & RsMst_Employee!EmployeeID & "")

   For I = 1 To MSHFlexGrid3.Rows - 2
       If MSHFlexGrid3.TextMatrix(I, 5) <> "" Then
          
          RsMst_EmpAssetMap.Find "EmpAssetMapID = " & MSHFlexGrid3.TextMatrix(I, 5)
          If RsMst_EmpAssetMap.EOF Then
             RsMst_EmpAssetMap.AddNew
             RsMst_EmpAssetMap!EmpAssetMapID = GetEmpAssetMapID()
          End If
       Else
          RsMst_EmpAssetMap.AddNew
          RsMst_EmpAssetMap!EmpAssetMapID = GetEmpAssetMapID()
       End If
       RsMst_EmpAssetMap!EmployeeID = TxtEmpID
       RsMst_EmpAssetMap!AssetID = GetAssetTypeID(MSHFlexGrid3.TextMatrix(I, 0))
       RsMst_EmpAssetMap!AssetStatusID = GetAssetStatusID(MSHFlexGrid3.TextMatrix(I, 1))
       RsMst_EmpAssetMap!AssetQty = Val(MSHFlexGrid3.TextMatrix(I, 2))
       RsMst_EmpAssetMap!Description = MSHFlexGrid3.TextMatrix(I, 3)
       RsMst_EmpAssetMap!VerifiedDate = MSHFlexGrid3.TextMatrix(I, 4)
       RsMst_EmpAssetMap!Remarks = MSHFlexGrid3.TextMatrix(I, 6)
       If MSHFlexGrid3.TextMatrix(I, 7) <> "" Then
          RsMst_EmpAssetMap!AMCExpDate = MSHFlexGrid3.TextMatrix(I, 7)
       End If
       RsMst_EmpAssetMap!AMCAmount = Val(MSHFlexGrid3.TextMatrix(I, 8))
       If IsNull(RsMst_EmpAssetMap!CreatedBy) = True Or RsMst_EmpAssetMap!CreatedBy = "" Then
          RsMst_EmpAssetMap!CreatedBy = Gen_UserLoginID
          RsMst_EmpAssetMap!CreatedDate = Now
       Else
          RsMst_EmpAssetMap!UpdatedBy = Gen_UserLoginID
          RsMst_EmpAssetMap!UpdatedDate = Now
       End If
       RsMst_EmpAssetMap.Update
   Next
Else
   tmp = "Delete from Mst_EmpAssetMap where EmployeeID = " & RsMst_Employee!EmployeeID & ""
   Call TmpTable
End If

Call SaveAllocationDetails
Call UpdateUsersStatus

Call GButtonLock(Me, True)
RsMst_Employee.Requery
RsMst_Employee.MoveFirst
RsMst_Employee.Find "EmployeeID = " & TxtEmpID & ""
If RsMst_Employee.EOF Then
   RsMst_Employee.MoveLast
End If
Call Indata

End Sub

Private Sub SaveAllocationDetails()
tmp = "Delete from Mst_EmployeeAllocationDetails Where EmployeeID=" & Val(TxtEmpID)
Call TmpTable

Call TableMst_EmployeeAllocationDetails(" Where EmployeeID=" & Val(TxtEmpID) & "")

For I = 1 To MSHAllocation.Rows - 2
    If MSHAllocation.TextMatrix(I, 1) <> "" Then
       RsMst_EmployeeAllocationDetails.AddNew
       RsMst_EmployeeAllocationDetails!EmpAllDetailID = GetEmpAllDetailID '("EmpAllDetailID", "Mst_EmployeeAllocationDetails")
       RsMst_EmployeeAllocationDetails!EmployeeID = Val(TxtEmpID)
       RsMst_EmployeeAllocationDetails!AllocationType = Left$(MSHAllocation.TextMatrix(I, 1), 1)
       RsMst_EmployeeAllocationDetails!AllocationID = Val(MSHAllocation.TextMatrix(I, 3))
       RsMst_EmployeeAllocationDetails!G_UID = GetGUID
       
       RsMst_EmployeeAllocationDetails.Update
       'RsMst_AcademicDetails.Requery
    End If
Next
End Sub

Private Function GetEmpAllDetailID() As Double
Dim Retval As Double
tmp = "select max(EmpAllDetailID) from Mst_EmployeeAllocationDetails"
Call TmpTable
tmp = IIf(IsNull(TmpRs.Fields(0)), 1, TmpRs.Fields(0) + 1)
GetEmpAllDetailID = tmp
End Function


Private Function GetEmpCMPMapID() As Double
Dim Retval As Double
tmp = "select max(EmpCMPMapID) from Mst_EmpCMPMapping"
Call TmpTable
tmp = IIf(IsNull(TmpRs.Fields(0)), 1, TmpRs.Fields(0) + 1)
GetEmpCMPMapID = tmp
End Function
Private Function GetEmpAssetMapID() As Double
Dim Retval As Double
tmp = "select max(EmpAssetMapID) from Mst_EmpAssetMap"
Call TmpTable
tmp = IIf(IsNull(TmpRs.Fields(0)), 1, TmpRs.Fields(0) + 1)
GetEmpAssetMapID = tmp
End Function
Public Sub Indata()


Frame2.Enabled = False
Frame3.Enabled = False
Frame5.Enabled = False

CmbDesignationID.Visible = False
CmbDirectReporting.Visible = False
CmbAdditionalReporting.Visible = False
CmbConnectedRAG.Visible = False
CmbFlag.Visible = False
CmbBaseLocationID.Visible = False


TxtDesignationID.Visible = True
TxtDirectReporting.Visible = True
'TxtAdditionalReporting.Visible = True
TxtConnectedRAG.Visible = True
TxtFlag.Visible = True
TxtBaseLocationID.Visible = True

CmdDesignationID.Visible = False
CmdLocationID.Visible = False
CmdCMPID.Visible = False
CmdAssetType.Visible = False
cmdAssetStatus.Visible = False
CmdBaseLocationID.Visible = False
ChkSelfReporting.Visible = False

TxtEmpNo.Locked = True
TxtEmpName.Locked = True
TxtEmpEmailID.Locked = True
TxtMobileNo.Locked = True
TxtAddress.Locked = True
ChkAllocateExpense.Enabled = False
ChkCMPAllocation.Enabled = False

TxtEmpID = ""
TxtHrDesignation = ""
TxtEmpNo = ""
TxtEmpName = ""
TxtEmpEmailID = ""
TxtMobileNo = ""
TxtAddress = ""
TxtDesignationID = ""
TxtDirectReporting = ""
'TxtAdditionalReporting = ""
TxtConnectedRAG = ""
TxtFlag = ""
TxtBaseLocationID = ""

mEmployeeID = RsMst_Employee!EmployeeID
TxtEmpID = RsMst_Employee!EmployeeID
TxtEmpNo = RsMst_Employee!EmployeeNo
TxtEmpName = RsMst_Employee!EmployeeName
TxtEmpEmailID = IIf(IsNull(RsMst_Employee!EmployeeEmailID), "", RsMst_Employee!EmployeeEmailID)
TxtMobileNo = IIf(IsNull(RsMst_Employee!MobileNo), "", RsMst_Employee!MobileNo)
TxtAddress = IIf(IsNull(RsMst_Employee!Address), "", RsMst_Employee!Address)
ChkAllocateExpense.Value = IIf(IsNull(RsMst_Employee!AllocateExpense), 0, RsMst_Employee!AllocateExpense)
ChkCMPAllocation.Value = IIf(IsNull(RsMst_Employee!CMPAllocation), 0, RsMst_Employee!CMPAllocation)
mOldDesignationID = RsMst_Employee!DesignationID
mDesignationID = RsMst_Employee!DesignationID
TxtDesignationID = GetDesignaion(mDesignationID)
mReportingDesignationID = GetReportingDesignation(mDesignationID)
mSecReportingDesignationID = GetSecReportingDesignation(mDesignationID)

mDirectReportingID = IIf(IsNull(RsMst_Employee!DirectReportingTo), "", RsMst_Employee!DirectReportingTo)
If mDirectReportingID <> "" Then
   TxtDirectReporting = GetEmployeeNonName(mDirectReportingID)
Else
   mDirectReportingID = Null
End If
mPreDirectReportingID = mDirectReportingID
'mAdditionalReportingID = IIf(IsNull(RsMst_Employee!ReportingToDRC), "", RsMst_Employee!ReportingToDRC)
'If mAdditionalReportingID <> "" Then
'   TxtAdditionalReporting = GetEmployeeNonName(mAdditionalReportingID)
'Else
'   mAdditionalReportingID = Null
'End If

mConnectedRAGID = RsMst_Employee!ConnectedToRAG
If mConnectedRAGID <> 0 Then
   TxtConnectedRAG = GetEmployeeNonName(mConnectedRAGID)
Else
   mConnectedRAGID = Null
End If


mFlag = RsMst_Employee!Flag


If mFlag = "Y" Then
   TxtFlag = "Active"
Else
   TxtFlag = "De-Active"
End If




mBaseLocationID = IIf(IsNull(RsMst_Employee!BaseLocationID), 0, RsMst_Employee!BaseLocationID)

If mBaseLocationID <> 0 Then
   TxtBaseLocationID = GetLocationName(mBaseLocationID)
Else
   mBaseLocationID = 0
End If


If RsMst_Employee!SelfReporting = 1 Then
   Label25.Visible = True
   ChkSelfReporting.Visible = True
Else
   Label25.Visible = False
   ChkSelfReporting.Visible = False
End If
ChkSelfReporting.Value = IIf(IsNull(RsMst_Employee!SelfReporting), 0, RsMst_Employee!SelfReporting)
ChkSelfReporting.Enabled = False

Call ClearTabs
TxtCreated = IIf(IsNull(RsMst_Employee!CreatedBy), "", RsMst_Employee!CreatedBy) & " :- " & IIf(IsNull(RsMst_Employee!CreatedDate), "", RsMst_Employee!CreatedDate)
TxtUpdated = IIf(IsNull(RsMst_Employee!UpdatedBy), "", RsMst_Employee!UpdatedBy) & " :- " & IIf(IsNull(RsMst_Employee!UpdatedDate), "", RsMst_Employee!UpdatedDate)

TxtHrDesignation = GetHRDesignationFromHRPro(TxtEmpNo)
TxtHRDirectReportee = GetHRDirectReporteeFromHRPro(TxtEmpNo, "D")
mHRLocationID = GetHRDirectReporteeFromHRPro(TxtEmpNo, "L")

Call Grid_Head

'Call TableMst_EmpCMPMap(" Where EmployeeID = " & RsMst_Employee!EmployeeID & "")

tmp = "Select  ML.LocationName,MC.CMPId,MC.CMPName,MC.Address,ML.LocationID,ECM.EmpCMPMapID,ML.SM_Prefix " & _
      " From Mst_EmpCMPMapping ECM  " & _
      " Inner Join Mst_Location ML On ECM.LocationID = ML.LocationID  " & _
      " Inner Join Mst_CMP MC On ECM.CMPID = MC.CMPID " & _
      " Where MC.CloseDate Is Null ANd ECM.EmployeeID = " & RsMst_Employee!EmployeeID & "  "

Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   For I = 1 To TmpRs1.RecordCount
       MSHFlexGrid2.TextMatrix(I, 0) = TmpRs1!LocationName 'GetLocationName(TmpRs1!LocationID)
       MSHFlexGrid2.TextMatrix(I, 1) = TmpRs1!CMPID
       MSHFlexGrid2.TextMatrix(I, 2) = TmpRs1!CMPName 'GetCMPName(TmpRs1!CMPID)
       MSHFlexGrid2.TextMatrix(I, 3) = IIf(IsNull(TmpRs1!Address), "", TmpRs1!Address) 'GetCMPAddress(TmpRs1!CMPID, "A")
       MSHFlexGrid2.TextMatrix(I, 4) = TmpRs1!LocationID
       MSHFlexGrid2.TextMatrix(I, 5) = TmpRs1!EmpCMPMapID
       MSHFlexGrid2.TextMatrix(I, 6) = TmpRs1!SM_Prefix
       TmpRs1.MoveNext
       MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
   Next
End If

Call Grid_Head2
'Call TableMst_EmpAssetMap(" Where EmployeeID = " & RsMst_Employee!EmployeeID & "")
'
'If RsMst_EmpAssetMap.RecordCount > 0 Then
'   For I = 1 To RsMst_EmpAssetMap.RecordCount
'       MSHFlexGrid3.TextMatrix(I, 0) = GetAssetType(RsMst_EmpAssetMap!AssetID)
'       MSHFlexGrid3.TextMatrix(I, 1) = GetAssetStatus(RsMst_EmpAssetMap!AssetStatusID)
'       MSHFlexGrid3.TextMatrix(I, 2) = RsMst_EmpAssetMap!AssetQty
'       MSHFlexGrid3.TextMatrix(I, 3) = RsMst_EmpAssetMap!Description
'       MSHFlexGrid3.TextMatrix(I, 4) = RsMst_EmpAssetMap!VerifiedDate
'       MSHFlexGrid3.TextMatrix(I, 5) = RsMst_EmpAssetMap!EmpAssetMapID
'       MSHFlexGrid3.TextMatrix(I, 6) = IIf(IsNull(RsMst_EmpAssetMap!Remarks), "", RsMst_EmpAssetMap!Remarks)
'       MSHFlexGrid3.TextMatrix(I, 7) = IIf(IsNull(RsMst_EmpAssetMap!AMCExpDate), "", RsMst_EmpAssetMap!AMCExpDate)
'       MSHFlexGrid3.TextMatrix(I, 8) = IIf(IsNull(RsMst_EmpAssetMap!AMCAmount), 0, RsMst_EmpAssetMap!AMCAmount)
'       RsMst_EmpAssetMap.MoveNext
'       MSHFlexGrid3.Rows = MSHFlexGrid3.Rows + 1
'   Next
'End If
Call InDataAllocation

SSTab1.Tab = 0

Call GButtonLock(Me, True)
End Sub

Private Sub InDataAllocation()
Call Grid_HeadAllocation

Call TableMst_EmployeeAllocationDetails(" Where EmployeeID=" & Val(TxtEmpID) & "")
MSHAllocation.Rows = RsMst_EmployeeAllocationDetails.RecordCount + 2
For r = 1 To RsMst_EmployeeAllocationDetails.RecordCount
    If LCase(RsMst_EmployeeAllocationDetails!AllocationType) = "s" Then
        tmp = "Select StateName from Mst_State Where StateID = " & RsMst_EmployeeAllocationDetails!AllocationID
        Call TmpTable
        If TmpRs.RecordCount > 0 Then
            MSHAllocation.TextMatrix(r, 1) = "State"
            MSHAllocation.TextMatrix(r, 2) = IIf(IsNull(TmpRs!StateName), "", TmpRs!StateName)
            MSHAllocation.TextMatrix(r, 3) = IIf(IsNull(RsMst_EmployeeAllocationDetails!AllocationID), "", RsMst_EmployeeAllocationDetails!AllocationID)
        End If
    ElseIf LCase(RsMst_EmployeeAllocationDetails!AllocationType) = "r" Then
        tmp = "Select Region from Mst_Region Where RegionID = " & RsMst_EmployeeAllocationDetails!AllocationID
        Call TmpTable
        If TmpRs.RecordCount > 0 Then
            MSHAllocation.TextMatrix(r, 1) = "Region"
            MSHAllocation.TextMatrix(r, 2) = IIf(IsNull(TmpRs!Region), "", TmpRs!Region)
            MSHAllocation.TextMatrix(r, 3) = IIf(IsNull(RsMst_EmployeeAllocationDetails!AllocationID), "", RsMst_EmployeeAllocationDetails!AllocationID)
        End If
    ElseIf LCase(RsMst_EmployeeAllocationDetails!AllocationType) = "l" Then
        tmp = "Select LocationName from Mst_Location Where LocationID = " & RsMst_EmployeeAllocationDetails!AllocationID
        Call TmpTable
        If TmpRs.RecordCount > 0 Then
            MSHAllocation.TextMatrix(r, 1) = "Location"
            MSHAllocation.TextMatrix(r, 2) = IIf(IsNull(TmpRs!LocationName), "", TmpRs!LocationName)
            MSHAllocation.TextMatrix(r, 3) = IIf(IsNull(RsMst_EmployeeAllocationDetails!AllocationID), "", RsMst_EmployeeAllocationDetails!AllocationID)
        End If
    End If
    RsMst_EmployeeAllocationDetails.MoveNext
Next
End Sub


Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_Employee.MoveNext
If RsMst_Employee.EOF Then
   RsMst_Employee.MoveLast
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
RsMst_Employee.MovePrevious
If RsMst_Employee.BOF Then
   RsMst_Employee.MoveFirst
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
RsMst_Employee.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsMst_Employee.MoveLast
Call Indata
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = Fasle
End Sub

Private Sub SearchCmd_Click()
Dim x As Integer
If SearchCmd.Caption = "Cancel" Then
   If RsMst_Employee.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   If RsMst_Employee.EOF Then
      RsMst_Employee.MoveFirst
   End If
   Call Indata
   Exit Sub
End If

Frame0.Height = Frame4.Height - 500
Frame0.Left = Frame1.Left
Frame0.Top = Frame1.Top
Frame0.Width = Frame4.Width - 400
Frame0.Visible = True

MSHFlexGrid1.Width = Frame0.Width - 300
MSHFlexGrid1.Height = Frame0.Height - 1200

''===========comment on 7 july 2012 for not clearing search grid=============
'MSHFlexGrid1.Clear
'Call Grid_Head1
'TxtSearchEmployee = ""
'CmbSearchStatus.Text = ""
'Option4.Value = True
'Label22.Caption = ""
''===========comment on 7 july 2012 for not clearing search grid=============

End Sub

Private Sub TxtAddress_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "TxtAddress")
End If
End Sub

Private Sub TxtAMCAmount_LostFocus()
If TxtAMCAmount = "" Then Exit Sub
If IsNumeric(TxtAMCAmount) = False Or Val(TxtAMCAmount) < 0 Then
   MsgBox "Invalid number format!!!"
   TxtAMCAmount.SetFocus
   Exit Sub
End If
TxtAMCAmount = Format(TxtAMCAmount, "#########0.00")

End Sub

Private Sub TxtAssetDescription_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "TxtAssetDescription")
End If
End Sub

Private Sub TxtAssetQty_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "TxtAssetQty")
End If
End Sub

Private Sub TxtAssetQty_LostFocus()
If TxtAssetQty = "" Then Exit Sub
If IsNumeric(TxtAssetQty) = False Or Val(TxtAssetQty) < 0 Then
   MsgBox "Invalid number format!!!"
   TxtAssetQty.SetFocus
   Exit Sub
End If
End Sub
Private Sub Option1_Click()
Label19.Visible = True
TxtSearchEmployee.Visible = True
TxtSearchEmployee = ""
'CmbSearchStatus.Visible = False
CmbSearchStatus.Text = ""
CmdGo.Visible = True
If Option1.Value = True Then
   Label19.Caption = "Search By EHOP Employee ID"
ElseIf Option2.Value = True Then
   Label19.Caption = "Search By Designation"
ElseIf Option5.Value = True Then
   Label19.Caption = "Search By Employee Name"
End If
End Sub

Private Sub TxtEmpEmailID_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "TxtEmpEmailID")
End If
End Sub

Private Sub TxtEmpId_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "TxtEmpID")
End If
End Sub

Private Sub TxtEmpName_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "TxtEmpName")
End If
End Sub

Private Sub TxtEmpNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "TxtEmpNo")
End If
End Sub

Private Sub TxtMobileNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "TxtMobileNo")
End If
End Sub

Private Sub TxtSearchEmployee_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "TxtSearchEmployee")
End If
End Sub
Private Sub Form_Activate()
If Me.WindowState = vbMaximized Then
   SetScrollBars
End If
End Sub
Private Sub SetScrollBars()
'    Frame1.Width = 10000
'    Frame1.Height = 20000
VScroll1.Top = 0
HScroll1.Left = 0
VScroll1.Left = Me.Width - (VScroll1.Width + 120)
HScroll1.Top = Me.Height - ((HScroll1.Height * 2) + 250)
VScroll1.Height = Me.Height - 500
HScroll1.Width = Me.Width - 500
VScroll1.Min = 0
HScroll1.Min = 0
VScroll1.Max = Me.ScaleHeight - Frame4.Height
VScroll1.LargeChange = Me.ScaleHeight
VScroll1.SmallChange = Me.ScaleHeight / 5
HScroll1.Max = Me.ScaleWidth - Frame4.Width
HScroll1.LargeChange = Me.ScaleWidth
HScroll1.SmallChange = Me.ScaleWidth / 5
If Frame4.Height >= Me.Height Then VScroll1.Enabled = True
If Frame4.Height < Me.Height Then VScroll1.Enabled = False
If Frame4.Width >= Me.Width Then HScroll1.Enabled = True
If Frame4.Width < Me.Width Then HScroll1.Enabled = False
End Sub
Private Sub VScroll1_Change()

   Dim l As Double
   Dim a As Double
   Dim x As Double
   
   x = VScroll1.Value
   a = Me.Height - Frame4.Height
   l = (a * x) / 1000
   Frame4.Top = -l

'Frame3.Top = VScroll1.Value
End Sub

Private Sub VScroll1_Scroll()

   Dim l As Double
   Dim a As Double
   Dim x As Double
   
   x = VScroll1.Value
   a = Me.Height - Frame4.Height
   l = (a * x) / 1000
   Frame4.Top = -l

'Frame3.Top = VScroll1.Value
End Sub

Private Sub HScroll1_Change()
   Dim l As Double
   Dim a As Double
   Dim x As Double
   
   x = HScroll1.Value
   a = Me.Width - Frame4.Width
   l = (a * x) / 1000
   Frame4.Left = -l
End Sub

Private Sub HScroll1_Scroll()
   Dim l As Double
   Dim a As Double
   Dim x As Double
   
   x = HScroll1.Value
   a = Me.Width - Frame4.Width
   l = (a * x) / 1000
   Frame4.Left = -l

End Sub

'Private Function CheckHRProStatus(mFlag_ As String) As Boolean
'Dim Retval As Boolean
'Retval = False
'
''===== if emp is active & department in CMP,CM & CMP, CMG
'If mFlag_ = "D" Then
'   tmp = "Select * from Mst_Employee Where EmpStatusID not in (2,4) And DepartmentID in (13,33,38) And EmployeeNo='" & Trim(TxtEmpNo) & "'"
'ElseIf mFlag_ = "A" Then
'   tmp = "Select * from Mst_Employee Where EmployeeNo='" & Trim(TxtEmpNo) & "' And DepartmentID Not in (13,33,38)" ''EmpStatusID in (2,4)
'End If
'Call TmpTableHRPro
'
'If TmpRsHRPro.RecordCount > 0 Then
'   Retval = True
'End If
'CheckHRProStatus = Retval
'End Function

Private Function CheckHRProStatus(mFlag_ As String) As Boolean
Dim Retval As Boolean
Retval = False

'===== if emp is active & department segment mapping to CMP
'If mFlag_ = "D" Then
'   tmp = "Select * from Mst_Employee Where EmpStatusID not in (2,4) And DepartmentID in (3,33,36) And EmployeeNo='" & Trim(TxtEmpNo) & "'"
'ElseIf mFlag_ = "A" Then
'   tmp = "Select * from Mst_Employee Where EmployeeNo='" & Trim(TxtEmpNo) & "' And DepartmentID not in (3,33,36)" ''EmpStatusID in (2,4) And "
'End If

'SegmentID = 1 CMP

If mFlag_ = "D" Then
   'tmp = "Select * from Mst_Employee Where EmpStatusID not in (2,4) And DepartmentID in (3,33,36) And EmployeeNo='" & Trim(TxtEmpNo) & "'"

'   tmp = " Select ME.EmployeeNo,ME.FullName,a.segmentID from Mst_Employee ME"
'   tmp = tmp & " Inner Join Mst_Department MD On ME.DepartmentID = MD.DepartmentID"
'   tmp = tmp & " Inner Join Mst_DeptSegmentMappingDetail a On MD.DepartmentID = a.DepartmentID"
'   tmp = tmp & " Where  ME.EmpStatusID in (2,4) And ME.EmployeeNo = '" & Trim(TxtEmpNo) & "' and a.SegmentID = 1"


   tmp = " Select ME.EmployeeNo,ME.FullName,a.BusinessSegmentID from Mst_Employee ME"
   tmp = tmp & " Inner Join Txn_EmpSegmentMapping MD On ME.EmployeeID = MD.EmployeeID"
   tmp = tmp & " Inner Join Txn_EmpSegmentMappingDetail a On MD.EmpSegMapID = a.EmpSegMapID"
   tmp = tmp & " Where ME.EmpStatusID in (2,4) And ME.EmployeeNo = '" & Trim(TxtEmpNo) & "' and a.BusinessSegmentID = 1"


ElseIf mFlag_ = "A" Then
'   tmp = "Select * from Mst_Employee Where EmployeeNo='" & Trim(TxtEmpNo) & "' And DepartmentID not in (3,33,36)" ''EmpStatusID in (2,4) And "

'    tmp = " Select ME.EmployeeNo,ME.FullName,a.segmentID from Mst_Employee ME"
'    tmp = tmp & " Inner Join Mst_Department MD On ME.DepartmentID = MD.DepartmentID"
'    tmp = tmp & " Inner Join Mst_DeptSegmentMappingDetail a On MD.DepartmentID = a.DepartmentID"
'    tmp = tmp & " Where ME.EmpStatusID not in (2,4) And ME.EmployeeNo = '" & Trim(TxtEmpNo) & "' and a.SegmentID = 1"


   tmp = " Select ME.EmployeeNo,ME.FullName,a.BusinessSegmentID from Mst_Employee ME"
   tmp = tmp & " Inner Join Txn_EmpSegmentMapping MD On ME.EmployeeID = MD.EmployeeID"
   tmp = tmp & " Inner Join Txn_EmpSegmentMappingDetail a On MD.EmpSegMapID = a.EmpSegMapID"
   tmp = tmp & " Where ME.EmpStatusID not in (2,4) And ME.EmployeeNo = '" & Trim(TxtEmpNo) & "' and a.BusinessSegmentID = 1"

End If

Call TmpTableHRPro

If TmpRsHRPro.RecordCount = 0 Then 'If TmpRsHR.RecordCount > 0 Then
   Retval = True
End If

CheckHRProStatus = Retval

End Function

Private Sub UpdateUsersStatus()
tmp = "Update Mst_Users Set Flag='" & mFlag & "' Where EmployeeNo='" & TxtEmpNo & "'"
Call TmpTable
End Sub

Public Sub CheckQQSMapping()

tmp = " Select QQSMAPID from Txn_QQSGodownMapping where EmployeeID=" & TxtEmpID.Text & " "
Call Tmp1Table
If TmpRs1.RecordCount > 0 Then
   If mPreDirectReportingID <> mDirectReportingID Then
      tmp = " Insert into Log_Txn_QQSGodownMappingDetail" & _
            " (ID," & _
            "  QQSMAPID," & _
            "  GodownID," & _
            "  UpdatedBy)" & _
            " ( Select Isnull((Select Max(ID) from Log_Txn_QQSGodownMappingDetail)+1,1) , QQSMAPID,GodownID,'" & Gen_UserLoginID & "'" & _
            " From Txn_QQSGodownMappingDetail Where QQSMapID= " & Val(TmpRs1!QQSMapID) & " )"

      Call Tmp5Table
       
      tmp = " Delete From Txn_QQSGodownMappingDetail where QQSMAPID=" & TmpRs1!QQSMapID & ""
      Call Tmp2Table
      tmp = " Update Txn_QQSGodownMapping Set FreeFlag='Y' Where QQSMAPID=" & TmpRs1!QQSMapID & ""
      Call Tmp3Table
   End If
End If
End Sub
