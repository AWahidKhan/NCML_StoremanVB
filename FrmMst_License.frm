VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmMst_License 
   Caption         =   "License Master"
   ClientHeight    =   6015
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10410
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6015
   ScaleWidth      =   10410
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   2895
      Left            =   -720
      TabIndex        =   38
      Top             =   9120
      Visible         =   0   'False
      Width           =   15135
      Begin VB.ComboBox CmbSEmployeeID 
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
         ItemData        =   "FrmMst_License.frx":0000
         Left            =   4080
         List            =   "FrmMst_License.frx":0002
         Sorted          =   -1  'True
         TabIndex        =   71
         Top             =   1630
         Width           =   5145
      End
      Begin VB.ComboBox CmbSLicenseType 
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
         ItemData        =   "FrmMst_License.frx":0004
         Left            =   120
         List            =   "FrmMst_License.frx":0006
         Sorted          =   -1  'True
         TabIndex        =   69
         Top             =   1630
         Width           =   3950
      End
      Begin VB.OptionButton OptIssuingAutho 
         Caption         =   "Issuing Authority"
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
         Left            =   4560
         TabIndex        =   49
         Top             =   240
         Width           =   1830
      End
      Begin VB.OptionButton OptLicenseNo 
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   840
         TabIndex        =   48
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton OptLicenseDate 
         Caption         =   "License Date"
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
         Left            =   6480
         TabIndex        =   47
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton OptLicExpDate 
         Caption         =   "License Expirey Date"
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
         Left            =   8160
         TabIndex        =   46
         Top             =   240
         Width           =   2295
      End
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   350
         Left            =   13560
         TabIndex        =   45
         Top             =   960
         Width           =   1125
      End
      Begin VB.OptionButton OptLicDesc 
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2280
         TabIndex        =   44
         Top             =   240
         Width           =   2070
      End
      Begin VB.TextBox TxtSearchLicense 
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
         Left            =   4140
         TabIndex        =   43
         Top             =   960
         Width           =   5055
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
         TabIndex        =   42
         Top             =   240
         Value           =   -1  'True
         Width           =   585
      End
      Begin VB.ComboBox CmbSearchLocation 
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
         ItemData        =   "FrmMst_License.frx":0008
         Left            =   9240
         List            =   "FrmMst_License.frx":000A
         Sorted          =   -1  'True
         TabIndex        =   41
         Top             =   960
         Width           =   2925
      End
      Begin VB.CommandButton CmdGo 
         Caption         =   "Go"
         Height          =   350
         Left            =   12360
         TabIndex        =   40
         Top             =   960
         Width           =   1125
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1215
         Left            =   120
         TabIndex        =   39
         Top             =   2040
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   2143
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSComCtl2.DTPicker DTPLicFromDate 
         Height          =   350
         Left            =   120
         TabIndex        =   50
         Top             =   960
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   609
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   48496641
         CurrentDate     =   40012
      End
      Begin MSComCtl2.DTPicker DTPLicToDate 
         Height          =   350
         Left            =   2130
         TabIndex        =   51
         Top             =   960
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   609
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   48496641
         CurrentDate     =   40012
      End
      Begin VB.Label Label13 
         Caption         =   "Name of person holding License Copy"
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
         Left            =   4905
         TabIndex        =   72
         Top             =   1380
         Width           =   3495
      End
      Begin VB.Label Label12 
         Caption         =   "License Type"
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
         Left            =   1368
         TabIndex        =   70
         Top             =   1380
         Width           =   1455
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
         Left            =   14400
         TabIndex        =   58
         Top             =   240
         Width           =   75
      End
      Begin VB.Label LblSearchLicense 
         Caption         =   "Search By License No "
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
         Left            =   5040
         TabIndex        =   57
         Top             =   600
         Width           =   3855
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
         TabIndex        =   56
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
         Left            =   11520
         TabIndex        =   55
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label24 
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
         Height          =   375
         Left            =   10320
         TabIndex        =   54
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblWHRDateFrom 
         Caption         =   "From Date"
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
         Left            =   570
         TabIndex        =   53
         Top             =   600
         Width           =   1035
      End
      Begin VB.Label LblWHRToDate 
         Caption         =   "To Date"
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
         Left            =   2685
         TabIndex        =   52
         Top             =   600
         Width           =   825
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1200
      Left            =   120
      TabIndex        =   31
      Top             =   7690
      Width           =   14925
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New"
         Height          =   400
         Left            =   100
         TabIndex        =   0
         Top             =   240
         Width           =   2930
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "Search"
         Height          =   400
         Left            =   11775
         TabIndex        =   20
         Top             =   240
         Width           =   2930
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   400
         Left            =   8850
         TabIndex        =   19
         Top             =   240
         Width           =   2930
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   400
         Left            =   5925
         TabIndex        =   18
         Top             =   240
         Width           =   2930
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   400
         Left            =   3030
         TabIndex        =   17
         Top             =   240
         Width           =   2900
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   400
         Left            =   105
         TabIndex        =   21
         Top             =   645
         Width           =   2930
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   400
         Left            =   3030
         TabIndex        =   22
         Top             =   645
         Width           =   2900
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   400
         Left            =   5925
         TabIndex        =   23
         Top             =   645
         Width           =   2930
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   400
         Left            =   8850
         TabIndex        =   24
         Top             =   645
         Width           =   2930
      End
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   400
         Left            =   11775
         TabIndex        =   25
         Top             =   645
         Width           =   2930
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2985
      Left            =   120
      TabIndex        =   27
      Top             =   0
      Width           =   14925
      Begin VB.TextBox TxtCapacity 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   12480
         MaxLength       =   100
         TabIndex        =   7
         Top             =   1800
         Width           =   2280
      End
      Begin VB.TextBox TxtIssuingAuthority 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   2040
         MaxLength       =   50
         TabIndex        =   6
         Top             =   1800
         Width           =   7320
      End
      Begin MSComCtl2.DTPicker TxtLicenseDate 
         Height          =   375
         Left            =   10200
         TabIndex        =   3
         Top             =   660
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   48496641
         CurrentDate     =   39889
      End
      Begin MSComCtl2.DTPicker TxtLicenseExpiryDate 
         Height          =   375
         Left            =   13440
         TabIndex        =   4
         Top             =   660
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   48496641
         CurrentDate     =   39889
      End
      Begin VB.TextBox TxtLicenseID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   360
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   240
         Width           =   1920
      End
      Begin VB.TextBox TxtLicenseDesc 
         Appearance      =   0  'Flat
         Height          =   600
         Left            =   2040
         MaxLength       =   100
         TabIndex        =   5
         Top             =   1110
         Width           =   12720
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
         Height          =   300
         Left            =   14370
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   1200
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox TxtLicenseNo 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   2040
         MaxLength       =   100
         TabIndex        =   2
         Top             =   675
         Width           =   6840
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
         Left            =   10200
         Sorted          =   -1  'True
         TabIndex        =   26
         Top             =   1155
         Visible         =   0   'False
         Width           =   4080
      End
      Begin VB.TextBox TxtLocationID 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   10200
         Locked          =   -1  'True
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   1155
         Visible         =   0   'False
         Width           =   4560
      End
      Begin VB.ComboBox CmbEmployeeID 
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
         TabIndex        =   8
         Top             =   2400
         Width           =   12720
      End
      Begin VB.TextBox TxtEmployeeID 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   2400
         Width           =   12720
      End
      Begin VB.ComboBox CmbLicenseType 
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
         Left            =   5160
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   9600
      End
      Begin VB.TextBox TxtLicenseType 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   240
         Width           =   9600
      End
      Begin VB.Label Label11 
         Caption         =   "Capacity for which license issued"
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
         Left            =   9480
         TabIndex        =   73
         Top             =   1890
         Width           =   2880
      End
      Begin VB.Label Label8 
         Caption         =   "Name of person holding License Copy"
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
         Left            =   120
         TabIndex        =   62
         Top             =   2400
         Width           =   1920
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "License Type"
         Height          =   195
         Left            =   4080
         TabIndex        =   60
         Top             =   300
         Width           =   960
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "License Expiry Date"
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
         Left            =   11640
         TabIndex        =   35
         Top             =   750
         Width           =   1710
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "License Date"
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
         Left            =   9000
         TabIndex        =   34
         Top             =   750
         Width           =   1140
      End
      Begin VB.Label Label2 
         Caption         =   "License Issuing Autority"
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
         Left            =   120
         TabIndex        =   33
         Top             =   1830
         Width           =   1800
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "License Description"
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
         Left            =   120
         TabIndex        =   32
         Top             =   1313
         Width           =   1695
      End
      Begin VB.Label LblCategory 
         AutoSize        =   -1  'True
         Caption         =   "License No"
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
         Left            =   120
         TabIndex        =   30
         Top             =   765
         Width           =   975
      End
      Begin VB.Label LblCategoryID 
         AutoSize        =   -1  'True
         Caption         =   "License ID"
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   330
         Width           =   765
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Location "
         Height          =   195
         Left            =   9480
         TabIndex        =   37
         Top             =   1200
         Visible         =   0   'False
         Width           =   660
      End
   End
   Begin MSComDlg.CommonDialog cdialog 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4695
      Left            =   120
      TabIndex        =   63
      Top             =   3000
      Width           =   14925
      _ExtentX        =   26326
      _ExtentY        =   8281
      _Version        =   393216
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
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
      TabCaption(0)   =   "Attachment Details"
      TabPicture(0)   =   "FrmMst_License.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "License Allocation"
      TabPicture(1)   =   "FrmMst_License.frx":0028
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "MSHFlexGrid3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame3"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Godown Mapping"
      TabPicture(2)   =   "FrmMst_License.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "MSHFlexGrid2"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Log History"
      TabPicture(3)   =   "FrmMst_License.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label9"
      Tab(3).Control(1)=   "Label10"
      Tab(3).Control(2)=   "TxtUpdated"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "TxtCreated"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).ControlCount=   4
      Begin VB.Frame Frame3 
         Height          =   1215
         Left            =   120
         TabIndex        =   80
         Top             =   360
         Width           =   14655
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
            ItemData        =   "FrmMst_License.frx":007C
            Left            =   3420
            List            =   "FrmMst_License.frx":007E
            Sorted          =   -1  'True
            TabIndex        =   14
            Top             =   675
            Width           =   6765
         End
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
            ItemData        =   "FrmMst_License.frx":0080
            Left            =   120
            List            =   "FrmMst_License.frx":0090
            Sorted          =   -1  'True
            TabIndex        =   13
            Top             =   675
            Width           =   3240
         End
         Begin VB.CommandButton CmdDeleteSAD 
            Caption         =   "Delete from List"
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
            Left            =   12360
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   720
            Width           =   1935
         End
         Begin VB.CommandButton CmdAdd2ListSAD 
            Caption         =   "Add to List"
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
            Left            =   10440
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   720
            Width           =   1935
         End
         Begin VB.Label lblstateregionlocation 
            AutoSize        =   -1  'True
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
            Left            =   6337
            TabIndex        =   82
            Top             =   360
            Width           =   930
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            Caption         =   "Allocation Type"
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
            Left            =   915
            TabIndex        =   81
            Top             =   360
            Width           =   1650
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
         Left            =   -72840
         Locked          =   -1  'True
         TabIndex        =   75
         TabStop         =   0   'False
         Top             =   840
         Width           =   7290
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
         Left            =   -72840
         Locked          =   -1  'True
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   1320
         Width           =   7290
      End
      Begin VB.Frame Frame2 
         Height          =   4215
         Left            =   -74880
         TabIndex        =   64
         Top             =   360
         Width           =   14700
         Begin VB.CommandButton CmdGetFile 
            Caption         =   "Get File"
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
            Left            =   12240
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   240
            Width           =   2070
         End
         Begin VB.CommandButton CmdAttachment 
            Caption         =   "Add Attachment"
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
            Left            =   12240
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   600
            Width           =   2070
         End
         Begin VB.CommandButton CmdRemoveAttachment 
            Caption         =   "Remove Attachment"
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
            Left            =   12240
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   960
            Width           =   2070
         End
         Begin VB.TextBox TxtAttachmentRemarks 
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
            Height          =   630
            Left            =   1365
            MaxLength       =   2000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   9
            Top             =   840
            Width           =   10755
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid5 
            Height          =   2550
            Left            =   120
            TabIndex        =   65
            Top             =   1560
            Width           =   14505
            _ExtentX        =   25585
            _ExtentY        =   4498
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
         Begin VB.Label lblAttachment 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Path of  file selected for attachment"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   120
            TabIndex        =   67
            Top             =   270
            Width           =   12000
         End
         Begin VB.Label Label6 
            Caption         =   "Attachment Remarks"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   120
            TabIndex        =   66
            Top             =   960
            Width           =   1185
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
         Height          =   4065
         Left            =   -74880
         TabIndex        =   78
         Top             =   480
         Width           =   14655
         _ExtentX        =   25850
         _ExtentY        =   7170
         _Version        =   393216
         FixedCols       =   0
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
         Height          =   2895
         Left            =   120
         TabIndex        =   79
         Top             =   1680
         Width           =   14715
         _ExtentX        =   25956
         _ExtentY        =   5106
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
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
      Begin VB.Label Label10 
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
         Left            =   -74640
         TabIndex        =   77
         Top             =   900
         Width           =   855
      End
      Begin VB.Label Label9 
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
         Left            =   -74640
         TabIndex        =   76
         Top             =   1380
         Width           =   885
      End
   End
End
Attribute VB_Name = "FrmMst_License"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg, mLocationID, mLicenseTypeID, mEmployeeID, mSEmployeeID As Variant
Dim File_Path As String, mStateLocGodID, mStateID, mCMPID, mGodownID As Variant
Dim strFilename As Variant

Private Sub CmbAllocationType_GotFocus()
tmp = CmbAllocationType.Text
CmbAllocationType.Text = tmp
End Sub

Private Sub CmbAllocationType_LostFocus()

If CmbAllocationType.Text = "" Then
   CmbAllocationTypeID.Text = ""
   CmbAllocationTypeID.Clear
   Exit Sub
End If

If LCase(CmbAllocationType.Text) <> "godown" And LCase(CmbAllocationType.Text) <> "location" And LCase(CmbAllocationType.Text) <> "state" And LCase(CmbAllocationType.Text) <> "cmp" Then
   MsgBox "Invalid Selection!!!"
   CmbAllocationType.SetFocus
   Exit Sub
End If

If LCase(CmbAllocationType.Text) = "state" Then
   lblstateregionlocation.Caption = "State"
ElseIf LCase(CmbAllocationType.Text) = "location" Then
   lblstateregionlocation.Caption = "Location"
ElseIf LCase(CmbAllocationType.Text) = "godown" Then
   lblstateregionlocation.Caption = "Godown"
ElseIf LCase(CmbAllocationType.Text) = "cmp" Then
   lblstateregionlocation.Caption = "CMP"
End If

If tmp <> CmbAllocationType.Text Then
   CmbAllocationTypeID.Text = ""
End If

End Sub


Private Sub CmbAllocationTypeID_GotFocus()

If CmbAllocationType.Text = "" Then
   MsgBox "Select Allocation Type!!!"
   CmbAllocationType.SetFocus
   Exit Sub
End If

CmbAllocationTypeID.Clear

If LCase(CmbAllocationType.Text) = "state" Then
   
   tmp = " Select Distinct MS.StateName As AllocationType,MS.StateID "
   tmp = tmp & " From Mst_Godown MG"
   tmp = tmp & " Inner Join Mst_CMP MC On MG.CMPID = MC.CMPID"
   tmp = tmp & " Inner Join Mst_Location ML On MC.LocationID = ML.LocationID"
   tmp = tmp & " Inner Join Mst_State MS On ML.StateID = MS.StateID"
   tmp = tmp & " Where MS.Flag = 'A' And MC.CloseDate Is Null And MG.CloseDate Is Null"
   tmp = tmp & " Order By MS.StateName"


   Call TmpTable
ElseIf LCase(CmbAllocationType.Text) = "location" Then
   
   tmp = " Select Distinct ML.LocationName As AllocationType,ML.LocationID,ML.StateID "
   tmp = tmp & " From Mst_Godown MG"
   tmp = tmp & " Inner Join Mst_CMP MC On MG.CMPID = MC.CMPID"
   tmp = tmp & " Inner Join Mst_Location ML On MC.LocationID = ML.LocationID"
   tmp = tmp & " Inner Join Mst_State MS On ML.StateID = MS.StateID"
   tmp = tmp & " Where MS.Flag = 'A' And MC.CloseDate Is Null And MG.CloseDate Is Null"
   tmp = tmp & " Order By ML.LocationName"
   
   Call TmpTable
ElseIf LCase(CmbAllocationType.Text) = "godown" Then
   
   tmp = " Select Distinct MC.CMPName+'~'+MG.GodownNo 'AllocationType',MG.GodownNo,MG.GodownID,ML.LocationID,MS.StateID,MC.CMPID,MC.CMPName"
   tmp = tmp & " From Mst_Godown MG"
   tmp = tmp & " Inner Join Mst_CMP MC On MG.CMPID = MC.CMPID"
   tmp = tmp & " Inner Join Mst_Location ML On MC.LocationID = ML.LocationID"
   tmp = tmp & " Inner Join Mst_State MS On ML.StateID = MS.StateID"
   tmp = tmp & " Where MS.Flag = 'A' And MC.CloseDate Is Null And MG.CloseDate Is Null"
   tmp = tmp & " Order By MC.CMPName,MG.GodownNo"

   Call TmpTable
ElseIf LCase(CmbAllocationType.Text) = "cmp" Then

   tmp = "Select Distinct MC.CMPName As AllocationType,MC.CMPID,ML.LocationID,MS.StateID "
   tmp = tmp & " From Mst_Godown MG"
   tmp = tmp & " Inner Join Mst_CMP MC On MG.CMPID = MC.CMPID"
   tmp = tmp & " Inner Join Mst_Location ML On MC.LocationID = ML.LocationID"
   tmp = tmp & " Inner Join Mst_State MS On ML.StateID = MS.StateID"
   tmp = tmp & " Where MS.Flag = 'A' And MC.CloseDate Is Null And MG.CloseDate Is Null"
   tmp = tmp & " Order By MC.CMPName"
   
   Call TmpTable
End If

If TmpRs.RecordCount = 0 Then
   MsgBox "No Record Found!!!"
   CmbAllocationTypeID.SetFocus
   Exit Sub
End If
For I = 1 To TmpRs.RecordCount
    CmbAllocationTypeID.AddItem TmpRs!AllocationType
    TmpRs.MoveNext
Next

End Sub

Private Sub CmbAllocationTypeID_LostFocus()
Dim strSplitString() As String
Dim CMPName_, GodownNo_ As String

If CmbAllocationTypeID.Text = "" Then
   mStateLocGodID = Null
   Exit Sub
End If

mStateID = Null
mLocationID = Null
mCMPID = Null
mGodownID = Null
    
If TmpRs.RecordCount > 0 Then TmpRs.MoveFirst

If LCase(CmbAllocationType.Text) = "state" Then
   TmpRs.Find "AllocationType = '" & CmbAllocationTypeID.Text & "'"
   If TmpRs.EOF Then
      MsgBox "Invalid Selection "
      CmbAllocationTypeID.SetFocus
      Exit Sub
   End If
   mStateLocGodID = TmpRs!StateID
   mStateID = TmpRs!StateID
   
ElseIf LCase(CmbAllocationType.Text) = "location" Then
    TmpRs.Find "AllocationType = '" & CmbAllocationTypeID.Text & "'"
    If TmpRs.EOF Then
       MsgBox "Invalid Selection "
       CmbAllocationTypeID.SetFocus
       Exit Sub
    End If
    mStateLocGodID = TmpRs!LocationID
    mStateID = TmpRs!StateID
    mLocationID = TmpRs!LocationID
    
ElseIf LCase(CmbAllocationType.Text) = "godown" Then
    strSplitString = Split(CmbAllocationTypeID, "~")
    
    If UBound(strSplitString) <> 1 Then
       MsgBox "Invalid Selection "
       CmbAllocationTypeID.SetFocus
       Exit Sub
    End If
    
    CMPName_ = Trim(strSplitString(0))
    GodownNo_ = Trim(strSplitString(1))

    TmpRs.Filter = "GodownNO = '" & GodownNo_ & "' ANd CMPName='" & CMPName_ & "'"
    If TmpRs.EOF Then
       MsgBox "Invalid Selection "
       CmbAllocationTypeID.SetFocus
       Exit Sub
    End If
    mStateLocGodID = TmpRs!GodownID
    mStateID = TmpRs!StateID
    mLocationID = TmpRs!LocationID
    mCMPID = TmpRs!CMPID
    mGodownID = TmpRs!GodownID
    TmpRs.Filter = ""
ElseIf LCase(CmbAllocationType.Text) = "cmp" Then
    TmpRs.Find "AllocationType = '" & CmbAllocationTypeID.Text & "'"
    If TmpRs.EOF Then
       MsgBox "Invalid Selection "
       CmbAllocationTypeID.SetFocus
       Exit Sub
    End If
    mStateLocGodID = TmpRs!CMPID
    mCMPID = TmpRs!CMPID
    mStateID = TmpRs!StateID
    mLocationID = TmpRs!LocationID
Else
    mStateLocGodID = Null
End If
End Sub


Private Sub CmbEmployeeID_GotFocus()
tmp = CmbEmployeeID.Text
Call TableMst_Employee("Where Flag = 'Y'")
CmbEmployeeID.Clear
If RsMst_Employee.RecordCount = 0 Then
   MsgBox "No Employee Found!!!"
   CmbEmployeeID.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Employee.RecordCount
    CmbEmployeeID.AddItem RsMst_Employee!EmployeeName & "~" & RsMst_Employee!EmployeeNo
    RsMst_Employee.MoveNext
Next
CmbEmployeeID.Text = tmp
End Sub


Private Sub CmbEmployeeID_LostFocus()
Dim strSplitString() As String
Dim EmpNo_, EmpName_ As String

If CmbEmployeeID.Text = "" Then
   mEmployeeID = Null
   Exit Sub
End If

strSplitString = Split(CmbEmployeeID, "~")
EmpName_ = Trim(strSplitString(0))
EmpNo_ = Trim(strSplitString(1))

RsMst_Employee.MoveFirst
RsMst_Employee.Find "EmployeeNo = '" & EmpNo_ & "'"

If RsMst_Employee.EOF Then
   MsgBox "Invalid Selection "
   CmbEmployeeID.SetFocus
   Exit Sub
End If
mEmployeeID = RsMst_Employee!EmployeeID
End Sub

Private Sub CmbLicenseType_GotFocus()
tmp = CmbLicenseType.Text
Call TableMst_LicenseType
CmbLicenseType.Clear
If RsMst_LicenseType.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_LicenseType.RecordCount
    CmbLicenseType.AddItem RsMst_LicenseType!LicenseTypeName
    RsMst_LicenseType.MoveNext
Next
CmbLicenseType.Text = tmp
End Sub

Private Sub CmbLicenseType_LostFocus()

If CmbLicenseType.Text = "" Then
   TxtCapacity.Visible = False
   Label11.Visible = False
   Label8.Visible = False
   CmbEmployeeID.Visible = False
   TxtEmployeeID.Visible = False
   Call ClearLicenseAllocationFrame
   Call Grid_Head2
   Frame3.Enabled = False
   Exit Sub
Else
   TxtCapacity.Visible = True
   Label11.Visible = True
   Label8.Visible = True
   CmbEmployeeID.Visible = True
   TxtEmployeeID.Visible = True
   Frame3.Enabled = True
End If

If CmbLicenseType.Text = "" Then Exit Sub

RsMst_LicenseType.MoveFirst
RsMst_LicenseType.Find "LicenseTypeName = '" & CmbLicenseType.Text & "'"
If RsMst_LicenseType.EOF Then
   MsgBox "Invalid selection "
   CmbLicenseType.SetFocus
   Exit Sub
End If
mLicenseTypeID = RsMst_LicenseType!LicenseTypeID

End Sub

Private Sub CmbLocationID_GotFocus()
tmp = CmbLocationID.Text
Call TableMst_Location
CmbLocationID.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
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
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocationID.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection "
   CmbLocationID.SetFocus
   Exit Sub
End If
mLocationID = RsMst_Location!LocationID
End Sub

Private Sub CmbSearchLocation_GotFocus()
tmp = CmbSearchLocation.Text
Call TableMst_Location
CmbSearchLocation.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbSearchLocation.AddItem RsMst_Location!Locationname
    RsMst_Location.MoveNext
Next
CmbSearchLocation.Text = tmp
End Sub

Private Sub CmbSearchLocation_LostFocus()
If CmbSearchLocation.Text = "" Then
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbSearchLocation.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection "
   CmbSearchLocation.SetFocus
   Exit Sub
End If

End Sub

Private Sub CmbSEmployeeID_GotFocus()
tmp = CmbSEmployeeID.Text
Call TableMst_Employee("Where Flag = 'Y'")
CmbSEmployeeID.Clear
If RsMst_Employee.RecordCount = 0 Then
   MsgBox "No Employee Found!!!"
   CmbSEmployeeID.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Employee.RecordCount
    CmbSEmployeeID.AddItem RsMst_Employee!EmployeeName & "~" & RsMst_Employee!EmployeeNo
    RsMst_Employee.MoveNext
Next
CmbSEmployeeID.Text = tmp
End Sub


Private Sub CmbSEmployeeID_LostFocus()
Dim strSplitString() As String
Dim EmpNo_, EmpName_ As String

If CmbSEmployeeID.Text = "" Then
   mEmployeeID = Null
   Exit Sub
End If

strSplitString = Split(CmbSEmployeeID, "~")
EmpName_ = Trim(strSplitString(0))
EmpNo_ = Trim(strSplitString(1))

RsMst_Employee.MoveFirst
RsMst_Employee.Find "EmployeeNo = '" & EmpNo_ & "'"

If RsMst_Employee.EOF Then
   MsgBox "Invalid Selection "
   CmbSEmployeeID.SetFocus
   Exit Sub
End If
mSEmployeeID = RsMst_Employee!EmployeeID

End Sub


Private Sub CmbSLicenseType_GotFocus()
tmp = CmbSLicenseType.Text
Call TableMst_LicenseType
CmbSLicenseType.Clear
If RsMst_LicenseType.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_LicenseType.RecordCount
    CmbSLicenseType.AddItem RsMst_LicenseType!LicenseTypeName
    RsMst_LicenseType.MoveNext
Next
CmbSLicenseType.Text = tmp
End Sub


Private Sub CmbSLicenseType_LostFocus()
If CmbSLicenseType.Text = "" Then
   Exit Sub
End If
RsMst_LicenseType.MoveFirst
RsMst_LicenseType.Find "LicenseTypeName = '" & CmbSLicenseType.Text & "'"
If RsMst_LicenseType.EOF Then
   MsgBox "Invalid selection "
   CmbSLicenseType.SetFocus
   Exit Sub
End If
End Sub


Private Sub CmdAdd2ListSAD_Click()
Dim r As Variant

If CmbAllocationType.Text = "" Then
   MsgBox "Blank Allocation Type!!! "
   CmbAllocationType.SetFocus
   Exit Sub
End If

If CmbAllocationTypeID.Text = "" Then
   MsgBox "Blank " & lblstateregionlocation.Caption & "!!! "
   CmbAllocationTypeID.SetFocus
   Exit Sub
End If

If ChkDuplicateAllocationType = False Then
   MsgBox "Duplicate Allocation Type not allowed !!!"
   CmbAllocationType.SetFocus
   Exit Sub
End If

If ChkDuplicateLicenseType(UCase(Left(CmbAllocationType.Text, 1))) = False Then
   MsgBox "Same license type already exist for " & CmbAllocationTypeID.Text & " " & lblstateregionlocation.Caption & " !!!"
   CmbAllocationType.SetFocus
   Exit Sub
End If

If CheckDuplicateGodown(UCase(Left(CmbAllocationType.Text, 1))) = False Then
   MsgBox "Duplicate State/Location/CMP/Godown Not Allowed !!!"
   CmbAllocationType.SetFocus
   Exit Sub
End If

If CmdAdd2ListSAD.Caption = "Update List" Then
   r = MSHFlexGrid3.RowSel
Else
   r = MSHFlexGrid3.Rows - 1
   MSHFlexGrid3.Rows = MSHFlexGrid3.Rows + 1
End If

MSHFlexGrid3.TextMatrix(r, 0) = r
MSHFlexGrid3.TextMatrix(r, 1) = mStateLocGodID
MSHFlexGrid3.TextMatrix(r, 2) = CmbAllocationType.Text
MSHFlexGrid3.TextMatrix(r, 3) = CmbAllocationTypeID.Text
MSHFlexGrid3.TextMatrix(r, 4) = UCase(Left(CmbAllocationType.Text, 1))

MSHFlexGrid3.TextMatrix(r, 5) = IIf(IsNull(mStateID), "", mStateID)
MSHFlexGrid3.TextMatrix(r, 6) = IIf(IsNull(mLocationID), "", mLocationID)
MSHFlexGrid3.TextMatrix(r, 7) = IIf(IsNull(mCMPID), "", mCMPID)
MSHFlexGrid3.TextMatrix(r, 8) = IIf(IsNull(mGodownID), "", mGodownID)

Call ClearLicenseAllocationFrame
CmbAllocationType.SetFocus

End Sub
Function ChkDuplicateAllocationType() As Boolean
Dim Retval As Boolean
Retval = True
If MSHFlexGrid3.Rows > 1 Then
   For u = 1 To MSHFlexGrid3.Rows - 1
       If MSHFlexGrid3.TextMatrix(u, 2) = CmbAllocationType.Text And MSHFlexGrid3.TextMatrix(u, 3) = CmbAllocationTypeID.Text Then
          If LCase(CmdAdd2ListSAD.Caption) = "add to list" Then
             Retval = False
             Exit For
          Else
             If u <> MSHFlexGrid3.RowSel Then
                Retval = False
             End If
          End If
       End If
   Next
End If
ChkDuplicateAllocationType = Retval
End Function
Function ChkDuplicateLicenseType(Flag As Variant) As Boolean
Dim Retval As Boolean
Retval = True

If Not IsNull(mStateID) Then
   tmp = " Select * From Txn_LicenseAllocationDetail TLAD"
   tmp = tmp & " Inner Join Mst_License TL on TLAD.LicenseID=TL.LicenseID"
   tmp = tmp & " Where TL.LicenseID<>" & Val(TxtLicenseID.Text) & " And TL.LicenseTypeID=" & Val(mLicenseTypeID) & " And TLAD.Flag='S' And TLAD.SLGID=" & Val(mStateID) & ""
   Call Tmp1Table
   If TmpRs1.RecordCount > 0 Then
      Retval = False
      GoTo Lbl
   End If
End If

If Not IsNull(mLocationID) Then
   tmp = " Select * From Txn_LicenseAllocationDetail TLAD"
   tmp = tmp & " Inner Join Mst_License TL on TLAD.LicenseID=TL.LicenseID"
   tmp = tmp & " Where TL.LicenseID<>" & Val(TxtLicenseID.Text) & " And TL.LicenseTypeID=" & Val(mLicenseTypeID) & " And TLAD.Flag='L' And TLAD.SLGID=" & Val(mLocationID) & ""
   Call Tmp1Table
   If TmpRs1.RecordCount > 0 Then
      Retval = False
      GoTo Lbl
   End If
End If

If Not IsNull(mCMPID) Then
   tmp = " Select * From Txn_LicenseAllocationDetail TLAD"
   tmp = tmp & " Inner Join Mst_License TL on TLAD.LicenseID=TL.LicenseID"
   tmp = tmp & " Where TL.LicenseID<>" & Val(TxtLicenseID.Text) & " And TL.LicenseTypeID=" & Val(mLicenseTypeID) & " And TLAD.Flag='C' And TLAD.SLGID=" & Val(mCMPID) & ""
   Call Tmp1Table
   If TmpRs1.RecordCount > 0 Then
      Retval = False
      GoTo Lbl
   End If
End If

If Not IsNull(mGodownID) Then
   tmp = " Select * From Txn_LicenseAllocationDetail TLAD"
   tmp = tmp & " Inner Join Mst_License TL on TLAD.LicenseID=TL.LicenseID"
   tmp = tmp & " Where TL.LicenseID<>" & Val(TxtLicenseID.Text) & " And TL.LicenseTypeID=" & Val(mLicenseTypeID) & " And TLAD.Flag='G' And TLAD.SLGID=" & Val(mGodownID) & ""
   Call Tmp1Table
   If TmpRs1.RecordCount > 0 Then
      Retval = False
      GoTo Lbl
   End If
End If

Lbl:
ChkDuplicateLicenseType = Retval

End Function
Private Sub CmdAttachment_Click()
Dim r As Variant

If lblAttachment.Caption = "File Attachment" Then
   MsgBox "Blank Attachment not allowed!!!"
   Exit Sub
End If

If TxtAttachmentRemarks = "" Then
   MsgBox "Blank Remarks not allowed!!!"
   TxtAttachmentRemarks.SetFocus
   Exit Sub
End If

If CmdAttachment.Caption = "Update Attachment" Then
   r = MSHFlexGrid3.RowSel
Else
   r = MSHFlexGrid5.Rows - 1
   MSHFlexGrid5.Rows = MSHFlexGrid5.Rows + 1
   MSHFlexGrid5.TextMatrix(r, 0) = ""
End If

MSHFlexGrid5.TextMatrix(r, 1) = GetFileFromPath(lblAttachment.Caption)
MSHFlexGrid5.TextMatrix(r, 2) = TxtAttachmentRemarks
MSHFlexGrid5.TextMatrix(r, 3) = lblAttachment.Caption
Call ClearFrame

End Sub

Private Sub CmdDeleteSAD_Click()
If CmbAllocationType.Text = "" Or CmbAllocationTypeID.Text = "" Then Exit Sub

I = MSHFlexGrid3.RowSel

If I > 0 Then
   If MSHFlexGrid3.TextMatrix(I, 1) <> "" Then
      For RO = I To MSHFlexGrid3.Rows - 2
          For C = 0 To MSHFlexGrid3.Cols - 1
              MSHFlexGrid3.TextMatrix(RO, C) = MSHFlexGrid3.TextMatrix(RO + 1, C)
          Next
      Next
      MSHFlexGrid3.Rows = MSHFlexGrid3.Rows - 1
   End If
End If
Call ClearLicenseAllocationFrame
End Sub

Private Sub CmdGetFile_Click()

With cdialog
     .InitDir = App.Path
     .DialogTitle = "Select File"
     .FilterIndex = 1
     .ShowOpen
End With

If cdialog.CancelError Then
   lblAttachment.Caption = ""
Else
   strFilename = cdialog.FileName
   If strFilename = "" Then Exit Sub
   lblAttachment.Refresh
   lblAttachment.Caption = strFilename
End If
End Sub

Private Sub CmdGo_Click()
Dim wc As Variant
wc = ""
tmp = ""

tmp = "Select Distinct ML.LicenseID,ML.LicenseNo,ML.LicenseDesc,MLo.LocationName,ML.IssuingAuthority,ML.LicenseDate,ML.LicenseExpiryDate,MLT.LicenseTypeName,ML.Capacity,ME.EmployeeName+'~'+ME.EmployeeNo As FullName "
tmp = tmp & " From Mst_License ML "
If CmbSearchLocation.Text = "" Then
   tmp = tmp & " Left Join Mst_Location MLo On  ML.LocationID = MLo.LocationID "
Else
   tmp = tmp & " Left Join Txn_LicenseGodownDetail TGD on ML.LicenseID = TGD.LicenseID"
   tmp = tmp & " Left Join Mst_Godown MG On TGD.GodownID = MG.GodownID"
   tmp = tmp & " Left Join Mst_CMP MC On MG.CMPID = MC.CMPID"
   tmp = tmp & " Left Join Mst_Location MLo On  MC.LocationID = MLo.LocationID"
End If
tmp = tmp & " Left Join Mst_LicenseType MLT on ML.LicenseTypeID=MLT.LicenseTypeID "
tmp = tmp & " Left Join Mst_Employee ME on ML.EmployeeID=ME.EmployeeID"

If OptLicenseDate.Value = True Then
   If Gen_DBType = "MDB" Then
      wc = " Where ML.LicenseDate Between #" & Format(DTPLicFromDate.Value, Gen_DatFormat) & "# AND #" & Format(DTPLicToDate.Value, Gen_DatFormat) & "#"
   Else
      wc = " Where ML.LicenseDate Between '" & Format(DTPLicFromDate.Value, Gen_DatFormat) & "' AND '" & Format(DTPLicToDate.Value, Gen_DatFormat) & "'"
   End If
End If

If OptLicenseNo.Value = True Then
   wc = " Where ML.LicenseNo Like '%" & TxtSearchLicense & "%'"
End If

If OptLicDesc.Value = True Then
    wc = " Where ML.LicenseDesc Like '%" & TxtSearchLicense & "%' "
End If

If OptIssuingAutho.Value = True Then
    wc = " Where ML.IssuingAuthority  Like '%" & TxtSearchLicense & "%'"
End If

If OptLicExpDate.Value = True Then
   If Gen_DBType = "MDB" Then
      wc = " Where ML.LicenseExpiryDate Between #" & Format(DTPLicFromDate.Value, Gen_DatFormat) & "# AND #" & Format(DTPLicToDate.Value, Gen_DatFormat) & "#"
   Else
      wc = " Where ML.LicenseExpiryDate Between '" & Format(DTPLicFromDate.Value, Gen_DatFormat) & "' AND '" & Format(DTPLicToDate.Value, Gen_DatFormat) & "'"
   End If
End If

If CmbSearchLocation.Text <> "" Then
   If wc = "" Then
      wc = " Where MLo.LocationName = '" & CmbSearchLocation.Text & "'"
   Else
      wc = wc & " And MLo.LocationName = '" & CmbSearchLocation.Text & "'"
   End If
End If

If CmbSLicenseType.Text <> "" Then
   If wc = "" Then
      wc = " Where MLT.LicenseTypeName = '" & CmbSLicenseType.Text & "'"
   Else
      wc = wc & " And MLT.LicenseTypeName = '" & CmbSLicenseType.Text & "'"
   End If
End If

If CmbSEmployeeID.Text <> "" Then
   If wc = "" Then
      wc = " Where ME.EmployeeID = " & mSEmployeeID & ""
   Else
      wc = wc & " And ME.EmployeeID = " & mSEmployeeID & ""
   End If
End If

tmp1 = " Order By ML.LicenseID "

tmp = tmp & " " & wc & " " & tmp1


Call Tmp3Table

Call Grid_Head

MSHFlexGrid1.Rows = TmpRs3.RecordCount + 2
For iu = 1 To TmpRs3.RecordCount
    For C = 0 To MSHFlexGrid1.Cols - 1
        MSHFlexGrid1.TextMatrix(iu, C) = IIf(IsNull(TmpRs3.Fields(C)), "", TmpRs3.Fields(C))
        If InStr(MSHFlexGrid1.TextMatrix(iu, C), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(iu, C)) Then
              MSHFlexGrid1.TextMatrix(iu, C) = Format(MSHFlexGrid1.TextMatrix(iu, C), "dd-mmm-yyyy")
           Else
              MSHFlexGrid1.TextMatrix(iu, C) = MSHFlexGrid1.TextMatrix(iu, C)
           End If
        Else
          MSHFlexGrid1.TextMatrix(iu, C) = MSHFlexGrid1.TextMatrix(iu, C)
        End If
    Next
       
    TmpRs3.MoveNext
Next
LblRecordCount.Refresh
LblRecordCount.Caption = TmpRs3.RecordCount

End Sub

Private Sub CmdLocationID_Click()
FrmMst_Location.Show
End Sub

Private Sub CmdRemoveAttachment_Click()
I = MSHFlexGrid5.RowSel
If I > 0 Then
    If MSHFlexGrid5.TextMatrix(I, 4) = "" And MSHFlexGrid5.TextMatrix(I, 1) <> "" Then
        For RO = I To MSHFlexGrid5.Rows - 2
            For C = 0 To MSHFlexGrid5.Cols - 1
                MSHFlexGrid5.TextMatrix(RO, C) = MSHFlexGrid5.TextMatrix(RO + 1, C)
            Next
        Next
        MSHFlexGrid5.Rows = MSHFlexGrid5.Rows - 1
    End If
End If

Call ClearFrame
End Sub

Private Sub ClearFrame()

lblAttachment.Caption = "File Attachment"
CmdAttachment.Caption = "Add Attachment"
TxtAttachmentRemarks = ""

End Sub

Private Sub SaveAttachmentDetails()
If MSHFlexGrid5.Rows <= 1 Then Exit Sub

Call TableMst_LicenseAttachment

For I = 1 To MSHFlexGrid5.Rows - 1
    If MSHFlexGrid5.TextMatrix(I, 1) = "" Then Exit For
    If MSHFlexGrid5.TextMatrix(I, 4) = "" Then
       RsMst_LicenseAttachment.AddNew
 
       RsMst_LicenseAttachment!TxnID = GetMaxAttachID
       RsMst_LicenseAttachment!LicenseID = TxtLicenseID.Text
       mFlag = SaveFileToDB(MSHFlexGrid5.TextMatrix(I, 3), RsMst_LicenseAttachment, "File")
       mFlag = GetFileFromPath(MSHFlexGrid5.TextMatrix(I, 3))
       RsMst_LicenseAttachment!FileName = "License" & TxtLicenseID.Text & "-" & mFlag
    Else
       RsMst_LicenseAttachment.MoveFirst
       RsMst_LicenseAttachment.Find "TxnID = " & MSHFlexGrid5.TextMatrix(I, 4) & ""
    End If
    RsMst_LicenseAttachment!Remarks = MSHFlexGrid5.TextMatrix(I, 2)
    RsMst_LicenseAttachment!UpdatedDate = Now
    RsMst_LicenseAttachment!UpdatedBy = Gen_UserLoginID
    RsMst_LicenseAttachment!G_UID = GetGUID
    If IsNull(RsMst_LicenseAttachment!CreatedBy) = True Or RsMst_LicenseAttachment!CreatedBy = "" Then
       RsMst_LicenseAttachment!CreatedBy = Gen_UserLoginID
       RsMst_LicenseAttachment!CreatedDate = Now
    Else
       RsMst_LicenseAttachment!UpdatedBy = Gen_UserLoginID
       RsMst_LicenseAttachment!UpdatedDate = Now
    End If
    RsMst_LicenseAttachment.Update
Next
End Sub

Private Function GetMaxAttachID() As Double
Dim Retval As Double
tmp = "Select max(TxnID) from Mst_LicenseAttachment"
Call TmpTableSMAttachment
Retval = IIf(IsNull(TmpRsSM_Attachment.Fields(0)), 0, TmpRsSM_Attachment.Fields(0))
Retval = Retval + 1
GetMaxAttachID = Retval
End Function

Private Sub IndataAttachmentDetails()
Call Grid_Head4
Call ClearFrame

Call TableMst_LicenseAttachment(" Where LicenseID = " & TxtLicenseID.Text & " ")

MSHFlexGrid5.Rows = RsMst_LicenseAttachment.RecordCount + 2
For I = 1 To RsMst_LicenseAttachment.RecordCount

    MSHFlexGrid5.TextMatrix(I, 0) = I
    MSHFlexGrid5.TextMatrix(I, 1) = RsMst_LicenseAttachment!FileName
    MSHFlexGrid5.TextMatrix(I, 2) = RsMst_LicenseAttachment!Remarks
    MSHFlexGrid5.TextMatrix(I, 3) = App.Path & "\Agreement Attachment\" & MSHFlexGrid5.TextMatrix(I, 1)
    File_Path = MSHFlexGrid5.TextMatrix(I, 3) 'App.Path & "\Project Attachment\"
    Call LoadFileFromDB(File_Path, RsMst_LicenseAttachment, "File")
    MSHFlexGrid5.TextMatrix(I, 4) = RsMst_LicenseAttachment!TxnID
    RsMst_LicenseAttachment.MoveNext
Next
lblAttachment.Enabled = False
TxtAttachmentRemarks.Locked = True
CmdAttachment.Enabled = False
CmdGetFile.Enabled = False
CmdRemoveAttachment.Enabled = False
End Sub
Private Sub IndataLicenseAllocationDetails()
Call Grid_Head3
Call ClearLicenseAllocationFrame

tmp = " Select ML.LicenseID,TLAD.AllocateID,TLAD.Flag,TLAD.SLGID,MS.StateName,MLO.LocationName,MCP.CMPNAME,MC.CMPName+'~'+MG.GodownNo 'AllocationType'"
tmp = tmp & " From Mst_License ML"
tmp = tmp & " inner join Txn_LicenseAllocationDetail TLAD on ML.LicenseID=TLAD.LicenseID"
tmp = tmp & " Left join Mst_state MS on TLAD.SLGID=MS.StateId"
tmp = tmp & " Left join Mst_Location MLO on TLAD.SLGID=MLO.LocationID"
tmp = tmp & " Left join Mst_Godown MG on TLAD.SLGID=MG.GodownId"
tmp = tmp & " Left join Mst_CMP MCP on TLAD.SLGID=MCP.CMPID"
tmp = tmp & " Left Join Mst_CMP MC On MG.CMPID = MC.CMPID"
tmp = tmp & " Where ML.LicenseID = " & TxtLicenseID.Text & ""

Call Tmp1Table

tmp = " Select ML.LocationID,ML.StateID From Mst_Location ML"
Call Tmp2Table

tmp = " Select MG.GodownID,ML.LocationID,MS.StateID,MC.CMPID"
tmp = tmp & " From Mst_Godown MG"
tmp = tmp & " Inner Join Mst_CMP MC On MG.CMPID = MC.CMPID"
tmp = tmp & " Inner Join Mst_Location ML On MC.LocationID = ML.LocationID"
tmp = tmp & " Inner Join Mst_State MS On ML.StateID = MS.StateID"
tmp = tmp & " Where MC.CloseDate Is Null And MG.CloseDate Is Null"
Call Tmp3Table

If TmpRs1.RecordCount > 0 Then
   MSHFlexGrid3.Rows = TmpRs1.RecordCount + 2
   For I = 1 To TmpRs1.RecordCount
       MSHFlexGrid3.TextMatrix(I, 0) = IIf(IsNull(TmpRs1!AllocateID), "", TmpRs1!AllocateID)
       MSHFlexGrid3.TextMatrix(I, 1) = IIf(IsNull(TmpRs1!SLGID), "", TmpRs1!SLGID)
       MSHFlexGrid3.TextMatrix(I, 4) = IIf(IsNull(TmpRs1!Flag), "", TmpRs1!Flag)

       If UCase(MSHFlexGrid3.TextMatrix(I, 4)) = "S" Then
          MSHFlexGrid3.TextMatrix(I, 2) = "State"
          MSHFlexGrid3.TextMatrix(I, 3) = TmpRs1!StateName
          MSHFlexGrid3.TextMatrix(I, 5) = IIf(IsNull(TmpRs1!SLGID), "", TmpRs1!SLGID)
       ElseIf UCase(MSHFlexGrid3.TextMatrix(I, 4)) = "L" Then
          MSHFlexGrid3.TextMatrix(I, 2) = "Location"
          MSHFlexGrid3.TextMatrix(I, 3) = TmpRs1!Locationname
          If TmpRs2.RecordCount > 0 Then TmpRs2.MoveFirst
          TmpRs2.Find "LocationID=" & MSHFlexGrid3.TextMatrix(I, 1) & ""
          If TmpRs2.EOF = False Then
             MSHFlexGrid3.TextMatrix(I, 5) = IIf(IsNull(TmpRs2!StateID), "", TmpRs2!StateID)
             MSHFlexGrid3.TextMatrix(I, 6) = IIf(IsNull(TmpRs2!LocationID), "", TmpRs2!LocationID)
          End If
       ElseIf UCase(MSHFlexGrid3.TextMatrix(I, 4)) = "G" Then
          MSHFlexGrid3.TextMatrix(I, 2) = "Godown"
          MSHFlexGrid3.TextMatrix(I, 3) = TmpRs1!AllocationType
          If TmpRs3.RecordCount > 0 Then TmpRs3.MoveFirst
          TmpRs3.Find "GodownID=" & MSHFlexGrid3.TextMatrix(I, 1) & ""
          If TmpRs3.EOF = False Then
             MSHFlexGrid3.TextMatrix(I, 5) = IIf(IsNull(TmpRs3!StateID), "", TmpRs3!StateID)
             MSHFlexGrid3.TextMatrix(I, 6) = IIf(IsNull(TmpRs3!LocationID), "", TmpRs3!LocationID)
             MSHFlexGrid3.TextMatrix(I, 7) = IIf(IsNull(TmpRs3!CMPID), "", TmpRs3!CMPID)
             MSHFlexGrid3.TextMatrix(I, 8) = IIf(IsNull(TmpRs3!GodownID), "", TmpRs3!GodownID)
          End If
       ElseIf UCase(MSHFlexGrid3.TextMatrix(I, 4)) = "C" Then
          MSHFlexGrid3.TextMatrix(I, 2) = "CMP"
          MSHFlexGrid3.TextMatrix(I, 3) = TmpRs1!CMPName
          If TmpRs3.RecordCount > 0 Then TmpRs3.MoveFirst
          TmpRs3.Find "CMPID=" & MSHFlexGrid3.TextMatrix(I, 1) & ""
          If TmpRs3.EOF = False Then
             MSHFlexGrid3.TextMatrix(I, 5) = IIf(IsNull(TmpRs3!StateID), "", TmpRs3!StateID)
             MSHFlexGrid3.TextMatrix(I, 6) = IIf(IsNull(TmpRs3!LocationID), "", TmpRs3!LocationID)
             MSHFlexGrid3.TextMatrix(I, 7) = IIf(IsNull(TmpRs3!CMPID), "", TmpRs3!CMPID)
             MSHFlexGrid3.TextMatrix(I, 8) = IIf(IsNull(TmpRs3!GodownID), "", TmpRs3!GodownID)
          End If
       End If
       TmpRs1.MoveNext
   Next
End If
End Sub
Private Sub Grid_Head4()
With MSHFlexGrid5
    .Clear
    .Rows = 2
    .Cols = 5
    .Font.Size = 10
    .WordWrap = True
    .TextMatrix(0, 0) = "SR No"
    .TextMatrix(0, 1) = "Attachment File Name"
    .TextMatrix(0, 2) = "Remarks"
    .TextMatrix(0, 3) = "File Path"
    .TextMatrix(0, 4) = "ID"
    .ColWidth(0) = 1000
    .ColWidth(1) = 5000
    .ColWidth(2) = 8200
    .ColWidth(3) = 0
    .ColWidth(4) = 0
End With
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
Call Grid_Head2
Call Grid_Head3

TxtSearchLicense = ""
CmbSearchLocation.Text = ""
Call TableMst_License
If RsMst_License.RecordCount = 0 Then
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
Public Sub Grid_Head2()
MSHFlexGrid2.Clear
MSHFlexGrid2.Rows = 3
MSHFlexGrid2.FixedRows = 2
MSHFlexGrid2.Cols = 10

MSHFlexGrid2.TextMatrix(1, 0) = ""
MSHFlexGrid2.TextMatrix(1, 1) = "CMP"
MSHFlexGrid2.TextMatrix(1, 2) = "Godown No"
MSHFlexGrid2.TextMatrix(1, 3) = "Location"
MSHFlexGrid2.TextMatrix(1, 4) = "State"
MSHFlexGrid2.TextMatrix(1, 5) = "Unit Type"
MSHFlexGrid2.TextMatrix(1, 6) = "Capacity"
MSHFlexGrid2.TextMatrix(1, 7) = "Start Date"
MSHFlexGrid2.TextMatrix(1, 8) = "Agreement End Date"
MSHFlexGrid2.TextMatrix(1, 9) = "GodownID"

MSHFlexGrid2.ColWidth(0) = 0
MSHFlexGrid2.ColWidth(1) = 1800
MSHFlexGrid2.ColWidth(2) = 1500
MSHFlexGrid2.ColWidth(3) = 2000
MSHFlexGrid2.ColWidth(4) = 2500
MSHFlexGrid2.ColWidth(5) = 2000
MSHFlexGrid2.ColWidth(6) = 1500
MSHFlexGrid2.ColWidth(7) = 1500
MSHFlexGrid2.ColWidth(8) = 1500
MSHFlexGrid2.ColWidth(9) = 0
End Sub
Public Sub Grid_Head3()
With MSHFlexGrid3
    .Clear
    .Rows = 2
    .FixedRows = 1
    .Cols = 9

    .TextMatrix(0, 0) = "Row ID"
    .TextMatrix(0, 1) = "State/Location/Godown ID"
    .TextMatrix(0, 2) = "Allocation Type"
    .TextMatrix(0, 3) = "State/Location/Godown"
    .TextMatrix(0, 4) = "Flag"
    
    .TextMatrix(0, 5) = "StateID"
    .TextMatrix(0, 6) = "LocationID"
    .TextMatrix(0, 7) = "CMPID"
    .TextMatrix(0, 8) = "GodownID"
    
    .ColWidth(0) = 0
    .ColWidth(1) = 0
    .ColWidth(2) = 3500
    .ColWidth(3) = 6500
    .ColWidth(4) = 0
    
    .ColWidth(5) = 1000
    .ColWidth(6) = 1000
    .ColWidth(7) = 1000
    .ColWidth(8) = 1000
End With
End Sub
Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Mst_License.xls")
Call Xls_Detail_Rpt
End Sub
Public Sub Xls_Detail_Rpt()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim oWS As Excel.Worksheet

Dim mRow As Integer
Dim mCol As Integer
MsgBox ("Wait till next message")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Mst_License.xls")
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

Private Sub MSHFlexGrid3_Click()
Dim I As Variant

I = MSHFlexGrid3.RowSel
If MSHFlexGrid3.TextMatrix(I, 0) = "" Then Exit Sub
CmbAllocationType.Text = MSHFlexGrid3.TextMatrix(I, 2)
CmbAllocationTypeID.Text = MSHFlexGrid3.TextMatrix(I, 3)

mStateLocGodID = MSHFlexGrid3.TextMatrix(I, 1)
mStateID = MSHFlexGrid3.TextMatrix(I, 5)
mLocationID = MSHFlexGrid3.TextMatrix(I, 6)
mCMPID = MSHFlexGrid3.TextMatrix(I, 7)
mGodownID = MSHFlexGrid3.TextMatrix(I, 8)

CmdAdd2ListSAD.Caption = "Update List"
CmbAllocationType.SetFocus
End Sub

Private Sub MSHFlexGrid5_Click()
Dim ans, r As Variant
Dim ID As Integer
r = MSHFlexGrid5.RowSel
If r < 1 Then Exit Sub
If MSHFlexGrid5.TextMatrix(r, 1) = "" Then Exit Sub

If SaveCmd.Enabled = False Then

   ans = MsgBox("Do You want to view this document ", vbYesNo)
   If ans = vbYes Then
      If Right(App.Path, 1) = "\" Then
         Shell "Explorer.exe " & App.Path & "Agreement Attachment", vbMaximizedFocus
      Else
         Shell "Explorer.exe " & App.Path & "\Agreement Attachment", vbMaximizedFocus
      End If
   End If
   Exit Sub
End If

lblAttachment.Caption = MSHFlexGrid5.TextMatrix(r, 3)
TxtAttachmentRemarks = MSHFlexGrid5.TextMatrix(r, 2)

CmdRemoveAttachment.Enabled = True

If MSHFlexGrid5.TextMatrix(r, 4) <> "" Then
   CmdRemoveAttachment.Enabled = False
End If

CmdAttachment.Caption = "Update Attachment"
End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_License.MoveNext
If RsMst_License.EOF Then
   RsMst_License.MoveLast
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

Private Sub OptAll_Click()
DTPLicFromDate.Enabled = False
DTPLicToDate.Enabled = False
End Sub

Private Sub OptIssuingAutho_Click()
Call OptLicenseNo_Click
End Sub

Private Sub OptLicDesc_Click()
Call OptLicenseNo_Click
End Sub

Private Sub OptLicenseDate_Click()
Call OptLicenseNo_Click
End Sub

Private Sub OptLicenseNo_Click()
LblSearchLicense.Visible = True
TxtSearchLicense.Visible = True
TxtSearchLicense = ""
CmbSearchLocation.Text = ""
CmdGo.Visible = True

DTPLicFromDate.Enabled = False
DTPLicToDate.Enabled = False

If OptLicenseDate.Value = True Then
    DTPLicFromDate.Enabled = True
    DTPLicToDate.Enabled = True
ElseIf OptLicenseNo.Value = True Then
   LblSearchLicense.Caption = "Search By WHR No"
ElseIf OptLicDesc.Value = True Then
   LblSearchLicense.Caption = "Search By License Description"
ElseIf OptIssuingAutho.Value = True Then
   LblSearchLicense.Caption = "Search By Issuing Authority"
ElseIf OptLicExpDate.Value = True Then
   DTPLicFromDate.Enabled = True
   DTPLicToDate.Enabled = True
End If
End Sub

Private Sub OptLicExpDate_Click()
Call OptLicenseNo_Click
End Sub

Private Sub PreviousCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_License.MovePrevious
If RsMst_License.BOF Then
   RsMst_License.MoveFirst
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
RsMst_License.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsMst_License.MoveLast
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
   Call GButtonLock(Me, True)
   Call Indata
   Exit Sub
End If
Frame0.Height = Me.Height - 500
Frame0.Left = Frame1.Left - 100
Frame0.Top = Frame1.Top - 100
Frame0.Width = Me.Width - 150
Frame0.Visible = True
MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 2000
RsMst_License.MoveFirst

End Sub
Public Sub Indata()
Call ClearFrame
TxtLicenseID.Locked = True
TxtLicenseNo.Locked = True
TxtLicenseDesc.Locked = True
TxtIssuingAuthority.Locked = True
TxtLicenseDate.Enabled = False
TxtLicenseExpiryDate.Enabled = False

CmbLocationID.Visible = False
TxtLocationID.Visible = True

CmbEmployeeID.Visible = False
TxtEmployeeID.Visible = True

CmbLicenseType.Visible = False
TxtLicenseType.Visible = True

tmp = " Select ML.LicenseID,ML.LicenseNo,ML.LicenseDesc,ML.IssuingAuthority,ML.LicenseDate,ML.LicenseExpiryDate,ML.LocationID,MLO.LocationName,"
tmp = tmp & " ML.LicenseTypeID,MLT.LicenseTypeName,ML.Capacity,ML.CreatedBy,ML.CreatedDate,ML.UpdatedBy,ML.UpdatedDate,ME.EmployeeName+'~'+ME.EmployeeNo as Person,ME.EmployeeID "
tmp = tmp & " From Mst_License ML"
tmp = tmp & " Left Join Mst_Location MLO on ML.LocationID=MLO.LocationID"
tmp = tmp & " Left Join Mst_LicenseType MLT on ML.LicenseTypeID=MLT.LicenseTypeID"
tmp = tmp & " Left Join Mst_Employee ME on ML.EmployeeID=ME.EmployeeID"
tmp = tmp & " Where ML.LicenseID=" & RsMst_License!LicenseID & ""
Call Tmp2Table
If TmpRs2.RecordCount > 0 Then
   TxtLicenseID = IIf(IsNull(TmpRs2!LicenseID), "", TmpRs2!LicenseID)
   TxtLicenseNo = IIf(IsNull(TmpRs2!LicenseNo), "", TmpRs2!LicenseNo)
   TxtLicenseDesc = IIf(IsNull(TmpRs2!LicenseDesc), "", TmpRs2!LicenseDesc)
   TxtIssuingAuthority = IIf(IsNull(TmpRs2!IssuingAuthority), "", TmpRs2!IssuingAuthority)
   TxtLicenseDate.Value = IIf(IsNull(TmpRs2!LicenseDate), "", TmpRs2!LicenseDate)
   TxtLicenseExpiryDate.Value = IIf(IsNull(TmpRs2!LicenseExpiryDate), "", TmpRs2!LicenseExpiryDate)
   mLocationID = TmpRs2!LocationID
   TxtLocationID = IIf(IsNull(TmpRs2!Locationname), "", TmpRs2!Locationname)
   TxtCreated = IIf(IsNull(TmpRs2!CreatedBy), "", TmpRs2!CreatedBy) & " :- " & IIf(IsNull(TmpRs2!CreatedDate), "", TmpRs2!CreatedDate)
   TxtUpdated = IIf(IsNull(TmpRs2!UpdatedBy), "", TmpRs2!UpdatedBy) & " :- " & IIf(IsNull(TmpRs2!UpdatedDate), "", TmpRs2!UpdatedDate)
   
   TxtLicenseType = IIf(IsNull(TmpRs2!LicenseTypeName), "", TmpRs2!LicenseTypeName)
   mLicenseTypeID = IIf(IsNull(TmpRs2!LicenseTypeID), Null, TmpRs2!LicenseTypeID)
   TxtCapacity = IIf(IsNull(TmpRs2!Capacity), "", TmpRs2!Capacity)
   mEmployeeID = IIf(IsNull(TmpRs2!EmployeeID), Null, TmpRs2!EmployeeID)
   TxtEmployeeID = IIf(IsNull(TmpRs2!Person), "", TmpRs2!Person)
End If

If RsMst_License.RecordCount = 1 Then
   Call GButtonLock_1(Me, True)
Else
   Call GButtonLock(Me, True)
End If

Call IndataAttachmentDetails

If Trim(TxtLicenseType.Text) <> "" Then
   Call IndataLicenseAllocationDetails
   Call IndataLicenseGodownMappingDetail
Else
   Call Grid_Head2
   Call Grid_Head3
End If
End Sub

Private Sub TxtCapacity_LostFocus()
If TxtCapacity.Text = "" Then Exit Sub

If IsNumeric(TxtCapacity) = False Or InStr(1, TxtCapacity.Text, ".", vbTextCompare) > 0 Then
   MsgBox "Invalid numeric format !!!!"
   TxtCapacity.SetFocus
   Exit Sub
End If

End Sub


Private Sub TxtLicenseID_LostFocus()
If TxtLicenseID = "" Then Exit Sub
If IsNumeric(TxtLicenseID) = False Then
   MsgBox "Invalid numeric format !!!!"
   TxtLicenseID.SetFocus
   Exit Sub
End If
TxtLicenseID = Val(TxtLicenseID)
End Sub
Private Sub AddCmd_Click()

Edit_Flg = "A"
Call GButtonLock(Me, False)

TxtLicenseID.Locked = True
TxtLicenseNo.Locked = False
TxtLicenseDesc.Locked = False
TxtIssuingAuthority.Locked = False
TxtLicenseDate.Enabled = True
TxtLicenseExpiryDate.Enabled = True

CmbLocationID.Visible = True
TxtLocationID.Visible = False

CmbLicenseType.Visible = True
TxtLicenseType.Visible = False

TxtCreated = ""
TxtUpdated = ""
TxtLicenseID = ""
TxtLicenseNo = ""
TxtLicenseDesc = ""
CmbLocationID.Text = ""
mLocationID = Null
TxtIssuingAuthority = ""
TxtLicenseDate.Value = Date
TxtLicenseExpiryDate.Value = Date
TxtLicenseType.Text = ""
CmbLicenseType.Text = ""

lblAttachment.Enabled = True
TxtAttachmentRemarks.Locked = False
CmdAttachment.Enabled = True
CmdGetFile.Enabled = True
CmdRemoveAttachment.Enabled = True

TxtCapacity.Text = ""
CmbEmployeeID.Text = ""
TxtEmployeeID.Text = ""

TxtCapacity.Visible = True
Label11.Visible = True
Label8.Visible = True
CmbEmployeeID.Visible = True
TxtEmployeeID.Visible = True

Call Grid_Head2
Call Grid_Head4
Call Grid_Head3

End Sub
Private Sub EditCmd_Click()
Edit_Flg = "E"
Call GButtonLock(Me, False)

TxtLicenseID.Locked = True

If LCase(Gen_UserRole) = "power" Or LCase(Gen_UserRole) = "superadmin" Or LCase(Gen_UserRole) = "headcmg" Or LCase(Gen_UserRole) = "admin" Or LCase(Gen_UserRole) = "fagadmin" Or LCase(Gen_UserRole) = "fag" Then
   TxtLicenseNo.Locked = False
   TxtLicenseDesc.Locked = False
   TxtIssuingAuthority.Locked = False
   TxtLicenseDate.Enabled = True
   TxtLicenseExpiryDate.Enabled = True
End If

CmbLocationID.Text = TxtLocationID

CmbLicenseType.Visible = True
TxtLicenseType.Visible = False
CmbLicenseType.Text = TxtLicenseType.Text

If TxtLicenseType.Text <> "" Then
   TxtCapacity.Visible = True
   CmbEmployeeID.Visible = True
   TxtEmployeeID.Visible = False
   CmbEmployeeID.Text = TxtEmployeeID.Text
   Frame3.Enabled = True
Else
   CmbEmployeeID.Visible = False
   TxtCapacity.Visible = False
   Frame3.Enabled = False
End If

lblAttachment.Enabled = True
TxtAttachmentRemarks.Locked = False
CmdAttachment.Enabled = True
CmdGetFile.Enabled = True
CmdRemoveAttachment.Enabled = True

End Sub
Private Sub DeleteCmd_Click()
Dim ans As Variant
Dim mGodownList As Variant

mGodownList = ""

On Error GoTo msgError
ans = MsgBox("Do you really want to DELETE this record???", vbYesNo)
If ans = vbYes Then
   
   If TxtLicenseType.Text <> "" Then
      mGodownList = GetGodownList()
   End If
   
   tmp = " Insert Into Log_MstLicenseDetail (ID,LicenseID,LicenseNo,LicenseDesc,IssuingAuthority,LicenseDate,LicenseExpiryDate,EmployeeID,LicenseTypeID,Capacity,Flag,CreatedBy,CreatedDate,GodownID)"
   tmp = tmp & " Select (Select Isnull(Max(ID),0) From Log_MstLicenseDetail)+1,LicenseID,LicenseNo,LicenseDesc,IssuingAuthority,LicenseDate,LicenseExpiryDate,EmployeeID,LicenseTypeID,Capacity,Flag,'" & Gen_UserLoginID & "','" & Now & "','" & mGodownList & "' From Mst_License Where LicenseID=" & Val(TxtLicenseID.Text) & ""
   Call Tmp4Table
   
   Call TableMst_LicenseAttachment("Where LicenseID=" & Val(TxtLicenseID.Text) & "")
   If RsMst_LicenseAttachment.RecordCount > 0 Then
      RsMst_LicenseAttachment.Delete
      RsMst_LicenseAttachment.Update
   End If
   
   tmp = " Delete From Txn_LicenseAllocationDetail Where LicenseID=" & Val(TxtLicenseID.Text) & ""
   Call Tmp1Table
   
   tmp = " Delete From Txn_LicenseGodownDetail Where LicenseID=" & Val(TxtLicenseID.Text) & ""
   Call Tmp1Table
   
   RsMst_License.Delete
   RsMst_License.Update
   If RsMst_License.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_License.MoveNext
   If RsMst_License.EOF() Then
      RsMst_License.MoveLast
   End If
   Call Indata
End If
Exit Sub

msgError:
MsgBox "Child record Present "
RsMst_License.CancelUpdate
End Sub
Private Sub SaveCmd_Click()


If TxtLicenseNo.Text = "" Then
   MsgBox "Blank License No Not Allowed !!! "
   TxtLicenseNo.SetFocus
   Exit Sub
End If


If TxtLicenseDesc.Text = "" Then
   MsgBox "Blank License Description  Not Allowed !!! "
   TxtLicenseDesc.SetFocus
   Exit Sub
End If

If TxtIssuingAuthority.Text = "" Then
   MsgBox "Blank Issuing Authority Not Allowed !!! "
   TxtIssuingAuthority.SetFocus
   Exit Sub
End If

If TxtLicenseExpiryDate.Value < TxtLicenseDate.Value Then
   MsgBox "License  date must be Less than  Expiry Date !!! "
   TxtLicenseDate.SetFocus
   Exit Sub
End If
If TxtLicenseDate.Value > Date Then
   MsgBox "Future Date is not Allowed !!!!! "
   TxtLicenseDate.SetFocus
   Exit Sub
End If

If CmbLicenseType.Text <> "" Then
   If TxtCapacity.Text = "" Then
      MsgBox "Blank capacity for which license issued Not Allowed"
      Exit Sub
   End If
   If CmbEmployeeID.Text = "" Then
      MsgBox "Blank name of person holding License Copy Not Allowed"
      Exit Sub
   End If
   
   If MSHFlexGrid5.TextMatrix(1, 1) = "" Then
      MsgBox "Blank Attachment Details Not Allowed "
      SSTab1.Tab = 0
      CmdGetFile.SetFocus
      Exit Sub
   End If
   
   If MSHFlexGrid3.TextMatrix(1, 1) = "" Then
      MsgBox "Blank License Allocation detail Not Allowed "
      SSTab1.Tab = 1
      CmbAllocationType.SetFocus
      Exit Sub
   End If
End If

If Edit_Flg = "A" Then
   If RsMst_License.RecordCount > 0 Then
      RsMst_License.MoveFirst
      RsMst_License.Find "LicenseNo= '" & TxtLicenseNo.Text & "'"
      If Not RsMst_License.EOF Then
        MsgBox "Duplicate License  Not Allowed !!! "
         TxtLicenseNo.SetFocus
         Exit Sub
      End If
   End If
   RsMst_License.AddNew
   RsMst_License!LicenseID = GetLicenseID
   RsMst_License!G_UID = GetGUID
   TxtLicenseID.Text = RsMst_License!LicenseID
Else
   RsMst_License.MoveFirst
   RsMst_License.Find "LicenseNo = '" & TxtLicenseNo.Text & "'"
   If Not RsMst_License.EOF Then
      If RsMst_License!LicenseID <> TxtLicenseID.Text Then
         MsgBox "Duplicate License Not Allowed !!! "
         TxtLicenseNo.SetFocus
         Exit Sub
      End If
   End If
   RsMst_License.MoveFirst
   RsMst_License.Find " LicenseID= " & TxtLicenseID.Text
End If

RsMst_License!LicenseID = TxtLicenseID.Text
RsMst_License!LicenseNo = TxtLicenseNo.Text
RsMst_License!LicenseDesc = TxtLicenseDesc.Text
RsMst_License!IssuingAuthority = TxtIssuingAuthority.Text
RsMst_License!LicenseDate = TxtLicenseDate.Value
RsMst_License!LicenseExpiryDate = TxtLicenseExpiryDate.Value

If CmbLicenseType.Text <> "" Then
   RsMst_License!LicenseTypeID = mLicenseTypeID
   RsMst_License!Capacity = TxtCapacity.Text
   RsMst_License!EmployeeID = mEmployeeID
End If

If CmbLocationID.Text <> "" Then
   RsMst_License!LocationID = mLocationID
Else
   RsMst_License!LocationID = 0
End If

If IsNull(RsMst_License!CreatedBy) = True Or RsMst_License!CreatedBy = "" Then
   RsMst_License!CreatedBy = Gen_UserLoginID
   RsMst_License!CreatedDate = Now
Else
   RsMst_License!UpdatedBy = Gen_UserLoginID
   RsMst_License!UpdatedDate = Now
End If
RsMst_License.Update

Call SaveAttachmentDetails

If Trim(CmbLicenseType.Text) <> "" Then
   Call SaveLicenseAllocationDetails
   Call GetGodown
   Call SaveLicenseGodownMappingDetail
End If
Call Indata

End Sub
Private Function GetLicenseID() As Double
Dim Retval As Double
tmp = "Select max(LicenseID) from Mst_License"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetLicenseID = Retval
End Function
Private Sub Grid_Head()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 10
MSHFlexGrid1.FixedRows = 1

MSHFlexGrid1.TextMatrix(0, 0) = "License ID"
MSHFlexGrid1.TextMatrix(0, 1) = "License No"
MSHFlexGrid1.TextMatrix(0, 2) = "License Description"
MSHFlexGrid1.TextMatrix(0, 3) = "Location"
MSHFlexGrid1.TextMatrix(0, 4) = "Issuing Authority"
MSHFlexGrid1.TextMatrix(0, 5) = "License Date"
MSHFlexGrid1.TextMatrix(0, 6) = "License Expiry Date"
MSHFlexGrid1.TextMatrix(0, 7) = "License Type"
MSHFlexGrid1.TextMatrix(0, 8) = "Capacity of license issued"
MSHFlexGrid1.TextMatrix(0, 9) = "Name of person holding License Copy"

MSHFlexGrid1.ColWidth(0) = 1000
MSHFlexGrid1.ColWidth(1) = 2400
MSHFlexGrid1.ColWidth(2) = 3000
MSHFlexGrid1.ColWidth(3) = 2000
MSHFlexGrid1.ColWidth(4) = 2000
MSHFlexGrid1.ColWidth(5) = 2000
MSHFlexGrid1.ColWidth(6) = 2000
MSHFlexGrid1.ColWidth(7) = 1500
MSHFlexGrid1.ColWidth(8) = 2000
MSHFlexGrid1.ColWidth(9) = 4000
End Sub
Private Sub MSHFlexGrid1_Click()
Dim I As Variant
I = MSHFlexGrid1.RowSel
If I <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(I, 0) = "" Then Exit Sub
RsMst_License.MoveFirst
RsMst_License.Find "LicenseID = " & MSHFlexGrid1.TextMatrix(I, 0)
If Not RsMst_License.EOF Then
   Call Indata
   Frame0.Visible = False
End If
End Sub


Public Sub ClearLicenseAllocationFrame()
CmbAllocationType.Text = ""
CmbAllocationTypeID = ""
CmdAdd2ListSAD.Caption = "Add to List"
End Sub

Public Sub SaveLicenseAllocationDetails()
If MSHFlexGrid3.Rows <= 1 Then Exit Sub

tmp = "Delete From Txn_LicenseAllocationDetail Where LicenseID=" & TxtLicenseID.Text & ""
Call Tmp1Table

If Trim(CmbLicenseType.Text) <> "" Then
   Call TableTxn_LicenseAllocationDetail("Where LicenseID=" & TxtLicenseID.Text & "")
   For I = 1 To MSHFlexGrid3.Rows - 1
       If MSHFlexGrid3.TextMatrix(I, 1) = "" Then Exit For
          RsTxn_LicenseAllocationDetail.AddNew
          RsTxn_LicenseAllocationDetail!AllocateID = GetMaxID("AllocateID", "Txn_LicenseAllocationDetail")
          RsTxn_LicenseAllocationDetail!LicenseID = TxtLicenseID.Text
          RsTxn_LicenseAllocationDetail!Flag = UCase(MSHFlexGrid3.TextMatrix(I, 4))
          RsTxn_LicenseAllocationDetail!SLGID = Val(MSHFlexGrid3.TextMatrix(I, 1))
          RsTxn_LicenseAllocationDetail!G_UID = GetGUID
          If IsNull(RsTxn_LicenseAllocationDetail!CreatedBy) = True Or RsTxn_LicenseAllocationDetail!CreatedBy = "" Then
             RsTxn_LicenseAllocationDetail!CreatedBy = Gen_UserLoginID
             RsTxn_LicenseAllocationDetail!CreatedDate = Now
          Else
             RsTxn_LicenseAllocationDetail!UpdatedBy = Gen_UserLoginID
             RsTxn_LicenseAllocationDetail!UpdatedDate = Now
          End If
          RsTxn_LicenseAllocationDetail.Update
   Next
End If
End Sub

Public Sub GetGodown()
Call Grid_Head2

tmp = " Select TGA.GodownID ,Max(TA.ExpiryDate) 'ExpDate' from Txn_GodownLessorAgreeDetail TGA "
tmp = tmp & " Inner Join Txn_LessorAgreement TA On TGA.LAID = TA.LAID "
tmp = tmp & " Where TA.Flag not in ('R') "
tmp = tmp & " Group BY TGA.GodownID "
      
Call Tmp3Table

tmp1 = " Select MC.CMPName,MG.GodownNO,ML.LocationName,MS.StateName,MUT.UnitType,MG.GodownCapacity"
tmp1 = tmp1 & " ,MG.StartDate,MG.AgreementEndDate,MG.GodownID "
tmp1 = tmp1 & " From Mst_Godown MG"
tmp1 = tmp1 & " Inner Join Mst_CMP MC On MG.CMPID = MC.CMPID"
tmp1 = tmp1 & " Inner Join Mst_UnitType MUT on MG.UnitTypeID=MUT.UnitTypeID"
tmp1 = tmp1 & " Inner Join Mst_Location ML On MC.LocationID = ML.LocationID"
tmp1 = tmp1 & " Inner Join Mst_State MS On ML.StateID = MS.StateID"

With MSHFlexGrid3
     For I = 1 To MSHFlexGrid3.Rows - 1
     
         If MSHFlexGrid3.TextMatrix(I, 4) = "S" Then
            tmp = tmp1 & " Where MS.StateID =" & MSHFlexGrid3.TextMatrix(I, 1) & " "
         ElseIf MSHFlexGrid3.TextMatrix(I, 4) = "L" Then
            tmp = tmp1 & " Where ML.LocationID =" & MSHFlexGrid3.TextMatrix(I, 1) & " "
         ElseIf MSHFlexGrid3.TextMatrix(I, 4) = "G" Then
            tmp = tmp1 & " Where MG.GodownID =" & MSHFlexGrid3.TextMatrix(I, 1) & ""
         ElseIf MSHFlexGrid3.TextMatrix(I, 4) = "C" Then
            tmp = tmp1 & " Where MC.CMPID =" & MSHFlexGrid3.TextMatrix(I, 1) & " "
         Else
            Exit For
         End If
         
         tmp = tmp & " And MG.CloseDate Is Null And MC.CloseDate IS Null"
         
         Call Tmp4Table
         If TmpRs4.RecordCount > 0 Then
            k = 1
            With MSHFlexGrid2
                 For iCtr = k To TmpRs4.RecordCount
                     j = MSHFlexGrid2.Rows - 1
                     MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
                     For C = 1 To MSHFlexGrid2.Cols - 1
                         MSHFlexGrid2.TextMatrix(j, C) = IIf(IsNull(TmpRs4.Fields(C - 1)), "", TmpRs4.Fields(C - 1))
                     Next
                     If LCase(TmpRs4!UnitType) = "leased" Then
                        If TmpRs3.RecordCount > 0 Then
                           TmpRs3.MoveFirst
                           TmpRs3.Find "GodownID = " & Val(MSHFlexGrid2.TextMatrix(iCtr, 9)) & ""
                           If TmpRs3.EOF = False Then
                              MSHFlexGrid2.TextMatrix(j, 8) = IIf(IsNull(TmpRs3!ExpDate), "", TmpRs3!ExpDate)
                           Else
                              MSHFlexGrid2.TextMatrix(j, 8) = ""
                           End If
                        End If
                     End If
                     If MSHFlexGrid2.TextMatrix(j, 8) <> "" Then
                        MSHFlexGrid2.TextMatrix(j, 8) = Format(MSHFlexGrid2.TextMatrix(j, 8), "dd-mmm-yyyy")
                     End If
                     If MSHFlexGrid2.TextMatrix(j, 7) <> "" Then
                        MSHFlexGrid2.TextMatrix(j, 7) = Format(MSHFlexGrid2.TextMatrix(j, 7), "dd-mmm-yyyy")
                     End If
                     TmpRs4.MoveNext
                 Next
            End With
         End If
     Next
End With
End Sub

Public Function CheckDuplicateGodown(Flag As Variant) As Boolean
Dim Retval As Boolean
Dim mCol As Variant
Dim mID As Variant
Dim mmGodownID, mmStateID, mmLocationID As Variant

Retval = True
If MSHFlexGrid3.Rows > 1 Then
   If Flag = "S" Then
      mCol = 5
      For u = 1 To MSHFlexGrid3.Rows - 1
          If MSHFlexGrid3.TextMatrix(u, 5) = mStateID Then
             If LCase(CmdAdd2ListSAD.Caption) = "add to list" Then
                Retval = False
                Exit For
             Else
                If u <> MSHFlexGrid3.RowSel Then
                   Retval = False
                End If
             End If
          End If
      Next
   ElseIf Flag = "L" Then
      mCol = 6
      For u = 1 To MSHFlexGrid3.Rows - 1
          If MSHFlexGrid3.TextMatrix(u, 5) = mLocationID And UCase(MSHFlexGrid3.TextMatrix(u, 4)) <> "L" Then
             If LCase(CmdAdd2ListSAD.Caption) = "add to list" Then
                Retval = False
                Exit For
             Else
                If u <> MSHFlexGrid3.RowSel Then
                   Retval = False
                End If
             End If
          End If
      Next
   ElseIf Flag = "G" Then
      mCol = 8
      For u = 1 To MSHFlexGrid3.Rows - 1
          If MSHFlexGrid3.TextMatrix(u, 5) = mGodownID And UCase(MSHFlexGrid3.TextMatrix(u, 4)) <> "G" Then
             If LCase(CmdAdd2ListSAD.Caption) = "add to list" Then
                Retval = False
                Exit For
             Else
                If u <> MSHFlexGrid3.RowSel Then
                   Retval = False
                End If
             End If
          End If
      Next
   ElseIf Flag = "C" Then
       mCol = 7
       For u = 1 To MSHFlexGrid3.Rows - 1
           If MSHFlexGrid3.TextMatrix(u, 5) = mCMPID And UCase(MSHFlexGrid3.TextMatrix(u, 4)) <> "C" Then
              If LCase(CmdAdd2ListSAD.Caption) = "add to list" Then
                 Retval = False
                 Exit For
              Else
                 If u <> MSHFlexGrid3.RowSel Then
                    Retval = False
                 End If
              End If
           End If
       Next
   End If
End If
CheckDuplicateGodown = Retval
End Function

Public Sub SaveLicenseGodownMappingDetail()
If MSHFlexGrid2.Rows <= 1 Then Exit Sub

tmp = "Delete From Txn_LicenseGodownDetail Where LicenseID=" & TxtLicenseID.Text & ""
Call Tmp1Table

If Trim(CmbLicenseType.Text) <> "" Then
   Call TableTxn_LicenseGodownDetail("Where LicenseID=" & TxtLicenseID.Text & "")
   For I = 2 To MSHFlexGrid2.Rows - 1
       If MSHFlexGrid2.TextMatrix(I, 1) = "" Then Exit For
          RsTxn_LicenseGodownDetail.AddNew
          RsTxn_LicenseGodownDetail!ID = GetMaxID("ID", "Txn_LicenseGodownDetail")
          RsTxn_LicenseGodownDetail!LicenseID = TxtLicenseID.Text
          RsTxn_LicenseGodownDetail!GodownID = IIf(MSHFlexGrid2.TextMatrix(I, 9) = "", Null, MSHFlexGrid2.TextMatrix(I, 9))
          RsTxn_LicenseGodownDetail!G_UID = GetGUID
          If IsNull(RsTxn_LicenseGodownDetail!CreatedBy) = True Or RsTxn_LicenseGodownDetail!CreatedBy = "" Then
             RsTxn_LicenseGodownDetail!CreatedBy = Gen_UserLoginID
             RsTxn_LicenseGodownDetail!CreatedDate = Now
          Else
             RsTxn_LicenseGodownDetail!UpdatedBy = Gen_UserLoginID
             RsTxn_LicenseGodownDetail!UpdatedDate = Now
          End If
          RsTxn_LicenseGodownDetail.Update
   Next
End If
End Sub

Public Sub IndataLicenseGodownMappingDetail()
Call Grid_Head2

tmp = " Select TGA.GodownID ,Max(TA.ExpiryDate) 'ExpDate' from Txn_GodownLessorAgreeDetail TGA "
tmp = tmp & " Inner Join Txn_LessorAgreement TA On TGA.LAID = TA.LAID "
tmp = tmp & " Where TA.Flag not in ('R') "
tmp = tmp & " Group BY TGA.GodownID "
      
Call Tmp3Table

tmp = " Select MC.CMPName,MG.GodownNO,ML.LocationName,MS.StateName,MUT.UnitType,MG.GodownCapacity"
tmp = tmp & " ,MG.StartDate,MG.AgreementEndDate,MG.GodownID "
tmp = tmp & " From Txn_LicenseGodownDetail TLGD"
tmp = tmp & " Inner join Mst_Godown MG on TLGD.GodownID=MG.GodownID "
tmp = tmp & " Inner Join Mst_CMP MC On MG.CMPID = MC.CMPID"
tmp = tmp & " Inner Join Mst_UnitType MUT on MG.UnitTypeID=MUT.UnitTypeID"
tmp = tmp & " Inner Join Mst_Location ML On MC.LocationID = ML.LocationID"
tmp = tmp & " Inner Join Mst_State MS On ML.StateID = MS.StateID"
tmp = tmp & " Where TLGD.LicenseID=" & TxtLicenseID.Text & ""
tmp = tmp & " Order by MC.CMPName,MG.GodownNO,ML.LocationName,MS.StateName,MUT.UnitType,MG.GodownCapacity"

Call Tmp4Table

If TmpRs4.RecordCount > 0 Then
   With MSHFlexGrid2
        For iCtr = 2 To TmpRs4.RecordCount + 1
            MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
            For C = 1 To MSHFlexGrid2.Cols - 1
                MSHFlexGrid2.TextMatrix(iCtr, C) = IIf(IsNull(TmpRs4.Fields(C - 1)), "", TmpRs4.Fields(C - 1))
            Next
            If LCase(TmpRs4!UnitType) = "leased" Then
               If TmpRs3.RecordCount > 0 Then
                  TmpRs3.MoveFirst
                  TmpRs3.Find "GodownID = " & Val(MSHFlexGrid2.TextMatrix(iCtr, 9)) & ""
                  If TmpRs3.EOF = False Then
                     MSHFlexGrid2.TextMatrix(iCtr, 8) = IIf(IsNull(TmpRs3!ExpDate), "", TmpRs3!ExpDate)
                  Else
                     MSHFlexGrid2.TextMatrix(iCtr, 8) = ""
                  End If
               End If
            End If
            MSHFlexGrid2.TextMatrix(iCtr, 8) = IIf(MSHFlexGrid2.TextMatrix(iCtr, 8) = "", "", Format(MSHFlexGrid2.TextMatrix(iCtr, 8), "dd-mmm-yyyy"))
            MSHFlexGrid2.TextMatrix(iCtr, 7) = IIf(MSHFlexGrid2.TextMatrix(iCtr, 7) = "", "", Format(MSHFlexGrid2.TextMatrix(iCtr, 7), "dd-mmm-yyyy"))
            TmpRs4.MoveNext
        Next
   End With
   MSHFlexGrid2.TextMatrix(0, 2) = Val(TmpRs4.RecordCount)
End If

End Sub

Public Function GetGodownList() As Variant
GetGodownList = ""
For I = 1 To MSHFlexGrid2.Rows - 1
    If MSHFlexGrid2.TextMatrix(I, 1) = "" Then Exit For
       If GetGodownList = "" Then
          GetGodownList = MSHFlexGrid2.TextMatrix(I, 9)
       Else
          GetGodownList = GetGodownList & "," & MSHFlexGrid2.TextMatrix(I, 9)
       End If
Next
GetGodownList = GetGodownList
End Function
