VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmRpt_CommTrackData 
   Caption         =   "Comm Track Data"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   60
      TabIndex        =   11
      Top             =   45
      Width           =   15015
      Begin VB.CheckBox Check2 
         Caption         =   "With NCDEX Code"
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
         Left            =   6960
         TabIndex        =   29
         Top             =   203
         Width           =   2295
      End
      Begin VB.CheckBox Check1 
         Caption         =   "For Opening Stock"
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
         Left            =   9600
         TabIndex        =   28
         Top             =   203
         Width           =   2175
      End
      Begin VB.ComboBox CmbBatch 
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
         Height          =   360
         Left            =   2760
         TabIndex        =   5
         Top             =   630
         Width           =   1365
      End
      Begin VB.OptionButton OptOld 
         Caption         =   "Old"
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
         TabIndex        =   2
         Top             =   690
         Width           =   855
      End
      Begin VB.OptionButton OptPending 
         Caption         =   "Pending"
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
         TabIndex        =   1
         Top             =   210
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.CommandButton CmdExport 
         Caption         =   "Export"
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
         Left            =   12390
         TabIndex        =   8
         Top             =   630
         Width           =   1215
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "Clear"
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
         Left            =   13605
         TabIndex        =   9
         Top             =   630
         Width           =   1215
      End
      Begin VB.CommandButton CmdGetReport 
         Caption         =   "Get Report"
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
         Left            =   9960
         TabIndex        =   6
         Top             =   630
         Width           =   1215
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
         Height          =   375
         Left            =   11175
         TabIndex        =   7
         Top             =   630
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker TxtFromDate 
         Height          =   375
         Left            =   2760
         TabIndex        =   3
         Top             =   143
         Width           =   1365
         _ExtentX        =   2408
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
         CustomFormat    =   "MM/dd/yyyy"
         Format          =   63832067
         CurrentDate     =   39850
         MinDate         =   38353
      End
      Begin MSComCtl2.DTPicker TxtToDate 
         Height          =   375
         Left            =   5280
         TabIndex        =   4
         Top             =   143
         Width           =   1380
         _ExtentX        =   2434
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
         CustomFormat    =   "MM/dd/yyyy"
         Format          =   63832067
         CurrentDate     =   39850
         MinDate         =   38353
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Batch No"
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
         Left            =   1680
         TabIndex        =   25
         Top             =   690
         Width           =   825
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   4320
         TabIndex        =   24
         Top             =   210
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "From  Date"
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
         Left            =   1680
         TabIndex        =   23
         Top             =   210
         Width           =   990
      End
   End
   Begin VB.Frame Frame2 
      Height          =   8775
      Left            =   60
      TabIndex        =   0
      Top             =   1125
      Width           =   15015
      Begin TabDlg.SSTab SSTab2 
         Height          =   8175
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   14805
         _ExtentX        =   26114
         _ExtentY        =   14420
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   1058
         TabCaption(0)   =   "DEPOSIT DATA"
         TabPicture(0)   =   "FrmRpt_CommTrackData.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "MSHFlexGrid10"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "WITHDRAWAL DATA"
         TabPicture(1)   =   "FrmRpt_CommTrackData.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "MSHFlexGrid11"
         Tab(1).ControlCount=   1
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid10 
            Height          =   7335
            Left            =   120
            TabIndex        =   26
            Top             =   720
            Width           =   14535
            _ExtentX        =   25638
            _ExtentY        =   12938
            _Version        =   393216
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid11 
            Height          =   7335
            Left            =   -74880
            TabIndex        =   27
            Top             =   720
            Width           =   14535
            _ExtentX        =   25638
            _ExtentY        =   12938
            _Version        =   393216
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
      Begin TabDlg.SSTab SSTab1 
         Height          =   8175
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   14420
         _Version        =   393216
         Tabs            =   11
         Tab             =   10
         TabsPerRow      =   11
         TabHeight       =   1058
         TabCaption(0)   =   "UOM_MSTR"
         TabPicture(0)   =   "FrmRpt_CommTrackData.frx":0038
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "MSHFlexGrid1"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "CLIENT_MSTR"
         TabPicture(1)   =   "FrmRpt_CommTrackData.frx":0054
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "MSHFlexGrid2"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "USER_MSTR"
         TabPicture(2)   =   "FrmRpt_CommTrackData.frx":0070
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "MSHFlexGrid3"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "ASS_MSTR"
         TabPicture(3)   =   "FrmRpt_CommTrackData.frx":008C
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "MSHFlexGrid4"
         Tab(3).ControlCount=   1
         TabCaption(4)   =   "LOCATION_MSTR"
         TabPicture(4)   =   "FrmRpt_CommTrackData.frx":00A8
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "MSHFlexGrid5"
         Tab(4).ControlCount=   1
         TabCaption(5)   =   "WH_MSTR"
         TabPicture(5)   =   "FrmRpt_CommTrackData.frx":00C4
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "MSHFlexGrid6"
         Tab(5).ControlCount=   1
         TabCaption(6)   =   "WSP_MSTR"
         TabPicture(6)   =   "FrmRpt_CommTrackData.frx":00E0
         Tab(6).ControlEnabled=   0   'False
         Tab(6).Control(0)=   "MSHFlexGrid7"
         Tab(6).ControlCount=   1
         TabCaption(7)   =   "COMMODITY_MSTR"
         TabPicture(7)   =   "FrmRpt_CommTrackData.frx":00FC
         Tab(7).ControlEnabled=   0   'False
         Tab(7).Control(0)=   "MSHFlexGrid8"
         Tab(7).ControlCount=   1
         TabCaption(8)   =   "WH_COMMODITY_MAP"
         TabPicture(8)   =   "FrmRpt_CommTrackData.frx":0118
         Tab(8).ControlEnabled=   0   'False
         Tab(8).Control(0)=   "MSHFlexGrid9"
         Tab(8).ControlCount=   1
         TabCaption(9)   =   "DEPOSIT DATA"
         TabPicture(9)   =   "FrmRpt_CommTrackData.frx":0134
         Tab(9).ControlEnabled=   0   'False
         Tab(9).ControlCount=   0
         TabCaption(10)  =   "WITHDRAWAL DATA"
         TabPicture(10)  =   "FrmRpt_CommTrackData.frx":0150
         Tab(10).ControlEnabled=   -1  'True
         Tab(10).ControlCount=   0
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   7335
            Left            =   -74880
            TabIndex        =   13
            Top             =   720
            Width           =   14535
            _ExtentX        =   25638
            _ExtentY        =   12938
            _Version        =   393216
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
            Height          =   7335
            Left            =   -74880
            TabIndex        =   14
            Top             =   720
            Width           =   14535
            _ExtentX        =   25638
            _ExtentY        =   12938
            _Version        =   393216
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
            Height          =   7335
            Left            =   -74880
            TabIndex        =   15
            Top             =   720
            Width           =   14535
            _ExtentX        =   25638
            _ExtentY        =   12938
            _Version        =   393216
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid4 
            Height          =   7335
            Left            =   -74880
            TabIndex        =   16
            Top             =   720
            Width           =   14535
            _ExtentX        =   25638
            _ExtentY        =   12938
            _Version        =   393216
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid5 
            Height          =   7335
            Left            =   -74880
            TabIndex        =   17
            Top             =   720
            Width           =   14535
            _ExtentX        =   25638
            _ExtentY        =   12938
            _Version        =   393216
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid6 
            Height          =   7335
            Left            =   -74880
            TabIndex        =   18
            Top             =   720
            Width           =   14535
            _ExtentX        =   25638
            _ExtentY        =   12938
            _Version        =   393216
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid7 
            Height          =   7335
            Left            =   -74880
            TabIndex        =   19
            Top             =   720
            Width           =   14535
            _ExtentX        =   25638
            _ExtentY        =   12938
            _Version        =   393216
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid8 
            Height          =   7335
            Left            =   -74880
            TabIndex        =   20
            Top             =   720
            Width           =   14535
            _ExtentX        =   25638
            _ExtentY        =   12938
            _Version        =   393216
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid9 
            Height          =   7335
            Left            =   -74880
            TabIndex        =   21
            Top             =   720
            Width           =   14535
            _ExtentX        =   25638
            _ExtentY        =   12938
            _Version        =   393216
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   8400
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
   End
End
Attribute VB_Name = "FrmRpt_CommTrackData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mBatchID As Long
Dim mGodownIDs As String
Dim mGSLs As String

Private Sub Grid_Head1()
CmdExcel.Enabled = False
CmdExport.Enabled = False
With MSHFlexGrid1
   .Clear
   .Rows = 2
   .Cols = 8
   
   .TextMatrix(0, 0) = "UOM_CODE"
   .TextMatrix(0, 1) = "UOM_NAME"
   .TextMatrix(0, 2) = "UOM_UNIT"
   .TextMatrix(0, 3) = "IS_FREEZED"
   .TextMatrix(0, 4) = "IS_ACTIVE"
   .TextMatrix(0, 5) = "UOM_NAME_0"
   .TextMatrix(0, 6) = "UOM_NAME_1"
   .TextMatrix(0, 7) = "UOM_NAME_2"
      
   .RowHeight(0) = 600
End With
End Sub

Private Sub Grid_Head2()
With MSHFlexGrid2
   .Clear
   .Rows = 2
   .Cols = 127
   
   .TextMatrix(0, 0) = "CLIENT_ID"
   .TextMatrix(0, 1) = "CLIENT_TYPE"
   .TextMatrix(0, 2) = "AUX_CLIENT_ID"
   .TextMatrix(0, 3) = "TRADING_CLIENT_ID"
   .TextMatrix(0, 4) = "CLIENT_NAME"
   .TextMatrix(0, 5) = "CLIENT_ADDRESS"
   .TextMatrix(0, 6) = "CLIENT_DOC_TYPE"
   .TextMatrix(0, 7) = "CLIENT_PHONE1"
   .TextMatrix(0, 8) = "CLIENT_PHONE2"
   .TextMatrix(0, 9) = "WH_CODE"
   .TextMatrix(0, 10) = "LOCATION_CODE"
   .TextMatrix(0, 11) = "CLIENT_EMAIL"
   .TextMatrix(0, 12) = "CLIENT_ORG_TYPE"
   .TextMatrix(0, 13) = "CLIENT_CONTACT"
   .TextMatrix(0, 14) = "OPERATOR_ID"
   .TextMatrix(0, 15) = "OPERATOR_NAME"
   .TextMatrix(0, 16) = "CONTACT_PERSON_NAME"
   .TextMatrix(0, 17) = "CLIENT_MOB_NO"
   .TextMatrix(0, 18) = "OPERATOR_CONTACT_NO"
   .TextMatrix(0, 19) = "CLIENT_DOC_NO"
   .TextMatrix(0, 20) = "CLIENT_PASSWORD"
   .TextMatrix(0, 21) = "EXCHANGE_CODE"
   .TextMatrix(0, 22) = "CM_ID"
   .TextMatrix(0, 23) = "DP_ID"
   .TextMatrix(0, 24) = "TRADING_ACC_ID"
   .TextMatrix(0, 25) = "APPROVAL_STATUS"
   .TextMatrix(0, 26) = "APPROVAL_DATE"
   .TextMatrix(0, 27) = "UCC_ID"
   .TextMatrix(0, 28) = "CLIENT_TEL"
   .TextMatrix(0, 29) = "CLIENT_TLX"
   .TextMatrix(0, 30) = "CLIENT_FAX"
   .TextMatrix(0, 31) = "PASSWORD"
   .TextMatrix(0, 32) = "OPERATOR_CATEGORY"
   .TextMatrix(0, 33) = "CONTACT_PERSON_CONTACT"
   .TextMatrix(0, 34) = "CREATED_BY"
   .TextMatrix(0, 35) = "CREATED_ON"
   .TextMatrix(0, 36) = "MODIFIED_BY"
   .TextMatrix(0, 37) = "MODIFIED_ON"
   .TextMatrix(0, 38) = "CMSACCNT"
   .TextMatrix(0, 39) = "CLIENT_NEW_EMAIL"
   .TextMatrix(0, 40) = "CLIENT_NAME_0"
   .TextMatrix(0, 41) = "CLIENT_NAME_1"
   .TextMatrix(0, 42) = "IS_FREEZED"
   .TextMatrix(0, 43) = "IS_ACTIVE"
   .TextMatrix(0, 44) = "PLEDGING_ALLOWED"
   .TextMatrix(0, 45) = "AS_ADDRESS"
   .TextMatrix(0, 46) = "AS_ALTERNATE_TEL"
   .TextMatrix(0, 47) = "AS_CITY"
   .TextMatrix(0, 48) = "AS_DESIGNATION"
   .TextMatrix(0, 49) = "AS_EMAIL_ID"
   .TextMatrix(0, 50) = "AS_FAX"
   .TextMatrix(0, 51) = "AS_MOBILE_NO"
   .TextMatrix(0, 52) = "AS_NAME"
   .TextMatrix(0, 53) = "AS_PAN_NAME"
   .TextMatrix(0, 54) = "AS_PINCODE"
   .TextMatrix(0, 55) = "AS_TEL"
   .TextMatrix(0, 56) = "BD_BANK_ACCOUNT_NO"
   .TextMatrix(0, 57) = "BD_BANK_CODE"
   .TextMatrix(0, 58) = "BD_BRANCH_CODE"
   .TextMatrix(0, 59) = "BD_IFSC_CODE"
   .TextMatrix(0, 60) = "CA_ALTERNATE_TEL"
   .TextMatrix(0, 61) = "CA_DESIGNATION"
   .TextMatrix(0, 62) = "CA_FAX"
   .TextMatrix(0, 63) = "CA_PINCODE"
   .TextMatrix(0, 64) = "CONTACT_PERSON_NO"
   .TextMatrix(0, 65) = "DATE_OF_BIRTH"
   .TextMatrix(0, 66) = "FIRST_LOGIN_FLAG"
   .TextMatrix(0, 67) = "MEMBERSHIP_TYPE"
   .TextMatrix(0, 68) = "MEMBER_SHIP_ID"
   .TextMatrix(0, 69) = "OPERATOR_TYPE"
   .TextMatrix(0, 70) = "ORG_CONSTITUTION"
   .TextMatrix(0, 71) = "PAN_NO"
   .TextMatrix(0, 72) = "PD_AMOUNT"
   .TextMatrix(0, 73) = "PD_APPLICANTS_BANK"
   .TextMatrix(0, 74) = "PD_APP_BANK_BRANCH"
   .TextMatrix(0, 75) = "PD_CLEARING_BANK"
   .TextMatrix(0, 76) = "PD_CL_BANK_BRANCH"
   .TextMatrix(0, 77) = "PD_IFSC_CODE"
   .TextMatrix(0, 78) = "PD_INSTRUMENT_NO"
   .TextMatrix(0, 79) = "PD_INSTRUMENT_TYPE"
   .TextMatrix(0, 80) = "PRCNT_REBATE"
   .TextMatrix(0, 81) = "RO_ADDRESS"
   .TextMatrix(0, 82) = "RO_ALTERNATE_TEL"
   .TextMatrix(0, 83) = "RO_CITY"
   .TextMatrix(0, 84) = "RO_EMAIL_ID"
   .TextMatrix(0, 85) = "RO_FAX"
   .TextMatrix(0, 86) = "RO_PINCODE"
   .TextMatrix(0, 87) = "RO_TEL"
   .TextMatrix(0, 88) = "RO_TELEX"
   .TextMatrix(0, 89) = "TAN_NO"
   .TextMatrix(0, 90) = "FIRST_CLIENT_ADDRESS"
   .TextMatrix(0, 91) = "FIRST_CITY"
   .TextMatrix(0, 92) = "FIRST_FAT_HUBBY"
   .TextMatrix(0, 93) = "FIRST_HOLDER_NAME"
   .TextMatrix(0, 94) = "FIRST_OCCUPATION"
   .TextMatrix(0, 95) = "FIRST_PAN"
   .TextMatrix(0, 96) = "FIRST_PINCODE"
   .TextMatrix(0, 97) = "MODE_OF_OPERATION"
   .TextMatrix(0, 98) = "NOMINEE_ADDRESS"
   .TextMatrix(0, 99) = "NOMINEE_AGE"
   .TextMatrix(0, 100) = "NOMINEE_CITY"
   .TextMatrix(0, 101) = "NOMINEE_NAME"
   .TextMatrix(0, 102) = "NOMINEE_OCCUPATION"
   .TextMatrix(0, 103) = "NOMINEE_PINCODE"
   .TextMatrix(0, 104) = "NOMINEE_RELATION"
   .TextMatrix(0, 105) = "SECOND_CITY"
   .TextMatrix(0, 106) = "SECOND_CLIENT_ADDRESS"
   .TextMatrix(0, 107) = "SECOND_FAT_HUBBY"
   .TextMatrix(0, 108) = "SECOND_HOLDER_NAME"
   .TextMatrix(0, 109) = "SECOND_OCCUPATION"
   .TextMatrix(0, 110) = "SECOND_PAN"
   .TextMatrix(0, 111) = "SECOND_PINCODE"
   .TextMatrix(0, 112) = "THIRD_CITY"
   .TextMatrix(0, 113) = "THIRD_CLIENT_ADDRESS"
   .TextMatrix(0, 114) = "THIRD_FAT_HUBBY"
   .TextMatrix(0, 115) = "THIRD_HOLDER_NAME"
   .TextMatrix(0, 116) = "THIRD_OCCUPATION"
   .TextMatrix(0, 117) = "THIRD_PAN"
   .TextMatrix(0, 118) = "THIRD_PINCODE"
   .TextMatrix(0, 119) = "UPLOAD_FOR_FIRST_HOLDER"
   .TextMatrix(0, 120) = "UPLOAD_FOR_SECOND_HOLDER"
   .TextMatrix(0, 121) = "UPLOAD_FOR_THIRD_HOLDER"
   .TextMatrix(0, 122) = "ORG_CLIENT_TYPE"
   .TextMatrix(0, 123) = "ACCOUNT_VALIDITY_DATE"
   .TextMatrix(0, 124) = "ACCOUNT_OPENING_CHARGES"
   .TextMatrix(0, 125) = "RENEWAL_CHARGES"
   .TextMatrix(0, 126) = "BRANCH_CODE"
   
   .RowHeight(0) = 600
End With
End Sub

Private Sub Grid_Head3()
With MSHFlexGrid3
   .Clear
   .Rows = 2
   .Cols = 48
   
   .TextMatrix(0, 0) = "USER_ID"
   .TextMatrix(0, 1) = "ORG_ID"
   .TextMatrix(0, 2) = "ORG_NAME"
   .TextMatrix(0, 3) = "ORG_CATEGORY"
   .TextMatrix(0, 4) = "USER_NAME"
   .TextMatrix(0, 5) = "USER_EMAIL"
   .TextMatrix(0, 6) = "USER_CONTACT"
   .TextMatrix(0, 7) = "USER_PHONE"
   .TextMatrix(0, 8) = "STD_ADDRESS"
   .TextMatrix(0, 9) = "USER_ROLES"
   .TextMatrix(0, 10) = "MAKER"
   .TextMatrix(0, 11) = "CHECKER"
   .TextMatrix(0, 12) = "WH_CODE"
   .TextMatrix(0, 13) = "DOCUMENT_NO"
   .TextMatrix(0, 14) = "DOCUMENT_TYPE"
   .TextMatrix(0, 15) = "USER_MOBILE"
   .TextMatrix(0, 16) = "USER_PASSWORD"
   .TextMatrix(0, 17) = "PASSWORD"
   .TextMatrix(0, 18) = "USER_CATEGORY"
   .TextMatrix(0, 19) = "FIRST_LOGIN_FLAG"
   .TextMatrix(0, 20) = "CRT_USR_ID"
   .TextMatrix(0, 21) = "CRT_TM"
   .TextMatrix(0, 22) = "MOD_USR_ID"
   .TextMatrix(0, 23) = "MOD_TM"
   .TextMatrix(0, 24) = "USER_COMM_GROUP"
   .TextMatrix(0, 25) = "USER_LANGUAGE"
   .TextMatrix(0, 26) = "IS_FREEZED"
   .TextMatrix(0, 27) = "IS_ACTIVE"
   .TextMatrix(0, 28) = "ORG_ROLE"
   .TextMatrix(0, 29) = "DEALER_CM"
   .TextMatrix(0, 30) = "DEALER_DM"
   .TextMatrix(0, 31) = "DEALER_MM"
   .TextMatrix(0, 32) = "DEALER_AO"
   .TextMatrix(0, 33) = "DEALER_BIDS"
   .TextMatrix(0, 34) = "ACC_CREATION"
   .TextMatrix(0, 35) = "FUNDS"
   .TextMatrix(0, 36) = "CST"
   .TextMatrix(0, 37) = "SUPER"
   .TextMatrix(0, 38) = "CMS"
   .TextMatrix(0, 39) = "BANK"
   .TextMatrix(0, 40) = "FIN"
   .TextMatrix(0, 41) = "ACTIVE"
   .TextMatrix(0, 42) = "LAST_LOGIN_TIME"
   .TextMatrix(0, 43) = "PASSWORD_EXPIRY_DATE"
   .TextMatrix(0, 44) = "IS_LOGGED_IN"
   .TextMatrix(0, 45) = "EXPOSURE_LIMIT"
   .TextMatrix(0, 46) = "CRT_SYS_NM"
   .TextMatrix(0, 47) = "DEALER_WH"
   
   .RowHeight(0) = 600
End With
End Sub

Private Sub Grid_Head4()
With MSHFlexGrid4
   .Clear
   .Rows = 2
   .Cols = 70
   
   .TextMatrix(0, 0) = "ASS_ID"
   .TextMatrix(0, 1) = "ASS_NAME"
   .TextMatrix(0, 2) = "ASS_EMAIL"
   .TextMatrix(0, 3) = "ASS_CONTACT"
   .TextMatrix(0, 4) = "ASS_PHONE"
   .TextMatrix(0, 5) = "ASS_PHONE_2"
   .TextMatrix(0, 6) = "ASS_FAX_NO"
   .TextMatrix(0, 7) = "ASS_ADDRESS"
   .TextMatrix(0, 8) = "ASS_DOCUMENT_TYPE"
   .TextMatrix(0, 9) = "ASS_DOCUMENT_NUMBER"
   .TextMatrix(0, 10) = "WH_CODE"
   .TextMatrix(0, 11) = "COMMODITY_CODE"
   .TextMatrix(0, 12) = "ASS_PASSWORD"
   .TextMatrix(0, 13) = "PASSWORD"
   .TextMatrix(0, 14) = "FIRST_LOGIN_FLAG"
   .TextMatrix(0, 15) = "CREATED_BY"
   .TextMatrix(0, 16) = "CREATED_ON"
   .TextMatrix(0, 17) = "MODIFIED_BY"
   .TextMatrix(0, 18) = "MODIFIED_ON"
   .TextMatrix(0, 19) = "ASS_NEW_EMAIL"
   .TextMatrix(0, 20) = "ASS_NAME_0"
   .TextMatrix(0, 21) = "ASS_NAME_1"
   .TextMatrix(0, 22) = "IS_FREEZED"
   .TextMatrix(0, 23) = "IS_ACTIVE"
   .TextMatrix(0, 24) = "ASSAYER_NAME_1"
   .TextMatrix(0, 25) = "ASS_ADDRESS_0"
   .TextMatrix(0, 26) = "ASS_ADDRESS_1"
   .TextMatrix(0, 27) = "AS_ADDRESS"
   .TextMatrix(0, 28) = "AS_ALTERNATE_TEL"
   .TextMatrix(0, 29) = "AS_CITY"
   .TextMatrix(0, 30) = "AS_DESIGNATION"
   .TextMatrix(0, 31) = "AS_EMAIL_ID"
   .TextMatrix(0, 32) = "AS_FAX"
   .TextMatrix(0, 33) = "AS_MOBILE_NO"
   .TextMatrix(0, 34) = "AS_NAME"
   .TextMatrix(0, 35) = "AS_PAN_NAME"
   .TextMatrix(0, 36) = "AS_PINCODE"
   .TextMatrix(0, 37) = "AS_TEL"
   .TextMatrix(0, 38) = "BD_BANK_ACCOUNT_NO"
   .TextMatrix(0, 39) = "BD_BANK_CODE"
   .TextMatrix(0, 40) = "BD_BRANCH_CODE"
   .TextMatrix(0, 41) = "BD_IFSC_CODE"
   .TextMatrix(0, 42) = "CA_ALTERNATE_TEL"
   .TextMatrix(0, 43) = "CA_CITY"
   .TextMatrix(0, 44) = "CA_CONTR_PERSON_NAME"
   .TextMatrix(0, 45) = "CA_DESIGNATION"
   .TextMatrix(0, 46) = "CA_PINCODE"
   .TextMatrix(0, 47) = "CA_TELEX"
   .TextMatrix(0, 48) = "DATE_OF_BIRTH"
   .TextMatrix(0, 49) = "MEMBERSHIP_TYPE"
   .TextMatrix(0, 50) = "MEMBER_SHIP_ID"
   .TextMatrix(0, 51) = "ORG_CATEGORY"
   .TextMatrix(0, 52) = "ORG_CONSTITUTION"
   .TextMatrix(0, 53) = "PAN_NO"
   .TextMatrix(0, 54) = "PD_AMOUNT"
   .TextMatrix(0, 55) = "PD_APPLICANTS_BANK"
   .TextMatrix(0, 56) = "PD_APP_BANK_BRANCH"
   .TextMatrix(0, 57) = "PD_CLEARING_BANK"
   .TextMatrix(0, 58) = "PD_CL_BANK_BRANCH"
   .TextMatrix(0, 59) = "PD_IFSC_CODE"
   .TextMatrix(0, 60) = "PD_INSTRUMENT_NO"
   .TextMatrix(0, 61) = "PD_INSTRUMENT_TYPE"
   .TextMatrix(0, 62) = "RO_ADDRESS"
   .TextMatrix(0, 63) = "RO_ALTERNATE_TEL"
   .TextMatrix(0, 64) = "RO_CITY"
   .TextMatrix(0, 65) = "RO_EMAIL_ID"
   .TextMatrix(0, 66) = "RO_FAX"
   .TextMatrix(0, 67) = "RO_PINCODE"
   .TextMatrix(0, 68) = "RO_TEL"
   .TextMatrix(0, 69) = "TAN_NO"
   
   .RowHeight(0) = 600
End With
End Sub

Private Sub Grid_Head5()
With MSHFlexGrid5
   .Clear
   .Rows = 2
   .Cols = 6
   
   .TextMatrix(0, 0) = "LOCATION_CODE"
   .TextMatrix(0, 1) = "LOCATION_NAME"
   .TextMatrix(0, 2) = "WH_CODE"
   .TextMatrix(0, 3) = "IS_FREEZED"
   .TextMatrix(0, 4) = "IS_ACTIVE"
   .TextMatrix(0, 5) = "LOCATION_NAME_1"

   .RowHeight(0) = 600
End With
End Sub

Private Sub Grid_Head6()
With MSHFlexGrid6
   .Clear
   .Rows = 2
   .Cols = 81

   .TextMatrix(0, 0) = "WH_CODE"
   .TextMatrix(0, 1) = "WH_SERVICE_PROVIDER"
   .TextMatrix(0, 2) = "WH_NAME"
   .TextMatrix(0, 3) = "LOCATION_CODE"
   .TextMatrix(0, 4) = "WH_LANG_CODE"
   .TextMatrix(0, 5) = "WH_ADDRESS"
   .TextMatrix(0, 6) = "WH_MODE"
   .TextMatrix(0, 7) = "WH_CAPACITY"
   .TextMatrix(0, 8) = "WH_NAME_0"
   .TextMatrix(0, 9) = "WH_ADDRESS_0"
   .TextMatrix(0, 10) = "WH_NAME_1"
   .TextMatrix(0, 11) = "WH_ADDRESS_1"
   .TextMatrix(0, 12) = "IS_FREEZED"
   .TextMatrix(0, 13) = "IS_ACTIVE"
   .TextMatrix(0, 14) = "PLEDGING_ALLOWED"
   .TextMatrix(0, 15) = "AS_ADDRESS"
   .TextMatrix(0, 16) = "AS_ALTERNATE_TEL"
   .TextMatrix(0, 17) = "AS_CITY"
   .TextMatrix(0, 18) = "AS_DESIGNATION"
   .TextMatrix(0, 19) = "AS_EMAIL_ID"
   .TextMatrix(0, 20) = "AS_FAX"
   .TextMatrix(0, 21) = "AS_MOBILE_NO"
   .TextMatrix(0, 22) = "AS_NAME"
   .TextMatrix(0, 23) = "AS_PAN_NAME"
   .TextMatrix(0, 24) = "AS_PINCODE"
   .TextMatrix(0, 25) = "AS_TEL"
   .TextMatrix(0, 26) = "BD_BANK_ACCOUNT_NO"
   .TextMatrix(0, 27) = "BD_BANK_CODE"
   .TextMatrix(0, 28) = "BD_BRANCH_CODE"
   .TextMatrix(0, 29) = "BD_IFSC_CODE"
   .TextMatrix(0, 30) = "BOUND_OR_UNBOUND"
   .TextMatrix(0, 31) = "CA_ALTERNATE_TEL"
   .TextMatrix(0, 32) = "CA_CITY"
   .TextMatrix(0, 33) = "CA_CONTR_PERSON_NAME"
   .TextMatrix(0, 34) = "CA_DESIGNATION"
   .TextMatrix(0, 35) = "CA_EMAIL_ID"
   .TextMatrix(0, 36) = "CA_FAX"
   .TextMatrix(0, 37) = "CA_MOBILE_NO"
   .TextMatrix(0, 38) = "CA_PINCODE"
   .TextMatrix(0, 39) = "CA_TEL"
   .TextMatrix(0, 40) = "CA_TELEX"
   .TextMatrix(0, 41) = "CREATED_BY"
   .TextMatrix(0, 42) = "CREATED_ON"
   .TextMatrix(0, 43) = "DATE_OF_BIRTH"
   .TextMatrix(0, 44) = "FIRST_LOGIN_FLAG"
   .TextMatrix(0, 45) = "MEMBERSHIP_TYPE"
   .TextMatrix(0, 46) = "MEMBER_SHIP_ID"
   .TextMatrix(0, 47) = "MODIFIED_BY"
   .TextMatrix(0, 48) = "MODIFIED_ON"
   .TextMatrix(0, 49) = "ORG_CONSTITUTION"
   .TextMatrix(0, 50) = "PAN_NO"
   .TextMatrix(0, 51) = "PASSWORD"
   .TextMatrix(0, 52) = "PD_AMOUNT"
   .TextMatrix(0, 53) = "PD_APPLICANTS_BANK"
   .TextMatrix(0, 54) = "PD_APP_BANK_BRANCH"
   .TextMatrix(0, 55) = "PD_CLEARING_BANK"
   .TextMatrix(0, 56) = "PD_CL_BANK_BRANCH"
   .TextMatrix(0, 57) = "PD_IFSC_CODE"
   .TextMatrix(0, 58) = "PD_INSTRUMENT_NO"
   .TextMatrix(0, 59) = "PD_INSTRUMENT_TYPE"
   .TextMatrix(0, 60) = "RO_ADDRESS"
   .TextMatrix(0, 61) = "RO_ALTERNATE_TEL"
   .TextMatrix(0, 62) = "RO_CITY"
   .TextMatrix(0, 63) = "RO_EMAIL_ID"
   .TextMatrix(0, 64) = "RO_FAX"
   .TextMatrix(0, 65) = "RO_PINCODE"
   .TextMatrix(0, 66) = "RO_TEL"
   .TextMatrix(0, 67) = "RO_TELEX"
   .TextMatrix(0, 68) = "TAN_NO"
   .TextMatrix(0, 69) = "WH_DOCUMENT_NUMBER"
   .TextMatrix(0, 70) = "WH_DOCUMENT_TYPE"
   .TextMatrix(0, 71) = "WH_LICENSE"
   .TextMatrix(0, 72) = "WH_LICENSE_VALIDITY"
   .TextMatrix(0, 73) = "WH_TYPE"
   .TextMatrix(0, 74) = "WH_NEW_EMAIL"
   .TextMatrix(0, 75) = "WH_PASSWORD"
   .TextMatrix(0, 76) = "WH_PHONE_2"
   .TextMatrix(0, 77) = "WH_CATEGORY"
   .TextMatrix(0, 78) = "INSURANCE_VALID_DATE"
   .TextMatrix(0, 79) = "ORG_CATEGORY"
   .TextMatrix(0, 80) = "EXCHANGE_TYPE"

   .RowHeight(0) = 600
End With
End Sub

Private Sub Grid_Head7()
With MSHFlexGrid7
   .Clear
   .Rows = 2
   .Cols = 66
   
   .TextMatrix(0, 0) = "WSP_ID"
   .TextMatrix(0, 1) = "WSP_NAME"
   .TextMatrix(0, 2) = "WSP_EMAIL"
   .TextMatrix(0, 3) = "WSP_CONTACT"
   .TextMatrix(0, 4) = "WSP_PHONE"
   .TextMatrix(0, 5) = "WSP_PHONE_2"
   .TextMatrix(0, 6) = "WSP_FAX_NO"
   .TextMatrix(0, 7) = "WSP_ADDRESS"
   .TextMatrix(0, 8) = "WSP_DOCUMENT_TYPE"
   .TextMatrix(0, 9) = "WSP_DOCUMENT_NUMBER"
   .TextMatrix(0, 10) = "WSP_PASSWORD"
   .TextMatrix(0, 11) = "PASSWORD"
   .TextMatrix(0, 12) = "FIRST_LOGIN_FLAG"
   .TextMatrix(0, 13) = "CREATED_BY"
   .TextMatrix(0, 14) = "CREATED_ON"
   .TextMatrix(0, 15) = "MODIFIED_BY"
   .TextMatrix(0, 16) = "MODIFIED_ON"
   .TextMatrix(0, 17) = "WSP_NEW_EMAIL"
   .TextMatrix(0, 18) = "ORG_CATEGORY"
   .TextMatrix(0, 19) = "ORG_CONSTITUTION"
   .TextMatrix(0, 20) = "DATE_OF_BIRTH"
   .TextMatrix(0, 21) = "PAN_NO"
   .TextMatrix(0, 22) = "TAN_NO"
   .TextMatrix(0, 23) = "MEMBERSHIP_TYPE"
   .TextMatrix(0, 24) = "RO_ADDRESS"
   .TextMatrix(0, 25) = "RO_TEL"
   .TextMatrix(0, 26) = "RO_CITY"
   .TextMatrix(0, 27) = "RO_FAX"
   .TextMatrix(0, 28) = "RO_PINCODE"
   .TextMatrix(0, 29) = "RO_ALTERNATE_TEL"
   .TextMatrix(0, 30) = "RO_EMAIL_ID"
   .TextMatrix(0, 31) = "CA_CITY"
   .TextMatrix(0, 32) = "CA_PINCODE"
   .TextMatrix(0, 33) = "CA_ALTERNATE_TEL"
   .TextMatrix(0, 34) = "CA_TELEX"
   .TextMatrix(0, 35) = "CA_CONTR_PERSON_NAME"
   .TextMatrix(0, 36) = "CA_DESIGNATION"
   .TextMatrix(0, 37) = "AS_NAME"
   .TextMatrix(0, 38) = "AS_TEL"
   .TextMatrix(0, 39) = "AS_DESIGNATION"
   .TextMatrix(0, 40) = "AS_ALTERNATE_TEL"
   .TextMatrix(0, 41) = "AS_ADDRESS"
   .TextMatrix(0, 42) = "AS_FAX"
   .TextMatrix(0, 43) = "AS_CITY"
   .TextMatrix(0, 44) = "AS_EMAIL_ID"
   .TextMatrix(0, 45) = "AS_PINCODE"
   .TextMatrix(0, 46) = "AS_PAN_NAME"
   .TextMatrix(0, 47) = "AS_MOBILE_NO"
   .TextMatrix(0, 48) = "BD_BANK_CODE"
   .TextMatrix(0, 49) = "BD_BANK_ACCOUNT_NO"
   .TextMatrix(0, 50) = "BD_IFSC_CODE"
   .TextMatrix(0, 51) = "BD_BRANCH_CODE"
   .TextMatrix(0, 52) = "PD_AMOUNT"
   .TextMatrix(0, 53) = "PD_INSTRUMENT_TYPE"
   .TextMatrix(0, 54) = "PD_CLEARING_BANK"
   .TextMatrix(0, 55) = "PD_APPLICANTS_BANK"
   .TextMatrix(0, 56) = "PD_CL_BANK_BRANCH"
   .TextMatrix(0, 57) = "PD_APP_BANK_BRANCH"
   .TextMatrix(0, 58) = "PD_INSTRUMENT_NO"
   .TextMatrix(0, 59) = "PD_IFSC_CODE"
   .TextMatrix(0, 60) = "MEMBER_SHIP_ID"
   .TextMatrix(0, 61) = "IS_FREEZED"
   .TextMatrix(0, 62) = "IS_ACTIVE"
   .TextMatrix(0, 63) = "INSURANCE_AMOUNT"
   .TextMatrix(0, 64) = "INSURANCE_POLICY"
   .TextMatrix(0, 65) = "INSURANCE_VALIDITY"
   
   .RowHeight(0) = 600
End With
End Sub

Private Sub Grid_Head8()
With MSHFlexGrid8
   .Clear
   .Rows = 2
   .Cols = 40
   
   .TextMatrix(0, 0) = "COMMODITY_CODE"
   .TextMatrix(0, 1) = "COMMODITY_GRP"
   .TextMatrix(0, 2) = "COMMODITY_NAME"
   .TextMatrix(0, 3) = "COMMODITY_MKT_UNIT"
   .TextMatrix(0, 4) = "COMMODITY_LANG_CODE"
   .TextMatrix(0, 5) = "COMMODITY_NAME_LANG_LOCAL"
   .TextMatrix(0, 6) = "MAX_EXPIRY_DATE"
   .TextMatrix(0, 7) = "COMMODITY_WH_UOM"
   .TextMatrix(0, 8) = "COMMODITY_CMSE_UOM"
   .TextMatrix(0, 9) = "PRECISION"
   .TextMatrix(0, 10) = "MAX_VALIDITY_DATE"
   .TextMatrix(0, 11) = "ASSAYING_REQUIRED"
   .TextMatrix(0, 12) = "STANDARD_DEDUCTION"
   .TextMatrix(0, 13) = "COMMODITY_CAT"
   .TextMatrix(0, 14) = "WH_CODE"
   .TextMatrix(0, 15) = "PRECISON"
   .TextMatrix(0, 16) = "AGGREGATION_TOLERANCE"
   .TextMatrix(0, 17) = "COMMODITY_NAME_0"
   .TextMatrix(0, 18) = "COMMODITY_NAME_1"
   .TextMatrix(0, 19) = "IS_FREEZED"
   .TextMatrix(0, 20) = "IS_ACTIVE"
   .TextMatrix(0, 21) = "NO_OF_DAYS"
   .TextMatrix(0, 22) = "TOLERANCE"
   .TextMatrix(0, 23) = "DEPOSIT_TYPE"
   .TextMatrix(0, 24) = "STANDARD_DEDUCTION_FOR_REVALID"
   .TextMatrix(0, 25) = "MAX_SAMPLE_QTY"
   .TextMatrix(0, 26) = "BILL_BASIS"
   .TextMatrix(0, 27) = "HZRD_TYPE"
   .TextMatrix(0, 28) = "RATE_UNIT"
   .TextMatrix(0, 29) = "SERVICE_TAX"
   .TextMatrix(0, 30) = "IS_FUNGIBLE"
   .TextMatrix(0, 31) = "EXP_MONTHS"
   .TextMatrix(0, 32) = "VAL_MONTHS_FRM_DEP"
   .TextMatrix(0, 33) = "NO_OF_REVAL"
   .TextMatrix(0, 34) = "FIRST_REVAL_MONTHS"
   .TextMatrix(0, 35) = "SECOND_REVAL_MONTHS"
   .TextMatrix(0, 36) = "THIRD_REVAL_MONTHS"
   .TextMatrix(0, 37) = "DENOMINATOR"
   .TextMatrix(0, 38) = "NUMERATOR"
   .TextMatrix(0, 39) = "COMMODITY_HAIR_CUT_PERCENTAGE"
   
   .RowHeight(0) = 600
End With
End Sub

Private Sub Grid_Head9()
With MSHFlexGrid9
   .Clear
   .Rows = 2
   .Cols = 6
   
   .TextMatrix(0, 0) = "WH_CODE"
   .TextMatrix(0, 1) = "WSP_CODE"
   .TextMatrix(0, 2) = "COMMODITY_CODE"
   .TextMatrix(0, 3) = "WH_COM_CODE"
   .TextMatrix(0, 4) = "IS_FREEZED"
   .TextMatrix(0, 5) = "IS_ACTIVE"
   
   .RowHeight(0) = 600
End With
End Sub

Private Sub Grid_Head10()
With MSHFlexGrid10
   .Clear
   .Rows = 2
   .Cols = 36
   
   
   .TextMatrix(0, 0) = "LocationCode"
   .TextMatrix(0, 1) = "WarehouseCode"
   .TextMatrix(0, 2) = "CommodityCode"
   .TextMatrix(0, 3) = "LotNo"
   .TextMatrix(0, 4) = "ClientID"
   .TextMatrix(0, 5) = "ActualQty"
   .TextMatrix(0, 6) = "NominalQty"
   .TextMatrix(0, 7) = "ActualGrade"
   .TextMatrix(0, 8) = "NominalGrade"
   .TextMatrix(0, 9) = "FED"
   .TextMatrix(0, 10) = "DepositType"
   .TextMatrix(0, 11) = "DepositDate"
   .TextMatrix(0, 12) = "WarehouseUOM"
   .TextMatrix(0, 13) = "GrossQty"
   .TextMatrix(0, 14) = "PackageType"
   .TextMatrix(0, 15) = "PackageSize"
   .TextMatrix(0, 16) = "NetQty"
   .TextMatrix(0, 17) = "WeighingOfficerCode"
   .TextMatrix(0, 18) = "WeighingDate"
   .TextMatrix(0, 19) = "AssayerName"
   .TextMatrix(0, 20) = "AssayerAgency"
   .TextMatrix(0, 21) = "IntAssayerRefno"
   .TextMatrix(0, 22) = "AssayerRefNo"
   .TextMatrix(0, 23) = "SampleNo"
   .TextMatrix(0, 24) = "SampleDate"
   .TextMatrix(0, 25) = "AssayDate"
   .TextMatrix(0, 26) = "RevalidationDate"
   .TextMatrix(0, 27) = "GradeParamName"
   .TextMatrix(0, 28) = "GradeParamValue"
   .TextMatrix(0, 29) = "BayNo"
   .TextMatrix(0, 30) = "StackNo"
   .TextMatrix(0, 31) = "SlotNo"
   .TextMatrix(0, 32) = "AssayerCode"
   .TextMatrix(0, 33) = "ReceiptNo"
   .TextMatrix(0, 34) = "PackageSize1"
   .TextMatrix(0, 35) = "NoofPack"
      
   If OptPending.Value = True Then
      .Cols = 40
      .TextMatrix(0, 36) = "NCDEX BALANCE BAGS"
      .TextMatrix(0, 37) = "NCDEX BALANCE WEIGHT"
      .TextMatrix(0, 38) = "CxTFNo"
      .TextMatrix(0, 39) = "SM_Prefix"
   End If
   .RowHeight(0) = 600
End With
End Sub

Private Sub Grid_Head11()
With MSHFlexGrid11
   .Clear
   .Rows = 2
   .Cols = 16
   
   '.TextMatrix(0, 0) = "TOTALNOOFLOTS"
   '.TextMatrix(0, 1) = "TXNNO"
   .TextMatrix(0, 0) = "DepTxnNumber"
   .TextMatrix(0, 1) = "DepositType"
   .TextMatrix(0, 2) = "ClientId"
   .TextMatrix(0, 3) = "WspId"
   .TextMatrix(0, 4) = "LocationCode"
   .TextMatrix(0, 5) = "WarehouseCode"
   .TextMatrix(0, 6) = "CommodityCode"
   .TextMatrix(0, 7) = "WithdrawerName"
   .TextMatrix(0, 8) = "DocumentType"
   .TextMatrix(0, 9) = "DocumentNumber"
   .TextMatrix(0, 10) = "ContactNumber"
   .TextMatrix(0, 11) = "LotNo"
   .TextMatrix(0, 12) = "ActualQty"
   .TextMatrix(0, 13) = "NominalQty "
   .TextMatrix(0, 14) = "ActualGrade"
   .TextMatrix(0, 15) = "NominalGrade"
   If OptPending.Value = True Then
      .Cols = 20
      .TextMatrix(0, 16) = "NCDEX BALANCE BAGS"
      .TextMatrix(0, 17) = "NCDEX BALANCE WEIGHT"
      '.TextMatrix(0, 18) = "NCDEX BALANCE WEIGHT"
      '.TextMatrix(0, 19) = "NCDEX BALANCE WEIGHT"
   End If
   .RowHeight(0) = 600
End With
End Sub

Private Sub Check1_Click()
If Check1.Value = vbChecked Then
   OptOld.Enabled = False
   OptPending.Enabled = False
   TxtFromDate.Enabled = False
   TxtToDate.Enabled = False
   CmbBatch.Enabled = False
   Check2.Enabled = False
   Check2.Value = vbUnchecked
   CmbBatch.Text = ""
Else
   OptOld.Enabled = True
   OptPending.Enabled = True
   OptPending.Value = True
   Call OptPending_Click
   Check2.Enabled = True
   Check2.Value = vbUnchecked
   'TxtFromDate.Enabled = False
   'TxtToDate.Enabled = False
   'CmbBatch.Enabled = False
End If
End Sub

Private Sub CmdClear_Click()
TxtFromDate.Value = Date - 1
TxtToDate.Value = Date
Call Grid_Head1
Call Grid_Head2
Call Grid_Head3
Call Grid_Head4
Call Grid_Head5
Call Grid_Head6
Call Grid_Head7
Call Grid_Head8
Call Grid_Head9
Call Grid_Head10
Call Grid_Head11
If Check1.Value = vbChecked Then
   Call Check1_Click
   Exit Sub
End If
OptPending.Value = True
Call OptPending_Click
End Sub
Private Sub CmdExport_Click()
If MSHFlexGrid10.Rows = 2 And MSHFlexGrid11.Rows = 2 Then
   MsgBox "No data available to export!!!"
   Exit Sub
End If
Dim mBatchIDXML As Long
Dim mDepositeDate As Variant
Dim mRow As Long
Dim m250Devisor As Integer
Dim mXMLFileName As String

mDepositeDate = Date
mBatchIDXML = mBatchID

MsgBox "Wait till next message!!!"
If Check1.Value = vbUnchecked Then
   If OptPending.Value = True Then
      For i = 1 To MSHFlexGrid10.Rows - 2
         If MSHFlexGrid10.TextMatrix(i, 39) <> "" Then
            tmp = " Update tcgsl Set CommtrackBatchDate = '" & Format(Date, "yyyy-mm-dd") & "', CommtrackBatchID = " & mBatchID
            tmp = tmp & " from txn_cxtf_gsl tcgsl  "
            tmp = tmp & " inner join Txn_CxTF_Details tcd on tcd.cxtfno = tcgsl.cxtfno and tcd.sm_prefix = tcgsl.sm_prefix and tcd.txntype = tcgsl.txntype"
            tmp = tmp & " inner join mst_gsl mgsl on mgsl.GSLID = tcgsl.gslid and mgsl.sm_prefix = tcgsl.sm_prefix"
            tmp = tmp & " where tcd.txntype = 'D' and tcd.flag = 'P' and isnull(tcgsl.commtrackbatchid,0) = 0 and tcd.ExchangeTypeID = 1 and tcd.CxTFDate between '" & Format(TxtFromDate.Value, "yyyy-mm-dd") & "' and '" & Format(TxtToDate.Value, "yyyy-mm-dd") & "'"
            tmp = tmp & " and mgsl.godownid in " & mGodownIDs
            tmp = tmp & " and tcgsl.cxtfno = " & MSHFlexGrid10.TextMatrix(i, 38) & " and tcgsl.sm_prefix = '" & MSHFlexGrid10.TextMatrix(i, 39) & "' "
            Call TmpTable
         End If
      Next
      
      For i = 1 To MSHFlexGrid11.Rows - 2
         If MSHFlexGrid11.TextMatrix(i, 19) <> "" Then
            tmp = " Update tcgsl Set CommtrackBatchDate = '" & Format(Date, "yyyy-mm-dd") & "', CommtrackBatchID = " & mBatchID
            tmp = tmp & " from txn_cxtf_gsl tcgsl  "
            tmp = tmp & " inner join Txn_CxTF_Details tcd on tcd.cxtfno = tcgsl.cxtfno and tcd.sm_prefix = tcgsl.sm_prefix and tcd.txntype = tcgsl.txntype"
            tmp = tmp & " inner join mst_gsl mgsl on mgsl.GSLID = tcgsl.gslid and mgsl.sm_prefix = tcgsl.sm_prefix"
            tmp = tmp & " where tcd.txntype = 'W' and tcd.flag = 'P' and isnull(tcgsl.commtrackbatchid,0) = 0 and tcd.ExchangeTypeID = 1 and tcd.CxTFDate between '" & Format(TxtFromDate.Value, "yyyy-mm-dd") & "' and '" & Format(TxtToDate.Value, "yyyy-mm-dd") & "'"
            tmp = tmp & " and tcgsl.cxtfno = " & MSHFlexGrid11.TextMatrix(i, 18) & " and tcgsl.sm_prefix = '" & MSHFlexGrid11.TextMatrix(i, 19) & "' "
            tmp = tmp & " and mgsl.godownid in " & mGodownIDs
            tmp = tmp & " and convert(varchar, mgsl.GSLID)+'_'+mgsl.SM_Prefix in (" '& mGSLs
            tmp = tmp & " select convert(varchar, mgsl.GSLID)+'_'+mgsl.SM_Prefix" ' GSL, mgsl.DepositeWeight, tcgsl.CxTFNo, sum(tcgsl.grossweight)"
            tmp = tmp & " from mst_gsl mgsl"
            tmp = tmp & " inner join txn_cxtf_gsl tcgsl on mgsl.GSLID = tcgsl.gslid and mgsl.sm_prefix = tcgsl.sm_prefix"
            tmp = tmp & " where tcgsl.txntype = 'W'"
            tmp = tmp & " group by mgsl.gslid, mgsl.sm_prefix, mgsl.DepositeWeight, tcgsl.CxTFNo"
            tmp = tmp & " Having mgsl.DepositeWeight = Sum(tcgsl.grossweight))"
            Call TmpTable
         End If
      Next
      
      For i = 1 To MSHFlexGrid10.Rows - 2
         If MSHFlexGrid10.TextMatrix(i, 39) = "" Then
            tmp = " Update tcgsl Set CommtrackBatchDate = '" & Format(Date, "yyyy-mm-dd") & "', CommtrackBatchID = " & mBatchID
            tmp = tmp & " from Txn_GSL_Transfer  tcgsl"
            tmp = tmp & " inner join mst_gsl mgsl on mgsl.GSLID = tcgsl.newgslid and mgsl.sm_prefix = tcgsl.sm_prefix"
            tmp = tmp & " where mgsl.ExchangeTypeID = 1 and isnull(tcgsl.commtrackbatchid,0) = 0 and tcgsl.GSLTransferDate between '" & Format(TxtFromDate.Value, "yyyy-mm-dd") & "' and '" & Format(TxtToDate.Value, "yyyy-mm-dd") & "'"
            tmp = tmp & " and mgsl.godownid in " & mGodownIDs
            tmp = tmp & " and tcgsl.GSLTransferID = " & MSHFlexGrid10.TextMatrix(i, 38)
            Call TmpTable
         End If
      Next
      For i = 1 To MSHFlexGrid11.Rows - 2
         If MSHFlexGrid11.TextMatrix(i, 19) = "" Then
            tmp = " Update tcgsl Set CommtrackBatchDate = '" & Format(Date, "yyyy-mm-dd") & "', CommtrackBatchID = " & mBatchID
            tmp = tmp & " from Txn_GSL_Transfer  tcgsl"
            tmp = tmp & " inner join mst_gsl mgsl on mgsl.GSLID = tcgsl.newgslid and mgsl.sm_prefix = tcgsl.sm_prefix"
            tmp = tmp & " where mgsl.ExchangeTypeID = 1 and isnull(tcgsl.commtrackbatchid,0) = 0 and tcgsl.GSLTransferDate between '" & Format(TxtFromDate.Value, "yyyy-mm-dd") & "' and '" & Format(TxtToDate.Value, "yyyy-mm-dd") & "'"
            tmp = tmp & " and mgsl.godownid in " & mGodownIDs
            tmp = tmp & " and tcgsl.GSLTransferID = " & MSHFlexGrid11.TextMatrix(i, 18)
            Call TmpTable
         End If
      Next
   Else
      mBatchIDXML = Val(CmbBatch.Text)
      tmp = " Select CommtrackBatchDate from txn_cxtf_gsl where CommtrackBatchID = " & Val(CmbBatch.Text)
      Call TmpTable
      If TmpRs.RecordCount > 0 Then
         mDepositeDate = IIf(IsNull(TmpRs!CommtrackBatchDate), Date, TmpRs!CommtrackBatchDate)
      End If
   End If
Else
   mDepositeDate = "2012-12-25"
   '--Update txn_cxtf_gsl
   tmp = " Update tcgsl Set CommtrackBatchDate = '2012-12-25', CommtrackBatchID = " & mBatchID
   tmp = tmp & " from txn_cxtf_gsl tcgsl"
   tmp = tmp & " inner join Txn_CxTF_Details tcd on tcd.cxtfno = tcgsl.cxtfno and tcd.sm_prefix = tcgsl.sm_prefix and tcd.txntype = tcgsl.txntype"
   tmp = tmp & " inner join mst_gsl mgsl on mgsl.GSLID = tcgsl.gslid and mgsl.sm_prefix = tcgsl.sm_prefix"
   tmp = tmp & " inner join mst_godown mg on mg.GodownID = mgsl.godownid"
   tmp = tmp & " inner join mst_cmp cmp on mg.cmpid = cmp.cmpid"
   tmp = tmp & " inner join mst_location ml on ml.locationid = cmp.locationid"
   tmp = tmp & " inner join mst_commodity mcom on mcom.commodityid = mgsl.commodityid"
   tmp = tmp & " where MGSL.ExchangeTypeID = 1 And MGSL.Flag = 'O' and mgsl.godownid in"
   tmp = tmp & " (Select Distinct GodownID from Mst_GSL Where Mst_GSL.ExchangeTypeID = 2 and flag not in ('E', 'J', 'Z') and balancebags > 0 )"
   tmp = tmp & " and isnull(mg.ncdexcode,'') <> '' and mgsl.CDTFDate <= '2012-12-25' and mgsl.BalanceWeight > 0"
   tmp = tmp & " and tcd.CxTFDate <= '2012-12-25' and isnull(tcgsl.CommtrackBatchID,0) = 0 "
   'tmp = tmp & " and convert(varchar, mgsl.gslid)+'_'+mgsl.sm_prefix not in (select distinct convert(varchar, gslid)+'_'+sm_prefix from txn_cxtf_gsl where isnull(commtrackbatchid, 0) <> 0)"
   Call TmpTable
   
   '--Update Txn_GSL_Transfer
   tmp = " Update tcgsl Set CommtrackBatchDate = '2012-12-25', CommtrackBatchID = " & mBatchID
   tmp = tmp & " from Txn_GSL_Transfer tcgsl"
   tmp = tmp & " inner join mst_gsl mgsl on (mgsl.GSLID = tcgsl.newgslid or mgsl.GSLID = tcgsl.oldgslid) and mgsl.sm_prefix = tcgsl.sm_prefix"
   tmp = tmp & " inner join mst_godown mg on mg.GodownID = mgsl.godownid"
   tmp = tmp & " inner join mst_cmp cmp on mg.cmpid = cmp.cmpid"
   tmp = tmp & " inner join mst_location ml on ml.locationid = cmp.locationid"
   tmp = tmp & " inner join mst_commodity mcom on mcom.commodityid = mgsl.commodityid"
   tmp = tmp & " where MGSL.ExchangeTypeID = 1 And MGSL.Flag = 'O' and mgsl.godownid in"
   tmp = tmp & " (Select Distinct GodownID from Mst_GSL Where Mst_GSL.ExchangeTypeID = 2 and flag not in ('E', 'J', 'Z') and balancebags > 0 )"
   tmp = tmp & " and isnull(mg.ncdexcode,'') <> '' and mgsl.CDTFDate <= '2012-12-25' and mgsl.BalanceWeight > 0"
   tmp = tmp & " and tcgsl.GSLTransferDate <= '2012-12-25' and isnull(tcgsl.CommtrackBatchID,0) = 0"
   'tmp = tmp & " and convert(varchar, mgsl.gslid)+'_'+mgsl.sm_prefix not in (select distinct convert(varchar, gslid)+'_'+sm_prefix from txn_cxtf_gsl where isnull(commtrackbatchid, 0) <> 0)"
   Call TmpTable
End If
Dim mXMLStart, mXMLEnd, mXMLDetail  As String
Dim oFs As New FileSystemObject

'mXMLStart = "<?xml version=""1.0"" standalone=""yes"" ?>" & vbCrLf
'mXMLStart = mXMLStart & "<SuccessReceipts>" & vbCrLf
'mXMLStart = mXMLStart & "<HeaderTable>" & vbCrLf
'mXMLStart = mXMLStart & vbTab & "<CMSPID>NCMSLWHMKR</CMSPID>" & vbCrLf
'mXMLStart = mXMLStart & vbTab & "<TotalNoOfLots>1</TotalNoOfLots>" & vbCrLf
'mXMLStart = mXMLStart & vbTab & "<DepositDate>" & Format(CDate(mDepositeDate), "yyyy-mm-dd hh:mm:ss") & "</DepositDate>" & vbCrLf
'mXMLStart = mXMLStart & vbTab & "<TxnNo>NCMSL0001</TxnNo>" & vbCrLf '" & mBatchID & "
'mXMLStart = mXMLStart & "</HeaderTable>" & vbCrLf
'mXMLStart = mXMLStart & vbTab & "<DetailTable>" & vbCrLf

mXMLEnd = vbTab & "</DetailTable>" & vbCrLf
mXMLEnd = mXMLEnd & "</SuccessReceipts>"

'--------------for deposit data
mXMLDetail = ""
ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid10.Rows + 3

m250Devisor = 1
'If CmbBatch.Text = "" Then
'   mFileName = App.Path & "\Excel\CommtrackDeposit" & mBatchID & "0" & m250Devisor & ".xml"
'Else
'   mFileName = App.Path & "\Excel\CommtrackDeposit" & CmbBatch.Text & "0" & m250Devisor & ".xml"
'End If

For mRow = 1 To MSHFlexGrid10.Rows - 2
   mXMLDetail = mXMLDetail & vbTab & vbTab & "<DetailTableDataRow>" & vbCrLf
   If OptPending.Value = True Then
      For c = 0 To MSHFlexGrid10.Cols - 5
         mXMLDetail = mXMLDetail & vbTab & vbTab & vbTab & " <" & Replace(Trim(MSHFlexGrid10.TextMatrix(0, c)), " ", "", 1) & ">" & Trim(MSHFlexGrid10.TextMatrix(mRow, c)) & "</" & Replace(Trim(MSHFlexGrid10.TextMatrix(0, c)), " ", "", 1) & ">" & vbCrLf
      Next
   ElseIf OptOld.Value = True Then
      For c = 0 To MSHFlexGrid10.Cols - 1
         mXMLDetail = mXMLDetail & vbTab & vbTab & vbTab & " <" & Replace(Trim(MSHFlexGrid10.TextMatrix(0, c)), " ", "", 1) & ">" & Trim(MSHFlexGrid10.TextMatrix(mRow, c)) & "</" & Replace(Trim(MSHFlexGrid10.TextMatrix(0, c)), " ", "", 1) & ">" & vbCrLf
      Next
   End If
   mXMLDetail = mXMLDetail & vbTab & vbTab & "</DetailTableDataRow>" & vbCrLf
   ProgressBar1.Value = ProgressBar1.Value + 1
   
   If mRow = 250 * m250Devisor Then
      mTxnNo = "NCMSL" & mBatchIDXML & "0" & m250Devisor
      mXMLFileName = Year(CDate(mDepositeDate)) & "_" & Day(CDate(mDepositeDate)) & "_" & Month(CDate(mDepositeDate)) & "_" & mTxnNo & "_Deposit"
      mFileName = App.Path & "\Excel\" & mXMLFileName & ".xml"
      
      oFs.CreateTextFile mFileName, True

      mXMLStart = "<?xml version=""1.0"" standalone=""yes"" ?>" & vbCrLf
      mXMLStart = mXMLStart & "<SuccessReceipts>" & vbCrLf
      mXMLStart = mXMLStart & "<HeaderTable>" & vbCrLf
      'mXMLStart = mXMLStart & vbTab & "<CMSPID>NCMSLWHMKR</CMSPID>" & vbCrLf
      mXMLStart = mXMLStart & vbTab & "<CMSEID></CMSEID>" & vbCrLf
      mXMLStart = mXMLStart & vbTab & "<TotalNoOfLots>250</TotalNoOfLots>" & vbCrLf
      mXMLStart = mXMLStart & vbTab & "<DepositDate>" & Format(CDate(mDepositeDate), "yyyy-mm-dd hh:mm:ss") & "</DepositDate>" & vbCrLf
      mXMLStart = mXMLStart & vbTab & "<TxnNo>" & mTxnNo & "</TxnNo>" & vbCrLf '" & mBatchID & "
      mXMLStart = mXMLStart & "</HeaderTable>" & vbCrLf
      mXMLStart = mXMLStart & vbTab & "<DetailTable>" & vbCrLf

      Open mFileName For Output As #1
         Print #1, mXMLStart & mXMLDetail & mXMLEnd
      Close #1
      mXMLDetail = ""
      If mRow <> MSHFlexGrid10.Rows - 2 Then
         m250Devisor = m250Devisor + 1
         'If CmbBatch.Text = "" Then
         '   mFileName = App.Path & "\Excel\CommtrackDeposit" & mBatchID & "0" & m250Devisor & ".xml"
         'Else
         '   mFileName = App.Path & "\Excel\CommtrackDeposit" & CmbBatch.Text & "0" & m250Devisor & ".xml"
         'End If
         'oFs.CreateTextFile mFileName, True
      End If
   ElseIf mRow = MSHFlexGrid10.Rows - 2 Then
      mTxnNo = "NCMSL" & mBatchIDXML & "0" & m250Devisor
      mXMLFileName = Year(CDate(mDepositeDate)) & "_" & Day(CDate(mDepositeDate)) & "_" & Month(CDate(mDepositeDate)) & "_" & mTxnNo & "_Deposit"
      mFileName = App.Path & "\Excel\" & mXMLFileName & ".xml"
      
      oFs.CreateTextFile mFileName, True

      mXMLStart = "<?xml version=""1.0"" standalone=""yes"" ?>" & vbCrLf
      mXMLStart = mXMLStart & "<SuccessReceipts>" & vbCrLf
      mXMLStart = mXMLStart & "<HeaderTable>" & vbCrLf
      'mXMLStart = mXMLStart & vbTab & "<CMSPID>NCMSLWHMKR</CMSPID>" & vbCrLf
      mXMLStart = mXMLStart & vbTab & "<CMSEID></CMSEID>" & vbCrLf
      mXMLStart = mXMLStart & vbTab & "<TotalNoOfLots>" & MSHFlexGrid10.Rows - 2 - (250 * (m250Devisor - 1)) & "</TotalNoOfLots>" & vbCrLf
      mXMLStart = mXMLStart & vbTab & "<DepositDate>" & Format(CDate(mDepositeDate), "yyyy-mm-dd hh:mm:ss") & "</DepositDate>" & vbCrLf
      mXMLStart = mXMLStart & vbTab & "<TxnNo>" & mTxnNo & "</TxnNo>" & vbCrLf '" & mBatchID & "
      mXMLStart = mXMLStart & "</HeaderTable>" & vbCrLf
      mXMLStart = mXMLStart & vbTab & "<DetailTable>" & vbCrLf
      
      Open mFileName For Output As #1
         Print #1, mXMLStart & mXMLDetail & mXMLEnd
      Close #1
   End If
Next
ProgressBar1.Value = ProgressBar1.Max

m250Devisor = 1
'---------------for withdraw data
mXMLDetail = ""
ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid11.Rows + 3
'If CmbBatch.Text = "" Then
'   mFileName = App.Path & "\Excel\CommtrackWithdraw" & mBatchID & "0" & m250Devisor & ".xml"
'Else
'   mFileName = App.Path & "\Excel\CommtrackWithdraw" & CmbBatch.Text & "0" & m250Devisor & ".xml"
'End If
'oFs.CreateTextFile mFileName, True

For mRow = 1 To MSHFlexGrid11.Rows - 2
   mXMLDetail = mXMLDetail & vbTab & vbTab & "<DetailTableDataRow>" & vbCrLf
   If OptPending.Value = True Then
      For c = 0 To MSHFlexGrid11.Cols - 5
         mXMLDetail = mXMLDetail & vbTab & vbTab & vbTab & " <" & Replace(Trim(MSHFlexGrid11.TextMatrix(0, c)), " ", "", 1) & ">" & Trim(MSHFlexGrid11.TextMatrix(mRow, c)) & "</" & Replace(Trim(MSHFlexGrid11.TextMatrix(0, c)), " ", "", 1) & ">" & vbCrLf
      Next
   ElseIf OptOld.Value = True Then
      For c = 0 To MSHFlexGrid11.Cols - 1
         mXMLDetail = mXMLDetail & vbTab & vbTab & vbTab & " <" & Replace(Trim(MSHFlexGrid11.TextMatrix(0, c)), " ", "", 1) & ">" & Trim(MSHFlexGrid11.TextMatrix(mRow, c)) & "</" & Replace(Trim(MSHFlexGrid11.TextMatrix(0, c)), " ", "", 1) & ">" & vbCrLf
      Next
   End If
   mXMLDetail = mXMLDetail & vbTab & vbTab & "</DetailTableDataRow>" & vbCrLf
   ProgressBar1.Value = ProgressBar1.Value + 1
   
   If mRow = 250 * m250Devisor Then
      mTxnNo = "NCMSLW" & mBatchIDXML & "0" & m250Devisor
      mXMLFileName = Year(CDate(mDepositeDate)) & "_" & Day(CDate(mDepositeDate)) & "_" & Month(CDate(mDepositeDate)) & "_" & mTxnNo & "_Withdraw"
      mFileName = App.Path & "\Excel\" & mXMLFileName & ".xml"
      
      oFs.CreateTextFile mFileName, True

      mXMLStart = "<?xml version=""1.0"" standalone=""yes"" ?>" & vbCrLf
      mXMLStart = mXMLStart & "<SuccessReceipts>" & vbCrLf
      mXMLStart = mXMLStart & "<HeaderTable>" & vbCrLf
      'mXMLStart = mXMLStart & vbTab & "<CMSPID>NCMSLWHMKR</CMSPID>" & vbCrLf
      mXMLStart = mXMLStart & vbTab & "<TotalNoOfLots>250</TotalNoOfLots>" & vbCrLf
      'mXMLStart = mXMLStart & vbTab & "<DepositDate>" & Format(CDate(mDepositeDate), "yyyy-mm-dd hh:mm:ss") & "</DepositDate>" & vbCrLf
      mXMLStart = mXMLStart & vbTab & "<TxnNo>" & mTxnNo & "</TxnNo>" & vbCrLf '" & mBatchID & "
      mXMLStart = mXMLStart & "</HeaderTable>" & vbCrLf
      mXMLStart = mXMLStart & vbTab & "<DetailTable>" & vbCrLf
      
      Open mFileName For Output As #1
         Print #1, mXMLStart & mXMLDetail & mXMLEnd
      Close #1
      mXMLDetail = ""
      If mRow <> MSHFlexGrid11.Rows - 2 Then
         m250Devisor = m250Devisor + 1
         'If CmbBatch.Text = "" Then
         '   mFileName = App.Path & "\Excel\CommtrackWithdraw" & mBatchID & "0" & m250Devisor & ".xml"
         'Else
         '   mFileName = App.Path & "\Excel\CommtrackWithdraw" & CmbBatch.Text & "0" & m250Devisor & ".xml"
         'End If
         'oFs.CreateTextFile mFileName, True
      End If
   ElseIf mRow = MSHFlexGrid11.Rows - 2 Then
      mTxnNo = "NCMSLW" & mBatchIDXML & "0" & m250Devisor
      mXMLFileName = Year(CDate(mDepositeDate)) & "_" & Day(CDate(mDepositeDate)) & "_" & Month(CDate(mDepositeDate)) & "_" & mTxnNo & "_Withdraw"
      mFileName = App.Path & "\Excel\" & mXMLFileName & ".xml"
      
      oFs.CreateTextFile mFileName, True

      mXMLStart = "<?xml version=""1.0"" standalone=""yes"" ?>" & vbCrLf
      mXMLStart = mXMLStart & "<SuccessReceipts>" & vbCrLf
      mXMLStart = mXMLStart & "<HeaderTable>" & vbCrLf
      'mXMLStart = mXMLStart & vbTab & "<CMSPID>NCMSLWHMKR</CMSPID>" & vbCrLf
      mXMLStart = mXMLStart & vbTab & "<TotalNoOfLots>" & MSHFlexGrid11.Rows - 2 - (250 * (m250Devisor - 1)) & "</TotalNoOfLots>" & vbCrLf
      'mXMLStart = mXMLStart & vbTab & "<DepositDate>" & Format(CDate(mDepositeDate), "yyyy-mm-dd hh:mm:ss") & "</DepositDate>" & vbCrLf
      mXMLStart = mXMLStart & vbTab & "<TxnNo>" & mTxnNo & "</TxnNo>" & vbCrLf '" & mBatchID & "
      mXMLStart = mXMLStart & "</HeaderTable>" & vbCrLf
      mXMLStart = mXMLStart & vbTab & "<DetailTable>" & vbCrLf
      Open mFileName For Output As #1
         Print #1, mXMLStart & mXMLDetail & mXMLEnd
      Close #1
   End If
Next

ProgressBar1.Value = ProgressBar1.Max
If OptPending.Value = True Or Check1.Value = vbChecked Then
   CmbBatch.AddItem mBatchID
   mBatchID = mBatchID + 1
End If
'Check1.Value = vbUnchecked
'Check1.Visible = False
CmdExport.Enabled = False
MsgBox "Data Successfully Converted into XML!!!", vbInformation
End Sub

Private Sub CmdGetReport_Click()
Call Grid_Head1
Call Grid_Head2
Call Grid_Head3
Call Grid_Head4
Call Grid_Head5
Call Grid_Head6
Call Grid_Head7
Call Grid_Head8
Call Grid_Head9
Call Grid_Head10
Call Grid_Head11
MsgBox "Wait till next message!!!"

tmp = " select godownid, sum(Balancebags) BalanceBags, sum(balanceweight) balanceweight from mst_gsl where exchangetypeid = 2 and flag not in ('E', 'J', 'Z') and balancebags > 0 group by godownid "
Call Tmp2Table
If TmpRs2.RecordCount > 0 Then
   mGodownIDs = "("
   For k = 1 To TmpRs2.RecordCount
      mGodownIDs = mGodownIDs & IIf(IsNull(TmpRs2!GodownID), 0, TmpRs2!GodownID) & " ,"
      TmpRs2.MoveNext
   Next
   mGodownIDs = mID(mGodownIDs, 1, Len(mGodownIDs) - 1) & " )"
End If

If Check1.Value = vbUnchecked Then
   If OptPending.Value = True Then
      If Format(TxtFromDate.Value, "yyyy-mm-dd") > Format(Date, "yyyy-mm-dd") Then
         MsgBox "Future month not allowed!!!"
         TxtFromDate.SetFocus
         Exit Sub
      End If
      If Format(TxtToDate.Value, "yyyy-mm-dd") > Format(Date, "yyyy-mm-dd") Then
         MsgBox "Future month not allowed!!!"
         TxtToDate.SetFocus
         Exit Sub
      End If
      If TxtFromDate.Value > TxtToDate.Value Then
         MsgBox "From Date Should Be Less Than To Date!!!"
         TxtFromDate.SetFocus
         Exit Sub
      End If
   
   ElseIf OptOld.Value = True Then
      If Trim(CmbBatch.Text) = "" Then
         MsgBox "Blank BatchID not allowed!!!"
         CmbBatch.SetFocus
         Exit Sub
      End If
   End If
End If
'Call Get_UOM_MSTR
CmdExcel.Enabled = True
CmdExport.Enabled = True
'Call Get_CLIENT_MSTR
'Call Get_USER_MSTR
'Call Get_ASS_MSTR
'Call Get_LOCATION_MSTR
'Call Get_WH_MSTR
'Call Get_WSP_MSTR
'Call Get_COMMODITY_MSTR
'Call Get_WH_COMMODITY_MAP
If Check1.Value = vbUnchecked Then
   Call Get_DEPOSITEDATA
   Call Get_WITHDRAWDATA
Else
   Call Get_OpeningDEPOSITEDATA
End If
MsgBox "Done!!!"

End Sub


Private Sub Get_WITHDRAWDATA()

tmp = "Select "
''' TOTALNOOFLOTS, convert(varchar, tcd.cxtfno) TXNNO,
tmp = tmp & " mgsl.CDTFNo 'DEPTXNNUMBER', 'UnAssayed' DEPOSITETYPE, '375' CLIENTID, '' WSPID,"
'tmp = tmp & " upper(tcgsl.sm_prefix) LOCATION_CODE, upper(cmp.cmpname)+'_'+upper(mg.godownno) WH_CODE, upper(mcom.Commodity) COMMODITY_CODE,"
tmp = tmp & " upper(ml.ncdexcode) LOCATION_CODE, upper(mg.ncdexcode) WH_CODE, upper(mcom.ncdexcode) COMMODITY_CODE,"
tmp = tmp & " '' WITHDRAWERNAME, '' DOCUMENTTYPE, '' DOCUMENTNUMBER, '' CONTACTNUMBER,"
tmp = tmp & " '' LOTNO,"
tmp = tmp & " sum(tcgsl.Grossweight) ACTUALQTY, sum(tcgsl.Dematweight) NOMINALQTY, 'DEFAULT' ACTUALGRADE, 'DEFAULT' NOMINALGRADE, mgsl.godownid, tcgsl.cxtfno, tcgsl.sm_prefix"
tmp = tmp & " from txn_cxtf_gsl tcgsl  "
tmp = tmp & " inner join Txn_CxTF_Details tcd on tcd.cxtfno = tcgsl.cxtfno and tcd.sm_prefix = tcgsl.sm_prefix and tcd.txntype = tcgsl.txntype"
tmp = tmp & " inner join mst_gsl mgsl on mgsl.GSLID = tcgsl.gslid and mgsl.sm_prefix = tcgsl.sm_prefix"
tmp = tmp & " inner join mst_godown mg on mg.GodownID = mgsl.godownid"
tmp = tmp & " inner join mst_cmp cmp on mg.cmpid = cmp.cmpid"
tmp = tmp & " inner join mst_location ml on ml.locationid= cmp.locationid"
tmp = tmp & " inner join mst_commodity mcom on mcom.commodityid = tcd.commodityid"
If CmbBatch.Text = "" Then
   tmp = tmp & " where tcd.txntype = 'W' and isnull(tcgsl.commtrackbatchid,0) = 0 and tcd.flag = 'P'  and tcd.ExchangeTypeID = 1 and tcd.CxTFDate between '" & Format(TxtFromDate.Value, "yyyy-mm-dd") & "' and '" & Format(TxtToDate.Value, "yyyy-mm-dd") & "'"
   tmp = tmp & " and mgsl.godownid in " & mGodownIDs
   tmp = tmp & " and convert(varchar, mgsl.GSLID)+'_'+mgsl.SM_Prefix in (" '& mGSLs
   tmp = tmp & " select convert(varchar, mgsl.GSLID)+'_'+mgsl.SM_Prefix" ' GSL, mgsl.DepositeWeight, tcgsl.CxTFNo, sum(tcgsl.grossweight)"
   tmp = tmp & " from mst_gsl mgsl"
   tmp = tmp & " inner join txn_cxtf_gsl tcgsl on mgsl.GSLID = tcgsl.gslid and mgsl.sm_prefix = tcgsl.sm_prefix"
   tmp = tmp & " where tcgsl.txntype = 'W'"
   tmp = tmp & " group by mgsl.gslid, mgsl.sm_prefix, mgsl.DepositeWeight, tcgsl.CxTFNo"
   tmp = tmp & " Having mgsl.DepositeWeight = Sum(tcgsl.grossweight))"
Else
   tmp = tmp & " where tcd.txntype = 'W' and tcgsl.commtrackbatchid = " & CmbBatch.Text
End If
tmp = tmp & " group by mgsl.CDTFNo, ml.ncdexcode, mg.ncdexcode, mcom.ncdexcode, mgsl.StackNo, mgsl.LotNo, mgsl.godownid, tcgsl.cxtfno, tcgsl.sm_prefix"
tmp = tmp & " Union"
''' TOTALNOOFLOTS, '' TXNNO,
tmp = tmp & " select"
tmp = tmp & " mgsl.CDTFNo 'DEPTXNNUMBER', 'UnAssayed' DEPOSITETYPE, '375' CLIENTID, '' WSPID,"
tmp = tmp & " upper(ml.ncdexcode) LOCATION_CODE, upper(mg.ncdexcode) WH_CODE, upper(mcom.ncdexcode) COMMODITY_CODE,"
tmp = tmp & " '' WITHDRAWERNAME, '' DOCUMENTTYPE, '' DOCUMENTNUMBER, '' CONTACTNUMBER,"
tmp = tmp & " '' LOTNO,"
tmp = tmp & " sum(tcgsl.TransferedWeight) ACTUALQTY, sum(tcgsl.TransferedWeight) NOMINALQTY, 'DEFAULT' ACTUALGRADE, 'DEFAULT' NOMINALGRADE, mgsl.godownid, tcgsl.GSLTransferID cxtfno, '' SM_Prefix "
tmp = tmp & " from Txn_GSL_Transfer  tcgsl"
tmp = tmp & " inner join mst_gsl mgsl on mgsl.GSLID = tcgsl.oldgslid and mgsl.sm_prefix = tcgsl.sm_prefix"
tmp = tmp & " inner join mst_godown mg on mg.GodownID = mgsl.godownid"
tmp = tmp & " inner join mst_cmp cmp on mg.cmpid = cmp.cmpid"
tmp = tmp & " inner join mst_location ml on ml.locationid= cmp.locationid"
tmp = tmp & " inner join mst_commodity mcom on mcom.commodityid = mgsl.commodityid"
If CmbBatch.Text = "" Then
   tmp = tmp & " where mgsl.ExchangeTypeID = 1 and isnull(tcgsl.commtrackbatchid,0) = 0 and tcgsl.GSLTransferDate between '" & Format(TxtFromDate.Value, "yyyy-mm-dd") & "' and '" & Format(TxtToDate.Value, "yyyy-mm-dd") & "'"
   tmp = tmp & " and mgsl.godownid in " & mGodownIDs
   tmp = tmp & " and convert(varchar, mgsl.GSLID)+'_'+mgsl.SM_Prefix in (" '& mGSLs
   tmp = tmp & " select convert(varchar, mgsl.GSLID)+'_'+mgsl.SM_Prefix" ' GSL, mgsl.DepositeWeight, tcgsl.CxTFNo, sum(tcgsl.grossweight)"
   tmp = tmp & " from mst_gsl mgsl"
   tmp = tmp & " inner join txn_cxtf_gsl tcgsl on mgsl.GSLID = tcgsl.gslid and mgsl.sm_prefix = tcgsl.sm_prefix"
   tmp = tmp & " where tcgsl.txntype = 'W'"
   tmp = tmp & " group by mgsl.gslid, mgsl.sm_prefix, mgsl.DepositeWeight, tcgsl.CxTFNo"
   tmp = tmp & " Having mgsl.DepositeWeight = Sum(tcgsl.grossweight))"
Else
   tmp = tmp & " where tcgsl.commtrackbatchid = " & CmbBatch.Text
   tmp = tmp & " and convert(varchar, mgsl.GSLID)+'_'+mgsl.SM_Prefix in (" '& mGSLs
   tmp = tmp & " select convert(varchar, mgsl.GSLID)+'_'+mgsl.SM_Prefix" ' GSL, mgsl.DepositeWeight, tcgsl.CxTFNo, sum(tcgsl.grossweight)"
   tmp = tmp & " from mst_gsl mgsl"
   tmp = tmp & " inner join txn_cxtf_gsl tcgsl on mgsl.GSLID = tcgsl.gslid and mgsl.sm_prefix = tcgsl.sm_prefix"
   tmp = tmp & " where tcgsl.txntype = 'W'"
   tmp = tmp & " group by mgsl.gslid, mgsl.sm_prefix, mgsl.DepositeWeight, tcgsl.CxTFNo"
   tmp = tmp & " Having mgsl.DepositeWeight = Sum(tcgsl.grossweight))"
End If
tmp = tmp & " group by mgsl.CDTFNo, ml.ncdexcode, mg.ncdexcode, mcom.ncdexcode, mgsl.StackNo, mgsl.LotNo, mgsl.godownid, tcgsl.GSLTransferID"
Call TmpTable

If TmpRs.RecordCount > 0 Then
   'CmdExport.Enabled = True
   ProgressBar1.Value = 0
   ProgressBar1.Max = TmpRs.RecordCount + 3
   With MSHFlexGrid11
      For k = 1 To TmpRs.RecordCount
         i = .Rows - 1
         If Check2.Value = vbChecked Then
            If IIf(IsNull(TmpRs!LOCATION_CODE), "", TmpRs!LOCATION_CODE) = "" Or IIf(IsNull(TmpRs!WH_CODE), "", TmpRs!WH_CODE) = "" Or IIf(IsNull(TmpRs!COMMODITY_CODE), "", TmpRs!COMMODITY_CODE) = "" Then
               GoTo NextRow
            End If
         End If
         '.TextMatrix(I, 0) = IIf(IsNull(TmpRs!TOTALNOOFLOTS), "", TmpRs!TOTALNOOFLOTS)
         '.TextMatrix(I, 1) = IIf(IsNull(TmpRs!TXNNO), "", TmpRs!TXNNO)
         .TextMatrix(i, 0) = IIf(IsNull(TmpRs!DEPTXNNUMBER), "", TmpRs!DEPTXNNUMBER)
         .TextMatrix(i, 1) = IIf(IsNull(TmpRs!DEPOSITETYPE), "", TmpRs!DEPOSITETYPE)
         .TextMatrix(i, 2) = IIf(IsNull(TmpRs!ClientID), "", TmpRs!ClientID)
         .TextMatrix(i, 3) = IIf(IsNull(TmpRs!WSPID), "", TmpRs!WSPID)
         .TextMatrix(i, 4) = IIf(IsNull(TmpRs!LOCATION_CODE), "", TmpRs!LOCATION_CODE)
         If CmdExport.Enabled = True And Trim(.TextMatrix(i, 4)) = "" Then
            CmdExport.Enabled = False
         End If
         .TextMatrix(i, 5) = IIf(IsNull(TmpRs!WH_CODE), "", TmpRs!WH_CODE)
         If CmdExport.Enabled = True And Trim(.TextMatrix(i, 5)) = "" Then
            CmdExport.Enabled = False
         End If
         .TextMatrix(i, 6) = IIf(IsNull(TmpRs!COMMODITY_CODE), "", TmpRs!COMMODITY_CODE)
         If CmdExport.Enabled = True And Trim(.TextMatrix(i, 6)) = "" Then
            CmdExport.Enabled = False
         End If
         .TextMatrix(i, 7) = IIf(IsNull(TmpRs!WITHDRAWERNAME), "", TmpRs!WITHDRAWERNAME)
         .TextMatrix(i, 8) = IIf(IsNull(TmpRs!DocumentType), "", TmpRs!DocumentType)
         .TextMatrix(i, 9) = IIf(IsNull(TmpRs!DOCUMENTNUMBER), "", TmpRs!DOCUMENTNUMBER)
         .TextMatrix(i, 10) = IIf(IsNull(TmpRs!CONTACTNUMBER), "", TmpRs!CONTACTNUMBER)
         .TextMatrix(i, 11) = IIf(IsNull(TmpRs!LotNo), "", TmpRs!LotNo)
         .TextMatrix(i, 12) = IIf(IsNull(TmpRs!ActualQty), "", TmpRs!ActualQty)
         .TextMatrix(i, 13) = IIf(IsNull(TmpRs!NOMINALQTY), "", TmpRs!NOMINALQTY)
         .TextMatrix(i, 14) = IIf(IsNull(TmpRs!ACTUALGRADE), "", TmpRs!ACTUALGRADE)
         .TextMatrix(i, 15) = IIf(IsNull(TmpRs!NOMINALGRADE), "", TmpRs!NOMINALGRADE)
         If OptPending.Value = True Then
            If TmpRs2.RecordCount > 0 Then
               TmpRs2.MoveFirst
               TmpRs2.Find "GodownID = " & TmpRs!GodownID
               If Not TmpRs2.EOF Then
                  .TextMatrix(i, 16) = IIf(IsNull(TmpRs2!BalanceBags), 0, TmpRs2!BalanceBags)
                  .TextMatrix(i, 17) = IIf(IsNull(TmpRs2!BalanceWeight), 0, TmpRs2!BalanceWeight)
               End If
            End If
            .TextMatrix(i, 18) = IIf(IsNull(TmpRs!CxTFNO), "", TmpRs!CxTFNO)
            .TextMatrix(i, 19) = IIf(IsNull(TmpRs!SM_Prefix), "", TmpRs!SM_Prefix)
         End If
         .Rows = .Rows + 1
NextRow:
         ProgressBar1.Value = ProgressBar1.Value + 1
         TmpRs.MoveNext
      Next
      ProgressBar1.Value = ProgressBar1.Max
   End With
End If
End Sub

Private Sub Get_DEPOSITEDATA()
tmp = " Select"
tmp = tmp & " upper(ml.NCDEXCODE) LOCATION_CODE, upper(mg.ncdexcode) WH_CODE, upper(mcom.ncdexcode) COMMODITY_CODE,"
tmp = tmp & " '' LOTNO, '375' CLIENTID,"
tmp = tmp & " sum(tcgsl.Grossweight) ACTUALQTY, sum(tcgsl.Dematweight) NOMINALQTY, 'DEFAULT' ACTUALGRADE, 'DEFAULT' NOMINALGRADE, '' FED, 'UnAssayed' DEPOSITETYPE,"
tmp = tmp & " tcd.CxTFDate DEPOSITEDATE, 'MT' WAREHOUSEUOM , sum(tcgsl.Grossweight) GROSSQTY, 'Bags' PACKAGETYPE, '' PACKAGESIZE,"
tmp = tmp & " sum(tcgsl.Dematweight) NETQTY, '' WEIGHINGOFFICERCODE, '' WEIGHINGDATE, '' ASSAYERNAME, '' ASSAYERAGENCY, '' INTASSAYERREFNO,"
tmp = tmp & " '' ASSAYERREFNO, '' SAMPLENO, '' SAMPLEDATE, '' ASSAYDATE, '' REVALIDATIONDATE, '' GRADEPARAMNAME, '' GRADEPARAMVALUE,"
'tmp = tmp & " upper(mg.ncdexcode) BAYNO, upper(mgsl.stackNo) STACKNO, upper(mg.ncdexcode)+'~'+upper(mgsl.StackNo)+'~'+upper(mgsl.LotNo) SLOTNO, '' ASSAYERCODE,"
tmp = tmp & " upper(mg.ncdexcode) BAYNO, upper(mgsl.stackNo) STACKNO, upper(mgsl.LotNo) SLOTNO, '' ASSAYERCODE,"
tmp = tmp & " 'D'+convert(varchar, mgsl.GSLID)+upper(mgsl.SM_prefix)+convert(varchar, tcgsl.CxtfNo) RECEIPTNO, '' PACKAGESIZE1, sum(tcgsl.NoOfBags) NOOFPACK, mgsl.godownid, tcgsl.CxtfNo, tcgsl.SM_Prefix"
tmp = tmp & " from txn_cxtf_gsl tcgsl "
tmp = tmp & " inner join Txn_CxTF_Details tcd on tcd.cxtfno = tcgsl.cxtfno and tcd.sm_prefix = tcgsl.sm_prefix and tcd.txntype = tcgsl.txntype"
tmp = tmp & " inner join mst_gsl mgsl on mgsl.GSLID = tcgsl.gslid and mgsl.sm_prefix = tcgsl.sm_prefix"
tmp = tmp & " inner join mst_godown mg on mg.GodownID = mgsl.godownid"
tmp = tmp & " inner join mst_cmp cmp on mg.cmpid = cmp.cmpid"
tmp = tmp & " inner join mst_location ml on ml.locationid = cmp.locationid"
tmp = tmp & " inner join mst_commodity mcom on mcom.commodityid = mgsl.commodityid"
If Trim(CmbBatch.Text) = "" Then
   tmp = tmp & " where tcd.txntype = 'D' and isnull(tcgsl.commtrackbatchid,0) = 0 and tcd.flag = 'P' and tcd.ExchangeTypeID = 1 and tcd.CxTFDate between '" & Format(TxtFromDate.Value, "yyyy-mm-dd") & "' and '" & Format(TxtToDate.Value, "yyyy-mm-dd") & "'"
   tmp = tmp & " and mgsl.godownid in " & mGodownIDs
Else
   tmp = tmp & " where tcd.txntype = 'D' and tcgsl.commtrackbatchid = " & CmbBatch.Text
End If
tmp = tmp & " group by ml.ncdexcode, mg.ncdexcode, mcom.ncdexcode, mgsl.StackNo, mgsl.LotNo, tcd.CxTFDate, mgsl.godownid, tcgsl.CxtfNo,  cmp.cmpname, mg.godownno, tcgsl.SM_Prefix, mgsl.GSLID, mgsl.SM_prefix "
tmp = tmp & " Union"
tmp = tmp & " select"
tmp = tmp & " upper(ml.ncdexcode) LOCATION_CODE, upper(mg.ncdexcode) WH_CODE, upper(mcom.ncdexcode) COMMODITY_CODE,"
tmp = tmp & " '' LOTNO, '375' CLIENTID,"
tmp = tmp & " sum(tcgsl.TransferedWeight) ACTUALQTY, sum(tcgsl.TransferedWeight) NOMINALQTY, 'DEFAULT' ACTUALGRADE, 'DEFAULT' NOMINALGRADE, '' FED, 'UnAssayed' DEPOSITETYPE,"
tmp = tmp & " tcgsl.GSLTransferDate DEPOSITEDATE, 'MT' WAREHOUSEUOM , sum(tcgsl.TransferedWeight) GROSSQTY, 'Bags' PACKAGETYPE, '' PACKAGESIZE,"
tmp = tmp & " sum(tcgsl.TransferedWeight) NETQTY, '' WEIGHINGOFFICERCODE, '' WEIGHINGDATE, '' ASSAYERNAME, '' ASSAYERAGENCY, '' INTASSAYERREFNO,"
tmp = tmp & " '' ASSAYERREFNO, '' SAMPLENO, '' SAMPLEDATE, '' ASSAYDATE, '' REVALIDATIONDATE, '' GRADEPARAMNAME, '' GRADEPARAMVALUE,"
tmp = tmp & " upper(mg.ncdexcode) BAYNO, upper(mgsl.stackNo) STACKNO, upper(mgsl.LotNo) SLOTNO, '' ASSAYERCODE,"
tmp = tmp & " 'T'+convert(varchar, mgsl.GSLID)+upper(mgsl.SM_prefix)+convert(varchar, tcgsl.GSLTransferID) RECEIPTNO, '' PACKAGESIZE1, sum(tcgsl.TransferedBags) NOOFPACK, mgsl.godownid, tcgsl.GSLTransferID CxtfNo, '' SM_Prefix"
tmp = tmp & " from Txn_GSL_Transfer  tcgsl"
tmp = tmp & " inner join mst_gsl mgsl on mgsl.GSLID = tcgsl.newgslid and mgsl.sm_prefix = tcgsl.sm_prefix"
tmp = tmp & " inner join mst_godown mg on mg.GodownID = mgsl.godownid"
tmp = tmp & " inner join mst_cmp cmp on mg.cmpid = cmp.cmpid"
tmp = tmp & " inner join mst_location ml on ml.locationid = cmp.locationid"
tmp = tmp & " inner join mst_commodity mcom on mcom.commodityid = mgsl.commodityid"
If Trim(CmbBatch.Text) = "" Then
   tmp = tmp & " where mgsl.ExchangeTypeID = 1 and isnull(tcgsl.commtrackbatchid,0) = 0 and tcgsl.GSLTransferDate between '" & Format(TxtFromDate.Value, "yyyy-mm-dd") & "' and '" & Format(TxtToDate.Value, "yyyy-mm-dd") & "'"
   tmp = tmp & " and mgsl.godownid in " & mGodownIDs
Else
   tmp = tmp & " where tcgsl.commtrackbatchid = " & CmbBatch.Text
End If
tmp = tmp & " group by ml.ncdexcode, mg.ncdexcode, mcom.ncdexcode, mgsl.StackNo, mgsl.LotNo, tcgsl.GSLTransferDate, mgsl.godownid, tcgsl.GSLTransferID,  cmp.cmpname, mg.godownno, mgsl.GSLID, mgsl.SM_prefix"
tmp = tmp & " having Sum (tcgsl.TransferedWeight) > 0 "
Call TmpTable

If TmpRs.RecordCount > 0 Then
   'CmdExport.Enabled = True
   ProgressBar1.Value = 0
   ProgressBar1.Max = TmpRs.RecordCount + 3
   With MSHFlexGrid10
      For k = 1 To TmpRs.RecordCount
         i = .Rows - 1
         If Check2.Value = vbChecked Then
            If IIf(IsNull(TmpRs!LOCATION_CODE), "", TmpRs!LOCATION_CODE) = "" Or IIf(IsNull(TmpRs!WH_CODE), "", TmpRs!WH_CODE) = "" Or IIf(IsNull(TmpRs!COMMODITY_CODE), "", TmpRs!COMMODITY_CODE) = "" Then
               GoTo NextRow
            End If
         End If
         .TextMatrix(i, 0) = IIf(IsNull(TmpRs!LOCATION_CODE), "", TmpRs!LOCATION_CODE)
         If CmdExport.Enabled = True And Trim(.TextMatrix(i, 0)) = "" Then
            CmdExport.Enabled = False
         End If
         .TextMatrix(i, 1) = IIf(IsNull(TmpRs!WH_CODE), "", TmpRs!WH_CODE)
         If CmdExport.Enabled = True And Trim(.TextMatrix(i, 1)) = "" Then
            CmdExport.Enabled = False
         End If
         .TextMatrix(i, 2) = IIf(IsNull(TmpRs!COMMODITY_CODE), "", TmpRs!COMMODITY_CODE)
         If CmdExport.Enabled = True And Trim(.TextMatrix(i, 2)) = "" Then
            CmdExport.Enabled = False
         End If
         .TextMatrix(i, 3) = IIf(IsNull(TmpRs!LotNo), "", TmpRs!LotNo)
         .TextMatrix(i, 4) = IIf(IsNull(TmpRs!ClientID), "", TmpRs!ClientID)
         .TextMatrix(i, 5) = IIf(IsNull(TmpRs!ActualQty), "", TmpRs!ActualQty)
         .TextMatrix(i, 6) = IIf(IsNull(TmpRs!NOMINALQTY), "", TmpRs!NOMINALQTY)
         .TextMatrix(i, 7) = IIf(IsNull(TmpRs!ACTUALGRADE), "", TmpRs!ACTUALGRADE)
         .TextMatrix(i, 8) = IIf(IsNull(TmpRs!NOMINALGRADE), "", TmpRs!NOMINALGRADE)
         .TextMatrix(i, 9) = IIf(IsNull(TmpRs!FED), "", TmpRs!FED)
         .TextMatrix(i, 10) = IIf(IsNull(TmpRs!DEPOSITETYPE), "", TmpRs!DEPOSITETYPE)
         .TextMatrix(i, 11) = IIf(IsNull(TmpRs!DEPOSITEDATE), "", Format(TmpRs!DEPOSITEDATE, "yyyy-mm-dd hh:MM:ss"))
         .TextMatrix(i, 12) = IIf(IsNull(TmpRs!WAREHOUSEUOM), "", TmpRs!WAREHOUSEUOM)
         .TextMatrix(i, 13) = IIf(IsNull(TmpRs!GROSSQTY), "", TmpRs!GROSSQTY)
         .TextMatrix(i, 14) = IIf(IsNull(TmpRs!PACKAGETYPE), "", TmpRs!PACKAGETYPE)
         .TextMatrix(i, 15) = IIf(IsNull(TmpRs!PACKAGESIZE), "", TmpRs!PACKAGESIZE)
         .TextMatrix(i, 16) = IIf(IsNull(TmpRs!NETQTY), "", TmpRs!NETQTY)
         .TextMatrix(i, 17) = IIf(IsNull(TmpRs!WEIGHINGOFFICERCODE), "", TmpRs!WEIGHINGOFFICERCODE)
         .TextMatrix(i, 18) = IIf(IsNull(TmpRs!WEIGHINGDATE), "", Format(TmpRs!WEIGHINGDATE, "dd-MMM-yyyy"))
         .TextMatrix(i, 19) = IIf(IsNull(TmpRs!AssayerName), "", TmpRs!AssayerName)
         .TextMatrix(i, 20) = IIf(IsNull(TmpRs!ASSAYERAGENCY), "", TmpRs!ASSAYERAGENCY)
         .TextMatrix(i, 21) = IIf(IsNull(TmpRs!INTASSAYERREFNO), "", TmpRs!INTASSAYERREFNO)
         .TextMatrix(i, 22) = IIf(IsNull(TmpRs!ASSAYERREFNO), "", TmpRs!ASSAYERREFNO)
         .TextMatrix(i, 23) = IIf(IsNull(TmpRs!SAMPLENO), "", TmpRs!SAMPLENO)
         .TextMatrix(i, 24) = IIf(IsNull(TmpRs!SAMPLEDATE), "", Format(TmpRs!SAMPLEDATE, "dd-MMM-yyyy"))
         .TextMatrix(i, 25) = IIf(IsNull(TmpRs!ASSAYDATE), "", Format(TmpRs!ASSAYDATE, "dd-MMM-yyyy"))
         .TextMatrix(i, 26) = IIf(IsNull(TmpRs!RevalidationDate), "", Format(TmpRs!RevalidationDate, "dd-MMM-yyyy"))
         .TextMatrix(i, 27) = IIf(IsNull(TmpRs!GRADEPARAMNAME), "", TmpRs!GRADEPARAMNAME)
         .TextMatrix(i, 28) = IIf(IsNull(TmpRs!GRADEPARAMVALUE), "", TmpRs!GRADEPARAMVALUE)
         .TextMatrix(i, 29) = IIf(IsNull(TmpRs!BAYNO), "", TmpRs!BAYNO)
         .TextMatrix(i, 30) = IIf(IsNull(TmpRs!StackNo), "", TmpRs!StackNo)
         .TextMatrix(i, 31) = RemoveSpChar1(IIf(IsNull(TmpRs!SLOTNO), "", TmpRs!SLOTNO))
         .TextMatrix(i, 32) = IIf(IsNull(TmpRs!ASSAYERCODE), "", TmpRs!ASSAYERCODE)
         .TextMatrix(i, 33) = IIf(IsNull(TmpRs!RECEIPTNO), "", TmpRs!RECEIPTNO)
         .TextMatrix(i, 34) = IIf(IsNull(TmpRs!PACKAGESIZE1), "", TmpRs!PACKAGESIZE1)
         .TextMatrix(i, 35) = IIf(IsNull(TmpRs!NOOFPACK), "", TmpRs!NOOFPACK)
         If OptPending.Value = True Then
            If TmpRs2.RecordCount > 0 Then
               TmpRs2.MoveFirst
               TmpRs2.Find "GodownID = " & TmpRs!GodownID
               If Not TmpRs2.EOF Then
                  .TextMatrix(i, 36) = IIf(IsNull(TmpRs2!BalanceBags), 0, TmpRs2!BalanceBags)
                  .TextMatrix(i, 37) = IIf(IsNull(TmpRs2!BalanceWeight), 0, TmpRs2!BalanceWeight)
               End If
            End If
            .TextMatrix(i, 38) = IIf(IsNull(TmpRs!CxTFNO), "", TmpRs!CxTFNO)
            .TextMatrix(i, 39) = IIf(IsNull(TmpRs!SM_Prefix), "", TmpRs!SM_Prefix)
         End If
         .Rows = .Rows + 1
NextRow:
         ProgressBar1.Value = ProgressBar1.Value + 1
         TmpRs.MoveNext
      Next
      ProgressBar1.Value = ProgressBar1.Max
   End With
End If
End Sub

Private Sub Get_OpeningDEPOSITEDATA()
'=========================================one time query===============================================================
'----------only for opening stock
'--Update txn_cxtf_gsl for opening stock
'tmp = " Update tcgsl Set CommtrackBatchDate = '2012-12-25', CommtrackBatchID = 1"
'tmp = tmp & " from txn_cxtf_gsl tcgsl"
'tmp = tmp & " inner join Txn_CxTF_Details tcd on tcd.cxtfno = tcgsl.cxtfno and tcd.sm_prefix = tcgsl.sm_prefix and tcd.txntype = tcgsl.txntype"
'tmp = tmp & " inner join mst_gsl mgsl on mgsl.GSLID = tcgsl.gslid and mgsl.sm_prefix = tcgsl.sm_prefix"
'tmp = tmp & " inner join mst_godown mg on mg.GodownID = mgsl.godownid"
'tmp = tmp & " inner join mst_cmp cmp on mg.cmpid = cmp.cmpid"
'tmp = tmp & " inner join mst_location ml on ml.locationid = cmp.locationid"
'tmp = tmp & " inner join mst_commodity mcom on mcom.commodityid = mgsl.commodityid"
'tmp = tmp & " where MGSL.ExchangeTypeID = 1 And MGSL.Flag = 'O' and mgsl.godownid in"
'tmp = tmp & " (Select Distinct GodownID from Mst_GSL Where Mst_GSL.ExchangeTypeID = 2 and flag not in ('E', 'J', 'Z') and balancebags > 0 )"
'tmp = tmp & " and isnull(mg.ncdexcode,'') <> '' and mgsl.CDTFDate <= '2012-12-25' and mgsl.BalanceWeight > 0"
'tmp = tmp & " and tcd.CxTFDate <= '2012-12-25' and isnull(tcgsl.CommtrackBatchID,0) = 0"
'Call TmpTable
'
'''--Update Txn_GSL_Transfer for opening stock
'tmp = tmp & " Update tcgsl Set CommtrackBatchDate = '2012-12-25', CommtrackBatchID = 1"
'tmp = tmp & " from Txn_GSL_Transfer tcgsl"
'tmp = tmp & " inner join mst_gsl mgsl on (mgsl.GSLID = tcgsl.newgslid or mgsl.GSLID = tcgsl.oldgslid) and mgsl.sm_prefix = tcgsl.sm_prefix"
'tmp = tmp & " inner join mst_godown mg on mg.GodownID = mgsl.godownid"
'tmp = tmp & " inner join mst_cmp cmp on mg.cmpid = cmp.cmpid"
'tmp = tmp & " inner join mst_location ml on ml.locationid = cmp.locationid"
'tmp = tmp & " inner join mst_commodity mcom on mcom.commodityid = mgsl.commodityid"
'tmp = tmp & " where MGSL.ExchangeTypeID = 1 And MGSL.Flag = 'O' and mgsl.godownid in"
'tmp = tmp & " (Select Distinct GodownID from Mst_GSL Where Mst_GSL.ExchangeTypeID = 2 and flag not in ('E', 'J', 'Z') and balancebags > 0 )"
'tmp = tmp & " and isnull(mg.ncdexcode,'') <> '' and mgsl.CDTFDate <= '2012-12-25' and mgsl.BalanceWeight > 0"
'tmp = tmp & " and tcgsl.GSLTransferDate <= '2012-12-25' and isnull(tcgsl.CommtrackBatchID,0) = 0"
'Call TmpTable

''----------data from gsl for opening stock
'tmp = " Select upper(ml.NCDEXCODE) LOCATION_CODE, upper(mg.ncdexcode) WH_CODE, upper(mcom.ncdexcode) COMMODITY_CODE, '' LOTNO, '375' CLIENTID, Sum(mgsl.BalanceWeight) ACTUALQTY, Sum(mgsl.BalanceWeight) NOMINALQTY, 'DEFAULT' ACTUALGRADE, 'DEFAULT' NOMINALGRADE, '' FED, 'UnAssayed' DEPOSITETYPE,"
'tmp = tmp & " mgsl.CDTFDate DEPOSITEDATE, 'MT' WAREHOUSEUOM , Sum(mgsl.BalanceWeight) GROSSQTY, 'Bags' PACKAGETYPE, '' PACKAGESIZE, Sum(mgsl.BalanceWeight) NETQTY, '' WEIGHINGOFFICERCODE, '' WEIGHINGDATE, '' ASSAYERNAME, '' ASSAYERAGENCY, '' INTASSAYERREFNO, '' ASSAYERREFNO, '' SAMPLENO, '' SAMPLEDATE, '' ASSAYDATE, '' REVALIDATIONDATE, '' GRADEPARAMNAME, '' GRADEPARAMVALUE, upper(mg.ncdexcode) BAYNO, upper(mgsl.stackNo) STACKNO, upper(cmp.cmpname)+'_'+upper(mg.godownno)+'~'+upper(mgsl.StackNo)+'~'+upper(mgsl.LotNo) SLOTNO, '' ASSAYERCODE, 'D'+convert(varchar, mgsl.CDTFNo) RECEIPTNO, '' PACKAGESIZE1, Sum(mgsl.BalanceBags) NOOFPACK, mgsl.godownid"
'tmp = tmp & " from mst_gsl mgsl"
'tmp = tmp & " inner join mst_godown mg on mg.GodownID = mgsl.godownid"
'tmp = tmp & " inner join mst_cmp cmp on mg.cmpid = cmp.cmpid"
'tmp = tmp & " inner join mst_location ml on ml.locationid = cmp.locationid"
'tmp = tmp & " inner join mst_commodity mcom on mcom.commodityid = mgsl.commodityid"
'tmp = tmp & " where MGSL.ExchangeTypeID = 1 And MGSL.Flag = 'O' and mgsl.godownid in"
'tmp = tmp & " (Select Distinct GodownID from Mst_GSL Where Mst_GSL.ExchangeTypeID = 2 and flag not in ('E', 'J', 'Z') and balancebags > 0 )"
'tmp = tmp & " and isnull(mg.ncdexcode,'') <> '' and mgsl.CDTFDate <= '2012-12-25' "
'tmp = tmp & " group by ml.ncdexcode, mg.ncdexcode, mcom.ncdexcode, mgsl.StackNo, mgsl.LotNo, mgsl.godownid, mgsl.CDTFDate, mgsl.CDTFNo, cmp.cmpname, mg.godownno"
'tmp = tmp & " Having Sum(mgsl.BalanceWeight) > 0"
'Call TmpTable
''=========================================one time query===============================================================

''=========================================query after adding some godown ncdex code====================================
tmp = " Select upper(ml.NCDEXCODE) LOCATION_CODE, upper(mg.ncdexcode) WH_CODE, upper(mcom.ncdexcode) COMMODITY_CODE, '' LOTNO, '375' CLIENTID, Sum(mgsl.BalanceWeight) ACTUALQTY, Sum(mgsl.BalanceWeight) NOMINALQTY, 'DEFAULT' ACTUALGRADE, 'DEFAULT' NOMINALGRADE, '' FED, 'UnAssayed' DEPOSITETYPE,"
tmp = tmp & " mgsl.CDTFDate DEPOSITEDATE, 'MT' WAREHOUSEUOM , Sum(mgsl.BalanceWeight) GROSSQTY, 'Bags' PACKAGETYPE, '' PACKAGESIZE, Sum(mgsl.BalanceWeight) NETQTY, '' WEIGHINGOFFICERCODE, '' WEIGHINGDATE, '' ASSAYERNAME, '' ASSAYERAGENCY, '' INTASSAYERREFNO, '' ASSAYERREFNO, '' SAMPLENO, '' SAMPLEDATE, '' ASSAYDATE, '' REVALIDATIONDATE, '' GRADEPARAMNAME, '' GRADEPARAMVALUE, upper(mg.ncdexcode) BAYNO, upper(mgsl.stackNo) STACKNO, upper(mgsl.LotNo) SLOTNO, '' ASSAYERCODE, 'D'+convert(varchar, mgsl.GSLID)+upper(mgsl.SM_prefix) RECEIPTNO, '' PACKAGESIZE1, Sum(mgsl.BalanceBags) NOOFPACK, mgsl.godownid"
tmp = tmp & " from mst_gsl mgsl"
tmp = tmp & " inner join mst_godown mg on mg.GodownID = mgsl.godownid"
tmp = tmp & " inner join mst_cmp cmp on mg.cmpid = cmp.cmpid"
tmp = tmp & " inner join mst_location ml on ml.locationid = cmp.locationid"
tmp = tmp & " inner join mst_commodity mcom on mcom.commodityid = mgsl.commodityid"
tmp = tmp & " where MGSL.ExchangeTypeID = 1 And MGSL.Flag = 'O' and mgsl.godownid in"
tmp = tmp & " (Select Distinct GodownID from Mst_GSL Where Mst_GSL.ExchangeTypeID = 2 and flag not in ('E', 'J', 'Z') and balancebags > 0 )"
tmp = tmp & " and isnull(mg.ncdexcode,'') <> '' and mgsl.CDTFDate <= '2012-12-25'"
tmp = tmp & " and convert(varchar, mgsl.gslid)+'_'+mgsl.sm_prefix not in (select distinct convert(varchar, gslid)+'_'+sm_prefix from txn_cxtf_gsl where isnull(commtrackbatchid, 0) <> 0)"
tmp = tmp & " group by ml.ncdexcode, mg.ncdexcode, mcom.ncdexcode, mgsl.StackNo, mgsl.LotNo, mgsl.godownid, mgsl.CDTFDate, mgsl.CDTFNo, cmp.cmpname, mg.godownno, mgsl.GSLID, mgsl.SM_prefix"
tmp = tmp & " Having Sum(mgsl.BalanceWeight) > 0"
Call TmpTable
''=========================================query after adding some godown ncdex code====================================

If TmpRs.RecordCount > 0 Then
   CmdExport.Enabled = True
   ProgressBar1.Value = 0
   ProgressBar1.Max = TmpRs.RecordCount + 3
   With MSHFlexGrid10
      .Rows = .Rows + TmpRs.RecordCount
      For i = 1 To TmpRs.RecordCount
         .TextMatrix(i, 0) = IIf(IsNull(TmpRs!LOCATION_CODE), "", TmpRs!LOCATION_CODE)
         If CmdExport.Enabled = True And Trim(.TextMatrix(i, 0)) = "" Then
            CmdExport.Enabled = False
         End If
         .TextMatrix(i, 1) = IIf(IsNull(TmpRs!WH_CODE), "", TmpRs!WH_CODE)
         If CmdExport.Enabled = True And Trim(.TextMatrix(i, 1)) = "" Then
            CmdExport.Enabled = False
         End If
         .TextMatrix(i, 2) = IIf(IsNull(TmpRs!COMMODITY_CODE), "", TmpRs!COMMODITY_CODE)
         If CmdExport.Enabled = True And Trim(.TextMatrix(i, 2)) = "" Then
            CmdExport.Enabled = False
         End If
         .TextMatrix(i, 3) = IIf(IsNull(TmpRs!LotNo), "", TmpRs!LotNo)
         .TextMatrix(i, 4) = IIf(IsNull(TmpRs!ClientID), "", TmpRs!ClientID)
         .TextMatrix(i, 5) = IIf(IsNull(TmpRs!ActualQty), "", TmpRs!ActualQty)
         .TextMatrix(i, 6) = IIf(IsNull(TmpRs!NOMINALQTY), "", TmpRs!NOMINALQTY)
         .TextMatrix(i, 7) = IIf(IsNull(TmpRs!ACTUALGRADE), "", TmpRs!ACTUALGRADE)
         .TextMatrix(i, 8) = IIf(IsNull(TmpRs!NOMINALGRADE), "", TmpRs!NOMINALGRADE)
         .TextMatrix(i, 9) = IIf(IsNull(TmpRs!FED), "", TmpRs!FED)
         .TextMatrix(i, 10) = IIf(IsNull(TmpRs!DEPOSITETYPE), "", TmpRs!DEPOSITETYPE)
         .TextMatrix(i, 11) = IIf(IsNull(TmpRs!DEPOSITEDATE), "", Format(TmpRs!DEPOSITEDATE, "yyyy-mm-dd hh:MM:ss"))
         .TextMatrix(i, 12) = IIf(IsNull(TmpRs!WAREHOUSEUOM), "", TmpRs!WAREHOUSEUOM)
         .TextMatrix(i, 13) = IIf(IsNull(TmpRs!GROSSQTY), "", TmpRs!GROSSQTY)
         .TextMatrix(i, 14) = IIf(IsNull(TmpRs!PACKAGETYPE), "", TmpRs!PACKAGETYPE)
         .TextMatrix(i, 15) = IIf(IsNull(TmpRs!PACKAGESIZE), "", TmpRs!PACKAGESIZE)
         .TextMatrix(i, 16) = IIf(IsNull(TmpRs!NETQTY), "", TmpRs!NETQTY)
         .TextMatrix(i, 17) = IIf(IsNull(TmpRs!WEIGHINGOFFICERCODE), "", TmpRs!WEIGHINGOFFICERCODE)
         .TextMatrix(i, 18) = IIf(IsNull(TmpRs!WEIGHINGDATE), "", Format(TmpRs!WEIGHINGDATE, "dd-MMM-yyyy"))
         .TextMatrix(i, 19) = IIf(IsNull(TmpRs!AssayerName), "", TmpRs!AssayerName)
         .TextMatrix(i, 20) = IIf(IsNull(TmpRs!ASSAYERAGENCY), "", TmpRs!ASSAYERAGENCY)
         .TextMatrix(i, 21) = IIf(IsNull(TmpRs!INTASSAYERREFNO), "", TmpRs!INTASSAYERREFNO)
         .TextMatrix(i, 22) = IIf(IsNull(TmpRs!ASSAYERREFNO), "", TmpRs!ASSAYERREFNO)
         .TextMatrix(i, 23) = IIf(IsNull(TmpRs!SAMPLENO), "", TmpRs!SAMPLENO)
         .TextMatrix(i, 24) = IIf(IsNull(TmpRs!SAMPLEDATE), "", Format(TmpRs!SAMPLEDATE, "dd-MMM-yyyy"))
         .TextMatrix(i, 25) = IIf(IsNull(TmpRs!ASSAYDATE), "", Format(TmpRs!ASSAYDATE, "dd-MMM-yyyy"))
         .TextMatrix(i, 26) = IIf(IsNull(TmpRs!RevalidationDate), "", Format(TmpRs!RevalidationDate, "dd-MMM-yyyy"))
         .TextMatrix(i, 27) = IIf(IsNull(TmpRs!GRADEPARAMNAME), "", TmpRs!GRADEPARAMNAME)
         .TextMatrix(i, 28) = IIf(IsNull(TmpRs!GRADEPARAMVALUE), "", TmpRs!GRADEPARAMVALUE)
         .TextMatrix(i, 29) = IIf(IsNull(TmpRs!BAYNO), "", TmpRs!BAYNO)
         .TextMatrix(i, 30) = IIf(IsNull(TmpRs!StackNo), "", TmpRs!StackNo)
         .TextMatrix(i, 31) = RemoveSpChar1(IIf(IsNull(TmpRs!SLOTNO), "", TmpRs!SLOTNO))
         '.TextMatrix(I, 31) = IIf(IsNull(TmpRs!SLOTNO), "", TmpRs!SLOTNO)
         .TextMatrix(i, 32) = IIf(IsNull(TmpRs!ASSAYERCODE), "", TmpRs!ASSAYERCODE)
         .TextMatrix(i, 33) = IIf(IsNull(TmpRs!RECEIPTNO), "", TmpRs!RECEIPTNO)
         .TextMatrix(i, 34) = IIf(IsNull(TmpRs!PACKAGESIZE1), "", TmpRs!PACKAGESIZE1)
         .TextMatrix(i, 35) = IIf(IsNull(TmpRs!NOOFPACK), "", TmpRs!NOOFPACK)
         If OptPending.Value = True Then
            If TmpRs2.RecordCount > 0 Then
               TmpRs2.MoveFirst
               TmpRs2.Find "GodownID = " & TmpRs!GodownID
               If Not TmpRs2.EOF Then
                  .TextMatrix(i, 36) = IIf(IsNull(TmpRs2!BalanceBags), 0, TmpRs2!BalanceBags)
                  .TextMatrix(i, 37) = IIf(IsNull(TmpRs2!BalanceWeight), 0, TmpRs2!BalanceWeight)
               End If
            End If
         End If
         ProgressBar1.Value = ProgressBar1.Value + 1
         TmpRs.MoveNext
      Next
      ProgressBar1.Value = ProgressBar1.Max
   End With
End If
End Sub

Public Function RemoveSpChar1(strToCheck) As String
Dim Rstr As Variant
Dim lng As Long
For lng = 1 To Len(strToCheck)
        Select Case Asc(mID$(strToCheck, lng, 1))
        Case 32
        Rstr = Rstr + mID$(strToCheck, lng, 1)
        Case 48 To 57
        ' Numbers 0 to 9
          Rstr = Rstr + mID$(strToCheck, lng, 1)
        Case 65 To 90
        ' Letters A to Z
          Rstr = Rstr + mID$(strToCheck, lng, 1)
        Case 97 To 122
        ' Letters a to z
          Rstr = Rstr + mID$(strToCheck, lng, 1)
        Case Else
             Rstr = Rstr + "_"
        End Select
Next lng
RemoveSpChar1 = Rstr
End Function


Private Sub Get_UOM_MSTR()

tmp = "select 'MT' UOM_CODE, 'METRIC TON' UOM_NAME, 'MT' UOM_UNIT, '' IS_FREEZED, 'A' IS_ACTIVE, '' UOM_NAME_0,"
tmp = tmp & " '' UOM_NAME_1, '' UOM_NAME_2"
Call TmpTable
If TmpRs.RecordCount > 0 Then
   ProgressBar1.Value = 0
   ProgressBar1.Max = TmpRs.RecordCount + 3
   With MSHFlexGrid1
      .Rows = .Rows + TmpRs.RecordCount
      For i = 1 To TmpRs.RecordCount
         .TextMatrix(i, 0) = IIf(IsNull(TmpRs!UOM_CODE), "", TmpRs!UOM_CODE)
         .TextMatrix(i, 1) = IIf(IsNull(TmpRs!UOM_NAME), "", TmpRs!UOM_NAME)
         .TextMatrix(i, 2) = IIf(IsNull(TmpRs!UOM_UNIT), "", TmpRs!UOM_UNIT)
         .TextMatrix(i, 3) = IIf(IsNull(TmpRs!IS_FREEZED), "", TmpRs!IS_FREEZED)
         .TextMatrix(i, 4) = IIf(IsNull(TmpRs!IS_ACTIVE), "", TmpRs!IS_ACTIVE)
         .TextMatrix(i, 5) = IIf(IsNull(TmpRs!UOM_NAME_0), "", TmpRs!UOM_NAME_0)
         .TextMatrix(i, 6) = IIf(IsNull(TmpRs!UOM_NAME_1), "", TmpRs!UOM_NAME_1)
         .TextMatrix(i, 7) = IIf(IsNull(TmpRs!UOM_NAME_2), "", TmpRs!UOM_NAME_2)
         ProgressBar1.Value = ProgressBar1.Value + 1
         TmpRs.MoveNext
      Next
   End With
End If
End Sub

Private Sub Get_CLIENT_MSTR()
tmp = " select"
tmp = tmp & " '1' CLIENT_ID, '' CLIENT_TYPE, '' AUX_CLIENT_ID, '' TRADING_CLIENT_ID, 'NCML' CLIENT_NAME, 'Gayatri Towers, 954 , Appasaheb Marathe Marg, Prabhadevi , Mumbai 400 025' CLIENT_ADDRESS,"
tmp = tmp & " '' CLIENT_DOC_TYPE, '02240419191' CLIENT_PHONE1, '' CLIENT_PHONE2, '' WH_CODE, '' LOCATION_CODE, '' CLIENT_EMAIL,"
tmp = tmp & " '' CLIENT_ORG_TYPE, '' CLIENT_CONTACT, '' OPERATOR_ID, '' OPERATOR_NAME, '' CONTACT_PERSON_NAME, '' CLIENT_MOB_NO,"
tmp = tmp & " '' OPERATOR_CONTACT_NO, '' CLIENT_DOC_NO, '' CLIENT_PASSWORD, '' EXCHANGE_CODE, '' CM_ID, '' DP_ID, '' TRADING_ACC_ID,"
tmp = tmp & " '' APPROVAL_STATUS, '' APPROVAL_DATE, '' UCC_ID, '' CLIENT_TEL, '' CLIENT_TLX, '' CLIENT_FAX, '' PASSWORD, '' OPERATOR_CATEGORY,"
tmp = tmp & " '' CONTACT_PERSON_CONTACT, '' CREATED_BY, '' CREATED_ON, '' MODIFIED_BY, '' MODIFIED_ON, '' CMSACCNT, '' CLIENT_NEW_EMAIL,"
tmp = tmp & " '' CLIENT_NAME_0, '' CLIENT_NAME_1, '' IS_FREEZED, 'A' IS_ACTIVE, '' PLEDGING_ALLOWED, '' AS_ADDRESS, '' AS_ALTERNATE_TEL,"
tmp = tmp & " '' AS_CITY, '' AS_DESIGNATION, '' AS_EMAIL_ID, '' AS_FAX, '' AS_MOBILE_NO, '' AS_NAME, '' AS_PAN_NAME, '' AS_PINCODE, '' AS_TEL,"
tmp = tmp & " '' BD_BANK_ACCOUNT_NO, '' BD_BANK_CODE, '' BD_BRANCH_CODE, '' BD_IFSC_CODE, '' CA_ALTERNATE_TEL, '' CA_DESIGNATION, '' CA_FAX,"
tmp = tmp & " '' CA_PINCODE, '' CONTACT_PERSON_NO, '' DATE_OF_BIRTH, '' FIRST_LOGIN_FLAG, '' MEMBERSHIP_TYPE, '' MEMBER_SHIP_ID,"
tmp = tmp & " '' OPERATOR_TYPE, '' ORG_CONSTITUTION, '' PAN_NO, '' PD_AMOUNT, '' PD_APPLICANTS_BANK, '' PD_APP_BANK_BRANCH, '' PD_CLEARING_BANK,"
tmp = tmp & " '' PD_CL_BANK_BRANCH, '' PD_IFSC_CODE,'' PD_INSTRUMENT_NO, '' PD_INSTRUMENT_TYPE, '' PRCNT_REBATE, '' RO_ADDRESS,"
tmp = tmp & " '' RO_ALTERNATE_TEL, '' RO_CITY, '' RO_EMAIL_ID, '' RO_FAX, '' RO_PINCODE, '' RO_TEL, '' RO_TELEX, '' TAN_NO, '' FIRST_CLIENT_ADDRESS, '' FIRST_CITY,"
tmp = tmp & " '' FIRST_FAT_HUBBY, '' FIRST_HOLDER_NAME, '' FIRST_OCCUPATION, '' FIRST_PAN, '' FIRST_PINCODE, '' MODE_OF_OPERATION, '' NOMINEE_ADDRESS,"
tmp = tmp & " '' NOMINEE_AGE, '' NOMINEE_CITY, '' NOMINEE_NAME, '' NOMINEE_OCCUPATION, '' NOMINEE_PINCODE, '' NOMINEE_RELATION,"
tmp = tmp & " '' SECOND_CITY, '' SECOND_CLIENT_ADDRESS, '' SECOND_FAT_HUBBY, '' SECOND_HOLDER_NAME, '' SECOND_OCCUPATION, '' SECOND_PAN,"
tmp = tmp & " '' SECOND_PINCODE, '' THIRD_CITY, '' THIRD_CLIENT_ADDRESS, '' THIRD_FAT_HUBBY, '' THIRD_HOLDER_NAME, '' THIRD_OCCUPATION,"
tmp = tmp & " '' THIRD_PAN, '' THIRD_PINCODE,'' UPLOAD_FOR_FIRST_HOLDER, '' UPLOAD_FOR_SECOND_HOLDER, '' UPLOAD_FOR_THIRD_HOLDER, '' ORG_CLIENT_TYPE,"
tmp = tmp & " '' ACCOUNT_VALIDITY_DATE, '' ACCOUNT_OPENING_CHARGES, '' RENEWAL_CHARGES, '' BRANCH_CODE"
Call TmpTable
If TmpRs.RecordCount > 0 Then
   ProgressBar1.Value = 0
   ProgressBar1.Max = TmpRs.RecordCount + 3
   With MSHFlexGrid2
      .Rows = .Rows + TmpRs.RecordCount
      For i = 1 To TmpRs.RecordCount
         .TextMatrix(i, 0) = IIf(IsNull(TmpRs!CLIENT_ID), "", TmpRs!CLIENT_ID)
         .TextMatrix(i, 1) = IIf(IsNull(TmpRs!CLIENT_TYPE), "", TmpRs!CLIENT_TYPE)
         .TextMatrix(i, 2) = IIf(IsNull(TmpRs!AUX_CLIENT_ID), "", TmpRs!AUX_CLIENT_ID)
         .TextMatrix(i, 3) = IIf(IsNull(TmpRs!TRADING_CLIENT_ID), "", TmpRs!TRADING_CLIENT_ID)
         .TextMatrix(i, 4) = IIf(IsNull(TmpRs!CLIENT_NAME), "", TmpRs!CLIENT_NAME)
         .TextMatrix(i, 5) = IIf(IsNull(TmpRs!CLIENT_ADDRESS), "", TmpRs!CLIENT_ADDRESS)
         .TextMatrix(i, 6) = IIf(IsNull(TmpRs!CLIENT_DOC_TYPE), "", TmpRs!CLIENT_DOC_TYPE)
         .TextMatrix(i, 7) = IIf(IsNull(TmpRs!CLIENT_PHONE1), "", TmpRs!CLIENT_PHONE1)
         .TextMatrix(i, 8) = IIf(IsNull(TmpRs!CLIENT_PHONE2), "", TmpRs!CLIENT_PHONE2)
         .TextMatrix(i, 9) = IIf(IsNull(TmpRs!WH_CODE), "", TmpRs!WH_CODE)
         .TextMatrix(i, 10) = IIf(IsNull(TmpRs!LOCATION_CODE), "", TmpRs!LOCATION_CODE)
         .TextMatrix(i, 11) = IIf(IsNull(TmpRs!CLIENT_EMAIL), "", TmpRs!CLIENT_EMAIL)
         .TextMatrix(i, 12) = IIf(IsNull(TmpRs!CLIENT_ORG_TYPE), "", TmpRs!CLIENT_ORG_TYPE)
         .TextMatrix(i, 13) = IIf(IsNull(TmpRs!CLIENT_CONTACT), "", TmpRs!CLIENT_CONTACT)
         .TextMatrix(i, 14) = IIf(IsNull(TmpRs!OPERATOR_ID), "", TmpRs!OPERATOR_ID)
         .TextMatrix(i, 15) = IIf(IsNull(TmpRs!OPERATOR_NAME), "", TmpRs!OPERATOR_NAME)
         .TextMatrix(i, 16) = IIf(IsNull(TmpRs!CONTACT_PERSON_NAME), "", TmpRs!CONTACT_PERSON_NAME)
         .TextMatrix(i, 17) = IIf(IsNull(TmpRs!CLIENT_MOB_NO), "", TmpRs!CLIENT_MOB_NO)
         .TextMatrix(i, 18) = IIf(IsNull(TmpRs!OPERATOR_CONTACT_NO), "", TmpRs!OPERATOR_CONTACT_NO)
         .TextMatrix(i, 19) = IIf(IsNull(TmpRs!CLIENT_DOC_NO), "", TmpRs!CLIENT_DOC_NO)
         .TextMatrix(i, 20) = IIf(IsNull(TmpRs!CLIENT_PASSWORD), "", TmpRs!CLIENT_PASSWORD)
         .TextMatrix(i, 21) = IIf(IsNull(TmpRs!EXCHANGE_CODE), "", TmpRs!EXCHANGE_CODE)
         .TextMatrix(i, 22) = IIf(IsNull(TmpRs!CM_ID), "", TmpRs!CM_ID)
         .TextMatrix(i, 23) = IIf(IsNull(TmpRs!DP_ID), "", TmpRs!DP_ID)
         .TextMatrix(i, 24) = IIf(IsNull(TmpRs!TRADING_ACC_ID), "", TmpRs!TRADING_ACC_ID)
         .TextMatrix(i, 25) = IIf(IsNull(TmpRs!APPROVAL_STATUS), "", TmpRs!APPROVAL_STATUS)
         .TextMatrix(i, 26) = IIf(IsNull(TmpRs!APPROVAL_DATE), "", TmpRs!APPROVAL_DATE)
         .TextMatrix(i, 27) = IIf(IsNull(TmpRs!UCC_ID), "", TmpRs!UCC_ID)
         .TextMatrix(i, 28) = IIf(IsNull(TmpRs!CLIENT_TEL), "", TmpRs!CLIENT_TEL)
         .TextMatrix(i, 29) = IIf(IsNull(TmpRs!CLIENT_TLX), "", TmpRs!CLIENT_TLX)
         .TextMatrix(i, 30) = IIf(IsNull(TmpRs!CLIENT_FAX), "", TmpRs!CLIENT_FAX)
         .TextMatrix(i, 31) = IIf(IsNull(TmpRs!Password), "", TmpRs!Password)
         .TextMatrix(i, 32) = IIf(IsNull(TmpRs!OPERATOR_CATEGORY), "", TmpRs!OPERATOR_CATEGORY)
         .TextMatrix(i, 33) = IIf(IsNull(TmpRs!CONTACT_PERSON_CONTACT), "", TmpRs!CONTACT_PERSON_CONTACT)
         .TextMatrix(i, 34) = IIf(IsNull(TmpRs!CREATED_BY), "", TmpRs!CREATED_BY)
         .TextMatrix(i, 35) = IIf(IsNull(TmpRs!CREATED_ON), "", TmpRs!CREATED_ON)
         .TextMatrix(i, 36) = IIf(IsNull(TmpRs!MODIFIED_BY), "", TmpRs!MODIFIED_BY)
         .TextMatrix(i, 37) = IIf(IsNull(TmpRs!MODIFIED_ON), "", TmpRs!MODIFIED_ON)
         .TextMatrix(i, 38) = IIf(IsNull(TmpRs!CMSACCNT), "", TmpRs!CMSACCNT)
         .TextMatrix(i, 39) = IIf(IsNull(TmpRs!CLIENT_NEW_EMAIL), "", TmpRs!CLIENT_NEW_EMAIL)
         .TextMatrix(i, 40) = IIf(IsNull(TmpRs!CLIENT_NAME_0), "", TmpRs!CLIENT_NAME_0)
         .TextMatrix(i, 41) = IIf(IsNull(TmpRs!CLIENT_NAME_1), "", TmpRs!CLIENT_NAME_1)
         .TextMatrix(i, 42) = IIf(IsNull(TmpRs!IS_FREEZED), "", TmpRs!IS_FREEZED)
         .TextMatrix(i, 43) = IIf(IsNull(TmpRs!IS_ACTIVE), "", TmpRs!IS_ACTIVE)
         .TextMatrix(i, 44) = IIf(IsNull(TmpRs!PLEDGING_ALLOWED), "", TmpRs!PLEDGING_ALLOWED)
         .TextMatrix(i, 45) = IIf(IsNull(TmpRs!AS_ADDRESS), "", TmpRs!AS_ADDRESS)
         .TextMatrix(i, 46) = IIf(IsNull(TmpRs!AS_ALTERNATE_TEL), "", TmpRs!AS_ALTERNATE_TEL)
         .TextMatrix(i, 47) = IIf(IsNull(TmpRs!AS_CITY), "", TmpRs!AS_CITY)
         .TextMatrix(i, 48) = IIf(IsNull(TmpRs!AS_DESIGNATION), "", TmpRs!AS_DESIGNATION)
         .TextMatrix(i, 49) = IIf(IsNull(TmpRs!AS_EMAIL_ID), "", TmpRs!AS_EMAIL_ID)
         .TextMatrix(i, 50) = IIf(IsNull(TmpRs!AS_FAX), "", TmpRs!AS_FAX)
         .TextMatrix(i, 51) = IIf(IsNull(TmpRs!AS_MOBILE_NO), "", TmpRs!AS_MOBILE_NO)
         .TextMatrix(i, 52) = IIf(IsNull(TmpRs!AS_NAME), "", TmpRs!AS_NAME)
         .TextMatrix(i, 53) = IIf(IsNull(TmpRs!AS_PAN_NAME), "", TmpRs!AS_PAN_NAME)
         .TextMatrix(i, 54) = IIf(IsNull(TmpRs!AS_PINCODE), "", TmpRs!AS_PINCODE)
         .TextMatrix(i, 55) = IIf(IsNull(TmpRs!AS_TEL), "", TmpRs!AS_TEL)
         .TextMatrix(i, 56) = IIf(IsNull(TmpRs!BD_BANK_ACCOUNT_NO), "", TmpRs!BD_BANK_ACCOUNT_NO)
         .TextMatrix(i, 57) = IIf(IsNull(TmpRs!BD_BANK_CODE), "", TmpRs!BD_BANK_CODE)
         .TextMatrix(i, 58) = IIf(IsNull(TmpRs!BD_BRANCH_CODE), "", TmpRs!BD_BRANCH_CODE)
         .TextMatrix(i, 59) = IIf(IsNull(TmpRs!BD_IFSC_CODE), "", TmpRs!BD_IFSC_CODE)
         .TextMatrix(i, 60) = IIf(IsNull(TmpRs!CA_ALTERNATE_TEL), "", TmpRs!CA_ALTERNATE_TEL)
         .TextMatrix(i, 61) = IIf(IsNull(TmpRs!CA_DESIGNATION), "", TmpRs!CA_DESIGNATION)
         .TextMatrix(i, 62) = IIf(IsNull(TmpRs!CA_FAX), "", TmpRs!CA_FAX)
         .TextMatrix(i, 63) = IIf(IsNull(TmpRs!CA_PINCODE), "", TmpRs!CA_PINCODE)
         .TextMatrix(i, 64) = IIf(IsNull(TmpRs!CONTACT_PERSON_NO), "", TmpRs!CONTACT_PERSON_NO)
         .TextMatrix(i, 65) = IIf(IsNull(TmpRs!DATE_OF_BIRTH), "", TmpRs!DATE_OF_BIRTH)
         .TextMatrix(i, 66) = IIf(IsNull(TmpRs!FIRST_LOGIN_FLAG), "", TmpRs!FIRST_LOGIN_FLAG)
         .TextMatrix(i, 67) = IIf(IsNull(TmpRs!MEMBERSHIP_TYPE), "", TmpRs!MEMBERSHIP_TYPE)
         .TextMatrix(i, 68) = IIf(IsNull(TmpRs!MEMBER_SHIP_ID), "", TmpRs!MEMBER_SHIP_ID)
         .TextMatrix(i, 69) = IIf(IsNull(TmpRs!OPERATOR_TYPE), "", TmpRs!OPERATOR_TYPE)
         .TextMatrix(i, 70) = IIf(IsNull(TmpRs!ORG_CONSTITUTION), "", TmpRs!ORG_CONSTITUTION)
         .TextMatrix(i, 71) = IIf(IsNull(TmpRs!PAN_NO), "", TmpRs!PAN_NO)
         .TextMatrix(i, 72) = IIf(IsNull(TmpRs!PD_AMOUNT), "", TmpRs!PD_AMOUNT)
         .TextMatrix(i, 73) = IIf(IsNull(TmpRs!PD_APPLICANTS_BANK), "", TmpRs!PD_APPLICANTS_BANK)
         .TextMatrix(i, 74) = IIf(IsNull(TmpRs!PD_APP_BANK_BRANCH), "", TmpRs!PD_APP_BANK_BRANCH)
         .TextMatrix(i, 75) = IIf(IsNull(TmpRs!PD_CLEARING_BANK), "", TmpRs!PD_CLEARING_BANK)
         .TextMatrix(i, 76) = IIf(IsNull(TmpRs!PD_CL_BANK_BRANCH), "", TmpRs!PD_CL_BANK_BRANCH)
         .TextMatrix(i, 77) = IIf(IsNull(TmpRs!PD_IFSC_CODE), "", TmpRs!PD_IFSC_CODE)
         .TextMatrix(i, 78) = IIf(IsNull(TmpRs!PD_INSTRUMENT_NO), "", TmpRs!PD_INSTRUMENT_NO)
         .TextMatrix(i, 79) = IIf(IsNull(TmpRs!PD_INSTRUMENT_TYPE), "", TmpRs!PD_INSTRUMENT_TYPE)
         .TextMatrix(i, 80) = IIf(IsNull(TmpRs!PRCNT_REBATE), "", TmpRs!PRCNT_REBATE)
         .TextMatrix(i, 81) = IIf(IsNull(TmpRs!RO_ADDRESS), "", TmpRs!RO_ADDRESS)
         .TextMatrix(i, 82) = IIf(IsNull(TmpRs!RO_ALTERNATE_TEL), "", TmpRs!RO_ALTERNATE_TEL)
         .TextMatrix(i, 83) = IIf(IsNull(TmpRs!RO_CITY), "", TmpRs!RO_CITY)
         .TextMatrix(i, 84) = IIf(IsNull(TmpRs!RO_EMAIL_ID), "", TmpRs!RO_EMAIL_ID)
         .TextMatrix(i, 85) = IIf(IsNull(TmpRs!RO_FAX), "", TmpRs!RO_FAX)
         .TextMatrix(i, 86) = IIf(IsNull(TmpRs!RO_PINCODE), "", TmpRs!RO_PINCODE)
         .TextMatrix(i, 87) = IIf(IsNull(TmpRs!RO_TEL), "", TmpRs!RO_TEL)
         .TextMatrix(i, 88) = IIf(IsNull(TmpRs!RO_TELEX), "", TmpRs!RO_TELEX)
         .TextMatrix(i, 89) = IIf(IsNull(TmpRs!TAN_NO), "", TmpRs!TAN_NO)
         .TextMatrix(i, 90) = IIf(IsNull(TmpRs!FIRST_CLIENT_ADDRESS), "", TmpRs!FIRST_CLIENT_ADDRESS)
         .TextMatrix(i, 91) = IIf(IsNull(TmpRs!FIRST_CITY), "", TmpRs!FIRST_CITY)
         .TextMatrix(i, 92) = IIf(IsNull(TmpRs!FIRST_FAT_HUBBY), "", TmpRs!FIRST_FAT_HUBBY)
         .TextMatrix(i, 93) = IIf(IsNull(TmpRs!FIRST_HOLDER_NAME), "", TmpRs!FIRST_HOLDER_NAME)
         .TextMatrix(i, 94) = IIf(IsNull(TmpRs!FIRST_OCCUPATION), "", TmpRs!FIRST_OCCUPATION)
         .TextMatrix(i, 95) = IIf(IsNull(TmpRs!FIRST_PAN), "", TmpRs!FIRST_PAN)
         .TextMatrix(i, 96) = IIf(IsNull(TmpRs!FIRST_PINCODE), "", TmpRs!FIRST_PINCODE)
         .TextMatrix(i, 97) = IIf(IsNull(TmpRs!MODE_OF_OPERATION), "", TmpRs!MODE_OF_OPERATION)
         .TextMatrix(i, 98) = IIf(IsNull(TmpRs!NOMINEE_ADDRESS), "", TmpRs!NOMINEE_ADDRESS)
         .TextMatrix(i, 99) = IIf(IsNull(TmpRs!NOMINEE_AGE), "", TmpRs!NOMINEE_AGE)
         .TextMatrix(i, 100) = IIf(IsNull(TmpRs!NOMINEE_CITY), "", TmpRs!NOMINEE_CITY)
         .TextMatrix(i, 101) = IIf(IsNull(TmpRs!NOMINEE_NAME), "", TmpRs!NOMINEE_NAME)
         .TextMatrix(i, 102) = IIf(IsNull(TmpRs!NOMINEE_OCCUPATION), "", TmpRs!NOMINEE_OCCUPATION)
         .TextMatrix(i, 103) = IIf(IsNull(TmpRs!NOMINEE_PINCODE), "", TmpRs!NOMINEE_PINCODE)
         .TextMatrix(i, 104) = IIf(IsNull(TmpRs!NOMINEE_RELATION), "", TmpRs!NOMINEE_RELATION)
         .TextMatrix(i, 105) = IIf(IsNull(TmpRs!SECOND_CITY), "", TmpRs!SECOND_CITY)
         .TextMatrix(i, 106) = IIf(IsNull(TmpRs!SECOND_CLIENT_ADDRESS), "", TmpRs!SECOND_CLIENT_ADDRESS)
         .TextMatrix(i, 107) = IIf(IsNull(TmpRs!SECOND_FAT_HUBBY), "", TmpRs!SECOND_FAT_HUBBY)
         .TextMatrix(i, 108) = IIf(IsNull(TmpRs!SECOND_HOLDER_NAME), "", TmpRs!SECOND_HOLDER_NAME)
         .TextMatrix(i, 109) = IIf(IsNull(TmpRs!SECOND_OCCUPATION), "", TmpRs!SECOND_OCCUPATION)
         .TextMatrix(i, 110) = IIf(IsNull(TmpRs!SECOND_PAN), "", TmpRs!SECOND_PAN)
         .TextMatrix(i, 111) = IIf(IsNull(TmpRs!SECOND_PINCODE), "", TmpRs!SECOND_PINCODE)
         .TextMatrix(i, 112) = IIf(IsNull(TmpRs!THIRD_CITY), "", TmpRs!THIRD_CITY)
         .TextMatrix(i, 113) = IIf(IsNull(TmpRs!THIRD_CLIENT_ADDRESS), "", TmpRs!THIRD_CLIENT_ADDRESS)
         .TextMatrix(i, 114) = IIf(IsNull(TmpRs!THIRD_FAT_HUBBY), "", TmpRs!THIRD_FAT_HUBBY)
         .TextMatrix(i, 115) = IIf(IsNull(TmpRs!THIRD_HOLDER_NAME), "", TmpRs!THIRD_HOLDER_NAME)
         .TextMatrix(i, 116) = IIf(IsNull(TmpRs!THIRD_OCCUPATION), "", TmpRs!THIRD_OCCUPATION)
         .TextMatrix(i, 117) = IIf(IsNull(TmpRs!THIRD_PAN), "", TmpRs!THIRD_PAN)
         .TextMatrix(i, 118) = IIf(IsNull(TmpRs!THIRD_PINCODE), "", TmpRs!THIRD_PINCODE)
         .TextMatrix(i, 119) = IIf(IsNull(TmpRs!UPLOAD_FOR_FIRST_HOLDER), "", TmpRs!UPLOAD_FOR_FIRST_HOLDER)
         .TextMatrix(i, 120) = IIf(IsNull(TmpRs!UPLOAD_FOR_SECOND_HOLDER), "", TmpRs!UPLOAD_FOR_SECOND_HOLDER)
         .TextMatrix(i, 121) = IIf(IsNull(TmpRs!UPLOAD_FOR_THIRD_HOLDER), "", TmpRs!UPLOAD_FOR_THIRD_HOLDER)
         .TextMatrix(i, 122) = IIf(IsNull(TmpRs!ORG_CLIENT_TYPE), "", TmpRs!ORG_CLIENT_TYPE)
         .TextMatrix(i, 123) = IIf(IsNull(TmpRs!ACCOUNT_VALIDITY_DATE), "", TmpRs!ACCOUNT_VALIDITY_DATE)
         .TextMatrix(i, 124) = IIf(IsNull(TmpRs!ACCOUNT_OPENING_CHARGES), "", TmpRs!ACCOUNT_OPENING_CHARGES)
         .TextMatrix(i, 125) = IIf(IsNull(TmpRs!RENEWAL_CHARGES), "", TmpRs!RENEWAL_CHARGES)
         .TextMatrix(i, 126) = IIf(IsNull(TmpRs!BRANCH_CODE), "", TmpRs!BRANCH_CODE)
         ProgressBar1.Value = ProgressBar1.Value + 1
         TmpRs.MoveNext
      Next
   End With
End If
End Sub

Private Sub Get_USER_MSTR()
tmp = " select"
tmp = tmp & " Employeeno  'USER_ID', '' ORG_ID, '' ORG_NAME, '' ORG_CATEGORY, EmployeeName 'USER_NAME', EmployeeEmailID USER_EMAIL,"
tmp = tmp & " MobileNo USER_CONTACT, '' USER_PHONE, Address STD_ADDRESS, '' USER_ROLES, '' MAKER, '' CHECKER, '' WH_CODE,"
tmp = tmp & " '' DOCUMENT_NO, '' DOCUMENT_TYPE, '' USER_MOBILE, '' USER_PASSWORD, '' 'PASSWORD', '' USER_CATEGORY, '' FIRST_LOGIN_FLAG,"
tmp = tmp & " '' CRT_USR_ID, '' CRT_TM, '' MOD_USR_ID, '' MOD_TM, '' USER_COMM_GROUP, '' USER_LANGUAGE, '' IS_FREEZED, 'A' IS_ACTIVE,"
tmp = tmp & " '' ORG_ROLE, '' DEALER_CM, '' DEALER_DM, '' DEALER_MM, '' DEALER_AO, '' DEALER_BIDS, '' ACC_CREATION, '' FUNDS, '' CST,"
tmp = tmp & " '' SUPER, '' CMS, '' BANK, '' FIN, '' ACTIVE, '' LAST_LOGIN_TIME, '' PASSWORD_EXPIRY_DATE, '' IS_LOGGED_IN, '' EXPOSURE_LIMIT,"
tmp = tmp & " '' CRT_SYS_NM, '' DEALER_WH"
tmp = tmp & " from mst_employee where employeeno in ('1308', '1648', '2118', '1714', '1729')"
Call TmpTable
If TmpRs.RecordCount > 0 Then
   ProgressBar1.Value = 0
   ProgressBar1.Max = TmpRs.RecordCount + 3
   With MSHFlexGrid3
      .Rows = .Rows + TmpRs.RecordCount
      For i = 1 To TmpRs.RecordCount
         .TextMatrix(i, 0) = IIf(IsNull(TmpRs!USER_ID), "", TmpRs!USER_ID)
         .TextMatrix(i, 1) = IIf(IsNull(TmpRs!ORG_ID), "", TmpRs!ORG_ID)
         .TextMatrix(i, 2) = IIf(IsNull(TmpRs!ORG_NAME), "", TmpRs!ORG_NAME)
         .TextMatrix(i, 3) = IIf(IsNull(TmpRs!ORG_CATEGORY), "", TmpRs!ORG_CATEGORY)
         .TextMatrix(i, 4) = IIf(IsNull(TmpRs!USER_NAME), "", TmpRs!USER_NAME)
         .TextMatrix(i, 5) = IIf(IsNull(TmpRs!USER_EMAIL), "", TmpRs!USER_EMAIL)
         .TextMatrix(i, 6) = IIf(IsNull(TmpRs!USER_CONTACT), "", TmpRs!USER_CONTACT)
         .TextMatrix(i, 7) = IIf(IsNull(TmpRs!USER_PHONE), "", TmpRs!USER_PHONE)
         .TextMatrix(i, 8) = IIf(IsNull(TmpRs!STD_ADDRESS), "", TmpRs!STD_ADDRESS)
         .TextMatrix(i, 9) = IIf(IsNull(TmpRs!USER_ROLES), "", TmpRs!USER_ROLES)
         .TextMatrix(i, 10) = IIf(IsNull(TmpRs!MAKER), "", TmpRs!MAKER)
         .TextMatrix(i, 11) = IIf(IsNull(TmpRs!CHECKER), "", TmpRs!CHECKER)
         .TextMatrix(i, 12) = IIf(IsNull(TmpRs!WH_CODE), "", TmpRs!WH_CODE)
         .TextMatrix(i, 13) = IIf(IsNull(TmpRs!DOCUMENT_NO), "", TmpRs!DOCUMENT_NO)
         .TextMatrix(i, 14) = IIf(IsNull(TmpRs!DOCUMENT_TYPE), "", TmpRs!DOCUMENT_TYPE)
         .TextMatrix(i, 15) = IIf(IsNull(TmpRs!USER_MOBILE), "", TmpRs!USER_MOBILE)
         .TextMatrix(i, 16) = IIf(IsNull(TmpRs!USER_PASSWORD), "", TmpRs!USER_PASSWORD)
         .TextMatrix(i, 17) = IIf(IsNull(TmpRs!Password), "", TmpRs!Password)
         .TextMatrix(i, 18) = IIf(IsNull(TmpRs!USER_CATEGORY), "", TmpRs!USER_CATEGORY)
         .TextMatrix(i, 19) = IIf(IsNull(TmpRs!FIRST_LOGIN_FLAG), "", TmpRs!FIRST_LOGIN_FLAG)
         .TextMatrix(i, 20) = IIf(IsNull(TmpRs!CRT_USR_ID), "", TmpRs!CRT_USR_ID)
         .TextMatrix(i, 21) = IIf(IsNull(TmpRs!CRT_TM), "", TmpRs!CRT_TM)
         .TextMatrix(i, 22) = IIf(IsNull(TmpRs!MOD_USR_ID), "", TmpRs!MOD_USR_ID)
         .TextMatrix(i, 23) = IIf(IsNull(TmpRs!MOD_TM), "", TmpRs!MOD_TM)
         .TextMatrix(i, 24) = IIf(IsNull(TmpRs!USER_COMM_GROUP), "", TmpRs!USER_COMM_GROUP)
         .TextMatrix(i, 25) = IIf(IsNull(TmpRs!USER_LANGUAGE), "", TmpRs!USER_LANGUAGE)
         .TextMatrix(i, 26) = IIf(IsNull(TmpRs!IS_FREEZED), "", TmpRs!IS_FREEZED)
         .TextMatrix(i, 27) = IIf(IsNull(TmpRs!IS_ACTIVE), "", TmpRs!IS_ACTIVE)
         .TextMatrix(i, 28) = IIf(IsNull(TmpRs!ORG_ROLE), "", TmpRs!ORG_ROLE)
         .TextMatrix(i, 29) = IIf(IsNull(TmpRs!DEALER_CM), "", TmpRs!DEALER_CM)
         .TextMatrix(i, 30) = IIf(IsNull(TmpRs!DEALER_DM), "", TmpRs!DEALER_DM)
         .TextMatrix(i, 31) = IIf(IsNull(TmpRs!DEALER_MM), "", TmpRs!DEALER_MM)
         .TextMatrix(i, 32) = IIf(IsNull(TmpRs!DEALER_AO), "", TmpRs!DEALER_AO)
         .TextMatrix(i, 33) = IIf(IsNull(TmpRs!DEALER_BIDS), "", TmpRs!DEALER_BIDS)
         .TextMatrix(i, 34) = IIf(IsNull(TmpRs!ACC_CREATION), "", TmpRs!ACC_CREATION)
         .TextMatrix(i, 35) = IIf(IsNull(TmpRs!FUNDS), "", TmpRs!FUNDS)
         .TextMatrix(i, 36) = IIf(IsNull(TmpRs!CST), "", TmpRs!CST)
         .TextMatrix(i, 37) = IIf(IsNull(TmpRs!SUPER), "", TmpRs!SUPER)
         .TextMatrix(i, 38) = IIf(IsNull(TmpRs!CMS), "", TmpRs!CMS)
         .TextMatrix(i, 39) = IIf(IsNull(TmpRs!Bank), "", TmpRs!Bank)
         .TextMatrix(i, 40) = IIf(IsNull(TmpRs!FIN), "", TmpRs!FIN)
         .TextMatrix(i, 41) = IIf(IsNull(TmpRs!Active), "", TmpRs!Active)
         .TextMatrix(i, 42) = IIf(IsNull(TmpRs!LAST_LOGIN_TIME), "", TmpRs!LAST_LOGIN_TIME)
         .TextMatrix(i, 43) = IIf(IsNull(TmpRs!PASSWORD_EXPIRY_DATE), "", TmpRs!PASSWORD_EXPIRY_DATE)
         .TextMatrix(i, 44) = IIf(IsNull(TmpRs!IS_LOGGED_IN), "", TmpRs!IS_LOGGED_IN)
         .TextMatrix(i, 45) = IIf(IsNull(TmpRs!EXPOSURE_LIMIT), "", TmpRs!EXPOSURE_LIMIT)
         .TextMatrix(i, 46) = IIf(IsNull(TmpRs!CRT_SYS_NM), "", TmpRs!CRT_SYS_NM)
         .TextMatrix(i, 47) = IIf(IsNull(TmpRs!DEALER_WH), "", TmpRs!DEALER_WH)
         ProgressBar1.Value = ProgressBar1.Value + 1
         TmpRs.MoveNext
      Next
   End With
End If
End Sub

Private Sub Get_ASS_MSTR()
tmp = " select '1' ASS_ID, 'COMMGRADE' ASS_NAME, '' ASS_EMAIL, '' ASS_CONTACT, '' ASS_PHONE, ''   ASS_PHONE_2, ''   ASS_FAX_NO, '' ASS_ADDRESS, "
tmp = tmp & " '' ASS_DOCUMENT_TYPE, '' ASS_DOCUMENT_NUMBER, '' WH_CODE, '' COMMODITY_CODE, ''  ASS_PASSWORD, '' PASSWORD, '' FIRST_LOGIN_FLAG,"
tmp = tmp & " '' CREATED_BY, '' CREATED_ON, '' MODIFIED_BY, '' MODIFIED_ON, '' ASS_NEW_EMAIL, '' ASS_NAME_0, '' ASS_NAME_1, '' IS_FREEZED, 'A' IS_ACTIVE,"
tmp = tmp & " '' ASSAYER_NAME_1, '' ASS_ADDRESS_0, '' ASS_ADDRESS_1, '' AS_ADDRESS, '' AS_ALTERNATE_TEL, '' AS_CITY, '' AS_DESIGNATION, '' AS_EMAIL_ID,"
tmp = tmp & " '' AS_FAX, '' AS_MOBILE_NO, '' AS_NAME, '' AS_PAN_NAME, '' AS_PINCODE, '' AS_TEL, '' BD_BANK_ACCOUNT_NO, '' BD_BANK_CODE,"
tmp = tmp & " '' BD_BRANCH_CODE, '' BD_IFSC_CODE, '' CA_ALTERNATE_TEL, '' CA_CITY, '' CA_CONTR_PERSON_NAME, '' CA_DESIGNATION, '' CA_PINCODE,"
tmp = tmp & " '' CA_TELEX, '' DATE_OF_BIRTH, '' MEMBERSHIP_TYPE, '' MEMBER_SHIP_ID, '' ORG_CATEGORY, '' ORG_CONSTITUTION, '' PAN_NO,"
tmp = tmp & " '' PD_AMOUNT, '' PD_APPLICANTS_BANK, '' PD_APP_BANK_BRANCH, '' PD_CLEARING_BANK, '' PD_CL_BANK_BRANCH, '' PD_IFSC_CODE, '' PD_INSTRUMENT_NO,"
tmp = tmp & " '' PD_INSTRUMENT_TYPE, '' RO_ADDRESS, '' RO_ALTERNATE_TEL, '' RO_CITY, '' RO_EMAIL_ID, '' RO_FAX, '' RO_PINCODE,"
tmp = tmp & " '' RO_TEL, '' TAN_NO"
Call TmpTable
If TmpRs.RecordCount > 0 Then
   ProgressBar1.Value = 0
   ProgressBar1.Max = TmpRs.RecordCount + 3
   With MSHFlexGrid4
      .Rows = .Rows + TmpRs.RecordCount
      For i = 1 To TmpRs.RecordCount
         .TextMatrix(i, 0) = IIf(IsNull(TmpRs!ASS_ID), "", TmpRs!ASS_ID)
         .TextMatrix(i, 1) = IIf(IsNull(TmpRs!ASS_NAME), "", TmpRs!ASS_NAME)
         .TextMatrix(i, 2) = IIf(IsNull(TmpRs!ASS_EMAIL), "", TmpRs!ASS_EMAIL)
         .TextMatrix(i, 3) = IIf(IsNull(TmpRs!ASS_CONTACT), "", TmpRs!ASS_CONTACT)
         .TextMatrix(i, 4) = IIf(IsNull(TmpRs!ASS_PHONE), "", TmpRs!ASS_PHONE)
         .TextMatrix(i, 5) = IIf(IsNull(TmpRs!ASS_PHONE_2), "", TmpRs!ASS_PHONE_2)
         .TextMatrix(i, 6) = IIf(IsNull(TmpRs!ASS_FAX_NO), "", TmpRs!ASS_FAX_NO)
         .TextMatrix(i, 7) = IIf(IsNull(TmpRs!ASS_ADDRESS), "", TmpRs!ASS_ADDRESS)
         .TextMatrix(i, 8) = IIf(IsNull(TmpRs!ASS_DOCUMENT_TYPE), "", TmpRs!ASS_DOCUMENT_TYPE)
         .TextMatrix(i, 9) = IIf(IsNull(TmpRs!ASS_DOCUMENT_NUMBER), "", TmpRs!ASS_DOCUMENT_NUMBER)
         .TextMatrix(i, 10) = IIf(IsNull(TmpRs!WH_CODE), "", TmpRs!WH_CODE)
         .TextMatrix(i, 11) = IIf(IsNull(TmpRs!COMMODITY_CODE), "", TmpRs!COMMODITY_CODE)
         .TextMatrix(i, 12) = IIf(IsNull(TmpRs!ASS_PASSWORD), "", TmpRs!ASS_PASSWORD)
         .TextMatrix(i, 13) = IIf(IsNull(TmpRs!Password), "", TmpRs!Password)
         .TextMatrix(i, 14) = IIf(IsNull(TmpRs!FIRST_LOGIN_FLAG), "", TmpRs!FIRST_LOGIN_FLAG)
         .TextMatrix(i, 15) = IIf(IsNull(TmpRs!CREATED_BY), "", TmpRs!CREATED_BY)
         .TextMatrix(i, 16) = IIf(IsNull(TmpRs!CREATED_ON), "", TmpRs!CREATED_ON)
         .TextMatrix(i, 17) = IIf(IsNull(TmpRs!MODIFIED_BY), "", TmpRs!MODIFIED_BY)
         .TextMatrix(i, 18) = IIf(IsNull(TmpRs!MODIFIED_ON), "", TmpRs!MODIFIED_ON)
         .TextMatrix(i, 19) = IIf(IsNull(TmpRs!ASS_NEW_EMAIL), "", TmpRs!ASS_NEW_EMAIL)
         .TextMatrix(i, 20) = IIf(IsNull(TmpRs!ASS_NAME_0), "", TmpRs!ASS_NAME_0)
         .TextMatrix(i, 21) = IIf(IsNull(TmpRs!ASS_NAME_1), "", TmpRs!ASS_NAME_1)
         .TextMatrix(i, 22) = IIf(IsNull(TmpRs!IS_FREEZED), "", TmpRs!IS_FREEZED)
         .TextMatrix(i, 23) = IIf(IsNull(TmpRs!IS_ACTIVE), "", TmpRs!IS_ACTIVE)
         .TextMatrix(i, 24) = IIf(IsNull(TmpRs!ASSAYER_NAME_1), "", TmpRs!ASSAYER_NAME_1)
         .TextMatrix(i, 25) = IIf(IsNull(TmpRs!ASS_ADDRESS_0), "", TmpRs!ASS_ADDRESS_0)
         .TextMatrix(i, 26) = IIf(IsNull(TmpRs!ASS_ADDRESS_1), "", TmpRs!ASS_ADDRESS_1)
         .TextMatrix(i, 27) = IIf(IsNull(TmpRs!AS_ADDRESS), "", TmpRs!AS_ADDRESS)
         .TextMatrix(i, 28) = IIf(IsNull(TmpRs!AS_ALTERNATE_TEL), "", TmpRs!AS_ALTERNATE_TEL)
         .TextMatrix(i, 29) = IIf(IsNull(TmpRs!AS_CITY), "", TmpRs!AS_CITY)
         .TextMatrix(i, 30) = IIf(IsNull(TmpRs!AS_DESIGNATION), "", TmpRs!AS_DESIGNATION)
         .TextMatrix(i, 31) = IIf(IsNull(TmpRs!AS_EMAIL_ID), "", TmpRs!AS_EMAIL_ID)
         .TextMatrix(i, 32) = IIf(IsNull(TmpRs!AS_FAX), "", TmpRs!AS_FAX)
         .TextMatrix(i, 33) = IIf(IsNull(TmpRs!AS_MOBILE_NO), "", TmpRs!AS_MOBILE_NO)
         .TextMatrix(i, 34) = IIf(IsNull(TmpRs!AS_NAME), "", TmpRs!AS_NAME)
         .TextMatrix(i, 35) = IIf(IsNull(TmpRs!AS_PAN_NAME), "", TmpRs!AS_PAN_NAME)
         .TextMatrix(i, 36) = IIf(IsNull(TmpRs!AS_PINCODE), "", TmpRs!AS_PINCODE)
         .TextMatrix(i, 37) = IIf(IsNull(TmpRs!AS_TEL), "", TmpRs!AS_TEL)
         .TextMatrix(i, 38) = IIf(IsNull(TmpRs!BD_BANK_ACCOUNT_NO), "", TmpRs!BD_BANK_ACCOUNT_NO)
         .TextMatrix(i, 39) = IIf(IsNull(TmpRs!BD_BANK_CODE), "", TmpRs!BD_BANK_CODE)
         .TextMatrix(i, 40) = IIf(IsNull(TmpRs!BD_BRANCH_CODE), "", TmpRs!BD_BRANCH_CODE)
         .TextMatrix(i, 41) = IIf(IsNull(TmpRs!BD_IFSC_CODE), "", TmpRs!BD_IFSC_CODE)
         .TextMatrix(i, 42) = IIf(IsNull(TmpRs!CA_ALTERNATE_TEL), "", TmpRs!CA_ALTERNATE_TEL)
         .TextMatrix(i, 43) = IIf(IsNull(TmpRs!CA_CITY), "", TmpRs!CA_CITY)
         .TextMatrix(i, 44) = IIf(IsNull(TmpRs!CA_CONTR_PERSON_NAME), "", TmpRs!CA_CONTR_PERSON_NAME)
         .TextMatrix(i, 45) = IIf(IsNull(TmpRs!CA_DESIGNATION), "", TmpRs!CA_DESIGNATION)
         .TextMatrix(i, 46) = IIf(IsNull(TmpRs!CA_PINCODE), "", TmpRs!CA_PINCODE)
         .TextMatrix(i, 47) = IIf(IsNull(TmpRs!CA_TELEX), "", TmpRs!CA_TELEX)
         .TextMatrix(i, 48) = IIf(IsNull(TmpRs!DATE_OF_BIRTH), "", TmpRs!DATE_OF_BIRTH)
         .TextMatrix(i, 49) = IIf(IsNull(TmpRs!MEMBERSHIP_TYPE), "", TmpRs!MEMBERSHIP_TYPE)
         .TextMatrix(i, 50) = IIf(IsNull(TmpRs!MEMBER_SHIP_ID), "", TmpRs!MEMBER_SHIP_ID)
         .TextMatrix(i, 51) = IIf(IsNull(TmpRs!ORG_CATEGORY), "", TmpRs!ORG_CATEGORY)
         .TextMatrix(i, 52) = IIf(IsNull(TmpRs!ORG_CONSTITUTION), "", TmpRs!ORG_CONSTITUTION)
         .TextMatrix(i, 53) = IIf(IsNull(TmpRs!PAN_NO), "", TmpRs!PAN_NO)
         .TextMatrix(i, 54) = IIf(IsNull(TmpRs!PD_AMOUNT), "", TmpRs!PD_AMOUNT)
         .TextMatrix(i, 55) = IIf(IsNull(TmpRs!PD_APPLICANTS_BANK), "", TmpRs!PD_APPLICANTS_BANK)
         .TextMatrix(i, 56) = IIf(IsNull(TmpRs!PD_APP_BANK_BRANCH), "", TmpRs!PD_APP_BANK_BRANCH)
         .TextMatrix(i, 57) = IIf(IsNull(TmpRs!PD_CLEARING_BANK), "", TmpRs!PD_CLEARING_BANK)
         .TextMatrix(i, 58) = IIf(IsNull(TmpRs!PD_CL_BANK_BRANCH), "", TmpRs!PD_CL_BANK_BRANCH)
         .TextMatrix(i, 59) = IIf(IsNull(TmpRs!PD_IFSC_CODE), "", TmpRs!PD_IFSC_CODE)
         .TextMatrix(i, 60) = IIf(IsNull(TmpRs!PD_INSTRUMENT_NO), "", TmpRs!PD_INSTRUMENT_NO)
         .TextMatrix(i, 61) = IIf(IsNull(TmpRs!PD_INSTRUMENT_TYPE), "", TmpRs!PD_INSTRUMENT_TYPE)
         .TextMatrix(i, 62) = IIf(IsNull(TmpRs!RO_ADDRESS), "", TmpRs!RO_ADDRESS)
         .TextMatrix(i, 63) = IIf(IsNull(TmpRs!RO_ALTERNATE_TEL), "", TmpRs!RO_ALTERNATE_TEL)
         .TextMatrix(i, 64) = IIf(IsNull(TmpRs!RO_CITY), "", TmpRs!RO_CITY)
         .TextMatrix(i, 65) = IIf(IsNull(TmpRs!RO_EMAIL_ID), "", TmpRs!RO_EMAIL_ID)
         .TextMatrix(i, 66) = IIf(IsNull(TmpRs!RO_FAX), "", TmpRs!RO_FAX)
         .TextMatrix(i, 67) = IIf(IsNull(TmpRs!RO_PINCODE), "", TmpRs!RO_PINCODE)
         .TextMatrix(i, 68) = IIf(IsNull(TmpRs!RO_TEL), "", TmpRs!RO_TEL)
         .TextMatrix(i, 69) = IIf(IsNull(TmpRs!TAN_NO), "", TmpRs!TAN_NO)
         ProgressBar1.Value = ProgressBar1.Value + 1
         TmpRs.MoveNext
      Next
   End With
End If
End Sub

Private Sub Get_LOCATION_MSTR()
tmp = " select  upper(ml.sm_prefix) LOCATION_CODE,  upper(ml.locationname) LOCATION_NAME,  '' WH_CODE,  '' IS_FREEZED,  'A' IS_ACTIVE,"
tmp = tmp & " upper(ml.locationname) LOCATION_NAME_1 from mst_location ml"
Call TmpTable
If TmpRs.RecordCount > 0 Then
   ProgressBar1.Value = 0
   ProgressBar1.Max = TmpRs.RecordCount + 3
   With MSHFlexGrid5
      .Rows = .Rows + TmpRs.RecordCount
      For i = 1 To TmpRs.RecordCount
         .TextMatrix(i, 0) = IIf(IsNull(TmpRs!LOCATION_CODE), "", TmpRs!LOCATION_CODE)
         .TextMatrix(i, 1) = IIf(IsNull(TmpRs!LOCATION_NAME), "", TmpRs!LOCATION_NAME)
         .TextMatrix(i, 2) = IIf(IsNull(TmpRs!WH_CODE), "", TmpRs!WH_CODE)
         .TextMatrix(i, 3) = IIf(IsNull(TmpRs!IS_FREEZED), "", TmpRs!IS_FREEZED)
         .TextMatrix(i, 4) = IIf(IsNull(TmpRs!IS_ACTIVE), "", TmpRs!IS_ACTIVE)
         .TextMatrix(i, 5) = IIf(IsNull(TmpRs!LOCATION_NAME_1), "", TmpRs!LOCATION_NAME_1)
         TmpRs.MoveNext
         ProgressBar1.Value = ProgressBar1.Value + 1
      Next
   End With
End If
End Sub

Private Sub Get_WH_MSTR()
tmp = " select upper(cmp.cmpname)+'_'+upper(mg.godownno) WH_CODE, '' WH_SERVICE_PROVIDER, upper(cmp.cmpname)+'_'+upper(mg.godownno) WH_NAME, upper(ml.sm_prefix) LOCATION_CODE, "
tmp = tmp & " '' WH_LANG_CODE, mg.address WH_ADDRESS, '' WH_MODE, mg.GodownCapacity WH_CAPACITY, '' WH_NAME_0, '' WH_ADDRESS_0, '' WH_NAME_1,"
tmp = tmp & " '' WH_ADDRESS_1, '' IS_FREEZED, 'A' IS_ACTIVE, '' PLEDGING_ALLOWED, mg.address AS_ADDRESS, cmp.telephoneno AS_ALTERNATE_TEL,"
tmp = tmp & " ml.locationname AS_CITY, '' AS_DESIGNATION, '' AS_EMAIL_ID, cmp.faxno AS_FAX, cmp.mobileno AS_MOBILE_NO, '' AS_NAME, '' AS_PAN_NAME,"
tmp = tmp & " '' AS_PINCODE, '' AS_TEL, '' BD_BANK_ACCOUNT_NO, '' BD_BANK_CODE, '' BD_BRANCH_CODE, '' BD_IFSC_CODE, '' BOUND_OR_UNBOUND,"
tmp = tmp & " '' CA_ALTERNATE_TEL, '' CA_CITY, '' CA_CONTR_PERSON_NAME, '' CA_DESIGNATION, '' CA_EMAIL_ID, '' CA_FAX, '' CA_MOBILE_NO, '' CA_PINCODE,"
tmp = tmp & " '' CA_TEL, '' CA_TELEX, '' CREATED_BY, '' CREATED_ON, '' DATE_OF_BIRTH, '' FIRST_LOGIN_FLAG, '' MEMBERSHIP_TYPE, '' MEMBER_SHIP_ID,"
tmp = tmp & " '' MODIFIED_BY, '' MODIFIED_ON,'' ORG_CONSTITUTION, '' PAN_NO, '' PASSWORD, '' PD_AMOUNT, '' PD_APPLICANTS_BANK, '' PD_APP_BANK_BRANCH,"
tmp = tmp & " '' PD_CLEARING_BANK, '' PD_CL_BANK_BRANCH, '' PD_IFSC_CODE, '' PD_INSTRUMENT_NO, '' PD_INSTRUMENT_TYPE, '' RO_ADDRESS,"
tmp = tmp & " '' RO_ALTERNATE_TEL, '' RO_CITY, '' RO_EMAIL_ID, '' RO_FAX, '' RO_PINCODE, '' RO_TEL, '' RO_TELEX, '' TAN_NO,"
tmp = tmp & " '' WH_DOCUMENT_NUMBER,'' WH_DOCUMENT_TYPE, '' WH_LICENSE, '' WH_LICENSE_VALIDITY, '' WH_TYPE, '' WH_NEW_EMAIL, '' WH_PASSWORD,"
tmp = tmp & " '' WH_PHONE_2, '' WH_CATEGORY, '' INSURANCE_VALID_DATE, '' ORG_CATEGORY, '' EXCHANGE_TYPE"
tmp = tmp & " from mst_godown mg inner join mst_cmp cmp on cmp.cmpid = mg.cmpid inner join mst_location ml on ml.locationid  = cmp.locationid"
Call TmpTable
If TmpRs.RecordCount > 0 Then
   ProgressBar1.Value = 0
   ProgressBar1.Max = TmpRs.RecordCount + 3
   With MSHFlexGrid6
      .Rows = .Rows + TmpRs.RecordCount
      For i = 1 To TmpRs.RecordCount
         .TextMatrix(i, 0) = IIf(IsNull(TmpRs!WH_CODE), "", TmpRs!WH_CODE)
         .TextMatrix(i, 1) = IIf(IsNull(TmpRs!WH_SERVICE_PROVIDER), "", TmpRs!WH_SERVICE_PROVIDER)
         .TextMatrix(i, 2) = IIf(IsNull(TmpRs!WH_NAME), "", TmpRs!WH_NAME)
         .TextMatrix(i, 3) = IIf(IsNull(TmpRs!LOCATION_CODE), "", TmpRs!LOCATION_CODE)
         .TextMatrix(i, 4) = IIf(IsNull(TmpRs!WH_LANG_CODE), "", TmpRs!WH_LANG_CODE)
         .TextMatrix(i, 5) = IIf(IsNull(TmpRs!WH_ADDRESS), "", TmpRs!WH_ADDRESS)
         .TextMatrix(i, 6) = IIf(IsNull(TmpRs!WH_MODE), "", TmpRs!WH_MODE)
         .TextMatrix(i, 7) = IIf(IsNull(TmpRs!WH_CAPACITY), "", TmpRs!WH_CAPACITY)
         .TextMatrix(i, 8) = IIf(IsNull(TmpRs!WH_NAME_0), "", TmpRs!WH_NAME_0)
         .TextMatrix(i, 9) = IIf(IsNull(TmpRs!WH_ADDRESS_0), "", TmpRs!WH_ADDRESS_0)
         .TextMatrix(i, 10) = IIf(IsNull(TmpRs!WH_NAME_1), "", TmpRs!WH_NAME_1)
         .TextMatrix(i, 11) = IIf(IsNull(TmpRs!WH_ADDRESS_1), "", TmpRs!WH_ADDRESS_1)
         .TextMatrix(i, 12) = IIf(IsNull(TmpRs!IS_FREEZED), "", TmpRs!IS_FREEZED)
         .TextMatrix(i, 13) = IIf(IsNull(TmpRs!IS_ACTIVE), "", TmpRs!IS_ACTIVE)
         .TextMatrix(i, 14) = IIf(IsNull(TmpRs!PLEDGING_ALLOWED), "", TmpRs!PLEDGING_ALLOWED)
         .TextMatrix(i, 15) = IIf(IsNull(TmpRs!AS_ADDRESS), "", TmpRs!AS_ADDRESS)
         .TextMatrix(i, 16) = IIf(IsNull(TmpRs!AS_ALTERNATE_TEL), "", TmpRs!AS_ALTERNATE_TEL)
         .TextMatrix(i, 17) = IIf(IsNull(TmpRs!AS_CITY), "", TmpRs!AS_CITY)
         .TextMatrix(i, 18) = IIf(IsNull(TmpRs!AS_DESIGNATION), "", TmpRs!AS_DESIGNATION)
         .TextMatrix(i, 19) = IIf(IsNull(TmpRs!AS_EMAIL_ID), "", TmpRs!AS_EMAIL_ID)
         .TextMatrix(i, 20) = IIf(IsNull(TmpRs!AS_FAX), "", TmpRs!AS_FAX)
         .TextMatrix(i, 21) = IIf(IsNull(TmpRs!AS_MOBILE_NO), "", TmpRs!AS_MOBILE_NO)
         .TextMatrix(i, 22) = IIf(IsNull(TmpRs!AS_NAME), "", TmpRs!AS_NAME)
         .TextMatrix(i, 23) = IIf(IsNull(TmpRs!AS_PAN_NAME), "", TmpRs!AS_PAN_NAME)
         .TextMatrix(i, 24) = IIf(IsNull(TmpRs!AS_PINCODE), "", TmpRs!AS_PINCODE)
         .TextMatrix(i, 25) = IIf(IsNull(TmpRs!AS_TEL), "", TmpRs!AS_TEL)
         .TextMatrix(i, 26) = IIf(IsNull(TmpRs!BD_BANK_ACCOUNT_NO), "", TmpRs!BD_BANK_ACCOUNT_NO)
         .TextMatrix(i, 27) = IIf(IsNull(TmpRs!BD_BANK_CODE), "", TmpRs!BD_BANK_CODE)
         .TextMatrix(i, 28) = IIf(IsNull(TmpRs!BD_BRANCH_CODE), "", TmpRs!BD_BRANCH_CODE)
         .TextMatrix(i, 29) = IIf(IsNull(TmpRs!BD_IFSC_CODE), "", TmpRs!BD_IFSC_CODE)
         .TextMatrix(i, 30) = IIf(IsNull(TmpRs!BOUND_OR_UNBOUND), "", TmpRs!BOUND_OR_UNBOUND)
         .TextMatrix(i, 31) = IIf(IsNull(TmpRs!CA_ALTERNATE_TEL), "", TmpRs!CA_ALTERNATE_TEL)
         .TextMatrix(i, 32) = IIf(IsNull(TmpRs!CA_CITY), "", TmpRs!CA_CITY)
         .TextMatrix(i, 33) = IIf(IsNull(TmpRs!CA_CONTR_PERSON_NAME), "", TmpRs!CA_CONTR_PERSON_NAME)
         .TextMatrix(i, 34) = IIf(IsNull(TmpRs!CA_DESIGNATION), "", TmpRs!CA_DESIGNATION)
         .TextMatrix(i, 35) = IIf(IsNull(TmpRs!CA_EMAIL_ID), "", TmpRs!CA_EMAIL_ID)
         .TextMatrix(i, 36) = IIf(IsNull(TmpRs!CA_FAX), "", TmpRs!CA_FAX)
         .TextMatrix(i, 37) = IIf(IsNull(TmpRs!CA_MOBILE_NO), "", TmpRs!CA_MOBILE_NO)
         .TextMatrix(i, 38) = IIf(IsNull(TmpRs!CA_PINCODE), "", TmpRs!CA_PINCODE)
         .TextMatrix(i, 39) = IIf(IsNull(TmpRs!CA_TEL), "", TmpRs!CA_TEL)
         .TextMatrix(i, 40) = IIf(IsNull(TmpRs!CA_TELEX), "", TmpRs!CA_TELEX)
         .TextMatrix(i, 41) = IIf(IsNull(TmpRs!CREATED_BY), "", TmpRs!CREATED_BY)
         .TextMatrix(i, 42) = IIf(IsNull(TmpRs!CREATED_ON), "", TmpRs!CREATED_ON)
         .TextMatrix(i, 43) = IIf(IsNull(TmpRs!DATE_OF_BIRTH), "", TmpRs!DATE_OF_BIRTH)
         .TextMatrix(i, 44) = IIf(IsNull(TmpRs!FIRST_LOGIN_FLAG), "", TmpRs!FIRST_LOGIN_FLAG)
         .TextMatrix(i, 45) = IIf(IsNull(TmpRs!MEMBERSHIP_TYPE), "", TmpRs!MEMBERSHIP_TYPE)
         .TextMatrix(i, 46) = IIf(IsNull(TmpRs!MEMBER_SHIP_ID), "", TmpRs!MEMBER_SHIP_ID)
         .TextMatrix(i, 47) = IIf(IsNull(TmpRs!MODIFIED_BY), "", TmpRs!MODIFIED_BY)
         .TextMatrix(i, 48) = IIf(IsNull(TmpRs!MODIFIED_ON), "", TmpRs!MODIFIED_ON)
         .TextMatrix(i, 49) = IIf(IsNull(TmpRs!ORG_CONSTITUTION), "", TmpRs!ORG_CONSTITUTION)
         .TextMatrix(i, 50) = IIf(IsNull(TmpRs!PAN_NO), "", TmpRs!PAN_NO)
         .TextMatrix(i, 51) = IIf(IsNull(TmpRs!Password), "", TmpRs!Password)
         .TextMatrix(i, 52) = IIf(IsNull(TmpRs!PD_AMOUNT), "", TmpRs!PD_AMOUNT)
         .TextMatrix(i, 53) = IIf(IsNull(TmpRs!PD_APPLICANTS_BANK), "", TmpRs!PD_APPLICANTS_BANK)
         .TextMatrix(i, 54) = IIf(IsNull(TmpRs!PD_APP_BANK_BRANCH), "", TmpRs!PD_APP_BANK_BRANCH)
         .TextMatrix(i, 55) = IIf(IsNull(TmpRs!PD_CLEARING_BANK), "", TmpRs!PD_CLEARING_BANK)
         .TextMatrix(i, 56) = IIf(IsNull(TmpRs!PD_CL_BANK_BRANCH), "", TmpRs!PD_CL_BANK_BRANCH)
         .TextMatrix(i, 57) = IIf(IsNull(TmpRs!PD_IFSC_CODE), "", TmpRs!PD_IFSC_CODE)
         .TextMatrix(i, 58) = IIf(IsNull(TmpRs!PD_INSTRUMENT_NO), "", TmpRs!PD_INSTRUMENT_NO)
         .TextMatrix(i, 59) = IIf(IsNull(TmpRs!PD_INSTRUMENT_TYPE), "", TmpRs!PD_INSTRUMENT_TYPE)
         .TextMatrix(i, 60) = IIf(IsNull(TmpRs!RO_ADDRESS), "", TmpRs!RO_ADDRESS)
         .TextMatrix(i, 61) = IIf(IsNull(TmpRs!RO_ALTERNATE_TEL), "", TmpRs!RO_ALTERNATE_TEL)
         .TextMatrix(i, 62) = IIf(IsNull(TmpRs!RO_CITY), "", TmpRs!RO_CITY)
         .TextMatrix(i, 63) = IIf(IsNull(TmpRs!RO_EMAIL_ID), "", TmpRs!RO_EMAIL_ID)
         .TextMatrix(i, 64) = IIf(IsNull(TmpRs!RO_FAX), "", TmpRs!RO_FAX)
         .TextMatrix(i, 65) = IIf(IsNull(TmpRs!RO_PINCODE), "", TmpRs!RO_PINCODE)
         .TextMatrix(i, 66) = IIf(IsNull(TmpRs!RO_TEL), "", TmpRs!RO_TEL)
         .TextMatrix(i, 67) = IIf(IsNull(TmpRs!RO_TELEX), "", TmpRs!RO_TELEX)
         .TextMatrix(i, 68) = IIf(IsNull(TmpRs!TAN_NO), "", TmpRs!TAN_NO)
         .TextMatrix(i, 69) = IIf(IsNull(TmpRs!WH_DOCUMENT_NUMBER), "", TmpRs!WH_DOCUMENT_NUMBER)
         .TextMatrix(i, 70) = IIf(IsNull(TmpRs!WH_DOCUMENT_TYPE), "", TmpRs!WH_DOCUMENT_TYPE)
         .TextMatrix(i, 71) = IIf(IsNull(TmpRs!WH_LICENSE), "", TmpRs!WH_LICENSE)
         .TextMatrix(i, 72) = IIf(IsNull(TmpRs!WH_LICENSE_VALIDITY), "", TmpRs!WH_LICENSE_VALIDITY)
         .TextMatrix(i, 73) = IIf(IsNull(TmpRs!WH_TYPE), "", TmpRs!WH_TYPE)
         .TextMatrix(i, 74) = IIf(IsNull(TmpRs!WH_NEW_EMAIL), "", TmpRs!WH_NEW_EMAIL)
         .TextMatrix(i, 75) = IIf(IsNull(TmpRs!WH_PASSWORD), "", TmpRs!WH_PASSWORD)
         .TextMatrix(i, 76) = IIf(IsNull(TmpRs!WH_PHONE_2), "", TmpRs!WH_PHONE_2)
         .TextMatrix(i, 77) = IIf(IsNull(TmpRs!WH_CATEGORY), "", TmpRs!WH_CATEGORY)
         .TextMatrix(i, 78) = IIf(IsNull(TmpRs!INSURANCE_VALID_DATE), "", TmpRs!INSURANCE_VALID_DATE)
         .TextMatrix(i, 79) = IIf(IsNull(TmpRs!ORG_CATEGORY), "", TmpRs!ORG_CATEGORY)
         .TextMatrix(i, 80) = IIf(IsNull(TmpRs!EXCHANGE_TYPE), "", TmpRs!EXCHANGE_TYPE)
         ProgressBar1.Value = ProgressBar1.Value + 1
         TmpRs.MoveNext
      Next
   End With
End If
End Sub

Private Sub Get_WSP_MSTR()
tmp = " select 1 'WSP_ID', 'NCML' WSP_NAME, '' WSP_EMAIL, '' WSP_CONTACT, '02240419191' WSP_PHONE, '' WSP_PHONE_2, '02240419193' WSP_FAX_NO,"
tmp = tmp & " 'Gayatri Towers, 954 , Appasaheb Marathe Marg, Prabhadevi , Mumbai 400 025' WSP_ADDRESS, '' WSP_DOCUMENT_TYPE, '' WSP_DOCUMENT_NUMBER,"
tmp = tmp & " '' WSP_PASSWORD, '' PASSWORD, '' FIRST_LOGIN_FLAG, '' CREATED_BY, '' CREATED_ON, '' MODIFIED_BY, '' MODIFIED_ON, '' WSP_NEW_EMAIL,"
tmp = tmp & " '' ORG_CATEGORY, '' ORG_CONSTITUTION, '' DATE_OF_BIRTH, '' PAN_NO, '' TAN_NO, '' MEMBERSHIP_TYPE, '' RO_ADDRESS, '' RO_TEL,"
tmp = tmp & " '' RO_CITY, '' RO_FAX, '' RO_PINCODE, '' RO_ALTERNATE_TEL, '' RO_EMAIL_ID, '' CA_CITY, '' CA_PINCODE, '' CA_ALTERNATE_TEL,"
tmp = tmp & " '' CA_TELEX, '' CA_CONTR_PERSON_NAME, '' CA_DESIGNATION, '' AS_NAME, '' AS_TEL, '' AS_DESIGNATION, '' AS_ALTERNATE_TEL,"
tmp = tmp & " '' AS_ADDRESS, '' AS_FAX, '' AS_CITY, '' AS_EMAIL_ID, '' AS_PINCODE, '' AS_PAN_NAME, '' AS_MOBILE_NO, '' BD_BANK_CODE, '' BD_BANK_ACCOUNT_NO,"
tmp = tmp & " '' BD_IFSC_CODE, '' BD_BRANCH_CODE, '' PD_AMOUNT, '' PD_INSTRUMENT_TYPE, '' PD_CLEARING_BANK, '' PD_APPLICANTS_BANK, '' PD_CL_BANK_BRANCH,"
tmp = tmp & " '' PD_APP_BANK_BRANCH, '' PD_INSTRUMENT_NO, '' PD_IFSC_CODE, '' MEMBER_SHIP_ID, '' IS_FREEZED, 'A' IS_ACTIVE, '' INSURANCE_AMOUNT,"
tmp = tmp & " '' INSURANCE_POLICY, '' INSURANCE_VALIDITY"
Call TmpTable
If TmpRs.RecordCount > 0 Then
   ProgressBar1.Value = 0
   ProgressBar1.Max = TmpRs.RecordCount + 3
   With MSHFlexGrid7
      .Rows = .Rows + TmpRs.RecordCount
      For i = 1 To TmpRs.RecordCount
         .TextMatrix(i, 0) = IIf(IsNull(TmpRs!WSP_ID), "", TmpRs!WSP_ID)
         .TextMatrix(i, 1) = IIf(IsNull(TmpRs!WSP_NAME), "", TmpRs!WSP_NAME)
         .TextMatrix(i, 2) = IIf(IsNull(TmpRs!WSP_EMAIL), "", TmpRs!WSP_EMAIL)
         .TextMatrix(i, 3) = IIf(IsNull(TmpRs!WSP_CONTACT), "", TmpRs!WSP_CONTACT)
         .TextMatrix(i, 4) = IIf(IsNull(TmpRs!WSP_PHONE), "", TmpRs!WSP_PHONE)
         .TextMatrix(i, 5) = IIf(IsNull(TmpRs!WSP_PHONE_2), "", TmpRs!WSP_PHONE_2)
         .TextMatrix(i, 6) = IIf(IsNull(TmpRs!WSP_FAX_NO), "", TmpRs!WSP_FAX_NO)
         .TextMatrix(i, 7) = IIf(IsNull(TmpRs!WSP_ADDRESS), "", TmpRs!WSP_ADDRESS)
         .TextMatrix(i, 8) = IIf(IsNull(TmpRs!WSP_DOCUMENT_TYPE), "", TmpRs!WSP_DOCUMENT_TYPE)
         .TextMatrix(i, 9) = IIf(IsNull(TmpRs!WSP_DOCUMENT_NUMBER), "", TmpRs!WSP_DOCUMENT_NUMBER)
         .TextMatrix(i, 10) = IIf(IsNull(TmpRs!WSP_PASSWORD), "", TmpRs!WSP_PASSWORD)
         .TextMatrix(i, 11) = IIf(IsNull(TmpRs!Password), "", TmpRs!Password)
         .TextMatrix(i, 12) = IIf(IsNull(TmpRs!FIRST_LOGIN_FLAG), "", TmpRs!FIRST_LOGIN_FLAG)
         .TextMatrix(i, 13) = IIf(IsNull(TmpRs!CREATED_BY), "", TmpRs!CREATED_BY)
         .TextMatrix(i, 14) = IIf(IsNull(TmpRs!CREATED_ON), "", TmpRs!CREATED_ON)
         .TextMatrix(i, 15) = IIf(IsNull(TmpRs!MODIFIED_BY), "", TmpRs!MODIFIED_BY)
         .TextMatrix(i, 16) = IIf(IsNull(TmpRs!MODIFIED_ON), "", TmpRs!MODIFIED_ON)
         .TextMatrix(i, 17) = IIf(IsNull(TmpRs!WSP_NEW_EMAIL), "", TmpRs!WSP_NEW_EMAIL)
         .TextMatrix(i, 18) = IIf(IsNull(TmpRs!ORG_CATEGORY), "", TmpRs!ORG_CATEGORY)
         .TextMatrix(i, 19) = IIf(IsNull(TmpRs!ORG_CONSTITUTION), "", TmpRs!ORG_CONSTITUTION)
         .TextMatrix(i, 20) = IIf(IsNull(TmpRs!DATE_OF_BIRTH), "", TmpRs!DATE_OF_BIRTH)
         .TextMatrix(i, 21) = IIf(IsNull(TmpRs!PAN_NO), "", TmpRs!PAN_NO)
         .TextMatrix(i, 22) = IIf(IsNull(TmpRs!TAN_NO), "", TmpRs!TAN_NO)
         .TextMatrix(i, 23) = IIf(IsNull(TmpRs!MEMBERSHIP_TYPE), "", TmpRs!MEMBERSHIP_TYPE)
         .TextMatrix(i, 24) = IIf(IsNull(TmpRs!RO_ADDRESS), "", TmpRs!RO_ADDRESS)
         .TextMatrix(i, 25) = IIf(IsNull(TmpRs!RO_TEL), "", TmpRs!RO_TEL)
         .TextMatrix(i, 26) = IIf(IsNull(TmpRs!RO_CITY), "", TmpRs!RO_CITY)
         .TextMatrix(i, 27) = IIf(IsNull(TmpRs!RO_FAX), "", TmpRs!RO_FAX)
         .TextMatrix(i, 28) = IIf(IsNull(TmpRs!RO_PINCODE), "", TmpRs!RO_PINCODE)
         .TextMatrix(i, 29) = IIf(IsNull(TmpRs!RO_ALTERNATE_TEL), "", TmpRs!RO_ALTERNATE_TEL)
         .TextMatrix(i, 30) = IIf(IsNull(TmpRs!RO_EMAIL_ID), "", TmpRs!RO_EMAIL_ID)
         .TextMatrix(i, 31) = IIf(IsNull(TmpRs!CA_CITY), "", TmpRs!CA_CITY)
         .TextMatrix(i, 32) = IIf(IsNull(TmpRs!CA_PINCODE), "", TmpRs!CA_PINCODE)
         .TextMatrix(i, 33) = IIf(IsNull(TmpRs!CA_ALTERNATE_TEL), "", TmpRs!CA_ALTERNATE_TEL)
         .TextMatrix(i, 34) = IIf(IsNull(TmpRs!CA_TELEX), "", TmpRs!CA_TELEX)
         .TextMatrix(i, 35) = IIf(IsNull(TmpRs!CA_CONTR_PERSON_NAME), "", TmpRs!CA_CONTR_PERSON_NAME)
         .TextMatrix(i, 36) = IIf(IsNull(TmpRs!CA_DESIGNATION), "", TmpRs!CA_DESIGNATION)
         .TextMatrix(i, 37) = IIf(IsNull(TmpRs!AS_NAME), "", TmpRs!AS_NAME)
         .TextMatrix(i, 38) = IIf(IsNull(TmpRs!AS_TEL), "", TmpRs!AS_TEL)
         .TextMatrix(i, 39) = IIf(IsNull(TmpRs!AS_DESIGNATION), "", TmpRs!AS_DESIGNATION)
         .TextMatrix(i, 40) = IIf(IsNull(TmpRs!AS_ALTERNATE_TEL), "", TmpRs!AS_ALTERNATE_TEL)
         .TextMatrix(i, 41) = IIf(IsNull(TmpRs!AS_ADDRESS), "", TmpRs!AS_ADDRESS)
         .TextMatrix(i, 42) = IIf(IsNull(TmpRs!AS_FAX), "", TmpRs!AS_FAX)
         .TextMatrix(i, 43) = IIf(IsNull(TmpRs!AS_CITY), "", TmpRs!AS_CITY)
         .TextMatrix(i, 44) = IIf(IsNull(TmpRs!AS_EMAIL_ID), "", TmpRs!AS_EMAIL_ID)
         .TextMatrix(i, 45) = IIf(IsNull(TmpRs!AS_PINCODE), "", TmpRs!AS_PINCODE)
         .TextMatrix(i, 46) = IIf(IsNull(TmpRs!AS_PAN_NAME), "", TmpRs!AS_PAN_NAME)
         .TextMatrix(i, 47) = IIf(IsNull(TmpRs!AS_MOBILE_NO), "", TmpRs!AS_MOBILE_NO)
         .TextMatrix(i, 48) = IIf(IsNull(TmpRs!BD_BANK_CODE), "", TmpRs!BD_BANK_CODE)
         .TextMatrix(i, 49) = IIf(IsNull(TmpRs!BD_BANK_ACCOUNT_NO), "", TmpRs!BD_BANK_ACCOUNT_NO)
         .TextMatrix(i, 50) = IIf(IsNull(TmpRs!BD_IFSC_CODE), "", TmpRs!BD_IFSC_CODE)
         .TextMatrix(i, 51) = IIf(IsNull(TmpRs!BD_BRANCH_CODE), "", TmpRs!BD_BRANCH_CODE)
         .TextMatrix(i, 52) = IIf(IsNull(TmpRs!PD_AMOUNT), "", TmpRs!PD_AMOUNT)
         .TextMatrix(i, 53) = IIf(IsNull(TmpRs!PD_INSTRUMENT_TYPE), "", TmpRs!PD_INSTRUMENT_TYPE)
         .TextMatrix(i, 54) = IIf(IsNull(TmpRs!PD_CLEARING_BANK), "", TmpRs!PD_CLEARING_BANK)
         .TextMatrix(i, 55) = IIf(IsNull(TmpRs!PD_APPLICANTS_BANK), "", TmpRs!PD_APPLICANTS_BANK)
         .TextMatrix(i, 56) = IIf(IsNull(TmpRs!PD_CL_BANK_BRANCH), "", TmpRs!PD_CL_BANK_BRANCH)
         .TextMatrix(i, 57) = IIf(IsNull(TmpRs!PD_APP_BANK_BRANCH), "", TmpRs!PD_APP_BANK_BRANCH)
         .TextMatrix(i, 58) = IIf(IsNull(TmpRs!PD_INSTRUMENT_NO), "", TmpRs!PD_INSTRUMENT_NO)
         .TextMatrix(i, 59) = IIf(IsNull(TmpRs!PD_IFSC_CODE), "", TmpRs!PD_IFSC_CODE)
         .TextMatrix(i, 60) = IIf(IsNull(TmpRs!MEMBER_SHIP_ID), "", TmpRs!MEMBER_SHIP_ID)
         .TextMatrix(i, 61) = IIf(IsNull(TmpRs!IS_FREEZED), "", TmpRs!IS_FREEZED)
         .TextMatrix(i, 62) = IIf(IsNull(TmpRs!IS_ACTIVE), "", TmpRs!IS_ACTIVE)
         .TextMatrix(i, 63) = IIf(IsNull(TmpRs!INSURANCE_AMOUNT), "", TmpRs!INSURANCE_AMOUNT)
         .TextMatrix(i, 64) = IIf(IsNull(TmpRs!INSURANCE_POLICY), "", TmpRs!INSURANCE_POLICY)
         .TextMatrix(i, 65) = IIf(IsNull(TmpRs!INSURANCE_VALIDITY), "", TmpRs!INSURANCE_VALIDITY)
         ProgressBar1.Value = ProgressBar1.Value + 1
         TmpRs.MoveNext
      Next
   End With
End If
End Sub

Private Sub Get_COMMODITY_MSTR()
tmp = " Select upper(Commodity) COMMODITY_CODE, '' COMMODITY_GRP, upper(Commodity) COMMODITY_NAME, '' COMMODITY_MKT_UNIT, '' COMMODITY_LANG_CODE"
tmp = tmp & " , '' COMMODITY_NAME_LANG_LOCAL, '' MAX_EXPIRY_DATE, '' COMMODITY_WH_UOM, '' COMMODITY_CMSE_UOM, '' PRECISION, '' MAX_VALIDITY_DATE"
tmp = tmp & " , '' ASSAYING_REQUIRED, '' STANDARD_DEDUCTION, '' COMMODITY_CAT, '' WH_CODE, '' PRECISON, '' AGGREGATION_TOLERANCE, '' COMMODITY_NAME_0"
tmp = tmp & " , '' COMMODITY_NAME_1, '' IS_FREEZED, 'A' IS_ACTIVE, '' NO_OF_DAYS, '' TOLERANCE, '' DEPOSIT_TYPE, '' STANDARD_DEDUCTION_FOR_REVALID"
tmp = tmp & " , '' MAX_SAMPLE_QTY, '' BILL_BASIS, '' HZRD_TYPE, '' RATE_UNIT, '' SERVICE_TAX, '' IS_FUNGIBLE, '' EXP_MONTHS, '' VAL_MONTHS_FRM_DEP"
tmp = tmp & " , '' NO_OF_REVAL, '' FIRST_REVAL_MONTHS, '' SECOND_REVAL_MONTHS, '' THIRD_REVAL_MONTHS, '' DENOMINATOR, '' NUMERATOR"
tmp = tmp & " , '' COMMODITY_HAIR_CUT_PERCENTAGE from mst_commodity"
Call TmpTable
If TmpRs.RecordCount > 0 Then
   ProgressBar1.Value = 0
   ProgressBar1.Max = TmpRs.RecordCount + 3
   With MSHFlexGrid8
      .Rows = .Rows + TmpRs.RecordCount
      For i = 1 To TmpRs.RecordCount
         .TextMatrix(i, 0) = IIf(IsNull(TmpRs!COMMODITY_CODE), "", TmpRs!COMMODITY_CODE)
         .TextMatrix(i, 1) = IIf(IsNull(TmpRs!COMMODITY_GRP), "", TmpRs!COMMODITY_GRP)
         .TextMatrix(i, 2) = IIf(IsNull(TmpRs!COMMODITY_NAME), "", TmpRs!COMMODITY_NAME)
         .TextMatrix(i, 3) = IIf(IsNull(TmpRs!COMMODITY_MKT_UNIT), "", TmpRs!COMMODITY_MKT_UNIT)
         .TextMatrix(i, 4) = IIf(IsNull(TmpRs!COMMODITY_LANG_CODE), "", TmpRs!COMMODITY_LANG_CODE)
         .TextMatrix(i, 5) = IIf(IsNull(TmpRs!COMMODITY_NAME_LANG_LOCAL), "", TmpRs!COMMODITY_NAME_LANG_LOCAL)
         .TextMatrix(i, 6) = IIf(IsNull(TmpRs!MAX_EXPIRY_DATE), "", TmpRs!MAX_EXPIRY_DATE)
         .TextMatrix(i, 7) = IIf(IsNull(TmpRs!COMMODITY_WH_UOM), "", TmpRs!COMMODITY_WH_UOM)
         .TextMatrix(i, 8) = IIf(IsNull(TmpRs!COMMODITY_CMSE_UOM), "", TmpRs!COMMODITY_CMSE_UOM)
         .TextMatrix(i, 9) = IIf(IsNull(TmpRs!Precision), "", TmpRs!Precision)
         .TextMatrix(i, 10) = IIf(IsNull(TmpRs!MAX_VALIDITY_DATE), "", TmpRs!MAX_VALIDITY_DATE)
         .TextMatrix(i, 11) = IIf(IsNull(TmpRs!ASSAYING_REQUIRED), "", TmpRs!ASSAYING_REQUIRED)
         .TextMatrix(i, 12) = IIf(IsNull(TmpRs!STANDARD_DEDUCTION), "", TmpRs!STANDARD_DEDUCTION)
         .TextMatrix(i, 13) = IIf(IsNull(TmpRs!COMMODITY_CAT), "", TmpRs!COMMODITY_CAT)
         .TextMatrix(i, 14) = IIf(IsNull(TmpRs!WH_CODE), "", TmpRs!WH_CODE)
         .TextMatrix(i, 15) = IIf(IsNull(TmpRs!PRECISON), "", TmpRs!PRECISON)
         .TextMatrix(i, 16) = IIf(IsNull(TmpRs!AGGREGATION_TOLERANCE), "", TmpRs!AGGREGATION_TOLERANCE)
         .TextMatrix(i, 17) = IIf(IsNull(TmpRs!COMMODITY_NAME_0), "", TmpRs!COMMODITY_NAME_0)
         .TextMatrix(i, 18) = IIf(IsNull(TmpRs!COMMODITY_NAME_1), "", TmpRs!COMMODITY_NAME_1)
         .TextMatrix(i, 19) = IIf(IsNull(TmpRs!IS_FREEZED), "", TmpRs!IS_FREEZED)
         .TextMatrix(i, 20) = IIf(IsNull(TmpRs!IS_ACTIVE), "", TmpRs!IS_ACTIVE)
         .TextMatrix(i, 21) = IIf(IsNull(TmpRs!NO_OF_DAYS), "", TmpRs!NO_OF_DAYS)
         .TextMatrix(i, 22) = IIf(IsNull(TmpRs!TOLERANCE), "", TmpRs!TOLERANCE)
         .TextMatrix(i, 23) = IIf(IsNull(TmpRs!DEPOSIT_TYPE), "", TmpRs!DEPOSIT_TYPE)
         .TextMatrix(i, 24) = IIf(IsNull(TmpRs!STANDARD_DEDUCTION_FOR_REVALID), "", TmpRs!STANDARD_DEDUCTION_FOR_REVALID)
         .TextMatrix(i, 25) = IIf(IsNull(TmpRs!MAX_SAMPLE_QTY), "", TmpRs!MAX_SAMPLE_QTY)
         .TextMatrix(i, 26) = IIf(IsNull(TmpRs!BILL_BASIS), "", TmpRs!BILL_BASIS)
         .TextMatrix(i, 27) = IIf(IsNull(TmpRs!HZRD_TYPE), "", TmpRs!HZRD_TYPE)
         .TextMatrix(i, 28) = IIf(IsNull(TmpRs!RATE_UNIT), "", TmpRs!RATE_UNIT)
         .TextMatrix(i, 29) = IIf(IsNull(TmpRs!SERVICE_TAX), "", TmpRs!SERVICE_TAX)
         .TextMatrix(i, 30) = IIf(IsNull(TmpRs!IS_FUNGIBLE), "", TmpRs!IS_FUNGIBLE)
         .TextMatrix(i, 31) = IIf(IsNull(TmpRs!EXP_MONTHS), "", TmpRs!EXP_MONTHS)
         .TextMatrix(i, 32) = IIf(IsNull(TmpRs!VAL_MONTHS_FRM_DEP), "", TmpRs!VAL_MONTHS_FRM_DEP)
         .TextMatrix(i, 33) = IIf(IsNull(TmpRs!NO_OF_REVAL), "", TmpRs!NO_OF_REVAL)
         .TextMatrix(i, 34) = IIf(IsNull(TmpRs!FIRST_REVAL_MONTHS), "", TmpRs!FIRST_REVAL_MONTHS)
         .TextMatrix(i, 35) = IIf(IsNull(TmpRs!SECOND_REVAL_MONTHS), "", TmpRs!SECOND_REVAL_MONTHS)
         .TextMatrix(i, 36) = IIf(IsNull(TmpRs!THIRD_REVAL_MONTHS), "", TmpRs!THIRD_REVAL_MONTHS)
         .TextMatrix(i, 37) = IIf(IsNull(TmpRs!DENOMINATOR), "", TmpRs!DENOMINATOR)
         .TextMatrix(i, 38) = IIf(IsNull(TmpRs!NUMERATOR), "", TmpRs!NUMERATOR)
         .TextMatrix(i, 39) = IIf(IsNull(TmpRs!COMMODITY_HAIR_CUT_PERCENTAGE), "", TmpRs!COMMODITY_HAIR_CUT_PERCENTAGE)
         TmpRs.MoveNext
         ProgressBar1.Value = ProgressBar1.Value + 1
      Next
   End With
End If
End Sub

Private Sub Get_WH_COMMODITY_MAP()
tmp = " select upper(cmp.cmpname)+'_'+upper(mg.godownno) WH_CODE,  '' WSP_CODE, upper(com.Commodity) COMMODITY_CODE, '' WH_COM_CODE,"
tmp = tmp & " '' IS_FREEZED, 'A' IS_ACTIVE from Mst_LocationCommodityMapping mlcm"
tmp = tmp & " inner join mst_cmp cmp on cmp.locationid = mlcm.locationid"
tmp = tmp & " inner join mst_godown mg on mg.cmpid = cmp.cmpid"
tmp = tmp & " inner join mst_commodity  com on com.commodityid = mlcm.commodityid"
Call TmpTable
If TmpRs.RecordCount > 0 Then
   ProgressBar1.Value = 0
   ProgressBar1.Max = TmpRs.RecordCount + 3
   With MSHFlexGrid9
      .Rows = .Rows + TmpRs.RecordCount
      For i = 1 To TmpRs.RecordCount
         .TextMatrix(i, 0) = IIf(IsNull(TmpRs!WH_CODE), "", TmpRs!WH_CODE)
         .TextMatrix(i, 1) = IIf(IsNull(TmpRs!WSP_CODE), "", TmpRs!WSP_CODE)
         .TextMatrix(i, 2) = IIf(IsNull(TmpRs!COMMODITY_CODE), "", TmpRs!COMMODITY_CODE)
         .TextMatrix(i, 3) = IIf(IsNull(TmpRs!WH_COM_CODE), "", TmpRs!WH_COM_CODE)
         .TextMatrix(i, 4) = IIf(IsNull(TmpRs!IS_FREEZED), "", TmpRs!IS_FREEZED)
         .TextMatrix(i, 5) = IIf(IsNull(TmpRs!IS_ACTIVE), "", TmpRs!IS_ACTIVE)
         ProgressBar1.Value = ProgressBar1.Value + 1
         TmpRs.MoveNext
      Next
   End With
End If
End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   MsgBox "Access Denied !!!!!!!!!"
   Frame1.Enabled = False
   Frame2.Enabled = False
   Exit Sub
End If
Call CmdClear_Click

tmp = "select distinct commtrackbatchid As BatchID from Txn_CxTF_GSL Where commtrackbatchid > 0  order by commtrackbatchid"
Call Tmp1Table
CmbBatch.Clear
If TmpRs1.RecordCount > 0 Then
   For i = 1 To TmpRs1.RecordCount
      CmbBatch.AddItem TmpRs1.Fields(0)
      mBatchID = TmpRs1.Fields(0) + 1
      TmpRs1.MoveNext
   Next
Else
   mBatchID = 1
End If

If Gen_UserLoginID = 2047 Or Gen_UserLoginID = 1076 Then
   Check1.Visible = True
   Check1.Value = vbChecked
   'Check1.Enabled = False
Else
   Check1.Visible = False
   Check1.Value = vbUnchecked
End If

End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_CommTrackData.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_CommTrackData.xls")

Set oWS = oWB.Worksheets("Sheet1")
oWS.Name = "DEPOSIT DATA"
oWS.ClearArrows
mRow = 1
ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid10.Rows
For i = 0 To MSHFlexGrid10.Rows - 1
   mRow = i + 1
   For c = 0 To MSHFlexGrid10.Cols - 1
      oWS.Cells(mRow, c + 1) = MSHFlexGrid10.TextMatrix(i, c)
   Next
   ProgressBar1.Value = ProgressBar1.Value + 1
Next
Set oWS = oWB.Worksheets("Sheet2")
oWS.Name = "WITHDRAWAL DATA"
oWS.ClearArrows
mRow = 1
ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid11.Rows
For i = 0 To MSHFlexGrid11.Rows - 1
   mRow = i + 1
   For c = 0 To MSHFlexGrid11.Cols - 1
      oWS.Cells(mRow, c + 1) = MSHFlexGrid11.TextMatrix(i, c)
   Next
   ProgressBar1.Value = ProgressBar1.Value + 1
Next



''--------Commented if not all data required
'Set oWS = oWB.Worksheets("Sheet1")
'oWS.Name = "UOM_MSTR"
'oWS.ClearArrows
'mRow = 1
'For I = 0 To MSHFlexGrid1.Rows - 1
'   mRow = I + 1
'   For c = 0 To MSHFlexGrid1.Cols - 1
'      oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(I, c)
'   Next
'Next
'
'Set oWS = oWB.Worksheets("Sheet2")
'oWS.Name = "CLIENT_MSTR"
'oWS.ClearArrows
'mRow = 1
'For I = 0 To MSHFlexGrid2.Rows - 1
'   mRow = I + 1
'   For c = 0 To MSHFlexGrid2.Cols - 1
'      oWS.Cells(mRow, c + 1) = MSHFlexGrid2.TextMatrix(I, c)
'   Next
'Next
'
'Set oWS = oWB.Worksheets("Sheet3")
'oWS.Name = "USER_MSTR"
'oWS.ClearArrows
'mRow = 1
'For I = 0 To MSHFlexGrid3.Rows - 1
'   mRow = I + 1
'   For c = 0 To MSHFlexGrid3.Cols - 1
'      oWS.Cells(mRow, c + 1) = MSHFlexGrid3.TextMatrix(I, c)
'   Next
'Next
'
'Set oWS = oWB.Worksheets("Sheet4")
'oWS.Name = "ASS_MSTR"
'oWS.ClearArrows
'mRow = 1
'For I = 0 To MSHFlexGrid4.Rows - 1
'   mRow = I + 1
'   For c = 0 To MSHFlexGrid4.Cols - 1
'      oWS.Cells(mRow, c + 1) = MSHFlexGrid4.TextMatrix(I, c)
'   Next
'Next
'
'Set oWS = oWB.Worksheets("Sheet5")
'oWS.Name = "LOCATION_MSTR"
'oWS.ClearArrows
'mRow = 1
'For I = 0 To MSHFlexGrid5.Rows - 1
'   mRow = I + 1
'   For c = 0 To MSHFlexGrid5.Cols - 1
'      oWS.Cells(mRow, c + 1) = MSHFlexGrid5.TextMatrix(I, c)
'   Next
'Next
'
'Set oWS = oWB.Worksheets("Sheet6")
'oWS.Name = "WH_MSTR"
'oWS.ClearArrows
'mRow = 1
'For I = 0 To MSHFlexGrid6.Rows - 1
'   mRow = I + 1
'   For c = 0 To MSHFlexGrid6.Cols - 1
'      oWS.Cells(mRow, c + 1) = MSHFlexGrid6.TextMatrix(I, c)
'   Next
'Next
'
'Set oWS = oWB.Worksheets("Sheet7")
'oWS.Name = "WSP_MSTR"
'oWS.ClearArrows
'mRow = 1
'For I = 0 To MSHFlexGrid7.Rows - 1
'   mRow = I + 1
'   For c = 0 To MSHFlexGrid7.Cols - 1
'      oWS.Cells(mRow, c + 1) = MSHFlexGrid7.TextMatrix(I, c)
'   Next
'Next
'
'Set oWS = oWB.Worksheets("Sheet8")
'oWS.Name = "COMMODITY_MSTR"
'oWS.ClearArrows
'mRow = 1
'For I = 0 To MSHFlexGrid8.Rows - 1
'   mRow = I + 1
'   For c = 0 To MSHFlexGrid8.Cols - 1
'      oWS.Cells(mRow, c + 1) = MSHFlexGrid8.TextMatrix(I, c)
'   Next
'Next
'
'Set oWS = oWB.Worksheets("Sheet9")
'oWS.Name = "WH_COMMODITY_MAP"
'oWS.ClearArrows
'mRow = 1
'For I = 0 To MSHFlexGrid9.Rows - 1
'   mRow = I + 1
'   For c = 0 To MSHFlexGrid9.Cols - 1
'      oWS.Cells(mRow, c + 1) = MSHFlexGrid9.TextMatrix(I, c)
'   Next
'Next
'
'Set oWS = oWB.Worksheets("Sheet10")
'oWS.Name = "DEPOSIT DATA"
'oWS.ClearArrows
'mRow = 1
'For I = 0 To MSHFlexGrid10.Rows - 1
'   mRow = I + 1
'   For c = 0 To MSHFlexGrid10.Cols - 1
'      oWS.Cells(mRow, c + 1) = MSHFlexGrid10.TextMatrix(I, c)
'   Next
'Next
'
'Set oWS = oWB.Worksheets("Sheet11")
'oWS.Name = "WITHDRAWAL DATA"
'oWS.ClearArrows
'mRow = 1
'For I = 0 To MSHFlexGrid11.Rows - 1
'   mRow = I + 1
'   For c = 0 To MSHFlexGrid11.Cols - 1
'      oWS.Cells(mRow, c + 1) = MSHFlexGrid11.TextMatrix(I, c)
'   Next
'Next



oWB.Save
MsgBox ("Excel File Created!!!")
oXL.Visible = True
End Sub

Private Sub OptOld_Click()
Call OptPending_Click
End Sub

Private Sub OptPending_Click()
Call Grid_Head1
Call Grid_Head2
Call Grid_Head3
Call Grid_Head4
Call Grid_Head5
Call Grid_Head6
Call Grid_Head7
Call Grid_Head8
Call Grid_Head9
Call Grid_Head10
Call Grid_Head11
If OptPending.Value = True Then
   CmbBatch.Text = ""
   CmbBatch.Enabled = False
   TxtFromDate.Enabled = True
   TxtToDate.Enabled = True
   Check2.Value = vbUnchecked
   Check2.Enabled = True
ElseIf OptOld.Value = True Then
   CmbBatch.Enabled = True
   TxtFromDate.Enabled = False
   TxtToDate.Enabled = False
   Check2.Value = vbUnchecked
   Check2.Enabled = False
End If
End Sub
