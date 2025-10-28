VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmTxn_UpdateLessorAgreementFlag 
   Caption         =   "Update Lessor Agreement Flag"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   2835
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   15015
      Begin VB.TextBox TxtServiceTaxID 
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
         Left            =   10320
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   240
         Width           =   4575
      End
      Begin VB.TextBox TxtStatus 
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
         Left            =   1905
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1665
      End
      Begin VB.TextBox TxtAgreeRecAs 
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
         Left            =   10320
         Locked          =   -1  'True
         MaxLength       =   200
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Top             =   720
         Width           =   4575
      End
      Begin VB.TextBox TxtAgreementID 
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
         Left            =   1890
         MaxLength       =   10
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   240
         Width           =   1665
      End
      Begin VB.TextBox TxtRemark 
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
         Height          =   615
         Left            =   5640
         Locked          =   -1  'True
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   1680
         Width           =   9255
      End
      Begin VB.TextBox TxtAdvAdjusted 
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
         Height          =   375
         Left            =   5640
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox TxtAdvanceRemain 
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
         Height          =   375
         Left            =   10320
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox TxtAdvanceGiven 
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
         Height          =   375
         Left            =   1890
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   9
         Top             =   1200
         Width           =   1680
      End
      Begin VB.TextBox TxtBalance 
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
         Height          =   375
         Left            =   13200
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CheckBox ChkAgreementNotReceived 
         Caption         =   "Agreement not Received"
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
         Left            =   10335
         TabIndex        =   2
         Top             =   2400
         Width           =   2760
      End
      Begin VB.CheckBox ChkBRAgreement 
         Caption         =   "BR Agreement"
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
         Height          =   240
         Left            =   5640
         TabIndex        =   7
         Top             =   2400
         Width           =   2640
      End
      Begin MSComCtl2.DTPicker TxtEntryDate 
         Height          =   375
         Left            =   5640
         TabIndex        =   13
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   111411201
         CurrentDate     =   39884
      End
      Begin MSComCtl2.DTPicker TxtStartDate 
         Height          =   375
         Left            =   1890
         TabIndex        =   14
         Top             =   720
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   111411201
         CurrentDate     =   39884
      End
      Begin MSComCtl2.DTPicker TxtExpiryDate 
         Height          =   375
         Left            =   5640
         TabIndex        =   15
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   111411201
         CurrentDate     =   39884
      End
      Begin MSComCtl2.DTPicker TxtAAggrStartDate 
         Height          =   375
         Left            =   1875
         TabIndex        =   19
         Top             =   2333
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   111411201
         CurrentDate     =   39884
      End
      Begin VB.Label LblAgreemID 
         AutoSize        =   -1  'True
         Caption         =   "Agreement ID"
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
         Left            =   135
         TabIndex        =   32
         Top             =   300
         Width           =   1230
      End
      Begin VB.Label LblEntryDate 
         AutoSize        =   -1  'True
         Caption         =   "Entry Date"
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
         Left            =   3960
         TabIndex        =   31
         Top             =   300
         Width           =   930
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   3960
         TabIndex        =   30
         Top             =   1867
         Width           =   825
      End
      Begin VB.Label LblAdvGiven 
         Caption         =   "Security Deposit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   135
         TabIndex        =   29
         Top             =   1102
         Width           =   1680
      End
      Begin VB.Label LblAdvAdjust 
         AutoSize        =   -1  'True
         Caption         =   "SD Adjusted"
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
         Left            =   3960
         TabIndex        =   28
         Top             =   1267
         Width           =   1125
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Start Date"
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
         Left            =   135
         TabIndex        =   27
         Top             =   787
         Width           =   885
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Expiry Date"
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
         Left            =   3960
         TabIndex        =   26
         Top             =   787
         Width           =   1035
      End
      Begin VB.Label LblAgrePecAs 
         AutoSize        =   -1  'True
         Caption         =   "Agreement Received As"
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
         Left            =   7440
         TabIndex        =   25
         Top             =   787
         Width           =   2535
      End
      Begin VB.Label LblAdvRem 
         AutoSize        =   -1  'True
         Caption         =   "SD Returned"
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
         Left            =   7440
         TabIndex        =   24
         Top             =   1267
         Width           =   1155
      End
      Begin VB.Label LblBalance 
         AutoSize        =   -1  'True
         Caption         =   "Balance"
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
         Left            =   12180
         TabIndex        =   23
         Top             =   1230
         Width           =   750
      End
      Begin VB.Label Label15 
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
         Height          =   255
         Left            =   135
         TabIndex        =   22
         Top             =   1733
         Width           =   960
      End
      Begin VB.Label Label16 
         Caption         =   "Service Tax Description"
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
         Left            =   7440
         TabIndex        =   21
         Top             =   300
         Width           =   2700
      End
      Begin VB.Label Label26 
         Caption         =   "Actual Agreement Start Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   120
         TabIndex        =   20
         Top             =   2258
         Width           =   1665
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   720
      Left            =   120
      TabIndex        =   1
      Top             =   7080
      Width           =   15015
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
         Height          =   405
         Left            =   11985
         TabIndex        =   5
         Top             =   195
         Width           =   2820
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   9165
         TabIndex        =   4
         Top             =   195
         Width           =   2820
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4140
      Left            =   120
      TabIndex        =   3
      Top             =   2880
      Width           =   14985
      _ExtentX        =   26432
      _ExtentY        =   7303
      _Version        =   393216
      Tabs            =   4
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
      TabCaption(0)   =   "Lessor Details"
      TabPicture(0)   =   "FrmTxn_UpdateLessorAgreementFlag.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "MSHFlexGrid2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Godown Details"
      TabPicture(1)   =   "FrmTxn_UpdateLessorAgreementFlag.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "MSHFlexGrid1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Agreement Attachment Details"
      TabPicture(2)   =   "FrmTxn_UpdateLessorAgreementFlag.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "MSHFlexGrid5"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Log Details"
      TabPicture(3)   =   "FrmTxn_UpdateLessorAgreementFlag.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label20"
      Tab(3).Control(1)=   "Label21"
      Tab(3).Control(2)=   "MSHFlexGrid4"
      Tab(3).Control(3)=   "TxtUpdated"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "TxtCreated"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).ControlCount=   5
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
         Height          =   360
         Left            =   -73830
         Locked          =   -1  'True
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   405
         Width           =   6210
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
         Height          =   360
         Left            =   -66150
         Locked          =   -1  'True
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   405
         Width           =   5970
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
         Height          =   3615
         Left            =   120
         TabIndex        =   35
         Top             =   360
         Width           =   14700
         _ExtentX        =   25929
         _ExtentY        =   6376
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   3720
         Left            =   -74880
         TabIndex        =   36
         Top             =   360
         Width           =   14715
         _ExtentX        =   25956
         _ExtentY        =   6562
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid4 
         Height          =   3135
         Left            =   -74880
         TabIndex        =   37
         Top             =   840
         Width           =   14700
         _ExtentX        =   25929
         _ExtentY        =   5530
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid5 
         Height          =   3630
         Left            =   -74880
         TabIndex        =   38
         Top             =   390
         Width           =   14745
         _ExtentX        =   26009
         _ExtentY        =   6403
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
      Begin VB.Label Label21 
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
         Left            =   -74880
         TabIndex        =   40
         Top             =   465
         Width           =   855
      End
      Begin VB.Label Label20 
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
         Left            =   -67320
         TabIndex        =   39
         Top             =   465
         Width           =   885
      End
   End
   Begin MSComDlg.CommonDialog cdialog 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmTxn_UpdateLessorAgreementFlag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mLocationID, mCMPID, mSCMPID, mLessorID, mGdLessorID, mSMPrefix, Edit_Flg As Variant
Dim mGodownID, mStatus, mServiceTaxID, mSStatus As Variant
Dim mExpiryDate, mSGodownID, mSLessorID, mPrevExpiryDate As Variant
Dim File_Path As String
Dim mMsg As Variant
Dim mAmount, mGodownRentAmount As Double
Dim mGodownExists As Boolean
Dim mGodownMsg As Variant

Private Sub ExitCmd_Click()
 Unload Me
End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
Dim wc As Variant

FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Frame1.Enabled = False
   ButtonFrm.Enabled = False
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If

Frame1.Enabled = True
ButtonFrm.Enabled = True
Call Grid_Head1 ''-- Grid 2
Call Grid_Head2 ''-- Grid 1
Call Grid_Head4 ''-- Grid 5
Call Grid_Head3 ''-- Grid 4

TxtAgreementID.Text = ""
TxtEntryDate.Value = Date
TxtServiceTaxID.Text = ""
TxtStartDate.Value = Date
TxtExpiryDate.Value = Date
TxtAgreeRecAs.Text = ""
TxtAdvanceGiven.Text = ""
TxtAdvAdjusted.Text = ""
TxtAdvanceRemain.Text = ""
TxtBalance.Text = ""
TxtStatus.Text = ""
TxtRemark.Text = ""
TxtAAggrStartDate.Value = Date
ChkBRAgreement.Value = vbUnchecked: ChkBRAgreement.Enabled = False
ChkAgreementNotReceived.Value = vbUnchecked
TxtCreated = ""
TxtUpdated = ""
SaveCmd.Enabled = False
ExitCmd.Enabled = True
SSTab1.Tab = 0
End Sub

Private Sub Grid_Head1()
With MSHFlexGrid2
    .Clear
    .Rows = 3
    .FixedRows = 2
    .Cols = 4

    .TextMatrix(1, 0) = "Lessor ID":        .ColWidth(0) = 0
    .TextMatrix(1, 1) = "Lessor Name":      .ColWidth(1) = 3400
    .TextMatrix(1, 2) = "Amount":           .ColWidth(2) = 1500
    .TextMatrix(1, 3) = "Remark":           .ColWidth(3) = 2800
End With
End Sub

Private Sub Grid_Head2()
With MSHFlexGrid1
    .Clear
    .Rows = 3
    .FixedRows = 2
    .Cols = 11
    .TextMatrix(1, 0) = "Godown ID":            .ColWidth(0) = 0
    .TextMatrix(1, 1) = "Location":             .ColWidth(1) = 2200
    .TextMatrix(1, 2) = "CMP Name":             .ColWidth(2) = 2200
    .TextMatrix(1, 3) = "Godown No":            .ColWidth(3) = 1300
    .TextMatrix(1, 4) = "Address":              .ColWidth(4) = 3000
    .TextMatrix(1, 5) = "Area in Sq.Ft":        .ColWidth(5) = 1300
    .TextMatrix(1, 6) = "Lessor ID":            .ColWidth(6) = 0
    .TextMatrix(1, 7) = "Lessor Name":          .ColWidth(7) = 2600
    .TextMatrix(1, 8) = "Rent Amount":          .ColWidth(8) = 1600
    .TextMatrix(1, 9) = "SM_Prefix":            .ColWidth(9) = 0
    .TextMatrix(1, 10) = "RentPaidUpTo":        .ColWidth(10) = 1200
End With
End Sub

Private Sub Grid_Head4()
With MSHFlexGrid5
    .Clear
    .Rows = 2
    .Cols = 5
    .Font.Size = 10
    .WordWrap = True
    .TextMatrix(0, 0) = "SR No":                    .ColWidth(0) = 1000
    .TextMatrix(0, 1) = "Attachment File Name":     .ColWidth(1) = 5000
    .TextMatrix(0, 2) = "Remarks":                  .ColWidth(2) = 8200
    .TextMatrix(0, 3) = "File Path":                .ColWidth(3) = 0
    .TextMatrix(0, 4) = "ID":                       .ColWidth(4) = 0
End With
End Sub

Private Sub Grid_Head3()
With MSHFlexGrid4
    .Clear
    .Rows = 2
    .FixedRows = 1
    .Cols = 3
    .TextMatrix(0, 0) = "Lessor Agreement ID":      .ColWidth(0) = 2200
    .TextMatrix(0, 1) = "Updated By":               .ColWidth(1) = 2000
    .TextMatrix(0, 2) = "Updated On":               .ColWidth(2) = 3000
End With
End Sub

Private Sub SaveCmd_Click()
Dim mResult As Boolean
If Val(TxtAgreementID) > 0 Then
'-- Ullhas Patil 20 April 2018
'    If ChkAgreementNotReceived.Value = vbChecked Then
'        MsgBox "If agreement is received then un-checked 'Agreement Not Received' and save."
'        ChkAgreementNotReceived.SetFocus
'        Exit Sub
'    Else
'        If RsTxn_LessorAgreement.RecordCount > 0 Then
'            If RsTxn_LessorAgreement!LAID = Val(TxtAgreementID) Then
'                RsTxn_LessorAgreement!AgreenotReceieveFlag = "0"
'                RsTxn_LessorAgreement!UpdatedDate = Now
'                RsTxn_LessorAgreement!UpdatedBy = Gen_UserLoginID
'                RsTxn_LessorAgreement.Update
'                Call Gen_Email
'                MsgBox "Done."
'                Call ClearFrame1
'                TxtAgreementID.Text = ""
'                TxtAgreementID.SetFocus
'            End If
'        End If
'    End If
    
    Select Case MsgBox(" Do you want to update " & IIf(ChkAgreementNotReceived.Value = vbChecked, " 'Agreement not Received' ", " 'Agreement Received' ") & " for Lessor Agreement ID : " & Val(TxtAgreementID), vbQuestion + vbYesNo)
        Case vbYes
            If RsTxn_LessorAgreement.RecordCount > 0 Then
                If RsTxn_LessorAgreement!LAID = Val(TxtAgreementID) Then
                    If ChkAgreementNotReceived.Value = vbChecked Then
                        RsTxn_LessorAgreement!AgreenotReceieveFlag = "1"    ''-- Agreement not Received
                    Else
                        RsTxn_LessorAgreement!AgreenotReceieveFlag = "0"    ''-- Agreement Received
                    End If
                    RsTxn_LessorAgreement!UpdatedDate = Now
                    RsTxn_LessorAgreement!UpdatedBy = Gen_UserLoginID
                    RsTxn_LessorAgreement.Update
                    Call Gen_Email
                    MsgBox "Done."
                    Call ClearFrame1
                    TxtAgreementID.Text = ""
                    TxtAgreementID.SetFocus
                End If
            End If
        Case vbNo
            Exit Sub
    End Select

End If
End Sub

Private Sub TxtAgreementID_LostFocus()
If TxtAgreementID = "" Then Call ClearFrame1: Exit Sub
If IsNumeric(TxtAgreementID) = False Or Val(TxtAgreementID) < 0 Then
   MsgBox "Invalid format!!!"
   Call ClearFrame1
   TxtAgreementID.SetFocus
   Exit Sub
End If
Call ClearFrame1
Call Indata
End Sub

Sub Indata()
If TxtAgreementID.Text <> "" Then
    Call TableTxn_LessorAgreement(" Where LAID = " & Val(TxtAgreementID) & " ")
    If RsTxn_LessorAgreement.RecordCount > 0 Then
        
        If LCase(IIf(IsNull(RsTxn_LessorAgreement!Flag), "", RsTxn_LessorAgreement!Flag)) = "p" Then
            MsgBox "Agreement is Pending"
            'TxtAgreementID.SetFocus
            Exit Sub
        ElseIf LCase(IIf(IsNull(RsTxn_LessorAgreement!Flag), "", RsTxn_LessorAgreement!Flag)) = "r" Then
            MsgBox "Agreement is Rejected"
            'TxtAgreementID.SetFocus
            Exit Sub
        End If
        
''--- Ullhas Patil 20 April 2018
'        If Val(IIf(IsNull(RsTxn_LessorAgreement!AgreenotReceieveFlag), 0, RsTxn_LessorAgreement!AgreenotReceieveFlag)) = 0 Then
'            MsgBox "Agreement already Receieved for Agreement ID : - " & TxtAgreementID & " ."
'            'TxtAgreementID.SetFocus
'            Exit Sub
'        End If
                
        TxtStartDate.Value = RsTxn_LessorAgreement!StartDate
        TxtAAggrStartDate.Value = RsTxn_LessorAgreement!AgreemenActualtStartDate
        TxtEntryDate.Value = RsTxn_LessorAgreement!EntryDate
        TxtExpiryDate.Value = RsTxn_LessorAgreement!ExpiryDate
        mServiceTaxID = IIf(IsNull(RsTxn_LessorAgreement!ServiceTaxID), 0, RsTxn_LessorAgreement!ServiceTaxID)
        mSerTax = GetServiceTaxDescription(IIf(IsNull(RsTxn_LessorAgreement!ServiceTaxID), 0, RsTxn_LessorAgreement!ServiceTaxID))
        TxtServiceTaxID.Text = IIf(mSerTax = 0, "", mSerTax)
        TxtAgreeRecAs = IIf(IsNull(RsTxn_LessorAgreement!AgreementRecvAs), "", RsTxn_LessorAgreement!AgreementRecvAs)
        TxtAdvanceGiven = Format(IIf(IsNull(RsTxn_LessorAgreement!AdvanceGiven), "", RsTxn_LessorAgreement!AdvanceGiven), "#####0.00")
        TxtAdvAdjusted = Format(IIf(IsNull(RsTxn_LessorAgreement!AdvanceAdjusted), "", RsTxn_LessorAgreement!AdvanceAdjusted), "#####0.00")
        TxtAdvanceRemain = Format(IIf(IsNull(RsTxn_LessorAgreement!AdvanceReturn), "", RsTxn_LessorAgreement!AdvanceReturn), "#####0.00")
        TxtBalance.Text = Format(IIf(IsNull(RsTxn_LessorAgreement!AdvanceBalance), "", RsTxn_LessorAgreement!AdvanceBalance), "#####0.00")
        TxtRemark = IIf(IsNull(RsTxn_LessorAgreement!Remark), "", RsTxn_LessorAgreement!Remark)
        TxtCreated = IIf(IsNull(RsTxn_LessorAgreement!CreatedBy), "", RsTxn_LessorAgreement!CreatedBy) & " :- " & IIf(IsNull(RsTxn_LessorAgreement!CreatedDate), "", RsTxn_LessorAgreement!CreatedDate)
        TxtUpdated = IIf(IsNull(RsTxn_LessorAgreement!UpdatedBy), "", RsTxn_LessorAgreement!UpdatedBy) & " :- " & IIf(IsNull(RsTxn_LessorAgreement!UpdatedDate), "", RsTxn_LessorAgreement!UpdatedDate)
        ChkAgreementNotReceived.Value = IIf(IsNull(RsTxn_LessorAgreement!AgreenotReceieveFlag), 0, RsTxn_LessorAgreement!AgreenotReceieveFlag)
        ChkBRAgreement.Value = IIf(RsTxn_LessorAgreement!BRAgreement = True, vbChecked, Unchecked)
        mStatus = IIf(IsNull(RsTxn_LessorAgreement!Flag), "", RsTxn_LessorAgreement!Flag)
        
        If LCase(mStatus) = "a" Then
           TxtStatus = "Approved"
        ElseIf LCase(mStatus) = "p" Then
           TxtStatus = "Pending"
        ElseIf LCase(mStatus) = "r" Then
           TxtStatus = "Rejected"
        Else
           TxtStatus = ""
        End If

        Call InDataLessor
        Call InDataGodown
        Call InDataLogHistory
        Call IndataAttachmentDetails
        SSTab1.Tab = 0
        SaveCmd.Enabled = True
        
        If Val(IIf(IsNull(RsTxn_LessorAgreement!AgreenotReceieveFlag), 0, RsTxn_LessorAgreement!AgreenotReceieveFlag)) = 0 Then
            MsgBox "Agreement already Receieved for Agreement ID : - " & TxtAgreementID & " ."
        Else
            MsgBox "Agreement not Receieved for Agreement ID : - " & TxtAgreementID & " ."
        End If
        
    Else
        MsgBox "Agreement Not Found."
        TxtAgreementID.SetFocus
        Exit Sub
    End If
End If
End Sub


Private Sub InDataLessor()
With MSHFlexGrid2
Call Grid_Head1
tmp = "Select TLD.LAID,TLD.LessorID,TLD.LessorAmt,TLD.Remark,ML.LessorName"
tmp = tmp & " From Txn_LessorAgreementDetail TLD"
tmp = tmp & " inner Join Mst_Lessor ML on TLD.LessorID=ML.LessorID"
tmp = tmp & " Where TLD.LAID=" & Val(TxtAgreementID) & ""
Call Tmp1Table
.Rows = TmpRs1.RecordCount + 3
For I = 2 To .Rows - 2
   .TextMatrix(I, 0) = TmpRs1!LessorID
   .TextMatrix(I, 1) = TmpRs1!LessorName
   .TextMatrix(I, 2) = Format(TmpRs1!LessorAmt, "#####0.00")
   .TextMatrix(I, 3) = TmpRs1!Remark
   .TextMatrix(0, 2) = Format(Val(.TextMatrix(0, 2)) + TmpRs1!LessorAmt, "#####0.00")
   TmpRs1.MoveNext
Next
End With
End Sub

Private Sub InDataLogHistory()
With MSHFlexGrid4
Call Grid_Head3
tmp = "Select LAID, UpdatedBy, UpdatedOn from Log_LessorAgreementHistory Where LAID=" & Val(TxtAgreementID) & ""
Call TmpTable
.Rows = .Rows + TmpRs.RecordCount
For I = 1 To .Rows - 2
   .TextMatrix(I, 0) = TmpRs!LAID
   .TextMatrix(I, 1) = IIf(IsNull(TmpRs!UpdatedBy), "", TmpRs!UpdatedBy)
   .TextMatrix(I, 2) = TmpRs!UpdatedOn
   TmpRs.MoveNext
Next
End With
End Sub

Private Sub InDataGodown()
With MSHFlexGrid1
Call Grid_Head2
tmp = "select TGD.LAID,TGD.GodownID,TGD.SM_Prefix,TGD.LessorID,TGD.RentAmount,"
tmp = tmp & " MG.GodownNo , MG.Address, MG.Area, MC.CMPName, ML.LocationName, MLes.LessorName,TGD.RentPaidUpTo"
tmp = tmp & " From Txn_GodownLessorAgreeDetail TGD"
tmp = tmp & " Inner Join Mst_Godown MG on TGD.GodownID=MG.GodownID"
tmp = tmp & " Inner Join Mst_CMP MC on MG.CMPID=MC.CMPID"
tmp = tmp & " Inner Join Mst_Location ML on MC.LocationID=ML.LocationID"
tmp = tmp & " inner Join Mst_Lessor MLes on TGD.LessorID=MLes.LessorID"
tmp = tmp & " Where TGD.LAID=" & Val(TxtAgreementID) & ""
tmp = tmp & " Order By TGD.GLADID"
Call TmpTable
.Rows = TmpRs.RecordCount + 3
For I = 2 To .Rows - 2
   .TextMatrix(I, 0) = TmpRs!GodownID
   .TextMatrix(I, 3) = TmpRs!GodownNo
   .TextMatrix(I, 5) = TmpRs!Area
   .TextMatrix(I, 4) = TmpRs!Address
   .TextMatrix(I, 2) = TmpRs!CMPName
   .TextMatrix(I, 1) = TmpRs!LocationName
   .TextMatrix(I, 6) = TmpRs!LessorID
   .TextMatrix(I, 7) = TmpRs!LessorName
   .TextMatrix(I, 8) = Format(TmpRs!RentAmount, "#####0.00")
   .TextMatrix(I, 9) = TmpRs!SM_Prefix
   .TextMatrix(I, 10) = IIf(IsNull(TmpRs!RentPaidUpTo), "", TmpRs!RentPaidUpTo)
   .TextMatrix(0, 8) = Format(Val(.TextMatrix(0, 8)) + TmpRs!RentAmount, "#####0.00")
   TmpRs.MoveNext
Next

End With
End Sub

Private Sub IndataAttachmentDetails()
With MSHFlexGrid5
Call Grid_Head4
Call TableTxn_LessorAgreeAttachment(" Where LAID = " & TxtAgreementID.Text & " ")
.Rows = RsTxn_LessorAgreeAttachment.RecordCount + 2
For I = 1 To RsTxn_LessorAgreeAttachment.RecordCount

    .TextMatrix(I, 0) = I
    .TextMatrix(I, 1) = RsTxn_LessorAgreeAttachment!FileName
    .TextMatrix(I, 2) = RsTxn_LessorAgreeAttachment!Remarks
    .TextMatrix(I, 3) = App.Path & "\Agreement Attachment\Lessor\" & MSHFlexGrid5.TextMatrix(I, 1)
    File_Path = .TextMatrix(I, 3) 'App.Path & "\Project Attachment\"
    Call LoadFileFromDB(File_Path, RsTxn_LessorAgreeAttachment, "File")
    .TextMatrix(I, 4) = RsTxn_LessorAgreeAttachment!TxnID
    RsTxn_LessorAgreeAttachment.MoveNext
Next
End With
End Sub

Function ClearFrame1()
Call Grid_Head1 ''-- Grid 2
Call Grid_Head2 ''-- Grid 1
Call Grid_Head4 ''-- Grid 5
Call Grid_Head3 ''-- Grid 4

'TxtAgreementID.Text = ""
TxtEntryDate.Value = Date
TxtServiceTaxID.Text = ""
TxtStartDate.Value = Date
TxtExpiryDate.Value = Date
TxtAgreeRecAs.Text = ""
TxtAdvanceGiven.Text = ""
TxtAdvAdjusted.Text = ""
TxtAdvanceRemain.Text = ""
TxtBalance.Text = ""
TxtStatus.Text = ""
TxtRemark.Text = ""
TxtAAggrStartDate.Value = Date
ChkBRAgreement.Value = vbUnchecked: ChkBRAgreement.Enabled = False
ChkAgreementNotReceived.Value = vbUnchecked
TxtCreated = ""
TxtUpdated = ""
SaveCmd.Enabled = False
ExitCmd.Enabled = True
End Function

Private Sub Gen_Email()
Dim TOADD As Variant
Dim TOCC, mMsg, mCMPMEmailList As Variant
Dim mR As Variant

mMsg = ""
TOADD = ""
TOCC = ""
mCMPMEmailList = ""

mMsg = "<table border=1 width='80%' align='center'>"
mMsg = mMsg & "<tr><td width='40%'><b>Agreement ID</b></td><td width='60%'>" & TxtAgreementID.Text & "</td></tr>"
mMsg = mMsg & "<tr><td width='40%'><b>Entry Date</b></td><td width='60%'>" & Format(TxtEntryDate, "DD-MMM-YYYY") & "</td></tr>"
mMsg = mMsg & "<tr><td width='40%'><b>Start Date</b></td><td width='60%'>" & Format(TxtStartDate, "DD-MMM-YYYY") & "</td></tr>"
mMsg = mMsg & "<tr><td width='40%'><b>Expiry Date</b></td><td width='60%'>" & Format(TxtExpiryDate, "DD-MMM-YYYY") & "</td></tr>"
mMsg = mMsg & "<tr><td width='40%'><b>Tax Description</b></td><td width='60%'>" & TxtServiceTaxID.Text & "</td></tr>"
mMsg = mMsg & "<tr><td width='40%'><b>Agreement Received As</b></td><td width='60%'>" & TxtAgreeRecAs.Text & "</td></tr>"
mMsg = mMsg & "<tr><td width='40%'><b>Security Deposit</b></td><td width='60%'>" & TxtAdvanceGiven.Text & "</td></tr>"
mMsg = mMsg & "<tr><td width='40%'><b>SD Adjusted</b></td><td width='60%'>" & TxtAdvAdjusted.Text & "</td></tr>"
mMsg = mMsg & "<tr><td width='40%'><b>SD Returned</b></td><td width='60%'>" & TxtAdvanceRemain.Text & "</td></tr>"
mMsg = mMsg & "<tr><td width='40%'><b>Balance</b></td><td width='60%'>" & TxtBalance.Text & "</td></tr>"
mMsg = mMsg & "<tr><td width='40%'><b>Status</b></td><td width='60%'>" & TxtStatus.Text & "</td></tr>"
mMsg = mMsg & "<tr><td width='40%'><b>Remark</b></td><td width='60%'>" & TxtRemark.Text & "</td></tr>"
mMsg = mMsg & "<tr><td width='40%'><b>BR Agreement</b></td><td width='60%'>" & IIf(ChkBRAgreement.Value = vbChecked, "Yes", "No") & "</td></tr>"
mMsg = mMsg & "<tr><td width='40%'><b>Agreement Received</b></td><td width='60%'>" & IIf(ChkAgreementNotReceived.Value = vbChecked, "No", "Yes") & "</td></tr>"
mMsg = mMsg & "<tr><td width='40%'><b>Updated By</b></td><td width='60%'>" & Gen_UserName & "~" & Gen_UserLoginID & "</td></tr>"
mMsg = mMsg & "<tr><td width='40%'><b>Updated On</b></td><td width='60%'>" & Format(Now, "DD-MMM-YYYY hh:mm:ss AM/PM") & "</td></tr>"
mMsg = mMsg & "</table>"

If mForwardID <> Gen_UserID Then
    TOADD = "cmg@ncml.com;accounts@ncml.com"  ''-- "ullhas.p@ncml.com"
End If

tmp = " Select ME.EmployeeID,ME.EmployeeNo,ME.EmployeeName,ME.EmployeeEmailID"
tmp = tmp & " From Mst_Employee ME"
tmp = tmp & " Inner Join Mst_EmpCMPMapping A On ME.EMployeeID = A.EmployeeID"
tmp = tmp & " Where ME.Flag = 'Y' And A.CMPID In ("
tmp = tmp & " Select MC.CMPID From Txn_GodownLessorAgreeDetail AA"
tmp = tmp & " Inner Join Mst_Godown MG On AA.GodownID = MG.GodownID"
tmp = tmp & " Inner Join Mst_CMP MC On MG.CMPID = MC.CMPID"
tmp = tmp & " Where AA.LAID = " & Val(TxtAgreementID) & ")"
Call TmpTable
If TmpRs.RecordCount > 0 Then
    For mR = 1 To TmpRs.RecordCount
        mCMPMEmailList = mCMPMEmailList & "," & IIf(IsNull(TmpRs!EmployeeEmailID), "", TmpRs!EmployeeEmailID)
        TmpRs.MoveNext
    Next mR
End If

TOCC = Gen_UserEmailID & mCMPMEmailList

mMsgSubject = "Update Lessor Agreement Receive Flag : " & Val(TxtAgreementID)
Call SendMail(TOADD, TOCC, mMsgSubject, mMsg, "")

End Sub

Private Function SendMail(ToADD_ As Variant, TOCC_ As Variant, MsgSubject_ As Variant, Message1_ As Variant, mMailTypeFlg As Variant)
Dim htmlStringhead As String
Dim htmlStringtail As String
Set oSMTP = New OSSMTP.SMTPSession
servertxt = "webmail.ncml.com"
htmlStringhead = "<html>"
htmlStringhead = htmlStringhead & " <head>"
htmlStringhead = htmlStringhead & "<style> body{ color: #000000;margin: 5px;font-family: Candara;} table, th, td { border: 1px solid black; border-collapse: collapse;} th { padding: 5px; text-align: center;} td { padding: 5px; text-align: left;} </style>"
htmlStringhead = htmlStringhead & "</head><body>"
htmlStringtail = "</body></html>"
   
''With oSMTP
''  'connection
''  .Server = servertxt
''  If LCase(mMailTypeFlg) = "n" Then
''     .MailFrom = "info@ncmsl.com"
''  Else
''     .MailFrom = Gen_UserEmailID
''     '.MailFrom = "ullhas.p@ncml.com"
''  End If
'''    .SendTo = "ullhas.p@ncml.com"
'''    .CC = Gen_UserEmailID
''  .SendTo = ToADD_
''  If ToADD_ <> TOCC_ Then
''     .CC = TOCC_
''  End If
''  If LCase(mMailTypeFlg) = "n" Then
''     '.Attachments.Add (mPDFFileName)
''  End If
''  .MessageSubject = MsgSubject_ & " (from Store-Man)" '"Confidential Mail"
''  .MessageHTML = htmlStringhead & Message1_ & htmlStringtail
''  .TimeStamp = Now()
''  If LCase(Gen_EmailConfigFlag) = "y" Then
''     .SendEmail
''  End If
''End With

Dim objSMTP
Set objSMTP = CreateObject("MailBee.SMTP")
objSMTP.LicenseKey = "MBC900-5F470FAD67-D8E202A2E22927335F24F5A667BFD196"
objSMTP.ServerName = Gen_ReadMailServerIPINI("MailIP") ' = "mail.ncml.com" '"mail.ncml.com"
If LCase(mMailTypeFlg) = "n" Then
   objSMTP.FromAddr = "info@ncmsl.com"
Else
   objSMTP.FromAddr = Gen_UserEmailID
End If
If IsNull(Trim(ToADD_)) = True Or Trim(ToADD_) = "" Then
   ToADD_ = "manu.ag@ncml.com"
End If
objSMTP.ToAddr = ToADD_
If ToADD_ <> TOCC_ Then
   objSMTP.CCAddr = TOCC_
End If
If LCase(mMailTypeFlg) = "n" Then
   'objSMTP.AddAttachment mPDFFileName
End If
objSMTP.Subject = MsgSubject_ & " (from Store-Man)" 'MsgSubject_ '"Message in HTML format"
objSMTP.BodyFormat = 1
objSMTP.BodyText = htmlStringhead & Message1_ & htmlStringtail 'htmlStringhead & MsgSubject_ & " :- " & Message1_ & htmlStringtail '"<html>This word is <b>bold</b></html>"
If LCase(Gen_EmailConfigFlag) = "y" Then
   objSMTP.Send
End If

End Function

