VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmTxn_SecurityAgreement 
   Caption         =   "Security Agreement"
   ClientHeight    =   6825
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9405
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6825
   ScaleWidth      =   9405
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2685
      Left            =   120
      TabIndex        =   12
      Top             =   -120
      Visible         =   0   'False
      Width           =   14835
      Begin VB.CommandButton CmdSearch 
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
         Height          =   350
         Left            =   12420
         TabIndex        =   18
         Top             =   1035
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
         Left            =   13440
         TabIndex        =   17
         Top             =   1035
         Width           =   1000
      End
      Begin VB.ComboBox CmbSSecurityAgency 
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
         Left            =   6150
         Sorted          =   -1  'True
         TabIndex        =   16
         Top             =   570
         Width           =   3765
      End
      Begin VB.ComboBox CmbSState 
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
         Left            =   3255
         Sorted          =   -1  'True
         TabIndex        =   15
         Top             =   570
         Width           =   2805
      End
      Begin VB.ComboBox CmbSStatus 
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
         ItemData        =   "FrmTxn_SecurityAgreement.frx":0000
         Left            =   10080
         List            =   "FrmTxn_SecurityAgreement.frx":000D
         Sorted          =   -1  'True
         TabIndex        =   14
         Top             =   570
         Width           =   2415
      End
      Begin VB.TextBox TxtSAgreeID 
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
         Height          =   360
         Left            =   12585
         MaxLength       =   12
         TabIndex        =   13
         Top             =   570
         Width           =   1890
      End
      Begin MSComCtl2.DTPicker STo 
         Height          =   375
         Left            =   1695
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   555
         Width           =   1455
         _ExtentX        =   2566
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
         Format          =   109969409
         CurrentDate     =   36494
      End
      Begin MSComCtl2.DTPicker SFrom 
         Height          =   375
         Left            =   135
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   555
         Width           =   1455
         _ExtentX        =   2566
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
         Format          =   109969409
         CurrentDate     =   36494
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1455
         Left            =   120
         TabIndex        =   21
         Top             =   1440
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   2566
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
      Begin VB.Label Label45 
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
         Height          =   255
         Left            =   1995
         TabIndex        =   29
         Top             =   255
         Width           =   855
      End
      Begin VB.Label Label44 
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
         Height          =   255
         Left            =   435
         TabIndex        =   28
         Top             =   255
         Width           =   855
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
         Left            =   4620
         TabIndex        =   27
         Top             =   1065
         Width           =   2415
      End
      Begin VB.Label Label22 
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
         Left            =   7095
         TabIndex        =   26
         Top             =   1080
         Width           =   90
      End
      Begin VB.Label Label12 
         Caption         =   "Security Agency"
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
         Left            =   7440
         TabIndex        =   25
         Top             =   240
         Width           =   1560
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   4560
         TabIndex        =   24
         Top             =   255
         Width           =   465
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   10935
         TabIndex        =   23
         Top             =   255
         Width           =   555
      End
      Begin VB.Label Label24 
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
         Left            =   12885
         TabIndex        =   22
         Top             =   255
         Width           =   1230
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2880
      Left            =   120
      TabIndex        =   41
      Top             =   0
      Width           =   15015
      Begin VB.ComboBox CmbSecurityAgency 
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
         Left            =   2400
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   750
         Width           =   11325
      End
      Begin VB.TextBox TxtSecurityAgencyID 
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
         Left            =   2400
         Locked          =   -1  'True
         MaxLength       =   250
         ScrollBars      =   2  'Vertical
         TabIndex        =   63
         Top             =   750
         Width           =   11265
      End
      Begin VB.ComboBox CmbState 
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
         Left            =   2400
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   1200
         Width           =   11325
      End
      Begin VB.ComboBox CmbServiceTax 
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
         ItemData        =   "FrmTxn_SecurityAgreement.frx":002E
         Left            =   11205
         List            =   "FrmTxn_SecurityAgreement.frx":0038
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   1650
         Width           =   2520
      End
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
         ItemData        =   "FrmTxn_SecurityAgreement.frx":0052
         Left            =   11205
         List            =   "FrmTxn_SecurityAgreement.frx":005F
         Sorted          =   -1  'True
         TabIndex        =   44
         Top             =   270
         Width           =   2520
      End
      Begin VB.TextBox TxtAgreementID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   375
         Left            =   2415
         Locked          =   -1  'True
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   270
         Width           =   1440
      End
      Begin VB.TextBox TxtRemark 
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
         Height          =   600
         Left            =   2415
         MaxLength       =   2000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   2160
         Width           =   11310
      End
      Begin MSComCtl2.DTPicker TxtEntryDate 
         Height          =   375
         Left            =   6960
         TabIndex        =   43
         Top             =   270
         Width           =   1455
         _ExtentX        =   2566
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
         Format          =   109969409
         CurrentDate     =   39884
      End
      Begin MSComCtl2.DTPicker TxtStartDate 
         Height          =   375
         Left            =   2415
         TabIndex        =   2
         Top             =   1680
         Width           =   1455
         _ExtentX        =   2566
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
         Format          =   109969409
         CurrentDate     =   39884
      End
      Begin MSComCtl2.DTPicker TxtExpiryDate 
         Height          =   375
         Left            =   6960
         TabIndex        =   3
         Top             =   1650
         Width           =   1455
         _ExtentX        =   2566
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
         Format          =   109969409
         CurrentDate     =   39884
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
         Left            =   11205
         Locked          =   -1  'True
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   270
         Width           =   2520
      End
      Begin VB.TextBox TxtServiceTaxID 
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
         Left            =   11205
         MaxLength       =   12
         TabIndex        =   46
         Top             =   1650
         Width           =   2520
      End
      Begin VB.TextBox TxtStateID 
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
         Left            =   2400
         Locked          =   -1  'True
         MaxLength       =   250
         ScrollBars      =   2  'Vertical
         TabIndex        =   64
         Top             =   1200
         Width           =   11265
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "State"
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
         Left            =   135
         TabIndex        =   62
         Top             =   1275
         Width           =   555
      End
      Begin VB.Label Label25 
         Caption         =   "Security Agency"
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
         Left            =   135
         TabIndex        =   59
         Top             =   810
         Width           =   2010
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
         TabIndex        =   53
         Top             =   330
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
         Left            =   4320
         TabIndex        =   52
         Top             =   330
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
         Left            =   135
         TabIndex        =   51
         Top             =   2280
         Width           =   825
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Agreement Start Date"
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
         Left            =   135
         TabIndex        =   50
         Top             =   1710
         Width           =   2235
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Agreement Expiry Date"
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
         Left            =   4200
         TabIndex        =   49
         Top             =   1710
         Width           =   2400
      End
      Begin VB.Label Label15 
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
         Height          =   255
         Left            =   9045
         TabIndex        =   48
         Top             =   330
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
         Left            =   9045
         TabIndex        =   47
         Top             =   1710
         Width           =   1380
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1080
      Left            =   120
      TabIndex        =   30
      Top             =   8760
      Width           =   15015
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   400
         Left            =   12105
         TabIndex        =   40
         Top             =   600
         Width           =   2820
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "Search"
         Height          =   400
         Left            =   12105
         TabIndex        =   39
         Top             =   195
         Width           =   2820
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   400
         Left            =   9165
         TabIndex        =   38
         Top             =   600
         Width           =   2940
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   400
         Left            =   9165
         TabIndex        =   37
         Top             =   195
         Width           =   2940
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   400
         Left            =   6225
         TabIndex        =   36
         Top             =   600
         Width           =   2940
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   400
         Left            =   6225
         TabIndex        =   35
         Top             =   195
         Width           =   2940
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   400
         Left            =   3165
         TabIndex        =   34
         Top             =   600
         Width           =   3060
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   400
         Left            =   3165
         TabIndex        =   33
         Top             =   195
         Width           =   3060
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   400
         Left            =   105
         TabIndex        =   32
         Top             =   600
         Width           =   3060
      End
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New"
         Height          =   400
         Left            =   105
         TabIndex        =   31
         Top             =   195
         Width           =   3060
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5820
      Left            =   120
      TabIndex        =   54
      Top             =   2880
      Width           =   14985
      _ExtentX        =   26432
      _ExtentY        =   10266
      _Version        =   393216
      Tab             =   1
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
      TabCaption(0)   =   "Security Guard Details"
      TabPicture(0)   =   "FrmTxn_SecurityAgreement.frx":0080
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "MSHFlexGrid2"
      Tab(0).Control(1)=   "Frame2"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Agreement Attachment"
      TabPicture(1)   =   "FrmTxn_SecurityAgreement.frx":009C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lblAttachment"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label4"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "MSHFlexGrid5"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "CmdGetFile"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "CmdAttachment"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "CmdRemoveAttachment"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "TxtAttachmentRemarks"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "Log Details"
      TabPicture(2)   =   "FrmTxn_SecurityAgreement.frx":00B8
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label20"
      Tab(2).Control(1)=   "Label21"
      Tab(2).Control(2)=   "MSHFlexGrid3"
      Tab(2).Control(3)=   "TxtUpdated"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "TxtCreated"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).ControlCount=   5
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
         Height          =   510
         Left            =   6120
         MaxLength       =   5000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   73
         Top             =   765
         Width           =   7515
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
         Left            =   11580
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   1410
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
         Left            =   9510
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   1410
         Width           =   2070
      End
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
         Left            =   7440
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   1410
         Width           =   2070
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
         Left            =   -72990
         Locked          =   -1  'True
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   480
         Width           =   7410
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
         Left            =   -72990
         Locked          =   -1  'True
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   885
         Width           =   7410
      End
      Begin VB.Frame Frame2 
         Height          =   1380
         Left            =   -74880
         TabIndex        =   55
         Top             =   360
         Width           =   14730
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
            ItemData        =   "FrmTxn_SecurityAgreement.frx":00D4
            Left            =   120
            List            =   "FrmTxn_SecurityAgreement.frx":00D6
            Sorted          =   -1  'True
            TabIndex        =   6
            Top             =   390
            Width           =   5655
         End
         Begin VB.ComboBox CmbSecurityDesignation 
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
            ItemData        =   "FrmTxn_SecurityAgreement.frx":00D8
            Left            =   5775
            List            =   "FrmTxn_SecurityAgreement.frx":00DA
            Sorted          =   -1  'True
            TabIndex        =   7
            Top             =   390
            Width           =   5295
         End
         Begin VB.TextBox TxtDetailRemark 
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
            Height          =   495
            Left            =   975
            MaxLength       =   2000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   9
            Top             =   825
            Width           =   10455
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
            Height          =   360
            Left            =   11160
            MaxLength       =   12
            TabIndex        =   8
            Top             =   390
            Width           =   1650
         End
         Begin VB.CommandButton CmdDeleteList 
            Caption         =   "Delete From List"
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
            Left            =   12870
            TabIndex        =   11
            Top             =   960
            Width           =   1710
         End
         Begin VB.CommandButton CmdAdd2List 
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
            Left            =   11535
            TabIndex        =   10
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label Label3 
            Caption         =   "Security Designation"
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
            Left            =   7312
            TabIndex        =   61
            Top             =   135
            Width           =   2220
         End
         Begin VB.Label Lbl 
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
            Height          =   300
            Left            =   2197
            TabIndex        =   60
            Top             =   135
            Width           =   1500
         End
         Begin VB.Label Label5 
            Caption         =   "Remark"
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
            TabIndex        =   57
            Top             =   975
            Width           =   1200
         End
         Begin VB.Label LblAmt 
            Caption         =   "Monthly Rate"
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
            Left            =   11258
            TabIndex        =   56
            Top             =   135
            Width           =   1455
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
         Height          =   3855
         Left            =   -74880
         TabIndex        =   58
         Top             =   1800
         Width           =   14700
         _ExtentX        =   25929
         _ExtentY        =   6800
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
         Height          =   4215
         Left            =   -74880
         TabIndex        =   67
         Top             =   1440
         Width           =   14700
         _ExtentX        =   25929
         _ExtentY        =   7435
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
         Height          =   3750
         Left            =   120
         TabIndex        =   74
         Top             =   1830
         Width           =   13545
         _ExtentX        =   23892
         _ExtentY        =   6615
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
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   7320
         TabIndex        =   76
         Top             =   480
         Width           =   2145
      End
      Begin VB.Label lblAttachment 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Path of  file selected for attachment"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1140
         Left            =   120
         TabIndex        =   75
         Top             =   510
         Width           =   5880
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
         Left            =   -74760
         TabIndex        =   69
         Top             =   540
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
         Left            =   -74760
         TabIndex        =   68
         Top             =   945
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
Attribute VB_Name = "FrmTxn_SecurityAgreement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mSecurityTypeID, mSecurityDesignationID, mSecurityAgencyID, mStateID, mStatus, Edit_Flg, mServiceTax, mPrevExpiryDate As Variant
Dim mSStatus As Variant
Dim File_Path As String
Private Sub Grid_Head1()

MSHFlexGrid2.Clear
MSHFlexGrid2.Rows = 2
MSHFlexGrid2.FixedRows = 1
MSHFlexGrid2.Cols = 7

MSHFlexGrid2.TextMatrix(0, 0) = "Detail ID"
MSHFlexGrid2.TextMatrix(0, 1) = "Security Type"
MSHFlexGrid2.TextMatrix(0, 2) = "Security Type ID"
MSHFlexGrid2.TextMatrix(0, 3) = "Security Designation"
MSHFlexGrid2.TextMatrix(0, 4) = "Security Designation ID"
MSHFlexGrid2.TextMatrix(0, 5) = "Rate"
MSHFlexGrid2.TextMatrix(0, 6) = "Remark"

MSHFlexGrid2.ColWidth(0) = 0
MSHFlexGrid2.ColWidth(1) = 3200
MSHFlexGrid2.ColWidth(2) = 0
MSHFlexGrid2.ColWidth(3) = 3200
MSHFlexGrid2.ColWidth(4) = 0
MSHFlexGrid2.ColWidth(5) = 1600
MSHFlexGrid2.ColWidth(6) = 2500

End Sub

Private Sub AddCmd_Click()

Edit_Flg = "A"
Call GButtonLock(Me, False)
Frame2.Enabled = True
Frame1.Enabled = True
SSTab1.Tab = 0
TxtAgreementID = ""
TxtEntryDate.Value = Date
TxtStartDate.Value = Date
TxtExpiryDate.Value = Date
TxtRemark = ""
CmbServiceTax.Visible = True
TxtServiceTaxID.Visible = False
CmbServiceTax.Text = ""

CmbState.Visible = True
TxtStateID.Visible = False
CmbState.Text = ""

TxtSecurityAgencyID.Visible = False
CmbSecurityAgency.Visible = True
CmbSecurityAgency.Text = ""

TxtCreated = ""
TxtUpdated = ""
CmbStatus.Visible = False
TxtStatus.Visible = True
CmbStatus.Text = "Pending"
TxtStatus.Text = "Pending"
mStatus = "P"

Call ClearFrame

Call Grid_Head1

Call Grid_Head2


lblAttachment.Enabled = True
TxtAttachmentRemarks.Locked = False
CmdAttachment.Enabled = True
CmdGetFile.Enabled = True
CmdRemoveAttachment.Enabled = True

Call Grid_Head4

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

Private Sub CmbSStatus_LostFocus()
If CmbSStatus.Text = "" Then
   Exit Sub
End If
If LCase(CmbSStatus.Text) <> "approved" And LCase(CmbSStatus.Text) <> "pending" And LCase(CmbSStatus.Text) <> "rejected" Then
   MsgBox "Invalid selection "
   CmbSStatus.SetFocus
   Exit Sub
End If
If LCase(CmbSStatus.Text) = "approved" Then
   mSStatus = "A"
ElseIf LCase(CmbSStatus.Text) = "rejected" Then
   mSStatus = "R"
ElseIf LCase(CmbSStatus.Text) = "pending" Then
   mSStatus = "P"
End If
End Sub

Private Sub CmdAdd2List_Click()

If CmbSecurityType.Text = "" Then
   MsgBox "Select Security Type"
   CmbSecurityType.SetFocus
   Exit Sub
End If

If CmbSecurityDesignation.Text = "" Then
   MsgBox "Select Security Designation"
   CmbSecurityDesignation.SetFocus
   Exit Sub
End If

If TxtRate.Text = "" Then
   MsgBox "Blank Monthly Rate not Allowed."
   TxtRate.SetFocus
   Exit Sub
End If

If CheckDuplicate = True Then
   MsgBox "Duplicate Security Type And Security Designation Not Allowed!!!"
   CmbSecurityDesignation.SetFocus
   Exit Sub
End If

If CmdAdd2List.Caption = "Update List" Then
   r = MSHFlexGrid2.RowSel
Else
   r = MSHFlexGrid2.Rows - 1
   MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
End If


MSHFlexGrid2.TextMatrix(r, 1) = CmbSecurityType.Text
MSHFlexGrid2.TextMatrix(r, 2) = mSecurityTypeID
MSHFlexGrid2.TextMatrix(r, 3) = CmbSecurityDesignation.Text
MSHFlexGrid2.TextMatrix(r, 4) = mSecurityDesignationID
MSHFlexGrid2.TextMatrix(r, 5) = Format(Val(TxtRate.Text), "#####0.00")
MSHFlexGrid2.TextMatrix(r, 6) = TxtDetailRemark.Text

Call ClearFrame

End Sub

Private Sub CmdDeleteList_Click()
I = MSHFlexGrid2.RowSel
If I > 0 Then
   If MSHFlexGrid2.TextMatrix(I, 1) <> "" Then
      For RO = I To MSHFlexGrid2.Rows - 2
          For C = 0 To MSHFlexGrid2.Cols - 1
              MSHFlexGrid2.TextMatrix(RO, C) = MSHFlexGrid2.TextMatrix(RO + 1, C)
          Next
      Next
      MSHFlexGrid2.Rows = MSHFlexGrid2.Rows - 1
   End If
End If

Call ClearFrame
End Sub
Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Txn_SecurityAgreement.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Txn_SecurityAgreement.xls")
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


Private Sub CmdSearch_Click()
Dim Wc1, mSerTax, mStatus As Variant

Call Grid_Head

If SFrom.Value > STo.Value Then
   MsgBox "From Date Must be less than To Date!!!"
   SFrom.SetFocus
   Exit Sub
End If

tmp = "Select TSA.SAID,TSA.Flag,MSA.AgencyName, MS.StateName,TSA.EntryDate,TSA.AgreeStartDate,TSA.AgreeExpiryDate,"
tmp = tmp & " TSA.ServiceTax , TSA.Remark"
tmp = tmp & " From Txn_SecurityAgreement TSA"
tmp = tmp & " Inner join Mst_SecurityAgency MSA on TSA.SecurityAgencyID=MSA.SecurityAgencyID"
tmp = tmp & " Inner join Mst_State MS on TSA.StateID=MS.StateID"

If Gen_DBType = "MDB" Then
   Wc1 = " Where TSA.EntryDate Between #" & Format(SFrom.Value, Gen_DatFormat) & "# And #" & Format(STo.Value, Gen_DatFormat) & "# "
Else
   Wc1 = " Where TSA.EntryDate Between '" & Format(SFrom.Value, Gen_DatFormat) & "' And '" & Format(STo.Value, Gen_DatFormat) & "' "
End If

If CmbSSecurityAgency.Text <> "" Then
   Wc1 = Wc1 & " And MSA.AgencyName='" & CmbSSecurityAgency.Text & "'"
End If

If CmbSState.Text <> "" Then
   Wc1 = Wc1 & " And MS.StateName='" & CmbSState.Text & "'"
End If

If CmbSStatus.Text <> "" Then
   Wc1 = Wc1 & " And TSA.Flag='" & mSStatus & "'"
End If

If TxtSAgreeID.Text <> "" Then
   Wc1 = Wc1 & " And TSA.SAID=" & Val(TxtSAgreeID) & ""
End If

tmp = tmp & Wc1 & " Order By TSA.SAID"

Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   MSHFlexGrid1.Rows = TmpRs1.RecordCount + 2
   For I = 1 To TmpRs1.RecordCount
   
       MSHFlexGrid1.TextMatrix(I, 0) = TmpRs1!SAID
       If LCase(TmpRs1!Flag) = "a" Then
          mStatus = "Approved"
       ElseIf LCase(TmpRs1!Flag) = "p" Then
          mStatus = "Pending"
       ElseIf LCase(TmpRs1!Flag) = "r" Then
          mStatus = "Rejected"
       End If
       
       MSHFlexGrid1.TextMatrix(I, 1) = mStatus
       
       MSHFlexGrid1.TextMatrix(I, 2) = TmpRs1!AgencyName
       MSHFlexGrid1.TextMatrix(I, 3) = TmpRs1!StateName
       MSHFlexGrid1.TextMatrix(I, 4) = Format(TmpRs1!EntryDate, "dd-MMM-yyyy")
       MSHFlexGrid1.TextMatrix(I, 5) = Format(TmpRs1!AgreeStartDate, "dd-MMM-yyyy")
       MSHFlexGrid1.TextMatrix(I, 6) = Format(TmpRs1!AgreeExpiryDate, "dd-MMM-yyyy")
       mStatus = ""
       If LCase(TmpRs1!ServiceTax) = "i" Then
          mStatus = "Inclusive"
       ElseIf LCase(TmpRs1!ServiceTax) = "e" Then
          mStatus = "Exclusive"
       End If
       MSHFlexGrid1.TextMatrix(I, 7) = mStatus
       MSHFlexGrid1.TextMatrix(I, 8) = TmpRs1!Remark
       
       TmpRs1.MoveNext
   Next
   
End If
Label22.Caption = TmpRs1.RecordCount
Label22.Refresh

End Sub

Private Sub EditCmd_Click()

If LCase(Gen_UserRole) <> "fag" And LCase(Gen_UserRole) <> "fagadmin" And LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "headcmg" Then
   If LCase(Gen_UserRole) <> "admin" Then
      If LCase(TxtStatus) = "approved" Or LCase(TxtStatus) = "rejected" Then
         MsgBox "This agreement is " & TxtStatus & ". You can't edit this agreement any more!!!"
         Exit Sub
      End If
   Else
      If LCase(TxtStatus) = "approved" Or LCase(TxtStatus) = "rejected" Then
         Edit_Flg = "E"
         Call GButtonLock(Me, False)
         lblAttachment.Enabled = True
         TxtAttachmentRemarks.Locked = False
         CmdAttachment.Enabled = True
         CmdGetFile.Enabled = True
         CmdRemoveAttachment.Enabled = True
         Exit Sub
      End If
   End If
End If
'
'If CheckNextAgreement = True Then
'   MsgBox "You Can't Edit this Record. New Agreement is created for one or more godowns"
'   Exit Sub
'End If

Edit_Flg = "E"

Call GButtonLock(Me, False)

Frame2.Enabled = True
Frame1.Enabled = True
'Or LCase(Gen_UserRole) = "superadmin"

If LCase(Gen_UserRole) = "fag" Or LCase(Gen_UserRole) = "power" Or LCase(Gen_UserRole) = "fagadmin" Or LCase(Gen_UserRole) = "headcmg" Then
   CmbStatus.Visible = True
   TxtStatus.Visible = False
   CmbStatus.Text = TxtStatus.Text
End If

CmbServiceTax.Visible = True
TxtServiceTaxID.Visible = False
CmbServiceTax.Text = TxtServiceTaxID.Text
CmbSecurityAgency.Text = TxtSecurityAgencyID.Text
CmbState.Text = TxtStateID.Text
lblAttachment.Enabled = True
TxtAttachmentRemarks.Locked = False
CmdAttachment.Enabled = True
CmdGetFile.Enabled = True
CmdRemoveAttachment.Enabled = True


SSTab1.Tab = 0


End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim wc As Variant

Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Frame2.Enabled = False
   Frame1.Enabled = False
   Call GButtonLockAD(Me)
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If
Call Grid_Head
SFrom.Value = Date - 365
STo.Value = Date
Label22.Caption = ""
CmbSSecurityAgency.Text = ""
CmbSState.Text = ""
CmbSStatus.Text = ""
TxtSAgreeID.Text = ""

Call TableTxn_SecurityAgreement(" Order by SAID")
If RsTxn_SecurityAgreement.RecordCount = 0 Then
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

Private Sub CmbStatus_LostFocus()
If CmbStatus.Text = "" Then
   Exit Sub
End If
If LCase(CmbStatus.Text) <> "approved" And LCase(CmbStatus.Text) <> "pending" And LCase(CmbStatus.Text) <> "rejected" Then
   MsgBox "Invalid selection "
   CmbStatus.SetFocus
   Exit Sub
End If
End Sub


Private Sub CmbSecurityAgency_GotFocus()
tmp = CmbSecurityAgency.Text
Call TableMst_SecurityAgency("Where Flag = 'Y'")
CmbSecurityAgency.Clear
If RsMst_SecurityAgency.RecordCount = 0 Then
   MsgBox "No Security Agency Found!!!"
   CmbSecurityAgency.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_SecurityAgency.RecordCount
    CmbSecurityAgency.AddItem RsMst_SecurityAgency!AgencyName
    RsMst_SecurityAgency.MoveNext
Next
CmbSecurityAgency.Text = tmp

End Sub

Private Sub CmbSecurityAgency_LostFocus()
If CmbSecurityAgency.Text = "" Then
   mSecurityAgencyID = Null
   Exit Sub
End If

RsMst_SecurityAgency.MoveFirst
RsMst_SecurityAgency.Find "AgencyName = '" & CmbSecurityAgency.Text & "'"

If RsMst_SecurityAgency.EOF Then
   MsgBox "Invalid Selection "
   CmbSecurityAgency.SetFocus
   Exit Sub
End If
mSecurityAgencyID = RsMst_SecurityAgency!SecurityAgencyID
End Sub


Private Sub CmbState_GotFocus()
tmp = CmbState.Text
Call TableMst_State(" Where Flag = 'A'")

CmbState.Clear
If RsMst_State.RecordCount = 0 Then
   MsgBox "No state found"
   CmdState.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_State.RecordCount
    CmbState.AddItem RsMst_State!StateName
    RsMst_State.MoveNext
Next
CmbState.Text = tmp
End Sub

Private Sub CmbState_LostFocus()
If CmbState.Text = "" Then
   Exit Sub
End If
RsMst_State.MoveFirst
RsMst_State.Find "StateName = '" & CmbState.Text & "'"

If RsMst_State.EOF Then
   MsgBox "Invalid State selection "
   CmbState.SetFocus
   Exit Sub
End If

mStateID = RsMst_State.Fields(0)

End Sub

Private Sub CmbServiceTax_LostFocus()
If CmbServiceTax.Text = "" Then Exit Sub
If LCase(CmbServiceTax.Text) <> "inclusive" And LCase(CmbServiceTax.Text) <> "exclusive" Then
   MsgBox "Invalid Selection"
   CmbServiceTax.SetFocus
   Exit Sub
End If
If LCase(CmbServiceTax.Text) = "inclusive" Then
   mServiceTax = "I"
Else
   mServiceTax = "E"
End If

End Sub

Private Sub CmbSecurityType_GotFocus()
tmp = CmbSecurityType.Text
Call TableMst_SecurityShifts("")


CmbSecurityType.Clear
If RsMst_SecurityShifts.RecordCount = 0 Then
   MsgBox "No Record found!!!"
   'CmbSecurityType.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_SecurityShifts.RecordCount
    CmbSecurityType.AddItem RsMst_SecurityShifts!SecurityShifts
    RsMst_SecurityShifts.MoveNext
Next
CmbSecurityType.Text = tmp
End Sub

Private Sub CmbSecurityType_LostFocus()
If CmbSecurityType.Text = "" Then Exit Sub

RsMst_SecurityShifts.MoveFirst
RsMst_SecurityShifts.Find "SecurityShifts = '" & CmbSecurityType.Text & "'"

If RsMst_SecurityShifts.EOF Then
   MsgBox "Invalid Selection!!!"
   CmbSecurityType.SetFocus
   Exit Sub
End If
mSecurityTypeID = RsMst_SecurityShifts!SecurityShiftsID
End Sub

Private Sub CmbSecurityDesignation_GotFocus()
tmp = CmbSecurityDesignation.Text
Call TableMst_SecurityDesignation("")

CmbSecurityDesignation.Clear
If RsMst_SecurityDesignation.RecordCount = 0 Then
   MsgBox "No Record found!!!"
   'CmbSecurityType.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_SecurityDesignation.RecordCount
    CmbSecurityDesignation.AddItem RsMst_SecurityDesignation!SecurityDesignation
    RsMst_SecurityDesignation.MoveNext
Next
CmbSecurityDesignation.Text = tmp

End Sub

Private Sub CmbSecurityDesignation_LostFocus()
If CmbSecurityDesignation.Text = "" Then Exit Sub
      
RsMst_SecurityDesignation.MoveFirst
RsMst_SecurityDesignation.Find "SecurityDesignation = '" & CmbSecurityDesignation.Text & "'"

If RsMst_SecurityDesignation.EOF Then
   MsgBox "Invalid Selection!!!"
   CmbSecurityDesignation.SetFocus
   Exit Sub
End If
mSecurityDesignationID = RsMst_SecurityDesignation!SecurityDesignationID
End Sub


Private Sub MSHFlexGrid1_Click()

Dim I As Variant

I = MSHFlexGrid1.RowSel
If I <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(I, 0) = "" Then Exit Sub
RsTxn_SecurityAgreement.MoveFirst
RsTxn_SecurityAgreement.Find "SAID= " & Val(MSHFlexGrid1.TextMatrix(I, 0)) & ""
If Not RsTxn_SecurityAgreement.EOF Then
   Call Indata
   Frame0.Visible = False
End If
End Sub

Private Sub MSHFlexGrid2_Click()
If SaveCmd.Enabled = False Then Exit Sub
Dim r As Variant
r = MSHFlexGrid2.RowSel
If r < 1 Or r = MSHFlexGrid2.Rows - 1 Then
   Call ClearFrame
   Exit Sub
End If
If MSHFlexGrid2.TextMatrix(r, 1) = "" Then Exit Sub
mSecurityTypeID = MSHFlexGrid2.TextMatrix(r, 2)
CmbSecurityType.Text = MSHFlexGrid2.TextMatrix(r, 1)
mSecurityDesignationID = MSHFlexGrid2.TextMatrix(r, 4)
CmbSecurityDesignation.Text = MSHFlexGrid2.TextMatrix(r, 3)
TxtRate.Text = MSHFlexGrid2.TextMatrix(r, 5)
TxtDetailRemark.Text = MSHFlexGrid2.TextMatrix(r, 6)

CmdAdd2List.Caption = "Update List"
CmbSecurityType.SetFocus

End Sub
Private Sub FirstCmd_Click()
RsTxn_SecurityAgreement.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsTxn_SecurityAgreement.MoveNext
If RsTxn_SecurityAgreement.EOF Then
   RsTxn_SecurityAgreement.MoveLast
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
Private Sub LastCmd_Click()
RsTxn_SecurityAgreement.MoveLast
Call Indata
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False
End Sub

Private Sub PreviousCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsTxn_SecurityAgreement.MovePrevious
If RsTxn_SecurityAgreement.BOF Then
   RsTxn_SecurityAgreement.MoveFirst
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


If Frame1.Enabled = True Then

   If TxtStartDate.Value > Date Then
      MsgBox " Future Agreement Start Date not allowed!!"
      TxtStartDate.SetFocus
      Exit Sub
   End If

   If TxtExpiryDate <= TxtStartDate Then
      MsgBox "Expiry Date should be Greater than Start Date."
      TxtExpiryDate.SetFocus
      Exit Sub
   End If

   If CmbSecurityAgency.Text = "" Then
      MsgBox "Select Security Agency!!"
      CmbSecurityAgency.SetFocus
      Exit Sub
   End If

   If CmbState.Text = "" Then
      MsgBox "Select State!!"
      CmbState.SetFocus
      Exit Sub
   End If

   If CheckPreviousExpiry() = False Then
      MsgBox "Agreemet of this Security Agency & State is already exist and will expire on " & mPrevExpiryDate & " ."
'   CmbSecurityAgency.SetFocus
      Exit Sub
   End If

   If MSHFlexGrid2.Rows < 3 Then
      MsgBox "Map atleast one security guard!!"
      Exit Sub
   End If

   If CmbStatus.Visible = True Then
      If LCase(CmbStatus.Text) = "approved" Then
         mStatus = "A"
      ElseIf LCase(CmbStatus.Text) = "rejected" Then
         mStatus = "R"
      ElseIf LCase(CmbStatus.Text) = "pending" Then
         mStatus = "P"
      End If
   Else
      mStatus = "P"
   End If
End If

If Edit_Flg = "A" Then
   RsTxn_SecurityAgreement.AddNew
   RsTxn_SecurityAgreement!SAID = getMaxSAID
   RsTxn_SecurityAgreement!G_UID = GetGUID
   TxtAgreementID.Text = RsTxn_SecurityAgreement!SAID
Else
   RsTxn_SecurityAgreement.MoveFirst
   RsTxn_SecurityAgreement.Find " SAID= " & Val(TxtAgreementID.Text) & ""
   
End If

RsTxn_SecurityAgreement!SAID = Val(TxtAgreementID.Text)
RsTxn_SecurityAgreement!EntryDate = Format(TxtEntryDate.Value, Gen_DatFormat)
RsTxn_SecurityAgreement!SecurityAgencyID = mSecurityAgencyID
RsTxn_SecurityAgreement!StateID = mStateID
RsTxn_SecurityAgreement!AgreeStartDate = Format(TxtStartDate.Value, Gen_DatFormat)
RsTxn_SecurityAgreement!AgreeExpiryDate = Format(TxtExpiryDate.Value, Gen_DatFormat)
RsTxn_SecurityAgreement!ServiceTax = mServiceTax
RsTxn_SecurityAgreement!Remark = TxtRemark.Text
RsTxn_SecurityAgreement!Flag = mStatus
If IsNull(RsTxn_SecurityAgreement!CreatedBy) = True Or RsTxn_SecurityAgreement!CreatedBy = "" Then
   RsTxn_SecurityAgreement!CreatedBy = Gen_UserLoginID
   RsTxn_SecurityAgreement!CreatedDate = Now
Else
   RsTxn_SecurityAgreement!UpdatedBy = Gen_UserLoginID
   RsTxn_SecurityAgreement!UpdatedDate = Now
End If

RsTxn_SecurityAgreement.Update

Call SaveSecurityDetails
Call SaveAttachmentDetails

SaveCmd.Enabled = False
Call Indata


End Sub

Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   If RsTxn_SecurityAgreement.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   RsTxn_SecurityAgreement.MoveFirst
   Call ClearFrame
   RsTxn_SecurityAgreement.Find "SAID = " & Val(TxtAgreementID) & ""
   If RsTxn_SecurityAgreement.EOF Then
      RsTxn_SecurityAgreement.MoveLast
   End If
   Call Indata
   Exit Sub
End If

Frame0.Height = Me.Height - 500
Frame0.Left = Frame1.Left - 100
Frame0.Top = Frame1.Top
Frame0.Width = Me.Width - 200
Frame0.Visible = True

MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 2500

RsTxn_SecurityAgreement.MoveFirst
RsTxn_SecurityAgreement.Filter = ""
End Sub

Private Sub TxtRate_Lostfocus()
If TxtRate.Text = "" Then Exit Sub
If IsNumeric(TxtRate) = False Then
   MsgBox "Invalid Numeric Format !!!!"
   TxtRate.SetFocus
   Exit Sub
End If
TxtRate = Val(TxtRate)
End Sub

Function CheckDuplicate() As Boolean
Dim Retval As Boolean
Retval = False
If MSHFlexGrid2.Rows > 1 Then
   For u = 1 To MSHFlexGrid2.Rows - 1
       If MSHFlexGrid2.TextMatrix(u, 2) = mSecurityTypeID And MSHFlexGrid2.TextMatrix(u, 4) = mSecurityDesignationID Then
          If LCase(CmdAdd2List.Caption) = "add to list" Then
             Retval = True
             Exit For
          Else
             If u <> MSHFlexGrid2.RowSel Then
                Retval = True
             End If
          End If
       End If
   Next
End If
CheckDuplicate = Retval
End Function

Private Sub ClearFrame()
CmbSecurityType.Text = ""
mSecurityDesignationID = Null
CmbSecurityDesignation.Text = ""
mSecurityTypeID = Null
TxtRate = ""
TxtDetailRemark.Text = ""
CmdAdd2List.Caption = "Add to List"
End Sub

Private Sub Grid_Head2()

MSHFlexGrid3.Clear
MSHFlexGrid3.Rows = 2
MSHFlexGrid3.FixedRows = 1
MSHFlexGrid3.Cols = 3

MSHFlexGrid3.TextMatrix(0, 0) = "Agreement ID"
MSHFlexGrid3.TextMatrix(0, 1) = "Updated By"
MSHFlexGrid3.TextMatrix(0, 2) = "Updated On"

MSHFlexGrid3.ColWidth(0) = 2200
MSHFlexGrid3.ColWidth(1) = 2000
MSHFlexGrid3.ColWidth(2) = 3000

End Sub

Private Function getMaxSAID() As Double
Dim Retval As Double
tmp = "Select max(SAID) from Txn_SecurityAgreement"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
getMaxSAID = Retval
End Function
Private Function getMaxSADID() As Double
Dim Retval As Double
tmp = "Select max(SADID) from Txn_SecurityAgreementDetail"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
getMaxSADID = Retval
End Function

Private Sub SaveSecurityDetails()

Call DeleteChildTableDetail

Call TableTxn_SecurityAgreementDetail(" Where SAID=" & Val(TxtAgreementID) & "")

For I = 1 To MSHFlexGrid2.Rows - 1
    If MSHFlexGrid2.TextMatrix(I, 1) <> "" Then
       RsTxn_SecurityAgreementDetail.AddNew
       RsTxn_SecurityAgreementDetail!G_UID = GetGUID
       RsTxn_SecurityAgreementDetail!SADID = getMaxSADID
       RsTxn_SecurityAgreementDetail!SecurityTypeID = Val(MSHFlexGrid2.TextMatrix(I, 2))
       RsTxn_SecurityAgreementDetail!SAID = Val(TxtAgreementID)
       RsTxn_SecurityAgreementDetail!SecurityDesignationID = Val(MSHFlexGrid2.TextMatrix(I, 4))
       RsTxn_SecurityAgreementDetail!MonthlyRate = Val(MSHFlexGrid2.TextMatrix(I, 5))
       RsTxn_SecurityAgreementDetail!Remark = MSHFlexGrid2.TextMatrix(I, 6)
       RsTxn_SecurityAgreementDetail.Update
       RsTxn_SecurityAgreementDetail.Requery
    End If
Next
End Sub

Sub DeleteChildTableDetail()
tmp = "Delete from Txn_SecurityAgreementDetail where SAID = " & Val(TxtAgreementID.Text) & ""
Call TmpTable

End Sub


Public Sub Indata()

Call ClearFrame

Frame2.Enabled = False
Frame1.Enabled = False
SSTab1.Tab = 0

Call Grid_Head1
Call Grid_Head2

CmbStatus.Visible = False
TxtStatus.Visible = True

CmbState.Visible = False
TxtStateID.Visible = True

CmbSecurityAgency.Visible = False
TxtSecurityAgencyID.Visible = True

CmbServiceTax.Visible = False
TxtServiceTaxID.Visible = True

TxtAgreementID = IIf(IsNull(RsTxn_SecurityAgreement!SAID), "", RsTxn_SecurityAgreement!SAID)
TxtStartDate.Value = RsTxn_SecurityAgreement!AgreeStartDate
TxtEntryDate.Value = RsTxn_SecurityAgreement!EntryDate
TxtExpiryDate.Value = RsTxn_SecurityAgreement!AgreeExpiryDate

mStateID = RsTxn_SecurityAgreement!StateID
mSecurityAgencyID = RsTxn_SecurityAgreement!SecurityAgencyID

TxtStateID.Text = GetStateName(RsTxn_SecurityAgreement!StateID) 'GetStateIndiaFromStateName(GetStateName(RsTxn_SecurityAgreement!StateID))
TxtSecurityAgencyID.Text = GetSecurityAgency(RsTxn_SecurityAgreement!SecurityAgencyID, "N")
mServiceTax = RsTxn_SecurityAgreement!ServiceTax
TxtServiceTaxID = IIf(IsNull(RsTxn_SecurityAgreement!ServiceTax), "", IIf(RsTxn_SecurityAgreement!ServiceTax = "I", "Inclusive", "Exclusive"))
TxtRemark = IIf(IsNull(RsTxn_SecurityAgreement!Remark), "", RsTxn_SecurityAgreement!Remark)
TxtCreated = IIf(IsNull(RsTxn_SecurityAgreement!CreatedBy), "", RsTxn_SecurityAgreement!CreatedBy) & " :- " & IIf(IsNull(RsTxn_SecurityAgreement!CreatedDate), "", RsTxn_SecurityAgreement!CreatedDate)
TxtUpdated = IIf(IsNull(RsTxn_SecurityAgreement!UpdatedBy), "", RsTxn_SecurityAgreement!UpdatedBy) & " :- " & IIf(IsNull(RsTxn_SecurityAgreement!UpdatedDate), "", RsTxn_SecurityAgreement!UpdatedDate)

mStatus = IIf(IsNull(RsTxn_SecurityAgreement!Flag), "", RsTxn_SecurityAgreement!Flag)
If LCase(mStatus) = "a" Then
   TxtStatus = "Approved"
ElseIf LCase(mStatus) = "p" Then
   TxtStatus = "Pending"
ElseIf LCase(mStatus) = "r" Then
   TxtStatus = "Rejected"
Else
   TxtStatus = ""
End If


'TxtG_UID = IIf(IsNull(RsTxn_SecurityAgreement!G_UID), "", RsTxn_SecurityAgreement!G_UID)

Call InDataSecurityDetail
Call IndataAttachmentDetails

'Call InDataLogHistory
If RsTxn_SecurityAgreement.RecordCount = 1 Then
   Call GButtonLock_1(Me, True)
Else
   Call GButtonLock(Me, True)
End If

End Sub

Private Sub InDataSecurityDetail()

tmp = "Select TSAD.SADID,MST.SecurityShifts,TSAD.SecurityTypeID,MSD.SecurityDesignation,TSAD.SecurityDesignationID ,TSAD.MonthlyRate,TSAD.Remark"
tmp = tmp & " from Txn_SecurityAgreementDetail TSAD"
tmp = tmp & " Inner join Mst_SecurityShifts MST on MST.SecurityShiftsID=TSAD.SecurityTypeID"
tmp = tmp & " Inner join Mst_SecurityDesignation MSD on MSD.SecurityDesignationID=TSAD.SecurityDesignationID"
tmp = tmp & " Where TSAD.SAID=" & Val(TxtAgreementID) & ""

Call Tmp1Table

Call Grid_Head1

MSHFlexGrid2.Rows = TmpRs1.RecordCount + 2
For I = 1 To MSHFlexGrid2.Rows - 2
   MSHFlexGrid2.TextMatrix(I, 0) = TmpRs1!SADID
   MSHFlexGrid2.TextMatrix(I, 1) = TmpRs1!SecurityShifts
   MSHFlexGrid2.TextMatrix(I, 2) = TmpRs1!SecurityTypeID
   MSHFlexGrid2.TextMatrix(I, 3) = TmpRs1!SecurityDesignation
   MSHFlexGrid2.TextMatrix(I, 4) = TmpRs1!SecurityDesignationID
   MSHFlexGrid2.TextMatrix(I, 5) = Format(TmpRs1!MonthlyRate, "#####0.00")
   MSHFlexGrid2.TextMatrix(I, 6) = IIf(IsNull(TmpRs1!Remark), "", TmpRs1!Remark)
   
   TmpRs1.MoveNext
Next
End Sub
''Private Sub InDataLogHistory()
''
''Call Grid_Head3
''tmp = "Select LAID, UpdatedBy, UpdatedOn from Log_LessorAgreementHistory Where LAID=" & Val(TxtAgreementID) & ""
''Call TmpTable
''MSHFlexGrid4.Rows = MSHFlexGrid4.Rows + TmpRs.RecordCount
''For I = 1 To MSHFlexGrid4.Rows - 2
''   MSHFlexGrid4.TextMatrix(I, 0) = TmpRs!LAID
''   MSHFlexGrid4.TextMatrix(I, 1) = IIf(IsNull(TmpRs!UpdatedBy), "", TmpRs!UpdatedBy)
''   MSHFlexGrid4.TextMatrix(I, 2) = TmpRs!UpdatedOn
''   TmpRs.MoveNext
''Next
''End Sub


Private Sub CmbSState_GotFocus()
tmp = CmbSState.Text
Call TableMst_State("")

CmbSState.Clear
If RsMst_State.RecordCount = 0 Then
   MsgBox "No state found"
   CmbSState.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_State.RecordCount
    CmbSState.AddItem RsMst_State!StateName
    RsMst_State.MoveNext
Next
CmbSState.Text = tmp
End Sub

Private Sub CmbSState_LostFocus()
If CmbSState.Text = "" Then
   Exit Sub
End If
RsMst_State.MoveFirst
RsMst_State.Find "StateName = '" & CmbSState.Text & "'"

If RsMst_State.EOF Then
   MsgBox "Invalid State selection "
   CmbSState.SetFocus
   Exit Sub
End If

mSStateID = RsMst_State.Fields(0)

End Sub


Private Sub CmbSSecurityAgency_GotFocus()
tmp = CmbSSecurityAgency.Text
Call TableMst_SecurityAgency("Where Flag = 'Y'")
CmbSSecurityAgency.Clear
If RsMst_SecurityAgency.RecordCount = 0 Then
   MsgBox "No Security Agency Found!!!"
   CmbSecurityAgency.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_SecurityAgency.RecordCount
    CmbSSecurityAgency.AddItem RsMst_SecurityAgency!AgencyName
    RsMst_SecurityAgency.MoveNext
Next
CmbSSecurityAgency.Text = tmp

End Sub

Private Sub CmbSSecurityAgency_LostFocus()
If CmbSSecurityAgency.Text = "" Then
   mSSecurityAgencyID = Null
   Exit Sub
End If

RsMst_SecurityAgency.MoveFirst
RsMst_SecurityAgency.Find "AgencyName = '" & CmbSSecurityAgency.Text & "'"

If RsMst_SecurityAgency.EOF Then
   MsgBox "Invalid Selection "
   CmbSSecurityAgency.SetFocus
   Exit Sub
End If
mSSecurityAgencyID = RsMst_SecurityAgency!SecurityAgencyID
End Sub


Private Sub Grid_Head()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.FixedRows = 1
MSHFlexGrid1.Cols = 9
MSHFlexGrid1.WordWrap = True

MSHFlexGrid1.TextMatrix(0, 0) = "Security Agreement ID"
MSHFlexGrid1.TextMatrix(0, 1) = "Status"
MSHFlexGrid1.TextMatrix(0, 2) = "Security Agency Name"
MSHFlexGrid1.TextMatrix(0, 3) = "State Name"
MSHFlexGrid1.TextMatrix(0, 4) = "Entry Date"
MSHFlexGrid1.TextMatrix(0, 5) = "Start Date"
MSHFlexGrid1.TextMatrix(0, 6) = "Expiry Date"
MSHFlexGrid1.TextMatrix(0, 7) = "Service Tax"
MSHFlexGrid1.TextMatrix(0, 8) = "Remark"

MSHFlexGrid1.ColWidth(0) = 1200
MSHFlexGrid1.ColWidth(1) = 1200
MSHFlexGrid1.ColWidth(2) = 3000
MSHFlexGrid1.ColWidth(3) = 2800
MSHFlexGrid1.ColWidth(4) = 1400
MSHFlexGrid1.ColWidth(5) = 1400
MSHFlexGrid1.ColWidth(6) = 1400
MSHFlexGrid1.ColWidth(7) = 1400
MSHFlexGrid1.ColWidth(8) = 3500

MSHFlexGrid1.RowHeight(0) = 800

End Sub


Private Function CheckPreviousExpiry() As Boolean

Dim Retval As Boolean
Retval = True

tmp = "Select TSA.SAID, TSA.AgreeExpiryDate as ExpiryDate  from "
tmp = tmp & " Txn_SecurityAgreement TSA where TSA.FLag <> 'R' "
tmp = tmp & " And TSA.StateID=" & mStateID & " And TSA.SecurityAgencyID=" & mSecurityAgencyID
If Val(TxtAgreementID.Text) > 0 Then
   tmp = tmp & " And TSA.SAID<> " & Val(TxtAgreementID.Text)
End If
tmp = tmp & " Order by TSA.AgreeExpiryDate Desc"

Call TmpTable

If TmpRs.RecordCount > 0 Then
 
   If CDate(TmpRs!ExpiryDate) > TxtStartDate.Value Then   ' for Check Expiry Date of New Agreement
      mPrevExpiryDate = TmpRs!ExpiryDate
      Retval = False
   End If

End If

CheckPreviousExpiry = Retval

End Function

Private Sub IndataAttachmentDetails()
Call Grid_Head4
Call ClearFrameAttachment

Call TableTxn_SecurityAgreeAttachment(" Where SAID = " & TxtAgreementID.Text & " ")

MSHFlexGrid5.Rows = RsTxn_SecurityAgreeAttachment.RecordCount + 2
For I = 1 To RsTxn_SecurityAgreeAttachment.RecordCount
'Security
    MSHFlexGrid5.TextMatrix(I, 0) = I
    MSHFlexGrid5.TextMatrix(I, 1) = RsTxn_SecurityAgreeAttachment!FileName
    MSHFlexGrid5.TextMatrix(I, 2) = RsTxn_SecurityAgreeAttachment!Remarks
    'MSHFlexGrid5.TextMatrix(i, 3) = App.Path & "\Agreement Attachment\" & MSHFlexGrid5.TextMatrix(i, 1)
    MSHFlexGrid5.TextMatrix(I, 3) = App.Path & "\Agreement Attachment\Security\" & MSHFlexGrid5.TextMatrix(I, 1)
    File_Path = MSHFlexGrid5.TextMatrix(I, 3) 'App.Path & "\Project Attachment\"
    Call LoadFileFromDB(File_Path, RsTxn_SecurityAgreeAttachment, "File")
    MSHFlexGrid5.TextMatrix(I, 4) = RsTxn_SecurityAgreeAttachment!TxnID
    RsTxn_SecurityAgreeAttachment.MoveNext
Next
lblAttachment.Enabled = False
TxtAttachmentRemarks.Locked = True
CmdAttachment.Enabled = False
CmdGetFile.Enabled = False
CmdRemoveAttachment.Enabled = False
End Sub

Private Sub ClearFrameAttachment()

lblAttachment.Caption = "File Attachment"
CmdAttachment.Caption = "Add Attachment"
TxtAttachmentRemarks = ""

End Sub

Private Sub SaveAttachmentDetails()

If MSHFlexGrid5.Rows <= 1 Then Exit Sub


Call TableTxn_SecurityAgreeAttachment

For I = 1 To MSHFlexGrid5.Rows - 1
    If MSHFlexGrid5.TextMatrix(I, 1) = "" Then Exit For
    If MSHFlexGrid5.TextMatrix(I, 4) = "" Then
       RsTxn_SecurityAgreeAttachment.AddNew
 
       RsTxn_SecurityAgreeAttachment!TxnID = GetMaxAttachID
       RsTxn_SecurityAgreeAttachment!SAID = TxtAgreementID
       mFlag = SaveFileToDB(MSHFlexGrid5.TextMatrix(I, 3), RsTxn_SecurityAgreeAttachment, "File")
       mFlag = GetFileFromPath(MSHFlexGrid5.TextMatrix(I, 3))
       RsTxn_SecurityAgreeAttachment!FileName = "SecurityAgreement" & TxtAgreementID & "-" & mFlag
    Else
       RsTxn_SecurityAgreeAttachment.MoveFirst
       RsTxn_SecurityAgreeAttachment.Find "TxnID = " & MSHFlexGrid5.TextMatrix(I, 4) & ""
    End If
    RsTxn_SecurityAgreeAttachment!Remarks = MSHFlexGrid5.TextMatrix(I, 2)
    RsTxn_SecurityAgreeAttachment!UpdatedDate = Now
    RsTxn_SecurityAgreeAttachment!UpdatedBy = Gen_UserLoginID
    RsTxn_SecurityAgreeAttachment!G_UID = GetGUID
    If IsNull(RsTxn_SecurityAgreeAttachment!CreatedBy) = True Or RsTxn_SecurityAgreeAttachment!CreatedBy = "" Then
       RsTxn_SecurityAgreeAttachment!CreatedBy = Gen_UserLoginID
       RsTxn_SecurityAgreeAttachment!CreatedDate = Now
    Else
       RsTxn_SecurityAgreeAttachment!UpdatedBy = Gen_UserLoginID
       RsTxn_SecurityAgreeAttachment!UpdatedDate = Now
    End If
    RsTxn_SecurityAgreeAttachment.Update
Next

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
   lblAttachment.Caption = strFilename 'GetFileFromPath(strFilename)
End If
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

Call ClearFrameAttachment
End Sub

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
Call ClearFrameAttachment

End Sub


Private Sub MSHFlexGrid5_Click()
Dim ans, r As Variant
Dim ID As Integer
r = MSHFlexGrid5.RowSel
If r < 1 Then Exit Sub
If MSHFlexGrid5.TextMatrix(r, 1) = "" Then Exit Sub

If SaveCmd.Enabled = False Then
'Security
   ans = MsgBox("Do You want to view this document ", vbYesNo)
   If ans = vbYes Then
      If Right(App.Path, 1) = "\" Then
         Shell "Explorer.exe " & App.Path & "Agreement Attachment\Security", vbMaximizedFocus
      Else
         Shell "Explorer.exe " & App.Path & "\Agreement Attachment\Security\", vbMaximizedFocus
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

Private Function GetMaxAttachID() As Double
Dim Retval As Double
tmp = "Select max(TxnID) from Txn_SecurityAgreeAttachment"
Call TmpTableSMAttachment
Retval = IIf(IsNull(TmpRsSM_Attachment.Fields(0)), 0, TmpRsSM_Attachment.Fields(0))
Retval = Retval + 1
GetMaxAttachID = Retval
End Function
