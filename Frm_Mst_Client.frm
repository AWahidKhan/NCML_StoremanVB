VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmMst_Client 
   Caption         =   "Client Master"
   ClientHeight    =   8280
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12765
   MDIChild        =   -1  'True
   ScaleHeight     =   8280
   ScaleWidth      =   12765
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   3285
      Left            =   120
      TabIndex        =   1
      Top             =   9480
      Visible         =   0   'False
      Width           =   14880
      Begin VB.OptionButton Option3 
         Caption         =   "Hr-Pro CRF ID"
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
         Height          =   375
         Left            =   4800
         TabIndex        =   115
         Top             =   180
         Width           =   1590
      End
      Begin VB.ComboBox CmbSCorpClient 
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
         ItemData        =   "Frm_Mst_Client.frx":0000
         Left            =   7050
         List            =   "Frm_Mst_Client.frx":000A
         Sorted          =   -1  'True
         TabIndex        =   107
         Top             =   1200
         Width           =   975
      End
      Begin VB.ComboBox CmbSRiskType 
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
         ItemData        =   "Frm_Mst_Client.frx":0017
         Left            =   13575
         List            =   "Frm_Mst_Client.frx":0024
         Sorted          =   -1  'True
         TabIndex        =   105
         Top             =   675
         Width           =   1185
      End
      Begin VB.ComboBox CmbSKYCFlag 
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
         ItemData        =   "Frm_Mst_Client.frx":003B
         Left            =   4920
         List            =   "Frm_Mst_Client.frx":0045
         Sorted          =   -1  'True
         TabIndex        =   100
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox TxtSCMSPCode 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1545
         MaxLength       =   200
         TabIndex        =   98
         Top             =   1200
         Width           =   2625
      End
      Begin VB.ComboBox CmbSSCMClient 
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
         ItemData        =   "Frm_Mst_Client.frx":0052
         Left            =   12240
         List            =   "Frm_Mst_Client.frx":005C
         Sorted          =   -1  'True
         TabIndex        =   94
         Top             =   660
         Width           =   1275
      End
      Begin VB.ComboBox CmbSConverted 
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
         ItemData        =   "Frm_Mst_Client.frx":0069
         Left            =   10965
         List            =   "Frm_Mst_Client.frx":0073
         Sorted          =   -1  'True
         TabIndex        =   90
         Top             =   675
         Width           =   1245
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
         ItemData        =   "Frm_Mst_Client.frx":0080
         Left            =   9645
         List            =   "Frm_Mst_Client.frx":008A
         Sorted          =   -1  'True
         TabIndex        =   87
         Top             =   675
         Width           =   1275
      End
      Begin VB.ComboBox CmbSReturn 
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
         ItemData        =   "Frm_Mst_Client.frx":00A0
         Left            =   8595
         List            =   "Frm_Mst_Client.frx":00AA
         Sorted          =   -1  'True
         TabIndex        =   82
         Top             =   675
         Width           =   1005
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Navcode"
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
         Left            =   3480
         TabIndex        =   77
         Top             =   233
         Width           =   1230
      End
      Begin VB.ComboBox CmbSearchExchange 
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
         ItemData        =   "Frm_Mst_Client.frx":00B7
         Left            =   6000
         List            =   "Frm_Mst_Client.frx":00B9
         Sorted          =   -1  'True
         TabIndex        =   69
         Top             =   675
         Width           =   2550
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
         TabIndex        =   59
         Top             =   233
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Client ID"
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
         Left            =   720
         TabIndex        =   58
         Top             =   233
         Width           =   1110
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Client Name"
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
         Left            =   1950
         TabIndex        =   57
         Top             =   233
         Width           =   1470
      End
      Begin VB.CommandButton CmdGo 
         Caption         =   "Go"
         Height          =   345
         Left            =   12705
         TabIndex        =   56
         Top             =   1200
         Width           =   1005
      End
      Begin VB.TextBox TxtSearchClient 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1545
         MaxLength       =   200
         TabIndex        =   55
         Top             =   675
         Width           =   4350
      End
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   345
         Left            =   13710
         TabIndex        =   33
         Top             =   1208
         Width           =   1005
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1215
         Left            =   75
         TabIndex        =   32
         Top             =   1665
         Width           =   14700
         _ExtentX        =   25929
         _ExtentY        =   2143
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label28 
         Caption         =   "Corporate Client"
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
         Left            =   6000
         TabIndex        =   108
         Top             =   1140
         Width           =   945
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   13680
         TabIndex        =   104
         Top             =   240
         Width           =   930
      End
      Begin VB.Label Label24 
         Caption         =   "KYC Flag"
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
         Left            =   4350
         TabIndex        =   101
         Top             =   1133
         Width           =   705
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "CMS-P Code"
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
         TabIndex        =   99
         Top             =   1260
         Width           =   1170
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "SCM Client"
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
         Left            =   12360
         TabIndex        =   95
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Converted"
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
         Left            =   11040
         TabIndex        =   91
         Top             =   240
         Width           =   1020
      End
      Begin VB.Label Label14 
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
         Left            =   10080
         TabIndex        =   88
         Top             =   240
         Width           =   555
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "Return"
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
         Left            =   8850
         TabIndex        =   83
         Top             =   240
         Width           =   585
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Exchange Type"
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
         Left            =   6615
         TabIndex        =   70
         Top             =   240
         Width           =   1425
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
         Left            =   8550
         TabIndex        =   61
         Top             =   1260
         Width           =   2415
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "--"
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
         Left            =   11070
         TabIndex        =   60
         Top             =   1260
         Width           =   165
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Client ID "
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
         TabIndex        =   54
         Top             =   720
         Width           =   780
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "KYD Attachments"
      Height          =   2295
      Left            =   10440
      TabIndex        =   117
      Top             =   120
      Width           =   9135
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
         Height          =   990
         Left            =   4920
         MaxLength       =   5000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   121
         Top             =   495
         Width           =   4035
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
         Left            =   6540
         Style           =   1  'Graphical
         TabIndex        =   120
         Top             =   1860
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
         Left            =   4470
         Style           =   1  'Graphical
         TabIndex        =   119
         Top             =   1860
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
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   118
         Top             =   1860
         Width           =   2070
      End
      Begin VB.Label Label34 
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
         Left            =   4845
         TabIndex        =   123
         Top             =   210
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
         Height          =   1005
         Left            =   120
         TabIndex        =   122
         Top             =   480
         Width           =   4680
      End
   End
   Begin VB.Frame Frame1 
      Height          =   8205
      Left            =   120
      TabIndex        =   34
      Top             =   120
      Width           =   10215
      Begin VB.TextBox TxtTANNo 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   6600
         MaxLength       =   10
         TabIndex        =   21
         Top             =   6480
         Width           =   3510
      End
      Begin VB.TextBox TxtPAN 
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
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   112
         TabStop         =   0   'False
         Top             =   6000
         Width           =   3510
      End
      Begin VB.TextBox TxtFirmName 
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
         Left            =   2175
         Locked          =   -1  'True
         TabIndex        =   111
         TabStop         =   0   'False
         Top             =   6900
         Width           =   7920
      End
      Begin VB.CheckBox ChkCorpClient 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   9330
         TabIndex        =   18
         Top             =   4755
         Width           =   225
      End
      Begin VB.ComboBox CmbRiskType 
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
         ItemData        =   "Frm_Mst_Client.frx":00BB
         Left            =   7680
         List            =   "Frm_Mst_Client.frx":00C8
         Sorted          =   -1  'True
         TabIndex        =   14
         Top             =   4320
         Width           =   1890
      End
      Begin VB.CheckBox ChkKYCFlag 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5910
         TabIndex        =   16
         Top             =   4755
         Width           =   255
      End
      Begin VB.TextBox TxtCMSPCode 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   15
         Top             =   4695
         Width           =   2445
      End
      Begin VB.CheckBox ChkSCMFlag 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5910
         TabIndex        =   13
         Top             =   4365
         Width           =   255
      End
      Begin VB.CheckBox ChkConverted 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3840
         TabIndex        =   12
         Top             =   4365
         Width           =   255
      End
      Begin VB.CheckBox ChkReturn 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   7560
         TabIndex        =   17
         Top             =   4755
         Width           =   225
      End
      Begin VB.TextBox TxtNAVClientName 
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
         Left            =   2175
         Locked          =   -1  'True
         TabIndex        =   80
         TabStop         =   0   'False
         Top             =   5565
         Width           =   7935
      End
      Begin VB.TextBox TxtNAVUpdated 
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
         Left            =   2175
         Locked          =   -1  'True
         TabIndex        =   78
         TabStop         =   0   'False
         Top             =   7740
         Width           =   7935
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
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   7320
         Width           =   3495
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
         Left            =   2175
         Locked          =   -1  'True
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   7335
         Width           =   3450
      End
      Begin VB.CommandButton CmdExchangeTypeID 
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
         Left            =   9720
         TabIndex        =   67
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton CmdDepositoryID 
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
         Left            =   9720
         TabIndex        =   63
         Top             =   1560
         Width           =   375
      End
      Begin VB.ComboBox CmbDepositoryID 
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
         Left            =   2160
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   1560
         Width           =   7440
      End
      Begin VB.TextBox TxtClientID 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   3
         Top             =   720
         Width           =   2445
      End
      Begin VB.TextBox TxtClientIDRow 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   360
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   240
         Width           =   2445
      End
      Begin VB.TextBox TxtDepositorID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   360
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   40
         Top             =   8280
         Visible         =   0   'False
         Width           =   3000
      End
      Begin VB.TextBox TxtDepositorName 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   2160
         MaxLength       =   200
         TabIndex        =   4
         Top             =   1125
         Width           =   7440
      End
      Begin VB.TextBox TxtAddress 
         Appearance      =   0  'Flat
         Height          =   1005
         Left            =   2160
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   1980
         Width           =   7440
      End
      Begin VB.TextBox TxtOrganization 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   2160
         MaxLength       =   255
         TabIndex        =   7
         Top             =   3060
         Width           =   7440
      End
      Begin VB.TextBox TxtTelephoneNo 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   10
         Top             =   3900
         Width           =   2445
      End
      Begin VB.TextBox TxtDepositoryID 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   200
         TabIndex        =   66
         Top             =   1560
         Visible         =   0   'False
         Width           =   7440
      End
      Begin VB.TextBox TxtClientGroupID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   360
         Left            =   6645
         Locked          =   -1  'True
         TabIndex        =   76
         TabStop         =   0   'False
         Top             =   705
         Visible         =   0   'False
         Width           =   2940
      End
      Begin VB.ComboBox CmbExchange 
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
         ItemData        =   "Frm_Mst_Client.frx":00DF
         Left            =   6630
         List            =   "Frm_Mst_Client.frx":00E1
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   2970
      End
      Begin VB.TextBox TxtEmailID 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   2160
         MaxLength       =   200
         TabIndex        =   8
         Top             =   3510
         Width           =   2445
      End
      Begin VB.TextBox TxtMobileNo 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   6630
         MaxLength       =   50
         TabIndex        =   11
         Top             =   3900
         Width           =   2940
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
         ItemData        =   "Frm_Mst_Client.frx":00E3
         Left            =   6630
         List            =   "Frm_Mst_Client.frx":00ED
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   3510
         Width           =   2970
      End
      Begin VB.TextBox TxtStatus 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   6630
         Locked          =   -1  'True
         TabIndex        =   85
         Top             =   3510
         Visible         =   0   'False
         Width           =   2940
      End
      Begin VB.TextBox TxtExchangeTypeID 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   6630
         Locked          =   -1  'True
         TabIndex        =   68
         Top             =   240
         Visible         =   0   'False
         Width           =   2940
      End
      Begin VB.TextBox TxtRiskType 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   7680
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   103
         Top             =   4320
         Width           =   1890
      End
      Begin VB.ComboBox CmbConvertedClientName 
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
         Left            =   2175
         Sorted          =   -1  'True
         TabIndex        =   19
         Top             =   5160
         Visible         =   0   'False
         Width           =   7935
      End
      Begin VB.ComboBox CmbCRFID 
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
         Left            =   2160
         Sorted          =   -1  'True
         TabIndex        =   20
         Top             =   6000
         Width           =   2415
      End
      Begin VB.TextBox TxtCRFID 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   110
         Top             =   6000
         Width           =   2415
      End
      Begin VB.TextBox TxtConvertedClientName 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2175
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   22
         Top             =   5160
         Visible         =   0   'False
         Width           =   7920
      End
      Begin VB.Label Label33 
         Caption         =   "TAN Number"
         Height          =   225
         Left            =   4920
         TabIndex        =   116
         Top             =   6548
         Width           =   1155
      End
      Begin VB.Label Label32 
         Caption         =   "PAN Number"
         Height          =   225
         Left            =   4920
         TabIndex        =   114
         Top             =   6068
         Width           =   1035
      End
      Begin VB.Label Label30 
         Caption         =   "CRF Firm Name"
         Height          =   225
         Left            =   240
         TabIndex        =   113
         Top             =   6960
         Width           =   1155
      End
      Begin VB.Label Label29 
         Caption         =   "HR-Pro CRF ID"
         Height          =   225
         Left            =   240
         TabIndex        =   109
         Top             =   6068
         Width           =   1875
      End
      Begin VB.Label Label27 
         Caption         =   "Corporate Client"
         Height          =   210
         Left            =   8040
         TabIndex        =   106
         Top             =   4770
         Width           =   1245
      End
      Begin VB.Label Label25 
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
         Height          =   210
         Left            =   6720
         TabIndex        =   102
         Top             =   4380
         Width           =   915
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "KYC Flag"
         Height          =   195
         Left            =   4920
         TabIndex        =   97
         Top             =   4785
         Width           =   660
      End
      Begin VB.Label Label19 
         Caption         =   "CMS-P Code"
         Height          =   240
         Left            =   240
         TabIndex        =   96
         Top             =   4755
         Width           =   1800
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "SCM Client"
         Height          =   195
         Left            =   4920
         TabIndex        =   93
         Top             =   4395
         Width           =   780
      End
      Begin VB.Label LblConvertedClientName 
         Caption         =   "Converted Client Name"
         Height          =   225
         Left            =   240
         TabIndex        =   92
         Top             =   5205
         Visible         =   0   'False
         Width           =   1875
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Converted"
         Height          =   195
         Left            =   2160
         TabIndex        =   89
         Top             =   4395
         Width           =   735
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Cheque Return"
         Height          =   195
         Left            =   6360
         TabIndex        =   84
         Top             =   4785
         Width           =   1080
      End
      Begin VB.Label Label11 
         Caption         =   "NAVISION Client Name"
         Height          =   240
         Left            =   240
         TabIndex        =   81
         Top             =   5625
         Width           =   1695
      End
      Begin VB.Label Label8 
         Caption         =   "NAV Updated"
         Height          =   255
         Left            =   240
         TabIndex        =   79
         Top             =   7800
         Width           =   1695
      End
      Begin VB.Label Label7 
         Caption         =   "Client Group"
         Height          =   240
         Left            =   4935
         TabIndex        =   75
         Top             =   750
         Width           =   1395
      End
      Begin VB.Label Label9 
         Caption         =   "Updated :"
         Height          =   255
         Left            =   5760
         TabIndex        =   74
         Top             =   7395
         Width           =   1695
      End
      Begin VB.Label Label10 
         Caption         =   "Created  :"
         Height          =   240
         Left            =   240
         TabIndex        =   73
         Top             =   7395
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "ClientID"
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
         TabIndex        =   65
         Top             =   810
         Width           =   1800
      End
      Begin VB.Label Label3 
         Caption         =   "Depository Name"
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
         TabIndex        =   64
         Top             =   1635
         Width           =   1800
      End
      Begin VB.Label Label5 
         Caption         =   "Client Code"
         Height          =   225
         Left            =   240
         TabIndex        =   53
         Top             =   308
         Width           =   1800
      End
      Begin VB.Label LblClientID 
         Caption         =   "DepositorID"
         Height          =   300
         Left            =   240
         TabIndex        =   39
         Top             =   8280
         Visible         =   0   'False
         Width           =   1800
      End
      Begin VB.Label LblClientName 
         Caption         =   "Client Name"
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
         TabIndex        =   41
         Top             =   1215
         Width           =   1800
      End
      Begin VB.Label LblAddress 
         Caption         =   "Address"
         Height          =   300
         Left            =   240
         TabIndex        =   43
         Top             =   2332
         Width           =   1800
      End
      Begin VB.Label LblOrganization 
         Caption         =   "Organization"
         Height          =   225
         Left            =   240
         TabIndex        =   35
         Top             =   3128
         Width           =   1800
      End
      Begin VB.Label LblEmailID 
         Caption         =   "EmailID"
         Height          =   225
         Left            =   240
         TabIndex        =   36
         Top             =   3578
         Width           =   1800
      End
      Begin VB.Label LblTelephoneNo 
         Caption         =   "Telephone No"
         Height          =   240
         Left            =   240
         TabIndex        =   37
         Top             =   3960
         Width           =   1800
      End
      Begin VB.Label LblMobileNo 
         Caption         =   "MobileNo"
         Height          =   210
         Left            =   4920
         TabIndex        =   38
         Top             =   3975
         Width           =   840
      End
      Begin VB.Label Label4 
         Caption         =   "Exchange Type"
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
         Left            =   4920
         TabIndex        =   62
         Top             =   240
         Width           =   1395
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Status"
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
         Left            =   4920
         TabIndex        =   86
         Top             =   3593
         Width           =   555
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1200
      Left            =   120
      TabIndex        =   42
      Top             =   8280
      Width           =   10215
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   400
         Left            =   105
         TabIndex        =   27
         Top             =   645
         Width           =   1995
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   400
         Left            =   2100
         TabIndex        =   28
         Top             =   645
         Width           =   1995
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   400
         Left            =   4095
         TabIndex        =   29
         Top             =   645
         Width           =   1995
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   400
         Left            =   6090
         TabIndex        =   30
         Top             =   645
         Width           =   1995
      End
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   400
         Left            =   8085
         TabIndex        =   31
         Top             =   645
         Width           =   1995
      End
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New"
         Height          =   400
         Left            =   105
         TabIndex        =   0
         Top             =   240
         Width           =   1995
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   400
         Left            =   2100
         TabIndex        =   23
         Top             =   240
         Width           =   1995
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   400
         Left            =   4080
         TabIndex        =   24
         Top             =   240
         Width           =   1995
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   400
         Left            =   6090
         TabIndex        =   25
         Top             =   240
         Width           =   1995
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "Search"
         Height          =   400
         Left            =   8085
         TabIndex        =   26
         Top             =   240
         Width           =   1995
      End
   End
   Begin VB.TextBox TxtFlag 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   360
      Left            =   2640
      MaxLength       =   1
      TabIndex        =   44
      Top             =   4680
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.ComboBox CmbExchangeTypeID 
      Height          =   315
      Left            =   2280
      TabIndex        =   47
      Top             =   1080
      Width           =   3600
   End
   Begin VB.TextBox TxtDepository 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   2400
      MaxLength       =   50
      TabIndex        =   50
      Top             =   600
      Width           =   7440
   End
   Begin VB.ComboBox CmbDepository 
      Height          =   315
      Left            =   2265
      TabIndex        =   49
      Top             =   1080
      Width           =   7455
   End
   Begin VB.CommandButton CmdDepository 
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
      Left            =   9960
      TabIndex        =   48
      Top             =   1080
      Width           =   375
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid5 
      Height          =   2670
      Left            =   10440
      TabIndex        =   124
      Top             =   2400
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   4710
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
   Begin MSComDlg.CommonDialog cdialog 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label LblDepositoryID 
      Caption         =   "DepositoryName"
      Height          =   300
      Left            =   360
      TabIndex        =   51
      Top             =   1080
      Width           =   1800
   End
   Begin VB.Label LblExchangeTypeID 
      Caption         =   "Exchange Type"
      Height          =   300
      Left            =   360
      TabIndex        =   46
      Top             =   1080
      Width           =   1800
   End
   Begin VB.Label LblFlag 
      Caption         =   "Flag"
      Enabled         =   0   'False
      Height          =   300
      Left            =   720
      TabIndex        =   45
      Top             =   5280
      Visible         =   0   'False
      Width           =   600
   End
End
Attribute VB_Name = "FrmMst_Client"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg, I, InmClientRowID As Variant
Dim mExchangeID, mDepositoryID, mDepositorID, tmpExchangeID As Double
Dim tmpDepositorID As String
Dim mExchangeTypeID As Double
Dim mClientGroupID As Double
Dim mConvertedClientIDRow As Double
Dim mCRFID As Variant
Dim File_Path As String

Dim oFs As New FileSystemObject
Dim oFs1 As New FileSystemObject
Dim oFolder As Folder
Dim oFile As File

Private Sub CmbConvertedClientName_GotFocus()
tmp1 = CmbConvertedClientName.Text
'mExchangeTypeID = GetExchangeID("NON NCDEX")
'Call TableMst_Client("where ExchangeTypeId = " & mExchangeTypeID & " And Converted = 1 ")
tmp = "Select * from Mst_Client Where ExchangeTypeId = 1 And Converted = 1 "
Call TmpTable

CmbConvertedClientName.Clear
If TmpRs.RecordCount = 0 Then
   MsgBox "No record found!!! "
   Exit Sub
End If
For I = 1 To TmpRs.RecordCount
    CmbConvertedClientName.AddItem TmpRs!ClientName
    TmpRs.MoveNext
Next
CmbConvertedClientName.Text = tmp1
End Sub
Private Sub CmbConvertedClientName_LostFocus()
If CmbConvertedClientName.Text = "" Then
   Exit Sub
End If
TmpRs.MoveFirst
TmpRs.Find "ClientName = '" & CmbConvertedClientName.Text & "'"
If TmpRs.EOF Then
   MsgBox "Invalid selection!!! "
   CmbConvertedClientName.SetFocus
   Exit Sub
End If
mConvertedClientIDRow = TmpRs!ClientIDRow

End Sub



Private Sub CmbCRFID_GotFocus()
tmp1 = CmbCRFID.Text
tmp = "Select CRFID,PAN,FirmName From Txn_CRF Where Flag='Y'"
Call TmpTableHRPro
CmbCRFID.Clear
If TmpRsHRPro.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To TmpRsHRPro.RecordCount
    CmbCRFID.AddItem TmpRsHRPro!CRFID
    TmpRsHRPro.MoveNext
Next
CmbCRFID.Text = tmp1
End Sub

Private Sub CmbCRFID_LostFocus()
If CmbCRFID.Text = "" Then
   mCRFID = Null
   Exit Sub
End If

If TmpRsHRPro.RecordCount > 0 Then TmpRsHRPro.MoveFirst
TmpRsHRPro.Find "CRFID = " & Val(CmbCRFID.Text) & ""
If TmpRsHRPro.EOF Then
   MsgBox "Invalid selection "
   mCRFID = Null
   CmbCRFID.SetFocus
   Exit Sub
End If
mCRFID = TmpRsHRPro!CRFID
TxtPAN.Text = IIf(IsNull(TmpRsHRPro!PAN), "", TmpRsHRPro!PAN)
TxtFirmName.Text = TmpRsHRPro!FirmName
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
    'If RsMst_ExchangeType.Fields(0) <> 1 Then
    CmbExchange.AddItem RsMst_ExchangeType!ExchangeType
    'End If
    RsMst_ExchangeType.MoveNext
Next
CmbExchange.Text = tmp
End Sub

Private Sub CmbExchange_LostFocus()
If CmbExchange.Text = "" Then
   ChkConverted.Enabled = False
   ChkSCMFlag.Enabled = False
   ChkCorpClient.Enabled = False
   Exit Sub
End If

RsMst_ExchangeType.MoveFirst
RsMst_ExchangeType.Find "ExchangeType = '" & CmbExchange.Text & "'"
If RsMst_ExchangeType.EOF Then
   MsgBox "Invalid selection "
   CmbExchange.SetFocus
   Exit Sub
End If
If CmbExchange.Text = "SPOT" Then
  Label3.Caption = "TM Name"
Else
  Label3.Caption = "Depository Name"
End If
mExchangeTypeID = RsMst_ExchangeType!ExchangeTypeID
CmbDepositoryID.Text = ""
Frame2.Enabled = False
If mExchangeTypeID = 1 Then
   ChkConverted.Enabled = True
   ChkSCMFlag.Enabled = True
   ChkCorpClient.Enabled = True
   Call Grid_Head4
Else
   ChkConverted.Enabled = False
   ChkSCMFlag.Enabled = False
   ChkCorpClient.Enabled = False

   Frame2.Enabled = True
   lblAttachment.Enabled = True
   TxtAttachmentRemarks.Locked = False
   CmdAttachment.Enabled = True
   CmdGetFile.Enabled = True
   CmdRemoveAttachment.Enabled = True
   Call Grid_Head4
End If

End Sub


Private Sub CmbClientGroup_GotFocus()
tmp = CmbClientGroup.Text
Call TableMst_ClientGroup
CmbClientGroup.Clear
If RsMst_ClientGroup.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_ClientGroup.RecordCount
    'If RsMst_ExchangeType.Fields(0) <> 1 Then
    CmbClientGroup.AddItem RsMst_ClientGroup!ClientGroupName
    'End If
    RsMst_ClientGroup.MoveNext
Next
CmbClientGroup.Text = tmp

End Sub

Private Sub CmbClientGroup_LostFocus()
If CmbClientGroup.Text = "" Then
   Exit Sub
End If

RsMst_ClientGroup.MoveFirst
RsMst_ClientGroup.Find "ClientGroupName = '" & CmbClientGroup.Text & "'"
If RsMst_ClientGroup.EOF Then
   MsgBox "Invalid selection "
   CmbClientGroup.SetFocus
   Exit Sub
End If

mClientGroupID = RsMst_ClientGroup!ClientGroupID

End Sub

Private Sub CmdClientGroup_Click()
FrmMst_ClientGroup.Show
End Sub



Private Sub CmbRiskType_LostFocus()
If CmbRiskType.Text = "" Then Exit Sub
If LCase(CmbRiskType.Text) <> "low" And LCase(CmbRiskType.Text) <> "high" And LCase(CmbRiskType.Text) <> "medium" Then
   MsgBox "Invalid selection!!! "
   CmbRiskType.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmbSConverted_LostFocus()
If CmbSConverted.Text = "" Then Exit Sub
If LCase(CmbSConverted.Text) <> "yes" And LCase(CmbSConverted.Text) <> "no" Then
   MsgBox "Invalid Selection"
   CmbSConverted.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmbSCorpClient_LostFocus()

If CmbSCorpClient.Text = "" Then Exit Sub
If LCase(CmbSCorpClient.Text) <> "yes" And LCase(CmbSCorpClient.Text) <> "no" Then
   MsgBox "Invalid Selection!!!"
   CmbSCorpClient.SetFocus
   Exit Sub
End If

End Sub

Private Sub CmbSKYCFlag_LostFocus()

If CmbSKYCFlag.Text = "" Then Exit Sub
If LCase(CmbSKYCFlag.Text) <> "yes" And LCase(CmbSKYCFlag.Text) <> "no" Then
   MsgBox "Invalid Selection!!!"
   CmbSKYCFlag.SetFocus
   Exit Sub
End If

End Sub

Private Sub CmbSReturn_LostFocus()
If CmbSReturn.Text = "" Then Exit Sub
If LCase(CmbSReturn.Text) <> "yes" And LCase(CmbSReturn.Text) <> "no" Then
   MsgBox "Invalid Selection"
   CmbSReturn.SetFocus
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


Private Sub CmbSSCMClient_LostFocus()
If CmbSSCMClient.Text = "" Then Exit Sub
If LCase(CmbSSCMClient.Text) <> "yes" And LCase(CmbSSCMClient.Text) <> "no" Then
   MsgBox "Invalid Selection!!!"
   CmbSSCMClient.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmbSStatus_LostFocus()
If CmbSStatus.Text = "" Then Exit Sub
If LCase(CmbSStatus.Text) <> "active" And LCase(CmbSStatus.Text) <> "deactive" Then
   MsgBox "Invalid Selection!!!"
   CmbSStatus.SetFocus
   Exit Sub
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

Private Sub CmdGo_Click()

tmp = "Select MC.ClientID,MC.DepositorID,MC.ClientName,ME.ExchangeType,MCG.ClientGroupName,"
tmp = tmp & " MD.DepositoryID + '~' + MD.DepositoryName As [DP Name],"
tmp = tmp & " MC.Flag , MC.Address, MC.Organization, MC.EmailID, MC.TelephoneNo, MC.MobileNo,"
tmp = tmp & " MC.NAVCode , MC.RNAVCode, MC.ClientIDRow,MC.FAGUpdated,MC.FAGUpdatedDate ,MC.NAVClientName, MC.ChequeReturn, MC.Status,MC.Converted,MC.SCMFlag,MCL.CLientName 'ConvertedClientName',MC.CMSPCode,MC.KYCFlag, MC.RiskType,MC.CorpClient,MC.CRFID,'' 'CRFFirmName',MC.TANNO "
tmp = tmp & " from Mst_Client MC "
tmp = tmp & " Inner Join Mst_ExchangeType ME on MC.ExchangeTypeID = ME.ExchangeTypeID"
tmp = tmp & " Inner Join Mst_Depository MD on MC.DepositoryID = MD.DepositoryID"
tmp = tmp & " Left Join  Mst_ClientGroup MCG on MC.ClientGroupID = MCG.ClientGroupID"
tmp = tmp & " Left Join  Mst_Client MCl on MC.ConvertedClientIDRow = MCl.ClientIDRow"

Dim wc As Variant
wc = ""
If Option5.Value = True Then
   wc = " Where MC.ClientName Like '%" & TxtSearchClient & "%' "
End If
If Option2.Value = True Then
   wc = " Where MC.Navcode Like '%" & TxtSearchClient & "%' or MC.RNavcode Like '%" & TxtSearchClient & "%'"
End If

If Option1.Value = True Then
   wc = " Where MC.ClientID  Like '%" & TxtSearchClient & "%'"
End If

If Option3.Value = True Then
   wc = " Where MC.CRFID  Like '%" & TxtSearchClient & "%'"
End If

If CmbSearchExchange.Text <> "" Then
   If wc = "" Then
      wc = " Where MC.ExchangeTypeID = " & mExchangeID & " "
   Else
      wc = wc & " And MC.ExchangeTypeID = " & mExchangeID & " "
   End If
End If

If LCase(CmbSReturn.Text) = "yes" Then
   If wc <> "" Then
      wc = wc & " AND MC.ChequeReturn='Y'"
   Else
      wc = " Where MC.ChequeReturn='Y'"
   End If
ElseIf LCase(CmbSReturn.Text) = "no" Then
   If wc <> "" Then
      wc = wc & " AND MC.ChequeReturn='N'"
   Else
      wc = " Where MC.ChequeReturn='N'"
   End If
End If

If CmbSStatus.Text <> "" Then
   If wc = "" Then
      wc = " Where MC.Status = '" & Left$(CmbSStatus.Text, 1) & "'"
   Else
      wc = wc & " And MC.Status = '" & Left$(CmbSStatus.Text, 1) & "'"
   End If
End If

If CmbSRiskType.Text <> "" Then
   If wc = "" Then
      wc = " Where MC.RiskType = '" & CmbSRiskType & "'"
   Else
      wc = wc & " And MC.RiskType = '" & CmbSRiskType & "'"
   End If
End If

If LCase(CmbSConverted.Text) = "yes" Then
   If wc <> "" Then
      wc = wc & " AND MC.Converted= 1 "
   Else
      wc = " Where MC.Converted= 1 "
   End If
ElseIf LCase(CmbSConverted.Text) = "no" Then
   If wc <> "" Then
      wc = wc & " AND MC.Converted = 0 "
   Else
      wc = " Where MC.Converted = 0 "
   End If
End If

If LCase(CmbSSCMClient.Text) = "yes" Then
   If wc <> "" Then
      wc = wc & " AND MC.SCMFlag = 1 "
   Else
      wc = " Where MC.SCMFlag= 1 "
   End If
ElseIf LCase(CmbSSCMClient.Text) = "no" Then
   If wc <> "" Then
      wc = wc & " AND MC.SCMFlag = 0 "
   Else
      wc = " Where MC.SCMFlag = 0 "
   End If
End If

If TxtSCMSPCode <> "" Then
   If wc <> "" Then
      wc = wc & " AND MC.CMSPCode like  '%" & TxtSCMSPCode & "%' "
   Else
      wc = " Where MC.CMSPCode like  '%" & TxtSCMSPCode & "%' "
   End If
End If

If LCase(CmbSKYCFlag.Text) = "yes" Then
   If wc <> "" Then
      wc = wc & " AND MC.KYCFlag = 1 "
   Else
      wc = " Where MC.KYCFlag= 1 "
   End If
ElseIf LCase(CmbSKYCFlag.Text) = "no" Then
   If wc <> "" Then
      wc = wc & " AND MC.KYCFlag = 0 "
   Else
      wc = " Where MC.KYCFlag = 0 "
   End If
End If

If LCase(CmbSCorpClient.Text) = "yes" Then
   If wc <> "" Then
      wc = wc & " AND MC.CorpClient = 1 "
   Else
      wc = " Where MC.CorpClient= 1 "
   End If
ElseIf LCase(CmbSCorpClient.Text) = "no" Then
   If wc <> "" Then
      wc = wc & " AND MC.CorpClient = 0 "
   Else
      wc = " Where MC.CorpClient = 0 "
   End If
End If
tmp = tmp & " " & wc & " Order by MC.ExchangeTypeID  ,MC.ClientName"
Call TmpTable

Call Grid_Head

For iu = 1 To TmpRs.RecordCount
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
    For C = 0 To MSHFlexGrid1.Cols - 1
       MSHFlexGrid1.TextMatrix(iu, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
    Next
    
    If Val(MSHFlexGrid1.TextMatrix(iu, 27)) > 0 Then
       MSHFlexGrid1.TextMatrix(iu, 28) = GetCRFFirmName(Val(MSHFlexGrid1.TextMatrix(iu, 27)))
    End If
    
    If LCase(MSHFlexGrid1.TextMatrix(iu, 18)) = "y" Then
       MSHFlexGrid1.TextMatrix(iu, 18) = "Yes"
    Else
       MSHFlexGrid1.TextMatrix(iu, 18) = "No"
    End If
    If LCase(MSHFlexGrid1.TextMatrix(iu, 19)) = "a" Then
       MSHFlexGrid1.TextMatrix(iu, 19) = "Active"
    Else
       MSHFlexGrid1.TextMatrix(iu, 19) = "Deactive"
    End If
    If LCase(MSHFlexGrid1.TextMatrix(iu, 20)) = "1" Then
       MSHFlexGrid1.TextMatrix(iu, 20) = "Yes"
    Else
       MSHFlexGrid1.TextMatrix(iu, 20) = "No"
    End If
    If LCase(MSHFlexGrid1.TextMatrix(iu, 21)) = "1" Then
       MSHFlexGrid1.TextMatrix(iu, 21) = "Yes"
    Else
       MSHFlexGrid1.TextMatrix(iu, 21) = "No"
    End If
    If LCase(MSHFlexGrid1.TextMatrix(iu, 26)) = "1" Then
       MSHFlexGrid1.TextMatrix(iu, 26) = "Yes"
    Else
       MSHFlexGrid1.TextMatrix(iu, 26) = "No"
    End If

    TmpRs.MoveNext
Next

Label20.Caption = "Total Records Found : " & TmpRs.RecordCount
Label20.Width = 3000
End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Call GButtonLockAD(Me)
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If
Option4.Value = True
Call Grid_Head
Call TableMst_Client(" Order By ClientIDRow ")
If RsMst_Client.RecordCount = 0 Then
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
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Mst_Client.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Mst_Client.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = I + 1 ' + 6
    For C = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(I, C), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(I, C)) Then
              oWS.Cells(mRow, C + 1) = Format(MSHFlexGrid1.TextMatrix(I, C), "dd-MMM-yyyy") ' "MM/dd/yyyy")
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
I = MSHFlexGrid2.RowSel
If I <= 0 Then Exit Sub
   CmbExchange.Enabled = True
'   If MSHFlexGrid2.TextMatrix(i, 8) <> "" Then
'      CmbExchange.Enabled = False
'   End If
   
   TxtClientID.Text = MSHFlexGrid2.TextMatrix(I, 2)
   TxtClientName.Text = MSHFlexGrid2.TextMatrix(I, 3)
   CmbExchange.Text = MSHFlexGrid2.TextMatrix(I, 1)
   mExchangeTypeID = GetExchangeID(CmbExchange.Text)
   CmbDepositoryID.Text = MSHFlexGrid2.TextMatrix(I, 4)
   mDepositoryID = GetDepositoryID(CmbDepositoryID.Text)
   CmdAdd2List.Caption = "Update List"
   
End Sub



Private Sub MSHFlexGrid5_Click()
Dim ans, r As Variant
Dim ID As Integer
r = MSHFlexGrid5.RowSel
If r < 1 Then Exit Sub
If MSHFlexGrid5.TextMatrix(r, 1) = "" Then Exit Sub

If SaveCmd.Enabled = False Then
'Fumigation And Spraying
   ans = MsgBox("Do You want to view this document ", vbYesNo)
   If ans = vbYes Then
      If Right(App.Path, 1) = "\" Then
         Shell "Explorer.exe " & App.Path & "Agreement Attachment\Client KYD\" & Val(TxtClientIDRow) & "", vbMaximizedFocus
      Else
         Shell "Explorer.exe " & App.Path & "\Agreement Attachment\Client KYD\" & Val(TxtClientIDRow) & "\", vbMaximizedFocus
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
RsMst_Client.MoveNext
If RsMst_Client.EOF Then
   RsMst_Client.MoveLast
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

Private Sub Option5_Click()
Call Option1_Click
End Sub

Private Sub PreviousCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_Client.MovePrevious
If RsMst_Client.BOF Then
   RsMst_Client.MoveFirst
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
RsMst_Client.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsMst_Client.MoveLast
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

Dim x As Integer
If SearchCmd.Caption = "Cancel" Then
   If RsMst_Client.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   Call FirstCmd_Click
   Exit Sub
End If
Frame0.Height = Me.Height - 900
Frame0.Left = Frame1.Left - 100
Frame0.Top = Frame1.Top - 100
Frame0.Width = Me.Width - 150
Frame0.Visible = True

MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 1800
''===========comment on 7 july 2012 for not clearing search grid=============
'Option4.Value = True
'Call Grid_Head
''===========comment on 7 july 2012 for not clearing search grid=============

End Sub

Private Sub TxtCMSPCode_LostFocus()
If TxtCMSPCode = "" Then Exit Sub
If SaveCmd.Enabled = False Then Exit Sub
If TxtCMSPCode.Locked = True Then Exit Sub

tmp = "Select * from Mst_Client Where CMSPCode = '" & TxtCMSPCode & "'"
Call TmpTable

If TmpRs.RecordCount > 0 Then
   If TxtClientIDRow <> TmpRs!ClientIDRow Then
      MsgBox "CMS-P Code already mapped to client " & TmpRs!ClientName & "~" & TmpRs!ClientID
      TxtCMSPCode.SetFocus
      Exit Sub
   End If
End If

End Sub

Private Sub TxtSearchClient_LostFocus()
'Dim x As Integer
'If TxtSearchClient = "" Then
'   Call Grid_Head
'   Label21.Caption = ""
'   Exit Sub
'End If
'tmp = "Select ClientID,DepositorID,ClientName,ExchangeTypeID,DepositoryID,Flag,Address,Organization,EmailID,TelephoneNo,MobileNo,ClientIDRow from Mst_Client Where"
'If Option1.Value = True Then
'   tmp = tmp & " ClientID  Like '%" & TxtSearchClient & "%'"  '
''ElseIf Option2.Value = True Then
''   tmp = "Select * from Mst_Client MC Inner Join  Mst_ExchangeType ME on MC.ExchangeTypeID = ME.ExchangeTypeID where ME.ExchangeType like '%" & TxtSearchClient & "%'"
'   'tmp = tmp & " DepositoryID Like '%" & TxtSearchClient & "%' "
'ElseIf Option5.Value = True Then
'   Wc = Wc & " ClientName Like '%" & TxtSearchClient & "%' "
'End If
'Call Tmp1Table
'If TmpRs1.RecordCount = 0 Then
'   MsgBox "No such Record Found!!!"
'   Exit Sub
'End If
'MSHFlexGrid1.Rows = TmpRs1.RecordCount + 1
'For x = 1 To TmpRs1.RecordCount
'    Call ShowGrid(x)
'    TmpRs1.MoveNext
'Next
End Sub
Private Sub Option1_Click()
'Label6.Visible = True
'TxtSearchClient.Visible = True
'TxtSearchClient = ""
'CmbSearchExchange.Visible = False
'CmbSearchExchange.Text = ""
'CmdGo.Visible = True
'If Option1.Value = True Then
'   Label6.Caption = "By Client ID"
'ElseIf Option5.Value = True Then
'   Label6.Caption = "By Client Name"
'End If
'
'TxtSearchClient.SetFocus
End Sub

Private Sub Option2_Click()
Call Option1_Click
End Sub

Private Sub Option3_Click()
Call Option1_Click
End Sub

Private Sub Option4_Click()
'Dim x As Integer
'Label6.Visible = False
'CmbSearchExchange.Visible = False
'TxtSearchClient.Visible = False
'CmdGo.Visible = False
'TmpRs3.MoveFirst
'Call Grid_Head
'MSHFlexGrid1.Rows = TmpRs3.RecordCount + 1
'For x = 1 To TmpRs3.RecordCount
'    Call ShowGrid(x)
'    TmpRs3.MoveNext
'Next
End Sub

Private Sub ShowGrid(I As Integer)
MSHFlexGrid1.ColWidth(0) = 1000
Label21.Refresh
If Option4.Value = False Then
   Call GridDetails(TmpRs1, I)
   Label21.Caption = TmpRs1.RecordCount
Else
   If TmpRs3.EOF Then Exit Sub
   Call GridDetails(TmpRs3, I)
   Label21.Caption = TmpRs3.RecordCount
   
End If
End Sub
Private Sub GridDetails(RSType As Variant, iu As Integer)
For C = 0 To MSHFlexGrid1.Cols - 1
   MSHFlexGrid1.TextMatrix(iu, C) = IIf(IsNull(RSType.Fields(C)), "", RSType.Fields(C))
Next
MSHFlexGrid1.TextMatrix(iu, 3) = GetExchangeName(MSHFlexGrid1.TextMatrix(iu, 3))
MSHFlexGrid1.TextMatrix(iu, 4) = GetDepositorName_ID(MSHFlexGrid1.TextMatrix(iu, 4), "N", GetExchangeID(MSHFlexGrid1.TextMatrix(iu, 3)))

End Sub

Private Sub Grid_Head()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 30 '29 '28 '26 '23 ' 19 '18 '17 '18
MSHFlexGrid1.TextMatrix(0, 0) = "ClientID"
MSHFlexGrid1.TextMatrix(0, 1) = "DepositorId"
MSHFlexGrid1.TextMatrix(0, 2) = "Client Name"
MSHFlexGrid1.TextMatrix(0, 3) = "ExchangeType"
MSHFlexGrid1.TextMatrix(0, 4) = "Client Group"
MSHFlexGrid1.TextMatrix(0, 5) = "DepositoryID"
MSHFlexGrid1.TextMatrix(0, 6) = "Flag"
MSHFlexGrid1.TextMatrix(0, 7) = "Address"
MSHFlexGrid1.TextMatrix(0, 8) = "Organization"
MSHFlexGrid1.TextMatrix(0, 9) = "Email"
MSHFlexGrid1.TextMatrix(0, 10) = "Tel No"
MSHFlexGrid1.TextMatrix(0, 11) = "Mobile No"
MSHFlexGrid1.TextMatrix(0, 12) = "NAV Code"
MSHFlexGrid1.TextMatrix(0, 13) = "R NAV Code"
MSHFlexGrid1.TextMatrix(0, 14) = "Row ID"
MSHFlexGrid1.TextMatrix(0, 15) = "FAG Updated"
MSHFlexGrid1.TextMatrix(0, 16) = "FAG Updated Date"
MSHFlexGrid1.TextMatrix(0, 17) = "NAV Client Name"
MSHFlexGrid1.TextMatrix(0, 18) = "Return Flag"
MSHFlexGrid1.TextMatrix(0, 19) = "Status"
MSHFlexGrid1.TextMatrix(0, 20) = "Converted"
MSHFlexGrid1.TextMatrix(0, 21) = "SCM FLag"
MSHFlexGrid1.TextMatrix(0, 22) = "Converted Client"
MSHFlexGrid1.TextMatrix(0, 23) = "CMS-P Code"
MSHFlexGrid1.TextMatrix(0, 24) = "KYC Flag"
MSHFlexGrid1.TextMatrix(0, 25) = "Risk Type"
MSHFlexGrid1.TextMatrix(0, 26) = "Corporate Client"
MSHFlexGrid1.TextMatrix(0, 27) = "CRFID"
MSHFlexGrid1.TextMatrix(0, 28) = "CRF Firm Name" 'GetCRMFirmName
MSHFlexGrid1.TextMatrix(0, 29) = "TAN No"

MSHFlexGrid1.ColWidth(0) = 800
MSHFlexGrid1.ColWidth(1) = 0
MSHFlexGrid1.ColWidth(2) = 1800
MSHFlexGrid1.ColWidth(3) = 1400
MSHFlexGrid1.ColWidth(4) = 1600
MSHFlexGrid1.ColWidth(5) = 1200
MSHFlexGrid1.ColWidth(6) = 600
MSHFlexGrid1.ColWidth(7) = 2000
MSHFlexGrid1.ColWidth(8) = 1800
MSHFlexGrid1.ColWidth(9) = 1600
MSHFlexGrid1.ColWidth(10) = 1600
MSHFlexGrid1.ColWidth(11) = 1600
MSHFlexGrid1.ColWidth(14) = 0
MSHFlexGrid1.ColWidth(15) = 1000
MSHFlexGrid1.ColWidth(16) = 1200
MSHFlexGrid1.ColWidth(27) = 1200
MSHFlexGrid1.ColWidth(28) = 1200
MSHFlexGrid1.ColWidth(29) = 1200
'MSHFlexGrid1.ColWidth(17) = 1200

End Sub

Public Sub Indata()

TxtDepositorID.Locked = True
TxtClientID.Locked = True
'TxtClientName.Locked = True
TxtDepositorName.Locked = True
CmbExchange.Visible = False
CmbExchange.Locked = True
CmbStatus.Visible = False
CmbStatus.Locked = True
CmdExchangeTypeID.Visible = False
CmbDepositoryID.Visible = False
CmbDepositoryID.Locked = True
CmdDepositoryID.Visible = False
TxtExchangeTypeID.Visible = True
TxtStatus.Visible = True
TxtDepositoryID.Visible = True
TxtClientGroupID.Visible = True

TxtFlag.Locked = True
TxtAddress.Locked = True
TxtOrganization.Locked = True
TxtEmailId.Locked = True
TxtTelephoneNo.Locked = True
TxtMobileNo.Locked = True
TxtClientGroupID.Visible = True

CmbRiskType.Visible = False
TxtRiskType.Visible = True
 
CmbCRFID.Visible = False
TxtCRFID.Visible = True

TxtClientIDRow = IIf(IsNull(RsMst_Client!ClientIDRow), "", RsMst_Client!ClientIDRow)
InmClientRowID = TxtClientIDRow

TxtClientID = IIf(IsNull(RsMst_Client!ClientID), "", RsMst_Client!ClientID)
'mDepositorID = TxtDepositorID.Text
TxtDepositorName = IIf(IsNull(RsMst_Client!ClientName), "", RsMst_Client!ClientName)

LblConvertedClientName.Visible = False
CmbConvertedClientName.Visible = False
TxtConvertedClientName.Visible = False

mExchangeTypeID = IIf(IsNull(RsMst_Client!ExchangeTypeID), "", RsMst_Client!ExchangeTypeID)
TxtExchangeTypeID = GetExchangeName(RsMst_Client!ExchangeTypeID) '(IsNull(RsMst_Client!ExchangeType), "", RsMst_Client!ExchangeType)


If LCase(IIf(IsNull(RsMst_Client!Status), "", RsMst_Client!Status)) = "a" Then
   TxtStatus.Text = "Active"
ElseIf LCase(IIf(IsNull(RsMst_Client!Status), "", RsMst_Client!Status)) = "d" Then
   TxtStatus.Text = "Deactive"
End If
mClientGroupID = IIf(IsNull(RsMst_Client!ClientGroupID), 0, RsMst_Client!ClientGroupID)
TxtClientGroupID = GetClientGroup(mClientGroupID)
If TxtExchangeTypeID = "SPOT" Then
  Label3.Caption = "TM Name"
Else
  Label3.Caption = "Depository Name"
End If

If LCase(IIf(IsNull(RsMst_Client!ChequeReturn), "", RsMst_Client!ChequeReturn)) = "y" Then
   ChkReturn.Value = vbChecked
ElseIf LCase(IIf(IsNull(RsMst_Client!ChequeReturn), "", RsMst_Client!ChequeReturn)) = "n" Then
   ChkReturn.Value = vbUnchecked
End If

mDepositoryID = IIf(IsNull(RsMst_Client!DepositoryID), "", RsMst_Client!DepositoryID)

'TxtDepositoryID = GetDepositorName(RsMst_Client!DepositoryID, "N") '(IsNull(RsMst_Client!DepositoryName), "", RsMst_Client!DepositoryName)

TxtDepositoryID = GetDepositorName_ID(RsMst_Client!DepositoryID, "N", mExchangeTypeID) '(IsNull(RsMst_Client!DepositoryName), "", RsMst_Client!DepositoryName)

'GetDepositorName_ID

TxtFlag = "0" 'IIf(IsNull(RsMst_Client!Flag), "", RsMst_Client!Flag)
TxtAddress = IIf(IsNull(RsMst_Client!Address), "", RsMst_Client!Address)
TxtOrganization = IIf(IsNull(RsMst_Client!Organization), "", RsMst_Client!Organization)
TxtEmailId = IIf(IsNull(RsMst_Client!EmailID), "", RsMst_Client!EmailID)
TxtTelephoneNo = IIf(IsNull(RsMst_Client!TelephoneNo), "", RsMst_Client!TelephoneNo)
TxtMobileNo = IIf(IsNull(RsMst_Client!MobileNo), "", RsMst_Client!MobileNo)
TxtNAVClientName = IIf(IsNull(RsMst_Client!NAVClientName), "", RsMst_Client!NAVClientName)
TxtRiskType = IIf(IsNull(RsMst_Client!RiskType), "", RsMst_Client!RiskType)

ChkConverted.Enabled = False
ChkConverted.Value = IIf(IsNull(RsMst_Client!Converted), 0, RsMst_Client!Converted)

ChkSCMFlag.Enabled = False
ChkSCMFlag.Value = IIf(IsNull(RsMst_Client!SCMFlag), 0, RsMst_Client!SCMFlag)

ChkCorpClient.Enabled = False
ChkCorpClient.Value = IIf(IsNull(RsMst_Client!CorpClient), 0, RsMst_Client!CorpClient)

mConvertedClientIDRow = IIf(IsNull(RsMst_Client!ConvertedClientID), 0, RsMst_Client!ConvertedClientID)
If mExchangeTypeID = 2 Then
   LblConvertedClientName.Visible = True
   CmbConvertedClientName.Visible = False
   TxtConvertedClientName.Visible = True
   TxtConvertedClientName.Text = ""
   tmp = "Select ClientName From Mst_Client Where Converted = 1 And ClientID = " & mConvertedClientIDRow & " "
   Call Tmp1Table
   If TmpRs1.RecordCount > 0 Then
      TxtConvertedClientName.Text = TmpRs1!ClientName
   End If
Else
   LblConvertedClientName.Visible = False
   CmbConvertedClientName.Visible = False
   TxtConvertedClientName.Visible = False
   TxtConvertedClientName = ""
End If

TxtCMSPCode.Locked = True
TxtCMSPCode = IIf(IsNull(RsMst_Client!CMSPCode), "", RsMst_Client!CMSPCode)
ChkKYCFlag.Enabled = False
ChkKYCFlag.Value = IIf(IsNull(RsMst_Client!KYCFlag), 0, RsMst_Client!KYCFlag)

TxtCreated = IIf(IsNull(RsMst_Client!CreatedBy), "", RsMst_Client!CreatedBy) & " :- " & IIf(IsNull(RsMst_Client!CreatedDate), "", RsMst_Client!CreatedDate)
TxtUpdated = IIf(IsNull(RsMst_Client!UpdatedBy), "", RsMst_Client!UpdatedBy) & " :- " & IIf(IsNull(RsMst_Client!UpdatedDate), "", RsMst_Client!UpdatedDate)
TxtNAVUpdated = IIf(IsNull(RsMst_Client!FAGUpdated), "", RsMst_Client!FAGUpdated) & " :- " & IIf(IsNull(RsMst_Client!FAGUpdatedDate), "", RsMst_Client!FAGUpdatedDate)

TxtCRFID = IIf(IsNull(RsMst_Client!CRFID), "", RsMst_Client!CRFID)
mCRFID = IIf(IsNull(RsMst_Client!CRFID), "", RsMst_Client!CRFID)

tmp = "Select CRFID,PAN,FirmName From Txn_CRF Where CRFID=" & Val(TxtCRFID.Text) & ""
Call TmpTableHRPro
If TmpRsHRPro.RecordCount > 0 Then
   TxtPAN.Text = IIf(IsNull(TmpRsHRPro!PAN), "", TmpRsHRPro!PAN)
   TxtFirmName.Text = TmpRsHRPro!FirmName
Else
   mCRFID = Null
   TxtPAN.Text = ""
   TxtFirmName.Text = ""
End If
''-- 09 Aug 2017 Ullhas Patil
TxtTANNO = IIf(IsNull(RsMst_Client!TANNO), "", RsMst_Client!TANNO)
TxtTANNO.Locked = True

Call ClearFrameAttachment
Call Grid_Head4
Frame2.Enabled = False
If mExchangeTypeID = 2 Then
   Call IndataAttachmentDetails
End If

If RsMst_Client.RecordCount = 1 Then
   Call GButtonLock_1(Me, True)
Else
   Call GButtonLock(Me, True)
End If

End Sub

Private Sub AddCmd_Click()
Edit_Flg = "A"

Call GButtonLock(Me, False)
TxtDepositorName.Locked = False
TxtClientID.Locked = False

CmbExchange.Locked = False
CmbExchange.Visible = True
TxtExchangeTypeID.Visible = False

CmbStatus.Locked = False
CmbStatus.Visible = True
TxtStatus.Visible = False

CmbCRFID.Visible = True
TxtCRFID.Visible = False

CmbDepositoryID.Locked = False
CmbDepositoryID.Visible = True
CmdDepositoryID.Visible = True
TxtDepositoryID.Visible = False

CmdExchangeTypeID.Visible = True
TxtFlag.Locked = False
TxtAddress.Locked = False
TxtOrganization.Locked = False
TxtEmailId.Locked = False
TxtTelephoneNo.Locked = False
TxtMobileNo.Locked = False

CmbRiskType.Visible = True
TxtRiskType.Visible = False
CmbRiskType.Text = ""

TxtStatus.Text = ""
ChkReturn.Value = vbUnchecked
TxtCreated = ""
TxtUpdated = ""
TxtClientIDRow = ""
TxtClientID = ""
TxtFlag = ""
TxtAddress = ""
TxtOrganization = ""
TxtEmailId = ""
TxtTelephoneNo = ""
TxtMobileNo = ""
TxtDepositorName.Text = ""
CmbExchange.Text = ""
CmbStatus.Text = ""
CmbDepositoryID.Text = ""
TxtClientGroupID = ""
mClientGroupID = 0
TxtNAVUpdated = ""
TxtNAVClientName = ""

ChkConverted.Enabled = False
ChkConverted.Value = 0

ChkSCMFlag.Enabled = False
ChkSCMFlag.Value = 0

TxtCMSPCode.Locked = False
TxtCMSPCode = ""
ChkKYCFlag.Enabled = True
ChkKYCFlag.Value = 0

LblConvertedClientName.Visible = False
CmbConvertedClientName.Visible = False
TxtConvertedClientName.Visible = False
mConvertedClientIDRow = 0

TxtPAN.Text = ""
TxtFirmName.Text = ""
CmbCRFID.Text = ""
TxtCRFID.Text = ""
mCRFID = Null
''-- 09 Aug 2017 Ullhas Patil
TxtTANNO.Locked = False
TxtTANNO = ""

Call ClearFrameAttachment
Call Grid_Head4
Frame2.Enabled = False
If mExchangeTypeID = 2 Then
   Frame2.Enabled = True
   lblAttachment.Enabled = True
   TxtAttachmentRemarks.Locked = False
   CmdAttachment.Enabled = True
   CmdGetFile.Enabled = True
   CmdRemoveAttachment.Enabled = True
End If


If RsMst_Client.RecordCount > 0 Then
   CmbExchange.SetFocus
End If
End Sub
Private Sub EditCmd_Click()
Edit_Flg = "E"
Call GButtonLock(Me, False)

'TxtClientID.Locked = False
'TxtExchangeTypeID.Locked = False
'CmbExchange.Visible = True
'CmbExchange.Locked = False
'CmdExchangeTypeID.Visible = True
'TxtExchangeTypeID.Visible = False
'TxtDepositoryID.Visible = False
'CmbDepositoryID.Visible = True
'CmbDepositoryID.Locked = False
'CmdDepositoryID.Visible = True
'TxtDepositoryID.Visible = False
TxtDepositorName.Locked = False
TxtFlag.Locked = False
TxtAddress.Locked = False
TxtOrganization.Locked = False
TxtEmailId.Locked = False
TxtTelephoneNo.Locked = False
TxtMobileNo.Locked = False
'TxtClientID.Locked = False
CmbExchange.Text = TxtExchangeTypeID
CmbStatus.Text = TxtStatus.Text
CmbDepositoryID.Text = TxtDepositoryID
TxtStatus.Visible = False
CmbStatus.Visible = True
CmbStatus.Locked = False
CmdDepository.Visible = True

CmbRiskType.Visible = True
TxtRiskType.Visible = False
CmbRiskType = TxtRiskType

CmbCRFID.Visible = True
TxtCRFID.Visible = False
CmbCRFID.Text = TxtCRFID.Text

If mExchangeTypeID = 1 Then
   ChkSCMFlag.Enabled = True
   tmp = "Select ConvertedClientIDRow from Mst_Client Where ConvertedClientID = " & Val(TxtClientIDRow) & "  "
   Call TmpTable
   If TmpRs.RecordCount > 0 Then
      ChkConverted.Enabled = False
   Else
      ChkConverted.Enabled = True
   End If
   
   tmp = "Select Distinct MC.ClientIDRow,MC.ClientName from Txn_CxTF_Details a"
   tmp = tmp & " Inner join Txn_CxTF_ClientDetail b On a.sm_Prefix = b.SM_Prefix And a.txntype = b.txntype and a.cxtfno = b.cxtfno"
   tmp = tmp & " Inner Join Mst_Client MC on b.clientidrow = MC.clientidrow"
   tmp = tmp & " Where a.TxnType = 'D' And a.SCMFlag = '1' And mc.ClientIDRow = " & Val(TxtClientIDRow) & ""
   
   'tmp = "Select Txn_CxTF_ClientDetail.ClientIDRow  from Txn_CxTF_ClientDetail Where Txn_CxTF_ClientDetail.ClientIDRow = " & TxtClientIDRow & " "
   Call TmpTable
   
   If TmpRs.RecordCount > 0 Then
      If ChkSCMFlag.Value = 0 Then
         ChkSCMFlag.Enabled = True
      Else
         ChkSCMFlag.Enabled = False
      End If
      
   Else
      ChkSCMFlag.Enabled = True
   End If
Else
   ChkConverted.Enabled = False
   ChkSCMFlag.Enabled = False
End If

If mExchangeTypeID = 2 Then
   LblConvertedClientName.Visible = True
   CmbConvertedClientName.Visible = True
   TxtConvertedClientName.Visible = False
   CmbConvertedClientName.Text = TxtConvertedClientName
   
   tmp = "Select ConvertedClientIDRow from Txn_RRNPWR Where ConvertedClientIDRow = " & mConvertedClientIDRow & "  "
   Call TmpTable
   
   If TmpRs.RecordCount > 0 Then
      CmbConvertedClientName.Enabled = False
   Else
      CmbConvertedClientName.Enabled = True
   End If
Else
   LblConvertedClientName.Visible = False
   CmbConvertedClientName.Visible = False
   TxtConvertedClientName.Visible = False
End If

TxtCMSPCode.Locked = False
ChkKYCFlag.Enabled = True

If TxtCMSPCode <> "" Then
   tmp = "Select * from Txn_CxTF_ClientDetail Where TxnType = 'D' And ClientIDRow = " & Val(TxtClientIDRow) & ""
   Call TmpTable
   If TmpRs.RecordCount > 0 Then
      TxtCMSPCode.Locked = True
   End If
End If

If ChkKYCFlag.Value = 1 Then
   tmp = "Select * from Txn_CxTF_ClientDetail Where TxnType = 'D' And ClientIDRow = " & Val(TxtClientIDRow) & ""
   Call TmpTable
   If TmpRs.RecordCount > 0 Then
      ChkKYCFlag.Enabled = False
   End If
End If

ChkCorpClient.Enabled = True

If mExchangeTypeID = 1 Then
   tmp = "Select * from Txn_CxTF_ClientDetail Where TxnType = 'D' And ClientIDRow = " & Val(TxtClientIDRow) & ""
   Call TmpTable
   If TmpRs.RecordCount > 0 Then
      If ChkCorpClient.Value = 1 Then
         ChkCorpClient.Enabled = False
      Else
         ChkCorpClient.Enabled = True
      End If
   End If
End If

''-- 09 Aug 2017 Ullhas Patil
TxtTANNO.Locked = False


Frame2.Enabled = False
If mExchangeTypeID = 2 Then
   Frame2.Enabled = True
   lblAttachment.Enabled = True
   TxtAttachmentRemarks.Locked = False
   CmdAttachment.Enabled = True
   CmdGetFile.Enabled = True
   CmdRemoveAttachment.Enabled = True
End If


End Sub
Private Function CheckForUseage() As Boolean
Dim Retval As Boolean

Retval = True

tmp = "select 'GSL' As  'Type', Mst_GSL.ClientIDRow  from Mst_GSL where Mst_GSL.ClientIDRow in( " & InmClientRowID & ") "
tmp = tmp & " union"
tmp = tmp & " select 'Txn_CDC' As 'Type',Txn_CDC.ClientIDRow  from Txn_CDC where Txn_CDC.ClientIDRow in( " & InmClientRowID & ") "
tmp = tmp & " union"
tmp = tmp & " select 'Mst_StorageContract' As 'Type',Mst_StorageContract.ClientIDRow  from Mst_StorageContract where Mst_StorageContract.ClientIDRow in( " & InmClientRowID & ") "
tmp = tmp & " union"
tmp = tmp & " select 'Txn_CxTF_ClientDetail' As 'Type',Txn_CxTF_ClientDetail.ClientIDRow  from Txn_CxTF_ClientDetail where Txn_CxTF_ClientDetail.ClientIDRow in( " & InmClientRowID & ") "
tmp = tmp & " union"
tmp = tmp & " select 'Txn_WHR_Detail' As 'Type',Txn_WHR_Detail.ClientIDRow  from Txn_WHR_Detail where Txn_WHR_Detail.ClientIDRow in( " & InmClientRowID & ") "
tmp = tmp & " union"
tmp = tmp & " select 'Txn_ISINRevalidation' As 'Type',Txn_ISINRevalidation.ClientRowID  from Txn_ISINRevalidation where Txn_ISINRevalidation.ClientRowID in( " & InmClientRowID & ") "
tmp = tmp & " union"
tmp = tmp & " select 'Txn_CxTF_GSL' As 'Type',Txn_CxTF_GSL.ClientIDRow  from Txn_CxTF_GSL where Txn_CxTF_GSL.ClientIDRow in( " & InmClientRowID & ") "
tmp = tmp & " union"
tmp = tmp & " select 'Txn_CxTF_ClientDetail' As 'Type',Txn_CxTF_ClientDetail.ClientIDRow  from Txn_CxTF_ClientDetail where Txn_CxTF_ClientDetail.ClientIDRow in( " & InmClientRowID & ") "
tmp = tmp & " union"
tmp = tmp & " select 'Txn_CDC_GSL_Detail' As 'Type',Txn_CDC_GSL_Detail.ClientIDRow  from Txn_CDC_GSL_Detail where Txn_CDC_GSL_Detail.ClientIDRow in( " & InmClientRowID & ") "
tmp = tmp & " union"
tmp = tmp & " select 'Txn_CDC_GSL_Detail' As 'Type',Txn_CDC_GSL_Detail.ClientIDRowSeller  from Txn_CDC_GSL_Detail where Txn_CDC_GSL_Detail.ClientIDRowSeller in( " & InmClientRowID & ") "
tmp = tmp & " union"
tmp = tmp & " select 'Mst_RRN' As 'Type',Mst_RRN.ClientIDRow  from Mst_RRN where Mst_RRN.ClientIDRow in( " & InmClientRowID & ") "
tmp = tmp & " union"
tmp = tmp & " select 'Txn_PledgeCreation' As 'Type',Txn_PledgeCreation.ClientIDRow  from Txn_PledgeCreation where Txn_PledgeCreation.ClientIDRow in( " & InmClientRowID & ") "
tmp = tmp & " union"
tmp = tmp & " select 'Txn_CashReceipt' As 'Type',Txn_CashReceipt.ClientIDRow  from Txn_CashReceipt where Txn_CashReceipt.ClientIDRow in( " & InmClientRowID & ") "

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


If mExchangeTypeID = 1 Then
   tmp = "Select * from Mst_Client Where ConvertedClientIDRow = " & Val(TxtClientIDRow) & " "
   Call TmpTable
   If TmpRs.RecordCount > 0 Then
      MsgBox "You can't delete this record since it is used in Client Master as converted client mapping !!!  "
      Exit Sub
   End If
End If

ans = MsgBox("Do you really want to DELETE this record???", vbYesNo)


If ans = vbYes Then
   RsMst_Client.Find "ClientIDRow = '" & Val(TxtClientIDRow) & "'"
   RsMst_Client.Delete
   RsMst_Client.Update
   If RsMst_Client.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_Client.MoveNext
   If RsMst_Client.EOF() Then
      RsMst_Client.MoveLast
   End If
   Call Indata
End If
Exit Sub
msgError:
MsgBox "Child record Present "
RsMst_Client.CancelUpdate
End Sub
'Function CheckDuplicateClient() As Boolean
'Dim RetVal As Boolean
'RetVal = False
'
'For j = 1 To MSHFlexGrid2.Rows - 1
'    If MSHFlexGrid2.TextMatrix(j, 2) = TxtClientID And _
'       MSHFlexGrid2.TextMatrix(j, 1) = CmbExchange.Text And _
'       MSHFlexGrid2.TextMatrix(j, 4) = CmbDepositoryID.Text Then
'       If CmdAdd2List.Caption = "Add to List" Then
'          RetVal = True
'          Exit For
'       Else
'          If j <> MSHFlexGrid2.RowSel Then
'             RetVal = True
'             Exit For
'          End If
'       End If
'    End If
'Next
'CheckDuplicateClient = RetVal
'End Function

Private Sub SaveCmd_Click()
If TxtDepositorName.Text = "" Then
   MsgBox "Blank DepositorName Not Allowed !!! "
   TxtDepositorName.SetFocus
   Exit Sub
End If
If TxtClientID.Text = "" Then
   MsgBox "Blank ClientID Not Allowed !!! "
   TxtClientID.SetFocus
   Exit Sub
End If
If CmbExchange.Text = "" Then
   MsgBox "Blank Exchange Type Not Allowed !!! "
   CmbExchange.SetFocus
   Exit Sub
End If
'If LCase(CmbExchange.Text) = "non ncdex" Then
'   If CmbClientGroup.Text = "" Then
'      MsgBox "Blank Client Group Not Allowed for Non NCDEX Client !!! "
'      CmbClientGroup.SetFocus
'      Exit Sub
'   End If
'End If
If CmbRiskType.Text = "" Then
   MsgBox "Blank Risk Type Not Allowed !!! "
   CmbRiskType.SetFocus
   Exit Sub
End If
If CmbDepositoryID.Text = "" Then
   MsgBox "Blank Depository Type Not Allowed!!!"
   CmbDepositoryID.SetFocus
   Exit Sub
End If
If CmbStatus.Text = "" Then
   MsgBox "Blank Status Not Allowed !!! "
   CmbStatus.SetFocus
   Exit Sub
End If
''--- 09 Aug 2017 Ullhas Patil
''If TxtTANNo.Text = "" Then
''   MsgBox "Blank TAN No Allowed !!! "
''   TxtTANNo.SetFocus
''   Exit Sub
''Else
If Trim(TxtTANNO.Text) <> "" Then
   If ValidateTANNo(UCase(TxtTANNO.Text)) = False Then
      MsgBox "Invalid TAN No !!!"
      TxtTANNO.SetFocus
      Exit Sub
   End If
End If

'If LCase(CmbStatus.Text) = "deactive" And TxtClientIDRow.Text <> "" Then
'   If CheckStockInvoiceClear = False Then
'      MsgBox "You can not deactive it due to stock or pending invoice!!!"
'      CmbStatus.SetFocus
'      Exit Sub
'   End If
'End If

'If CmbCRFID.Text = "" Then
'   MsgBox "Blank CRFID Not Allowed !!! "
'   CmbCRFID.SetFocus
'   Exit Sub
'End If

If Edit_Flg = "A" Then
   If RsMst_Client.RecordCount > 0 Then
      RsMst_Client.MoveFirst
      tmp = "ClientID = '" & TxtClientID.Text & "' And DepositoryID = '" & mDepositoryID & "' And ExchangeTypeID = " & mExchangeTypeID & ""
      'RsMst_Client.Find "ClientID = '" & TxtClientID.Text & "'"
      RsMst_Client.Filter = tmp
      If Not RsMst_Client.EOF Then
         MsgBox "Duplicate Client Not Allowed !!!"
         TxtClientID.SetFocus
         Exit Sub
      Else
         If mExchangeTypeID = 1 Then
            If CheckDupliClient = False Then
               MsgBox "Duplicate Client Name Not Allowed !!!"
               TxtDepositorName.SetFocus
               Exit Sub
            End If
'            RsMst_Client.Filter = ""
'            RsMst_Client.Filter = "ClientName = '" & TxtDepositorName & "' And ExchangeTypeID = " & mExchangeTypeID & " "
'            If Not RsMst_Client.EOF Then
'               MsgBox "Duplicate Client Name Not Allowed !!!"
'               TxtDepositorName.SetFocus
'               Exit Sub
'            End If
         End If
         RsMst_Client.Filter = ""
      End If
      ''-- Alter for duplicate TAN No 09 Aug 2017 Ullhas Patil
      RsMst_Client.Filter = ""
      RsMst_Client.MoveFirst
      If Trim(TxtTANNO.Text) <> "" Then
      RsMst_Client.Find ("TANNO ='" & Trim(TxtTANNO.Text) & "'")
      If Not RsMst_Client.EOF Then
        Select Case MsgBox("Duplicate TAN No Found with another Client, Do you want to save the Client with same TAN No?", vbYesNo)
            Case vbNo
                TxtTANNO.SetFocus
                Exit Sub
        End Select
      End If
      End If
   End If
   RsMst_Client.AddNew
   
   'RsMst_Client!ClientID = GetDepositorID
   'TxtDepositorID.Text = RsMst_Client!ClientID
   RsMst_Client!ClientIDRow = GetRowID
   RsMst_Client!G_UID = GetGUID
Else
   RsMst_Client.MoveFirst
   tmp = "ClientID = '" & TxtClientID.Text & "' And DepositoryID = '" & mDepositoryID & "' And ExchangeTypeID = " & mExchangeTypeID & ""
   'RsMst_Client.Find "ClientID = '" & TxtClientID.Text & "'"
   RsMst_Client.Filter = tmp
   If Not RsMst_Client.EOF Then
      If RsMst_Client!ClientIDRow <> TxtClientIDRow.Text Then
         MsgBox "Duplicate Client Not Allowed !!! "
         TxtClientID.SetFocus
         Exit Sub
      Else
         If mExchangeTypeID = 1 Then
            If CheckDupliClient = False Then
               If RsMst_Client!ClientIDRow <> TxtClientIDRow.Text Then
                  MsgBox "Duplicate Client Name Not Allowed !!!"
                  TxtDepositorName.SetFocus
                  Exit Sub
               End If
            End If
'            RsMst_Client.Filter = ""
'            RsMst_Client.Filter = "ClientName = '" & TxtDepositorName & "' And ExchangeTypeID = " & mExchangeTypeID & " "
'            If Not RsMst_Client.EOF Then
'               If RsMst_Client!ClientIDRow <> TxtClientIDRow.Text Then
'                  MsgBox "Duplicate Client Name Not Allowed !!!"
'                  TxtDepositorName.SetFocus
'                  Exit Sub
'               End If
'            End If
         End If
      End If
   End If
   RsMst_Client.Filter = ""
   RsMst_Client.MoveFirst
   RsMst_Client.Find "ClientIDRow = " & TxtClientIDRow.Text
   If RsMst_Client.EOF Then
      MsgBox "No Such Record Found!!! "
      Exit Sub
   End If
   
   ''-- Alter for duplicate TAN No 09 Aug 2017 Ullhas Patil
   RsMst_Client.Filter = ""
   RsMst_Client.MoveFirst
   If Trim(TxtTANNO.Text) <> "" Then
   RsMst_Client.Find ("TANNO ='" & Trim(TxtTANNO.Text) & "'")
   If Not RsMst_Client.EOF Then
    If RsMst_Client!ClientIDRow <> TxtClientIDRow.Text Then
      Select Case MsgBox("Duplicate TAN No Found with another Client, Do you want to save the Client with same TAN No?", vbYesNo)
          Case vbNo
              TxtTANNO.SetFocus
              Exit Sub
      End Select
    End If
   End If
   End If
   RsMst_Client.Filter = ""
   RsMst_Client.MoveFirst
   RsMst_Client.Find "ClientIDRow = " & TxtClientIDRow.Text
   
End If

RsMst_Client!ClientID = TxtClientID.Text
RsMst_Client!ClientName = TxtDepositorName.Text
RsMst_Client!ExchangeTypeID = mExchangeTypeID
If LCase(CmbExchange.Text) = "non ncdex" And mClientGroupID <> 0 Then
   RsMst_Client!ClientGroupID = mClientGroupID
Else
   RsMst_Client!ClientGroupID = Null
End If
RsMst_Client!DepositoryID = mDepositoryID
RsMst_Client!Flag = "0" 'TxtFlag.Text
RsMst_Client!Address = TxtAddress.Text
RsMst_Client!Organization = TxtOrganization.Text
RsMst_Client!EmailID = TxtEmailId.Text
RsMst_Client!TelephoneNo = TxtTelephoneNo.Text
RsMst_Client!MobileNo = TxtMobileNo.Text
RsMst_Client!Status = Left$(CmbStatus.Text, 1)
If mExchangeTypeID = 1 Then
   RsMst_Client!Converted = ChkConverted.Value
   RsMst_Client!SCMFlag = ChkSCMFlag.Value
   RsMst_Client!CorpClient = ChkCorpClient.Value
Else
   RsMst_Client!Converted = 0
   RsMst_Client!SCMFlag = 0
   RsMst_Client!CorpClient = 0
End If

If CmbConvertedClientName.Visible = True Then 'If mExchangeTypeID = 1 Then
   RsMst_Client!ConvertedClientIDRow = mConvertedClientIDRow
Else
   RsMst_Client!ConvertedClientIDRow = Null
End If

RsMst_Client!CMSPCode = Trim(TxtCMSPCode)
RsMst_Client!KYCFlag = ChkKYCFlag.Value
RsMst_Client!RiskType = CmbRiskType.Text
''--- 09 Aug 2017 Ullhas Patil
RsMst_Client!TANNO = TxtTANNO.Text

If IsNull(RsMst_Client!CreatedBy) = True Or RsMst_Client!CreatedBy = "" Then
   RsMst_Client!CreatedBy = Gen_UserLoginID
   RsMst_Client!CreatedDate = Now
Else
   RsMst_Client!UpdatedBy = Gen_UserLoginID
   RsMst_Client!UpdatedDate = Now
End If
'RsMst_Client!DepositorID = ""
RsMst_Client!CRFID = IIf(IsNull(mCRFID), Null, mCRFID)
RsMst_Client.Update


If mExchangeTypeID = 2 Then
   Call SaveAttachmentDetails
Else
   tmp = "Delete from Mst_ClientAttachment WHere ClientIDRow = " & Val(TxtClientIDRow) & ""
   Call TmpTableSMAttachment
End If

'RsMst_Client.Requery
Dim om_ClientIDRow As Variant
om_ClientIDRow = RsMst_Client!ClientIDRow
RsMst_Client.Find "ClientIDRow = " & om_ClientIDRow

'If RsMst_Client.EOF Then
'   RsMst_Client.MoveFirst
'End If
Call Indata
Call GButtonLock(Me, True)
End Sub

'Private Function CheckStockInvoiceClear() As Boolean
'CheckStockInvoiceClear = True
'tmp = " Select MS.StateName, ML.LocationName, TI.TIID, TI.NAVCode, TI.TIAmount,"
'tmp = tmp & " (Select isnull(Sum(TCRT.RecAmount),0) as RecAmount From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.CRID= TCRT.CRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID)+ isnull(TI.AdjAmountFAG,0) +"
'tmp = tmp & " (Select isnull(Sum(TCRT.RecTDS),0) as RecTDS From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.CRID= TCRT.CRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID) +"
'tmp = tmp & " (Select isnull(Sum(TCRT.RecTDS),0) as RecTDS From Txn_CRTDSAdjustmentDetail TCRT Inner Join Txn_CRTDSAdjustment TCR on TCR.TDSAID= TCRT.TDSAID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.Flag<>'R'And TIID=TI.TIID )+"
'tmp = tmp & " (Select isnull(Sum(TCRT.RecOtherDed),0) as RecOtherDed From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.CRID= TCRT.CRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID)+"
'tmp = tmp & " (Select isnull(Sum(TCRT.RecOtherDed),0) as RecOtherDed From Txn_CRTDSAdjustmentDetail TCRT Inner Join Txn_CRTDSAdjustment TCR on TCR.TDSAID= TCRT.TDSAID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.Flag<>'R'And TIID=TI.TIID) as RecAmount,"
'tmp = tmp & " (Select top 1 ClientName From Mst_Client Where NAVCode = TI.NAVCode Or RNAVCode = TI.NAVCode) as ClientName,"
'tmp = tmp & " (Select top 1 DepositoryName From Mst_Depository Where NAVCode = TI.NAVCode) as DepositoryName"
'tmp = tmp & " From  Txn_TIDetails TI"
'tmp = tmp & " Left Join Mst_Location ML on TI.LocationID=ML.LocationID"
'tmp = tmp & " Left Join Mst_State MS on ML.StateID=MS.StateID "
'tmp = tmp & " Where TI.ClientIDRow = " & TxtClientIDRow.Text
'tmp = tmp & " And TIAmount-((select isnull(Sum(RecAmount)+Sum(RecTDS)+Sum(RecOtherDed),0) From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.CRID= TCRT.CRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCRT.TIID=TI.TIID And TCR.ReturnFlag<>'Y')+ (Select isnull(Sum(RecTDS)+Sum(RecOtherDed),0)  From Txn_CRTDSAdjustmentDetail TCRT Inner Join Txn_CRTDSAdjustment TCR on TCR.TDSAID = TCRT.TDSAID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCRT.TIID=TI.TIID And TCR.Flag <> 'R')+AdjAmountFAG)>0"
'tmp = tmp & " group by ML.LocationName, MS.StateName, TI.TIID, TI.NAVCode, TI.TIAmount, TI.AdjAmountFAG, TI.NAVCode"
'Call Tmp7Table
'If TmpRs7.RecordCount > 0 Then
'   CheckStockInvoiceClear = False
'   Exit Function
'End If
'
'tmp = " select * from mst_gsl where exchangetypeid = 1 and balancebags > 0 and ClientIDRow = " & TxtClientIDRow.Text
'Call Tmp7Table
'If TmpRs7.RecordCount > 0 Then
'   CheckStockInvoiceClear = False
'   Exit Function
'End If
'End Function

Private Function GetDepositorID() As Double
Dim Retval As Double
If Gen_DBType = "MDB" Then
    tmp = "Select max(val(ClientID)) from Mst_Client where DepositorId is Null"
Else
    tmp = "Select max(convert(decimal(10,0),ClientID)) from Mst_Client where DepositorId is Null"
End If
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))

Retval = Retval + 1
GetDepositorID = Retval
End Function
Private Function GetRowID() As Double
Dim Retval As Double
tmp = "Select max(ClientIDRow) from Mst_Client "
Call Tmp1Table
Retval = IIf(IsNull(TmpRs1.Fields(0)), 0, TmpRs1.Fields(0))
Retval = Retval + 1
TmpRs1.Close
GetRowID = Retval
End Function

Private Sub CmdExchangeTypeID_Click()
FrmMst_ExchangeType.Show
End Sub
Private Sub CmdDepositoryID_Click()
    FrmMst_Depository.Show
End Sub

Private Sub CmbDepositoryID_GotFocus()
If CmbExchange.Text = "" Then Exit Sub
tmp = CmbDepositoryID.Text
Call TableMst_Depository(" Where ExchangeTypeID = " & mExchangeTypeID)
CmbDepositoryID.Clear
If RsMst_Depository.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Depository.RecordCount
  If UCase(CmbExchange.Text) = "SPOT" Or UCase(CmbExchange.Text) = "NCDEX" Then
    If RsMst_Depository!DepositoryName <> "NCMSL" Then
      CmbDepositoryID.AddItem RsMst_Depository!DepositoryName & "~" & RsMst_Depository!DepositoryID
    End If
  ElseIf UCase(CmbExchange.Text) = "NON NCDEX" Then
    If RsMst_Depository!DepositoryName = "NCMSL" Then
      CmbDepositoryID.AddItem RsMst_Depository!DepositoryName & "~" & RsMst_Depository!DepositoryID
    End If
     'CmbDepositoryID.AddItem RsMst_Depository!DepositoryName
  End If
  RsMst_Depository.MoveNext
Next
CmbDepositoryID.Text = tmp
End Sub
Private Sub CmbDepositoryID_LostFocus()
Dim Str_() As String

If CmbDepositoryID.Text = "" Then
   Exit Sub
End If

Str_ = Split(CmbDepositoryID.Text, "~")

If UBound(Str_) <> 1 Then
   MsgBox "Invalid selection !!!!! "
   CmbDepositoryID.SetFocus
   Exit Sub
End If

RsMst_Depository.MoveFirst
RsMst_Depository.Find "DepositoryID = '" & Str_(1) & "'"
If RsMst_Depository.EOF Then
   MsgBox "Invalid selection "
   CmbDepositoryID.SetFocus
   Exit Sub
ElseIf UCase(CmbExchange.Text) = "SPOT" Or UCase(CmbExchange.Text) = "NCDEX" Then
   If Str_(0) = "NCMSL" Then
    MsgBox "Invalid selection "
    CmbDepositoryID.SetFocus
    Exit Sub
   End If
ElseIf UCase(CmbExchange.Text) = "NON NCDEX" Then
   If Str_(0) <> "NCMSL" Then
    MsgBox "Invalid selection "
    CmbDepositoryID.SetFocus
    Exit Sub
   End If
End If
mDepositoryID = Str_(1) 'RsMst_Depository!DepositoryID

End Sub

Private Sub TxtClientID_LostFocus()
TxtClientID = Trim(TxtClientID)
End Sub


Private Sub TxtClientName_LostFocus()
TxtClientName = Trim(TxtClientName)
End Sub

Private Sub TxtDepositorName_LostFocus()
TxtDepositorName = Trim(TxtDepositorName)
End Sub

Private Sub TxtSearchClient12_LostFocus()
Dim x As Integer
RsMst_Client.MoveFirst
Call Grid_Head
If TxtSearchClient <> "" Then
    RsMst_Client.Find ("ClientID = '" & Trim(TxtSearchClient) & "'")
    If RsMst_Client.EOF Then
      MsgBox "No Such Client Exist in DataBase!!!"
      TxtSearchClient.SetFocus
      Exit Sub
    End If
  Call TableMst_Client("Where ClientID = '" & Trim(TxtSearchClient) & "'")
  MSHFlexGrid1.Rows = RsMst_Client.RecordCount + 1
  For x = 1 To RsMst_Client.RecordCount
     Call ShowGrid(x)
     RsMst_Client.MoveNext
  Next
Else
 'RsMst_ISIN.Sort
 Call TableMst_Client(" Order By ClientIDRow ")
 MSHFlexGrid1.Rows = RsMst_Client.RecordCount + 1

 For x = 1 To RsMst_Client.RecordCount
     Call ShowGrid(x)
     RsMst_Client.MoveNext
 Next

End If
End Sub

Private Sub MSHFlexGrid1_Click()
Dim I As Variant
I = MSHFlexGrid1.RowSel
If I <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(I, 14) = "" Then Exit Sub
RsMst_Client.MoveFirst
RsMst_Client.Find "ClientIDRow = " & MSHFlexGrid1.TextMatrix(I, 14) & ""
If Not RsMst_Client.EOF Then
   Call Indata
   Frame0.Visible = False
End If

'Call TableMst_Client(" Order By ClientIDRow ")
End Sub
Private Sub CmbSearchExchange_GotFocus()
mExchangeID = Null
tmp = CmbSearchExchange.Text
Call TableMst_ExchangeType
CmbSearchExchange.Clear
If RsMst_ExchangeType.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_ExchangeType.RecordCount
    CmbSearchExchange.AddItem RsMst_ExchangeType!ExchangeType
    RsMst_ExchangeType.MoveNext
Next
CmbSearchExchange.Text = tmp

End Sub

Private Sub CmbSearchExchange_LostFocus()
Dim x As Integer
If CmbSearchExchange.Text = "" Then
   Exit Sub
End If
RsMst_ExchangeType.MoveFirst
RsMst_ExchangeType.Find "ExchangeType = '" & CmbSearchExchange.Text & "'"
If RsMst_ExchangeType.EOF Then
   MsgBox "Invalid selection "
   CmbSearchExchange.SetFocus
   Exit Sub
End If
mExchangeID = RsMst_ExchangeType!ExchangeTypeID

End Sub
Private Function CheckDupliClient() As Boolean
Dim Retval As Boolean
Retval = True
RsMst_Client.Filter = ""
RsMst_Client.Filter = "ClientName = '" & TxtDepositorName & "' And ExchangeTypeID = " & mExchangeTypeID & " "
If Not RsMst_Client.EOF Then
   Retval = False
End If
CheckDupliClient = Retval
End Function

Private Sub TxtTANNo_LostFocus()
If TxtTANNO.Locked = False Then
    If Trim(TxtTANNO.Text) <> "" Then
       TxtTANNO = UCase(TxtTANNO)
       If ValidateTANNo(TxtTANNO) = False Then
        MsgBox "Invalid TAN No"
        TxtTANNO.SetFocus
        Exit Sub
       End If
    End If
End If
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

Private Sub SaveAttachmentDetails()

If MSHFlexGrid5.Rows <= 1 Then Exit Sub


Call TableMst_ClientAttachment

For I = 1 To MSHFlexGrid5.Rows - 1
    If MSHFlexGrid5.TextMatrix(I, 1) = "" Then Exit For
    If MSHFlexGrid5.TextMatrix(I, 4) = "" Then
       RsMst_ClientAttachment.AddNew
 
       'RsMst_ClientAttachment!TxnID = GetMaxAttachID
       RsMst_ClientAttachment!ClientIDRow = TxtClientIDRow
       mFlag = SaveFileToDB(MSHFlexGrid5.TextMatrix(I, 3), RsMst_ClientAttachment, "File")
       mFlag = GetFileFromPath(MSHFlexGrid5.TextMatrix(I, 3))
       RsMst_ClientAttachment!FileName = "CLientKYD" & Val(TxtClientIDRow) & "-" & mFlag
    Else
       RsMst_ClientAttachment.MoveFirst
       RsMst_ClientAttachment.Find "AttachmentID = " & MSHFlexGrid5.TextMatrix(I, 4) & ""
    End If
    RsMst_ClientAttachment!Remarks = MSHFlexGrid5.TextMatrix(I, 2)
    RsMst_ClientAttachment!UpdatedDate = Now
    RsMst_ClientAttachment!UpdatedBy = Gen_UserLoginID
    RsMst_ClientAttachment!G_UID = GetGUID
    If IsNull(RsMst_ClientAttachment!CreatedBy) = True Or RsMst_ClientAttachment!CreatedBy = "" Then
       RsMst_ClientAttachment!CreatedBy = Gen_UserLoginID
       RsMst_ClientAttachment!CreatedDate = Now
    Else
       RsMst_ClientAttachment!UpdatedBy = Gen_UserLoginID
       RsMst_ClientAttachment!UpdatedDate = Now
    End If
    RsMst_ClientAttachment.Update
Next

End Sub

Private Function GetMaxAttachID() As Double
Dim Retval As Double
tmp = "Select max(TxnID) from Mst_ClientAttachment"
Call TmpTableSMAttachment
Retval = IIf(IsNull(TmpRsSM_Attachment.Fields(0)), 0, TmpRsSM_Attachment.Fields(0))
Retval = Retval + 1
GetMaxAttachID = Retval
End Function


Private Sub IndataAttachmentDetails()
Call Grid_Head4
Call ClearFrameAttachment

Call TableMst_ClientAttachment(" Where ClientIDRow = " & TxtClientIDRow.Text & " ")

If Right(App.Path, 1) = "\" Then
   File_Path = App.Path & "Agreement Attachment\Client KYD\" & Val(TxtClientIDRow)
Else
   File_Path = App.Path & "\Agreement Attachment\Client KYD\" & Val(TxtClientIDRow)
End If
If oFs.FolderExists(File_Path) = True Then
   Set oFolder = oFs.GetFolder(File_Path)
   For Each oFile In oFolder.Files
       Call oFs.DeleteFile(oFile.Path, True)
   Next
Else
   If Not oFs.FolderExists(File_Path) Then
      Call oFs.CreateFolder(File_Path)
   End If
End If


MSHFlexGrid5.Rows = RsMst_ClientAttachment.RecordCount + 2
For I = 1 To RsMst_ClientAttachment.RecordCount
'Client KYD
    MSHFlexGrid5.TextMatrix(I, 0) = I
    MSHFlexGrid5.TextMatrix(I, 1) = IIf(IsNull(RsMst_ClientAttachment!FileName), "", RsMst_ClientAttachment!FileName)
    MSHFlexGrid5.TextMatrix(I, 2) = IIf(IsNull(RsMst_ClientAttachment!Remarks), "", RsMst_ClientAttachment!Remarks)
    MSHFlexGrid5.TextMatrix(I, 3) = App.Path & "\Agreement Attachment\Client KYD\" & Val(TxtClientIDRow) & "\" & MSHFlexGrid5.TextMatrix(I, 1)
    File_Path = MSHFlexGrid5.TextMatrix(I, 3) 'App.Path & "\Project Attachment\"
    Call LoadFileFromDB(File_Path, RsMst_ClientAttachment, "File")
    MSHFlexGrid5.TextMatrix(I, 4) = RsMst_ClientAttachment!AttachmentID
    RsMst_ClientAttachment.MoveNext
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

