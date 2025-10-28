VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmTxn_ClientFeedBackDocuments 
   Caption         =   "Client Feedback Documents"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   1935
      Left            =   120
      TabIndex        =   26
      Top             =   9600
      Visible         =   0   'False
      Width           =   15000
      Begin VB.ComboBox CmbSClientID 
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
         ItemData        =   "FrmTxn_ClientFeedBackDocuments.frx":0000
         Left            =   4920
         List            =   "FrmTxn_ClientFeedBackDocuments.frx":0002
         TabIndex        =   48
         Top             =   187
         Width           =   9975
      End
      Begin VB.ComboBox CmbSExchangeID 
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
         ItemData        =   "FrmTxn_ClientFeedBackDocuments.frx":0004
         Left            =   1680
         List            =   "FrmTxn_ClientFeedBackDocuments.frx":0006
         Sorted          =   -1  'True
         TabIndex        =   46
         Top             =   187
         Width           =   1815
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
         Left            =   13875
         TabIndex        =   28
         Top             =   630
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
         Left            =   12850
         TabIndex        =   27
         Top             =   630
         Width           =   1000
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1455
         Left            =   75
         TabIndex        =   29
         Top             =   1050
         Width           =   14835
         _ExtentX        =   26167
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
         Left            =   11475
         TabIndex        =   30
         Top             =   675
         Width           =   90
      End
      Begin VB.Label Label6 
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
         Height          =   255
         Left            =   3600
         TabIndex        =   47
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label5 
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
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   240
         Width           =   1695
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
         TabIndex        =   31
         Top             =   660
         Width           =   2415
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
      Height          =   4935
      Left            =   240
      TabIndex        =   13
      Top             =   3510
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   8705
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
   Begin VB.Frame ButtonFrm 
      Height          =   1080
      Left            =   120
      TabIndex        =   15
      Top             =   8520
      Width           =   15015
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
         Left            =   105
         TabIndex        =   25
         Top             =   195
         Width           =   3060
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
         Left            =   105
         TabIndex        =   24
         Top             =   600
         Width           =   3060
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
         Left            =   3165
         TabIndex        =   23
         Top             =   195
         Width           =   3060
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
         Left            =   3165
         TabIndex        =   22
         Top             =   600
         Width           =   3060
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
         Left            =   6225
         TabIndex        =   21
         Top             =   195
         Width           =   2940
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
         Left            =   6225
         TabIndex        =   20
         Top             =   600
         Width           =   2940
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
         Left            =   9165
         TabIndex        =   19
         Top             =   195
         Width           =   2940
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
         Left            =   9165
         TabIndex        =   18
         Top             =   600
         Width           =   2940
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
         Left            =   12105
         TabIndex        =   17
         Top             =   195
         Width           =   2700
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
         Left            =   12105
         TabIndex        =   16
         Top             =   600
         Width           =   2700
      End
   End
   Begin VB.Frame Frame1 
      Height          =   8415
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   15015
      Begin VB.TextBox TxtTrancationDate 
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
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   240
         Width           =   1575
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
         Left            =   8840
         Locked          =   -1  'True
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   1580
         Width           =   6055
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
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   1580
         Width           =   6015
      End
      Begin VB.TextBox TxtRNAVCode 
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
         Left            =   13200
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1147
         Width           =   1695
      End
      Begin VB.TextBox TxtNAVCode 
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
         Left            =   9960
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1147
         Width           =   1695
      End
      Begin VB.TextBox TxtDepository 
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
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1147
         Width           =   7215
      End
      Begin VB.PictureBox FrameAttachment 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1300
         Left            =   120
         ScaleHeight     =   1275
         ScaleWidth      =   14745
         TabIndex        =   35
         Top             =   2040
         Width           =   14775
         Begin VB.TextBox TxtAttachmentRemarks 
            Appearance      =   0  'Flat
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2057
               SubFormatType   =   0
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Left            =   1200
            MaxLength       =   1000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   10
            Top             =   480
            Width           =   10935
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
            Height          =   375
            Left            =   12240
            TabIndex        =   12
            Top             =   820
            Width           =   2445
         End
         Begin VB.CommandButton CmdAddAttachment 
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
            Height          =   375
            Left            =   12240
            TabIndex        =   11
            Top             =   430
            Width           =   2445
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
            Height          =   375
            Left            =   12240
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   50
            Width           =   2445
         End
         Begin VB.Label lblAttachment 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Left            =   1200
            TabIndex        =   9
            Top             =   45
            Width           =   10935
         End
         Begin VB.Label Label73 
            Caption         =   "Remark"
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
            Left            =   50
            TabIndex        =   37
            Top             =   713
            Width           =   1215
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
            Left            =   50
            TabIndex        =   36
            Top             =   60
            Width           =   1215
         End
      End
      Begin VB.TextBox TxtTrancationID 
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
         Left            =   1455
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   240
         Width           =   1575
      End
      Begin VB.ComboBox CmbClientID 
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
         ItemData        =   "FrmTxn_ClientFeedBackDocuments.frx":0008
         Left            =   1440
         List            =   "FrmTxn_ClientFeedBackDocuments.frx":000A
         TabIndex        =   3
         Top             =   700
         Width           =   13455
      End
      Begin VB.ComboBox CmbExchangeTypeID 
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
         ItemData        =   "FrmTxn_ClientFeedBackDocuments.frx":000C
         Left            =   8295
         List            =   "FrmTxn_ClientFeedBackDocuments.frx":000E
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox TxtExchangeType 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
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
         Left            =   8280
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox TxtClientName 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
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
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   700
         Width           =   13455
      End
      Begin MSComDlg.CommonDialog CDialog 
         Left            =   10560
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Trancation Date"
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
         Left            =   3120
         TabIndex        =   50
         Top             =   300
         Width           =   1440
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
         Left            =   7680
         TabIndex        =   44
         Top             =   1640
         Width           =   1815
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
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   1640
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "R-NAV Code"
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
         Left            =   11880
         TabIndex        =   40
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "NAV Code"
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
         Left            =   8880
         TabIndex        =   39
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Depository"
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
         TabIndex        =   38
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Trancation ID"
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
         TabIndex        =   34
         Top             =   300
         Width           =   1335
      End
      Begin VB.Label LblClient_WHR_RRN 
         Caption         =   "Client Name"
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
         TabIndex        =   33
         Top             =   753
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Exchange Type"
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
         Left            =   6480
         TabIndex        =   32
         Top             =   300
         Width           =   1695
      End
   End
End
Attribute VB_Name = "FrmTxn_ClientFeedBackDocuments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oFs As New FileSystemObject
Dim oFs1 As New FileSystemObject
Dim oFolder As Folder
Dim oFile As File
Dim File_Path As Variant
Dim Edit_Flg As String
Dim mExchangeTypeID, mClientIDRow As Variant
Dim mSExchangeTypeID, mSClientIDRow As Variant

Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)
Frame1.Enabled = True
FrameAttachment.Enabled = True

TxtTrancationID = ""
TxtTrancationDate = Format(Date, "DD-MMM-YYYY")
TxtExchangeType = "": TxtExchangeType.Visible = False
CmbExchangeTypeID = "": CmbExchangeTypeID.Visible = True

TxtClientName = "": TxtClientName.Visible = False
CmbClientID = "": CmbClientID.Visible = True

TxtDepository = ""
TxtNAVCode = ""
TxtRNAVCode = ""
TxtCreated = ""
TxtUpdated = ""
Call Grid_HeadAttachment
If RsTxn_ClientFeedbackDoc.RecordCount > 0 Then
   CmbExchangeTypeID.SetFocus
Else
   SearchCmd.Enabled = False
End If
End Sub

Private Sub CmbClientId_GotFocus()
    If CmbExchangeTypeID = "" Then
       MsgBox "Select Exchange Type.", vbOKOnly + vbInformation, App.Title
       CmbExchangeTypeID.SetFocus
       Exit Sub
    End If
    
    tmp = CmbClientID.Text
    Call TableMst_Client(" Where Status = 'A' AND ExchangeTypeID = " & mExchangeTypeID & " Order By ClientName ")
    CmbClientID.Clear
    If RsMst_Client.RecordCount = 0 Then
       MsgBox "No record found!!!"
       Exit Sub
    End If
    For I = 1 To RsMst_Client.RecordCount
        CmbClientID.AddItem IIf(IsNull(RsMst_Client!ClientName), "", RsMst_Client!ClientName) & "#" & IIf(IsNull(RsMst_Client!ClientID), "", RsMst_Client!ClientID) & "#" & IIf(IsNull(RsMst_Client!ClientIDRow), "", RsMst_Client!ClientIDRow)
        RsMst_Client.MoveNext
    Next
    CmbClientID.Text = tmp
End Sub

Private Sub CmbClientId_LostFocus()
Dim SplitClient As Variant
If CmbClientID.Text = "" Then
   TxtDepository = ""
   TxtNAVCode = ""
   TxtRNAVCode = ""
   Exit Sub
End If

SplitClient = Split(CmbClientID, "#")
If UBound(SplitClient) <> 2 Then
    MsgBox "Invalid Client.", vbOKOnly + vbInformation, App.Title
    CmbClientID.SetFocus
    Exit Sub
End If

RsMst_Client.MoveFirst
RsMst_Client.Find "ClientIDRow = " & SplitClient(2)
If RsMst_Client.EOF Then
   MsgBox "Invalid selection!!!"
   CmbClientID.SetFocus
   Exit Sub
End If
mClientIDRow = RsMst_Client!ClientIDRow
   TxtDepository = ""
   TxtNAVCode = ""
   TxtRNAVCode = ""
tmp = " SELECT MS.ExchangeTypeID,MS.ClientID,MS.ClientIDRow,MS.ClientName,MS.DepositoryID,MD.DepositoryName,MS.NAVCode,MS.RNAVCode "
tmp = tmp & " FROM Mst_Client AS MS Inner Join Mst_Depository AS MD On MS.DepositoryID = MD.DepositoryID Where MS.ClientIDRow = " & mClientIDRow
Call TmpTable
If TmpRs.RecordCount > 0 Then
   TxtDepository = IIf(IsNull(TmpRs!DepositoryName), "", TmpRs!DepositoryName)
   TxtNAVCode = IIf(IsNull(TmpRs!NAVCode), "", TmpRs!NAVCode)
   TxtRNAVCode = IIf(IsNull(TmpRs!RNAVCode), "", TmpRs!RNAVCode)
End If
End Sub

Private Sub CmbSClientID_GotFocus()
    If CmbSExchangeID = "" Then
       MsgBox "Select Exchange Type.", vbOKOnly + vbInformation, App.Title
       CmbSExchangeID.SetFocus
       Exit Sub
    End If
    
    tmp = CmbSClientID.Text
    Call TableMst_Client(" Where ExchangeTypeID = " & mSExchangeTypeID & " Order By ClientName ")
    CmbSClientID.Clear
    If RsMst_Client.RecordCount = 0 Then
       MsgBox "No record found!!!"
       Exit Sub
    End If
    For I = 1 To RsMst_Client.RecordCount
        CmbSClientID.AddItem RsMst_Client!ClientName & "#" & RsMst_Client!ClientID & "#" & RsMst_Client!ClientIDRow
        RsMst_Client.MoveNext
    Next
    CmbSClientID.Text = tmp
End Sub

Private Sub CmbSClientID_LostFocus()
Dim SplitClient As Variant
If CmbSClientID.Text = "" Then
   Exit Sub
End If

SplitClient = Split(CmbSClientID, "#")
If UBound(SplitClient) <> 2 Then
    MsgBox "Invalid Client.", vbOKOnly + vbInformation, App.Title
    CmbSClientID.SetFocus
    Exit Sub
End If

RsMst_Client.MoveFirst
RsMst_Client.Find "ClientIDRow = " & SplitClient(2)
If RsMst_Client.EOF Then
   MsgBox "Invalid selection!!!"
   CmbSClientID.SetFocus
   Exit Sub
End If
mSClientIDRow = RsMst_Client!ClientIDRow
End Sub

Private Sub CmbSExchangeID_GotFocus()
tmp = CmbSExchangeID.Text
Call TableMst_ExchangeType
CmbSExchangeID.Clear
If RsMst_ExchangeType.RecordCount = 0 Then
   MsgBox "No record found!!!"
   Exit Sub
End If
For I = 1 To RsMst_ExchangeType.RecordCount
    CmbSExchangeID.AddItem RsMst_ExchangeType!ExchangeType
    RsMst_ExchangeType.MoveNext
Next
CmbSExchangeID.Text = tmp
End Sub

Private Sub CmbSExchangeID_LostFocus()
If CmbSExchangeID.Text = "" Then
   CmbSClientID = ""
   Exit Sub
End If
If CmbSExchangeID.Text <> tmp Then
   CmbSClientID = ""
End If
RsMst_ExchangeType.MoveFirst
RsMst_ExchangeType.Find "ExchangeType = '" & CmbSExchangeID.Text & "'"
If RsMst_ExchangeType.EOF Then
   MsgBox "Invalid selection!!!"
   CmbSExchangeID.SetFocus
   Exit Sub
End If
mSExchangeTypeID = RsMst_ExchangeType!ExchangeTypeID
End Sub

Private Sub CmdAddAttachment_Click()
Dim r As Variant
With MSHFlexGrid2
If lblAttachment.Caption = "" Then
   MsgBox "Blank Attachment not allowed!!!", vbOKOnly + vbInformation, App.Title
   Exit Sub
End If

If TxtAttachmentRemarks = "" Then
   MsgBox "Blank Remarks not allowed!!!", vbOKOnly + vbInformation, App.Title
   TxtAttachmentRemarks.SetFocus
   Exit Sub
End If

If CmdAddAttachment.Caption = "Update Attachment" Then
   r = .RowSel
   If DuplicateFileName(GetFileFromPath(lblAttachment.Caption), CInt(r)) = True Then
      MsgBox "Duplicate File Name not allowed."
      Exit Sub
   End If
Else
   r = .Rows - 1
   If DuplicateFileName(GetFileFromPath(lblAttachment.Caption)) = True Then
      MsgBox "Duplicate File Name not allowed."
      Exit Sub
   End If
   .Rows = .Rows + 1
End If
.TextMatrix(r, 0) = ""
.TextMatrix(r, 1) = GetFileFromPath(lblAttachment.Caption)
.TextMatrix(r, 2) = TxtAttachmentRemarks
.TextMatrix(r, 3) = lblAttachment.Caption
Call ClearFrameAttachment
End With
End Sub

Private Function DuplicateFileName(FileName_ As String, Optional Index As Integer) As Boolean
Dim mR, mC As Variant
    With MSHFlexGrid2
       For mR = 1 To .Rows - 2
            If Index = 0 Then
                If LCase(FileName_) = LCase(.TextMatrix(mR, 1)) Then
                    DuplicateFileName = True
                    Exit Function
                End If
            Else
                If LCase(FileName_) = LCase(.TextMatrix(mR, 1)) Then
                   If Index <> 0 And Index <> mR Then
                    DuplicateFileName = True
                    Exit Function
                   End If
                End If
            End If
       Next mR
    End With
DuplicateFileName = False
End Function

Private Sub ClearFrameAttachment()
    lblAttachment.Caption = ""
    TxtAttachmentRemarks = ""
    CmdAddAttachment.Caption = "Add Attachment"
End Sub

Private Sub CmdExcel_Click()
    Gen_mTimeStamp = GetTimeStamp
    Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Txn_ClientFeedbackDocuments.xls")
    Call Xls_Detail_Rpt
End Sub

Public Sub Xls_Detail_Rpt()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Double
Dim mCol As Double
MsgBox ("Wait till next message")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Txn_ClientFeedbackDocuments.xls")
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
MsgBox ("Excel file created!!! ")
oXL.Visible = True
End Sub


Private Sub CmdGetFile_Click()
With CDialog
     .InitDir = App.Path
     .DialogTitle = "Select File"
     .FilterIndex = 1
     .Filter = "*.*"
     .ShowOpen
End With
If CDialog.CancelError Then
   lblAttachment.Caption = ""
Else
   strFilename = CDialog.FileName
   If strFilename = "" Then Exit Sub
   lblAttachment.Refresh
   lblAttachment.Caption = strFilename 'GetFileFromPath(strFilename)
End If
End Sub

Private Sub CmdGo_Click()
Dim Wc As Variant
Wc = ""

If CmbSExchangeID <> "" Then
   Wc = " Where MS.ExchangeTypeID = " & mSExchangeTypeID & ""
End If

If CmbSClientID <> "" Then
   If Wc = "" Then
      Wc = " Where TCFD.ClientIDRow = " & mSClientIDRow
   Else
      Wc = Wc & " And TCFD.ClientIDRow = " & mSClientIDRow
   End If
End If

tmp = " SELECT TCFD.TrancationID,TCFD.TrancationDate,TCFD.ClientIDRow,MS.ClientID,MS.ClientName,ME.ExchangeType,MD.DepositoryName "
tmp = tmp & " FROM Txn_ClientFeedbackDoc As TCFD "
tmp = tmp & " Inner Join Mst_Client As Ms On TCFD.ClientIDRow  = Ms.ClientIDRow "
tmp = tmp & " Inner Join Mst_ExchangeType As ME On MS.ExchangeTypeID  = ME.ExchangeTypeID "
tmp = tmp & " Inner Join Mst_Depository As MD On MS.DepositoryID = MD.DepositoryID "
tmp = tmp & Wc & " Order By TCFD.TrancationID "
Call TmpTable
Call Grid_Head
For iu = 1 To TmpRs.RecordCount
    With MSHFlexGrid1
        .TextMatrix(iu, 0) = IIf(IsNull(TmpRs!TrancationID), "", TmpRs!TrancationID)
        .TextMatrix(iu, 1) = IIf(IsNull(TmpRs!TrancationDate), "", Format(TmpRs!TrancationDate, "DD-MMM-YYYY"))
        .TextMatrix(iu, 2) = IIf(IsNull(TmpRs!ClientID), "", TmpRs!ClientID)
        .TextMatrix(iu, 3) = IIf(IsNull(TmpRs!ClientName), "", TmpRs!ClientName)
        .TextMatrix(iu, 4) = IIf(IsNull(TmpRs!ExchangeType), "", TmpRs!ExchangeType)
        .TextMatrix(iu, 5) = IIf(IsNull(TmpRs!DepositoryName), "", TmpRs!DepositoryName)
        .Rows = .Rows + 1
        TmpRs.MoveNext
    End With
Next
Label22.Refresh
Label22.Caption = TmpRs.RecordCount
End Sub

Private Sub CmdRemoveAttachment_Click()
    I = MSHFlexGrid2.RowSel
    If I > 0 Then
       If lblAttachment.Caption <> "" Then
        'If MSHFlexGrid2.TextMatrix(I, 4) = "" And MSHFlexGrid2.TextMatrix(I, 1) <> "" Then
            For RO = I To MSHFlexGrid2.Rows - 2
                For c = 0 To MSHFlexGrid2.Cols - 1
                    MSHFlexGrid2.TextMatrix(RO, c) = MSHFlexGrid2.TextMatrix(RO + 1, c)
                Next
            Next
            MSHFlexGrid2.Rows = MSHFlexGrid2.Rows - 1
        'End If
       End If
    End If
    Call ClearFrameAttachment
End Sub

Private Sub DeleteCmd_Click()

Dim ans As Variant
On Error GoTo msgError

ans = MsgBox("Do you really want to DELETE this Record??", vbYesNo)
If ans = vbYes Then
   
    ''-- Delete Related Attachment
    tmp = " DELETE FROM Txn_ClientFeedbackDocAttachment WHERE TrancationID = " & Val(TxtTrancationID) & " "
    Call TmpTableSMAttachment
    ''-------------------------
   
   RsTxn_ClientFeedbackDoc.Delete
   RsTxn_ClientFeedbackDoc.Update
   RsTxn_ClientFeedbackDoc.Requery
   If RsTxn_ClientFeedbackDoc.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsTxn_ClientFeedbackDoc.MoveNext
   If RsTxn_ClientFeedbackDoc.EOF() Then
      RsTxn_ClientFeedbackDoc.MoveLast
   End If
   Call Indata
End If
Exit Sub
msgError:
MsgBox "Child record Present "
RsTxn_ClientFeedbackDoc.CancelUpdate
End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Call GButtonLockAD(Me)
   MsgBox "Access denied!!!"
   Exit Sub
End If
  
Call Grid_HeadAttachment
Call Grid_Head
Call TableTxn_ClientFeedbackDoc
If RsTxn_ClientFeedbackDoc.RecordCount = 0 Then
   If UserType = "Auditor" Then
      Call GButtonLock(Me, False)
      SaveCmd.Enabled = False
      SearchCmd.Enabled = False
      Exit Sub
   End If
   Call AddCmd_Click
   Exit Sub
End If
RsTxn_ClientFeedbackDoc.MoveLast
Call GButtonLock(Me, True)
Call Indata
End Sub

Private Sub Indata()
Frame1.Enabled = False

TxtTrancationID = IIf(IsNull(RsTxn_ClientFeedbackDoc!TrancationID), "", RsTxn_ClientFeedbackDoc!TrancationID)
TxtTrancationDate = IIf(IsNull(RsTxn_ClientFeedbackDoc!TrancationDate), "", Format(RsTxn_ClientFeedbackDoc!TrancationDate, "DD-MMM-YYYY"))
mClientIDRow = IIf(IsNull(RsTxn_ClientFeedbackDoc!ClientIDRow), "", RsTxn_ClientFeedbackDoc!ClientIDRow)

CmbExchangeTypeID = "": CmbExchangeTypeID.Visible = False
CmbClientID = "": CmbClientID.Visible = False
TxtExchangeType = "": TxtExchangeType.Visible = True
TxtClientName = "": TxtClientName.Visible = True
TxtDepository = ""
TxtNAVCode = ""
TxtRNAVCode = ""

'tmp = " SELECT MS.ExchangeTypeID,MS.ClientID,MS.ClientIDRow,MS.ClientName,MS.DepositoryID,MD.DepositoryName,MS.NAVCode,MS.RNAVCode "
'tmp = tmp & " FROM Mst_Client AS MS Inner Join Mst_Depository AS MD On MS.DepositoryID = MD.DepositoryID Where MS.ClientIDRow = " & mClientIDRow
tmp = " SELECT MS.ExchangeTypeID,ME.ExchangeType,MS.ClientID,MS.ClientIDRow,MS.ClientName,MS.DepositoryID,MD.DepositoryName,MS.NAVCode,MS.RNAVCode"
tmp = tmp & "  FROM Mst_Client AS MS Inner Join Mst_Depository AS MD On MS.DepositoryID = MD.DepositoryID Inner Join Mst_ExchangeType AS ME On MS.ExchangeTypeID = ME.ExchangeTypeID Where MS.ClientIDRow = " & mClientIDRow

Call TmpTable
If TmpRs.RecordCount > 0 Then
   mExchangeTypeID = IIf(IsNull(TmpRs!ExchangeTypeID), "", TmpRs!ExchangeTypeID)
   TxtExchangeType = IIf(IsNull(TmpRs!ExchangeType), "", TmpRs!ExchangeType)
   CmbExchangeTypeID = TxtExchangeType
   TxtClientName = IIf(IsNull(TmpRs!ClientName), "", TmpRs!ClientName) & "#" & IIf(IsNull(TmpRs!ClientID), "", TmpRs!ClientID) & "#" & IIf(IsNull(TmpRs!ClientIDRow), "", TmpRs!ClientIDRow)
   CmbClientID = TxtClientName
   TxtDepository = IIf(IsNull(TmpRs!DepositoryName), "", TmpRs!DepositoryName)
   TxtNAVCode = IIf(IsNull(TmpRs!NAVCode), "", TmpRs!NAVCode)
   TxtRNAVCode = IIf(IsNull(TmpRs!RNAVCode), "", TmpRs!RNAVCode)
End If

TxtCreated = IIf(IsNull(RsTxn_ClientFeedbackDoc!CreatedBy), "", RsTxn_ClientFeedbackDoc!CreatedBy) & " :- " & IIf(IsNull(RsTxn_ClientFeedbackDoc!CreatedDate), "", RsTxn_ClientFeedbackDoc!CreatedDate)
TxtUpdated = IIf(IsNull(RsTxn_ClientFeedbackDoc!UpdatedBy), "", RsTxn_ClientFeedbackDoc!UpdatedBy) & " :- " & IIf(IsNull(RsTxn_ClientFeedbackDoc!UpdatedDate), "", RsTxn_ClientFeedbackDoc!UpdatedDate)
Call InDataAttachment
Call GButtonLock(Me, True)
    
End Sub

Private Sub InDataAttachment()
Dim mR, mC As Double
Call Grid_HeadAttachment
Call ClearFrameAttachment
Call TableTxn_ClientFeedbackDocAttachment(" Where TrancationID = " & Val(TxtTrancationID) & "")
If RsTxn_ClientFeedbackDocAttachment.RecordCount > 0 Then
   
    If Right(App.Path, 1) = "\" Then
       File_Path = App.Path & "Agreement Attachment\Client Feedback Documents\" & Gen_UserLoginID
    Else
       File_Path = App.Path & "\Agreement Attachment\Client Feedback Documents\" & Gen_UserLoginID
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
   
   For mR = 1 To RsTxn_ClientFeedbackDocAttachment.RecordCount
   With MSHFlexGrid2
        .Rows = .Rows + 1
        .TextMatrix(mR, 0) = mR
        .TextMatrix(mR, 1) = IIf(IsNull(RsTxn_ClientFeedbackDocAttachment!FileName), "", RsTxn_ClientFeedbackDocAttachment!FileName)
        .TextMatrix(mR, 2) = IIf(IsNull(RsTxn_ClientFeedbackDocAttachment!Remarks), "", RsTxn_ClientFeedbackDocAttachment!Remarks)
        .TextMatrix(mR, 3) = IIf(IsNull(RsTxn_ClientFeedbackDocAttachment!FileName), "", File_Path & "\" & RsTxn_ClientFeedbackDocAttachment!FileName)
        Call LoadFileFromDB(.TextMatrix(mR, 3), RsTxn_ClientFeedbackDocAttachment, "File")
        .TextMatrix(mR, 4) = IIf(IsNull(RsTxn_ClientFeedbackDocAttachment!AttachmentID), "", RsTxn_ClientFeedbackDocAttachment!AttachmentID)
        RsTxn_ClientFeedbackDocAttachment.MoveNext
   End With
   Next mR
End If
FrameAttachment.Enabled = False
End Sub


Private Sub Grid_HeadAttachment()
With MSHFlexGrid2
    .Clear
    .Rows = 2
    .Cols = 5
    .Font.Size = 10
    .WordWrap = True
    .TextMatrix(0, 0) = "SR No":                    .ColWidth(0) = 1000
    .TextMatrix(0, 1) = "Attachment File Name":     .ColWidth(1) = 6000
    .TextMatrix(0, 2) = "Remarks":                  .ColWidth(2) = 6000
    .TextMatrix(0, 3) = "File Path":                .ColWidth(3) = 0
    .TextMatrix(0, 4) = "ID":                       .ColWidth(4) = 0
End With
End Sub

Private Sub Grid_Head()
With MSHFlexGrid1
    .Clear
    .Rows = 2
    .Cols = 6
    .Font.Size = 10
    .WordWrap = True
    .TextMatrix(0, 0) = "Trancation ID":            .ColWidth(0) = 1500
    .TextMatrix(0, 1) = "Trancation Date":          .ColWidth(1) = 1600
    .TextMatrix(0, 2) = "Client ID":                .ColWidth(2) = 1200
    .TextMatrix(0, 3) = "Client Name":              .ColWidth(3) = 5000
    .TextMatrix(0, 4) = "Exchange Type":            .ColWidth(4) = 2000
    .TextMatrix(0, 5) = "Depositor Name":           .ColWidth(5) = 3000
End With
End Sub


Private Sub CmbExchangeTypeID_GotFocus()
tmp = CmbExchangeTypeID.Text
Call TableMst_ExchangeType
CmbExchangeTypeID.Clear
If RsMst_ExchangeType.RecordCount = 0 Then
   MsgBox "No record found!!!"
   Exit Sub
End If
For I = 1 To RsMst_ExchangeType.RecordCount
    CmbExchangeTypeID.AddItem RsMst_ExchangeType!ExchangeType
    RsMst_ExchangeType.MoveNext
Next
CmbExchangeTypeID.Text = tmp
End Sub

Private Sub CmbExchangeTypeID_LostFocus()
If CmbExchangeTypeID.Text = "" Then
   CmbClientID = ""
   TxtDepository = ""
   TxtNAVCode = ""
   TxtRNAVCode = ""
   Exit Sub
End If
If CmbExchangeTypeID.Text <> tmp Then
   CmbClientID = ""
   TxtDepository = ""
   TxtNAVCode = ""
   TxtRNAVCode = ""
End If
RsMst_ExchangeType.MoveFirst
RsMst_ExchangeType.Find "ExchangeType = '" & CmbExchangeTypeID.Text & "'"
If RsMst_ExchangeType.EOF Then
   MsgBox "Invalid selection!!!"
   CmbExchangeTypeID.SetFocus
   Exit Sub
End If
mExchangeTypeID = RsMst_ExchangeType!ExchangeTypeID
End Sub

Private Sub LastCmd_Click()
RsTxn_ClientFeedbackDoc.MoveLast
Call Indata
End Sub

Private Sub MSHFlexGrid1_Click()
Dim I As Variant
I = MSHFlexGrid1.RowSel
If I <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(I, 0) = "" Then Exit Sub
RsTxn_ClientFeedbackDoc.MoveFirst
RsTxn_ClientFeedbackDoc.Find "TrancationID = " & MSHFlexGrid1.TextMatrix(I, 0)
If Not RsTxn_ClientFeedbackDoc.EOF Then
   Call Indata
   Frame0.Visible = False
End If
End Sub

Private Sub MSHFlexGrid2_Click()
Dim ans, r As Variant
Dim ID As Integer
r = MSHFlexGrid2.RowSel
If r = MSHFlexGrid2.Rows - 1 Then
    Call ClearFrameAttachment
    Exit Sub
End If
If r < 1 Then Exit Sub
If MSHFlexGrid2.TextMatrix(r, 1) = "" Then Exit Sub

If SaveCmd.Enabled = False Then
   ans = MsgBox("Do You want to view this document ", vbYesNo)
   If ans = vbYes Then
      Shell "Explorer.exe " & File_Path, vbMaximizedFocus
        'If Right(App.Path, 1) = "\" Then
        '   Shell "Explorer.exe " & File_Path, vbMaximizedFocus
        'Else
        '   Shell "Explorer.exe " & App.Path & "\Withdrawal\", vbMaximizedFocus
        'End If
   End If
   Exit Sub
End If

lblAttachment.Caption = MSHFlexGrid2.TextMatrix(r, 3)
TxtAttachmentRemarks = MSHFlexGrid2.TextMatrix(r, 2)
CmdRemoveAttachment.Enabled = True
'If MSHFlexGrid12.TextMatrix(r, 4) <> "" Then
'   CmdRemoveAttachment.Enabled = False
'End If
CmdAddAttachment.Caption = "Update Attachment"
End Sub


Private Sub EditCmd_Click()
    Edit_Flg = "E"
    Call GButtonLock(Me, False)
    'CmbClientID.Visible = True
    'TxtClientName.Visible = False
    'CmbClientID = TxtClientName
    Frame1.Enabled = True
    FrameAttachment.Enabled = True
End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub FirstCmd_Click()
RsTxn_ClientFeedbackDoc.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsTxn_ClientFeedbackDoc.MoveNext
If RsTxn_ClientFeedbackDoc.EOF Then
   RsTxn_ClientFeedbackDoc.MoveLast
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
RsTxn_ClientFeedbackDoc.MovePrevious
If RsTxn_ClientFeedbackDoc.BOF Then
   RsTxn_ClientFeedbackDoc.MoveFirst
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

If CmbExchangeTypeID.Text = "" Then
   MsgBox "Blank Exchange Type not allowed.", vbOKOnly + vbInformation, App.Title
   CmbExchangeTypeID.SetFocus
   Exit Sub
End If

If CmbClientID.Text = "" Then
   MsgBox "Blank Client Name not allowed.", vbOKOnly + vbInformation, App.Title
   CmbClientID.SetFocus
   Exit Sub
End If

If MSHFlexGrid2.Rows < 3 Then
    MsgBox "Blank Client Feedback Documents not allowed.", vbOKOnly + vbInformation, App.Title
    CmdGetFile.SetFocus
    Exit Sub
End If

'If Val(TxtTrancationID) <> 0 Then
'   MsgBox "Truncation ID already available.", vbOKOnly + vbInformation, App.Title
'   Exit Sub
'End If

If Edit_Flg = "A" Then
    RsTxn_ClientFeedbackDoc.AddNew
    RsTxn_ClientFeedbackDoc!TrancationID = GetMaxID("TrancationID", "Txn_ClientFeedbackDoc")
    RsTxn_ClientFeedbackDoc!TrancationDate = Date
    TxtTrancationID = RsTxn_ClientFeedbackDoc!TrancationID
    RsTxn_ClientFeedbackDoc!G_UID = GetGUID
ElseIf Edit_Flg = "E" Then
    RsTxn_ClientFeedbackDoc.MoveFirst
    RsTxn_ClientFeedbackDoc.Find " TrancationID = " & Val(TxtTrancationID)
End If

    RsTxn_ClientFeedbackDoc!ClientIDRow = mClientIDRow
    'RsTxn_ClientFeedbackDoc!Flag = ""
    'RsTxn_ClientFeedbackDoc!Remarks = ""
    If IsNull(RsTxn_ClientFeedbackDoc!CreatedBy) = True Or RsTxn_ClientFeedbackDoc!CreatedBy = "" Then
       RsTxn_ClientFeedbackDoc!CreatedBy = Gen_UserLoginID
       RsTxn_ClientFeedbackDoc!CreatedDate = Now
    Else
       RsTxn_ClientFeedbackDoc!UpdatedBy = Gen_UserLoginID
       RsTxn_ClientFeedbackDoc!UpdatedDate = Now
    End If
    RsTxn_ClientFeedbackDoc.Update
    Call SaveAttachment
    RsTxn_ClientFeedbackDoc.Requery
    RsTxn_ClientFeedbackDoc.Find "TrancationID = " & TxtTrancationID
    Call Indata
    Exit Sub

End Sub

Private Function GetMaxAttachID() As Double
Dim Retval As Double
tmp = "Select max(AttachmentID) from Txn_ClientFeedbackDocAttachment"
Call TmpTableSMAttachment
Retval = IIf(IsNull(TmpRsSM_Attachment.Fields(0)), 0, TmpRsSM_Attachment.Fields(0))
Retval = Retval + 1
GetMaxAttachID = Retval
End Function

Private Sub SaveAttachment()
Dim mAttachmentID As Variant
Dim mR, mC As Double
With MSHFlexGrid2
    If MSHFlexGrid2.Rows <= 1 Then Exit Sub
    Call TableTxn_ClientFeedbackDocAttachment(" Where TrancationID = " & Val(TxtTrancationID) & "")

    For mR = 1 To MSHFlexGrid2.Rows - 2
        If .TextMatrix(mR, 1) = "" Then Exit For
         If .TextMatrix(mR, 0) = "" Then
             If .TextMatrix(mR, 4) = "" Then
                RsTxn_ClientFeedbackDocAttachment.AddNew
                RsTxn_ClientFeedbackDocAttachment!AttachmentID = GetMaxAttachID
                .TextMatrix(mR, 4) = RsTxn_ClientFeedbackDocAttachment!AttachmentID
                RsTxn_ClientFeedbackDocAttachment!G_UID = GetGUID
                RsTxn_ClientFeedbackDocAttachment!TrancationID = Val(TxtTrancationID)
             Else
                RsTxn_ClientFeedbackDocAttachment.MoveFirst
                RsTxn_ClientFeedbackDocAttachment.Find "AttachmentID = " & .TextMatrix(mR, 4) & ""
             End If
            
             mFlag = SaveFileToDB(.TextMatrix(mR, 3), RsTxn_ClientFeedbackDocAttachment, "File")
             mFlag = GetFileFromPath(.TextMatrix(mR, 3))
             If RsTxn_ClientFeedbackDocAttachment!FileName <> mFlag Then
                RsTxn_ClientFeedbackDocAttachment!FileName = TxtTrancationID & "-" & .TextMatrix(mR, 4) & "-" & mFlag
             Else
                If IsNull(RsTxn_ClientFeedbackDocAttachment!FileName) Then
                    RsTxn_ClientFeedbackDocAttachment!FileName = TxtTrancationID & "-" & .TextMatrix(mR, 4) & "-" & mFlag
                Else
                    RsTxn_ClientFeedbackDocAttachment!FileName = mFlag
                End If
             End If
             RsTxn_ClientFeedbackDocAttachment!Remarks = Trim(.TextMatrix(mR, 2))
             If IsNull(RsTxn_ClientFeedbackDocAttachment!CreatedBy) = True Or RsTxn_ClientFeedbackDocAttachment!CreatedBy = "" Then
                RsTxn_ClientFeedbackDocAttachment!CreatedBy = Gen_UserLoginID
                RsTxn_ClientFeedbackDocAttachment!CreatedDate = Now
             Else
                RsTxn_ClientFeedbackDocAttachment!UpdatedBy = Gen_UserLoginID
                RsTxn_ClientFeedbackDocAttachment!UpdatedDate = Now
             End If
             RsTxn_ClientFeedbackDocAttachment.Update
         End If
        mAttachmentID = mAttachmentID & "," & .TextMatrix(mR, 4)
    Next
    
    If mAttachmentID <> "" Then
       mAttachmentID = Right(mAttachmentID, Len(mAttachmentID) - 1)
       If mAttachmentID <> "" Then
            tmp = " DELETE FROM Txn_ClientFeedbackDocAttachment WHERE TrancationID = " & Val(TxtTrancationID) & " AND AttachmentID NOT IN (" & mAttachmentID & ") "
            Call TmpTableSMAttachment
       End If
    End If
    
End With
End Sub

Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   If RsTxn_ClientFeedbackDoc.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   If RsTxn_ClientFeedbackDoc.EOF Then
      RsTxn_ClientFeedbackDoc.MoveFirst
   End If
   Call Indata
   Exit Sub
End If

Frame0.Height = Me.Height - 1000
Frame0.Left = Frame1.Left
Frame0.Top = Frame1.Top
Frame0.Width = Me.Width - 300

MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 1000
RsTxn_ClientFeedbackDoc.MoveFirst
Frame0.Visible = True
End Sub
