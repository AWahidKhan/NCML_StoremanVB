VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmTxn_NewCashReceipt 
   Caption         =   "New Cash Receipt"
   ClientHeight    =   7905
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14775
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame6 
      BackColor       =   &H008080FF&
      Caption         =   "Return Details"
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
      Left            =   4200
      TabIndex        =   136
      Top             =   3120
      Visible         =   0   'False
      Width           =   7095
      Begin VB.CommandButton CmdCRR 
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
         Left            =   6600
         TabIndex        =   186
         TabStop         =   0   'False
         Top             =   1830
         Width           =   375
      End
      Begin VB.ComboBox CmbCRR 
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
         ItemData        =   "FrmTxn_NewCashReceipt.frx":0000
         Left            =   1800
         List            =   "FrmTxn_NewCashReceipt.frx":0010
         Sorted          =   -1  'True
         TabIndex        =   147
         Top             =   1800
         Width           =   4575
      End
      Begin VB.TextBox TxtAdhocAmount 
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
         Left            =   1800
         MaxLength       =   15
         TabIndex        =   148
         Top             =   2400
         Width           =   1920
      End
      Begin VB.TextBox TxtReturnRemark 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   1800
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   146
         Top             =   1200
         Width           =   4560
      End
      Begin VB.TextBox TxtEntryNo 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1800
         MaxLength       =   25
         TabIndex        =   139
         Top             =   285
         Visible         =   0   'False
         Width           =   1560
      End
      Begin VB.CommandButton CmdSaveReturn 
         BackColor       =   &H00E0E0E0&
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
         Left            =   1680
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   138
         Top             =   2880
         Width           =   1440
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cancel"
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
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   137
         Top             =   2880
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker TxtReturnDate 
         Height          =   375
         Left            =   1800
         TabIndex        =   141
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
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
         Format          =   48103425
         CurrentDate     =   39884
      End
      Begin VB.Label Label55 
         BackColor       =   &H008080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Cheque dishonour Reason"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   120
         TabIndex        =   185
         Top             =   1650
         Width           =   1800
      End
      Begin VB.Label Label54 
         BackStyle       =   0  'Transparent
         Caption         =   "Cheque dishonour Charges"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   120
         TabIndex        =   184
         Top             =   2355
         Width           =   1845
      End
      Begin VB.Label Label22 
         BackColor       =   &H008080FF&
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
         Height          =   300
         Left            =   120
         TabIndex        =   149
         Top             =   1290
         Width           =   1260
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Return Date"
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
         TabIndex        =   142
         Top             =   750
         Width           =   1800
      End
      Begin VB.Label Label18 
         BackColor       =   &H008080FF&
         Caption         =   "Entry No"
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
         TabIndex        =   140
         Top             =   315
         Visible         =   0   'False
         Width           =   1260
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H008080FF&
      Caption         =   "Approve Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   4320
      TabIndex        =   192
      Top             =   4200
      Visible         =   0   'False
      Width           =   6375
      Begin VB.TextBox TxtApproveRemarks 
         Appearance      =   0  'Flat
         Height          =   960
         Left            =   1080
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   195
         Top             =   360
         Width           =   5160
      End
      Begin VB.CommandButton CmdApproveSave 
         BackColor       =   &H00E0E0E0&
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
         Left            =   2160
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   194
         Top             =   1440
         Width           =   1440
      End
      Begin VB.CommandButton CmdApproveCancel 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cancel"
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
         Left            =   3600
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   193
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label58 
         BackColor       =   &H008080FF&
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
         Height          =   300
         Left            =   120
         TabIndex        =   196
         Top             =   450
         Width           =   1260
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H008080FF&
      Caption         =   "Reject Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   4320
      TabIndex        =   170
      Top             =   3240
      Visible         =   0   'False
      Width           =   6375
      Begin VB.CommandButton CmdCancelReject 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cancel"
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
         Left            =   3000
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   174
         Top             =   1440
         Width           =   1455
      End
      Begin VB.CommandButton CmdSaveReject 
         BackColor       =   &H00E0E0E0&
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
         Left            =   1560
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   173
         Top             =   1440
         Width           =   1440
      End
      Begin VB.TextBox TxtRejectRemark 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   1560
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   172
         Top             =   840
         Width           =   4680
      End
      Begin MSComCtl2.DTPicker TxtRejectDate 
         Height          =   375
         Left            =   1560
         TabIndex        =   171
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
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
         Format          =   48103425
         CurrentDate     =   39884
      End
      Begin VB.Label Label43 
         BackStyle       =   0  'Transparent
         Caption         =   "Reject Date"
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
         TabIndex        =   176
         Top             =   390
         Width           =   1800
      End
      Begin VB.Label Label40 
         BackColor       =   &H008080FF&
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
         Height          =   300
         Left            =   120
         TabIndex        =   175
         Top             =   930
         Width           =   1260
      End
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   9015
      Left            =   15120
      TabIndex        =   0
      Top             =   120
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   480
      TabIndex        =   25
      Top             =   10080
      Width           =   10935
   End
   Begin VB.TextBox TxtG_UID 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
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
      Left            =   15120
      Locked          =   -1  'True
      TabIndex        =   100
      TabStop         =   0   'False
      Top             =   1320
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Frame Frame4 
      Height          =   9735
      Left            =   30
      TabIndex        =   31
      Top             =   0
      Width           =   15255
      Begin VB.Frame Frame3 
         BackColor       =   &H008080FF&
         Caption         =   "Hologram No "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   5160
         TabIndex        =   32
         Top             =   3960
         Visible         =   0   'False
         Width           =   5055
         Begin VB.CheckBox Check2 
            BackColor       =   &H008080FF&
            Caption         =   "Generate Excel"
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   1440
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.CommandButton CmdCancel 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Cancel"
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
            Left            =   2880
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   840
            Width           =   1455
         End
         Begin VB.CommandButton CmdSaveHologram 
            BackColor       =   &H00E0E0E0&
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
            Left            =   1440
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   840
            Width           =   1440
         End
         Begin VB.TextBox TxtHologramNo 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   1440
            MaxLength       =   25
            TabIndex        =   33
            Top             =   240
            Width           =   2880
         End
         Begin VB.Label Label33 
            BackColor       =   &H008080FF&
            Caption         =   "Hologram No"
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
            TabIndex        =   37
            Top             =   270
            Width           =   1260
         End
      End
      Begin VB.Frame Frame0 
         Height          =   2445
         Left            =   0
         TabIndex        =   38
         Top             =   6240
         Visible         =   0   'False
         Width           =   14880
         Begin VB.ComboBox CmbSReject 
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
            ItemData        =   "FrmTxn_NewCashReceipt.frx":002C
            Left            =   11520
            List            =   "FrmTxn_NewCashReceipt.frx":0036
            Sorted          =   -1  'True
            TabIndex        =   182
            Top             =   1095
            Width           =   1125
         End
         Begin VB.TextBox TxtSClientName 
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
            Left            =   6045
            MaxLength       =   100
            TabIndex        =   179
            TabStop         =   0   'False
            Top             =   1755
            Width           =   5325
         End
         Begin VB.TextBox TxtSInstrumentNo 
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
            Left            =   120
            MaxLength       =   100
            TabIndex        =   160
            TabStop         =   0   'False
            Top             =   1740
            Width           =   2775
         End
         Begin VB.CheckBox ChkBatchNo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Check1"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   5175
            TabIndex        =   158
            Top             =   1793
            Width           =   225
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
            ItemData        =   "FrmTxn_NewCashReceipt.frx":0043
            Left            =   13590
            List            =   "FrmTxn_NewCashReceipt.frx":004D
            Sorted          =   -1  'True
            TabIndex        =   152
            Top             =   375
            Width           =   1245
         End
         Begin VB.ComboBox CmbSAdvance 
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
            ItemData        =   "FrmTxn_NewCashReceipt.frx":005A
            Left            =   12360
            List            =   "FrmTxn_NewCashReceipt.frx":0064
            Sorted          =   -1  'True
            TabIndex        =   150
            Top             =   375
            Width           =   1245
         End
         Begin VB.TextBox TxtSDrawnOn 
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
            Left            =   6165
            MaxLength       =   100
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   1095
            Width           =   2445
         End
         Begin VB.ComboBox CmbSDepositedBank 
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
            ItemData        =   "FrmTxn_NewCashReceipt.frx":0071
            Left            =   2965
            List            =   "FrmTxn_NewCashReceipt.frx":0073
            Sorted          =   -1  'True
            TabIndex        =   47
            Top             =   1095
            Width           =   3165
         End
         Begin VB.ComboBox CmbSDepositBank 
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
            ItemData        =   "FrmTxn_NewCashReceipt.frx":0075
            Left            =   8645
            List            =   "FrmTxn_NewCashReceipt.frx":007F
            Sorted          =   -1  'True
            TabIndex        =   46
            Top             =   1095
            Width           =   2805
         End
         Begin VB.TextBox TxtSCRID 
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
            Left            =   11325
            MaxLength       =   15
            TabIndex        =   45
            TabStop         =   0   'False
            Top             =   375
            Width           =   1005
         End
         Begin VB.TextBox TxtAmountTo 
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
            Left            =   10200
            MaxLength       =   15
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   375
            Width           =   1110
         End
         Begin VB.TextBox TxtAmountFrom 
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
            Left            =   8685
            MaxLength       =   15
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   390
            Width           =   1125
         End
         Begin VB.ComboBox CmbSLocation 
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
            ItemData        =   "FrmTxn_NewCashReceipt.frx":00A6
            Left            =   2845
            List            =   "FrmTxn_NewCashReceipt.frx":00A8
            Sorted          =   -1  'True
            TabIndex        =   42
            Top             =   375
            Width           =   3165
         End
         Begin VB.ComboBox CmbSInstrumentType 
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
            ItemData        =   "FrmTxn_NewCashReceipt.frx":00AA
            Left            =   6075
            List            =   "FrmTxn_NewCashReceipt.frx":00BA
            Sorted          =   -1  'True
            TabIndex        =   41
            Top             =   375
            Width           =   1965
         End
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
            Left            =   12690
            TabIndex        =   40
            Top             =   1095
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
            Left            =   13725
            TabIndex        =   39
            Top             =   1095
            Width           =   1000
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   1575
            Left            =   120
            TabIndex        =   49
            Top             =   2160
            Width           =   10215
            _ExtentX        =   18018
            _ExtentY        =   2778
            _Version        =   393216
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSComCtl2.DTPicker STo 
            Height          =   360
            Left            =   1475
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   375
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   635
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
            Format          =   48103425
            CurrentDate     =   36494
         End
         Begin MSComCtl2.DTPicker SFrom 
            Height          =   360
            Left            =   105
            TabIndex        =   51
            TabStop         =   0   'False
            Top             =   375
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   635
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
            Format          =   48103425
            CurrentDate     =   36494
         End
         Begin MSComCtl2.DTPicker STo1 
            Height          =   360
            Left            =   1550
            TabIndex        =   52
            TabStop         =   0   'False
            Top             =   1095
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   635
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
            Format          =   48103425
            CurrentDate     =   36494
         End
         Begin MSComCtl2.DTPicker SFrom1 
            Height          =   360
            Left            =   135
            TabIndex        =   53
            TabStop         =   0   'False
            Top             =   1095
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   635
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
            Format          =   48103425
            CurrentDate     =   36494
         End
         Begin VB.Label Label53 
            AutoSize        =   -1  'True
            Caption         =   "Reject"
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
            Left            =   11790
            TabIndex        =   183
            Top             =   840
            Width           =   585
         End
         Begin VB.Label Label51 
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
            Left            =   8265
            TabIndex        =   180
            Top             =   1485
            Width           =   1245
         End
         Begin VB.Label Label34 
            Caption         =   "Cheque/DD/RTGS No"
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
            Left            =   427
            TabIndex        =   161
            Top             =   1470
            Width           =   2160
         End
         Begin VB.Label Label32 
            Caption         =   "Approved by FAG"
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
            Left            =   3360
            TabIndex        =   159
            Top             =   1800
            Width           =   1665
         End
         Begin VB.Label Label31 
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
            Height          =   225
            Left            =   13920
            TabIndex        =   153
            Top             =   120
            Width           =   840
         End
         Begin VB.Label Label23 
            Caption         =   "Advance"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   12525
            TabIndex        =   151
            Top             =   120
            Width           =   840
         End
         Begin VB.Label Label49 
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
            Left            =   1800
            TabIndex        =   66
            Top             =   795
            Width           =   855
         End
         Begin VB.Label Label48 
            Caption         =   "Deposit From Date"
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
            TabIndex        =   65
            Top             =   795
            Width           =   1575
         End
         Begin VB.Label Label47 
            Caption         =   "Drawn on Bank"
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
            TabIndex        =   64
            Top             =   795
            Width           =   1245
         End
         Begin VB.Label Label42 
            Caption         =   "Deposited Bank"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   3840
            TabIndex        =   63
            Top             =   795
            Width           =   1680
         End
         Begin VB.Label Label41 
            Caption         =   "Deposit to Bank"
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
            Left            =   9330
            TabIndex        =   62
            Top             =   795
            Width           =   1590
         End
         Begin VB.Label Label19 
            Caption         =   "ID"
            Height          =   255
            Left            =   11760
            TabIndex        =   61
            Top             =   120
            Width           =   405
         End
         Begin VB.Label Label11 
            Caption         =   "To"
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
            Left            =   9840
            TabIndex        =   60
            Top             =   435
            Width           =   255
         End
         Begin VB.Label Label10 
            Caption         =   "From"
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
            Left            =   8100
            TabIndex        =   59
            Top             =   405
            Width           =   645
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
            Left            =   360
            TabIndex        =   58
            Top             =   120
            Width           =   855
         End
         Begin VB.Label Label3 
            Caption         =   "Amount Range"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   9120
            TabIndex        =   57
            Top             =   120
            Width           =   1800
         End
         Begin VB.Label Label36 
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
            Height          =   300
            Left            =   4080
            TabIndex        =   56
            Top             =   120
            Width           =   960
         End
         Begin VB.Label Label38 
            Caption         =   "Instrument Type"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   6600
            TabIndex        =   55
            Top             =   120
            Width           =   1200
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
            Left            =   1800
            TabIndex        =   54
            Top             =   120
            Width           =   855
         End
      End
      Begin VB.Frame Frame1 
         Height          =   3960
         Left            =   120
         TabIndex        =   80
         Top             =   120
         Width           =   14895
         Begin VB.CheckBox ChkAppAM 
            Caption         =   "Approve by AM"
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
            Height          =   255
            Left            =   11520
            TabIndex        =   189
            Top             =   1613
            Visible         =   0   'False
            Width           =   1740
         End
         Begin VB.TextBox TxtChequeReturnChargesAmount 
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
            Left            =   9840
            MaxLength       =   15
            TabIndex        =   10
            Top             =   1560
            Visible         =   0   'False
            Width           =   1605
         End
         Begin VB.TextBox TxtRejectFlag 
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
            Left            =   13260
            Locked          =   -1  'True
            TabIndex        =   168
            TabStop         =   0   'False
            Top             =   3480
            Width           =   1440
         End
         Begin VB.ComboBox CmbBankName 
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
            Left            =   2010
            Sorted          =   -1  'True
            TabIndex        =   13
            Top             =   2970
            Width           =   6165
         End
         Begin VB.TextBox TxtFAGRemarks 
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
            Height          =   465
            Left            =   2010
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   15
            Top             =   3390
            Width           =   9795
         End
         Begin VB.TextBox TxtReturnFlag 
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
            Left            =   10980
            Locked          =   -1  'True
            TabIndex        =   144
            TabStop         =   0   'False
            Top             =   195
            Width           =   1440
         End
         Begin VB.TextBox TxtBankName 
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
            Left            =   2010
            Locked          =   -1  'True
            TabIndex        =   125
            TabStop         =   0   'False
            Top             =   2970
            Visible         =   0   'False
            Width           =   6165
         End
         Begin VB.TextBox TxtSightedDate 
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
            Left            =   13860
            Locked          =   -1  'True
            TabIndex        =   123
            TabStop         =   0   'False
            Top             =   3330
            Visible         =   0   'False
            Width           =   1440
         End
         Begin VB.TextBox TxtDrawnOn 
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
            Left            =   2010
            MaxLength       =   150
            TabIndex        =   9
            Top             =   1545
            Width           =   5550
         End
         Begin VB.TextBox TxtInstrumentNo 
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
            Left            =   9840
            MaxLength       =   25
            TabIndex        =   7
            Top             =   1080
            Width           =   1605
         End
         Begin VB.TextBox TxtCRID 
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
            Left            =   2010
            Locked          =   -1  'True
            TabIndex        =   83
            TabStop         =   0   'False
            Top             =   195
            Width           =   2400
         End
         Begin VB.TextBox TxtPreFix 
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
            Left            =   13260
            Locked          =   -1  'True
            TabIndex        =   82
            TabStop         =   0   'False
            Top             =   195
            Width           =   1440
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
            Left            =   2010
            Sorted          =   -1  'True
            TabIndex        =   3
            Top             =   647
            Width           =   5565
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
            Left            =   7620
            TabIndex        =   81
            TabStop         =   0   'False
            Top             =   660
            Width           =   375
         End
         Begin VB.TextBox TxtInstrumentAmount 
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
            Left            =   6000
            MaxLength       =   15
            TabIndex        =   6
            Top             =   1099
            Width           =   1560
         End
         Begin VB.ComboBox CmbInstrumentType 
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
            ItemData        =   "FrmTxn_NewCashReceipt.frx":00D6
            Left            =   2010
            List            =   "FrmTxn_NewCashReceipt.frx":00E6
            Sorted          =   -1  'True
            TabIndex        =   5
            Top             =   1099
            Width           =   2445
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
            Height          =   435
            Left            =   2010
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   12
            Top             =   2460
            Width           =   12720
         End
         Begin VB.CheckBox ChkAdv 
            Caption         =   "Advance"
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
            Height          =   255
            Left            =   8040
            TabIndex        =   2
            Top             =   240
            Width           =   1140
         End
         Begin MSComCtl2.DTPicker TxtInstrumentDate 
            Height          =   375
            Left            =   13260
            TabIndex        =   8
            Top             =   1095
            Width           =   1440
            _ExtentX        =   2540
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
            Format          =   48103425
            CurrentDate     =   39884
         End
         Begin MSComCtl2.DTPicker TxtCRDate 
            Height          =   375
            Left            =   6120
            TabIndex        =   1
            Top             =   195
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
            Format          =   48103425
            CurrentDate     =   39884
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
            Left            =   2010
            Locked          =   -1  'True
            TabIndex        =   84
            TabStop         =   0   'False
            Top             =   647
            Width           =   5550
         End
         Begin VB.TextBox TxtReceivedFrom 
            Appearance      =   0  'Flat
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2057
               SubFormatType   =   3
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
            Left            =   9840
            MaxLength       =   100
            ScrollBars      =   2  'Vertical
            TabIndex        =   4
            Top             =   615
            Width           =   4860
         End
         Begin MSComCtl2.DTPicker TxtFAGPostingDate 
            Height          =   375
            Left            =   13245
            TabIndex        =   14
            Top             =   2970
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
            Format          =   48103425
            CurrentDate     =   39884
         End
         Begin MSComCtl2.DTPicker TxtDepDate 
            Height          =   375
            Left            =   9915
            TabIndex        =   156
            Top             =   2925
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
            Format          =   48103425
            CurrentDate     =   39884
         End
         Begin VB.TextBox TxtBranchName 
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
            Left            =   2010
            MaxLength       =   150
            TabIndex        =   11
            Top             =   2025
            Width           =   12645
         End
         Begin VB.Label Label52 
            Caption         =   "Cheque Return Charges Amount"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   8040
            TabIndex        =   181
            Top             =   1493
            Visible         =   0   'False
            Width           =   1800
         End
         Begin VB.Label Label39 
            Caption         =   "Reject Flag"
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
            Left            =   11850
            TabIndex        =   169
            Top             =   3540
            Width           =   1095
         End
         Begin VB.Label LblDepoDate 
            Caption         =   "Deposit Date"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   8445
            TabIndex        =   157
            Top             =   2970
            Width           =   1335
         End
         Begin VB.Label Label21 
            Caption         =   "Return Flag"
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
            Left            =   9840
            TabIndex        =   145
            Top             =   255
            Width           =   1095
         End
         Begin VB.Label Label16 
            Caption         =   "Prefix"
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
            Left            =   12600
            TabIndex        =   134
            Top             =   240
            Width           =   975
         End
         Begin VB.Label LblDepoRemark 
            Caption         =   "Deposit Remark "
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
            Left            =   240
            TabIndex        =   127
            Top             =   3495
            Width           =   1695
         End
         Begin VB.Label LblDepoBank 
            Caption         =   "Deposited Bank Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   240
            TabIndex        =   126
            Top             =   2940
            Width           =   1575
         End
         Begin VB.Label LblCRSightedDate 
            Caption         =   "Credit Sighted Date"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   12435
            TabIndex        =   124
            Top             =   3285
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label LblPostDate 
            Caption         =   "FAG Posting Date"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Left            =   11850
            TabIndex        =   122
            Top             =   2925
            Width           =   1455
         End
         Begin VB.Label Label29 
            Caption         =   "Received from"
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
            Left            =   8040
            TabIndex        =   96
            Top             =   675
            Width           =   1935
         End
         Begin VB.Label Label9 
            Caption         =   "Prefix"
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
            Left            =   14640
            TabIndex        =   95
            Top             =   2520
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.Label Label8 
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
            Height          =   285
            Left            =   240
            TabIndex        =   94
            Top             =   685
            Width           =   1440
         End
         Begin VB.Label Label7 
            Caption         =   "Instrument Date"
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
            Left            =   11520
            TabIndex        =   93
            Top             =   1140
            Width           =   1800
         End
         Begin VB.Label Label6 
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
            Height          =   300
            Left            =   4620
            TabIndex        =   92
            Top             =   1125
            Width           =   840
         End
         Begin VB.Label Label5 
            Caption         =   "Instrument No"
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
            Left            =   8040
            TabIndex        =   91
            Top             =   1125
            Width           =   1560
         End
         Begin VB.Label Label4 
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
            Height          =   300
            Left            =   240
            TabIndex        =   90
            Top             =   2520
            Width           =   1800
         End
         Begin VB.Label Label2 
            Caption         =   "Instrument Type"
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
            Left            =   240
            TabIndex        =   89
            Top             =   1129
            Width           =   1800
         End
         Begin VB.Label Label1 
            Caption         =   "Branch Name"
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
            Left            =   225
            TabIndex        =   88
            Top             =   2085
            Width           =   1800
         End
         Begin VB.Label LblFlag 
            Caption         =   "Receipt Date"
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
            Left            =   4620
            TabIndex        =   87
            Top             =   225
            Width           =   1800
         End
         Begin VB.Label LblStateID 
            Caption         =   "Drawn On"
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
            Left            =   225
            TabIndex        =   86
            Top             =   1575
            Width           =   1800
         End
         Begin VB.Label LblLocationID 
            Caption         =   "Receipt No"
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
            Left            =   240
            TabIndex        =   85
            Top             =   225
            Width           =   1200
         End
      End
      Begin VB.Frame ButtonFrm 
         Height          =   1095
         Left            =   135
         TabIndex        =   67
         Top             =   8430
         Width           =   14985
         Begin VB.CommandButton CmdApproveAM 
            Caption         =   "Approve by AM"
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
            Height          =   820
            Left            =   12405
            TabIndex        =   191
            Top             =   165
            Width           =   2475
         End
         Begin VB.CommandButton CmdReject 
            Caption         =   "Reject"
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
            Height          =   400
            Left            =   10650
            TabIndex        =   190
            Top             =   585
            Width           =   1755
         End
         Begin VB.CommandButton CmdReturn 
            Caption         =   "Return"
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
            Height          =   400
            Left            =   10650
            TabIndex        =   143
            Top             =   165
            Width           =   1755
         End
         Begin VB.CommandButton CmdPrintPreview 
            Caption         =   "Print Preview"
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
            Left            =   8880
            TabIndex        =   79
            Top             =   165
            Width           =   1755
         End
         Begin VB.CommandButton CmdPrint 
            Caption         =   "&Print"
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
            Left            =   8880
            TabIndex        =   78
            Top             =   585
            Width           =   1755
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
            Left            =   7125
            TabIndex        =   77
            Top             =   165
            Width           =   1755
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
            Left            =   5370
            TabIndex        =   76
            Top             =   165
            Width           =   1755
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
            Left            =   3615
            TabIndex        =   75
            Top             =   165
            Width           =   1755
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
            Height          =   400
            Left            =   1860
            TabIndex        =   74
            Top             =   165
            Width           =   1755
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
            Left            =   100
            TabIndex        =   73
            Top             =   165
            Width           =   1755
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
            Left            =   7125
            TabIndex        =   72
            Top             =   585
            Width           =   1755
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
            Left            =   5370
            TabIndex        =   71
            Top             =   585
            Width           =   1755
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
            Left            =   3615
            TabIndex        =   70
            Top             =   585
            Width           =   1755
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
            Left            =   1860
            TabIndex        =   69
            Top             =   585
            Width           =   1755
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
            Left            =   100
            TabIndex        =   68
            Top             =   585
            Width           =   1755
         End
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   4350
         Left            =   120
         TabIndex        =   97
         Top             =   4095
         Width           =   14895
         _ExtentX        =   26273
         _ExtentY        =   7673
         _Version        =   393216
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
         TabCaption(0)   =   "Client Details"
         TabPicture(0)   =   "FrmTxn_NewCashReceipt.frx":0102
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "MSHFlexGrid2"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Frame5"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Frame7"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "TI Details"
         TabPicture(1)   =   "FrmTxn_NewCashReceipt.frx":011E
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame2"
         Tab(1).Control(1)=   "MSHFlexGrid3"
         Tab(1).ControlCount=   2
         TabCaption(2)   =   "Log Details"
         TabPicture(2)   =   "FrmTxn_NewCashReceipt.frx":013A
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "TxtRejectedBy"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "TxtFAGUpdated"
         Tab(2).Control(1).Enabled=   0   'False
         Tab(2).Control(2)=   "TxtUpdated"
         Tab(2).Control(2).Enabled=   0   'False
         Tab(2).Control(3)=   "TxtCreated"
         Tab(2).Control(3).Enabled=   0   'False
         Tab(2).Control(4)=   "Label50"
         Tab(2).Control(5)=   "Label15"
         Tab(2).Control(6)=   "Label12"
         Tab(2).Control(7)=   "Label13"
         Tab(2).ControlCount=   8
         Begin VB.TextBox TxtRejectedBy 
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
            Left            =   -73920
            Locked          =   -1  'True
            TabIndex        =   177
            TabStop         =   0   'False
            Top             =   2475
            Width           =   8970
         End
         Begin VB.Frame Frame7 
            Height          =   555
            Left            =   120
            TabIndex        =   162
            Top             =   1185
            Width           =   14655
            Begin VB.CommandButton CmdCDeleteFromList 
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
               Left            =   12795
               TabIndex        =   24
               Top             =   135
               Width           =   1710
            End
            Begin VB.CommandButton CmdCAddToList 
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
               Left            =   11205
               TabIndex        =   23
               Top             =   120
               Width           =   1575
            End
            Begin VB.TextBox TxtCAmt 
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
               Left            =   9555
               MaxLength       =   15
               TabIndex        =   22
               Top             =   135
               Width           =   1560
            End
            Begin VB.TextBox TxtCRemark 
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
               Left            =   915
               MaxLength       =   255
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   20
               Top             =   135
               Width           =   4005
            End
            Begin VB.ComboBox CmbNAVCode 
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
               Left            =   6120
               Sorted          =   -1  'True
               TabIndex        =   21
               Top             =   120
               Width           =   1575
            End
            Begin VB.Label LblCAmt 
               Caption         =   "Excess Amt Recvd"
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
               Left            =   7770
               TabIndex        =   165
               Top             =   165
               Width           =   1800
            End
            Begin VB.Label Label30 
               AutoSize        =   -1  'True
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
               TabIndex        =   164
               Top             =   210
               Width           =   720
            End
            Begin VB.Label Label35 
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
               Left            =   5040
               TabIndex        =   163
               Top             =   180
               Width           =   975
            End
         End
         Begin VB.TextBox TxtFAGUpdated 
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
            Left            =   -73935
            Locked          =   -1  'True
            TabIndex        =   154
            TabStop         =   0   'False
            Top             =   1845
            Width           =   8970
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
            Left            =   -73935
            Locked          =   -1  'True
            TabIndex        =   129
            TabStop         =   0   'False
            Top             =   1185
            Width           =   8970
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
            Left            =   -73935
            Locked          =   -1  'True
            TabIndex        =   128
            TabStop         =   0   'False
            Top             =   600
            Width           =   8970
         End
         Begin VB.Frame Frame5 
            Height          =   900
            Left            =   120
            TabIndex        =   110
            Top             =   330
            Width           =   14655
            Begin VB.CheckBox Check3 
               Caption         =   "Show DP "
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
               Left            =   2445
               TabIndex        =   133
               Top             =   165
               Width           =   1140
            End
            Begin VB.ComboBox CmbDPName 
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
               Left            =   5760
               Sorted          =   -1  'True
               TabIndex        =   132
               Top             =   120
               Visible         =   0   'False
               Width           =   495
            End
            Begin VB.ComboBox CmbCNonNcdexClient 
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
               Left            =   4920
               Sorted          =   -1  'True
               TabIndex        =   18
               Top             =   120
               Visible         =   0   'False
               Width           =   495
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
               ItemData        =   "FrmTxn_NewCashReceipt.frx":0156
               Left            =   120
               List            =   "FrmTxn_NewCashReceipt.frx":0158
               Sorted          =   -1  'True
               TabIndex        =   16
               Top             =   480
               Width           =   3255
            End
            Begin VB.TextBox TxtClientID 
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
               Left            =   3480
               TabIndex        =   17
               Top             =   480
               Width           =   1575
            End
            Begin VB.ComboBox CmbClientDP 
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
               Left            =   8520
               Sorted          =   -1  'True
               TabIndex        =   19
               Top             =   480
               Width           =   4335
            End
            Begin VB.TextBox TxtCClientName 
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
               Left            =   5040
               Locked          =   -1  'True
               TabIndex        =   120
               TabStop         =   0   'False
               Top             =   480
               Width           =   3495
            End
            Begin VB.TextBox TxtCDPID 
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
               Left            =   8520
               Locked          =   -1  'True
               TabIndex        =   117
               TabStop         =   0   'False
               Top             =   480
               Width           =   1215
            End
            Begin VB.TextBox TxtCDPName 
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
               Left            =   9720
               Locked          =   -1  'True
               TabIndex        =   116
               TabStop         =   0   'False
               Top             =   480
               Width           =   3135
            End
            Begin VB.Label Label14 
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
               Left            =   660
               TabIndex        =   121
               Top             =   150
               Width           =   1695
            End
            Begin VB.Label Label24 
               Caption         =   "ClientDPID"
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
               Left            =   9480
               TabIndex        =   115
               Top             =   180
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.Label Label25 
               Caption         =   "DP Name "
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
               Left            =   10920
               TabIndex        =   114
               Top             =   150
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.Label Label26 
               Caption         =   "DP ID"
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
               Left            =   8775
               TabIndex        =   113
               Top             =   180
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.Label Label27 
               Caption         =   "Client Name "
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
               Left            =   6480
               TabIndex        =   112
               Top             =   150
               Width           =   1095
            End
            Begin VB.Label Label28 
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
               Left            =   3780
               TabIndex        =   111
               Top             =   150
               Width           =   975
            End
         End
         Begin VB.Frame Frame2 
            Height          =   1545
            Left            =   -74880
            TabIndex        =   101
            Top             =   360
            Width           =   14655
            Begin VB.TextBox TxtTILocation 
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
               Left            =   7815
               Locked          =   -1  'True
               TabIndex        =   187
               TabStop         =   0   'False
               Top             =   1050
               Width           =   2625
            End
            Begin VB.TextBox TxtTIClient 
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
               Left            =   810
               Locked          =   -1  'True
               TabIndex        =   166
               TabStop         =   0   'False
               Top             =   1065
               Width           =   5865
            End
            Begin VB.TextBox TxtOtherDed 
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
               Left            =   12210
               MaxLength       =   15
               TabIndex        =   28
               Top             =   540
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.TextBox TxtTIDate 
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
               Height          =   360
               Left            =   1770
               Locked          =   -1  'True
               TabIndex        =   119
               TabStop         =   0   'False
               Top             =   540
               Width           =   1650
            End
            Begin VB.TextBox TxtTINo 
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
               Height          =   360
               Left            =   90
               Locked          =   -1  'True
               TabIndex        =   118
               TabStop         =   0   'False
               Top             =   540
               Width           =   1680
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
               Left            =   13905
               TabIndex        =   30
               Top             =   525
               Visible         =   0   'False
               Width           =   1710
            End
            Begin VB.TextBox TxtBalRec 
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
               Height          =   360
               Left            =   8490
               Locked          =   -1  'True
               TabIndex        =   103
               TabStop         =   0   'False
               Top             =   540
               Width           =   1905
            End
            Begin VB.TextBox TxtRecAmt 
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
               Left            =   5040
               MaxLength       =   15
               TabIndex        =   26
               Top             =   540
               Width           =   1680
            End
            Begin VB.CommandButton CmdAdd2List 
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
               Height          =   375
               Left            =   10575
               TabIndex        =   29
               Top             =   1020
               Width           =   1335
            End
            Begin VB.TextBox TxtBalAmount 
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
               Height          =   360
               Left            =   3420
               Locked          =   -1  'True
               TabIndex        =   102
               Top             =   540
               Width           =   1620
            End
            Begin VB.TextBox TxtTDSAmt 
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
               Left            =   6720
               MaxLength       =   15
               TabIndex        =   27
               Top             =   540
               Width           =   1770
            End
            Begin VB.Label Label56 
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
               Left            =   6840
               TabIndex        =   188
               Top             =   1095
               Width           =   900
            End
            Begin VB.Label Label37 
               AutoSize        =   -1  'True
               Caption         =   "Client"
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
               Left            =   120
               TabIndex        =   167
               Top             =   1080
               Width           =   600
            End
            Begin VB.Label Label17 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               Caption         =   "Other Deduction"
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
               TabIndex        =   135
               Top             =   210
               Visible         =   0   'False
               Width           =   1440
            End
            Begin VB.Label LblChequeAmt 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               Caption         =   "Balance Receivable"
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
               Left            =   8505
               TabIndex        =   109
               Top             =   210
               Width           =   1875
            End
            Begin VB.Label LblTDSAmt 
               AutoSize        =   -1  'True
               Caption         =   "Received Amt"
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
               Left            =   5160
               TabIndex        =   108
               Top             =   210
               Width           =   1470
            End
            Begin VB.Label LblPINo 
               Caption         =   "TI No.(Voucher No)"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   90
               TabIndex        =   107
               Top             =   210
               Width           =   1680
            End
            Begin VB.Label LblInvoiceAmount 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               Caption         =   "Balance Amt."
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
               Left            =   3600
               TabIndex        =   106
               Top             =   210
               Width           =   1395
            End
            Begin VB.Label LblPIDate 
               AutoSize        =   -1  'True
               Caption         =   "TI Date"
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
               Left            =   2160
               TabIndex        =   105
               Top             =   210
               Width           =   780
            End
            Begin VB.Label Label46 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               Caption         =   "TDS Deducted"
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
               Left            =   6840
               TabIndex        =   104
               Top             =   210
               Width           =   1560
            End
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
            Height          =   2490
            Left            =   120
            TabIndex        =   98
            Top             =   1725
            Width           =   14670
            _ExtentX        =   25876
            _ExtentY        =   4392
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
            Height          =   2295
            Left            =   -74880
            TabIndex        =   99
            Top             =   1920
            Width           =   14670
            _ExtentX        =   25876
            _ExtentY        =   4048
            _Version        =   393216
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
         Begin VB.Label Label50 
            Caption         =   "Rejected By             :"
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
            Left            =   -74880
            TabIndex        =   178
            Top             =   2400
            Width           =   975
         End
         Begin VB.Label Label15 
            Caption         =   "FAG Updated :"
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
            Left            =   -74895
            TabIndex        =   155
            Top             =   1770
            Width           =   975
         End
         Begin VB.Label Label12 
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
            TabIndex        =   131
            Top             =   1245
            Width           =   975
         End
         Begin VB.Label Label13 
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
            TabIndex        =   130
            Top             =   660
            Width           =   1695
         End
      End
   End
End
Attribute VB_Name = "FrmTxn_NewCashReceipt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mAgentID, mClientIDRow, mSDepoBankID As Variant
Dim mClientNAVCode As String
Dim mClientRNAVCode As String
Dim mChequeReturn As String
Dim mDPNAVCode As String
Dim mExchangeTypeID As Double
Dim mLocationID As Double
Dim Edit_Flg As Variant
Dim mFAGBankID As Double
Dim mPrintFlag As Variant
Dim Wc1 As String
Dim mRType As Variant
Dim mPreMsg As Variant
Dim TOADD, semail, mMsgSubject As Variant
Dim TOCC, mMsg, mGridMsg As Variant
Dim strSplitString() As String
Dim EmpNo_, EmpName_ As String
Dim mTIMsg, mClientMsg, mTIDetailMsg, mPINo, mCRRID, mGenerateAdHoc, mChequeDishonourCharges As Variant
Dim pgval As Double
Dim objPDF As New mjwPDF
Dim mLeft As Double
Dim mTop As Double
Dim mHeight As Double
Dim mRowHeight As Double
Dim mRow As Double
Dim mGlobalLeft As Double
Dim mTxtTop As Double
Dim mTxtLeft As Double
Dim mFontSize As Integer
Dim mTxtFromTop As Double
Dim mPDFFileName As Variant
Dim PntPrv  As Variant

Dim mStrw() As String
Dim mFlagA As Variant
Dim mErrMessage, mStateID, mStateIndiaID, mDPID As Variant
Dim mClientOldNAVCode, mClientOldRNAVCode, mDPOldNAVCode As Variant

Private Sub Grid_Head1()
MSHFlexGrid2.Clear
MSHFlexGrid2.Rows = 3
MSHFlexGrid2.FixedRows = 2
MSHFlexGrid2.Cols = 14

MSHFlexGrid2.TextMatrix(1, 0) = "ExchangeTypeID"
MSHFlexGrid2.TextMatrix(1, 1) = "Exchange Type"
MSHFlexGrid2.TextMatrix(1, 2) = "Client ID"
MSHFlexGrid2.TextMatrix(1, 3) = "Client Name"
MSHFlexGrid2.TextMatrix(1, 4) = "DP ID"
MSHFlexGrid2.TextMatrix(1, 5) = "DP Name"
MSHFlexGrid2.TextMatrix(1, 6) = "Excess Recvd. Amount"
MSHFlexGrid2.TextMatrix(1, 7) = "Remark"
MSHFlexGrid2.TextMatrix(1, 8) = "ClientRowID"
MSHFlexGrid2.TextMatrix(1, 9) = "ShowDP"
MSHFlexGrid2.TextMatrix(1, 10) = "NAVCode"
MSHFlexGrid2.TextMatrix(1, 11) = "RNAVCode"
MSHFlexGrid2.TextMatrix(1, 12) = "AdvNAVCode"
MSHFlexGrid2.TextMatrix(1, 13) = "ChequeReturn"

MSHFlexGrid2.ColWidth(0) = 0
MSHFlexGrid2.ColWidth(1) = 2000
MSHFlexGrid2.ColWidth(2) = 1400
MSHFlexGrid2.ColWidth(3) = 4200
MSHFlexGrid2.ColWidth(4) = 1400
MSHFlexGrid2.ColWidth(5) = 2200
MSHFlexGrid2.ColWidth(6) = 1600
MSHFlexGrid2.ColWidth(7) = 3800
MSHFlexGrid2.ColWidth(8) = 0
MSHFlexGrid2.ColWidth(9) = 0
MSHFlexGrid2.ColWidth(10) = 1500
MSHFlexGrid2.ColWidth(11) = 0
MSHFlexGrid2.ColWidth(12) = 1500
MSHFlexGrid2.ColWidth(13) = 0
End Sub
Private Sub Grid_Head2()

MSHFlexGrid3.Clear
MSHFlexGrid3.Rows = 3
MSHFlexGrid3.Cols = 21
MSHFlexGrid3.FixedRows = 2
MSHFlexGrid3.WordWrap = True

MSHFlexGrid3.TextMatrix(1, 0) = "TI No (Voucher No)"
MSHFlexGrid3.TextMatrix(1, 1) = "TI Date"
MSHFlexGrid3.TextMatrix(1, 2) = "TI Amount"
MSHFlexGrid3.TextMatrix(1, 3) = "Received Amt"
MSHFlexGrid3.TextMatrix(1, 4) = "TDS Deducted"
MSHFlexGrid3.TextMatrix(1, 5) = "Other Deduction"
MSHFlexGrid3.TextMatrix(1, 6) = "Balance Amt"
MSHFlexGrid3.TextMatrix(1, 7) = "Received Amt against this CR"
MSHFlexGrid3.TextMatrix(1, 8) = "TDS Deducted against this CR"
MSHFlexGrid3.TextMatrix(1, 9) = "Other Deduction against this CR"
MSHFlexGrid3.TextMatrix(1, 10) = "Balance Receivable after this CR"
MSHFlexGrid3.TextMatrix(1, 11) = "TIID"
MSHFlexGrid3.TextMatrix(1, 12) = "ClientIDRow/DPID"
MSHFlexGrid3.TextMatrix(1, 13) = "ClientID/DPID"
MSHFlexGrid3.TextMatrix(1, 14) = "Client Name"
MSHFlexGrid3.TextMatrix(1, 15) = "PostInvoice No"
MSHFlexGrid3.TextMatrix(1, 16) = "Entry No"
MSHFlexGrid3.TextMatrix(1, 17) = "DPFlag"
MSHFlexGrid3.TextMatrix(1, 18) = "CR Advance Allocation ID"
MSHFlexGrid3.TextMatrix(1, 19) = "Location"
MSHFlexGrid3.TextMatrix(1, 20) = "TI NAV Code"


MSHFlexGrid3.ColWidth(0) = 1600
MSHFlexGrid3.ColWidth(1) = 1400
MSHFlexGrid3.ColWidth(2) = 1200
MSHFlexGrid3.ColWidth(3) = 1200
MSHFlexGrid3.ColWidth(4) = 1200
MSHFlexGrid3.ColWidth(5) = 1200
MSHFlexGrid3.ColWidth(6) = 1200
MSHFlexGrid3.ColWidth(7) = 1300
MSHFlexGrid3.ColWidth(8) = 1500
MSHFlexGrid3.ColWidth(9) = 1500
MSHFlexGrid3.ColWidth(10) = 1300
MSHFlexGrid3.ColWidth(11) = 0
MSHFlexGrid3.ColWidth(12) = 0
MSHFlexGrid3.ColWidth(13) = 1400
MSHFlexGrid3.ColWidth(14) = 2000
MSHFlexGrid3.ColWidth(17) = 0
MSHFlexGrid3.ColWidth(18) = 1500
MSHFlexGrid3.ColWidth(19) = 1800

MSHFlexGrid3.RowHeight(1) = 800

End Sub

Private Sub AddCmd_Click()

Edit_Flg = "A"
Call GButtonLock(Me, False)
Frame1.Enabled = True
Frame5.Enabled = True
Frame7.Enabled = True
Frame2.Enabled = True
TxtInstrumentAmount.Locked = False
TxtChequeReturnChargesAmount.Locked = False
CmdReturn.Enabled = False
MSHFlexGrid2.Enabled = True
Call Grid_Head1
Call Grid_Head2
'TxtCRDate.Enabled = True
TxtInstrumentDate.Enabled = True
CmbInstrumentType.Visible = True
TxtG_UID = ""
CmbLocationID.Visible = True
TxtLocationID.Visible = False
TxtCRID.Locked = True
TxtInstrumentNo.Locked = False
TxtDrawnOn.Locked = False
TxtBranchName.Locked = False
TxtRemark.Locked = False
mPrintFlag = ""
mFlagA = ""
SSTab1.Tab = 0
mClientOldNAVCode = ""
mClientOldRNAVCode = ""
mDPOldNAVCode = ""

TxtCRID = ""
TxtInstrumentNo = ""
TxtCRDate.Value = Date
TxtInstrumentDate = Date
CmbInstrumentType.Text = ""
CmbLocationID.Text = ""
TxtReceivedFrom.Text = ""
TxtInstrumentAmount = ""
TxtChequeReturnChargesAmount = ""
Call ClearClientTabControl
Call ClearTITabControl
TxtDrawnOn = ""
TxtBranchName = ""
TxtRemark = ""
TxtCreated = ""
TxtUpdated = ""
TxtRejectedBy = ""
'TxtDepositeDate.Text = ""
TxtBankName.Text = ""
'TxtRemark1.Text = ""
TxtSightedDate = ""
TxtReturnFlag = ""
TxtRejectFlag = ""
TxtPreFix = ""
ChkAdv.Value = 0
''Call ChkAdv_Click
CmdPrint.Enabled = False
CmdReject.Enabled = False
CmdPrintPreview.Enabled = False


LblDepoDate.Visible = False
LblPostDate.Visible = False
TxtFAGPostingDate.Visible = False
TxtDepDate.Visible = False
TxtFAGRemarks = ""
If LCase(Gen_UserRole) <> "fag" And LCase(Gen_UserRole) <> "fagadmin" And LCase(Gen_UserRole) <> "power" Then
   TxtBankName = ""
   CmbBankName.Visible = False
   TxtBankName.Visible = True
   TxtFAGRemarks.Locked = True
Else
   CmbBankName.Visible = True
   TxtBankName.Visible = False
   TxtFAGRemarks.Locked = False
End If
CmbBankName.Text = ""
TxtFAGRemarks = ""
ChkAppAM.Visible = False

If RsTxn_NewCashReceipt.RecordCount > 0 Then
    TxtCRDate.SetFocus
End If

End Sub

Private Sub ButtonFrm_Click()
Frame6.Visible = False
Frame8.Visible = False
End Sub

''Private Sub ChkAdv_Click()
''Call Grid_Head1
''Call Grid_Head2
''Call ClearClientTabControl
''Call ClearTITabControl
''If ChkAdv.Value = 1 Then
''   LblCAmt.Visible = True
''   TxtCAmt.Visible = True
''   CmbNAVCode.Visible = True
''   Label35.Visible = True
''   Frame2.Enabled = False
''
''Else
''   LblCAmt.Visible = False
''   TxtCAmt.Visible = False
''   CmbNAVCode.Visible = False
''   Label35.Visible = False
''   Frame2.Enabled = True
''End If
''
''End Sub

Private Sub CmbClientDP_LostFocus()
Dim Str_() As String
Str_ = Split(CmbClientDP.Text, "~")
If UBound(Str_) <> 1 Then
   MsgBox "Invalid selection !!!!! "
   If CmbClientDP.Visible = True Then
      CmbClientDP.SetFocus
   Else
      CmbClientDP.SetFocus
   End If
   Exit Sub
End If
TmpRs1.MoveFirst


TmpRs1.MoveFirst
TmpRs1.Find "DepositoryID  = '" & Str_(0) & "'"
If TmpRs1.EOF Then
   MsgBox "Invalid selection !!!!! "
   CmbClientDP.SetFocus
   Exit Sub
End If

TxtCClientName = TmpRs1!ClientName
TxtCDPID = TmpRs1!DepositoryID
TxtCDPName = TmpRs1!DepositoryName
CmbClientDP.Visible = False
mClientIDRow = TmpRs1!ClientID
mClientOldNAVCode = IIf(IsNull(TmpRs1!NavCode), "", TmpRs1!NavCode)
mChequeReturn = IIf(IsNull(TmpRs1!ChequeReturn), "N", TmpRs1!ChequeReturn)
mClientOldRNAVCode = IIf(IsNull(TmpRs1!RNAVCode), "", TmpRs1!RNAVCode)

'CmbNAVCode.Clear

If mClientRNAVCode <> "" And mClientNAVCode <> "" Then
'   CmbNAVCode.AddItem mClientNAVCode
'   CmbNAVCode.AddItem mClientRNAVCode
ElseIf mClientNAVCode <> "" Then
'   CmbNAVCode.AddItem mClientNAVCode
'   CmbNAVCode.Text = mClientNAVCode
ElseIf mClientRNAVCode <> "" Then
'   CmbNAVCode.AddItem mClientRNAVCode
'   CmbNAVCode.Text = mClientRNAVCode
End If

CmdCAddToList.SetFocus
End Sub

Private Sub CmbCNonNcdexClient_GotFocus()

tmp = CmbCNonNcdexClient.Text
Call TableMst_Client("where ExchangeTypeId = " & mExchangeTypeID)
CmbCNonNcdexClient.Clear
If RsMst_Client.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Client.RecordCount
    CmbCNonNcdexClient.AddItem RsMst_Client!ClientName
    RsMst_Client.MoveNext
Next
CmbCNonNcdexClient.Text = tmp
End Sub

Private Sub CmbCNonNcdexClient_LostFocus()
If CmbCNonNcdexClient.Text = "" Then
   Exit Sub
End If
RsMst_Client.MoveFirst
RsMst_Client.Find "ClientName = '" & CmbCNonNcdexClient.Text & "'"
If RsMst_Client.EOF Then
   MsgBox "Invalid selection "
   CmbCNonNcdexClient.SetFocus
   Exit Sub
End If
mClientIDRow = RsMst_Client!ClientID
mChequeReturn = IIf(IsNull(RsMst_Client!ChequeReturn), "N", RsMst_Client!ChequeReturn)
mClientOldNAVCode = IIf(IsNull(RsMst_Client!NavCode), "", RsMst_Client!NavCode)
mClientOldRNAVCode = IIf(IsNull(RsMst_Client!RNAVCode), "", RsMst_Client!RNAVCode)

'CmbNAVCode.Clear

If mClientRNAVCode <> "" And mClientNAVCode <> "" Then
'   CmbNAVCode.AddItem mClientNAVCode
'   CmbNAVCode.AddItem mClientRNAVCode
ElseIf mClientNAVCode <> "" Then
'   CmbNAVCode.AddItem mClientNAVCode
'   CmbNAVCode.Text = mClientNAVCode
ElseIf mClientRNAVCode <> "" Then
'   CmbNAVCode.AddItem mClientRNAVCode
'   CmbNAVCode.Text = mClientRNAVCode
End If


Dim ans As Variant
ans = GetClientNameByRow(RsMst_Client!ClientID)
TxtClientID = ans(0)
TxtCClientName = ans(1)
TxtCDPID = ans(2)
TxtCDPName = ans(3)



End Sub

Private Sub CmbCRR_GotFocus()

tmp = CmbCRR.Text

Call TableMst_ChequeReturnReason("")

CmbCRR.Clear
If RsMst_ChequeReturnReason.RecordCount = 0 Then
   MsgBox "No record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_ChequeReturnReason.RecordCount
    CmbCRR.AddItem RsMst_ChequeReturnReason!Reason
    RsMst_ChequeReturnReason.MoveNext
Next
CmbCRR.Text = tmp
End Sub

Private Sub CmbCRR_LostFocus()
If CmbCRR.Text = "" Then Exit Sub

RsMst_ChequeReturnReason.MoveFirst
RsMst_ChequeReturnReason.Find "Reason = '" & CmbCRR.Text & "'"
If RsMst_ChequeReturnReason.EOF Then
   MsgBox "Invalid selection "
   CmbCRR.SetFocus
   Exit Sub
End If

mCRRID = RsMst_ChequeReturnReason!CRRID
mGenerateAdHoc = IIf(IsNull(RsMst_ChequeReturnReason!GenerateAdHoc), "N", RsMst_ChequeReturnReason!GenerateAdHoc)
mChequeDishonourCharges = IIf(IsNull(RsMst_ChequeReturnReason!ReturnCharges), 0, RsMst_ChequeReturnReason!ReturnCharges)
If LCase(mGenerateAdHoc) = "n" Then
   TxtAdhocAmount = ""
   TxtAdhocAmount.Locked = True
Else
   TxtAdhocAmount = mChequeDishonourCharges
   TxtAdhocAmount.Locked = False
End If

End Sub

Private Sub CmbNAVCode_GotFocus()

tmp1 = CmbNAVCode.Text

If CmbExchangeTypeID.Text = "" Then
   MsgBox "Select Client!!!"
   CmbExchangeTypeID.SetFocus
   Exit Sub
End If

If Val(mClientIDRow) = 0 And mDPID = 0 Then
   MsgBox "Select Client!!!"
   CmbExchangeTypeID.SetFocus
   Exit Sub
End If


If Gen_WcState = "" Then
   If Check3.Value = 0 Then
      tmp = " Select Distinct GSTNAVCode From("
      tmp = tmp & " Select Distinct a.GSTNavCode 'GSTNAVCode'   from Mst_ClientStateWiseGST a"
      tmp = tmp & " Inner Join Mst_Client b On a.ClientID = b.ClientID"
      tmp = tmp & " Where  a.GSTNavCode is not null and a.GSTNavCode <> '' And a.ClientID = " & mClientIDRow & " "
      tmp = tmp & " Union"
      tmp = tmp & " Select Distinct a.RGSTNavCode 'GSTNAVCode'   from Mst_ClientStateWiseGST a"
      tmp = tmp & " Inner Join Mst_Client b On a.ClientID = b.ClientID"
      tmp = tmp & " Where  a.RGSTNavCode is not null and a.RGSTNavCode <> '' And a.ClientID = " & mClientIDRow & " "
      tmp = tmp & " ) A Group By GSTNAVCode"
   Else
      tmp = "Select Distinct a.GSTNavCode 'GSTNAVCode'   from Mst_DepositoryStateWiseGST a"
      tmp = tmp & " Inner Join Mst_Depository b On a.DepositoryID = b.DepositoryID"
      tmp = tmp & " Where  a.GSTNavCode is not null and a.GSTNavCode <> '' And a.DepositoryID = '" & mDPID & "'"
   End If
   
   Call Tmp4Table
Else
   tmp = "Select Distinct StateIndiaID  from Mst_State Where StateID = " & mStateID & ""
   Call Tmp4Table
   mStateIndiaID = IIf(IsNull(TmpRs4!StateIndiaID), 0, TmpRs4!StateIndiaID)
   
   If Check3.Value = 0 Then
      tmp = " Select Distinct GSTNAVCode From("
      tmp = tmp & " Select Distinct a.GSTNavCode 'GSTNAVCode'   from Mst_ClientStateWiseGST a"
      tmp = tmp & " Inner Join Mst_Client b On a.ClientID = b.ClientID"
      tmp = tmp & " Where  a.GSTNavCode is not null and a.GSTNavCode <> '' And a.ClientID = " & mClientIDRow & " And a.StateIndiaID = " & mStateIndiaID & " "
      tmp = tmp & " Union"
      tmp = tmp & " Select Distinct a.RGSTNavCode 'GSTNAVCode'   from Mst_ClientStateWiseGST a"
      tmp = tmp & " Inner Join Mst_Client b On a.ClientID = b.ClientID"
      tmp = tmp & " Where  a.RGSTNavCode is not null and a.RGSTNavCode <> '' And a.ClientID = " & mClientIDRow & " And a.StateIndiaID = " & mStateIndiaID & " "
      tmp = tmp & " ) A Group By GSTNAVCode"
   
   Else
      tmp = "Select Distinct a.GSTNavCode 'GSTNAVCode'   from Mst_DepositoryStateWiseGST a"
      tmp = tmp & " Inner Join Mst_Depository b On a.DepositoryID = b.DepositoryID"
      tmp = tmp & " Where  a.GSTNavCode is not null and a.GSTNavCode <> '' And a.DepositoryID = '" & mDPID & "' And a.StateIndiaID = " & mStateIndiaID & " "
   
   End If
   
   Call Tmp4Table
   
   
End If

CmbNAVCode.Clear

If TmpRs4.RecordCount = 0 Then
   MsgBox "No record found !!!! "
   Exit Sub
End If

For I = 1 To TmpRs4.RecordCount
    CmbNAVCode.AddItem TmpRs4!GSTNavCode
    TmpRs4.MoveNext
Next

CmbNAVCode.Text = tmp1


End Sub

Private Sub CmbNAVCode_LostFocus()

If CmbNAVCode.Text = "" Then
   Exit Sub
End If

TmpRs4.MoveFirst
TmpRs4.Find "GSTNAVCode = '" & CmbNAVCode.Text & "'"

If TmpRs4.EOF Then
   MsgBox "Invalid GST NAV Code Selected!!!"
   CmbNAVCode.SetFocus
   Exit Sub
End If

If Check3.Value = 0 Then
   mClientNAVCode = IIf(IsNull(TmpRs4!GSTNavCode), "", TmpRs4!GSTNavCode)
   mClientRNAVCode = ""
   mDPNAVCode = ""
Else
   mClientNAVCode = ""
   mClientRNAVCode = ""
   mDPNAVCode = IIf(IsNull(TmpRs4!GSTNavCode), "", TmpRs4!GSTNavCode)
End If
   
End Sub

Private Sub CmbSReject_LostFocus()
If CmbSReject.Text = "" Then Exit Sub
If LCase(CmbSReject.Text) <> "yes" And LCase(CmbSReject.Text) <> "no" Then
   MsgBox "Invalid Selection"
   CmbSReject.SetFocus
   Exit Sub
End If

End Sub

Private Sub CmdAdd2List_Click()
If TxtTINo = "" Then Exit Sub
If TxtRecAmt = "" Then
   MsgBox "Blank 'Received Amount' not allowed!!"
   TxtRecAmt.SetFocus
   Exit Sub
End If

If TxtTDSAmt = "" Then
   MsgBox "Blank 'TDS Deducted' not allowed!!"
   TxtTDSAmt.SetFocus
   Exit Sub
End If

If Val(TxtBalRec.Text) < 0 Then
   MsgBox "Balance Receivable cannot be less than zero."
   TxtRecAmt.SetFocus
   Exit Sub
End If


r = MSHFlexGrid3.RowSel

MSHFlexGrid3.TextMatrix(0, 8) = Val(MSHFlexGrid3.TextMatrix(0, 8)) - Val(MSHFlexGrid3.TextMatrix(r, 8))
MSHFlexGrid3.TextMatrix(0, 7) = Val(MSHFlexGrid3.TextMatrix(0, 7)) - Val(MSHFlexGrid3.TextMatrix(r, 7))
MSHFlexGrid3.TextMatrix(0, 9) = Val(MSHFlexGrid3.TextMatrix(0, 9)) - Val(MSHFlexGrid3.TextMatrix(r, 9))

MSHFlexGrid3.TextMatrix(r, 7) = TxtRecAmt.Text
MSHFlexGrid3.TextMatrix(r, 8) = TxtTDSAmt.Text
MSHFlexGrid3.TextMatrix(r, 9) = TxtOtherDed.Text
MSHFlexGrid3.TextMatrix(r, 10) = TxtBalRec.Text

MSHFlexGrid3.TextMatrix(0, 8) = Val(MSHFlexGrid3.TextMatrix(0, 8)) + Val(MSHFlexGrid3.TextMatrix(r, 8))
MSHFlexGrid3.TextMatrix(0, 7) = Val(MSHFlexGrid3.TextMatrix(0, 7)) + Val(MSHFlexGrid3.TextMatrix(r, 7))
MSHFlexGrid3.TextMatrix(0, 9) = Val(MSHFlexGrid3.TextMatrix(0, 9)) + Val(MSHFlexGrid3.TextMatrix(r, 9))
Call ClearTITabControl
End Sub

Private Sub CmdApproveAM_Click()
If LCase(Gen_UserRole) = "amsandrms" Or LCase(Gen_UserRole) = "headcmg" Then
   If RsTxn_NewCashReceipt!FlagApproval = "AMS" Then
      MsgBox "Cash Receipt already Approve by AMs "
      Exit Sub
   End If
   Frame9.Visible = True
   ButtonFrm.Enabled = False
End If
End Sub

Private Sub CmdApproveCancel_Click()
Frame9.Visible = False
ButtonFrm.Enabled = True
End Sub

Private Sub CmdApproveSave_Click()

If TxtApproveRemarks.Text = "" Then
   MsgBox "Blank Remark not allowed!!"
   TxtApproveRemarks.SetFocus
   Exit Sub
End If

RsTxn_NewCashReceipt!FlagApproval = "AMS"
RsTxn_NewCashReceipt!ApproveDate = Date
RsTxn_NewCashReceipt!ApprovalRemark = TxtApproveRemarks.Text
RsTxn_NewCashReceipt!ApproveBy = Gen_UserLoginID
RsTxn_NewCashReceipt.Update
RsTxn_NewCashReceipt.Requery
RsTxn_NewCashReceipt.MoveFirst
RsTxn_NewCashReceipt.Find "G_UID= '" & TxtG_UID & "'"
If RsTxn_NewCashReceipt.EOF Then
   RsTxn_NewCashReceipt.MoveLast
End If
Frame9.Visible = False
ButtonFrm.Enabled = True
CmdApproveAM.Enabled = False
Call Indata

End Sub

Private Sub CmdCAddToList_Click()

If CmbLocationID.Text = "" Then
   MsgBox "Blank Location not allowed!"
   CmbLocationID.SetFocus
   Exit Sub
End If

If CmbExchangeTypeID.Text = "" Then
   MsgBox "Blank Exchange Type not allowed!"
   CmbExchangeTypeID.SetFocus
   Exit Sub
End If

If Check3.Value = 0 Then
   If CmbClientDP.Visible = True Then
      MsgBox "Invalid selection !!! "
      Exit Sub
   End If
   If CmbCNonNcdexClient.Visible = True Then
      If CmbCNonNcdexClient.Text = "" Then
         MsgBox "Invalid selection !!! "
         CmbCNonNcdexClient.SetFocus
         Exit Sub
      End If
   End If
   If TxtClientID = "" Then
      MsgBox "Invalid Client ID !!! "
      Exit Sub
   End If
Else
   If CmbDPName.Visible = True Then
      If CmbDPName.Text = "" Then
         MsgBox "Blank DP Name Not allowed !!! "
         CmbDPName.SetFocus
         Exit Sub
      End If
   End If
End If

If CmbNAVCode.Text = "" Then
   MsgBox "Please select Client NAV Code!! "
   CmbNAVCode.SetFocus
   Exit Sub
End If


'If Check3.Value = 0 Then
'   If LCase(CmbNAVCode.Text) <> LCase(mClientNAVCode) And LCase(CmbNAVCode.Text) <> LCase(mClientRNAVCode) Then
'      MsgBox "Please select Client NAV Code from drop down only!! "
'      CmbNAVCode.SetFocus
'      Exit Sub
'   End If
'Else
'   If LCase(CmbNAVCode.Text) <> LCase(mDPNAVCode) Then
'      MsgBox "Please select Client NAV Code from drop down only!! "
'      CmbNAVCode.SetFocus
'      Exit Sub
'   End If
'End If

If CheckDuplicateClient = True Then
   MsgBox "Duplicate Client ID  not allowd !!!"
   TxtClientID.SetFocus
   Exit Sub
End If

If CmdCAddToList.Caption = "Update List" Then
   r = MSHFlexGrid2.RowSel
'   If MSHFlexGrid2.TextMatrix(r, 12) <> CmbNAVCode.Text Then
'      Call DeleteFromTIDetail
'   End If
'   Call UpdateTIDetails
   MSHFlexGrid2.TextMatrix(0, 6) = Format(Val(MSHFlexGrid2.TextMatrix(0, 6)) - Val(MSHFlexGrid2.TextMatrix(r, 6)))
      
Else
  r = MSHFlexGrid2.Rows - 1
  MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
  Call AddTIDetails
End If

MSHFlexGrid2.TextMatrix(r, 0) = mExchangeTypeID
MSHFlexGrid2.TextMatrix(r, 1) = CmbExchangeTypeID.Text
If Check3.Value = 0 Then
   MSHFlexGrid2.TextMatrix(r, 2) = TxtClientID
   MSHFlexGrid2.TextMatrix(r, 3) = TxtCClientName
   MSHFlexGrid2.TextMatrix(r, 4) = TxtCDPID
   MSHFlexGrid2.TextMatrix(r, 5) = TxtCDPName
   MSHFlexGrid2.TextMatrix(r, 8) = mClientIDRow
   MSHFlexGrid2.TextMatrix(r, 10) = mClientNAVCode
   MSHFlexGrid2.TextMatrix(r, 11) = mClientRNAVCode
Else
   MSHFlexGrid2.TextMatrix(r, 2) = ""
   MSHFlexGrid2.TextMatrix(r, 3) = ""
   MSHFlexGrid2.TextMatrix(r, 4) = TxtCDPID
   MSHFlexGrid2.TextMatrix(r, 5) = TxtCDPName
   MSHFlexGrid2.TextMatrix(r, 8) = TxtCDPID
   MSHFlexGrid2.TextMatrix(r, 10) = mDPNAVCode
End If
MSHFlexGrid2.TextMatrix(r, 6) = Format(Val(TxtCAmt), "######0.00")
MSHFlexGrid2.TextMatrix(0, 6) = Val(MSHFlexGrid2.TextMatrix(0, 6)) + Val(MSHFlexGrid2.TextMatrix(r, 6))
MSHFlexGrid2.TextMatrix(0, 6) = Format(MSHFlexGrid2.TextMatrix(0, 6), "######0.00")
MSHFlexGrid2.TextMatrix(r, 7) = TxtCRemark.Text
MSHFlexGrid2.TextMatrix(r, 9) = Check3.Value
MSHFlexGrid2.TextMatrix(r, 12) = CmbNAVCode.Text
MSHFlexGrid2.TextMatrix(r, 13) = mChequeReturn
If mChequeReturn = "Y" Then
    'If LCase(CmbInstrumentType.Text) <> "cash" And LCase(CmbInstrumentType.Text) <> "dd" Then
    If LCase(CmbInstrumentType.Text) <> "cash" And LCase(CmbInstrumentType.Text) <> "dd" And LCase(CmbInstrumentType.Text) <> "rtgs" Then
        MsgBox "This Client cheque was dishonoured in past, hence we do not accept cheque from this client any more!!!"
        CmbInstrumentType.Text = "Cash"
        CmbInstrumentType_LostFocus
        CmbInstrumentType.Text = ""
'        TxtInstrumentAmount.Text = ""
    End If
End If
Call ClearClientTabControl

End Sub
Private Sub UpdateTIDetails()

'wc1 = ""
'If LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "admin" And LCase(Gen_UserRole) <> "superadmin" And LCase(Gen_UserRole) <> "fag" And LCase(Gen_UserRole) <> "fag admin" And LCase(Gen_UserRole) <> "head cmg" Then
'   wc1 = " And (LocationID=" & mLocationID & " or LocationID is null)"
'End If
'
'If Check3.Value = 0 Then
'   If mClientRNAVCode <> "" Then
'      wc1 = wc1 & " And (TI.NAVCode = '" & mClientNAVCode & "' or TI.NAVCode = '" & mClientRNAVCode & "')"
'   Else
'      wc1 = wc1 & " And TI.NAVCode = '" & mClientNAVCode & "'"
'   End If
'Else
'   wc1 = wc1 & " And TI.NAVCode='" & mDPNAVCode & "'"
'End If
'
'Call GetClientsRemInvoices(wc1, TxtG_UID, "")
'
'''tmp = " Select TCRT.TIID,Sum(TCRT.RecAmount) as RecAmount,Sum(RecTDS) as RecTDSAmount,Sum(RecOtherDed)as RecOtherDed  From Txn_NewCashReceiptTIDetail TCRT"
'''tmp = tmp & " Inner Join Txn_NewCashReceipt TCR on TCR.CRID= TCRT.CRID and TCR.SM_Prefix=TCRT.SM_Prefix"
'''tmp = tmp & " Where TCR.G_UID<>'" & TxtG_UID & "' And TCR.ReturnFlag<>'Y'"
'''tmp = tmp & " Group By  TCRT.TIID"
'''
'''Call Tmp4Table
'
'If TmpRs.RecordCount > 0 Then
'   For gr = 1 To TmpRs.RecordCount
'       If CheckDuplicateTINo(TmpRs!EntryNo) = False Then
'          rn = MSHFlexGrid3.Rows - 1
'          MSHFlexGrid3.TextMatrix(rn, 0) = TmpRs!TINo
'          MSHFlexGrid3.TextMatrix(rn, 1) = IIf(IsNull(TmpRs!TIDate), "", Format(TmpRs!TIDate, "dd-MMM-yyyy"))
'          MSHFlexGrid3.TextMatrix(rn, 2) = IIf(IsNull(TmpRs!TIAmount), 0, TmpRs!TIAmount)
'          MSHFlexGrid3.TextMatrix(rn, 3) = IIf(IsNull(TmpRs!RecAmount), 0, TmpRs!RecAmount)
'          MSHFlexGrid3.TextMatrix(rn, 4) = IIf(IsNull(TmpRs!RecTDS), 0, TmpRs!RecTDS)
'          MSHFlexGrid3.TextMatrix(rn, 5) = IIf(IsNull(TmpRs!RecOtherDed), 0, TmpRs!RecOtherDed)
'          MSHFlexGrid3.TextMatrix(rn, 6) = Format(Val(MSHFlexGrid3.TextMatrix(rn, 2)) - (Val(MSHFlexGrid3.TextMatrix(rn, 3)) + Val(MSHFlexGrid3.TextMatrix(rn, 4)) + Val(MSHFlexGrid3.TextMatrix(rn, 5))), "#####0.00")
'          MSHFlexGrid3.TextMatrix(rn, 11) = TmpRs!TIID
'          If Check3.Value = 0 Then
'             MSHFlexGrid3.TextMatrix(rn, 12) = mClientIDRow
'             MSHFlexGrid3.TextMatrix(rn, 13) = TxtClientID
'             MSHFlexGrid3.TextMatrix(rn, 14) = TxtCClientName
'             MSHFlexGrid3.TextMatrix(rn, 17) = "N"
'          Else
'             MSHFlexGrid3.TextMatrix(rn, 12) = TxtCDPID
'             MSHFlexGrid3.TextMatrix(rn, 13) = TxtCDPID
'             MSHFlexGrid3.TextMatrix(rn, 14) = TxtCDPName
'             MSHFlexGrid3.TextMatrix(rn, 17) = "Y"
'          End If
'
'          MSHFlexGrid3.TextMatrix(rn, 15) = IIf(IsNull(TmpRs!PostInvoiceNo), 0, TmpRs!PostInvoiceNo)
'          MSHFlexGrid3.TextMatrix(rn, 16) = IIf(IsNull(TmpRs!EntryNo), 0, TmpRs!EntryNo)
''          MSHFlexGrid3.TextMatrix(rn, 18) = IIf(IsNull(TmpRs!CRAAID), 0, TmpRs!CRAAID)
'          MSHFlexGrid3.Rows = MSHFlexGrid3.Rows + 1
'       End If
'       TmpRs.MoveNext
'   Next
'End If


End Sub
Private Sub AddTIDetails()

'''tmp = " Select TINo,TIDate,TIAmount,TIID,TI.ClientIDRow,ClientID,ClientName,TI.DepositoryID,DepositoryName,TI.PostInvoiceNo,TI.EntryNo From  Txn_TIDetails TI "
'''tmp = tmp & " Left Join Mst_Client MC on TI.ClientIDRow=MC.ClientIDRow"
'''tmp = tmp & " Left Join Mst_Depository MD on TI.DepositoryID=MD.DepositoryID"
'''tmp = tmp & " Where  TINo <> '' And TIAmount-(select isnull(Sum(RecAmount)+Sum(RecTDS)+Sum(RecOtherDed),0) From Txn_NewCashReceiptTIDetail TCRT"
'''tmp = tmp & " Inner Join Txn_NewCashReceipt TCR on TCR.CRID= TCRT.CRID and TCR.SM_Prefix=TCRT.SM_Prefix"
'''tmp = tmp & " Where TCRT.TIID=TI.TIID and TCR.G_UID<>'" & TxtG_UID & "' And TCR.ReturnFlag<>'Y')>0 "
'''
'''If LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "admin" And LCase(Gen_UserRole) <> "superadmin" And LCase(Gen_UserRole) <> "fag" And LCase(Gen_UserRole) <> "head cmg" Then
'''   tmp = tmp & " And (LocationID=" & mLocationID & " or LocationID is null)"
''''Else
''''   tmp = tmp & " And (LocationID=" & mLocationID & " or LocationID is null)"
'''End If
'''
'''If Check3.Value = 0 Then
'''   If mClientRNAVCode <> "" Then
'''      tmp = tmp & " And (TI.NAVCode = '" & mClientNAVCode & "' or TI.NAVCode = '" & mClientRNAVCode & "')"
'''   Else
'''      tmp = tmp & " And TI.NAVCode = '" & mClientNAVCode & "'"
'''   End If
'''Else
'''   tmp = tmp & " And TI.NAVCode='" & mDPNAVCode & "'"
'''End If
'''
'''Call Tmp3Table

Wc1 = ""
If LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "admin" And LCase(Gen_UserRole) <> "superadmin" And LCase(Gen_UserRole) <> "fag" And LCase(Gen_UserRole) <> "fagadmin" And LCase(Gen_UserRole) <> "headcmg" Then
   Wc1 = " And (TI.LocationID=" & mLocationID & " or TI.LocationID is null)"
End If

If Check3.Value = 0 Then
   If mClientRNAVCode <> "" Then
      Wc1 = Wc1 & " And (TI.NAVCode = '" & mClientNAVCode & "' or TI.NAVCode = '" & mClientRNAVCode & "' Or TI.NAVCode = '" & mClientOldNAVCode & "' Or TI.NAVCode = '" & mClientOldRNAVCode & "' )"
   Else
      If mClientOldRNAVCode = "" Then
         Wc1 = Wc1 & " And (TI.NAVCode = '" & mClientNAVCode & "' Or TI.NAVCode = '" & mClientOldNAVCode & "') "
      Else
         Wc1 = Wc1 & " And (TI.NAVCode = '" & mClientNAVCode & "' Or TI.NAVCode = '" & mClientOldNAVCode & "' Or TI.NAVCode = '" & mClientOldRNAVCode & "') "
      End If
   End If
Else
   Wc1 = Wc1 & " And (TI.NAVCode='" & mDPNAVCode & "' Or TI.NAVCode = '" & mDPOldNAVCode & "') "
End If

Call GetClientsRemInvoices(Wc1, TxtG_UID, "")

''tmp = " Select TCRT.TIID,Sum(TCRT.RecAmount) as RecAmount,Sum(RecTDS) as RecTDSAmount,Sum(RecOtherDed)as RecOtherDed  From Txn_NewCashReceiptTIDetail TCRT"
''tmp = tmp & " Inner Join Txn_NewCashReceipt TCR on TCR.CRID= TCRT.CRID and TCR.SM_Prefix=TCRT.SM_Prefix"
''tmp = tmp & " Where TCR.G_UID<>'" & TxtG_UID & "' And TCR.ReturnFlag<>'Y'"
''tmp = tmp & " Group By  TCRT.TIID"
''
''Call Tmp4Table

If TmpRs.RecordCount > 0 Then
   For gr = 1 To TmpRs.RecordCount
       If CheckDuplicateTINo(TmpRs!EntryNo) = False Then
          rn = MSHFlexGrid3.Rows - 1
          MSHFlexGrid3.TextMatrix(rn, 0) = TmpRs!TINo
          MSHFlexGrid3.TextMatrix(rn, 1) = IIf(IsNull(TmpRs!TIDate), "", Format(TmpRs!TIDate, "dd-MMM-yyyy"))
          MSHFlexGrid3.TextMatrix(rn, 2) = IIf(IsNull(TmpRs!TIAmount), 0, TmpRs!TIAmount)
          MSHFlexGrid3.TextMatrix(rn, 3) = IIf(IsNull(TmpRs!RecAmount), 0, TmpRs!RecAmount)
          MSHFlexGrid3.TextMatrix(rn, 3) = Val(MSHFlexGrid3.TextMatrix(rn, 3)) + IIf(IsNull(TmpRs!AdjAmountFAG), 0, TmpRs!AdjAmountFAG)
          MSHFlexGrid3.TextMatrix(rn, 4) = IIf(IsNull(TmpRs!RecTDS), 0, TmpRs!RecTDS)
          MSHFlexGrid3.TextMatrix(rn, 5) = IIf(IsNull(TmpRs!RecOtherDed), 0, TmpRs!RecOtherDed)
          MSHFlexGrid3.TextMatrix(rn, 6) = Format(Val(MSHFlexGrid3.TextMatrix(rn, 2)) - (Val(MSHFlexGrid3.TextMatrix(rn, 3)) + Val(MSHFlexGrid3.TextMatrix(rn, 4)) + Val(MSHFlexGrid3.TextMatrix(rn, 5))), "#####0.00")
          MSHFlexGrid3.TextMatrix(rn, 11) = TmpRs!TIID
          If Check3.Value = 0 Then
             MSHFlexGrid3.TextMatrix(rn, 12) = mClientIDRow
             MSHFlexGrid3.TextMatrix(rn, 13) = TxtClientID
             MSHFlexGrid3.TextMatrix(rn, 14) = TxtCClientName
             MSHFlexGrid3.TextMatrix(rn, 17) = "N"
          Else
             MSHFlexGrid3.TextMatrix(rn, 12) = TxtCDPID
             MSHFlexGrid3.TextMatrix(rn, 13) = TxtCDPID
             MSHFlexGrid3.TextMatrix(rn, 14) = TxtCDPName
             MSHFlexGrid3.TextMatrix(rn, 17) = "Y"
          End If
        
          MSHFlexGrid3.TextMatrix(rn, 15) = IIf(IsNull(TmpRs!PostInvoiceNo), 0, TmpRs!PostInvoiceNo)
          MSHFlexGrid3.TextMatrix(rn, 16) = IIf(IsNull(TmpRs!EntryNo), 0, TmpRs!EntryNo)
'          MSHFlexGrid3.TextMatrix(rn, 18) = IIf(IsNull(TmpRs!CRAAID), 0, TmpRs!CRAAID)
          MSHFlexGrid3.TextMatrix(rn, 19) = IIf(IsNull(TmpRs!Locationname), "", TmpRs!Locationname)
          MSHFlexGrid3.TextMatrix(rn, 20) = IIf(IsNull(TmpRs!NavCode), "", TmpRs!NavCode)
          MSHFlexGrid3.Rows = MSHFlexGrid3.Rows + 1
       End If
       TmpRs.MoveNext
   Next
End If
End Sub

Private Sub DeleteFromTIDetail()
I = MSHFlexGrid2.RowSel
Dim mDelRows As Variant
If I > 0 Then
   If MSHFlexGrid2.TextMatrix(I, 1) <> "" Then
      mDelRows = 0
Restart:
        For r = 2 To MSHFlexGrid3.Rows - 1
            If MSHFlexGrid3.TextMatrix(r, 12) = MSHFlexGrid2.TextMatrix(I, 8) Then
               MSHFlexGrid3.TextMatrix(0, 8) = Val(MSHFlexGrid3.TextMatrix(0, 8)) - Val(MSHFlexGrid3.TextMatrix(r, 8))
               MSHFlexGrid3.TextMatrix(0, 7) = Val(MSHFlexGrid3.TextMatrix(0, 7)) - Val(MSHFlexGrid3.TextMatrix(r, 7))
               MSHFlexGrid3.TextMatrix(0, 9) = Val(MSHFlexGrid3.TextMatrix(0, 9)) - Val(MSHFlexGrid3.TextMatrix(r, 9))
               For RO = r To MSHFlexGrid3.Rows - 2
                   For C = 0 To MSHFlexGrid3.Cols - 1
                       MSHFlexGrid3.TextMatrix(RO, C) = MSHFlexGrid3.TextMatrix(RO + 1, C)
                   Next
               Next
               MSHFlexGrid3.Rows = MSHFlexGrid3.Rows - 1
               GoTo Restart
            End If
        Next
   End If
End If

End Sub


Private Sub CmdCDeleteFromList_Click()
Call DeleteFromTIDetail
I = MSHFlexGrid2.RowSel
If I > 1 Then
   If MSHFlexGrid2.TextMatrix(I, 0) <> "" Then
      ''==================================================
      'If ChkAdv.Value = 1 Then
      '   MSHFlexGrid2.TextMatrix(0, 6) = Val(MSHFlexGrid2.TextMatrix(0, 6)) - Val(MSHFlexGrid2.TextMatrix(I, 6))
      'End If
      MSHFlexGrid2.TextMatrix(0, 6) = Val(MSHFlexGrid2.TextMatrix(0, 6)) - Val(MSHFlexGrid2.TextMatrix(I, 6))
      ''==================================================
      For RO = I To MSHFlexGrid2.Rows - 2
          For C = 0 To MSHFlexGrid2.Cols - 1
              MSHFlexGrid2.TextMatrix(RO, C) = MSHFlexGrid2.TextMatrix(RO + 1, C)
          Next
      Next
      MSHFlexGrid2.Rows = MSHFlexGrid2.Rows - 1
      MSHFlexGrid2.row = 0
   End If
End If
Call ClearClientTabControl
End Sub

Function CheckDuplicateClient() As Boolean
Dim Retval As Boolean
Retval = False
If MSHFlexGrid2.Rows > 3 Then
   For u = 1 To MSHFlexGrid2.Rows - 1
       'If MSHFlexGrid2.TextMatrix(u, 2) = TxtClientID And MSHFlexGrid2.TextMatrix(u, 4) = TxtCDPID Then 'And LCase(MSHFlexGrid2.TextMatrix(u, 12)) = LCase(CmbNAVCode.Text) Then
       If MSHFlexGrid2.TextMatrix(u, 2) = TxtClientID And MSHFlexGrid2.TextMatrix(u, 4) = TxtCDPID And LCase(MSHFlexGrid2.TextMatrix(u, 12)) = LCase(CmbNAVCode.Text) Then
          If LCase(CmdCAddToList.Caption) = "add to list" Then
             Retval = True
             Exit For
          Else
             If u <> MSHFlexGrid2.RowSel Then
                Retval = True
                Exit For
             End If
          End If
       End If
   Next
End If
CheckDuplicateClient = Retval
End Function




Private Sub CmdCRR_Click()
FrmMst_ChequeReturnReason.Show
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Txn_NewCashReceipt.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Txn_NewCashReceipt.xls")
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
Private Sub CmdReject_Click()

If LCase(Gen_UserRole) <> "fag" And LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "fagadmin" And LCase(Gen_UserRole) <> "headcmg" Then
   MsgBox "Only FAG/Head CMG user can reject the Cash Receipt!!!"
   Exit Sub
End If
    
If RsTxn_NewCashReceipt!BatchID <> 0 Then
   MsgBox "Batch is generated already. You can't reject this receipt!!!! "
   Exit Sub
End If

If IsNull(RsTxn_NewCashReceipt!FAGPostingDate) = False Then
   tmp = "Select * from Txn_NewCashReceipt Where SM_Prefix='" & TxtPreFix & "' And NewCRID = " & TxtCRID & " "
   Call Tmp4Table
   
   If TmpRs4.RecordCount > 0 Then
      If TmpRs4!BatchID <> 0 Then
         MsgBox "Receipt is approved by FAG and Batch ID " & TmpRs4!BatchID & " is generated. You can't reject this receipt!!!! "
         Exit Sub
      End If
   End If
End If

ButtonFrm.Enabled = False
Frame8.Visible = True
TxtRejectDate.Value = Date
TxtRejectRemark = ""
TxtRejectDate.SetFocus

End Sub

Private Sub CmdReturn_Click()
Frame6.Visible = True
TxtEntryNo = ""
TxtReturnDate.Value = Date
TxtReturnRemark = ""
'ChkGenerateAdHocInvoice.Value = vbUnchecked
CmbCRR.Text = ""
TxtAdhocAmount = ""
TxtReturnDate.SetFocus
End Sub

Private Sub CmdSaveReject_Click()

If TxtRejectDate.Value > Date Then
   MsgBox "Future Date not allowed!!"
   TxtRejectDate.SetFocus
   Exit Sub
End If

If Format(Gen_BlockingDate, "yyyy-mm") >= Format(TxtRejectDate.Value, "yyyy-mm") Then
   MsgBox "Cash Receipt Reject Date allowed for next month of " & Format(Gen_BlockingDate, "MMM-yyyy") & "!!! "
   TxtRejectDate.SetFocus
   Exit Sub
End If

If TxtRejectRemark.Text = "" Then
   MsgBox "Blank Remark not allowed!!"
   TxtRejectRemark.SetFocus
   Exit Sub
End If

Call DeleteTIDetail

RsTxn_NewCashReceipt!RejectFlag = "Y"
RsTxn_NewCashReceipt!RejectDate = TxtRejectDate.Value
RsTxn_NewCashReceipt!RejectRemark = TxtRejectRemark.Text
RsTxn_NewCashReceipt!RejectedBy = Gen_UserLoginID
RsTxn_NewCashReceipt.Update
RsTxn_NewCashReceipt.Requery
RsTxn_NewCashReceipt.MoveFirst
RsTxn_NewCashReceipt.Find "G_UID= '" & TxtG_UID & "'"
If RsTxn_NewCashReceipt.EOF Then
   RsTxn_NewCashReceipt.MoveLast
End If
Frame8.Visible = False
ButtonFrm.Enabled = True

mRType = "Reject"
Call Gen_Email

Call Indata


End Sub
Private Sub Gen_Email()


mMsg = ""
'mMsg = "<tr>Pledge Details (From Store-Man)</tr>"
mPreMsg = ""

If LCase(mRType) = "return" Then
   mPreMsg = "<P> Please find below the " & "cheque returned by bank " & "details pertaining to your location and take appropiate corrective steps."
Else
   mPreMsg = "<P> Please find below the " & "cheque rejected by FAG " & "details pertaining to your location and take appropiate corrective steps."
End If

mPreMsg = mPreMsg & "<br><br>"
'mMsg = mMsg & "<tr><td colspan=2> " & mPreMsg & "  </td></tr>"
mMsg = mMsg & "<table border = 1><tr><td> CR No </td><td> " & TxtLocationID & " - " & TxtCRID & " </td></tr>"
mMsg = mMsg & "<tr><td> CR Date </td><td> " & Format(TxtCRDate.Value, "dd-mmm-yyyy") & "  </td></tr>"
mMsg = mMsg & "<tr><td> Location Name </td><td> " & TxtLocationID & "  </td></tr>"
mMsg = mMsg & "<tr><td> Instrument Type </td><td> " & CmbInstrumentType.Text & "  </td></tr>"
mMsg = mMsg & "<tr><td> Instrument No </td><td> " & TxtInstrumentNo & "  </td></tr>"
mMsg = mMsg & "<tr><td> Instrument Date </td><td> " & Format(TxtInstrumentDate, "dd-mmm-yyyy") & "  </td></tr>"
mMsg = mMsg & "<tr><td> Instrument Amount  </td><td> " & TxtInstrumentAmount & "  </td></tr>"
'mMsg = mMsg & "<tr><td> Cheque Return Charges Amount  </td><td> " & TxtChequeReturnChargesAmount & "  </td></tr>"
mMsg = mMsg & "<tr><td> Drawn On  </td><td> " & TxtDrawnOn & "  </td></tr>"
mMsg = mMsg & "<tr><td> Branch Name  </td><td> " & TxtBranchName & "  </td></tr>"
mMsg = mMsg & "<tr><td> Remarks  </td><td> " & TxtRemark & "  </td></tr>"

If LCase(mRType) <> "reject" Then
   mMsg = mMsg & "<tr><td> FAG Remarks  </td><td> " & TxtFAGRemarks & "  </td></tr>"
End If

If LCase(mRType) = "reject" Then
   mMsg = mMsg & "<tr><td> Reason for rejection  </td><td> " & TxtRejectRemark & "  </td></tr></table>"
ElseIf LCase(mRType) = "return" Then
   mMsg = mMsg & "<tr><td> Reason for return  </td><td> " & TxtReturnRemark & "  </td></tr></table>"
End If


mMsg = mMsg & "<br><br>"

mClientMsg = "Client Details" & "<br><br>"

mGridMsg = ""
For I = 1 To MSHFlexGrid2.Rows - 1
    If mGridMsg = "" Then
       mGridMsg = "<table border = 1><tr><td> " & MSHFlexGrid2.TextMatrix(I, 1) & " </td><td> " & MSHFlexGrid2.TextMatrix(I, 2) & " </td><td> " & MSHFlexGrid2.TextMatrix(I, 3) & " </td><td> " & MSHFlexGrid2.TextMatrix(I, 4) & " </td><td> " & MSHFlexGrid2.TextMatrix(I, 5) & " </td><td> " & Val(MSHFlexGrid2.TextMatrix(I, 6)) & " </td><td> " & MSHFlexGrid2.TextMatrix(I, 7) & " </td></tr>"
    Else
       mGridMsg = mGridMsg & "<tr><td> " & MSHFlexGrid2.TextMatrix(I, 1) & " </td><td> " & MSHFlexGrid2.TextMatrix(I, 2) & " </td><td> " & MSHFlexGrid2.TextMatrix(I, 3) & " </td><td> " & MSHFlexGrid2.TextMatrix(I, 4) & " </td><td> " & MSHFlexGrid2.TextMatrix(I, 5) & " </td><td> " & Val(MSHFlexGrid2.TextMatrix(I, 6)) & " </td><td> " & MSHFlexGrid2.TextMatrix(I, 7) & " </td></tr>"
    End If
Next

If mGridMsg <> "" Then
   mGridMsg = mGridMsg & "</table>"
End If
'
'mMsg = mMsg & "<tr><td> Details </td><td><table border = 1>" & mGridMsg & "</table></td></tr>"
'tmp = "Select * from Mst_StateDRCMapping Where StateID = " & mStateID & " And Flag = 'A' "

mClientMsg = mClientMsg & mGridMsg

mClientMsg = mClientMsg & "<br><br>"


mTIMsg = ""
mTIDetailMsg = ""

If MSHFlexGrid3.Rows > 3 Then

   mTIMsg = "TI Details" & "<br><br>"
   For I = 1 To MSHFlexGrid3.Rows - 1
       If mTIDetailMsg = "" Then
          mTIDetailMsg = "<table border = 1><tr><td> " & MSHFlexGrid3.TextMatrix(I, 0) & " </td><td> " & MSHFlexGrid3.TextMatrix(I, 1) & " </td><td> " & MSHFlexGrid3.TextMatrix(I, 2) & " </td><td> " & MSHFlexGrid3.TextMatrix(I, 3) & " </td><td> " & MSHFlexGrid3.TextMatrix(I, 4) & " </td><td> " & MSHFlexGrid3.TextMatrix(I, 5) & " </td><td> " & MSHFlexGrid3.TextMatrix(I, 6) & " </td><td> " & MSHFlexGrid3.TextMatrix(I, 7) & " </td><td> " & MSHFlexGrid3.TextMatrix(I, 8) & " </td><td> " & MSHFlexGrid3.TextMatrix(I, 9) & " </td><td> " & MSHFlexGrid3.TextMatrix(I, 10) & " </td></tr>"
       Else
          mTIDetailMsg = mTIDetailMsg & "<tr><td> " & MSHFlexGrid3.TextMatrix(I, 0) & " </td><td> " & MSHFlexGrid3.TextMatrix(I, 1) & " </td><td> " & MSHFlexGrid3.TextMatrix(I, 2) & " </td><td> " & MSHFlexGrid3.TextMatrix(I, 3) & " </td><td> " & MSHFlexGrid3.TextMatrix(I, 4) & " </td><td> " & MSHFlexGrid3.TextMatrix(I, 5) & " </td><td> " & MSHFlexGrid3.TextMatrix(I, 6) & " </td><td> " & MSHFlexGrid3.TextMatrix(I, 7) & " </td><td> " & MSHFlexGrid3.TextMatrix(I, 8) & " </td><td> " & MSHFlexGrid3.TextMatrix(I, 9) & " </td><td> " & MSHFlexGrid3.TextMatrix(I, 10) & " </td></tr>"
       End If
   Next
   If mTIDetailMsg <> "" Then
      mTIDetailMsg = mTIDetailMsg & "</table>"
   End If
End If

mTIMsg = mTIMsg & mTIDetailMsg

mTIMsg = mTIMsg & "<br><br>"


tmp = " Select ME.EmployeeEmailID 'Emp',ME1.EmployeeEmailID 'AM',ME2.EmployeeEmailID 'RM'"
tmp = tmp & " From Mst_Employee ME"
tmp = tmp & " Inner Join Mst_Employee ME1 On ME.DirectReportingTo = ME1.EmployeeID"
tmp = tmp & " Inner Join Mst_Employee ME2 On ME1.DirectReportingTo = ME2.EmployeeID"
tmp = tmp & " Where ME.EmployeeNo = '" & RsTxn_NewCashReceipt!CreatedBy & "' And ME.Flag <> 'N'"


Call TmpTable

TOADD = ""
TOCC = ""
If TmpRs.RecordCount > 0 Then
   TOADD = IIf(IsNull(TmpRs!EMP), "", TmpRs!EMP)
   TOCC = IIf(IsNull(TmpRs!AM), "", TmpRs!AM)
   If TOCC = "" Then
      TOCC = IIf(IsNull(TmpRs!RM), "", TmpRs!RM)
   Else
      TOCC = TOCC & "," & IIf(IsNull(TmpRs!RM), "", TmpRs!RM)
   End If
End If
'TOCC = "manish.g@ncml.com,jayaprakash.g@ncml.com,state.cmdrc@ncml.com,whr.ccg@ncml.com," 'Gen_UserEmailID '& "," & mForwardToEmailID

If TOCC = "" Then
   TOCC = "cr.return@ncml.com"
Else
   TOCC = TOCC & "," & "cr.return@ncml.com"
End If

If LCase(mRType) = "return" Then
   mMsgSubject = "Cheque Return details"
ElseIf LCase(mRType) = "reject" Then
   mMsgSubject = "Cheque Reject details"
End If

'Call SendMail(TOADD, TOCC, mMsgSubject, mMsg)

mMsg = mMsg & mClientMsg & mTIMsg

Call SendMail(TOADD, TOCC, mMsgSubject, mMsg)


End Sub
Private Function SendMail(ToADD_ As Variant, TOCC_ As Variant, MsgSubject_ As Variant, Message1_ As Variant)
Dim htmlStringhead As String
Dim htmlStringtail As String
Set oSMTP = New OSSMTP.SMTPSession

servertxt = "mail.ncml.com"

'htmlStringhead = "<html><body><font size= 3 face= Zurich BT color = Blue> " & mPreMsg & " <table border = 1>"
'htmlStringtail = "</table></Font></body></html>"
   
htmlStringhead = "<html><body><font size= 3 face= Zurich BT color = Blue> " & mPreMsg & " "
htmlStringtail = "</Font></body></html>"
   
'With oSMTP
'  'connection
'  .Server = servertxt
'  'message
'  .MailFrom = Gen_UserEmailID  '""'TOCC_ '"mayur.d@ncml.com"
'  If IsNull(Trim(ToADD_)) = True Or Trim(ToADD_) = "" Then
'     ToADD_ = "sandeep.m@ncml.com"
'  End If
'  .SendTo = ToADD_ '"parag.r@ncml.com" 'ToADD_
'  TOCC_ = TOCC_ '& "," & mGen
'  If ToADD_ <> TOCC_ Then
'     .CC = TOCC_
'  End If
'  '.BCC = "mayur.d@ncml.com"
'
''  .SendTo = "parag.r@ncml.com"
''  .CC = "prajakta.p@ncml.com"
'
'  .MessageSubject = MsgSubject_ & " (from Store-Man)" '"Confidential Mail"
'  .MessageHTML = htmlStringhead & Message1_ & htmlStringtail
'  .TimeStamp = Now()
'  If LCase(Gen_EmailConfigFlag) = "y" Then
'     .SendEmail
'  End If
'End With

Dim objSMTP
Set objSMTP = CreateObject("MailBee.SMTP")
objSMTP.LicenseKey = "MBC900-5F470FAD67-D8E202A2E22927335F24F5A667BFD196"
objSMTP.ServerName = Gen_ReadMailServerIPINI("MailIP") ' = "mail.ncml.com" '"mail.ncml.com"
objSMTP.FromAddr = Gen_UserEmailID
If IsNull(Trim(ToADD_)) = True Or Trim(ToADD_) = "" Then
   ToADD_ = "manu.ag@ncml.com"
End If
objSMTP.ToAddr = ToADD_
If ToADD_ <> TOCC_ Then
   objSMTP.CCAddr = TOCC_
End If

objSMTP.Subject = MsgSubject_ & " (from Store-Man)" 'MsgSubject_ '"Message in HTML format"
objSMTP.BodyFormat = 1
objSMTP.BodyText = htmlStringhead & Message1_ & htmlStringtail 'htmlStringhead & MsgSubject_ & " :- " & Message1_ & htmlStringtail '"<html>This word is <b>bold</b></html>"
If LCase(Gen_EmailConfigFlag) = "y" Then
   objSMTP.Send
End If


End Function



Private Sub CmdSaveReturn_Click()

'If TxtEntryNo.Text = "" Then
'   MsgBox "Blank entry No not allowed!!"
'   TxtEntryNo.SetFocus
'   Exit Sub
'End If
If TxtReturnDate.Value > Date Then
   MsgBox "Future Date not allowed!!"
   TxtReturnDate.SetFocus
   Exit Sub
End If
If TxtReturnDate.Value < CDate(TxtFAGPostingDate.Value) Then  'CDate(TxtDepositeDate.Text) Then
   MsgBox "'Return Date' should not less than 'FAG Posting Date' !!"
   TxtReturnDate.SetFocus
   Exit Sub
End If

If Format(Gen_BlockingDate, "yyyy-mm") >= Format(TxtReturnDate.Value, "yyyy-mm") Then
   MsgBox "Return Date allowed for next month of " & Format(Gen_BlockingDate, "MMM-yyyy") & "!!! "
   TxtReturnDate.SetFocus
   Exit Sub
End If

If TxtReturnRemark.Text = "" Then
   MsgBox "Blank Remark not allowed!!"
   TxtReturnRemark.SetFocus
   Exit Sub
End If

If CmbCRR.Text = "" Then
   MsgBox "Blank Cheque disonour Reason not allowed!!"
   CmbCRR.SetFocus
   Exit Sub
End If

If mGenerateAdHoc = "Y" Then  ' If ChkGenerateAdHocInvoice.Value = vbChecked And TxtAdhocAmount.Text = "" Then
   If TxtAdhocAmount.Text = "" Then
      MsgBox "Blank Cheque Dishonour Charges Amount not allowed!!"
      TxtAdhocAmount.SetFocus
      Exit Sub
   End If
   If Val(mChequeDishonourCharges) > 0 Then
      If Val(TxtAdhocAmount.Text) < Val(mChequeDishonourCharges) Then
         MsgBox "Cheque Dishonour Charges should not less than " & mChequeDishonourCharges & "!!!"
         TxtAdhocAmount.SetFocus
         Exit Sub
      End If
   End If
End If

TxtReturnRemark = Replace(TxtReturnRemark, ",", "-")
If Trim(TxtEntryNo.Text) <> "" Then
   RsTxn_NewCashReceipt!EntryNo = TxtEntryNo.Text
End If
RsTxn_NewCashReceipt!ReturnDate = TxtReturnDate.Value
RsTxn_NewCashReceipt!ReturnFlag = "Y"
RsTxn_NewCashReceipt!ReturnRemark = TxtReturnRemark.Text
RsTxn_NewCashReceipt.Update
Frame6.Visible = False
CmdReturn.Enabled = False

mRType = "Return"

Call Gen_Email

'''For Updating ChequeReturn in  Client And Depositor  Master
For I = 2 To MSHFlexGrid2.Rows - 1
    If MSHFlexGrid2.TextMatrix(I, 0) <> "" Then
        tmp = " Update Mst_Client set ChequeReturn = 'Y' Where NAVCode = '" & MSHFlexGrid2.TextMatrix(I, 10) & "' Or RNAVCode = '" & MSHFlexGrid2.TextMatrix(I, 11) & "'  "
        Call TmpTable
        
        tmp = " Update Mst_Depository set ChequeReturn = 'Y' Where NAVCode = '" & MSHFlexGrid2.TextMatrix(I, 10) & "' "
        Call TmpTable
    End If
Next
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''Insert entry in Ad-Hoc Invoice Table Txn_TempInvoice''''''''''''''''''''''''''''''
If mGenerateAdHoc = "Y" Then 'If ChkGenerateAdHocInvoice.Value = vbChecked Then
   Call SaveAdHocInvoice
End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Call SaveChequeDishonour

End Sub

Private Sub SaveChequeDishonour()
tmp = " Select * From Txn_ChequeDishonour"
Call Tmp1Table

TmpRs1.AddNew
'TmpRs1!ChequeDishonourID = GetMaxChequeDishonourID
TmpRs1!G_UID = GetGUID
TmpRs1!NewCRID = TxtCRID
TmpRs1!CRRID = mCRRID
TmpRs1!SM_Prefix = TxtPreFix
TmpRs1!CRDate = TxtCRDate.Value
TmpRs1!LocationID = mLocationID
TmpRs1!InstrumentType = CmbInstrumentType.Text
TmpRs1!DrawnOn = TxtDrawnOn
TmpRs1!BranchName = TxtBranchName
TmpRs1!InstrumentNo = TxtInstrumentNo
TmpRs1!InstrumentDate = TxtInstrumentDate.Value
TmpRs1!InstrumentAmount = Val(TxtInstrumentAmount)
TmpRs1!ReceivedFrom = TxtReceivedFrom
TmpRs1!Remark = TxtRemark
TmpRs1!ReturnRemark = TxtReturnRemark
TmpRs1!ChequeReturnChargesAmount = Val(TxtChequeReturnChargesAmount)
TmpRs1!CreatedBy = Gen_UserLoginID
TmpRs1!CreatedDate = Now
If mGenerateAdHoc = "Y" Then
    TmpRs1!GenerateAdHocInvoice = 1
    TmpRs1!PINo = mPINo
Else
    TmpRs1!GenerateAdHocInvoice = 0
End If
TmpRs1.Update
End Sub

Private Sub SaveAdHocInvoice()
Call TableTxn_TempInvoice

RsTxn_TempInvoice.AddNew
'RsTxn_TempInvoice!PINo = GetMaxPINo
'mPINo = RsTxn_TempInvoice!PINo
RsTxn_TempInvoice!G_UID = GetGUID
mFromDate = Format(TxtCRDate.Year & "-" & TxtCRDate.Month & "-" & "01", "yyyy-mm-dd")
RsTxn_TempInvoice!InvoiceDate = CDate(mFromDate)
RsTxn_TempInvoice!PIDate = Date 'Format(TxtPIDate.Value, Gen_DatFormat) ' Remaining
RsTxn_TempInvoice!InvoiceTypeID = 9 '''From Mst_IncomeExpenseType
RsTxn_TempInvoice!LocationID = mLocationID
RsTxn_TempInvoice!ExchangeTypeID = Val(MSHFlexGrid2.TextMatrix(2, 0))
RsTxn_TempInvoice!CommodityID = 1 'For Zeera's CommodityID is 1
If Val(MSHFlexGrid2.TextMatrix(2, 9)) = 1 Then
   RsTxn_TempInvoice!DPID = MSHFlexGrid2.TextMatrix(2, 8)
   RsTxn_TempInvoice!ClientID = Null
Else
   RsTxn_TempInvoice!ClientID = Val(MSHFlexGrid2.TextMatrix(2, 8))
   RsTxn_TempInvoice!DPID = Null
End If
RsTxn_TempInvoice!BatchID = 0
RsTxn_TempInvoice!RBatchID = 0
RsTxn_TempInvoice!ServiceTaxID = 1
RsTxn_TempInvoice!Amount = Format(Val(TxtAdhocAmount.Text), "#####0.00")
RsTxn_TempInvoice!Quantity = 0 'Format(Val(TxtQty.Text), "#####0.000")
RsTxn_TempInvoice!Flag = "P" ' Left(CmbStatus.Text, 1)
RsTxn_TempInvoice!Remarks = TxtRemark.Text
RsTxn_TempInvoice!CreatedBy = Gen_UserLoginID
RsTxn_TempInvoice!CreatedDate = Now

RsTxn_TempInvoice.Update

End Sub

Private Function GetMaxPINo() As Double
Dim Retval As Double
tmp = "Select max(PINo) from Txn_TempInvoice"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetMaxPINo = Retval
End Function

Private Function GetMaxChequeDishonourID() As Double
Dim Retval As Double
tmp = "Select max(ChequeDishonourID) from Txn_ChequeDishonour"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetMaxChequeDishonourID = Retval
End Function


Private Sub CmdSearch_Click()

Call Grid_Head

If SFrom.Value > STo.Value Then
   MsgBox "'From Date' must be less than 'To date'!!!"
   SFrom.SetFocus
   Exit Sub
End If

If SFrom1.Value > STo1.Value Then
   MsgBox "'Deposit From Date' must be less than 'To date'!!!"
   SFrom1.SetFocus
   Exit Sub
End If

If Val(TxtAmountFrom) > Val(TxtAmountTo) Then
   MsgBox "In Amount range 'From amount' must be less than 'To amount'!!!"
   TxtAmountFrom.SetFocus
   Exit Sub
End If

'tmp = "Select CR.CRID,LM.LocationName,CR.CRDate,CR.InstrumentType,CR.DrawnOn,CR.BranchName,CR.InstrumentNo,CR.InstrumentDate,"
'tmp = tmp & " CR.InstrumentAmount ,CR.ReceivedFrom, CR.Remark, CR.Flag, CR.HologramNo,CR.G_UID,CR.CRDepositID,TCRDD.DepositDate,TCRDD.Remark 'DepositRemarks',MFB.BankName +'~'+ MFB.AccountNo as 'Bank',CR.AdvanceFlag,CR.ReturnFlag"
'tmp = tmp & " From Txn_NewCashReceipt CR"
'tmp = tmp & " Inner Join Mst_Location LM On CR.LocationID = LM.LocationID"
''tmp = tmp & " Left Join Txn_NewCashReceiptDepositDetails TCRDD On CR.CRDepositID = TCRDD.CRDepositID "
''tmp = tmp & " Left Join Mst_FAGBank MFB On TCRDD.FAGBankID = MFB.FAGBankID "
'tmp = tmp & " Left Join Mst_FAGBank MFB On CR.FAGBankID = MFB.FAGBankID "
'
''tmp = tmp & " Where CR.CRDate between '2009-08-20 00:00:00' and '2009-08-20 00:00:00'"


tmp = "Select CR.NewCRID,LM.LocationName,CR.CRDate,CR.InstrumentType,CR.DrawnOn,CR.BranchName,CR.InstrumentNo,CR.InstrumentDate,"
tmp = tmp & " CR.InstrumentAmount ,CR.ReceivedFrom, CR.Remark, CR.Flag, CR.HologramNo,CR.G_UID,'',CR.DepositDate,CR.FAGPostingDate,CR.FAGRemarks 'DepositRemarks',MFB.BankName +'~'+ MFB.AccountNo as 'Bank',CR.AdvanceFlag,CR.ReturnFlag,CR.BatchID,CR.RejectFlag,CR.RejectDate,CR.RejectRemark,"
tmp = tmp & " MUC.EmployeeName+'~'+MUC.EmployeeNo 'CreatedBy',CR.CreatedDate,MUU.EmployeeName+'~'+MUU.EmployeeNo 'UpdatedBy',CR.UpdatedDate,MUF.EmployeeName+'~'+MUF.EmployeeNo 'FAGUpdatedBy',CR.FAGUpdatedDate,CR.RBatchID,"
tmp = tmp & " CR.FlagApproval,MUA.EmployeeName+'~'+MUA.EmployeeNo 'AMsApproveBy',CR.ApprovedDate,CR.ApprovalRemark"
tmp = tmp & " From Txn_NewCashReceipt CR"
tmp = tmp & " Inner Join Mst_Location LM On CR.LocationID = LM.LocationID"
'tmp = tmp & " Left Join Txn_NewCashReceiptDepositDetails TCRDD On CR.CRDepositID = TCRDD.CRDepositID "
'tmp = tmp & " Left Join Mst_FAGBank MFB On TCRDD.FAGBankID = MFB.FAGBankID "
tmp = tmp & " Left Join Mst_FAGBank MFB On CR.FAGBankID = MFB.FAGBankID "
tmp = tmp & " Left Join Mst_Users MUC On CR.CreatedBy = MUC.EmployeeNo"
tmp = tmp & " Left Join Mst_Users MUU On CR.UpdatedBy = MUU.EmployeeNo"
tmp = tmp & " Left Join Mst_Users MUF On CR.FAGUpdatedBy = MUF.EmployeeNo"
tmp = tmp & " Left Join Mst_Users MUA On CR.ApprovedBy = MUA.EmployeeNo"
If TxtSClientName <> "" Then
   tmp = tmp & " left join Txn_NewCashReceiptClientDetail TCRD on CR.NewCRID=TCRD.NewCRID and TCRD.SM_Prefix=CR.SM_Prefix "
   tmp = tmp & " left join Mst_Client MC on MC.ClientID=TCRD.ClientID"
   tmp = tmp & " Left join Mst_Depository DP on DP.DepositoryID=TCRD.DPID"
End If
If Gen_DBType = "MDB" Then
   tmp = tmp & " Where CR.CRDate Between #" & Format(SFrom, Gen_DatFormat) & "# And #" & Format(STo, Gen_DatFormat) & "#) "
Else
   tmp = tmp & " where CR.CRDate Between '" & Format(SFrom, Gen_DatFormat) & "' And '" & Format(STo, Gen_DatFormat) & "' "
End If

If TxtAmountFrom <> "" And TxtAmountTo <> "" Then
   tmp = tmp & " AND CR.InstrumentAmount between " & Val(TxtAmountFrom) & " And " & Val(TxtAmountTo) & " "
End If

'If CmbSCmp.Text <> "" Then
'   Wc = GenWc(Wc, mCMPName, "Mst_CMP.CMPID", "N")
'Else
'   If Gen_WcCMP <> "" Then
'      tmp = Replace(Gen_WcCMP, "CMPID", "mst_CMP.CMPID")
'      Wc = GenWc(Wc, tmp, "", "I")
'   End If
'End If


If CmbSLocation.Text <> "" Then
   tmp = tmp & " AND LM.LocationName = '" & CmbSLocation.Text & "' "
Else
   If Gen_WcLocation <> "" Then
      tmp = tmp & " AND LM." & Trim(Gen_WcLocation)
   End If
End If
If CmbSInstrumentType.Text <> "" Then
   tmp = tmp & " AND CR.InstrumentType = '" & CmbSInstrumentType.Text & "' "
End If
If TxtSCRID <> "" Then
   tmp = tmp & " And CR.NewCRID = " & Val(TxtSCRID)
End If


'If CmbSDepositBank.Text = "Without Bank Name" Then
'   tmp = tmp & " And MFB.BankName is Null"
'   If Gen_DBType = "MDB" Then
'      tmp = tmp & " And (TCRDD.DepositDate Between #" & Format(SFrom1, Gen_DatFormat) & "# And #" & Format(STo1, Gen_DatFormat) & "#)) Or TCRDD.DepositDate is Null"
'   Else
'      tmp = tmp & " And (TCRDD.DepositDate Between '" & Format(SFrom1, Gen_DatFormat) & "' And '" & Format(STo1, Gen_DatFormat) & "') Or TCRDD.DepositDate is Null "
'   End If
'ElseIf CmbSDepositBank.Text = "With Bank Name" Then
'   tmp = tmp & " And MFB.BankName is not Null"
'   If Gen_DBType = "MDB" Then
'      tmp = tmp & " And TCRDD.DepositDate Between #" & Format(SFrom1, Gen_DatFormat) & "# And #" & Format(STo1, Gen_DatFormat) & "#) "
'   Else
'      tmp = tmp & " And TCRDD.DepositDate Between '" & Format(SFrom1, Gen_DatFormat) & "' And '" & Format(STo1, Gen_DatFormat) & "' "
'   End If
'End If

'
'If CmbSDepositedBank.Text <> "" Then
'   tmp = tmp & " AND TCRDD.FAGBankID = " & mSDepoBankID & " "
'End If

If CmbSDepositBank.Text = "Without Bank Name" Then
   tmp = tmp & " And MFB.BankName is Null"
   If Gen_DBType = "MDB" Then
      tmp = tmp & " And (CR.DepositDate Between #" & Format(SFrom1, Gen_DatFormat) & "# And #" & Format(STo1, Gen_DatFormat) & "#)) Or CR.DepositDate is Null"
   Else
      tmp = tmp & " And (CR.DepositDate Between '" & Format(SFrom1, Gen_DatFormat) & "' And '" & Format(STo1, Gen_DatFormat) & "') Or CR.DepositDate is Null "
   End If
ElseIf CmbSDepositBank.Text = "With Bank Name" Then
   tmp = tmp & " And MFB.BankName is not Null"
   If Gen_DBType = "MDB" Then
      tmp = tmp & " And CR.DepositDate Between #" & Format(SFrom1, Gen_DatFormat) & "# And #" & Format(STo1, Gen_DatFormat) & "#) "
   Else
      tmp = tmp & " And CR.DepositDate Between '" & Format(SFrom1, Gen_DatFormat) & "' And '" & Format(STo1, Gen_DatFormat) & "' "
   End If
End If


If CmbSDepositedBank.Text <> "" Then
   tmp = tmp & " AND CR.FAGBankID = " & mSDepoBankID & " "
End If

If TxtSDrawnOn.Text <> "" Then
   tmp = tmp & " AND CR.DrawnOn Like '%" & TxtSDrawnOn & "%' "
End If

If TxtSInstrumentNo.Text <> "" Then
   tmp = tmp & " AND CR.InstrumentNo Like '%" & TxtSInstrumentNo & "%' "
End If

If LCase(CmbSAdvance.Text) = "yes" Then
   tmp = tmp & " AND CR.AdvanceFlag= 1"
ElseIf LCase(CmbSAdvance.Text) = "no" Then
   tmp = tmp & " AND CR.AdvanceFlag= 0"
End If

If LCase(CmbSReturn.Text) = "yes" Then
   tmp = tmp & " AND CR.ReturnFlag='Y'"
ElseIf LCase(CmbSReturn.Text) = "no" Then
   tmp = tmp & " AND CR.ReturnFlag= 'N'"
End If

If LCase(CmbSReject.Text) = "yes" Then
   tmp = tmp & " AND CR.RejectFlag='Y'"
ElseIf LCase(CmbSReject.Text) = "no" Then
   tmp = tmp & " AND CR.RejectFlag= 'N'"
End If

If ChkBatchNo.Value = 0 Then
   tmp = tmp & " AND CR.BatchID = 0 "
End If

If TxtSClientName <> "" Then
   tmp = tmp & " And (MC.ClientName Like '%" & TxtSClientName & "%' or DP.DepositoryName Like '%" & TxtSClientName & "%')"
End If

Call TmpTable

For iu = 1 To TmpRs.RecordCount
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
    For C = 0 To MSHFlexGrid1.Cols - 3
        MSHFlexGrid1.TextMatrix(iu, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
    Next
    MSHFlexGrid1.TextMatrix(iu, 2) = Format(MSHFlexGrid1.TextMatrix(iu, 2), "dd-mmm-yyyy")
    If MSHFlexGrid1.TextMatrix(iu, 7) <> "" Then
       MSHFlexGrid1.TextMatrix(iu, 7) = Format(MSHFlexGrid1.TextMatrix(iu, 7), "dd-mmm-yyyy")
    End If
    If MSHFlexGrid1.TextMatrix(iu, 15) <> "" Then
       MSHFlexGrid1.TextMatrix(iu, 15) = Format(MSHFlexGrid1.TextMatrix(iu, 15), "dd-mmm-yyyy")
    End If
    If MSHFlexGrid1.TextMatrix(iu, 16) <> "" Then
       MSHFlexGrid1.TextMatrix(iu, 16) = Format(MSHFlexGrid1.TextMatrix(iu, 16), "dd-mmm-yyyy")
    End If
    If MSHFlexGrid1.TextMatrix(iu, 26) <> "" Then
       MSHFlexGrid1.TextMatrix(iu, 26) = Format(MSHFlexGrid1.TextMatrix(iu, 26), "dd-mmm-yyyy")
    End If
    If MSHFlexGrid1.TextMatrix(iu, 28) <> "" Then
       MSHFlexGrid1.TextMatrix(iu, 28) = Format(MSHFlexGrid1.TextMatrix(iu, 28), "dd-mmm-yyyy")
    End If
    If MSHFlexGrid1.TextMatrix(iu, 30) <> "" Then
       MSHFlexGrid1.TextMatrix(iu, 30) = Format(MSHFlexGrid1.TextMatrix(iu, 30), "dd-mmm-yyyy")
    End If
    If MSHFlexGrid1.TextMatrix(iu, 19) = "1" Then
       MSHFlexGrid1.TextMatrix(iu, 19) = "Yes"
    Else
       MSHFlexGrid1.TextMatrix(iu, 19) = "No"
    End If
    If MSHFlexGrid1.TextMatrix(iu, 20) = "Y" Then
       MSHFlexGrid1.TextMatrix(iu, 20) = "Yes"
    Else
       MSHFlexGrid1.TextMatrix(iu, 20) = "No"
    End If
    If MSHFlexGrid1.TextMatrix(iu, 22) = "Y" Then
       MSHFlexGrid1.TextMatrix(iu, 22) = "Yes"
    Else
       MSHFlexGrid1.TextMatrix(iu, 22) = "No"
    End If
    If TmpRs!FlagApproval = "AMS" Then
       MSHFlexGrid1.TextMatrix(iu, 32) = IIf(IsNull(TmpRs!AMsApproveBy), "", TmpRs!AMsApproveBy)
       MSHFlexGrid1.TextMatrix(iu, 33) = IIf(IsNull(TmpRs!ApprovedDate), "", TmpRs!ApprovedDate)
       MSHFlexGrid1.TextMatrix(iu, 34) = IIf(IsNull(TmpRs!ApprovalRemark), "", TmpRs!ApprovalRemark)
    End If
    TmpRs.MoveNext
Next

End Sub

Private Sub Command1_Click()
TxtEntryNo = ""
TxtReturnDate.Value = Date
TxtReturnRemark = ""
'ChkGenerateAdHocInvoice.Value = vbUnchecked
CmbCRR.Text = ""
TxtAdhocAmount = ""
Frame6.Visible = False
End Sub


Private Sub CmdCancelReject_Click()
Frame8.Visible = False
ButtonFrm.Enabled = True
End Sub



Private Sub DeleteCmd_Click()


If RsTxn_NewCashReceipt!RejectFlag = "Y" Then
   MsgBox "Receipt is rejected. You can't delete this receipt!!!! "
   Exit Sub
End If

If RsTxn_NewCashReceipt!Flag = "P" Then
   MsgBox "Print out is generated already. You can't delete this receipt!!!! "
   Exit Sub
End If


tmp = " Select TCCD.NewCRID,TCCD.SM_PRefix "
tmp = tmp & " From Txn_CITFCRDetail TCCD "
tmp = tmp & " Inner Join Txn_CITFCR TCC On TCCD.CITFCRID = TCC.CITFCRID "
tmp = tmp & " Where TCC.Flag = 'N' And TCCD.NewCRID = " & TxtCRID & " And TCCD.SM_Prefix = '" & TxtPreFix & "' "

Call Tmp4Table

If TmpRs4.RecordCount > 0 Then
   MsgBox "Cash Receipt used in CITF CR Mapping module. Can't Delete!!!"
   Exit Sub
End If

Dim ans As Variant
On Error GoTo msgError

ans = MsgBox("Do you really want to DELETE this Record??", vbYesNo)
If ans = vbYes Then
   Call DeleteClientDetail
   Call DeleteTIDetail
   
   RsTxn_NewCashReceipt.Delete
   RsTxn_NewCashReceipt.Update
   RsTxn_NewCashReceipt.Requery
   If RsTxn_NewCashReceipt.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsTxn_NewCashReceipt.MoveNext
   If RsTxn_NewCashReceipt.EOF() Then
      RsTxn_NewCashReceipt.MoveLast
   End If
   Call Indata
End If
Exit Sub
msgError:
MsgBox "Child record Present "
RsTxn_NewCashReceipt.CancelUpdate
End Sub

Private Sub EditCmd_Click()

If RsTxn_NewCashReceipt!RejectFlag = "Y" Then
   MsgBox "Receipt is rejected. You can't edit this receipt!!!! "
   Exit Sub
End If

If (LCase(Gen_UserRole) = "fag" Or LCase(Gen_UserRole) = "power" Or LCase(Gen_UserRole) = "fagadmin") Then
'   If LCase(mFlagA) <> "ams" And mFlagA <> "sys" Then
'      MsgBox "Still AMs not approve this Cash Receipt"
'      Exit Sub
'   End If
   If RsTxn_NewCashReceipt!BatchID <> 0 Then
      MsgBox "Batch is generated already. You can't edit this receipt!!!! "
      Exit Sub
   End If
   Frame1.Enabled = True
Else
   If IsNull(RsTxn_NewCashReceipt!FAGPostingDate) = False Then
      MsgBox "Already Posted by FAG. You can't edit this receipt!!!! "
      Exit Sub
   End If
   If RsTxn_NewCashReceipt!Flag = "P" Then
      MsgBox "Print out is generated already. You can't edit this receipt!!!! "
      Exit Sub
   End If
   Frame1.Enabled = False
End If

Edit_Flg = "E"

CmbLocationID.Text = TxtLocationID

If (LCase(Gen_UserRole) = "fag" Or LCase(Gen_UserRole) = "power" Or LCase(Gen_UserRole) = "fagadmin") Then
   CmbBankName.Visible = True
   TxtBankName.Visible = False
   CmbBankName.Text = TxtBankName
End If

Call AppendTIDetails

Frame5.Enabled = True
Frame7.Enabled = True
Frame2.Enabled = True

CmdPrint.Enabled = False
CmdReject.Enabled = False
CmdPrintPreview.Enabled = False
CmdApproveAM.Enabled = False

TxtInstrumentAmount.Locked = False
TxtChequeReturnChargesAmount.Locked = False
For I = 2 To MSHFlexGrid3.Rows - 1
    If MSHFlexGrid3.TextMatrix(I, 18) <> "" Then
       TxtInstrumentAmount.Locked = True
       TxtChequeReturnChargesAmount.Locked = True
    End If
Next

Call GButtonLock(Me, False)

End Sub



Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Call GButtonLockAD(Me)
   Frame1.Enabled = False
   SSTab1.Enabled = False
   CmdPrint.Enabled = False
   CmdReject.Enabled = False
   CmdPrintPreview.Enabled = False
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If
Frame1.Enabled = True
SSTab1.Enabled = True
Call Grid_Head1
Call Grid_Head2
CmdPrint.Enabled = True
CmdPrintPreview.Enabled = True
wc = ""
If Gen_WcLocation <> "" Then
   wc = "Where " & Gen_WcLocation
End If

Call Grid_Head
SFrom.Value = Date - 730
STo.Value = Date
SFrom1.Value = Date - 730
STo1.Value = Date
CmbSLocation.Text = ""
CmbSInstrumentType.Text = ""
TxtAmountFrom = ""
TxtAmountTo = ""
TxtSDrawnOn = ""
CmbSAdvance.Text = ""
CmbSReturn.Text = ""
TxtSCRID = ""
TxtSInstrumentNo = ""
TxtSClientName = ""

Call TableTxn_NewCashReceipt(" " & wc & "  Order by SM_Prefix,NewCRID")
If RsTxn_NewCashReceipt.RecordCount = 0 Then
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
Private Sub CmbLocationID_GotFocus()
tmp = CmbLocationID.Text

If Gen_WcLocation = "" Then
   Call TableMst_Location("")
Else
   Call TableMst_Location("where " & Gen_WcLocation)
End If

CmbLocationID.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No record found !!!! "
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

If tmp <> CmbLocationID.Text Then
    Call Grid_Head1
    Call Grid_Head2
    Call ClearClientTabControl
    Call ClearTITabControl
End If

mLocationID = RsMst_Location!LocationID
TxtPreFix = GetPreFixByLocation(RsMst_Location!LocationID, "P")
mStateID = RsMst_Location!StateID

If ChkInstrumentNo = False Then
   If LCase(CmbInstrumentType.Text) = "cheque" Then
      MsgBox "Cheque No already entered for Cash Receipt - " & TmpRs1!NewCRID & "-" & GetLocationName(TmpRs1!LocationID) & " "
   ElseIf LCase(CmbInstrumentType.Text) = "dd" Then
      MsgBox "DD No already entered for Cash Receipt - " & TmpRs1!NewCRID & "-" & GetLocationName(TmpRs1!LocationID) & " "
   ElseIf LCase(CmbInstrumentType.Text) = "rtgs" Then
      MsgBox "RTGS No already entered for Cash Receipt - " & TmpRs1!NewCRID & "-" & GetLocationName(TmpRs1!LocationID) & " "
   End If
   CmbLocationID.SetFocus
   Exit Sub
End If

End Sub

Private Function CheckReturnClient() As Boolean
Dim Retval As Boolean
Retval = False

For I = 2 To MSHFlexGrid2.Rows - 1
    If MSHFlexGrid2.TextMatrix(I, 0) <> "" Then
        If MSHFlexGrid2.TextMatrix(I, 13) = "Y" Then
            Retval = True
            Exit For
        End If
    End If
Next
CheckReturnClient = Retval
End Function

Private Sub CmbInstrumentType_LostFocus()
If CmbInstrumentType.Text = "" Then Exit Sub
CmbInstrumentType.Text = Trim(CmbInstrumentType.Text)

If LCase(CmbInstrumentType.Text) = "cash" Then
   TxtDrawnOn.Enabled = False
   TxtBranchName.Enabled = False
   TxtInstrumentNo.Enabled = False
   TxtInstrumentDate.Enabled = False
   TxtDrawnOn = ""
   TxtBranchName = ""
   TxtInstrumentNo = ""
   Label5.Caption = "Instrument No"
   Label7.Caption = "Instrument Date"
'   TxtInstrumentDate = ""
   Exit Sub
Else
   TxtDrawnOn.Enabled = True
   TxtBranchName.Enabled = True
   TxtInstrumentNo.Enabled = True
   TxtInstrumentDate.Enabled = True

End If
If LCase(CmbInstrumentType.Text) = "cheque" Then
   Label5.Caption = "Cheque No"
   Label7.Caption = "Cheque Date"
   If CheckReturnClient = True Then
       'MsgBox "You can not select Instrument Type Cheque!!!"
       MsgBox "This Client cheque was dishonoured in past, hence we do not accept cheque from this client any more!!!"
       CmbInstrumentType.SetFocus
       Exit Sub
   End If
ElseIf LCase(CmbInstrumentType.Text) = "dd" Then
   Label5.Caption = "DD No"
   Label7.Caption = "DD Date"
   
ElseIf LCase(CmbInstrumentType.Text) = "rtgs" Then
   Label5.Caption = "RTGS No"
   Label7.Caption = "RTGS Date"
   'If CheckReturnClient = True Then
   '   MsgBox "You can not select Instrument Type RTGS!!!"
   '   CmbInstrumentType.SetFocus
   '   Exit Sub
   'End If
Else
   Label5.Caption = "Instrument No"
   Label7.Caption = "Instrument Date"
End If

If LCase(CmbInstrumentType.Text) = "cheque" Or LCase(CmbInstrumentType.Text) = "dd" Or LCase(CmbInstrumentType.Text) = "rtgs" Then
   TxtDrawnOn.Enabled = True
   TxtBranchName.Enabled = True
   TxtInstrumentNo.Enabled = True
   TxtInstrumentDate.Enabled = True
   Exit Sub
End If

MsgBox "Invalid selection !!!!"
CmbInstrumentType.SetFocus

End Sub


Private Sub CmbExchangeTypeID_GotFocus()
tmp = CmbExchangeTypeID.Text
Call TableMst_ExchangeType
CmbExchangeTypeID.Clear
If RsMst_ExchangeType.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_ExchangeType.RecordCount
    CmbExchangeTypeID.AddItem RsMst_ExchangeType!ExchangeType
    RsMst_ExchangeType.MoveNext
Next
CmbExchangeTypeID.Text = tmp
End Sub
'----
Private Sub CmbExchangeTypeID_LostFocus()
If CmbExchangeTypeID.Text = "" Then
   mExchangeTypeID = 0
   CmbCNonNcdexClient.Clear
   CmbClientDP.Clear
   Exit Sub
End If
RsMst_ExchangeType.MoveFirst
RsMst_ExchangeType.Find "ExchangeType = '" & CmbExchangeTypeID.Text & "'"
If RsMst_ExchangeType.EOF Then
   MsgBox "Invalid selection "
   CmbExchangeTypeID.SetFocus
   Exit Sub
End If

If mExchangeTypeID <> RsMst_ExchangeType!ExchangeTypeID Then
   CmbCNonNcdexClient.Clear
   CmbClientDP.Clear
   CmbDPName.Clear
End If
mExchangeTypeID = RsMst_ExchangeType!ExchangeTypeID

If mExchangeTypeID = 1 Then
   CmbCNonNcdexClient.Width = 9400
   CmbCNonNcdexClient.Top = TxtClientID.Top
   CmbCNonNcdexClient.Left = TxtClientID.Left
   CmbCNonNcdexClient.Visible = True
   Label25.Visible = False
   Label26.Visible = False
   Label28.Visible = False
   Check3.Value = 0
   Check3.Enabled = False
Else
   Label25.Visible = True
   Label26.Visible = True
   Label28.Visible = True
   CmbCNonNcdexClient.Visible = False
   Check3.Enabled = True
   Call Check3_Click
End If

End Sub

Private Sub Form_Resize()

If Me.WindowState = vbMaximized Then
   SetScrollBars
End If

End Sub



Private Sub MSHFlexGrid1_Click()
Dim I As Variant
I = MSHFlexGrid1.RowSel
If I <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(I, 0) = "" Then Exit Sub
RsTxn_NewCashReceipt.MoveFirst
'RsTxn_NewCashReceipt.Find "CRID = " & Val(MSHFlexGrid1.TextMatrix(I, 0))
RsTxn_NewCashReceipt.Find "G_UID= '" & MSHFlexGrid1.TextMatrix(I, 13) & "'"

If Not RsTxn_NewCashReceipt.EOF Then
   Call Indata
   Frame0.Visible = False
End If

End Sub

Private Sub MSHFlexGrid2_Click()
If SaveCmd.Enabled = False Then Exit Sub
r = MSHFlexGrid2.RowSel

If r < 2 Or r = MSHFlexGrid2.Rows - 1 Then
   Call ClearClientTabControl
   Exit Sub
End If
If MSHFlexGrid2.TextMatrix(r, 0) = "" Then Exit Sub
CmdCAddToList.Caption = "Update List"

mExchangeTypeID = MSHFlexGrid2.TextMatrix(r, 0)
'CmbNAVCode.Clear
If mExchangeTypeID = 1 Then
   CmbCNonNcdexClient.Width = 9400
   CmbCNonNcdexClient.Top = TxtClientID.Top
   CmbCNonNcdexClient.Left = TxtClientID.Left
   CmbCNonNcdexClient.Visible = True
   CmbCNonNcdexClient.Text = MSHFlexGrid2.TextMatrix(r, 3)
   Label25.Visible = False
   Label26.Visible = False
   Label28.Visible = False
Else
   Label25.Visible = True
   Label26.Visible = True
   Label28.Visible = True
   CmbCNonNcdexClient.Visible = False
End If
Check3.Value = Val(MSHFlexGrid2.TextMatrix(r, 9))
CmbExchangeTypeID.Text = MSHFlexGrid2.TextMatrix(r, 1)
TxtClientID = MSHFlexGrid2.TextMatrix(r, 2)
TxtCClientName.Text = MSHFlexGrid2.TextMatrix(r, 3)
TxtCDPID.Text = MSHFlexGrid2.TextMatrix(r, 4)
TxtCDPName.Text = MSHFlexGrid2.TextMatrix(r, 5)
TxtCAmt.Text = MSHFlexGrid2.TextMatrix(r, 6)
TxtCRemark = MSHFlexGrid2.TextMatrix(r, 7)

If Check3.Value = 0 Then
   mClientIDRow = MSHFlexGrid2.TextMatrix(r, 8)
   mClientNAVCode = MSHFlexGrid2.TextMatrix(r, 10)
   mClientRNAVCode = MSHFlexGrid2.TextMatrix(r, 11)
'   CmbNAVCode.AddItem mClientNAVCode
'   CmbNAVCode.AddItem mClientRNAVCode
Else
   mDPID = MSHFlexGrid2.TextMatrix(r, 8)
   mDPNAVCode = MSHFlexGrid2.TextMatrix(r, 10)
'   CmbNAVCode.AddItem mDPNAVCode
End If

CmbNAVCode.Text = MSHFlexGrid2.TextMatrix(r, 12)
mChequeReturn = MSHFlexGrid2.TextMatrix(r, 13)

If Check3.Value = 1 Then
   If CmbDPName.Visible = True Then
      CmbDPName.Text = MSHFlexGrid2.TextMatrix(r, 5) & "~" & MSHFlexGrid2.TextMatrix(r, 4)
   End If
End If
Frame5.Enabled = False
'CmbExchangeTypeID.SetFocus
End Sub

Private Sub MSHFlexGrid3_Click()
If SaveCmd.Enabled = False Then Exit Sub
r = MSHFlexGrid3.RowSel
If MSHFlexGrid3.TextMatrix(r, 0) = "" Then Exit Sub
If Val(MSHFlexGrid3.TextMatrix(r, 18)) <> 0 Then
   MsgBox "This invoice is allocated through Advance Allocation Screen . Can't Edit!!!"
   Exit Sub
End If
If r <= 1 Then Exit Sub

TxtTINo.Text = MSHFlexGrid3.TextMatrix(r, 0)
TxtTIDate.Text = MSHFlexGrid3.TextMatrix(r, 1)
TxtBalAmount = MSHFlexGrid3.TextMatrix(r, 6)
TxtRecAmt = MSHFlexGrid3.TextMatrix(r, 7)
TxtTDSAmt = MSHFlexGrid3.TextMatrix(r, 8)
TxtOtherDed = MSHFlexGrid3.TextMatrix(r, 9)
TxtBalRec = MSHFlexGrid3.TextMatrix(r, 8)
TxtTIClient = MSHFlexGrid3.TextMatrix(r, 13) & "~" & MSHFlexGrid3.TextMatrix(r, 14)
TxtTILocation = MSHFlexGrid3.TextMatrix(r, 19)
TxtRecAmt.SetFocus
End Sub

Private Sub SaveCmd_Click()
On Error GoTo MsgErr

mErrMessage = "Save"

If Edit_Flg = "A" Then

'   If TxtCRDate.Value < Date - 30 Then
'      MsgBox "Only one month Back Date allowed !!! "
'      TxtCRDate.SetFocus
'      Exit Sub
'   End If

   If Month(TxtCRDate) < 6 And Year(TxtCRDate) <= 2011 Then
      MsgBox " Cash Receipt earlier to June 2011 is not allowed  !!!"
      TxtCRDate.SetFocus
      Exit Sub
   End If
    
End If

If CmbLocationID.Text = "" Then
   MsgBox "Blank Location Name not allowed !!! "
   CmbLocationID.SetFocus
   Exit Sub
End If

If TxtReceivedFrom.Text = "" Then
   MsgBox "Blank Received From not allowed !!! "
   TxtReceivedFrom.SetFocus
   Exit Sub
End If

If CmbInstrumentType.Text = "" Then
   MsgBox "Blank Instrument Type not allowed !!! "
   CmbInstrumentType.SetFocus
   Exit Sub
End If

If Val(TxtInstrumentAmount) <= 0 Then
   MsgBox "Blank Instrument Amount Not Allowed !!! "
   TxtInstrumentAmount.SetFocus
   Exit Sub
End If

'If Val(TxtChequeReturnChargesAmount) < 0 Then
'   MsgBox "Blank Cheque Return Charges Amount Not Allowed !!! "
'   TxtChequeReturnChargesAmount.SetFocus
'   Exit Sub
'End If
'
'If Val(TxtChequeReturnChargesAmount) > Val(TxtInstrumentAmount) Then
'   MsgBox "Cheque Return Charges Amount should not greater than Instrument Amount!!! "
'   TxtChequeReturnChargesAmount.SetFocus
'   Exit Sub
'End If

If LCase(CmbInstrumentType.Text) <> "cash" Then
   If TxtDrawnOn.Text = "" Then
      MsgBox "Blank 'Drawn On' not allowed !!! "
      TxtDrawnOn.SetFocus
      Exit Sub
   End If
   If TxtBranchName.Text = "" Then
      MsgBox "Blank Branch Name not allowed !!! "
      TxtBranchName.SetFocus
      Exit Sub
   End If
   If TxtInstrumentNo.Text = "" Then
      MsgBox "Blank Instrument No not allowed !!! "
      TxtInstrumentNo.SetFocus
      Exit Sub
   End If
End If

If TxtRemark.Text = "" Then
   MsgBox "Blank Remark not allowed !!! "
   TxtRemark.SetFocus
   Exit Sub
End If

If TxtDepDate.Visible = True Then
    If TxtCRDate.Value > TxtDepDate.Value Then
       MsgBox "Cash Receipt Date should not greater than Deposit Date!!!"
       TxtDepDate.SetFocus
       Exit Sub
    End If

    If Format(Gen_BlockingDate, "yyyy-mm") >= Format(TxtDepDate.Value, "yyyy-mm") Then
       MsgBox "Deposite Date allowed for next month of " & Format(Gen_BlockingDate, "MMM-yyyy") & "!!!"
       TxtDepDate.SetFocus
       Exit Sub
    End If

End If

If TxtFAGPostingDate.Visible = True Then
   If TxtDepDate.Value > TxtFAGPostingDate.Value Then
      MsgBox "FAG Posting date must be greater than Deposit Date!!!"
      TxtFAGPostingDate.SetFocus
      Exit Sub
   End If

   If Format(Gen_BlockingDate, "yyyy-mm") >= Format(TxtFAGPostingDate.Value, "yyyy-mm") Then
      MsgBox "FAG Posting Date allowed for next month of " & Format(Gen_BlockingDate, "MMM-yyyy") & "!!!"
      TxtFAGPostingDate.SetFocus
      Exit Sub
   End If

End If

If MSHFlexGrid2.Rows < 4 Then
   MsgBox "Blank Client Details not allowed !!! "
   SSTab1.Tab = 0
   CmbExchangeTypeID.SetFocus
   Exit Sub
End If

''If ChkAdv.Value = 1 Then
''   If Val(TxtInstrumentAmount) <> Val(MSHFlexGrid2.TextMatrix(0, 6)) Then
''      MsgBox "Total Client wise amount must be same as Instrument Amount!"
''      'TxtInstrumentAmount.SetFocus
''      Exit Sub
''   End If
''Else
''   If MSHFlexGrid3.Rows < 4 Then
''      MsgBox "Blank TI Details not allowed !!! "
''      SSTab1.Tab = 0
''      Exit Sub
''   End If
''   If Val(TxtInstrumentAmount) <> Val(MSHFlexGrid3.TextMatrix(0, 7)) Then
''      MsgBox "Total Received amount must be same as Instrument Amount!"
''      'TxtInstrumentAmount.SetFocus
''      Exit Sub
''   End If
''
''End If

If MSHFlexGrid3.Rows > 3 Then
   If Val(MSHFlexGrid3.TextMatrix(0, 7)) = 0 And Val(MSHFlexGrid3.TextMatrix(0, 8)) = 0 Then
      Dim ans As Variant
      ans = MsgBox("You have not Updated Tax Invoice detail against Cash Receipt. Do you want to Continue [Yes /No]??", vbYesNo)
      If ans = vbNo Then
         SSTab1.Tab = 1
         Exit Sub
      End If
   End If
End If

If TxtFAGPostingDate.Visible = True Then
   If LCase(mFlagA) <> "sys" And LCase(mFlagA) <> "ams" Then
      If (LCase(Gen_UserRole) = "fag" Or LCase(Gen_UserRole) = "power" Or LCase(Gen_UserRole) = "fagadmin") Then
         If CheckMultiClient = True Then
            If MsgBox("These cash receipt contains multiple clients and not approved by concerned AMs. Do you want to Continue [Yes /No]??", vbYesNo) = vbNo Then
               Exit Sub
            End If
         End If
      End If
   End If
End If

mTotrec = 0

For I = 2 To MSHFlexGrid3.Rows - 1
    If MSHFlexGrid3.TextMatrix(I, 18) = "" Then
       mTotrec = Val(mTotrec) + Val(MSHFlexGrid3.TextMatrix(I, 7))
       mTotrec = Format(mTotrec, "#######0.00")
    End If
Next

'If Val(TxtInstrumentAmount) - Val(TxtChequeReturnChargesAmount) <> Val(MSHFlexGrid2.TextMatrix(0, 6)) + Val(mTotrec) Then 'If Val(TxtInstrumentAmount) <> Val(MSHFlexGrid2.TextMatrix(0, 6)) + Val(MSHFlexGrid3.TextMatrix(0, 7)) Then
'   MsgBox "'Total Excess Received amount' in Client Details tab + 'Total Received Amount against this CR' in TI Details tab must be equal to Instrument Amount - Cheque Return Charges Amount!!!"
   'TxtInstrumentAmount.SetFocus
'   Exit Sub
'End If

'mInsAMt = Val(TxtInstrumentAmount)
'mAdvAmt = Format(Val(MSHFlexGrid2.TextMatrix(0, 6)), "#######0.00")

If Format(Val(TxtInstrumentAmount), "#######0.00") <> Format(Val(MSHFlexGrid2.TextMatrix(0, 6)) + Val(mTotrec), "#######0.00") Then 'If Val(TxtInstrumentAmount) <> Val(MSHFlexGrid2.TextMatrix(0, 6)) + Val(MSHFlexGrid3.TextMatrix(0, 7)) Then
'If Val(mInsAMt) <> Val(mAdvAmt) + Val(mTotrec) Then 'If Val(TxtInstrumentAmount) <> Val(MSHFlexGrid2.TextMatrix(0, 6)) + Val(MSHFlexGrid3.TextMatrix(0, 7)) Then
   MsgBox "'Total Excess Received amount' in Client Details tab + 'Total Received Amount against this CR' in TI Details tab must be equal to Instrument Amount!!!"
   Exit Sub
End If

'''Comment When Added Cheque Return Charges Amount ----------------------------------

'If Val(TxtInstrumentAmount) <> Val(MSHFlexGrid2.TextMatrix(0, 6)) + Val(mTotrec) Then 'If Val(TxtInstrumentAmount) <> Val(MSHFlexGrid2.TextMatrix(0, 6)) + Val(MSHFlexGrid3.TextMatrix(0, 7)) Then
'   MsgBox "'Total Excess Received amount' in Client Details tab + 'Total Received Amount against this CR' in TI Details tab must be equal to Instrument Amount!!!"
'   'TxtInstrumentAmount.SetFocus
'   Exit Sub
'End If

'''Comment When Added Cheque Return Charges Amount ----------------------------------


If Edit_Flg = "A" Then
'   If (LCase(Gen_UserRole) = "fag" Or LCase(Gen_UserRole) = "power" Or LCase(Gen_UserRole) = "fagadmin") Then
'      If CheckMultiClient = True Then
'         If MsgBox("Are you sure to save this cash recepit with approval of AMs?", vbYesNo) = vbNo Then
'            Exit Sub
'         End If
'      End If
'   End If
   RsTxn_NewCashReceipt.AddNew
   RsTxn_NewCashReceipt!CRID = GetCRID
   RsTxn_NewCashReceipt!G_UID = GetGUID
   'TxtCRID.Text = RsTxn_NewCashReceipt!CRID
   TxtG_UID = RsTxn_NewCashReceipt!G_UID
Else
   RsTxn_NewCashReceipt.MoveFirst
   RsTxn_NewCashReceipt.Find "G_UID= '" & TxtG_UID & "'"
End If

'RsTxn_NewCashReceipt!CRID = TxtCRID
RsTxn_NewCashReceipt!CRDate = TxtCRDate.Value
RsTxn_NewCashReceipt!InstrumentType = CmbInstrumentType.Text
RsTxn_NewCashReceipt!DrawnOn = TxtDrawnOn
RsTxn_NewCashReceipt!BranchName = TxtBranchName
RsTxn_NewCashReceipt!InstrumentNo = TxtInstrumentNo
RsTxn_NewCashReceipt!InstrumentDate = TxtInstrumentDate.Value
RsTxn_NewCashReceipt!InstrumentAmount = Val(TxtInstrumentAmount)
RsTxn_NewCashReceipt!ChequeReturnChargesAmount = Val(TxtChequeReturnChargesAmount)
RsTxn_NewCashReceipt!Remark = TxtRemark
If Val(MSHFlexGrid2.TextMatrix(0, 6)) > 0 Then
   RsTxn_NewCashReceipt!AdvanceFlag = 1
Else
   RsTxn_NewCashReceipt!AdvanceFlag = 0
End If
RsTxn_NewCashReceipt!LocationID = mLocationID
RsTxn_NewCashReceipt!SM_Prefix = TxtPreFix
RsTxn_NewCashReceipt!ReceivedFrom = TxtReceivedFrom
RsTxn_NewCashReceipt!BatchID = 0
RsTxn_NewCashReceipt!RBatchID = 0
RsTxn_NewCashReceipt!ReturnFlag = "N"
RsTxn_NewCashReceipt!RejectFlag = "N"

If CheckMultiClient = False Then
   RsTxn_NewCashReceipt!FlagApproval = "SYS"
   RsTxn_NewCashReceipt!ApproveDate = Date
   RsTxn_NewCashReceipt!ApprovalRemark = TxtApproveRemarks.Text
   RsTxn_NewCashReceipt!ApproveBy = Gen_UserLoginID
End If

If LCase(Gen_UserRole) = "fag" Or LCase(Gen_UserRole) = "power" Or LCase(Gen_UserRole) = "fagadmin" Then
   If CmbBankName.Text <> "" Then
      If TxtFAGPostingDate.Visible = True Then
         RsTxn_NewCashReceipt!FAGPostingDate = TxtFAGPostingDate
      End If
      RsTxn_NewCashReceipt!FAGBankID = mFAGBankID 'CmbBankName.Visible = True
      RsTxn_NewCashReceipt!FAGRemarks = TxtFAGRemarks
      RsTxn_NewCashReceipt!DepositDate = TxtDepDate
      If mPrintFlag = "" Then
         RsTxn_NewCashReceipt!Flag = "P"
         RsTxn_NewCashReceipt!HologramNo = 0
      End If
      RsTxn_NewCashReceipt!FAGUpdatedBy = Gen_UserLoginID
      RsTxn_NewCashReceipt!FAGUpdatedDate = Now
   End If
Else
   RsTxn_NewCashReceipt!FAGPostingDate = Null
   RsTxn_NewCashReceipt!FAGBankID = Null 'CmbBankName.Visible = True
   RsTxn_NewCashReceipt!FAGRemarks = Null
End If

If IsNull(RsTxn_NewCashReceipt!CreatedBy) = True Or RsTxn_NewCashReceipt!CreatedBy = "" Then
   RsTxn_NewCashReceipt!CreatedBy = Gen_UserLoginID
   RsTxn_NewCashReceipt!CreatedDate = Now
Else
   If RsTxn_NewCashReceipt!FAGPostingDate = Null Then
      RsTxn_NewCashReceipt!UpdatedBy = Gen_UserLoginID
      RsTxn_NewCashReceipt!UpdatedDate = Now
   End If
End If

RsTxn_NewCashReceipt.Update

If Edit_Flg = "A" Then
   tmp = "Select * from NewCashReceipt Where G_UID= '" & TxtG_UID & "'"
   Call Tmp1Table
   If TmpRs1.RecordCount > 0 Then
      TxtCRID.Text = TmpRs1!NewCRID
   End If
End If


Call SaveClientDetail
Call SaveTIDetail

RsTxn_NewCashReceipt.Requery
RsTxn_NewCashReceipt.MoveFirst
RsTxn_NewCashReceipt.Find "G_UID= '" & TxtG_UID & "'"
If RsTxn_NewCashReceipt.EOF Then
   RsTxn_NewCashReceipt.MoveLast
End If

Call Indata

If IsNull(RsTxn_NewCashReceipt!DepositDate) And mFlagA = "" Then
   Call Gen_Mail
End If
Exit Sub

MsgErr:
MsgBox Err.Description
Call Gen_EmailForError("parag.r@ncml.com", "", Err.Description, mErrMessage)
Unload Me
End Sub

Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   If RsTxn_NewCashReceipt.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   RsTxn_NewCashReceipt.MoveFirst
   RsTxn_NewCashReceipt.Find "G_UID= '" & TxtG_UID & "'"
   If RsTxn_NewCashReceipt.EOF Then
      RsTxn_NewCashReceipt.MoveLast
   End If
   Call Indata
   Exit Sub
End If
Frame0.Height = Frame4.Height - 200
Frame0.Left = Frame1.Left - 100
Frame0.Top = Frame1.Top - 100
Frame0.Width = Frame4.Width - 150
Frame0.Visible = True
MSHFlexGrid1.Width = Frame0.Width - 400
MSHFlexGrid1.Height = Frame0.Height - 2400
RsTxn_NewCashReceipt.Filter = ""
RsTxn_NewCashReceipt.MoveFirst

''===========comment on 7 july 2012 for not clearing search grid=============
'MSHFlexGrid1.Clear
'MSHFlexGrid1.Cols = 11
'Call Grid_Head
'SFrom.Value = Date - 730
'STo.Value = Date
'SFrom1.Value = Date - 730
'STo1.Value = Date
'CmbSLocation.Text = ""
'CmbSInstrumentType.Text = ""
'TxtAmountFrom = ""
'TxtAmountTo = ""
'TxtSDrawnOn = ""
'CmbSAdvance.Text = ""
'CmbSReturn.Text = ""
'TxtSCRID = ""
'TxtSInstrumentNo = ""
'TxtSClientName = ""
''===========comment on 7 july 2012 for not clearing search grid=============

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
Frame6.Visible = False
Frame8.Visible = False
ButtonFrm.Enabled = True
End Sub

Private Sub TxtAdhocAmount_LostFocus()
If TxtAdhocAmount.Text = "" Then Exit Sub
'If ChkGenerateAdHocInvoice.Value = vbUnchecked Then
'   TxtAdhocAmount.Text = ""
'   Exit Sub
'End If
If IsNumeric(TxtAdhocAmount.Text) = False Or Val(TxtAdhocAmount.Text) < 0 Then
   MsgBox "Invalid Numeric Format !!!"
   TxtAdhocAmount.SetFocus
   Exit Sub
End If
If Val(TxtAdhocAmount) > Val(TxtInstrumentAmount) Then
   MsgBox "Ad Hoc Amount should not greater than Instrument Amount!!!"
   TxtAdhocAmount.SetFocus
   Exit Sub
End If

TxtAdhocAmount = Round(Val(TxtAdhocAmount), 2)
End Sub

Private Sub TxtCAmt_LostFocus()

If TxtCAmt.Text = "" Then Exit Sub
If IsNumeric(TxtCAmt.Text) = False Or Val(TxtCAmt.Text) < 0 Then
   MsgBox "Invalid Numeric Format !!!"
   TxtCAmt.SetFocus
   Exit Sub
End If
TxtCAmt = Round(Val(TxtCAmt), 2)
TxtCAmt = Format(TxtCAmt, "#######0.00")
End Sub

Private Sub TxtChequeReturnChargesAmount_LostFocus()
If TxtChequeReturnChargesAmount.Text = "" Then Exit Sub
If IsNumeric(TxtChequeReturnChargesAmount.Text) = False Then 'Or Val(TxtChequeReturnChargesAmount.Text) < 0 Then
   MsgBox "Invalid Numeric Format !!!"
   TxtInstrTxtChequeReturnChargesAmountumentAmount.SetFocus
   Exit Sub
End If
TxtChequeReturnChargesAmount = Round(Val(TxtChequeReturnChargesAmount), 2)

End Sub

Private Sub TxtClientID_LostFocus()
CmbClientDP.Visible = False
If TxtClientID = "" Then Exit Sub
If CmbExchangeTypeID.Text = "" Then Exit Sub

Dim ans As Variant
ans = GetClientNameDtl(TxtClientID, mExchangeTypeID)

If ans(0, 0) = False Then
   MsgBox "Not Found !!!"
   TxtClientID.SetFocus
   Exit Sub
End If
If ans(3, 0) > 1 Then
   CmbClientDP.Visible = True
   CmbClientDP.SetFocus
End If
CmbClientDP.Text = ""
CmbClientDP.Clear
tmp = " SELECT Mst_Depository.DepositoryID, Mst_Depository.DepositoryName, Mst_Client.ClientID, " & _
      " Mst_Client.ClientName,Mst_Client.ClientID ,Mst_Client.NAVCode,Mst_Client.RNAVCode, Mst_Client.ChequeReturn " & _
      " FROM Mst_Depository INNER JOIN Mst_Client ON Mst_Depository.DepositoryID = Mst_Client.DepositoryID " & _
      " WHERE Mst_Depository.ExchangeTypeID = " & mExchangeTypeID & _
      "  And  (((Mst_Client.ExchangeTypeID)=" & mExchangeTypeID & ") AND ((Mst_Client.ClientID)='" & TxtClientID & "'));"

Call Tmp1Table

For op = 1 To TmpRs1.RecordCount
    If CmbClientDP.Text = "" Then
       CmbClientDP.Text = TmpRs1!DepositoryID & " ~ " & TmpRs1!DepositoryName
    End If
    CmbClientDP.AddItem TmpRs1!DepositoryID & " ~ " & TmpRs1!DepositoryName
    TmpRs1.MoveNext
Next
If ans(3, 0) <= 1 Then
   CmbClientDP_LostFocus
End If

End Sub

Sub ClearClientTabControl()
TxtClientID = ""
TxtCClientName = ""
TxtCDPID = ""
TxtCDPName = ""
CmbClientDP.Text = ""
CmbClientDP.Visible = False
TxtCRemark.Text = ""
TxtCAmt = ""
CmbNAVCode.Clear
CmbNAVCode.Text = ""
CmbExchangeTypeID.Text = ""
mExchangeTypeID = 0
CmbCNonNcdexClient.Text = ""
CmdCAddToList.Caption = "Add to List"
Check3.Value = 0
CmbDPName.Text = ""
Frame5.Enabled = True
mClientIDRow = 0
mDPID = 0
End Sub
Sub ClearTITabControl()
TxtTINo = ""
TxtTIDate = ""
TxtBalAmount = ""
TxtRecAmt = ""
TxtTDSAmt = ""
TxtBalRec = ""
TxtOtherDed = ""
TxtTIClient = ""
TxtTILocation = ""
End Sub


Private Sub TxtInstrumentAmount_LostFocus()

If TxtInstrumentAmount.Text = "" Then Exit Sub
If IsNumeric(TxtInstrumentAmount.Text) = False Or Val(TxtInstrumentAmount.Text) < 0 Then
   MsgBox "Invalid Numeric Format !!!"
   TxtInstrumentAmount.SetFocus
   Exit Sub
End If
TxtInstrumentAmount = Round(Val(TxtInstrumentAmount), 2)

If ChkInstrumentNo = False Then
   If LCase(CmbInstrumentType.Text) = "cheque" Then
      MsgBox "Cheque No already entered for Cash Receipt - " & TmpRs1!CRID & "-" & GetLocationName(TmpRs1!LocationID) & " "
   ElseIf LCase(CmbInstrumentType.Text) = "dd" Then
      MsgBox "DD No already entered for Cash Receipt - " & TmpRs1!CRID & "-" & GetLocationName(TmpRs1!LocationID) & " "
   ElseIf LCase(CmbInstrumentType.Text) = "rtgs" Then
      MsgBox "RTGS No already entered for Cash Receipt - " & TmpRs1!CRID & "-" & GetLocationName(TmpRs1!LocationID) & " "
   End If
   TxtInstrumentAmount.SetFocus
   TxtInstrumentNo = ""
   Exit Sub
End If

End Sub

Private Sub TxtInstrumentDate_LostFocus()
If LCase(CmbInstrumentType.Text) <> "cheque" And TxtInstrumentDate.Value > Date Then
   MsgBox "Future date not allowed!!"
   TxtInstrumentDate.SetFocus
   Exit Sub
End If
End Sub

Private Sub TxtInstrumentNo_LostFocus()
If TxtInstrumentNo = "" Then Exit Sub
If CmbLocationID.Text = "" Then
   MsgBox "Select Location!!!"
   TxtInstrumentNo = ""
   CmbLocationID.SetFocus
   Exit Sub
End If

If IsClean(TxtInstrumentNo) = False Then
   MsgBox "No special character allowed !!!"
   TxtInstrumentNo.SetFocus
   Exit Sub
End If

'TxtInstrumentNo = Trim(TxtInstrumentNo)
TxtInstrumentNo = Replace(TxtInstrumentNo, Chr(32), "")



If ChkInstrumentNo = False Then
   If LCase(CmbInstrumentType.Text) = "cheque" Then
      MsgBox "Cheque No already entered for Cash Receipt - " & TmpRs1!CRID & "-" & GetLocationName(TmpRs1!LocationID) & " "
   ElseIf LCase(CmbInstrumentType.Text) = "dd" Then
      MsgBox "DD No already entered for Cash Receipt - " & TmpRs1!NewCRID & "-" & GetLocationName(TmpRs1!LocationID) & " "
   ElseIf LCase(CmbInstrumentType.Text) = "rtgs" Then
      MsgBox "RTGS No already entered for Cash Receipt - " & TmpRs1!NewCRID & "-" & GetLocationName(TmpRs1!LocationID) & " "
   End If
   TxtInstrumentNo.SetFocus
   Exit Sub
End If

End Sub
Private Function ChkInstrumentNo() As Boolean
Dim Retval As Boolean
Retval = True

If TxtInstrumentNo = "" Then GoTo l 'Exit Function
If TxtInstrumentAmount = "" Then GoTo l 'Exit Function
If CmbLocationID.Text = "" Then GoTo l 'Exit Function


'*******'
'Commented On 03 Feb 2012 Mail By Mayur Sir and Ashok Sir for checking Duli on No, Location and Amount

'tmp = "Select * from Txn_NewCashReceipt Where InstrumentType <> 'Cash' and InstrumentNo = '" & TxtInstrumentNo & "' "
'*******'

tmp = "Select * from Txn_NewCashReceipt Where InstrumentType <> 'Cash' and InstrumentNo = '" & TxtInstrumentNo & "' and InstrumentAmount = " & Val(TxtInstrumentAmount) & " And LocationID = " & mLocationID & " And RejectFlag <> 'Y' And ReturnFlag <> 'Y' "
Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   If Edit_Flg = "A" Then
      Retval = False
   Else
      If TxtCRID <> TmpRs1!NewCRID Then 'And mLocationID <> TmpRs1!LocationID Then
         Retval = False
      End If
   End If
   
End If

l: ChkInstrumentNo = Retval
End Function

Private Sub TxtRecAmt_LostFocus()
If TxtRecAmt.Text = "" Then Exit Sub
If IsNumeric(TxtRecAmt.Text) = False Or Val(TxtRecAmt.Text) < 0 Then
   MsgBox "Invalid Numeric Format !!!"
   TxtRecAmt.SetFocus
   Exit Sub
End If
TxtRecAmt = Round(Val(TxtRecAmt), 2)
If Val(TxtRecAmt) > Val(TxtBalAmount) Then
   MsgBox "Received Amount should be less than or equal to Balance Amount!"
   TxtRecAmt = 0
   TxtRecAmt.SetFocus
   Exit Sub
End If
'TxtBalRec.Text = Format(Val(TxtBalAmount) - (Val(TxtRecAmt) + Val(TxtTDSAmt)), "#####0.00")
TxtBalRec.Text = Format(Val(TxtBalAmount) - (Val(TxtRecAmt) + Val(TxtTDSAmt) + Val(TxtOtherDed)), "#####0.00")
If Val(TxtBalRec.Text) < 0 Then
   MsgBox "Balance Receivable cannot be less than zero. "
   TxtRecAmt = 0
   TxtRecAmt.SetFocus
End If
End Sub

Private Sub DeleteClientDetail()
tmp = "Delete From Txn_NewCashReceiptClientDetail Where NewCRID=" & Val(TxtCRID) & " And SM_Prefix='" & TxtPreFix & "'"
Call Tmp4Table
End Sub

Private Sub SaveClientDetail()
mErrMessage = "SaveClientDetail"

Call DeleteClientDetail
Call TableTxn_NewCashReceiptClientDetail(" Where NewCRID= " & Val(TxtCRID))
For I = 2 To MSHFlexGrid2.Rows - 1
    If MSHFlexGrid2.TextMatrix(I, 0) <> "" Then
       RsTxn_NewCashReceiptClientDetail.AddNew
       'RsTxn_NewCashReceiptClientDetail!CRCDID = getCRClientDetailMaxID
       RsTxn_NewCashReceiptClientDetail!NewCRID = Val(TxtCRID)
       RsTxn_NewCashReceiptClientDetail!SM_Prefix = TxtPreFix
       RsTxn_NewCashReceiptClientDetail!G_UID = GetGUID
       RsTxn_NewCashReceiptClientDetail!ExchangeTypeID = Val(MSHFlexGrid2.TextMatrix(I, 0))
       RsTxn_NewCashReceiptClientDetail!ShowDP = Val(MSHFlexGrid2.TextMatrix(I, 9))
       If Val(MSHFlexGrid2.TextMatrix(I, 9)) = 1 Then
          RsTxn_NewCashReceiptClientDetail!DPID = MSHFlexGrid2.TextMatrix(I, 8)
       Else
          RsTxn_NewCashReceiptClientDetail!ClientID = Val(MSHFlexGrid2.TextMatrix(I, 8))
       End If
       RsTxn_NewCashReceiptClientDetail!Remark = MSHFlexGrid2.TextMatrix(I, 7)
       RsTxn_NewCashReceiptClientDetail!NavCode = MSHFlexGrid2.TextMatrix(I, 10)
       If Val(MSHFlexGrid2.TextMatrix(I, 6)) > 0 Then
          RsTxn_NewCashReceiptClientDetail!Amount = Format(Val(MSHFlexGrid2.TextMatrix(I, 6)), "#####0.00")
          RsTxn_NewCashReceiptClientDetail!AdvNAVCode = MSHFlexGrid2.TextMatrix(I, 12)
       End If
       RsTxn_NewCashReceiptClientDetail.Update
       RsTxn_NewCashReceiptClientDetail.Requery
    End If
Next
End Sub
Function getCRClientDetailMaxID() As Double
Dim Retval As Double
Retval = 0
tmp = "select max(CRCDID) from Txn_NewCashReceiptClientDetail "
Call TmpTable
If TmpRs.RecordCount > 0 Then
   Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
End If
Retval = Retval + 1
TmpRs.Close
getCRClientDetailMaxID = Retval

End Function

Function getCRTIDetailMaxID() As Double
Dim Retval As Double
Retval = 0
tmp = "select max(CRTIDID) from Txn_NewCashReceiptTIDetail "
Call TmpTable
If TmpRs.RecordCount > 0 Then
   Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
End If
Retval = Retval + 1
TmpRs.Close
getCRTIDetailMaxID = Retval

End Function

Public Sub Indata()

Frame1.Enabled = False

PntPrv = "N"
Frame2.Enabled = False
Frame6.Visible = False
Frame8.Visible = False
Frame9.Visible = False

ButtonFrm.Enabled = True
Call ClearClientTabControl
Call ClearTITabControl
Frame5.Enabled = False
Frame7.Enabled = False
CmdReturn.Enabled = False
CmdApproveAM.Enabled = False

CmbLocationID.Visible = False
CmdLocationID.Visible = False
TxtLocationID.Visible = True
CmbLocationID.Text = ""

TxtCRID = RsTxn_NewCashReceipt!NewCRID
TxtCRDate = RsTxn_NewCashReceipt!CRDate
CmbInstrumentType = RsTxn_NewCashReceipt!InstrumentType

If LCase(CmbInstrumentType.Text) = "cheque" Then
   Label5.Caption = "Cheque No"
   Label7.Caption = "Cheque Date"
ElseIf LCase(CmbInstrumentType.Text) = "dd" Then
   Label5.Caption = "DD No"
   Label7.Caption = "DD Date"
ElseIf LCase(CmbInstrumentType.Text) = "rtgs" Then
   Label5.Caption = "RTGS No"
   Label7.Caption = "RTGS Date"
Else
   Label5.Caption = "Instrument No"
   Label7.Caption = "Instrument Date"
End If

TxtDrawnOn = RsTxn_NewCashReceipt!DrawnOn
TxtBranchName = RsTxn_NewCashReceipt!BranchName
TxtInstrumentNo = RsTxn_NewCashReceipt!InstrumentNo
TxtInstrumentDate = RsTxn_NewCashReceipt!InstrumentDate
ChkAdv.Value = IIf(IsNull(RsTxn_NewCashReceipt!AdvanceFlag), 0, RsTxn_NewCashReceipt!AdvanceFlag)
TxtInstrumentAmount = RsTxn_NewCashReceipt!InstrumentAmount
TxtChequeReturnChargesAmount = IIf(IsNull(RsTxn_NewCashReceipt!ChequeReturnChargesAmount), 0, RsTxn_NewCashReceipt!ChequeReturnChargesAmount)
TxtReceivedFrom = RsTxn_NewCashReceipt!ReceivedFrom
TxtRemark = RsTxn_NewCashReceipt!Remark
TxtPreFix = RsTxn_NewCashReceipt!SM_Prefix
mLocationID = IIf(IsNull(RsTxn_NewCashReceipt!LocationID), 0, RsTxn_NewCashReceipt!LocationID)
TxtLocationID = GetLocationName(mLocationID)
TxtG_UID = RsTxn_NewCashReceipt!G_UID
LblDepoDate.Visible = False
LblPostDate.Visible = False
TxtFAGPostingDate.Visible = False
TxtDepDate.Visible = False
mPrintFlag = IIf(IsNull(RsTxn_NewCashReceipt!Flag), "", RsTxn_NewCashReceipt!Flag)
mFAGBankID = IIf(IsNull(RsTxn_NewCashReceipt!FAGBankID), 0, RsTxn_NewCashReceipt!FAGBankID)
TxtBankName.Visible = True
CmbBankName.Visible = False
TxtFAGRemarks.Locked = True
mFlagA = IIf(IsNull(RsTxn_NewCashReceipt!FlagApproval), "", RsTxn_NewCashReceipt!FlagApproval)

If mFAGBankID <> 0 Then
   TxtBankName = GetFAGBankName(mFAGBankID, "N") & "~" & GetFAGBankName(mFAGBankID, "AC")
   
   LblDepoDate.Visible = True
   TxtDepDate.Visible = True
   If IsNull(RsTxn_NewCashReceipt!DepositDate) = False Then
      TxtDepDate.Value = RsTxn_NewCashReceipt!DepositDate
'   Else
'      TxtDepDate.Value = Date 'RsTxn_NewCashReceipt!DepositDate
   End If
   If IsNull(RsTxn_NewCashReceipt!FAGPostingDate) = False Then 'Or RsTxn_NewCashReceipt!BatchID <> 0 Then
      LblPostDate.Visible = True
      TxtFAGPostingDate.Visible = True
      TxtFAGPostingDate = RsTxn_NewCashReceipt!FAGPostingDate
      If LCase(Gen_UserRole) = "fag" Or LCase(Gen_UserRole) = "power" Or LCase(Gen_UserRole) = "fagadmin" Then
         If RsTxn_NewCashReceipt!ReturnFlag = "N" And (LCase(CmbInstrumentType.Text) = "cheque" Or LCase(CmbInstrumentType.Text) = "rtgs" Or LCase(CmbInstrumentType.Text) = "cash") Then
            CmdReturn.Enabled = True
         End If
      End If
   End If
   TxtFAGRemarks = IIf(IsNull(RsTxn_NewCashReceipt!FAGRemarks), "", RsTxn_NewCashReceipt!FAGRemarks)
Else
   TxtBankName = "" 'GetFAGBankName(mFAGBankID, "N") & "~" & GetFAGBankName(mFAGBankID, "AC")
   TxtFAGRemarks = ""
End If


If IsNull(RsTxn_NewCashReceipt!SightedDate) = False Then
   TxtSightedDate.Text = Format(RsTxn_NewCashReceipt!SightedDate, "dd-MMM-yyyy")
Else
   TxtSightedDate.Text = ""
End If
If RsTxn_NewCashReceipt!ReturnFlag = "Y" Then
   TxtReturnFlag = "Yes"
Else
   TxtReturnFlag = "No"
End If
If RsTxn_NewCashReceipt!RejectFlag = "Y" Then
   TxtRejectFlag = "Yes"
   CmdReject.Enabled = False
Else
   TxtRejectFlag = "No"
   CmdReject.Enabled = True
End If
TxtRejectedBy = IIf(IsNull(RsTxn_NewCashReceipt!RejectedBy), "", RsTxn_NewCashReceipt!RejectedBy) & " :- " & IIf(IsNull(RsTxn_NewCashReceipt!RejectDate), "", Format(RsTxn_NewCashReceipt!RejectDate, "dd-mmm-yyyy hh:mm:ss"))
TxtCreated = IIf(IsNull(RsTxn_NewCashReceipt!CreatedBy), "", RsTxn_NewCashReceipt!CreatedBy) & " :- " & IIf(IsNull(RsTxn_NewCashReceipt!CreatedDate), "", Format(RsTxn_NewCashReceipt!CreatedDate, "dd-mmm-yyyy hh:mm:ss"))
TxtUpdated = IIf(IsNull(RsTxn_NewCashReceipt!UpdatedBy), "", RsTxn_NewCashReceipt!UpdatedBy) & " :- " & IIf(IsNull(RsTxn_NewCashReceipt!UpdatedDate), "", Format(RsTxn_NewCashReceipt!UpdatedDate, "dd-mmm-yyyy hh:mm:ss"))
SSTab1.Tab = 0
TxtFAGUpdated = IIf(IsNull(RsTxn_NewCashReceipt!FAGUpdatedBy), "", RsTxn_NewCashReceipt!FAGUpdatedBy) & " :- " & IIf(IsNull(RsTxn_NewCashReceipt!FAGUpdatedDate), "", Format(RsTxn_NewCashReceipt!FAGUpdatedDate, "dd-mmm-yyyy hh:mm:ss"))
TxtG_UID = RsTxn_NewCashReceipt!G_UID
'Call ChkAdv_Click
Call InDataClientDetails
Call InDataTIDetails

'Call InDataDepositDetails(IIf(IsNull(RsTxn_NewCashReceipt!CRDepositID), 0, RsTxn_NewCashReceipt!CRDepositID))

If RsTxn_NewCashReceipt.RecordCount = 1 Then
   Call GButtonLock_1(Me, True)
Else
   Call GButtonLock(Me, True)
End If
CmdPrint.Enabled = True
CmdPrintPreview.Enabled = True
'CmdReject.Enabled = True

If mFlagA = "" Then
   If (LCase(Gen_UserRole) = "amsandrms" Or LCase(Gen_UserRole) = "headcmg") And CheckMultiClient = True Then
      CmdApproveAM.Enabled = True
   End If
End If
End Sub

Private Sub InDataClientDetails()
Dim ans As Variant
Call Grid_Head1
Call TableTxn_NewCashReceiptClientDetail(" Where NewCRID = " & Val(TxtCRID) & " And SM_Prefix='" & TxtPreFix & "'")
MSHFlexGrid2.Rows = RsTxn_NewCashReceiptClientDetail.RecordCount + 3
For I = 2 To MSHFlexGrid2.Rows - 2
    MSHFlexGrid2.TextMatrix(I, 0) = RsTxn_NewCashReceiptClientDetail!ExchangeTypeID
    MSHFlexGrid2.TextMatrix(I, 1) = GetExchangeName(RsTxn_NewCashReceiptClientDetail!ExchangeTypeID)
    MSHFlexGrid2.TextMatrix(I, 9) = IIf(IsNull(RsTxn_NewCashReceiptClientDetail!ShowDP), 0, RsTxn_NewCashReceiptClientDetail!ShowDP)
    MSHFlexGrid2.TextMatrix(I, 10) = IIf(IsNull(RsTxn_NewCashReceiptClientDetail!NavCode), "", RsTxn_NewCashReceiptClientDetail!NavCode)
    If MSHFlexGrid2.TextMatrix(I, 9) = 0 Then
       ans = GetClientNameByRow(RsTxn_NewCashReceiptClientDetail!ClientID)
       MSHFlexGrid2.TextMatrix(I, 2) = ans(0)
       MSHFlexGrid2.TextMatrix(I, 3) = ans(1) 'TmpRs!ClientName
       MSHFlexGrid2.TextMatrix(I, 4) = ans(2) 'TmpRs!DepositoryID
       MSHFlexGrid2.TextMatrix(I, 5) = ans(3) 'TmpRs!DepositoryName
       MSHFlexGrid2.TextMatrix(I, 8) = RsTxn_NewCashReceiptClientDetail!ClientID
       
       If Trim(MSHFlexGrid2.TextMatrix(I, 10)) = "" Then
          tmp = "select NAVCode,RNAVCode from mst_Client Where ClientID=" & RsTxn_NewCashReceiptClientDetail!ClientID
          Call Tmp4Table
          If TmpRs4.RecordCount > 0 Then
             MSHFlexGrid2.TextMatrix(I, 10) = IIf(IsNull(TmpRs4!NavCode), "", TmpRs4!NavCode)
             MSHFlexGrid2.TextMatrix(I, 11) = IIf(IsNull(TmpRs4!RNAVCode), "", TmpRs4!RNAVCode)
          End If
       End If
    Else
       MSHFlexGrid2.TextMatrix(I, 4) = RsTxn_NewCashReceiptClientDetail!DPID
       MSHFlexGrid2.TextMatrix(I, 5) = GetDepositoryName(RsTxn_NewCashReceiptClientDetail!DPID, RsTxn_NewCashReceiptClientDetail!ExchangeTypeID)
       MSHFlexGrid2.TextMatrix(I, 8) = RsTxn_NewCashReceiptClientDetail!DPID
       If Trim(MSHFlexGrid2.TextMatrix(I, 10)) = "" Then
          tmp = "select NAVCode from Mst_Depository Where DepositoryID='" & RsTxn_NewCashReceiptClientDetail!DPID & "'"
          Call Tmp4Table
          If TmpRs4.RecordCount > 0 Then
             MSHFlexGrid2.TextMatrix(I, 10) = IIf(IsNull(TmpRs4!NavCode), "", TmpRs4!NavCode)
          End If
       End If
    End If
    MSHFlexGrid2.TextMatrix(I, 12) = IIf(IsNull(RsTxn_NewCashReceiptClientDetail!AdvNAVCode), "", RsTxn_NewCashReceiptClientDetail!AdvNAVCode)
    MSHFlexGrid2.TextMatrix(I, 6) = IIf(IsNull(RsTxn_NewCashReceiptClientDetail!Amount), "", RsTxn_NewCashReceiptClientDetail!Amount)
    MSHFlexGrid2.TextMatrix(I, 7) = IIf(IsNull(RsTxn_NewCashReceiptClientDetail!Remark), "", RsTxn_NewCashReceiptClientDetail!Remark)
    MSHFlexGrid2.TextMatrix(0, 6) = Val(MSHFlexGrid2.TextMatrix(0, 6)) + Val(MSHFlexGrid2.TextMatrix(I, 6))
    
    RsTxn_NewCashReceiptClientDetail.MoveNext
Next

End Sub

Private Sub InDataTIDetails()
Dim ans As Variant
Call Grid_Head2

'Call TableTxn_NewCashReceiptTIDetail(" Where CRID = " & Val(TxtCRID))

tmp = " Select TTI.TINo,TTI.TIDate,TTI.TIAmount,"
tmp = tmp & " TI.RecAmount,TI.RecTDS,TI.RecOtherDed,TI.TIID,TI.ClientID,ClientIDRow,ClientName,TI.DPID,DepositoryName,TTI.PostInvoiceNo,TTI.EntryNo,TI.CRAAID,ML.LocationName,TTI.NAVCode "
tmp = tmp & " From Txn_NewCashReceiptTIDetail TI"
tmp = tmp & " inner Join  Txn_TIDetails TTI on TI.TIID=TTI.TIID"
tmp = tmp & " Left Join Mst_Client MC on TI.ClientID=MC.ClientID"
tmp = tmp & " Left Join Mst_Depository MD on TI.DPID=MD.DepositoryID"
tmp = tmp & " Left Join Mst_Location ML on TTI.LocationID=ML.LocationID"
tmp = tmp & " Where TI.NewCRID = " & Val(TxtCRID) & " And TI.SM_Prefix='" & TxtPreFix & "'"

Call Tmp4Table

Wc1 = ""
Wc1 = " And TI.TIID in (Select TIID from Txn_NewCashReceiptTIDetail TCRT"
Wc1 = Wc1 & " Where TCRT.NewCRID =  " & Val(TxtCRID) & " and TCRT.SM_Prefix= '" & TxtPreFix.Text & "')"

Call GetClientsRemInvoices(Wc1, TxtG_UID, "", , "show")

MSHFlexGrid3.Rows = TmpRs4.RecordCount + 3
For I = 2 To MSHFlexGrid3.Rows - 2
    MSHFlexGrid3.TextMatrix(I, 0) = TmpRs4!TINo
    MSHFlexGrid3.TextMatrix(I, 1) = Format(TmpRs4!TIDate, "dd-MMM-yyyy")
    MSHFlexGrid3.TextMatrix(I, 2) = TmpRs4!TIAmount
    
    If TmpRs.RecordCount > 0 Then TmpRs.MoveFirst
    TmpRs.Find "TIID=" & TmpRs4!TIID
    If TmpRs.EOF = False Then
       MSHFlexGrid3.TextMatrix(I, 3) = IIf(IsNull(TmpRs!RecAmount), 0, TmpRs!RecAmount)
       MSHFlexGrid3.TextMatrix(I, 3) = Val(MSHFlexGrid3.TextMatrix(I, 3)) + TmpRs!AdjAmountFAG
       MSHFlexGrid3.TextMatrix(I, 4) = IIf(IsNull(TmpRs!RecTDS), 0, TmpRs!RecTDS)
       MSHFlexGrid3.TextMatrix(I, 5) = IIf(IsNull(TmpRs!RecOtherDed), 0, TmpRs!RecOtherDed)
    Else
       MSHFlexGrid3.TextMatrix(I, 3) = 0
       MSHFlexGrid3.TextMatrix(I, 4) = 0
       MSHFlexGrid3.TextMatrix(I, 5) = 0
    End If
    
    MSHFlexGrid3.TextMatrix(I, 6) = Format(Val(MSHFlexGrid3.TextMatrix(I, 2)) - (Val(MSHFlexGrid3.TextMatrix(I, 3)) + Val(MSHFlexGrid3.TextMatrix(I, 4)) + Val(MSHFlexGrid3.TextMatrix(I, 5))), "#####0.00")
    MSHFlexGrid3.TextMatrix(I, 7) = TmpRs4!RecAmount
    MSHFlexGrid3.TextMatrix(I, 8) = TmpRs4!RecTDS
    MSHFlexGrid3.TextMatrix(I, 9) = IIf(IsNull(TmpRs4!RecOtherDed), 0, TmpRs4!RecOtherDed)
    MSHFlexGrid3.TextMatrix(I, 10) = Format(Val(MSHFlexGrid3.TextMatrix(I, 6)) - (Val(MSHFlexGrid3.TextMatrix(I, 7)) + Val(MSHFlexGrid3.TextMatrix(I, 8)) + Val(MSHFlexGrid3.TextMatrix(I, 9))), "#####0.00")
    MSHFlexGrid3.TextMatrix(I, 11) = TmpRs4!TIID
    If IsNull(TmpRs4!ClientID) = True Then
       MSHFlexGrid3.TextMatrix(I, 12) = IIf(IsNull(TmpRs4!DPID), "", TmpRs4!DPID)
       MSHFlexGrid3.TextMatrix(I, 13) = IIf(IsNull(TmpRs4!DPID), "", TmpRs4!DPID)
       MSHFlexGrid3.TextMatrix(I, 14) = IIf(IsNull(TmpRs4!DepositoryName), "", TmpRs4!DepositoryName)
    Else
       MSHFlexGrid3.TextMatrix(I, 12) = IIf(IsNull(TmpRs4!ClientID), "", TmpRs4!ClientID)
       MSHFlexGrid3.TextMatrix(I, 13) = IIf(IsNull(TmpRs4!ClientIDRow), "", TmpRs4!ClientIDRow)
       MSHFlexGrid3.TextMatrix(I, 14) = IIf(IsNull(TmpRs4!ClientName), "", TmpRs4!ClientName)
    End If
    
    MSHFlexGrid3.TextMatrix(I, 15) = IIf(IsNull(TmpRs4!PostInvoiceNo), 0, TmpRs4!PostInvoiceNo)
    MSHFlexGrid3.TextMatrix(I, 16) = IIf(IsNull(TmpRs4!EntryNo), 0, TmpRs4!EntryNo)

    MSHFlexGrid3.TextMatrix(I, 18) = IIf(IsNull(TmpRs4!CRAAID), "", TmpRs4!CRAAID)
    MSHFlexGrid3.TextMatrix(I, 19) = IIf(IsNull(TmpRs4!Locationname), "", TmpRs4!Locationname)
    MSHFlexGrid3.TextMatrix(I, 20) = IIf(IsNull(TmpRs4!NavCode), "", TmpRs4!NavCode)
    
    MSHFlexGrid3.TextMatrix(0, 7) = Val(MSHFlexGrid3.TextMatrix(0, 7)) + Val(MSHFlexGrid3.TextMatrix(I, 7))
    MSHFlexGrid3.TextMatrix(0, 8) = Val(MSHFlexGrid3.TextMatrix(0, 8)) + Val(MSHFlexGrid3.TextMatrix(I, 8))
    'If MSHFlexGrid3.TextMatrix(I, 18) = "" Then
    '   MSHFlexGrid3.TextMatrix(0, 7) = Val(MSHFlexGrid3.TextMatrix(0, 7)) + Val(MSHFlexGrid3.TextMatrix(I, 7))
    'End If
    MSHFlexGrid3.TextMatrix(0, 9) = Val(MSHFlexGrid3.TextMatrix(0, 9)) + Val(MSHFlexGrid3.TextMatrix(I, 9))
    
    TmpRs4.MoveNext
Next

End Sub

Private Function GetCRID() As Double
Dim Retval As Double
tmp = "Select max(CRID) from Txn_NewCashReceipt Where SM_Prefix='" & TxtPreFix & "'"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetCRID = Retval
End Function
Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsTxn_NewCashReceipt.MoveNext
If RsTxn_NewCashReceipt.EOF Then
   RsTxn_NewCashReceipt.MoveLast
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
RsTxn_NewCashReceipt.MovePrevious
If RsTxn_NewCashReceipt.BOF Then
   RsTxn_NewCashReceipt.MoveFirst
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
RsTxn_NewCashReceipt.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsTxn_NewCashReceipt.MoveLast
Call Indata
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False
End Sub
Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub Grid_Head()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 35

MSHFlexGrid1.TextMatrix(0, 0) = "Receipt ID"
MSHFlexGrid1.TextMatrix(0, 1) = "Location"
MSHFlexGrid1.TextMatrix(0, 2) = "Receipt Date"
MSHFlexGrid1.TextMatrix(0, 3) = "Type"
MSHFlexGrid1.TextMatrix(0, 4) = "Drawn On"
MSHFlexGrid1.TextMatrix(0, 5) = "Branch Name"
MSHFlexGrid1.TextMatrix(0, 6) = "Instrument No"
MSHFlexGrid1.TextMatrix(0, 7) = "Instrument Date"
MSHFlexGrid1.TextMatrix(0, 8) = "Instrument Amount"
MSHFlexGrid1.TextMatrix(0, 9) = "Received From"
MSHFlexGrid1.TextMatrix(0, 10) = "Remark"
MSHFlexGrid1.TextMatrix(0, 11) = "Print Flag"
MSHFlexGrid1.TextMatrix(0, 12) = "Hologram No"
MSHFlexGrid1.TextMatrix(0, 13) = "G_UID"
MSHFlexGrid1.TextMatrix(0, 14) = "Cash Receipt Deposit ID"
MSHFlexGrid1.TextMatrix(0, 15) = "Deposit Date"
MSHFlexGrid1.TextMatrix(0, 16) = "FAG Posting Date"
MSHFlexGrid1.TextMatrix(0, 17) = "FAG Remarks"
MSHFlexGrid1.TextMatrix(0, 18) = "Bank"
MSHFlexGrid1.TextMatrix(0, 19) = "Advance Flag"
MSHFlexGrid1.TextMatrix(0, 20) = "Return Flag"
MSHFlexGrid1.TextMatrix(0, 21) = "Batch ID"
MSHFlexGrid1.TextMatrix(0, 22) = "Reject Flag"
MSHFlexGrid1.TextMatrix(0, 23) = "Reject Date"
MSHFlexGrid1.TextMatrix(0, 24) = "Reject Remarks"
MSHFlexGrid1.TextMatrix(0, 25) = "Created By"
MSHFlexGrid1.TextMatrix(0, 26) = "Created On"
MSHFlexGrid1.TextMatrix(0, 27) = "Updated By"
MSHFlexGrid1.TextMatrix(0, 28) = "Updated On"
MSHFlexGrid1.TextMatrix(0, 29) = "FAG Updated By"
MSHFlexGrid1.TextMatrix(0, 30) = "FAG Updated On"
MSHFlexGrid1.TextMatrix(0, 31) = "Return Batch ID"
MSHFlexGrid1.TextMatrix(0, 32) = "AMs Approved By"
MSHFlexGrid1.TextMatrix(0, 33) = "AMs Approved On"
MSHFlexGrid1.TextMatrix(0, 34) = "AMs Approve Remarks"

MSHFlexGrid1.ColWidth(0) = 1200
MSHFlexGrid1.ColWidth(1) = 1000
MSHFlexGrid1.ColWidth(2) = 1200
MSHFlexGrid1.ColWidth(3) = 1400
MSHFlexGrid1.ColWidth(4) = 1800
MSHFlexGrid1.ColWidth(5) = 1800
MSHFlexGrid1.ColWidth(6) = 1600
MSHFlexGrid1.ColWidth(7) = 1400
MSHFlexGrid1.ColWidth(8) = 1800
MSHFlexGrid1.ColWidth(9) = 1800
MSHFlexGrid1.ColWidth(10) = 3400
MSHFlexGrid1.ColWidth(11) = 1400
MSHFlexGrid1.ColWidth(12) = 1400
MSHFlexGrid1.ColWidth(13) = 0
MSHFlexGrid1.ColWidth(14) = 1400
MSHFlexGrid1.ColWidth(15) = 1400
MSHFlexGrid1.ColWidth(16) = 1400
MSHFlexGrid1.ColWidth(17) = 5000
MSHFlexGrid1.ColWidth(18) = 1000
MSHFlexGrid1.ColWidth(19) = 1000
MSHFlexGrid1.ColWidth(20) = 1000

MSHFlexGrid1.ColWidth(32) = 2000
MSHFlexGrid1.ColWidth(33) = 2000
MSHFlexGrid1.ColWidth(34) = 4000

MSHFlexGrid1.WordWrap = True
MSHFlexGrid1.RowHeight(0) = 600

End Sub

Private Sub CmbSLocation_GotFocus()
tmp = CmbSLocation.Text
'Call TableMst_Location(" Where LocationId > 0 ")

If Gen_WcLocation = "" Then
   Call TableMst_Location("")
Else
   Call TableMst_Location("where " & Gen_WcLocation)
End If

CmbSLocation.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbSLocation.AddItem RsMst_Location!Locationname
    RsMst_Location.MoveNext
Next
CmbSLocation.Text = tmp
End Sub

Private Sub CmbSLocation_LostFocus()
If CmbSLocation.Text = "" Then
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbSLocation.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection "
   CmbSLocation.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmbSInstrumentType_LostFocus()
If CmbSInstrumentType.Text = "" Then Exit Sub

If LCase(CmbSInstrumentType.Text) <> "cash" And LCase(CmbSInstrumentType.Text) <> "cheque" And LCase(CmbSInstrumentType.Text) <> "dd" And LCase(CmbSInstrumentType.Text) <> "rtgs" Then
   MsgBox "Invalid Selection"
   CmbSInstrumentType.SetFocus
   Exit Sub
End If

End Sub


Private Sub TxtAmountFrom_LostFocus()
If TxtAmountFrom = "" Then Exit Sub
If IsNumeric(TxtAmountFrom) = False Or Val(TxtAmountFrom) < 0 Then
   MsgBox "Invalid numeric format!!!"
   TxtAmountFrom.SetFocus
   Exit Sub
End If
End Sub

Private Sub TxtAmountTo_LostFocus()
If TxtAmountTo = "" Then Exit Sub
If IsNumeric(TxtAmountTo) = False Or Val(TxtAmountTo) < 0 Then
   MsgBox "Invalid numeric format!!!"
   TxtAmountTo.SetFocus
   Exit Sub
End If
End Sub

Private Sub TxtCRDate_LostFocus()
If TxtCRDate.Value > Date Then
   MsgBox "Future Cash Receipt Date is not allowed !!!! "
   TxtCRDate.SetFocus
End If

If Format(Gen_BlockingDate, "yyyy-mm") >= Format(TxtCRDate.Value, "yyyy-mm") Then
   MsgBox "Cash Receipt Date allowed for next month of " & Format(Gen_BlockingDate, "MMM-yyyy") & "!!! "
   TxtCRDate.SetFocus
   Exit Sub
End If

End Sub

Private Sub CmbSDepositedBank_GotFocus()
tmp = CmbSDepositedBank.Text
Call TableMst_FAGBank("")
CmbSDepositedBank.Clear
If RsMst_FAGBank.RecordCount = 0 Then
   MsgBox "No Bank found"
   CmbSDepositedBank.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_FAGBank.RecordCount
    CmbSDepositedBank.AddItem RsMst_FAGBank!BankName & "~" & RsMst_FAGBank!AccountNo
    RsMst_FAGBank.MoveNext
Next
CmbSDepositedBank.Text = tmp
End Sub
Private Sub CmbSDepositedBank_LostFocus()
Dim strSplitString() As String
Dim AccountNo_, BankName_ As String

If CmbSDepositedBank.Text = "" Then
   Exit Sub
End If

strSplitString = Split(CmbSDepositedBank, "~")

If UBound(strSplitString) <> 1 Then
   MsgBox "Invalid Selection!!!"
   CmbSDepositedBank.SetFocus
   Exit Sub
End If

BankName_ = Trim(strSplitString(0))
AccountNo_ = Trim(strSplitString(1))

RsMst_FAGBank.MoveFirst
RsMst_FAGBank.Find "AccountNo = '" & AccountNo_ & "'"

If RsMst_FAGBank.EOF Then
   MsgBox "Invalid Deposited Bank selection "
   CmbSDepositedBank.SetFocus
   Exit Sub
End If
mSDepoBankID = RsMst_FAGBank!FAGBankID
End Sub

Private Sub CmbSDepositBank_LostFocus()
If CmbSDepositBank.Text = "" Then Exit Sub
If CmbSDepositBank.Text = "With Bank Name" Or CmbSDepositBank.Text = "Without Bank Name" Then
   Exit Sub
End If
MsgBox "Invalid selection"
CmbSDepositBank.SetFocus
End Sub

Private Sub DeleteTIDetail()

tmp = "Select TIID,RecAmount,RecTDs,RecOtherDed From Txn_NewCashReceiptTIDetail Where NewCRID= " & Val(TxtCRID) & " and SM_Prefix='" & TxtPreFix & "'"
Call TmpTable

Call TableTxn_TIDetails(" where TIID in (Select TIID From Txn_NewCashReceiptTIDetail Where NewCRID= " & Val(TxtCRID) & " and SM_Prefix='" & TxtPreFix & "')")

For I = 1 To TmpRs.RecordCount
    If RsTxn_TIDetails.RecordCount > 0 Then RsTxn_TIDetails.MoveFirst
    RsTxn_TIDetails.Find "TIID=" & TmpRs!TIID
    If RsTxn_TIDetails.EOF = False Then
       RsTxn_TIDetails!RecAmount = IIf(IsNull(RsTxn_TIDetails!RecAmount), 0, RsTxn_TIDetails!RecAmount) - IIf(IsNull(TmpRs!RecAmount), 0, TmpRs!RecAmount)
       RsTxn_TIDetails!RecTDSAmount = IIf(IsNull(RsTxn_TIDetails!RecTDSAmount), 0, RsTxn_TIDetails!RecTDSAmount) - IIf(IsNull(TmpRs!RecTDS), 0, TmpRs!RecTDS)
       RsTxn_TIDetails!RecOtherDed = IIf(IsNull(RsTxn_TIDetails!RecOtherDed), 0, RsTxn_TIDetails!RecOtherDed) - IIf(IsNull(TmpRs!RecOtherDed), 0, TmpRs!RecOtherDed)
       RsTxn_TIDetails.Update
    End If
    TmpRs.MoveNext
Next

tmp = "Delete From Txn_NewCashReceiptTIDetail Where NewCRID=" & Val(TxtCRID) & "  And SM_Prefix='" & TxtPreFix & "'"
Call Tmp4Table

End Sub

Private Sub SaveTIDetail()

mErrMessage = "SaveTIDetail"

Call DeleteTIDetail

mTIIDs = "(0"
For I = 2 To MSHFlexGrid3.Rows - 1
    If MSHFlexGrid3.TextMatrix(I, 0) <> "" And (Val(MSHFlexGrid3.TextMatrix(I, 7)) > 0 Or Val(MSHFlexGrid3.TextMatrix(I, 8)) > 0 Or Val(MSHFlexGrid3.TextMatrix(I, 9)) > 0) Then
       mTIIDs = mTIIDs & "," & Val(MSHFlexGrid3.TextMatrix(I, 11))
    End If
Next
mTIIDs = mTIIDs & ")"


'If RsTxn_TIDetails.State = 1 Then RsTxn_TIDetails.Close

Call TableTxn_NewCashReceiptTIDetail(" Where NewCRID= " & Val(TxtCRID))

Call TableTxn_TIDetails(" Where TIID in " & mTIIDs & "")

For I = 2 To MSHFlexGrid3.Rows - 1
    If MSHFlexGrid3.TextMatrix(I, 0) <> "" And (Val(MSHFlexGrid3.TextMatrix(I, 7)) > 0 Or Val(MSHFlexGrid3.TextMatrix(I, 8)) > 0 Or Val(MSHFlexGrid3.TextMatrix(I, 9)) > 0) Then
       RsTxn_NewCashReceiptTIDetail.AddNew
'       RsTxn_NewCashReceiptTIDetail!CRTIDID = getCRTIDetailMaxID
       RsTxn_NewCashReceiptTIDetail!NewCRID = Val(TxtCRID)
       RsTxn_NewCashReceiptTIDetail!SM_Prefix = TxtPreFix
       RsTxn_NewCashReceiptTIDetail!G_UID = GetGUID
       RsTxn_NewCashReceiptTIDetail!TIID = Val(MSHFlexGrid3.TextMatrix(I, 11))
   '    RsTxn_NewCashReceiptTIDetail!TotRecAmt = Val(MSHFlexGrid3.TextMatrix(I, 3))
    '   RsTxn_NewCashReceiptTIDetail!TotTDSAmt = Val(MSHFlexGrid3.TextMatrix(I, 4))
       RsTxn_NewCashReceiptTIDetail!RecAmount = Val(MSHFlexGrid3.TextMatrix(I, 7))
       RsTxn_NewCashReceiptTIDetail!RecTDS = Val(MSHFlexGrid3.TextMatrix(I, 8))
       RsTxn_NewCashReceiptTIDetail!RecOtherDed = Val(MSHFlexGrid3.TextMatrix(I, 9))
       If MSHFlexGrid3.TextMatrix(I, 18) <> "" Then
          RsTxn_NewCashReceiptTIDetail!CRAAID = MSHFlexGrid3.TextMatrix(I, 18)
       Else
          RsTxn_NewCashReceiptTIDetail!CRAAID = Null
       End If
       If MSHFlexGrid3.TextMatrix(I, 17) = "Y" Then
          RsTxn_NewCashReceiptTIDetail!DPID = MSHFlexGrid3.TextMatrix(I, 12)
       Else
          RsTxn_NewCashReceiptTIDetail!ClientID = Val(MSHFlexGrid3.TextMatrix(I, 12))
       End If
       
       RsTxn_NewCashReceiptTIDetail.Update
       RsTxn_NewCashReceiptTIDetail.Requery
       If RsTxn_TIDetails.RecordCount > 0 Then RsTxn_TIDetails.MoveFirst
       RsTxn_TIDetails.Find "TIID=" & Val(MSHFlexGrid3.TextMatrix(I, 11))
       If RsTxn_TIDetails.EOF = False Then
          RsTxn_TIDetails!RecAmount = IIf(IsNull(RsTxn_TIDetails!RecAmount), 0, RsTxn_TIDetails!RecAmount) + Val(MSHFlexGrid3.TextMatrix(I, 7))
          RsTxn_TIDetails!RecTDSAmount = IIf(IsNull(RsTxn_TIDetails!RecTDSAmount), 0, RsTxn_TIDetails!RecTDSAmount) + Val(MSHFlexGrid3.TextMatrix(I, 8))
          RsTxn_TIDetails!RecOtherDed = IIf(IsNull(RsTxn_TIDetails!RecOtherDed), 0, RsTxn_TIDetails!RecOtherDed) + Val(MSHFlexGrid3.TextMatrix(I, 9))
          RsTxn_TIDetails.Update
       End If
    End If
Next
End Sub




Private Sub TxtTDSAmt_LostFocus()
If TxtTDSAmt.Text = "" Then Exit Sub
If IsNumeric(TxtTDSAmt.Text) = False Or Val(TxtTDSAmt.Text) < 0 Then
   MsgBox "Invalid Numeric Format !!!"
   TxtTDSAmt.SetFocus
   Exit Sub
End If
TxtTDSAmt = Round(Val(TxtTDSAmt), 2)
'TxtBalRec.Text = Format(Val(TxtBalAmount) - (Val(TxtRecAmt) + Val(TxtTDSAmt)), "#####0.00")
TxtBalRec.Text = Format(Val(TxtBalAmount) - (Val(TxtRecAmt) + Val(TxtTDSAmt) + Val(TxtOtherDed)), "#####0.00")
If Val(TxtBalRec.Text) < 0 Then
   MsgBox "Balance Receivable cannot be less than zero. "
   TxtTDSAmt.Text = 0
   TxtTDSAmt.SetFocus
   Exit Sub
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

Private Sub CmbDPName_GotFocus()
If CmbExchangeTypeID.Text = "" Then
   MsgBox "Please select Exchange Type!!!"
   CmbDPName.Clear
   CmbExchangeTypeID.SetFocus
   Exit Sub
End If
tmp = CmbDPName.Text

Call TableMst_Depository("Where ExchangeTypeId = " & mExchangeTypeID & " ")

CmbDPName.Clear
If RsMst_Depository.RecordCount = 0 Then
   MsgBox "No record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Depository.RecordCount
    CmbDPName.AddItem RsMst_Depository!DepositoryName & "~" & RsMst_Depository!DepositoryID
    RsMst_Depository.MoveNext
Next
CmbDPName.Text = tmp

End Sub

Private Sub CmbDPName_LostFocus()

Dim strSplitString() As String
Dim DPID_, DPName_ As String

If CmbDPName.Text = "" Then
   TxtCDPID = ""
   TxtCDPName = ""
   Exit Sub
End If
strSplitString = Split(CmbDPName, "~")
DPName_ = Trim(strSplitString(0))
DPID_ = Trim(strSplitString(1))

RsMst_Depository.MoveFirst
RsMst_Depository.Find "DepositoryID = '" & DPID_ & "'"
If RsMst_Depository.EOF Then
   MsgBox "Invalid selection "
   CmbDPName.SetFocus
   Exit Sub
End If

mDPID = RsMst_Depository!DepositoryID
mDPOldNAVCode = IIf(IsNull(RsMst_Depository!NavCode), "", RsMst_Depository!NavCode)
mChequeReturn = IIf(IsNull(RsMst_Depository!ChequeReturn), "", RsMst_Depository!ChequeReturn)
TxtCDPID = mDPID
TxtCDPName = DPName_ 'CmbDPName.Text

'CmbNAVCode.AddItem mDPNAVCode
'CmbNAVCode.Text = mDPNAVCode

End Sub

Private Sub Check3_Click()
If Check3.Value = 1 Then
   Label25.Visible = False
   Label26.Visible = False
   Label28.Visible = False
   CmbDPName.Visible = True
   CmbDPName.Width = 9400 'CmbNonNcdexClient.Width
   CmbDPName.Top = TxtClientID.Top 'CmbNonNcdexClient.Top
   CmbDPName.Left = TxtClientID.Left 'CmbNonNcdexClient.Left
   CmbCNonNcdexClient.Visible = False
   SSTab1.Tab = 0
Else
   Label25.Visible = True
   Label26.Visible = True
   Label28.Visible = True
'   CmbNonNcdexClient.Visible = True
'   CmbNonNcdexClient.Width = CmbDPName.Width
'   CmbNonNcdexClient.Top = CmbDPName.Top
'   CmbNonNcdexClient.Left = CmbDPName.Left
   CmbDPName.Visible = False
End If

End Sub

Private Sub InDataDepositDetails(CRDepositeID_ As Variant)
Call TableTxn_NewCRDepositDetails(" Where CRDepositID= " & CRDepositeID_)

If RsTxn_NewCRDepositDetails.RecordCount > 0 Then
   TxtDepositeDate.Text = Format(RsTxn_NewCRDepositDetails!DepositDate, "dd-MMM-yyyy")
   TxtBankName.Text = GetFAGBankName(RsTxn_NewCRDepositDetails!FAGBankID, "N") & "~" & GetFAGBankName(RsTxn_NewCRDepositDetails!FAGBankID, "AC")
   TxtRemark1.Text = RsTxn_NewCRDepositDetails!Remark
   If LCase(Gen_UserRole) = "fag" Or LCase(Gen_UserRole) = "power" Or LCase(Gen_UserRole) = "fagadmin" Then
      If RsTxn_NewCashReceipt!BatchID <> 0 And RsTxn_NewCashReceipt!ReturnFlag = "N" And (LCase(CmbInstrumentType.Text) = "cheque" Or LCase(CmbInstrumentType.Text) = "rtgs") Then
         CmdReturn.Enabled = True
      End If
   End If
Else
   TxtDepositeDate.Text = ""
   TxtBankName.Text = ""
   TxtRemark1.Text = ""
End If
End Sub

Private Sub CmdPrintPreview_Click()

'Gen_mTimeStamp = GetTimeStamp
'Call Create_Xls(Gen_mTimeStamp & Gen_UserName & "_NewCashReceipt-" & Trim(TxtPreFix) & "" & Trim(TxtCRID) & ".xls")
'Call Xls_Print_Rpt
'Call Xls_Print_Rpt_New
PntPrv = "Y"

Gen_mTimeStamp = GetTimeStamp
Call PDF_Print_Rpt


End Sub
Public Sub Create_Xls(FileName As String)
'Dim oExcel As Object
'Dim oBook As Object
'Dim oSheet As Object
'Dim Pat As String
Pat = App.Path
Set oExcel = CreateObject("Excel.Application")
Set oBook = oExcel.Workbooks.Add
Gen_FLName = FileName
oBook.SaveAs (Pat & "\Excel\" & FileName)

Set oSheet = oBook.Worksheets(1)
With oExcel
'.Application.Sheets.Delete
.SheetsInNewWorkbook = 15 'MSHFlexGrid2.Rows - 2
.Workbooks.Add
End With
'oBook.SaveAs (Pat & "\" & FileName)
oExcel.Quit
'oExcel.Close
End Sub

Public Sub Xls_Print_Rpt_New()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim oWS As Excel.Worksheet
Dim mRow As Integer
'Dim mCol As Integer
Dim ans As Variant
tmp = "Select max(ReturnCharges) ReturnCharges from Mst_ChequeReturnReason"
Call TmpTable
mMaxReturnCharges = 0
If TmpRs.RecordCount > 0 Then
    mMaxReturnCharges = IIf(IsNull(TmpRs!ReturnCharges), 0, TmpRs!ReturnCharges)
End If
MsgBox ("Wait till next message")
Set oXL = New Excel.Application
'Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "_NewCashReceipt-" & Trim(TxtPreFix) & "" & Trim(TxtCRID) & ".xls")

I = 2

Set oWS = oWB.Worksheets("Sheet" & 1)
''oWS.PageSetup.RightMargin = 0.05
'oWS.PageSetup.TopMargin = 0.05
'oWS.PageSetup.RightMargin = 0.05
'oWS.PageSetup.TopMargin = 4.6
oWS.Name = TxtPreFix & "-" & TxtCRID  'MSHFlexGrid2.TextMatrix(i, 23)
oWS.ClearArrows
mRow = 6

oWS.Range("a1:g1").Merge
oWS.Range("a:a").ColumnWidth = 10.86
oWS.Range("b:b").ColumnWidth = 12.86
oWS.Range("c:c").ColumnWidth = 12.86
oWS.Range("d:d").ColumnWidth = 12.86
oWS.Range("e:e").ColumnWidth = 12.86 '32.14
oWS.Range("f:f").ColumnWidth = 12.86
oWS.Range("g:g").ColumnWidth = 12.86
oWS.Range("h:h").ColumnWidth = 12.86
oWS.Range("i:i").ColumnWidth = 12.86

tmp = "a" & mRow & ":" & "i" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "NATIONAL COLLATERAL MANAGEMENT SERVICES LIMITED"
oWS.Cells(mRow, C + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 10

'oWS.Cells(1, 1).Font.ColorIndex = 4
mRow = mRow + 1

tmp = "a" & mRow & ":" & "i" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "Location -" & TxtLocationID
oWS.Cells(mRow, C + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 10

mRow = mRow + 1

tmp = "a" & mRow & ":" & "i" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = CmbInstrumentType.Text & " Receipt"
oWS.Cells(mRow, C + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 10

mRow = mRow + 1

oWS.Cells(mRow, C + 1) = "Receipt No"
oWS.Cells(mRow, C + 1).Font.Size = 7
oWS.Cells(mRow, C + 1).Font.Bold = True

tmp = "b" & mRow & ":" & "c" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 2) = TxtPreFix & "-" & TxtCRID
oWS.Cells(mRow, C + 2).Font.Size = 7
oWS.Cells(mRow, C + 2).HorizontalAlignment = -4108

oWS.Cells(mRow, C + 4) = "Receipt Date"
oWS.Cells(mRow, C + 4).Font.Size = 7
oWS.Cells(mRow, C + 4).Font.Bold = True

tmp = "e" & mRow & ":" & "i" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 5) = " " & Format(TxtCRDate, "dd-mmm-yyyy")
oWS.Cells(mRow, C + 5).Font.Size = 7
oWS.Cells(mRow, C + 5).HorizontalAlignment = -4108


mRow = mRow + 1

Dim Rsp As String

Rsp = NumStrConv(Val(TxtInstrumentAmount))
'Rsp1 = NumStrConv(Val(TxtInstrumentAmount))
tmp = "a" & mRow & ":" & "i" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "Received with thanks from Shri " & TxtReceivedFrom & " a sum of Rs " & TxtInstrumentAmount & " ( Rupees " & Rsp & " )" & _
                         " as per the details enclosed, by " & CmbInstrumentType & " "
'No:................. (instrument no) dated ................. ( instrument date) drawn on ............................ (drawn on Bank) ........................ branch (drawn on branch)."
If CmbInstrumentType.Text <> "Cash" Then
   oWS.Cells(mRow, C + 1) = oWS.Cells(mRow, C + 1) & " No: " & TxtInstrumentNo & " dated " & TxtInstrumentDate & " drawn on " & TxtDrawnOn & " " & TxtBranchName & " branch (subject to realisation). "
End If
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).WrapText = True
oWS.Cells(mRow, C + 1).HorizontalAlignment = 2
oWS.Cells(mRow, C + 1).Font.Size = 10
oWS.Cells(mRow, C + 1).RowHeight = 60

mRow = mRow + 1
'mCol = 5

tmp = "a" & mRow & ":" & "d" & mRow
oWS.Range(tmp).Merge

tmp = "e" & mRow & ":" & "i" & mRow
oWS.Range(tmp).Merge

oWS.Cells(mRow, C + 5) = "Affix Hologram here"
oWS.Cells(mRow, C + 5).Font.Size = 7
oWS.Cells(mRow, C + 5).Font.Bold = True
oWS.Cells(mRow, C + 5).RowHeight = 45

mRow = mRow + 1

tmp = "a" & mRow & ":" & "d" & mRow
oWS.Range(tmp).Merge

oWS.Cells(mRow, C + 1) = "Name and Signature of authorised representative of NCMSL"
oWS.Cells(mRow, C + 1).Font.Size = 7
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).WrapText = True

tmp = "e" & mRow & ":" & "i" & mRow
oWS.Range(tmp).Merge

If IsNull(RsTxn_NewCashReceipt!HologramNo) Then
   oWS.Cells(mRow, C + 5) = "Hologram No:- Preview"
Else
  oWS.Cells(mRow, C + 5) = "Hologram No:- " & RsTxn_NewCashReceipt!HologramNo
End If
oWS.Cells(mRow, C + 5).Font.Size = 7
oWS.Cells(mRow, C + 5).Font.Bold = True
oWS.Cells(mRow, C + 5).RowHeight = 40

mRow = mRow + 1

oWS.Cells(mRow, C + 1) = "Created By :-"
oWS.Cells(mRow, C + 1).Font.Size = 7
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).WrapText = True

tmp = "b" & mRow & ":" & "c" & mRow
oWS.Range(tmp).Merge

Dim strSplitString() As String
Dim EmpNo_, CreatedDateTime_ As String

If TxtCreated <> "" Then
   strSplitString = Split(TxtCreated, ":-")
   EmpNo_ = Trim(strSplitString(0))
   CreatedDateTime_ = Trim(strSplitString(1))
Else
   EmpNo_ = "" 'Trim(strSplitString(0))
   CreatedDateTime_ = "" 'Trim(strSplitString(1))

End If

oWS.Cells(mRow, C + 2) = EmpNo_ & "-" & GetCreatedByName(EmpNo_)
oWS.Cells(mRow, C + 2).Font.Size = 7
oWS.Cells(mRow, C + 2).Font.Bold = True
oWS.Cells(mRow, C + 2).WrapText = True
oWS.Cells(mRow, C + 2).HorizontalAlignment = -4108

oWS.Cells(mRow, C + 4) = "Created On :-"
oWS.Cells(mRow, C + 4).Font.Size = 7
oWS.Cells(mRow, C + 4).Font.Bold = True
oWS.Cells(mRow, C + 4).WrapText = True

tmp = "e" & mRow & ":" & "i" & mRow
oWS.Range(tmp).Merge

oWS.Cells(mRow, C + 5) = CreatedDateTime_
oWS.Cells(mRow, C + 5).Font.Size = 7
oWS.Cells(mRow, C + 5).Font.Bold = True
oWS.Cells(mRow, C + 5).WrapText = True
oWS.Cells(mRow, C + 5).HorizontalAlignment = -4108

If LCase(CmbInstrumentType.Text) = "cheque" Then
   mRow = mRow + 1

   tmp = "a" & mRow & ":" & "i" & mRow
   oWS.Range(tmp).Merge
   oWS.Cells(mRow, C + 1) = "Note :cheque dishonour would attract charges of Rs. " & mMaxReturnCharges & "."
   oWS.Cells(mRow, C + 1).Font.Size = 7
   oWS.Cells(mRow, C + 1).Font.Bold = True
   oWS.Cells(mRow, C + 1).WrapText = True
End If
mRow = mRow + 1
tmp = "a" & mRow & ":" & "i" & mRow
oWS.Range(tmp).Merge


k = mRow

For Cell = 6 To k
    tmp = "a" & Cell & ":" & "i" & Cell
    oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
    oWS.Range(tmp).Font.Name = "Zurich BT"
    'oWS.Range(tmp).Font.Size = 7
    oWS.Range(tmp).WrapText = True
Next Cell

''
''If Gen_ExcelPassword <> "" Then
''   oWS.Protect Gen_ExcelPassword, , , , , , True, True
''End If



'Set oWS = oWB.Worksheets("Sheet" & 2)


'oWS.Name = "Cash Receipt Details (Invoice-CDTF related)"
mRow = mRow + 2

If ChkAdv.Value = 1 Then
   k = mRow
   'oWS.Name = "Advance Details"
   oWS.Range("a1:g1").Merge
   oWS.Range("a:a").ColumnWidth = 11
   oWS.Range("b:b").ColumnWidth = 11
   oWS.Range("c:c").ColumnWidth = 23
   oWS.Range("d:d").ColumnWidth = 11
   oWS.Range("e:e").ColumnWidth = 22 '32.14
   oWS.Range("f:f").ColumnWidth = 10
    
    mStart = mRow
   tmp = "a" & mRow & ":" & "f" & mRow
   oWS.Range(tmp).Merge
   oWS.Cells(mRow, C + 1) = "Cash Receipt Details (Advance related) for Receipt No: " & TxtPreFix & "-" & TxtCRID & " Dated " & Format(TxtCRDate, "dd-mmm-yyyy")
   'oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
   oWS.Cells(mRow, C + 1).Font.Bold = True
   oWS.Cells(mRow, C + 1).Font.Size = 10
    
   mRow = mRow + 1
    
   oWS.Cells(mRow, C + 1) = "Exchange Type"
   oWS.Cells(mRow, C + 1).Font.Size = 7
   oWS.Cells(mRow, C + 1).Font.Bold = True
   
   oWS.Cells(mRow, C + 2) = "Client ID"
   oWS.Cells(mRow, C + 2).Font.Size = 7
   oWS.Cells(mRow, C + 2).Font.Bold = True
   
   oWS.Cells(mRow, C + 3) = "Client Name"
   oWS.Cells(mRow, C + 3).Font.Size = 7
   oWS.Cells(mRow, C + 3).Font.Bold = True
   
   oWS.Cells(mRow, C + 4) = "DP ID"
   oWS.Cells(mRow, C + 4).Font.Size = 7
   oWS.Cells(mRow, C + 4).Font.Bold = True
   
   oWS.Cells(mRow, C + 5) = "DP Name"
   oWS.Cells(mRow, C + 5).Font.Size = 7
   oWS.Cells(mRow, C + 5).Font.Bold = True
   
   oWS.Cells(mRow, C + 6) = "Amount"
   oWS.Cells(mRow, C + 6).Font.Size = 7
   oWS.Cells(mRow, C + 6).Font.Bold = True
   
   mRow = mRow + 1
    
   For k = 2 To MSHFlexGrid2.Rows - 2
       oWS.Cells(mRow, C + 1) = MSHFlexGrid2.TextMatrix(k, 1)
       oWS.Cells(mRow, C + 1).Font.Size = 7
       
       oWS.Cells(mRow, C + 2) = MSHFlexGrid2.TextMatrix(k, 2)
       oWS.Cells(mRow, C + 2).Font.Size = 7
        
       oWS.Cells(mRow, C + 3) = MSHFlexGrid2.TextMatrix(k, 3)
       oWS.Cells(mRow, C + 3).Font.Size = 7
       
       If LCase(MSHFlexGrid2.TextMatrix(k, 1)) <> "non ncdex" Then
          oWS.Cells(mRow, C + 4) = MSHFlexGrid2.TextMatrix(k, 4)
          oWS.Cells(mRow, C + 4).Font.Size = 7
           
          oWS.Cells(mRow, C + 5) = MSHFlexGrid2.TextMatrix(k, 5)
          oWS.Cells(mRow, C + 5).Font.Size = 7
       Else
          oWS.Cells(mRow, C + 4) = "" 'MSHFlexGrid2.TextMatrix(k, 4)
          oWS.Cells(mRow, C + 4).Font.Size = 7
           
          oWS.Cells(mRow, C + 5) = "" 'MSHFlexGrid2.TextMatrix(k, 5)
          oWS.Cells(mRow, C + 5).Font.Size = 7
       End If
       

       
       oWS.Cells(mRow, C + 6) = MSHFlexGrid2.TextMatrix(k, 6)
       oWS.Cells(mRow, C + 6).Font.Size = 7
      
       mRow = mRow + 1
       
   Next k

   k = mRow

    For Cell = mStart To k
        tmp = "a" & Cell & ":" & "f" & Cell
        oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
        oWS.Range(tmp).Font.Name = "Zurich BT"
        'oWS.Range(tmp).Font.Size = 7
        oWS.Range(tmp).WrapText = True
    Next Cell
End If

mTIDetails = 0
If MSHFlexGrid3.Rows > 3 Then
   If Val(MSHFlexGrid3.TextMatrix(0, 7)) <> 0 Or Val(MSHFlexGrid3.TextMatrix(0, 8)) <> 0 Then
      mTIDetails = 1
   End If
End If
   
   
If ChkAdv.Value = 0 Or (ChkAdv.Value = 1 And mTIDetails = 1) Then
    'oWS.Name = "TI Details"
   ' oWS.Range("a1:g1").Merge
    'oWS.Range("a:a").ColumnWidth = 15
    'oWS.Range("b:b").ColumnWidth = 22
    'oWS.Range("c:c").ColumnWidth = 15
    'oWS.Range("d:d").ColumnWidth = 12
    'oWS.Range("e:e").ColumnWidth = 13
    'oWS.Range("f:f").ColumnWidth = 13
    'oWS.Range("g:g").ColumnWidth = 10 '32.14
    'oWS.Range("h:h").ColumnWidth = 13
    
    oWS.Range("a:a").ColumnWidth = 11
    oWS.Range("b:b").ColumnWidth = 16.86
    oWS.Range("c:c").ColumnWidth = 11
    oWS.Range("d:d").ColumnWidth = 10
    oWS.Range("e:e").ColumnWidth = 8
    oWS.Range("f:f").ColumnWidth = 8
    oWS.Range("g:g").ColumnWidth = 8 '32.14
    oWS.Range("h:h").ColumnWidth = 9
    oWS.Range("i:i").ColumnWidth = 11
        
    mStart = mRow
    tmp = "a" & mRow & ":" & "i" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 1) = "Cash Receipt Details (Invoice related) for Receipt No: " & TxtPreFix & "-" & TxtCRID & " Dated " & Format(TxtCRDate, "dd-mmm-yyyy")
    'oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
    oWS.Cells(mRow, C + 1).Font.Bold = True
    oWS.Cells(mRow, C + 1).Font.Size = 10
    
    mRow = mRow + 1
    'mCol = 5
    
    ''New
    oWS.Cells(mRow, C + 1).RowHeight = 30
    oWS.Cells(mRow, C + 1) = "ClientID /DPID"
    oWS.Cells(mRow, C + 1).Font.Size = 7
    oWS.Cells(mRow, C + 1).Font.Bold = True
    
    oWS.Cells(mRow, C + 2) = "Client Name"
    oWS.Cells(mRow, C + 2).Font.Size = 7
    oWS.Cells(mRow, C + 2).Font.Bold = True
    
    oWS.Cells(mRow, C + 3) = "Invoice No"
    oWS.Cells(mRow, C + 3).Font.Size = 7
    oWS.Cells(mRow, C + 3).Font.Bold = True
    
    oWS.Cells(mRow, C + 4) = "Invoice Date"
    oWS.Cells(mRow, C + 4).Font.Size = 7
    oWS.Cells(mRow, C + 4).Font.Bold = True
    
    oWS.Cells(mRow, C + 5) = "Invoice Amount"
    oWS.Cells(mRow, C + 5).Font.Size = 7
    oWS.Cells(mRow, C + 5).Font.Bold = True
    
    oWS.Cells(mRow, C + 6) = "Received Amount"
    oWS.Cells(mRow, C + 6).Font.Size = 7
    oWS.Cells(mRow, C + 6).Font.Bold = True
    
    oWS.Cells(mRow, C + 7) = "TDS Amount"
    oWS.Cells(mRow, C + 7).Font.Size = 7
    oWS.Cells(mRow, C + 7).Font.Bold = True
    
    oWS.Cells(mRow, C + 8) = "Other Deduction"
    oWS.Cells(mRow, C + 8).Font.Size = 7
    oWS.Cells(mRow, C + 8).Font.Bold = True
    
    oWS.Cells(mRow, C + 9) = "PostInvoice No"
    oWS.Cells(mRow, C + 9).Font.Size = 7
    oWS.Cells(mRow, C + 9).Font.Bold = True
    
    mRow = mRow + 1
    
    For k = 2 To MSHFlexGrid3.Rows - 2
    'k = 1
        If MSHFlexGrid3.TextMatrix(k, 18) = "" Then
           oWS.Cells(mRow, C + 1) = MSHFlexGrid3.TextMatrix(k, 13)
           oWS.Cells(mRow, C + 1).Font.Size = 7
            
           oWS.Cells(mRow, C + 2) = MSHFlexGrid3.TextMatrix(k, 14)
           oWS.Cells(mRow, C + 2).Font.Size = 7
          
           oWS.Cells(mRow, C + 3) = MSHFlexGrid3.TextMatrix(k, 0)
           oWS.Cells(mRow, C + 3).Font.Size = 7
            
           oWS.Cells(mRow, C + 4) = " " & Format(MSHFlexGrid3.TextMatrix(k, 1), "dd-mmm-yyyy")
           oWS.Cells(mRow, C + 4).Font.Size = 7
           
           oWS.Cells(mRow, C + 5) = Val(MSHFlexGrid3.TextMatrix(k, 2))
           oWS.Cells(mRow, C + 5).Font.Size = 7
            
           oWS.Cells(mRow, C + 6) = Val(MSHFlexGrid3.TextMatrix(k, 7))
           oWS.Cells(mRow, C + 6).Font.Size = 7
           
           oWS.Cells(mRow, C + 7) = Val(MSHFlexGrid3.TextMatrix(k, 8))
           oWS.Cells(mRow, C + 7).Font.Size = 7
           
           oWS.Cells(mRow, C + 8) = Val(MSHFlexGrid3.TextMatrix(k, 9))
           oWS.Cells(mRow, C + 8).Font.Size = 7
           
           oWS.Cells(mRow, C + 9) = MSHFlexGrid3.TextMatrix(k, 15)
           oWS.Cells(mRow, C + 9).Font.Size = 7
            
           mRow = mRow + 1
        End If
    Next k
    For Cell = mStart To mRow
        tmp = "a" & Cell & ":" & "i" & Cell
        oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
        oWS.Range(tmp).Font.Name = "Zurich BT"
        'oWS.Range(tmp).Font.Size = 7
        oWS.Range(tmp).WrapText = True
    Next Cell
End If


If Gen_ExcelPassword <> "" Then
   oWS.Protect Gen_ExcelPassword, , , , , , True, True
End If
oWB.Save

Call Log_Print_History(TxtPreFix, "CR", TxtCRID, Gen_FLName)

MsgBox ("Excel file created !!! ")
oXL.Visible = True

End Sub

Private Sub CmdCancel_Click()
TxtHologramNo = ""
Frame3.Visible = False
End Sub
Private Sub CmdPrint_Click()
'If RsTxn_NewCashReceipt!HologramNo = "" Or IsNull(RsTxn_NewCashReceipt!HologramNo) Then
'   If LCase(UserType) = "power" Then
'      Check2.Visible = True
'   End If
'   Frame3.Visible = True
'   TxtHologramNo.SetFocus
'   Call TableMst_Hologram
'   Exit Sub
'Else
'   TxtHologramNo = ""
'   Frame3.Visible = False
'End If
'RsTxn_NewCashReceipt!Flag = "P"
'RsTxn_NewCashReceipt.Update
'Gen_mTimeStamp = GetTimeStamp
'Call Create_Xls(Gen_mTimeStamp & Gen_UserName & "_NewCashReceipt-" & Trim(TxtPreFix) & "" & Trim(TxtCRID) & ".xls")
''Call Xls_Print_Rpt
'Call Xls_Print_Rpt_New

If RsTxn_NewCashReceipt!HologramNo = "" Or IsNull(RsTxn_NewCashReceipt!HologramNo) Then
   If LCase(UserType) = "power" Then
      Check2.Visible = True
   End If
   
   If PntPrv = "N" Then
      ans = MsgBox("We strongly recommend that you take a look at Print-Preview first before going ahead with this Print option. Do you still want to continue?", vbYesNo)
      If ans = vbNo Then
         Exit Sub
      End If
   End If
   
   Frame3.Visible = True
   TxtHologramNo.SetFocus
   Call TableMst_Hologram
   Exit Sub
Else
   TxtHologramNo = ""
   Frame3.Visible = False
End If

PntPrv = "N"
RsTxn_NewCashReceipt!Flag = "P"
RsTxn_NewCashReceipt.Update

Gen_mTimeStamp = GetTimeStamp
Call PDF_Print_Rpt


End Sub
Private Sub TxtHologramNo_LostFocus()
'If TxtHologramNo = "" Then Exit Sub
'If Check_HologramNo(TxtHologramNo) = False Then
'   TxtHologramNo.SetFocus
'   CmdSaveHologram.Enabled = False
'   Exit Sub
'End If
'CmdSaveHologram.Enabled = True
'CmdSaveHologram.SetFocus

If TxtHologramNo = "" Then Exit Sub
'If Check_HologramNo(TxtHologramNo) = False Then
'   TxtHologramNo.SetFocus
'   CmdSaveHologram.Enabled = False
'   Exit Sub
'End If


If TxtHologramNo <> "0" Then

   If Check_HologramNo(TxtHologramNo) = False Then
      TxtHologramNo.SetFocus
      CmdSaveHologram.Enabled = False
      Exit Sub
   End If
    
   If CheckHologramInRange(TxtHologramNo, mLocationID) = False Then
      MsgBox "Hologram No not found in any range."
      TxtHologramNo.SetFocus
      CmdSaveHologram.Enabled = False
      Exit Sub
   End If
End If

CmdSaveHologram.Enabled = True
CmdSaveHologram.SetFocus

End Sub
Private Sub CmdSaveHologram_Click()


'On Error GoTo msgError

Dim ans As Variant
If TxtHologramNo = "" Then
   MsgBox "Blank Hologram No not allowed!!!"
   TxtHologramNo.SetFocus
   Exit Sub
End If

ans = MsgBox("Are you sure you want to save this Hologram No??", vbYesNo)
If ans = vbYes Then
   If TxtHologramNo <> "0" Then
      Call TableMst_Hologram
      RsMst_Hologram.AddNew
      RsMst_Hologram!NewHologramNo = TxtHologramNo
      RsMst_Hologram!DocType = "Cash Receipt"
      RsMst_Hologram!DocNo = TxtCRID
      RsMst_Hologram!SM_Prefix = TxtPreFix
      RsMst_Hologram!CreatedBy = Gen_UserLoginID
      RsMst_Hologram.Update
      RsMst_Hologram.Close
   End If
   
   RsTxn_NewCashReceipt!HologramNo = TxtHologramNo
   RsTxn_NewCashReceipt!Flag = "P"
   RsTxn_NewCashReceipt.Update
   TxtHologramNo = ""
   Frame3.Visible = False
   Gen_mTimeStamp = GetTimeStamp
   If Check2.Value = 1 Then
      Call Create_Xls(Gen_mTimeStamp & Gen_UserName & "_NewCashReceipt-" & Trim(TxtPreFix) & "" & Trim(TxtCRID) & ".xls")
       'Call Xls_Print_Rpt
      Call Xls_Print_Rpt_New
   End If
Else
   TxtHologramNo.SetFocus
End If

Exit Sub
msgError:
MsgBox Err.Description, vbExclamation, "Error" '"Some error occurred during process!!!"

End Sub

Private Sub TxtOtherDed_LostFocus()
If TxtOtherDed.Text = "" Then Exit Sub
If IsNumeric(TxtOtherDed.Text) = False Or Val(TxtOtherDed.Text) < 0 Then
   MsgBox "Invalid Numeric Format !!!"
   TxtOtherDed.SetFocus
   Exit Sub
End If
TxtOtherDed = Round(Val(TxtOtherDed), 2)
TxtBalRec.Text = Format(Val(TxtBalAmount) - (Val(TxtRecAmt) + Val(TxtTDSAmt) + Val(TxtOtherDed)), "#####0.00")
If Val(TxtBalRec.Text) < 0 Then
   MsgBox "Balance Receivable cannot be less than zero."
   TxtOtherDed.Text = 0
   TxtOtherDed.SetFocus
    Exit Sub
End If
End Sub

Private Sub CmbSAdvance_LostFocus()
If CmbSAdvance.Text = "" Then Exit Sub
If LCase(CmbSAdvance.Text) <> "yes" And LCase(CmbSAdvance.Text) <> "no" Then
   MsgBox "Invalid Selection"
   CmbSAdvance.SetFocus
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
Private Sub CmbBankName_GotFocus()
tmp = CmbBankName.Text
Call TableMst_FAGBank("")
CmbBankName.Clear
If RsMst_FAGBank.RecordCount = 0 Then
   MsgBox "No Bank found"
   CmbBankName.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_FAGBank.RecordCount
    CmbBankName.AddItem RsMst_FAGBank!BankName & "~" & RsMst_FAGBank!AccountNo
    RsMst_FAGBank.MoveNext
Next
CmbBankName.Text = tmp
End Sub
Private Sub CmbBankName_LostFocus()
Dim strSplitString() As String
Dim AccountNo_, BankName_ As String

If CmbBankName.Text = "" Then
   mFAGBankID = 0
   LblDepoDate.Visible = False
   LblPostDate.Visible = False
   TxtFAGPostingDate.Visible = False
   TxtDepDate.Visible = False
   TxtFAGRemarks = ""
   TxtFAGRemarks.Locked = True
   Exit Sub
End If

strSplitString = Split(CmbBankName, "~")

If UBound(strSplitString) <> 1 Then
   MsgBox "Invalid Selection!!!"
   CmbBankName.SetFocus
   Exit Sub
End If

BankName_ = Trim(strSplitString(0))
AccountNo_ = Trim(strSplitString(1))

RsMst_FAGBank.MoveFirst
RsMst_FAGBank.Find "AccountNo = '" & AccountNo_ & "'"

If RsMst_FAGBank.EOF Then
   MsgBox "Invalid FAG Bank selection "
   CmbBankName.SetFocus
   Exit Sub
End If
mFAGBankID = RsMst_FAGBank!FAGBankID

LblPostDate.Visible = True
TxtFAGPostingDate.Visible = True
TxtFAGPostingDate.Enabled = True
If TxtDepDate.Visible = False Then
   TxtDepDate.Value = TxtCRDate.Value  'Commented on 03 Jan 2018 after mail from Deepali (FAG) Date
   TxtDepDate.Visible = True
   TxtDepDate.Enabled = True
   LblDepoDate.Visible = True
End If

TxtFAGPostingDate.Value = TxtCRDate.Value  'Commented on 03 Jan 2018 after mail from Deepali (FAG) Date

TxtFAGRemarks.Locked = False
TxtFAGRemarks = ""

End Sub

Private Function CheckDuplicateTINo(EntryNo_ As Double) As Boolean
Dim Retval As Boolean
Retval = False
For I = 2 To MSHFlexGrid3.Rows - 1
    If MSHFlexGrid3.TextMatrix(I, 0) <> "" Then
       If Val(MSHFlexGrid3.TextMatrix(I, 16)) = EntryNo_ Then
          Retval = True
          Exit For
       End If
    End If
Next
CheckDuplicateTINo = Retval
End Function

Private Sub AppendTIDetails()

Wc1 = ""
If LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "admin" And LCase(Gen_UserRole) <> "superadmin" And LCase(Gen_UserRole) <> "fag" And LCase(Gen_UserRole) <> "fagadmin" And LCase(Gen_UserRole) <> "headcmg" Then
   Wc1 = " And (TI.LocationID=" & mLocationID & " or TI.LocationID is null)"
End If

Wc1 = Wc1 & " And (TI.NAVCode in "
Wc1 = Wc1 & " (Select distinct TCRD.NAVCode from Txn_NewCashReceiptClientDetail TCRD"
Wc1 = Wc1 & " inner join Mst_Client MC on MC.ClientID=TCRD.ClientID"
Wc1 = Wc1 & " Where TCRD.NewCRID =  " & Val(TxtCRID) & " and TCRD.SM_Prefix= '" & TxtPreFix.Text & "') OR"
Wc1 = Wc1 & " TI.NAVCode in "
Wc1 = Wc1 & " (Select distinct MC.NAVCode from Txn_NewCashReceiptClientDetail TCRD"
Wc1 = Wc1 & " inner join Mst_Client MC on MC.ClientID=TCRD.ClientID"
Wc1 = Wc1 & " Where TCRD.NewCRID =  " & Val(TxtCRID) & " and TCRD.SM_Prefix= '" & TxtPreFix.Text & "') OR"
Wc1 = Wc1 & " TI.NAVCode in "
Wc1 = Wc1 & " (Select distinct RNAVCode from Txn_NewCashReceiptClientDetail TCRD"
Wc1 = Wc1 & " inner join Mst_Client MC on MC.ClientID=TCRD.ClientID"
Wc1 = Wc1 & " Where TCRD.NewCRID =  " & Val(TxtCRID) & " and TCRD.SM_Prefix= '" & TxtPreFix.Text & "') OR"
Wc1 = Wc1 & " TI.NAVCode in "
Wc1 = Wc1 & " (Select distinct TCRD.NAVCode from Txn_NewCashReceiptClientDetail TCRD"
Wc1 = Wc1 & " inner join Mst_Depository MC on MC.DepositoryID=TCRD.DPID"
Wc1 = Wc1 & " Where TCRD.NewCRID =  " & Val(TxtCRID) & " and TCRD.SM_Prefix= '" & TxtPreFix.Text & "'))"

Wc1 = Wc1 & " And TI.TIID not in (Select TIID from Txn_NewCashReceiptTIDetail TCRT"
Wc1 = Wc1 & " Where TCRT.NewCRID =  " & Val(TxtCRID) & " and TCRT.SM_Prefix= '" & TxtPreFix.Text & "')"

Call GetClientsRemInvoices(Wc1, TxtG_UID, "")

If TmpRs.RecordCount > 0 Then
   For gr = 1 To TmpRs.RecordCount
       If CheckDuplicateTINo(TmpRs!EntryNo) = False Then
          rn = MSHFlexGrid3.Rows - 1
          MSHFlexGrid3.TextMatrix(rn, 0) = TmpRs!TINo
          MSHFlexGrid3.TextMatrix(rn, 1) = IIf(IsNull(TmpRs!TIDate), "", Format(TmpRs!TIDate, "dd-MMM-yyyy"))
          MSHFlexGrid3.TextMatrix(rn, 2) = IIf(IsNull(TmpRs!TIAmount), 0, TmpRs!TIAmount)
          MSHFlexGrid3.TextMatrix(rn, 3) = IIf(IsNull(TmpRs!RecAmount), 0, TmpRs!RecAmount)
          MSHFlexGrid3.TextMatrix(rn, 4) = IIf(IsNull(TmpRs!RecTDS), 0, TmpRs!RecTDS)
          MSHFlexGrid3.TextMatrix(rn, 5) = IIf(IsNull(TmpRs!RecOtherDed), 0, TmpRs!RecOtherDed)
          MSHFlexGrid3.TextMatrix(rn, 6) = Format(Val(MSHFlexGrid3.TextMatrix(rn, 2)) - (Val(MSHFlexGrid3.TextMatrix(rn, 3)) + Val(MSHFlexGrid3.TextMatrix(rn, 4)) + Val(MSHFlexGrid3.TextMatrix(rn, 5))), "#####0.00")
          MSHFlexGrid3.TextMatrix(rn, 11) = TmpRs!TIID
          
          tmp = "Select ShowDP,DPID,TCRD.ClientID,MC.ClientID,MC.ClientName,DepositoryName,Amount"
          tmp = tmp & " From Txn_NewCashReceiptClientDetail TCRD"
          tmp = tmp & " left join Mst_Client MC on MC.ClientID=TCRD.ClientID"
          tmp = tmp & " Left join Mst_Depository MD on MD.DepositoryID=TCRD.DPID"
          tmp = tmp & " Where  TCRD.NewCRID = " & Val(TxtCRID) & " and TCRD.SM_Prefix= '" & TxtPreFix & "' And ( MC.NAVCode='" & TmpRs!NavCode & "' or MC.RNAVCode='" & TmpRs!NavCode & "' or MD.NAVCode='" & TmpRs!NavCode & "')"
       
          Call Tmp6Table
       
          If TmpRs6.RecordCount > 0 Then
             If TmpRs6!ShowDP = 0 Then
                MSHFlexGrid3.TextMatrix(rn, 12) = TmpRs6!ClientID
                MSHFlexGrid3.TextMatrix(rn, 13) = TmpRs6!ClientID
                MSHFlexGrid3.TextMatrix(rn, 14) = TmpRs6!ClientName
                MSHFlexGrid3.TextMatrix(rn, 17) = "N"
             Else
                MSHFlexGrid3.TextMatrix(rn, 12) = IIf(IsNull(TmpRs6!DPID), "", TmpRs6!DPID)
                MSHFlexGrid3.TextMatrix(rn, 13) = IIf(IsNull(TmpRs6!DPID), "", TmpRs6!DPID)
                MSHFlexGrid3.TextMatrix(rn, 14) = IIf(IsNull(TmpRs6!DepositoryName), "", TmpRs6!DepositoryName)
                MSHFlexGrid3.TextMatrix(rn, 17) = "Y"
             End If
          End If
       
                  
          MSHFlexGrid3.TextMatrix(rn, 15) = IIf(IsNull(TmpRs!PostInvoiceNo), 0, TmpRs!PostInvoiceNo)
          MSHFlexGrid3.TextMatrix(rn, 16) = IIf(IsNull(TmpRs!EntryNo), 0, TmpRs!EntryNo)
'          MSHFlexGrid3.TextMatrix(rn, 18) = IIf(IsNull(TmpRs!CRAAID), 0, TmpRs!CRAAID)
          MSHFlexGrid3.TextMatrix(rn, 19) = IIf(IsNull(TmpRs!Locationname), "", TmpRs!Locationname)
          MSHFlexGrid3.Rows = MSHFlexGrid3.Rows + 1
       End If
       TmpRs.MoveNext
   Next
End If
End Sub

'Private Function GetAllocatedAmt(mClntDPID As Variant) As Double
'Dim mRetval As Double
'mRetval = 0
'
'tmp = "Select sum(RecAmount) from Txn_NewCashReceiptTIDetail Where CRID=" & Val(TxtCRID) & " and SM_Prefix='" & TxtPreFix & "' "
'tmp = tmp & " And isnull(CRAAID,0)<>0 And (ClientID=" & Val(mClntDPID) & " or DPID='" & mClntDPID & "')"
'
'Call TmpTable
'
'If TmpRs.RecordCount > 0 Then
'   mRetval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
'End If
'GetAllocatedAmt = mRetval
'End Function

Private Sub PDF_Print_Rpt()
Dim mHeadingTop  As Double
Dim mHeadingLeft  As Double
Dim ans As Variant
Dim a1 As Variant
Set objPDF = New mjwPDF

I = 2
' Create a simple PDF file using the mjwPDF class

Gen_mTimeStamp = GetTimeStamp
'Call Create_Xls(Gen_mTimeStamp & Gen_UserName & "_CDTF-" & Trim(TxtPreFix) & "" & Trim(TxtCxTFNo) & ".xls")
    
    
' Set the PDF title and filename
objPDF.PDFTitle = "Cash Receipt PDF Document"
objPDF.PDFFileName = App.Path & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "_Cash Receipt-" & Trim(TxtPreFix) & "" & Trim(TxtCRID) & ".pdf"
mPDFFileName = App.Path & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "_Cash Receipt-" & Trim(TxtPreFix) & "" & Trim(TxtCRID) & ".pdf"
' We must tell the class where the PDF fonts are located
objPDF.PDFLoadAfm = App.Path & "\Fonts"

' Set the file properties
objPDF.PDFSetLayoutMode = LAYOUT_DEFAULT
objPDF.PDFFormatPage = FORMAT_A4
objPDF.PDFOrientation = ORIENT_PORTRAIT
objPDF.PDFSetUnit = UNIT_PT
objPDF.PDFSetViewerPreferences = VIEW_HIDEMENUBAR & VIEW_HIDETOOLBAR & VIEW_HIDEWINDOWUI & VIEW_HIDERIGHTCLICK '& VIEW_DISABLEFILEMENU


' Lets us set see the bookmark pane when we view the PDF
'objPDF.PDFUseOutlines = True

' View the PDF file after we create it
objPDF.PDFView = True
'    objPDF.PDFWatermark = "NCMSL"
' Begin our PDF document

objPDF.PDFBeginDoc
    ' Lets add a heading
objPDF.PDFSetFont FONT_ARIAL, 12, FONT_BOLD
objPDF.PDFSetDrawColor = vbWhite
objPDF.PDFSetTextColor = vbBlack
objPDF.PDFSetAlignement = ALIGN_CENTER
objPDF.PDFSetBorder = BORDER_ALL
objPDF.PDFSetFill = True

objPDF.PDFCell " ", 15, 15, _
objPDF.PDFGetPageWidth - 30, objPDF.PDFGetPageHeight - 20

If PntPrv = "Y" Then
   objPDF.PDFWatermark = "Preview"
Else
   objPDF.PDFWatermark = "NCML" '"NCMSL"
End If

objPDF.PDFSetFont FONT_ARIAL, 12, FONT_BOLD
objPDF.PDFSetDrawColor = vbWhite
objPDF.PDFSetTextColor = vbBlack
objPDF.PDFSetAlignement = ALIGN_CENTER
objPDF.PDFSetBorder = BORDER_ALL
objPDF.PDFSetFill = True
   
pgval = ((objPDF.PDFGetPageWidth - 18))
  
mRowHeight = 15
mGlobalLeft = 12
'objPDF.PDFDrawPolygon Array(12, 15, objPDF.PDFGetPageWidth - 18, 15, objPDF.PDFGetPageWidth - 18, 30, 12, 30)
mLeft = mGlobalLeft
mTop = 15
mHeight = mTop + mRowHeight
mTxtFromTop = mRowHeight + 2
 
 'Call DrawTable(10, 2, 15, 30, 12)

Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)

mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
mStrLength = Len("NATIONAL COLLATERAL MANAGEMENT SERVICES LIMITED")
mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
mTxtLeft = mPrintPoint - (mStrLength * 2) '* 5) '270
mFontSize = 12

Call PrintText(mFontSize, FONT_NORMAL, ALIGN_CENTER, mTxtLeft - 10, mTxtTop, "NATIONAL COLLATERAL MANAGEMENT SERVICES LIMITED")

mHeight = mTop + mRowHeight
mLeft = mGlobalLeft

Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)

mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
mStrLength = Len("Location - " & TxtLocationID)
mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
mTxtLeft = mPrintPoint - (mStrLength * 2) '* 5) '270
mFontSize = 10

Call PrintText(mFontSize, FONT_BOLD, ALIGN_CENTER, mTxtLeft - 26, mTxtTop - 3, "Location - " & TxtLocationID)

mHeight = mTop + mRowHeight
mLeft = mGlobalLeft

Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)

mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
mStrLength = Len(CmbInstrumentType.Text & " Receipt")
mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
mTxtLeft = mPrintPoint - (mStrLength * 2) '* 5) '270
mFontSize = 10

Call PrintText(mFontSize, FONT_BOLD, ALIGN_CENTER, mTxtLeft - 26, mTxtTop - 3, CmbInstrumentType.Text & " Receipt")

mHeight = mTop + mRowHeight
mLeft = mGlobalLeft

''mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
''mStrLength = Len("Depositor Details")
''mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
''mTxtLeft = mPrintPoint '- (mStrLength) '* 5) '270
mFontSize = 7
Dim mDispTextSize  As Integer
mDispTextSize = 7

mTxtTop = mTxtTop + mRowHeight
Dim Rsp As String
Rsp = NumStrConv(Val(TxtInstrumentAmount))

Call DrawTable(, 4, mTop, mHeight, mLeft, mRowHeight)

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Receipt No")
Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop - 5, TxtPreFix & "-" & TxtCRID)
'--
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 280, mTxtTop - 5, "Receipt Date")
Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 430, mTxtTop - 5, Format(TxtCRDate, "dd-MMM-yyyy"))
'--

mTxtTop = mTxtTop + mRowHeight + mRowHeight

mstr = "Received with thanks from Shri " & TxtReceivedFrom & " a sum of Rs " & TxtInstrumentAmount & " ( Rupees " & Rsp & " ) as per the details enclosed, by " & CmbInstrumentType & ""
If CmbInstrumentType.Text <> "Cash" Then
   mstr = mstr & " No: " & TxtInstrumentNo & " dated " & Format(TxtInstrumentDate.Value, "dd-MMM-yyyy") & " drawn on " & TxtDrawnOn & " " & TxtBranchName & " branch (subject to realisation). "
End If

'------------------------------------------------------------


Dim mStrAdd As String
Dim mTotWords As Variant
mMax = 22 'mMax = 38
mPara = 0 'For Row Height
t1 = 1
mTotWords = CountWords(mstr)
'mLen = CountWords(TxtAddress) 'Len(TxtAddress)
mLen = RoundUp(mTotWords / mMax) '+ 1
If mLen = 0 Then
   mLen = 1
End If
mStrAdd = ""
a = 0
For T = 1 To mLen
    mStrAdd = ""
    If T > 1 Then
       t1 = mMax + t1
       mPara = mPara + 10
       a = t1 - 1
       If mTotWords - a < 23 Then
          For x = a To mTotWords - 1
             mStrAdd = mStrAdd & mStrw(x) & " "
             a = a + 1
          Next
       Else
          For x = a To a + 21
             mStrAdd = mStrAdd & mStrw(x) & " "
             a = a + 1
          Next
       End If
    Else
       If mTotWords < 23 Then
          For x = 0 To mTotWords - 1
             mStrAdd = mStrAdd & mStrw(x) & " "
          Next
       Else
          For x = 0 To 21
             mStrAdd = mStrAdd & mStrw(x) & " "
          Next
       End If
    End If
    'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_CENTER, mTxtLeft, mTxtTop + mPara, Mid$(TxtAddress, t1, mMax))
    
    Call PrintText(mFontSize, FONT_NORMAL, ALIGN_CENTER, mGlobalLeft + 5, mTxtTop + mPara - 5 - 7, mStrAdd)
Next



'-------------------------------------------------------------
'on 26 July 2012
'mMax = 133
'mPara = 7
't1 = 1
'mLen = Len(mstr)
'
'If mLen > mMax Then
'   mLen = RoundUp(mLen / mMax)
'Else
'   mLen = 1
'End If
'For t = 1 To mLen
'    If t > 1 Then
'       t1 = mMax + t1
'       mPara = mPara + 9
'    End If
'    'Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 20, mTxtTop + mPara - 5 - 7, Mid$(mstr, t1, mMax))
'
'    Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 5, mTxtTop + mPara - 5 - 7, Mid$(mstr, t1, mMax))
'Next

'-------------------------------------------------------------


a1 = mTxtTop + mPara - 5

''Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 20, mTxtTop - 5, mstr)
''
''Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 20, mTxtTop - 5, mstr)

objPDF.PDFDrawLine mGlobalLeft, Val(a1), objPDF.PDFGetPageWidth - 18, Val(a1)

mTxtTop = a1 + mRowHeight + mRowHeight + mRowHeight

Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 430, mTxtTop - 5, "Affix Hologram here")

objPDF.PDFDrawLine mGlobalLeft + 421, mTxtTop, mGlobalLeft + 421, mTxtTop - (mRowHeight + mRowHeight + mRowHeight)

objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop


mTxtTop = mTxtTop + mRowHeight + mRowHeight

Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Name and Signature of authorised representative of NCMSL")

Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 430, mTxtTop - 5, "Hologram No: " & RsTxn_NewCashReceipt!HologramNo)

objPDF.PDFDrawLine mGlobalLeft + 421, mTxtTop, mGlobalLeft + 421, mTxtTop - (mRowHeight + mRowHeight)

objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop

mTxtTop = mTxtTop + mRowHeight

Dim strSplitString() As String
Dim EmpNo_, CreatedDateTime_ As String

If TxtCreated <> "" Then
   strSplitString = Split(TxtCreated, ":-")
   EmpNo_ = Trim(strSplitString(0))
   CreatedDateTime_ = Trim(strSplitString(1))
Else
   EmpNo_ = "" 'Trim(strSplitString(0))
   CreatedDateTime_ = "" 'Trim(strSplitString(1))
End If

''oWS.Cells(mRow, c + 2) = EmpNo_ & "-" & GetCreatedByName(EmpNo_)
Call DrawTable(, 4, mTxtTop, mTxtTop - mRowHeight, mLeft, mRowHeight)
mTxtTop = mTxtTop + mRowHeight
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Created By : ")
Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop - 5, EmpNo_ & "-" & GetCreatedByName(EmpNo_))
'--
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 280, mTxtTop - 5, "Created On : ")
Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 430, mTxtTop - 5, CreatedDateTime_)
'--
If LCase(CmbInstrumentType.Text) = "cheque" Then
mTxtTop = mTxtTop + mRowHeight
Call PrintText(mDispTextSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Note :cheque dishonour would attract charges of RS. 500.")
objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop
End If
mTxtTop = mTxtTop + mRowHeight + mRowHeight
objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop



If Val(MSHFlexGrid3.TextMatrix(0, 7)) > 0 Then
    mTxtTop = mTxtTop + mRowHeight
    Call PrintText(mDispTextSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Cash Receipt Details (Invoice related) for Receipt No: " & TxtPreFix & "-" & TxtCRID & " Dated " & Format(TxtCRDate, "dd-mmm-yyyy"))
    objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop
    mTxtTop = mTxtTop + mRowHeight
    Call PrintText(mDispTextSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "ClientID /DPID")
    Call PrintText(mDispTextSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 60, mTxtTop - 5, "Client Name")
    Call PrintText(mDispTextSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 170, mTxtTop - 5, "Invoice No")
    Call PrintText(mDispTextSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 230, mTxtTop - 5, "Invoice Date")
    Call PrintText(mDispTextSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 300, mTxtTop - 5, "Invoice Amount")
    Call PrintText(mDispTextSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 360, mTxtTop - 5, "Received Amount")
    Call PrintText(mDispTextSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 430, mTxtTop - 5, "TDS Amount")
    Call PrintText(mDispTextSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 490, mTxtTop - 5, "Other Deduction")
    
    objPDF.PDFDrawLine mGlobalLeft + 58, mTxtTop - mRowHeight, mGlobalLeft + 58, mTxtTop
    objPDF.PDFDrawLine mGlobalLeft + 168, mTxtTop - mRowHeight, mGlobalLeft + 168, mTxtTop
    objPDF.PDFDrawLine mGlobalLeft + 228, mTxtTop - mRowHeight, mGlobalLeft + 228, mTxtTop
    objPDF.PDFDrawLine mGlobalLeft + 298, mTxtTop - mRowHeight, mGlobalLeft + 298, mTxtTop
    objPDF.PDFDrawLine mGlobalLeft + 358, mTxtTop - mRowHeight, mGlobalLeft + 358, mTxtTop
    objPDF.PDFDrawLine mGlobalLeft + 428, mTxtTop - mRowHeight, mGlobalLeft + 428, mTxtTop
    objPDF.PDFDrawLine mGlobalLeft + 488, mTxtTop - mRowHeight, mGlobalLeft + 488, mTxtTop
    
    objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop
     mTxtTop = mTxtTop + (mRowHeight)
     
    For k = 2 To MSHFlexGrid3.Rows - 2
        If MSHFlexGrid3.TextMatrix(k, 18) = "" Then
           If mTxtTop > objPDF.PDFGetPageHeight Then
              Call CheckForNewPage
           End If
           Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, MSHFlexGrid3.TextMatrix(k, 13))
           mMax = 20
           mPara = 7
           t1 = 1
           mLen = Len(MSHFlexGrid3.TextMatrix(k, 14))
            
           If mLen > mMax Then
              mLen = RoundUp(mLen / mMax)
           Else
              mLen = 1
           End If
           For T = 1 To mLen
               If T > 1 Then
                  t1 = mMax + t1
                  mPara = mPara + 7
               End If
               Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 60, mTxtTop + mPara - 5 - 7, mID$(MSHFlexGrid3.TextMatrix(k, 14), t1, mMax))
           Next
            
           a1 = mTxtTop + mPara - 5
        
            'Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 60, mTxtTop - 5, MSHFlexGrid3.TextMatrix(k, 14))
           
'           Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 170, mTxtTop - 5, MSHFlexGrid3.TextMatrix(k, 0))
           
           mMax = 8
           mPara = 4
           t1 = 1
           mLen = Len(MSHFlexGrid3.TextMatrix(k, 0))
            
           If mLen > mMax Then
              mLen = RoundUp(mLen / mMax)
           Else
              mLen = 1
           End If
           For T = 1 To mLen
               If T > 1 Then
                  t1 = mMax + t1
                  mPara = mPara + 7
               End If
               Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 170, mTxtTop + mPara - 2 - 7, mID$(MSHFlexGrid3.TextMatrix(k, 0), t1, mMax))
           Next
            
           a1 = mTxtTop + mPara - 2
           
           
           
           Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 230, mTxtTop - 5, Format(MSHFlexGrid3.TextMatrix(k, 1), "dd-MMM-yyyy"))
           Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 300, mTxtTop - 5, Format(MSHFlexGrid3.TextMatrix(k, 2), "#####0.00"))
           Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 360, mTxtTop - 5, Format(MSHFlexGrid3.TextMatrix(k, 7), "#####0.00"))
           Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 430, mTxtTop - 5, Format(MSHFlexGrid3.TextMatrix(k, 8), "#####0.00"))
           Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 490, mTxtTop - 5, Format(MSHFlexGrid3.TextMatrix(k, 9), "#####0.00"))
            
           objPDF.PDFDrawLine mGlobalLeft + 58, mTxtTop - mRowHeight, mGlobalLeft + 58, Val(a1)
           objPDF.PDFDrawLine mGlobalLeft + 168, mTxtTop - mRowHeight, mGlobalLeft + 168, Val(a1)
           objPDF.PDFDrawLine mGlobalLeft + 228, mTxtTop - mRowHeight, mGlobalLeft + 228, Val(a1)
           objPDF.PDFDrawLine mGlobalLeft + 298, mTxtTop - mRowHeight, mGlobalLeft + 298, Val(a1)
           objPDF.PDFDrawLine mGlobalLeft + 358, mTxtTop - mRowHeight, mGlobalLeft + 358, Val(a1)
           objPDF.PDFDrawLine mGlobalLeft + 428, mTxtTop - mRowHeight, mGlobalLeft + 428, Val(a1)
           objPDF.PDFDrawLine mGlobalLeft + 488, mTxtTop - mRowHeight, mGlobalLeft + 488, Val(a1)
           objPDF.PDFDrawLine mGlobalLeft, Val(a1), objPDF.PDFGetPageWidth - 18, Val(a1)
           
           mTxtTop = a1 + mRowHeight
        End If
     Next
End If

If Val(MSHFlexGrid2.TextMatrix(0, 6)) > 0 Then
    mTxtTop = mTxtTop + mRowHeight
     objPDF.PDFDrawLine mGlobalLeft, mTxtTop - mRowHeight, objPDF.PDFGetPageWidth - 18, mTxtTop - mRowHeight
    Call PrintText(mDispTextSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Cash Receipt Details (Advance related) for Receipt No: " & TxtPreFix & "-" & TxtCRID & " Dated " & Format(TxtCRDate, "dd-mmm-yyyy"))
    objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop
    mTxtTop = mTxtTop + mRowHeight
    Call PrintText(mDispTextSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "ExchangeType")
    Call PrintText(mDispTextSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 60, mTxtTop - 5, "Client ID")
    Call PrintText(mDispTextSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 130, mTxtTop - 5, "Client Name")
    Call PrintText(mDispTextSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 300, mTxtTop - 5, "DP ID")
    Call PrintText(mDispTextSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 360, mTxtTop - 5, "DP Name")
    Call PrintText(mDispTextSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 490, mTxtTop - 5, "Advance Amount")
        
    objPDF.PDFDrawLine mGlobalLeft + 58, mTxtTop - mRowHeight, mGlobalLeft + 58, mTxtTop
    objPDF.PDFDrawLine mGlobalLeft + 128, mTxtTop - mRowHeight, mGlobalLeft + 128, mTxtTop
'    objPDF.PDFDrawLine mGlobalLeft + 228, mTxtTop - mRowHeight, mGlobalLeft + 228, mTxtTop
    objPDF.PDFDrawLine mGlobalLeft + 298, mTxtTop - mRowHeight, mGlobalLeft + 298, mTxtTop
    objPDF.PDFDrawLine mGlobalLeft + 358, mTxtTop - mRowHeight, mGlobalLeft + 358, mTxtTop
    objPDF.PDFDrawLine mGlobalLeft + 488, mTxtTop - mRowHeight, mGlobalLeft + 488, mTxtTop
    
    objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop
     mTxtTop = mTxtTop + (mRowHeight)
     
    For k = 2 To MSHFlexGrid2.Rows - 2
        If mTxtTop > objPDF.PDFGetPageHeight Then
           Call CheckForNewPage
        End If
        Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, MSHFlexGrid2.TextMatrix(k, 1))
        Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 60, mTxtTop - 5, MSHFlexGrid2.TextMatrix(k, 2))
        
        mMax = 35 '45
        mPara = 7
        t1 = 1
        mLen = Len(MSHFlexGrid2.TextMatrix(k, 3))
        
        If mLen > mMax Then
           mLen = RoundUp(mLen / mMax)
        Else
           mLen = 1
        End If
        For T = 1 To mLen
            If T > 1 Then
               t1 = mMax + t1
               mPara = mPara + 7
            End If
            Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 130, mTxtTop + mPara - 5 - 7, mID$(MSHFlexGrid2.TextMatrix(k, 3), t1, mMax))
        Next
        
        a1 = mTxtTop + mPara - 3
    

        Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 300, mTxtTop - 5, MSHFlexGrid2.TextMatrix(k, 4))
        Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 360, mTxtTop - 5, MSHFlexGrid2.TextMatrix(k, 5))
        Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 490, mTxtTop - 5, Format(MSHFlexGrid2.TextMatrix(k, 6), "#####0.00"))
    
'         mTxtTop = mTxtTop + (mRowHeight)
    
        objPDF.PDFDrawLine mGlobalLeft + 58, mTxtTop - mRowHeight, mGlobalLeft + 58, Val(a1) 'mTxtTop
        objPDF.PDFDrawLine mGlobalLeft + 128, mTxtTop - mRowHeight, mGlobalLeft + 128, Val(a1) 'mTxtTop
    '    objPDF.PDFDrawLine mGlobalLeft + 228, mTxtTop - mRowHeight, mGlobalLeft + 228, mTxtTop
        objPDF.PDFDrawLine mGlobalLeft + 298, mTxtTop - mRowHeight, mGlobalLeft + 298, Val(a1) 'mTxtTop
        objPDF.PDFDrawLine mGlobalLeft + 358, mTxtTop - mRowHeight, mGlobalLeft + 358, Val(a1) 'mTxtTop
        objPDF.PDFDrawLine mGlobalLeft + 488, mTxtTop - mRowHeight, mGlobalLeft + 488, Val(a1) 'mTxtTop
        
        objPDF.PDFDrawLine mGlobalLeft, Val(a1), objPDF.PDFGetPageWidth - 18, Val(a1)
        
        mTxtTop = mTxtTop + (mRowHeight)
    Next
End If

objPDF.PDFEndDoc

End Sub


Private Sub CheckForNewPage()
   objPDF.PDFWatermark = ""
   objPDF.PDFEndPage
   objPDF.PDFNewPage
   objPDF.PDFSetFont FONT_ARIAL, 12, FONT_BOLD
   objPDF.PDFSetDrawColor = vbWhite
   objPDF.PDFSetTextColor = vbBlack
   objPDF.PDFSetAlignement = ALIGN_CENTER
   objPDF.PDFSetBorder = BORDER_ALL
   objPDF.PDFSetFill = True

   objPDF.PDFCell " ", 15, 15, _
   objPDF.PDFGetPageWidth - 30, objPDF.PDFGetPageHeight - 20
   If PntPrv = "Y" Then
      objPDF.PDFWatermark = "Preview"
   Else
      objPDF.PDFWatermark = "NCML" '"NCMSL"
   End If
   
   objPDF.PDFSetFont FONT_ARIAL, 12, FONT_BOLD
   objPDF.PDFSetDrawColor = vbWhite
   objPDF.PDFSetTextColor = vbBlack
   objPDF.PDFSetAlignement = ALIGN_CENTER
   objPDF.PDFSetBorder = BORDER_ALL
   objPDF.PDFSetFill = True

   mLeft = mGlobalLeft
   mTxtTop = 30
   mHeight = mTxtTop + mRowHeight
   mTxtFromTop = mRowHeight + 2
End Sub

Private Sub DrawTable(Optional mROw_ As Double, Optional mCol_ As Double, Optional mTop_ As Double, Optional mHeight_ As Double, Optional mLeft_ As Double, Optional mRowHeight_ As Double)

If mROw_ = 0 Then mROw_ = 1
If mCol_ = 0 Then mCol_ = 1
If mLeft_ = 0 Then mLeft_ = mmasterLeft
If mTop_ = 0 Then mTop_ = 0
If mHeight_ = 0 Then mHeight_ = mTop_ + mRowHeight_

Dim pgval_ As Double

For ir = 1 To mROw_
    pgval_ = pgval
    pgval_ = pgval_ / mCol_
    mWidth_ = pgval_
    For ic = 1 To mCol_
        objPDF.PDFDrawPolygon Array(mLeft_, mTop_, mWidth_, mTop_, mWidth_, mHeight_, mLeft_, mHeight_)
        mLeft_ = pgval_ * ic
        mWidth_ = pgval_ * (ic + 1)
    Next
    mTop_ = mHeight_
    mHeight_ = mHeight_ + mRowHeight_
    mLeft_ = mGlobalLeft
Next

mTop = mTop_
mRow = mROw_ + 1

End Sub
Private Sub PrintText(mFontSize_ As Integer, mFontFormat_ As PDFFontStl, mAlign_ As Variant, mTxtLeft_ As Double, mTxtTop_ As Double, mTexttoPrint As String)

'        objPDF.PDFSetFont FONT_ARIAL, 12, FONT_BOLD
'        objPDF.PDFSetAlignement = ALIGN_CENTER
        
objPDF.PDFSetFont FONT_ARIAL, mFontSize_, mFontFormat_
objPDF.PDFSetAlignement = mAlign_
objPDF.PDFTextOut mTexttoPrint, mTxtLeft_, mTxtTop_

End Sub

Function CountWords(Str As Variant) As Variant


'Dim abc() As String

If IsNull(Str) Then Str = ""
Dim cwRegEx As Variant
Set cwRegEx = New RegExp
cwRegEx.Pattern = "\s+"
cwRegEx.Global = True
CountWords = UBound(Split(cwRegEx.Replace(Trim(Str), " "), " ")) + 1
mStrw() = Split(cwRegEx.Replace(Trim(Str), " "), " ")


End Function

Public Function CheckMultiClient() As Boolean
CheckMultiClient = False

For iCtr = 2 To MSHFlexGrid2.Rows - 1
    If Trim(MSHFlexGrid2.TextMatrix(iCtr, 9)) = "1" Then
       CheckMultiClient = False
       Exit Function
    End If
Next

For iCtr = 2 To MSHFlexGrid2.Rows - 1
    If Trim(MSHFlexGrid2.TextMatrix(iCtr, 10)) <> "" Then
       For jCtr = 2 To MSHFlexGrid2.Rows - 1
           If Trim(MSHFlexGrid2.TextMatrix(jCtr, 10)) <> "" Then
              If iCtr <> jCtr Then
                 If Trim(MSHFlexGrid2.TextMatrix(iCtr, 10)) <> Trim(MSHFlexGrid2.TextMatrix(jCtr, 10)) Then
                    CheckMultiClient = True
                    Exit Function
                 End If
           End If
           End If
       Next
    End If
Next
CheckMultiClient = False
End Function

Public Sub Gen_Mail()
Dim TOADD As Variant
Dim TOCC, mMsg, mGridMsg As Variant
Dim strSplitString() As String
Dim EmpNo_, EmpName_ As String

mMsg = "": TOADD = "": TOCC = ""

mMsg = ""
mMsg = "Dear Sir/Madam,"
mMsg = mMsg & "<br><br>"
mMsg = mMsg & "Cash Receipt Number : " & TxtCRID.Text & " dated " & TxtCRDate & " requires your approval in StoreMan since it is pertaining to more than one client. Please verify the "
mMsg = mMsg & "<br>"
mMsg = mMsg & "correctness of this cash receipt (Transaction->Cash Receipt Transction->Issue of Cash Receipt menu item) by clicking on 'Approved by AM' button."
mMsg = mMsg & "<br><br>"
mMsg = mMsg & "If you require any further details, please get in touch with " & Gen_UserName & "."
mMsg = mMsg & "<br><br>"
mMsg = mMsg & "(Auto mail generated from StoreMan)"
mMsg = mMsg & "<br><br>"

tmp = " Select Distinct EMP1.EmployeeID as 'AMEmpID',EMP1.EmployeeNo as 'AMEmpNo',EMP1.EmployeeName as 'AMName',EMP1.EmployeeEmailID as 'AMEmailID'"
tmp = tmp & " From Mst_EmpCMPMapping EGM"
tmp = tmp & " Inner Join Mst_Employee EMP on EGM.EmployeeID = EMP.EmployeeID"
tmp = tmp & " Inner Join Mst_Employee EMP1 on EMP.DirectReportingTo = EMP1.EmployeeID"
tmp = tmp & " Inner Join Mst_CMP MC on EGM.CMPID = MC.CMPID"
tmp = tmp & " Inner Join Mst_Location ML on EGM.LocationID = ML.LocationID"
tmp = tmp & " Where EMP.Flag = 'Y' And ML.LocationID = " & mLocationID & ""

Call Tmp3Table
If TmpRs3.RecordCount > 0 Then
   TOADD = IIf(IsNull(TmpRs3!AMEmailID), "", TmpRs3!AMEmailID) '""
End If
 
TOCC = Gen_UserEmailID

mMsgSubject = "Cash Receipt awaiting your approval; cmg"

'MsgBox "TO:" & TOADD & " "
'MsgBox "CC:" & TOCC & ""
Call SendMail(TOADD, TOCC, mMsgSubject, mMsg)

End Sub
Private Sub Gen_EmailForError(TOMail As Variant, CCMail As Variant, ErrorMsg As Variant, Mode As Variant)
If ErrorMsg <> "" Then
   mMsg = ""
   mMsg = "<tr>Cash Receipt " & Val(TxtCRID) & "  And Location " & TxtPreFix & " (from Store-Man) " & Mode & "</tr>"
   mMsg = mMsg & "<tr><td> Error Description </td><td> " & ErrorMsg & " </td></tr>"
   mMsg = mMsg & "<tr><td> Login By </td><td> " & Gen_UserLoginID & ":" & Gen_UserName & "  </td></tr>"
   mMsg = mMsg & "<tr><td> Entry Date/Time </td><td> " & Now & "  </td></tr>"
    
   mMsgSubject = "Error Found in Cash Receipt"
   TOADD = TOMail
   semail = CCMail
   Call SendMail1(TOADD, semail, mMsgSubject, mMsg)
End If
End Sub
Public Function SendMail1(ToADD_ As Variant, semail_ As Variant, MsgSubject_ As Variant, Message1_ As Variant)

Dim htmlStringhead As String
Dim htmlStringtail As String
Set oSMTP = New OSSMTP.SMTPSession

'servertxt = "10.0.0.10"
servertxt = "mail.ncml.com"


htmlStringhead = "<html><body><font size= 3 face= Zurich BT><table border = 1>"
htmlStringtail = "</table></Font></body></html>"
   
'With oSMTP
'  'connection
'  .Server = servertxt
'  'message
'  .MailFrom = Gen_UserEmailID  '"mayur.d@ncml.com"
'  If ToADD_ = "" Or IsNull(ToADD_) = True Then
'     ToADD_ = "parag.r@ncml.com"
'  End If
'  .SendTo = ToADD_
'  TOCC_ = semail_
'  If ToADD_ <> TOCC_ Then
'     .CC = TOCC_
'  End If
''  .BCC = "mayur.d@ncml.com"
'  .MessageSubject = MsgSubject_
'  .MessageHTML = htmlStringhead & Message1_ & htmlStringtail
'  .TimeStamp = Now()
'  If LCase(Gen_EmailConfigFlag) = "y" Then
'     .SendEmail
'  End If
'End With

Dim objSMTP
Set objSMTP = CreateObject("MailBee.SMTP")
objSMTP.LicenseKey = "MBC900-5F470FAD67-D8E202A2E22927335F24F5A667BFD196"
objSMTP.ServerName = Gen_ReadMailServerIPINI("MailIP") ' = "mail.ncml.com" '"mail.ncml.com"
objSMTP.FromAddr = Gen_UserEmailID
If IsNull(Trim(ToADD_)) = True Or Trim(ToADD_) = "" Then
   ToADD_ = "manu.ag@ncml.com"
End If
objSMTP.ToAddr = ToADD_
TOCC_ = semail_
If ToADD_ <> TOCC_ Then
   objSMTP.CCAddr = TOCC_
End If

objSMTP.Subject = MsgSubject_ '& " (from Store-Man)" 'MsgSubject_ '"Message in HTML format"
objSMTP.BodyFormat = 1
objSMTP.BodyText = htmlStringhead & Message1_ & htmlStringtail 'htmlStringhead & MsgSubject_ & " :- " & Message1_ & htmlStringtail '"<html>This word is <b>bold</b></html>"
If LCase(Gen_EmailConfigFlag) = "y" Then
   objSMTP.Send
End If

End Function

