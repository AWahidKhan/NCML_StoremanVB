VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmTxn_PledgeCashReceipt 
   Caption         =   "Pledge Cash Receipt"
   ClientHeight    =   7470
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10215
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   7470
   ScaleWidth      =   10215
   WindowState     =   2  'Maximized
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
      Left            =   5400
      TabIndex        =   0
      Top             =   3960
      Visible         =   0   'False
      Width           =   6375
      Begin VB.TextBox TxtRejectRemark 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   1560
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Top             =   840
         Width           =   4680
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
         TabIndex        =   15
         Top             =   1440
         Width           =   1440
      End
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
         TabIndex        =   11
         Top             =   1440
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker TxtRejectDate 
         Height          =   375
         Left            =   1560
         TabIndex        =   17
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
         Format          =   107413505
         CurrentDate     =   39884
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
         TabIndex        =   19
         Top             =   930
         Width           =   1260
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
         TabIndex        =   18
         Top             =   390
         Width           =   1800
      End
   End
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
      Height          =   2295
      Left            =   5400
      TabIndex        =   20
      Top             =   3960
      Visible         =   0   'False
      Width           =   6375
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
         Left            =   3000
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   1800
         Width           =   1455
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
         Left            =   1560
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   1800
         Width           =   1440
      End
      Begin VB.TextBox TxtEntryNo 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1575
         MaxLength       =   25
         TabIndex        =   22
         Top             =   285
         Visible         =   0   'False
         Width           =   1560
      End
      Begin VB.TextBox TxtReturnRemark 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   1560
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   21
         Top             =   1200
         Width           =   4680
      End
      Begin MSComCtl2.DTPicker TxtReturnDate 
         Height          =   375
         Left            =   1560
         TabIndex        =   25
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
         Format          =   107413505
         CurrentDate     =   39884
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
         TabIndex        =   28
         Top             =   315
         Visible         =   0   'False
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
         TabIndex        =   27
         Top             =   750
         Width           =   1800
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
         TabIndex        =   26
         Top             =   1290
         Width           =   1260
      End
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   9615
      Left            =   15090
      TabIndex        =   29
      Top             =   120
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      TabIndex        =   30
      Top             =   9840
      Width           =   15135
   End
   Begin VB.Frame Frame4 
      Height          =   9735
      Left            =   0
      TabIndex        =   32
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
         Left            =   6000
         TabIndex        =   33
         Top             =   3720
         Visible         =   0   'False
         Width           =   5055
         Begin VB.TextBox TxtHologramNo 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   1440
            MaxLength       =   25
            TabIndex        =   37
            Top             =   240
            Width           =   2880
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
            TabIndex        =   36
            Top             =   840
            Width           =   1440
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
         Begin VB.CheckBox Check2 
            BackColor       =   &H008080FF&
            Caption         =   "Generate Excel"
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   1440
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   1695
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
            TabIndex        =   38
            Top             =   270
            Width           =   1260
         End
      End
      Begin VB.Frame Frame0 
         Height          =   2445
         Left            =   360
         TabIndex        =   39
         Top             =   5520
         Visible         =   0   'False
         Width           =   14520
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
            TabIndex        =   53
            Top             =   1095
            Width           =   1000
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
            TabIndex        =   52
            Top             =   1095
            Width           =   1000
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
            ItemData        =   "FrmTxn_PledgeCashReceipt.frx":0000
            Left            =   6075
            List            =   "FrmTxn_PledgeCashReceipt.frx":0010
            Sorted          =   -1  'True
            TabIndex        =   51
            Top             =   375
            Width           =   1965
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
            ItemData        =   "FrmTxn_PledgeCashReceipt.frx":002C
            Left            =   2845
            List            =   "FrmTxn_PledgeCashReceipt.frx":002E
            Sorted          =   -1  'True
            TabIndex        =   50
            Top             =   375
            Width           =   3165
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
            TabIndex        =   49
            TabStop         =   0   'False
            Top             =   390
            Width           =   1125
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
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   375
            Width           =   1110
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
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   375
            Width           =   1005
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
            ItemData        =   "FrmTxn_PledgeCashReceipt.frx":0030
            Left            =   8645
            List            =   "FrmTxn_PledgeCashReceipt.frx":003A
            Sorted          =   -1  'True
            TabIndex        =   46
            Top             =   1095
            Width           =   2805
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
            ItemData        =   "FrmTxn_PledgeCashReceipt.frx":0061
            Left            =   2965
            List            =   "FrmTxn_PledgeCashReceipt.frx":0063
            Sorted          =   -1  'True
            TabIndex        =   45
            Top             =   1095
            Width           =   3165
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
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   1095
            Width           =   2445
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
            ItemData        =   "FrmTxn_PledgeCashReceipt.frx":0065
            Left            =   12360
            List            =   "FrmTxn_PledgeCashReceipt.frx":006F
            Sorted          =   -1  'True
            TabIndex        =   43
            Top             =   375
            Width           =   1245
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
            ItemData        =   "FrmTxn_PledgeCashReceipt.frx":007C
            Left            =   13590
            List            =   "FrmTxn_PledgeCashReceipt.frx":0086
            Sorted          =   -1  'True
            TabIndex        =   42
            Top             =   375
            Width           =   1245
         End
         Begin VB.CheckBox ChkBatchNo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Check1"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   5175
            TabIndex        =   41
            Top             =   1793
            Width           =   225
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
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   1740
            Width           =   2775
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   1575
            Left            =   120
            TabIndex        =   54
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
            TabIndex        =   55
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
            Format          =   107413505
            CurrentDate     =   36494
         End
         Begin MSComCtl2.DTPicker SFrom 
            Height          =   360
            Left            =   105
            TabIndex        =   56
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
            Format          =   107413505
            CurrentDate     =   36494
         End
         Begin MSComCtl2.DTPicker STo1 
            Height          =   360
            Left            =   1550
            TabIndex        =   57
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
            Format          =   107413505
            CurrentDate     =   36494
         End
         Begin MSComCtl2.DTPicker SFrom1 
            Height          =   360
            Left            =   135
            TabIndex        =   58
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
            Format          =   107413505
            CurrentDate     =   36494
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
            TabIndex        =   75
            Top             =   120
            Width           =   855
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
            TabIndex        =   74
            Top             =   120
            Width           =   1200
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
            TabIndex        =   73
            Top             =   120
            Width           =   960
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
            TabIndex        =   72
            Top             =   120
            Width           =   1800
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
            Left            =   480
            TabIndex        =   71
            Top             =   120
            Width           =   495
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
            TabIndex        =   70
            Top             =   405
            Width           =   645
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
            TabIndex        =   69
            Top             =   435
            Width           =   255
         End
         Begin VB.Label Label19 
            Caption         =   "ID"
            Height          =   255
            Left            =   11760
            TabIndex        =   68
            Top             =   120
            Width           =   405
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
            TabIndex        =   67
            Top             =   795
            Width           =   1590
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
            TabIndex        =   66
            Top             =   795
            Width           =   1680
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
            TabIndex        =   65
            Top             =   795
            Width           =   1245
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
            TabIndex        =   64
            Top             =   795
            Width           =   1575
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
            TabIndex        =   63
            Top             =   795
            Width           =   855
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
            TabIndex        =   62
            Top             =   120
            Width           =   840
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
            TabIndex        =   61
            Top             =   120
            Width           =   840
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
            TabIndex        =   60
            Top             =   1800
            Width           =   1665
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
            TabIndex        =   59
            Top             =   1470
            Width           =   2160
         End
      End
      Begin VB.Frame ButtonFrm 
         Height          =   1095
         Left            =   135
         TabIndex        =   110
         Top             =   8445
         Width           =   14985
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
            TabIndex        =   124
            Top             =   585
            Width           =   2115
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
            Left            =   2220
            TabIndex        =   123
            Top             =   585
            Width           =   2115
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
            Left            =   4335
            TabIndex        =   122
            Top             =   585
            Width           =   2115
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
            Left            =   6450
            TabIndex        =   121
            Top             =   585
            Width           =   2115
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
            Left            =   8565
            TabIndex        =   120
            Top             =   585
            Width           =   2115
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
            TabIndex        =   119
            Top             =   165
            Width           =   2115
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
            Left            =   2220
            TabIndex        =   118
            Top             =   165
            Width           =   2115
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
            Left            =   4335
            TabIndex        =   117
            Top             =   165
            Width           =   2115
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
            Left            =   6450
            TabIndex        =   116
            Top             =   165
            Width           =   2115
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
            Left            =   8565
            TabIndex        =   115
            Top             =   165
            Width           =   2115
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
            Left            =   10680
            TabIndex        =   114
            Top             =   585
            Width           =   2115
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
            Left            =   10680
            TabIndex        =   113
            Top             =   165
            Width           =   2115
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
            Left            =   12810
            TabIndex        =   112
            Top             =   165
            Width           =   1995
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
            Left            =   12810
            TabIndex        =   111
            Top             =   585
            Width           =   1995
         End
      End
      Begin VB.Frame Frame1 
         Height          =   3480
         Left            =   120
         TabIndex        =   76
         Top             =   120
         Width           =   14895
         Begin VB.TextBox TxtExcAmount 
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
            Left            =   9840
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   153
            Top             =   3000
            Width           =   1590
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
            TabIndex        =   3
            Top             =   615
            Width           =   4860
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
            Left            =   9840
            MaxLength       =   150
            TabIndex        =   9
            Top             =   1551
            Width           =   4860
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
            Left            =   8160
            TabIndex        =   86
            Top             =   240
            Width           =   1140
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
            TabIndex        =   10
            Top             =   1980
            Width           =   12720
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
            ItemData        =   "FrmTxn_PledgeCashReceipt.frx":0093
            Left            =   2010
            List            =   "FrmTxn_PledgeCashReceipt.frx":00A3
            Sorted          =   -1  'True
            TabIndex        =   4
            Top             =   1099
            Width           =   2460
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
            TabIndex        =   5
            Top             =   1099
            Width           =   1560
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
            TabIndex        =   85
            TabStop         =   0   'False
            Top             =   660
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
            Left            =   2010
            Sorted          =   -1  'True
            TabIndex        =   2
            Top             =   647
            Width           =   5565
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
            TabIndex        =   84
            TabStop         =   0   'False
            Top             =   195
            Width           =   1440
         End
         Begin VB.TextBox TxtPCRID 
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
            TabIndex        =   6
            Top             =   1080
            Width           =   1605
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
            TabIndex        =   8
            Top             =   1551
            Width           =   5550
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
            Left            =   14760
            Locked          =   -1  'True
            TabIndex        =   82
            TabStop         =   0   'False
            Top             =   3000
            Visible         =   0   'False
            Width           =   1440
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
            TabIndex        =   80
            TabStop         =   0   'False
            Top             =   195
            Width           =   1440
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
            TabIndex        =   79
            Top             =   2910
            Width           =   5775
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
            TabIndex        =   78
            Top             =   2490
            Width           =   5565
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
            TabIndex        =   77
            TabStop         =   0   'False
            Top             =   3000
            Width           =   1440
         End
         Begin MSComCtl2.DTPicker TxtInstrumentDate 
            Height          =   375
            Left            =   13260
            TabIndex        =   7
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
            Format          =   107413505
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
            Format          =   107413505
            CurrentDate     =   39884
         End
         Begin MSComCtl2.DTPicker TxtFAGPostingDate 
            Height          =   375
            Left            =   13245
            TabIndex        =   88
            Top             =   2490
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
            Format          =   107413505
            CurrentDate     =   39884
         End
         Begin MSComCtl2.DTPicker TxtDepDate 
            Height          =   375
            Left            =   9840
            TabIndex        =   89
            Top             =   2490
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
            Format          =   107413505
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
            TabIndex        =   87
            TabStop         =   0   'False
            Top             =   647
            Width           =   5550
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
            TabIndex        =   81
            TabStop         =   0   'False
            Top             =   2490
            Visible         =   0   'False
            Width           =   5565
         End
         Begin VB.Label Label25 
            Caption         =   "Excess Amount"
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
            Left            =   8160
            TabIndex        =   154
            Top             =   3030
            Width           =   1680
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
            TabIndex        =   109
            Top             =   225
            Width           =   1200
         End
         Begin VB.Label LblStateID 
            Caption         =   "Drawn On"
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
            Left            =   225
            TabIndex        =   108
            Top             =   1581
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
            TabIndex        =   107
            Top             =   225
            Width           =   1800
         End
         Begin VB.Label Label1 
            Caption         =   "Branch Name"
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
            Left            =   8160
            TabIndex        =   106
            Top             =   1604
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
            TabIndex        =   105
            Top             =   1129
            Width           =   1800
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
            TabIndex        =   104
            Top             =   2040
            Width           =   1800
         End
         Begin VB.Label Label5 
            Caption         =   "Instrument No"
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
            Left            =   8160
            TabIndex        =   103
            Top             =   1125
            Width           =   1560
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
            TabIndex        =   102
            Top             =   1125
            Width           =   960
         End
         Begin VB.Label Label7 
            Caption         =   "Instrument Date"
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
            Left            =   11610
            TabIndex        =   101
            Top             =   1140
            Width           =   1800
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
            TabIndex        =   100
            Top             =   685
            Width           =   1440
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
            TabIndex        =   99
            Top             =   2520
            Visible         =   0   'False
            Width           =   720
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
            Left            =   8160
            TabIndex        =   98
            Top             =   675
            Width           =   1935
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
            Left            =   11610
            TabIndex        =   97
            Top             =   2445
            Width           =   1455
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
            Left            =   14160
            TabIndex        =   96
            Top             =   3000
            Visible         =   0   'False
            Width           =   1455
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
            TabIndex        =   95
            Top             =   2460
            Width           =   1575
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
            TabIndex        =   94
            Top             =   3015
            Width           =   1695
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
            TabIndex        =   93
            Top             =   240
            Width           =   975
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
            TabIndex        =   92
            Top             =   255
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
            Left            =   8160
            TabIndex        =   91
            Top             =   2490
            Width           =   1335
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
            Left            =   11610
            TabIndex        =   90
            Top             =   3060
            Width           =   1095
         End
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   4830
         Left            =   120
         TabIndex        =   125
         Top             =   3615
         Width           =   14895
         _ExtentX        =   26273
         _ExtentY        =   8520
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
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
         TabCaption(0)   =   "Pledge Details"
         TabPicture(0)   =   "FrmTxn_PledgeCashReceipt.frx":00BF
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "MSHFlexGrid2"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Frame2"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Log Details"
         TabPicture(1)   =   "FrmTxn_PledgeCashReceipt.frx":00DB
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "TxtCreated"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "TxtUpdated"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "TxtFAGUpdated"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "TxtRejectedBy"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "Label13"
         Tab(1).Control(5)=   "Label12"
         Tab(1).Control(6)=   "Label15"
         Tab(1).Control(7)=   "Label50"
         Tab(1).ControlCount=   8
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
            Left            =   -73920
            Locked          =   -1  'True
            TabIndex        =   143
            TabStop         =   0   'False
            Top             =   480
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
            Left            =   -73920
            Locked          =   -1  'True
            TabIndex        =   142
            TabStop         =   0   'False
            Top             =   1065
            Width           =   8970
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
            Left            =   -73920
            Locked          =   -1  'True
            TabIndex        =   141
            TabStop         =   0   'False
            Top             =   1725
            Width           =   8970
         End
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
            Left            =   -73905
            Locked          =   -1  'True
            TabIndex        =   140
            TabStop         =   0   'False
            Top             =   2355
            Width           =   8970
         End
         Begin VB.Frame Frame2 
            Height          =   1545
            Left            =   120
            TabIndex        =   126
            Top             =   360
            Width           =   14655
            Begin VB.TextBox TxtClientName 
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
               Left            =   4920
               Locked          =   -1  'True
               TabIndex        =   151
               TabStop         =   0   'False
               Top             =   1080
               Width           =   7455
            End
            Begin VB.TextBox TxtTotRecAmt 
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
               Left            =   6630
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   149
               TabStop         =   0   'False
               Top             =   660
               Width           =   1890
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
               Left            =   10425
               TabIndex        =   13
               Top             =   660
               Width           =   1920
            End
            Begin VB.TextBox TxtPledgeDate 
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
               Left            =   1650
               Locked          =   -1  'True
               TabIndex        =   132
               TabStop         =   0   'False
               Top             =   660
               Width           =   1650
            End
            Begin VB.TextBox txtPledgeID 
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
               Left            =   90
               TabIndex        =   12
               Top             =   660
               Width           =   1560
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
               Left            =   12840
               TabIndex        =   131
               Top             =   1080
               Width           =   1710
            End
            Begin VB.TextBox TxtBalAmt 
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
               Left            =   8520
               Locked          =   -1  'True
               TabIndex        =   130
               TabStop         =   0   'False
               Top             =   660
               Width           =   1905
            End
            Begin VB.TextBox TxtPledgeFees 
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
               Left            =   3300
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   129
               TabStop         =   0   'False
               Top             =   660
               Width           =   1680
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
               Height          =   405
               Left            =   12840
               TabIndex        =   14
               Top             =   660
               Width           =   1695
            End
            Begin VB.TextBox TxtWHRNo 
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
               Left            =   1650
               Locked          =   -1  'True
               TabIndex        =   128
               TabStop         =   0   'False
               Top             =   1095
               Width           =   1650
            End
            Begin VB.TextBox TxtDiscount 
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
               Left            =   4980
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   127
               TabStop         =   0   'False
               Top             =   660
               Width           =   1650
            End
            Begin VB.Label Label24 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
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
               Height          =   240
               Left            =   3435
               TabIndex        =   152
               Top             =   1155
               Width           =   1125
            End
            Begin VB.Label Label17 
               Alignment       =   2  'Center
               Caption         =   "Total Received Amount"
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
               Left            =   6630
               TabIndex        =   150
               Top             =   135
               Width           =   1680
            End
            Begin VB.Label Label14 
               Alignment       =   2  'Center
               Caption         =   "Received Amount against this CR"
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
               Left            =   10320
               TabIndex        =   148
               Top             =   120
               Width           =   2175
            End
            Begin VB.Label LblChequeAmt 
               Alignment       =   2  'Center
               Caption         =   "Balance Amount"
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
               Left            =   8580
               TabIndex        =   138
               Top             =   300
               Width           =   1635
            End
            Begin VB.Label LblTDSAmt 
               AutoSize        =   -1  'True
               Caption         =   "Pledge Fee"
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
               Left            =   3585
               TabIndex        =   137
               Top             =   240
               Width           =   1065
            End
            Begin VB.Label LblPINo 
               AutoSize        =   -1  'True
               Caption         =   "Pledge ID"
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
               Left            =   330
               TabIndex        =   136
               Top             =   240
               Width           =   1050
            End
            Begin VB.Label LblInvoiceAmount 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               Caption         =   "WHR No"
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
               TabIndex        =   135
               Top             =   1140
               Width           =   825
            End
            Begin VB.Label LblPIDate 
               AutoSize        =   -1  'True
               Caption         =   "Pledge Date"
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
               Left            =   1920
               TabIndex        =   134
               Top             =   240
               Width           =   1140
            End
            Begin VB.Label Label46 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               Caption         =   "Discount"
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
               Left            =   5400
               TabIndex        =   133
               Top             =   240
               Width           =   780
            End
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
            Height          =   2775
            Left            =   120
            TabIndex        =   139
            Top             =   1920
            Width           =   14670
            _ExtentX        =   25876
            _ExtentY        =   4895
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
            Left            =   -74865
            TabIndex        =   147
            Top             =   540
            Width           =   1695
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
            Left            =   -74865
            TabIndex        =   146
            Top             =   1125
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
            Left            =   -74880
            TabIndex        =   145
            Top             =   1650
            Width           =   975
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
            Left            =   -74865
            TabIndex        =   144
            Top             =   2280
            Width           =   975
         End
      End
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
      Left            =   15090
      Locked          =   -1  'True
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   1320
      Visible         =   0   'False
      Width           =   1080
   End
End
Attribute VB_Name = "FrmTxn_PledgeCashReceipt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mAgentID, mClientIDRow, mSDepoBankID As Variant
Dim mClientNAVCode As String
Dim mClientRNAVCode As String
Dim mDPNAVCode As String
Dim mExchangeTypeID As Double
Dim mLocationID As Double
Dim Edit_Flg As Variant
Dim mFAGBankID As Double
Dim mPrintFlag As Variant
Dim Wc1 As String
Dim mRType As Variant
Dim mPreMsg As Variant

Dim TOADD As Variant
Dim TOCC, mMsg, mGridMsg As Variant
Dim strSplitString() As String
Dim EmpNo_, EmpName_ As String

'Declaration for PDF report
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

Private Sub Grid_Head2()

MSHFlexGrid2.Clear
MSHFlexGrid2.Rows = 3
MSHFlexGrid2.Cols = 11
MSHFlexGrid2.FixedRows = 2
MSHFlexGrid2.WordWrap = True

MSHFlexGrid2.TextMatrix(1, 0) = "Pledge ID"
MSHFlexGrid2.TextMatrix(1, 1) = "Pledge Date"
MSHFlexGrid2.TextMatrix(1, 2) = "WHR No"
MSHFlexGrid2.TextMatrix(1, 3) = "Client Name"
MSHFlexGrid2.TextMatrix(1, 4) = "Pledge Fees"
MSHFlexGrid2.TextMatrix(1, 5) = "Discount"
MSHFlexGrid2.TextMatrix(1, 6) = "Received Amount"
MSHFlexGrid2.TextMatrix(1, 7) = "Balance Receivable"
MSHFlexGrid2.TextMatrix(1, 8) = "Received Amt against this CR"
MSHFlexGrid2.TextMatrix(1, 9) = "Print Flag"
MSHFlexGrid2.TextMatrix(1, 10) = "Pledged Bank"

MSHFlexGrid2.ColWidth(0) = 1500
MSHFlexGrid2.ColWidth(1) = 1500
MSHFlexGrid2.ColWidth(2) = 1300
MSHFlexGrid2.ColWidth(3) = 3200
MSHFlexGrid2.ColWidth(4) = 1300
MSHFlexGrid2.ColWidth(5) = 1300
MSHFlexGrid2.ColWidth(6) = 1300
MSHFlexGrid2.ColWidth(7) = 1300
MSHFlexGrid2.ColWidth(8) = 1500
MSHFlexGrid2.ColWidth(9) = 500
MSHFlexGrid2.ColWidth(10) = 0

MSHFlexGrid2.RowHeight(1) = 800

End Sub

Private Sub AddCmd_Click()

Edit_Flg = "A"
Call GButtonLock(Me, False)
Frame1.Enabled = True
Frame2.Enabled = True
TxtInstrumentAmount.Locked = False
CmdReturn.Enabled = False
MSHFlexGrid2.Enabled = True
Call Grid_Head2
TxtCRDate.Enabled = True
TxtInstrumentDate.Enabled = True
CmbInstrumentType.Visible = True
TxtG_UID = ""
CmbLocationID.Visible = True
TxtLocationID.Visible = False
TxtPCRID.Locked = True
TxtInstrumentNo.Locked = False
TxtDrawnOn.Locked = False
TxtBranchName.Locked = False
TxtRemark.Locked = False
mPrintFlag = ""

SSTab1.Tab = 0
TxtExcAmount = ""
TxtPCRID = ""
TxtInstrumentNo = ""
TxtCRDate.Value = Date
TxtInstrumentDate = Date
CmbInstrumentType.Text = ""
CmbLocationID.Text = ""
TxtReceivedFrom.Text = ""
TxtInstrumentAmount = ""
Call ClearPledgeTabControl

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

If RsTxn_PledgeCashReceipt.RecordCount > 0 Then
    TxtCRDate.SetFocus
End If

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
    CmbLocationID.AddItem RsMst_Location!LocationName
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
    Call Grid_Head2
    Call ClearPledgeTabControl
End If

mLocationID = RsMst_Location!LocationID
TxtPreFix = GetPreFixByLocation(RsMst_Location!LocationID, "P")

If ChkInstrumentNo = False Then
   If LCase(CmbInstrumentType.Text) = "cheque" Then
      MsgBox "Cheque No already entered for Cash Receipt - " & TmpRs1!PCRID & "-" & GetLocationName(TmpRs1!LocationID) & " "
   ElseIf LCase(CmbInstrumentType.Text) = "dd" Then
      MsgBox "DD No already entered for Cash Receipt - " & TmpRs1!PCRID & "-" & GetLocationName(TmpRs1!LocationID) & " "
   ElseIf LCase(CmbInstrumentType.Text) = "rtgs" Then
      MsgBox "RTGS No already entered for Cash Receipt - " & TmpRs1!PCRID & "-" & GetLocationName(TmpRs1!LocationID) & " "
   End If
   CmbLocationID.SetFocus
   Exit Sub
End If

End Sub
Private Function ChkInstrumentNo() As Boolean
Dim Retval As Boolean
Retval = True

If TxtInstrumentNo = "" Then GoTo l 'Exit Function
If TxtInstrumentAmount = "" Then GoTo l 'Exit Function
If CmbLocationID.Text = "" Then GoTo l 'Exit Function

'Commented On 03 Feb 2012 Mail By Mayur Sir and Ashok Sir for checking Duli on No, Location and Amount

'tmp = "Select * from Txn_NewCashReceipt Where InstrumentType <> 'Cash' and InstrumentNo = '" & TxtInstrumentNo & "' "
'*******'

tmp = "Select * from Txn_PledgeCashReceipt Where InstrumentType <> 'Cash' and InstrumentNo = '" & TxtInstrumentNo & "' and InstrumentAmount = " & Val(TxtInstrumentAmount) & " And LocationID = " & mLocationID & " And isnull(RejectFlag,'') <> 'Y' And isnull(ReturnFlag,'') <> 'Y' "
Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   If Edit_Flg = "A" Then
      Retval = False
   Else
      If TxtPCRID <> TmpRs1!PCRID Then 'And mLocationID <> TmpRs1!LocationID Then
         Retval = False
      End If
   End If
   
End If

l: ChkInstrumentNo = Retval
End Function
Private Sub CmdAdd2List_Click()

If TxtPledgeID = "" Then
   MsgBox "Blank 'PledgeID' not allowed!!"
   TxtPledgeID.SetFocus
   Exit Sub
End If

If TxtRecAmt = "" Then
   MsgBox "Blank 'Received Amount' not allowed!!"
   TxtRecAmt.SetFocus
   Exit Sub
End If

If (Val(TxtBalAmt.Text) - Val(TxtRecAmt)) < 0 Then
   MsgBox "Received Amount must be less than or equal to 'Balance Amount'."
   TxtRecAmt.SetFocus
   Exit Sub
End If
If Validate_DuplicateGrid = False Then
   MsgBox "Duplicate Pledge ID not allowed!!"
   TxtPledgeID.SetFocus
   Exit Sub
End If
If CmdAdd2List.Caption = "Update List" Then
   r = MSHFlexGrid2.RowSel
   MSHFlexGrid2.TextMatrix(0, 8) = Format(Val(MSHFlexGrid2.TextMatrix(0, 8)) - Val(MSHFlexGrid2.TextMatrix(r, 8)))
Else
   r = MSHFlexGrid2.Rows - 1
   MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
End If

MSHFlexGrid2.TextMatrix(r, 0) = TxtPledgeID.Text
MSHFlexGrid2.TextMatrix(r, 1) = TxtPledgeDate.Text
MSHFlexGrid2.TextMatrix(r, 2) = TxtWHRNo.Text
MSHFlexGrid2.TextMatrix(r, 3) = TxtClientName.Text
MSHFlexGrid2.TextMatrix(r, 4) = TxtPledgeFees.Text
MSHFlexGrid2.TextMatrix(r, 5) = TxtDiscount.Text
MSHFlexGrid2.TextMatrix(r, 6) = TxtTotRecAmt.Text
MSHFlexGrid2.TextMatrix(r, 7) = TxtBalAmt.Text
MSHFlexGrid2.TextMatrix(r, 8) = TxtRecAmt.Text

MSHFlexGrid2.TextMatrix(0, 8) = Val(MSHFlexGrid2.TextMatrix(0, 8)) + Val(MSHFlexGrid2.TextMatrix(r, 8))
TxtExcAmount = Val(TxtInstrumentAmount) - Val(MSHFlexGrid2.TextMatrix(0, 8))
Call ClearPledgeTabControl
TxtPledgeID.SetFocus
End Sub

Private Sub CmdCancel_Click()
TxtHologramNo = ""
Frame3.Visible = False
End Sub

Private Sub CmdCancelReject_Click()
Frame8.Visible = False
ButtonFrm.Enabled = True

End Sub

Private Sub CmdDeleteList_Click()
I = MSHFlexGrid2.RowSel
If MSHFlexGrid2.TextMatrix(r, 9) = "P" Then Exit Sub
If I > 1 Then
   If MSHFlexGrid2.TextMatrix(I, 0) <> "" Then
      MSHFlexGrid2.TextMatrix(0, 8) = Val(MSHFlexGrid2.TextMatrix(0, 8)) - Val(MSHFlexGrid2.TextMatrix(I, 8))
      For RO = I To MSHFlexGrid2.Rows - 2
          For C = 0 To MSHFlexGrid2.Cols - 1
              MSHFlexGrid2.TextMatrix(RO, C) = MSHFlexGrid2.TextMatrix(RO + 1, C)
          Next
      Next
      MSHFlexGrid2.Rows = MSHFlexGrid2.Rows - 1
      MSHFlexGrid2.row = 0
   End If
End If
TxtExcAmount = Val(TxtInstrumentAmount) - Val(MSHFlexGrid2.TextMatrix(0, 8))
Call ClearPledgeTabControl
End Sub

Private Sub CmdPrint_Click()
If RsTxn_PledgeCashReceipt!HologramNo = "" Or IsNull(RsTxn_PledgeCashReceipt!HologramNo) Then
   If LCase(UserType) = "power" Then
      Check2.Visible = True
   End If
   Frame3.Visible = True
   TxtHologramNo.SetFocus
   Call TableMst_Hologram
   Exit Sub
Else
   TxtHologramNo = ""
   Frame3.Visible = False
End If
RsTxn_PledgeCashReceipt!Flag = "P"
RsTxn_PledgeCashReceipt.Update
Gen_mTimeStamp = GetTimeStamp
''Call Create_Xls(Gen_mTimeStamp & Gen_UserName & "_PledgeCashReceipt-" & Trim(TxtPreFix) & "" & Trim(TxtPCRID) & ".xls")
'''Call Xls_Print_Rpt
''Call Xls_Print_Rpt_New
PntPrv = "N"
Call PDF_Print_Rpt

End Sub

Private Sub CmdReject_Click()

If LCase(Gen_UserRole) <> "fag" And LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "fagadmin" Then
   MsgBox "Only FAG user can reject the Cash Receipt!!!"
   Exit Sub
End If
    
If RsTxn_PledgeCashReceipt!BatchID <> 0 Then
   MsgBox "Batch is generated already. You can't reject this receipt!!!! "
   Exit Sub
End If

If IsNull(RsTxn_PledgeCashReceipt!FAGPostingDate) = False Then
   MsgBox "Receipt is approved by FAG. You can't reject this receipt!!!! "
   Exit Sub
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
TxtReturnDate.SetFocus

End Sub

Private Sub CmdSaveHologram_Click()

On Error GoTo msgError

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
      RsMst_Hologram!DocType = "Pledge Cash Receipt"
      RsMst_Hologram!DocNo = TxtPCRID
      RsMst_Hologram!SM_Prefix = TxtPreFix
      RsMst_Hologram!CreatedBy = Gen_UserLoginID
      RsMst_Hologram.Update
      RsMst_Hologram.Close
   End If
   RsTxn_PledgeCashReceipt!HologramNo = TxtHologramNo
   RsTxn_PledgeCashReceipt!Flag = "P"
   RsTxn_PledgeCashReceipt.Update
   TxtHologramNo = ""
   Frame3.Visible = False
   Gen_mTimeStamp = GetTimeStamp
   If Check2.Value = 1 Then
      PntPrv = "N"
      Call PDF_Print_Rpt
   End If
Else
   TxtHologramNo.SetFocus
End If

Exit Sub
msgError:
MsgBox Err.Description, vbExclamation, "Error" '"Some error occurred during process!!!"

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

Call DeletePledgeDetail

RsTxn_PledgeCashReceipt!RejectFlag = "Y"
RsTxn_PledgeCashReceipt!RejectDate = TxtRejectDate.Value
RsTxn_PledgeCashReceipt!RejectRemark = TxtRejectRemark.Text
RsTxn_PledgeCashReceipt!RejectedBy = Gen_UserLoginID
RsTxn_PledgeCashReceipt.Update
RsTxn_PledgeCashReceipt.Requery
RsTxn_PledgeCashReceipt.MoveFirst
RsTxn_PledgeCashReceipt.Find "G_UID= '" & TxtG_UID & "'"

If RsTxn_PledgeCashReceipt.EOF Then
   RsTxn_PledgeCashReceipt.MoveLast
End If

Frame8.Visible = False
ButtonFrm.Enabled = True

mRType = "Reject"

'Call Gen_Email
Call Indata

End Sub

Private Sub CmdSaveReturn_Click()

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
   MsgBox "Return Date allowed for next month of " & Format(Gen_BlockingDate, "MMM-yyyy") & "!!!"
   TxtReturnDate.SetFocus
   Exit Sub
End If

If TxtReturnRemark.Text = "" Then
   MsgBox "Blank Remark not allowed!!"
   TxtReturnRemark.SetFocus
   Exit Sub
End If
TxtReturnRemark = Replace(TxtReturnRemark, ",", "-")
'RsTxn_PledgeCashReceipt!EntryNo = TxtEntryNo.Text
Call DeletePledgeDetail
RsTxn_PledgeCashReceipt!ReturnDate = TxtReturnDate.Value
RsTxn_PledgeCashReceipt!ReturnFlag = "Y"
RsTxn_PledgeCashReceipt!ReturnRemark = TxtReturnRemark.Text
RsTxn_PledgeCashReceipt.Update
RsTxn_PledgeCashReceipt.Requery
RsTxn_PledgeCashReceipt.MoveFirst
RsTxn_PledgeCashReceipt.Find "G_UID= '" & TxtG_UID & "'"
If RsTxn_PledgeCashReceipt.EOF Then
   RsTxn_PledgeCashReceipt.MoveLast
End If

Frame6.Visible = False
CmdReturn.Enabled = False

mRType = "Return"

''Call Gen_Email

End Sub

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


tmp = "Select CR.PCRID,LM.LocationName,CR.PCRDate,CR.InstrumentType,CR.DrawnOn,CR.BranchName,CR.InstrumentNo,CR.InstrumentDate,"
tmp = tmp & " CR.InstrumentAmount ,CR.ReceivedFrom, CR.Remark, CR.Flag,CR.G_UID,CR.DepositDate,CR.FAGPostingDate,CR.FAGRemarks 'DepositRemarks',MFB.BankName +'~'+ MFB.AccountNo as 'Bank',CR.AdvanceFlag,CR.ReturnFlag,CR.BatchID,CR.RejectFlag,CR.RejectDate,CR.RejectRemark"
tmp = tmp & " From Txn_PledgeCashReceipt CR"
tmp = tmp & " Inner Join Mst_Location LM On CR.LocationID = LM.LocationID"
tmp = tmp & " Left Join Mst_FAGBank MFB On CR.FAGBankID = MFB.FAGBankID "
If Gen_DBType = "MDB" Then
   tmp = tmp & " Where CR.PCRDate Between #" & Format(SFrom, Gen_DatFormat) & "# And #" & Format(STo, Gen_DatFormat) & "#) "
Else
    tmp = tmp & " where CR.PCRDate Between '" & Format(SFrom, Gen_DatFormat) & "' And '" & Format(STo, Gen_DatFormat) & "' "
End If

If TxtAmountFrom <> "" And TxtAmountTo <> "" Then
   tmp = tmp & " AND CR.InstrumentAmount between " & Val(TxtAmountFrom) & " And " & Val(TxtAmountTo) & " "
End If

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
   tmp = tmp & " And CR.PCRID = " & Val(TxtSCRID)
End If

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

If ChkBatchNo.Value = 0 Then
   tmp = tmp & " AND CR.BatchID = 0 "
End If

Call TmpTable

For iu = 1 To TmpRs.RecordCount
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
    For C = 0 To MSHFlexGrid1.Cols - 1
       MSHFlexGrid1.TextMatrix(iu, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
    Next
    MSHFlexGrid1.TextMatrix(iu, 2) = Format(MSHFlexGrid1.TextMatrix(iu, 2), "dd-mmm-yyyy")
    If MSHFlexGrid1.TextMatrix(iu, 7) <> "" Then
       MSHFlexGrid1.TextMatrix(iu, 7) = Format(MSHFlexGrid1.TextMatrix(iu, 7), "dd-mmm-yyyy")
    End If
    If MSHFlexGrid1.TextMatrix(iu, 13) <> "" Then
       MSHFlexGrid1.TextMatrix(iu, 13) = Format(MSHFlexGrid1.TextMatrix(iu, 13), "dd-mmm-yyyy")
    End If
    If MSHFlexGrid1.TextMatrix(iu, 14) <> "" Then
       MSHFlexGrid1.TextMatrix(iu, 14) = Format(MSHFlexGrid1.TextMatrix(iu, 14), "dd-mmm-yyyy")
    End If
    If MSHFlexGrid1.TextMatrix(iu, 17) = "1" Then
       MSHFlexGrid1.TextMatrix(iu, 17) = "Yes"
    Else
       MSHFlexGrid1.TextMatrix(iu, 17) = "No"
    End If
    If MSHFlexGrid1.TextMatrix(iu, 18) = "Y" Then
       MSHFlexGrid1.TextMatrix(iu, 18) = "Yes"
    Else
       MSHFlexGrid1.TextMatrix(iu, 18) = "No"
    End If
    If MSHFlexGrid1.TextMatrix(iu, 20) = "Y" Then
       MSHFlexGrid1.TextMatrix(iu, 20) = "Yes"
    Else
       MSHFlexGrid1.TextMatrix(iu, 20) = "No"
    End If
    TmpRs.MoveNext
Next
End Sub

Private Sub Command1_Click()
TxtEntryNo = ""
TxtReturnDate.Value = Date
TxtReturnRemark = ""
Frame6.Visible = False
End Sub

Private Sub DeleteCmd_Click()

If RsTxn_PledgeCashReceipt!RejectFlag = "Y" Then
   MsgBox "Receipt is rejected. You can't delete this receipt!!!! "
   Exit Sub
End If

If RsTxn_PledgeCashReceipt!Flag = "P" Then
   MsgBox "Print out is generated already. You can't delete this receipt!!!! "
   Exit Sub
End If

Dim ans As Variant
On Error GoTo msgError

ans = MsgBox("Do you really want to DELETE this Record??", vbYesNo)
If ans = vbYes Then
   Call DeletePledgeDetail
   
   RsTxn_PledgeCashReceipt.Delete
   RsTxn_PledgeCashReceipt.Update
   RsTxn_PledgeCashReceipt.Requery
   If RsTxn_PledgeCashReceipt.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsTxn_PledgeCashReceipt.MoveNext
   If RsTxn_PledgeCashReceipt.EOF() Then
      RsTxn_PledgeCashReceipt.MoveLast
   End If
   Call Indata
End If
Exit Sub
msgError:
MsgBox "Child record Present "
RsTxn_PledgeCashReceipt.CancelUpdate

End Sub

Private Sub EditCmd_Click()

If RsTxn_PledgeCashReceipt!RejectFlag = "Y" Then
   MsgBox "Receipt is rejected. You can't edit this receipt!!!! "
   Exit Sub
End If

If (LCase(Gen_UserRole) = "fag" Or LCase(Gen_UserRole) = "power" Or LCase(Gen_UserRole) = "fagadmin") Then
   If RsTxn_PledgeCashReceipt!BatchID <> 0 Then
      MsgBox "Batch is generated already. You can't edit this receipt!!!! "
      Exit Sub
   End If
   Frame1.Enabled = True
Else
   If IsNull(RsTxn_PledgeCashReceipt!FAGPostingDate) = False Then
      MsgBox "Already Posted by FAG. You can't edit this receipt!!!! "
      Exit Sub
   End If
   If RsTxn_PledgeCashReceipt!Flag = "P" Then
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

Frame2.Enabled = True

CmdPrint.Enabled = False
CmdReject.Enabled = False
CmdPrintPreview.Enabled = False

TxtInstrumentAmount.Locked = False

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

Call TableTxn_PledgeCashReceipt(" " & wc & "  Order by SM_Prefix,PCRID")


If RsTxn_PledgeCashReceipt.RecordCount = 0 Then
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
Private Sub ClearPledgeTabControl()
TxtPledgeID.Text = ""
TxtBalAmt = ""
TxtClientName = ""
TxtDiscount = ""
TxtPledgeDate = ""
TxtPledgeFees = ""
TxtTotRecAmt = ""
TxtRecAmt = ""
TxtWHRNo = ""
CmdAdd2List.Caption = "Add to List"
TxtPledgeID.Locked = False
'txtPledgeID.SetFocus
End Sub

Private Sub MSHFlexGrid1_Click()
Dim I As Variant
I = MSHFlexGrid1.RowSel
If I <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(I, 0) = "" Then Exit Sub
RsTxn_PledgeCashReceipt.MoveFirst
RsTxn_PledgeCashReceipt.Find "G_UID= '" & MSHFlexGrid1.TextMatrix(I, 12) & "'"

If Not RsTxn_PledgeCashReceipt.EOF Then
   Call Indata
   Frame0.Visible = False
End If

End Sub

Private Sub MSHFlexGrid2_Click()
If SaveCmd.Enabled = False Then Exit Sub
r = MSHFlexGrid2.RowSel

If r < 2 Or r = MSHFlexGrid2.Rows - 1 Then
   Call ClearPledgeTabControl
   Exit Sub
End If
If MSHFlexGrid2.TextMatrix(r, 0) = "" Then Exit Sub
If MSHFlexGrid2.TextMatrix(r, 9) = "P" Then Exit Sub
CmdAdd2List.Caption = "Update List"

TxtPledgeID.Text = MSHFlexGrid2.TextMatrix(r, 0)
TxtPledgeDate.Text = MSHFlexGrid2.TextMatrix(r, 1)
TxtWHRNo.Text = MSHFlexGrid2.TextMatrix(r, 2)
TxtClientName.Text = MSHFlexGrid2.TextMatrix(r, 3)
TxtPledgeFees.Text = MSHFlexGrid2.TextMatrix(r, 4)
TxtDiscount.Text = MSHFlexGrid2.TextMatrix(r, 5)
TxtTotRecAmt.Text = Val(MSHFlexGrid2.TextMatrix(r, 6))
TxtBalAmt.Text = MSHFlexGrid2.TextMatrix(r, 7)
TxtRecAmt.Text = MSHFlexGrid2.TextMatrix(r, 8)
TxtPledgeID.Locked = True
TxtRecAmt.SetFocus

End Sub

Private Sub SaveCmd_Click()

''If Edit_Flg = "A" Then
''   If Month(TxtCRDate) < 6 And Year(TxtCRDate) <= 2011 Then
''      MsgBox " Cash Receipt earlier to June 2011 is not allowed  !!!"
''      TxtCRDate.SetFocus
''      Exit Sub
''   End If
''
''End If

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

If TxtDepDate.Visible = True Then
    If TxtCRDate.Value > TxtDepDate.Value Then
       MsgBox "Cash Receipt Date should not greater than Deposite Date!!!"
       TxtDepDate.SetFocus
       Exit Sub
    End If
    If Format(Gen_BlockingDate, "yyyy-mm") >= Format(TxtDepDate.Value, "yyyy-mm") Then
       MsgBox "Deposit Date allowed for next month of " & Format(Gen_BlockingDate, "MMM-yyyy") & "!!! "
       TxtDepDate.SetFocus
       Exit Sub
    End If
End If

If TxtFAGPostingDate.Visible = True Then
   If TxtDepDate.Value > TxtFAGPostingDate.Value Then
      MsgBox "Deposite Date should not greater than FAG Posting Date!!!"
      TxtFAGPostingDate.SetFocus
      Exit Sub
   End If
   If Format(Gen_BlockingDate, "yyyy-mm") >= Format(TxtFAGPostingDate.Value, "yyyy-mm") Then
       MsgBox "FAG Posting Date allowed for next month of " & Format(Gen_BlockingDate, "MMM-yyyy") & "!!! "
       TxtFAGPostingDate.SetFocus
       Exit Sub
    End If
End If


If TxtRemark.Text = "" Then
   MsgBox "Blank Remark not allowed !!! "
   TxtRemark.SetFocus
   Exit Sub
End If

If MSHFlexGrid2.Rows < 4 Then
   MsgBox "Blank Pledge Details not allowed !!! "
   SSTab1.Tab = 0
   TxtPledgeID.SetFocus
   Exit Sub
End If

If Val(TxtInstrumentAmount) < Val(MSHFlexGrid2.TextMatrix(0, 8)) Then
   MsgBox "'Total Received Amount against this CR' must be less than or equal to Instrument Amount!!!"
   Exit Sub
End If

If Edit_Flg = "A" Then
   RsTxn_PledgeCashReceipt.AddNew
   RsTxn_PledgeCashReceipt!PCRID = GetPCRID
   RsTxn_PledgeCashReceipt!G_UID = GetGUID
   TxtPCRID.Text = RsTxn_PledgeCashReceipt!PCRID
   TxtG_UID = RsTxn_PledgeCashReceipt!G_UID
Else
   RsTxn_PledgeCashReceipt.MoveFirst
   RsTxn_PledgeCashReceipt.Filter = "G_UID= '" & TxtG_UID & "'"
End If

RsTxn_PledgeCashReceipt!PCRID = TxtPCRID
RsTxn_PledgeCashReceipt!PCRDate = TxtCRDate.Value
RsTxn_PledgeCashReceipt!InstrumentType = CmbInstrumentType.Text
RsTxn_PledgeCashReceipt!DrawnOn = TxtDrawnOn
RsTxn_PledgeCashReceipt!BranchName = TxtBranchName
RsTxn_PledgeCashReceipt!InstrumentNo = TxtInstrumentNo
RsTxn_PledgeCashReceipt!InstrumentDate = TxtInstrumentDate.Value
RsTxn_PledgeCashReceipt!InstrumentAmount = Val(TxtInstrumentAmount)
RsTxn_PledgeCashReceipt!ExcessAmount = Val(TxtExcAmount)
RsTxn_PledgeCashReceipt!Remark = TxtRemark
If Val(MSHFlexGrid2.TextMatrix(0, 8)) < Val(TxtInstrumentAmount) Then
   RsTxn_PledgeCashReceipt!AdvanceFlag = 1
Else
   RsTxn_PledgeCashReceipt!AdvanceFlag = 0
End If
RsTxn_PledgeCashReceipt!LocationID = mLocationID
RsTxn_PledgeCashReceipt!SM_Prefix = TxtPreFix
RsTxn_PledgeCashReceipt!ReceivedFrom = TxtReceivedFrom
RsTxn_PledgeCashReceipt!BatchID = 0
RsTxn_PledgeCashReceipt!RBatchID = 0
RsTxn_PledgeCashReceipt!ReturnFlag = "N"
RsTxn_PledgeCashReceipt!RejectFlag = "N"

If LCase(Gen_UserRole) = "fag" Or LCase(Gen_UserRole) = "power" Or LCase(Gen_UserRole) = "fagadmin" Then
   If CmbBankName.Text <> "" Then
      If TxtFAGPostingDate.Visible = True Then
         RsTxn_PledgeCashReceipt!FAGPostingDate = TxtFAGPostingDate
      End If
      RsTxn_PledgeCashReceipt!FAGBankID = mFAGBankID 'CmbBankName.Visible = True
      RsTxn_PledgeCashReceipt!FAGRemarks = TxtFAGRemarks
      RsTxn_PledgeCashReceipt!DepositDate = TxtDepDate
      If mPrintFlag = "" Then
         RsTxn_PledgeCashReceipt!Flag = "P"
      End If
      RsTxn_PledgeCashReceipt!FAGUpdatedBy = Gen_UserLoginID
      RsTxn_PledgeCashReceipt!FAGUpdatedDate = Now
   End If
Else
   RsTxn_PledgeCashReceipt!FAGPostingDate = Null
   RsTxn_PledgeCashReceipt!FAGBankID = Null 'CmbBankName.Visible = True
   RsTxn_PledgeCashReceipt!FAGRemarks = Null
End If
  
If IsNull(RsTxn_PledgeCashReceipt!CreatedBy) = True Or RsTxn_PledgeCashReceipt!CreatedBy = "" Then
   RsTxn_PledgeCashReceipt!CreatedBy = Gen_UserLoginID
   RsTxn_PledgeCashReceipt!CreatedDate = Now
Else
   RsTxn_PledgeCashReceipt!UpdatedBy = Gen_UserLoginID
   RsTxn_PledgeCashReceipt!UpdatedDate = Now
End If

RsTxn_PledgeCashReceipt.Update

Call SavePledgeDetail

RsTxn_PledgeCashReceipt.Filter = ""
RsTxn_PledgeCashReceipt.Requery
RsTxn_PledgeCashReceipt.MoveFirst
RsTxn_PledgeCashReceipt.Find "G_UID= '" & TxtG_UID & "'"
If RsTxn_PledgeCashReceipt.EOF Then
   RsTxn_PledgeCashReceipt.MoveLast
End If

Call Indata

End Sub
Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   If RsTxn_PledgeCashReceipt.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   RsTxn_PledgeCashReceipt.MoveFirst
   RsTxn_PledgeCashReceipt.Find "G_UID= '" & TxtG_UID & "'"
   If RsTxn_PledgeCashReceipt.EOF Then
      RsTxn_PledgeCashReceipt.MoveLast
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
RsTxn_PledgeCashReceipt.Filter = ""
RsTxn_PledgeCashReceipt.MoveFirst

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
''===========comment on 7 july 2012 for not clearing search grid=============

End Sub

Private Sub TxtExcAmount_LostFocus()
'If TxtExcAmount.Text = "" Then Exit Sub
'If IsNumeric(TxtExcAmount.Text) = False Or Val(TxtExcAmount.Text) < 0 Then
'   MsgBox "Invalid Numeric Format !!!"
'   TxtExcAmount.SetFocus
'   Exit Sub
'End If
'TxtExcAmount = Round(Val(TxtExcAmount), 2)

End Sub
Private Sub TxtHologramNo_LostFocus()

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

Private Sub TxtInstrumentAmount_LostFocus()
If TxtInstrumentAmount.Text = "" Then Exit Sub
If IsNumeric(TxtInstrumentAmount.Text) = False Or Val(TxtInstrumentAmount.Text) < 0 Then
   MsgBox "Invalid Numeric Format !!!"
   TxtInstrumentAmount.SetFocus
   Exit Sub
End If
TxtInstrumentAmount = Round(Val(TxtInstrumentAmount), 2)
TxtExcAmount = Val(TxtInstrumentAmount) - Val(MSHFlexGrid2.TextMatrix(0, 8))
If ChkInstrumentNo = False Then
   If LCase(CmbInstrumentType.Text) = "cheque" Then
      MsgBox "Cheque No already entered for Cash Receipt - " & TmpRs1!PCRID & "-" & GetLocationName(TmpRs1!LocationID) & " "
   ElseIf LCase(CmbInstrumentType.Text) = "dd" Then
      MsgBox "DD No already entered for Cash Receipt - " & TmpRs1!PCRID & "-" & GetLocationName(TmpRs1!LocationID) & " "
   ElseIf LCase(CmbInstrumentType.Text) = "rtgs" Then
      MsgBox "RTGS No already entered for Cash Receipt - " & TmpRs1!PCRID & "-" & GetLocationName(TmpRs1!LocationID) & " "
   End If
   TxtInstrumentAmount.SetFocus
   TxtInstrumentNo = ""
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

TxtInstrumentNo = Replace(TxtInstrumentNo, Chr(32), "")

If ChkInstrumentNo = False Then
   If LCase(CmbInstrumentType.Text) = "cheque" Then
      MsgBox "Cheque No already entered for Cash Receipt - " & TmpRs1!PCRID & "-" & GetLocationName(TmpRs1!LocationID) & " "
   ElseIf LCase(CmbInstrumentType.Text) = "dd" Then
      MsgBox "DD No already entered for Cash Receipt - " & TmpRs1!PCRID & "-" & GetLocationName(TmpRs1!LocationID) & " "
   ElseIf LCase(CmbInstrumentType.Text) = "rtgs" Then
      MsgBox "RTGS No already entered for Cash Receipt - " & TmpRs1!PCRID & "-" & GetLocationName(TmpRs1!LocationID) & " "
   End If
   TxtInstrumentNo.SetFocus
   Exit Sub
End If

End Sub

Private Sub txtPledgeID_LostFocus()
If TxtPledgeID.Locked = True Then Exit Sub
   
If CmbLocationID.Text = "" Then
   MsgBox "Select Location!!"
   CmbLocationID.SetFocus
   Exit Sub
End If

If TxtPledgeID.Text = "" Then Exit Sub

'tmp = "Select TW.PledgeID,TW.WHRNo,PledgeDate,PledgeFees,isnull(Discount,0)Discount,isnull(ActualFees,0)ActualFees,isnull(RecFees,0)RecFees,MC.ClientName,TW.CMPID from Txn_PledgeCreation TP"
'tmp = tmp & " left Join Txn_WHR_Detail TW on TW.WHRNo=TP.WHRNo and TW.SM_Prefix=Tp.SM_Prefix"
'tmp = tmp & " left Join Mst_Client MC on MC.ClientIDRow=TP.ClientIDRow"
'tmp = tmp & " Where TW.PledgeID=" & Val(txtPledgeID) & " And TW.SM_Prefix='" & TxtPreFix & "'" ' And isnull(TP.Flag,'')='P'"

tmp = "Select TW.PledgeID,TW.WHRNo,PledgeDate,PledgeFees,isnull(Discount,0)Discount,isnull(ActualFees,0)ActualFees,isnull(RecFees,0)RecFees,MC.ClientName,TW.CMPID from Txn_PledgeCreation TP"
tmp = tmp & " left Join Txn_WHR_Detail TW on TW.WHRNo=TP.WHRNo and TW.SM_Prefix=Tp.SM_Prefix"
tmp = tmp & " left Join Mst_Client MC on MC.ClientIDRow=TP.ClientIDRow"
tmp = tmp & " Where TP.PledgeID=" & Val(TxtPledgeID) & " And TP.SM_Prefix='" & TxtPreFix & "'" ' And isnull(TP.Flag,'')='P'"


Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   TxtPledgeDate.Text = Format(TmpRs1!PledgeDate, "dd-MMM-yyyy")
   TxtPledgeFees.Text = TmpRs1!PledgeFees
   TxtDiscount.Text = TmpRs1!Discount
   TxtTotRecAmt.Text = TmpRs1!RecFees
   TxtBalAmt.Text = (TmpRs1!PledgeFees - TmpRs1!Discount) - TmpRs1!RecFees
   mWHRNo = TmpRs1!WHRNo
   TxtWHRNo = IIf(IsNull(TmpRs1!CMPID), 0, TmpRs1!CMPID) & "-" & Val(mWHRNo) - (IIf(IsNull(TmpRs1!CMPID), 0, TmpRs1!CMPID) * 1000000)
  
   TxtClientName.Text = TmpRs1!ClientName
   If Val(TxtBalAmt) <= 0 Then
      MsgBox " Balance Amount for this Pledge ID is zero!!"
      TxtPledgeID.SetFocus
      Exit Sub
   End If
Else
   MsgBox " Pledge ID not found!!"
   Call ClearPledgeTabControl
   TxtPledgeID.SetFocus
   Exit Sub
End If

End Sub

Private Sub SavePledgeDetail()
Call DeletePledgeDetail
Call TableTxn_PledgeCashReceiptDetail(" Where SM_Prefix='" & TxtPreFix & "'")
Call TableTxn_PledgeCreation(" Where SM_Prefix='" & TxtPreFix & "'")
For I = 2 To MSHFlexGrid2.Rows - 1
    If MSHFlexGrid2.TextMatrix(I, 0) <> "" And Val(MSHFlexGrid2.TextMatrix(I, 8)) > 0 Then
       RsTxn_PledgeCashReceiptDetail.AddNew
       RsTxn_PledgeCashReceiptDetail!PCRDID = getPCRDetailMaxID
       RsTxn_PledgeCashReceiptDetail!PCRID = Val(TxtPCRID)
       RsTxn_PledgeCashReceiptDetail!SM_Prefix = TxtPreFix
       RsTxn_PledgeCashReceiptDetail!G_UID = GetGUID
       RsTxn_PledgeCashReceiptDetail!PledgeID = Val(MSHFlexGrid2.TextMatrix(I, 0))
       RsTxn_PledgeCashReceiptDetail!RecAmount = Val(MSHFlexGrid2.TextMatrix(I, 8))
       
       RsTxn_PledgeCashReceiptDetail.Update
       RsTxn_PledgeCashReceiptDetail.Requery
       If RsTxn_PledgeCreation.RecordCount > 0 Then RsTxn_PledgeCreation.MoveFirst
       RsTxn_PledgeCreation.Find "PledgeID=" & Val(MSHFlexGrid2.TextMatrix(I, 0))
       If RsTxn_PledgeCreation.EOF = False Then
          RsTxn_PledgeCreation!RecFees = IIf(IsNull(RsTxn_PledgeCreation!RecFees), 0, RsTxn_PledgeCreation!RecFees) + Val(MSHFlexGrid2.TextMatrix(I, 8))
          RsTxn_PledgeCreation.Update
       End If
    End If
Next
End Sub

Function getPCRDetailMaxID() As Double
Dim Retval As Double
Retval = 0
tmp = "select max(PCRDID) from Txn_PledgeCashReceiptDetail "
Call TmpTable
If TmpRs.RecordCount > 0 Then
   Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
End If
Retval = Retval + 1
TmpRs.Close
getPCRDetailMaxID = Retval

End Function

Private Sub DeletePledgeDetail()

tmp = "Select PledgeID,RecAmount From Txn_PledgeCashReceiptDetail Where PCRID= " & Val(TxtPCRID) & " and SM_Prefix='" & TxtPreFix & "'"
Call TmpTable

Call TableTxn_PledgeCreation(" where PledgeID in (Select PledgeID From Txn_PledgeCashReceiptDetail Where PCRID= " & Val(TxtPCRID) & " and SM_Prefix='" & TxtPreFix & "') and SM_Prefix='" & TxtPreFix & "'")

For I = 1 To TmpRs.RecordCount
    If RsTxn_PledgeCreation.RecordCount > 0 Then RsTxn_PledgeCreation.MoveFirst
    RsTxn_PledgeCreation.Find "PledgeID=" & TmpRs!PledgeID
    If RsTxn_PledgeCreation.EOF = False Then
       RsTxn_PledgeCreation!RecFees = IIf(IsNull(RsTxn_PledgeCreation!RecFees), 0, RsTxn_PledgeCreation!RecFees) - IIf(IsNull(TmpRs!RecAmount), 0, TmpRs!RecAmount)
       RsTxn_PledgeCreation.Update
    End If
    TmpRs.MoveNext
Next

tmp = "Delete From Txn_PledgeCashReceiptDetail Where PCRID=" & Val(TxtPCRID) & "  And SM_Prefix='" & TxtPreFix & "'"
Call Tmp4Table

End Sub

Private Function GetPCRID() As Double
Dim Retval As Double

tmp = "Select max(PCRID) from Txn_PledgeCashReceipt Where SM_Prefix='" & TxtPreFix & "'"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetPCRID = Retval
End Function


Public Sub Indata()

Frame1.Enabled = False

Frame2.Enabled = False
ButtonFrm.Enabled = True
Call ClearPledgeTabControl
CmdReturn.Enabled = False
CmbLocationID.Visible = False
CmdLocationID.Visible = False
TxtLocationID.Visible = True
CmbLocationID.Text = ""

TxtPCRID = RsTxn_PledgeCashReceipt!PCRID
TxtCRDate = RsTxn_PledgeCashReceipt!PCRDate
CmbInstrumentType = RsTxn_PledgeCashReceipt!InstrumentType
TxtInstrumentNo.Enabled = True
TxtInstrumentDate.Enabled = True
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
   TxtInstrumentNo.Enabled = False
   TxtInstrumentDate.Enabled = False
End If

TxtDrawnOn = RsTxn_PledgeCashReceipt!DrawnOn
TxtBranchName = RsTxn_PledgeCashReceipt!BranchName
TxtInstrumentNo = RsTxn_PledgeCashReceipt!InstrumentNo
TxtInstrumentDate = RsTxn_PledgeCashReceipt!InstrumentDate
ChkAdv.Value = IIf(IsNull(RsTxn_PledgeCashReceipt!AdvanceFlag), 0, RsTxn_PledgeCashReceipt!AdvanceFlag)
TxtInstrumentAmount = RsTxn_PledgeCashReceipt!InstrumentAmount
TxtReceivedFrom = RsTxn_PledgeCashReceipt!ReceivedFrom
TxtRemark = RsTxn_PledgeCashReceipt!Remark
TxtPreFix = RsTxn_PledgeCashReceipt!SM_Prefix
mLocationID = IIf(IsNull(RsTxn_PledgeCashReceipt!LocationID), 0, RsTxn_PledgeCashReceipt!LocationID)
TxtLocationID = GetLocationName(mLocationID)
TxtG_UID = RsTxn_PledgeCashReceipt!G_UID
LblDepoDate.Visible = False
LblPostDate.Visible = False
TxtFAGPostingDate.Visible = False
TxtDepDate.Visible = False
mPrintFlag = IIf(IsNull(RsTxn_PledgeCashReceipt!Flag), "", RsTxn_PledgeCashReceipt!Flag)
mFAGBankID = IIf(IsNull(RsTxn_PledgeCashReceipt!FAGBankID), 0, RsTxn_PledgeCashReceipt!FAGBankID)
TxtExcAmount = IIf(IsNull(RsTxn_PledgeCashReceipt!ExcessAmount), 0, RsTxn_PledgeCashReceipt!ExcessAmount)
TxtBankName.Visible = True
CmbBankName.Visible = False
TxtFAGRemarks.Locked = True
If mFAGBankID <> 0 Then
   TxtBankName = GetFAGBankName(mFAGBankID, "N") & "~" & GetFAGBankName(mFAGBankID, "AC")
   LblDepoDate.Visible = True
   TxtDepDate.Visible = True
   If IsNull(RsTxn_PledgeCashReceipt!DepositDate) = False Then
      TxtDepDate.Value = RsTxn_PledgeCashReceipt!DepositDate
'   Else
'      TxtDepDate.Value = Date 'RsTxn_PledgeCashReceipt!DepositDate
   End If
   If IsNull(RsTxn_PledgeCashReceipt!FAGPostingDate) = False Then
      LblPostDate.Visible = True
      TxtFAGPostingDate.Visible = True
      TxtFAGPostingDate = RsTxn_PledgeCashReceipt!FAGPostingDate
      If LCase(Gen_UserRole) = "fag" Or LCase(Gen_UserRole) = "power" Or LCase(Gen_UserRole) = "fagadmin" Then
         If RsTxn_PledgeCashReceipt!ReturnFlag = "N" And (LCase(CmbInstrumentType.Text) = "cheque" Or LCase(CmbInstrumentType.Text) = "rtgs") Then
            CmdReturn.Enabled = True
         End If
      End If
   End If
      TxtFAGRemarks = RsTxn_PledgeCashReceipt!FAGRemarks
Else
   TxtBankName = "" 'GetFAGBankName(mFAGBankID, "N") & "~" & GetFAGBankName(mFAGBankID, "AC")
   TxtFAGRemarks = ""
End If


If IsNull(RsTxn_PledgeCashReceipt!SightedDate) = False Then
   TxtSightedDate.Text = Format(RsTxn_PledgeCashReceipt!SightedDate, "dd-MMM-yyyy")
Else
   TxtSightedDate.Text = ""
End If
If RsTxn_PledgeCashReceipt!ReturnFlag = "Y" Then
   TxtReturnFlag = "Yes"
Else
   TxtReturnFlag = "No"
End If
If RsTxn_PledgeCashReceipt!RejectFlag = "Y" Then
   TxtRejectFlag = "Yes"
   CmdReject.Enabled = False
Else
   TxtRejectFlag = "No"
   CmdReject.Enabled = True
End If

TxtRejectedBy = IIf(IsNull(RsTxn_PledgeCashReceipt!RejectedBy), "", RsTxn_PledgeCashReceipt!RejectedBy) & " :- " & IIf(IsNull(RsTxn_PledgeCashReceipt!RejectDate), "", Format(RsTxn_PledgeCashReceipt!RejectDate, "dd-mmm-yyyy hh:mm:ss"))
TxtCreated = IIf(IsNull(RsTxn_PledgeCashReceipt!CreatedBy), "", RsTxn_PledgeCashReceipt!CreatedBy) & " :- " & IIf(IsNull(RsTxn_PledgeCashReceipt!CreatedDate), "", Format(RsTxn_PledgeCashReceipt!CreatedDate, "dd-mmm-yyyy hh:mm:ss"))
TxtUpdated = IIf(IsNull(RsTxn_PledgeCashReceipt!UpdatedBy), "", RsTxn_PledgeCashReceipt!UpdatedBy) & " :- " & IIf(IsNull(RsTxn_PledgeCashReceipt!UpdatedDate), "", Format(RsTxn_PledgeCashReceipt!UpdatedDate, "dd-mmm-yyyy hh:mm:ss"))
SSTab1.Tab = 0
TxtFAGUpdated = IIf(IsNull(RsTxn_PledgeCashReceipt!FAGUpdatedBy), "", RsTxn_PledgeCashReceipt!FAGUpdatedBy) & " :- " & IIf(IsNull(RsTxn_PledgeCashReceipt!FAGUpdatedDate), "", Format(RsTxn_PledgeCashReceipt!FAGUpdatedDate, "dd-mmm-yyyy hh:mm:ss"))
TxtG_UID = RsTxn_PledgeCashReceipt!G_UID

Call InDataPledgeDetails

Call GButtonLock(Me, True)

CmdPrint.Enabled = True
CmdPrintPreview.Enabled = True
'CmdReject.Enabled = True
End Sub

Private Sub InDataPledgeDetails()
Dim ans As Variant

Call Grid_Head2

tmp = "Select TP.PledgeID,TW.WHRNo,TP.PledgeDate,TP.PledgeFees,isnull(TP.Discount,0)Discount,isnull(TP.ActualFees,0)ActualFees,isnull(TP.RecFees,0)RecFees,MC.ClientName,TW.CMPID,TPC.RecAmount,TP.Flag,Bank.BankName + '~' + Branch.BranchName as BankName "
tmp = tmp & " from Txn_PledgeCashReceiptDetail TPC"
tmp = tmp & " inner Join Txn_PledgeCreation TP on TPC.PledgeID=TP.PledgeID and TPC.SM_Prefix=Tp.SM_Prefix"
tmp = tmp & " Inner Join Mst_Bank Bank On TP.BankID = Bank.BankID "
tmp = tmp & " Inner Join Mst_Branch Branch On TP.BranchID = Branch.BranchID "
tmp = tmp & " Left Join Txn_WHR_Detail TW on TW.WHRNo=TP.WHRNo and TW.SM_Prefix=TP.SM_Prefix"
tmp = tmp & " Left Join Mst_Client MC on MC.ClientIDRow=TP.ClientIDRow"
tmp = tmp & " Where TPC.PCRID=" & Val(TxtPCRID) & " And TPC.SM_Prefix='" & TxtPreFix & "' "

Call Tmp4Table

MSHFlexGrid2.Rows = TmpRs4.RecordCount + 3
For I = 2 To MSHFlexGrid2.Rows - 2
    MSHFlexGrid2.TextMatrix(I, 0) = TmpRs4!PledgeID
    MSHFlexGrid2.TextMatrix(I, 1) = Format(TmpRs4!PledgeDate, "dd-MMM-yyyy")
    mWHRNo = TmpRs4!WHRNo
    MSHFlexGrid2.TextMatrix(I, 2) = IIf(IsNull(TmpRs4!CMPID), 0, TmpRs4!CMPID) & "-" & Val(mWHRNo) - (IIf(IsNull(TmpRs4!CMPID), 0, TmpRs4!CMPID) * 1000000)
    MSHFlexGrid2.TextMatrix(I, 3) = IIf(IsNull(TmpRs4!ClientName), "", TmpRs4!ClientName)
    MSHFlexGrid2.TextMatrix(I, 4) = IIf(IsNull(TmpRs4!PledgeFees), 0, TmpRs4!PledgeFees)
    MSHFlexGrid2.TextMatrix(I, 5) = IIf(IsNull(TmpRs4!Discount), 0, TmpRs4!Discount)
    MSHFlexGrid2.TextMatrix(I, 6) = IIf(IsNull(TmpRs4!RecFees), 0, TmpRs4!RecFees) - IIf(IsNull(TmpRs4!RecAmount), "", TmpRs4!RecAmount)
    MSHFlexGrid2.TextMatrix(I, 7) = IIf(IsNull(TmpRs4!PledgeFees), 0, TmpRs4!PledgeFees) - IIf(IsNull(TmpRs4!Discount), 0, TmpRs4!Discount) - Val(MSHFlexGrid2.TextMatrix(I, 6))
    MSHFlexGrid2.TextMatrix(I, 8) = IIf(IsNull(TmpRs4!RecAmount), "", TmpRs4!RecAmount)
    MSHFlexGrid2.TextMatrix(I, 9) = IIf(IsNull(TmpRs4!Flag), "", TmpRs4!Flag)
    MSHFlexGrid2.TextMatrix(I, 10) = IIf(IsNull(TmpRs4!BankName), "", TmpRs4!BankName)
    MSHFlexGrid2.TextMatrix(0, 8) = Val(MSHFlexGrid2.TextMatrix(0, 8)) + Val(MSHFlexGrid2.TextMatrix(I, 8))
    
    TmpRs4.MoveNext
Next

End Sub


Private Function Validate_DuplicateGrid() As Boolean
Dim Retval As Boolean
Retval = True
If MSHFlexGrid2.Rows = 2 Then
   Retval = True
Else
    For I = 1 To MSHFlexGrid2.Rows - 1
        If MSHFlexGrid2.TextMatrix(I, 0) <> "" Then
           If CmdAdd2List.Caption = "Update List" Then
              If I <> MSHFlexGrid2.RowSel Then
                 If Val(MSHFlexGrid2.TextMatrix(I, 0)) = TxtPledgeID Then
                    Retval = False
                 End If
              End If
           Else
              If Val(MSHFlexGrid2.TextMatrix(I, 0)) = TxtPledgeID Then
                 Retval = False
              End If
           End If
        End If
    Next
End If
Validate_DuplicateGrid = Retval
End Function


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
   TxtDepDate.Value = Date
   TxtDepDate.Visible = True
   TxtDepDate.Enabled = True
   LblDepoDate.Visible = True
End If

TxtFAGPostingDate.Value = Date

TxtFAGRemarks.Locked = False
TxtFAGRemarks = ""

End Sub


Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsTxn_PledgeCashReceipt.MoveNext
If RsTxn_PledgeCashReceipt.EOF Then
   RsTxn_PledgeCashReceipt.MoveLast
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
RsTxn_PledgeCashReceipt.MovePrevious
If RsTxn_PledgeCashReceipt.BOF Then
   RsTxn_PledgeCashReceipt.MoveFirst
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
RsTxn_PledgeCashReceipt.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsTxn_PledgeCashReceipt.MoveLast
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
MSHFlexGrid1.Cols = 23
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
MSHFlexGrid1.TextMatrix(0, 12) = "G_UID"
MSHFlexGrid1.TextMatrix(0, 13) = "Deposit Date"
MSHFlexGrid1.TextMatrix(0, 14) = "FAG Posting Date"
MSHFlexGrid1.TextMatrix(0, 15) = "FAG Remarks"
MSHFlexGrid1.TextMatrix(0, 16) = "Bank"
MSHFlexGrid1.TextMatrix(0, 17) = "Advance Flag"
MSHFlexGrid1.TextMatrix(0, 18) = "Return Flag"
MSHFlexGrid1.TextMatrix(0, 19) = "Batch ID"
MSHFlexGrid1.TextMatrix(0, 20) = "Reject Flag"
MSHFlexGrid1.TextMatrix(0, 21) = "Reject Date"
MSHFlexGrid1.TextMatrix(0, 22) = "Reject Remarks"

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
MSHFlexGrid1.ColWidth(12) = 0
MSHFlexGrid1.ColWidth(13) = 1400
MSHFlexGrid1.ColWidth(14) = 1400
MSHFlexGrid1.ColWidth(15) = 1400
MSHFlexGrid1.ColWidth(16) = 1400
MSHFlexGrid1.ColWidth(17) = 1000
MSHFlexGrid1.ColWidth(18) = 1000
MSHFlexGrid1.ColWidth(19) = 1000
MSHFlexGrid1.ColWidth(20) = 1000

MSHFlexGrid1.WordWrap = True
MSHFlexGrid1.RowHeight(0) = 600

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
    CmbSLocation.AddItem RsMst_Location!LocationName
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

Private Sub TxtRecAmt_LostFocus()
If TxtRecAmt.Text = "" Then Exit Sub
If IsNumeric(TxtRecAmt.Text) = False Or Val(TxtRecAmt.Text) < 0 Then
   MsgBox "Invalid Numeric Format !!!"
   TxtRecAmt.SetFocus
   Exit Sub
End If
If Val(TxtInstrumentAmount) < Val(TxtRecAmt) Then
   MsgBox "'Received Amount against this CR' must be less than or equal to Instrument Amount!!!"
   TxtRecAmt.SetFocus
   Exit Sub
End If

TxtRecAmt = Round(Val(TxtRecAmt), 2)

End Sub

Private Sub CmdPrintPreview_Click()
Gen_mTimeStamp = GetTimeStamp
''Call Create_Xls(Gen_mTimeStamp & Gen_UserName & "_PledgeCashReceipt-" & Trim(TxtPreFix) & "" & Trim(TxtPCRID) & ".xls")
''
'''Call Xls_Print_Rpt
''
''Call Xls_Print_Rpt_New
PntPrv = "Y"
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
MsgBox ("Wait till next message")
Set oXL = New Excel.Application
'Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "_PledgeCashReceipt-" & Trim(TxtPreFix) & "" & Trim(TxtPCRID) & ".xls")

I = 2

Set oWS = oWB.Worksheets("Sheet" & 1)
''oWS.PageSetup.RightMargin = 0.05
'oWS.PageSetup.TopMargin = 0.05
'oWS.PageSetup.RightMargin = 0.05
'oWS.PageSetup.TopMargin = 4.6
oWS.Name = TxtPreFix & "-" & TxtPCRID  'MSHFlexGrid2.TextMatrix(i, 23)
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

tmp = "a" & mRow & ":" & "h" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "NATIONAL COLLATERAL MANAGEMENT SERVICES LIMITED"
oWS.Cells(mRow, C + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 10

'oWS.Cells(1, 1).Font.ColorIndex = 4
mRow = mRow + 1

tmp = "a" & mRow & ":" & "h" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "Location -" & TxtLocationID
oWS.Cells(mRow, C + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 10

mRow = mRow + 1

tmp = "a" & mRow & ":" & "h" & mRow
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
oWS.Cells(mRow, C + 2) = TxtPreFix & "-" & TxtPCRID
oWS.Cells(mRow, C + 2).Font.Size = 7
oWS.Cells(mRow, C + 2).HorizontalAlignment = -4108

oWS.Cells(mRow, C + 4) = "Receipt Date"
oWS.Cells(mRow, C + 4).Font.Size = 7
oWS.Cells(mRow, C + 4).Font.Bold = True

tmp = "e" & mRow & ":" & "h" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 5) = " " & Format(TxtCRDate, "dd-mmm-yyyy")
oWS.Cells(mRow, C + 5).Font.Size = 7
oWS.Cells(mRow, C + 5).HorizontalAlignment = -4108


mRow = mRow + 1

Dim Rsp As String
Rsp = NumStrConv(Val(TxtInstrumentAmount))
tmp = "a" & mRow & ":" & "h" & mRow
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

tmp = "e" & mRow & ":" & "h" & mRow
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

tmp = "e" & mRow & ":" & "h" & mRow
oWS.Range(tmp).Merge

If IsNull(RsTxn_PledgeCashReceipt!HologramNo) Then
   oWS.Cells(mRow, C + 5) = "Hologram No:- Preview"
Else
  oWS.Cells(mRow, C + 5) = "Hologram No:- " & RsTxn_PledgeCashReceipt!HologramNo
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

tmp = "e" & mRow & ":" & "h" & mRow
oWS.Range(tmp).Merge

oWS.Cells(mRow, C + 5) = CreatedDateTime_
oWS.Cells(mRow, C + 5).Font.Size = 7
oWS.Cells(mRow, C + 5).Font.Bold = True
oWS.Cells(mRow, C + 5).WrapText = True
oWS.Cells(mRow, C + 5).HorizontalAlignment = -4108

If LCase(CmbInstrumentType.Text) = "cheque" Then
   mRow = mRow + 1

   tmp = "a" & mRow & ":" & "h" & mRow
   oWS.Range(tmp).Merge
   oWS.Cells(mRow, C + 1) = "Note :cheque dishonour would attract charges of RS. 500."
   oWS.Cells(mRow, C + 1).Font.Size = 7
   oWS.Cells(mRow, C + 1).Font.Bold = True
   oWS.Cells(mRow, C + 1).WrapText = True
End If
mRow = mRow + 1
tmp = "a" & mRow & ":" & "h" & mRow
oWS.Range(tmp).Merge


k = mRow

For Cell = 6 To k
    tmp = "a" & Cell & ":" & "h" & Cell
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

    'oWS.Name = "TI Details"
   ' oWS.Range("a1:g1").Merge
    oWS.Range("a:a").ColumnWidth = 15
    oWS.Range("b:b").ColumnWidth = 22
    oWS.Range("c:c").ColumnWidth = 15
    oWS.Range("d:d").ColumnWidth = 12
    oWS.Range("e:e").ColumnWidth = 13
    oWS.Range("f:f").ColumnWidth = 13
    oWS.Range("g:g").ColumnWidth = 10 '32.14
    oWS.Range("h:h").ColumnWidth = 13
    
    mStart = mRow
    tmp = "a" & mRow & ":" & "h" & mRow
    oWS.Range(tmp).Merge
    oWS.Cells(mRow, C + 1) = "Cash Receipt Details (Pledge related) for Receipt No: " & TxtPreFix & "-" & TxtPCRID & " Dated " & Format(TxtCRDate, "dd-mmm-yyyy")
    'oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
    oWS.Cells(mRow, C + 1).Font.Bold = True
    oWS.Cells(mRow, C + 1).Font.Size = 10
    
    mRow = mRow + 1
    'mCol = 5
    
    ''New
    oWS.Cells(mRow, C + 1) = "Pledge ID"
    oWS.Cells(mRow, C + 1).Font.Size = 7
    oWS.Cells(mRow, C + 1).Font.Bold = True
    
    oWS.Cells(mRow, C + 2) = "Pledge Date"
    oWS.Cells(mRow, C + 2).Font.Size = 7
    oWS.Cells(mRow, C + 2).Font.Bold = True
    
    oWS.Cells(mRow, C + 3) = "WHR No"
    oWS.Cells(mRow, C + 3).Font.Size = 7
    oWS.Cells(mRow, C + 3).Font.Bold = True
    
    oWS.Cells(mRow, C + 4) = "Client Name"
    oWS.Cells(mRow, C + 4).Font.Size = 7
    oWS.Cells(mRow, C + 4).Font.Bold = True
    
    oWS.Cells(mRow, C + 5) = "Pledge Fees"
    oWS.Cells(mRow, C + 5).Font.Size = 7
    oWS.Cells(mRow, C + 5).Font.Bold = True
    
    oWS.Cells(mRow, C + 6) = "Discount"
    oWS.Cells(mRow, C + 6).Font.Size = 7
    oWS.Cells(mRow, C + 6).Font.Bold = True
    
    oWS.Cells(mRow, C + 7) = "Received Amount"
    oWS.Cells(mRow, C + 7).Font.Size = 7
    oWS.Cells(mRow, C + 7).Font.Bold = True
    
    mRow = mRow + 1
    
    For k = 2 To MSHFlexGrid2.Rows - 2
    'k = 1
        oWS.Cells(mRow, C + 1) = MSHFlexGrid2.TextMatrix(k, 0)
        oWS.Cells(mRow, C + 1).Font.Size = 7
        
        oWS.Cells(mRow, C + 2) = " " & Format(MSHFlexGrid2.TextMatrix(k, 1), "dd-mmm-yyyy")
        oWS.Cells(mRow, C + 2).Font.Size = 7
        
        oWS.Cells(mRow, C + 3) = MSHFlexGrid2.TextMatrix(k, 2)
        oWS.Cells(mRow, C + 3).Font.Size = 7
        
        oWS.Cells(mRow, C + 4) = MSHFlexGrid2.TextMatrix(k, 3)
        oWS.Cells(mRow, C + 4).Font.Size = 7
        
        oWS.Cells(mRow, C + 5) = Val(MSHFlexGrid2.TextMatrix(k, 4))
        oWS.Cells(mRow, C + 5).Font.Size = 7
        
        oWS.Cells(mRow, C + 6) = Val(MSHFlexGrid2.TextMatrix(k, 5))
        oWS.Cells(mRow, C + 6).Font.Size = 7
        
        oWS.Cells(mRow, C + 7) = Val(MSHFlexGrid2.TextMatrix(k, 6))
        oWS.Cells(mRow, C + 7).Font.Size = 7
        
        mRow = mRow + 1
    Next k
    For Cell = mStart To mRow
        tmp = "a" & Cell & ":" & "h" & Cell
        oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
        oWS.Range(tmp).Font.Name = "Zurich BT"
        'oWS.Range(tmp).Font.Size = 7
        oWS.Range(tmp).WrapText = True
    Next Cell

If Gen_ExcelPassword <> "" Then
   oWS.Protect Gen_ExcelPassword, , , , , , True, True
End If
oWB.Save

Call Log_Print_History(TxtPreFix, "CR", TxtPCRID, Gen_FLName)

MsgBox ("Excel file created !!! ")
oXL.Visible = True

End Sub


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
objPDF.PDFTitle = "Pledge Cash Receipt PDF Document"
objPDF.PDFFileName = App.Path & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "_Pledge Cash Receipt-" & Trim(TxtPreFix) & "" & Trim(TxtPCRID) & ".pdf"
mPDFFileName = App.Path & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "_Pledge Cash Receipt-" & Trim(TxtPreFix) & "" & Trim(TxtPCRID) & ".pdf"
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
   objPDF.PDFWatermark = "NCMSL"
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
mStrLength = Len("Pledge " & CmbInstrumentType.Text & " Receipt")
mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
mTxtLeft = mPrintPoint - (mStrLength * 2) '* 5) '270
mFontSize = 10

Call PrintText(mFontSize, FONT_BOLD, ALIGN_CENTER, mTxtLeft - 26, mTxtTop - 3, "Pledge " & CmbInstrumentType.Text & " Receipt")

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
Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop - 5, TxtPreFix & "-" & TxtPCRID)
'--
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 280, mTxtTop - 5, "Receipt Date")
Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 430, mTxtTop - 5, Format(TxtCRDate, "dd-MMM-yyyy"))
'--

mTxtTop = mTxtTop + mRowHeight + mRowHeight

mstr = "Received with thanks from Shri " & TxtReceivedFrom & " a sum of Rs " & TxtInstrumentAmount & " ( Rupees " & Rsp & " ) as per the details enclosed, by " & CmbInstrumentType & ""
If CmbInstrumentType.Text <> "Cash" Then
   mstr = mstr & " No: " & TxtInstrumentNo & " dated " & Format(TxtInstrumentDate.Value, "dd-MMM-yyyy") & " drawn on " & TxtDrawnOn & " " & TxtBranchName & " branch (subject to realisation). "
End If

mMax = 150
mPara = 7
t1 = 1
mLen = Len(mstr)

If mLen > mMax Then
   mLen = RoundUp(mLen / mMax)
Else
   mLen = 1
End If
For T = 1 To mLen
    If T > 1 Then
       t1 = mMax + t1
       mPara = mPara + 9
    End If
    Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 20, mTxtTop + mPara - 5 - 7, mID$(mstr, t1, mMax))
Next

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

Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 430, mTxtTop - 5, "Hologram No: " & IIf(IsNull(RsTxn_PledgeCashReceipt!HologramNo), "", RsTxn_PledgeCashReceipt!HologramNo))

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

mTxtTop = mTxtTop + mRowHeight
Call PrintText(mDispTextSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Cash Receipt Details (Pledge related) for Receipt No: " & TxtPreFix & "-" & TxtPCRID & " Dated " & Format(TxtCRDate, "dd-mmm-yyyy"))
objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop
mTxtTop = mTxtTop + mRowHeight

Call PrintText(mDispTextSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Pledge ID")
Call PrintText(mDispTextSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 50, mTxtTop - 5, "Pledge Date")
Call PrintText(mDispTextSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 100, mTxtTop - 5, "WHR No")
Call PrintText(mDispTextSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop - 5, "Pledged Bank")
Call PrintText(mDispTextSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 240, mTxtTop - 5, "Client Name")
Call PrintText(mDispTextSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 400, mTxtTop - 5, "Pledge Fees")
Call PrintText(mDispTextSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 450, mTxtTop - 5, "Discount")
Call PrintText(mDispTextSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 500, mTxtTop - 5, "Received Amount")

objPDF.PDFDrawLine mGlobalLeft + 48, mTxtTop - mRowHeight, mGlobalLeft + 48, mTxtTop
objPDF.PDFDrawLine mGlobalLeft + 98, mTxtTop - mRowHeight, mGlobalLeft + 98, mTxtTop
objPDF.PDFDrawLine mGlobalLeft + 138, mTxtTop - mRowHeight, mGlobalLeft + 138, mTxtTop
objPDF.PDFDrawLine mGlobalLeft + 238, mTxtTop - mRowHeight, mGlobalLeft + 238, mTxtTop
objPDF.PDFDrawLine mGlobalLeft + 398, mTxtTop - mRowHeight, mGlobalLeft + 398, mTxtTop
objPDF.PDFDrawLine mGlobalLeft + 448, mTxtTop - mRowHeight, mGlobalLeft + 448, mTxtTop
objPDF.PDFDrawLine mGlobalLeft + 498, mTxtTop - mRowHeight, mGlobalLeft + 498, mTxtTop


objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop
 mTxtTop = mTxtTop + (mRowHeight)
 
For k = 2 To MSHFlexGrid2.Rows - 2
    If mTxtTop > objPDF.PDFGetPageHeight Then
       Call CheckForNewPage
    End If
    
    Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, MSHFlexGrid2.TextMatrix(k, 0))
    Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 50, mTxtTop - 5, MSHFlexGrid2.TextMatrix(k, 1))
    Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 100, mTxtTop - 5, MSHFlexGrid2.TextMatrix(k, 2))
    
    mMax = 26
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
        Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop + mPara - 5 - 7, mID$(MSHFlexGrid2.TextMatrix(k, 10), t1, mMax))
    Next
    
    
    mMax = 26
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
        Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 240, mTxtTop + mPara - 5 - 7, mID$(MSHFlexGrid2.TextMatrix(k, 3), t1, mMax))
    Next
    
    a1 = mTxtTop + mPara - 5

    Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 400, mTxtTop - 5, Format(MSHFlexGrid2.TextMatrix(k, 4), "#####0.00"))
    Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 450, mTxtTop - 5, Format(MSHFlexGrid2.TextMatrix(k, 5), "#####0.00"))
    Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 500, mTxtTop - 5, Format(MSHFlexGrid2.TextMatrix(k, 8), "#####0.00"))
    
     
    objPDF.PDFDrawLine mGlobalLeft + 48, mTxtTop - mRowHeight, mGlobalLeft + 48, Val(a1)
    objPDF.PDFDrawLine mGlobalLeft + 98, mTxtTop - mRowHeight, mGlobalLeft + 98, Val(a1)
    objPDF.PDFDrawLine mGlobalLeft + 138, mTxtTop - mRowHeight, mGlobalLeft + 138, Val(a1)
    objPDF.PDFDrawLine mGlobalLeft + 238, mTxtTop - mRowHeight, mGlobalLeft + 238, Val(a1)
    objPDF.PDFDrawLine mGlobalLeft + 398, mTxtTop - mRowHeight, mGlobalLeft + 398, Val(a1)
    objPDF.PDFDrawLine mGlobalLeft + 448, mTxtTop - mRowHeight, mGlobalLeft + 448, Val(a1)
    objPDF.PDFDrawLine mGlobalLeft + 498, mTxtTop - mRowHeight, mGlobalLeft + 498, Val(a1)
    objPDF.PDFDrawLine mGlobalLeft, Val(a1), objPDF.PDFGetPageWidth - 18, Val(a1)
    
    mTxtTop = a1 + mRowHeight
Next

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
      objPDF.PDFWatermark = "NCMSL"
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
    'mLeft_ = mmasterLeft

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




