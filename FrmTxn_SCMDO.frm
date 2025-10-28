VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmTxn_SCMDO 
   Caption         =   "SCM DO"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.VScrollBar VScroll1 
      Height          =   9015
      Left            =   15120
      TabIndex        =   52
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      TabIndex        =   51
      Top             =   10320
      Visible         =   0   'False
      Width           =   10935
   End
   Begin VB.PictureBox Frame4 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   9975
      Left            =   0
      ScaleHeight     =   9945
      ScaleWidth      =   15105
      TabIndex        =   10
      Top             =   0
      Width           =   15135
      Begin VB.PictureBox Frame0 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1695
         Left            =   240
         ScaleHeight     =   1665
         ScaleWidth      =   14985
         TabIndex        =   11
         Top             =   3600
         Visible         =   0   'False
         Width           =   15015
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
            ItemData        =   "FrmTxn_SCMDO.frx":0000
            Left            =   4440
            List            =   "FrmTxn_SCMDO.frx":000D
            Sorted          =   -1  'True
            TabIndex        =   88
            Top             =   1425
            Width           =   2085
         End
         Begin VB.TextBox TxtSWHRNo 
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
            Height          =   315
            Left            =   1485
            MaxLength       =   50
            TabIndex        =   86
            Top             =   1425
            Width           =   1905
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
            Height          =   360
            Left            =   13725
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   1380
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
            Height          =   360
            Left            =   12720
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   1380
            Width           =   1000
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
            ItemData        =   "FrmTxn_SCMDO.frx":002E
            Left            =   2865
            List            =   "FrmTxn_SCMDO.frx":0030
            Sorted          =   -1  'True
            TabIndex        =   15
            Top             =   450
            Width           =   3030
         End
         Begin VB.ComboBox CmbSCommodity 
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
            ItemData        =   "FrmTxn_SCMDO.frx":0032
            Left            =   1485
            List            =   "FrmTxn_SCMDO.frx":0034
            Sorted          =   -1  'True
            TabIndex        =   14
            Top             =   960
            Width           =   4065
         End
         Begin VB.ComboBox CmbSClientName 
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
            ItemData        =   "FrmTxn_SCMDO.frx":0036
            Left            =   6840
            List            =   "FrmTxn_SCMDO.frx":0038
            Sorted          =   -1  'True
            TabIndex        =   13
            Top             =   960
            Width           =   7935
         End
         Begin VB.ComboBox CmbSCMPName 
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
            ItemData        =   "FrmTxn_SCMDO.frx":003A
            Left            =   5925
            List            =   "FrmTxn_SCMDO.frx":003C
            Sorted          =   -1  'True
            TabIndex        =   12
            Top             =   450
            Width           =   8850
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   1455
            Left            =   120
            TabIndex        =   18
            Top             =   1860
            Width           =   14595
            _ExtentX        =   25744
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
         Begin MSComCtl2.DTPicker TxtSFrom 
            Height          =   360
            Left            =   120
            TabIndex        =   19
            Top             =   450
            Width           =   1335
            _ExtentX        =   2355
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
         Begin MSComCtl2.DTPicker TxtSTo 
            Height          =   360
            Left            =   1485
            TabIndex        =   20
            Top             =   450
            Width           =   1335
            _ExtentX        =   2355
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
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   3360
            Width           =   14595
            _ExtentX        =   25744
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   1
            Scrolling       =   1
         End
         Begin VB.Label Label6 
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
            Height          =   225
            Left            =   3840
            TabIndex        =   89
            Top             =   1493
            Width           =   810
         End
         Begin VB.Label Label32 
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
            TabIndex        =   87
            Top             =   1462
            Width           =   960
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
            Left            =   8280
            TabIndex        =   29
            Top             =   1440
            Width           =   2415
         End
         Begin VB.Label Label45 
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
            Left            =   1785
            TabIndex        =   28
            Top             =   120
            Width           =   735
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
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
            Height          =   240
            Left            =   315
            TabIndex        =   27
            Top             =   120
            Width           =   945
         End
         Begin VB.Label Label22 
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
            Left            =   10890
            TabIndex        =   26
            Top             =   1440
            Width           =   165
         End
         Begin VB.Label Label11 
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
            Left            =   3998
            TabIndex        =   25
            Top             =   120
            Width           =   765
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Commodity"
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
            TabIndex        =   24
            Top             =   1020
            Width           =   1020
         End
         Begin VB.Label Label14 
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
            Left            =   5715
            TabIndex        =   23
            Top             =   1020
            Width           =   1095
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "CMP Name"
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
            Left            =   9833
            TabIndex        =   22
            Top             =   120
            Width           =   1035
         End
      End
      Begin VB.PictureBox Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   6555
         Left            =   120
         ScaleHeight     =   6525
         ScaleWidth      =   14850
         TabIndex        =   37
         Top             =   60
         Width           =   14880
         Begin VB.TextBox TxtPrevDOBags 
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
            Left            =   1755
            Locked          =   -1  'True
            TabIndex        =   90
            TabStop         =   0   'False
            Top             =   2325
            Width           =   2385
         End
         Begin VB.TextBox TxtBalanceQty 
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
            Left            =   12315
            Locked          =   -1  'True
            TabIndex        =   82
            TabStop         =   0   'False
            Top             =   2325
            Width           =   1380
         End
         Begin VB.TextBox TxtBalanceBags 
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
            Left            =   8475
            Locked          =   -1  'True
            TabIndex        =   81
            TabStop         =   0   'False
            Top             =   2325
            Width           =   2385
         End
         Begin VB.TextBox TxtPrevDOQty 
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
            Left            =   5190
            Locked          =   -1  'True
            TabIndex        =   80
            TabStop         =   0   'False
            Top             =   2325
            Width           =   1380
         End
         Begin VB.TextBox TxtPledgeQty 
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
            Left            =   12315
            Locked          =   -1  'True
            TabIndex        =   77
            TabStop         =   0   'False
            Top             =   1875
            Width           =   1380
         End
         Begin VB.TextBox TxtPledgeBags 
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
            Left            =   8475
            Locked          =   -1  'True
            TabIndex        =   76
            TabStop         =   0   'False
            Top             =   1875
            Width           =   2385
         End
         Begin VB.TextBox TxtApproved 
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
            Left            =   1755
            Locked          =   -1  'True
            TabIndex        =   74
            TabStop         =   0   'False
            Top             =   5580
            Width           =   5745
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
            Left            =   8625
            Locked          =   -1  'True
            TabIndex        =   71
            TabStop         =   0   'False
            Top             =   5160
            Width           =   6630
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
            Height          =   360
            Left            =   1755
            Locked          =   -1  'True
            TabIndex        =   70
            TabStop         =   0   'False
            Top             =   5160
            Width           =   5745
         End
         Begin VB.TextBox TxtWHRDate 
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
            Left            =   8475
            Locked          =   -1  'True
            TabIndex        =   69
            TabStop         =   0   'False
            Top             =   1440
            Width           =   2385
         End
         Begin VB.TextBox TxtDONoOfBags 
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
            Left            =   1755
            TabIndex        =   2
            Top             =   2805
            Width           =   2385
         End
         Begin VB.TextBox TxtDOQty 
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
            Left            =   5190
            TabIndex        =   3
            Top             =   2805
            Width           =   1380
         End
         Begin VB.TextBox TxtPledgeID 
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
            Left            =   12315
            Locked          =   -1  'True
            TabIndex        =   65
            TabStop         =   0   'False
            Top             =   1440
            Width           =   2445
         End
         Begin VB.TextBox TxtWHRNo 
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
            Left            =   8475
            TabIndex        =   59
            Top             =   135
            Visible         =   0   'False
            Width           =   510
         End
         Begin VB.TextBox TxtWHRNoOfBags 
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
            Left            =   1755
            Locked          =   -1  'True
            TabIndex        =   58
            TabStop         =   0   'False
            Top             =   1875
            Width           =   2385
         End
         Begin VB.TextBox TxtWHRQuantity 
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
            Left            =   5190
            Locked          =   -1  'True
            TabIndex        =   57
            TabStop         =   0   'False
            Top             =   1875
            Width           =   1380
         End
         Begin VB.TextBox TxtGodownNo 
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
            Left            =   1755
            Locked          =   -1  'True
            TabIndex        =   56
            TabStop         =   0   'False
            Top             =   1440
            Width           =   4815
         End
         Begin VB.TextBox TxtWHRNoNew 
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
            Left            =   8475
            TabIndex        =   1
            Top             =   135
            Width           =   2310
         End
         Begin VB.TextBox TxtApprovalRemarks 
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
            Height          =   900
            Left            =   1755
            MaxLength       =   5000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   6
            Top             =   4200
            Width           =   12945
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
            ItemData        =   "FrmTxn_SCMDO.frx":003E
            Left            =   8475
            List            =   "FrmTxn_SCMDO.frx":004B
            Sorted          =   -1  'True
            TabIndex        =   4
            Top             =   2805
            Width           =   2385
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
            Height          =   900
            Left            =   1755
            MaxLength       =   5000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   5
            Top             =   3240
            Width           =   12945
         End
         Begin VB.TextBox TxtDOID 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   360
            Left            =   1755
            Locked          =   -1  'True
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   135
            Width           =   2400
         End
         Begin MSComCtl2.DTPicker TxtDODate 
            Height          =   360
            Left            =   5190
            TabIndex        =   39
            Top             =   135
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
            CurrentDate     =   39884
         End
         Begin VB.TextBox TxtLocation 
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
            Left            =   12315
            Locked          =   -1  'True
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   135
            Width           =   2460
         End
         Begin VB.TextBox TxtCommodity 
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
            Left            =   8475
            Locked          =   -1  'True
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   585
            Width           =   6285
         End
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
            Left            =   1755
            Locked          =   -1  'True
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   1035
            Width           =   13005
         End
         Begin VB.TextBox TxtStatus 
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
            Left            =   8475
            Locked          =   -1  'True
            TabIndex        =   53
            TabStop         =   0   'False
            Top             =   2805
            Width           =   2385
         End
         Begin VB.TextBox TxtCMPName 
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
            Left            =   1755
            Locked          =   -1  'True
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   600
            Width           =   4815
         End
         Begin VB.Label Label29 
            Caption         =   "Prev. DO Bags"
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
            Left            =   150
            TabIndex        =   91
            Top             =   2385
            Width           =   1350
         End
         Begin VB.Label Label31 
            Caption         =   "Balance Qty"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   11190
            TabIndex        =   85
            Top             =   2235
            Width           =   945
         End
         Begin VB.Label Label30 
            Caption         =   "Balance No of Bags"
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
            Left            =   7080
            TabIndex        =   84
            Top             =   2265
            Width           =   1395
         End
         Begin VB.Label Label28 
            Caption         =   "Prev. DO Qty"
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
            Left            =   4320
            TabIndex        =   83
            Top             =   2265
            Width           =   900
         End
         Begin VB.Label Label27 
            Caption         =   "Pledge Qty"
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
            Left            =   11190
            TabIndex        =   79
            Top             =   1935
            Width           =   1065
         End
         Begin VB.Label Label26 
            Caption         =   "Pledge Bags"
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
            Left            =   7080
            TabIndex        =   78
            Top             =   1935
            Width           =   1350
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "Approved  :"
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
            Left            =   150
            TabIndex        =   75
            Top             =   5640
            Width           =   1035
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Updated  :"
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
            Left            =   7650
            TabIndex        =   73
            Top             =   5220
            Width           =   930
         End
         Begin VB.Label Label9 
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
            Left            =   150
            TabIndex        =   72
            Top             =   5220
            Width           =   855
         End
         Begin VB.Label Label24 
            Caption         =   "DO No of Bags"
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
            Left            =   150
            TabIndex        =   68
            Top             =   2865
            Width           =   1590
         End
         Begin VB.Label Label20 
            Caption         =   "DO Qty"
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
            Left            =   4320
            TabIndex        =   67
            Top             =   2865
            Width           =   945
         End
         Begin VB.Label Label19 
            Caption         =   "Pledge ID"
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
            Left            =   11190
            TabIndex        =   66
            Top             =   1500
            Width           =   990
         End
         Begin VB.Label Label18 
            Caption         =   "WHR No"
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
            Left            =   7080
            TabIndex        =   64
            Top             =   180
            Width           =   1335
         End
         Begin VB.Label Label10 
            Caption         =   "WHR Date"
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
            Left            =   7080
            TabIndex        =   63
            Top             =   1500
            Width           =   1215
         End
         Begin VB.Label Label16 
            Caption         =   "WHR No of Bags"
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
            Left            =   150
            TabIndex        =   62
            Top             =   1935
            Width           =   1590
         End
         Begin VB.Label Label8 
            Caption         =   "WHR Qty"
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
            Left            =   4320
            TabIndex        =   61
            Top             =   1935
            Width           =   945
         End
         Begin VB.Label Label21 
            Caption         =   "Godown No"
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
            Left            =   150
            TabIndex        =   60
            Top             =   1500
            Width           =   1455
         End
         Begin VB.Label Label7 
            Caption         =   "Approval Remarks"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   150
            TabIndex        =   55
            Top             =   4410
            Width           =   1260
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
            Left            =   7080
            TabIndex        =   54
            Top             =   2858
            Width           =   960
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Location"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   11190
            TabIndex        =   50
            Top             =   165
            Width           =   840
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Commodity"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   7080
            TabIndex        =   49
            Top             =   645
            Width           =   1080
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Client Name"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   150
            TabIndex        =   48
            Top             =   1095
            Width           =   1170
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "CMP Name"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   150
            TabIndex        =   47
            Top             =   645
            Width           =   1035
         End
         Begin VB.Label LbClaimID 
            AutoSize        =   -1  'True
            Caption         =   "DO ID"
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
            Left            =   150
            TabIndex        =   46
            Top             =   195
            Width           =   540
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Remarks"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   150
            TabIndex        =   45
            Top             =   3570
            Width           =   810
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            Caption         =   "DO Date"
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
            TabIndex        =   44
            Top             =   195
            Width           =   780
         End
      End
      Begin VB.PictureBox ButtonFrm 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1035
         Left            =   120
         ScaleHeight     =   1005
         ScaleWidth      =   14865
         TabIndex        =   30
         Top             =   6720
         Width           =   14895
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
            Left            =   270
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   525
            Width           =   2835
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
            Left            =   3105
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   525
            Width           =   2835
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
            Left            =   5940
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   525
            Width           =   2835
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
            Left            =   8775
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   525
            Width           =   2835
         End
         Begin VB.CommandButton AddCmd 
            Caption         =   "&Add New "
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
            Left            =   270
            Style           =   1  'Graphical
            TabIndex        =   0
            Top             =   120
            Width           =   2835
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
            Left            =   3105
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   120
            Width           =   2835
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
            Left            =   5940
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   120
            Width           =   2835
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
            Left            =   8775
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   120
            Width           =   2835
         End
         Begin VB.CommandButton SearchCmd 
            Caption         =   "&Search"
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
            Left            =   11610
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   120
            Width           =   2835
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
            Left            =   11610
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   525
            Width           =   2835
         End
      End
   End
End
Attribute VB_Name = "FrmTxn_SCMDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mLocationID, mSM_Prefix As Variant
Dim Edit_Flg As Variant
Dim mSCommodityID, mSLocationID, mCommodityID As Variant

Private Sub AddCmd_Click()


Edit_Flg = "A"
Frame1.Enabled = True
TxtWHRNoNew.Locked = False
TxtWHRNoNew = ""
TxtWHRNo = ""
TxtDOID = ""
TxtDODate.Value = Date

Call ClearDetail

TxtRemarks.Text = ""
TxtApprovalRemarks = ""
CmbStatus.Text = "Approved"
TxtStatus = "Approved"
TxtApprovalRemarks.Locked = True
CmbStatus.Visible = False
TxtStatus.Visible = True

TxtDONoOfBags.Locked = False
TxtDOQty.Locked = False

TxtCreated = ""
TxtUpdated = ""
TxtApproved = ""


Call GButtonLock(Me, False)


End Sub
Private Sub CmbSStatus_LostFocus()
If Trim(CmbSStatus.Text) = "" Then Exit Sub

If LCase(CmbSStatus.Text) <> "pending" And LCase(CmbSStatus.Text) <> "approved" And LCase(CmbSStatus.Text) <> "rejected" Then
   MsgBox "Invalid Selection!!!"
   CmbSStatus.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmbStatus_LostFocus()
If CmbStatus.Text = "" Then Exit Sub
   
If LCase(CmbStatus.Text) <> "pending" And LCase(CmbStatus.Text) <> "approved" And LCase(CmbStatus.Text) <> "rejected" Then
   MsgBox "Invalid Selection!!!"
   CmbStatus.SetFocus
   Exit Sub
End If

If LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "headcmg" And LCase(Gen_UserRole) <> "admin" And LCase(Gen_UserRole) <> "superadmin" Then
   If LCase(CmbStatus.Text) = "approved" Then
      MsgBox "Only CMG Department can approve DO!!!"
      CmbStatus.SetFocus
      Exit Sub
   End If
End If

If LCase(CmbStatus.Text) = "pending" Then
   TxtApprovalRemarks = ""
   TxtApprovalRemarks.Locked = True
Else
   TxtApprovalRemarks.Locked = False
End If
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Txn_SCMDO.xls")
Call Xls_Detail_Rpt
End Sub

Public Sub Xls_Detail_Rpt()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Integer
Dim mCol As Integer
MsgBox ("Wait till Next Message!!!")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Txn_SCMDO.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid1.Rows - 1
   mRow = I + 1
   For C = 0 To MSHFlexGrid1.Cols - 1
      If C <> 6 Then
         oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
      Else
         oWS.Cells(mRow, C + 1) = "'" & MSHFlexGrid1.TextMatrix(I, C)
      End If
   Next
Next
oWB.Save
MsgBox ("Excel File Created!!!")
oXL.Visible = True
End Sub

Private Sub CmdGo_Click()

Dim wc As Variant

Call Grid_Head
'MsgBox "Wait till next message!!!"
Label22.Caption = ""
wc = ""

If Gen_DBType = "MDB" Then
   wc = " Where TSCM.DODate Between #" & Format(TxtSFrom.Value, Gen_DatFormat) & "# And #" & Format(TxtSTo.Value, Gen_DatFormat) & "# "
Else
   wc = " Where TSCM.DODate Between '" & Format(TxtSFrom.Value, Gen_DatFormat) & "' And '" & Format(TxtSTo.Value, Gen_DatFormat) & "' "
End If

If Trim(CmbSLocation.Text) <> "" Then
   wc = wc & " And LM.LocationName = '" & CmbSLocation.Text & "'"
Else
   If Gen_WcLocation <> "" Then
      wc = wc & " And LM." & Gen_WcLocation
   End If
End If
If Trim(CmbSCMPName.Text) <> "" Then
   wc = wc & " And CMP.CMPName = '" & CmbSCMPName.Text & "'"
Else
   If Ge_WcCMP <> "" Then
      wc = wc & " And CMP." & Gen_WcCMP
   End If
End If

If Trim(CmbSCommodity.Text) <> "" Then
   wc = wc & " And ComM.Commodity = '" & CmbSCommodity.Text & "'"
End If

If Trim(CmbSClientName.Text) <> "" Then
   wc = wc & " And CM.ClientName = '" & CmbSClientName.Text & "'"
End If

If Trim(CmbSStatus.Text) <> "" Then
   wc = wc & " And TSCM.Flag= '" & Left$(CmbSStatus.Text, 1) & "'"
End If


If Trim(TxtSWHRNo) <> "" Then
   Dim strSplitString() As String
   Dim mSWHRNo As Variant
   strSplitString = Split(TxtSWHRNo.Text, "-")
   tmp1 = UBound(strSplitString)
   If tmp1 <> 1 Then
      MsgBox "Invalid WHR No selection"
      TxtSearchPledge.SetFocus
      Exit Sub
   End If
   mSWHRNo = Val(strSplitString(0)) * 1000000 + Val(strSplitString(1))
   wc = wc & " And TSCM.WHRNo= " & mSWHRNo & ""
End If


tmp = " SELECT  TSCM.DOID,TSCM.DODate,TWD.WHRNo,TWD.SM_Prefix,CMP.CMPName,CM.ClientName,ComM.Commodity,TWD.WHRDate,"
tmp = tmp & " TWD.ClientIDRow , TWD.CommodityID, TWD.CMPID, TWD.WHRNoofBags, TWD.WHRQuantity"
tmp = tmp & " ,TWD.PledgeID,TWD.CMPID,"
tmp = tmp & " TSCM.DOBags,TSCM.DOQty,TSCM.Remarks,TSCM.ApprovalRemarks,"
tmp = tmp & " LM.LocationName,LM.LocationID,GM.GodownNo,TSCM.Flag "
tmp = tmp & " FROM  Txn_SCMDO TSCM "
tmp = tmp & " Inner Join Txn_WHR_Detail TWD On TSCM.WHRNo = TWD.WHRNo"
tmp = tmp & " Inner Join Mst_Location LM ON TSCM.LocationID=LM.LocationID"
tmp = tmp & " Inner Join Mst_CMP CMP ON TWD.CMPID=CMP.CMPID"
tmp = tmp & " Inner Join Mst_Godown GM ON TWD.GodownID=GM.GodownID"
tmp = tmp & " Inner Join Mst_Client CM ON TWD.ClientIDRow=CM.ClientIDRow"
tmp = tmp & " Inner Join Mst_Commodity ComM ON TWD.CommodityID=ComM.CommodityID"
tmp = tmp & " "
tmp = tmp & wc
tmp = tmp & " Order by TSCM.DOID"

Call TmpTable
If TmpRs.RecordCount > 0 Then
   CmdExcel.Enabled = True
   With MSHFlexGrid1
      .Rows = .Rows + TmpRs.RecordCount
      For iu = 1 To TmpRs.RecordCount
          .TextMatrix(iu, 0) = IIf(IsNull(TmpRs!DOID), "", TmpRs!DOID) '"DO ID"
          .TextMatrix(iu, 1) = IIf(IsNull(TmpRs!DODate), "", Format(TmpRs!DODate, "dd-mmm-yyyy")) '"DO Date"
          .TextMatrix(iu, 2) = IIf(IsNull(TmpRs!LocationName), "", TmpRs!LocationName) '"Location Name"
          .TextMatrix(iu, 3) = IIf(IsNull(TmpRs!CMPName), "", TmpRs!CMPName) '"CMP"
          .TextMatrix(iu, 4) = IIf(IsNull(TmpRs!Commodity), "", TmpRs!Commodity) '"Commodity"
          .TextMatrix(iu, 5) = IIf(IsNull(TmpRs!ClientName), "", TmpRs!ClientName) '"Client Name"
          .TextMatrix(iu, 6) = IIf(IsNull(TmpRs!CMPID), 0, TmpRs!CMPID) & "-" & TmpRs!WHRNo - (IIf(IsNull(TmpRs!CMPID), 0, TmpRs!CMPID) * 1000000) '"WHR No"
          .TextMatrix(iu, 7) = IIf(IsNull(TmpRs!DOBags), "", TmpRs!DOBags) '"DO No Of Bags"
          .TextMatrix(iu, 8) = IIf(IsNull(TmpRs!DOQty), "", TmpRs!DOQty) '"DO Qty"
          .TextMatrix(iu, 9) = IIf(IsNull(TmpRs!Flag), "", TmpRs!Flag) '"Status"
          .TextMatrix(iu, 10) = IIf(IsNull(TmpRs!Remarks), "", TmpRs!Remarks) '"Remarks"
          .TextMatrix(iu, 11) = IIf(IsNull(TmpRs!ApprovalRemarks), "", TmpRs!ApprovalRemarks) '"Approval Remarks"
         
          If LCase(.TextMatrix(iu, 9)) = "p" Then
             .TextMatrix(iu, 9) = "Pending"
          ElseIf LCase(.TextMatrix(iu, 9)) = "a" Then
             .TextMatrix(iu, 9) = "Approved"
          ElseIf LCase(.TextMatrix(iu, 9)) = "r" Then
             .TextMatrix(iu, 9) = "Rejected"
          End If
          TmpRs.MoveNext
      Next
   End With
   Label22.Caption = TmpRs.RecordCount
End If

End Sub


Private Sub EditCmd_Click()

Edit_Flg = "E"

If LCase(TxtStatus.Text) = "rejected" Or LCase(TxtStatus.Text) = "approved" Then
   MsgBox "This DO entry is " & TxtStatus & " . Can't Edit!!!"
   Exit Sub
End If

Call GButtonLock(Me, False)


Frame1.Enabled = True
TxtDODate.Enabled = False
TxtWHRNoNew.Locked = True
TxtDONoOfBags.Locked = True
TxtDOQty.Locked = True

CmbStatus.Visible = True
TxtStatus.Visible = False
CmbStatus.Text = TxtStatus


End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub Form_Load()

Dim wc As Variant

Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Frame1.Enabled = False
   Call GButtonLockAD(Me)
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If

'Call Grid_Head1
'SFrom.Value = Date - 365
'STo.Value = Date
'Label22.Caption = ""
'CmbSLocationID.Text = ""
'TxtHologramNo.Text = ""


TxtSFrom.Value = Date - 30
TxtSTo.Value = Date
CmbSLocation.Text = ""
CmbSCMPName.Text = ""
CmbSCommodity.Text = ""
CmbSClientName.Text = ""
TxtSWHRNo = ""
Label22.Caption = ""
CmbSStatus.Text = ""
Call Grid_Head

Frame1.Enabled = True
If Gen_WcLocation <> "" Then
   Call TableTxn_SCMDO(" Where " & Gen_WcLocation & " Order by DOID")
Else
   Call TableTxn_SCMDO(" Order by DOID")
End If

If RsTxn_SCMDO.RecordCount = 0 Then
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

Private Sub MSHFlexGrid1_Click()
Dim I As Variant

I = MSHFlexGrid1.RowSel
If I <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(I, 0) = "" Then Exit Sub
RsTxn_SCMDO.MoveFirst
RsTxn_SCMDO.Find "DOID= " & Val(MSHFlexGrid1.TextMatrix(I, 0)) & ""
If Not RsTxn_SCMDO.EOF Then
   Call Indata
   Frame0.Visible = False
End If
End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsTxn_SCMDO.MoveNext
If RsTxn_SCMDO.EOF Then
   RsTxn_SCMDO.MoveLast
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
RsTxn_SCMDO.MovePrevious
If RsTxn_SCMDO.BOF Then
   RsTxn_SCMDO.MoveFirst
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


Private Sub LastCmd_Click()
RsTxn_SCMDO.MoveLast
Call Indata
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False
End Sub
Private Sub FirstCmd_Click()
RsTxn_SCMDO.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub

Private Sub SaveCmd_Click()

If TxtWHRNoNew = "" Then
   MsgBox "Blank WHR No Not allowed!!!"
   TxtWHRNoNew.SetFocus
   Exit Sub
End If

If TxtLocation = "" Then
   MsgBox "Please provide WHR No !!!"
   TxtWHRNoNew.SetFocus
   Exit Sub
End If

If Val(TxtDONoOfBags) = 0 Then
   MsgBox "Blank DO Bags not allowed!!!"
   TxtDONoOfBags.SetFocus
   Exit Sub
End If

If Val(TxtDOQty) = 0 Then
   MsgBox "Blank DO Qty not allowed!!!"
   TxtDOQty.SetFocus
   Exit Sub
End If

If Edit_Flg = "A" Then

   If Val(TxtDONoOfBags) > Val(TxtWHRNoOfBags) Then
      MsgBox "DO No Of Bags must be less than WHR No of Bags!!!"
      TxtDONoOfBags.SetFocus
      Exit Sub
   End If

   If Val(TxtDOQty) > Val(TxtWHRQuantity) Then
      MsgBox "DO Qty must be less than WHR Qty!!!"
      TxtDOQty.SetFocus
      Exit Sub
   End If

   If Val(TxtDONoOfBags) > Val(TxtPledgeBags) Then
      MsgBox "DO No Of Bags must be less than Pledge No of Bags!!!"
      TxtDONoOfBags.SetFocus
      Exit Sub
   End If

   If Val(TxtDOQty) > Val(TxtPledgeQty) Then
      MsgBox "DO Qty must be less than Pledge Qty!!!"
      TxtDOQty.SetFocus
      Exit Sub
   End If


   If Val(TxtDONoOfBags) > Val(TxtPledgeBags) - Val(TxtPrevDOBags) Then
      MsgBox "DO No Of Bags must be less than Diffrence of Pledge and Previous DO No of Bags!!!"
      TxtDONoOfBags.SetFocus
      Exit Sub
   End If

   If Val(Format(Val(TxtDOQty), "#######0.000")) > Val(Format(Val(TxtPledgeQty) - Val(TxtPrevDOQty), "#######0.000")) Then
      MsgBox "DO Qty must be less than Diffrence of Pledge and Previous DO Qty!!!"
      TxtDOQty.SetFocus
      Exit Sub
   End If
   
   
   If Val(TxtPledgeBags) - Val(TxtPrevDOBags) = Val(TxtDONoOfBags) Then
      If Format(Val(TxtDOQty), "#######0.000") > Format(Val(TxtPledgeQty) - Val(TxtPrevDOQty), "#######0.000") Then 'If Val(TxtPledgeQty) - Val(TxtPrevDOQty) <> Val(TxtDOQty) Then
         MsgBox "When all the bags are released, all the quantity are also required to be released !!!! "
         TxtDONoOfBags.SetFocus
         Exit Sub
      End If
   End If

   If Val(TxtPrevDOQty) <> Val(TxtDOQty) = Val(TxtDOQty) Then
      If Val(TxtPledgeBags) - Val(TxtPrevDOBags) <> Val(TxtDONoOfBags) Then
         MsgBox "When all the quantity are released, all the  bags are also required to be released !!!! "
         TxtDOQty.SetFocus
         Exit Sub
      End If
   End If
'
   tmp = Validate_ComBags(mCommodityID, Val(TxtDONoOfBags), Val(TxtDOQty))
   If Len(tmp) > 1 Then
      ans = MsgBox(tmp, vbYesNo) 'MsgBox tmp
      If ans = vbNo Then
         TxtDOQty.SetFocus
         Exit Sub
      End If
   End If
   
ElseIf Edit_Flg = "E" Then
  If LCase(CmbStatus.Text) = "approved" Then
     If TxtApprovalRemarks = "" Then
        MsgBox "Blank Approval Remarks not allowed!!!"
        TxtApprovalRemarks.SetFocus
        Exit Sub
     End If
  End If
End If

If Edit_Flg = "A" Then
   RsTxn_SCMDO.AddNew
   RsTxn_SCMDO!DOID = GetMaxDOID
   RsTxn_SCMDO!G_UID = GetGUID
   TxtDOID.Text = RsTxn_SCMDO!DOID
Else
   RsTxn_SCMDO.MoveFirst
   RsTxn_SCMDO.Find " DOID= '" & Val(TxtDOID.Text) & "'"
End If

RsTxn_SCMDO!DODate = Format(TxtDODate.Value, Gen_DatFormat)
RsTxn_SCMDO!LocationID = mLocationID
RsTxn_SCMDO!WHRNo = TxtWHRNo
RsTxn_SCMDO!Flag = Left(CmbStatus.Text, 1)
RsTxn_SCMDO!Remarks = Trim(TxtRemarks.Text)
RsTxn_SCMDO!ApprovalRemarks = Trim(TxtApprovalRemarks.Text)
RsTxn_SCMDO!DOBags = Val(TxtDONoOfBags.Text)
RsTxn_SCMDO!DOQty = Val(TxtDOQty.Text)

RsTxn_SCMDO!PrevDOBags = Val(TxtPrevDOBags.Text)
RsTxn_SCMDO!PrevDOQty = Val(TxtPrevDOQty.Text)
RsTxn_SCMDO!BalanceBags = Val(TxtBalanceBags.Text)
RsTxn_SCMDO!BalanceQty = Val(TxtBalanceQty.Text)

If IsNull(RsTxn_SCMDO!CreatedBy) = True Or RsTxn_SCMDO!CreatedBy = "" Then
   RsTxn_SCMDO!CreatedBy = Gen_UserLoginID
   RsTxn_SCMDO!CreatedDate = Now
Else
   RsTxn_SCMDO!UpdatedBy = Gen_UserLoginID
   RsTxn_SCMDO!UpdatedDate = Now
End If

RsTxn_SCMDO.Update

SaveCmd.Enabled = False

Call Indata

End Sub
Private Function GetMaxDOID() As Double
Dim Retval As Double
tmp = "Select max(DOID) from Txn_SCMDO"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetMaxDOID = Retval
End Function

Private Sub SearchCmd_Click()

Dim x As Integer
If SearchCmd.Caption = "Cancel" Then
   If RsTxn_SCMDO.RecordCount = 0 Then Exit Sub

   Call GButtonLock(Me, True)
   If RsTxn_SCMDO.EOF Then
      RsTxn_SCMDO.MoveFirst
   End If
   Call Indata
   Exit Sub
End If


Frame0.Height = Frame4.Height - 200
Frame0.Left = Frame1.Left
Frame0.Top = Frame1.Top
Frame0.Width = Frame1.Width
Frame0.Visible = True
'MSHFlexGrid1.Clear
MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 2000


End Sub

Private Sub TxtDONoOfBags_LostFocus()
If SaveCmd.Enabled = False Then Exit Sub
If TxtDONoOfBags.Locked = True Then Exit Sub
If TxtDONoOfBags = "" Then Exit Sub
If Not IsNumeric(TxtDONoOfBags) Or Val(TxtDONoOfBags) < 0 Then
   MsgBox "Invalid numeric format !!!! "
   TxtDONoOfBags.SetFocus
End If

If Val(TxtDONoOfBags) > Val(TxtPledgeBags) - Val(TxtPrevDOBags) Then
   MsgBox "DO Bags should be less than or equal to balance no of bags!!!" 'diffrence of Total of Pledge Bags and Total of Previous DO No of Bags!!!"
   TxtDONoOfBags.SetFocus
   Exit Sub
End If


End Sub

Private Sub TxtDOQty_LostFocus()
If SaveCmd.Enabled = False Then Exit Sub
If TxtDOQty.Locked = True Then Exit Sub
If TxtDOQty = "" Then Exit Sub

If Not IsNumeric(TxtDOQty) Or Val(TxtDOQty) < 0 Then
   MsgBox "Invalid numeric format !!!! "
   TxtDOQty.SetFocus
End If

If Val(Format(Val(TxtDOQty), "#######0.000")) > Val(Format(Val(TxtPledgeQty) - Val(TxtPrevDOQty), "#######0.000")) Then
   MsgBox "DO Qty should be less than or equal to balance qty!!!" ' diffrence of Total of Pledge Qty and Total of Previous DO Qty!!!"
   TxtDOQty.SetFocus
   Exit Sub
End If

End Sub

Private Sub TxtWHRNo_LostFocus()

If SaveCmd.Enabled = False Then Exit Sub
If TxtWHRNoNew.Text = "" Then
   Call ClearDetail
   Exit Sub
End If
Dim strSplitString() As String
If TxtWHRNoNew.Text = "" Then
   Call ClearDetail
   Exit Sub
End If
strSplitString = Split(TxtWHRNoNew.Text, "-")
tmp = UBound(strSplitString)
If tmp <> 1 Then
   MsgBox "Invalid WHR Number!!!"
   TxtWHRNoNew.SetFocus
   Exit Sub
End If
TxtWHRNo = (Val(strSplitString(0)) * 1000000) + Val(strSplitString(1))
   
'Call TableTxn_WHRDetail(" SCMFLag = '1' And WHRNo = " & Val(TxtWHRNo) & "  ")

tmp = " SELECT  TWD.WHRNo,TWD.SM_Prefix,CMP.CMPName,CM.ClientName,ComM.Commodity,TWD.WHRDate,"
tmp = tmp & " TWD.ClientIDRow , TWD.CommodityID, TWD.CMPID, TWD.WHRNoofBags, TWD.WHRQuantity"
tmp = tmp & " ,TWD.PledgeID,TWD.GodownID,TWD.DeliveredQuantity,TWD.DeliveredNoOfBags,"
tmp = tmp & " TWD.SCMFlag , Sum(TPCG.PledgeNoOfBags) 'PledgeBags',Sum(TPCG.PledgeQuantity) 'PledgeQty',LM.LocationID,GM.GodownNo,LM.LocationName "
tmp = tmp & " FROM  Txn_WHR_Detail TWD"
tmp = tmp & " Inner Join Mst_Location LM ON TWD.SM_Prefix=LM.SM_Prefix"
tmp = tmp & " Inner Join Mst_CMP CMP ON TWD.CMPID=CMP.CMPID"
tmp = tmp & " Inner Join Mst_Godown GM ON TWD.GodownID=GM.GodownID"
tmp = tmp & " Inner Join Mst_Client CM ON TWD.ClientIDRow=CM.ClientIDRow"
tmp = tmp & " Inner Join Mst_Commodity ComM ON TWD.CommodityID=ComM.CommodityID"
tmp = tmp & " left join Txn_PledgeCreation TPC on  TWD.PledgeID = TPC.PledgeID And TWD.SM_Prefix = TPC.SM_Prefix And TPC.ExchangeTypeID = 1"
tmp = tmp & " left join Txn_PledgeCreation_GSL TPCG on  TWD.PledgeID = TPCG.PledgeID And TWD.SM_Prefix = TPCG.SM_Prefix"
If Gen_WcLocation <> "" Then
   tmp = tmp & " Where LM." & Gen_WcLocation & "  And TWD.SCMFlag = '1' And TWD.WHRNo = " & Val(TxtWHRNo) & ""
Else
   tmp = tmp & " Where TWD.SCMFlag = '1' And TWD.WHRNo = " & Val(TxtWHRNo) & ""
End If

tmp = tmp & " Group By TWD.WHRNo,TWD.SM_Prefix,CMP.CMPName,CM.ClientName,ComM.Commodity,TWD.WHRDate,"
tmp = tmp & " TWD.ClientIDRow , TWD.CommodityID, TWD.CMPID, TWD.WHRNoofBags, TWD.WHRQuantity,"
tmp = tmp & " TWD.PledgeID,TWD.GodownID,TWD.DeliveredQuantity,TWD.DeliveredNoOfBags,TWD.SCMFlag,LM.LocationID,GM.GodownNo,LM.LocationName"
tmp = tmp & " Having Sum(TPCG.PledgeNoOfBags) > 0 "

Call TmpTable

If TmpRs.RecordCount > 0 Then TmpRs.MoveFirst
TmpRs.Find "WHRNo = " & TxtWHRNo.Text & ""

If TmpRs.EOF Then
   MsgBox "Invalid selection!!!"
   TxtWHRNoNew.SetFocus
   Exit Sub
End If

If TmpRs!DeliveredNoOfBags = TmpRs!WHRNoOfBags Then
   MsgBox "WHR already withdrawn!!!"
   TxtWHRNoNew.SetFocus
   Exit Sub
End If
 
mWHRNo = TmpRs!WHRNo
TxtWHRDate = Format(TmpRs!WHRDate, "dd-mmm-yyyy")
TxtClientName = TmpRs!ClientName
mCommodityID = IIf(IsNull(TmpRs!CommodityID), "", TmpRs!CommodityID)
TxtCommodity = IIf(IsNull(TmpRs!Commodity), "", TmpRs!Commodity)
TxtWHRNoOfBags = IIf(IsNull(TmpRs!WHRNoOfBags), "", TmpRs!WHRNoOfBags)
TxtWHRQuantity = IIf(IsNull(TmpRs!WHRQuantity), "", TmpRs!WHRQuantity)
mSM_Prefix = IIf(IsNull(TmpRs!SM_Prefix), "", TmpRs!SM_Prefix)
TxtPledgeBags = IIf(IsNull(TmpRs!PledgeBags), "", TmpRs!PledgeBags)
TxtPledgeQty = IIf(IsNull(TmpRs!PledgeQty), "", TmpRs!PledgeQty)
TxtWHRQuantity = Format(TxtWHRQuantity, "#########0.000")
mCMPID = IIf(IsNull(TmpRs!CMPID), "", TmpRs!CMPID)
TxtCMPName = IIf(IsNull(TmpRs!CMPName), "", TmpRs!CMPName)
mGodownID = IIf(IsNull(TmpRs!GodownID), "", TmpRs!GodownID)
TxtGodownNo = IIf(IsNull(TmpRs!GodownNo), "", TmpRs!GodownNo)
mLocationID = IIf(IsNull(TmpRs!LocationID), "", TmpRs!LocationID)
TxtLocation = IIf(IsNull(TmpRs!LocationName), "", TmpRs!LocationName)
TxtPledgeID = IIf(IsNull(TmpRs!PledgeID), "", TmpRs!PledgeID)


tmp = "Select TR.PledgeID,TR.SM_Prefix,Sum(TRG.ReleasedNoOfBags) 'RBags',Sum(TRG.ReleasedQuantity) 'RQty' "
tmp = tmp & " From Txn_PledgeRelease  TR"
tmp = tmp & " Inner Join Txn_PledgeRelease_GSL TRG On TR.PledgeReleaseID = TRG.PledgeReleaseID And TR.PledgeID = TRG.PledgeID And TR.SM_Prefix = TRG.SM_Prefix"
tmp = tmp & " Where TR.PledgeID = " & Val(TxtPledgeID) & " ANd TR.SM_Prefix = '" & mSM_Prefix & "'"
tmp = tmp & " Group By TR.PledgeID,TR.SM_Prefix "

Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   TxtPledgeBags = Val(TxtPledgeBags) + IIf(IsNull(TmpRs1!RBags), 0, TmpRs1!RBags)
   TxtPledgeQty = Val(TxtPledgeQty) + IIf(IsNull(TmpRs1!RQty), 0, TmpRs1!RQty)
End If

TxtPrevDOBags = 0
TxtPrevDOQty = 0

tmp = "SELECT  TWD.WHRNo,TWD.SM_Prefix,Sum(TSCM.DOBags) 'PrevDOBags' ,Sum(TSCM.DOQty) 'PrevDOQty'"
tmp = tmp & " FROM  Txn_SCMDO TSCM"
tmp = tmp & " Inner Join Txn_WHR_Detail TWD On TSCM.WHRNo = TWD.WHRNo"
tmp = tmp & " Where TWD.SCMFlag = '1' And TWD.WHRNo = " & TxtWHRNo & " And TSCM.Flag not in ('R')"
tmp = tmp & " Group By TWD.WHRNo,TWD.SM_Prefix"
Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   TxtPrevDOBags = IIf(IsNull(TmpRs1!PrevDOBags), 0, TmpRs1!PrevDOBags)
   TxtPrevDOQty = IIf(IsNull(TmpRs1!PrevDOQty), 0, TmpRs1!PrevDOQty)
End If

TxtBalanceBags = Val(TxtPledgeBags) - Val(TxtPrevDOBags)
TxtBalanceQty = Round(Val(TxtPledgeQty) - Val(TxtPrevDOQty), 3)

'TxtDONoOfBags = 0
'TxtDOQty = 0


End Sub
Private Sub ClearDetail()

TxtLocation = ""
TxtCMPName = ""
TxtClientName = ""
TxtCommodity = ""
TxtGodownNo = ""
TxtWHRNoOfBags = ""
TxtWHRQuantity = ""
TxtPledgeID = ""
TxtPledgeBags = ""
TxtPledgeQty = ""
TxtPrevDOBags = ""
TxtPrevDOQty = ""
CmbStatus.Text = "Approved"
TxtStatus = "Approved"
TxtWHRDate = ""
TxtDONoOfBags = ""
TxtDOQty = ""
TxtBalanceBags = ""
TxtBalanceQty = ""

'TxtRemarks = ""
'TxtApprovalRemarks = ""
'TxtCreated = ""
'TxtUpdated = ""
'TxtApproved = ""


End Sub


Private Sub TxtWHRNoNew_LostFocus()

If SaveCmd.Enabled = False Then Exit Sub
If TxtWHRNoNew.Locked = True Then Exit Sub

Call TxtWHRNo_LostFocus

End Sub

Public Sub Indata()

Frame1.Enabled = False

Call ClearDetail

tmp = " SELECT  TSCM.DOID,TSCM.DODate,TWD.WHRNo,TWD.SM_Prefix,CMP.CMPName,CM.ClientName,ComM.Commodity,TWD.WHRDate,"
tmp = tmp & " TWD.ClientIDRow , TWD.CommodityID, TWD.CMPID, TWD.WHRNoofBags, TWD.WHRQuantity"
tmp = tmp & " ,TWD.PledgeID,TWD.GodownID,TWD.CMPID," 'TWD.DeliveredQuantity,TWD.DeliveredNoOfBags,"
tmp = tmp & " TWD.SCMFlag, TSCM.DOBags,TSCM.DOQty,TSCM.Remarks,TSCM.ApprovalRemarks,TSCM.CreatedBy,TSCM.CreatedDate," ', Sum(TPCG.PledgeNoOfBags) 'PledgeBags',Sum(TPCG.PledgeQuantity) 'PledgeQty'"
tmp = tmp & " TSCM.UpdatedBy,TSCM.UpdatedDate,LM.SM_Prefix,LM.LocationName,LM.LocationID,TWD.DeliveredNoOfBags,TWD.DeliveredQuantity,GM.GodownNo,TSCM.Flag,TSCM.PrevDOBags,TSCM.PrevDOQty,TSCM.BalanceBags,TSCM.BalanceQty "
tmp = tmp & " FROM  Txn_SCMDO TSCM "
tmp = tmp & " Inner Join Txn_WHR_Detail TWD On TSCM.WHRNo = TWD.WHRNo"
tmp = tmp & " Inner Join Mst_Location LM ON TSCM.LocationID=LM.LocationID"
tmp = tmp & " Inner Join Mst_CMP CMP ON TWD.CMPID=CMP.CMPID"
tmp = tmp & " Inner Join Mst_Godown GM ON TWD.GodownID=GM.GodownID"
tmp = tmp & " Inner Join Mst_Client CM ON TWD.ClientIDRow=CM.ClientIDRow"
tmp = tmp & " Inner Join Mst_Commodity ComM ON TWD.CommodityID=ComM.CommodityID"

If Gen_WcLocation <> "" Then
   tmp = tmp & " Where LM." & Gen_WcLocation & "  And TWD.SCMFlag = '1' And TWD.WHRNo = " & Val(RsTxn_SCMDO!WHRNo) & " And TSCM.DOID = " & Val(RsTxn_SCMDO!DOID) & " "
Else
   tmp = tmp & " Where TWD.SCMFlag = '1' And TWD.WHRNo = " & Val(RsTxn_SCMDO!WHRNo) & " And TSCM.DOID = " & Val(RsTxn_SCMDO!DOID) & ""
End If
'tmp = tmp & " Group By TWD.WHRNo,TWD.SM_Prefix,CMP.CMPName,CM.ClientName,ComM.Commodity,TWD.WHRDate,"
'tmp = tmp & " TWD.ClientIDRow , TWD.CommodityID, TWD.CMPID, TWD.WHRNoofBags, TWD.WHRQuantity,"
'tmp = tmp & " TWD.PledgeID,TWD.GodownID,TWD.SCMFlag"
Call TmpTable


'TxtWHRNoNew = RsTxn_WHRDetail!CMPID & "-" & (Val(TxtWHRNo) - (RsTxn_WHRDetail!CMPID * 1000000))
TxtDOID.Text = TmpRs!DOID
TxtDODate.Value = TmpRs!DODate
mLocationID = TmpRs!LocationID
TxtWHRNo = TmpRs!WHRNo
TxtWHRNoNew = TmpRs!CMPID & "-" & (Val(TxtWHRNo) - (TmpRs!CMPID * 1000000))


mWHRNo = TmpRs!WHRNo
TxtWHRDate = Format(TmpRs!WHRDate, "dd-mmm-yyyy")
TxtClientName = TmpRs!ClientName
mCommodityID = IIf(IsNull(TmpRs!CommodityID), "", TmpRs!CommodityID)
TxtCommodity = IIf(IsNull(TmpRs!Commodity), "", TmpRs!Commodity)
TxtWHRNoOfBags = IIf(IsNull(TmpRs!WHRNoOfBags), "", TmpRs!WHRNoOfBags)
TxtWHRQuantity = IIf(IsNull(TmpRs!WHRQuantity), "", TmpRs!WHRQuantity)
mSM_Prefix = IIf(IsNull(TmpRs!SM_Prefix), "", TmpRs!SM_Prefix)

'TxtBalanceBags = TxtWHRNoOfBags - IIf(IsNull(TmpRs!DeliveredNoOfBags), 0, TmpRs!DeliveredNoOfBags)
'TxtBalanceQty = TxtWHRQuantity - IIf(IsNull(TmpRs!DeliveredQuantity), 0, TmpRs!DeliveredQuantity)
TxtPledgeBags = IIf(IsNull(TmpRs!WHRNoOfBags), "", TmpRs!WHRNoOfBags)
TxtPledgeQty = IIf(IsNull(TmpRs!WHRQuantity), "", TmpRs!WHRQuantity)


TxtWHRQuantity = Format(TxtWHRQuantity, "#########0.000")
mCMPID = IIf(IsNull(TmpRs!CMPID), "", TmpRs!CMPID)
TxtCMPName = IIf(IsNull(TmpRs!CMPName), "", TmpRs!CMPName)
mGodownID = IIf(IsNull(TmpRs!GodownID), "", TmpRs!GodownID)
TxtGodownNo = IIf(IsNull(TmpRs!GodownNo), "", TmpRs!GodownNo)
TxtLocation = IIf(IsNull(TmpRs!LocationName), "", TmpRs!LocationName)


CmbStatus.Visible = False
TxtStatus.Visible = True
CmbStatus.Text = TxtStatus

mFlag = IIf(IsNull(TmpRs!Flag), "P", TmpRs!Flag)
If LCase(mFlag) = "p" Then
   TxtStatus = "Pending"
ElseIf LCase(mFlag) = "a" Then
   TxtStatus = "Approved"
ElseIf LCase(mFlag) = "r" Then
   TxtStatus = "Rejected"
End If
CmbStatus.Text = TxtStatus

TxtWHRNoOfBags = IIf(IsNull(TmpRs!WHRNoOfBags), "", TmpRs!WHRNoOfBags)
TxtWHRQuantity = IIf(IsNull(TmpRs!WHRQuantity), "", TmpRs!WHRQuantity)

TxtPledgeID = IIf(IsNull(TmpRs!PledgeID), "", TmpRs!PledgeID)

TxtRemarks.Text = IIf(IsNull(TmpRs!Remarks), "", TmpRs!Remarks)
TxtApprovalRemarks.Text = IIf(IsNull(TmpRs!ApprovalRemarks), "", TmpRs!ApprovalRemarks)


TxtDONoOfBags = IIf(IsNull(TmpRs!DOBags), 0, TmpRs!DOBags)
TxtDOQty = IIf(IsNull(TmpRs!DOQty), 0, TmpRs!DOQty)

TxtPrevDOBags = IIf(IsNull(TmpRs!PrevDOBags), 0, TmpRs!PrevDOBags)
TxtPrevDOQty = IIf(IsNull(TmpRs!PrevDOQty), 0, TmpRs!PrevDOQty)

TxtBalanceBags = IIf(IsNull(TmpRs!BalanceBags), 0, TmpRs!BalanceBags)
TxtBalanceQty = IIf(IsNull(TmpRs!BalanceQty), 0, TmpRs!BalanceQty)

TxtCreated = IIf(IsNull(TmpRs!CreatedBy), "", TmpRs!CreatedBy) & " :- " & IIf(IsNull(TmpRs!CreatedDate), "", TmpRs!CreatedDate)
TxtUpdated = IIf(IsNull(TmpRs!UpdatedBy), "", TmpRs!UpdatedBy) & " :- " & IIf(IsNull(TmpRs!UpdatedDate), "", TmpRs!UpdatedDate)

If RsTxn_SCMDO.RecordCount = 1 Then
   Call GButtonLock_1(Me, True)
Else
   Call GButtonLock(Me, True)
End If
End Sub


Private Sub CmbSLocation_GotFocus()
tmp = CmbSLocation.Text
Call TableMst_Location
CmbSLocation.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
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
   CmbSCMPName.Text = ""
   CmbSCommodity.Text = ""
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbSLocation.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection "
   CmbSLocation.SetFocus
   Exit Sub
End If

If tmp <> CmbSLocation.Text Then
   CmbSCMPName.Text = ""
   CmbSCommodity.Text = ""
End If

mSLocationID = RsMst_Location!LocationID
End Sub


Private Sub CmbSCMPName_GotFocus()
If CmbSLocation.Text = "" Then
   MsgBox "Select Location!!!"
   CmbSLocation.SetFocus
   Exit Sub
End If

tmp = CmbSCMPName.Text
Call TableMst_CMP(" Where CloseDate is null And LocationID = " & mSLocationID)
CmbSCMPName.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_CMP.RecordCount
    CmbSCMPName.AddItem RsMst_CMP!CMPName
    RsMst_CMP.MoveNext
Next
CmbSCMPName.Text = tmp

End Sub

Private Sub CmbSCMPName_LostFocus()
If CmbSCMPName.Text = "" Then Exit Sub

RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPName = '" & CmbSCMPName.Text & "'"
If RsMst_CMP.EOF Then
   MsgBox "Invalid selection "
   CmbSCMPName.SetFocus
   Exit Sub
End If

mSCMPID = RsMst_CMP!CMPID

End Sub


Private Sub CmbSCommodity_GotFocus()

'If Trim(CmbSLocation.Text) = "" Then
'   MsgBox "First Select Location!!!"
'   CmbSLocation.SetFocus
'   Exit Sub
'End If

tmp1 = CmbSCommodity.Text

tmp = "Select MCom.CommodityID, MCom.Commodity From Mst_Commodity MCom "
'tmp = tmp & " Inner Join Mst_LocationCOmmodityMapping MLCM on MLCM.CommodityID = MCom.CommodityID "
'tmp = tmp & " Where MLCM.LocationID = " & mSLocationID
Call Tmp7Table

CmbSCommodity.Clear
If TmpRs7.RecordCount = 0 Then
   MsgBox "No Record found!!!"
   Exit Sub
End If

For I = 1 To TmpRs7.RecordCount
   CmbSCommodity.AddItem TmpRs7!Commodity
   TmpRs7.MoveNext
Next
CmbSCommodity.Text = tmp1

End Sub

Private Sub CmbSCommodity_LostFocus()
If Trim(CmbSCommodity.Text) = "" Then Exit Sub

TmpRs7.MoveFirst
TmpRs7.Find "Commodity = '" & CmbSCommodity.Text & "'"
If TmpRs7.EOF Then
   MsgBox "Invalid selection!!!"
   CmbSCommodity.SetFocus
   Exit Sub
End If

mSCommodityID = TmpRs7!CommodityID

End Sub

Private Sub CmbSClientName_GotFocus()
tmp = CmbSClientName.Text
Call TableMst_Client("where ExchangeTypeId = 1 And SCMFlag = '1' ")
CmbSClientName.Clear
If RsMst_Client.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Client.RecordCount
    CmbSClientName.AddItem RsMst_Client!ClientName
    RsMst_Client.MoveNext
Next
CmbSClientName.Text = tmp
End Sub
Private Sub CmbSClientName_LostFocus()
If CmbSClientName.Text = "" Then
   Exit Sub
End If
RsMst_Client.MoveFirst
RsMst_Client.Find "ClientName = '" & CmbSClientName.Text & "'"
If RsMst_Client.EOF Then
   MsgBox "Invalid selection "
   CmbSClientName.SetFocus
   Exit Sub
End If
End Sub

Private Sub Grid_Head()
CmdExcel.Enabled = False
With MSHFlexGrid1
   .Clear
   .Rows = 2
   .Cols = 12
   .TextMatrix(0, 0) = "DO ID"
   .TextMatrix(0, 1) = "DO Date"
   .TextMatrix(0, 2) = "Location Name"
   .TextMatrix(0, 3) = "CMP"
   .TextMatrix(0, 4) = "Commodity"
   .TextMatrix(0, 5) = "Client Name"
   .TextMatrix(0, 6) = "WHR No"
   .TextMatrix(0, 7) = "DO No Of Bags"
   .TextMatrix(0, 8) = "DO Qty"
   .TextMatrix(0, 9) = "Status"
   .TextMatrix(0, 10) = "Remarks"
   .TextMatrix(0, 11) = "Approval Remarks"
   
   .RowHeight(0) = 400
End With
End Sub
