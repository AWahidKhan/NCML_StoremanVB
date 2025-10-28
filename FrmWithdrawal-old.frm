VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmWithdrawal 
   Caption         =   "Withdrawal Transaction Detail"
   ClientHeight    =   9855
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14985
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9855
   ScaleWidth      =   14985
   WindowState     =   2  'Maximized
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
      TabIndex        =   119
      Top             =   4080
      Visible         =   0   'False
      Width           =   4695
      Begin VB.CommandButton CmdCancel 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cancel"
         Height          =   400
         Left            =   2880
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   122
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton CmdSaveHologram 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Save"
         Height          =   400
         Left            =   1440
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   121
         Top             =   840
         Width           =   1440
      End
      Begin VB.TextBox TxtHologramNo 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1440
         MaxLength       =   25
         TabIndex        =   120
         Top             =   360
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
         TabIndex        =   123
         Top             =   390
         Width           =   1260
      End
   End
   Begin VB.Frame Frame0 
      Height          =   645
      Left            =   3360
      TabIndex        =   104
      Top             =   360
      Visible         =   0   'False
      Width           =   1575
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   350
         Left            =   80
         TabIndex        =   105
         Top             =   150
         Width           =   1000
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1455
         Left            =   70
         TabIndex        =   106
         Top             =   600
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   2566
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.PictureBox Picture3 
      Height          =   1005
      Left            =   120
      ScaleHeight     =   945
      ScaleWidth      =   14880
      TabIndex        =   102
      Top             =   8760
      Width           =   14940
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New "
         Height          =   375
         Left            =   105
         TabIndex        =   31
         Top             =   75
         Width           =   2415
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   375
         Left            =   2520
         TabIndex        =   28
         Top             =   75
         Width           =   2415
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   4935
         TabIndex        =   32
         Top             =   75
         Width           =   2415
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   7350
         TabIndex        =   33
         Top             =   75
         Width           =   2415
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "&Search"
         Height          =   375
         Left            =   9750
         TabIndex        =   29
         Top             =   75
         Width           =   2430
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   375
         Left            =   105
         TabIndex        =   34
         Top             =   450
         Width           =   2415
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   375
         Left            =   2520
         TabIndex        =   36
         Top             =   450
         Width           =   2415
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   375
         Left            =   4935
         TabIndex        =   37
         Top             =   450
         Width           =   2415
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   375
         Left            =   7350
         TabIndex        =   38
         Top             =   450
         Width           =   2415
      End
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   375
         Left            =   9765
         TabIndex        =   30
         Top             =   450
         Width           =   2415
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Print"
         Height          =   750
         Left            =   12180
         TabIndex        =   103
         Top             =   75
         Width           =   2655
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   8595
      Left            =   120
      ScaleHeight     =   8535
      ScaleWidth      =   14850
      TabIndex        =   0
      Top             =   120
      Width           =   14910
      Begin TabDlg.SSTab SSTab1 
         Height          =   6420
         Left            =   120
         TabIndex        =   43
         Top             =   2040
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   11324
         _Version        =   393216
         Tabs            =   4
         Tab             =   3
         TabsPerRow      =   4
         TabHeight       =   520
         ForeColor       =   128
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "RWC Detail"
         TabPicture(0)   =   "FrmWithdrawal.frx":0000
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "PicRRN"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Truck Details"
         TabPicture(1)   =   "FrmWithdrawal.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Pic_TTD"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "GSL Details"
         TabPicture(2)   =   "FrmWithdrawal.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Pic_GCTD"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "Weight Adjustment guide"
         TabPicture(3)   =   "FrmWithdrawal.frx":0054
         Tab(3).ControlEnabled=   -1  'True
         Tab(3).Control(0)=   "MSHFlexGrid5"
         Tab(3).Control(0).Enabled=   0   'False
         Tab(3).ControlCount=   1
         Begin VB.PictureBox Pic_TTD 
            Height          =   5535
            Left            =   -74880
            ScaleHeight     =   5475
            ScaleWidth      =   14475
            TabIndex        =   79
            Top             =   600
            Width           =   14535
            Begin VB.TextBox TxtTTDGrossWt 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
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
               Left            =   8400
               TabIndex        =   16
               Top             =   420
               Width           =   1815
            End
            Begin VB.TextBox TxtTTDTruckNo 
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
               Left            =   120
               TabIndex        =   12
               Top             =   420
               Width           =   2415
            End
            Begin VB.TextBox TxtWeighBridgNo 
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
               Left            =   2520
               TabIndex        =   13
               Top             =   420
               Width           =   2535
            End
            Begin VB.TextBox TxtTTDTotalWt 
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
               Left            =   5040
               TabIndex        =   14
               Top             =   420
               Width           =   1575
            End
            Begin VB.TextBox TxtTTDTareWt 
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
               Left            =   6600
               TabIndex        =   15
               Top             =   420
               Width           =   1815
            End
            Begin VB.CommandButton CmdTTDAdd2List 
               Caption         =   "Add to List"
               Height          =   375
               Left            =   12600
               TabIndex        =   18
               Top             =   960
               Width           =   1575
            End
            Begin VB.TextBox txtTTDRemarks 
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
               Left            =   10215
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   17
               Top             =   420
               Width           =   4200
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
               Height          =   3930
               Left            =   120
               TabIndex        =   80
               Top             =   1440
               Width           =   14055
               _ExtentX        =   24791
               _ExtentY        =   6932
               _Version        =   393216
               WordWrap        =   -1  'True
               AllowUserResizing=   3
               _NumberOfBands  =   1
               _Band(0).Cols   =   2
            End
            Begin VB.Label Label18 
               Caption         =   "Gross Weight"
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
               Left            =   8520
               TabIndex        =   86
               Top             =   120
               Width           =   1455
            End
            Begin VB.Label Label30 
               Caption         =   "Truck No."
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
               Left            =   600
               TabIndex        =   85
               Top             =   120
               Width           =   1095
            End
            Begin VB.Label Label15 
               Caption         =   "Weighbridge Slip No."
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
               Left            =   2880
               TabIndex        =   84
               Top             =   120
               Width           =   1935
            End
            Begin VB.Label Label16 
               Caption         =   "Total Weight"
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
               Left            =   5160
               TabIndex        =   83
               Top             =   120
               Width           =   1335
            End
            Begin VB.Label Label17 
               Caption         =   "Tare Weight"
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
               Left            =   6840
               TabIndex        =   82
               Top             =   120
               Width           =   1335
            End
            Begin VB.Label Label6 
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
               Height          =   255
               Left            =   10920
               TabIndex        =   81
               Top             =   120
               Width           =   855
            End
         End
         Begin VB.PictureBox Pic_GCTD 
            Height          =   5775
            Left            =   -74880
            ScaleHeight     =   5715
            ScaleWidth      =   14475
            TabIndex        =   59
            Top             =   480
            Width           =   14535
            Begin VB.TextBox TxtGSLAvailableNofBags 
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
               Left            =   3720
               Locked          =   -1  'True
               TabIndex        =   69
               TabStop         =   0   'False
               Top             =   540
               Width           =   1800
            End
            Begin VB.TextBox TxtGCTDNoofBags 
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
               Left            =   5520
               TabIndex        =   22
               Top             =   540
               Width           =   2295
            End
            Begin VB.TextBox TxtGSLPledgedWeight 
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
               Left            =   1290
               Locked          =   -1  'True
               TabIndex        =   115
               TabStop         =   0   'False
               Top             =   1440
               Width           =   1185
            End
            Begin VB.TextBox TxtGSLPledgedNofBags 
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
               Left            =   1290
               Locked          =   -1  'True
               TabIndex        =   113
               TabStop         =   0   'False
               Top             =   1020
               Width           =   1185
            End
            Begin VB.ComboBox CmbGCTDClient 
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
               ItemData        =   "FrmWithdrawal.frx":0070
               Left            =   5400
               List            =   "FrmWithdrawal.frx":0072
               TabIndex        =   26
               Top             =   1380
               Visible         =   0   'False
               Width           =   4095
            End
            Begin VB.TextBox TxtAdjWt 
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
               Left            =   9270
               TabIndex        =   24
               Top             =   540
               Width           =   1560
            End
            Begin VB.CheckBox Check2 
               Caption         =   "Spillage"
               Height          =   255
               Left            =   2730
               TabIndex        =   107
               Top             =   1080
               Width           =   1170
            End
            Begin VB.ComboBox CmbLot 
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
               ItemData        =   "FrmWithdrawal.frx":0074
               Left            =   2480
               List            =   "FrmWithdrawal.frx":0076
               Sorted          =   -1  'True
               TabIndex        =   21
               Top             =   540
               Width           =   1230
            End
            Begin VB.ComboBox CmbStack 
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
               ItemData        =   "FrmWithdrawal.frx":0078
               Left            =   1290
               List            =   "FrmWithdrawal.frx":007A
               Sorted          =   -1  'True
               TabIndex        =   20
               Top             =   540
               Width           =   1185
            End
            Begin VB.ComboBox CmbGodownNo 
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
               ItemData        =   "FrmWithdrawal.frx":007C
               Left            =   0
               List            =   "FrmWithdrawal.frx":007E
               Sorted          =   -1  'True
               TabIndex        =   19
               Top             =   540
               Width           =   1290
            End
            Begin VB.TextBox TxtGSLISINno 
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
               Left            =   5400
               Locked          =   -1  'True
               TabIndex        =   77
               TabStop         =   0   'False
               Top             =   1380
               Width           =   4575
            End
            Begin VB.TextBox TxtGSLWeightGainLoss 
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
               Left            =   13170
               Locked          =   -1  'True
               TabIndex        =   75
               TabStop         =   0   'False
               Top             =   540
               Visible         =   0   'False
               Width           =   1245
            End
            Begin VB.TextBox TxtGSLNetWeight 
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
               Left            =   11880
               Locked          =   -1  'True
               TabIndex        =   73
               TabStop         =   0   'False
               Top             =   540
               Width           =   1305
            End
            Begin VB.TextBox TxtGSLAvailableGrossWgt 
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
               Left            =   4905
               Locked          =   -1  'True
               TabIndex        =   71
               TabStop         =   0   'False
               Top             =   540
               Visible         =   0   'False
               Width           =   1590
            End
            Begin VB.TextBox TxtGCTDremarks 
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
               Left            =   5400
               MaxLength       =   255
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   25
               Top             =   960
               Width           =   9015
            End
            Begin VB.CommandButton CmdGCTDAdd2List 
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
               Left            =   12795
               TabIndex        =   27
               Top             =   1395
               Width           =   1605
            End
            Begin VB.TextBox TxtGCTDGrossWt 
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
               Left            =   7800
               TabIndex        =   23
               Top             =   540
               Width           =   1500
            End
            Begin VB.TextBox TxtGGunnyWeight 
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
               Left            =   10800
               Locked          =   -1  'True
               TabIndex        =   60
               TabStop         =   0   'False
               Top             =   540
               Width           =   1095
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid4 
               Height          =   3765
               Left            =   120
               TabIndex        =   61
               Top             =   1920
               Width           =   14295
               _ExtentX        =   25215
               _ExtentY        =   6641
               _Version        =   393216
               WordWrap        =   -1  'True
               AllowUserResizing=   3
               _NumberOfBands  =   1
               _Band(0).Cols   =   2
            End
            Begin VB.Label Label11 
               Caption         =   "Pledged Wt."
               Height          =   255
               Left            =   60
               TabIndex        =   116
               Top             =   1500
               Width           =   1155
            End
            Begin VB.Label Label9 
               Caption         =   "Pledged Bags"
               Height          =   255
               Left            =   60
               TabIndex        =   114
               Top             =   1080
               Width           =   1155
            End
            Begin VB.Label LblClient_WHR_RRN 
               Caption         =   "Client Name ~ DP Name"
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
               Left            =   4200
               TabIndex        =   110
               Top             =   1320
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.Label Label1 
               Caption         =   "Adjusted Weight"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   9405
               TabIndex        =   108
               Top             =   45
               Width           =   1095
            End
            Begin VB.Label Label32 
               Caption         =   "ISIN No."
               Height          =   255
               Left            =   4185
               TabIndex        =   78
               Top             =   1440
               Width           =   1080
            End
            Begin VB.Label Label31 
               Caption         =   "Weight  Gain / Loss"
               Height          =   375
               Left            =   13200
               TabIndex        =   76
               Top             =   75
               Visible         =   0   'False
               Width           =   1185
            End
            Begin VB.Label Label26 
               Caption         =   "Electronic Weight"
               Height          =   495
               Left            =   12120
               TabIndex        =   74
               Top             =   0
               Width           =   810
            End
            Begin VB.Label Label13 
               Caption         =   "Available Electronic Weight"
               Height          =   390
               Left            =   4950
               TabIndex        =   72
               Top             =   75
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.Label Label12 
               Caption         =   "Available Bags"
               Height          =   255
               Left            =   3780
               TabIndex        =   70
               Top             =   120
               Width           =   1155
            End
            Begin VB.Label Label25 
               Caption         =   "Remarks"
               Height          =   255
               Left            =   4200
               TabIndex        =   68
               Top             =   1020
               Width           =   975
            End
            Begin VB.Label Label24 
               Caption         =   "Delivered Gross Weight"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Left            =   7860
               TabIndex        =   67
               Top             =   60
               Width           =   1215
            End
            Begin VB.Label Label23 
               Caption         =   "Delivered Bags"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Left            =   6075
               TabIndex        =   66
               Top             =   60
               Width           =   1440
            End
            Begin VB.Label Label22 
               Caption         =   "Lot No"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   2760
               TabIndex        =   65
               Top             =   120
               Width           =   735
            End
            Begin VB.Label Label21 
               Caption         =   "Stack No"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1440
               TabIndex        =   64
               Top             =   120
               Width           =   975
            End
            Begin VB.Label Label20 
               Caption         =   "Godown No"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   63
               Top             =   120
               Width           =   1095
            End
            Begin VB.Label Label27 
               Caption         =   "Gunny Weight"
               Height          =   495
               Left            =   11040
               TabIndex        =   62
               Top             =   0
               Width           =   855
            End
         End
         Begin VB.PictureBox PicRRN 
            Height          =   5655
            Left            =   -74880
            ScaleHeight     =   5595
            ScaleWidth      =   14475
            TabIndex        =   44
            Top             =   600
            Width           =   14535
            Begin VB.TextBox TxtRRNBags 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               Height          =   375
               Left            =   11520
               Locked          =   -1  'True
               TabIndex        =   112
               TabStop         =   0   'False
               Top             =   120
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.TextBox TxtBalanceBags 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               Height          =   375
               Left            =   13005
               Locked          =   -1  'True
               TabIndex        =   111
               TabStop         =   0   'False
               Top             =   120
               Visible         =   0   'False
               Width           =   1440
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
               Left            =   6300
               Sorted          =   -1  'True
               TabIndex        =   10
               Top             =   1080
               Visible         =   0   'False
               Width           =   6105
            End
            Begin VB.TextBox TxtDPName 
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
               Left            =   8025
               Locked          =   -1  'True
               TabIndex        =   98
               TabStop         =   0   'False
               Top             =   1080
               Width           =   4350
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
               Height          =   375
               Left            =   1800
               Locked          =   -1  'True
               TabIndex        =   57
               TabStop         =   0   'False
               Top             =   1080
               Width           =   4470
            End
            Begin VB.TextBox TxtDPID 
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
               Left            =   6285
               Locked          =   -1  'True
               TabIndex        =   55
               TabStop         =   0   'False
               Top             =   1080
               Width           =   1725
            End
            Begin VB.TextBox TxtClientID 
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
               Left            =   30
               Locked          =   -1  'True
               TabIndex        =   9
               TabStop         =   0   'False
               Top             =   1080
               Width           =   1755
            End
            Begin VB.TextBox TxtRRNRemarks 
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
               Height          =   390
               Left            =   7590
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   8
               Top             =   360
               Width           =   6840
            End
            Begin VB.CommandButton CmdAdd2ListRRN 
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
               Left            =   12480
               TabIndex        =   11
               Top             =   1065
               Width           =   1935
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
               Left            =   5760
               Locked          =   -1  'True
               TabIndex        =   47
               TabStop         =   0   'False
               Top             =   360
               Width           =   1800
            End
            Begin VB.TextBox TxtRRNNo 
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
               Left            =   30
               Locked          =   -1  'True
               TabIndex        =   7
               Top             =   360
               Width           =   1830
            End
            Begin VB.TextBox TxtISINno 
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
               TabIndex        =   46
               TabStop         =   0   'False
               Top             =   360
               Width           =   2010
            End
            Begin VB.TextBox TxtRRNQty 
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
               Left            =   3915
               Locked          =   -1  'True
               TabIndex        =   45
               TabStop         =   0   'False
               Top             =   360
               Width           =   1815
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
               Height          =   3930
               Left            =   60
               TabIndex        =   48
               Top             =   1560
               Width           =   14340
               _ExtentX        =   25294
               _ExtentY        =   6932
               _Version        =   393216
               WordWrap        =   -1  'True
               AllowUserResizing=   3
               _NumberOfBands  =   1
               _Band(0).Cols   =   2
            End
            Begin VB.Label lblNameDid 
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
               Left            =   7380
               TabIndex        =   109
               Top             =   795
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.Label LblDPName 
               Caption         =   "DP Name"
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
               Left            =   8640
               TabIndex        =   99
               Top             =   810
               Width           =   3255
            End
            Begin VB.Label Label10 
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
               Left            =   3450
               TabIndex        =   58
               Top             =   795
               Width           =   1095
            End
            Begin VB.Label LblDPID 
               Caption         =   "DP-ID"
               Height          =   255
               Left            =   6285
               TabIndex        =   56
               Top             =   840
               Width           =   1575
            End
            Begin VB.Label Label8 
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
               Height          =   255
               Left            =   600
               TabIndex        =   54
               Top             =   810
               Width           =   855
            End
            Begin VB.Label Label42 
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
               Height          =   255
               Left            =   10485
               TabIndex        =   53
               Top             =   90
               Width           =   975
            End
            Begin VB.Label Label38 
               Caption         =   "Balance Qty"
               Height          =   255
               Left            =   6158
               TabIndex        =   52
               Top             =   75
               Width           =   975
            End
            Begin VB.Label LblRRNQty 
               Caption         =   "RRN Qty"
               Height          =   255
               Left            =   4350
               TabIndex        =   51
               Top             =   75
               Width           =   1545
            End
            Begin VB.Label LblISINNo 
               Caption         =   "ISIN No."
               Height          =   255
               Left            =   2318
               TabIndex        =   50
               Top             =   105
               Width           =   1155
            End
            Begin VB.Label LblRRNNo 
               Alignment       =   2  'Center
               Caption         =   "RRN No"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   30
               TabIndex        =   49
               Top             =   105
               Width           =   1815
            End
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid5 
            Height          =   5445
            Left            =   120
            TabIndex        =   125
            Top             =   720
            Width           =   14415
            _ExtentX        =   25426
            _ExtentY        =   9604
            _Version        =   393216
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin VB.PictureBox Picture2 
         Height          =   1860
         Left            =   120
         ScaleHeight     =   1800
         ScaleWidth      =   14595
         TabIndex        =   40
         Top             =   120
         Width           =   14655
         Begin VB.TextBox TxtManualCDTFNo 
            Alignment       =   1  'Right Justify
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
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   126
            Top             =   1350
            Width           =   1815
         End
         Begin VB.CommandButton CmdAgent 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   11400
            TabIndex        =   124
            TabStop         =   0   'False
            Top             =   960
            Width           =   375
         End
         Begin VB.ComboBox CmbAgent 
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
            Left            =   5280
            Sorted          =   -1  'True
            TabIndex        =   5
            Top             =   960
            Width           =   6015
         End
         Begin VB.TextBox TxtPreFix 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
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
            Left            =   11445
            Locked          =   -1  'True
            TabIndex        =   117
            TabStop         =   0   'False
            Top             =   930
            Visible         =   0   'False
            Width           =   975
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
            ItemData        =   "FrmWithdrawal.frx":0080
            Left            =   1215
            List            =   "FrmWithdrawal.frx":0082
            Sorted          =   -1  'True
            TabIndex        =   1
            Top             =   930
            Width           =   1815
         End
         Begin VB.ComboBox CmbCommodityID 
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
            Left            =   5280
            Sorted          =   -1  'True
            TabIndex        =   4
            Top             =   510
            Width           =   6015
         End
         Begin VB.ComboBox CmbCMPID 
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
            Left            =   5280
            Sorted          =   -1  'True
            TabIndex        =   3
            Top             =   90
            Width           =   6015
         End
         Begin VB.TextBox TxtAgent 
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
            Left            =   5280
            ScrollBars      =   2  'Vertical
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   930
            Width           =   6000
         End
         Begin VB.TextBox TxtCGunnyWt 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
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
            Left            =   13530
            Locked          =   -1  'True
            TabIndex        =   89
            TabStop         =   0   'False
            Top             =   90
            Width           =   975
         End
         Begin VB.TextBox TxtCLotSize 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
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
            Left            =   13530
            Locked          =   -1  'True
            TabIndex        =   88
            TabStop         =   0   'False
            Top             =   510
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.TextBox TxtCVariation 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
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
            Left            =   13530
            Locked          =   -1  'True
            TabIndex        =   87
            TabStop         =   0   'False
            Top             =   930
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Check1"
            Height          =   255
            Left            =   4050
            TabIndex        =   2
            Top             =   945
            Visible         =   0   'False
            Width           =   225
         End
         Begin VB.TextBox TxtCxTFNo 
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
            Left            =   1215
            Locked          =   -1  'True
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   90
            Width           =   1815
         End
         Begin MSComCtl2.DTPicker TxtCxTFDate 
            Height          =   360
            Left            =   1215
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   510
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   635
            _Version        =   393216
            CalendarBackColor=   -2147483648
            Format          =   48627713
            CurrentDate     =   39083
            MinDate         =   39083
         End
         Begin VB.TextBox TxtCMPID 
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
            Left            =   5280
            Locked          =   -1  'True
            TabIndex        =   100
            TabStop         =   0   'False
            Top             =   120
            Width           =   6015
         End
         Begin VB.TextBox TxtCommodityID 
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
            Left            =   5280
            Locked          =   -1  'True
            TabIndex        =   101
            TabStop         =   0   'False
            Top             =   525
            Width           =   6015
         End
         Begin VB.Label Label35 
            Caption         =   "Manual CITF #"
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
            Left            =   120
            TabIndex        =   127
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            Caption         =   "Location ID"
            Height          =   255
            Left            =   11520
            TabIndex        =   118
            Top             =   660
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label Label14 
            Caption         =   "Commodity"
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
            Left            =   4080
            TabIndex        =   97
            Top             =   563
            Width           =   1215
         End
         Begin VB.Label Label41 
            Alignment       =   1  'Right Justify
            Caption         =   "Lot Size"
            Height          =   255
            Left            =   11925
            TabIndex        =   96
            Top             =   570
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label Label5 
            Caption         =   "Txn Date"
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
            TabIndex        =   95
            Top             =   563
            Width           =   1215
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "Rejected "
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
            TabIndex        =   94
            Top             =   930
            Visible         =   0   'False
            Width           =   840
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label28 
            Caption         =   "Agent"
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
            Left            =   4590
            TabIndex        =   93
            Top             =   990
            Width           =   855
         End
         Begin VB.Label Label4 
            Caption         =   "CMP"
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
            Left            =   4680
            TabIndex        =   92
            Top             =   150
            Width           =   615
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Gunny Wt. per Bag"
            Height          =   255
            Left            =   11925
            TabIndex        =   91
            Top             =   150
            Width           =   1455
         End
         Begin VB.Label Label43 
            Alignment       =   1  'Right Justify
            Caption         =   "% Variation "
            Height          =   255
            Left            =   11925
            TabIndex        =   90
            Top             =   990
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label Label7 
            Caption         =   "Mode"
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
            TabIndex        =   42
            Top             =   983
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "CITF No"
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
            TabIndex        =   41
            Top             =   143
            Width           =   1215
         End
      End
   End
End
Attribute VB_Name = "FrmWithdrawal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg, Diff_wt As Variant
Dim mAgentID, mCMPID, mCommodityId, mContractID, mGodownID, mGSLID, mRRNinGCTD As Double
Dim mExchangeTypeID As Double
Dim mClientIDRow As Double
Dim mRWCID As Double
Dim mGSLClientIDRow As Double
Dim mRWCBalanceQty As Double
Dim mPledgeID, mTimeStamp As Variant
Private Sub Grid_Head_Error()

MSHFlexGrid5.Clear
MSHFlexGrid5.Rows = 3
MSHFlexGrid5.FixedRows = 2
MSHFlexGrid5.Cols = 7
MSHFlexGrid5.TextMatrix(1, 0) = ""
If mExchangeTypeID = 1 Then
   MSHFlexGrid5.TextMatrix(1, 1) = "Client Name"
   MSHFlexGrid5.TextMatrix(1, 2) = "WHR Balance Qty"
Else
   MSHFlexGrid5.TextMatrix(1, 1) = "ISIN No"
   MSHFlexGrid5.TextMatrix(1, 2) = "RRN Balance Qty"
End If
MSHFlexGrid5.TextMatrix(1, 3) = "Delivered Qty (Electronic Weight)"
MSHFlexGrid5.TextMatrix(1, 4) = "Adjusted Weight"
MSHFlexGrid5.TextMatrix(1, 5) = "Excess Qty Delivered"
MSHFlexGrid5.TextMatrix(1, 6) = "Partial Qty Delivered"

MSHFlexGrid5.ColWidth(0) = 500
MSHFlexGrid5.ColWidth(1) = 3000
MSHFlexGrid5.ColWidth(2) = 2000
MSHFlexGrid5.ColWidth(3) = 2500
MSHFlexGrid5.ColWidth(4) = 1800
MSHFlexGrid5.ColWidth(5) = 1800
MSHFlexGrid5.ColWidth(6) = 1800


End Sub

Private Sub Grid_Head_RRN()

MSHFlexGrid2.Clear
MSHFlexGrid2.Rows = 3
MSHFlexGrid2.FixedRows = 2
MSHFlexGrid2.Cols = 16

MSHFlexGrid2.TextMatrix(1, 0) = ""
MSHFlexGrid2.TextMatrix(1, 1) = "RRN No"
MSHFlexGrid2.TextMatrix(1, 2) = "ISIN No"
MSHFlexGrid2.TextMatrix(1, 3) = "Client ID"
MSHFlexGrid2.TextMatrix(1, 4) = "Client Name"
MSHFlexGrid2.TextMatrix(1, 5) = "DP ID"
MSHFlexGrid2.TextMatrix(1, 6) = "DP Name"
MSHFlexGrid2.TextMatrix(1, 7) = "RRN Qty"
MSHFlexGrid2.TextMatrix(1, 8) = "Balance Qty (excluding this CITF)"
MSHFlexGrid2.TextMatrix(1, 9) = "Delivered Qty (Electronic Weight)"
MSHFlexGrid2.TextMatrix(1, 10) = "Remark"
MSHFlexGrid2.TextMatrix(1, 11) = "ClientRowID"
MSHFlexGrid2.TextMatrix(1, 12) = "R/W/C ID"
MSHFlexGrid2.TextMatrix(1, 13) = "WHR Bags"
MSHFlexGrid2.TextMatrix(1, 14) = "Balance Bags"
MSHFlexGrid2.TextMatrix(1, 15) = "Delivered Bags"


MSHFlexGrid2.ColWidth(0) = 500
MSHFlexGrid2.ColWidth(1) = 1300
MSHFlexGrid2.ColWidth(2) = 1200
MSHFlexGrid2.ColWidth(3) = 1200
MSHFlexGrid2.ColWidth(4) = 1500
MSHFlexGrid2.ColWidth(5) = 1500
MSHFlexGrid2.ColWidth(6) = 1500
MSHFlexGrid2.ColWidth(10) = 1500
MSHFlexGrid2.ColWidth(11) = 0
MSHFlexGrid2.ColWidth(12) = 0
MSHFlexGrid2.ColWidth(13) = 0
MSHFlexGrid2.ColWidth(14) = 1000
MSHFlexGrid2.ColWidth(15) = 1000

MSHFlexGrid2.RowHeight(1) = 800

End Sub
Private Sub Grid_Head_TTD()
MSHFlexGrid3.Clear
MSHFlexGrid3.Rows = 3
MSHFlexGrid3.FixedRows = 2

MSHFlexGrid3.Cols = 7
MSHFlexGrid3.TextMatrix(1, 0) = ""
MSHFlexGrid3.TextMatrix(1, 1) = "Truck No"
MSHFlexGrid3.TextMatrix(1, 2) = "Weighbridge Slip No"
MSHFlexGrid3.TextMatrix(1, 3) = "Total Weight"
MSHFlexGrid3.TextMatrix(1, 4) = "Tare Weight"
MSHFlexGrid3.TextMatrix(1, 5) = "Gross Weight"
MSHFlexGrid3.TextMatrix(1, 6) = "Remark"

MSHFlexGrid3.ColWidth(0) = 0
MSHFlexGrid3.ColWidth(1) = 2000
MSHFlexGrid3.ColWidth(2) = 2500
MSHFlexGrid3.ColWidth(3) = 1500
MSHFlexGrid3.ColWidth(4) = 1500
MSHFlexGrid3.ColWidth(5) = 1500
MSHFlexGrid3.ColWidth(6) = 4500

End Sub
Private Sub Grid_Head_GCTD()
MSHFlexGrid4.Clear
MSHFlexGrid4.Rows = 3
MSHFlexGrid4.FixedRows = 2

MSHFlexGrid4.Cols = 18
MSHFlexGrid4.TextMatrix(1, 0) = ""
MSHFlexGrid4.TextMatrix(1, 1) = "Godown No"
MSHFlexGrid4.TextMatrix(1, 2) = "Stack No"
MSHFlexGrid4.TextMatrix(1, 3) = "Lot No"
MSHFlexGrid4.TextMatrix(1, 4) = "Available Bags"
MSHFlexGrid4.TextMatrix(1, 5) = "Available Ele. Wt."
MSHFlexGrid4.TextMatrix(1, 6) = "No of Bags"
MSHFlexGrid4.TextMatrix(1, 7) = "Gross Wt.    (A) "
MSHFlexGrid4.TextMatrix(1, 8) = "Adjusted Wt. (B)"
MSHFlexGrid4.TextMatrix(1, 9) = "Gunny Weight (C)"
MSHFlexGrid4.TextMatrix(1, 10) = "Electronic Weight (A+B-C)"
MSHFlexGrid4.TextMatrix(1, 11) = "Weight Gain/Loss"
MSHFlexGrid4.TextMatrix(1, 12) = "ISIN No"
MSHFlexGrid4.TextMatrix(1, 13) = "Remark"
MSHFlexGrid4.TextMatrix(1, 14) = "GSLID"
MSHFlexGrid4.TextMatrix(1, 15) = "Godown Close"
MSHFlexGrid4.TextMatrix(1, 16) = "ClientIDRow"
MSHFlexGrid4.TextMatrix(1, 17) = "SpillageLot"


MSHFlexGrid4.ColWidth(0) = 0
MSHFlexGrid4.ColWidth(1) = 1200
MSHFlexGrid4.ColWidth(2) = 1200
MSHFlexGrid4.ColWidth(3) = 1200

MSHFlexGrid4.ColWidth(4) = 0
MSHFlexGrid4.ColWidth(5) = 0

MSHFlexGrid4.ColWidth(12) = 1500
MSHFlexGrid4.ColWidth(13) = 2000
MSHFlexGrid4.ColWidth(14) = 0
MSHFlexGrid4.ColWidth(15) = 0
MSHFlexGrid4.ColWidth(16) = 0
MSHFlexGrid4.ColWidth(11) = 0
MSHFlexGrid4.RowHeight(1) = 800
End Sub
Private Sub AddCmd_Click()
Edit_Flg = "A"
Call Grid_Head_RRN
Call Grid_Head_Error
Call Grid_Head_TTD
Call Grid_Head_GCTD
Call TableMst_RRN

TxtCxTFDate = CDate("01/01/2009")
TxtCMPID.Visible = False
TxtCommodityID.Visible = False
CmbCMPID.Visible = True
CmbCommodityID.Visible = True
CmbExchangeTypeID.Enabled = True
PicRRN.Enabled = True
Pic_GCTD.Enabled = True
Pic_TTD.Enabled = True
TxtRRNNo.Locked = False
TxtAgent.Locked = False
TxtCxTFNo = ""
TxtAgent = ""
Check1.Value = 0
CmbExchangeTypeID.Text = ""
CmbCMPID.Text = ""
CmbCommodityID.Text = ""
TxtCxTFDate.Enabled = False
Check1.Enabled = True
Check1.Value = 0
If RsTxn_CxTF_Details.RecordCount > 0 Then
   If CmbExchangeTypeID.Enabled = True Then
      CmbExchangeTypeID.SetFocus
   End If
End If
Call CheckRejection(Check1.Value)
Call GButtonLock(Me, False)
CmdPrint.Enabled = False
TxtAgent.Locked = False
CmbAgent.Visible = True
CmdAgent.Visible = True
TxtAgent.Visible = False
CmbAgent.Text = ""
TxtTTDGrossWt.Locked = True

End Sub
Private Sub Check1_Click()
'Check1.Enabled = False
'SSTab1.Tab = 0
'MsgBox SSTab1.Caption

If mExchangeTypeID = 3 Then 'SPOT
   Call CheckRejection(Check1.Value)
   Call SetCaptions("CDC No", "CDC-No", "CDC-No", "CDC Qty", "TM ID", "Trading Member Name", True)
   Call CDCDetails(mCMPID, mCommodityId, TxtRRNNo.Text)
Else
   Call CheckRejection(Check1.Value)
   If Check1.Value = 1 Then
      Call SetCaptions("client details", "CDTF No", "Client Name", "Rejected Qty", "DP ID", "DP Name", True)
   Else
      Call SetCaptions("Remat Request No. (RRN)", "RRN No", "ISIN No", "RRN Qty", "DP ID", "DP Name", True)
   End If
End If
End Sub
Private Sub CheckRejection(Value_)
If mExchangeTypeID = 1 Then Exit Sub

Call Clear_RWCTabControl
Call Grid_Head_RRN

Label8.Top = 810
TxtClientID.Top = 1080
Label42.Top = 90
Label42.Left = 10485
TxtRRNRemarks.Top = 360
TxtRRNRemarks.Left = 7590

Label10.Top = Label8.Top
TxtClientName.Top = TxtClientID.Top
LblDPID.Top = Label8.Top
LblDPName.Top = Label8.Top
lblNameDid.Top = Label8.Top
CmbClientDP.Top = TxtClientID.Top
TxtDPID.Top = TxtClientID.Top
TxtDPName.Top = TxtClientID.Top



LblRRNQty.Top = 75
LblRRNQty.Left = 4470
TxtRRNQty.Top = 360
TxtRRNQty.Left = 3915

TxtBalanceQty.Top = 360
TxtBalanceQty.Left = 5760
Label38.Top = 75
Label38.Left = 6158


LblRRNNo.Visible = True
TxtRRNNo.Visible = True
LblISINNo.Visible = True
TxtISINno.Visible = True
TxtRRNQty.Visible = True
LblRRNQty.Visible = True
TxtBalanceQty.Visible = True
CmbGCTDClient.Visible = False
LblClient_WHR_RRN.Visible = False
TxtGSLISINno.Visible = True
Label38.Visible = True
TxtClientID.Locked = True
TxtClientID.BackColor = &H80000000
MSHFlexGrid2.ColWidth(1) = 1300
MSHFlexGrid2.ColWidth(2) = 1200
MSHFlexGrid2.ColWidth(7) = 1000
MSHFlexGrid2.ColWidth(8) = 1000
MSHFlexGrid2.ColWidth(9) = 1000
MSHFlexGrid2.ColWidth(10) = 1500


If mExchangeTypeID = 2 And Value_ = 1 Then
   SSTab1.TabCaption(0) = "Client Detail "
   LblRRNNo.Visible = False
   TxtRRNNo.Visible = False
   LblISINNo.Visible = False
   TxtISINno.Visible = False
   
   TxtRRNQty.Visible = True
   LblRRNQty.Visible = True
   
   TxtBalanceQty.Visible = True
   Label38.Visible = True
   
   CmbGCTDClient.Visible = True
   LblClient_WHR_RRN.Visible = True
   TxtGSLISINno.Visible = False
   
   TxtClientID.Locked = False
   TxtClientID.BackColor = &H80000004
   MSHFlexGrid2.ColWidth(1) = 0
   MSHFlexGrid2.ColWidth(2) = 0
'   MSHFlexGrid2.ColWidth(7) = 0
'   MSHFlexGrid2.ColWidth(8) = 0
'   MSHFlexGrid2.ColWidth(9) = 0
   MSHFlexGrid2.ColWidth(10) = 1500
   Label8.Top = 105
   TxtClientID.Top = 360
   Label42.Top = 810
   Label42.Left = 3450
   TxtRRNRemarks.Top = 1080
   TxtRRNRemarks.Left = 50
   TxtRRNRemarks.Width = 8500
   
   LblRRNQty.Top = 810
   LblRRNQty.Left = 9000
   TxtRRNQty.Top = 1080
   TxtRRNQty.Left = 8650
   
   Label38.Top = 810
   Label38.Left = 10500
   TxtBalanceQty.Top = 1080
   TxtBalanceQty.Left = 10500
   
   Label10.Top = Label8.Top
   
   TxtClientName.Top = TxtClientID.Top
   LblDPID.Top = Label8.Top
   LblDPName.Top = Label8.Top
   lblNameDid.Top = Label8.Top
   CmbClientDP.Top = TxtClientID.Top
   TxtDPID.Top = TxtClientID.Top
   TxtDPName.Top = TxtClientID.Top
    

ElseIf mExchangeTypeID = 2 And Value_ = 0 Then
       SSTab1.TabCaption(0) = "Remat Request No. (RRN)"
'       LblRRNNo.Visible = True
'       TxtRRNNo.Visible = True
'       LblISINNo.Visible = True
'       TxtISINno.Visible = True
'       TxtRRNQty.Visible = True
'       LblRRNQty.Visible = True
'       TxtBalanceQty.Visible = True
'       CmbGCTDClient.Visible = False
'       LblClient_WHR_RRN.Visible = False
'       TxtGSLISINno.Visible = True
'       Label38.Visible = True
'       TxtClientID.Locked = True
'       TxtClientID.BackColor = &H80000000
'       MSHFlexGrid2.ColWidth(1) = 1300
'       MSHFlexGrid2.ColWidth(2) = 1200
'       MSHFlexGrid2.ColWidth(7) = 1000
'       MSHFlexGrid2.ColWidth(8) = 1000
'       MSHFlexGrid2.ColWidth(9) = 1000
'       MSHFlexGrid2.ColWidth(10) = 1500
'       Label8.Top = 810
'       TxtClientID.Top = 1080
'       Label42.Top = 90
'       Label42.Left = 10485
'       TxtRRNRemarks.Top = 360
'       TxtRRNRemarks.Left = 7590
       
'
'       LblRRNQty.Top = 75
'       LblRRNQty.Left = 4470
'       TxtRRNQty.Top = 360
'       TxtRRNQty.Left = 3915
'
'       TxtBalanceQty.Top = 360
'       TxtBalanceQty.Left = 5760
'       Label38.Top = 75
'       Label38.Left = 6158
End If


End Sub

Private Sub Check2_Click()
CmbGCTDClient.Visible = False
TxtGSLISINno.Visible = False
If Check2.Value = 1 Then
   CmbGCTDClient.Visible = True
Else
   If mExchangeTypeID = 2 And Check1.Value = 1 Then
      CmbGCTDClient.Visible = True
   Else
      TxtGSLISINno.Visible = True
   End If
End If
CmbLot.Text = ""

If mExchangeTypeID = 3 And Check2.Value = 0 Then
   CmbGodownNo.Enabled = False
   CmbStack.Enabled = False
   CmbLot.Enabled = False
   CmdGCTDAdd2List.Enabled = False
Else
   CmbGodownNo.Enabled = True
   CmbStack.Enabled = True
   CmbLot.Enabled = True
   CmdGCTDAdd2List.Caption = "Add to List"
   CmdGCTDAdd2List.Enabled = True
End If




End Sub

Private Sub CmbAgent_GotFocus()
tmp = CmbAgent.Text
Call TableMst_Agent
CmbAgent.Clear
If RsMst_Agent.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_Agent.RecordCount
    CmbAgent.AddItem RsMst_Agent!AgentName
    RsMst_Agent.MoveNext
Next
CmbAgent.Text = tmp
End Sub
'----
Private Sub CmbAgent_LostFocus()
If CmbAgent.Text = "" Then
   Exit Sub
End If
RsMst_Agent.MoveFirst
RsMst_Agent.Find "agentName = '" & CmbAgent.Text & "'"
If RsMst_Agent.EOF Then
   MsgBox "Invalid  selection "
   CmbAgent.SetFocus
   Exit Sub
End If
mAgentID = RsMst_Agent!AgentID
SSTab1.Tab = 0
If TxtRRNNo.Visible = True And TxtRRNNo.Enabled = True Then
   TxtRRNNo.SetFocus
End If
End Sub

Private Sub CmbCMPID_GotFocus()
tmp = CmbCMPID.Text
Call TableMst_CMP(" Where CloseDate is Null")
CmbCMPID.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_CMP.RecordCount
   CmbCMPID.AddItem RsMst_CMP!CMPName
   RsMst_CMP.MoveNext
Next
CmbCMPID.Text = tmp
End Sub
'----
Private Sub CmbCMPID_LostFocus()
If CmbCMPID.Text = "" Then
   Exit Sub
End If
RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPName = '" & CmbCMPID.Text & "'"
If RsMst_CMP.EOF Then
   MsgBox "Invalid  selection !!!! "
   CmbCMPID.SetFocus
   Exit Sub
End If
mCMPID = RsMst_CMP!CMPID
TxtPreFix = GetPreFix(RsMst_CMP!CMPID, "P")
If mExchangeTypeID = 1 Then 'NON NCDEX
   Call WHRDetails(mCMPID, mCommodityId, TxtRRNNo.Text)
End If
If mExchangeTypeID = 2 Then 'NCDEX
   Call RRNDetails(mCMPID, mCommodityId, TxtRRNNo.Text)
End If
If mExchangeTypeID = 3 Then 'SPOT
   Call CDCDetails(mCMPID, mCommodityId, TxtRRNNo.Text)
End If
Call Clear_GCTDTabControl
Call Grid_Head_GCTD
 
TxtCxTFDate.Value = IIf(IsNull(RsMst_CMP!CurrentDate), RsMst_CMP!StartDate, RsMst_CMP!CurrentDate) + 1
'If SaveCmd.Enabled = True Then
'   MsgBox "CITF Date is " & TxtCxTFDate
'End If

End Sub
Private Sub CmbCommodityID_GotFocus()
If CmbExchangeTypeID.Text <> "" Then
   If CmbCMPID.Text <> "" Then
      tmp = CmbCommodityID.Text
      Call TableMst_Commodity
      CmbCommodityID.Clear
      If RsMst_Commodity.RecordCount = 0 Then
         MsgBox "No Record found !!!! "
         Exit Sub
      End If
      For i = 1 To RsMst_Commodity.RecordCount
          CmbCommodityID.AddItem RsMst_Commodity!Commodity
          RsMst_Commodity.MoveNext
      Next
      CmbCommodityID.Text = tmp
   Else
      MsgBox "Please Select CMP!!!!!"
      CmbCMPID.SetFocus
      Exit Sub
   End If
Else
   MsgBox "Please Select Mode !!!!"
   CmbExchangeTypeID.SetFocus
   Exit Sub
End If
End Sub
'----
Private Sub CmbCommodityID_LostFocus()
If CmbCommodityID.Text = "" Then
   Exit Sub
End If

RsMst_Commodity.MoveFirst
RsMst_Commodity.Find "Commodity = '" & CmbCommodityID.Text & "'"
If RsMst_Commodity.EOF Then
   MsgBox "Invalid  selection "
   CmbCommodityID.SetFocus
   Exit Sub
End If
mCommodityId = RsMst_Commodity!CommodityID
Call TableMst_CommodityDetail("Where CommodityId = " & mCommodityId & " and ExchangeTypeId = " & mExchangeTypeID)
If RsMst_CommodityDetail.RecordCount > 0 Then
   TxtCGunnyWt = RsMst_CommodityDetail!GunnyWt
   TxtCLotSize = IIf(IsNull(RsMst_CommodityDetail!LotSize), "", RsMst_CommodityDetail!LotSize)
   TxtCVariation = IIf(IsNull(RsMst_CommodityDetail!Variation), "", RsMst_CommodityDetail!Variation)
Else
   MsgBox "No Record Found For Exchange" & CmbExchangeTypeID.Text & " - " & CmbCommodityID.Text
   CmbCommodityID.SetFocus
   Exit Sub
End If
Call Clear_RWCTabControl
'Call Grid_Head_RRN
Call Clear_TTDTabControl
Call Grid_Head_TTD
Call Clear_GCTDTabControl
Call Grid_Head_GCTD

'Call CheckRejection(Check1.Value)
'Call Check1_Click
Call DisplayRWCModification

'If mExchangeTypeID = 1 Then 'NON NCDEX
'   Call WHRDetails(mCMPID, mCommodityID, TxtRRNNo.Text)
'End If
'If mExchangeTypeID = 2 Then 'NCDEX
'   Call RRNDetails(mCMPID, mCommodityID, TxtRRNNo.Text)
'End If
'If mExchangeTypeID = 3 Then 'SPOT
'   Call CDCDetails(mCMPID, mCommodityID, TxtRRNNo.Text)
'End If

'SSTab1.Tab = 0
End Sub

Private Sub CmbExchangeTypeID_GotFocus()
tmp = CmbExchangeTypeID.Text
Call TableMst_ExchangeType
CmbExchangeTypeID.Clear
If RsMst_ExchangeType.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_ExchangeType.RecordCount
    CmbExchangeTypeID.AddItem RsMst_ExchangeType!ExchangeType
    RsMst_ExchangeType.MoveNext
Next
CmbExchangeTypeID.Text = tmp
End Sub
'----
Private Sub CmbExchangeTypeID_LostFocus()
If CmbExchangeTypeID.Text = "" Then
   Exit Sub
End If
RsMst_ExchangeType.MoveFirst
RsMst_ExchangeType.Find "ExchangeType = '" & CmbExchangeTypeID.Text & "'"
If RsMst_ExchangeType.EOF Then
   MsgBox "Invalid  selection "
   CmbExchangeTypeID.SetFocus
   Exit Sub
End If
mExchangeTypeID = RsMst_ExchangeType!ExchangeTypeID
CmbCommodityID.Text = ""
mCommodityId = 0
TxtCGunnyWt = ""
TxtCLotSize = ""
TxtCVariation = ""

Call Clear_RWCTabControl
Call Clear_TTDTabControl
Call Clear_GCTDTabControl
Call Grid_Head_RRN
Call Grid_Head_TTD
Call Grid_Head_GCTD


Call DisplayRWCModification
Check1.Value = 0
'Call Check1_Click

End Sub
Private Sub DisplayRWCModification()

If mExchangeTypeID = 1 Then 'NON NCDEX
   Call SetCaptions("WHR No", "WHR No", "Client Name", "WHR Qty", "Depository ID", "Depository Name", False)
   Call WHRDetails(mCMPID, mCommodityId, TxtRRNNo.Text)
End If
If mExchangeTypeID = 2 Then 'NCDEX
   If Check1.Value = 0 Then
      Call SetCaptions("Remat Request No. (RRN)", "RRN No", "ISIN No", "RRN Qty", "DP ID", "DP Name", True)
      Call RRNDetails(mCMPID, mCommodityId, TxtRRNNo.Text)
   Else
      Call Check1_Click
   End If
End If
If mExchangeTypeID = 3 Then 'SPOT
   Call SetCaptions("CDC No", "CDC-No", "CDC-No", "CDC Qty", "TM ID", "Trading Member Name", True)
   Call CDCDetails(mCMPID, mCommodityId, TxtRRNNo.Text)
End If

End Sub
Private Sub SetCaptions(TabCaption_, RRNNo_, ISINNo_, RRNQty_, DPID_, DPName_, Yn)

SSTab1.TabCaption(0) = TabCaption_
LblRRNNo.Caption = RRNNo_
LblISINNo.Caption = ISINNo_
LblRRNQty.Caption = RRNQty_
LblDPID.Caption = DPID_
LblDPName.Caption = DPName_
Label32.Caption = ISINNo_
Label29.Visible = Yn
Check1.Visible = Yn
If Yn = False Then
   MSHFlexGrid2.ColWidth(2) = 0
   LblISINNo.Visible = Yn
   TxtISINno.Visible = Yn
   If mExchangeTypeID = 3 Then
      Label32.Visible = Yn
      TxtGSLISINno.Visible = Yn
   Else
      Label32.Visible = True
      TxtGSLISINno.Visible = True
   End If
Else
   If Check1.Value = 0 Then
      LblISINNo.Visible = Yn
      TxtISINno.Visible = Yn
      Label32.Visible = Yn
      TxtGSLISINno.Visible = Yn
   Else
      If mExchangeTypeID = 2 Then
         SSTab1.TabCaption(0) = "Client Detail "
      End If
      If mExchangeTypeID = 3 Then
         SSTab1.TabCaption(0) = "CDC Detail "
      End If
      LblISINNo.Visible = False
      TxtISINno.Visible = False
      If mExchangeTypeID = 3 Then
         Label32.Visible = True
         TxtGSLISINno.Visible = True
      End If
   End If
End If
MSHFlexGrid2.TextMatrix(1, 1) = RRNNo_
MSHFlexGrid2.TextMatrix(1, 5) = DPID_
MSHFlexGrid2.TextMatrix(1, 6) = DPName_
MSHFlexGrid2.TextMatrix(1, 7) = RRNQty_
MSHFlexGrid4.TextMatrix(1, 12) = ISINNo_

End Sub

Private Sub CmdAgent_Click()
FrmMst_Agent.Show
End Sub

Private Sub CmdCancel_Click()
TxtHologramNo = ""
Frame3.Visible = False
End Sub

'Private Sub CmdPrint_Click()
''chaaay23
'End Sub

Private Sub DeleteCmd_Click()
If RsTxn_CxTF_Details!Flag = "P" Then
   MsgBox "Print out is Generated, You Can't Edit This CDTF   !!!! "
   Exit Sub
End If
Dim ans As Variant
'On Error GoTo msgError
ans = MsgBox("Are you Sure you want to Delete this Record", vbYesNo)
If ans = vbYes Then
   Call DeleteChildTableDetail_RRN
   tmp = "select * from Txn_CxTF_TruckDetail where sm_prefix = '" & TxtPreFix & "' And TxnType = 'W'  And  CxTFNo = " & TxtCxTFNo
   Call TmpTable
   TmpRs.Close
   RsTxn_CxTF_Details.Delete
   RsTxn_CxTF_Details.Update
   If RsTxn_CxTF_Details.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsTxn_CxTF_Details.MoveNext
   If RsTxn_CxTF_Details.EOF() Then
      RsTxn_CxTF_Details.MoveLast
   End If
   Call InData
End If
Exit Sub
msgError:
MsgBox "Child Record Present "
RsTxn_CxTF_Details.CancelUpdate
End Sub

Sub DeleteChildTableDetail_RRN()
Call TableMst_GSL("where  sm_prefix = '" & TxtPreFix & "'")
tmp = "select * from Txn_CxTF_GSL where  sm_prefix = '" & TxtPreFix & "' And TxnType = 'W' And CxTFNo = " & TxtCxTFNo
Call TmpTable
For i = 1 To TmpRs.RecordCount
    RsMst_GSL.MoveFirst
    RsMst_GSL.Find "GSLID = " & TmpRs!GSLID
    If Not RsMst_GSL.EOF Then
       RsMst_GSL!CDTFBags = IIf(IsNull(RsMst_GSL!CDTFBags), 0, RsMst_GSL!CDTFBags) - Val(MSHFlexGrid4.TextMatrix(i + 1, 6))
       RsMst_GSL!CDTFWeight = IIf(IsNull(RsMst_GSL!CDTFWeight), 0, RsMst_GSL!CDTFWeight) - Val(MSHFlexGrid4.TextMatrix(i + 1, 10))
       RsMst_GSL!BalanceBags = IIf(IsNull(RsMst_GSL!BalanceBags), 0, RsMst_GSL!BalanceBags) + Val(MSHFlexGrid4.TextMatrix(i + 1, 6))
       
       RsMst_GSL!BalanceWeight = IIf(IsNull(RsMst_GSL!BalanceWeight), 0, RsMst_GSL!BalanceWeight) + Val(MSHFlexGrid4.TextMatrix(i + 1, 10))
       RsMst_GSL!WithdrawalAdjWt = IIf(IsNull(RsMst_GSL!WithdrawalAdjWt), 0, RsMst_GSL!WithdrawalAdjWt) - Val(MSHFlexGrid4.TextMatrix(i + 1, 8))
       RsMst_GSL.Update
       If Not IsNull(RsMst_GSL!Flag1) Then
          RsMst_GSL!Flag = RsMst_GSL!Flag1
       End If
       RsMst_GSL.Update
    End If
    TmpRs.MoveNext
Next

If mExchangeTypeID = 2 Then
   Call TableMst_RRN
End If
If mExchangeTypeID = 1 Then
   Call TableTxn_WHRDetail(" where  sm_prefix = '" & TxtPreFix & "'")
End If
    
If mExchangeTypeID = 3 Then
   Call TableTxn_CDC_GSLDetail(" Where  sm_prefix = '" & TxtPreFix & "'")
End If
For i = 2 To MSHFlexGrid2.Rows - 1
    
    If mExchangeTypeID = 2 Then
       Call TableMst_RRN
       If Check1.Value = 0 Then
          RsMst_RRN.MoveFirst
          If MSHFlexGrid2.TextMatrix(i, 12) <> "" Then
             RsMst_RRN.Find "RRNID = " & Val(MSHFlexGrid2.TextMatrix(i, 12))
             If Not RsMst_RRN.EOF Then
                RsMst_RRN!DeliveredQty = RsMst_RRN!DeliveredQty - Val(MSHFlexGrid2.TextMatrix(i, 9))
                RsMst_RRN.Update
                RsMst_RRN.Requery
             End If
          End If
       End If
    End If
    
    
    If mExchangeTypeID = 1 Then
       If Check1.Value = 0 Then
          RsTxn_WHRDetail.MoveFirst
          RsTxn_WHRDetail.Find "WHRNo = " & Val(MSHFlexGrid2.TextMatrix(i, 12))
          If Not RsTxn_WHRDetail.EOF Then
             RsTxn_WHRDetail!DeliveredQuantity = RsTxn_WHRDetail!DeliveredQuantity - Val(MSHFlexGrid2.TextMatrix(i, 9))
             RsTxn_WHRDetail!DeliveredNoOfBags = RsTxn_WHRDetail!DeliveredNoOfBags - Val(MSHFlexGrid2.TextMatrix(i, 15))
             RsTxn_WHRDetail.Update
             RsTxn_WHRDetail.Requery
          End If
       End If
    End If
    
    If mExchangeTypeID = 3 Then
       If Check1.Value = 0 Then
          RsTxn_CDC_GSLDetail.MoveFirst
          RsTxn_CDC_GSLDetail.Find "CDCNo = " & Val(MSHFlexGrid2.TextMatrix(i, 12))
          If Not RsTxn_CDC_GSLDetail.EOF Then
             RsTxn_CDC_GSLDetail!DeliveredQty = RsTxn_CDC_GSLDetail!DeliveredQty - Val(MSHFlexGrid2.TextMatrix(i, 9))
             RsTxn_CDC_GSLDetail!Flag = "P"
             RsTxn_CDC_GSLDetail.Update
             RsTxn_CDC_GSLDetail.Requery
          End If
       End If
    End If


Next
tmp = "Delete from Txn_CxTF_ClientDetail where  sm_prefix = '" & TxtPreFix & "' And TxnType = 'W' And  CxTFNo = " & TxtCxTFNo
Call TmpTable
tmp = "Delete from Txn_CxTF_GSL where  sm_prefix = '" & TxtPreFix & "' And TxnType = 'W' And  CxTFNo = " & TxtCxTFNo
Call TmpTable
tmp = "Delete from Txn_CxTF_TruckDetail where  sm_prefix = '" & TxtPreFix & "' And TxnType = 'W'  And  CxTFNo = " & TxtCxTFNo
Call TmpTable
tmp = "select * from Txn_CxTF_GSL where  sm_prefix = '" & TxtPreFix & "' And TxnType = 'W'  And  CxTFNo = " & TxtCxTFNo
Call TmpTable
TmpRs.Close

End Sub




Private Sub SSTab1_Click(PreviousTab As Integer)

If SSTab1.Tab = 0 Then
   If SaveCmd.Enabled = True Then
      Call UpdateRRNGridStock
   End If
End If
If SSTab1.Tab = 3 Then
   If SaveCmd.Enabled = True Then
      Call UpdateErrorGrid
   End If
End If




End Sub
Private Sub UpdateErrorGrid()
Call Grid_Head_Error
For i = 2 To MSHFlexGrid2.Rows - 1
    If mExchangeTypeID = 1 Then
       If MSHFlexGrid2.TextMatrix(i, 11) <> "" Then
          Call UpdateDataISIN_WHR(MSHFlexGrid2.TextMatrix(i, 4), Val(MSHFlexGrid2.TextMatrix(i, 8)), "R", 0)
       End If
   Else
       If MSHFlexGrid2.TextMatrix(i, 2) <> "" Then
          Call UpdateDataISIN_WHR(MSHFlexGrid2.TextMatrix(i, 2), Val(MSHFlexGrid2.TextMatrix(i, 8)), "R", 0)
       End If
    End If
Next

For i = 2 To MSHFlexGrid4.Rows - 1
    If MSHFlexGrid4.TextMatrix(i, 12) <> "" Then
       If mExchangeTypeID = 1 Then
          Call UpdateDataISIN_WHR(MSHFlexGrid4.TextMatrix(i, 12), Val(MSHFlexGrid4.TextMatrix(i, 10)), "G", Val(MSHFlexGrid4.TextMatrix(i, 8)))
       Else
          If MSHFlexGrid4.TextMatrix(i, 17) = 1 Then
             Call UpdateDataISIN_WHR(MSHFlexGrid4.TextMatrix(i, 12), Val(MSHFlexGrid4.TextMatrix(i, 10)), "G", Val(MSHFlexGrid4.TextMatrix(i, 8)))
          Else
             Call UpdateDataISIN_WHR(MSHFlexGrid4.TextMatrix(i, 12), Val(MSHFlexGrid4.TextMatrix(i, 10)), "G", Val(MSHFlexGrid4.TextMatrix(i, 8)))
          End If
       End If
    End If
Next


For u = 2 To MSHFlexGrid5.Rows - 1
    MSHFlexGrid5.TextMatrix(0, 2) = Val(MSHFlexGrid5.TextMatrix(0, 2)) + Val(MSHFlexGrid5.TextMatrix(u, 2))
    MSHFlexGrid5.TextMatrix(0, 3) = Val(MSHFlexGrid5.TextMatrix(0, 3)) + Val(MSHFlexGrid5.TextMatrix(u, 3))
    MSHFlexGrid5.TextMatrix(0, 4) = Val(MSHFlexGrid5.TextMatrix(0, 4)) + Val(MSHFlexGrid5.TextMatrix(u, 4))
    MSHFlexGrid5.TextMatrix(0, 5) = Val(MSHFlexGrid5.TextMatrix(0, 5)) + Val(MSHFlexGrid5.TextMatrix(u, 5))
    MSHFlexGrid5.TextMatrix(0, 6) = Val(MSHFlexGrid5.TextMatrix(0, 6)) + Val(MSHFlexGrid5.TextMatrix(u, 6))
    
    MSHFlexGrid5.TextMatrix(u, 2) = Format(MSHFlexGrid5.TextMatrix(u, 2), "######0.000")
    MSHFlexGrid5.TextMatrix(u, 3) = Format(MSHFlexGrid5.TextMatrix(u, 3), "######0.000")
    MSHFlexGrid5.TextMatrix(u, 4) = Format(MSHFlexGrid5.TextMatrix(u, 4), "######0.000")
    MSHFlexGrid5.TextMatrix(u, 5) = Format(MSHFlexGrid5.TextMatrix(u, 5), "######0.000")
    MSHFlexGrid5.TextMatrix(u, 6) = Format(MSHFlexGrid5.TextMatrix(u, 6), "######0.000")
    
    MSHFlexGrid5.TextMatrix(0, 2) = Format(MSHFlexGrid5.TextMatrix(0, 2), "######0.000")
    MSHFlexGrid5.TextMatrix(0, 3) = Format(MSHFlexGrid5.TextMatrix(0, 3), "######0.000")
    MSHFlexGrid5.TextMatrix(0, 4) = Format(MSHFlexGrid5.TextMatrix(0, 4), "######0.000")
    MSHFlexGrid5.TextMatrix(0, 5) = Format(MSHFlexGrid5.TextMatrix(0, 5), "######0.000")
    MSHFlexGrid5.TextMatrix(0, 6) = Format(MSHFlexGrid5.TextMatrix(0, 6), "######0.000")
    MSHFlexGrid5.Row = u
    If Val(MSHFlexGrid5.TextMatrix(u, 5)) <> 0 Then
       MSHFlexGrid5.Col = 5
       MSHFlexGrid5.CellBackColor = vbYellow
       MSHFlexGrid5.Col = 1
       MSHFlexGrid5.CellBackColor = vbYellow
   End If
Next

End Sub

Private Sub UpdateDataISIN_WHR(ISIN_ As String, Wght_ As Double, Mod_ As String, AdjWght_ As Double)
Dim u As Variant
For u = 2 To MSHFlexGrid5.Rows - 1
    If MSHFlexGrid5.TextMatrix(u, 1) = ISIN_ Then
       If Mod_ = "R" Then
          MSHFlexGrid5.TextMatrix(u, 2) = Val(MSHFlexGrid5.TextMatrix(u, 2)) + Wght_
       Else
          MSHFlexGrid5.TextMatrix(u, 3) = Val(MSHFlexGrid5.TextMatrix(u, 3)) + Wght_
          MSHFlexGrid5.TextMatrix(u, 4) = Val(MSHFlexGrid5.TextMatrix(u, 4)) + AdjWght_
          If Val(MSHFlexGrid5.TextMatrix(u, 3)) > Val(MSHFlexGrid5.TextMatrix(u, 2)) Then
             MSHFlexGrid5.TextMatrix(u, 5) = Val(MSHFlexGrid5.TextMatrix(u, 3)) - (Val(MSHFlexGrid5.TextMatrix(u, 2)))
             MSHFlexGrid5.TextMatrix(u, 6) = 0
          Else
             MSHFlexGrid5.TextMatrix(u, 6) = Val(MSHFlexGrid5.TextMatrix(u, 2)) - (Val(MSHFlexGrid5.TextMatrix(u, 3)))
          End If
       End If
       Exit Sub
    End If
Next
If Mod_ = "R" Then
   u = MSHFlexGrid5.Rows - 1
   MSHFlexGrid5.Rows = MSHFlexGrid5.Rows + 1
   MSHFlexGrid5.TextMatrix(u, 1) = ISIN_
   MSHFlexGrid5.TextMatrix(u, 2) = Wght_
End If
End Sub

Private Sub UpdateRRNGridStock()
Diff_wt = 0
For i = 2 To MSHFlexGrid2.Rows - 1
    MSHFlexGrid2.TextMatrix(i, 9) = 0 '"Delived Qty"
    MSHFlexGrid2.TextMatrix(i, 15) = 0 '"Delived Bags"
Next
MSHFlexGrid2.TextMatrix(0, 9) = 0 '"Delived Qty"
MSHFlexGrid2.TextMatrix(0, 15) = 0 '"Delived Bags"
For i = 2 To MSHFlexGrid4.Rows - 1
    If MSHFlexGrid4.TextMatrix(i, 12) <> "" Then
       If mExchangeTypeID = 1 Then
          If MSHFlexGrid4.TextMatrix(i, 17) = 1 Then
             Call UpdateDataClient(MSHFlexGrid4.TextMatrix(i, 16), Val(MSHFlexGrid4.TextMatrix(i, 10)), 0)
          Else
             Call UpdateDataClient(MSHFlexGrid4.TextMatrix(i, 16), Val(MSHFlexGrid4.TextMatrix(i, 10)), Val(MSHFlexGrid4.TextMatrix(i, 6)))
          End If
       Else
          If MSHFlexGrid4.TextMatrix(i, 17) = 1 Then
             Call UpdateDataClient(MSHFlexGrid4.TextMatrix(i, 12), Val(MSHFlexGrid4.TextMatrix(i, 10)), 0)
          Else
             Call UpdateDataClient(MSHFlexGrid4.TextMatrix(i, 12), Val(MSHFlexGrid4.TextMatrix(i, 10)), Val(MSHFlexGrid4.TextMatrix(i, 6)))
          End If
       End If
    End If
Next
End Sub
Sub UpdateDataClient(ISIN_ As Variant, EWeight As Double, DBags As Double)

Dim Ci, Bal, PQty, Cl_IS As Variant
PQty = EWeight


For i = 2 To MSHFlexGrid2.Rows - 1
    If mExchangeTypeID = 1 Then
       Cl_IS = MSHFlexGrid2.TextMatrix(i, 11)
    ElseIf mExchangeTypeID = 2 Then
       If Check1.Value = 0 Then
          Cl_IS = MSHFlexGrid2.TextMatrix(i, 2)
       Else
          Cl_IS = MSHFlexGrid2.TextMatrix(i, 4)
       End If
    Else
       Cl_IS = MSHFlexGrid2.TextMatrix(i, 1)
    End If
    If LCase(Cl_IS) = LCase(ISIN_) Then
       Bal = Val(MSHFlexGrid2.TextMatrix(i, 8)) - Val(MSHFlexGrid2.TextMatrix(i, 9))
       If Bal > 0 Then
          If Bal >= (EWeight) Then
             MSHFlexGrid2.TextMatrix(i, 9) = Val(MSHFlexGrid2.TextMatrix(i, 9)) + EWeight '"Weight"
             MSHFlexGrid2.TextMatrix(0, 9) = Val(MSHFlexGrid2.TextMatrix(0, 9)) + EWeight '"Weight"
             EWeight = 0
          Else
             EWeight = EWeight - Bal
             MSHFlexGrid2.TextMatrix(i, 9) = Val(MSHFlexGrid2.TextMatrix(i, 9)) + Bal '"Weight"
             MSHFlexGrid2.TextMatrix(0, 9) = Val(MSHFlexGrid2.TextMatrix(0, 9)) + Bal '"Weight"
          End If
       End If
       If EWeight = 0 Then Exit For
    End If
Next

' Adj No of Bags
If mExchangeTypeID = 1 Or (mExchangeTypeID = 2 And Check1.Value = 1) Then
   PQty = DBags
   For i = 2 To MSHFlexGrid2.Rows - 1
       If mExchangeTypeID = 1 Then
          Cl_IS = MSHFlexGrid2.TextMatrix(i, 11)
       ElseIf mExchangeTypeID = 2 Then
          If Check1.Value = 0 Then
             Cl_IS = MSHFlexGrid2.TextMatrix(i, 2)
          Else
             Cl_IS = MSHFlexGrid2.TextMatrix(i, 4)
          End If
       
       Else
          Cl_IS = MSHFlexGrid2.TextMatrix(i, 1)
       End If
       If Cl_IS = ISIN_ Then
           Bal = Val(MSHFlexGrid2.TextMatrix(i, 14)) - Val(MSHFlexGrid2.TextMatrix(i, 15))
           If Bal > 0 Then
              If Bal >= (DBags) Then
                 MSHFlexGrid2.TextMatrix(i, 15) = Val(MSHFlexGrid2.TextMatrix(i, 15)) + DBags '"Weight"
                 MSHFlexGrid2.TextMatrix(0, 15) = Val(MSHFlexGrid2.TextMatrix(0, 15)) + DBags '"Weight"
                 EWeight = 0
              Else
                 DBags = DBags - Bal
                 MSHFlexGrid2.TextMatrix(i, 15) = Val(MSHFlexGrid2.TextMatrix(i, 15)) + Bal '"Weight"
                 MSHFlexGrid2.TextMatrix(0, 15) = Val(MSHFlexGrid2.TextMatrix(0, 15)) + Bal '"Weight"
              End If
           End If
           If DBags = 0 Then Exit For
        End If
    Next
    If DBags <> 0 Then
       Diff_wt = 1
    End If
End If
If EWeight <> 0 Then
   Diff_wt = 1
End If
End Sub



Private Sub TxtClientID_LostFocus()
CmbClientDP.Visible = False
If TxtClientID = "" Then Exit Sub
If CmbExchangeTypeID.Text = "" Then Exit Sub

If CmbCommodityID.Text = "" Then Exit Sub
If CmbCMPID.Text = "" Then Exit Sub
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
tmp = "SELECT Mst_Depository.DepositoryID, Mst_Depository.DepositoryName, Mst_Client.ClientID, " & _
      "Mst_Client.ClientName,  Mst_Client.ClientIDRow " & _
      "FROM Mst_Depository INNER JOIN Mst_Client ON Mst_Depository.DepositoryID = Mst_Client.DepositoryID " & _
      "WHERE Mst_Depository.ExchangeTypeID = " & mExchangeTypeID & _
      " And  (((Mst_Client.ExchangeTypeID)=" & mExchangeTypeID & ") AND ((Mst_Client.ClientID)='" & TxtClientID & "'));"
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

Private Sub CmbClientDP_LostFocus()

Dim Str_() As String
Str_ = Split(CmbClientDP.Text, "~")
If UBound(Str_) <> 1 Then
   MsgBox "Invalid Selection !!!!! "
   CmbClientDP.SetFocus
   Exit Sub
End If
TmpRs1.MoveFirst
TmpRs1.Find "DepositoryID  = '" & Str_(0) & "'"
If TmpRs1.EOF Then
   MsgBox "Invalid Selection !!!!! "
   CmbClientDP.SetFocus
   Exit Sub
End If


TxtClientName = TmpRs1!ClientName
TxtDPID = TmpRs1!DepositoryID
TxtDPName = TmpRs1!DepositoryName
CmbClientDP.Visible = False
mClientIDRow = TmpRs1!ClientIDRow


tmp = "SELECT Sum(Mst_GSL.BalanceWeight) AS SBalWeight, Sum(Mst_GSL.BalanceBags) AS SBalBags, "
tmp = tmp & "  Sum(Mst_GSL.DepositeWeight) AS SDWeight,Sum(Mst_GSL.DepositeBags) AS SDBags FROM Mst_GSL "
tmp = tmp & " WHERE (((Mst_GSL.SM_Prefix)='" & TxtPreFix & "') AND ((Mst_GSL.CMPID)=" & mCMPID & ") AND "
tmp = tmp & " ((Mst_GSL.CommodityID)=" & mCommodityId & " ) AND ((Mst_GSL.ClientIDRow)=" & mClientIDRow & ") AND "
tmp = tmp & " ((Mst_GSL.ExchangeTypeID)= " & mExchangeTypeID & ") AND ((Mst_GSL.Flag)='J'));"
Call TmpTable

TxtRRNQty = TmpRs.Fields(2)
TxtBalanceQty = TmpRs.Fields(0)
TxtBalanceBags = TmpRs.Fields(1)
TxtRRNBags = TmpRs.Fields(3)
TxtRRNNo = mClientIDRow
CmdAdd2ListRRN.SetFocus
End Sub

Private Sub CmbGodownNo_GotFocus()
If CmbCMPID.Text = "" Then
   MsgBox "Select CMP !!!! "
   CmbCMPID.SetFocus
   Exit Sub
End If
If CmbCommodityID.Text <> "" Then
   tmp = CmbGodownNo.Text
   Call TableMst_Godown(" Where CloseDate is null And CMPID = " & mCMPID)
   CmbGodownNo.Clear
   If RsMst_Godown.RecordCount = 0 Then
      MsgBox "No Record found !!!! "
      Exit Sub
   End If
   For i = 1 To RsMst_Godown.RecordCount
       CmbGodownNo.AddItem RsMst_Godown!GodownNo
       RsMst_Godown.MoveNext
   Next
   CmbGodownNo.Text = tmp
Else
   MsgBox "Please Select Commodity!!!"
   CmbCommodityID.SetFocus
   Exit Sub
End If
End Sub
'----
Private Sub CmbGodownNo_LostFocus()
If CmbGodownNo.Text = "" Then
   Exit Sub
End If
RsMst_Godown.MoveFirst
RsMst_Godown.Find "GodownNo = '" & CmbGodownNo.Text & "'"
If RsMst_Godown.EOF Then
   MsgBox "Invalid  selection "
   CmbGodownNo.SetFocus
   Exit Sub
End If
mGodownID = RsMst_Godown!GodownID
End Sub
Private Sub CmbLot_GotFocus()
If CmbStack.Text = "" Then
   MsgBox "Please Select Stack No !!!"
   CmbStack.SetFocus
   CmbLot.Text = ""
   Exit Sub
End If
tmp = CmbLot.Text
If Check2.Value = 1 Then
   Call TableMst_GSL(" Where  (SpillageFlag = '1' And CMPID = " & mCMPID & " And CommodityID = " & mCommodityId & "  And GodownID = " & mGodownID & " And StackNo = '" & CmbStack.Text & "' And Flag <> 'E' )")
ElseIf mExchangeTypeID <> 1 Then
   If Check1.Value = 0 Then
      Call TableMst_GSL(" Where ((ExchangeTypeId = " & mExchangeTypeID & " And CMPID = " & mCMPID & " And CommodityID = " & mCommodityId & " And GodownID = " & mGodownID & " And StackNo = '" & CmbStack.Text & "' And ISINID <> '0' And Flag <> 'E')) ")
   ElseIf Check1.Value = 1 Then
      Call TableMst_GSL(" Where ((ExchangeTypeId = " & mExchangeTypeID & " And CMPID = " & mCMPID & " And CommodityID = " & mCommodityId & " And GodownID = " & mGodownID & " And StackNo = '" & CmbStack.Text & "' And ISINID <> '0' And (Flag <> 'E' And Flag = 'J'))) ")
   End If
Else
   Call TableMst_GSL(" Where  (ExchangeTypeId = " & mExchangeTypeID & " And CMPID = " & mCMPID & " And CommodityID = " & mCommodityId & "  And GodownID = " & mGodownID & " And StackNo = '" & CmbStack.Text & "' And Flag <> 'E' )")
End If
CmbLot.Clear
If RsMst_GSL.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_GSL.RecordCount
    CmbLot.AddItem RsMst_GSL!LotNo
    RsMst_GSL.MoveNext
Next i
CmbLot.Text = tmp
End Sub
'----
Private Sub CmbLot_LostFocus()
Dim ans As Variant
If CmbLot.Text = "" Then
   Exit Sub
End If
RsMst_GSL.MoveFirst
RsMst_GSL.Find "LotNo = '" & CmbLot.Text & "'"
If RsMst_GSL.EOF Then
   MsgBox "Invalid  selection "
   CmbLot.SetFocus
   Exit Sub
End If
mGSLID = RsMst_GSL!GSLID
mGSLClientIDRow = IIf(IsNull(RsMst_GSL!ClientIDRow), 0, RsMst_GSL!ClientIDRow)
TxtGSLAvailableNofBags = IIf(IsNull(RsMst_GSL!DepositeBags), 0, RsMst_GSL!DepositeBags) - IIf(IsNull(RsMst_GSL!CDTFBags), 0, RsMst_GSL!CDTFBags)
TxtGSLAvailableGrossWgt = IIf(IsNull(RsMst_GSL!DepositeWeight), 0, RsMst_GSL!DepositeWeight) - IIf(IsNull(RsMst_GSL!CDTFWeight), 0, RsMst_GSL!CDTFWeight)
mPledgeID = IIf(IsNull(RsMst_GSL!PledgeNo), "", RsMst_GSL!PledgeNo)
TxtGSLPledgedNofBags = IIf(IsNull(RsMst_GSL!PledgeBags), "", RsMst_GSL!PledgeBags)
TxtGSLPledgedWeight = IIf(IsNull(RsMst_GSL!PledgeWeight), "", RsMst_GSL!PledgeWeight)
If mExchangeTypeID <> 1 Then
   If Check1.Value = 0 Then
      TxtGSLISINno = IIf(IsNull(RsMst_GSL!ISINID), "", RsMst_GSL!ISINID)
   End If
Else
   TxtGSLISINno = GetClientName(mGSLClientIDRow) 'IIf(IsNull(RsMst_GSL!ClientIDRow), "", RsMst_GSL!ClientIDRow)
End If
End Sub
Private Sub CmbStack_GotFocus()
If CmbGodownNo.Text = "" Or mGodownID = "" Then    'IsNull(mGodownID)
   MsgBox "Please Select Godown No!!!"
   Exit Sub
End If
tmp = CmbStack.Text
Call TableMst_GSL(" Where CMPID = " & mCMPID & " And GodownID = " & mGodownID, "D")
CmbStack.Clear
If RsMst_GSL.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_GSL.RecordCount
    CmbStack.AddItem RsMst_GSL!StackNo
    RsMst_GSL.MoveNext
Next
CmbStack.Text = tmp
End Sub
'----
Private Sub CmbStack_LostFocus()
If CmbStack.Text = "" Then
   Exit Sub
End If
RsMst_GSL.MoveFirst
RsMst_GSL.Find "StackNo = '" & CmbStack.Text & "'"
If RsMst_GSL.EOF Then
   MsgBox "Invalid  selection "
   CmbStack.SetFocus
   Exit Sub
End If

End Sub
Private Sub CmdAdd2ListRRN_Click()
If CmbExchangeTypeID.Text = "" Then
   MsgBox "Select Mode"
   CmbExchangeTypeID.SetFocus
   Exit Sub
End If

If CmbCMPID.Text = "" Then
   MsgBox "Select CMP"
   CmbCMPID.SetFocus
   Exit Sub
End If
If CmbCommodityID.Text = "" Then
   MsgBox "Select Commodity"
   CmbCommodityID.SetFocus
   Exit Sub
End If
Dim r As Variant
If Check1.Value = 0 Then
   If TxtRRNNo = "" Then
      If mExchangeTypeID = 1 Then 'NON NCDEX
         MsgBox "Blank WHR No Not Allowed!!! "
      ElseIf mExchangeTypeID = 2 Then 'NCDEX
         MsgBox "Blank RRN No Not Allowed!!! "
      ElseIf mExchangeTypeID = 3 Then 'SPOT
         MsgBox "Blank CDC No Not Allowed!!! "
      End If
      TxtRRNNo.SetFocus
      Exit Sub
    End If
Else
    If TxtClientID = "" Then
       MsgBox "Blank Client ID Not Allowed!!!"
       TxtClientID.SetFocus
       Exit Sub
    End If
End If
If Check1.Value = 0 Then
   If CheckDuplicateRWC = True Then
     'MsgBox "Duplicate Client ID  not allowd !!!"
      If mExchangeTypeID = 1 Then 'NON NCDEX
         MsgBox "Duplicate WHR No Not Allowed!!! "
      ElseIf mExchangeTypeID = 2 Then 'NCDEX
         MsgBox "Duplicate RRN No Not Allowed!!! "
      ElseIf mExchangeTypeID = 3 Then 'SPOT
         MsgBox "Duplicate CDC No Not Allowed!!! "
      End If
      TxtRRNNo.SetFocus
      Exit Sub
   End If
Else
   If CheckDuplicateClient = True Then
      MsgBox "Duplicate Client ID  not allowd !!!"
      TxtClientID.SetFocus
      Exit Sub
   End If
End If
If Check1.Value = 0 Then
   If Val(TxtBalanceQty) = 0 Then
      If mExchangeTypeID = 1 Then
         If Val(TxtBalanceBags) = 0 Then
            MsgBox "All Quantity is Delivered!!!"
            TxtRRNNo.SetFocus
            Exit Sub
         End If
      Else
         MsgBox "All Quantity is Delivered!!!"
         TxtRRNNo.SetFocus
         Exit Sub
      End If
   End If
End If
If mExchangeTypeID = 3 Then
   If Val(TxtBalanceQty) = 0 Then
      MsgBox "All Quantity is Delivered!!!"
      TxtRRNNo.SetFocus
      Exit Sub
   End If
End If

If CmdAdd2ListRRN.Caption = "Update List" Then
   r = MSHFlexGrid2.RowSel
   MSHFlexGrid2.TextMatrix(0, 7) = Val(MSHFlexGrid2.TextMatrix(0, 7)) - Val(MSHFlexGrid2.TextMatrix(r, 7))
   MSHFlexGrid2.TextMatrix(0, 8) = Val(MSHFlexGrid2.TextMatrix(0, 8)) - Val(MSHFlexGrid2.TextMatrix(r, 8))
   MSHFlexGrid2.TextMatrix(0, 9) = Val(MSHFlexGrid2.TextMatrix(0, 9)) - Val(MSHFlexGrid2.TextMatrix(r, 9))
   
   MSHFlexGrid2.TextMatrix(0, 13) = Val(MSHFlexGrid2.TextMatrix(0, 13)) - Val(MSHFlexGrid2.TextMatrix(r, 13))
   MSHFlexGrid2.TextMatrix(0, 14) = Val(MSHFlexGrid2.TextMatrix(0, 14)) - Val(MSHFlexGrid2.TextMatrix(r, 14))
   MSHFlexGrid2.TextMatrix(0, 15) = Val(MSHFlexGrid2.TextMatrix(0, 15)) - Val(MSHFlexGrid2.TextMatrix(r, 15))
   
Else
   If mExchangeTypeID = 3 Then 'SPOT
      Call DisplayGSLData(mGSLID)
   End If
   r = MSHFlexGrid2.Rows - 1
   MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
End If

MSHFlexGrid2.TextMatrix(r, 1) = TxtRRNNo
MSHFlexGrid2.TextMatrix(r, 2) = TxtISINno
MSHFlexGrid2.TextMatrix(r, 3) = TxtClientID
MSHFlexGrid2.TextMatrix(r, 4) = TxtClientName
MSHFlexGrid2.TextMatrix(r, 5) = TxtDPID
MSHFlexGrid2.TextMatrix(r, 6) = TxtDPName
MSHFlexGrid2.TextMatrix(r, 7) = TxtRRNQty
MSHFlexGrid2.TextMatrix(r, 8) = TxtBalanceQty
MSHFlexGrid2.TextMatrix(r, 9) = Val(TxtRRNQty) - Val(TxtBalanceQty)
MSHFlexGrid2.TextMatrix(r, 10) = TxtRRNRemarks
MSHFlexGrid2.TextMatrix(r, 11) = mClientIDRow
MSHFlexGrid2.TextMatrix(r, 12) = mRWCID
MSHFlexGrid2.TextMatrix(r, 13) = TxtRRNBags
MSHFlexGrid2.TextMatrix(r, 14) = TxtBalanceBags

 
 
'If mExchangeTypeID = 1 Then
'   MSHFlexGrid2.TextMatrix(r, 12) = mRWCID
'ElseIf mExchangeTypeID = 2 Then
'   MSHFlexGrid2.TextMatrix(r, 12) = mRWCID
'ElseIf mExchangeTypeID = 3 Then
'   MSHFlexGrid2.TextMatrix(r, 12) = mRWCID
'End If
MSHFlexGrid2.TextMatrix(0, 7) = Val(MSHFlexGrid2.TextMatrix(0, 7)) + Val(MSHFlexGrid2.TextMatrix(r, 7))
MSHFlexGrid2.TextMatrix(0, 8) = Val(MSHFlexGrid2.TextMatrix(0, 8)) + Val(MSHFlexGrid2.TextMatrix(r, 8))
MSHFlexGrid2.TextMatrix(0, 9) = Val(MSHFlexGrid2.TextMatrix(0, 9)) + Val(MSHFlexGrid2.TextMatrix(r, 9))

MSHFlexGrid2.TextMatrix(0, 13) = Val(MSHFlexGrid2.TextMatrix(0, 13)) + Val(MSHFlexGrid2.TextMatrix(r, 13))
MSHFlexGrid2.TextMatrix(0, 14) = Val(MSHFlexGrid2.TextMatrix(0, 14)) + Val(MSHFlexGrid2.TextMatrix(r, 14))
MSHFlexGrid2.TextMatrix(0, 15) = Val(MSHFlexGrid2.TextMatrix(0, 15)) + Val(MSHFlexGrid2.TextMatrix(r, 15))


Call Clear_RWCTabControl
TxtRRNNo.Locked = False
If TxtRRNNo.Visible = True And TxtRRNNo.Enabled = True Then
   TxtRRNNo.SetFocus
End If
End Sub
Function CheckDuplicateRWC() As Boolean
Dim RetVal As Boolean
RetVal = False
If MSHFlexGrid2.Rows > 3 Then
   For u = 2 To MSHFlexGrid2.Rows - 1
       If MSHFlexGrid2.TextMatrix(u, 1) = TxtRRNNo Then
          If LCase(CmdAdd2ListRRN.Caption) = "add to list" Then
             RetVal = True
             Exit For
          Else
             If u <> MSHFlexGrid2.RowSel Then
                RetVal = True
             End If
          End If
       End If
   Next
End If

CheckDuplicateRWC = RetVal
End Function

Private Sub CmdGCTDAdd2List_Click()

Dim r As Variant

If (CmbGodownNo.Text) = "" Then
   MsgBox "Blank Godown No Not Allowed !!! "
   If CmbGodownNo.Enabled = True Then
      CmbGodownNo.SetFocus
   End If
   Exit Sub
End If

If (CmbStack.Text) = "" Then
   MsgBox "Blank Stack  Not Allowed !!! "
   CmbStack.SetFocus
   Exit Sub
End If
If (CmbLot.Text) = "" Then
   MsgBox "Blank Lot Not Allowed !!! "
   CmbLot.SetFocus
   Exit Sub
End If

If CheckDuplicateGSL() = True Then
   MsgBox "Duplicate G.S.L.  Not Allowed !!! "
   CmbGodownNo.SetFocus
   Exit Sub
End If

If Check2.Value = 1 Then
   If Val(TxtGCTDNoofBags.Text) < 0 Then
      MsgBox "Invalid No of Bags !!! "
      TxtGCTDNoofBags.SetFocus
      Exit Sub
   End If
Else
   If Val(TxtGCTDNoofBags.Text) <= 0 Then
      MsgBox "Invalid No of Bags !!! "
      TxtGCTDNoofBags.SetFocus
      Exit Sub
   End If
End If
If Val(TxtGCTDGrossWt) <= 0 Then
   MsgBox "Invalid Weight !!! "
   TxtGCTDGrossWt.SetFocus
   Exit Sub
End If

If Check2.Value = 0 Then
    tmp = Validate_ComBags(mCommodityId, Val(TxtGCTDNoofBags), Val(TxtGCTDGrossWt))
    If Len(tmp) > 1 Then
      MsgBox tmp
      TxtGCTDGrossWt.SetFocus
      Exit Sub
    End If

End If
If MSHFlexGrid2.Rows <= 3 Then
   If mExchangeTypeID = 1 Then 'NON NCDEX
      MsgBox "Missing WHR Detail !!!! "
   ElseIf mExchangeTypeID = 2 Then 'NCDEX
      MsgBox "Missing RRN Detail !!!! "
   ElseIf mExchangeTypeID = 3 Then 'SPOT
      MsgBox "Missing CDC Detail !!!! "
   End If
   Exit Sub
End If

If mExchangeTypeID = 1 Then
   If CheckClientinRRNTab() = True Then
      MsgBox "No such Client Added in WHR Details"
      CmbLot.SetFocus
      Exit Sub
   End If
End If
If mExchangeTypeID = 2 Then 'NCDEX
   If Check1.Value = 0 Then
      If CheckISINinRRNTab() = True Then
         MsgBox "No such ISIN Added in RRN Details"
         CmbLot.SetFocus
         Exit Sub
      End If
   Else
      If CheckClientinRRNTab() = True Then
         MsgBox "No such Client Added in CDTF Details"
         CmbLot.SetFocus
         Exit Sub
      End If
      If CmbGCTDClient.Text = "" Then
         MsgBox "Blank Client-DP not Allowed!!!"
         CmbGCTDClient.SetFocus
         Exit Sub
      End If
   End If
End If
   
If Val(TxtGCTDNoofBags) > (Val(TxtGSLAvailableNofBags) - Val(TxtGSLPledgedNofBags)) Then
   MsgBox "Delivered bags are more than available unpledged no of bags !!!!"
End If
If CmdGCTDAdd2List.Caption = "Update List" Then
   r = MSHFlexGrid4.RowSel
   MSHFlexGrid4.TextMatrix(0, 4) = Val(MSHFlexGrid4.TextMatrix(0, 4)) - Val(MSHFlexGrid4.TextMatrix(r, 4))
   MSHFlexGrid4.TextMatrix(0, 5) = Val(MSHFlexGrid4.TextMatrix(0, 5)) - Val(MSHFlexGrid4.TextMatrix(r, 5))
   MSHFlexGrid4.TextMatrix(0, 6) = Val(MSHFlexGrid4.TextMatrix(0, 6)) - Val(MSHFlexGrid4.TextMatrix(r, 6))
   MSHFlexGrid4.TextMatrix(0, 7) = Val(MSHFlexGrid4.TextMatrix(0, 7)) - Val(MSHFlexGrid4.TextMatrix(r, 7))
   MSHFlexGrid4.TextMatrix(0, 8) = Val(MSHFlexGrid4.TextMatrix(0, 8)) - Val(MSHFlexGrid4.TextMatrix(r, 8))
   MSHFlexGrid4.TextMatrix(0, 9) = Val(MSHFlexGrid4.TextMatrix(0, 9)) - Val(MSHFlexGrid4.TextMatrix(r, 9))
   MSHFlexGrid4.TextMatrix(0, 10) = Val(MSHFlexGrid4.TextMatrix(0, 10)) - Val(MSHFlexGrid4.TextMatrix(r, 10))
   MSHFlexGrid4.TextMatrix(0, 11) = Val(MSHFlexGrid4.TextMatrix(0, 11)) - Val(MSHFlexGrid4.TextMatrix(r, 11))
Else
   r = MSHFlexGrid4.Rows - 1
   MSHFlexGrid4.Rows = MSHFlexGrid4.Rows + 1
End If

MSHFlexGrid4.TextMatrix(r, 1) = CmbGodownNo.Text
MSHFlexGrid4.TextMatrix(r, 2) = CmbStack.Text
MSHFlexGrid4.TextMatrix(r, 3) = CmbLot.Text
MSHFlexGrid4.TextMatrix(r, 4) = Val(TxtGSLAvailableNofBags)
MSHFlexGrid4.TextMatrix(r, 5) = Val(TxtGSLAvailableGrossWgt)
MSHFlexGrid4.TextMatrix(r, 6) = Val(TxtGCTDNoofBags)
MSHFlexGrid4.TextMatrix(r, 7) = Val(TxtGCTDGrossWt)
MSHFlexGrid4.TextMatrix(r, 8) = Val(TxtAdjWt)
MSHFlexGrid4.TextMatrix(r, 9) = Val(TxtGGunnyWeight)
MSHFlexGrid4.TextMatrix(r, 10) = Val(TxtGSLNetWeight)
MSHFlexGrid4.TextMatrix(r, 11) = Val(TxtGSLWeightGainLoss)
'If Check2.Value = 1 Or Check1.Value = 1 Then
'   MSHFlexGrid4.TextMatrix(r, 12) = CmbGCTDClient.Text
'Else
'   MSHFlexGrid4.TextMatrix(r, 12) = TxtGSLISINno
'End If


MSHFlexGrid4.Col = 7
MSHFlexGrid4.Row = r
MSHFlexGrid4.CellBackColor = &HFFC0C0


If CmbGCTDClient.Visible Then
   MSHFlexGrid4.TextMatrix(r, 12) = CmbGCTDClient.Text
Else
   MSHFlexGrid4.TextMatrix(r, 12) = TxtGSLISINno
End If


MSHFlexGrid4.TextMatrix(r, 13) = TxtGCTDremarks
MSHFlexGrid4.TextMatrix(r, 14) = mGSLID
MSHFlexGrid4.TextMatrix(r, 16) = IIf(IsNull(mGSLClientIDRow), "", mGSLClientIDRow)
MSHFlexGrid4.TextMatrix(r, 17) = Check2.Value

MSHFlexGrid4.TextMatrix(0, 4) = Val(MSHFlexGrid4.TextMatrix(0, 4)) + Val(MSHFlexGrid4.TextMatrix(r, 4))
MSHFlexGrid4.TextMatrix(0, 5) = Val(MSHFlexGrid4.TextMatrix(0, 5)) + Val(MSHFlexGrid4.TextMatrix(r, 5))
MSHFlexGrid4.TextMatrix(0, 6) = Val(MSHFlexGrid4.TextMatrix(0, 6)) + Val(MSHFlexGrid4.TextMatrix(r, 6))
MSHFlexGrid4.TextMatrix(0, 7) = Val(MSHFlexGrid4.TextMatrix(0, 7)) + Val(MSHFlexGrid4.TextMatrix(r, 7))
MSHFlexGrid4.TextMatrix(0, 8) = Val(MSHFlexGrid4.TextMatrix(0, 8)) + Val(MSHFlexGrid4.TextMatrix(r, 8))
MSHFlexGrid4.TextMatrix(0, 9) = Val(MSHFlexGrid4.TextMatrix(0, 9)) + Val(MSHFlexGrid4.TextMatrix(r, 9))
MSHFlexGrid4.TextMatrix(0, 10) = Val(MSHFlexGrid4.TextMatrix(0, 10)) + Val(MSHFlexGrid4.TextMatrix(r, 10))
MSHFlexGrid4.TextMatrix(0, 11) = Val(MSHFlexGrid4.TextMatrix(0, 11)) + Val(MSHFlexGrid4.TextMatrix(r, 11))


Call Clear_GCTDTabControl
If CmbGodownNo.Enabled = True Then
   CmbGodownNo.SetFocus
End If

End Sub
Function CheckISINinRRNTab() As Boolean
Dim RetVal As Boolean
Dim TxtGSLISINno_ As Variant
RetVal = False
If Check2.Value = 0 Then
   TxtGSLISINno_ = TxtGSLISINno
Else
    TxtGSLISINno_ = CmbGCTDClient.Text
End If

If MSHFlexGrid2.Rows > 3 Then
   For u = 2 To MSHFlexGrid2.Rows - 1
       If MSHFlexGrid2.TextMatrix(u, 2) <> "" Then
          If LCase(TxtGSLISINno_) = LCase(MSHFlexGrid2.TextMatrix(u, 2)) Then
             mRRNinGCTD = MSHFlexGrid2.TextMatrix(u, 12)
             RetVal = False
             Exit For
           Else
              RetVal = True
           End If
       End If
   Next
End If
CheckISINinRRNTab = RetVal
End Function
Function CheckClientinRRNTab() As Boolean
Dim RetVal As Boolean
Dim TxtGSLISINno_ As Variant
RetVal = True


If MSHFlexGrid2.Rows > 3 Then
   For u = 2 To MSHFlexGrid2.Rows - 1
       If Check2.Value = 0 Then
          If mGSLClientIDRow = IIf(IsNull(MSHFlexGrid2.TextMatrix(u, 11)), "", MSHFlexGrid2.TextMatrix(u, 11)) Then
             RetVal = False
             Exit For
          End If
       Else
          If MSHFlexGrid2.TextMatrix(u, 4) <> "" Then
             If CmbGCTDClient.Text = IIf(IsNull(MSHFlexGrid2.TextMatrix(u, 4)), "", MSHFlexGrid2.TextMatrix(u, 4)) Then
                mGSLClientIDRow = IIf(IsNull(MSHFlexGrid2.TextMatrix(u, 11)), "", MSHFlexGrid2.TextMatrix(u, 11))
                RetVal = False
                Exit For
             End If
          End If
       End If
   Next
End If
CheckClientinRRNTab = RetVal
End Function
Function CheckDuplicateGSL() As Boolean
Dim RetVal As Boolean
RetVal = False
If MSHFlexGrid4.Rows > 3 Then
   For u = 1 To MSHFlexGrid4.Rows - 1
       If MSHFlexGrid4.TextMatrix(u, 1) = CmbGodownNo.Text And _
          MSHFlexGrid4.TextMatrix(u, 2) = CmbStack.Text And _
          MSHFlexGrid4.TextMatrix(u, 3) = CmbLot.Text Then
          If LCase(CmdGCTDAdd2List.Caption) = "add to list" Then
             RetVal = True
             Exit For
          Else
             If u <> MSHFlexGrid4.RowSel Then
                RetVal = True
             End If
          End If
       End If
   Next
End If

CheckDuplicateGSL = RetVal
End Function

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
SSTab1.Tab = 0
Call TableTxn_CxTF_Details("Where TxnType = 'W' Order By SM_Prefix,CMPID,CxTFNo  ")
If RsTxn_CxTF_Details.RecordCount = 0 Then
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
Call InData


'Call AddCmd_Click
End Sub
Private Sub CmdTTDAdd2List_Click()
Dim r As Variant
If CmbExchangeTypeID.Text = "" Then
   MsgBox "Select Mode"
   CmbExchangeTypeID.SetFocus
   Exit Sub
End If

If CmbCMPID.Text = "" Then
   MsgBox "Select CMP"
   CmbCMPID.SetFocus
   Exit Sub
End If
If CmbCommodityID.Text = "" Then
   MsgBox "Select Commodity"
   CmbCommodityID.SetFocus
   Exit Sub
End If

If TxtTTDTruckNo = "" Then
   MsgBox "Blank Truck No  !!! "
   TxtTTDTruckNo.SetFocus
   Exit Sub
End If
If Val(TxtTTDTotalWt) <= 0 Then
   MsgBox "Invalid Weight !!! "
   TxtTTDTotalWt.SetFocus
   Exit Sub
End If
If Val(TxtTTDGrossWt) <= 0 Then
   MsgBox "Invalid Weight !!! "
   TxtTTDGrossWt.SetFocus
   Exit Sub
End If
tmp = Format(Val(TxtTTDTotalWt) - Val(TxtTTDTareWt), "#######0.000")
r = Format(Val(TxtTTDGrossWt), "#######0.000")

tmp = (Val(TxtTTDTotalWt) - Val(TxtTTDTareWt))
r = (Val(TxtTTDGrossWt))

If (tmp) <> r Then
   MsgBox "Invalid Weight !!! "
   Exit Sub
End If
If CmdTTDAdd2List.Caption = "Update List" Then
   r = MSHFlexGrid3.RowSel
   MSHFlexGrid3.TextMatrix(0, 3) = Val(MSHFlexGrid3.TextMatrix(0, 3)) - Val(MSHFlexGrid3.TextMatrix(r, 3))
   MSHFlexGrid3.TextMatrix(0, 4) = Val(MSHFlexGrid3.TextMatrix(0, 4)) - Val(MSHFlexGrid3.TextMatrix(r, 4))
   MSHFlexGrid3.TextMatrix(0, 5) = Val(MSHFlexGrid3.TextMatrix(0, 5)) - Val(MSHFlexGrid3.TextMatrix(r, 5))
Else
   r = MSHFlexGrid3.Rows - 1
   MSHFlexGrid3.Rows = MSHFlexGrid3.Rows + 1
End If
MSHFlexGrid3.TextMatrix(r, 1) = TxtTTDTruckNo
MSHFlexGrid3.TextMatrix(r, 2) = TxtWeighBridgNo
MSHFlexGrid3.TextMatrix(r, 3) = TxtTTDTotalWt
MSHFlexGrid3.TextMatrix(r, 4) = TxtTTDTareWt
MSHFlexGrid3.TextMatrix(r, 5) = TxtTTDGrossWt
MSHFlexGrid3.TextMatrix(r, 6) = txtTTDRemarks

MSHFlexGrid3.Col = 5
MSHFlexGrid3.Row = r
MSHFlexGrid3.CellBackColor = &HFFC0C0


MSHFlexGrid3.TextMatrix(0, 3) = Val(MSHFlexGrid3.TextMatrix(0, 3)) + Val(MSHFlexGrid3.TextMatrix(r, 3))
MSHFlexGrid3.TextMatrix(0, 4) = Val(MSHFlexGrid3.TextMatrix(0, 4)) + Val(MSHFlexGrid3.TextMatrix(r, 4))
MSHFlexGrid3.TextMatrix(0, 5) = Val(MSHFlexGrid3.TextMatrix(0, 5)) + Val(MSHFlexGrid3.TextMatrix(r, 5))
Call Clear_TTDTabControl
TxtTTDTruckNo.SetFocus
End Sub
Sub Clear_TTDTabControl()
TxtTTDTruckNo = ""
TxtWeighBridgNo = ""
TxtTTDTotalWt = ""
TxtTTDTareWt = ""
TxtTTDGrossWt = ""
txtTTDRemarks = ""
CmdTTDAdd2List.Caption = "Add to List"
TxtTTDGrossWt.Locked = True
End Sub

Private Sub MSHFlexGrid2_Click()
CmdAdd2ListRRN.Caption = "Update List"
TxtRRNNo.Locked = False
r = MSHFlexGrid2.RowSel
If r <= 0 Then Exit Sub

If Val(MSHFlexGrid2.TextMatrix(r, 9)) <> 0 Then
   MsgBox "You can't modify this row since corresponding GSL details are already selected !!!!!"
   Call Clear_RWCTabControl
   Exit Sub
End If


TxtRRNNo = MSHFlexGrid2.TextMatrix(r, 1)
TxtISINno = MSHFlexGrid2.TextMatrix(r, 2)
TxtClientID = MSHFlexGrid2.TextMatrix(r, 3)
TxtClientName = MSHFlexGrid2.TextMatrix(r, 4)
TxtDPID = MSHFlexGrid2.TextMatrix(r, 5)
TxtDPName = MSHFlexGrid2.TextMatrix(r, 6)
TxtRRNQty = MSHFlexGrid2.TextMatrix(r, 7)
TxtBalanceQty = MSHFlexGrid2.TextMatrix(r, 8)
TxtRRNRemarks = MSHFlexGrid2.TextMatrix(r, 10)
mClientIDRow = Val(MSHFlexGrid2.TextMatrix(r, 11))

End Sub

Private Sub MSHFlexGrid3_Click()
  
CmdTTDAdd2List.Caption = "Update List"
r = MSHFlexGrid3.RowSel
If r <= 1 Then Exit Sub
If MSHFlexGrid3.TextMatrix(r, 1) = "" Then
   CmdTTDAdd2List.Caption = "Add to List"
   Exit Sub
End If
TxtTTDTruckNo = MSHFlexGrid3.TextMatrix(r, 1)
TxtWeighBridgNo = MSHFlexGrid3.TextMatrix(r, 2)
TxtTTDTotalWt = MSHFlexGrid3.TextMatrix(r, 3)
TxtTTDTareWt = MSHFlexGrid3.TextMatrix(r, 4)
TxtTTDGrossWt = MSHFlexGrid3.TextMatrix(r, 5)
txtTTDRemarks = MSHFlexGrid3.TextMatrix(r, 6)
End Sub
Private Sub MSHFlexGrid4_Click()
Call Clear_GCTDTabControl
r = MSHFlexGrid4.RowSel
If r <= 1 Then Exit Sub
If MSHFlexGrid4.TextMatrix(r, 15) <> "" Then
   MsgBox "This godown is Closed you can't Edit !!!!!"
   Exit Sub
End If

If MSHFlexGrid4.TextMatrix(r, 1) = "" Then
   Call Clear_GCTDTabControl
   Exit Sub
End If

'CmdGCTDAdd2List.Caption = "Update List"  Error Found while  edit
'CmdGCTDAdd2List.Enabled = True

CmbGodownNo.Text = MSHFlexGrid4.TextMatrix(r, 1)

TxtGSLAvailableNofBags = MSHFlexGrid4.TextMatrix(r, 4)
TxtGSLAvailableGrossWgt = MSHFlexGrid4.TextMatrix(r, 5)
TxtGCTDNoofBags = MSHFlexGrid4.TextMatrix(r, 6)
TxtGCTDGrossWt = MSHFlexGrid4.TextMatrix(r, 7)
TxtAdjWt = MSHFlexGrid4.TextMatrix(r, 8)
TxtGGunnyWeight = MSHFlexGrid4.TextMatrix(r, 9)
TxtGSLNetWeight = MSHFlexGrid4.TextMatrix(r, 10)
TxtGSLWeightGainLoss = MSHFlexGrid4.TextMatrix(r, 11)
TxtGSLISINno = MSHFlexGrid4.TextMatrix(r, 12)
TxtGCTDremarks = MSHFlexGrid4.TextMatrix(r, 13)
mGSLID = MSHFlexGrid4.TextMatrix(r, 14)
'mGodownID = Val(MSHFlexGrid4.TextMatrix(r, 15))
'mRWCID = MSHFlexGrid4.TextMatrix(r, 16)
Check2.Value = Val(MSHFlexGrid4.TextMatrix(r, 17))

CmdGCTDAdd2List.Caption = "Update List"
CmdGCTDAdd2List.Enabled = True

If CmbGCTDClient.Visible = True Then
   CmbGCTDClient.Text = MSHFlexGrid4.TextMatrix(r, 12)
End If
Call CmbGodownNo_GotFocus
Call CmbGodownNo_LostFocus
CmbStack.Text = MSHFlexGrid4.TextMatrix(r, 2)
Call CmbStack_GotFocus
Call CmbStack_LostFocus
CmbLot.Text = MSHFlexGrid4.TextMatrix(r, 3)
Call CmbLot_GotFocus
Call CmbLot_LostFocus

If MSHFlexGrid4.TextMatrix(r, 0) <> "" Or mExchangeTypeID = 3 Then
   CmbGodownNo.Enabled = False
   CmbStack.Enabled = False
   CmbLot.Enabled = False
Else
   CmbGodownNo.Enabled = True
   CmbStack.Enabled = True
   CmbLot.Enabled = True
End If
If mExchangeTypeID = 3 Then
   Check2.Enabled = False
   Call Check2_Click
End If
CmbLot.Text = MSHFlexGrid4.TextMatrix(r, 3)
CmdGCTDAdd2List.Enabled = True
End Sub

Private Sub SaveCmd_Click()
Dim ans As Variant
Call UpdateRRNGridStock
Call UpdateErrorGrid

If CmbExchangeTypeID.Text = "" Then
   MsgBox "Select mode "
   CmbExchangeTypeID.SetFocus
   Exit Sub
End If

If CmbCMPID.Text = "" Then
   MsgBox "Select CMP "
   CmbCMPID.SetFocus
   Exit Sub
End If

If CmbAgent.Text = "" Then
   MsgBox "Select Agent "
   CmbAgent.SetFocus
   Exit Sub
End If

If CmbCommodityID.Text = "" Then
   MsgBox "Select Commodity "
   CmbCommodityID.SetFocus
   Exit Sub
End If

If MSHFlexGrid2.Rows <= 3 Then
   If mExchangeTypeID = 1 Then 'NON NCDEX
      MsgBox "Missing WHR Detail !!!! "
   ElseIf mExchangeTypeID = 2 Then 'NCDEX
      MsgBox "Missing RRN Detail !!!! "
   ElseIf mExchangeTypeID = 3 Then 'SPOT
      MsgBox "Missing CDC Detail !!!! "
   End If
   Exit Sub
End If
If MSHFlexGrid3.Rows <= 3 Then
   MsgBox "Missing Truck wise Detail !!!! "
   Exit Sub
End If
If MSHFlexGrid4.Rows <= 3 Then
   MsgBox "Missing Godown Stack Lot wise Detail !!!! "
   Exit Sub
End If

If Val(MSHFlexGrid3.TextMatrix(0, 5)) <> Val(MSHFlexGrid4.TextMatrix(0, 7)) Then
   MsgBox "Gross Weight in 'Truck Detail' tab is different from Gross Weight in 'GSL Detail' tab!!!! "
   Exit Sub
End If
If mExchangeTypeID = 1 Then
   If Val(MSHFlexGrid2.TextMatrix(0, 15)) <> Val(MSHFlexGrid4.TextMatrix(0, 6)) Then
      MsgBox "'Delivered Bags' in 'WHR No' tab is different from delivered bags in 'GSL Detail' tab!!!! "
      Exit Sub
   End If
End If
If mExchangeTypeID <= 2 Then
   SSTab1.Tab = 3
   If Val(MSHFlexGrid5.TextMatrix(0, 5)) > 0 Then
      tmp = "Excess deliveries are made against certain ISINs which are highlighted in yellow." & _
             "Please make appropriate adjustment entries against those ISINs in the 'GSL Details' tab." & _
             "Please note that these adjustment entries should be NEGATIVE entries."
      MsgBox tmp
      SSTab1.Tab = 3
      Exit Sub
   End If
   If Val(MSHFlexGrid5.TextMatrix(0, 4)) < 0 And Val(MSHFlexGrid5.TextMatrix(0, 6)) > 0 Then
      If Abs(Val(MSHFlexGrid5.TextMatrix(0, 4))) < Val(MSHFlexGrid5.TextMatrix(0, 6)) Then
         
         For i = 2 To MSHFlexGrid5.Rows - 1
             If Val(MSHFlexGrid5.TextMatrix(i, 6)) > 0 Then
                MSHFlexGrid5.Row = i
                MSHFlexGrid5.Col = 1
                MSHFlexGrid5.CellBackColor = vbRed
             End If
         Next
         
         
'         tmp = "You have made a net adjustment of " & Abs(Val(MSHFlexGrid5.TextMatrix(0, 4))) & " MT towards Excess " & _
'               "Delivery. Additionally you have also short delivered to the tune of " & Val(MSHFlexGrid5.TextMatrix(0, 6)) & " MT. " & _
'               "You need to adjust this Excess Delivery Adjustment by a corresponding Short Delivery Adjustment " & _
'               "Hence please make appropriate Adjustment Entry of " & Abs(Val(MSHFlexGrid5.TextMatrix(0, 4))) & "  MT " & _
'               "in those ISINs where short delivery are made."
'
         tmp = "Please make Adjustment Entry of " & Abs(Val(MSHFlexGrid5.TextMatrix(0, 4))) & "  MT " & _
               "in 'GSL Details' tab for some of those ISINs which are highlighted in red."
         MsgBox tmp
'         For i = 2 To MSHFlexGrid5.Rows - 1
'             If Val(MSHFlexGrid5.TextMatrix(i, 6)) > 0 Then
'                MSHFlexGrid5.Row = i
'                MSHFlexGrid5.Col = 6
'                MSHFlexGrid5.CellBackColor = vbRed
'                MSHFlexGrid5.Col = 1
'                MSHFlexGrid5.CellBackColor = vbRed
'             End If
'         Next
         'ans = MsgBox(tmp, vbOK)
                 'MsgBox "Create entries  for adjusted weight for  " & Abs(Val(MSHFlexGrid5.TextMatrix(0, 4))) & " MT  'GSL Detail' tab !!!!"
         
      Else
'         tmp = "You have made a net adjustment of " & Abs(Val(MSHFlexGrid5.TextMatrix(0, 4))) & " MT towards Excess " & _
'               "Delivery. Additionally you have also short delivered to the tune of " & Val(MSHFlexGrid5.TextMatrix(0, 6)) & " MT. " & _
'               "You need to adjust this Excess Delivery Adjustment by a corresponding Short Delivery Adjustment " & _
'               "Hence please make appropriate Adjustment Entry of " & Abs(Val(MSHFlexGrid5.TextMatrix(0, 6))) & "  MT " & _
'               "in those ISINs where short delivery are made."
         
         tmp = "Please make Adjustment Entry of " & Abs(Val(MSHFlexGrid5.TextMatrix(0, 6))) & "  MT " & _
               "in 'GSL Details' tab for some of those ISINs which are highlighted in red."
         MsgBox tmp
         
         'MsgBox "Create entries  for adjusted weight for  " & Abs(Val(MSHFlexGrid5.TextMatrix(0, 6))) & " MT  'GSL Detail' tab !!!!"
      End If
      Exit Sub
   End If
     
End If
If Val(MSHFlexGrid2.TextMatrix(0, 9)) <> Val(MSHFlexGrid4.TextMatrix(0, 10)) Then
   If mExchangeTypeID = 1 Then
      MsgBox "'Delivered Qty' in 'WHR No' tab is different from 'Electronic Weight' in 'GSL Detail' tab!!!! "
   ElseIf mExchangeTypeID = 2 Then
     If Check1.Value = 0 Then
        MsgBox "'Delivered Qty' in 'RRN' tab is different from 'Electronic Weight' in 'GSL Detail' tab!!!! "
     ElseIf Check1.Value = 1 Then
        MsgBox "Delivered Quantity in 'Client Detail tab' is not equal to Electronic Weight in 'Delivery Lot(G~S~L) Details' "
     End If
   ElseIf mExchangeTypeID = 3 Then
        MsgBox "'Delivered Qty' in 'CDC No' tab is different from 'Electronic Weight' in 'GSL Detail' tab!!!! "
   End If
   Exit Sub
End If

'MsgBox ("Since no Edit or Delete option available for CITF, Data once saved can't be modified therefor go through all the tabs carfully and ensure correctness of data entered by you !!! ",vbYesNo
tmp = MsgBox("Data once saved can't be modified. Are you sure you want to save !!!!", vbYesNo)
If tmp = vbNo Then
   Exit Sub
End If

If Edit_Flg = "A" Then
   RsTxn_CxTF_Details.AddNew
   RsTxn_CxTF_Details!CxTFNo = GetCxTFNo
   RsTxn_CxTF_Details!G_UID = GetGUID
   TxtCxTFNo.Text = RsTxn_CxTF_Details!CxTFNo
Else
   RsTxn_CxTF_Details.MoveFirst
   RsTxn_CxTF_Details.Find " CxTFNo= " & TxtCxTFNo.Text
    'Call DeleteChildTableDetail
End If

RsTxn_CxTF_Details!CxTFNo = TxtCxTFNo.Text
RsTxn_CxTF_Details!ManualCDTFNo = TxtManualCDTFNo
RsTxn_CxTF_Details!CxTFDate = TxtCxTFDate.Value
RsTxn_CxTF_Details!CMPID = mCMPID
RsTxn_CxTF_Details!CommodityID = mCommodityId
RsTxn_CxTF_Details!ExchangeTypeID = mExchangeTypeID
RsTxn_CxTF_Details!BookedSpaceFlag = Check1.Value
'RsTxn_CxTF_Details!Flag = TxtFlag.Text
RsTxn_CxTF_Details!Agent = TxtAgent.Text
RsTxn_CxTF_Details!AgentID = mAgentID

'RsTxn_CxTF_Details!Remark = TxtRemark.Text
RsTxn_CxTF_Details!SM_Prefix = TxtPreFix

RsTxn_CxTF_Details!TxnType = "W"
If mContractID <> "" Then
   RsTxn_CxTF_Details!ContractID = mContractID
End If
RsTxn_CxTF_Details.Update
Call SaveClient
Call SaveTruck
Call SaveGSL
SaveCmd.Enabled = False
Check1.Value = 0
Call InData

End Sub
Private Function GetCxTFNo() As Double
Dim RetVal As Double
tmp = "Select max(CxTFNo) from Txn_CxTF_Details Where  sm_prefix = '" & TxtPreFix & "' And TxnType = 'W'  "
Call TmpTable
RetVal = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
RetVal = RetVal + 1
GetCxTFNo = RetVal
End Function
Private Sub SaveClient()
Call TableTxn_CxTF_ClientDetail(" Where  sm_prefix = '" & TxtPreFix & "' And CxTFNo = " & TxtCxTFNo & " And TxnType = 'W' ")
If mExchangeTypeID = 2 Then
   Call TableMst_RRN
End If
If mExchangeTypeID = 1 Then
   Call TableTxn_WHRDetail(" where  sm_prefix = '" & TxtPreFix & "'")
End If
If mExchangeTypeID = 3 Then
   Call TableTxn_CDC_GSLDetail(" Where  sm_prefix = '" & TxtPreFix & "'")
End If

For i = 2 To MSHFlexGrid2.Rows - 1
    If MSHFlexGrid2.TextMatrix(i, 1) <> "" Then
       RsTxn_CxTF_ClientDetail.AddNew
       RsTxn_CxTF_ClientDetail!ClientDetailID = getCxtf_ClientMax
       RsTxn_CxTF_ClientDetail!G_UID = GetGUID
       RsTxn_CxTF_ClientDetail!CxTFNo = TxtCxTFNo
       RsTxn_CxTF_ClientDetail!ClientIDRow = MSHFlexGrid2.TextMatrix(i, 11)
       RsTxn_CxTF_ClientDetail!TxnType = "W"
       RsTxn_CxTF_ClientDetail!RRNID = MSHFlexGrid2.TextMatrix(i, 12)
       RsTxn_CxTF_ClientDetail!AdjRRNQty = MSHFlexGrid2.TextMatrix(i, 9)
       RsTxn_CxTF_ClientDetail!AdjRRNNoofBags = Val(MSHFlexGrid2.TextMatrix(i, 15))
       
       RsTxn_CxTF_ClientDetail!SM_Prefix = TxtPreFix
       RsTxn_CxTF_ClientDetail.Update
       RsTxn_CxTF_ClientDetail.Requery
       If mExchangeTypeID = 1 Then
          If Check1.Value = 0 Then
             RsTxn_WHRDetail.MoveFirst
             RsTxn_WHRDetail.Find "WHRNo = " & Val(MSHFlexGrid2.TextMatrix(i, 12))
             If Not RsTxn_WHRDetail.EOF Then
                RsTxn_WHRDetail!DeliveredQuantity = RsTxn_WHRDetail!DeliveredQuantity + Val(MSHFlexGrid2.TextMatrix(i, 9))
                RsTxn_WHRDetail!DeliveredNoOfBags = RsTxn_WHRDetail!DeliveredNoOfBags + Val(MSHFlexGrid2.TextMatrix(i, 15))
                
                RsTxn_WHRDetail.Update
                RsTxn_WHRDetail.Requery
             End If
          End If
       End If
       If mExchangeTypeID = 2 Then
          If Check1.Value = 0 Then
             RsMst_RRN.MoveFirst
             RsMst_RRN.Find "RRNID = " & Val(MSHFlexGrid2.TextMatrix(i, 12))
             If Not RsMst_RRN.EOF Then
                RsMst_RRN!DeliveredQty = RsMst_RRN!DeliveredQty + Val(MSHFlexGrid2.TextMatrix(i, 9))
                RsMst_RRN.Update
                RsMst_RRN.Requery
             End If
          End If
       End If
    
       If mExchangeTypeID = 3 Then
          If Check1.Value = 0 Then
             RsTxn_CDC_GSLDetail.MoveFirst
             RsTxn_CDC_GSLDetail.Find "CDCNo = " & Val(MSHFlexGrid2.TextMatrix(i, 12))
             If Not RsTxn_CDC_GSLDetail.EOF Then
                RsTxn_CDC_GSLDetail!DeliveredQty = RsTxn_CDC_GSLDetail!DeliveredQty + Val(MSHFlexGrid2.TextMatrix(i, 9))
                RsTxn_CDC_GSLDetail!Flag = "W"
                RsTxn_CDC_GSLDetail.Update
                RsTxn_CDC_GSLDetail.Requery
             End If
          End If
       End If
    
    End If
Next


End Sub
Function getCxtf_ClientMax() As Double
Dim RetVal As Double
RetVal = 0
tmp = "select max(ClientDetailID) from Txn_CxTF_ClientDetail where sm_prefix = '" & TxtPreFix & "'"
Call TmpTable
If TmpRs.RecordCount > 0 Then
   RetVal = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
End If
RetVal = RetVal + 1
TmpRs.Close
getCxtf_ClientMax = RetVal

End Function
Private Sub SaveTruck()
Call TableTxn_CxTF_TruckDetail(" Where sm_prefix = '" & TxtPreFix & "' And CxTFNo = " & TxtCxTFNo & " And TxnType = 'W'")
For i = 2 To MSHFlexGrid3.Rows - 1
    If MSHFlexGrid3.TextMatrix(i, 1) <> "" Then
       RsTxn_CxTF_TruckDetail.AddNew
       RsTxn_CxTF_TruckDetail!TruckDetailID = getCxtf_TruckMax
       RsTxn_CxTF_TruckDetail!G_UID = GetGUID
       RsTxn_CxTF_TruckDetail!CxTFNo = TxtCxTFNo
       RsTxn_CxTF_TruckDetail!TruckNo = MSHFlexGrid3.TextMatrix(i, 1)
       RsTxn_CxTF_TruckDetail!WeighBridgeSlipno = MSHFlexGrid3.TextMatrix(i, 2)
       RsTxn_CxTF_TruckDetail!TotalWeight = MSHFlexGrid3.TextMatrix(i, 3)
       RsTxn_CxTF_TruckDetail!TareWeight = Val(MSHFlexGrid3.TextMatrix(i, 4))
       RsTxn_CxTF_TruckDetail!GrossWeight = MSHFlexGrid3.TextMatrix(i, 5)
       RsTxn_CxTF_TruckDetail!Remarks = MSHFlexGrid3.TextMatrix(i, 6)
       RsTxn_CxTF_TruckDetail!TxnType = "W"
       RsTxn_CxTF_TruckDetail!SM_Prefix = TxtPreFix
       RsTxn_CxTF_TruckDetail.Update
       RsTxn_CxTF_TruckDetail.Requery
    End If
Next
End Sub
Function getCxtf_TruckMax() As Double
Dim RetVal As Double
RetVal = 0
tmp = "select max(TruckDetailID) from Txn_CxTF_TruckDetail where sm_prefix = '" & TxtPreFix & "'"
Call TmpTable
If TmpRs.RecordCount > 0 Then
   RetVal = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
End If
RetVal = RetVal + 1
TmpRs.Close
getCxtf_TruckMax = RetVal

End Function
Private Sub SaveGSL()
Dim PQty, TQty As Variant

Call TableTxn_CxTF_GSL(" Where sm_prefix = '" & TxtPreFix & "' And CxTFNo = " & TxtCxTFNo & " And TxnType = 'W' ")
Call TableMst_GSL(" where sm_prefix = '" & TxtPreFix & "'")
For i = 2 To MSHFlexGrid4.Rows - 1
    If MSHFlexGrid4.TextMatrix(i, 1) <> "" Then
       RsTxn_CxTF_GSL.AddNew
       RsTxn_CxTF_GSL!GSLDetailID = getCxtf_GSLMax
       RsTxn_CxTF_GSL!G_UID = GetGUID
       RsMst_GSL.MoveFirst
       RsMst_GSL.Find "GSLID = " & Val(MSHFlexGrid4.TextMatrix(i, 14))
       
       RsTxn_CxTF_GSL!CxTFNo = TxtCxTFNo
       RsTxn_CxTF_GSL!NoofBags = Val(MSHFlexGrid4.TextMatrix(i, 6))
       RsTxn_CxTF_GSL!GrossWeight = Val(MSHFlexGrid4.TextMatrix(i, 7))
       RsTxn_CxTF_GSL!AdjNCMSLWeight = Val(MSHFlexGrid4.TextMatrix(i, 8))
       'RsTxn_CxTF_GSL!AdjNClientWeight = Val(MSHFlexGrid4.TextMatrix(i, 8))
       RsTxn_CxTF_GSL!GunnyWeight = Val(MSHFlexGrid4.TextMatrix(i, 9))
       RsTxn_CxTF_GSL!DematWeight = Val(MSHFlexGrid4.TextMatrix(i, 10))
       RsTxn_CxTF_GSL!WeightGainLoss = Val(MSHFlexGrid4.TextMatrix(i, 11)) 'Weight Gain Or Loss
'       If mExchangeTypeID = 2 Then
'          RsTxn_CxTF_GSL!ClientIDRow = MSHFlexGrid4.TextMatrix(i, 12)
'       End If
       
       RsTxn_CxTF_GSL!Remarks = MSHFlexGrid4.TextMatrix(i, 13)
       RsTxn_CxTF_GSL!GSLID = MSHFlexGrid4.TextMatrix(i, 14)
       RsTxn_CxTF_GSL!TxnType = "W"
       RsTxn_CxTF_GSL!SM_Prefix = TxtPreFix
       RsTxn_CxTF_GSL.Update
       RsTxn_CxTF_GSL.Requery
       If Not RsMst_GSL.EOF Then
          RsMst_GSL!CDTFBags = IIf(IsNull(RsMst_GSL!CDTFBags), 0, RsMst_GSL!CDTFBags) + Val(MSHFlexGrid4.TextMatrix(i, 6))
          RsMst_GSL!CDTFWeight = IIf(IsNull(RsMst_GSL!CDTFWeight), 0, RsMst_GSL!CDTFWeight) + Val(MSHFlexGrid4.TextMatrix(i, 10))
          RsMst_GSL!BalanceBags = IIf(IsNull(RsMst_GSL!BalanceBags), 0, RsMst_GSL!BalanceBags) - Val(MSHFlexGrid4.TextMatrix(i, 6))
          RsMst_GSL!BalanceWeight = IIf(IsNull(RsMst_GSL!BalanceWeight), 0, RsMst_GSL!BalanceWeight) - Val(MSHFlexGrid4.TextMatrix(i, 10))
          
          RsMst_GSL!WithdrawalAdjWt = IIf(IsNull(RsMst_GSL!WithdrawalAdjWt), 0, RsMst_GSL!WithdrawalAdjWt) + Val(MSHFlexGrid4.TextMatrix(i, 8))
          RsMst_GSL!Flag1 = RsMst_GSL!Flag
          If Val(RsMst_GSL!BalanceBags) = 0 Then
             RsMst_GSL!Flag1 = RsMst_GSL!Flag
             RsMst_GSL!Flag = "E"
          End If
          RsMst_GSL.Update
       End If

    End If
Next
End Sub
Private Sub InDataTruck()
Call TableTxn_CxTF_TruckDetail("  Where sm_prefix = '" & TxtPreFix & "' And CxTFNo = " & TxtCxTFNo & " And TxnType = 'W'")
MSHFlexGrid3.Rows = MSHFlexGrid3.Rows + RsTxn_CxTF_TruckDetail.RecordCount
For i = 2 To MSHFlexGrid3.Rows - 2
    MSHFlexGrid3.TextMatrix(i, 0) = RsTxn_CxTF_TruckDetail!TruckDetailID
    MSHFlexGrid3.TextMatrix(i, 1) = RsTxn_CxTF_TruckDetail!TruckNo
    MSHFlexGrid3.TextMatrix(i, 2) = RsTxn_CxTF_TruckDetail!WeighBridgeSlipno
    MSHFlexGrid3.TextMatrix(i, 3) = RsTxn_CxTF_TruckDetail!TotalWeight
    MSHFlexGrid3.TextMatrix(i, 4) = RsTxn_CxTF_TruckDetail!TareWeight
    MSHFlexGrid3.TextMatrix(i, 5) = RsTxn_CxTF_TruckDetail!GrossWeight
    MSHFlexGrid3.TextMatrix(i, 6) = RsTxn_CxTF_TruckDetail!Remarks
   
   
    MSHFlexGrid3.Col = 5
    MSHFlexGrid3.Row = i
    MSHFlexGrid3.CellBackColor = &HFFC0C0
    
    MSHFlexGrid3.TextMatrix(0, 3) = Val(MSHFlexGrid3.TextMatrix(0, 3)) + RsTxn_CxTF_TruckDetail!TotalWeight
    MSHFlexGrid3.TextMatrix(0, 4) = Val(MSHFlexGrid3.TextMatrix(0, 4)) + RsTxn_CxTF_TruckDetail!TareWeight
    MSHFlexGrid3.TextMatrix(0, 5) = Val(MSHFlexGrid3.TextMatrix(0, 5)) + RsTxn_CxTF_TruckDetail!GrossWeight
    RsTxn_CxTF_TruckDetail.MoveNext
Next
End Sub
Function getCxtf_GSLMax() As Double
Dim RetVal As Double
RetVal = 0
tmp = "select max(GSLDetailID) from Txn_CxTF_GSL where sm_prefix = '" & TxtPreFix & "'"
Call TmpTable
If TmpRs.RecordCount > 0 Then
   RetVal = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
End If
RetVal = RetVal + 1
TmpRs.Close
getCxtf_GSLMax = RetVal

End Function
Private Sub InDataClient()
Dim ans As Variant
Call TableTxn_CxTF_ClientDetail(" Where sm_prefix = '" & TxtPreFix & "' And CxTFNo = " & TxtCxTFNo & " And TxnType = 'W' ")
If mExchangeTypeID = 1 Then
   Call TableTxn_WHRDetail(" where sm_prefix = '" & TxtPreFix & "'")
End If
If mExchangeTypeID = 2 Then
   Call TableMst_RRN
End If
If mExchangeTypeID = 3 Then
   Call TableTxn_CDC_GSLDetail(" where sm_prefix = '" & TxtPreFix & "'")
   Call TableMst_GSL(" where sm_prefix = '" & TxtPreFix & "'")
End If

MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + RsTxn_CxTF_ClientDetail.RecordCount
For i = 2 To MSHFlexGrid2.Rows - 2
    MSHFlexGrid2.TextMatrix(i, 0) = RsTxn_CxTF_ClientDetail!ClientDetailID
    Call GetRWCDetails(i, mExchangeTypeID, RsTxn_CxTF_ClientDetail!RRNID)
    ans = GetClientNameByRow(RsTxn_CxTF_ClientDetail!ClientIDRow)
    MSHFlexGrid2.TextMatrix(i, 3) = ans(0)
    MSHFlexGrid2.TextMatrix(i, 4) = ans(1) 'TmpRs!ClientName
    MSHFlexGrid2.TextMatrix(i, 5) = ans(2) 'TmpRs!DepositoryID
    MSHFlexGrid2.TextMatrix(i, 6) = ans(3) 'TmpRs!DepositoryName
    If mExchangeTypeID = 1 Then
       RsTxn_WHRDetail.Find "WHRNo = " & RsTxn_CxTF_ClientDetail!RRNID
       MSHFlexGrid2.TextMatrix(i, 7) = Val(RsTxn_WHRDetail!WHRQuantity)
       MSHFlexGrid2.TextMatrix(i, 8) = Format(Val(RsTxn_WHRDetail!WHRQuantity) - Val(RsTxn_WHRDetail!DeliveredQuantity), "#######0.000")
       MSHFlexGrid2.TextMatrix(i, 9) = RsTxn_CxTF_ClientDetail!AdjRRNQty
       
       MSHFlexGrid2.TextMatrix(i, 13) = Val(RsTxn_WHRDetail!WHRNoOfBags)
       MSHFlexGrid2.TextMatrix(i, 14) = Format(Val(RsTxn_WHRDetail!WHRNoOfBags) - Val(RsTxn_WHRDetail!DeliveredNoOfBags), "#######0.000")
       MSHFlexGrid2.TextMatrix(i, 15) = RsTxn_CxTF_ClientDetail!AdjRRNNoofBags
    ElseIf mExchangeTypeID = 2 Then
       If Check1.Value = 0 Then
          RsMst_RRN.Find "RRNID = " & RsTxn_CxTF_ClientDetail!RRNID
          MSHFlexGrid2.TextMatrix(i, 7) = RsMst_RRN!RRNQty
          MSHFlexGrid2.TextMatrix(i, 8) = Format(RsMst_RRN!RRNQty - RsMst_RRN!DeliveredQty, "#######0.000")
          MSHFlexGrid2.TextMatrix(i, 9) = RsTxn_CxTF_ClientDetail!AdjRRNQty
       ElseIf Check1.Value = 1 Then
          
            MSHFlexGrid2.TextMatrix(i, 9) = RsTxn_CxTF_ClientDetail!AdjRRNQty
            MSHFlexGrid2.TextMatrix(i, 15) = RsTxn_CxTF_ClientDetail!AdjRRNNoofBags
       End If
    ElseIf mExchangeTypeID = 3 Then
       RsTxn_CDC_GSLDetail.Find "CDCNo = " & RsTxn_CxTF_ClientDetail!RRNID
       RsMst_GSL.MoveFirst
       RsMst_GSL.Find "GSLID = " & RsTxn_CDC_GSLDetail!GSLID
       MSHFlexGrid2.TextMatrix(i, 7) = Val(RsMst_GSL!DepositeWeight)
       MSHFlexGrid2.TextMatrix(i, 8) = Format(Val(RsMst_GSL!DepositeWeight) - Val(RsMst_GSL!CDTFWeight), "#######0.000")
       MSHFlexGrid2.TextMatrix(i, 9) = Val(RsMst_GSL!CDTFWeight)
    End If
    MSHFlexGrid2.TextMatrix(i, 11) = RsTxn_CxTF_ClientDetail!ClientIDRow
    MSHFlexGrid2.TextMatrix(i, 12) = RsTxn_CxTF_ClientDetail!RRNID
    MSHFlexGrid2.TextMatrix(0, 7) = Val(MSHFlexGrid2.TextMatrix(0, 7)) + Val(MSHFlexGrid2.TextMatrix(i, 7))
    MSHFlexGrid2.TextMatrix(0, 8) = Format(Val(MSHFlexGrid2.TextMatrix(0, 8)) + Val(MSHFlexGrid2.TextMatrix(i, 8)), "#######0.000")
    MSHFlexGrid2.TextMatrix(0, 9) = Val(MSHFlexGrid2.TextMatrix(0, 9)) + Val(MSHFlexGrid2.TextMatrix(i, 9))
    RsTxn_CxTF_ClientDetail.MoveNext
Next
'MsgBox TmpRs.State

End Sub
Private Sub GetRWCDetails(i_, mExchangeTypeID_, RWCID_)
'Dim mRWCID_ As Double
If mExchangeTypeID_ = 1 Then
   Call TableTxn_WHRDetail("Where sm_prefix = '" & TxtPreFix & "' And WHRNo = " & RWCID_ & " ")
   
   mRWCID_ = mCMPID & "-" & (RsTxn_WHRDetail!WHRNo - mCMPID * 1000000)
   
   mRWCQty_ = RsTxn_WHRDetail!WHRQuantity
ElseIf mExchangeTypeID_ = 2 And Check1.Value = 0 Then
   Call TableMst_RRN("Where RRNID = " & RWCID_ & " ")
   mRWCID_ = RsMst_RRN!RRNNo
   mISINNo_ = RsMst_RRN!ISINID
   mRWCQty_ = RsMst_RRN!RRNQty
ElseIf mExchangeTypeID_ = 3 Then
   Call TableTxn_CDC_GSLDetail("Where sm_prefix = '" & TxtPreFix & "' And CDCNo = " & RWCID_ & "")
   mRWCID_ = RsTxn_CDC_GSLDetail!CDCNo
   mRWCQty_ = RsTxn_CDC_GSLDetail!StorageAmount
End If
MSHFlexGrid2.TextMatrix(i_, 1) = mRWCID_
MSHFlexGrid2.TextMatrix(i_, 2) = mISINNo_
MSHFlexGrid2.TextMatrix(i_, 7) = mRWCQty_
'MSHFlexGrid2.TextMatrix(x, 8) = mRWCID_
'MSHFlexGrid2.TextMatrix(x, 9) = mRWCID_
End Sub

Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   If RsTxn_CxTF_Details.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   Call FirstCmd_Click
   Exit Sub
End If

Frame0.Height = Me.Height - 350
Frame0.Left = Picture1.Left - 100
Frame0.Top = Picture1.Top - 100
Frame0.Width = Me.Width - 150
Frame0.Visible = True
MSHFlexGrid1.Clear
MSHFlexGrid1.Cols = 3
MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 1000
RsTxn_CxTF_Details.MoveFirst
Call Grid_Head
MSHFlexGrid1.Rows = RsTxn_CxTF_Details.RecordCount + 1
MSHFlexGrid1.ColWidth(0) = 1000
For i = 1 To RsTxn_CxTF_Details.RecordCount
    MSHFlexGrid1.TextMatrix(i, 0) = RsTxn_CxTF_Details!CxTFNo
    MSHFlexGrid1.TextMatrix(i, 1) = RsTxn_CxTF_Details!CxTFDate
    MSHFlexGrid1.TextMatrix(i, 2) = GetCMPName(RsTxn_CxTF_Details!CMPID)
    MSHFlexGrid1.TextMatrix(i, 3) = GetCommodityName(RsTxn_CxTF_Details!CommodityID)
    MSHFlexGrid1.TextMatrix(i, 4) = GetExchangeName(RsTxn_CxTF_Details!ExchangeTypeID)
    MSHFlexGrid1.TextMatrix(i, 5) = RsTxn_CxTF_Details!SM_Prefix
    MSHFlexGrid1.TextMatrix(i, 6) = IIf(IsNull(RsTxn_CxTF_Details!Flag), "", RsTxn_CxTF_Details!Flag)
    MSHFlexGrid1.TextMatrix(i, 7) = IIf(IsNull(RsTxn_CxTF_Details!Agent), "", RsTxn_CxTF_Details!Agent)
    MSHFlexGrid1.TextMatrix(i, 8) = IIf(IsNull(RsTxn_CxTF_Details!Remark), "", RsTxn_CxTF_Details!Remark)
    MSHFlexGrid1.TextMatrix(i, 9) = "Withdrawal"
    RsTxn_CxTF_Details.MoveNext
Next
End Sub
Private Sub Grid_Head()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 11
MSHFlexGrid1.TextMatrix(0, 0) = "CITF No"
MSHFlexGrid1.TextMatrix(0, 1) = "CITF Date"
MSHFlexGrid1.TextMatrix(0, 2) = "CMP"
MSHFlexGrid1.TextMatrix(0, 3) = "Commodity"
MSHFlexGrid1.TextMatrix(0, 4) = "Exchange Type"
MSHFlexGrid1.TextMatrix(0, 5) = "Prefix"
MSHFlexGrid1.TextMatrix(0, 6) = "Flag"
MSHFlexGrid1.TextMatrix(0, 7) = "Agent"
MSHFlexGrid1.TextMatrix(0, 8) = "Remark"
MSHFlexGrid1.TextMatrix(0, 9) = "Txn Type"
MSHFlexGrid1.TextMatrix(0, 10) = "Booked Space ID"

End Sub
Private Sub MSHFlexGrid1_Click()
Dim i As Variant
i = MSHFlexGrid1.RowSel
If i <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(i, 0) = "" Then Exit Sub
RsTxn_CxTF_Details.MoveFirst
RsTxn_CxTF_Details.Find "sm_prefix  = '" & MSHFlexGrid1.TextMatrix(i, 5) & "'"
If Not RsTxn_CxTF_Details.EOF Then
   RsTxn_CxTF_Details.Find "CxTFNo = " & MSHFlexGrid1.TextMatrix(i, 0)
   Call InData
   Frame0.Visible = False
End If
End Sub
Private Sub CmdExcel_Click()
Call Gen_Create_Xls("Txn_CxTF_Details.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & "Txn_CxTF_Details.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For i = 0 To MSHFlexGrid1.Rows - 1
    mRow = i + 1 ' + 6
    For c = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(i, c), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(i, c)) Then
              oWS.Cells(mRow, c + 1) = Format(MSHFlexGrid1.TextMatrix(i, c), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(i, c)
           End If
        Else
           oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(i, c)
        End If
    Next
Next
oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub

Private Sub TxtAdjWt_LostFocus()
If TxtAdjWt = "" Then
   TxtAdjWt = Val(TxtAdjWt)
Else
    If IsNumeric(TxtAdjWt) = False Then
       MsgBox "Invalid Number format !!!"
       TxtAdjWt.SetFocus
       Exit Sub
    End If
End If
Call CalDmatWt
End Sub
Private Sub TxtCxTFDate_LostFocus()
If TxtCxTFDate.Value > Date Then
   MsgBox "Future Date is not Allowed !!!! "
   TxtCxTFDate.SetFocus
End If
End Sub
Private Sub TxtGCTDGrossWt_LostFocus()
If TxtGCTDGrossWt = "" Then Exit Sub

If Not IsNumeric(TxtGCTDGrossWt) Or Val(TxtGCTDGrossWt) < 0 Then
   MsgBox "Invalid Number format !!!"
   TxtGCTDGrossWt.SetFocus
   Exit Sub
End If
Call CalDmatWt

End Sub
Private Sub TxtGCTDNoofBags_LostFocus()
If TxtGCTDNoofBags = "" Then Exit Sub
If Not IsNumeric(TxtGCTDNoofBags) Or Val(TxtGCTDNoofBags) < 0 Then
   MsgBox "Invalid Number Format!!!"
   TxtGCTDNoofBags.SetFocus
   Exit Sub
End If
If Val(TxtGSLAvailableNofBags) < Val(TxtGCTDNoofBags) Then
   MsgBox "Deliver No Of Bags Must Be Less Than Available No Of Bags!!!"
   TxtGCTDNoofBags.SetFocus
   Exit Sub
End If
Call CalDmatWt
End Sub
Private Sub CalDmatWt()

TxtGGunnyWeight = Format((Val(TxtGCTDNoofBags) * Val(TxtCGunnyWt)) / 1000, "#######0.000")
TxtGSLNetWeight = Format(Val(TxtGCTDGrossWt) + Val(TxtAdjWt) - (Val(TxtGGunnyWeight)), "#######0.000")
If Val(TxtGSLAvailableNofBags) = Val(TxtGCTDNoofBags) Then
   TxtGSLWeightGainLoss = Format((Val(TxtGCTDGrossWt) - Val(TxtGGunnyWeight)) - (Val(TxtGSLAvailableGrossWgt)), "#######0.000")
Else
   TxtGSLWeightGainLoss = 0
End If
End Sub

Private Sub TxtRecoverClient_LostFocus()
If TxtRecoverClient = "" Then Exit Sub
If IsNumeric(TxtRecoverClient) = False Or Val(TxtRecoverClient) < 0 Then
   MsgBox "Invalid Numeric Format!!!"
   TxtRecoverClient.SetFocus
   Exit Sub
End If
End Sub

Private Sub TxtRecoverNCMSL_LostFocus()
If TxtRecoverNCMSL = "" Then Exit Sub
If IsNumeric(TxtRecoverNCMSL) = False Or Val(TxtRecoverNCMSL) < 0 Then
   MsgBox "Invalid Numeric Format!!!"
   TxtRecoverNCMSL.SetFocus
   Exit Sub
End If
End Sub

Private Sub TxtRRNNo_KeyPress(KeyAscii As Integer)
tmp = KeyAscii
If KeyAscii = 14 Then
   SSTab1.Tab = 1
   TxtTTDTruckNo.SetFocus
End If
End Sub

Private Sub TxtRRNNo_LostFocus()
If CmbExchangeTypeID.Text = "" Then
   MsgBox "Please Select Mode!!!"
   CmbExchangeTypeID.SetFocus
   Exit Sub
End If
If TxtRRNNo = "" Then
   Exit Sub
End If
If mExchangeTypeID = 1 Then 'NON NCDEX
   Dim strSplitString() As String
   strSplitString = Split(TxtRRNNo.Text, "-")
   tmp = UBound(strSplitString)
   If tmp <> 1 Then
      MsgBox "Invalid WHR No Selection"
      TxtRRNNo.SetFocus
      Exit Sub
   End If
   tmp = (Val(strSplitString(0)) * 1000000) + Val(strSplitString(1))
   Call WHRDetails(mCMPID, mCommodityId, tmp)
End If
If mExchangeTypeID = 2 Then 'NCDEX
   Call RRNDetails(mCMPID, mCommodityId, TxtRRNNo.Text)
End If
If mExchangeTypeID = 3 Then 'SPOT
   Call CDCDetails(mCMPID, mCommodityId, TxtRRNNo.Text)
End If
End Sub
Private Sub RRNDetails(mCMPID_, mCommodityID_, mTxtRRNNo_)

If mExchangeTypeID = 0 Or IsNull(mExchangeTypeID) Then Exit Sub
If mCMPID_ = 0 Or IsNull(mCMPID_) Then Exit Sub
If mCommodityID_ = 0 Or IsNull(mCommodityID_) Then Exit Sub
If mTxtRRNNo_ = "" Then Exit Sub

Call TableMst_RRN("Where RRNNo = '" & mTxtRRNNo_ & "'")
If RsMst_RRN.RecordCount = 0 Then
   MsgBox "No Record Found!!!"
   TxtRRNNo.SetFocus
   Exit Sub
End If
RsMst_RRN.MoveFirst
RsMst_RRN.Find "RRNNo = '" & TxtRRNNo & "'"
If Not RsMst_RRN.EOF Then
   TxtISINno = IIf(IsNull(RsMst_RRN!ISINID), "", RsMst_RRN!ISINID)
   Call TableMst_ISIN("Where ISINID = '" & TxtISINno & "'")
   If Not RsMst_ISIN.EOF Then
      If mCMPID_ <> RsMst_ISIN!CMPID Then
         MsgBox "ISIN CMP And Above CMP Mismatch!!!"
         TxtRRNNo.SetFocus
         Exit Sub
      End If
      If mCommodityID_ <> RsMst_ISIN!CommodityID Then
         MsgBox "ISIN Commodity And Above Selected Commodity Mismatch!!!"
         TxtRRNNo.SetFocus
         Exit Sub
      End If
  End If
  mRWCID = RsMst_RRN!RRNID
  TxtRRNQty = IIf(IsNull(RsMst_RRN!RRNQty), "", RsMst_RRN!RRNQty)
  TxtBalanceQty = Format(Val(TxtRRNQty) - Val(RsMst_RRN!DeliveredQty), "#######0.000")
  mClientIDRow = RsMst_RRN!ClientIDRow
  Call GetClientDetails(mClientIDRow)
End If
End Sub
Private Sub WHRDetails(mCMPID_, mCommodityID_, mTxtRRNNo_)
If mExchangeTypeID = 0 Or IsNull(mExchangeTypeID) Then Exit Sub
If mCMPID_ = 0 Or IsNull(mCMPID_) Then Exit Sub
If mCommodityID_ = 0 Or IsNull(mCommodityID_) Then Exit Sub
If mTxtRRNNo_ = "" Then Exit Sub

'Call TableTxn_WHRDetail("Where CMPID = " & mCMPID_ & " And CommodityId = " & mCommodityID_ & "")
Call TableTxn_WHRDetail("Where  sm_prefix = '" & TxtPreFix & "' And WHRNo = " & mTxtRRNNo_ & "")
If Check1.Value = 0 Then
   If RsTxn_WHRDetail.RecordCount = 0 Then
      MsgBox "No Record Found!!!"
      TxtRRNNo.SetFocus
      Exit Sub
   End If
   RsTxn_WHRDetail.MoveFirst
   RsTxn_WHRDetail.Find "WHRNo = " & mTxtRRNNo_
   If RsTxn_WHRDetail!Flag <> "P" Or IsNull(RsTxn_WHRDetail!Flag) Then
      MsgBox " WHR Print is not Generated !!!!! "
      Exit Sub
   End If
   
   If Not RsTxn_WHRDetail.EOF Then
      If mCMPID_ <> RsTxn_WHRDetail!CMPID Then
         MsgBox "WHR CMP And Above CMP Mismatch!!!"
         CmbCMPID.SetFocus
         Exit Sub
      End If
      If mCommodityID_ <> RsTxn_WHRDetail!CommodityID Then
         MsgBox "WHR Commodity And Above Selected Commodity Mismatch!!!"
         CmbCommodityID.SetFocus
         Exit Sub
      End If
      mRWCID = RsTxn_WHRDetail!WHRNo
      TxtRRNQty = IIf(IsNull(RsTxn_WHRDetail!WHRQuantity), "0", RsTxn_WHRDetail!WHRQuantity)
      TxtBalanceQty = IIf(IsNull(RsTxn_WHRDetail!WHRQuantity), "0", RsTxn_WHRDetail!WHRQuantity) - IIf(IsNull(RsTxn_WHRDetail!DeliveredQuantity), "0", RsTxn_WHRDetail!DeliveredQuantity)
      
      TxtRRNBags = IIf(IsNull(RsTxn_WHRDetail!WHRNoOfBags), "0", RsTxn_WHRDetail!WHRNoOfBags)
      TxtBalanceBags = Val(TxtRRNBags) - IIf(IsNull(RsTxn_WHRDetail!DeliveredNoOfBags), 0, RsTxn_WHRDetail!DeliveredNoOfBags)
      
      mClientIDRow = RsTxn_WHRDetail!ClientIDRow
      If RsTxn_WHRDetail!PledgeID <> 0 Then
         MsgBox "Warning : This Warehouse Receipt is already Pledged!!!"
      End If
      Call GetClientDetails(mClientIDRow)
   End If
End If
End Sub
Private Sub CDCDetails(mCMPID_, mCommodityID_, mTxtRRNNo_)
If mExchangeTypeID = 0 Or IsNull(mExchangeTypeID) Then Exit Sub
If mCMPID_ = 0 Or IsNull(mCMPID_) Then Exit Sub
If mCommodityID_ = 0 Or IsNull(mCommodityID_) Then Exit Sub
If mTxtRRNNo_ = "" Then Exit Sub

Call TableTxn_CDC_GSLDetail("Where  sm_prefix = '" & TxtPreFix & "' And CDCNo = " & mTxtRRNNo_)
If RsTxn_CDC_GSLDetail.RecordCount = 0 Then
   MsgBox "No Such CDC Record Found!!!"
   Exit Sub
End If

If Check1.Value = 0 Then
   
   
   If RsTxn_CDC_GSLDetail!Flag <> "P" Then
      
      If RsTxn_CDC_GSLDetail!Flag = "F" Then
         MsgBox "Payout is not executed !!!! "
         TxtRRNNo.SetFocus
         Exit Sub
      End If
      If RsTxn_CDC_GSLDetail!Flag = "T" Then
         MsgBox "This CDC is reissued.It can't be withdrawn !!!! "
         TxtRRNNo.SetFocus
         Exit Sub
      End If
      If RsTxn_CDC_GSLDetail!Flag = "W" Then
         MsgBox "This CDC is already withdrawn !!!! "
         TxtRRNNo.SetFocus
         Exit Sub
      End If
   End If
Else
   tmp = "SELECT Txn_CDC_Assayer_Detail.AcceptenceFlag " & _
   "FROM Txn_CDC_Assayer_Detail INNER JOIN Txn_CDC_GSL_Detail ON (Txn_CDC_Assayer_Detail.SM_Prefix = " & _
   "Txn_CDC_GSL_Detail.SM_Prefix) AND (Txn_CDC_Assayer_Detail.CDCID = Txn_CDC_GSL_Detail.CDCID) " & _
   "WHERE (((Txn_CDC_GSL_Detail.CDCNo)=" & mTxtRRNNo_ & ") AND ((Txn_CDC_GSL_Detail.SM_Prefix)= '" & TxtPreFix & "'));"
   Call Tmp1Table
   If TmpRs1.RecordCount > 0 Then
      If TmpRs1.Fields(0) <> 0 Then
         MsgBox "Selected CDC is not Rejected, you can't withdraw this CDC !!!!! "
         TxtRRNNo.SetFocus
         Exit Sub
      End If
   End If
End If

mCDCID = RsTxn_CDC_GSLDetail!CDCID
Call TableTxn_CDC("Where sm_prefix = '" & TxtPreFix & "' And  CDCID =  " & mCDCID & " ")
If Not RsTxn_CDC.EOF Then
   If mCMPID_ <> RsTxn_CDC!CMPID Then
      MsgBox "CDC CMP and Above CMP Selected Mismatch!!!"
      CmbCMPID.SetFocus
      Exit Sub
   End If
   If RsTxn_CDC!PrintFlag <> "P" Then
      MsgBox "CDC Print is not Generated !!!"
      TxtRRNNo.SetFocus
      Exit Sub
   End If
   
   If mCommodityID_ <> RsTxn_CDC!CommodityID Then
      MsgBox "CDC Commodity and Above Commodity Selected Mismatch!!!"
      CmbCommodityID.SetFocus
      Exit Sub
   End If
End If

mRWCID = RsTxn_CDC_GSLDetail!CDCNo
mGSLID = RsTxn_CDC_GSLDetail!GSLID
mClientIDRow = RsTxn_CDC!ClientIDRow
Call TableMst_GSL(" Where sm_prefix = '" & TxtPreFix & "' And  GSLID = " & RsTxn_CDC_GSLDetail!GSLID)
TxtRRNQty = RsMst_GSL!DepositeWeight
TxtBalanceQty = RsMst_GSL!DepositeWeight - RsMst_GSL!CDTFWeight

Call GetClientDetails(mClientIDRow)
'Call DisplayGSLData(mGSLID)

End Sub
Private Sub GetClientDetails(mClientIDRow_ As Double)
If mClientIDRow_ = 0 Or IsNull(mClientIDRow_) Then Exit Sub
ans = GetClientNameByRow(mClientIDRow_)
If ans(0) = "" Then
   MsgBox "Not Found!!!"
   Exit Sub
End If
TxtClientID = ans(0)
TxtClientName = ans(1)
TxtDPID = ans(2)
TxtDPName = ans(3)

End Sub
Private Sub DisplayGSLData(mGSLID_)
Dim r As Variant
r = MSHFlexGrid4.Rows - 1
MSHFlexGrid4.Rows = MSHFlexGrid4.Rows + 1
Call TableMst_GSL("Where sm_prefix = '" & TxtPreFix & "' And GSLID = " & mGSLID_ & " ")
If RsMst_GSL.RecordCount = 0 Then
   'MsgBox "No GSL Attach To This CDC!!!"
   Exit Sub
End If


MSHFlexGrid4.TextMatrix(r, 1) = GetGodownName(RsMst_GSL!GodownID, "N")
MSHFlexGrid4.TextMatrix(r, 2) = RsMst_GSL!StackNo
MSHFlexGrid4.TextMatrix(r, 3) = RsMst_GSL!LotNo
MSHFlexGrid4.TextMatrix(r, 4) = IIf(IsNull(RsMst_GSL!DepositeBags), 0, RsMst_GSL!DepositeBags) - IIf(IsNull(RsMst_GSL!CDTFBags), 0, RsMst_GSL!CDTFBags)
MSHFlexGrid4.TextMatrix(r, 5) = IIf(IsNull(RsMst_GSL!DepositeWeight), 0, RsMst_GSL!DepositeWeight) - IIf(IsNull(RsMst_GSL!CDTFWeight), 0, RsMst_GSL!CDTFWeight)
MSHFlexGrid4.TextMatrix(r, 8) = Val(TxtAdjWt)
MSHFlexGrid4.TextMatrix(r, 9) = Format((Val(MSHFlexGrid4.TextMatrix(r, 4)) * Val(TxtCGunnyWt)) / 1000, "#######0.000")
MSHFlexGrid4.TextMatrix(r, 10) = IIf(IsNull(RsMst_GSL!DepositeWeight), 0, RsMst_GSL!DepositeWeight) - IIf(IsNull(RsMst_GSL!CDTFWeight), 0, RsMst_GSL!CDTFWeight) 'Val(TxtGSLNetWeight)
MSHFlexGrid4.TextMatrix(r, 6) = IIf(IsNull(RsMst_GSL!DepositeBags), 0, RsMst_GSL!DepositeBags) - IIf(IsNull(RsMst_GSL!CDTFBags), 0, RsMst_GSL!CDTFBags)
MSHFlexGrid4.TextMatrix(r, 7) = Val(MSHFlexGrid4.TextMatrix(r, 10)) - Val(MSHFlexGrid4.TextMatrix(r, 8)) + Val(MSHFlexGrid4.TextMatrix(r, 9))   'IIf(IsNull(RsMst_GSL!DepositeWeight), 0, RsMst_GSL!DepositeWeight) - IIf(IsNull(RsMst_GSL!CDTFWeight), 0, RsMst_GSL!CDTFWeight)
MSHFlexGrid4.TextMatrix(r, 11) = Val(TxtGSLWeightGainLoss)
MSHFlexGrid4.TextMatrix(r, 12) = TxtRRNNo
MSHFlexGrid4.TextMatrix(r, 13) = TxtGCTDremarks
MSHFlexGrid4.TextMatrix(r, 14) = mGSLID_
'MSHFlexGrid4.TextMatrix(r, 15) = mGodownID
'MSHFlexGrid4.TextMatrix(r, 16) = mRWCID
 
MSHFlexGrid4.TextMatrix(0, 4) = Val(MSHFlexGrid4.TextMatrix(0, 4)) + Val(MSHFlexGrid4.TextMatrix(r, 4))
MSHFlexGrid4.TextMatrix(0, 5) = Val(MSHFlexGrid4.TextMatrix(0, 5)) + Val(MSHFlexGrid4.TextMatrix(r, 5))
MSHFlexGrid4.TextMatrix(0, 6) = Val(MSHFlexGrid4.TextMatrix(0, 6)) + Val(MSHFlexGrid4.TextMatrix(r, 6))
MSHFlexGrid4.TextMatrix(0, 7) = Val(MSHFlexGrid4.TextMatrix(0, 7)) + Val(MSHFlexGrid4.TextMatrix(r, 7))
MSHFlexGrid4.TextMatrix(0, 8) = Val(MSHFlexGrid4.TextMatrix(0, 8)) + Val(MSHFlexGrid4.TextMatrix(r, 8))
MSHFlexGrid4.TextMatrix(0, 9) = Val(MSHFlexGrid4.TextMatrix(0, 9)) + Val(MSHFlexGrid4.TextMatrix(r, 9))
MSHFlexGrid4.TextMatrix(0, 10) = Val(MSHFlexGrid4.TextMatrix(0, 10)) + Val(MSHFlexGrid4.TextMatrix(r, 10))
MSHFlexGrid4.TextMatrix(0, 11) = Val(MSHFlexGrid4.TextMatrix(0, 11)) + Val(MSHFlexGrid4.TextMatrix(r, 11))

'MSHFlexGrid4.ColWidth(12) = 0
Call Clear_GCTDTabControl
CmdGCTDAdd2List.Enabled = False
End Sub
Private Sub TxtTTDTareWt_LostFocus()
If TxtTTDTareWt = "" Then Exit Sub
If Not IsNumeric(TxtTTDTareWt) Then
   MsgBox "Invalid Number Format"
   TxtTTDTareWt.SetFocus
End If
Call CalTruckWt
End Sub

Private Sub TxtTTDGrossWt_LostFocus()
If TxtTTDGrossWt = "" Then Exit Sub
If Not IsNumeric(TxtTTDGrossWt) Then
   MsgBox "Invalid Number Format"
   TxtTTDGrossWt.SetFocus
End If

End Sub

Private Sub TxtTTDTotalWt_LostFocus()
If TxtTTDTotalWt = "" Then Exit Sub
If Not IsNumeric(TxtTTDTotalWt) Then
   MsgBox "Invalid Number Format"
   TxtTTDTotalWt.SetFocus
End If
Call CalTruckWt
End Sub

Private Sub CalTruckWt()
TxtTTDGrossWt = Val(TxtTTDTotalWt) - Val(TxtTTDTareWt)
End Sub
Private Sub Clear_RWCTabControl()
TxtRRNNo = ""
TxtISINno = ""
TxtRRNQty = ""
TxtBalanceQty = ""
TxtAdjustedQty = ""
TxtRRNRemarks = ""
TxtClientID = ""
TxtClientName = ""
TxtDPID = ""
TxtDPName = ""
Check2.Value = 0
CmdAdd2ListRRN.Caption = "Add to List"
'Call Grid_Head_RRN
End Sub
Private Sub Clear_GCTDTabControl()
CmbGodownNo.Text = ""
CmbStack.Text = ""
CmbLot.Text = ""
'CmbGCTDClient.Text = ""
TxtGCTDNoofBags = ""
TxtGCTDGrossWt = ""
TxtGGunnyWeight = ""
TxtGSLISINno = ""
TxtGSLNetWeight = ""
TxtGCTDremarks = ""
TxtAdjWt = ""
TxtGSLAvailableNofBags = ""
TxtGSLAvailableGrossWgt = ""
CmbGCTDClient.Text = ""
TxtGSLWeightGainLoss = ""
'CmbNonNcdexClient.Text = ""
mGSLID = 0
CmdGCTDAdd2List.Caption = "Add to List"
Check2.Enabled = False
Call EnDnGCTDTab(True)
If mExchangeTypeID = 3 Then
   CmbGodownNo.Enabled = False
   CmbStack.Enabled = False
   CmbLot.Enabled = False
End If
mPledgeID = ""
End Sub
Private Sub EnDnGCTDTab(Yn As Boolean)
Dim Ny As Boolean
Ny = IIf(Yn, False, True)

CmdGCTDAdd2List.Enabled = Yn
CmbGodownNo.Enabled = Yn
CmbStack.Enabled = Yn
CmbLot.Enabled = Yn
Check2.Enabled = Yn
TxtGCTDNoofBags.Locked = Ny
TxtGCTDGrossWt.Locked = Ny
TxtAdjWt.Locked = Ny
TxtGCTDremarks.Locked = Ny

End Sub
Public Sub InData()

PicRRN.Enabled = False
Pic_GCTD.Enabled = False
Pic_TTD.Enabled = False

Call Grid_Head_TTD
Call Grid_Head_RRN
Call Grid_Head_GCTD

TxtCxTFNo.Locked = True
TxtCxTFDate.Enabled = False
TxtCMPID.Locked = True
CmbCMPID.Visible = False
TxtCMPID.Visible = True
TxtCommodityID.Locked = True
CmbCommodityID.Visible = False
TxtCommodityID.Visible = True
CmbExchangeTypeID.Enabled = False
TxtAgent.Locked = True
Check1.Enabled = False

TxtAgent.Locked = True
CmbAgent.Visible = False
CmdAgent.Visible = False
TxtAgent.Visible = True

TxtCxTFNo = IIf(IsNull(RsTxn_CxTF_Details!CxTFNo), "", RsTxn_CxTF_Details!CxTFNo)
TxtManualCDTFNo = IIf(IsNull(RsTxn_CxTF_Details!ManualCDTFNo), "", RsTxn_CxTF_Details!ManualCDTFNo)

TxtPreFix = IIf(IsNull(RsTxn_CxTF_Details!SM_Prefix), "", RsTxn_CxTF_Details!SM_Prefix)
TxtCxTFDate.Value = IIf(IsNull(RsTxn_CxTF_Details!CxTFDate), "", RsTxn_CxTF_Details!CxTFDate)
TxtCMPID = GetCMPName(IIf(IsNull(RsTxn_CxTF_Details!CMPID), "", RsTxn_CxTF_Details!CMPID))
mCMPID = IIf(IsNull(RsTxn_CxTF_Details!CMPID), "", RsTxn_CxTF_Details!CMPID)
TxtCommodityID = GetCommodityName(IIf(IsNull(RsTxn_CxTF_Details!CommodityID), "", RsTxn_CxTF_Details!CommodityID))
mCommodityId = IIf(IsNull(RsTxn_CxTF_Details!CommodityID), "", RsTxn_CxTF_Details!CommodityID)
CmbExchangeTypeID.Text = GetExchangeName(IIf(IsNull(RsTxn_CxTF_Details!ExchangeTypeID), "", RsTxn_CxTF_Details!ExchangeTypeID))
mExchangeTypeID = IIf(IsNull(RsTxn_CxTF_Details!ExchangeTypeID), "", RsTxn_CxTF_Details!ExchangeTypeID)
Check1.Value = IIf(IsNull(RsTxn_CxTF_Details!BookedSpaceFlag), "", RsTxn_CxTF_Details!BookedSpaceFlag)
TxtFlag = IIf(IsNull(RsTxn_CxTF_Details!Flag), "", RsTxn_CxTF_Details!Flag)
TxtAgent = IIf(IsNull(RsTxn_CxTF_Details!Agent), "", RsTxn_CxTF_Details!Agent)

If Not IsNull(RsTxn_CxTF_Details!AgentID) Then
   mAgentID = IIf(IsNull(RsTxn_CxTF_Details!AgentID), "", RsTxn_CxTF_Details!AgentID)
   TxtAgent = GetAgentName(IIf(IsNull(RsTxn_CxTF_Details!AgentID), "", RsTxn_CxTF_Details!AgentID))
End If

TxtRemark = IIf(IsNull(RsTxn_CxTF_Details!Remark), "", RsTxn_CxTF_Details!Remark)
TxtTxnType = IIf(IsNull(RsTxn_CxTF_Details!TxnType), "", RsTxn_CxTF_Details!TxnType)
'TxtContractID = GetContract(IIf(IsNull(RsTxn_CxTF_Details!ContractID), "", RsTxn_CxTF_Details!ContractID))
mContractID = IIf(IsNull(RsTxn_CxTF_Details!ContractID), "", RsTxn_CxTF_Details!ContractID)
CmbCMPID.Text = TxtCMPID

'Call CmbExchangeTypeID_GotFocus
'Call CmbExchangeTypeID_LostFocus
'Call CmbCMPID_GotFocus
'Call CmbCMPID_LostFocus
CmbCommodityID.Text = TxtCommodityID
Call CmbCommodityID_GotFocus
Call CmbCommodityID_LostFocus
If mExchangeTypeID = 1 Then 'NON NCDEX
   Call SetCaptions("WHR No / CDTF No", "WHR-CDTF No", "ClientIDRow", "WHR Qty", "Depository ID", "Depository Name", False)
End If
If mExchangeTypeID = 2 Then 'NCDEX
   If Check1.Value = 1 Then
      Call SetCaptions("client details", "CDTF No", "Client Name", "Rejected Qty", "DP ID", "DP Name", True)
   Else
      Call SetCaptions("Remat Request No. (RRN)", "RRN No", "ISIN No", "RRN Qty", "DP ID", "DP Name", True)
   End If
End If
If mExchangeTypeID = 3 Then 'SPOT
   Call SetCaptions("CDC No", "CDC-No", "CDC-No", "CDC Qty", "TM ID", "Trading Member Name", True)
End If
Call InDataClient
Call InDataTruck
Call InDataGSL
If RsTxn_CxTF_Details.RecordCount = 1 Then
   Call GButtonLock_1(Me, True)
Else
   Call GButtonLock(Me, True)
End If
Frame3.Visible = False
EditCmd.Enabled = False
'DeleteCmd.Enabled = False
CmdPrint.Enabled = True
End Sub
Private Sub InDataGSL()
Call TableTxn_CxTF_GSL("  Where  sm_prefix = '" & TxtPreFix & "' And CxTFNo = " & TxtCxTFNo & " And TxnType = 'W'")
MSHFlexGrid4.Rows = MSHFlexGrid4.Rows + RsTxn_CxTF_GSL.RecordCount
If RsTxn_CxTF_GSL.RecordCount = 0 Then Exit Sub

Dim ans As Variant

For i = 2 To MSHFlexGrid4.Rows - 2
    
    
    MSHFlexGrid4.Col = 7
    MSHFlexGrid4.Row = i
    MSHFlexGrid4.CellBackColor = &HFFC0C0


    ans = GetGSLByGLSID(RsTxn_CxTF_GSL!GSLID, TxtPreFix)
    MSHFlexGrid4.TextMatrix(i, 0) = RsTxn_CxTF_GSL!GSLDetailID
    
    MSHFlexGrid4.TextMatrix(i, 1) = ans(3) ' TmpRs!GodownNo
    MSHFlexGrid4.TextMatrix(i, 2) = ans(4) 'TmpRs!StackNo
    MSHFlexGrid4.TextMatrix(i, 3) = ans(5) 'TmpRs!LotNo
    
    If MSHFlexGrid4.TextMatrix(1, 12) = "Client Name" Then
       MSHFlexGrid4.TextMatrix(i, 12) = GetClientName(IIf(IsNull(ans(9)), 0, ans(9)))
    Else
       MSHFlexGrid4.TextMatrix(i, 12) = IIf(IsNull(ans(10)) Or ans(10) = 0, "", ans(10))
    End If
    MSHFlexGrid4.TextMatrix(i, 15) = IIf(IsNull(ans(6)), "", ans(6))
    MSHFlexGrid4.TextMatrix(i, 6) = RsTxn_CxTF_GSL!NoofBags
    MSHFlexGrid4.TextMatrix(i, 7) = RsTxn_CxTF_GSL!GrossWeight
    MSHFlexGrid4.TextMatrix(i, 8) = RsTxn_CxTF_GSL!AdjNCMSLWeight
    MSHFlexGrid4.TextMatrix(i, 9) = RsTxn_CxTF_GSL!GunnyWeight
    MSHFlexGrid4.TextMatrix(i, 10) = RsTxn_CxTF_GSL!DematWeight
    'MSHFlexGrid4.TextMatrix(i, 11) = Weight Gain/LossIif(Isnull(RsTxn_CxTF_GSL!DematWeight),0,RsTxn_CxTF_GSL!De
    MSHFlexGrid4.TextMatrix(i, 13) = RsTxn_CxTF_GSL!Remarks
    
    If Not (IsNull(RsTxn_CxTF_GSL!ClientIDRow)) Then
       ans = GetClientNameByRow(RsTxn_CxTF_GSL!ClientIDRow)
       MSHFlexGrid4.TextMatrix(i, 14) = RsTxn_CxTF_GSL!GSLID
       MSHFlexGrid4.TextMatrix(i, 16) = RsTxn_CxTF_GSL!ClientIDRow
       'MSHFlexGrid4.TextMatrix(i, 17) = RRNID
    End If
    MSHFlexGrid4.Col = 7
    MSHFlexGrid4.Row = i
    MSHFlexGrid4.CellBackColor = &HFFC0C0



    MSHFlexGrid4.TextMatrix(0, 4) = Val(MSHFlexGrid4.TextMatrix(0, 4)) + Val(MSHFlexGrid4.TextMatrix(i, 4))
    MSHFlexGrid4.TextMatrix(0, 5) = Val(MSHFlexGrid4.TextMatrix(0, 5)) + Val(MSHFlexGrid4.TextMatrix(i, 5))
    MSHFlexGrid4.TextMatrix(0, 6) = Val(MSHFlexGrid4.TextMatrix(0, 6)) + Val(MSHFlexGrid4.TextMatrix(i, 6))
    MSHFlexGrid4.TextMatrix(0, 7) = Val(MSHFlexGrid4.TextMatrix(0, 7)) + Val(MSHFlexGrid4.TextMatrix(i, 7))
    MSHFlexGrid4.TextMatrix(0, 8) = Val(MSHFlexGrid4.TextMatrix(0, 8)) + Val(MSHFlexGrid4.TextMatrix(i, 8))
    MSHFlexGrid4.TextMatrix(0, 9) = Val(MSHFlexGrid4.TextMatrix(0, 9)) + Val(MSHFlexGrid4.TextMatrix(i, 9))
    MSHFlexGrid4.TextMatrix(0, 10) = Val(MSHFlexGrid4.TextMatrix(0, 10)) + Val(MSHFlexGrid4.TextMatrix(i, 10))
    MSHFlexGrid4.TextMatrix(0, 11) = Val(MSHFlexGrid4.TextMatrix(0, 11)) + Val(MSHFlexGrid4.TextMatrix(i, 11))
    
    If mExchangeTypeID = 3 Then
       MSHFlexGrid4.TextMatrix(i, 12) = "" 'IIf(IsNull(Ans(10)) Or Ans(10) = 0, "", Ans(10))
       MSHFlexGrid4.ColWidth(12) = 0
    End If
    
    RsTxn_CxTF_GSL.MoveNext
Next
End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsTxn_CxTF_Details.MoveNext
If RsTxn_CxTF_Details.EOF Then
   RsTxn_CxTF_Details.MoveLast
   LF_Flag = "Y"
End If
Call InData
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
RsTxn_CxTF_Details.MovePrevious
If RsTxn_CxTF_Details.BOF Then
   RsTxn_CxTF_Details.MoveFirst
   LF_Flag = "Y"
End If
Call InData
If LF_Flag = "Y" Then
    PreviousCmd.Enabled = False
    FirstCmd.Enabled = False
End If
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub FirstCmd_Click()
RsTxn_CxTF_Details.MoveFirst
Call InData
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsTxn_CxTF_Details.MoveLast
Call InData
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False
End Sub
Private Sub CmbGCTDClient_GotFocus()
tmp = CmbGCTDClient.Text
CmbGCTDClient.Clear
If MSHFlexGrid2.Rows < 4 Then Exit Sub
For i = 2 To MSHFlexGrid2.Rows - 2
    If mExchangeTypeID = 2 Then
       If Check1.Value = 0 Then
          CmbGCTDClient.AddItem MSHFlexGrid2.TextMatrix(i, 2)
       Else
          CmbGCTDClient.AddItem MSHFlexGrid2.TextMatrix(i, 4)
       End If
    ElseIf mExchangeTypeID = 1 Then
       CmbGCTDClient.AddItem MSHFlexGrid2.TextMatrix(i, 4)
    ElseIf mExchangeTypeID = 3 Then
       CmbGCTDClient.AddItem MSHFlexGrid2.TextMatrix(i, 1)
       
    End If
Next i
CmbGCTDClient.Text = tmp
End Sub
Private Sub CmbGCTDClient_LostFocus()
If CmbGCTDClient.Text = "" Then Exit Sub
For i = 2 To MSHFlexGrid2.Rows - 2
    If mExchangeTypeID = 2 Then
       If Check1.Value = 0 Then
        If CmbGCTDClient.Text = MSHFlexGrid2.TextMatrix(i, 2) Then
           Exit Sub
        End If
       Else
          If CmbGCTDClient.Text = MSHFlexGrid2.TextMatrix(i, 4) Then
             mGSLClientIDRow = Val(MSHFlexGrid2.TextMatrix(i, 11))
             Exit Sub
          End If
       
       End If
    ElseIf mExchangeTypeID = 1 Then
        If CmbGCTDClient.Text = MSHFlexGrid2.TextMatrix(i, 4) Then
           mGSLClientIDRow = Val(MSHFlexGrid2.TextMatrix(i, 11))
           Exit Sub
        End If
    ElseIf mExchangeTypeID = 3 Then
        If CmbGCTDClient.Text = MSHFlexGrid2.TextMatrix(i, 1) Then
           mGSLClientIDRow = Val(MSHFlexGrid2.TextMatrix(i, 11))
           Exit Sub
        End If
        
    End If
Next i

MsgBox "Invalid Selection "
CmbGCTDClient.SetFocus

End Sub
Function CheckDuplicateClient() As Boolean
Dim RetVal As Boolean
RetVal = False
If MSHFlexGrid2.Rows > 3 Then
   For u = 1 To MSHFlexGrid2.Rows - 1
       If MSHFlexGrid2.TextMatrix(u, 3) = TxtClientID And MSHFlexGrid2.TextMatrix(u, 5) = TxtDPID Then
          If LCase(CmdAdd2ListRRN.Caption) = "add to list" Then
             RetVal = True
             Exit For
          Else
             If u <> MSHFlexGrid2.RowSel Then
                RetVal = True
             End If
          End If
       End If
   Next
End If

CheckDuplicateClient = RetVal
End Function

Private Sub TxtTTDTruckNo_KeyPress(KeyAscii As Integer)
tmp = KeyAscii
If KeyAscii = 14 Then
   SSTab1.Tab = 2
   TxtTTDTruckNo.SetFocus
End If
End Sub
Private Sub CmdPrint_Click()

If RsTxn_CxTF_Details!HologramNo = "" Or IsNull(RsTxn_CxTF_Details!HologramNo) Then
   Frame3.Visible = True
   TxtHologramNo.SetFocus
   Call TableMst_Hologram
   Exit Sub
Else
   TxtHologramNo = ""
   Frame3.Visible = False
End If
RsTxn_CxTF_Details!Flag = "P"
RsTxn_CxTF_Details.Update

Gen_mTimeStamp = GetTimeStamp
Call Create_Xls(Gen_mTimeStamp & Gen_UserName & "_CITF-" & Trim(TxtPreFix) & "" & Trim(TxtCxTFNo) & ".xls")


Call Xls_Print_Rpt
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
oBook.SaveAs (Pat & "\excel\" & FileName)

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

Public Sub Xls_Print_Rpt()

'Dim oXL As Excel.Application
'Dim oWB As Excel.Workbook
'Dim oWS As Excel.Worksheet
'Dim mRow As Integer
'Dim mCol As Integer
Dim ans As Variant
MsgBox ("Wait till Next Message")
Set oXL = New Excel.Application
'Dim Pat As String
Pat = App.Path
'Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & mTimeStamp & Gen_UserName & "_CITF-" & Trim(TxtPreFix) & "" & Trim(TxtCxTFNo) & ".xls")
 
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "_CITF-" & Trim(TxtPreFix) & "" & Trim(TxtCxTFNo) & ".xls")
 
i = 2
Set oWS = oWB.Worksheets("Sheet" & 1)
oWS.PageSetup.RightMargin = 0.05
oWS.Name = TxtPreFix & "-" & TxtCxTFNo  'MSHFlexGrid2.TextMatrix(i, 23)
oWS.ClearArrows
mRow = 4

oWS.Range("a1:f1").Merge
oWS.Range("a:a").ColumnWidth = 10.86
oWS.Range("b:b").ColumnWidth = 12.86
oWS.Range("c:c").ColumnWidth = 12.86
oWS.Range("d:d").ColumnWidth = 12.86
oWS.Range("e:e").ColumnWidth = 17.86 '32.14
oWS.Range("f:f").ColumnWidth = 17.86


tmp = "a" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = GetCMPAddress(RsTxn_CxTF_Details!CMPID, "D") '& "-" & GetCMPName(RsTxn_CxTF_Details!CMPID) '& "-" & GetCMPAddress(RsTxn_CxTF_Details!CMPID)
oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).Font.Size = 10
'oWS.Cells(mRow, c + 1).RowHeight = 24
mRow = mRow + 1
'oWS.Range("A1:E1").Borders.Color = RGB(0, 0, 0)

tmp = "a" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = "Commodity Issue Transaction Form (CITF)"
oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).Font.Size = 10
oWS.Cells(mRow, c + 1).RowHeight = 24


mRow = mRow + 1

tmp = "a" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = "(This is NOT an instrument of title to the goods)"
oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).Font.Size = 7


mRow = mRow + 1

oWS.Cells(mRow, c + 1) = "CITF No"
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).WrapText = True
'oWS.Range("b5:c5").Merge
tmp = "b" & mRow & ":" & "c" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1).Font.Size = 7
oWS.Cells(mRow, c + 2) = TxtPreFix & "-" & TxtCxTFNo
oWS.Cells(mRow, c + 2).HorizontalAlignment = 2
oWS.Cells(mRow, c + 2).Font.Size = 7


oWS.Cells(mRow, c + 4) = "Date"
oWS.Cells(mRow, c + 4).Font.Size = 7
oWS.Cells(mRow, c + 4).Font.Bold = True

'oWS.Range(tmp).Merge
tmp = "e" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 5) = Format(TxtCxTFDate, "dd-mmm-yyyy")
oWS.Cells(mRow, c + 5).HorizontalAlignment = 2
oWS.Cells(mRow, c + 5).Font.Size = 7

mRow = mRow + 1

oWS.Cells(mRow, c + 1) = "Location"
oWS.Cells(mRow, c + 1).Font.Size = 7
oWS.Cells(mRow, c + 1).Font.Bold = True
tmp = "b" & mRow & ":" & "c" & mRow

oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 2) = GetLocationFromCMP(RsTxn_CxTF_Details!CMPID)
oWS.Cells(mRow, c + 2).HorizontalAlignment = 2
oWS.Cells(mRow, c + 2).Font.Size = 7

oWS.Cells(mRow, c + 4) = "State"
oWS.Cells(mRow, c + 4).Font.Size = 7
oWS.Cells(mRow, c + 4).Font.Bold = True

tmp = "e" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 5) = GetStateFromLocation(GetLocationIDFromCMP(RsTxn_CxTF_Details!CMPID))
oWS.Cells(mRow, c + 5).HorizontalAlignment = 2
oWS.Cells(mRow, c + 5).Font.Size = 7

mRow = mRow + 1

oWS.Cells(mRow, c + 1) = "CMP ID & Name"
oWS.Cells(mRow, c + 1).Font.Size = 7
oWS.Cells(mRow, c + 1).WrapText = True
oWS.Cells(mRow, c + 1).Font.Bold = True

tmp = "b" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 2) = RsTxn_CxTF_Details!CMPID & "-" & GetCMPName(RsTxn_CxTF_Details!CMPID)
oWS.Cells(mRow, c + 2).HorizontalAlignment = 2
oWS.Cells(mRow, c + 2).Font.Size = 7
oWS.Cells(mRow, c + 2).WrapText = True

mRow = mRow + 1

oWS.Cells(mRow, c + 1) = "CMP Address"
oWS.Cells(mRow, c + 1).Font.Size = 7
oWS.Cells(mRow, c + 1).WrapText = True
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).RowHeight = 50

tmp = "b" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 2) = GetCMPAddress(RsTxn_CxTF_Details!CMPID, "A")
oWS.Cells(mRow, c + 2).HorizontalAlignment = 2
oWS.Cells(mRow, c + 2).Font.Size = 7
oWS.Cells(mRow, c + 2).WrapText = True

mRow = mRow + 1

oWS.Cells(mRow, c + 1) = "Transaction Type"
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).Font.Size = 7
oWS.Cells(mRow, c + 1).WrapText = True

tmp = "b" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 2) = CmbExchangeTypeID.Text
oWS.Cells(mRow, c + 2).Font.Size = 7

mRow = mRow + 1

oWS.Cells(mRow, c + 1) = "Name Of Agent"
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).Font.Size = 7

tmp = "b" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 2) = TxtAgent
oWS.Cells(mRow, c + 2).Font.Size = 7

mRow = mRow + 1

oWS.Cells(mRow, c + 1) = "Commodity"
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).Font.Size = 7

tmp = "b" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 2) = TxtCommodityID
oWS.Cells(mRow, c + 2).Font.Size = 7

mRow = mRow + 1

tmp = "a" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = "Withdrawer Details"
oWS.Cells(mRow, c + 1).Font.Size = 9
oWS.Cells(mRow, c + 1).Font.Bold = True

mRow = mRow + 1
If CmbExchangeTypeID.Text = "NCDEX" Then
   oWS.Cells(mRow, c + 1) = "RRN No"
ElseIf CmbExchangeTypeID.Text = "Non NCDEX" Then
   oWS.Cells(mRow, c + 1) = "WHR No"
ElseIf CmbExchangeTypeID.Text = "SPOT" Then
   oWS.Cells(mRow, c + 1) = "CDC No"
End If
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).Font.Size = 7
oWS.Cells(mRow, c + 2) = "Client-ID"
oWS.Cells(mRow, c + 2).Font.Bold = True
oWS.Cells(mRow, c + 2).Font.Size = 7

If mExchangeTypeID <> 2 Then
   tmp = "c" & mRow & ":" & "d" & mRow
   oWS.Range(tmp).Merge
Else
   tmp = "c" & mRow & ":" & "e" & mRow
   oWS.Range(tmp).Merge
End If

oWS.Cells(mRow, c + 3) = "Client Name"
oWS.Cells(mRow, c + 3).Font.Bold = True
oWS.Cells(mRow, c + 3).Font.Size = 7
If mExchangeTypeID <> 2 Then
   oWS.Cells(mRow, c + 5) = "No. of Bags"
   oWS.Cells(mRow, c + 5).Font.Bold = True
   oWS.Cells(mRow, c + 5).Font.Size = 7
End If

oWS.Cells(mRow, c + 6) = "Electronic Weight in MT"
oWS.Cells(mRow, c + 6).Font.Bold = True
oWS.Cells(mRow, c + 6).Font.Size = 7

mRow = mRow + 1

For k = 2 To MSHFlexGrid2.Rows - 2
'k = 1
    oWS.Cells(mRow, c + 1) = "'" & MSHFlexGrid2.TextMatrix(k, 1)
    oWS.Cells(mRow, c + 1).Font.Size = 7
    oWS.Cells(mRow, c + 2) = MSHFlexGrid2.TextMatrix(k, 3)
    oWS.Cells(mRow, c + 2).Font.Size = 7
    If mExchangeTypeID <> 2 Then
       tmp = "c" & mRow & ":" & "d" & mRow
       oWS.Range(tmp).Merge
    Else
       tmp = "c" & mRow & ":" & "e" & mRow
       oWS.Range(tmp).Merge
    End If
    oWS.Cells(mRow, c + 3) = MSHFlexGrid2.TextMatrix(k, 4)
    oWS.Cells(mRow, c + 3).Font.Size = 7
    If mExchangeTypeID <> 2 Then
       oWS.Cells(mRow, c + 5) = Format(MSHFlexGrid2.TextMatrix(k, 15), "#######0.000")
       oWS.Cells(mRow, c + 5).Font.Size = 7
    End If
    oWS.Cells(mRow, c + 6) = Format(MSHFlexGrid2.TextMatrix(k, 9), "#######0.000")
    oWS.Cells(mRow, c + 6).Font.Size = 7
    mRow = mRow + 1
Next k

If mExchangeTypeID <> 2 Then
   tmp = "c" & mRow & ":" & "d" & mRow
   oWS.Range(tmp).Merge
Else
   tmp = "c" & mRow & ":" & "e" & mRow
   oWS.Range(tmp).Merge
End If
mRow = mRow + 1

If mExchangeTypeID <> 2 Then
   tmp = "a" & mRow & ":" & "d" & mRow
   oWS.Range(tmp).Merge
Else
   tmp = "a" & mRow & ":" & "e" & mRow
   oWS.Range(tmp).Merge
End If
oWS.Cells(mRow, c + 1) = "Total"
oWS.Cells(mRow, c + 1).Font.Size = 7
oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, c + 1).Font.Bold = True

If mExchangeTypeID <> 2 Then
   oWS.Cells(mRow, c + 5) = Format(MSHFlexGrid2.TextMatrix(0, 15), "#######0.000")
   oWS.Cells(mRow, c + 5).Font.Size = 7
End If
oWS.Cells(mRow, c + 6) = Format(MSHFlexGrid2.TextMatrix(0, 9), "#######0.000")
oWS.Cells(mRow, c + 6).Font.Size = 7

mRow = mRow + 1

tmp = "a" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = "Truck-wise Withdrawal Details"
oWS.Cells(mRow, c + 1).Font.Size = 9
oWS.Cells(mRow, c + 1).Font.Bold = True

mRow = mRow + 1

oWS.Cells(mRow, c + 1) = "Truck No"
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).Font.Size = 7
oWS.Cells(mRow, c + 2) = "Total Weight(MT)"
oWS.Cells(mRow, c + 2).Font.Bold = True
oWS.Cells(mRow, c + 2).Font.Size = 7
oWS.Cells(mRow, c + 3) = "Tare Weight(MT)"
oWS.Cells(mRow, c + 3).Font.Bold = True
oWS.Cells(mRow, c + 3).Font.Size = 7
oWS.Cells(mRow, c + 4) = "Gross Weight"
oWS.Cells(mRow, c + 4).Font.Bold = True
oWS.Cells(mRow, c + 4).Font.Size = 7
oWS.Cells(mRow, c + 5) = "Weighbridge Slip No"
oWS.Cells(mRow, c + 5).Font.Bold = True
oWS.Cells(mRow, c + 5).Font.Size = 7
oWS.Cells(mRow, c + 6) = "Remarks"
oWS.Cells(mRow, c + 6).Font.Bold = True
oWS.Cells(mRow, c + 6).Font.Size = 7

mRow = mRow + 1

For k = 2 To MSHFlexGrid3.Rows - 2
'k = 1
    oWS.Cells(mRow, c + 1) = MSHFlexGrid3.TextMatrix(k, 1)
    oWS.Cells(mRow, c + 1).Font.Size = 7
    oWS.Cells(mRow, c + 2) = Format(MSHFlexGrid3.TextMatrix(k, 3), "#######0.000")
    oWS.Cells(mRow, c + 2).Font.Size = 7
    oWS.Cells(mRow, c + 3) = Format(MSHFlexGrid3.TextMatrix(k, 4), "#######0.000")
    oWS.Cells(mRow, c + 3).Font.Size = 7
    oWS.Cells(mRow, c + 4) = Format(MSHFlexGrid3.TextMatrix(k, 5), "#######0.000")
    oWS.Cells(mRow, c + 4).Font.Size = 7
    oWS.Cells(mRow, c + 5) = MSHFlexGrid3.TextMatrix(k, 2)
    oWS.Cells(mRow, c + 5).Font.Size = 7
    oWS.Cells(mRow, c + 6) = MSHFlexGrid3.TextMatrix(k, 6)
    oWS.Cells(mRow, c + 6).Font.Size = 7
    mRow = mRow + 1
Next k
mRow = mRow + 1

tmp = "a" & mRow & ":" & "c" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = "Total"
oWS.Cells(mRow, c + 1).Font.Size = 7
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108

oWS.Cells(mRow, c + 4) = Format(MSHFlexGrid3.TextMatrix(0, 5), "#######0.000")
oWS.Cells(mRow, c + 4).Font.Size = 7

tmp = "e" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge

tmp = "e" & mRow & ":" & "e" & (mRow + 1)
oWS.Range(tmp).Merge

oWS.Cells(mRow, c + 5) = "Affix Hologram here."
oWS.Cells(mRow, c + 5).Font.Size = 7
oWS.Cells(mRow, c + 5).Font.Bold = True


mRow = mRow + 1

tmp = "a" & mRow & ":" & "c" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = "Less Gunny Weight(MT)"
oWS.Cells(mRow, c + 1).Font.Size = 7
oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, c + 1).Font.Bold = True

oWS.Cells(mRow, c + 4) = Format(MSHFlexGrid4.TextMatrix(0, 9), "#######0.000")
oWS.Cells(mRow, c + 4).Font.Size = 7

tmp = "e" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge

tmp = "e" & mRow & ":" & "e" & (mRow + 1)
oWS.Range(tmp).Merge

mRow = mRow + 1

tmp = "a" & mRow & ":" & "c" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = "Less Standard Deductions(MT)"
oWS.Cells(mRow, c + 1).Font.Size = 7
oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, c + 1).Font.Bold = True

oWS.Cells(mRow, c + 4) = "" '
oWS.Cells(mRow, c + 4).Font.Size = 7

tmp = "e" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge

tmp = "e" & mRow & ":" & "e" & (mRow + 1)
oWS.Range(tmp).Merge

mRow = mRow + 1

tmp = "a" & mRow & ":" & "c" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = "Add/Less Adjustments(MT)"
oWS.Cells(mRow, c + 1).Font.Size = 7
oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, c + 1).Font.Bold = True

oWS.Cells(mRow, c + 4) = Format(MSHFlexGrid4.TextMatrix(0, 8), "#######0.000")
oWS.Cells(mRow, c + 4).Font.Size = 7

tmp = "e" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge

If mExchangeTypeID = 2 Then
   tmp = "e" & mRow & ":" & "e" & (mRow + 1)
   oWS.Range(tmp).Merge
End If
mRow = mRow + 1


tmp = "a" & mRow & ":" & "c" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = "Electronic Weight(MT)"
oWS.Cells(mRow, c + 1).Font.Size = 7
oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, c + 1).Font.Bold = True

oWS.Cells(mRow, c + 4) = Format(MSHFlexGrid4.TextMatrix(0, 10), "#######0.000")
oWS.Cells(mRow, c + 4).Font.Size = 7

mRow = mRow + 1

If mExchangeTypeID = 2 Then
   tmp = "a" & mRow & ":" & "c" & mRow
   oWS.Range(tmp).Merge
   oWS.Cells(mRow, c + 1) = "Total No. of Bags"
   oWS.Cells(mRow, c + 1).Font.Size = 7
   oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
   oWS.Cells(mRow, c + 1).Font.Bold = True

   oWS.Cells(mRow, c + 4) = Format(MSHFlexGrid4.TextMatrix(0, 6)) 'Format(MSHFlexGrid4.TextMatrix(0, 8), "#######0.000")
   oWS.Cells(mRow, c + 4).Font.Size = 7
'   tmp = "e" & mRow & ":" & "f" & mRow
'   oWS.Range(tmp).Merge
   
'   tmp = "e" & mRow & ":" & "e" & (mRow + 1)
'   oWS.Range(tmp).Merge
   
   
   oWS.Cells(mRow, c + 5) = "Hologram No"
   oWS.Cells(mRow, c + 5).Font.Size = 7
   oWS.Cells(mRow, c + 5).Font.Bold = True

   oWS.Cells(mRow, c + 6) = RsTxn_CxTF_Details!HologramNo
   oWS.Cells(mRow, c + 6).Font.Size = 7
   oWS.Cells(mRow, c + 6).HorizontalAlignment = -4108
   mRow = mRow + 1
   
Else
   oWS.Cells(mRow - 1, c + 5) = "Hologram No"
   oWS.Cells(mRow - 1, c + 5).Font.Size = 7
   oWS.Cells(mRow - 1, c + 5).Font.Bold = True

   oWS.Cells(mRow - 1, c + 6) = RsTxn_CxTF_Details!HologramNo
   oWS.Cells(mRow - 1, c + 6).Font.Size = 7
   oWS.Cells(mRow - 1, c + 6).HorizontalAlignment = -4108
End If




mRow = mRow + 1
mRow = mRow + 1
mRow = mRow + 1

tmp = "a" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = "Name & Signature of CMP Manager/CMP DRC:"
oWS.Cells(mRow, c + 1).Font.Size = 9
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).RowHeight = 40
mRow = mRow + 1

tmp = "a" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = "Received the above mentioned goods in good condition."
oWS.Cells(mRow, c + 1).Font.Size = 10
oWS.Cells(mRow, c + 1).WrapText = True
oWS.Cells(mRow, c + 1).Font.Name = "Zurich BT"

mRow = mRow + 1

tmp = "a" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
mRow = mRow + 1

tmp = "a" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
mRow = mRow + 1


tmp = "a" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge

mRow = mRow + 1

tmp = "a" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge

mRow = mRow + 1

tmp = "a" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = "Name & Signature of Withdrawer/Agent of Withdrawer:"
oWS.Cells(mRow, c + 1).Font.Size = 9
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).Font.Name = "Zurich BT"
k = mRow - 6

For Cell = 4 To k
    tmp = "a" & Cell & ":" & "f" & Cell
    oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
    oWS.Range(tmp).Font.Name = "Zurich BT"
    'oWS.Range(tmp).Font.Size = 7
    oWS.Range(tmp).WrapText = True
Next Cell
    
If Gen_ExcelPassword <> "" Then
    oWS.Protect Gen_ExcelPassword, , , , , , True, True
End If
oWB.Save
Call Log_Print_History(TxtPreFix, "CITF", TxtCxTFNo, Gen_FLName)
    
    
'oWB.Close
'oXL.Quit
MsgBox ("Excel File Created !!! ")
oXL.Visible = True





End Sub


Private Sub TxtHologramNo_LostFocus()
If TxtHologramNo = "" Then Exit Sub
If RsMst_Hologram.RecordCount > 0 Then
   RsMst_Hologram.MoveFirst
   RsMst_Hologram.Find "HologramNo = '" & TxtHologramNo & "' "
End If
If Not RsMst_Hologram.EOF Then
   
   If RsMst_Hologram!DocType = "WHR" Then
      tmp = "Select CMPID from Txn_WHR_Detail Where WHRNo = " & RsMst_Hologram!DocNo & ""
      Call Tmp4Table
      If TmpRs4.RecordCount > 0 Then
         MsgBox "Hologram Already Used In " & RsMst_Hologram!DocType & " " & TmpRs4!CMPID & "-" & (Val(RsMst_Hologram!DocNo) - TmpRs4!CMPID * 1000000)
         TmpRs4.Close
      End If
   Else
      MsgBox "Hologram Already Used In " & RsMst_Hologram!DocType & "-" & RsMst_Hologram!DocNo
   End If
   TxtHologramNo.SetFocus
   CmdSaveHologram.Enabled = False
   Exit Sub
End If
CmdSaveHologram.Enabled = True
CmdSaveHologram.SetFocus
End Sub

Private Sub CmdSaveHologram_Click()
Dim ans As Variant
If TxtHologramNo = "" Then
   MsgBox "Blank Hologram No not allowed!!!"
   TxtHologramNo.SetFocus
   Exit Sub
End If
ans = MsgBox("Are you Sure you want to save this HologramNo", vbYesNo)
If ans = vbYes Then
   RsMst_Hologram.AddNew
   RsMst_Hologram!HologramNo = TxtHologramNo
   RsMst_Hologram!DocType = "CITF"
   RsMst_Hologram!DocNo = TxtCxTFNo
   RsMst_Hologram.Update
   RsTxn_CxTF_Details!HologramNo = TxtHologramNo
   RsTxn_CxTF_Details.Update
   TxtHologramNo = ""
   Frame3.Visible = False
   RsTxn_CxTF_Details!Flag = "P"
   RsTxn_CxTF_Details.Update
   Gen_mTimeStamp = GetTimeStamp
   Call Create_Xls(Gen_mTimeStamp & Gen_UserName & "_CITF-" & Trim(TxtPreFix) & "" & Trim(TxtCxTFNo) & ".xls")
   Call Xls_Print_Rpt
Else
   TxtHologramNo.SetFocus
End If

End Sub
