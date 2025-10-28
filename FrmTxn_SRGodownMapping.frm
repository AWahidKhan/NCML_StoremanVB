VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmTxn_SRGodownMapping 
   Caption         =   "Storage Reservation Godown Mapping"
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
      Height          =   4095
      Left            =   3600
      TabIndex        =   65
      Top             =   1200
      Visible         =   0   'False
      Width           =   15015
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
         Left            =   10440
         Sorted          =   -1  'True
         TabIndex        =   76
         Top             =   420
         Width           =   4485
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
         Left            =   1875
         Sorted          =   -1  'True
         TabIndex        =   75
         Top             =   420
         Width           =   3945
      End
      Begin VB.ComboBox CmbSClient 
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
         Left            =   5880
         Sorted          =   -1  'True
         TabIndex        =   74
         Top             =   420
         Width           =   4500
      End
      Begin VB.TextBox TxtSContractID 
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
         Left            =   135
         TabIndex        =   73
         Top             =   420
         Width           =   1650
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
         Left            =   9600
         TabIndex        =   68
         Top             =   840
         Width           =   1125
      End
      Begin VB.TextBox TxtSearchDescription 
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
         Left            =   3120
         TabIndex        =   67
         Top             =   840
         Width           =   4815
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
         Height          =   345
         Left            =   8400
         TabIndex        =   66
         Top             =   840
         Width           =   1125
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   2460
         Left            =   120
         TabIndex        =   69
         Top             =   1380
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   4339
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
      Begin VB.Label Label39 
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
         Height          =   255
         Left            =   12142
         TabIndex        =   80
         Top             =   120
         Width           =   1080
      End
      Begin VB.Label Label42 
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
         Height          =   225
         Left            =   3435
         TabIndex        =   79
         Top             =   120
         Width           =   825
      End
      Begin VB.Label Label5 
         Caption         =   "Client"
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
         Left            =   7838
         TabIndex        =   78
         Top             =   120
         Width           =   585
      End
      Begin VB.Label Label20 
         Caption         =   "Reservation ID"
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
         Left            =   233
         TabIndex        =   77
         Top             =   180
         Width           =   1455
      End
      Begin VB.Label Label19 
         Caption         =   "Search By Storage Description"
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
         TabIndex        =   72
         Top             =   900
         Width           =   2895
      End
      Begin VB.Label Label22 
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
         Left            =   13680
         TabIndex        =   71
         Top             =   900
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
         Left            =   11160
         TabIndex        =   70
         Top             =   885
         Width           =   2415
      End
   End
   Begin VB.Frame Frame3 
      Height          =   8775
      Left            =   120
      TabIndex        =   17
      Top             =   0
      Width           =   11085
      Begin VB.TextBox TxtMappingID 
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
         Left            =   1815
         Locked          =   -1  'True
         TabIndex        =   56
         Top             =   165
         Width           =   1680
      End
      Begin VB.TextBox TxtBookedTo 
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
         Left            =   5235
         Locked          =   -1  'True
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   3000
         Width           =   1320
      End
      Begin VB.TextBox TxtBookedFrom 
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   3000
         Width           =   1320
      End
      Begin VB.CommandButton CmdStorageReservation 
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
         Height          =   315
         Left            =   8280
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   600
         Width           =   375
      End
      Begin VB.ComboBox CmbStorageDescription 
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
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   570
         Width           =   6360
      End
      Begin VB.TextBox TxtStorageDescription 
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
         Locked          =   -1  'True
         MaxLength       =   150
         TabIndex        =   25
         Top             =   570
         Width           =   6360
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
         Height          =   360
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   24
         Top             =   1365
         Width           =   6360
      End
      Begin VB.TextBox TxtBookedAsPer 
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
         Left            =   6015
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   1770
         Width           =   2130
      End
      Begin VB.TextBox TxtBookedRate 
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   2160
         Width           =   1320
      End
      Begin VB.TextBox TxtStorageCharges 
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   2580
         Width           =   4170
      End
      Begin VB.TextBox TxtBookedQty 
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
         Left            =   4635
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   2160
         Width           =   1320
      End
      Begin VB.TextBox TxtLocationID 
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   960
         Width           =   6360
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
         Height          =   360
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1770
         Width           =   2640
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   5310
         Left            =   120
         TabIndex        =   37
         Top             =   3390
         Width           =   10860
         _ExtentX        =   19156
         _ExtentY        =   9366
         _Version        =   393216
         Tabs            =   2
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
         TabCaption(0)   =   "Godown Details"
         TabPicture(0)   =   "FrmTxn_SRGodownMapping.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "MSHFlexGrid2"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Frame2"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Log History"
         TabPicture(1)   =   "FrmTxn_SRGodownMapping.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label25"
         Tab(1).Control(1)=   "Label24"
         Tab(1).Control(2)=   "TxtUpdated"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "TxtCreated"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).ControlCount=   4
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
            Left            =   -73215
            Locked          =   -1  'True
            TabIndex        =   60
            TabStop         =   0   'False
            Top             =   1080
            Width           =   8535
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
            Left            =   -73215
            Locked          =   -1  'True
            TabIndex        =   59
            TabStop         =   0   'False
            Top             =   1845
            Width           =   8535
         End
         Begin VB.Frame Frame2 
            Height          =   2535
            Left            =   120
            TabIndex        =   38
            Top             =   360
            Width           =   10680
            Begin VB.TextBox TxtDetailsID 
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               Height          =   360
               Left            =   1725
               Locked          =   -1  'True
               TabIndex        =   63
               TabStop         =   0   'False
               Top             =   2040
               Visible         =   0   'False
               Width           =   1200
            End
            Begin VB.CommandButton CmdGodownNo 
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
               Height          =   315
               Left            =   9990
               TabIndex        =   53
               TabStop         =   0   'False
               Top             =   173
               Width           =   375
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
               Left            =   7785
               Sorted          =   -1  'True
               TabIndex        =   4
               Top             =   150
               Width           =   2145
            End
            Begin VB.ComboBox CmbCMPName 
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
               Left            =   1725
               Sorted          =   -1  'True
               TabIndex        =   3
               Top             =   150
               Width           =   4560
            End
            Begin VB.CommandButton CmdCMPName 
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
               Height          =   315
               Left            =   6360
               TabIndex        =   44
               TabStop         =   0   'False
               Top             =   173
               Width           =   375
            End
            Begin VB.TextBox TxtGodownAddress 
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               Height          =   630
               Left            =   1725
               Locked          =   -1  'True
               MaxLength       =   255
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   43
               TabStop         =   0   'False
               Top             =   555
               Width           =   8235
            End
            Begin VB.TextBox TxtGodownCapacity 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               Height          =   360
               Left            =   1725
               Locked          =   -1  'True
               TabIndex        =   42
               TabStop         =   0   'False
               Top             =   1635
               Width           =   2520
            End
            Begin VB.TextBox TxtBagsCapacity 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               Height          =   360
               Left            =   6105
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   41
               TabStop         =   0   'False
               Top             =   1230
               Width           =   1875
            End
            Begin VB.TextBox TxtUnitTypeID 
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               Height          =   360
               Left            =   1725
               Locked          =   -1  'True
               TabIndex        =   40
               TabStop         =   0   'False
               Top             =   1230
               Width           =   2520
            End
            Begin VB.TextBox TxtReservationQty 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   360
               Left            =   6120
               MaxLength       =   50
               TabIndex        =   5
               Top             =   1635
               Width           =   1845
            End
            Begin VB.CommandButton CmdAddToList 
               Caption         =   "Add To List"
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
               Left            =   6360
               TabIndex        =   6
               Top             =   2040
               Width           =   1695
            End
            Begin VB.CommandButton CmdRemoveFromList 
               Caption         =   "Remove From List"
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
               Left            =   8070
               TabIndex        =   39
               Top             =   2040
               Width           =   1935
            End
            Begin VB.Label Label4 
               Caption         =   "Details ID"
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
               TabIndex        =   64
               Top             =   2070
               Visible         =   0   'False
               Width           =   1440
            End
            Begin VB.Label lblGodownNo 
               Caption         =   "Godown No"
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
               Left            =   6915
               TabIndex        =   51
               Top             =   210
               Width           =   1095
            End
            Begin VB.Label LblCMPName 
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
               Height          =   300
               Left            =   75
               TabIndex        =   50
               Top             =   180
               Width           =   1800
            End
            Begin VB.Label Label1 
               Caption         =   "Godown Address"
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
               Left            =   75
               TabIndex        =   49
               Top             =   720
               Width           =   1680
            End
            Begin VB.Label LblGodownCapacity 
               Caption         =   "Capacity in MT"
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
               Left            =   90
               TabIndex        =   48
               Top             =   1665
               Width           =   1440
            End
            Begin VB.Label LblBagsCapacity 
               Caption         =   "Capacity in Bags"
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
               Left            =   4410
               TabIndex        =   47
               Top             =   1260
               Width           =   1560
            End
            Begin VB.Label LblUnitTypeID 
               Caption         =   "Unit Type"
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
               TabIndex        =   46
               Top             =   1260
               Width           =   960
            End
            Begin VB.Label Label11 
               Caption         =   "Reservation Qty"
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
               Left            =   4410
               TabIndex        =   45
               Top             =   1710
               Width           =   1770
            End
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
            Height          =   2280
            Left            =   120
            TabIndex        =   52
            Top             =   2940
            Width           =   10590
            _ExtentX        =   18680
            _ExtentY        =   4022
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
         Begin VB.Label Label24 
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
            TabIndex        =   62
            Top             =   1140
            Width           =   1695
         End
         Begin VB.Label Label25 
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
            TabIndex        =   61
            Top             =   1905
            Width           =   975
         End
      End
      Begin MSComCtl2.DTPicker TxtMappingDate 
         Height          =   375
         Left            =   5160
         TabIndex        =   1
         Top             =   165
         Width           =   1560
         _ExtentX        =   2752
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
         Format          =   3997697
         CurrentDate     =   39884
      End
      Begin VB.Label Label8 
         Caption         =   "Mapping Date"
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
         Left            =   3840
         TabIndex        =   58
         Top             =   195
         Width           =   1800
      End
      Begin VB.Label Label3 
         Caption         =   "Mapping ID"
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
         TabIndex        =   57
         Top             =   195
         Width           =   1800
      End
      Begin VB.Label Label2 
         Caption         =   "Storage Description"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   120
         TabIndex        =   35
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label LblClientID 
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
         Height          =   225
         Left            =   120
         TabIndex        =   34
         Top             =   1440
         Width           =   1800
      End
      Begin VB.Label LblLocationID 
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
         Height          =   225
         Left            =   120
         TabIndex        =   33
         Top             =   1035
         Width           =   1800
      End
      Begin VB.Label LblCommodityID 
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
         Height          =   225
         Left            =   120
         TabIndex        =   32
         Top             =   1845
         Width           =   1800
      End
      Begin VB.Label LblBookedFrom 
         Caption         =   "Booked From"
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
         TabIndex        =   31
         Top             =   3015
         Width           =   1800
      End
      Begin VB.Label LblBookedTo 
         Caption         =   "Extended Upto"
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
         Left            =   3480
         TabIndex        =   30
         Top             =   3030
         Width           =   1560
      End
      Begin VB.Label LblBookedAsPer 
         Caption         =   "Booked As Per"
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
         Left            =   4635
         TabIndex        =   29
         Top             =   1800
         Width           =   1425
      End
      Begin VB.Label LblBookedRate 
         Caption         =   "Booked Rate Per Month"
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
         Left            =   120
         TabIndex        =   28
         Top             =   2100
         Width           =   1575
      End
      Begin VB.Label LblStorageCharges 
         Caption         =   "Storage Charges"
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
         TabIndex        =   27
         Top             =   2640
         Width           =   1800
      End
      Begin VB.Label LblBookedQty 
         Caption         =   "Booked Qty"
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
         Left            =   3480
         TabIndex        =   26
         Top             =   2190
         Width           =   1335
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1065
      Left            =   120
      TabIndex        =   0
      Top             =   8760
      Width           =   11085
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
         Left            =   225
         TabIndex        =   16
         Top             =   180
         Width           =   2115
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
         Left            =   225
         TabIndex        =   15
         Top             =   585
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
         Left            =   2340
         TabIndex        =   7
         Top             =   180
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
         Left            =   2340
         TabIndex        =   14
         Top             =   585
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
         Left            =   4455
         TabIndex        =   13
         Top             =   180
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
         Left            =   4455
         TabIndex        =   12
         Top             =   585
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
         Left            =   6570
         TabIndex        =   11
         Top             =   180
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
         Left            =   6570
         TabIndex        =   10
         Top             =   585
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
         Left            =   8685
         TabIndex        =   8
         Top             =   180
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
         Left            =   8685
         TabIndex        =   9
         Top             =   585
         Width           =   2115
      End
   End
End
Attribute VB_Name = "FrmTxn_SRGodownMapping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg, mLocationID As Variant
Dim mCMPID, mGodownID, mStorageReservationDate As Variant
Dim mContractID As Double
Public Sub Grid_Head()
MSHFlexGrid2.Clear
MSHFlexGrid2.Cols = 8
MSHFlexGrid2.Rows = 3
MSHFlexGrid2.FixedRows = 2

MSHFlexGrid2.TextMatrix(1, 0) = "CMP"
MSHFlexGrid2.TextMatrix(1, 1) = "Godown"
MSHFlexGrid2.TextMatrix(1, 2) = "Unit Type"
MSHFlexGrid2.TextMatrix(1, 3) = "Capacity in Bags"
MSHFlexGrid2.TextMatrix(1, 4) = "Capacity in MT"
MSHFlexGrid2.TextMatrix(1, 5) = "Reservation Qty"
MSHFlexGrid2.TextMatrix(1, 6) = "GodownID"
'MSHFlexGrid2.TextMatrix(1, 7) = "SRGodownMapDetailID"
MSHFlexGrid2.TextMatrix(1, 7) = "Address"

MSHFlexGrid2.ColWidth(0) = 3000
MSHFlexGrid2.ColWidth(1) = 2400
MSHFlexGrid2.ColWidth(2) = 1600
MSHFlexGrid2.ColWidth(3) = 1000
MSHFlexGrid2.ColWidth(4) = 1000
MSHFlexGrid2.ColWidth(5) = 1200
MSHFlexGrid2.ColWidth(6) = 0
'MSHFlexGrid2.ColWidth(7) = 0
MSHFlexGrid2.ColWidth(7) = 0

MSHFlexGrid2.RowHeight(1) = 500

End Sub

Private Sub AddCmd_Click()

Edit_Flg = "A"
TxtMappingID = ""
'TxtMappingDate.Enabled = True
TxtMappingDate = Date

CmbStorageDescription.Visible = True
TxtStorageDescription.Visible = False
CmbStorageDescription.Enabled = True
CmbStorageDescription.Text = ""

CmdStorageReservation.Visible = True
Call ClearContractDetails
CmbCMPName.Text = ""
Call ClearGodownDetails
Frame2.Enabled = True

TxtCreated = ""
TxtUpdated = ""
SSTab1.Tab = 0
If RsTxn_SRGodownMapping.RecordCount > 0 Then
   If CmbStorageDescription.Enabled = True Then
      CmbStorageDescription.SetFocus
   End If
End If

Call Grid_Head
Call GButtonLock(Me, False)

End Sub

Private Sub CmbSClient_GotFocus()

tmp = CmbSClient.Text
Call TableMst_Client("Where ExchangeTypeID = " & GetExchangeID("Non NCDEX"))
CmbSClient.Clear
If RsMst_Client.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Client.RecordCount
    CmbSClient.AddItem RsMst_Client!ClientName
    RsMst_Client.MoveNext
Next
CmbSClient.Text = tmp


End Sub

Private Sub CmbSClient_LostFocus()
If CmbSClient.Text = "" Then
   Exit Sub
End If
RsMst_Client.MoveFirst
RsMst_Client.Find "ClientName = '" & CmbSClient.Text & "'"
If RsMst_Client.EOF Then
   MsgBox "Invalid selection "
   CmbSClient.SetFocus
   Exit Sub
End If

End Sub

Private Sub CmbSCommodity_GotFocus()
tmp = CmbSCommodity.Text
Call TableMst_Commodity
CmbSCommodity.Clear
If RsMst_Commodity.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Commodity.RecordCount
    CmbSCommodity.AddItem RsMst_Commodity!Commodity
    RsMst_Commodity.MoveNext
Next
CmbSCommodity.Text = tmp
End Sub

Private Sub CmbSCommodity_LostFocus()
If CmbSCommodity.Text = "" Then
   Exit Sub
End If
RsMst_Commodity.MoveFirst
RsMst_Commodity.Find "Commodity = '" & CmbSCommodity.Text & "'"
If RsMst_Commodity.EOF Then
   MsgBox "Invalid selection "
   CmbSCommodity.SetFocus
   Exit Sub
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

Private Sub CmbStorageDescription_GotFocus()
tmp = CmbStorageDescription.Text
Call TableMst_StorageContract '(" Where CloseDate is null And LocationID = " & mLocationID)
CmbStorageDescription.Clear
If RsMst_StorageContract.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_StorageContract.RecordCount
    CmbStorageDescription.AddItem RsMst_StorageContract!ContractName
    RsMst_StorageContract.MoveNext
Next
CmbStorageDescription.Text = tmp

End Sub
Private Sub CmbStorageDescription_LostFocus()

If CmbStorageDescription.Text = "" Then
   Call ClearContractDetails
   CmbCMPName.Text = ""
   Call ClearGodownDetails
   mContractID = 0
   Exit Sub
End If

RsMst_StorageContract.MoveFirst
RsMst_StorageContract.Find "ContractName = '" & CmbStorageDescription.Text & "'"
If RsMst_StorageContract.EOF Then
   MsgBox "Invalid selection "
   CmbStorageDescription.SetFocus
   Exit Sub
End If
If mContractID <> RsMst_StorageContract!ContractID Then
   Call ClearContractDetails
   CmbCMPName.Text = ""
   Call ClearGodownDetails
   Call Grid_Head
End If
mContractID = RsMst_StorageContract!ContractID



Call GetContractDetails(mContractID)

End Sub
Private Sub GetContractDetails(mContractID_ As Double)


'If CmbStorageDescription.Text = "" Then Exit Sub

tmp = "Select MSC.ContractID,MC.ClientName,ML.LocationID,ML.LocationName,MComm.Commodity,MSC.BookedFrom,MSC.BookedTo,MSC.BookedAsPer, " & _
      " MSC.BookedRate , MSC.StorageCharges, MSC.BookedQty, MSC.ContractName,MSC.ContractDate,MSC.ExtendedPeriod,MSC.ContractName " & _
      " From Mst_StorageContract MSC " & _
      " Inner Join Mst_Client MC On MSC.ClientID = MC.ClientID " & _
      " Inner Join Mst_Location ML On MSC.LocationID = ML.LocationID " & _
      " Left Join Mst_Commodity MComm On MSC.CommodityID = MComm.CommodityID  " & _
      " Where MSC.ContractID = " & mContractID_ & ""

Call TmpTable

If TmpRs.RecordCount > 0 Then
   TxtStorageDescription = IIf(IsNull(TmpRs!ContractName), "", TmpRs!ContractName)
   TxtLocationID = IIf(IsNull(TmpRs!LocationName), "", TmpRs!LocationName)
   TxtClientID = IIf(IsNull(TmpRs!ClientName), "", TmpRs!ClientName)
   TxtCommodityID = IIf(IsNull(TmpRs!Commodity), "", TmpRs!Commodity)
   TxtBookedAsPer = IIf(IsNull(TmpRs!BookedAsPer), "", TmpRs!BookedAsPer)
   TxtBookedRate = IIf(IsNull(TmpRs!BookedRate), 0, TmpRs!BookedRate)
   TxtBookedQty = IIf(IsNull(TmpRs!BookedQty), "", TmpRs!BookedQty)
   TxtStorageCharges = IIf(IsNull(TmpRs!StorageCharges), "", TmpRs!StorageCharges)
   TxtBookedFrom = IIf(IsNull(TmpRs!BookedFrom), "", TmpRs!BookedFrom)
   TxtBookedTo = IIf(IsNull(TmpRs!ExtendedPeriod), "", TmpRs!ExtendedPeriod)
   mLocationID = TmpRs!LocationID
   mStorageReservationDate = TmpRs!ContractDate
End If


End Sub
Private Sub GetGodownDetails(mCMPID_ As Variant, mGodownID_ As Variant)

If CmbCMPName.Text = "" Then Exit Sub
If CmbGodownNo.Text = "" Then Exit Sub

tmp = "Select MG.GodownID,MUT.UnitType,MG.BagsCapacity,MG.GodownCapacity,MG.Address " & _
       " From Mst_Godown MG " & _
       " Inner Join Mst_UnitType MUT On MG.UnitTypeID = MUT.UnitTypeID " & _
       " Where MG.GodownID = " & mGodownID_ & " And CMPID = " & mCMPID_ & ""


Call TmpTable

If TmpRs.RecordCount > 0 Then

   TxtGodownAddress = IIf(IsNull(TmpRs!Address), "", TmpRs!Address)
   TxtUnitTypeID = IIf(IsNull(TmpRs!UnitType), "", TmpRs!UnitType)
   TxtGodownCapacity = IIf(IsNull(TmpRs!GodownCapacity), "", TmpRs!GodownCapacity)
   TxtBagsCapacity = IIf(IsNull(TmpRs!BagsCapacity), "", TmpRs!BagsCapacity)

End If


End Sub

Private Sub ClearContractDetails()

TxtLocationID = "" 'IIf(IsNull(TmpRs!LocationName), "", TmpRs!LocationName)"
TxtClientID = "" 'IIf(IsNull(TmpRs!ClientName), "", TmpRs!ClientName)
TxtCommodityID = "" 'IIf(IsNull(TmpRs!Commodity), "", TmpRs!Commodity)
TxtBookedAsPer = "" 'IIf(IsNull(TmpRs!BookedAsPer), "", TmpRs!BookedAsPer)
TxtBookedRate = "" 'IIf(IsNull(TmpRs!BookedRate), 0, TmpRs!BookedRate)
TxtBookedQty = "" 'IIf(IsNull(TmpRs!BookedQty), "", TmpRs!BookedQty)
TxtStorageCharges = "" 'IIf(IsNull(TmpRs!StorageCharges), "", TmpRs!StorageCharges)
TxtBookedFrom = "" 'IIf(IsNull(TmpRs!BookedFrom), "", TmpRs!BookedFrom)
TxtBookedTo = "" 'IIf(IsNull(TmpRs!BookedTo), "", TmpRs!BookedTo)
mLocationID = Null 'TmpRs!LocationID

End Sub

Private Sub ClearGodownDetails()

'CmbCMPName.Text = ""
CmbGodownNo.Text = ""
TxtUnitTypeID = ""
TxtGodownAddress = ""
TxtGodownCapacity = ""
TxtBagsCapacity = ""
TxtReservationQty = ""
CmdAddToList.Caption = "Add To List"
CmdRemoveFromList.Enabled = False


End Sub

Private Sub CmdAddtoList_Click()

If CmbCMPName.Text = "" Then
   MsgBox "Blank CMP Not Allowed!!!"
   CmbCMPName.SetFocus
   Exit Sub
End If

If CmbGodownNo.Text = "" Then
   MsgBox "Blank Godown Not Allowed!!!"
   CmbGodownNo.SetFocus
   Exit Sub
End If

'If Val(TxtReservationQty) = 0 Then
'   MsgBox "Blank Resevation Qty Not Allowed!!!"
'   TxtReservationQty.SetFocus
'   Exit Sub
'End If

If TxtReservationQty = "" Then
   MsgBox "Blank Resevation Qty Not Allowed!!!"
   TxtReservationQty.SetFocus
   Exit Sub
End If


Dim I, ans As Variant

If Val(TxtReservationQty) > Val(TxtGodownCapacity) Then
   ans = MsgBox("Warning :- Reservation Qty is more than Godown Capacity (in MT). Are you Sure you want to continue", vbYesNo)
   If ans = vbNo Then
      TxtReservationQty.SetFocus
      Exit Sub
   End If
End If

ans = Validate_DuplicateGrid()
If ans = False Then
   MsgBox "Duplicate Entry !!!! "
   Exit Sub
End If

If CmdAddToList.Caption = "Add To List" Then
   I = MSHFlexGrid2.Rows - 1
   MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
Else
   I = MSHFlexGrid2.RowSel
   MSHFlexGrid2.TextMatrix(0, 3) = Val(MSHFlexGrid2.TextMatrix(0, 3)) - Val(MSHFlexGrid2.TextMatrix(I, 3))
   MSHFlexGrid2.TextMatrix(0, 4) = Val(MSHFlexGrid2.TextMatrix(0, 4)) - Val(MSHFlexGrid2.TextMatrix(I, 4))
   MSHFlexGrid2.TextMatrix(0, 5) = Val(MSHFlexGrid2.TextMatrix(0, 5)) - Val(MSHFlexGrid2.TextMatrix(I, 5))
End If


MSHFlexGrid2.TextMatrix(I, 0) = CmbCMPName.Text  '"CMP"
MSHFlexGrid2.TextMatrix(I, 1) = CmbGodownNo.Text  '"Godown"
MSHFlexGrid2.TextMatrix(I, 2) = TxtUnitTypeID '"Unit Type"
MSHFlexGrid2.TextMatrix(I, 3) = Val(TxtBagsCapacity) '"Capacity in Bags"
MSHFlexGrid2.TextMatrix(I, 4) = Val(TxtGodownCapacity) '  "Capacity in MT"
MSHFlexGrid2.TextMatrix(I, 5) = Val(TxtReservationQty) '"Reservation Qty"
MSHFlexGrid2.TextMatrix(I, 6) = mGodownID '"GodownID"
'MSHFlexGrid2.TextMatrix(i, 7) = TxtDetailsID '"" '"SRGodownMapDetailID"
MSHFlexGrid2.TextMatrix(I, 7) = TxtGodownAddress


MSHFlexGrid2.TextMatrix(0, 3) = Val(MSHFlexGrid2.TextMatrix(0, 3)) + Val(MSHFlexGrid2.TextMatrix(I, 3))
MSHFlexGrid2.TextMatrix(0, 4) = Val(MSHFlexGrid2.TextMatrix(0, 4)) + Val(MSHFlexGrid2.TextMatrix(I, 4))
MSHFlexGrid2.TextMatrix(0, 5) = Val(MSHFlexGrid2.TextMatrix(0, 5)) + Val(MSHFlexGrid2.TextMatrix(I, 5))



CmdAddToList.Caption = "Add To List"
CmbCMPName.Text = ""
Call ClearGodownDetails
CmbCMPName.SetFocus

End Sub

Private Sub CmdCMPName_Click()
FrmMst_CMP.Show
End Sub
Private Sub CmdGo_Click()
Dim wc As Variant


Call Grid_Head1

tmp = " Select TSRGM.SRGodownMappingID,TSRGM.SRGodownMappingDate, " & _
      " MSC.ContractName , ML.LocationName, MC.ClientName, MComm.Commodity, MSC.StorageCharges,MSC.ContractID,MSC.BookedFrom ,MSC.ExtendedPeriod," & _
      " c.CMPName,g.GodownNo,MU.UnitType,g.BagsCapacity,g.GodownCapacity,TSRGMD.ReservationQty  " & _
      " From Txn_SRGodownMapping TSRGM  " & _
      " Inner Join Mst_StorageContract MSC On TSRGM.ContractID = MSC.ContractID  " & _
      " Inner Join Mst_Location ML On MSC.LocationID = ML.LocationID  " & _
      " left Join Mst_Commodity MComm On MSC.CommodityID = MComm.CommodityID  " & _
      " Inner Join Mst_Client MC On MSC.ClientID = MC.ClientID  " & _
      " Inner Join Txn_SRGodownMapDetails TSRGMD On TSRGM.SRGodownMappingID = TSRGMD.SRGodownMappingID " & _
      " Inner Join Mst_Godown g on TSRGMD.GodownID = g.GodownID " & _
      " Inner Join Mst_CMP C on g.CMPID = c.CMPID " & _
      " Inner Join Mst_UnitType MU on g.UnitTypeID = MU.UnitTypeID "

wc = ""

If TxtSearchDescription <> "" Then
   wc = " Where MSC.ContractName like '%" & TxtSearchDescription & "%'"
End If


If CmbSLocation.Text <> "" Then
   If wc <> "" Then
      wc = wc & " And ML.LocationName = '" & CmbSLocation.Text & "'"
   Else
      wc = " Where ML.LocationName = '" & CmbSLocation.Text & "'"
   End If
End If

If CmbSClient.Text <> "" Then
   If wc <> "" Then
      wc = wc & " And MC.ClientName = '" & CmbSClient.Text & "'"
   Else
      wc = " Where MC.ClientName = '" & CmbSClient.Text & "'"
   End If
End If

If CmbSCommodity.Text <> "" Then
   If wc <> "" Then
      wc = wc & " And MComm.Commodity = '" & CmbSCommodity.Text & "'"
   Else
      wc = " Where MComm.Commodity = '" & CmbSCommodity.Text & "'"
   End If
End If

If TxtSContractID <> "" Then
   If wc <> "" Then
      wc = wc & " And MSC.ContractID = " & Val(TxtSContractID) & ""
   Else
      wc = " Where MSC.ContractID = " & Val(TxtSContractID) & ""
   End If
End If

tmp1 = " Order by TSRGM.SRGodownMappingID "

tmp = tmp & wc & tmp1

Call TmpTable

Call Grid_Head1
For iu = 1 To TmpRs.RecordCount
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
    For C = 0 To MSHFlexGrid1.Cols - 1
       MSHFlexGrid1.TextMatrix(iu, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
    Next
    MSHFlexGrid1.TextMatrix(iu, 1) = Format(MSHFlexGrid1.TextMatrix(iu, 1), "dd-mmm-yyyy")
    MSHFlexGrid1.TextMatrix(iu, 8) = Format(MSHFlexGrid1.TextMatrix(iu, 8), "dd-mmm-yyyy")
    MSHFlexGrid1.TextMatrix(iu, 9) = Format(MSHFlexGrid1.TextMatrix(iu, 9), "dd-mmm-yyyy")
    TmpRs.MoveNext
Next

Label22.Refresh
Label22.Caption = TmpRs.RecordCount

End Sub

Private Sub CmdGodownNo_Click()
FrmMst_CMP.Show
End Sub

Private Sub CmdRemoveFromList_Click()
If MSHFlexGrid2.Rows < 3 Then Exit Sub
Dim ans As Variant
If CmbCMPName.Text = "" Then
   MsgBox "Please Select atleast one row from grid!!!"
   Exit Sub
End If
ans = MsgBox("Are you Sure you want to Delete this Record ", vbYesNo)
If ans = vbYes Then
   I = MSHFlexGrid2.RowSel
   If I > 0 Then
      If MSHFlexGrid2.TextMatrix(I, 0) <> "" Then
         
         MSHFlexGrid2.TextMatrix(0, 3) = Val(MSHFlexGrid2.TextMatrix(0, 3)) - Val(MSHFlexGrid2.TextMatrix(I, 3))
         MSHFlexGrid2.TextMatrix(0, 4) = Val(MSHFlexGrid2.TextMatrix(0, 4)) - Val(MSHFlexGrid2.TextMatrix(I, 4))
         MSHFlexGrid2.TextMatrix(0, 5) = Val(MSHFlexGrid2.TextMatrix(0, 5)) - Val(MSHFlexGrid2.TextMatrix(I, 5))

         For RO = I To MSHFlexGrid2.Rows - 2
             For C = 0 To MSHFlexGrid2.Cols - 1
                 MSHFlexGrid2.TextMatrix(RO, C) = MSHFlexGrid2.TextMatrix(RO + 1, C)
             Next
         Next
         MSHFlexGrid2.Rows = MSHFlexGrid2.Rows - 1
      End If
   End If
   CmbCMPName.Text = ""
   Call ClearGodownDetails
   'CmdAddToGrid.Caption = "Add To Grid"
   CmdRemoveFromList.Enabled = False
Else
   CmbCMPName.SetFocus
End If


End Sub

Private Sub CmdStorageReservation_Click()
FrmMst_StorageContract.Show
End Sub

Private Sub DeleteCmd_Click()
Dim ans As Variant
On Error GoTo msgError
ans = MsgBox("Are you Sure you want to Delete this Record ", vbYesNo)
If ans = vbYes Then
   tmp = "Delete from Txn_SRGodownMapDetails Where SRGodownMappingID = " & TxtMappingID
   Call TmpTable

'   tmp = "SELECT * from Mst_EmpCMPMapping where EmployeeID = " & TxtEmpID
'   Call TmpTable
   
   RsTxn_SRGodownMapping.Requery
   RsTxn_SRGodownMapping.Find "SRGodownMappingID = " & TxtMappingID
   If RsTxn_SRGodownMapping.EOF Then Exit Sub
   RsTxn_SRGodownMapping.Delete
   RsTxn_SRGodownMapping.Update
   If RsTxn_SRGodownMapping.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsTxn_SRGodownMapping.MoveNext
   If RsTxn_SRGodownMapping.EOF() Then
      RsTxn_SRGodownMapping.MoveLast
   End If
   Call Indata
   'TmpRs.Close
End If

Exit Sub
msgError:
MsgBox "Child Record Present "
RsTxn_SRGodownMapping.CancelUpdate
End Sub

Private Sub EditCmd_Click()

Edit_Flg = "E"

CmbStorageDescription.Visible = True
TxtStorageDescription.Visible = False
CmbStorageDescription.Enabled = False
CmbStorageDescription.Text = TxtStorageDescription

CmdStorageReservation.Visible = False

Call ClearGodownDetails
Frame2.Enabled = True

SSTab1.Tab = 0

Call GButtonLock(Me, False)
End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub Form_Load()

Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Call GButtonLockAD(Me)
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If
Call Grid_Head1
TxtSearchDescription = ""
Label22.Caption = ""
Call Grid_Head
Call TableTxn_SRGodownMapping
If RsTxn_SRGodownMapping.RecordCount = 0 Then
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

Private Sub CmbCMPName_GotFocus()
If TxtLocationID.Text = "" Then
   MsgBox "Select Storage Description!!!"
   CmbStorageDescription.SetFocus
   Exit Sub
End If
tmp = CmbCMPName.Text
Call TableMst_CMP(" Where CloseDate is null And LocationID = " & mLocationID)
CmbCMPName.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_CMP.RecordCount
    CmbCMPName.AddItem RsMst_CMP!CMPName
    RsMst_CMP.MoveNext
Next
CmbCMPName.Text = tmp

End Sub

Private Sub CmbCMPName_LostFocus()
If CmbCMPName.Text = "" Then
   CmbGodownNo.Text = ""
   mCMPID = 0
   Exit Sub
End If

RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPName = '" & CmbCMPName.Text & "'"
If RsMst_CMP.EOF Then
   MsgBox "Invalid selection "
   CmbCMPName.SetFocus
   Exit Sub
End If
If mCMPID <> RsMst_CMP!CMPID Then
   CmbGodownNo.Text = ""
   
   Call ClearGodownDetails
End If
mCMPID = RsMst_CMP!CMPID

End Sub

Private Sub CmbGodownNo_GotFocus()
If CmbCMPName.Text = "" Then
   MsgBox "Select CMP !!!! "
   CmbCMPName.SetFocus
   Exit Sub
End If
tmp = CmbGodownNo.Text
Call TableMst_Godown("Where CloseDate is null And CMPID = " & mCMPID)
CmbGodownNo.Clear
If RsMst_Godown.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Godown.RecordCount
    CmbGodownNo.AddItem RsMst_Godown!GodownNo
    RsMst_Godown.MoveNext
Next
CmbGodownNo.Text = tmp
End Sub

Private Sub CmbGodownNo_LostFocus()
If CmbGodownNo.Text = "" Then
   mGodownID = Null
   Exit Sub
End If
RsMst_Godown.MoveFirst
RsMst_Godown.Find "GodownNo = '" & CmbGodownNo.Text & "'"
If RsMst_Godown.EOF Then
   MsgBox "Invalid selection "
   If CmbGodownNo.Enabled = True Then
      CmbGodownNo.SetFocus
   End If
   Exit Sub
End If
If mGodownID <> RsMst_Godown!GodownID Then

End If
mGodownID = RsMst_Godown!GodownID
Call GetGodownDetails(mCMPID, mGodownID)

End Sub

Private Sub MSHFlexGrid1_Click()
Dim x As Variant
Dim mCaId As Variant

x = MSHFlexGrid1.RowSel

mCaId = MSHFlexGrid1.TextMatrix(x, 0)

RsTxn_SRGodownMapping.MoveFirst

If RsTxn_SRGodownMapping.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
RsTxn_SRGodownMapping.MoveFirst
RsTxn_SRGodownMapping.Find "SRGodownMappingID = " & Val(mCaId) & ""
If RsTxn_SRGodownMapping.EOF Then
   RsTxn_SRGodownMapping.MoveLast
End If
Frame0.Visible = False
Call Indata
End Sub

Private Sub MSHFlexGrid2_Click()
If SaveCmd.Enabled = False Then Exit Sub

Dim I As Variant

I = MSHFlexGrid2.RowSel
If MSHFlexGrid2.TextMatrix(I, 0) = "" Then Exit Sub
CmbCMPName.Text = MSHFlexGrid2.TextMatrix(I, 0)
mCMPID = GetCMPID(MSHFlexGrid2.TextMatrix(I, 0))
CmbGodownNo.Text = MSHFlexGrid2.TextMatrix(I, 1)
TxtUnitTypeID = MSHFlexGrid2.TextMatrix(I, 2)
TxtBagsCapacity = Val(MSHFlexGrid2.TextMatrix(I, 3))
TxtGodownCapacity = Val(MSHFlexGrid2.TextMatrix(I, 4))
TxtReservationQty = Val(MSHFlexGrid2.TextMatrix(I, 5))
mGodownID = MSHFlexGrid2.TextMatrix(I, 6)
'TxtDetailsID = MSHFlexGrid2.TextMatrix(i, 7)
TxtGodownAddress = MSHFlexGrid2.TextMatrix(I, 7)


CmdAddToList.Caption = "Update List"
CmdRemoveFromList.Enabled = True
CmbCMPName.SetFocus
 
End Sub

Private Function Validate_DuplicateGrid() As Boolean
Dim Retval As Boolean
Retval = True
If MSHFlexGrid2.Rows = 2 Then
   Retval = True
Else
    For I = 1 To MSHFlexGrid2.Rows - 1
        If MSHFlexGrid2.TextMatrix(I, 0) <> "" Then
           If CmdAddToList.Caption = "Update List" Then
              If I <> MSHFlexGrid2.RowSel Then
                 If Val(MSHFlexGrid2.TextMatrix(I, 6)) = mGodownID Then
                    Retval = False
                 End If
              End If
           Else
              If Val(MSHFlexGrid2.TextMatrix(I, 6)) = mGodownID Then
                 Retval = False
              End If
           End If
        End If
    Next
End If
Validate_DuplicateGrid = Retval
End Function

Private Sub SaveCmd_Click()

If CmbStorageDescription.Text = "" Then
   MsgBox "Blank Storage Description not allowed!!!"
   CmbStorageDescription.SetFocus
   Exit Sub
End If

If MSHFlexGrid2.Rows < 4 Then
   MsgBox "No godown map to this storage reservation !!!"
   CmbCMPName.SetFocus
   Exit Sub
End If

If Edit_Flg = "A" Then
   If RsTxn_SRGodownMapping.RecordCount > 0 Then
      RsTxn_SRGodownMapping.MoveFirst
   End If
   RsTxn_SRGodownMapping.Find "ContractID = " & mContractID & ""
   If Not RsTxn_SRGodownMapping.EOF Then
      MsgBox "Duplicate Contract for mapping not Allowed"
      CmbStorageDescription.SetFocus
      Exit Sub
   End If
   RsTxn_SRGodownMapping.AddNew
   RsTxn_SRGodownMapping!SRGodownMappingID = GetMaxMappingID
   TxtMappingID = RsTxn_SRGodownMapping!SRGodownMappingID
   RsTxn_SRGodownMapping!G_UID = GetGUID
Else
    RsTxn_SRGodownMapping.MoveFirst
    RsTxn_SRGodownMapping.Find "ContractID = " & mContractID & ""
    If Not RsTxn_SRGodownMapping.EOF Then
       If RsTxn_SRGodownMapping!SRGodownMappingID <> TxtMappingID Then
          MsgBox "Duplicate Contract for mapping not Allowed"
          CmbStorageDescription.SetFocus
          Exit Sub
       End If
    End If
   RsTxn_SRGodownMapping.MoveFirst
   RsTxn_SRGodownMapping.Find "SRGodownMappingID = " & TxtMappingID & ""
End If

RsTxn_SRGodownMapping!SRGodownMappingDate = TxtMappingDate.Value
RsTxn_SRGodownMapping!ContractID = mContractID

If IsNull(RsTxn_SRGodownMapping!CreatedBy) = True Or RsTxn_SRGodownMapping!CreatedBy = "" Then
   RsTxn_SRGodownMapping!CreatedBy = Gen_UserLoginID
   RsTxn_SRGodownMapping!CreatedDate = Now
Else
   RsTxn_SRGodownMapping!UpdatedBy = Gen_UserLoginID
   RsTxn_SRGodownMapping!UpdatedDate = Now
End If
RsTxn_SRGodownMapping.Update

If MSHFlexGrid2.Rows > 2 Then
   tmp = "Delete from Txn_SRGodownMapDetails Where SRGodownMappingID = " & TxtMappingID & ""
   Call TmpTable
      
   Call TableTxn_SRGodownMapDetails(" Where SRGodownMappingID = " & TxtMappingID & "")
   x = 2
   For I = 1 To MSHFlexGrid2.Rows - 2
       If MSHFlexGrid2.TextMatrix(x, 0) = "" Then Exit For
'       If MSHFlexGrid2.TextMatrix(x, 7) <> "" Then
'          RsTxn_SRGodownMapDetails.Find "SRGodownMapDetailsID = " & MSHFlexGrid2.TextMatrix(x, 7)
'       Else
        RsTxn_SRGodownMapDetails.AddNew
'        RsTxn_SRGodownMapDetails!SRGodownMapDetailsID = GetMaxDetailsID()
        RsTxn_SRGodownMapDetails!G_UID = GetGUID
'       End If
       RsTxn_SRGodownMapDetails!SRGodownMappingID = TxtMappingID
       RsTxn_SRGodownMapDetails!GodownID = MSHFlexGrid2.TextMatrix(x, 6)
       RsTxn_SRGodownMapDetails!ReservationQty = Val(MSHFlexGrid2.TextMatrix(x, 5))
       If IsNull(RsTxn_SRGodownMapDetails!CreatedBy) = True Or RsTxn_SRGodownMapDetails!CreatedBy = "" Then
          RsTxn_SRGodownMapDetails!CreatedBy = Gen_UserLoginID
          RsTxn_SRGodownMapDetails!CreatedDate = Now
       Else
          RsTxn_SRGodownMapDetails!UpdatedBy = Gen_UserLoginID
          RsTxn_SRGodownMapDetails!UpdatedDate = Now
       End If
       x = x + 1
       RsTxn_SRGodownMapDetails.Update
   Next
End If

Call Indata

End Sub
Private Function GetMaxMappingID() As Double
Dim Retval As Double
tmp = "Select max(SRGodownMappingID) from Txn_SRGodownMapping "
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetMaxMappingID = Retval
End Function

Function GetMaxDetailsID() As Double
Dim Retval As Double
Retval = 0
tmp = "Select max(SRGodownMapDetailsID) from Txn_SRGodownMapDetails " 'where ClaimID = " & TxtClaimID & " And SM_prefix = '" & TxtPreFix & "'"
Call TmpTable
If TmpRs.RecordCount > 0 Then
   Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
End If
Retval = Retval + 1
TmpRs.Close
GetMaxDetailsID = Retval

End Function

Private Sub Indata()

TxtMappingID = RsTxn_SRGodownMapping!SRGodownMappingID
TxtMappingDate = RsTxn_SRGodownMapping!SRGodownMappingDate
TxtMappingDate.Enabled = False
CmbStorageDescription.Visible = False
TxtStorageDescription.Visible = True
CmdStorageReservation.Visible = False
mContractID = RsTxn_SRGodownMapping!ContractID
'TxtStorageDescription = GetContract(mContractID)
Call GetContractDetails(mContractID)

TxtCreated = IIf(IsNull(RsTxn_SRGodownMapping!CreatedBy), "", RsTxn_SRGodownMapping!CreatedBy) & " :- " & IIf(IsNull(RsTxn_SRGodownMapping!CreatedDate), "", RsTxn_SRGodownMapping!CreatedDate)
TxtUpdated = IIf(IsNull(RsTxn_SRGodownMapping!UpdatedBy), "", RsTxn_SRGodownMapping!UpdatedBy) & " :- " & IIf(IsNull(RsTxn_SRGodownMapping!UpdatedDate), "", RsTxn_SRGodownMapping!UpdatedDate)


Frame2.Enabled = False

Call Grid_Head
Call Indata_MappingDetails

Call GButtonLock(Me, True)

End Sub

Private Sub Indata_MappingDetails()

CmbCMPName.Text = ""
Call ClearGodownDetails

tmp = " Select MC.CMPName,MG.GodownNo,MU.UnitType,MG.BagsCapacity,MG.GodownCapacity, " & _
      " TSRGD.ReservationQty , MG.GodownID, MG.Address " & _
      " From Txn_SRGodownMapDetails TSRGD " & _
      " Inner Join Mst_Godown MG On TSRGD.GodownID = MG.GodownID " & _
      " Inner Join Mst_CMP MC On MG.CMPID = MC.CMPID " & _
      " Inner Join Mst_UnitType MU On MG.UnitTypeID = MU.UnitTypeID " & _
      " Where TSRGD.SRGodownMappingID = " & TxtMappingID & " "
      
Call TmpTable
If TmpRs.RecordCount > 0 Then
   x = 2
   For I = 1 To TmpRs.RecordCount
       MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
       For C = 0 To MSHFlexGrid2.Cols - 1
           MSHFlexGrid2.TextMatrix(x, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
       Next
       MSHFlexGrid2.TextMatrix(0, 3) = Val(MSHFlexGrid2.TextMatrix(0, 3)) + Val(MSHFlexGrid2.TextMatrix(x, 3))
       MSHFlexGrid2.TextMatrix(0, 4) = Val(MSHFlexGrid2.TextMatrix(0, 4)) + Val(MSHFlexGrid2.TextMatrix(x, 4))
       MSHFlexGrid2.TextMatrix(0, 5) = Val(MSHFlexGrid2.TextMatrix(0, 5)) + Val(MSHFlexGrid2.TextMatrix(x, 5))
       TmpRs.MoveNext
       x = x + 1
   Next
End If
End Sub
Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsTxn_SRGodownMapping.MoveNext
If RsTxn_SRGodownMapping.EOF Then
   RsTxn_SRGodownMapping.MoveLast
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
RsTxn_SRGodownMapping.MovePrevious
If RsTxn_SRGodownMapping.BOF Then
   RsTxn_SRGodownMapping.MoveFirst
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
RsTxn_SRGodownMapping.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsTxn_SRGodownMapping.MoveLast
Call Indata
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False
End Sub

Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   If RsTxn_SRGodownMapping.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   Call Indata
   Exit Sub
End If

Frame0.Height = Me.Height - 500
Frame0.Left = Frame3.Left
Frame0.Top = Frame3.Top
Frame0.Width = Me.Width - 400
Frame0.Visible = True
MSHFlexGrid1.Width = Frame0.Width - 300
MSHFlexGrid1.Height = Frame0.Height - 1600

''===========comment on 7 july 2012 for not clearing search grid=============
'MSHFlexGrid1.Clear
'Call Grid_Head1
'TxtSearchDescription = ""
'Label22.Caption = ""
''===========comment on 7 july 2012 for not clearing search grid=============

End Sub

Public Sub Grid_Head1()
MSHFlexGrid1.Clear
MSHFlexGrid1.Cols = 16 '7
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.FixedRows = 1

MSHFlexGrid1.TextMatrix(0, 0) = "Mapping ID"
MSHFlexGrid1.TextMatrix(0, 1) = "Mapping Date"
MSHFlexGrid1.TextMatrix(0, 2) = "Storage Description"
MSHFlexGrid1.TextMatrix(0, 3) = "Location"
MSHFlexGrid1.TextMatrix(0, 4) = "Client"
MSHFlexGrid1.TextMatrix(0, 5) = "Commodity"
MSHFlexGrid1.TextMatrix(0, 6) = "Storage Charges"
MSHFlexGrid1.TextMatrix(0, 7) = "Contract ID"
MSHFlexGrid1.TextMatrix(0, 8) = "Booked From"
MSHFlexGrid1.TextMatrix(0, 9) = "Extended Period"
MSHFlexGrid1.TextMatrix(0, 10) = "CMP Name"
MSHFlexGrid1.TextMatrix(0, 11) = "Godown No"
MSHFlexGrid1.TextMatrix(0, 12) = "Unit Type"
MSHFlexGrid1.TextMatrix(0, 13) = "Bags Capacity"
MSHFlexGrid1.TextMatrix(0, 14) = "Godown Capacity"
MSHFlexGrid1.TextMatrix(0, 15) = "Reservation Qty"


MSHFlexGrid1.ColWidth(0) = 1000
MSHFlexGrid1.ColWidth(1) = 1400
MSHFlexGrid1.ColWidth(2) = 3600
MSHFlexGrid1.ColWidth(3) = 2000
MSHFlexGrid1.ColWidth(4) = 2000
MSHFlexGrid1.ColWidth(5) = 2200
MSHFlexGrid1.ColWidth(6) = 1000

MSHFlexGrid1.RowHeight(0) = 400

End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Txn_SRGodownMapping.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Txn_SRGodownMapping.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = I + 1 ' + 6
    For C = 0 To MSHFlexGrid1.Cols - 1
        If C = 11 Then
           oWS.Cells(mRow, C + 1) = "'" & MSHFlexGrid1.TextMatrix(I, C)
        Else
           If InStr(MSHFlexGrid1.TextMatrix(I, C), "/") > 0 Then
              If IsDate(MSHFlexGrid1.TextMatrix(I, C)) Then
                 oWS.Cells(mRow, C + 1) = Format(MSHFlexGrid1.TextMatrix(I, C), "MM/dd/yyyy")
              Else
                 oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
              End If
           Else
              oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
           End If
        End If
    Next
Next
oWB.Save
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub

Private Sub TxtReservationQty_LostFocus()
If TxtReservationQty = "" Then Exit Sub

If IsNumeric(TxtReservationQty) = False Or Val(TxtReservationQty) < 0 Then
   MsgBox "Invalid numeric format!!!"
   TxtReservationQty.SetFocus
   Exit Sub
End If
End Sub

