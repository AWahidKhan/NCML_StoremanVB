VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmTxn_GSL_OpeningBalance 
   Caption         =   "FrmTxn_GSL_OpeningBalance"
   ClientHeight    =   5610
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   5610
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   3885
      Left            =   120
      TabIndex        =   45
      Top             =   5520
      Visible         =   0   'False
      Width           =   12345
      Begin VB.CommandButton CmdSearch 
         Caption         =   "Search"
         Height          =   350
         Left            =   11040
         TabIndex        =   54
         Top             =   960
         Width           =   1000
      End
      Begin VB.ComboBox CmbSCmp 
         Height          =   315
         Left            =   240
         Sorted          =   -1  'True
         TabIndex        =   53
         Top             =   600
         Width           =   4725
      End
      Begin VB.ComboBox CmbSGodown 
         Height          =   315
         Left            =   5085
         Sorted          =   -1  'True
         TabIndex        =   52
         Top             =   600
         Width           =   2565
      End
      Begin VB.ComboBox CmbSClient 
         Height          =   315
         Left            =   1200
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   1020
         Visible         =   0   'False
         Width           =   3765
      End
      Begin VB.ComboBox CmbSExchangeType 
         Height          =   315
         Left            =   11340
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   600
         Width           =   2565
      End
      Begin VB.ComboBox CmbSCommodity 
         Height          =   315
         Left            =   7680
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   600
         Width           =   3645
      End
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   350
         Left            =   12000
         TabIndex        =   48
         Top             =   960
         Width           =   1000
      End
      Begin VB.ComboBox CmbPreFix 
         Height          =   315
         Left            =   6045
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   960
         Width           =   1725
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1455
         Left            =   120
         TabIndex        =   47
         Top             =   1440
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   2566
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label2 
         Caption         =   "CMP"
         Height          =   300
         Left            =   1200
         TabIndex        =   60
         Top             =   240
         Width           =   1320
      End
      Begin VB.Label Label3 
         Caption         =   "Commodity"
         Height          =   300
         Left            =   8040
         TabIndex        =   59
         Top             =   240
         Width           =   1320
      End
      Begin VB.Label Label4 
         Caption         =   "Godown"
         Height          =   300
         Left            =   5160
         TabIndex        =   58
         Top             =   240
         Width           =   1320
      End
      Begin VB.Label Label5 
         Caption         =   "Exchange Type"
         Height          =   300
         Left            =   11400
         TabIndex        =   57
         Top             =   240
         Width           =   1320
      End
      Begin VB.Label Label6 
         Caption         =   "Client Name"
         Height          =   300
         Left            =   240
         TabIndex        =   56
         Top             =   1140
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.Label Label8 
         Caption         =   "Location ID"
         Height          =   300
         Left            =   5160
         TabIndex        =   55
         Top             =   1080
         Width           =   1320
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1200
      Left            =   120
      TabIndex        =   32
      Top             =   4320
      Width           =   11205
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   400
         Left            =   8865
         TabIndex        =   41
         Top             =   600
         Width           =   1980
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "Search"
         Height          =   400
         Left            =   8865
         TabIndex        =   40
         Top             =   195
         Width           =   1980
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   400
         Left            =   6765
         TabIndex        =   39
         Top             =   600
         Width           =   2100
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   400
         Left            =   6765
         TabIndex        =   38
         Top             =   195
         Width           =   2100
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   400
         Left            =   4665
         TabIndex        =   37
         Top             =   600
         Width           =   2100
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   400
         Left            =   4665
         TabIndex        =   36
         Top             =   195
         Width           =   2100
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   400
         Left            =   2445
         TabIndex        =   35
         Top             =   600
         Width           =   2220
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   400
         Left            =   2445
         TabIndex        =   12
         Top             =   195
         Width           =   2220
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   400
         Left            =   225
         TabIndex        =   34
         Top             =   600
         Width           =   2220
      End
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New"
         Height          =   400
         Left            =   225
         TabIndex        =   33
         Top             =   195
         Width           =   2220
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4305
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   11205
      Begin VB.ComboBox CmbClientDP 
         Height          =   315
         Left            =   2400
         Sorted          =   -1  'True
         TabIndex        =   65
         Top             =   2955
         Visible         =   0   'False
         Width           =   8250
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
         Left            =   7680
         Locked          =   -1  'True
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   2970
         Width           =   2970
      End
      Begin VB.TextBox TxtClient 
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
         Left            =   2400
         TabIndex        =   6
         Top             =   2100
         Width           =   2940
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
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   66
         Top             =   2955
         Width           =   2910
      End
      Begin VB.TextBox TxtG_Uid 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   360
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   3720
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.CommandButton CmdISIN 
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
         Left            =   5400
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   1680
         Width           =   375
      End
      Begin VB.TextBox TxtGSLID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   360
         Left            =   600
         Locked          =   -1  'True
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   270
         Width           =   840
      End
      Begin VB.TextBox TxtPreFix 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   360
         Left            =   7680
         Locked          =   -1  'True
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   1650
         Width           =   1440
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
         Height          =   315
         Left            =   5400
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   1200
         Width           =   375
      End
      Begin VB.CommandButton CmdClientId 
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
         Left            =   10680
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   2550
         Width           =   375
      End
      Begin VB.CommandButton CmdCommodityId 
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
         Left            =   10680
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   1185
         Width           =   375
      End
      Begin VB.ComboBox CmbCommodityID 
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
         Left            =   7680
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   1185
         Width           =   2895
      End
      Begin VB.TextBox TxtCommodityID 
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
         Left            =   7680
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   1185
         Width           =   2640
      End
      Begin VB.ComboBox CmbExchangeTypeID 
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
         TabIndex        =   3
         Top             =   1185
         Width           =   2925
      End
      Begin VB.TextBox TxtExchangeTypeID 
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
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1185
         Width           =   2880
      End
      Begin VB.ComboBox CmbGodownID 
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
         TabIndex        =   2
         Top             =   765
         Width           =   8175
      End
      Begin VB.CommandButton CmdGodownID 
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
         Left            =   10695
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   780
         Width           =   375
      End
      Begin VB.TextBox TxtGodownID 
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
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   765
         Width           =   7935
      End
      Begin VB.TextBox TxtDepositWeight 
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
         Left            =   7665
         MaxLength       =   10
         TabIndex        =   11
         Top             =   3855
         Width           =   3015
      End
      Begin VB.TextBox TxtDepositBags 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   2385
         MaxLength       =   10
         TabIndex        =   10
         Top             =   3855
         Width           =   2925
      End
      Begin VB.ComboBox CmbClientID 
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
         TabIndex        =   7
         Top             =   2550
         Width           =   8265
      End
      Begin VB.TextBox TxtISINID 
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
         TabIndex        =   5
         Top             =   1650
         Width           =   2925
      End
      Begin VB.TextBox TxtLotNo 
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
         Left            =   7665
         MaxLength       =   10
         TabIndex        =   9
         Top             =   3390
         Width           =   3015
      End
      Begin VB.TextBox TxtStackNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   2385
         MaxLength       =   10
         TabIndex        =   8
         Top             =   3405
         Width           =   2925
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
         Left            =   2400
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   315
         Width           =   8175
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
         Left            =   10680
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   315
         Width           =   375
      End
      Begin VB.TextBox TxtCMPName 
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
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   315
         Width           =   7920
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
         Height          =   360
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   2550
         Width           =   8220
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
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
         Left            =   600
         TabIndex        =   70
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label Label30 
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
         Height          =   495
         Left            =   6105
         TabIndex        =   69
         Top             =   2925
         Width           =   1575
      End
      Begin VB.Label lblNameDid 
         Alignment       =   1  'Right Justify
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
         Left            =   600
         TabIndex        =   68
         Top             =   2970
         Visible         =   0   'False
         Width           =   1740
      End
      Begin VB.Label LblGSLID 
         Caption         =   "Sr No"
         Height          =   180
         Left            =   120
         TabIndex        =   62
         Top             =   360
         Width           =   840
      End
      Begin VB.Label Label7 
         Caption         =   "Location ID"
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
         Left            =   6120
         TabIndex        =   44
         Top             =   1680
         Width           =   1080
      End
      Begin VB.Label LblLotNo 
         Caption         =   "Lot No"
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
         Left            =   6105
         TabIndex        =   31
         Top             =   3405
         Width           =   1200
      End
      Begin VB.Label LblQuantity 
         Caption         =   "Quantity"
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
         Left            =   6105
         TabIndex        =   30
         Top             =   3855
         Width           =   840
      End
      Begin VB.Label LblCommodityID 
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
         Height          =   300
         Left            =   6120
         TabIndex        =   29
         Top             =   1185
         Width           =   1320
      End
      Begin VB.Label LblGodownID 
         Alignment       =   1  'Right Justify
         Caption         =   "Godown"
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
         Left            =   495
         TabIndex        =   26
         Top             =   765
         Width           =   1800
      End
      Begin VB.Label LblNoOfBags 
         Alignment       =   1  'Right Justify
         Caption         =   "No Of Bags"
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
         Left            =   480
         TabIndex        =   25
         Top             =   3855
         Width           =   1800
      End
      Begin VB.Label LblISINID 
         Alignment       =   1  'Right Justify
         Caption         =   "ISIN No"
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
         Left            =   1080
         TabIndex        =   24
         Top             =   1650
         Width           =   1200
      End
      Begin VB.Label LblClientID 
         Alignment       =   1  'Right Justify
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
         Height          =   300
         Left            =   495
         TabIndex        =   23
         Top             =   2550
         Width           =   1800
      End
      Begin VB.Label LblExchangeTypeID 
         Alignment       =   1  'Right Justify
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
         Height          =   300
         Left            =   495
         TabIndex        =   22
         Top             =   1185
         Width           =   1800
      End
      Begin VB.Label LblStackNo 
         Alignment       =   1  'Right Justify
         Caption         =   "Stack No"
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
         Left            =   480
         TabIndex        =   21
         Top             =   3405
         Width           =   1800
      End
      Begin VB.Label LblCMPName 
         Alignment       =   1  'Right Justify
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
         Left            =   495
         TabIndex        =   20
         Top             =   315
         Width           =   1800
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "DP ID "
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
         TabIndex        =   71
         Top             =   3000
         Width           =   1800
      End
   End
End
Attribute VB_Name = "FrmTxn_GSL_OpeningBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg, msSM_prefix As Variant
Dim mExchangeTypeID As Double
Dim mCMPName, mCommodityID, mGodownID, mClientIDRow As Double
Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)

CmbCMPName.Visible = True
CmdCMPName.Visible = True
TxtCMPName.Visible = False
CmdISIN.Visible = True
CmbGodownID.Visible = True
CmdGodownID.Visible = True
TxtGodownID.Visible = False
TxtG_UID = ""
CmbExchangeTypeID.Visible = True
CmdExchangeTypeID.Visible = True
TxtExchangeTypeID.Visible = False

CmbClientId.Visible = True
CmdClientID.Visible = True
TxtClientID.Visible = False

CmbCommodityID.Visible = True
CmdCommodityID.Visible = True
TxtCommodityID.Visible = False

TxtISINID.Locked = False

TxtStackNo.Locked = False
TxtLotNo.Locked = False
TxtDepositBags.Locked = False
TxtDepositWeight.Locked = False

'CmbCMPName = ""
'CmbCommodityID = ""
'CmbGodownID = ""
'CmbClientId.Text = ""
'CmbExchangeTypeID.Text = ""
'TxtStackNo = ""
TxtLotNo = ""
TxtDepositWeight = ""
'TxtDepositBags = ""
'TxtISINID = ""
'TxtPreFix = ""
TxtClient = ""
TxtClientID = ""
TxtDPID = ""
TxtDPName = ""

Call CmbExchangeTypeID_LostFocus

If RsTxn_GSL_OpeningBalance.RecordCount > 0 Then
   CmbCMPName.SetFocus
End If
End Sub

Private Sub CmbClientDP_LostFocus()

Dim Str_() As String
Str_ = Split(CmbClientDP.Text, "~")
If UBound(Str_) <> 1 Then
   MsgBox "Invalid selection!!! "
   CmbClientDP.SetFocus
   Exit Sub
End If
TmpRs1.MoveFirst


TmpRs1.MoveFirst
TmpRs1.Find "DepositoryID  = '" & Str_(0) & "'"
If TmpRs1.EOF Then
   MsgBox "Invalid selection!!! "
   CmbClientDP.SetFocus
   Exit Sub
End If

TxtClientID.Visible = True
TxtClientID = TmpRs1!ClientName
TxtDPID = TmpRs1!DepositoryID
TxtDPName = TmpRs1!DepositoryName
CmbClientDP.Visible = False
mClientIDRow = TmpRs1!ClientIDRow

End Sub

Private Sub CmdClientID_Click()
FrmMst_Client.Show
End Sub

Private Sub CmdCMPName_Click()
FrmMst_CMP.Show
End Sub

Private Sub CmdCommodityId_Click()
FrmMst_Commodity.Show
End Sub

Private Sub CmdExchangeTypeID_Click()
FrmMst_ExchangeType.Show
End Sub

Private Sub CmdGodownID_Click()
FrmMst_CMP.Show
End Sub



Private Sub CmdISIN_Click()
FrmMst_ISIN.Show
End Sub

Private Sub EditCmd_Click()

Call TableMst_GSL(" where sm_prefix = '" & TxtPreFix & "'")
If RsMst_GSL.RecordCount > 0 Then
   MsgBox "Import process already executed for selected location,you can't edit this Record!!!"
   Exit Sub
End If

Edit_Flg = "E"
Call GButtonLock(Me, False)

CmbCMPName.Visible = True
CmdCMPName.Visible = True
TxtCMPName.Visible = False

CmbGodownID.Visible = True
CmdGodownID.Visible = True
TxtGodownID.Visible = False

CmbExchangeTypeID.Visible = True
CmdExchangeTypeID.Visible = True
TxtExchangeTypeID.Visible = False

If mExchangeTypeID = 1 Then
   CmbClientId.Visible = True
   TxtClientID.Visible = False
Else
   CmbClientId.Visible = False
   TxtClientID.Visible = True
End If
CmdClientID.Visible = True

CmbCommodityID.Visible = True
TxtCommodityID.Visible = False
CmdCommodityID.Visible = True
TxtISINID.Locked = False

TxtStackNo.Locked = False
TxtLotNo.Locked = False
TxtDepositBags.Locked = False
TxtDepositWeight.Locked = False

CmbCMPName.Text = TxtCMPName
CmbCommodityID.Text = TxtCommodityID
CmbGodownID.Text = TxtGodownID
CmbClientId.Text = TxtClientID
CmbExchangeTypeID.Text = TxtExchangeTypeID

If RsTxn_GSL_OpeningBalance.RecordCount > 0 Then
   CmbCMPName.SetFocus
End If
End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Call GButtonLockAD(Me)
   MsgBox "Access denied!!!"
   Exit Sub
End If


Call TableTxn_GSL_OpeningBalance(" Order By CMPID,GodownID,StackNo,LotNo ")

If RsTxn_GSL_OpeningBalance.RecordCount = 0 Then
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
tmp = CmbCMPName.Text
Call TableMst_CMP
CmbCMPName.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "No record found!!! "
   Exit Sub
End If
For I = 1 To RsMst_CMP.RecordCount
    CmbCMPName.AddItem RsMst_CMP!CMPName
    RsMst_CMP.MoveNext
Next
CmbCMPName.Text = tmp
End Sub
'----
Private Sub CmbCMPName_LostFocus()
If CmbCMPName.Text = "" Then
   CmbGodownID.Text = ""
   TxtStackNo = ""
   TxtLotNo = ""
   TxtISINID = ""
   mCMPName = Null
   Exit Sub
End If
RsMst_CMP.MoveFirst


RsMst_CMP.Find "CMPName = '" & CmbCMPName.Text & "'"
If RsMst_CMP.EOF Then
   MsgBox "Invalid selection!!! "
   CmbCMPName.SetFocus
   Exit Sub
End If
If mCMPName <> RsMst_CMP!CMPID Then
   CmbGodownID.Text = ""
   TxtStackNo = ""
   TxtLotNo = ""
   TxtISINID = ""
End If

mCMPName = RsMst_CMP!CMPID

TxtPreFix = GetPreFix(RsMst_CMP!CMPID, "P")
End Sub
Private Sub CmbGodownID_GotFocus()
If CmbCMPName.Text = "" Then
   MsgBox "Select CMP first!!! "
   CmbCMPName.SetFocus
   Exit Sub
End If
tmp = CmbGodownID.Text
Call TableMst_Godown("Where CMPID = " & mCMPName)
CmbGodownID.Clear
If RsMst_Godown.RecordCount = 0 Then
   MsgBox "No record found!!! "
   Exit Sub
End If
For I = 1 To RsMst_Godown.RecordCount
    CmbGodownID.AddItem RsMst_Godown!GodownNo
    RsMst_Godown.MoveNext
Next
CmbGodownID.Text = tmp
End Sub
'----
Private Sub CmbGodownID_LostFocus()
If CmbGodownID.Text = "" Then
   TxtStackNo = ""
   TxtLotNo = ""
   mGodownID = Null
   Exit Sub
End If
RsMst_Godown.MoveFirst
If mGodownID <> RsMst_Godown!CMPID Then
   TxtStackNo = ""
   TxtLotNo = ""
End If
RsMst_Godown.Find "GodownNo = '" & CmbGodownID.Text & "'"
If RsMst_Godown.EOF Then
   MsgBox "Invalid selection!!! "
   CmbGodownID.SetFocus
   Exit Sub
End If
mGodownID = RsMst_Godown!GodownID
End Sub
Private Sub CmbExchangeTypeID_GotFocus()
tmp = CmbExchangeTypeID.Text
Call TableMst_ExchangeType
CmbExchangeTypeID.Clear
If RsMst_ExchangeType.RecordCount = 0 Then
   MsgBox "No record found!!! "
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
   CmbCommodityID.Text = ""
   Exit Sub
End If
RsMst_ExchangeType.MoveFirst

RsMst_ExchangeType.Find "ExchangeType = '" & CmbExchangeTypeID.Text & "'"
If RsMst_ExchangeType.EOF Then
   MsgBox "Invalid selection!!! "
   CmbExchangeTypeID.SetFocus
   Exit Sub
End If
If mExchangeTypeID <> RsMst_ExchangeType!ExchangeTypeID Then
   'CmbCommodityID.Text = ""
   TxtISINID = ""
End If
mExchangeTypeID = RsMst_ExchangeType!ExchangeTypeID

TxtClient = ""
TxtClientID = ""
TxtDPID = ""
TxtDPName = ""
CmbClientId = ""
LblISINID.Caption = ""
If mExchangeTypeID = 2 Then
   LblISINID.Caption = "ISIN No"
   TxtISINID.Enabled = True
   TxtClient.Enabled = True
   TxtClientID.Visible = True
   CmbClientId.Visible = False
   CmdISIN.Enabled = True
ElseIf mExchangeTypeID = 3 Then
   LblISINID.Caption = "CDC No"
   TxtISINID.Enabled = True
   TxtClient.Enabled = True
   TxtClientID.Visible = True
   CmbClientId.Visible = False
   CmdISIN.Enabled = True
Else
   
   TxtISINID.Enabled = False
   TxtClient.Enabled = False
   TxtClientID.Visible = True
   CmbClientId.Visible = True
   CmdISIN.Enabled = False
End If


End Sub
Private Sub CmbCommodityID_GotFocus()
If CmbExchangeTypeID.Text <> "" Then
   tmp = CmbCommodityID.Text
   Call TableMst_Commodity
   CmbCommodityID.Clear
   If RsMst_Commodity.RecordCount = 0 Then
      MsgBox "No record found!!! "
      Exit Sub
   End If
   For I = 1 To RsMst_Commodity.RecordCount
       CmbCommodityID.AddItem RsMst_Commodity!Commodity
       RsMst_Commodity.MoveNext
   Next
   CmbCommodityID.Text = tmp
 Else
   MsgBox "Please select Exchange Mode!!!"
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
   MsgBox "Invalid selection!!! "
   CmbCommodityID.SetFocus
   Exit Sub
End If

If mCommodityID <> RsMst_Commodity!CommodityID Then
   TxtISINID = ""
End If
mCommodityID = RsMst_Commodity!CommodityID

Call TableMst_CommodityDetail("Where CommodityId = " & mCommodityID & " and ExchangeTypeId = " & mExchangeTypeID)
 If RsMst_CommodityDetail.RecordCount > 0 Then
 Else
    MsgBox "No record found for Exchange" & CmbExchangeTypeID.Text & " - " & CmbCommodityID.Text
    CmbCommodityID.SetFocus
    Exit Sub
 End If
'SSTab1.Tab = 0
End Sub

Private Sub CmbClientId_GotFocus()
tmp = CmbClientId.Text
Call TableMst_Client("where ExchangeTypeId = 1")
CmbClientId.Clear
If RsMst_Client.RecordCount = 0 Then
   MsgBox "No record found!!! "
   Exit Sub
End If
For I = 1 To RsMst_Client.RecordCount
    CmbClientId.AddItem RsMst_Client!ClientName
    RsMst_Client.MoveNext
Next
CmbClientId.Text = tmp
End Sub

Private Sub CmbClientId_LostFocus()
If CmbClientId.Text = "" Then
   Exit Sub
End If
RsMst_Client.MoveFirst
RsMst_Client.Find "ClientName = '" & CmbClientId.Text & "'"
If RsMst_Client.EOF Then
   MsgBox "Invalid selection!!! "
   CmbClientId.SetFocus
   Exit Sub
End If
mClientIDRow = RsMst_Client!ClientIDRow
'Dim ans As Variant
'ans = GetClientNameByRow(RsMst_Client!ClientIDRow)
'TxtClientID = ans(0)
'TxtClientName = ans(1)
'TxtDPID = ans(2)
'TxtDPName = ans(3)

End Sub

Private Sub SaveCmd_Click()
Call TableMst_GSL(" where sm_prefix = '" & TxtPreFix & "'")
If RsMst_GSL.RecordCount > 0 Then
   MsgBox "Import process already executed for selected location. You can't delete this record!!!"
   Exit Sub
End If

If CmbCMPName.Text = "" Then
   MsgBox "Select CMP Name"
   CmbCMPName.SetFocus
   Exit Sub
End If
If CmbGodownID.Text = "" Then
   MsgBox "Select Godown "
   CmbGodownID.SetFocus
   Exit Sub
End If
If CmbExchangeTypeID.Text = "" Then
   MsgBox "Select Exchange "
   CmbExchangeTypeID.SetFocus
   Exit Sub
End If

If CmbCommodityID.Text = "" Then
   MsgBox "Select Commodity "
   CmbCommodityID.SetFocus
   Exit Sub
End If



If mExchangeTypeID <> 1 Then
   
   If TxtISINID = "" And TxtClient = "" Then
      If mExchangeTypeID = 2 Then
         MsgBox "ISIN No/Client ID can not be blank!!!"
      ElseIf mExchangeTypeID = 3 Then
         MsgBox "CDC No/Client ID can not be blank!!!"
      End If
      TxtISINID.SetFocus
      Exit Sub
   End If
Else
   If CmbClientId.Text = "" Then
      MsgBox "Select Client Name"
      CmbClientId.SetFocus
      Exit Sub
   End If
End If

If TxtStackNo.Text = "" Then
   MsgBox "Blank Stack No not allowed!!! "
   TxtStackNo.SetFocus
   Exit Sub
End If

If TxtLotNo.Text = "" Then
   MsgBox "Blank Lot No not allowed!!! "
   TxtLotNo.SetFocus
   Exit Sub
End If

If TxtDepositBags.Text = "" Then
   MsgBox "No Of Bags cannot be empty!!! "
   TxtDepositBags.SetFocus
   Exit Sub
End If

If TxtDepositWeight.Text = "" Then
   MsgBox "Quantity cannot be empty!!! "
   TxtDepositWeight.SetFocus
   Exit Sub
End If

If Edit_Flg = "A" Then
   If RsTxn_GSL_OpeningBalance.RecordCount > 0 Then
      RsTxn_GSL_OpeningBalance.MoveFirst
      tmp = "sm_prefix = '" & TxtPreFix & "' And CMPID = " & mCMPName & " And GodownID = " & mGodownID & " And StackNo = '" & TxtStackNo.Text & "' And LotNo = '" & TxtLotNo & "'"
      RsTxn_GSL_OpeningBalance.Filter = tmp
      If RsTxn_GSL_OpeningBalance.RecordCount > 0 Then
        MsgBox "Duplicate GSL not allowed!!! "
         TxtLotNo.SetFocus
         RsTxn_GSL_OpeningBalance.Filter = ""
         Exit Sub
      End If
   End If
   RsTxn_GSL_OpeningBalance.AddNew
   RsTxn_GSL_OpeningBalance!GSLID = GetGSLID
   TxtGSLID.Text = RsTxn_GSL_OpeningBalance!GSLID
   RsTxn_GSL_OpeningBalance!G_UID = GetGUID
   TxtG_UID = RsTxn_GSL_OpeningBalance!G_UID
Else
   RsTxn_GSL_OpeningBalance.MoveFirst
   tmp = "sm_prefix = '" & TxtPreFix & "' And CMPID = " & mCMPName & " And GodownID = " & mGodownID & " And StackNo = '" & TxtStackNo.Text & "' And LotNo = '" & TxtLotNo & "'"
   RsTxn_GSL_OpeningBalance.Filter = tmp
   If Not RsTxn_GSL_OpeningBalance.EOF Then
      If RsTxn_GSL_OpeningBalance!GSLID <> TxtGSLID.Text Then
         MsgBox "Duplicate GSL not allowed!!! "
         TxtStackNo.SetFocus
         Exit Sub
      End If
   End If
   'tmp = "sm_prefix = '" & TxtPreFix & "'"
   'RsTxn_GSL_OpeningBalance.Filter = tmp
   RsTxn_GSL_OpeningBalance.MoveFirst
   RsTxn_GSL_OpeningBalance.Find " G_UID= '" & TxtG_UID & "'"
End If

RsTxn_GSL_OpeningBalance!GSLID = TxtGSLID.Text
RsTxn_GSL_OpeningBalance!SM_Prefix = TxtPreFix
RsTxn_GSL_OpeningBalance!CMPID = mCMPName
RsTxn_GSL_OpeningBalance!GodownID = mGodownID

RsTxn_GSL_OpeningBalance!ExchangeTypeID = mExchangeTypeID
RsTxn_GSL_OpeningBalance!CommodityID = mCommodityID

If TxtISINID <> "" Then
   RsTxn_GSL_OpeningBalance!ISINID = TxtISINID.Text
   RsTxn_GSL_OpeningBalance!ClientIDRow = 8
Else
   RsTxn_GSL_OpeningBalance!ISINID = ""
   RsTxn_GSL_OpeningBalance!ClientIDRow = mClientIDRow
End If
RsTxn_GSL_OpeningBalance!StackNo = TxtStackNo.Text
RsTxn_GSL_OpeningBalance!LotNo = TxtLotNo.Text
RsTxn_GSL_OpeningBalance!DepositBags = Val(TxtDepositBags.Text)
RsTxn_GSL_OpeningBalance!DepositWeight = Val(TxtDepositWeight.Text)
RsTxn_GSL_OpeningBalance.Update
RsTxn_GSL_OpeningBalance.Filter = ""
RsTxn_GSL_OpeningBalance.MoveFirst
RsTxn_GSL_OpeningBalance.Find "g_uid = '" & TxtG_UID & "'"
If RsTxn_GSL_OpeningBalance.EOF Then
   RsTxn_GSL_OpeningBalance.MoveFirst
End If
Call Indata

End Sub

Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   If RsTxn_GSL_OpeningBalance.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   'Call FirstCmd_Click
   Call Indata
   Exit Sub
End If
Frame0.Height = Me.Height - 900
Frame0.Left = Frame1.Left - 100
Frame0.Top = Frame1.Top - 100
Frame0.Width = Me.Width - 150
Frame0.Visible = True
MSHFlexGrid1.Clear
MSHFlexGrid1.Cols = 3
MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 1500
RsTxn_GSL_OpeningBalance.MoveFirst
Call Grid_Head

CmbSCmp.Text = ""
CmbSGodown.Text = ""
CmbSCommodity.Text = ""
CmbSExchangeType.Text = ""
CmbSClient.Text = ""

MSHFlexGrid1.ColWidth(0) = 1000
Dim ans As Variant
End Sub

Private Sub TxtClient_LostFocus()

If SaveCmd.Enabled = False Then Exit Sub
CmbClientDP.Visible = False
If TxtClient = "" Then
   mClientIDRow = 0
   Exit Sub
End If
 'If CmbExchangeTypeID.Text = "" Then Exit Sub
 
Dim ans As Variant
mExchangeTypeID = GetExchangeID(CmbExchangeTypeID.Text)
ans = GetClientNameDtl(TxtClient, mExchangeTypeID)

If ans(0, 0) = False Then
   MsgBox "Not found!!!"
   TxtClientID.SetFocus
   Exit Sub
End If
If ans(3, 0) > 1 Then
   CmbClientDP.Visible = True
End If
CmbClientDP.Text = ""
CmbClientDP.Clear
tmp = "SELECT Mst_Depository.DepositoryID, Mst_Depository.DepositoryName, Mst_Client.ClientID, " & _
      "Mst_Client.ClientName,  Mst_Client.ClientIDRow " & _
      "FROM Mst_Depository INNER JOIN Mst_Client ON Mst_Depository.DepositoryID = Mst_Client.DepositoryID " & _
      "WHERE Mst_Depository.ExchangeTypeID = " & mExchangeTypeID & _
      "  And  (((Mst_Client.ExchangeTypeID)=" & mExchangeTypeID & ") AND ((Mst_Client.ClientID)='" & TxtClient & "'));"
            
'      "WHERE (((Mst_Depository.ExchangeTypeID)=" & mExchangeTypeID & ") AND ((Mst_Client.ClientID)='" & TxtClientID & "'));"
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
Else
   If CmbClientDP.Visible = True Then
      CmbClientDP.SetFocus
   End If
End If
'If SaveCmd.Enabled = True Then
  'Call GetGridData(mClientIDRow, mExchangeTypeID, mCMPID, mCommodityID)
'End If


End Sub

Private Sub TxtDepositBags_LostFocus()
If TxtDepositBags = "" Then Exit Sub
If IsNumeric(TxtDepositBags) = False Or Val(TxtDepositBags) < 0 Then
   MsgBox "Invalid numeric format!!!"
   TxtDepositBags.SetFocus
   Exit Sub
End If

TxtDepositBags = Round(Val(TxtDepositBags), "0")

End Sub

Private Sub TxtDepositWeight_LostFocus()
If TxtDepositWeight = "" Then Exit Sub
If IsNumeric(TxtDepositWeight) = False Or Val(TxtDepositWeight) < 0 Then
   MsgBox "Invalid numeric format!!!"
   TxtDepositWeight.SetFocus
   Exit Sub
End If
End Sub

Private Sub TxtISINID_LostFocus()
If SaveCmd.Enabled = False Then Exit Sub
If TxtISINID.Text = "" Then
   TxtClient.Enabled = True
   TxtClient.SetFocus
   Exit Sub
End If

Call TableMst_ISIN("where ISINID = '" & TxtISINID.Text & "'")
If RsMst_ISIN.RecordCount = 0 Then
   MsgBox "There is no such ISIN!!!"
   TxtISINID.SetFocus
   Exit Sub
End If
If mCommodityID <> IIf(IsNull(RsMst_ISIN!CommodityID), "", RsMst_ISIN!CommodityID) Then
   MsgBox "Commodity mismatch!!! "
   TxtISINID.SetFocus
   Exit Sub
End If

If mCMPName <> IIf(IsNull(RsMst_ISIN!CMPID), "", RsMst_ISIN!CMPID) Then
   MsgBox "Commodity mismatch!!! "
   TxtISINID.SetFocus
   Exit Sub
End If
TxtClient = ""
TxtClientID = ""
TxtDPID = ""
TxtDPName = ""

TxtClient.Enabled = False
  
End Sub

Private Sub TxtLotNo_LostFocus()
TxtLotNo = Trim(TxtLotNo)

End Sub

Private Sub TxtStackNo_LostFocus()
TxtStackNo = Trim(TxtStackNo)
End Sub

Public Sub Indata()
CmbCMPName.Visible = False
CmdCMPName.Visible = False
TxtCMPName.Visible = True

CmbGodownID.Visible = False
CmdGodownID.Visible = False
TxtGodownID.Visible = True
CmdISIN.Visible = False

CmbExchangeTypeID.Visible = False
CmdExchangeTypeID.Visible = False
TxtExchangeTypeID.Visible = True

CmbClientId.Visible = False
CmdClientID.Visible = False
TxtClientID.Visible = True

CmbCommodityID.Visible = False
CmdCommodityID.Visible = False
TxtCommodityID.Visible = True

TxtISINID.Locked = True

TxtStackNo.Locked = True
TxtLotNo.Locked = True
TxtDepositBags.Locked = True
TxtDepositWeight.Locked = True

'RsTxn_GSL_OpeningBalance
TxtGSLID = IIf(IsNull(RsTxn_GSL_OpeningBalance!GSLID), "", RsTxn_GSL_OpeningBalance!GSLID)
TxtPreFix = IIf(IsNull(RsTxn_GSL_OpeningBalance!SM_Prefix), "", RsTxn_GSL_OpeningBalance!SM_Prefix)

TxtCommodityID = ""
TxtExchangeTypeID = ""
TxtISINID = ""
'TxtClientID = ""
TxtClient = ""
TxtClientID = ""
TxtDPID = ""
TxtDPName = ""

mCMPID = IIf(IsNull(RsTxn_GSL_OpeningBalance!CMPID), "", RsTxn_GSL_OpeningBalance!CMPID)
mCMPName = IIf(IsNull(RsTxn_GSL_OpeningBalance!CMPID), "", RsTxn_GSL_OpeningBalance!CMPID)
TxtCMPName = GetCMPName(RsTxn_GSL_OpeningBalance!CMPID) 'IIf(IsNull(RsTxn_GSL_OpeningBalance!CMPID), "", RsTxn_GSL_OpeningBalance!CMPID)
If RsTxn_GSL_OpeningBalance!CommodityID <> "" Then
   mCommodityID = IIf(IsNull(RsTxn_GSL_OpeningBalance!CommodityID), "", RsTxn_GSL_OpeningBalance!CommodityID)
   TxtCommodityID = GetCommodityName(IIf(IsNull(RsTxn_GSL_OpeningBalance!CommodityID), "0", RsTxn_GSL_OpeningBalance!CommodityID))
End If
mGodownID = IIf(IsNull(RsTxn_GSL_OpeningBalance!GodownID), "", RsTxn_GSL_OpeningBalance!GodownID)
TxtGodownID = GetGodownName(RsTxn_GSL_OpeningBalance!GodownID, "N") 'IIf(IsNull(RsTxn_GSL_OpeningBalance!GodownNo), "", RsTxn_GSL_OpeningBalance!GodownNo)

TxtStackNo = IIf(IsNull(RsTxn_GSL_OpeningBalance!StackNo), "", RsTxn_GSL_OpeningBalance!StackNo)
TxtLotNo = IIf(IsNull(RsTxn_GSL_OpeningBalance!LotNo), "", RsTxn_GSL_OpeningBalance!LotNo)
TxtDepositWeight = IIf(IsNull(RsTxn_GSL_OpeningBalance!DepositWeight), "", RsTxn_GSL_OpeningBalance!DepositWeight)
TxtDepositBags = IIf(IsNull(RsTxn_GSL_OpeningBalance!DepositBags), "", RsTxn_GSL_OpeningBalance!DepositBags)

mExchangeTypeID = IIf(IsNull(RsTxn_GSL_OpeningBalance!ExchangeTypeID), "", RsTxn_GSL_OpeningBalance!ExchangeTypeID)
If mExchangeTypeID <> 0 Then
   TxtExchangeTypeID = GetExchangeName(RsTxn_GSL_OpeningBalance!ExchangeTypeID) 'IIf(IsNull(RsTxn_GSL_OpeningBalance!ExchangeType), "", RsTxn_GSL_OpeningBalance!ExchangeType)
End If

mClientIDRow = 0
If Not IsNull(RsTxn_GSL_OpeningBalance!ClientIDRow) Then
   mClientIDRow = IIf(IsNull(RsTxn_GSL_OpeningBalance!ClientIDRow), 0, RsTxn_GSL_OpeningBalance!ClientIDRow)
End If
If mClientIDRow <> 0 Then
   
   If mExchangeTypeID = 1 Then
      TxtClientID = GetClientName(RsTxn_GSL_OpeningBalance!ClientIDRow) '(IsNull(RsTxn_GSL_OpeningBalance!ClientName), "", RsTxn_GSL_OpeningBalance!ClientName)
   Else
      ans = GetClientNameByRow(mClientIDRow)
      TxtClient = ans(0)
      TxtClientID = ans(1)
      TxtDPID = ans(2)
      TxtDPName = ans(3)
   End If
End If

mISINID = IIf(IsNull(RsTxn_GSL_OpeningBalance!ISINID), "0", RsTxn_GSL_OpeningBalance!ISINID)
TxtISINID = IIf(IsNull(RsTxn_GSL_OpeningBalance!ISINID), "0", RsTxn_GSL_OpeningBalance!ISINID)

TxtG_UID = IIf(IsNull(RsTxn_GSL_OpeningBalance!G_UID), "0", RsTxn_GSL_OpeningBalance!G_UID)

If RsTxn_GSL_OpeningBalance.RecordCount = 1 Then
   Call GButtonLock_1(Me, True)
Else
   Call GButtonLock(Me, True)
End If
End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsTxn_GSL_OpeningBalance.MoveNext
If RsTxn_GSL_OpeningBalance.EOF Then
   RsTxn_GSL_OpeningBalance.MoveLast
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
RsTxn_GSL_OpeningBalance.MovePrevious
If RsTxn_GSL_OpeningBalance.BOF Then
   RsTxn_GSL_OpeningBalance.MoveFirst
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
RsTxn_GSL_OpeningBalance.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsTxn_GSL_OpeningBalance.MoveLast
Call Indata
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False
End Sub
Private Sub Grid_Head()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 13

MSHFlexGrid1.TextMatrix(0, 0) = "GSL ID"
MSHFlexGrid1.TextMatrix(0, 1) = "Location ID"
MSHFlexGrid1.TextMatrix(0, 2) = "CMP Name"
MSHFlexGrid1.TextMatrix(0, 3) = "Commodity"
MSHFlexGrid1.TextMatrix(0, 4) = "Godown No"
MSHFlexGrid1.TextMatrix(0, 5) = "Stack No"
MSHFlexGrid1.TextMatrix(0, 6) = "Lot No"
MSHFlexGrid1.TextMatrix(0, 7) = "Deposited Bags"
MSHFlexGrid1.TextMatrix(0, 8) = "Deposited Weight"
MSHFlexGrid1.TextMatrix(0, 9) = "Exchange Type"
MSHFlexGrid1.TextMatrix(0, 10) = "Client Name"
MSHFlexGrid1.TextMatrix(0, 11) = "ISIN ID / CDC No"
MSHFlexGrid1.TextMatrix(0, 12) = "GUID"

MSHFlexGrid1.RowHeight(0) = 500
MSHFlexGrid1.ColWidth(9) = 1000
MSHFlexGrid1.ColWidth(12) = 0
End Sub

Private Sub MSHFlexGrid1_Click()
Dim I As Variant
I = MSHFlexGrid1.RowSel
If I <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(I, 0) = "" Then Exit Sub
RsTxn_GSL_OpeningBalance.Filter = ""
RsTxn_GSL_OpeningBalance.MoveFirst

'tmp = "G_UID = '" & MSHFlexGrid1.TextMatrix(i, 12) & "'"
'RsTxn_GSL_OpeningBalance.Find "GSLID = " & (MSHFlexGrid1.TextMatrix(i, 0))
RsTxn_GSL_OpeningBalance.Find ("G_UID = '" & MSHFlexGrid1.TextMatrix(I, 12) & "'")
If Not RsTxn_GSL_OpeningBalance.EOF Then
   Call Indata
   'RsTxn_GSL_OpeningBalance.Filter = ""
   Frame0.Visible = False
End If
End Sub

Private Sub CmbPreFix_GotFocus()
tmp = CmbPreFix.Text
Call TableMst_StoreManLocations("")

CmbPreFix.Clear
If RsMst_StoreManLocations.RecordCount = 0 Then
   MsgBox "No Bank found"
   CmbPreFix.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_StoreManLocations.RecordCount
    CmbPreFix.AddItem RsMst_StoreManLocations!SM_Prefix
    RsMst_StoreManLocations.MoveNext
Next
CmbPreFix.Text = tmp
End Sub

Private Sub CmbPreFix_LostFocus()
If CmbPreFix.Text = "" Then
   Exit Sub
End If
RsMst_StoreManLocations.MoveFirst
RsMst_StoreManLocations.Find "SM_Prefix = '" & CmbPreFix.Text & "'"

If RsMst_StoreManLocations.EOF Then
   MsgBox "Invalid Location ID selection!!!"
   CmbPreFix.SetFocus
   Exit Sub
End If
msSM_prefix = RsMst_StoreManLocations!SM_Prefix

End Sub

Private Sub CmbSCmp_GotFocus()
tmp = CmbSCmp.Text
Call TableMst_CMP
CmbSCmp.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "No record found!!! "
   Exit Sub
End If
For I = 1 To RsMst_CMP.RecordCount
    CmbSCmp.AddItem RsMst_CMP!CMPName
    RsMst_CMP.MoveNext
Next
CmbSCmp.Text = tmp
End Sub
Private Sub CmbSCmp_LostFocus()
If CmbSCmp.Text = "" Then
   Exit Sub
End If
RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPName = '" & CmbSCmp.Text & "'"
If RsMst_CMP.EOF Then
   MsgBox "Invalid selection "
   CmbSCmp.SetFocus
   Exit Sub
End If
mCMPName = RsMst_CMP!CMPID
End Sub

Private Sub CmbSCommodity_GotFocus()
tmp = CmbSCommodity.Text
Call TableMst_Commodity
CmbSCommodity.Clear
If RsMst_Commodity.RecordCount = 0 Then
   MsgBox "No record found!!! "
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
   MsgBox "Invalid selection!!! "
   CmbSCommodity.SetFocus
   Exit Sub
End If

mCommodityID = RsMst_Commodity!CommodityID

End Sub

Private Sub CmbSExchangeType_GotFocus()
tmp = CmbSExchangeType.Text
Call TableMst_ExchangeType
CmbSExchangeType.Clear
If RsMst_ExchangeType.RecordCount = 0 Then
   MsgBox "No record found!!! "
   Exit Sub
End If
For I = 1 To RsMst_ExchangeType.RecordCount
    CmbSExchangeType.AddItem RsMst_ExchangeType!ExchangeType
    RsMst_ExchangeType.MoveNext
Next
CmbSExchangeType.Text = tmp
End Sub

Private Sub CmbSExchangeType_LostFocus()
If CmbSExchangeType.Text = "" Then
   Exit Sub
End If
RsMst_ExchangeType.MoveFirst
RsMst_ExchangeType.Find "ExchangeType = '" & CmbSExchangeType.Text & "'"
If RsMst_ExchangeType.EOF Then
   MsgBox "Invalid selection!!! "
   CmbSExchangeType.SetFocus
   Exit Sub
End If
mExchangeTypeID = RsMst_ExchangeType!ExchangeTypeID

End Sub

Private Sub CmbSGodown_GotFocus()
If CmbSCmp.Text = "" Then
   MsgBox "Select CMP first!!! "
   CmbSCmp.SetFocus
   Exit Sub
End If
tmp = CmbSGodown.Text
Call TableMst_Godown("Where CMPID = " & mCMPName)
CmbSGodown.Clear
If RsMst_Godown.RecordCount = 0 Then
   MsgBox "No record found!!! "
   Exit Sub
End If
For I = 1 To RsMst_Godown.RecordCount
    CmbSGodown.AddItem RsMst_Godown!GodownNo
    RsMst_Godown.MoveNext
Next
CmbGodownID.Text = tmp
End Sub

Private Sub CmbSGodown_LostFocus()
If CmbSGodown.Text = "" Then
   Exit Sub
End If
RsMst_Godown.MoveFirst
RsMst_Godown.Find "GodownNo = '" & CmbSGodown.Text & "'"
If RsMst_Godown.EOF Then
   MsgBox "Invalid selection!!! "
   CmbSGodown.SetFocus
   Exit Sub
End If
mGodownID = RsMst_Godown!GodownID
End Sub

Private Sub CmdSearch_Click()
Dim Wc As String
Call Grid_Head
If CmbSCmp.Text <> "" Then
   Wc = GenWc(Wc, mCMPName, "Mst_CMP.CMPID", "N")
End If

If CmbSGodown.Text <> "" Then
   Wc = GenWc(Wc, mGodownID, "Txn_GSL_OpeningBalance.GodownID", "N")
End If
If CmbSCommodity.Text <> "" Then
   Wc = GenWc(Wc, mCommodityID, "Mst_Commodity.CommodityID", "N")
End If
If CmbSExchangeType.Text <> "" Then
   Wc = GenWc(Wc, mExchangeTypeID, "Mst_ExchangeType.ExchangeTypeID", "N")
End If
If CmbPreFix.Text <> "" Then
   Wc = GenWc(Wc, msSM_prefix, "Txn_GSL_OpeningBalance.SM_prefix", "S")
End If

'tmp = "Select * from Txn_GSL_OpeningBalance " & Wc
tmp = "SELECT  Txn_GSL_OpeningBalance.GSLID,Mst_ExchangeType.ExchangeType , Mst_Commodity.Commodity, "
tmp = tmp & " Mst_Client.ClientName, Mst_CMP.CMPName,Mst_Godown.GodownNo , Txn_GSL_OpeningBalance.StackNo, "
tmp = tmp & " Txn_GSL_OpeningBalance.LotNo, Txn_GSL_OpeningBalance.DepositBags,Txn_GSL_OpeningBalance.DepositWeight , "
tmp = tmp & " Txn_GSL_OpeningBalance.SM_Prefix,Txn_GSL_OpeningBalance.CMPID , Txn_GSL_OpeningBalance.godownid, "
tmp = tmp & " Txn_GSL_OpeningBalance.ExchangeTypeID , Txn_GSL_OpeningBalance.CommodityID, Txn_GSL_OpeningBalance.ISINID, Txn_GSL_OpeningBalance.G_UID "
tmp = tmp & " FROM  Txn_GSL_OpeningBalance Left JOIN"
tmp = tmp & " Mst_Client ON Txn_GSL_OpeningBalance.ClientIDRow = Mst_Client.ClientIDRow INNER JOIN"
tmp = tmp & " Mst_CMP ON Txn_GSL_OpeningBalance.CMPID = Mst_CMP.CMPID INNER JOIN"
tmp = tmp & " Mst_Commodity ON Txn_GSL_OpeningBalance.CommodityID = Mst_Commodity.CommodityID INNER JOIN"
tmp = tmp & " Mst_ExchangeType ON Txn_GSL_OpeningBalance.ExchangeTypeID = Mst_ExchangeType.ExchangeTypeID INNER JOIN"
tmp = tmp & " Mst_Godown ON Txn_GSL_OpeningBalance.GodownID = Mst_Godown.godownid " & Wc
tmp = tmp & " ORDER BY Mst_ExchangeType.ExchangeType, Txn_GSL_OpeningBalance.ClientIDRow, Mst_Commodity.Commodity, Mst_CMP.CMPName , "
tmp = tmp & " Mst_Godown.GodownNo , Txn_GSL_OpeningBalance.StackNo, Txn_GSL_OpeningBalance.LotNo "

Call TmpTable
MSHFlexGrid1.Rows = TmpRs.RecordCount + 2
For I = 1 To TmpRs.RecordCount
 
    MSHFlexGrid1.TextMatrix(I, 0) = TmpRs!GSLID
    MSHFlexGrid1.TextMatrix(I, 1) = TmpRs!SM_Prefix
    MSHFlexGrid1.TextMatrix(I, 2) = TmpRs!CMPName
    MSHFlexGrid1.TextMatrix(I, 3) = TmpRs!Commodity
    MSHFlexGrid1.TextMatrix(I, 4) = TmpRs!GodownNo
    MSHFlexGrid1.TextMatrix(I, 5) = TmpRs!StackNo
    MSHFlexGrid1.TextMatrix(I, 6) = TmpRs!LotNo
    MSHFlexGrid1.TextMatrix(I, 7) = TmpRs!DepositBags
    MSHFlexGrid1.TextMatrix(I, 8) = Format(TmpRs!DepositWeight, "########0.000")
    MSHFlexGrid1.TextMatrix(I, 9) = TmpRs!ExchangeType
    MSHFlexGrid1.TextMatrix(I, 10) = IIf(IsNull(TmpRs!ClientName), "", TmpRs!ClientName)
    MSHFlexGrid1.TextMatrix(I, 11) = IIf(IsNull(TmpRs!ISINID), "", TmpRs!ISINID)
    MSHFlexGrid1.TextMatrix(I, 12) = IIf(IsNull(TmpRs!G_UID), "", TmpRs!G_UID)
    TmpRs.MoveNext
Next
End Sub
Function GenWc(wc_ As String, ID As Variant, Field_ As String, Type_ As String) As String
Dim Retval
Retval = wc_
If Retval = "" Then
   If Type_ = "S" Then
      Retval = "Where " & Field_ & "= '" & ID & "'"
   Else
      Retval = "Where " & Field_ & "= " & ID
   End If
Else
   If Type_ = "S" Then
      Retval = Retval & " And " & Field_ & "= '" & ID & "'"
   Else
      Retval = Retval & " And " & Field_ & "= " & ID
   End If
End If
GenWc = Retval
End Function

Private Function GetGSLID() As Double
Dim Retval As Double
tmp = "Select max(GSLID) from Txn_GSL_OpeningBalance where SM_Prefix = '" & TxtPreFix & "'"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetGSLID = Retval
End Function

Private Sub CmdExcel_Click()
Call Gen_Create_Xls(Gen_UserName & "_" & "Txn_GSL_OpeningBalance.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_UserName & "_" & "Txn_GSL_OpeningBalance.xls")
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
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub
Private Sub DeleteCmd_Click()
Call TableMst_GSL(" where sm_prefix = '" & TxtPreFix & "'")
If RsMst_GSL.RecordCount > 0 Then
   MsgBox "Import process already executed for selected location,you can't delete this Record!!!"
   Exit Sub
End If

Dim ans As Variant
On Error GoTo msgError
ans = MsgBox("Do you really want to DELETE this Record???", vbYesNo)
If ans = vbYes Then
   RsTxn_GSL_OpeningBalance.MoveFirst
   RsTxn_GSL_OpeningBalance.Find "G_UID = '" & TxtG_UID & "' "
   RsTxn_GSL_OpeningBalance.Delete
   RsTxn_GSL_OpeningBalance.Update
   If RsTxn_GSL_OpeningBalance.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsTxn_GSL_OpeningBalance.MoveNext
   If RsTxn_GSL_OpeningBalance.EOF() Then
      RsTxn_GSL_OpeningBalance.MoveLast
   End If
   Call Indata
End If
Exit Sub
msgError:
MsgBox "Child record Present"
RsTxn_GSL_OpeningBalance.CancelUpdate
End Sub
