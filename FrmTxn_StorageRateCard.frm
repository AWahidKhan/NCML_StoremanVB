VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmTxn_StorageRateCard 
   Caption         =   "Storage Rate Card"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame ButtonFrm 
      Height          =   1200
      Left            =   120
      TabIndex        =   37
      Top             =   7920
      Width           =   15015
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   400
         Left            =   12105
         TabIndex        =   47
         Top             =   600
         Width           =   2820
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "Search"
         Height          =   400
         Left            =   12105
         TabIndex        =   46
         Top             =   195
         Width           =   2820
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   400
         Left            =   9165
         TabIndex        =   45
         Top             =   600
         Width           =   2940
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   400
         Left            =   9165
         TabIndex        =   44
         Top             =   195
         Width           =   2940
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   400
         Left            =   6225
         TabIndex        =   43
         Top             =   600
         Width           =   2940
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   400
         Left            =   6225
         TabIndex        =   42
         Top             =   195
         Width           =   2940
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   400
         Left            =   3285
         TabIndex        =   41
         Top             =   600
         Width           =   2940
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   400
         Left            =   3285
         TabIndex        =   40
         Top             =   195
         Width           =   2940
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   400
         Left            =   345
         TabIndex        =   39
         Top             =   600
         Width           =   2940
      End
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New"
         Height          =   400
         Left            =   345
         TabIndex        =   38
         Top             =   195
         Width           =   2940
      End
   End
   Begin VB.Frame Frame1 
      Height          =   7815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   15135
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
         Left            =   2265
         Locked          =   -1  'True
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   1920
         Width           =   4095
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
         Left            =   8385
         Locked          =   -1  'True
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   1920
         Width           =   4095
      End
      Begin VB.TextBox TxtPreFix 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   375
         Left            =   6960
         Locked          =   -1  'True
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   195
         Width           =   1455
      End
      Begin VB.Frame Frame2 
         Caption         =   "Exception "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   120
         TabIndex        =   36
         Top             =   2400
         Width           =   14775
         Begin VB.ComboBox CmbCMPID 
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
            Left            =   2040
            Sorted          =   -1  'True
            TabIndex        =   16
            Top             =   1065
            Width           =   4140
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
            Left            =   7560
            Sorted          =   -1  'True
            TabIndex        =   17
            Top             =   1065
            Width           =   4125
         End
         Begin VB.TextBox TxtPremium 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   2040
            MaxLength       =   50
            TabIndex        =   11
            Top             =   600
            Width           =   1875
         End
         Begin VB.TextBox TxtSpDiscount 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   120
            MaxLength       =   100
            TabIndex        =   10
            Top             =   600
            Width           =   1860
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
            Left            =   13080
            TabIndex        =   56
            Top             =   1058
            Width           =   1590
         End
         Begin VB.ComboBox CmbClientDP 
            Height          =   315
            Left            =   9720
            Sorted          =   -1  'True
            TabIndex        =   13
            Top             =   600
            Visible         =   0   'False
            Width           =   5055
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
            Left            =   6240
            Locked          =   -1  'True
            TabIndex        =   51
            TabStop         =   0   'False
            Top             =   600
            Width           =   3495
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
            Left            =   3960
            TabIndex        =   12
            Top             =   600
            Width           =   2295
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
            Left            =   9600
            Locked          =   -1  'True
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   600
            Width           =   1215
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
            Left            =   10800
            Locked          =   -1  'True
            TabIndex        =   49
            TabStop         =   0   'False
            Top             =   600
            Width           =   3855
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
            Left            =   11805
            TabIndex        =   18
            Top             =   1058
            Width           =   1215
         End
         Begin VB.ComboBox CmbNonNcdexClient 
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
            Left            =   6240
            Sorted          =   -1  'True
            TabIndex        =   15
            Top             =   240
            Visible         =   0   'False
            Width           =   495
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
            Left            =   10920
            TabIndex        =   61
            Top             =   300
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label LblCMPName 
            Caption         =   "CMP"
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
            TabIndex        =   60
            Top             =   1080
            Width           =   600
         End
         Begin VB.Label lblGodownNo 
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
            Left            =   6360
            TabIndex        =   59
            Top             =   1080
            Width           =   1095
         End
         Begin VB.Label Label3 
            Caption         =   "Premium"
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
            Left            =   2400
            TabIndex        =   58
            Top             =   240
            Width           =   1020
         End
         Begin VB.Label Label5 
            Caption         =   "Special Discount"
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
            TabIndex        =   57
            Top             =   240
            Width           =   1800
         End
         Begin VB.Label Label2 
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
            Left            =   4920
            TabIndex        =   55
            Top             =   300
            Width           =   975
         End
         Begin VB.Label Label8 
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
            Left            =   6960
            TabIndex        =   54
            Top             =   300
            Width           =   1095
         End
         Begin VB.Label Label1 
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
            Left            =   9960
            TabIndex        =   53
            Top             =   300
            Width           =   735
         End
         Begin VB.Label Label11 
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
            Left            =   12120
            TabIndex        =   52
            Top             =   300
            Width           =   1095
         End
      End
      Begin VB.TextBox TxtToBagSize 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   325
         Left            =   13440
         MaxLength       =   50
         TabIndex        =   8
         Top             =   1120
         Width           =   1380
      End
      Begin VB.TextBox TxtRatePerBag 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   325
         Left            =   13440
         TabIndex        =   9
         Top             =   1530
         Width           =   1380
      End
      Begin VB.TextBox TxtFromBagSize 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   325
         Left            =   13440
         MaxLength       =   100
         TabIndex        =   7
         Top             =   705
         Width           =   1380
      End
      Begin VB.CommandButton CmdCommodityGroupID 
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
         Height          =   360
         Left            =   8640
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   1485
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
         Left            =   1200
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   600
         Width           =   7200
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
         Left            =   8640
         TabIndex        =   25
         Top             =   630
         Width           =   375
      End
      Begin VB.TextBox TxtSRCID 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   375
         Left            =   1215
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   195
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker FromDate 
         Height          =   330
         Left            =   10455
         TabIndex        =   4
         Top             =   705
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         Format          =   112197633
         CurrentDate     =   39884
      End
      Begin MSComCtl2.DTPicker ToDate 
         Height          =   330
         Left            =   10455
         TabIndex        =   5
         Top             =   1125
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         Format          =   112197633
         CurrentDate     =   39884
      End
      Begin MSComCtl2.DTPicker TillDate 
         Height          =   330
         Left            =   10455
         TabIndex        =   6
         Top             =   1530
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         Format          =   112197633
         CurrentDate     =   39884
      End
      Begin MSComCtl2.DTPicker TxtSRCDate 
         Height          =   375
         Left            =   4200
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   195
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   112197633
         CurrentDate     =   39884
      End
      Begin VB.TextBox TxtLocationID 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1200
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   600
         Width           =   7200
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
         Height          =   3570
         Left            =   120
         TabIndex        =   48
         Top             =   4080
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   6297
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
      Begin VB.ComboBox CmbCommodityGroupID 
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
         Left            =   1200
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   1485
         Width           =   7200
      End
      Begin VB.TextBox TxtCommodityGroupID 
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
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   1485
         Width           =   7200
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
         ItemData        =   "FrmTxn_StorageRateCard.frx":0000
         Left            =   1200
         List            =   "FrmTxn_StorageRateCard.frx":0002
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   1050
         Width           =   7200
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
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   1050
         Width           =   7200
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
         Left            =   120
         TabIndex        =   67
         Top             =   1980
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
         Left            =   7350
         TabIndex        =   66
         Top             =   1980
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Location ID"
         Height          =   300
         Left            =   5850
         TabIndex        =   63
         Top             =   240
         Width           =   1080
      End
      Begin VB.Label LblBagsSizeTo 
         Caption         =   "Bag Size To"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   12075
         TabIndex        =   35
         Top             =   1125
         Width           =   1200
      End
      Begin VB.Label LblRatePerBag 
         Caption         =   "Rate Per Bag"
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
         Left            =   12075
         TabIndex        =   34
         Top             =   1590
         Width           =   1200
      End
      Begin VB.Label LblBagSizeFrom 
         Caption         =   "Bag Size From"
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
         Left            =   12075
         TabIndex        =   33
         Top             =   720
         Width           =   1440
      End
      Begin VB.Label Label7 
         Caption         =   "Mode"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   90
         TabIndex        =   32
         Top             =   1125
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Commodity"
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
         TabIndex        =   31
         Top             =   1590
         Width           =   1215
      End
      Begin VB.Label LblLocationID 
         Caption         =   "Location"
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
         Left            =   90
         TabIndex        =   27
         Top             =   720
         Width           =   960
      End
      Begin VB.Label LblWHRNo 
         Caption         =   "SRC ID"
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
         TabIndex        =   24
         Top             =   255
         Width           =   900
      End
      Begin VB.Label LblWHRDate 
         Caption         =   "From Date"
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
         Left            =   9330
         TabIndex        =   23
         Top             =   720
         Width           =   1125
      End
      Begin VB.Label Label10 
         Caption         =   "Till Date"
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
         Left            =   9330
         TabIndex        =   22
         Top             =   1530
         Width           =   960
      End
      Begin VB.Label Label9 
         Caption         =   "To Date"
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
         Left            =   9330
         TabIndex        =   21
         Top             =   1125
         Width           =   960
      End
      Begin VB.Label Label41 
         Caption         =   "SRC Date"
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
         Left            =   2835
         TabIndex        =   20
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FrmTxn_StorageRateCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg, mExchangeTypeID, mCommodityGroupId, mLOcationID, mGodownID, mCMPID As Variant

Private Sub Grid_Head()
MSHFlexGrid2.Clear
MSHFlexGrid2.Rows = 2
MSHFlexGrid2.FixedRows = 1
MSHFlexGrid2.Cols = 12

MSHFlexGrid2.TextMatrix(0, 0) = "Row ID"
MSHFlexGrid2.TextMatrix(0, 1) = "Special Discount"
MSHFlexGrid2.TextMatrix(0, 2) = "Premium "
MSHFlexGrid2.TextMatrix(0, 3) = "Client ID"
MSHFlexGrid2.TextMatrix(0, 4) = "Client Name"
MSHFlexGrid2.TextMatrix(0, 5) = "DP ID"
MSHFlexGrid2.TextMatrix(0, 6) = "DP Name "
MSHFlexGrid2.TextMatrix(0, 7) = "ClientRowID"
MSHFlexGrid2.TextMatrix(0, 8) = "CMP"
MSHFlexGrid2.TextMatrix(0, 9) = "Godown No"
MSHFlexGrid2.TextMatrix(0, 10) = "CMPID"
MSHFlexGrid2.TextMatrix(0, 11) = "GodownID"

MSHFlexGrid2.RowHeight(0) = 800
MSHFlexGrid2.ColWidth(0) = 500
MSHFlexGrid2.ColWidth(1) = 1200
MSHFlexGrid2.ColWidth(2) = 1200
MSHFlexGrid2.ColWidth(3) = 1200
MSHFlexGrid2.ColWidth(4) = 2500
MSHFlexGrid2.ColWidth(5) = 1200
MSHFlexGrid2.ColWidth(6) = 2500
MSHFlexGrid2.ColWidth(7) = 0
MSHFlexGrid2.ColWidth(8) = 2500
MSHFlexGrid2.ColWidth(9) = 1600
MSHFlexGrid2.ColWidth(10) = 0
MSHFlexGrid2.ColWidth(11) = 0

End Sub

Private Sub CmbCommodityID_Change()

End Sub

Private Sub CmbCMPName_Change()

End Sub

Private Sub CmbCMPID_GotFocus()
    If CmbLocationID.Text = "" Then
       MsgBox "Please select 'Location' from  dropdown list!!!"
       CmbLocationID.SetFocus
       Exit Sub
    End If
    tmp = CmbCMPID.Text
    Call TableMst_CMP(" where LocationID = " & mLOcationID)
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
   MsgBox "Invalid selection "
   CmbCMPID.SetFocus
   Exit Sub
End If
mCMPID = RsMst_CMP!CMPID
TxtPreFix = GetPreFix(RsMst_CMP!CMPID, "P")

End Sub

Private Sub CmbCommodityGroupID_GotFocus()

tmp = CmbCommodityGroupID.Text
Call TableMst_CommodityGroup
CmbCommodityGroupID.Clear
If RsMst_CommodityGroup.RecordCount = 0 Then
  MsgBox "No Record found !!!! "
  Exit Sub
End If
For i = 1 To RsMst_CommodityGroup.RecordCount
   CmbCommodityGroupID.AddItem RsMst_CommodityGroup!COmmodityGroup
   RsMst_CommodityGroup.MoveNext
Next
CmbCommodityGroupID.Text = tmp
    
End Sub

'----
Private Sub CmbCommodityGroupID_LostFocus()
If CmbCommodityGroupID.Text = "" Then
   Exit Sub
End If

RsMst_CommodityGroup.MoveFirst
RsMst_CommodityGroup.Find "CommodityGroup = '" & CmbCommodityGroupID.Text & "'"
If RsMst_CommodityGroup.EOF Then
   MsgBox "Invalid selection "
   CmbCommodityGroupID.SetFocus
   Exit Sub
End If
mCommodityGroupId = RsMst_CommodityGroup!CommodityGroupID
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
   MsgBox "Invalid selection "
   CmbExchangeTypeID.SetFocus
   Exit Sub
End If

'If mExchangeTypeID <> RsMst_ExchangeType!ExchangeTypeID Then
'   Call Grid_Head_Client
'   Call Grid_Head_GCTD
'   Call ClearClientTabControl
'   Call ClearGCTDTabControl
'   CmbContractID.Text = ""
'   Check1.Value = 0
'   CmdClientAdd2List.Enabled = True
'End If
mExchangeTypeID = RsMst_ExchangeType!ExchangeTypeID


End Sub

Private Sub CmbGodownNo_GotFocus()
If CmbCMPID.Text = "" Then
   MsgBox "Select CMP !!!! "
   CmbCMPID.SetFocus
   Exit Sub
End If
tmp = CmbGodownNo.Text
Call TableMst_Godown(" Where CloseDate is null And CMPID = " & mCMPID)
CmbGodownNo.Clear
If RsMst_Godown.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_Godown.RecordCount
    CmbGodownNo.AddItem RsMst_Godown!godownno
    RsMst_Godown.MoveNext
Next
CmbGodownNo.Text = tmp

End Sub
'----
Private Sub CmbGodownNo_LostFocus()
If CmbGodownNo.Text = "" Then
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
mGodownID = RsMst_Godown!GodownID
End Sub

Private Sub CmbLocationID_GotFocus()

tmp = CmbLocationID.Text
Call TableMst_Location
CmbLocationID.Clear
If RsMst_Location.RecordCount = 0 Then
  MsgBox "No Record found !!!! "
  Exit Sub
End If
For i = 1 To RsMst_Location.RecordCount
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
RsMst_Location.Find "Locationname = '" & CmbLocationID.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection "
   CmbLocationID.SetFocus
   Exit Sub
End If
If mLOcationID <> RsMst_Location!LocationID Then
   CmbCMPID.Text = ""
   CmbGodownNo.Text = ""
   mGodownID = 0
   mCMPID = 0
   
End If
mLOcationID = RsMst_Location!LocationID

End Sub

Private Sub CmdCommodityGroupID_Click()
FrmMst_CommodityGroup.Show
End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call Grid_Head
End Sub

