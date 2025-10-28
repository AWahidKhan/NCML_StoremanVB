VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm_Spillage 
   Caption         =   "Spillage Transaction"
   ClientHeight    =   7905
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10710
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   7905
   ScaleWidth      =   10710
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   1485
      Left            =   120
      TabIndex        =   42
      Top             =   9120
      Visible         =   0   'False
      Width           =   15075
      Begin VB.ComboBox CmbSTxnType 
         Height          =   315
         ItemData        =   "Frm_Spillage.frx":0000
         Left            =   6420
         List            =   "Frm_Spillage.frx":000A
         Sorted          =   -1  'True
         TabIndex        =   72
         Top             =   810
         Width           =   2685
      End
      Begin VB.ComboBox CmbLocation 
         Height          =   315
         Left            =   3195
         Sorted          =   -1  'True
         TabIndex        =   52
         Top             =   330
         Width           =   3165
      End
      Begin VB.ComboBox CmbSGodown 
         Height          =   315
         Left            =   9627
         Sorted          =   -1  'True
         TabIndex        =   54
         Top             =   330
         Width           =   2925
      End
      Begin VB.ComboBox CmbSStack 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "Frm_Spillage.frx":0026
         Left            =   12595
         List            =   "Frm_Spillage.frx":0028
         Sorted          =   -1  'True
         TabIndex        =   55
         Top             =   330
         Width           =   2295
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "Clear"
         Height          =   350
         Left            =   13890
         TabIndex        =   60
         Top             =   795
         Width           =   1000
      End
      Begin VB.ComboBox CmbSCommodity 
         Height          =   315
         Left            =   1680
         Sorted          =   -1  'True
         TabIndex        =   56
         Top             =   813
         Width           =   4725
      End
      Begin VB.ComboBox CmbSCmp 
         Height          =   315
         Left            =   6420
         Sorted          =   -1  'True
         TabIndex        =   53
         Top             =   330
         Width           =   3165
      End
      Begin VB.CommandButton CmdGo 
         Caption         =   "Go"
         Height          =   350
         Left            =   11880
         TabIndex        =   57
         Top             =   795
         Width           =   1000
      End
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   350
         Left            =   12885
         TabIndex        =   59
         Top             =   795
         Width           =   1000
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1455
         Left            =   75
         TabIndex        =   43
         Top             =   1320
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   2566
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSComCtl2.DTPicker STo 
         Height          =   300
         Left            =   1680
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   337
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   48496641
         CurrentDate     =   36494
      End
      Begin MSComCtl2.DTPicker SFrom 
         Height          =   300
         Left            =   120
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   330
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   48496641
         CurrentDate     =   36494
      End
      Begin VB.Label Label18 
         Caption         =   "Bags Type"
         Height          =   180
         Left            =   7335
         TabIndex        =   73
         Top             =   630
         Width           =   855
      End
      Begin VB.Label Label16 
         Caption         =   "Location"
         Height          =   180
         Left            =   4410
         TabIndex        =   69
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label15 
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
         Left            =   11715
         TabIndex        =   67
         Top             =   855
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
         Left            =   9240
         TabIndex        =   66
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label Label14 
         Caption         =   "Godown"
         Height          =   240
         Left            =   10725
         TabIndex        =   65
         Top             =   120
         Width           =   720
      End
      Begin VB.Label Label21 
         Caption         =   "Stack No"
         Height          =   180
         Left            =   13375
         TabIndex        =   64
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         Caption         =   "Commodity"
         Height          =   195
         Left            =   465
         TabIndex        =   63
         Top             =   873
         Width           =   765
      End
      Begin VB.Label Label42 
         Caption         =   "CMP"
         Height          =   300
         Left            =   7755
         TabIndex        =   62
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label44 
         Caption         =   "From Date"
         Height          =   255
         Left            =   420
         TabIndex        =   61
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label45 
         Caption         =   "To Date"
         Height          =   255
         Left            =   1980
         TabIndex        =   58
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   6630
      Left            =   120
      ScaleHeight     =   6570
      ScaleWidth      =   8955
      TabIndex        =   21
      Top             =   120
      Width           =   9015
      Begin VB.TextBox TxtBalanceMadeUp 
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
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   84
         TabStop         =   0   'False
         Top             =   2280
         Width           =   1950
      End
      Begin VB.TextBox TxtExchangeTypeID 
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
         Left            =   6360
         Locked          =   -1  'True
         TabIndex        =   79
         TabStop         =   0   'False
         Top             =   5730
         Width           =   2070
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
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   78
         TabStop         =   0   'False
         Top             =   6150
         Width           =   6870
      End
      Begin VB.TextBox TxtMadeUpBags 
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   76
         TabStop         =   0   'False
         Top             =   2280
         Width           =   1950
      End
      Begin VB.TextBox TxtNewGSLID 
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   4320
         Visible         =   0   'False
         Width           =   6855
      End
      Begin VB.ComboBox CmbTxnType 
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
         ItemData        =   "Frm_Spillage.frx":002A
         Left            =   1560
         List            =   "Frm_Spillage.frx":0034
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   600
         Width           =   6885
      End
      Begin VB.TextBox TxtGSLFlag 
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
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   3600
         Visible         =   0   'False
         Width           =   270
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   5265
         Width           =   6855
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   4815
         Width           =   6855
      End
      Begin VB.TextBox TxtPreFix 
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
         Left            =   7680
         Locked          =   -1  'True
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   120
         Width           =   990
      End
      Begin VB.CommandButton CmdCommodityID 
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
         Left            =   8520
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   1845
         Width           =   375
      End
      Begin VB.CommandButton CmdGSL 
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
         Left            =   8520
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   1440
         Width           =   375
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
         Left            =   1560
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   1845
         Width           =   6885
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
         Height          =   1080
         Left            =   1560
         MaxLength       =   1000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   3180
         Width           =   6855
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
         Left            =   1560
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   1440
         Width           =   1980
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
         ItemData        =   "Frm_Spillage.frx":0050
         Left            =   4800
         List            =   "Frm_Spillage.frx":0052
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   1440
         Width           =   1215
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
         ItemData        =   "Frm_Spillage.frx":0054
         Left            =   6960
         List            =   "Frm_Spillage.frx":0056
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   1440
         Width           =   1470
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
         Left            =   1560
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   1035
         Width           =   6885
      End
      Begin VB.TextBox TxtSpillageQuantity 
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
         Left            =   5880
         TabIndex        =   9
         Top             =   2760
         Width           =   2550
      End
      Begin VB.TextBox TxtSpillageNoOfBags 
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
         Left            =   1560
         TabIndex        =   8
         Top             =   2760
         Width           =   2190
      End
      Begin VB.TextBox TxtNoOfBags 
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   5730
         Width           =   2085
      End
      Begin VB.TextBox TxtCommodity 
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   1845
         Width           =   6870
      End
      Begin VB.TextBox TxtSpillageID 
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   120
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker SpillageDate 
         Height          =   375
         Left            =   5160
         TabIndex        =   1
         Top             =   120
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   48496641
         CurrentDate     =   39884
      End
      Begin VB.TextBox TxtCMPID 
         Appearance      =   0  'Flat
         Height          =   325
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   1050
         Width           =   6855
      End
      Begin VB.TextBox TxtGodownID 
         Appearance      =   0  'Flat
         Height          =   325
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1965
      End
      Begin VB.TextBox TxtStack 
         Appearance      =   0  'Flat
         Height          =   325
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1185
      End
      Begin VB.TextBox TxtLot 
         Appearance      =   0  'Flat
         Height          =   325
         Left            =   6960
         Locked          =   -1  'True
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox TxtTxnType 
         Appearance      =   0  'Flat
         Height          =   325
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   615
         Width           =   6855
      End
      Begin VB.Label Label24 
         Caption         =   "Balance Made up bags "
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
         TabIndex        =   85
         Top             =   2340
         Width           =   2160
      End
      Begin VB.Label LblExchangeTypeID 
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
         Height          =   300
         Left            =   4920
         TabIndex        =   81
         Top             =   5760
         Width           =   1440
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
         Height          =   300
         Left            =   120
         TabIndex        =   80
         Top             =   6180
         Width           =   1800
      End
      Begin VB.Label Label20 
         Caption         =   "No of Made up bags so far "
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
         Left            =   120
         TabIndex        =   77
         Top             =   2205
         Width           =   1770
      End
      Begin VB.Label LblNewGSL 
         Caption         =   "New GSL"
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
         TabIndex        =   75
         Top             =   4380
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label17 
         Caption         =   "Bags Type"
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
         TabIndex        =   71
         Top             =   630
         Width           =   1335
      End
      Begin VB.Label Label13 
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
         Left            =   120
         TabIndex        =   49
         Top             =   5325
         Width           =   975
      End
      Begin VB.Label Label12 
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
         TabIndex        =   48
         Top             =   4875
         Width           =   1695
      End
      Begin VB.Label Label11 
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
         Height          =   255
         Left            =   6600
         TabIndex        =   45
         Top             =   180
         Width           =   1095
      End
      Begin VB.Label Label10 
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
         Height          =   240
         Left            =   120
         TabIndex        =   39
         Top             =   3600
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Quantity in MT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4320
         TabIndex        =   38
         Top             =   2790
         Width           =   1575
      End
      Begin VB.Label Label7 
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
         Left            =   120
         TabIndex        =   37
         Top             =   2790
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Deposit No Of Bags in Original Lot"
         Height          =   450
         Left            =   120
         TabIndex        =   35
         Top             =   5685
         Width           =   1440
      End
      Begin VB.Label Label4 
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
         Left            =   120
         TabIndex        =   33
         Top             =   1905
         Width           =   1335
      End
      Begin VB.Label Label22 
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
         Height          =   255
         Left            =   6120
         TabIndex        =   30
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label3 
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
         Height          =   255
         Left            =   3720
         TabIndex        =   29
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label2 
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
         Left            =   135
         TabIndex        =   26
         Top             =   1485
         Width           =   1335
      End
      Begin VB.Label Label1 
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
         Left            =   120
         TabIndex        =   25
         Top             =   1065
         Width           =   1335
      End
      Begin VB.Label Label26 
         Caption         =   "Txn  Date"
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
         Left            =   3600
         TabIndex        =   24
         Top             =   180
         Width           =   1575
      End
      Begin VB.Label Label9 
         Caption         =   "Txn ID"
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
         TabIndex        =   23
         Top             =   180
         Width           =   1335
      End
   End
   Begin VB.PictureBox Picture3 
      Height          =   975
      Left            =   120
      ScaleHeight     =   915
      ScaleWidth      =   8970
      TabIndex        =   14
      Top             =   6795
      Width           =   9030
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New "
         Height          =   375
         Left            =   240
         TabIndex        =   0
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   375
         Left            =   1935
         TabIndex        =   11
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   3630
         TabIndex        =   20
         Top             =   120
         Width           =   1680
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   5310
         TabIndex        =   19
         Top             =   120
         Width           =   1680
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "&Search"
         Height          =   375
         Left            =   6990
         TabIndex        =   12
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   495
         Width           =   1695
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   375
         Left            =   1935
         TabIndex        =   17
         Top             =   495
         Width           =   1695
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   375
         Left            =   3630
         TabIndex        =   16
         Top             =   495
         Width           =   1680
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   375
         Left            =   5310
         TabIndex        =   15
         Top             =   495
         Width           =   1680
      End
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   375
         Left            =   6990
         TabIndex        =   13
         Top             =   495
         Width           =   1695
      End
   End
   Begin VB.Label Label27 
      Caption         =   "Made Up Bags Information"
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
      Left            =   9360
      TabIndex        =   87
      Top             =   3600
      Width           =   2895
   End
   Begin VB.Label Label25 
      Caption         =   "Spillage Information"
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
      Left            =   9360
      TabIndex        =   86
      Top             =   0
      Width           =   2175
   End
   Begin VB.Label Label19 
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"Frm_Spillage.frx":0058
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3900
      Left            =   9360
      TabIndex        =   83
      Top             =   3840
      Width           =   5685
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"Frm_Spillage.frx":0257
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3000
      Left            =   9360
      TabIndex        =   82
      Top             =   240
      Width           =   5685
   End
End
Attribute VB_Name = "Frm_Spillage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mGodownID, mCommodityID As Double
Dim mGSLID As Double
Dim Edit_Flg As Variant
Dim mCMPID As Double
Dim mSpillageNoOfBags As Double
Dim mSpillageQuantity As Double
Dim mSGodownID, mSCMPID, mSLocationID As Double
Dim mSSMPrefix       As Variant
Dim mConnectedGSLID, mNewGSLID, mExchangeTypeID, mClientIDRow As Variant

Private Sub AddCmd_Click()

Edit_Flg = "A"
TxtGSLFlag = ""
CmbCMPID.Visible = True
CmbGodownNo.Visible = True
CmbStack.Visible = True
CmbLot.Visible = True
CmbCommodityID.Visible = False
TxtCMPID.Visible = False
TxtGodownID.Visible = False
TxtStack.Visible = False
TxtLot.Visible = False
CmdCommodityID.Visible = False
TxtSpillageID = ""
SpillageDate.Value = Date
'SpillageDate.Enabled = True
CmbCMPID.Text = ""
CmbGodownNo.Text = ""
CmbStack.Text = ""
CmbLot.Text = ""
CmbCommodityID.Text = ""
TxtCommodity = ""
TxtNoofBags = ""
TxtQuantity = ""
TxtSpillageNoOfBags = ""
TxtSpillageQuantity = ""
TxtRemarks = ""
TxtCreated = ""
TxtUpdated = ""
TxtPreFix = ""
TxtSpillageNoOfBags.Locked = False
TxtSpillageQuantity.Locked = False
TxtRemarks.Locked = False
CmdGSL.Visible = True

CmbTxnType.Visible = True
TxtTxnType.Visible = False
CmbTxnType.Text = ""


TxtMadeupBags = ""
TxtBalanceMadeUp = ""
TxtExchangeTypeID = ""
TxtClientID = ""
TxtNewGSLID = ""

TxtExchangeTypeID.Visible = False
TxtClientID.Visible = False
LblExchangeTypeID.Visible = False
LblClientID.Visible = False


Call GButtonLock(Me, False)

End Sub

Private Sub CmbLocation_GotFocus()
tmp = CmbLocation.Text

If Gen_WcLocation = "" Then
   Call TableMst_Location("")
Else
   Call TableMst_Location("where " & Gen_WcLocation)
End If
CmbLocation.Clear

If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Location found"
   CmbLocation.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbLocation.AddItem RsMst_Location!Locationname
    RsMst_Location.MoveNext
Next
CmbLocation.Text = tmp
End Sub

Private Sub CmbLocation_LostFocus()
If CmbLocation.Text = "" Then
   CmbSCmp.Text = ""
   CmbSGodown.Text = ""
   CmbSStack.Text = ""
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocation.Text & "'"

If RsMst_Location.EOF Then
   MsgBox "Invalid Location selection "
   CmbLocation.SetFocus
   Exit Sub
End If
CmbSCmp.Clear
mSLocationID = RsMst_Location.Fields("LocationID")
mSSMPrefix = RsMst_Location.Fields("SM_Prefix")
End Sub

Private Sub CmbSTxnType_LostFocus()
If CmbSTxnType.Text = "" Then Exit Sub

If LCase(CmbSTxnType.Text) <> "made up bags" And LCase(CmbSTxnType.Text) <> "spillage" Then
   MsgBox "Invalid Selection!!!"
   CmbSTxnType.SetFocus
   Exit Sub
End If


End Sub

Private Sub CmbTxnType_GotFocus()
tmp = CmbTxnType.Text
End Sub

Private Sub CmbTxnType_LostFocus()

TxtSpillageQuantity.Locked = False
LblNewGSL.Visible = False
TxtNewGSLID.Visible = False
TxtExchangeTypeID.Visible = False
TxtClientID.Visible = False
LblExchangeTypeID.Visible = False
LblClientID.Visible = False

TxtMadeupBags = ""
TxtExchangeTypeID = ""
TxtClientID = ""

If CmbTxnType.Text = "" Then
   CmbCMPID.Text = ""
   CmbGodownNo.Text = ""
   CmbStack.Text = ""
   CmbLot.Text = ""
   CmbCommodityID.Text = ""
   TxtCommodity.Text = ""
   mCommodityID = 0
   mGSLID = 0
   Exit Sub
End If

If LCase(CmbTxnType.Text) <> "made up bags" And LCase(CmbTxnType.Text) <> "spillage" Then
   MsgBox "Invalid Selection!!!"
   CmbTxnType.SetFocus
   Exit Sub
End If

If tmp <> CmbTxnType.Text Then
   CmbCMPID.Text = ""
   CmbGodownNo.Text = ""
   CmbStack.Text = ""
   CmbLot.Text = ""
   CmbCommodityID.Text = ""
   TxtCommodity.Text = ""
   mCommodityID = 0
   mGSLID = 0
End If

If LCase(CmbTxnType.Text) = "made up bags" Then
   TxtSpillageQuantity = 0
   TxtSpillageQuantity.Locked = True
   LblNewGSL.Visible = True
   TxtNewGSLID.Visible = True
   TxtNewGSLID = ""
   TxtExchangeTypeID.Visible = True
   TxtClientID.Visible = True
   LblExchangeTypeID.Visible = True
   LblClientID.Visible = True
End If

End Sub

Private Sub CmdClear_Click()

SFrom.Value = CDate(From_GenDate) - 900 'Date
STo.Value = Date
CmbSCmp.Text = ""
CmbSGodown.Text = ""
CmbSStack.Text = ""
CmbSCommodity.Text = ""
CmbLocation.Text = ""
Label15.Caption = ""


End Sub

Private Sub CmdCommodityID_Click()
FrmMst_Commodity.Show
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "SpillageTxn.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "SpillageTxn.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = I + 1 ' + 6
    For C = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(I, C), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(I, C)) Then
              oWS.Cells(mRow, C + 1) = Format(MSHFlexGrid1.TextMatrix(I, C), "dd/MMM/yyyy")
           Else
              oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
           End If
        Else
           oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
        End If
    Next
Next
oWB.Save
Set oWB = Nothing
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub

Private Sub CmdGo_Click()
Dim wc As String
Call Grid_Head

If STo.Value < SFrom.Value Then
   MsgBox "To date must be always greater than from date!!!"
   STo.SetFocus
   Exit Sub
End If

tmp = "Select TS.Txn_Spillage_ID,TS.SM_Prefix,TS.SpillageDate,MC.CMPName,MG.GodownNo,MGSL.StackNo, " & _
      " MGSL.LotNo , MComm.Commodity, TS.SpillageNoOfBags, TS.SpillageQuantity, TS.Remarks, TS.G_UID,TS.Flag " & _
      " From Txn_Spillage TS  " & _
      " Inner Join Mst_GSL MGSL On TS.Mst_GSL_ID = MGSL.Mst_GSL_ID And TS.SM_prefix = MGSL.SM_prefix  " & _
      " Inner Join Mst_CMP MC On MGSL.CMPID = MC.CMPID  " & _
      " Inner Join Mst_Location ML On MC.LocationID = ML.LocationID" & _
      " Inner Join Mst_Godown MG On MGSL.GodownID = MG.GodownID  " & _
      " Inner Join Mst_Commodity MComm On TS.CommodityID = MComm.CommodityID  "

wc = ""
If Gen_DBType = "MDB" Then
   wc = " Where (TS.SpillageDate Between #" & Format(SFrom, Gen_DatFormat) & "# And #" & Format(STo, Gen_DatFormat) & "#) "
Else
   wc = " Where (TS.SpillageDate Between '" & Format(SFrom, Gen_DatFormat) & "' And '" & Format(STo, Gen_DatFormat) & "') "
End If

If CmbSCmp.Text <> "" Then
   wc = wc & " And MC.CMPName = '" & CmbSCmp.Text & "'"
End If

If CmbSGodown.Text <> "" Then
   wc = wc & " And MG.GodownNo = '" & CmbSGodown.Text & "'"
End If

If CmbSStack.Text <> "" Then
   wc = wc & " And MGSL.StackNo = '" & CmbSStack.Text & "'"
End If

If CmbSCommodity.Text <> "" Then
   wc = wc & " And MComm.Commodity = '" & CmbSCommodity.Text & "'"
End If

If CmbLocation.Text <> "" Then
   wc = wc & " And MC.LocationID = '" & mSLocationID & "'"
Else
   If Gen_WcLocation <> "" Then
      wc = wc & " And MC." & Gen_WcLocation
   End If
End If

If CmbSTxnType.Text <> "" Then
   wc = wc & " And TS.Flag = '" & Left(CmbSTxnType.Text, 1) & "'"
End If

tmp = tmp & wc

Call TmpTable

MSHFlexGrid1.Rows = TmpRs.RecordCount + 1

For I = 1 To TmpRs.RecordCount
    For C = 0 To TmpRs.Fields.Count - 1
        MSHFlexGrid1.TextMatrix(I, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
    Next
    If LCase(MSHFlexGrid1.TextMatrix(I, 12)) = "s" Then
       MSHFlexGrid1.TextMatrix(I, 12) = "Spillage"
    Else
       MSHFlexGrid1.TextMatrix(I, 12) = "Made Up Bags"
    End If
    
    
    TmpRs.MoveNext
Next

Label15.Caption = TmpRs.RecordCount

End Sub

Private Sub CmdGSL_Click()
Gen_GSLID(0) = False
Gen_GSLID(1) = mCMPID
Gen_GSLID(2) = mGodownID
Gen_GSLID(3) = CmbStack.Text
Gen_GSLID(4) = CmbLot.Text
Gen_GSLID(5) = CmbCMPID.Text
Gen_GSLID(6) = CmbGodownNo.Text


FrmMst_GSL.Show
End Sub

Private Sub DeleteCmd_Click()

Dim ans As Variant
'On Error GoTo msgError

mGSLID = RsTxn_Spillage!Mst_GSL_ID
ans = MsgBox("Do you really want to DELETE this record???", vbYesNo)
If ans = vbYes Then
   If LCase(CmbTxnType.Text) = "spillage" Then
      Call TableMst_GSL("Where sm_prefix ='" & TxtPreFix & "' And Mst_GSL_ID = " & mGSLID)
   ElseIf LCase(CmbTxnType.Text) = "made up bags" Then
      Call TableMst_GSL("Where sm_prefix ='" & TxtPreFix & "' And Mst_GSL_ID = " & mConnectedGSLID)
   End If
   If RsMst_GSL.RecordCount > 0 Then
      If Val(RsMst_GSL!BalanceBags) - Val(TxtSpillageNoOfBags) < 0 Then
         MsgBox "You Can't delete this Recored since some Bags is withdran from this lot !!! "
         Exit Sub
      End If
   
      If Val(RsMst_GSL!BalanceWeight) - Val(TxtSpillageQuantity) < 0 Then
         MsgBox "You Can't delete this Recored since some quantity is withdran from this lot !!! "
         Exit Sub
      End If
    
       RsMst_GSL!DepositeBags = Val(RsMst_GSL!DepositeBags) - Val(TxtSpillageNoOfBags)
       RsMst_GSL!DepositeWeight = Val(RsMst_GSL!DepositeWeight) - Val(TxtSpillageQuantity)
       RsMst_GSL!BalanceBags = Val(RsMst_GSL!BalanceBags) - Val(TxtSpillageNoOfBags)
       RsMst_GSL!BalanceWeight = Val(RsMst_GSL!BalanceWeight) - Val(TxtSpillageQuantity)
       If LCase(CmbTxnType.Text) = "spillage" Then
          If RsMst_GSL!DepositeWeight = 0 And RsMSt_CDTFWeight = 0 Then
             RsMst_GSL!CommodityID = Null
          End If
          tmp = "Delete from Txn_GSL Where Txn_Spillage_ID = " & Val(TxtSpillageID) & ""
          Call Tmp1Table
       
       ElseIf LCase(CmbTxnType.Text) = "made up bags" Then
          If RsMst_GSL!DepositeBags = 0 And RsMst_GSL!BalanceBags = 0 Then
    
             tmp = "Delete from Txn_GSL Where Mst_GSL_ID = " & mGSLID & ""
             Call Tmp1Table
    
             tmp = "Update Mst_GSL Set Connected_Mst_GSL_ID = 0 Where SM_Prefix = '" & TxtPreFix & "' And Mst_GSL_ID = " & mGSLID & ""
             Call Tmp1Table
             
             tmp = "Update Txn_Spillage Set NewMst_GSL_ID =null Where SM_Prefix = '" & TxtPreFix & "' And NewMst_GSL_ID = " & mNewGSLID & ""
             Call Tmp1Table
             
             RsMst_GSL.Delete
          End If
       End If
       RsMst_GSL.Update
   Else
       tmp = "Delete from Txn_GSL Where Mst_GSL_ID = " & mGSLID & ""
       Call Tmp1Table
       tmp = "Delete from Txn_GSL Where Txn_Spillage_ID = " & Val(TxtSpillageID) & ""
       Call Tmp1Table
   
   
   End If
   
   
   RsTxn_Spillage.Delete
   RsTxn_Spillage.Update
    
   If RsTxn_Spillage.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
    
   RsTxn_Spillage.MoveNext
   If RsTxn_Spillage.EOF() Then
      RsTxn_Spillage.MoveLast
   End If
   Call Indata
End If

Exit Sub

msgError:
MsgBox "Child record Present "

End Sub
Private Sub EditCmd_Click()

If UCase(TxtGSLFlag) = "E" Or UCase(TxtGSLFlag) = "Z" Then
   MsgBox "Lot Transfered or Fully delivered you can't edit !!!"
   Exit Sub
End If

'tmp = " select tad.godownid from Txn_AuditDetail tad"
'tmp = tmp & " inner join Txn_Audit ta on ta.auditid = tad.auditid"
'tmp = tmp & " where ta.BusinessSegmentID = 1 and tad.godownid = " & mGodownID
'tmp = tmp & " Union"
tmp = " select tm.godownid from Txn_MonthlyCompliance tm"
tmp = tmp & " where tm.BusinessSegmentID = 1 and tm.godownid = " & mGodownID
Call TmpTableHRPro

If TmpRsHRPro.RecordCount > 0 Then
   'If LCase(Gen_UserRole) <> "admin" And LCase(Gen_UserRole) <> "superadmin" And LCase(Gen_UserRole) <> "head cmg" And LCase(Gen_UserRole) <> "power" Then
   If LCase(Gen_UserRole) <> "power" Then
      MsgBox "Godown is already audited, you can not edit!!!"
      Exit Sub
   Else
      If MsgBox("Godown is already audited, Do you want to modify record?", vbYesNo) = vbNo Then Exit Sub
   End If
End If

Edit_Flg = "E"

If Val(TxtNoofBags) = Val(TxtSpillageNoOfBags) Then
   TxtCommodity.Visible = False
   CmbCommodityID.Text = TxtCommodity.Text
   CmbCommodityID.Visible = True
   CmbCommodityID.SetFocus
End If
TxtSpillageNoOfBags.Locked = False
If LCase(TxtTxnType) <> "made up bags" Then
   TxtSpillageQuantity.Locked = False
End If
mSpillageNoOfBags = TxtSpillageNoOfBags
mSpillageQuantity = TxtSpillageQuantity
TxtRemarks.Locked = False

TxtMadeupBags = Val(TxtMadeupBags) - Val(mSpillageNoOfBags)
TxtBalanceMadeUp = Val(TxtBalanceMadeUp) - Val(mSpillageNoOfBags)

Call GButtonLock(Me, False)

End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub
Private Sub CmbCMPID_GotFocus()
tmp = CmbCMPID.Text
wc = ""
If Gen_WcCMP <> "" Then
   wc = " And " & Gen_WcCMP
End If

Call TableMst_CMP(" where CloseDate is null " & wc)

CmbCMPID.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_CMP.RecordCount
   CmbCMPID.AddItem RsMst_CMP!CMPName
   RsMst_CMP.MoveNext
Next
CmbCMPID.Text = tmp

End Sub
'----
Private Sub CmbCMPID_LostFocus()
If CmbCMPID.Text = "" Then
   CmbCommodityID.Text = ""
   Exit Sub
End If
RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPName = '" & CmbCMPID.Text & "'"
If RsMst_CMP.EOF Then
   MsgBox "Invalid selection "
   CmbCMPID.SetFocus
   Exit Sub
End If
If mCMPID <> RsMst_CMP!CMPID Then
  CmbGodownNo.Clear
  CmbStack.Clear
  CmbLot.Clear
End If
mCMPID = RsMst_CMP!CMPID
TxtPreFix = GetPreFix(RsMst_CMP!CMPID, "P")
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
For I = 1 To RsMst_Godown.RecordCount
    CmbGodownNo.AddItem RsMst_Godown!GodownNo
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
   CmbGodownNo.SetFocus
   Exit Sub
End If
If mGodownID <> RsMst_Godown!GodownID Then
   CmbStack.Text = ""
   CmbLot.Text = ""
   mCmbStack = ""
   mCmbLot = ""
End If
mGodownID = RsMst_Godown!GodownID
End Sub
Private Sub CmbStack_GotFocus()

If CmbTxnType.Text = "" Then
  MsgBox "Please Select Bags Type!!!"
  CmbTxnType.SetFocus
  Exit Sub
End If

If CmbGodownNo.Text = "" Or mGodownID = "" Then
  MsgBox "Please Select Godown No!!!"
  Exit Sub
End If
'MsgBox mGodownID & " " & IsNull(mGodownID)
'If IsNull(mGodownID) Then
'  MsgBox "Please Select Godown No!!!"
'  Exit Sub
'End If

tmp = CmbStack.Text
If LCase(CmbTxnType.Text) = "spillage" Then
   Call TableMst_GSL(" Where  Status not in ('Z','E') And sm_prefix ='" & TxtPreFix & "' And SpillageFlag = '1' And CMPID = " & mCMPID & " And GodownID = " & mGodownID, "D")
ElseIf LCase(CmbTxnType.Text) = "made up bags" Then
   Call TableMst_GSL(" Where  Status Not in  ('B','Z','E') And sm_prefix ='" & TxtPreFix & "' And SpillageFlag = '0' And MadeUpFlag = '0' And CMPID = " & mCMPID & " And GodownID = " & mGodownID, "D")
End If
CmbStack.Clear
If RsMst_GSL.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_GSL.RecordCount
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
   MsgBox "Invalid selection "
   CmbStack.SetFocus
   Exit Sub
End If
If mCmbStack <> CmbStack.Text Then
   CmbLot.Text = ""
   
   mCmbLot = ""
End If
mCmbStack = CmbStack.Text
End Sub
Private Sub CmbLot_GotFocus()

If CmbStack.Text = "" Then
  MsgBox "Please Select Stack No !!!"
  CmbStack.SetFocus
  Exit Sub
End If

tmp = CmbLot.Text

If LCase(CmbTxnType.Text) = "spillage" Then
   Call TableMst_GSL(" Where  sm_prefix ='" & TxtPreFix & "' And SpillageFlag = '1' And CMPID = " & mCMPID & " And GodownID = " & mGodownID & " And StackNo = '" & CmbStack.Text & "' And Status not in ('Z','E')")
ElseIf LCase(CmbTxnType.Text) = "made up bags" Then
   Call TableMst_GSL(" Where  Status not in ('B','Z','E') and sm_prefix ='" & TxtPreFix & "' And SpillageFlag = '0' And MadeUpFlag = '0' And CMPID = " & mCMPID & " And GodownID = " & mGodownID & " And StackNo = '" & CmbStack.Text & "'")
End If

CmbLot.Clear
If RsMst_GSL.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_GSL.RecordCount
    CmbLot.AddItem RsMst_GSL!LotNo
    RsMst_GSL.MoveNext
Next I
CmbLot.Text = tmp
End Sub
'----
Private Sub CmbLot_LostFocus()
If CmbLot.Text = "" Then
   Exit Sub
End If
RsMst_GSL.MoveFirst
RsMst_GSL.Find "LotNo = '" & CmbLot.Text & "'"
If RsMst_GSL.EOF Then
   MsgBox "Invalid selection "
   CmbLot.SetFocus
   Exit Sub
End If
mGSLID = RsMst_GSL!Mst_GSL_ID

If LCase(CmbTxnType.Text) = "made up bags" Then
   
   tmp = " Select TCD.* from Txn_CxTF_Details TCD"
   tmp = tmp & "  inner join Txn_CxTF_GSL TCG on TCG.CxTF_DetailsID = TCD.CxTF_DetailsID And TCG.TxnType = TCD.TxnType And TCG.SM_Prefix = TCD.SM_Prefix"
   tmp = tmp & " where isnull(TCD.Flag,'') <> 'P' And TCD.TxnType = 'D' and TCG.Mst_GSL_ID = " & mGSLID & " and TCG.SM_prefix = '" & TxtPreFix & "'"
   
   Call Tmp5Table
   
   If TmpRs5.RecordCount > 0 Then
'      MsgBox "One of Deposit of this lot not printed. Can't generate Made up Lot!!!"
      MsgBox "One or more CDTFs for this lot not printed. Can't transfer!!!"
      CmbLot.SetFocus
      Exit Sub
   End If

End If

mCommodityID = IIf(IsNull(RsMst_GSL!CommodityID), 0, RsMst_GSL!CommodityID)
If mCommodityID = 0 Then
  'MsgBox "No Transaction Done Earlier on This GSL"
   CmbCommodityID.Visible = True
   TxtCommodity.Visible = False
   CmbCommodityID.Clear
   CmdCommodityID.Visible = True
   CmbCommodityID.SetFocus
   Exit Sub
Else
   If Not IsNull(RsMst_GSL!ExchangeTypeID) Then
      TxtExchangeTypeID.Text = GetExchangeName(RsMst_GSL!ExchangeTypeID)
   End If
    
   If Not IsNull(RsMst_GSL!ClientID) Then
      TxtClientID.Text = GetClientName(RsMst_GSL!ClientID)
   End If
   TxtCommodity = GetCommodityName(mCommodityID)
   TxtNoofBags = RsMst_GSL!DepositeBags
   TxtQuantity = RsMst_GSL!DepositeWeight
End If

mConnectedGSLID = IIf(IsNull(RsMst_GSL!Connected_Mst_GSL_ID), 0, RsMst_GSL!Connected_Mst_GSL_ID)
'TxtNewGSLID = mConnectedGSLID
If mConnectedGSLID <> 0 Then
   tmp = "Select b.godownno+'~'+a.stackNo+'~'+a.lotno as 'GSL',a.DepositeBags,a.balancebags from Mst_GSL a "
   tmp = tmp & " inner join mst_godown b On a.godownid = b.godownid Where a.Mst_GSL_ID = " & mConnectedGSLID & " And a.sm_prefix = '" & TxtPreFix & "'"
   
   Call TmpTable
   TxtNewGSLID = ""
   TxtMadeupBags = ""
   If TmpRs.RecordCount > 0 Then
      TxtNewGSLID = IIf(IsNull(TmpRs!GSL), "", TmpRs!GSL)
      TxtMadeupBags = IIf(IsNull(TmpRs!DepositeBags), 0, TmpRs!DepositeBags)
      TxtBalanceMadeUp = IIf(IsNull(TmpRs!BalanceBags), 0, TmpRs!BalanceBags)
   End If
   
End If

End Sub
Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Call GButtonLockAD(Me)
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If
Call Grid_Head
Call CmdClear_Click

Call TableTxn_Spillage

If RsTxn_Spillage.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
Call GButtonLock(Me, True)
Call Indata
End Sub
Private Sub Indata()

TxtSpillageID = IIf(IsNull(RsTxn_Spillage!Txn_Spillage_ID), "0", RsTxn_Spillage!Txn_Spillage_ID)
SpillageDate.Value = IIf(IsNull(RsTxn_Spillage!SpillageDate), "0", RsTxn_Spillage!SpillageDate)
TxtPreFix = IIf(IsNull(RsTxn_Spillage!SM_Prefix), "", RsTxn_Spillage!SM_Prefix)
LblNewGSL.Visible = False
TxtNewGSLID.Visible = False
TxtExchangeTypeID.Visible = False
TxtClientID.Visible = False
LblExchangeTypeID.Visible = False
LblClientID.Visible = False


mTxnType = IIf(IsNull(RsTxn_Spillage!Flag), "S", RsTxn_Spillage!Flag)

If mTxnType = "S" Then
   TxtTxnType = "Spillage"
Else
   TxtTxnType = "Made Up Bags"
   LblNewGSL.Visible = True
   TxtNewGSLID.Visible = True
   TxtExchangeTypeID.Visible = True
   TxtClientID.Visible = True
   LblExchangeTypeID.Visible = True
   LblClientID.Visible = True
End If

mGSLID = RsTxn_Spillage!Mst_GSL_ID
'Ans = GetGSLByGLSID(RsTxn_Spillage!GSLID)
Call TableMst_GSL("Where sm_prefix ='" & TxtPreFix & "' And Mst_GSL_ID = " & mGSLID)
mCMPID = RsMst_GSL!CMPID
mGodownID = RsMst_GSL!GodownID
TxtCMPID = GetCMPName(mCMPID)
TxtGodownID = GetGodownName(RsMst_GSL!GodownID, "N") 'IIf(IsNull(RsTxn_Spillage!SpillageID), "0", RsTxn_Spillage!SpillageID)
TxtStack = RsMst_GSL!StackNo
TxtLot = RsMst_GSL!LotNo 'Ans(5)
TxtGSLFlag = RsMst_GSL!Status   'Ans(5)

TxtCommodity = GetCommodityName(RsTxn_Spillage!CommodityID)
mCommodityID = IIf(IsNull(RsMst_GSL!CommodityID), RsTxn_Spillage!CommodityID, RsMst_GSL!CommodityID)

TxtExchangeTypeID = ""
TxtClientID = ""

If Not IsNull(RsMst_GSL!ExchangeTypeID) Then
   TxtExchangeTypeID.Text = GetExchangeName(RsMst_GSL!ExchangeTypeID)
End If
 
If Not IsNull(RsMst_GSL!ClientID) Then
   TxtClientID.Text = GetClientName(RsMst_GSL!ClientID)
End If

TxtNoofBags = IIf(IsNull(RsMst_GSL!DepositeBags), "0", RsMst_GSL!DepositeBags)
TxtQuantity = IIf(IsNull(RsMst_GSL!DepositeWeight), "0", RsMst_GSL!DepositeWeight) 'RsMst_GSL!DepositeWeight
TxtSpillageNoOfBags = IIf(IsNull(RsTxn_Spillage!SpillageNoOfBags), "0", RsTxn_Spillage!SpillageNoOfBags)
TxtSpillageQuantity = IIf(IsNull(RsTxn_Spillage!SpillageQuantity), "0", RsTxn_Spillage!SpillageQuantity)
TxtRemarks = IIf(IsNull(RsTxn_Spillage!Remarks), "", RsTxn_Spillage!Remarks)

TxtCreated = IIf(IsNull(RsTxn_Spillage!CreatedBy), "", RsTxn_Spillage!CreatedBy) & " :- " & IIf(IsNull(RsTxn_Spillage!CreatedDate), "", RsTxn_Spillage!CreatedDate)
TxtUpdated = IIf(IsNull(RsTxn_Spillage!UpdatedBy), "", RsTxn_Spillage!UpdatedBy) & " :- " & IIf(IsNull(RsTxn_Spillage!UpdatedDate), "", RsTxn_Spillage!UpdatedDate)
TxtNewGSLID = ""
mConnectedGSLID = IIf(IsNull(RsMst_GSL!Connected_Mst_GSL_ID), 0, RsMst_GSL!Connected_Mst_GSL_ID)
If mTxnType <> "S" Then
   mNewGSLID = IIf(IsNull(RsMst_GSL!Connected_Mst_GSL_ID), 0, RsMst_GSL!Connected_Mst_GSL_ID)
'   TxtNewGSLID = TxtGodownID & "~" & TxtStack & "~" & "MadeUp" & TxtLot

   tmp = "Select b.godownno+'~'+a.stackNo+'~'+a.lotno as 'GSL',a.DepositeBags,a.balancebags from Mst_GSL a "
   tmp = tmp & " inner join mst_godown b On a.godownid = b.godownid Where a.Mst_GSL_ID = " & mNewGSLID & " And a.sm_prefix = '" & TxtPreFix & "'"
   Call TmpTable
   
   TxtNewGSLID = ""
   TxtMadeupBags = ""
   If TmpRs.RecordCount > 0 Then
      TxtNewGSLID = IIf(IsNull(TmpRs!GSL), "", TmpRs!GSL)
      TxtMadeupBags = IIf(IsNull(TmpRs!DepositeBags), 0, TmpRs!DepositeBags)
      TxtBalanceMadeUp = IIf(IsNull(TmpRs!BalanceBags), 0, TmpRs!BalanceBags)
   End If
End If

TxtSpillageNoOfBags.Locked = True
TxtSpillageQuantity.Locked = True
TxtRemarks.Locked = True
CmbCMPID.Visible = False
CmbGodownNo.Visible = False
CmbStack.Visible = False
CmbLot.Visible = False
CmbCommodityID.Visible = False
CmdCommodityID.Visible = False
CmdGSL.Visible = False
TxtCMPID.Visible = True
TxtGodownID.Visible = True
TxtStack.Visible = True
TxtLot.Visible = True
TxtCommodity.Visible = True
SpillageDate.Enabled = False
CmbTxnType.Visible = False
TxtTxnType.Visible = True
CmbTxnType.Text = TxtTxnType
Call GButtonLock(Me, True)

End Sub

Private Sub SaveCmd_Click()

If SpillageDate.Value > Date Then
   MsgBox "Spillage date can't be future date!!!"
   SpillageDate.SetFocus
   Exit Sub
End If

If Val(TxtSpillageNoOfBags) < 0 Then
   MsgBox "Negative No of bags not allowed !!!"
   TxtSpillageNoOfBags.SetFocus
   Exit Sub
End If

If Val(TxtSpillageQuantity) < 0 Then
   MsgBox "Negative quantity not allowed !!!"
   TxtSpillageQuantity.SetFocus
   Exit Sub
End If

If Edit_Flg = "A" Then
   If CmbTxnType.Text = "" Then
      MsgBox "Blank Bags Type not allowed!!!"
      CmbTxnType.SetFocus
      Exit Sub
   End If
   If CmbCMPID.Text = "" Then
      MsgBox "Blank CMP not allowed!!!"
      CmbCMPID.SetFocus
      Exit Sub
   End If
   If CmbGodownNo.Text = "" Then
      MsgBox "Blank Godown No not allowed!!!"
      CmbGodownNo.SetFocus
      Exit Sub
   End If
   If CmbStack.Text = "" Then
      MsgBox "Blank Stack No not allowed!!!"
      CmbStack.SetFocus
      Exit Sub
   End If
   If CmbLot.Text = "" Then
      MsgBox "Blank Lot No not allowed!!!"
      CmbLot.SetFocus
      Exit Sub
   End If
   If CmbCommodityID.Visible = True Then
      If CmbCommodityID.Text = "" Then
         MsgBox "Blank Commodity not allowed!!!"
         CmbCommodityID.SetFocus
         Exit Sub
      End If
   End If
End If

If LCase(CmbTxnType.Text) = "spillage" Then
   If Val(TxtSpillageQuantity) <= 0 Then
      MsgBox "Please enter spillage weight!!!"
      TxtSpillageQuantity.SetFocus
      Exit Sub
   End If
Else
   If Val(TxtSpillageNoOfBags) <= 0 Then
      MsgBox "Please enter Made Up Bags!!!"
      TxtSpillageNoOfBags.SetFocus
      Exit Sub
   End If

   If Val(TxtBalanceMadeUp) < 0 Then
      MsgBox "-ve balance of Made Up Bags not allowed!!!"
      TxtSpillageNoOfBags.SetFocus
      Exit Sub
   End If

End If

If TxtRemarks = "" Then
   MsgBox "Blank Remarks not allowed!!!"
   TxtRemarks.SetFocus
   Exit Sub
End If

If Edit_Flg = "A" Then
  RsTxn_Spillage.AddNew
  RsTxn_Spillage!SpillageID = GetSpillageID
  RsTxn_Spillage!G_UID = GetGUID
  TxtSpillageID = RsTxn_Spillage!SpillageID
Else
  RsTxn_Spillage.Find "Txn_Spillage_ID = " & TxtSpillageID
End If

Call TableMst_GSL(" Where sm_prefix ='" & TxtPreFix & "'")
RsMst_GSL.Find "Mst_GSL_ID = " & mGSLID
If Not RsMst_GSL.EOF Then
   If LCase(CmbTxnType.Text) = "spillage" Then
      RsMst_GSL!CommodityID = mCommodityID
      If Edit_Flg = "E" Then
         If (Val(IIf(IsNull(RsMst_GSL!BalanceWeight), "0", RsMst_GSL!BalanceWeight)) - Val(mSpillageQuantity)) + (Val(TxtSpillageQuantity)) < 0 Then
            MsgBox "You can't delete this record since some quantity is withdrawn from this lot !!! "
            Exit Sub
         End If
         If (Val(RsMst_GSL!BalanceBags) - Val(mSpillageNoOfBags)) + Val(TxtSpillageNoOfBags) < 0 Then
            MsgBox "You Can't delete this record since some bags are withdrawn from this lot !!! "
            Exit Sub
         End If
         RsMst_GSL!DepositeBags = Val(IIf(IsNull(RsMst_GSL!DepositeBags), "0", RsMst_GSL!DepositeBags)) - (Val(mSpillageNoOfBags))
         RsMst_GSL!DepositeWeight = Val(IIf(IsNull(RsMst_GSL!DepositeWeight), "0", RsMst_GSL!DepositeWeight)) - (Val(mSpillageQuantity))
         RsMst_GSL!BalanceBags = Val(IIf(IsNull(RsMst_GSL!BalanceBags), "0", RsMst_GSL!BalanceBags)) - (Val(mSpillageNoOfBags))
         RsMst_GSL!BalanceWeight = Val(IIf(IsNull(RsMst_GSL!BalanceWeight), "0", RsMst_GSL!BalanceWeight)) - (Val(mSpillageQuantity))
      End If
      RsMst_GSL!DepositeBags = Val(IIf(IsNull(RsMst_GSL!DepositeBags), "0", RsMst_GSL!DepositeBags)) + (Val(TxtSpillageNoOfBags))
      RsMst_GSL!DepositeWeight = Val(IIf(IsNull(RsMst_GSL!DepositeWeight), "0", RsMst_GSL!DepositeWeight)) + (Val(TxtSpillageQuantity))
      RsMst_GSL!BalanceBags = Val(IIf(IsNull(RsMst_GSL!BalanceBags), "0", RsMst_GSL!BalanceBags)) + (Val(TxtSpillageNoOfBags))
      RsMst_GSL!BalanceWeight = Val(IIf(IsNull(RsMst_GSL!BalanceWeight), "0", RsMst_GSL!BalanceWeight)) + (Val(TxtSpillageQuantity))
      RsMst_GSL!Status = "O"
      RsMst_GSL.Update
   Else
'      If Edit_Flg = "E" Then
'         If (Val(RsMst_GSL!BalanceBags) - Val(mSpillageNoOfBags)) + Val(TxtSpillageNoOfBags) < 0 Then
'            MsgBox "You Can't delete this record since some bags are withdrawn from this lot !!! "
'            Exit Sub
'         End If
'         RsMst_GSL!BalanceBags = Val(IIf(IsNull(RsMst_GSL!BalanceBags), "0", RsMst_GSL!BalanceBags)) - (Val(mSpillageNoOfBags))
'      End If
'      RsMst_GSL!MadeUpBags = Val(IIf(IsNull(RsMst_GSL!MadeUpBags), "0", RsMst_GSL!MadeUpBags)) - (Val(TxtSpillageNoOfBags))
'      RsMst_GSL!BalanceBags = Val(IIf(IsNull(RsMst_GSL!BalanceBags), "0", RsMst_GSL!BalanceBags)) + (Val(TxtSpillageNoOfBags))
'      RsMst_GSL.Update
   
      If mConnectedGSLID = 0 Then
         'mConnectedGSLID = GetMaxGSLID
'         mExchangeTypeID = RsMst_GSL!ExchangeTypeID
'         mClientIDRow = RsMst_GSL!ClientIDRow
'         mCommodityID = RsMst_GSL!CommodityID
'         mSCMFlag = RsMst_GSL!SCMFlagID
         
         If (Val(TxtSpillageNoOfBags) + Val(TxtMadeupBags)) > Format((Val(TxtNoofBags) * (20 / 100)), "#######0") Then
            MsgBox "Made up Bags are " & (Val(TxtSpillageNoOfBags) + Val(TxtMadeupBags)) & " which is more than 20% of deposited bags " & TxtNoofBags & " . Can't create Made up lot!!!"
            If Edit_Flg = "A" Then
               RsTxn_Spillage.CancelUpdate
            End If
            Exit Sub
         End If
         Call CreateMadeUpGSL
         RsMst_GSL!Connected_Mst_GSL_ID = mNewGSLID
         RsMst_GSL.Update
         tmp = "Select * from Mst_GSL Where MadeupFlag = '1' And Mst_GSL_ID = " & mNewGSLID & " And SM_Prefix = '" & TxtPreFix & "'"
         Call TmpTable
      
      Else
         
         tmp = "Select * from Mst_GSL Where MadeupFlag = '1' And Mst_GSL_ID = " & mConnectedGSLID & " And SM_Prefix = '" & TxtPreFix & "' "
         Call TmpTable
         
         If Edit_Flg = "A" Then
            If Val(TxtSpillageNoOfBags) > Format((Val(TxtNoofBags) * (20 / 100)), "#######0") - Val(TxtMadeupBags) Then
               MsgBox "Made up Bags will be " & Val(TxtSpillageNoOfBags) + Val(TxtMadeupBags) & " after this deposit which is more than 20% of deposited bags " & TxtNoofBags & " . Can't update Made up lot!!!"
               RsTxn_Spillage.CancelUpdate
               Exit Sub
            End If
         Else
            If Val(TxtSpillageNoOfBags) > Format((Val(TxtNoofBags) * (20 / 100)), "#######0") - Val(TxtMadeupBags) Then
               MsgBox "Made up Bags will be " & Val(TxtSpillageNoOfBags) + Val(TxtMadeupBags) & " after this deposit which is more than 20% of deposited bags " & TxtNoofBags & " . Can't update Made up lot!!!"
               RsTxn_Spillage.CancelUpdate
               Exit Sub
            End If
         End If
      End If
      If Edit_Flg = "E" Then
'         If (Val(IIf(IsNull(RsMst_GSL!BalanceWeight), "0", RsMst_GSL!BalanceWeight)) - Val(mSpillageQuantity)) + (Val(TxtSpillageQuantity)) < 0 Then
'            MsgBox "You can't delete this record since some quantity is withdrawn from this lot !!! "
'            Exit Sub
'         End If
'         If (Val(RsMst_GSL!BalanceBags) - Val(mSpillageNoOfBags)) + Val(TxtSpillageNoOfBags) < 0 Then
'            MsgBox "You Can't delete this record since some bags are withdrawn from this lot !!! "
'            Exit Sub
'         End If
         TmpRs!DepositeBags = Val(IIf(IsNull(TmpRs!DepositeBags), "0", TmpRs!DepositeBags)) - (Val(mSpillageNoOfBags))
         TmpRs!DepositeWeight = Val(IIf(IsNull(TmpRs!DepositeWeight), "0", TmpRs!DepositeWeight)) - (Val(mSpillageQuantity))
         TmpRs!BalanceBags = Val(IIf(IsNull(TmpRs!BalanceBags), "0", TmpRs!BalanceBags)) - (Val(mSpillageNoOfBags))
         TmpRs!BalanceWeight = Val(IIf(IsNull(TmpRs!BalanceWeight), "0", TmpRs!BalanceWeight)) - (Val(mSpillageQuantity))
      End If
      TmpRs!DepositeBags = Val(IIf(IsNull(TmpRs!DepositeBags), "0", TmpRs!DepositeBags)) + (Val(TxtSpillageNoOfBags))
      TmpRs!DepositeWeight = Val(IIf(IsNull(TmpRs!DepositeWeight), "0", TmpRs!DepositeWeight)) + (Val(TxtSpillageQuantity))
      TmpRs!BalanceBags = Val(IIf(IsNull(TmpRs!BalanceBags), "0", TmpRs!BalanceBags)) + (Val(TxtSpillageNoOfBags))
      TmpRs!BalanceWeight = Val(IIf(IsNull(TmpRs!BalanceWeight), "0", TmpRs!BalanceWeight)) + (Val(TxtSpillageQuantity))
      If TmpRs!CDTFBags < TmpRs!DepositeBags Then
         TmpRs!Status = "O"
      End If
      TmpRs.Update
   End If
End If

RsTxn_Spillage!SpillageDate = SpillageDate.Value
RsTxn_Spillage!Flag = Left(CmbTxnType.Text, 1)
RsTxn_Spillage!Mst_GSL_ID = mGSLID
If LCase(RsTxn_Spillage!Flag) = "m" Then
   If mConnectedGSLID = 0 Then
      RsTxn_Spillage!NewMst_GSL_ID = mNewGSLID
   Else
      RsTxn_Spillage!NewMst_GSL_ID = mConnectedGSLID
   End If
End If
RsTxn_Spillage!CommodityID = mCommodityID
RsTxn_Spillage!SpillageNoOfBags = (Val(TxtSpillageNoOfBags))
RsTxn_Spillage!SpillageQuantity = (Val(TxtSpillageQuantity))
RsTxn_Spillage!Remarks = TxtRemarks.Text
RsTxn_Spillage!SM_Prefix = TxtPreFix

If IsNull(RsTxn_Spillage!CreatedBy) = True Or RsTxn_Spillage!CreatedBy = "" Then
   RsTxn_Spillage!CreatedBy = Gen_UserLoginID
   RsTxn_Spillage!CreatedDate = Now
Else
   RsTxn_Spillage!UpdatedBy = Gen_UserLoginID
   RsTxn_Spillage!UpdatedDate = Now
End If

RsTxn_Spillage.Update

Call Indata

End Sub

Private Function GetSpillageID() As Double
Dim Retval As Double
tmp = "Select max(SpillageID) from Txn_Spillage where sm_prefix ='" & TxtPreFix & "'"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetSpillageID = Retval
End Function
Private Sub CmbCommodityID_GotFocus()

If CmbCMPID.Visible = True Then
    If CmbCMPID.Text = "" Then
       MsgBox "Select CMP!!!"
       CmbCMPID.SetFocus
       Exit Sub
    End If
End If
'tmp = CmbCommodityID.Text
'Call TableMst_Commodity

tmp1 = CmbCommodityID.Text
'Call TableMst_Commodity

tmp = "Select MCom.CommodityID, MCom.Commodity From Mst_Commodity MCom "
tmp = tmp & " Inner Join Mst_LocationCOmmodityMapping MLCM on MLCM.CommodityID = MCom.CommodityID "
tmp = tmp & " Inner Join Mst_Location ML on ML.LocationID = MLCM.LocationID "
tmp = tmp & " Where ML.SM_Prefix = '" & TxtPreFix & "'"

Call Tmp7Table

CmbCommodityID.Clear
If TmpRs7.RecordCount = 0 Then
   MsgBox "No record found !!!! "
   Exit Sub
End If
For I = 1 To TmpRs7.RecordCount
    CmbCommodityID.AddItem TmpRs7!Commodity
    TmpRs7.MoveNext
Next
CmbCommodityID.Text = tmp1

End Sub
'----
Private Sub CmbCommodityID_LostFocus()
If CmbCommodityID.Text = "" Then
   Exit Sub
End If

TmpRs7.MoveFirst
TmpRs7.Find "Commodity = '" & CmbCommodityID.Text & "'"
If TmpRs7.EOF Then
   MsgBox "Invalid selection "
   CmbCommodityID.SetFocus
   Exit Sub
End If
mCommodityID = TmpRs7!CommodityID
End Sub
Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsTxn_Spillage.MoveNext
If RsTxn_Spillage.EOF Then
   RsTxn_Spillage.MoveLast
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
RsTxn_Spillage.MovePrevious
If RsTxn_Spillage.BOF Then
   RsTxn_Spillage.MoveFirst
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
RsTxn_Spillage.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsTxn_Spillage.MoveLast
Call Indata
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False
End Sub

Private Sub SearchCmd_Click()
Dim ans As Variant
If SearchCmd.Caption = "Cancel" Then
   If RsTxn_Spillage.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   'Call FirstCmd_Click
   Call Indata
   Exit Sub
End If
Frame0.Height = Me.Height - 700
Frame0.Left = Picture1.Left - 100
Frame0.Top = Picture1.Top - 100
Frame0.Width = Me.Width - 150
Frame0.Visible = True
MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 1400
RsTxn_Spillage.MoveFirst

''===========comment on 7 july 2012 for not clearing search grid=============
'MSHFlexGrid1.Clear
'MSHFlexGrid1.Cols = 3
'Call Grid_Head
''===========comment on 7 july 2012 for not clearing search grid=============

'Call CmdClear_Click

End Sub
Private Sub MSHFlexGrid1_Click()
Dim I As Variant
I = MSHFlexGrid1.RowSel
If I <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(I, 0) = "" Then Exit Sub
RsTxn_Spillage.MoveFirst
RsTxn_Spillage.Find "G_UID = '" & MSHFlexGrid1.TextMatrix(I, 11) & "' "
If Not RsTxn_Spillage.EOF Then
   Call Indata
   Frame0.Visible = False
End If
End Sub
Private Sub Grid_Head()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 13
MSHFlexGrid1.TextMatrix(0, 0) = "Spillage ID"
MSHFlexGrid1.TextMatrix(0, 1) = "SM_Prefix"
MSHFlexGrid1.TextMatrix(0, 2) = "Spillage Date"
MSHFlexGrid1.TextMatrix(0, 3) = "CMP"
MSHFlexGrid1.TextMatrix(0, 4) = "Godown No"
MSHFlexGrid1.TextMatrix(0, 5) = "Stack No"
MSHFlexGrid1.TextMatrix(0, 6) = "Lot No"
MSHFlexGrid1.TextMatrix(0, 7) = "Commodity"
MSHFlexGrid1.TextMatrix(0, 8) = "Spillage No Of Bags"
MSHFlexGrid1.TextMatrix(0, 9) = "Spillage Quantity"
MSHFlexGrid1.TextMatrix(0, 10) = "Remarks"
MSHFlexGrid1.TextMatrix(0, 11) = "G_UID"
MSHFlexGrid1.TextMatrix(0, 12) = "Bags Type"

MSHFlexGrid1.ColWidth(0) = 1000
MSHFlexGrid1.ColWidth(1) = 1000
MSHFlexGrid1.ColWidth(2) = 1100
MSHFlexGrid1.ColWidth(3) = 1800
MSHFlexGrid1.ColWidth(4) = 1000
MSHFlexGrid1.ColWidth(5) = 800
MSHFlexGrid1.ColWidth(6) = 800
MSHFlexGrid1.ColWidth(7) = 1200
MSHFlexGrid1.ColWidth(8) = 1600
MSHFlexGrid1.ColWidth(9) = 1400
MSHFlexGrid1.ColWidth(10) = 3000
MSHFlexGrid1.ColWidth(11) = 0
MSHFlexGrid1.ColWidth(12) = 2000

MSHFlexGrid1.FixedRows = 1
MSHFlexGrid1.FixedCols = 1
End Sub
Private Sub TxtSpillageNoOfBags_LostFocus()
If TxtSpillageNoOfBags = "" Then Exit Sub
If IsNumeric(TxtSpillageNoOfBags) = False Then
   MsgBox "Invalid numeric format!!!"
   TxtSpillageNoOfBags.SetFocus
   Exit Sub
End If
If Val(TxtSpillageNoOfBags) < 0 Then
   MsgBox "Negative No of bags not allowed !!!"
   TxtSpillageNoOfBags.SetFocus
   Exit Sub
End If

TxtSpillageNoOfBags = Round(TxtSpillageNoOfBags, 0)
End Sub
Private Sub TxtSpillageQuantity_LostFocus()
If TxtSpillageQuantity = "" Then Exit Sub
If IsNumeric(TxtSpillageQuantity) = False Then
   MsgBox "Invalid numeric format!!!"
   TxtSpillageQuantity.SetFocus
   Exit Sub
End If
If Val(TxtSpillageQuantity) < 0 Then
   MsgBox "Negative quantity not allowed !!!"
   TxtSpillageQuantity.SetFocus
   Exit Sub
End If

TxtSpillageQuantity = Format(TxtSpillageQuantity, "#######0.000")

End Sub

Private Sub CmbSCmp_GotFocus()
If CmbLocation.Text = "" Then
   MsgBox "Select Location first !!! "
   CmbLocation.SetFocus
   Exit Sub
End If
tmp = CmbSCmp.Text
Call TableMst_CMP(" where LocationID=" & mSLocationID)
CmbSCmp.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
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
   mSCMPID = 0
   Exit Sub
End If
RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPName = '" & CmbSCmp.Text & "'"
If RsMst_CMP.EOF Then
   MsgBox "Invalid selection "
   CmbSCmp.SetFocus
   Exit Sub
End If
mSCMPID = RsMst_CMP!CMPID
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

Private Sub CmbSGodown_GotFocus()
If CmbSCmp.Text = "" Then
   MsgBox "Select CMP first !!! "
   CmbSCmp.SetFocus
   Exit Sub
End If
tmp = CmbSGodown.Text
Call TableMst_Godown("Where CMPID = " & mSCMPID)
CmbSGodown.Clear
If RsMst_Godown.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Godown.RecordCount
    CmbSGodown.AddItem RsMst_Godown!GodownNo
    RsMst_Godown.MoveNext
Next
CmbSGodown.Text = tmp
End Sub

Private Sub CmbSGodown_LostFocus()
If CmbSGodown.Text = "" Then
   mSGodownID = 0
   Exit Sub
End If
RsMst_Godown.MoveFirst
RsMst_Godown.Find "GodownNo = '" & CmbSGodown.Text & "'"
If RsMst_Godown.EOF Then
   MsgBox "Invalid selection "
   CmbSGodown.SetFocus
   Exit Sub
End If
mSGodownID = RsMst_Godown!GodownID
End Sub

Private Sub CmbSStack_GotFocus()

If CmbSGodown.Text = "" Or mSGodownID = "" Then
  MsgBox "Please Select Godown No!!!"
  Exit Sub
End If
tmp = CmbSStack.Text
Call TableMst_GSL(" Where CMPID = " & mSCMPID & " And GodownID = " & mSGodownID, "D")
CmbSStack.Clear
If RsMst_GSL.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_GSL.RecordCount
    CmbSStack.AddItem RsMst_GSL!StackNo
    RsMst_GSL.MoveNext
Next

CmbSStack.Text = tmp
End Sub
'----
Private Sub CmbSStack_LostFocus()
If CmbSStack.Text = "" Then
   Exit Sub
End If
RsMst_GSL.MoveFirst
RsMst_GSL.Find "StackNo = '" & CmbSStack.Text & "'"
If RsMst_GSL.EOF Then
   MsgBox "Invalid selection "
   CmbSStack.SetFocus
   Exit Sub
End If

End Sub

Private Function GetMaxGSLID() As Double
Dim Retval As Double
tmp = "Select max(GSLID) from Mst_GSL Where SM_Prefix = '" & TxtPreFix & "'"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetMaxGSLID = Retval
End Function

Private Sub CreateMadeUpGSL()

tmp = "Select * from Mst_GSL Where SM_Prefix = '" & TxtPreFix & "'"
Call Tmp1Table

TmpRs1.AddNew
'TmpRs1!GSLID = GetMaxGSLID
TmpRs1!G_UID = GetGUID
mNewG_UID = TmpRs1!G_UID
'mNewGSLID = TmpRs1!GSLID
TmpRs1!Connected_Mst_GSL_ID = mGSLID
TmpRs1!SM_Prefix = TxtPreFix
TmpRs1!CMPID = mCMPID
TmpRs1!GodownID = mGodownID
TmpRs1!StackNo = CmbStack.Text
TmpRs1!LotNo = "MUB" & CmbLot.Text
TmpRs1!ExchangeTypeID = RsMst_GSL!ExchangeTypeID
TmpRs1!ClientID = RsMst_GSL!ClientID
TmpRs1!CommodityID = RsMst_GSL!CommodityID 'mCommodityID
TmpRs1!SpillageFlag = RsMst_GSL!SpillageFlag
TmpRs1!MadeupFlag = 1
TmpRs1!Status = "O"
TmpRs1!SCMFlag = RsMst_GSL!SCMFlag
TmpRs1!PledgeNo = RsMst_GSL!PledgeNo
'TmpRs1!ISINID = RsMst_GSL!ISINID
'TmpRs1!CdtfNo = RsMst_GSL!CdtfNo
TmpRs1!CDTFDate = RsMst_GSL!CDTFDate
'TmpRs1!FirstISINID = RsMst_GSL!FirstISINID
TmpRs1!SpillageFlag = RsMst_GSL!SpillageFlag
'TmpRs1!DepositAdjWt = RsMst_GSL!DepositAdjWt
'TmpRs1!WithdrawalAdjWt = RsMst_GSL!WithdrawalAdjWt
TmpRs1!LCDTF_Date = RsMst_GSL!LCDTF_Date
'TmpRs1!GrossWt = RsMst_GSL!GrossWt
'TmpRs1!KRNo = RsMst_GSL!KRNo
'TmpRs1!KRNoBatchID = RsMst_GSL!KRNoBatchID
'TmpRs1!KRDate = RsMst_GSL!KRDate
TmpRs1!ContractID = RsMst_GSL!ContractID
TmpRs1!BagSize = RsMst_GSL!BagSize
TmpRs1!BagEquivalent = RsMst_GSL!BagEquivalent
'TmpRs1!WtQty = RsMst_GSL!WtQty
TmpRs1!CRAmount = RsMst_GSL!CRAmount
TmpRs1!CRDate = RsMst_GSL!CRDate
TmpRs1!NContractID = RsMst_GSL!NContractID
TmpRs1!RateCardType = RsMst_GSL!RateCardType
TmpRs1!BillingCycleID = RsMst_GSL!BillingCycleID
TmpRs1!BillingUnit = RsMst_GSL!BillingUnit
TmpRs1!Rate = RsMst_GSL!Rate
TmpRs1!RateForDailyPnL = RsMst_GSL!RateForDailyPnL
TmpRs1!DepConDID = RsMst_GSL!DepConDID
TmpRs1!NoOfDays = RsMst_GSL!NoOfDays
TmpRs1!State = "A"
TmpRs1!GSLID = RsMst_GSL!GSLID



If IsNull(TmpRs1!CreatedBy) = True Or TmpRs1!CreatedBy = "" Then
   TmpRs1!CreatedBy = Gen_UserLoginID
   TmpRs1!CreatedDate = Now
Else
   TmpRs1!UpdatedBy = Gen_UserLoginID
   TmpRs1!UpdatedDate = Now
End If

TmpRs1.Update

tmp = "Select Mst_GSL_ID From Mst_GSL Where G_UID='" & mNewG_UID & "'"
Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
'   mCxTF_DetailsID = TmpRs!CxTF_DetailsID
   mNewGSLID = TmpRs1!Mst_GSL_ID
End If




End Sub

