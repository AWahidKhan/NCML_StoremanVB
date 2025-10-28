VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm_Pledge_Creation 
   Caption         =   "Pledge Creation"
   ClientHeight    =   9690
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14370
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9690
   ScaleWidth      =   14370
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   1245
      Left            =   14520
      TabIndex        =   46
      Top             =   0
      Visible         =   0   'False
      Width           =   735
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   350
         Left            =   80
         TabIndex        =   47
         Top             =   150
         Width           =   1000
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1455
         Left            =   120
         TabIndex        =   48
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
      Height          =   975
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   14205
      TabIndex        =   32
      Top             =   8520
      Width           =   14265
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   375
         Left            =   11070
         TabIndex        =   17
         Top             =   495
         Width           =   2655
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   375
         Left            =   8430
         TabIndex        =   38
         Top             =   495
         Width           =   2640
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   375
         Left            =   5790
         TabIndex        =   37
         Top             =   495
         Width           =   2640
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   375
         Left            =   3135
         TabIndex        =   36
         Top             =   495
         Width           =   2655
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   375
         Left            =   480
         TabIndex        =   35
         Top             =   495
         Width           =   2655
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "&Search"
         Height          =   375
         Left            =   11070
         TabIndex        =   18
         Top             =   120
         Width           =   2655
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   8430
         TabIndex        =   34
         Top             =   120
         Width           =   2640
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   5790
         TabIndex        =   33
         Top             =   120
         Width           =   2640
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   375
         Left            =   3135
         TabIndex        =   16
         Top             =   120
         Width           =   2655
      End
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New "
         Height          =   375
         Left            =   480
         TabIndex        =   0
         Top             =   120
         Width           =   2655
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   8445
      Left            =   0
      ScaleHeight     =   8385
      ScaleWidth      =   14205
      TabIndex        =   19
      Top             =   0
      Width           =   14265
      Begin VB.Frame Frame3 
         Height          =   1635
         Left            =   120
         TabIndex        =   69
         Top             =   2880
         Width           =   13815
         Begin VB.TextBox TxtPledgeFee 
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
            Left            =   11400
            TabIndex        =   11
            Top             =   705
            Width           =   2040
         End
         Begin VB.TextBox TxtPledgeReqNo 
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
            TabIndex        =   9
            Top             =   1095
            Width           =   5025
         End
         Begin VB.ComboBox CmbBankName 
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
            Sorted          =   -1  'True
            TabIndex        =   7
            Top             =   240
            Width           =   5025
         End
         Begin VB.ComboBox CmbBranchName 
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
            Sorted          =   -1  'True
            TabIndex        =   8
            Top             =   660
            Width           =   5025
         End
         Begin VB.CommandButton CmdBankName 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   7560
            TabIndex        =   71
            Top             =   255
            Width           =   375
         End
         Begin VB.CommandButton CmdBranchName 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   7560
            TabIndex        =   70
            Top             =   660
            Width           =   375
         End
         Begin MSComCtl2.DTPicker PledgeRequestDate 
            Height          =   375
            Left            =   11400
            TabIndex        =   10
            Top             =   240
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   661
            _Version        =   393216
            Format          =   20709377
            CurrentDate     =   39884
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
            Left            =   2400
            TabIndex        =   72
            TabStop         =   0   'False
            Top             =   240
            Width           =   5010
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
            Height          =   375
            Left            =   2400
            TabIndex        =   73
            TabStop         =   0   'False
            Top             =   660
            Width           =   5025
         End
         Begin VB.Label Label11 
            Caption         =   "Pledge Creation Fee"
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
            Left            =   8280
            TabIndex        =   78
            Top             =   750
            Width           =   2535
         End
         Begin VB.Label Label5 
            Caption         =   "Pledge Request No "
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
            TabIndex        =   77
            Top             =   1125
            Width           =   2775
         End
         Begin VB.Label Label38 
            Caption         =   "Bank Name"
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
            TabIndex        =   76
            Top             =   285
            Width           =   1455
         End
         Begin VB.Label Label39 
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
            Height          =   285
            Left            =   120
            TabIndex        =   75
            Top             =   705
            Width           =   1815
         End
         Begin VB.Label Label18 
            Caption         =   "Pledge Request Date"
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
            TabIndex        =   74
            Top             =   300
            Width           =   2895
         End
      End
      Begin VB.PictureBox Picture2 
         Height          =   2745
         Left            =   120
         ScaleHeight     =   2685
         ScaleWidth      =   13920
         TabIndex        =   20
         Top             =   120
         Width           =   13980
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
            Left            =   1920
            TabIndex        =   3
            Top             =   975
            Width           =   2670
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
            Height          =   375
            Left            =   10155
            Locked          =   -1  'True
            TabIndex        =   55
            TabStop         =   0   'False
            Top             =   45
            Width           =   1515
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
            Left            =   6120
            Locked          =   -1  'True
            TabIndex        =   53
            TabStop         =   0   'False
            Top             =   1800
            Width           =   1575
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
            TabIndex        =   52
            TabStop         =   0   'False
            Top             =   1800
            Width           =   6135
         End
         Begin VB.ComboBox CmbClientDP 
            Height          =   315
            Left            =   6120
            Sorted          =   -1  'True
            TabIndex        =   5
            Top             =   960
            Visible         =   0   'False
            Width           =   7215
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
            Left            =   1905
            Locked          =   -1  'True
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   1800
            Visible         =   0   'False
            Width           =   2670
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
            ItemData        =   "Frm_Pledge_Creation.frx":0000
            Left            =   1920
            List            =   "Frm_Pledge_Creation.frx":0002
            Sorted          =   -1  'True
            TabIndex        =   1
            Top             =   120
            Width           =   2655
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
            Height          =   435
            Left            =   1920
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   6
            Top             =   2175
            Width           =   11940
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
            Left            =   12315
            Locked          =   -1  'True
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   1400
            Width           =   1515
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
            Left            =   9585
            Locked          =   -1  'True
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   1400
            Width           =   1605
         End
         Begin VB.TextBox TxtLicenseNo 
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
            Left            =   13560
            Locked          =   -1  'True
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   1395
            Visible         =   0   'False
            Width           =   525
         End
         Begin VB.TextBox TxtPledgeID 
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
            Left            =   6105
            Locked          =   -1  'True
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   45
            Width           =   2700
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
            Left            =   6120
            Locked          =   -1  'True
            TabIndex        =   4
            Top             =   500
            Width           =   2670
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
            Height          =   375
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   1395
            Width           =   2670
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
            Left            =   6120
            Locked          =   -1  'True
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   960
            Width           =   7710
         End
         Begin MSComCtl2.DTPicker PledgeDate 
            Height          =   375
            Left            =   1905
            TabIndex        =   2
            Top             =   525
            Width           =   2685
            _ExtentX        =   4736
            _ExtentY        =   661
            _Version        =   393216
            Format          =   20709377
            CurrentDate     =   39884
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
            Left            =   1560
            TabIndex        =   15
            Top             =   975
            Visible         =   0   'False
            Width           =   510
         End
         Begin MSComCtl2.DTPicker WHRDate 
            Height          =   375
            Left            =   6120
            TabIndex        =   39
            Top             =   1400
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   661
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   20709377
            CurrentDate     =   39884
         End
         Begin VB.Label Label31 
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
            Left            =   9000
            TabIndex        =   56
            Top             =   60
            Width           =   1695
         End
         Begin VB.Label lblDpIDName 
            Caption         =   "DP ID : Name"
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
            Left            =   4800
            TabIndex        =   54
            Top             =   1860
            Width           =   1335
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
            Left            =   120
            TabIndex        =   51
            Top             =   1860
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label Label20 
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
            Left            =   120
            TabIndex        =   49
            Top             =   150
            Width           =   1680
         End
         Begin VB.Label Label19 
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
            Left            =   135
            TabIndex        =   45
            Top             =   2265
            Width           =   1215
         End
         Begin VB.Label Label17 
            Caption         =   "WHR Qty."
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
            Left            =   11250
            TabIndex        =   44
            Top             =   1460
            Width           =   945
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
            Left            =   7980
            TabIndex        =   42
            Top             =   1460
            Width           =   1590
         End
         Begin VB.Label Label14 
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
            Left            =   4800
            TabIndex        =   40
            Top             =   1460
            Width           =   1215
         End
         Begin VB.Label Label2 
            Caption         =   "License No"
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
            Left            =   12240
            TabIndex        =   31
            Top             =   1455
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label Label9 
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
            Height          =   255
            Left            =   4800
            TabIndex        =   30
            Top             =   90
            Width           =   1095
         End
         Begin VB.Label Label6 
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
            Left            =   4800
            TabIndex        =   29
            Top             =   660
            Width           =   1590
         End
         Begin VB.Label Label26 
            Caption         =   "Pledge  Date"
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
            TabIndex        =   28
            Top             =   585
            Width           =   1215
         End
         Begin VB.Label Label7 
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
            Left            =   135
            TabIndex        =   27
            Top             =   1455
            Width           =   1455
         End
         Begin VB.Label Label8 
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
            Left            =   4800
            TabIndex        =   26
            Top             =   1020
            Width           =   1365
         End
         Begin VB.Label Label1 
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
            Left            =   120
            TabIndex        =   25
            Top             =   1035
            Width           =   1335
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
         Height          =   2745
         Left            =   120
         TabIndex        =   68
         Top             =   5520
         Width           =   13785
         _ExtentX        =   24315
         _ExtentY        =   4842
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   120
         TabIndex        =   57
         Top             =   4560
         Width           =   13770
         Begin VB.TextBox TxtUnpledgeWeight 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
            Height          =   360
            Left            =   7200
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   61
            TabStop         =   0   'False
            Top             =   420
            Width           =   1485
         End
         Begin VB.TextBox TxtUnpledgeNoOfBags 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
            Height          =   360
            Left            =   5475
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   60
            TabStop         =   0   'False
            Top             =   420
            Width           =   1785
         End
         Begin VB.TextBox TxtCMP 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
            Height          =   360
            Left            =   120
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   59
            TabStop         =   0   'False
            Top             =   420
            Width           =   3075
         End
         Begin VB.TextBox TxtGSL 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
            Height          =   360
            Left            =   3120
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   58
            TabStop         =   0   'False
            Top             =   420
            Width           =   2415
         End
         Begin VB.CommandButton CmdUpdate 
            Caption         =   "Update List"
            Height          =   360
            Left            =   12345
            TabIndex        =   14
            Top             =   420
            Width           =   1380
         End
         Begin VB.TextBox TxtPledgeNoOfBags 
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
            Left            =   8730
            TabIndex        =   12
            Top             =   420
            Width           =   1785
         End
         Begin VB.TextBox TxtPledgeQuantity 
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
            Left            =   10545
            TabIndex        =   13
            Top             =   420
            Width           =   1620
         End
         Begin VB.Label Label15 
            Caption         =   "Unpledged No Of Bags"
            Height          =   300
            Left            =   5520
            TabIndex        =   67
            Top             =   135
            Width           =   1710
         End
         Begin VB.Label Label10 
            Caption         =   "CMP"
            Height          =   300
            Left            =   1267
            TabIndex        =   66
            Top             =   165
            Width           =   1020
         End
         Begin VB.Label Label12 
            Caption         =   "G-S-L"
            Height          =   300
            Left            =   4050
            TabIndex        =   65
            Top             =   120
            Width           =   660
         End
         Begin VB.Label Label13 
            Caption         =   "Unpledged  Qty."
            Height          =   300
            Left            =   7560
            TabIndex        =   64
            Top             =   120
            Width           =   1260
         End
         Begin VB.Label Label3 
            Caption         =   "Pledged No Of Bags"
            Height          =   255
            Left            =   9000
            TabIndex        =   63
            Top             =   120
            Width           =   1575
         End
         Begin VB.Label Label4 
            Caption         =   "Pledged Qty."
            Height          =   255
            Left            =   10920
            TabIndex        =   62
            Top             =   120
            Width           =   1335
         End
      End
   End
End
Attribute VB_Name = "Frm_Pledge_Creation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg As Variant
Dim mClientIDRow As Double
Dim mCommodityId As Double
Dim mExchangeTypeID As Double
Dim mCMPID As Double
Dim mBankID As Double
Dim mBranchID As Double
Dim mWHRNo As Double
Dim mISINID As Variant
Dim mGodownID As Variant

Private Sub AddCmd_Click()
Edit_Flg = "A"
TxtWHRNo.Locked = False

CmbBankName.Visible = True
CmbBranchName.Visible = True
CmdBankName.Visible = True
CmdBranchName.Visible = True
TxtBankName.Visible = False
TxtBranchName.Visible = False
CmbExchange.Enabled = True
PledgeDate.Value = Date
PledgeDate.Enabled = True
PledgeRequestDate.Value = Date
CmdUpdate.Enabled = False
TxtPledgeNoOfBags.Locked = True
TxtPledgeQuantity.Locked = True
TxtPledgeReqNo.Locked = False
TxtPledgeFee.Locked = False
PledgeRequestDate.Enabled = True
TxtRemarks.Locked = False
CmbExchange.Text = ""
mExchangeTypeID = 0
TxtDPID.Visible = False
TxtDPName.Visible = False
TxtClientID.BackColor = -2147483648#
TxtPreFix = ""
TxtWHRNoNew = ""
Call ClearData
Call GButtonLock(Me, False)
Call TableTxn_WHRDetail
Call TableMst_ISIN
End Sub
Private Sub CmbBranchName_GotFocus()
If CmbBankName.Text = "" Then Exit Sub
If IsNull(mBankID) Then Exit Sub
tmp = CmbBranchName.Text
Call TableMst_Branch(" Where BankID =  " & mBankID & " Order by  BranchID")
CmbBranchName.Clear
For i = 1 To RsMst_Branch.RecordCount
    CmbBranchName.AddItem RsMst_Branch!BranchName
    RsMst_Branch.MoveNext
Next
CmbBranchName.Text = tmp
End Sub
Private Sub CmbBranchName_LostFocus()
If CmbBranchName.Text = "" Then
   Exit Sub
End If
RsMst_Branch.MoveFirst
RsMst_Branch.Find "BranchName = '" & CmbBranchName.Text & "'"
If RsMst_Branch.EOF Then
   MsgBox "Invalid  selection "
   CmbBranchName.SetFocus
   Exit Sub
End If
mBranchID = RsMst_Branch!BranchID
End Sub

Private Sub CmbExchange_GotFocus()
tmp = CmbExchange.Text
Call TableMst_ExchangeType
CmbExchange.Clear
If RsMst_ExchangeType.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_ExchangeType.RecordCount
    If RsMst_ExchangeType!ExchangeType <> "SPOT" Then
      CmbExchange.AddItem RsMst_ExchangeType!ExchangeType
    End If
    RsMst_ExchangeType.MoveNext
Next
CmbExchange.Text = tmp
TxtCommodity = ""
TxtLicenseNo = ""
End Sub

Private Sub CmbExchange_LostFocus()
If CmbExchange.Text = "" Then
   Exit Sub
End If
RsMst_ExchangeType.MoveFirst
RsMst_ExchangeType.Find "ExchangeType = '" & CmbExchange.Text & "'"
If RsMst_ExchangeType.EOF Then
   MsgBox "Invalid  selection "
   CmbExchange.SetFocus
   Exit Sub
End If
If CmbExchange.Text = "NCDEX" Then
  Call EnDn("NCDEX", False)
  mClientIDRow = 0
  TxtClientID.BackColor = -2147483643
  TxtClientID.Enabled = True
  TxtClientID.Locked = False
  lblDpIDName.Visible = True
  TxtDPID.Visible = True
  TxtDPName.Visible = True
Else
  Call EnDn("Non NCDEX", True)
  WHRDate.Enabled = False
  TxtClientID.BackColor = -2147483648#
  TxtClientID.Enabled = False
  TxtClientID.Locked = True
  
  lblDpIDName.Visible = False
  TxtDPID.Visible = False
  TxtDPName.Visible = False
End If
Call Grid_Head
TxtWHRNoNew.SetFocus
mExchangeTypeID = RsMst_ExchangeType!ExchangeTypeID
End Sub
Private Sub CmbBankName_GotFocus()
tmp = CmbBankName.Text
Call TableMst_Bank("")

CmbBankName.Clear
If RsMst_Bank.RecordCount = 0 Then
   MsgBox "No Bank found"
   'CmdInsuranceCo.SetFocus
   Exit Sub
End If
For i = 1 To RsMst_Bank.RecordCount
    CmbBankName.AddItem RsMst_Bank!BankName
    RsMst_Bank.MoveNext
Next
CmbBankName.Text = tmp
End Sub
Private Sub CmbBankName_LostFocus()
If CmbBankName.Text = "" Then
   Exit Sub
End If
RsMst_Bank.MoveFirst
RsMst_Bank.Find "BankName = '" & CmbBankName.Text & "'"
If RsMst_Bank.EOF Then
   MsgBox "Invalid Bank selection "
   CmbBankName.SetFocus
   Exit Sub
End If
mBankID = RsMst_Bank!BankID
If CmbBankName.Text <> tmp Then
   CmbBranchName.Text = ""
   TxtPledgeReqNo = ""
   TxtPledgeFee = ""
End If
End Sub

Private Sub CmbNonNcdexClient_Change()

End Sub

Private Sub CmdBankName_Click()
  FrmMst_Bank.Show
End Sub

Private Sub CmdBranchName_Click()
  FrmMst_Branch.Show
End Sub

Private Sub CmdUpdate_Click()
Dim strSplitString() As String
If TxtPledgeNoOfBags = "" Then
  MsgBox "Please Enter Pledge No Of Bags"
  TxtPledgeNoOfBags.SetFocus
  Exit Sub
End If
If TxtPledgeQuantity = "" Then
  MsgBox "Please Enter Pledge Quantity!!!"
  TxtPledgeQuantity.SetFocus
  Exit Sub
End If


If Val(TxtUnpledgeNoOfBags) = TxtPledgeNoOfBags Then
   If Val(TxtUnpledgeWeight) <> Val(TxtPledgeQuantity) Then
      MsgBox "When all the bags are pledged, all the quantity are also required to be pledged !!!! "
      TxtPledgeQuantity.SetFocus
      Exit Sub
   End If
End If

If Val(TxtUnpledgeWeight) = TxtPledgeQuantity Then
   If Val(TxtUnpledgeNoOfBags) <> Val(TxtPledgeNoOfBags) Then
      MsgBox "When all the quantity are pledged, all the  bags are also required to be pledged !!!! "
      TxtPledgeQuantity.SetFocus
      Exit Sub
   End If
End If

If TxtUnpledgeNoOfBags = 0 And TxtPledgeNoOfBags = 0 Then

Else
    If Val(TxtUnpledgeNoOfBags) <> TxtPledgeNoOfBags Then
       tmp = Validate_ComBags(mCommodityId, Val(TxtPledgeNoOfBags), Val(TxtPledgeQuantity))
       If Len(tmp) > 1 Then
          MsgBox tmp
          TxtPledgeQuantity.SetFocus
          Exit Sub
       End If
    End If
End If

MSHFlexGrid2.TextMatrix(0, 7) = Val(MSHFlexGrid2.TextMatrix(0, 7)) - Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 7))
MSHFlexGrid2.TextMatrix(0, 8) = Val(MSHFlexGrid2.TextMatrix(0, 8)) - Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 8))

If Val(TxtPledgeNoOfBags) = 0 And Val(TxtPledgeQuantity) = 0 Then
    MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 0) = ""
Else
    MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 0) = "*"
End If

MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 1) = TxtCMP
mCMPID = GetCMPID(TxtCMP)
strSplitString = Split(TxtGSL, "~")
MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 2) = Trim(strSplitString(0))
MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 3) = Trim(strSplitString(1))
MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 4) = Trim(strSplitString(2))
MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 5) = Val(TxtUnpledgeNoOfBags) - Val(TxtPledgeNoOfBags)
MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 6) = Format(Val(TxtUnpledgeWeight) - Val(TxtPledgeQuantity), "########0.000")
MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 7) = Val(TxtPledgeNoOfBags)
MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 8) = Val(TxtPledgeQuantity)


MSHFlexGrid2.TextMatrix(0, 7) = Val(MSHFlexGrid2.TextMatrix(0, 7)) + Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 7))
MSHFlexGrid2.TextMatrix(0, 8) = Val(MSHFlexGrid2.TextMatrix(0, 8)) + Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 8))

TxtCMP = ""
TxtGSL = ""
TxtUnpledgeNoOfBags = ""
TxtUnpledgeWeight = ""
TxtPledgeNoOfBags = ""
TxtPledgeQuantity = ""
TxtPledgeNoOfBags.Locked = True
TxtPledgeQuantity.Locked = True
CmdUpdate.Enabled = False
End Sub

Private Sub DeleteCmd_Click()
Dim strSplitString() As String
Dim ans As Variant
On Error GoTo msgError
ans = MsgBox("Are you Sure you want to Delete this Record", vbYesNo)
If ans = vbYes Then
   Call TableTxn_PledgeRelease("Where PledgeId = " & TxtPledgeID & "")
   If RsTxn_PledgeRelease.RecordCount > 0 Then
      MsgBox "Released Pledge Can't Delete!!!"
      Exit Sub
   End If
   
   'tmp = "Select * from Mst_GSL where sm_prefix = '" & TxtPreFix & "' And GSLID = " & MSHFlexGrid2.TextMatrix(2, 9) & " "
   RsMst_GSL.MoveFirst
   For x = 2 To MSHFlexGrid2.Rows - 1
       If MSHFlexGrid2.TextMatrix(x, 1) = "" Then Exit For
       RsMst_GSL.Filter = " sm_prefix = '" & TxtPreFix & "' And GSLID = " & MSHFlexGrid2.TextMatrix(x, 9) & " And ExchangeTypeID = " & mExchangeTypeID & " "
       If RsMst_GSL.RecordCount > 0 Then
          strSplitString = Split(RsMst_GSL!PledgeNo, "~")
          For i = 0 To UBound(strSplitString)
              If TxtPledgeID = strSplitString(i) Then
                 If i = 0 Then
                    RsMst_GSL!PledgeNo = ""
                 End If
              Else
                 If i = 0 Then
                    RsMst_GSL!PledgeNo = strSplitString(i)
                 Else
                    If Trim(RsMst_GSL!PledgeNo) = "" Then
                       RsMst_GSL!PledgeNo = strSplitString(i)
                    Else
                       RsMst_GSL!PledgeNo = Trim(RsMst_GSL!PledgeNo) & "~" & strSplitString(i)
                    End If
                 End If
              End If
          Next
          If RsMst_GSL!PledgeBags > 0 Then
             RsMst_GSL!PledgeBags = Val(RsMst_GSL!PledgeBags) - Val(MSHFlexGrid2.TextMatrix(x, 7))
             'RsMst_GSL!PledgeWeight = Val(IIf(IsNull(RsMst_GSL!PledgeWeight), "0", RsMst_GSL!PledgeWeight)) + Val(MSHFlexGrid2.TextMatrix(i, 8)) 'MSHFlexGrid2.TextMatrix(i, 8)
             'RsMst_GSL.Update
          End If
          If RsMst_GSL!PledgeWeight > 0 Then
             RsMst_GSL!PledgeWeight = Val(RsMst_GSL!PledgeWeight) - Val(MSHFlexGrid2.TextMatrix(x, 8))
          End If
          RsMst_GSL.Update
          RsMst_GSL.Filter = " GSLID <> Null"
       Else
          Exit Sub
       End If
   Next
   If mExchangeTypeID = 1 Then
      RsTxn_WHRDetail.MoveFirst
      RsTxn_WHRDetail.Find "WHRNo = " & mWHRNo
      If RsTxn_WHRDetail.EOF = False Then
         RsTxn_WHRDetail!PledgeID = Null
         RsTxn_WHRDetail.Update
      End If
   End If
   
   tmp = "Delete from Txn_PledgeCreation_GSL where PledgeID = " & TxtPledgeID
   Call TmpTable
   
   RsTxn_PledgeCreation.Delete
   RsTxn_PledgeCreation.Update
   
   'RsTxn_PledgeCreation_GSL.MoveFirst
   If RsTxn_PledgeCreation.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsTxn_PledgeCreation.MoveNext
   If RsTxn_PledgeCreation.EOF() Then
      RsTxn_PledgeCreation.MoveLast
   End If
   Call InData
End If
Exit Sub
msgError:
MsgBox "Child Record Present "
RsTxn_PledgeCreation.CancelUpdate
End Sub

Private Sub EditCmd_Click()
Edit_Flg = "E"
End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub

Public Sub Grid_Head()
MSHFlexGrid2.Clear
MSHFlexGrid2.Cols = 10
MSHFlexGrid2.Rows = 3
MSHFlexGrid2.FixedRows = 2

MSHFlexGrid2.TextMatrix(1, 0) = "  "
MSHFlexGrid2.TextMatrix(1, 1) = "CMP Name"
MSHFlexGrid2.TextMatrix(1, 2) = "Godown No"
MSHFlexGrid2.TextMatrix(1, 3) = "Stack No"
MSHFlexGrid2.TextMatrix(1, 4) = "Lot No"
MSHFlexGrid2.TextMatrix(1, 5) = "Unpledged  No Of Bags"
MSHFlexGrid2.TextMatrix(1, 6) = "Unpledged  Qty"
MSHFlexGrid2.TextMatrix(1, 7) = "Pledged No Of Bags"
MSHFlexGrid2.TextMatrix(1, 8) = "Pledged Qty"
MSHFlexGrid2.TextMatrix(1, 9) = "PledgeGSLID"


MSHFlexGrid2.ColWidth(0) = 400
MSHFlexGrid2.ColWidth(1) = 3200
MSHFlexGrid2.ColWidth(2) = 1200
MSHFlexGrid2.ColWidth(3) = 1000
MSHFlexGrid2.ColWidth(4) = 1000
MSHFlexGrid2.ColWidth(5) = 1800
MSHFlexGrid2.ColWidth(6) = 1800
MSHFlexGrid2.ColWidth(7) = 1600
MSHFlexGrid2.ColWidth(8) = 1600
MSHFlexGrid2.ColWidth(9) = 0

MSHFlexGrid2.RowHeight(1) = 400
End Sub
Public Sub Grid_Head1()
MSHFlexGrid1.Clear
MSHFlexGrid1.Cols = 10
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.FixedRows = 1

MSHFlexGrid1.TextMatrix(0, 0) = "PledgeID"
MSHFlexGrid1.TextMatrix(0, 1) = "Pledge Date"
MSHFlexGrid1.TextMatrix(0, 2) = "WHR No/ISINID"
MSHFlexGrid1.TextMatrix(0, 3) = "Bank Name"
MSHFlexGrid1.TextMatrix(0, 4) = "Branch Name"
MSHFlexGrid1.TextMatrix(0, 5) = "Pledge Request No"
MSHFlexGrid1.TextMatrix(0, 6) = "Pledge Fees"
MSHFlexGrid1.TextMatrix(0, 7) = "Pledge Request Date"
MSHFlexGrid1.TextMatrix(0, 8) = "Remarks"
MSHFlexGrid1.TextMatrix(0, 9) = "Exchange"

MSHFlexGrid1.ColWidth(0) = 900
MSHFlexGrid1.ColWidth(1) = 1000
MSHFlexGrid1.ColWidth(2) = 1200
MSHFlexGrid1.ColWidth(3) = 1600
MSHFlexGrid1.ColWidth(4) = 1400
MSHFlexGrid1.ColWidth(5) = 1600
MSHFlexGrid1.ColWidth(6) = 1400
MSHFlexGrid1.ColWidth(7) = 1800
MSHFlexGrid1.ColWidth(8) = 2800
MSHFlexGrid1.ColWidth(9) = 1000
End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Call GButtonLockAD(Me)
   MsgBox "Access Denied !!!!!!!!!"
   Exit Sub
End If

Call Grid_Head
'SSTab1.Tab = 0
Call TableTxn_PledgeCreation
If RsTxn_PledgeCreation.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
Call InData

End Sub

Private Sub ClearData()
TxtPledgeID = ""
TxtWHRNo = ""
TxtGodownNo = ""
TxtWHRNoOfBags = ""
TxtWHRQuantity = ""
TxtClientID = ""
TxtClientName = ""
TxtCommodity = ""
TxtLicenseNo = ""
CmbBankName.Text = ""
CmbBranchName.Text = ""
TxtPledgeReqNo = ""
TxtPledgeFee = ""
TxtRemarks = ""
TxtDPID = ""
TxtDPName = ""

Call Grid_Head
End Sub

Private Sub GetGridData(mExchangeTypeID_, mCommodityID_, mClientIDRow_, mCMPID_)
Call Grid_Head
If mExchangeTypeID_ = "" Or mExchangeTypeID_ = 0 Then Exit Sub
If mCommodityID_ = "" Or mCommodityID_ = 0 Then Exit Sub

If Label1.Caption = "WHR No" Then
  If mClientIDRow_ = "" Or mClientIDRow_ = 0 Then Exit Sub
  If mGodownID = "" Or mGodownID = 0 Then Exit Sub
End If
If mCMPID_ = "" Or mCMPID_ = 0 Then Exit Sub

If CmbExchange.Text = "Non NCDEX" Then
   tmp = "SELECT Mst_GSL.GSLID, Mst_GSL.StackNo, Mst_GSL.LotNo, (Mst_GSL.BalanceBags-Mst_GSL.PledgeBags) as UnpledgeBags, (Mst_GSL.BalanceWeight-Mst_GSL.PledgeWeight) as UnpledgeWeight, " & _
         "Mst_GSL.PledgeBags, Mst_GSL.PledgeWeight, Mst_GSL.PledgeNo, Mst_CMP.CMPName, Mst_Godown.GodownNo " & _
         "FROM Mst_Godown INNER JOIN (Mst_CMP INNER JOIN Mst_GSL ON Mst_CMP.CMPID = Mst_GSL.CMPID) ON Mst_Godown.GodownID = Mst_GSL.GodownID " & _
         "WHERE SM_Prefix = '" & TxtPreFix & "' And (Mst_GSL.Flag <> 'E' ) And (((Mst_GSL.ExchangeTypeID)= " & mExchangeTypeID_ & ")AND ((Mst_GSL.CMPID)=" & mCMPID_ & ") AND  " & _
         "((Mst_GSL.CommodityID)=" & mCommodityID_ & ") And Mst_GSL.GodownID = " & mGodownID & _
         "AND ((Mst_GSL.DepositeWeight - Mst_GSL.PledgeWeight) > 0)) AND ((Mst_GSL.ClientIDRow)=" & mClientIDRow_ & ")"
Else
   tmp = "SELECT Mst_GSL.GSLID, Mst_GSL.StackNo, Mst_GSL.LotNo, (Mst_GSL.BalanceBags-Mst_GSL.PledgeBags) as UnpledgeBags, (Mst_GSL.BalanceWeight-Mst_GSL.PledgeWeight) as UnpledgeWeight, " & _
         "Mst_GSL.PledgeBags, Mst_GSL.PledgeWeight, Mst_GSL.PledgeNo, Mst_CMP.CMPName, Mst_Godown.GodownNo " & _
         "FROM Mst_Godown INNER JOIN (Mst_CMP INNER JOIN Mst_GSL ON Mst_CMP.CMPID = Mst_GSL.CMPID) ON Mst_Godown.GodownID = Mst_GSL.GodownID " & _
         "WHERE SM_Prefix = '" & TxtPreFix & "' And (Mst_GSL.Flag <> 'E' ) And (((Mst_GSL.ExchangeTypeID)= " & mExchangeTypeID_ & ")AND ((Mst_GSL.CMPID)=" & mCMPID_ & ") AND  " & _
         "((Mst_GSL.CommodityID)=" & mCommodityID_ & ") AND ((Mst_GSL.ISINID) = '" & TxtWHRNo.Text & "')) AND ((Mst_GSL.DepositeWeight - Mst_GSL.PledgeWeight) > 0)"
End If

Call TmpTable

If TmpRs.RecordCount = 0 Then
  MsgBox "No Record For Grid"
  Exit Sub
End If

k = 1
For i = 0 To TmpRs.RecordCount - 1
  k = MSHFlexGrid2.Rows - 1
  MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
  MSHFlexGrid2.TextMatrix(k, 0) = ""
  MSHFlexGrid2.TextMatrix(k, 1) = TmpRs!CMPName
  MSHFlexGrid2.TextMatrix(k, 2) = TmpRs!GodownNo
  MSHFlexGrid2.TextMatrix(k, 3) = TmpRs!StackNo
  MSHFlexGrid2.TextMatrix(k, 4) = TmpRs!LotNo
  MSHFlexGrid2.TextMatrix(k, 5) = TmpRs!UnPledgeBags
  MSHFlexGrid2.TextMatrix(k, 6) = TmpRs!UnPledgeWeight
  'MSHFlexGrid2.TextMatrix(k, 7) = TmpRs!PledgeBags
  'MSHFlexGrid2.TextMatrix(k, 8) = TmpRs!PledgeWeight
  MSHFlexGrid2.TextMatrix(k, 9) = TmpRs!GSLID
  TmpRs.MoveNext
Next

End Sub
Private Sub MSHFlexGrid2_Click()
If SaveCmd.Enabled = False Then Exit Sub
TxtCMP = ""
TxtGSL = ""
TxtUnpledgeNoOfBags = ""
TxtUnpledgeWeight = ""
TxtPledgeQuantity = ""
TxtPledgeNoOfBags = ""
CmdUpdate.Enabled = False


If MSHFlexGrid2.Rows > 3 Then
    If MSHFlexGrid2.RowSel <> "0" Then
       If MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 1) = "" Then Exit Sub
       TxtPledgeNoOfBags.Locked = False
       TxtPledgeQuantity.Locked = False
       TxtCMP = MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 1)
       TxtGSL = MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 2) & "~" & MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 3) & "~" & MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 4)
       TxtUnpledgeNoOfBags = Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 5)) + Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 7))
       TxtUnpledgeWeight = Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 6)) + Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 8))
       
       
       If Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 7)) = 0 Then
          TxtPledgeNoOfBags = Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 5))
       Else
          TxtPledgeNoOfBags = Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 7))
       End If
       If Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 8)) = 0 Then
          TxtPledgeQuantity = Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 6))
       Else
          TxtPledgeQuantity = Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 8))
       End If
       
       CmdUpdate.Enabled = True
    End If
End If

End Sub



Private Sub SaveCmd_Click()
If PledgeDate.Value > Date Then
  MsgBox "Pledge Date Should Not Be Future Date!!!"
  PledgeDate.SetFocus
  Exit Sub
End If

If CmbExchange.Text = "Non NCDEX" Then
   If PledgeDate.Value < WHRDate.Value Then
      MsgBox "Pledge Date must be greater than WHR Date !!!"
      PledgeDate.SetFocus
      Exit Sub
   End If
Else
   If TxtClientID = "" Then
      MsgBox "Blank Client ID not allowed  !!!"
      TxtClientID.SetFocus
      Exit Sub
   End If

End If
If PledgeRequestDate > PledgeDate.Value Then
   MsgBox "Pledge Request Date must be less than or equal to Pledge date !!!!"
   PledgeRequestDate.SetFocus
   Exit Sub
End If
If TxtWHRNo.Text = "" Then
   If Label1.Caption = "WHR No" Then
      MsgBox "Blank WHR No Not Allowed!!!"
   Else
      MsgBox "Blank ISIN No Not Allowed!!!"
   End If
   TxtWHRNo.SetFocus
   Exit Sub
End If
If CmbBankName.Text = "" Then
   MsgBox "Provide Pledged to Bank details in 'Pledge Detail' tab!!!"
   CmbBankName.SetFocus
   Exit Sub
End If
If CmbBranchName.Text = "" Then
   MsgBox "Blank Branch Name Not Allowed!!!"
   CmbBranchName.SetFocus
   Exit Sub
End If
If TxtPledgeReqNo = "" Then
   MsgBox "Blank Pledge Request No Not Allowed!!!"
   TxtPledgeReqNo.SetFocus
   Exit Sub
End If
'If TxtPledgeFee = "" Then
'   MsgBox "Blank Pledge Fee Receviable Not Allowed!!!"
'   TxtPledgeFee.SetFocus
'   Exit Sub
'End If

If PledgeRequestDate.Value > Date Then
  MsgBox "Pledge Request Date Should Not Be Future Date!!!"
  PledgeRequestDate.SetFocus
  Exit Sub
End If

If MSHFlexGrid2.Rows < 4 Then
   MsgBox "Blank Grid Not Allowed !!!"
   TxtWHRNo.SetFocus
   Exit Sub
End If

If ChkSelection() = False Then
    MsgBox "Please Select Atleast One Row From Grid"
    MSHFlexGrid2.SetFocus
    Exit Sub
End If

If Label1.Caption = "WHR No" Then
   If Val(MSHFlexGrid2.TextMatrix(0, 7)) <> Val(TxtWHRNoOfBags) Then
      MsgBox "Total Pledge No Of Bags Should Be Equal To WHR No Of Bags"
      MSHFlexGrid2.SetFocus
      Exit Sub
   End If
   If Val(MSHFlexGrid2.TextMatrix(0, 8)) <> Val(TxtWHRQuantity) Then
      MsgBox "Total Pledge Quantity Should Be Equal To WHR Quantity"
      MSHFlexGrid2.SetFocus
      Exit Sub
   End If
End If
If Edit_Flg = "A" Then
   RsTxn_PledgeCreation.AddNew
   RsTxn_PledgeCreation!PledgeID = GetPledgeID
   RsTxn_PledgeCreation!G_UID = GetGUID
   TxtPledgeID = RsTxn_PledgeCreation!PledgeID
Else

End If

RsTxn_PledgeCreation!PledgeDate = PledgeDate.Value
If Label1.Caption = "WHR No" Then
  RsTxn_PledgeCreation!WHRNo = mWHRNo
  RsTxn_PledgeCreation!ISINID = "0"
Else
  RsTxn_PledgeCreation!WHRNo = 0
  RsTxn_PledgeCreation!ISINID = mISINID
End If
RsTxn_PledgeCreation!ExchangeTypeID = mExchangeTypeID
RsTxn_PledgeCreation!BankID = mBankID
RsTxn_PledgeCreation!BranchID = mBranchID
RsTxn_PledgeCreation!PledgeReqNo = TxtPledgeReqNo
RsTxn_PledgeCreation!PledgeFees = Val(TxtPledgeFee)
RsTxn_PledgeCreation!PledgeReqDate = PledgeRequestDate.Value
RsTxn_PledgeCreation!Remarks = IIf(IsNull(TxtRemarks), "", TxtRemarks)
RsTxn_PledgeCreation!ClientIDRow = mClientIDRow
RsTxn_PledgeCreation!sm_prefix = TxtPreFix
RsTxn_PledgeCreation.Update
If Label1.Caption = "WHR No" Then
  RsTxn_WHRDetail!PledgeID = TxtPledgeID
  RsTxn_WHRDetail.Update
End If
Call SavePledgeCreation_GSL
End Sub
Private Sub SavePledgeCreation_GSL()
Call TableTxn_PledgeCreation_GSL("Where SM_Prefix = '" & TxtPreFix & "' And PledgeID = " & TxtPledgeID)
Call TableMst_GSL(" Where SM_Prefix = '" & TxtPreFix & "'")
For i = 2 To MSHFlexGrid2.Rows - 1
  If MSHFlexGrid2.TextMatrix(i, 1) = "" Then Exit For
  If MSHFlexGrid2.TextMatrix(i, 0) = "*" Then
    RsTxn_PledgeCreation_GSL.AddNew
    RsTxn_PledgeCreation_GSL!G_UID = GetGUID
    RsTxn_PledgeCreation_GSL!PledgeNo = GetPledgeNo
    RsTxn_PledgeCreation_GSL!PledgeID = TxtPledgeID
    RsTxn_PledgeCreation_GSL!GSLID = MSHFlexGrid2.TextMatrix(i, 9)
    RsTxn_PledgeCreation_GSL!PledgeNoOfBags = MSHFlexGrid2.TextMatrix(i, 7)
    RsTxn_PledgeCreation_GSL!PledgeQuantity = MSHFlexGrid2.TextMatrix(i, 8)
    RsTxn_PledgeCreation_GSL!sm_prefix = TxtPreFix
    RsTxn_PledgeCreation_GSL.Update
    RsMst_GSL.MoveFirst
    RsMst_GSL.Find "GSLID = " & Val(MSHFlexGrid2.TextMatrix(i, 9))
    If Not RsMst_GSL.EOF Then
      If IsNull(RsMst_GSL!PledgeNo) Or RsMst_GSL!PledgeNo = 0 Or RsMst_GSL!PledgeNo = "" Then
        RsMst_GSL!PledgeNo = TxtPledgeID
      Else
        RsMst_GSL!PledgeNo = Trim(RsMst_GSL!PledgeNo) & "~" & TxtPledgeID
      End If
      RsMst_GSL!PledgeBags = Val(IIf(IsNull(RsMst_GSL!PledgeBags), "0", RsMst_GSL!PledgeBags)) + Val(MSHFlexGrid2.TextMatrix(i, 7))
      RsMst_GSL!PledgeWeight = Val(IIf(IsNull(RsMst_GSL!PledgeWeight), "0", RsMst_GSL!PledgeWeight)) + Val(MSHFlexGrid2.TextMatrix(i, 8)) 'MSHFlexGrid2.TextMatrix(i, 8)
      RsMst_GSL.Update
    End If
    RsTxn_PledgeCreation_GSL.Requery
  End If
Next
Call InData
End Sub
Private Sub InData()
mExchangeTypeID = RsTxn_PledgeCreation!ExchangeTypeID
CmbExchange.Text = GetExchangeName(mExchangeTypeID)

If CmbExchange.Text = "NCDEX" Then
  Call EnDn("NCDEX", False)
  mClientIDRow = 0
  TxtClientID.BackColor = -2147483643
  TxtClientID.Enabled = True
  TxtClientID.Locked = True
  lblDpIDName.Visible = True
  TxtDPID.Visible = True
  TxtDPName.Visible = True
Else
  Call EnDn("Non NCDEX", True)
  WHRDate.Enabled = False
  mClientIDRow = 0
  TxtClientID.BackColor = -2147483648#
  TxtClientID.Enabled = False
  TxtClientID.Locked = True
  lblDpIDName.Visible = False
  TxtDPID.Visible = False
  TxtDPName.Visible = False
End If
TxtPledgeID = IIf(IsNull(RsTxn_PledgeCreation!PledgeID), "0", RsTxn_PledgeCreation!PledgeID)
PledgeDate = RsTxn_PledgeCreation!PledgeDate
TxtPreFix = RsTxn_PledgeCreation!sm_prefix
If Label1.Caption = "WHR No" Then
    
    TxtWHRNo = RsTxn_PledgeCreation!WHRNo
    Call TableTxn_WHRDetail("Where WHRNo = " & TxtWHRNo & " And sm_prefix = '" & TxtPreFix & "'")
    WHRDate.Value = RsTxn_WHRDetail!WHRDate
    mClientIDRow = RsTxn_WHRDetail!ClientIDRow
    TxtWHRNoNew = RsTxn_WHRDetail!CMPID & "-" & (Val(TxtWHRNo) - (RsTxn_WHRDetail!CMPID * 1000000))
    ans = GetClientNameByRow(mClientIDRow)
    If ans(0) = "" Then
       MsgBox "Client ID  Not Found!!!"
       Exit Sub
    End If
    TxtClientID.Text = ans(0)
    TxtClientName = ans(1)
    mCommodityId = IIf(IsNull(RsTxn_WHRDetail!CommodityID), "", RsTxn_WHRDetail!CommodityID)
    TxtLicenseNo = GetLicenseNo(RsTxn_WHRDetail!LicenseID)
    TxtWHRNoOfBags = IIf(IsNull(RsTxn_WHRDetail!WHRNoOfBags), "", RsTxn_WHRDetail!WHRNoOfBags)
    TxtWHRQuantity = IIf(IsNull(RsTxn_WHRDetail!WHRQuantity), "", RsTxn_WHRDetail!WHRQuantity)
    TxtGodownNo = GetGodownName(IIf(IsNull(RsTxn_WHRDetail!GodownID), "", RsTxn_WHRDetail!GodownID), "N")
Else
    mClientIDRow = RsTxn_PledgeCreation!ClientIDRow
    ans = GetClientNameByRow(mClientIDRow)
    If ans(0) = "" Then
       'MsgBox "Client ID Not Found!!!"
       Exit Sub
    End If
    TxtClientID.Text = ans(0)
    TxtClientName = ans(1)
    TxtDPID.Text = ans(2)
    TxtDPName = ans(3)
    
    TxtWHRNoNew = RsTxn_PledgeCreation!ISINID
    TxtWHRNo = RsTxn_PledgeCreation!ISINID
    Call TableMst_ISIN("Where ISINID = '" & TxtWHRNo & "'")
    mCommodityId = IIf(IsNull(RsMst_ISIN!CommodityID), "", RsMst_ISIN!CommodityID)
    mCMPID = IIf(IsNull(RsMst_ISIN!CMPID), "", RsMst_ISIN!CMPID)
    TxtLicenseNo = GetCMPName(mCMPID)
End If

TxtCommodity = GetCommodityName(mCommodityId)
TxtRemarks = IIf(IsNull(RsTxn_PledgeCreation!Remarks), "", RsTxn_PledgeCreation!Remarks)
TxtBankName = GetBankName(RsTxn_PledgeCreation!BankID, "N")
TxtBranchName = GetBranchName(RsTxn_PledgeCreation!BranchID, "N")
TxtPledgeReqNo = IIf(IsNull(RsTxn_PledgeCreation!PledgeReqNo), "", RsTxn_PledgeCreation!PledgeReqNo)
TxtPledgeFee = IIf(IsNull(RsTxn_PledgeCreation!PledgeFees), "", RsTxn_PledgeCreation!PledgeFees)
PledgeRequestDate.Value = RsTxn_PledgeCreation!PledgeReqDate
CmbBankName.Visible = False
CmbBranchName.Visible = False
TxtWHRNo.Visible = False
TxtWHRNo.Locked = True
TxtBankName.Visible = True
TxtBranchName.Visible = True
TxtPledgeFee.Locked = True
TxtPledgeReqNo.Locked = True
PledgeDate.Enabled = False
PledgeRequestDate.Enabled = False
CmdBankName.Visible = False
CmdBranchName.Visible = False
TxtBankName.Locked = True
TxtBranchName.Locked = True
TxtRemarks.Locked = True
TxtPledgeNoOfBags.Locked = True
TxtPledgeQuantity.Locked = True
CmbExchange.Enabled = False
CmdUpdate.Enabled = False
Call Grid_Head
Call TableMst_GSL(" Where Sm_prefix = '" & TxtPreFix & "'")
Call TableTxn_PledgeCreation_GSL("Where sm_prefix = '" & TxtPreFix & "' And PledgeID = " & TxtPledgeID)
MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + RsTxn_PledgeCreation_GSL.RecordCount
k = 1
For i = 0 To RsTxn_PledgeCreation_GSL.RecordCount - 1
  k = k + 1
  RsMst_GSL.Find "GSLID = " & RsTxn_PledgeCreation_GSL!GSLID
  If Not RsMst_GSL.EOF Then
     MSHFlexGrid2.TextMatrix(k, 1) = GetCMPName(RsMst_GSL!CMPID)
     MSHFlexGrid2.TextMatrix(k, 2) = GetGodownName(RsMst_GSL!GodownID, "N")
     MSHFlexGrid2.TextMatrix(k, 3) = RsMst_GSL!StackNo
     MSHFlexGrid2.TextMatrix(k, 4) = RsMst_GSL!LotNo
     MSHFlexGrid2.TextMatrix(k, 5) = "0" 'RsMst_GSL!WHRBags
     MSHFlexGrid2.TextMatrix(k, 6) = "0" 'RsMst_GSL!WHRWeight
     MSHFlexGrid2.TextMatrix(k, 7) = RsTxn_PledgeCreation_GSL!PledgeNoOfBags
     MSHFlexGrid2.TextMatrix(k, 8) = RsTxn_PledgeCreation_GSL!PledgeQuantity
     MSHFlexGrid2.TextMatrix(k, 9) = RsTxn_PledgeCreation_GSL!GSLID
  End If
  RsTxn_PledgeCreation_GSL.MoveNext
Next i
Call GButtonLock(Me, True)
EditCmd.Enabled = False
'DeleteCmd.Enabled = False

End Sub
Private Function GetPledgeID() As Double
Dim RetVal As Double
tmp = "Select max(PledgeID) from Txn_PledgeCreation where SM_Prefix = '" & TxtPreFix & "'"
Call TmpTable
RetVal = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
RetVal = RetVal + 1
GetPledgeID = RetVal
End Function
Private Function GetPledgeNo() As Double
Dim RetVal As Double
tmp = "Select max(PledgeNo) from Txn_PledgeCreation_GSL where SM_Prefix = '" & TxtPreFix & "'"
Call TmpTable
RetVal = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
RetVal = RetVal + 1
GetPledgeNo = RetVal
End Function

Private Sub SearchCmd_Click()
Dim ans As Variant
If SearchCmd.Caption = "Cancel" Then
   If RsTxn_PledgeCreation.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   Call FirstCmd_Click
   Exit Sub
End If
Frame0.Height = Me.Height - 100
Frame0.Left = Picture1.Left - 100
Frame0.Top = Picture1.Top - 100
Frame0.Width = Me.Width - 150
Frame0.Visible = True
MSHFlexGrid1.Clear
MSHFlexGrid1.Cols = 3
MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 1000
RsTxn_PledgeCreation.MoveFirst
Call Grid_Head1
MSHFlexGrid1.Rows = RsTxn_PledgeCreation.RecordCount + 1
RsTxn_PledgeCreation.MoveFirst
For i = 1 To RsTxn_PledgeCreation.RecordCount
    MSHFlexGrid1.TextMatrix(i, 0) = IIf(IsNull(RsTxn_PledgeCreation!PledgeID), "", RsTxn_PledgeCreation!PledgeID)
    MSHFlexGrid1.TextMatrix(i, 1) = IIf(IsNull(RsTxn_PledgeCreation!PledgeDate), "", RsTxn_PledgeCreation!PledgeDate)
    If RsTxn_PledgeCreation!WHRNo <> 0 Then
      'MSHFlexGrid1.TextMatrix(i, 2) = IIf(IsNull(RsTxn_PledgeCreation!WHRNo), "", RsTxn_PledgeCreation!WHRNo)
      Call TableTxn_WHRDetail("Where sm_prefix = '" & RsTxn_PledgeCreation!sm_prefix & "' And WHRNo = " & RsTxn_PledgeCreation!WHRNo)
      MSHFlexGrid1.TextMatrix(i, 2) = RsTxn_WHRDetail!CMPID & "-" & ((RsTxn_PledgeCreation!WHRNo) - RsTxn_WHRDetail!CMPID * 1000000)
    Else
      MSHFlexGrid1.TextMatrix(i, 2) = IIf(IsNull(RsTxn_PledgeCreation!ISINID), "", RsTxn_PledgeCreation!ISINID)
    End If
    mBankID = IIf(IsNull(RsTxn_PledgeCreation!BankID), "", RsTxn_PledgeCreation!BankID)
    MSHFlexGrid1.TextMatrix(i, 3) = GetBankName(mBankID, "N")
    mBranchID = IIf(IsNull(RsTxn_PledgeCreation!BranchID), "", RsTxn_PledgeCreation!BranchID)
    MSHFlexGrid1.TextMatrix(i, 4) = GetBranchName(mBranchID, "N")
    MSHFlexGrid1.TextMatrix(i, 5) = IIf(IsNull(RsTxn_PledgeCreation!PledgeReqNo), "", RsTxn_PledgeCreation!PledgeReqNo)
    MSHFlexGrid1.TextMatrix(i, 6) = IIf(IsNull(RsTxn_PledgeCreation!PledgeFees), "", RsTxn_PledgeCreation!PledgeFees)
    MSHFlexGrid1.TextMatrix(i, 7) = IIf(IsNull(RsTxn_PledgeCreation!PledgeReqDate), "", RsTxn_PledgeCreation!PledgeReqDate)
    MSHFlexGrid1.TextMatrix(i, 8) = IIf(IsNull(RsTxn_PledgeCreation!Remarks), "", RsTxn_PledgeCreation!Remarks)
    MSHFlexGrid1.TextMatrix(i, 9) = GetExchangeName(RsTxn_PledgeCreation!ExchangeTypeID)
    RsTxn_PledgeCreation.MoveNext
Next
End Sub
Private Sub MSHFlexGrid1_Click()
Dim i As Variant
i = MSHFlexGrid1.RowSel
If i <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(i, 0) = "" Then Exit Sub
RsTxn_PledgeCreation.MoveFirst
RsTxn_PledgeCreation.Find "PledgeID = " & Val(MSHFlexGrid1.TextMatrix(i, 0))
If Not RsTxn_PledgeCreation.EOF Then
   Call InData
   Frame0.Visible = False
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
End Sub

Private Sub Text1_Change()

End Sub

Private Sub TxtClientID_LostFocus()
CmbClientDP.Visible = False
If TxtClientID = "" Then Exit Sub

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
      "  And  (((Mst_Client.ExchangeTypeID)=" & mExchangeTypeID & ") AND ((Mst_Client.ClientID)='" & TxtClientID & "'));"
      
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
End If
End Sub

Private Sub TxtPledgeFee_LostFocus()
If TxtPledgeFee = "" Then Exit Sub
If IsNumeric(TxtPledgeFee) = False Then
   MsgBox "Invalid Numeric Format!!!"
   TxtPledgeFee.SetFocus
   Exit Sub
End If
End Sub
Private Sub TxtPledgeNoOfBags_LostFocus()
If TxtPledgeNoOfBags = "" Then Exit Sub
If IsNumeric(TxtPledgeNoOfBags) = False Then
   MsgBox "Invalid Number Format!!!"
   TxtPledgeNoOfBags.SetFocus
   Exit Sub
End If
If Val(TxtPledgeNoOfBags) < 0 Then
   MsgBox "Pledge No Of Bags Can't Be Negetive!!!"
   TxtPledgeNoOfBags.SetFocus
   Exit Sub
End If
If Val(TxtPledgeNoOfBags) > Val(TxtUnpledgeNoOfBags) Then
   MsgBox "Pledge No Of Bags Should Be Less Than Unpledge No Of Bags"
   TxtPledgeNoOfBags.SetFocus
   Exit Sub
End If
End Sub
Private Sub TxtPledgeQuantity_LostFocus()
If TxtPledgeQuantity = "" Then Exit Sub
If IsNumeric(TxtPledgeQuantity) = False Then
   MsgBox "Invalid Number Format!!!"
   TxtPledgeQuantity.SetFocus
   Exit Sub
End If
If Val(TxtPledgeQuantity) < 0 Then
   MsgBox "Pledge Quantity Can't Be Negetive!!!"
   TxtPledgeQuantity.SetFocus
   Exit Sub
End If
If Val(TxtPledgeQuantity) > Val(TxtUnpledgeWeight) Then
   MsgBox "Pledge Quantity Should Be Less Than Unpledge Quantity"
   TxtPledgeQuantity.SetFocus
   Exit Sub
End If
'CmdUpdate.Enabled = True
End Sub
Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsTxn_PledgeCreation.MoveNext
If RsTxn_PledgeCreation.EOF Then
   RsTxn_PledgeCreation.MoveLast
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
RsTxn_PledgeCreation.MovePrevious
If RsTxn_PledgeCreation.BOF Then
   RsTxn_PledgeCreation.MoveFirst
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
RsTxn_PledgeCreation.MoveFirst
Call InData
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsTxn_PledgeCreation.MoveLast
Call InData
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False
End Sub
Function ChkSelection() As Boolean
Dim RetVal As Boolean
RetVal = False
If MSHFlexGrid2.Rows > 3 Then
    For i = 2 To MSHFlexGrid2.Rows - 1
      If MSHFlexGrid2.TextMatrix(i, 0) = "*" Then
        RetVal = True
        Exit For
      End If
    Next
End If
ChkSelection = RetVal
End Function
Private Sub EnDn(Exg, Yn)

If Exg = "Non NCDEX" Then
  Label1.Caption = "WHR No"
  Label2.Caption = "License No"
ElseIf Exg = "NCDEX" Then
  Label1.Caption = "ISIN ID"
  Label2.Caption = "CMP"
End If
TxtClientID.Enabled = Yn
TxtClientName.Enabled = Yn
WHRDate.Enabled = Yn
TxtWHRNoOfBags.Enabled = Yn
TxtWHRQuantity.Enabled = Yn
'Label6.Enabled = Yn
'Label8.Enabled = Yn
Label16.Enabled = Yn
Label17.Enabled = Yn
Label14.Enabled = Yn

TxtClientID.Text = ""
TxtClientName.Text = ""
TxtWHRNoOfBags.Text = ""
TxtWHRQuantity.Text = ""

End Sub

Private Sub TxtWHRNo_LostFocus()
If SaveCmd.Enabled = False Then Exit Sub
If TxtWHRNoNew.Text = "" Then
   Call Grid_Head
   Exit Sub
End If
If CmbExchange.Text = "Non NCDEX" Then
   If RsTxn_WHRDetail.RecordCount <= 0 Then
      MsgBox "No WHR Data found !!!! "
      TxtWHRNoNew.SetFocus
      Exit Sub
   End If
   Dim strSplitString() As String
   If TxtWHRNoNew.Text = "" Then
      Exit Sub
   End If
   strSplitString = Split(TxtWHRNoNew.Text, "-")
   tmp = UBound(strSplitString)
   If tmp <> 1 Then
      MsgBox "Invalid WHR No Selection"
      TxtWHRNoNew.SetFocus
      Exit Sub
   End If
   TxtWHRNo = (Val(strSplitString(0)) * 1000000) + Val(strSplitString(1))
    RsTxn_WHRDetail.MoveFirst
    RsTxn_WHRDetail.Find "WHRNo = " & TxtWHRNo.Text & ""
    If RsTxn_WHRDetail.EOF Then
       MsgBox "Invalid  selection "
       Call Grid_Head
       TxtWHRNoNew.SetFocus
       Exit Sub
    End If
    If RsTxn_WHRDetail!Flag <> "P" Or IsNull(RsTxn_WHRDetail!Flag) Then
       MsgBox " WHR Print is not Generated !!!!! "
       TxtWHRNoNew.SetFocus
       Exit Sub
    End If
    If RsTxn_WHRDetail!DeliveredNoOfBags = RsTxn_WHRDetail!WHRNoOfBags Then
       MsgBox "WHR Already Withdrawan  "
       TxtWHRNoNew.SetFocus
       Exit Sub
    End If
    If RsTxn_WHRDetail!PledgeID <> 0 Then
       MsgBox "WHR Already Pledge against Pledge ID " & RsTxn_WHRDetail!PledgeID
       TxtWHRNoNew.SetFocus
       Exit Sub
    End If
    
    mWHRNo = RsTxn_WHRDetail!WHRNo
    WHRDate.Value = RsTxn_WHRDetail!WHRDate
    mClientIDRow = RsTxn_WHRDetail!ClientIDRow
    ans = GetClientNameByRow(mClientIDRow)
    If ans(0) = "" Then
       MsgBox "Not Found!!!"
       TxtWHRNoNew.SetFocus
       Exit Sub
    End If
    TxtClientID.Text = ans(0)
    TxtClientName = ans(1)
    mCommodityId = IIf(IsNull(RsTxn_WHRDetail!CommodityID), "", RsTxn_WHRDetail!CommodityID)
    TxtCommodity = GetCommodityName(mCommodityId)
    TxtLicenseNo = GetLicenseNo(RsTxn_WHRDetail!LicenseID)
    TxtWHRNoOfBags = IIf(IsNull(RsTxn_WHRDetail!WHRNoOfBags), "", RsTxn_WHRDetail!WHRNoOfBags)
    TxtWHRQuantity = IIf(IsNull(RsTxn_WHRDetail!WHRQuantity), "", RsTxn_WHRDetail!WHRQuantity)
    TxtPreFix = IIf(IsNull(RsTxn_WHRDetail!sm_prefix), "", RsTxn_WHRDetail!sm_prefix)
    TxtWHRNoOfBags = TxtWHRNoOfBags - IIf(IsNull(RsTxn_WHRDetail!DeliveredNoOfBags), 0, RsTxn_WHRDetail!DeliveredNoOfBags)
    TxtWHRQuantity = TxtWHRQuantity - IIf(IsNull(RsTxn_WHRDetail!DeliveredQuantity), 0, RsTxn_WHRDetail!DeliveredQuantity)
    TxtWHRQuantity = Format(TxtWHRQuantity, "#########0.000")
    mCMPID = IIf(IsNull(RsTxn_WHRDetail!CMPID), "", RsTxn_WHRDetail!CMPID)
    mGodownID = IIf(IsNull(RsTxn_WHRDetail!GodownID), "", RsTxn_WHRDetail!GodownID)
    If Not IsNull(RsTxn_WHRDetail!GodownID) Then
       TxtGodownNo = GetGodownName(RsTxn_WHRDetail!GodownID, "N")
    End If
    Call GetGridData(mExchangeTypeID, mCommodityId, mClientIDRow, mCMPID)
Else
    TxtWHRNo = TxtWHRNoNew
    RsMst_ISIN.MoveFirst
    RsMst_ISIN.Find "ISINID = '" & TxtWHRNo.Text & "'"
    If RsMst_ISIN.EOF Then
       MsgBox "Invalid  selection "
       Call Grid_Head
       TxtWHRNoNew.SetFocus
       Exit Sub
    End If
    mISINID = RsMst_ISIN!ISINID
    mCommodityId = IIf(IsNull(RsMst_ISIN!CommodityID), "", RsMst_ISIN!CommodityID)
    TxtCommodity = GetCommodityName(mCommodityId)
    TxtLicenseNo = GetCMPName(RsMst_ISIN!CMPID)
    mCMPID = IIf(IsNull(RsMst_ISIN!CMPID), "", RsMst_ISIN!CMPID)
    TxtPreFix = GetPreFix(RsMst_ISIN!CMPID, "P")
    Call GetGridData(mExchangeTypeID, mCommodityId, mClientIDRow, mCMPID)

End If

End Sub

Private Sub TxtWHRNoNew_LostFocus()
Call TxtWHRNo_LostFocus ' = TxtWHRNoNew
TxtCMP = ""
TxtGSL = ""
TxtUnpledgeNoOfBags = ""
TxtUnpledgeWeight = ""
TxtPledgeNoOfBags = ""
TxtPledgeQuantity = ""
TxtPledgeNoOfBags.Locked = True
TxtPledgeQuantity.Locked = True
CmdUpdate.Enabled = False

End Sub
